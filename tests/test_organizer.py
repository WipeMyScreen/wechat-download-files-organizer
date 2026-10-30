"""Integration tests against disposable files; never uses the real WeChat root."""
import collections
import ctypes
import hashlib
import json
import msvcrt
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import unittest
import unicodedata
import uuid

POWERSHELL = os.environ.get("WECHAT_TEST_POWERSHELL", "powershell.exe")
# Each engine must discover its own modules rather than inherit another engine's.
POWERSHELL_ENV = {k: v for k, v in os.environ.items() if k.upper() != "PSMODULEPATH"}

SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "organize.ps1"
CATEGORIES = ["word", "excel", "ppt", "压缩包", "DWG", "PDF", "音频", "其他"]
EXTENSIONS = {
    "word": ".doc .docx .docm .dot .dotx .dotm .wps .wpt",
    "excel": ".xls .xlsx .xlsm .xlsb .xlt .xltx .xltm .csv .et .ett",
    "ppt": ".ppt .pptx .pptm .pot .potx .potm .pps .ppsx .ppsm .dps .dpt",
    "压缩包": ".zip .rar .7z .tar .gz .tgz .bz2 .xz .zst .cab",
    "DWG": ".dwg",
    "PDF": ".pdf",
    "音频": ".mp3 .wav .flac .m4a .m4b .aac .ogg .opus .wma .amr .aif .aiff .ape .alac .mid .midi .silk",
    "其他": ".png .yml .dxf",
}


def run(root, apply=False, check=True, include_records=False, as_json=True, deduplicate=False, migrate_audio=False):
    args = [POWERSHELL, "-NoProfile", "-NonInteractive", "-File", str(SCRIPT), "-RootPath", str(root)]
    if as_json:
        args.append("-AsJson")
    if include_records:
        args.append("-IncludeRecords")
    if apply:
        args.append("-Apply")
    if deduplicate:
        args.append("-Deduplicate")
    if migrate_audio:
        args.append("-MigrateLegacyAudio")
    result = subprocess.run(args, capture_output=True, encoding="utf-8", env=POWERSHELL_ENV)
    if not check:
        return result
    if result.returncode:
        raise AssertionError(result.stderr)
    return json.loads(result.stdout) if as_json else result.stdout


def discover(base, *extra):
    result = subprocess.run(
        [POWERSHELL, "-NoProfile", "-NonInteractive", "-File", str(SCRIPT), "-ListAccounts",
         "-WechatBasePath", str(base), "-AsJson", *extra],
        capture_output=True, encoding="utf-8", env=POWERSHELL_ENV,
    )
    if result.returncode:
        raise AssertionError(result.stderr)
    return json.loads(result.stdout)


def put(path, content=b"x" * 1024):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(content)
    return path


def put_sparse(path, size):
    """Use NTFS sparse files to test GB logical sizes without writing GB of data."""
    path.parent.mkdir(parents=True, exist_ok=True)
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel.DeviceIoControl.argtypes = [ctypes.c_void_p, ctypes.c_uint32, ctypes.c_void_p,
                                      ctypes.c_uint32, ctypes.c_void_p, ctypes.c_uint32,
                                      ctypes.POINTER(ctypes.c_uint32), ctypes.c_void_p]
    kernel.DeviceIoControl.restype = ctypes.c_int
    with path.open("wb") as stream:
        returned = ctypes.c_uint32()
        handle = msvcrt.get_osfhandle(stream.fileno())
        if not kernel.DeviceIoControl(handle, 0x900C4, None, 0, None, 0, ctypes.byref(returned), None):
            raise ctypes.WinError(ctypes.get_last_error())
        stream.truncate(size)
    return path


def snapshot(root):
    return {
        str(p.relative_to(root)): hashlib.sha256(p.read_bytes()).hexdigest() if p.is_file() else "DIR"
        for p in root.rglob("*")
    }


class ExclusiveFile:
    def __init__(self, path):
        self.path = path

    def __enter__(self):
        self.kernel = ctypes.WinDLL("kernel32", use_last_error=True)
        self.kernel.CreateFileW.argtypes = [ctypes.c_wchar_p, ctypes.c_uint32, ctypes.c_uint32,
                                            ctypes.c_void_p, ctypes.c_uint32, ctypes.c_uint32, ctypes.c_void_p]
        self.kernel.CreateFileW.restype = ctypes.c_void_p
        self.kernel.CloseHandle.argtypes = [ctypes.c_void_p]
        flags = 0x02000000 if self.path.is_dir() else 0x80  # FILE_FLAG_BACKUP_SEMANTICS for directories
        self.handle = self.kernel.CreateFileW(str(self.path), 0x80000000, 0, None, 3, flags, None)
        if self.handle == ctypes.c_void_p(-1).value:
            raise ctypes.WinError(ctypes.get_last_error())
        return self

    def __exit__(self, *args):
        self.kernel.CloseHandle(self.handle)


class OrganizerTests(unittest.TestCase):
    def setUp(self):
        self.restore_recycled = False
        self.temp_path = SCRIPT.parents[2] / ("wechat-organizer-test-" + uuid.uuid4().hex + " 中文 空格")
        self.temp_path.mkdir()
        self.root = self.temp_path / "file"
        self.root.mkdir()
        self.output = self.root / "整理归纳"

    def make_symlink(self, *args, **kwargs):
        try:
            os.symlink(*args, **kwargs)
        except OSError as exc:
            if getattr(exc, "winerror", None) == 1314:
                self.skipTest("Windows symbolic-link privilege unavailable")
            raise

    def tearDown(self):
        assert self.temp_path.resolve().parent == SCRIPT.parents[2].resolve()
        assert self.temp_path.name.startswith("wechat-organizer-test-")
        if self.restore_recycled:
            restored = subprocess.run([POWERSHELL, "-NoProfile", "-NonInteractive", "-File",
                                       str(SCRIPT.parents[1] / "tests/restore-test-recycle.ps1"),
                                       "-TestRoot", str(self.temp_path)],
                                      capture_output=True, encoding="utf-8", env=POWERSHELL_ENV)
            self.assertEqual(restored.returncode, 0, restored.stderr)
        shutil.rmtree(self.temp_path)

    def discover_default_locations(self, profile, *extra, system_profile=None, check=True):
        # Substitute the OS profile lookup and USERPROFILE in a disposable copy;
        # exercise default discovery without reading real WeChat accounts.
        script = SCRIPT.read_text(encoding="utf-8-sig")
        lookup = "[Environment]::GetFolderPath('UserProfile')"
        self.assertEqual(script.count(lookup), 1)
        fallback = profile if system_profile is None else system_profile
        script = script.replace(lookup, "'" + str(fallback).replace("'", "''") + "'")
        test_script = self.temp_path / "discover-defaults.ps1"
        test_script.write_text(script, encoding="utf-8-sig")
        test_env = {k: v for k, v in POWERSHELL_ENV.items() if k.upper() != "USERPROFILE"}
        if profile is not None:
            test_env["USERPROFILE"] = str(profile)
        result = subprocess.run(
            [POWERSHELL, "-NoProfile", "-NonInteractive", "-File", str(test_script),
             "-ListAccounts", "-AsJson", *extra],
            capture_output=True, encoding="utf-8", env=test_env,
        )
        if not check:
            return result
        self.assertEqual(result.returncode, 0, result.stderr)
        return json.loads(result.stdout)

    def test_default_discovery_combines_both_locations_without_selecting(self):
        profile = self.temp_path / "用户目录"
        documents = profile / "Documents"
        roots = [profile / "xwechat_files", documents / "xwechat_files"]
        files = [put(base / "同名账号" / "msg" / "file" / "原件.docx") for base in roots]
        before = [snapshot(base) for base in roots]
        result = self.discover_default_locations(profile)
        self.assertTrue(result["Success"])
        self.assertEqual(result["SearchRoots"], [str(base) for base in roots])
        self.assertNotIn("SearchRoot", result)
        self.assertEqual(result["CandidateCount"], 2)
        self.assertEqual([c["RootPath"] for c in result["Candidates"]], [str(f.parent) for f in files])
        self.assertTrue(result["SelectionRequired"])
        self.assertEqual(before, [snapshot(base) for base in roots])

    def test_default_discovery_finds_either_location_alone(self):
        for location in ["profile", "documents"]:
            with self.subTest(location=location):
                profile = self.temp_path / location / "用户"
                documents = profile / "Documents"
                parent = profile if location == "profile" else documents
                source = put(parent / "xwechat_files" / "账号" / "msg" / "file" / "资料.pdf")
                result = self.discover_default_locations(profile)
                self.assertTrue(result["Success"])
                self.assertEqual(result["CandidateCount"], 1)
                self.assertFalse(result["SelectionRequired"])
                self.assertEqual(result["Candidates"][0]["RootPath"], str(source.parent))
                missing = (documents if location == "profile" else profile) / "xwechat_files"
                self.assertFalse(missing.exists())

    def test_default_discovery_missing_locations_does_not_create_them(self):
        profile = self.temp_path / "用户"
        documents = profile / "Documents"
        result = self.discover_default_locations(profile)
        self.assertEqual(result["Candidates"], [])
        self.assertTrue(result["SelectionRequired"])
        # PowerShell itself may initialize caches under the test USERPROFILE.
        self.assertFalse((profile / "xwechat_files").exists())
        self.assertFalse((documents / "xwechat_files").exists())

    def test_default_discovery_prefers_real_environment_over_sandbox_identity(self):
        profile = self.temp_path / "用户"
        put(profile / "xwechat_files" / "账号" / "msg" / "file" / "资料.pdf")
        sandbox = self.temp_path / "CodexSandboxOffline"
        put(sandbox / "xwechat_files" / "错误账号" / "msg" / "file" / "资料.pdf")
        result = self.discover_default_locations(profile, system_profile=sandbox)
        self.assertEqual(result["SearchRoots"], [str(profile / "xwechat_files"),
                                               str(profile / "Documents" / "xwechat_files")])
        self.assertEqual(result["CandidateCount"], 1)
        self.assertEqual(result["Candidates"][0]["AccountFolder"], "账号")

    def test_default_discovery_falls_back_to_real_system_profile(self):
        profile = self.temp_path / "用户"
        source = put(profile / "Documents" / "xwechat_files" / "账号" / "msg" / "file" / "资料.pdf")
        for environment_profile in [None, self.temp_path / "CodexSandboxOnline"]:
            with self.subTest(environment_profile=environment_profile):
                result = self.discover_default_locations(environment_profile, system_profile=profile)
                self.assertEqual(result["Candidates"][0]["RootPath"], str(source.parent))

    def test_default_discovery_refuses_sandbox_only_profiles(self):
        sandbox = self.temp_path / "CodexSandboxOffline"
        source = put(sandbox / "xwechat_files" / "账号" / "msg" / "file" / "资料.pdf")
        result = self.discover_default_locations(sandbox, check=False)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("无法识别真实用户目录", result.stderr)
        self.assertTrue(source.exists())

    def test_explicit_discovery_path_replaces_both_defaults(self):
        profile = self.temp_path / "用户"
        documents = profile / "Documents"
        for parent in [profile, documents]:
            put(parent / "xwechat_files" / "默认账号" / "msg" / "file" / "资料.pdf")
        custom = self.temp_path / "自定义存储"
        source = put(custom / "指定账号" / "msg" / "file" / "资料.pdf")
        result = self.discover_default_locations(profile, "-WechatBasePath", str(custom))
        self.assertEqual(result["SearchRoots"], [str(custom)])
        self.assertNotIn("SearchRoot", result)
        self.assertEqual(result["Candidates"], [{"AccountFolder": "指定账号", "RootPath": str(source.parent)}])

    def test_trailing_separators_preserve_paths_and_discovery(self):
        source = put(self.root / "中文.docx")
        result = run(str(self.root) + "\\")
        self.assertTrue(result["Success"])
        self.assertEqual(result["RootPath"], str(self.root))
        self.assertEqual(result["FileCount"], 1)
        self.assertTrue(source.exists())
        base = self.temp_path / "wechat"
        put(base / "账号" / "msg" / "file" / "样本.pdf")
        found = discover(str(base) + "\\")
        self.assertEqual(found["SearchRoots"], [str(base)])
        self.assertEqual(found["CandidateCount"], 1)

    def test_discovery_lists_accounts_without_modifying_or_selecting_them(self):
        base = self.temp_path / "xwechat_files"
        first = put(base / "账号A" / "msg" / "file" / "月份" / "原件.docx")
        second = put(base / "账号B" / "msg" / "file" / "原件.pdf")
        (base / "all_users").mkdir()
        (base / "Backup").mkdir()
        before = snapshot(base)
        result = discover(base)
        self.assertTrue(result["Success"])
        self.assertEqual(result["Mode"], "SelectAccount")
        self.assertTrue(result["SelectionRequired"])
        self.assertEqual(result["CandidateCount"], 2)
        self.assertEqual([r["AccountFolder"] for r in result["Candidates"]], ["账号A", "账号B"])
        self.assertEqual(before, snapshot(base))
        selected = run(first.parents[1], apply=True)
        self.assertTrue(selected["Success"])
        self.assertTrue(second.exists())
        self.assertFalse((second.parent / "整理归纳").exists())

    def test_discovery_one_account_does_not_require_separate_selection(self):
        base = self.temp_path / "one-account"
        source = put(base / "only" / "msg" / "file" / "原件.docx")
        result = discover(base)
        self.assertEqual(result["CandidateCount"], 1)
        self.assertFalse(result["SelectionRequired"])
        self.assertTrue(source.exists())
        self.assertFalse((source.parent / "整理归纳").exists())

    def test_discovery_missing_or_empty_location_does_not_create_directories(self):
        missing = self.temp_path / "missing"
        result = discover(missing)
        self.assertEqual(result["Candidates"], [])
        self.assertTrue(result["SelectionRequired"])
        self.assertFalse(missing.exists())
        missing.mkdir()
        self.assertEqual(discover(missing)["Candidates"], [])

    def test_missing_root_cannot_preview_or_apply_implicit_account(self):
        before = snapshot(self.root)
        for extra in [[], ["-Apply"]]:
            result = subprocess.run([POWERSHELL, "-NoProfile", "-NonInteractive", "-File", str(SCRIPT), *extra],
                                    capture_output=True, encoding="utf-8", cwd=self.root, env=POWERSHELL_ENV)
            self.assertNotEqual(result.returncode, 0)
            self.assertIn("未选择微信账号目录", result.stderr)
        self.assertEqual(before, snapshot(self.root))

    def test_discovery_skips_linked_accounts_and_msg_directories(self):
        base = self.temp_path / "links"
        base.mkdir()
        outside_account = self.temp_path / "outside-account"
        put(outside_account / "msg" / "file" / "原件.docx")
        account_link = base / "linked-account"
        self.make_symlink(outside_account, account_link, target_is_directory=True)
        ordinary = base / "ordinary-account"
        ordinary.mkdir()
        msg_link = ordinary / "msg"
        self.make_symlink(outside_account / "msg", msg_link, target_is_directory=True)
        try:
            result = discover(base)
            self.assertTrue(result["Success"])
            self.assertEqual(result["Candidates"], [])
            self.assertEqual(set(result["SkippedLinks"]), {str(account_link), str(msg_link)})
        finally:
            account_link.unlink()
            msg_link.unlink()

    def test_preview_has_no_side_effects_and_matches_apply(self):
        put(self.root / "2026-01" / "深层" / "报告 [修订] '中文.DOCX")
        put(self.root / "根文件", b"abc")
        before = snapshot(self.root)
        preview = run(self.root)
        self.assertTrue(preview["Success"])
        self.assertEqual(preview["FileCount"], 2)
        self.assertEqual(before, snapshot(self.root))
        self.assertFalse(self.output.exists())
        applied = run(self.root, apply=True)
        self.assertEqual(preview["CategoryCounts"], applied["CategoryCounts"])
        self.assertEqual(preview["Run"]["Statistics"], applied["Inventory"]["Statistics"])
        self.assertEqual(preview["Run"]["TotalSize"], applied["Inventory"]["TotalSize"])
        self.assertTrue((self.root / "2026-01" / "深层").is_dir())
        self.assertTrue((self.output / "其他" / "根文件").exists())

    def test_json_records_are_opt_in_and_leave_summary_unchanged(self):
        source = put(self.root / "月份" / "同名.docx", b"new")
        put(self.output / "word" / "同名.docx", b"existing")
        before = snapshot(self.root)
        summary = run(self.root)
        detailed = run(self.root, include_records=True)
        self.assertNotIn("Records", summary)
        records = detailed.pop("Records")
        self.assertEqual(summary, detailed)
        self.assertEqual(before, snapshot(self.root))
        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]["Source"], str(source))
        self.assertEqual(records[0]["Status"], "Planned")
        self.assertTrue(records[0]["Renamed"])

        applied = run(self.root, apply=True, include_records=True)
        self.assertEqual(applied["Records"][0]["Status"], "Moved")
        self.assertEqual(applied["Records"][0]["Destination"], records[0]["Destination"])
        self.assertEqual(summary["CategoryCounts"], applied["CategoryCounts"])
        self.assertNotIn("Records", run(self.root, apply=True))
        self.assertEqual(run(self.root, include_records=True)["Records"], [])

    def test_preview_sizes_use_kb_mb_gb_without_writing(self):
        samples = {
            "word": ("资料.docx", 1023 * 1024, "1023.00 KB"),
            "excel": ("资料.xlsx", 1024 ** 2, "1.00 MB"),
            "ppt": ("资料.pptx", 1536 * 1024, "1.50 MB"),
            "压缩包": ("资料.zip", 1024 ** 3, "1.00 GB"),
            "DWG": ("资料.dwg", 1024 ** 3 - 1, "1024.00 MB"),
            "PDF": ("空文件.pdf", 0, "0.00 KB"),
            "其他": ("零碎字节.bin", 3, "0.00 KB"),
        }
        for filename, size, _ in samples.values():
            put_sparse(self.root / "月份" / filename, size)
        before = {str(p): (p.stat().st_size, p.stat().st_mtime_ns) for p in self.root.rglob("*")}
        preview = run(self.root)
        text = run(self.root, as_json=False)
        self.assertTrue(preview["Success"])
        self.assertTrue(preview["Run"]["Complete"])
        self.assertEqual(preview["Run"]["TotalFiles"], 7)
        stats = {row["Folder"]: row for row in preview["Run"]["Statistics"]}
        for category, (_, size, display) in samples.items():
            self.assertEqual(stats[category]["FileCount"], 1)
            self.assertEqual(stats[category]["Size"], display)
            self.assertNotIn("SizeKB", stats[category])
            self.assertRegex(text, rf"{re.escape(category)}\s+\|\s+1 个\s+\|\s+{re.escape(display)}")
        self.assertEqual(preview["Run"]["TotalSize"], f"{sum(s[1] for s in samples.values()) / 1024 ** 3:.2f} GB")
        self.assertRegex(text, r"总计\s+\|\s+7 个\s+\|\s+2\.00 GB")
        self.assertIsNone(preview["ReportPath"])
        self.assertFalse(self.output.exists())
        self.assertEqual(before, {str(p): (p.stat().st_size, p.stat().st_mtime_ns) for p in self.root.rglob("*")})

    def test_preview_total_uses_raw_sizes_and_excludes_already_organized_files(self):
        put(self.root / "一月" / "同名.docx", b"a" * (768 * 1024))
        put(self.root / "二月" / "资料.pdf", b"b" * (768 * 1024))
        put(self.output / "word" / "同名.docx", b"existing")
        put(self.output / "自定义" / "资料.bin", b"excluded")
        before = snapshot(self.root)
        result = run(self.root)
        self.assertEqual(result["Run"]["TotalFiles"], 2)
        self.assertEqual(result["Run"]["TotalSize"], "1.50 MB")
        self.assertEqual(result["RenamedCount"], 1)
        stats = {row["Folder"]: row for row in result["Run"]["Statistics"]}
        self.assertEqual(stats["word"]["Size"], "768.00 KB")
        self.assertEqual(stats["PDF"]["Size"], "768.00 KB")
        self.assertNotIn("excel", stats)
        self.assertNotIn("自定义", stats)
        self.assertEqual(before, snapshot(self.root))

    def test_empty_preview_hides_zero_rows_and_has_no_output(self):
        result = run(self.root)
        self.assertEqual(result["Run"]["TotalFiles"], 0)
        self.assertEqual(result["Run"]["TotalSize"], "0.00 KB")
        self.assertTrue(result["Run"]["Complete"])
        self.assertEqual(result["Run"]["Statistics"], [])
        self.assertTrue(all(r["FileCount"] == 0 and r["Size"] == "0.00 KB" for r in result["Run"]["Statistics"]))
        self.assertFalse(self.output.exists())

    def test_preview_marks_unreadable_directory_as_incomplete(self):
        blocked = put(self.root / "暂不可读" / "资料.pdf", b"hidden").parent
        put(self.root / "可读.docx", b"x" * 2048)
        with ExclusiveFile(blocked):
            result = run(self.root)
            text = run(self.root, as_json=False)
        self.assertFalse(result["Success"])
        self.assertFalse(result["Run"]["Complete"])
        self.assertEqual(result["Run"]["TotalFiles"], 1)
        self.assertEqual(result["Run"]["TotalSize"], "2.00 KB")
        self.assertIn("预览统计不完整", text)
        self.assertTrue(any(i["Stage"] == "Scan" for i in result["Failures"]))
        self.assertFalse(self.output.exists())

    def test_text_records_are_opt_in_for_preview_and_apply(self):
        source = put(self.root / "月份" / "报告.docx")
        summary = run(self.root, as_json=False)
        detailed = run(self.root, as_json=False, include_records=True)
        self.assertNotIn(str(source), summary)
        self.assertRegex(summary, r"word\s+\|\s+1 个\s+\|\s+1\.00 KB")
        self.assertIn(f"[Planned] {source} -> ", detailed)
        self.assertEqual(summary, "\n".join(line for line in detailed.splitlines()
                                         if " -> " not in line) + "\n")
        applied = run(self.root, apply=True, as_json=False, include_records=True)
        self.assertIn(f"[Moved] {source} -> ", applied)
        second = put(self.root / "第二份.docx")
        applied_summary = run(self.root, apply=True, as_json=False)
        self.assertNotIn(str(second), applied_summary)
        self.assertIn("统计报告：", applied_summary)

    def test_every_extension_duplicates_statistics_and_repeat(self):
        expected_counts = {}
        expected_bytes = {}
        for category, extensions in EXTENSIONS.items():
            expected_counts[category] = 0
            expected_bytes[category] = 0
            for index, ext in enumerate(extensions.split()):
                data = bytes([index + 1]) * (1024 + index)
                put(self.root / "月份" / "嵌套" / (f"文件[{index}]" + ext.upper()), data)
                expected_counts[category] += 1
                expected_bytes[category] += len(data)
        for month in ["一月", "二月"]:
            put(self.root / month / "同名.docx", b"a" * 2048)
        put(self.output / "word" / "同名.docx", b"old")
        expected_counts["word"] += 3
        expected_bytes["word"] += 4099
        extra = put(self.output / "自定义" / "内层" / "自定义.bin", b"z" * 3072)
        put(self.output / "existing.txt", b"excluded root report-like file")
        legacy_csv = put(self.output / "移动记录_20261021_025643_309_7e40b010.csv", b"legacy audit")
        contents_before = collections.Counter(
            hashlib.sha256(p.read_bytes()).hexdigest()
            for p in self.root.rglob("*") if p.is_file()
        )
        result = run(self.root, apply=True)
        self.assertTrue(result["Success"], result["Failures"])
        self.assertEqual(result["RenamedCount"], 2)
        self.assertEqual((self.output / "word" / "同名.docx").read_bytes(), b"old")
        self.assertTrue((self.output / "word" / "同名 (1).docx").is_file())
        self.assertTrue((self.output / "word" / "同名 (2).docx").is_file())
        stats = {r["Folder"]: r for r in result["Inventory"]["Statistics"]}
        for category in CATEGORIES:
            self.assertEqual(stats[category]["FileCount"], expected_counts[category])
            self.assertEqual(stats[category]["Size"], f"{expected_bytes[category] / 1024:.2f} KB")
        self.assertEqual(stats["自定义"]["FileCount"], 1)
        self.assertEqual(stats["自定义"]["Size"], "3.00 KB")
        report_path = Path(result["ReportPath"])
        self.assertTrue(report_path.read_bytes().startswith(b"\xef\xbb\xbf"))
        report = report_path.read_text(encoding="utf-8-sig")
        self.assertNotIn("字节", report)
        self.assertNotIn("MiB", report)
        self.assertNotIn("CSV", report)
        self.assertRegex(report, r"统计时间：\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}（北京时间）")
        total_count = sum(expected_counts.values()) + 1
        total_kb = (sum(expected_bytes.values()) + extra.stat().st_size) / 1024
        self.assertRegex(report, rf"总计\s+\|\s+{total_count} 个\s+\|\s+{total_kb:.2f} KB")
        moved_contents = collections.Counter(
            hashlib.sha256(p.read_bytes()).hexdigest()
            for p in self.output.rglob("*") if p.is_file()
            and p != report_path
        )
        self.assertEqual(contents_before, moved_contents)
        repeated = run(self.root, apply=True)
        self.assertTrue(repeated["Success"], repeated["Failures"])
        self.assertEqual(repeated["FileCount"], 0)
        self.assertEqual(repeated["Inventory"]["Statistics"], result["Inventory"]["Statistics"])
        self.assertEqual(repeated["ReportPath"], result["ReportPath"])
        self.assertEqual(len(list(self.output.glob("整理统计.txt"))), 1)
        self.assertEqual(list(self.output.glob("移动记录_*.csv")), [legacy_csv])
        self.assertEqual(legacy_csv.read_bytes(), b"legacy audit")
        self.assertEqual(run(self.root)["FileCount"], 0)

    def test_only_txt_report_is_generated_without_csv_logs(self):
        put(self.root / "月份" / "用户数据.csv", b"name,value\nitem,1\n")
        first = run(self.root, apply=True)
        self.assertTrue(first["Success"])
        self.assertTrue((self.output / "excel" / "用户数据.csv").is_file())
        self.assertNotIn("LogPath", first)
        self.assertEqual([p.name for p in self.output.iterdir() if p.is_file()], ["整理统计.txt"])
        self.assertFalse(list(self.output.glob("移动记录_*.csv")))
        repeated = run(self.root, apply=True)
        self.assertTrue(repeated["Success"], repeated["Failures"])
        self.assertEqual(repeated["FileCount"], 0)
        self.assertEqual(first["Inventory"]["Statistics"], repeated["Inventory"]["Statistics"])
        self.assertEqual([p.name for p in self.output.iterdir() if p.is_file()], ["整理统计.txt"])
        self.assertFalse(list(self.output.glob("移动记录_*.csv")))

    def helper(self, body, **data):
        config = self.temp_path / "helper.json"
        config.write_text(json.dumps(dict(Script=str(SCRIPT), Root=str(self.root), **data)), encoding="utf-8-sig")
        script = self.temp_path / "helper.ps1"
        script.write_text('param([string]$ConfigPath)\n$ErrorActionPreference="Stop"\n'
                          '[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)\n'
                          '$cfg=Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json\n'
                          '. $cfg.Script -ListAccounts -WechatBasePath (Join-Path $cfg.Root "absent") -AsJson | Out-Null\n'
                          '. (Join-Path (Split-Path -Parent $cfg.Script) "deduplicate.ps1")\n' + body,
                          encoding="utf-8-sig")
        response = subprocess.run([POWERSHELL, "-NoProfile", "-NonInteractive", "-File", str(script), str(config)],
                                  capture_output=True, encoding="utf-8", env=POWERSHELL_ENV)
        self.assertEqual(response.returncode, 0, response.stderr)
        return json.loads(response.stdout)

    def test_audio_extensions_and_opt_in_audio_migration_are_previewed(self):
        for index, ext in enumerate(EXTENSIONS["音频"].split()):
            put(self.root / "月份" / f"audio-{index}{ext.upper()}", bytes([index]))
        old = put(self.output / "其他" / "archive" / "saved.MP3", b"old audio")
        put(self.output / "音频" / "saved.MP3", b"different")
        before = snapshot(self.root)
        preview = run(self.root, migrate_audio=True, deduplicate=True)
        self.assertTrue(preview["Success"], preview["Failures"])
        self.assertEqual(preview["CategoryCounts"]["音频"], 18)
        self.assertEqual(before, snapshot(self.root))
        result = run(self.root, migrate_audio=True, apply=True)
        self.assertTrue(result["Success"], result["Failures"])
        self.assertFalse(old.exists())
        self.assertEqual((self.output / "音频" / "saved (1).MP3").read_bytes(), b"old audio")
        self.assertEqual([r["Folder"] for r in result["Inventory"]["Statistics"]], ["音频"])
        self.assertEqual(result["Inventory"]["Statistics"][0]["FileCount"], 19)
        self.assertTrue((self.output / "其他" / "archive").is_dir())
        self.assertFalse((self.output / "PDF").exists())

    def test_only_empty_builtin_roots_are_removed_and_zero_byte_files_remain(self):
        (self.output / "word").mkdir(parents=True)
        (self.output / "PDF" / "empty-child").mkdir(parents=True)
        (self.output / "custom-empty").mkdir()
        month = self.root / "empty-month"
        month.mkdir()
        put(self.root / "empty.mp3", b"")
        result = run(self.root, apply=True)
        self.assertTrue(result["Success"], result["Failures"])
        self.assertFalse((self.output / "word").exists())
        self.assertTrue((self.output / "PDF" / "empty-child").exists())
        self.assertTrue((self.output / "custom-empty").exists())
        self.assertTrue(month.exists())
        self.assertEqual(result["Inventory"]["Statistics"], [{"Folder": "音频", "FileCount": 1, "Size": "0.00 KB"}])
        self.assertEqual(len(result["CategoryCounts"]), 8)

    def test_preview_dedup_respects_category_boundary_and_keeps_existing(self):
        new = put(self.root / "older-new.docx", b"identical")
        old = put(self.output / "word" / "keep.docx", b"identical")
        put(self.output / "word" / "copy.docx", b"identical")
        self.helper("[IO.File]::SetCreationTimeUtc($cfg.Keeper,[datetime]'2000-01-01'); 'true'", Keeper=str(old))
        put(self.output / "PDF" / "same.pdf", b"identical")
        put(self.output / "custom" / "one.docx", b"identical")
        put(self.output / "custom" / "two.docx", b"identical")
        put(self.output / "root-record.csv", b"identical")
        put(self.root / "different.docx", b"different")
        before = snapshot(self.root)
        preview = run(self.root, deduplicate=True, include_records=True)
        self.assertTrue(preview["Success"], preview["Failures"])
        self.assertEqual(preview["DuplicateCount"], 2)
        self.assertEqual(preview["RecycledCount"], 0)
        self.assertTrue(preview["DuplicateCheckComplete"])
        planned = [r for r in preview["Records"] if r["Status"] == "DuplicatePlanned"]
        self.assertTrue(all(r["KeptPath"] == str(old) for r in planned))
        self.assertTrue(any(r["Source"] == str(new) for r in planned))
        self.assertEqual(before, snapshot(self.root))
        self.assertIsNone(preview["ReportPath"])

    def test_apply_dedup_recycles_history_and_new_copies_and_is_repeatable(self):
        put(self.root / "oldest-new.docx", b"duplicate data")
        kept = put(self.output / "word" / "keep.docx", b"duplicate data")
        put(self.output / "word" / "history-one.docx", b"duplicate data")
        put(self.output / "word" / "history-two.docx", b"duplicate data")
        self.helper("[IO.File]::SetCreationTimeUtc($cfg.Keeper,[datetime]'2000-01-01'); 'true'", Keeper=str(kept))
        put(self.root / "unique.docx", b"unique")
        self.restore_recycled = True
        result = run(self.root, apply=True, deduplicate=True, include_records=True)
        self.assertTrue(result["Success"], result["Failures"])
        self.assertEqual(result["DuplicateCount"], 3)
        self.assertEqual(result["RecycledCount"], 3)
        self.assertEqual(result["Run"]["TotalFiles"], 2)
        self.assertEqual(result["Inventory"]["TotalFiles"], 2)
        self.assertEqual(kept.read_bytes(), b"duplicate data")
        self.assertEqual(len(list((self.output / "word").iterdir())), 2)
        details = [r for r in result["Records"] if r["Status"] == "Recycled"]
        self.assertEqual(len(details), 3)
        self.assertTrue(all(r["KeptPath"] == str(kept) for r in details))
        report = Path(result["ReportPath"]).read_text(encoding="utf-8-sig")
        self.assertIn("3 个重复副本移入回收站", report)
        repeated = run(self.root, apply=True, deduplicate=True)
        self.assertTrue(repeated["Success"], repeated["Failures"])
        self.assertEqual(repeated["DuplicateCount"], 0)
        self.assertEqual(repeated["RecycledCount"], 0)
        self.assertEqual(result["Inventory"]["Statistics"], repeated["Inventory"]["Statistics"])

    def test_creation_time_tie_uses_full_path_and_empty_files_are_duplicates(self):
        a = put(self.output / "word" / "a.docx", b"")
        z = put(self.output / "word" / "z.docx", b"")
        self.helper("[IO.File]::SetCreationTimeUtc($cfg.A,[datetime]'2000-01-01'); [IO.File]::SetCreationTimeUtc($cfg.Z,[datetime]'2000-01-01'); 'true'", A=str(a), Z=str(z))
        result = run(self.root, deduplicate=True, include_records=True)
        self.assertTrue(result["Success"], result["Failures"])
        self.assertEqual(result["DuplicateCount"], 1)
        self.assertEqual(result["DuplicateSize"], "0.00 KB")
        record = next(r for r in result["Records"] if r["Status"] == "DuplicatePlanned")
        self.assertEqual(record["KeptPath"], str(a))
        self.assertEqual(record["Source"], str(z))
        self.assertTrue(a.exists() and z.exists())

    def test_existing_audio_is_preferred_over_new_duplicate_audio(self):
        put(self.root / "new.mp3", b"same audio")
        put(self.output / "其他" / "existing.mp3", b"same audio")
        self.restore_recycled = True
        result = run(self.root, migrate_audio=True, apply=True, deduplicate=True, include_records=True)
        self.assertTrue(result["Success"], result["Failures"])
        self.assertEqual(result["RecycledCount"], 1)
        self.assertEqual(result["Inventory"]["TotalFiles"], 1)
        self.assertEqual((self.output / "音频" / "existing.mp3").read_bytes(), b"same audio")
        self.assertFalse((self.output / "其他").exists())
        self.assertFalse((self.output / "音频" / "new.mp3").exists())

    def test_unavailable_native_interface_keeps_copies_and_reports_failure(self):
        kept = put(self.output / "word" / "keep.docx", b"same")
        copy = put(self.output / "word" / "copy.docx", b"same")
        package = self.temp_path / "broken-interface-package"
        shutil.copytree(SCRIPT.parents[1] / "scripts", package)
        (package / "recycle-native.cs").write_text("invalid C sharp source", encoding="utf-8")
        response = subprocess.run([POWERSHELL, "-NoProfile", "-NonInteractive", "-File", str(package / "organize.ps1"),
                                   "-RootPath", str(self.root), "-Apply", "-Deduplicate", "-AsJson"],
                                  capture_output=True, encoding="utf-8", env=POWERSHELL_ENV)
        self.assertEqual(response.returncode, 0, response.stderr)
        result = json.loads(response.stdout)
        self.assertFalse(result["Success"])
        self.assertEqual(result["RecycledCount"], 0)
        self.assertEqual(result["RecycleFailedCount"], 1)
        self.assertEqual(result["Inventory"]["TotalFiles"], 2)
        self.assertTrue(kept.exists() and copy.exists())
        self.assertIn("回收失败 1 个", Path(result["ReportPath"]).read_text(encoding="utf-8-sig"))

    def test_plain_move_never_deduplicates_and_cross_category_files_are_preserved(self):
        put(self.root / "first.docx", b"same")
        put(self.root / "second.docx", b"same")
        plain = run(self.root, apply=True)
        self.assertEqual(plain["RecycledCount"], 0)
        self.assertEqual(plain["Inventory"]["TotalFiles"], 2)
        (self.output / "word" / "second.docx").unlink()
        put(self.output / "PDF" / "same.pdf", b"same")
        put(self.output / "custom" / "a.docx", b"same")
        put(self.output / "custom" / "b.docx", b"same")
        result = run(self.root, apply=True, deduplicate=True)
        self.assertTrue(result["Success"], result["Failures"])
        self.assertEqual(result["DuplicateCount"], 0)
        self.assertEqual(result["Inventory"]["TotalFiles"], 4)

    def test_unreadable_duplicate_and_changed_files_are_retained(self):
        kept = put(self.output / "word" / "keep.docx", b"same")
        duplicate = put(self.output / "word" / "copy.docx", b"same")
        with ExclusiveFile(duplicate):
            preview = run(self.root, deduplicate=True)
        self.assertFalse(preview["Success"])
        self.assertFalse(preview["DuplicateCheckComplete"])
        self.assertEqual(preview["RecycledCount"], 0)
        result = self.helper(r'''
$entries=@((New-DuplicateEntry (Get-Item -LiteralPath $cfg.Keeper) 'word' $true), (New-DuplicateEntry (Get-Item -LiteralPath $cfg.Candidate) 'word' $true))
$pairs=@(Get-DuplicatePlan $entries)
[IO.File]::WriteAllText($pairs[0].Candidate.Path,'changed')
$script:called=$false
$failed=$false
try { $null=Invoke-DuplicateRecycle $pairs[0] -RecycleAction {param($p) $script:called=$true; [pscustomobject]@{Recycled=$true}} } catch {$failed=$true}
[pscustomobject]@{Failed=$failed;Called=$script:called;KeptExists=(Test-Path -LiteralPath $pairs[0].Keeper.Path);CandidateExists=(Test-Path -LiteralPath $pairs[0].Candidate.Path)} | ConvertTo-Json
''', Keeper=str(kept), Candidate=str(duplicate))
        self.assertTrue(result["Failed"])
        self.assertFalse(result["Called"])
        self.assertTrue(result["KeptExists"] and result["CandidateExists"])

    def test_recycle_refusal_cancellation_and_unavailable_backend_retain_files(self):
        kept = put(self.output / "word" / "keep.docx", b"same")
        duplicate = put(self.output / "word" / "copy.docx", b"same")
        result = self.helper(r'''
$pairs=@(Get-DuplicatePlan @((New-DuplicateEntry (Get-Item -LiteralPath $cfg.Keeper) 'word' $true), (New-DuplicateEntry (Get-Item -LiteralPath $cfg.Candidate) 'word' $true)))
$failures=0
foreach($action in @({param($p) throw 'interface unavailable'}, {param($p) [pscustomobject]@{Recycled=$false;Error='cancelled'}}, {param($p) [pscustomobject]@{Recycled=$false;Error='cannot recycle'}})) {
    try {$null=Invoke-DuplicateRecycle $pairs[0] -RecycleAction $action} catch {$failures++}
}
Add-Type -Path (Join-Path (Split-Path -Parent $cfg.Script) 'recycle-native.cs')
$sink=[WechatOrganizerRecycleV1.NativeRecycle+Sink]::new($cfg.Candidate,$pairs[0].Candidate.Hash,4,0,0)
$veto=$sink.PreDeleteItem(0,$null)
$sink.Dispose()
[pscustomobject]@{Failures=$failures;Veto=$veto;KeptExists=(Test-Path -LiteralPath $cfg.Keeper);CandidateExists=(Test-Path -LiteralPath $cfg.Candidate)} | ConvertTo-Json
''', Keeper=str(kept), Candidate=str(duplicate))
        self.assertEqual(result["Failures"], 3)
        self.assertLess(result["Veto"], 0)
        self.assertTrue(result["KeptExists"] and result["CandidateExists"])

    def test_empty_run_and_rounded_fractional_kb(self):
        empty = run(self.root, apply=True)
        self.assertTrue(empty["Success"])
        self.assertEqual(empty["Inventory"]["Statistics"], [])
        self.assertTrue(all(r["FileCount"] == 0 and r["Size"] == "0.00 KB" for r in empty["Inventory"]["Statistics"]))
        put(self.root / "fraction.pdf", b"1234567890")
        result = run(self.root, apply=True)
        pdf = next(r for r in result["Inventory"]["Statistics"] if r["Folder"] == "PDF")
        self.assertEqual(pdf["Size"], "0.01 KB")

    def test_pdf_category_handles_uppercase_duplicates_and_repeated_runs(self):
        put(self.output / "PDF" / "报告.pdf", b"existing")
        put(self.root / "一月" / "报告.pdf", b"first")
        put(self.root / "二月" / "报告.PDF", b"second")
        before = snapshot(self.root)
        preview = run(self.root)
        self.assertTrue(preview["Success"])
        self.assertEqual(preview["CategoryCounts"]["PDF"], 2)
        self.assertEqual(before, snapshot(self.root))
        result = run(self.root, apply=True)
        self.assertTrue(result["Success"], result["Failures"])
        self.assertEqual(result["RenamedCount"], 2)
        self.assertEqual(result["CategoryCounts"]["PDF"], 2)
        contents = {path.read_bytes() for path in (self.output / "PDF").iterdir()}
        self.assertEqual(contents, {b"existing", b"first", b"second"})
        self.assertEqual((self.output / "PDF" / "报告.pdf").read_bytes(), b"existing")
        self.assertFalse((self.output / "其他").exists())
        stats = next(row for row in result["Inventory"]["Statistics"] if row["Folder"] == "PDF")
        self.assertEqual(stats["FileCount"], 3)
        self.assertRegex(Path(result["ReportPath"]).read_text(encoding="utf-8-sig"), r"PDF\s+\|\s+3 个\s+\|")
        repeated = run(self.root, apply=True)
        self.assertEqual(repeated["FileCount"], 0)
        self.assertEqual(repeated["Inventory"]["Statistics"], result["Inventory"]["Statistics"])

    def test_report_automatically_selects_kb_mb_and_gb(self):
        sizes = {
            "word": 0,
            "excel": 1023 * 1024,
            "ppt": 1024 ** 2,
            "压缩包": 1536 * 1024,
            "DWG": 1023 * 1024 ** 2,
            "PDF": 0,
            "其他": 1024 ** 3,
        }
        expected = {
            "word": "0.00 KB", "excel": "1023.00 KB", "ppt": "1.00 MB",
            "压缩包": "1.50 MB", "DWG": "1023.00 MB", "PDF": "0.00 KB", "其他": "1.00 GB",
        }
        for category, size in sizes.items():
            if size:
                put_sparse(self.output / category / "大小测试.bin", size)
        result = run(self.root, apply=True)
        self.assertTrue(result["Success"], result["Failures"])
        stats = {row["Folder"]: row for row in result["Inventory"]["Statistics"]}
        report = Path(result["ReportPath"]).read_text(encoding="utf-8-sig")
        for category, display in expected.items():
            with self.subTest(category=category):
                if sizes[category] == 0:
                    self.assertNotIn(category, stats)
                    continue
                self.assertEqual(stats[category]["Size"], display)
                self.assertNotIn("SizeKB", stats[category])
                self.assertRegex(report, rf"(?m)^{re.escape(category)}\s+\|\s+\d+ 个\s+\|\s+{re.escape(display)}$")
        self.assertRegex(report, r"总计\s+\|\s+5 个\s+\|\s+2\.00 GB")
        self.assertNotIn("字节", report)

    def test_total_chooses_its_own_unit_from_original_bytes(self):
        put(self.output / "word" / "一.docx", b"x" * (768 * 1024))
        put(self.output / "excel" / "二.xlsx", b"x" * (768 * 1024))
        result = run(self.root, apply=True)
        self.assertTrue(result["Success"])
        stats = {row["Folder"]: row for row in result["Inventory"]["Statistics"]}
        self.assertEqual(stats["word"]["Size"], "768.00 KB")
        self.assertEqual(stats["excel"]["Size"], "768.00 KB")
        report = Path(result["ReportPath"]).read_text(encoding="utf-8-sig")
        self.assertRegex(report, r"总计\s+\|\s+2 个\s+\|\s+1\.50 MB")
        put(self.output / "word" / "一.docx", b"a" * (1024 ** 2))
        put(self.output / "excel" / "二.xlsx", b"b")
        repeated = run(self.root, apply=True)
        self.assertTrue(repeated["Success"])
        changed_report = Path(repeated["ReportPath"]).read_text(encoding="utf-8-sig")
        self.assertRegex(changed_report, r"总计\s+\|\s+2 个\s+\|\s+1\.00 MB")
        self.assertEqual(repeated["ReportPath"], result["ReportPath"])

    def test_report_columns_align_for_chinese_long_names_and_different_sizes(self):
        names = ["很长的自定义资料分类-Quarterly-reports-2026", "资料e\u0301", "📁图纸"]
        for index, name in enumerate(names):
            put(self.output / name / "sample.bin", b"x" * (index + 1))
        for index in range(12):
            put(self.output / "word" / f"{index}.docx")
        put(self.output / "excel" / "大文件.xlsx", b"x" * (120 * 1024))
        result = run(self.root, apply=True)
        self.assertTrue(result["Success"], result["Failures"])
        report = Path(result["ReportPath"]).read_text(encoding="utf-8-sig")

        def width(text):
            return sum(0 if unicodedata.combining(c) else 2 if unicodedata.east_asian_width(c) in "WF" else 1 for c in text)

        rows = [line for line in report.splitlines() if " | " in line]
        self.assertEqual(len(rows), 1 + 2 + len(names) + 1)
        baseline = [width(cell) for cell in rows[0].split(" | ")]
        for row in rows:
            cells = row.split(" | ")
            self.assertEqual([width(cell) for cell in cells], baseline)
            self.assertFalse(cells[0].startswith(" "))
            self.assertEqual(cells[1], cells[1].rstrip())
            self.assertEqual(cells[2], cells[2].rstrip())
        for name in names:
            self.assertTrue(any(row.split(" | ")[0].rstrip() == name for row in rows))
        self.assertIn("新宋体", report)
        rules = [line for line in report.splitlines() if set(line) <= set("=-+") and line]
        self.assertTrue(rules)
        self.assertTrue(all(len(rule) == width(rows[0]) for rule in rules))

    def test_locked_source_is_retained_and_report_counts_actual_files(self):
        blocked = put(self.root / "月份" / "占用.docx", b"locked")
        put(self.root / "月份" / "正常.pdf", b"x" * 2048)
        with ExclusiveFile(blocked):
            result = run(self.root, apply=True)
            detailed = run(self.root, apply=True, include_records=True)
        self.assertFalse(result["Success"])
        self.assertEqual(result["FailedCount"], 1)
        self.assertEqual(result["Run"]["TotalFiles"], 1)
        self.assertEqual(result["Run"]["TotalSize"], "2.00 KB")
        self.assertFalse(result["Run"]["Complete"])
        self.assertTrue(result["Inventory"]["Complete"])
        self.assertTrue(blocked.exists())
        self.assertTrue((self.output / "PDF" / "正常.pdf").exists())
        self.assertEqual(sum(r["FileCount"] for r in result["Inventory"]["Statistics"]), 1)
        self.assertTrue(Path(result["ReportPath"]).exists())
        self.assertNotIn("Records", result)
        self.assertTrue(any(i["Stage"] == "Move" and i["Path"] == str(blocked) and i["Message"]
                            for i in result["Failures"]))
        self.assertTrue(any(r["Status"] == "Failed" and r["Error"] for r in detailed["Records"]))
        self.assertFalse(list(self.output.glob("移动记录_*.csv")))

    def test_locked_report_is_reported_as_failure(self):
        first = run(self.root, apply=True)
        with ExclusiveFile(Path(first["ReportPath"])):
            result = run(self.root, apply=True)
        self.assertFalse(result["Success"])
        self.assertIsNone(result["ReportPath"])
        self.assertTrue(any(i["Stage"] == "Report" for i in result["Failures"]))
        self.assertFalse(list(self.output.glob(".统计_*.tmp")))

    def test_unreadable_statistics_directory_marks_report_incomplete(self):
        blocked = put(self.output / "暂不可读" / "资料.bin", b"blocked").parent
        put(self.root / "正常.docx", b"visible")
        with ExclusiveFile(blocked):
            result = run(self.root, apply=True)
        self.assertFalse(result["Success"])
        self.assertFalse(result["Inventory"]["Complete"])
        self.assertTrue(result["Run"]["Complete"])
        self.assertTrue(any(i["Stage"] == "Statistics" for i in result["Failures"]))
        self.assertIn("统计不完整", Path(result["ReportPath"]).read_text(encoding="utf-8-sig"))
        self.assertTrue((blocked / "资料.bin").exists())
        self.assertTrue((self.output / "word" / "正常.docx").exists())

    def test_directory_links_are_skipped_and_linked_target_is_rejected(self):
        outside = self.temp_path / "outside"
        outside.mkdir()
        put(outside / "不能移动.docx", b"outside")
        link = self.root / "目录链接"
        self.make_symlink(outside, link, target_is_directory=True)
        try:
            result = run(self.root, apply=True)
            self.assertTrue(result["Success"])
            self.assertEqual(result["FileCount"], 0)
            self.assertIn(str(link), result["SkippedLinks"])
            self.assertTrue((outside / "不能移动.docx").exists())
            result = run(link, check=False)
            self.assertNotEqual(result.returncode, 0)
            target_link = self.output / "额外链接"
            self.make_symlink(outside, target_link, target_is_directory=True)
            try:
                stats = run(self.root, apply=True)
                self.assertNotIn("额外链接", [r["Folder"] for r in stats["Inventory"]["Statistics"]])
            finally:
                target_link.unlink()
            if (self.output / "word").exists():
                (self.output / "word").rmdir()
            self.make_symlink(outside, self.output / "word", target_is_directory=True)
            try:
                rejected = run(self.root, apply=True, check=False)
                self.assertNotEqual(rejected.returncode, 0)
                self.assertTrue((outside / "不能移动.docx").exists())
            finally:
                (self.output / "word").unlink()
        finally:
            link.unlink()

    def test_file_symlink_is_skipped(self):
        outside = put(self.temp_path / "external.docx")
        link = self.root / "链接.docx"
        try:
            self.make_symlink(outside, link)
        except OSError as exc:
            self.skipTest(f"Windows file-symlink privilege unavailable: {exc.winerror}")
        try:
            result = run(self.root, apply=True)
            self.assertEqual(result["FileCount"], 0)
            self.assertIn(str(link), result["SkippedLinks"])
            self.assertTrue(outside.exists())
        finally:
            link.unlink()

    def test_selected_root_and_discovery_reject_linked_ancestors(self):
        outside = self.temp_path / "outside"
        source = put(outside / "account" / "msg" / "file" / "原件.docx")
        link = self.temp_path / "linked-base"
        self.make_symlink(outside, link, target_is_directory=True)
        try:
            selected = link / "account" / "msg" / "file"
            rejected = run(selected, apply=True, check=False)
            self.assertNotEqual(rejected.returncode, 0)
            result = subprocess.run(
                [POWERSHELL, "-NoProfile", "-NonInteractive", "-File", str(SCRIPT),
                 "-ListAccounts", "-WechatBasePath", str(link / "account"), "-AsJson"],
                capture_output=True, encoding="utf-8", env=POWERSHELL_ENV,
            )
            self.assertNotEqual(result.returncode, 0)
            self.assertTrue(source.exists())
            self.assertFalse((source.parent / "整理归纳").exists())
        finally:
            link.unlink()

    def test_run_counts_are_separate_from_inventory_and_repeat(self):
        for index in range(100):
            put(self.output / "word" / f"old-{index}.docx", b"old")
        for index in range(10):
            put(self.root / f"new-{index}.pdf", b"newdata")
        before = snapshot(self.root)
        preview = run(self.root)
        self.assertEqual(preview["Run"]["TotalFiles"], 10)
        self.assertEqual(preview["Run"]["Phase"], "Planned")
        self.assertIsNone(preview["Inventory"])
        self.assertEqual(before, snapshot(self.root))
        result = run(self.root, apply=True)
        self.assertEqual(result["Run"]["Phase"], "Completed")
        self.assertEqual(result["Run"]["TotalFiles"], 10)
        self.assertEqual(result["Inventory"]["TotalFiles"], 110)
        self.assertEqual(result["Run"]["Statistics"], preview["Run"]["Statistics"])
        for key in ["Statistics", "StatisticsComplete", "TotalFiles", "TotalSize"]:
            self.assertNotIn(key, result)
        report = Path(result["ReportPath"]).read_text(encoding="utf-8-sig")
        self.assertIn("本次移动成功：10 个", report)
        self.assertRegex(report, r"总计\s+\|\s+110 个")
        repeated = run(self.root, apply=True)
        self.assertEqual(repeated["Run"]["TotalFiles"], 0)
        self.assertEqual(repeated["Inventory"]["TotalFiles"], 110)

    def test_legacy_audio_stays_put_by_default_even_with_dedup(self):
        old = put(self.output / "其他" / "old.mp3", b"same audio")
        put(self.root / "new.mp3", b"same audio")
        before = snapshot(self.root)
        preview = run(self.root, deduplicate=True)
        self.assertEqual(preview["Run"]["TotalFiles"], 1)
        self.assertEqual(preview["Run"]["AudioMigrationCount"], 0)
        self.assertEqual(preview["DuplicateCount"], 0)
        self.assertEqual(before, snapshot(self.root))
        result = run(self.root, apply=True, deduplicate=True)
        self.assertTrue(result["Success"], result["Failures"])
        self.assertFalse(result["MigrateLegacyAudio"])
        self.assertEqual(result["Run"]["TotalFiles"], 1)
        self.assertEqual(result["Inventory"]["TotalFiles"], 2)
        self.assertEqual(result["RecycledCount"], 0)
        self.assertEqual(old.read_bytes(), b"same audio")
        self.assertEqual((self.output / "音频" / "new.mp3").read_bytes(), b"same audio")
        planned = run(self.root, migrate_audio=True, deduplicate=True, include_records=True)
        self.assertEqual(planned["Run"]["TotalFiles"], 1)
        self.assertEqual(planned["Run"]["AudioMigrationCount"], 1)
        self.assertEqual(planned["DuplicateCount"], 1)
        self.assertTrue(old.exists())

    def test_existing_target_file_is_rejected_before_moving(self):
        source = put(self.root / "原件.docx")
        put(self.output / "word", b"not a directory")
        result = run(self.root, apply=True, check=False)
        self.assertNotEqual(result.returncode, 0)
        self.assertTrue(source.exists())


if __name__ == "__main__":
    # Policy is a prerequisite; the test harness must not silently relax it.
    policy = subprocess.run(
        [POWERSHELL, "-NoProfile", "-NonInteractive", "-Command", "Get-ExecutionPolicy"],
        capture_output=True, encoding="utf-8", errors="replace", env=POWERSHELL_ENV,
    )
    if policy.returncode:
        sys.exit("Cannot read PowerShell execution policy: " + policy.stderr)
    if policy.stdout.strip() == "Restricted":
        sys.exit("Execution policy is Restricted. Tests stopped without changing policy. "
                 "See README: explicit user approval is required before policy configuration.")
    unittest.main(verbosity=2)
