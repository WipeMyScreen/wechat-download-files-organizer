"""Integration tests against disposable files; never uses the real WeChat root."""
import collections
import ctypes
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import unittest
import uuid

SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "organize.ps1"
CATEGORIES = ["word", "excel", "ppt", "压缩包", "DWG", "其他"]
EXTENSIONS = {
    "word": ".doc .docx .docm .dot .dotx .dotm .wps .wpt",
    "excel": ".xls .xlsx .xlsm .xlsb .xlt .xltx .xltm .csv .et .ett",
    "ppt": ".ppt .pptx .pptm .pot .potx .potm .pps .ppsx .ppsm .dps .dpt",
    "压缩包": ".zip .rar .7z .tar .gz .tgz .bz2 .xz .zst .cab",
    "DWG": ".dwg",
    "其他": ".pdf .png .yml .dxf",
}


def run(root, apply=False, check=True):
    args = ["pwsh", "-NoProfile", "-File", str(SCRIPT), "-RootPath", str(root), "-AsJson"]
    if apply:
        args.append("-Apply")
    result = subprocess.run(args, capture_output=True, encoding="utf-8")
    if not check:
        return result
    if result.returncode:
        raise AssertionError(result.stderr)
    return json.loads(result.stdout)


def discover(base, *extra):
    result = subprocess.run(
        ["pwsh", "-NoProfile", "-File", str(SCRIPT), "-ListAccounts",
         "-WechatBasePath", str(base), "-AsJson", *extra],
        capture_output=True, encoding="utf-8",
    )
    if result.returncode:
        raise AssertionError(result.stderr)
    return json.loads(result.stdout)


def put(path, content=b"x" * 1024):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(content)
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
        self.handle = self.kernel.CreateFileW(str(self.path), 0x80000000, 0, None, 3, 0x80, None)
        if self.handle == ctypes.c_void_p(-1).value:
            raise ctypes.WinError(ctypes.get_last_error())
        return self

    def __exit__(self, *args):
        self.kernel.CloseHandle(self.handle)


class OrganizerTests(unittest.TestCase):
    def setUp(self):
        self.temp_path = SCRIPT.parents[2] / ("wechat-organizer-test-" + uuid.uuid4().hex)
        self.temp_path.mkdir()
        self.root = self.temp_path / "file"
        self.root.mkdir()
        self.output = self.root / "整理归纳"

    def tearDown(self):
        assert self.temp_path.resolve().parent == SCRIPT.parents[2].resolve()
        assert self.temp_path.name.startswith("wechat-organizer-test-")
        shutil.rmtree(self.temp_path)

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

    def test_discovery_one_account_still_requires_selection(self):
        base = self.temp_path / "one-account"
        source = put(base / "only" / "msg" / "file" / "原件.docx")
        result = discover(base)
        self.assertEqual(result["CandidateCount"], 1)
        self.assertTrue(result["SelectionRequired"])
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
            result = subprocess.run(["pwsh", "-NoProfile", "-File", str(SCRIPT), *extra],
                                    capture_output=True, encoding="utf-8", cwd=self.root)
            self.assertNotEqual(result.returncode, 0)
            self.assertIn("未选择微信账号目录", result.stderr)
        self.assertEqual(before, snapshot(self.root))

    def test_discovery_skips_linked_accounts_and_msg_directories(self):
        base = self.temp_path / "links"
        base.mkdir()
        outside_account = self.temp_path / "outside-account"
        put(outside_account / "msg" / "file" / "原件.docx")
        account_link = base / "linked-account"
        os.symlink(outside_account, account_link, target_is_directory=True)
        ordinary = base / "ordinary-account"
        ordinary.mkdir()
        msg_link = ordinary / "msg"
        os.symlink(outside_account / "msg", msg_link, target_is_directory=True)
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
        self.assertTrue((self.root / "2026-01" / "深层").is_dir())
        self.assertTrue((self.output / "其他" / "根文件").exists())

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
        stats = {r["Folder"]: r for r in result["Statistics"]}
        for category in CATEGORIES:
            self.assertEqual(stats[category]["FileCount"], expected_counts[category])
            self.assertEqual(stats[category]["SizeKB"], f"{expected_bytes[category] / 1024:.2f}")
        self.assertEqual(stats["自定义"]["FileCount"], 1)
        self.assertEqual(stats["自定义"]["SizeKB"], "3.00")
        report_path = Path(result["ReportPath"])
        report = report_path.read_text(encoding="utf-8-sig")
        self.assertNotIn("字节", report)
        self.assertNotIn("MiB", report)
        self.assertNotIn("CSV", report)
        self.assertRegex(report, r"统计时间：\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}（北京时间）")
        total_count = sum(expected_counts.values()) + 1
        total_kb = (sum(expected_bytes.values()) + extra.stat().st_size) / 1024
        self.assertRegex(report, rf"总计\s+{total_count} 个\s+{total_kb:.2f} KB")
        moved_contents = collections.Counter(
            hashlib.sha256(p.read_bytes()).hexdigest()
            for p in self.output.rglob("*") if p.is_file()
            and p != report_path
        )
        self.assertEqual(contents_before, moved_contents)
        repeated = run(self.root, apply=True)
        self.assertTrue(repeated["Success"])
        self.assertEqual(repeated["FileCount"], 0)
        self.assertEqual(repeated["Statistics"], result["Statistics"])
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
        self.assertTrue(repeated["Success"])
        self.assertEqual(repeated["FileCount"], 0)
        self.assertEqual(first["Statistics"], repeated["Statistics"])
        self.assertEqual([p.name for p in self.output.iterdir() if p.is_file()], ["整理统计.txt"])
        self.assertFalse(list(self.output.glob("移动记录_*.csv")))

    def test_empty_run_and_rounded_fractional_kb(self):
        empty = run(self.root, apply=True)
        self.assertTrue(empty["Success"])
        self.assertEqual(len(empty["Statistics"]), 6)
        self.assertTrue(all(r["FileCount"] == 0 and r["SizeKB"] == "0.00" for r in empty["Statistics"]))
        put(self.root / "fraction.pdf", b"1234567890")
        result = run(self.root, apply=True)
        other = next(r for r in result["Statistics"] if r["Folder"] == "其他")
        self.assertEqual(other["SizeKB"], "0.01")

    def test_locked_source_is_retained_and_report_counts_actual_files(self):
        blocked = put(self.root / "月份" / "占用.docx", b"locked")
        put(self.root / "月份" / "正常.pdf", b"x" * 2048)
        with ExclusiveFile(blocked):
            result = run(self.root, apply=True)
        self.assertFalse(result["Success"])
        self.assertEqual(result["FailedCount"], 1)
        self.assertTrue(blocked.exists())
        self.assertTrue((self.output / "其他" / "正常.pdf").exists())
        self.assertEqual(sum(r["FileCount"] for r in result["Statistics"]), 1)
        self.assertTrue(Path(result["ReportPath"]).exists())
        self.assertTrue(any(r["Status"] == "Failed" and r["Error"] for r in result["Records"]))
        self.assertFalse(list(self.output.glob("移动记录_*.csv")))

    def test_locked_report_is_reported_as_failure(self):
        first = run(self.root, apply=True)
        with ExclusiveFile(Path(first["ReportPath"])):
            result = run(self.root, apply=True)
        self.assertFalse(result["Success"])
        self.assertIsNone(result["ReportPath"])
        self.assertTrue(any(i["Stage"] == "Report" for i in result["Failures"]))
        self.assertFalse(list(self.output.glob(".统计_*.tmp")))

    def test_directory_links_are_skipped_and_linked_target_is_rejected(self):
        outside = self.temp_path / "outside"
        outside.mkdir()
        put(outside / "不能移动.docx", b"outside")
        link = self.root / "目录链接"
        os.symlink(outside, link, target_is_directory=True)
        try:
            result = run(self.root, apply=True)
            self.assertTrue(result["Success"])
            self.assertEqual(result["FileCount"], 0)
            self.assertIn(str(link), result["SkippedLinks"])
            self.assertTrue((outside / "不能移动.docx").exists())
            result = run(link, check=False)
            self.assertNotEqual(result.returncode, 0)
            target_link = self.output / "额外链接"
            os.symlink(outside, target_link, target_is_directory=True)
            try:
                stats = run(self.root, apply=True)
                self.assertNotIn("额外链接", [r["Folder"] for r in stats["Statistics"]])
            finally:
                target_link.unlink()
            (self.output / "word").rmdir()
            os.symlink(outside, self.output / "word", target_is_directory=True)
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
            os.symlink(outside, link)
        except OSError as exc:
            self.skipTest(f"Windows file-symlink privilege unavailable: {exc.winerror}")
        try:
            result = run(self.root, apply=True)
            self.assertEqual(result["FileCount"], 0)
            self.assertIn(str(link), result["SkippedLinks"])
            self.assertTrue(outside.exists())
        finally:
            link.unlink()

    def test_existing_target_file_is_rejected_before_moving(self):
        source = put(self.root / "原件.docx")
        put(self.output / "word", b"not a directory")
        result = run(self.root, apply=True, check=False)
        self.assertNotEqual(result.returncode, 0)
        self.assertTrue(source.exists())


if __name__ == "__main__":
    unittest.main(verbosity=2)
