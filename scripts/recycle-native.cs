// Windows Shell adapter. A missing recycle flag vetoes deletion before it starts.
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Threading;

namespace WechatOrganizerRecycleV1
{
    public sealed class Result
    {
        public bool Recycled;
        public string Error;
        public string RecycledPath;
    }

    public static class NativeRecycle
    {
        public static bool CanRecycle(uint flags) { return (flags & 0x80) != 0; }

        public static Result RecycleDuplicate(string path, string keeper, string hash, long length,
            long creationTicks, long writeTicks)
        {
            Result result = null;
            Thread thread = new Thread(delegate()
            {
                try { result = Run(path, keeper, hash, length, creationTicks, writeTicks); }
                catch (Exception ex) { result = new Result { Error = ex.Message }; }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
            return result;
        }

        static Result Run(string path, string keeper, string hash, long length, long creation, long write)
        {
            if (String.Equals(Path.GetFullPath(path), Path.GetFullPath(keeper), StringComparison.OrdinalIgnoreCase))
                throw new IOException("Candidate and keeper must be different files.");
            CheckPlainFile(path);
            CheckPlainFile(keeper);
            // Deny changes or removal of the verified keeper until recycling finishes.
            using (FileStream kept = new FileStream(keeper, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                if (kept.Length != length || !String.Equals(Hash(kept), hash, StringComparison.OrdinalIgnoreCase))
                    throw new IOException("Keeper changed; no duplicate was recycled.");
                IFileOperation operation = null;
                IShellItem item = null;
                Sink sink = new Sink(path, hash, length, creation, write);
                try
                {
                    operation = (IFileOperation)new FileOperation();
                    // Recycle requested; no UI, no elevation, no silent permanent-delete fallback.
                    Throw(operation.SetOperationFlags(0x00080000 | 0x00100000 | 0x0004 | 0x0010 | 0x0400));
                    Guid iid = typeof(IShellItem).GUID;
                    Throw(SHCreateItemFromParsingName(path, IntPtr.Zero, ref iid, out item));
                    Throw(operation.DeleteItem(item, sink));
                    int hr = operation.PerformOperations();
                    bool aborted;
                    int state = operation.GetAnyOperationsAborted(out aborted);
                    if (hr < 0 || state < 0 || aborted || !sink.Completed || File.Exists(path))
                        return new Result { Error = sink.Error ?? String.Format("Recycle refused or failed (0x{0:X8}); permanent deletion is disabled.", hr) };
                    return new Result { Recycled = true, RecycledPath = sink.RecycledPath };
                }
                finally
                {
                    sink.Dispose();
                    if (item != null) Marshal.ReleaseComObject(item);
                    if (operation != null) Marshal.ReleaseComObject(operation);
                }
            }
        }

        internal static void CheckPlainFile(string path)
        {
            FileAttributes a = File.GetAttributes(path);
            if ((a & (FileAttributes.ReparsePoint | FileAttributes.Directory)) != 0)
                throw new IOException("Only regular files may be recycled.");
            DirectoryInfo parent = new FileInfo(path).Directory;
            while (parent != null)
            {
                if ((parent.Attributes & FileAttributes.ReparsePoint) != 0)
                    throw new IOException("Linked ancestors are excluded.");
                parent = parent.Parent;
            }
        }

        internal static string Hash(Stream stream)
        {
            stream.Position = 0;
            using (SHA256 sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "");
        }
        static void Throw(int hr) { if (hr < 0) Marshal.ThrowExceptionForHR(hr); }

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = true)]
        static extern int SHCreateItemFromParsingName(string path, IntPtr bind, ref Guid iid, out IShellItem item);

        [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
        public sealed class Sink : IFileOperationProgressSink, IDisposable
        {
            readonly string path, hash;
            readonly long length, creation, write;
            FileStream guard;
            bool approved;
            public bool Completed;
            public string Error, RecycledPath;
            public Sink(string p, string h, long l, long c, long w) { path = p; hash = h; length = l; creation = c; write = w; }
            public int PreDeleteItem(uint flags, IShellItem item)
            {
                try
                {
                    // This callback is mandatory: requesting recycling alone is insufficient.
                    if (!CanRecycle(flags)) throw new IOException("Windows cannot recycle this item; deletion was vetoed.");
                    CheckPlainFile(path);
                    FileInfo info = new FileInfo(path);
                    if (info.Length != length || info.CreationTimeUtc.Ticks != creation || info.LastWriteTimeUtc.Ticks != write)
                        throw new IOException("Candidate changed; deletion was vetoed.");
                    // Shell may rename/delete it, but writers cannot alter its contents.
                    guard = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read | FileShare.Delete);
                    if (guard.Length != length || !String.Equals(Hash(guard), hash, StringComparison.OrdinalIgnoreCase))
                        throw new IOException("Candidate content changed; deletion was vetoed.");
                    approved = true;
                    return 0;
                }
                catch (Exception ex) { Error = ex.Message; Dispose(); return unchecked((int)0x80070032); }
            }
            public int PostDeleteItem(uint flags, IShellItem item, int hr, IShellItem created)
            {
                Completed = approved && CanRecycle(flags) && hr >= 0;
                if (hr < 0) Error = String.Format("Windows recycle failed (0x{0:X8}).", hr);
                if (Completed && created != null)
                {
                    IntPtr text;
                    if (created.GetDisplayName(0x80058000, out text) >= 0 && text != IntPtr.Zero)
                    {
                        try { RecycledPath = Marshal.PtrToStringUni(text); }
                        finally { Marshal.FreeCoTaskMem(text); }
                    }
                }
                Dispose();
                return 0;
            }
            public void Dispose() { if (guard != null) { guard.Dispose(); guard = null; } }
            public int StartOperations() { return 0; }
            public int FinishOperations(int result) { return 0; }
            public int PreRenameItem(uint f, IShellItem i, string n) { return 0; }
            public int PostRenameItem(uint f, IShellItem i, string n, int h, IShellItem c) { return 0; }
            public int PreMoveItem(uint f, IShellItem i, IShellItem d, string n) { return 0; }
            public int PostMoveItem(uint f, IShellItem i, IShellItem d, string n, int h, IShellItem c) { return 0; }
            public int PreCopyItem(uint f, IShellItem i, IShellItem d, string n) { return 0; }
            public int PostCopyItem(uint f, IShellItem i, IShellItem d, string n, int h, IShellItem c) { return 0; }
            public int PreNewItem(uint f, IShellItem d, string n) { return 0; }
            public int PostNewItem(uint f, IShellItem d, string n, string t, uint a, int h, IShellItem c) { return 0; }
            public int UpdateProgress(uint total, uint done) { return 0; }
            public int ResetTimer() { return 0; }
            public int PauseTimer() { return 0; }
            public int ResumeTimer() { return 0; }
        }

        [ComImport, Guid("3AD05575-8857-4850-9277-11B85BDB8E09")] class FileOperation { }
        [ComImport, Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IShellItem
        {
            [PreserveSig] int BindToHandler(IntPtr b, ref Guid h, ref Guid i, out IntPtr p);
            [PreserveSig] int GetParent(out IShellItem p);
            [PreserveSig] int GetDisplayName(uint type, out IntPtr name);
            [PreserveSig] int GetAttributes(uint mask, out uint attributes);
            [PreserveSig] int Compare(IShellItem other, uint hint, out int order);
        }
        [ComImport, Guid("947AAB5F-0A5C-4C13-B4D6-4BF7836FC9F8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IFileOperation
        {
            [PreserveSig] int Advise(IFileOperationProgressSink s, out uint cookie);
            [PreserveSig] int Unadvise(uint cookie);
            [PreserveSig] int SetOperationFlags(uint flags);
            [PreserveSig] int SetProgressMessage([MarshalAs(UnmanagedType.LPWStr)] string message);
            [PreserveSig] int SetProgressDialog(IntPtr dialog);
            [PreserveSig] int SetProperties(IntPtr changes);
            [PreserveSig] int SetOwnerWindow(IntPtr owner);
            [PreserveSig] int ApplyPropertiesToItem(IShellItem item);
            [PreserveSig] int ApplyPropertiesToItems([MarshalAs(UnmanagedType.IUnknown)] object items);
            [PreserveSig] int RenameItem(IShellItem item, [MarshalAs(UnmanagedType.LPWStr)] string name, IFileOperationProgressSink sink);
            [PreserveSig] int RenameItems([MarshalAs(UnmanagedType.IUnknown)] object items, [MarshalAs(UnmanagedType.LPWStr)] string name);
            [PreserveSig] int MoveItem(IShellItem item, IShellItem destination, [MarshalAs(UnmanagedType.LPWStr)] string name, IFileOperationProgressSink sink);
            [PreserveSig] int MoveItems([MarshalAs(UnmanagedType.IUnknown)] object items, IShellItem destination);
            [PreserveSig] int CopyItem(IShellItem item, IShellItem destination, [MarshalAs(UnmanagedType.LPWStr)] string name, IFileOperationProgressSink sink);
            [PreserveSig] int CopyItems([MarshalAs(UnmanagedType.IUnknown)] object items, IShellItem destination);
            [PreserveSig] int DeleteItem(IShellItem item, IFileOperationProgressSink sink);
            [PreserveSig] int DeleteItems([MarshalAs(UnmanagedType.IUnknown)] object items);
            [PreserveSig] int NewItem(IShellItem destination, uint attributes, [MarshalAs(UnmanagedType.LPWStr)] string name, [MarshalAs(UnmanagedType.LPWStr)] string template, IFileOperationProgressSink sink);
            [PreserveSig] int PerformOperations();
            [PreserveSig] int GetAnyOperationsAborted([MarshalAs(UnmanagedType.Bool)] out bool aborted);
        }
        [ComImport, Guid("04B0F1A7-9490-44BC-96E1-4296A31252E2"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IFileOperationProgressSink
        {
            [PreserveSig] int StartOperations();
            [PreserveSig] int FinishOperations(int result);
            [PreserveSig] int PreRenameItem(uint f, IShellItem i, [MarshalAs(UnmanagedType.LPWStr)] string n);
            [PreserveSig] int PostRenameItem(uint f, IShellItem i, [MarshalAs(UnmanagedType.LPWStr)] string n, int h, IShellItem c);
            [PreserveSig] int PreMoveItem(uint f, IShellItem i, IShellItem d, [MarshalAs(UnmanagedType.LPWStr)] string n);
            [PreserveSig] int PostMoveItem(uint f, IShellItem i, IShellItem d, [MarshalAs(UnmanagedType.LPWStr)] string n, int h, IShellItem c);
            [PreserveSig] int PreCopyItem(uint f, IShellItem i, IShellItem d, [MarshalAs(UnmanagedType.LPWStr)] string n);
            [PreserveSig] int PostCopyItem(uint f, IShellItem i, IShellItem d, [MarshalAs(UnmanagedType.LPWStr)] string n, int h, IShellItem c);
            [PreserveSig] int PreDeleteItem(uint f, IShellItem i);
            [PreserveSig] int PostDeleteItem(uint f, IShellItem i, int h, IShellItem c);
            [PreserveSig] int PreNewItem(uint f, IShellItem d, [MarshalAs(UnmanagedType.LPWStr)] string n);
            [PreserveSig] int PostNewItem(uint f, IShellItem d, [MarshalAs(UnmanagedType.LPWStr)] string n, [MarshalAs(UnmanagedType.LPWStr)] string t, uint a, int h, IShellItem c);
            [PreserveSig] int UpdateProgress(uint total, uint done);
            [PreserveSig] int ResetTimer();
            [PreserveSig] int PauseTimer();
            [PreserveSig] int ResumeTimer();
        }
    }
}
