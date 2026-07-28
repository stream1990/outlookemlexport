using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

namespace OutlookEmlExport.Export
{
    internal sealed class StoreInfo
    {
        public string DisplayName { get; set; }
        public object Store { get; set; }

        public override string ToString()
        {
            return DisplayName ?? "(unnamed)";
        }
    }

    internal sealed class ExportOptions
    {
        public ExportOptions()
        {
            SaveEvery = 50;
        }

        public string OutputRoot { get; set; }
        public IList<StoreInfo> SelectedStores { get; set; }
        public int SaveEvery { get; set; }
        public bool Verbose { get; set; }
        public bool ResetState { get; set; }
        public HashSet<string> ExtraSkipFolders { get; set; }
        public CancellationToken CancellationToken { get; set; }
        public Action<string> Log { get; set; }
    }

    internal sealed class ExportResult
    {
        public int Exported { get; set; }
        public int Skipped { get; set; }
        public int Failed { get; set; }
        public int Folders { get; set; }
        public string StatePath { get; set; }
        public string OutputRoot { get; set; }
        public bool Cancelled { get; set; }
    }

    /// <summary>
    /// 通过 COM late-binding 连接本机 Outlook，导出已缓存邮件为 EML。
    /// 不调用 SendAndReceive，不主动拉新邮件。
    /// </summary>
    internal sealed class OutlookExporter : IDisposable
    {
        public const int OlMail = 43;
        public const int OlFolder = 2;

        public static readonly HashSet<string> DefaultSkipFolders = CreateDefaultSkipFolders();

        private static HashSet<string> CreateDefaultSkipFolders()
        {
            var set = new HashSet<string>(StringComparer.Ordinal);
            string[] names = new string[]
            {
                "日历", "Calendar",
                "联系人", "Contacts",
                "任务", "Tasks",
                "日记", "Journal",
                "便笺", "Notes",
                "RSS 订阅", "RSS Feeds",
                "同步问题", "Sync Issues",
                "快速步骤设置", "Quick Step Settings",
                "会话操作设置", "Conversation Action Settings",
            };
            foreach (var n in names)
                set.Add(n);
            return set;
        }

        private object _outlook;
        private object _namespace;
        private bool _comInited;

        public void Connect()
        {
            if (_namespace != null)
                return;

            int hr = NativeMethods.CoInitializeEx(IntPtr.Zero, NativeMethods.COINIT_APARTMENTTHREADED);
            // S_OK / S_FALSE 时需要 Uninit
            if (hr == 0 || hr == 1)
                _comInited = true;

            Type t = Type.GetTypeFromProgID("Outlook.Application");
            if (t == null)
                throw new InvalidOperationException("未找到 Outlook.Application，请确认已安装桌面版 Outlook。");

            _outlook = Activator.CreateInstance(t);
            dynamic app = _outlook;
            _namespace = app.GetNamespace("MAPI");
        }

        public List<StoreInfo> ListStores()
        {
            Connect();
            var result = new List<StoreInfo>();
            dynamic ns = _namespace;
            int count;
            try { count = Convert.ToInt32(ns.Stores.Count); }
            catch { count = 0; }

            for (int i = 1; i <= count; i++)
            {
                try
                {
                    dynamic store = ns.Stores.Item(i);
                    string name = Utils.ComStr(store.DisplayName);
                    if (string.IsNullOrEmpty(name))
                        name = "Store_" + i;
                    result.Add(new StoreInfo { DisplayName = name, Store = store });
                }
                catch
                {
                    // skip bad store
                }
            }
            return result;
        }

        public ExportResult Export(ExportOptions options)
        {
            if (options == null)
                throw new ArgumentNullException("options");
            if (string.IsNullOrWhiteSpace(options.OutputRoot))
                throw new ArgumentException("输出目录不能为空。");
            if (options.SelectedStores == null || options.SelectedStores.Count == 0)
                throw new ArgumentException("请至少选择一个账号。");

            Connect();

            string outputRoot = Path.GetFullPath(options.OutputRoot);
            Utils.EnsureDir(outputRoot);

            var skipNames = new HashSet<string>(DefaultSkipFolders, StringComparer.Ordinal);
            if (options.ExtraSkipFolders != null)
            {
                foreach (var s in options.ExtraSkipFolders)
                {
                    if (!string.IsNullOrWhiteSpace(s))
                        skipNames.Add(s.Trim());
                }
            }

            string statePath = Path.Combine(outputRoot, ".export_state.json");
            if (options.ResetState && File.Exists(statePath))
                File.Delete(statePath);

            var state = new ExportState(statePath);
            Action<string> log = options.Log;
            if (log == null)
                log = delegate { };
            CancellationToken ct = options.CancellationToken;
            int saveEvery = options.SaveEvery > 0 ? options.SaveEvery : 50;

            log("输出目录: " + outputRoot);
            log("开始导出…");
            log(string.Empty);

            bool cancelled = false;

            foreach (var stInfo in options.SelectedStores)
            {
                ct.ThrowIfCancellationRequested();

                string storeName = stInfo.DisplayName ?? "Account";
                string storeDirName = Utils.SafeName(storeName, 100);
                string storeDir = Path.Combine(outputRoot, storeDirName);

                object root = null;
                try
                {
                    dynamic store = stInfo.Store;
                    root = store.GetRootFolder();
                }
                catch (Exception ex)
                {
                    log("[跳过] 无法打开: " + storeName + " — " + ex.Message);
                    continue;
                }

                if (root == null)
                {
                    log("[跳过] 无法打开: " + storeName);
                    continue;
                }

                log("=== 账号: " + storeName + " ===");
                Utils.EnsureDir(storeDir);

                try
                {
                    dynamic dRoot = root;
                    int topCount;
                    try { topCount = Convert.ToInt32(dRoot.Folders.Count); }
                    catch { topCount = 0; }

                    if (topCount == 0)
                    {
                        string rootName = Utils.SafeName(Utils.ComStr(dRoot.Name) ?? "Root");
                        ExportFolder(root, Path.Combine(storeDir, rootName), state, skipNames,
                            saveEvery, options.Verbose, log, ct);
                    }
                    else
                    {
                        for (int i = 1; i <= topCount; i++)
                        {
                            ct.ThrowIfCancellationRequested();
                            try
                            {
                                dynamic top = dRoot.Folders.Item(i);
                                string topName = Utils.SafeName(Utils.ComStr(top.Name) ?? ("Folder_" + i));
                                ExportFolder(top, Path.Combine(storeDir, topName), state, skipNames,
                                    saveEvery, options.Verbose, log, ct);
                            }
                            catch (OperationCanceledException)
                            {
                                throw;
                            }
                            catch (Exception ex)
                            {
                                log("  [顶级文件夹失败] " + ex.Message);
                                if (options.Verbose)
                                    log(ex.ToString());
                            }
                        }
                    }
                }
                catch (OperationCanceledException)
                {
                    cancelled = true;
                    state.Save();
                    log("已取消。下次运行会按断点跳过已导出的邮件。");
                    break;
                }

                state.Save();
                log(string.Empty);
            }

            state.Save();

            if (!cancelled)
            {
                log("—— 完成 ——");
                log(string.Format(
                    "新导出: {0}  |  跳过(已导出过): {1}  |  失败: {2}  |  遍历文件夹: {3}",
                    GetStat(state, "exported"),
                    GetStat(state, "skipped"),
                    GetStat(state, "failed"),
                    GetStat(state, "folders")));
                log("断点文件: " + statePath);
                log("邮件目录: " + outputRoot);
            }

            return new ExportResult
            {
                Exported = GetStat(state, "exported"),
                Skipped = GetStat(state, "skipped"),
                Failed = GetStat(state, "failed"),
                Folders = GetStat(state, "folders"),
                StatePath = statePath,
                OutputRoot = outputRoot,
                Cancelled = cancelled,
            };
        }

        private static int GetStat(ExportState state, string key)
        {
            int v;
            if (state.Stats.TryGetValue(key, out v))
                return v;
            return 0;
        }

        private void ExportFolder(
            object folder,
            string destDir,
            ExportState state,
            HashSet<string> skipNames,
            int saveEvery,
            bool verbose,
            Action<string> log,
            CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();

            string name = Utils.ComStr(GetDynProp(folder, "Name"));
            if (string.IsNullOrEmpty(name))
                name = "folder";

            if (skipNames.Contains(name.Trim()))
            {
                if (verbose)
                    log("  [跳过文件夹] " + name);
                return;
            }

            Utils.EnsureDir(destDir);
            state.Inc("folders");

            int exportedHere = 0;
            int failedHere = 0;
            int skippedHere = 0;

            try
            {
                foreach (var item in IterMailItems(folder))
                {
                    ct.ThrowIfCancellationRequested();
                    try
                    {
                        int cls = 0;
                        try { cls = Convert.ToInt32(GetDynProp(item, "Class") ?? 0); }
                        catch { cls = 0; }

                        if (cls != OlMail)
                        {
                            string mc = Utils.ComStr(GetDynProp(item, "MessageClass"));
                            if (string.IsNullOrEmpty(mc) || !mc.StartsWith("IPM.Note", StringComparison.OrdinalIgnoreCase))
                                continue;
                        }

                        string entryId = Utils.ComStr(GetDynProp(item, "EntryID"));
                        if (string.IsNullOrEmpty(entryId))
                        {
                            entryId = Utils.ShortId(
                                Utils.ComStr(GetDynProp(item, "Subject"))
                                + Take(Utils.ComStr(GetDynProp(item, "Body")), 200),
                                40);
                        }

                        if (state.IsDone(entryId))
                        {
                            skippedHere++;
                            state.Inc("skipped");
                            continue;
                        }

                        byte[] raw = EmlBuilder.BuildEmlBytes(item);
                        string outPath = Utils.UniqueEmlPath(destDir, EmlBuilder.MailFilename(item));
                        File.WriteAllBytes(outPath, raw);

                        state.MarkDone(entryId);
                        state.Inc("exported");
                        exportedHere++;

                        if (GetStat(state, "exported") % saveEvery == 0)
                        {
                            state.Save();
                            log(string.Format(
                                "  …已导出 {0} 封 (跳过 {1}, 失败 {2})",
                                GetStat(state, "exported"),
                                GetStat(state, "skipped"),
                                GetStat(state, "failed")));
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        throw;
                    }
                    catch (Exception ex)
                    {
                        failedHere++;
                        state.Inc("failed");
                        if (verbose)
                        {
                            string subj = string.Empty;
                            try { subj = Utils.ComStr(GetDynProp(item, "Subject")); }
                            catch { /* ignore */ }
                            log("  [失败] " + (string.IsNullOrEmpty(subj) ? "(unknown)" : subj) + ": " + ex.Message);
                        }
                    }
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                log("  [读文件夹失败] " + name + ": " + ex.Message);
            }

            string displayPath = destDir.Replace('\\', '/');
            log(string.Format("  {0}  →  新导出 {1}, 跳过 {2}, 失败 {3}",
                displayPath, exportedHere, skippedHere, failedHere));

            // 子文件夹
            try
            {
                dynamic dFolder = folder;
                int subCount;
                try { subCount = Convert.ToInt32(dFolder.Folders.Count); }
                catch { subCount = 0; }

                for (int i = 1; i <= subCount; i++)
                {
                    ct.ThrowIfCancellationRequested();
                    try
                    {
                        dynamic sub = dFolder.Folders.Item(i);
                        string subName = Utils.SafeName(Utils.ComStr(sub.Name) ?? ("sub_" + i));
                        ExportFolder(sub, Path.Combine(destDir, subName), state, skipNames,
                            saveEvery, verbose, log, ct);
                    }
                    catch (OperationCanceledException)
                    {
                        throw;
                    }
                    catch (Exception ex)
                    {
                        log("  [子文件夹失败] " + ex.Message);
                    }
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch
            {
                // ignore folder enumeration errors
            }
        }

        private static IEnumerable<object> IterMailItems(object folder)
        {
            dynamic dFolder = folder;
            dynamic items = dFolder.Items;

            try { items.Sort("[ReceivedTime]", true); }
            catch { /* ignore */ }

            object source = items;
            try
            {
                dynamic restricted = items.Restrict(
                    "@SQL=\"http://schemas.microsoft.com/mapi/proptag/0x001A001F\" LIKE 'IPM.Note%'");
                source = restricted;
            }
            catch
            {
                source = items;
            }

            // GetFirst / GetNext 更稳
            bool usedIterator = false;
            object firstItem = null;
            try
            {
                dynamic src = source;
                firstItem = src.GetFirst();
                usedIterator = true;
            }
            catch
            {
                usedIterator = false;
            }

            if (usedIterator)
            {
                dynamic src = source;
                object item = firstItem;
                while (item != null)
                {
                    yield return item;
                    try { item = src.GetNext(); }
                    catch { break; }
                }
                yield break;
            }

            int count;
            try { count = Convert.ToInt32(((dynamic)source).Count); }
            catch { count = 0; }

            for (int i = 1; i <= count; i++)
            {
                object item = null;
                try { item = ((dynamic)source).Item(i); }
                catch { continue; }
                if (item != null)
                    yield return item;
            }
        }

        private static object GetDynProp(object comObj, string name)
        {
            if (comObj == null)
                return null;
            try
            {
                dynamic d = comObj;
                switch (name)
                {
                    case "Name": return d.Name;
                    case "Class": return d.Class;
                    case "MessageClass": return d.MessageClass;
                    case "EntryID": return d.EntryID;
                    case "Subject": return d.Subject;
                    case "Body": return d.Body;
                    default:
                        return comObj.GetType().InvokeMember(
                            name,
                            System.Reflection.BindingFlags.GetProperty,
                            null, comObj, null);
                }
            }
            catch
            {
                return null;
            }
        }

        private static string Take(string s, int n)
        {
            if (string.IsNullOrEmpty(s))
                return string.Empty;
            return s.Length <= n ? s : s.Substring(0, n);
        }

        public void Dispose()
        {
            try
            {
                if (_namespace != null)
                {
                    Marshal.FinalReleaseComObject(_namespace);
                    _namespace = null;
                }
            }
            catch { /* ignore */ }

            try
            {
                if (_outlook != null)
                {
                    Marshal.FinalReleaseComObject(_outlook);
                    _outlook = null;
                }
            }
            catch { /* ignore */ }

            if (_comInited)
            {
                try { NativeMethods.CoUninitialize(); }
                catch { /* ignore */ }
                _comInited = false;
            }
        }

        private static class NativeMethods
        {
            public const uint COINIT_APARTMENTTHREADED = 0x2;

            [DllImport("ole32.dll")]
            public static extern int CoInitializeEx(IntPtr pvReserved, uint dwCoInit);

            [DllImport("ole32.dll")]
            public static extern void CoUninitialize();
        }
    }
}
