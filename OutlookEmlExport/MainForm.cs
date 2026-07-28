using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OutlookEmlExport.Export;

namespace OutlookEmlExport
{
    public partial class MainForm : Form
    {
        private OutlookExporter _exporter;
        private CancellationTokenSource _cts;
        private bool _exporting;

        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            txtOutput.Text = AppDomain.CurrentDomain.BaseDirectory.TrimEnd(
                Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

            AppendLog("从本机 Outlook 导出已缓存/已下载邮件为 .eml");
            AppendLog("- 不触发发送/接收，不主动拉新邮件");
            AppendLog("- 一个账号(Store)一个目录，文件夹层级映射为子目录");
            AppendLog("- 支持断点续传（.export_state.json）");
            AppendLog(string.Empty);
            AppendLog("请先点「列出账号」，勾选要导出的账号后点「导出」。");
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = "选择 EML 输出根目录";
                dlg.ShowNewFolderButton = true;
                if (Directory.Exists(txtOutput.Text))
                    dlg.SelectedPath = txtOutput.Text;

                if (dlg.ShowDialog(this) == DialogResult.OK)
                    txtOutput.Text = dlg.SelectedPath;
            }
        }

        private void btnListAccounts_Click(object sender, EventArgs e)
        {
            if (_exporting)
                return;

            btnListAccounts.Enabled = false;
            try
            {
                AppendLog("连接本机 Outlook（只读本地存储，不触发发送/接收）…");
                Cursor = Cursors.WaitCursor;

                EnsureExporter();
                List<StoreInfo> stores = _exporter.ListStores();

                lstAccounts.Items.Clear();
                if (stores.Count == 0)
                {
                    AppendLog("未找到任何 Outlook 数据文件/账号。请先打开 Outlook 并确认已登录。");
                    MessageBox.Show(this,
                        "未找到任何 Outlook 数据文件/账号。\n请先打开 Outlook 并确认已登录。",
                        "提示",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                AppendLog("发现 " + stores.Count + " 个账号/数据文件：");
                for (int i = 0; i < stores.Count; i++)
                {
                    // 默认全勾选
                    lstAccounts.Items.Add(stores[i], true);
                    AppendLog("  " + (i + 1) + ". " + stores[i].DisplayName);
                }
                AppendLog("已默认全选，可取消不需要的账号后点「导出」。");
            }
            catch (Exception ex)
            {
                AppendLog("[错误] 列出账号失败: " + ex.Message);
                MessageBox.Show(this,
                    "列出账号失败：\n" + ex.Message +
                    "\n\n请确认已安装桌面版 Outlook，并至少打开过一次。",
                    "错误",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
                btnListAccounts.Enabled = true;
            }
        }

        private void btnSelectAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < lstAccounts.Items.Count; i++)
                lstAccounts.SetItemChecked(i, true);
        }

        private void btnSelectNone_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < lstAccounts.Items.Count; i++)
                lstAccounts.SetItemChecked(i, false);
        }

        private async void btnExport_Click(object sender, EventArgs e)
        {
            if (_exporting)
                return;

            if (lstAccounts.Items.Count == 0)
            {
                MessageBox.Show(this, "请先点「列出账号」。", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var selected = new List<StoreInfo>();
            for (int i = 0; i < lstAccounts.Items.Count; i++)
            {
                if (lstAccounts.GetItemChecked(i) && lstAccounts.Items[i] is StoreInfo)
                    selected.Add((StoreInfo)lstAccounts.Items[i]);
            }

            if (selected.Count == 0)
            {
                MessageBox.Show(this, "请至少勾选一个账号。", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string output = (txtOutput.Text ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(output))
            {
                MessageBox.Show(this, "请指定输出目录。", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                Directory.CreateDirectory(output);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "无法创建输出目录：\n" + ex.Message, "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            _exporting = true;
            _cts = new CancellationTokenSource();
            SetUiExporting(true);

            AppendLog(string.Empty);
            AppendLog("======== 开始导出 ========");
            AppendLog("选中账号 " + selected.Count + " 个：");
            foreach (var s in selected)
                AppendLog("  - " + s.DisplayName);

            try
            {
                EnsureExporter();
                var options = new ExportOptions();
                options.OutputRoot = output;
                options.SelectedStores = selected;
                options.SaveEvery = 50;
                options.Verbose = chkVerbose.Checked;
                options.ResetState = chkResetState.Checked;
                options.CancellationToken = _cts.Token;
                options.Log = delegate(string msg)
                {
                    if (IsDisposed)
                        return;
                    BeginInvoke(new Action(delegate { AppendLog(msg); }));
                };

                ExportResult result = await RunExportStaAsync(options, _cts.Token).ConfigureAwait(true);

                if (result.Cancelled)
                {
                    AppendLog("导出已取消。");
                }
                else
                {
                    MessageBox.Show(this,
                        string.Format(
                            "导出完成。\n\n新导出: {0}\n跳过: {1}\n失败: {2}\n文件夹: {3}\n\n目录:\n{4}",
                            result.Exported, result.Skipped, result.Failed, result.Folders, result.OutputRoot),
                        "完成",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
            catch (OperationCanceledException)
            {
                AppendLog("导出已取消。下次运行会按断点跳过已导出的邮件。");
            }
            catch (Exception ex)
            {
                AppendLog("[错误] 导出失败: " + ex.Message);
                if (chkVerbose.Checked)
                    AppendLog(ex.ToString());
                MessageBox.Show(this, "导出失败：\n" + ex.Message, "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _exporting = false;
                if (_cts != null)
                {
                    _cts.Dispose();
                    _cts = null;
                }
                SetUiExporting(false);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            try
            {
                if (_cts != null)
                    _cts.Cancel();
                AppendLog("正在取消…");
            }
            catch
            {
                // ignore
            }
        }

        private Task<ExportResult> RunExportStaAsync(ExportOptions options, CancellationToken ct)
        {
            var tcs = new TaskCompletionSource<ExportResult>();
            var thread = new Thread(delegate()
            {
                try
                {
                    // 导出线程自己持有 COM；列表时拿到的 COM 对象在另一线程不可用，需按名重绑
                    using (var exporter = new OutlookExporter())
                    {
                        List<StoreInfo> stores = exporter.ListStores();
                        var byName = new Dictionary<string, StoreInfo>(StringComparer.OrdinalIgnoreCase);
                        foreach (var s in stores)
                        {
                            if (!byName.ContainsKey(s.DisplayName))
                                byName[s.DisplayName] = s;
                        }

                        var rebound = new List<StoreInfo>();
                        foreach (var want in options.SelectedStores)
                        {
                            StoreInfo found;
                            if (byName.TryGetValue(want.DisplayName, out found))
                                rebound.Add(found);
                            else if (options.Log != null)
                                options.Log("[警告] 未找到账号: " + want.DisplayName);
                        }

                        options.SelectedStores = rebound;
                        options.CancellationToken = ct;
                        ExportResult result = exporter.Export(options);
                        tcs.TrySetResult(result);
                    }
                }
                catch (OperationCanceledException)
                {
                    tcs.TrySetCanceled();
                }
                catch (Exception ex)
                {
                    tcs.TrySetException(ex);
                }
            });
            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread.Name = "OutlookEmlExport";
            thread.Start();
            return tcs.Task;
        }

        private void EnsureExporter()
        {
            if (_exporter == null)
                _exporter = new OutlookExporter();
        }

        private void SetUiExporting(bool exporting)
        {
            btnListAccounts.Enabled = !exporting;
            btnSelectAll.Enabled = !exporting;
            btnSelectNone.Enabled = !exporting;
            btnBrowse.Enabled = !exporting;
            btnExport.Enabled = !exporting;
            lstAccounts.Enabled = !exporting;
            txtOutput.Enabled = !exporting;
            chkVerbose.Enabled = !exporting;
            chkResetState.Enabled = !exporting;
            btnCancel.Enabled = exporting;
            progressBar.MarqueeAnimationSpeed = exporting ? 30 : 0;
            if (!exporting)
                progressBar.Value = 0;
        }

        private void AppendLog(string line)
        {
            if (txtLog.TextLength > 0)
                txtLog.AppendText(Environment.NewLine);
            txtLog.AppendText(line ?? string.Empty);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (_exporting)
            {
                DialogResult r = MessageBox.Show(this,
                    "正在导出，确定要退出并取消吗？",
                    "确认",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                if (r != DialogResult.Yes)
                {
                    e.Cancel = true;
                    return;
                }
                try
                {
                    if (_cts != null)
                        _cts.Cancel();
                }
                catch { /* ignore */ }
            }

            try
            {
                if (_exporter != null)
                    _exporter.Dispose();
            }
            catch { /* ignore */ }
            _exporter = null;

            base.OnFormClosing(e);
        }
    }
}
