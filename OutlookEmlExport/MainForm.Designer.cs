namespace OutlookEmlExport
{
    partial class MainForm
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            this.lblAccounts = new System.Windows.Forms.Label();
            this.lstAccounts = new System.Windows.Forms.CheckedListBox();
            this.btnListAccounts = new System.Windows.Forms.Button();
            this.btnSelectAll = new System.Windows.Forms.Button();
            this.btnSelectNone = new System.Windows.Forms.Button();
            this.lblOutput = new System.Windows.Forms.Label();
            this.txtOutput = new System.Windows.Forms.TextBox();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.chkVerbose = new System.Windows.Forms.CheckBox();
            this.chkResetState = new System.Windows.Forms.CheckBox();
            this.btnExport = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.txtLog = new System.Windows.Forms.TextBox();
            this.lblLog = new System.Windows.Forms.Label();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            //
            // lblAccounts
            //
            this.lblAccounts.AutoSize = true;
            this.lblAccounts.Location = new System.Drawing.Point(12, 15);
            this.lblAccounts.Name = "lblAccounts";
            this.lblAccounts.Size = new System.Drawing.Size(197, 12);
            this.lblAccounts.TabIndex = 0;
            this.lblAccounts.Text = "Outlook 账号 / 数据文件（勾选导出）";
            //
            // lstAccounts
            //
            this.lstAccounts.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lstAccounts.CheckOnClick = true;
            this.lstAccounts.FormattingEnabled = true;
            this.lstAccounts.HorizontalScrollbar = true;
            this.lstAccounts.Location = new System.Drawing.Point(14, 35);
            this.lstAccounts.Name = "lstAccounts";
            this.lstAccounts.Size = new System.Drawing.Size(658, 148);
            this.lstAccounts.TabIndex = 1;
            //
            // btnListAccounts
            //
            this.btnListAccounts.Location = new System.Drawing.Point(14, 195);
            this.btnListAccounts.Name = "btnListAccounts";
            this.btnListAccounts.Size = new System.Drawing.Size(120, 28);
            this.btnListAccounts.TabIndex = 2;
            this.btnListAccounts.Text = "列出账号";
            this.btnListAccounts.UseVisualStyleBackColor = true;
            this.btnListAccounts.Click += new System.EventHandler(this.btnListAccounts_Click);
            //
            // btnSelectAll
            //
            this.btnSelectAll.Location = new System.Drawing.Point(140, 195);
            this.btnSelectAll.Name = "btnSelectAll";
            this.btnSelectAll.Size = new System.Drawing.Size(80, 28);
            this.btnSelectAll.TabIndex = 3;
            this.btnSelectAll.Text = "全选";
            this.btnSelectAll.UseVisualStyleBackColor = true;
            this.btnSelectAll.Click += new System.EventHandler(this.btnSelectAll_Click);
            //
            // btnSelectNone
            //
            this.btnSelectNone.Location = new System.Drawing.Point(226, 195);
            this.btnSelectNone.Name = "btnSelectNone";
            this.btnSelectNone.Size = new System.Drawing.Size(80, 28);
            this.btnSelectNone.TabIndex = 4;
            this.btnSelectNone.Text = "全不选";
            this.btnSelectNone.UseVisualStyleBackColor = true;
            this.btnSelectNone.Click += new System.EventHandler(this.btnSelectNone_Click);
            //
            // lblOutput
            //
            this.lblOutput.AutoSize = true;
            this.lblOutput.Location = new System.Drawing.Point(12, 240);
            this.lblOutput.Name = "lblOutput";
            this.lblOutput.Size = new System.Drawing.Size(53, 12);
            this.lblOutput.TabIndex = 5;
            this.lblOutput.Text = "输出目录";
            //
            // txtOutput
            //
            this.txtOutput.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtOutput.Location = new System.Drawing.Point(71, 236);
            this.txtOutput.Name = "txtOutput";
            this.txtOutput.Size = new System.Drawing.Size(520, 21);
            this.txtOutput.TabIndex = 6;
            //
            // btnBrowse
            //
            this.btnBrowse.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnBrowse.Location = new System.Drawing.Point(597, 234);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(75, 25);
            this.btnBrowse.TabIndex = 7;
            this.btnBrowse.Text = "浏览…";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            //
            // chkVerbose
            //
            this.chkVerbose.AutoSize = true;
            this.chkVerbose.Location = new System.Drawing.Point(14, 275);
            this.chkVerbose.Name = "chkVerbose";
            this.chkVerbose.Size = new System.Drawing.Size(120, 16);
            this.chkVerbose.TabIndex = 8;
            this.chkVerbose.Text = "详细失败信息";
            this.chkVerbose.UseVisualStyleBackColor = true;
            //
            // chkResetState
            //
            this.chkResetState.AutoSize = true;
            this.chkResetState.Location = new System.Drawing.Point(150, 275);
            this.chkResetState.Name = "chkResetState";
            this.chkResetState.Size = new System.Drawing.Size(216, 16);
            this.chkResetState.TabIndex = 9;
            this.chkResetState.Text = "忽略断点重新导出（reset state）";
            this.chkResetState.UseVisualStyleBackColor = true;
            //
            // btnExport
            //
            this.btnExport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExport.Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, System.Drawing.FontStyle.Bold);
            this.btnExport.Location = new System.Drawing.Point(492, 268);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(100, 30);
            this.btnExport.TabIndex = 10;
            this.btnExport.Text = "导出";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            //
            // btnCancel
            //
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.Enabled = false;
            this.btnCancel.Location = new System.Drawing.Point(598, 268);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(74, 30);
            this.btnCancel.TabIndex = 11;
            this.btnCancel.Text = "取消";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            //
            // lblLog
            //
            this.lblLog.AutoSize = true;
            this.lblLog.Location = new System.Drawing.Point(12, 310);
            this.lblLog.Name = "lblLog";
            this.lblLog.Size = new System.Drawing.Size(29, 12);
            this.lblLog.TabIndex = 12;
            this.lblLog.Text = "日志";
            //
            // progressBar
            //
            this.progressBar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar.Location = new System.Drawing.Point(14, 328);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(658, 12);
            this.progressBar.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
            this.progressBar.MarqueeAnimationSpeed = 0;
            this.progressBar.TabIndex = 13;
            //
            // txtLog
            //
            this.txtLog.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtLog.Font = new System.Drawing.Font("Consolas", 9F);
            this.txtLog.Location = new System.Drawing.Point(14, 348);
            this.txtLog.Multiline = true;
            this.txtLog.Name = "txtLog";
            this.txtLog.ReadOnly = true;
            this.txtLog.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtLog.Size = new System.Drawing.Size(658, 201);
            this.txtLog.TabIndex = 14;
            this.txtLog.WordWrap = false;
            //
            // MainForm
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(684, 561);
            this.Controls.Add(this.txtLog);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.lblLog);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.chkResetState);
            this.Controls.Add(this.chkVerbose);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.txtOutput);
            this.Controls.Add(this.lblOutput);
            this.Controls.Add(this.btnSelectNone);
            this.Controls.Add(this.btnSelectAll);
            this.Controls.Add(this.btnListAccounts);
            this.Controls.Add(this.lstAccounts);
            this.Controls.Add(this.lblAccounts);
            this.MinimumSize = new System.Drawing.Size(640, 480);
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Outlook EML 导出";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.Label lblAccounts;
        private System.Windows.Forms.CheckedListBox lstAccounts;
        private System.Windows.Forms.Button btnListAccounts;
        private System.Windows.Forms.Button btnSelectAll;
        private System.Windows.Forms.Button btnSelectNone;
        private System.Windows.Forms.Label lblOutput;
        private System.Windows.Forms.TextBox txtOutput;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.CheckBox chkVerbose;
        private System.Windows.Forms.CheckBox chkResetState;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.TextBox txtLog;
        private System.Windows.Forms.Label lblLog;
        private System.Windows.Forms.ProgressBar progressBar;
    }
}
