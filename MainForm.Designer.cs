namespace ExtractSWF
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.exportBtn = new System.Windows.Forms.Button();
            this.openFolderBtn = new System.Windows.Forms.Button();
            this.dragPanel = new System.Windows.Forms.Panel();
            this.dragLabel = new System.Windows.Forms.Label();
            this.fileList = new System.Windows.Forms.ListBox();
            this.clearListBtn = new System.Windows.Forms.Button();
            this.aboutBtn = new System.Windows.Forms.Button();
            this.flashFileProgressBar = new System.Windows.Forms.ProgressBar();
            this.fileProgressBar = new System.Windows.Forms.ProgressBar();
            this.fileNumLbl = new System.Windows.Forms.Label();
            this.flashNumLbl = new System.Windows.Forms.Label();
            this.folderLbl = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // exportBtn
            // 
            this.exportBtn.Location = new System.Drawing.Point(12, 307);
            this.exportBtn.Name = "exportBtn";
            this.exportBtn.Size = new System.Drawing.Size(125, 25);
            this.exportBtn.TabIndex = 4;
            this.exportBtn.Text = "Export SWF";
            this.exportBtn.UseVisualStyleBackColor = true;
            this.exportBtn.Click += new System.EventHandler(this.exportBtn_Click);
            // 
            // openFolderBtn
            // 
            this.openFolderBtn.Location = new System.Drawing.Point(146, 307);
            this.openFolderBtn.Name = "openFolderBtn";
            this.openFolderBtn.Size = new System.Drawing.Size(197, 25);
            this.openFolderBtn.TabIndex = 5;
            this.openFolderBtn.Text = "Open Folder";
            this.openFolderBtn.UseVisualStyleBackColor = true;
            this.openFolderBtn.Click += new System.EventHandler(this.openFolderBtn_Click);
            // 
            // dragPanel
            // 
            this.dragPanel.AllowDrop = true;
            this.dragPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.dragPanel.Location = new System.Drawing.Point(12, 25);
            this.dragPanel.Name = "dragPanel";
            this.dragPanel.Size = new System.Drawing.Size(331, 276);
            this.dragPanel.TabIndex = 6;
            this.dragPanel.DragDrop += new System.Windows.Forms.DragEventHandler(this.dragPanel_DragDrop);
            this.dragPanel.DragEnter += new System.Windows.Forms.DragEventHandler(this.dragPanel_DragEnter);
            // 
            // dragLabel
            // 
            this.dragLabel.AutoSize = true;
            this.dragLabel.Location = new System.Drawing.Point(12, 9);
            this.dragLabel.Name = "dragLabel";
            this.dragLabel.Size = new System.Drawing.Size(260, 13);
            this.dragLabel.TabIndex = 7;
            this.dragLabel.Text = "Drag your presentation(s) and the output folder below:";
            // 
            // fileList
            // 
            this.fileList.FormattingEnabled = true;
            this.fileList.HorizontalScrollbar = true;
            this.fileList.Location = new System.Drawing.Point(349, 63);
            this.fileList.Name = "fileList";
            this.fileList.Size = new System.Drawing.Size(273, 238);
            this.fileList.TabIndex = 8;
            // 
            // clearListBtn
            // 
            this.clearListBtn.Location = new System.Drawing.Point(353, 307);
            this.clearListBtn.Name = "clearListBtn";
            this.clearListBtn.Size = new System.Drawing.Size(142, 24);
            this.clearListBtn.TabIndex = 10;
            this.clearListBtn.Text = "Clear list";
            this.clearListBtn.UseVisualStyleBackColor = true;
            this.clearListBtn.Click += new System.EventHandler(this.clearListBtn_Click);
            // 
            // aboutBtn
            // 
            this.aboutBtn.Location = new System.Drawing.Point(501, 307);
            this.aboutBtn.Name = "aboutBtn";
            this.aboutBtn.Size = new System.Drawing.Size(121, 24);
            this.aboutBtn.TabIndex = 11;
            this.aboutBtn.Text = "About";
            this.aboutBtn.UseVisualStyleBackColor = true;
            this.aboutBtn.Click += new System.EventHandler(this.aboutBtn_Click);
            // 
            // flashFileProgressBar
            // 
            this.flashFileProgressBar.Location = new System.Drawing.Point(80, 383);
            this.flashFileProgressBar.Name = "flashFileProgressBar";
            this.flashFileProgressBar.Size = new System.Drawing.Size(542, 26);
            this.flashFileProgressBar.TabIndex = 12;
            // 
            // fileProgressBar
            // 
            this.fileProgressBar.Location = new System.Drawing.Point(80, 351);
            this.fileProgressBar.Name = "fileProgressBar";
            this.fileProgressBar.Size = new System.Drawing.Size(542, 26);
            this.fileProgressBar.TabIndex = 13;
            // 
            // fileNumLbl
            // 
            this.fileNumLbl.AutoSize = true;
            this.fileNumLbl.Location = new System.Drawing.Point(12, 351);
            this.fileNumLbl.Name = "fileNumLbl";
            this.fileNumLbl.Size = new System.Drawing.Size(30, 13);
            this.fileNumLbl.TabIndex = 14;
            this.fileNumLbl.Text = "(n/a)";
            // 
            // flashNumLbl
            // 
            this.flashNumLbl.AutoSize = true;
            this.flashNumLbl.Location = new System.Drawing.Point(12, 383);
            this.flashNumLbl.Name = "flashNumLbl";
            this.flashNumLbl.Size = new System.Drawing.Size(30, 13);
            this.flashNumLbl.TabIndex = 15;
            this.flashNumLbl.Text = "(n/a)";
            // 
            // folderLbl
            // 
            this.folderLbl.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.folderLbl.Location = new System.Drawing.Point(349, 25);
            this.folderLbl.Multiline = true;
            this.folderLbl.Name = "folderLbl";
            this.folderLbl.ReadOnly = true;
            this.folderLbl.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal;
            this.folderLbl.Size = new System.Drawing.Size(273, 32);
            this.folderLbl.TabIndex = 16;
            this.folderLbl.Text = "(no folder selected)";
            this.folderLbl.WordWrap = false;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(634, 421);
            this.Controls.Add(this.folderLbl);
            this.Controls.Add(this.flashNumLbl);
            this.Controls.Add(this.fileNumLbl);
            this.Controls.Add(this.fileProgressBar);
            this.Controls.Add(this.flashFileProgressBar);
            this.Controls.Add(this.aboutBtn);
            this.Controls.Add(this.clearListBtn);
            this.Controls.Add(this.fileList);
            this.Controls.Add(this.dragLabel);
            this.Controls.Add(this.dragPanel);
            this.Controls.Add(this.openFolderBtn);
            this.Controls.Add(this.exportBtn);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "MainForm";
            this.Text = "Extract SWF";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button exportBtn;
        private System.Windows.Forms.Button openFolderBtn;
        private System.Windows.Forms.Panel dragPanel;
        private System.Windows.Forms.Label dragLabel;
        private System.Windows.Forms.ListBox fileList;
        private System.Windows.Forms.Button clearListBtn;
        private System.Windows.Forms.Button aboutBtn;
        private System.Windows.Forms.ProgressBar flashFileProgressBar;
        private System.Windows.Forms.ProgressBar fileProgressBar;
        private System.Windows.Forms.Label fileNumLbl;
        private System.Windows.Forms.Label flashNumLbl;
        private System.Windows.Forms.TextBox folderLbl;
    }
}

