namespace PomMergeAcctsExcel
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
            this.btnSrcSelect = new System.Windows.Forms.Button();
            this.btnDestSelect = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txtSrcFileName = new System.Windows.Forms.TextBox();
            this.txtDestFileName = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnExecute = new System.Windows.Forms.Button();
            this.pb1 = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            // 
            // btnSrcSelect
            // 
            this.btnSrcSelect.Location = new System.Drawing.Point(685, 12);
            this.btnSrcSelect.Name = "btnSrcSelect";
            this.btnSrcSelect.Size = new System.Drawing.Size(75, 23);
            this.btnSrcSelect.TabIndex = 0;
            this.btnSrcSelect.Text = "选择文件";
            this.btnSrcSelect.UseVisualStyleBackColor = true;
            this.btnSrcSelect.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // btnDestSelect
            // 
            this.btnDestSelect.Location = new System.Drawing.Point(685, 46);
            this.btnDestSelect.Name = "btnDestSelect";
            this.btnDestSelect.Size = new System.Drawing.Size(75, 23);
            this.btnDestSelect.TabIndex = 1;
            this.btnDestSelect.Text = "选择路径";
            this.btnDestSelect.UseVisualStyleBackColor = true;
            this.btnDestSelect.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(36, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(179, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "请选择您想要合并的Excel文件：";
            // 
            // txtSrcFileName
            // 
            this.txtSrcFileName.Enabled = false;
            this.txtSrcFileName.Location = new System.Drawing.Point(216, 14);
            this.txtSrcFileName.Name = "txtSrcFileName";
            this.txtSrcFileName.ReadOnly = true;
            this.txtSrcFileName.Size = new System.Drawing.Size(452, 21);
            this.txtSrcFileName.TabIndex = 3;
            // 
            // txtDestFileName
            // 
            this.txtDestFileName.Enabled = false;
            this.txtDestFileName.Location = new System.Drawing.Point(216, 48);
            this.txtDestFileName.Name = "txtDestFileName";
            this.txtDestFileName.ReadOnly = true;
            this.txtDestFileName.Size = new System.Drawing.Size(452, 21);
            this.txtDestFileName.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 51);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(203, 12);
            this.label2.TabIndex = 4;
            this.label2.Text = "请选择合并后Excel文件保存的路径：";
            // 
            // btnExecute
            // 
            this.btnExecute.Location = new System.Drawing.Point(685, 81);
            this.btnExecute.Name = "btnExecute";
            this.btnExecute.Size = new System.Drawing.Size(75, 23);
            this.btnExecute.TabIndex = 6;
            this.btnExecute.Text = "执行";
            this.btnExecute.UseVisualStyleBackColor = true;
            this.btnExecute.Click += new System.EventHandler(this.button1_Click);
            // 
            // pb1
            // 
            this.pb1.Location = new System.Drawing.Point(14, 81);
            this.pb1.Name = "pb1";
            this.pb1.Size = new System.Drawing.Size(654, 23);
            this.pb1.TabIndex = 7;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(772, 121);
            this.Controls.Add(this.pb1);
            this.Controls.Add(this.btnExecute);
            this.Controls.Add(this.txtDestFileName);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtSrcFileName);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnDestSelect);
            this.Controls.Add(this.btnSrcSelect);
            this.Name = "MainForm";
            this.Text = "合并Excel中相同账户";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnSrcSelect;
        private System.Windows.Forms.Button btnDestSelect;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtSrcFileName;
        private System.Windows.Forms.TextBox txtDestFileName;
        private System.Windows.Forms.Label label2;
        public System.Windows.Forms.Button btnExecute;
        public System.Windows.Forms.ProgressBar pb1;
    }
}

