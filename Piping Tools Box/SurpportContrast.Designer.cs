namespace Piping_Tools_Box
{
    partial class SupportContrast
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SupportContrast));
            this.label1 = new System.Windows.Forms.Label();
            this.tboldexcel = new System.Windows.Forms.TextBox();
            this.btnoldexcel = new System.Windows.Forms.Button();
            this.btnnewexcel = new System.Windows.Forms.Button();
            this.tbnewexcel = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnStart = new System.Windows.Forms.Button();
            this.btnOpenExcel = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.pgbRun = new System.Windows.Forms.ProgressBar();
            this.runstate = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Cursor = System.Windows.Forms.Cursors.Default;
            this.label1.Location = new System.Drawing.Point(49, 51);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(56, 17);
            this.label1.TabIndex = 0;
            this.label1.Text = "老表格：";
            // 
            // tboldexcel
            // 
            this.tboldexcel.Location = new System.Drawing.Point(104, 47);
            this.tboldexcel.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.tboldexcel.Name = "tboldexcel";
            this.tboldexcel.Size = new System.Drawing.Size(336, 23);
            this.tboldexcel.TabIndex = 2;
            // 
            // btnoldexcel
            // 
            this.btnoldexcel.Location = new System.Drawing.Point(461, 42);
            this.btnoldexcel.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnoldexcel.Name = "btnoldexcel";
            this.btnoldexcel.Size = new System.Drawing.Size(87, 32);
            this.btnoldexcel.TabIndex = 3;
            this.btnoldexcel.Text = "导入老表";
            this.btnoldexcel.UseVisualStyleBackColor = true;
            this.btnoldexcel.Click += new System.EventHandler(this.btnoldexcel_Click);
            // 
            // btnnewexcel
            // 
            this.btnnewexcel.Location = new System.Drawing.Point(461, 103);
            this.btnnewexcel.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnnewexcel.Name = "btnnewexcel";
            this.btnnewexcel.Size = new System.Drawing.Size(87, 32);
            this.btnnewexcel.TabIndex = 6;
            this.btnnewexcel.Text = "导入新表";
            this.btnnewexcel.UseVisualStyleBackColor = true;
            this.btnnewexcel.Click += new System.EventHandler(this.btnnewexcel_Click);
            // 
            // tbnewexcel
            // 
            this.tbnewexcel.Location = new System.Drawing.Point(104, 108);
            this.tbnewexcel.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.tbnewexcel.Name = "tbnewexcel";
            this.tbnewexcel.Size = new System.Drawing.Size(336, 23);
            this.tbnewexcel.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(49, 113);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(56, 17);
            this.label2.TabIndex = 4;
            this.label2.Text = "新表格：";
            // 
            // btnStart
            // 
            this.btnStart.Location = new System.Drawing.Point(461, 153);
            this.btnStart.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(87, 32);
            this.btnStart.TabIndex = 7;
            this.btnStart.Text = "开始";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // btnOpenExcel
            // 
            this.btnOpenExcel.Location = new System.Drawing.Point(461, 199);
            this.btnOpenExcel.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnOpenExcel.Name = "btnOpenExcel";
            this.btnOpenExcel.Size = new System.Drawing.Size(87, 32);
            this.btnOpenExcel.TabIndex = 8;
            this.btnOpenExcel.Text = "打开文件";
            this.btnOpenExcel.UseVisualStyleBackColor = true;
            this.btnOpenExcel.Click += new System.EventHandler(this.btnOpenExcel_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(49, 161);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(56, 17);
            this.label3.TabIndex = 9;
            this.label3.Text = "进度条：";
            // 
            // pgbRun
            // 
            this.pgbRun.Location = new System.Drawing.Point(104, 168);
            this.pgbRun.Name = "pgbRun";
            this.pgbRun.Size = new System.Drawing.Size(336, 10);
            this.pgbRun.TabIndex = 10;
            // 
            // runstate
            // 
            this.runstate.AutoSize = true;
            this.runstate.Location = new System.Drawing.Point(364, 199);
            this.runstate.Name = "runstate";
            this.runstate.Size = new System.Drawing.Size(0, 17);
            this.runstate.TabIndex = 11;
            // 
            // SupportContrast
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.ClientSize = new System.Drawing.Size(583, 262);
            this.Controls.Add(this.runstate);
            this.Controls.Add(this.pgbRun);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnOpenExcel);
            this.Controls.Add(this.btnStart);
            this.Controls.Add(this.btnnewexcel);
            this.Controls.Add(this.tbnewexcel);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnoldexcel);
            this.Controls.Add(this.tboldexcel);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "SupportContrast";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "SupportContrast--SHEN";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tboldexcel;
        private System.Windows.Forms.Button btnoldexcel;
        private System.Windows.Forms.Button btnnewexcel;
        private System.Windows.Forms.TextBox tbnewexcel;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.Button btnOpenExcel;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ProgressBar pgbRun;
        private System.Windows.Forms.Label runstate;
    }
}