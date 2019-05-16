namespace Piping_Tools_Box
{
    partial class SpoolgenExcel
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SpoolgenExcel));
            this.read = new System.Windows.Forms.Button();
            this.tbBoxPath = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.pgbtrack = new System.Windows.Forms.ProgressBar();
            this.rtbTrack = new System.Windows.Forms.RichTextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.tbTrackCount = new System.Windows.Forms.TextBox();
            this.btnCheck = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.pgbweld = new System.Windows.Forms.ProgressBar();
            this.button1 = new System.Windows.Forms.Button();
            this.tbWeldsCount = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.rtbWelds = new System.Windows.Forms.RichTextBox();
            this.tbTrackFinish = new System.Windows.Forms.TextBox();
            this.tbWeldFinish = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // read
            // 
            this.read.Font = new System.Drawing.Font("Microsoft YaHei", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.read.Location = new System.Drawing.Point(571, 27);
            this.read.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.read.Name = "read";
            this.read.Size = new System.Drawing.Size(98, 32);
            this.read.TabIndex = 4;
            this.read.Text = "读取并导出";
            this.read.UseVisualStyleBackColor = true;
            this.read.Click += new System.EventHandler(this.read_Click);
            // 
            // tbBoxPath
            // 
            this.tbBoxPath.AllowDrop = true;
            this.tbBoxPath.Location = new System.Drawing.Point(131, 30);
            this.tbBoxPath.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tbBoxPath.Name = "tbBoxPath";
            this.tbBoxPath.Size = new System.Drawing.Size(416, 26);
            this.tbBoxPath.TabIndex = 6;
            this.tbBoxPath.DragDrop += new System.Windows.Forms.DragEventHandler(this.tbBoxPath_DragDrop);
            this.tbBoxPath.DragEnter += new System.Windows.Forms.DragEventHandler(this.tbBoxPath_DragEnter);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft YaHei", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(38, 29);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(82, 20);
            this.label1.TabIndex = 7;
            this.label1.Text = "读取文件夹:";
            // 
            // pgbtrack
            // 
            this.pgbtrack.Location = new System.Drawing.Point(225, 345);
            this.pgbtrack.Name = "pgbtrack";
            this.pgbtrack.Size = new System.Drawing.Size(400, 10);
            this.pgbtrack.TabIndex = 8;
            // 
            // rtbTrack
            // 
            this.rtbTrack.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.rtbTrack.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.rtbTrack.Location = new System.Drawing.Point(65, 112);
            this.rtbTrack.Name = "rtbTrack";
            this.rtbTrack.Size = new System.Drawing.Size(306, 120);
            this.rtbTrack.TabIndex = 11;
            this.rtbTrack.Text = "";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(61, 79);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(152, 17);
            this.label2.TabIndex = 12;
            this.label2.Text = "读取到的Material文件名：";
            // 
            // tbTrackCount
            // 
            this.tbTrackCount.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.tbTrackCount.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.tbTrackCount.Location = new System.Drawing.Point(225, 250);
            this.tbTrackCount.Name = "tbTrackCount";
            this.tbTrackCount.Size = new System.Drawing.Size(146, 19);
            this.tbTrackCount.TabIndex = 13;
            // 
            // btnCheck
            // 
            this.btnCheck.Location = new System.Drawing.Point(150, 275);
            this.btnCheck.Name = "btnCheck";
            this.btnCheck.Size = new System.Drawing.Size(144, 28);
            this.btnCheck.TabIndex = 14;
            this.btnCheck.Text = "查看TrackingList";
            this.btnCheck.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btnCheck.UseVisualStyleBackColor = true;
            this.btnCheck.Click += new System.EventHandler(this.btnCheck_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(103, 338);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(116, 17);
            this.label3.TabIndex = 15;
            this.label3.Text = "TrackingList进度条:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.Location = new System.Drawing.Point(103, 364);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(114, 17);
            this.label4.TabIndex = 17;
            this.label4.Text = "WeldingList进度条:";
            // 
            // pgbweld
            // 
            this.pgbweld.Location = new System.Drawing.Point(225, 371);
            this.pgbweld.Name = "pgbweld";
            this.pgbweld.Size = new System.Drawing.Size(400, 10);
            this.pgbweld.TabIndex = 16;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(463, 275);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(144, 28);
            this.button1.TabIndex = 21;
            this.button1.Text = "查看WeldingList";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // tbWeldsCount
            // 
            this.tbWeldsCount.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.tbWeldsCount.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.tbWeldsCount.Location = new System.Drawing.Point(554, 250);
            this.tbWeldsCount.Name = "tbWeldsCount";
            this.tbWeldsCount.Size = new System.Drawing.Size(146, 19);
            this.tbWeldsCount.TabIndex = 20;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label5.Location = new System.Drawing.Point(390, 79);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(140, 17);
            this.label5.TabIndex = 19;
            this.label5.Text = "读取到的Welds文件名：";
            // 
            // rtbWelds
            // 
            this.rtbWelds.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.rtbWelds.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.rtbWelds.Location = new System.Drawing.Point(394, 112);
            this.rtbWelds.Name = "rtbWelds";
            this.rtbWelds.Size = new System.Drawing.Size(306, 120);
            this.rtbWelds.TabIndex = 18;
            this.rtbWelds.Text = "";
            // 
            // tbTrackFinish
            // 
            this.tbTrackFinish.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.tbTrackFinish.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.tbTrackFinish.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.tbTrackFinish.Location = new System.Drawing.Point(631, 338);
            this.tbTrackFinish.Name = "tbTrackFinish";
            this.tbTrackFinish.Size = new System.Drawing.Size(69, 16);
            this.tbTrackFinish.TabIndex = 22;
            // 
            // tbWeldFinish
            // 
            this.tbWeldFinish.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.tbWeldFinish.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.tbWeldFinish.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.tbWeldFinish.Location = new System.Drawing.Point(631, 364);
            this.tbWeldFinish.Name = "tbWeldFinish";
            this.tbWeldFinish.Size = new System.Drawing.Size(69, 16);
            this.tbWeldFinish.TabIndex = 23;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft YaHei", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label6.Location = new System.Drawing.Point(201, 414);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(279, 14);
            this.label6.TabIndex = 24;
            this.label6.Text = "Copyright©2019 BOMESC All Rights Reserved Design By SHEN\r\n";
            // 
            // SpoolgenExcel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.ClientSize = new System.Drawing.Size(752, 437);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.tbWeldFinish);
            this.Controls.Add(this.tbTrackFinish);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.tbWeldsCount);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.rtbWelds);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.pgbweld);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnCheck);
            this.Controls.Add(this.tbTrackCount);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.rtbTrack);
            this.Controls.Add(this.pgbtrack);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tbBoxPath);
            this.Controls.Add(this.read);
            this.Font = new System.Drawing.Font("Microsoft YaHei", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "SpoolgenExcel";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "SpoolgenExcel-BHP--SHEN";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button read;
        private System.Windows.Forms.TextBox tbBoxPath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ProgressBar pgbtrack;
        private System.Windows.Forms.RichTextBox rtbTrack;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox tbTrackCount;
        private System.Windows.Forms.Button btnCheck;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ProgressBar pgbweld;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox tbWeldsCount;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.RichTextBox rtbWelds;
        private System.Windows.Forms.TextBox tbTrackFinish;
        private System.Windows.Forms.TextBox tbWeldFinish;
        private System.Windows.Forms.Label label6;
    }
}

