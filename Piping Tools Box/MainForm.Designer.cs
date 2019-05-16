namespace Piping_Tools_Box
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
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.tsbSupportContrast = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.tsbPipMaterialCode = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripButton1 = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // toolStrip1
            // 
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsbSupportContrast,
            this.toolStripSeparator1,
            this.tsbPipMaterialCode,
            this.toolStripSeparator2,
            this.toolStripButton1,
            this.toolStripSeparator3});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(667, 93);
            this.toolStrip1.TabIndex = 0;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // tsbSupportContrast
            // 
            this.tsbSupportContrast.AutoSize = false;
            this.tsbSupportContrast.Image = global::Piping_Tools_Box.Properties.Resources.support;
            this.tsbSupportContrast.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.tsbSupportContrast.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tsbSupportContrast.Name = "tsbSupportContrast";
            this.tsbSupportContrast.Size = new System.Drawing.Size(90, 90);
            this.tsbSupportContrast.Text = "支架号对比";
            this.tsbSupportContrast.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.tsbSupportContrast.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.tsbSupportContrast.Click += new System.EventHandler(this.tsbSupportContrast_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 93);
            // 
            // tsbPipMaterialCode
            // 
            this.tsbPipMaterialCode.AccessibleRole = System.Windows.Forms.AccessibleRole.None;
            this.tsbPipMaterialCode.AutoSize = false;
            this.tsbPipMaterialCode.Image = ((System.Drawing.Image)(resources.GetObject("tsbPipMaterialCode.Image")));
            this.tsbPipMaterialCode.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.tsbPipMaterialCode.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tsbPipMaterialCode.Name = "tsbPipMaterialCode";
            this.tsbPipMaterialCode.Overflow = System.Windows.Forms.ToolStripItemOverflow.Never;
            this.tsbPipMaterialCode.Size = new System.Drawing.Size(68, 90);
            this.tsbPipMaterialCode.Text = "文件校验";
            this.tsbPipMaterialCode.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.tsbPipMaterialCode.Click += new System.EventHandler(this.tsbPipMaterialCode_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(6, 93);
            // 
            // toolStripButton1
            // 
            this.toolStripButton1.AutoSize = false;
            this.toolStripButton1.Image = global::Piping_Tools_Box.Properties.Resources.spoolgen;
            this.toolStripButton1.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.toolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton1.Name = "toolStripButton1";
            this.toolStripButton1.Size = new System.Drawing.Size(90, 90);
            this.toolStripButton1.Text = "Spoolgen表格";
            this.toolStripButton1.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.toolStripButton1.Click += new System.EventHandler(this.tsbSpoolgenExcel_Click);
            // 
            // toolStripSeparator3
            // 
            this.toolStripSeparator3.Name = "toolStripSeparator3";
            this.toolStripSeparator3.Size = new System.Drawing.Size(6, 93);
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.ClientSize = new System.Drawing.Size(667, 95);
            this.Controls.Add(this.toolStrip1);
            this.Cursor = System.Windows.Forms.Cursors.Default;
            this.Font = new System.Drawing.Font("Microsoft YaHei", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.ForeColor = System.Drawing.SystemColors.ControlText;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "FormMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Piping Tools";
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton tsbSupportContrast;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripButton tsbPipMaterialCode;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripButton toolStripButton1;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator3;
    }
}

