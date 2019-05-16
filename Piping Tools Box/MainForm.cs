using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Piping_Tools_Box
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void tsbSupportContrast_Click(object sender, EventArgs e)
        {
            SupportContrast surpportContrast = new SupportContrast();
            surpportContrast.ShowDialog();
        }

        private void tsbPipMaterialCode_Click(object sender, EventArgs e)
        {
            PipMaterialCode pipMaterialCode = new PipMaterialCode();
            pipMaterialCode.ShowDialog();
        }       

        private void tsbSpoolgenExcel_Click(object sender, EventArgs e)
        {
            SpoolgenExcel spoolgenExcel = new SpoolgenExcel();
            spoolgenExcel.ShowDialog();
        }
    }
}
