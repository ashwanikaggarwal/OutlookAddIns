using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookAddIns.Forms
{
    public partial class frmSetPath : Form
    {
        private readonly Properties.Settings settings = Properties.Settings.Default;
        public frmSetPath()
        {
            InitializeComponent();
        }

        private void frmSetPath_Load(object sender, EventArgs e)
        {
            txtPath.Text = settings.Database;
        }

        private void frmSetPath_DragDrop(object sender, DragEventArgs e)
        {
            string[] filePaths = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            txtPath.Text = filePaths[0];
        }

        private void frmSetPath_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            DialogResult fd = filePathDialog.ShowDialog();
            if (fd == DialogResult.OK)
            {
                txtPath.Text = filePathDialog.FileName;
            }
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            settings["Database"] = txtPath.Text;
            settings.Save();
            Close();
        }
    }
}
