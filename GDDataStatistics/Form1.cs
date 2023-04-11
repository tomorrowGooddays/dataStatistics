using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GDDataStatistics
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// 选择文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = true;//该值确定是否可以选择多个文件
            dialog.Title = "请选择文件";
            dialog.Filter = "所有文件(*.*)|*.xlsx";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                this.fileName.Text = dialog.FileName;
                ShowInfo($"选择文件{dialog.FileName}");
            }
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {
        }

        /// <summary>
        /// 选择文件夹
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void fileName_TextChanged(object sender, EventArgs e)
        {

        }


        public void ShowInfo(string msg)
        {
            this.BeginInvoke((Action)(()=>{
                this.log.AppendText($"{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}  {msg}");
                this.log.AppendText(Environment.NewLine);
                this.log.ScrollToCaret();
            }));
        }
    }
}
