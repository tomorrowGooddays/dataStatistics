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
            dialog.Multiselect = false;//该值确定是否可以选择多个文件
            dialog.Title = "请选择文件";
            dialog.Filter = "所有文件(*.*)|*.xlsx";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                this.fileName.Text = dialog.FileName;
                ShowInfo($"选择文件{dialog.FileName}");

                Dictionary<string, Dictionary<string, double>> dataDic = ExcelDataFactory.LoadExcelData(this.fileName.Text);

                DataTable dataTable = DataMergeTool.ConvertData(dataDic);
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
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件路径";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    this.fileName.Text = dialog.SelectedPath;
                    ShowInfo($"选择文件夹{this.fileName.Text}");

                    //返回指定目录中的文件的名称（绝对路径）
                    string[] files = System.IO.Directory.GetFiles(dialog.SelectedPath);
                    //获取Test文件夹下所有文件名称
                    List<string> filePaths = System.IO.Directory.GetFiles(dialog.SelectedPath, "*.xlsx", System.IO.SearchOption.AllDirectories).ToList();

                    if (filePaths != null && filePaths.Count > 0)
                    {
                        MessageBox.Show($"当前文件夹有:{filePaths.Count()}个文件准备开始处理");
                        List<Dictionary<string, Dictionary<string, double>>> dataList = new List<Dictionary<string, Dictionary<string, double>>>();
                        foreach (var filePath in filePaths)
                        {
                            ShowInfo($"开始处理文件：{filePath}。");

                            Dictionary<string, Dictionary<string, double>> dataDic = ExcelDataFactory.LoadExcelData(filePath);

                            if (dataDic != null)
                            {
                                dataList.Add(dataDic);
                            }

                            ShowInfo($"文件：{filePath}处理完成。");
                        }

                        if (dataList != null && dataList.Count > 0)
                        {
                            ShowInfo("========开始合并数据==========");
                            var dicData = DataMergeTool.MergeData(dataList);
                            //把dataTable显示在页面
                            DataTable dataTable = DataMergeTool.ConvertData(dicData);
                        }

                        MessageBox.Show("所有文件处理完成，请查看导出结果");
                    }
                    else
                    {
                        MessageBox.Show("当前路径下没有可执行的Excel文件，请确认路径是否选错");
                    }
                }
                catch (Exception ex)
                {
                    ShowInfo(ex.Message + ex.StackTrace);
                    MessageBox.Show(ex.Message);
                }
            }

        }

        private void fileName_TextChanged(object sender, EventArgs e)
        {

        }


        public void ShowInfo(string msg)
        {
            this.log.AppendText($"{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}  {msg}");
            this.log.AppendText(Environment.NewLine);
            this.log.ScrollToCaret();
        }
    }
}
