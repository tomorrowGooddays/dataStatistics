using GDDataStatistics.Model;
using Newtonsoft.Json;
using NPOI.OpenXmlFormats.Shared;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
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
                try
                {
                    this.label1.Text = "文件名称：";
                    this.fileName.Text = dialog.FileName;
                    ShowInfo($"选择文件{dialog.FileName}");
                    ShowInfo($"开始处理文件：{dialog.FileName}。");

                    this.BtnEnabled(false);
                    //ExcelDataFactory.ReadTotalDistrictCodeName(this.fileName.Text);
                    //var dic = ExcelDataFactory.LoadExcelDataDistrict(this.fileName.Text);

                    Dictionary<string, Dictionary<string, double>> dataDic = ExcelDataFactory.LoadExcelData(this.fileName.Text);

                    DataTable dataTable = DataMergeTool.ConvertData(dataDic);

                    dataGridView1.DataSource = dataTable;

                    string filepathAndName = DataMergeTool.ExportData(dialog.FileName, dataDic);

                    ShowInfo($"文件处理完成，请查看导出结果。{filepathAndName}");
                    MessageBox.Show($"文件处理完成，请查看导出结果:{filepathAndName}");
                }
                catch (Exception ex)
                {
                    ShowInfo(ex.Message + ex.StackTrace);
                    MessageBox.Show(ex.Message);
                }
            }

            this.BtnEnabled(true);
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
                    this.label1.Text = "文件夹路径：";
                    this.fileName.Text = dialog.SelectedPath;
                    ShowInfo($"选择文件夹{this.fileName.Text}");

                    //返回指定目录中的文件的名称（绝对路径）
                    string[] files = System.IO.Directory.GetFiles(dialog.SelectedPath);
                    //获取Test文件夹下所有文件名称
                    List<string> filePaths = System.IO.Directory.GetFiles(dialog.SelectedPath, "*.xlsx", System.IO.SearchOption.AllDirectories).ToList();

                    if (filePaths != null && filePaths.Count > 0)
                    {
                        MessageBox.Show($"当前文件夹有:{filePaths.Count()}个文件准备开始处理");

                        this.BtnEnabled(false);

                        #region 全量统计
                        //DataStatisticsByTotal(filePaths);
                        #endregion

                        #region 按区分批统计

                        DataStatisticsByDistrict(filePaths);
                        #endregion
                    }
                    else
                    {
                        MessageBox.Show("当前路径下没有可执行的Excel文件，请确认路径是否选错");
                    }

                    MessageBox.Show("所有文件处理完成，请查看导出结果");

                    this.BtnEnabled(true);
                }
                catch (Exception ex)
                {
                    ShowInfo(ex.Message + ex.StackTrace);
                    MessageBox.Show(ex.Message);

                    this.BtnEnabled(true);
                }
            }

        }

        private void DataStatisticsByTotal(List<string> filePaths)
        {
            List<ExcelDataInfo> dataList = new List<ExcelDataInfo>();
            foreach (var filePath in filePaths)
            {
                ShowInfo($"开始处理文件：{filePath}。");

                string fileName = Path.GetFileNameWithoutExtension(filePath);

                Dictionary<string, Dictionary<string, double>> dataDic = ExcelDataFactory.LoadExcelData(filePath);

                string dataDicString = JsonConvert.SerializeObject(dataDic);
                if (!string.IsNullOrWhiteSpace(dataDicString))
                {
                    Dictionary<string, Dictionary<string, double>> dataJson = JsonConvert.DeserializeObject<Dictionary<string, Dictionary<string, double>>>(dataDicString);

                    ExcelDataInfo dataInfo = new ExcelDataInfo()
                    {
                        FileName = fileName,
                        DataList = dataJson
                    };

                    dataList.Add(dataInfo);
                }
            }

            if (dataList != null && dataList.Count > 0)
            {
                ShowInfo("========开始导出全量统计数据==========");

                DataMergeTool.ExportDataByName(this.fileName.Text, dataList);

                //var dicData = DataMergeTool.MergeData(dataList);
                //把dataTable显示在页面
                //DataTable dataTable = DataMergeTool.ConvertData(dicData);

                //dataGridView1.DataSource = dataTable;

                //string filePathAndName = $"{this.fileName.Text}\\计算结果.xlsx";
                //DataMergeTool.ExportData(filePathAndName, dicData);

                ShowInfo("全量统计数据处理完成");
            }
        }

        private void DataStatisticsByDistrict(List<string> filePaths)
        {
            ShowInfo("开始处理分区数据统计");

            List<ExcelDataDistrictInfo> dataDistrcitList = new List<ExcelDataDistrictInfo>();
            foreach (var filePath in filePaths)
            {
                ShowInfo($"开始处理文件：{filePath}。");

                string fileName = Path.GetFileNameWithoutExtension(filePath);

                ExcelDataFactory.ReadTotalDistrictCodeName(filePath);

                var result = ExcelDataFactory.LoadExcelDataDistrict(filePath);
                Dictionary<string, Dictionary<string, Dictionary<string, double>>> dataDic = result.Item1;
                Dictionary<string, string> dataDic2 = result.Item2;
                Dictionary<string, double> dataDic3 = result.Item3;

                string dataDicString2 = JsonConvert.SerializeObject(dataDic2);
                string dataDicString3 = JsonConvert.SerializeObject(dataDic3);
                string dataDicString = JsonConvert.SerializeObject(dataDic);
                if (!string.IsNullOrWhiteSpace(dataDicString))
                {
                    Dictionary<string, Dictionary<string, Dictionary<string, double>>> dataJson = JsonConvert.DeserializeObject<Dictionary<string, Dictionary<string, Dictionary<string, double>>>>(dataDicString);

                    Dictionary<string, string> dicJson2 = JsonConvert.DeserializeObject<Dictionary<string, string>>(dataDicString2);
                    Dictionary<string, double> dicJson3 = JsonConvert.DeserializeObject<Dictionary<string, double>>(dataDicString3);

                    ExcelDataDistrictInfo dataInfo = new ExcelDataDistrictInfo()
                    {
                        FileName = fileName,
                        DataList = dataJson,
                        DistrictNameAndCodeMap = dicJson2,
                        DistrictTotalAmount = dicJson3
                    };

                    dataDistrcitList.Add(dataInfo);
                }
            }

            if (dataDistrcitList != null && dataDistrcitList.Count > 0)
            {
                ShowInfo("========开始导出分区统计数据==========");

                DataMergeTool.ExportDataByNameWithDistrcit(this.fileName.Text, dataDistrcitList);

                ShowInfo("分区统计数据处理完成");
            }
        }

        private void BtnEnabled(bool enable)
        {
            this.button1.Enabled = enable;
            this.button2.Enabled = enable;

            if (enable)
            {
                this.tabControl1.SelectedTab = this.tabPage1;
            }
            else
            {
                this.tabControl1.SelectedTab = this.tabPage2;
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

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
