using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using K12.Data;
using Aspose.Cells;
using System.IO;

using FISCA.Data;

namespace KaoHsiungJHSemesterYearDomainFailCount
{
    public partial class SemesterSettingForm : BaseForm
    {
        string semester = "";

        public SemesterSettingForm()
        {
            InitializeComponent();
        }

        
        private void buttonX1_Click(object sender, EventArgs e)
        {
            buttonX1.Enabled = false;
            buttonX2.Enabled = false;
            RunWork();
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void SemesterSettingForm_Load(object sender, EventArgs e)
        {
            //加入系統預設定學年 加減二年
            comboBoxEx1.Items.Add(int.Parse(School.DefaultSchoolYear) - 2);
            comboBoxEx1.Items.Add(int.Parse(School.DefaultSchoolYear) - 1);
            comboBoxEx1.Items.Add(int.Parse(School.DefaultSchoolYear));
            //comboBoxEx1.Items.Add(int.Parse(School.DefaultSchoolYear) + 1);
            //comboBoxEx1.Items.Add(int.Parse(School.DefaultSchoolYear) + 2);
        }

        public void RunWork()
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("報表列印中",20);
            BackgroundWorker _BW;
            _BW = new BackgroundWorker();

             semester = "" + comboBoxEx1.SelectedItem;

            _BW.DoWork += new DoWorkEventHandler(_BW_DoWork);
            _BW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_BW_RunWorkerCompleted);
            _BW.RunWorkerAsync();
        }

        private void _BW_DoWork(object sender, DoWorkEventArgs e)
        {
            QueryHelper _Q = new QueryHelper();

            DataTable dt_source;
            
            
            dt_source = _Q.Select( SqlString.sql(semester));

            //開始列印
            Workbook wb = new Workbook();

            Worksheet ws = wb.Worksheets[0];

            ws.Name = "全校學年領域不及格人數";

            Cells cs = ws.Cells;

            int row_index = 0;

            cs[row_index, 0].Value = "學年度";
            cs[row_index, 1].Value = "年級";
            cs[row_index, 2].Value = "年級人數";
            cs[row_index, 3].Value = "語文不及格人數";
            cs[row_index, 4].Value = "數學不及格人數";
            cs[row_index, 5].Value = "社會不及格人數";
            cs[row_index, 6].Value = "自然與生活科技不及格人數";
            cs[row_index, 7].Value = "健康與體育不及格人數";
            cs[row_index, 8].Value = "藝術與人文不及格人數";
            cs[row_index, 9].Value = "綜合活動不及格人數";
            cs[row_index, 10].Value = "不及格領域數1人數";
            cs[row_index, 11].Value = "不及格領域數2人數";
            cs[row_index, 12].Value = "不及格領域數3人數";
            cs[row_index, 13].Value = "不及格領域數4人數";
            cs[row_index, 14].Value = "不及格領域數5人數";
            cs[row_index, 15].Value = "不及格領域數6人數";
            cs[row_index, 16].Value = "不及格領域數7人數";
            cs[row_index, 17].Value = "補考人數";
            cs[row_index, 18].Value = "補考通過人數";

            // 調整寬度
            cs.Columns[3].Width = 20;
            cs.Columns[4].Width = 20;
            cs.Columns[5].Width = 20;
            cs.Columns[6].Width = 20;
            cs.Columns[7].Width = 20;
            cs.Columns[8].Width = 20;
            cs.Columns[9].Width = 20;
            cs.Columns[10].Width = 20;
            cs.Columns[11].Width = 20;
            cs.Columns[12].Width = 20;
            cs.Columns[13].Width = 20;
            cs.Columns[14].Width = 20;
            cs.Columns[15].Width = 20;
            cs.Columns[16].Width = 20;
            cs.Columns[17].Width = 20;
            cs.Columns[18].Width = 20;
            cs.Columns[19].Width = 20;


            row_index++;

            foreach (DataRow dr in dt_source.Rows)
            {
                cs[row_index, 0].Value = dr[0];
                cs[row_index, 1].Value = dr[1];
                cs[row_index, 2].Value = dr[2];
                cs[row_index, 3].Value = dr[3];
                cs[row_index, 4].Value = dr[4];
                cs[row_index, 5].Value = dr[5];
                cs[row_index, 6].Value = dr[6];
                cs[row_index, 7].Value = dr[7];
                cs[row_index, 8].Value = dr[8];
                cs[row_index, 9].Value = dr[9];
                cs[row_index, 10].Value = dr[10];
                cs[row_index, 11].Value = dr[11];
                cs[row_index, 12].Value = dr[12];
                cs[row_index, 13].Value = dr[13];
                cs[row_index, 14].Value = dr[14];
                cs[row_index, 15].Value = dr[15];
                cs[row_index, 16].Value = dr[16];
                cs[row_index, 17].Value = dr[17];
                cs[row_index, 18].Value = dr[18];

                row_index++;
            }
            e.Result = wb;
        }

        private void _BW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Workbook wb = e.Result as Workbook;

            if (wb == null)
                return;

            FISCA.Presentation.MotherForm.SetStatusBarMessage("報表列印完成",100);

            SaveFileDialog save = new SaveFileDialog();
            save.Title = "另存新檔";
            save.FileName = "全校學年領域不及格人數.xls";
            save.Filter = "Excel檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";

            if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    wb.Save(save.FileName, Aspose.Cells.SaveFormat.Excel97To2003);
                    System.Diagnostics.Process.Start(save.FileName);
                }
                catch
                {
                    MessageBox.Show("檔案儲存失敗");
                }
            }

            this.Close();
        }
        private void comboBoxEx1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if ("" + comboBoxEx1.SelectedItem != "")
            {
                buttonX1.Enabled = true;
            }
        }
    }
}
