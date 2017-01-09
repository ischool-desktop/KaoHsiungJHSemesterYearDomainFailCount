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
using Aspose.Cells;
using System.IO;
using JHSchool.Data;
using K12.Data;


namespace KaoHsiungJHSemesterYearDomainFailCount
{

    // 2016/5/30 穎驊開始正式製作高雄 統計全校學期領域不及格人數報表

    public partial class Mention_To_Start : BaseForm
    {

        private List<string> StudentIDs;
        int? _schoolYear;
        int? _semester;
        List<K12.Data.StudentRecord> StudentList;
        List<String> DomainList = new List<string>();
        BackgroundWorker _BW;


        //decimal DomainScore1 = 0;
        //decimal DomainScore2 = 0;
        //decimal FirstCredit = 0;
        //decimal SecondCredit = 0;
        //decimal SchoolYearDomainScore = 0;

        // 一年級學生總數
        int Grade1StudentNumber = 0;
        // 二年級學生總數
        int Grade2StudentNumber = 0;
        // 三年級學生總數
        int Grade3StudentNumber = 0;

        // 全學年 國語文領域成績不及格人數
        int Chinese_DomainScore_FailedCount = 0;
        // 全學年 語文領域成績不及格人數
        int English_DomainScore_FailedCount = 0;
        // 全學年 數學領域成績不及格人數
        int Math_DomainScore_FailedCount = 0;
        // 全學年 社會領域成績不及格人數
        int Social_DomainScore_FailedCount = 0;
        // 全學年 自然與生活科技領域成績不及格人數
        int Nature_Tech_DomainScore_FailedCount = 0;
        // 全學年 藝術與人文領域成績不及格人數
        int Art_Humanity_DomainScore_FailedCount = 0;
        // 全學年 健康與體育領域成績不及格人數
        int Hygiene_PE_DomainScore_FailedCount = 0;
        // 全學年 綜合領域成績不及格人數
        int Integrated_Activities_Domain_FailedCount = 0;
        // 全學年 語文領域成績不及格人數
        int Language_DomainScore_FailedCount = 0;

        // 上學期資料
        Dictionary<String, JHSemesterScoreRecord> JSR_Arrange1 = new Dictionary<string, JHSemesterScoreRecord>();

        //下學期資料
        Dictionary<String, JHSemesterScoreRecord> JSR_Arrange2 = new Dictionary<string, JHSemesterScoreRecord>();

        // 計算一年級領域不及格者Dictionary
        Dictionary<String, int> TotalDomainFailCountBook_Grade1 = new Dictionary<string, int>();

        // 計算二年級領域不及格者Dictionary
        Dictionary<String, int> TotalDomainFailCountBook_Grade2 = new Dictionary<string, int>();

        // 計算三年級領域不及格者Dictionary
        Dictionary<String, int> TotalDomainFailCountBook_Grade3 = new Dictionary<string, int>();

        //計算各年級的人數
        Dictionary<String, int> EachGradeStudentCount = new Dictionary<string, int>();

        //紀錄領域不及格者之詳情，作為報表輸出的詳情使用
        Dictionary<String, StudentDomainFailRecord> StudentFailedDomainRecordList = new Dictionary<string, StudentDomainFailRecord>();

        //紀錄領域有補考紀錄者之詳情，作為報表輸出的詳情使用
        Dictionary<String, StudentMakeUpScoreRecord> StudentMakeUpDomainRecordList = new Dictionary<string, StudentMakeUpScoreRecord>();
        
        //紀錄一年級的領域不及格積累次數 ， EX: student.ID = 12345 次數 : 五個領域不及格  >> 記錄成[12345,5]
        Dictionary<String, int> EachGradeStudentAccumulatedFailCount_Grade1 = new Dictionary<string, int>();

        //紀錄二年級的領域不及格積累次數 
        Dictionary<String, int> EachGradeStudentAccumulatedFailCount_Grade2 = new Dictionary<string, int>();

        //紀錄三年級的領域不及格積累次數 
        Dictionary<String, int> EachGradeStudentAccumulatedFailCount_Grade3 = new Dictionary<string, int>();

        //計算各年級預警、輔導及補救措施成效的人數(原本 學年成績不及格，經加入補考成績後變成及格)
        Dictionary<String, int> EachGradeMakeUpStudentCount = new Dictionary<string, int>();

        //計算各年級預警、輔導及補救措施成效的成功人數
        Dictionary<String, int> EachGradeMakeUpStudentCount_sucessful = new Dictionary<string, int>();

        // 計入已有補考紀錄 之學生ID List 用來排除 重覆學生
        List<string> Has_MakeUpStudentIDList = new List<string>();
    
        public Mention_To_Start()
        {
            InitializeComponent();

            // 選擇全部的學生 
            List<JHStudentRecord> allstudent = JHStudent.SelectAll();

            List<string> gyStudent = new List<string>();

            foreach (JHStudentRecord each in allstudent)
            {
                if (each.Status == 0 && each.Class != null)
                {
                    gyStudent.Add(each.ID);
                }
            }


            StudentIDs = gyStudent;

            // 取得該學校當學年
            _schoolYear = Int32.Parse(K12.Data.School.DefaultSchoolYear);

            //取得該學校當學期 (本程式目前未用到，因為是算學年領域成績)
            _semester = Int32.Parse(K12.Data.School.DefaultSemester);

            StudentList = K12.Data.Student.SelectByIDs(StudentIDs);

            _BW = new BackgroundWorker();

            //  加入所有的領域名稱 (高雄地區國中為九大領域)
            //DomainList.Add("國語文");
            //DomainList.Add("英語");
            DomainList.Add("語文");
            DomainList.Add("數學");
            DomainList.Add("社會");
            DomainList.Add("自然與生活科技");
            DomainList.Add("藝術與人文");
            DomainList.Add("健康與體育");
            DomainList.Add("綜合活動");


            //加入 所有的領域名稱，作為分開各年級數各領域不及格人數用，目前報表輸出，已將國語文、英語合併為"語文"一領域，故註解掉

            //TotalDomainFailCountBook_Grade1.Add("國語文", Chinese_DomainScore_FailedCount);
            //TotalDomainFailCountBook_Grade1.Add("英語", English_DomainScore_FailedCount);
            TotalDomainFailCountBook_Grade1.Add("數學", Math_DomainScore_FailedCount);
            TotalDomainFailCountBook_Grade1.Add("社會", Social_DomainScore_FailedCount);
            TotalDomainFailCountBook_Grade1.Add("自然與生活科技", Nature_Tech_DomainScore_FailedCount);
            TotalDomainFailCountBook_Grade1.Add("藝術與人文", Art_Humanity_DomainScore_FailedCount);
            TotalDomainFailCountBook_Grade1.Add("健康與體育", Hygiene_PE_DomainScore_FailedCount);
            TotalDomainFailCountBook_Grade1.Add("綜合活動", Integrated_Activities_Domain_FailedCount);
            TotalDomainFailCountBook_Grade1.Add("語文", Language_DomainScore_FailedCount);

            //TotalDomainFailCountBook_Grade2.Add("國語文", Chinese_DomainScore_FailedCount);
            //TotalDomainFailCountBook_Grade2.Add("英語", English_DomainScore_FailedCount);
            TotalDomainFailCountBook_Grade2.Add("數學", Math_DomainScore_FailedCount);
            TotalDomainFailCountBook_Grade2.Add("社會", Social_DomainScore_FailedCount);
            TotalDomainFailCountBook_Grade2.Add("自然與生活科技", Nature_Tech_DomainScore_FailedCount);
            TotalDomainFailCountBook_Grade2.Add("藝術與人文", Art_Humanity_DomainScore_FailedCount);
            TotalDomainFailCountBook_Grade2.Add("健康與體育", Hygiene_PE_DomainScore_FailedCount);
            TotalDomainFailCountBook_Grade2.Add("綜合活動", Integrated_Activities_Domain_FailedCount);
            TotalDomainFailCountBook_Grade2.Add("語文", Language_DomainScore_FailedCount);

            //TotalDomainFailCountBook_Grade3.Add("國語文", Chinese_DomainScore_FailedCount);
            //TotalDomainFailCountBook_Grade3.Add("英語", English_DomainScore_FailedCount);
            TotalDomainFailCountBook_Grade3.Add("數學", Math_DomainScore_FailedCount);
            TotalDomainFailCountBook_Grade3.Add("社會", Social_DomainScore_FailedCount);
            TotalDomainFailCountBook_Grade3.Add("自然與生活科技", Nature_Tech_DomainScore_FailedCount);
            TotalDomainFailCountBook_Grade3.Add("藝術與人文", Art_Humanity_DomainScore_FailedCount);
            TotalDomainFailCountBook_Grade3.Add("健康與體育", Hygiene_PE_DomainScore_FailedCount);
            TotalDomainFailCountBook_Grade3.Add("綜合活動", Integrated_Activities_Domain_FailedCount);
            TotalDomainFailCountBook_Grade3.Add("語文", Language_DomainScore_FailedCount);

            //每個年級 人數統計
            EachGradeStudentCount.Add("Grade1", Grade1StudentNumber);
            EachGradeStudentCount.Add("Grade2", Grade2StudentNumber);
            EachGradeStudentCount.Add("Grade3", Grade3StudentNumber);

            //每個年級 有補考紀錄的人數統計
            EachGradeMakeUpStudentCount.Add("Grade1",0);
            EachGradeMakeUpStudentCount.Add("Grade2", 0);
            EachGradeMakeUpStudentCount.Add("Grade3", 0);

            //每個年級 有補考紀錄、且最終學年成績平均及格(>60分) 的人數統計
            EachGradeMakeUpStudentCount_sucessful.Add("Grade1", 0);
            EachGradeMakeUpStudentCount_sucessful.Add("Grade2", 0);
            EachGradeMakeUpStudentCount_sucessful.Add("Grade3", 0);

            foreach(var sr in StudentList)
            {
                //整理一年級不及格領域累積次數
                if (sr.Class != null && (sr.Class.GradeYear == 1 || sr.Class.GradeYear == 7))                
                {
                    EachGradeStudentAccumulatedFailCount_Grade1.Add(sr.ID, 0);                
                }

                //整理二年級不及格領域累積次數
                if (sr.Class != null && (sr.Class.GradeYear == 2 || sr.Class.GradeYear == 8))
                {
                    EachGradeStudentAccumulatedFailCount_Grade2.Add(sr.ID, 0);
                }

                //整理三年級不及格領域累積次數
                if (sr.Class != null && (sr.Class.GradeYear == 2 || sr.Class.GradeYear == 9))
                {
                    EachGradeStudentAccumulatedFailCount_Grade3.Add(sr.ID, 0);
                }               
            }            
        }

        private void _BW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("報表列印完成");
            Workbook wb = e.Result as Workbook;

            if (wb == null)
                return;

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
        }

        private void _BW_DoWork(object sender, DoWorkEventArgs e)
        {
            GetStudentDomainScore();

            //開始列印
            Workbook wb = new Workbook(new MemoryStream(Properties.Resources.計算全校學年領域不及格人數));
            //wb.Worksheets.AddCopy(0);
            
            wb.Worksheets[0].Name = "全校全年級各領域不及格人數統計表";
            wb.Worksheets[1].Name = "學年領域不及格學生明細";
            wb.Worksheets[2].Name = "學習情形及補救一覽表OK (含偏遠)";
            wb.Worksheets[3].Name = "補考學生資料明細";

            #region 第一頁
            Cells cs0 = wb.Worksheets[0].Cells;

            cs0[1, 3].PutValue(EachGradeStudentCount["Grade1"]);
            cs0[2, 3].PutValue(TotalDomainFailCountBook_Grade1["語文"]);
            cs0[3, 3].PutValue(TotalDomainFailCountBook_Grade1["數學"]);
            cs0[4, 3].PutValue(TotalDomainFailCountBook_Grade1["社會"]);
            cs0[5, 3].PutValue(TotalDomainFailCountBook_Grade1["自然與生活科技"]);
            cs0[6, 3].PutValue(TotalDomainFailCountBook_Grade1["藝術與人文"]);
            cs0[7, 3].PutValue(TotalDomainFailCountBook_Grade1["健康與體育"]);
            cs0[8, 3].PutValue(TotalDomainFailCountBook_Grade1["綜合活動"]);

            cs0[11, 3].PutValue(EachGradeStudentCount["Grade2"]);
            cs0[12, 3].PutValue(TotalDomainFailCountBook_Grade2["語文"]);
            cs0[13, 3].PutValue(TotalDomainFailCountBook_Grade2["數學"]);
            cs0[14, 3].PutValue(TotalDomainFailCountBook_Grade2["社會"]);
            cs0[15, 3].PutValue(TotalDomainFailCountBook_Grade2["自然與生活科技"]);
            cs0[16, 3].PutValue(TotalDomainFailCountBook_Grade2["藝術與人文"]);
            cs0[17, 3].PutValue(TotalDomainFailCountBook_Grade2["健康與體育"]);
            cs0[18, 3].PutValue(TotalDomainFailCountBook_Grade2["綜合活動"]);

            cs0[21, 3].PutValue(EachGradeStudentCount["Grade3"]);
            cs0[22, 3].PutValue(TotalDomainFailCountBook_Grade3["語文"]);
            cs0[23, 3].PutValue(TotalDomainFailCountBook_Grade3["數學"]);
            cs0[24, 3].PutValue(TotalDomainFailCountBook_Grade3["社會"]);
            cs0[25, 3].PutValue(TotalDomainFailCountBook_Grade3["自然與生活科技"]);
            cs0[26, 3].PutValue(TotalDomainFailCountBook_Grade3["藝術與人文"]);
            cs0[27, 3].PutValue(TotalDomainFailCountBook_Grade3["健康與體育"]);
            cs0[28, 3].PutValue(TotalDomainFailCountBook_Grade3["綜合活動"]); 
            #endregion

            #region 第二頁

            Cells cs1 = wb.Worksheets[1].Cells;

            //領域不及格者計數器
            int FailStudentCounter = 0;

            foreach (String FailedStudent in StudentFailedDomainRecordList.Keys)
            {
                cs1[FailStudentCounter + 1, 0].PutValue(StudentFailedDomainRecordList[FailedStudent].Grade);
                cs1[FailStudentCounter + 1, 1].PutValue(StudentFailedDomainRecordList[FailedStudent].Class);
                cs1[FailStudentCounter + 1, 2].PutValue(StudentFailedDomainRecordList[FailedStudent].Student);
                cs1[FailStudentCounter + 1, 3].PutValue(StudentFailedDomainRecordList[FailedStudent].SchoolYear);
                cs1[FailStudentCounter + 1, 4].PutValue(StudentFailedDomainRecordList[FailedStudent].FailedDomainName);
                cs1[FailStudentCounter + 1, 5].PutValue(StudentFailedDomainRecordList[FailedStudent]._first_domain_score);
                cs1[FailStudentCounter + 1, 6].PutValue(StudentFailedDomainRecordList[FailedStudent]._first_domain_credit);
                cs1[FailStudentCounter + 1, 7].PutValue(StudentFailedDomainRecordList[FailedStudent]._second_domain_score);
                cs1[FailStudentCounter + 1, 8].PutValue(StudentFailedDomainRecordList[FailedStudent]._second_domain_credit);
                cs1[FailStudentCounter + 1, 9].PutValue(StudentFailedDomainRecordList[FailedStudent]._school_year_domain_score);

                if (StudentFailedDomainRecordList[FailedStudent]._memo != null)
                {
                    cs1[FailStudentCounter + 1, 10].PutValue(StudentFailedDomainRecordList[FailedStudent]._memo);
                }
                FailStudentCounter++;


            } 
            #endregion

            //wb.Worksheets.RemoveAt(0);

            #region 第三頁
		  Cells cs2 = wb.Worksheets[2].Cells;


            int fail_1_grade_1 = 0;
            int fail_2_grade_1 = 0;
            int fail_3_grade_1 = 0;
            int fail_4_grade_1 = 0;
            int fail_5_grade_1 = 0;
            int fail_6_grade_1 = 0;
            int fail_7_grade_1 = 0;
            int fail_4up_grade_1 = 0;  // 一年級 4個領域以上不及格人數

            int fail_1_grade_2 = 0;
            int fail_2_grade_2 = 0;
            int fail_3_grade_2 = 0;
            int fail_4_grade_2 = 0;
            int fail_5_grade_2 = 0;
            int fail_6_grade_2 = 0;
            int fail_7_grade_2 = 0;
            int fail_4up_grade_2 = 0;   // 二年級 4個領域以上不及格人數

            int fail_1_grade_3 = 0;
            int fail_2_grade_3 = 0;
            int fail_3_grade_3 = 0;
            int fail_4_grade_3 = 0;
            int fail_5_grade_3 = 0;
            int fail_6_grade_3 = 0;
            int fail_7_grade_3 = 0;
            int fail_4up_grade_3 = 0;   // 三年級 4個領域以上不及格人數

            // 處理一年級累積不及格領域筆數人次
            #region 一年級
            foreach (var AccumulatedfailRecord in EachGradeStudentAccumulatedFailCount_Grade1)
            {
                if (AccumulatedfailRecord.Value == 1)
                {
                    fail_1_grade_1++;
                }
                if (AccumulatedfailRecord.Value == 2)
                {
                    fail_2_grade_1++;
                }
                if (AccumulatedfailRecord.Value == 3)
                {
                    fail_3_grade_1++;
                }
                if (AccumulatedfailRecord.Value == 4)
                {
                    fail_4_grade_1++;
                }
                if (AccumulatedfailRecord.Value == 5)
                {
                    fail_5_grade_1++;
                }
                if (AccumulatedfailRecord.Value == 6)
                {
                    fail_6_grade_1++;
                }
                if (AccumulatedfailRecord.Value == 7)
                {
                    fail_7_grade_1++;
                }
            }
            //一年級 四領域以上不及格人次
            fail_4up_grade_1 = fail_4_grade_1 + fail_5_grade_1 + fail_6_grade_1 + fail_7_grade_1; 
            #endregion

            // 處理二年級累積不及格領域筆數人次
            #region 二年級
            foreach (var AccumulatedfailRecord in EachGradeStudentAccumulatedFailCount_Grade2)
            {
                if (AccumulatedfailRecord.Value == 1)
                {
                    fail_1_grade_2++;
                }
                if (AccumulatedfailRecord.Value == 2)
                {
                    fail_2_grade_2++;
                }
                if (AccumulatedfailRecord.Value == 3)
                {
                    fail_3_grade_2++;
                }
                if (AccumulatedfailRecord.Value == 4)
                {
                    fail_4_grade_2++;
                }
                if (AccumulatedfailRecord.Value == 5)
                {
                    fail_5_grade_2++;
                }
                if (AccumulatedfailRecord.Value == 6)
                {
                    fail_6_grade_2++;
                }
                if (AccumulatedfailRecord.Value == 7)
                {
                    fail_7_grade_2++;
                }
            }
            //二年級 四領域以上不及格人次
            fail_4up_grade_2 = fail_4_grade_2 + fail_5_grade_2 + fail_6_grade_2 + fail_7_grade_2; 
            #endregion

            // 處理三年級累積不及格領域筆數人次
            #region 三年級
            foreach (var AccumulatedfailRecord in EachGradeStudentAccumulatedFailCount_Grade3)
            {
                if (AccumulatedfailRecord.Value == 1)
                {
                    fail_1_grade_3++;
                }
                if (AccumulatedfailRecord.Value == 2)
                {
                    fail_2_grade_3++;
                }
                if (AccumulatedfailRecord.Value == 3)
                {
                    fail_3_grade_3++;
                }
                if (AccumulatedfailRecord.Value == 4)
                {
                    fail_4_grade_3++;
                }
                if (AccumulatedfailRecord.Value == 5)
                {
                    fail_5_grade_3++;
                }
                if (AccumulatedfailRecord.Value == 6)
                {
                    fail_6_grade_3++;
                }
                if (AccumulatedfailRecord.Value == 7)
                {
                    fail_7_grade_3++;
                }
            }
            //三年級 四領域以上不及格人次
            fail_4up_grade_3 = fail_4_grade_3 + fail_5_grade_3 + fail_6_grade_3 + fail_7_grade_3; 
            #endregion


            //標題(故意留"○○縣(市)" ，是為了以後還有機會給其他縣市用到)
            cs2[0, 0].PutValue("○○縣(市)" + _schoolYear + "學年度國民中小學學生學習情形及補教教學成效一覽表");

            cs2[10, 2].PutValue(EachGradeStudentCount["Grade1"]);
            cs2[11, 2].PutValue(EachGradeStudentCount["Grade2"]);
            cs2[12, 2].PutValue(EachGradeStudentCount["Grade3"]);
            cs2[13, 2].PutValue(EachGradeStudentCount["Grade1"] + EachGradeStudentCount["Grade3"] + EachGradeStudentCount["Grade3"]);


            // 一年級 、七年級
            cs2[10, 3].PutValue(TotalDomainFailCountBook_Grade1["語文"]);
            cs2[10, 4].PutValue(TotalDomainFailCountBook_Grade1["數學"]);
            cs2[10, 5].PutValue(TotalDomainFailCountBook_Grade1["社會"]);
            cs2[10, 6].PutValue(TotalDomainFailCountBook_Grade1["自然與生活科技"]);
            cs2[10, 7].PutValue(TotalDomainFailCountBook_Grade1["藝術與人文"]);
            cs2[10, 8].PutValue(TotalDomainFailCountBook_Grade1["健康與體育"]);
            cs2[10, 9].PutValue(TotalDomainFailCountBook_Grade1["綜合活動"]);

            // 二年級 、八年級
            cs2[11, 3].PutValue(TotalDomainFailCountBook_Grade2["語文"]);
            cs2[11, 4].PutValue(TotalDomainFailCountBook_Grade2["數學"]);
            cs2[11, 5].PutValue(TotalDomainFailCountBook_Grade2["社會"]);
            cs2[11, 6].PutValue(TotalDomainFailCountBook_Grade2["自然與生活科技"]);
            cs2[11, 7].PutValue(TotalDomainFailCountBook_Grade2["藝術與人文"]);
            cs2[11, 8].PutValue(TotalDomainFailCountBook_Grade2["健康與體育"]);
            cs2[11, 9].PutValue(TotalDomainFailCountBook_Grade2["綜合活動"]);

            // 三年級 、九年級
            cs2[12, 3].PutValue(TotalDomainFailCountBook_Grade3["語文"]);
            cs2[12, 4].PutValue(TotalDomainFailCountBook_Grade3["數學"]);
            cs2[12, 5].PutValue(TotalDomainFailCountBook_Grade3["社會"]);
            cs2[12, 6].PutValue(TotalDomainFailCountBook_Grade3["自然與生活科技"]);
            cs2[12, 7].PutValue(TotalDomainFailCountBook_Grade3["藝術與人文"]);
            cs2[12, 8].PutValue(TotalDomainFailCountBook_Grade3["健康與體育"]);
            cs2[12, 9].PutValue(TotalDomainFailCountBook_Grade3["綜合活動"]);

            // 一年級 、七年級 學生不及格領域數 累計記數
            cs2[10, 10].PutValue(fail_1_grade_1);
            cs2[10, 11].PutValue(fail_2_grade_1);
            cs2[10, 12].PutValue(fail_3_grade_1);
            cs2[10, 13].PutValue(fail_4_grade_1);
            cs2[10, 14].PutValue(fail_5_grade_1);
            cs2[10, 15].PutValue(fail_6_grade_1);
            cs2[10, 16].PutValue(fail_7_grade_1);
            cs2[10, 17].PutValue(fail_4up_grade_1);
            cs2[10, 18].PutValue(decimal.Parse(""+ fail_4up_grade_1) / decimal.Parse(""+ EachGradeStudentCount["Grade1"]));     //比率

            // 二年級 、八年級 學生不及格領域數 累計記數
            cs2[11, 10].PutValue(fail_1_grade_2);
            cs2[11, 11].PutValue(fail_2_grade_2);
            cs2[11, 12].PutValue(fail_3_grade_2);
            cs2[11, 13].PutValue(fail_4_grade_2);
            cs2[11, 14].PutValue(fail_5_grade_2);
            cs2[11, 15].PutValue(fail_6_grade_2);
            cs2[11, 16].PutValue(fail_7_grade_2);
            cs2[11, 17].PutValue(fail_4up_grade_2);
            cs2[11, 18].PutValue(decimal.Parse("" + fail_4up_grade_2) / decimal.Parse("" + EachGradeStudentCount["Grade2"]));   //比率

            // 三年級 、九年級 學生不及格領域數 累計記數
            cs2[12, 10].PutValue(fail_1_grade_3);
            cs2[12, 11].PutValue(fail_2_grade_3);
            cs2[12, 12].PutValue(fail_3_grade_3);
            cs2[12, 13].PutValue(fail_4_grade_3);
            cs2[12, 14].PutValue(fail_5_grade_3);
            cs2[12, 15].PutValue(fail_6_grade_3);
            cs2[12, 16].PutValue(fail_7_grade_3);
            cs2[12, 17].PutValue(fail_4up_grade_3);
            cs2[12, 18].PutValue(decimal.Parse("" + fail_4up_grade_3) / decimal.Parse("" + EachGradeStudentCount["Grade3"]));   //比率

            // 一年級 、七年級 補救人數、成功人數
            cs2[10, 22].PutValue(EachGradeMakeUpStudentCount["Grade1"]);
            cs2[10, 23].PutValue(EachGradeMakeUpStudentCount_sucessful["Grade1"]);

            // 二年級 、八年級 補救人數、成功人數
            cs2[11, 22].PutValue(EachGradeMakeUpStudentCount["Grade2"]);
            cs2[11, 23].PutValue(EachGradeMakeUpStudentCount_sucessful["Grade2"]);

            // 三年級 、九年級 補救人數、成功人數
            cs2[12, 22].PutValue(EachGradeMakeUpStudentCount["Grade3"]);
            cs2[12, 23].PutValue(EachGradeMakeUpStudentCount_sucessful["Grade3"]); 
	#endregion


            #region 第四頁
            Cells cs3 = wb.Worksheets[3].Cells;

            //領域補考者計數器
            int MakeUpStudentCounter = 0;

            foreach (String MakeUpStudent in StudentMakeUpDomainRecordList.Keys)
            {
                cs3[MakeUpStudentCounter + 1, 0].PutValue(StudentMakeUpDomainRecordList[MakeUpStudent].Grade);
                cs3[MakeUpStudentCounter + 1, 1].PutValue(StudentMakeUpDomainRecordList[MakeUpStudent].Class);
                cs3[MakeUpStudentCounter + 1, 2].PutValue(StudentMakeUpDomainRecordList[MakeUpStudent].Student);
                cs3[MakeUpStudentCounter + 1, 3].PutValue(StudentMakeUpDomainRecordList[MakeUpStudent].SchoolYear);
                cs3[MakeUpStudentCounter + 1, 4].PutValue(StudentMakeUpDomainRecordList[MakeUpStudent].MakeUpDomainName);
                cs3[MakeUpStudentCounter + 1, 5].PutValue(StudentMakeUpDomainRecordList[MakeUpStudent]._first_domain_score);
                cs3[MakeUpStudentCounter + 1, 6].PutValue(StudentMakeUpDomainRecordList[MakeUpStudent]._first_domain_origin_score);
                cs3[MakeUpStudentCounter + 1, 7].PutValue(StudentMakeUpDomainRecordList[MakeUpStudent]._first_domain_makeup_score);
                cs3[MakeUpStudentCounter + 1, 8].PutValue(StudentMakeUpDomainRecordList[MakeUpStudent]._first_domain_makeup_credit);
                cs3[MakeUpStudentCounter + 1, 9].PutValue(StudentMakeUpDomainRecordList[MakeUpStudent]._second_domain_score);
                cs3[MakeUpStudentCounter + 1, 10].PutValue(StudentMakeUpDomainRecordList[MakeUpStudent]._second_domain_origin_score);
                cs3[MakeUpStudentCounter + 1, 11].PutValue(StudentMakeUpDomainRecordList[MakeUpStudent]._second_domain_makeup_score);
                cs3[MakeUpStudentCounter + 1, 12].PutValue(StudentMakeUpDomainRecordList[MakeUpStudent]._second_domain_makeup_credit);

                cs3[MakeUpStudentCounter + 1, 13].PutValue(StudentMakeUpDomainRecordList[MakeUpStudent]._school_year_domain_score);

                if (StudentMakeUpDomainRecordList[MakeUpStudent]._memo != null)
                {
                    cs1[MakeUpStudentCounter + 1, 14].PutValue(StudentMakeUpDomainRecordList[MakeUpStudent]._memo);
                }

                MakeUpStudentCounter++;
            } 
            #endregion
            
            wb.Worksheets.ActiveSheetIndex = 0;
            e.Result = wb;
        }




        public void GetStudentDomainScore()
        {

            // 處理學生的每一筆成績，處理好後，分上下學期，使用特定的KEY值，分別加入上下學期成績的Dictionary。
            foreach (JHSemesterScoreRecord jsr in JHSchool.Data.JHSemesterScore.SelectByStudentIDs(StudentIDs))
            {
                // 該比學期成績資料，學年必須等於當下學期的學年，且學生的目前狀態是一般生(Status == 0)
                if (jsr.SchoolYear == _schoolYear && jsr.Student.Status == 0)
                {
                    if (jsr.Semester == 1)
                    {   
                        //  使用學生ID+ 底線 _ + 學年 可以保證這KEY 獨一無二，(記得底線_要加，否則電腦分不清99 還是9+9)
                        String Sem1Key1 = jsr.RefStudentID + "_" + jsr.SchoolYear + "1";

                        JSR_Arrange1.Add(Sem1Key1, jsr);
                    }

                    if (jsr.Semester == 2)
                    {
                        String Sem2Key2 = jsr.RefStudentID + "_" + jsr.SchoolYear + "2";

                        JSR_Arrange2.Add(Sem2Key2, jsr);
                    }
                }
            }

            // 處理目前所有學生，假如狀態是一般生，且有班級，取得KEY值，進一步去做領域計算
            foreach (StudentRecord sr in StudentList)
            {
                // 計算 補考成績是否有過
                bool pass_makeup = false;

                if (sr.Status == 0 && sr.Class != null)
                {
                    String Sem1Key1 = sr.ID + "_" + _schoolYear + "1";

                    String Sem2Key2 = sr.ID + "_" + _schoolYear + "2";

                    //使用DomainList去查出每一筆成績
                    foreach (String DomainName in DomainList)
                    {
                        if (JSR_Arrange1.ContainsKey(Sem1Key1) && JSR_Arrange2.ContainsKey(Sem2Key2))
                        {
                            // 使用EachJSRCalculator方法，傳入上下學期成績資料、領域名稱，判斷該學年該學生該領域成績有沒有及格
                            bool pass = EachJSRCalculator(JSR_Arrange1[Sem1Key1], JSR_Arrange2[Sem2Key2], DomainName);

                            if (!pass_makeup) // 只要有一領域 補考成績過了(pass_makeup ==true) 就不會再進來重算(就算 此學生 有兩個以上的補考成績也是一樣，因為所要統計的是"人頭"不是人次。)
                            {
                                pass_makeup = EachJSRCalculator_For_MakeUp(JSR_Arrange1[Sem1Key1], JSR_Arrange2[Sem2Key2], DomainName);
                            }
                            
                            //不及格就進行數數
                            //一年級、七年級
                            if (!pass && (sr.Class.GradeYear == 1 || sr.Class.GradeYear == 7))
                            {
                                TotalDomainFailCountBook_Grade1[DomainName]++;

                                if (!EachGradeStudentAccumulatedFailCount_Grade1.ContainsKey(sr.ID))
                                {
                                    EachGradeStudentAccumulatedFailCount_Grade1.Add(sr.ID, 0);

                                    EachGradeStudentAccumulatedFailCount_Grade1[sr.ID]++;
                                }
                                else
                                {
                                    EachGradeStudentAccumulatedFailCount_Grade1[sr.ID]++;                                
                                }

                            }
                            //二年級、八年級
                            if (!pass && (sr.Class.GradeYear == 2 || sr.Class.GradeYear == 8))
                            {
                                TotalDomainFailCountBook_Grade2[DomainName]++;

                                if (!EachGradeStudentAccumulatedFailCount_Grade2.ContainsKey(sr.ID))
                                {
                                    EachGradeStudentAccumulatedFailCount_Grade2.Add(sr.ID, 0);

                                    EachGradeStudentAccumulatedFailCount_Grade2[sr.ID]++;
                                }
                                else
                                {
                                    EachGradeStudentAccumulatedFailCount_Grade2[sr.ID]++;
                                }

                            }
                            //三年級、九年級
                            if (!pass && (sr.Class.GradeYear == 3 || sr.Class.GradeYear == 9))
                            {
                                TotalDomainFailCountBook_Grade3[DomainName]++;

                                if (!EachGradeStudentAccumulatedFailCount_Grade3.ContainsKey(sr.ID))
                                {
                                    EachGradeStudentAccumulatedFailCount_Grade3.Add(sr.ID, 0);

                                    EachGradeStudentAccumulatedFailCount_Grade3[sr.ID]++;
                                }
                                else
                                {
                                    EachGradeStudentAccumulatedFailCount_Grade3[sr.ID]++;
                                }

                            }

                        }
                    }

                }

                if (pass_makeup) 
                {
                    if (sr.Class != null) 
                    {
                        if (sr.Class.GradeYear == 1 || sr.Class.GradeYear == 7) 
                        {
                            EachGradeMakeUpStudentCount_sucessful["Grade1"]++;                                                                        
                        }
                        if (sr.Class.GradeYear == 2 || sr.Class.GradeYear == 8)
                        {
                            EachGradeMakeUpStudentCount_sucessful["Grade2"]++;
                        }
                        if (sr.Class.GradeYear == 3 || sr.Class.GradeYear == 9)
                        {
                            EachGradeMakeUpStudentCount_sucessful["Grade3"]++;
                        }                                        
                    }                                
                }

                //用來數各年級的學生數量
                if (sr.Class != null)
                {
                    if (sr.Class.GradeYear == 1 || sr.Class.GradeYear == 7)
                    {
                        EachGradeStudentCount["Grade1"]++;
                    }

                    if (sr.Class.GradeYear == 2 || sr.Class.GradeYear == 8)
                    {
                        EachGradeStudentCount["Grade2"]++;
                    }
                    if (sr.Class.GradeYear == 3 || sr.Class.GradeYear == 9)
                    {
                        EachGradeStudentCount["Grade3"]++;
                    }
                }

            }

        }


        public bool EachJSRCalculator(JHSemesterScoreRecord JSR, JHSemesterScoreRecord JSR2, String DomainName)
        {

            JHSemesterScoreRecord jsr1 = JSR;
            JHSemesterScoreRecord jsr2 = JSR2;

            // 領域總分，領域總加權數
            decimal scoreSum = 0, creditCount = 0;

            if (jsr1.Domains.ContainsKey(DomainName))
            {
                scoreSum += (decimal)jsr1.Domains[DomainName].Score * (decimal)jsr1.Domains[DomainName].Credit;
                creditCount += (decimal)jsr1.Domains[DomainName].Credit;
            }

            if (jsr2.Domains.ContainsKey(DomainName))
            {
                scoreSum += (decimal)jsr2.Domains[DomainName].Score * (decimal)jsr2.Domains[DomainName].Credit;
                creditCount += (decimal)jsr2.Domains[DomainName].Credit;
            }

            bool pass = true;

            if (creditCount == 0)
                return true;

            // 學年領域成績 ，因為國中 目前(2017/1/9) 並沒有 "學年領域成績" 的概念， 所以要在這邊手動算
            decimal SchoolYearDomainScore = Math.Round((scoreSum) / (creditCount), 2, MidpointRounding.AwayFromZero);

            if (SchoolYearDomainScore < 60)
            {
                pass = false;

                String FailedDomainBookKey = jsr1.Student.ID + "_" + jsr1.SchoolYear + "_" + DomainName;

                // 此物件用來記錄該學生該領域不及格詳情(穎驊在本程式自訂)
                StudentDomainFailRecord SDFR = new StudentDomainFailRecord();

                SDFR.Grade = jsr1.Student.Class.GradeYear;
                SDFR.Class = jsr1.Student.Class.Name;
                SDFR.Student = jsr1.Student.Name;
                SDFR.SchoolYear = jsr1.SchoolYear;
                SDFR.FailedDomainName = DomainName;

                if (jsr1.Domains.ContainsKey(DomainName))
                {
                    SDFR._first_domain_score = jsr1.Domains[DomainName].Score;
                    SDFR._first_domain_credit = jsr1.Domains[DomainName].Credit;
                }
                else
                {
                    // 恩正說，備註暫時先不要用
                    //SDFR._memo = "學生在此項目的學期領域成績、權數資料不全，將導致系統計算錯誤，請至系統檢查核對";
                }

                if (jsr2.Domains.ContainsKey(DomainName))
                {
                    SDFR._second_domain_score = jsr2.Domains[DomainName].Score;
                    SDFR._second_domain_credit = jsr2.Domains[DomainName].Credit;
                }
                else
                {
                    // 恩正說，備註暫時先不要用
                    //SDFR._memo = "學生在此項目的學期領域成績、權數資料不全，將導致系統計算錯誤，請至系統檢查核對";
                }

                SDFR._school_year_domain_score = SchoolYearDomainScore;

                StudentFailedDomainRecordList.Add(FailedDomainBookKey, SDFR);

            }
            else
            {
                pass = true;
            }


            // 此物件用來記錄  所有有的補考紀錄
            StudentMakeUpScoreRecord SMUSR = new StudentMakeUpScoreRecord();
            
            String MakeUpDomainBookKey = jsr1.Student.ID + "_" + jsr1.SchoolYear + "_" + DomainName;

            bool first_sem_has_makeup = false;
            bool second_sem_has_makeup = false;

            // 假如第一學期有補考成績
            if (jsr1.Domains.ContainsKey(DomainName))
            {
                if ( jsr1.Domains[DomainName].ScoreMakeup.HasValue)
                {                                       

                    SMUSR.Grade = jsr1.Student.Class.GradeYear;
                    SMUSR.Class = jsr1.Student.Class.Name;
                    SMUSR.Student = jsr1.Student.Name;
                    SMUSR.SchoolYear = jsr1.SchoolYear;
                    SMUSR.MakeUpDomainName = DomainName;     
                    SMUSR._first_domain_score = jsr1.Domains[DomainName].Score;
                    SMUSR._first_domain_origin_score = jsr1.Domains[DomainName].ScoreOrigin;
                    SMUSR._first_domain_makeup_score = jsr1.Domains[DomainName].ScoreMakeup;
                    SMUSR._first_domain_makeup_credit = jsr1.Domains[DomainName].Credit;                                                         
                    SMUSR._school_year_domain_score = SchoolYearDomainScore;

                    first_sem_has_makeup = true; 
                }                
            }

            // 假如第二學期有補考成績
            if (jsr2.Domains.ContainsKey(DomainName))
            {
                if (jsr2.Domains[DomainName].ScoreMakeup.HasValue)
                {
                    SMUSR.Grade = jsr1.Student.Class.GradeYear;
                    SMUSR.Class = jsr1.Student.Class.Name;
                    SMUSR.Student = jsr1.Student.Name;
                    SMUSR.SchoolYear = jsr1.SchoolYear;
                    SMUSR.MakeUpDomainName = DomainName;            
                    SMUSR._second_domain_score = jsr2.Domains[DomainName].Score;
                    SMUSR._second_domain_origin_score = jsr2.Domains[DomainName].ScoreOrigin;
                    SMUSR._second_domain_makeup_score = jsr2.Domains[DomainName].ScoreMakeup;
                    SMUSR._second_domain_makeup_credit = jsr2.Domains[DomainName].Credit;                    
                    SMUSR._school_year_domain_score = SchoolYearDomainScore;

                    second_sem_has_makeup = true;
                }

                
            }
            if (first_sem_has_makeup || second_sem_has_makeup)
            {
                //加入 有補考紀錄學生 明細
                StudentMakeUpDomainRecordList.Add(MakeUpDomainBookKey, SMUSR);            
            }

            return pass;

        }


        // 計算   預警、輔導及補救措施成效 (在本ischool 系統 只有補考)
        public bool EachJSRCalculator_For_MakeUp(JHSemesterScoreRecord JSR, JHSemesterScoreRecord JSR2, String DomainName)
        {

            JHSemesterScoreRecord jsr1 = JSR;
            JHSemesterScoreRecord jsr2 = JSR2;

            // 領域總分，領域總加權數
            decimal scoreSum = 0, creditCount = 0;

            if (jsr1.Domains.ContainsKey(DomainName))
            {
                scoreSum += (decimal)jsr1.Domains[DomainName].Score * (decimal)jsr1.Domains[DomainName].Credit;
                creditCount += (decimal)jsr1.Domains[DomainName].Credit;
            }

            if (jsr2.Domains.ContainsKey(DomainName))
            {
                scoreSum += (decimal)jsr2.Domains[DomainName].Score * (decimal)jsr2.Domains[DomainName].Credit;
                creditCount += (decimal)jsr2.Domains[DomainName].Credit;
            }

            bool pass = true;

            bool have_makeup_score = false;

            // 假如第一學期有補考成績
            if (jsr1.Domains.ContainsKey(DomainName))
            {
                if (jsr1.Domains[DomainName].ScoreMakeup.HasValue)
                {
                    have_makeup_score = true;

                    // 穎驊註解，會這樣處理，是因為 關於補考紀錄， 局端 要數的是"人頭" 而不是"人次"，
                    // 也就是說，如果同一學生 上有五筆不同領域的補考紀錄， 其中三筆是補考後 全學年領域成績及格，
                    // 這樣 系統 會算 :
                    // 1.該年級已啟動預警、輔導及補救措施之人數+1
                    // 2.經由預警、輔導及補救措施而達及格標準之人數 +1
                    // 僅會 加1人頭

                    if (!Has_MakeUpStudentIDList.Contains(jsr1.Student.ID))
                    {
                        Has_MakeUpStudentIDList.Add(jsr1.Student.ID);

                        if (jsr1.Student.Class != null)
                        {
                            if (jsr1.Student.Class.GradeYear == 1)
                            {
                                EachGradeMakeUpStudentCount["Grade1"]++;

                            }
                            if (jsr1.Student.Class.GradeYear == 2)
                            {
                                EachGradeMakeUpStudentCount["Grade2"]++;

                            }
                            if (jsr1.Student.Class.GradeYear == 3)
                            {
                                EachGradeMakeUpStudentCount["Grade3"]++;
                            }
                        }
                    }
                    else
                    {
                    }
                }
            }

            // 假如第二學期有補考成績
            if (jsr2.Domains.ContainsKey(DomainName))
            {
                if (jsr2.Domains[DomainName].ScoreMakeup.HasValue)
                {
                    have_makeup_score = true;

                    // 穎驊註解，會這樣處理，是因為 關於補考紀錄， 局端 要數的是"人頭" 而不是"人次"，
                    // 也就是說，如果同一學生 上有五筆不同領域的補考紀錄， 其中三筆是補考後 全學年領域成績及格，
                    // 這樣 系統 會算 :
                    // 1.該年級已啟動預警、輔導及補救措施之人數+1
                    // 2.經由預警、輔導及補救措施而達及格標準之人數 +1
                    // 僅會 加1人頭

                    if (!Has_MakeUpStudentIDList.Contains(jsr1.Student.ID))
                    {
                        Has_MakeUpStudentIDList.Add(jsr1.Student.ID);

                        if (jsr1.Student.Class != null)
                        {
                            if (jsr1.Student.Class.GradeYear == 1)
                            {
                                EachGradeMakeUpStudentCount["Grade1"]++;

                            }
                            if (jsr1.Student.Class.GradeYear == 2)
                            {
                                EachGradeMakeUpStudentCount["Grade2"]++;

                            }
                            if (jsr1.Student.Class.GradeYear == 3)
                            {
                                EachGradeMakeUpStudentCount["Grade3"]++;
                            }
                        }
                    }
                    else
                    {
                    }
                }
            }

            decimal SchoolYearDomainScore = 0;

            // 假如有補考成績 、但是總加權學分為0  當成沒有過(因為算出來的 學年領域成績 =0 <60)。
            if (creditCount == 0 && have_makeup_score) 
            {
                return false;
            }

            // 只要加權總學分 不為0，就算學年領域成績
            if (creditCount != 0 )
            {
                SchoolYearDomainScore = Math.Round((scoreSum) / (creditCount), 2, MidpointRounding.AwayFromZero);
            }
            
            // 若有補考成績、 學年領域成績>  60 ，就是 補考及格成功 ， 反之 若有補考成績 學年成績卻沒有 > 60 、或是 沒有補考成績 學年成績 本來就> 60 ， 其都不代表 補考及格成功。
            if (SchoolYearDomainScore > 60  && have_makeup_score )
            {
                pass = true;
            }
            else
            {
                pass = false;
            }

            return pass;
        }


        public void RunWork()
        {
            _BW.DoWork += new DoWorkEventHandler(_BW_DoWork);
            _BW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_BW_RunWorkerCompleted);
            _BW.RunWorkerAsync();
        }


        // 點選"取消" 後關閉視窗
        private void btn_Exit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // 點選"開始計算" 後開始進行運算
        private void btn_StartCalculate_Click(object sender, EventArgs e)
        {
            if (_BW.IsBusy)
            {
                MessageBox.Show("系統忙碌,請稍後再試...");
            }
            else
            {
                RunWork();
                // 要記得在功能Run完後Close()，否則使用者若是再按下殘留在主畫面的計算按鈕，會造成重覆使用一樣沒清乾淨的變數，造成當機跳出
                this.Close();
            }
        }

    }
}
