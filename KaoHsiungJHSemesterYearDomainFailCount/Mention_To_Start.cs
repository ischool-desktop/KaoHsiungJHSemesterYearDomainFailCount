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

            //取得該學校當學期 (本程式目前未用到)
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


            EachGradeStudentCount.Add("Grade1", Grade1StudentNumber);
            EachGradeStudentCount.Add("Grade2", Grade2StudentNumber);
            EachGradeStudentCount.Add("Grade3", Grade3StudentNumber);

        }

        private void _BW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("報表列印完成");
            Workbook wb = e.Result as Workbook;

            if (wb == null)
                return;

            SaveFileDialog save = new SaveFileDialog();
            save.Title = "另存新檔";
            save.FileName = "全校學期領域不及格人數.xls";
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
            Workbook wb = new Workbook(new MemoryStream(Properties.Resources.計算全校學期領域不及格人數));


            //wb.Worksheets.AddCopy(0);

            wb.Worksheets[0].Name = "全校全年級各領域不及格人數統計表";
            wb.Worksheets[1].Name = "學年領域不及格學生明細";
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
            //wb.Worksheets.RemoveAt(0);

            wb.Worksheets.ActiveSheetIndex = 0;

            e.Result = wb;
        }




        public void GetStudentDomainScore()
        {


            // 處理學生的每一筆成績，處理好後，分上下學期，使用特定的KEY值，分別加入上下學期成績的Dictionary。
            foreach (JHSemesterScoreRecord jsr in JHSchool.Data.JHSemesterScore.SelectByStudentIDs(StudentIDs))
            {
                // 該比學期成績資料，學年必須等於當下學期，且學生的目前狀態是一般生(Status == 0)
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

                            //不及格就進行數數
                            if (!pass && sr.Class.GradeYear == 1)
                            {

                                TotalDomainFailCountBook_Grade1[DomainName]++;

                            }
                            if (!pass && sr.Class.GradeYear == 2)
                            {

                                TotalDomainFailCountBook_Grade2[DomainName]++;

                            }
                            if (!pass && sr.Class.GradeYear == 3)
                            {

                                TotalDomainFailCountBook_Grade3[DomainName]++;

                            }
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
