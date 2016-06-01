using System;
using System.Collections.Generic;
using System.Text;
using FISCA;
using FISCA.Presentation;
using K12.Presentation;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using FISCA.Permission;


namespace KaoHsiungJHSemesterYearDomainFailCount
{
    public static class Program
    {
        internal static ModuleMode Mode { get; private set; }

        [MainMethod()]
        public static void Main()
        {


            MenuButton item = JHSchool.Affair.EduAdmin.Instance.RibbonBarItems["批次作業/檢視"]["成績作業"];
            item["全年級學年領域統計不及格人數報表"].Enable = Permissions.年級學年領域統計不及格人數報表權限;
            item["全年級學年領域統計不及格人數報表"].Click += delegate
            {

                
                Mention_To_Start MTS = new Mention_To_Start();

                MTS.ShowDialog();
                
            };
            Catalog detail1 = RoleAclSource.Instance["教務作業"];
            detail1.Add(new RibbonFeature(Permissions.全年級學年領域統計不及格人數報表, "全年級學年領域統計不及格人數報表"));

                     
}

    }

    internal enum ModuleMode
    {
        /// <summary>
        /// 新竹
        /// </summary>
        HsinChu,
        /// <summary>
        /// 高雄
        /// </summary>
        KaoHsiung
    }
}
