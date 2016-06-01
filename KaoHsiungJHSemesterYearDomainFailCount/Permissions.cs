using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KaoHsiungJHSemesterYearDomainFailCount
{
    class Permissions
    {

        public static string 全年級學年領域統計不及格人數報表 { get { return "KaoHsiungJHSemesterYearDomainFailCount"; } }
        public static bool 年級學年領域統計不及格人數報表權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[全年級學年領域統計不及格人數報表].Executable;
            }
        }
    }
}
