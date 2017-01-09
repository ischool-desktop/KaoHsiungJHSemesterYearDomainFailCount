using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KaoHsiungJHSemesterYearDomainFailCount
{
    class StudentMakeUpScoreRecord
    {
        // 摘要: 
        //     有補考者成績資訊欄位 (穎驊自訂)

        public StudentMakeUpScoreRecord()
        { 
        }
      
        ///<summary>年級</summary>
        public int? Grade { get; set; }
                
        //班級， 這邊超級蠢，沒辦法直接打class，系統會認定是單字XD

        ///<summary>班級</summary>
        public string Class { get; set; }

        ///<summary>學生</summary>
        public string Student { get; set; }

        ///<summary>學年</summary>
        public int? SchoolYear { get; set; }

        ///<summary>補考領域名稱</summary>
        public string MakeUpDomainName { get; set; }

        ///<summary>補考領域第一學期成績</summary>
        public decimal? _first_domain_score { get; set; }

        ///<summary>補考領域第一學期原始成績</summary>
        public decimal? _first_domain_origin_score { get; set; }

        ///<summary>補考領域第一學期補考成績</summary>
        public decimal? _first_domain_makeup_score { get; set; }

       
        ///<summary>補考領域第一學期權數</summary>
        public decimal? _first_domain_makeup_credit { get; set; }

        ///<summary>補考領域第二學期成績</summary>
        public decimal? _second_domain_score { get; set; }

        ///<summary>補考領域第二學期原始成績</summary>
        public decimal? _second_domain_origin_score { get; set; }

        ///<summary>補考領域第二學期補考成績</summary>
        public decimal? _second_domain_makeup_score { get; set; }

        
        ///<summary>補考領域第二學期權數</summary>
        public decimal? _second_domain_makeup_credit { get; set; }

        
        ///<summary>計算後全學年領域成績</summary>
        public decimal? _school_year_domain_score { get; set; }

        
        ///<summary> 備註，用來記錄例外狀況</summary>
        public String _memo { get; set; }






    }
}
