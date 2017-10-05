using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KaoHsiungJHSemesterYearDomainFailCount
{
    class SqlString
    {

        public static string sql(string semester)
        {


            string sql =
            @"WITH 年級資料 AS (
	SELECT id, school_year, MAX(to_number( NULLIF(grade_year, ''), '999999999999.9999')) as grade_year
	FROM (
		SELECT id, 
			" + semester + @"  as school_year, 
			1 as semester,
			xpath_string('<root>'|| sems_history ||'</root>','/root/History[ @SchoolYear=''" + semester + @"'' and @Semester=''1'' ]/@GradeYear') as grade_year
		FROM student
		UNION
		SELECT id, 
			" + semester + @" as school_year, 
			2 as semester,
			xpath_string('<root>'|| sems_history ||'</root>','/root/History[ @SchoolYear=''" + semester + @"'' and @Semester=''2'' ]/@GradeYear') as grade_year
		FROM student
	) as grade_year_mapping
	GROUP BY id, school_year
	HAVING MAX(to_number( NULLIF(grade_year, ''), '999999999999.9999')) is not null
), 成績資料 AS (
	SELECT b.ref_student_id, b.school_year, 年級資料.grade_year, b.領域, b.領域成績, CASE WHEN b.領域成績 >= 60  THEN true ELSE false END as pass
	FROM (
		SELECT ref_student_id, school_year, 領域, AVG(to_number( NULLIF(領域成績, ''), '999999999999.9999')) as 領域成績
		FROM (
			SELECT ref_student_id, school_year, semester,
				unnest(xpath('/R/Domains/Domain/@成績', ('<R>'||sems_subj_score.score_info||'</R>')::xml))::text as 領域成績,
				unnest(xpath('/R/Domains/Domain/@領域', ('<R>'||sems_subj_score.score_info||'</R>')::xml))::text as 領域
			FROM sems_subj_score 
			WHERE school_year in (" + semester + @")
		) as a
		GROUP BY ref_student_id, school_year, 領域
	) as b
		LEFT OUTER JOIN 年級資料 on 年級資料.id = b.ref_student_id AND 年級資料.school_year = b.school_year
	WHERE b.領域成績 is not null
		AND b.領域 in (
			'語文',
			'數學',
			'社會',
			'自然與生活科技',
			'健康與體育',
			'藝術與人文',
			'綜合活動'
		)
	ORDER BY school_year, grade_year, ref_student_id, 領域
), 年級人數統計 AS (
	SELECT school_year, grade_year, COUNT(*)
	FROM 年級資料
	GROUP BY school_year, grade_year
), 個人不及格領域數統計 AS (
	SELECT ref_student_id, school_year, grade_year, SUM(CASE WHEN pass THEN 0 ELSE 1 END) AS 不及格領域數
	FROM 成績資料
	GROUP BY ref_student_id, school_year, grade_year
), 補考資料 AS (
	SELECT 年級資料.grade_year, COUNT(*) AS 補考人數, SUM(CASE WHEN 補考成功 THEN 1 ELSE 0 END) AS 補考通過人數
	FROM (
		SELECT ref_student_id, school_year, ( CASE WHEN MAX(to_number( NULLIF(補考成績, ''), '999999999999.9999')) >= 60 THEN true ELSE false END ) as 補考成功
		FROM (
			SELECT ref_student_id, school_year, 
				unnest(xpath('/R/Domains/Domain/@補考成績', ('<R>'||sems_subj_score.score_info||'</R>')::xml))::text as 補考成績,
				unnest(xpath('/R/Domains/Domain/@成績', ('<R>'||sems_subj_score.score_info||'</R>')::xml))::text as 領域成績,
				unnest(xpath('/R/Domains/Domain/@領域', ('<R>'||sems_subj_score.score_info||'</R>')::xml))::text as 領域
			FROM sems_subj_score 
			WHERE school_year in (" + semester + @")
			UNION ALL 
			SELECT ref_student_id, school_year,  
				unnest(xpath('/R/SemesterSubjectScoreInfo/Subject/@補考成績', ('<R>'||sems_subj_score.score_info||'</R>')::xml))::text  as 補考成績,
				'0'  as 領域成績,
				unnest(xpath('/R/SemesterSubjectScoreInfo/Subject/@領域', ('<R>'||sems_subj_score.score_info||'</R>')::xml))::text  as 領域
			FROM sems_subj_score 
			WHERE school_year in (" + semester + @")
		) as a
		WHERE to_number( NULLIF(補考成績, ''), '999999999999.9999') is not null
			AND a.領域 in (
				'語文',
				'數學',
				'社會',
				'自然與生活科技',
				'健康與體育',
				'藝術與人文',
				'綜合活動'
			)
		GROUP BY ref_student_id, school_year
	) as b
		LEFT OUTER JOIN 年級資料 on 年級資料.id = b.ref_student_id AND 年級資料.school_year = b.school_year
	GROUP BY 年級資料.grade_year
)
SELECT 
	年級人數統計.school_year AS 學年度,
	年級人數統計.grade_year % 6 AS 年級,
	年級人數統計.count AS 年級人數,
	語文不及格人數統計.sum AS 語文不及格人數,
	數學不及格人數統計.sum AS 數學不及格人數,
	社會不及格人數統計.sum AS 社會不及格人數,
	自然與生活科技不及格人數統計.sum AS 自然與生活科技不及格人數,
	健康與體育不及格人數統計.sum AS 健康與體育不及格人數,
	藝術與人文不及格人數統計.sum AS 藝術與人文不及格人數,
	綜合活動不及格人數統計.sum AS 綜合活動不及格人數,
	
	不及格領域數1.count AS 不及格領域數1人數,
	不及格領域數2.count AS 不及格領域數2人數,
	不及格領域數3.count AS 不及格領域數3人數,
	不及格領域數4.count AS 不及格領域數4人數,
	不及格領域數5.count AS 不及格領域數5人數,
	不及格領域數6.count AS 不及格領域數6人數,
	不及格領域數7.count AS 不及格領域數7人數,
	補考資料.補考人數,
	補考資料.補考通過人數
	
	--CAST( 語文不及格人數統計.sum AS float) / 年級人數統計.count AS 語文不及格人數比例,
	--CAST( 數學不及格人數統計.sum AS float) / 年級人數統計.count AS 數學不及格人數比例,
	--CAST( 社會不及格人數統計.sum AS float) / 年級人數統計.count AS 社會不及格人數比例,
	--CAST( 自然與生活科技不及格人數統計.sum AS float) / 年級人數統計.count AS 自然與生活科技不及格人數比例,
	--CAST( 健康與體育不及格人數統計.sum AS float) / 年級人數統計.count AS 健康與體育不及格人數比例,
	--CAST( 藝術與人文不及格人數統計.sum AS float) / 年級人數統計.count AS 藝術與人文不及格人數比例,
	--CAST( 綜合活動不及格人數統計.sum AS float) / 年級人數統計.count AS 綜合活動不及格人數比例
FROM
	年級人數統計 
	LEFT OUTER JOIN (
		SELECT school_year, grade_year, SUM(CASE WHEN pass THEN 0 ELSE 1 END) AS sum
		FROM 成績資料
		WHERE 領域 = '語文'
		GROUP BY school_year, grade_year
	) AS 語文不及格人數統計
		ON 語文不及格人數統計.school_year = 年級人數統計.school_year
		AND 語文不及格人數統計.grade_year = 年級人數統計.grade_year
	LEFT OUTER JOIN (
		SELECT school_year, grade_year, SUM(CASE WHEN pass THEN 0 ELSE 1 END) AS sum
		FROM 成績資料
		WHERE 領域 = '數學'
		GROUP BY school_year, grade_year
	) AS 數學不及格人數統計
		ON 數學不及格人數統計.school_year = 年級人數統計.school_year
		AND 數學不及格人數統計.grade_year = 年級人數統計.grade_year
	LEFT OUTER JOIN (
		SELECT school_year, grade_year, SUM(CASE WHEN pass THEN 0 ELSE 1 END) AS sum
		FROM 成績資料
		WHERE 領域 = '社會'
		GROUP BY school_year, grade_year
	) AS 社會不及格人數統計
		ON 社會不及格人數統計.school_year = 年級人數統計.school_year
		AND 社會不及格人數統計.grade_year = 年級人數統計.grade_year
	LEFT OUTER JOIN (
		SELECT school_year, grade_year, SUM(CASE WHEN pass THEN 0 ELSE 1 END) AS sum
		FROM 成績資料
		WHERE 領域 = '自然與生活科技'
		GROUP BY school_year, grade_year
	) AS 自然與生活科技不及格人數統計
		ON 自然與生活科技不及格人數統計.school_year = 年級人數統計.school_year
		AND 自然與生活科技不及格人數統計.grade_year = 年級人數統計.grade_year
	LEFT OUTER JOIN (
		SELECT school_year, grade_year, SUM(CASE WHEN pass THEN 0 ELSE 1 END) AS sum
		FROM 成績資料
		WHERE 領域 = '健康與體育'
		GROUP BY school_year, grade_year
	) AS 健康與體育不及格人數統計
		ON 健康與體育不及格人數統計.school_year = 年級人數統計.school_year
		AND 健康與體育不及格人數統計.grade_year = 年級人數統計.grade_year
	LEFT OUTER JOIN (
		SELECT school_year, grade_year, SUM(CASE WHEN pass THEN 0 ELSE 1 END) AS sum
		FROM 成績資料
		WHERE 領域 = '藝術與人文'
		GROUP BY school_year, grade_year
	) AS 藝術與人文不及格人數統計
		ON 藝術與人文不及格人數統計.school_year = 年級人數統計.school_year
		AND 藝術與人文不及格人數統計.grade_year = 年級人數統計.grade_year
	LEFT OUTER JOIN (
		SELECT school_year, grade_year, SUM(CASE WHEN pass THEN 0 ELSE 1 END) AS sum
		FROM 成績資料
		WHERE 領域 = '綜合活動'
		GROUP BY school_year, grade_year
	) AS 綜合活動不及格人數統計
		ON 綜合活動不及格人數統計.school_year = 年級人數統計.school_year
		AND 綜合活動不及格人數統計.grade_year = 年級人數統計.grade_year
	LEFT OUTER JOIN (
		SELECT school_year, grade_year, count(*)
		FROM 個人不及格領域數統計
		WHERE 不及格領域數 = '1'
		GROUP BY school_year, grade_year
	) AS 不及格領域數1
		ON 不及格領域數1.school_year = 年級人數統計.school_year
		AND 不及格領域數1.grade_year = 年級人數統計.grade_year
	LEFT OUTER JOIN (
		SELECT school_year, grade_year, count(*)
		FROM 個人不及格領域數統計
		WHERE 不及格領域數 = '2'
		GROUP BY school_year, grade_year
	) AS 不及格領域數2
		ON 不及格領域數2.school_year = 年級人數統計.school_year
		AND 不及格領域數2.grade_year = 年級人數統計.grade_year
	LEFT OUTER JOIN (
		SELECT school_year, grade_year, count(*)
		FROM 個人不及格領域數統計
		WHERE 不及格領域數 = '3'
		GROUP BY school_year, grade_year
	) AS 不及格領域數3
		ON 不及格領域數3.school_year = 年級人數統計.school_year
		AND 不及格領域數3.grade_year = 年級人數統計.grade_year
	LEFT OUTER JOIN (
		SELECT school_year, grade_year, count(*)
		FROM 個人不及格領域數統計
		WHERE 不及格領域數 = '4'
		GROUP BY school_year, grade_year
	) AS 不及格領域數4
		ON 不及格領域數4.school_year = 年級人數統計.school_year
		AND 不及格領域數4.grade_year = 年級人數統計.grade_year
	LEFT OUTER JOIN (
		SELECT school_year, grade_year, count(*)
		FROM 個人不及格領域數統計
		WHERE 不及格領域數 = '5'
		GROUP BY school_year, grade_year
	) AS 不及格領域數5
		ON 不及格領域數5.school_year = 年級人數統計.school_year
		AND 不及格領域數5.grade_year = 年級人數統計.grade_year
	LEFT OUTER JOIN (
		SELECT school_year, grade_year, count(*)
		FROM 個人不及格領域數統計
		WHERE 不及格領域數 = '6'
		GROUP BY school_year, grade_year
	) AS 不及格領域數6
		ON 不及格領域數6.school_year = 年級人數統計.school_year
		AND 不及格領域數6.grade_year = 年級人數統計.grade_year
	LEFT OUTER JOIN (
		SELECT school_year, grade_year, count(*)
		FROM 個人不及格領域數統計
		WHERE 不及格領域數 = '7'
		GROUP BY school_year, grade_year
	) AS 不及格領域數7
		ON 不及格領域數7.school_year = 年級人數統計.school_year
		AND 不及格領域數7.grade_year = 年級人數統計.grade_year
	LEFT OUTER JOIN 補考資料
		ON 補考資料.grade_year = 年級人數統計.grade_year
ORDER BY 年級人數統計.grade_year";


            return sql;

        }



//        public string sql =
//            @"WITH 年級資料 AS (
//	SELECT id, school_year, MAX(to_number( NULLIF(grade_year, ''), '999999999999.9999')) as grade_year
//	FROM (
//		SELECT id, 
//			105 as school_year, 
//			1 as semester,
//			xpath_string('<root>'|| sems_history ||'</root>','/root/History[ @SchoolYear=''105'' and @Semester=''1'' ]/@GradeYear') as grade_year
//		FROM student
//		UNION
//		SELECT id, 
//			105 as school_year, 
//			2 as semester,
//			xpath_string('<root>'|| sems_history ||'</root>','/root/History[ @SchoolYear=''105'' and @Semester=''2'' ]/@GradeYear') as grade_year
//		FROM student
//	) as grade_year_mapping
//	GROUP BY id, school_year
//	HAVING MAX(to_number( NULLIF(grade_year, ''), '999999999999.9999')) is not null
//), 成績資料 AS (
//	SELECT b.ref_student_id, b.school_year, 年級資料.grade_year, b.領域, b.領域成績, CASE WHEN b.領域成績 >= 60  THEN true ELSE false END as pass
//	FROM (
//		SELECT ref_student_id, school_year, 領域, AVG(to_number( NULLIF(領域成績, ''), '999999999999.9999')) as 領域成績
//		FROM (
//			SELECT ref_student_id, school_year, semester,
//				unnest(xpath('/R/Domains/Domain/@成績', ('<R>'||sems_subj_score.score_info||'</R>')::xml))::text as 領域成績,
//				unnest(xpath('/R/Domains/Domain/@領域', ('<R>'||sems_subj_score.score_info||'</R>')::xml))::text as 領域
//			FROM sems_subj_score 
//			WHERE school_year in (105)
//		) as a
//		GROUP BY ref_student_id, school_year, 領域
//	) as b
//		LEFT OUTER JOIN 年級資料 on 年級資料.id = b.ref_student_id AND 年級資料.school_year = b.school_year
//	WHERE b.領域成績 is not null
//		AND b.領域 in (
//			'語文',
//			'數學',
//			'社會',
//			'自然與生活科技',
//			'健康與體育',
//			'藝術與人文',
//			'綜合活動'
//		)
//	ORDER BY school_year, grade_year, ref_student_id, 領域
//), 年級人數統計 AS (
//	SELECT school_year, grade_year, COUNT(*)
//	FROM 年級資料
//	GROUP BY school_year, grade_year
//), 個人不及格領域數統計 AS (
//	SELECT ref_student_id, school_year, grade_year, SUM(CASE WHEN pass THEN 0 ELSE 1 END) AS 不及格領域數
//	FROM 成績資料
//	GROUP BY ref_student_id, school_year, grade_year
//), 補考資料 AS (
//	SELECT 年級資料.grade_year, COUNT(*) AS 補考人數, SUM(CASE WHEN 補考成功 THEN 1 ELSE 0 END) AS 補考通過人數
//	FROM (
//		SELECT ref_student_id, school_year, ( CASE WHEN MAX(to_number( NULLIF(補考成績, ''), '999999999999.9999')) >= 60 THEN true ELSE false END ) as 補考成功
//		FROM (
//			SELECT ref_student_id, school_year, 
//				unnest(xpath('/R/Domains/Domain/@補考成績', ('<R>'||sems_subj_score.score_info||'</R>')::xml))::text as 補考成績,
//				unnest(xpath('/R/Domains/Domain/@成績', ('<R>'||sems_subj_score.score_info||'</R>')::xml))::text as 領域成績,
//				unnest(xpath('/R/Domains/Domain/@領域', ('<R>'||sems_subj_score.score_info||'</R>')::xml))::text as 領域
//			FROM sems_subj_score 
//			WHERE school_year in (105)
//			UNION ALL 
//			SELECT ref_student_id, school_year,  
//				unnest(xpath('/R/SemesterSubjectScoreInfo/Subject/@補考成績', ('<R>'||sems_subj_score.score_info||'</R>')::xml))::text  as 補考成績,
//				'0'  as 領域成績,
//				unnest(xpath('/R/SemesterSubjectScoreInfo/Subject/@領域', ('<R>'||sems_subj_score.score_info||'</R>')::xml))::text  as 領域
//			FROM sems_subj_score 
//			WHERE school_year in (105)
//		) as a
//		WHERE to_number( NULLIF(補考成績, ''), '999999999999.9999') is not null
//			AND a.領域 in (
//				'語文',
//				'數學',
//				'社會',
//				'自然與生活科技',
//				'健康與體育',
//				'藝術與人文',
//				'綜合活動'
//			)
//		GROUP BY ref_student_id, school_year
//	) as b
//		LEFT OUTER JOIN 年級資料 on 年級資料.id = b.ref_student_id AND 年級資料.school_year = b.school_year
//	GROUP BY 年級資料.grade_year
//)
//SELECT 
//	年級人數統計.school_year AS 學年度,
//	年級人數統計.grade_year % 6 AS 年級,
//	年級人數統計.count AS 年級人數,
//	語文不及格人數統計.sum AS 語文不及格人數,
//	數學不及格人數統計.sum AS 數學不及格人數,
//	社會不及格人數統計.sum AS 社會不及格人數,
//	自然與生活科技不及格人數統計.sum AS 自然與生活科技不及格人數,
//	健康與體育不及格人數統計.sum AS 健康與體育不及格人數,
//	藝術與人文不及格人數統計.sum AS 藝術與人文不及格人數,
//	綜合活動不及格人數統計.sum AS 綜合活動不及格人數,
	
//	不及格領域數1.count AS 不及格領域數1人數,
//	不及格領域數2.count AS 不及格領域數2人數,
//	不及格領域數3.count AS 不及格領域數3人數,
//	不及格領域數4.count AS 不及格領域數4人數,
//	不及格領域數5.count AS 不及格領域數5人數,
//	不及格領域數6.count AS 不及格領域數6人數,
//	不及格領域數7.count AS 不及格領域數7人數,
//	補考資料.補考人數,
//	補考資料.補考通過人數
	
//	--CAST( 語文不及格人數統計.sum AS float) / 年級人數統計.count AS 語文不及格人數比例,
//	--CAST( 數學不及格人數統計.sum AS float) / 年級人數統計.count AS 數學不及格人數比例,
//	--CAST( 社會不及格人數統計.sum AS float) / 年級人數統計.count AS 社會不及格人數比例,
//	--CAST( 自然與生活科技不及格人數統計.sum AS float) / 年級人數統計.count AS 自然與生活科技不及格人數比例,
//	--CAST( 健康與體育不及格人數統計.sum AS float) / 年級人數統計.count AS 健康與體育不及格人數比例,
//	--CAST( 藝術與人文不及格人數統計.sum AS float) / 年級人數統計.count AS 藝術與人文不及格人數比例,
//	--CAST( 綜合活動不及格人數統計.sum AS float) / 年級人數統計.count AS 綜合活動不及格人數比例
//FROM
//	年級人數統計 
//	LEFT OUTER JOIN (
//		SELECT school_year, grade_year, SUM(CASE WHEN pass THEN 0 ELSE 1 END) AS sum
//		FROM 成績資料
//		WHERE 領域 = '語文'
//		GROUP BY school_year, grade_year
//	) AS 語文不及格人數統計
//		ON 語文不及格人數統計.school_year = 年級人數統計.school_year
//		AND 語文不及格人數統計.grade_year = 年級人數統計.grade_year
//	LEFT OUTER JOIN (
//		SELECT school_year, grade_year, SUM(CASE WHEN pass THEN 0 ELSE 1 END) AS sum
//		FROM 成績資料
//		WHERE 領域 = '數學'
//		GROUP BY school_year, grade_year
//	) AS 數學不及格人數統計
//		ON 數學不及格人數統計.school_year = 年級人數統計.school_year
//		AND 數學不及格人數統計.grade_year = 年級人數統計.grade_year
//	LEFT OUTER JOIN (
//		SELECT school_year, grade_year, SUM(CASE WHEN pass THEN 0 ELSE 1 END) AS sum
//		FROM 成績資料
//		WHERE 領域 = '社會'
//		GROUP BY school_year, grade_year
//	) AS 社會不及格人數統計
//		ON 社會不及格人數統計.school_year = 年級人數統計.school_year
//		AND 社會不及格人數統計.grade_year = 年級人數統計.grade_year
//	LEFT OUTER JOIN (
//		SELECT school_year, grade_year, SUM(CASE WHEN pass THEN 0 ELSE 1 END) AS sum
//		FROM 成績資料
//		WHERE 領域 = '自然與生活科技'
//		GROUP BY school_year, grade_year
//	) AS 自然與生活科技不及格人數統計
//		ON 自然與生活科技不及格人數統計.school_year = 年級人數統計.school_year
//		AND 自然與生活科技不及格人數統計.grade_year = 年級人數統計.grade_year
//	LEFT OUTER JOIN (
//		SELECT school_year, grade_year, SUM(CASE WHEN pass THEN 0 ELSE 1 END) AS sum
//		FROM 成績資料
//		WHERE 領域 = '健康與體育'
//		GROUP BY school_year, grade_year
//	) AS 健康與體育不及格人數統計
//		ON 健康與體育不及格人數統計.school_year = 年級人數統計.school_year
//		AND 健康與體育不及格人數統計.grade_year = 年級人數統計.grade_year
//	LEFT OUTER JOIN (
//		SELECT school_year, grade_year, SUM(CASE WHEN pass THEN 0 ELSE 1 END) AS sum
//		FROM 成績資料
//		WHERE 領域 = '藝術與人文'
//		GROUP BY school_year, grade_year
//	) AS 藝術與人文不及格人數統計
//		ON 藝術與人文不及格人數統計.school_year = 年級人數統計.school_year
//		AND 藝術與人文不及格人數統計.grade_year = 年級人數統計.grade_year
//	LEFT OUTER JOIN (
//		SELECT school_year, grade_year, SUM(CASE WHEN pass THEN 0 ELSE 1 END) AS sum
//		FROM 成績資料
//		WHERE 領域 = '綜合活動'
//		GROUP BY school_year, grade_year
//	) AS 綜合活動不及格人數統計
//		ON 綜合活動不及格人數統計.school_year = 年級人數統計.school_year
//		AND 綜合活動不及格人數統計.grade_year = 年級人數統計.grade_year
//	LEFT OUTER JOIN (
//		SELECT school_year, grade_year, count(*)
//		FROM 個人不及格領域數統計
//		WHERE 不及格領域數 = '1'
//		GROUP BY school_year, grade_year
//	) AS 不及格領域數1
//		ON 不及格領域數1.school_year = 年級人數統計.school_year
//		AND 不及格領域數1.grade_year = 年級人數統計.grade_year
//	LEFT OUTER JOIN (
//		SELECT school_year, grade_year, count(*)
//		FROM 個人不及格領域數統計
//		WHERE 不及格領域數 = '2'
//		GROUP BY school_year, grade_year
//	) AS 不及格領域數2
//		ON 不及格領域數2.school_year = 年級人數統計.school_year
//		AND 不及格領域數2.grade_year = 年級人數統計.grade_year
//	LEFT OUTER JOIN (
//		SELECT school_year, grade_year, count(*)
//		FROM 個人不及格領域數統計
//		WHERE 不及格領域數 = '3'
//		GROUP BY school_year, grade_year
//	) AS 不及格領域數3
//		ON 不及格領域數3.school_year = 年級人數統計.school_year
//		AND 不及格領域數3.grade_year = 年級人數統計.grade_year
//	LEFT OUTER JOIN (
//		SELECT school_year, grade_year, count(*)
//		FROM 個人不及格領域數統計
//		WHERE 不及格領域數 = '4'
//		GROUP BY school_year, grade_year
//	) AS 不及格領域數4
//		ON 不及格領域數4.school_year = 年級人數統計.school_year
//		AND 不及格領域數4.grade_year = 年級人數統計.grade_year
//	LEFT OUTER JOIN (
//		SELECT school_year, grade_year, count(*)
//		FROM 個人不及格領域數統計
//		WHERE 不及格領域數 = '5'
//		GROUP BY school_year, grade_year
//	) AS 不及格領域數5
//		ON 不及格領域數5.school_year = 年級人數統計.school_year
//		AND 不及格領域數5.grade_year = 年級人數統計.grade_year
//	LEFT OUTER JOIN (
//		SELECT school_year, grade_year, count(*)
//		FROM 個人不及格領域數統計
//		WHERE 不及格領域數 = '6'
//		GROUP BY school_year, grade_year
//	) AS 不及格領域數6
//		ON 不及格領域數6.school_year = 年級人數統計.school_year
//		AND 不及格領域數6.grade_year = 年級人數統計.grade_year
//	LEFT OUTER JOIN (
//		SELECT school_year, grade_year, count(*)
//		FROM 個人不及格領域數統計
//		WHERE 不及格領域數 = '7'
//		GROUP BY school_year, grade_year
//	) AS 不及格領域數7
//		ON 不及格領域數7.school_year = 年級人數統計.school_year
//		AND 不及格領域數7.grade_year = 年級人數統計.grade_year
//	LEFT OUTER JOIN 補考資料
//		ON 補考資料.grade_year = 年級人數統計.grade_year
//ORDER BY 年級人數統計.grade_year";

        
    }
}