
package com.tradevan.util.report;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFFooter;
import org.apache.poi.hssf.usermodel.HSSFPrintSetup;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import com.tradevan.util.dao.RdbCommonDao;
import com.tradevan.util.report.RptFL015W.ReportBean;
import com.tradevan.util.Utility;
/***
 * 個別農漁會對其政策性農業專案貸款檢查缺失改善處理情形
 * @author 2808
 *
 */
public class RptFL017W {
	/**Report執行SQL*/
	private static String QUERY_SQL =  "" ;
	/**報表名稱*/
	private static String FILE_NAME = "" ;
	
	private static String UI_VAL = "" ;
	
	private static String QUERY_STD = "" ;
	private static String QUERY_ETD = "" ;
	private static String QUERY_VAL = "" ;
	private static String STAT_TYPE = "" ;
	
	private static Connection con = null ;
	
	
	private static ArrayList <ReportBean> DATA_LIST = new ArrayList <ReportBean>(); 
	
	//private static HSSFRow row = null; //宣告一列
	//private static HSSFCell cell = null; //宣告一個儲存格
	
	/**農漁會SQL*/
    private static void   getReportSQL1() {
    	StringBuffer sql = new StringBuffer() ;
    	sql.append(" select ");
		sql.append(" ex_type, ");//--查核類別 FEB:金管會檢查報告 AGRI:農業金庫查核 BOAF:農金局訪查  
		sql.append(" decode(ex_type,'FEB','金管會檢查報告','AGRI','農業金庫查核','BOAF','農金局訪查','') as ex_type_name,");//--查核類別
		sql.append(" frm_exdef.ex_no,");//--查核報告編號
		sql.append(" decode(ex_type,'FEB',frm_exdef.ex_no,'AGRI',substr(frm_exdef.ex_no,0,3)||'年第'||substr(frm_exdef.ex_no,4,2)||'季','BOAF',F_TRANSCHINESEDATE(to_date(frm_exdef.ex_no,'yyyymmdd')),'') as ex_no_list,");// --報表顯示的檢查報告編號或查核季別或訪查日期
		//sql.append(" --frm_exdef.def_seq, ");//--缺失序號
		sql.append(" loan_name,");//--借款人名稱
		sql.append(" F_TRANSCHINESEDATE(loan_date) as loan_date,");//--貸款日期
		sql.append(" frm_exdef.loan_item,loan_item_name,");//--貸款種類
		sql.append(" loan_amt,");//--貸款金額
		sql.append(" F_COMBEX_DEF_CASE (frm_exmaster.bank_no,frm_exmaster.ex_no,loan_name,to_char(loan_date,'yyyy/mm/dd'),frm_exdef.loan_item,loan_amt) as def_case,");// --缺失摘要(用來排序用)
		sql.append(" frm_snrtdoc.non_loan_amt,");//--不符規定貸款金額
		sql.append(" frm_snrtdoc.pay_amt,");//--繳還補貼息
		sql.append(" re_pay_amt,");// --少計補貼息金額
		sql.append(" frm_refunddoc.refund_amt,");// --溢繳金額
		sql.append(" decode(frm_snrtdoc.pay_amt,null,0,'',0,frm_snrtdoc.pay_amt) + decode(re_pay_amt,null,0,'',0,re_pay_amt) - decode(frm_refunddoc.refund_amt,null,0,'',0,frm_refunddoc.refund_amt)  as pay_amt_count,");// --繳還補貼息金額(淨額)=原已繳補貼息+少繳-溢繳
		sql.append(" F_TRANSCHINESEDATE(frm_snrtdoc.pay_date) as pay_date, ");//--繳還日期
		sql.append(" F_GETFRM_MEMO(frm_exdef.bank_no,frm_exdef.ex_no,frm_exdef.loan_name,to_char(frm_exdef.loan_date,'yyyy/mm/dd'),frm_exdef.loan_item,frm_exdef.loan_amt) as pay_status ");//備註
		sql.append(" from frm_exdef ");
		sql.append(" left join (select * from bn01 where m_year=100)bn01 on frm_exdef.bank_no=bn01.bank_no ");
		sql.append(" left join frm_loan_item on frm_exdef.loan_item = frm_loan_item.loan_item ");
		sql.append(" left join frm_exmaster on frm_exdef.ex_no=frm_exmaster.ex_no and "); 
		sql.append(" frm_exdef.bank_no=frm_exmaster.bank_no") ;
		sql.append(" left join frm_def_item on frm_exdef.def_type = frm_def_item.def_type and frm_exdef.def_case = frm_def_item.def_case ");
		sql.append(" left join frm_snrtdoc on frm_exdef.ex_no=frm_snrtdoc.ex_no and frm_exdef.bank_no=frm_snrtdoc.bank_no and frm_exdef.def_seq=frm_snrtdoc.def_seq ");
		sql.append(" left join frm_agsncorrdoc on frm_exdef.ex_no=frm_agsncorrdoc.ex_no and frm_exdef.bank_no=frm_agsncorrdoc.bank_no and frm_exdef.def_seq=frm_agsncorrdoc.def_seq ");
		sql.append(" left join frm_refunddoc on frm_exdef.ex_no=frm_refunddoc.ex_no and frm_exdef.bank_no=frm_refunddoc.bank_no and frm_exdef.def_seq=frm_refunddoc.def_seq ");
		sql.append(" where frm_exdef.bank_no=? ");//'UI.受檢單位 ex:6030016' --受檢單位(若統計分類=農漁會別,才加入)
		sql.append(" and frm_snrtdoc.audit_id in ('A2','A3') ");// --核處情形為A2:繳還補貼息 A3:調整貸款期限
		if(!"".equals(QUERY_STD) && !"".equals(QUERY_ETD)) {
			sql.append(" and TO_CHAR(frm_snrtdoc.pay_date ,'yyyymmdd') BETWEEN ? and ? ");//'UI.繳還日期begin ex:20160101' AND 'UI.繳還日期end ex:20161009'  --繳還日期有值時,才加入
		}
		sql.append(" group by "); 
		sql.append(" ex_type,frm_exdef.ex_no,loan_name,loan_date,frm_exdef.loan_item,loan_item_name,loan_amt,F_COMBEX_DEF_CASE");
		sql.append(" (frm_exmaster.bank_no,frm_exmaster.ex_no,loan_name,to_char(loan_date,'yyyy/mm/dd'),frm_exdef.loan_item,loan_amt),frm_snrtdoc.non_loan_amt,frm_snrtdoc.pay_amt,re_pay_amt,frm_refunddoc.refund_amt,frm_snrtdoc.pay_date,");
		sql.append(" F_GETFRM_MEMO(frm_exdef.bank_no,frm_exdef.ex_no,frm_exdef.loan_name,to_char(frm_exdef.loan_date,'yyyy/mm/dd'),frm_exdef.loan_item,frm_exdef.loan_amt)") ;
		sql.append(" order by ");
		sql.append(" ex_type asc,ex_no asc,def_case ");
		QUERY_SQL = sql.toString() ;
    }
    public static void getReportSQL1CaseCount() {
    	StringBuffer sql = new StringBuffer() ; 
    	sql.append(" Select count(loan_name) as case_count , ") ;
    	sql.append(" sum(non_loan_amt) as non_loan_amt_sum,  ") ;
    	sql.append(" sum(pay_amt_count) as pay_amt_count  ");
    	sql.append(" from ( ") ;
    	getReportSQL1() ;
    	sql.append(QUERY_SQL) ;
    	sql.append(" ) ") ;
    	QUERY_SQL = sql.toString() ;
    }
    /**貸款種類 SQL*/
    private static void   getReportSQL2 () { 
    	StringBuffer sql = new StringBuffer() ;
    	sql.append(" select ");
		sql.append(" frm_exdef.bank_no,");//--農漁會別.機構代碼  統計分類不為農漁會別才加入
		sql.append(" bank_name,");//--農漁會別.機構名稱 統計分類不為農漁會別才加入
		sql.append(" ex_type,");//--查核類別 FEB:金管會檢查報告 AGRI:農業金庫查核 BOAF:農金局訪查  
		sql.append(" decode(ex_type,'FEB','金管會檢查報告','AGRI','農業金庫查核','BOAF','農金局訪查','') as ex_type_name,");//--查核類別
		sql.append(" frm_exdef.ex_no,");//--查核報告編號
		sql.append(" decode(ex_type,'FEB',frm_exdef.ex_no,'AGRI',substr(frm_exdef.ex_no,0,3)||'年第'||substr(frm_exdef.ex_no,4,2)||'季','BOAF',F_TRANSCHINESEDATE(to_date(frm_exdef.ex_no,'yyyymmdd')),'') as ex_no_list, ");//--報表顯示的檢查報告編號或查核季別或訪查日期
		//sql.append(" --frm_exdef.def_seq,");//--缺失序號
		sql.append(" loan_name,");//--借款人名稱
		sql.append(" F_TRANSCHINESEDATE(loan_date) as loan_date,");//--貸款日期
		sql.append(" frm_exdef.loan_item,loan_item_name,");//--貸款種類
		sql.append(" loan_amt,");//--貸款金額
		sql.append(" F_COMBEX_DEF_CASE (frm_exmaster.bank_no,frm_exmaster.ex_no,loan_name,to_char(loan_date,'yyyy/mm/dd'),frm_exdef.loan_item,loan_amt) as def_case, ");//--缺失摘要(用來排序用)
		sql.append(" frm_snrtdoc.non_loan_amt,");//--不符規定貸款金額
		sql.append(" frm_snrtdoc.pay_amt,");//--繳還補貼息
		sql.append(" re_pay_amt,");// --少計補貼息金額
		sql.append(" frm_refunddoc.refund_amt, ");//--溢繳金額
		sql.append(" decode(frm_snrtdoc.pay_amt,null,0,'',0,frm_snrtdoc.pay_amt) + decode(re_pay_amt,null,0,'',0,re_pay_amt) - decode(frm_refunddoc.refund_amt,null,0,'',0,frm_refunddoc.refund_amt)  as pay_amt_count,");// --繳還補貼息金額(淨額)=原已繳補貼息+少繳-溢繳
		sql.append(" F_TRANSCHINESEDATE(frm_snrtdoc.pay_date) as pay_date, ");//--繳還日期
		sql.append(" F_GETFRM_MEMO(frm_exdef.bank_no,frm_exdef.ex_no,frm_exdef.loan_name,to_char(frm_exdef.loan_date,'yyyy/mm/dd'),frm_exdef.loan_item,frm_exdef.loan_amt) as pay_status ");//備註
		sql.append(" from frm_exdef ");
		sql.append(" left join (select * from bn01 where m_year=100)bn01 on frm_exdef.bank_no=bn01.bank_no ");
		sql.append(" left join frm_loan_item on frm_exdef.loan_item = frm_loan_item.loan_item ");
		sql.append(" left join frm_exmaster on frm_exdef.ex_no=frm_exmaster.ex_no and "); 
		sql.append(" frm_exdef.bank_no=frm_exmaster.bank_no") ;
		sql.append(" left join frm_def_item on frm_exdef.def_type = frm_def_item.def_type and frm_exdef.def_case = frm_def_item.def_case ");
		sql.append(" left join frm_snrtdoc on frm_exdef.ex_no=frm_snrtdoc.ex_no and frm_exdef.bank_no=frm_snrtdoc.bank_no and frm_exdef.def_seq=frm_snrtdoc.def_seq ");
		sql.append(" left join frm_agsncorrdoc on frm_exdef.ex_no=frm_agsncorrdoc.ex_no and frm_exdef.bank_no=frm_agsncorrdoc.bank_no and frm_exdef.def_seq=frm_agsncorrdoc.def_seq ");
		sql.append(" left join frm_refunddoc on frm_exdef.ex_no=frm_refunddoc.ex_no and frm_exdef.bank_no=frm_refunddoc.bank_no and frm_exdef.def_seq=frm_refunddoc.def_seq ");
		sql.append(" where frm_exdef.loan_item = ? ") ;//'UI.貸款種類 ex:03' --貸款種類(若統計分類=貸款種類別,才加入)
		sql.append(" and frm_snrtdoc.audit_id in ('A2','A3') ");// --核處情形為A2:繳還補貼息 A3:調整貸款期限
		if(!"".equals(QUERY_STD) && !"".equals(QUERY_ETD)) {
			sql.append(" and TO_CHAR(frm_snrtdoc.pay_date ,'yyyymmdd') BETWEEN ? and ? ");//'UI.繳還日期begin ex:20160101' AND 'UI.繳還日期end ex:20161009'  --繳還日期有值時,才加入
		}
		sql.append(" group by "); 
		sql.append(" frm_exdef.bank_no,bank_name, ") ;
		sql.append(" ex_type,frm_exdef.ex_no,loan_name,loan_date,frm_exdef.loan_item,loan_item_name,loan_amt,F_COMBEX_DEF_CASE");
		sql.append(" (frm_exmaster.bank_no,frm_exmaster.ex_no,loan_name,to_char(loan_date,'yyyy/mm/dd'),frm_exdef.loan_item,loan_amt),frm_snrtdoc.non_loan_amt,frm_snrtdoc.pay_amt,re_pay_amt,frm_refunddoc.refund_amt,frm_snrtdoc.pay_date,");
		sql.append(" F_GETFRM_MEMO(frm_exdef.bank_no,frm_exdef.ex_no,frm_exdef.loan_name,to_char(frm_exdef.loan_date,'yyyy/mm/dd'),frm_exdef.loan_item,frm_exdef.loan_amt)") ;
		sql.append(" order by frm_exdef.bank_no,bank_name,");// 統計分類不為農漁會別才加入
		sql.append(" ex_type asc,ex_no asc,def_case ");
		QUERY_SQL = sql.toString() ;
    }
    public static void getReportSQL2CaseCount() {
    	StringBuffer sql = new StringBuffer() ; 
    	sql.append(" Select count(loan_name) as case_count , ") ;
    	sql.append(" sum(non_loan_amt) as non_loan_amt_sum,  ") ;
    	sql.append(" sum(pay_amt_count) as pay_amt_count  ");
    	sql.append(" from ( ") ;
    	getReportSQL2() ;
    	sql.append(QUERY_SQL) ;
    	sql.append(" ) ") ;
    	QUERY_SQL = sql.toString() ;
    }
    public static void getBankCntSQL2() {
    	StringBuffer sql = new StringBuffer() ; 
    	sql.append(" Select count(bank_no) as bankCnt ") ;
    	sql.append(" from ( Select bank_no,count(loan_name) as case_count  ") ;
    	sql.append(" from ( ") ;
    	getReportSQL2() ;
    	sql.append(QUERY_SQL) ;
    	sql.append(" ) group by bank_no ") ;
    	sql.append(" )");
    	QUERY_SQL = sql.toString() ;
    }
    /**缺失態樣大類*/
    private static void   getReportSQL3 () { 
    	StringBuffer sql = new StringBuffer() ;
    	sql.append(" select ");
		sql.append(" frm_exdef.bank_no,");//--農漁會別.機構代碼  統計分類不為農漁會別才加入
		sql.append(" bank_name,");//--農漁會別.機構名稱 統計分類不為農漁會別才加入
		sql.append(" ex_type,");//--查核類別 FEB:金管會檢查報告 AGRI:農業金庫查核 BOAF:農金局訪查  
		sql.append(" decode(ex_type,'FEB','金管會檢查報告','AGRI','農業金庫查核','BOAF','農金局訪查','') as ex_type_name,");//--查核類別
		sql.append(" frm_exdef.ex_no,");//--查核報告編號
		sql.append(" decode(ex_type,'FEB',frm_exdef.ex_no,'AGRI',substr(frm_exdef.ex_no,0,3)||'年第'||substr(frm_exdef.ex_no,4,2)||'季','BOAF',F_TRANSCHINESEDATE(to_date(frm_exdef.ex_no,'yyyymmdd')),'') as ex_no_list, ");//--報表顯示的檢查報告編號或查核季別或訪查日期
		//sql.append(" --frm_exdef.def_seq,");//--缺失序號
		sql.append(" loan_name,");//--借款人名稱
		sql.append(" F_TRANSCHINESEDATE(loan_date) as loan_date,");//--貸款日期
		sql.append(" frm_exdef.loan_item,loan_item_name,");//--貸款種類
		sql.append(" loan_amt,");//--貸款金額
		sql.append(" F_COMBEX_DEF_CASE (frm_exmaster.bank_no,frm_exmaster.ex_no,loan_name,to_char(loan_date,'yyyy/mm/dd'),frm_exdef.loan_item,loan_amt) as def_case, ");//--缺失摘要(用來排序用)
		sql.append(" frm_snrtdoc.non_loan_amt,");//--不符規定貸款金額
		sql.append(" frm_snrtdoc.pay_amt,");//--繳還補貼息
		sql.append(" re_pay_amt,");// --少計補貼息金額
		sql.append(" frm_refunddoc.refund_amt, ");//--溢繳金額
		sql.append(" decode(frm_snrtdoc.pay_amt,null,0,'',0,frm_snrtdoc.pay_amt) + decode(re_pay_amt,null,0,'',0,re_pay_amt) - decode(frm_refunddoc.refund_amt,null,0,'',0,frm_refunddoc.refund_amt)  as pay_amt_count,");// --繳還補貼息金額(淨額)=原已繳補貼息+少繳-溢繳
		sql.append(" F_TRANSCHINESEDATE(frm_snrtdoc.pay_date) as pay_date, ");//--繳還日期
		sql.append(" F_GETFRM_MEMO(frm_exdef.bank_no,frm_exdef.ex_no,frm_exdef.loan_name,to_char(frm_exdef.loan_date,'yyyy/mm/dd'),frm_exdef.loan_item,frm_exdef.loan_amt) as pay_status ");//備註
		sql.append(" from frm_exdef ");
		sql.append(" left join (select * from bn01 where m_year=100)bn01 on frm_exdef.bank_no=bn01.bank_no ");
		sql.append(" left join frm_loan_item on frm_exdef.loan_item = frm_loan_item.loan_item ");
		sql.append(" left join frm_exmaster on frm_exdef.ex_no=frm_exmaster.ex_no and "); 
		sql.append(" frm_exdef.bank_no=frm_exmaster.bank_no") ;
		sql.append(" left join frm_def_item on frm_exdef.def_type = frm_def_item.def_type and frm_exdef.def_case = frm_def_item.def_case ");
		sql.append(" left join frm_snrtdoc on frm_exdef.ex_no=frm_snrtdoc.ex_no and frm_exdef.bank_no=frm_snrtdoc.bank_no and frm_exdef.def_seq=frm_snrtdoc.def_seq ");
		sql.append(" left join frm_agsncorrdoc on frm_exdef.ex_no=frm_agsncorrdoc.ex_no and frm_exdef.bank_no=frm_agsncorrdoc.bank_no and frm_exdef.def_seq=frm_agsncorrdoc.def_seq ");
		sql.append(" left join frm_refunddoc on frm_exdef.ex_no=frm_refunddoc.ex_no and frm_exdef.bank_no=frm_refunddoc.bank_no and frm_exdef.def_seq=frm_refunddoc.def_seq ");
		sql.append(" where frm_exdef.def_type= ? ") ;
		sql.append(" and frm_snrtdoc.audit_id in ('A2','A3') ");// --核處情形為A2:繳還補貼息 A3:調整貸款期限
		if(!"".equals(QUERY_STD) && !"".equals(QUERY_ETD)) {
			sql.append(" and TO_CHAR(frm_snrtdoc.pay_date ,'yyyymmdd') BETWEEN ? and ? ");//'UI.繳還日期begin ex:20160101' AND 'UI.繳還日期end ex:20161009'  --繳還日期有值時,才加入
		}
		sql.append(" group by "); 
		sql.append(" frm_exdef.bank_no,bank_name, ") ;
		sql.append(" ex_type,frm_exdef.ex_no,loan_name,loan_date,frm_exdef.loan_item,loan_item_name,loan_amt,F_COMBEX_DEF_CASE");
		sql.append(" (frm_exmaster.bank_no,frm_exmaster.ex_no,loan_name,to_char(loan_date,'yyyy/mm/dd'),frm_exdef.loan_item,loan_amt),frm_snrtdoc.non_loan_amt,frm_snrtdoc.pay_amt,re_pay_amt,frm_refunddoc.refund_amt,frm_snrtdoc.pay_date,");
		sql.append(" F_GETFRM_MEMO(frm_exdef.bank_no,frm_exdef.ex_no,frm_exdef.loan_name,to_char(frm_exdef.loan_date,'yyyy/mm/dd'),frm_exdef.loan_item,frm_exdef.loan_amt)") ;
		sql.append(" order by frm_exdef.bank_no,bank_name,");// 統計分類不為農漁會別才加入
		sql.append(" ex_type asc,ex_no asc,def_case ");
		QUERY_SQL = sql.toString() ;
    }
    public static void getReportSQL3CaseCount() {
    	StringBuffer sql = new StringBuffer() ; 
    	sql.append(" Select count(loan_name) as case_count , ") ;
    	sql.append(" sum(non_loan_amt) as non_loan_amt_sum,  ") ;
    	sql.append(" sum(pay_amt_count) as pay_amt_count  ");
    	sql.append(" from ( ") ;
    	getReportSQL3() ;
    	sql.append(QUERY_SQL) ;
    	sql.append(" ) ") ;
    	QUERY_SQL = sql.toString() ;
    }
    public static void getBankCntSQL3() {
    	StringBuffer sql = new StringBuffer() ; 
    	sql.append(" Select count(bank_no) as bankCnt ") ;
    	sql.append(" from ( Select bank_no,count(loan_name) as case_count  ") ;
    	sql.append(" from ( ") ;
    	getReportSQL3() ;
    	sql.append(QUERY_SQL) ;
    	sql.append(" ) group by bank_no ") ;
    	sql.append(" )");
    	QUERY_SQL = sql.toString() ;
    }
    private static void   getReportSQL4 () { 
    	StringBuffer sql = new StringBuffer() ;
    	sql.append(" select ");
		sql.append(" frm_exdef.bank_no,");//--農漁會別.機構代碼  統計分類不為農漁會別才加入
		sql.append(" bank_name,");//--農漁會別.機構名稱 統計分類不為農漁會別才加入
		sql.append(" ex_type,");//--查核類別 FEB:金管會檢查報告 AGRI:農業金庫查核 BOAF:農金局訪查  
		sql.append(" decode(ex_type,'FEB','金管會檢查報告','AGRI','農業金庫查核','BOAF','農金局訪查','') as ex_type_name,");//--查核類別
		sql.append(" frm_exdef.ex_no,");//--查核報告編號
		sql.append(" decode(ex_type,'FEB',frm_exdef.ex_no,'AGRI',substr(frm_exdef.ex_no,0,3)||'年第'||substr(frm_exdef.ex_no,4,2)||'季','BOAF',F_TRANSCHINESEDATE(to_date(frm_exdef.ex_no,'yyyymmdd')),'') as ex_no_list, ");//--報表顯示的檢查報告編號或查核季別或訪查日期
		//sql.append(" --frm_exdef.def_seq,");//--缺失序號
		sql.append(" loan_name,");//--借款人名稱
		sql.append(" F_TRANSCHINESEDATE(loan_date) as loan_date,");//--貸款日期
		sql.append(" frm_exdef.loan_item,loan_item_name,");//--貸款種類
		sql.append(" loan_amt,");//--貸款金額
		sql.append(" F_COMBEX_DEF_CASE (frm_exmaster.bank_no,frm_exmaster.ex_no,loan_name,to_char(loan_date,'yyyy/mm/dd'),frm_exdef.loan_item,loan_amt) as def_case, ");//--缺失摘要(用來排序用)
		sql.append(" frm_snrtdoc.non_loan_amt,");//--不符規定貸款金額
		sql.append(" frm_snrtdoc.pay_amt,");//--繳還補貼息
		sql.append(" re_pay_amt,");// --少計補貼息金額
		sql.append(" frm_refunddoc.refund_amt, ");//--溢繳金額
		sql.append(" decode(frm_snrtdoc.pay_amt,null,0,'',0,frm_snrtdoc.pay_amt) + decode(re_pay_amt,null,0,'',0,re_pay_amt) - decode(frm_refunddoc.refund_amt,null,0,'',0,frm_refunddoc.refund_amt)  as pay_amt_count,");// --繳還補貼息金額(淨額)=原已繳補貼息+少繳-溢繳
		sql.append(" F_TRANSCHINESEDATE(frm_snrtdoc.pay_date) as pay_date, ");//--繳還日期
		sql.append(" F_GETFRM_MEMO(frm_exdef.bank_no,frm_exdef.ex_no,frm_exdef.loan_name,to_char(frm_exdef.loan_date,'yyyy/mm/dd'),frm_exdef.loan_item,frm_exdef.loan_amt) as pay_status ");//備註
		sql.append(" from frm_exdef ");
		sql.append(" left join (select * from bn01 where m_year=100)bn01 on frm_exdef.bank_no=bn01.bank_no ");
		sql.append(" left join frm_loan_item on frm_exdef.loan_item = frm_loan_item.loan_item ");
		sql.append(" left join frm_exmaster on frm_exdef.ex_no=frm_exmaster.ex_no and "); 
		sql.append(" frm_exdef.bank_no=frm_exmaster.bank_no") ;
		sql.append(" left join frm_def_item on frm_exdef.def_type = frm_def_item.def_type and frm_exdef.def_case = frm_def_item.def_case ");
		sql.append(" left join frm_snrtdoc on frm_exdef.ex_no=frm_snrtdoc.ex_no and frm_exdef.bank_no=frm_snrtdoc.bank_no and frm_exdef.def_seq=frm_snrtdoc.def_seq ");
		sql.append(" left join frm_agsncorrdoc on frm_exdef.ex_no=frm_agsncorrdoc.ex_no and frm_exdef.bank_no=frm_agsncorrdoc.bank_no and frm_exdef.def_seq=frm_agsncorrdoc.def_seq ");
		sql.append(" left join frm_refunddoc on frm_exdef.ex_no=frm_refunddoc.ex_no and frm_exdef.bank_no=frm_refunddoc.bank_no and frm_exdef.def_seq=frm_refunddoc.def_seq ");
		sql.append(" where frm_exmaster.ex_type= ? ") ;
		sql.append(" and frm_snrtdoc.audit_id in ('A2','A3') ");// --核處情形為A2:繳還補貼息 A3:調整貸款期限
		if(!"".equals(QUERY_STD) && !"".equals(QUERY_ETD)) {
			sql.append(" and TO_CHAR(frm_snrtdoc.pay_date ,'yyyymmdd') BETWEEN ? and ? ");//'UI.繳還日期begin ex:20160101' AND 'UI.繳還日期end ex:20161009'  --繳還日期有值時,才加入
		}
		sql.append(" group by "); 
		sql.append(" frm_exdef.bank_no,bank_name, ") ;
		sql.append(" ex_type,frm_exdef.ex_no,loan_name,loan_date,frm_exdef.loan_item,loan_item_name,loan_amt,F_COMBEX_DEF_CASE");
		sql.append(" (frm_exmaster.bank_no,frm_exmaster.ex_no,loan_name,to_char(loan_date,'yyyy/mm/dd'),frm_exdef.loan_item,loan_amt),frm_snrtdoc.non_loan_amt,frm_snrtdoc.pay_amt,re_pay_amt,frm_refunddoc.refund_amt,frm_snrtdoc.pay_date,");
		sql.append(" F_GETFRM_MEMO(frm_exdef.bank_no,frm_exdef.ex_no,frm_exdef.loan_name,to_char(frm_exdef.loan_date,'yyyy/mm/dd'),frm_exdef.loan_item,frm_exdef.loan_amt)") ;
		sql.append(" order by frm_exdef.bank_no,bank_name,");// 統計分類不為農漁會別才加入
		sql.append(" ex_type asc,ex_no asc,def_case ");
		QUERY_SQL = sql.toString() ;
    }
    public static void getReportSQL4CaseCount() {
    	StringBuffer sql = new StringBuffer() ; 
    	sql.append(" Select count(loan_name) as case_count , ") ;
    	sql.append(" sum(non_loan_amt) as non_loan_amt_sum,  ") ;
    	sql.append(" sum(pay_amt_count) as pay_amt_count  ");
    	sql.append(" from ( ") ;
    	getReportSQL4() ;
    	sql.append(QUERY_SQL) ;
    	sql.append(" ) ") ;
    	QUERY_SQL = sql.toString() ;
    }
    public static void getBankCntSQL4() {
    	StringBuffer sql = new StringBuffer() ; 
    	sql.append(" Select count(bank_no) as bankCnt ") ;
    	sql.append(" from ( Select bank_no,count(loan_name) as case_count  ") ;
    	sql.append(" from ( ") ;
    	getReportSQL4() ;
    	sql.append(QUERY_SQL) ;
    	sql.append(" ) group by bank_no ") ;
    	sql.append(" )");
    	QUERY_SQL = sql.toString() ;
    }
    private static void getQueryData (Connection con , ResultSet rs ) {
    	QUERY_SQL = "" ;
    	try {
    		if("1".equals(STAT_TYPE)) {
    			getReportSQL1() ;
    		} else if("2".equals(STAT_TYPE)) {
    			getReportSQL2() ;
    		} else if("3".equals(STAT_TYPE)) {
    			getReportSQL3() ;
    		} else if("4".equals(STAT_TYPE)) {
    			getReportSQL4() ;
    		}
    		//System.out.println(" query sql :"+QUERY_SQL); 
    		PreparedStatement  stmt = con.prepareStatement(QUERY_SQL) ;
    		stmt.setString(1, QUERY_VAL) ; 
    		//System.out.println(" param1:"+QUERY_VAL);
    		if(!"".equals(QUERY_STD) && !"".equals(QUERY_ETD)) {
    			stmt.setString(2, transDateStyle(2,QUERY_STD));
    			stmt.setString(3, transDateStyle(2,QUERY_ETD));
    			//System.out.println(" param2:"+transDateStyle(2,QUERY_STD));
    			//System.out.println(" param3:"+transDateStyle(2,QUERY_ETD));
    		}
    		rs = stmt.executeQuery() ;
    		DATA_LIST.clear(); 
    		while(rs.next() ) {
    			ReportBean bean = new ReportBean () ;
    			if(!"1".equals(STAT_TYPE)) {
    				bean.setBank_no(rs.getString("bank_no")==null?"":rs.getString("bank_no"));
        			bean.setBank_name(rs.getString("bank_name")==null?"":rs.getString("bank_name"));
    			}
    			bean.setEx_type(rs.getString("ex_type")==null?"":rs.getString("ex_type"));
    			bean.setEx_type_name(rs.getString("ex_type_name")==null?"":rs.getString("ex_type_name"));
    			bean.setEx_no(rs.getString("ex_no")==null?"":rs.getString("ex_no"));
    			bean.setEx_no_list(rs.getString("ex_no_list")==null?"":rs.getString("ex_no_list"));
    			bean.setLoan_name(rs.getString("loan_name")==null?"":rs.getString("loan_name"));
    			bean.setLoan_date(rs.getString("loan_date")==null?"":rs.getString("loan_date"));
    			bean.setLoan_item(rs.getString("loan_item")==null?"":rs.getString("loan_item"));
    			bean.setLoan_item_name(rs.getString("loan_item_name")==null?"":rs.getString("loan_item_name"));
    			bean.setLoan_amt(rs.getString("loan_amt")==null?"":rs.getString("loan_amt"));
    			bean.setDef_case(rs.getString("def_case")==null?"":rs.getString("def_case"));
    			bean.setNon_loan_amt(rs.getString("non_loan_amt")==null?"":rs.getString("non_loan_amt"));
    			bean.setPay_amt(rs.getString("pay_amt")==null?"":rs.getString("pay_amt"));
    			bean.setRe_pay_amt(rs.getString("re_pay_amt")==null?"":rs.getString("re_pay_amt"));
    			bean.setPay_amt_count(rs.getString("pay_amt_count")==null?"":rs.getString("pay_amt_count"));
    			bean.setPay_date(rs.getString("pay_date")==null?"":rs.getString("pay_date"));
    			bean.setPay_status(rs.getString("pay_status")==null?"":rs.getString("pay_status"));
    			
    			DATA_LIST.add(bean) ;
    		}
    		rs.close(); 
    		stmt.close(); 
    	} catch(Exception e) {
    		e.printStackTrace(); 
    	} 
    }
    private static void getQueryDataCaseCnt (Connection con , ResultSet rs ) {
    	QUERY_SQL = "" ;
    	try {
    		
    		if("1".equals(STAT_TYPE)) {
    			getReportSQL1CaseCount() ;
    		} else if("2".equals(STAT_TYPE)) {
    			getReportSQL2CaseCount() ;
    		} else if("3".equals(STAT_TYPE)) {
    			getReportSQL3CaseCount() ;
    		} else if("4".equals(STAT_TYPE)) {
    			getReportSQL4CaseCount() ;
    		}
    		System.out.println(" getQueryDataCaseCnt1 SQL="+QUERY_SQL);
    		PreparedStatement  stmt = con.prepareStatement(QUERY_SQL) ;
    		stmt.setString(1, QUERY_VAL) ; 
    		System.out.println(" param1="+QUERY_VAL );
    		if(!"".equals(QUERY_STD) && !"".equals(QUERY_ETD)) {
    			stmt.setString(2, transDateStyle(2,QUERY_STD));
    			stmt.setString(3, transDateStyle(2,QUERY_ETD));
    			
    		}
    		rs = stmt.executeQuery() ;
    		DATA_LIST.clear(); 
    		while(rs.next() ) {
    			ReportBean bean = new ReportBean () ;
    			bean.setCase_count(Utility.getTrimString(rs.getString("case_count")));
    			bean.setNon_loan_amt_sum(Utility.getTrimString(rs.getString("non_loan_amt_sum")));
    			bean.setPay_amt_count(Utility.getTrimString(rs.getString("pay_amt_count")));
    			DATA_LIST.add(bean) ;
    		}
    		rs.close(); 
    		stmt.close(); 
    	} catch(Exception e) {
    		e.printStackTrace(); 
    	} 
    }
    private static String getQueryBankCnt (Connection con , ResultSet rs ) {
    	QUERY_SQL = "" ;
    	String bankCnt = "0";
    	try {
    		if("2".equals(STAT_TYPE)) {
    			getBankCntSQL2() ;
    		} else if("3".equals(STAT_TYPE)) {
    			getBankCntSQL3() ;
    		} else if("4".equals(STAT_TYPE)) {
    			getBankCntSQL4() ;
    		}
    		System.out.println(" getQueryBankCnt SQL="+QUERY_SQL);
    		PreparedStatement  stmt = con.prepareStatement(QUERY_SQL) ;
    		stmt.setString(1, QUERY_VAL) ; 
    		System.out.println(" param1="+QUERY_VAL );
    		if(!"".equals(QUERY_STD) && !"".equals(QUERY_ETD)) {
    			stmt.setString(2, transDateStyle(2,QUERY_STD));
    			stmt.setString(3, transDateStyle(2,QUERY_ETD));
    			
    		}
    		rs = stmt.executeQuery() ;
    		//DATA_LIST.clear();
    		if(rs.next()) {
    			bankCnt = Utility.getTrimString(rs.getString(1)) ;
    		}
    		rs.close(); 
    		stmt.close(); 
    	} catch(Exception e) {
    		e.printStackTrace(); 
    	} 
    	return bankCnt ;
    }
    private static HSSFWorkbook  getExcelSimple() {
		String openFile= "" ;
		if("1".equals(STAT_TYPE)) {
			openFile = "FL017W收回補貼息案件明細表_農漁會別.xls" ;
		} else if("2".equals(STAT_TYPE)) {
			openFile = "FL017W收回補貼息案件明細表_貸款種類別.xls" ;
		} else if("3".equals(STAT_TYPE)) {
			openFile = "FL017W收回補貼息案件明細表_缺失態樣別.xls" ;
		} else if("4".equals(STAT_TYPE)) {
			openFile = "FL017W收回補貼息案件明細表_查核類別.xls" ;
		}
        //File reportDir = new File(Utility.getProperties("reportDir"));
		FileInputStream finput = null;
		HSSFWorkbook wb = null ;
		try {
			System.out.println(" simple file address="+Utility.getProperties("xlsDir") + System.getProperty("file.separator")+ openFile);
			
			System.out.println(" out put file at ="+Utility.getProperties("reportDir")+ System.getProperty("file.separator") + FILE_NAME );
			File xlsDir = new File(Utility.getProperties("xlsDir"));
			finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+ openFile );
			//設定FileINputStream讀取Excel檔
            POIFSFileSystem fs = new POIFSFileSystem( finput );
            if(fs==null){System.out.println("open 範本檔失敗");} else System.out.println("open 範本檔成功");
            wb = new HSSFWorkbook(fs);
            if(wb==null){System.out.println("open工作表失敗");}else System.out.println("open 工作表 成功");
		}catch(Exception e) {
			e.printStackTrace();
		}
		return wb ;
	}
	private static void outPutExcel(HSSFWorkbook wb , HSSFSheet sheet ) {
	    	
	    	try {
	    		System.out.println(" output file address="+Utility.getProperties("reportDir")+ System.getProperty("file.separator")+ FILE_NAME);
	    		
	    		File reportDir = new File(Utility.getProperties("reportDir"));
				if (!reportDir.exists()) {
					if (!Utility.mkdirs(Utility.getProperties("reportDir"))) {
						//ERROR_MSG += Utility.getProperties("reportDir") + "目錄新增失敗";
						System.out.println(Utility.getProperties("reportDir") + "目錄新增失敗") ;
					}
				}
				
				
				FileOutputStream fout = 
						new FileOutputStream(reportDir + System.getProperty("file.separator")+ FILE_NAME);
				
				//HSSFFooter footer = sheet.getFooter();
				//footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
				//footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
				wb.write(fout);
				//儲存
				fout.close();
	    	}catch(Exception e) {
	    		e.printStackTrace();
	    	}
	    }
	private static void setExcelCellValue(HSSFSheet sheet ,HSSFRow row,HSSFCell cell,int rowNo,short cellNo,String value,HSSFCellStyle sty) {
    	try {
    		row = sheet.createRow(rowNo) ;
    		cell = row.createCell(cellNo) ;
    		cell.setCellStyle(sty); 
    		cell.setCellValue(value); 
    		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
    	}catch(Exception e) {
    		e.printStackTrace();
    	}
    }
    
	public static String createRpt(Map h){
		
		
		try {
			//查詢參數處理
			STAT_TYPE = Utility.getTrimString(h.get("stat_tp")) ; //查詢類別
			QUERY_VAL = Utility.getTrimString(h.get("query_val")) ; 
			QUERY_STD = Utility.getTrimString(h.get("begDate")) ;
			QUERY_ETD = Utility.getTrimString(h.get("endDate")) ;
			UI_VAL = Utility.getTrimString(h.get("ui_val")) ;
			
			HSSFSheet sheet = null ;
		    HSSFWorkbook wb = null ;
		    HSSFRow row = null ; 
		    HSSFCell cell = null ;
		    HSSFCellStyle simpleStyle = null;
		    HSSFCellStyle right_style = null;
		    HSSFCellStyle statem_style = null;
		    int rowNo = 0 ;
		    //System.out.println("1.報表名稱");
			//1.報表名稱
			FILE_NAME = "RptFL017W"+getYYYYMMDDHHMMSS()+".xls" ;
			//2.取得SQL
			//System.out.println("2.取得SQL");
			//getReportSQL1() ;
			//3.取得連線
			System.out.println("3.取得連線");
			con = (new RdbCommonDao("")).newConnection();
			ResultSet rs = null ;
			//4.取得資料
			//System.out.println("4.取得資料");
			getQueryData(con,rs) ;  
			//5.產生Excel
			
			//6.get 報表範本
			//System.out.println("6.get報表範本") ;
			wb = getExcelSimple () ;
			sheet = wb.getSheetAt(0);//讀取第一個工作表，宣告其為sheet 
			HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
			if("1".equals(STAT_TYPE)) {
				ps.setScale( ( short )84 ); //列印縮放百分比
				ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
				ps.setLandscape( true ); // 設定橫式
				
				simpleStyle = sheet.getRow((short)4).getCell((short)0).getCellStyle() ;
				right_style = sheet.getRow((short)4).getCell((short)5).getCellStyle();
    			statem_style = sheet.getRow((short)4).getCell((short)9).getCellStyle();
			} else if("2".equals(STAT_TYPE)) {
				ps.setScale( ( short )76 ); //列印縮放百分比
				ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
				ps.setLandscape( true ); // 設定橫式
				
    			simpleStyle = sheet.getRow((short)4).getCell((short)0).getCellStyle() ;
    			right_style = sheet.getRow((short)4).getCell((short)5).getCellStyle();
    			statem_style = sheet.getRow((short)4).getCell((short)9).getCellStyle();
			} else if("3".equals(STAT_TYPE)) {
				ps.setScale( ( short )68 ); //列印縮放百分比
				ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
				ps.setLandscape( true ); // 設定橫式
				simpleStyle = sheet.getRow((short)4).getCell((short)0).getCellStyle() ;
				right_style = sheet.getRow((short)4).getCell((short)6).getCellStyle();
    			statem_style = sheet.getRow((short)4).getCell((short)10).getCellStyle();
			} else if("4".equals(STAT_TYPE)) {
				ps.setScale( ( short )80 ); //列印縮放百分比
				ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
				ps.setLandscape( true ); // 設定橫式
    			simpleStyle = sheet.getRow((short)4).getCell((short)0).getCellStyle() ;
    			right_style = sheet.getRow((short)4).getCell((short)5).getCellStyle();
    			statem_style = sheet.getRow((short)4).getCell((short)9).getCellStyle();
			}
			//繳還日期 
			row = sheet.getRow((short)1) ;
			cell = row.getCell((short)0) ;
			if(!"".equals(QUERY_STD) && !"".equals(QUERY_ETD))  {
				cell.setCellValue("繳還日期："+transChtDateFormate(QUERY_STD)+"至"+transChtDateFormate(QUERY_ETD));
			} else {
				cell.setCellValue("");
			}
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			//農漁會別
			System.out.println("start set report data..........");
			rowNo = 4 ;
			if("1".equals(STAT_TYPE)) {
				
				row = sheet.getRow((short)2) ;
				cell = row.getCell((short)0) ;
				cell.setCellValue("農漁會別："+UI_VAL); 
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				for(int i=0 ; i < DATA_LIST.size() ;i++) {
					ReportBean rb = DATA_LIST.get(i) ;
					int cellNo = 0 ;
					if(i==0) {
						row = sheet.getRow(rowNo) ;
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(rb.getEx_type_name()); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(rb.getEx_no_list()); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(rb.getLoan_name()); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(rb.getLoan_date()); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(rb.getLoan_item_name()); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(NumberFormatMoney(rb.getLoan_amt())); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(NumberFormatMoney(rb.getNon_loan_amt()));
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(NumberFormatMoney(rb.getPay_amt_count())); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(rb.getPay_date()); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(rb.getPay_status()); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					} else {
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,rb.getEx_type_name(),simpleStyle) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,rb.getEx_no_list(),simpleStyle) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,rb.getLoan_name(),simpleStyle) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,rb.getLoan_date(),simpleStyle) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,rb.getLoan_item_name(),simpleStyle) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,NumberFormatMoney(rb.getLoan_amt()),right_style) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,NumberFormatMoney(rb.getNon_loan_amt()),right_style) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,NumberFormatMoney(rb.getPay_amt_count()),right_style) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,rb.getPay_date(),simpleStyle) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,rb.getPay_status(),statem_style) ;
					}
					++rowNo ;
				}
			} else if("2".equals(STAT_TYPE)) {
    			
    			row = sheet.getRow((short)2) ;
				cell = row.getCell((short)0) ;
				cell.setCellValue("貸款種類："+UI_VAL); 
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				for(int i=0 ; i < DATA_LIST.size() ;i++) {
					ReportBean rb = DATA_LIST.get(i) ;
					int cellNo = 0 ;
					if(i==0) {
						row = sheet.getRow(rowNo) ;
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(rb.getBank_no().concat(rb.getBank_name())); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(rb.getEx_type_name()); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(rb.getEx_no_list()); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(rb.getLoan_name()); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(rb.getLoan_date()); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(NumberFormatMoney(rb.getLoan_amt())); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(NumberFormatMoney(rb.getNon_loan_amt()));
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(NumberFormatMoney(rb.getPay_amt_count())); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(rb.getPay_date()); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(rb.getPay_status()); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					} else {
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,rb.getBank_no().concat(rb.getBank_name()),simpleStyle) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,rb.getEx_type_name(),simpleStyle) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,rb.getEx_no_list(),simpleStyle) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,rb.getLoan_name(),simpleStyle) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,rb.getLoan_date(),simpleStyle) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,NumberFormatMoney(rb.getLoan_amt()),right_style) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,NumberFormatMoney(rb.getNon_loan_amt()),right_style) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,NumberFormatMoney(rb.getPay_amt_count()),right_style) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,rb.getPay_date(),simpleStyle) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,rb.getPay_status(),statem_style) ;
					}
					++rowNo ;
				}
			} else if("3".equals(STAT_TYPE)) {
				
				row = sheet.getRow((short)2) ;
				cell = row.getCell((short)0) ;
				cell.setCellValue("缺失態樣別："+UI_VAL); 
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				for(int i=0 ; i < DATA_LIST.size() ;i++) {
					ReportBean rb = DATA_LIST.get(i) ;
					int cellNo = 0 ;
					if(i==0) {
						row = sheet.getRow(rowNo) ;
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(rb.getBank_no().concat(rb.getBank_name())); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(rb.getEx_type_name()); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(rb.getEx_no_list()); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(rb.getLoan_name()); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(rb.getLoan_date()); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(rb.getLoan_item_name()); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(NumberFormatMoney(rb.getLoan_amt())); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(NumberFormatMoney(rb.getNon_loan_amt()));
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(NumberFormatMoney(rb.getPay_amt_count())); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(rb.getPay_date()); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(rb.getPay_status()); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					} else {
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,rb.getBank_no().concat(rb.getBank_name()),simpleStyle) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,rb.getEx_type_name(),simpleStyle) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,rb.getEx_no_list(),simpleStyle) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,rb.getLoan_name(),simpleStyle) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,rb.getLoan_date(),simpleStyle) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,rb.getLoan_item_name(),simpleStyle) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,NumberFormatMoney(rb.getLoan_amt()),right_style) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,NumberFormatMoney(rb.getNon_loan_amt()),right_style) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,NumberFormatMoney(rb.getPay_amt_count()),right_style) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,rb.getPay_date(),simpleStyle) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,rb.getPay_status(),statem_style) ;
					}
					++rowNo ;
				}
    		} else if("4".equals(STAT_TYPE)) {
    			String cellName1 = "FEB".equals(QUERY_VAL)?"檢查報告編號":"AGRI".equals(QUERY_VAL)?"查核季別" : "訪查日期" ;
    			
    			//set查核類別
    			row = sheet.getRow((short)2) ;
				cell = row.getCell((short)0) ;
				//System.out.println("----done1");
				cell.setCellValue("查核類別："+UI_VAL); 
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				//System.out.println("----done2");
				//set 欄位名稱
				row = sheet.getRow((short)3) ;
				cell = row.getCell((short)1) ;
				cell.setCellValue(cellName1); 
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				for(int i=0 ; i < DATA_LIST.size() ;i++) {
					System.out.println("----done3");
					ReportBean rb = DATA_LIST.get(i) ;
					int cellNo = 0 ;
					if(i==0) {
						row = sheet.getRow(rowNo) ;
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(rb.getBank_no().concat(rb.getBank_name())); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(rb.getEx_no_list()); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(rb.getLoan_name()); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(rb.getLoan_date()); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(rb.getLoan_item_name()); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(NumberFormatMoney(rb.getLoan_amt())); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(NumberFormatMoney(rb.getNon_loan_amt()));
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(NumberFormatMoney(rb.getPay_amt_count())); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(rb.getPay_date()); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell = row.getCell((short)cellNo++) ;
						cell.setCellValue(rb.getPay_status()); 
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					} else {
						System.out.println("----done4");
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,rb.getBank_no().concat(rb.getBank_name()),simpleStyle) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,rb.getEx_no_list(),simpleStyle) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,rb.getLoan_name(),simpleStyle) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,rb.getLoan_date(),simpleStyle) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,rb.getLoan_item_name(),simpleStyle) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,NumberFormatMoney(rb.getLoan_amt()),right_style) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,NumberFormatMoney(rb.getNon_loan_amt()),right_style) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,NumberFormatMoney(rb.getPay_amt_count()),right_style) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,rb.getPay_date(),simpleStyle) ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)cellNo++,rb.getPay_status(),statem_style) ;
					}
					++rowNo ;
				}
    		}
			//案件數 
			//補三行空白
			int stCell = 0 ; 
			int edCell = 0 ;
			if("3".equals(STAT_TYPE)) {
				edCell = 11 ;
			} else {
				edCell = 10 ;
			}
			for(int k=0 ; k < 3 ; k++) {
				for(stCell = 0 ; stCell < edCell ; stCell++) {
					setExcelCellValue(sheet ,row,cell,rowNo,(short)stCell,"",simpleStyle) ;
				}
				++rowNo ;
			}
			//列印合計數
			getQueryDataCaseCnt(con ,rs ) ;
			for(int i=0 ; i < DATA_LIST.size() ;i++) {
				ReportBean rb = DATA_LIST.get(i) ;
				if("1".equals(STAT_TYPE) || "4".equals(STAT_TYPE)){
					setExcelCellValue(sheet ,row,cell,rowNo,(short)0,"1".equals(STAT_TYPE)?"": NumberFormat(getQueryBankCnt(con,rs)),right_style) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)1,"",right_style) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)2,NumberFormat(rb.getCase_count()),right_style) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)3,"",right_style) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)4,"",right_style) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)5,"",right_style) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)6,NumberFormatMoney(rb.getNon_loan_amt_sum()),right_style) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)7,NumberFormatMoney(rb.getPay_amt_count()),right_style) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)8,"",right_style) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)9,"",right_style) ;
				} else if("2".equals(STAT_TYPE)) {
					setExcelCellValue(sheet ,row,cell,rowNo,(short)0,NumberFormat(getQueryBankCnt(con,rs)),right_style) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)1,"",right_style) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)2,"",right_style) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)3,NumberFormat(rb.getCase_count()),right_style) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)4,"",right_style) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)5,"",right_style) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)6,NumberFormatMoney(rb.getNon_loan_amt_sum()),right_style) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)7,NumberFormatMoney(rb.getPay_amt_count()),right_style) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)8,"",right_style) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)9,"",right_style) ;
				} else {
					setExcelCellValue(sheet ,row,cell,rowNo,(short)0,"1".equals(STAT_TYPE)?"": NumberFormat(getQueryBankCnt(con,rs)),right_style) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)1,"",right_style) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)2,"",right_style) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)3,NumberFormat(rb.getCase_count()),right_style) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)4,"",right_style) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)5,"",right_style) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)6,"",right_style) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)7,NumberFormatMoney(rb.getNon_loan_amt_sum()),right_style) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)8,NumberFormatMoney(rb.getPay_amt_count()),right_style) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)9,"",right_style) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)10,"",right_style) ;
				}
				
			}
			outPutExcel(wb ,sheet) ; 
			
			
		}catch(Exception e){
			//System.out.println("createRpt Error:"+e+e.getMessage());
			e.printStackTrace();
		}finally {
			try {
				con.commit(); 
				con.close();
			}catch(Exception e){
				e.printStackTrace();
			}
		}
		return FILE_NAME;
	}

	static class ReportBean {
    	private String bank_no = "";
    	private String bank_name = "";
    	
    	private String ex_type = "";
    	private String ex_type_name = "";
    	private String ex_no = "";
    	private String ex_no_list = "";
    	private String def_seq = "";
    	private String loan_name = "";
    	private String loan_date = "";
    	private String loan_item = "";
    	private String loan_item_name ="";
    	private String loan_amt = "";
    	private String def_case = "";
    	private String non_loan_amt = "";
    	private String pay_amt = "";
    	private String re_pay_amt = "";
    	private String refund_amt = "";
    	private String pay_amt_count = "";
    	private String pay_date = "";
    	private String pay_status = "";
    	private String case_count = "";
    	private String non_loan_amt_sum  = "";
    	
    	
		public String getBank_no() {
			return bank_no;
		}
		public void setBank_no(String bank_no) {
			this.bank_no = bank_no;
		}
		public String getBank_name() {
			return bank_name;
		}
		public void setBank_name(String bank_name) {
			this.bank_name = bank_name;
		}
		public String getEx_type() {
			return ex_type;
		}
		public void setEx_type(String e_type) {
			this.ex_type = e_type;
		}
		public String getEx_type_name() {
			return ex_type_name;
		}
		public void setEx_type_name(String ex_type_name) {
			this.ex_type_name = ex_type_name;
		}
		public String getEx_no() {
			return ex_no;
		}
		public void setEx_no(String ex_no) {
			this.ex_no = ex_no;
		}
		public String getEx_no_list() {
			return ex_no_list;
		}
		public void setEx_no_list(String ex_no_list) {
			this.ex_no_list = ex_no_list;
		}
		public String getDef_seq() {
			return def_seq;
		}
		public void setDef_seq(String def_seq) {
			this.def_seq = def_seq;
		}
		public String getLoan_name() {
			return loan_name;
		}
		public void setLoan_name(String loan_name) {
			this.loan_name = loan_name;
		}
		public String getLoan_date() {
			return loan_date;
		}
		public void setLoan_date(String loan_date) {
			this.loan_date = loan_date;
		}
		public String getLoan_item_name() {
			return loan_item_name;
		}
		public void setLoan_item_name(String loan_item_name) {
			this.loan_item_name = loan_item_name;
		}
		public String getLoan_amt() {
			return loan_amt;
		}
		public void setLoan_amt(String loan_amt) {
			this.loan_amt = loan_amt;
		}
		public String getDef_case() {
			return def_case;
		}
		public void setDef_case(String def_case) {
			this.def_case = def_case;
		}
		public String getNon_loan_amt() {
			return non_loan_amt;
		}
		public void setNon_loan_amt(String non_loan_amt) {
			this.non_loan_amt = non_loan_amt;
		}
		public String getPay_amt() {
			return pay_amt;
		}
		public void setPay_amt(String pay_amt) {
			this.pay_amt = pay_amt;
		}
		public String getRe_pay_amt() {
			return re_pay_amt;
		}
		public void setRe_pay_amt(String re_pay_amt) {
			this.re_pay_amt = re_pay_amt;
		}
		public String getRefund_amt() {
			return refund_amt;
		}
		public void setRefund_amt(String refund_amt) {
			this.refund_amt = refund_amt;
		}
		public String getPay_amt_count() {
			return pay_amt_count;
		}
		public void setPay_amt_count(String pay_amt_count) {
			this.pay_amt_count = pay_amt_count;
		}
		public String getPay_date() {
			return pay_date;
		}
		public void setPay_date(String pay_date) {
			this.pay_date = pay_date;
		}
		public String getPay_status() {
			return pay_status;
		}
		public void setPay_status(String pay_status) {
			this.pay_status = pay_status;
		}
		public String getLoan_item() {
			return loan_item;
		}
		public void setLoan_item(String loan_item) {
			this.loan_item = loan_item;
		}
		public String getCase_count() {
			return case_count;
		}
		public void setCase_count(String case_count) {
			this.case_count = case_count;
		}
		public String getNon_loan_amt_sum() {
			return non_loan_amt_sum;
		}
		public void setNon_loan_amt_sum(String non_loan_amt_sum) {
			this.non_loan_amt_sum = non_loan_amt_sum;
		}
		
    	
    }
	/***
	 * 日期轉換
	 * @param p 1:西元轉民國  2:民國轉西元
	 * @param d 日期(西元8碼,民國7碼)
	 * @return
	 */
	private static String transDateStyle (int p , String d ) {
		String fd = "";
		switch( p) {
		case 1 : //西元轉民國 
			if(8==d.length()) {
				fd =  String.valueOf(Integer.parseInt(d.substring(0,4))-1911); 
				fd += d.substring(4, 8) ;
			}
			break ;
		case 2 : 
			if(7==d.length()) {
				fd =  String.valueOf(Integer.parseInt(d.substring(0,3))+1911); 
				fd += d.substring(3, 7) ;
			}
			break ;
		}
		return fd ;
	}
	private static String NumberFormatMoney(String m ) {
		String t = "" ;
		if(!"".equals(m)) {
			DecimalFormat  fat = new DecimalFormat ("$###,###") ;
			t = fat.format(Double.parseDouble(m)) ;
		} 
		return t ;
	}
	private static String NumberFormat(String m ) {
		String t = "" ;
		if(!"".equals(m)) {
			DecimalFormat  fat = new DecimalFormat ("###,###") ;
			t = fat.format(Double.parseDouble(m)) ;
		} 
		return t ;
	}
	private static String transChtDateFormate(String d) {
		if(d.length()!=7) {
			return "" ;
		}
		return d.substring(0,3).concat("年").concat(d.substring(3,5)).concat("月").concat(d.substring(5,7)).concat("日");
	}
	private static String getYYYYMMDDHHMMSS(){
    	Calendar rightNow = Calendar.getInstance();
        String year = (new Integer(rightNow.get(Calendar.YEAR))).toString();
        String month = (new Integer(rightNow.get(Calendar.MONTH) + 1)).toString();
        String day = (new Integer(rightNow.get(Calendar.DAY_OF_MONTH))).toString();
        String hour = (new Integer(rightNow.get(Calendar.HOUR_OF_DAY))).toString();
        String minute = (new Integer(rightNow.get(Calendar.MINUTE))).toString();
        String second = (new Integer(rightNow.get(Calendar.SECOND))).toString();

        if (month.length() == 1) {
            month = "0" + month;
        }
        if (day.length() == 1) {
            day = "0" + day;
        }
        if (hour.length() == 1) {
            hour = "0" + hour;
        }
        if (minute.length() == 1) {
            minute = "0" + minute;
        }
        if (second.length() == 1) {
            second = "0" + second;

        }
        return (year + month + day + hour + minute + second);
    }
}
