
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
public class RptFL016W {
	/**Report執行SQL*/
	private static String QUERY_SQL =  "" ;
	/**報表名稱*/
	private static String FILE_NAME = "" ;
	/**錯誤訊息*/
	private static String ERROR_MSG = "" ;
	/**受檢單位名稱*/
	private static String TBANK_UI = "" ;
	private static String TBANK = "" ;
	/**查核類別*/
	private static String EX_TYPE_UI = "" ;
	private static String EX_NO = "" ; 
	private static String UI_VAL = "";
	/**查核類別對應項目*/
	private static String EX_NO_LIST_UI="";
	
	private static Connection con = null ;
	
	   
	private static ArrayList <ReportBean> DATA_LIST = new ArrayList <ReportBean>(); 
	
	//private static HSSFRow row = null; //宣告一列
	//private static HSSFCell cell = null; //宣告一個儲存格
	
	/**取得缺失改善情形報表SQL*/
    private static void   getReportSQL1() {
    	StringBuffer sql = new StringBuffer() ;
    	sql.append(" select frm_snrtdoc.doc_type,");//--發文性質代碼 A:陳述意見/B:核處/C:結案
		sql.append(" cmuse_name as doc_type_name,");// --缺失改善情形
		sql.append(" F_TRANSCHINESEDATE(doc_date) as doc_date,");//--農金局發文.發文日期
		sql.append(" docno,");//--農金局發文.發文文號
		sql.append(" F_TRANSCHINESEDATE(bank_rt_doc_date) as bank_rt_doc_date,");//--農漁會回文.收文日期
		sql.append(" bank_rt_docno,");//--農漁會回文.收文文號
		sql.append(" loan_name,");//--核處情形.個案名稱-借款人名稱
		sql.append(" frm_snrtdoc.loan_item,loan_item_name,");//--核處情形.個案名稱-貸款種類
		sql.append(" F_COMBEX_DEF_CASE (bank_no,ex_no,loan_name,to_char(loan_date,'yyyy/mm/dd'),frm_snrtdoc.loan_item,loan_amt) as def_case, ");//--缺失摘要
		sql.append(" decode(frm_snrtdoc.doc_type,'C','改善情形洽悉(結案)',audit_case) as audit_case,");//--核處情形.核處內容
		sql.append(" decode(audit_id_c2,'1','裁處糾正','') as audit_id_c2, ");//--有值時,在核處情形.核處內容顯示一筆裁處糾正
		sql.append(" decode(audit_id_c1,'1','裁處罰鍰','') as audit_id_c1");//--有值時,在核處情形.核處內容顯示一筆裁處罰鍰
		sql.append(" from (select frm_snrtdoc.*,ex_kind,loan_name,loan_item,loan_amt,loan_date,def_type,def_case "); 
		sql.append(" from frm_snrtdoc ");
		sql.append(" left join frm_exdef on frm_snrtdoc.ex_no = frm_exdef.ex_no ");
		sql.append(" and frm_snrtdoc.bank_no=frm_exdef.bank_no and frm_snrtdoc.def_seq=frm_exdef.def_seq ");
		sql.append(" )frm_snrtdoc "); 
		sql.append(" left join (select * from cdshareno where cmuse_div='048')cdshareno ");
		sql.append(" on frm_snrtdoc.doc_type=cdshareno.cmuse_id ");
		sql.append(" left join frm_loan_item on frm_snrtdoc.loan_item = frm_loan_item.loan_item ");
		sql.append(" left join frm_def_item on frm_snrtdoc.def_type=frm_def_item.def_type and frm_snrtdoc.def_case = frm_def_item.def_case ");
		sql.append(" left join frm_audit_item on frm_snrtdoc.doc_type=frm_audit_item.doc_type and frm_snrtdoc.audit_type=frm_audit_item.audit_type and frm_snrtdoc.audit_id=frm_audit_item.audit_id ");
		sql.append(" Where bank_no = ? ") ;
		sql.append(" and ex_no = ? ") ;
		
		//sql.append(" where bank_no='UI.受檢單位 ex:6030016' ");
		//sql.append(" and ex_no='UI.查核報告編號ex_no ex:101C087' ");
		sql.append(" group by frm_snrtdoc.doc_type,cmuse_name,doc_date,docno,bank_rt_doc_date,bank_rt_docno,loan_name,frm_snrtdoc.loan_item,loan_item_name, ");
		sql.append(" F_COMBEX_DEF_CASE (bank_no,ex_no,loan_name,to_char(loan_date,'yyyy/mm/dd'),frm_snrtdoc.loan_item,loan_amt),audit_case,audit_id_c2,audit_id_c1 ");
		sql.append(" order by frm_snrtdoc.doc_type,doc_date,F_COMBEX_DEF_CASE (bank_no,ex_no,loan_name,to_char(loan_date,'yyyy/mm/dd'),frm_snrtdoc.loan_item,loan_amt) ");
		System.out.println("==SQL="+sql.toString());


		QUERY_SQL = sql.toString() ;
    }
    //繳還補貼息情形報表查詢SQL
    private static void   getReportSQL2 () { 
    	StringBuffer sql = new StringBuffer() ;
    	sql.append(" select F_TRANSCHINESEDATE(frm_snrtdoc.doc_date) as doc_date,");//--農金局.發文日期
		sql.append(" frm_snrtdoc.docno,");//--農金局.發文文號
		sql.append(" F_TRANSCHINESEDATE(ag_rt_doc_date) as ag_rt_doc_date, ");//--金庫來文.收文日期
		sql.append(" ag_rt_docno, ");//--金庫來文.收文文號
		sql.append(" pay_amt_sum, ");//--繳還金額
		sql.append(" F_TRANSCHINESEDATE(frm_agsncorrdoc.doc_date) as toag_doc_date,");//--農金局.發文日期(有疑義個案.發文至金庫)
		sql.append(" frm_agsncorrdoc.docno as toag_docno, ");//--農金局.發文文號(有疑義個案.發文至金庫)
		sql.append(" decode(ag_flag,'0','核無不妥','1','尚有疑義','') as ag_flag ");//--核處情形
		sql.append(" from frm_snrtdoc "); 
		sql.append(" left join frm_agsncorrdoc on  frm_snrtdoc.ex_no=frm_agsncorrdoc.ex_no ");
		sql.append(" and frm_snrtdoc.bank_no = frm_agsncorrdoc.bank_no ");
		sql.append(" and frm_snrtdoc.def_seq = frm_agsncorrdoc.def_seq ");
		sql.append(" where frm_snrtdoc.bank_no=? ");
		sql.append(" and frm_snrtdoc.ex_no=?");
		sql.append(" and audit_id in ('A2','A3')");// --核處情形A2:繳還補貼息 A3:調整期限繳還溢領補貼息
		sql.append(" group by frm_snrtdoc.doc_date,frm_snrtdoc.docno,ag_rt_doc_date,ag_rt_docno,pay_amt_sum,frm_agsncorrdoc.doc_date,frm_agsncorrdoc.docno,ag_flag ");
		sql.append(" union ");
		//sql.append(" --有疑義發文至金庫及金庫更正函
		sql.append(" select "); 
		sql.append(" F_TRANSCHINESEDATE(frm_agsncorrdoc.doc_date) as toag_doc_date,");//--農金局.發文日期(有疑義個案.發文至金庫)
		sql.append(" frm_agsncorrdoc.docno as toag_docno, ");//--農金局.發文文號(有疑義個案.發文至金庫)
		sql.append(" F_TRANSCHINESEDATE(corr_doc_date) as corr_doc_date, ");//--金庫更正函.收文日期
		sql.append(" corr_docno, ");//--金庫更正函.收文文號
		sql.append(" over_amt, ");//--繳還(退還)金額
		sql.append(" F_TRANSCHINESEDATE(frm_refunddoc.doc_date) as doc_date, ");//--退還溢繳金額發文至農漁會.發文日期
		sql.append(" frm_refunddoc.docno, ");//--退還溢繳金額發文至農漁會.發文文號
		sql.append(" CASE         ");                                                               
		sql.append(" WHEN (case_status=0 and over_amt=0) THEN '核無不妥' ");
		sql.append(" WHEN (over_amt>0) THEN '少計需補繳'||over_amt||'元' ");
		sql.append(" WHEN (over_amt<0) THEN '退還溢繳'||abs(over_amt)||'元' ");
		sql.append(" ELSE '' ");
		sql.append(" END AS over_flag ");//--金庫更正函.回覆結果
		sql.append(" from  ");
		sql.append(" (select ex_no,bank_no,doc_date,docno,corr_doc_date,corr_docno,case_status,");
		//sql.append(" sum(re_pay_amt) - sum(over_amt) as over_amt ");
		sql.append(" sum(re_pay_amt),sum(over_amt),decode(sum(re_pay_amt),null,0,sum(re_pay_amt)) ");
		sql.append(" - decode(sum(over_amt),null,0,sum(over_amt)) as over_amt ");
		sql.append(" from ");
		sql.append(" ( ");
		sql.append(" select frm_agsncorrdoc.ex_no,frm_agsncorrdoc.bank_no,loan_name,loan_date,loan_item,loan_amt,doc_date,docno,corr_doc_date,corr_docno,agcorr_flag,re_pay_amt,frm_agsncorrdoc.over_amt,frm_exmaster.case_status");// --frm_agsncorrdoc.def_seq,def_case 
		sql.append(" from frm_agsncorrdoc  ");
		sql.append(" left join frm_exdef on  frm_agsncorrdoc.ex_no=frm_exdef.ex_no and frm_agsncorrdoc.bank_no = frm_exdef.bank_no and frm_agsncorrdoc.def_seq=frm_exdef.def_seq ");
		sql.append(" left join frm_exmaster on frm_agsncorrdoc.ex_no = frm_exmaster.ex_no and frm_agsncorrdoc.bank_no =frm_exmaster.bank_no ");
		sql.append(" where corr_docno is not null ");
		sql.append(" group by  frm_agsncorrdoc.ex_no,frm_agsncorrdoc.bank_no,loan_name,loan_date,loan_item,loan_amt,doc_date,docno,corr_doc_date,corr_docno,agcorr_flag,re_pay_amt,frm_agsncorrdoc.over_amt,frm_exmaster.case_status ");
		sql.append(" )group by ex_no,bank_no,doc_date,docno,corr_doc_date,corr_docno,case_status)frm_agsncorrdoc ");
		sql.append(" left join frm_refunddoc on frm_agsncorrdoc.ex_no=frm_refunddoc.ex_no and frm_agsncorrdoc.bank_no=frm_refunddoc.bank_no "); 
		sql.append(" where frm_agsncorrdoc.bank_no=? ");
		sql.append(" and frm_agsncorrdoc.ex_no= ? ");
		sql.append(" group by frm_agsncorrdoc.doc_date,frm_agsncorrdoc.docno,corr_docno,corr_doc_date,over_amt,case_status,frm_refunddoc.doc_date,frm_refunddoc.docno ");
		
		QUERY_SQL = sql.toString() ;
    }
    private static void getQueryData2 (Connection con , ResultSet rs ) {
    	try {
    		PreparedStatement  stmt = con.prepareStatement(QUERY_SQL) ; 
    		stmt.setString(1, TBANK); 
    		stmt.setString(2, EX_NO); 
    		stmt.setString(3, TBANK); 
    		stmt.setString(4, EX_NO);
    		rs = stmt.executeQuery() ;
    		DATA_LIST.clear(); 
    		while(rs.next() ) {
    			ReportBean bean = new ReportBean () ; 
    			bean.setDoc_date(rs.getString("doc_date")==null?"":rs.getString("doc_date"));
    			bean.setDocNo(rs.getString("docNo")==null?"":rs.getString("docNo"));
    			bean.setAg_rt_doc_date(rs.getString("ag_rt_doc_date")==null?"":rs.getString("ag_rt_doc_date"));
    			bean.setAg_rt_docno(rs.getString("ag_rt_docno")==null?"":rs.getString("ag_rt_docno"));
    			bean.setPay_amt_sum(rs.getString("pay_amt_sum")==null?"":rs.getString("pay_amt_sum"));
    			bean.setToag_doc_date(rs.getString("toag_doc_date")==null?"":rs.getString("toag_doc_date"));
    			bean.setToag_docno(rs.getString("toag_docno")==null?"":rs.getString("toag_docno"));
    			bean.setAg_flag(rs.getString("ag_flag")==null?"":rs.getString("ag_flag"));
    			DATA_LIST.add(bean) ;
    		}
    		rs.close(); 
    		stmt.close(); 
    	} catch(Exception e) {
    		e.printStackTrace(); 
    	} 
    }
    private static void getQueryData1 (Connection con , ResultSet rs ) {
    	
    	try {
    		PreparedStatement  stmt = con.prepareStatement(QUERY_SQL) ;
    		//System.out.println("FL016W sql:"+QUERY_SQL.toString()); 
    		//System.out.println("p1:"+TBANK); 
    		//System.out.println("p2:"+EX_NO);
    		stmt.setString(1, TBANK); 
    		stmt.setString(2, EX_NO); 
    		
    		rs = stmt.executeQuery() ;
    		DATA_LIST.clear(); 
    		while(rs.next() ) {
    			ReportBean bean = new ReportBean () ;
    			bean.setDoc_type(rs.getString("doc_type")==null?"":rs.getString("doc_type"));
    			bean.setDoc_type_name(rs.getString("doc_type_name")==null?"":rs.getString("doc_type_name"));
    			bean.setDoc_date(rs.getString("doc_date")==null?"":rs.getString("doc_date"));
    			bean.setDocNo(rs.getString("docno")==null?"":rs.getString("docno"));
    			bean.setBank_rt_doc_date(rs.getString("bank_rt_doc_date")==null?"":rs.getString("bank_rt_doc_date"));
    			bean.setBank_rt_docNo(rs.getString("bank_rt_docno")==null?"":rs.getString("bank_rt_docno"));
    			bean.setLoan_name(rs.getString("loan_name")==null?"":rs.getString("loan_name"));
    			bean.setLoan_item(rs.getString("loan_item")==null?"":rs.getString("loan_item"));
    			bean.setLoan_item_name(rs.getString("loan_item_name")==null?"":rs.getString("loan_item_name"));
    			bean.setDef_case(rs.getString("def_case")==null?"":rs.getString("def_case"));
    			bean.setAudit_case(rs.getString("audit_case")==null?"":rs.getString("audit_case"));
    			bean.setAudit_id_c2(rs.getString("audit_id_c2")==null?"":rs.getString("audit_id_c2"));
    			bean.setAudit_id_c1(rs.getString("audit_id_c1")==null?"":rs.getString("audit_id_c1"));
    			DATA_LIST.add(bean) ;
    		}
    		rs.close(); 
    		stmt.close(); 
    	} catch(Exception e) {
    		e.printStackTrace(); 
    	} 
    }
    private static HSSFWorkbook  getExcelSimple() {
		String openFile= "FL016W個別農漁會對其缺失改善處理情形.xls" ;
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
						ERROR_MSG += Utility.getProperties("reportDir") + "目錄新增失敗";
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
    		row = sheet.createRow((short)rowNo) ;
    		cell = row.createCell((short)cellNo) ;
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
			EX_TYPE_UI = Utility.getTrimString(h.get("ex_Type")) ;
			TBANK_UI = Utility.getTrimString(h.get("tbank_ui")) ;
			TBANK = Utility.getTrimString(h.get("tbank")) ;
			EX_NO = Utility.getTrimString(h.get("ex_no")) ; 
			UI_VAL = Utility.getTrimString(h.get("ui_value")) ;
			HSSFSheet sheet = null ;
		    HSSFWorkbook wb = null ;
		    HSSFRow row = null ; 
		    HSSFCell cell = null ;
		    HSSFCellStyle simpleStyle = null;
		    
		    int rowNo = 0 ;
		    System.out.println("1.報表名稱");
			//1.報表名稱
			FILE_NAME = "RptFL016W"+getYYYYMMDDHHMMSS()+".xls" ;
			//2.取得SQL
			System.out.println("2.取得SQL");
			getReportSQL1() ;
			//3.取得連線
			System.out.println("3.取得連線");
			con = (new RdbCommonDao("")).newConnection();
			ResultSet rs = null ;
			//4.取得資料
			System.out.println("4.取得資料");
			getQueryData1(con,rs) ;  
			//5.產生Excel
			
			//6.get 報表範本
			System.out.println("6.get報表範本") ;
			wb = getExcelSimple () ;
			System.out.println("7.設定報表格式") ;
			sheet = wb.getSheetAt(0);//讀取第一個工作表，宣告其為sheet 
    		HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
    		ps.setScale( ( short )100 ); //列印縮放百分比
			ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
			ps.setLandscape( true ); // 設定橫式
			
			//取得欄位style 
			HSSFCellStyle cellLeft1 = sheet.getRow((short)7).getCell((short)6).getCellStyle() ;
			simpleStyle = sheet.getRow((short)7).getCell((short)0).getCellStyle() ;
			//農漁會別:
			row = sheet.getRow((short)1) ; 
			cell = row.getCell((short)1) ;
			cell.setCellValue(TBANK_UI) ;
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			//查核類別 :
			
			if("FEB".equals(EX_TYPE_UI)){
				row = sheet.getRow((short)2) ;
				cell = row.getCell((short)1) ;
				cell.setCellValue("■金管會檢查報告");
				cell.setEncoding(HSSFCell.ENCODING_UTF_16); 
				cell = row.getCell((short)4) ; 
				cell.setCellValue(UI_VAL); 
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			}  else if ("AGRI".equals(EX_TYPE_UI)){
				row = sheet.getRow((short)3) ;
				cell = row.getCell((short)1) ;
				cell.setCellValue("■農業金庫查核");
				cell.setEncoding(HSSFCell.ENCODING_UTF_16); 
				cell = row.getCell((short)4) ; 
				cell.setCellValue(UI_VAL); 
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			} else {
				row = sheet.getRow((short)4) ;
				cell = row.getCell((short)1) ;
				cell.setCellValue("■農金局訪查");
				cell.setEncoding(HSSFCell.ENCODING_UTF_16); 
				cell = row.getCell((short)4) ; 
				cell.setCellValue(UI_VAL); 
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			}
			
			rowNo = 7 ;
			
			for(int i=0 ; i < DATA_LIST.size() ;i++) {
				ReportBean rb = DATA_LIST.get(i) ;
				if("A".equals(rb.getDoc_type())){
					row = sheet.getRow((short)rowNo) ;
					cell = row.getCell((short)0) ;
					cell.setCellValue(rb.getDoc_type_name()); 
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					//發文日期
					cell = row.getCell((short)1) ;
					cell.setCellValue(rb.getDoc_date()); 
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					//發文文號
					cell = row.getCell((short)2) ;
					cell.setCellValue(rb.getDocNo()); 
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					//收文日期
					cell = row.getCell((short)3) ;
					cell.setCellValue(rb.getBank_rt_doc_date()); 
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					//收文文號
					cell = row.getCell((short)4) ;
					cell.setCellValue(rb.getBank_rt_docNo()); 
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					//個案名稱
					cell = row.getCell((short)5) ;
					cell.setCellValue(rb.getLoan_item_name()+rb.getLoan_item_name()); 
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					//核處內容
					cell = row.getCell((short)6) ;
					cell.setCellValue(rb.getAudit_case()); 
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					rowNo++;
				}
			}
			int rowStart = rowNo ;
			int rowEnd = 0 ;
			boolean flg1 = false ;//列印裁處糾正
			boolean flg2 = false ;//列印裁處罰鍰
			boolean flg3 = false ;//裁處糾正 裁處罰鍰 資要
			int seqNo =  0 ;
			String tmpDocNo = "" ;//發文文號 temp 
			for(int i=0 ; i < DATA_LIST.size() ;i++) {
				ReportBean rb = DATA_LIST.get(i) ;
				if(!"A".equals(rb.getDoc_type())){
					row = sheet.createRow((short)rowNo) ;
					if("".equals(tmpDocNo)) {
						if("裁處糾正".equals(rb.getAudit_id_c2())) {
							System.out.println(" 需 列印裁處糾正");
							flg1 = true ;
							flg3 = true ;
						}
						if("裁處罰鍰".equals(rb.getAudit_id_c1())) {
							System.out.println(" 需 列印裁處罰鍰");
							flg2 = true ;
							flg3 = true ;
						}
						if(!flg3) {
							rowEnd = rowNo ;
							setExcelCellValue(sheet ,row,cell,rowNo,(short)0,transChtNum(++seqNo), simpleStyle);
							//發文日期
							setExcelCellValue(sheet ,row,cell,rowNo,(short)1,rb.getDoc_date(), simpleStyle);
							//發文文號
							setExcelCellValue(sheet ,row,cell,rowNo,(short)2,rb.getDocNo(), simpleStyle);
							//收文日期
							setExcelCellValue(sheet ,row,cell,rowNo,(short)3,rb.getBank_rt_doc_date(), simpleStyle);
							//收文文號
							setExcelCellValue(sheet ,row,cell,rowNo,(short)4,rb.getBank_rt_docNo(), simpleStyle);
							//個案名稱
							setExcelCellValue(sheet ,row,cell,rowNo,(short)5,rb.getLoan_name()+rb.getLoan_item_name(), simpleStyle);
							//核處內容
							setExcelCellValue(sheet ,row,cell,rowNo++,(short)6,rb.getAudit_case(), cellLeft1);
						} else {
							++seqNo;
						}
						flg3 = false  ; //清除註記
					} else if (tmpDocNo.equals(rb.getDocNo())) {
						if("裁處糾正".equals(rb.getAudit_id_c2())) {
							System.out.println(" 需 列印裁處糾正");
							flg1 = true ;
							flg3 = true ;
						}
						if("裁處罰鍰".equals(rb.getAudit_id_c1())) {
							System.out.println(" 需 列印裁處罰鍰");
							flg2 = true ;
							flg3 = true ;
						}
						
						System.out.println("相同文號列印資訊");
						if(!flg3) {
							rowEnd = rowNo ;
							setExcelCellValue(sheet ,row,cell,rowNo,(short)0,transChtNum(seqNo), simpleStyle);
							//發文日期
							setExcelCellValue(sheet ,row,cell,rowNo,(short)1,rb.getDoc_date(), simpleStyle);
							//發文文號
							setExcelCellValue(sheet ,row,cell,rowNo,(short)2,rb.getDocNo(), simpleStyle);
							//收文日期
							setExcelCellValue(sheet ,row,cell,rowNo,(short)3,rb.getBank_rt_doc_date(), simpleStyle);
							//收文文號
							setExcelCellValue(sheet ,row,cell,rowNo,(short)4,rb.getBank_rt_docNo(), simpleStyle);
							//個案名稱
							setExcelCellValue(sheet ,row,cell,rowNo,(short)5,rb.getLoan_name()+rb.getLoan_item_name(), simpleStyle);
							//核處內容
							setExcelCellValue(sheet ,row,cell,rowNo++,(short)6,rb.getAudit_case(), cellLeft1);
						} else {
							--rowNo;//該行不列印,不跳行
						}
						flg3 = false  ; //清除註記
					} else {
						System.out.println("不同文號列印");
						if(flg1) {
							rowEnd = rowNo ;
							setExcelCellValue(sheet ,row,cell,rowNo,(short)0,"", simpleStyle);
							setExcelCellValue(sheet ,row,cell,rowNo,(short)1,"", simpleStyle);
							setExcelCellValue(sheet ,row,cell,rowNo,(short)2,"", simpleStyle);
							setExcelCellValue(sheet ,row,cell,rowNo,(short)3,"", simpleStyle);
							setExcelCellValue(sheet ,row,cell,rowNo,(short)4,"", simpleStyle);
							setExcelCellValue(sheet ,row,cell,rowNo,(short)5,"", simpleStyle);
							setExcelCellValue(sheet ,row,cell,rowNo++,(short)6,"裁處糾正", cellLeft1);
							flg1 = false ;
						}
						if(flg2) {
							rowEnd = rowNo ;
							setExcelCellValue(sheet ,row,cell,rowNo,(short)0,"", simpleStyle);
							setExcelCellValue(sheet ,row,cell,rowNo,(short)1,"", simpleStyle);
							setExcelCellValue(sheet ,row,cell,rowNo,(short)2,"", simpleStyle);
							setExcelCellValue(sheet ,row,cell,rowNo,(short)3,"", simpleStyle);
							setExcelCellValue(sheet ,row,cell,rowNo,(short)4,"", simpleStyle);
							setExcelCellValue(sheet ,row,cell,rowNo,(short)5,"", simpleStyle);
							setExcelCellValue(sheet ,row,cell,rowNo,(short)6,"裁處罰鍰", cellLeft1);
							flg2 = false ;
						}
						
						sheet.addMergedRegion(new Region((short)(rowStart), (short) 0, rowEnd, (short) 0));
						sheet.addMergedRegion(new Region((short)(rowStart), (short) 1, rowEnd, (short) 1));
						sheet.addMergedRegion(new Region((short)(rowStart), (short) 2, rowEnd, (short) 2));
						++rowNo ;
						rowStart = rowNo ;
						rowEnd = rowNo ;
						setExcelCellValue(sheet ,row,cell,rowNo,(short)0,transChtNum(++seqNo), simpleStyle);
						//發文日期
						setExcelCellValue(sheet ,row,cell,rowNo,(short)1,rb.getDoc_date(), simpleStyle);
						//發文文號
						setExcelCellValue(sheet ,row,cell,rowNo,(short)2,rb.getDocNo(), simpleStyle);
						//收文日期
						setExcelCellValue(sheet ,row,cell,rowNo,(short)3,rb.getBank_rt_doc_date(), simpleStyle);
						//收文文號
						setExcelCellValue(sheet ,row,cell,rowNo,(short)4,rb.getBank_rt_docNo(), simpleStyle);
						//個案名稱
						setExcelCellValue(sheet ,row,cell,rowNo,(short)5,rb.getLoan_name()+rb.getLoan_item_name(), simpleStyle);
						//核處內容
						setExcelCellValue(sheet ,row,cell,rowNo++,(short)6,rb.getAudit_case(), cellLeft1);
					}
					
					tmpDocNo = rb.getDocNo() ;
				}
			}
			System.out.println("判斷 最後一筆缺失改善情形欄位合併 ") ;
			if(flg1) {
				rowEnd = rowNo ;
				setExcelCellValue(sheet ,row,cell,rowNo,(short)0,"", simpleStyle);
				setExcelCellValue(sheet ,row,cell,rowNo,(short)1,"", simpleStyle);
				setExcelCellValue(sheet ,row,cell,rowNo,(short)2,"", simpleStyle);
				setExcelCellValue(sheet ,row,cell,rowNo,(short)3,"", simpleStyle);
				setExcelCellValue(sheet ,row,cell,rowNo,(short)4,"", simpleStyle);
				setExcelCellValue(sheet ,row,cell,rowNo,(short)5,"", simpleStyle);
				setExcelCellValue(sheet ,row,cell,rowNo++,(short)6,"裁處糾正", cellLeft1);
				flg1 = false ;
			}
			if(flg2) {
				rowEnd = rowNo ;
				setExcelCellValue(sheet ,row,cell,rowNo,(short)0,"", simpleStyle);
				setExcelCellValue(sheet ,row,cell,rowNo,(short)1,"", simpleStyle);
				setExcelCellValue(sheet ,row,cell,rowNo,(short)2,"", simpleStyle);
				setExcelCellValue(sheet ,row,cell,rowNo,(short)3,"", simpleStyle);
				setExcelCellValue(sheet ,row,cell,rowNo,(short)4,"", simpleStyle);
				setExcelCellValue(sheet ,row,cell,rowNo,(short)5,"", simpleStyle);
				setExcelCellValue(sheet ,row,cell,rowNo,(short)6,"裁處罰鍰", cellLeft1);
				flg2 = false ;
			}
			
			System.out.println("判斷 是否最後一次需合併儲存格 ") ;
			if(rowStart != rowEnd  && seqNo >0 ) {
				System.out.println("需要合併");
				System.out.println("rowStart:"+rowStart);
				System.out.println("rowEnd:"+(rowNo));
				sheet.addMergedRegion(new Region((short)(rowStart), (short) 0, rowEnd, (short) 0));
				sheet.addMergedRegion(new Region((short)(rowStart), (short) 1, rowEnd, (short) 1));
				sheet.addMergedRegion(new Region((short)(rowStart), (short) 2, rowEnd, (short) 2));
			}
			
						
			//繳還補貼息情形報表查詢 
			++rowNo ;
			++rowNo ;
			HSSFCellStyle subTitle = sheet.getRow((short)6).getCell((short)1).getCellStyle() ;
			HSSFCellStyle subTitle2 = sheet.getRow((short)6).getCell((short)0).getCellStyle() ;
			//HSSFCellStyle subTitle12Font = getColumn_LeftStyle(wb);
				        
			setExcelCellValue(sheet ,row,cell,rowNo,(short)0,"繳還補貼息情形", subTitle2);
			setExcelCellValue(sheet ,row,cell,rowNo,(short)1,"農業金庫來文", subTitle);
			setExcelCellValue(sheet ,row,cell,rowNo,(short)2,"", subTitle);
			setExcelCellValue(sheet ,row,cell,rowNo,(short)3,"", subTitle);
			setExcelCellValue(sheet ,row,cell,rowNo,(short)4,"農金局發文", subTitle);
			setExcelCellValue(sheet ,row,cell,rowNo,(short)5,"", subTitle);
			setExcelCellValue(sheet ,row,cell,rowNo,(short)6,"核處情形", subTitle);
			sheet.addMergedRegion(new Region((short)(rowNo), (short) 1, rowNo, (short) 3));
			sheet.addMergedRegion(new Region((short)(rowNo), (short) 4, rowNo, (short) 5));
			
			rowNo++;
			setExcelCellValue(sheet ,row,cell,rowNo,(short)0,"", subTitle);
			setExcelCellValue(sheet ,row,cell,rowNo,(short)1,"收文日期", subTitle);
			setExcelCellValue(sheet ,row,cell,rowNo,(short)2,"收文文號", subTitle);
			setExcelCellValue(sheet ,row,cell,rowNo,(short)3,"繳還(退還)金額", subTitle);
			setExcelCellValue(sheet ,row,cell,rowNo,(short)4,"發文日期", subTitle);
			setExcelCellValue(sheet ,row,cell,rowNo,(short)5,"發文文號", subTitle);
			setExcelCellValue(sheet ,row,cell,rowNo,(short)6,"", subTitle);
			sheet.addMergedRegion(new Region((short)(rowNo)-1, (short) 0, rowNo, (short) 0));
			sheet.addMergedRegion(new Region((short)(rowNo)-1, (short) 6, rowNo, (short) 6));
			
			//end
			getReportSQL2 () ;
			getQueryData2 (con , rs ) ;
			
			for(int i=0 ; i < DATA_LIST.size() ;i++) { 
				ReportBean rb = DATA_LIST.get(i) ;
				rowNo++; 
				setExcelCellValue(sheet ,row,cell,rowNo,(short)0,transChtNum(i+1), simpleStyle);
				setExcelCellValue(sheet ,row,cell,rowNo,(short)1,rb.getAg_rt_doc_date(), simpleStyle);
				setExcelCellValue(sheet ,row,cell,rowNo,(short)2,rb.getAg_rt_docno(), simpleStyle);
				setExcelCellValue(sheet ,row,cell,rowNo,(short)3,NumberFormat(rb.getPay_amt_sum()), reportUtil.getRightStyleF12(wb));
				setExcelCellValue(sheet ,row,cell,rowNo,(short)4,rb.getToag_doc_date(), simpleStyle);
				setExcelCellValue(sheet ,row,cell,rowNo,(short)5,rb.getToag_docno(), simpleStyle);
				setExcelCellValue(sheet ,row,cell,rowNo,(short)6,rb.getAg_flag(), simpleStyle);
			}
			outPutExcel(wb ,sheet) ; 
			
			
		}catch(Exception e){
			System.out.println("createRpt Error1:"+e);
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

	private static String transChtNum(int k) {
		String tmp = "";
		String t= String.valueOf(k) ;
		String c= "十" ; 
		String [] st = {"零","一","二","三","四","五","六","七","八","九"}; 
		if(t.length()==1) {
			tmp = st[k] ;
		} else if(t.length()==2){
			if(t.substring(0,1).equals("1")) {
				if(t.substring(1,2).equals("0")) {
					tmp = c ;
				} else {
					tmp = c+st[Integer.parseInt(t.substring(1,2))] ;
				}
			} else {
				tmp = st[Integer.parseInt(t.substring(0,1))] +c +st[Integer.parseInt(t.substring(1,2))];
			}
		}
		return "第"+tmp+"次" ;
	}
	
	
	
	
	static class ReportBean {
    	private String doc_type = "" ;
    	private String doc_type_name = "";
    	private String doc_date = "";
    	private String docNo = "";
    	private String bank_rt_doc_date = "" ;
    	private String bank_rt_docNo = "";
    	private String loan_name="";
    	private String loan_item = "";
    	private String loan_item_name = "";
    	private String def_case = "";
    	private String audit_case = "" ;
    	private String audit_id_c2 = "";
    	private String audit_id_c1 = "";
    	private String ag_rt_doc_date = "";
    	private String ag_rt_docno = "";
    	private String pay_amt_sum = "";
    	private String toag_doc_date = "";
    	private String toag_docno  = "";
    	private String ag_flag = "" ;
    	
		public String getDoc_type() {
			return doc_type;
		}
		public void setDoc_type(String doc_type) {
			this.doc_type = doc_type;
		}
		public String getDoc_type_name() {
			return doc_type_name;
		}
		public void setDoc_type_name(String doc_type_name) {
			this.doc_type_name = doc_type_name;
		}
		public String getDoc_date() {
			return doc_date;
		}
		public void setDoc_date(String doc_date) {
			this.doc_date = doc_date;
		}
		public String getDocNo() {
			return docNo;
		}
		public void setDocNo(String docNo) {
			this.docNo = docNo;
		}
		public String getBank_rt_doc_date() {
			return bank_rt_doc_date;
		}
		public void setBank_rt_doc_date(String bank_rt_doc_date) {
			this.bank_rt_doc_date = bank_rt_doc_date;
		}
		public String getBank_rt_docNo() {
			return bank_rt_docNo;
		}
		public void setBank_rt_docNo(String bank_rt_docNo) {
			this.bank_rt_docNo = bank_rt_docNo;
		}
		public String getLoan_name() {
			return loan_name;
		}
		public void setLoan_name(String loan_name) {
			this.loan_name = loan_name;
		}
		public String getLoan_item_name() {
			return loan_item_name;
		}
		public void setLoan_item_name(String loan_item_name) {
			this.loan_item_name = loan_item_name;
		}
		public String getDef_case() {
			return def_case;
		}
		public void setDef_case(String def_case) {
			this.def_case = def_case;
		}
		public String getAudit_case() {
			return audit_case;
		}
		public void setAudit_case(String audit_case) {
			this.audit_case = audit_case;
		}
		public String getAudit_id_c2() {
			return audit_id_c2;
		}
		public void setAudit_id_c2(String audit_id_c2) {
			this.audit_id_c2 = audit_id_c2;
		}
		public String getAudit_id_c1() {
			return audit_id_c1;
		}
		public void setAudit_id_c1(String audit_id_c1) {
			this.audit_id_c1 = audit_id_c1;
		}
		public String getLoan_item() {
			return loan_item;
		}
		public void setLoan_item(String loan_item) {
			this.loan_item = loan_item;
		}
		public String getAg_rt_doc_date() {
			return ag_rt_doc_date;
		}
		public void setAg_rt_doc_date(String ag_rt_doc_date) {
			this.ag_rt_doc_date = ag_rt_doc_date;
		}
		public String getAg_rt_docno() {
			return ag_rt_docno;
		}
		public void setAg_rt_docno(String ag_rt_docno) {
			this.ag_rt_docno = ag_rt_docno;
		}
		public String getPay_amt_sum() {
			return pay_amt_sum;
		}
		public void setPay_amt_sum(String pay_amt_sum) {
			this.pay_amt_sum = pay_amt_sum;
		}
		public String getToag_doc_date() {
			return toag_doc_date;
		}
		public void setToag_doc_date(String toag_doc_date) {
			this.toag_doc_date = toag_doc_date;
		}
		public String getToag_docno() {
			return toag_docno;
		}
		public void setToag_docno(String toag_docno) {
			this.toag_docno = toag_docno;
		}
		public String getAg_flag() {
			return ag_flag;
		}
		public void setAg_flag(String ag_flag) {
			this.ag_flag = ag_flag;
		}
    	
    			
		
		
    	
    }
	public static HSSFCellStyle getColumn_LeftStyle(HSSFWorkbook wb){
		HSSFFont f = wb.createFont();
		//set font 1 to 12 point type
		f.setFontHeightInPoints((short) 12);
		//make it blue
		//f.setColor( (short)0xc );
		//make it bold
		//arial is the default font
		//f.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
    
		HSSFCellStyle columnStyle = wb.createCellStyle(); 
		columnStyle = HssfStyle.setStyle( columnStyle, f,
                                 new String[] {
                                 "BORDER", "PHL", "PVC", "F12",
                                 "WRAP"} );
		return columnStyle;
}
	private static String NumberFormat(String m ) {
		String t = "" ;
		if(!"".equals(m)) {
			DecimalFormat  fat = new DecimalFormat ("$###,###") ;
			t = fat.format(Double.parseDouble(m)) ; 
		}
		DecimalFormat  fat = new DecimalFormat ("$###,###") ;
		return t ;
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
