
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

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFPrintSetup;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import com.tradevan.util.Utility;
import com.tradevan.util.dao.RdbCommonDao;

/***
 * 政策性農業專案貸款尚未繳還補貼息明細表
 * @author 2808
 *
 */
public class RptFL013W {
	/**Report執行SQL*/
	private static String QUERY_SQL =  "" ;
	/**報表名稱*/
	private static String FILE_NAME = "" ;
	/**錯誤訊息*/
	private static String ERROR_MSG = "" ;
	
	private static Connection con = null ;
	
	private static ArrayList <ReportBean> DATA_LIST = new ArrayList <ReportBean>(); 
	
	//private static HSSFRow row = null; //宣告一列
	//private static HSSFCell cell = null; //宣告一個儲存格
	
	/**取得報表查詢SQL*/
    private static void   getReportSQL() {
    	StringBuffer sql = new StringBuffer() ;
    	sql.append(" select frm_snrtdoc.bank_no,");//--農漁會別.機構代號
		sql.append(" bank_name,");//--農漁會別.機構名稱
		sql.append(" decode(frm_exmaster.ex_type,'FEB','金管會檢查報告','AGRI','農業金庫查核','BOAF','農金局訪查','') as ex_type_name, ");//--查核類別
		sql.append(" frm_snrtdoc.ex_no, ");//--查核報告編號
		sql.append(" decode(ex_type,'FEB',frm_snrtdoc.ex_no,'AGRI',substr(frm_snrtdoc.ex_no,0,3)||'年第'||substr(frm_snrtdoc.ex_no,4,2)||'季','BOAF',F_TRANSCHINESEDATE(to_date(frm_snrtdoc.ex_no,'yyyymmdd')),'') as ex_no_list,");// --報表需顯示的檢查報告編號或查核季別或訪查日期
		sql.append(" F_TRANSCHINESEDATE(frm_snrtdoc.doc_date) as doc_date,");//--農金局發文.日期
		sql.append(" frm_snrtdoc.docno,");//--農金局發文.文號 
		sql.append(" frm_exdef.case_count,");//--農金局發文.案件數
		sql.append(" frm_exdef.loan_amt_sum,");//--農金局發文.貸款金額加總
		sql.append(" wlx01.telno,");//--電話
		sql.append(" F_COMBWLX01_MNAME(frm_snrtdoc.bank_no,'4') as mname,");// --信用部主任
		sql.append(" F_COMBWLX01_AuditNAME(frm_snrtdoc.bank_no) as audit_name ");// --稽核
		sql.append(" from frm_snrtdoc ");
		sql.append(" left join (select * from bn01 where m_year=100)bn01 on frm_snrtdoc.bank_no=bn01.bank_no ");
		sql.append(" left join frm_exmaster on frm_snrtdoc.ex_no=frm_exmaster.ex_no  and ");
		sql.append(" frm_snrtdoc.bank_no=frm_exmaster.bank_no  ");
		sql.append(" left join (select * from wlx01 where m_year=100)wlx01 on frm_snrtdoc.bank_no=wlx01.bank_no ");
		sql.append(" left join ( ");
		sql.append(" select  ex_no,bank_no,doc_date,docno,count(*) as case_count,sum(loan_amt) as loan_amt_sum ");
		sql.append(" from ( "); 
		sql.append(" select frm_exdef.ex_no,frm_exdef.bank_no,doc_date,docno,loan_name,loan_date,loan_amt ");
		sql.append(" from frm_exdef "); 
		sql.append(" left join frm_snrtdoc on frm_exdef.ex_no = frm_snrtdoc.ex_no and frm_exdef.bank_no=frm_snrtdoc.bank_no and frm_exdef.def_seq=frm_snrtdoc.def_seq ");
		sql.append(" where ex_kind='C'  ");//--個案
		sql.append(" and frm_snrtdoc.audit_id in ('A2','A3') ");//--需繳還補貼息
		sql.append(" and frm_snrtdoc.pay_date is null ");//--尚未繳還補貼息
		sql.append(" group by frm_exdef.ex_no,frm_exdef.bank_no,doc_date,docno,loan_name,loan_date,loan_amt ");
		sql.append(" order by ex_no,bank_no,doc_date,docno)a ");
		sql.append(" group by ex_no,bank_no,doc_date,docno ");
		sql.append(" )frm_exdef on frm_snrtdoc.ex_no=frm_exdef.ex_no and frm_snrtdoc.bank_no=frm_exdef.bank_no and frm_snrtdoc.doc_date = frm_exdef.doc_date and frm_snrtdoc.docno=frm_exdef.docno ");
		sql.append(" where doc_type='B' ");//--核處
		sql.append(" and audit_id in ('A2','A3') ");//--需繳還補貼息
		sql.append(" and pay_date is null ");//--尚未繳還補貼息
		sql.append(" group by frm_snrtdoc.bank_no,bank_name,frm_exmaster.ex_type,");
		sql.append(" frm_snrtdoc.ex_no,frm_snrtdoc.doc_date,frm_snrtdoc.docno,");
		sql.append(" frm_exdef.case_count,frm_exdef.loan_amt_sum,wlx01.telno,");
		sql.append(" F_COMBWLX01_MNAME(frm_snrtdoc.bank_no,'4'),F_COMBWLX01_AuditNAME(frm_snrtdoc.bank_no) ");
		sql.append(" order by frm_snrtdoc.bank_no asc,ex_no asc,doc_date desc ");


		QUERY_SQL = sql.toString() ;
    }
    
    
    private static void getQueryData (Connection con , ResultSet rs ) {
    	
    	try {
    		PreparedStatement  stmt = con.prepareStatement(QUERY_SQL) ;
    		rs = stmt.executeQuery() ;
    		DATA_LIST.clear(); 
    		while(rs.next() ) {
    			ReportBean bean = new ReportBean () ;
    			bean.setBank_no(rs.getString("bank_no")==null?"":rs.getString("bank_no"));
    			bean.setBank_nm(rs.getString("bank_name")==null?"":rs.getString("bank_name"));
    			bean.setEx_type_name(rs.getString("ex_type_name")==null?"":rs.getString("ex_type_name"));
    			bean.setEx_no(rs.getString("ex_no")==null?"":rs.getString("ex_no"));
    			bean.setEx_no_list(rs.getString("ex_no_list")==null?"":rs.getString("ex_no_list"));
    			bean.setDoc_date(rs.getString("doc_date")==null?"":rs.getString("doc_date"));
    			bean.setDocNo(rs.getString("docno")==null?"":rs.getString("docno"));
    			bean.setCase_count(rs.getString("case_count")==null?"":rs.getString("case_count"));
    			bean.setLoan_amt_sum(rs.getString("loan_amt_sum")==null?"":rs.getString("loan_amt_sum"));
    			bean.setTelNo(rs.getString("telno")==null?"":rs.getString("telno"));
    			bean.setmName(rs.getString("mname")==null?"":rs.getString("mname"));
    			bean.setAudit_name(rs.getString("audit_name")==null?"":rs.getString("audit_name"));
    			DATA_LIST.add(bean) ;
    		}
    		rs.close(); 
    		stmt.close(); 
    	} catch(Exception e) {
    		e.printStackTrace(); 
    	} 
    }
    
    private static HSSFWorkbook  getExcelSimple() {
		String openFile= "FL013W政策性農業專案貸款尚未繳還補貼息明細表.xls" ;
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
    		row = sheet.createRow(rowNo) ;
    		cell = row.createCell(cellNo) ;
    		cell.setCellStyle(sty); 
    		cell.setCellValue(value); 
    		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
    	}catch(Exception e) {
    		e.printStackTrace();
    	}
    }
	public static String createRpt(){
		try {
			HSSFSheet sheet = null ;
		    HSSFWorkbook wb = null ;
		    HSSFRow row = null ; 
		    HSSFCell cell = null ;
		    HSSFCellStyle simpleStyle = null;
		    //金額靠右
		    HSSFCellStyle simpleStyleRight = null;
		    
		    int rowNo = 0 ;
		    System.out.println("1.報表名稱");
			//1.報表名稱
			FILE_NAME = "RptFL013W"+getYYYYMMDDHHMMSS()+".xls" ;
			//2.取得SQL
			System.out.println("2.取得SQL");
			getReportSQL() ;
			//3.取得連線
			System.out.println("3.取得連線");
			con = (new RdbCommonDao("")).newConnection();//DBManager.newQryConnection(). ;
			ResultSet rs = null ;
			//4.取得資料
			System.out.println("4.取得資料");
			getQueryData(con,rs) ;  
			//5.產生Excel
			
			//6.get 報表範本
			System.out.println("6.get報表範本") ;
			wb = getExcelSimple () ;
			System.out.println("7.設定報表格式") ;
			sheet = wb.getSheetAt(0);//讀取第一個工作表，宣告其為sheet 
    		HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
    		ps.setScale( ( short )90 ); //列印縮放百分比
			ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
			ps.setLandscape( true ); // 設定橫式
			//set Query Date
			String queryDate = "查詢日期："+Utility.getCHTYYMMDD("yy").concat("年").concat(Utility.getCHTYYMMDD("mm")).concat("月").concat(Utility.getCHTYYMMDD("dd")).concat("日") ;
			row  = sheet.getRow(1) ;
			cell = row.getCell((short)0) ;
			cell.setCellValue(queryDate); 
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			//取得欄位style 
			row = sheet.getRow(4) ;
			cell = row.getCell((short)0) ;
			simpleStyle = cell.getCellStyle() ;
			simpleStyleRight = row.getCell((short)5).getCellStyle() ;
			rowNo = 5 ;
			for(int i=0 ; i < DATA_LIST.size() ;i++) {
				ReportBean rb = DATA_LIST.get(i) ;
				if(i==0) {
					cell = row.getCell((short)0) ; 
					cell.setCellValue(rb.getBank_no().concat(rb.getBank_nm()));
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell = row.getCell((short)1) ;
					cell.setCellValue(rb.getEx_type_name());
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell = row.getCell((short)2) ;
					cell.setCellValue(rb.getEx_no_list());
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell = row.getCell((short)3) ;
					cell.setCellValue(rb.getDoc_date());
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell = row.getCell((short)4) ;
					cell.setCellValue(rb.getDocNo());
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell = row.getCell((short)5) ;
					cell.setCellValue(rb.getCase_count());
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell = row.getCell((short)6) ;
					cell.setCellValue(rb.getLoan_amt_sum());
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell = row.getCell((short)7) ;
					cell.setCellValue(rb.getTelNo());
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell = row.getCell((short)8) ;
					cell.setCellValue(rb.getmName());
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell = row.getCell((short)9) ;
					cell.setCellValue(rb.getAudit_name());
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				} else {
					//row = sheet.createRow(++rowNo);
					setExcelCellValue(sheet ,row,cell,rowNo,(short)0,rb.getBank_no().concat(rb.getBank_nm()),simpleStyle) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)1,rb.getEx_type_name(),simpleStyle) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)2,rb.getEx_no_list(),simpleStyle) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)3,rb.getDoc_date(),simpleStyle) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)4,rb.getDocNo(),simpleStyle) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)5,rb.getCase_count(),simpleStyle) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)6,rb.getLoan_amt_sum(),simpleStyleRight) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)7,rb.getTelNo(),simpleStyle) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)8,rb.getmName(),simpleStyle) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)9,rb.getAudit_name(),simpleStyle) ;
					rowNo++ ;
				}
			}
			outPutExcel(wb ,sheet) ;
		}catch(Exception e){
			System.out.println("createRpt Error:"+e+e.getMessage());
		}finally {
			try {
				con.commit(); 
				con.close();
			}catch(Exception e){
				e.printStackTrace();
			}
		}
		return FILE_NAME ;
	}

	
	
	private static String NumberFormat(String m ) {
		DecimalFormat  fat = new DecimalFormat ("$###,###") ;
		return fat.format(Double.parseDouble(m)) ;
	}
	static class ReportBean {
    	private String bank_no = "" ;
    	private String bank_nm = "";
    	private String ex_type_name ;
    	private String ex_no  = "" ;
    	private String ex_no_list = "";
    	private String doc_date = "";
    	private String docNo = "";
    	private String case_count="";
    	private String loan_amt_sum ="";
    	private String telNo = "" ;
    	private String mName = "";
    	private String audit_name = "";
    	
		public String getBank_no() {
			return bank_no;
		}
		public void setBank_no(String bank_no) {
			this.bank_no = bank_no;
		}
		public String getBank_nm() {
			return bank_nm;
		}
		public void setBank_nm(String bank_nm) {
			this.bank_nm = bank_nm;
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
		public String getCase_count() {
			return case_count;
		}
		public void setCase_count(String case_count) {
			this.case_count = case_count;
		}
		public String getLoan_amt_sum() {
			if("".equals(loan_amt_sum)) {
				loan_amt_sum = "0" ;
			} else {
				loan_amt_sum = NumberFormat(loan_amt_sum) ;
			}
			return loan_amt_sum;
		}
		public void setLoan_amt_sum(String loan_amt_sum) {
			this.loan_amt_sum = loan_amt_sum;
		}
		public String getTelNo() {
			return telNo;
		}
		public void setTelNo(String telNo) {
			this.telNo = telNo;
		}
		public String getmName() {
			return mName;
		}
		public void setmName(String mName) {
			this.mName = mName;
		}
		public String getAudit_name() {
			return audit_name;
		}
		public void setAudit_name(String audit_name) {
			this.audit_name = audit_name;
		}
		
    	
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
