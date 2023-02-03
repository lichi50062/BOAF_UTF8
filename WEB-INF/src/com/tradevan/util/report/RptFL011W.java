package com.tradevan.util.report;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.ArrayList;
import java.util.Calendar;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFooter;
import org.apache.poi.hssf.usermodel.HSSFPrintSetup;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import com.jcraft.jsch.Logger;
import com.tradevan.util.Utility;
import com.tradevan.util.dao.RdbCommonDao;

public class RptFL011W {
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
		sql.append(" decode(frm_exmaster.ex_type,'FEB','金管會檢查報告','AGRI','農業金庫查核','BOAF','農金局訪查','') as ex_type_name,");//--查核類別
		sql.append(" frm_snrtdoc.ex_no, ");//--查核報告編號
		// --報表需顯示的檢查報告編號或查核季別或訪查日期
		sql.append(" decode(ex_type,'FEB',frm_snrtdoc.ex_no,'AGRI',substr(frm_snrtdoc.ex_no,0,3)||'年第'||substr(frm_snrtdoc.ex_no,4,2)||'季','BOAF',F_TRANSCHINESEDATE(to_date(frm_snrtdoc.ex_no,'yyyymmdd')),'') as ex_no_list,");
		sql.append(" F_TRANSCHINESEDATE(doc_date) as doc_date,");//--農金局發文.日期
		sql.append(" docno,");//--農金局發文.文號 
		sql.append(" doc_type,");//--發文性質代碼 A:陳述意見/B:核處/C:結案
		sql.append(" cmuse_name as doc_type_name, ");//--農金局發文.發文性質名稱
		sql.append(" F_TRANSCHINESEDATE(limitdate) as limitdate,");//--應回覆日期
		sql.append(" wlx01.telno, ");//--電話
		sql.append(" F_COMBWLX01_MNAME(frm_snrtdoc.bank_no,'4') as mname, ");//--信用部主任
		sql.append(" F_COMBWLX01_AuditNAME(frm_snrtdoc.bank_no) as audit_name");// --稽核
		sql.append(" from frm_snrtdoc ");
		sql.append(" left join (select * from bn01 where m_year=100)bn01 on frm_snrtdoc.bank_no=bn01.bank_no ");
		sql.append(" left join (select * from cdshareno where cmuse_div='048')cdshareno on frm_snrtdoc.doc_type=cdshareno.cmuse_id "); 
		sql.append(" left join frm_exmaster on frm_snrtdoc.ex_no=frm_exmaster.ex_no and ");
		sql.append(" frm_snrtdoc.bank_no=frm_exmaster.bank_no") ;
		sql.append(" left join (select * from wlx01 where m_year=100)wlx01 on frm_snrtdoc.bank_no=wlx01.bank_no ");
		sql.append(" where doc_type='B'");//--核處 
		sql.append(" and audit_id in ('A1','B1') ");//--限期改善具報
		sql.append(" and bank_rt_docno is null ");
		sql.append(" and sysdate > limitdate ");//--超過限期函報日
		sql.append(" group by frm_snrtdoc.bank_no,bank_name,frm_exmaster.ex_type,frm_snrtdoc.ex_no,frm_snrtdoc.ex_no,doc_date,docno,doc_type,cmuse_name,limitdate,wlx01.telno,frm_snrtdoc.bank_no,frm_snrtdoc.bank_no ");
		sql.append(" order by frm_snrtdoc.bank_no asc,ex_no asc,doc_date desc ");
		QUERY_SQL = sql.toString() ;
    }
    
    private static HSSFWorkbook  getExcelSimple() {
		String openFile= "FL011W逾期未函報缺失改善情形農漁會明細查詢表.xls" ;
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
    
    
    private static void getQueryData (Connection con , ResultSet rs ) {
    	
    	try {
    		PreparedStatement  stmt = con.prepareStatement(QUERY_SQL) ;
    		rs = stmt.executeQuery() ;
    		DATA_LIST.clear(); 
    		while(rs.next() ) {
    			ReportBean bean = new ReportBean () ;
    			bean.setBank_no(rs.getString("bank_no")==null?"":rs.getString("bank_no"));
    			bean.setBank_name(rs.getString("bank_name")==null?"":rs.getString("bank_name"));
    			bean.setEx_type_name(rs.getString("ex_type_name")==null?"":rs.getString("ex_type_name"));
    			bean.setEx_no(rs.getString("ex_no")==null?"":rs.getString("ex_no"));
    			bean.setEx_no_list(rs.getString("ex_no_list")==null?"":rs.getString("ex_no_list"));
    			bean.setDoc_date(rs.getString("doc_date")==null?"":rs.getString("doc_date"));
    			bean.setDocNo(rs.getString("docNo")==null?"":rs.getString("docNo"));
    			bean.setDoc_type(rs.getString("doc_type")==null?"":rs.getString("doc_type"));
    			bean.setDoc_type_name(rs.getString("doc_type_name")==null?"":rs.getString("doc_type_name"));
    			bean.setLimitDate(rs.getString("limitDate")==null?"":rs.getString("limitDate"));
    			bean.setTelNo(rs.getString("telNo")==null?"":rs.getString("telNo"));
    			bean.setmName(rs.getString("mName")==null?"":rs.getString("mName"));
    			bean.setAudit_name(rs.getString("audit_name")==null?"":rs.getString("audit_name"));
    			DATA_LIST.add(bean) ;
    		}
    		rs.close(); 
    		stmt.close(); 
    	} catch(Exception e) {
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
		    int rowNo = 0 ;
			System.out.println("1.報表名稱");
			//1.報表名稱
			FILE_NAME = "RPFFL011W"+getYYYYMMDDHHMMSS()+".xls" ;
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
    		ps.setScale( ( short )85 ); //列印縮放百分比
			ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
			ps.setLandscape( true ); // 設定橫式
			//wb.SetRepeatingRowsAndColumns(sheetNo, firstColumn, secondColumn, startRow, endRow);
			wb.setRepeatingRowsAndColumns(0, 0, 9, 2, 3);
			
			//set Query Date
			String queryDate = "查詢日期："+Utility.getCHTYYMMDD("yy").concat("年").concat(Utility.getCHTYYMMDD("mm")).concat("月").concat(Utility.getCHTYYMMDD("dd")).concat("日") ;
			row  = sheet.getRow(1) ;
			cell = row.getCell((short)9) ;
			cell.setCellValue(queryDate); 
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			//取得欄位style 
			row = sheet.getRow(4) ;
			cell = row.getCell((short)0) ;
			simpleStyle = cell.getCellStyle() ;
			rowNo = 5 ;
			//System.out.println(" data row="+DATA_LIST.size());
			for(int i=0 ; i < DATA_LIST.size() ;i++) {
				ReportBean rb = DATA_LIST.get(i) ;
				if(i==0) {
					cell = row.getCell((short)0) ; 
					cell.setCellValue(rb.getBank_no().concat(rb.getBank_name()));
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell = row.getCell((short)1) ;
					cell.setCellValue(rb.getEx_type_name());
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell = row.getCell((short)2) ;
					cell.setCellValue(rb.getEx_no());
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell = row.getCell((short)3) ;
					cell.setCellValue(rb.getDoc_date());
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell = row.getCell((short)4) ;
					cell.setCellValue(rb.getDocNo());
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell = row.getCell((short)5) ;
					cell.setCellValue(rb.getDoc_type_name());
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell = row.getCell((short)6) ;
					cell.setCellValue(rb.getLimitDate());
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
					setExcelCellValue(sheet ,row,cell,rowNo,(short)0,rb.getBank_no().concat(rb.getBank_name()),simpleStyle) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)1,rb.getEx_type_name(),simpleStyle) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)2,rb.getEx_no(),simpleStyle) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)3,rb.getDoc_date(),simpleStyle) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)4,rb.getDocNo(),simpleStyle) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)5,rb.getDoc_type_name(),simpleStyle) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)6,rb.getLimitDate(),simpleStyle) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)7,rb.getTelNo(),simpleStyle) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)8,rb.getmName(),simpleStyle) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)9,rb.getAudit_name(),simpleStyle) ;
					rowNo++ ;
				}
			}
			
			outPutExcel(wb ,sheet) ;
			
		}catch(Exception e){
			System.out.println("createRpt Error:"+e+e.getMessage());
		} finally {
			try {
				con.commit(); 
				con.close();
			}catch(Exception e){
				e.printStackTrace();
			}
		}
		return FILE_NAME ;
	}

	
	
	static class ReportBean {
    	private String bank_no = "" ;  //農漁會別.機構代號
    	private String bank_name = ""; //農漁會別.機構名稱
    	private String ex_type_name = ""; //查核類別
    	private String ex_no ="" ;  //查核報告編號
    	private String ex_no_list = "" ; //報表需顯示的檢查報告編號或查核季別或訪查日期
    	private String doc_date = ""; //農金局發文.日期
    	private String docNo = "";  //農金局發文.文號 
    	private String doc_type= "";//發文性質代碼 A:陳述意見/B:核處/C:結案
    	private String doc_type_name = ""; //農金局發文.發文性質名稱
    	private String limitDate = "";  //應回覆日期
    	private String telNo = "" ; //電話
    	private String mName = ""; //信用部主任
    	private String audit_name = ""; //稽核
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
		public String getLimitDate() {
			return limitDate;
		}
		public void setLimitDate(String limitDate) {
			this.limitDate = limitDate;
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
