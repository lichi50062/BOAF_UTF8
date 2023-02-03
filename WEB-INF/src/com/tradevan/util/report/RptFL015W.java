
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
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFFooter;
import org.apache.poi.hssf.usermodel.HSSFPrintSetup;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import com.tradevan.util.Utility;
import com.tradevan.util.dao.RdbCommonDao;
import com.tradevan.util.report.RptFL014W.ReportBean;
/***
 * 政策性農業專案貸款糾正及罰鍰處分一覽表
 * @author 2808
 *
 */
public class RptFL015W {
	/**Report執行SQL*/
	private static String QUERY_SQL =  "" ;
	/**報表名稱*/
	private static String FILE_NAME = "" ;
	/**錯誤訊息*/
	private static String ERROR_MSG = "" ;
	
	private static Connection con = null ;
	/**發文期間 起日*/
	private static String QUERY_SD= "" ;
	/**發文期間 迄日*/
	private static String QUERY_ED= "" ;
	
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
		sql.append(" decode(ex_type,'FEB',frm_snrtdoc.ex_no,'AGRI',substr(frm_snrtdoc.ex_no,0,3)||'年第'||substr(frm_snrtdoc.ex_no,4,2)||'季','BOAF',F_TRANSCHINESEDATE(to_date(frm_snrtdoc.ex_no,'yyyymmdd')),'') as ex_no_list,");// --報表需顯示的檢查報告編號或查核季別或訪查日期
		sql.append(" F_TRANSCHINESEDATE(frm_snrtdoc.doc_date) as doc_date,");//--核處日期
		sql.append(" frm_snrtdoc.docno,");//--核處文號 
		sql.append(" decode(audit_id_c2,'1','O','X') as audit_id_c2,");//--行政處分.糾正
		sql.append(" decode(audit_id_c1,'1','O','X') as audit_id_c1,");//--行政處分.罰鍰
		sql.append(" fine_amt,");//--罰鍰金額
		sql.append(" F_TRANSCHINESEDATE(fine_date) as fine_date ");//--繳款日期
		sql.append(" from frm_snrtdoc ");
		sql.append(" left join (select * from bn01 where m_year=100)bn01 on frm_snrtdoc.bank_no=bn01.bank_no ");
		sql.append(" left join frm_exmaster on frm_snrtdoc.ex_no=frm_exmaster.ex_no and ");
		sql.append(" frm_snrtdoc.bank_no=frm_exmaster.bank_no ") ;
		sql.append(" where doc_type='B' ");//-核處
		sql.append(" and audit_id in ('C0') ");//--行政處分
		if(!"".equals(QUERY_SD) && !"".equals(QUERY_ED)) {
			sql.append(" and to_char(doc_date,'yyyymmdd') between ? and ? ") ;
		}
 		//sql.append(" and TO_CHAR(doc_date ,'yyyymmdd') BETWEEN 'UI.發文日期 begin ex:20160901' AND 'UI.發文日期 end ex:20160930' ");// --發文日期有值時,才加入
		sql.append(" group by frm_snrtdoc.bank_no,bank_name,frm_exmaster.ex_type,frm_snrtdoc.ex_no,frm_snrtdoc.doc_date,frm_snrtdoc.docno,audit_id_c2,audit_id_c1,fine_amt,fine_date ");
		sql.append(" order by frm_snrtdoc.bank_no asc,ex_no asc,doc_date desc");
		//System.out.println("sql:"+sql.toString()) ;



		QUERY_SQL = sql.toString() ;
    }
    
    
    private static void getQueryData (Connection con , ResultSet rs ) {
    	
    	try {
    		PreparedStatement  stmt = con.prepareStatement(QUERY_SQL) ;
    		if(!"".equals(QUERY_SD) && !"".equals(QUERY_ED)) {
    			stmt.setString(1, transDateStyle(2,QUERY_SD));
    			stmt.setString(2, transDateStyle(2,QUERY_ED));
    		}
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
    			bean.setAudit_id_c2(rs.getString("audit_id_c2")==null?"":rs.getString("audit_id_c2"));
    			bean.setAudit_id_c1(rs.getString("audit_id_c1")==null?"":rs.getString("audit_id_c1"));
    			bean.setFine_amt(rs.getString("fine_amt")==null?"":rs.getString("fine_amt"));
    			bean.setFine_date(rs.getString("fine_date")==null?"":rs.getString("fine_date"));
    			DATA_LIST.add(bean) ;
    		}
    		rs.close(); 
    		stmt.close(); 
    	} catch(Exception e) {
    		e.printStackTrace(); 
    	} 
    }
    private static HSSFWorkbook  getExcelSimple() {
		String openFile= "FL015W政策性農業專案貸款糾正及罰鍰處分一覽表.xls" ;
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
    
	public static String createRpt(String begDate, String endDate){
		QUERY_SD = begDate ; 
		QUERY_ED = endDate ;
		String subTitle = "" ;//查詢條件
		try {
			HSSFSheet sheet = null ;
		    HSSFWorkbook wb = null ;
		    HSSFRow row = null ; 
		    HSSFCell cell = null ;
		    HSSFCellStyle simpleStyle = null;
		    //金額靠右
		    HSSFCellStyle simpleStyleRight = null;
		    //內文至中 
		    HSSFCellStyle simpleStyleCenter = null;
		    
		    int rowNo = 0 ;
		    System.out.println("1.報表名稱");
			//1.報表名稱
			FILE_NAME = "RptFL015W"+getYYYYMMDDHHMMSS()+".xls" ;
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
    		ps.setScale( ( short )100 ); //列印縮放百分比
			ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
			ps.setLandscape( true ); // 設定橫式
			if(!"".equals(QUERY_SD) && !"".equals(QUERY_ED)) {
				subTitle = "發文日期："+transChtDateFormate(QUERY_SD)
						   +"至"+
						   transChtDateFormate(QUERY_ED) ;
			}
			row  = sheet.getRow(1) ;
			cell = row.getCell((short)0) ;
			cell.setCellValue(subTitle); 
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			//取得欄位style 
			row = sheet.getRow(5) ;
			cell = row.getCell((short)0) ;
			simpleStyle = cell.getCellStyle() ;
			simpleStyleRight = row.getCell((short)7).getCellStyle() ;
			simpleStyleCenter = row.getCell((short)5).getCellStyle() ;
			rowNo = 6 ;
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
					cell.setCellValue(rb.getAudit_id_c2());
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell = row.getCell((short)6) ;
					cell.setCellValue(rb.getAudit_id_c1());
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell = row.getCell((short)7) ;
					cell.setCellValue(rb.getFine_amt());
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell = row.getCell((short)8) ;
					cell.setCellValue(rb.getFine_date());
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				} else {
					//row = sheet.createRow(++rowNo);
					setExcelCellValue(sheet ,row,cell,rowNo,(short)0,rb.getBank_no().concat(rb.getBank_nm()),simpleStyle) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)1,rb.getEx_type_name(),simpleStyle) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)2,rb.getEx_no_list(),simpleStyle) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)3,rb.getDoc_date(),simpleStyle) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)4,rb.getDocNo(),simpleStyle) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)5,rb.getAudit_id_c2(),simpleStyleCenter) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)6,rb.getAudit_id_c1(),simpleStyleCenter) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)7,rb.getFine_amt(),simpleStyleRight) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)8,rb.getFine_date(),simpleStyle) ;
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
		return FILE_NAME;
	}

	
	
	private static String NumberFormat(String m ) {
		String t = "" ;
		if(!"".equals(m)) {
			DecimalFormat  fat = new DecimalFormat ("$###,###") ;
			t = fat.format(Double.parseDouble(m)) ; 
		} 
		return t ; 
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
	private static String transChtDateFormate(String d) {
		if(d.length()!=7) {
			return "" ;
		}
		return d.substring(0,3).concat("年").concat(d.substring(3,5)).concat("月").concat(d.substring(5,7)).concat("日");
	}
	static class ReportBean {
    	private String bank_no = "" ;
    	private String bank_nm = "";
    	private String ex_type_name ;
    	private String ex_no  = "" ;
    	private String ex_no_list = "";
    	private String doc_date = "";
    	private String docNo = "";
    	private String audit_id_c2 = "" ;
    	private String audit_id_c1 = "" ;
    	private String fine_amt = "";
    	private String fine_date = "" ;
    	
    	
    	
    	public String getFine_amt(){
    		if("".equals(fine_amt)) {
    			fine_amt = "0" ;
    		} else {
    			fine_amt = NumberFormat(fine_amt) ;
    		}
    		return fine_amt ;
    	}
    	public void setFine_amt(String v) {
    		fine_amt = v ;
    	}
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
		public String getFine_date() {
			return fine_date;
		}
		public void setFine_date(String fine_date) {
			this.fine_date = fine_date;
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
