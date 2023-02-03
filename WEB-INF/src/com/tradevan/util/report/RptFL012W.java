
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
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFFooter;
import org.apache.poi.hssf.usermodel.HSSFPrintSetup;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import com.tradevan.util.DBManager;
import com.tradevan.util.Utility;
import com.tradevan.util.dao.DataObject;
import com.tradevan.util.dao.RdbCommonDao;
import com.tradevan.util.report.RptFL011W.ReportBean;
/***
 * 政策性農業專案貸款未結缺失案件明細表
 * @author 2808
 * 105.12.01 add 查詢條件,增加查核類別(FEB:金管會檢查報告/AGRI:農業金庫查核/BOAF:農金局訪查/ALL:全部)/金額單位;查核類別預設在全部 by 2968
 */
public class RptFL012W {
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
    private static List getReportSQL(String ex_type,String unit) {
    	StringBuffer sql = new StringBuffer() ;
    	List paramList = new ArrayList();//共同參數
    	sql.append(" select frm_exmaster.bank_no,");//--農漁會.機構代號
		sql.append(" bank_name,");//--農漁會.機構名稱
		sql.append(" decode(frm_exmaster.ex_type,'FEB','金管會檢查報告','AGRI','農業金庫查核','BOAF','農金局訪查','') as ex_type_name,");//--查核類別名稱
		sql.append(" frm_exmaster.ex_no,");// --查核報告編號
		sql.append(" decode(frm_exmaster.ex_type,'FEB',frm_exmaster.ex_no,'AGRI',substr(frm_exmaster.ex_no,0,3)||'年第'||substr(frm_exmaster.ex_no,4,2)||'季','BOAF',F_TRANSCHINESEDATE(to_date(frm_exmaster.ex_no,'yyyymmdd')),'') as ex_no_list,");// --報表需顯示的檢查報告編號或查核季別或訪查日期
		sql.append(" loan_name, ");//--借款人名稱
		sql.append(" F_TRANSCHINESEDATE(loan_date) as loan_date,");//--貸款日期
		sql.append(" loan_item_name,");//--貸款種類名稱
		sql.append(" round(loan_amt/?) as loan_amt, ");//--貸款金額
		paramList.add(unit);
		sql.append(" F_COMBEX_DEF_CASE (frm_exmaster.bank_no,frm_exmaster.ex_no,loan_name,to_char(loan_date,'yyyy/mm/dd'),");
		sql.append(" frm_exdef.loan_item,loan_amt) as def_case ");//--缺失摘要
		sql.append(" from frm_exdef ");
		sql.append(" left join frm_loan_item on frm_exdef.loan_item = frm_loan_item.loan_item ");
		sql.append(" left join frm_exmaster on frm_exdef.ex_no=frm_exmaster.ex_no and ");
		sql.append(" frm_exdef.bank_no=frm_exmaster.bank_no ") ;
		sql.append(" left join frm_def_item on frm_exdef.def_type = frm_def_item.def_type and frm_exdef.def_case = frm_def_item.def_case ");
		sql.append(" left join (select * from bn01 where m_year=100)bn01 on frm_exdef.bank_no=bn01.bank_no");
		sql.append(" where  frm_exmaster.case_status !='0' ");//--未結案
		if(!"ALL".equals(ex_type)){
			sql.append(" and frm_exmaster.ex_type = ? ");
			paramList.add(ex_type);
		}
		sql.append(" group by frm_exmaster.bank_no,bank_name,frm_exmaster.ex_type,");
		sql.append(" frm_exmaster.ex_no,loan_name,loan_date,loan_item_name,loan_amt,");
		sql.append(" F_COMBEX_DEF_CASE (frm_exmaster.bank_no,frm_exmaster.ex_no,loan_name, ");
		sql.append(" to_char(loan_date,'yyyy/mm/dd'),frm_exdef.loan_item,loan_amt) ");
		sql.append(" order by frm_exmaster.bank_no,ex_no,def_case ");
		
		
	    List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"bank_no,bank_name,ex_type_name,ex_no,ex_no_list,loan_name,loan_date,loan_item_name,loan_amt,def_case");
		System.out.println("getReportSQL.size()="+dbData.size());
		return dbData;
		//QUERY_SQL = sql.toString() ;
    }
    /**
     * get 農漁會家數合計SQL.
     * 
     * */
    
	private static List getReportSQL2(String ex_type) {
		StringBuffer sql = new StringBuffer();
		List paramList = new ArrayList();//共同參數
		sql.append(" Select count(*) as bank_count ") ; //農漁會家數
		sql.append(" from ( ") ;
		sql.append(" Select bank_no,count(*) ") ;
		sql.append(" from frm_exmaster ") ;
		sql.append(" where frm_exmaster.case_status!='0' ") ; //未結案
		if(!"ALL".equals(ex_type)){
			sql.append(" and frm_exmaster.ex_type = ? ");
			paramList.add(ex_type);
		}
		sql.append(" group by bank_no ) a ") ;
		
		List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"bank_count");
		System.out.println("getReportSQL2.size()="+dbData.size());
		return dbData;
		//QUERY_SQL = sql.toString() ;
	}
    
    private static void getQueryData (List qList) {
    	DATA_LIST.clear(); 
    	for(int i=0;i<qList.size();i++){
    		ReportBean bean = new ReportBean () ;
    		DataObject obj = (DataObject)qList.get(i);
			bean.setBank_no(obj.getValue("bank_no")==null?"":(String)obj.getValue("bank_no"));
			bean.setBank_nm(obj.getValue("bank_name")==null?"":(String)obj.getValue("bank_name"));
			bean.setEx_type_name(obj.getValue("ex_type_name")==null?"":(String)obj.getValue("ex_type_name"));
			bean.setEx_no(obj.getValue("ex_no")==null?"":(String)obj.getValue("ex_no"));
			bean.setEx_no_list(obj.getValue("ex_no_list")==null?"":(String)obj.getValue("ex_no_list"));
			bean.setLoan_name(obj.getValue("loan_name")==null?"":(String)obj.getValue("loan_name"));
			bean.setLoan_date(obj.getValue("loan_date")==null?"":(String)obj.getValue("loan_date"));
			bean.setLoan_item_name(obj.getValue("loan_item_name")==null?"":(String)obj.getValue("loan_item_name"));
			bean.setLoan_amt(obj.getValue("loan_amt")==null?"":obj.getValue("loan_amt").toString());
			bean.setDef_case(obj.getValue("def_case")==null?"":(String)obj.getValue("def_case"));
			DATA_LIST.add(bean) ;
    	}
    	/*
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
    			bean.setLoan_name(rs.getString("loan_name")==null?"":rs.getString("loan_name"));
    			bean.setLoan_date(rs.getString("loan_date")==null?"":rs.getString("loan_date"));
    			bean.setLoan_item_name(rs.getString("loan_item_name")==null?"":rs.getString("loan_item_name"));
    			bean.setLoan_amt(rs.getString("loan_amt")==null?"":rs.getString("loan_amt"));
    			bean.setDef_case(rs.getString("def_case")==null?"":rs.getString("def_case"));
    			DATA_LIST.add(bean) ;
    		}
    		rs.close(); 
    		stmt.close(); 
    	} catch(Exception e) {
    		e.printStackTrace(); 
    	} */
    }
    private static String getQueryData2 (List qList) {
    	/*ReportBean bean = new ReportBean () ;
    	try {
    		PreparedStatement  stmt = con.prepareStatement(QUERY_SQL) ;
    		rs = stmt.executeQuery() ;
    		if(rs.next()) {
    			bean.setBank_count(rs.getString(1));
    		}
    		
    		rs.close(); 
    		stmt.close(); 
    	} catch(Exception e) {
    		e.printStackTrace(); 
    	} */
    	ReportBean bean = new ReportBean () ;
    	for(int i=0;i<qList.size();i++){
    		DataObject obj = (DataObject)qList.get(i);
    		bean.setBank_count(obj.getValue("bank_count")==null?"":obj.getValue("bank_count").toString());
    	}
    	return bean.getBank_count();
    }
    /*private static String getQueryData2 (Connection con , ResultSet rs ) {
    	ReportBean bean = new ReportBean () ;
    	try {
    		PreparedStatement  stmt = con.prepareStatement(QUERY_SQL) ;
    		rs = stmt.executeQuery() ;
    		if(rs.next()) {
    			bean.setBank_count(rs.getString(1));
    		}
    		
    		rs.close(); 
    		stmt.close(); 
    	} catch(Exception e) {
    		e.printStackTrace(); 
    	} 
    	return bean.getBank_count();
    }*/
    private static HSSFWorkbook  getExcelSimple() {
		String openFile= "FL012W政策性農業專案貸款未結缺失案件明細表.xls" ;
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
	public static String createRpt(String ex_type,String unit){
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
			FILE_NAME = "RptFL012W"+getYYYYMMDDHHMMSS()+".xls" ;
			//2.取得SQL
			System.out.println("2.取得SQL");
			List qList= getReportSQL(ex_type,unit) ;
			//3.取得連線
			System.out.println("3.取得連線");
			con = (new RdbCommonDao("")).newConnection();//DBManager.newQryConnection(). ;
			ResultSet rs = null ;
			//4.取得資料
			System.out.println("4.取得資料");
			getQueryData(qList) ;  
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
			//set Query Date
			String queryDate = "查詢日期："+Utility.getCHTYYMMDD("yy").concat("年").concat(Utility.getCHTYYMMDD("mm")).concat("月").concat(Utility.getCHTYYMMDD("dd")).concat("日") ;
			row  = sheet.getRow(1) ;
			cell = row.getCell((short)7) ;
			cell.setCellValue(queryDate); 
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell = row.getCell((short)0) ;
		    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	        cell.setCellValue("單位:新臺幣 " + Utility.getUnitName(unit));//105.12.01 add 顯示金額單位 	        
			//取得欄位style 
			row = sheet.getRow(4) ;
			cell = row.getCell((short)0) ;
			simpleStyle = cell.getCellStyle() ;
			simpleStyleRight = row.getCell((short)6).getCellStyle() ;
			rowNo = 5 ;
			System.out.println(" data row="+DATA_LIST.size());
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
					cell.setCellValue(rb.getLoan_name());
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell = row.getCell((short)4) ;
					cell.setCellValue(rb.getLoan_date());
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell = row.getCell((short)5) ;
					cell.setCellValue(rb.getLoan_item_name());
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell = row.getCell((short)6) ;
					cell.setCellValue(rb.getLoan_amt());
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell = row.getCell((short)7) ;
					cell.setCellValue(rb.getDef_case());
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				} else {
					//row = sheet.createRow(++rowNo);
					setExcelCellValue(sheet ,row,cell,rowNo,(short)0,rb.getBank_no().concat(rb.getBank_nm()),simpleStyle) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)1,rb.getEx_type_name(),simpleStyle) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)2,rb.getEx_no_list(),simpleStyle) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)3,rb.getLoan_name(),simpleStyle) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)4,rb.getLoan_date(),simpleStyle) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)5,rb.getLoan_item_name(),simpleStyle) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)6,rb.getLoan_amt(),simpleStyleRight) ;
					setExcelCellValue(sheet ,row,cell,rowNo,(short)7,rb.getDef_case(),simpleStyle) ;
					
					rowNo++ ;
				}
			}
			//=====================================================================
			//do 農漁會家數合計 
			int cellNo = 0 ;
			for(int k=0 ; k < 2 ; k++) {
				setExcelCellValue(sheet ,row,cell,rowNo,(short)0,"",simpleStyle) ;
				setExcelCellValue(sheet ,row,cell,rowNo,(short)1,"",simpleStyle) ;
				setExcelCellValue(sheet ,row,cell,rowNo,(short)2,"",simpleStyle) ;
				setExcelCellValue(sheet ,row,cell,rowNo,(short)3,"",simpleStyle) ;
				setExcelCellValue(sheet ,row,cell,rowNo,(short)4,"",simpleStyle) ;
				setExcelCellValue(sheet ,row,cell,rowNo,(short)5,"",simpleStyle) ;
				setExcelCellValue(sheet ,row,cell,rowNo,(short)6,"",simpleStyleRight) ;
				setExcelCellValue(sheet ,row,cell,rowNo,(short)7,"",simpleStyle) ;
				rowNo++ ;
			}
			
			List qList2 = getReportSQL2(ex_type) ;
			setExcelCellValue(sheet ,row,cell,rowNo,(short)0,getQueryData2(qList2),simpleStyleRight) ;
			setExcelCellValue(sheet ,row,cell,rowNo,(short)1,"",simpleStyle) ;
			setExcelCellValue(sheet ,row,cell,rowNo,(short)2,"",simpleStyle) ;
			setExcelCellValue(sheet ,row,cell,rowNo,(short)3,String.valueOf(DATA_LIST.size()),simpleStyleRight) ;
			setExcelCellValue(sheet ,row,cell,rowNo,(short)4,"",simpleStyle) ;
			setExcelCellValue(sheet ,row,cell,rowNo,(short)5,"",simpleStyle) ;
			setExcelCellValue(sheet ,row,cell,rowNo,(short)6,"",simpleStyleRight) ;
			setExcelCellValue(sheet ,row,cell,rowNo,(short)7,"",simpleStyle) ;
					
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
		DecimalFormat  fat = new DecimalFormat ("$###,###") ;
		return fat.format(Double.parseDouble(m)) ;
	}
	static class ReportBean {
    	private String bank_no = "" ;
    	private String bank_nm = "";
    	private String ex_type_name ;
    	private String ex_no  = "" ;
    	private String ex_no_list = "";
    	private String loan_name = "";
    	private String loan_date = ""; 
    	private String loan_item_name = "";
    	private String loan_amt = "";
    	private String def_case = "";
    	private String bank_count = "" ;
    	
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
			return NumberFormat("".equals(loan_amt)?"0" : loan_amt );
		}
		public void setLoan_amt(String loan_amt) {
			this.loan_amt = loan_amt;
		}
		public String getDef_case() {
			//return def_case;
			String st = "" ;
			String [] sp = def_case.split("&") ;
			for(int i= 1 ; i<sp.length ;i++) {
				st += sp[i]+"\r\n" ;
			}
			
			return st ; 
		}
		public void setDef_case(String def_case) {
			this.def_case = def_case;
		}
		public String getBank_count() {
			return bank_count;
		}
		public void setBank_count(String bank_count) {
			this.bank_count = bank_count;
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
