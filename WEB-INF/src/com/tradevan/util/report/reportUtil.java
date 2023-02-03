/*
 *  Created on 2005/1/13
 *  95.08.23 add 無框內文置左 by 2295
 *  95.09.21 add 報表欄位名稱用--有框內文置左 by 2295
 *  95.09.21 add 違返規則欄位--無框內文置中,紅色+底線 by 2295
 *  95.09.21 add 違返規則欄位--無框內文置右,紅色+底線 by 2295
 *  95.10.03 add 儲存報表格式檔.讀取報表格式檔.刪除報表格式檔.取得範本資料 by 2295
 *  95.12.04 add 儲存/讀取範本.增加結束日期 by 2295 
 *  99.12.23 add 儲存報表格式檔.讀取報表格式檔(for BR基本報表用) by 2295
 * 102.11.19 add 原QueryDB改套用QueryDB_SQLParam by 2295    
 * 103.05.06 add saveReport()/readReport()增加loan_item貸款項目別  by 2968
 * 104.03.09 add saveReport()/readReport()增加utf-8編碼 by 2295
 * 104.03.11 add saveReport_BR()/readReport_BR()增加utf-8編碼 by 2295
 * 110.10.08 add saveReport()/readReport()增加S_DAY/E_DAY/sysType/loginFlag by 2295
 */
package com.tradevan.util.report;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpSession;
import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.Region;

public class reportUtil {
	public HSSFCellStyle defaultStyle ;
    public HSSFCellStyle noBorderDefaultStyle ;
    
    //設定樣式和位置(請精減style物件的使用量，以免style物件太多excel報表無法開啟)
    //有框內文置中
    public HSSFCellStyle getDefaultStyle(HSSFWorkbook wb){ 
    	   HSSFCellStyle defaultStyle1 = wb.createCellStyle(); 
    	   defaultStyle1 = HssfStyle.setStyle( defaultStyle1, wb.createFont(),
                                       new String[] {
                                       "BORDER", "PHC", "PVC", "F10",
                                       "WRAP"} );
    	   return defaultStyle1;
    } 
    //有框內文置中.紅色字体
    public HSSFCellStyle getDefaultRedStyle(HSSFWorkbook wb){ 
    	   HSSFCellStyle defaultStyle1 = wb.createCellStyle(); 
    	   HSSFFont f = wb.createFont();
   		   //set font 1 to 12 point type
   		   f.setFontHeightInPoints((short) 10);   		 
   		   //make it red
   		   f.setColor( HSSFFont.COLOR_RED );
    	   defaultStyle1 = HssfStyle.setStyle( defaultStyle1, f,
                                       new String[] {
                                       "BORDER", "PHC", "PVC", "F10",
                                       "WRAP"} );
    	   return defaultStyle1;
    } 
    //無框內文置中
    public HSSFCellStyle getNoBorderDefaultStyle(HSSFWorkbook wb){ 
    	    HSSFCellStyle noBorderDefaultStyle = wb.createCellStyle(); 
    		noBorderDefaultStyle = HssfStyle.setStyle( noBorderDefaultStyle,
    									       wb.createFont(), new String[] {
      										   "PHC", "PVC", "F10", "WRAP"} );
    		return noBorderDefaultStyle;
    } 
   //95.08.23 add 無框內文置左
    public HSSFCellStyle getNoBorderLeftStyle(HSSFWorkbook wb){ 
    	    HSSFCellStyle noBorderLeftStyle = wb.createCellStyle(); 
    		noBorderLeftStyle = HssfStyle.setStyle( noBorderLeftStyle,
    									       wb.createFont(), new String[] {
      										   "PHL", "PVC", "F10", "WRAP"} );
    		return noBorderLeftStyle;
    } 
	//自定需求style
    //標題用
    public HSSFCellStyle getTitleStyle(HSSFWorkbook wb){
    	   HSSFCellStyle titleStyle = wb.createCellStyle(); 
    	   titleStyle = HssfStyle.setStyle( titleStyle, wb.createFont(),
                                     new String[] {
                                     "PHC", "PVC", "F24"} );
    	  return titleStyle; 
    } 
    //報表欄位名稱用--有框內文置中
    public HSSFCellStyle getColumnStyle(HSSFWorkbook wb){
    		HSSFFont f = wb.createFont();
    		//set font 1 to 12 point type
    		f.setFontHeightInPoints((short) 12);
    		//make it blue
    		f.setColor( (short)0xc );
    		//make it bold
    		//arial is the default font
    		f.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        
    		HSSFCellStyle columnStyle = wb.createCellStyle(); 
    		columnStyle = HssfStyle.setStyle( columnStyle, f,
                                     new String[] {
                                     "BORDER", "PHC", "PVC", "F10",
                                     "WRAP"} );
    		return columnStyle;
    }
  //報表欄位名稱用--有框內文置左 //95.09.21 add by 2295    
    public HSSFCellStyle getColumn_LeftStyle(HSSFWorkbook wb){
    		HSSFFont f = wb.createFont();
    		//set font 1 to 12 point type
    		f.setFontHeightInPoints((short) 12);
    		//make it blue
    		f.setColor( (short)0xc );
    		//make it bold
    		//arial is the default font
    		f.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        
    		HSSFCellStyle columnStyle = wb.createCellStyle(); 
    		columnStyle = HssfStyle.setStyle( columnStyle, f,
                                     new String[] {
                                     "BORDER", "PHL", "PVC", "F10",
                                     "WRAP"} );
    		return columnStyle;
    }    
    //95.09.21 add 違返規則欄位--無框內文置中,紅色+底線 
    public HSSFCellStyle getCenter_Red_UnderlineStyle(HSSFWorkbook wb){
    		HSSFFont f = wb.createFont();
    		//set font 1 to 12 point type
    		f.setFontHeightInPoints((short) 10);
    		//make it single (normal) underline 
    		f.setUnderline(HSSFFont.U_SINGLE); 
    		//make it red
    		f.setColor( HSSFFont.COLOR_RED );
    		
    		HSSFCellStyle columnStyle = wb.createCellStyle(); 
    		columnStyle = HssfStyle.setStyle( columnStyle, f,
                                     new String[] {
                                     "BORDER", "PHC", "PVC", "F10",
                                     "WRAP"} );
    		return columnStyle;
    }
    //95.09.21 add 違返規則欄位--無框內文置右,紅色+底線 
    public HSSFCellStyle getRight_Red_UnderlineStyle(HSSFWorkbook wb){
    		HSSFFont f = wb.createFont();
    		//set font 1 to 12 point type
    		f.setFontHeightInPoints((short) 10);
    		//make it single (normal) underline 
    		f.setUnderline(HSSFFont.U_SINGLE); 
    		//make it red
    		f.setColor( HSSFFont.COLOR_RED );
    		
    		HSSFCellStyle columnStyle = wb.createCellStyle(); 
    		columnStyle = HssfStyle.setStyle( columnStyle, f,
                                     new String[] {
                                     "BORDER", "PHR", "PVC", "F10",
                                     "WRAP"} );
    		return columnStyle;
    }
    //有框內文置左
    public HSSFCellStyle getLeftStyle(HSSFWorkbook wb){
    		HSSFCellStyle leftStyle = wb.createCellStyle(); 
    		leftStyle = HssfStyle.setStyle( leftStyle, wb.createFont(),
                                    new String[] {
                                    "BORDER", "PHL", "PVC", "F10",
                                    "WRAP"} );
    		return leftStyle;	
    }                                                  
    //有框內文置右
    public HSSFCellStyle getRightStyle(HSSFWorkbook wb){
    		HSSFCellStyle rightStyle = wb.createCellStyle(); 
    		rightStyle = HssfStyle.setStyle( rightStyle, wb.createFont(),
                                     new String[] {
                                     "BORDER", "PHR", "PVC", "F10",
                                     "WRAP"} );
           return rightStyle; 
    }
    //有框內文置右
    public static HSSFCellStyle getRightStyleF12(HSSFWorkbook wb){
    		HSSFCellStyle rightStyle = wb.createCellStyle();    		
    		rightStyle = HssfStyle.setStyle( rightStyle, wb.createFont(),
                                     new String[] {
                                     "BORDER", "PHR", "PVC", "F12",
                                     "WRAP"} );
    		return rightStyle;
    }
   //有框內文置中
    public HSSFCellStyle getCenterStyleF12(HSSFWorkbook wb){ 
    	   HSSFCellStyle defaultStyle1 = wb.createCellStyle();    	   
    	   defaultStyle1 = HssfStyle.setStyle( defaultStyle1, wb.createFont(),
                                       new String[] {
                                       "BORDER", "PHC", "PVC", "F12",
                                       "WRAP"} );
    	   return defaultStyle1;
    } 
    //有框小字
    public HSSFCellStyle getSmallFontStyle(HSSFWorkbook wb){
    		HSSFCellStyle smallFontStyle = wb.createCellStyle(); 
    		smallFontStyle = HssfStyle.setStyle( smallFontStyle, wb.createFont(),
                                         new String[] {
                                         "BORDER", "PHC", "PVC", "F08",
                                         "WRAP"} );
    		return smallFontStyle;
    }
    
    //無框置右
    public HSSFCellStyle getNoBoderStyle(HSSFWorkbook wb){
    		HSSFCellStyle noBoderStyle = wb.createCellStyle(); 
    		noBoderStyle = HssfStyle.setStyle( noBoderStyle, wb.createFont(),
                                       new String[] {
                                       "PHR", "PVC", "F10", "WRAP"} );
    		return noBoderStyle;
    }
    
    public void setDefaultStyle( HSSFCellStyle szdefaultStyle){
    	this.defaultStyle = szdefaultStyle;
    }
    
    public void setNoBorderDefaultStyle( HSSFCellStyle sznoBorderDefaultStyle){
    	this.noBorderDefaultStyle = sznoBorderDefaultStyle;
    }
	public  void setBorder( HSSFCellStyle style ) {
        style.setBorderBottom( HSSFCellStyle.BORDER_THIN );
        style.setBorderLeft( HSSFCellStyle.BORDER_THIN );
        style.setBorderRight( HSSFCellStyle.BORDER_THIN );
        style.setBorderTop( HSSFCellStyle.BORDER_THIN );
        style.setVerticalAlignment( HSSFCellStyle.VERTICAL_CENTER );
        style.setAlignment( HSSFCellStyle.ALIGN_CENTER );
    }


    public  void createCell( HSSFWorkbook wb, HSSFRow row, short column,
                             String value,int[] columnLen ) {
        createCell( wb, row, column, value, true ,columnLen);
    }


    public  void createCell( HSSFWorkbook wb, HSSFRow row, short column,
                             String value, boolean hasBorder,int[] columnLen ) {
        HSSFCell cell = row.createCell( column );
        cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
        cell.setCellValue( value );
        int len = value.getBytes().length;
        if ( len > columnLen[column - 1] ) {
            columnLen[column - 1] = len;
        }
        HSSFCellStyle style = wb.createCellStyle();
        if ( hasBorder ) {
            setBorder( style );
            //style.setAlignment(HSSFCellStyle.ALIGN_CENTER);

        }
        cell.setCellStyle( style );
    }


    public void createCell( HSSFWorkbook wb, HSSFRow row, short column,
                             String value, HSSFCellStyle style ) {
        HSSFCell cell = row.createCell( column );
        cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
        cell.setCellValue( value );
        cell.setCellStyle( style );
    }


    public void createCell( HSSFWorkbook wb, HSSFRow row, short column,
                             int value ) {
        createCell( wb, row, column, String.valueOf( value ), this.defaultStyle );
    }


    public void createCell( HSSFWorkbook wb, HSSFRow row, short column,
                             int value, boolean hasBorder ) {
        if ( hasBorder == true ) {
            createCell( wb, row, column, String.valueOf( value ),
                        this.defaultStyle );
        } else {
            createCell( wb, row, column, String.valueOf( value ),
                        this.noBorderDefaultStyle );
        }
    }
    
    //95.11.03 add 儲存報表格式檔========================================================
    //檔案名稱 ex:登入者帳號_程式名稱_範本名稱_建置日期.txt
    //95.12.04 add 結束日期
    //104.03.09 add 寫入檔案時加入utf-8編碼
    //110.10.08 add S_DAY/E_DAY/sysType/loginFlag
    public static String saveReport(HttpServletRequest request,String lguser_id,String template,String report_no){
    	HttpSession session = request.getSession();          	
    	String bank_type = (session.getAttribute("nowbank_type")==null)?"":(String)session.getAttribute("nowbank_type");
    	String actMsg = "";
    	String filename = "";
    	boolean isSaveLoanItem = false;
    	if("DS056W".equals(report_no)||"DS057W".equals(report_no)||"DS058W".equals(report_no)||"DS059W".equals(report_no)
                ||"DS060W".equals(report_no)||"DS061W".equals(report_no)||"DS062W".equals(report_no)
                ||"DS063W".equals(report_no)||"DS064W".equals(report_no)||"DS065W".equals(report_no)){
    	    isSaveLoanItem = true;
        }
    	try{    	
    	    String nowDate = Utility.getDateFormat("yyyyMMdd");
	        filename = lguser_id+"_"+report_no+"_"+bank_type+"_"+template+"_"+nowDate+".txt";
      		File profileDir = new File(Utility.getProperties("profileDir"));        
    		if(!profileDir.exists()){
    			if(!Utility.mkdirs(Utility.getProperties("profileDir"))){
    	   			actMsg = actMsg + Utility.getProperties("profileDir")+"目錄新增失敗";
     			}    
    		}
    		File savefile = new File(Utility.getProperties("profileDir")+System.getProperty("file.separator")+ filename);
    		if(savefile.exists()) savefile.delete();
        	FileOutputStream fos = new FileOutputStream(Utility.getProperties("profileDir")+System.getProperty("file.separator")+ filename);
			BufferedOutputStream bos = new BufferedOutputStream(fos);
			PrintStream ps = new PrintStream(bos,false,"UTF-8");//104.03.09 add
			
			String cancel_no_data = "CANCEL_NO="+(String)session.getAttribute("CANCEL_NO");//營運中/裁撤別
			String hsien_id_data = "HSIEN_ID="+(String)session.getAttribute("HSIEN_ID");//縣市別
			String BankList_data = "BankList="+(String)session.getAttribute("BankList");//金融機構代號
			String SortList_data = "SortList="+(String)session.getAttribute("SortList");//排序欄位
			String SortBy_data = "SortBy="+(String)session.getAttribute("SortBy");//遞增.遞減
			String btnFieldList_data = "btnFieldList="+(String)session.getAttribute("btnFieldList");//報表欄位
			String S_YEAR = "S_YEAR="+(String)session.getAttribute("S_YEAR");//年-begin
			String S_MONTH = "S_MONTH="+(String)session.getAttribute("S_MONTH");//月
			String Unit = "Unit="+(String)session.getAttribute("Unit");//金額單位
			//String acc_div = "acc_div="+(String)session.getAttribute("acc_div");//報表類別			   		
			ps.println(S_YEAR);
			ps.println(S_MONTH);
			if(isSaveLoanItem){
			    String loan_item = "loan_item="+(String)session.getAttribute("loan_item");//貸款項目別
			    ps.println(loan_item);
			}else{
			    String E_YEAR = "E_YEAR="+(String)session.getAttribute("E_YEAR");//年-end//95.12.04 add 結束日期
	            String E_MONTH = "E_MONTH="+(String)session.getAttribute("E_MONTH");//月
			    ps.println(E_YEAR);//95.12.04 add 結束日期
                ps.println(E_MONTH);
                if("DS094W".equals(report_no)){//110.10.08 add
                    String S_DAY = "S_DAY="+(String)session.getAttribute("S_DAY");
    	            String E_DAY = "E_DAY="+(String)session.getAttribute("E_DAY");
    	            String sysType = "sysType="+(String)session.getAttribute("sysType");
    	            String loginFlag = "loginFlag="+(String)session.getAttribute("loginFlag");
    	            ps.println(S_DAY);
                    ps.println(E_DAY); 
                    ps.println(sysType);
                    ps.println(loginFlag);
	            }
			}
			//ps.println(acc_div);			   
			ps.println(Unit);
			ps.println(cancel_no_data);
			ps.println(hsien_id_data);
			ps.println(BankList_data);
			ps.println(btnFieldList_data);
			ps.println(SortList_data);
			ps.println(SortBy_data);
			ps.close();
			bos.close();		
			fos.close();										
		}catch(Exception e){
				System.out.println("SaveReport Error:"+e+e.getMessage());
				actMsg += "SaveReport Error:"+e+e.getMessage();
		}
		return actMsg;
    } 
    
    // 99.12.23 儲存報表格式檔(BR基本報表用)========================================================
    //104.03.10 add 寫入檔案時加入utf-8編碼
    public static String saveReport_BR(HttpServletRequest request,String lguser_id,String report_no){
    	HttpSession session = request.getSession();      
    	String actMsg = "";
    	try{
          		File profileDir = new File(Utility.getProperties("profileDir"));        
        		if(!profileDir.exists()){
        			if(!Utility.mkdirs(Utility.getProperties("profileDir"))){
        	   			actMsg = actMsg + Utility.getProperties("profileDir")+"目錄新增失敗";
         			}    
        		}
        		File savefile = new File(Utility.getProperties("profileDir")+System.getProperty("file.separator")+ lguser_id+"_"+report_no+".txt");
        		if(savefile.exists()) savefile.delete();
            	FileOutputStream fos = new FileOutputStream(Utility.getProperties("profileDir")+System.getProperty("file.separator")+ lguser_id+"_"+report_no+".txt");
				BufferedOutputStream bos = new BufferedOutputStream(fos);
				PrintStream ps = new PrintStream(bos,false,"UTF-8");//104.03.11 add
				String s_year_data = "S_YEAR="+(String)session.getAttribute("S_YEAR");//99.12.23 add
				String s_month_data = "S_MONTH="+(String)session.getAttribute("S_MONTH");//99.12.23 add
				String cancel_no_data = "CANCEL_NO="+(String)session.getAttribute("CANCEL_NO");
				String hsien_id_data = "HSIEN_ID="+(String)session.getAttribute("HSIEN_ID");
				String BankList_data = "BankList="+(String)session.getAttribute("BankList");
				String SortList_data = "SortList="+(String)session.getAttribute("SortList");
				String SortBy_data = "SortBy="+(String)session.getAttribute("SortBy");
				String btnFieldList_data = "btnFieldList="+(String)session.getAttribute("btnFieldList");
				ps.println(s_year_data);//99.12.23 add
				ps.println(s_month_data);//99.12.23 add
				ps.println(cancel_no_data);
				ps.println(hsien_id_data);
				ps.println(BankList_data);
				ps.println(btnFieldList_data);
				ps.println(SortList_data);
				ps.println(SortBy_data);
				ps.close();
				bos.close();		
				fos.close();						
		}catch(Exception e){
				System.out.println("SaveReport Error:"+e+e.getMessage());
				actMsg += "SaveReport Error";
		}
		return actMsg;
    } 
    
    //95.11.03 add 讀取報表格式檔========================================================
    //ex:登入者帳號_程式名稱_範本名稱_建置日期.txt
    //95.12.04 add 結束日期
    //103.03.09 add 中文讀取時,轉utf-8
    //110.10.08 add S_DAY/E_DAY/sysType/loginFlag 
    public static String readReport(HttpServletRequest request,String lguser_id/*建置者帳號*/,String template/*範本名稱*/,String nowDate/*範本建置日期*/,String report_no/*報表名稱*/){
    	String sztemp="";
    	String CANCEL_NO_file = "";//營運中/裁撤別
        String HSIEN_ID_file = "";//縣市別
        String BankList_file = "";//金融機構代號
        String btnFieldList_file = "";//報表欄位
        String SortList_file = "";//排序欄位
        String SortBy_file = "";//遞增.遞減
        String S_YEAR = "";//年-begin
		String S_MONTH = "";//月
		String E_YEAR = "";//年-end//95.12.04 add 結束日期
		String E_MONTH = "";//月
		String S_DAY = "";//日 //110.10.08 add 
		String E_DAY = "";//日 //110.10.08 add
		String Unit = "";//金額單位
		String acc_div = "";//報表類別
		String loan_item = "";//貸款項目別
		String sysType = "";//系統類別 //110.10.08 add
	    String loginFlag = "";//登入狀態 //110.10.08 add
	   	String filename="";		   	
        String actMsg = "";
        boolean isGetLoanItem = false;
        if("DS056W".equals(report_no)||"DS057W".equals(report_no)||"DS058W".equals(report_no)||"DS059W".equals(report_no)
                ||"DS060W".equals(report_no)||"DS061W".equals(report_no)||"DS062W".equals(report_no)
                ||"DS063W".equals(report_no)||"DS064W".equals(report_no)||"DS065W".equals(report_no)){
            isGetLoanItem = true;
        }
        HttpSession session = request.getSession();          
        String bank_type = (session.getAttribute("nowbank_type")==null)?"":(String)session.getAttribute("nowbank_type");
    	try{
    	    filename = lguser_id+"_"+report_no+"_"+bank_type+"_"+template+"_"+nowDate;
    		File WorkFile = new File(Utility.getProperties("profileDir")+System.getProperty("file.separator")+filename);	        	
			if(WorkFile.exists()){			
				FileInputStream fis = new FileInputStream(Utility.getProperties("profileDir")+System.getProperty("file.separator")+ filename);	
				BufferedInputStream bis = new BufferedInputStream(fis);
				DataInputStream dis = new DataInputStream(fis);							
				while((sztemp = dis.readLine()) != null){
				    sztemp = sztemp;
				    if(sztemp.indexOf("CANCEL_NO=") != -1){
				       CANCEL_NO_file = sztemp.substring(sztemp.indexOf("CANCEL_NO=")+10,sztemp.length());				       
				    }
				    if(sztemp.indexOf("HSIEN_ID=") != -1){
				       HSIEN_ID_file = sztemp.substring(sztemp.indexOf("HSIEN_ID=")+9,sztemp.length());				       
				    }
				    if(sztemp.indexOf("BankList=") != -1){
				       BankList_file = sztemp.substring(sztemp.indexOf("BankList=")+9,sztemp.length());		
				       BankList_file = Utility.ISOtoUTF8(BankList_file);//104.03.09 add
				    }
				    if(sztemp.indexOf("btnFieldList=") != -1){
				       btnFieldList_file = sztemp.substring(sztemp.indexOf("btnFieldList=")+13,sztemp.length());	
				       btnFieldList_file = Utility.ISOtoUTF8(btnFieldList_file);//104.03.09 add
				    }	
				    if(sztemp.indexOf("SortList=") != -1){
				       SortList_file = sztemp.substring(sztemp.indexOf("SortList=")+9,sztemp.length());	
				       SortList_file = Utility.ISOtoUTF8(SortList_file);//104.03.09 add
				    }
				    if(sztemp.indexOf("SortBy=") != -1){
				       SortBy_file = sztemp.substring(sztemp.indexOf("SortBy=")+7,sztemp.length());				       
				    }
				    if(sztemp.indexOf("S_YEAR=") != -1){
				       S_YEAR = sztemp.substring(sztemp.indexOf("S_YEAR=")+7,sztemp.length());				       
				    }
				    if(sztemp.indexOf("S_MONTH=") != -1){
				       S_MONTH = sztemp.substring(sztemp.indexOf("S_MONTH=")+8,sztemp.length());				       
				    }
				    //95.12.04 add 結束日期
				    if(sztemp.indexOf("E_YEAR=") != -1){
					   E_YEAR = sztemp.substring(sztemp.indexOf("E_YEAR=")+7,sztemp.length());				       
					}
					if(sztemp.indexOf("E_MONTH=") != -1){
					   E_MONTH = sztemp.substring(sztemp.indexOf("E_MONTH=")+8,sztemp.length());				       
					}
					
				    if("DS094W".equals(report_no) && sztemp.indexOf("S_DAY=") != -1){//110.10.08 add
					   S_DAY = sztemp.substring(sztemp.indexOf("S_DAY=")+6,sztemp.length());				       
					}
				    if("DS094W".equals(report_no) && sztemp.indexOf("E_DAY=") != -1){//110.10.08 add
					   E_DAY = sztemp.substring(sztemp.indexOf("E_DAY=")+6,sztemp.length());				       
					}
					if(sztemp.indexOf("E_MONTH=") != -1){
					   E_MONTH = sztemp.substring(sztemp.indexOf("E_MONTH=")+8,sztemp.length());				       
					}
					
				    if(sztemp.indexOf("Unit=") != -1){
				       Unit = sztemp.substring(sztemp.indexOf("Unit=")+5,sztemp.length());				       
				    }
				    if(sztemp.indexOf("acc_div=") != -1){
				       acc_div = sztemp.substring(sztemp.indexOf("acc_div=")+8,sztemp.length());				       
				    }
				    if("DS094W".equals(report_no) && sztemp.indexOf("sysType=") != -1){//系統類別 //110.10.08 add
				    	sysType = sztemp.substring(sztemp.indexOf("sysType=")+8,sztemp.length());				       
					}
				    if("DS094W".equals(report_no) && sztemp.indexOf("loginFlag=") != -1){//登入狀態 //110.10.08 add
				    	loginFlag = sztemp.substring(sztemp.indexOf("loginFlag=")+10,sztemp.length());				       
					}
				    if(isGetLoanItem){
    				    if(sztemp.indexOf("loan_item=") != -1){
    				       loan_item = sztemp.substring(sztemp.indexOf("loan_item=")+10,sztemp.length());                       
    	                }
				    }
				}//end of while	
				System.out.println("CANCEL_NO_file="+CANCEL_NO_file);
				System.out.println("HSIEN_ID_file="+HSIEN_ID_file);
				System.out.println("BankList_file="+BankList_file);
				System.out.println("btnFieldList_file="+btnFieldList_file);
				System.out.println("SortList_file="+SortList_file);
				System.out.println("SortBy_file="+SortBy_file);
				System.out.println("S_YEAR="+S_YEAR);
				System.out.println("S_MONTH="+S_MONTH);				
				System.out.println("E_YEAR="+E_YEAR);//95.12.04 add 結束日期
				System.out.println("E_MONTH="+E_MONTH);
				if("DS094W".equals(report_no)){
				System.out.println("S_DAY="+S_DAY);
				System.out.println("E_DAY="+E_DAY);
				System.out.println("sysType="+sysType);
				System.out.println("loginFlag="+loginFlag);
				}
				System.out.println("Unit="+Unit);
				if(isGetLoanItem)System.out.println("loan_item="+loan_item);
				//System.out.println("acc_div="+acc_div);
				
				session.setAttribute("CANCEL_NO",CANCEL_NO_file);
				session.setAttribute("HSIEN_ID",HSIEN_ID_file);
				session.setAttribute("BankList",BankList_file);
				session.setAttribute("btnFieldList",btnFieldList_file);		
				session.setAttribute("SortList",SortList_file);
				session.setAttribute("SortBy",SortBy_file);
				session.setAttribute("S_YEAR",S_YEAR);
				session.setAttribute("S_MONTH",S_MONTH);				
				session.setAttribute("E_YEAR",E_YEAR);//95.12.04 add 結束日期
				session.setAttribute("E_MONTH",E_MONTH);
				if("DS094W".equals(report_no)){
					session.setAttribute("S_DAY",S_DAY);
					session.setAttribute("E_DAY",E_DAY);		
					session.setAttribute("sysType",sysType);
					session.setAttribute("loginFlag",loginFlag);
				}
				session.setAttribute("Unit",Unit);
				if(isGetLoanItem)session.setAttribute("loan_item",loan_item);
				//session.setAttribute("acc_div",acc_div);				
			}else{//end of workfile exist
			   actMsg = "無已存之報表格式檔";
			}
		}catch(Exception e){
			System.out.println("readReport Error:"+e+e.getMessage());
			actMsg = "readReport Error:"+e+e.getMessage();			
		}	
		return actMsg;
    } 
    
    // 99.12.23 add讀取報表格式檔(BR基本報表用)========================================================
    //103.03.11 add 中文讀取時,轉utf-8
    public static String readReport_BR(HttpServletRequest request,String lguser_id,String report_no){
    	String sztemp="";
    	String S_YEAR_file = "";//99.12.23 add
    	String S_MONTH_file = "";//99.12.23 add
    	String CANCEL_NO_file = "";
        String HSIEN_ID_file = "";
        String BankList_file = "";
        String btnFieldList_file = "";
        String SortList_file = "";
        String SortBy_file = "";
        
        String actMsg = "";
        HttpSession session = request.getSession();      
    	try{
    		File WorkFile = new File(Utility.getProperties("profileDir")+System.getProperty("file.separator")+ lguser_id+"_"+report_no+".txt");	        	
			if(WorkFile.exists()){			
				FileInputStream fis = new FileInputStream(Utility.getProperties("profileDir")+System.getProperty("file.separator")+ lguser_id+"_"+report_no+".txt");	
				BufferedInputStream bis = new BufferedInputStream(fis);
				DataInputStream dis = new DataInputStream(fis);							
				while((sztemp = dis.readLine()) != null){
				    //sztemp = Utility.ISOtoBig5(sztemp);
				    if(sztemp.indexOf("S_YEAR=") != -1){//99.12.23 add
				       S_YEAR_file = sztemp.substring(sztemp.indexOf("S_YEAR=")+7,sztemp.length());				       
				    }
				    if(sztemp.indexOf("S_MONTH=") != -1){//99.12.23 add
				       S_MONTH_file = sztemp.substring(sztemp.indexOf("S_MONTH=")+8,sztemp.length());				       
				    }
				    if(sztemp.indexOf("CANCEL_NO=") != -1){
				       CANCEL_NO_file = sztemp.substring(sztemp.indexOf("CANCEL_NO=")+10,sztemp.length());				       
				    }
				    if(sztemp.indexOf("HSIEN_ID=") != -1){
				       HSIEN_ID_file = sztemp.substring(sztemp.indexOf("HSIEN_ID=")+9,sztemp.length());				       
				    }
				    if(sztemp.indexOf("BankList=") != -1){
				       BankList_file = sztemp.substring(sztemp.indexOf("BankList=")+9,sztemp.length());
				       BankList_file = Utility.ISOtoUTF8(BankList_file);//104.03.11 add
				    }
				    if(sztemp.indexOf("btnFieldList=") != -1){
				       btnFieldList_file = sztemp.substring(sztemp.indexOf("btnFieldList=")+13,sztemp.length());
				       btnFieldList_file = Utility.ISOtoUTF8(btnFieldList_file);//104.03.11 add
				    }	
				    if(sztemp.indexOf("SortList=") != -1){
				       SortList_file = sztemp.substring(sztemp.indexOf("SortList=")+9,sztemp.length());	
				       SortList_file = Utility.ISOtoUTF8(SortList_file);//104.03.11 add
				    }
				    if(sztemp.indexOf("SortBy=") != -1){
				       SortBy_file = sztemp.substring(sztemp.indexOf("SortBy=")+7,sztemp.length());				       
				    }
				    				
				}//end of while	
				System.out.println("S_YEAR_file="+S_YEAR_file);//99.12.23 add
				System.out.println("S_MONTH_file="+S_MONTH_file);//99.12.23 add
				System.out.println("CANCEL_NO_file="+CANCEL_NO_file);
				System.out.println("HSIEN_ID_file="+HSIEN_ID_file);
				System.out.println("BankList_file="+BankList_file);
				System.out.println("btnFieldList_file="+btnFieldList_file);
				System.out.println("SortList_file="+SortList_file);
				System.out.println("SortBy_file="+SortBy_file);
				
				session.setAttribute("S_YEAR",S_YEAR_file);//99.12.23 add
				session.setAttribute("S_MONTH",S_MONTH_file);//99.12.23 add
				session.setAttribute("CANCEL_NO",CANCEL_NO_file);
				session.setAttribute("HSIEN_ID",HSIEN_ID_file);
				session.setAttribute("BankList",BankList_file);
				session.setAttribute("btnFieldList",btnFieldList_file);		
				session.setAttribute("SortList",SortList_file);
				session.setAttribute("SortBy",SortBy_file);
				
			}else{//end of workfile exist
			   actMsg = "無已存之報表格式檔";
			}
		}catch(Exception e){
			System.out.println("readReport Error:"+e+e.getMessage());
			actMsg = "readReport Error";
		}	
		return actMsg;
    } 
    //95.11.03 add 刪除報表格式檔========================================================
    //ex:登入者帳號_程式名稱_範本名稱_建置日期.txt
    public static String deleteReport(HttpServletRequest request,String lguser_id/*建置者帳號*/,String template/*範本名稱*/,String nowDate/*範本建置日期*/,String report_no){    	
    	String filename="";		   	
        String actMsg = "";
        HttpSession session = request.getSession();              
        String bank_type = (session.getAttribute("nowbank_type")==null)?"":(String)session.getAttribute("nowbank_type");
    	try{
    	    filename = lguser_id+"_"+report_no+"_"+bank_type+"_"+template+"_"+nowDate;
    	    System.out.println("deleteReport="+filename);
    		File WorkFile = new File(Utility.getProperties("profileDir")+System.getProperty("file.separator")+filename);	        	
			if(WorkFile.exists()){							
			   if(!WorkFile.delete()){
			      actMsg = "刪除報表格式檔失敗";
			   }	
			}else{//end of workfile exist
			   actMsg = "無已存之報表格式檔";
			}
		}catch(Exception e){
			System.out.println("deleteReport Error:"+e+e.getMessage());
			actMsg = "readReport Error:"+e+e.getMessage();			
		}	
		return actMsg;
    }
    //95.11.03 add 取得範本資料====================================================================================== 
    public static List getTemplateList(String lguser_id,String bank_type,String report_no){
        List templateList = new LinkedList();
        try{        
        HashMap hm = new HashMap();
        File WorkFile = new File(Utility.getProperties("profileDir")+System.getProperty("file.separator"));	        	
        String[] fname = WorkFile.list();//====列出此目錄下的所有檔案===================
        File checkfile;		   
        List tmpFileName=null;
        String createName="";
        String template="";
        String createDate="";
        String tmpDate="";
        List wtt01 = null;
        //StringBuffer template_length = new StringBuffer("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;");
        String tmpTemplate = "";
        //StringBuffer createName_length = new StringBuffer("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;");
        String tmpCreateName = "";
        System.out.println("template.start="+lguser_id+"_"+report_no+"_"+bank_type);
		//=====判斷是否為目錄,若不是才加入範本檔案=============
		for(int c=0;c<fname.length;c++){	
		    checkfile = new File(Utility.getProperties("profileDir")+System.getProperty("file.separator")+fname[c]);
		    if(!checkfile.isDirectory()){	
		    	tmpFileName = Utility.getStringTokenizerData(fname[c],"_");     
		    	if(!fname[c].startsWith(lguser_id+"_"+report_no+"_"+bank_type)){
		    	   continue;
		    	}          
		    	//ex:filename = lguser_id+"_DS012W_"+bank_type+"_"+template+"_"+nowDate+".txt";
		    	if(tmpFileName.size() >= 5){    
		    	   wtt01 =  getWTT01((String)tmpFileName.get(0));
		    	   createName = (String)((DataObject)wtt01.get(0)).getValue("muser_name");//建罝者姓名
		    	   template = (String)tmpFileName.get(3);//範本名稱  
		    	   tmpDate = (String)tmpFileName.get(4);//建置日期  
		    	   System.out.println("tmpDate="+tmpDate);
		    	   System.out.println("tmpDate="+tmpDate.substring(0,4));
		    	   System.out.println("tmpDate="+tmpDate.substring(4,6));
		    	   System.out.println("tmpDate="+tmpDate.substring(6,8));
		    	   createDate = tmpDate.substring(0,4)+"/"+tmpDate.substring(4,6)+"/"+tmpDate.substring(6,8);
		    	   System.out.println("createDate="+createDate);
		    	   createDate = Utility.getCHTdate(createDate,1);//建置日期//ex: 2002/09/23->91年09月23日		    	   
		    	   System.out.println("template="+template);    	  
		    	   System.out.println("createName="+createName);    	  
		    	   System.out.println("createDate="+createDate); 
		    	   /*
		    	   tmpTemplate = template_length.replace(0,Utility.toBig5Convert(template).length(),template).toString(); 
		    	   System.out.println(tmpTemplate);
		    	   if(!Utility.toBig5Convert(tmpTemplate).substring(Utility.toBig5Convert(template).length(),Utility.toBig5Convert(tmpTemplate).indexOf(";")+1).equals("&nbsp;")){
		    	      tmpTemplate = tmpTemplate.substring(0,template.length())+Utility.toBig5Convert(tmpTemplate).substring(Utility.toBig5Convert(tmpTemplate).indexOf(";")+1,Utility.toBig5Convert(tmpTemplate).length());
		    	      System.out.println("showTemplate="+tmpTemplate);		    	      
		    	   }
		    	   tmpCreateName = createName_length.replace(0,Utility.toBig5Convert(createName).length(),createName).toString(); 
		    	   System.out.println(tmpCreateName);
		    	   System.out.println(Utility.toBig5Convert(createName).length());
		    	   System.out.println(Utility.toBig5Convert(tmpCreateName).indexOf(";"));
		    	   if(!Utility.toBig5Convert(tmpCreateName).substring(Utility.toBig5Convert(createName).length(),Utility.toBig5Convert(tmpCreateName).indexOf(";")+1).equals("&nbsp;")){
		    	      tmpCreateName = tmpCreateName.substring(0,createName.length())+Utility.toBig5Convert(tmpCreateName).substring(Utility.toBig5Convert(tmpCreateName).indexOf(";")+1,Utility.toBig5Convert(tmpCreateName).length());
		    	      System.out.println("showCreateUser"+tmpCreateName);		    	      
		    	   }
		    	   hm.put("showTemplate",tmpTemplate);//顯示的範本名稱
		    	   hm.put("template",template);//實際的範本名稱
		    	   hm.put("showCreateUser",tmpCreateName);//顯示的建置者名稱
		    	   hm.put("createUser",(String)tmpFileName.get(0));//建置者id
		    	   hm.put("showUpdateDate",createDate);//顯示的建置日期		
		    	   hm.put("updateDate",(String)tmpFileName.get(4));//實際的建置日期		
		    	   */
		    	   hm.put("showTemplate",template);//顯示的範本名稱
		    	   hm.put("template",template);//實際的範本名稱
		    	   hm.put("showCreateUser",createName);//顯示的建置者名稱
		    	   hm.put("createUser",(String)tmpFileName.get(0));//建置者id
		    	   hm.put("showUpdateDate",createDate);//顯示的建置日期		
		    	   hm.put("updateDate",(String)tmpFileName.get(4));//實際的建置日期		
		    	   
		    	   templateList.add(hm);
		    	   hm = new HashMap();
		    	   tmpFileName=null;
		    	}//end of tmpFileName   
		    }//end of 此檔不是目錄
		}//end of for    
		}catch(Exception e){
		    System.out.println("getTemplateList Error:"+e.getMessage());
		}
		return templateList;
    }
    
    //取得WTT01該使用者帳號資料
    public static List getWTT01(String muser_id){
            List paramList = new ArrayList();
    		//查詢條件        		
    		String sqlCmd = " select * from WTT01,MUSER_DATA "
    					  + " where WTT01.muser_id=?"      					    		    		  
    					  + " and WTT01.muser_id = MUSER_DATA.muser_id(+)";
    		paramList.add(muser_id);
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
            return dbData;
    } 
    
    public static void copyRows(HSSFSheet pSourceSheet, HSSFSheet pTargetSheet, int pStartRow,
            int pEndRow, int pPosition, HSSFWorkbook wb) {
           HSSFRow sourceRow = null;
           HSSFRow targetRow = null;
           HSSFCell sourceCell = null;
           HSSFCell targetCell = null;
           HSSFSheet sourceSheet = null;
           HSSFSheet targetSheet = null;
           Region region = null;
           int cType;
           int i;
           short j;
           int targetRowFrom;
           int targetRowTo;
           if ((pStartRow == -1) || (pEndRow == -1)) {
            return;
           }
           
           sourceSheet = pSourceSheet;
           targetSheet = pTargetSheet;
           //copy合併的單元格
           for (i = 0; i < sourceSheet.getNumMergedRegions(); i++) {
            region = sourceSheet.getMergedRegionAt(i);
            if ((region.getRowFrom() >= pStartRow)
              && (region.getRowTo() <= pEndRow)) {
             targetRowFrom = region.getRowFrom() - pStartRow + pPosition;
             targetRowTo = region.getRowTo() - pStartRow + pPosition;
             region.setRowFrom(targetRowFrom);
             region.setRowTo(targetRowTo);
             targetSheet.addMergedRegion(region);
            }
           }
           //設置列寬
           //如果是同一頁就不需要設置列寬，否則會有問題
           if (pSourceSheet != pTargetSheet) {
            for (i = pStartRow; i <= pEndRow; i++) {
             sourceRow = sourceSheet.getRow(i);
             if (sourceRow != null) {
              for (j = sourceRow.getFirstCellNum(); j < sourceRow
                .getLastCellNum(); j++) {
               targetSheet.setColumnWidth(j, sourceSheet
                 .getColumnWidth(j));
              }
              break;
             }
            }
           }
           //拷貝並填充數據
           for(i = pStartRow;i <= pEndRow; i++) {
               sourceRow = sourceSheet.getRow(i);
               if (sourceRow == null) {
                continue;
               }
               targetRow = targetSheet.createRow(i - pStartRow + pPosition);
               targetRow.setHeight(sourceRow.getHeight());
               for (j = sourceRow.getFirstCellNum(); j < sourceRow.getLastCellNum(); j++) {
                   sourceCell = sourceRow.getCell(j);
                   if (sourceCell == null) {
                    continue;
                   }
                   targetCell = targetRow.createCell(j);
                   targetCell.setEncoding(sourceCell.getEncoding());
                   targetCell.setCellStyle(sourceCell.getCellStyle());
                   cType = sourceCell.getCellType();
                   targetCell.setCellType(cType);
                   switch (cType) {
                      case HSSFCell.CELL_TYPE_BOOLEAN:
                       targetCell.setCellValue(sourceCell.getBooleanCellValue());
                       break;
                      case HSSFCell.CELL_TYPE_ERROR:
                       targetCell
                         .setCellErrorValue(sourceCell.getErrorCellValue());
                       break;
                      case HSSFCell.CELL_TYPE_FORMULA:
                       // parseFormula
                       targetCell.setCellFormula(parseFormula(sourceCell
                         .getCellFormula()));
                       break;
                      case HSSFCell.CELL_TYPE_NUMERIC:
                       targetCell.setCellValue(sourceCell.getNumericCellValue());
                       break;
                      case HSSFCell.CELL_TYPE_STRING:
                       targetCell.setCellValue(sourceCell.getStringCellValue());
                       break;
                   }//end of switch
               }//end of for
           }
    }
    //這是為了解決單元格中設置了函數的複製問題
     private static String parseFormula(String pPOIFormula) {
          final String cstReplaceString = "ATTR(semiVolatile)"; //$NON-NLS-1$
          StringBuffer result = null;
          int index;
          result = new StringBuffer();
          index = pPOIFormula.indexOf(cstReplaceString);
          if (index >= 0) {
           result.append(pPOIFormula.substring(0, index));
           result.append(pPOIFormula.substring(index
             + cstReplaceString.length()));
          } else {
           result.append(pPOIFormula);
          }
          return result.toString();
    }
     
    public static void insertRow(HSSFWorkbook wb, HSSFSheet sheet,  int starRow, int rows) { 
           sheet.shiftRows(starRow + 1, sheet.getLastRowNum(), rows); 
           starRow = starRow - 1;
           for (int i = 0; i < rows; i++) { 
                HSSFRow sourceRow = null;
                HSSFRow targetRow = null;
                HSSFCell sourceCell = null;
                HSSFCell targetCell = null;
                short m;
                starRow = starRow + 1;
                sourceRow = sheet.getRow(starRow);
                targetRow = sheet.createRow(starRow + 1);
                targetRow.setHeight(sourceRow.getHeight()); 
                for (m = sourceRow.getFirstCellNum(); m < sourceRow.getPhysicalNumberOfCells(); m++) {
                    sourceCell = sourceRow.getCell(m);
                    targetCell = targetRow.createCell(m);
                    targetCell.setEncoding(sourceCell.getEncoding());
                    targetCell.setCellStyle(sourceCell.getCellStyle());
                    targetCell.setCellType(sourceCell.getCellType());
                    
                }
           }       
    }
}
