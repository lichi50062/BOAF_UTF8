//105.07.11 add by 2968  
//110.08.25 fix 原993000農會盈餘取消申報,因與992120資料相同,改抓992120農農全體事業本期損益 by 2295
//111.01.26 調整 A05.910401/910402/910403無法取得資料問題 by 2295
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*;
import java.io.*;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR075W {
    public static String createRpt(String S_YEAR,String S_MONTH,String bank_code,String unit,String bank_type,String muser_id){
    	
    	String errMsg = "";
		StringBuffer sql = new StringBuffer () ;
		ArrayList paramList = new ArrayList() ;
		String unit_name=Utility.getUnitName(unit);
		FileInputStream finput = null;
		List baseData1 = new ArrayList();
		List baseData2 = new ArrayList();
		List dbData = new ArrayList();
		List dbData_1 = new ArrayList();
		List dbData_2 = new ArrayList();
		List dbData_3 = new ArrayList();
		List dbData_4 = new ArrayList();
		List dbData_5 = new ArrayList();
    	String bank_name="";
    	String u_year = "100" ;
		if(S_YEAR!=null && Integer.parseInt(S_YEAR) <= 99) {
			u_year = "99" ;
		}
		
		Utility.printLogTime("RptFR075W begin time");	
		try{	
			
			File xlsDir = new File(Utility.getProperties("xlsDir"));        
        	File reportDir = new File(Utility.getProperties("reportDir"));       
	    	        
    		if(!xlsDir.exists()){
     			if(!Utility.mkdirs(Utility.getProperties("xlsDir"))){
     		   		errMsg +=Utility.getProperties("xlsDir")+"目錄新增失敗";
     			}    
    		}
    		if(!reportDir.exists()){
     			if(!Utility.mkdirs(Utility.getProperties("reportDir"))){
     		   		errMsg +=Utility.getProperties("reportDir")+"目錄新增失敗";
     			}    
    		}
    		
            finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+"農漁會信用部財務經營報表.xls" );
			
			System.out.println(xlsDir + System.getProperty("file.separator")+"農漁會信用部財務經營報表.xls");
	  	    //設定FileINputStream讀取Excel檔
	  		POIFSFileSystem fs = new POIFSFileSystem( finput );
	  		HSSFWorkbook wb = new HSSFWorkbook(fs);
	  		HSSFSheet sheet = wb.getSheetAt(0);//讀取第一個工作表，宣告其為sheet 
	  		HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
	        //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
	        //sheet.setAutobreaks(true); //自動分頁
			
	        //設定頁面符合列印大小
	        //sheet.setAutobreaks( false );
	        ps.setScale( ( short )70 ); //列印縮放百分比
	        ps.setLandscape(false);//設定橫印
	       // HSSFFooter footer = sheet.getFooter();
	        ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
	  		//wb.setSheetName(0,"test");
	  		finput.close();
	  		
	  		HSSFRow row=null;//宣告一列 
	  		HSSFCell cell=null;//宣告一個儲存格  		
	  		//如果你需要使用換行符,你需要設置  
  		    //單元格的樣式wrap=true,代碼如下:  
	  		HSSFCellStyle cs = wb.createCellStyle();  
  		  	cs.setWrapText(true);  
  		  	cs.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
  		  	cs.setBorderTop(HSSFCellStyle.BORDER_THIN);
  		  	cs.setBorderLeft(HSSFCellStyle.BORDER_THIN);
  		  	cs.setBorderRight(HSSFCellStyle.BORDER_THIN);
  		  	cs.setBorderBottom(HSSFCellStyle.BORDER_THIN);
  		  	HSSFFont font = wb.createFont();
  		  	font.setFontHeightInPoints((short) 12);
  		  	cs.setFont(font);
  		  	
	  		//基本資料.總幹事/理事長/信用部主任/常務監事,若有值才將姓名填入報表欄位
	  		baseData1 = getBaseSQL1(bank_code);
	  		//基本資料.信用部員工/分部數
	  		baseData2 = getBaseSQL2(u_year,bank_code);
	  		/*
	  		 財務狀況(當月份及近5年12月底資料).N年X月/N-1年12月/N-2年12月/N-3年12月/N-4年12月/N05年12月
	  		*/
            dbData = getSQL(S_YEAR,S_MONTH,u_year,bank_type,bank_code,unit);//UI.查詢年月 
            dbData_1 = getSQL(String.valueOf(Integer.parseInt(S_YEAR)-1),"12",u_year,bank_type,bank_code,unit);//UI.查詢年度-1年12月
            dbData_2 = getSQL(String.valueOf(Integer.parseInt(S_YEAR)-2),"12",u_year,bank_type,bank_code,unit);//UI.查詢年度-2年12月
            dbData_3 = getSQL(String.valueOf(Integer.parseInt(S_YEAR)-3),"12",u_year,bank_type,bank_code,unit);//UI.查詢年度-3年12月
            dbData_4 = getSQL(String.valueOf(Integer.parseInt(S_YEAR)-4),"12",u_year,bank_type,bank_code,unit);//UI.查詢年度-4年12月
            dbData_5 = getSQL(String.valueOf(Integer.parseInt(S_YEAR)-5),"12",u_year,bank_type,bank_code,unit);//UI.查詢年度-5年12月
            
		  	//基本資料
  		  	if(baseData2.size()>0){
  		  		DataObject obj = (DataObject) baseData2.get(0);
  		  		bank_name = obj.getValue("bank_name")==null?"":obj.getValue("bank_name").toString();
	  		  	row=sheet.getRow(2);
			  	cell=row.getCell((short)5);//信用部員工
			  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			  	cell.setCellValue(Utility.setCommaFormat(obj.getValue("credit_staff_num")==null?"0":obj.getValue("credit_staff_num").toString()));
			  	
			  	row=sheet.getRow(3);
			  	cell=row.getCell((short)5);//分部數	       	
			  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			  	cell.setCellValue(Utility.setCommaFormat(obj.getValue("bank_count")==null?"0":obj.getValue("bank_count").toString()));
		  	}
  		  	row=sheet.getRow(0);
	  		cell=row.getCell((short)0);	       	
	  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	   	    cell.setCellValue(bank_name);
		  	row=sheet.getRow(0);
	  		cell=row.getCell((short)6);	       	
	  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	   	    cell.setCellValue("單位："+unit_name);
	   	    //基本資料
  		  	if(baseData1.size()>0){
  		  		for(int i=0;i<baseData1.size();i++){
  		  			DataObject obj = (DataObject) baseData1.get(i);
  		  			String cmuse_name = obj.getValue("cmuse_name")==null?"":obj.getValue("cmuse_name").toString();//職別
  		  			String name = obj.getValue("name")==null?"":obj.getValue("name").toString();
  		  			if(!"".equals(name)){
	  		  			if("總幹事".equals(cmuse_name)){
		  		  			row=sheet.getRow(2);
						  	cell=row.getCell((short)1);	       	
	  		  			}else if("理事長".equals(cmuse_name)){
	  		  				row=sheet.getRow(2);
		  		  			cell=row.getCell((short)3);       	
	  		  			}else if("信用部主任".equals(cmuse_name)){
		  		  			row=sheet.getRow(3);
						  	cell=row.getCell((short)1);	       	
	  		  			}else if("常務監事".equals(cmuse_name)){
	  		  				row=sheet.getRow(3);
		  		  			cell=row.getCell((short)3);      	
	  		  			}
		  		  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					  	cell.setCellValue(name);
  		  			}
  		  		}
  		  	}
  		  	row=sheet.getRow(5);
	  		cell=row.getCell((short)1);	       	
	  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	   	    cell.setCellValue(String.valueOf(Integer.parseInt(S_YEAR)-5)+"年12月");
	   	    row=sheet.getRow(5);
	  		cell=row.getCell((short)2);	       	
	  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	   	    cell.setCellValue(String.valueOf(Integer.parseInt(S_YEAR)-4)+"年12月");
	   	    row=sheet.getRow(5);
	  		cell=row.getCell((short)3);	       	
	  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	   	    cell.setCellValue(String.valueOf(Integer.parseInt(S_YEAR)-3)+"年12月");
	   	    row=sheet.getRow(5);
	  		cell=row.getCell((short)4);	       	
	  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	   	    cell.setCellValue(String.valueOf(Integer.parseInt(S_YEAR)-2)+"年12月");
	   	    row=sheet.getRow(5);
	  		cell=row.getCell((short)5);	       	
	  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	   	    cell.setCellValue(String.valueOf(Integer.parseInt(S_YEAR)-1)+"年12月");
		  	row=sheet.getRow(5);
	  		cell=row.getCell((short)6);	       	
	  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	   	    cell.setCellValue(S_YEAR+"年"+S_MONTH+"月");
	   	    DataObject obj = new DataObject();
  	        
	  		  	//N-5年12月
  	        	obj = new DataObject();
	  		  	if(dbData_5.size()>0) {
	  		  		obj = (DataObject) dbData_5.get(0);
	  		  	}
	  		  	setVal(sheet,row,cell,obj,(short)1,cs);
			  
	  		  	//N-4年12月
	  		  	obj = new DataObject();
	  		  	if(dbData_4.size()>0){	
	  		  		obj = (DataObject) dbData_4.get(0);
			  	}
	  		  	setVal(sheet,row,cell,obj,(short)2,cs);
	  		  	
	  		  	//N-3年12月
	  		  	obj = new DataObject();
	  		  	if(dbData_3.size()>0){	
	  		  		obj = (DataObject) dbData_3.get(0);
	  		  	}
	  		  	setVal(sheet,row,cell,obj,(short)3,cs);
	  		  	
	  		  	//N-2年12月
	  		  	obj = new DataObject();
	  		  	if(dbData_2.size()>0){	
	  		  		obj = (DataObject) dbData_2.get(0);
			  	}
	  		  	setVal(sheet,row,cell,obj,(short)4,cs);
	  		  	
	  		  	//N-1年12月
	  		  	obj = new DataObject();
	  		  	if(dbData_1.size()>0){	
	  		  		obj = (DataObject) dbData_1.get(0);
			  	}
	  		  	setVal(sheet,row,cell,obj,(short)5,cs);
	  		  
	  		  	//N年○月
	  		  	obj = new DataObject();
	  		  	if(dbData.size()>0){	
	  		  		obj = (DataObject) dbData.get(0);
	  		  	}
	  		  	setVal(sheet,row,cell,obj,(short)6,cs);
		  		//放款依對象分
	  		  	row=sheet.getRow(27);
			  	cell=row.getCell((short)1);//會員-金額	       	
			  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			  	cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_992140")==null?"0":obj.getValue("field_992140").toString()));
			  	cell=row.getCell((short)2);//會員-比率	       	
			  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			  	cell.setCellValue(obj.getValue("field_992140_rate")==null?0.00:Double.parseDouble(obj.getValue("field_992140_rate").toString()));
			  	cell=row.getCell((short)5);//全國農業金庫股票-金額	       	
			  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			  	cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_910401")==null?"0":obj.getValue("field_910401").toString()));
			  	cell=row.getCell((short)6);//全國農業金庫股票-比率	       	
			  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			  	cell.setCellValue(obj.getValue("field_910401_rate")==null?0.00:Double.parseDouble(obj.getValue("field_910401_rate").toString()));
			  	
			  	row=sheet.getRow(28);
			  	cell=row.getCell((short)1);//贊助會員-金額	       	
			  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			  	cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_990410")==null?"0":obj.getValue("field_990410").toString()));
			  	cell=row.getCell((short)2);//贊助會員-比率	       	
			  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			  	cell.setCellValue(obj.getValue("field_990410_rate")==null?0.00:Double.parseDouble(obj.getValue("field_990410_rate").toString()));
			  	cell=row.getCell((short)5);//合作金庫股票-金額	       	
			  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			  	cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_910403")==null?"0":obj.getValue("field_910403").toString()));
			  	cell=row.getCell((short)6);//合作金庫股票-比率	       	
			  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			  	cell.setCellValue(obj.getValue("field_910403_rate")==null?0.00:Double.parseDouble(obj.getValue("field_910403_rate").toString()));
			  	
			  	row=sheet.getRow(29);
			  	cell=row.getCell((short)1);//非會員-金額	       	
			  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			  	cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_990610_990611")==null?"0":obj.getValue("field_990610_990611").toString()));
			  	cell=row.getCell((short)2);//非會員-比率	       	
			  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			  	cell.setCellValue(obj.getValue("field_990610_rate")==null?0.00:Double.parseDouble(obj.getValue("field_990610_rate").toString()));
			  	cell=row.getCell((short)5);//財金資訊(股)公司股票-金額	       	
			  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			  	cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_910402")==null?"0":obj.getValue("field_910402").toString()));
			  	cell=row.getCell((short)6);//財金資訊(股)公司股票-比率	       	
			  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			  	cell.setCellValue(obj.getValue("field_910402_rate")==null?0.00:Double.parseDouble(obj.getValue("field_910402_rate").toString()));
			  	
			  	row=sheet.getRow(30);
			  	cell=row.getCell((short)1);//總計-金額	       	
			  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			  	cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_sum_1")==null?"0":obj.getValue("field_sum_1").toString()));
			  	cell=row.getCell((short)2);//總計-比率	       	
			  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			  	cell.setCellValue(obj.getValue("field_sum1_rate")==null?0.00:Double.parseDouble(obj.getValue("field_sum1_rate").toString()));
			  	cell=row.getCell((short)5);//其他(含基金)-金額	       	
			  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			  	cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_other")==null?"0":obj.getValue("field_other").toString()));
			  	cell=row.getCell((short)6);//其他(含基金)-比率	       	
			  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			  	cell.setCellValue(obj.getValue("field_other_rate")==null?0.00:Double.parseDouble(obj.getValue("field_other_rate").toString()));	
	  		  	
	  	    
  	        
  	        HSSFFooter footer=sheet.getFooter();
			footer.setCenter( "Page:"+HSSFFooter.page()+" of "+HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
			
	        FileOutputStream fout=new FileOutputStream(reportDir + System.getProperty("file.separator")+"農漁會信用部財務經營報表.xls");
	        wb.write(fout);
	        //儲存 
	        fout.close();
	        Utility.printLogTime("RptFR075W end time");	
		}catch(Exception e){
			System.out.println("createRpt Error:"+e+e.getMessage());
		}
		return errMsg;
	}
    private static void setVal(HSSFSheet sheet,HSSFRow row,HSSFCell cell,DataObject obj,short cNum,HSSFCellStyle cs){
    		row=sheet.getRow(6);//存款
		  	cell=row.getCell(cNum);	       	
		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		    cell.setCellStyle(cs);
		  	cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_debit")==null?"0":obj.getValue("field_debit").toString()));
		  	row=sheet.getRow(7);//放款
		  	cell=row.getCell(cNum);	       	
		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		    cell.setCellStyle(cs);
		    cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_credit")==null?"0":obj.getValue("field_credit").toString()));
		  	row=sheet.getRow(8);//存放比率
		  	cell=row.getCell(cNum);	       	
		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		  	cell.setCellValue(obj.getValue("field_dc_rate")==null?0.00:Double.parseDouble(obj.getValue("field_dc_rate").toString()));
		  	row=sheet.getRow(9);//資本適足率
		  	cell=row.getCell(cNum);	       	
		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		  	cell.setCellValue(obj.getValue("field_captial_rate")==null?0.00:Double.parseDouble(obj.getValue("field_captial_rate").toString()));
		    row=sheet.getRow(10);//備抵呆帳覆蓋率
		  	cell=row.getCell(cNum);	       	
		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		  	cell.setCellValue(obj.getValue("field_backup_over_rate")==null?0.00:Double.parseDouble(obj.getValue("field_backup_over_rate").toString()));
		    row=sheet.getRow(11);//放款覆蓋率
		  	cell=row.getCell(cNum);	       	
		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		  	cell.setCellValue(obj.getValue("field_backup_credit_rate")==null?0.00:Double.parseDouble(obj.getValue("field_backup_credit_rate").toString()));
		  	row=sheet.getRow(12);//備抵呆帳
		  	cell=row.getCell(cNum);	       	
		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		  	cell.setCellStyle(cs);
		  	cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_backup")==null?"0":obj.getValue("field_backup").toString()));
		  	row=sheet.getRow(13);//逾放金額
		  	cell=row.getCell(cNum);	       	
		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		  	cell.setCellStyle(cs);
		  	cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_over")==null?"0":obj.getValue("field_over").toString()));
		  	row=sheet.getRow(14);//逾放比率
		  	cell=row.getCell(cNum);	       	
		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		  	cell.setCellValue(obj.getValue("field_over_rate")==null?0.00:Double.parseDouble(obj.getValue("field_over_rate").toString()));
		  	row=sheet.getRow(15);//信用部盈餘
		  	cell=row.getCell(cNum);	       	
		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		  	cell.setCellStyle(cs);
		  	cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_320300")==null?"0":obj.getValue("field_320300").toString()));
		  	row=sheet.getRow(16);//農會盈餘
		  	cell=row.getCell(cNum);	       	
		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		  	cell.setCellStyle(cs);
		  	cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_992120")==null?"0":obj.getValue("field_992120").toString()));
		  	row=sheet.getRow(17);//信用部盈餘占農漁會比率
		  	cell=row.getCell(cNum);	       	
		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		  	cell.setCellValue(obj.getValue("field_320300_rate")==null?0.00:Double.parseDouble(obj.getValue("field_320300_rate").toString()));
		  	row=sheet.getRow(18);//信用部淨值
		  	cell=row.getCell(cNum);	       	
		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		  	cell.setCellStyle(cs);
		  	cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_net")==null?"0":obj.getValue("field_net").toString()));
		  	row=sheet.getRow(19);//農漁會淨值
		  	cell=row.getCell(cNum);	       	
		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		  	cell.setCellStyle(cs);
		  	cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_992110")==null?"0":obj.getValue("field_992110").toString()));
		  	row=sheet.getRow(20);//信用部淨值占農漁會比率
		  	cell=row.getCell(cNum);	       	
		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		  	cell.setCellValue(obj.getValue("field_net_rate")==null?0.00:Double.parseDouble(obj.getValue("field_net_rate").toString()));
		  	row=sheet.getRow(21);//內部融資
		  	cell=row.getCell(cNum);	       	
		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		  	cell.setCellStyle(cs);
		  	cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_120700")==null?"0":obj.getValue("field_120700").toString()));
		  	row=sheet.getRow(22);//建築放款餘額
		  	cell=row.getCell(cNum);	       	
		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		  	cell.setCellStyle(cs);
		  	cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_992710")==null?"0":obj.getValue("field_992710").toString()));
		  	row=sheet.getRow(23);//專案農貸餘額
		  	cell=row.getCell(cNum);	       	
		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		  	cell.setCellStyle(cs);
		  	cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_120600")==null?"0":obj.getValue("field_120600").toString()));
		  	row=sheet.getRow(24);//活存比率
		  	cell=row.getCell(cNum);	       	
		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		  	cell.setCellValue(obj.getValue("field_debit_rate")==null?0.00:Double.parseDouble(obj.getValue("field_debit_rate").toString()));
    	
    }
    //基本資料.總幹事/理事長/信用部主任/常務監事,若有值才將姓名填入報表欄位
    private static List getBaseSQL1(String bank_code){
    	StringBuffer sqlCmd = new StringBuffer () ;
		ArrayList paramList = new ArrayList() ;
		sqlCmd.append("select cmuse_name,");//--職別
		sqlCmd.append("       name ");//--姓名
		sqlCmd.append(" from WLX01_M,cdshareno where bank_no=? ");
		sqlCmd.append("  and WLX01_M.POSITION_CODE = cdshareno.CMUSE_ID and cdshareno.CMUSE_DIV='005' ");
		sqlCmd.append("  and abdicate_code !='Y' ");
		sqlCmd.append("order by position_code,abdicate_date ");
		paramList.add(bank_code) ;
        List dbData =DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList, "cmuse_name,name");
    	System.out.println("******Base1 dbData.size()="+dbData.size());
    	return dbData;
    }
    //基本資料.信用部員工/分部數
    private static List getBaseSQL2(String u_year,String bank_code){
    	StringBuffer sqlCmd = new StringBuffer () ;
		ArrayList paramList = new ArrayList() ;
		sqlCmd.append("select ba01.bank_no,");
		sqlCmd.append("	      ba01.bank_name,");//--農漁信用部名稱
		sqlCmd.append("       credit_staff_num,");//--信用部員工x人
		sqlCmd.append("       bank_count ");//--分部數
		sqlCmd.append("from (select * from wlx01 where m_year=?)wlx01 "); 
		sqlCmd.append("left join (select * from ba01 where m_year=? and bank_type in ('6','7') and bank_kind='0')ba01 on wlx01.bank_no = ba01.bank_no "); 
		sqlCmd.append("left join (select ba01.pbank_no,count(*) as bank_count ");
		sqlCmd.append("             from (select * from wlx02 where m_year=?)wlx02 ");
		sqlCmd.append("		    left join (select * from ba01 where m_year=? and bank_type in ('6','7') and bank_kind='1')ba01 on wlx02.tbank_no = ba01.pbank_no and wlx02.bank_no=ba01.bank_no ");
		sqlCmd.append("            where cancel_no != 'Y' ");
		sqlCmd.append("            group by pbank_no)wlx02 on wlx01.bank_no=wlx02.pbank_no ");
		sqlCmd.append("where wlx01.bank_no=? ");
		sqlCmd.append("and cancel_no != 'Y' ");
		paramList.add(u_year) ;
		paramList.add(u_year) ;
		paramList.add(u_year) ;
		paramList.add(u_year) ;
		paramList.add(bank_code) ;
        List dbData =DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList, "bank_no,bank_name,credit_staff_num,bank_count");
    	System.out.println("******Base2 dbData.size()="+dbData.size());
    	return dbData;
    }
    //財務狀況.N年X月/N-1年12月/N-2年12月/N-3年12月/N-4年12月/N05年12月,依據查詢年月代入條件
    private static List getSQL(String S_YEAR,String S_MONTH,String u_year,String bank_type,String bank_code,String unit){
    	StringBuffer sqlCmd = new StringBuffer () ;
		ArrayList paramList = new ArrayList() ;
		sqlCmd.append("select bank_no,bank_name,");
		sqlCmd.append("      round(field_DEBIT/?,0) as field_DEBIT,");//--存款
		sqlCmd.append("      round(field_CREDIT/?,0) as field_CREDIT,");//--放款
		sqlCmd.append("       field_dc_rate,");//--存放比率
		sqlCmd.append("       round(field_CAPTIAL /  1000 ,2)  as   field_CAPTIAL_RATE,");//--資本適足率
		sqlCmd.append("       decode(a01.field_OVER,0,0,round(a01.field_BACKUP /  a01.field_OVER *100 ,2)) as   field_BACKUP_OVER_RATE,");//--備抵呆帳覆蓋率=備抵呆帳/逾期放款
		sqlCmd.append("       decode(a01.field_CREDIT,0,0,round(a01.field_BACKUP /  a01.field_CREDIT *100 ,2)) as   field_BACKUP_CREDIT_RATE,");//--放款覆蓋率=備抵呆帳/放款
		sqlCmd.append("       round(field_BACKUP/?,0) as field_BACKUP,");//--備抵呆帳
		sqlCmd.append("       round(field_OVER/?,0) as field_OVER,");//--逾放金額
		sqlCmd.append("       decode(a01.field_CREDIT,0,0,round(a01.field_OVER /  a01.field_CREDIT *100 ,2))  as   field_OVER_RATE,");//--逾放比率
		sqlCmd.append("       round(field_320300/?,0) as field_320300,");//--信用部盈餘
		sqlCmd.append("       round(field_992120/?,0) as field_992120,");//--農會盈餘 //110.08.25原993000農會盈餘取消申報,因與992120資料相同,改抓992120農農全體事業本期損益 
		sqlCmd.append("       decode(field_992120,0,0,round(field_320300 / field_992120 *100 ,2))  as   field_320300_RATE,");//--信用部盈餘占農漁會比率
		sqlCmd.append("       round(field_NET/?,0) as field_NET,");//--信用部淨值
		sqlCmd.append("       round(field_992110/?,0) as field_992110,");//--農漁會淨值
		sqlCmd.append("       decode(field_992110,0,0,round(field_NET /field_992110 *100 ,2))  as   field_NET_RATE,");//--信用部淨值占農漁會比率
		sqlCmd.append("       round(field_120700/?,0) as field_120700,");//--內部融資
		sqlCmd.append("       round(field_992710/?,0) as field_992710,");//--建築放款餘額
		sqlCmd.append("       round(field_120600/?,0) as field_120600,");//--專案農貸餘額
		sqlCmd.append("       decode(field_DEBIT,0,0,round((field_220100+field_220200+field_220300+field_220400+field_220500) / field_DEBIT *100 ,2))  as field_DEBIT_RATE,");//--活存比率
		sqlCmd.append("       round(field_992140/?,0) as field_992140,");//--會員
		sqlCmd.append("       decode(field_sum_1,0,0,round(field_992140 / field_sum_1 *100 ,2)) as   field_992140_RATE,");//--會員比率
		sqlCmd.append("       round(field_990410/?,0) as field_990410,");//--贊助會員
		sqlCmd.append("       decode(field_sum_1,0,0,round(field_990410 / field_sum_1 *100 ,2)) as   field_990410_RATE,");//--贊助會員比率
		sqlCmd.append("       round((field_990610_990611)/?,0) as field_990610_990611,");//--非會員
		sqlCmd.append("       decode(field_sum_1,0,0,round(field_990610_990611 / field_sum_1 *100 ,2)) as   field_990610_RATE,");//--非會員比率
		sqlCmd.append("       round(field_sum_1/?,0) as field_sum_1,");//--放款依對象總計
		sqlCmd.append("       decode(field_sum_1,0,0,round(field_sum_1 / field_sum_1 *100 ,2)) as   field_sum1_RATE,");//--放款依對象總計比率
		sqlCmd.append("       round(field_910401/?,0) as field_910401,");//--全國農業金庫股票
		sqlCmd.append("       decode(field_sum_2,0,0,round(field_910401 / field_sum_2 *100 ,2)) as   field_910401_RATE,");//--全國農業金庫股票比率
		sqlCmd.append("       round(field_910403/?,0) as field_910403,");//--合作金庫股票
		sqlCmd.append("       decode(field_sum_2,0,0,round(field_910403 / field_sum_2 *100 ,2)) as   field_910403_RATE,");//--合作金庫股票比率
		sqlCmd.append("       round(field_910402/?,0) as field_910402,");//--財金資訊(股)公司股票
		sqlCmd.append("       decode(field_sum_2,0,0,round(field_910402 / field_sum_2 *100 ,2)) as   field_910402_RATE,");//--財金資訊(股)公司股票比率
		sqlCmd.append("       round(field_other/?,0) as field_other,");//--其他(含基金)
		sqlCmd.append("       decode(field_sum_2,0,0,round(field_other / field_sum_2 *100 ,2)) as   field_other_RATE ");//--其他(含基金)比率 
		for(int i=0;i<19;i++){
        	paramList.add(unit) ;
        }
		sqlCmd.append("from ");
		sqlCmd.append("(    ");
		sqlCmd.append("  select a01.bank_no,a01.bank_name,");
		sqlCmd.append("         field_DEBIT,field_CREDIT,field_dc_rate,field_CAPTIAL,field_BACKUP,field_OVER,field_320300,field_992120,field_NET,");
		sqlCmd.append("         field_992110,field_120700,field_992710,field_120600, field_220100,field_220200,field_220300,field_220400,field_220500,");
		sqlCmd.append("         field_992140,field_990410,round((field_990610-field_990611)/1,0) as field_990610_990611,");//--非會員
		sqlCmd.append("         round((field_992140+field_990410+field_990610-field_990611)/1,0) as field_sum_1,");//--放款依對象分總計
		sqlCmd.append("         field_910401,field_910403,field_910402,");
		sqlCmd.append("         round((field_130200-field_910401-field_910403-field_910402+field_130100)/1,0) as field_other,");//--其他(含基金)
		sqlCmd.append("         round((field_130200+field_130100)/1,0) as field_sum_2 ");//--投資業務合計
		sqlCmd.append("  from  (select a01.bank_no , a01.BANK_NAME,");
		sqlCmd.append("          SUM(field_DEBIT)  field_DEBIT , SUM(field_CREDIT) field_CREDIT, SUM(field_BACKUP)  field_BACKUP, SUM(field_OVER)   field_OVER,");
		sqlCmd.append("          SUM(field_320300) field_320300, SUM(field_NET)    field_NET,    SUM(field_120700)  field_120700, SUM(field_120600) field_120600,");
		sqlCmd.append("          SUM(field_220100) field_220100, SUM(field_220200) field_220200, SUM(field_220300)  field_220300, SUM(field_220400) field_220400,");
		sqlCmd.append("          SUM(field_220500) field_220500, SUM(field_130100) field_130100, SUM(field_130200)  field_130200, SUM(field_990410) field_990410,");
		sqlCmd.append("          SUM(field_990610) field_990610, SUM(field_990611) field_990611, SUM(field_CAPTIAL) field_CAPTIAL,SUM(field_910401) field_910401,");
		sqlCmd.append("          SUM(field_910402) field_910402, SUM(field_910403) field_910403, SUM(field_992120)  field_992120, SUM(field_992110) field_992110,");
		sqlCmd.append("          SUM(field_992710) field_992710, SUM(field_992140) field_992140, SUM(field_dc_rate) field_dc_rate ");
		sqlCmd.append("   from ");
		sqlCmd.append("  (select bn01.bank_no , bn01.BANK_NAME,");
		sqlCmd.append("          round(sum(decode(a01.acc_code,'220000',amt,0)) /1,0) as field_DEBIT,");//--存款
		sqlCmd.append("          round(sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /1,0) as  field_CREDIT,");//--放款
		sqlCmd.append("          round(sum(decode(a01.acc_code, '120800',amt,'150300',amt,0)) /1,0) as  field_BACKUP,");//--備抵呆帳
		sqlCmd.append("          round(sum(decode(a01.acc_code,'990000',amt,0)) /1,0) as field_OVER,");//--逾期放款
		sqlCmd.append("          round(sum(decode(a01.acc_code,'320300',amt,0)) /1,0) as field_320300,");//--信用部盈餘
		sqlCmd.append("          round(sum(decode(bn01.bank_type,'6',decode(a01.acc_code,'310000',amt,'320000',amt,0),'7',decode(a01.acc_code,'300000',amt,0),0)) /1,0) as field_NET,");//--淨值
		sqlCmd.append("          round(sum(decode(a01.acc_code,'120700',amt,0)) /1,0) as field_120700,");//--內部融資
		sqlCmd.append("          round(sum(decode(a01.acc_code,'120600',amt,0)) /1,0) as field_120600,");//--專案農貸餘額
		sqlCmd.append("          round(sum(decode(a01.acc_code,'220100',amt,0)) /1,0) as field_220100,");
		sqlCmd.append("          round(sum(decode(a01.acc_code,'220200',amt,0)) /1,0) as field_220200,");
		sqlCmd.append("          round(sum(decode(a01.acc_code,'220300',amt,0)) /1,0) as field_220300,");
		sqlCmd.append("          round(sum(decode(a01.acc_code,'220400',amt,0)) /1,0) as field_220400,");
		sqlCmd.append("          round(sum(decode(a01.acc_code,'220500',amt,0)) /1,0) as field_220500,");
		sqlCmd.append("          round(sum(decode(a01.acc_code,'130200',amt,0)) /1,0) as field_130200,");
		sqlCmd.append("          round(sum(decode(a01.acc_code,'130100',amt,0)) /1,0) as field_130100 ");
		sqlCmd.append("  from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
		sqlCmd.append("        left join (select (CASE WHEN (a01.m_year <= 102) THEN '102' ");
		sqlCmd.append("                                WHEN (a01.m_year > 102) THEN '103' ");
		sqlCmd.append("                                ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01 ");
		sqlCmd.append("                   where m_year=? and m_month=? ");
		sqlCmd.append("                   ) a01  on  bn01.bank_no = a01.bank_code ");
		sqlCmd.append("        where a01.bank_code=? ");
		paramList.add(u_year) ;
		paramList.add(S_YEAR) ;
        paramList.add(S_MONTH) ;
		paramList.add(bank_code) ;
		sqlCmd.append("        group by a01.m_year,a01.m_month,bn01.bank_no,bn01.BANK_NAME ");
		sqlCmd.append("  )a01,");
		sqlCmd.append("  (select bn01.bank_no as bank_code, bn01.BANK_NAME,");
		sqlCmd.append("          round(sum(decode(acc_code,'990410',amt,0)) /1,0) as field_990410,");
		sqlCmd.append("          round(sum(decode(acc_code,'990610',amt,0)) /1,0) as field_990610,");
		sqlCmd.append("          round(sum(decode(acc_code,'990611',amt,0)) /1,0) as field_990611 ");
		sqlCmd.append("   from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
		sqlCmd.append("   left join (select * from a02 where m_year=? and m_month=?)a02 on bn01.bank_no = a02.bank_code ");
		sqlCmd.append("   where bank_code=? ");
		paramList.add(u_year) ;
		paramList.add(S_YEAR) ;
        paramList.add(S_MONTH) ;
		paramList.add(bank_code) ;
		sqlCmd.append("   group by bn01.bank_no,bn01.BANK_NAME ");
		sqlCmd.append("  )a02,");
		sqlCmd.append("  (select bn01.bank_no as bank_code,  bn01.bank_name,");
		sqlCmd.append("          round(sum(decode(a05.acc_code,'91060P',amt,0)) /1,0) as field_CAPTIAL,");
		sqlCmd.append("          round(sum(decode(a05.acc_code,'910401',amt,0)) /1,0) as field_910401,");
		sqlCmd.append("          round(sum(decode(a05.acc_code,'910402',amt,0)) /1,0) as field_910402,");
		sqlCmd.append("          round(sum(decode(a05.acc_code,'910403',amt,0)) /1,0) as field_910403 ");
		sqlCmd.append("   from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
		sqlCmd.append("   left join (select * from a05 where m_year=? and m_month=? ) a05 on  bn01.bank_no = a05.bank_code ");
		sqlCmd.append("   where a05.bank_code=? ");
		paramList.add(u_year) ;
		paramList.add(S_YEAR) ;
        paramList.add(S_MONTH) ;
		paramList.add(bank_code) ;
		sqlCmd.append("   group by bn01.bank_no,bn01.BANK_NAME ");
		sqlCmd.append("  )a05,");
		sqlCmd.append("  (select bn01.bank_no as bank_code, bn01.BANK_NAME,");
		sqlCmd.append("          round(sum(decode(a99.acc_code,'992120',amt,0)) /1,0) as field_992120,");//110.08.25原993000農會盈餘取消申報,因與992120資料相同,改抓992120農農全體事業本期損益
		sqlCmd.append("          round(sum(decode(a99.acc_code,'992110',amt,0)) /1,0) as field_992110,");
		sqlCmd.append("          round(sum(decode(a99.acc_code,'992710',amt,0)) /1,0) as field_992710,");
		sqlCmd.append("          round(sum(decode(a99.acc_code,'992140',amt,0)) /1,0) as field_992140 ");
		sqlCmd.append("   from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
		sqlCmd.append("   left join (select * from a99 where m_year=? and m_month=?)a99 on bn01.bank_no = a99.bank_code ");
		sqlCmd.append("   where bank_code=? ");
		paramList.add(u_year) ;
		paramList.add(S_YEAR) ;
        paramList.add(S_MONTH) ;
		paramList.add(bank_code) ;
		sqlCmd.append("   group by bn01.bank_no,bn01.BANK_NAME ");
		sqlCmd.append("  )a99,");
		sqlCmd.append("  (select bn01.bank_no as bank_code, bn01.BANK_NAME,");
		sqlCmd.append("          sum(decode(acc_code,'field_dc_rate',amt,0))  as field_dc_rate ");
		sqlCmd.append("  from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
		sqlCmd.append("  left join (select * from a01_operation where m_year=? and m_month=?)a01_operation on bn01.bank_no = a01_operation.bank_code ");
		sqlCmd.append("  where bank_code=? ");
		paramList.add(u_year) ;
		paramList.add(S_YEAR) ;
        paramList.add(S_MONTH) ;
		paramList.add(bank_code) ;
		sqlCmd.append("  group by bn01.bank_no,bn01.BANK_NAME ");
		sqlCmd.append("  )a01_operation ");
		sqlCmd.append("  where a01.bank_no = a02.bank_code(+) and a01.bank_no = a05.bank_code(+)  and a01.bank_no=a99.bank_code(+) and a01.bank_no=a01_operation.bank_code(+) ");
		sqlCmd.append("  and a01.bank_no=? ");
		paramList.add(bank_code) ;
		sqlCmd.append("  GROUP BY a01.bank_no,a01.BANK_NAME ");
		sqlCmd.append("  )a01 ");
		sqlCmd.append(")a01 ");

        List dbData =DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList, 
        		"field_debit,field_credit,field_dc_rate,field_captial_rate,field_backup_over_rate,field_backup_credit_rate,"
        		+"field_backup,field_over,field_over_rate,field_320300,field_992120,field_320300_rate,field_992110,field_net,field_net_rate,"
        		+"field_120700,field_992710,field_120600,field_debit_rate,field_992140,field_992140_rate,field_990410,field_990410_rate,"
        		+"field_990610_990611,field_990610_rate,field_sum_1,field_sum1_rate,field_910401,field_910401_rate,field_910403,"
        		+"field_910403_rate,field_910402,field_910402_rate,field_other,field_other_rate");
    	System.out.println("*****"+S_YEAR+"年"+S_MONTH+"月  dbData.size()="+dbData.size());
    	return dbData;
    }
}
