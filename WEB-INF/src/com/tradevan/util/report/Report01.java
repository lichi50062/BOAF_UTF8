/*
 * Created on 2005/1/11
 *
 * TODO To change the template for this generated file go to
 * Window - Preferences - Java - Code Style - Code Templates
 */
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.poifs.filesystem.*;
import java.io.*;
import java.text.*;
import java.util.*;
import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;
/**
 * @author 2295
 *
 * TODO To change the template for this generated type comment go to
 * Window - Preferences - Java - Code Style - Code Templates
 */
public class Report01 {
    public static String createRpt(String m_year,String m_month,String bank_code){
		String errMsg = "";
		List dbData = null;
		String sqlCmd = "";
		Properties A01Data = new Properties();		
		String acc_code = "";
		String amt = "";		
		reportUtil reportUtil = new reportUtil();
		try{	
			System.out.println("信用部損益表.xls");
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
			FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+"信用部損益表.xls" );
			
	  	    //設定FileINputStream讀取Excel檔
	  		POIFSFileSystem fs = new POIFSFileSystem( finput );
	  		HSSFWorkbook wb = new HSSFWorkbook(fs);
	  		HSSFSheet sheet = wb.getSheetAt(0);//讀取第一個工作表，宣告其為sheet 
	  		HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
	  		//HSSFCellStyle cs = wb.createCellStyle();
	  		//HSSFDataFormat df = wb.createDataFormat();
	  		//cs.setDataFormat(df.getFormat("###,###,##0"));

	        //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
	        //sheet.setAutobreaks(true); //自動分頁
			
	        //設定頁面符合列印大小
	        sheet.setAutobreaks( false );
	        ps.setScale( ( short )65 ); //列印縮放百分比

	        ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
	  		//wb.setSheetName(0,"test");
	  		finput.close();
	  		
	  		HSSFRow row=null;//宣告一列 
	  		HSSFCell cell=null;//宣告一個儲存格  		
	  		
	  		short i=0;
	  		short y=0;	  
	  		
	  		sqlCmd = " select * from A01 LEFT JOIN ncacno ON A01.acc_code = ncacno.acc_code" 
	  			   + " where A01.m_year="+m_year 
				   + "   and A01.m_month="+m_month 
				   + "   and A01.bank_code='"+bank_code+"'" 
				   + "   and ncacno.acc_div='02'" 
				   + " order by ncacno.acc_range";
	  		dbData = DBManager.QueryDB(sqlCmd,"m_year,m_month,amt");	  	         

  			for(i=0;i<dbData.size();i++){
  				acc_code = (String)((DataObject)dbData.get(i)).getValue("acc_code");
  				amt = (((DataObject)dbData.get(i)).getValue("amt")).toString();
	       	    A01Data.setProperty(acc_code,amt);
	       	    System.out.println("acc_code="+acc_code+":amt="+amt);
	       	}
  			
	  		row=sheet.getRow(1);
	  		cell=row.getCell((short)0);	       	
	  		//設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
	  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		//===========================================================
	       	cell.setCellValue("中華民國      94     年         1        月                1  日至    94   年       1    月      31   日");	       	
			
	  		//以巢狀迴圈讀取所有儲存格資料 
	  		System.out.println("total row ="+sheet.getLastRowNum());
	  		for(i=3;i<=sheet.getLastRowNum();i++){    		
	    		row=sheet.getRow(i);	    		
	       		cell=row.getCell((short)1);
	       		System.out.print((int)cell.getNumericCellValue()+"=");
	       		amt = A01Data.getProperty(String.valueOf((int)cell.getNumericCellValue()));
	       		cell=row.getCell((short)2);
	       		//cell.setCellStyle(cs);
	       		cell.setEncoding( HSSFCell.ENCODING_UTF_16 );	       			       		
	       		cell.setCellValue(amt);
	       		cell=row.getCell((short)4);	       			 
	       		amt = A01Data.getProperty(String.valueOf((int)cell.getNumericCellValue()));	       		
	       		cell=row.getCell((short)5);
	       		//cell.setCellStyle(cs);
	       		cell.setCellValue(amt);
	       		
	  		}
	  		
	        FileOutputStream fout=new FileOutputStream(reportDir + System.getProperty("file.separator")+"信用部損益表.xls");
	        wb.write(fout);
	        //儲存 
	        fout.close();
		}catch(Exception e){
			System.out.println("createRpt Error:"+e+e.getMessage());
		}
		return errMsg;
	}		 
}
