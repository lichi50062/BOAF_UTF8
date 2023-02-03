/*
 * Created on 2007/07/12-13 存款帳戶分級差異化管理統計表_農金局身份 by 2295
 * 99.03.26 fix 根據查詢年度.100年以後取得新縣市別.100年以前取得舊縣市別 
 * 				使用PreparedStatement;並列印轉換後的SQL;套用QueryDB_SQLParam by 2295
 */

package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*;
import java.io.*;
import java.text.*;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR045W {
  public static String createRpt(String S_YEAR, String S_MONTH,
                                 String bank_type,String lguser_name,List bank_list) {
    

    String errMsg = "";
    List dbData = null;
    
    int rowNum=0;
    DataObject bean = null;
    reportUtil reportUtil = new reportUtil();
	HSSFCellStyle cs_right = null; 
	HSSFCellStyle cs_center = null;
    SimpleDateFormat logformat = new SimpleDateFormat("yyyy年MM月dd日  HH:mm:ss  ");
    Date nowlog = new Date();
    String bank_code="";
    String bank_name="";
    String warnaccount_cnt="";
	String limitaccount_cnt="";
	String erroraccount_cnt="";
	String otheraccount_cnt="";
	String depositaccount_tcnt="";
	String last_warnaccount_cnt="";
	String last_limitaccount_cnt="";
	String last_erroraccount_cnt="";
	String last_otheraccount_cnt="";
	String last_depositaccount_tcnt="";
	String quarter_warnaccount_cnt="";
	String quarter_limitaccount_cnt="";
	String quarter_erroraccount_cnt="";
	String quarter_otheraccount_cnt="";
	String quarter_depositaccount_tcnt="";
	String bank_no="";
	String bn01_m_year = "";
	List paramList = new ArrayList();
    try {
       	
         //99.03.26 add 查詢年度100年以前.縣市別不同===============================  	     
     	 bn01_m_year = (Integer.parseInt(S_YEAR) < 100)?"99":"100"; 
     	 //=====================================================================
         File xlsDir = new File(Utility.getProperties("xlsDir"));
         File reportDir = new File(Utility.getProperties("reportDir"));
    
         if (!xlsDir.exists()) {
           if (!Utility.mkdirs(Utility.getProperties("xlsDir"))) {
             errMsg += Utility.getProperties("xlsDir") + "目錄新增失敗";
           }
         }
         if (!reportDir.exists()) {
           if (!Utility.mkdirs(Utility.getProperties("reportDir"))) {
             errMsg += Utility.getProperties("reportDir") + "目錄新增失敗";
           }
         }
         FileInputStream finput = null;
    
         //input the standard report form      
         finput = new FileInputStream(xlsDir +System.getProperty("file.separator") +"存款帳戶分級差異化管理統計表.xls");
    
         //設定FileINputStream讀取Excel檔
         POIFSFileSystem fs = new POIFSFileSystem(finput);
         HSSFWorkbook wb = new HSSFWorkbook(fs);
         HSSFSheet sheet = wb.getSheetAt(0); //讀取第一個工作表，宣告其為sheet
         HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
         //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
         //sheet.setAutobreaks(true); //自動分頁
    
         //設定頁面符合列印大小
         sheet.setAutobreaks(false);
         ps.setScale( (short) 50); //列印縮放百分比
    
         ps.setPaperSize( (short) 9); //設定紙張大小 A4
         //wb.setSheetName(0,"test");
         finput.close();
    
         HSSFRow row = null; //宣告一列
         HSSFCell cell = null; //宣告一個儲存格
    
         short i = 0;
         short y = 0;
         cs_right = reportUtil.getRightStyleF12(wb);
         cs_center = reportUtil.getCenterStyleF12(wb);
         //把月份跟年作轉換 去取得上月的月份跟年份
         int m_year = Integer.parseInt(S_YEAR);
         int m_month = Integer.parseInt(S_MONTH);
         int last_year = m_year;
         int last_month = m_month - 1;
         int quarter_year = m_year;
         int quarter_month = m_month;
         if (last_month == 0) { //本月如果是1月的話 上月為12月
           last_month = 12;
           last_year--;
         }
         if(m_month >= 1 && m_month <= 3){      	
         	quarter_month = 12;
         	quarter_year--;
         }
         if(m_month >= 4 && m_month <= 6){      	
         	quarter_month = 3;
         }
         if(m_month >= 7 && m_month <= 9){      	
         	quarter_month = 6;
         }
         if(m_month >= 10 && m_month <= 12){      	
         	quarter_month = 9;
         }
         /*
          * select nowA08.bank_code,
          		    nowA08.warnaccount_cnt,
          		    nowA08.warnaccount_cnt - lastA08.warnaccount_cnt as last_warnaccount_cnt,
          		    nowA08.warnaccount_cnt - quarterA08.warnaccount_cnt as quarter_warnaccount_cnt,
          		    nowA08.limitaccount_cnt,
          		    nowA08.limitaccount_cnt - lastA08.limitaccount_cnt as last_limitaccount_cnt,
          		    nowA08.limitaccount_cnt - quarterA08.limitaccount_cnt as quarter_limitaccount_cnt,
          		    nowA08.erroraccount_cnt,
          		    nowA08.erroraccount_cnt - lastA08.erroraccount_cnt as last_erroraccount_cnt,
          		    nowA08.erroraccount_cnt - quarterA08.erroraccount_cnt as quarter_erroraccount_cnt,
          		    nowA08.otheraccount_cnt,
          		    nowA08.otheraccount_cnt - lastA08.otheraccount_cnt as last_otheraccount_cnt,
          		    nowA08.otheraccount_cnt - quarterA08.otheraccount_cnt as quarter_otheraccount_cnt,
          		    nowA08.depositaccount_tcnt,
          		    nowA08.depositaccount_tcnt - lastA08.depositaccount_tcnt as last_depositaccount_tcnt,
          		    nowA08.depositaccount_tcnt - quarterA08.depositaccount_tcnt as quarter_depositaccount_tcnt
             from 
             	    (select * from a08 where m_year=96 and m_month=6)nowA08 
             	    left join (select * from a08 where m_year=96 and m_month=5)lastA08 on nowA08.bank_code = lastA08.bank_code 
             	    left join (select * from a08 where m_year=96 and m_month=3)quarterA08 on nowA08.bank_code = quarterA08.bank_code
             union
             select '9999999','合計',sum(warnaccount_cnt) as warnaccount_cnt,
             	      sum(last_warnaccount_cnt) as last_warnaccount_cnt,
             	      sum(quarter_warnaccount_cnt) as quarter_warnaccount_cnt,
             	      sum(limitaccount_cnt) as limitaccount_cnt,
             	      sum(last_limitaccount_cnt) as last_limitaccount_cnt,
             	      sum(quarter_limitaccount_cnt) as quarter_limitaccount_cnt,
             	      sum(erroraccount_cnt) as  erroraccount_cnt,
             	      sum(last_erroraccount_cnt) as last_erroraccount_cnt,
             	      sum(quarter_erroraccount_cnt) as quarter_erroraccount_cnt,
             	      sum(otheraccount_cnt) as otheraccount_cnt,
             	      sum(last_otheraccount_cnt) as last_otheraccount_cnt,
             	      sum(quarter_otheraccount_cnt) as quarter_otheraccount_cnt,
             	      sum(depositaccount_tcnt) as depositaccount_tcnt,
             	      sum(last_depositaccount_tcnt) as last_depositaccount_tcnt,
             	      sum(quarter_depositaccount_tcnt) as quarter_depositaccount_tcnt
             from
             	     (
             	     select nowA08.bank_code,bn01.bank_name,
             	     	    nowA08.warnaccount_cnt,
             	     	    nowA08.warnaccount_cnt - lastA08.warnaccount_cnt as last_warnaccount_cnt,
             	     	    nowA08.warnaccount_cnt - quarterA08.warnaccount_cnt as quarter_warnaccount_cnt,
             	     	    nowA08.limitaccount_cnt,
             	     	    nowA08.limitaccount_cnt - lastA08.limitaccount_cnt as last_limitaccount_cnt,
             	     	    nowA08.limitaccount_cnt - quarterA08.limitaccount_cnt as quarter_limitaccount_cnt,
             	     	    nowA08.erroraccount_cnt,
             	     	    nowA08.erroraccount_cnt - lastA08.erroraccount_cnt as last_erroraccount_cnt,
             	     	    nowA08.erroraccount_cnt - quarterA08.erroraccount_cnt as quarter_erroraccount_cnt,
             	     	    nowA08.otheraccount_cnt,
             	     	    nowA08.otheraccount_cnt - lastA08.otheraccount_cnt as last_otheraccount_cnt,
             	     	    nowA08.otheraccount_cnt - quarterA08.otheraccount_cnt as quarter_otheraccount_cnt,
             	     	    nowA08.depositaccount_tcnt,
             	     	    nowA08.depositaccount_tcnt - lastA08.depositaccount_tcnt as last_depositaccount_tcnt,
             	     	    nowA08.depositaccount_tcnt - quarterA08.depositaccount_tcnt as quarter_depositaccount_tcnt
                    from 
                    	   (select * from a08 where m_year=96 and m_month=6)nowA08 
                    	   left join (select * from a08 where m_year=96 and m_month=5)lastA08 on nowA08.bank_code = lastA08.bank_code 
                    	   left join (select * from a08 where m_year=96 and m_month=3)quarterA08 on nowA08.bank_code = quarterA08.bank_code
                    	   left join bn01 on nowA08.bank_code = bn01.bank_no and bn01.m_year=100)
             order by bank_code
          */
         for(int idx_bank_no=0;idx_bank_no<bank_list.size();idx_bank_no++){
         	  bank_no += "'"+(String)((List)bank_list.get(idx_bank_no)).get(0)+"'";
         	  if(idx_bank_no != bank_list.size()-1){
        	  	 bank_no += ",";
        	  }
         }
         System.out.println("bank_no="+bank_no);
         
         
         StringBuffer table_a08 = new StringBuffer();
         
         table_a08 = table_a08.append(" select nowA08.bank_code,bn01.bank_name, ");
         table_a08 = table_a08.append(" 	   nowA08.warnaccount_cnt,");
         table_a08 = table_a08.append("		   nowA08.warnaccount_cnt - lastA08.warnaccount_cnt as last_warnaccount_cnt,");
	     table_a08 = table_a08.append("		   nowA08.warnaccount_cnt - quarterA08.warnaccount_cnt as quarter_warnaccount_cnt,");
	     table_a08 = table_a08.append("		   nowA08.limitaccount_cnt,");
	     table_a08 = table_a08.append("		   nowA08.limitaccount_cnt - lastA08.limitaccount_cnt as last_limitaccount_cnt,");
	     table_a08 = table_a08.append("		   nowA08.limitaccount_cnt - quarterA08.limitaccount_cnt as quarter_limitaccount_cnt,");
	     table_a08 = table_a08.append("    	   nowA08.erroraccount_cnt,");
	     table_a08 = table_a08.append("		   nowA08.erroraccount_cnt - lastA08.erroraccount_cnt as last_erroraccount_cnt,");
	     table_a08 = table_a08.append("		   nowA08.erroraccount_cnt - quarterA08.erroraccount_cnt as quarter_erroraccount_cnt,");
	     table_a08 = table_a08.append("		   nowA08.otheraccount_cnt,");
	     table_a08 = table_a08.append("		   nowA08.otheraccount_cnt - lastA08.otheraccount_cnt as last_otheraccount_cnt,");
	     table_a08 = table_a08.append("		   nowA08.otheraccount_cnt - quarterA08.otheraccount_cnt as quarter_otheraccount_cnt,");
	     table_a08 = table_a08.append("		   nowA08.depositaccount_tcnt,");
	     table_a08 = table_a08.append("		   nowA08.depositaccount_tcnt - lastA08.depositaccount_tcnt as last_depositaccount_tcnt,");
	     table_a08 = table_a08.append("		   nowA08.depositaccount_tcnt - quarterA08.depositaccount_tcnt as quarter_depositaccount_tcnt");
	     table_a08 = table_a08.append(" from "); 
	     table_a08 = table_a08.append("	  	(select * from a08 where m_year=? and m_month=? and bank_code in ("+bank_no+"))nowA08 "); 
	     table_a08 = table_a08.append("	  	left join (select * from a08 where m_year=? and m_month=?)lastA08 on nowA08.bank_code = lastA08.bank_code"); 
	     table_a08 = table_a08.append("	  	left join (select * from a08 where m_year=? and m_month=?)quarterA08 on nowA08.bank_code = quarterA08.bank_code");
	     table_a08 = table_a08.append(" 	left join bn01 on nowA08.bank_code = bn01.bank_no and bn01.m_year = "+bn01_m_year);
	    
	     paramList.add(String.valueOf(m_year));
		 paramList.add(String.valueOf(m_month));
		 paramList.add(String.valueOf(last_year));
		 paramList.add(String.valueOf(last_month));
		 paramList.add(String.valueOf(quarter_year));
		 paramList.add(String.valueOf(quarter_month));
		 
         StringBuffer sqlCmd = new StringBuffer();   
         sqlCmd = sqlCmd.append(table_a08);
         sqlCmd = sqlCmd.append(" union");
	   	 sqlCmd = sqlCmd.append(" select '9999999' as bank_code,'合計' as bank_name,");
	   	 sqlCmd = sqlCmd.append(" 	    sum(warnaccount_cnt) as warnaccount_cnt,");
	   	 sqlCmd = sqlCmd.append("		sum(last_warnaccount_cnt) as last_warnaccount_cnt,");
	   	 sqlCmd = sqlCmd.append("		sum(quarter_warnaccount_cnt) as quarter_warnaccount_cnt,");
	   	 sqlCmd = sqlCmd.append("		sum(limitaccount_cnt) as limitaccount_cnt,");
	   	 sqlCmd = sqlCmd.append("		sum(last_limitaccount_cnt) as last_limitaccount_cnt,");
	   	 sqlCmd = sqlCmd.append("		sum(quarter_limitaccount_cnt) as quarter_limitaccount_cnt,");
	   	 sqlCmd = sqlCmd.append("	    sum(erroraccount_cnt) as  erroraccount_cnt,");
	   	 sqlCmd = sqlCmd.append("		sum(last_erroraccount_cnt) as last_erroraccount_cnt,");
	   	 sqlCmd = sqlCmd.append("		sum(quarter_erroraccount_cnt) as quarter_erroraccount_cnt,");
	   	 sqlCmd = sqlCmd.append("	   	sum(otheraccount_cnt) as otheraccount_cnt,");
	   	 sqlCmd = sqlCmd.append("		sum(last_otheraccount_cnt) as last_otheraccount_cnt,");
	   	 sqlCmd = sqlCmd.append("		sum(quarter_otheraccount_cnt) as quarter_otheraccount_cnt,");
	   	 sqlCmd = sqlCmd.append("		sum(depositaccount_tcnt) as depositaccount_tcnt,");
	   	 sqlCmd = sqlCmd.append("		sum(last_depositaccount_tcnt) as last_depositaccount_tcnt,");
	   	 sqlCmd = sqlCmd.append("		sum(quarter_depositaccount_tcnt) as quarter_depositaccount_tcnt");
	   	 sqlCmd = sqlCmd.append(" from (");
	   	 sqlCmd = sqlCmd.append(table_a08);
	   	 sqlCmd = sqlCmd.append(")");
	   	 sqlCmd = sqlCmd.append(" order by bank_code ");
	   	 
	   	 paramList.add(String.valueOf(m_year));
		 paramList.add(String.valueOf(m_month));
		 paramList.add(String.valueOf(last_year));
		 paramList.add(String.valueOf(last_month));
		 paramList.add(String.valueOf(quarter_year));
		 paramList.add(String.valueOf(quarter_month));
    
         dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"warnaccount_cnt,last_warnaccount_cnt,quarter_warnaccount_cnt," +
         									"limitaccount_cnt,last_limitaccount_cnt,quarter_limitaccount_cnt,"+
											"erroraccount_cnt,last_erroraccount_cnt,quarter_erroraccount_cnt,"+										 
											"otheraccount_cnt,last_otheraccount_cnt,quarter_otheraccount_cnt,"+
	   										"depositaccount_tcnt,last_depositaccount_tcnt,quarter_depositaccount_tcnt");
    
         
         System.out.println("dbData.size=" + dbData.size());
         //取得當前日期
         Calendar rightNow = Calendar.getInstance();
         
         String year = String.valueOf(rightNow.get(Calendar.YEAR) - 1911);
        
         nowlog = rightNow.getTime();
        
         //設定報表表頭資料============================================
         row = sheet.getRow(1);
         cell = row.getCell( (short) 0);
         cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示      
         cell.setCellValue("農漁會信用部存款帳戶分級差異化管理統計表");
        
         row=sheet.getRow(2);
         cell=row.getCell((short)15);	       	
         cell.setEncoding(HSSFCell.ENCODING_UTF_16);	
         cell.setCellValue("列印日期："+year+(logformat.format(nowlog)).toString().substring(4));  	
         row=sheet.getRow(3);
         
         cell=row.getCell((short)0);	       	
         cell.setEncoding(HSSFCell.ENCODING_UTF_16);	  	 		
         cell.setCellValue("列印人員："+lguser_name);
         
         //判斷a1是不是null，是的話表示沒有資料
         if (dbData == null || dbData.size() ==0) {
         	  row=sheet.getRow(2);
              cell=row.getCell((short)0);	       	
              cell.setEncoding(HSSFCell.ENCODING_UTF_16);
              cell.setCellValue("民國 " + S_YEAR + " 年 " + S_MONTH + " 月底 無資料存在");
            
         }else {
         	//設定日期===================================================
            row = sheet.getRow(2);
            cell = row.getCell( (short) 0);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
            cell.setCellValue("民國 " + S_YEAR + " 年 " + S_MONTH + " 月底");
          
            rowNum = 6;
            for(int idx=0;idx<dbData.size();idx++){
           	    //System.out.println("idx="+idx);
           	    bean = (DataObject)dbData.get(idx);
           	    bank_code = (bean.getValue("bank_code") == null)?"":(String)bean.getValue("bank_code");
           	    if(bank_code.equals("9999999") && idx==0){
           	    	row=sheet.getRow(2);
                    cell=row.getCell((short)0);	       	
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue("民國 " + S_YEAR + " 年 " + S_MONTH + " 月底 無資料存在");                
           	    	continue; 
           	    }
	    	    bank_name= (bean.getValue("bank_name") == null)?"":(String)bean.getValue("bank_name");
	    	    //System.out.println(bank_code);
           	    //System.out.println(bank_name);
           	    warnaccount_cnt=(bean.getValue("warnaccount_cnt") == null)?"":(bean.getValue("warnaccount_cnt")).toString();
           	    limitaccount_cnt=(bean.getValue("limitaccount_cnt") == null)?"":(bean.getValue("limitaccount_cnt")).toString();
           	    erroraccount_cnt=(bean.getValue("erroraccount_cnt") == null)?"":(bean.getValue("erroraccount_cnt")).toString();
           	    otheraccount_cnt=(bean.getValue("otheraccount_cnt") == null)?"":(bean.getValue("otheraccount_cnt")).toString();
           	    depositaccount_tcnt=(bean.getValue("depositaccount_tcnt") == null)?"":(bean.getValue("depositaccount_tcnt")).toString();
           	    last_warnaccount_cnt=(bean.getValue("last_warnaccount_cnt") == null)?"":(bean.getValue("last_warnaccount_cnt")).toString();
           	    last_limitaccount_cnt=(bean.getValue("last_limitaccount_cnt") == null)?"":(bean.getValue("last_limitaccount_cnt")).toString();
           	    last_erroraccount_cnt=(bean.getValue("last_erroraccount_cnt") == null)?"":(bean.getValue("last_erroraccount_cnt")).toString();
           	    last_otheraccount_cnt=(bean.getValue("last_otheraccount_cnt") == null)?"":(bean.getValue("last_otheraccount_cnt")).toString();
           	    last_depositaccount_tcnt=(bean.getValue("last_depositaccount_tcnt") == null)?"":(bean.getValue("last_depositaccount_tcnt")).toString();
           	    quarter_warnaccount_cnt=(bean.getValue("quarter_warnaccount_cnt") == null)?"":(bean.getValue("quarter_warnaccount_cnt")).toString();
           	    quarter_limitaccount_cnt=(bean.getValue("quarter_limitaccount_cnt") == null)?"":(bean.getValue("quarter_limitaccount_cnt")).toString();
           	    quarter_erroraccount_cnt=(bean.getValue("quarter_erroraccount_cnt") == null)?"":(bean.getValue("quarter_erroraccount_cnt")).toString();
           	    quarter_otheraccount_cnt=(bean.getValue("quarter_otheraccount_cnt") == null)?"":(bean.getValue("quarter_otheraccount_cnt")).toString();
           	    quarter_depositaccount_tcnt=(bean.getValue("quarter_depositaccount_tcnt") == null)?"":(bean.getValue("quarter_depositaccount_tcnt")).toString();
           	    warnaccount_cnt = warnaccount_cnt.equals("")?"":((Integer.parseInt(warnaccount_cnt) < 0)?warnaccount_cnt:Utility.setCommaFormat(warnaccount_cnt));     		
        	    limitaccount_cnt = limitaccount_cnt.equals("")?"":((Integer.parseInt(limitaccount_cnt) < 0)?limitaccount_cnt:Utility.setCommaFormat(limitaccount_cnt));
        	    erroraccount_cnt = erroraccount_cnt.equals("")?"":((Integer.parseInt(erroraccount_cnt) < 0)?erroraccount_cnt:Utility.setCommaFormat(erroraccount_cnt));        	
           	    otheraccount_cnt = otheraccount_cnt.equals("")?"":((Integer.parseInt(otheraccount_cnt) < 0)?otheraccount_cnt:Utility.setCommaFormat(otheraccount_cnt));        	
        	    depositaccount_tcnt = depositaccount_tcnt.equals("")?"":((Integer.parseInt(depositaccount_tcnt) < 0)?depositaccount_tcnt:Utility.setCommaFormat(depositaccount_tcnt));
           	    last_warnaccount_cnt = last_warnaccount_cnt.equals("")?"":((Integer.parseInt(last_warnaccount_cnt) < 0)?last_warnaccount_cnt:Utility.setCommaFormat(last_warnaccount_cnt));
           	    //System.out.println("last_warnaccount_cnt="+last_warnaccount_cnt);
           	    last_limitaccount_cnt = last_limitaccount_cnt.equals("")?"":((Integer.parseInt(last_limitaccount_cnt) < 0)?last_limitaccount_cnt:Utility.setCommaFormat(last_limitaccount_cnt));
           	    //System.out.println("last_limitaccount_cnt="+last_limitaccount_cnt);
           	    last_erroraccount_cnt = last_erroraccount_cnt.equals("")?"":((Integer.parseInt(last_erroraccount_cnt) < 0)?last_erroraccount_cnt:Utility.setCommaFormat(last_erroraccount_cnt));
           	    //System.out.println("last_erroraccount_cnt="+last_erroraccount_cnt);
           	    last_otheraccount_cnt = last_otheraccount_cnt.equals("")?"":((Integer.parseInt(last_otheraccount_cnt) < 0)?last_otheraccount_cnt:Utility.setCommaFormat(last_otheraccount_cnt));
           	    //System.out.println("last_otheraccount_cnt="+last_otheraccount_cnt);
           	    last_depositaccount_tcnt = last_depositaccount_tcnt.equals("")?"":((Integer.parseInt(last_depositaccount_tcnt) < 0)?last_depositaccount_tcnt:Utility.setCommaFormat(last_depositaccount_tcnt));
           	    //System.out.println("last_depositaccount_tcnt="+last_depositaccount_tcnt);
           	    quarter_warnaccount_cnt = quarter_warnaccount_cnt.equals("")?"":((Integer.parseInt(quarter_warnaccount_cnt) < 0)?quarter_warnaccount_cnt:Utility.setCommaFormat(quarter_warnaccount_cnt));
           	    //System.out.println("quarter_warnaccount_cnt="+quarter_warnaccount_cnt);
           	    quarter_limitaccount_cnt = quarter_limitaccount_cnt.equals("")?"":((Integer.parseInt(quarter_limitaccount_cnt) < 0)?quarter_limitaccount_cnt:Utility.setCommaFormat(quarter_limitaccount_cnt));
           	    //System.out.println("quarter_limitaccount_cnt="+quarter_limitaccount_cnt);
           	    quarter_erroraccount_cnt = quarter_erroraccount_cnt.equals("")?"":((Integer.parseInt(quarter_erroraccount_cnt) < 0)?quarter_erroraccount_cnt:Utility.setCommaFormat(quarter_erroraccount_cnt));
           	    //System.out.println("quarter_erroraccount_cnt="+quarter_erroraccount_cnt);
           	    quarter_otheraccount_cnt = quarter_otheraccount_cnt.equals("")?"":((Integer.parseInt(quarter_otheraccount_cnt) < 0)?quarter_otheraccount_cnt:Utility.setCommaFormat(quarter_otheraccount_cnt));
           	    //System.out.println("quarter_otheraccount_cnt="+quarter_otheraccount_cnt);
           	    quarter_depositaccount_tcnt = quarter_depositaccount_tcnt.equals("")?"":((Integer.parseInt(quarter_depositaccount_tcnt) < 0)?quarter_depositaccount_tcnt:Utility.setCommaFormat(quarter_depositaccount_tcnt));
           	    //System.out.println("quarter_depositaccount_tcnt="+quarter_depositaccount_tcnt);		
        	    	
           	    
           	    row = sheet.createRow(rowNum);
	   	        for(int cellcount=0;cellcount<18;cellcount++){
	    	    	cell = row.createCell( (short)cellcount);
	            	cell.setEncoding(HSSFCell.ENCODING_UTF_16);	 	
	            	if(cellcount == 0 || cellcount == 1){
	            	   cell.setCellStyle(cs_center);
	            	}else{
	            	   cell.setCellStyle(cs_right);
	            	}
	            	if(cellcount == 0)  cell.setCellValue(bank_code.equals("9999999")?"":bank_code);		        
	   		        if(cellcount == 1)  cell.setCellValue(bank_name);	 
	   		        if(cellcount == 2)  cell.setCellValue( S_YEAR + "/" + S_MONTH );	 
	   		        if(cellcount == 3)  cell.setCellValue(warnaccount_cnt);	 
	   		        if(cellcount == 4)  cell.setCellValue((last_warnaccount_cnt.startsWith("-")==false && last_warnaccount_cnt.length() > 1?"+":"")+last_warnaccount_cnt);
	   		        if(cellcount == 5)  cell.setCellValue((quarter_warnaccount_cnt.startsWith("-")==false && quarter_warnaccount_cnt.length() > 1?"+":"")+quarter_warnaccount_cnt);
	   		        if(cellcount == 6)  cell.setCellValue(limitaccount_cnt);	 
	   		        if(cellcount == 7)  cell.setCellValue((last_limitaccount_cnt.startsWith("-")==false && last_limitaccount_cnt.length() > 1?"+":"")+last_limitaccount_cnt);
	   		        if(cellcount == 8)  cell.setCellValue((quarter_limitaccount_cnt.startsWith("-")==false && quarter_limitaccount_cnt.length() > 1?"+":"")+quarter_limitaccount_cnt);			    	 
	   		        if(cellcount == 9)  cell.setCellValue(erroraccount_cnt);
	   		        if(cellcount == 10)  cell.setCellValue((last_erroraccount_cnt.startsWith("-")==false && last_erroraccount_cnt.length() > 1?"+":"")+last_erroraccount_cnt);
	   		        if(cellcount == 11)  cell.setCellValue((quarter_erroraccount_cnt.startsWith("-")==false && quarter_erroraccount_cnt.length() > 1?"+":"")+quarter_erroraccount_cnt);	 
	   		        if(cellcount == 12)  cell.setCellValue(otheraccount_cnt);
	   		        if(cellcount == 13)  cell.setCellValue((last_otheraccount_cnt.startsWith("-")==false && last_otheraccount_cnt.length() > 1?"+":"")+last_otheraccount_cnt);
	   		        if(cellcount == 14)  cell.setCellValue((quarter_otheraccount_cnt.startsWith("-")==false && quarter_otheraccount_cnt.length() > 1?"+":"")+quarter_otheraccount_cnt);	 
	   		        if(cellcount == 15)  cell.setCellValue(depositaccount_tcnt);
	   		        if(cellcount == 16)  cell.setCellValue((last_depositaccount_tcnt.startsWith("-")==false && last_depositaccount_tcnt.length() > 1?"+":"")+last_depositaccount_tcnt);
	   		        if(cellcount == 17)  cell.setCellValue((quarter_depositaccount_tcnt.startsWith("-")==false && quarter_depositaccount_tcnt.length() > 1?"+":"")+quarter_depositaccount_tcnt);
	   		    }//end of cellcount	 		  	  	 	
	   	        rowNum++;
            }//end of dbData        
         } //end of else dbData.size() != 0
    
         FileOutputStream fout = null;     
         fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + "存款帳戶分級差異化管理統計表.xls");
        
         HSSFFooter footer = sheet.getFooter();
         footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
         footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
         wb.write(fout);
         //儲存
         fout.close();
       }catch (Exception e) {
         System.out.println("RptFR045W.createRpt Error:" + e + e.getMessage());
       }       
       return errMsg;
     }
}
