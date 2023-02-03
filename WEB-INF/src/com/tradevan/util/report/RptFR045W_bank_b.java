/*
 *  Created on 2007/07/12-13 存款帳戶分級差異化管理統計表_縣市政府 by 2295 
 *  96.12.14 fix 修改備註文字 2295
 *  99.03.26 fix 根據查詢年度.100年以後取得新縣市別.100年以前取得舊縣市別 
 * 				 使用PreparedStatement;並列印轉換後的SQL;套用QueryDB_SQLParam by 2295
 * 100.02.16 fix 縣市政府無法下載報表 by 2295
 * 102.11.19 add 原QueryDB改套用QueryDB_SQLParam by 2295    
 */

package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*;
import java.io.*;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;


public class RptFR045W_bank_b {
  public static String createRpt(String S_YEAR, String S_MONTH,String m2_name ) {    

    String errMsg = "";
    List dbData = null;    
    int rowNum=0;
    DataObject bean = null;
    reportUtil reportUtil = new reportUtil();
	HSSFCellStyle cs_right = null; 
	HSSFCellStyle cs_center = null;
  
    String bank_code="";
    String bank_name="";
    String warnaccount_cnt="";
	String limitaccount_cnt="";
	String erroraccount_cnt="";
	String otheraccount_cnt="";
	String depositaccount_tcnt="";
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
         finput = new FileInputStream(xlsDir +System.getProperty("file.separator") +"存款帳戶分級差異化管理統計表_縣市政府.xls");
         
         //設定FileINputStream讀取Excel檔
         POIFSFileSystem fs = new POIFSFileSystem(finput);
         HSSFWorkbook wb = new HSSFWorkbook(fs);
         HSSFSheet sheet = wb.getSheetAt(0); //讀取第一個工作表，宣告其為sheet
         HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
         //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
         //sheet.setAutobreaks(true); //自動分頁
         
         //設定頁面符合列印大小
         sheet.setAutobreaks(false);
         ps.setScale( (short) 100); //列印縮放百分比
         
         ps.setPaperSize( (short) 9); //設定紙張大小 A4
         //wb.setSheetName(0,"test");
         finput.close();
         
         HSSFRow row = null; //宣告一列
         HSSFCell cell = null; //宣告一個儲存格
         
         short i = 0;
         short y = 0;
         cs_right = reportUtil.getRightStyle(wb);
         cs_center = reportUtil.getDefaultStyle(wb);
         
         int m_year = Integer.parseInt(S_YEAR);
         int m_month = Integer.parseInt(S_MONTH);      
         int quarter_year = m_year;//上季底
         int quarter_month = m_month;//上季底     
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
         /*縣市政府機關.其所屬農漁會.合計.與上季底之比較
          select nowA08.bank_code,bn01.bank_name,--各農漁會明細
          		  nowA08.warnaccount_cnt,
          		  nowA08.limitaccount_cnt,
          		  nowA08.erroraccount_cnt,
          		  nowA08.otheraccount_cnt,
          		  nowA08.depositaccount_tcnt
          from  (select * from a08 
 	      		  LEFT JOIN WLX01 on bank_code = WLX01.bank_no and wlx01.m_year=100
 	      		  where m_year=96 and m_month=6 and WLX01.m2_name = '9510010'	 
 	      		  )nowA08
 	      		  left join bn01 on nowA08.bank_code = bn01.bank_no
          union	   
          select '9999998' as bank_code,'合計' as bank_name, --合計
          		  sum(warnaccount_cnt) as warnaccount_cnt,
          		  sum(limitaccount_cnt) as limitaccount_cnt,
          		  sum(erroraccount_cnt) as  erroraccount_cnt,
          		  sum(otheraccount_cnt) as otheraccount_cnt,
          		  sum(depositaccount_tcnt) as depositaccount_tcnt
          from  (select * from a08 
 	      		  LEFT JOIN WLX01 on bank_code = WLX01.bank_no and wlx01.m_year=100
 	      		  where m_year=96 and m_month=6 and WLX01.m2_name = '9510010'	 
 	      		  )nowA08
 	      		  left join bn01 on nowA08.bank_code = bn01.bank_no
          union	
          select '9999999' as bank_code,'合計與上季底之比較' as bank_name,--合計與上季底之比較
           	  sum(quarter_warnaccount_cnt) as quarter_warnaccount_cnt,
           	  sum(quarter_limitaccount_cnt) as quarter_limitaccount_cnt,
           	  sum(quarter_erroraccount_cnt) as quarter_erroraccount_cnt,
           	  sum(quarter_otheraccount_cnt) as quarter_otheraccount_cnt,
           	  sum(quarter_depositaccount_tcnt) as quarter_depositaccount_tcnt
          from ( select nowA08.bank_code,bn01.bank_name,
             	  	     nowA08.warnaccount_cnt - quarterA08.warnaccount_cnt as quarter_warnaccount_cnt,
             	  	     nowA08.limitaccount_cnt - quarterA08.limitaccount_cnt as quarter_limitaccount_cnt,
             	  	     nowA08.erroraccount_cnt - quarterA08.erroraccount_cnt as quarter_erroraccount_cnt,
             	  	     nowA08.otheraccount_cnt - quarterA08.otheraccount_cnt as quarter_otheraccount_cnt,
             	  	     nowA08.depositaccount_tcnt - quarterA08.depositaccount_tcnt as quarter_depositaccount_tcnt
             	  from (select * from a08 
 	      		 	    LEFT JOIN WLX01 on bank_code = WLX01.bank_no and wlx01.m_year=100
 	      		 	    where m_year=96 and m_month=6 and WLX01.m2_name = '9510010'	 
 	      		 	    )nowA08              	   
                 left join (select * from a08 where m_year=96 and m_month=3)quarterA08 on nowA08.bank_code = quarterA08.bank_code
                 left join bn01 on nowA08.bank_code = bn01.bank_no	 and bn01.m_year=100   
	            )   
          order by bank_code    
         */
     
      
         StringBuffer sqlCmd = new StringBuffer();
      	 sqlCmd = sqlCmd.append(" select nowA08.bank_code,bn01.bank_name,");//--各農漁會明細
      	 sqlCmd = sqlCmd.append(" 		nowA08.warnaccount_cnt,");
		 sqlCmd = sqlCmd.append(" 		nowA08.limitaccount_cnt,");
		 sqlCmd = sqlCmd.append("	    nowA08.erroraccount_cnt,");
		 sqlCmd = sqlCmd.append("		nowA08.otheraccount_cnt,");
		 sqlCmd = sqlCmd.append("		nowA08.depositaccount_tcnt");
		 sqlCmd = sqlCmd.append(" from  (select * from a08 "); 
		 sqlCmd = sqlCmd.append("	    left join (select * from wlx01 where m_year=?)wlx01 on bank_code = wlx01.bank_no"); 
		 sqlCmd = sqlCmd.append(" 		where a08.m_year=? and m_month=?");
		 sqlCmd = sqlCmd.append(" 		and wlx01.m2_name = ?");
		 sqlCmd = sqlCmd.append(" 		)nowA08 left join (select * from bn01 where m_year=?)bn01 on nowA08.bank_code = bn01.bank_no");
		 sqlCmd = sqlCmd.append(" union ");	   
		 sqlCmd = sqlCmd.append(" select '9999998' as bank_code,'合計' as bank_name,");// --合計
		 sqlCmd = sqlCmd.append("		sum(warnaccount_cnt) as warnaccount_cnt,");
		 sqlCmd = sqlCmd.append("		sum(limitaccount_cnt) as limitaccount_cnt,");
		 sqlCmd = sqlCmd.append("		sum(erroraccount_cnt) as  erroraccount_cnt,");
		 sqlCmd = sqlCmd.append("		sum(otheraccount_cnt) as otheraccount_cnt,");
		 sqlCmd = sqlCmd.append("		sum(depositaccount_tcnt) as depositaccount_tcnt");
		 sqlCmd = sqlCmd.append(" from  (select * from a08 "); 
		 sqlCmd = sqlCmd.append("	    left join (select * from wlx01 where m_year=?)wlx01 on bank_code = WLX01.bank_no "); 
		 sqlCmd = sqlCmd.append(" 		where a08.m_year=? and m_month=?");
		 sqlCmd = sqlCmd.append(" 		and wlx01.m2_name = ?");
		 sqlCmd = sqlCmd.append(" 	    )nowA08 left join (select * from bn01 where m_year=?)bn01 on nowA08.bank_code = bn01.bank_no");
		 sqlCmd = sqlCmd.append(" union ");	
		 sqlCmd = sqlCmd.append(" select '9999999' as bank_code,'與上季底比較增減戶數(請以+,- 表示)' as bank_name,");//--合計與上季底之比較
		 sqlCmd = sqlCmd.append("		sum(quarter_warnaccount_cnt) as quarter_warnaccount_cnt,");
		 sqlCmd = sqlCmd.append("		sum(quarter_limitaccount_cnt) as quarter_limitaccount_cnt,");
		 sqlCmd = sqlCmd.append("		sum(quarter_erroraccount_cnt) as quarter_erroraccount_cnt,");
		 sqlCmd = sqlCmd.append("	    sum(quarter_otheraccount_cnt) as quarter_otheraccount_cnt,");
		 sqlCmd = sqlCmd.append("		sum(quarter_depositaccount_tcnt) as quarter_depositaccount_tcnt");
		 sqlCmd = sqlCmd.append(" from ( select nowA08.bank_code,bn01.bank_name,");
         sqlCmd = sqlCmd.append("		 	    nowA08.warnaccount_cnt - quarterA08.warnaccount_cnt as quarter_warnaccount_cnt,");
         sqlCmd = sqlCmd.append("	 		    nowA08.limitaccount_cnt - quarterA08.limitaccount_cnt as quarter_limitaccount_cnt,");
         sqlCmd = sqlCmd.append("			    nowA08.erroraccount_cnt - quarterA08.erroraccount_cnt as quarter_erroraccount_cnt,");
         sqlCmd = sqlCmd.append("			    nowA08.otheraccount_cnt - quarterA08.otheraccount_cnt as quarter_otheraccount_cnt,");
         sqlCmd = sqlCmd.append(" 		   		nowA08.depositaccount_tcnt - quarterA08.depositaccount_tcnt as quarter_depositaccount_tcnt");
		 sqlCmd = sqlCmd.append("    	 from (select * from a08 "); 
		 sqlCmd = sqlCmd.append("    	 	   left join (select * from wlx01 where m_year=?)wlx01 on bank_code = wlx01.bank_no"); 
		 sqlCmd = sqlCmd.append("		       where a08.m_year=? and m_month=?");
		 sqlCmd = sqlCmd.append("	           and wlx01.m2_name = ?");
		 sqlCmd = sqlCmd.append(" 	           )nowA08 ");             	   
		 sqlCmd = sqlCmd.append(" 		 left join (select * from a08 where m_year=? and m_month=?");
		 sqlCmd = sqlCmd.append("			        )quarterA08 on nowA08.bank_code = quarterA08.bank_code ");
		 sqlCmd = sqlCmd.append("		 left join (select * from bn01 where m_year=?)bn01 on nowA08.bank_code = bn01.bank_no ");   
		 sqlCmd = sqlCmd.append(" 	  ) ");   
		 sqlCmd = sqlCmd.append(" order by bank_code "); 
		 paramList.add(bn01_m_year);		 
		 paramList.add(String.valueOf(m_year));
		 paramList.add(String.valueOf(m_month));
		 paramList.add(String.valueOf(m2_name));
		 paramList.add(bn01_m_year);
		 
		 paramList.add(bn01_m_year);
		 paramList.add(String.valueOf(m_year));
		 paramList.add(String.valueOf(m_month));
		 paramList.add(String.valueOf(m2_name));
		 paramList.add(bn01_m_year);
		 
		 paramList.add(bn01_m_year);
		 paramList.add(String.valueOf(m_year));
		 paramList.add(String.valueOf(m_month));
		 paramList.add(String.valueOf(m2_name));
		 
		 paramList.add(String.valueOf(quarter_year));
		 paramList.add(String.valueOf(quarter_month));
		 paramList.add(bn01_m_year);
		 
         dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"warnaccount_cnt,limitaccount_cnt,erroraccount_cnt,otheraccount_cnt,depositaccount_tcnt");
         
         
         System.out.println("dbData.size=" + dbData.size());
         
         //設定報表表頭資料============================================
         row = sheet.getRow(1);
         cell = row.getCell( (short) 1);
         cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示      
         String m2_bank_name = "";
         paramList = new ArrayList();
         paramList.add(m2_name);
         List dbData_bank_name = DBManager.QueryDB_SQLParam("select bank_name from bn01 where bank_no=?",paramList,""); 
         if(dbData_bank_name != null && dbData_bank_name.size() != 0){
         	 bean = (DataObject)dbData_bank_name.get(0);
         	 m2_bank_name = (String)bean.getValue("bank_name");
         }
         cell.setCellValue(m2_bank_name);
         
         row=sheet.getRow(2);
         cell=row.getCell((short)1);	       	
         cell.setEncoding(HSSFCell.ENCODING_UTF_16);
         cell.setCellValue(S_YEAR+"年"+S_MONTH+"月底");  	
         
         if (dbData == null || dbData.size() ==0) {
         	 row=sheet.getRow(2);
             cell=row.getCell((short)2);	       	
             cell.setEncoding(HSSFCell.ENCODING_UTF_16);
             cell.setCellValue("無資料存在");         
         }else {
         	rowNum = 4;
            for(int idx=0;idx<dbData.size();idx++){
           	    //System.out.println("idx="+idx);
           	    bean = (DataObject)dbData.get(idx);
           	    bank_code = (bean.getValue("bank_code") == null)?"":(String)bean.getValue("bank_code");
           	    bank_name= (bean.getValue("bank_name") == null)?"":(String)bean.getValue("bank_name");
	 	           //System.out.println(bank_code);
           	    //System.out.println(bank_name);
           	    warnaccount_cnt=(bean.getValue("warnaccount_cnt") == null)?"":(bean.getValue("warnaccount_cnt")).toString();
           	    limitaccount_cnt=(bean.getValue("limitaccount_cnt") == null)?"":(bean.getValue("limitaccount_cnt")).toString();
           	    erroraccount_cnt=(bean.getValue("erroraccount_cnt") == null)?"":(bean.getValue("erroraccount_cnt")).toString();
           	    otheraccount_cnt=(bean.getValue("otheraccount_cnt") == null)?"":(bean.getValue("otheraccount_cnt")).toString();
           	    depositaccount_tcnt=(bean.getValue("depositaccount_tcnt") == null)?"":(bean.getValue("depositaccount_tcnt")).toString();
           	    
           	    row = sheet.createRow(rowNum);
		           for(int cellcount=0;cellcount<6;cellcount++){
	 	   	        cell = row.createCell( (short)cellcount);
	            		cell.setEncoding(HSSFCell.ENCODING_UTF_16);	 	
	            		if(cellcount == 0 ){
	            		   cell.setCellStyle(cs_center);
	            		}else{
	            		   cell.setCellStyle(cs_right);
	            		}
	            		if(bank_code.equals("9999999")){
	            			if(cellcount == 0) cell.setCellValue(bank_name);		        
	            			if(cellcount == 1 && !warnaccount_cnt.equals("")) cell.setCellValue(warnaccount_cnt.startsWith("-")?warnaccount_cnt:"+"+Utility.setCommaFormat(warnaccount_cnt));	 
	            			if(cellcount == 2 && !limitaccount_cnt.equals("")) cell.setCellValue(limitaccount_cnt.startsWith("-")?limitaccount_cnt:"+"+Utility.setCommaFormat(limitaccount_cnt));	 
	            			if(cellcount == 3 && !erroraccount_cnt.equals("")) cell.setCellValue(erroraccount_cnt.startsWith("-")?erroraccount_cnt:"+"+Utility.setCommaFormat(erroraccount_cnt));
	            			if(cellcount == 4 && !otheraccount_cnt.equals("")) cell.setCellValue(otheraccount_cnt.startsWith("-")?otheraccount_cnt:"+"+Utility.setCommaFormat(otheraccount_cnt));
	            			if(cellcount == 5 && !depositaccount_tcnt.equals("")) cell.setCellValue(depositaccount_tcnt.startsWith("-")?depositaccount_tcnt:"+"+Utility.setCommaFormat(depositaccount_tcnt));
	            		}else{
	            			if(cellcount == 0) cell.setCellValue(bank_name);		        
	            			if(cellcount == 1) cell.setCellValue(Utility.setCommaFormat(warnaccount_cnt));	 
	            			if(cellcount == 2) cell.setCellValue(Utility.setCommaFormat(limitaccount_cnt));	 
	            			if(cellcount == 3) cell.setCellValue(Utility.setCommaFormat(erroraccount_cnt));
	            			if(cellcount == 4) cell.setCellValue(Utility.setCommaFormat(otheraccount_cnt));
	            			if(cellcount == 5) cell.setCellValue(Utility.setCommaFormat(depositaccount_tcnt));
	            		}
		   	    }//end of cellcount	 		  	  	 	
		        rowNum++;
           }//end of dbData        
         } //end of else dbData.size() != 0
         
         row=sheet.createRow(rowNum);
         cell=row.createCell((short)0);	       	
         cell.setEncoding(HSSFCell.ENCODING_UTF_16);
         cell.setCellValue("備註：");
         rowNum++;
         row=sheet.createRow(rowNum);
         cell=row.createCell((short)0);	       	
         cell.setEncoding(HSSFCell.ENCODING_UTF_16);
         cell.setCellValue("1.存款帳戶指支票存款、活期存款、活期儲蓄存款及定期存款帳戶。");
         rowNum++;
         row=sheet.createRow(rowNum);
         cell=row.createCell((short)0);	       	
         cell.setEncoding(HSSFCell.ENCODING_UTF_16);
         cell.setCellValue("2.本表由各地縣市政府下載列印後，於每季終了次月20日前以書面報送");      					 
         
         rowNum++;
         row=sheet.createRow(rowNum);
         cell=row.createCell((short)0);	       	
         cell.setEncoding(HSSFCell.ENCODING_UTF_16);
         cell.setCellValue("   農業金融局(原第一次書面填報為95年9月底資料)。");     
         
         FileOutputStream fout = null;     
         fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + "存款帳戶分級差異化管理統計表_縣市政府.xls");
         
         HSSFFooter footer = sheet.getFooter();
         footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
         footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
         wb.write(fout);
         //儲存
         fout.close();
    }
    catch (Exception e) {
      System.out.println("RptFR045W_bank_b.createRpt Error:" + e + e.getMessage());
    }
    
    return errMsg;
  }
}
