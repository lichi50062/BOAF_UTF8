/* 
 *  96.11.30 create 農漁會信用部金融卡發卡及ATM裝設情形資料_明細表 by 2295
 * 100.02.22 fix 根據查詢年度.100年以後取得新縣市別.100年以前取得舊縣市別 
 	  		      使用PreparedStatement;並列印轉換後的SQL;套用QueryDB_SQLParam by 2295
 * 100.04.26 fix 修改排列順序.以直轄市先顯示.由北到南 by 2295 		
 * 102.11.19 add 原QueryDB改套用QueryDB_SQLParam by 2295      
 */ 

package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR025WB {
  public static String createRpt(String S_YEAR, String S_MONTH,String bank_type,String Unit,String showEng) {  
    System.out.println("aaaaaaaaaaaa"+showEng);
      String errMsg = "";
    List dbData = null;
    String sqlCmd = "";
    int rowNum=7;
    SimpleDateFormat logformat = new SimpleDateFormat("yyyy年MM月dd日  HH:mm:ss  ");
    String[] colname = {"bank_no", "bank_name", "atm_cnt","monthamt","yearamt","mtrancnt","ytrancnt"
    					,"push_cnt","cancel_cnt","use_cnt","debitcard_monthamt"
						,"debitcard_yearamt","debitcard_mtrancnt","debitcard_ytrancnt"};	
    String filename="農漁會信用部金融卡發卡及ATM裝設情形資料_明細表.xls";
    String cd01_table = "";
    String wlx01_m_year = "";
	List paramList = new ArrayList();
    try {
      //100.02.22 add 查詢年度100年以前.縣市別不同===============================
	  cd01_table = (Integer.parseInt(S_YEAR) < 100)?"cd01_99":""; 
	  wlx01_m_year = (Integer.parseInt(S_YEAR) < 100)?"99":"100"; 
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
    
      finput = new FileInputStream(xlsDir + System.getProperty("file.separator") + filename);

      //設定FileINputStream讀取Excel檔
      POIFSFileSystem fs = new POIFSFileSystem(finput);
      HSSFWorkbook wb = new HSSFWorkbook(fs);
      HSSFSheet sheet = wb.getSheetAt(0); //讀取第一個工作表，宣告其為sheet
      HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
      //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
      //sheet.setAutobreaks(true); //自動分頁

      //設定頁面符合列印大小
      sheet.setAutobreaks(false);
      ps.setScale( (short) 80); //列印縮放百分比
      ps.setPaperSize( (short) 9); //設定紙張大小 A4
      
      finput.close();

      HSSFRow row = null; //宣告一列
      HSSFCell cell = null; //宣告一個儲存格
      HSSFCellStyle cellStyle = wb.createCellStyle();    
      //明細內容用
  	  cellStyle = HssfStyle.setStyle( cellStyle, wb.createFont(),
              					   new String[] {
              					   "BORDER", "PHR", "PVC", "F10",
              				  	   "WRAP"} );
  	  //單位代號.機構名稱用
  	  HSSFCellStyle cellStyle_name = wb.createCellStyle();    
  	  cellStyle_name = HssfStyle.setStyle( cellStyle_name, wb.createFont(),
            					   new String[] {
            					   "BORDER", "PHL", "PVC", "F10",
            				  	   "WRAP"} );
      short i = 0;
      short y = 0;
      
      sqlCmd = "select  wlx05_m_atm.bank_no,ba01.bank_name,wlx01.english,e.fr001w_output_order,"
          	 + "	    SUM(ATM_CNT)                                 		AS  atm_cnt, "//ATM-機器台數(歷累計數)
			 + "		SUM(PUSH_DebitCard_CNT + PUSH_BinCard_CNT)   		AS  push_cnt, "//金融卡-發卡張數
			 + "		SUM(CANC_DebitCard_CNT + CANC_BinCard_CNT)   		AS  cancel_cnt, "//金融卡-停卡張數
			 + "		SUM(USE_DebitCard_CNT  + USE_BinCard_CNT)   		AS  use_cnt, "//金融卡-流通張數 
			 + "		round((SUM(DebitCard_Month_Tran_AMT)/?),0)   AS  debitcard_monthamt, "//金融卡-交易金額.本月合計
			 + "		round((SUM(DebitCard_Year_AccTran_AMT)/?),0) AS  debitcard_yearamt, "//金融卡-交易金額.本年累計
          	 + "		SUM(debitcard_month_tran_cnt)                       AS  debitcard_mtrancnt, "//金融卡-交易次數.本月合計
			 + "		SUM(debitcard_year_acctran_cnt)                     AS  debitcard_ytrancnt, "//金融卡-交易次數.本年累計
			 + "		round((SUM(Month_Tran_AMT)/?),0)      		AS  monthamt, "//ATM-交易金額.本月合計
			 + "		round((SUM(Year_AccTran_AMT)/?),0)    		AS  yearamt, "//ATM-交易金額.本年累計
			 + "		SUM(Month_Tran_CNT)                          		AS  mtrancnt, "//ATM-交易次數.本月合計
			 + "		SUM(Year_AccTran_CNT)                        		AS  ytrancnt "//ATM-交易次數.本年累計
			 + "from  WLX05_M_ATM  ,  (select * from ba01 where m_year=?)ba01 " 
			 + "LEFT JOIN (SELECT * FROM wlx01 WHERE m_year = ?) wlx01 ON wlx01.bank_no = ba01.bank_no, " 
			 + "(select * from v_bank_location where m_year=?)e  "
			 + "where (WLX05_M_ATM.m_year= ? "   
			 + " and WLX05_M_ATM.M_MONTH = ?) and "
			 + "(WLX05_M_ATM.BANK_NO=ba01.bank_no AND ba01.bank_type=?) "
			 + " and WLX05_M_ATM.bank_no = e.bank_no"
			 + " group by WLX05_M_ATM.bank_no,ba01.bank_name,wlx01.english,e.fr001w_output_order "
			 + " union "
			 + "select  '9999999','合計','','999', "
	         + "	    SUM(ATM_CNT)                                 		AS  atm_cnt, "//ATM-機器台數(歷累計數)
			 + "		SUM(PUSH_DebitCard_CNT + PUSH_BinCard_CNT)   		AS  push_cnt, "//金融卡-發卡張數
			 + "		SUM(CANC_DebitCard_CNT + CANC_BinCard_CNT)   		AS  cancel_cnt, "//金融卡-停卡張數
			 + "		SUM(USE_DebitCard_CNT  + USE_BinCard_CNT)   		AS  use_cnt, "//金融卡-流通張數 
			 + "		round((SUM(DebitCard_Month_Tran_AMT)/?),0)   AS  debitcard_monthamt, "//金融卡-交易金額.本月合計
			 + "		round((SUM(DebitCard_Year_AccTran_AMT)/?),0) AS  debitcard_yearamt, "//金融卡-交易金額.本年累計
	         + "		SUM(debitcard_month_tran_cnt)                       AS  debitcard_mtrancnt, "//金融卡-交易次數.本月合計
			 + "		SUM(debitcard_year_acctran_cnt)                     AS  debitcard_ytrancnt, "//金融卡-交易次數.本年累計
			 + "		round((SUM(Month_Tran_AMT)/?),0)      		AS  monthamt, "//ATM-交易金額.本月合計
			 + "		round((SUM(Year_AccTran_AMT)/?),0)    		AS  yearamt, "//ATM-交易金額.本年累計
			 + "		SUM(Month_Tran_CNT)                          		AS  mtrancnt, "//ATM-交易次數.本月合計
			 + "		SUM(Year_AccTran_CNT)                        		AS  ytrancnt "//ATM-交易次數.本年累計
			 + "from  WLX05_M_ATM  ,  (select * from ba01 where m_year=?)ba01 "
			 + "where (WLX05_M_ATM.m_year= ? "   
			 + " and WLX05_M_ATM.M_MONTH = ?) and "
			 + "(WLX05_M_ATM.BANK_NO=ba01.bank_no AND ba01.bank_type=?) "			 
			 + " order by fr001w_output_order,bank_no ";
      paramList.add(Unit);
      paramList.add(Unit);
      paramList.add(Unit);
      paramList.add(Unit);
      paramList.add(wlx01_m_year);
      paramList.add(wlx01_m_year);
      paramList.add(wlx01_m_year);
      paramList.add(S_YEAR);
      paramList.add(S_MONTH);
      paramList.add(bank_type);
      paramList.add(Unit);
      paramList.add(Unit);
      paramList.add(Unit);
      paramList.add(Unit);
      paramList.add(wlx01_m_year);
      paramList.add(S_YEAR);
      paramList.add(S_MONTH);
      paramList.add(bank_type);
      
      dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"atm_cnt,push_cnt,cancel_cnt,use_cnt,debitcard_monthamt,debitcard_yearamt,debitcard_mtrancnt,debitcard_ytrancnt,monthamt,yearamt,mtrancnt,ytrancnt");
      System.out.println("dbData.size=" + dbData.size());
      
      //設定報表表頭資料============================================
      row = sheet.getRow(0);
      cell = row.getCell( (short) 0);
      cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
      if(bank_type.equals("6")) {
        cell.setCellValue("農會信用部金融卡發卡及ATM裝設情形資料");
      } else {
        cell.setCellValue("漁會信用部金融卡發卡及ATM裝設情形資料");
      }
      
      if ( dbData == null || dbData.size() == 0) {
      	row = sheet.getRow(1);
        cell = row.getCell( (short) 0);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16);//設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        //設定無資料
        cell.setCellValue("民國 " + S_YEAR + " 年 " + S_MONTH + " 月 無資料存在");
      }else {
      	
      	DataObject bean = null;
      	row = sheet.getRow(1);
        cell = row.getCell( (short) 0);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16);//設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        //設定無資料
        cell.setCellValue("民國 " + S_YEAR + " 年 " + S_MONTH + " 月 ");
        //加上列印日期
        Calendar rightNow = Calendar.getInstance();
		String year = String.valueOf(rightNow.get(Calendar.YEAR)-1911);
		String month = String.valueOf(rightNow.get(Calendar.MONTH)+1);
		String day = String.valueOf(rightNow.get(Calendar.DAY_OF_MONTH));
		row=(sheet.getRow(2)==null)? sheet.createRow(2) : sheet.getRow(2);
		cell = row.getCell( (short) 0);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16);//設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue("列印日期："+year+"年"+month+"月"+day+"日"+(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa")).substring(10,19));
        //列印單位
        row=(sheet.getRow(3)==null)? sheet.createRow(3) : sheet.getRow(3);
		cell = row.getCell( (short) 0);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16);//設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue("單位：新台幣"+Utility.getUnitName(Unit)+"、％");
      	for(int rowcount=0;rowcount<dbData.size();rowcount++){
      		row = sheet.createRow(rowNum++);     
      		bean = (DataObject)dbData.get(rowcount);
    		for(int cellcount=0;cellcount<14;cellcount++){	
    			cell = row.createCell( (short) cellcount);
    			cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
    			//System.out.println("cellcount = "+cellcount);
    			if(cellcount <= 1){
    				cell.setCellStyle(cellStyle_name);
    				if(((String)bean.getValue(colname[cellcount])).equals("9999999")){
    					cell.setCellValue("");
    				}else{
    				    if(cellcount == 1){
    				        if("true".equals(showEng)){
    				            cell.setCellValue((String)bean.getValue("bank_name")+"\n"+(bean.getValue("english")==null?"":bean.getValue("english")).toString()); 
    				        }else{
    				            cell.setCellValue((String)bean.getValue(colname[cellcount])); 
    				        }
    				        
    				    }else{
    				        cell.setCellValue((String)bean.getValue(colname[cellcount])); 
    				    }
    				}
    			}else{
    			   cell.setCellStyle(cellStyle);
    			   cell.setCellValue(Utility.setCommaFormat( ((bean.getValue(colname[cellcount]) == null)?"":bean.getValue(colname[cellcount])).toString())); 
    			}
    		}//end of cellcount
      	}//end of rowcount       
      } //end of else dbData.size() != 0
      
      row = sheet.createRow(rowNum++);  
      cell = row.createCell( (short) 0);
	  cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
	  cell.setCellValue("註：");
	  row = sheet.createRow(rowNum++);  
      cell = row.createCell( (short) 0);
	  cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
	  cell.setCellValue("1.ATM機器台數包括營業廳內及營業廳外總台數。");													
	  row = sheet.createRow(rowNum++);  
      cell = row.createCell( (short) 0);
	  cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
	  cell.setCellValue("2.發卡張數指新發卡及掛失補發卡，停卡張數系指掛失卡及註銷卡，流通卡數指發卡張數扣除停卡張數。");
	  row = sheet.createRow(rowNum++);  
      cell = row.createCell( (short) 0);
	  cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
	  cell.setCellValue("3.交易次數不包括查詢次數及失敗次數。");
	  row = sheet.createRow(rowNum++);  
      cell = row.createCell( (short) 0);
	  cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示	 
	  cell.setCellValue("4.ATM 交易次數及交易金額為持有本會或其他金融機構發行之金融卡在本會自動化服務機器使用之次數及金額。");													
	  row = sheet.createRow(rowNum++);  
      cell = row.createCell( (short) 0);
	  cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
	  cell.setCellValue("5.金融卡交易次數及交易金額為本會客戶持有本會金融卡在本會或其他金融機構之自動化服務機器使用之次數及金額。");
      
      FileOutputStream fout = null;      
      fout = new FileOutputStream(reportDir +System.getProperty("file.separator") + (bank_type.equals("6")?"農":"漁")+"會信用部金融卡發卡及ATM裝設情形資料_明細表.xls");      
      HSSFFooter footer = sheet.getFooter();
      footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
      footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));	

      wb.write(fout);
      //儲存
      fout.close();

    }catch (Exception e) {
      System.out.println("createRpt Error:" + e + e.getMessage());
    }

    return errMsg;
  }
}
