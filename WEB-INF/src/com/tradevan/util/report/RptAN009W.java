/*
 * Created on 2006/10/23 by ABYSS Brenda
 * 某年底各縣市農漁會信用部逾期放款、比率及存放款及比率彙總表
 *
 * 2006/12/18 by Abyss Brenda
 * 1.新增其他縣市別
 * 2.金額四捨五入至整數位
 *
 * 2006/12/22 by Abyss Brenda
 * 1.新增可區分農漁會
 * 2.新增選擇金額單位
 *
 * 2006/12/26 by Abyss Brenda
 * 燕貞提出存放比率計算方式修改
 * 1.(X)修正後方式總額：有邏輯判斷的部份，需先將該縣市別的所有金額合計後（含農漁會），再做判斷。
 * 2.(Y)修正後存款總額：各信用部（含農漁會）依公式計算出Y值後（SUM(a:j)-i/2)，再做合計。
 * 
 * 2009/04.14 fix 縣市合併&查詢方式 by 2808
 * fixed 102.06.03 sql by 2968
 */

package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.poifs.filesystem.*;
import java.io.*;
import java.util.*;
import com.tradevan.util.DBManager;
import com.tradevan.util.Utility;
import com.tradevan.util.Utility_report;
import com.tradevan.util.dao.DataObject;

public class RptAN009W {
  public static String createRpt(String S_YEAR, String bankType ,String tmUnit) {
    boolean debug = false;
    if(debug) System.out.println("RptAN009W Start ...");

    //金額單位
    String unitName = tmUnit.substring(0,tmUnit.indexOf(";"));
    String unit = tmUnit.substring(tmUnit.indexOf(";") + 1);

    // 取得當下年月資料轉換西元年到民國年
    Calendar c = Calendar.getInstance();
    //String nowYear = Integer.toString(c.get(Calendar.YEAR) - 1911);
    //String nowMonth = Integer.toString(c.get(Calendar.MONTH) + 1);
    //if(!S_YEAR.equals(nowYear)) nowMonth = "12";

    //2006-12-19修改只取得12月份資料
    int nowYear = c.get(Calendar.YEAR) - 1911;
    int qYear = Integer.parseInt(S_YEAR);
    String nowMonth = "12";
    
    String errMsg = "";
    String u_year = "100" ; //判斷縣市合併用
    
    if(S_YEAR==null || Integer.parseInt(S_YEAR) < 100 ) {
    	u_year  = "99" ;
    }
    try {
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
      System.out.println("某年底各縣市農漁會信用部逾期放款、比率及存放款及比率彙總表.xls");

      finput = new FileInputStream(xlsDir +
                                   System.getProperty("file.separator") +
                                   "某年底各縣市農漁會信用部逾期放款、比率及存放款及比率彙總表.xls");

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

      //設定報表表頭資料============================================
      row = sheet.getRow(0);
      cell = row.getCell( (short) 0);
      cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
      String tmBankType = "農漁";
      if(bankType.equals("6")){
        tmBankType = "農";
      }else if(bankType.equals("7")){
        tmBankType = "漁";
      }
      cell.setCellValue(S_YEAR + " 年底各縣市"+tmBankType+"會信用部逾期放款、比率及存放款及比率彙總表");

      //設定年月及單位資料============================================
      row = sheet.getRow(1);
      cell = row.getCell( (short) 0);
      cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
      cell.setCellValue("單位：新台幣 " + unitName);

      //if(qYear >= nowYear){
      if(false) {
        row = sheet.getRow(3);
        cell = row.getCell( (short) 0);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue("查無資料");
        sheet.addMergedRegion(new Region((short)3, (short)0,(short)3, (short)11));//跨行(第幾行,開始欄位數,跨幾行,結束欄位數)
      }else{
    	  StringBuffer sql = new StringBuffer () ;
          List paramList = new ArrayList() ;
          DataObject bean = null;
          sql.append(getReportSql(bankType,u_year)) ;
          paramList = getReportParameter(bankType,u_year,unit,S_YEAR);
          List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "hsien_id,hsien_name,fr001w_output_order,field_seq,field_over," +
          		"field_debit,field_credit,field_dc_rate,field_over_rate");
          
          System.out.println("layer 1 dbData.size()="+dbData.size());
          long sum1 = 0l;
          long sum2 = 0l;
          if(dbData.size() > 0 ) {
          int k = 0 ;
          for(int i=0;i<dbData.size();i++){
				bean = (DataObject) dbData.get(i);
				String hsien_id = bean.getValue("hsien_id")==null?"":bean.getValue("hsien_id").toString() ;
				String hsien_name = bean.getValue("hsien_name")==null?"":bean.getValue("hsien_name").toString() ;
				String fr001w_output_order = bean.getValue("fr001w_output_order")==null?"":bean.getValue("fr001w_output_order").toString();
				System.out.println("縣市:"+hsien_name) ;
				if("其他".equals(hsien_name)||"029".equals(fr001w_output_order)){
					continue ;  
				}
				String field_over      = bean.getValue("field_over")==null?"0":bean.getValue("field_over").toString();
				String field_debit     = bean.getValue("field_debit")==null?"0":bean.getValue("field_debit").toString();
				String field_credit    = bean.getValue("field_credit")==null?"0":bean.getValue("field_credit").toString(); 
				String field_dc_rate   = bean.getValue("field_dc_rate")==null?"0":bean.getValue("field_dc_rate").toString();
				String field_over_rate = bean.getValue("field_over_rate")==null?"0":bean.getValue("field_over_rate").toString();
				System.out.println("逾放金額:"+field_over) ;
				/*System.out.println("逾放金額"+field_over) ;
				System.out.println("逾放比率:"+field_over_rate) ;
				System.out.println("存款總額:"+field_debit) ;
				System.out.println("放款總額:"+field_credit) ;
				System.out.println("存放比率:"+field_dc_rate) ;*/
				
		        row = sheet.getRow(k + 3);
	            cell = row.getCell( (short) 0);
	            cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
	            cell.setCellValue(hsien_name);
	            
	            cell = row.getCell( (short) 2);
	            cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
	            cell.setCellValue(Utility.setCommaFormat(field_over));
	            
	            cell = row.getCell( (short) 4);
	            cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
	            if ("0".equals(field_over_rate)) field_over_rate = "0.00";
	            cell.setCellValue(field_over_rate);
	            
	            cell = row.getCell( (short) 6);
	            cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
	            cell.setCellValue(Utility.setCommaFormat(field_debit));
	           
	            cell = row.getCell( (short) 8);
	            cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
	            cell.setCellValue(Utility.setCommaFormat(field_credit));
	            
	            cell = row.getCell( (short) 10);
	            cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
	            cell.setCellValue(field_dc_rate);
	            
		        sum1 += field_over==null? 0 : Long.parseLong(field_over);
		        sum2 += field_credit==null? 0 :Long.parseLong(field_credit);
		        k++ ;
		        
            }
          	//設定年月資料及備註============================================
            row = sheet.getRow(30);
            cell = row.getCell( (short) 0);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
            cell.setCellValue("註：" + S_YEAR + " 年 " + nowMonth +
                            " 月底全體" + tmBankType + "會信用部逾期放款比率為 " +
                            Utility_report.round(Long.toString(sum1 * 100),
                                                 Long.toString(sum2), 2));
          
         }else {
	    	  row = sheet.getRow(3);
	          cell = row.getCell( (short) 0);
	          cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
	          cell.setCellValue("查無資料");
	          sheet.addMergedRegion(new Region((short)3, (short)0,(short)3, (short)11));//跨行(第幾行,開始欄位數,跨幾行,結束欄位數)	
         }
      }
      FileOutputStream fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + "某年底各縣市農漁會信用部逾期放款、比率及存放款及比率彙總表.xls");
      HSSFFooter footer = sheet.getFooter();
      footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
      footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));

      wb.write(fout);
      //儲存
      fout.close();
    }catch (Exception e) {
      System.out.println("//RptAN009W createRpt() Have Error.....");
      e.printStackTrace();
      System.out.println(e);
      System.out.println(e.getMessage());
      System.out.println("//-------------------------------------");
    }

    if(debug) System.out.println("RptAN009W End ...");
    return errMsg;
  }

  
  /***
   * 取得報表SQL.
   * 
   * @param bankType
   * @return
   */
  private static String getReportSql(String bankType,String u_year){
	  String cd01Table = "cd01" ;
	  StringBuffer sql = new StringBuffer() ;
	  System.out.println("getReportSql_uyear:"+u_year) ;
	  if(Integer.parseInt(u_year)<100) {
		  cd01Table = "cd01_99" ;
	  }
	  sql.append("select a01.hsien_id , a01.hsien_name, a01.fr001w_output_order, "); 
	  sql.append("       round(field_OVER /?,0)    as field_over, "); //--逾放金額 
	  sql.append("       round(field_DEBIT /?,0)   as field_debit, "); //--存款總額 
	  sql.append("       round(field_CREDIT /?,0)  as field_credit, "); //--放款總額
	  sql.append("       decode(a01.fieldI_Y,0,0, ");                                  
	  sql.append("              round((a01.fieldI_XA  ");                              
	  sql.append("                  + decode(sign(a01.fieldI_XB1 - a01.fieldI_XB2),-1,0,(a01.fieldI_XB1 - a01.fieldI_XB2)) ");   
	  sql.append("                  + decode(sign(a01.fieldI_XC1 - a01.fieldI_XC2),-1,0,(a01.fieldI_XC1 - a01.fieldI_XC2)) ");   
	  sql.append("                  + decode(sign(a01.fieldI_XD1 - a01.fieldI_XD2),-1,0,(a01.fieldI_XD1 - a01.fieldI_XD2)) ");   
	  sql.append("                  + decode(sign(a01.fieldI_XE1 - a01.fieldI_XE2),-1,0,(a01.fieldI_XE1 - a01.fieldI_XE2)) ");   
	  sql.append("                  - decode(sign(a01.fieldI_XF1 - a01.fieldI_XF3 -  a01.fieldI_XF2),-1,0,(a01.fieldI_XF1 - a01.fieldI_XF3 - a01.fieldI_XF2)) ");   
	  sql.append("                    )/a01.fieldI_Y * 100,2)) as field_dc_rate , "); //--存放比率   
	  sql.append("       decode(a01.field_credit,0,0,round(a01.field_over /  a01.field_credit *100 ,2))  as field_over_rate "); //--逾放比率          
	  sql.append("from ( ");
	  sql.append("      select a01.hsien_id ,  a01.hsien_name, a01.FR001W_output_order, ");
	  sql.append("             SUM(field_OVER) field_OVER, SUM(field_DEBIT) field_DEBIT,SUM(field_CREDIT) field_CREDIT,SUM(fieldI_XA) fieldI_XA, ");                
	  sql.append("             SUM(fieldI_XB1) fieldI_XB1, SUM(fieldI_XB2) fieldI_XB2, SUM(fieldI_XC1) fieldI_XC1,SUM(fieldI_XC2) fieldI_XC2, ");               
	  sql.append("             SUM(fieldI_XD1) fieldI_XD1, SUM(fieldI_XD2) fieldI_XD2, SUM(fieldI_XE1) fieldI_XE1,SUM(fieldI_XE2) fieldI_XE2, ");               
	  sql.append("             SUM(fieldI_XF1) fieldI_XF1,SUM(fieldI_XF3) fieldI_XF3,SUM(fieldI_XF2)  fieldI_XF2, SUM(fieldI_Y) fieldI_Y ");                  
	  sql.append("      from ( select nvl(cd01.hsien_id,' ') as hsien_id , ");    
	  sql.append("                    nvl(cd01.hsien_name,'OTHER') as hsien_name, ");     
	  sql.append("                    cd01.FR001W_output_order as FR001W_output_order, ");
	  sql.append("                    round(sum(decode(a01.acc_code,'990000',amt,0)) /1,0) as field_OVER, ");    
	  sql.append("                    round(sum(decode(a01.acc_code,'220000',amt,0)) /1,0) as field_DEBIT, ");            
	  sql.append("                    round(sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /1,0) as field_CREDIT, ");   
	  sql.append("                    decode(YEAR_TYPE,'102',round(sum(decode(bank_type,'6',decode(a01.acc_code,'120101',amt,'120102',amt,'120200',amt,'120301',amt,'120302',amt,'120700',amt,'150200',amt,0) ");            
	  sql.append("                                                                     ,'7',decode(a01.acc_code,'120101',amt,'120102',amt,'120300',amt,'120401',amt,'120402',amt,'120700',amt,'150200',amt,0))) /1,0), ");
	  sql.append("                                     '103',round(sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120200',amt,'120301',amt,'120302',amt,'120700',amt,'150200',amt,0)) /1,0),0) as fieldI_XA, ");   
	  sql.append("                    decode(YEAR_TYPE,'102',round(sum(decode(bank_type,'6',decode(a01.acc_code,'120401',amt,'120402',amt,0),'7',decode(a01.acc_code,'120201',amt,'120202',amt,0))) /1,0), ");
	  sql.append("                                     '103', round(sum(decode(a01.acc_code,'120401',amt,'120402',amt,0)) /1,0),0) as fieldI_XB1, ");   
	  sql.append("                    decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'240305',amt,'251300',amt,'310800',amt, 0),'7',decode(a01.acc_code,'240205',amt, '310800',amt,0))), ");
	  sql.append("                                     '103',sum(decode(bank_type,'6',decode(a01.acc_code,'240305',amt,'251300',amt,'310800',amt, 0),'7',decode(a01.acc_code,'240305',amt, '251200',amt,0))),0)  as fieldI_XB2, ");   
	  sql.append("                    round(sum(decode(a01.acc_code,'120501',amt,'120502',amt,0)) /1,0) as fieldI_XC1, ");         
	  sql.append("                    decode(YEAR_TYPE,'102', round(sum(decode(bank_type,'6',decode(a01.acc_code,'240301',amt,'240302',amt,'240303',amt,'240304',amt,0), ");                                                             
	  sql.append("                                                                       '7',decode(a01.acc_code,'240201',amt,'240202',amt,'240203',amt,'240204',amt,0)) ) /1,0) , ");
	  sql.append("                                     '103', round(sum(decode(a01.acc_code,'240301',amt,'240302',amt,'240303',amt,'240304',amt,0)) /1,0),0) as fieldI_XC2, ");                            
	  sql.append("                    round(sum(decode(a01.acc_code,'120600',amt,0)) /1,0) as fieldI_XD1, ");        
	  sql.append("                    decode(YEAR_TYPE,'102',round(sum(decode(bank_type,'6',decode(a01.acc_code,'240200',amt,0),'7',decode(a01.acc_code,'240300',amt,0)) ) /1,0) , ");
	  sql.append("                                     '103',round(sum(decode(a01.acc_code,'240200',amt,0)) /1,0)  ,0) as fieldI_XD2, ");  
	  sql.append("                    round(sum(decode(a01.acc_code,'150100',amt,0)) /1,0) as fieldI_XE1, ");                                                                               
	  sql.append("                    round(sum(decode(a01.acc_code,'250100',amt,0)) /1,0) as fieldI_XE2, ");                                                                              
	  sql.append("                    round(sum(decode(a01.acc_code,'310000',amt,'320000',amt,0)) /1,0) as fieldI_XF1, ");                                                                               
	  sql.append("                    decode(YEAR_TYPE,'102',round(sum(decode(a01.acc_code,'310800',amt,0)) /1,0), ");
	  sql.append("                                     '103',round(sum(decode(bank_type,'6',decode(a01.acc_code,'310800',amt,0), '7',0) ) /1,0),0)  as fieldI_XF3, ");                                                                                  
	  sql.append("                    round(sum(decode(a01.acc_code,'140000',amt,0)) /1,0) as fieldI_XF2, ");                                                                               
	  sql.append("                    round((sum(decode(a01.acc_code,'220100',amt,'220200',amt,'220300',amt,'220400',amt,'220500',amt,'220600',amt,'220700',amt,'220800',amt, '220900',amt,'221000',amt,0)) ");                                                                                                   
	  sql.append("                          - round(sum(decode(a01.acc_code,'220900',amt,0))/2,0)) /1,0) as fieldI_Y ");                                                                                 
	  sql.append("              from  (select * from ").append(cd01Table).append(" cd01 where cd01.hsien_id <> 'Y') cd01 "); 
	  sql.append("              left join (select cd01.hsien_id,cd01.hsien_name,cd01.fr001w_output_order,bn01.bank_type,bn01.bank_no,bn01.bank_name ");
	  sql.append("                         from ").append(cd01Table).append(" cd01, (select * from wlx01 where m_year=? )wlx01, (select * from bn01 where m_year=? ");
	  if("".equals(bankType)) {
          sql.append(" and bank_type in ('6','7') ");
      }else{
          sql.append(" and bank_type in (?) ");
      }
	  sql.append("                                 )bn01  ");
	  sql.append("                         where wlx01.hsien_id=cd01.hsien_id ");  
	  sql.append("                         and wlx01.bank_no=bn01.bank_no)wlx01 on cd01.hsien_id=wlx01.hsien_id ");     
	  sql.append("              left join (select (CASE WHEN (a01.m_year <= 102) THEN '102' ");
	  sql.append("                                      WHEN (a01.m_year > 102) THEN '103' ");
	  sql.append("                                      ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01 "); 
	  sql.append("                         where  a01.m_year  = ? and a01.m_month  = 12 ) a01 on  wlx01.bank_no = a01.bank_code ");                                                                                                                             
	  sql.append("               group by a01.YEAR_TYPE,nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order ");                              
	  sql.append("          ) a01 ");                   
	  sql.append("      where a01.hsien_id <> 'Y' ");
	  sql.append("      GROUP  BY a01.hsien_id ,a01.hsien_name, a01.FR001W_output_order ");
	  sql.append(")a01 ");               
	  sql.append("order by FR001W_output_order ");
	  /*sql.append(" select T1.hsien_id,T2.hsien_name,T2.FR001W_output_order, ");
	  sql.append(" T1.COUNT_SEQ,T1.field_SEQ,T1.field_OVER,T1.field_DEBIT,T1.field_CREDIT,");
	  sql.append(" T1.field_DC_RATE,T1.field_OVER_RATE ");
	  sql.append(" from ").append(cd01Table).append(" T2 ,");
	  sql.append(" (select a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order,");
	  sql.append("        COUNT_SEQ,  field_SEQ,                                   ");
	  sql.append("        round(field_OVER /1,0)    as field_OVER,      ");//--逾放金額 
	  sql.append("        round(field_DEBIT /1,0)   as field_DEBIT,     ");//--存款總額 
	  sql.append("        round(field_CREDIT /1,0)  as field_CREDIT,    ");//--放款總額
	  sql.append("        decode(a01.fieldI_Y,0,0,                                 ");
	  sql.append("               round((a01.fieldI_XA                              ");
	  sql.append("                   + decode(sign(a01.fieldI_XB1 - a01.fieldI_XB2),-1,0,(a01.fieldI_XB1 - a01.fieldI_XB2))  ");
	  sql.append("                   + decode(sign(a01.fieldI_XC1 - a01.fieldI_XC2),-1,0,(a01.fieldI_XC1 - a01.fieldI_XC2))  ");
	  sql.append("                   + decode(sign(a01.fieldI_XD1 - a01.fieldI_XD2),-1,0,(a01.fieldI_XD1 - a01.fieldI_XD2))  ");
	  sql.append("                   + decode(sign(a01.fieldI_XE1 - a01.fieldI_XE2),-1,0,(a01.fieldI_XE1 - a01.fieldI_XE2))  ");
	  sql.append("                   - decode(sign(a01.fieldI_XF1 - a01.fieldI_XF3 -  a01.fieldI_XF2),-1,0,(a01.fieldI_XF1 - a01.fieldI_XF3 - a01.fieldI_XF2))  ");
	  sql.append("                     )/a01.fieldI_Y * 100,2))  as  field_DC_RATE ,   	    ");// --存放比率   
	  sql.append("        decode(a01.field_CREDIT,0,0,round(a01.field_OVER /  a01.field_CREDIT *100 ,2))  as field_OVER_RATE       	           		");//--逾放比率
	  sql.append(" from (  ");
	  sql.append("        select a01.hsien_id ,");
	  sql.append("               a01.hsien_name,");
	  sql.append("               a01.FR001W_output_order,");
	  sql.append("               a01.bank_type,  ");
	  sql.append("            	 ' ' AS  bank_no ,");
	  sql.append("               ' ' AS  bank_name,  ");
	  sql.append("               COUNT(*)  AS  COUNT_SEQ,                      ");
	  sql.append("        		 'A90'  as  field_SEQ,                          ");
	  sql.append("               SUM(field_OVER)      field_OVER,              ");
	  sql.append("               SUM(field_DEBIT)     field_DEBIT,             ");
	  sql.append("               SUM(field_CREDIT)    field_CREDIT,            ");
	  sql.append("               SUM(fieldI_XA)       fieldI_XA,               ");
	  sql.append("               SUM(fieldI_XB1)      fieldI_XB1,              ");
	  sql.append("               SUM(fieldI_XB2)      fieldI_XB2,              ");
	  sql.append("               SUM(fieldI_XC1)      fieldI_XC1,              ");
	  sql.append("               SUM(fieldI_XC2)      fieldI_XC2,              ");
	  sql.append("               SUM(fieldI_XD1)      fieldI_XD1,              ");
	  sql.append("               SUM(fieldI_XD2)      fieldI_XD2,              ");
	  sql.append("               SUM(fieldI_XE1)      fieldI_XE1,              ");
	  sql.append("               SUM(fieldI_XE2)      fieldI_XE2,              ");
	  sql.append("               SUM(fieldI_XF1)      fieldI_XF1,              ");
	  sql.append("               SUM(fieldI_XF3)      fieldI_XF3,              ");
	  sql.append("               SUM(fieldI_XF2)      fieldI_XF2,              ");
	  sql.append("               SUM(fieldI_Y)        fieldI_Y                 ");
	  sql.append(" 	   from ( select nvl(cd01.hsien_id,' ')       as  hsien_id ,   ");
	  sql.append("        				 nvl(cd01.hsien_name,'OTHER') as  hsien_name,    ");
	  sql.append("        				 cd01.FR001W_output_order     as  FR001W_output_order,bn01.bank_type,  ");
	  sql.append("        				 bn01.bank_no ,  bn01.bank_name,					                   ");
	  sql.append("                      round(sum(decode(a01.acc_code,'990000',amt,0)) /?,0)        as field_OVER,   ");
	  sql.append("                      round(sum(decode(a01.acc_code,'220000',amt,0)) /?,0)        as field_DEBIT,           ");
	  sql.append("                      round(sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /?,0) as  field_CREDIT,");  
	  sql.append("              		 round(sum(decode(bank_type,'6',decode(a01.acc_code,'120101',amt,'120102',amt,");
	  sql.append("                                           					  			'120200',amt,'120301',amt,");
	  sql.append("       			                					                    '120302',amt,'120700',amt,");
	  sql.append("                  		  	                           					'150200',amt,0)           ");
	  sql.append("                                       			,'7',decode(a01.acc_code,'120101',amt,'120102',amt,                       ");
	  sql.append("               		                      						  		 '120300',amt,'120401',amt,                                 ");
	  sql.append("                    	                    				  				 '120402',amt,'120700',amt,                               ");
	  sql.append("                       	                 							     '150200',amt,0))) /?,0)     as fieldI_XA,                  ");
	  sql.append("              		round(sum(decode(bank_type,'6',decode(a01.acc_code,'120401',amt,'120402',amt,0),'7',");
	  sql.append("                      decode(a01.acc_code,'120201',amt,'120202',amt,0))) /?,0)     as fieldI_XB1,  ");
	  sql.append("              		sum(decode(bank_type,'6',decode(a01.acc_code,'240305',amt,'251300',amt,'310800',amt, 0),'7',decode(a01.acc_code,'240205',amt, '310800',amt,0)))  as fieldI_XB2,  ");
	  sql.append(" 				        round(sum(decode(a01.acc_code,'120501',amt,'120502',amt,0)) /?,0)     as fieldI_XC1,                                                                                   ");
	  sql.append(" 				        round(sum(decode(bank_type,'6',decode(a01.acc_code,'240301',amt,'240302',amt,'240303',amt,'240304',amt,0),                                                            ");
	  sql.append("                                               '7',decode(a01.acc_code,'240201',amt,'240202',amt,'240203',amt,'240204',amt,0)) ) /?,0)     as fieldI_XC2,                           ");
	  sql.append("                      round(sum(decode(a01.acc_code,'120600',amt,0)) /?,0)              as fieldI_XD1,                                                                              ");
	  sql.append("                      round(sum(decode(a01.acc_code,'240200',amt,0)) /?,0)              as fieldI_XD2,                                                                              ");
	  sql.append("                      round(sum(decode(a01.acc_code,'150100',amt,0)) /?,0)              as fieldI_XE1,                                                                              ");
	  sql.append("                      round(sum(decode(a01.acc_code,'250100',amt,0)) /?,0)              as fieldI_XE2,                                                                              ");
	  sql.append("                      round(sum(decode(a01.acc_code,'310000',amt,'320000',amt,0)) /?,0) as fieldI_XF1,                                                                              ");
	  sql.append("                      round(sum(decode(a01.acc_code,'310800',amt,0)) /?,0)     		   as fieldI_XF3,                                                                                 ");
	  sql.append("                      round(sum(decode(a01.acc_code,'140000',amt,0)) /?,0)              as fieldI_XF2,                                                                              ");
	  sql.append("                      round((sum(decode(a01.acc_code,'220100',amt,'220200',amt,                                                                                                     ");
	  sql.append("                                                     '220300',amt,'220400',amt,                                                                                                     ");
	  sql.append("                                                     '220500',amt,'220600',amt,                                                                                                     ");
	  sql.append("                                                     '220700',amt,'220800',amt,                                                                                                     ");
	  sql.append("                                                     '220900',amt,'221000',amt,0))                                                                                                  ");
	  sql.append("                            - round(sum(decode(a01.acc_code,'220900',amt,0))/2,0)) /?,0) as fieldI_Y                                                                                ");
	  sql.append("       		  from  (select * from ").append(cd01Table).append(" cd01 where cd01.hsien_id <> 'Y') cd01                                                                                                      ");
	  sql.append("        		  left join (select * from wlx01 where m_year=? )wlx01 on wlx01.hsien_id=cd01.hsien_id ");
	  sql.append("        		  left join (select * from bn01 where m_year=? )bn01 on wlx01.bank_no=bn01.bank_no  ");
	  if(!"".equals(bankType)) {
		  sql.append("                                and bn01.bank_type=? ");
	  }
	  sql.append("        		  left join (select * from a01 where  a01.m_year  = ? and a01.m_month  = 12 ) a01                                                                                      ");
	  sql.append("        		            on  bn01.bank_no = a01.bank_code                                                                                                                            ");
	  sql.append("        		  group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_type,bn01.bank_no ,  bn01.BANK_NAME                                   ");
	  sql.append(" 		    ) a01                  ");
	  sql.append(" 		where  a01.bank_no <> ' '  ");
	  sql.append(" 		GROUP  BY a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order,a01.bank_type  ");
	  sql.append(" 	 ) a01");
	  sql.append("  ) T1 ");
	  sql.append(" Where T2.HSIEN_NAME=T1.hsien_name(+) ");
	  sql.append(" order by T2.FR001W_output_order ");*/
	
	  return sql.toString(); 
  }
  /***
   * 取得SQL參數.
   * 
   * @param bankType
   * @param u_year
   * @param unit
   * @param S_YEAR
   * @return
   */
  private static List getReportParameter(String bankType,String u_year,String unit,String S_YEAR) {
      List paramList = new ArrayList() ;
	  paramList.add(unit);
      paramList.add(unit);
      paramList.add(unit);
      paramList.add(u_year);
      paramList.add(u_year);
	  if(!"".equals(bankType)){
	      paramList.add(bankType);
	  }
	  paramList.add(S_YEAR);
	  return paramList ;
  }
}
