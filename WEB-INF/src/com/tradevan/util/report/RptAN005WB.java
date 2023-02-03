/*
 * Created on 2006/11/08 by ABYSS Brenda
 * 年底縣市別之全體農漁會信用部放款金額及放款平均餘額表
 *
 * 2006/12/18 by Abyss Brenda
 * 1.新增其他縣市別
 * 2.金額四捨五入至整數位
 *
 * 2006/12/22 by Abyss Brenda
 * 1.新增可區分農漁會
 * 2.新增選擇金額單位
 * 
 * 2010/04/14 by 2808
 * 1.因應縣市合併調整SQL
 * 2.修改查詢方式以PreparedStatement
 * 2010/11/5 by 2808
 * 1.修改縣市合併排序調整
 * 
 * 2013/11/22 fix 100年後農會報表格式 by 2968
 * 2014/12/24 fix 桃園縣升格調整 by 2968 
 */

package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.poifs.filesystem.*;
import java.io.*;
import java.util.*;
import java.sql.*;

import com.tradevan.util.DBManager;
import com.tradevan.util.Utility;
import com.tradevan.util.Utility_report;
import com.tradevan.util.dao.DataObject;



public class RptAN005WB {
    public static String createRpt(String S_YEAR, String bankType ,String tmUnit) {
        boolean debug = true;
        if(debug) System.out.println("RptAN005WB Start ...");

        //金額單位
        String unitName = tmUnit.substring(0,tmUnit.indexOf(";"));
        String unit = tmUnit.substring(tmUnit.indexOf(";") + 1);

        // 取得當下年月資料轉換西元年到民國年
        Calendar c = Calendar.getInstance();
        int nowYear = c.get(Calendar.YEAR) - 1911;
        int qYear = Integer.parseInt(S_YEAR);

        String errMsg = "";
        Connection conn = null;
        PreparedStatement pst = null;
        ResultSet rs = null;
        HashMap <String,String[]>dataMap = new HashMap<String,String[]>();  //存放結果
        String sqlCmd = null;
        long t_count = 0;  //總單位數
        long t_amount = 0;  //總放款金額
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

            String fileName = "年底縣市別之全體農漁會信用部放款金額及放款平均餘額表";
            if("100".equals(u_year) && "6".equals(bankType)){
                fileName = "年底縣市別之全體農會信用部放款金額及放款平均餘額表_100";
            }
            //input the standard report form
            System.out.println(fileName+".xls");

            finput = new FileInputStream(xlsDir +
                                   System.getProperty("file.separator") +
                                   fileName + ".xls");
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
            String title = " 年底縣市別之全體農漁會信用部放款金額及放款平均餘額表";
            if(bankType.equals("6")){
                title = " 年底縣市別之全體農會信用部放款金額及放款平均餘額表";
            }else if(bankType.equals("7")){
                title = " 年底縣市別之全體漁會信用部放款金額及放款平均餘額表";
            }
            cell.setCellValue(S_YEAR + title);

            //設定年月及單位資料============================================
            row = sheet.getRow(1);
            cell = row.getCell( (short) 0);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
            cell.setCellValue("單位：新台幣 " + unitName);

            //if(qYear >= nowYear){
            if(false) { //測試用=============
                row = sheet.getRow(3);
                cell = row.getCell( (short) 0);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
                cell.setCellValue("查無資料");
                sheet.addMergedRegion(new Region((short)3, (short)0,(short)3, (short)7));//跨行(第幾行,開始欄位數,跨幾行,結束欄位數)
            }else{

                /*ArrayList <String > cityLs = getCityList(u_year);
                StringBuffer sql = new StringBuffer () ;
                List paramList = new ArrayList() ;
                sql.append(getReportSql(bankType,u_year)) ;
                paramList = getReportParameter(u_year,S_YEAR,unit,bankType);
                List <DataObject> dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "hsien_name,count_seq,field_120000");
                System.out.println("layer 1 dbData.size()="+dbData.size());
                ArrayList <String> ls = new ArrayList<String> () ;*/
                StringBuffer sql = new StringBuffer () ;
                List paramList = new ArrayList() ;
                
                ArrayList <String >cityLs = getCityList(u_year) ;
             
                //DataObject bean = null;
                sql.append(getReportSql(bankType,u_year)) ;
                paramList = getReportParameter(u_year,S_YEAR,unit,bankType);
                List <DataObject>dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "hsien_name,count_seq,field_120000,field_120000_avg,fr001w_output_order");
                long total1 = 0l ;
                long total2 = 0l ;
                long total3 = 0l ;
                System.out.println("layer 1 dbData.size()="+dbData.size());
                int num = 0;
                int i = 0 ;
                for(DataObject bean : dbData) {
                    String dbHsienId = String.valueOf(bean.getValue("hsien_id"));
                    String dbHsienName = String.valueOf(bean.getValue("hsien_name")) ;
                    System.out.println("get city from sql :"+dbHsienName);
                    /*ls.add(dbHsienName) ;
                    String dbDetal[] = new String[3]; //存放內容：單位數、放款總額
                    dbDetal[0] = bean.getValue("count_seq")==null? "0" : bean.getValue("count_seq").toString(); 
                    String dbAmount = bean.getValue("field_120000")==null? "0" :bean.getValue("field_120000").toString();
                    dbDetal[1] = Utility_report.add(dbDetal[1], dbAmount);
                    dbDetal[2] = Utility_report.round(dbDetal[1], dbDetal[0], 2);
                    dataMap.put(dbHsienName, dbDetal);*/
                    String fr001w_output_order = bean.getValue("fr001w_output_order").toString() ;
                    String dbDetal[] = new String[3]; //存放內容：單位數、存款總額、存款平均
                    if(!"999".equals(fr001w_output_order)) {
                        dbDetal[0] = bean.getValue("count_seq")==null? "0" : bean.getValue("count_seq").toString(); //取得所有縣市的單位數
                        dbDetal[1] = bean.getValue("field_120000")==null? "0": bean.getValue("field_120000").toString(); //依查詢年度取得各縣市的放款總額
                        dbDetal[2] = bean.getValue("field_120000_avg")==null? "0": bean.getValue("field_120000_avg").toString();
                        total1 += Long.parseLong(dbDetal[0]) ;
                        total2 += Long.parseLong(dbDetal[1]) ;
                        total3 += Long.parseLong(dbDetal[2]) ;
                        dataMap.put(dbHsienName, dbDetal);
                        if("未歸屬縣市".equals(dbHsienName)) {
                            continue ;
                        }
                        if("100".equals(u_year) && ("6".equals(bankType)||"".equals(bankType))){
                            if("f".equals(dbHsienId)||"A".equals(dbHsienId)) num = 3;
                            else if("H".equals(dbHsienId)||"b".equals(dbHsienId)) num = 4;
                            else if("d".equals(dbHsienId)||"e".equals(dbHsienId)) num = 5;
                            else if("G".equals(dbHsienId)||"K".equals(dbHsienId)) num = 6;
                            else if("J".equals(dbHsienId)||"M".equals(dbHsienId)) num = 7;
                            else if("N".equals(dbHsienId)||"Q".equals(dbHsienId)) num = 8;
                            else if("P".equals(dbHsienId)||"I".equals(dbHsienId)) num = 9;
                            else if("T".equals(dbHsienId)||"X".equals(dbHsienId)) num = 10;
                            else if("V".equals(dbHsienId)||"O".equals(dbHsienId)) num = 11;
                            else if("C".equals(dbHsienId)||"Z".equals(dbHsienId)) num = 12;
                            else if("U".equals(dbHsienId)||"h".equals(dbHsienId)) num = 13;
                            row = sheet.getRow(num);
                            //f新北市、b臺中市、e高雄市、G宜蘭縣、J新竹縣、N彰化縣、P雲林縣、T屏東縣、V花蓮縣、C基隆市、U嘉義市
                            if("f".equals(dbHsienId)||"H".equals(dbHsienId)||"d".equals(dbHsienId)||"G".equals(dbHsienId)||"J".equals(dbHsienId)
                                    ||"N".equals(dbHsienId)||"P".equals(dbHsienId)||"T".equals(dbHsienId)||"V".equals(dbHsienId)||"C".equals(dbHsienId)
                                    ||"U".equals(dbHsienId)) {
                                
                                cell = row.getCell( (short) 0);
                                cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
                                cell.setCellValue(dbHsienName);
                                
                                cell = row.getCell( (short) 1);
                                cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
                                cell.setCellValue(dbDetal[0]);
                              
                                cell = row.getCell( (short) 2);
                                cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
                                cell.setCellValue(Utility.setCommaFormat(Utility_report.round(dbDetal[1], unit, 0)));
                                
                                cell = row.getCell( (short) 3);
                                cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
                                cell.setCellValue(Utility.setCommaFormat(Utility_report.round(dbDetal[2], unit, 0)));
                                
                            }else{
                                
                                cell = row.getCell( (short) 4);
                                cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
                                cell.setCellValue(dbHsienName);
                                  
                                cell = row.getCell( (short) 5);
                                cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
                                cell.setCellValue(dbDetal[0]);
                                  
                                cell = row.getCell( (short) 6);
                                cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
                                cell.setCellValue(Utility.setCommaFormat(Utility_report.round(dbDetal[1], unit, 0)));
                                  
                                cell = row.getCell( (short) 7);
                                cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
                                cell.setCellValue(Utility.setCommaFormat(Utility_report.round(dbDetal[2], unit, 0)));
                            }
                        }else{
                            row = sheet.getRow(num+3);
                            if (i % 2 == 0) {
                                
                                cell = row.getCell( (short) 0);
                                cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
                                cell.setCellValue(dbHsienName);
                                
                                cell = row.getCell( (short) 1);
                                cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
                                cell.setCellValue(dbDetal[0]);
                              
                                cell = row.getCell( (short) 2);
                                cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
                                cell.setCellValue(Utility.setCommaFormat(Utility_report.round(dbDetal[1], unit, 0)));
                                
                                cell = row.getCell( (short) 3);
                                cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
                                cell.setCellValue(Utility.setCommaFormat(Utility_report.round(dbDetal[2], unit, 0)));
                                
                            }else {
                                
                                cell = row.getCell( (short) 4);
                                cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
                                cell.setCellValue(dbHsienName);
                                  
                                cell = row.getCell( (short) 5);
                                cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
                                cell.setCellValue(dbDetal[0]);
                                  
                                cell = row.getCell( (short) 6);
                                cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
                                cell.setCellValue(Utility.setCommaFormat(Utility_report.round(dbDetal[1], unit, 0)));
                                  
                                cell = row.getCell( (short) 7);
                                cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
                                cell.setCellValue(Utility.setCommaFormat(Utility_report.round(dbDetal[2], unit, 0)));
                                  
                                num++;
                            }
                        }
                        i++ ;
                    }
                }
                
                //總計====================================
                row = sheet.getRow(15);
                cell = row.getCell( (short) 0);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
                cell.setCellValue("總計");
                cell = row.getCell( (short) 1);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
                cell.setCellValue(total1);
                cell = row.getCell( (short) 2);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
                cell.setCellValue(Utility.setCommaFormat(Utility_report.round(String.valueOf(total2), unit, 0)));
                cell = row.getCell( (short) 3);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
                cell.setCellValue(Utility.setCommaFormat(Utility_report.round(String.valueOf(total3), unit, 0)));
                
                //100未歸屬縣市/99其他====================================
                String dbDetal[] = (String[]) dataMap.get("未歸屬縣市");
                cell = row.getCell( (short) 4);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
                if("100".equals(u_year)){
                    cell.setCellValue("未歸屬縣市");
                }else{
                    cell.setCellValue("其它");
                }
                cell = row.getCell( (short) 5);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
                if(dbDetal == null){
                  cell.setCellValue("0");
                }else{
                  cell.setCellValue(dbDetal[0]);
                }
                cell = row.getCell( (short) 6);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
                if(dbDetal == null){
                  cell.setCellValue("0");
                }else{
                  cell.setCellValue(Utility.setCommaFormat(Utility_report.round(dbDetal[1], unit, 0)));
                }
                cell = row.getCell( (short) 7);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
                if(dbDetal == null){
                  cell.setCellValue("0");
                }else{
                  cell.setCellValue(Utility.setCommaFormat(Utility_report.round(dbDetal[2], unit, 0)));
                }
            }

            FileOutputStream fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + "年底縣市別之全體農漁會信用部放款金額及放款平均餘額表.xls");
            HSSFFooter footer = sheet.getFooter();
            footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
            footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));

            wb.write(fout);
            //儲存
            fout.close();
        }catch (Exception e) {
            System.out.println("//RptAN005WB createRpt() Have Error.....");
            e.printStackTrace();
            System.out.println("//-------------------------------------");
        }finally {
            try {
                conn.close();
            }catch (Exception sqlEx) {
                conn = null;
            }
        }

        if(debug) System.out.println("RptAN005WB End ...");
        return errMsg;
    }
  
  private static String getReportSql (String bankType,String u_year) {
	  StringBuffer sql = new StringBuffer() ;
	  if("100".equals(u_year)) {
		  sql.append(" select * from ");
		  sql.append(" (             ");
		  sql.append(" select a01.hsien_id ,a01.hsien_name,     ");                                                             
		  sql.append("        a01.FR001W_output_order,          ");
		  sql.append("        count(*) as count_seq,            ");   //--單位數
		  sql.append(" 	      sum(a01.field_120000) field_120000,");//--放款總額 
		  sql.append(" 	      round(sum(a01.field_120000) /count(*),0)  as field_120000_avg  "); //--放款平均餘額
		  sql.append(" from (    ");
		  sql.append(" 	  select nvl(cd01.hsien_id,' ')        as  hsien_id ,   ");                                        
		  sql.append("           nvl(cd01.hsien_name,'OTHER')  as  hsien_name, ");                                       
		  sql.append(" 			 cd01.FR001W_output_order      as  FR001W_output_order,");                                    
		  sql.append(" 			 bn01.bank_no ,  bn01.BANK_NAME,	                   ");
		  sql.append("    	   	 round(sum(amt)/?,0)  as field_120000              ");
		  sql.append(" 	  from  (select * from cd01 where cd01.hsien_id <> 'Y') cd01   ");
		  sql.append(" 	  left join wlx01 on wlx01.hsien_id=cd01.hsien_id ");
		  sql.append("                    and wlx01.m_year = ?          ");
		  sql.append(" 	  left join bn01 on wlx01.bank_no=bn01.bank_no  ");
		  if(!"".equals(bankType)) {
			  sql.append("                   and bn01.bank_type in (?) ");
		  }
		  sql.append("                   and bn01.m_year = wlx01.m_year ");
		  sql.append("                   and bn01.m_year = ? ");
		  sql.append(" 	  left join (select * from a01 where m_year = ? and m_month = 12 and acc_code in ('120000','120800','150300')) a01  ");
		  sql.append("       	   		 on  bn01.bank_no = a01.bank_code												");
		  sql.append(" 	  group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),     ");
		  sql.append(" 	        cd01.FR001W_output_order,bn01.bank_no ,  bn01.BANK_NAME     ");
		  sql.append(" 	  ) a01                                                             ");
		  sql.append(" where a01.bank_no <> ' '                                             ");
		  sql.append(" group by a01.hsien_id , a01.hsien_name,a01.FR001W_output_order       ");
		  sql.append(" )     ");//--各縣市統計
		  sql.append(" union ");
		  sql.append(" (     ");//--未歸屬縣市統計  
		  sql.append(" select 'Y' as hsien_id ,'未歸屬縣市' as hsien_name,                        ");                                         
		  sql.append("        '990' as FR001W_output_order,                                 ");
		  sql.append("        count(*) as count_seq,        "); //--單位數
		  sql.append(" 	   sum(a01.field_120000) field_120000,                   ");//--放款總額 
		  sql.append(" 	   round(sum(a01.field_120000) /count(*),0)  as field_120000_avg   ");//--放款平均餘額
		  sql.append(" from (select bank_code, 	        ");
		  sql.append("    	   		round(sum(amt) /?,0)  as field_120000   ");
		  sql.append(" 	  from  (select * from a01 where m_year = ? and m_month = 12 and acc_code in ('120000','120800','150300')) a01 ");
		  sql.append(" 	  where BANK_CODE NOT IN (SELECT BANK_NO FROM BN01 WHERE BANK_TYPE IN ('6','7')) ");
		  sql.append(" 	  group by a01.bank_code ");
		  sql.append(" 	  ) a01  ");
		  sql.append(" ) ");
		  sql.append(" union ");
		  sql.append(" ( ");//--全部總計
		  sql.append(" select ' ' as hsien_id ,'總計' as hsien_name,       ");                                                           
		  sql.append("        '999' as FR001W_output_order,                ");
		  sql.append("        count(*) as count_seq,               ");//--單位數
		  sql.append(" 	   sum(a01.field_120000) field_120000,   ");//--放款總額
		  sql.append(" 	   round(sum(a01.field_120000) /count(*),0)  as field_120000_avg ");//--放款平均餘額
		  sql.append(" from (	");//--其他縣市  
		  sql.append(" 	  select '','','',bank_code as bank_no,'', 	                                   ");
		  sql.append("    	   		 round(sum(amt) /?,0)  as field_120000                               ");
		  sql.append(" 	  from  (select * from a01 where m_year = ? and m_month = 12 and acc_code in ('120000','120800','150300')) a01");
		  sql.append(" 	  where BANK_CODE NOT IN (SELECT BANK_NO FROM BN01 WHERE BANK_TYPE IN ('6','7')) ");
		  sql.append(" 	  group by a01.bank_code  ");
		  sql.append(" 	  union ");//--各縣市
		  sql.append(" 	  select nvl(cd01.hsien_id,' ')        as  hsien_id ,         ");         
		  sql.append("           nvl(cd01.hsien_name,'OTHER')  as  hsien_name,    ");        
		  sql.append(" 			 cd01.FR001W_output_order      as  FR001W_output_order,");     
		  sql.append(" 			 bn01.bank_no ,  bn01.BANK_NAME,	                  ");
		  sql.append("    	   	 round(sum(amt) /?,0)  as field_120000            ");
		  sql.append(" 	  from  (select * from cd01 where cd01.hsien_id <> 'Y') cd01                                                   ");
		  sql.append(" 	  left join wlx01 on wlx01.hsien_id=cd01.hsien_id ");
		  sql.append("                    and wlx01.m_year = ?                                       ");
		  sql.append(" 	  left join bn01 on wlx01.bank_no=bn01.bank_no  ");
		  if(!"".equals(bankType)) {
			  sql.append("               and bn01.bank_type in (?) ");
		  }
		  sql.append("                   and bn01.m_year = wlx01.m_year ");
		  sql.append("                   and bn01.m_year = ? ");
		  sql.append(" 	  left join (select * from a01 where m_year = ? and m_month = 12 and acc_code in ('120000','120800','150300')) a01  ");
		  sql.append("       	   		 on  bn01.bank_no = a01.bank_code										");
		  sql.append(" 	  group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'), ");
		  sql.append(" 	        cd01.FR001W_output_order,bn01.bank_no ,  bn01.BANK_NAME	"); 
		  sql.append("      )a01			");
		  sql.append(" where bank_no <> ' ' ");
		  sql.append(" )                    ");
		  sql.append(" order by fr001w_output_order  ");
	  }else {
		  sql.append("SELECT * ");
		  sql.append("FROM   (SELECT a01.hsien_id, ");
		  sql.append("               a01.hsien_name, ");
		  sql.append("               a01.fr001w_output_order , ");
		  sql.append("               COUNT(*)                                   AS count_seq,      ");
		  sql.append("               SUM(a01.field_120000)                      field_120000, ");
		  sql.append("               Round(SUM(a01.field_120000) / COUNT(*), 0) AS field_120000_avg  ");
		  sql.append("        FROM   (SELECT Nvl(cd01.hsien_id, ' ')       AS hsien_id, ");
		  sql.append("                       Nvl(cd01.hsien_name, 'OTHER') AS hsien_name, ");
		  sql.append("                       cd01.fr001w_output_order      AS fr001w_output_order, ");
		  sql.append("                       bn01.bank_no, ");
		  sql.append("                       bn01.bank_name, ");
		  sql.append("                       Round(SUM(amt) / ?, 0)        AS field_120000 ");
		  sql.append("                FROM   (SELECT * ");
		  sql.append("                        FROM   cd01_99 cd01 ");
		  sql.append("                        WHERE  cd01.hsien_id <> 'Y') cd01 ");
		  sql.append("                       LEFT JOIN wlx01 ");
		  sql.append("                         ON wlx01.hsien_id = cd01.hsien_id ");
		  sql.append("                            AND wlx01.m_year = ? ");
		  sql.append("                       LEFT JOIN bn01 ");
		  sql.append("                         ON wlx01.bank_no = bn01.bank_no ");
		  if(!"".equals(bankType)) {
			  sql.append("                            AND bn01.bank_type IN ( ? ) ");
		  }
		  sql.append("                            AND bn01.m_year = wlx01.m_year ");
		  sql.append("                            AND bn01.m_year = ? ");
		  sql.append("                       LEFT JOIN (SELECT * ");
		  sql.append("                                  FROM   a01 ");
		  sql.append("                                  WHERE  m_year = ? ");
		  sql.append("                                         AND m_month = 12 ");
		  sql.append("                                         AND acc_code IN ( '120000', '120800', '150300' )) a01 ");
		  sql.append("                         ON bn01.bank_no = a01.bank_code ");
		  sql.append("                GROUP  BY Nvl(cd01.hsien_id, ' '), ");
		  sql.append("                          Nvl(cd01.hsien_name, 'OTHER'), ");
		  sql.append("                          cd01.fr001w_output_order, ");
		  sql.append("                          bn01.bank_no, ");
		  sql.append("                          bn01.bank_name) a01 ");
		  sql.append("        WHERE  a01.bank_no <> ' ' ");
		  sql.append("        GROUP  BY a01.hsien_id, ");
		  sql.append("                  a01.hsien_name, ");
		  sql.append("                  a01.fr001w_output_order) ");
		  sql.append("UNION ");
		  sql.append("( ");
		  sql.append("SELECT 'Y' AS hsien_id, ");
		  sql.append("       '未歸屬縣市' AS hsien_name, ");
		  sql.append("       '990' AS fr001w_output_order, ");
		  sql.append("       COUNT(*) AS count_seq,       ");
		  sql.append("       SUM(a01.field_120000) field_120000, ");
		  sql.append("       Round(SUM(a01.field_120000) / COUNT(*), 0) AS field_120000_avg ");
		  sql.append(" FROM   (SELECT bank_code, ");
		  sql.append("                Round(SUM(amt) / ?, 0) AS field_120000 ");
		  sql.append("         FROM   (SELECT * ");
		  sql.append("                 FROM   a01 ");
		  sql.append("                 WHERE  m_year = ? ");
		  sql.append("                        AND m_month = 12 ");
		  sql.append("                        AND acc_code IN ( '120000', '120800', '150300' )) a01 ");
		  sql.append("         WHERE  bank_code NOT IN (SELECT bank_no ");
		  sql.append("                                  FROM   bn01 ");
		  sql.append("                                  WHERE  bank_type IN ( '6', '7' )) ");
		  sql.append("         GROUP  BY a01.bank_code) a01) ");
		  sql.append("UNION ");
		  sql.append("( ");
		  sql.append("SELECT ' ' AS hsien_id, ");
		  sql.append("       '總計' AS hsien_name, ");
		  sql.append("       '999' AS fr001w_output_order, ");
		  sql.append("       COUNT(*) AS count_seq,       ");
		  sql.append("       SUM(a01.field_120000) field_120000, ");
		  sql.append("       Round(SUM(a01.field_120000) / COUNT(*), 0) AS field_120000_avg ");
		  sql.append(" FROM   ( ");

		  sql.append("        SELECT '', ");
		  sql.append("               '', ");
		  sql.append("               '', ");
		  sql.append("               bank_code AS bank_no, ");
		  sql.append("               '', ");
		  sql.append("               Round(SUM(amt) / ?, 0) AS field_120000 ");
		  sql.append("        FROM   (SELECT * ");
		  sql.append("                FROM   a01 ");
		  sql.append("                WHERE  m_year = ? ");
		  sql.append("                       AND m_month = 12 ");
		  sql.append("                       AND acc_code IN ( '120000', '120800', '150300' )) a01 ");
		  sql.append("        WHERE  bank_code NOT IN (SELECT bank_no ");
		  sql.append("                                 FROM   bn01 ");
		  sql.append("                                 WHERE  bank_type IN ( '6', '7' )) ");
		  sql.append("        GROUP  BY a01.bank_code ");
		  sql.append("         UNION ");
		  sql.append("         SELECT Nvl(cd01.hsien_id, ' ')       AS hsien_id, ");
		  sql.append("                Nvl(cd01.hsien_name, 'OTHER') AS hsien_name, ");
		  sql.append("                cd01.fr001w_output_order      AS fr001w_output_order, ");
		  sql.append("                bn01.bank_no, ");
		  sql.append("                bn01.bank_name, ");
		  sql.append("                Round(SUM(amt) / ?, 0)        AS field_120000 ");
		  sql.append("         FROM   (SELECT * ");
		  sql.append("                 FROM   cd01_99 cd01 ");
		  sql.append("                 WHERE  cd01.hsien_id <> 'Y') cd01 ");
		  sql.append("                LEFT JOIN wlx01 ");
		  sql.append("                  ON wlx01.hsien_id = cd01.hsien_id ");
		  sql.append("                     AND wlx01.m_year = ? ");
		  sql.append("                LEFT JOIN bn01 ");
		  sql.append("                  ON wlx01.bank_no = bn01.bank_no ");
		  if(!"".equals(bankType)) {
			  sql.append("                     AND bn01.bank_type IN ( ? ) ");
		  }
		  sql.append("                     AND bn01.m_year = wlx01.m_year ");
		  sql.append("                     AND bn01.m_year = ? ");
		  sql.append("                LEFT JOIN (SELECT * ");
		  sql.append("                           FROM   a01 ");
		  sql.append("                           WHERE  m_year = ? ");
		  sql.append("                                  AND m_month = 12 ");
		  sql.append("                                  AND acc_code IN ( '120000', '120800', '150300' )) a01 ");
		  sql.append("                  ON bn01.bank_no = a01.bank_code ");
		  sql.append("         GROUP  BY Nvl(cd01.hsien_id, ' '), ");
		  sql.append("                   Nvl(cd01.hsien_name, 'OTHER'), ");
		  sql.append("                   cd01.fr001w_output_order, ");
		  sql.append("                   bn01.bank_no, ");
		  sql.append("                   bn01.bank_name)a01 ");
		  sql.append(" WHERE  bank_no <> ' ') ");
		  sql.append("ORDER  BY fr001w_output_order ");
	  }
	  return sql.toString() ;
  }
  /***
   * SQL的parameter.
   * 
   * @param u_year
   * @param S_YEAR
   * @param unit
   * @param bankType
   * @return
   */
  private static List getReportParameter(String u_year,String S_YEAR,String unit,String bankType) {
	  List paramList = new ArrayList() ;
	  if(!"".equals(bankType)) {
			paramList.add(unit     );
			paramList.add(u_year   );
			paramList.add(bankType );
			paramList.add(u_year   );
			paramList.add(S_YEAR   );
			paramList.add(unit     );
			paramList.add(S_YEAR   );
			paramList.add(unit     );
			paramList.add(S_YEAR   );
			paramList.add(unit     );
			paramList.add(u_year   );
			paramList.add(bankType );
			paramList.add(u_year   );
			paramList.add(S_YEAR   );
		}else {
			paramList.add(unit   );
			paramList.add(u_year );
			paramList.add(u_year );
			paramList.add(S_YEAR );
			paramList.add(unit   );
			paramList.add(S_YEAR );
			paramList.add(unit   );
			paramList.add(S_YEAR );
			paramList.add(unit   );
			paramList.add(u_year );
			paramList.add(u_year );
			paramList.add(S_YEAR );
		}
	  return paramList ;
  }
  private static ArrayList<String> getCityList(String uYear) throws Exception {
	  StringBuffer sql = new StringBuffer () ;
      List paramList = new ArrayList() ;
      if("100".equals(uYear)) {
    	  sql.append(" select HSIEN_NAME  from cd01 order by FR001W_OUTPUT_ORDER ");
      }else {
    	  sql.append(" select HSIEN_NAME  from cd01_99 order by fr001W_OUTPUT_ORDER ");
      }
      ArrayList <String > cityLs = new ArrayList<String>() ;
      List <DataObject> dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "HSIEN_NAME");
      for(DataObject b : dbData ) {
    	  cityLs.add(b.getValue("hsien_name").toString()) ;
      }
      return cityLs ;
  }
}
