/*
 *  98.06.10 create 全體農漁會信用部各會員別放款金額一覽表 by 2295
 *  98.06.16 add 傳至檢查局 by 2295
 *  99.09.10 fix 根據查詢年度.100年以後取得新縣市別.100年以前取得舊縣市別 
  			    使用PreparedStatement;並列印轉換後的SQL;套用QueryDB_SQLParam by 2295
 * 102.02.05 add a02.amt_name by 2295    
 * 102.11.19 add 原QueryDB改套用QueryDB_SQLParam by 2295 
 * 103.09.01 fix 計算公式 by 2295
 */
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import java.io.*;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR054W {	 
	 public static String createRpt(String m_year,String m_month,String unit,String febxlsFlag,HSSFWorkbook wb) {    

	    String errMsg = "";	    
	    String sqlCmd = ""; 
	    String condition = "";
	    List dbData = null;	    
	    int rowNum=0;
	    DataObject bean = null;
	    DataObject bean_sub = null;
        DataObject bean_sub1 = null;
	    reportUtil reportUtil = new reportUtil();
		HSSFCellStyle cs_right = null; 
		HSSFCellStyle cs_center = null;
	    HSSFCellStyle cs_left = null;
	    String hsien_name="";//縣市別
	    String bank_name="";	  
	    String amt992140="";//--會員-放款總金額(A)
	    String amt992510="";//--會員-逾放金額(B)
	    String b_a_rate="";//--會員-比例(B/A)		   
	    String amt990410="";//--贊助會員-放款總金額(C)
	    String amt992520="";//--贊助會員-逾放金額(D)
	    String d_c_rate="";//--贊助會員-比例(D/C)		
	    String amt990610="";//--非會員-放款總金額(E)
	    String amt990611="";//--非會員-對縣市政府貸款
	    String amt992530="";//--非會員-逾放金額(F)
	    String f_e_rate="";//--非會員-比例(F/E)		
		String amt120000="";//--放款總額(G)
		String b_d_f_g_rate="";//--比例(B+D+F/G)
		
	    String unit_name="";
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
	      if(febxlsFlag.equals("")){//原全體農漁會信用部各會員別放款金額一覽表
	      	//input the standard report form      
	      	finput = new FileInputStream(xlsDir +System.getProperty("file.separator") +"全體農漁會信用部各會員別放款金額一覽表.xls");      
	      
	      	//設定FileINputStream讀取Excel檔
	      	POIFSFileSystem fs = new POIFSFileSystem(finput);	      
	        wb = new HSSFWorkbook(fs);
	      }else{
	      	System.out.println("檢查局格式:全體農漁會信用部各會員別放款金額一覽表");
	      }
	      HSSFSheet sheet = wb.getSheetAt((febxlsFlag.equals("")?0:2)); //讀取第一個工作表，宣告其為sheet
	      HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
	      //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
	      //sheet.setAutobreaks(true); //自動分頁

	      //設定頁面符合列印大小
	      sheet.setAutobreaks(false);
	      ps.setScale( (short) 95); //列印縮放百分比	      
	      ps.setPaperSize( (short) 9); //設定紙張大小 A4
	      //wb.setSheetName(0,"test");
	      
	      if(febxlsFlag.equals("")) finput.close();

	      HSSFRow row = null; //宣告一列
	      HSSFCell cell = null; //宣告一個儲存格	      
	      cs_right = reportUtil.getRightStyle(wb);
	      cs_center = reportUtil.getDefaultStyle(wb);
	      cs_left = reportUtil.getNoBorderLeftStyle(wb);
	      m_year = String.valueOf(Integer.parseInt(m_year));
	      m_month = String.valueOf(Integer.parseInt(m_month));
	      
	      unit_name = Utility.getUnitName(unit);//取得單位名稱
	      String cd01_table = "";
          String wlx01_m_year = "";
	      StringBuffer sql = new StringBuffer();
	      List paramList = new ArrayList();//傳入參數
	      //99.09.10 add 查詢年度100年以前.縣市別不同===============================
	      cd01_table = (Integer.parseInt(m_year) < 100)?"cd01_99":""; 
	      wlx01_m_year = (Integer.parseInt(m_year) < 100)?"99":"100"; 
	      //=====================================================================    
	      sql.append(" select nvl(cd01.hsien_id,' ')  as  hsien_id ,"); 
	      sql.append(" 		  nvl(cd01.hsien_name,'OTHER') as  hsien_name,"); 
	      sql.append(" 		  cd01.fr001w_output_order  as fr001w_output_order,"); 
	      sql.append("		  bn01.bank_no,bn01.bank_name,");
	      sql.append("		  round(sum(decode(a.acc_code,'992140',amt,0)) /?,0) as amt992140,");//--會員-放款總金額(A)
	      sql.append("		  round(sum(decode(a.acc_code,'992510',amt,0)) /?,0) as amt992510,");//--會員-逾放金額(B)
	      sql.append("        decode(sum(decode(a.acc_code,'992140',amt,0)),0,0,round(sum(decode(a.acc_code,'992510',amt,0)) /  sum(decode(a.acc_code,'992140',amt,0)) *100 ,2)) as b_a_rate,");//--B/A
	      //sql.append("		  decode(sum(decode(a.acc_code,'992140',amt,0)),0,0,round(sum(decode(a.acc_code,'992510',amt,0)) *10000/sum(decode(a.acc_code,'992140',amt,0)) ,0)) as b_a_rate,");//--B/A	   
	      sql.append("		  round(sum(decode(a.acc_code,'990410',amt,0)) /?,0) as amt990410,");//--贊助會員-放款總金額(C) 
	      sql.append("		  round(sum(decode(a.acc_code,'992520',amt,0)) /?,0) as amt992520,");//--贊助會員-逾放金額(D) 
	      sql.append("        decode(sum(decode(a.acc_code,'990410',amt,0)),0,0,round(sum(decode(a.acc_code,'992520',amt,0)) /  sum(decode(a.acc_code,'990410',amt,0)) *100 ,2)) as d_c_rate,");//--D/C
	      //sql.append("		  decode(sum(decode(a.acc_code,'990410',amt,0)),0,0,round(sum(decode(a.acc_code,'992520',amt,0)) *10000/sum(decode(a.acc_code,'990410',amt,0)) ,0)) as d_c_rate,");//--D/C
	      sql.append("		  round(sum(decode(a.acc_code,'990610',amt,0)) /?,0) as amt990610,");//--非會員-放款總金額(E) 
	      sql.append("		  round(sum(decode(a.acc_code,'990611',amt,0)) /?,0) as amt990611,");//--非會員-對縣市政府貸款
	      sql.append("		  round(sum(decode(a.acc_code,'992530',amt,0)) /?,0) as amt992530,");//--非會員-逾放金額(F)
	      sql.append("        decode(sum(decode(a.acc_code,'990610',amt,0)),0,0,round(sum(decode(a.acc_code,'992530',amt,0)) /  sum(decode(a.acc_code,'990610',amt,0)) *100 ,2)) as f_e_rate,");//--F/E
	      //sql.append("		  decode(sum(decode(a.acc_code,'990610',amt,0)),0,0,round(sum(decode(a.acc_code,'992530',amt,0)) *10000/sum(decode(a.acc_code,'990610',amt,0)) ,0)) as f_e_rate,");//--F/E
	      sql.append("		  round(sum(decode(a.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /?,0) as amt120000,");//--放款總額(G)
	      sql.append("		  round(sum(decode(a.acc_code,'992510',amt,'992520',amt,'992530',amt,0)) /?,0) as b_d_f,");//--B+D+F
	      sql.append("        decode(sum(decode(a.acc_code,'120000',amt,'120800',amt,'150300',amt,0)),0,0,round(sum(decode(a.acc_code,'992510',amt,'992520',amt,'992530',amt,0)) /  sum(decode(a.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) *100 ,2)) as b_d_f_g_rate");//--B+D+F/G
	      //sql.append("		  decode(sum(decode(a.acc_code,'120000',amt,'120800',amt,'150300',amt,0)),0,0,round(sum(decode(a.acc_code,'992510',amt,'992520',amt,'992530',amt,0))*10000/sum(decode(a.acc_code,'120000',amt,'120800',amt,'150300',amt,0)),0)) as b_d_f_g_rate");//--B+D+F/G 	   
	      sql.append(" from (select * from "+cd01_table+" cd01 where cd01.hsien_id <> 'Y') cd01"); 
	      sql.append(" left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id");	      
	      sql.append(" left join (select * from bn01 where m_year=?)bn01 on wlx01.bank_no=bn01.bank_no");	      
	      sql.append(" left join (select * from a99"); 
	      sql.append(" where a99.m_year=? and a99.m_month=?");	      
	      sql.append(" and acc_code in ('992140','992510','992520','992530')");
	      sql.append(" union");
	      sql.append(" select m_year,m_month,bank_code,acc_code,amt from a02");
	      sql.append(" where a02.m_year=? and a02.m_month= ?");	      
	      sql.append(" and acc_code in ('990410','990610','990611')"); 
	      sql.append(" union");
	      sql.append(" select * from a01");
	      sql.append(" where a01.m_year=? and a01.m_month= ?");	      
	      sql.append(" and acc_code in ('120000','120800','150300')");			  
	      sql.append(" )a on  bn01.bank_no = a.bank_code");
	      sql.append(" where bn01.bank_no <> ' '");  
	      sql.append(" group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.fr001w_output_order,bn01.bank_no,bn01.bank_name"); 
	      sql.append(" order by cd01.fr001w_output_order");
	      paramList.add(unit);
          paramList.add(unit);
          paramList.add(unit);
          paramList.add(unit);
          paramList.add(unit);
          paramList.add(unit);
          paramList.add(unit);
          paramList.add(unit);
          paramList.add(unit);
          paramList.add(wlx01_m_year);
          paramList.add(wlx01_m_year);
          paramList.add(m_year);
          paramList.add(m_month);
          paramList.add(m_year);
          paramList.add(m_month);
          paramList.add(m_year);
          paramList.add(m_month);
	      dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"amt992140,amt992510,b_a_rate,amt990410,amt992520,d_c_rate,amt990610,amt990611,amt992530,f_e_rate,amt120000,b_d_f,b_d_f_g_rate");
	      System.out.println("dbData.size=" + dbData.size());	      
	      //設定報表表頭資料============================================
	   	  row = sheet.getRow(1);
	   	  cell = row.getCell( (short) 0);	   	  
	   	  cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	      if (dbData != null && dbData.size() != 0) {
		   	  cell.setCellValue("中華民國"+m_year+"年"+m_month+"月");
		   	  cell = row.getCell( (short) 12);	   	  
		   	  cell.setEncoding(HSSFCell.ENCODING_UTF_16);		   	
		   	  cell.setCellValue("單位:新臺幣 " + unit_name + ",%");  
		   	  rowNum = 3;	      	  
	      	  for(int i=0;i<dbData.size();i++){	      	      
	      	      bean = (DataObject)dbData.get(i);	      	    
	      	      hsien_name = String.valueOf(bean.getValue("hsien_name"));//縣市別
	      	      bank_name = String.valueOf(bean.getValue("bank_name"));//農漁會名稱	  
	      	      amt992140=(bean.getValue("amt992140") == null)?"":(bean.getValue("amt992140")).toString();//--會員-放款總金額(A)
	      	      amt992510=(bean.getValue("amt992510") == null)?"":(bean.getValue("amt992510")).toString();//--會員-逾放金額(B)	      	      		   
	      	      b_a_rate=(bean.getValue("b_a_rate") == null)?"":(bean.getValue("b_a_rate")).toString();//--B/A
	      	      amt990410=(bean.getValue("amt990410") == null)?"":(bean.getValue("amt990410")).toString();//--贊助會員-放款總金額(C) 
	      	      amt992520=(bean.getValue("amt992520") == null)?"":(bean.getValue("amt992520")).toString();//--贊助會員-逾放金額(D)	      	      		   
	      	      d_c_rate=(bean.getValue("d_c_rate") == null)?"":(bean.getValue("d_c_rate")).toString();//--D/C
	      	      amt990610=(bean.getValue("amt990610") == null)?"":(bean.getValue("amt990610")).toString();//--非會員-放款總金額(E) 
	      	      amt990611=(bean.getValue("amt990611") == null)?"":(bean.getValue("amt990611")).toString();//--非會員-對縣市政府貸款	      	      		   
	      	      amt992530=(bean.getValue("amt992530") == null)?"":(bean.getValue("amt992530")).toString();//--非會員-逾放金額(F)
	      	      f_e_rate=(bean.getValue("f_e_rate") == null)?"":(bean.getValue("f_e_rate")).toString();//--F/E 
	      	      amt120000=(bean.getValue("amt120000") == null)?"":(bean.getValue("amt120000")).toString();//--放款總額(G)
	      	      b_d_f_g_rate=(bean.getValue("b_d_f_g_rate") == null)?"":(bean.getValue("b_d_f_g_rate")).toString();//--B+D+F/G 
	      	      //b_a_rate = String.valueOf(Float.parseFloat(Utility.getRound(b_a_rate,"10"))/10);
	      	      //d_c_rate = String.valueOf(Float.parseFloat(Utility.getRound(d_c_rate,"10"))/10);
	      	      //f_e_rate = String.valueOf(Float.parseFloat(Utility.getRound(f_e_rate,"10"))/10);
	      	      //b_d_f_g_rate = String.valueOf(Float.parseFloat(Utility.getRound(b_d_f_g_rate,"10"))/10);
	      	      rowNum++;
	    	  	  row = sheet.createRow(rowNum);
   	    	  	  //列印各機構明細資料
				  for(int cellcount=0;cellcount<=13;cellcount++){			 	      
			 	     cell=row.createCell((short)cellcount);			 		
			    	 cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			    	 cell.setCellStyle(cs_right);		   
			    	 if(cellcount == 0) cell.setCellValue(hsien_name);//縣市別
			    	 if(cellcount == 1) cell.setCellValue(bank_name);//農漁會名稱
			    	 if(cellcount == 2) cell.setCellValue(Utility.setCommaFormat(amt992140));//會員-放款總金額(A) 
			    	 if(cellcount == 3) cell.setCellValue(Utility.setCommaFormat(amt992510));//會員-逾放金額(B)
			    	 if(cellcount == 4) cell.setCellValue(Utility.setCommaFormat(b_a_rate));//會員-比例(B/A)		    	  
			    	 if(cellcount == 5) cell.setCellValue(Utility.setCommaFormat(amt990410));//贊助會員-放款總金額(C)
			    	 if(cellcount == 6) cell.setCellValue(Utility.setCommaFormat(amt992520));//贊助會員-逾放金額(D)
			    	 if(cellcount == 7) cell.setCellValue(Utility.setCommaFormat(d_c_rate));//贊助會員-比例(D/C)
			    	 if(cellcount == 8) cell.setCellValue(Utility.setCommaFormat(amt990610));///非會員-放款總金額(E) 
			    	 if(cellcount == 9) cell.setCellValue(Utility.setCommaFormat(amt990611));//非會員-對縣市政府貸款
			    	 if(cellcount == 10) cell.setCellValue(Utility.setCommaFormat(amt992530));//非會員-逾放金額(F)
			    	 if(cellcount == 11) cell.setCellValue(Utility.setCommaFormat(f_e_rate));//非會員-比例(F/E)	 
			    	 if(cellcount == 12) cell.setCellValue(Utility.setCommaFormat(amt120000));//放款總額(G)
			    	 if(cellcount == 13) cell.setCellValue(Utility.setCommaFormat(b_d_f_g_rate));//比例B+D+F/G
			    	 
				  }//end of cellcount
				      
				     
	      	  }//end of bean
	      }else{ //end of else dbData.size() != 0
	      	 cell.setCellValue("中華民國"+m_year+"年"+m_month+"月無資料存在");
	      }
	     
	      if(febxlsFlag.equals("")){//原全體農漁會信用部各會員別放款金額一覽表	      	
	         FileOutputStream fout = null;     
	         fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + "全體農漁會信用部各會員別放款金額一覽表.xls");
	     
	         HSSFFooter footer = sheet.getFooter();
	         footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
	         footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
	         wb.write(fout);
	         //儲存
	         fout.close();
	      }
	    }catch (Exception e) {
	      System.out.println("RptFR054W.createRpt Error:" + e + e.getMessage());
	    }
	    
	    return errMsg;
	  }
}
