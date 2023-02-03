/*
 * 97.06.18 create 應予評估資產彙總表
 * 99.04.28 fix 縣市合併問題 and sql injection by 2808
 *100.02.22 fix 中央存保無法下載A10文字檔 by 2295 
 *104.03.10 fix A10增加申報欄位  by2968
 *106.05.04 add 報表增加非授信資產可能遭受損失(投資及其他資產)【註5】 by 2295
 */
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import java.io.*;
import java.util.*;

import com.tradevan.util.DownLoad;
import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR047W {
	private static String[][] print_table = {};
	final private static String[][] print_table1 = {{"loan2_amt",   "990001"},
											       {"loan3_amt",   "990002"},	
												   {"loan4_amt",   "990003"},	
												   {"loan_sum",    "990004"},	
												   {"invest2_amt", "990005"},
											       {"invest3_amt", "990006"},	
												   {"invest4_amt", "990007"},	
												   {"invest_sum",  "990008"},
												   {"other2_amt",  "990009"},
											       {"other3_amt",  "990010"},	
												   {"other4_amt",  "990011"},	
												   {"other_sum",   "990012"},
												   {"type2_sum",   "990013"},
											       {"type3_sum",   "990014"},	
												   {"type4_sum",   "990015"},	
												   {"type_sum",    "990016"},
										};
	final private static String[][] print_table2 = {{"loan1_amt",   "990001"},
												   {"loan2_amt",   "990002"},
											       {"loan3_amt",   "990003"},	
												   {"loan4_amt",   "990004"},	
												   {"loan_sum",    "990005"},
												   {"invest1_amt", "990006"},
												   {"invest2_amt", "990007"},
											       {"invest3_amt", "990008"},	
												   {"invest4_amt", "990009"},	
												   {"invest_sum",  "990010"},
												   {"other1_amt",  "990011"},
												   {"other2_amt",  "990012"},
											       {"other3_amt",  "990013"},	
												   {"other4_amt",  "990014"},	
												   {"other_sum",   "990015"},
												   {"type1_sum",   "990016"},
												   {"type2_sum",   "990017"},
											       {"type3_sum",   "990018"},	
												   {"type4_sum",   "990019"},	
												   {"type_sum",    "990020"},
												   {"loan1_baddebt",   "990021"},
												   {"loan2_baddebt",   "990022"},
											       {"loan3_baddebt",   "990023"},	
												   {"loan4_baddebt",   "990024"},	
												   {"loan_baddebt_sum",    "990025"},
												   {"build1_baddebt",   "990026"},
												   {"build2_baddebt",   "990027"},
											       {"build3_baddebt",   "990028"},	
												   {"build4_baddebt",   "990029"},	
												   {"build_baddebt_sum",    "990030"},
												   {"above_loan1_amt",   "990031"},
												   {"above_loan2_amt",   "990032"},
											       {"above_loan3_amt",   "990033"},	
												   {"above_loan4_amt",   "990034"},	
												   {"above_loan_sum",   "990035"},
												   {"amt990036",   "990036"},
												   {"amt990037",   "990037"},
												   {"baddebt_noenough", "990038"},
												   {"baddebt_104",   "990039"},
											       {"baddebt_105",   "990040"},	
												   {"baddebt_106",   "990041"},	
												   {"baddebt_107",   "990042"},
												   {"baddebt_108",   "990043"},
												   {"property_loss", "990044"},/*106.05.04 add*/
												   
	};	
	 public static String createRpt(String M_YEAR, String M_MONTH,String unit,String bank_code ) {    

	    String errMsg = "";
	    List dbData = null;
	    String sqlCmd = "";    
	    int rowNum=0;
	    DataObject bean = null;
	    reportUtil reportUtil = new reportUtil();
		HSSFCellStyle cs_right = null; 
		HSSFCellStyle cs_center = null;
		HSSFCellStyle cs_nbLeft = null;
	   
	    String data_year="";
	    String data_month="";
	    String bank_name="";
	    String loan1_amt="";
	    String loan2_amt="";//放款-列二類
	  	String loan3_amt="";//放款-列三類     
	  	String loan4_amt="";//放款-列四類   
	  	String loan_sum="";//放款-合計  
	  	String invest1_amt="";
	  	String invest2_amt="";//投資-列二類   
	  	String invest3_amt="";//投資-列三類   
	  	String invest4_amt="";//投資-列四類   
	  	String invest_sum="";//投資-合計
	  	String other1_amt="";
	  	String other2_amt="";//其他-列二類    
	  	String other3_amt="";//其他-列三類    
	  	String other4_amt="";//其他-列四類    
	  	String other_sum="";//其他-合計
	  	String type1_sum="";
	  	String type2_sum="";//列二類合計
	  	String type3_sum="";//列三類合計
	  	String type4_sum="";//列三類合計
	  	String type_sum ="";//合計
	  	String loan1_baddebt     = "";
	  	String loan2_baddebt     = "";
	  	String loan3_baddebt	 = "";
	  	String loan4_baddebt	 = "";
	  	String loan_baddebt_sum  = "";
	  	String build1_baddebt    = "";
	  	String build2_baddebt    = "";
	  	String build3_baddebt	 = "";
	  	String build4_baddebt	 = "";
	  	String build_baddebt_sum = "";
	  	String above_loan1_amt   = "";
	  	String above_loan2_amt   = "";
	  	String above_loan3_amt	 = "";
	  	String above_loan4_amt	 = "";
	  	String above_loan_sum    = "";
	  	String amt990036         = "";
	  	String amt990037         = "";
	  	String baddebt_noenough  = "";
	  	String baddebt_104       = "";
	  	String baddebt_105	     = "";
	  	String baddebt_106	     = "";
	  	String baddebt_107       = "";
	  	String baddebt_108       = "";
	  	String baddebt_flag      = "";
	  	String baddebt_delay     = "";
	  	String property_loss     = "";
	  	StringBuffer sql = new StringBuffer() ;
	  	ArrayList paramList = new ArrayList() ;
	  	String u_year = "100" ;
	  	if(M_YEAR==null || Integer.parseInt(M_YEAR)<=99 ) {
	  		u_year ="99" ;
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
	      finput = new FileInputStream(xlsDir +System.getProperty("file.separator") +"應予評估資產彙總表.xls");
	      System.out.println("RptFR047W.java="+xlsDir +System.getProperty("file.separator") +"應予評估資產彙總表.xls");
	      //設定FileINputStream讀取Excel檔
	      POIFSFileSystem fs = new POIFSFileSystem(finput);
	      HSSFWorkbook wb = new HSSFWorkbook(fs);
	      HSSFSheet sheet = wb.getSheetAt(0); //讀取第一個工作表，宣告其為sheet
	      HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
	      //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
	      //sheet.setAutobreaks(true); //自動分頁

	      //設定頁面符合列印大小
	      sheet.setAutobreaks(false);
	      ps.setScale( (short) 75); //列印縮放百分比	      
	      ps.setPaperSize( (short) 9); //設定紙張大小 A4
	      //wb.setSheetName(0,"test");
	      finput.close();

	      HSSFRow row = null; //宣告一列
	      HSSFCell cell = null; //宣告一個儲存格

	      short i = 0;
	      short y = 0;
	      cs_right = reportUtil.getRightStyle(wb);
	      cs_center = reportUtil.getDefaultStyle(wb);
	      cs_nbLeft = reportUtil.getNoBorderLeftStyle(wb);
	      String unit_name = Utility.getUnitName(unit); 
	   	  /*
	      select m_year,m_month,bank_name,loan2_amt,loan3_amt,loan4_amt,loan2_amt+loan3_amt+loan4_amt as loan_sum,
	      		 invest2_amt,invest3_amt,invest4_amt,invest2_amt+invest3_amt+invest4_amt as invest_sum,
	      		 other2_amt,other3_amt,other4_amt,other2_amt+other3_amt+other4_amt as other_sum,
	      		 loan2_amt+invest2_amt+other2_amt as type2_sum,
	      		 loan3_amt+invest3_amt+other3_amt as type3_sum,
	      		 loan4_amt+invest4_amt+other4_amt as type4_sum,
	      		 loan2_amt+invest2_amt+other2_amt+loan3_amt+invest3_amt+other3_amt+loan4_amt+invest4_amt+other4_amt as type_sum
          from a10 left join bn01 on a10.bank_code = bn01.bank_no
          where m_year=97
          and m_month=6
          and bank_code='6030016'  
          */
	      
	      sql.append(" select a10.m_year,m_month,bank_name,"
	    		 + " round(loan1_amt/?,0) as loan1_amt,"
	      		 + " round(loan2_amt/?,0) as loan2_amt,"
				 + " round(loan3_amt/?,0) as loan3_amt,"
				 + " round(loan4_amt/?,0) as loan4_amt,"
				 + " round((loan1_amt+loan2_amt+loan3_amt+loan4_amt)/?,0) as loan_sum,"
				 + " round(invest1_amt/?,0) as invest1_amt,"
	      	     + " round(invest2_amt/?,0) as invest2_amt,"
				 + " round(invest3_amt/?,0) as invest3_amt,"
				 + " round(invest4_amt/?,0) as invest4_amt,"
				 + " round((invest1_amt+invest2_amt+invest3_amt+invest4_amt)/?,0) as invest_sum,"
				 + " round(other1_amt/?,0) as other1_amt,"
				 + " round(other2_amt/?,0) as other2_amt,"
				 + " round(other3_amt/?,0) as other3_amt,"
				 + " round(other4_amt/?,0) as other4_amt,"
				 + " round((other1_amt+other2_amt+other3_amt+other4_amt)/?,0) as other_sum, "
				 + " round((loan1_amt+invest1_amt+other1_amt)/?,0) as type1_sum,"
				 + " round((loan2_amt+invest2_amt+other2_amt)/?,0) as type2_sum,"
				 + " round((loan3_amt+invest3_amt+other3_amt)/?,0) as type3_sum,"
				 + " round((loan4_amt+invest4_amt+other4_amt)/?,0) as type4_sum,"
				 + " round((loan1_amt+invest1_amt+other1_amt+loan2_amt+invest2_amt+other2_amt+loan3_amt+invest3_amt+other3_amt+loan4_amt+invest4_amt+other4_amt)/?,0) as type_sum,"
				 + " round(loan1_baddebt/?,0) as loan1_baddebt,"
				 + " round(loan2_baddebt/?,0) as loan2_baddebt,"
				 + " round(loan3_baddebt/?,0) as loan3_baddebt,"
				 + " round(loan4_baddebt/?,0) as loan4_baddebt,"
				 + " round((loan1_baddebt+loan2_baddebt+loan3_baddebt+loan4_baddebt)/?,0) as loan_baddebt_sum,"
				 + " round(build1_baddebt/?,0) as build1_baddebt,"
				 + " round(build2_baddebt/?,0) as build2_baddebt,"
				 + " round(build3_baddebt/?,0) as build3_baddebt,"
				 + " round(build4_baddebt/?,0) as build4_baddebt,"
				 + " round((build1_baddebt+build2_baddebt+build3_baddebt+build4_baddebt)/?,0) as build_baddebt_sum,"
				 + " round(loan1_amt*0.01/?,0) as above_loan1_amt,"
				 + " round(loan2_amt*0.02/?,0) as above_loan2_amt,"
				 + " round(loan3_amt*0.5/?,0) as above_loan3_amt,"
				 + " round(loan4_amt/?,0) as above_loan4_amt,"
				 + " round(loan1_amt*0.01/?,0)+round(loan2_amt*0.02/?,0)+round(loan3_amt*0.5/?,0)+round(loan4_amt/?,0) as above_loan_sum,"
				 + " decode(baddebt_flag,'Y','1','0') as amt990036,"//--中央存保下載文字檔使用(990036)
				 + " decode(baddebt_flag,'N','1','0') as amt990037,"//--中央存保下載文字檔使用(990037)
				 + " round(baddebt_noenough/?,0) as baddebt_noenough,"
				 + " round(baddebt_104/?,0) as baddebt_104,"
				 + " round(baddebt_105/?,0) as baddebt_105,"
				 + " round(baddebt_106/?,0) as baddebt_106,"
				 + " round(baddebt_107/?,0) as baddebt_107,"
				 + " round(baddebt_108/?,0) as baddebt_108," 
				 + " round(property_loss/?,0) as property_loss," 
				 + " baddebt_flag,"//--Y: 是-帳列備呆等於或大於最低標準,N:否-小於
				 + " baddebt_delay "//--Y:無法一年內提足 N:可於1年內提足
				 + " from a10 left join (select * from bn01 where m_year=?  )bn01  on a10.bank_code = bn01.bank_no " 
				 + " where a10.m_year= ? "
				 + " and m_month= ? "
				 + " and bank_code= ? "); 
	      for(int k=0;k<45;k++){
	    	  paramList.add(unit) ;
	      }
	     paramList.add(u_year) ;
	     paramList.add(M_YEAR) ;
	     paramList.add(M_MONTH) ;
	     paramList.add(bank_code) ;
	     
	      dbData =  DBManager.QueryDB_SQLParam(sql.toString(), paramList, "m_year,m_month,loan1_amt,loan2_amt,loan3_amt,loan4_amt,loan_sum,invest1_amt,invest2_amt,invest3_amt,invest4_amt,invest_sum,other1_amt,other2_amt,other3_amt,other4_amt,other_sum,type1_sum,type2_sum,type3_sum,type4_sum,type_sum,"+
															    		  "loan1_baddebt,loan2_baddebt,loan3_baddebt,loan4_baddebt,loan_baddebt_sum,"+
															    		  "build1_baddebt,build2_baddebt,build3_baddebt,build4_baddebt,build_baddebt_sum,"+
															    		  "above_loan1_amt,above_loan2_amt,above_loan3_amt,above_loan4_amt,above_loan_sum,"+
															    		  "amt990036,amt990037,baddebt_noenough,baddebt_104,baddebt_105,baddebt_106,baddebt_107,baddebt_108,baddebt_flag,baddebt_delay,property_loss");
	      System.out.println("dbData.size=" + dbData.size());
	     
	      //設定報表表頭資料============================================
	      
	      row=sheet.getRow(1);
	      cell=row.getCell((short)1);
	      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	      cell.setCellValue("民國"+M_YEAR+"年"+M_MONTH+"月"+((dbData == null || dbData.size() ==0)?"無資料存在":""));  	
	      
	      row = sheet.getRow(2);
	      cell = row.getCell( (short) 1);
	      cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
	      if(dbData != null && dbData.size() != 0){
	      	 bean = (DataObject)dbData.get(0);
	      	 bank_name = (String)bean.getValue("bank_name");
	      }
	      cell.setCellValue("農漁會信用部名稱："+bank_name);	      
	      System.out.println("單位="+unit_name);
	      row = sheet.getRow(3);                        
  	   	  cell = row.getCell( (short)8);
  	   	  cell.setEncoding(HSSFCell.ENCODING_UTF_16);
  	   	  cell.setCellValue("單位:新臺幣 " + unit_name + ",%");
	      
	      if (dbData == null || dbData.size() ==0) {      	      
	      }else {
	      	rowNum = 6;
	      	bean = (DataObject)dbData.get(0);
	      	data_year=(bean.getValue("m_year") == null)?"":(bean.getValue("m_year")).toString();
	        data_month=(bean.getValue("m_month") == null)?"":(bean.getValue("m_month")).toString();
	        loan1_amt=(bean.getValue("loan1_amt") == null)?"0":(bean.getValue("loan1_amt")).toString();
	        loan2_amt=(bean.getValue("loan2_amt") == null)?"0":(bean.getValue("loan2_amt")).toString();//放款-列二類
			loan3_amt=(bean.getValue("loan3_amt") == null)?"0":(bean.getValue("loan3_amt")).toString();//放款-列三類
			loan4_amt=(bean.getValue("loan4_amt") == null)?"0":(bean.getValue("loan4_amt")).toString();//放款-列四類
			loan_sum=(bean.getValue("loan_sum") == null)?"0":(bean.getValue("loan_sum")).toString();//放款-合計
			invest1_amt=(bean.getValue("invest1_amt") == null)?"0":(bean.getValue("invest1_amt")).toString();
			invest2_amt=(bean.getValue("invest2_amt") == null)?"0":(bean.getValue("invest2_amt")).toString();//投資-列二類
			invest3_amt=(bean.getValue("invest3_amt") == null)?"0":(bean.getValue("invest3_amt")).toString();//投資-列三類		
			invest4_amt=(bean.getValue("invest4_amt") == null)?"0":(bean.getValue("invest4_amt")).toString();//投資-列四類  
			invest_sum=(bean.getValue("invest_sum") == null)?"0":(bean.getValue("invest_sum")).toString();//投資-合計
			other1_amt=(bean.getValue("other1_amt") == null)?"0":(bean.getValue("other1_amt")).toString();
			other2_amt=(bean.getValue("other2_amt") == null)?"0":(bean.getValue("other2_amt")).toString();//其他-列二類
			other3_amt=(bean.getValue("other3_amt") == null)?"0":(bean.getValue("other3_amt")).toString();//其他-列三類		
			other4_amt=(bean.getValue("other4_amt") == null)?"0":(bean.getValue("other4_amt")).toString();//其他-列四類 	
			other_sum=(bean.getValue("other_sum") == null)?"0":(bean.getValue("other_sum")).toString();//其他-合計
			type1_sum=(bean.getValue("type1_sum") == null)?"0":(bean.getValue("type1_sum")).toString();
			type2_sum=(bean.getValue("type2_sum") == null)?"0":(bean.getValue("type2_sum")).toString();//列二類-合計
			type3_sum=(bean.getValue("type3_sum") == null)?"0":(bean.getValue("type3_sum")).toString();//列三類-合計
			type4_sum=(bean.getValue("type4_sum") == null)?"0":(bean.getValue("type4_sum")).toString();//列四類-合計
			type_sum=(bean.getValue("type_sum") == null)?"0":(bean.getValue("type_sum")).toString();//總合計-合計
			loan1_baddebt=(bean.getValue("loan1_baddebt") == null)?"0":(bean.getValue("loan1_baddebt")).toString();
			loan2_baddebt=(bean.getValue("loan2_baddebt") == null)?"0":(bean.getValue("loan2_baddebt")).toString();
			loan3_baddebt=(bean.getValue("loan3_baddebt") == null)?"0":(bean.getValue("loan3_baddebt")).toString();
			loan4_baddebt=(bean.getValue("loan4_baddebt") == null)?"0":(bean.getValue("loan4_baddebt")).toString();
		  	loan_baddebt_sum=(bean.getValue("loan_baddebt_sum") == null)?"0":(bean.getValue("loan_baddebt_sum")).toString();
		  	build1_baddebt=(bean.getValue("build1_baddebt") == null)?"0":(bean.getValue("build1_baddebt")).toString();
		  	build2_baddebt=(bean.getValue("build2_baddebt") == null)?"0":(bean.getValue("build2_baddebt")).toString();
		  	build3_baddebt=(bean.getValue("build3_baddebt") == null)?"0":(bean.getValue("build3_baddebt")).toString();
		  	build4_baddebt=(bean.getValue("build4_baddebt") == null)?"0":(bean.getValue("build4_baddebt")).toString();
		  	build_baddebt_sum=(bean.getValue("build_baddebt_sum") == null)?"0":(bean.getValue("build_baddebt_sum")).toString();
		  	above_loan1_amt=(bean.getValue("above_loan1_amt") == null)?"0":(bean.getValue("above_loan1_amt")).toString();
		  	above_loan2_amt=(bean.getValue("above_loan2_amt") == null)?"0":(bean.getValue("above_loan2_amt")).toString();
		  	above_loan3_amt=(bean.getValue("above_loan3_amt") == null)?"0":(bean.getValue("above_loan3_amt")).toString();
		  	above_loan4_amt=(bean.getValue("above_loan4_amt") == null)?"0":(bean.getValue("above_loan4_amt")).toString();
		  	above_loan_sum=(bean.getValue("above_loan_sum") == null)?"0":(bean.getValue("above_loan_sum")).toString();
		  	amt990036=(bean.getValue("amt990036") == null)?"0":(bean.getValue("amt990036")).toString();
		  	amt990037=(bean.getValue("amt990037") == null)?"0":(bean.getValue("amt990037")).toString();
		  	baddebt_noenough  = (bean.getValue("baddebt_noenough") == null)?"0":(bean.getValue("baddebt_noenough")).toString();
		  	baddebt_104=(bean.getValue("baddebt_104") == null)?"":(bean.getValue("baddebt_104")).toString();
		  	baddebt_105=(bean.getValue("baddebt_105") == null)?"":(bean.getValue("baddebt_105")).toString();
		  	baddebt_106=(bean.getValue("baddebt_106") == null)?"":(bean.getValue("baddebt_106")).toString();
		  	baddebt_107=(bean.getValue("baddebt_107") == null)?"":(bean.getValue("baddebt_107")).toString();
		  	baddebt_108=(bean.getValue("baddebt_108") == null)?"":(bean.getValue("baddebt_108")).toString();
		  	baddebt_flag =(bean.getValue("baddebt_flag") == null)?"":(bean.getValue("baddebt_flag")).toString();
		  	baddebt_delay =(bean.getValue("baddebt_delay") == null)?"":(bean.getValue("baddebt_delay")).toString();
		  	property_loss =(bean.getValue("property_loss") == null)?"":(bean.getValue("property_loss")).toString();//106.05.04 add
			//應予評估資產-放款
			row = sheet.getRow(rowNum);
			for(int cellcount=5;cellcount<10;cellcount++){
		 	    cell = row.getCell( (short)cellcount);
		    	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		    	cell.setCellStyle(cs_right);
		    	if(cellcount == 5) cell.setCellValue(Utility.setCommaFormat(loan1_amt));
		    	if(cellcount == 6) cell.setCellValue(Utility.setCommaFormat(loan2_amt));	 
		    	if(cellcount == 7) cell.setCellValue(Utility.setCommaFormat(loan3_amt));
		    	if(cellcount == 8) cell.setCellValue(Utility.setCommaFormat(loan4_amt));
		    	if(cellcount == 9) cell.setCellValue(Utility.setCommaFormat(loan_sum));	     		
			}//end of cellcount	 		  	  	 	
			
			//應予評估資產-投資
			rowNum++;
			row = sheet.getRow(rowNum);
			for(int cellcount=5;cellcount<10;cellcount++){
		 	    cell = row.getCell( (short)cellcount);
		    	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		    	cell.setCellStyle(cs_right);
		    	if(cellcount == 5) cell.setCellValue(Utility.setCommaFormat(invest1_amt));
		    	if(cellcount == 6) cell.setCellValue(Utility.setCommaFormat(invest2_amt));	 
		    	if(cellcount == 7) cell.setCellValue(Utility.setCommaFormat(invest3_amt));
		    	if(cellcount == 8) cell.setCellValue(Utility.setCommaFormat(invest4_amt));
		    	if(cellcount == 9) cell.setCellValue(Utility.setCommaFormat(invest_sum));	     		
			}//end of cellcount
			//應予評估資產-其他
			rowNum++;
			row = sheet.getRow(rowNum);
			for(int cellcount=5;cellcount<10;cellcount++){
		 	    cell = row.getCell( (short)cellcount);
		    	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		    	cell.setCellStyle(cs_right);
		    	if(cellcount == 5) cell.setCellValue(Utility.setCommaFormat(other1_amt));
		    	if(cellcount == 6) cell.setCellValue(Utility.setCommaFormat(other2_amt));	 
		    	if(cellcount == 7) cell.setCellValue(Utility.setCommaFormat(other3_amt));
		    	if(cellcount == 8) cell.setCellValue(Utility.setCommaFormat(other4_amt));
		    	if(cellcount == 9) cell.setCellValue(Utility.setCommaFormat(other_sum));	     		
			}//end of cellcount
			//應予評估資產-合計
			rowNum++;
			row = sheet.getRow(rowNum);
			for(int cellcount=5;cellcount<10;cellcount++){
		 	    cell = row.getCell( (short)cellcount);
		    	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		    	cell.setCellStyle(cs_right);
		    	if(cellcount == 5) cell.setCellValue(Utility.setCommaFormat(type1_sum));
		    	if(cellcount == 6) cell.setCellValue(Utility.setCommaFormat(type2_sum));	 
		    	if(cellcount == 7) cell.setCellValue(Utility.setCommaFormat(type3_sum));
		    	if(cellcount == 8) cell.setCellValue(Utility.setCommaFormat(type4_sum));
		    	if(cellcount == 9) cell.setCellValue(Utility.setCommaFormat(type_sum));	     		
			}//end of cellcount
			//帳列備抵呆帳-放款
			rowNum+=4;
			row = sheet.getRow(rowNum);
			for(int cellcount=5;cellcount<10;cellcount++){
		 	    cell = row.getCell( (short)cellcount);
		    	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		    	cell.setCellStyle(cs_right);
		    	if(cellcount == 5) cell.setCellValue(Utility.setCommaFormat(loan1_baddebt));
		    	if(cellcount == 6) cell.setCellValue(Utility.setCommaFormat(loan2_baddebt));	 
		    	if(cellcount == 7) cell.setCellValue(Utility.setCommaFormat(loan3_baddebt));
		    	if(cellcount == 8) cell.setCellValue(Utility.setCommaFormat(loan4_baddebt));
		    	if(cellcount == 9) cell.setCellValue(Utility.setCommaFormat(loan_baddebt_sum));	     		
			}//end of cellcount
			//帳列備抵呆帳-放款
			rowNum++;
			row = sheet.getRow(rowNum);
			for(int cellcount=5;cellcount<10;cellcount++){
		 	    cell = row.getCell( (short)cellcount);
		    	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		    	cell.setCellStyle(cs_right);
		    	if(cellcount == 5) cell.setCellValue(Utility.setCommaFormat(build1_baddebt));
		    	if(cellcount == 6) cell.setCellValue(Utility.setCommaFormat(build2_baddebt));	 
		    	if(cellcount == 7) cell.setCellValue(Utility.setCommaFormat(build3_baddebt));
		    	if(cellcount == 8) cell.setCellValue(Utility.setCommaFormat(build4_baddebt));
		    	if(cellcount == 9) cell.setCellValue(Utility.setCommaFormat(build_baddebt_sum));	     		
			}//end of cellcount
			//依規定應提列最低標準之備抵呆帳-放款
			rowNum+=4;
			row = sheet.getRow(rowNum);
			for(int cellcount=5;cellcount<10;cellcount++){
		 	    cell = row.getCell( (short)cellcount);
		    	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		    	cell.setCellStyle(cs_right);
		    	if(cellcount == 5) cell.setCellValue(Utility.setCommaFormat(above_loan1_amt));
		    	if(cellcount == 6) cell.setCellValue(Utility.setCommaFormat(above_loan2_amt));	 
		    	if(cellcount == 7) cell.setCellValue(Utility.setCommaFormat(above_loan3_amt));
		    	if(cellcount == 8) cell.setCellValue(Utility.setCommaFormat(above_loan4_amt));
		    	if(cellcount == 9) cell.setCellValue(Utility.setCommaFormat(above_loan_sum));	     		
			}//end of cellcount
			
			//非授信資產可能遭受損失(投資及其他資產)
			rowNum++;
			row = sheet.getRow(rowNum);
			cell = row.getCell( (short)9);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellStyle(cs_right);
            cell.setCellValue(Utility.setCommaFormat(property_loss));//106.05.04 add    
            
			//備抵呆帳是否依規定提足
			rowNum+=3;
			if(!"".equals(baddebt_flag)){
				if("N".equals(baddebt_flag)){
					rowNum++;
				}
				row = sheet.getRow(rowNum);
				cell = row.getCell( (short)1);
		    	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellValue("■");
				if("N".equals(baddebt_flag)){
					//備抵呆帳提列不足額部份
					rowNum++;
					row = sheet.getRow(rowNum);
					cell = row.getCell( (short)3);
			    	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellValue("備抵呆帳提列不足額部分"+Utility.setCommaFormat(baddebt_noenough));
					rowNum+=2;
					if(!"".equals(baddebt_delay)){//Y:無法一年內提足 N:可於1年內提足
						if("N".equals(baddebt_delay)){
							row = sheet.getRow(rowNum);
							cell = row.getCell( (short)3);
					    	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					    	cell.setCellValue("■");
							cell = row.getCell( (short)4);
					    	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
							cell.setCellValue("可於1年內提足(即於104年底前提足)，預計提撥金額"+Utility.setCommaFormat(baddebt_104));
						}else if("Y".equals(baddebt_delay)){
							rowNum++;
							row = sheet.getRow(rowNum);
							cell = row.getCell( (short)3);
					    	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
							cell.setCellValue("■");
							rowNum++;
							if(!"".equals(baddebt_105)){
								row = sheet.getRow(rowNum);
								cell = row.getCell( (short)4);
						    	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
								cell.setCellValue("■");
								cell = row.getCell( (short)6);
						    	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
								cell.setCellValue(Utility.setCommaFormat(baddebt_105));
							}
							rowNum++;
							if(!"".equals(baddebt_106)){
								row = sheet.getRow(rowNum);
								cell = row.getCell( (short)4);
						    	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
								cell.setCellValue("■");
								cell = row.getCell( (short)6);
						    	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
								cell.setCellValue(Utility.setCommaFormat(baddebt_106));
							}
							rowNum++;
							if(!"".equals(baddebt_107)){
								row = sheet.getRow(rowNum);
								cell = row.getCell( (short)4);
						    	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
								cell.setCellValue("■");
								cell = row.getCell( (short)6);
						    	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
								cell.setCellValue(Utility.setCommaFormat(baddebt_107));
							}
							rowNum++;
							if(!"".equals(baddebt_108)){
								row = sheet.getRow(rowNum);
								cell = row.getCell( (short)4);
						    	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
								cell.setCellValue("■");
								cell = row.getCell( (short)6);
						    	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
								cell.setCellValue(Utility.setCommaFormat(baddebt_108));
							}
						}
					}
				}
			}
			
	      } //end of else dbData.size() != 0
	      
	      
	      FileOutputStream fout = null;     
	      fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + "應予評估資產彙總表.xls");
	     
	      HSSFFooter footer = sheet.getFooter();
	      footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
	      footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
	      wb.write(fout);
	      //儲存
	      fout.close();
	    }
	    catch (Exception e) {
	    	System.out.println("RptFR047W.createRpt Error:" + e + e.getMessage());
	    }
	    
	    return errMsg;
	  }
	 
	 
	  //96.04.24 fix A10中央存保格式檔案下載改用回傳List 
	  public static List printData(String M_YEAR, String M_MONTH, String bank_code, String bank_type,String unit) {	                               
	  	String errMsg = "";
	  	List dbData = null;
	  	String sqlCmd = "";
	  		
	  	HashMap h = new HashMap();//儲存data
	  	List bank_code_list = new LinkedList();		
	  	HashMap bank_code_h = new HashMap();//儲存每個bank_code的data		  	
	  	List return_List = new LinkedList();
	  	StringBuffer sql = new StringBuffer() ;
	  	ArrayList paramList = new ArrayList() ;
	  	String u_year = "100" ;
	  	if(M_YEAR==null || Integer.parseInt(M_YEAR)<=99) {
	  		u_year = "99" ;
 	  	}
	  	System.out.println("printData.M_MONTH="+M_MONTH);
	  	String yymm = M_YEAR+M_MONTH;
	  	try {
	  		System.out.println("******yymm="+Integer.parseInt(yymm));
	  		sql.append(" select a10.m_year,m_month,bank_code,"
		    		 + " round(loan1_amt/'1',0) as loan1_amt,"
		      		 + " round(loan2_amt/'1',0) as loan2_amt,"
					 + " round(loan3_amt/'1',0) as loan3_amt,"
					 + " round(loan4_amt/'1',0) as loan4_amt,"
					 + " round((loan1_amt+loan2_amt+loan3_amt+loan4_amt)/'1',0) as loan_sum,"
					 + " round(invest1_amt/'1',0) as invest1_amt,"
		      	     + " round(invest2_amt/'1',0) as invest2_amt,"
					 + " round(invest3_amt/'1',0) as invest3_amt,"
					 + " round(invest4_amt/'1',0) as invest4_amt,"
					 + " round((invest1_amt+invest2_amt+invest3_amt+invest4_amt)/'1',0) as invest_sum,"
					 + " round(other1_amt/'1',0) as other1_amt,"
					 + " round(other2_amt/'1',0) as other2_amt,"
					 + " round(other3_amt/'1',0) as other3_amt,"
					 + " round(other4_amt/'1',0) as other4_amt,"
					 + " round((other1_amt+other2_amt+other3_amt+other4_amt)/'1',0) as other_sum, "
					 + " round((loan1_amt+invest1_amt+other1_amt)/'1',0) as type1_sum,"
					 + " round((loan2_amt+invest2_amt+other2_amt)/'1',0) as type2_sum,"
					 + " round((loan3_amt+invest3_amt+other3_amt)/'1',0) as type3_sum,"
					 + " round((loan4_amt+invest4_amt+other4_amt)/'1',0) as type4_sum,"
					 + " round((loan1_amt+invest1_amt+other1_amt+loan2_amt+invest2_amt+other2_amt+loan3_amt+invest3_amt+other3_amt+loan4_amt+invest4_amt+other4_amt)/'1',0) as type_sum,"
					 + " round(loan1_baddebt/'1',0) as loan1_baddebt,"
					 + " round(loan2_baddebt/'1',0) as loan2_baddebt,"
					 + " round(loan3_baddebt/'1',0) as loan3_baddebt,"
					 + " round(loan4_baddebt/'1',0) as loan4_baddebt,"
					 + " round((loan1_baddebt+loan2_baddebt+loan3_baddebt+loan4_baddebt)/'1',0) as loan_baddebt_sum,"
					 + " round(build1_baddebt/'1',0) as build1_baddebt,"
					 + " round(build2_baddebt/'1',0) as build2_baddebt,"
					 + " round(build3_baddebt/'1',0) as build3_baddebt,"
					 + " round(build4_baddebt/'1',0) as build4_baddebt,"
					 + " round((build1_baddebt+build2_baddebt+build3_baddebt+build4_baddebt)/'1',0) as build_baddebt_sum,"
					 + " round(loan1_amt*0.01/'1',0) as above_loan1_amt,"
					 + " round(loan2_amt*0.02/'1',0) as above_loan2_amt,"
					 + " round(loan3_amt*0.5/'1',0) as above_loan3_amt,"
					 + " round(loan4_amt/'1',0) as above_loan4_amt,"
					 + " round(loan1_amt*0.01/'1',0)+round(loan2_amt*0.02/'1',0)+round(loan3_amt*0.5/'1',0)+round(loan4_amt/'1',0) as above_loan_sum,"
					 + " decode(baddebt_flag,'Y','1','0') as amt990036,"//--中央存保下載文字檔使用(990036)
					 + " decode(baddebt_flag,'N','1','0') as amt990037,"//--中央存保下載文字檔使用(990037)
					 + " round(baddebt_noenough/'1',0) as baddebt_noenough,"
					 + " round(baddebt_104/'1',0) as baddebt_104,"
					 + " round(baddebt_105/'1',0) as baddebt_105,"
					 + " round(baddebt_106/'1',0) as baddebt_106,"
					 + " round(baddebt_107/'1',0) as baddebt_107,"
					 + " round(baddebt_108/'1',0) as baddebt_108," 
					 + " round(property_loss/'1',0) as property_loss," 
					 + " baddebt_flag,"//--Y: 是-帳列備呆等於或大於最低標準,N:否-小於
					 + " baddebt_delay "//--Y:無法一年內提足 N:可於1年內提足
					 + " from a10 left join (select * from bn01 where m_year=?  )bn01  on a10.bank_code = bn01.bank_no " 
					 + " where a10.m_year= ? "
					 + " and m_month= ? "
					 + " and "+bank_code);  
		     paramList.add(u_year) ;
		     paramList.add(M_YEAR) ;
		     paramList.add(M_MONTH) ;
		     dbData =  DBManager.QueryDB_SQLParam(sql.toString(), paramList, "m_year,m_month,loan1_amt,loan2_amt,loan3_amt,loan4_amt,loan_sum,invest1_amt,invest2_amt,invest3_amt,invest4_amt,invest_sum,other1_amt,other2_amt,other3_amt,other4_amt,other_sum,type1_sum,type2_sum,type3_sum,type4_sum,type_sum,"+
		    		  "loan1_baddebt,loan2_baddebt,loan3_baddebt,loan4_baddebt,loan_baddebt_sum,"+
		    		  "build1_baddebt,build2_baddebt,build3_baddebt,build4_baddebt,build_baddebt_sum,"+
		    		  "above_loan1_amt,above_loan2_amt,above_loan3_amt,above_loan4_amt,above_loan_sum,"+
		    		  "amt990036,amt990037,baddebt_noenough,baddebt_104,baddebt_105,baddebt_106,baddebt_107,baddebt_108,baddebt_flag,baddebt_delay,property_loss");
	  		if(Integer.parseInt(yymm)>=10403){
			     print_table = print_table2;
	  		}else{
	  			print_table = print_table1;
	  			/*sql.append(" select a10.m_year,m_month,bank_code,"
		        		+ " round(loan2_amt/1,0) as loan2_amt,"
		  			 	+ " round(loan3_amt/1,0) as loan3_amt,"
		  			 	+ " round(loan4_amt/1,0) as loan4_amt,"
		  			 	+ " round((loan2_amt+loan3_amt+loan4_amt)/1,0) as loan_sum,"
		        	    + " round(invest2_amt/1,0) as invest2_amt,"
		  			    + " round(invest3_amt/1,0) as invest3_amt,"
		  			    + " round(invest4_amt/1,0) as invest4_amt,"
		  			    + " round((invest2_amt+invest3_amt+invest4_amt)/1,0) as invest_sum,"
		  			    + " round(other2_amt/1,0) as other2_amt,"
		  			    + " round(other3_amt/1,0) as other3_amt,"
		  			    + " round(other4_amt/1,0) as other4_amt,"
		  			    + " round((other2_amt+other3_amt+other4_amt)/1,0) as other_sum, "
		  			    + " round((loan2_amt+invest2_amt+other2_amt)/1,0) as type2_sum,"
		  			    + " round((loan3_amt+invest3_amt+other3_amt)/1,0) as type3_sum,"
		  			    + " round((loan4_amt+invest4_amt+other4_amt)/1,0) as type4_sum,"
		  			    + " round((loan2_amt+invest2_amt+other2_amt+loan3_amt+invest3_amt+other3_amt+loan4_amt+invest4_amt+other4_amt)/1,0) as type_sum"
		  			    + " from a10 left join (select * from bn01 where m_year= ?) bn01  on a10.bank_code = bn01.bank_no " 
		  			    + " where a10.m_year= ? "
		  			    + " and m_month= ? ");
	  			paramList.add(u_year) ;
			    paramList.add(M_YEAR) ;
			    paramList.add(M_MONTH) ;
		  		if(bank_code.length() == 7){
		  			sql.append("  and bank_code =  ? ") ;
		  			paramList.add(bank_code) ;
				}else{
					sql.append("and a10.").append(bank_code) ;
				}
		  		sql.append(" order by a10.bank_code ");*/
	  		}
	  		
		    sql.setLength(0) ;
	  		paramList.clear();
	  		System.out.println("dbData.size()="+dbData.size());
	  			// 取出資料存入MAP
	  			for (int k = 0; k < dbData.size(); k++) {				
		  			DataObject obj = (DataObject) dbData.get(k);
		  			h = new HashMap();//儲存data
		  			for (int i = 0; i < print_table.length; i++) {
		  				if(obj.getValue(print_table[i][0]) != null){//990001/amt	
		  				   h.put(print_table[i][1], obj.getValue(print_table[i][0]).toString());
		  				   //System.out.println(print_table[i][0]+"="+obj.getValue(print_table[i][0]).toString());
		  				}
		  			}		
		  			bank_code_list.add((String)obj.getValue("bank_code"));
		  			bank_code_h.put((String)obj.getValue("bank_code"),h);
		  		}
		  		
		  		System.out.println("bank_code_h.size()="+bank_code_h.size());
		  		System.out.println("bank_code_list.size()="+bank_code_list.size());
		  		for(int j = 0;j<bank_code_list.size();j++){
		  			for(int i = 0; i < print_table.length; i++) {
		  				h = (HashMap)bank_code_h.get((String)bank_code_list.get(j));	
		  				if(h.get(print_table[i][1]) != null){
		  						return_List.add(DownLoad.fillStuff(M_YEAR, "L", "0", 3)//年			   	   			   	   
			                                   +DownLoad.fillStuff(M_MONTH, "L", "0", 2)//月
			                                   +DownLoad.fillStuff((String)bank_code_list.get(j), "R", "0", 7)//機構代號
			  								   +print_table[i][1]//科目編號
			  						           +DownLoad.fillStuff((String)h.get(print_table[i][1]), "L", "0", 0, 14)//金額
		  								   ); 
		  				
		  										
		  			    } 
		  			}
		  		} 
	  		
	  	} catch (Exception e) {
	  		System.out.println("printData Error:" + e + e.getMessage());
	  	}
	  	System.out.println("return_List="+return_List.size());
	  	return return_List;
	  }
	 
}
