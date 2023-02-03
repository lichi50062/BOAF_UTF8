/*
  95.11.27 add 法定比率分析統計表(中央存保用) by 2295
  96.03.09 fix 990140存款總額(含公庫存庫).因含公庫存款.需扣除990150公庫存款 by 2295
  96.04.14 add 取得最近一年平均存放比率,列印月份往前推11個月
           add 取得最近一季存放比率,列印月份往前推2個月 by 2295
  96.04.17 add A99_extra中央存保格式A02.A01.A99檔案下載 by 2295
  96.04.24 fix A99_extra中央存保格式A02.A01.A99檔案下載改用回傳List by 2295
  96.04.26 fix A99_extra中央存保格式A02.A01.A99檔案下載科目代號改為6碼 by 2295
  96.07.27 fix (13)/(14)/(15)加印99141Y最近決算年度 by 2295
  96.08.30 fix 992170.(35)以轄區外縣市土地建物為擔(副)保品之戶數.直接顯示(不除以金額單位) by 2295
  96.10.30 fix 取得放款餘額(季平均),列印月份往前推2個月120000+120800+150300 
  			  加總資產負債表中的放款總額(120000+120800+150300),再除以3 by 2295
  97.01.29 add 990612非會員政策性農業專案貸款 by 2295 
  97.02.20 add 992550無擔保消費性貸款中之逾期放款
               992650無擔保消費性貸款中之應予觀察放款 by 2295
  97.04.17 add 980000中央存保文字檔.編號固定.位置變動.編碼不變
  97.04.18 990620帳面金額已內含公庫存款.不用再加990630公庫存款 by 2295
  98.03.25 fix tmpAmt = new BigDecimal((double)tmpAmt.multiply(tmp05).intValue()) by 2295
  98.07.27 add 990511非會員無擔保消費性政策貸款    
  			  990512非會員無擔保消費性非政策貸款 by 2295   			  
  98.12.11 add 992810投資全國農業金庫尚未攤提之損失 by 2295
 100.02.22 fix 根據查詢年度.100年以後取得新縣市別.100年以前取得舊縣市別 
 			   使用PreparedStatement;並列印轉換後的SQL;套用QueryDB_SQLParam by 2295
 100.04.21 fix 查詢年月條件 by 2295 		
 102.02.05 add a02.amt_name by 2295	 
 102.11.19 add 原QueryDB改套用QueryDB_SQLParam by 2295    
 102.12.25 fix 調整printData SQL by 2295  
 104.04.24 fix 調整新DB.SQL無法適用問題 by 2295
 108.09.17 add 108年10月以後增加列印(58)購置住宅放款餘額(990711)、(59)房屋修繕放款餘額(990712) by 2295
 111.06.07 fix 調整111年9月以後,(32)本月份持有非由政府發行之債券及票卷餘額改至A02.990860申報 by 2295
 */
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*;
import java.io.*;
import java.math.BigDecimal;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;
import com.tradevan.util.DownLoad;

public class RptFR0066WB {
	static String[] table = null;
	final private static String[] table_10809 = {"990110_12", "990120_12", "990130_12", "990140_12", "990150_12",
			                               "990110_3", "990120_3", "990130_3", "990140_3", "990150_3",
			                               "990210", "990220", "990230", "990230", "990320", "992100", "992110", "992120", "992130", "990420",
										   "990620", "992140", "990410", "990610", "990630", "991020", "990810", "990910", "990920", "990710",
										   "120000", "992440", "992170", "992180", "992190", "992200", "992300", "992510", "992520", "992530",
										   "992540", "992610", "992620", "992630", "992640", "990611", "992710", "992720", "992730", "990510",
										   "990511", "990512", "992150", "992550", "992650", "990612", "992810"
										   };	
	final private static String[] table_10810 = {"990110_12", "990120_12", "990130_12", "990140_12", "990150_12",
        							             "990110_3", "990120_3", "990130_3", "990140_3", "990150_3",
        							             "990210", "990220", "990230", "990230", "990320", "992100", "992110", "992120", "992130", "990420",
        							             "990620", "992140", "990410", "990610", "990630", "991020", "990810", "990910", "990920", "990710",
        							             "120000", "992440", "992170", "992180", "992190", "992200", "992300", "992510", "992520", "992530",
        							             "992540", "992610", "992620", "992630", "992640", "990611", "992710", "992720", "992730", "990510",
        							             "990511", "990512", "992150", "992550", "992650", "990612", "992810", "990711", "990712"
		   								   };	
	static String[] print_table = null;
	final private static String[] print_table_10809 = {"990110_12", "990120_12", "990130_12", "990140_12", "990150_12",
            									 "990110_3", "990120_3", "990130_3", "990140_3", "990150_3",
												 "990210", "990220", "99141y","990230_990240","990230", "990320",
												 "992100", "992110", "992120", "992130", "990420","990620", "992140", 
												 "990410", "990610", "990630", "991020", "990810", "990910", "990920", "990710",
												 "120000", "992440", "992170", "992180", "992190", "992200", "992300", "992510", 
												 "992520", "992530", "992540", "992610", "992620", "992630", "992640", "990611", 
												 "992710", "992720", "992730", "990510", "990511", "990512", "992150", "992550", 
												 "992650", "990612", "992810"
			   									};	
	final private static String[] print_table_10810 = {"990110_12", "990120_12", "990130_12", "990140_12", "990150_12",
		 											   "990110_3", "990120_3", "990130_3", "990140_3", "990150_3",
		 											   "990210", "990220", "99141y","990230_990240","990230", "990320",
		 											   "992100", "992110", "992120", "992130", "990420","990620", "992140", 
		 											   "990410", "990610", "990630", "991020", "990810", "990910", "990920", "990710",
		 											   "120000", "992440", "992170", "992180", "992190", "992200", "992300", "992510", 
		 											   "992520", "992530", "992540", "992610", "992620", "992630", "992640", "990611", 
		 											   "992710", "992720", "992730", "990510", "990511", "990512", "992150", "992550", 
		 											   "992650", "990612", "992810", "990711", "990712"
													  };	
	public static String createRpt(String S_YEAR, String S_MONTH, String bank_code, String bank_type,String BANK_NAME,String unit) {	                               
		String errMsg = "";
		List dbData = null;
		String sqlCmd = "";
		String ncacno="";
		String unit_name="";
		BigDecimal tmpZero=new BigDecimal("0");		
		HashMap h = new HashMap();//儲存data
		String cd01_table = "";
        String wlx01_m_year = "";
		List paramList = new ArrayList();
		List a02_table_paramList = new ArrayList(); 
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
			System.out.println("RptFR0066WB.bank_type="+bank_type);
			FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator") + (bank_type.equals("6")?"農會":"漁會")+"信用部法定比率分析統計表"+(((Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH)) >= 10810)?"_10810":"")+".xls");
			ncacno = bank_type.equals("6")?"ncacno":"ncacno_7";	
			// 設定FileINputStream讀取Excel檔
			POIFSFileSystem fs = new POIFSFileSystem(finput);
			HSSFWorkbook wb = new HSSFWorkbook(fs);
			HSSFSheet sheet = wb.getSheetAt(0);// 讀取第一個工作表，宣告其為sheet
			HSSFPrintSetup ps = sheet.getPrintSetup(); // 取得設定
			// sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
			// sheet.setAutobreaks(true); //自動分頁

			// 設定頁面符合列印大小
			sheet.setAutobreaks(false);
			HSSFFooter footer = sheet.getFooter();
			ps.setScale((short) 88); // 列印縮放百分比

			ps.setPaperSize((short) 9); // 設定紙張大小 A4			
			finput.close();

			HSSFRow row = null;// 宣告一列
			HSSFCell cell = null;// 宣告一個儲存格

			short rowNo = 0;
			// short y = 0;
			sqlCmd = "select bank_name from bn01 where m_year=? and bank_no=?";
			paramList.add(wlx01_m_year);
			paramList.add(bank_code);
			
			dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");
			if(dbData != null && dbData.size() !=0){
			    BANK_NAME = (String)((DataObject) dbData.get(0)).getValue("bank_name");
			}
			//96.04.14 add 取得最近一年平均存放比率,列印月份往前推11個月
			/*select bank_code
			,round(sum(decode(acc_code,'990110',amt,0))/12,0) as amt990110
			,round(sum(decode(acc_code,'990120',amt,0))/12,0) as amt990120
			,round(sum(decode(acc_code,'990130',amt,0))/12,0) as amt990130
			,round((sum(decode(acc_code,'990140',amt,0))  - sum(decode(acc_code,'990150',amt,0)))/12,0)  as amt990140
			,round((sum(decode(acc_code,'990150',amt,0))/2)/12,0) as amt990150
			from a02
			where to_char(m_year * 100 + m_month) >= (select TO_CHAR(ADD_MONTHS(TO_DATE(200704,'YYYYMM'),-11),'YYYYMM') - 191100 from dual)
			  and to_char(m_year * 100 + m_month) <= 9604
			group by bank_code  
			*/
			paramList.clear();
			sqlCmd = "  select bank_code "
				   + " ,round(sum(decode(acc_code,'990110',amt,0))/12/?,0) as amt990110 "
				   + " ,round(sum(decode(acc_code,'990120',amt,0))/12/?,0) as amt990120 "
				   + " ,round(sum(decode(acc_code,'990130',amt,0))/12/?,0) as amt990130 "
				   + " ,round((sum(decode(acc_code,'990140',amt,0))  - sum(decode(acc_code,'990150',amt,0)))/12/?,0)  as amt990140 "
				   + " ,round((sum(decode(acc_code,'990150',amt,0))/2)/12/?,0) as amt990150 "
				   + " from a02 "
				   + " where to_char(m_year * 100 + m_month) >= (select TO_CHAR(ADD_MONTHS(TO_DATE(?,'YYYY/MM/DD'),-11),'YYYYMM') - 191100 from dual) "
				   + " and to_char(m_year * 100 + m_month) <= to_number(?)"
				   + " and bank_code = ?"
				   + " group by bank_code ";
			paramList.add(unit);
			paramList.add(unit);
			paramList.add(unit);
			paramList.add(unit);
			paramList.add(unit);
			paramList.add(Utility.getFullDate(S_YEAR,S_MONTH,"01"));
			paramList.add(S_YEAR+S_MONTH);
			paramList.add(bank_code);
			dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList, "amt990110,amt990120,amt990130,amt990140,amt990150");
			// 取出資料存入MAP
			System.out.println("dbData.size()="+dbData.size());
			
			for (int k = 0; k < dbData.size(); k++) {
				DataObject obj = (DataObject) dbData.get(k);
				if(obj.getValue("amt990110") != null){
				   h.put("990110_12", obj.getValue("amt990110").toString());				
				}
				if(obj.getValue("amt990120") != null){
					   h.put("990120_12", obj.getValue("amt990120").toString());				
				}
				if(obj.getValue("amt990130") != null){
					   h.put("990130_12", obj.getValue("amt990130").toString());				
				}
				if(obj.getValue("amt990140") != null){
					   h.put("990140_12", obj.getValue("amt990140").toString());				
				}
				if(obj.getValue("amt990150") != null){
					   h.put("990150_12", obj.getValue("amt990150").toString());				
				}
			}
	        //96.04.14 add 取得最近一季存放比率,列印月份往前推2個月
			/*
			select bank_code
			,round(sum(decode(acc_code,'990110',amt,0))/3,0) as amt990110
			,round(sum(decode(acc_code,'990120',amt,0))/3,0) as amt990120
			,round(sum(decode(acc_code,'990130',amt,0))/3,0) as amt990130
			,round((sum(decode(acc_code,'990140',amt,0))  - sum(decode(acc_code,'990150',amt,0)))/3,0)  as amt990140
			,round((sum(decode(acc_code,'990150',amt,0))/2)/3,0) as amt990150
			from a02
			where to_char(m_year * 100 + m_month) >= (select TO_CHAR(ADD_MONTHS(TO_DATE(200704,'YYYYMM'),-2),'YYYYMM') - 191100 from dual)
			  and to_char(m_year * 100 + m_month) <= 9604
			group by bank_code  
			*/  
			paramList.clear();
			sqlCmd = "  select bank_code "
				   + " ,round(sum(decode(acc_code,'990110',amt,0))/3/?,0) as amt990110 "
				   + " ,round(sum(decode(acc_code,'990120',amt,0))/3/?,0) as amt990120 "
				   + " ,round(sum(decode(acc_code,'990130',amt,0))/3/?,0) as amt990130 "
				   + " ,round((sum(decode(acc_code,'990140',amt,0))  - sum(decode(acc_code,'990150',amt,0)))/3/?,0)  as amt990140 "
				   + " ,round((sum(decode(acc_code,'990150',amt,0))/2)/3/?,0) as amt990150 "
				   + " from a02 "
				   + " where to_char(m_year * 100 + m_month) >= (select TO_CHAR(ADD_MONTHS(TO_DATE(?,'YYYY/MM/DD'),-2),'YYYYMM') - 191100 from dual) "
				   + " and to_char(m_year * 100 + m_month) <= to_number(?)"
				   + " and bank_code = ?"
				   + " group by bank_code ";
			paramList.add(unit);
			paramList.add(unit);
			paramList.add(unit);
			paramList.add(unit);
			paramList.add(unit);
			paramList.add(Utility.getFullDate(S_YEAR,S_MONTH,"01"));
			paramList.add(S_YEAR+S_MONTH);
			paramList.add(bank_code);
			
			dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList, "amt990110,amt990120,amt990130,amt990140,amt990150");
			// 取出資料存入MAP
			System.out.println("dbData.size()="+dbData.size());
			
			for (int k = 0; k < dbData.size(); k++) {
				DataObject obj = (DataObject) dbData.get(k);
				if(obj.getValue("amt990110") != null){
				   h.put("990110_3", obj.getValue("amt990110").toString());				
				}
				if(obj.getValue("amt990120") != null){
					   h.put("990120_3", obj.getValue("amt990120").toString());				
				}
				if(obj.getValue("amt990130") != null){
					   h.put("990130_3", obj.getValue("amt990130").toString());				
				}
				if(obj.getValue("amt990140") != null){
					   h.put("990140_3", obj.getValue("amt990140").toString());				
				}
				if(obj.getValue("amt990150") != null){
					   h.put("990150_3", obj.getValue("amt990150").toString());				
				}
			}
			
			/*96.10.30 add 取得放款餘額(季平均),列印月份往前推2個月120000+120800+150300
			select bank_code,round(sum(decode(acc_code,'120000',amt,'120800',amt,'150300',amt,0))/3,0) as amt120000
			from a01
			where to_char(m_year * 100 + m_month) >= (select TO_CHAR(ADD_MONTHS(TO_DATE(200709,'YYYYMM'),-2),'YYYYMM') - 191100 from dual)
			  and to_char(m_year * 100 + m_month) <= 9609
			  and bank_code = '6050018'
			group by bank_code  
			*/		
			paramList.clear();
			sqlCmd = "  select bank_code, "
				   + "  round(sum(decode(acc_code,'120000',amt,'120800',amt,'150300',amt,0))/3/?,0) as amt120000 "
				   + " from a01 "
				   + " where to_char(m_year * 100 + m_month) >= (select TO_CHAR(ADD_MONTHS(TO_DATE(?,'YYYY/MM/DD'),-2),'YYYYMM') - 191100 from dual) "
				   + " and to_char(m_year * 100 + m_month) <= to_number(?)"
				   + " and bank_code = ?"
				   + " group by bank_code ";
			paramList.add(unit);
			paramList.add(Utility.getFullDate(S_YEAR,S_MONTH,"01"));
			paramList.add(S_YEAR+S_MONTH);
			paramList.add(bank_code);
			dbData = DBManager.QueryDB_SQLParam(sqlCmd, paramList,"amt120000");
			// 取出資料存入MAP
			System.out.println("dbData.size()="+dbData.size());
			//(33)放款餘額(季平均)
			for (int k = 0; k < dbData.size(); k++) {
				DataObject obj = (DataObject) dbData.get(k);
				if(obj.getValue("amt120000") != null){
				   h.put("120000", obj.getValue("amt120000").toString());				
				}				
			}			
			
			//98.08.30 fix 992170.(35)以轄區外縣市土地建物為擔(副)保品之戶數.直接顯示.
			paramList.clear();
			sqlCmd = " select m_year,m_month,"+ncacno+".acc_code,"+ncacno+".acc_name,decode("+ncacno+".acc_code,'992190',amt,'992200',amt,'992300',amt,'99141Y',amt,'992170',amt,round(nvl(a02.amt,0)/?,0)) as amt "
		       + " from "+ncacno+" left join ("
			   + " 						  select m_year,m_month,bank_code,acc_code,amt from a02 " 
			   + " 					      where m_year = ? and m_month = ?"
			   + " 						  and   bank_code = ?"
			   + " 						  union " 
			   + " 						  select * from a99 " 
			   + " 						  where m_year = ? and m_month = ?" 
			   + " 						  and   bank_code = ?"			  
			   + " 						)a02 "
			   + "						on "+ncacno+".acc_code=a02.acc_code "
			   + " where (acc_tr_type='A02' or acc_tr_type='A99')"
			   + " and m_year is not null and m_month is not null"
			   + " order by "+ncacno+".acc_tr_type,"+ncacno+".acc_range";
			paramList.add(unit);
			paramList.add(S_YEAR);
			paramList.add(S_MONTH);
			paramList.add(bank_code);
			paramList.add(S_YEAR);
			paramList.add(S_MONTH);
			paramList.add(bank_code);
			
			dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList, "m_year,m_month,amt");
			// 取出資料存入MAP
			System.out.println("dbData.size()="+dbData.size());
			//HashMap h = new HashMap();
			for (int k = 0; k < dbData.size(); k++) {
				DataObject obj = (DataObject) dbData.get(k);
				if(obj.getValue("amt") != null){
				   h.put(obj.getValue("acc_code"), obj.getValue("amt").toString());				
				}
			}
			
			rowNo = 0;			
			row = sheet.getRow(rowNo);
			cell = row.getCell((short) 0);
			// 設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue(BANK_NAME );
			String lastday = Utility.getCHTdate(Utility.getLastDay(String.valueOf(Integer.parseInt(S_YEAR)+1911)+"/"+S_MONTH,"yyyy/mm"),1);
			
            rowNo=2;
			if (h.isEmpty()) {
				row = sheet.getRow(rowNo);
				cell = row.getCell((short) 0);
				// 設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);			
				cell.setCellValue("基準日:"+S_YEAR + "年" + S_MONTH + "月"+lastday.substring(0,10).substring(8,10)+"日無資料存在                          ");
			} else {			    
				row = sheet.getRow(rowNo);
				cell = row.getCell((short) 0);				
				if (unit.equals("1")){
   	 	   	 		unit_name="元";
   	 	   	 	}else if (unit.equals("1000")){   	 	   		
   	 	   	 	   	unit_name="仟元";
   	       	    }else if (unit.equals("10000")){
   	       	       	unit_name="萬元";
   	 	   	 	}else if (unit.equals("1000000")){
   	 	   	 	   	unit_name="百萬元";
   	 	   	 	}else if (unit.equals("10000000")){
   	 	   	 	   	unit_name="千萬元";
   	       	    }else if (unit.equals("100000000")){
   	       	       	unit_name="億元";
   	 	   	 	}
   	       	    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
   	       	    cell.setCellValue("基準日:"+S_YEAR + "年" + S_MONTH + "月"+lastday.substring(0,10).substring(8,10)+"日              單位：新台幣"+unit_name+"、％");   	       	    
				rowNo = 4;
				//double tmpAmt = 0.0;	
				BigDecimal tmp05=new BigDecimal("0.5");
				BigDecimal tmpAmt = null;
				String str99141Y="";
				str99141Y= (String) h.get("99141Y");//96.07.27 add 列印99141Y最近決算年度
				System.out.println("str99141Y="+str99141Y);		
				if((Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH)) >= 10810){
					table = table_10810;
				}else{
					table = table_10809;
				}
				for (int i = 0; i < table.length; i++) {
					//System.out.print("Row =" + rowNo+":"+table[i]);
					row = sheet.getRow(rowNo++);
					cell = row.getCell((short) 2);
					String amt = "0";
					if(h.get(table[i]) == null) continue;
					if(table[i].equals("990140")){//990140-990150
					   //96.3.9 fix 990140存款總額(含公庫存庫).因含公庫存款.需扣除990150公庫存款 	
					   tmpAmt = new BigDecimal(((String) h.get(table[i])));
					   tmpAmt = tmpAmt.add( new BigDecimal(((String) h.get("990150"))).negate());//990140-990150					    
					   amt = Utility.setCommaFormat(tmpAmt.toString());
					}else if(table[i].equals("990150")){//990150--(5)1/2公庫存款
					   if((String)h.get(table[i]) != null){
					      tmpAmt = new BigDecimal(((String) h.get(table[i])));					      
					      if(((String) h.get(table[i]) != null) && Integer.parseInt((String) h.get(table[i])) != 0 ){					        
					        tmpAmt = new BigDecimal((double)tmpAmt.multiply(tmp05).intValue());//990150/2 98.03.25 fix 					       
					        amt = Utility.setCommaFormat(tmpAmt.toString());
					      }
					   }					   
					}else if(table[i].equals("990230") && i==12){
					    //990230--(13)上 (　)年度信用部決算淨值(扣除前一年度最近一次檢查應補提未提足之備抵呆帳)
					    tmpAmt = new BigDecimal(((String) h.get(table[i])));
					    tmpAmt = tmpAmt.add( new BigDecimal(((String) h.get("990240"))).negate());//990230-990240					    
					    amt = Utility.setCommaFormat(tmpAmt.toString());
					    //96.07.27 加印99141Y最近決算年度
					    cell = row.getCell((short) 1);
					    if(!str99141Y.equals("")){		
					    	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					    	cell.setCellValue("(13)上 ("+str99141Y+")年度信用部決算淨值(扣除前一年度最近一次檢查應補提未提足之備抵呆帳)");
					    }
					    cell = row.getCell((short) 2);
					}else if(table[i].equals("990230") && i==13){
					    //990230--(14)上 (　)年度信用部決算淨值
						amt = Utility.setCommaFormat((String) h.get(table[i]));
						//96.07.27 加印99141Y最近決算年度
					    cell = row.getCell((short) 1);
					    if(!str99141Y.equals("")){		
					    	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					    	cell.setCellValue("(14)上 ("+str99141Y+")年度信用部決算淨值");
					    }
					    cell = row.getCell((short) 2);  
					}else if(table[i].equals("990320")){
					    //990320--(15)上 (　)年度全體決算淨值
						amt = Utility.setCommaFormat((String) h.get(table[i])); 
						//96.07.27 加印99141Y最近決算年度
					    cell = row.getCell((short) 1);
					    if(!str99141Y.equals("")){	
					    	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					    	cell.setCellValue("(15)上 ("+str99141Y+")年度全體決算淨值");
					    }
					    cell = row.getCell((short) 2);        
					}else if(table[i].equals("990620")){//97.04.18帳面金額已內含公庫存款
					    //990620--(21)非會員存款總額（含公庫存款）
					    tmpAmt = new BigDecimal(((String) h.get(table[i])));//990620
					    //tmpAmt = tmpAmt.add( new BigDecimal((String) h.get("990630")));//990620+990630					    
					    amt = Utility.setCommaFormat(tmpAmt.toString());
					}else if(table[i].equals("990630")){
					    //990630--(25) 1/2公庫存款
					    if((String)h.get(table[i]) != null){
					        tmpAmt = new BigDecimal(((String) h.get(table[i])));					      
						    if(((String) h.get(table[i]) != null) && Integer.parseInt((String) h.get(table[i])) != 0 ){					        
						        tmpAmt = new BigDecimal((double)tmpAmt.multiply(tmp05).intValue());//990630/2 98.03.25 fix 				         
						        amt = Utility.setCommaFormat(tmpAmt.toString());
						    }
						}
					}else if(table[i].equals("992190") || table[i].equals("992200") || table[i].equals("992300")){//員工人數
					    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					    amt = Utility.setCommaFormat((String) h.get(table[i]));
					    amt += "人";
					    
					}else if(table[i].equals("992440")){//若為111年9月以後,改讀取990860							   
						 //111.9月後,改讀取A02.990860
						 if((Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH)) >= 11109){
						   tmpAmt = new BigDecimal(((String) h.get("990860")));
						 }
						 amt = Utility.setCommaFormat(tmpAmt.toString());
					}else{
					    amt = Utility.setCommaFormat((String) h.get(table[i]));    
					}
					//System.out.println(":amt =" + amt);
					if (!amt.equals("0")) {
						cell.setCellValue(amt);
					}
				}
			}			

            //設定涷結欄位            
            footer.setCenter( HSSFFooter.page());		                                 
			FileOutputStream fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + (bank_type.equals("6")?"農會":"漁會")+"信用部法定比率分析統計表.xls");
			wb.write(fout);
			// 儲存
			fout.close();
		} catch (Exception e) {
			System.out.println("createRpt Error:" + e + e.getMessage());
		}
		return errMsg;
	}
	
	//96.04.17 add A99_extra中央存保格式A02.A01.A99檔案下載
	//96.04.24 fix A99_extra中央存保格式A02.A01.A99檔案下載改用回傳List 
	public static List printData(String S_YEAR, String S_MONTH, String bank_code, String bank_type,String unit) {	                               
		String errMsg = "";
		List dbData = null;
		String sqlCmd = "";
		String ncacno="";
		String unit_name="";
		BigDecimal tmpZero=new BigDecimal("0");		
		HashMap h = new HashMap();//儲存data
		List bank_code_list = new LinkedList();		
		HashMap bank_code_h = new HashMap();//儲存最近一年data		
		String printResult="";
		List return_List = new LinkedList();
		String cd01_table = "";
        String wlx01_m_year = "";
		List paramList = new ArrayList();
		List a02_table_paramList = new ArrayList(); 
		try {
			//100.02.22 add 查詢年度100年以前.縣市別不同===============================
  	    	cd01_table = (Integer.parseInt(S_YEAR) < 100)?"cd01_99":""; 
  	    	wlx01_m_year = (Integer.parseInt(S_YEAR) < 100)?"99":"100"; 
  	    	//=====================================================================   
			ncacno = bank_type.equals("6")?"ncacno":"ncacno";	
			/* 取得最近一年/一季平均存放比率
			 * select a02_12.bank_code,
			 		  a02_12.amt990110_12,a02_12.amt990120_12,a02_12.amt990130_12,a02_12.amt990140_12,a02_12.amt990150_12,
			 		  a02_3.amt990110_3,a02_3.amt990120_3,a02_3.amt990130_3,a02_3.amt990140_3,a02_3.amt990150_3 
			   from ( select bank_code  -- 最近一年
				    	   ,round(sum(decode(acc_code,'990110',amt,0))/12/1,0) as amt990110_12  
				    	   ,round(sum(decode(acc_code,'990120',amt,0))/12/1,0) as amt990120_12  
				    	   ,round(sum(decode(acc_code,'990130',amt,0))/12/1 ,0) as amt990130_12  
				    	   ,round((sum(decode(acc_code,'990140',amt,0))  - sum(decode(acc_code,'990150',amt,0)))/12/1,0)  as amt990140_12  
				    	   ,round((sum(decode(acc_code,'990150',amt,0))/2)/12/1,0) as amt990150_12  
				    from a02  
				    where to_char(m_year * 100 + m_month) >= (select TO_CHAR(ADD_MONTHS(TO_DATE('20060501','YYYY/MM/DD'),-11),'YYYYMM') - 191100 from dual)  
				    and to_char(m_year * 100 + m_month) <= 9505
				    group by bank_code 
				    )a02_12 left join ( select bank_code --最近一季 
				    						   ,round(sum(decode(acc_code,'990110',amt,0))/3/1,0) as amt990110_3
				    						   ,round(sum(decode(acc_code,'990120',amt,0))/3/1,0) as amt990120_3  
				    						   ,round(sum(decode(acc_code,'990130',amt,0))/3/1 ,0) as amt990130_3  
				    						   ,round((sum(decode(acc_code,'990140',amt,0))  - sum(decode(acc_code,'990150',amt,0)))/3/1,0)  as amt990140_3  
				    						   ,round((sum(decode(acc_code,'990150',amt,0))/2)/3/1,0) as amt990150_3  
				                        from a02  
				                        where to_char(m_year * 100 + m_month) >= (select TO_CHAR(ADD_MONTHS(TO_DATE('20060501','YYYY/MM/DD'),-2),'YYYYMM') - 191100 from dual)  
				                        and to_char(m_year * 100 + m_month) <= 9505
				                        group by bank_code )a02_3 on a02_12.bank_code = a02_3.bank_code
                            left join (select bank_code --(33)放款餘額(季平均)
                            				 ,round(sum(decode(acc_code,'120000',amt,'120800',amt,'150300',amt,0))/3,0) as amt120000
                            		   from a01
                            		   where to_char(m_year * 100 + m_month) >= (select TO_CHAR(ADD_MONTHS(TO_DATE(200709,'YYYYMM'),-2),'YYYYMM') - 191100 from dual)
                            		   and to_char(m_year * 100 + m_month) <= 9609						 
                            		   group by bank_code)a01_3 on a02_12.bank_code=a01_3.bank_code				                        	
               order by a02_12.bank_code		
			 */
			
	        
			String a02_table = " (select a.m_year,a.m_month,a.bank_code "
							 + " ,round(sum(decode(acc_code,'990210',amt,0))/1,0) as amt990210 "
							 + " ,round(sum(decode(acc_code,'990220',amt,0))/1,0) as amt990220 "
							 + " ,round(sum(decode(acc_code,'99141Y',amt,0))/1,0) as amt99141Y "
							 + " ,round((sum(decode(acc_code,'990230',amt,0)) - sum(decode(acc_code,'990240',amt,0)))/1,0) as amt990230_990240 "
							 + " ,round(sum(decode(acc_code,'990230',amt,0))/1,0) as amt990230 "
							 + " ,round(sum(decode(acc_code,'990320',amt,0))/1,0) as amt990320 "
							 + " ,round(sum(decode(acc_code,'992100',amt,0))/1,0) as amt992100 "
							 + " ,round(sum(decode(acc_code,'992110',amt,0))/1,0) as amt992110 "
							 + " ,round(sum(decode(acc_code,'992120',amt,0))/1,0) as amt992120 "
							 + " ,round(sum(decode(acc_code,'992130',amt,0))/1,0) as amt992130 "
							 + " ,round(sum(decode(acc_code,'990420',amt,0))/1,0) as amt990420 "
							 + " ,round(sum(decode(acc_code,'990620',amt,0))/1,0) as amt990620 "//97.04.18帳面金額已內含公庫存款.不用再加990630公庫存款
							 //+ " ,round(sum(decode(acc_code,'990620',amt,'990630',amt,0))/1,0)  as amt990620 //97.04.18 fix 
							 + " ,round(sum(decode(acc_code,'992140',amt,0))/1,0) as amt992140 "
							 + " ,round(sum(decode(acc_code,'990410',amt,0))/1,0) as amt990410 "
							 + " ,round(sum(decode(acc_code,'990610',amt,0))/1,0) as amt990610 "
							 + " ,round(sum(decode(acc_code,'990630',amt,0))/2/1,0) as amt990630 "//1/2公庫存款
							 + " ,round(sum(decode(acc_code,'991020',amt,0))/1,0) as amt991020 "
							 + " ,round(sum(decode(acc_code,'990510',amt,0))/1,0) as amt990510 "
							 + " ,round(sum(decode(acc_code,'990511',amt,0))/1,0) as amt990511 "//98.07.27 add 990511非會員無擔保消費性政策貸款
							 + " ,round(sum(decode(acc_code,'990512',amt,0))/1,0) as amt990512 "//98.07.27 add 990512非會員無擔保消費性非政策貸款
							 + " ,round(sum(decode(acc_code,'992150',amt,0))/1,0) as amt992150 "
							 + " ,round(sum(decode(acc_code,'990810',amt,0))/1,0) as amt990810 "
							 + " ,round(sum(decode(acc_code,'990910',amt,0))/1,0) as amt990910 "                             
							 + " ,round(sum(decode(acc_code,'990920',amt,0))/1,0) as amt990920 "                               
							 + " ,round(sum(decode(acc_code,'990710',amt,0))/1,0) as amt990710 "                              
							 + " ,round(sum(decode(acc_code,'120000',amt,0))/1,0) as amt120000 " //放款總額=120000+120800+150300 //(33)放款餘額(季平均)                            
							 + " ,round(sum(decode(acc_code,'992440',amt,0))/1,0) as amt992440 "                               
							 + " ,round(sum(decode(acc_code,'990860',amt,0))/1,0) as amt990860 " //111.06.07 add 880860持有非由政府發行之債券及票卷餘額 
							 + " ,round(sum(decode(acc_code,'992170',amt,0))/1,0) as amt992170 "                               
							 + " ,round(sum(decode(acc_code,'992180',amt,0))/1,0) as amt992180 "          
							 + " ,round(sum(decode(acc_code,'992190',amt,0))/1,0) as amt992190 "                               
							 + " ,round(sum(decode(acc_code,'992200',amt,0))/1,0) as amt992200 "                               
							 + " ,round(sum(decode(acc_code,'992300',amt,0))/1,0) as amt992300 "                               
							 + " ,round(sum(decode(acc_code,'992510',amt,0))/1,0) as amt992510 "                               
							 + " ,round(sum(decode(acc_code,'992520',amt,0))/1,0) as amt992520 "                               
							 + " ,round(sum(decode(acc_code,'992530',amt,0))/1,0) as amt992530 "                               
							 + " ,round(sum(decode(acc_code,'992540',amt,0))/1,0) as amt992540 "                               
							 + " ,round(sum(decode(acc_code,'992610',amt,0))/1,0) as amt992610 "   
							 + " ,round(sum(decode(acc_code,'992620',amt,0))/1,0) as amt992620 "                               
							 + " ,round(sum(decode(acc_code,'992630',amt,0))/1,0) as amt992630 "                               
							 + " ,round(sum(decode(acc_code,'992640',amt,0))/1,0) as amt992640 "                               
							 + " ,round(sum(decode(acc_code,'990611',amt,0))/1,0) as amt990611 "                               
							 + " ,round(sum(decode(acc_code,'992710',amt,0))/1,0) as amt992710 "                               
							 + " ,round(sum(decode(acc_code,'992720',amt,0))/1,0) as amt992720 "                               
							 + " ,round(sum(decode(acc_code,'992730',amt,0))/1,0) as amt992730 "
							 + " ,round(sum(decode(acc_code,'992550',amt,0))/1,0) as amt992550 "//無擔保消費性貸款中之逾期放款
							 + " ,round(sum(decode(acc_code,'992650',amt,0))/1,0) as amt992650 "//無擔保消費性貸款中之應予觀察放款
							 + " ,round(sum(decode(acc_code,'990612',amt,0))/1,0) as amt990612 "//非會員放款中之政策性農業專案貸款餘額
							 + " ,round(sum(decode(acc_code,'992810',amt,0))/1,0) as amt992810 "//98.07.27 add 992810投資全國農業金庫尚未攤提之損失
							 + " ,round(sum(decode(acc_code,'990711',amt,0))/1,0) as amt990711 "//108.09.17 add 990711購置住宅放款餘額
							 + " ,round(sum(decode(acc_code,'990712',amt,0))/1,0) as amt990712 "//108.09.17 add 990712房屋修繕放款餘額
							 + " from ( "
							 + "  select a02.m_year,a02.m_month,a02.bank_code,"+ncacno+".acc_code,"+ncacno+".acc_name,decode("+ncacno+".acc_code,'992190',amt,'992200',amt,'992300',amt,'99141Y',amt,round(nvl(a02.amt,0)/1,0)) as amt "
						     + "  from "+ncacno+" left join ("
							 + " 						  select m_year,m_month,bank_code,acc_code,amt from a02 " 
							 + " 					      where a02.m_year = ? and a02.m_month = ?" 
							 //+ " 						  and   bank_code = '" + bank_code+"'"
							 + " 						  union " 
							 + " 						  select * from a99 " 
							 + " 						  where a99.m_year = ? and a99.m_month = ?"
							 //+ " 						  and   bank_code = '" + bank_code+"'"
							 + " 						)a02 "
							 + "						on "+ncacno+".acc_code=a02.acc_code "
							 + " where ("+ncacno+".acc_tr_type='A02' or "+ncacno+".acc_tr_type='A99')"
							 + " and a02.m_year is not null and a02.m_month is not null"
							 + " group by a02.m_year,a02.m_month,a02.bank_code,"+ncacno+".acc_tr_type,"+ncacno+".acc_code,"+ncacno+".acc_name, "
							 + " decode("+ncacno+".acc_code,'992190',amt,'992200',amt,'992300',amt,'99141Y',amt,round(nvl(a02.amt,0)/1,0))"
							 + " order by a02.m_year,a02.m_month,a02.bank_code,"+ncacno+".acc_tr_type"
							 + " )a group by a.m_year,a.m_month,a.bank_code)a02 ";

			a02_table_paramList.add(S_YEAR);
			a02_table_paramList.add(S_MONTH);
			a02_table_paramList.add(S_YEAR);
			a02_table_paramList.add(S_MONTH);
			
			
			sqlCmd = " select ba01.bank_type,a02_12.bank_code,"
				   + " 	 	  a02_12.amt990110_12,a02_12.amt990120_12,a02_12.amt990130_12,a02_12.amt990140_12,a02_12.amt990150_12,"
				   + "		  a02_3.amt990110_3,a02_3.amt990120_3,a02_3.amt990130_3,a02_3.amt990140_3,a02_3.amt990150_3,a01_3.amt120000,"
				   + "        amt990210,amt990220,amt99141Y,amt990230_990240,"
				   + "        amt990230,amt990320,amt992100,amt992110,amt992120,amt992130,amt990420,amt990620,amt992140,amt990410,amt990610,amt990630,"        
				   + "        amt991020,amt990510,amt990511,amt990512,amt992150,amt990810,amt990910,amt990920,amt990710,a02.amt120000,amt992440,amt990860,amt992170,"        
				   + "        amt992180,amt992190,amt992200,amt992300,amt992510,amt992520,amt992530,amt992540,amt992610,amt992620,"        
				   + "        amt992630,amt992640,amt990611,amt992710,amt992720,amt992730,amt992550,amt992650,amt990612,amt992810,amt990711,amt990712 "				    
				   + " from ( select bank_code " // -- 最近一年
		    	   + " 				 ,round(sum(decode(acc_code,'990110',amt,0))/12/1,0) as amt990110_12 "
		    	   + "				 ,round(sum(decode(acc_code,'990120',amt,0))/12/1,0) as amt990120_12 " 
		    	   + " 				 ,round(sum(decode(acc_code,'990130',amt,0))/12/1 ,0) as amt990130_12 " 
		    	   + "				 ,round((sum(decode(acc_code,'990140',amt,0))  - sum(decode(acc_code,'990150',amt,0)))/12/1,0)  as amt990140_12 "  
		    	   + "				 ,round((sum(decode(acc_code,'990150',amt,0))/2)/12/1,0) as amt990150_12 " 
				   + "		  from a02 " 
				   + " 		  where to_char(m_year * 100 + m_month) >= (select TO_CHAR(ADD_MONTHS(TO_DATE(?,'YYYY/MM/DD'),-11),'YYYYMM') - 191100 from dual)"  
				   + " 		  and to_char(m_year * 100 + m_month) <= to_number(?)"
				   + " 		  group by bank_code )a02_12 "
				   + "		  left join ( select bank_code "//--最近一季 
				   + "				   		     ,round(sum(decode(acc_code,'990110',amt,0))/3/1,0) as amt990110_3 "
				   + "				   			 ,round(sum(decode(acc_code,'990120',amt,0))/3/1,0) as amt990120_3 " 
				   + "				   			 ,round(sum(decode(acc_code,'990130',amt,0))/3/1 ,0) as amt990130_3 " 
				   + "				   			 ,round((sum(decode(acc_code,'990140',amt,0))  - sum(decode(acc_code,'990150',amt,0)))/3/1,0)  as amt990140_3 "  
				   + "				   			 ,round((sum(decode(acc_code,'990150',amt,0))/2)/3/1,0) as amt990150_3  "
				   + "             		  from a02  "
				   + "             		  where to_char(m_year * 100 + m_month) >= (select TO_CHAR(ADD_MONTHS(TO_DATE(?,'YYYY/MM/DD'),-2),'YYYYMM') - 191100 from dual)"  
				   + "             		  and to_char(m_year * 100 + m_month) <= to_number(?)" 
				   + "             		  group by bank_code )a02_3 on a02_12.bank_code = a02_3.bank_code "	
				   + "		  left join ( select bank_code "//--(33)放款餘額(季平均).96.10.30 fix 120000+120800+150300
       			   + "		 				 	 ,round(sum(decode(acc_code,'120000',amt,'120800',amt,'150300',amt,0))/3,0) as amt120000"
				   + "				      from a01"
				   + "             		  where to_char(m_year * 100 + m_month) >= (select TO_CHAR(ADD_MONTHS(TO_DATE(?,'YYYY/MM/DD'),-2),'YYYYMM') - 191100 from dual)"  
				   + "             		  and to_char(m_year * 100 + m_month) <= to_number(?)" 		 
				   + "					  group by bank_code)a01_3 on a02_12.bank_code=a01_3.bank_code"	
				   + " left join " + a02_table + " on a02_12.bank_code = a02.bank_code "	
				   + " left join (select bank_no,bank_name,bank_type from ba01 where ba01.m_year=?)ba01 on ba01.bank_no = a02_12.bank_code ";
			paramList.add(Utility.getFullDate(S_YEAR,S_MONTH,"01"));
			paramList.add(S_YEAR+S_MONTH);
			paramList.add(Utility.getFullDate(S_YEAR,S_MONTH,"01"));
			paramList.add(S_YEAR+S_MONTH);
			paramList.add(Utility.getFullDate(S_YEAR,S_MONTH,"01"));
			paramList.add(S_YEAR+S_MONTH);
			for(int i=0;i<a02_table_paramList.size();i++){
				paramList.add(a02_table_paramList.get(i));
			}
			paramList.add(wlx01_m_year);
			if(bank_code.length() == 7){
			   sqlCmd += " where a02_12.bank_code = ?";
			   paramList.add(bank_code);
			}else{
			   sqlCmd += " where a02_12."+bank_code;
			}
			sqlCmd += " order by a02_12.bank_code ";	
			dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList, "amt990110_12,amt990120_12,amt990130_12,amt990140_12,amt990150_12,amt990110_3,amt990120_3,amt990130_3,amt990140_3,amt990150_3,amt120000"
											  + ",amt990210,amt990220,amt99141y,amt990230_990240,"
											  + "amt990230,amt990320,amt992100,amt992110,amt992120,amt992130,amt990420,amt990620,amt992140,amt990410,amt990610,amt990630,"        
											  + "amt991020,amt990510,amt990511,amt990512,amt992150,amt990810,amt990910,amt990920,amt990710,amt120000,amt992440,amt990860,amt992170,"        
											  + "amt992180,amt992190,amt992200,amt992300,amt992510,amt992520,amt992530,amt992540,amt992610,amt992620,"        
											  + "amt992630,amt992640,amt990611,amt992710,amt992720,amt992730,amt992550,amt992650,amt990612,amt992810,amt990711,amt990712");
			// 取出資料存入MAP
			System.out.println("dbData.size()="+dbData.size());
			if((Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH)) >= 10810){
				print_table = print_table_10810;
			}else{
				print_table = print_table_10809;
			}
			for (int k = 0; k < dbData.size(); k++) {				
				DataObject obj = (DataObject) dbData.get(k);
				h = new HashMap();//儲存data
				for (int i = 0; i < print_table.length; i++) {
					if(obj.getValue("amt"+print_table[i]) != null){	
						//111.6月後,改讀取A02.990860
					   if(print_table[i].equals("992440") && ((Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH)) >= 11103)){
						   h.put("amt"+print_table[i], obj.getValue("amt990860").toString());
						   System.out.println((String)obj.getValue("bank_code")+".amt"+print_table[i]+"="+obj.getValue("amt990860").toString());
					   }else{
						   h.put("amt"+print_table[i], obj.getValue("amt"+print_table[i]).toString());
						   System.out.println((String)obj.getValue("bank_code")+".amt"+print_table[i]+"="+obj.getValue("amt"+print_table[i]).toString());
					   }
					}else{
					   System.out.println((String)obj.getValue("bank_code")+".amt"+print_table[i]+"= null");
					}
				}		
				bank_code_list.add((String)obj.getValue("bank_code"));
				bank_code_h.put((String)obj.getValue("bank_code"),h);
			}
			
			int countidx = 98000;
			for(int j = 0;j<bank_code_list.size();j++){
				countidx = 980000;//96.04.26 A99_extra中央存保格式A02.A01.A99檔案下載科目代號改為6碼
				for(int i = 0; i < print_table.length; i++) {	
					countidx++;
					h = (HashMap)bank_code_h.get((String)bank_code_list.get(j));	
					if(h.get("amt"+print_table[i]) != null){
						return_List.add(DownLoad.fillStuff(S_YEAR, "L", "0", 3)//年			   	   			   	   
   	                                +  DownLoad.fillStuff(S_MONTH, "L", "0", 2)//月
   	                                +  DownLoad.fillStuff((String)bank_code_list.get(j), "R", "0", 7)//機構代號
									+  String.valueOf(countidx)//科目編號
							        +  DownLoad.fillStuff((String)h.get("amt"+print_table[i]), "L", "0", 0, 14)//金額
									); 
						/*printResult += DownLoad.fillStuff(S_YEAR, "L", "0", 3)//年			   	   			   	   
   	                                +  DownLoad.fillStuff(S_MONTH, "L", "0", 2)//月
   	                                +  DownLoad.fillStuff((String)bank_code_list.get(j), "R", "0", 7)//機構代號
									+  String.valueOf(countidx)//科目編號
							        +  DownLoad.fillStuff((String)h.get("amt"+print_table[i]), "L", "0", 0, 14)//金額
									+ "\n";*/							
				    } 
				}
			}   
			/*
			if(printResult.lastIndexOf("\n") != -1){
			   printResult = printResult.substring(0,printResult.lastIndexOf("\n"));
			}*/
			//System.out.println("printResult="+printResult);
		} catch (Exception e) {
			System.out.println("printData Error:" + e + e.getMessage());
		}
		//return printResult;
		return return_List;
	}
}
