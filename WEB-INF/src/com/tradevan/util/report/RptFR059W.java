/*
 *101.07.30 create by 2968
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

public class RptFR059W {
	/*final private static String[][] print_table = {{"loan2_amt",   "990001"},
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
												   
	};	*/

 	
	 public static String createRpt(String M_YEAR, String M_MONTH,String unit,String bank_code ) {    

	    String errMsg = "";
	    List dbData = null;
	    String sqlCmd = "";    
	    int rowNum=0;
	    DataObject bean = null;
	    reportUtil reportUtil = new reportUtil();
		HSSFCellStyle cs_right = null; 
		HSSFCellStyle cs_center = null;
		HSSFCellStyle cs_left = null;
	   
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
	      String filename="M106_貸款用途分析表.xls";
	      //input the standard report form      
	      finput = new FileInputStream(xlsDir +System.getProperty("file.separator") +filename);
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
	      cs_left = reportUtil.getLeftStyle(wb);
	      String unit_name = Utility.getUnitName(unit); 
	   	  /*
	      select loan_use_no,
            cdshareno.cmuse_name,--貸款用途
            guarantee_cnt_year,--當年度件數
            loan_amt_year,--當年度貸款金額
            guarantee_amt_year,--當年度保證金額
            guarantee_cnt_sum,--累計件數
            loan_amt_sum,--累計貸款金額
            guarantee_amt_sum--累計保證金額
            from m106
            left join (select * from cdshareno where cmuse_div='039')cdshareno on M106.LOAN_USE_NO = CDSHARENO.CMUSE_ID
            where m_year = UI.查詢年度 ex:100 and m_month=UI.查詢月份 ex:12
            order by to_number(cdshareno.output_order)
  
          */
	      
	      sql.append("  select  loan_use_no,"
	      		 + "            cdshareno.cmuse_name,"
				 + "            guarantee_cnt_year,"
				 + "            loan_amt_year,"
				 + "            guarantee_amt_year,"
	      	     + "            guarantee_cnt_sum,"
				 + "            loan_amt_sum,"
				 + "            guarantee_amt_sum "
				 + "            from m106 "
				 + "            left join (select * from cdshareno where cmuse_div='039')cdshareno on M106.LOAN_USE_NO = CDSHARENO.CMUSE_ID "
				 + "            where m_year = ? and m_month=? "
				 + "            order by to_number(cdshareno.input_order)" );
	      paramList.add(M_YEAR);
	      paramList.add(M_MONTH);
	     
	      dbData =  DBManager.QueryDB_SQLParam(sql.toString(), paramList, "loan_use_no,cdshareno.cmuse_name,guarantee_cnt_year,loan_amt_year,guarantee_amt_year,guarantee_cnt_sum,loan_amt_sum,guarantee_amt_sum");
	      System.out.println("dbData.size=" + dbData.size());
	     
	      //設定報表表頭資料============================================
	      row=sheet.getRow(3);
	      cell=row.getCell((short)1);	       	
	      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	      cell.setCellValue("中華民國"+M_YEAR+"年"+M_MONTH+"月"+((dbData == null || dbData.size() ==0)?"無資料存在":""));  	
	      
	      if (dbData != null || dbData.size() !=0) { 
            rowNum = 7;
            for (int k=0; k<dbData.size();k++){
	            bean = (DataObject)dbData.get(k);
                String loan_use_no=(bean.getValue("loan_use_no") == null)?"":(bean.getValue("loan_use_no")).toString();
                String cmuse_name=(bean.getValue("cmuse_name") == null)?"":(bean.getValue("cmuse_name")).toString();
                String guarantee_cnt_year=(bean.getValue("guarantee_cnt_year") == null)?"0":(bean.getValue("guarantee_cnt_year")).toString();
                String loan_amt_year=(bean.getValue("loan_amt_year") == null)?"0":(bean.getValue("loan_amt_year")).toString();
                String guarantee_amt_year=(bean.getValue("guarantee_amt_year") == null)?"0":(bean.getValue("guarantee_amt_year")).toString();
                String guarantee_cnt_sum=(bean.getValue("guarantee_cnt_sum") == null)?"0":(bean.getValue("guarantee_cnt_sum")).toString();
                String loan_amt_sum=(bean.getValue("loan_amt_sum") == null)?"0":(bean.getValue("loan_amt_sum")).toString();
                String guarantee_amt_sum=(bean.getValue("guarantee_amt_sum") == null)?"0":(bean.getValue("guarantee_amt_sum")).toString();
                
                row = sheet.createRow(rowNum);
                
                //貸款用途
                cell = row.createCell( (short)1);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                cell.setCellStyle(cs_left);                    
                cell.setCellValue(cmuse_name);
                //當年度件數
                cell = row.createCell( (short)2);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                cell.setCellStyle(cs_right);                    
                cell.setCellValue(Utility.setCommaFormat(guarantee_cnt_year));
                //當年度貸款金額
                cell = row.createCell( (short)3);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                cell.setCellStyle(cs_right);                    
                cell.setCellValue(Utility.setCommaFormat(loan_amt_year));
                //當年度保證金額
                cell = row.createCell( (short)4);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                cell.setCellStyle(cs_right);                    
                cell.setCellValue(Utility.setCommaFormat(guarantee_amt_year));
                //累計件數
                cell = row.createCell( (short)5);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                cell.setCellStyle(cs_right);                    
                cell.setCellValue(Utility.setCommaFormat(guarantee_cnt_sum));
                //累計貸款金額
                cell = row.createCell( (short)6);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                cell.setCellStyle(cs_right);                    
                cell.setCellValue(Utility.setCommaFormat(loan_amt_sum));
                //累計保證金額
                cell = row.createCell( (short)7);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                cell.setCellStyle(cs_right);                    
                cell.setCellValue(Utility.setCommaFormat(guarantee_amt_sum));
                
                rowNum++;
                
	      	} // end of for
	      } //end of else dbData.size() != 0
	      
	      
	      FileOutputStream fout = null;     
	      fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + filename);
	     
	      HSSFFooter footer = sheet.getFooter();
	      footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
	      footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
	      wb.write(fout);
	      //儲存
	      fout.close();
	    }
	    catch (Exception e) {
	    	System.out.println("RptFR059W.createRpt Error:" + e + e.getMessage());
	    }
	    
	    return errMsg;
	  }
	 
	 /*
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
	  	try {
	  		 sql.append(" select a10.m_year,m_month,bank_code,"
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
			   //sqlCmd += " and bank_code = '"+bank_code+"'";
	  			sql.append("  and bank_code =  ? ") ;
	  			paramList.add(bank_code) ;
			}else{
			    //sqlCmd += " and a10."+bank_code;
				sql.append("and a10.").append(bank_code) ;
			}
			//sqlCmd += " order by a10.bank_code ";
	  		sql.append(" order by a10.bank_code ");
	  		
	  		dbData = DBManager.QueryDB_SQLParam(sql.toString(), paramList, "m_year,m_month,loan2_amt,loan3_amt,loan4_amt,loan_sum,invest2_amt,invest3_amt,invest4_amt,invest_sum,other2_amt,other3_amt,other4_amt,other_sum,type2_sum,type3_sum,type4_sum,type_sum"); 
	  			//DBManager.QueryDB(sqlCmd,"m_year,m_month,loan2_amt,loan3_amt,loan4_amt,loan_sum,invest2_amt,invest3_amt,invest4_amt,invest_sum,other2_amt,other3_amt,other4_amt,other_sum,type2_sum,type3_sum,type4_sum,type_sum");
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
	  }*/
	 
}
