/*
  Created on 2006/03/30 by 2495
  99.04.13 fix 縣市合併SQL調整 && 修改查詢方式為preparedstatement by 2808
 103.01.16 add 臺灣省農會更名為中華民國農會增加說明 by 2295
 103.12.23 fix 調整title出現奇怪的線 by 2295
*/
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import java.io.*;
import java.util.*;
import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR036W{      
	
  	public static String createRpt(String S_YEAR,String S_MONTH,String unit,String bank_type,String bank_list,String rptStyle) {
  		System.out.println("start RptFR036W------------------");
		String errMsg = "";		
		String S_YEAR_last="";
		String S_MONTH_last="";	 
		String bank_type_name=(bank_type.equals("6"))?"農會":"漁會";	
		String filename="";
		String openfile="";
		int rowcount=0;
		int flag=0;
		String u_year = "100" ; //縣市合併判斷用 
        if(S_YEAR==null || Integer.parseInt(S_YEAR) < 100) {
        	u_year = "99" ;
        }
				
        if(bank_type.equals("6")){	
				if(rptStyle.equals("0")){
					filename="全體農會信用部統一農貸資料總表.xls";
				}else{
					filename="全體農會信用部統一農貸資料明細表.xls";
				}
				openfile = filename; 
				
		}else{
				if(rptStyle.equals("0")){
					filename="全體漁會信用部統一漁貸資料總表.xls";
				}else{
					filename="全體漁會信用部統一漁貸資料明細表.xls";
				}
				openfile = filename; 
		}
	   	   	 
	   try{				
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
		    FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+ openfile );			
			
	  	    //設定FileINputStream讀取Excel檔
	  		POIFSFileSystem fs = new POIFSFileSystem( finput );
	  		if(fs==null){System.out.println("open 範本檔失敗");} 
	  		HSSFWorkbook wb = new HSSFWorkbook(fs);
	  		if(wb==null){System.out.println("open工作表失敗");}
	  		HSSFSheet sheet = wb.getSheetAt(0);//讀取第一個工作表，宣告其為sheet 
	  		if(sheet==null){System.out.println("open sheet 失敗");}
	  		HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
	        //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
	        //sheet.setAutobreaks(true); //自動分頁
			
	        //設定頁面符合列印大小
	        sheet.setAutobreaks( false );
	        ps.setScale( ( short )70 ); //列印縮放百分比

	        ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
	        
	        //設定表頭 為固定 先設欄的起始再設列的起始
	        wb.setRepeatingRowsAndColumns(0, 1, 10, 1, 3);
	  		
	  		finput.close();
	  		
	  		
	  		HSSFRow row=null;//宣告一列 
	  		HSSFCell cell=null;//宣告一個儲存格

			HSSFFont ft = wb.createFont();
			HSSFCellStyle cs = wb.createCellStyle();
			HSSFFont ft2 = wb.createFont();
			HSSFCellStyle cs2 = wb.createCellStyle();
			HSSFCellStyle cs3 = wb.createCellStyle();
			HSSFFont f = wb.createFont();
			HSSFCellStyle c = wb.createCellStyle();
		     
			 
			
      		 row = sheet.getRow(0);
      		 cell = row.getCell( (short) 0);
      		 cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        	 //cell.setCellValue( bank_type_name +"管理人員名冊");
		 	 if(bank_type.equals("6")){	
				if(rptStyle.equals("0")){
					cell.setCellValue("全體農會信用部統一農貸資料總表");
				}
				if(rptStyle.equals("1")){
					cell.setCellValue("全體農會信用部統一農貸資料明細表");
				}
			}
			if(bank_type.equals("7")){
				if(rptStyle.equals("0"))
				{
					cell.setCellValue("全體漁會信用部統一漁貸資料總表");
				}
				if(rptStyle.equals("1"))
				{
					cell.setCellValue("全體漁會信用部統一漁貸資料明細表");
				}
			}	
		 
  		       

			int rowNum =1;
	
			//取得bank_list的bank_no
			String bank_id = "";
			String sqlCmd="";	
	        StringBuffer sql = new StringBuffer () ;
	        List paramList = new ArrayList() ;
			if(rptStyle.equals("0")){ //總表
				sql.append(getReport1SQl(u_year,unit)) ;
				//paramList = getReport1_paramList( u_year ,unit,S_YEAR , S_MONTH, bank_type) ;
				paramList.add(u_year) ;
		  		paramList.add(bank_type) ;
		  		paramList.add(u_year) ;
		  		paramList.add(S_YEAR) ;
		  		paramList.add(S_MONTH) ;
		  		paramList.add(u_year) ;
		  		paramList.add(bank_type) ;
		  		paramList.add(u_year) ;
		  		paramList.add(S_YEAR) ;
		  		paramList.add(S_MONTH) ;
			}
	
			if(rptStyle.equals("1") ){	
				//================================================================================================================		
				if(rptStyle.equals("1")){ //明細
					if("100".equals(u_year)) {
						sql.append(" select * from    ");
						sql.append(" (                ");
						sql.append("  select nvl(cd01.hsien_id,' ') as  hsien_id ,                 ");
						sql.append("         nvl(cd01.hsien_name,'OTHER') as  hsien_name,          ");
						sql.append("         cd01.FR001W_output_order     as  FR001W_output_order, ");
						sql.append("         bn01.bank_no ,  bn01.BANK_NAME , CreditMonth_Cnt,     ");
						sql.append("         Round(CreditMonth_Amt/? ,0)  as CreditMonth_Amt,       ");
						sql.append("         CreditYear_Cnt_Acc,                                   ");
						sql.append("         Round(CreditYear_Amt_Acc/?,0) as CreditYear_Amt_Acc,  ");
						sql.append("         Credit_Cnt,                                           ");
						sql.append("         Round(Credit_Bal/?,0)  as Credit_Bal,                 ");
						sql.append("         OverCreditMonth_Cnt,                                  ");
						sql.append("         Round(OverCreditMonth_Amt/?,0) as  OverCreditMonth_Amt,");
						sql.append("         OverCredit_Cnt,                                        ");
						sql.append("         Round(OverCredit_Bal/?,0)  as  OverCredit_Bal          ");
						sql.append("  from  (select * from cd01 where cd01.hsien_id <> 'Y') cd01    ");
						sql.append("  left join wlx01 on wlx01.hsien_id=cd01.hsien_id               ");
						sql.append("  left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type= ? ");
						sql.append("  	  	   		and bn01.m_year = wlx01.m_year and bn01.m_year = ?     ");
						sql.append("  left join (select * from WLX07_M_CREDIT where m_year  =   ? and m_month =  ?) a01 ");
						sql.append(" 		    on bn01.bank_no = a01.bank_no");
						sql.append(" ) a01							");
						sql.append(" where a01.bank_no  <>  ' ' ");
						sql.append(" and a01.bank_no in (").append(getQueryBankList(bank_list)).append(") ");
						sql.append(" and  a01.CreditYear_Amt_Acc >= 0 ");
						sql.append(" order by  a01.FR001W_output_order, a01.hsien_id,  a01.bank_no ");
					}else {
						sql.append(" select * from   ");
						sql.append(" (               ");
						sql.append("  select nvl(cd01.hsien_id,' ') as  hsien_id ,                      ");
						sql.append("         nvl(cd01.hsien_name,'OTHER') as  hsien_name,               ");
						sql.append("         cd01.FR001W_output_order     as  FR001W_output_order,      ");
						sql.append("         bn01.bank_no ,  bn01.BANK_NAME , CreditMonth_Cnt,          ");
						sql.append("         Round(CreditMonth_Amt/?,0)  as CreditMonth_Amt,            ");
						sql.append("         CreditYear_Cnt_Acc,                                        ");
						sql.append("         Round(CreditYear_Amt_Acc/?,0) as CreditYear_Amt_Acc,       ");
						sql.append("         Credit_Cnt,                                                ");
						sql.append("         Round(Credit_Bal/?,0)  as Credit_Bal,                      ");
						sql.append("         OverCreditMonth_Cnt,                                       ");
						sql.append("         Round(OverCreditMonth_Amt/?,0) as  OverCreditMonth_Amt,    ");
						sql.append("         OverCredit_Cnt,                                            ");
						sql.append("         Round(OverCredit_Bal/?,0)  as  OverCredit_Bal              ");
						sql.append("  from  (select * from cd01_99 cd01 where cd01.hsien_id <> 'Y') cd01");
						sql.append("  left join wlx01 on wlx01.hsien_id=cd01.hsien_id                   ");
						sql.append("  left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=?                ");
						sql.append("  	  	   		and bn01.m_year = wlx01.m_year and bn01.m_year = ?       ");
						sql.append("  left join (select * from WLX07_M_CREDIT where m_year  =   ? and m_month =  ?) a01  ");
						sql.append(" 		    on bn01.bank_no = a01.bank_no                                         ");
						sql.append(" ) a01   ");
						sql.append(" where a01.bank_no  <>  ' '  and  a01.CreditYear_Amt_Acc >= 0         ");
						sql.append(" and a01.bank_no in (").append(getQueryBankList(bank_list)).append(") ");
						sql.append(" order by  a01.FR001W_output_order, a01.hsien_id,  a01.bank_no        ");
					}
					System.out.println("明細表SQL:"+sql.toString()) ;
			
					paramList.add(unit) ;
					paramList.add(unit) ;
					paramList.add(unit) ;
					paramList.add(unit) ;
					paramList.add(unit) ;
					paramList.add(bank_type) ;
					paramList.add(u_year) ;
					paramList.add(S_YEAR) ;
					paramList.add(S_MONTH) ;
				//==============================
		    
				}
		

			}    

			List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"hsien_id,hsien_name,fr001w_output_order,bank_no,bank_name,creditmonth_cnt,creditmonth_amt,credityear_cnt_acc,credityear_amt_acc,credit_cnt,credit_bal,overcreditmonth_cnt,overcreditmonth_amt,overcredit_cnt,overcredit_bal");
		
 				
			System.out.print("明細抓出的dbData="+dbData.size());
			DataObject bean = null;
			String unit_name="";
			if (rptStyle.equals("1")) {
				if (dbData.size() != 0) {
					// 列印年度
					row = (sheet.getRow(1) == null) ? sheet.createRow(1): sheet.getRow(1);
					cell = row.getCell((short) 3);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);

					ft.setFontName("標楷體");
					//ft.setFontHeightInPoints((short) 14);
					cs.setFont(ft);
					
					

					cs.setBorderTop(HSSFCellStyle.BORDER_THIN);
					cs.setBorderBottom(HSSFCellStyle.BORDER_THIN);
					cs.setBorderLeft(HSSFCellStyle.BORDER_THIN);
					cs.setBorderRight(HSSFCellStyle.BORDER_THIN);

					cs.setAlignment(HSSFCellStyle.ALIGN_CENTER);
					//cell.setCellStyle(cs);
					unit_name = Utility.getUnitName(unit);

					if (S_MONTH.equals("0")) {
						cell.setCellValue("中華民國" + S_YEAR + "年度");
					} else {
						cell.setCellValue("             中華民國" + S_YEAR + "年"
								+ S_MONTH
								+ "月                               單位：新台幣"
								+ unit_name + "、％");
					}

					// 列印單位
					cell = row.getCell((short) 11);
					f.setFontName("標楷體");
					f.setFontHeightInPoints((short) 14);
					cs.setFont(f);
					//cell.setCellStyle(cs);
					rowNum = 4;
					
				} else {
					// 列印年度
					row = (sheet.getRow(1) == null) ? sheet.createRow(1): sheet.getRow(1);
					cell = row.getCell((short) 3);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);

					//f.setFontName("標楷體");
					//f.setFontHeightInPoints((short) 14);
					//cs.setFont(f);
					cs.setBorderTop(HSSFCellStyle.BORDER_THIN);
					cs.setBorderBottom(HSSFCellStyle.BORDER_THIN);
					cs.setBorderLeft(HSSFCellStyle.BORDER_THIN);
					cs.setBorderRight(HSSFCellStyle.BORDER_THIN);

					cs.setAlignment(HSSFCellStyle.ALIGN_CENTER);
					//cell.setCellStyle(cs);
					if (S_MONTH.equals("0")) {
						cell.setCellValue("中華民國" + S_YEAR + "年度");
					} else {
						cell.setCellValue("            中華民國" + S_YEAR + "年"
								+ S_MONTH + "月無資料存在");
					}
					System.out.println("debug---無資料");
				}

				if (dbData != null) {
					for (int k = 0; k < dbData.size(); k++) {
						bean = (DataObject) dbData.get(k);
						String hsien_id = String.valueOf(bean.getValue("hsien_id"));
						String hsien_name = String.valueOf(bean.getValue("hsien_name"));
						hsien_name ="null".equals(hsien_name)?"" :hsien_name ;
						String fr001w_output_order = String.valueOf(bean.getValue("fr001w_output_order"));
						String bank_no = String.valueOf(bean.getValue("bank_no"));
						bank_no = "null".equals(bank_no)? "" : bank_no ;
						String bank_name = String.valueOf(bean.getValue("bank_name"));
						bank_name = "null".equals(bank_name) ? "" : bank_name ;
						String creditmonth_cnt = String.valueOf(bean.getValue("creditmonth_cnt"));
						creditmonth_cnt = "null".equals(creditmonth_cnt) ? "0" : creditmonth_cnt ;
						String creditmonth_amt = String.valueOf(bean.getValue("creditmonth_amt"));
						creditmonth_amt  = "null".equals(creditmonth_amt) ? "0" : creditmonth_amt ;
						String credityear_cnt_acc = String.valueOf(bean.getValue("credityear_cnt_acc"));
						credityear_cnt_acc = "null".equals(credityear_cnt_acc) ? "0" :credityear_cnt_acc ;
						String credityear_amt_acc = String.valueOf(bean.getValue("credityear_amt_acc"));
						String credit_cnt = String.valueOf(bean.getValue("credit_cnt"));
						String credit_bal = String.valueOf(bean.getValue("credit_bal"));
						String overcreditmonth_cnt = String.valueOf(bean.getValue("overcreditmonth_cnt"));
						String overcreditmonth_amt = String.valueOf(bean.getValue("overcreditmonth_amt"));
						String overcredit_cnt = String.valueOf(bean.getValue("overcredit_cnt"));
						String overcredit_bal = String.valueOf(bean.getValue("overcredit_bal"));

						
						rowNum++;
						row = (sheet.getRow(rowNum) == null) ? sheet.createRow(rowNum) : sheet.getRow(rowNum);
						for (int cellcount = 0; cellcount < 13; cellcount++) {
							System.out.println("cellcount=" + cellcount);
							cell = row.createCell((short) cellcount);
							cell.setEncoding(HSSFCell.ENCODING_UTF_16);
							cs2.setBorderTop(HSSFCellStyle.BORDER_THIN);
							cs2.setBorderBottom(HSSFCellStyle.BORDER_THIN);
							cs2.setBorderLeft(HSSFCellStyle.BORDER_THIN);
							cs2.setBorderRight(HSSFCellStyle.BORDER_THIN);

							ft.setFontName("標楷體");
							ft.setFontHeightInPoints((short) 12);
							cs2.setFont(ft);
							cs2.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
							cell.setCellStyle(cs2);
							if (cellcount == 0) {
								cell.setCellValue(hsien_name);
							}
							if (cellcount == 1) {
								cell.setCellValue(bank_no);
							}
							if (cellcount == 2) {
								cs3.setBorderTop(HSSFCellStyle.BORDER_THIN);
								cs3.setBorderBottom(HSSFCellStyle.BORDER_THIN);
								cs3.setBorderLeft(HSSFCellStyle.BORDER_THIN);
								cs3.setBorderRight(HSSFCellStyle.BORDER_THIN);
								ft2.setFontName("標楷體");
								ft2.setFontHeightInPoints((short) 12);
								cs3.setFont(ft2);
								cs3.setAlignment(HSSFCellStyle.ALIGN_LEFT);
								cell.setCellStyle(cs3);
								cell.setCellValue(bank_name);
							}
							if (cellcount == 3) {
								cell.setCellValue(creditmonth_cnt);
							}
							if (cellcount == 4) {
								cell.setCellValue(creditmonth_amt);
							}
							if (cellcount == 5) {
								cell.setCellValue(credityear_cnt_acc);
							}
							if (cellcount == 6) {
								cell.setCellValue(credityear_amt_acc);
							}
							if (cellcount == 7) {
								cell.setCellValue(credit_cnt);
							}
							if (cellcount == 8) {
								cell.setCellValue(credit_bal);
							}
							if (cellcount == 9) {
								cell.setCellValue(overcreditmonth_cnt);
							}
							if (cellcount == 10) {
								cell.setCellValue(overcreditmonth_amt);
							}
							if (cellcount == 11) {
								cell.setCellValue(overcredit_cnt);
							}
							if (cellcount == 12) {
								cell.setCellValue(overcredit_bal);
							}
						}
					}
				}
			}
		else{  //總表
			if (dbData.size() > 0) {
					// 列印年度
					System.out.println("dbData.size()=" + dbData.size());
					row = (sheet.getRow(1) == null) ? sheet.createRow(1): sheet.getRow(1);
					cell = row.getCell((short) 1);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);

					f.setFontName("標楷體");
					f.setFontHeightInPoints((short) 10);
					cs2.setFont(f);
					
					//cell.setCellStyle(cs2);
					unit_name = Utility.getUnitName(unit) ;
					
					if (S_MONTH.equals("0")) {
						cell.setCellValue("中華民國" + S_YEAR + "年度");
					} else {
						cell.setCellValue("中華民國" + S_YEAR + "年" + S_MONTH+ "月                             單位：新台幣"+ unit_name + "、％ ");
					}

					// 列印單位
					cell = row.getCell((short) 9);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					f.setFontName("標楷體");
					f.setFontHeightInPoints((short) 10);
					//cs2.setFont(f);
					
					rowNum = 4;
					System.out.print("總表資料dbData.size()" + dbData.size());
				} else {
					System.out.print("總表尚無資料");
					// 列印年度
					System.out.println("dbData.size()=" + dbData.size());
					row = (sheet.getRow(1) == null) ? sheet.createRow(1): sheet.getRow(1);
					cell = row.getCell((short) 1);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);

					f.setFontName("標楷體");
					f.setFontHeightInPoints((short) 10);
					cs.setFont(f);
					cs.setAlignment(HSSFCellStyle.ALIGN_CENTER);
					cell.setCellStyle(cs);
					if (S_MONTH.equals("0")) {
						cell.setCellValue("中華民國" + S_YEAR + "年度");
					} else {
						cell.setCellValue("中華民國" + S_YEAR + "年" + S_MONTH+ "月無資料存在");
					}

				}
				if (dbData != null) {
					for (int k = 0; k < dbData.size(); k++) {
						bean = (DataObject) dbData.get(k);
						String hsien_id = String.valueOf(bean.getValue("hsien_id"));
						if (hsien_id.equals("null"))
							hsien_id = "";
						String hsien_name = String.valueOf(bean.getValue("hsien_name"));
						if (hsien_name.equals("null"))
							hsien_name = "";
						String fr001w_output_order = String.valueOf(bean.getValue("fr001w_output_order"));
						if (fr001w_output_order.equals("null"))
							fr001w_output_order = "";
						String bank_no = String.valueOf(bean.getValue("bank_no"));
						if (bank_no.equals("null"))
							bank_no = "";
						String bank_name = String.valueOf(bean.getValue("bank_name"));
						if (bank_name.equals("null"))
							bank_name = "";
						String creditmonth_cnt = String.valueOf(bean.getValue("creditmonth_cnt"));
						if (creditmonth_cnt.equals("null"))
							creditmonth_cnt = "";
						String creditmonth_amt = String.valueOf(bean.getValue("creditmonth_amt"));
						if (creditmonth_amt.equals("null"))
							creditmonth_amt = "";
						String credityear_cnt_acc = String.valueOf(bean.getValue("credityear_cnt_acc"));
						if (credityear_cnt_acc.equals("null"))
							credityear_cnt_acc = "";
						String credityear_amt_acc = String.valueOf(bean.getValue("credityear_amt_acc"));
						if (credityear_amt_acc.equals("null"))
							credityear_amt_acc = "";
						String credit_cnt = String.valueOf(bean.getValue("credit_cnt"));
						if (credit_cnt.equals("null"))
							credit_cnt = "";
						String credit_bal = String.valueOf(bean.getValue("credit_bal"));
						if (credit_bal.equals("null"))
							credit_bal = "";
						String overcreditmonth_cnt = String.valueOf(bean.getValue("overcreditmonth_cnt"));
						if (overcreditmonth_cnt.equals("null"))
							overcreditmonth_cnt = "";
						String overcreditmonth_amt = String.valueOf(bean.getValue("overcreditmonth_amt"));
						if (overcreditmonth_amt.equals("null"))
							overcreditmonth_amt = "";
						String overcredit_cnt = String.valueOf(bean.getValue("overcredit_cnt"));
						if (overcredit_cnt.equals("null"))
							overcredit_cnt = "";
						String overcredit_bal = String.valueOf(bean.getValue("overcredit_bal"));
						if (overcredit_bal.equals("null"))
							overcredit_bal = "";

						rowNum++;
						row = (sheet.getRow(rowNum) == null) ? sheet.createRow(rowNum) : sheet.getRow(rowNum);
						for (int cellcount = 0; cellcount < 11; cellcount++) {
							cell = row.createCell((short) cellcount);
							cell.setEncoding(HSSFCell.ENCODING_UTF_16);

							cs2.setBorderTop(HSSFCellStyle.BORDER_THIN);
							cs2.setBorderBottom(HSSFCellStyle.BORDER_THIN);
							cs2.setBorderLeft(HSSFCellStyle.BORDER_THIN);
							cs2.setBorderRight(HSSFCellStyle.BORDER_THIN);

							ft.setFontName("標楷體");
							ft.setFontHeightInPoints((short) 12);
							cs2.setFont(ft);
							cs2.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
							cell.setCellStyle(cs2);
							if (cellcount == 0) {
								cell.setCellValue(hsien_name);
							}
							if (cellcount == 1) {
								cell.setCellValue(creditmonth_cnt);
							}
							if (cellcount == 2) {
								cell.setCellValue(creditmonth_amt);
							}
							if (cellcount == 3) {
								cell.setCellValue(credityear_cnt_acc);
							}
							if (cellcount == 4) {
								cell.setCellValue(credityear_amt_acc);
							}
							if (cellcount == 5) {
								cell.setCellValue(credit_cnt);
							}
							if (cellcount == 6) {
								cell.setCellValue(credit_bal);
							}
							if (cellcount == 7) {
								cell.setCellValue(overcreditmonth_cnt);
							}
							if (cellcount == 8) {
								cell.setCellValue(overcreditmonth_amt);
							}
							if (cellcount == 9) {
								cell.setCellValue(overcredit_cnt);
							}
							if (cellcount == 10) {
								cell.setCellValue(overcredit_bal);
							}
						}

					}
					rowNum++;
		            row = sheet.createRow(rowNum);
		            cell = row.createCell( (short)0);
		            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		            cell.setCellValue("縣市別欄之其他(農會)係指原臺灣省農會，該農會於102年5月22日更名為中華民國農會。");		           
				}
			}
    
		HSSFFooter footer=sheet.getFooter();
	    footer.setCenter( "Page:" +HSSFFooter.page() +" of " +HSSFFooter.numPages() );
		footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
		FileOutputStream fout=new FileOutputStream(reportDir+ System.getProperty("file.separator")+ openfile);
		wb.write(fout);
	    //儲存 
	    fout.close();
	    System.out.println("儲存完成");
	        
		}catch(Exception e){
			System.out.println("createRpt Error:"+e+e.getMessage());
		}
		return errMsg;
	}	
  	/***
  	 * 取得總表SQL.
  	 * 
  	 * @param u_year 
  	 * @return
  	 */
  	private static String getReport1SQl(String u_year ,String unit){
  		StringBuffer sql = new StringBuffer () ;
  		String cd01Table = "cd01" ;
  		if("99".equals(u_year)) {
  			cd01Table = "cd01_99" ;
  		}
  		sql.append(" select nvl(cd01.hsien_id,' ') as  hsien_id ,                              ");
  		sql.append("        nvl(cd01.hsien_name,'OTHER')  as  hsien_name,                      ");
  		sql.append("        cd01.FR001W_output_order      as  FR001W_output_order,             ");
  		sql.append("        sum(a01.CreditMonth_Cnt)  as CreditMonth_Cnt,                      ");
  		sql.append("        Round(sum(a01.CreditMonth_Amt)/").append(unit).append(",0)  as CreditMonth_Amt,           ");
  		sql.append("        Sum(a01.CreditYear_Cnt_Acc) as CreditYear_Cnt_Acc,                 ");
  		sql.append("        Round(sum(a01.CreditYear_Amt_Acc)/").append(unit).append(",0)  as CreditYear_Amt_Acc,     ");
  		sql.append("        sum(a01.Credit_Cnt) as Credit_Cnt,                                 ");
  		sql.append("        Round(sum(a01.Credit_Bal)/").append(unit).append(",0)  as Credit_Bal,                     ");
  		sql.append("        Sum(a01.OverCreditMonth_Cnt)  as OverCreditMonth_Cnt,              ");
  		sql.append("        Round(sum(a01.OverCreditMonth_Amt)/").append(unit).append(",0)   as  OverCreditMonth_Amt, ");
  		sql.append("        Sum(a01.OverCredit_Cnt)   as OverCredit_Cnt,                       ");
  		sql.append("        Round(sum(a01.OverCredit_Bal)/").append(unit).append(",0)  as  OverCredit_Bal             ");
  		sql.append(" from  (select * from ").append(cd01Table).append(" where hsien_id <> 'Y') cd01                ");
  		
  		sql.append(" left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = ?    ");
  		sql.append(" left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=? and bn01.m_year = wlx01.m_year and bn01.m_year = ? ");
  		sql.append(" left join (select * from WLX07_M_CREDIT where m_year  = ? and m_month =?) a01 on  bn01.bank_no = a01.bank_no             ");
  		sql.append(" group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order                                     ");
  		sql.append(" union                                                                                                                     ");
  		sql.append(" select ' ' as  hsien_id ,                                                                                                 ");
  		sql.append("        '合計'  hsien_name,                                                                                                ");
  		sql.append("        '999' as  FR001W_output_order,                                                                                     ");
  		sql.append("        sum(a01.CreditMonth_Cnt)  as CreditMonth_Cnt,                                                                      ");
  		sql.append("        Round(sum(a01.CreditMonth_Amt)/").append(unit).append(",0)  as CreditMonth_Amt,                                                           ");
  		sql.append("        Sum(a01.CreditYear_Cnt_Acc) as CreditYear_Cnt_Acc,                                                                 ");
  		sql.append("        Round(sum(a01.CreditYear_Amt_Acc)/").append(unit).append(",0)  as CreditYear_Amt_Acc,                                                     ");
  		sql.append("        sum(a01.Credit_Cnt)  as Credit_Cnt,                                                                                ");
  		sql.append("        Round(sum(a01.Credit_Bal)/").append(unit).append(",0)  as Credit_Bal,                                                                     ");
  		sql.append("        Sum(a01.OverCreditMonth_Cnt)  as OverCreditMonth_Cnt,                                                              ");
  		sql.append("        Round(sum(a01.OverCreditMonth_Amt)/").append(unit).append(",0)   as  OverCreditMonth_Amt,                                                 ");
  		sql.append("        Sum(a01.OverCredit_Cnt)   as OverCredit_Cnt,                                                                       ");
  		sql.append("        Round(sum(a01.OverCredit_Bal)/").append(unit).append(",0)  as  OverCredit_Bal                                                             ");
  		sql.append(" from  (select * from ").append(cd01Table).append(" where hsien_id <> 'Y') cd01                                                                ");
  		sql.append(" left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = ? ");
  		sql.append(" left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=? and bn01.m_year = wlx01.m_year and bn01.m_year = ? ");
  		sql.append(" left join (select * from WLX07_M_CREDIT where m_year  = ? and m_month = ?) a01 on  bn01.bank_no = a01.bank_no           ");
  		sql.append(" order by FR001W_output_order,hsien_id  ");
  		
  		return sql.toString();
  	}
  	/***
  	 * 取得總表的PareparedStatement參數.
  	 * 
  	 * @param u_year
  	 * @param unit
  	 * @param S_YEAR
  	 * @param S_MONTH
  	 * @param bank_type
  	 * @return
  	 */
  	private static List getReport1_paramList(String u_year ,String unit,String S_YEAR ,String S_MONTH,String bank_type) {
  		List paramList = new ArrayList() ;
  		if("99".equals(u_year) ) {
  			paramList.add(unit) ;
			paramList.add(unit) ;
			paramList.add(unit) ;
			paramList.add(unit) ;
			paramList.add(unit) ;
			paramList.add(u_year) ;
			paramList.add(bank_type) ;
			paramList.add(u_year) ;
			paramList.add(S_YEAR) ;
			paramList.add(S_MONTH);
			paramList.add(unit) ;
			paramList.add(unit) ;
			paramList.add(unit) ;
			paramList.add(unit) ;
			paramList.add(unit) ;
			paramList.add(u_year) ;
			paramList.add(bank_type) ;
			paramList.add(u_year) ;
			paramList.add(S_YEAR) ;
			paramList.add(S_MONTH);
  		}else {
  			paramList.add(unit);
  			paramList.add(unit);
  			paramList.add(unit);
  			paramList.add(unit);
  			paramList.add(unit);
  			paramList.add(u_year);
  			paramList.add(bank_type) ;
  			paramList.add(u_year);
  			paramList.add(S_YEAR) ;
  			paramList.add(S_MONTH) ;
  			paramList.add(unit);
  			paramList.add(unit);
  			paramList.add(unit);
  			paramList.add(unit);
  			paramList.add(unit);
  			paramList.add(u_year);
  			paramList.add(bank_type) ;
  			paramList.add(u_year);
  			paramList.add(S_YEAR) ;
  			paramList.add(S_MONTH) ;
  		}
  		return paramList ;
  	}
  	/***
  	 * 取得前端畫面查詢的bankNo,重組為SQL.
  	 * 
  	 * 
  	 * @param bank_list
  	 * @return
  	 */
  	private static String getQueryBankList(String bank_list) {
  		System.out.println("bankno source is :"+bank_list) ;
  		String s ="" ;
  		if(bank_list!=null) {
  			String [] s1 = bank_list.split(",") ;
  			for(String tmp : s1) {
  				tmp  = tmp.substring(0, 7) ;
  				System.out.println(tmp) ;
  				if("".equals(s)) {
  					s = "'".concat(tmp).concat("'") ;
  				}else {
  					s += ",'".concat(tmp).concat("'") ;
  				}
  			}
  			
  		}
  		System.out.println("after program is :"+s) ;
  		return s ;
  	}
}



