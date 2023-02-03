//104.07.24 add by 2968
package com.tradevan.util.report;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.Region;
import java.io.*;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR072W {
	public static String createRpt(String s_year,String s_month) {    
		String errMsg="";
		int rowNum=0;
		int i=0;
		int j=0;
		String filename="全體農漁會信用部人員配置情形表.xls";
        //99.09.10 add 查詢年度100年以前.縣市別不同===============================
        String cd01_table = (Integer.parseInt(s_year) < 100)?"cd01_99":""; 
        String wlx01_m_year = (Integer.parseInt(s_year) < 100)?"99":"100"; 
	    //=====================================================================    

		try {
			File xlsDir=new File(Utility.getProperties("xlsDir"));
			File reportDir=new File(Utility.getProperties("reportDir"));
	
			if(!xlsDir.exists()){
				if(!Utility.mkdirs(Utility.getProperties("xlsDir"))){
			   		errMsg+=Utility.getProperties("xlsDir")+"目錄新增失敗";
				}
			}
			if(!reportDir.exists()){
				if(!Utility.mkdirs(Utility.getProperties("reportDir"))){
			   		errMsg+=Utility.getProperties("reportDir")+"目錄新增失敗";
				}
			}
			
			//String openfile="建築貸款占信用部決算淨值逾100%明細表.xls";
			System.out.println("open file "+filename);
	
			FileInputStream finput=new FileInputStream(xlsDir+System.getProperty("file.separator")+filename );
	
		    //設定FileINputStream讀取Excel檔
			POIFSFileSystem fs=new POIFSFileSystem( finput );
			if(fs==null){System.out.println("open 範本檔失敗");} else System.out.println("open 範本檔成功");
			HSSFWorkbook wb=new HSSFWorkbook(fs);
			if(wb==null){System.out.println("open工作表失敗");}else System.out.println("open 工作表 成功");
			HSSFSheet sheet=wb.getSheetAt(0);//讀取第一個工作表，宣告其為sheet
			if(sheet==null){System.out.println("open sheet 失敗");}else System.out.println("open sheet 成功");
			HSSFPrintSetup ps=sheet.getPrintSetup(); //取得設定
		    //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
		    //sheet.setAutobreaks(true); //自動分頁
	
		    //設定頁面符合列印大小
		    sheet.setAutobreaks( false );
		    //ps.setScale( ( short )65 ); //列印縮放百分比
		    ps.setLandscape( true ); // 設定橫印
		    //ps.setPaperSize( ( short )8 ); //設定紙張大小 A3
		    ps.setPaperSize( (short) 9); //設定紙張大小 A4 (A3:8/A4:9)
			
			finput.close();
			
			HSSFFont ft = wb.createFont();
            HSSFCellStyle cs = wb.createCellStyle();
            ft.setFontHeightInPoints((short)12);
            ft.setFontName("標楷體");
            cs.setFont(ft);
            cs.setBorderLeft(HSSFCellStyle.BORDER_THIN);  
            cs.setBorderTop(HSSFCellStyle.BORDER_THIN);   
            cs.setBorderRight(HSSFCellStyle.BORDER_THIN); 
            cs.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            cs.setAlignment(HSSFCellStyle.ALIGN_CENTER);
			HSSFFont ft1 = wb.createFont();
            HSSFCellStyle cs1 = wb.createCellStyle();
            ft1.setFontHeightInPoints((short)12);
            ft1.setFontName("標楷體");
            cs1.setFont(ft1);
            cs1.setBorderLeft(HSSFCellStyle.BORDER_THIN);  
            cs1.setBorderTop(HSSFCellStyle.BORDER_THIN);   
            cs1.setBorderRight(HSSFCellStyle.BORDER_THIN); 
            cs1.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            cs1.setAlignment(HSSFCellStyle.ALIGN_LEFT);
            HSSFFont ft2 = wb.createFont();
            HSSFCellStyle cs2 = wb.createCellStyle();
            ft2.setFontHeightInPoints((short)12);
            ft2.setFontName("標楷體");
            cs2.setFont(ft2);
            cs2.setBorderLeft(HSSFCellStyle.BORDER_THIN);  
            cs2.setBorderTop(HSSFCellStyle.BORDER_THIN);   
            cs2.setBorderRight(HSSFCellStyle.BORDER_THIN); 
            cs2.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            cs2.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
            HSSFCellStyle cs3 = wb.createCellStyle();
            cs3.setFont(ft1);
            cs3.setBorderLeft(HSSFCellStyle.BORDER_NONE);  
            cs3.setBorderTop(HSSFCellStyle.BORDER_THIN);   
            cs3.setBorderRight(HSSFCellStyle.BORDER_NONE); 
            cs3.setBorderBottom(HSSFCellStyle.BORDER_NONE);
            cs3.setAlignment(HSSFCellStyle.ALIGN_LEFT);
            HSSFCellStyle cs4 = wb.createCellStyle();
            cs4.setFont(ft1);
            cs4.setBorderLeft(HSSFCellStyle.BORDER_NONE);  
            cs4.setBorderTop(HSSFCellStyle.BORDER_NONE);   
            cs4.setBorderRight(HSSFCellStyle.BORDER_NONE); 
            cs4.setBorderBottom(HSSFCellStyle.BORDER_NONE);
            cs4.setAlignment(HSSFCellStyle.ALIGN_LEFT);
			HSSFRow row=null;//宣告一列
			HSSFCell cell=null;//宣告一個儲存格
			
			
		    List dbData=getRPT_MONTHData(s_year,s_month,wlx01_m_year);
			System.out.println("清單.size()="+dbData.size());
			if(dbData==null || dbData.size()==0) dbData=getWLX01Data(s_year,s_month,wlx01_m_year);
			
			//列印報表名稱
			row=(sheet.getRow(1)==null)? sheet.createRow(1) : sheet.getRow(1);
			cell=row.getCell((short)1);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue(s_year+"年"+s_month+"月農(漁)會信用部人員配置情形表");
			row=(sheet.getRow(2)==null)? sheet.createRow(2) : sheet.getRow(2);
			cell=row.getCell((short)10);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("單位：人、%");
			rowNum = 5;
			
			if (dbData !=null && dbData.size() !=0) {
				for(j=0;j<dbData.size();j++){
					DataObject obj = (DataObject)dbData.get(j);
					String bank_name = obj.getValue("bank_name")==null?"":obj.getValue("bank_name").toString();
					String credit_staff = obj.getValue("credit_staff")==null?"0":obj.getValue("credit_staff").toString();
					String credit_staff_rate = obj.getValue("credit_staff_rate")==null?"0":obj.getValue("credit_staff_rate").toString();
					String skill_staff = obj.getValue("skill_staff")==null?"0":obj.getValue("skill_staff").toString();
					String skill_staff_rate = obj.getValue("skill_staff_rate")==null?"0":obj.getValue("skill_staff_rate").toString();
					String manual_staff = obj.getValue("manual_staff")==null?"0":obj.getValue("manual_staff").toString();
					String manual_staff_rate = obj.getValue("manual_staff_rate")==null?"0":obj.getValue("manual_staff_rate").toString();
					String temp_staff = obj.getValue("temp_staff")==null?"0":obj.getValue("temp_staff").toString();
					String temp_staff_rate = obj.getValue("temp_staff_rate")==null?"0":obj.getValue("temp_staff_rate").toString();
					String credit_staff_num = obj.getValue("credit_staff_num")==null?"0":obj.getValue("credit_staff_num").toString();
					
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
					for(int l=0;l<11;l++){
		                cell = row.createCell( (short)l);
		            }
					
					cell=row.getCell((short)1);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs1);
					cell.setCellValue(bank_name);
					
					cell=row.getCell((short)2);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(credit_staff));
					
					cell=row.getCell((short)3);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(credit_staff_rate);
					
					cell=row.getCell((short)4);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(skill_staff));
					
					cell=row.getCell((short)5);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(skill_staff_rate);
					
					cell=row.getCell((short)6);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(manual_staff));
					
					cell=row.getCell((short)7);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(manual_staff_rate);
					
					cell=row.getCell((short)8);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(temp_staff));
					
					cell=row.getCell((short)9);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(temp_staff_rate);
					
					cell=row.getCell((short)10);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(credit_staff_num));
				
					rowNum++;
					
				}
			}
			
			HSSFFooter footer=sheet.getFooter();
			footer.setCenter( "Page:"+HSSFFooter.page()+" of "+HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
	
			FileOutputStream fout=new FileOutputStream(reportDir+System.getProperty("file.separator")+filename);
			wb.write(fout);
			//儲存
			fout.close();
			System.out.println("儲存完成");
		}catch(Exception e) {
			System.out.println("createRpt Error:"+e+e.getMessage());
		}
		return errMsg;
	}
	public static List getRPT_MONTHData(String syear,String smonth,String wlx01_m_year){
		StringBuffer sql = new StringBuffer();
		List paramList = new ArrayList();//共同參數
		sql.append("select bank_code,");
		sql.append("      bank_name,");//單位名稱
		sql.append("      credit_staff,");//正式職員.人數
		sql.append("      decode(credit_staff,0,0,round(credit_staff /  credit_staff_num *100 ,2))  as   credit_staff_rate,");//正式職員.比率
		sql.append("      skill_staff,");//技工.人數
		sql.append("      decode(skill_staff,0,0,round(skill_staff /  credit_staff_num *100 ,2))  as  skill_staff_rate,");//技工.比率
		sql.append("      manual_staff,");//工友.人數
		sql.append("      decode(manual_staff,0,0,round(manual_staff /  credit_staff_num *100 ,2))  as   manual_staff_rate,");//工友.比率
		sql.append("      temp_staff,");//特約人員.人數
		sql.append("      decode(temp_staff,0,0,round(temp_staff /  credit_staff_num *100 ,2))  as  temp_staff_rate,");//特約人員.比率
		sql.append("      credit_staff_num ");//總人數
		sql.append("from ( ");
		sql.append("select rpt_month.m_year,m_month,bank_code,bank_name,");
		sql.append("        sum(decode(acc_code,'credit_staff',amt,0)) as credit_staff ,");//正式職員.人數
		sql.append("        sum(decode(acc_code,'skill_staff',amt,0)) as  skill_staff,");//技工.人數
		sql.append("        sum(decode(acc_code,'manual_staff',amt,0)) as  manual_staff ,");//工友.人數
		sql.append("        sum(decode(acc_code,'temp_staff',amt,0)) as temp_staff,");//特約人員.人數
		sql.append("        sum(decode(acc_code,'credit_staff_num',amt,0)) as credit_staff_num ");//總人數
		sql.append("from rpt_month left join (select * from bn01 where m_year=?)bn01 on rpt_month.bank_code=bn01.bank_no ");
		sql.append("where report_no=? ");
		paramList.add(wlx01_m_year);
		paramList.add("FR072W");
		sql.append("and rpt_month.bank_type in ('6','7') ");
		sql.append("group by rpt_month.m_year,m_month,bank_code,bank_name ");
		/*sql.append("union ");
		sql.append("select CAST(? AS int), CAST(? AS int),wlx01.bank_no,");
		sql.append("      bank_name,");//單位名稱
		sql.append("      credit_staff,");//正式職員.人數     
		sql.append("      skill_staff,");//技工.人數    
		sql.append("      manual_staff,");//工友.人數    
		sql.append("      temp_staff,");//特約人員.人數   
		sql.append("      credit_staff_num ");//總人數         
		sql.append("from (select * from wlx01 where m_year=?)wlx01 ");
		sql.append("left join (select * from bn01 where m_year=?)bn01 on wlx01.bank_no=bn01.bank_no ");
		sql.append("where bn01.bank_type in ('6','7') ");
		sql.append("and bn_type != '2' ");
		paramList.add(syear);
		paramList.add(smonth);
		paramList.add(wlx01_m_year);
		paramList.add(wlx01_m_year);*/
		sql.append(")a ");
		sql.append("where m_year=? ");
		sql.append("and m_month=? ");
		sql.append("order by bank_code ");
		paramList.add(syear);
		paramList.add(smonth);
		
	    List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"bank_code,bank_name,credit_staff,credit_staff_rate,skill_staff,skill_staff_rate,manual_staff,manual_staff_rate,temp_staff,temp_staff_rate,credit_staff_num");
		System.out.println("dbData.size()="+dbData.size());
		return dbData;
	}
	public static List getWLX01Data(String syear,String smonth,String wlx01_m_year){
		StringBuffer sql = new StringBuffer();
		List paramList = new ArrayList();//共同參數
		sql.append("select bank_code,");
		sql.append("      bank_name,");//單位名稱
		sql.append("      credit_staff,");//正式職員.人數
		sql.append("      decode(credit_staff,0,0,round(credit_staff /  credit_staff_num *100 ,2))  as   credit_staff_rate,");//正式職員.比率
		sql.append("      skill_staff,");//技工.人數
		sql.append("      decode(skill_staff,0,0,round(skill_staff /  credit_staff_num *100 ,2))  as  skill_staff_rate,");//技工.比率
		sql.append("      manual_staff,");//工友.人數
		sql.append("      decode(manual_staff,0,0,round(manual_staff /  credit_staff_num *100 ,2))  as   manual_staff_rate,");//工友.比率
		sql.append("      temp_staff,");//特約人員.人數
		sql.append("      decode(temp_staff,0,0,round(temp_staff /  credit_staff_num *100 ,2))  as  temp_staff_rate,");//特約人員.比率
		sql.append("      credit_staff_num ");//總人數
		sql.append("from ( ");
		/*sql.append("select rpt_month.m_year,m_month,bank_code,bank_name,");
		sql.append("        sum(decode(acc_code,'credit_staff',amt,0)) as credit_staff ,");//正式職員.人數
		sql.append("        sum(decode(acc_code,'skill_staff',amt,0)) as  skill_staff,");//技工.人數
		sql.append("        sum(decode(acc_code,'manual_staff',amt,0)) as  manual_staff ,");//工友.人數
		sql.append("        sum(decode(acc_code,'temp_staff',amt,0)) as temp_staff,");//特約人員.人數
		sql.append("        sum(decode(acc_code,'credit_staff_num',amt,0)) as credit_staff_num ");//總人數
		sql.append("from rpt_month left join (select * from bn01 where m_year=?)bn01 on rpt_month.bank_code=bn01.bank_no ");
		sql.append("where report_no=? ");
		paramList.add(wlx01_m_year);
		paramList.add("FR072W");
		sql.append("and rpt_month.bank_type in ('6','7') ");
		sql.append("group by rpt_month.m_year,m_month,bank_code,bank_name ");
		sql.append("union ");*/
		sql.append("select CAST(? AS int)as m_year, CAST(? AS int)as m_month ,wlx01.bank_no as bank_code,");
		sql.append("      bank_name,");//單位名稱
		sql.append("      credit_staff,");//正式職員.人數     
		sql.append("      skill_staff,");//技工.人數    
		sql.append("      manual_staff,");//工友.人數    
		sql.append("      temp_staff,");//特約人員.人數   
		sql.append("      credit_staff_num ");//總人數         
		sql.append("from (select * from wlx01 where m_year=?)wlx01 ");
		sql.append("left join (select * from bn01 where m_year=?)bn01 on wlx01.bank_no=bn01.bank_no ");
		sql.append("where bn01.bank_type in ('6','7') ");
		sql.append("and bn_type != '2' ");
		paramList.add(syear);
		paramList.add(smonth);
		paramList.add(wlx01_m_year);
		paramList.add(wlx01_m_year);
		sql.append(")a ");
		sql.append("where m_year=? ");
		sql.append("and m_month=? ");
		sql.append("order by bank_code ");
		paramList.add(syear);
		paramList.add(smonth);
		
	    List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"bank_code,bank_name,credit_staff,credit_staff_rate,skill_staff,skill_staff_rate,manual_staff,manual_staff_rate,temp_staff,temp_staff_rate,credit_staff_num");
		System.out.println("dbData.size()="+dbData.size());
		return dbData;
	}
}