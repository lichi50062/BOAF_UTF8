/*
104.03.14 add by 2968  
104.03.24 fix 實際資料有2筆,但只顯示1筆資料問題 by 2295        
*/
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

public class RptFR069W {
	public static String createRpt(String s_year,String s_month,String unit) {    
		System.out.println("s_month="+s_month);
		String errMsg="";
		StringBuffer sqlCmd = new StringBuffer();
		String unit_name="";
		int rowNum=0;
		int i=0;
		int j=0;
		String filename="全體農漁會信用部應予評估資產彙總表.xls";
		StringBuffer field = new StringBuffer();
		StringBuffer group = new StringBuffer();
		List dbData1=null;
		List dbData2=null;
		List paramList = new ArrayList();//共同參數
		String cd01_table = "";
        String wlx01_m_year = "";
        //99.09.10 add 查詢年度100年以前.縣市別不同===============================
	    cd01_table = (Integer.parseInt(s_year) < 100)?"cd01_99":""; 
	    wlx01_m_year = (Integer.parseInt(s_year) < 100)?"99":"100"; 
	    //=====================================================================    
		
	    reportUtil reportUtil=new reportUtil();
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
			
			String openfile="全體農漁會信用部應予評估資產彙總表.xls";
			System.out.println("open file "+openfile);
	
			FileInputStream finput=new FileInputStream(xlsDir+System.getProperty("file.separator")+openfile );
	
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
		    ps.setPaperSize( ( short )8 ); //設定紙張大小 A4
			
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
            
			HSSFRow row=null;//宣告一列
			HSSFCell cell=null;//宣告一個儲存格
	        
			//paramList.add(wlx01_m_year);
			//paramList.add(bank_type);
            
			sqlCmd.append("select a10.m_year,m_month,bank_type,");
		    sqlCmd.append(getFieldSQL());
		    for(int k=1;k<=26;k++){
	            paramList.add(unit);
	        }
		    paramList.add(wlx01_m_year);
		    paramList.add(s_year);
			paramList.add(s_month);
		    sqlCmd.append("group by a10.m_year,m_month,bank_type ");
		    sqlCmd.append("union ");
		    sqlCmd.append("select a10.m_year,m_month,'ALL' as bank_type,");
		    sqlCmd.append(getFieldSQL());
		    for(int k=1;k<=26;k++){
	            paramList.add(unit);
	        }
		    paramList.add(wlx01_m_year);
		    paramList.add(s_year);
			paramList.add(s_month);
		    sqlCmd.append("group by a10.m_year,m_month ");
		    dbData1=DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,bank_type,loan1_amt,loan2_amt,loan3_amt,loan4_amt,"+
																    		"loan_sum,invest_sum,other_sum,all_sum,"+
																    		"loan1_baddebt,loan2_baddebt,loan3_baddebt,loan4_baddebt,loan_baddebt_sum,"+
																    		"build1_baddebt,build2_baddebt,build3_baddebt,build4_baddebt,build_baddebt_sum,"+
																    		"above_loan1_amt,above_loan2_amt,above_loan3_amt,above_loan4_amt,above_loan_sum");
			System.out.println("表1~表4的dbData1.size()="+dbData1.size());
			
			sqlCmd.setLength(0);
			paramList.clear();
			sqlCmd.append("select bank_code,bank_name,");
			sqlCmd.append("round(sum(loan1_baddebt)/?,0) as loan1_baddebt,");//--帳列備抵呆帳(實提).放款.第一類
			sqlCmd.append("round(sum(loan2_baddebt)/?,0) as loan2_baddebt,");//--帳列備抵呆帳(實提).放款.第二類
			sqlCmd.append("round(sum(loan3_baddebt)/?,0) as loan3_baddebt,");//--帳列備抵呆帳(實提).放款.第三類
			sqlCmd.append("round(sum(loan4_baddebt)/?,0) as loan4_baddebt,");//--帳列備抵呆帳(實提).放款.第四類
			sqlCmd.append("round(sum(loan1_baddebt+loan2_baddebt+loan3_baddebt+loan4_baddebt)/?,0) as loan_baddebt_sum,");//--帳列備抵呆帳.放款.加總
			sqlCmd.append("round(sum(loan1_amt)*0.01/?,0) as above_loan1_amt,");//--最低標準之備抵呆帳.放款.第一類
			sqlCmd.append("round(sum(loan2_amt)*0.02/?,0) as above_loan2_amt,");//--最低標準之備抵呆帳.放款.第二類
			sqlCmd.append("round(sum(loan3_amt)*0.5/?,0) as above_loan3_amt,");//--最低標準之備抵呆帳.放款.第三類
			sqlCmd.append("round(sum(loan4_amt)/?,0) as above_loan4_amt,");//--最低標準之備抵呆帳.放款.第四類
			sqlCmd.append("round(sum(loan1_amt)*0.01/?,0)+round(sum(loan2_amt)*0.02/?,0)+round(sum(loan3_amt)*0.5/?,0)+round(sum(loan4_amt)/?,0) as above_loan_sum,");//--最低標準之備抵呆帳.放款.合計
			sqlCmd.append("round(sum(baddebt_noenough)/?,0) as baddebt_noenough,");//--提列備抵呆帳不足金額
			sqlCmd.append("round(sum(baddebt_104)/?,0) as baddebt_104,");//--104年底前,提撥金額
			sqlCmd.append("round(sum(baddebt_105)/?,0) as baddebt_105,");//--105年底前,提撥金額
			sqlCmd.append("round(sum(baddebt_106)/?,0) as baddebt_106,");//--106年底前,提撥金額
			sqlCmd.append("round(sum(baddebt_107)/?,0) as baddebt_107,");//--107年底前,提撥金額
			sqlCmd.append("round(sum(baddebt_108)/?,0) as baddebt_108 ");//--108年底前,提撥金額
			sqlCmd.append("from a10 left join (select * from bn01 where m_year=?)bn01 on a10.bank_code=bn01.bank_no ");
			sqlCmd.append("where a10.m_year=? ");
			sqlCmd.append("and m_month=? ");
			sqlCmd.append("and baddebt_flag='N' ");
			sqlCmd.append("group by bank_code,bank_name ");
			for(int k=1;k<=19;k++){
	            paramList.add(unit);
	        }
			paramList.add(wlx01_m_year);
			paramList.add(s_year);
			paramList.add(s_month);
			dbData2=DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"bank_code,bank_name,loan1_baddebt,loan2_baddebt,loan3_baddebt,loan4_baddebt,loan_baddebt_sum,"+
																			"above_loan1_amt,above_loan2_amt,above_loan3_amt,above_loan4_amt,above_loan_sum,"+
																			"baddebt_noenough,baddebt_104,baddebt_105,baddebt_106,baddebt_107,baddebt_108");
			System.out.println("表5的dbData2.size()="+dbData2.size());
			unit_name=Utility.getUnitName(unit);		
			//列印報表名稱
			row=(sheet.getRow(1)==null)? sheet.createRow(1) : sheet.getRow(1);
			cell=row.getCell((short)1);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue(s_year+"年"+s_month+"月底應予評估資產彙總資料");
			row=(sheet.getRow(3)==null)? sheet.createRow(3) : sheet.getRow(3);
			cell=row.getCell((short)1);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("【表1】全體農漁會信用部"+s_year+"年"+s_month+"月底應予評估資產");
			row=(sheet.getRow(4)==null)? sheet.createRow(4) : sheet.getRow(4);
			cell=row.getCell((short)4);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("單位：新台幣　"+unit_name);
			row=(sheet.getRow(3)==null)? sheet.createRow(3) : sheet.getRow(3);
			cell=row.getCell((short)7);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("【表2】全體農漁會信用部"+s_year+"年"+s_month+"月底帳列備抵呆帳(實提)");
			row=(sheet.getRow(4)==null)? sheet.createRow(4) : sheet.getRow(4);
			cell=row.getCell((short)9);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("單位：新台幣　"+unit_name);
			row=(sheet.getRow(3)==null)? sheet.createRow(3) : sheet.getRow(3);
			cell=row.getCell((short)12);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("【表3】全體農漁會信用部"+s_year+"年"+s_month+"月底帳列「建築貸款之備抵呆帳」");
			row=(sheet.getRow(4)==null)? sheet.createRow(4) : sheet.getRow(4);
			cell=row.getCell((short)14);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("單位：新台幣　"+unit_name);
			row=(sheet.getRow(3)==null)? sheet.createRow(3) : sheet.getRow(3);
			cell=row.getCell((short)17);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("【表4】全體農漁會信用部"+s_year+"年"+s_month+"月底依規定應提列最低標準之備抵呆帳(應提)");
			row=(sheet.getRow(4)==null)? sheet.createRow(4) : sheet.getRow(4);
			cell=row.getCell((short)19);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("單位：新台幣　"+unit_name);
			row=(sheet.getRow(15)==null)? sheet.createRow(15) : sheet.getRow(15);
			cell=row.getCell((short)1);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("【表5】"+s_year+"年"+s_month+"月底備抵呆帳提列不足之農漁會信用部");
			row=(sheet.getRow(16)==null)? sheet.createRow(16) : sheet.getRow(16);
			cell=row.getCell((short)20);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("單位：新台幣　"+unit_name);
			
			if (dbData1 !=null && dbData1.size() !=0) {
				for(j=0;j<dbData1.size();j++){
					String bank_type = ((DataObject)dbData1.get(j)).getValue("bank_type")==null?"":(((DataObject)dbData1.get(j)).getValue("bank_type")).toString();
					String loan1_amt = ((DataObject)dbData1.get(j)).getValue("loan1_amt")==null?"0":(((DataObject)dbData1.get(j)).getValue("loan1_amt")).toString();
					String loan2_amt = ((DataObject)dbData1.get(j)).getValue("loan2_amt")==null?"0":(((DataObject)dbData1.get(j)).getValue("loan2_amt")).toString();
					String loan3_amt = ((DataObject)dbData1.get(j)).getValue("loan3_amt")==null?"0":(((DataObject)dbData1.get(j)).getValue("loan3_amt")).toString();
					String loan4_amt = ((DataObject)dbData1.get(j)).getValue("loan4_amt")==null?"0":(((DataObject)dbData1.get(j)).getValue("loan4_amt")).toString();
					String loan_sum = ((DataObject)dbData1.get(j)).getValue("loan_sum")==null?"0":(((DataObject)dbData1.get(j)).getValue("loan_sum")).toString();
					String invest_sum = ((DataObject)dbData1.get(j)).getValue("invest_sum")==null?"0":(((DataObject)dbData1.get(j)).getValue("invest_sum")).toString();
					String other_sum = ((DataObject)dbData1.get(j)).getValue("other_sum")==null?"0":(((DataObject)dbData1.get(j)).getValue("other_sum")).toString();
					String all_sum = ((DataObject)dbData1.get(j)).getValue("all_sum")==null?"0":(((DataObject)dbData1.get(j)).getValue("all_sum")).toString();
					String loan1_baddebt = ((DataObject)dbData1.get(j)).getValue("loan1_baddebt")==null?"0":(((DataObject)dbData1.get(j)).getValue("loan1_baddebt")).toString();
					String loan2_baddebt = ((DataObject)dbData1.get(j)).getValue("loan2_baddebt")==null?"0":(((DataObject)dbData1.get(j)).getValue("loan2_baddebt")).toString();
					String loan3_baddebt = ((DataObject)dbData1.get(j)).getValue("loan3_baddebt")==null?"0":(((DataObject)dbData1.get(j)).getValue("loan3_baddebt")).toString();
					String loan4_baddebt = ((DataObject)dbData1.get(j)).getValue("loan4_baddebt")==null?"0":(((DataObject)dbData1.get(j)).getValue("loan4_baddebt")).toString();
					String loan_baddebt_sum = ((DataObject)dbData1.get(j)).getValue("loan_baddebt_sum")==null?"0":(((DataObject)dbData1.get(j)).getValue("loan_baddebt_sum")).toString();
					String build1_baddebt = ((DataObject)dbData1.get(j)).getValue("build1_baddebt")==null?"0":(((DataObject)dbData1.get(j)).getValue("build1_baddebt")).toString();
					String build2_baddebt = ((DataObject)dbData1.get(j)).getValue("build2_baddebt")==null?"0":(((DataObject)dbData1.get(j)).getValue("build2_baddebt")).toString();
					String build3_baddebt = ((DataObject)dbData1.get(j)).getValue("build3_baddebt")==null?"0":(((DataObject)dbData1.get(j)).getValue("build3_baddebt")).toString();
					String build4_baddebt = ((DataObject)dbData1.get(j)).getValue("build4_baddebt")==null?"0":(((DataObject)dbData1.get(j)).getValue("build4_baddebt")).toString();
					String build_baddebt_sum = ((DataObject)dbData1.get(j)).getValue("build_baddebt_sum")==null?"0":(((DataObject)dbData1.get(j)).getValue("build_baddebt_sum")).toString();
					String above_loan1_amt = ((DataObject)dbData1.get(j)).getValue("above_loan1_amt")==null?"0":(((DataObject)dbData1.get(j)).getValue("above_loan1_amt")).toString();
					String above_loan2_amt = ((DataObject)dbData1.get(j)).getValue("above_loan2_amt")==null?"0":(((DataObject)dbData1.get(j)).getValue("above_loan2_amt")).toString();
					String above_loan3_amt = ((DataObject)dbData1.get(j)).getValue("above_loan3_amt")==null?"0":(((DataObject)dbData1.get(j)).getValue("above_loan3_amt")).toString();
					String above_loan4_amt = ((DataObject)dbData1.get(j)).getValue("above_loan4_amt")==null?"0":(((DataObject)dbData1.get(j)).getValue("above_loan4_amt")).toString();
					String above_loan_sum = ((DataObject)dbData1.get(j)).getValue("above_loan_sum")==null?"0":(((DataObject)dbData1.get(j)).getValue("above_loan_sum")).toString();
					//表1
					rowNum = 6;
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
					if("6".equals(bank_type))cell=row.getCell((short)3);
					else if("7".equals(bank_type))cell=row.getCell((short)4);
					else if("ALL".equals(bank_type))cell=row.getCell((short)5);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(loan1_amt));
					rowNum++;
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
					if("6".equals(bank_type))cell=row.getCell((short)3);
					else if("7".equals(bank_type))cell=row.getCell((short)4);
					else if("ALL".equals(bank_type))cell=row.getCell((short)5);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(loan2_amt));
					rowNum++;
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
					if("6".equals(bank_type))cell=row.getCell((short)3);
					else if("7".equals(bank_type))cell=row.getCell((short)4);
					else if("ALL".equals(bank_type))cell=row.getCell((short)5);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(loan3_amt));
					rowNum++;
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
					if("6".equals(bank_type))cell=row.getCell((short)3);
					else if("7".equals(bank_type))cell=row.getCell((short)4);
					else if("ALL".equals(bank_type))cell=row.getCell((short)5);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(loan4_amt));
					rowNum++;
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
					if("6".equals(bank_type))cell=row.getCell((short)3);
					else if("7".equals(bank_type))cell=row.getCell((short)4);
					else if("ALL".equals(bank_type))cell=row.getCell((short)5);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(loan_sum));
					rowNum++;
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
					if("6".equals(bank_type))cell=row.getCell((short)3);
					else if("7".equals(bank_type))cell=row.getCell((short)4);
					else if("ALL".equals(bank_type))cell=row.getCell((short)5);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(invest_sum));
					rowNum++;
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
					if("6".equals(bank_type))cell=row.getCell((short)3);
					else if("7".equals(bank_type))cell=row.getCell((short)4);
					else if("ALL".equals(bank_type))cell=row.getCell((short)5);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(other_sum));
					rowNum++;
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
					if("6".equals(bank_type))cell=row.getCell((short)3);
					else if("7".equals(bank_type))cell=row.getCell((short)4);
					else if("ALL".equals(bank_type))cell=row.getCell((short)5);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(all_sum));
					//表2
					rowNum = 6;
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
					if("6".equals(bank_type))cell=row.getCell((short)8);
					else if("7".equals(bank_type))cell=row.getCell((short)9);
					else if("ALL".equals(bank_type))cell=row.getCell((short)10);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(loan1_baddebt));
					rowNum++;
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
					if("6".equals(bank_type))cell=row.getCell((short)8);
					else if("7".equals(bank_type))cell=row.getCell((short)9);
					else if("ALL".equals(bank_type))cell=row.getCell((short)10);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(loan2_baddebt));
					rowNum++;
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
					if("6".equals(bank_type))cell=row.getCell((short)8);
					else if("7".equals(bank_type))cell=row.getCell((short)9);
					else if("ALL".equals(bank_type))cell=row.getCell((short)10);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(loan3_baddebt));
					rowNum++;
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
					if("6".equals(bank_type))cell=row.getCell((short)8);
					else if("7".equals(bank_type))cell=row.getCell((short)9);
					else if("ALL".equals(bank_type))cell=row.getCell((short)10);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(loan4_baddebt));
					rowNum++;
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
					if("6".equals(bank_type))cell=row.getCell((short)8);
					else if("7".equals(bank_type))cell=row.getCell((short)9);
					else if("ALL".equals(bank_type))cell=row.getCell((short)10);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(loan_baddebt_sum));
					//表3
					rowNum = 6;
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
					if("6".equals(bank_type))cell=row.getCell((short)13);
					else if("7".equals(bank_type))cell=row.getCell((short)14);
					else if("ALL".equals(bank_type))cell=row.getCell((short)15);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(build1_baddebt));
					rowNum++;
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
					if("6".equals(bank_type))cell=row.getCell((short)13);
					else if("7".equals(bank_type))cell=row.getCell((short)14);
					else if("ALL".equals(bank_type))cell=row.getCell((short)15);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(build2_baddebt));
					rowNum++;
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
					if("6".equals(bank_type))cell=row.getCell((short)13);
					else if("7".equals(bank_type))cell=row.getCell((short)14);
					else if("ALL".equals(bank_type))cell=row.getCell((short)15);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(build3_baddebt));
					rowNum++;
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
					if("6".equals(bank_type))cell=row.getCell((short)13);
					else if("7".equals(bank_type))cell=row.getCell((short)14);
					else if("ALL".equals(bank_type))cell=row.getCell((short)15);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(build4_baddebt));
					rowNum++;
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
					if("6".equals(bank_type))cell=row.getCell((short)13);
					else if("7".equals(bank_type))cell=row.getCell((short)14);
					else if("ALL".equals(bank_type))cell=row.getCell((short)15);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(build_baddebt_sum));
					//表4
					rowNum = 6;
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
					if("6".equals(bank_type))cell=row.getCell((short)18);
					else if("7".equals(bank_type))cell=row.getCell((short)19);
					else if("ALL".equals(bank_type))cell=row.getCell((short)20);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(above_loan1_amt));
					rowNum++;
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
					if("6".equals(bank_type))cell=row.getCell((short)18);
					else if("7".equals(bank_type))cell=row.getCell((short)19);
					else if("ALL".equals(bank_type))cell=row.getCell((short)20);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(above_loan2_amt));
					rowNum++;
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
					if("6".equals(bank_type))cell=row.getCell((short)18);
					else if("7".equals(bank_type))cell=row.getCell((short)19);
					else if("ALL".equals(bank_type))cell=row.getCell((short)20);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(above_loan3_amt));
					rowNum++;
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
					if("6".equals(bank_type))cell=row.getCell((short)18);
					else if("7".equals(bank_type))cell=row.getCell((short)19);
					else if("ALL".equals(bank_type))cell=row.getCell((short)20);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(above_loan4_amt));
					rowNum++;
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
					if("6".equals(bank_type))cell=row.getCell((short)18);
					else if("7".equals(bank_type))cell=row.getCell((short)19);
					else if("ALL".equals(bank_type))cell=row.getCell((short)20);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(above_loan_sum));
				}
			}
			
			if (dbData2 !=null && dbData2.size() !=0) {
				rowNum=20;
				short k=1;
				for(j=0;j<dbData2.size();j++){				   
					String bank_name = ((DataObject)dbData2.get(j)).getValue("bank_name")==null?"":(((DataObject)dbData2.get(j)).getValue("bank_name")).toString();
					String loan1_baddebt = ((DataObject)dbData2.get(j)).getValue("loan1_baddebt")==null?"0":(((DataObject)dbData2.get(j)).getValue("loan1_baddebt")).toString();
					String loan2_baddebt = ((DataObject)dbData2.get(j)).getValue("loan2_baddebt")==null?"0":(((DataObject)dbData2.get(j)).getValue("loan2_baddebt")).toString();
					String loan3_baddebt = ((DataObject)dbData2.get(j)).getValue("loan3_baddebt")==null?"0":(((DataObject)dbData2.get(j)).getValue("loan3_baddebt")).toString();
					String loan4_baddebt = ((DataObject)dbData2.get(j)).getValue("loan4_baddebt")==null?"0":(((DataObject)dbData2.get(j)).getValue("loan4_baddebt")).toString();
					String loan_baddebt_sum = ((DataObject)dbData2.get(j)).getValue("loan_baddebt_sum")==null?"0":(((DataObject)dbData2.get(j)).getValue("loan_baddebt_sum")).toString();					
					String above_loan1_amt = ((DataObject)dbData2.get(j)).getValue("above_loan1_amt")==null?"0":(((DataObject)dbData2.get(j)).getValue("above_loan1_amt")).toString();
					String above_loan2_amt = ((DataObject)dbData2.get(j)).getValue("above_loan2_amt")==null?"0":(((DataObject)dbData2.get(j)).getValue("above_loan2_amt")).toString();
					String above_loan3_amt = ((DataObject)dbData2.get(j)).getValue("above_loan3_amt")==null?"0":(((DataObject)dbData2.get(j)).getValue("above_loan3_amt")).toString();
					String above_loan4_amt = ((DataObject)dbData2.get(j)).getValue("above_loan4_amt")==null?"0":(((DataObject)dbData2.get(j)).getValue("above_loan4_amt")).toString();
					String above_loan_sum = ((DataObject)dbData2.get(j)).getValue("above_loan_sum")==null?"0":(((DataObject)dbData2.get(j)).getValue("above_loan_sum")).toString();					
					String baddebt_noenough = ((DataObject)dbData2.get(j)).getValue("baddebt_noenough")==null?"0":(((DataObject)dbData2.get(j)).getValue("baddebt_noenough")).toString();
					String baddebt_104 = ((DataObject)dbData2.get(j)).getValue("baddebt_104")==null?"0":(((DataObject)dbData2.get(j)).getValue("baddebt_104")).toString();
					String baddebt_105 = ((DataObject)dbData2.get(j)).getValue("baddebt_105")==null?"0":(((DataObject)dbData2.get(j)).getValue("baddebt_105")).toString();
					String baddebt_106 = ((DataObject)dbData2.get(j)).getValue("baddebt_106")==null?"0":(((DataObject)dbData2.get(j)).getValue("baddebt_106")).toString();
					String baddebt_107 = ((DataObject)dbData2.get(j)).getValue("baddebt_107")==null?"0":(((DataObject)dbData2.get(j)).getValue("baddebt_107")).toString();
					String baddebt_108 = ((DataObject)dbData2.get(j)).getValue("baddebt_108")==null?"0":(((DataObject)dbData2.get(j)).getValue("baddebt_108")).toString();
					
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
					for(int l=0;l<23;l++){
		                cell = row.createCell( (short)l);
		            }
					
					cell=row.getCell((short)1);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs);
					cell.setCellValue(k);
					
					cell=row.getCell((short)2);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs1);
					cell.setCellValue(bank_name);
					cell=row.getCell((short)3);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellStyle(cs1);
					sheet.addMergedRegion( new Region( ( short )rowNum, ( short )2,
                              ( short )rowNum,
                              ( short )3 ) );            
					
					cell=row.getCell((short)4);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(loan1_baddebt));
					
					cell=row.getCell((short)5);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(loan2_baddebt));
					
					cell=row.getCell((short)6);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(loan3_baddebt));
					cell=row.getCell((short)7);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellStyle(cs2);
					sheet.addMergedRegion( new Region( ( short )rowNum, ( short )6,
                            ( short )rowNum,
                            ( short )7 ) );            
					
					
					cell=row.getCell((short)8);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(loan4_baddebt));
					
					cell=row.getCell((short)9);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(loan_baddebt_sum));
					
					cell=row.getCell((short)10);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(above_loan1_amt));
					
					cell=row.getCell((short)11);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(above_loan2_amt));
					cell=row.getCell((short)12);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellStyle(cs2);
					sheet.addMergedRegion( new Region( ( short )rowNum, ( short )11,
                            ( short )rowNum,
                            ( short )12 ) );            
					
					cell=row.getCell((short)13);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(above_loan3_amt));
					
					cell=row.getCell((short)14);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(above_loan4_amt));
					
					cell=row.getCell((short)15);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(above_loan_sum));
					
					cell=row.getCell((short)16);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(baddebt_noenough));
					cell=row.getCell((short)17);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellStyle(cs2);
					sheet.addMergedRegion( new Region( ( short )rowNum, ( short )16,
                            ( short )rowNum,
                            ( short )17 ) );            
					
					cell=row.getCell((short)18);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(baddebt_104));
					
					cell=row.getCell((short)19);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(baddebt_105));
					
					cell=row.getCell((short)20);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(baddebt_106));
					
					cell=row.getCell((short)21);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(baddebt_107));
					
					cell=row.getCell((short)22);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(baddebt_108));
					
					rowNum++;
					k++;
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

	public static String getFieldSQL() {
		StringBuffer sql = new StringBuffer();
			sql.append("round(sum(loan1_amt)/?,0) as loan1_amt,");//--放款.第一類-表1
			sql.append("round(sum(loan2_amt)/?,0) as loan2_amt,");//--放款.第二類-表1
			sql.append("round(sum(loan3_amt)/?,0) as loan3_amt,");//--放款.第三類-表1
			sql.append("round(sum(loan4_amt)/?,0) as loan4_amt,");//--放款.第四類-表1
			sql.append("round(sum(loan1_amt+loan2_amt+loan3_amt+loan4_amt)/?,0) as loan_sum,");//--放款.加總-表1
			sql.append("round(sum(invest1_amt+invest2_amt+invest3_amt+invest4_amt)/?,0) as invest_sum,");//--投資.加總-表1
			sql.append("round(sum(other1_amt+other2_amt+other3_amt+other4_amt)/?,0) as other_sum,");//--其他.加總-表1
			sql.append("round(sum(loan1_amt+loan2_amt+loan3_amt+loan4_amt+invest1_amt+invest2_amt+invest3_amt+invest4_amt+other1_amt+other2_amt+other3_amt+other4_amt)/?,0) as all_sum,");//--合計.加總-表1
			sql.append("round(sum(loan1_baddebt)/?,0) as loan1_baddebt,");//--帳列備抵呆帳(實提).放款.第一類-表2
			sql.append("round(sum(loan2_baddebt)/?,0) as loan2_baddebt,");//--帳列備抵呆帳(實提).放款.第二類-表2
			sql.append("round(sum(loan3_baddebt)/?,0) as loan3_baddebt,");//--帳列備抵呆帳(實提).放款.第三類-表2
			sql.append("round(sum(loan4_baddebt)/?,0) as loan4_baddebt,");//--帳列備抵呆帳(實提).放款.第四類-表2
			sql.append("round(sum(loan1_baddebt+loan2_baddebt+loan3_baddebt+loan4_baddebt)/?,0) as loan_baddebt_sum,");//--帳列備抵呆帳.放款.加總-表2
			sql.append("round(sum(build1_baddebt)/?,0) as build1_baddebt,");//--帳列備抵呆帳(實提).建築放款.第一類-表3
			sql.append("round(sum(build2_baddebt)/?,0) as build2_baddebt,");//--帳列備抵呆帳(實提).建築放款.第二類-表3
			sql.append("round(sum(build3_baddebt)/?,0) as build3_baddebt,");//--帳列備抵呆帳(實提).建築放款.第三類-表3
			sql.append("round(sum(build4_baddebt)/?,0) as build4_baddebt,");//--帳列備抵呆帳(實提).建築放款.第四類-表3
			sql.append("round(sum(build1_baddebt+build2_baddebt+build3_baddebt+build4_baddebt)/?,0) as build_baddebt_sum,");//--帳列備抵呆帳.建築放款.加總-表3
			sql.append("round(sum(loan1_amt)*0.01/?,0) as above_loan1_amt,");//--最低標準之備抵呆帳.放款.第一類-表4
			sql.append("round(sum(loan2_amt)*0.02/?,0) as above_loan2_amt,");//--最低標準之備抵呆帳.放款.第二類-表4
			sql.append("round(sum(loan3_amt)*0.5/?,0) as above_loan3_amt,");//--最低標準之備抵呆帳.放款.第三類-表4
			sql.append("round(sum(loan4_amt)/?,0) as above_loan4_amt,");//--最低標準之備抵呆帳.放款.第四類-表4
			sql.append("round(sum(loan1_amt)*0.01/?,0)+round(sum(loan2_amt)*0.02/?,0)+round(sum(loan3_amt)*0.5/?,0)+round(sum(loan4_amt)/?,0) as above_loan_sum ");//--最低標準之備抵呆帳.放款.合計-表4
			sql.append("from a10 left join (select * from bn01 where m_year=?)bn01 on a10.bank_code=bn01.bank_no ");
			sql.append("where a10.m_year=? ");
			sql.append("and m_month=? ");
			sql.append("and bank_type in ('6','7') ");
		return sql.toString();
	}
}