/*
    101.07 created by 2968  	
    101.11.14 add 列印金額單位    by 2295		   
    101.12.18 當年度累計的保證餘額,修改成讀取總累計的保證餘額 by 2295   
    106.10.16 add [M]台中商業銀行 by 2295
*/
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.Region;

import java.io.*;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR062W {
    public static String createRpt(String s_year,String s_month,String unit,String bank_type){
		String errMsg = "";
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList(); 
		int rowNum=0;
		int j=0;
		String filename="農業信用保證基金業務統計.xls";
		//===============================
	    String u_year = (Integer.parseInt(s_year) < 100)?"99":"100"; 
	    Calendar now = Calendar.getInstance();
	    String now_YEAR  = String.valueOf(now.get(Calendar.YEAR)-1911); //回覆值為西元年故需-1911取得民國年;
	    //=====================================================================    
	   
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
    		
    		String openfile="";
    		
    		 if(Integer.parseInt(s_year) *100 +  Integer.parseInt(s_month) >= 10605 ){//106.10.16 add
    		     openfile = "農業信用保證基金業務統計.xls";
             }else{
                 openfile = "農業信用保證基金業務統計_10605.xls";
             }
    		
    		 
    		System.out.println("open file "+openfile);
			FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+ openfile );

	  	    //設定FileINputStream讀取Excel檔
			POIFSFileSystem fs = new POIFSFileSystem(finput); //讀取檔案
            HSSFWorkbook wb = new HSSFWorkbook(fs);//建立活頁簿
            HSSFSheet sheet = wb.getSheetAt(0);// 讀取第一個工作表，宣告其為sheet
            HSSFSheet sheet1 = wb.getSheetAt(1);
            HSSFPrintSetup ps = sheet.getPrintSetup(); // 取得設定
	        //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
	        //sheet.setAutobreaks(true); //自動分頁

	  		//設定頁面符合列印大小
	        sheet.setAutobreaks( false );
	        //ps.setScale( ( short )64 ); //列印縮放百分比
	        ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
	  		//wb.setSheetName(0,"test");
	  		finput.close();
	  		HSSFRow row=null;//宣告一列
	  		HSSFCell cell=null;//宣告一個儲存格
	  		
	  		reportUtil reportUtil = new reportUtil();
            HSSFCellStyle cs_right = reportUtil.getRightStyle(wb);
            HSSFCellStyle cs_center = reportUtil.getDefaultStyle(wb);
            HSSFCellStyle nb_left = reportUtil.getNoBorderLeftStyle(wb);
            HSSFCellStyle nb_center = reportUtil.getNoBorderDefaultStyle(wb);
            HSSFCellStyle nb_right = reportUtil.getNoBoderStyle(wb);//無框置右 
            
	  		//一、保證項目別
	  		sqlCmd.append(" select m201.guarantee_item_no,");
	  		sqlCmd.append("        cdshareno.cmuse_name,");
	  		sqlCmd.append("        m201.guarantee_cnt_year,");
	  		sqlCmd.append("        round(m201.loan_amt_year / ?,0)  as loan_amt_year,");
	  		sqlCmd.append("        round(m201.guarantee_amt_year / ?,0)  as guarantee_amt_year,");
	  		sqlCmd.append("        round(m201.guarantee_bal_sum / ?,0)  as guarantee_bal_sum,");//101.12.18 修改讀取總累計的保證餘額
	  		sqlCmd.append("        decode(m201_sum.guarantee_bal_sum,0,0,round(m201.guarantee_bal_sum /  m201_sum.guarantee_bal_sum *100 ,2))  as   guarantee_bal_rate,");
	  		sqlCmd.append("        m201.guarantee_cnt_sum,");
	  		sqlCmd.append("        round(m201.loan_amt_sum / ?,0)  as loan_amt_sum,");
	  		sqlCmd.append("        round(m201.guarantee_amt_sum / ?,0)  as guarantee_amt_sum ");
	  		paramList.add(unit);
            paramList.add(unit);
            paramList.add(unit);
            paramList.add(unit);
            paramList.add(unit);
	  		sqlCmd.append(" from M201 left join (select * from cdshareno where cmuse_div='040')cdshareno on M201.GUARANTEE_ITEM_NO = CDSHARENO.CMUSE_ID,");
            sqlCmd.append("  (select * from m201 where m_year= ? and m_month= ? and guarantee_item_no='0')m201_sum ");
            sqlCmd.append("  where m201.m_year = ? and m201.m_month= ? ");
            sqlCmd.append("  and m201.m_year = m201_sum.m_year and m201.m_month =  m201_sum.m_month ");
            sqlCmd.append("  order by to_number(cdshareno.output_order) ");
            paramList.add(s_year);
            paramList.add(s_month);
            paramList.add(s_year);
            paramList.add(s_month);
            List qList1 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"guarantee_item_no,cmuse_name,guarantee_cnt_year,loan_amt_year,guarantee_amt_year,guarantee_bal_sum,guarantee_bal_rate,guarantee_cnt_sum,loan_amt_sum,guarantee_amt_sum");
			System.out.println("一、保證項目別 qList1.size()--"+qList1.size());
            
			//二、保證貸款用途別
            sqlCmd = new StringBuffer();
            paramList = new ArrayList();
            sqlCmd.append(" select loan_use_no,");
            sqlCmd.append("        cdshareno.cmuse_name,");
            sqlCmd.append("        guarantee_cnt_year,");
            sqlCmd.append("        round(m106.loan_amt_year / ?,0)  as loan_amt_year,");
            sqlCmd.append("        round(m106.guarantee_amt_year / ?,0)  as guarantee_amt_year,");
            sqlCmd.append("        guarantee_cnt_sum,");
            sqlCmd.append("        round(m106.loan_amt_sum / ?,0)  as loan_amt_sum,");
            sqlCmd.append("        round(m106.guarantee_amt_sum / ?,0)  as guarantee_amt_sum ");
            sqlCmd.append("        from m106 ");
            sqlCmd.append("        left join (select * from cdshareno where cmuse_div='039')cdshareno on M106.LOAN_USE_NO = CDSHARENO.CMUSE_ID ");
            paramList.add(unit);
            paramList.add(unit);
            paramList.add(unit);
            paramList.add(unit);
            sqlCmd.append("  where m_year = ? and m_month= ? ");
            sqlCmd.append("  order by to_number(cdshareno.output_order) ");
            paramList.add(s_year);
            paramList.add(s_month);
            List qList2 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"loan_use_no,cmuse_name,guarantee_cnt_year,loan_amt_year,guarantee_amt_year,guarantee_cnt_sum,loan_amt_sum,guarantee_amt_sum");
            System.out.println("二、保證貸款用途別 qList2.size()--"+qList2.size());
            
            //三、送保農業金融機構合計
            sqlCmd = new StringBuffer();
            paramList = new ArrayList();
            sqlCmd.append(" select m206.loan_unit,M206_LOAN_UNIT.out_loan_name,");
            sqlCmd.append("        m206.guarantee_cnt_year,");
            sqlCmd.append("        round(m206.loan_amt_year / ?,0)  as loan_amt_year,");
            sqlCmd.append("        round(m206.guarantee_amt_year / ?,0)  as guarantee_amt_year,");
            sqlCmd.append("        round(m206.guarantee_bal_year / ?,0)  as guarantee_bal_year,");
            sqlCmd.append("        decode(m206_sum.guarantee_bal_year,0,0,round(m206.guarantee_bal_year /  m206_sum.guarantee_bal_year *100 ,2))  as  guarantee_bal_rate,");
            sqlCmd.append("        m206.guarantee_cnt_sum,");
            sqlCmd.append("        round(m206.loan_amt_sum / ?,0)  as loan_amt_sum,");
            sqlCmd.append("        round(m206.guarantee_amt_sum /?,0)  as guarantee_amt_sum ");
            sqlCmd.append("        from m206 ");
            paramList.add(unit);
            paramList.add(unit);
            paramList.add(unit);
            paramList.add(unit);
            paramList.add(unit);
            sqlCmd.append(" left join m206_loan_unit on M206.LOAN_UNIT = M206_LOAN_UNIT.LOAN_UNIT_NO,");
            sqlCmd.append("  (select * from m206 where m_year= ? and m_month= ? and loan_unit='0')m206_sum , ");
            sqlCmd.append("  (select m_year,m_month,count(*) from m206_loan_unit ");
            if(Integer.parseInt(s_year) *100 +  Integer.parseInt(s_month) >= 10605 ){//106.10.16 add
                sqlCmd.append(" where m_year=? and m_month >= ?");
            }
            sqlCmd.append("  group by m_year,m_month)loan_unit_year ");
            sqlCmd.append(" where m206.m_year*100+m206.m_month >= loan_unit_year.m_year *100+loan_unit_year.m_month");
            sqlCmd.append(" and m206.m_year= ? and m206.m_month= ? ");
            sqlCmd.append(" and m206.m_year = m206_sum.m_year and m206.m_month =  m206_sum.m_month");
            sqlCmd.append(" and loan_unit_year.m_year = m206_loan_unit.m_year");
            sqlCmd.append(" and loan_unit_year.m_month = m206_loan_unit.m_month");
            sqlCmd.append(" order by to_number(m206_loan_unit.output_order)");
            paramList.add(s_year);
            paramList.add(s_month);
            if(Integer.parseInt(s_year) *100 +  Integer.parseInt(s_month) >= 10605 ){//106.05增加台中商銀
                paramList.add("106");
                paramList.add("05");
            }
            paramList.add(s_year);
            paramList.add(s_month);
            List qList3 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"loan_unit,out_loan_name,guarantee_cnt_year,loan_amt_year,guarantee_amt_year,guarantee_bal_year,guarantee_bal_rate,guarantee_cnt_sum,loan_amt_sum,guarantee_amt_sum");
            System.out.println("三、保證貸款用途別 qList3.size()--"+qList3.size());
            
            //title日期
            String getData=(((qList1 == null || qList1.size() ==0)&&(qList2 == null || qList2.size() ==0)&&(qList3 == null || qList3.size() ==0))?"無資料存在":"");
            row=(sheet.getRow(3)==null)? sheet.createRow(3) : sheet.getRow(3);
            cell=row.getCell((short)0);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellValue("中華民國"+s_year+"年"+s_month+"月"+getData);
            row=(sheet.getRow(32)==null)? sheet.createRow(32) : sheet.getRow(32);
            cell=row.getCell((short)0);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellValue("中華民國"+s_year+"年"+s_month+"月"+getData);
            String unit_name = "("+Utility.getUnitName(unit)+")";//取得單位名稱
            //101.11.14 add 列印金額單位    by 2295     
            row=(sheet.getRow(6)==null)? sheet.createRow(6) : sheet.getRow(6);           
            for(int i=3;i<=9;i++){
                //System.out.println("i="+i);
                if(i==6 || i==7) continue;
                cell=row.getCell((short)i);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                cell.setCellValue(unit_name);                
            }
            if (qList1.size() != 0){
                rowNum = 8;
                for(j=0;j<qList1.size();j++){
                    DataObject bean = (DataObject) qList1.get(j);
                    row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
                    String guarantee_item_no=((bean.getValue("guarantee_item_no") == null)?"":bean.getValue("guarantee_item_no")).toString();
                    String cmuse_name=((bean.getValue("cmuse_name") == null)?"":bean.getValue("cmuse_name")).toString();
                    String guarantee_cnt_year=((bean.getValue("guarantee_cnt_year") == null)?"":bean.getValue("guarantee_cnt_year")).toString();
                    String loan_amt_year = ((bean.getValue("loan_amt_year") == null)?"":bean.getValue("loan_amt_year")).toString();
                    String guarantee_amt_year = ((bean.getValue("guarantee_amt_year") == null)?"":bean.getValue("guarantee_amt_year")).toString();
                    String guarantee_bal_sum = ((bean.getValue("guarantee_bal_sum") == null)?"":bean.getValue("guarantee_bal_sum")).toString();//101.12.18修改成讀取總累計的保證餘額
                    String guarantee_bal_rate = ((bean.getValue("guarantee_bal_rate") == null)?"":bean.getValue("guarantee_bal_rate")).toString();
                    String guarantee_cnt_sum=((bean.getValue("guarantee_cnt_sum") == null)?"":bean.getValue("guarantee_cnt_sum")).toString();
                    String loan_amt_sum = ((bean.getValue("loan_amt_sum") == null)?"":bean.getValue("loan_amt_sum")).toString();
                    String guarantee_amt_sum = ((bean.getValue("guarantee_amt_sum") == null)?"":bean.getValue("guarantee_amt_sum")).toString();
                    System.out.println(cmuse_name+", "+guarantee_cnt_year+", "+loan_amt_year+", "+guarantee_amt_year+", "+guarantee_bal_sum+", "+guarantee_bal_rate+", "+guarantee_cnt_sum+", "+loan_amt_sum+", "+guarantee_amt_sum);
                    
                    //保證項目別
                    cell=row.getCell((short)0);                    
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    if("合計".equals(cmuse_name)){
                        cell.setCellValue("一、保證項目別"+cmuse_name);
                   // }else{
                   //     cell.setCellValue(getChNumber(j)+cmuse_name);
                    }
                    //當年度保證件數
                    cell=row.getCell((short)2);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    if("0".equals(guarantee_cnt_year)){
                        cell.setCellValue("-");
                    }else{
                        cell.setCellValue(Utility.setCommaFormat(guarantee_cnt_year));
                    }
                    //當年度融資金額(百萬元)
                    cell=row.getCell((short)3);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    if("0".equals(loan_amt_year)){
                        cell.setCellValue("-");
                    }else{
                        cell.setCellValue(Utility.setCommaFormat(loan_amt_year));
                    }
                    //當年度保證金額(百萬元)
                    cell=row.getCell((short)4);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    if("0".equals(guarantee_amt_year)){
                        cell.setCellValue("-");
                    }else{
                        cell.setCellValue(Utility.setCommaFormat(guarantee_amt_year));
                    }
                    //當年度保證餘額(百萬元)-->101.12.18修改成讀取總累計的保證餘額
                    cell=row.getCell((short)5);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    if("0".equals(guarantee_bal_sum)){
                        cell.setCellValue("-");
                    }else{
                        cell.setCellValue(Utility.setCommaFormat(guarantee_bal_sum));
                    }
                    //當年度保證餘額-結構比
                    cell=row.getCell((short)6);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    if("0".equals(guarantee_bal_rate)){
                        cell.setCellValue("-");
                    }else{
                        cell.setCellValue(getPercentNumber_3(guarantee_bal_rate));
                    }
                    //總累計保證件數
                    cell=row.getCell((short)7);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    if("0".equals(guarantee_cnt_sum)){
                        cell.setCellValue("-");
                    }else{
                        cell.setCellValue(Utility.setCommaFormat(guarantee_cnt_sum));
                    }
                    //總累計融資金額(百萬元)
                    cell=row.getCell((short)8);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    if("0".equals(loan_amt_sum)){
                        cell.setCellValue("-");
                    }else{
                        cell.setCellValue(Utility.setCommaFormat(loan_amt_sum));
                    }
                    //總累計保證金額(百萬元)
                    cell=row.getCell((short)9);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    if("0".equals(guarantee_amt_sum)){
                        cell.setCellValue("-");
                    }else{
                        cell.setCellValue(Utility.setCommaFormat(guarantee_amt_sum));
                    }
                    
                    rowNum++;
                    if(qList2.size() != 0 && j==0){
                        row=sheet.getRow(16);
                        //保證貸款用途別合計-保證餘額
                        cell=row.getCell((short)5);
                        cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                        if("0".equals(guarantee_bal_sum)){
                            cell.setCellValue("-");
                        }else{
                            cell.setCellValue(Utility.setCommaFormat(guarantee_bal_sum));
                        }
                        //保證貸款用途別合計-結構比
                        cell=row.getCell((short)6);
                        cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                        if("0".equals(guarantee_bal_rate)){
                            cell.setCellValue("-");
                        }else{
                            cell.setCellValue(getPercentNumber_3(guarantee_bal_rate));
                        }
                    }
                }
            }
            if (qList2.size() != 0){
                rowNum = 16;
                for(j=0;j<qList2.size();j++){
                    DataObject bean = (DataObject) qList2.get(j);
                    row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
                    String loan_use_no=((bean.getValue("loan_use_no") == null)?"":bean.getValue("loan_use_no")).toString();
                    String cmuse_name=((bean.getValue("cmuse_name") == null)?"":bean.getValue("cmuse_name")).toString();
                    String guarantee_cnt_year=((bean.getValue("guarantee_cnt_year") == null)?"0":bean.getValue("guarantee_cnt_year")).toString();
                    String loan_amt_year = ((bean.getValue("loan_amt_year") == null)?"":bean.getValue("loan_amt_year")).toString();
                    String guarantee_amt_year = ((bean.getValue("guarantee_amt_year") == null)?"0":bean.getValue("guarantee_amt_year")).toString();
                    String guarantee_cnt_sum=((bean.getValue("guarantee_cnt_sum") == null)?"0":bean.getValue("guarantee_cnt_sum")).toString();
                    String loan_amt_sum = ((bean.getValue("loan_amt_sum") == null)?"0":bean.getValue("loan_amt_sum")).toString();
                    String guarantee_amt_sum = ((bean.getValue("guarantee_amt_sum") == null)?"0":bean.getValue("guarantee_amt_sum")).toString();
                    System.out.println(cmuse_name+", "+guarantee_cnt_year+", "+loan_amt_year+", "+guarantee_amt_year+", "+guarantee_cnt_sum+", "+loan_amt_sum+", "+guarantee_amt_sum);
                    
                    //保證貸款用途別
                    cell=row.getCell((short)0);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    if("合計".equals(cmuse_name)){
                        cell.setCellValue("二、保證貸款用途別"+cmuse_name);
                   // }else{
                   //     cell.setCellValue(getChNumber(j)+cmuse_name);
                    }
                    
                    //當年度件數
                    cell=row.getCell((short)2);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    if("0".equals(guarantee_cnt_year)){
                        cell.setCellValue("-");
                    }else{
                        cell.setCellValue(Utility.setCommaFormat(guarantee_cnt_year));
                    }
                    //當年度融資金額
                    cell=row.getCell((short)3);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    if("0".equals(loan_amt_year)){
                        cell.setCellValue("-");
                    }else{
                        cell.setCellValue(Utility.setCommaFormat(loan_amt_year));
                    }
                    //當年度保證金額
                    cell=row.getCell((short)4);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    if("0".equals(guarantee_amt_year)){
                        cell.setCellValue("-");
                    }else{
                        cell.setCellValue(Utility.setCommaFormat(guarantee_amt_year));
                    }
                    //總累計件數
                    cell=row.getCell((short)7);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    if("0".equals(guarantee_cnt_sum)){
                        cell.setCellValue("-");
                    }else{
                        cell.setCellValue(Utility.setCommaFormat(guarantee_cnt_sum));
                    }
                    //總累計融資金額
                    cell=row.getCell((short)8);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    if("0".equals(loan_amt_sum)){
                        cell.setCellValue("-");
                    }else{
                        cell.setCellValue(Utility.setCommaFormat(loan_amt_sum));
                    }
                    //總累計保證金額
                    cell=row.getCell((short)9);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    if("0".equals(guarantee_amt_sum)){
                        cell.setCellValue("-");
                    }else{
                        cell.setCellValue(Utility.setCommaFormat(guarantee_amt_sum));
                    }
                    rowNum++;
                }
            }
            
            //101.11.14 add 列印金額單位    by 2295
            row=(sheet.getRow(35)==null)? sheet.createRow(35) : sheet.getRow(35);
            for(int i=3;i<=9;i++){
                if(i==6 || i==7) continue;
                cell=row.getCell((short)i);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                cell.setCellValue(unit_name);               
            }
           
            if (qList3.size() != 0){
                rowNum = 37;
                for(j=0;j<qList3.size();j++){
                    DataObject bean = (DataObject) qList3.get(j);
                    System.out.println("j="+j);
                    
                    row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
                    String loan_unit=((bean.getValue("loan_unit") == null)?"":bean.getValue("loan_unit")).toString();
                    String out_loan_name=((bean.getValue("out_loan_name") == null)?"":bean.getValue("out_loan_name")).toString();
                    String guarantee_cnt_year=((bean.getValue("guarantee_cnt_year") == null)?"":bean.getValue("guarantee_cnt_year")).toString();
                    String loan_amt_year = ((bean.getValue("loan_amt_year") == null)?"":bean.getValue("loan_amt_year")).toString();
                    String guarantee_amt_year = ((bean.getValue("guarantee_amt_year") == null)?"":bean.getValue("guarantee_amt_year")).toString();
                    String guarantee_bal_year = ((bean.getValue("guarantee_bal_year") == null)?"":bean.getValue("guarantee_bal_year")).toString();
                    String guarantee_bal_rate = ((bean.getValue("guarantee_bal_rate") == null)?"":bean.getValue("guarantee_bal_rate")).toString();
                    String guarantee_cnt_sum=((bean.getValue("guarantee_cnt_sum") == null)?"":bean.getValue("guarantee_cnt_sum")).toString();
                    String loan_amt_sum = ((bean.getValue("loan_amt_sum") == null)?"":bean.getValue("loan_amt_sum")).toString();
                    String guarantee_amt_sum = ((bean.getValue("guarantee_amt_sum") == null)?"":bean.getValue("guarantee_amt_sum")).toString();
                    System.out.println(out_loan_name+", "+guarantee_cnt_year+", "+loan_amt_year+", "+guarantee_amt_year+", "+guarantee_bal_year+", "+guarantee_bal_rate+", "+guarantee_cnt_sum+", "+loan_amt_sum+", "+guarantee_amt_sum);
                    //貸款機構別
                    cell=row.getCell((short)0);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    if("合計".equals(out_loan_name)){
                        cell.setCellValue("三、送保農業金融機構"+out_loan_name);                        
                    }else{
                       if(out_loan_name.length()>=16){                          
                           String str1=out_loan_name.substring(0, 17);
                           String str2=out_loan_name.substring(17, out_loan_name.length());
                           out_loan_name=str1+"\n"+"　"+str2;
                           row.setHeight((short)(2*(short)0x120));
                           cell.setCellValue("　"+out_loan_name);                          
                       }else{                          
                           cell.setCellValue("　"+out_loan_name); 
                       }
                       
                        
                    }
                    //當年度保證案件件數
                    cell=row.getCell((short)2);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    if("0".equals(guarantee_cnt_year)){
                        cell.setCellValue("-");
                    }else{
                        cell.setCellValue(Utility.setCommaFormat(guarantee_cnt_year));
                    }
                   
                    //當年度融資金額(百萬元)
                    cell=row.getCell((short)3);                    
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);                   
                    if("0".equals(loan_amt_year)){                       
                        cell.setCellValue("-");
                    }else{                       
                        cell.setCellValue(Utility.setCommaFormat(loan_amt_year));
                    }
                    
                    //當年度保證金額(百萬元)
                    cell=row.getCell((short)4);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    if("0".equals(guarantee_amt_year)){
                        cell.setCellValue("-");
                    }else{
                        cell.setCellValue(Utility.setCommaFormat(guarantee_amt_year));
                    }
                    
                    //當年度保證餘額(百萬元)
                    cell=row.getCell((short)5);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    if("0".equals(guarantee_bal_year)){
                        cell.setCellValue("-");
                    }else{
                        cell.setCellValue(Utility.setCommaFormat(guarantee_bal_year));
                    }
                    
                    //當年度保證餘額-結構比
                    cell=row.getCell((short)6);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    if("0".equals(guarantee_bal_rate)){
                        cell.setCellValue("-");
                    }else{
                        cell.setCellValue(getPercentNumber_3(guarantee_bal_rate));
                    }
                    
                    //總累計保證件數
                    cell=row.getCell((short)7);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    if("0".equals(guarantee_cnt_sum)){
                        cell.setCellValue("-");
                    }else{
                        cell.setCellValue(Utility.setCommaFormat(guarantee_cnt_sum));
                    }
                    
                    //總累計融資金額(百萬元)
                    cell=row.getCell((short)8);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    if("0".equals(loan_amt_sum)){
                        cell.setCellValue("-");
                    }else{
                        cell.setCellValue(Utility.setCommaFormat(loan_amt_sum));
                    }
                    
                    //總累計保證金額(百萬元)
                    cell=row.getCell((short)9);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    if("0".equals(guarantee_amt_sum)){
                        cell.setCellValue("-");
                    }else{
                        cell.setCellValue(Utility.setCommaFormat(guarantee_amt_sum));
                    }
                    
                    rowNum++;
                }
            }
           
            // 建表結束--------------------------------------
	        HSSFFooter footer = sheet.getFooter();
	        footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));

	        FileOutputStream fout=new FileOutputStream(reportDir + System.getProperty("file.separator")+filename);
	        wb.write(fout);
	        //儲存
	        fout.close();
	        System.out.println("儲存完成");
		}catch(Exception e){
			System.out.println("FR062W.createRpt Error:"+e+e.getMessage());
		}
		return errMsg;
	}
    
    /*
    public static String getChNumber(int n){
        String chNumber="";
        try{
            if (n==1){
                chNumber="(一)";
            }else if (n==2){                 
                chNumber="(二)";
            }else if (n==3){
                chNumber="(三)";
            }else if (n==4){
                chNumber="(四)";
            }else if (n==5){
                chNumber="(五)";
            }else if (n==6){
                chNumber="(六)";
            }else if (n==7){                 
                chNumber="(七)";
            }else if (n==8){
                chNumber="(八)";
            }else if (n==9){
                chNumber="(九)";
            }else if (n==10){
                chNumber="(十)";
            }else if (n==11){
                chNumber="(十一)";
            }else if (n==12){
                chNumber="(十二)";
            }else if (n==13){
                chNumber="(十三)";
            }else if (n==14){
                chNumber="(十四)";
            }else if (n==15){
                chNumber="(十五)";
            }else if (n==16){
                chNumber="(十六)";
            }else if (n==17){
                chNumber="(十七)";
            }else if (n==18){
                chNumber="(十八)";
            }else if (n==19){
                chNumber="(十九)";
            }else if (n==20){
                chNumber="(二十)";
            }
            
            
        }catch(Exception e){
            System.out.println("getMonthName Error:"+e.getMessage());
        }
        return chNumber;   
    }*/
    public static String getPercentNumber_3(String szNumString) {
        try {
            double dDouble=Double.parseDouble(szNumString);
            DecimalFormat df_md = new DecimalFormat("0.00");
            return df_md.format(dDouble).toString();            
        }catch (Exception e) {
            System.out.println(e);return szNumString;
        }
    }
    public static void copyRows(HSSFSheet pSourceSheet, HSSFSheet pTargetSheet, int pStartRow,
            int pEndRow, int pPosition, HSSFWorkbook wb) {
           HSSFRow sourceRow = null;
           HSSFRow targetRow = null;
           HSSFCell sourceCell = null;
           HSSFCell targetCell = null;
           HSSFSheet sourceSheet = null;
           HSSFSheet targetSheet = null;
           Region region = null;
           int cType;
           int i;
           short j;
           int targetRowFrom;
           int targetRowTo;
           if ((pStartRow == -1) || (pEndRow == -1)) {
            return;
           }
           
           sourceSheet = pSourceSheet;
           targetSheet = pTargetSheet;
           //copy合併的單元格
           for (i = 0; i < sourceSheet.getNumMergedRegions(); i++) {
            region = sourceSheet.getMergedRegionAt(i);
            if ((region.getRowFrom() >= pStartRow)
              && (region.getRowTo() <= pEndRow)) {
             targetRowFrom = region.getRowFrom() - pStartRow + pPosition;
             targetRowTo = region.getRowTo() - pStartRow + pPosition;
             region.setRowFrom(targetRowFrom);
             region.setRowTo(targetRowTo);
             targetSheet.addMergedRegion(region);
            }
           }
           //設置列寬
           //如果是同一頁就不需要設置列寬，否則會有問題
           if (pSourceSheet != pTargetSheet) {
            for (i = pStartRow; i <= pEndRow; i++) {
             sourceRow = sourceSheet.getRow(i);
             if (sourceRow != null) {
              for (j = sourceRow.getFirstCellNum(); j < sourceRow
                .getLastCellNum(); j++) {
               targetSheet.setColumnWidth(j, sourceSheet
                 .getColumnWidth(j));
              }
              break;
             }
            }
           }
           //拷貝並填充數據
           for (i = pStartRow;i <= pEndRow; i++) {
            sourceRow = sourceSheet.getRow(i);
            if (sourceRow == null) {
             continue;
            }
            targetRow = targetSheet.createRow(i - pStartRow + pPosition);
            targetRow.setHeight(sourceRow.getHeight());
            for (j = sourceRow.getFirstCellNum(); j < sourceRow
              .getLastCellNum(); j++) {
             sourceCell = sourceRow.getCell(j);
             if (sourceCell == null) {
              continue;
             }
             targetCell = targetRow.createCell(j);
             targetCell.setEncoding(sourceCell.getEncoding());
             targetCell.setCellStyle(sourceCell.getCellStyle());
             cType = sourceCell.getCellType();
             targetCell.setCellType(cType);
             switch (cType) {
             case HSSFCell.CELL_TYPE_BOOLEAN:
              targetCell.setCellValue(sourceCell.getBooleanCellValue());
              break;
             case HSSFCell.CELL_TYPE_ERROR:
              targetCell
                .setCellErrorValue(sourceCell.getErrorCellValue());
              break;
             case HSSFCell.CELL_TYPE_FORMULA:
              // parseFormula
              targetCell.setCellFormula(parseFormula(sourceCell
                .getCellFormula()));
              break;
             case HSSFCell.CELL_TYPE_NUMERIC:
              targetCell.setCellValue(sourceCell.getNumericCellValue());
              break;
             case HSSFCell.CELL_TYPE_STRING:
              targetCell.setCellValue(sourceCell.getStringCellValue());
              break;
             }
            }
           }
          }
          //這是為了解決單元格中設置了函數的複製問題
     private static String parseFormula(String pPOIFormula) {
        final String cstReplaceString = "ATTR(semiVolatile)"; //$NON-NLS-1$
          StringBuffer result = null;
           int index;
           result = new StringBuffer();
           index = pPOIFormula.indexOf(cstReplaceString);
           if (index >= 0) {
            result.append(pPOIFormula.substring(0, index));
            result.append(pPOIFormula.substring(index
              + cstReplaceString.length()));
           } else {
            result.append(pPOIFormula);
           }
           return result.toString();
          }
}
