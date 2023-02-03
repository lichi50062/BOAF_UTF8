/*
 * Created on 2006/10/25 by ABYSS Allen
 * AN007W 多個年度農漁會信用部簡明損益比較表
 * fixed 99.06.07 sql injection by 2808
 * fixed 102.06.03 sql by 2968
 * fixed 102.08.30 外包的connection拿掉 by 2968
 */

package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.poifs.filesystem.*;
import java.io.*;
import java.sql.*;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.List;

import com.tradevan.util.DBManager;
import com.tradevan.util.Utility;
import com.tradevan.util.Utility_report;
import com.tradevan.util.dao.DataObject;
import com.tradevan.util.dao.RdbCommonDao;

public class RptAN007W {

	public static String createRpt(String startYear, String bankType, int priceUtil) {
	    NumberFormat nf = NumberFormat.getInstance();
        nf.setMinimumFractionDigits(2); // 若小數點不足二位，則補足二位
	    String m_year = "";
	    String field_520100 = "";
	    String field_520100_rate = "";
	    String field_520200 = "";
	    String field_520200_rate = "";
	    String field_521700_521800 = "";
	    String field_110300_rate = "";
	    String field_521500 = "";
	    String field_521500_rate = "";
	    String field_other = "";
	    String field_other_rate = "";
	    String field_320300 = "";
	    String field_320300_rate = "";
	    String field_total1 = "";
	    String field_total1_rate = "";
	    String field_420100 = "";
	    String field_420100_rate = "";
	    String field_420300 = "";
	    String field_420300_rate = "";
	    String field_420170 = "";
	    String field_420170_rate = "";
	    String field_420400 = "";
	    String field_420400_rate = "";
	    String field_420700 = "";
	    String field_420700_rate = "";
	    String field_total2 = "";
	    String field_total2_rate = "";
		String[] priceUtilStr = new String[]{"元","仟元","萬元","百萬元","仟萬元","億元"};
		DataObject bean = null;
		System.out.println("RptAN007 createRpt() Debug Start ...");
		String errMsg = "";
		StringBuffer sqlCmd = new StringBuffer () ;
		List paramList = new ArrayList () ;
		String u_year = "99" ;
		if(!"".equals(startYear) && Integer.parseInt(startYear) >99) {
			u_year = "100" ;
		}
		int priceIndex=1;
		if(priceUtil==1){
			priceIndex=0;
		}else if(priceUtil==1000){
			priceIndex=1;
		}else if(priceUtil==10000){
			priceIndex=2;
		}else if(priceUtil==1000000){
			priceIndex=3;
		}else if(priceUtil==10000000){
			priceIndex=4;
		}else if(priceUtil==100000000){
			priceIndex=5;
		}
		System.out.println("bankType="+bankType+";priceUtil="+priceUtil+";priceIndex="+priceIndex+";priceUtilStr="+priceUtilStr[priceIndex]);
		try {
			sqlCmd.append("select m_year, ");
			sqlCmd.append("       round(field_520100 /?,0) as  field_520100, ");//存款利息支出
			sqlCmd.append("       decode(field_total1,0,0,round(field_520100 / field_total1 *100 ,2))  as field_520100_rate, ");//存款利息支出.%
			sqlCmd.append("       round(field_520200 /?,0) as  field_520200, ");//借款利息支出
			sqlCmd.append("       decode(field_total1,0,0,round(field_520200 / field_total1 *100 ,2))  as  field_520200_rate, ");//借款利息支出.%
			sqlCmd.append("       round(field_521700_521800 /?,0) as  field_521700_521800, ");//管理及會議費用
			sqlCmd.append("       decode(field_total1,0,0,round(field_521700_521800 / field_total1 *100 ,2))  as  field_110300_rate, ");
			sqlCmd.append("       round(field_521500 /?,0) as  field_521500, ");//用人費用
			sqlCmd.append("       decode(field_total1,0,0,round(field_521500 / field_total1 *100 ,2))  as field_521500_rate, ");
			sqlCmd.append("       round(field_other /?,0) as  field_other, ");//其他支出
			sqlCmd.append("       decode(field_total1,0,0,round(field_other / field_total1 *100 ,2))  as field_other_rate, ");
			sqlCmd.append("       round(field_320300 /?,0) as  field_320300, ");//本期損益
			sqlCmd.append("       decode(field_total1,0,0,round(field_320300 / field_total1 *100 ,2))  as field_320300_rate, ");
			sqlCmd.append("       round(field_total1 /?,0) as  field_total1, ");//支出合計
			sqlCmd.append("       decode(field_total1,0,0,round(field_total1 / field_total1 *100 ,2))  as field_total1_rate, ");
			sqlCmd.append("       round(field_420100 /?,0) as  field_420100, ");//放款利息收入
			sqlCmd.append("       decode(field_total2,0,0,round(field_420100 / field_total2 *100 ,2))  as field_420100_rate, ");
			sqlCmd.append("       round(field_420300 /?,0) as  field_420300, ");//存儲利息收入
			sqlCmd.append("       decode(field_total2,0,0,round(field_420300 / field_total2 *100 ,2))  as field_420300_rate, ");
			sqlCmd.append("       round(field_420170 /?,0) as  field_420170, ");//內部融資利息收入
			sqlCmd.append("       decode(field_total2,0,0,round(field_420170 / field_total2 *100 ,2))  as field_420170_rate, ");
			sqlCmd.append("       round(field_420400 /?,0) as  field_420400, ");//代辦業務及手續費收入
			sqlCmd.append("       decode(field_total2,0,0,round(field_420400 / field_total2 *100 ,2))  as field_420400_rate, ");
			sqlCmd.append("       round(field_420700 /?,0) as  field_420700, ");//其他收入
			sqlCmd.append("       decode(field_total2,0,0,round(field_420700 / field_total2 *100 ,2))  as field_420700_rate, ");
			sqlCmd.append("       round(field_total2 /?,0) as  field_total2, ");//收入合計
			sqlCmd.append("       decode(field_total2,0,0,round(field_total2 / field_total2 *100 ,2))  as field_total2_rate ");
			paramList.add(priceUtil) ;
            paramList.add(priceUtil) ;
            paramList.add(priceUtil) ;
            paramList.add(priceUtil) ;
            paramList.add(priceUtil) ;
            paramList.add(priceUtil) ;
            paramList.add(priceUtil) ;
            paramList.add(priceUtil) ;
            paramList.add(priceUtil) ;
            paramList.add(priceUtil) ;
            paramList.add(priceUtil) ;
            paramList.add(priceUtil) ;
            paramList.add(priceUtil) ;
			sqlCmd.append("from ");
			sqlCmd.append(" (  ");  
			sqlCmd.append("     select m_year, ");
			sqlCmd.append("              sum(field_520100) as field_520100, ");
			sqlCmd.append("              sum(field_520200) as field_520200, ");
			sqlCmd.append("              sum(field_521700_521800) as field_521700_521800, ");
			sqlCmd.append("              sum(field_521500) as field_521500, ");
			sqlCmd.append("              sum(field_520000)-sum(field_520100)-sum(field_520200)-sum(field_521700_521800)-sum(field_521500) as field_other, ");
			sqlCmd.append("              sum(field_320300) as field_320300, ");
			sqlCmd.append("              sum(field_520100)+sum(field_520200)+ sum(field_521700_521800) +sum(field_521500) +(sum(field_520000)-sum(field_520100)-sum(field_520200)-sum(field_521700_521800)-sum(field_521500)) + sum(field_320300) as field_total1, ");
			sqlCmd.append("              sum(field_420100) as field_420100, ");
			sqlCmd.append("              sum(field_420300) as field_420300, ");
			sqlCmd.append("              sum(field_420170) as field_420170, ");
			sqlCmd.append("              sum(field_420400) as field_420400, ");
			sqlCmd.append("              sum(field_420700) as field_420700, ");
			sqlCmd.append("              sum(field_420100)+sum(field_420300)+sum(field_420170)+ sum(field_420400)+sum(field_420700)  as field_total2   ");          
			sqlCmd.append("       from (   ");   
			sqlCmd.append("             select a01.m_year, ");
			sqlCmd.append("                      sum(decode(a01.acc_code,'520000',amt,0))  as  field_520000, ");
			sqlCmd.append("                      sum(decode(a01.acc_code,'520100',amt,0))  as  field_520100, ");
			sqlCmd.append("                      sum(decode(a01.acc_code,'520200',amt,0))  as  field_520200, ");
			sqlCmd.append("                      sum(decode(a01.acc_code,'521700',amt,'521800',amt,0))  as  field_521700_521800, ");
			sqlCmd.append("                      sum(decode(a01.acc_code,'521500',amt,0))  as  field_521500, ");
			sqlCmd.append("                      sum(decode(a01.acc_code,'320300',amt,0))  as  field_320300, ");
			sqlCmd.append("                      sum(decode(a01.acc_code,'420100',amt,0))  as  field_420100, ");                    
			sqlCmd.append("                      decode(YEAR_TYPE,'102',decode(bn01.bank_type,'6',sum(decode(a01.acc_code,'420300',amt,0)),'7',sum(decode(a01.acc_code,'420200',amt,0)),0),'103',sum(decode(a01.acc_code,'420300',amt,0)),0)  as  field_420300, ");
			sqlCmd.append("                      decode(bn01.bank_type,'6',sum(decode(a01.acc_code,'420170',amt,0)),0)  as  field_420170, ");
			sqlCmd.append("                      decode(YEAR_TYPE,'102',decode(bn01.bank_type,'6',sum(decode(a01.acc_code,'420400',amt,'422200',amt,'420900',amt,0)),'7',sum(decode(a01.acc_code,'420400',amt,0)),0),'103',sum(decode(a01.acc_code,'420400',amt,'420900',amt,0)),0)  as  field_420400, ");
			sqlCmd.append("                      sum(decode(a01.acc_code,'420700',amt,0))  as  field_420700 ");               
			sqlCmd.append("               from (select  (CASE WHEN (a01.m_year <= 102) THEN '102' ");
			sqlCmd.append("                                   WHEN (a01.m_year > 102) THEN '103' ");
			sqlCmd.append("                              ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01  ");
			sqlCmd.append("                              where m_year in (?,?) and m_month=12  ");
			paramList.add(Integer.parseInt(startYear)-1) ;
			paramList.add(startYear) ;
			sqlCmd.append("                              and a01.acc_code in ('320300','420100','420200','420300','420170','420400','422200','420900','420700','520000','520100','520200','521700','521800','521500')  ");                     
			sqlCmd.append("                              )a01, ");
			sqlCmd.append("                              (select * from bn01 where m_year=? ");
			paramList.add(u_year) ;
            if("ALL".equals(bankType)){
                sqlCmd.append("             and bank_type in ('6','7') " );
            }else{
                sqlCmd.append("             and bank_type in (?) " );
                paramList.add(bankType) ;
            }
			sqlCmd.append("                              and bn01.bn_type <> '2')bn01 ");
			sqlCmd.append("               where a01.bank_code = bn01.bank_no ");            
			sqlCmd.append("               group by a01.m_year,bn01.bank_type,YEAR_TYPE ");
			sqlCmd.append("             )a01 ");
			sqlCmd.append("        group by m_year ");
			sqlCmd.append(")  "); 
			List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,
                    "m_year,field_520100,field_520100_rate,field_520200,field_520200_rate,"+
                    "field_521700_521800,field_110300_rate,field_521500,field_521500_rate,"+
                    "field_other,field_other_rate,field_320300,field_320300_rate,field_total1,field_total1_rate,"+
                    "field_420100,field_420100_rate,field_420300,field_420300_rate,field_420170,field_420170_rate,"+
                    "field_420400,field_420400_rate,field_420700,field_420700_rate,field_total2,field_total2_rate");
            System.out.println("dbData.size=" + dbData.size()); 
            //====開始製作Excel檔案
			HSSFWorkbook wb = new HSSFWorkbook();
			HSSFSheet sheet = wb.createSheet("New Sheet 1");
			HSSFPrintSetup ps = sheet.getPrintSetup(); //取得列印設定
			//設定頁面符合列印大小
            sheet.setZoom(68, 100); // 螢幕上看到的縮放大小
            sheet.setAutobreaks(false); //自動分頁
            ps.setScale((short)75); //列印縮放百分比
            ps.setPaperSize((short)9); //設定紙張大小 A4
            ps.setLandscape(true); // 設定橫印
			HSSFRow row = null; //宣告一列
			HSSFDataFormat format = wb.createDataFormat();
			short totalColumnNums = (short)10;
			System.out.println("totalColumnNums="+totalColumnNums);
			short[] columnLen = new short[]{27,5,15,5,15,27,5,15,5,15};
    		for(int ai=0; ai<columnLen.length; ai++){
    			sheet.setColumnWidth((short)ai, (short)(256*(columnLen[ai]+4)));
    		}
    		
    		HSSFFont defaultFont = wb.createFont();
			defaultFont.setFontHeightInPoints((short)14);
			HSSFFont numberFont = wb.createFont();
			numberFont.setFontHeightInPoints((short)12);
			reportUtil reportUtil = new reportUtil();
			HSSFCellStyle titleStyle = reportUtil.getTitleStyle(wb); //標題用
			HSSFCellStyle defaultStyle = reportUtil.getDefaultStyle(wb);//有框內文置中
			defaultStyle.setFont(defaultFont);
			HSSFCellStyle rightStyle = reportUtil.getRightStyle(wb);//有框內文置右
			rightStyle.setFont(numberFont);
			HSSFCellStyle leftStyle = reportUtil.getLeftStyle(wb);
			leftStyle.setFont(defaultFont);
			HSSFCellStyle numberRightStyle = reportUtil.getRightStyle(wb);//有框整數值置右
            numberRightStyle.setDataFormat(format.getFormat("#,##0"));
            numberRightStyle.setFont(numberFont);
            HSSFCellStyle doubleRightStyle = reportUtil.getRightStyle(wb);//有框小數點置右
            doubleRightStyle.setDataFormat(format.getFormat("#,##0.00"));
            doubleRightStyle.setFont(numberFont);
			HSSFCellStyle noBorderRightStyle = reportUtil.getNoBoderStyle(wb);
			noBorderRightStyle.setFont(defaultFont);
			reportUtil.setDefaultStyle(defaultStyle);
			
			//設定報表表頭資料 開始============================================
			row=sheet.createRow(0);
			String titleStr="";
			if(bankType.equals("ALL")){
				titleStr = "多個年度農漁會信用部簡明損益比較表";
			}else if(bankType.equals("6")){
				titleStr = "多個年度農會信用部簡明損益比較表";
			}else{
				titleStr = "多個年度漁會信用部簡明損益比較表";
			}
			reportUtil.createCell(wb, row, (short)0, titleStr, titleStyle);
			for(int ci=1; ci<totalColumnNums;ci++){
				reportUtil.createCell( wb, row, (short)ci, "", titleStyle);
			}
			sheet.addMergedRegion(new Region((short)0, (short)0, (short)0, (short)(totalColumnNums-1)));
			row=sheet.createRow(1);
			row.setHeightInPoints(35.0F);
			reportUtil.createCell(wb, row, (short)0, "單位:"+priceUtilStr[priceIndex], noBorderRightStyle);
			for(int ci=1; ci<totalColumnNums;ci++){
				reportUtil.createCell( wb, row, (short)ci, "", noBorderRightStyle);
			}
			sheet.addMergedRegion(new Region((short)1, (short)0, (short)1, (short)(totalColumnNums-1)));
			row=sheet.createRow(2);
			row.setHeightInPoints(40.0F);
			reportUtil.createCell(wb, row, (short)0, "支  出", defaultStyle);
			reportUtil.createCell(wb, row, (short)1, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)2, "", defaultStyle);
			sheet.addMergedRegion(new Region((short)2, (short)1, (short)2, (short)2));
			reportUtil.createCell(wb, row, (short)3, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)4, "", defaultStyle);
			sheet.addMergedRegion(new Region((short)2, (short)3, (short)2, (short)4));
			reportUtil.createCell(wb, row, (short)5, "收  入", defaultStyle);
			reportUtil.createCell(wb, row, (short)6, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)7, "", defaultStyle);
			sheet.addMergedRegion(new Region((short)2, (short)6, (short)2, (short)7));
			reportUtil.createCell(wb, row, (short)8, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)9, "", defaultStyle);
			sheet.addMergedRegion(new Region((short)2, (short)8, (short)2, (short)9));
			row=sheet.createRow(3);
			row.setHeightInPoints(40.0F);
			reportUtil.createCell(wb, row, (short)0, "", defaultStyle);
			sheet.addMergedRegion(new Region((short)2, (short)0, (short)3, (short)0));
			reportUtil.createCell(wb, row, (short)1, "％", defaultStyle);
			reportUtil.createCell(wb, row, (short)2, "金  額", defaultStyle);
			reportUtil.createCell(wb, row, (short)3, "％", defaultStyle);
			reportUtil.createCell(wb, row, (short)4, "金  額", defaultStyle);
			reportUtil.createCell(wb, row, (short)5, "", defaultStyle);
			sheet.addMergedRegion(new Region((short)2, (short)5, (short)3, (short)5));
			reportUtil.createCell(wb, row, (short)6, "％", defaultStyle);
			reportUtil.createCell(wb, row, (short)7, "金額", defaultStyle);
			reportUtil.createCell(wb, row, (short)8, "％", defaultStyle);
			reportUtil.createCell(wb, row, (short)9, "金額", defaultStyle);
			row=sheet.createRow(4);
			row.setHeightInPoints(40.0F);
			reportUtil.createCell(wb, row, (short)0, "存款利息支出", leftStyle);
			reportUtil.createCell(wb, row, (short)5, "存款利息收入", leftStyle);
			row=sheet.createRow(5);
			row.setHeightInPoints(40.0F);
			reportUtil.createCell(wb, row, (short)0, "借款利息支出", leftStyle);
			reportUtil.createCell(wb, row, (short)5, "存儲利息收入", leftStyle);
			row=sheet.createRow(6);
			row.setHeightInPoints(40.0F);
			reportUtil.createCell(wb, row, (short)0, "管理及會議費用", leftStyle);
			reportUtil.createCell(wb, row, (short)5, "內部融資利息收入", leftStyle);
			row=sheet.createRow(7);
			row.setHeightInPoints(40.0F);
			reportUtil.createCell(wb, row, (short)0, "用人費用", leftStyle);
			reportUtil.createCell(wb, row, (short)5, "代辦業務及手續費收入", leftStyle);
			row=sheet.createRow(8);
			row.setHeightInPoints(40.0F);
			reportUtil.createCell(wb, row, (short)0, "其他支出", leftStyle);
			reportUtil.createCell(wb, row, (short)5, "其他收入", leftStyle);
			row=sheet.createRow(9);
			row.setHeightInPoints(40.0F);
			reportUtil.createCell(wb, row, (short)0, "本期損益", leftStyle);
			reportUtil.createCell(wb, row, (short)5, "", leftStyle);
			row=sheet.createRow(10);
			row.setHeightInPoints(40.0F);
			reportUtil.createCell(wb, row, (short)0, "合計", leftStyle);
			reportUtil.createCell(wb, row, (short)5, "合計", leftStyle);
			//wb.setRepeatingRowsAndColumns(0, 0, 9, 0, 9);//設為固定表頭(第幾個sheet,起始欄,終止欄,起始列,終止列)
			//設定報表表頭資料 結束============================================
			
			HSSFRow rowTmp=sheet.getRow(2);
			int cellNo = 1;
			if(dbData.size()>0){
			    for(int i=0;i<dbData.size();i++){  
                    bean = (DataObject)dbData.get(i);
                    m_year = (bean.getValue("m_year")==null)?"":(bean.getValue("m_year")).toString();
                    field_520100 = (bean.getValue("field_520100")==null)?"0":Utility.setCommaFormat((bean.getValue("field_520100")).toString());
                    field_520100_rate = (bean.getValue("field_520100_rate")==null)?"0.00":nf.format(bean.getValue("field_520100_rate"));
                    field_520200 = (bean.getValue("field_520200")==null)?"0":Utility.setCommaFormat((bean.getValue("field_520200")).toString());
                    field_520200_rate = (bean.getValue("field_520200_rate")==null)?"0.00":nf.format(bean.getValue("field_520200_rate"));
                    field_521700_521800 = (bean.getValue("field_521700_521800")==null)?"0":Utility.setCommaFormat((bean.getValue("field_521700_521800")).toString());
                    field_110300_rate = (bean.getValue("field_110300_rate")==null)?"0.00":nf.format(bean.getValue("field_110300_rate"));
                    field_521500 = (bean.getValue("field_521500")==null)?"0":Utility.setCommaFormat((bean.getValue("field_521500")).toString());
                    field_521500_rate = (bean.getValue("field_521500_rate")==null)?"0.00":nf.format(bean.getValue("field_521500_rate"));
                    field_other = (bean.getValue("field_other")==null)?"0":Utility.setCommaFormat((bean.getValue("field_other")).toString());
                    field_other_rate = (bean.getValue("field_other_rate")==null)?"0.00":nf.format(bean.getValue("field_other_rate"));
                    field_320300 = (bean.getValue("field_320300")==null)?"0":Utility.setCommaFormat((bean.getValue("field_320300")).toString());
                    field_320300_rate = (bean.getValue("field_320300_rate")==null)?"0.00":nf.format(bean.getValue("field_320300_rate"));
                    field_total1 = (bean.getValue("field_total1")==null)?"0":Utility.setCommaFormat((bean.getValue("field_total1")).toString());
                    field_total1_rate = (bean.getValue("field_total1_rate")==null)?"0.00":nf.format(bean.getValue("field_total1_rate"));
                    
                    field_420100 = (bean.getValue("field_420100")==null)?"0":Utility.setCommaFormat((bean.getValue("field_420100")).toString());
                    field_420100_rate = (bean.getValue("field_420100_rate")==null)?"0.00":nf.format(bean.getValue("field_420100_rate"));
                    field_420300 = (bean.getValue("field_420300")==null)?"0":Utility.setCommaFormat((bean.getValue("field_420300")).toString());
                    field_420300_rate = (bean.getValue("field_420300_rate")==null)?"0.00":nf.format(bean.getValue("field_420300_rate"));
                    field_420170 = (bean.getValue("field_420170")==null)?"0":Utility.setCommaFormat((bean.getValue("field_420170")).toString());
                    field_420170_rate = (bean.getValue("field_420170_rate")==null)?"0.00":nf.format(bean.getValue("field_420170_rate"));
                    field_420400 = (bean.getValue("field_420400")==null)?"0":Utility.setCommaFormat((bean.getValue("field_420400")).toString());
                    field_420400_rate = (bean.getValue("field_420400_rate")==null)?"0.00":nf.format(bean.getValue("field_420400_rate"));
                    field_420700 = (bean.getValue("field_420700")==null)?"0":Utility.setCommaFormat((bean.getValue("field_420700")).toString());
                    field_420700_rate = (bean.getValue("field_420700_rate")==null)?"0.00":nf.format(bean.getValue("field_420700_rate"));
                    field_total2 = (bean.getValue("field_total2")==null)?"0":Utility.setCommaFormat((bean.getValue("field_total2")).toString());
                    field_total2_rate = (bean.getValue("field_total2_rate")==null)?"0.00":nf.format(bean.getValue("field_total2_rate"));
                    if(dbData.size()==1){
                        if(m_year==startYear){ //缺前一年度資料
                            rowTmp=sheet.getRow(2);
                            reportUtil.createCell(wb, rowTmp, (short)cellNo, Integer.parseInt(startYear)-1+"年度", defaultStyle);//設年度
                            reportUtil.createCell(wb, rowTmp, (short)(cellNo+1), "", defaultStyle);
                            sheet.addMergedRegion(new Region((short)2, (short)cellNo, (short)2, (short)(cellNo+1)));
                            reportUtil.createCell(wb, rowTmp, (short)(cellNo+5), Integer.parseInt(startYear)-1+"年度", defaultStyle);
                            reportUtil.createCell(wb, rowTmp, (short)(cellNo+6), "", defaultStyle);
                            sheet.addMergedRegion(new Region((short)2, (short)(cellNo+5), (short)2, (short)(cellNo+6)));
                            //支出-%
                            rowTmp=sheet.getRow(4);
                            for(int r=4;r<=10;r++){
                                reportUtil.createCell(wb, sheet.getRow(r), (short)cellNo, "0.00", doubleRightStyle);
                            }
                            //支出-金額
                            sheet.setColumnWidth((short)(cellNo+1),(short)4800);
                            for(int r=4;r<=10;r++){
                                reportUtil.createCell(wb, sheet.getRow(r), (short)(cellNo+1), "0", numberRightStyle);
                            }
                            //收入-%
                            for(int r=4;r<=10;r++){
                                if(r==9){
                                    reportUtil.createCell(wb, sheet.getRow(r), (short)(cellNo+5), "", doubleRightStyle);
                                }else{
                                    reportUtil.createCell(wb, sheet.getRow(r), (short)(cellNo+5), "0.00", doubleRightStyle);
                                }
                            }
                            //收入-金額
                            sheet.setColumnWidth((short)(cellNo+6),(short)4800);
                            for(int r=4;r<=10;r++){
                                if(r==9){
                                    reportUtil.createCell(wb, sheet.getRow(r), (short)(cellNo+6), "", numberRightStyle);
                                }else{
                                    reportUtil.createCell(wb, sheet.getRow(r), (short)(cellNo+6), "0", numberRightStyle);
                                }
                            }
                            cellNo+=2;
                        }else{ //缺查詢年度資料
                            cellNo+=2;
                            rowTmp=sheet.getRow(2);
                            reportUtil.createCell(wb, rowTmp, (short)cellNo, startYear+"年度", defaultStyle);//設年度
                            reportUtil.createCell(wb, rowTmp, (short)(cellNo+1), "", defaultStyle);
                            sheet.addMergedRegion(new Region((short)2, (short)cellNo, (short)2, (short)(cellNo+1)));
                            reportUtil.createCell(wb, rowTmp, (short)(cellNo+5), startYear+"年度", defaultStyle);
                            reportUtil.createCell(wb, rowTmp, (short)(cellNo+6), "", defaultStyle);
                            sheet.addMergedRegion(new Region((short)2, (short)(cellNo+5), (short)2, (short)(cellNo+6)));
                            //支出-%
                            rowTmp=sheet.getRow(4);
                            for(int r=4;r<=10;r++){
                                reportUtil.createCell(wb, sheet.getRow(r), (short)cellNo, "0.00", doubleRightStyle);
                            }
                            //支出-金額
                            sheet.setColumnWidth((short)(cellNo+1),(short)4800);
                            for(int r=4;r<=10;r++){
                                reportUtil.createCell(wb, sheet.getRow(r), (short)(cellNo+1), "0", numberRightStyle);
                            }
                            //收入-%
                            for(int r=4;r<=10;r++){
                                if(r==9){
                                    reportUtil.createCell(wb, sheet.getRow(r), (short)(cellNo+5), "", doubleRightStyle);
                                }else{
                                    reportUtil.createCell(wb, sheet.getRow(r), (short)(cellNo+5), "0.00", doubleRightStyle);
                                }
                            }
                            //收入-金額
                            sheet.setColumnWidth((short)(cellNo+6),(short)4800);
                            for(int r=4;r<=10;r++){
                                if(r==9){
                                    reportUtil.createCell(wb, sheet.getRow(r), (short)(cellNo+6), "", numberRightStyle);
                                }else{
                                    reportUtil.createCell(wb, sheet.getRow(r), (short)(cellNo+6), "0", numberRightStyle);
                                }
                            }
                            cellNo-=2;
                        }
                    }
                    rowTmp=sheet.getRow(2);
                    reportUtil.createCell(wb, rowTmp, (short)cellNo, m_year+"年度", defaultStyle);//設年度
                    reportUtil.createCell(wb, rowTmp, (short)(cellNo+1), "", defaultStyle);
                    sheet.addMergedRegion(new Region((short)2, (short)cellNo, (short)2, (short)(cellNo+1)));
                    reportUtil.createCell(wb, rowTmp, (short)(cellNo+5), m_year+"年度", defaultStyle);
                    reportUtil.createCell(wb, rowTmp, (short)(cellNo+6), "", defaultStyle);
                    sheet.addMergedRegion(new Region((short)2, (short)(cellNo+5), (short)2, (short)(cellNo+6)));
                    //支出-%
                    reportUtil.createCell(wb, sheet.getRow(4), (short)cellNo, field_520100_rate, doubleRightStyle);
                    reportUtil.createCell(wb, sheet.getRow(5), (short)cellNo, field_520200_rate, doubleRightStyle);
                    reportUtil.createCell(wb, sheet.getRow(6), (short)cellNo, field_110300_rate, doubleRightStyle);
                    reportUtil.createCell(wb, sheet.getRow(7), (short)cellNo, field_521500_rate, doubleRightStyle);
                    reportUtil.createCell(wb, sheet.getRow(8), (short)cellNo, field_other_rate, doubleRightStyle);
                    reportUtil.createCell(wb, sheet.getRow(9), (short)cellNo, field_320300_rate, doubleRightStyle);
                    reportUtil.createCell(wb, sheet.getRow(10), (short)cellNo, field_total1_rate, doubleRightStyle);
                    //支出-金額
                    sheet.setColumnWidth((short)(cellNo+1),(short)4800);
                    reportUtil.createCell(wb, sheet.getRow(4), (short)(cellNo+1), field_520100, numberRightStyle);
                    reportUtil.createCell(wb, sheet.getRow(5), (short)(cellNo+1), field_520200, numberRightStyle);
                    reportUtil.createCell(wb, sheet.getRow(6), (short)(cellNo+1), field_521700_521800, numberRightStyle);
                    reportUtil.createCell(wb, sheet.getRow(7), (short)(cellNo+1), field_521500, numberRightStyle);
                    reportUtil.createCell(wb, sheet.getRow(8), (short)(cellNo+1), field_other, numberRightStyle);
                    reportUtil.createCell(wb, sheet.getRow(9), (short)(cellNo+1), field_320300, numberRightStyle);
                    reportUtil.createCell(wb, sheet.getRow(10), (short)(cellNo+1), field_total1, numberRightStyle);
                    //收入-%
                    reportUtil.createCell(wb, sheet.getRow(4), (short)(cellNo+5), field_420100_rate, doubleRightStyle);
                    reportUtil.createCell(wb, sheet.getRow(5), (short)(cellNo+5), field_420300_rate, doubleRightStyle);
                    reportUtil.createCell(wb, sheet.getRow(6), (short)(cellNo+5), field_420170_rate, doubleRightStyle);
                    reportUtil.createCell(wb, sheet.getRow(7), (short)(cellNo+5), field_420400_rate, doubleRightStyle);
                    reportUtil.createCell(wb, sheet.getRow(8), (short)(cellNo+5), field_420700_rate, doubleRightStyle);
                    reportUtil.createCell(wb, sheet.getRow(9), (short)(cellNo+5), "", doubleRightStyle);
                    reportUtil.createCell(wb, sheet.getRow(10), (short)(cellNo+5), field_total2_rate, doubleRightStyle);
                    //收入-金額
                    sheet.setColumnWidth((short)(cellNo+6),(short)4800);
                    reportUtil.createCell(wb, sheet.getRow(4), (short)(cellNo+6), field_420100, numberRightStyle);
                    reportUtil.createCell(wb, sheet.getRow(5), (short)(cellNo+6), field_420300, numberRightStyle);
                    reportUtil.createCell(wb, sheet.getRow(6), (short)(cellNo+6), field_420170, numberRightStyle);
                    reportUtil.createCell(wb, sheet.getRow(7), (short)(cellNo+6), field_420400, numberRightStyle);
                    reportUtil.createCell(wb, sheet.getRow(8), (short)(cellNo+6), field_420700, numberRightStyle);
                    reportUtil.createCell(wb, sheet.getRow(9), (short)(cellNo+6), "", numberRightStyle);
                    reportUtil.createCell(wb, sheet.getRow(10), (short)(cellNo+6), field_total2, numberRightStyle);
                    cellNo+=2;
			    }
			}else{
			    for(int i=1;i<=2;i++){
			        if(i==1){
			            m_year = String.valueOf(Integer.parseInt(startYear)-1);
			        }else{
			            m_year = startYear;
			        }
    			    rowTmp=sheet.getRow(2);
                    reportUtil.createCell(wb, rowTmp, (short)cellNo, m_year+"年度", defaultStyle);//設年度
                    reportUtil.createCell(wb, rowTmp, (short)(cellNo+1), "", defaultStyle);
                    sheet.addMergedRegion(new Region((short)2, (short)cellNo, (short)2, (short)(cellNo+1)));
                    reportUtil.createCell(wb, rowTmp, (short)(cellNo+5), Integer.parseInt(startYear)-1+"年度", defaultStyle);
                    reportUtil.createCell(wb, rowTmp, (short)(cellNo+6), "", defaultStyle);
                    sheet.addMergedRegion(new Region((short)2, (short)(cellNo+5), (short)2, (short)(cellNo+6)));
                    //支出-%
                    rowTmp=sheet.getRow(4);
                    for(int r=4;r<=10;r++){
                        reportUtil.createCell(wb, sheet.getRow(r), (short)cellNo, "0.00", doubleRightStyle);
                    }
                    //支出-金額
                    sheet.setColumnWidth((short)(cellNo+1),(short)4800);
                    for(int r=4;r<=10;r++){
                        reportUtil.createCell(wb, sheet.getRow(r), (short)(cellNo+1), "0", numberRightStyle);
                    }
                    //收入-%
                    for(int r=4;r<=10;r++){
                        if(r==9){
                            reportUtil.createCell(wb, sheet.getRow(r), (short)(cellNo+5), "", doubleRightStyle);
                        }else{
                            reportUtil.createCell(wb, sheet.getRow(r), (short)(cellNo+5), "0.00", doubleRightStyle);
                        }
                    }
                    //收入-金額
                    sheet.setColumnWidth((short)(cellNo+6),(short)4800);
                    for(int r=4;r<=10;r++){
                        if(r==9){
                            reportUtil.createCell(wb, sheet.getRow(r), (short)(cellNo+6), "", numberRightStyle);
                        }else{
                            reportUtil.createCell(wb, sheet.getRow(r), (short)(cellNo+6), "0", numberRightStyle);
                        }
                    }
                    cellNo+=2;
			    }
			}
			File reportDir = new File(Utility.getProperties("reportDir"));
			if (!reportDir.exists()) {
				if (!Utility.mkdirs(Utility.getProperties("reportDir"))) {
					errMsg += Utility.getProperties("reportDir") + "目錄新增失敗";
				}
			}
			FileOutputStream fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + titleStr+".xls");
			HSSFFooter footer = sheet.getFooter();
			footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
			wb.write(fout);
			//儲存
			fout.close();
		}catch (Exception e) {
			System.out.println("//RptAN007 createRpt() Have Error.....");
			e.printStackTrace();
			System.out.println("//-------------------------------------");
		}
		System.out.println("RptAN007 createRpt() Debug End ...");
		return errMsg;
	}
	private static void setPreparedStatementParameter(PreparedStatement pst,List paramList) throws Exception{
		for(int i = 0 ;i< paramList.size() ;i++) {
			pst.setString(i+1,(String)paramList.get(i)) ;
		}
	}
}
