/*
 * Created on 2006/10/24 by ABYSS Allen
 * 多個年度農漁會信用部放款結構及變動表
 * fixed 99.06.04 sql injection by 2808
 * fixed 102.06.03 sql by 2968
 * fixed 102.08.30 外包的connection拿掉 by 2968
 */

package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.Region;
import java.io.*;
import java.sql.*;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.List;

import com.tradevan.util.DBManager;
import com.tradevan.util.Utility;
import com.tradevan.util.dao.DataObject;
import com.tradevan.util.dao.RdbCommonDao;

public class RptAN004WB {

	public static String createRpt(String startYear, String endYear, String bankType, int priceUtil) {
	    NumberFormat nf = NumberFormat.getInstance();
	    nf.setMinimumFractionDigits(2);                        // 若小數點不足二位，則補足二位
	    String m_year = "";
	    String field_120101 = "";
	    String field_120101_rate = "";
	    String field_120102 = "";
	    String field_120102_rate = "";
	    String field_120200 = "";
	    String field_120200_rate = "";
	    String field_120401 = "";
	    String field_120401_rate = "";
	    String field_subtotal1 = "";
	    String field_subtotal1_rate = "";
	    String field_120501 = "";
	    String field_120501_rate = "";
	    String field_120602 = "";
	    String field_120602_rate = "";
	    String field_120601 = "";
	    String field_120601_rate = "";
	    String field_120603 = "";
	    String field_120603_rate = "";
	    String field_120604 = "";
	    String field_120604_rate = "";
	    String field_subtotal2 = "";
	    String field_subtotal2_rate = "";
	    String field_120700 = "";
	    String field_120700_rate = "";
	    String field_150200 = "";
	    String field_150200_rate = "";
	    String field_total = "";
	    String field_total_rate = "";
	    String[] priceUtilStr = new String[]{"元","仟元","萬元","百萬元","仟萬元","億元"};
		System.out.println("RptAN004WB createRpt() Debug Start ...");
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
		try {
			sqlCmd.append("select m_year," );
            sqlCmd.append("       round(field_120101 /?,0) as field_120101," );//   --一般放款.無擔保放款
            sqlCmd.append("       decode(field_total,0,0,round(field_120101 / field_total *100 ,2)) as field_120101_rate," );
            sqlCmd.append("       round(field_120102 /?,0) as field_120102," );//  --一般放款.擔保放款
            sqlCmd.append("       decode(field_total,0,0,round(field_120102 / field_total *100 ,2)) as field_120102_rate," );
            sqlCmd.append("       round(field_120200 /?,0) as field_120200," );//  --一般放款.貼現及透支
            sqlCmd.append("       decode(field_total,0,0,round(field_120200 / field_total *100 ,2)) as field_120200_rate," ); 
            sqlCmd.append("       round(field_120401 /?,0) as field_120401," );//   --一般放款.統一農.漁貸
            sqlCmd.append("       decode(field_total,0,0,round(field_120401 / field_total *100 ,2)) as field_120401_rate," ); 
            sqlCmd.append("       round(field_subtotal1/?,0) as field_subtotal1," );//--一般放款.小計
            sqlCmd.append("       decode(field_total,0,0,round(field_subtotal1 / field_total *100 ,2)) as field_subtotal1_rate," ); 
            sqlCmd.append("       round(field_120501 /?,0) as field_120501," );// --專案放款.專案放款
            sqlCmd.append("       decode(field_total,0,0,round(field_120501 / field_total *100 ,2)) as field_120501_rate," ); 
            sqlCmd.append("       round(field_120602 /?,0) as field_120602," );//  --專案放款.農業發展基金放款.農機放款
            sqlCmd.append("       decode(field_total,0,0,round(field_120602 / field_total *100 ,2)) as field_120602_rate," ); 
            sqlCmd.append("       round(field_120601 /?,0) as field_120601," );// --專案放款.農業發展基金放款.農建放款
            sqlCmd.append("       decode(field_total,0,0,round(field_120601 / field_total *100 ,2)) as field_120601_rate," ); 
            sqlCmd.append("       round(field_120603 /?,0) as field_120603," );// --專案放款.農業發展基金放款.購地放款(其他)
            sqlCmd.append("       decode(field_total,0,0,round(field_120603 / field_total *100 ,2)) as field_120603_rate," ); 
            sqlCmd.append("       round(field_120604 /?,0) as field_120604," );// --專案放款.農業發展基金放款.農宅放款
            sqlCmd.append("       decode(field_total,0,0,round(field_120604 / field_total *100 ,2)) as field_120604_rate," ); 
            sqlCmd.append("       round(field_subtotal2/?,0) as field_subtotal2," );//--專案放款.小計
            sqlCmd.append("       decode(field_total,0,0,round(field_subtotal2 / field_total *100 ,2)) as field_subtotal2_rate," ); 
            sqlCmd.append("       round(field_120700 /?,0) as field_120700," );//--內部融資
            sqlCmd.append("       decode(field_total,0,0,round(field_120700 / field_total *100 ,2)) as field_120700_rate," ); 
            sqlCmd.append("       round(field_150200 /?,0) as field_150200," ); //--催收款項
            sqlCmd.append("       decode(field_total,0,0,round(field_150200 / field_total *100 ,2)) as field_150200_rate," ); 
            sqlCmd.append("       round(field_total /?,0) as field_total," );//--合計
            sqlCmd.append("       decode(field_total,0,0,round(field_total / field_total *100 ,2)) as field_total_rate " );
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
            paramList.add(priceUtil) ;
            sqlCmd.append("from " );
            sqlCmd.append("(" );
            sqlCmd.append("       select m_year," );
            sqlCmd.append("              sum(field_120101) as field_120101," );
            sqlCmd.append("              sum(field_120102) as field_120102," );
            sqlCmd.append("              sum(field_120200) as field_120200," );
            sqlCmd.append("              sum(field_120401) as field_120401," );
            sqlCmd.append("              sum(field_120101)+sum(field_120102)+sum(field_120200)+sum(field_120401) as field_subtotal1," );//--一般放款.小計
            sqlCmd.append("              sum(field_120501) as field_120501," );
            sqlCmd.append("              sum(field_120602) as field_120602," );
            sqlCmd.append("              sum(field_120601) as field_120601," );
            sqlCmd.append("              sum(field_120603) as field_120603," );
            sqlCmd.append("              sum(field_120604) as field_120604," );
            sqlCmd.append("              sum(field_120501)+sum(field_120602)+sum(field_120601)+sum(field_120603)+sum(field_120604) as field_subtotal2," );//--專案放款.小計
            sqlCmd.append("              sum(field_120700) as field_120700," );
            sqlCmd.append("              sum(field_150200) as field_150200," );
            sqlCmd.append("              sum(field_120101)+sum(field_120102)+sum(field_120200)+sum(field_120401)+sum(field_120501)+sum(field_120602)+sum(field_120601)+sum(field_120603)+sum(field_120604)+sum(field_120700)+sum(field_150200) as field_total " );//--合計
            sqlCmd.append("       from (" );
            sqlCmd.append("               select a01.m_year," );
            sqlCmd.append("                      sum(decode(a01.acc_code,'120101',amt,0)) as field_120101," );
            sqlCmd.append("                      sum(decode(a01.acc_code,'120102',amt,0)) as field_120102," );
            sqlCmd.append("                      decode(YEAR_TYPE,'102',decode(bn01.bank_type,'6',sum(decode(a01.acc_code,'120200',amt,'120301',amt,'120302',amt,0)),'7',sum(decode(a01.acc_code,'120300',amt,'120401',amt,'120402',amt,0)),0),'103',sum(decode(a01.acc_code,'120200',amt,'120301',amt,'120302',amt,0)),0 ) as field_120200," );    
            sqlCmd.append("                      decode(YEAR_TYPE,'102',decode(bn01.bank_type,'6',sum(decode(a01.acc_code,'120401',amt,'120402',amt,0)),'7',sum(decode(a01.acc_code,'120201',amt,'120202',amt,0)),0),'103',sum(decode(a01.acc_code,'120401',amt,'120402',amt,0)),0) as field_120401," );
            sqlCmd.append("                      sum(decode(a01.acc_code,'120501',amt,'120502',amt,0)) as field_120501," );
            sqlCmd.append("                      sum(decode(a01.acc_code,'120602',amt,0)) as field_120602," );
            sqlCmd.append("                      sum(decode(a01.acc_code,'120601',amt,0)) as field_120601," );
            sqlCmd.append("                      decode(YEAR_TYPE,'102',decode(bn01.bank_type,'6',sum(decode(a01.acc_code,'120603',amt,0)),'7',sum(decode(a01.acc_code,'120604',amt,0)),0),'103',sum(decode(a01.acc_code,'120603',amt,0)),0) as field_120603," );
            sqlCmd.append("                      decode(YEAR_TYPE,'102',decode(bn01.bank_type,'6',sum(decode(a01.acc_code,'120604',amt,0)),'7',sum(decode(a01.acc_code,'120603',amt,0)),0),'103',sum(decode(a01.acc_code,'120604',amt,0)),0) as field_120604," );
            sqlCmd.append("                      sum(decode(a01.acc_code,'120700',amt,0)) as field_120700," );
            sqlCmd.append("                      sum(decode(a01.acc_code,'150200',amt,0)) as field_150200 " );
            sqlCmd.append("               from (select  (CASE WHEN (a01.m_year <= 102) THEN '102'" );
            sqlCmd.append("                                   WHEN (a01.m_year > 102) THEN '103'" );
            sqlCmd.append("                                ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01 " );
            sqlCmd.append("                       where m_year in (? ");
            paramList.add(startYear) ;
            for(int y=Integer.parseInt(startYear)+1; y<=Integer.parseInt(endYear);y++){
                sqlCmd.append(",?" );    
                paramList.add(y) ;
            }
            sqlCmd.append("                     ) and m_month=12)a01,(select * from bn01 where m_year=?)bn01 " );
            paramList.add(u_year) ;
            sqlCmd.append("               where a01.bank_code = bn01.bank_no" );
            sqlCmd.append("                 and bn01.bn_type <> '2' " );
            if("ALL".equals(bankType)){
                sqlCmd.append("             and bank_type in ('6','7') " );
            }else{
                sqlCmd.append("             and bank_type in (?) " );
                paramList.add(bankType) ;
            }
            sqlCmd.append("                 and a01.acc_code in ('120101','120102','120200','120201','120202','120300','120301','120302','120401','120402','120501','120502','120601','120602','120603','120604','120700','150200') " );
            sqlCmd.append("               group by a01.m_year,bn01.bank_type,YEAR_TYPE " );
            sqlCmd.append("       )a01 " );
            sqlCmd.append("       group by m_year " );
            sqlCmd.append(")" );
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,
                    "m_year,field_120101,field_120101_rate,field_120102,field_120102_rate,field_120200," +
                    "field_120200_rate,field_120401,field_120401_rate,field_subtotal1,field_subtotal1_rate," + 
                    "field_120501,field_120501_rate,field_120602,field_120602_rate,field_120601,field_120601_rate," +
                    "field_120603,field_120603_rate,field_120604,field_120604_rate,field_subtotal2,field_subtotal2_rate," +
                    "field_120700,field_120700_rate,field_150200,field_150200_rate,field_total,field_total_rate");
            System.out.println("dbData.size=" + dbData.size());
            
			HSSFWorkbook wb = new HSSFWorkbook();
			HSSFSheet sheet = wb.createSheet("New Sheet 1");
			HSSFPrintSetup ps = sheet.getPrintSetup(); //取得列印設定
			//設定頁面符合列印大小
            sheet.setZoom(65, 100); // 螢幕上看到的縮放大小
            sheet.setAutobreaks(false); //自動分頁            
            ps.setScale((short)75); //列印縮放百分比
            ps.setPaperSize((short)9); //設定紙張大小 A4
            ps.setLandscape(true); // 設定橫印
			HSSFRow row = null; //宣告一列
			HSSFDataFormat format = wb.createDataFormat();
			reportUtil reportUtil = new reportUtil();
			HSSFCellStyle titleStyle = reportUtil.getTitleStyle(wb); //標題用
			HSSFCellStyle defaultStyle = reportUtil.getDefaultStyle(wb);//有框內文置中
			HSSFFont defaultFont = wb.createFont();
			defaultFont.setFontHeightInPoints((short)12);
			HSSFFont numberFont = wb.createFont();
			numberFont.setFontHeightInPoints((short)12);
			defaultStyle.setFont(defaultFont);
			HSSFCellStyle rightStyle = reportUtil.getRightStyle(wb);//有框內文置右
			rightStyle.setFont(defaultFont);
			HSSFCellStyle numberRightStyle = reportUtil.getRightStyle(wb);//有框整數值置右
			numberRightStyle.setDataFormat(format.getFormat("#,##0"));
			numberRightStyle.setFont(numberFont);
			HSSFCellStyle doubleRightStyle = reportUtil.getRightStyle(wb);//有框小數點置右
			doubleRightStyle.setDataFormat(format.getFormat("#,##0.00"));
			doubleRightStyle.setFont(numberFont);
			HSSFCellStyle noBorderRightStyle = reportUtil.getNoBoderStyle(wb);
			noBorderRightStyle.setFont(defaultFont);
			reportUtil.setDefaultStyle(defaultStyle);
			sheet.setColumnWidth((short)0,(short)1200);
            sheet.setColumnWidth((short)1,(short)1200);
            sheet.setColumnWidth((short)2,(short)4800);
			//設定報表表頭資料 開始============================================
			row=sheet.createRow(0);
			String titleStr="";
			if(bankType.equals("ALL")){
				titleStr = "多個年度農漁會信用部放款結構及變動表";
			}else if(bankType.equals("6")){
				titleStr = "多個年度農會信用部放款結構及變動表";
			}else{
				titleStr = "多個年度漁會信用部放款結構及變動表";
			}
			int qAllYear = Integer.parseInt(endYear)-Integer.parseInt(startYear)+1;
			int repeatTime = 1;
			if(qAllYear>5 && (qAllYear%5)>0){
				repeatTime++;
			}else if(dbData.size()==0){ //即使撈出0筆資料也能印出title
			    repeatTime++;
			}
			for(int ri=0;ri<repeatTime;ri++){
				int startIndex=0;
				if(ri>0){
					startIndex=(ri*10)+3;
				}
				reportUtil.createCell(wb, row, (short)startIndex, titleStr, titleStyle);
				int endIndex=12;
				if(ri>0){
					endIndex=((ri+1)*10)+2;
				}
				for(int ci=(startIndex+1); ci<endIndex;ci++){
					reportUtil.createCell( wb, row, (short)ci, "", titleStyle);
				}
				sheet.addMergedRegion(new Region((short)0, (short)startIndex, (short)0, (short)endIndex));
			}
			
			row=sheet.createRow(1);
			row.setHeightInPoints(20.0F);
			for(int ri=0;ri<repeatTime;ri++){
				int startIndex=0;
				if(ri>0){
					startIndex=(ri*10)+3;
				}
				reportUtil.createCell(wb, row, (short)startIndex, "單位：新台幣 "+priceUtilStr[priceIndex], noBorderRightStyle);
				int endIndex=12;
				if(ri>0){
					endIndex=((ri+1)*10)+2;
				}
				for(int ci=(startIndex+1); ci<endIndex;ci++){
					reportUtil.createCell( wb, row, (short)ci, "", titleStyle);
				}
				sheet.addMergedRegion(new Region((short)1, (short)startIndex, (short)1, (short)endIndex));
			}
			row=sheet.createRow(2);
			row.setHeightInPoints(20.0F);
			reportUtil.createCell(wb, row, (short)0, "年度別", rightStyle);
			reportUtil.createCell(wb, row, (short)1, "", rightStyle);
			reportUtil.createCell(wb, row, (short)2, "", rightStyle);
			row=sheet.createRow(3);
			row.setHeightInPoints(20.0F);
			reportUtil.createCell(wb, row, (short)0, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)1, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)2, "", defaultStyle);
			sheet.addMergedRegion(new Region((short)2, (short)0, (short)3, (short)2));
			row=sheet.createRow(4);
			row.setHeightInPoints(34.0F);
			reportUtil.createCell(wb, row, (short)0, "放款項目", defaultStyle);
			reportUtil.createCell(wb, row, (short)1, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)2, "", defaultStyle);
			sheet.addMergedRegion(new Region((short)4, (short)0, (short)4, (short)2));
			row=sheet.createRow(5);
			row.setHeightInPoints(34.0F);
			reportUtil.createCell(wb, row, (short)0, "一般放款", defaultStyle);
			reportUtil.createCell(wb, row, (short)1, "無擔保", defaultStyle);
			reportUtil.createCell(wb, row, (short)2, "", defaultStyle);
			sheet.addMergedRegion(new Region((short)5, (short)1, (short)5, (short)2));
			row=sheet.createRow(6);
			row.setHeightInPoints(34.0F);
			reportUtil.createCell(wb, row, (short)0, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)1, "擔保放款", defaultStyle);
			reportUtil.createCell(wb, row, (short)2, "", defaultStyle);
			sheet.addMergedRegion(new Region((short)6, (short)1, (short)6, (short)2));
			row=sheet.createRow(7);
			row.setHeightInPoints(34.0F);
			reportUtil.createCell(wb, row, (short)0, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)1, "貼現及透支", defaultStyle);
			reportUtil.createCell(wb, row, (short)2, "", defaultStyle);
			sheet.addMergedRegion(new Region((short)7, (short)1, (short)7, (short)2));
			row=sheet.createRow(8);
			row.setHeightInPoints(34.0F);
			reportUtil.createCell(wb, row, (short)0, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)1, "統一農貸", defaultStyle);
			reportUtil.createCell(wb, row, (short)2, "", defaultStyle);
			sheet.addMergedRegion(new Region((short)8, (short)1, (short)8, (short)2));
			row=sheet.createRow(9);
			row.setHeightInPoints(34.0F);
			reportUtil.createCell(wb, row, (short)0, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)1, "小計", defaultStyle);
			reportUtil.createCell(wb, row, (short)2, "", defaultStyle);
			sheet.addMergedRegion(new Region((short)9, (short)1, (short)9, (short)2));
			sheet.addMergedRegion(new Region((short)5, (short)0, (short)9, (short)0));
			row=sheet.createRow(10);
			row.setHeightInPoints(34.0F);
			reportUtil.createCell(wb, row, (short)0, "專案放款", defaultStyle);
			reportUtil.createCell(wb, row, (short)1, "專案放款", defaultStyle);
			reportUtil.createCell(wb, row, (short)2, "", defaultStyle);
			sheet.addMergedRegion(new Region((short)10, (short)1, (short)10, (short)2));
			row=sheet.createRow(11);
			row.setHeightInPoints(34.0F);
			reportUtil.createCell(wb, row, (short)0, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)1, "農業發展基金放款", defaultStyle);
			reportUtil.createCell(wb, row, (short)2, "農機放款", defaultStyle);
			row=sheet.createRow(12);
			row.setHeightInPoints(34.0F);
			reportUtil.createCell(wb, row, (short)0, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)1, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)2, "農建放款", defaultStyle);
			row=sheet.createRow(13);
			row.setHeightInPoints(34.0F);
			reportUtil.createCell(wb, row, (short)0, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)1, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)2, "購地放款(其他)", defaultStyle);
			row=sheet.createRow(14);
			row.setHeightInPoints(34.0F);
			reportUtil.createCell(wb, row, (short)0, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)1, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)2, "農宅放款", defaultStyle);
			sheet.addMergedRegion(new Region((short)11, (short)1, (short)14, (short)1));
			row=sheet.createRow(15);
			row.setHeightInPoints(34.0F);
			reportUtil.createCell(wb, row, (short)0, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)1, "小計", defaultStyle);
			reportUtil.createCell(wb, row, (short)2, "", defaultStyle);
			sheet.addMergedRegion(new Region((short)15, (short)1, (short)15, (short)2));
			sheet.addMergedRegion(new Region((short)10, (short)0, (short)15, (short)0));
			row=sheet.createRow(16);
			row.setHeightInPoints(34.0F);
			reportUtil.createCell(wb, row, (short)0, "內部融資", defaultStyle);
			reportUtil.createCell(wb, row, (short)1, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)2, "", defaultStyle);
			sheet.addMergedRegion(new Region((short)16, (short)0, (short)16, (short)2));
			row=sheet.createRow(17);
			row.setHeightInPoints(34.0F);
			reportUtil.createCell(wb, row, (short)0, "催收款項", defaultStyle);
			reportUtil.createCell(wb, row, (short)1, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)2, "", defaultStyle);
			sheet.addMergedRegion(new Region((short)17, (short)0, (short)17, (short)2));
			row=sheet.createRow(18);
			row.setHeightInPoints(34.0F);
			reportUtil.createCell(wb, row, (short)0, "合計", defaultStyle);
			reportUtil.createCell(wb, row, (short)1, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)2, "", defaultStyle);
			sheet.addMergedRegion(new Region((short)18, (short)0, (short)18, (short)2));
			wb.setRepeatingRowsAndColumns(0, 0, 2, 2, 18);//設為固定表頭(第幾個sheet,起始欄,終止欄,起始列,終止列)
			//設定報表表頭資料 結束============================================
			//設定儲存格資料============================================
			int rowNo=4;
			int cellNo=2;
			DataObject bean = null;
			int qYear = Integer.parseInt(startYear);
			HSSFRow rowTmp=sheet.getRow(2);
			if(dbData.size()>0){
    			for(int i=0;i<dbData.size();i++){  
                    bean = (DataObject)dbData.get(i);
                    m_year = (bean.getValue("m_year")==null)?"":(bean.getValue("m_year")).toString();
                    field_120101 = (bean.getValue("field_120101")==null)?"0":Utility.setCommaFormat((bean.getValue("field_120101")).toString());
                    field_120101_rate = (bean.getValue("field_120101_rate")==null)?"0.00":nf.format(bean.getValue("field_120101_rate"));
                    field_120102 = (bean.getValue("field_120102")==null)?"0":Utility.setCommaFormat((bean.getValue("field_120102")).toString());
                    field_120102_rate = (bean.getValue("field_120102_rate")==null)?"0.00":nf.format(bean.getValue("field_120102_rate"));
                    field_120200 = (bean.getValue("field_120200")==null)?"0":Utility.setCommaFormat((bean.getValue("field_120200")).toString());
                    field_120200_rate = (bean.getValue("field_120200_rate")==null)?"0.00":nf.format(bean.getValue("field_120200_rate"));
                    field_120401 = (bean.getValue("field_120401")==null)?"0":Utility.setCommaFormat((bean.getValue("field_120401")).toString());
                    field_120401_rate = (bean.getValue("field_120401_rate")==null)?"0.00":nf.format(bean.getValue("field_120401_rate"));
                    field_subtotal1 = (bean.getValue("field_subtotal1")==null)?"0":Utility.setCommaFormat((bean.getValue("field_subtotal1")).toString());
                    field_subtotal1_rate = (bean.getValue("field_subtotal1_rate")==null)?"0.00":nf.format(bean.getValue("field_subtotal1_rate")); 
                    field_120501 = (bean.getValue("field_120501")==null)?"0":Utility.setCommaFormat((bean.getValue("field_120501")).toString());
                    field_120501_rate = (bean.getValue("field_120501_rate")==null)?"0.00":nf.format(bean.getValue("field_120501_rate"));
                    field_120602 = (bean.getValue("field_120602")==null)?"0":Utility.setCommaFormat((bean.getValue("field_120602")).toString());
                    field_120602_rate = (bean.getValue("field_120602_rate")==null)?"0.00":nf.format(bean.getValue("field_120602_rate"));
                    field_120601 = (bean.getValue("field_120601")==null)?"0":Utility.setCommaFormat((bean.getValue("field_120601")).toString());
                    field_120601_rate = (bean.getValue("field_120601_rate")==null)?"0.00":nf.format(bean.getValue("field_120601_rate"));
                    field_120603 = (bean.getValue("field_120603")==null)?"0":Utility.setCommaFormat((bean.getValue("field_120603")).toString());
                    field_120603_rate = (bean.getValue("field_120603_rate")==null)?"0.00":nf.format(bean.getValue("field_120603_rate"));
                    field_120604 = (bean.getValue("field_120604")==null)?"0":Utility.setCommaFormat((bean.getValue("field_120604")).toString());
                    field_120604_rate = (bean.getValue("field_120604_rate")==null)?"0.00":nf.format(bean.getValue("field_120604_rate"));
                    field_subtotal2 = (bean.getValue("field_subtotal2")==null)?"0":Utility.setCommaFormat((bean.getValue("field_subtotal2")).toString());
                    field_subtotal2_rate = (bean.getValue("field_subtotal2_rate")==null)?"0.00":nf.format(bean.getValue("field_subtotal2_rate"));
                    field_120700 = (bean.getValue("field_120700")==null)?"0":Utility.setCommaFormat((bean.getValue("field_120700")).toString());
                    field_120700_rate = (bean.getValue("field_120700_rate")==null)?"0.00":nf.format(bean.getValue("field_120700_rate"));
                    field_150200 = (bean.getValue("field_150200")==null)?"0":Utility.setCommaFormat((bean.getValue("field_150200")).toString());
                    field_150200_rate = (bean.getValue("field_150200_rate")==null)?"0.00":nf.format(bean.getValue("field_150200_rate"));
                    field_total = (bean.getValue("field_total")==null)?"0":Utility.setCommaFormat((bean.getValue("field_total")).toString());
                    field_total_rate = (bean.getValue("field_total_rate")==null)?"0.00":nf.format(bean.getValue("field_total_rate"));
                    row=sheet.getRow(rowNo);
                    cellNo++;
                    if(i==0 && Integer.parseInt(m_year)>Integer.parseInt(startYear)){
                        for(int y=Integer.parseInt(startYear); y<Integer.parseInt(m_year);y++){
                            rowTmp=sheet.getRow(2);
                            reportUtil.createCell(wb, rowTmp, (short)cellNo, y+"年度 資料不完整", defaultStyle);//設年度
                            reportUtil.createCell(wb, rowTmp, (short)(cellNo+1), "", defaultStyle);
                            rowTmp=sheet.getRow(3);
                            reportUtil.createCell(wb, rowTmp, (short)cellNo, "", defaultStyle);//設年度
                            reportUtil.createCell(wb, rowTmp, (short)(cellNo+1), "", defaultStyle);
                            sheet.addMergedRegion(new Region((short)2, (short)cellNo, (short)3, (short)(cellNo+1)));
                            rowTmp=sheet.getRow(4);
                            sheet.setColumnWidth((short)cellNo,(short)4500);
                            reportUtil.createCell(wb, rowTmp, (short)cellNo, "金額", rightStyle);
                            reportUtil.createCell(wb, rowTmp, (short)(cellNo+1), "%", rightStyle);
                            for(int r=5;r<=18;r++){
                                reportUtil.createCell(wb, sheet.getRow(r), (short)cellNo, "", numberRightStyle);
                                reportUtil.createCell(wb, sheet.getRow(r), (short)(cellNo+1), "", doubleRightStyle);
                            }
                            cellNo+=2;
                            qYear++;
                        }
                    }
                    if(Integer.parseInt(m_year)!=qYear){
                        rowTmp=sheet.getRow(2);
                        reportUtil.createCell(wb, rowTmp, (short)cellNo, qYear+"年度 資料不完整", defaultStyle);//設年度
                        reportUtil.createCell(wb, rowTmp, (short)(cellNo+1), "", defaultStyle);
                        rowTmp=sheet.getRow(3);
                        reportUtil.createCell(wb, rowTmp, (short)cellNo, "", defaultStyle);//設年度
                        reportUtil.createCell(wb, rowTmp, (short)(cellNo+1), "", defaultStyle);
                        sheet.addMergedRegion(new Region((short)2, (short)cellNo, (short)3, (short)(cellNo+1)));
                        rowTmp=sheet.getRow(4);
                        sheet.setColumnWidth((short)cellNo,(short)4500);
                        reportUtil.createCell(wb, rowTmp, (short)cellNo, "金額", rightStyle);
                        reportUtil.createCell(wb, rowTmp, (short)(cellNo+1), "%", rightStyle);
                        for(int r=5;r<=18;r++){
                            reportUtil.createCell(wb, sheet.getRow(r), (short)cellNo, "", numberRightStyle);
                            reportUtil.createCell(wb, sheet.getRow(r), (short)(cellNo+1), "", doubleRightStyle);
                        }
                        cellNo++;
                    }else{
                        rowTmp=sheet.getRow(2);
                        reportUtil.createCell(wb, rowTmp, (short)cellNo, m_year+"年度", defaultStyle);//設年度
                        reportUtil.createCell(wb, rowTmp, (short)(cellNo+1), "", defaultStyle);
                        rowTmp=sheet.getRow(3);
                        reportUtil.createCell(wb, rowTmp, (short)cellNo, "", defaultStyle);//設年度
                        reportUtil.createCell(wb, rowTmp, (short)(cellNo+1), "", defaultStyle);
                        sheet.addMergedRegion(new Region((short)2, (short)cellNo, (short)3, (short)(cellNo+1)));
                        rowTmp=sheet.getRow(4);
                        sheet.setColumnWidth((short)cellNo,(short)4500);
                        reportUtil.createCell(wb, rowTmp, (short)cellNo, "金額", rightStyle);
                        reportUtil.createCell(wb, rowTmp, (short)(cellNo+1), "%", rightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+1), (short)cellNo, field_120101, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+2), (short)cellNo, field_120102, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+3), (short)cellNo, field_120200, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+4), (short)cellNo, field_120401, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+5), (short)cellNo, field_subtotal1, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+6), (short)cellNo, field_120501, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+7), (short)cellNo, field_120602, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+8), (short)cellNo, field_120601, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+9), (short)cellNo, field_120603, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+10), (short)cellNo, field_120604, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+11), (short)cellNo, field_subtotal2, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+12), (short)cellNo, field_120700, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+13), (short)cellNo, field_150200, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+14), (short)cellNo, field_total, numberRightStyle);
                        cellNo++;
                        reportUtil.createCell(wb, sheet.getRow(rowNo+1), (short)cellNo, field_120101_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+2), (short)cellNo, field_120102_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+3), (short)cellNo, field_120200_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+4), (short)cellNo, field_120401_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+5), (short)cellNo, field_subtotal1_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+6), (short)cellNo, field_120501_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+7), (short)cellNo, field_120602_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+8), (short)cellNo, field_120601_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+9), (short)cellNo, field_120603_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+10), (short)cellNo, field_120604_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+11), (short)cellNo, field_subtotal2_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+12), (short)cellNo, field_120700_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+13), (short)cellNo, field_150200_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+14), (short)cellNo, field_total_rate, doubleRightStyle);
                    }
                    if(i==dbData.size()-1 && Integer.parseInt(m_year)<Integer.parseInt(endYear)){
                        for(int y=Integer.parseInt(m_year)+1; y<=Integer.parseInt(endYear);y++){
                            cellNo++;
                            rowTmp=sheet.getRow(2);
                            reportUtil.createCell(wb, rowTmp, (short)cellNo, y+"年度 資料不完整", defaultStyle);//設年度
                            reportUtil.createCell(wb, rowTmp, (short)(cellNo+1), "", defaultStyle);
                            rowTmp=sheet.getRow(3);
                            reportUtil.createCell(wb, rowTmp, (short)cellNo, "", defaultStyle);//設年度
                            reportUtil.createCell(wb, rowTmp, (short)(cellNo+1), "", defaultStyle);
                            sheet.addMergedRegion(new Region((short)2, (short)cellNo, (short)3, (short)(cellNo+1)));
                            rowTmp=sheet.getRow(4);
                            sheet.setColumnWidth((short)cellNo,(short)4500);
                            reportUtil.createCell(wb, rowTmp, (short)cellNo, "金額", rightStyle);
                            reportUtil.createCell(wb, rowTmp, (short)(cellNo+1), "%", rightStyle);
                            for(int r=5;r<=18;r++){
                                reportUtil.createCell(wb, sheet.getRow(r), (short)cellNo, "", numberRightStyle);
                                reportUtil.createCell(wb, sheet.getRow(r), (short)(cellNo+1), "", doubleRightStyle);
                            }
                            cellNo++;
                        }
                    }
                    qYear++;
    			}
			}else{
			    cellNo++;
                for(int y=Integer.parseInt(startYear); y<=Integer.parseInt(endYear);y++){
                    rowTmp=sheet.getRow(2);
                    reportUtil.createCell(wb, rowTmp, (short)cellNo, y+"年度 資料不完整", defaultStyle);//設年度
                    reportUtil.createCell(wb, rowTmp, (short)(cellNo+1), "", defaultStyle);
                    rowTmp=sheet.getRow(3);
                    reportUtil.createCell(wb, rowTmp, (short)cellNo, "", defaultStyle);//設年度
                    reportUtil.createCell(wb, rowTmp, (short)(cellNo+1), "", defaultStyle);
                    sheet.addMergedRegion(new Region((short)2, (short)cellNo, (short)3, (short)(cellNo+1)));
                    rowTmp=sheet.getRow(4);
                    sheet.setColumnWidth((short)cellNo,(short)4500);
                    reportUtil.createCell(wb, rowTmp, (short)cellNo, "金額", rightStyle);
                    reportUtil.createCell(wb, rowTmp, (short)(cellNo+1), "%", rightStyle);
                    for(int r=5;r<=18;r++){
                        reportUtil.createCell(wb, sheet.getRow(r), (short)cellNo, "", numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(r), (short)(cellNo+1), "", doubleRightStyle);
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
			System.out.println("//RptAN004WB createRpt() Have Error.....");
			//e.printStackTrace();
			System.out.println("//-------------------------------------");
		}
		System.out.println("RptAN004WB createRpt() Debug End ...");
		return errMsg;
	}
	private static void setPreparedStatementParameter(PreparedStatement pst,List paramList) throws Exception{
		for(int i = 0 ;i< paramList.size() ;i++) {
			pst.setString(i+1,(String)paramList.get(i)) ;
		}
	}
}
