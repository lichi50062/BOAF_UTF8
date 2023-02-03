/*
 * Created on 2006/10/25 by ABYSS Allen
 * AN006W 多個年度農漁會信用部簡明資產負債表
 * fixed 99.06.04 sql injection by 2808
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

public class RptAN006W {

	public static String createRpt(String startYear, String endYear, String bankType, int priceUtil) {
	    NumberFormat nf = NumberFormat.getInstance();
        nf.setMinimumFractionDigits(2);// 若小數點不足二位，則補足二位
	    String m_year =""; String field_110000 =""; String field_110000_rate =""; String field_110100 =""; String field_110100_rate ="";
	    String field_110300 =""; String field_110300_rate =""; String field_110400 =""; String field_110400_rate =""; 
	    String field_110600 =""; String field_110600_rate =""; String field_111100 =""; String field_111100_rate ="";
	    String field_110000_other =""; String field_110000_other_rate =""; String field_120000 =""; String field_120000_rate ="";
	    String field_130000 =""; String field_130000_rate =""; String field_140000 =""; String field_140000_rate ="";
	    String field_150000 =""; String field_150000_rate =""; String field_160000 =""; String field_160000_rate ="";
	    String field_190000 =""; String field_190000_rate =""; String field_210000 =""; String field_210000_rate ="";
	    String field_210400 =""; String field_210400_rate =""; String field_210700 =""; String field_210700_rate ="";
	    String field_210000_other =""; String field_210000_other_rate =""; String field_220000 =""; String field_220000_rate ="";
	    String field_240000 =""; String field_240000_rate =""; String field_250000 =""; String field_250000_rate ="";
	    String field_260000 =""; String field_260000_rate =""; String field_400000 =""; String field_400000_rate ="";
	    String field_310000 =""; String field_310000_rate =""; String field_320200 =""; String field_320200_rate ="";
	    String field_320300 =""; String field_320300_rate =""; String field_300000 =""; String field_300000_rate ="";
	    String field_600000 =""; String field_600000_rate ="";
	    String[] priceUtilStr = new String[]{"元","仟元","萬元","百萬元","仟萬元","億元"};
		System.out.println("RptAN006 createRpt() Debug Start ...");
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
			sqlCmd.append("       round(field_110000 /?,0) as  field_110000, ");   //流動資產 
			sqlCmd.append("       decode(field_190000,0,0,round(field_110000 / field_190000 *100 ,2))  as  field_110000_rate, ");// 流動資產.%
			sqlCmd.append("       round(field_110100 /?,0) as  field_110100, ");   //庫存現金
			sqlCmd.append("       decode(field_190000,0,0,round(field_110100 / field_190000 *100 ,2))  as  field_110100_rate, ");//庫存現金.%
			sqlCmd.append("       round(field_110300 /?,0) as  field_110300, ");   //存放行庫
			sqlCmd.append("       decode(field_190000,0,0,round(field_110300 / field_190000 *100 ,2))  as  field_110300_rate, ");
			sqlCmd.append("       round(field_110400 /?,0) as  field_110400, ");   //繳存存款準備
			sqlCmd.append("       decode(field_190000,0,0,round(field_110400 / field_190000 *100 ,2))  as field_110400_rate, ");
			sqlCmd.append("       round(field_110600 /?,0) as  field_110600, ");   //有價證券
			sqlCmd.append("       decode(field_190000,0,0,round(field_110600 / field_190000 *100 ,2))  as field_110600_rate, ");
			sqlCmd.append("       round(field_111100 /?,0) as  field_111100, ");   //應收利息
			sqlCmd.append("       decode(field_190000,0,0,round(field_111100 / field_190000 *100 ,2))  as field_111100_rate, ");
			sqlCmd.append("       round(field_110000_other /?,0) as  field_110000_other, ");   //其他流動資產
			sqlCmd.append("       decode(field_190000,0,0,round(field_110000_other / field_190000 *100 ,2))  as field_110000_other_rate, ");
			sqlCmd.append("       round(field_120000 /?,0) as  field_120000, ");   //放款
			sqlCmd.append("       decode(field_190000,0,0,round(field_120000 / field_190000 *100 ,2))  as field_120000_rate, ");
			sqlCmd.append("       round(field_130000 /?,0) as  field_130000, ");   //基金及出資
			sqlCmd.append("       decode(field_190000,0,0,round(field_130000 / field_190000 *100 ,2))  as field_130000_rate, ");
			sqlCmd.append("       round(field_140000 /?,0) as  field_140000, ");   //固定資產淨額
			sqlCmd.append("       decode(field_190000,0,0,round(field_140000 / field_190000 *100 ,2))  as field_140000_rate, ");
			sqlCmd.append("       round(field_150000 /?,0) as  field_150000, ");   //其他資產
			sqlCmd.append("       decode(field_190000,0,0,round(field_150000 / field_190000 *100 ,2))  as field_150000_rate, ");
			sqlCmd.append("       round(field_160000 /?,0) as  field_160000, ");   //往來
			sqlCmd.append("       decode(field_190000,0,0,round(field_160000 / field_190000 *100 ,2))  as field_160000_rate, ");
			sqlCmd.append("       round(field_190000 /?,0) as  field_190000, ");   //資產總計
			sqlCmd.append("       decode(field_190000,0,0,round(field_190000 / field_190000 *100 ,2))  as field_190000_rate, ");
			sqlCmd.append("       round(field_210000 /?,0) as  field_210000, ");   //流動負債
			sqlCmd.append("       decode(field_600000,0,0,round(field_210000 / field_600000 *100 ,2))  as  field_210000_rate, ");
			sqlCmd.append("       round(field_210400 /?,0) as  field_210400, ");   //短期借款
			sqlCmd.append("       decode(field_600000,0,0,round(field_210400 / field_600000 *100 ,2))  as  field_210400_rate, ");
			sqlCmd.append("       round(field_210700 /?,0) as  field_210700, ");   //應付利息
			sqlCmd.append("       decode(field_600000,0,0,round(field_210700 / field_600000 *100 ,2))  as  field_210700_rate, ");
			sqlCmd.append("       round(field_210000_other /?,0) as  field_210000_other, ");   //其他流動負債
			sqlCmd.append("       decode(field_600000,0,0,round(field_210000_other / field_600000 *100 ,2))  as  field_210000_other_rate, ");
			sqlCmd.append("       round(field_220000 /?,0) as  field_220000, ");   //存款
			sqlCmd.append("       decode(field_600000,0,0,round(field_220000 / field_600000 *100 ,2))  as  field_220000_rate, ");
			sqlCmd.append("       round(field_240000 /?,0) as  field_240000, ");   //長期負債
			sqlCmd.append("       decode(field_600000,0,0,round(field_240000 / field_600000 *100 ,2))  as  field_240000_rate, ");
			sqlCmd.append("       round(field_250000 /?,0) as  field_250000, ");   //其他負債
			sqlCmd.append("       decode(field_600000,0,0,round(field_250000 / field_600000 *100 ,2))  as  field_250000_rate, ");
			sqlCmd.append("       round(field_260000 /?,0) as  field_260000, ");   //往來
			sqlCmd.append("       decode(field_600000,0,0,round(field_260000 / field_600000 *100 ,2))  as  field_260000_rate, ");
			sqlCmd.append("       round(field_400000 /?,0) as  field_400000, ");   //負債總計
			sqlCmd.append("       decode(field_600000,0,0,round(field_400000 / field_600000 *100 ,2))  as field_400000_rate, ");
			sqlCmd.append("       round(field_310000 /?,0) as  field_310000, ");   //事業資金及公積
			sqlCmd.append("       decode(field_600000,0,0,round(field_310000 / field_600000 *100 ,2))  as field_310000_rate, ");
			sqlCmd.append("       round(field_320200 /?,0) as  field_320200, ");   //前期損益
			sqlCmd.append("       decode(field_600000,0,0,round(field_320200 / field_600000 *100 ,2))  as field_320200_rate, ");
			sqlCmd.append("       round(field_320300 /?,0) as  field_320300, ");   //本期損益
			sqlCmd.append("       decode(field_600000,0,0,round(field_320300 / field_600000 *100 ,2))  as field_320300_rate, ");
			sqlCmd.append("       round(field_300000 /?,0) as  field_300000, ");   //淨值總計
			sqlCmd.append("       decode(field_600000,0,0,round(field_300000 / field_600000 *100 ,2))  as field_300000_rate, ");
			sqlCmd.append("       round(field_600000 /?,0) as  field_600000, ");   //負債及淨值合計
			sqlCmd.append("       decode(field_600000,0,0,round(field_600000 / field_600000 *100 ,2))  as field_600000_rate ");
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
			sqlCmd.append("              sum(field_110000) as field_110000, ");
			sqlCmd.append("              sum(field_110100) as field_110100, ");
			sqlCmd.append("              sum(field_110300) as field_110300, ");
			sqlCmd.append("              sum(field_110400) as field_110400, ");
			sqlCmd.append("              sum(field_110600) as field_110600, ");
			sqlCmd.append("              sum(field_111100) as field_111100, ");
			sqlCmd.append("              sum(field_110000)+sum(field_110800)-sum(field_110100)-sum(field_110300)-sum(field_110400)-sum(field_110600)-sum(field_111100) as field_110000_other, ");
			sqlCmd.append("              sum(field_120000) as field_120000, ");
			sqlCmd.append("              sum(field_130000) as field_130000, ");
			sqlCmd.append("              sum(field_140000) as field_140000, ");
			sqlCmd.append("              sum(field_150000) as field_150000, ");
			sqlCmd.append("              sum(field_160000) as field_160000, ");
			sqlCmd.append("              sum(field_190000) as field_190000, ");
			sqlCmd.append("              sum(field_210000) as field_210000, ");              
			sqlCmd.append("              sum(field_210400) as field_210400, ");
			sqlCmd.append("              sum(field_210700) as field_210700, ");
			sqlCmd.append("              sum(field_210000)-sum(field_210400)-sum(field_210700) as field_210000_other, ");
			sqlCmd.append("              sum(field_220000) as field_220000, ");
			sqlCmd.append("              sum(field_240000) as field_240000, ");
			sqlCmd.append("              sum(field_250000) as field_250000, ");
			sqlCmd.append("              sum(field_260000) as field_260000, ");
			sqlCmd.append("              sum(field_400000) as field_400000, ");
			sqlCmd.append("              sum(field_310000) as field_310000, ");
			sqlCmd.append("              sum(field_320200) as field_320200, ");
			sqlCmd.append("              sum(field_320300) as field_320300, ");
			sqlCmd.append("              sum(field_300000) as field_300000, ");
			sqlCmd.append("              sum(field_600000) as field_600000 ");
			sqlCmd.append("       from ( ");
			sqlCmd.append("            select a01.m_year, ");
			sqlCmd.append("                      sum(decode(a01.acc_code,'110000',amt,0))  as  field_110000, ");
			sqlCmd.append("                      sum(decode(a01.acc_code,'110100',amt,0))  as  field_110100, ");
			sqlCmd.append("                      decode(YEAR_TYPE,'102',decode(bn01.bank_type,'6',sum(decode(a01.acc_code,'110300',amt,0)),'7',sum(decode(a01.acc_code,'110200',amt,0)),0),'103',sum(decode(a01.acc_code,'110300',amt,0)),0)  as  field_110300, ");
			sqlCmd.append("                      sum(decode(a01.acc_code,'110400',amt,0))  as  field_110400, ");
			sqlCmd.append("                      decode(bn01.bank_type,'6',sum(decode(a01.acc_code,'110600',amt,0))-sum(decode(a01.acc_code,'112600',amt,0)),'7',decode(YEAR_TYPE,'102',sum(decode(a01.acc_code,'110600',amt,0))-sum(decode(a01.acc_code,'110700',amt,0)),'103',sum(decode(a01.acc_code,'110600',amt,0))-sum(decode(a01.acc_code,'112300',amt,0)),0),0) as field_110600, ");
			sqlCmd.append("                      decode(YEAR_TYPE,'102',decode(bn01.bank_type,'6',sum(decode(a01.acc_code,'111100',amt,0)),'7',sum(decode(a01.acc_code,'111200',amt,0)),0),'103',sum(decode(a01.acc_code,'111100',amt,0)),0)  as  field_111100, ");
			sqlCmd.append("                      decode(YEAR_TYPE,'102',decode(bn01.bank_type,'6',sum(decode(a01.acc_code,'110800',amt,0))+sum(decode(a01.acc_code,'111000',amt,0)),'7',sum(decode(a01.acc_code,'111100',amt,0)),0),'103',sum(decode(a01.acc_code,'110800',amt,0))+sum(decode(a01.acc_code,'111000',amt,0)),0) as field_110800, ");
			sqlCmd.append("                      sum(decode(a01.acc_code,'120000',amt,0))  as  field_120000, ");
			sqlCmd.append("                      sum(decode(a01.acc_code,'130000',amt,0))  as  field_130000, ");
			sqlCmd.append("                      sum(decode(a01.acc_code,'140000',amt,0))  as  field_140000, ");
			sqlCmd.append("                      sum(decode(a01.acc_code,'150000',amt,0))  as  field_150000, ");
			sqlCmd.append("                      sum(decode(a01.acc_code,'160000',amt,0))  as  field_160000, ");
			sqlCmd.append("                      decode(bn01.bank_type,'6',sum(decode(a01.acc_code,'190000',amt,0)),'7',sum(decode(a01.acc_code,'100000',amt,0)),0) as field_190000, ");
			sqlCmd.append("                      sum(decode(a01.acc_code,'210000',amt,0))  as  field_210000, ");
			sqlCmd.append("                      decode(YEAR_TYPE,'102',decode(bn01.bank_type,'6',sum(decode(a01.acc_code,'210400',amt,0)),'7',sum(decode(a01.acc_code,'210200',amt,0)),0),'103',sum(decode(a01.acc_code,'210400',amt,0)),0)  as  field_210400, ");
			sqlCmd.append("                      sum(decode(a01.acc_code,'210700',amt,0))  as  field_210700, ");
			sqlCmd.append("                      sum(decode(a01.acc_code,'220000',amt,0))  as  field_220000, ");
			sqlCmd.append("                      sum(decode(a01.acc_code,'240000',amt,0))  as  field_240000, ");
			sqlCmd.append("                      sum(decode(a01.acc_code,'250000',amt,0))  as  field_250000, ");
			sqlCmd.append("                      sum(decode(a01.acc_code,'260000',amt,0))  as  field_260000, ");
			sqlCmd.append("                      decode(bn01.bank_type,'6',sum(decode(a01.acc_code,'400000',amt,0))-sum(decode(a01.acc_code,'310000',amt,0))-sum(decode(a01.acc_code,'320000',amt,0)),'7',sum(decode(a01.acc_code,'200000',amt,0)),0) as field_400000, ");
			sqlCmd.append("                      sum(decode(a01.acc_code,'310000',amt,0))  as  field_310000, ");
			sqlCmd.append("                      sum(decode(a01.acc_code,'320200',amt,0))  as  field_320200, ");
			sqlCmd.append("                      sum(decode(a01.acc_code,'320300',amt,0))  as  field_320300, ");
			sqlCmd.append("                      decode(bn01.bank_type,'6',sum(decode(a01.acc_code,'310000',amt,0))+sum(decode(a01.acc_code,'320000',amt,0)),'7',sum(decode(a01.acc_code,'300000',amt,0)),0) as field_300000, ");
			sqlCmd.append("                      decode(bn01.bank_type,'6',sum(decode(a01.acc_code,'400000',amt,0)),'7',sum(decode(a01.acc_code,'600000',amt,0)),0) as field_600000 ");
			sqlCmd.append("               from (select  (CASE WHEN (a01.m_year <= 102) THEN '102' ");
			sqlCmd.append("                                   WHEN (a01.m_year > 102) THEN '103' ");
			sqlCmd.append("                              ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01  ");
			sqlCmd.append("                              where m_year in (? ");
            paramList.add(startYear) ;
            for(int y=Integer.parseInt(startYear)+1; y<=Integer.parseInt(endYear);y++){
                sqlCmd.append(",?" );    
                paramList.add(y) ;
            }
			sqlCmd.append("                              ) and m_month=12 and a01.acc_code in ('110000','110100','110200','110300','110400','110600','110700','112300','112600','111100','111200','110800','111000','120000','130000','140000','150000','160000','190000','100000','210000','210400','210200','210700','220000','240000','250000','260000','400000','310000','320000','200000','310000','320200','320300','310000','320000','300000','600000')  ");                     
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
			sqlCmd.append("               where a01.bank_code = bn01.bank_no  ");            
			sqlCmd.append("               group by a01.m_year,bn01.bank_type,YEAR_TYPE ");
			sqlCmd.append("               )a01 ");
			sqlCmd.append("       group by m_year ");
			sqlCmd.append(")  ");  
			List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,
			        "m_year,field_110000,field_110000_rate,field_110100,field_110100_rate,field_110300,field_110300_rate,"+
			        "field_110400,field_110400_rate,field_110600,field_110600_rate,field_111100,field_111100_rate,"+
			        "field_110000_other,field_110000_other_rate,field_120000,field_120000_rate,field_130000,field_130000_rate,"+
			        "field_140000,field_140000_rate,field_150000,field_150000_rate,field_160000,field_160000_rate,"+
			        "field_190000,field_190000_rate,field_210000,field_210000_rate,field_210400,field_210400_rate,"+
			        "field_210700,field_210700_rate,field_210000_other,field_210000_other_rate,field_220000,field_220000_rate,"+
			        "field_240000,field_240000_rate,field_250000,field_250000_rate,field_260000,field_260000_rate,"+
			        "field_400000,field_400000_rate,ield_310000,field_310000_rate,field_320200,field_320200_rate,"+
			        "field_320300,field_320300_rate,field_300000,field_300000_rate,field_600000,field_600000_rate");
            System.out.println("dbData.size=" + dbData.size());
			//====開始製作Excel檔案
			HSSFWorkbook wb = new HSSFWorkbook();
			HSSFSheet sheet = wb.createSheet("New Sheet 1");
			HSSFPrintSetup ps = sheet.getPrintSetup(); //取得列印設定
			//設定頁面符合列印大小
            sheet.setZoom(75, 100); // 螢幕上看到的縮放大小
            sheet.setAutobreaks(false); //自動分頁
            ps.setScale((short)74); //列印縮放百分比
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
			HSSFCellStyle leftStyle = reportUtil.getLeftStyle(wb);
			HSSFFont subTitleFont = wb.createFont();
			subTitleFont.setFontHeightInPoints((short)14);
			subTitleFont.setBoldweight((short)30);
			leftStyle.setFont(subTitleFont);
			HSSFCellStyle custLeftStyle = reportUtil.getNoBorderLeftStyle(wb);
			custLeftStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
			custLeftStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
			custLeftStyle.setFont(defaultFont);
			HSSFCellStyle numberRightStyle = reportUtil.getRightStyle(wb);//有框整數值置右
			numberRightStyle.setDataFormat(format.getFormat("#,##0"));
			numberRightStyle.setFont(numberFont);
			HSSFCellStyle doubleRightStyle = reportUtil.getRightStyle(wb);//有框小數點置右
			doubleRightStyle.setDataFormat(format.getFormat("#,##0.00"));
			doubleRightStyle.setFont(numberFont);
			HSSFCellStyle noBorderRightStyle = reportUtil.getNoBoderStyle(wb);
			noBorderRightStyle.setFont(defaultFont);
			reportUtil.setDefaultStyle(defaultStyle);
			sheet.setColumnWidth((short)0,(short)3500);
			sheet.setColumnWidth((short)1,(short)3500);
			//設定報表表頭資料 開始============================================
			row=sheet.createRow(0);
			String titleStr="";
			if(bankType.equals("ALL")){
				titleStr = "多個年度農漁會信用部簡明資產負債表";
			}else if(bankType.equals("6")){
				titleStr = "多個年度農會信用部簡明資產負債表";
			}else{
				titleStr = "多個年度漁會信用部簡明資產負債表";
			}
			int qAllYear = Integer.parseInt(endYear)-Integer.parseInt(startYear)+1;
            int repeatTime = 1;
            if(qAllYear>5 && (qAllYear%5)>0){
				repeatTime++;
            }
			for(int ri=0;ri<repeatTime;ri++){
				int startIndex=0;
				if(ri>0){
					startIndex=(ri*10)+2;
				}
				reportUtil.createCell(wb, row, (short)startIndex, titleStr, titleStyle);
				int endIndex=((ri+1)*10)+1;
				for(int ci=(startIndex+1); ci<endIndex;ci++){
					reportUtil.createCell( wb, row, (short)ci, "", titleStyle);
				}
				sheet.addMergedRegion(new Region((short)0, (short)startIndex, (short)0, (short)endIndex));
			}
			row=sheet.createRow(1);
			row.setHeightInPoints(18.0F);
			for(int ri=0;ri<repeatTime;ri++){
				int startIndex=0;
				if(ri>0){
					startIndex=(ri*10)+2;
				}
				reportUtil.createCell(wb, row, (short)startIndex, "單位：新台幣 "+priceUtilStr[priceIndex], noBorderRightStyle);
				int endIndex=((ri+1)*10)+1;
				for(int ci=(startIndex+1); ci<endIndex;ci++){
					reportUtil.createCell( wb, row, (short)ci, "", titleStyle);
				}
				sheet.addMergedRegion(new Region((short)1, (short)startIndex, (short)1, (short)endIndex));
			}
			row=sheet.createRow(2);
			row.setHeightInPoints(18.0F);
			reportUtil.createCell(wb, row, (short)0, "年度別", rightStyle);
			reportUtil.createCell(wb, row, (short)1, "", rightStyle);
			sheet.addMergedRegion(new Region((short)2, (short)0, (short)2, (short)1));
			row=sheet.createRow(3);
			row.setHeightInPoints(18.0F);
			reportUtil.createCell(wb, row, (short)0, "項目", leftStyle);
			reportUtil.createCell(wb, row, (short)1, "", leftStyle);
			sheet.addMergedRegion(new Region((short)3, (short)0, (short)3, (short)1));
			row=sheet.createRow(4);
			row.setHeightInPoints(18.0F);
			reportUtil.createCell(wb, row, (short)0, "  流動資產", custLeftStyle);
			reportUtil.createCell(wb, row, (short)1, "", custLeftStyle);
			sheet.addMergedRegion(new Region((short)4, (short)0, (short)4, (short)1));
			row=sheet.createRow(5);
			row.setHeightInPoints(18.0F);
			reportUtil.createCell(wb, row, (short)0, "    庫存現金", custLeftStyle);
			reportUtil.createCell(wb, row, (short)1, "", custLeftStyle);
			sheet.addMergedRegion(new Region((short)5, (short)0, (short)5, (short)1));
			row=sheet.createRow(6);
			row.setHeightInPoints(18.0F);
			reportUtil.createCell(wb, row, (short)0, "    存放行庫", custLeftStyle);
			reportUtil.createCell(wb, row, (short)1, "", custLeftStyle);
			sheet.addMergedRegion(new Region((short)6, (short)0, (short)6, (short)1));
			row=sheet.createRow(7);
			row.setHeightInPoints(18.0F);
			reportUtil.createCell(wb, row, (short)0, "    繳存存款準備", custLeftStyle);
			reportUtil.createCell(wb, row, (short)1, "", custLeftStyle);
			sheet.addMergedRegion(new Region((short)7, (short)0, (short)7, (short)1));
			row=sheet.createRow(8);
			row.setHeightInPoints(18.0F);
			reportUtil.createCell(wb, row, (short)0, "    有價證卷", custLeftStyle);
			reportUtil.createCell(wb, row, (short)1, "", custLeftStyle);
			sheet.addMergedRegion(new Region((short)8, (short)0, (short)8, (short)1));
			row=sheet.createRow(9);
			row.setHeightInPoints(18.0F);
			reportUtil.createCell(wb, row, (short)0, "    應收利息", custLeftStyle);
			reportUtil.createCell(wb, row, (short)1, "", custLeftStyle);
			sheet.addMergedRegion(new Region((short)9, (short)0, (short)9, (short)1));
			row=sheet.createRow(10);
			row.setHeightInPoints(18.0F);
			reportUtil.createCell(wb, row, (short)0, "    其他流動資產", custLeftStyle);
			reportUtil.createCell(wb, row, (short)1, "", custLeftStyle);
			sheet.addMergedRegion(new Region((short)10, (short)0, (short)10, (short)1));
			row=sheet.createRow(11);
			row.setHeightInPoints(18.0F);
			reportUtil.createCell(wb, row, (short)0, "  放  款", custLeftStyle);
			reportUtil.createCell(wb, row, (short)1, "", custLeftStyle);
			sheet.addMergedRegion(new Region((short)11, (short)0, (short)11, (short)1));
			row=sheet.createRow(12);
			row.setHeightInPoints(18.0F);
			reportUtil.createCell(wb, row, (short)0, "  基金及出資", custLeftStyle);
			reportUtil.createCell(wb, row, (short)1, "", custLeftStyle);
			sheet.addMergedRegion(new Region((short)12, (short)0, (short)12, (short)1));
			row=sheet.createRow(13);
			row.setHeightInPoints(18.0F);
			reportUtil.createCell(wb, row, (short)0, "  固定資產淨額", custLeftStyle);
			reportUtil.createCell(wb, row, (short)1, "", custLeftStyle);
			sheet.addMergedRegion(new Region((short)13, (short)0, (short)13, (short)1));
			row=sheet.createRow(14);
			row.setHeightInPoints(18.0F);
			reportUtil.createCell(wb, row, (short)0, "  其他資產", custLeftStyle);
			reportUtil.createCell(wb, row, (short)1, "", custLeftStyle);
			sheet.addMergedRegion(new Region((short)14, (short)0, (short)14, (short)1));
			row=sheet.createRow(15);
			row.setHeightInPoints(18.0F);
			reportUtil.createCell(wb, row, (short)0, "  往  來", custLeftStyle);
			reportUtil.createCell(wb, row, (short)1, "", custLeftStyle);
			sheet.addMergedRegion(new Region((short)15, (short)0, (short)15, (short)1));
			row=sheet.createRow(16);
			row.setHeightInPoints(18.0F);
			reportUtil.createCell(wb, row, (short)0, "資產總計", leftStyle);
			reportUtil.createCell(wb, row, (short)1, "", leftStyle);
			sheet.addMergedRegion(new Region((short)16, (short)0, (short)16, (short)1));
			row=sheet.createRow(17);
			row.setHeightInPoints(18.0F);
			reportUtil.createCell(wb, row, (short)0, "  流動負債", custLeftStyle);
			reportUtil.createCell(wb, row, (short)1, "", custLeftStyle);
			sheet.addMergedRegion(new Region((short)17, (short)0, (short)17, (short)1));
			row=sheet.createRow(18);
			row.setHeightInPoints(18.0F);
			reportUtil.createCell(wb, row, (short)0, "    短期借款", custLeftStyle);
			reportUtil.createCell(wb, row, (short)1, "", custLeftStyle);
			sheet.addMergedRegion(new Region((short)18, (short)0, (short)18, (short)1));
			row=sheet.createRow(19);
			row.setHeightInPoints(18.0F);
			reportUtil.createCell(wb, row, (short)0, "    應付利息", custLeftStyle);
			reportUtil.createCell(wb, row, (short)1, "", custLeftStyle);
			sheet.addMergedRegion(new Region((short)19, (short)0, (short)19, (short)1));
			row=sheet.createRow(20);
			row.setHeightInPoints(18.0F);
			reportUtil.createCell(wb, row, (short)0, "    其他流動負債", custLeftStyle);
			reportUtil.createCell(wb, row, (short)1, "", custLeftStyle);
			sheet.addMergedRegion(new Region((short)20, (short)0, (short)20, (short)1));
			row=sheet.createRow(21);
			row.setHeightInPoints(18.0F);
			reportUtil.createCell(wb, row, (short)0, "  存款", custLeftStyle);
			reportUtil.createCell(wb, row, (short)1, "", custLeftStyle);
			sheet.addMergedRegion(new Region((short)21, (short)0, (short)21, (short)1));
			row=sheet.createRow(22);
			row.setHeightInPoints(18.0F);
			reportUtil.createCell(wb, row, (short)0, "  長期負債", custLeftStyle);
			reportUtil.createCell(wb, row, (short)1, "", custLeftStyle);
			sheet.addMergedRegion(new Region((short)22, (short)0, (short)22, (short)1));
			row=sheet.createRow(23);
			row.setHeightInPoints(18.0F);
			reportUtil.createCell(wb, row, (short)0, "  其他負債", custLeftStyle);
			reportUtil.createCell(wb, row, (short)1, "", custLeftStyle);
			sheet.addMergedRegion(new Region((short)23, (short)0, (short)23, (short)1));
			row=sheet.createRow(24);
			row.setHeightInPoints(18.0F);
			reportUtil.createCell(wb, row, (short)0, "  往  來", custLeftStyle);
			reportUtil.createCell(wb, row, (short)1, "", custLeftStyle);
			sheet.addMergedRegion(new Region((short)24, (short)0, (short)24, (short)1));
			row=sheet.createRow(25);
			row.setHeightInPoints(18.0F);
			reportUtil.createCell(wb, row, (short)0, "負債總計", leftStyle);
			reportUtil.createCell(wb, row, (short)1, "", leftStyle);
			sheet.addMergedRegion(new Region((short)25, (short)0, (short)25, (short)1));
			row=sheet.createRow(26);
			row.setHeightInPoints(18.0F);
			reportUtil.createCell(wb, row, (short)0, "  事業資金及公積", custLeftStyle);
			reportUtil.createCell(wb, row, (short)1, "", custLeftStyle);
			sheet.addMergedRegion(new Region((short)26, (short)0, (short)26, (short)1));
			row=sheet.createRow(27);
			row.setHeightInPoints(18.0F);
			reportUtil.createCell(wb, row, (short)0, "  其他公積", custLeftStyle);
			reportUtil.createCell(wb, row, (short)1, "", custLeftStyle);
			sheet.addMergedRegion(new Region((short)27, (short)0, (short)27, (short)1));
			row=sheet.createRow(28);
			row.setHeightInPoints(18.0F);
			reportUtil.createCell(wb, row, (short)0, "  前期損益", custLeftStyle);
			reportUtil.createCell(wb, row, (short)1, "", custLeftStyle);
			sheet.addMergedRegion(new Region((short)28, (short)0, (short)28, (short)1));
			row=sheet.createRow(29);
			row.setHeightInPoints(18.0F);
			reportUtil.createCell(wb, row, (short)0, "  本期損益", custLeftStyle);
			reportUtil.createCell(wb, row, (short)1, "", custLeftStyle);
			sheet.addMergedRegion(new Region((short)29, (short)0, (short)29, (short)1));
			row=sheet.createRow(30);
			row.setHeightInPoints(18.0F);
			reportUtil.createCell(wb, row, (short)0, "淨值總計", leftStyle);
			reportUtil.createCell(wb, row, (short)1, "", leftStyle);
			sheet.addMergedRegion(new Region((short)30, (short)0, (short)30, (short)1));
			row=sheet.createRow(31);
			row.setHeightInPoints(18.0F);
			reportUtil.createCell(wb, row, (short)0, "負債及淨值合計", leftStyle);
			reportUtil.createCell(wb, row, (short)1, "", leftStyle);
			sheet.addMergedRegion(new Region((short)31, (short)0, (short)31, (short)1));
			wb.setRepeatingRowsAndColumns(0, 0, 1, 2, 30);//設為固定表頭(第幾個sheet,起始欄,終止欄,起始列,終止列)
			//設定報表表頭資料 結束============================================
//			設定儲存格資料============================================
			int rowNo=3;
            int cellNo=1;
            DataObject bean = null;
            HSSFRow rowTmp=sheet.getRow(2);
            int qYear = Integer.parseInt(startYear);
            if(dbData.size()>0){
                for(int i=0;i<dbData.size();i++){  
                    bean = (DataObject)dbData.get(i);
                    m_year = (bean.getValue("m_year")==null)?"":(bean.getValue("m_year")).toString();
                    field_110000 = (bean.getValue("field_110000")==null)?"0":Utility.setCommaFormat((bean.getValue("field_110000")).toString());
                    field_110000_rate = (bean.getValue("field_110000_rate")==null)?"0.00":nf.format(bean.getValue("field_110000_rate"));
                    field_110100 = (bean.getValue("field_110100")==null)?"0":Utility.setCommaFormat((bean.getValue("field_110100")).toString());
                    field_110100_rate = (bean.getValue("field_110100_rate")==null)?"0.00":nf.format(bean.getValue("field_110100_rate"));
                    field_110300 = (bean.getValue("field_110300")==null)?"0":Utility.setCommaFormat((bean.getValue("field_110300")).toString());
                    field_110300_rate = (bean.getValue("field_110300_rate")==null)?"0.00":nf.format(bean.getValue("field_110300_rate"));
                    field_110400 = (bean.getValue("field_110400")==null)?"0":Utility.setCommaFormat((bean.getValue("field_110400")).toString());
                    field_110400_rate = (bean.getValue("field_110400_rate")==null)?"0.00":nf.format(bean.getValue("field_110400_rate"));
                    field_110600 = (bean.getValue("field_110600")==null)?"0":Utility.setCommaFormat((bean.getValue("field_110600")).toString());
                    field_110600_rate = (bean.getValue("field_110600_rate")==null)?"0.00":nf.format(bean.getValue("field_110600_rate"));
                    field_111100 = (bean.getValue("field_111100")==null)?"0":Utility.setCommaFormat((bean.getValue("field_111100")).toString());
                    field_111100_rate = (bean.getValue("field_111100_rate")==null)?"0.00":nf.format(bean.getValue("field_111100_rate"));
                    field_110000_other = (bean.getValue("field_110000_other")==null)?"0":Utility.setCommaFormat((bean.getValue("field_110000_other")).toString());
                    field_110000_other_rate = (bean.getValue("field_110000_other_rate")==null)?"0.00":nf.format(bean.getValue("field_110000_other_rate"));
                    field_120000 = (bean.getValue("field_120000")==null)?"0":Utility.setCommaFormat((bean.getValue("field_120000")).toString());
                    field_120000_rate = (bean.getValue("field_120000_rate")==null)?"0.00":nf.format(bean.getValue("field_120000_rate"));
                    field_130000 = (bean.getValue("field_130000")==null)?"0":Utility.setCommaFormat((bean.getValue("field_130000")).toString());
                    field_130000_rate = (bean.getValue("field_130000_rate")==null)?"0.00":nf.format(bean.getValue("field_130000_rate"));
                    field_140000 = (bean.getValue("field_140000")==null)?"0":Utility.setCommaFormat((bean.getValue("field_140000")).toString());
                    field_140000_rate = (bean.getValue("field_140000_rate")==null)?"0.00":nf.format(bean.getValue("field_140000_rate"));
                    field_150000 = (bean.getValue("field_150000")==null)?"0":Utility.setCommaFormat((bean.getValue("field_150000")).toString());
                    field_150000_rate = (bean.getValue("field_150000_rate")==null)?"0.00":nf.format(bean.getValue("field_150000_rate"));
                    field_160000 = (bean.getValue("field_160000")==null)?"0":Utility.setCommaFormat((bean.getValue("field_160000")).toString());
                    field_160000_rate = (bean.getValue("field_160000_rate")==null)?"0.00":nf.format(bean.getValue("field_160000_rate"));
                    field_190000 = (bean.getValue("field_190000")==null)?"0":Utility.setCommaFormat((bean.getValue("field_190000")).toString());
                    field_190000_rate = (bean.getValue("field_190000_rate")==null)?"0.00":nf.format(bean.getValue("field_190000_rate"));
                    field_210000 = (bean.getValue("field_210000")==null)?"0":Utility.setCommaFormat((bean.getValue("field_210000")).toString());
                    field_210000_rate = (bean.getValue("field_210000_rate")==null)?"0.00":nf.format(bean.getValue("field_210000_rate"));
                    field_210400 = (bean.getValue("field_210400")==null)?"0":Utility.setCommaFormat((bean.getValue("field_210400")).toString());
                    field_210400_rate = (bean.getValue("field_210400_rate")==null)?"0.00":nf.format(bean.getValue("field_210400_rate"));
                    field_210700 = (bean.getValue("field_210700")==null)?"0":Utility.setCommaFormat((bean.getValue("field_210700")).toString());
                    field_210700_rate = (bean.getValue("field_210700_rate")==null)?"0.00":nf.format(bean.getValue("field_210700_rate"));
                    field_210000_other = (bean.getValue("field_210000_other")==null)?"0":Utility.setCommaFormat((bean.getValue("field_210000_other")).toString());
                    field_210000_other_rate = (bean.getValue("field_210000_other_rate")==null)?"0.00":nf.format(bean.getValue("field_210000_other_rate"));
                    field_220000 = (bean.getValue("field_220000")==null)?"0":Utility.setCommaFormat((bean.getValue("field_220000")).toString());
                    field_220000_rate = (bean.getValue("field_220000_rate")==null)?"0.00":nf.format(bean.getValue("field_220000_rate"));
                    field_240000 = (bean.getValue("field_240000")==null)?"0":Utility.setCommaFormat((bean.getValue("field_240000")).toString());
                    field_240000_rate = (bean.getValue("field_240000_rate")==null)?"0.00":nf.format(bean.getValue("field_240000_rate"));
                    field_250000 = (bean.getValue("field_250000")==null)?"0":Utility.setCommaFormat((bean.getValue("field_250000")).toString());
                    field_250000_rate = (bean.getValue("field_250000_rate")==null)?"0.00":nf.format(bean.getValue("field_250000_rate"));
                    field_260000 = (bean.getValue("field_260000")==null)?"0":Utility.setCommaFormat((bean.getValue("field_260000")).toString());
                    field_260000_rate = (bean.getValue("field_260000_rate")==null)?"0.00":nf.format(bean.getValue("field_260000_rate"));
                    field_400000 = (bean.getValue("field_400000")==null)?"0":Utility.setCommaFormat((bean.getValue("field_400000")).toString());
                    field_400000_rate = (bean.getValue("field_400000_rate")==null)?"0.00":nf.format(bean.getValue("field_400000_rate"));
                    field_310000 = (bean.getValue("field_310000")==null)?"0":Utility.setCommaFormat((bean.getValue("field_310000")).toString());
                    field_310000_rate = (bean.getValue("field_310000_rate")==null)?"0.00":nf.format(bean.getValue("field_310000_rate"));
                    field_320200 = (bean.getValue("field_320200")==null)?"0":Utility.setCommaFormat((bean.getValue("field_320200")).toString());
                    field_320200_rate = (bean.getValue("field_320200_rate")==null)?"0.00":nf.format(bean.getValue("field_320200_rate"));
                    field_320300 = (bean.getValue("field_320300")==null)?"0":Utility.setCommaFormat((bean.getValue("field_320300")).toString());
                    field_320300_rate = (bean.getValue("field_320300_rate")==null)?"0.00":nf.format(bean.getValue("field_320300_rate"));
                    field_300000 = (bean.getValue("field_300000")==null)?"0":Utility.setCommaFormat((bean.getValue("field_300000")).toString());
                    field_300000_rate = (bean.getValue("field_300000_rate")==null)?"0.00":nf.format(bean.getValue("field_300000_rate"));
                    field_600000 = (bean.getValue("field_600000")==null)?"0":Utility.setCommaFormat((bean.getValue("field_600000")).toString());
                    field_600000_rate = (bean.getValue("field_600000_rate")==null)?"0.00":nf.format(bean.getValue("field_600000_rate"));
                    row=sheet.getRow(rowNo);
                    cellNo++;
                    //
                    if(i==0 && Integer.parseInt(m_year)>Integer.parseInt(startYear)){
                        for(int y=Integer.parseInt(startYear); y<Integer.parseInt(m_year);y++){
                            rowTmp=sheet.getRow(2);
                            reportUtil.createCell(wb, rowTmp, (short)cellNo, y+"年度 資料不完整", defaultStyle);//設年度
                            reportUtil.createCell(wb, rowTmp, (short)(cellNo+1), "", defaultStyle);
                            sheet.addMergedRegion(new Region((short)2, (short)cellNo, (short)2, (short)(cellNo+1)));
                            rowTmp=sheet.getRow(3);
                            sheet.setColumnWidth((short)cellNo,(short)4500);
                            reportUtil.createCell(wb, rowTmp, (short)cellNo, "金額", rightStyle);
                            reportUtil.createCell(wb, rowTmp, (short)(cellNo+1), "%", rightStyle);
                            for(int r=4;r<=31;r++){
                                reportUtil.createCell(wb, sheet.getRow(r), (short)cellNo, "", numberRightStyle);
                                reportUtil.createCell(wb, sheet.getRow(r), (short)(cellNo+1), "", doubleRightStyle);
                            }
                            cellNo+=2;
                            qYear++;
                        }
                    }
                    //
                    if(Integer.parseInt(m_year)!=qYear){
                        rowTmp=sheet.getRow(2);
                        reportUtil.createCell(wb, rowTmp, (short)cellNo, qYear+"年度 資料不完整", defaultStyle);//設年度
                        reportUtil.createCell(wb, rowTmp, (short)(cellNo+1), "", defaultStyle);
                        sheet.addMergedRegion(new Region((short)2, (short)cellNo, (short)2, (short)(cellNo+1)));
                        rowTmp=sheet.getRow(3);
                        sheet.setColumnWidth((short)cellNo,(short)4500);
                        reportUtil.createCell(wb, rowTmp, (short)cellNo, "金額", rightStyle);
                        reportUtil.createCell(wb, rowTmp, (short)(cellNo+1), "%", rightStyle);
                        for(int r=4;r<=31;r++){
                            reportUtil.createCell(wb, sheet.getRow(r), (short)cellNo, "", numberRightStyle);
                            reportUtil.createCell(wb, sheet.getRow(r), (short)(cellNo+1), "", doubleRightStyle);
                        }
                        cellNo++;
                    }else{
                        rowTmp=sheet.getRow(2);
                        reportUtil.createCell(wb, rowTmp, (short)cellNo, m_year+"年度", defaultStyle);//設年度
                        reportUtil.createCell(wb, rowTmp, (short)(cellNo+1), "", defaultStyle);
                        sheet.addMergedRegion(new Region((short)2, (short)cellNo, (short)2, (short)(cellNo+1)));
                        rowTmp=sheet.getRow(3);
                        sheet.setColumnWidth((short)cellNo,(short)4500);
                        reportUtil.createCell(wb, rowTmp, (short)cellNo, "金額", rightStyle);
                        reportUtil.createCell(wb, rowTmp, (short)(cellNo+1), "%", rightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+1), (short)cellNo, field_110000, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+2), (short)cellNo, field_110100, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+3), (short)cellNo, field_110300, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+4), (short)cellNo, field_110400, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+5), (short)cellNo, field_110600, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+6), (short)cellNo, field_111100, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+7), (short)cellNo, field_110000_other, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+8), (short)cellNo, field_120000, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+9), (short)cellNo, field_130000, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+10), (short)cellNo, field_140000, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+11), (short)cellNo, field_150000, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+12), (short)cellNo, field_160000, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+13), (short)cellNo, field_190000, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+14), (short)cellNo, field_210000, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+15), (short)cellNo, field_210400, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+16), (short)cellNo, field_210700, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+17), (short)cellNo, field_210000_other, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+18), (short)cellNo, field_220000, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+19), (short)cellNo, field_240000, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+20), (short)cellNo, field_250000, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+21), (short)cellNo, field_260000, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+22), (short)cellNo, field_400000, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+23), (short)cellNo, field_310000, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+24), (short)cellNo, "0", numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+25), (short)cellNo, field_320200, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+26), (short)cellNo, field_320300, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+27), (short)cellNo, field_300000, numberRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+28), (short)cellNo, field_600000, numberRightStyle);
                        cellNo++;
                        reportUtil.createCell(wb, sheet.getRow(rowNo+1), (short)cellNo, field_110000_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+2), (short)cellNo, field_110100_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+3), (short)cellNo, field_110300_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+4), (short)cellNo, field_110400_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+5), (short)cellNo, field_110600_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+6), (short)cellNo, field_111100_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+7), (short)cellNo, field_110000_other_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+8), (short)cellNo, field_120000_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+9), (short)cellNo, field_130000_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+10), (short)cellNo, field_140000_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+11), (short)cellNo, field_150000_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+12), (short)cellNo, field_160000_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+13), (short)cellNo, field_190000_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+14), (short)cellNo, field_210000_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+15), (short)cellNo, field_210400_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+16), (short)cellNo, field_210700_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+17), (short)cellNo, field_210000_other_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+18), (short)cellNo, field_220000_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+19), (short)cellNo, field_240000_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+20), (short)cellNo, field_250000_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+21), (short)cellNo, field_260000_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+22), (short)cellNo, field_400000_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+23), (short)cellNo, field_310000_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+24), (short)cellNo, "0.00", doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+25), (short)cellNo, field_320200_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+26), (short)cellNo, field_320300_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+27), (short)cellNo, field_300000_rate, doubleRightStyle);
                        reportUtil.createCell(wb, sheet.getRow(rowNo+28), (short)cellNo, field_600000_rate, doubleRightStyle);
                    }
                    //
                    if(i==dbData.size()-1 && Integer.parseInt(m_year)<Integer.parseInt(endYear)){
                        for(int y=Integer.parseInt(m_year)+1; y<=Integer.parseInt(endYear);y++){
                            cellNo++;
                            rowTmp=sheet.getRow(2);
                            reportUtil.createCell(wb, rowTmp, (short)cellNo, y+"年度 資料不完整", defaultStyle);//設年度
                            reportUtil.createCell(wb, rowTmp, (short)(cellNo+1), "", defaultStyle);
                            sheet.addMergedRegion(new Region((short)2, (short)cellNo, (short)2, (short)(cellNo+1)));
                            rowTmp=sheet.getRow(3);
                            sheet.setColumnWidth((short)cellNo,(short)4500);
                            reportUtil.createCell(wb, rowTmp, (short)cellNo, "金額", rightStyle);
                            reportUtil.createCell(wb, rowTmp, (short)(cellNo+1), "%", rightStyle);
                            for(int r=4;r<=31;r++){
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
                    sheet.addMergedRegion(new Region((short)2, (short)cellNo, (short)2, (short)(cellNo+1)));
                    rowTmp=sheet.getRow(3);
                    sheet.setColumnWidth((short)cellNo,(short)4500);
                    reportUtil.createCell(wb, rowTmp, (short)cellNo, "金額", rightStyle);
                    reportUtil.createCell(wb, rowTmp, (short)(cellNo+1), "%", rightStyle);
                    for(int r=4;r<=31;r++){
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
			System.out.println("//RptAN006 createRpt() Have Error.....");
			e.printStackTrace();
			System.out.println("//-------------------------------------");
		}
		System.out.println("RptAN006 createRpt() Debug End ...");
		return errMsg;
	}
	private static void setPreparedStatementParameter(PreparedStatement pst,List paramList) throws Exception{
		for(int i = 0 ;i< paramList.size() ;i++) {
			pst.setString(i+1,(String)paramList.get(i)) ;
		}
	}
	public static String MakesUpZero(String str) {
        if("0".equals(str)){
            str="0.00";
        }
        return str;
    }
}
