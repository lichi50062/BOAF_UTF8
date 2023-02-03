/*
	94.03.11 Fix 金額資料為零是不輸出0，改為輸出空白及欄位資料右靠處理
	94.08.09 fix當查詢條件>94年時,抓取=12月份當作94年底資料..以此類推 by 2295
	94.08.09 fix 月份統計僅顯示當年及上一年的全年月份資料 by 2295
	94.08.09 fix 本月與上月增減%-->若上月為"0"則為上一年的12月 by 2295
	94.08.10 fix 本月與上年同月比增減%=(本月-上年同月)/上年同月..若上年同月為0時,則/1 by 2295
			 fix 本月與上月比增減%=(本月-上月)/上月..若上月為0時,則/1 by 2295
	94.08.11 fix title設成標楷體 by 2295
	94.08.23 fix 更改公式 by 2295
	94.09.20 fix	調整SQL statement以解決存放款三欄位值未顯示。	jwang
   101.11.06 fix 調整SQL(跑太久) by 2295
   102.11.19 add 原QueryDB改套用QueryDB_SQLParam by 2295
*/
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.poifs.filesystem.*;
import org.apache.poi.hssf.record.ExtendedFormatRecord;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import java.io.*;
import java.text.*;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR002W {

	public static String createRpt(String s_year,String s_month,String unit,String datestate,String bank_type) {
		String errMsg="";
		String sqlCmd="";
		String unit_name="";
		int rowNum=0;
		int i=0;
		int j=0;
		int k=0;
		String hsien_id_sum[]={"","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p"};
		String month_name[]={"","一","二","三","四","五","六","七","八","九","十","十一","十二"};
		String filename=(bank_type.equals("6"))?"全體農會按縣市別經營指標變化表.xls":"全體漁會按縣市別經營指標變化表.xls";
		String bank_type_name=(bank_type.equals("6"))?"農會":"漁會";
		String s_year_last="";
		String s_month_last="";
		String div=(Integer.parseInt(s_year)==94 && Integer.parseInt(s_month)==6)?"1":"2";
		String wlx01_m_year = "";
        List paramList = new ArrayList(); 
        List roundrule_paramList = new ArrayList(); 
        List yearrule_paramList = new ArrayList(); 
        List monthrule1_paramList = new ArrayList(); 
        List monthrule_paramList = new ArrayList();
        wlx01_m_year = (Integer.parseInt(s_year) < 100)?"99":"100"; 
		String roundrule="round(sum(a01.fieldb)/?,0) as fieldb,"+
						 "round(sum(a01.fieldc)/?,0) as fieldc,"+
						 "decode(sum(a01.fieldb),0,0,round(sum(a01.fieldc)/sum(a01.fieldb),4)) as fieldd,"+
						 "round(sum(a01.fielde)/?,0) as fielde,"+
						 "decode(sum(a01.field_120000),0,0,round(sum(a01.fielde)/sum(a01.field_120000),4)) as fieldf,"+
						 "round(sum(a01.fieldg)/?,0) as fieldg,"+
						 "round(sum(a01.fieldh)/?,0) as fieldh,"+
						 "decode(sum(a01.fielde),0,0,round(sum(a01.fieldh)/sum(a01.fielde),4)) as fieldi,"+
						 "decode(sum(a01.field_120000),0,0,round(sum(a01.fieldh)/sum(a01.field_120000),4)) as fieldj,"+
						 "round(sum(a01.fieldk)/?,0) as fieldk,"+
						 "round(sum(a01.fieldl)/?,0) as fieldl,"+
						 "round(sum(a01.fieldm)/?,0) as fieldm,"+
						 "round(sum(a01.fieldn)/?,0) as fieldn,"+
						 "round(sum(a01.fieldo)/?,0) as fieldo,"+
						 "round(sum(a01.fieldp)/?,0) as fieldp";
		roundrule_paramList.add(unit);
		roundrule_paramList.add(unit);
		roundrule_paramList.add(unit);
		roundrule_paramList.add(unit);
		roundrule_paramList.add(unit);
		roundrule_paramList.add(unit);
		roundrule_paramList.add(unit);
		roundrule_paramList.add(unit);
		roundrule_paramList.add(unit);
		roundrule_paramList.add(unit);
		roundrule_paramList.add(unit);
		//針對存放款本月與上月平均（即年底及11月）
		String rule1="round(sum(decode(a01.acc_code,'220000',amt,0))-(round(sum(decode(a01.acc_code,'220900',amt,0))/2,0))/2,0) as fieldb,"+
					 "round((sum(decode(a01.acc_code,'310000',amt,0))+sum(decode(a01.acc_code,'320000',amt,0)))/2,0) as b_01,"+
					 "round(sum(decode(a01.acc_code,'140000',amt,0))/2,0) as c_01,"+
					 "round(sum(decode(a01.acc_code,'120000',amt,0))/2,0) as a_01,"+
					 "round((sum(decode(a01.acc_code,'120000',amt,0))-sum(decode(a01.acc_code,'310000',amt,0))+sum(decode(a01.acc_code,'320000',amt,0))+sum(decode(a01.acc_code,'140000',amt,0)))/2,0) as a_01_1";

		String yearrule="(select a01.m_year,bank_code,"+rule1+" from a01 left join (select * from ba01 where m_year = ?)ba01 on a01.bank_code=ba01.bank_no and ba01.bank_type=? where a01.m_year between 94 and ? and a01.m_month in (12,11)"+
					    " group by a01.m_year,a01.bank_code) a01";
		yearrule_paramList.add(wlx01_m_year);
		yearrule_paramList.add(bank_type);
		yearrule_paramList.add(String.valueOf(Integer.parseInt(s_year)-2));
	   //當年及上一年中當月及其上月的平均，但94年6月才開始申報，當月之分母為1
	   String monthrule1="round(sum(decode(a01.fieldb,0,0,a01.fieldb))/?,0) as fieldb,"+
						 "round(sum(decode(a01.b_01,0,0,a01.b_01))/?,0) as b_01,"+
						 "round(sum(decode(a01.c_01,0,0,a01.c_01))/?,0) as c_01,"+
						 "round(sum(decode(a01.a_01,0,0,a01.a_01))/?,0) as a_01,"+
						 "round(sum(decode(a01.a_01_1,0,0,a01.a_01_1))/?,0) as a_01_1";
	   monthrule1_paramList.add(div);
	   monthrule1_paramList.add(div);
	   monthrule1_paramList.add(div);
	   monthrule1_paramList.add(div);
	   monthrule1_paramList.add(div);
	   String monthrule2="sum(decode(a01.acc_code,'220000',amt,0))-(round(sum(decode(a01.acc_code,'220900',amt,0))/2,0)) as fieldb,"+
						 "sum(decode(a01.acc_code,'310000',amt,'320000',amt,0)) as b_01,"+
						 "sum(decode(a01.acc_code,'140000',amt,0)) as c_01,"+
						 "sum(decode(a01.acc_code,'120000',amt,0)) as a_01,"+
						 "sum(decode(a01.acc_code,'120000',amt,0))-sum(decode(a01.acc_code,'310000',amt,'320000',amt,0))+sum(decode(a01.acc_code,'140000',amt,0)) as a_01_1";

		String monthrule="((select ((to_char(a01.m_year+1911)||LPAD(to_char(a01.m_month),2,'0'))-191100 ) as yymm,bank_code,"+//當年全月先變為西元年再改為民國年月
						 monthrule2+
						 " from a01 left join (select * from ba01 where m_year=?)ba01 on a01.bank_code=ba01.bank_no and ba01.bank_type=? where a01.m_year=? group by a01.m_year,a01.m_month,a01.bank_code) union all "+
						 "(select decode(lpad(to_char(a01.m_month),2,'0'),'12',((to_char(a01.m_year+1912)||'01')-191100),((to_char(a01.m_year+1911)|| LPAD(to_char(a01.m_month),2,'0'))+1)-191100 ) as yymm,bank_code,"+//上一年全月往前加一個月再變成民國年月
						 monthrule2+
						 " from a01 left join (select * from ba01 where m_year=?)ba01 on a01.bank_code=ba01.bank_no and ba01.bank_type=? where a01.m_year=? group by a01.m_year,a01.m_month,a01.bank_code)) a01";
		monthrule_paramList.add(wlx01_m_year);
		monthrule_paramList.add(bank_type);
		monthrule_paramList.add(s_year);
		monthrule_paramList.add(wlx01_m_year);
		monthrule_paramList.add(bank_type);
		monthrule_paramList.add(String.valueOf(Integer.parseInt(s_year)-1));
		//針對逾期放款至淨值以一個月計（限為12月）
		String rule2="0 as fieldb,0 as fieldc,0 as fieldd,"+
					 "sum(decode(a01.acc_code,'990000',amt,0)) as fielde,"+
					 "round(decode(sum(decode(a01.acc_code,'120000',amt,0)),0,0,sum(decode(a01.acc_code,'990000',amt,0))/sum(decode(a01.acc_code,'120000',amt,0))),4) as fieldf,"+
					 "sum(decode(a01.acc_code,'150200',amt,0)) as fieldg,"+
					 "sum(decode(a01.acc_code,'120800',amt,'150300',amt,0)) as fieldh,"+
					 "round(decode(sum(decode(a01.acc_code,'990000',amt,0)),0,0,sum(decode(a01.acc_code,'120800',amt,'150300',amt,0))/sum(decode(a01.acc_code,'990000',amt,0))),4) as fieldi,"+
					 "round(decode(sum(decode(a01.acc_code,'120000',amt,0)),0,0,sum(decode(a01.acc_code,'120800',amt,'150300',amt,0))/sum(decode(a01.acc_code,'120000',amt,0))),4) as fieldj,"+
	  				 "sum(decode(a01.acc_code,'140000',amt,0)) as fieldl,"+
					 "sum(decode(a01.acc_code,'320300',amt,0)) as fieldm,"+
					 "sum(decode(a01.acc_code,'310000',amt,0)) as fieldn,"+
					 "sum(decode(a01.acc_code,'320000',amt,0)) as fieldo,"+
					 "sum(decode(a01.acc_code,'120000',amt,0)) as field_120000,";
		if (bank_type.equals("6")) {
			  rule2+="sum(decode(a01.acc_code,'120700',amt,0)) as fieldk,sum(decode(a01.acc_code,'310000',amt,'320000',amt,0)) as fieldp";
		}else if (bank_type.equals("7")) {
			  rule2+="sum(decode(a01.acc_code,'120700',amt,'120900',amt,0)) as fieldk,sum(decode(a01.acc_code,'300000',amt,0)) as fieldp";
		}

        reportUtil reportUtil=new reportUtil();
		try {
			File xlsDir=new File(Utility.getProperties("xlsDir"));
        	File reportDir=new File(Utility.getProperties("reportDir"));

    		if (!xlsDir.exists()) {
     			if (!Utility.mkdirs(Utility.getProperties("xlsDir"))) {
     		   		errMsg +=Utility.getProperties("xlsDir")+"目錄新增失敗";
     			}
    		}
    		if (!reportDir.exists()) {
     			if (!Utility.mkdirs(Utility.getProperties("reportDir"))) {
     		   		errMsg +=Utility.getProperties("reportDir")+"目錄新增失敗";
     			}
    		}

			FileInputStream finput=new FileInputStream(xlsDir +System.getProperty("file.separator")+filename );

	  	    //設定FileINputStream讀取Excel檔
	  		POIFSFileSystem fs=new POIFSFileSystem( finput );
	  		if(fs==null){System.out.println("開啟範本檔失敗");} else System.out.println("開啟範本檔成功="+xlsDir +System.getProperty("file.separator")+filename);
	  		HSSFWorkbook wb=new HSSFWorkbook(fs);
	  		if(wb==null){System.out.println("開啟工作表失敗");}else System.out.println("開啟工作表成功");
	  		HSSFSheet sheet=wb.getSheetAt(0);//讀取第一個工作表，宣告其為sheet
	  		if(sheet==null){System.out.println("開啟工作表失敗");}else System.out.println("開啟工作表成功");
	  		HSSFPrintSetup ps=sheet.getPrintSetup(); //取得設定
	        //sheet.setZoom(80,100); // 螢幕上看到的縮放大小
	        //sheet.setAutobreaks(true); //自動分頁

	        //設定頁面符合列印大小
	        sheet.setAutobreaks( false );
	        ps.setScale( ( short )72 ); //列印縮放百分比

	        ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
	  		//wb.setSheetName(0,"test");
	  		finput.close();

	  		HSSFRow row=null;//宣告一列
	  		HSSFCell cell=null;//宣告一個儲存格

			//年度統計
			String sqlCmd_yearsum="select a01.m_year,"+roundrule+" from (select a01.m_year,bank_code,sum(fieldb) as fieldb,sum(fieldc) as fieldc,sum(fieldd) as fieldd,"+
								  "sum(fielde) as fielde,sum(fieldf) as fieldf,sum(fieldg) as fieldg,sum(fieldh) as fieldh,sum(fieldi) as fieldi,sum(fieldj) as fieldj,"+
								  "sum(fieldk) as fieldk,sum(fieldl) as fieldl,sum(fieldm) as fieldm,sum(fieldn) as fieldn,sum(fieldo) as fieldo,sum(fieldp) as fieldp,"+
								  "sum(field_120000) as field_120000 from "+
								  "((select m_year,bank_code,fieldb,a_01 as fieldc,decode(fieldb,0,0,round(a_01/fieldb,4)) as fieldd,0 as fielde,0 as fieldf,0 as fieldg,0 as fieldh,"+
								  "0 as fieldi,0 as fieldj,0 as fieldk,0 as fieldl,0 as fieldm,0 as fieldn,0 as fieldo,0 as fieldp,0 as field_120000 from "+yearrule+" where b_01>c_01)"+
								  " union all "+
								  "(select a01.m_year,bank_code,fieldb,a_01 as fieldc,decode(fieldb,0,0,round(a_01_1/fieldb,4)) as fieldd,0 as fielde,0 as fieldf,0 as fieldg,0 as fieldh,"+
								  "0 as fieldi,0 as fieldj,0 as fieldk,0 as fieldl,0 as fieldm,0 as fieldn,0 as fieldo,0 as fieldp,0 as field_120000 from "+yearrule+" where b_01<=c_01)"+
								  " union all "+
								  "(select a01.m_year,bank_code,"+rule2+" from a01 left join (select * from ba01 where m_year=?)ba01 on a01.bank_code=ba01.bank_no and ba01.bank_type=? where a01.m_year between 94 and ? and a01.m_month=12"+
								  " group by a01.m_year,a01.bank_code)) a01 group by a01.m_year,a01.bank_code) a01 group by a01.m_year order by a01.m_year";
			List sqlCmd_yearsum_paramList = new ArrayList();
			//roundrule參數
			for(int roundrulei=0;roundrulei<roundrule_paramList.size();roundrulei++){
			    sqlCmd_yearsum_paramList.add(roundrule_paramList.get(roundrulei));
	        }
			//yearrule參數
            for(int yearrulei=0;yearrulei<yearrule_paramList.size();yearrulei++){
                sqlCmd_yearsum_paramList.add(yearrule_paramList.get(yearrulei));
            }
            //yearrule參數
            for(int yearrulei=0;yearrulei<yearrule_paramList.size();yearrulei++){
                sqlCmd_yearsum_paramList.add(yearrule_paramList.get(yearrulei));
            }
            sqlCmd_yearsum_paramList.add(wlx01_m_year);
            sqlCmd_yearsum_paramList.add(bank_type);
            sqlCmd_yearsum_paramList.add(String.valueOf(Integer.parseInt(s_year)-2));
			List dbData_yearsum=DBManager.QueryDB_SQLParam(sqlCmd_yearsum,sqlCmd_yearsum_paramList,"fieldB,fieldC,fieldD,fieldE,fieldF,fieldG,fieldH,fieldI,fieldJ,fieldK,fieldL,fieldM,fieldN,fieldO,fieldP");
			System.out.println("年度統計的dbData_yearsum.size()="+dbData_yearsum.size());

			//月份統計
			String sqlCmd_monthsum="select a01.yymm,"+roundrule+" from ((select yymm,sum(a01.fieldb) as fieldb,sum(a01.a_01) as fieldc,"+
								   "decode(sum(decode(a01.fieldb,0,0,a01.fieldb)),0,0,round(sum(decode(a01.a_01,0,0,a01.a_01))/sum(decode(a01.fieldb,0,0,a01.fieldb)),4)) as fieldd,"+
								   "0 as fielde,0 as fieldf,0 as fieldg,0 as fieldh,0 as fieldi,0 as fieldj,0 as fieldk,0 as fieldl,0 as fieldm,0 as fieldn,0 as fieldo,0 as fieldp,"+
								   "0 as field_120000 from (select a01.yymm,a01.bank_code,"+monthrule1+" from "+monthrule+" where b_01>c_01 group by a01.yymm,a01.bank_code) a01 group by a01.yymm)"+
								   " union all "+
								   "(select yymm,sum(a01.fieldb) as fieldb,sum(a01.a_01) as fieldc,"+
								   "decode(sum(decode(a01.fieldb,0,0,a01.fieldb)),0,0,round(sum(decode(a01.a_01_1,0,0,a01.a_01_1))/sum(decode(a01.fieldb,0,0,a01.fieldb)),4)) as fieldd,"+
								   "0 as fielde,0 as fieldf,0 as fieldg,0 as fieldh,0 as fieldi,0 as fieldj,0 as fieldk,0 as fieldl,0 as fieldm,0 as fieldn,0 as fieldo,0 as fieldp,"+
								   "0 as field_120000 from (select a01.yymm,a01.bank_code,"+monthrule1+" from "+monthrule+" where b_01<=c_01 group by a01.yymm,a01.bank_code) a01 group by a01.yymm)"+
								   " union all "+
								   "(select ((to_char(a01.m_year+1911)||lpad(to_char(a01.m_month),2,'0'))-191100 ) as yymm,"+rule2+" from a01,ba01 where a01.m_year between ? and ? "
								   +" and  a01.bank_code=ba01.bank_no and ba01.bank_type=? group by a01.m_year,a01.m_month)) a01"+
								   " group by a01.yymm order by a01.yymm";
			List sqlCmd_monthsum_paramList = new ArrayList();
			//roundrule參數
            for(int roundrulei=0;roundrulei<roundrule_paramList.size();roundrulei++){
                sqlCmd_monthsum_paramList.add(roundrule_paramList.get(roundrulei));
            }
            //monthrule1參數
            for(int monthrule1i=0;monthrule1i<monthrule1_paramList.size();monthrule1i++){
                sqlCmd_monthsum_paramList.add(monthrule1_paramList.get(monthrule1i));
            }
            //monthrule參數
            for(int monthrulei=0;monthrulei<monthrule_paramList.size();monthrulei++){
                sqlCmd_monthsum_paramList.add(monthrule_paramList.get(monthrulei));
            }
            //monthrule1參數
            for(int monthrule1i=0;monthrule1i<monthrule1_paramList.size();monthrule1i++){
                sqlCmd_monthsum_paramList.add(monthrule1_paramList.get(monthrule1i));
            }
            //monthrule參數
            for(int monthrulei=0;monthrulei<monthrule_paramList.size();monthrulei++){
                sqlCmd_monthsum_paramList.add(monthrule_paramList.get(monthrulei));
            }
            sqlCmd_monthsum_paramList.add(String.valueOf(Integer.parseInt(s_year)-1));
            sqlCmd_monthsum_paramList.add(s_year);
            sqlCmd_monthsum_paramList.add(bank_type);
			List dbData_monthsum=DBManager.QueryDB_SQLParam(sqlCmd_monthsum,sqlCmd_monthsum_paramList,"yymm,fieldB,fieldC,fieldD,fieldE,fieldF,fieldG,fieldH,fieldI,fieldJ,fieldK,fieldL,fieldM,fieldN,fieldO,fieldP");
			System.out.println("月份統計的dbData_monthsum.size()="+dbData_monthsum.size());

			//列印年月
			row=(sheet.getRow(1)==null)? sheet.createRow(1) : sheet.getRow(1);
			cell=row.getCell((short)1);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("中華民國"+s_year+"年"+s_month+"月");

			//列印日期
			if (datestate.equals("1")) {
				Calendar rightNow=Calendar.getInstance();
				String year=String.valueOf(rightNow.get(Calendar.YEAR)-1911);
				String month=String.valueOf(rightNow.get(Calendar.MONTH)+1);
				String day=String.valueOf(rightNow.get(Calendar.DAY_OF_MONTH));
				cell=row.getCell((short)0);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellValue("列印日期："+year+"年"+month+"月"+day+"日");
            }

			//列印金額單位
			cell=row.getCell((short)13);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			if (unit.equals("1")) {
				unit_name="元";
		   	}
		   		else if (unit.equals("1000")) {
	   	    		unit_name="千元";
   	        	}
   	        	else if (unit.equals("10000")) {
   	        		unit_name="萬元";
	   	    	}
	   	    	else if (unit.equals("1000000")) {
	   	    		unit_name="百萬元";
	   	    	}
	   	    	else if (unit.equals("10000000")) {
	   	    		unit_name="千萬元";
   	        	}
   	        	else if (unit.equals("100000000")) {
   	        		unit_name="億元";
	   	    	}
			cell.setCellValue("金額單位：新台幣"+unit_name+",％");

			short top=0,down=0;

			//年度統計
			rowNum=4;
			System.out.println("rowNum="+rowNum);
			//查詢民國96年（含）後資料才印前年度統計
			if (Integer.parseInt(s_year)>=96) {
				for (j=0;j<dbData_yearsum.size();j++) {
			        if (rowNum>sheet.getLastRowNum()) {
						row=sheet.createRow((short)rowNum );
			        }else {
						row=sheet.getRow((short)rowNum);
			        }
			        System.out.println("yymm="+(((DataObject)dbData_monthsum.get(j)).getValue("yymm")).toString());
			        insertCell(null,false,0,(((DataObject)dbData_monthsum.get(j)).getValue("yymm")).toString().trim()+"年度",wb,row,(short)0,(short)26,(short)1,top,down,(short)0,HSSFCellStyle.BORDER_THIN);
			        for (k=1;k<=15;k++) {
			    		//System.out.println("k="+k);
			    		if (k==15) {
							insertCell(dbData_yearsum,true,j,"field"+hsien_id_sum[k],wb,row,(short)k,(short)64,(short)0,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);
			    		}
			    			else {
								insertCell(dbData_yearsum,true,j,"field"+hsien_id_sum[k],wb,row,(short)k,(short)64,(short)0,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);
			    			}
		            }
			        rowNum++;
				}
			}

			//月份統計
			String m_year="";
			String m_month="";
			String datarange="";
			DataObject lastYearMonth=null;//上年同月資料
			DataObject nowMonth=null;//本月資料
			DataObject lastMonth=null;//上月資料
			String tmp_datarange="";
			int begin=0;
			int end=0;
			monthLoop:
			for (j=0;j<dbData_monthsum.size();j++) {
				//datarange=(String)((DataObject)dbData_monthsum.get(j)).getValue("yymm");
				datarange=(((DataObject)dbData_monthsum.get(j)).getValue("yymm")).toString();
				if(wlx01_m_year.equals("100")){
				    tmp_datarange=datarange.substring(0,3).trim();
				}else{
				    tmp_datarange=datarange.substring(0,2).trim();
				}
				System.out.println("datarange="+datarange);
				System.out.println("m_year="+m_year);
				//if(datarange.indexOf("年") !=-1){
				   //年份不同時
				   if(!m_year.equals(tmp_datarange)){
				   	   System.out.println("m_year="+m_year);
				   	   System.out.println("年份不同");
				   	   if(wlx01_m_year.equals("100")){
				   	       m_year=datarange.substring(0,3).trim();
				   	       System.out.println("nowyear="+m_year);
				   	   }else{
				   	       m_year=datarange.substring(0,2).trim();
	                       System.out.println("nowyear="+m_year);
				   	   }
				       //if(Integer.parseInt(m_year) > 93) break monthLoop;

				       if(rowNum > sheet.getLastRowNum()){
				       	  row=sheet.createRow((short)rowNum );
				       }else{
				          row=sheet.getRow((short)rowNum);
				       }
				       insertCell(dbData_yearsum,false,0,m_year+"年",wb,row,(short)0,(short)26,(short)1,top,down,(short)0,HSSFCellStyle.BORDER_THIN);
				       insertBlank(15,wb,(short)0,(short)0,row);//印出年度的空值表格
				       rowNum++;
				       begin=rowNum;//此年度月份的起始
				       for(i=1;i<=12;i++){
				       	   if(rowNum > sheet.getLastRowNum()){
					       	  row=sheet.createRow((short)rowNum );
					       }else{
					          row=sheet.getRow((short)rowNum);
					       }
				       	   insertCell(dbData_yearsum,false,0,"  "+month_name[i]+"月",wb,row,(short)0,(short)26,(short)1,top,down,(short)0,HSSFCellStyle.BORDER_THIN);
				       	   insertBlank(15,wb,(short)0,(short)0,row);//印出每月的空值表格
				       	   end=rowNum;//此年度月份的結束
				       	   rowNum++;
				       }
				       System.out.println("begin="+begin);
				       System.out.println("end="+end);
				       if(wlx01_m_year.equals("100")){
				           m_month=datarange.substring(3,datarange.length()).trim();
	                       System.out.println("年份不同.month="+m_month);
				       }else{
				           m_month=datarange.substring(2,datarange.length()).trim();
	                       System.out.println("年份不同.month="+m_month);
				       }
				      

				       for(int monthidx=begin;monthidx<=end;monthidx++){
				       	   row=sheet.getRow((short)monthidx);
				       	   cell=row.getCell((short)0);
				       	   System.out.println("eachmonth="+cell.getStringCellValue().trim());
				       	   System.out.println("nowmonth="+month_name[Integer.parseInt(m_month)]+"月");
				       	   //取得上年同月資料
					   	   if((Integer.parseInt(m_year)==(Integer.parseInt(s_year)-1))//上年
					   	   && (Integer.parseInt(m_month)==Integer.parseInt(s_month))){//同月
					   	      lastYearMonth=(DataObject)dbData_monthsum.get(j);
					   	      System.out.println("取得上年同月");
					   	   }
				       	   if((cell.getStringCellValue().trim()).equals(month_name[Integer.parseInt(m_month)]+"月" )){
				       	   	   //System.out.println("insert month data");
				       	       for(k=1;k<=15;k++){
					       	       //row=sheet.getRow((short)rowNum);
				    		       if(k==15){
				    		       	  //System.out.println("field"+hsien_id_sum[k]);
					       	          insertCell(dbData_monthsum,true,j,"field"+hsien_id_sum[k],wb,row,(short)k,(short)64,(short)0,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);
				    		       }else{
				    		          insertCell(dbData_monthsum,true,j,"field"+hsien_id_sum[k],wb,row,(short)k,(short)64,(short)0,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);
				    		       }
			                   }//end of 印出每筆名細
				       	       continue monthLoop;
				       	    }
				       }
				   }else{
				   	   System.out.println("年份相同");
				   	   //System.out.println("nowyear="+datarange.substring(0,datarange.indexOf("年")).trim());
				   	   System.out.println("datarange="+datarange);
				   	   if(wlx01_m_year.equals("100")){
				   	       System.out.println("nowyear="+datarange.substring(0,3).trim());
				   	       System.out.println("m_year="+m_year);	   
				   	       m_month=datarange.substring(3,datarange.length()).trim();
	                       System.out.println("年份相同.m_month="+m_month);
				   	   }else{
				   	       System.out.println("nowyear="+datarange.substring(0,2).trim());
				   	       System.out.println("m_year="+m_year);
				   	       m_month=datarange.substring(2,datarange.length()).trim();
				   	       System.out.println("年份相同.m_month="+m_month);
				   	   }
				   	   //m_month=datarange.substring(datarange.indexOf("年")+1,datarange.indexOf("月")).trim();
				   	  
				   	   //取得本月資料
				   	   if( Integer.parseInt(m_year)==Integer.parseInt(s_year)
				   	   	&& Integer.parseInt(m_month)==Integer.parseInt(s_month)){//本年本月
				   	   	  System.out.println("取得本月資料"+s_year+m_month);
				   	      nowMonth=(DataObject)dbData_monthsum.get(j);
				   	   }
				   	   //取得上月資料,若本月為1月,則取得上年12月資料
				   	   if(Integer.parseInt(s_month)==1){
 				   	      //取得上年12月資料
					   	   if((Integer.parseInt(m_year)==(Integer.parseInt(s_year)-1))//上年
					   	   && Integer.parseInt(m_month)==12){//12月
					   	      lastMonth=(DataObject)dbData_monthsum.get(j);
					   	      System.out.println("取得上年12月"+m_year+m_month);
					   	   }
				   	   }else if((Integer.parseInt(m_month)==(Integer.parseInt(s_month)-1))
				   	         &&(Integer.parseInt(m_year)==Integer.parseInt(s_year))){//本年上月
				   	         System.out.println("取得上月資料"+m_month);
				   	         lastMonth=(DataObject)dbData_monthsum.get(j);
				   	   }
				   	   //取得上年同月資料
				   	   if((Integer.parseInt(m_year)==(Integer.parseInt(s_year)-1))//上年
				   	   && (Integer.parseInt(m_month)==Integer.parseInt(s_month))){//同月
				   	      lastYearMonth=(DataObject)dbData_monthsum.get(j);
				   	      System.out.println("取得上年同月"+m_year+m_month);
				   	   }
				   	   for(int monthidx=begin;monthidx<=end;monthidx++){
				       	   row=sheet.getRow((short)monthidx);
				       	   cell=row.getCell((short)0);
				       	   System.out.println("eachmonth="+cell.getStringCellValue().trim());
				       	   System.out.println("nowmonth="+month_name[Integer.parseInt(m_month)]+"月");

				       	   for(int idx=1;idx<month_name.length;idx++){
				     	       if((cell.getStringCellValue().trim()).equals(month_name[Integer.parseInt(m_month)]+"月" )){
				     	       	   //System.out.println("insert month data");
				       	           for(k=1;k<=15;k++){
					       	           //row=sheet.getRow((short)rowNum);
				    		           if(k==15){
				    		           	  //System.out.println("field"+hsien_id_sum[k]);
					       	              insertCell(dbData_monthsum,true,j,"field"+hsien_id_sum[k],wb,row,(short)k,(short)64,(short)0,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);
				    		           }else{
				    		              insertCell(dbData_monthsum,true,j,"field"+hsien_id_sum[k],wb,row,(short)k,(short)64,(short)0,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);
				    		           }
			                       }//end of 印出每筆名細
				       	           continue monthLoop;
				       	       }
				       	   }
				       }
				   }
				//}
			}//end of 屬於該年度月份的統計資料
			System.out.println("本月="+nowMonth);
			if(nowMonth !=null){
			   for(k=1;k<=15;k++){
		           System.out.print((nowMonth.getValue("field"+hsien_id_sum[k])).toString()+":");
               }//end of 印出每筆名細
			   System.out.println("");
			}
			System.out.println("上月="+lastMonth);
			if(lastMonth !=null){
			   for(k=1;k<=15;k++){
		           System.out.print((lastMonth.getValue("field"+hsien_id_sum[k])).toString()+":");
               }//end of 印出每筆名細
			}
			System.out.println("上年同月="+lastYearMonth);
			if(lastYearMonth !=null){
			   for(k=1;k<=15;k++){
		           System.out.print((lastYearMonth.getValue("field"+hsien_id_sum[k])).toString()+":");
               }//end of 印出每筆名細
			}
			//本月與上月比增減
			if (rowNum>sheet.getLastRowNum()) {
				row=sheet.createRow((short)rowNum);
		    }
		    	else {
					row=sheet.getRow((short)rowNum);
		    	}
			insertCell(null,false,0," ",wb,row,(short)0,(short)26,(short)1,top,down,(short)0,HSSFCellStyle.BORDER_THIN);
		    insertBlank(15,wb,(short)0,(short)0,row);//先空一行
		    rowNum++;
		    if (rowNum>sheet.getLastRowNum()) {
				row=sheet.createRow((short)rowNum);
		    }
		    	else {
					row=sheet.getRow((short)rowNum);
		    	}
		    insertCell(null,false,0,"本月與上月比增減%",wb,row,(short)0,(short)26,(short)1,top,down,(short)0,HSSFCellStyle.BORDER_THIN);
		    //(本月-上月)/上月
		    String lastMonthdata="";
		    double data1=0.0;
		    double data2=0.0;
		    double data3=0.0;
		    if(lastMonth !=null){
		    	System.out.println("上月 !=null");
		    	for(k=1;k<=15;k++){
		    		lastMonthdata=lastMonth.getValue("field"+hsien_id_sum[k]).toString();
		    		if(lastMonthdata.equals("0")){//上月為"0"時,顯示"--"
		    		   if(k==15){
		    		   	  insertCell(null,false,0,"--",wb,row,(short)k,(short)64,(short)0,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);
		    		   }else{
		    		   	  insertCell(null,false,0,"--",wb,row,(short)k,(short)64,(short)0,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);
		    		   }
		    		}else{
		    		   /*本月*/data1=(nowMonth==null)?Double.parseDouble("0.0"):Double.parseDouble(nowMonth.getValue("field"+hsien_id_sum[k]).toString());
		    		   /*上月*/data2=Double.parseDouble(lastMonth.getValue("field"+hsien_id_sum[k]).toString());
		    		   /*(本月-上月)/上月*/data3=(data1 - data2)/data2;
		    		   if(k==15){
		    		   	  insertCell(null,false,0,String.valueOf(data3),wb,row,(short)k,(short)64,(short)0,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);
		    		   }else{
		    		   	  insertCell(null,false,0,String.valueOf(data3),wb,row,(short)k,(short)64,(short)0,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);
		    		   }
		    		}
		    		lastMonthdata="";
				    data1=0.0;
				    data2=0.0;
				    data3=0.0;
			    }//end of 印出每筆名細
		    }else{
		    	System.out.println("上月==null");
			    insertBlank(15,wb,(short)0,(short)0,row);
			}//end of 本月與上月比增減

			rowNum++;

			//本月與上年同月比增減%
		    if(rowNum > sheet.getLastRowNum()){
		       row=sheet.createRow((short)rowNum );
		    }else{
		       row=sheet.getRow((short)rowNum);
		    }
		    insertCell(null,false,0,"本月與上年同月比增減%",wb,row,(short)0,(short)26,(short)1,top,down,(short)0,HSSFCellStyle.BORDER_THIN);
		    //本月-上年同月/上年同月
		    String lastYearMonthdata="";
		    if(lastYearMonth !=null){
		    	for(k=1;k<=15;k++){
		    		lastYearMonthdata=lastYearMonth.getValue("field"+hsien_id_sum[k]).toString();
		    		if(lastYearMonthdata.equals("0")){//上年同月為"0"時,顯示"--"
		    		   if(k==15){
		    		   	  insertCell(null,false,0,"--",wb,row,(short)k,(short)64,(short)0,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);
		    		   }else{
		    		   	  insertCell(null,false,0,"--",wb,row,(short)k,(short)64,(short)0,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);
		    		   }
		    		}else{
		    		   ///*本月*/data1=Double.parseDouble(nowMonth.getValue("field"+hsien_id_sum[k]).toString());
		    		   /*本月*/data1=(nowMonth==null)?Double.parseDouble("0.0"):Double.parseDouble(nowMonth.getValue("field"+hsien_id_sum[k]).toString());
		    		   /*上年同月*/data2=Double.parseDouble(lastYearMonth.getValue("field"+hsien_id_sum[k]).toString());
		    		   /*(本月-上年同月)/上年同月*/data3=(data1 - data2)/data2;
		    		   if(k==15){
		    		   	  insertCell(null,false,0,String.valueOf(data3),wb,row,(short)k,(short)64,(short)0,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);
		    		   }else{
		    		   	  insertCell(null,false,0,String.valueOf(data3),wb,row,(short)k,(short)64,(short)0,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);
		    		   }
		    		}
		    		lastYearMonthdata="";
				    data1=0.0;
				    data2=0.0;
				    data3=0.0;
			    }//end of 印出每筆名細
		    }else{
			       insertBlank(15,wb,(short)0,(short)0,row);
			}//end of 本月與上年同月比增減

			rowNum++;

			if(rowNum > sheet.getLastRowNum()){
			   row=sheet.createRow((short)rowNum );
			}else{
			   row=sheet.getRow((short)rowNum);
			}
			insertCell(null,false,0," ",wb,row,(short)0,(short)26,(short)1,top,HSSFCellStyle.BORDER_MEDIUM,(short)0,HSSFCellStyle.BORDER_THIN);
			insertBlank(15,wb,(short)0,HSSFCellStyle.BORDER_MEDIUM,row);//先空一行
			rowNum++;
			if(rowNum > sheet.getLastRowNum()){
			   row=sheet.createRow((short)rowNum );
			}else{
			   row=sheet.getRow((short)rowNum);
			}

			insertCell(null,false,0,"資料來源 :根據各"+bank_type_name+"信用部資料彙編。",wb,row,(short)0,(short)64,(short)0,top,(short)0,(short)0,(short)0);

			HSSFFooter footer=sheet.getFooter();
	        footer.setCenter( "Page:" +HSSFFooter.page() +" of " +HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));


	       	FileOutputStream fout=new FileOutputStream(reportDir +System.getProperty("file.separator")+filename);
	        wb.write(fout);
	        //儲存
	        fout.close();
	        System.out.println("儲存完成");
		}catch(Exception e){
			System.out.println("createRpt Error:"+e+e.getMessage());
		}
		return errMsg;
	}

	public static void insertCell(List dbData,boolean getstate,int index,String Item,HSSFWorkbook wb,HSSFRow row,short j,
	                              short bg,short fp,
							      short bordertop,short borderbottom,short borderleft,short borderright)
	{
		try{
		String insertValue="";
  		if(getstate) insertValue=(((DataObject)dbData.get(index)).getValue(Item)).toString();
  		else         insertValue=Item;
		//System.out.println("insertValue="+insertValue);
	    HSSFCell cell=(row.getCell(j)==null)? row.createCell(j) : row.getCell(j);
	    HSSFCellStyle cs1=wb.createCellStyle();
	    HSSFCellStyle cs2=cell.getCellStyle();
	    //System.out.println("getFillPattern="+cs2.getFillPattern());
	    //System.out.println("getFillForegroundColor="+cs2.getFillForegroundColor());
	    //System.out.println("getFillBackgroundColor="+cs2.getFillBackgroundColor());
	    /*System.out.println("setBorderTop="+cs2.getBorderTop());
	    System.out.println("setBorderBottom="+cs2.getBorderBottom());
	    System.out.println("setBorderLeft="+cs2.getBorderLeft());
	    System.out.println("setBorderRight="+cs2.getBorderRight());
	    */
	    cs1.setBorderTop(bordertop);
	    cs1.setBorderBottom(borderbottom);
	    cs1.setBorderLeft(borderleft);
	    cs1.setBorderRight(borderright);
	    cs1.setFillPattern(fp);
	    cs1.setFillForegroundColor(bg);

	    if(insertValue.indexOf("資料來源") !=-1){
	        cs1.setAlignment(HSSFCellStyle.ALIGN_GENERAL);	//94.03.11 add by egg
	    }else{
	    	cs1.setAlignment((short) 3);	//94.03.11 add by egg
	    }
	    //94.8.11 fix title設成標楷體
	    if(j==0){
	       HSSFFont f=wb.createFont();
	       f.setFontHeightInPoints((short) 12);
	       f.setFontName("標楷體");
	       cs1.setFont(f);
	    }
	    cell.setCellStyle(cs1);
	    cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
	    //double value=0;

	    if(getstate){
	       if(!insertValue.equals("0"))cell.setCellValue(Utility.setCommaFormat(insertValue));
	    }else{
	       cell.setCellValue(insertValue);
	    }

	    /*
	    try{
	    	cs1.setDataFormat((short)3);	// "#,##0"
	    	cell.setCellValue(Double.parseDouble(insertValue));
	    }catch(NumberFormatException e){
	    	cell.setCellValue(insertValue);
	    }*/
		}catch(Exception e){
			System.out.println("insertCell Error:"+e+e.getMessage());
		}
	}

	public static void insertBlank(int maxlength,HSSFWorkbook wb,short top,short down,HSSFRow row){
	    for(int k=1;k<=maxlength;k++){
	        //System.out.println("k="+k);
	        if(k==15){
   	           insertCell(null,false,0," ",wb,row,(short)k,(short)64,(short)0,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);
	        }else{
	           insertCell(null,false,0," ",wb,row,(short)k,(short)64,(short)0,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);
	        }
       }//end of insert 空值表格
	}
}