/*
 * Created on 2006/11/02 by ABYSS Allen
 * RptCG001WA 稽核記錄統計總表
 */

package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.Region;
import java.io.*;
import java.sql.*;
import java.util.*;
import com.tradevan.util.Utility;
import com.tradevan.util.dao.RdbCommonDao;
import com.tradevan.util.report.reportUtil;

public class RptCG001WA {

	public static String createRpt(String startYear,String startMonth, String endYear, String endMonth, String reportType
			, String bankType, String[] tableName, HashMap staticBankNameMap) {
		boolean debug = false;
		startYear = Integer.parseInt(startYear)+1911+"";
		endYear = Integer.parseInt(endYear)+1911+"";
		String[] endDayStr = new String[]{"31","28","31","30","31","30","31","31","30","31","30","31"};
		if(debug) System.out.println("RptCG001WA createRpt() Debug Start ...");
		String errMsg = "";
		String sqlCmd = "";
		Connection conn = null;
		PreparedStatement pst = null;
		Statement st = null;
		ResultSet rs = null;
		HSSFCellStyle defaultStyle;
	    HSSFCellStyle rightStyle;
	    HSSFCellStyle leftStyle;
	    HSSFCellStyle noBorderDefaultStyle;
	    HSSFCellStyle noBorderLeftStyle;
	    HSSFCellStyle titleStyle;
		Calendar calendar = Calendar.getInstance();
		calendar.add(2,-2);
		calendar.set(calendar.get(1),calendar.get(2),1);
		if(debug) System.out.println("calendar date="+calendar.get(1)+"/"+(calendar.get(2)+1)+"/"+calendar.get(5));
		try {
			conn =(new RdbCommonDao("")).newConnection();
			int monthIndex = (Integer.parseInt(endYear)-Integer.parseInt(startYear))*12+(Integer.parseInt(endMonth)-Integer.parseInt(startMonth))+1;//一共查詢幾個月份
			int xArrayLength = monthIndex*3;//陣列長度
			short totalColumnNums = (short)xArrayLength;
			if(debug) System.out.println("monthIndex="+monthIndex+";xArrayLength="+xArrayLength+";totalColumnNums="+totalColumnNums);
			String queryYear="",queryMonth="";
			List bankList = new ArrayList();
//==== 先查詢出所有機關名稱
			sqlCmd="SELECT BANK_NO, BANK_NAME FROM BN01 WHERE 1=1";
			if(!bankType.equals("ALL")){
				sqlCmd+=" AND BANK_TYPE=?";
			}
			sqlCmd+=" ORDER BY BANK_NAME";
			if(debug) System.out.println("1. sqlCmd="+sqlCmd);
			pst = conn.prepareStatement(sqlCmd);
			if(!bankType.equals("ALL")){
				pst.setString(1, bankType);
			}
			rs=pst.executeQuery();
			while(rs.next()){
				String[] bankTmp = new String[]{rs.getString(1),rs.getString(2)};
				bankList.add(bankTmp);
			}
			rs.close();
			pst.close();
			if(bankType.equals("2")||bankType.equals("ALL")){
				String[] bankTmp = new String[]{"empty","農業金融局"};
				bankList.add(bankTmp);
			}
			st = conn.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			HashMap resultMap = new HashMap();//計劃裡面以bank_no+查詢年月當key,裡面放long[3]
//=== 依年月查詢
			for(int mi=0;mi<monthIndex;mi++){
				if((Integer.parseInt(startMonth)+mi)>12){
					queryYear = String.valueOf(Integer.parseInt(startYear)+1);
					queryMonth=(Integer.parseInt(startMonth)+mi-12)+"";
				}else{
					queryYear = startYear;
					queryMonth = String.valueOf(Integer.parseInt(startMonth)+mi);
				}
				Calendar queryDate = Calendar.getInstance();
				queryDate.set(Integer.parseInt(queryYear),Integer.parseInt(queryMonth)-1,5);
				if(debug) System.out.println("queryYear="+queryYear+";queryMonth="+queryMonth+"; queryDate.after(calendar)="+queryDate.after(calendar));
				//=== 若為三個月內的查詢
				if(queryDate.after(calendar)){
					String startQueryDate = Integer.parseInt(queryYear)+"/"+queryMonth+"/1";
					String endQueryDate = Integer.parseInt(queryYear)+"/"+queryMonth+"/"+endDayStr[(Integer.parseInt(queryMonth))-1];
					if(debug) System.out.println("startQueryDate="+startQueryDate+";endQueryDate="+endQueryDate);
					for(int ai=0;ai<tableName.length;ai++){
						if(tableName[ai].equals("WTT07")){
							sqlCmd="SELECT w.TBANK_NO, COUNT(a.RESULT_P) NUMS FROM WTT07 a, WTT01 w"
								+" WHERE w.MUSER_ID=a.MUSER_ID AND RESULT_P='XOO'"
								+" AND a.INPUT_DATE BETWEEN to_date('"+startQueryDate+"', 'yyyy/mm/dd') AND to_date('"
								+endQueryDate+"', 'yyyy/mm/dd')";
							if(!bankType.equals("ALL")){
								sqlCmd+=" AND w.BANK_TYPE='"+bankType+"'";
							}
							sqlCmd+=" GROUP BY w.TBANK_NO ORDER BY w.TBANK_NO";
						}else if(tableName[ai].equals("MUSER_DATA_log")){
							sqlCmd="SELECT w.TBANK_NO, a.UUPDATE_TYPE_C, COUNT(a.UUPDATE_TYPE_C) NUMS FROM "+tableName[ai]+" a, WTT01 w"
								+" WHERE w.MUSER_ID=a.USER_ID_C AND a.UPDATE_DATE_C BETWEEN to_date('"+startQueryDate
								+"', 'yyyy/mm/dd') AND to_date('"+endQueryDate+"', 'yyyy/mm/dd')";
							if(!bankType.equals("ALL")){
								sqlCmd+=" AND w.BANK_TYPE='"+bankType+"'";
							}
							sqlCmd+=" GROUP BY w.TBANK_NO, a.UUPDATE_TYPE_C ORDER BY w.TBANK_NO";
						}else{
							sqlCmd="SELECT w.TBANK_NO, a.UPDATE_TYPE_C, COUNT(a.UPDATE_TYPE_C) NUMS FROM "+tableName[ai]+" a, WTT01 w"
								+" WHERE w.MUSER_ID=a.USER_ID_C AND a.UPDATE_DATE_C BETWEEN to_date('"+startQueryDate
								+"', 'yyyy/mm/dd') AND to_date('"+endQueryDate+"', 'yyyy/mm/dd')";
							if(!bankType.equals("ALL")){
								sqlCmd+=" AND w.BANK_TYPE='"+bankType+"'";
							}
							sqlCmd+=" GROUP BY w.TBANK_NO, a.UPDATE_TYPE_C ORDER BY w.TBANK_NO";
						}
						if(debug) System.out.println("sqlCmd="+sqlCmd);
						rs = st.executeQuery(sqlCmd);
						if(tableName[ai].equals("WTT07")){
							while(rs.next()){
								long dataValueTmp[] = new long[3];
								String bankNoTmp = rs.getString(1)==null?"empty":rs.getString(1);
								if(bankNoTmp.trim().length()==0)
									bankNoTmp="empty";
								long nums = rs.getLong(2);
								if(debug) System.out.println("resultMap.get("+bankNoTmp+"_"+queryYear+queryMonth+")="+resultMap.get(bankNoTmp+"_"+queryYear+queryMonth));
								if(resultMap.get(bankNoTmp+"_"+queryYear+queryMonth)!=null){
									dataValueTmp = (long[])resultMap.get(bankNoTmp+"_"+queryYear+queryMonth);
								}
								if(debug) System.out.print("加到 L:"+dataValueTmp[2]+"+"+nums);
								dataValueTmp[2]+=nums;
								if(debug) System.out.println("="+dataValueTmp[2]);
								if(debug) System.out.println("加到map內: key="+bankNoTmp+"_"+queryYear+queryMonth);
								resultMap.put(bankNoTmp+"_"+queryYear+queryMonth, dataValueTmp);
							}
						}else{
							while(rs.next()){
								long dataValueTmp[] = new long[3];
								String bankNoTmp = rs.getString(1)==null?"empty":rs.getString(1);
								if(bankNoTmp.trim().length()==0)
									bankNoTmp="empty";
								String changeType = rs.getString(2);
								long nums = rs.getLong(3);
								if(debug) System.out.println("resultMap.get("+bankNoTmp+"_"+queryYear+queryMonth+")="+resultMap.get(bankNoTmp+"_"+queryYear+queryMonth));
								if(resultMap.get(bankNoTmp+"_"+queryYear+queryMonth)!=null){
									dataValueTmp = (long[])resultMap.get(bankNoTmp+"_"+queryYear+queryMonth);
								}
								if(changeType.equals("U")){
									if(debug) System.out.print("加到 U: "+dataValueTmp[0]+"+"+nums);
									dataValueTmp[0]+=nums;
									if(debug) System.out.println("="+dataValueTmp[0]);
								}else if(changeType.equals("D")){
									if(debug) System.out.print("加到 D:"+dataValueTmp[1]+"+"+nums);
									dataValueTmp[1]+=nums;
									if(debug) System.out.println("="+dataValueTmp[1]);
								}else if(changeType.equals("L")){
									if(debug) System.out.print("加到 L:"+dataValueTmp[2]+"+"+nums);
									dataValueTmp[2]+=nums;
									if(debug) System.out.println("="+dataValueTmp[2]);
								}
								if(debug) System.out.println("加到map內: key="+bankNoTmp+"_"+queryYear+queryMonth);
								resultMap.put(bankNoTmp+"_"+queryYear+queryMonth, dataValueTmp);
							}
						}
					}
				}else{
					for(int ai=0;ai<tableName.length;ai++){
						sqlCmd="SELECT BANK_NO, UPDATE_NUM, DELETE_NUM, DOWNLOAD_NUM FROM STATISTICS_BAK"
							+" WHERE TB_NAME='"+tableName[ai]+"' AND M_YEAR="+(Integer.parseInt(queryYear)-1911)
							+" AND M_MONTH="+queryMonth;
						if(!bankType.equals("ALL")){
							sqlCmd+=" AND BANK_TYPE='"+bankType+"'";
						}
						sqlCmd+=" ORDER BY BANK_NO";
						if(debug) System.out.println("sqlCmd="+sqlCmd);
						rs = st.executeQuery(sqlCmd);
						while(rs.next()){
							String bankNoTmp = rs.getString(1)==null?"empty":rs.getString(1);
							if(bankNoTmp.trim().length()==0)
								bankNoTmp="empty";
							long updateNums = rs.getLong(2);
							long deleteNums = rs.getLong(3);
							long downloadNums = rs.getLong(4);
							if(debug) System.out.println("updateNums="+updateNums+";deleteNums="+deleteNums+";downloadNums="+downloadNums);
							//long dataValueTmp[] = new long[]{updateNums,deleteNums,downloadNums};
							long dataValueTmp[] = new long[3];
							if(debug) System.out.println("resultMap.get("+bankNoTmp+"_"+queryYear+queryMonth+")="+resultMap.get(bankNoTmp+"_"+queryYear+queryMonth));
							if(resultMap.get(bankNoTmp+"_"+queryYear+queryMonth)!=null){
								dataValueTmp = (long[])resultMap.get(bankNoTmp+"_"+queryYear+queryMonth);
							}
							if(debug) System.out.print("加到 U: "+dataValueTmp[0]+"+"+updateNums);
							dataValueTmp[0]+=updateNums;
							if(debug) System.out.println("="+dataValueTmp[0]);
							if(debug) System.out.print("加到 D: "+dataValueTmp[1]+"+"+deleteNums);
							dataValueTmp[1]+=deleteNums;
							if(debug) System.out.println("="+dataValueTmp[1]);
							if(debug) System.out.print("加到 L: "+dataValueTmp[2]+"+"+downloadNums);
							dataValueTmp[2]+=downloadNums;
							if(debug) System.out.println("="+dataValueTmp[2]);
							resultMap.put(bankNoTmp+"_"+queryYear+queryMonth, dataValueTmp);
						}
					}
					rs.close();
					pst.close();
				}
			}
			HSSFWorkbook wb = new HSSFWorkbook();
			HSSFSheet sheet = wb.createSheet("New Sheet 1"); //讀取第一個工作表，宣告其為sheet
			HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
			sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
			//設定頁面符合列印大小
			sheet.setAutobreaks(false);//自動分頁
			ps.setScale((short) 75); //列印縮放百分比
			//ps.setFitWidth((short)1);
			ps.setPaperSize( (short) 9); //設定紙張大小 A4
			HSSFRow row = null; //宣告一列
			reportUtil reportUtil = new reportUtil();
			HSSFFont defaultFont = wb.createFont();
			defaultFont.setFontHeightInPoints((short)12);
			defaultStyle = reportUtil.getDefaultStyle(wb);//有框內文置中
			defaultStyle.setFont(defaultFont);
			rightStyle = reportUtil.getRightStyle(wb);//有框內文置右
			rightStyle.setFont(defaultFont);
			leftStyle = reportUtil.getLeftStyle(wb);//有框內文置左
			leftStyle.setFont(defaultFont);
    		noBorderDefaultStyle = reportUtil.getNoBorderDefaultStyle(wb);//無框內文置中
    		noBorderDefaultStyle.setFont(defaultFont);
    		noBorderLeftStyle = reportUtil.getNoBorderLeftStyle(wb);//無框內文置左
    		noBorderLeftStyle.setFont(defaultFont);
			reportUtil.setDefaultStyle(defaultStyle);
			reportUtil.setNoBorderDefaultStyle(noBorderDefaultStyle);			
    		titleStyle = reportUtil.getTitleStyle(wb); //標題用
    		//設定欄位長度----Start 
    		short[] columnLen = new short[totalColumnNums];
    		columnLen[0]=28;
    		for(int ai=1; ai<columnLen.length; ai++){
    			columnLen[ai]=5;
    		}
    		for(int ai=0; ai<columnLen.length; ai++){
    			sheet.setColumnWidth((short)ai, (short)(256*(columnLen[ai]+4)));
    		}
    		//設定欄位長度----End 
    		//=== 先填Title Start -----
			row = sheet.createRow(0);
			int repeatTime = Math.abs(totalColumnNums/9);
			if(totalColumnNums>9 && (totalColumnNums%9)>0){
				repeatTime++;
			}
			if(debug) System.out.println("repeatTime="+repeatTime);
			for(int ri=0;ri<repeatTime;ri++){
				int startIndex=0;
				if(ri>0){
					startIndex=(ri*9)+1;
				}
				reportUtil.createCell(wb, row, (short)startIndex, "稽核記錄統計總表", titleStyle);
				int endIndex=(ri+1)*9;
				for(int ci=(startIndex+1); ci<endIndex;ci++){
					reportUtil.createCell( wb, row, (short)ci, "", titleStyle);
				}
				if(debug) System.out.println(ri+":startIndex="+startIndex+";endIndex="+endIndex);
				sheet.addMergedRegion(new Region((short)0, (short)startIndex, (short)0, (short)endIndex));
			}
			row = sheet.createRow(1);
			HSSFRow row2 = sheet.createRow(2);
			reportUtil.createCell(wb, row, (short)0, "時間", rightStyle);
			reportUtil.createCell(wb, row2, (short)0, "使用者", leftStyle);
			for(int mi=0;mi<monthIndex;mi++){
				if((Integer.parseInt(startMonth)+mi)>12){
					queryYear = String.valueOf(Integer.parseInt(startYear)+1);
					queryMonth=(Integer.parseInt(startMonth)+mi-12)+"";
				}else{
					queryYear = startYear;
					queryMonth = String.valueOf(Integer.parseInt(startMonth)+mi);
				}
				reportUtil.createCell( wb, row, (short)(1+(mi*3)), String.valueOf(Integer.parseInt(queryYear)-1911)+"年"+queryMonth+"月", defaultStyle);
				reportUtil.createCell( wb, row, (short)(2+(mi*3)), "", defaultStyle);
				reportUtil.createCell( wb, row, (short)(3+(mi*3)), "", defaultStyle);
				sheet.addMergedRegion(new Region((short)1, (short)(1+(mi*3)), (short)1, (short)(3+(mi*3))));
				reportUtil.createCell( wb, row2, (short)(1+(mi*3)), "異動", defaultStyle);
				reportUtil.createCell( wb, row2, (short)(2+(mi*3)), "刪除", defaultStyle);
				reportUtil.createCell( wb, row2, (short)(3+(mi*3)), "下載", defaultStyle);
			}
			wb.setRepeatingRowsAndColumns(0, 0, 0, 0, (short)2); //設定表頭 為固定 先設欄的起始再設列的起始
			//=== 先填Title End -----
			//判斷dbData.size()是不是0，是的話表示沒有資料
			if (bankList.size()==0) {
				row = sheet.getRow(3);
				reportUtil.createCell( wb, row, (short)0, "查無資料", titleStyle);
				for(int ci=1; ci<totalColumnNums;ci++){
					reportUtil.createCell( wb, row, (short)ci, "", titleStyle);
				}
				sheet.addMergedRegion(new Region((short)3, (short)0, (short)3, totalColumnNums));
			}else {
				//設定儲存格資料============================================
				for(int li=0;li<bankList.size(); li++){
					String[] bankData = (String[])bankList.get(li);
					String bankNoTmp = bankData[0];
					String bankNameTmp = bankData[1];
					for(int mi=0;mi<monthIndex;mi++){
						if((Integer.parseInt(startMonth)+mi)>12){
							queryYear = String.valueOf(Integer.parseInt(startYear)+1);
							queryMonth=(Integer.parseInt(startMonth)+mi-12)+"";
						}else{
							queryYear = startYear;
							queryMonth = String.valueOf(Integer.parseInt(startMonth)+mi);
						}
						short cellNo = (short)(1+(mi*3));
						row=sheet.createRow(li+3);
						//=== 處理第一列時,需將機構名稱寫入
						if(mi==0){
							reportUtil.createCell(wb, row, (short)0, bankNameTmp, leftStyle);
						}
						long dataValueTmp[] = new long[3];
						if(debug) System.out.println("get value from map: Key="+bankNoTmp+"_"+queryYear+queryMonth+"; value="+resultMap.get(bankNoTmp+"_"+queryYear+queryMonth));
						if(resultMap.get(bankNoTmp+"_"+queryYear+queryMonth)!=null){
							dataValueTmp = (long[])resultMap.get(bankNoTmp+"_"+queryYear+queryMonth);
						}
						for(int ci=0; ci<3; ci++){
							reportUtil.createCell(wb, row, (short)(cellNo+ci), String.valueOf(dataValueTmp[ci]), rightStyle);
						}
					}
				}//end of for
			} //end of else ((DataObject) dbData.get(0)).getValue("m_year") == null
			File reportDir = new File(Utility.getProperties("reportDir"));
			if (!reportDir.exists()) {
				if (!Utility.mkdirs(Utility.getProperties("reportDir"))) {
					errMsg += Utility.getProperties("reportDir") + "目錄新增失敗";
				}
			}
			FileOutputStream fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + "稽核記錄統計總表.xls");
			HSSFFooter footer = sheet.getFooter();
			footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
			wb.write(fout);
			//儲存
			fout.close();
		}catch (Exception e) {
			System.out.println("//RptCG001WA createRpt() Have Error.....");
			e.printStackTrace();
			System.out.println("//-------------------------------------");
		}finally{
			try{
				conn.close();
			}catch(Exception sqlEx){
				conn=null;
			}
		}
		if(debug) System.out.println("RptCG001WA createRpt() Debug End ...");
		return errMsg;
	}
}
