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

public class RptCG060WA {

	public static String createRpt(String startYear, String startMonth, String endYear, String endMonth, String bankType) {
		boolean debug = false;
		startYear = Integer.parseInt(startYear)+1911+"";
		endYear = Integer.parseInt(endYear)+1911+"";
		String[] endDayStr = new String[]{"31","28","31","30","31","30","31","31","30","31","30","31"};
		if(debug) System.out.println("RptCG060WA createRpt() Debug Start ...");
		String errMsg = "";
		String sqlCmd = "";
		Connection conn = null;
		PreparedStatement pst = null;
		ResultSet rs = null;
		HSSFCellStyle defaultStyle;
	    HSSFCellStyle rightStyle;
	    HSSFCellStyle leftStyle;
	    HSSFCellStyle noBorderDefaultStyle;
	    HSSFCellStyle titleStyle;
		Calendar calendar = Calendar.getInstance();
		String nowDate = calendar.get(1) + "/" + (calendar.get(2) + 1);
		if(debug) System.out.println("nowDate=" + nowDate);
		calendar.add(2,-2);
		calendar.set(calendar.get(1),calendar.get(2),1);
		if(debug) System.out.println("calendar date="+calendar.get(1)+"/"+(calendar.get(2)+1)+"/"+calendar.get(5));
		try {
			conn =(new RdbCommonDao("")).newConnection();
			int monthIndex = (Integer.parseInt(endYear)-Integer.parseInt(startYear))*12+(Integer.parseInt(endMonth)-Integer.parseInt(startMonth))+1;//一共查詢幾個月份
			int xArrayLength = monthIndex*2;//陣列長度
			short totalColumnNums = (short)xArrayLength;
			if(debug) System.out.println("monthIndex="+monthIndex+";xArrayLength="+xArrayLength);
			String queryYear="",queryMonth="";
			List bankList = new ArrayList();
//==== 先查詢出所有機關名稱
			sqlCmd="SELECT BANK_NO, BANK_NAME FROM BN01 WHERE 1=1";
			if(!bankType.equals("ALL")){
				sqlCmd+=" AND BANK_TYPE=?";
			}
			sqlCmd+=" ORDER BY BANK_NAME";
			if(debug) System.out.println("sqlCmd="+sqlCmd);
			pst = conn.prepareStatement(sqlCmd);
			if(!bankType.equals("ALL")){
				pst.setString(1, bankType);
			}
			rs=pst.executeQuery();
			while(rs.next()){
				String[] bankTmp = new String[]{rs.getString(1),rs.getString(2)};
				bankList.add(bankTmp);
			}
			if(bankType.equals("2")||bankType.equals("ALL")){
				String[] bankTmp = new String[]{"empty","農業金融局"};
				bankList.add(bankTmp);
			}
			HashMap resultMap = new HashMap();//裡面以bank_no+查詢年月當key,裡面放String[2]
			HashMap threeMonthMap = new HashMap();//裡面以bank_no,裡面放三個月內登入次數
//=== 依年月查詢
			for(int mi=0;mi<monthIndex;mi++){
				if((Integer.parseInt(startMonth)+mi)>12){
					queryYear = String.valueOf(Integer.parseInt(startYear)+1);
					queryMonth=(Integer.parseInt(startMonth)+mi-12)+"";
				}else{
					queryYear = startYear;
					queryMonth = String.valueOf(Integer.parseInt(startMonth)+mi);
				}
				if (((Integer.parseInt(queryYear) % 4 == 0) && (Integer.parseInt(queryYear) % 100 != 0)) || (Integer.parseInt(queryYear) % 400 == 0))
					endDayStr[1] = "29";
				else
					endDayStr[1] = "28";
				Calendar queryDate = Calendar.getInstance();
				queryDate.set(Integer.parseInt(queryYear),Integer.parseInt(queryMonth)-1,5);
				if(debug) System.out.println("queryYear="+queryYear+";queryMonth="+queryMonth+"; queryDate.after(calendar)="+queryDate.after(calendar));
				if(queryDate.after(calendar)){
					sqlCmd="SELECT w.TBANK_NO, COUNT(a.SERIALNO) NUMS FROM WTT06 a, WTT01 w"
						+" WHERE w.MUSER_ID=a.MUSER_ID AND TO_CHAR(a.INPUT_DATE, 'yyyy/mm')  = ?";
					if(!bankType.equals("ALL")){
						sqlCmd+=" AND w.BANK_TYPE=?";
					}
					sqlCmd+=" GROUP BY w.TBANK_NO ORDER BY w.TBANK_NO";
					if(debug) System.out.println("sqlCmd="+sqlCmd);
					pst = conn.prepareStatement(sqlCmd);
					String startQueryDate = Integer.parseInt(queryYear)+"/"+queryMonth+"/1";
					String endQueryDate = Integer.parseInt(queryYear)+"/"+queryMonth+"/"+endDayStr[(Integer.parseInt(queryMonth))-1];

                                        if(queryMonth.length() < 2){
                                          startQueryDate = queryYear + "/0" + queryMonth;
                                        }else{
                                          startQueryDate = queryYear + "/" + queryMonth;
                                        }
					pst.setString(1, startQueryDate);
					if(!bankType.equals("ALL")){
						pst.setString(2, bankType);
					}
					//if(debug) System.out.println("startQueryDate="+startQueryDate+";endQueryDate="+endQueryDate);
					rs = pst.executeQuery();
					while(rs.next()){
						String dataValueTmp[] = new String[2];
						String bankNoTmp = rs.getString(1)==null?"empty":rs.getString(1);
						if(bankNoTmp.trim().length()==0)
							bankNoTmp="empty";
						String nums = rs.getString(2);
						if(resultMap.get(bankNoTmp+"_"+queryYear+queryMonth)!=null){
							dataValueTmp = (String[])resultMap.get(bankNoTmp+"_"+queryYear+queryMonth);
						}
						dataValueTmp[0]=nums;
						dataValueTmp[1]="";
						if(debug) System.out.println("加到map內: key="+bankNoTmp+"_"+queryYear+queryMonth);
						resultMap.put(bankNoTmp+"_"+queryYear+queryMonth, dataValueTmp);
					}

                                        //查詢的日期
                                        Calendar befor3MDate = Calendar.getInstance();
                                        befor3MDate.set(Integer.parseInt(queryYear),Integer.parseInt(queryMonth) - 1, 1);
                                        endQueryDate = befor3MDate.get(1) + "/" + (befor3MDate.get(2) + 1);
                                        //前三個月
                                        befor3MDate.add(2, -2);
                                        startQueryDate = befor3MDate.get(1) + "/" + (befor3MDate.get(2) + 1);
                                        if (debug) System.out.println("startQueryDate=" + startQueryDate +";endQueryDate=" + endQueryDate);

                                        //判斷三個月內是否有登入
                                        if(nowDate.equals(queryYear + "/" + queryMonth)){
                                          sqlCmd = "SELECT w.TBANK_NO, COUNT(a.SERIALNO) NUMS FROM WTT06 a, WTT01 w"
                                              + " WHERE w.MUSER_ID=a.MUSER_ID AND a.INPUT_DATE BETWEEN to_date(?, 'yyyy/mm') AND to_date(?, 'yyyy/mm')";
                                          if (!bankType.equals("ALL")) {
                                            sqlCmd += " AND w.BANK_TYPE=?";
                                          }
                                          sqlCmd += " GROUP BY w.TBANK_NO ORDER BY w.TBANK_NO";
                                          if(debug) System.out.println("sqlCmd="+sqlCmd);
                                          pst = conn.prepareStatement(sqlCmd);
                                          if (debug) System.out.println("startQueryDate=" + startQueryDate +";endQueryDate=" + endQueryDate);

                                          pst.setString(1, startQueryDate);
                                          pst.setString(2, endQueryDate);
                                          if (!bankType.equals("ALL")) {
                                            pst.setString(3, bankType);
                                          }
                                          rs = pst.executeQuery();
                                          while (rs.next()) {
                                            threeMonthMap.put(rs.getString(1) + "_" + queryYear + queryMonth,rs.getString(2));
                                          }
                                        }else{
                                          sqlCmd = "SELECT w.TBANK_NO, COUNT(a.SERIALNO) NUMS FROM WTT06 a, WTT01 w"
                                              + " WHERE w.MUSER_ID=a.MUSER_ID AND a.INPUT_DATE BETWEEN to_date(?, 'yyyy/mm') AND to_date(?, 'yyyy/mm')";
                                          if (!bankType.equals("ALL")) {
                                            sqlCmd += " AND w.BANK_TYPE=?";
                                          }
                                          sqlCmd += " GROUP BY w.TBANK_NO ORDER BY w.TBANK_NO";
                                          if(debug) System.out.println("sqlCmd="+sqlCmd);
                                          pst = conn.prepareStatement(sqlCmd);
                                          if (debug) System.out.println("startQueryDate=" + startQueryDate +";endQueryDate=" + endQueryDate);
                                          pst.setString(1, startQueryDate);
                                          pst.setString(2, endQueryDate);
                                          if (!bankType.equals("ALL")) {
                                            pst.setString(3, bankType);
                                          }
                                          rs = pst.executeQuery();
                                          while (rs.next()) {
                                            threeMonthMap.put(rs.getString(1) + "_" + queryYear + queryMonth,rs.getString(2));
                                          }

                                          //===========================================
                                          sqlCmd = "SELECT BANK_NO, LOGIN_NUM FROM STATISTICS_BAK WHERE TB_NAME=?"
                                              + " AND TO_DATE(CONCAT(TO_NUMBER(M_YEAR+1911), CONCAT('/', M_MONTH)),'YYYY/MM')"
                                              + " BETWEEN TO_DATE(?, 'YYYY/MM') AND TO_DATE(?, 'YYYY/MM')";
                                          if (!bankType.equals("ALL")) {
                                            sqlCmd += " AND BANK_TYPE=?";
                                          }
                                          sqlCmd += " ORDER BY BANK_NO";
                                          if(debug) System.out.println("sqlCmd="+sqlCmd);
                                          pst = conn.prepareStatement(sqlCmd);
                                          pst.setString(1, "WTT06");
                                          pst.setString(2, startQueryDate);
                                          pst.setString(3, endQueryDate);
                                          if (!bankType.equals("ALL")) {
                                            pst.setString(4, bankType);
                                          }
                                          rs = pst.executeQuery();
                                          while (rs.next()) {
                                            String bankNo2 = rs.getString(1);
                                            String lginNums = rs.getString(2);
                                            if (threeMonthMap.get(bankNo2 + "_" + queryYear + queryMonth) != null) {
                                              lginNums = Integer.parseInt(lginNums) + Integer.parseInt( (String) threeMonthMap.get(bankNo2 + "_" + queryYear + queryMonth)) + "";
                                            }
                                            threeMonthMap.put(bankNo2 + "_" + queryYear + queryMonth, lginNums);
                                          }
                                        }
				}else{
					sqlCmd="SELECT BANK_NO, LOGIN_NUM FROM STATISTICS_BAK "
                                            + "WHERE TB_NAME=? AND TO_NUMBER(M_YEAR+1911)=? AND M_MONTH=?";
					if(!bankType.equals("ALL")){
						sqlCmd+=" AND BANK_TYPE=?";
					}
					sqlCmd+=" ORDER BY BANK_NO";
					if(debug) System.out.println("sqlCmd="+sqlCmd);
					pst = conn.prepareStatement(sqlCmd);
					pst.setString(1, "WTT06");
					pst.setString(2,queryYear);
					pst.setString(3,queryMonth);
					if(!bankType.equals("ALL")){
						pst.setString(4, bankType);
					}
					rs = pst.executeQuery();
					while(rs.next()){
						String bankNoTmp = rs.getString(1)==null?"empty":rs.getString(1);
						if(bankNoTmp.trim().length()==0)
							bankNoTmp="empty";
						String loginNums = rs.getString(2);
						String dataValueTmp[] = new String[]{loginNums,""};
						resultMap.put(bankNoTmp+"_"+queryYear+queryMonth, dataValueTmp);
					}
					sqlCmd="SELECT BANK_NO, LOGIN_NUM FROM STATISTICS_BAK WHERE TB_NAME=?";
					sqlCmd+=" AND TO_DATE(CONCAT(TO_NUMBER(M_YEAR+1911), CONCAT('/', CONCAT(M_MONTH,'/1'))),'YYYY/MM/DD')";
					sqlCmd+=" BETWEEN TO_DATE(?, 'YYYY/MM/DD') AND TO_DATE(?, 'YYYY/MM/DD')";
					if(!bankType.equals("ALL")){
						sqlCmd+=" AND BANK_TYPE=?";
					}
					sqlCmd+=" ORDER BY BANK_NO";
					if(debug) System.out.println("sqlCmd="+sqlCmd);
					pst = conn.prepareStatement(sqlCmd);
					pst.setString(1, "WTT06");
					Calendar befor3MDate = Calendar.getInstance();
					befor3MDate.set(Integer.parseInt(queryYear),Integer.parseInt(queryMonth)-1,1);
					befor3MDate.add(2, -3);
					String startQueryDate = befor3MDate.get(1)+"/"+(befor3MDate.get(2)+1)+"/1";
					befor3MDate.add(2, 2);
					String endQueryDate = befor3MDate.get(1)+"/"+(befor3MDate.get(2)+1)+"/"+endDayStr[befor3MDate.get(2)];
					if(debug) System.out.println("startQueryDate="+startQueryDate+";endQueryDate="+endQueryDate);
					pst.setString(2, startQueryDate);
					pst.setString(3, endQueryDate);
					if(!bankType.equals("ALL")){
						pst.setString(4, bankType);
					}
					rs = pst.executeQuery();
					while(rs.next()){
						String bankNo2 = rs.getString(1);
						String lginNums = rs.getString(2);
						if(threeMonthMap.get(bankNo2+"_"+queryYear+queryMonth)!=null){
							lginNums = Integer.parseInt(lginNums)+ Integer.parseInt((String)threeMonthMap.get(bankNo2+"_"+queryYear+queryMonth))+"";
						}
						threeMonthMap.put(bankNo2+"_"+queryYear+queryMonth,lginNums);
					}
				}
			}
			HSSFWorkbook wb = new HSSFWorkbook();
			HSSFSheet sheet = wb.createSheet("sheet1"); //讀取第一個工作表，宣告其為sheet
			HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
			//設定頁面符合列印大小
			sheet.setAutobreaks(false);
			ps.setScale((short) 80); //列印縮放百分比
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
    		HSSFCellStyle smallFontStyle = reportUtil.getDefaultStyle(wb);
    		HSSFFont smallFont = wb.createFont();
    		smallFont.setFontHeightInPoints((short)10);
    		smallFontStyle.setFont(smallFont);
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
			int repeatTime = Math.abs(totalColumnNums/8);
			if(totalColumnNums>8 && (totalColumnNums%8)>0){
				repeatTime++;
			}
			if(debug) System.out.println("repeatTime="+repeatTime);
			for(int ri=0;ri<repeatTime;ri++){
				int startIndex=0;
				if(ri>0){
					startIndex=(ri*8)+1;
				}
				reportUtil.createCell(wb, row, (short)startIndex, "使用者帳號期間未使用總表", titleStyle);
				int endIndex=(ri+1)*8;
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
				reportUtil.createCell( wb, row, (short)(1+(mi*2)), String.valueOf(Integer.parseInt(queryYear)-1911)+"年"+queryMonth+"月", defaultStyle);
				reportUtil.createCell( wb, row, (short)(2+(mi*2)), "", defaultStyle);
				sheet.addMergedRegion(new Region((short)1, (short)(1+(mi*2)), (short)1, (short)(2+(mi*2))));
				reportUtil.createCell( wb, row2, (short)(1+(mi*2)), "登入次數", smallFontStyle);
				reportUtil.createCell( wb, row2, (short)(2+(mi*2)), "狀態", defaultStyle);
			}
			wb.setRepeatingRowsAndColumns(0, 0, 0, 0, (short)2); //設定表頭 為固定 先設欄的起始再設列的起始
			//=== 先填Title End -----
			//判斷dbData.size()是不是0，是的話表示沒有資料
			short rowNo=3;
			if (bankList.size()==0) {
				row = sheet.getRow(rowNo++);
				reportUtil.createCell( wb, row, (short)0, "查無資料", titleStyle);
				for(int ci=1; ci<=totalColumnNums;ci++){
					reportUtil.createCell( wb, row, (short)ci, "", titleStyle);
				}
				sheet.addMergedRegion(new Region((short)0, (short)0, (short)0, totalColumnNums));
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
						short cellNo = (short)(1+(mi*2));
						rowNo=(short)(li+3);
						row=sheet.createRow(rowNo);
						//=== 處理第一列時,需將機構名稱寫入
						if(mi==0){
							reportUtil.createCell(wb, row, (short)0, bankNameTmp, leftStyle);
						}
						String dataValueTmp[] = new String[2];
						if(debug) System.out.println("get value from map: Key="+bankNoTmp+"_"+queryYear+queryMonth+"; value="+resultMap.get(bankNoTmp+"_"+queryYear+queryMonth));
						if(resultMap.get(bankNoTmp+"_"+queryYear+queryMonth)!=null){
							dataValueTmp = (String[])resultMap.get(bankNoTmp+"_"+queryYear+queryMonth);
						}else{
							dataValueTmp[0]="0";
							dataValueTmp[1]="N";
						}
						reportUtil.createCell(wb, row, (short)(cellNo), String.valueOf(dataValueTmp[0]), rightStyle);
						if(dataValueTmp[0].equals("0") || dataValueTmp[1].equals("N")){
							//System.out.println("3 month:"+threeMonthMap.get(bankNoTmp+"_"+queryYear+queryMonth));
							if(threeMonthMap.get(bankNoTmp+"_"+queryYear+queryMonth)!=null){
								reportUtil.createCell(wb, row, (short)(cellNo+1), "☆", rightStyle);
							}else{
								reportUtil.createCell(wb, row, (short)(cellNo+1), "☆☆☆", smallFontStyle);
							}
						}else{
							reportUtil.createCell(wb, row, (short)(cellNo+1), "", rightStyle);
						}
					}
				}//end of for
			}
			row = sheet.createRow(rowNo);
			reportUtil.createCell( wb, row, (short)0, "註解：狀態項目表示若使用者帳號該月未登入則為「☆」，超過三個月都未登入則顯示為「☆☆☆」，空白則表示使用狀態正常。", leftStyle);
			for(int ci=1; ci<=totalColumnNums;ci++){
				reportUtil.createCell( wb, row, (short)ci, "", leftStyle);
			}
			sheet.addMergedRegion(new Region((short)rowNo, (short)0, (short)rowNo, totalColumnNums));
			File reportDir = new File(Utility.getProperties("reportDir"));
			if (!reportDir.exists()) {
				if (!Utility.mkdirs(Utility.getProperties("reportDir"))) {
					errMsg += Utility.getProperties("reportDir") + "目錄新增失敗";
				}
			}
			FileOutputStream fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + "使用者帳號期間未使用總表.xls");
			HSSFFooter footer = sheet.getFooter();
			footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
			wb.write(fout);
			//儲存
			fout.close();
		}catch (Exception e) {
			System.out.println("//RptCG060WA createRpt() Have Error.....");
			e.printStackTrace();
			System.out.println("//-------------------------------------");
		}finally{
			try{
				conn.close();
			}catch(Exception sqlEx){
				conn=null;
			}
		}
		if(debug) System.out.println("RptCG060WA createRpt() Debug End ...");
		return errMsg;
	}
}
