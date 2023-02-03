/*
 * Created on 2006/11/06 by ABYSS Allen
 * RptCG001WA 稽核記錄統計總表
 */

package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.poifs.filesystem.*;
import java.io.*;
import java.sql.*;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.dao.RdbCommonDao;
import com.tradevan.util.report.reportUtil;

public class RptCG001WB {

	public static String createRpt(String startYear,String startMonth, String endYear, String endMonth, String reportType
			, String bankType, String[] tableName, HashMap staticBankNameMap) {
		boolean debug = false;
		startYear = Integer.parseInt(startYear)+1911+"";
		endYear = Integer.parseInt(endYear)+1911+"";
		String[] endDayStr = new String[]{"31","28","31","30","31","30","31","31","30","31","30","31"};
		if(debug) System.out.println("----------------------RptCG001WB createRpt() Debug Start ...");
		String errMsg = "";
		String sqlCmd = "";
		Connection conn = null;
		PreparedStatement pst = null;
		Statement st = null;
		ResultSet rs = null;HSSFCellStyle defaultStyle;
	    HSSFCellStyle rightStyle;
	    HSSFCellStyle leftStyle;
	    HSSFCellStyle noBorderDefaultStyle;
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
			List bankList = new ArrayList();//存放機關名稱
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
			if(bankType.equals("2") || bankType.equals("ALL")){
				String[] bankTmp = new String[]{"empty","農業金融局"};
				bankList.add(bankTmp);
			}
			HashMap userListMap = new HashMap();//計劃裡面以Bank_no當Key,裡面放List
			HashMap resultMap = new HashMap();//計劃裡面以bank_no_查詢年月_UserId當key,裡面放long[3]的值
			st = conn.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
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
					for(int ai=0;ai<tableName.length;ai++){
						if(tableName[ai].equals("WTT07")){
							sqlCmd="SELECT w.TBANK_NO, w.MUSER_NAME, a.MUSER_ID, COUNT(a.RESULT_P) NUMS FROM WTT07 a, WTT01 w"
								+" WHERE w.MUSER_ID=a.MUSER_ID AND RESULT_P='XOO'"
								+" AND a.INPUT_DATE BETWEEN to_date('"+startQueryDate+"', 'yyyy/mm/dd') AND to_date('"
								+endQueryDate+"', 'yyyy/mm/dd')";
							if(!bankType.equals("ALL")){
								sqlCmd+=" AND w.BANK_TYPE='"+bankType+"'";
							}
							sqlCmd+=" GROUP BY w.TBANK_NO, w.MUSER_NAME, a.MUSER_ID ORDER BY w.TBANK_NO";
						}else if(tableName[ai].equals("MUSER_DATA_log")){
							sqlCmd="SELECT w.TBANK_NO, a.UUPDATE_TYPE_C, a.USER_ID_C, a.USER_NAME_C, COUNT(a.UUPDATE_TYPE_C) NUMS"
								+" FROM "+tableName[ai]+" a, WTT01 w"
								+" WHERE w.MUSER_ID=a.USER_ID_C AND a.UPDATE_DATE_C BETWEEN to_date('"
								+startQueryDate+"', 'yyyy/mm/dd') AND to_date('"+endQueryDate+"', 'yyyy/mm/dd')";
							if(!bankType.equals("ALL")){
								sqlCmd+=" AND w.BANK_TYPE='"+bankType+"'";
							}
							sqlCmd+=" GROUP BY w.TBANK_NO, a.UUPDATE_TYPE_C, a.USER_ID_C, a.USER_NAME_C ORDER BY w.TBANK_NO, a.USER_NAME_C";
						}else{
							sqlCmd="SELECT w.TBANK_NO, a.UPDATE_TYPE_C, a.USER_ID_C, a.USER_NAME_C, COUNT(a.UPDATE_TYPE_C) NUMS"
								+" FROM "+tableName[ai]+" a, WTT01 w"
								+" WHERE w.MUSER_ID=a.USER_ID_C AND a.UPDATE_DATE_C BETWEEN to_date('"+startQueryDate
								+"', 'yyyy/mm/dd') AND to_date('"+endQueryDate+"', 'yyyy/mm/dd')";
							if(!bankType.equals("ALL")){
								sqlCmd+=" AND w.BANK_TYPE='"+bankType+"'";
							}
							sqlCmd+=" GROUP BY w.TBANK_NO, a.UPDATE_TYPE_C, a.USER_ID_C, a.USER_NAME_C ORDER BY w.TBANK_NO, a.USER_NAME_C";
						}
						if(debug) System.out.println("1. sqlCmd="+sqlCmd);
						rs = st.executeQuery(sqlCmd);
						List userDataList = null;
						if(tableName[ai].equals("WTT07")){
							while(rs.next()){
								long dataValueTmp[] = new long[3];
								String bankNoTmp = rs.getString(1)==null?"empty":rs.getString(1);
								if(bankNoTmp.trim().length()==0)
									bankNoTmp="empty";
								String userNameTmp = rs.getString(2);
								String userIdTmp = rs.getString(3);//使用者ID
								long nums = rs.getLong(4);
								if(debug) System.out.println("nums="+nums);
								//=== 先將使用者加入該機構內
								if(userListMap.get(bankNoTmp)!=null){
									userDataList = (List)userListMap.get(bankNoTmp);
								}else{
									userDataList = new ArrayList();
								}
								boolean isExit=false;
								for(int li=0;li<userDataList.size();li++){
									String[] userInfo = (String[])userDataList.get(li);
									if(userInfo[0].equals(userIdTmp)){
										isExit=true;
										break;
									}
								}
								if(!isExit){
									userDataList.add(new String[]{userIdTmp, userNameTmp});
								}
								userListMap.put(bankNoTmp, userDataList);
								//===將該機構在該月份的加總資料叫出
								long[] monthSum = null;
								if(debug) System.out.println("resultMap.get("+bankNoTmp+"_"+queryYear+queryMonth+"_SUM)="+resultMap.get(bankNoTmp+"_"+queryYear+queryMonth+"_SUM"));
								if(resultMap.get(bankNoTmp+"_"+queryYear+queryMonth+"_SUM")==null){
									monthSum=new long[3];
								}else{
									monthSum=(long[])resultMap.get(bankNoTmp+"_"+queryYear+queryMonth+"_SUM");
								}
								//===將該使用者在該月份的資料取出
								if(debug) System.out.println("resultMap.get("+bankNoTmp+"_"+queryYear+queryMonth+"_"+userIdTmp+")="+resultMap.get(bankNoTmp+"_"+queryYear+queryMonth+"_"+userIdTmp));
								if(resultMap.get(bankNoTmp+"_"+queryYear+queryMonth+"_"+userIdTmp)!=null){
									dataValueTmp=(long[])resultMap.get(bankNoTmp+"_"+queryYear+queryMonth+"_"+userIdTmp);
								}else{
									dataValueTmp=new long[3];
								}
								if(debug) System.out.print(dataValueTmp[2]+"+"+nums+"=");
								dataValueTmp[2]+=nums;
								if(debug) System.out.println(""+dataValueTmp[2]);
								if(debug) System.out.print(monthSum[2]+"+"+nums+"=");
								monthSum[2]+=nums;
								if(debug) System.out.println(""+monthSum[2]);
								resultMap.put(bankNoTmp+"_"+queryYear+queryMonth+"_"+userIdTmp, dataValueTmp);
								resultMap.put(bankNoTmp+"_"+queryYear+queryMonth+"_SUM", monthSum);
							}
						}else{
							while(rs.next()){
								long dataValueTmp[] = new long[3];
								String bankNoTmp = rs.getString(1)==null?"empty":rs.getString(1);
								if(bankNoTmp.trim().length()==0)
									bankNoTmp="empty";
								String changeType = rs.getString(2);//異動別
								String userIdTmp = rs.getString(3);//使用者ID
								String userNameTmp = rs.getString(4);//使用者名稱
								long nums = rs.getLong(5);
								if(debug) System.out.println("nums="+nums);
								//=== 先將使用者加入該機構內
								if(userListMap.get(bankNoTmp)!=null){
									userDataList = (List)userListMap.get(bankNoTmp);
								}else{
									userDataList = new ArrayList();
								}
								boolean isExit=false;
								for(int li=0;li<userDataList.size();li++){
									String[] userInfo = (String[])userDataList.get(li);
									if(userInfo[0].equals(userIdTmp)){
										isExit=true;
										break;
									}
								}
								if(!isExit){
									userDataList.add(new String[]{userIdTmp, userNameTmp});
								}
								userListMap.put(bankNoTmp, userDataList);
								//===將該機構在該月份的加總資料叫出
								long[] monthSum = null;
								if(debug) System.out.println("resultMap.get("+bankNoTmp+"_"+queryYear+queryMonth+"_SUM)="+resultMap.get(bankNoTmp+"_"+queryYear+queryMonth+"_SUM"));
								if(resultMap.get(bankNoTmp+"_"+queryYear+queryMonth+"_SUM")==null){
									monthSum=new long[3];
								}else{
									monthSum=(long[])resultMap.get(bankNoTmp+"_"+queryYear+queryMonth+"_SUM");
								}
								//===將該使用者在該月份的資料取出
								if(debug) System.out.println("resultMap.get("+bankNoTmp+"_"+queryYear+queryMonth+"_"+userIdTmp+")="+resultMap.get(bankNoTmp+"_"+queryYear+queryMonth+"_"+userIdTmp));
								if(resultMap.get(bankNoTmp+"_"+queryYear+queryMonth+"_"+userIdTmp)!=null){
									dataValueTmp=(long[])resultMap.get(bankNoTmp+"_"+queryYear+queryMonth+"_"+userIdTmp);
								}else{
									dataValueTmp=new long[3];
								}
								if(changeType.equals("U")){
									if(debug) System.out.print("U:"+dataValueTmp[0]+"+"+nums+"=");
									dataValueTmp[0]+=nums;
									if(debug) System.out.println(dataValueTmp[0]);
									if(debug) System.out.print("U:"+monthSum[0]+"+"+nums+"=");
									monthSum[0]+=nums;
									if(debug) System.out.println(monthSum[0]);
								}else if(changeType.equals("D")){
									if(debug) System.out.print("D:"+dataValueTmp[1]+"+"+nums+"=");
									dataValueTmp[1]+=nums;
									if(debug) System.out.println(dataValueTmp[1]);
									if(debug) System.out.print("D:"+monthSum[1]+"+"+nums+"=");
									monthSum[1]+=nums;
									if(debug) System.out.println(monthSum[1]);
								}else if(changeType.equals("L")){
									if(debug) System.out.print("L:"+dataValueTmp[2]+"+"+nums+"=");
									dataValueTmp[2]+=nums;
									if(debug) System.out.println(dataValueTmp[2]);
									if(debug) System.out.print("L:"+monthSum[2]+"+"+nums+"=");
									monthSum[2]+=nums;
									if(debug) System.out.println(monthSum[2]);
								}
								resultMap.put(bankNoTmp+"_"+queryYear+queryMonth+"_"+userIdTmp, dataValueTmp);
								resultMap.put(bankNoTmp+"_"+queryYear+queryMonth+"_SUM", monthSum);
							}
						}
					}
				}else{
					for(int ai=0;ai<tableName.length;ai++){
						sqlCmd="SELECT BANK_NO, UPDATE_NUM, DELETE_NUM, DOWNLOAD_NUM, USER_ID, USER_NAME FROM STATISTICS_BAK"
							+" WHERE TB_NAME='"+tableName[ai]+"' AND M_YEAR="+(Integer.parseInt(queryYear)-1911)
							+" AND M_MONTH="+queryMonth;
						if(!bankType.equals("ALL")){
							sqlCmd+=" AND BANK_TYPE='"+bankType+"'";
						}
						sqlCmd+=" ORDER BY BANK_NO, USER_NAME";
						if(debug) System.out.println("sqlCmd="+sqlCmd);
						rs = st.executeQuery(sqlCmd);
						List userDataList = new ArrayList();
						while(rs.next()){
							String bankNoTmp = rs.getString(1);
							long updateNums = rs.getLong(2);
							long deleteNums = rs.getLong(3);
							long downloadNums = rs.getLong(4);
							String userIdTmp=rs.getString(5);
							String userNameTmp=rs.getString(6);
							//long dataValueTmp[] = new long[]{updateNums,deleteNums,downloadNums};
							//=== 先將使用者加入該機構內
							if(userListMap.get(bankNoTmp)!=null){
								userDataList = (List)userListMap.get(bankNoTmp);
							}else{
								userDataList = new ArrayList();
							}
							boolean isExit=false;
							for(int li=0;li<userDataList.size();li++){
								String[] userInfo = (String[])userDataList.get(li);
								if(userInfo[0].equals(userIdTmp)){
									isExit=true;
									break;
								}
							}
							if(!isExit){
								userDataList.add(new String[]{userIdTmp, userNameTmp});
							}
							userListMap.put(bankNoTmp, userDataList);
							//===將該機構在該月份的加總資料叫出
							long[] dataValueTmp = null;
							if(resultMap.get(bankNoTmp+"_"+queryYear+queryMonth+"_"+userIdTmp)==null){
								dataValueTmp=new long[3];
							}else{
								dataValueTmp = (long[])resultMap.get(bankNoTmp+"_"+queryYear+queryMonth+"_"+userIdTmp);
							}
							dataValueTmp[0]+=updateNums;
							dataValueTmp[1]+=deleteNums;
							dataValueTmp[2]+=downloadNums;
							resultMap.put(bankNoTmp+"_"+queryYear+queryMonth+"_"+userIdTmp, dataValueTmp);
							long[] monthSum = null;
							if(resultMap.get(bankNoTmp+"_"+queryYear+queryMonth+"_SUM")==null){
								monthSum=new long[3];
							}else{
								monthSum=(long[])resultMap.get(bankNoTmp+"_"+queryYear+queryMonth+"_SUM");
							}
							monthSum[0]+=updateNums;
							monthSum[1]+=deleteNums;
							monthSum[2]+=downloadNums;
							resultMap.put(bankNoTmp+"_"+queryYear+queryMonth+"_SUM", monthSum);
						}
					}
				}
			}
			if(debug){
				Set keySet = userListMap.keySet();
				Iterator keyIter = keySet.iterator();
				while(keyIter.hasNext()){
					String keyName = (String)keyIter.next();
					List bankUserList = (List)userListMap.get(keyName);
					System.out.println("UserListMap keyName="+keyName);
					for(int li=0; li<bankUserList.size(); li++){
						String[] userInfo = (String[])bankUserList.get(li);
						System.out.println("      User id="+userInfo[0]+"; name="+userInfo[1]);
					}
				}
				keySet = resultMap.keySet();
				keyIter = keySet.iterator();
				while(keyIter.hasNext()){
					String keyName = (String)keyIter.next();
					long[] valueTmp = (long[])resultMap.get(keyName);
					System.out.println("resultMap keyName="+keyName);
					for(int ai=0; ai<valueTmp.length; ai++){
						System.out.println("      "+ai+"value="+valueTmp[ai]);
					}
				}
			}
			//設定報表表頭資料============================================
			HSSFWorkbook wb = new HSSFWorkbook();
			HSSFSheet sheet = wb.createSheet("New Sheet 1");
			HSSFPrintSetup ps = sheet.getPrintSetup();
			sheet.setAutobreaks(false);
			ps.setScale((short) 75); //列印縮放百分比
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
				reportUtil.createCell(wb, row, (short)startIndex, "稽核記錄統計明細表", titleStyle);
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
				for(int ci=1; ci<=totalColumnNums;ci++){
					reportUtil.createCell( wb, row, (short)ci, "", titleStyle);
				}
				sheet.addMergedRegion(new Region((short)3, (short)0, (short)3, totalColumnNums));
			}else {
				int rowIndex=3;
				//設定儲存格資料============================================
				for(int li=0;li<bankList.size(); li++){
					String[] bankData = (String[])bankList.get(li);
					String bankNoTmp = bankData[0];
					String bankNameTmp = bankData[1];
					//===先將機構名稱寫上
					row = sheet.createRow(rowIndex);
					reportUtil.createCell(wb, row, (short)0, bankNameTmp, leftStyle);
					for(int ci=1; ci<=totalColumnNums;ci++){
						reportUtil.createCell( wb, row, (short)ci, "", defaultStyle);
					}
					//==該機構下使用者列表
					List bankUserList = userListMap.get(bankNoTmp)==null?new ArrayList():(List)userListMap.get(bankNoTmp);
					for(int li2=0;li2<bankUserList.size();li2++){
						String[] userData = (String[])bankUserList.get(li2);
						rowIndex++;
						for(int mi=0;mi<monthIndex;mi++){
							if((Integer.parseInt(startMonth)+mi)>12){
								queryYear = String.valueOf(Integer.parseInt(startYear)+1);
								queryMonth=(Integer.parseInt(startMonth)+mi-12)+"";
							}else{
								queryYear = startYear;
								queryMonth = String.valueOf(Integer.parseInt(startMonth)+mi);
							}
							short cellNo = (short)(1+(mi*3));
							row = sheet.createRow(rowIndex);
							if(mi==0){//年月的第一次亦將使用者名稱填上
								//===先將使用者名稱填上
								reportUtil.createCell(wb, row, (short)0, userData[1], rightStyle);
							}
							if(resultMap.get(bankNoTmp+"_"+queryYear+queryMonth+"_"+userData[0])!=null){
								long[] dataTmp =(long[]) resultMap.get(bankNoTmp+"_"+queryYear+queryMonth+"_"+userData[0]);
								reportUtil.createCell(wb, row, (short)cellNo, String.valueOf(dataTmp[0]), rightStyle);
								reportUtil.createCell(wb, row, (short)(cellNo+1), String.valueOf(dataTmp[1]), rightStyle);
								reportUtil.createCell(wb, row, (short)(cellNo+2), String.valueOf(dataTmp[2]), rightStyle);
							}else{
								reportUtil.createCell(wb, row, (short)cellNo, "0", rightStyle);
								reportUtil.createCell(wb, row, (short)(cellNo+1), "0", rightStyle);
								reportUtil.createCell(wb, row, (short)(cellNo+2), "0", rightStyle);
							}
						}//end for every Year/Month
					}//End For every User
					//=== 寫合計
					rowIndex++;
					row = sheet.createRow(rowIndex);
					reportUtil.createCell(wb, row, (short)0, "合計", rightStyle);
					for(int mi=0;mi<monthIndex;mi++){
						if((Integer.parseInt(startMonth)+mi)>12){
							queryYear = String.valueOf(Integer.parseInt(startYear)+1);
							queryMonth=(Integer.parseInt(startMonth)+mi-12)+"";
						}else{
							queryYear = startYear;
							queryMonth = String.valueOf(Integer.parseInt(startMonth)+mi);
						}
						short cellNo = (short)(1+(mi*3));
						if(resultMap.get(bankNoTmp+"_"+queryYear+queryMonth+"_SUM")!=null){
							long[] dataTmp =(long[]) resultMap.get(bankNoTmp+"_"+queryYear+queryMonth+"_SUM");
							reportUtil.createCell(wb, row, (short)cellNo, String.valueOf(dataTmp[0]), rightStyle);
							reportUtil.createCell(wb, row, (short)(cellNo+1), String.valueOf(dataTmp[1]), rightStyle);
							reportUtil.createCell(wb, row, (short)(cellNo+2), String.valueOf(dataTmp[2]), rightStyle);
						}else{
							reportUtil.createCell(wb, row, (short)cellNo, "0", rightStyle);
							reportUtil.createCell(wb, row, (short)(cellNo+1), "0", rightStyle);
							reportUtil.createCell(wb, row, (short)(cellNo+2), "0", rightStyle);
						}
					}
					rowIndex++;
				}//end of for
			} //有資料
			File reportDir = new File(Utility.getProperties("reportDir"));
			if (!reportDir.exists()) {
				if (!Utility.mkdirs(Utility.getProperties("reportDir"))) {
					errMsg += Utility.getProperties("reportDir") + "目錄新增失敗";
				}
			}
			FileOutputStream fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + "稽核記錄統計明細表.xls");
			HSSFFooter footer = sheet.getFooter();
			footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
			wb.write(fout);
			//儲存
			fout.close();
		}catch (Exception e) {
			System.out.println("//RptCG001WB createRpt() Have Error.....");
			e.printStackTrace();
			System.out.println("//-------------------------------------");
		}finally{
			try{
				conn.close();
			}catch(Exception sqlEx){
				conn=null;
			}
		}
		if(debug) System.out.println("RptCG001WB createRpt() Debug End ...");
		return errMsg;
	}
}
