/*
 * Created on 2006/11/29 by ABYSS Allen
 * RptCG001WA 使用者帳號數量統計概況表
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

public class RptCG061W {

	public static String createRpt(String startYear,String startMonth, String endYear, String endMonth, String bankType) {
		boolean debug = false;
		String[] endDayStr = new String[]{"31","28","31","30","31","30","31","31","30","31","30","31"};
		if(debug) System.out.println("RptCG061W createRpt() Debug Start ...");
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
			HashMap lastMonthMap = new HashMap();//裡面以bank_no,裡面放上個月內登入次數
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
				if(debug) System.out.println("queryYear="+queryYear+";queryMonth="+queryMonth);
				sqlCmd="SELECT TBANK_NO, COUNT(MUSER_ID) NUMS FROM WTT01"
					+" WHERE TO_CHAR(ADD_DATE, 'yyyy/mm')  = ? ";
				if(!bankType.equals("ALL")){
					sqlCmd+=" AND BANK_TYPE=?";
				}
				sqlCmd+=" GROUP BY TBANK_NO ORDER BY TBANK_NO";
				if(debug) System.out.println("sqlCmd="+sqlCmd);
				pst = conn.prepareStatement(sqlCmd);
				String startQueryDate = "";
                                if(queryMonth.length() < 2){
                                  startQueryDate = (Integer.parseInt(queryYear) + 1911)+"/0" + queryMonth;
                                }else{
                                  startQueryDate = (Integer.parseInt(queryYear) + 1911)+"/" + queryMonth;
                                }
				String endQueryDate = (Integer.parseInt(queryYear) + 1911)+"/"+queryMonth+"/"+endDayStr[(Integer.parseInt(queryMonth))-1];
				pst.setString(1, startQueryDate);
				if(!bankType.equals("ALL")){
					pst.setString(2, bankType);
				}
				if(debug) System.out.println("startQueryDate="+startQueryDate+";endQueryDate="+endQueryDate);
				rs = pst.executeQuery();
				while(rs.next()){
					String bankNoTmp = rs.getString(1)==null?"empty":rs.getString(1);
					if(bankNoTmp.trim().length()==0)
						bankNoTmp="empty";
					int nums = rs.getInt(2);
					//if(debug) System.out.println("加到map內: key="+bankNoTmp+"_"+queryYear+queryMonth);
					resultMap.put(bankNoTmp+"_"+queryYear+queryMonth, new Integer(nums));
				}
				Calendar beforMDate = Calendar.getInstance();
				beforMDate.set(Integer.parseInt(queryYear),Integer.parseInt(queryMonth)-1,1);
				beforMDate.add(2, -1);
				sqlCmd="SELECT TBANK_NO, COUNT(MUSER_ID) NUMS FROM WTT01"
					+" WHERE TO_CHAR(ADD_DATE, 'yyyy/mm')  = ? ";
				if(!bankType.equals("ALL")){
					sqlCmd+=" AND BANK_TYPE=?";
				}
				sqlCmd+=" GROUP BY TBANK_NO ORDER BY TBANK_NO";
				if(debug) System.out.println("sqlCmd="+sqlCmd);
				pst = conn.prepareStatement(sqlCmd);
                                if((beforMDate.get(2)+1) < 10){
                                  startQueryDate = (beforMDate.get(1)+1911)+"/0" + (beforMDate.get(2)+1);
                                }else{
                                  startQueryDate = (beforMDate.get(1)+1911)+"/" + (beforMDate.get(2)+1);
                                }
				endQueryDate = (beforMDate.get(1)+1911)+"/"+(beforMDate.get(2)+1)+"/"+endDayStr[beforMDate.get(2)];
				pst.setString(1, startQueryDate);
				if(!bankType.equals("ALL")){
					pst.setString(2, bankType);
				}
				if(debug) System.out.println("startQueryDate="+startQueryDate+";endQueryDate="+endQueryDate);
				rs = pst.executeQuery();
				while(rs.next()){
					lastMonthMap.put(rs.getString(1)+"_"+queryYear+queryMonth, new Integer(rs.getInt(2)));
				}
			}
			//設定報表表頭資料============================================
			HSSFWorkbook wb = new HSSFWorkbook();
			HSSFSheet sheet = wb.createSheet("sheet1"); //讀取第一個工作表，宣告其為sheet
			HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
			//設定頁面符合列印大小
			sheet.setAutobreaks(false);//自動分頁
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
    			columnLen[ai]=(short)(5+((ai%2)*2));
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
				reportUtil.createCell(wb, row, (short)startIndex, "使用者帳號數量統計概況表", titleStyle);
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
			reportUtil.createCell(wb, row2, (short)0, "機構類別", leftStyle);
			for(int mi=0;mi<monthIndex;mi++){
				if((Integer.parseInt(startMonth)+mi)>12){
					queryYear = String.valueOf(Integer.parseInt(startYear)+1);
					queryMonth=(Integer.parseInt(startMonth)+mi-12)+"";
				}else{
					queryYear = startYear;
					queryMonth = String.valueOf(Integer.parseInt(startMonth)+mi);
				}
				reportUtil.createCell( wb, row, (short)(1+(mi*2)), queryYear+"年"+queryMonth+"月", defaultStyle);
				reportUtil.createCell( wb, row, (short)(2+(mi*2)), "", defaultStyle);
				sheet.addMergedRegion(new Region((short)1, (short)(1+(mi*2)), (short)1, (short)(2+(mi*2))));
				reportUtil.createCell( wb, row2, (short)(1+(mi*2)), "使用者數", defaultStyle);
				reportUtil.createCell( wb, row2, (short)(2+(mi*2)), "增減數", defaultStyle);
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
				sheet.addMergedRegion(new Region((short)0, (short)0, (short)0, totalColumnNums));
			}else {
				//設定儲存格資料============================================
				short rowNo=3;
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
						//=== 若超過原本的規劃的長度時,需建立row
						rowNo = (short)(li+3);
						row=sheet.createRow(rowNo);
						//=== 處理第一列時,需將機構名稱寫入
						if(mi==0){
							reportUtil.createCell(wb, row, (short)0, bankNameTmp, leftStyle);
						}
						int thisMonthNums=0,lastMonthNums=0;
						if(resultMap.get(bankNoTmp+"_"+queryYear+queryMonth)!=null){
							thisMonthNums=((Integer)resultMap.get(bankNoTmp+"_"+queryYear+queryMonth)).intValue();
						}
						if(lastMonthMap.get(bankNoTmp+"_"+queryYear+queryMonth)!=null){
							lastMonthNums=((Integer)lastMonthMap.get(bankNoTmp+"_"+queryYear+queryMonth)).intValue();
						}
						reportUtil.createCell(wb, row, (short)cellNo, String.valueOf(thisMonthNums), rightStyle);
						reportUtil.createCell(wb, row, (short)(cellNo+1), String.valueOf((thisMonthNums-lastMonthNums)), rightStyle);
					}
				}//end of for
			} //end of else ((DataObject) dbData.get(0)).getValue("m_year") == null
			File reportDir = new File(Utility.getProperties("reportDir"));
			if (!reportDir.exists()) {
				if (!Utility.mkdirs(Utility.getProperties("reportDir"))) {
					errMsg += Utility.getProperties("reportDir") + "目錄新增失敗";
				}
			}
			FileOutputStream fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + "使用者帳號數量統計概況表.xls");
			HSSFFooter footer = sheet.getFooter();
			footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
			wb.write(fout);
			//儲存
			fout.close();
		}catch (Exception e) {
			System.out.println("//RptCG061W createRpt() Have Error.....");
			e.printStackTrace();
			System.out.println("//-------------------------------------");
		}finally{
			try{
				conn.close();
			}catch(Exception sqlEx){
				conn=null;
			}
		}
		if(debug) System.out.println("RptCG061W createRpt() Debug End ...");
		return errMsg;
	}
}
