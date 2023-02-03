/*
 * Created on 2006/10/18 by ABYSS Allen
 * 歷年來全體農漁會信用部簡明存款結構比較表
 * fixed 99.06.04 sql injection by 2808
 */

package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.Region; 
import java.io.*;
import java.util.*;
import java.sql.*;
import com.tradevan.util.Utility;
import com.tradevan.util.Utility_report;
import com.tradevan.util.dao.RdbCommonDao;

public class RptAN002W {

	public static String createRpt(String startYear, String endYear, String bankType, int priceUtil) {
		String[] priceUtilStr = new String[]{"元","仟元","萬元","百萬元","仟萬元","億元"};
		boolean debug = true;
		if(debug) System.out.println("RptAN002 createRpt() Debug Start ...");
		String errMsg = "";
		List dbData = new ArrayList();
		//String sqlCmd = null;
		StringBuffer sqlCmd = new StringBuffer () ;
		List paramList = new ArrayList () ;
		String u_year = "99" ;
		if(!"".equals(startYear) && Integer.parseInt(startYear) >99) {
			u_year = "100" ;
		}
		Connection conn = null;
		PreparedStatement pst = null;
		ResultSet rs = null;
		short totalColumnNums = (short)7;
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
			conn =(new RdbCommonDao("")).newConnection();
			if(debug) System.out.println("conn="+conn);
			long projectAmount=0,pressAmount=0,sumAmount=0,preSumAmount=0;
			for(int yi=Integer.parseInt(startYear); yi<=Integer.parseInt(endYear);yi++){
				sqlCmd.append("SELECT SUM(a.AMT) FROM A01 a, (select * from bn01 where m_year=?) b WHERE a.BANK_CODE=b.BANK_NO AND a.M_YEAR=?"
					+" AND (a.ACC_CODE=? OR a.ACC_CODE=?) AND a.M_MONTH=12 AND a.BANK_CODE != ? " );
				paramList.add(u_year) ;
				paramList.add(new Integer(yi).toString()); 
				paramList.add("120501"); 
				paramList.add("120502"); 
				paramList.add("8888888"); 
				if(!bankType.equals("ALL")){
					sqlCmd.append(" AND b.BANK_TYPE=? " );
					paramList.add(bankType) ;
				}
				if(debug) System.out.println("sqlCmd="+sqlCmd);
				String[] data = new String[7];
				pst = conn.prepareStatement(sqlCmd.toString());
				setPreparedStatementParameter(pst,paramList) ; 
				
				//pst.setString(1,new Integer(yi).toString());
				//pst.setString(2,"120501");
				//pst.setString(3,"120502");
				rs = pst.executeQuery();
				if(rs.next()){
					projectAmount=rs.getLong(1);//本次專案放款
				}else{
					projectAmount=0;
				}
				rs.close();
				pst.close();
				paramList.clear() ;
				sqlCmd.setLength( 0) ;
				
				if(debug) System.out.println(yi+" projectAmount="+projectAmount);
				sqlCmd.append("SELECT SUM(a.AMT) FROM A01 a, (select * from bn01 where m_year=?) b WHERE a.BANK_CODE=b.BANK_NO AND a.M_YEAR=?"
					+" AND a.M_MONTH=12 AND a.ACC_CODE=? AND a.BANK_CODE != ? " );
				paramList.add(u_year) ;
				paramList.add(new Integer(yi).toString()) ;
				paramList.add("150200") ;
				paramList.add("8888888") ;
				if(!bankType.equals("ALL")){
					sqlCmd.append(" AND b.BANK_TYPE=? " );
					paramList.add(bankType) ;
				}
				if(debug) System.out.println("sqlCmd="+sqlCmd);
				pst = conn.prepareStatement(sqlCmd.toString() );
				//pst.setString(1,new Integer(yi).toString());
				//pst.setString(2,"150200");
				setPreparedStatementParameter(pst,paramList) ;
				rs = pst.executeQuery();
				if(rs.next()){
					pressAmount=rs.getLong(1);//本次催收款
				}else{
					pressAmount=0;
				}
				rs.close();
				pst.close();
				paramList.clear() ;
				sqlCmd.setLength( 0) ;
				
				if(debug) System.out.println(yi+" pressAmount="+pressAmount);
				sqlCmd.append("SELECT SUM(a.AMT) FROM A01 a, (select * from bn01 where m_year=? ) b WHERE a.BANK_CODE=b.BANK_NO AND a.M_YEAR=?"
					+" AND a.M_MONTH=12 AND (a.ACC_CODE=? OR a.ACC_CODE=? OR a.ACC_CODE=?) AND a.BANK_CODE != ? " );
				paramList.add(u_year) ;
				paramList.add(new Integer(yi).toString()) ;
				paramList.add("120000") ;
				paramList.add("120800") ;
				paramList.add("150300") ;
				paramList.add("8888888") ;
				if(!bankType.equals("ALL")){
					sqlCmd.append(" AND b.BANK_TYPE=? " );
					paramList.add(bankType) ;
				}
				if(debug) System.out.println("sqlCmd="+sqlCmd);
				pst = conn.prepareStatement(sqlCmd.toString());
				setPreparedStatementParameter(pst,paramList) ;
				//pst.setString(1,new Integer(yi).toString());
				//pst.setString(2,"120000");
				//pst.setString(3,"120800");
				//pst.setString(4,"150300");
				rs = pst.executeQuery();
				if(rs.next()){
					sumAmount=rs.getLong(1);//本次放款總額
				}else{
					sumAmount=0;
				}
				rs.close();
				pst.close();
				paramList.clear() ;
				sqlCmd.setLength(0) ;
				
				if(debug) System.out.println(yi+" sumAmount="+sumAmount);
				if(sumAmount>0){
					data[0]=new Integer(yi).toString();//年度
					data[1]=new Long(sumAmount-projectAmount).toString();
					data[1]=String.valueOf(Math.round(Double.parseDouble(data[1])/priceUtil));
					data[2]=new Long(projectAmount).toString();
					data[2]=String.valueOf(Math.round(Double.parseDouble(data[2])/priceUtil));
					data[3]=new Long(pressAmount).toString();
					data[3]=String.valueOf(Math.round(Double.parseDouble(data[3])/priceUtil));
					data[4]=new Long(sumAmount).toString();
					data[4]=String.valueOf(Math.round(Double.parseDouble(data[4])/priceUtil));
					data[5]=new Long(sumAmount-preSumAmount).toString();
					data[5]=String.valueOf(Math.round(Double.parseDouble(data[5])/priceUtil));
					if(preSumAmount>0){
						data[6]=Utility_report.round(Long.toString((sumAmount-preSumAmount)*100),Long.toString(preSumAmount),2);
					}else{
						data[6]="0.00";
					}
				}else{
					data[0]=new Integer(yi).toString();//年度
					data[1]="";
					data[2]="";
					data[3]="";
					data[4]="";
					data[5]="";
					data[6]="";
				}				
				dbData.add(data);
				preSumAmount=sumAmount;
			}
			if(debug) System.out.println("dbData.size=" + dbData.size());
			//設定報表表頭資料============================================
			File reportDir = new File(Utility.getProperties("reportDir"));
			if (!reportDir.exists()) {
				if (!Utility.mkdirs(Utility.getProperties("reportDir"))) {
					errMsg += Utility.getProperties("reportDir") + "目錄新增失敗";
				}
			}
			HSSFWorkbook wb = new HSSFWorkbook();
			HSSFSheet sheet = wb.createSheet("New Sheet 1");
			HSSFPrintSetup ps = sheet.getPrintSetup(); //取得列印設定
			//設定頁面符合列印大小
            sheet.setZoom(100, 100); // 螢幕上看到的縮放大小
            sheet.setAutobreaks(false); //自動分頁            
            ps.setScale((short)70); //列印縮放百分比
            ps.setPaperSize((short)9); //設定紙張大小 A4            
            //ps.setLandscape(true); // 設定橫印
            //ps.setFitWidth((short)1);//調整成幾頁寬
            //ps.setFitHeight((short)1);//調整成幾頁高
            short[] columnLen = new short[]{5,15,15,15,15,15,10};
    		for(int ai=0; ai<columnLen.length; ai++){
    			sheet.setColumnWidth((short)ai, (short)(256*(columnLen[ai]+4)));
    		}
			HSSFDataFormat format = wb.createDataFormat();
			reportUtil reportUtil = new reportUtil();
			HSSFCellStyle titleStyle = reportUtil.getTitleStyle(wb); //標題用
			HSSFFont defaultFont = wb.createFont();
			defaultFont.setFontHeightInPoints((short)12);
			HSSFCellStyle defaultStyle = reportUtil.getDefaultStyle(wb);//有框內文置中
			defaultStyle.setFont(defaultFont);
			HSSFCellStyle leftStyle = reportUtil.getLeftStyle(wb);
			leftStyle.setFont(defaultFont);
			HSSFCellStyle numberRightStyle = reportUtil.getRightStyle(wb);//有框內文置右
			numberRightStyle.setDataFormat(format.getFormat("#,##0"));
			numberRightStyle.setFont(defaultFont);
			HSSFCellStyle doubleRightStyle = reportUtil.getRightStyle(wb);//有框內文置右
			doubleRightStyle.setDataFormat(format.getFormat("#,##0.00"));
			doubleRightStyle.setFont(defaultFont);
			HSSFCellStyle noBorderRightStyle = reportUtil.getNoBoderStyle(wb);
			noBorderRightStyle.setFont(defaultFont);
			reportUtil.setDefaultStyle(defaultStyle);
			HSSFRow row = null; //宣告一列
			//設定報表表頭資料 開始============================================
			row=sheet.createRow(0);
			String titleStr="";
			if(bankType.equals("ALL")){
				titleStr = "歷年來全體農漁會信用部簡明放款結構比較表";
			}else if(bankType.equals("6")){
				titleStr = "歷年來全體農會信用部簡明放款結構比較表";
			}else{
				titleStr = "歷年來全體漁會信用部簡明放款結構比較表";
			}
			reportUtil.createCell( wb, row, (short)0, titleStr, titleStyle);
			for(int ci=1; ci<totalColumnNums;ci++){
				reportUtil.createCell( wb, row, (short)ci, "", titleStyle);
			}
			sheet.addMergedRegion(new Region((short)0, (short)0, (short)0, (short)(totalColumnNums-1)));
			row=sheet.createRow(1);
			row.setHeightInPoints(20.0F);
			reportUtil.createCell(wb, row, (short)0, "單位：新台幣 "+priceUtilStr[priceIndex], noBorderRightStyle);
			for(int ci=1; ci<totalColumnNums;ci++){
				reportUtil.createCell( wb, row, (short)ci, "", noBorderRightStyle);
			}
			sheet.addMergedRegion(new Region((short)1, (short)0, (short)1, (short)(totalColumnNums-1)));
			row=sheet.createRow(2);
			row.setHeightInPoints(24.0F);
			reportUtil.createCell(wb,row,(short)0, "年度", defaultStyle);
			reportUtil.createCell(wb,row,(short)1, "一般放款", defaultStyle);
			reportUtil.createCell(wb,row,(short)2, "專案放款", defaultStyle);
			reportUtil.createCell(wb,row,(short)3, "催收款", defaultStyle);
			reportUtil.createCell(wb,row,(short)4, "放款總額", defaultStyle);
			reportUtil.createCell(wb,row,(short)5, "增減額", defaultStyle);
			reportUtil.createCell(wb,row,(short)6, "增減率", defaultStyle);
			wb.setRepeatingRowsAndColumns(0, 0, (totalColumnNums-1), 0, 2);//設為固定表頭(第幾個sheet,起始欄,終止欄,起始列,終止列)
			//設定報表表頭資料 結束============================================
			//判斷dbData.size()是不是0，是的話表示沒有資料
			if (dbData.size() == 0) {
				int rowNo=2;
				for(int yi=Integer.parseInt(startYear); yi<=Integer.parseInt(endYear);yi++){
					row=sheet.createRow(++rowNo);
					row.setHeightInPoints(30.0F);
					reportUtil.createCell(wb,row,(short)0, yi+"", defaultStyle);
					reportUtil.createCell(wb,row,(short)1, "", leftStyle);
					for(short cellNo=2; cellNo<totalColumnNums; cellNo++){
						reportUtil.createCell(wb,row,(short)cellNo, "", leftStyle);
					}
					//sheet.addMergedRegion(new Region((short)rowNo, (short)1, (short)rowNo, (short)(totalColumnNums-1)));
				}
			}else {
				//設定儲存格資料============================================
				int rowNo=2;
				for(int li=0;li < dbData.size();li++){
					HSSFCell cell=null;
					row=sheet.createRow(++rowNo);
					row.setHeightInPoints(30.0F);
					String[] dataTmp = (String[])dbData.get(li);
					reportUtil.createCell(wb,row,(short)0, dataTmp[0], defaultStyle);
					if(dataTmp[1].trim().length()==0){
						reportUtil.createCell(wb,row,(short)1, "", leftStyle);
						reportUtil.createCell(wb,row,(short)2, dataTmp[2], leftStyle);
						reportUtil.createCell(wb,row,(short)3, dataTmp[3], leftStyle);
						reportUtil.createCell(wb,row,(short)4, dataTmp[4], leftStyle);
						reportUtil.createCell(wb,row,(short)5, dataTmp[5], leftStyle);
						reportUtil.createCell(wb,row,(short)6, dataTmp[6], leftStyle);
						//sheet.addMergedRegion(new Region((short)rowNo, (short)1, (short)rowNo, (short)(totalColumnNums-1)));
					}else{
						cell = row.createCell((short)1);
						cell.setCellStyle(numberRightStyle);
						cell.setCellValue(Long.parseLong(dataTmp[1]));
						cell = row.createCell((short)2);
						cell.setCellStyle(numberRightStyle);
						cell.setCellValue(Long.parseLong(dataTmp[2]));
						cell = row.createCell((short)3);
						cell.setCellStyle(numberRightStyle);
						cell.setCellValue(Long.parseLong(dataTmp[3]));
						cell = row.createCell((short)4);
						cell.setCellStyle(numberRightStyle);
						cell.setCellValue(Long.parseLong(dataTmp[4]));
						cell = row.createCell((short)5);
						cell.setCellStyle(numberRightStyle);
						cell.setCellValue(Long.parseLong(dataTmp[5]));
					}					
					reportUtil.createCell(wb,row,(short)6, dataTmp[6], doubleRightStyle);
				}//end of for
			} //end of else ((DataObject) dbData.get(0)).getValue("m_year") == null
			FileOutputStream fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + titleStr+".xls");
			HSSFFooter footer = sheet.getFooter();
			footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
			wb.write(fout);
			//儲存
			fout.close();
		}catch (Exception e) {
			System.out.println("//RptAN002 createRpt() Have Error.....");
			e.printStackTrace();
			System.out.println("//-------------------------------------");
		}finally{
			try{
				if(rs != null){
	               rs.close();
	               rs = null;//104.10.06
	            }
	            if(pst != null){
	               pst.close();
	               pst = null;//104.10.06
	            }
	            if(!conn.isClosed()){//104.10.06    
	               conn.close();
	               conn = null;
	            }
			}catch(Exception sqlEx){
				conn=null;
			}
		}
		if(debug) System.out.println("RptAN002 createRpt() Debug Start ...");
		return errMsg;
	}
	
	private static void setPreparedStatementParameter(PreparedStatement pst,List paramList) throws Exception{
		for(int i = 0 ;i< paramList.size() ;i++) {
			pst.setString(i+1,(String)paramList.get(i)) ;
		}
	}
}
