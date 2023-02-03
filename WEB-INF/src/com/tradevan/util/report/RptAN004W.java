/*
 * Created on 2006/10/24 by ABYSS Allen
 * 多個年度農漁會信用部存款結構及變動表
 * fixed 99.06.04 sql injection by 2808
 */

package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.poifs.filesystem.*;
import java.io.*;
import java.sql.*;
import java.util.ArrayList;
import java.util.List;

import com.tradevan.util.Utility;
import com.tradevan.util.Utility_report;
import com.tradevan.util.dao.RdbCommonDao;

public class RptAN004W {

	public static String createRpt(String startYear, String endYear, String bankType, int priceUtil) {
		String[] priceUtilStr = new String[]{"元","仟元","萬元","百萬元","仟萬元","億元"};
		boolean debug = false;
		if(debug) System.out.println("RptAN004 createRpt() Debug Start ...");
		String errMsg = "";
		StringBuffer sqlCmd = new StringBuffer () ;
		List paramList = new ArrayList () ;
		String u_year = "99" ;
		if(!"".equals(startYear) && Integer.parseInt(startYear) >99) {
			u_year = "100" ;
		}
		Connection conn = null;
		PreparedStatement pst = null;
		ResultSet rs = null;
		try {
			conn =(new RdbCommonDao("")).newConnection();
			int yearIndex = Integer.parseInt(endYear)-Integer.parseInt(startYear)+1;
			if(debug) System.out.println("yearIndex="+yearIndex);
			String[][] dataValue = new String[yearIndex*2][14];
			int arrayIndex=0;
			short totalColumnNums = (short)(yearIndex*2);
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
			if(debug) System.out.println("priceUtil="+priceUtil+";priceIndex="+priceIndex+";priceUtilStr="+priceUtilStr[priceIndex]);
			for(int yi=Integer.parseInt(startYear); yi<=Integer.parseInt(endYear);yi++){
				long a=0, b=0, c=0, d=0, e=0, f=0,g=0,h=0, x=0, y=0, z=0, tempTotalNums=0;
				sqlCmd.setLength(0) ;
				sqlCmd.append("SELECT SUM(a.AMT), a.ACC_CODE FROM A01 a, (select * from bn01 where m_year=?) b WHERE a.BANK_CODE=b.BANK_NO AND a.M_YEAR=? AND a.M_MONTH=12"
					+" AND a.ACC_CODE IN (?,?,?,?,?,?,?,?)"
					+" AND a.BANK_CODE != ?" );
				paramList.add(u_year) ;
				paramList.add(new Integer(yi).toString()) ;
				paramList.add("220100") ;
				paramList.add("220300") ;
				paramList.add("220400") ;
				paramList.add("220500") ;
				paramList.add("220600") ;
				paramList.add("220700") ;
				paramList.add("220800") ;
				paramList.add("220900") ;
				paramList.add("8888888") ;
				if(!bankType.equals("ALL")){
					sqlCmd.append(" AND b.BANK_TYPE= ? " );
					paramList.add(bankType) ;
				}
				sqlCmd.append(" GROUP BY a.M_YEAR, a.ACC_CODE" );
				if(debug) System.out.println("sqlCmd="+sqlCmd);
				pst = conn.prepareStatement(sqlCmd.toString());
				//pst.setString(1,new Integer(yi).toString());
				setPreparedStatementParameter(pst,paramList) ;
				rs = pst.executeQuery();
				while(rs.next()){
					long valueTmp = rs.getLong(1);
					String accCodeTmp = rs.getString(2);
					if(accCodeTmp.equals("220100")){
						a=valueTmp;
						if(debug) System.out.println(yi+";支票存款="+a);
					}else if(accCodeTmp.equals("220300")){
						b=valueTmp;
						if(debug) System.out.println(yi+";活期存款="+b);
					}else if(accCodeTmp.equals("220400")){
						c=valueTmp;
						if(debug) System.out.println(yi+";活期儲蓄存款="+c);
					}else if(accCodeTmp.equals("220500")){
						d=valueTmp;
						if(debug) System.out.println(yi+";員工活期儲蓄存款="+d);
					}else if(accCodeTmp.equals("220600")){
						e=valueTmp;
						if(debug) System.out.println(yi+";定期存款="+e);
					}else if(accCodeTmp.equals("220700")){
						f=valueTmp;
						if(debug) System.out.println(yi+";定期儲蓄存款="+f);
					}else if(accCodeTmp.equals("220800")){
						g=valueTmp;
						if(debug) System.out.println(yi+";員工儲蓄存款="+g);
					}else if(accCodeTmp.equals("220900")){
						h=valueTmp;
						if(debug) System.out.println(yi+";公庫存款="+h);
					}
				}
				sqlCmd.setLength(0);
				paramList.clear() ;
				x=a+b+c+d;
				if(debug) System.out.println(yi+";x="+x);
				y=e+f+g;
				if(debug) System.out.println(yi+";y="+y);
				z=x+y+h;
				if(debug) System.out.println(yi+";z="+z);
				dataValue[arrayIndex*2][0]=yi+"年度";
				dataValue[(arrayIndex*2)+1][0]=yi+"年度";
				dataValue[arrayIndex*2][1]="金額";
				dataValue[(arrayIndex*2)+1][1]="％";
				if(a>0){
					dataValue[arrayIndex*2][2]=Utility.setCommaFormat(Utility_report.round(a+"",priceUtil+"",0));
					dataValue[(arrayIndex*2)+1][2]=Utility.setCommaFormat(Utility_report.round((a*100)+"",z+"",2));
				}else{
					dataValue[arrayIndex*2][2]="0";
					dataValue[(arrayIndex*2)+1][2]="0.00";
				}
				if(b>0){
					dataValue[arrayIndex*2][3]=Utility.setCommaFormat(Utility_report.round(b+"",priceUtil+"",0));
					dataValue[(arrayIndex*2)+1][3]=Utility.setCommaFormat(Utility_report.round((b*100)+"",z+"",2));
				}else{
					dataValue[arrayIndex*2][3]="0";
					dataValue[(arrayIndex*2)+1][3]="0.00";
				}
				if(c>0){
					dataValue[arrayIndex*2][4]=Utility.setCommaFormat(Utility_report.round(c+"",priceUtil+"",0));
					dataValue[(arrayIndex*2)+1][4]=Utility.setCommaFormat(Utility_report.round((c*100)+"",z+"",2));
				}else{
					dataValue[arrayIndex*2][4]="0";
					dataValue[(arrayIndex*2)+1][4]="0.00";
				}
				if(d>0){
					dataValue[arrayIndex*2][5]=Utility.setCommaFormat(Utility_report.round(d+"",priceUtil+"",0));
					dataValue[(arrayIndex*2)+1][5]=Utility.setCommaFormat(Utility_report.round((d*100)+"",z+"",2));
				}else{
					dataValue[arrayIndex*2][5]="0";
					dataValue[(arrayIndex*2)+1][5]="0.00";
				}
				if(x>0){
					dataValue[arrayIndex*2][6]=Utility.setCommaFormat(Utility_report.round(x+"",priceUtil+"",0));
					dataValue[(arrayIndex*2)+1][6]=Utility.setCommaFormat(Utility_report.round((x*100)+"",z+"",2));
				}else{
					dataValue[arrayIndex*2][6]="0";
					dataValue[(arrayIndex*2)+1][6]="0.00";
				}
				if(e>0){
					dataValue[arrayIndex*2][7]=Utility.setCommaFormat(Utility_report.round(e+"",priceUtil+"",0));
					dataValue[(arrayIndex*2)+1][7]=Utility.setCommaFormat(Utility_report.round((e*100)+"",z+"",2));
				}else{
					dataValue[arrayIndex*2][7]="0";
					dataValue[(arrayIndex*2)+1][7]="0.00";
				}
				if(f>0){
					dataValue[arrayIndex*2][8]=Utility.setCommaFormat(Utility_report.round(f+"",priceUtil+"",0));
					dataValue[(arrayIndex*2)+1][8]=Utility.setCommaFormat(Utility_report.round((f*100)+"",z+"",2));
				}else{
					dataValue[arrayIndex*2][8]="0";
					dataValue[(arrayIndex*2)+1][8]="0.00";
				}
				if(g>0){
					dataValue[arrayIndex*2][9]=Utility.setCommaFormat(Utility_report.round(g+"",priceUtil+"",0));
					dataValue[(arrayIndex*2)+1][9]=Utility.setCommaFormat(Utility_report.round((g*100)+"",z+"",2));
				}else{
					dataValue[arrayIndex*2][9]="0";
					dataValue[(arrayIndex*2)+1][9]="0.00";
				}
				if(y>0){
					dataValue[arrayIndex*2][10]=Utility.setCommaFormat(Utility_report.round(y+"",priceUtil+"",0));
					dataValue[(arrayIndex*2)+1][10]=Utility.setCommaFormat(Utility_report.round((y*100)+"",z+"",2));
				}else{
					dataValue[arrayIndex*2][10]="0";
					dataValue[(arrayIndex*2)+1][10]="0.00";
				}
				if((x+y)>0){
					dataValue[arrayIndex*2][11]=Utility.setCommaFormat(Utility_report.round((x+y)+"",priceUtil+"",0));
					dataValue[(arrayIndex*2)+1][11]=Utility.setCommaFormat(Utility_report.round(((x+y)*100)+"",z+"",2));
				}else{
					dataValue[arrayIndex*2][11]="0";
					dataValue[(arrayIndex*2)+1][11]="0.00";
				}
				if(h>0){
					dataValue[arrayIndex*2][12]=Utility.setCommaFormat(Utility_report.round(h+"",priceUtil+"",0));
					dataValue[(arrayIndex*2)+1][12]=Utility.setCommaFormat(Utility_report.round((h*100)+"",z+"",2));
				}else{
					dataValue[arrayIndex*2][12]="0";
					dataValue[(arrayIndex*2)+1][12]="0.00";
				}
				dataValue[arrayIndex*2][13]=Utility.setCommaFormat(Utility_report.round(z+"",priceUtil+"",0));
				dataValue[(arrayIndex*2)+1][13]="100.00";
				tempTotalNums=a+b+c+d+e+f+g+h+x+y+z;
				if(tempTotalNums==0){
					for(int di=2;di<=13;di++){
						if(di==2){
							dataValue[(arrayIndex*2)][0]+=" 資料不完整";
						}
						dataValue[(arrayIndex*2)][di]="";
						dataValue[(arrayIndex*2)+1][di]="";
					}
				}
				arrayIndex++;
			}
			if(debug){
				for(int x=0; x<dataValue.length; x++){
					for(int y=0;y<dataValue[x].length;y++){
						System.out.println("dataValue["+x+"]["+y+"]="+dataValue[x][y]);
					}
				}
			}
			HSSFWorkbook wb = new HSSFWorkbook();
			HSSFSheet sheet = wb.createSheet("New Sheet 1");
			HSSFPrintSetup ps = sheet.getPrintSetup(); //取得列印設定
			//設定頁面符合列印大小
            sheet.setZoom(70, 100); // 螢幕上看到的縮放大小
            sheet.setAutobreaks(false); //自動分頁            
            ps.setScale((short)75); //列印縮放百分比
            ps.setPaperSize((short)9); //設定紙張大小 A4
            ps.setLandscape(true); // 設定橫印
			HSSFRow row = null; //宣告一列
			short[] columnLen = new short[totalColumnNums+3];
    		columnLen[0]=1;
    		columnLen[1]=1;
    		columnLen[2]=18;
    		for(int ai=3; ai<columnLen.length; ai+=2){
    			columnLen[ai]=15;
    			columnLen[ai+1]=5;
    		}
    		for(int ai=0; ai<columnLen.length; ai++){
    			sheet.setColumnWidth((short)ai, (short)(256*(columnLen[ai]+4)));
    		}
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
			if(debug) System.out.println("totalColumnNums="+totalColumnNums);
			//設定報表表頭資料 開始============================================
			row=sheet.createRow(0);
			String titleStr="";
			if(bankType.equals("ALL")){
				titleStr = "多個年度農漁會信用部存款結構及變動表";
			}else if(bankType.equals("6")){
				titleStr = "多個年度農會信用部存款結構及變動表";
			}else{
				titleStr = "多個年度漁會信用部存款結構及變動表";
			}
			int repeatTime = Math.abs(totalColumnNums/10);
			if(totalColumnNums>10 && (totalColumnNums%10)>0){
				repeatTime++;
			}
			if(debug) System.out.println("repeatTime="+repeatTime);
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
				if(debug) System.out.println(ri+":startIndex="+startIndex+";endIndex="+endIndex);
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
			row.setHeightInPoints(36.0F);
			reportUtil.createCell(wb, row, (short)0, "存款項目", defaultStyle);
			reportUtil.createCell(wb, row, (short)1, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)2, "", defaultStyle);
			sheet.addMergedRegion(new Region((short)4, (short)0, (short)4, (short)2));
			row=sheet.createRow(5);
			row.setHeightInPoints(40.0F);
			reportUtil.createCell(wb, row, (short)0, "一般存款", defaultStyle);
			reportUtil.createCell(wb, row, (short)1, "活期性", defaultStyle);
			reportUtil.createCell(wb, row, (short)2, "支票存款", defaultStyle);
			row=sheet.createRow(6);
			row.setHeightInPoints(40.0F);
			reportUtil.createCell(wb, row, (short)0, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)1, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)2, "活期存款", defaultStyle);
			row=sheet.createRow(7);
			row.setHeightInPoints(40.0F);
			reportUtil.createCell(wb, row, (short)0, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)1, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)2, "活期儲蓄存款", defaultStyle);
			row=sheet.createRow(8);
			row.setHeightInPoints(40.0F);
			reportUtil.createCell(wb, row, (short)0, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)1, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)2, "員工活期儲蓄存款", defaultStyle);
			row=sheet.createRow(9);
			row.setHeightInPoints(40.0F);
			reportUtil.createCell(wb, row, (short)0, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)1, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)2, "小計", defaultStyle);
			sheet.addMergedRegion(new Region((short)5, (short)1, (short)9, (short)1));
			row=sheet.createRow(10);
			row.setHeightInPoints(40.0F);
			reportUtil.createCell(wb, row, (short)0, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)1, "定期性", defaultStyle);
			reportUtil.createCell(wb, row, (short)2, "定期存款", defaultStyle);
			row=sheet.createRow(11);
			row.setHeightInPoints(40.0F);
			reportUtil.createCell(wb, row, (short)0, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)1, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)2, "定期儲蓄存款", defaultStyle);
			row=sheet.createRow(12);
			row.setHeightInPoints(40.0F);
			reportUtil.createCell(wb, row, (short)0, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)1, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)2, "員工儲蓄存款", defaultStyle);
			row=sheet.createRow(13);
			row.setHeightInPoints(40.0F);
			reportUtil.createCell(wb, row, (short)0, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)1, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)2, "小計", defaultStyle);
			sheet.addMergedRegion(new Region((short)10, (short)1, (short)13, (short)1));
			row=sheet.createRow(14);
			row.setHeightInPoints(40.0F);
			reportUtil.createCell(wb, row, (short)0, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)1, "小計", defaultStyle);
			reportUtil.createCell(wb, row, (short)2, "", defaultStyle);
			sheet.addMergedRegion(new Region((short)14, (short)1, (short)14, (short)2));
			sheet.addMergedRegion(new Region((short)5, (short)0, (short)14, (short)0));
			row=sheet.createRow(15);
			row.setHeightInPoints(40.0F);
			reportUtil.createCell(wb, row, (short)0, "公庫存款", defaultStyle);
			reportUtil.createCell(wb, row, (short)1, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)2, "", defaultStyle);
			sheet.addMergedRegion(new Region((short)15, (short)0, (short)15, (short)2));
			row=sheet.createRow(16);
			row.setHeightInPoints(40.0F);
			reportUtil.createCell(wb, row, (short)0, "合計", defaultStyle);
			reportUtil.createCell(wb, row, (short)1, "", defaultStyle);
			reportUtil.createCell(wb, row, (short)2, "", defaultStyle);
			sheet.addMergedRegion(new Region((short)16, (short)0, (short)16, (short)2));
			wb.setRepeatingRowsAndColumns(0, 0, 2, 2, 16);//設為固定表頭(第幾個sheet,起始欄,終止欄,起始列,終止列)
			//設定報表表頭資料 結束============================================
			//設定儲存格資料============================================
			int rowNo=1;
			int cellNo=0;
			for(int ai=0; ai<dataValue.length; ai+=2){
				row=sheet.getRow(rowNo++);
				cellNo=ai+3;
				for(int ai2=1; ai2<dataValue[ai].length; ai2++){
					rowNo=(ai2+3);
					row=sheet.getRow(rowNo);
					if(ai2==1){
						HSSFRow rowTmp=sheet.getRow(2);
						reportUtil.createCell(wb, rowTmp, (short)cellNo, dataValue[ai][0], defaultStyle);//設年度
						reportUtil.createCell(wb, rowTmp, (short)(cellNo+1), "", defaultStyle);
						rowTmp=sheet.getRow(3);
						reportUtil.createCell(wb, rowTmp, (short)cellNo, "", defaultStyle);//設年度
						reportUtil.createCell(wb, rowTmp, (short)(cellNo+1), "", defaultStyle);
						sheet.addMergedRegion(new Region((short)2, (short)cellNo, (short)3, (short)(cellNo+1)));
						reportUtil.createCell(wb, rowTmp, (short)cellNo, dataValue[ai][ai2], defaultStyle);//設金額
						reportUtil.createCell(wb, rowTmp, (short)(cellNo+1), dataValue[ai+1][ai2], defaultStyle);//比例
					}
					if(dataValue[ai][ai2].trim().length()>0){
						reportUtil.createCell(wb, row, (short)cellNo, dataValue[ai][ai2], numberRightStyle);//設金額
					}else{
						reportUtil.createCell(wb, row, (short)cellNo, "", rightStyle);//設金額
					}
					if(dataValue[ai+1][ai2].trim().length()>0){
						reportUtil.createCell(wb, row, (short)(cellNo+1), dataValue[ai+1][ai2], doubleRightStyle);//比例
					}else{
						reportUtil.createCell(wb, row, (short)(cellNo+1), "", rightStyle);//比例
					}
				}
			}//end of for
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
			System.out.println("//RptAN004 createRpt() Have Error.....");
			//e.printStackTrace();
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
		if(debug) System.out.println("RptAN004 createRpt() Debug End ...");
		return errMsg;
	}
	private static void setPreparedStatementParameter(PreparedStatement pst,List paramList) throws Exception{
		for(int i = 0 ;i< paramList.size() ;i++) {
			pst.setString(i+1,(String)paramList.get(i)) ;
		}
	}
}
