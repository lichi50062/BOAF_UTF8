/*
 * Created on 2005/1/11
 *
 * TODO To change the template for this generated file go to
 * Window - Preferences - Java - Code Style - Code Templates
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
import com.tradevan.util.report.reportUtil;
/**
 * @author 2295
 *
 * TODO To change the template for this generated type comment go to
 * Window - Preferences - Java - Code Style - Code Templates
 */
public class RptTC36 {
    public static String createRpt(String reportno,String s_year){
		String errMsg = "";
		List dbData = null;
		String sqlCmd = "";
		List paramList = new ArrayList();
		reportUtil reportUtil = new reportUtil();
		String u_year = (Integer.parseInt(s_year) < 100)?"99":"100"; 
		System.out.println("Class RptTC36.java createRpt Start...");
		System.out.println("input reportno="+reportno);
		try{
			System.out.println("金融業務檢查缺失改善情形報告表.xls");
			File xlsDir = new File(Utility.getProperties("xlsDir"));
        	File reportDir = new File(Utility.getProperties("reportDir"));
			
    		if(!xlsDir.exists()){
     			if(!Utility.mkdirs(Utility.getProperties("xlsDir"))){
     		   		errMsg +=Utility.getProperties("xlsDir")+"目錄新增失敗";
     			}
    		}
    		if(!reportDir.exists()){
     			if(!Utility.mkdirs(Utility.getProperties("reportDir"))){
     		   		errMsg +=Utility.getProperties("reportDir")+"目錄新增失敗";
     			}
    		}
    		
			FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+"金融業務檢查缺失改善情形報告表.xls" );
			
			System.out.println("xlsDir="+Utility.getProperties("xlsDir"));
			System.out.println("reportDir="+Utility.getProperties("reportDir"));
			System.out.println("Open excel 完成");

	  	    //設定FileINputStream讀取Excel檔
	  		POIFSFileSystem fs = new POIFSFileSystem( finput );
	  		if(fs==null){System.out.println("open 範本檔失敗");} else System.out.println("open 範本檔成功");
	  		HSSFWorkbook wb = new HSSFWorkbook(fs);
	  		if(wb==null){System.out.println("open工作表失敗");}else System.out.println("open 工作表 成功");
	  		HSSFSheet sheet = wb.getSheetAt(0);//讀取第一個工作表，宣告其為sheet
	  		if(sheet==null){System.out.println("open sheet 失敗");}else System.out.println("open sheet 成功");
	  		HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
	        //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
	        //sheet.setAutobreaks(true); //自動分頁

	        //設定頁面符合列印大小
	        sheet.setAutobreaks( false );
	        ps.setScale( ( short )76 ); //列印縮放百分比
	  		ps.setLandscape( true ); // 設定橫式
	        ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
	  		//wb.setSheetName(0,"test");
	  		finput.close();

	  		HSSFRow row=null;//宣告一列
	  		HSSFCell cell=null;//宣告一個儲存格

	  		short i=0;
	  		short y=0;
	  		sqlCmd 	= " SELECT A.ORIGINUNT_ID,C.RT_DOCNO,E.CMUSE_NAME,A.BANK_NO,D.BANK_NAME,B.ITEM_NO,B.EX_CONTENT,B.COMMENTT,C.DIGEST,B.AUDIT_RESULT,E.CMUSE_NAME as CDSHARENO_1 "
					+ " FROM EXREPORTF A,EXDEFGOODF B, EXDG_HISTORYF C, (select * from BA01 where m_year= ? )D, CDSHARENO E, CDSHARENO F "
					+ " WHERE A.REPORTNO = B.REPORTNO "
					+ " AND   ( B.REPORTNO = C.REPORTNO(+) AND  B.REPORTNO_SEQ = C.REPORTNO_SEQ(+)) "
					+ " AND	 A.BANK_NO	= D.BANK_NO "
					+ " AND  B.REPORTNO_SEQ = C.REPORTNO_SEQ "
					+ " AND (A.ORIGINUNT_ID = E.CMUSE_ID AND E.CMUSE_DIV = '024') "
					+ " AND ((B.AUDIT_RESULT = F.CMUSE_ID AND F.CMUSE_DIV = '026') OR B.AUDIT_RESULT IS NULL) "
					+ " AND	 A.REPORTNO = ? "
					+ " ORDER BY A.ORIGINUNT_ID,A.BANK_NO ";
			paramList.add(u_year);
			paramList.add(reportno);
			//System.out.println("sql="+sqlCmd);
			dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
			
			int j=0;
			short top=0,down=0;
			System.out.println("dbData.size="+dbData.size());
  			for(i=0;i<dbData.size();i++){
  				j=0;
  				
				top  = (i==0)?HSSFCellStyle.BORDER_MEDIUM:HSSFCellStyle.BORDER_NONE;
  				down = (i==dbData.size()-1)?HSSFCellStyle.BORDER_MEDIUM:HSSFCellStyle.BORDER_NONE;
				
				if(i==0){
					//System.out.println("print report header");
					//System.out.println("cmuse_name="+((DataObject)dbData.get(i)).getValue("cmuse_name"));
					//System.out.println("bank_name="+((DataObject)dbData.get(i)).getValue("bank_name"));
					j=0;
					//第二列表頭資料
					String chkUnit= "檢查單位: "+ ((DataObject)dbData.get(i)).getValue("cmuse_name");
					System.out.println("chkUnit="+chkUnit);
					row=(sheet.getRow(i+2)==null)? sheet.createRow(i+2) : sheet.getRow(i+2);
					insertCell(dbData,false,i,chkUnit,wb,row,(short)j, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);
					//System.out.println("第2列表頭資料i="+(i+2)+",j="+j+"cmuse_name="+((DataObject)dbData.get(i)).getValue("cmuse_name"));
					//第三列表頭資料
					String chkBank = "受檢單位代號: "+ ((DataObject)dbData.get(i)).getValue("bank_no") + " " +((DataObject)dbData.get(i)).getValue("bank_name");
					System.out.println("chkBank="+chkBank);
					row=(sheet.getRow(i+3)==null)? sheet.createRow(i+3) : sheet.getRow(i+3);
  					insertCell(dbData,false,i,chkBank,wb,row,(short)j, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);
  					j++;
					//System.out.println("第3列表頭資料i="+(i+3)+",j="+j+"bank_no="+((DataObject)dbData.get(i)).getValue("bank_no"));
  					//insertCell(dbData,false,i,"bank_name",wb,row,(short)j, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);
  					//System.out.println("第3列表頭資料i="+(i+3)+",j="+j+"bank_name="+((DataObject)dbData.get(i)).getValue("bank_name"));
  					//取得目前日期資料
  					Calendar cal = Calendar.getInstance();
  					String Year=(String.valueOf(cal.get(Calendar.YEAR)-1911));
  					String Month=(cal.get(Calendar.MONTH)+1) < 10 ? 0+(String.valueOf(cal.get(Calendar.MONTH)+1)):(String.valueOf(cal.get(Calendar.MONTH)+1));
  					String Date=(cal.get(Calendar.DATE)+1) < 10 ? 0+(String.valueOf(cal.get(Calendar.DATE)+1)):(String.valueOf(cal.get(Calendar.DATE)+1));
  					String cur_Date="列印日期: "+Year+"/"+Month+"/"+Date;
  					System.out.println("Month="+Month);
  					System.out.println("Date="+Date);
  					System.out.println("cur_Date="+cur_Date);
					row=(sheet.getRow(i+3)==null)? sheet.createRow(i+3) : sheet.getRow(i+3);
  					insertCell(dbData,false,i,cur_Date,wb,row,(short)(j+3), (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);j++;
  					//System.out.println("第3列表頭資料i="+(i+3)+",j="+j+"cur_Date="+((DataObject)dbData.get(i)).getValue("cur_Date"));
				}
				
				j=0;
				System.out.println("print report body");
				row=(sheet.getRow(i+06)==null)? sheet.createRow(i+06) : sheet.getRow(i+06);
				row.setHeightInPoints((float)16.5);
  				insertCell(dbData,true,i,"ex_content",wb,row,(short)j, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);j++;
  				insertCell(dbData,true,i,"rt_docno",wb,row,(short)j, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);j++;
  				insertCell(dbData,true,i,"commentt",wb,row,(short)j, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);j++;
  				insertCell(dbData,true,i,"digest",wb,row,(short)j, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);j++;
  				insertCell(dbData,true,i,"",wb,row,(short)j, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);j++;
  				insertCell(dbData,true,i,"cdshareno_1",wb,row,(short)j, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);j++;
  			}
			//sheet.addMergedRegion( new Region( ( short )(i+06), ( short )0,
            //                                   ( short )(i+06),
            //                                   ( short )16));
	       	//產生Excel檔案處理
	       	System.out.println(reportDir + System.getProperty("file.separator")+"金融業務檢查缺失改善情形報告表.xls");	        
	        FileOutputStream fout=new FileOutputStream(reportDir + System.getProperty("file.separator")+"金融業務檢查缺失改善情形報告表.xls");
	        wb.write(fout);
	        //儲存
	        fout.close();
	        System.out.println("儲存完成");
		}catch(Exception e){
			System.out.println("createRpt Error:"+e+e.getMessage());
		}
		System.out.println("Class RptTC36.java createRpt End...");
		return errMsg;
	}

	public static void insertCell(List dbData,boolean getstate,int index,String Item,HSSFWorkbook wb,HSSFRow row,short j,
	                              short bg,short bordertop,short borderbutton,short borderleft,short borderright){
		String insertValue="";
		System.out.println("RptTC36_Method:insertCell Start..");
		System.out.println("getstat="+getstate);
		System.out.println("index="+index);
		System.out.println("item="+Item);
		//System.out.println("wb="+wb);
		//System.out.println("row="+row);
		//System.out.println("current j="+j);
		//System.out.println("bg="+bg);
		//System.out.println("bordertop="+bordertop);
		//System.out.println("borderbutton="+borderbutton);
		//System.out.println("borderleft="+borderleft);
		//System.out.println("borderright="+borderright);
  		if(getstate) {
  			if(Item.equals("")){
  				insertValue="";
  			}else{
  				insertValue=((DataObject)dbData.get(index)).getValue(Item)== null ? "" :(String)((DataObject)dbData.get(index)).getValue(Item);
  			}
  			System.out.println("getstat = true and insertValue="+insertValue);   
  		}else{
  		    insertValue= Item;
  			System.out.println("getstat = false and insertValue="+insertValue);   
  		}
		
	    HSSFCell cell=(row.getCell(j)==null)? row.createCell(j) : row.getCell(j);
	    HSSFCellStyle cs1 = wb.createCellStyle();
	    
	    //設定邊框線條
	    if(getstate) {
	    	//System.out.println("設定邊框線條");
	    	HSSFCellStyle cs2 = cell.getCellStyle();
	    	//System.out.println("cs2.getBorderTop()="+cs2.getBorderTop());
	    	//System.out.println("cs2.getBorderBottom()="+cs2.getBorderBottom());
	    	//System.out.println("cs2.getWrapText()="+cs2.getWrapText());
	    	//cs1.setBorderTop(bordertop);
	    	cs1.setBorderTop((short)1);			//BORDER_THIN
	    	cs1.setBorderBottom((short)1);		//BORDER_THIN
	    	cs1.setBorderLeft((short)1);		//BORDER_THIN
	    	cs1.setBorderRight((short)1);		//BORDER_THIN
	    	cs1.setFillPattern((short)0);		//NO_FILL
	    	cs1.setAlignment((short)1);			//設定Cell 水平置中
			cs1.setVerticalAlignment((short)3);	//設定Cell 垂直靠上
			//cs1.setFillPattern((short)15);
			//cs1.setWrapText(true);
	    	//cs1.setBorderBottom(borderbutton);
			//cell.setCellStyle(cs1);
	    	//cs1.setFillForegroundColor(bg);
	    }
	    
	    HSSFFont f = wb.createFont();
	    f.setFontHeightInPoints((short)13.5);
	    //if(j==0) f.setFontName("標楷體");
	    f.setFontName("標楷體");
	    cs1.setFont(f);
	    System.out.println("Value="+insertValue);
	    System.out.println("字體為="+f.getFontName());
	    System.out.println("getIndention()="+cs1.getIndention());
	    System.out.println("getWrapText()="+cs1.getWrapText());
	    cell.setCellStyle(cs1);
	    //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
	    cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
	    double value=0;
	    try{
	    	//cs1.setDataFormat((short)3);	// "#,##0"
	    	cell.setCellValue(Double.parseDouble(insertValue));
	    }catch(NumberFormatException e){
	    	System.out.println("NumberFormatException...");
	    	cell.setCellValue(insertValue);
	    }
	    System.out.println("RptTC36_Method:insertCell End..");
	}
}
