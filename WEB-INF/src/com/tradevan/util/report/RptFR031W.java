//2013.11.19 add 原QueryDB改套用QueryDB_SQLParam by 2295
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

public class RptFR031W {
    public static String createRpt(String s_year,String s_month,String e_year,String e_month,String datestate){
		String errMsg = "";
		List dbData = null;
		String sqlCmd = "";
		Properties A01Data = new Properties();		
		String acc_code = "";
		String amt = "";
		reportUtil reportUtil = new reportUtil();
		try{	
			System.out.println("農業信用保證基金業務統計(一).xls");
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
			FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+"農業信用保證基金業務統計(一).xls" );
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

	        ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
	  		//wb.setSheetName(0,"test");
	  		finput.close();
	  		
	  		HSSFRow row=null;//宣告一列 
	  		HSSFCell cell=null;//宣告一個儲存格  		
	  		
	  		short i=0;
	  		short y=0;	  
	  		List paramList = new ArrayList();
			sqlCmd = "select * "
			       + "from ( "
			       + "	(	select m_year, "
			       + "		       m_month, "
			       + "			   sum(decode(guarantee_item_no,'0',guarantee_cnt,0)) \"type0_cnt\", "	//總計 件數 
			       + "			   round(sum(decode(guarantee_item_no,'0',guarantee_amt,0))/1000) \"type0_amt\", "	//總計 保證金額 
			       + "			   sum(decode(guarantee_item_no,'1',guarantee_cnt,0)) \"type1_cnt\", "	//加速農村建設貸款件數 
			       + "			   round(sum(decode(guarantee_item_no,'1',guarantee_amt,0))/1000) \"type1_amt\", "	//加速農村建設貸款保證金額 
			       + "			   sum(decode(guarantee_item_no,'2',guarantee_cnt,0)) \"type2_cnt\", "	//農業發展基金農機貸款件數 
			       + "			   round(sum(decode(guarantee_item_no,'2',guarantee_amt,0))/1000) \"type2_amt\", "	//農業發展基金農機貸款保證金額 
			       + "			   sum(decode(guarantee_item_no,'3',guarantee_cnt,0)) \"type3_cnt\", "	//農業專案貸款件數 
			       + "			   round(sum(decode(guarantee_item_no,'3',guarantee_amt,0))/1000) \"type3_amt\", "	//農業專案貸款保證金額 
			       + "			   sum(decode(guarantee_item_no,'4',guarantee_cnt,0)) \"type4_cnt\", "	// 統  一 農  貸 件數 
			       + "			   round(sum(decode(guarantee_item_no,'4',guarantee_amt,0))/1000) \"type4_amt\", "	// 統  一 農  貸 保證金額 
			       + "			   sum(decode(guarantee_item_no,'5',guarantee_cnt,0)) \"type5_cnt\", "	// 輔 導 漁  貸 件數 
			       + "			   round(sum(decode(guarantee_item_no,'5',guarantee_amt,0))/1000) \"type5_amt\", "	// 輔 導 漁  貸 保證金額 
			       + "			   sum(decode(guarantee_item_no,'6',guarantee_cnt,0)) \"type6_cnt\", "	//農漁會會員一般貸款 件數 
			       + "			   round(sum(decode(guarantee_item_no,'6',guarantee_amt,0))/1000) \"type6_amt\", "	//農漁會會員一般貸款 保證金額 
			       + "			   sum(decode(guarantee_item_no,'8',guarantee_cnt,0)) \"type8_cnt\", "	//其他政策性貸款 件數 
			       + "			   round(sum(decode(guarantee_item_no,'8',guarantee_amt,0))/1000) \"type8_amt\", "	//其他政策性貸款 保證金額 
			       + "			   0 \"minmonth\", "	//本年份之起始月份
			       + "			   0 \"maxmonth\", "	//本年份之結束月份
			       + "			   null "
			       + "		from m01 "
			       + "		where lpad(m_year,3,'0')||lpad(m_month,2,'0')>=lpad(?,3,'0') || lpad(?,2,'0') and "
			       + "		      lpad(m_year,3,'0')||lpad(m_month,2,'0')<=lpad(?,3,'0') || lpad(?,2,'0') and "
			       + "		      substr(data_range,4,2) in ('MM','MT') "
			       + "		group by m_year, "
			       + "		         m_month "
			       + "	)union( "
			       + "		select m_year, "
			       + "			   0 as m_month, "
			       + "			   sum(decode(guarantee_item_no,'0',guarantee_cnt,0)) \"type0_cnt\", "	//總計 件數 
			       + "			   round(sum(decode(guarantee_item_no,'0',guarantee_amt,0))/1000) \"type0_amt\", "	//總計 保證金額 
			       + "			   sum(decode(guarantee_item_no,'1',guarantee_cnt,0)) \"type1_cnt\", "	//加速農村建設貸款件數 
			       + "			   round(sum(decode(guarantee_item_no,'1',guarantee_amt,0))/1000) \"type1_amt\", "	//加速農村建設貸款保證金額 
			       + "			   sum(decode(guarantee_item_no,'2',guarantee_cnt,0)) \"type2_cnt\", "	//農業發展基金農機貸款件數 
			       + "			   round(sum(decode(guarantee_item_no,'2',guarantee_amt,0))/1000) \"type2_amt\", "	//農業發展基金農機貸款保證金額 
			       + "			   sum(decode(guarantee_item_no,'3',guarantee_cnt,0)) \"type3_cnt\", "	//農業專案貸款件數 
			       + "			   round(sum(decode(guarantee_item_no,'3',guarantee_amt,0))/1000) \"type3_amt\", "	//農業專案貸款保證金額 
			       + "			   sum(decode(guarantee_item_no,'4',guarantee_cnt,0)) \"type4_cnt\", "	// 統  一 農  貸 件數 
			       + "			   round(sum(decode(guarantee_item_no,'4',guarantee_amt,0))/1000) \"type4_amt\", "	// 統  一 農  貸 保證金額 
			       + "			   sum(decode(guarantee_item_no,'5',guarantee_cnt,0)) \"type5_cnt\", "	// 輔 導 漁  貸 件數 
			       + "			   round(sum(decode(guarantee_item_no,'5',guarantee_amt,0))/1000) \"type5_amt\", "	// 輔 導 漁  貸 保證金額 
			       + "			   sum(decode(guarantee_item_no,'6',guarantee_cnt,0)) \"type6_cnt\", "	//農漁會會員一般貸款 件數 
			       + "			   round(sum(decode(guarantee_item_no,'6',guarantee_amt,0))/1000) \"type6_amt\", "	//農漁會會員一般貸款 保證金額 
			       + "			   sum(decode(guarantee_item_no,'8',guarantee_cnt,0)) \"type8_cnt\", "	//其他政策性貸款 件數 
			       + "			   round(sum(decode(guarantee_item_no,'8',guarantee_amt,0))/1000) \"type8_amt\", "	//其他政策性貸款 保證金額 
			       + "			   min(m_month) \"minmonth\", "	//本年份之起始月份
			       + "			   max(m_month) \"maxmonth\", "	//本年份之結束月份
			       + "			   null "
			       + "		from m01 "
			       + "		where lpad(m_year,3,'0')||lpad(m_month,2,'0')>=lpad(?,3,'0') || lpad(?,2,'0') and "
			       + "		      lpad(m_year,3,'0')||lpad(m_month,2,'0')<=lpad(?,3,'0') || lpad(?,2,'0') and "
			       + "		      substr(data_range,4,2) in ('MM','MT') "
			       + "		group by m_year "
			       + "	) "
			       + ")order by m_year,m_month ";
			paramList.add(s_year);
			paramList.add(s_month);
			paramList.add(e_year);
			paramList.add(e_month);
			paramList.add(s_year);
			paramList.add(s_month);
			paramList.add(e_year);
			paramList.add(e_month);
			System.out.println("sql="+sqlCmd);
			dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"m_year,m_month,type0_cnt,type0_amt,type1_cnt,type1_amt,type2_cnt,type2_amt,type3_cnt,type3_amt,type4_cnt,type4_amt,type5_cnt,type5_amt,type6_cnt,type6_amt,type8_cnt,type8_amt,minmonth,maxmonth");	  	         
			System.out.println("dbData="+dbData.size());
			//加上列印日期
			if(datestate.equals("1")){
				row=(sheet.getRow(1)==null)? sheet.createRow(1) : sheet.getRow(1);
				Calendar rightNow = Calendar.getInstance();
				String year = String.valueOf(rightNow.get(Calendar.YEAR)-1911);
				String month = String.valueOf(rightNow.get(Calendar.MONTH)+1);
				String day = String.valueOf(rightNow.get(Calendar.DAY_OF_MONTH));
				insertCell(dbData,false,0,"列印日期："+year+"年"+month+"月"+day+"日",wb,row,(short)0, (short)65,(short)0,HSSFCellStyle.BORDER_MEDIUM,(short)0,(short)0);
            }
			int j=0;
			short top=0,down=0;
  			for(i=0;i<dbData.size();i++){
  				j=0;
  				row=(sheet.getRow(i+38)==null)? sheet.createRow(i+38) : sheet.getRow(i+38);
  				row.setHeightInPoints((float)30);
  				String m_year 	= (((DataObject)dbData.get(i)).getValue("m_year")).toString();
  				String m_month 	= (((DataObject)dbData.get(i)).getValue("m_month")).toString();
  				String minmonth = (((DataObject)dbData.get(i)).getValue("minmonth")).toString();
  				String maxmonth = (((DataObject)dbData.get(i)).getValue("maxmonth")).toString();
  				top  = (i==0)?HSSFCellStyle.BORDER_MEDIUM:HSSFCellStyle.BORDER_NONE;
  				down = (i==dbData.size()-1)?HSSFCellStyle.BORDER_MEDIUM:HSSFCellStyle.BORDER_NONE;
  				
  				if(m_month.equals("0")) insertCell(dbData,false,i,m_year+"年"+minmonth+"-"+maxmonth+"月",wb,row,(short)j, (short)26,
  				                                   top,down,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_MEDIUM);
				else   				    insertCell(dbData,false,i,"      "+m_month+"月",wb,row,(short)j, (short)26,
  				                                   top,down,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_MEDIUM);
				j++;
  				insertCell(dbData,true,i,"type0_cnt",wb,row,(short)j, (short)26,top,down,HSSFCellStyle.BORDER_MEDIUM,HSSFCellStyle.BORDER_THIN);j++;
  				insertCell(dbData,true,i,"type0_amt",wb,row,(short)j, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);j++;
  				insertCell(dbData,true,i,"type1_cnt",wb,row,(short)j, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);j++;
  				insertCell(dbData,true,i,"type1_amt",wb,row,(short)j, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);j++;
  				insertCell(dbData,true,i,"type2_cnt",wb,row,(short)j, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);j++;
  				insertCell(dbData,true,i,"type2_amt",wb,row,(short)j, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);j++;
  				insertCell(dbData,true,i,"type3_cnt",wb,row,(short)j, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);j++;
  				insertCell(dbData,true,i,"type3_amt",wb,row,(short)j, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);j++;
  				insertCell(dbData,true,i,"type4_cnt",wb,row,(short)j, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);j++;
  				insertCell(dbData,true,i,"type4_amt",wb,row,(short)j, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);j++;
  				insertCell(dbData,true,i,"type5_cnt",wb,row,(short)j, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);j++;
  				insertCell(dbData,true,i,"type5_amt",wb,row,(short)j, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);j++;
  				insertCell(dbData,true,i,"type6_cnt",wb,row,(short)j, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);j++;
  				insertCell(dbData,true,i,"type6_amt",wb,row,(short)j, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);j++;
  				insertCell(dbData,true,i,"type8_cnt",wb,row,(short)j, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);j++;
  				insertCell(dbData,true,i,"type8_amt",wb,row,(short)j, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE);j++;
	       	}

			sheet.addMergedRegion( new Region( ( short )(i+38), ( short )0,
                                               ( short )(i+38),
                                               ( short )16));
	       	row=(sheet.getRow(i+38)==null)? sheet.createRow(i+38) : sheet.getRow(i+38);i++;
	       	row.setHeightInPoints((float)30);
	       	//insertCell(dbData,true,i,"type6_cnt",wb,row,(short)j, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);j++;
	       	insertCell(dbData,false,i,"資料來源：財團法人農業信用保證基金",wb,row,(short)0, (short)65,(short)0,(short)0,(short)0,(short)0);
			sheet.addMergedRegion( new Region( ( short )(i+38), ( short )0,
                                               ( short )(i+38),
                                               ( short )16));
			row=(sheet.getRow(i+38)==null)? sheet.createRow(i+38) : sheet.getRow(i+38);
			row.setHeightInPoints((float)30);
			insertCell(dbData,false,i,"說　　明：本基金於民國73年3月正式開辦",wb,row,(short)0, (short)65,(short)0,(short)0,(short)0,(short)0);
	       	//insertCell(dbData,false,i,"說　　明：本基金於民國73年3月正式開辦",wb,row,(short)0, 0,
  			//	       HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_MEDIUM);
	       	System.out.println(reportDir + System.getProperty("file.separator")+"農業信用保證基金業務統計(一).xls");
	        FileOutputStream fout=new FileOutputStream(reportDir + System.getProperty("file.separator")+"農業信用保證基金業務統計(一).xls");
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
	                              short bg,
							      short bordertop,short borderbutton,short borderleft,short borderright)
	{
		String insertValue="";
  		if(getstate) insertValue= (((DataObject)dbData.get(index)).getValue(Item)).toString();
  		else         insertValue= Item;
		System.out.println("insertValue="+insertValue);
	    HSSFCell cell=(row.getCell(j)==null)? row.createCell(j) : row.getCell(j); 
	    HSSFCellStyle cs1 = wb.createCellStyle();
	    //HSSFCellStyle cs1 = cell.getCellStyle();
	    /*System.out.println("getFillPattern="+cs2.getFillPattern());
	    System.out.println("getFillForegroundColor="+cs2.getFillForegroundColor());
	    System.out.println("getFillBackgroundColor="+cs2.getFillBackgroundColor());
	    System.out.println("setBorderTop="+cs2.getBorderTop());
	    System.out.println("setBorderBottom="+cs2.getBorderBottom());
	    System.out.println("setBorderLeft="+cs2.getBorderLeft());
	    System.out.println("setBorderRight="+cs2.getBorderRight());
	    */
	    cs1.setBorderTop(bordertop);
	    cs1.setBorderBottom(borderbutton);
	    cs1.setBorderLeft(borderleft);
	    cs1.setBorderRight(borderright);
	    cs1.setFillPattern((short)1);
	    cs1.setFillForegroundColor(bg);
	    HSSFFont f = wb.createFont();
	    f.setFontHeightInPoints((short) 12);
	    if(j==0) f.setFontName("標楷體");
	    cs1.setFont(f);
	    cell.setCellStyle(cs1);
	    cell.setEncoding( HSSFCell.ENCODING_UTF_16 );	       			       		
	    double value=0;
	    try{
	    	cs1.setDataFormat((short)3);	// "#,##0"
	    	cell.setCellValue(Double.parseDouble(insertValue));
	    }catch(NumberFormatException e){
	    	cell.setCellValue(insertValue);
	    }
	}
}
