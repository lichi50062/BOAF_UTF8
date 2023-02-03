/*
 * Created on 2005/12/14 by 4183 lilic0c0
 * fixed on 2006/1/20 by 4183 lilic0c0 (修改金額單位先除會有rounding error的問題)
 * 95.12.05 fix sql 缺少fieldI_XF3的問題 by 2295
 * 96.01.02 fix sql 明細金額加上單位別 by 2295 
 * 99.09.15 fix 根據查詢年度.100年以後取得新縣市別.100年以前取得舊縣市別 
 * 			  使用PreparedStatement;並列印轉換後的SQL;套用QueryDB_SQLParam by 2295
 * 102.11.19 add 原QueryDB改套用QueryDB_SQLParam by 2295    
 */
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import java.io.*;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class FR002WA_Excel {	
	/* ===============================================
	 * create 經營指標變化表 report
	 * ===============================================
	 */
	public static String createRpt(String s_year,String s_month,String unit,String bank_type,String HSIEN_ID){
		
		String errMsg = "";
		String bank_type_name =( bank_type.equals("6") )?"農會":"漁會"; 
		String unit_name = Utility.getUnitName(unit);//取得單位名稱
		String hsienName = getHsienName(s_year,HSIEN_ID);  //取得縣市名稱
      	
      	try{
      		
      		//生出報表及設定格式===========================  
      	
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
						
			String openfile="各縣市各農漁會信用部各年月經營指標變化表.xls";//要去開啟的範本檔
						
			System.out.println("開啟檔:" + openfile);
			FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+ openfile );			
				
			//設定FileINputStream讀取Excel檔
			
			//新增一個xls unit
			POIFSFileSystem fs = new POIFSFileSystem( finput );
			
			if(fs==null){
				System.out.println("open 範本檔失敗");
			} else{ 
				System.out.println("open 範本檔成功");
			}
			
			//新增一個sheet
			HSSFWorkbook wb = new HSSFWorkbook(fs);
			
			if(wb==null){
				System.out.println("open工作表失敗");
			} else {
				System.out.println("open 工作表成功");
			}
			
			//對第一個sheet工作
			HSSFSheet sheet = wb.getSheetAt(0);			//讀取第一個工作表，宣告其為sheet 
						
			if(sheet==null){
				System.out.println("open sheet 失敗");
			}else {
				System.out.println("open sheet 成功");
			}
			
			//做屬性設定
			HSSFPrintSetup ps = sheet.getPrintSetup(); 	//取得設定
			//sheet.setZoom(80, 100); 					// 螢幕上看到的縮放大小
			//sheet.setAutobreaks(true); 				//自動分頁
			
			//設定頁面符合列印大小
			sheet.setAutobreaks( false );
			ps.setScale( ( short )70 ); 				//列印縮放百分比
			
			ps.setPaperSize( ( short )9 ); 				//設定紙張大小 A4
			
			//設定表頭 為固定 先設欄的起始再設列的起始
			wb.setRepeatingRowsAndColumns(0, 1, 21, 2, 3);
			
			finput.close();
			
			HSSFRow row=null;//宣告一列 
			HSSFCell cell=null;//宣告一個儲存格

			//建表開始 ===================================
			
			int rowNum =4;//表頭有4列是已經有且固定的資料，所以row從第5(index為4)列開始生
			int columnNum =19;//本excel總共有19行
			//取得變化表資料
      		List dbData = getData(s_year,s_month,unit,bank_type,HSIEN_ID);
      		DataObject bean;
      		
      		if(dbData == null){
      			System.out.println("dbData is null !!");
			} else {
	 			
	 			//first row
	  			row = sheet.getRow(0);                                                                        
	   			cell = row.getCell((short)4);
	   			cell.setEncoding(HSSFCell.ENCODING_UTF_16);//顯示中文字
	   			cell.setCellValue(hsienName +"各"+bank_type_name+"信用部"+s_year+"年"+s_month+"月"+"經營指標變化表");
	   			
	   			//second row
				row = sheet.getRow(1);
				cell = row.getCell((short)16);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellValue("單位:新臺幣 " + unit_name + ",%"); 
				
				//create all new row and cell for stroe table data 
				HSSFCellStyle cs = wb.createCellStyle();
				HSSFFont ft = wb.createFont();
				ft.setFontHeightInPoints((short)10);
				
				HSSFCellStyle csBankName = wb.createCellStyle();
				HSSFFont ftBankName = wb.createFont();
	 			ftBankName.setFontHeightInPoints((short)10);
				ftBankName.setFontName("標楷體");
	 			
				//set format for cell
	 			cs.setFont(ft);                                
				cs.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
				cs.setBorderTop(HSSFCellStyle.BORDER_THIN);
				cs.setBorderBottom(HSSFCellStyle.BORDER_THIN);
				cs.setBorderLeft(HSSFCellStyle.BORDER_THIN);
				cs.setBorderRight(HSSFCellStyle.BORDER_THIN);
				cs.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
				
				csBankName.setFont(ftBankName);
				csBankName.setAlignment(HSSFCellStyle.ALIGN_LEFT);
				csBankName.setBorderTop(HSSFCellStyle.BORDER_THIN);   
				csBankName.setBorderBottom(HSSFCellStyle.BORDER_THIN);
				csBankName.setBorderLeft(HSSFCellStyle.BORDER_THIN);  
				csBankName.setBorderRight(HSSFCellStyle.BORDER_THIN); 
				
				for(int i=0;i<dbData.size();i++){

					row = sheet.createRow(rowNum+i);
					
					for(int j=0;j<columnNum;j++){
						cell = row.createCell( (short) j);
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					}
					//取得資料
					bean = (DataObject) dbData.get(i);
					
					//塞資料
					row = sheet.getRow(rowNum+i);
					cell = row.getCell((short)0);
					cell.setCellValue(bean.getValue("bank_name").toString());
					cell.setCellStyle(csBankName); 
					
					cell = row.getCell((short)1);
					cell.setCellValue(Utility.setCommaFormat(bean.getValue("field_debit").toString()));
					cell.setCellStyle(cs); 
					
					cell = row.getCell((short)2);
					cell.setCellValue(Utility.setCommaFormat(bean.getValue("field_debit_l").toString()));
					cell.setCellStyle(cs); 
					
					cell = row.getCell((short)3);
					cell.setCellValue(Utility.setCommaFormat(bean.getValue("field_debit_lc").toString()));
					cell.setCellStyle(cs);
					
					cell = row.getCell((short)4);
					cell.setCellValue(Utility.setCommaFormat(bean.getValue("field_debit_lc_rate").toString()));
					cell.setCellStyle(cs);
					
					cell = row.getCell((short)5);
					cell.setCellValue(Utility.setCommaFormat(bean.getValue("field_credit").toString()));
					cell.setCellStyle(cs);
					
					cell = row.getCell((short)6);
					cell.setCellValue(Utility.setCommaFormat(bean.getValue("field_credit_l").toString()));
					cell.setCellStyle(cs);
					
					cell = row.getCell((short)7);
					cell.setCellValue(Utility.setCommaFormat(bean.getValue("field_credit_lc").toString()));
					cell.setCellStyle(cs);
					
					cell = row.getCell((short)8);
					cell.setCellValue(Utility.setCommaFormat(bean.getValue("field_credit_lc_rate").toString()));
					cell.setCellStyle(cs);
					
					cell = row.getCell((short)9);
					cell.setCellValue(Utility.setCommaFormat(bean.getValue("field_dc_rate").toString()));
					cell.setCellStyle(cs);
					
					cell = row.getCell((short)10);
					cell.setCellValue(Utility.setCommaFormat(bean.getValue("field_dc_rate_l").toString()));
					cell.setCellStyle(cs);
					
					cell = row.getCell((short)11);
					cell.setCellValue(Utility.setCommaFormat(bean.getValue("field_120700").toString()));
					cell.setCellStyle(cs);
					
					cell = row.getCell((short)12);
					cell.setCellValue(Utility.setCommaFormat(bean.getValue("field_over").toString()));
					cell.setCellStyle(cs);
					
					cell = row.getCell((short)13);
					cell.setCellValue(Utility.setCommaFormat(bean.getValue("c_field_over_rate").toString()));
					cell.setCellStyle(cs);
					
					cell = row.getCell((short)14);
					cell.setCellValue(Utility.setCommaFormat(bean.getValue("field_over_l").toString()));
					cell.setCellStyle(cs);
					
					cell = row.getCell((short)15);
					cell.setCellValue(Utility.setCommaFormat(bean.getValue("c_field_over_rate_l").toString()));
					cell.setCellStyle(cs);
					
					cell = row.getCell((short)16);
					cell.setCellValue(Utility.setCommaFormat(bean.getValue("field_310000").toString()));
					cell.setCellStyle(cs);
					
					cell = row.getCell((short)17);
					cell.setCellValue(Utility.setCommaFormat(bean.getValue("field_320000").toString()));
					cell.setCellStyle(cs);
					
					cell = row.getCell((short)18);
					cell.setCellValue(Utility.setCommaFormat(bean.getValue("field_220900").toString()));
					cell.setCellStyle(cs); 
				}
				
			}
			//建表結束 =================================
			HSSFFooter footer = sheet.getFooter();
			footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
			
			FileOutputStream fout=new FileOutputStream(reportDir + System.getProperty("file.separator")+
													   "各縣市各" + bank_type_name + "信用部各年月經營指標變化表.xls");  
			
			wb.write(fout);//儲存
			fout.close();
			
			System.out.println("儲存完成");
			
		}catch(Exception e){
			System.out.println("createRpt Error:"+e+e.getMessage());                                                
		}
		
		return errMsg;
	}//end of createreport
	
	
	/*	===============================================
	 *	由縣市ID取得縣市名稱
	 *  ===============================================
 	 */ 
	public static String getHsienName(String s_year,String HSIEN_ID){
	    List paramList = new ArrayList();
		//99.09.15 add 查詢年度100年以前.縣市別不同===============================
	    String cd01_table = (Integer.parseInt(s_year) < 100)?"cd01_99":"";
	    //=====================================================================    
	    paramList.add(HSIEN_ID);
		List dbData = DBManager.QueryDB_SQLParam("select hsien_name from "+cd01_table+" cd01 where hsien_id = ?",paramList,"");
		String rtnStr="";
		if(dbData.size()>0){
		    rtnStr=String.valueOf(((DataObject)dbData.get(0)).getValue("hsien_name"));
		}
		return rtnStr;
	}//end of getHsienName
	
	/*  ===============================================
	 *	取得變化明細表的各項資料
	 *  ===============================================
	 */
	public static List getData(String s_year,String s_month,String unit,String bank_type,String HSIEN_ID){
		
		String l_year = s_year;
		String l_month = "";
		//99.09.15 add 查詢年度100年以前.縣市別不同===============================
	    String cd01_table = (Integer.parseInt(s_year) < 100)?"cd01_99":""; 
	    String wlx01_m_year = (Integer.parseInt(s_year) < 100)?"99":"100"; 
	    //=====================================================================    
		//計算上個月份的年跟月(如果本月是一月，年份會不同)
		if(Integer.parseInt(s_month) == 1){
			l_month = "12";
			int temp = Integer.parseInt(l_year)-1;
			l_year = String.valueOf(temp);
		}
		else{
			int temp = Integer.parseInt(s_month)-1;
			l_month = String.valueOf(temp--);
		}
		
		List dbData = null;
		StringBuffer sqlCmd = new StringBuffer();
		StringBuffer sql_Sum = new StringBuffer();			//合計的sql語法
		StringBuffer sql_Sum_upper = new StringBuffer();	//合計上半部的sql語法
		StringBuffer sql_Sum_Combine =new StringBuffer();	//所有下轄農漁會資料合併後的sql語法
		StringBuffer sql_Detail = new StringBuffer();		//明細的sql語法
		StringBuffer sql_Detail_upper = new StringBuffer();	//明細上半部的sql語法
		StringBuffer sql_Combine =new StringBuffer();		//上下月份資料合併的sql語法
		StringBuffer sql_MonthData =new StringBuffer();		//一個月份資料的sql語法                                                                                                            
		StringBuffer sql_ThisMonth =new StringBuffer();		//這個月的資料的sql語法
		StringBuffer sql_LastMonth =new StringBuffer();		//上個月的資料的sql語法
		List paramList = new ArrayList();//傳入參數
		List paramList_sql_ThisMonth = new ArrayList();//傳入參數
		List paramList_sql_LastMonth = new ArrayList();//傳入參數
		List paramList_sql_Combine = new ArrayList();//傳入參數
		List paramList_sql_Detail = new ArrayList();//傳入參數
		List paramList_sql_Sum = new ArrayList();//傳入參數

		sqlCmd.append(" Select a01.*  ");
		sqlCmd.append(" from(  ");
		                      //明細資料 ================================
		sqlCmd.append("       select a01.hsien_id, a01.hsien_name, a01.FR001W_output_order, a01.bank_no, a01.BANK_NAME, "); 
		sqlCmd.append("               round(field_DEBIT /?,0) as field_DEBIT, ");//--存款(含公庫存款).本月份
		sqlCmd.append("               round(field_DEBIT_L /?,0) as field_DEBIT_L, ");//--存款(含公庫存款).上月份
		sqlCmd.append("               round((field_DEBIT  -  field_DEBIT_L) /?,0) as field_DEBIT_LC, ");//--存款(含公庫存款).增減金額
		sqlCmd.append("               decode( field_DEBIT_L,0,0,round((field_DEBIT - field_DEBIT_L)/field_DEBIT_L * 100,2)) as field_DEBIT_LC_RATE, ");//--存款(含公庫存款).增減%  
		sqlCmd.append("               round(field_CREDIT /?,0) as field_CREDIT, ");//--放款.本月份
		sqlCmd.append("               round(field_CREDIT_L /?,0) as field_CREDIT_L, ");//--放款.上月份
		sqlCmd.append("               round((field_CREDIT  - field_CREDIT_L) /?,0) as field_CREDIT_LC, ");//--放款.增減金額
		sqlCmd.append("               decode( field_CREDIT_L,0,0,round((field_CREDIT - field_CREDIT_L)/field_CREDIT_L * 100,2)) as field_CREDIT_LC_RATE, ");//--放款.增減% 
		sqlCmd.append("               decode( a01.fieldI_Y,0,0,round(  ");
		sqlCmd.append("               (a01.fieldI_XA    +  decode( sign(a01.fieldI_XB1 - a01.fieldI_XB2),-1,0,(a01.fieldI_XB1 - a01.fieldI_XB2)) +  "); 
		sqlCmd.append("               decode( sign(a01.fieldI_XC1 - a01.fieldI_XC2),-1,0,(a01.fieldI_XC1 - a01.fieldI_XC2)) +  "); 
		sqlCmd.append("               decode( sign(a01.fieldI_XD1 - a01.fieldI_XD2),-1,0,(a01.fieldI_XD1 - a01.fieldI_XD2)) +  ");
		sqlCmd.append("                decode( sign(a01.fieldI_XE1 - a01.fieldI_XE2),-1,0,(a01.fieldI_XE1 - a01.fieldI_XE2)) - ");  
		sqlCmd.append("                decode( sign(a01.fieldI_XF1 - a01.fieldI_XF3 - a01.fieldI_XF2),-1,0,(a01.fieldI_XF1 - a01.fieldI_XF3 - a01.fieldI_XF2))) ");
		sqlCmd.append("                / a01.fieldI_Y * 100,2)) as field_DC_RATE, ");//--存放比率.本月份 
		sqlCmd.append("               decode( a01.fieldI_Y_L,0,0, ");
		sqlCmd.append("               round( (a01.fieldI_XA_L +  decode( sign(a01.fieldI_XB1_L - a01.fieldI_XB2_L),-1,0,(a01.fieldI_XB1_L - a01.fieldI_XB2_L)) + ");  
		sqlCmd.append("               decode( sign(a01.fieldI_XC1_L - a01.fieldI_XC2_L),-1,0,(a01.fieldI_XC1_L - a01.fieldI_XC2_L)) + ");  
		sqlCmd.append("               decode( sign(a01.fieldI_XD1_L - a01.fieldI_XD2_L),-1,0,(a01.fieldI_XD1_L - a01.fieldI_XD2_L)) + ");  
		sqlCmd.append("               decode( sign(a01.fieldI_XE1_L - a01.fieldI_XE2_L),-1,0,(a01.fieldI_XE1_L - a01.fieldI_XE2_L)) - ");  
		sqlCmd.append("               decode( sign(a01.fieldI_XF1_L - a01.fieldI_XF2_L),-1,0,(a01.fieldI_XF1_L - a01.fieldI_XF2_L)) ) ");
		sqlCmd.append("               / a01.fieldI_Y_L * 100,2)) as field_DC_RATE_L, ");//--存放比率.上月份 
		sqlCmd.append("               round(field_120700 /?,0) as field_120700, ");//--內部融資
		sqlCmd.append("               round(field_OVER /?,0) as field_OVER, ");//--逾期放款.本月份
		sqlCmd.append("               round(field_OVER_L /?,0) as field_OVER_L, "); //--逾期放款.上月份
		sqlCmd.append("               decode(a01.field_CREDIT,0,0,round(a01.field_OVER / a01.field_CREDIT *100 ,2)) as C_field_OVER_RATE, ");//--逾期放款.本月份.%  
		sqlCmd.append("               decode(a01.field_CREDIT_L,0,0,round(a01.field_OVER_L / a01.field_CREDIT_L *100 ,2)) as C_field_OVER_RATE_L, ");//--逾期放款.上月份.%  
		sqlCmd.append("               round(field_310000 /?,0) as field_310000, ");//--事業資金及公積
		sqlCmd.append("               round(field_320000 /?,0) as field_320000, ");//--盈虧及損益
		sqlCmd.append("               round(field_220900 /?,0) as field_220900  ");//--公庫存款
		for(int k=1;k<=12;k++){
		    paramList.add(unit);
		}
		sqlCmd.append("        from (  select a01.hsien_id, a01.hsien_name, a01.FR001W_output_order, a01.bank_no, a01.BANK_NAME, "); 
		sqlCmd.append("                 sum(field_120700_L) field_120700_L, sum(field_120700)    field_120700, ");  
		sqlCmd.append("                 sum(field_310000_L) field_310000_L, sum(field_310000)    field_310000, ");  
		sqlCmd.append("                 sum(field_320000_L) field_320000_L, sum(field_320000)    field_320000, ");  
		sqlCmd.append("                 sum(field_220900_L) field_220900_L, sum(field_220900)    field_220900, ");  
		sqlCmd.append("                 sum(field_OVER_L)   field_OVER_L,   sum(field_OVER)      field_OVER,   ");  
		sqlCmd.append("                 sum(field_DEBIT_L)  field_DEBIT_L,  sum(field_DEBIT)     field_DEBIT,  ");  
		sqlCmd.append("                 sum(field_CREDIT_L) field_CREDIT_L, sum(field_CREDIT)    field_CREDIT, ");  
		sqlCmd.append("                 sum(fieldI_XA_L)     fieldI_XA_L,     sum(fieldI_XA)        fieldI_XA, ");     
		sqlCmd.append("                 sum(fieldI_XB1_L)    fieldI_XB1_L,     sum(fieldI_XB1)    fieldI_XB1,  ");   
		sqlCmd.append("                 sum(fieldI_XB2_L)    fieldI_XB2_L,     sum(fieldI_XB2)    fieldI_XB2,  ");   
		sqlCmd.append("                 sum(fieldI_XC1_L)    fieldI_XC1_L,     sum(fieldI_XC1)    fieldI_XC1,  ");   
		sqlCmd.append("                 sum(fieldI_XC2_L)    fieldI_XC2_L,     sum(fieldI_XC2)    fieldI_XC2,  ");   
		sqlCmd.append("                 sum(fieldI_XD1_L)    fieldI_XD1_L,     sum(fieldI_XD1)    fieldI_XD1,  ");   
		sqlCmd.append("                 sum(fieldI_XD2_L)    fieldI_XD2_L,     sum(fieldI_XD2)    fieldI_XD2,  ");   
		sqlCmd.append("                 sum(fieldI_XE1_L)    fieldI_XE1_L,     sum(fieldI_XE1)    fieldI_XE1,  ");   
		sqlCmd.append("                 sum(fieldI_XE2_L)    fieldI_XE2_L,     sum(fieldI_XE2)    fieldI_XE2,  ");   
		sqlCmd.append("                 sum(fieldI_XF1_L)    fieldI_XF1_L,     sum(fieldI_XF1)    fieldI_XF1,  ");   
		sqlCmd.append("                 sum(fieldI_XF2_L)    fieldI_XF2_L,     sum(fieldI_XF2)    fieldI_XF2,  ");   
		sqlCmd.append("                 SUM(fieldI_XF3)     fieldI_XF3, ");  
		sqlCmd.append("                 sum(fieldI_Y_L)        fieldI_Y_L, sum(fieldI_Y)    fieldI_Y  ");
		sqlCmd.append("              from (  ");//--本月份資料
		sqlCmd.append("                     select a01.hsien_id, a01.hsien_name, a01.FR001W_output_order, a01.bank_no, a01.BANK_NAME, "); 
		sqlCmd.append("                     0 as field_120700_L, field_120700, 0 as field_310000_L,     field_310000,  0 as field_320000_L, field_320000, "); 
		sqlCmd.append("                     0 as  field_220900_L,    field_220900,  0 as field_OVER_L,   field_OVER,0 as field_DEBIT_L,field_DEBIT,       ");
		sqlCmd.append("                     0 as field_CREDIT_L, field_CREDIT, 0 as fieldI_XA_L,        fieldI_XA,     0 as fieldI_XB1_L,   fieldI_XB1,   "); 
		sqlCmd.append("                     0 as fieldI_XB2_L,       fieldI_XB2,    0 as fieldI_XC1_L,   fieldI_XC1,   0 as fieldI_XC2_L,       fieldI_XC2, ");    
		sqlCmd.append("                     0 as fieldI_XD1_L,   fieldI_XD1,   0 as fieldI_XD2_L,       fieldI_XD2,    0 as fieldI_XE1_L,   fieldI_XE1,     ");
		sqlCmd.append("                     0 as fieldI_XE2_L,       fieldI_XE2,    0 as fieldI_XF1_L,   fieldI_XF1,   0 as fieldI_XF2_L,       fieldI_XF2, ");  
		sqlCmd.append("                     fieldI_XF3, 0 as fieldI_Y_L,     fieldI_Y ");     
		sqlCmd.append("                    from (  select nvl(cd01.hsien_id,' ') as hsien_id,nvl(cd01.hsien_name,'OTHER') as hsien_name, ");  
		sqlCmd.append(" cd01.FR001W_output_order as FR001W_output_order,bn01.bank_no,bn01.BANK_NAME, ");  
		sqlCmd.append("                             round(sum(decode(a01.acc_code,'120700',amt,0)) /1,0) as field_120700, ");  
		sqlCmd.append("                             round(sum(decode(a01.acc_code,'310000',amt,0)) /1,0) as field_310000, ");  
		sqlCmd.append("                             round(sum(decode(a01.acc_code,'320000',amt,0)) /1,0) as field_320000, ");  
		sqlCmd.append("                             round(sum(decode(a01.acc_code,'220900',amt,0)) /1,0) as field_220900, ");  
		sqlCmd.append("                             round(sum(decode(a01.acc_code,'990000',amt,0)) /1,0) as field_OVER,   ");  
		sqlCmd.append("                             round(sum(decode(a01.acc_code,'220000',amt,0)) /1,0) as field_DEBIT,  ");  
		sqlCmd.append("                             round(sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /1,0) as field_CREDIT, ");
		sqlCmd.append("                             decode(YEAR_TYPE,'102',decode(bank_type,'6',round(sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120200',amt,'120301',amt,'120302',amt,'120700',amt,'150200',amt,0)) /1,0), ");                  
		sqlCmd.append("                                                                      '7',round(sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120300',amt,'120401',amt,'120402',amt,'120700',amt,'150200',amt,0)) /1,0)),  ");
		sqlCmd.append("                                              '103',round(sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120200',amt,'120301',amt,'120302',amt,'120700',amt,'150200',amt,0)) /1,0),0) as fieldI_XA, ");  
		sqlCmd.append("                             decode(YEAR_TYPE,'102',decode(bank_type,'6',round(sum(decode(a01.acc_code,'120401',amt,'120402',amt,0)) /1,0), ");                   
		sqlCmd.append("                                                                      '7',round(sum(decode(a01.acc_code,'120201',amt,'120202',amt,0)) /1,0)), ");
		sqlCmd.append("                                              '103',round(sum(decode(a01.acc_code,'120401',amt,'120402',amt,0)) /1,0),0) as fieldI_XB1, ");  
		sqlCmd.append("                             decode(YEAR_TYPE,'102',decode(bank_type,'6',sum(decode(a01.acc_code,'240305',amt,'251300',amt,'310800',amt,0)), ");                  
		sqlCmd.append("                                                                      '7',sum(decode(a01.acc_code,'240205',amt,'310800',amt,0))), ");
		sqlCmd.append("                                              '103',decode(bank_type,'6',sum(decode(a01.acc_code,'240305',amt,'251300',amt,'310800',amt,0)), ");                  
		sqlCmd.append("                                                                      '7',sum(decode(a01.acc_code,'240305',amt,'251200',amt,0))),0)  as fieldI_XB2, ");  
		sqlCmd.append("                             round(sum(decode(a01.acc_code,'120501',amt,'120502',amt,0)) /1,0)    as fieldI_XC1, ");      
		sqlCmd.append("                             decode(YEAR_TYPE,'102',decode(bank_type,'6',round(sum(decode(a01.acc_code,'240301',amt,'240302',amt,'240303',amt,'240304',amt,0)) /1,0), ");  
		sqlCmd.append("                                                                       '7',round(sum(decode(a01.acc_code,'240201',amt,'240202',amt,'240203',amt,'240204',amt,0)) /1,0)), ");
		sqlCmd.append("                                               '103',round(sum(decode(a01.acc_code,'240301',amt,'240302',amt,'240303',amt,'240304',amt,0)) /1,0),0) as fieldI_XC2, ");  
		sqlCmd.append("                             round(sum(decode(a01.acc_code,'120600',amt,0)) /1,0)                  as fieldI_XD1, "); 
		sqlCmd.append("                             decode(YEAR_TYPE,'102',decode(bank_type,'6',round(sum(decode(a01.acc_code,'240200',amt,0)) /1,0), ");
		sqlCmd.append("                                                                     '7',round(sum(decode(a01.acc_code,'240300',amt,0)) /1,0)),  ");
		sqlCmd.append("                                              '103',round(sum(decode(a01.acc_code,'240200',amt,0)) /1,0),0)  as fieldI_XD2, "); 
		sqlCmd.append("                             round(sum(decode(a01.acc_code,'150100',amt,0)) /1,0)                  as fieldI_XE1, "); 
		sqlCmd.append("                             round(sum(decode(a01.acc_code,'250100',amt,0)) /1,0)                  as fieldI_XE2, "); 
		sqlCmd.append("                             round(sum(decode(a01.acc_code,'310000',amt,'320000',amt,0)) /1,0)  as fieldI_XF1, "); 
		sqlCmd.append("                             decode(YEAR_TYPE,'102',round(sum(decode(a01.acc_code,'310800',amt,0)) /1,0), ");
		sqlCmd.append("                                              '103',decode(bank_type,'6',round(sum(decode(a01.acc_code,'310800',amt,0)) /1,0),7,0),0)    as fieldI_XF3, "); 
		sqlCmd.append("                             round(sum(decode(a01.acc_code,'140000',amt,0)) /1,0)                  as fieldI_XF2, "); 
		sqlCmd.append("                             round((sum(decode(a01.acc_code,'220100',amt,'220200',amt,'220300',amt,'220400',amt,  '220500',amt,'220600',amt,'220700',amt,'220800',amt,'220900',amt,'221000',amt,0))  ");
		sqlCmd.append("                             -  round(sum(decode(a01.acc_code,'220900',amt,0))/2,0)) /1,0)   as fieldI_Y ");  
		sqlCmd.append("                         from (select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id = ? )   cd01  ");
		paramList.add(HSIEN_ID);
		sqlCmd.append("                         left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id  "); 
		sqlCmd.append("                         left join (select * from bn01 where m_year=?)bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=? ");
		paramList.add(wlx01_m_year);
		paramList.add(wlx01_m_year);
		paramList.add(bank_type);
		sqlCmd.append("                         left join (select  (CASE WHEN (a01.m_year <= 102) THEN '102' ");
		sqlCmd.append("                                    WHEN (a01.m_year > 102) THEN '103' ");
		sqlCmd.append("                               ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01)a01 on a01.m_year = ? and a01.m_month = ? and a01.bank_code=bn01.bank_no "); 
		paramList.add(s_year);
        paramList.add(s_month);
		sqlCmd.append("                         group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order, a01.YEAR_TYPE, bn01.bank_type,bn01.bank_no ,  bn01.BANK_NAME  ");
		sqlCmd.append("                        ) a01  ");  
		sqlCmd.append("                  where  a01.bank_no <> ' '  ");
		sqlCmd.append("                  union all  "); 
		                                 //上月份資料 ================================
		sqlCmd.append("                  select a01.hsien_id, a01.hsien_name, a01.FR001W_output_order, a01.bank_no, a01.BANK_NAME, "); 
		sqlCmd.append("                        field_120700 as field_120700_L, 0 as field_120700, field_310000 as field_310000_L, 0 as field_310000, ");  
		sqlCmd.append("                        field_320000 as field_320000_L, 0 as field_320000, field_220900 as field_220900_L, 0 as field_220900, ");  
		sqlCmd.append("                        field_OVER   as field_OVER_L,   0 as field_OVER,   field_DEBIT  as field_DEBIT_L,  0 as field_DEBIT,  ");  
		sqlCmd.append("                        field_CREDIT as field_CREDIT_L, 0 as field_CREDIT, fieldI_XA    as fieldI_XA_L,    0 as fieldI_XA,    ");  
		sqlCmd.append("                        fieldI_XB1   as fieldI_XB1_L,   0 as fieldI_XB1,   fieldI_XB2   as fieldI_XB2_L,   0 as fieldI_XB2,   ");  
		sqlCmd.append("                        fieldI_XC1   as fieldI_XC1_L,   0 as fieldI_XC1,   fieldI_XC2   as fieldI_XC2_L,   0 as fieldI_XC2,   ");  
		sqlCmd.append("                        fieldI_XD1   as fieldI_XD1_L,   0 as fieldI_XD1,   fieldI_XD2   as fieldI_XD2_L,   0 as fieldI_XD2,   ");  
		sqlCmd.append("                        fieldI_XE1   as fieldI_XE1_L,   0 as fieldI_XE1,   fieldI_XE2   as fieldI_XE2_L,   0 as fieldI_XE2,   ");  
		sqlCmd.append("                        fieldI_XF1   as fieldI_XF1_L,   0 as fieldI_XF1,   fieldI_XF2   as fieldI_XF2_L,   0 as fieldI_XF2,   ");  
		sqlCmd.append("                        0 as fieldI_XF3, fieldI_Y       as fieldI_Y_L,     0 as fieldI_Y  ");    
		sqlCmd.append("                  from (  select nvl(cd01.hsien_id,' ') as hsien_id,nvl(cd01.hsien_name,'OTHER') as hsien_name,  cd01.FR001W_output_order as FR001W_output_order,bn01.bank_no,bn01.BANK_NAME, ");  
		sqlCmd.append("                              round(sum(decode(a01.acc_code,'120700',amt,0)) /1,0) as field_120700, ");  
		sqlCmd.append("                              round(sum(decode(a01.acc_code,'310000',amt,0)) /1,0) as field_310000, ");  
		sqlCmd.append("                              round(sum(decode(a01.acc_code,'320000',amt,0)) /1,0) as field_320000, ");  
		sqlCmd.append("                              round(sum(decode(a01.acc_code,'220900',amt,0)) /1,0) as field_220900, ");  
		sqlCmd.append("                              round(sum(decode(a01.acc_code,'990000',amt,0)) /1,0) as field_OVER,   ");  
		sqlCmd.append("                              round(sum(decode(a01.acc_code,'220000',amt,0)) /1,0) as field_DEBIT,  ");  
		sqlCmd.append("                              round(sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /1,0) as field_CREDIT, ");  
		sqlCmd.append("                              decode(YEAR_TYPE,'102',decode(bank_type,'6',round(sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120200',amt,'120301',amt,'120302',amt,'120700',amt,'150200',amt,0)) /1,0),  ");                 
		sqlCmd.append("                                                                      '7',round(sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120300',amt,'120401',amt,'120402',amt,'120700',amt,'150200',amt,0)) /1,0)), ");
		sqlCmd.append("                                              '103',round(sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120200',amt,'120301',amt,'120302',amt,'120700',amt,'150200',amt,0)) /1,0),0) as fieldI_XA,        ");
		sqlCmd.append("                              decode(YEAR_TYPE,'102',decode(bank_type,'6',round(sum(decode(a01.acc_code,'120401',amt,'120402',amt,0)) /1,0),   ");                 
		sqlCmd.append("                                                                       '7',round(sum(decode(a01.acc_code,'120201',amt,'120202',amt,0)) /1,0)), ");
		sqlCmd.append("                                               '103',round(sum(decode(a01.acc_code,'120401',amt,'120402',amt,0)) /1,0),0) as fieldI_XB1,       ");
		sqlCmd.append("                              decode(YEAR_TYPE,'102',decode(bank_type,'6',sum(decode(a01.acc_code,'240305',amt,'251300',amt,'310800',amt,0)),  ");                 
		sqlCmd.append("                                                                       '7',sum(decode(a01.acc_code,'240205',amt,'310800',amt,0))),             ");
		sqlCmd.append("                                               '103',decode(bank_type,'6',sum(decode(a01.acc_code,'240305',amt,'251300',amt,'310800',amt,0)),  ");                 
		sqlCmd.append("                                                                       '7',sum(decode(a01.acc_code,'240305',amt,'251200',amt,0))),0)  as fieldI_XB2, ");  
		sqlCmd.append("                              round(sum(decode(a01.acc_code,'120501',amt,'120502',amt,0)) /1,0)    as fieldI_XC1, ");      
		sqlCmd.append("                              decode(YEAR_TYPE,'102',decode(bank_type,'6',round(sum(decode(a01.acc_code,'240301',amt,'240302',amt,'240303',amt,'240304',amt,0)) /1,0), ");  
		sqlCmd.append("                                                                        '7',round(sum(decode(a01.acc_code,'240201',amt,'240202',amt,'240203',amt,'240204',amt,0)) /1,0)), ");
		sqlCmd.append("                                                '103',round(sum(decode(a01.acc_code,'240301',amt,'240302',amt,'240303',amt,'240304',amt,0)) /1,0),0) as fieldI_XC2, ");  
		sqlCmd.append("                              round(sum(decode(a01.acc_code,'120600',amt,0)) /1,0)                  as fieldI_XD1, "); 
		sqlCmd.append("                              decode(YEAR_TYPE,'102',decode(bank_type,'6',round(sum(decode(a01.acc_code,'240200',amt,0)) /1,0),   ");
		sqlCmd.append("                                                                      '7',round(sum(decode(a01.acc_code,'240300',amt,0)) /1,0)),  ");
		sqlCmd.append("                                               '103',round(sum(decode(a01.acc_code,'240200',amt,0)) /1,0),0)  as fieldI_XD2,      ");
		sqlCmd.append("                              round(sum(decode(a01.acc_code,'150100',amt,0)) /1,0)                  as fieldI_XE1, "); 
		sqlCmd.append("                              round(sum(decode(a01.acc_code,'250100',amt,0)) /1,0)                  as fieldI_XE2, "); 
		sqlCmd.append("                              round(sum(decode(a01.acc_code,'310000',amt,'320000',amt,0)) /1,0)  as fieldI_XF1,    ");
		sqlCmd.append("                              decode(YEAR_TYPE,'102',round(sum(decode(a01.acc_code,'310800',amt,0)) /1,0),         ");
		sqlCmd.append("                                               '103',decode(bank_type,'6',round(sum(decode(a01.acc_code,'310800',amt,0)) /1,0),7,0),0)    as fieldI_XF3, "); 
		sqlCmd.append("                              round(sum(decode(a01.acc_code,'140000',amt,0)) /1,0)                  as fieldI_XF2, ");
		sqlCmd.append("                              round((sum(decode(a01.acc_code,'220100',amt,'220200',amt,'220300',amt,'220400',amt,  '220500',amt,'220600',amt,'220700',amt,'220800',amt,'220900',amt,'221000',amt,0)) -  "); 
		sqlCmd.append("                              round(sum(decode(a01.acc_code,'220900',amt,0))/2,0)) /1,0)   as fieldI_Y      ");
		sqlCmd.append("                         from (select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id = ? )   cd01   ");
		paramList.add(HSIEN_ID);
		sqlCmd.append("                         left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id  "); 
		sqlCmd.append("                         left join (select * from bn01 where m_year=?)bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=? ");
		paramList.add(wlx01_m_year);
		paramList.add(wlx01_m_year);
		paramList.add(bank_type);
		sqlCmd.append("                         left join (select  (CASE WHEN (a01.m_year <= 102) THEN '102'  ");
		sqlCmd.append("                                    WHEN (a01.m_year > 102) THEN '103'  ");
		sqlCmd.append("                               ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01)a01 on a01.m_year = ? and a01.m_month = ? and a01.bank_code=bn01.bank_no ");
		paramList.add(l_year);
        paramList.add(l_month);
		sqlCmd.append("                         group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,  a01.YEAR_TYPE, bn01.bank_type,bn01.bank_no ,  bn01.BANK_NAME ");
		sqlCmd.append("                     ) a01  "); 
		sqlCmd.append("                 where  a01.bank_no <> ' ' ");  
		sqlCmd.append("              ) a01  ");
		sqlCmd.append("              group by a01.hsien_id, a01.hsien_name, a01.FR001W_output_order, a01.bank_no, a01.BANK_NAME  "); 
		sqlCmd.append("     ) a01 ");  //--明細資料
		sqlCmd.append("     union all ");  
		                    //合計資料 ================================
		sqlCmd.append("     select a01.hsien_id, a01.hsien_name, '999' as FR001W_output_order, '9999999' as bank_no, '合計' as BANK_NAME, ");
		sqlCmd.append("            round(field_DEBIT /?,0) as field_DEBIT,  ");
		sqlCmd.append("            round(field_DEBIT_L /?,0) as field_DEBIT_L, round((field_DEBIT  -  field_DEBIT_L) /?,0) as field_DEBIT_LC,  ");
		sqlCmd.append("            decode( field_DEBIT_L,0,0,round( (field_DEBIT - field_DEBIT_L) / field_DEBIT_L * 100,2 ) )as field_DEBIT_LC_RATE, ");
		sqlCmd.append("            round(field_CREDIT /?,0) as field_CREDIT, round(field_CREDIT_L /?,0) as field_CREDIT_L, ");
		sqlCmd.append("            round((field_CREDIT  - field_CREDIT_L) /?,0) as field_CREDIT_LC, ");
		sqlCmd.append("            decode( field_CREDIT_L,0,0,round((field_CREDIT - field_CREDIT_L) / field_CREDIT_L * 100,2))as field_CREDIT_LC_RATE, "); 
		sqlCmd.append("            decode( a01.fieldI_Y,0,0,round( (a01.fieldI_XA + decode(    sign(a01.fieldI_XB1 - a01.fieldI_XB2),-1,0, (a01.fieldI_XB1 - a01.fieldI_XB2)) + decode(    sign(a01.fieldI_XC1 - a01.fieldI_XC2),-1,0,(a01.fieldI_XC1 - a01.fieldI_XC2))+ decode( sign(a01.fieldI_XD1 - a01.fieldI_XD2),-1,0,(a01.fieldI_XD1 - a01.fieldI_XD2))    + decode( sign(a01.fieldI_XE1 - a01.fieldI_XE2),-1,0,(a01.fieldI_XE1 - a01.fieldI_XE2))    - decode( sign(a01.fieldI_XF1 - a01.fieldI_XF3 - a01.fieldI_XF2),-1,0,(a01.fieldI_XF1 - a01.fieldI_XF3 - a01.fieldI_XF2)) )/ a01.fieldI_Y * 100,2))as field_DC_RATE, "); 
		sqlCmd.append("            decode( a01.fieldI_Y_L,0,0,round( (a01.fieldI_XA_L + decode( sign(a01.fieldI_XB1_L - a01.fieldI_XB2_L),-1,0, (a01.fieldI_XB1_L - a01.fieldI_XB2_L)) + decode( sign(a01.fieldI_XC1_L - a01.fieldI_XC2_L),-1,0,(a01.fieldI_XC1_L - a01.fieldI_XC2_L)) + decode( sign(a01.fieldI_XD1_L - a01.fieldI_XD2_L),-1,0,(a01.fieldI_XD1_L - a01.fieldI_XD2_L))    + decode( sign(a01.fieldI_XE1_L - a01.fieldI_XE2_L),-1,0,(a01.fieldI_XE1_L - a01.fieldI_XE2_L))    - decode( sign(a01.fieldI_XF1_L - a01.fieldI_XF2_L),-1,0,(a01.fieldI_XF1_L - a01.fieldI_XF2_L)))/a01.fieldI_Y_L * 100,2) )as field_DC_RATE_L, "); 
		sqlCmd.append("            round(field_120700 /?,0) as field_120700, ");
		sqlCmd.append("            round(field_OVER /?,0) as field_OVER,round(field_OVER_L /?,0) as field_OVER_L, ");
		sqlCmd.append("            decode(    a01.field_CREDIT,0,0,round(a01.field_OVER / a01.field_CREDIT *100 ,2) )as C_field_OVER_RATE,decode(    a01.field_CREDIT_L,0,0,round(a01.field_OVER_L / a01.field_CREDIT_L *100 ,2))as C_field_OVER_RATE_L, "); 
		sqlCmd.append("            round(field_310000 /?,0) as field_310000, ");
		sqlCmd.append("            round(field_320000 /?,0) as field_320000, ");
		sqlCmd.append("            round(field_220900 /?,0) as field_220900  ");
		for(int k=1;k<=12;k++){
            paramList.add(unit);
        }
		sqlCmd.append("     from( select a01.hsien_id,a01.hsien_name,        ");
		sqlCmd.append("                 sum(field_120700_L) field_120700_L,    sum(field_120700) field_120700,    ");
		sqlCmd.append("                 sum(field_310000_L)    field_310000_L,    sum(field_310000) field_310000, ");  
		sqlCmd.append("                 sum(field_320000_L)    field_320000_L,    sum(field_320000) field_320000, ");  
		sqlCmd.append("                 sum(field_220900_L)    field_220900_L,    sum(field_220900) field_220900, ");  
		sqlCmd.append("                 sum(field_OVER_L)    field_OVER_L,    sum(field_OVER)   field_OVER,       "); 
		sqlCmd.append("                 sum(field_DEBIT_L)     field_DEBIT_L,    sum(field_DEBIT)  field_DEBIT,   "); 
		sqlCmd.append("                 sum(field_CREDIT_L)    field_CREDIT_L,    sum(field_CREDIT) field_CREDIT, ");  
		sqlCmd.append("                 sum(fieldI_XA_L)    fieldI_XA_L,    sum(fieldI_XA)      fieldI_XA,        ");
		sqlCmd.append("                 sum(fieldI_XB1_L)    fieldI_XB1_L,    sum(fieldI_XB1)      fieldI_XB1,    "); 
		sqlCmd.append("                 sum(fieldI_XB2_L)    fieldI_XB2_L,    sum(fieldI_XB2)      fieldI_XB2,    "); 
		sqlCmd.append("                 sum(fieldI_XC1_L)    fieldI_XC1_L,    sum(fieldI_XC1)      fieldI_XC1,    "); 
		sqlCmd.append("                 sum(fieldI_XC2_L)    fieldI_XC2_L,    sum(fieldI_XC2)      fieldI_XC2,    "); 
		sqlCmd.append("                 sum(fieldI_XD1_L)    fieldI_XD1_L,    sum(fieldI_XD1)      fieldI_XD1,    "); 
		sqlCmd.append("                 sum(fieldI_XD2_L)    fieldI_XD2_L,    sum(fieldI_XD2)      fieldI_XD2,    "); 
		sqlCmd.append("                 sum(fieldI_XE1_L)    fieldI_XE1_L,    sum(fieldI_XE1)      fieldI_XE1,    "); 
		sqlCmd.append("                 sum(fieldI_XE2_L)    fieldI_XE2_L,    sum(fieldI_XE2)      fieldI_XE2,    "); 
		sqlCmd.append("                 sum(fieldI_XF1_L)    fieldI_XF1_L,    sum(fieldI_XF1)      fieldI_XF1,    "); 
		sqlCmd.append("                 SUM(fieldI_XF3)     fieldI_XF3,     sum(fieldI_XF2_L) fieldI_XF2_L,       ");
		sqlCmd.append("                 sum(fieldI_XF2)        fieldI_XF2, ");    
		sqlCmd.append("                 sum(fieldI_Y_L)      fieldI_Y_L,    sum(fieldI_Y)        fieldI_Y  ");   
		sqlCmd.append("           from(  select a01.hsien_id, a01.hsien_name, a01.FR001W_output_order, a01.bank_no, a01.BANK_NAME, "); 
		sqlCmd.append("                     sum(field_120700_L)  field_120700_L,    sum(field_120700)    field_120700,  sum(field_310000_L)  field_310000_L,    sum(field_310000) ");    
		sqlCmd.append("                     field_310000,  sum(field_320000_L)  field_320000_L,    sum(field_320000)    field_320000,  sum(field_220900_L)  field_220900_L,    sum(field_220900)    field_220900, ");  
		sqlCmd.append("                     sum(field_OVER_L)    field_OVER_L,      sum(field_OVER)      field_OVER,    sum(field_DEBIT_L)   field_DEBIT_L,     sum(field_DEBIT)     field_DEBIT,      ");
		sqlCmd.append("                     sum(field_CREDIT_L)  field_CREDIT_L,    sum(field_CREDIT)    field_CREDIT,  sum(fieldI_XA_L)       fieldI_XA_L,        sum(fieldI_XA)        fieldI_XA,    ");   
		sqlCmd.append("                     sum(fieldI_XB1_L)      fieldI_XB1_L,        sum(fieldI_XB1)     fieldI_XB1,    sum(fieldI_XB2_L)      fieldI_XB2_L,        sum(fieldI_XB2)     fieldI_XB2, ");    
		sqlCmd.append("                     sum(fieldI_XC1_L)      fieldI_XC1_L,sum(fieldI_XC1)     fieldI_XC1,    sum(fieldI_XC2_L)      fieldI_XC2_L,        sum(fieldI_XC2)     fieldI_XC2,   ");  
		sqlCmd.append("                     sum(fieldI_XD1_L)      fieldI_XD1_L,    sum(fieldI_XD1)     fieldI_XD1,sum(fieldI_XD2_L)        fieldI_XD2_L,    sum(fieldI_XD2)        fieldI_XD2,  ");   
		sqlCmd.append("                     sum(fieldI_XE1_L)      fieldI_XE1_L,        sum(fieldI_XE1)     fieldI_XE1,    sum(fieldI_XE2_L)      fieldI_XE2_L,        sum(fieldI_XE2)     fieldI_XE2, ");    
		sqlCmd.append("                     sum(fieldI_XF1_L)      fieldI_XF1_L,        sum(fieldI_XF1)     fieldI_XF1,    sum(fieldI_XF2_L)      fieldI_XF2_L,        sum(fieldI_XF2)     fieldI_XF2, ");    
		sqlCmd.append("                     SUM(fieldI_XF3)      fieldI_XF3,  sum(fieldI_Y_L)        fieldI_Y_L,        sum(fieldI_Y)        fieldI_Y  ");
		sqlCmd.append("                  from ( select a01.hsien_id, a01.hsien_name, a01.FR001W_output_order, a01.bank_no, a01.BANK_NAME, "); 
		sqlCmd.append("                          0 as field_120700_L, field_120700, 0 as field_310000_L,     field_310000,  0 as field_320000_L, field_320000,    ");
		sqlCmd.append("                          0 as  field_220900_L,    field_220900,  0 as field_OVER_L,   field_OVER,   0 as field_DEBIT_L,      field_DEBIT, ");   
		sqlCmd.append("                          0 as field_CREDIT_L, field_CREDIT, 0 as fieldI_XA_L,        fieldI_XA,     0 as fieldI_XB1_L,   fieldI_XB1,      ");
		sqlCmd.append("                          0 as fieldI_XB2_L,       fieldI_XB2,    0 as fieldI_XC1_L,   fieldI_XC1,   0 as fieldI_XC2_L,       fieldI_XC2,  ");   
		sqlCmd.append("                          0 as fieldI_XD1_L,   fieldI_XD1,   0 as fieldI_XD2_L,       fieldI_XD2,    0 as fieldI_XE1_L,   fieldI_XE1,      ");
		sqlCmd.append("                          0 as fieldI_XE2_L,       fieldI_XE2,    0 as fieldI_XF1_L,   fieldI_XF1,   0 as fieldI_XF2_L,       fieldI_XF2,  "); 
		sqlCmd.append("                          fieldI_XF3, 0 as fieldI_Y_L,     fieldI_Y  ");
		sqlCmd.append("                        from (  select nvl(cd01.hsien_id,' ') as hsien_id,nvl(cd01.hsien_name,'OTHER') as hsien_name,  cd01.FR001W_output_order as FR001W_output_order,bn01.bank_no,bn01.BANK_NAME, ");  
		sqlCmd.append("                                  round(sum(decode(a01.acc_code,'120700',amt,0)) /1,0) as field_120700, ");  
		sqlCmd.append("                                  round(sum(decode(a01.acc_code,'310000',amt,0)) /1,0) as field_310000, ");  
		sqlCmd.append("                                  round(sum(decode(a01.acc_code,'320000',amt,0)) /1,0) as field_320000, ");  
		sqlCmd.append("                                  round(sum(decode(a01.acc_code,'220900',amt,0)) /1,0) as field_220900, ");  
		sqlCmd.append("                                  round(sum(decode(a01.acc_code,'990000',amt,0)) /1,0) as field_OVER,   ");  
		sqlCmd.append("                                  round(sum(decode(a01.acc_code,'220000',amt,0)) /1,0) as field_DEBIT,  ");  
		sqlCmd.append("                                  round(sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /1,0) as field_CREDIT, ");  
		sqlCmd.append("                                  decode(YEAR_TYPE,'102',decode(bank_type,'6',round(sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120200',amt,'120301',amt,'120302',amt,'120700',amt,'150200',amt,0)) /1,0),   ");                
		sqlCmd.append("                                                                           '7',round(sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120300',amt,'120401',amt,'120402',amt,'120700',amt,'150200',amt,0)) /1,0)), ");
		sqlCmd.append("                                          '103',round(sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120200',amt,'120301',amt,'120302',amt,'120700',amt,'150200',amt,0)) /1,0),0) as fieldI_XA, ");  
		sqlCmd.append("                                  decode(YEAR_TYPE,'102',decode(bank_type,'6',round(sum(decode(a01.acc_code,'120401',amt,'120402',amt,0)) /1,0),   ");                 
		sqlCmd.append("                                                                           '7',round(sum(decode(a01.acc_code,'120201',amt,'120202',amt,0)) /1,0)), ");
		sqlCmd.append("                                                   '103',round(sum(decode(a01.acc_code,'120401',amt,'120402',amt,0)) /1,0),0) as fieldI_XB1,       ");
		sqlCmd.append("                                  decode(YEAR_TYPE,'102',decode(bank_type,'6',sum(decode(a01.acc_code,'240305',amt,'251300',amt,'310800',amt,0)),  ");                 
		sqlCmd.append("                                                                           '7',sum(decode(a01.acc_code,'240205',amt,'310800',amt,0))),             ");
		sqlCmd.append("                                                   '103',decode(bank_type,'6',sum(decode(a01.acc_code,'240305',amt,'251300',amt,'310800',amt,0)),  ");                 
		sqlCmd.append("                                                                           '7',sum(decode(a01.acc_code,'240305',amt,'251200',amt,0))),0)  as fieldI_XB2, ");  
		sqlCmd.append("                                  round(sum(decode(a01.acc_code,'120501',amt,'120502',amt,0)) /1,0)    as fieldI_XC1, ");      
		sqlCmd.append("                                  decode(YEAR_TYPE,'102',decode(bank_type,'6',round(sum(decode(a01.acc_code,'240301',amt,'240302',amt,'240303',amt,'240304',amt,0)) /1,0),    ");
		sqlCmd.append("                                                                            '7',round(sum(decode(a01.acc_code,'240201',amt,'240202',amt,'240203',amt,'240204',amt,0)) /1,0)), ");
		sqlCmd.append("                                                    '103',round(sum(decode(a01.acc_code,'240301',amt,'240302',amt,'240303',amt,'240304',amt,0)) /1,0),0) as fieldI_XC2,       ");
		sqlCmd.append("                                  round(sum(decode(a01.acc_code,'120600',amt,0)) /1,0)                  as fieldI_XD1,  ");
		sqlCmd.append("                                  decode(YEAR_TYPE,'102',decode(bank_type,'6',round(sum(decode(a01.acc_code,'240200',amt,0)) /1,0),  ");
		sqlCmd.append("                                                                          '7',round(sum(decode(a01.acc_code,'240300',amt,0)) /1,0)), ");
		sqlCmd.append("                                                   '103',round(sum(decode(a01.acc_code,'240200',amt,0)) /1,0),0)  as fieldI_XD2,     ");
		sqlCmd.append("                                  round(sum(decode(a01.acc_code,'150100',amt,0)) /1,0)                  as fieldI_XE1, "); 
		sqlCmd.append("                                  round(sum(decode(a01.acc_code,'250100',amt,0)) /1,0)                  as fieldI_XE2, "); 
		sqlCmd.append("                                  round(sum(decode(a01.acc_code,'310000',amt,'320000',amt,0)) /1,0)  as fieldI_XF1,    ");
		sqlCmd.append("                                  decode(YEAR_TYPE,'102',round(sum(decode(a01.acc_code,'310800',amt,0)) /1,0),  ");
		sqlCmd.append("                                                   '103',decode(bank_type,'6',round(sum(decode(a01.acc_code,'310800',amt,0)) /1,0),7,0),0)    as fieldI_XF3, "); 
		sqlCmd.append("                                  round(sum(decode(a01.acc_code,'140000',amt,0)) /1,0)                  as fieldI_XF2, "); 
		sqlCmd.append("                                             round((sum(decode(a01.acc_code,'220100',amt,'220200',amt,'220300',amt,'220400',amt,  '220500',amt,'220600',amt,'220700',amt,'220800',amt,'220900',amt,'221000',amt,0))  ");
		sqlCmd.append("                                             -  round(sum(decode(a01.acc_code,'220900',amt,0))/2,0)) /1,0)   as fieldI_Y  "); 
		sqlCmd.append("                                from (select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id = ? )   cd01  ");
		paramList.add(HSIEN_ID);
		sqlCmd.append("                                left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id  "); 
		sqlCmd.append("                                left join (select * from bn01 where m_year=?)bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=?  ");
		paramList.add(wlx01_m_year);
		paramList.add(wlx01_m_year);
		paramList.add(bank_type);
		sqlCmd.append("                                left join (select  (CASE WHEN (a01.m_year <= 102) THEN '102'  ");
		sqlCmd.append("                                                         WHEN (a01.m_year > 102) THEN '103'   ");
		sqlCmd.append("                                                    ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01)a01 on a01.m_year = ? and a01.m_month = ? and a01.bank_code=bn01.bank_no  ");
		paramList.add(s_year);
        paramList.add(s_month);
		sqlCmd.append("                                group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,  a01.YEAR_TYPE, bn01.bank_type,bn01.bank_no ,  bn01.BANK_NAME ");
		sqlCmd.append("                            ) a01  ");  
		sqlCmd.append("                     where  a01.bank_no <> ' '  ");
		sqlCmd.append("                     union all   ");
		sqlCmd.append("                     select a01.hsien_id, a01.hsien_name, a01.FR001W_output_order, a01.bank_no, a01.BANK_NAME, "); 
		sqlCmd.append("                     field_120700 as field_120700_L, 0 as field_120700, field_310000 as field_310000_L, 0 as field_310000, ");  
		sqlCmd.append("                     field_320000 as field_320000_L, 0 as field_320000, field_220900 as field_220900_L, 0 as field_220900, ");  
		sqlCmd.append("                     field_OVER   as field_OVER_L,   0 as field_OVER,   field_DEBIT  as field_DEBIT_L,  0 as field_DEBIT,  ");  
		sqlCmd.append("                     field_CREDIT as field_CREDIT_L, 0 as field_CREDIT, fieldI_XA    as fieldI_XA_L,    0 as fieldI_XA,    ");  
		sqlCmd.append("                     fieldI_XB1   as fieldI_XB1_L,   0 as fieldI_XB1,   fieldI_XB2   as fieldI_XB2_L,   0 as fieldI_XB2,   ");  
		sqlCmd.append("                     fieldI_XC1   as fieldI_XC1_L,   0 as fieldI_XC1,   fieldI_XC2   as fieldI_XC2_L,   0 as fieldI_XC2,   ");  
		sqlCmd.append("                     fieldI_XD1   as fieldI_XD1_L,   0 as fieldI_XD1,   fieldI_XD2   as fieldI_XD2_L,   0 as fieldI_XD2,   ");  
		sqlCmd.append("                     fieldI_XE1   as fieldI_XE1_L,   0 as fieldI_XE1,   fieldI_XE2   as fieldI_XE2_L,   0 as fieldI_XE2,   ");  
		sqlCmd.append("                     fieldI_XF1   as fieldI_XF1_L,   0 as fieldI_XF1,   fieldI_XF2   as fieldI_XF2_L,   0 as fieldI_XF2,   ");  
		sqlCmd.append("                     0 as fieldI_XF3, fieldI_Y       as fieldI_Y_L,     0 as fieldI_Y "); 
		sqlCmd.append("                     from (  select nvl(cd01.hsien_id,' ') as hsien_id,nvl(cd01.hsien_name,'OTHER') as hsien_name,  cd01.FR001W_output_order as FR001W_output_order,bn01.bank_no,bn01.BANK_NAME, ");  
		sqlCmd.append("                               round(sum(decode(a01.acc_code,'120700',amt,0)) /1,0) as field_120700, ");  
		sqlCmd.append("                               round(sum(decode(a01.acc_code,'310000',amt,0)) /1,0) as field_310000, ");  
		sqlCmd.append("                               round(sum(decode(a01.acc_code,'320000',amt,0)) /1,0) as field_320000, ");  
		sqlCmd.append("                               round(sum(decode(a01.acc_code,'220900',amt,0)) /1,0) as field_220900, ");  
		sqlCmd.append("                               round(sum(decode(a01.acc_code,'990000',amt,0)) /1,0) as field_OVER,   ");  
		sqlCmd.append("                               round(sum(decode(a01.acc_code,'220000',amt,0)) /1,0) as field_DEBIT,  ");  
		sqlCmd.append("                               round(sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /1,0) as field_CREDIT, ");  
		sqlCmd.append("                               decode(YEAR_TYPE,'102',decode(bank_type,'6',round(sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120200',amt,'120301',amt,'120302',amt,'120700',amt,'150200',amt,0)) /1,0),   ");                
		sqlCmd.append("                                                                        '7',round(sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120300',amt,'120401',amt,'120402',amt,'120700',amt,'150200',amt,0)) /1,0)), ");
		sqlCmd.append("                                                '103',round(sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120200',amt,'120301',amt,'120302',amt,'120700',amt,'150200',amt,0)) /1,0),0) as fieldI_XA,        ");
		sqlCmd.append("                               decode(YEAR_TYPE,'102',decode(bank_type,'6',round(sum(decode(a01.acc_code,'120401',amt,'120402',amt,0)) /1,0),   ");                 
		sqlCmd.append("                                                                        '7',round(sum(decode(a01.acc_code,'120201',amt,'120202',amt,0)) /1,0)), ");
		sqlCmd.append("                                                '103',round(sum(decode(a01.acc_code,'120401',amt,'120402',amt,0)) /1,0),0) as fieldI_XB1,       ");
		sqlCmd.append("                               decode(YEAR_TYPE,'102',decode(bank_type,'6',sum(decode(a01.acc_code,'240305',amt,'251300',amt,'310800',amt,0)),  ");                 
		sqlCmd.append("                                                                        '7',sum(decode(a01.acc_code,'240205',amt,'310800',amt,0))),             ");
		sqlCmd.append("                                                '103',decode(bank_type,'6',sum(decode(a01.acc_code,'240305',amt,'251300',amt,'310800',amt,0)),  ");                 
		sqlCmd.append("                                                                        '7',sum(decode(a01.acc_code,'240305',amt,'251200',amt,0))),0)  as fieldI_XB2, ");  
		sqlCmd.append("                               round(sum(decode(a01.acc_code,'120501',amt,'120502',amt,0)) /1,0)    as fieldI_XC1, ");      
		sqlCmd.append("                               decode(YEAR_TYPE,'102',decode(bank_type,'6',round(sum(decode(a01.acc_code,'240301',amt,'240302',amt,'240303',amt,'240304',amt,0)) /1,0),  "); 
		sqlCmd.append("                                                                         '7',round(sum(decode(a01.acc_code,'240201',amt,'240202',amt,'240203',amt,'240204',amt,0)) /1,0)), ");
		sqlCmd.append("                                                 '103',round(sum(decode(a01.acc_code,'240301',amt,'240302',amt,'240303',amt,'240304',amt,0)) /1,0),0) as fieldI_XC2, ");  
		sqlCmd.append("                               round(sum(decode(a01.acc_code,'120600',amt,0)) /1,0)                  as fieldI_XD1, "); 
		sqlCmd.append("                               decode(YEAR_TYPE,'102',decode(bank_type,'6',round(sum(decode(a01.acc_code,'240200',amt,0)) /1,0),  ");
		sqlCmd.append("                                                                       '7',round(sum(decode(a01.acc_code,'240300',amt,0)) /1,0)), ");
		sqlCmd.append("                                                '103',round(sum(decode(a01.acc_code,'240200',amt,0)) /1,0),0)  as fieldI_XD2,     ");
		sqlCmd.append("                               round(sum(decode(a01.acc_code,'150100',amt,0)) /1,0)                  as fieldI_XE1, "); 
		sqlCmd.append("                               round(sum(decode(a01.acc_code,'250100',amt,0)) /1,0)                  as fieldI_XE2, "); 
		sqlCmd.append("                               round(sum(decode(a01.acc_code,'310000',amt,'320000',amt,0)) /1,0)  as fieldI_XF1,    ");
		sqlCmd.append("                               decode(YEAR_TYPE,'102',round(sum(decode(a01.acc_code,'310800',amt,0)) /1,0),         ");
		sqlCmd.append("                                                '103',decode(bank_type,'6',round(sum(decode(a01.acc_code,'310800',amt,0)) /1,0),7,0),0)    as fieldI_XF3, "); 
		sqlCmd.append("                               round(sum(decode(a01.acc_code,'140000',amt,0)) /1,0)                  as fieldI_XF2, "); 
		sqlCmd.append("                               round((sum(decode(a01.acc_code,'220100',amt,'220200',amt,'220300',amt,'220400',amt,  '220500',amt,'220600',amt,'220700',amt,'220800',amt,'220900',amt,'221000',amt,0)) "); 
		sqlCmd.append("                               -  round(sum(decode(a01.acc_code,'220900',amt,0))/2,0)) /1,0)   as fieldI_Y  "); 
		sqlCmd.append("                           from (select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id = ? )   cd01 ");
		paramList.add(HSIEN_ID);
		sqlCmd.append("                           left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id ");  
		sqlCmd.append("                           left join (select * from bn01 where m_year=?)bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=? "); 
		paramList.add(wlx01_m_year);
		paramList.add(wlx01_m_year);
		paramList.add(bank_type);
		sqlCmd.append("                           left join (select  (CASE WHEN (a01.m_year <= 102) THEN '102' ");
		sqlCmd.append("                                                   WHEN (a01.m_year > 102) THEN '103'   ");
		sqlCmd.append("                                              ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01)a01 on a01.m_year = ? and a01.m_month = ? and a01.bank_code=bn01.bank_no ");
		paramList.add(l_year);
        paramList.add(l_month);
		sqlCmd.append("                           group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,  a01.YEAR_TYPE, bn01.bank_type,bn01.bank_no ,  bn01.BANK_NAME ");
		sqlCmd.append("                          ) a01 ");  
		sqlCmd.append("                   where  a01.bank_no <> ' ' ");  
		sqlCmd.append("                ) a01 ");  
		sqlCmd.append("                group by a01.hsien_id, a01.hsien_name, a01.FR001W_output_order, a01.bank_no, a01.BANK_NAME "); 
		sqlCmd.append("          ) a01 ");  
		sqlCmd.append("          group by a01.hsien_id ,  a01.hsien_name "); 
		sqlCmd.append("       ) a01 "); //--合計資料
		sqlCmd.append(" ) a01 "); 
		sqlCmd.append(" order by a01.hsien_id, a01.FR001W_output_order DESC ");

		//查詢
		dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,
				"field_debit,field_debit_l,field_debit_lc,field_debit_lc_rate,field_credit,field_credit_l,"+
				"field_credit_lc,field_credit_lc_rate,field_dc_rate,field_dc_rate_l,field_120700,field_over,"+
				"field_over_l,c_field_over_rate,c_field_over_rate_l,field_310000,field_320000,field_220900");
		System.out.println("dbData.size() = " + dbData.size());
		return dbData;
	}//end of getData
	                                                                                                         
}//end of class                                                                                                                       