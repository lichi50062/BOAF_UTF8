/*
 * Created on  95.1.4  by 4183 lilic0c0
 * 97.07.22 fix 更改列印日期為當天日期 by 2295
 */
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;

import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import java.io.*;
import com.tradevan.util.Utility;

public class RtpFX008W_Print{
	
	/* ===============================================
	 * create 申請延長處分期限審核表 report
	 * ===============================================
	 */
	public static String createRpt(String bank_name,String dureassure_no,String debtname,String year,String month,String day,String dureassuresite,String accountamt,String applydelayyear_month,String applydelayreason,String damage_yn,String disposal_fact_yn,String disposal_plan_yn){
		
		String errMsg = "";
		String unit = "1";
		String unit_name = getUnitName(unit); 		//取得顯示的金額單位
      	
      	try{
      		
      		//生出報表及設定格式===========================  
      		System.out.println("FR008W_Print.createRpt Starting........");
      	
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
						
			String openfile="信用部承受擔保品申請延長處分期限審核表.xls";//要去開啟的範本檔
						
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
			//sheet.setZoom(80, 100); 					//螢幕上看到的縮放大小
			//sheet.setAutobreaks(true); 				//自動分頁
			
			//設定頁面符合列印大小
			sheet.setAutobreaks( false );
			ps.setScale( ( short )84 ); 				//列印縮放百分比
			
			ps.setPaperSize( ( short )9 ); 				//設定紙張大小 A4
			
			//設定表頭 為固定 先設欄的起始再設列的起始
			//wb.setRepeatingRowsAndColumns(0, 1, 17, 2, 3);
			
			finput.close();
			
			HSSFRow row=null;//宣告一列 
			HSSFCell cell=null;//宣告一個儲存格

			//建表開始 ===================================
			
			int rowNum =4;//表頭有4列是已經有且固定的資料，所以row從第5(index為4)列開始生
			int columnNum =11;//本excel總共有11行
			
			//取得申報表資料
      		//List dbData = getData(s_year,s_month,unit,bank_type);
      		//DataObject bean;
      		
      		//first row
	  		row = sheet.getRow(0);                                                                        
	   		cell = row.getCell((short)0);
	   		cell.setEncoding(HSSFCell.ENCODING_UTF_16);//顯示中文字
	   		cell.setCellValue(bank_name+"承受擔保品申請延長處分期限審核表");
				System.out.println(bank_name+"承受擔保品申請延長處分期限審核表");
      	
	 			//second row                                    
				row = sheet.getRow(1);                          
				cell = row.getCell((short)1);                  
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				/*
				String temp_month=month.substring(1,2);    
				cell.setCellValue(year+ "年"+temp_month+"月"+day +"日" );
				*/
				//97.07.22 fix 更改列印日期為今日
				cell.setCellValue(Utility.getCHTdate(Utility.getDateFormat("yyyy/MM/dd"),1));
				cell = row.getCell((short)9);                  
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);     
				cell.setCellValue("單位:新臺幣 " + unit_name );  
				
				//create all new row and cell for stroe table data 
				HSSFCellStyle cs = wb.createCellStyle();
				HSSFFont ft = wb.createFont();
				ft.setFontHeightInPoints((short)10);
				
				HSSFCellStyle csAccont = wb.createCellStyle();
				HSSFCellStyle csYN = wb.createCellStyle();
					 			
				//set format for cell
	 			cs.setFont(ft);                                
				cs.setAlignment(HSSFCellStyle.ALIGN_LEFT);
				cs.setBorderTop(HSSFCellStyle.BORDER_THIN);
				cs.setBorderBottom(HSSFCellStyle.BORDER_THIN);
				cs.setBorderLeft(HSSFCellStyle.BORDER_THIN);
				cs.setBorderRight(HSSFCellStyle.BORDER_THIN);
				cs.setWrapText(true);
				csAccont.setFont(ft);
				csAccont.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
				csAccont.setBorderTop(HSSFCellStyle.BORDER_THIN);
				csAccont.setBorderBottom(HSSFCellStyle.BORDER_THIN);
				csAccont.setBorderLeft(HSSFCellStyle.BORDER_THIN);
				csAccont.setBorderRight(HSSFCellStyle.BORDER_THIN);
				csAccont.setDataFormat((short)3);
				csYN.setFont(ft);
				csYN.setAlignment(HSSFCellStyle.ALIGN_CENTER);
				csYN.setBorderTop(HSSFCellStyle.BORDER_THIN);
				csYN.setBorderBottom(HSSFCellStyle.BORDER_THIN);
				csYN.setBorderLeft(HSSFCellStyle.BORDER_THIN);
				csYN.setBorderRight(HSSFCellStyle.BORDER_THIN);
				
				//進入for-loop
				String lastBankName =" ";//紀錄上一個的農漁會名稱(用來判別是否要印出名稱用)
							
				int i=0;
					
					row = sheet.createRow(rowNum+i);
					
					for(int j=0;j<columnNum;j++){
						cell = row.createCell( (short) j);
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					}
					//取得資料
					//bean = (DataObject) dbData.get(i);
					
					//塞資料
					row = sheet.getRow(rowNum+i);
									
					//信用部名稱		
					cell = row.getCell((short)0);			
					cell.setCellValue(bank_name);
					cell.setCellStyle(cs); 
					//承受擔保編號
					cell = row.getCell((short)1);
					cell.setCellValue(dureassure_no);
					cell.setCellStyle(cs); 
				
					//借款人
					cell = row.getCell((short)2);
					cell.setCellValue(debtname);
					cell.setCellStyle(cs); 
					
					//承受日期
					
					cell = row.getCell((short)3);
					cell.setCellValue(year+month+day);
					cell.setCellStyle(cs);
								
					//承受擔保品座落
					cell = row.getCell((short)4);
					cell.setCellValue(dureassuresite);
					cell.setCellStyle(cs);
					
					//申請延長期間
					cell = row.getCell((short)5);
					cell.setCellValue(applydelayyear_month);
					cs.setAlignment(HSSFCellStyle.ALIGN_LEFT);
					cell.setCellStyle(cs);
					
					//帳列金額
					cell = row.getCell((short)6);
					cell.setCellValue(accountamt);
					cell.setCellStyle(csAccont);
								
		
					//申請延長理由
					cell = row.getCell((short)7);
					cell.setCellValue(applydelayreason);
					cell.setCellStyle(cs);
			
					
				
					//是否提足備抵跌價損失
					/*
					cell = row.getCell((short)8);
					cell.setCellValue(damage_yn);
					cell.setCellStyle(csYN);
					
					//是否有積極處分之事實
					cell = row.getCell((short)9);
					cell.setCellValue(disposal_fact_yn);
					cell.setCellStyle(csYN);
					
					//未來處分計劃是否合理可行
					cell = row.getCell((short)10);
					cell.setCellValue(disposal_plan_yn);
					cell.setCellStyle(csYN);
				*/
					
				 //是否提足備抵跌價損失
					cell = row.getCell((short)8);
					String temp_damage_yn = damage_yn;
					
					if(temp_damage_yn.equals("Y"))
									damage_yn="是";
					if(temp_damage_yn.equals("N"))
									damage_yn="否";
					cell.setCellValue(damage_yn);
					cell.setCellStyle(csYN);
					
					//是否有積極處分之事實
					cell = row.getCell((short)9);
					String temp_disposal_fact_yn = disposal_fact_yn;
					
					if(temp_disposal_fact_yn.equals("Y"))
									disposal_fact_yn="是";
					if(temp_disposal_fact_yn.equals("N"))
									disposal_fact_yn="否";
					cell.setCellValue(disposal_fact_yn);
					cell.setCellStyle(csYN);
					
					//未來處分計劃是否合理可行
					cell = row.getCell((short)10);
					String temp_disposal_plan_yn = disposal_plan_yn;
					
					if(temp_disposal_plan_yn.equals("Y"))
									disposal_plan_yn="是";
					if(temp_disposal_plan_yn.equals("N"))
									disposal_plan_yn="否";
					cell.setCellValue(disposal_plan_yn);
					cell.setCellStyle(csYN);
			
			
			//建表結束 =================================
			HSSFFooter footer = sheet.getFooter();
			footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
			
			FileOutputStream fout=new FileOutputStream(reportDir + System.getProperty("file.separator")+"信用部承受擔保品申請延長處分期限審核表.xls");  
			
			wb.write(fout);//儲存
			fout.close();
			System.out.println("儲存完成");
			
			
		}catch(Exception e){
			System.out.println("createRpt Error:"+e+e.getMessage());                                                
		}
		
		System.out.println("FR008W_Print.createRpt Ending........");
		
		return errMsg;
	}//end of createreport
	
	/*  ===============================================
	 *	取得顯示的金額單位
	 *  ===============================================
 	 */ 
	public static String getUnitName(String unit){
		
		String unit_name ="";
		
		//設定顯示的金額單位
		if(unit.equals("1")){
			unit_name="元";
		}else if (unit.equals("1000")){
			unit_name="千元";
		}else if (unit.equals("10000")){
			unit_name="萬元";
		}else if (unit.equals("1000000")){
			unit_name="百萬元";
		}else if (unit.equals("10000000")){
			unit_name="千萬元";
		}else if (unit.equals("100000000")){
			unit_name ="億元";
		}
		
		return unit_name;
	}//end of getUnitName
	
	                                                                                                         
}//end of class                                                                                                                       