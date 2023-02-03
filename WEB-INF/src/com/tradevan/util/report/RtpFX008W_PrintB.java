/*
 * Created on  95.1.4  by 4183 lilic0c0
 * 97.11.13 fix 是否提足備抵跌價損失/是否有積極處分之事實/未來處分計劃是否合理可行/審核結果.為"N"->否."N"->不合格 by 2295
 * 99.12.06 fix sqlInjection by 2808
 */
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.*;
import java.util.*;
import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RtpFX008W_PrintB{
	
	/* ===============================================
	 * create 申請延長處分期限審核表 report
	 * ===============================================
	 */
	public static String createRpt(String s_year,String s_month,String unit,String title_bank_name,String bank_type,String m2_name){
		
		String errMsg = "";
		String bank_type_name =( bank_type.equals("6") )?"農會":"漁會"; 
		String unit_name = getUnitName(unit); 		//取得顯示的金額單位
		 
      	try{
      		
      		//生出報表及設定格式===========================  
      		System.out.println("FR035W_Excel.createRpt Starting........");
      	
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
						
			String openfile="縣市政府承受擔保品申請延長處分期限審核表.xls";//要去開啟的範本檔
						
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
			wb.setRepeatingRowsAndColumns(0, 1, 17, 2, 3);
			
			finput.close();
			
			HSSFRow row=null;//宣告一列 
			HSSFCell cell=null;//宣告一個儲存格

			//建表開始 ===================================
			
			int rowNum =4;//表頭有4列是已經有且固定的資料，所以row從第5(index為4)列開始生
			int columnNum =12;//本excel總共有12行
			
			//取得申報表資料
      		List dbData = getData(s_year,s_month,unit,m2_name);
      		DataObject bean;
      		System.out.println("dbData.size="+dbData.size());
      		//first row
	  		row = sheet.getRow(0);                                                                        
	   		cell = row.getCell((short)0);
	   		cell.setEncoding(HSSFCell.ENCODING_UTF_16);//顯示中文字
	   		cell.setCellValue(title_bank_name+"政府承受擔保品申請延長處分期限審核表");
				System.out.println(title_bank_name+"政府承受擔保品申請延長處分期限審核表");
      
      		if(dbData == null){
      			System.out.println("dbData is null !!");
      			
      			//second row                                    
				row = sheet.getRow(1);                          
				cell = row.getCell((short)2);                  
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);     
				cell.setCellValue(s_year+" 年第"+s_month+" 季沒有資料" );
				
			} else {
	 			//second row                                    
				row = sheet.getRow(1);                          
				cell = row.getCell((short)2);                  
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);     
				cell.setCellValue(s_year+" 年第"+s_month+" 季" );
				
				
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
				int i = 0;	
				for(i=0;i<dbData.size();i++){
					
					row = sheet.createRow(rowNum+i);
					
					for(int j=0;j<22;j++){
						cell = row.createCell( (short) j);
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					}
					//取得資料
					bean = (DataObject) dbData.get(i);
					
					String flag =	bean.getValue("auditresult_yn").toString();					
					if(flag.equals("Y")){
					//塞資料
					row = sheet.getRow(rowNum+i);
									
					//信用部名稱
					if(!lastBankName.equals(bean.getValue("bank_name").toString())){	
						cell = row.getCell((short)0);
						cell.setCellValue(bean.getValue("bank_name").toString());
						cell.setCellStyle(cs);
						lastBankName = bean.getValue("bank_name").toString();
					}
					else{
						cell = row.getCell((short)0);
						cell.setCellValue(" ");
						cell.setCellStyle(cs);
					}
					
					//承受擔保編號
					cell = row.getCell((short)1);
					cell.setCellValue(bean.getValue("dureassure_no").toString());
					cell.setCellStyle(cs); 
				
					//借款人
					cell = row.getCell((short)2);
					cell.setCellValue(bean.getValue("debtname").toString());
					cell.setCellStyle(cs); 
					
					//承受日期
					
					cell = row.getCell((short)3);
					cell.setCellValue(bean.getValue("duredate").toString());
					cell.setCellStyle(cs);
								
					//承受擔保品座落
					cell = row.getCell((short)4);
					cell.setCellValue(bean.getValue("dureassuresite").toString());
					cell.setCellStyle(cs);
					
					//帳列金額
					cell = row.getCell((short)5);
					cell.setCellValue(Utility.setCommaFormat(bean.getValue("accountamt").toString()));
					cell.setCellStyle(csAccont);
					
					//申請延長期間
					cell = row.getCell((short)6);
					cell.setCellValue(bean.getValue("applydelayyear_month").toString());
					cs.setAlignment(HSSFCellStyle.ALIGN_LEFT);
					cell.setCellStyle(cs);
		
					//申請延長理由
					cell = row.getCell((short)7);
					cell.setCellValue(bean.getValue("applydelayreason").toString());
					cell.setCellStyle(cs);
							
				 
					//是否提足備抵跌價損失
					cell = row.getCell((short)10);
					String temp_damage_yn = bean.getValue("damage_yn").toString();
					String damage_yn="";
					if(temp_damage_yn.equals("Y"))
							damage_yn="是";
					if(temp_damage_yn.equals("N"))
							damage_yn="否";
					cell.setCellValue(damage_yn);
					cell.setCellStyle(csYN);
					
					//是否有積極處分之事實
					cell = row.getCell((short)11);					
					String temp_disposal_fact_yn = bean.getValue("disposal_fact_yn").toString();
					String disposal_fact_yn="";
					if(temp_disposal_fact_yn.equals("Y"))
							disposal_fact_yn="是";
					if(temp_disposal_fact_yn.equals("N"))
							disposal_fact_yn="否";
					cell.setCellValue(disposal_fact_yn);
					cell.setCellStyle(csYN);
					
					//未來處分計劃是否合理可行
					cell = row.getCell((short)12);
					String temp_disposal_plan_yn = bean.getValue("disposal_plan_yn").toString();
					String disposal_plan_yn="";
					if(temp_disposal_plan_yn.equals("Y"))
							disposal_plan_yn="是";
					if(temp_disposal_plan_yn.equals("N"))
							disposal_plan_yn="否";
					cell.setCellValue(disposal_plan_yn);
					cell.setCellStyle(csYN);
					
					//審核結果
					cell = row.getCell((short)13);
					String temp_auditresult_yn = bean.getValue("auditresult_yn").toString();
					String auditresult_yn="";
					if(temp_auditresult_yn.equals("Y"))
									auditresult_yn="合格";
					if(temp_auditresult_yn.equals("N"))
									auditresult_yn="不合格";
					cell.setCellValue(auditresult_yn);
					cell.setCellStyle(csYN);					
					
					//延長期間
					cell = row.getCell((short)18);
					System.out.println("audit_applydelayyear_month:"+bean.getValue("audit_applydelayyear_month").toString());
					cell.setCellValue(bean.getValue("audit_applydelayyear_month").toString());
					//String temp="否";
					//cell.setCellValue(temp);
					cs.setAlignment(HSSFCellStyle.ALIGN_LEFT);
					cell.setCellStyle(csYN);
					
					//延長期限
					cell = row.getCell((short)19);	
					System.out.println("audit_duredate:"+bean.getValue("audit_duredate").toString());
					cell.setCellValue(bean.getValue("audit_duredate").toString());
					//String ttemp="否";
					//cell.setCellValue(ttemp);
					cs.setAlignment(HSSFCellStyle.ALIGN_LEFT);
					cell.setCellStyle(csYN);
					
				}//end of inner-for
				
				}
				
			}//end of outer-for
			
      		//表尾
			rowNum=rowNum+dbData.size();
			row = sheet.createRow(rowNum+1);		
			for(int j=0;j<columnNum;j++){
			cell = row.createCell( (short) j);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			}
			
			cell = row.getCell((short)0);
			cell.setCellValue("局長");
			cell = row.getCell((short)1);
			cell.setCellValue("副局長");
			cell = row.getCell((short)4);
			cell.setCellValue("科長(技正)");
			cell = row.getCell((short)6);
			cell.setCellValue("課(股)長");
			cell = row.getCell((short)11);
			cell.setCellValue("承辦人");
      		
			//建表結束 =================================
			HSSFFooter footer = sheet.getFooter();
			footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
			
			FileOutputStream fout=new FileOutputStream(reportDir + System.getProperty("file.separator")+"縣市政府承受擔保品申請延長處分期限審核表.xls");  
			
			wb.write(fout);//儲存
			fout.close();
			System.out.println("儲存完成");
			
			
		}catch(Exception e){
			System.out.println("createRpt Error:"+e+e.getMessage());                                                
		}
		
		System.out.println("FR035W_Excel.createRpt Ending........");
		
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
	
	/*  ===============================================
	 *	取得申請延長處分期限審核表的各項資料
	 *  ===============================================
	 *  added 縣市合併 by 2808 99.12.08
	 */
	public static List getData(String s_year,String s_month,String unit,String m2_name){
		String m_year= "99" ;
		String cd01Table = "cd01_99" ;
		if(Integer.parseInt(s_year) > 99) {
			m_year = "100" ;
			cd01Table = "cd01" ;
		}
		List dbData = null;
		String sqlCmd = "";
		List paramList =new ArrayList() ;
	sqlCmd =" select * from (select nvl(cd01.hsien_id,' ') as  hsien_id , "+
					" nvl(cd01.hsien_name,'OTHER')  as  hsien_name, "+  
				  " cd01.FR001W_output_order     as  FR001W_output_order, "+  
				  " bn01.bank_no ,  bn01.BANK_NAME, nvl(DureAssure_NO,0) as DureAssure_NO,  "+ 
				  " nvl(DebtName, ' ')    as DebtName, "+  
				  " nvl(to_char(DureDate,'yyyy')-1911||to_char(DureDate,'/mm/dd'),'') as DureDate,  "+ 
				  " nvl(DureAssureSite, ' ') as DureAssureSite , "+  
				  " Round(AccountAmt/1,0)  as AccountAmt, "+  
				  " (decode(ApplyDelayYear,0,' ', Ltrim(ApplyDelayYear)|| '年') ||   "+
				  " decode(ApplyDelayMonth,0,' ',Ltrim(ApplyDelayMonth)|| '個月')) as ApplyDelayYear_MONTH,   "+
				  " ApplyDelayReason,(decode(nvl(Audit_ApplyDelayYear,0),0,' ', Ltrim(Audit_ApplyDelayYear)|| '年') ||   "+
				  " decode(nvl(Audit_ApplyDelayMonth,0),0,' ',Ltrim(Audit_ApplyDelayMonth)|| '個月')) as Audit_ApplyDelayYear_MONTH, "+  
				  " nvl(to_char(Audit_DureDate,'yyyy')-1911||to_char(Audit_DureDate,'/mm/dd'),' ') as Audit_DureDate,   "+
				  " nvl(Damage_YN, ' ') as Damage_YN, nvl(Disposal_Fact_YN, ' ') as  Disposal_Fact_YN,   "+
				  " nvl(Disposal_Plan_YN, ' ') as  Disposal_Plan_YN, nvl(AuditResult_YN, ' ') as  AuditResult_YN,   "+
				  " nvl(ApplyOK_DocNo, ' ') as  ApplyOK_DocNo, "+  
				  " nvl(to_char(ApplyOK_Date,'yyyy')-1911||to_char(ApplyOK_Date,'/mm/dd'),' ') as ApplyOK_Date,   "+	   
				  " nvl(Report_BOAF_DocNo, ' ') as  Report_BOAF_DocNo, nvl(to_char(Report_BOAF_Date,'yyyy')-1911||   "+
				  " to_char(Report_BOAF_Date,'/mm/dd'),' ') as Report_BOAF_Date from  (   "+
				  " select * from "+cd01Table+" where hsien_id <> 'Y') cd01 left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id    "+
				  " left join (select * from bn01 where m_year=? )bn01 on wlx01.bank_no=bn01.bank_no  "+ 
				  " and WLX01.m2_name =?"+
				  " and bn01.bank_type in ('6','7')  "+
				  " left join   WLX08_S_GAGE   a01  on  bn01.bank_no = a01.bank_no  "+ 
				  " and a01.M_YEAR =   ? AND  a01.M_Quarter  =   ? "+ 
				  " ) Temp1  WHERE  BANK_NO  <> ' ' and  DebtName <> ' '  order by FR001W_output_order, hsien_id, bank_no, DureAssure_NO ";
		paramList.add(m_year) ;
		paramList.add(m_year) ;
		paramList.add(m2_name) ;
		paramList.add(s_year) ;
		paramList.add(s_month) ;
		//查詢
		dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"hsien_id,hsien_name,fr001w_output_order,bank_no,bank_name,dureassure_no,debtname,duredate,dureassuresite,"+
										  "accountamt,applydelayyear_month,applydelayreason,audit_applydelayyear_month,audit_duredate,damage_yn,"+
 										  "disposal_fact_yn,disposal_plan_yn,auditresult_yn,applyok_docno,applyok_date,report_boaf_docno,report_boaf_date");
		
		return dbData;
	}//end of getData
	                                                                                                         
}//end of class                                                                                                                       