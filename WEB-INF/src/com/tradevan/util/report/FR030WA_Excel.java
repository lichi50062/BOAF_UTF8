/*
 * Created on  95.1.6  by 4183 lilic0c0
 * 99.04.12 fix 因應縣市合併修改SQL &Fix sql以preparedstatement方式查詢 by 2808
 * 99.11.09 fix 單位排序 by 2808
 * 
 */
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import java.io.*;
import java.util.*;
import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class FR030WA_Excel {
	
	/* ===============================================
	 * create 支票存款戶數與餘額明細 report
	 * ===============================================
	 */
	public static String createRpt(String s_year,String s_month,String unit,String bank_type,List bank_list){
		int u_year = s_year==null? 99 :Integer.parseInt(s_year)  ; //判斷縣市合併用參數
		
		if(u_year >=100 ) {
			u_year = 100 ;
		}else {
			u_year = 99 ;
		}
		String errMsg = "";
		String bank_type_name =( bank_type.equals("6") )?"農會":"漁會"; 
		String unit_name = Utility.getUnitName(unit); 		//取得顯示的金額單位
		DataObject bean = null;
      	try{
      		
      		//生出報表及設定格式===========================  
      		System.out.println("FR030WA_Excel.createRpt Starting........");
      	
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
						
			String openfile="台灣區農會信用部支票存款戶數與餘額明細表.xls";//要去開啟的範本檔
						
			//System.out.println("開啟檔:" + openfile);
			FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+ openfile );			
				
			//設定FileINputStream讀取Excel檔
			
			//新增一個xls unit
			POIFSFileSystem fs = new POIFSFileSystem( finput );
			
			if(fs==null){
				System.out.println("open 範本檔失敗");
			} 
			
			//新增一個sheet
			HSSFWorkbook wb = new HSSFWorkbook(fs);
			
			if(wb==null){
				System.out.println("open工作表失敗");
			} 
			
			//對第一個sheet工作
			HSSFSheet sheet = wb.getSheetAt(0);			//讀取第一個工作表，宣告其為sheet 
						
			if(sheet==null){
				System.out.println("open sheet 失敗");
			}
			
			//做屬性設定
			HSSFPrintSetup ps = sheet.getPrintSetup(); 	//取得設定
			//sheet.setZoom(80, 100); 					//螢幕上看到的縮放大小
			//sheet.setAutobreaks(true); 				//自動分頁
			
			//設定頁面符合列印大小
			sheet.setAutobreaks( false );
			ps.setScale( ( short )95 ); 				//列印縮放百分比
			
			ps.setPaperSize( ( short )9 ); 				//設定紙張大小 A4
			
			//設定表頭 為固定 先設欄的起始再設列的起始
			wb.setRepeatingRowsAndColumns(0, 1, 8, 2, 3);
			
			finput.close();
			
			HSSFRow row=null;//宣告一列 
			HSSFCell cell=null;//宣告一個儲存格

			//建表開始 ===================================
			
			int rowNum =4;//表頭有4列是已經有且固定的資料，所以row從第5(index為4)列開始生
			int columnNum =9;//本excel總共有18行
			
			//取得申報表資料
      		List dbData = getData(s_year,s_month,unit,bank_type,u_year);
      		
      		
      		//first row
	  		row = sheet.getRow(0);                                                                        
	   		cell = row.getCell((short)0);
	   		cell.setEncoding(HSSFCell.ENCODING_UTF_16);//顯示中文字
	   		cell.setCellValue("台灣區"+bank_type_name+"信用部支票存款戶數與餘額明細表");
			
			System.out.println(dbData.size()+"dbData.size()");
			
      		if(dbData == null || dbData.size() == 0){
      			System.out.println("dbData is null !!");
      			
      			//second row                                    
				row = sheet.getRow(1);                          
				cell = row.getCell((short)1);                  
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);      
				cell.setCellValue("民國 "+s_year+" 年 "+s_month+" 月底沒有資料" );
				cell = row.getCell((short)6);                  
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);     
				cell.setCellValue("單位:新臺幣 " + unit_name );  
			} else {
	 			//second row                                    
				row = sheet.getRow(1);                          
				cell = row.getCell((short)1);                  
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);     
				cell.setCellValue("民國 "+s_year+" 年 "+s_month+" 月底" );
				cell = row.getCell((short)6);                  
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);     
				cell.setCellValue("單位:新臺幣 " + unit_name );  
				
				//create all new row and cell for stroe table data 
				HSSFCellStyle cs = wb.createCellStyle();
				HSSFCellStyle csAccont = wb.createCellStyle();
					 			
				//set format for cell                                
				cs.setAlignment(HSSFCellStyle.ALIGN_LEFT);
				cs.setBorderTop(HSSFCellStyle.BORDER_THIN);
				cs.setBorderBottom(HSSFCellStyle.BORDER_THIN);
				cs.setBorderLeft(HSSFCellStyle.BORDER_THIN);
				cs.setBorderRight(HSSFCellStyle.BORDER_THIN);
				cs.setWrapText(true);
				csAccont.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
				csAccont.setBorderTop(HSSFCellStyle.BORDER_THIN);
				csAccont.setBorderBottom(HSSFCellStyle.BORDER_THIN);
				csAccont.setBorderLeft(HSSFCellStyle.BORDER_THIN);
				csAccont.setBorderRight(HSSFCellStyle.BORDER_THIN);
				csAccont.setDataFormat((short)3);
						
				
				
				//填入動態的資料
				String output ="";
				String bankNumber ="";
				for(int i=0;i<dbData.size();i++){
					
					//取得資料
					bean = (DataObject) dbData.get(i);
					
					//取得編號bank_no
					bankNumber = (bean.getValue("bank_no")==null)?" ":String.valueOf(bean.getValue("bank_no"));
					//如果沒有被選取的話就不做
					if(!isSelected(bank_list,bankNumber))
						continue;
					
					//去取得row如果沒有這個row就新增
					row = (sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
					rowNum++;
					
					for(int j=0;j<columnNum;j++){
						
						cell = (row.getCell((short) j)== null)?  row.createCell( (short) j):row.getCell((short) j);
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					}
									
					//信用部名稱
					output = (bean.getValue("bank_name")==null)?" ":String.valueOf(bean.getValue("bank_name"));
					cell = row.getCell((short)0);
					cell.setCellValue(output);
					cell.setCellStyle(cs);

					
					//總計戶數
					output = (bean.getValue("check_cnt_tot")==null)? " ":String.valueOf(bean.getValue("check_cnt_tot"));
					cell = row.getCell((short)1);
					cell.setCellValue(output);
					cell.setCellStyle(csAccont); 
				
					//總計餘額
					output = (bean.getValue("check_bal_tot")==null)? " ":String.valueOf(bean.getValue("check_bal_tot"));
					cell = row.getCell((short)2);
					cell.setCellValue(output);
					cell.setCellStyle(csAccont); 
					
					//正會員戶數
					output = (bean.getValue("checkbank_cnt")==null)? " ":String.valueOf(bean.getValue("checkbank_cnt"));
					cell = row.getCell((short)3);
					cell.setCellValue(output);
					cell.setCellStyle(csAccont);
								
					//正會員餘額
					output = (bean.getValue("checkbank_bal")==null)? " ":String.valueOf(bean.getValue("checkbank_bal"));
					cell = row.getCell((short)4);
					cell.setCellValue(output);
					cell.setCellStyle(csAccont);
					
					//贊助會員戶數
					output = (bean.getValue("checkbank_cnt_s")==null)? " ":String.valueOf(bean.getValue("checkbank_cnt_s"));
					cell = row.getCell((short)5);
					cell.setCellValue(output);
					cell.setCellStyle(csAccont);
			
					//贊助會員餘額
					output = (bean.getValue("checkbank_bal_s")==null)? " ":String.valueOf(bean.getValue("checkbank_bal_s"));
					cell = row.getCell((short)6);
					cell.setCellValue(output);
					cell.setCellStyle(csAccont);
		
					//非會員戶數
					output = (bean.getValue("checkbank_cnt_n")==null)? " ":String.valueOf(bean.getValue("checkbank_cnt_n"));
					cell = row.getCell((short)7);
					cell.setCellValue(output);
					cell.setCellStyle(csAccont);
			
					//非會員餘額
					output = (bean.getValue("checkbank_bal_n")==null)? " ":String.valueOf(bean.getValue("checkbank_bal_n"));
					cell = row.getCell((short)8);
					cell.setCellValue(output);
					cell.setCellStyle(csAccont);
					
				}//end of inner-for
			}//end of outer-for
			
			//建表結束 =================================
			HSSFFooter footer = sheet.getFooter();
			footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
			
			FileOutputStream fout=new FileOutputStream(reportDir + System.getProperty("file.separator")+
													   "台灣區" + bank_type_name + "信用部支票存款戶數與餘額明細表.xls");  
			
			wb.write(fout);//儲存
			fout.close();
			System.out.println("儲存完成");
			
			
		}catch(Exception e){
			System.out.println("createRpt Error:"+e+e.getMessage());                                                
		}
		
		System.out.println("FR030WA_Excel.createRpt Ending........");
		
		return errMsg;
	}//end of createreport
	
	
	/*  
	 *  ===============================================
	 *	取得支票存款戶數與餘額明細的各項資料
	 *  ===============================================
	 */
	public static List getData(String s_year,String s_month,String unit,String bank_type,int u_year){
		StringBuffer sql = new StringBuffer () ;
		List paramList = new ArrayList();
		if(u_year>=100) {
			sql.append(" select a01.* from ");
			sql.append(" ( select nvl(cd01.hsien_id,' ') as hsien_id ,nvl(cd01.hsien_name,'OTHER')  as  hsien_name,   ");
			sql.append("          cd01.FR001W_output_order as FR001W_output_order,bn01.bank_no, bn01.BANK_NAME,       ");
			sql.append("          (CheckBank_Cnt + CheckBank_Cnt_S + CheckBank_Cnt_N) as Check_cnt_TOT,               ");
			sql.append("          Round( (CheckBank_Bal + CheckBank_Bal_S + CheckBank_Bal_N)/?,0) as Check_Bal_TOT,   ");
			sql.append("          CheckBank_Cnt, Round(CheckBank_Bal  /?,0) As CheckBank_Bal,                         ");
			sql.append("          CheckBank_Cnt_S, Round( CheckBank_Bal_S  /?,0) as CheckBank_Bal_S,                  ");
			sql.append("          CheckBank_Cnt_N, Round(CheckBank_Bal_N  /?,0) as CheckBank_Bal_N                    ");
			sql.append("   from  (select * from cd01 where cd01.hsien_id <> 'Y') cd01                                 ");
			sql.append("   left join (SELECT * FROM WLX01 WHERE M_YEAR=? )wlx01 on wlx01.hsien_id=cd01.hsien_id ");
			sql.append("   left join (SELECT * FROM BN01 WHERE BANK_TYPE=? AND M_YEAR=? )bn01 on wlx01.bank_no=bn01.bank_no  ");
			sql.append("   left join (select * from WLX07_M_CHECKBANK																			");
			sql.append(" 		     where m_year  = ? and m_month =  ?) a01 on  bn01.bank_no = a01.bank_no ");
			sql.append(" ) a01 ,v_bank_location T2  ");
			sql.append(" where a01.bank_no  <>  ' '  and  a01.CheckBank_Cnt > 0                           ");
			sql.append(" and a01.bank_no = T2.bank_No and T2.m_year=? ");
			sql.append(" order by  T2.FR001W_output_order, a01.hsien_id,  a01.bank_no                    ");
		}else {
			sql.append(" select a01.* from ");
			sql.append(" ( select nvl(cd01.hsien_id,' ') as hsien_id ,nvl(cd01.hsien_name,'OTHER')  as  hsien_name,										");
			sql.append("          cd01.FR001W_output_order as FR001W_output_order,bn01.bank_no, bn01.BANK_NAME, (CheckBank_Cnt +      ");
			sql.append("          CheckBank_Cnt_S + CheckBank_Cnt_N) as Check_cnt_TOT, Round( (CheckBank_Bal + CheckBank_Bal_S +      ");
			sql.append("          CheckBank_Bal_N)/?,0) as Check_Bal_TOT, CheckBank_Cnt, Round(CheckBank_Bal  /?,0) As CheckBank_Bal, ");
			sql.append("          CheckBank_Cnt_S, Round( CheckBank_Bal_S  /?,0) as CheckBank_Bal_S, CheckBank_Cnt_N,                 ");
			sql.append("          Round(CheckBank_Bal_N  /?,0) as CheckBank_Bal_N                                                     ");
			sql.append("   from  (select * from cd01_99 cd01 where cd01.hsien_id <> 'Y') cd01                                         ");
			sql.append("   left join (select * from wlx01 where m_year=? )wlx01 on wlx01.hsien_id=cd01.hsien_id ");
			sql.append("   left join (select * from bn01 where bank_type=? and m_year=? )bn01 on wlx01.bank_no=bn01.bank_no ");
			sql.append("   left join (select * from WLX07_M_CHECKBANK																			");
			sql.append(" 		     where m_year  = ? and m_month =  ?) a01 on  bn01.bank_no = a01.bank_no  ");
			sql.append(" ) a01   ,v_bank_location T2 ");
			sql.append(" where a01.bank_no  <>  ' '  and  a01.CheckBank_Cnt > 0                           ");
			sql.append(" and a01.bank_no = T2.bank_No and T2.m_year=? ");
			sql.append(" order by  T2.FR001W_output_order, a01.hsien_id,  a01.bank_no                    ");
		}
		paramList.add(unit);
		paramList.add(unit);
		paramList.add(unit);
		paramList.add(unit);
		paramList.add(u_year) ;
		paramList.add(bank_type);
		paramList.add(u_year) ;
		paramList.add(s_year) ;
		paramList.add(s_month) ;
		paramList.add(u_year) ;
		List dbData = null;
		
		dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "hsien_id,hsien_name,fr001w_output_order,bank_no,bank_name,"+
										  "check_cnt_tot,check_bal_tot,checkbank_cnt,checkbank_bal,checkbank_cnt_s,"+
										  "checkbank_bal_s,checkbank_cnt_n,checkbank_bal_n,");		
		
		
		return dbData;
	}//end of getData
	
	/*  
	 *	===============================================
	 *	判斷有沒有被選取的信用部(有:true 沒有 :false)
	 *  ===============================================
	 */
	
	public static boolean isSelected(List bank_list,String bank_no){
		List temp;
		String bank_id ="";
		boolean flg = false ;
		for(int v=0;v<bank_list.size();v++){
     		temp = (List)(bank_list.get(v));
     		bank_id = String.valueOf(temp.get(0));    
     		//System.out.println("所選的bank_id:"+bank_id) ;
     		if(bank_id.equals(bank_no)){ 
				//System.out.println("bank_id ="+bank_id);
				flg = true ;
			}
     		if(flg) break ;
     	}
		
     	return flg;   
	}
                                                                  
}//end of class                                                                                                                       