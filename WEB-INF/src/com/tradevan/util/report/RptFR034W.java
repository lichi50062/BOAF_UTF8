/*
  Created on 2006/01/03 by 4180
  97.04.17 fix 即使為"0"也要顥示 by 2295
  99.04.13 fix 因應縣市合併調整SQL &修改SQL查詢方式以preparedStatement
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

public class RptFR034W {
  	public static String createRpt(String s_year,String s_month,String unit,String bank_type,List bank_list){
  		String errMsg = "";		
		String s_year_last="";
		String s_month_last="";	 
		String bank_type_name=(bank_type.equals("6"))?"農會":"漁會";	
		String unit_name = Utility.getUnitName(unit);//取得單位名稱
        String filename="";
        String u_year = "100" ; //縣市合併判斷用 
        if(s_year==null || Integer.parseInt(s_year) < 100) {
        	u_year = "99" ;
        }
		filename=(bank_type.equals("6"))?"農漁會信用部警示帳戶調查統計明細表.xls":"農漁會信用部警示帳戶調查統計明細表.xls";		
	
	
	
		try{				
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
    		String openfile="農漁會信用部警示帳戶調查統計明細表.xls";
    		System.out.println("open file "+openfile);
			FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+ openfile );			
			
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
	        ps.setScale( ( short )80 ); //列印縮放百分比

	        ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
	        
	        //設定表頭 為固定 先設欄的起始再設列的起始
	        wb.setRepeatingRowsAndColumns(0, 1, 10, 1, 3);
	  		
	  		finput.close();
	  		
	  		
	  		HSSFRow row=null;//宣告一列 
	  		HSSFCell cell=null;//宣告一個儲存格

			HSSFFont ft = wb.createFont();
			HSSFCellStyle cs = wb.createCellStyle();
			HSSFFont ft2 = wb.createFont();
			HSSFCellStyle cs2 = wb.createCellStyle();

		     
			 
		
      		row = sheet.getRow(0);
      		cell = row.getCell( (short) 0);
      		cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        	cell.setCellValue("全體"+ bank_type_name +"信用部警示帳戶調查統計明細表");	
		    
		    ft.setFontHeightInPoints((short)18);
		    ft.setFontName("標楷體");
		    cs.setFont(ft);
		    cs.setAlignment(HSSFCellStyle.ALIGN_CENTER);        	   
        	cell.setCellStyle(cs);
  		      
		    
			row= sheet.getRow(1);													
      		cell = row.getCell( (short) 0);		
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);			
			cell.setCellValue(" 單位："+unit_name+"，戶，筆");									   
        	ft2.setFontHeightInPoints((short)12);
		    ft2.setFontName("標楷體");
		    cs2.setFont(ft2);
		    cs2.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
        	cell.setCellStyle(cs2);

			int rowNum =5;
			//取得bank_list的bank_no
			String bank_id = "";	
     		if(bank_list!=null){	
     			StringBuffer sql = new StringBuffer() ;
     			List paramList = new ArrayList() ;
     			String cd01Table = "cd01_99" ;
     			if("100".equals(u_year)){
     				cd01Table = "cd01" ;
     			}
     			sql.append(" select   *                  ");                                                                   
 				sql.append(" from (                      ");                                                                    
 				sql.append("       select nvl(cd01.hsien_id,' ')       as  hsien_id ,             ");                                      
 				sql.append("              nvl(cd01.hsien_name,'OTHER') as  hsien_name,            ");                                
 				sql.append("              cd01.FR001W_output_order     as  FR001W_output_order,   ");                                 
 				sql.append("              bn01.bank_no ,  bn01.BANK_NAME,                         ");                                 
 				sql.append("              to_char(a01.WarnAccount_TCnt) WarnAccount_TCnt ,  ");                                 
 				sql.append("              to_char(round(a01.WarnAccount_Tbal/ ?,0))  as  WarnAccount_Tbal, ");                                  
 				sql.append("              to_char(a01.WarnAccount_Remit_TCnt)WarnAccount_Remit_TCnt ,");                                        
 				sql.append("              to_char(a01.WarnAccount_Refund_Apply_Cnt)WarnAccount_Refund_Apply_Cnt, ");                                        
 				sql.append("              to_char(round(a01.WarnAccount_Refund_Apply_Amt/?,0))  as  WarnAccount_Refund_Apply_Amt,    ");       
 				sql.append("              to_char(a01.WarnAccount_Refund_Cnt) WarnAccount_Refund_Cnt,  ");
 				sql.append("              to_char(round(a01.WarnAccount_Refund_Amt / ?,0))  as  WarnAccount_Refund_Amt               ");        
 				sql.append("       from  (select * from ").append(cd01Table).append("  where hsien_id <> 'Y') cd01                       ");               
 				sql.append("       left join (select * from wlx01 where m_year=? )wlx01 on wlx01.hsien_id=cd01.hsien_id     ");                            
 				sql.append("       left join (select * from bn01 where bank_type=? and m_year=?)bn01 on wlx01.bank_no=bn01.bank_no  ");      
 				sql.append("       left join (select * from WLX09_S_WARNING where m_year  =  ? and m_quarter =  ? ) a01 on  bn01.bank_no = a01.bank_no      ");                                                            
 				sql.append(" ) a01   where a01.bank_no  <>  ' '  ");
 				sql.append("         and a01.bank_no in (").append(getSelectBankNo(bank_list)).append(" )  ");
 				sql.append(" order by  a01.FR001W_output_order, a01.hsien_id,  a01.bank_no   ");
     			
     			paramList.add(unit) ;
     			paramList.add(unit) ;
     			paramList.add(unit) ;
     			paramList.add(u_year) ;
     			paramList.add(bank_type);
     			paramList.add(u_year) ;
     			paramList.add(s_year) ;
     			paramList.add(s_month) ;
     			List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"hsien_id,hsien_name,bank_no,bank_name,warnaccount_tcnt"+
											   "warnaccount_tbal,warnaccount_remit_tcnt,warnaccount_refund_apply_cnt"+
											   "warnaccount_refund_apply_amt,warnaccount_refund_cnt,warnaccount_refund_amt");
     			sql = null ;
     			paramList = null ;
               //================================================================================================================		
               DataObject bean = null;
     		   if(dbData!=null){
     			  String bank_name_temp ="";
     			  
                  for(int k=0;k<dbData.size();k++){
                   	bean = (DataObject)dbData.get(k);
                   	//Object [] testObj = bean.getKeys();
                   	//for(int i=0 ; i< testObj.length ; i++) {
                   	//	System.out.println("123:"+Utility.getTrimString(testObj[i])) ;
                   	//}
                   	String hsien_id = String.valueOf(bean.getValue("hsien_id"));
 	              	String hsien_name = String.valueOf(bean.getValue("hsien_name")); 
 	              	String bank_no = String.valueOf(bean.getValue("bank_no"));
 	              	String bank_name = String.valueOf(bean.getValue("bank_name"));
 	              	String warnaccount_tcnt =  Utility.setCommaFormat(Utility.getTrimString(bean.getValue("warnaccount_tcnt")));
	              	String warnaccount_tbal =  Utility.setCommaFormat(Utility.getTrimString(bean.getValue("warnaccount_tbal")));
	              	String warnaccount_remit_tcnt =  Utility.setCommaFormat(Utility.getTrimString(bean.getValue("warnaccount_remit_tcnt")));
	              	String warnaccount_refund_apply_cnt =  Utility.setCommaFormat(Utility.getTrimString(bean.getValue("warnaccount_refund_apply_cnt")));
	              	String warnaccount_refund_apply_amt =  Utility.setCommaFormat(Utility.getTrimString(bean.getValue("warnaccount_refund_apply_amt")));
	              	String warnaccount_refund_cnt =  Utility.setCommaFormat(Utility.getTrimString(bean.getValue("warnaccount_refund_cnt")));
	              	String warnaccount_refund_amt =  Utility.setCommaFormat(Utility.getTrimString(bean.getValue("warnaccount_refund_amt")));
	              
	              		
	              	HSSFFont f = wb.createFont();
	              	HSSFCellStyle c = wb.createCellStyle();
	              	row = sheet.createRow(rowNum++);                 
              	    for(int cellcount=0;cellcount<8;cellcount++){ 	              	                 
                		cell = row.createCell( (short)cellcount);
                		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                		if(cellcount==0 && !bank_name_temp.equals(bank_name))
                  	    { cell.setCellValue(bank_name);	bank_name_temp = bank_name;}
                  	    else if(cellcount==1)
                  	      cell.setCellValue(warnaccount_tcnt);   
                  	    else if(cellcount==2)
                  	      cell.setCellValue(warnaccount_tbal);	
                   	    else if(cellcount==3)
                  	    cell.setCellValue(warnaccount_remit_tcnt);	
                          else if(cellcount==4)
                  	    cell.setCellValue(warnaccount_refund_apply_cnt);
                  	    else if(cellcount==5)
                          cell.setCellValue(warnaccount_refund_apply_amt);   
                  	    else if(cellcount==6)
                  	    cell.setCellValue(warnaccount_refund_cnt);		
                		    else if(cellcount==7)
                  	    cell.setCellValue(warnaccount_refund_amt);    
                  	       
              	        f.setFontHeightInPoints((short)12);
              	        f.setFontName("標楷體");
              	        c.setFont(f);
              	        c.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
              	        c.setBorderTop((short)0);
                        c.setBorderBottom(HSSFCellStyle.BORDER_THIN);
                        c.setBorderLeft(HSSFCellStyle.BORDER_THIN);
                        if(cellcount==7)
                        c.setBorderRight(HSSFCellStyle.BORDER_THIN);
                        cell.setCellStyle(c);	
                    }	//end of cell set
                    
                  }//end of dbData size
                  rowNum = dbData.size()+rowNum;
	           }

		    }//end if
			HSSFFooter footer=sheet.getFooter();
	        footer.setCenter( "Page:" +HSSFFooter.page() +" of " +HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
		    FileOutputStream fout=new FileOutputStream(reportDir+ System.getProperty("file.separator")+ openfile);
		    wb.write(fout);
	        //儲存 
	        fout.close();
	        System.out.println("儲存完成");
	        
		}catch(Exception e){
			System.out.println("createRpt Error:"+e+e.getMessage());
			e.printStackTrace();
		}
		return errMsg;
	}	
  	/***
     * 回傳已選的單位.
     * 
     * @param bank_list
     * @return
     */
    private static String getSelectBankNo(List bank_list) {
    	List temp;
    	String bank_id ="" ;
    	for (int v = 0; v < bank_list.size(); v++) {
			temp = (List) (bank_list.get(v));
			if("".equals(bank_id)) {
				bank_id = "'"+String.valueOf(temp.get(0))+"'"; 
			}else {
				bank_id += ",'"+String.valueOf(temp.get(0))+"'"; 
			}
			
		}
    	return bank_id ;
    }
}



