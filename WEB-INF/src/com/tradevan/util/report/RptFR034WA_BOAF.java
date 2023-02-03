/*
  Created on 2006/11/06 by 2495
  fix sqlInjection by 2808
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

public class RptFR034WA_BOAF{
  	public static String createRpt(String s_year,String s_month,String unit,String bank_type)
	{
    
		String errMsg = "";		
		String s_year_last="";
		String s_month_last="";	 
		String bank_type_name=(bank_type.equals("6"))?"農會":"漁會";
		//added by 2808 99.12.08
		String m_year= "99" ;
		String cd01Table = "cd01_99" ;
		if(Integer.parseInt(s_year) > 99) {
			m_year = "100" ;
			cd01Table = "cd01" ;
		}
		String unit_name = "";

           if (unit.equals("1")){
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
		
		String filename="";
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
			 cell.setCellValue(s_year+" 年 "+s_month+" 季                            單位："+unit_name+"，戶，筆");									   
        	 ft2.setFontHeightInPoints((short)12);
		     ft2.setFontName("標楷體");
		     cs2.setFont(ft2);
		     cs2.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
        	 cell.setCellStyle(cs2);

	int rowNum =5;

	String bank_id = "";	
    
		


//================================================================================================================	
List paramList =new ArrayList() ;
String sqlCmd =                                                                                           
"  select   *                                                                                    "+
" from (                                                                                         "+
" select nvl(cd01.hsien_id,' ')       as  hsien_id ,                                             "+
"        nvl(cd01.hsien_name,'OTHER')  as  hsien_name,                                           "+
"        cd01.FR001W_output_order     as  FR001W_output_order,                                   "+
"        bn01.bank_no ,  bn01.BANK_NAME,                                                         "+
"        a01.WarnAccount_TCnt,                                                                   "+
" Round(a01.WarnAccount_Tbal/?,0)  as  WarnAccount_Tbal,                                  "+
" a01.WarnAccount_Remit_TCnt,                                                                    "+
" a01.WarnAccount_Refund_Apply_Cnt,                                                              "+
" round(a01.WarnAccount_Refund_Apply_Amt/?,0)  as  WarnAccount_Refund_Apply_Amt,          "+
" a01.WarnAccount_Refund_Cnt,                                                                    "+
" round(a01.WarnAccount_Refund_Amt /?,0)  as  WarnAccount_Refund_Amt                      "+
" from  (select * from "+cd01Table+" where hsien_id <> 'Y') cd01                                     "+
" left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id           " ;
paramList.add(m_year) ;
sqlCmd +="        left join (select * from bn01 where m_year=? )bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=?       " ;
paramList.add(m_year);
sqlCmd +="        left join (select * from WLX09_S_WARNING where m_year  = ? and m_quarter =?) a01    "+
" on  bn01.bank_no = a01.bank_no                                                                 "+
" ) a01   where a01.bank_no  <>  ' '  and  a01.WarnAccount_TCnt > 0                              "+
"          order by  a01.FR001W_output_order, a01.hsien_id,  a01.bank_no                         ";
paramList.add(unit) ;
paramList.add(unit) ;
paramList.add(unit) ;
paramList.add(bank_type) ;
paramList.add(s_year) ;
paramList.add(s_month) ;
//==========================================================================================================		
		List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"hsien_id,hsien_name,bank_no,bank_name,warnaccount_tcnt"+
											   "warnaccount_tbal,warnaccount_remit_tcnt,warnaccount_refund_apply_cnt"+
											   "warnaccount_refund_apply_amt,warnaccount_refund_cnt,warnaccount_refund_amt");
     DataObject bean = null;
     if(dbData!=null){
     	String bank_name_temp ="";
    System.out.println("dbData.size():"+dbData.size());	
    for(int k=0;k<dbData.size();k++){
     	bean = (DataObject)dbData.get(k);
     	
 		String hsien_id = String.valueOf(bean.getValue("hsien_id"));
 		String hsien_name = String.valueOf(bean.getValue("hsien_name")); 
 		String bank_no = String.valueOf(bean.getValue("bank_no"));
 		String bank_name = String.valueOf(bean.getValue("bank_name"));
 		String warnaccount_tcnt =  Utility.setCommaFormat(String.valueOf(bean.getValue("warnaccount_tcnt")));
		String warnaccount_tbal =  Utility.setCommaFormat(String.valueOf(bean.getValue("warnaccount_tbal")));
		String warnaccount_remit_tcnt =  Utility.setCommaFormat(String.valueOf(bean.getValue("warnaccount_remit_tcnt")));
		String warnaccount_refund_apply_cnt =  Utility.setCommaFormat(String.valueOf(bean.getValue("warnaccount_refund_apply_cnt")));
		String warnaccount_refund_apply_amt =  Utility.setCommaFormat(String.valueOf(bean.getValue("warnaccount_refund_apply_amt")));
		String warnaccount_refund_cnt =  Utility.setCommaFormat(String.valueOf(bean.getValue("warnaccount_refund_cnt")));
		String warnaccount_refund_amt =  Utility.setCommaFormat(String.valueOf(bean.getValue("warnaccount_refund_amt")));
	
		
		HSSFFont f = wb.createFont();
		HSSFCellStyle c = wb.createCellStyle();
			   
		//將所有抓到的資料，結合所選取的縣市別來填表	   
		List temp ;	  
		
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
		}
		return errMsg;
	}	
}



