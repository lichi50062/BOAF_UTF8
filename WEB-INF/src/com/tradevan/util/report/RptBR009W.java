/*
  Created on 2006/03/30 農漁會管理人員名冊 by 2495
  2010.12.30 fix 根據查詢年度.100年以後取得新縣市別.100年以前取得舊縣市別 
 				 使用PreparedStatement;並列印轉換後的SQL;套用QueryDB_SQLParam by 2295
*/
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import java.io.*;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptBR009W{
  	public static String createRpt(String s_year,String s_month,String bank_type,List bank_list){
    		String errMsg = "";	
			String bank_type_name=(bank_type.equals("6"))?"農會":"漁會";	
			String filename="";
			String temp_audit_name="",temp_audit_telno="",temp_it_name="",temp_it_telno="",t_1="",t_2="",temp_t_1="",temp_t_2="";
			String cd01_table = "";//99.12.30 add
	        String wlx01_m_year = "";//99.12.30 add
			int count=0;
			List paramList = new ArrayList(); 	
			StringBuffer sqlCmd = new StringBuffer();
			int flag=0,flag_print=0;
			filename=(bank_type.equals("6"))?"農漁會管理人員名冊.xls":"農漁會管理人員名冊.xls";		
			String report_no = "BR009W";
			try
			{				
				//99.12.30 add 查詢年度100年以前.縣市別不同===============================
	  	    	cd01_table = (Integer.parseInt(s_year) < 100)?"cd01_99":""; 
	  	    	wlx01_m_year = (Integer.parseInt(s_year) < 100)?"99":"100"; 
	  	    	//=====================================================================   
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
    		String openfile="農漁會管理人員名冊.xls";
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
	        ps.setScale( ( short )70 ); //列印縮放百分比

	        ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
	        
	        //設定表頭 為固定 先設欄的起始再設列的起始
	        //wb.setRepeatingRowsAndColumns(0, 1, 10, 1, 3);
	  		
	  		finput.close();
	  		
	  		
	  		HSSFRow row=null;//宣告一列 
	  		HSSFCell cell=null;//宣告一個儲存格

			HSSFFont ft = wb.createFont();
			HSSFCellStyle cs = wb.createCellStyle();
			HSSFFont ft2 = wb.createFont();
			HSSFCellStyle cs2 = wb.createCellStyle();
			HSSFFont f = wb.createFont();
			HSSFCellStyle c = wb.createCellStyle();
		     
			 
			
      		row = sheet.getRow(0);
      		cell = row.getCell( (short) 0);
      		cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        	cell.setCellValue( bank_type_name +"管理人員名冊");	
		    f.setFontName("標楷體");
		    f.setFontHeightInPoints((short)18);	
		    cs.setFont(f);	     
		    cs.setAlignment(HSSFCellStyle.ALIGN_CENTER);       	   
            cell.setCellStyle(cs);
  		       

	        int rowNum =0;
	        //取得bank_list的bank_no
	        String bank_id = "";	
            if(bank_list!=null){
               //================================================================================================================
            	//總機構
                String  SQL_TUP_UNIT = 
                " select '0' as t_1, "+
                " '1' as t_2, "+
                " '0' as t_3, "+
                " aa.bank_no as tbank_no, "+  
                " aa.bank_no as bank_no,  "+
                " aa.bank_name, aa.HSIEN_ID, aa.area_id, cd01.HSIEN_NAME  as  addr_1, "+ 
                " CD02.AREA_name  as addr_2, "+
                " aa.ADDR  as addr_3, "+ 
                " aa.TELNO, aa.FAX, aa.EMAIL, aa.WEB_SITE, aa.AUDIT_NAME, aa.AUDIT_TELNO, aa.IT_NAME, aa.IT_TELNO "+
                " from "+ 
                " (select wlx01.* , bn01.bank_name "+
                " from (select * from bn01 where m_year=?)bn01, (select * from wlx01 where m_year=?)wlx01 "+
 			    " where bn01.bank_type = ? and "+
                " bn01.bank_no  = wlx01.bank_no and "+
                " wlx01.CANCEL_NO <> 'Y') aa "+  
                " left join cd01 on  aa.HSIEN_ID = CD01.HSIEN_ID "+
                " left join cd02 on  aa.AREA_ID =  CD02.AREA_ID ";
                paramList.add(wlx01_m_year);
                paramList.add(wlx01_m_year);
                paramList.add(bank_type);
                System.out.println("SQL_TUP_UNIT: "+SQL_TUP_UNIT);
                
               //總機構_高階主管	
               String SQL_TUP_RANK = 
               " select '0' as t_1,"+
               "        '2' as t_2,"+
               " aa.position_code as t_3, "+
               " aa.bank_no as tbank_no, "+
               " aa.bank_no as bank_no, "+
               " aa.bank_name, aa.HSIEN_ID,  ' ' as area_id, "+
               " ab.CmUSE_name  as  addr_1, "+
               " aa.name    as addr_2, "+
               " ' '        as addr_3, "+
               " aa.TELNO, aa.FAX,   aa.EMAIL,  ' ' as  WEB_SITE, "+ 
               " ' ' as  AUDIT_NAME, "+
               " ' ' as AUDIT_TELNO, "+   
               " ' ' as IT_NAME, "+
               " ' ' as IT_TELNO "+
               " from "+
               " ( select wlx01_m.* , bn01.bank_name, wlx01.HSIEN_ID "+
               " from (select * from bn01 where m_year=?)bn01, wlx01_m, (select * from wlx01 where m_year=?)wlx01 "+
			   " where bn01.bank_type = ?  and "+
               " bn01.bank_no  = wlx01.bank_no and "+
               " wlx01.CANCEL_NO <> 'Y'  and "+ 
               " bn01.bank_no = wlx01_m.bank_no and "+
               " WLX01_M.ABDICATE_CODE <> 'Y' "+
               " ) aa "+  
               " left join CDShareNO ab on  ab.CmUSE_Div = '005' and "+
               " aa.position_code = ab.CmUSE_id ";
               paramList.add(wlx01_m_year);
               paramList.add(wlx01_m_year);
               paramList.add(bank_type);
               //分支機構
               String  SQL_SUB_UNIT = 
               " select '1' as t_1, "+
               " '1' as t_2, "+
               " '0' as t_3, "+ 
               " aa.tbank_no as tbank_no, "+  
               " aa.bank_no  as bank_no,  "+
               " aa.bank_name, "+
               " aa.HSIEN_ID_1 as  HSIEN_ID, "+
               " aa.area_id, cd01.HSIEN_NAME  as  addr_1, "+
               " CD02.AREA_name  as addr_2, "+
               " aa.ADDR  as addr_3, "+
               " aa.TELNO, aa.FAX, aa.EMAIL, aa.WEB_SITE, "+ 
               " ' ' as  AUDIT_NAME, "+
               " ' ' as AUDIT_TELNO, "+   
               " ' ' as IT_NAME, "+
               " ' ' as IT_TELNO "+
               " from "+ 
               " (select wlx02.*, bn02.BANK_NAME,  wlx01.HSIEN_ID as HSIEN_ID_1 "+
               " from (select * from bn02 where m_year=?)bn02, (select * from wlx02 where m_year=?)wlx02, (select * from wlx01 where m_year=?)wlx01 "+
			   " where bn02.bank_type = ?  and "+
			   " bn02.tbank_no  = wlx01.bank_no and "+
               " wlx01.CANCEL_NO <> 'Y' and "+
               " bn02.tbank_no = wlx02.tbank_no  and "+
               " bn02.bank_no = wlx02.bank_no    and "+
               " wlx02.CANCEL_NO <> 'Y') aa "+  
               " left join cd01 on  aa.HSIEN_ID = CD01.HSIEN_ID "+
               " left join cd02 on  aa.AREA_ID =  CD02.AREA_ID ";
               paramList.add(wlx01_m_year);
               paramList.add(wlx01_m_year);
               paramList.add(wlx01_m_year);
               paramList.add(bank_type);
               System.out.println("SQL_SUB_UNIT: "+SQL_SUB_UNIT);
               
               //分支機構_高階主管
               String  SQL_SUB_RANK = 
               " select '1' as t_1, "+
               "        '2' as t_2, "+
               " aa.position_code as t_3, "+ 
               " aa.tbank_no as tbank_no, "+  
               " aa.bank_no as bank_no, "+
               " aa.bank_name, aa.HSIEN_ID_1 as HSIEN_ID, "+  
               " ' ' as area_id, "+
               " ab.CmUSE_name  as  addr_1, "+
               " aa.name        as addr_2,  "+
               " ' '            as addr_3,  "+
               " aa.TELNO, ' '  as FAX, "+
               " aa.EMAIL, ' '  as  WEB_SITE, "+ 
               " ' ' as  AUDIT_NAME, "+
               " ' ' as  AUDIT_TELNO, "+   
               " ' ' as IT_NAME, "+
               " ' ' as IT_TELNO "+
               " from "+
               " (select wlx02_M.*, wlx02.tbank_no,  bn02.BANK_NAME, "+
               " wlx01.HSIEN_ID as HSIEN_ID_1 "+
               " from (select * from bn02 where m_year=?)bn02, (select * from wlx02 where m_year=?)wlx02, (select * from wlx01 where m_year=?)wlx01, wlx02_M "+
               " where bn02.bank_type = ?  and "+
               " bn02.tbank_no  = wlx01.bank_no and "+
               " wlx01.CANCEL_NO <> 'Y' and "+
               " bn02.tbank_no = wlx02.tbank_no and "+
               " bn02.bank_no = wlx02.bank_no   and "+
               " wlx02.CANCEL_NO <> 'Y'         and "+
               " bn02.bank_no = wlx02_m.bank_no and "+
               " WLX02_M.ABDICATE_CODE <> 'Y' ) aa  "+ 
               " left join CDShareNO ab on  ab.CmUSE_Div = '007' and "+  
               " aa.position_code = ab.CmUSE_id ";
               paramList.add(wlx01_m_year);
               paramList.add(wlx01_m_year);
               paramList.add(wlx01_m_year);
               paramList.add(bank_type);
               System.out.println("SQL_SUB_RANK: "+SQL_SUB_RANK);
               
               //組合
               String  SQL_COMBINE_FINAL = " select * from ("+ SQL_TUP_UNIT + " UNION ALL "+ SQL_TUP_RANK + " UNION ALL "+ SQL_SUB_UNIT + " UNION ALL "+ SQL_SUB_RANK + " ) temp_result  order by HSIEN_ID, tbank_no, bank_no, t_1, t_2, t_3";                                      
              
               //==========================================================================================================		
			   List dbData = DBManager.QueryDB_SQLParam(SQL_COMBINE_FINAL.toString(),paramList,"t_1,t_2,t_3,tbank_no,bank_no,addr_1,addr_2,addr_3,audit_name,audit_telno,it_name,it_telno");		
    		   System.out.print("抓出的dbData"+String.valueOf(dbData.size()) );

     		   DataObject bean = null;
     		  if(dbData!=null){
     			 String bank_name_temp ="";
      
      
    for(int k=0;k<dbData.size();k++){
        bean = (DataObject)dbData.get(k);
     	 		
 		t_1 = String.valueOf(bean.getValue("t_1"));
 		t_2 = String.valueOf(bean.getValue("t_2")); 
 		String t_3 = String.valueOf(bean.getValue("t_3"));
 		String tbank_no = String.valueOf(bean.getValue("tbank_no"));
 		
 		String bank_no = String.valueOf(bean.getValue("bank_no"));
 		String bank_name = String.valueOf(bean.getValue("bank_name"));
 		String area_id = String.valueOf(bean.getValue("area_id"));
 		String addr_1 = String.valueOf(bean.getValue("addr_1"));
 		String addr_2 = String.valueOf(bean.getValue("addr_2"));
 		String addr_3 = String.valueOf(bean.getValue("addr_3"));
 		String telno = String.valueOf(bean.getValue("telno"));
 		String fax = String.valueOf(bean.getValue("fax")); 
 		String email = String.valueOf(bean.getValue("email"));
 		String web_site = String.valueOf(bean.getValue("web_site"));
 		String audit_name = String.valueOf(bean.getValue("audit_name"));
 		String audit_telno = String.valueOf(bean.getValue("audit_telno"));
 		String it_name = String.valueOf(bean.getValue("it_name"));
 		String it_telno = String.valueOf(bean.getValue("it_telno"));
    
   	    if(t_1.equals("null")){t_1="";}
		if(t_2.equals("null")){t_2="";}
		if(t_3.equals("null")){t_3="";}
		if(tbank_no.equals("null")){tbank_no="";}
		if(bank_no.equals("null")){bank_no="";}
		if(bank_name.equals("null")){bank_name="";}
		if(area_id.equals("null")){area_id="";}
		if(addr_1.equals("null")){addr_1="";}
		if(addr_2.equals("null")){addr_2="";}
		if(addr_3.equals("null")){addr_3="";}
		if(telno.equals("null")){telno="";}
		if(fax.equals("null")){fax="";}
		if(email.equals("null")){email="";}
		if(web_site.equals("null")){web_site="";}
		
		if(audit_name.equals("null")){audit_name="";}
		if(audit_telno.equals("null")){audit_telno="";}
		if(it_name.equals("null")){it_name="";}
		if(it_telno.equals("null")){it_telno="";}
		
			   
		//將所有抓到的資料，結合所選取的縣市別來填表	   
		List temp ;	  
		boolean isSelected = false;  //判斷是否有在選取的清單中
		   
		for(int v=0;v<bank_list.size();v++){
     	    temp = (List)(bank_list.get(v));
     	    bank_id = String.valueOf(temp.get(0));
     	    if(bank_id.equals(tbank_no))
     	    { 
     	     	 isSelected=true;
     	       break;
     	    }
     	}
     	
     	
     	   
     	   	       
     	if(isSelected){
     		if(temp_t_1.equals("0")&&temp_t_2.equals("2")&&t_1.equals("0")&&t_2.equals("1")&&flag_print==0){
     			flag_print=1;
     			rowNum=rowNum+1;
     		    for(int cellcount=0;cellcount<8;cellcount++){
			        row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
			        cell=row.createCell((short)cellcount);
			        cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			        c.setAlignment(HSSFCellStyle.ALIGN_CENTER);			        
			        c.setBorderLeft(HSSFCellStyle.BORDER_THIN);
			        c.setBorderRight(HSSFCellStyle.BORDER_THIN);
			        cell.setCellStyle(c);    				     				
      			    if(cellcount==0) cell.setCellValue("稽核聯絡人員");      				
      			    if(cellcount==1) cell.setCellValue(temp_audit_name);   					
      			    if(cellcount==2){       										
      			    	sheet.addMergedRegion(new Region(rowNum,(short)2,rowNum,(short)3));     			   
      			    	cell.setCellValue(temp_audit_telno);
      			    }
      			    if(cellcount==5){    					
      			    	sheet.addMergedRegion(new Region(rowNum,(short)5,rowNum,(short)6)); 
      			    	cell.setCellValue("");
      			    }     		
      	        }//end of cellcount
      	        
      	        rowNum=rowNum+1;
                for(int cellcount=0;cellcount<8;cellcount++){  	           			
		        	row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
		        	cell=row.createCell((short)cellcount);
		        	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		        	c.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		        	c.setBorderTop(HSSFCellStyle.BORDER_THIN);
		        	c.setBorderBottom(HSSFCellStyle.BORDER_THIN);	
		        	c.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		        	c.setBorderRight(HSSFCellStyle.BORDER_THIN);
		        	cell.setCellStyle(c);    				     				
      	        	if(cellcount==0) cell.setCellValue("資訊聯絡人員");      	     	
      	        	if(cellcount==1) cell.setCellValue(temp_it_name);   					      	     	      				
      	        	if(cellcount==2){       												
      	        		sheet.addMergedRegion(new Region(rowNum,(short)2,rowNum,(short)3));     			   
      	        		cell.setCellValue(temp_it_telno);
      	        	}
      	        	if(cellcount==5){    					
      	        		sheet.addMergedRegion(new Region(rowNum,(short)5,rowNum,(short)6)); 
      	        		cell.setCellValue("");
      	        	}     		
      	        }//end of cellcount
     		}
     		
     		if(t_1.equals("0")&&t_2.equals("1")){     		     		  
      		    rowNum=rowNum+2;    		  
      		    row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);    		  
      			sheet.addMergedRegion(new Region(rowNum,(short)0,rowNum,(short)3)); 
      			cell = row.createCell( (short)0);
      			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		
      			cell.setCellValue(bank_no+"  "+bank_name);      			
      			System.out.print("TEST............="+bank_no+"  "+bank_name);
      			sheet.addMergedRegion(new Region(rowNum,(short)4,rowNum,(short)7)); 
      			cell = row.createCell( (short)4);
      			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
      			cell.setCellValue(area_id+"  "+addr_1+addr_2);
      			
      			rowNum++;
      			row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
      		
				sheet.addMergedRegion(new Region(rowNum,(short)0,rowNum,(short)3)); 
      			cell = row.createCell( (short)0);
      			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
      			cell.setCellValue(" ");      			
      			
      			sheet.addMergedRegion(new Region(rowNum,(short)4,rowNum,(short)7)); 
      			cell = row.createCell( (short)4);
      			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
      			cell.setCellValue(addr_3);
						
						
				rowNum++;
      			row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
      			sheet.addMergedRegion(new Region(rowNum,(short)0,rowNum,(short)3)); 
      			cell = row.createCell( (short)0);
      			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
      			cell.setCellValue("電話: "+telno);      			
      			
      			sheet.addMergedRegion(new Region(rowNum,(short)4,rowNum,(short)7)); 
      			cell = row.createCell( (short)4);
      			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
      			cell.setCellValue("聯絡信箱:"+email);
						
				rowNum++;
      			row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
      		 
      			sheet.addMergedRegion(new Region(rowNum,(short)0,rowNum,(short)3)); 
      			cell = row.createCell( (short)0);
      			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
      			cell.setCellValue("傳真: "+fax);      			
      			
      			sheet.addMergedRegion(new Region(rowNum,(short)4,rowNum,(short)7)); 
      			cell = row.createCell( (short)4);
      			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
      			cell.setCellValue("網址:"+web_site);
      			      			

      			rowNum=rowNum+1;
      			row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
      			for(int cellcount=0;cellcount<8;cellcount++)
      			{ 
      				     				
      				cell=row.createCell((short)cellcount);
      				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
      				c.setAlignment(HSSFCellStyle.ALIGN_CENTER);
      				//c.setBorderTop(HSSFCellStyle.BORDER_THIN);
      				//c.setBorderBottom(HSSFCellStyle.BORDER_THIN);	
      				c.setBorderLeft(HSSFCellStyle.BORDER_THIN);
      				c.setBorderRight(HSSFCellStyle.BORDER_THIN);
      				cell.setCellStyle(c); 
      				if(cellcount==0) cell.setCellValue("  職稱");      				
      				if(cellcount==1) cell.setCellValue("  姓名");
      				if(cellcount==2){       										
      					sheet.addMergedRegion(new Region(rowNum,(short)2,rowNum,(short)3));     			   
      					cell.setCellValue(" 電話");
      				}
      				if(cellcount==4){
      					cell.setCellValue("   傳真");
      				}      				
      				if(cellcount==5){    					
      					sheet.addMergedRegion(new Region(rowNum,(short)5,rowNum,(short)6)); 
      					cell.setCellValue("    E-MAIL");
      				}     				
      				if(cellcount==7) cell.setCellValue(" 備註");      				 									
      			}   			   			   					  			      					     		
      	    } //end of cell set
      if(t_1.equals("0")&&t_2.equals("2")){
      	rowNum=rowNum+1;      
      	for(int cellcount=0;cellcount<8;cellcount++){   
      		        row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
      		        cell=row.createCell((short)cellcount);
      				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
      				c.setAlignment(HSSFCellStyle.ALIGN_CENTER);
			      	c.setBorderLeft(HSSFCellStyle.BORDER_THIN);
			      	c.setBorderRight(HSSFCellStyle.BORDER_THIN);
			      	cell.setCellStyle(c);    	   				     				
      				if(cellcount==0) cell.setCellValue(addr_1);      				
      				if(cellcount==1) cell.setCellValue(addr_2);   					
      				
      				if(cellcount==2){       										
      					sheet.addMergedRegion(new Region(rowNum,(short)2,rowNum,(short)3));     			   
      					cell.setCellValue(telno);
      				}
      				if(cellcount==4) cell.setCellValue(fax);
      				      				
      				if(cellcount==5){    					
      					sheet.addMergedRegion(new Region(rowNum,(short)5,rowNum,(short)6)); 
      					cell.setCellValue(email);
      				}     				
      				if(cellcount==7){
      					cell.setCellValue(" ");
      				}    									
      	} 
      	flag=0;   
      	temp_t_1=t_1;
      	temp_t_2=t_2;   	     	   			   			   			
      }
      if(!audit_name.equals(" ")) temp_audit_name=audit_name;      
      if(!audit_telno.equals(" ")) temp_audit_telno=audit_telno;      
      if(!it_name.equals(" ")) temp_it_name=it_name;
      if(!it_telno.equals(" ")) temp_it_telno=it_telno;
            
      if(t_1.equals("1")&&t_2.equals("1"))
      { 	
      	if(flag==0)
      	{
      		flag_print=1;      		
      		flag=1;
      		rowNum=rowNum+1;
     		for(int cellcount=0;cellcount<8;cellcount++){
			      	row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
			      	cell=row.createCell((short)cellcount);
			      	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			      	c.setAlignment(HSSFCellStyle.ALIGN_CENTER);			      	
			      	c.setBorderLeft(HSSFCellStyle.BORDER_THIN);
			      	c.setBorderRight(HSSFCellStyle.BORDER_THIN);
			      	cell.setCellStyle(c);
      				if(cellcount==0) cell.setCellValue("稽核聯絡人員");
      				if(cellcount==1) cell.setCellValue(temp_audit_name);
      				if(cellcount==2){ 							
      					sheet.addMergedRegion(new Region(rowNum,(short)2,rowNum,(short)3));     			   
      					cell.setCellValue(temp_audit_telno);
      				}
      				if(cellcount==5){ 	
      					sheet.addMergedRegion(new Region(rowNum,(short)5,rowNum,(short)6)); 
      					cell.setCellValue("");
      				}     		
      	} 
      	
      	rowNum=rowNum+1;
        for(int cellcount=0;cellcount<8;cellcount++)
      	{  
	      			
			      	row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
			      	cell=row.createCell((short)cellcount);
			      	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			      	c.setAlignment(HSSFCellStyle.ALIGN_CENTER);
			      	c.setBorderTop(HSSFCellStyle.BORDER_THIN);
			      	c.setBorderBottom(HSSFCellStyle.BORDER_THIN);	
			      	c.setBorderLeft(HSSFCellStyle.BORDER_THIN);
			      	c.setBorderRight(HSSFCellStyle.BORDER_THIN);
			      	cell.setCellStyle(c);    				     				
      				if(cellcount==0) cell.setCellValue("資訊聯絡人員");      				
      				if(cellcount==1) cell.setCellValue(temp_it_name);
      				if(cellcount==2){       												
      					sheet.addMergedRegion(new Region(rowNum,(short)2,rowNum,(short)3));     			   
      					cell.setCellValue(temp_it_telno);
      				}
      				if(cellcount==5){    					
      					sheet.addMergedRegion(new Region(rowNum,(short)5,rowNum,(short)6)); 
      					cell.setCellValue("");
      				}     		
      	}    	
      }
     
              rowNum=rowNum+2;
      		  row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum); 
      		  
      			sheet.addMergedRegion(new Region(rowNum,(short)0,rowNum,(short)3)); 
      			cell = row.createCell( (short)0);
      			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
      			cell.setCellValue(bank_no+"  "+bank_name);      			
      			
      			sheet.addMergedRegion(new Region(rowNum,(short)4,rowNum,(short)7)); 
      			cell = row.createCell( (short)4);
      			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
      			cell.setCellValue(area_id+"  "+addr_1+addr_2);
      			rowNum++;
      			row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
      			sheet.addMergedRegion(new Region(rowNum,(short)0,rowNum,(short)3)); 
      			cell = row.createCell( (short)0);
      			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
      			cell.setCellValue("");      			
      			
      			sheet.addMergedRegion(new Region(rowNum,(short)4,rowNum,(short)7)); 
      			cell = row.createCell( (short)4);
      			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
      			cell.setCellValue(addr_3);
				rowNum++;
      			row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
      			sheet.addMergedRegion(new Region(rowNum,(short)0,rowNum,(short)3)); 
      			cell = row.createCell( (short)0);
      			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
      			cell.setCellValue("電話: "+telno);      			
      			
      			sheet.addMergedRegion(new Region(rowNum,(short)4,rowNum,(short)7)); 
      			cell = row.createCell( (short)4);
      			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
      			cell.setCellValue("聯絡信箱:"+email);
				rowNum++;
      			row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
      		  	sheet.addMergedRegion(new Region(rowNum,(short)0,rowNum,(short)3)); 
      			cell = row.createCell( (short)0);
      			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
      			cell.setCellValue("傳真: "+fax);      			
      			
      			sheet.addMergedRegion(new Region(rowNum,(short)4,rowNum,(short)7)); 
      			cell = row.createCell( (short)4);
      			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
      			cell.setCellValue("網址:"+web_site);
						
      			rowNum=rowNum+1;
      			row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
      			for(int cellcount=0;cellcount<8;cellcount++)
      			{      				
      				cell=row.createCell((short)cellcount);
      				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
      				c.setAlignment(HSSFCellStyle.ALIGN_CENTER);
      				c.setBorderTop(HSSFCellStyle.BORDER_THIN);
      				c.setBorderBottom(HSSFCellStyle.BORDER_THIN);	
      				c.setBorderLeft(HSSFCellStyle.BORDER_THIN);
      				c.setBorderRight(HSSFCellStyle.BORDER_THIN);
      				cell.setCellStyle(c);
      				if(cellcount==0) cell.setCellValue("  職稱");      				
      				if(cellcount==1) cell.setCellValue("  姓名"); 
      				if(cellcount==2){       										
      					sheet.addMergedRegion(new Region(rowNum,(short)2,rowNum,(short)3));     			   
      					cell.setCellValue(" 電話");
      				}
      				if(cellcount==4) cell.setCellValue("   傳真");      				      				
      				if(cellcount==5){    					
      					sheet.addMergedRegion(new Region(rowNum,(short)5,rowNum,(short)6)); 
      					cell.setCellValue("    E-MAIL");
      				}     				
      				if(cellcount==7) cell.setCellValue(" 備註");      				   									
      			}
      			  	   	
      }
    
       
    if(t_1.equals("1")&&t_2.equals("2"))
     {
     	flag_print=1;
     	rowNum=rowNum+1;      
      	for(int  cellcount=0;cellcount<8;cellcount++)
      	{   
      		    row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
      		    cell=row.createCell((short)cellcount);
      				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
      				c.setAlignment(HSSFCellStyle.ALIGN_CENTER);
			      	c.setBorderTop(HSSFCellStyle.BORDER_THIN);
			      	c.setBorderBottom(HSSFCellStyle.BORDER_THIN);	
			      	c.setBorderLeft(HSSFCellStyle.BORDER_THIN);
			      	c.setBorderRight(HSSFCellStyle.BORDER_THIN);
			      	cell.setCellStyle(c);    	   				     				
      				if(cellcount==0)
      				{     					
      					cell.setCellValue(addr_1);
      				}
      				if(cellcount==1)
      				{  
      					cell.setCellValue(addr_2);   					
      				}      				
      				if(cellcount==2)
      				{       										
      					sheet.addMergedRegion(new Region(rowNum,(short)2,rowNum,(short)3));     			   
      					cell.setCellValue(telno);
      				}
      				if(cellcount==4)
      				{
      					cell.setCellValue(fax);
      				}      				
      				if(cellcount==5)
      				{    					
      					sheet.addMergedRegion(new Region(rowNum,(short)5,rowNum,(short)6)); 
      					cell.setCellValue(email);
      				}     				
      				if(cellcount==7)
      				{
      					cell.setCellValue(" ");
      				}    									
      	} 
       	 	
     }   	   	        
     }//end isSelected
                    			   			
    }//end of dbData size
 
 // add by 2495
    
    
     		
   if(flag_print==0)
   { 
    if(t_1.equals("1") && t_2.equals("2"))
   {
   	     rowNum=rowNum+1;
     		 for(int cellcount=0;cellcount<8;cellcount++)
      	 {  
	      			
			      	row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
			      	cell=row.createCell((short)cellcount);
			      	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			      	c.setAlignment(HSSFCellStyle.ALIGN_CENTER);
			      	//c.setBorderTop(HSSFCellStyle.BORDER_THIN);
			      	//c.setBorderBottom(HSSFCellStyle.BORDER_THIN);	
			      	c.setBorderLeft(HSSFCellStyle.BORDER_THIN);
			      	c.setBorderRight(HSSFCellStyle.BORDER_THIN);
			      	cell.setCellStyle(c);    				     				
      				if(cellcount==0)
      				{     					
      					cell.setCellValue("稽核聯絡人員");
      				}
      				if(cellcount==1)
      				{       					
      					cell.setCellValue(temp_audit_name);   					
      				}      				
      				if(cellcount==2)
      				{       										
      					sheet.addMergedRegion(new Region(rowNum,(short)2,rowNum,(short)3));     			   
      					cell.setCellValue(temp_audit_telno);
      				}
      				if(cellcount==5)
      				{    					
      					sheet.addMergedRegion(new Region(rowNum,(short)5,rowNum,(short)6)); 
      					cell.setCellValue("");
      				}     		
      	} 
      	
      	rowNum=rowNum+1;
        for(int cellcount=0;cellcount<8;cellcount++)
      	{  
	      			
			      	row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
			      	cell=row.createCell((short)cellcount);
			      	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			      	c.setAlignment(HSSFCellStyle.ALIGN_CENTER);
			      	c.setBorderTop(HSSFCellStyle.BORDER_THIN);
			      	c.setBorderBottom(HSSFCellStyle.BORDER_THIN);	
			      	c.setBorderLeft(HSSFCellStyle.BORDER_THIN);
			      	c.setBorderRight(HSSFCellStyle.BORDER_THIN);
			      	cell.setCellStyle(c);    				     				
      				if(cellcount==0)
      				{     					
      					cell.setCellValue("資訊聯絡人員");
      				}
      				if(cellcount==1)
      				{  
      					cell.setCellValue(temp_it_name);   					
      				}      				
      				if(cellcount==2)
      				{       												
      					sheet.addMergedRegion(new Region(rowNum,(short)2,rowNum,(short)3));     			   
      					cell.setCellValue(temp_it_telno);
      				}
      				if(cellcount==5)
      				{    					
      					sheet.addMergedRegion(new Region(rowNum,(short)5,rowNum,(short)6)); 
      					cell.setCellValue("");
      				}     		
      	}    	
   }
  }           
   }
   

   
   
   
                         
}//end if
 else 
    {
    	
    		row = sheet.createRow(rowNum);
    		cell = row.createCell( (short)0);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16);
        cell.setCellValue("尚無資料");  
        f.setFontHeightInPoints((short)12);
		    f.setFontName("標楷體");
		    c.setFont(f);
		    c.setAlignment(HSSFCellStyle.ALIGN_LEFT);
            cell.setCellStyle(c);	   
    }
    
   
		HSSFFooter footer=sheet.getFooter();
	        footer.setCenter( "Page:" +HSSFFooter.page() +" of " +HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
		FileOutputStream fout=new FileOutputStream(reportDir+ System.getProperty("file.separator")+ report_no+".xls");
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



