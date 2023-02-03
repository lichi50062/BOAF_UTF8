/*
  Created on 2005/11/26 by 4180
  96.01.03 fix setup_date/cancel_date == null不轉日期 by 2295
  96.01.22 add 有.無裝設紀錄,都要寫入縣市別.機構代號.機構名稱,拿掉顯示字型 by 2295
  99.04.12 fix 1.因應縣市合併調整SQL 2.修改SQL查詢以pareparedStatement方式 by 2808 
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

public class RptFR025WA {
  	public static String createRpt(String s_year,String s_month,String bank_type,List bank_list){
    	List paramList = new ArrayList();
    	int u_year = 99 ; //縣市合併判斷年分用------------------------
    	s_year = s_year==null?"": s_year.trim() ;
    	if(!"".equals(s_year) && Integer.parseInt(s_year) >=100) {
    		u_year = 100 ;
    	}else {
    		u_year = 99 ;
    	}
    	System.out.println("U_YEAR:"+u_year) ;
    	//END縣市合併判斷年分用
		String errMsg = "";			
		String s_year_last="";
		String s_month_last="";	 
		String filename="";
		String hsien_id ="";
		String hsien_name =""; 
		String bank_no ="";
		String bank_name ="";
		String atm_cnt ="";
		String report_type ="";
		String site_name ="";
		String addr ="";
   		String setup_date="";   
		String cancel_type ="";
		String cancel_date ="";
		String property_no ="";
		String machine_name ="";
		String comment_m ="";
		String setupdate = ""; 
		String canceldate = "";
		filename=(bank_type.equals("6"))?"農漁會信用部ATM裝設台數及異動明細表.xls":"農漁會信用部ATM裝設台數及異動明細表.xls";		
		String bank_typeNm =  "6".equals(bank_type) ? "農會" : "漁會" ; 
		try{				
			System.out.println("xlsDir:"+Utility.getProperties("xlsDir"));
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
    		String openfile="農漁會信用部ATM裝設台數及異動明細表.xls";
    		System.out.println("open file "+openfile);
			FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+ openfile );			
			
	  	    //設定FileINputStream讀取Excel檔
	  		POIFSFileSystem fs = new POIFSFileSystem( finput );
	  		if(fs==null){System.out.println("open 範本檔failed");} else System.out.println("open 範本檔succece");
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
	        wb.setRepeatingRowsAndColumns(0, 1, 12, 2, 3);	  		
	  		finput.close();	  		
	  		
	  		HSSFRow row=null;//宣告一列 
	  		HSSFCell cell=null;//宣告一個儲存格

			HSSFFont ft = wb.createFont();
			HSSFCellStyle cs = wb.createCellStyle();
			HSSFFont ft2 = wb.createFont();
			HSSFCellStyle cs2 = wb.createCellStyle();		     
			 
			row = sheet.createRow(0);
			for(int j=0;j<13;j++){
      		    cell = row.createCell( (short)j);
      		}
      		cell = row.getCell( (short) 5);
      		cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        	cell.setCellValue("全體"+ bank_typeNm +"信用部ATM裝設台數及異動明細表");	
		    ft.setFontHeightInPoints((short)20);
		    ft.setFontName("標楷體");
		    cs.setFont(ft);
		    cs.setAlignment(HSSFCellStyle.ALIGN_CENTER);        	   
        	cell.setCellStyle(cs);  		       
		     
			row= sheet.createRow(1);													
			for(int t=0;t<13;t++){
      		    cell = row.createCell( (short)t);
      		}
      		
      		cell = row.getCell( (short) 5);		
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);			
			cell.setCellValue(" 中華民國"+s_year+"年 " + s_month + "月");									   
        	ft2.setFontHeightInPoints((short)12);
		    ft2.setFontName("標楷體");
		    cs2.setFont(ft2);
		    cs2.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        	cell.setCellStyle(cs2);
	
			int rowNum =4;
			//取得bank_list的bank_no
			String bank_id = "";	
    	    if(bank_list!=null){	
    	    	paramList.add(u_year);
    	    	paramList.add(bank_type);
    	    	paramList.add(s_year);
    	    	paramList.add(s_month);
    	    	paramList.add(u_year);
    	    	paramList.add(bank_type);
    	    	List dbData = DBManager.QueryDB_SQLParam(getReportSQl(u_year),paramList, "hsien_id,hsien_name,bank_no,bank_name,atm_cnt,"+
																				   "report_type,site_name,addr,setup_date,cancel_type,"+
																		   		   "cancel_date,property_no,machine_name,comment_m");
					
				DataObject bean = null;
                if(dbData!=null){
                   String hsien_name_temp ="";
                   String bank_no_temp ="";
                   String bank_name_temp ="";                    
                   for(int k=0;k<dbData.size();k++){
                    	bean = (DataObject)dbData.get(k);                    	
                		hsien_id = String.valueOf(bean.getValue("hsien_id"));
                		hsien_name = String.valueOf(bean.getValue("hsien_name")); 
                		bank_no = String.valueOf(bean.getValue("bank_no"));
                		bank_name = String.valueOf(bean.getValue("bank_name"));
                		atm_cnt = String.valueOf(bean.getValue("atm_cnt"));
                		report_type = String.valueOf(bean.getValue("report_type"));
                		site_name = String.valueOf(bean.getValue("site_name"));
                		addr = String.valueOf(bean.getValue("addr"));
	               		setup_date =String.valueOf(bean.getValue("setup_date"));   
                		cancel_type = String.valueOf(bean.getValue("cancel_type"));
                		cancel_date = "";//String.valueOf(bean.getValue("cancel_date"));
                		property_no = String.valueOf(bean.getValue("property_no"));
                		machine_name = String.valueOf( bean.getValue("machine_name"));
                		comment_m = String.valueOf(bean.getValue("comment_m"));
                		setupdate = ""; 
                		canceldate = "";
                		//96.01.03 fix cancel_date == null的處理 by 2295
                		if(bean.getValue("cancel_date") != null){
                		   cancel_date = String.valueOf(bean.getValue("cancel_date"));
                		}                		
	               		cancel_type = (cancel_type.equals("null"))?" ":cancel_type;
	               		cancel_date = (cancel_date.equals("null"))?"":cancel_date;    	
	               		comment_m = (comment_m.equals("null"))?" ":comment_m;  
                	    if(!cancel_date.equals("") && cancel_date.length() == 6){//96.01.03 fix cancel_date == null的處理 by 2295
                	        canceldate = cancel_date.substring(0,2)+"/"+cancel_date.substring(2,4)+"/"+cancel_date.substring(4,6);
                	    }	
	               		if(!setup_date.equals("") && setup_date.length() == 6){//96.01.03 fix setup_date == null的處理 by 2295
	               	    	//System.out.println("960103.setup_date="+setup_date);
	               	    	setupdate = setup_date.substring(0,2)+"/"+setup_date.substring(2,4)+"/"+setup_date.substring(4,6);	
	               		}
	               		HSSFFont f = wb.createFont();
	               		HSSFCellStyle c = wb.createCellStyle();
	               		//System.out.println("cancel_date="+cancel_date);	   
	               		//System.out.println("setup_date="+setup_date);
	               		//將所有抓到的資料，結合所選取的縣市別來填表	   
	               		List temp ;	  
	               		boolean isSelected = false;    
	               	 	for(int v=0;v<bank_list.size();v++){
                    		temp = (List)(bank_list.get(v));
                    		bank_id = String.valueOf(temp.get(0));
                    		if(bank_id.equals(bank_no)){ 
                    		   isSelected=true;
                    		   break;
                    		}
                    	}        
                    	if(isSelected){
                    	   row = sheet.createRow(rowNum++);                 
                		   for(int cellcount=0;cellcount<12;cellcount++){ 		                 
                     		   cell = row.createCell( (short)cellcount);
                     		   cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                     		   //96.01.22 add 有.無裝設紀錄,都要寫入縣市別.機構代號.機構名稱
                     		   if(cellcount==0 && !hsien_name_temp.equals(hsien_name)){
                     		  	  cell.setCellValue(hsien_name);	
                     		  	  hsien_name_temp = hsien_name;
                     		   }else if(cellcount==1 && !bank_no_temp.equals(bank_no)){ 
                     		   	  cell.setCellValue(bank_no);      
                     		   	  bank_no_temp = bank_no;
                     		   }else if(cellcount==2 && !bank_name_temp.equals(bank_name)){ 
                     		   	  cell.setCellValue(bank_name);	
                     		   	  bank_name_temp = bank_name;
                     		   }
                     		   if(!setup_date.equals("null")){ //有裝設紀錄，寫進各欄位
                     		      if(cellcount==3){ 
                     		      	 cell.setCellValue(site_name);	
                     		      }else if(cellcount==4){ 
                     		      	 cell.setCellValue(addr);
                     		      }else if(cellcount==6){ 
                     		      	 cell.setCellValue(setupdate);   
                		          }else if(cellcount==7){ 
                		          	 cell.setCellValue(cancel_type);		
                    			  }else if(cellcount==8){ 
                    			  	 cell.setCellValue(canceldate);    
                   				  }else if(cellcount==9){ 
                   				  	 cell.setCellValue(property_no);
                				  }else if(cellcount==10){ 
                				  	 cell.setCellValue(machine_name);		
    	    					  }else if(cellcount==11){ 
    	    					  	 cell.setCellValue(comment_m);     	    					  
    	    					  }
	               	              //f.setFontHeightInPoints((short)12);//96.01.22 fix 拿掉顯示字型
	               	              //f.setFontName("標楷體");
	               	              //c.setFont(f);
	               	          	  c.setAlignment(HSSFCellStyle.ALIGN_LEFT);
	               	          	  c.setBorderTop((short)0);
                                  c.setBorderBottom(HSSFCellStyle.BORDER_THIN);
                                  if(cellcount!=5){   //地址欄較長，故佔三欄                       
                                     c.setBorderLeft(HSSFCellStyle.BORDER_THIN);                                     
                                  }   
                                  if(cellcount==11){
                                  	c.setBorderRight(HSSFCellStyle.BORDER_THIN);                                    		
                                  }
                                  cell.setCellStyle(c);	
                       	   	    }else if(setup_date.equals("null")){ //無裝設紀錄，寫進裝設臺數 
                       	           //System.out.println("setup_date == null");
                       	   	  	  if(cellcount==3){
                       	   	  	  	cell.setCellValue("機器台數(歷累計數)");	
                       	   	  	  }else if(cellcount==4){
                       	   	  	  	cell.setCellValue(atm_cnt);        
                       	   	  	  	c.setAlignment(HSSFCellStyle.ALIGN_CENTER);
                       	   	  	  }
                       	   	  	  //f.setFontHeightInPoints((short)12);//96.01.22 fix 拿掉顯示字型
                       	   	  	  //f.setFontName("標楷體");
                       	   	  	  //c.setFont(f);
                       	   	  	  c.setAlignment(HSSFCellStyle.ALIGN_CENTER);
                       	   	  	  c.setBorderTop((short)0);
                                  c.setBorderBottom(HSSFCellStyle.BORDER_THIN); 
                                  if(cellcount<=3){
                                  	 c.setBorderLeft(HSSFCellStyle.BORDER_THIN);                                   	  
                                  }
                                  if(cellcount==3 || cellcount==11){
                                  	 c.setBorderRight(HSSFCellStyle.BORDER_THIN);                                  	
                                  }
                                  cell.setCellStyle(c);     	            	
                       	        }//end of 無裝設紀錄，寫進裝設臺數                
                     	   }//end of cell set.end of for
                     	}//end isSelected
                   }//end of dbData size
                   rowNum = dbData.size()+rowNum;
                }//end of dbData is not null
			}//end if bank_list
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
  	/***
  	 * 取得報表SQL.
  	 * 
  	 * @param u_year
  	 * @return
  	 */
  	private static String getReportSQl (int u_year) {
  		StringBuffer sql = new StringBuffer() ;
  		if(u_year==99) {
  			sql.append(" select hsien_id ,  hsien_name, FR001W_output_order, ");
  			sql.append("        bank_no ,  BANK_NAME,           ");
  			sql.append(" 	   ATM_CNT,   report_type,            ");
  			sql.append(" 	   SITE_Name, ADDR,     ");
  			sql.append(" 	   SETUP_DATE,  CANCEL_TYPE,          ");
  			sql.append(" 	   CANCEL_DATE,  PROPERTY_NO,         ");
  			sql.append(" 	   MACHINE_NAME, Comment_M            ");
  			sql.append(" from(      ");
  			sql.append("       select * from      ");
  			sql.append("       (    ");
  			sql.append("           select nvl(cd01.hsien_id,' ')       as  hsien_id ,       ");
  			sql.append("    nvl(cd01.hsien_name,'OTHER')  as  hsien_name,     ");
  			sql.append("    cd01.FR001W_output_order     as  FR001W_output_order,           ");
  			sql.append("    bn01.bank_no ,  bn01.BANK_NAME,     ");
  			sql.append("           	     ATM_CNT, ");
  			sql.append("           	     'Z99'  as  report_type,");
  			sql.append("           	     ''  AS SITE_Name,  '' as ADDR,       ");
  			sql.append("           	     ''  AS SETUP_DATE,  ''  as  CANCEL_TYPE,           ");
  			sql.append("           	     ''  AS CANCEL_DATE, ''  AS   PROPERTY_NO,          ");
  			sql.append("           	     ''  AS MACHINE_NAME,  '' AS   Comment_M            ");
  			sql.append("           from  (select * from cd01_99 where cd01_99.hsien_id <> 'Y') cd01       ");
  			sql.append("           		 left join wlx01 on wlx01.hsien_id=cd01.hsien_id and m_year = ?    ");
  			sql.append("    left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=? and bn01.m_year = wlx01.m_year  ");
  			sql.append("    left join (select * from WLX05_M_ATM  a01    ");
  			sql.append("           					where a01.m_year=  ? and a01.m_month  = ?)a01 on  bn01.bank_no = a01.bank_no        ");
  			sql.append("       ) Temp1  WHERE ATM_CNT >= 0  AND  BANK_NO  <> ' '       ");
  			sql.append("       union all     ");
  			sql.append("       select * from ");
  			sql.append("       (             ");
  			sql.append("           select nvl(cd01.hsien_id,' ')       as  hsien_id ,  ");
  			sql.append("    nvl(cd01.hsien_name,'OTHER')  as  hsien_name,");
  			sql.append("    cd01.FR001W_output_order     as  FR001W_output_order,      ");
  			sql.append("    bn01.bank_no ,  bn01.BANK_NAME,");
  			sql.append("           	     0   as  ATM_CNT,  ");
  			sql.append("    'A01'  as  report_type,        ");
  			sql.append("           	     nvl(a01.SITE_Name,' ')  as  SITE_Name,        ");
  			sql.append("           	     (xcd01.hsien_name || cd02.AREA_NAME ||  a01.ADDR)  as ADDR, ");
  			sql.append("           	     decode(a01.SETUP_DATE,nvl(a01.SETUP_DATE,''),to_char(to_char(a01.SETUP_DATE,'yyyymmdd')-'19110000'),'')			");
  			sql.append("    as SETUP_DATE,          ");
  			sql.append("           	     decode(a01.CANCEL_TYPE,'1','遷移','2','裁撤','')  as  CANCEL_TYPE, ");
  			sql.append("           	     decode(a01.CANCEL_DATE,nvl(a01.CANCEL_DATE,''),to_char(to_char(a01.CANCEL_DATE,'yyyymmdd')-'19110000'),'')   ");
  			sql.append("    as  CANCEL_DATE,        ");
  			sql.append("           	     nvl(a01.PROPERTY_NO,'' )  as  PROPERTY_NO,           ");
  			sql.append("           	     nvl(a01.MACHINE_NAME,'')  as  MACHINE_NAME,          ");
  			sql.append("           	     nvl(a01.Comment_M,'')     as  Comment_M");
  			sql.append("           from  (select * from cd01_99 where cd01_99.hsien_id <> 'Y') cd01         ");
  			sql.append("           		 left join wlx01 on wlx01.hsien_id=cd01.hsien_id and m_year = ?      ");
  			sql.append("    left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=? and bn01.m_year = wlx01.m_year         ");
  			sql.append("    left join (select * from WLX05_ATM_setup)  a01 on  bn01.bank_no = a01.bank_no   ");
  			sql.append("    left join  (select * from cd01 where cd01.hsien_id <> 'Y') xcd01 on  a01.hsien_id=xcd01.hsien_id            ");
  			sql.append("           		 left join  cd02 on  a01.AREA_ID=cd02.area_id           ");
  			sql.append("       ) Temp2  WHERE SITE_NAME <> ' '    ");
  			sql.append(" ) aaa        ");
  			sql.append(" order by FR001W_output_order, hsien_id, bank_no, report_type, SETUP_DATE           ");
  		}else {
  			sql.append(" select hsien_id ,  hsien_name, FR001W_output_order, ");
	    	sql.append("        bank_no ,  BANK_NAME,                        ");
	    	sql.append(" 	   ATM_CNT,   report_type,                         ");
	    	sql.append(" 	   SITE_Name, ADDR,                                ");
	    	sql.append(" 	   SETUP_DATE,  CANCEL_TYPE,                       ");
	    	sql.append(" 	   CANCEL_DATE,  PROPERTY_NO,                      ");
	    	sql.append(" 	   MACHINE_NAME, Comment_M                         ");
	    	sql.append(" from(                                               ");
	    	sql.append("       select * from                                 ");
	    	sql.append("       (                                             ");
	    	sql.append("           select nvl(cd01.hsien_id,' ')       as  hsien_id ,           ");
	    	sql.append("                  nvl(cd01.hsien_name,'OTHER')  as  hsien_name,         ");
	    	sql.append("                  cd01.FR001W_output_order     as  FR001W_output_order, ");
	    	sql.append("                  bn01.bank_no ,  bn01.BANK_NAME,                       ");
	    	sql.append("           	     ATM_CNT,                                               ");
	    	sql.append("           	     'Z99'  as  report_type,                                ");
	    	sql.append("           	     ''  AS SITE_Name,  '' as ADDR,                         ");
	    	sql.append("           	     ''  AS SETUP_DATE,  ''  as  CANCEL_TYPE,               ");
	    	sql.append("           	     ''  AS CANCEL_DATE, ''  AS   PROPERTY_NO,              ");
	    	sql.append("           	     ''  AS MACHINE_NAME,  '' AS   Comment_M                ");
	    	sql.append("           from  (select * from cd01 where cd01.hsien_id <> 'Y') cd01   ");
	    	sql.append("           		 left join wlx01 on wlx01.hsien_id=cd01.hsien_id and m_year = ? ");
	    	sql.append("                  left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=? and bn01.m_year = wlx01.m_year");
	    	sql.append("                  left join (select * from WLX05_M_ATM  a01 ");
	    	sql.append("           					where a01.m_year=  ? and a01.m_month  = ? )a01 on  bn01.bank_no = a01.bank_no");
	    	sql.append("       ) Temp1  WHERE ATM_CNT >= 0  AND  BANK_NO  <> ' '                                         ");
	    	sql.append("       union all                                                                                 ");
	    	sql.append("       select * from                                                                             ");
	    	sql.append("       (                                                                                         ");
	    	sql.append("           select nvl(cd01.hsien_id,' ')       as  hsien_id ,                                    ");
	    	sql.append("                  nvl(cd01.hsien_name,'OTHER')  as  hsien_name,                                  ");
	    	sql.append("                  cd01.FR001W_output_order     as  FR001W_output_order,                          ");
	    	sql.append("                  bn01.bank_no ,  bn01.BANK_NAME,                                                ");
	    	sql.append("           	     0   as  ATM_CNT,                                                                ");
	    	sql.append("                  'A01'  as  report_type,                                                        ");
	    	sql.append("           	     nvl(a01.SITE_Name,' ')  as  SITE_Name,                                          ");
	    	sql.append("           	     (xcd01.hsien_name || cd02.AREA_NAME ||  a01.ADDR)  as ADDR,                     ");
	    	sql.append("           	     decode(a01.SETUP_DATE,nvl(a01.SETUP_DATE,''),to_char(to_char(a01.SETUP_DATE,'yyyymmdd')-'19110000'),'')");
	    	sql.append("                  as SETUP_DATE,                                                                                        ");
	    	sql.append("           	     decode(a01.CANCEL_TYPE,'1','遷移','2','裁撤','')  as  CANCEL_TYPE,                                     ");
	    	sql.append("           	     decode(a01.CANCEL_DATE,nvl(a01.CANCEL_DATE,''),to_char(to_char(a01.CANCEL_DATE,'yyyymmdd')-'19110000'),'')");
	    	sql.append("                  as  CANCEL_DATE,                                                                                         ");
	    	sql.append("           	     nvl(a01.PROPERTY_NO,'' )  as  PROPERTY_NO,                                                                ");
	    	sql.append("           	     nvl(a01.MACHINE_NAME,'')  as  MACHINE_NAME,                                                               ");
	    	sql.append("           	     nvl(a01.Comment_M,'')     as  Comment_M                                                                   ");
	    	sql.append("           from  (select * from cd01 where cd01.hsien_id <> 'Y') cd01                                                      ");
	    	sql.append("           		 left join wlx01 on wlx01.hsien_id=cd01.hsien_id and m_year = ?                                            ");
	    	sql.append("                  left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type= ?  and bn01.m_year = wlx01.m_year      ");
	    	sql.append("                  left join (select * from WLX05_ATM_setup)  a01 on  bn01.bank_no = a01.bank_no                            ");
	    	sql.append("                  left join  (select * from cd01 where cd01.hsien_id <> 'Y') xcd01 on  a01.hsien_id=xcd01.hsien_id         ");
	    	sql.append("           		 left join  cd02 on  a01.AREA_ID=cd02.area_id                                                                ");
	    	sql.append("       ) Temp2  WHERE SITE_NAME <> ' '                                                                                     ");
	    	sql.append(" ) aaa   ");
	    	sql.append(" order by FR001W_output_order, hsien_id, bank_no, report_type, SETUP_DATE ");
  		}
  		return sql.toString() ;
  	}
}