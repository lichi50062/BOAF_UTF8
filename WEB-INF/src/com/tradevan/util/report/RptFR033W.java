/*
  Created on 2006/01/03 by 4180
  96.11.21 add 委外項目.委外範圍;報表區分以縣市別排序、以委外項目別排序 by 2295
  97.07.09 fix 結束日期.可不輸入.值為空白 by 2295
  99.04.12 fix 因應縣市合併調整SQL & 查詢方式改以preparedstatement by 2808
*/
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.Region;

import java.io.*;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR033W {
  	public static String createRpt(String bank_type,List bank_list,String rpt_type,String s_year) 
	{
  		String u_year = "100" ;
  		if(s_year==null || Integer.parseInt(s_year)<=99) {
  			u_year = "99" ;
  		}
  		
		DataObject bean = null; 
		List paramList = new ArrayList();
		String errMsg = "";
		String s_year_last="";
		String s_month_last="";
		String bank_type_name=(bank_type.equals("6"))?"農會":"漁會";
		String filename="全體農漁會委外催收委外之對象一覽表_"+(rpt_type.equals("01")?"縣市別":"委外項目")+".xls";
		String[] out_item_name = {"","一","二","三","四","五","六","七","八","九","十",
   			    				  "十一","十二","十三","十四","十五","十六","十七","十八","十九","二十"};
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
    		
    		System.out.println("open file "+filename);
			FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+ filename );

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
			
			HSSFFont f = wb.createFont();
			HSSFCellStyle c = wb.createCellStyle();
			HSSFCellStyle c_center = wb.createCellStyle();
      		row = sheet.getRow(0);
      		cell = row.getCell( (short) 0);
      		
      		cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        	cell.setCellValue("全體"+ bank_type_name +"委外內部作業資料一覽表");
        	
		    ft.setFontHeightInPoints((short)18);
		    ft.setFontName("標楷體");
		    cs.setFont(ft);
		    cs.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        	cell.setCellStyle(cs);
        	//內容的字型.可跳行設定
        	f.setFontHeightInPoints((short)12);
            f.setFontName("標楷體");
            c.setFont(f);
            c.setAlignment(HSSFCellStyle.ALIGN_LEFT);
            c.setBorderTop((short)0);
            c.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            c.setBorderLeft(HSSFCellStyle.BORDER_THIN);
            c.setWrapText(true);
            
            c_center.setFont(f);
            c_center.setAlignment(HSSFCellStyle.ALIGN_CENTER);
            c_center.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER );            
            c_center.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            c_center.setBorderLeft(HSSFCellStyle.BORDER_THIN);
            c_center.setWrapText(true);
            
			row= sheet.createRow(1);
			for(int t=0;t<11;t++){
      		 cell = row.createCell( (short)t);
      		}
			
	        int rowNum =4;
	        //取得bank_list的bank_no
	        String bank_id = "";
	        StringBuffer sql = new StringBuffer() ;
     		if(bank_list!=null){
     			System.out.println("列印方式BY 縣市別");
     		   //================================================================================================================
     		   if("01".equals(rpt_type)){//以縣市別區分
     			    if("100".equals(u_year) ) {
	     			    sql.append(" select Temp1.* from (");                                                  					  
		       			sql.append(" select nvl(cd01.hsien_id,' ')        as  hsien_id ,            ");                                              
		       			sql.append("        nvl(cd01.hsien_name,'OTHER')  as  hsien_name,           ");                                         
		       			sql.append("        cd01.FR001W_output_order      as  FR001W_output_order,  ");                                         
		       			sql.append("        bn01.bank_no ,  bn01.BANK_NAME,OutCompanyName, OutContractName, OutContractTel,          ");        
		       			sql.append(" 	   ltrim((substr(lpad((to_char(OUT_Begin_DATE,'yyyymmdd')-19110000),7,'0'),1,3) || '/' ||      ");     
		       			sql.append(" 	   substr(lpad((to_char(OUT_Begin_DATE,'yyyymmdd')-19110000),7,'0'),4,2) || '/' ||             ");     
		       			sql.append(" 	   substr(lpad((to_char(OUT_Begin_DATE,'yyyymmdd')-19110000),7,'0'),6,2)),'0')  as OUT_Begin_DATE,   ");
		       			sql.append("        decode(OUT_End_DATE,'','',ltrim((substr(lpad((to_char(OUT_End_DATE,'yyyymmdd')-19110000),7,'0'),1,3) || '/'  ||  ");          
		       			sql.append("        substr(lpad((to_char(OUT_End_DATE,'yyyymmdd')-19110000),7,'0'),4,2) || '/'  ||                                   ");
		       			sql.append("        substr(lpad((to_char(OUT_End_DATE,'yyyymmdd')-19110000),7,'0'),6,2)),'0'))  as OUT_End_DATE,                     ");
		       			sql.append("        BankComplainName,BankComplainTel,OutComment,out_item,out_range                                                   ");
		       			sql.append(" from  ");
		       			sql.append(" ( select * from cd01 where cd01.hsien_id <> 'Y') cd01 ");
		       			sql.append(" Left join (select * from wlx01 where m_year= ? )wlx01 on wlx01.hsien_id=cd01.hsien_id ");     
		       			sql.append(" Left join (select * from bn01 where bank_type=? and m_year=? ) bn01 on wlx01.bank_no=bn01.bank_no ");   
		       			sql.append(" Left join WLX06_M_OUTPUSH  a01 on  bn01.bank_no = a01.bank_no ");
		       			sql.append(" ) Temp1 , v_bank_location T2  ");
		       			sql.append(" WHERE OutCompanyName <> ' ' ") ;
		       			sql.append(" And Temp1.bank_no in (").append(getSelectBankNo(bank_list)).append(")");
		       			sql.append(" AND Temp1.bank_No = T2.bank_No AND T2.m_year =? ");
		       			sql.append(" Order by T2.FR001W_output_order, Temp1.hsien_id, Temp1.bank_no, Temp1.OUT_Begin_DATE    ");
		       			paramList.add(u_year) ;
           			    paramList.add(bank_type) ;
           			    paramList.add(u_year) ;
           			    paramList.add(u_year) ;
     			    }else {
     			    	String cd01Table = "cd01_99" ;
     			    	if("100".equals(u_year) ) {
     			    		cd01Table = "cd01" ;
     			    	}
     			    	sql.append(" select Temp1.* from (         ");                                                             					  
     			    	sql.append(" select ");
     			    	sql.append(" nvl(cd01.hsien_id,' ') As  hsien_id ");                                          
     			    	sql.append(" ,nvl(cd01.hsien_name,'OTHER')  as  hsien_name");                                          
     			    	sql.append(" ,cd01.FR001W_output_order  As  FR001W_output_order ");                                          
     			    	sql.append(" ,bn01.bank_no ,  bn01.BANK_NAME,OutCompanyName, OutContractName, OutContractTel");             
     			    	sql.append(" ,Ltrim((substr(lpad((to_char(OUT_Begin_DATE,'yyyymmdd')-19110000),7,'0'),1,3) || '/' ||                             ");
     			    	sql.append(" 	   substr(lpad((to_char(OUT_Begin_DATE,'yyyymmdd')-19110000),7,'0'),4,2) || '/' ||                                    ");
     			    	sql.append(" 	   substr(lpad((to_char(OUT_Begin_DATE,'yyyymmdd')-19110000),7,'0'),6,2)),'0')  as OUT_Begin_DATE,                    ");
     			    	sql.append("        decode(OUT_End_DATE,'','',ltrim((substr(lpad((to_char(OUT_End_DATE,'yyyymmdd')-19110000),7,'0'),1,3) || '/'  || ");           
     			    	sql.append("        substr(lpad((to_char(OUT_End_DATE,'yyyymmdd')-19110000),7,'0'),4,2) || '/'  ||                                  ");
     			    	sql.append("        substr(lpad((to_char(OUT_End_DATE,'yyyymmdd')-19110000),7,'0'),6,2)),'0'))  as OUT_End_DATE,                    ");
     			    	sql.append("        BankComplainName,BankComplainTel,OutComment,out_item,out_range                                                  ");
     			    	sql.append(" from ");
     			    	sql.append(" (select * from ").append(cd01Table).append(" where hsien_id <> 'Y') cd01 ");
     			    	sql.append("  Left join (select * from wlx01 where m_year=? ) wlx01 on wlx01.hsien_id=cd01.hsien_id ");     
     			    	sql.append("  Left join (select * from bn01 where bank_type=? and m_year=? ) bn01 on wlx01.bank_no=bn01.bank_no  ");    
     			    	sql.append("  Left join WLX06_M_OUTPUSH  a01 on  bn01.bank_no = a01.bank_no                                                   ");
     			    	sql.append(" ) Temp1  , v_bank_location T2 ");
     			    	sql.append(" WHERE   Temp1.OutCompanyName <> ' ' ");
     			    	sql.append("         and Temp1.bank_no in (").append(getSelectBankNo(bank_list)).append(")");   
     			    	sql.append(" AND Temp1.bank_No = T2.bank_No AND T2.m_year =? ");
     			    	sql.append("  order by T2.FR001W_output_order, Temp1.hsien_id, Temp1.bank_no, Temp1.OUT_Begin_DATE ");
     			    	paramList.add(u_year) ;
           			    paramList.add(bank_type) ;
           			    paramList.add(u_year) ;
           			    paramList.add(u_year) ;
     			    }
       			    
       			    System.out.println("====this sql is :"+sql.toString()) ;
     		   }else{//以委外項目做區分
     			  sql.append(" select T1.*,T2.FR001W_output_order from (  ") ;
     			  sql.append(" select bn01.bank_name,a.*, b.countdata ") ;
     			  sql.append(" from (select bank_no,out_item,seq_no,outcompanyname,outcontractname, ");
				  sql.append("	   		  outcontracttel,bankcomplainname,bankcomplaintel,");
	   		      sql.append("	   		  outcomment,user_id,user_name,update_date,out_range,");
	   		      sql.append(" 	     	  ltrim((substr(lpad((to_char(OUT_Begin_DATE,'yyyymmdd')-19110000),7,'0'),1,3) || '/' ||          ");
			      sql.append(" 	     	  substr(lpad((to_char(OUT_Begin_DATE,'yyyymmdd')-19110000),7,'0'),4,2) || '/' ||                 ");
			      sql.append(" 	     	  substr(lpad((to_char(OUT_Begin_DATE,'yyyymmdd')-19110000),7,'0'),6,2)),'0')  as OUT_Begin_DATE, ");
			      sql.append("        	  decode(OUT_End_DATE,'','',ltrim((substr(lpad((to_char(OUT_End_DATE,'yyyymmdd')-19110000),7,'0'),1,3) || '/'  ||           ");
			      sql.append("        	  substr(lpad((to_char(OUT_End_DATE,'yyyymmdd')-19110000),7,'0'),4,2) || '/'  ||                  ");
			      sql.append("        	  substr(lpad((to_char(OUT_End_DATE,'yyyymmdd')-19110000),7,'0'),6,2)),'0'))  as OUT_End_DATE      ");
			      sql.append("	   from WLX06_M_OUTPUSH ");
		   		  sql.append("       order by out_item,seq_no)a ");
		   		  sql.append(" left join (select bank_no,out_item, ");
		   		  sql.append(" 				   count(*) as countdata ");  
		   		  sql.append(" 		    from WLX06_M_OUTPUSH "); 
		   		  sql.append("	   	    group by bank_no,out_item)b on a.bank_no=b.bank_no and a.out_item=b.out_item ");
				  sql.append(" Left join (select * from bn01 where bank_type=? and m_year=? ) bn01 on a.bank_no=bn01.bank_no  ");
				  sql.append(" where bank_name is not null ");
				  sql.append("       and a.bank_no in (").append(getSelectBankNo(bank_list)).append(")"); 
				  sql.append(" ) T1 ,  v_bank_location T2 ");
				  sql.append(" Where T1.Bank_No = T2.Bank_No And T2.m_year=?  ");
				  sql.append(" Order by T2.FR001W_output_order ");
				  paramList.add(bank_type) ;
				  paramList.add(u_year) ;
				  paramList.add(u_year) ;
     		   }
     		  
               //==========================================================================================================
     		   List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "hsien_id,hsien_name,bank_no,bank_name,outcompanyname,"+
											          "outcontractname,outcontracttel,out_begin_date,out_end_date"+
											          "bankcomplainname,bankcomplaintel,outcomment,out_item,out_range,countdata");
			   
               System.out.println("抓出的dbData.size="+String.valueOf(dbData.size()) );
               
     		   if(dbData!=null){
     			  String bank_name_temp ="",hsien_id = "",hsien_name = "";
                  String bank_no = "",bank_name = "",outcompanyname = "";
                  String outcontractname = "",outcontracttel = "", out_begin_date = "";
                  String out_end_date = "",bankcomplainname ="", bankcomplaintel = "";
                  String outcomment = "";
                  String out_item = "",out_range = "";
                  String count_data = "",pre_out_item="";	
                  int begin_row = 0,end_row=0;
                  bean = (DataObject)dbData.get(0); 
                  pre_out_item = (bean.getValue("out_item")==null)?"": String.valueOf(bean.getValue("out_item"));
                  begin_row = rowNum;
                  for(int k=0;k<dbData.size();k++){
                   	  bean = (DataObject)dbData.get(k);                
 	              	  hsien_id = String.valueOf(bean.getValue("hsien_id"));
 	              	  bank_no = String.valueOf(bean.getValue("bank_no"));
 	              	  
 	              	  hsien_name = String.valueOf(bean.getValue("hsien_name"));
 	              	 
 	              	  bank_name = String.valueOf(bean.getValue("bank_name"));
 	              	 
 	              	  outcompanyname = String.valueOf(bean.getValue("outcompanyname"));
 	              	  
 	              	  outcontractname = String.valueOf(bean.getValue("outcontractname"));
 	              	  
 	              	  outcontracttel = String.valueOf(bean.getValue("outcontracttel"));
 	              	  
 	              	  out_begin_date = String.valueOf(bean.getValue("out_begin_date"));
 	              	  
	              	  out_end_date = String.valueOf(bean.getValue("out_end_date"));
	              	  
	              	  bankcomplainname = String.valueOf(bean.getValue("bankcomplainname"));
	              	  
	              	  bankcomplaintel = String.valueOf(bean.getValue("bankcomplaintel"));
	              	  
	              	  outcomment = String.valueOf(bean.getValue("outcomment"));  
	              	              
	           		  outcomment = (outcomment.equals("null"))?" ":outcomment;
	           		  
	           		  //96.11.20 add 委外項外.委外範圍
					  out_item = "null".equals(String.valueOf(bean.getValue("out_item")))? "0" : String.valueOf(bean.getValue("out_item"));	
					  
					  out_range = "null".equals(String.valueOf(bean.getValue("out_range")))?"" : String.valueOf(bean.getValue("out_range"));
					  
					  count_data = String.valueOf(bean.getValue("count_data"));				  
					  
		              
		              
     	           	  
 	           		 row = sheet.createRow(rowNum++);
	  			     for(int cellcount=0;cellcount<11;cellcount++){
  		                 cell = row.createCell( (short)cellcount);
  		                 cell.setEncoding(HSSFCell.ENCODING_UTF_16);
  		                 if(rpt_type.equals("01")){//縣市別區分
  		                 	if(cellcount==0 && !bank_name_temp.equals(bank_name)){//農漁會名稱 
  		                 		cell.setCellValue(bank_name);	
  		                 		bank_name_temp = bank_name;
  		                 	}else if(cellcount==1)//委外項目
  		                 		cell.setCellValue(out_item_name[Integer.parseInt(out_item)]);
  		                 	else if(cellcount==2)//受委託機構-機構名稱				  
  		                 		cell.setCellValue(outcompanyname);
  		                 	else if(cellcount==3)//受委託機構-聯絡人
  		                 		cell.setCellValue(outcontractname);
  		                 	else if(cellcount==4)//受委託機構-聯絡電話
  		                 		cell.setCellValue(outcontracttel);
  		                 	else if(cellcount==5)//受委託機構-受託起始日期
  		                 		cell.setCellValue(out_begin_date);
  		                 	else if(cellcount==6){//受委託機構-受託結束日期
  		                 		cell.setCellValue(out_end_date);      		                 		
  		                 	}else if(cellcount==7)//委外事項範圍
  		                 		cell.setCellValue(out_range);        	                 
  		                 	else if(cellcount==8)//信用部申訴窗口-聯絡人
  		                 		cell.setCellValue(bankcomplainname);
  		                 	else if(cellcount==9)//信用部申訴窗口-專線電話
  		                 		cell.setCellValue(bankcomplaintel);
  		                 	else if(cellcount==10)//備註
  		                 		cell.setCellValue(outcomment);
  		                 }else{//委外項目區分      		                 	
  		                 	if(cellcount==0){//委外項目
  		                 		if(!out_item.equals("")){
  		                 		   cell.setCellValue(out_item_name[Integer.parseInt(out_item)]);
  		                 		}      		                 		
  		                 	}else if(cellcount==1)//農漁會名稱
  		                 		cell.setCellValue(bank_name);
  		                 	else if(cellcount==2)//受委託機構-機構名稱				  
  		                 		cell.setCellValue(outcompanyname);
  		                 	else if(cellcount==3)//受委託機構-聯絡人	
  		                 		cell.setCellValue(outcontractname);
  		                 	else if(cellcount==4)//受委託機構-聯絡電話
  		                 		cell.setCellValue(outcontracttel);
  		                 	else if(cellcount==5)//受委託機構-受託起始日期
  		                 		cell.setCellValue(out_begin_date);
  		                 	else if(cellcount==6)//受委託機構-受託結束日期
  		                 		cell.setCellValue(out_end_date);
  		                 	else if(cellcount==7)//委外事項範圍
  		                 		cell.setCellValue(out_range);        	                 
  		                 	else if(cellcount==8)//信用部申訴窗口-聯絡人
  		                 		cell.setCellValue(bankcomplainname);
  		                 	else if(cellcount==9)//信用部申訴窗口-專線電話
  		                 		cell.setCellValue(bankcomplaintel);
  		                 	else if(cellcount==10)//備註
  		                 		cell.setCellValue(outcomment);      		                 	
  		                 	
  		                 	if(!pre_out_item.equals(out_item)){
  		                 		end_row = rowNum - 2;
  		                 		System.out.println("begin_row="+begin_row+":end_row="+end_row);      		                 		
  		                 		sheet.addMergedRegion( new Region( ( short )begin_row, ( short )0,
  	                                ( short )end_row,
  	                                ( short )0) );
  		                 		begin_row = rowNum -1;
  		                 		pre_out_item = out_item;
  		                    }
  		                 }
                         if(cellcount==10)
                            c.setBorderRight(HSSFCellStyle.BORDER_THIN);
                         cell.setCellStyle(c);
                         if(cellcount==0) cell.setCellStyle(c_center);
  			 		 }//end of cell set
      				 
                  }//end of dbData size
               	  rowNum = dbData.size()+rowNum;
		       }//end of data is not null
		    }else{//bank_list ==null
    		   row = sheet.createRow(rowNum);
    		   cell = row.createCell( (short)0);
               cell.setEncoding(HSSFCell.ENCODING_UTF_16);
               cell.setCellValue("尚無資料");
               f.setFontHeightInPoints((short)12);
		       f.setFontName("標楷體");
		       c.setFont(f);
		       c.setAlignment(HSSFCellStyle.ALIGN_LEFT);
               cell.setCellStyle(c);
    	    }//end of bank_list

			HSSFFooter footer=sheet.getFooter();
	        footer.setCenter( "Page:" +HSSFFooter.page() +" of " +HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
			FileOutputStream fout=new FileOutputStream(reportDir+ System.getProperty("file.separator")+ filename);
			wb.write(fout);
	        //儲存
	        fout.close();
	        System.out.println("儲存完成");

		}catch(Exception e){
			System.out.println("createRpt Error:"+e.getMessage());
			System.out.println(e) ;
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