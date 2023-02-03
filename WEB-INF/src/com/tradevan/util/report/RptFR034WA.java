/*
	95.01.04 add 明細表 by 4180
    99.04.13 fix 縣市合併SQL調整 && 修改查詢方式為preparedstatement by 2808
   103.01.16 add 臺灣省農會更名為中華民國農會增加說明 by 2295
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

public class RptFR034WA {
  	public static String createRpt(String s_year,String s_month,String unit,String bank_type){
	    String errMsg = "";
		
		
		String bank_type_name=(bank_type.equals("6"))?"農會":"漁會";
		String unit_name =  Utility.getUnitName(unit);//取得單位名稱
		String u_year = "100" ; //縣市合併判斷用 
        if(s_year==null || Integer.parseInt(s_year) < 100) {
        	u_year = "99" ;
        }
        System.out.println("query_year:"+s_year) ;
        System.out.println("u_year:"+u_year) ;
		String filename="農漁會信用部警示帳戶調查統計總表.xls";
			
		reportUtil reportUtil = new reportUtil();
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
    		 String openfile="農漁會信用部警示帳戶調查統計總表.xls";
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
	  		 //wb.setSheetName(0,"test");
	  		 //設定表頭 為固定 先設欄的起始再設列的起始
	         wb.setRepeatingRowsAndColumns(0, 1, 8, 2, 4);
	  		 
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
        	 cell.setCellValue("全體"+ bank_type_name +"信用部警示帳戶調查統計總表");	
		 
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
		
        	 StringBuffer sql = new StringBuffer() ;
        	 List paramList = new ArrayList() ;
        	 sql.append(getReportSQL(u_year)) ;
        	 
        	 paramList.add(unit) ;
        	 paramList.add(unit) ;
        	 paramList.add(unit) ;
        	 paramList.add(u_year) ;
        	 paramList.add(bank_type) ;
        	 paramList.add(u_year) ;
        	 paramList.add(s_year) ;
        	 paramList.add(s_month) ;
        	 paramList.add(unit) ;
        	 paramList.add(unit) ;
        	 paramList.add(unit) ;
        	 paramList.add(u_year) ;
        	 paramList.add(bank_type) ;
        	 paramList.add(u_year) ;
        	 paramList.add(s_year) ;
        	 paramList.add(s_month) ;
        	 List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"hsien_id,hsien_name,warnaccount_tcnt"+
									  "warnaccount_tbal,warnaccount_remit_tcnt,warnaccount_refund_apply_cnt"+
									  "warnaccount_refund_apply_amt,warnaccount_refund_cnt,warnaccount_refund_amt");
		    
			HSSFFont f = wb.createFont();
			HSSFCellStyle c = wb.createCellStyle();
			int rowNum=5;
			DataObject bean = (DataObject)dbData.get(0);
			String setValue = "";
	     	
	 		String hsien_id = String.valueOf(bean.getValue("hsien_id"));
	 		String hsien_name = String.valueOf(bean.getValue("hsien_name")); 
	 		String warnaccount_tcnt =  Utility.setCommaFormat(String.valueOf(bean.getValue("warnaccount_tcnt")));
			String warnaccount_tbal =  Utility.setCommaFormat(String.valueOf(bean.getValue("warnaccount_tbal")));
			String warnaccount_remit_tcnt =  Utility.setCommaFormat(String.valueOf(bean.getValue("warnaccount_remit_tcnt")));
			String warnaccount_refund_apply_cnt =  Utility.setCommaFormat(String.valueOf(bean.getValue("warnaccount_refund_apply_cnt")));
			String warnaccount_refund_apply_amt =  Utility.setCommaFormat(String.valueOf(bean.getValue("warnaccount_refund_apply_amt")));
			String warnaccount_refund_cnt =  Utility.setCommaFormat(String.valueOf(bean.getValue("warnaccount_refund_cnt")));
			String warnaccount_refund_amt =  Utility.setCommaFormat(String.valueOf(bean.getValue("warnaccount_refund_amt")));
	
		    //印出合計資料 
		    row = sheet.createRow(rowNum);  
		    for(int cellcount=0;cellcount<8;cellcount++){
		    	cell = row.createCell( (short)cellcount);
      		    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
      		            
  		        if(cellcount==0)
    	          setValue = hsien_name;
    	        else if(cellcount==1)
    	          setValue = warnaccount_tcnt;   
    	        else if(cellcount==2)
    	          setValue = warnaccount_tbal;	
     	        else if(cellcount==3)
    	          setValue = warnaccount_remit_tcnt;	
                else if(cellcount==4)
    	          setValue = warnaccount_refund_apply_cnt;
    	        else if(cellcount==5)
                  setValue = warnaccount_refund_apply_amt;   
    	        else if(cellcount==6)
    	          setValue = warnaccount_refund_cnt;		
  		        else if(cellcount==7)
    	          setValue = warnaccount_refund_amt;    
        	    
    	       	setValue = ( setValue.equals("null"))?"0":setValue;
    	       	cell.setCellValue(setValue);
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
              }          
              //印出各縣市資料 
              rowNum++;               
			  for(int rowcount=1;rowcount<dbData.size();rowcount++) {
				  //從1開始取得資料===================
				  bean = (DataObject)dbData.get(rowcount);
     	          
	 		      hsien_id = String.valueOf(bean.getValue("hsien_id"));
	 		      hsien_name = String.valueOf(bean.getValue("hsien_name"));   
	 		      warnaccount_tcnt =  Utility.setCommaFormat(String.valueOf(bean.getValue("warnaccount_tcnt")));
			      warnaccount_tbal =  Utility.setCommaFormat(String.valueOf(bean.getValue("warnaccount_tbal")));
			      warnaccount_remit_tcnt =  Utility.setCommaFormat(String.valueOf(bean.getValue("warnaccount_remit_tcnt")));
			      warnaccount_refund_apply_cnt =  Utility.setCommaFormat(String.valueOf(bean.getValue("warnaccount_refund_apply_cnt")));
			      warnaccount_refund_apply_amt =  Utility.setCommaFormat(String.valueOf(bean.getValue("warnaccount_refund_apply_amt")));
			      warnaccount_refund_cnt =  Utility.setCommaFormat(String.valueOf(bean.getValue("warnaccount_refund_cnt")));
			      warnaccount_refund_amt =  Utility.setCommaFormat(String.valueOf(bean.getValue("warnaccount_refund_amt")));
		          
			      row = sheet.createRow(rowNum);			       
			      for(int cellcount=0;cellcount<8;cellcount++){
					 cell = row.createCell( (short)cellcount);
	  		         cell.setEncoding(HSSFCell.ENCODING_UTF_16);
  		          
	  		         if(cellcount==0)
	    	           { setValue = hsien_name;}
	    	       	 else if(cellcount==1)
	    	             setValue = warnaccount_tcnt;   
	    	         else if(cellcount==2)
	    	           setValue = warnaccount_tbal;	
	     	         else if(cellcount==3)
	    	           setValue = warnaccount_remit_tcnt;	
	                 else if(cellcount==4)
	    	           setValue = warnaccount_refund_apply_cnt;
	    	         else if(cellcount==5)
	                   setValue = warnaccount_refund_apply_amt;   
	    	         else if(cellcount==6)
	    	           setValue = warnaccount_refund_cnt;		
	  		         else if(cellcount==7)
	    	             setValue = warnaccount_refund_amt;    
        	      
	        	     setValue = ( setValue.equals("null"))?"0":setValue;
	        	     cell.setCellValue(setValue);
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
                  }          
                  rowNum++;               
			  }		        	
			  
              row = sheet.createRow(rowNum);
              cell = row.createCell( (short)0);
              cell.setEncoding(HSSFCell.ENCODING_UTF_16);
              cell.setCellValue("縣市別欄之其他(農會)係指原臺灣省農會，該農會於102年5月22日更名為中華民國農會。");            
						
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
  	 * 取得總表SQL BY 年度
  	 * @param u_year
  	 * @return
  	 */
  	private static String getReportSQL(String u_year) {
  		StringBuffer sql = new StringBuffer() ;
  		if("99".equals(u_year)) {
  			 sql.append(" select ' '    as hsien_id ,  "); //--合計
	       	 sql.append("        '合計' as hsien_name, "); 
	       	 sql.append("        '000'  as FR001W_output_order,      "); 
	       	 sql.append("        sum(a01.WarnAccount_TCnt)  as WarnAccount_TCnt,              ");
	       	 sql.append("        round(sum(a01.WarnAccount_Tbal)/?,0)  as  WarnAccount_Tbal,  ");//1.unit
	       	 sql.append("        Sum(a01.WarnAccount_Remit_TCnt)   as WarnAccount_Remit_TCnt, ");
	       	 sql.append("        Sum(a01.WarnAccount_Refund_Apply_Cnt)   as WarnAccount_Refund_Apply_Cnt,              ");
	       	 sql.append("        round(sum(a01.WarnAccount_Refund_Apply_Amt)/ ? ,0)  as  WarnAccount_Refund_Apply_Amt,  ");//2.unit
	       	 sql.append("        sum(a01.WarnAccount_Refund_Cnt)  as WarnAccount_Refund_Cnt,                           ");
	       	 sql.append("        round(sum(a01.WarnAccount_Refund_Amt) /?,0)  as  WarnAccount_Refund_Amt               ");//3.unit
	       	 sql.append(" from  (select * from cd01_99 cd01 where cd01.hsien_id <> 'Y') cd01                           ");
	       	 sql.append(" left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = ?                        ");//4.u_year
	       	 sql.append(" left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type= ? and bn01.m_year = wlx01.m_year and bn01.m_year = ? ");//5.u_year
	       	 sql.append(" left join (select * from WLX09_S_WARNING where m_year  = ?  and m_quarter = ? ) a01 on  bn01.bank_no = a01.bank_no      ");
	       	 sql.append(" union     ");
	       	 sql.append(" select  * ");//--各縣市合計
	       	 sql.append(" from (    ");
	       	 sql.append("       select nvl(cd01.hsien_id,' ')       as  hsien_id ,                     ");
	       	 sql.append("              nvl(cd01.hsien_name,'OTHER') as  hsien_name,                    ");
	       	 sql.append("              cd01.FR001W_output_order     as  FR001W_output_order,           ");
	       	 sql.append("              sum(a01.WarnAccount_TCnt)    as WarnAccount_TCnt,               ");
	       	 sql.append("              round(sum(a01.WarnAccount_Tbal)/ ?,0)  as  WarnAccount_Tbal,     ");
	       	 sql.append("              Sum(a01.WarnAccount_Remit_TCnt)   as WarnAccount_Remit_TCnt,    ");
	       	 sql.append("              Sum(a01.WarnAccount_Refund_Apply_Cnt)   as WarnAccount_Refund_Apply_Cnt,           ");
	       	 sql.append("              round(sum(a01.WarnAccount_Refund_Apply_Amt)/?,0)  as  WarnAccount_Refund_Apply_Amt,");
	       	 sql.append("              sum(a01.WarnAccount_Refund_Cnt)  as WarnAccount_Refund_Cnt,                        ");
	       	 sql.append("              round(sum(a01.WarnAccount_Refund_Amt) /?,0)  as  WarnAccount_Refund_Amt            ");
	       	 sql.append("       from  (select * from cd01_99 cd01 where cd01.hsien_id <> 'Y') cd01                        ");
	       	 sql.append("       left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = ?                     ");
	       	 sql.append("       left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type= ? and bn01.m_year = wlx01.m_year and bn01.m_year = ? ");
	       	 sql.append("       left join (select * from WLX09_S_WARNING where m_year  = ?  and m_quarter = ? ) a01                                     ");
	       	 sql.append("                  on  bn01.bank_no = a01.bank_no                                                                               ");
	       	 sql.append("       group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order                                   ");
	       	 sql.append("       ) a01 ");
	       	 sql.append(" order by  FR001W_output_order, hsien_id ");
  		}else {
  			sql.append("SELECT ' '                                                 AS hsien_id, ");
  			sql.append("       '合計'                                              AS hsien_name, ");
  			sql.append("       '000'                                               AS fr001w_output_order, ");
  			sql.append("       SUM(a01.warnaccount_tcnt)                           AS warnaccount_tcnt, ");
  			sql.append("       Round(SUM(a01.warnaccount_tbal) / ?, 0)             AS warnaccount_tbal, ");
  			sql.append("       SUM(a01.warnaccount_remit_tcnt)                     AS warnaccount_remit_tcnt, ");
  			sql.append("       SUM(a01.warnaccount_refund_apply_cnt)               AS warnaccount_refund_apply_cnt, ");
  			sql.append("       Round(SUM(a01.warnaccount_refund_apply_amt) / ?, 0) AS warnaccount_refund_apply_amt, ");
  			sql.append("       SUM(a01.warnaccount_refund_cnt)                     AS warnaccount_refund_cnt, ");
  			sql.append("       Round(SUM(a01.warnaccount_refund_amt) / ?, 0)       AS warnaccount_refund_amt ");
  			sql.append("FROM   (SELECT * ");
  			sql.append("        FROM   cd01 ");
  			sql.append("        WHERE  cd01.hsien_id <> 'Y') cd01 ");
  			sql.append("       left join wlx01 ");
  			sql.append("         ON wlx01.hsien_id = cd01.hsien_id ");
  			sql.append("            AND wlx01.m_year = ? ");
  			sql.append("       left join bn01 ");
  			sql.append("         ON wlx01.bank_no = bn01.bank_no ");
  			sql.append("            AND bn01.bank_type = ? ");
  			sql.append("            AND bn01.m_year = wlx01.m_year ");
  			sql.append("            AND bn01.m_year = ? ");
  			sql.append("       left join (SELECT * ");
  			sql.append("                  FROM   wlx09_s_warning ");
  			sql.append("                  WHERE  m_year = ? ");
  			sql.append("                         AND m_quarter = ?) a01 ");
  			sql.append("         ON bn01.bank_no = a01.bank_no ");
  			sql.append("UNION ");
  			sql.append("SELECT *  ");
  			sql.append("FROM   (SELECT Nvl(cd01.hsien_id, ' ')                             AS hsien_id, ");
  			sql.append("               Nvl(cd01.hsien_name, 'OTHER')                       AS hsien_name, ");
  			sql.append("               cd01.fr001w_output_order                            AS fr001w_output_order, ");
  			sql.append("               SUM(a01.warnaccount_tcnt)                           AS warnaccount_tcnt, ");
  			sql.append("               Round(SUM(a01.warnaccount_tbal) / ?, 0)             AS warnaccount_tbal, ");
  			sql.append("               SUM(a01.warnaccount_remit_tcnt)                     AS warnaccount_remit_tcnt, ");
  			sql.append("               SUM(a01.warnaccount_refund_apply_cnt)               AS warnaccount_refund_apply_cnt, ");
  			sql.append("               Round(SUM(a01.warnaccount_refund_apply_amt) / ?, 0) AS warnaccount_refund_apply_amt, ");
  			sql.append("               SUM(a01.warnaccount_refund_cnt)                     AS warnaccount_refund_cnt, ");
  			sql.append("               Round(SUM(a01.warnaccount_refund_amt) / ?, 0)       AS warnaccount_refund_amt ");
  			sql.append("        FROM   (SELECT * ");
  			sql.append("                FROM   cd01 ");
  			sql.append("                WHERE  cd01.hsien_id <> 'Y') cd01 ");
  			sql.append("               left join wlx01 ");
  			sql.append("                 ON wlx01.hsien_id = cd01.hsien_id ");
  			sql.append("                    AND wlx01.m_year = ? ");
  			sql.append("               left join bn01 ");
  			sql.append("                 ON wlx01.bank_no = bn01.bank_no ");
  			sql.append("                    AND bn01.bank_type = ? ");
  			sql.append("                    AND bn01.m_year = wlx01.m_year ");
  			sql.append("                    AND bn01.m_year = ? ");
  			sql.append("               left join (SELECT * ");
  			sql.append("                          FROM   wlx09_s_warning ");
  			sql.append("                          WHERE  m_year = ? ");
  			sql.append("                                 AND m_quarter = ?) a01 ");
  			sql.append("                 ON bn01.bank_no = a01.bank_no ");
  			sql.append("        GROUP  BY Nvl(cd01.hsien_id, ' '), ");
  			sql.append("                  Nvl(cd01.hsien_name, 'OTHER'), ");
  			sql.append("                  cd01.fr001w_output_order) a01 ");
  			sql.append("ORDER  BY fr001w_output_order, ");
  			sql.append("          hsien_id ");
  			
  		}
  		return sql.toString();
  	}
}



