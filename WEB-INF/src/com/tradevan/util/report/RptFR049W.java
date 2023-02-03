/*
 *  97.09.03 create 農漁會信用部違反農金局及其子法而遭罰鍰明細表 by 2295
 *  99.05.26 fixed 縣市合併 & sql injection by 2808
 * 101.05.09 add 處分方式  by 2295
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

public class RptFR049W {	 
	 public static String createRpt(String cityType,String begDate,String endDate,String begY,String violate_type) {    

	    String errMsg = "";	    
	    StringBuffer sqlCmd = new StringBuffer();
	    List sqlCmdList = new ArrayList () ;
	    StringBuffer condition =  new StringBuffer();
	    List conditionList =new ArrayList() ;
	    List dbData = null;	    
	    int rowNum=0;
	    DataObject bean = null;
	    reportUtil reportUtil = new reportUtil();
		HSSFCellStyle cs_right = null; 
		HSSFCellStyle cs_center = null;
	    HSSFCellStyle cs_left = null;
	    String bank_name="";//受處分機構
	  	String violate_date ="";//受處分日期    
	  	String title="";//主旨
	  	String content="";//說明
	  	String law_content="";//法令依據
	  	
	  	String u_year = "99" ;
		String cd01Table = "cd01_99" ;
		StringBuffer bn01= new StringBuffer() ;
		StringBuffer wlx01 = new StringBuffer() ;
	    try {
	    	if("".equals(begY) || Integer.parseInt(begY)>99) {
	    		u_year = "100" ;
	    		cd01Table = "cd01" ;
	    	}
	    	bn01.append("(select BANK_NO,BANK_NAME,BN_TYPE,BANK_TYPE,ADD_USER,ADD_NAME,ADD_DATE,BANK_B_NAME,KIND_1,KIND_2,BN_TYPE2,EXCHANGE_NO from bn01 where m_year=? )") ;
			wlx01.append("(select BANK_NO ,ENGLISH,SETUP_APPROVAL_UNT,SETUP_DATE,SETUP_NO,CHG_LICENSE_DATE,");
			wlx01.append(" CHG_LICENSE_NO,CHG_LICENSE_REASON,START_DATE,BUSINESS_ID,HSIEN_ID,AREA_ID,ADDR,");
			wlx01.append(" TELNO,FAX,EMAIL,WEB_SITE,CENTER_FLAG,CENTER_NO,STAFF_NUM,IT_HSIEN_ID,IT_AREA_ID,");
			wlx01.append(" IT_ADDR,IT_NAME,IT_TELNO,AUDIT_HSIEN_ID,AUDIT_AREA_ID,AUDIT_ADDR,AUDIT_NAME,AUDIT_TELNO,FLAG,OPEN_DATE,M2_NAME,");
			wlx01.append(" HSIEN_DIV_1,CANCEL_NO,CANCEL_DATE from wlx01 where m_year =? )");
			
	      File xlsDir = new File(Utility.getProperties("xlsDir"));
	      File reportDir = new File(Utility.getProperties("reportDir"));

	      if (!xlsDir.exists()) {
	        if (!Utility.mkdirs(Utility.getProperties("xlsDir"))) {
	          errMsg += Utility.getProperties("xlsDir") + "目錄新增失敗";
	        }
	      }
	      if (!reportDir.exists()) {
	        if (!Utility.mkdirs(Utility.getProperties("reportDir"))) {
	          errMsg += Utility.getProperties("reportDir") + "目錄新增失敗";
	        }
	      }
	      FileInputStream finput = null;

	      //input the standard report form      
	      finput = new FileInputStream(xlsDir +System.getProperty("file.separator") +"農漁會信用部違反農金局及其子法而遭處分明細表.xls");

	      //設定FileINputStream讀取Excel檔
	      POIFSFileSystem fs = new POIFSFileSystem(finput);
	      HSSFWorkbook wb = new HSSFWorkbook(fs);
	      HSSFSheet sheet = wb.getSheetAt(0); //讀取第一個工作表，宣告其為sheet
	      HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
	      //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
	      //sheet.setAutobreaks(true); //自動分頁

	      //設定頁面符合列印大小
	      sheet.setAutobreaks(false);
	      ps.setScale( (short) 100); //列印縮放百分比

	      ps.setPaperSize( (short) 9); //設定紙張大小 A4
	      //wb.setSheetName(0,"test");
	      finput.close();

	      HSSFRow row = null; //宣告一列
	      HSSFCell cell = null; //宣告一個儲存格
	      
	      cs_right = reportUtil.getRightStyle(wb);
	      cs_center = reportUtil.getDefaultStyle(wb);
	      cs_left = reportUtil.getNoBorderLeftStyle(wb);
	   	  
	      sqlCmd.append(" select bn01.bank_no,bn01.bank_name,"					  
			     + " ((TO_CHAR(violate_date,'yyyy')-1911)||'/'|| TO_CHAR(violate_date,'mm/dd')) as violate_date,"
			     + " TO_CHAR(violate_date,'yyyy/mm/dd') as violate_date_1,"
			     + " violate_type,title,content,law_content "
			     + " from mis_violatelaw "
			     + " left join ").append(bn01.toString()).append(" bn01 on mis_violatelaw.bank_no = bn01.bank_no ");
	      sqlCmdList.add(u_year) ;
 	      if(!cityType.equals("") && !cityType.equals("ALL")) {
	      	  condition.append( (condition.length() > 0 ? " and":"")+" mis_violatelaw.BANK_NO in (select BANK_NO  from ").append(wlx01.toString()).append(" WLX01  where HSIEN_ID =  ? ) ");
	      	  conditionList.add(u_year) ;
	      	  conditionList.add(cityType) ;
	      }

	      if(!begDate.equals("") && !endDate.equals("")){
	      	  condition.append((condition.length() > 0 ? " and":"")+" TO_CHAR(violate_date, 'yyyy/mm/dd') BETWEEN ? AND ? " );
	      	  conditionList.add(begDate) ;
	      	  conditionList.add(endDate) ;
	      }
	      
	      if(!violate_type.equals("")) {
              condition.append( (condition.length() > 0 ? " and":"")+" (").append(violate_type.toString()).append(" ) ");             
          }
	      if(condition.length() > 0) sqlCmd.append("where ").append(condition.toString());
	      sqlCmd.append(" ORDER BY violate_date" );
	      for(int i=0 ;i<conditionList.size();i++) {
	    	  sqlCmdList.add(conditionList.get(i)) ;
	      }
          dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),sqlCmdList,"violate_date,violate_date_1");	      
	      System.out.println("dbData.size=" + dbData.size());
	      
	      //設定報表表頭資料============================================
	      row=sheet.getRow(2);
	      cell=row.getCell((short)0);	       	
	      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	      cell.setCellValue("受處分區間:"+Utility.getCHTdate(begDate,1)+" ~ "+Utility.getCHTdate(endDate,1));  	
	       
	      if (dbData != null && dbData.size() != 0) {
	          
	          List paramList = new ArrayList();
	          StringBuffer sql = new StringBuffer();
	          List violate_type_list = null;
	          DataObject violate_type_bean = null;
	          sql.append(" select cmuse_id,cmuse_name from cdshareno where cmuse_div=? order by input_order"); 
	          paramList.add("038");       
	          List violate_type_dbdate = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"");//處分方式代碼檔                    
	          
	      	  rowNum = 4;
	      	  for(int j=0;j<dbData.size();j++){
	      	      bank_name="";violate_date ="";title="";content="";law_content="";
	      	      bean = (DataObject)dbData.get(j);
	      	      bank_name=(bean.getValue("bank_name") == null)?"":(String)bean.getValue("bank_name");//受處分機構
	      	      violate_date =(bean.getValue("violate_date") == null)?"":(bean.getValue("violate_date")).toString();//受處分日期  	  	 
	      	      title=(bean.getValue("title") == null)?"":(String)bean.getValue("title");//主旨   
	      	      content=(bean.getValue("content") == null)?"":(String)bean.getValue("content");//說明
	      	      law_content=(bean.getValue("law_content") == null)?"":(String)bean.getValue("law_content");//法令依據 
	    	  	  violate_type = (bean.getValue("violate_type") == null)?"":(String)bean.getValue("violate_type");//處分方式.101.05.09 add
	    	  	  row = sheet.createRow(rowNum);
	    	  	  
	    	  	  //System.out.println("violate_type="+violate_type);
	    	  	  violate_type_list= new LinkedList();	            
	              if(!"".equals(violate_type)){
	                  violate_type_list = Utility.getStringTokenizerData(violate_type,":");
	                  violate_type = "";
	                  for(int i=0;i<violate_type_list.size();i++){//處分方式
	                      //System.out.println("i="+(String)violate_type_list.get(i));
	                      violate_type_loop:
	                      for(int k=0;k<violate_type_dbdate.size();k++){//代碼檔
	                          violate_type_bean = (DataObject)violate_type_dbdate.get(k);   
	                          if(((String)violate_type_list.get(i)).equals(violate_type_bean.getValue("cmuse_id"))){
	                             violate_type += (violate_type.length() > 0 ?"\n":"")+violate_type_bean.getValue("cmuse_name");
	                             //System.out.println(violate_type);
	                             break violate_type_loop;
	                          }
	                      }//end of violate_type_dbdate
	                  }//end of violate_type_list
	              }//end of violate_type != ''	                
	    	  	  
	    	  	  //列印各機構明細資料
				  for(int cellcount=0;cellcount<=5;cellcount++){			 	      
			 	      cell=row.createCell((short)cellcount);			 		
			    	  cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			    	  cell.setCellStyle(cs_center);		   
			    	  if(cellcount == 0) cell.setCellValue(bank_name);	
			    	  if(cellcount == 1) cell.setCellValue(violate_date);
			    	  if(cellcount == 2) cell.setCellValue(violate_type);   
			    	  if(cellcount == 3) cell.setCellValue(title);	 
			    	  if(cellcount == 4) cell.setCellValue(content);
			    	  if(cellcount == 5) cell.setCellValue(law_content);			    	  
				  }//end of cellcount	
				  rowNum++;
	      	  }
	      } //end of else dbData.size() != 0
	      
	      
	      FileOutputStream fout = null;     
	      fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + "農漁會信用部違反農金局及其子法而遭處分明細表.xls");
	     
	      HSSFFooter footer = sheet.getFooter();
	      footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
	      footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
	      wb.write(fout);
	      //儲存
	      fout.close();
	    }
	    catch (Exception e) {
	      System.out.println("RptFR049W.createRpt Error:" + e + e.getMessage());
	    }
	    
	    return errMsg;
	  }
}
