/*
 *  97.09.02-03 create 辦理基層金融機構變現性資產查核表 by 2295
 *  98.01.09 查核人員直接顯示 by 2295
 * 100.01.31 fix 根據查詢年度.100年以後取得新縣市別.100年以前取得舊縣市別 
 * 				 使用PreparedStatement;並列印轉換後的SQL;套用QueryDB_SQLParam by 2295
 * 109.02.26 fix 調整查核項目.查核結果欄位顯示長度 by 2295
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

public class RptFR048W {
	 /* M_YEAR :查核年核
	  * M2_NAME :地方主管機關代碼
	  * */
	 public static String createRpt(String M_YEAR, String M2_NAME ) {    

	    String errMsg = "";	    
	    String sqlCmd = "";  
	    List dbData = null;
	    String bank_name="";
	    int rowNum=0;
	    DataObject bean = null;
	    reportUtil reportUtil = new reportUtil();
		HSSFCellStyle cs_right = null; 
		HSSFCellStyle cs_center = null;
	    HSSFCellStyle cs_left = null;
	    String examine="";//受檢單位
	  	String check_date ="";//查核日期     
	  	String name="";//查核人員
	  	String chinese_name="";//查核人員.中文名稱
	  	String item="";//查核項目
	  	String result="";//查核結果   
	  	String content="";//處理情形   
	  	String remark="";//備註
    	List paramList = new ArrayList();
    	String cd01_table = "";
        String wlx01_m_year = "";
	    try {
	     //100.01.31 add 查詢年度100年以前.縣市別不同===============================
  	     cd01_table = (Integer.parseInt(M_YEAR) < 100)?"cd01_99":""; 
  	     wlx01_m_year = (Integer.parseInt(M_YEAR) < 100)?"99":"100"; 
  	     //=====================================================================    
      		
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
	      finput = new FileInputStream(xlsDir +System.getProperty("file.separator") +"縣市政府辦理基層金融機構變現性資產查核表.xls");

	      //設定FileINputStream讀取Excel檔
	      POIFSFileSystem fs = new POIFSFileSystem(finput);
	      HSSFWorkbook wb = new HSSFWorkbook(fs);
	      HSSFSheet sheet = wb.getSheetAt(0); //讀取第一個工作表，宣告其為sheet
	      HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
	      //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
	      //sheet.setAutobreaks(true); //自動分頁

	      //設定頁面符合列印大小
	      sheet.setAutobreaks(false);
	      ps.setScale( (short) 96); //列印縮放百分比 //109.02.26 fix
	   
	      ps.setPaperSize( (short) 9); //設定紙張大小 A4
	      
	      ps.setLandscape( true ); // 設定橫印 //109.02.26 fix
	      //wb.setSheetName(0,"test");
	      finput.close();

	      HSSFRow row = null; //宣告一列
	      HSSFCell cell = null; //宣告一個儲存格
	      
	      cs_right = reportUtil.getRightStyle(wb);
	      cs_center = reportUtil.getDefaultStyle(wb);
	      cs_left = reportUtil.getNoBorderLeftStyle(wb);
	   	  /*
	   	  --各明細資料
	   	  select m_year,m2_name,examine,to_char(check_date,'YYYY/MM/DD') as check_date,
	   	  		 name,item,result,content,remark,ba01.bank_name
	   	  from BOAF_ASSETCHECK 
	   	  left join ba01 on BOAF_ASSETCHECK.examine = ba01.bank_no 
	   	  where m_year=97				  
	   	  and m2_name='AAAA000'					  
	   	  order by m_year,check_date
	      --總機構/分支機構.查核比率(小數第2位.四捨五入)
          select a.tbank_sum,c.tbank_check,round(c.tbank_check/a.tbank_sum*100,2) as tbank_check_rate,
          	     b.bank_sum,d.bank_check,round(d.bank_check/b.bank_sum*100,2) as bank_check_rate
          from (select count(*) as tbank_sum
          		from bn01 left join wlx01 on bn01.bank_no = wlx01.bank_no
          		where m2_name='AAAA000'--該縣市政府轄區下的總機構
          		and bn01.bn_type <> '2')a,
          	   (select count(*) as bank_sum
          		from bn02 left join wlx01 on bn02.TBANK_NO = wlx01.BANK_NO
          		where m2_name='AAAA000'
          		and bn_type <> '2')b, --該縣市政府轄區下的分支機構
               (select count(*) as tbank_check
                from BOAF_ASSETCHECK
                left join ba01 on BOAF_ASSETCHECK.examine = ba01.bank_no
                where BOAF_ASSETCHECK.m2_name = 'AAAA000' and m_year=97
                and ba01.BANK_KIND='0')c, --實際查核總機構
               (select count(*) as bank_check
                from BOAF_ASSETCHECK
                left join ba01 on BOAF_ASSETCHECK.examine = ba01.bank_no
                where BOAF_ASSETCHECK.m2_name = 'AAAA000' and m_year=97
                and ba01.BANK_KIND='1')d --實際查核分支機構
          */
	      //各農漁會明細資料
	      sqlCmd = " select bank_name from ba01 where bank_no=? and m_year=?";
	      paramList.add(M2_NAME);
	      paramList.add(wlx01_m_year);
	      dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");
	      
	      //該機關所屬使用者
	      List name_List = DBManager.QueryDB_SQLParam("select muser_id,muser_name from wtt01 where tbank_no=?",paramList,"");
	      paramList.clear();
	      if(dbData != null && dbData.size() != 0){
	      	 bean = (DataObject)dbData.get(0);
	      	 bank_name = (String)bean.getValue("bank_name");
	      	 bank_name = Utility.ISOtoBig5(Utility.Big5toISO(bank_name).substring(0,6));
	      }
	      sqlCmd = " select BOAF_ASSETCHECK.m_year,m2_name,examine,"
	      		 + " to_char(check_date,'YYYY/MM/DD') as check_date,"
				 + " name,item,result,content,remark,ba01.bank_name"
	      		 + " from BOAF_ASSETCHECK " 
				 + " left join (select * from ba01 where m_year=?)ba01 on BOAF_ASSETCHECK.examine = ba01.bank_no " 
				 + " where BOAF_ASSETCHECK.m_year=?"				  
				 + " and m2_name=?"					  
				 + " order by BOAF_ASSETCHECK.m_year,check_date ";
	      paramList.add(wlx01_m_year);
	      paramList.add(M_YEAR);
	      paramList.add(M2_NAME);
	      dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"m_year,check_date");
	      //媲機構/分支機構.查核比率(小數第2位.四捨五入)
	      paramList.clear();
	      sqlCmd = " select a.tbank_sum,c.tbank_check,round(c.tbank_check/a.tbank_sum*100,2) as tbank_check_rate, "
          	     + " b.bank_sum,d.bank_check,round(d.bank_check/b.bank_sum*100,2) as bank_check_rate"
				 + " from (select count(*) as tbank_sum"
          		 + " from (select * from bn01 where m_year=?)bn01 left join (select * from wlx01 where m_year=?)wlx01 on bn01.bank_no = wlx01.bank_no "
          		 + " where m2_name=?"//--該縣市政府轄區下的總機構
          		 + " and bn01.bn_type <> '2')a,"
          	     + " (select count(*) as bank_sum"
          		 + " from (select * from bn02 where m_year=?)bn02 left join (select * from wlx01 where m_year=?)wlx01 on bn02.TBANK_NO = wlx01.BANK_NO "
          		 + " where m2_name=?"
          		 + " and bn_type <> '2')b, "//--該縣市政府轄區下的分支機構
                 + " (select count(*) as tbank_check "
                 + " from BOAF_ASSETCHECK "
                 + " left join (select * from ba01 where m_year=?)ba01 on BOAF_ASSETCHECK.examine = ba01.bank_no "
                 + " where BOAF_ASSETCHECK.m2_name = ?"
                 + " and BOAF_ASSETCHECK.m_year=?"
                 + " and ba01.BANK_KIND='0')c, "//--實際查核總機構
                 + " (select count(*) as bank_check "
                 + " from BOAF_ASSETCHECK"
                 + " left join (select * from ba01 where m_year=?)ba01 on BOAF_ASSETCHECK.examine = ba01.bank_no"
                 + " where BOAF_ASSETCHECK.m2_name = ?"
                 + " and BOAF_ASSETCHECK.m_year=?"
                 + " and ba01.BANK_KIND='1')d ";//--實際查核分支機構
	      paramList.add(wlx01_m_year);
	      paramList.add(wlx01_m_year);
	      paramList.add(M2_NAME);
	      paramList.add(wlx01_m_year);
	      paramList.add(wlx01_m_year);
	      paramList.add(M2_NAME);
	      paramList.add(wlx01_m_year);
	      paramList.add(M2_NAME);
	      paramList.add(M_YEAR);
	      paramList.add(wlx01_m_year);
	      paramList.add(M2_NAME);
	      paramList.add(M_YEAR);
	      
	      List dbData_total = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"tbank_sum,tbank_check,tbank_check_rate,bank_sum,bank_check,bank_check_rate");
	        
	      System.out.println("dbData.size=" + dbData.size());
	      System.out.println("dbData_total.size=" + dbData_total.size());
	      //設定報表表頭資料============================================
	      
	      row=sheet.getRow(1);
	      cell=row.getCell((short)0);	       	
	      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	      cell.setCellValue(bank_name+"政府"+M_YEAR+"年「辦理基層金融機構變現性資產查核」"+((dbData == null || dbData.size() ==0)?"無資料存在":"執行情形報告表"));  	
	       
	      if (dbData != null && dbData.size() != 0) {
	      	  rowNum = 4;
	      	  for(int j=0;j<dbData.size();j++){
	      	      chinese_name="";examine="";bank_name="";item="";name="";result="";content="";remark="";check_date="";
	      	      bean = (DataObject)dbData.get(j);
	      	      examine=(bean.getValue("bank_name") == null)?"":(String)bean.getValue("bank_name");//受檢單位
	    	  	  check_date =(bean.getValue("check_date") == null)?"":(bean.getValue("check_date")).toString();//查核日期
	    	  	  if(check_date.length() > 0){
	    	  	  	check_date = Utility.getCHTdate(check_date,0);
	    	  	  }
	    	  	  name=(bean.getValue("name") == null)?"":(String)bean.getValue("name");//查核人員   
	    	  	  item=(bean.getValue("item") == null)?"":(String)bean.getValue("item");//查核項目
	    	  	  result=(bean.getValue("result") == null)?"":(String)bean.getValue("result");//查核結果   
	    	  	  content=(bean.getValue("content") == null)?"":(String)bean.getValue("content");//處理情形   
	    	  	  remark=(bean.getValue("remark") == null)?"":(String)bean.getValue("remark");//備註
	    	  	  /*
	    	  	  for(int i=0;i<name_List.size();i++){                            
                    if(name.indexOf((String)((DataObject)name_List.get(i)).getValue("muser_id")) != -1){
                       chinese_name += (String)((DataObject)name_List.get(i)).getValue("muser_name");
                       if(i < name_List.size() -1) chinese_name +="\n";
                    }     
                  }
                  */
	    	  	  row = sheet.createRow(rowNum);
	    	  	  //列印各機構明細資料
				  for(int cellcount=0;cellcount<=6;cellcount++){			 	      
			 	      cell=row.createCell((short)cellcount);			 		
			    	  cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			    	  cell.setCellStyle(cs_center);		   
			    	  if(cellcount == 0) cell.setCellValue(examine);	
			    	  if(cellcount == 1) cell.setCellValue(check_date);	
			    	  if(cellcount == 2) cell.setCellValue(name);//98.01.09 查核人員直接顯示   	 
			    	  if(cellcount == 3) cell.setCellValue(item);
			    	  if(cellcount == 4) cell.setCellValue(result);
			    	  if(cellcount == 5) cell.setCellValue(content);	    
			    	  if(cellcount == 6) cell.setCellValue(remark);
				  }//end of cellcount	
				  rowNum++;
	      	  }
	      	  rowNum++;
	      	  //列印結尾機構數.查核機構.查核比率 
	      	  if (dbData_total != null && dbData_total.size() != 0) {	
	      	      bean = (DataObject)dbData_total.get(0);
	      	      //總機構
	      	      row = sheet.createRow(rowNum);
			      for(int cellcount=0;cellcount<=6;cellcount++){			 	      
		 	          cell=row.createCell((short)cellcount);			 		
		    	      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		    	      cell.setCellStyle(cs_left);		   
		    	      if(cellcount == 0) cell.setCellValue("總機構數:"+(bean.getValue("tbank_sum")).toString());
		    	      if(cellcount == 2) cell.setCellValue("實際查核總機構數:"+(bean.getValue("tbank_check")).toString());
		    	      if(cellcount == 5) cell.setCellValue("查核比率:"+(bean.getValue("tbank_check_rate")).toString()+"%");		    	  
			      }//end of cellcount	
			      sheet.addMergedRegion( new Region( ( short )rowNum, ( short )2,
							 						 ( short )rowNum,
													 ( short )3) );
			      rowNum++;
	      	      row = sheet.createRow(rowNum);
	      	      //分支機構
			      for(int cellcount=0;cellcount<=6;cellcount++){			 	      
		 	          cell=row.createCell((short)cellcount);			 		
		    	      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		    	      cell.setCellStyle(cs_left);		   
		    	      if(cellcount == 0) cell.setCellValue("分支機構數:"+(bean.getValue("bank_sum")).toString());
		    	      if(cellcount == 2) cell.setCellValue("實際查核分支機構數:"+(bean.getValue("bank_check")).toString());
		    	      if(cellcount == 5) cell.setCellValue("查核比率:"+(bean.getValue("bank_check_rate")).toString()+"%");		    	  
			      }//end of cellcount	
			      sheet.addMergedRegion( new Region( ( short )rowNum, ( short )2,
			      									 ( short )rowNum,
													 ( short )3) );
			      rowNum++;
			      row = sheet.createRow(rowNum);
			      //簽名區
			      for(int cellcount=0;cellcount<=6;cellcount++){			 	      
		 	          cell=row.createCell((short)cellcount);			 		
		    	      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		    	      cell.setCellStyle(cs_left);		   
		    	      if(cellcount == 0) cell.setCellValue("填表:");
		    	      if(cellcount == 2) cell.setCellValue("課長:");
		    	      if(cellcount == 5) cell.setCellValue("局長:");		    	  
			      }//end of cellcount
	      	  }    
	      } //end of else dbData.size() != 0
	      
	      
	      FileOutputStream fout = null;     
	      fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + "縣市政府辦理基層金融機構變現性資產查核表.xls");
	     
	      HSSFFooter footer = sheet.getFooter();
	      footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
	      footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
	      wb.write(fout);
	      //儲存
	      fout.close();
	    }
	    catch (Exception e) {
	      System.out.println("RptFR048W.createRpt Error:" + e + e.getMessage());
	    }
	    
	    return errMsg;
	  }
}
