/*
 * Created on 2005/12/04 by 4180
 * fixed on 2006/01/20 by 4180
 * SQL中的除以金額單位(unit)移除，改於資料送出前處理
 * fixed on 2006/01/24 by 4180
 * 存放比率公式修正
 * 2006.06.14 add 農漁會存放款彙明細表 by 2295
 * 2006.06.21 fix 農漁會總表存放比率sql by 2295
 * 2010.03.25 fix 根據查詢年度.100年以後取得新縣市別.100年以前取得舊縣市別 
 * 				  使用PreparedStatement;並列印轉換後的SQL;套用QueryDB_SQLParam by 2295
 * 2010.09.23 fix 100年(含)以後.明細表加印新北市/台中市/台南市 by 2295
 * 2010.12.01 fix wlx01.bn01 sql by 2295
 * 2012.02.17 fix 部份加總欄位區分農漁會科目 by 2295
 *            add 漁會增加110217存放行庫-合作金庫-公庫存款 by 2295
 * 			  add 區分農/漁會.99年以前/100年以後,總計位置 by 2295
 * 2012.06.14 add 103年(含)以後SQL by2968
 * 2013.09.03 fix 110257科目代號改為110256
 * 2014.01.16 add 臺灣省農會更名為中華民國農會增加說明 by 2295
 * 2014.12.24 fix 桃園縣升格調整 by 2968
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

public class RptFR040WA{

    public static String createRpt(String s_year,String s_month,String unit,String bank_type,String rptStyle){

    	System.out.println("inpute rptStyle = "+rptStyle);
    	System.out.println("明細表開始");
		String errMsg = "";
		String unit_name="";		
		int i=0;
		int j=0;
		String s_year_last="";
		String s_month_last="";
		String hsien_id_sum[]={"a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p"};
		String field_name[]={"bank_no","bank_name","field_220100","field_220200","field_220300","field_220400","field_220500","field_220600",
				 			 "field_220700","field_220800","field_220900","field_221000","field_220000","field_110321","field_110322",
							 "field_110323","field_110324","field_110325","field_110326","field_110327","field_110320","field_110301",
							 "field_110302","field_110303","field_110304","field_110305","field_110306","field_110307","field_110310",
							 "field_110311","field_110312","field_110313","field_110300"};

		String bank_type_name=(bank_type.equals("6"))?"農會":"漁會";
		String cd01_table = "";
        String wlx01_m_year = "";
		List paramList = new ArrayList();
		int flag=0;
		int rowcount_Taipei=0,rowcount_Kaohsiung=0;
		
		String filename="";
		String openfile="";
		String insertValue="";
		filename="農漁會"+((rptStyle.equals("0"))?"存放款彙總表.xls":"存放款彙明細表.xls");
		openfile = bank_type_name+((rptStyle.equals("0"))?"存放款彙總表.xls":"存放款彙明細表.xls");
		String hsien_name = ""; 
		reportUtil reportUtil = new reportUtil();
		try{
			//99.03.24 add 查詢年度100年以前.縣市別不同===============================
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
	        ps.setScale( ( short )70 ); //列印縮放百分比

	        ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
	  		//wb.setSheetName(0,"test");
	  		finput.close();
	  		
	  		HSSFRow row=null;//宣告一列 
	  		HSSFCell cell=null;//宣告一個儲存格
	  		String div=(Integer.parseInt(s_year)==94 && Integer.parseInt(s_month)==6)?"1":"2";
	  		if(Integer.parseInt(s_month) == 1) { 
			    s_year_last   =  String.valueOf(Integer.parseInt(s_year) - 1); 
			    s_month_last  =  "12";
			}else{ 
			    s_year_last   =  s_year; 
			    s_month_last = String.valueOf(Integer.parseInt(s_month) - 1);			    
			}

			String report_head="																"+s_year+"年" + s_month + "月"+bank_type_name+"信用部存款及轉存行庫明細表";			
			//列印年度
			row=(sheet.getRow(0)==null)? sheet.createRow(0) : sheet.getRow(0);													
			cell=row.getCell((short)1);	
			System.out.println(report_head);	
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);						
			if(s_month.equals("0")){
				cell.setCellValue("  					中華民國　"+s_year+"　年度");
			}else {																								
				cell.setCellValue(report_head);							
			}
			
			//列印單位		
			row=(sheet.getRow(1)==null)? sheet.createRow(1) : sheet.getRow(1);						
			cell=row.getCell((short)24);			
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);												
			unit_name = Utility.getUnitName(unit);
			//cell.setCellValue(" 單位：新台幣"+unit_name+"、％");
			insertValue=" 單位：新台幣"+unit_name+"、％";						
			insertCell(insertValue,1,24,wb,row,sheet,cell);							
			StringBuffer sqlCmd = new StringBuffer();
			sqlCmd.append("select hsien_id, hsien_name, FR001W_output_order, bank_no , BANK_NAME, COUNT_SEQ, field_SEQ, "); 
			sqlCmd.append("       round(field_220100 /?,0) as  field_220100, ");  // 信用部-支票存款
			sqlCmd.append("       round(field_220200 /?,0) as  field_220200, ");  // 信用部-保付支票
			sqlCmd.append("       round(field_220300 /?,0) as  field_220300, ");  // 信用部-活期支票
			sqlCmd.append("       round(field_220400 /?,0) as  field_220400, ");  // 信用部-活期儲蓄存款
			sqlCmd.append("       round(field_220500 /?,0) as  field_220500, ");  // 信用部-員工活期儲蓄存款
			sqlCmd.append("       round(field_220600 /?,0) as  field_220600, ");  // 信用部-定期存款
			sqlCmd.append("       round(field_220700 /?,0) as  field_220700, ");  // 信用部-定期儲蓄存款
			sqlCmd.append("       round(field_220800 /?,0) as  field_220800, ");  // 信用部-員工定期儲蓄存款
			sqlCmd.append("       round(field_220900 /?,0) as  field_220900, ");  // 信用部-公庫存款
			sqlCmd.append("       round(field_221000 /?,0) as  field_221000, ");  // 信用部-本會支票
			sqlCmd.append("       round(field_220000 /?,0) as  field_220000, ");  // 信用部-小計
			sqlCmd.append("       round(field_110321 /?,0) as  field_110321, ");  // 農業金庫-支票存款
			sqlCmd.append("       round(field_110322 /?,0) as  field_110322, ");  // 農業金庫-活期存款
			sqlCmd.append("       round(field_110323 /?,0) as  field_110323, ");  // 農業金庫-活期儲蓄存款
			sqlCmd.append("       round(field_110324 /?,0) as  field_110324, ");  // 農業金庫-定期存款
			sqlCmd.append("       round(field_110325 /?,0) as  field_110325, ");  // 農業金庫-定期儲蓄存款
			sqlCmd.append("       round(field_110326 /?,0) as  field_110326, ");  // 農業金庫-公庫存款
			sqlCmd.append("       round(field_110327 /?,0) as  field_110327, ");  // 農業金庫-金資專戶
			sqlCmd.append("       round(field_110320 /?,0) as  field_110320, ");  // 農業金庫-小計
			sqlCmd.append("       round(field_110301 /?,0) as  field_110301, ");  // 合作金庫-支票存款
			sqlCmd.append("       round(field_110302 /?,0) as  field_110302, ");  // 合作金庫-活期存款
			sqlCmd.append("       round(field_110303 /?,0) as  field_110303, ");  // 合作金庫-活期儲蓄存款
			sqlCmd.append("       round(field_110304 /?,0) as  field_110304, ");  // 合作金庫-定期存款
			sqlCmd.append("       round(field_110305 /?,0) as  field_110305, ");  // 合作金庫-定期儲蓄存款
			sqlCmd.append("       round(field_110306 /?,0) as  field_110306, ");  // 合作金庫-公庫存款
			sqlCmd.append("       round(field_110307 /?,0) as  field_110307, ");  // 合作金庫-金資專戶
			sqlCmd.append("       round(field_110310 /?,0) as  field_110310, ");  // 合作金庫-小計
			sqlCmd.append("       round(field_110311 /?,0) as  field_110311, ");  // 其他行庫-農民銀行
			sqlCmd.append("       round(field_110312 /?,0) as  field_110312, ");  // 其他行庫-土地銀行
			sqlCmd.append("       round(field_110313 /?,0) as  field_110313, ");  // 其他行庫-其他銀行
			sqlCmd.append("       round(field_110300 /?,0) as  field_110300  ");  // 其他行庫-合計
			for(int k=1;k<=31; k++){
                paramList.add(unit);
            }
			sqlCmd.append(" from (  ");
			//sql_總計
			sqlCmd.append("      select  ' ' AS hsien_id, '總 計' AS hsien_name, '001' AS FR001W_output_order, ");
			sqlCmd.append("                ' ' AS bank_no, ' ' AS BANK_NAME,  COUNT(*) AS COUNT_SEQ, 'A99' as field_SEQ, ");
			sqlCmd.append("                sum(a01.field_220100) field_220100, ");
			sqlCmd.append("                sum(a01.field_220200) field_220200, ");
			sqlCmd.append("                sum(a01.field_220300) field_220300, ");
			sqlCmd.append("                sum(a01.field_220400) field_220400, ");
			sqlCmd.append("                sum(a01.field_220500) field_220500, ");
			sqlCmd.append("                sum(a01.field_220600) field_220600, ");
			sqlCmd.append("                sum(a01.field_220700) field_220700, ");
			sqlCmd.append("                sum(a01.field_220800) field_220800, ");
			sqlCmd.append("                sum(a01.field_220900) field_220900, ");
			sqlCmd.append("                sum(a01.field_221000) field_221000, ");
			sqlCmd.append("                sum(a01.field_220000) field_220000, ");
			sqlCmd.append("                sum(a01.field_110321) field_110321, ");
			sqlCmd.append("                sum(a01.field_110322) field_110322, ");
			sqlCmd.append("                sum(a01.field_110323) field_110323, ");
			sqlCmd.append("                sum(a01.field_110324) field_110324, ");
			sqlCmd.append("                sum(a01.field_110325) field_110325, ");
			sqlCmd.append("                sum(a01.field_110326) field_110326, ");
			sqlCmd.append("                sum(a01.field_110327) field_110327, ");
			sqlCmd.append("                sum(a01.field_110320) field_110320, ");
			sqlCmd.append("                sum(a01.field_110301) field_110301, ");
			sqlCmd.append("                sum(a01.field_110302) field_110302, ");
			sqlCmd.append("                sum(a01.field_110303) field_110303, ");
			sqlCmd.append("                sum(a01.field_110304) field_110304, ");
			sqlCmd.append("                sum(a01.field_110305) field_110305, ");
			sqlCmd.append("                sum(a01.field_110306) field_110306, ");
			sqlCmd.append("                sum(a01.field_110307) field_110307, ");
			sqlCmd.append("                sum(a01.field_110310) field_110310, ");
			sqlCmd.append("                sum(a01.field_110311) field_110311, ");
			sqlCmd.append("                sum(a01.field_110312) field_110312, ");
			sqlCmd.append("                sum(a01.field_110313) field_110313, ");
			sqlCmd.append("                sum(a01.field_110300) field_110300  ");
			sqlCmd.append("        from( ");                             
			sqlCmd.append("                  select  nvl(cd01.hsien_id,' ') as hsien_id , ");                                                                    
			sqlCmd.append("                        nvl(cd01.hsien_name,'OTHER') as hsien_name, ");                                                                  
			sqlCmd.append("                        cd01.FR001W_output_order as FR001W_output_order, ");                                                          
			sqlCmd.append("                        bn01.bank_no, bn01.BANK_NAME, ");  
			sqlCmd.append("                        sum(decode(a01.acc_code,'220100',amt,0))  as field_220100   , ");          
			sqlCmd.append("                        sum(decode(a01.acc_code,'220200',amt,0))  as field_220200    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'220300',amt,0))  as field_220300    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'220400',amt,0))  as field_220400    , ");       
			sqlCmd.append("                        sum(decode(a01.acc_code,'220500',amt,0))  as field_220500    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'220600',amt,0))  as field_220600    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'220700',amt,0))  as field_220700    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'220800',amt,0))  as field_220800    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'220900',amt,0))  as field_220900    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'221000',amt,0))  as field_221000    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'220000',amt,0))  as field_220000    , ");
			sqlCmd.append("                      decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110321',amt,0),'7',decode(a01.acc_code,'110251',amt,0))),'103',sum(decode(a01.acc_code,'110321',amt,0)),0)  as field_110321    , ");
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110322',amt,0),'7',decode(a01.acc_code,'110252',amt,0))),'103',sum(decode(a01.acc_code,'110322',amt,0)),0)  as field_110322    , ");        
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110323',amt,0),'7',decode(a01.acc_code,'110253',amt,0))),'103',sum(decode(a01.acc_code,'110323',amt,0)),0)  as field_110323    , ");        
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110324',amt,0),'7',decode(a01.acc_code,'110254',amt,0))),'103',sum(decode(a01.acc_code,'110324',amt,0)),0)  as field_110324    , ");        
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110325',amt,0),'7',decode(a01.acc_code,'110255',amt,0))),'103',sum(decode(a01.acc_code,'110325',amt,0)),0)  as field_110325    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'110326',amt,0))     as field_110326    , ");        
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110327',amt,0),'7',decode(a01.acc_code,'110256',amt,0))),'103',sum(decode(a01.acc_code,'110327',amt,0)),0)  as field_110327    , ");        
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110320',amt,0),'7',decode(a01.acc_code,'110250',amt,0))),'103',sum(decode(a01.acc_code,'110320',amt,0)),0)  as field_110320    , ");        
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110301',amt,0),'7',decode(a01.acc_code,'110211',amt,0))),'103',sum(decode(a01.acc_code,'110301',amt,0)),0)  as field_110301    , ");        
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110302',amt,0),'7',decode(a01.acc_code,'110212',amt,0))),'103',sum(decode(a01.acc_code,'110302',amt,0)),0)  as field_110302    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110303',amt,0),'7',decode(a01.acc_code,'110213',amt,0))),'103',sum(decode(a01.acc_code,'110303',amt,0)),0)  as field_110303    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110304',amt,0),'7',decode(a01.acc_code,'110214',amt,0))),'103',sum(decode(a01.acc_code,'110304',amt,0)),0)  as field_110304    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110305',amt,0),'7',decode(a01.acc_code,'110215',amt,0))),'103',sum(decode(a01.acc_code,'110305',amt,0)),0)  as field_110305    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110306',amt,0),'7',decode(a01.acc_code,'110217',amt,0))),'103',sum(decode(a01.acc_code,'110306',amt,0)),0)  as field_110306    , ");
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110307',amt,0),'7',decode(a01.acc_code,'110216',amt,0))),'103',sum(decode(a01.acc_code,'110307',amt,0)),0)  as field_110307    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110310',amt,0),'7',decode(a01.acc_code,'110210',amt,0))),'103',sum(decode(a01.acc_code,'110310',amt,0)),0)  as field_110310    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110311',amt,0),'7',decode(a01.acc_code,'110220',amt,0))),'103',sum(decode(a01.acc_code,'110311',amt,0)),0)  as field_110311    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110312',amt,0),'7',decode(a01.acc_code,'110230',amt,0))),'103',sum(decode(a01.acc_code,'110312',amt,0)),0)  as field_110312    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110313',amt,0),'7',decode(a01.acc_code,'110240',amt,0))),'103',sum(decode(a01.acc_code,'110313',amt,0)),0)  as field_110313    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110300',amt,0),'7',decode(a01.acc_code,'110200',amt,0))),'103',sum(decode(a01.acc_code,'110300',amt,0)),0)  as field_110300      "); 
			sqlCmd.append("          from( select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y' ) cd01 ");
			sqlCmd.append("                left join(  select bn01.bank_type,bn01.bank_no,bn01.bank_name,wlx01.hsien_id ");                     
			sqlCmd.append("                            from (select * from bn01 where m_year=? and bank_type in (?))bn01, (select * from wlx01 where m_year=?)wlx01 "); 
			sqlCmd.append("                            where bn01.bank_no = wlx01.bank_no)bn01 on bn01.hsien_id=cd01.hsien_id ");     
			sqlCmd.append("                left join (select (CASE WHEN (a01.m_year <= 102) THEN '102' ");
			sqlCmd.append("                                                   WHEN (a01.m_year > 102) THEN '103' ");
			sqlCmd.append("                                                   ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01 "); 
			sqlCmd.append("                               where a01.m_year = ? and a01.m_month = ? ) a01 on  bn01.bank_no = a01.bank_code ");
			sqlCmd.append("         group by a01.YEAR_TYPE,nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no ,  bn01.BANK_NAME ");
			sqlCmd.append("         ) a01 "); 
			sqlCmd.append("        where a01.bank_no <> ' ' ");            
			paramList.add(wlx01_m_year);           
            paramList.add(bank_type);
            paramList.add(wlx01_m_year);
            paramList.add(s_year);
            paramList.add(s_month);
			sqlCmd.append("        UNION ALL ");
			//sql_明細
			sqlCmd.append("        select  a01.hsien_id, a01.hsien_name, a01.FR001W_output_order, ");
			sqlCmd.append("                 a01.bank_no, a01.BANK_NAME, "); 
			sqlCmd.append("                 1 AS COUNT_SEQ, ");
			sqlCmd.append("                 'A01' as field_SEQ, ");
			sqlCmd.append("                sum(a01.field_220100) field_220100, ");
			sqlCmd.append("                sum(a01.field_220200) field_220200, ");
			sqlCmd.append("                sum(a01.field_220300) field_220300, ");
			sqlCmd.append("                sum(a01.field_220400) field_220400, ");
			sqlCmd.append("                sum(a01.field_220500) field_220500, ");
			sqlCmd.append("                sum(a01.field_220600) field_220600, ");
			sqlCmd.append("                sum(a01.field_220700) field_220700, ");
			sqlCmd.append("                sum(a01.field_220800) field_220800, ");
			sqlCmd.append("                sum(a01.field_220900) field_220900, ");
			sqlCmd.append("                sum(a01.field_221000) field_221000, ");
			sqlCmd.append("                sum(a01.field_220000) field_220000, ");
			sqlCmd.append("                sum(a01.field_110321) field_110321, ");
			sqlCmd.append("                sum(a01.field_110322) field_110322, ");
			sqlCmd.append("                sum(a01.field_110323) field_110323, ");
			sqlCmd.append("                sum(a01.field_110324) field_110324, ");
			sqlCmd.append("                sum(a01.field_110325) field_110325, ");
			sqlCmd.append("                sum(a01.field_110326) field_110326, ");
			sqlCmd.append("                sum(a01.field_110327) field_110327, ");
			sqlCmd.append("                sum(a01.field_110320) field_110320, ");
			sqlCmd.append("                sum(a01.field_110301) field_110301, ");
			sqlCmd.append("                sum(a01.field_110302) field_110302, ");
			sqlCmd.append("                sum(a01.field_110303) field_110303, ");
			sqlCmd.append("                sum(a01.field_110304) field_110304, ");
			sqlCmd.append("                sum(a01.field_110305) field_110305, ");
			sqlCmd.append("                sum(a01.field_110306) field_110306, ");
			sqlCmd.append("                sum(a01.field_110307) field_110307, ");
			sqlCmd.append("                sum(a01.field_110310) field_110310, ");
			sqlCmd.append("                sum(a01.field_110311) field_110311, ");
			sqlCmd.append("                sum(a01.field_110312) field_110312, ");
			sqlCmd.append("                sum(a01.field_110313) field_110313, ");
			sqlCmd.append("                sum(a01.field_110300) field_110300  ");
			sqlCmd.append("        from(  ");                            
			sqlCmd.append("                  select  nvl(cd01.hsien_id,' ') as hsien_id , ");                                                                    
			sqlCmd.append("                        nvl(cd01.hsien_name,'OTHER') as hsien_name, ");                                                                  
			sqlCmd.append("                        cd01.FR001W_output_order as FR001W_output_order, ");                                                          
			sqlCmd.append("                        bn01.bank_no, bn01.BANK_NAME, ");  
			sqlCmd.append("                        sum(decode(a01.acc_code,'220100',amt,0))  as field_220100   , ");          
			sqlCmd.append("                        sum(decode(a01.acc_code,'220200',amt,0))  as field_220200    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'220300',amt,0))  as field_220300    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'220400',amt,0))  as field_220400    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'220500',amt,0))  as field_220500    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'220600',amt,0))  as field_220600    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'220700',amt,0))  as field_220700    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'220800',amt,0))  as field_220800    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'220900',amt,0))  as field_220900    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'221000',amt,0))  as field_221000    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'220000',amt,0))  as field_220000    , ");
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110321',amt,0),'7',decode(a01.acc_code,'110251',amt,0))),'103',sum(decode(a01.acc_code,'110321',amt,0)),0)  as field_110321    , ");
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110322',amt,0),'7',decode(a01.acc_code,'110252',amt,0))),'103',sum(decode(a01.acc_code,'110322',amt,0)),0)  as field_110322    , ");        
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110323',amt,0),'7',decode(a01.acc_code,'110253',amt,0))),'103',sum(decode(a01.acc_code,'110323',amt,0)),0)  as field_110323    , ");        
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110324',amt,0),'7',decode(a01.acc_code,'110254',amt,0))),'103',sum(decode(a01.acc_code,'110324',amt,0)),0)  as field_110324    , ");        
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110325',amt,0),'7',decode(a01.acc_code,'110255',amt,0))),'103',sum(decode(a01.acc_code,'110325',amt,0)),0)  as field_110325    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'110326',amt,0))     as field_110326    , ");        
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110327',amt,0),'7',decode(a01.acc_code,'110256',amt,0))),'103',sum(decode(a01.acc_code,'110327',amt,0)),0)  as field_110327    , ");        
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110320',amt,0),'7',decode(a01.acc_code,'110250',amt,0))),'103',sum(decode(a01.acc_code,'110320',amt,0)),0)  as field_110320    , ");        
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110301',amt,0),'7',decode(a01.acc_code,'110211',amt,0))),'103',sum(decode(a01.acc_code,'110301',amt,0)),0)  as field_110301    , ");        
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110302',amt,0),'7',decode(a01.acc_code,'110212',amt,0))),'103',sum(decode(a01.acc_code,'110302',amt,0)),0)  as field_110302    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110303',amt,0),'7',decode(a01.acc_code,'110213',amt,0))),'103',sum(decode(a01.acc_code,'110303',amt,0)),0)  as field_110303    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110304',amt,0),'7',decode(a01.acc_code,'110214',amt,0))),'103',sum(decode(a01.acc_code,'110304',amt,0)),0)  as field_110304    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110305',amt,0),'7',decode(a01.acc_code,'110215',amt,0))),'103',sum(decode(a01.acc_code,'110305',amt,0)),0)  as field_110305    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110306',amt,0),'7',decode(a01.acc_code,'110217',amt,0))),'103',sum(decode(a01.acc_code,'110306',amt,0)),0)  as field_110306    , ");
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110307',amt,0),'7',decode(a01.acc_code,'110216',amt,0))),'103',sum(decode(a01.acc_code,'110307',amt,0)),0)  as field_110307    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110310',amt,0),'7',decode(a01.acc_code,'110210',amt,0))),'103',sum(decode(a01.acc_code,'110310',amt,0)),0)  as field_110310    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110311',amt,0),'7',decode(a01.acc_code,'110220',amt,0))),'103',sum(decode(a01.acc_code,'110311',amt,0)),0)  as field_110311    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110312',amt,0),'7',decode(a01.acc_code,'110230',amt,0))),'103',sum(decode(a01.acc_code,'110312',amt,0)),0)  as field_110312    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110313',amt,0),'7',decode(a01.acc_code,'110240',amt,0))),'103',sum(decode(a01.acc_code,'110313',amt,0)),0)  as field_110313    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110300',amt,0),'7',decode(a01.acc_code,'110200',amt,0))),'103',sum(decode(a01.acc_code,'110300',amt,0)),0)  as field_110300      ");
			sqlCmd.append("                  from(select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y') cd01 ");
			sqlCmd.append("                  left join(select bn01.bank_type,bn01.bank_no,bn01.bank_name,wlx01.hsien_id  ");                    
			sqlCmd.append("                           from (select * from bn01 where m_year=? and bank_type in(?))bn01, (select * from wlx01 where m_year=?)wlx01 "); 
			sqlCmd.append("                          where bn01.bank_no = wlx01.bank_no)bn01 on bn01.hsien_id=cd01.hsien_id ");     
			sqlCmd.append("                  left join (select (CASE WHEN (a01.m_year <= 102) THEN '102' ");
			sqlCmd.append("                                                   WHEN (a01.m_year > 102) THEN '103' ");
			sqlCmd.append("                                                   ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01 ");
			sqlCmd.append("                               where a01.m_year =? and a01.m_month =?  ) a01 on  bn01.bank_no = a01.bank_code ");
			sqlCmd.append("                 group by a01.YEAR_TYPE,nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no ,bn01.BANK_NAME "); 
			sqlCmd.append("          ) a01 "); 
			sqlCmd.append("          where a01.bank_no <> ' ' ");
			sqlCmd.append("          group by a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order,a01.bank_no ,a01.BANK_NAME ");            
			paramList.add(wlx01_m_year);           
            paramList.add(bank_type);
            paramList.add(wlx01_m_year);
            paramList.add(s_year);
            paramList.add(s_month);
			sqlCmd.append("          UNION ALL ");
			//sql_各縣市小計
			sqlCmd.append("          select a01.hsien_id ,a01.hsien_name, a01.FR001W_output_order, ");
			sqlCmd.append("                   ' ' AS  bank_no, ' '   AS  BANK_NAME, "); 
			sqlCmd.append("                   COUNT(*)  AS  COUNT_SEQ, ");
			sqlCmd.append("                   'A90'  as  field_SEQ, ");
			sqlCmd.append("                sum(a01.field_220100) field_220100, ");
			sqlCmd.append("                sum(a01.field_220200) field_220200, ");
			sqlCmd.append("                sum(a01.field_220300) field_220300, ");
			sqlCmd.append("                sum(a01.field_220400) field_220400, ");
			sqlCmd.append("                sum(a01.field_220500) field_220500, ");
			sqlCmd.append("                sum(a01.field_220600) field_220600, ");
			sqlCmd.append("                sum(a01.field_220700) field_220700, ");
			sqlCmd.append("                sum(a01.field_220800) field_220800, ");
			sqlCmd.append("                sum(a01.field_220900) field_220900, ");
			sqlCmd.append("                sum(a01.field_221000) field_221000, ");
			sqlCmd.append("                sum(a01.field_220000) field_220000, ");
			sqlCmd.append("                sum(a01.field_110321) field_110321, ");
			sqlCmd.append("                sum(a01.field_110322) field_110322, ");
			sqlCmd.append("                sum(a01.field_110323) field_110323, ");
			sqlCmd.append("                sum(a01.field_110324) field_110324, ");
			sqlCmd.append("                sum(a01.field_110325) field_110325, ");
			sqlCmd.append("                sum(a01.field_110326) field_110326, ");
			sqlCmd.append("                sum(a01.field_110327) field_110327, ");
			sqlCmd.append("                sum(a01.field_110320) field_110320, ");
			sqlCmd.append("                sum(a01.field_110301) field_110301, ");
			sqlCmd.append("                sum(a01.field_110302) field_110302, ");
			sqlCmd.append("                sum(a01.field_110303) field_110303, ");
			sqlCmd.append("                sum(a01.field_110304) field_110304, ");
			sqlCmd.append("                sum(a01.field_110305) field_110305, ");
			sqlCmd.append("                sum(a01.field_110306) field_110306, ");
			sqlCmd.append("                sum(a01.field_110307) field_110307, ");
			sqlCmd.append("                sum(a01.field_110310) field_110310, ");
			sqlCmd.append("                sum(a01.field_110311) field_110311, ");
			sqlCmd.append("                sum(a01.field_110312) field_110312, ");
			sqlCmd.append("                sum(a01.field_110313) field_110313, ");
			sqlCmd.append("                sum(a01.field_110300) field_110300  ");
			sqlCmd.append("        from(  ");                            
			sqlCmd.append("                  select  nvl(cd01.hsien_id,' ') as hsien_id , ");                                                                    
			sqlCmd.append("                        nvl(cd01.hsien_name,'OTHER') as hsien_name, ");                                                                  
			sqlCmd.append("                        cd01.FR001W_output_order as FR001W_output_order, ");                                                          
			sqlCmd.append("                        bn01.bank_no, bn01.BANK_NAME, ");  
			sqlCmd.append("                        sum(decode(a01.acc_code,'220100',amt,0))  as field_220100   , ");          
			sqlCmd.append("                        sum(decode(a01.acc_code,'220200',amt,0))  as field_220200    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'220300',amt,0))  as field_220300    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'220400',amt,0))  as field_220400    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'220500',amt,0))  as field_220500    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'220600',amt,0))  as field_220600    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'220700',amt,0))  as field_220700    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'220800',amt,0))  as field_220800    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'220900',amt,0))  as field_220900    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'221000',amt,0))  as field_221000    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'220000',amt,0))  as field_220000    , ");
			sqlCmd.append("                      decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110321',amt,0),'7',decode(a01.acc_code,'110251',amt,0))),'103',sum(decode(a01.acc_code,'110321',amt,0)),0)  as field_110321    , ");
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110322',amt,0),'7',decode(a01.acc_code,'110252',amt,0))),'103',sum(decode(a01.acc_code,'110322',amt,0)),0)  as field_110322    , ");        
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110323',amt,0),'7',decode(a01.acc_code,'110253',amt,0))),'103',sum(decode(a01.acc_code,'110323',amt,0)),0)  as field_110323    , ");        
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110324',amt,0),'7',decode(a01.acc_code,'110254',amt,0))),'103',sum(decode(a01.acc_code,'110324',amt,0)),0)  as field_110324    , ");        
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110325',amt,0),'7',decode(a01.acc_code,'110255',amt,0))),'103',sum(decode(a01.acc_code,'110325',amt,0)),0)  as field_110325    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'110326',amt,0))     as field_110326    , ");        
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110327',amt,0),'7',decode(a01.acc_code,'110256',amt,0))),'103',sum(decode(a01.acc_code,'110327',amt,0)),0)  as field_110327    , ");        
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110320',amt,0),'7',decode(a01.acc_code,'110250',amt,0))),'103',sum(decode(a01.acc_code,'110320',amt,0)),0)  as field_110320    , ");        
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110301',amt,0),'7',decode(a01.acc_code,'110211',amt,0))),'103',sum(decode(a01.acc_code,'110301',amt,0)),0)  as field_110301    , ");        
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110302',amt,0),'7',decode(a01.acc_code,'110212',amt,0))),'103',sum(decode(a01.acc_code,'110302',amt,0)),0)  as field_110302    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110303',amt,0),'7',decode(a01.acc_code,'110213',amt,0))),'103',sum(decode(a01.acc_code,'110303',amt,0)),0)  as field_110303    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110304',amt,0),'7',decode(a01.acc_code,'110214',amt,0))),'103',sum(decode(a01.acc_code,'110304',amt,0)),0)  as field_110304    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110305',amt,0),'7',decode(a01.acc_code,'110215',amt,0))),'103',sum(decode(a01.acc_code,'110305',amt,0)),0)  as field_110305    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110306',amt,0),'7',decode(a01.acc_code,'110217',amt,0))),'103',sum(decode(a01.acc_code,'110306',amt,0)),0)  as field_110306    , ");
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110307',amt,0),'7',decode(a01.acc_code,'110216',amt,0))),'103',sum(decode(a01.acc_code,'110307',amt,0)),0)  as field_110307    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110310',amt,0),'7',decode(a01.acc_code,'110210',amt,0))),'103',sum(decode(a01.acc_code,'110310',amt,0)),0)  as field_110310    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110311',amt,0),'7',decode(a01.acc_code,'110220',amt,0))),'103',sum(decode(a01.acc_code,'110311',amt,0)),0)  as field_110311    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110312',amt,0),'7',decode(a01.acc_code,'110230',amt,0))),'103',sum(decode(a01.acc_code,'110312',amt,0)),0)  as field_110312    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110313',amt,0),'7',decode(a01.acc_code,'110240',amt,0))),'103',sum(decode(a01.acc_code,'110313',amt,0)),0)  as field_110313    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110300',amt,0),'7',decode(a01.acc_code,'110200',amt,0))),'103',sum(decode(a01.acc_code,'110300',amt,0)),0)  as field_110300  ");
			sqlCmd.append("                    from (select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y') cd01 ");
			sqlCmd.append("                  left join(select bn01.bank_type,bn01.bank_no,bn01.bank_name,wlx01.hsien_id  ");                    
			sqlCmd.append("                              from (select * from bn01 where m_year=? and bank_type in (?))bn01, (select * from wlx01 where m_year=?)wlx01 "); 
			sqlCmd.append("                             where bn01.bank_no = wlx01.bank_no)bn01 on bn01.hsien_id=cd01.hsien_id ");     
			sqlCmd.append("                  left join (select (CASE WHEN (a01.m_year <= 102) THEN '102' ");
			sqlCmd.append("                                                   WHEN (a01.m_year > 102) THEN '103' ");
			sqlCmd.append("                                                   ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01 ");
			sqlCmd.append("                               where a01.m_year =? and a01.m_month =?) a01 on  bn01.bank_no = a01.bank_code ");
			sqlCmd.append("                 group by a01.YEAR_TYPE,nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no ,bn01.BANK_NAME "); 
			sqlCmd.append("             ) a01 "); 
			sqlCmd.append("            where a01.bank_no <> ' ' ");
			sqlCmd.append("            group by a01.hsien_id,  a01.hsien_name,  a01.FR001W_output_order ");    
			paramList.add(wlx01_m_year);           
            paramList.add(bank_type);
            paramList.add(wlx01_m_year);
            paramList.add(s_year);
            paramList.add(s_month);
			sqlCmd.append("          UNION ALL ");
			//sql_台灣省小計
			sqlCmd.append("        select ' ' AS hsien_id, '臺灣省' AS hsien_name, '002' AS FR001W_output_order, ");
			sqlCmd.append("                 ' ' AS bank_no,  ' ' AS BANK_NAME, "); 
			sqlCmd.append("                 COUNT(*)  AS  COUNT_SEQ, ");
			sqlCmd.append("                 'A92'  as  field_SEQ, ");
			sqlCmd.append("                sum(a01.field_220100) field_220100, ");
			sqlCmd.append("                sum(a01.field_220200) field_220200, ");
			sqlCmd.append("                sum(a01.field_220300) field_220300, ");
			sqlCmd.append("                sum(a01.field_220400) field_220400, ");
			sqlCmd.append("                sum(a01.field_220500) field_220500, ");
			sqlCmd.append("                sum(a01.field_220600) field_220600, ");
			sqlCmd.append("                sum(a01.field_220700) field_220700, ");
			sqlCmd.append("                sum(a01.field_220800) field_220800, ");
			sqlCmd.append("                sum(a01.field_220900) field_220900, ");
			sqlCmd.append("                sum(a01.field_221000) field_221000, ");
			sqlCmd.append("                sum(a01.field_220000) field_220000, ");
			sqlCmd.append("                sum(a01.field_110321) field_110321, ");
			sqlCmd.append("                sum(a01.field_110322) field_110322, ");
			sqlCmd.append("                sum(a01.field_110323) field_110323, ");
			sqlCmd.append("                sum(a01.field_110324) field_110324, ");
			sqlCmd.append("                sum(a01.field_110325) field_110325, ");
			sqlCmd.append("                sum(a01.field_110326) field_110326, ");
			sqlCmd.append("                sum(a01.field_110327) field_110327, ");
			sqlCmd.append("                sum(a01.field_110320) field_110320, ");
			sqlCmd.append("                sum(a01.field_110301) field_110301, ");
			sqlCmd.append("                sum(a01.field_110302) field_110302, ");
			sqlCmd.append("                sum(a01.field_110303) field_110303, ");
			sqlCmd.append("                sum(a01.field_110304) field_110304, ");
			sqlCmd.append("                sum(a01.field_110305) field_110305, ");
			sqlCmd.append("                sum(a01.field_110306) field_110306, ");
			sqlCmd.append("                sum(a01.field_110307) field_110307, ");
			sqlCmd.append("                sum(a01.field_110310) field_110310, ");
			sqlCmd.append("                sum(a01.field_110311) field_110311, ");
			sqlCmd.append("                sum(a01.field_110312) field_110312, ");
			sqlCmd.append("                sum(a01.field_110313) field_110313, ");
			sqlCmd.append("                sum(a01.field_110300) field_110300  ");
			sqlCmd.append("        from( ");                            
			sqlCmd.append("                  select  nvl(cd01.hsien_id,' ') as hsien_id , ");                                                                    
			sqlCmd.append("                        nvl(cd01.hsien_name,'OTHER') as hsien_name, ");                                                                  
			sqlCmd.append("                        cd01.FR001W_output_order as FR001W_output_order, ");                                                          
			sqlCmd.append("                        bn01.bank_no,  bn01.BANK_NAME, ");  
			sqlCmd.append("                        sum(decode(a01.acc_code,'220100',amt,0))  as field_220100   , ");          
			sqlCmd.append("                        sum(decode(a01.acc_code,'220200',amt,0))  as field_220200    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'220300',amt,0))  as field_220300    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'220400',amt,0))  as field_220400    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'220500',amt,0))  as field_220500    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'220600',amt,0))  as field_220600    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'220700',amt,0))  as field_220700    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'220800',amt,0))  as field_220800    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'220900',amt,0))  as field_220900    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'221000',amt,0))  as field_221000    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'220000',amt,0))  as field_220000    , ");
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110321',amt,0),'7',decode(a01.acc_code,'110251',amt,0))),'103',sum(decode(a01.acc_code,'110321',amt,0)),0)  as field_110321    , ");
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110322',amt,0),'7',decode(a01.acc_code,'110252',amt,0))),'103',sum(decode(a01.acc_code,'110322',amt,0)),0)  as field_110322    , ");        
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110323',amt,0),'7',decode(a01.acc_code,'110253',amt,0))),'103',sum(decode(a01.acc_code,'110323',amt,0)),0)  as field_110323    , ");        
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110324',amt,0),'7',decode(a01.acc_code,'110254',amt,0))),'103',sum(decode(a01.acc_code,'110324',amt,0)),0)  as field_110324    , ");        
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110325',amt,0),'7',decode(a01.acc_code,'110255',amt,0))),'103',sum(decode(a01.acc_code,'110325',amt,0)),0)  as field_110325    , ");        
			sqlCmd.append("                        sum(decode(a01.acc_code,'110326',amt,0))     as field_110326    , ");        
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110327',amt,0),'7',decode(a01.acc_code,'110256',amt,0))),'103',sum(decode(a01.acc_code,'110327',amt,0)),0)  as field_110327    , ");        
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110320',amt,0),'7',decode(a01.acc_code,'110250',amt,0))),'103',sum(decode(a01.acc_code,'110320',amt,0)),0)  as field_110320    , ");        
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110301',amt,0),'7',decode(a01.acc_code,'110211',amt,0))),'103',sum(decode(a01.acc_code,'110301',amt,0)),0)  as field_110301    , ");        
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110302',amt,0),'7',decode(a01.acc_code,'110212',amt,0))),'103',sum(decode(a01.acc_code,'110302',amt,0)),0)  as field_110302    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110303',amt,0),'7',decode(a01.acc_code,'110213',amt,0))),'103',sum(decode(a01.acc_code,'110303',amt,0)),0)  as field_110303    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110304',amt,0),'7',decode(a01.acc_code,'110214',amt,0))),'103',sum(decode(a01.acc_code,'110304',amt,0)),0)  as field_110304    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110305',amt,0),'7',decode(a01.acc_code,'110215',amt,0))),'103',sum(decode(a01.acc_code,'110305',amt,0)),0)  as field_110305    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110306',amt,0),'7',decode(a01.acc_code,'110217',amt,0))),'103',sum(decode(a01.acc_code,'110306',amt,0)),0)  as field_110306    , ");
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110307',amt,0),'7',decode(a01.acc_code,'110216',amt,0))),'103',sum(decode(a01.acc_code,'110307',amt,0)),0)  as field_110307    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110310',amt,0),'7',decode(a01.acc_code,'110210',amt,0))),'103',sum(decode(a01.acc_code,'110310',amt,0)),0)  as field_110310    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110311',amt,0),'7',decode(a01.acc_code,'110220',amt,0))),'103',sum(decode(a01.acc_code,'110311',amt,0)),0)  as field_110311    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110312',amt,0),'7',decode(a01.acc_code,'110230',amt,0))),'103',sum(decode(a01.acc_code,'110312',amt,0)),0)  as field_110312    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110313',amt,0),'7',decode(a01.acc_code,'110240',amt,0))),'103',sum(decode(a01.acc_code,'110313',amt,0)),0)  as field_110313    , ");    
			sqlCmd.append("                        decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110300',amt,0),'7',decode(a01.acc_code,'110200',amt,0))),'103',sum(decode(a01.acc_code,'110300',amt,0)),0)  as field_110300   ");
			sqlCmd.append("                  from (select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y' and cd01.Hsien_div = '2') cd01 ");
			sqlCmd.append("                  left join(select bn01.bank_type,bn01.bank_no,bn01.bank_name,wlx01.hsien_id  ");                    
			sqlCmd.append("                              from (select * from bn01 where m_year=? and bank_type=?)bn01, (select * from wlx01 where m_year=?)wlx01 "); 
			sqlCmd.append("                             where bn01.bank_no = wlx01.bank_no)bn01 on bn01.hsien_id=cd01.hsien_id  ");    
			sqlCmd.append("                  left join (select (CASE WHEN (a01.m_year <= 102) THEN '102' ");
			sqlCmd.append("                                                   WHEN (a01.m_year > 102) THEN '103' ");
			sqlCmd.append("                                                   ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01 ");
			sqlCmd.append("                               where a01.m_year =? and a01.m_month =? ) a01 on  bn01.bank_no = a01.bank_code ");
			sqlCmd.append("                 group by a01.YEAR_TYPE,nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no ,  bn01.BANK_NAME ");
			sqlCmd.append("              ) a01  ");
			sqlCmd.append("             where a01.bank_no <> ' ' ");
			paramList.add(wlx01_m_year);           
            paramList.add(bank_type);
            paramList.add(wlx01_m_year);
            paramList.add(s_year);
            paramList.add(s_month);
			sqlCmd.append("        ) a01  ORDER by    FR001W_output_order, field_SEQ,  hsien_id ,  bank_no ");

			List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,
                                			        "hsien_id,hsien_name,FR001W_output_order,bank_no,BANK_NAME,COUNT_SEQ,field_SEQ,"+
                                			        "field_220100,field_220200,field_220300,field_220400,field_220500,field_220600,"+
                                			        "field_220700,field_220800,field_220900,field_221000,field_220000,field_110321,"+
                                			        "field_110322,field_110323,field_110324,field_110325,field_110326,field_110327,"+
                                			        "field_110320,field_110301,field_110302,field_110303,field_110304,field_110305,"+
                                			        "field_110306,field_110307,field_110310,field_110311,field_110312,field_110313,field_110300");
			System.out.println("dbData.size() ="+dbData.size());
			System.out.println("明細表----------------------------");
			
			int rowNum=4;
			
			DataObject bean = null;
			for(int rowcount=2;rowcount<dbData.size();rowcount++){		
				bean = (DataObject)dbData.get(rowcount);
				for(int cellcount=0;cellcount<33;cellcount++){	
					insertValue = (bean.getValue(field_name[cellcount]) == null)?"":(bean.getValue(field_name[cellcount])).toString();														
					if(cellcount==1 && insertValue.equals(" ")){					
						insertValue = (bean.getValue("hsien_name") == null)?"":(bean.getValue("hsien_name")).toString();	
						flag=1;
					}
					insertCell(insertValue,rowNum,cellcount,wb,row,sheet,cell);
					//if(insertValue.equals("台北市")) rowcount_Taipei=rowcount;		
					//if(insertValue.equals("高雄市"))	rowcount_Kaohsiung=rowcount;
                }//end of cellcount
				if(flag==1){
					flag=0;
					rowNum++;
					insertValue="";
					for(int cellcount=0;cellcount<33;cellcount++){
						insertCell(insertValue,rowNum,cellcount,wb,row,sheet,cell);
					} 
					rowNum++; 
				}else rowNum++;                             
			}//end of rowcount
			
			//列印總計	
			bean = (DataObject)dbData.get(0);
			for(int cellcount=0;cellcount<33;cellcount++){	
				insertValue = (bean.getValue(field_name[cellcount]) == null)?"":(bean.getValue(field_name[cellcount])).toString();
				if(cellcount==1 && insertValue.equals(" ")){					
					insertValue = (bean.getValue("hsien_name") == null)?"":(bean.getValue("hsien_name")).toString();	
					flag=1;
				}
				insertCell(insertValue,rowNum,cellcount,wb,row,sheet,cell);//總計列印.第一次
				if(Integer.parseInt(s_year) <= 99){					
					rowNum=(bank_type.equals("6")?(rowNum+6):(rowNum+5));//101.02.04 fix 區分農/漁會.99年以前總計位置
				}else{
					rowNum=(bank_type.equals("6")?(rowNum+10):(rowNum+8));//101.02.04 fix 區分農/漁會.100年以後總計位置
				}
				insertCell(insertValue,rowNum,cellcount,wb,row,sheet,cell);//總計列印.第二次	
				if(Integer.parseInt(s_year) <= 99){					
					rowNum=(bank_type.equals("6")?(rowNum-6):(rowNum-5));//101.02.04 fix 區分農/漁會.99年以前總計位置
				}else{
					rowNum=(bank_type.equals("6")?(rowNum-10):(rowNum-8));//101.02.04 fix 區分農/漁會.100年以後總計位置
				}
            }//end of 列印總計
			rowNum=rowNum+3;
			
		    for(i=0;i<dbData.size();i++){
		   		bean = (DataObject)dbData.get(i);
		   		if(String.valueOf(bean.getValue("bank_name")).equals(" ")){		   		
		   		   hsien_name = (bean.getValue("hsien_name") == null)?"":(bean.getValue("hsien_name")).toString();		   		   
		   		   if((Integer.parseInt(s_year) <= 99 && (hsien_name.equals("台北市") || hsien_name.equals("高雄市")))/*99年(含)以前.列印台北市.高雄市*/
		   		   || (Integer.parseInt(s_year) >= 100 && (hsien_name.equals("新北市") || hsien_name.equals("臺北市") || hsien_name.equals("桃園市")
		   		       || hsien_name.equals("臺中市") || hsien_name.equals("臺南市") || hsien_name.equals("高雄市") )))/*99.09.23 100年(含)以後.明細表加印新北市/台中市/台南市*/ 	
		   		   {
		   			    for(int cellcount=0;cellcount<33;cellcount++){
					        insertValue = (bean.getValue(field_name[cellcount]) == null)?"":(bean.getValue(field_name[cellcount])).toString();
					        if(cellcount==1 && insertValue.equals(" ")){					
						       insertValue = (bean.getValue("hsien_name") == null)?"":(bean.getValue("hsien_name")).toString();	
						       flag=1;
						    }
					   	    insertCell(insertValue,rowNum,cellcount,wb,row,sheet,cell);				
		                }
		   			    rowNum++; 
		   		   }
		   		}
		   		
		    }	  
		    
		    
			//列印臺灣省	
			bean = (DataObject)dbData.get(1);
			for(int cellcount=0;cellcount<33;cellcount++){	
				insertValue = (bean.getValue(field_name[cellcount]) == null)?"":(bean.getValue(field_name[cellcount])).toString();
				if(cellcount==1 && insertValue.equals(" ")){					
					insertValue = (bean.getValue("hsien_name") == null)?"":(bean.getValue("hsien_name")).toString();	
					flag=1;
				}
				insertCell(insertValue,rowNum,cellcount,wb,row,sheet,cell);			
            }//end of 列印臺灣省
			
	 		HSSFCellStyle cs = wb.createCellStyle();	
	 		cs.setAlignment(HSSFCellStyle.ALIGN_LEFT);
	 		rowNum =rowNum+4;
	 		row = sheet.createRow(rowNum);
	 		cell = row.createCell( (short)0);
	 		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	 		cell.setCellValue("1.資料來源:依各農(漁)會信用部由自有電腦設備或委由相關資訊中心以網際網路傳送之資料彙編");
	 		cell.setCellStyle(cs);
	 		sheet.addMergedRegion(new Region(rowNum,(short)0,rowNum,(short)10));
	 		rowNum++;
	 		row = sheet.createRow(rowNum);
	 		cell = row.createCell( (short)0);
	 		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	 		cell.setCellValue("2.備抵呆帳金額(農會)部份:係放款及催收之備抵呆帳合計數,(漁會)部份係放款及透支與催收之備抵呆帳合計數");
	 		cell.setCellStyle(cs);
	 		sheet.addMergedRegion(new Region(rowNum,(short)0,rowNum,(short)10));
	 		rowNum++;
            row = sheet.createRow(rowNum);
            cell = row.createCell( (short)0);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellValue("3.地區別欄之其他(農會)係指原臺灣省農會，該農會於102年5月22日更名為中華民國農會");
            cell.setCellStyle(cs);
            sheet.addMergedRegion(new Region(rowNum,(short)0,rowNum,(short)10));
	 		rowNum =rowNum+2;
	 		HSSFCellStyle cs1 = cell.getCellStyle();//會套用原本excel所設定的格式		  
	 		row = sheet.createRow(rowNum);
	 		cell = row.createCell( (short)0);
	 		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	 		HSSFFont ft = wb.createFont();
			ft.setFontName("標楷體");
			ft.setFontHeightInPoints((short)12);	
			cs.setFont(ft);
			cell.setCellValue("覆核");
			cell.setCellStyle(cs1);
			cell = row.createCell( (short)2);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("副理");
			cell.setCellStyle(cs1);
			cell = row.createCell( (short)4);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
    		cell.setCellValue("經理");
    		cell.setCellStyle(cs1);
		  
    		HSSFFooter footer=sheet.getFooter();
    		footer.setCenter( "Page:" +HSSFFooter.page() +" of " +HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
			FileOutputStream fout=new FileOutputStream(reportDir+ System.getProperty("file.separator")+ openfile);
			wb.write(fout);
	        //儲存 
	        fout.close();
	        System.out.println("儲存完成");		
		}catch(Exception e){
			System.out.println("RptFR040WA.createRpt Error:"+e+e.getMessage());
		}
		return errMsg;
	}
	
	public static void insertCell(String insertValue,int rowNum,int cellcount,HSSFWorkbook wb,HSSFRow row,HSSFSheet sheet,HSSFCell cell)
	{
			row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
			HSSFCellStyle cs1 = cell.getCellStyle();//會套用原本excel所設定的格式
		  	cell=row.createCell((short)cellcount);
		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		  	cs1.setBorderTop(HSSFCellStyle.BORDER_THIN);
			cs1.setBorderBottom(HSSFCellStyle.BORDER_THIN); 
		    cs1.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		    cs1.setBorderRight(HSSFCellStyle.BORDER_THIN);	 		
			cs1.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
			cell.setCellStyle(cs1);	
			if(cellcount < 2)
				cell.setCellValue(insertValue);
			else	
				cell.setCellValue(Utility.setCommaFormat(insertValue));		
				
			if(rowNum==1&&(cellcount==24||cellcount==25||cellcount==26||cellcount==27))
			{
				cell.setCellValue(insertValue);
			}
		     	
	}
}
