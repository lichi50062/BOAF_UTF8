/*
 * Created on 2005/12/04 by 4180
 * fixed on 2006/01/20 by 4180
 * SQL中的除以金額單位(unit)移除，改於資料送出前處理
 * fixed on 2006/01/24 by 4180
 * 存放比率公式修正
 * 2006.06.14 add 農漁會存放款彙總表 by 2295
 * 2006.06.21 fix 農漁會總表存放比率sql by 2295
 * 2010.03.24 fix 根據查詢年度.100年以後取得新縣市別.100年以前取得舊縣市別 
 * 				使用PreparedStatement;並列印轉換後的SQL;套用QueryDB_SQLParam by 2295
 * 2010.12.01 fix wlx01.bn01 sql by 2295
 * 2012.02.13 fix 部份加總欄位區分農漁會科目 by 2295
 *            add 漁會增加110217存放行庫-合作金庫-公庫存款 2295
 * 2012.06.14 add 103年(含)以後SQL by2968
 * 2013.09.02 fix 110257科目代號改為110256
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

public class RptFR040WW{

    public static String createRpt(String s_year,String s_month,String unit,String bank_type,String rptStyle){

        System.out.println("inpute s_month = "+s_month);
        System.out.println("彙總表開始");	
        String errMsg = "";
		String unit_name="";		
		int i=0;
		int j=0;
		String s_year_last="";
		String s_month_last="";
		String hsien_id_sum[]={"a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p"};		 
		String bank_type_name=(bank_type.equals("6"))?"農會":"漁會";
		String field_name[]={"hsien_name","field_220100","field_220200","field_220300","field_220400","field_220500","field_220600",
				  			 "field_220700","field_220800","field_220900","field_221000","field_220000","field_110321","field_110322",
							 "field_110323","field_110324","field_110325","field_110326","field_110327","field_110320","field_110301",
							 "field_110302","field_110303","field_110304","field_110305","field_110306","field_110307","field_110310",
							 "field_110311","field_110312","field_110313","field_110300"};
	
		String filename="";
		String openfile="";
		String insertValue="";
		String cd01_table = "";
        String wlx01_m_year = "";
		List paramList = new ArrayList();
		StringBuffer sql = new StringBuffer();
		filename="農漁會"+((rptStyle.equals("0"))?"存放款彙總表.xls":"存放款彙明細表.xls");
		openfile = bank_type_name+((rptStyle.equals("0"))?"存放款彙總表.xls":"存放款彙明細表.xls");; 
		
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

      
	  		String report_head="																"+s_year+"年" + s_month + "月"+bank_type_name+"信用部存款及轉存行庫總表";
     
			
			//列印年度
			row=(sheet.getRow(0)==null)? sheet.createRow(0) : sheet.getRow(0);													
			cell=row.getCell((short)0);			
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);			
			if(s_month.equals("0")){
				cell.setCellValue("  					中華民國　"+s_year+"　年度");
			}else {																	
				cell.setCellValue(report_head);							
			}
			
			//列印單位		
			row=(sheet.getRow(1)==null)? sheet.createRow(1) : sheet.getRow(1);						
			cell=row.getCell((short)23);			
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			unit_name = Utility.getUnitName(unit);
			insertValue=" 單位：新台幣"+unit_name+"、％";						
			insertCell(insertValue,1,23,wb,row,sheet,cell,s_year,bank_type);		
			StringBuffer sqlCom = new StringBuffer();
            //組共用sql
            sqlCom.append( " sum(a01.field_220100) field_220100,");
            sqlCom.append( " sum(a01.field_220200) field_220200,");
            sqlCom.append( " sum(a01.field_220300) field_220300,");
            sqlCom.append( " sum(a01.field_220400) field_220400,");
            sqlCom.append( " sum(a01.field_220500) field_220500,");
            sqlCom.append( " sum(a01.field_220600) field_220600,");
            sqlCom.append( " sum(a01.field_220700) field_220700,");
            sqlCom.append( " sum(a01.field_220800) field_220800,");
            sqlCom.append( " sum(a01.field_220900) field_220900,");
            sqlCom.append( " sum(a01.field_221000) field_221000,");
            sqlCom.append( " sum(a01.field_220000) field_220000,");
            sqlCom.append( " sum(a01.field_110321) field_110321,");
            sqlCom.append( " sum(a01.field_110322) field_110322,");
            sqlCom.append( " sum(a01.field_110323) field_110323,");
            sqlCom.append( " sum(a01.field_110324) field_110324,");
            sqlCom.append( " sum(a01.field_110325) field_110325,");
            sqlCom.append( " sum(a01.field_110326) field_110326,");
            sqlCom.append( " sum(a01.field_110327) field_110327,");
            sqlCom.append( " sum(a01.field_110320) field_110320,");
            sqlCom.append( " sum(a01.field_110301) field_110301,");
            sqlCom.append( " sum(a01.field_110302) field_110302,");
            sqlCom.append( " sum(a01.field_110303) field_110303,");
            sqlCom.append( " sum(a01.field_110304) field_110304,");
            sqlCom.append( " sum(a01.field_110305) field_110305,");
            sqlCom.append( " sum(a01.field_110306) field_110306,");
            sqlCom.append( " sum(a01.field_110307) field_110307,");
            sqlCom.append( " sum(a01.field_110310) field_110310,");
            sqlCom.append( " sum(a01.field_110311) field_110311,");
            sqlCom.append( " sum(a01.field_110312) field_110312,");
            sqlCom.append( " sum(a01.field_110313) field_110313,");
            sqlCom.append( " sum(a01.field_110300) field_110300 ");
            sqlCom.append( " from (");                             
            sqlCom.append( " select nvl(cd01.hsien_id,' ')       as  hsien_id ,");                                                                    
            sqlCom.append( " nvl(cd01.hsien_name,'OTHER')  as  hsien_name,");                                                                  
            sqlCom.append( " cd01.FR001W_output_order     as  FR001W_output_order,");                                                          
            sqlCom.append( " bn01.bank_no ,  bn01.BANK_NAME,");        
            sqlCom.append( " round(sum(decode(a01.acc_code,'220100',amt,0)) /?,0)     as field_220100    ,");        
            sqlCom.append( " round(sum(decode(a01.acc_code,'220200',amt,0)) /?,0)     as field_220200    ,");        
            sqlCom.append( " round(sum(decode(a01.acc_code,'220300',amt,0)) /?,0)     as field_220300    ,");        
            sqlCom.append( " round(sum(decode(a01.acc_code,'220400',amt,0)) /?,0)     as field_220400    ,");        
            sqlCom.append( " round(sum(decode(a01.acc_code,'220500',amt,0)) /?,0)     as field_220500    ,");        
            sqlCom.append( " round(sum(decode(a01.acc_code,'220600',amt,0)) /?,0)     as field_220600    ,");        
            sqlCom.append( " round(sum(decode(a01.acc_code,'220700',amt,0)) /?,0)     as field_220700    ,");        
            sqlCom.append( " round(sum(decode(a01.acc_code,'220800',amt,0)) /?,0)     as field_220800    ,");        
            sqlCom.append( " round(sum(decode(a01.acc_code,'220900',amt,0)) /?,0)     as field_220900    ,");        
            sqlCom.append( " round(sum(decode(a01.acc_code,'221000',amt,0)) /?,0)     as field_221000    ,");        
            sqlCom.append( " round(sum(decode(a01.acc_code,'220000',amt,0)) /?,0)     as field_220000    ,");    
            sqlCom.append( " round(decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110321',amt,0),'7',decode(a01.acc_code,'110251',amt,0))),'103',sum(decode(a01.acc_code,'110321',amt,0)),0)/?,0) as field_110321  ,");
            sqlCom.append( " round(decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110322',amt,0),'7',decode(a01.acc_code,'110252',amt,0))),'103',sum(decode(a01.acc_code,'110322',amt,0)),0)/?,0)     as field_110322  ,");    
            sqlCom.append( " round(decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110323',amt,0),'7',decode(a01.acc_code,'110253',amt,0))),'103',sum(decode(a01.acc_code,'110323',amt,0)),0)/?,0)     as field_110323  ,");
            sqlCom.append( " round(decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110324',amt,0),'7',decode(a01.acc_code,'110254',amt,0))),'103',sum(decode(a01.acc_code,'110324',amt,0)),0)/?,0)     as field_110324  ,");
            sqlCom.append( " round(decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110325',amt,0),'7',decode(a01.acc_code,'110255',amt,0))),'103',sum(decode(a01.acc_code,'110325',amt,0)),0)/?,0)     as field_110325  ,");        
            sqlCom.append( " round(sum(decode(a01.acc_code,'110326',amt,0)) /?,0)     as field_110326    ,");        
            sqlCom.append( " round(decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110327',amt,0),'7',decode(a01.acc_code,'110256',amt,0))),'103',sum(decode(a01.acc_code,'110327',amt,0)),0)/?,0)     as field_110327  ,");
            sqlCom.append( " round(decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110320',amt,0),'7',decode(a01.acc_code,'110250',amt,0))),'103',sum(decode(a01.acc_code,'110320',amt,0)),0)/?,0)     as field_110320  ,");    
            sqlCom.append( " round(decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110301',amt,0),'7',decode(a01.acc_code,'110211',amt,0))),'103',sum(decode(a01.acc_code,'110301',amt,0)),0)/?,0)     as field_110301  ,");
            sqlCom.append( " round(decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110302',amt,0),'7',decode(a01.acc_code,'110212',amt,0))),'103',sum(decode(a01.acc_code,'110302',amt,0)),0)/?,0)     as field_110302  ,");
            sqlCom.append( " round(decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110303',amt,0),'7',decode(a01.acc_code,'110213',amt,0))),'103',sum(decode(a01.acc_code,'110303',amt,0)),0)/?,0)     as field_110303  ,");
            sqlCom.append( " round(decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110304',amt,0),'7',decode(a01.acc_code,'110214',amt,0))),'103',sum(decode(a01.acc_code,'110304',amt,0)),0)/?,0)     as field_110304  ,");
            sqlCom.append( " round(decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110305',amt,0),'7',decode(a01.acc_code,'110215',amt,0))),'103',sum(decode(a01.acc_code,'110305',amt,0)),0)/?,0)     as field_110305  ,");
            sqlCom.append( " round(decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110306',amt,0),'7',decode(a01.acc_code,'110217',amt,0))),'103',sum(decode(a01.acc_code,'110306',amt,0)),0)/?,0)     as field_110306  ,");//101.02.02 add 漁會增加110217存放行庫-合作金庫-公庫存款    
            sqlCom.append( " round(decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110307',amt,0),'7',decode(a01.acc_code,'110216',amt,0))),'103',sum(decode(a01.acc_code,'110307',amt,0)),0)/?,0)     as field_110307  ,");
            sqlCom.append( " round(decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110310',amt,0),'7',decode(a01.acc_code,'110210',amt,0))),'103',sum(decode(a01.acc_code,'110310',amt,0)),0)/?,0)     as field_110310  ,");
            sqlCom.append( " round(decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110311',amt,0),'7',decode(a01.acc_code,'110220',amt,0))),'103',sum(decode(a01.acc_code,'110311',amt,0)),0)/?,0)     as field_110311  ,");
            sqlCom.append( " round(decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110312',amt,0),'7',decode(a01.acc_code,'110230',amt,0))),'103',sum(decode(a01.acc_code,'110312',amt,0)),0)/?,0)     as field_110312  ,");    
            sqlCom.append( " round(decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110313',amt,0),'7',decode(a01.acc_code,'110240',amt,0))),'103',sum(decode(a01.acc_code,'110313',amt,0)),0)/?,0)     as field_110313  ,");
            sqlCom.append( " round(decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110300',amt,0),'7',decode(a01.acc_code,'110200',amt,0))),'103',sum(decode(a01.acc_code,'110300',amt,0)),0)/?,0)     as field_110300    ");       
            
            //組SQL_總計
            StringBuffer sqlSubTotalHead = new StringBuffer();
            sqlSubTotalHead.append( " select  ' '  AS  hsien_id ,  '總 計'   AS hsien_name,  '001'  AS FR001W_output_order,");
            sqlSubTotalHead.append( " ' ' AS  bank_no ,     ' '   AS  BANK_NAME,"); 
            sqlSubTotalHead.append( " COUNT(*)  AS  COUNT_SEQ,");
            sqlSubTotalHead.append( " 'A99'  as  field_SEQ,");
            
            StringBuffer sqlSubTotalTail = new StringBuffer();
            sqlSubTotalTail.append( " from  (select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y' ) cd01");
            //cd01.hsien_id,cd01.hsien_name,cd01.fr001w_ouput_order,bn01.bank_no,bn01.bank_name
            //99.12.01
            sqlSubTotalTail.append( " left join(select bn01.bank_type,bn01.bank_no,bn01.bank_name,wlx01.hsien_id ");                   
            sqlSubTotalTail.append( " from (select * from bn01 where m_year=? and bank_type in(?))bn01,(select * from wlx01 where m_year=?)wlx01 ");
            sqlSubTotalTail.append( " where bn01.bank_no = wlx01.bank_no)bn01 on bn01.hsien_id=cd01.hsien_id ");    
            //sqlSubTotalTail.append( " left join (select * from a01 where m_year = ? and m_month = ? ) a01 on  bn01.bank_no = a01.bank_code");
            sqlSubTotalTail.append("           left join (select (CASE WHEN (a01.m_year <= 102) THEN '102' ");
            sqlSubTotalTail.append("                                                   WHEN (a01.m_year > 102) THEN '103' ");
            sqlSubTotalTail.append("                                                   ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01 "); 
            sqlSubTotalTail.append("                               where a01.m_year = ? and a01.m_month = ? ) a01 on  bn01.bank_no = a01.bank_code ");
            sqlSubTotalTail.append( " group by  a01.YEAR_TYPE,nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no ,  bn01.BANK_NAME");
            sqlSubTotalTail.append( " ) a01 ");
            sqlSubTotalTail.append( " where a01.bank_no <> ' '");
            for(int k=1;k<=31;k++){
                paramList.add(unit);
            }
            paramList.add(wlx01_m_year);
            paramList.add(bank_type);
            paramList.add(wlx01_m_year);            
            paramList.add(s_year);
            paramList.add(s_month);
            
            StringBuffer sqlSubTotal = new StringBuffer();          
            sqlSubTotal.append(sqlSubTotalHead);
            sqlSubTotal.append(sqlCom);
            sqlSubTotal.append(sqlSubTotalTail);
        
            //SQL_各縣市_小計
            StringBuffer sqlCountyHead =  new StringBuffer();
            sqlCountyHead.append(" select a01.hsien_id ,a01.hsien_name, a01.FR001W_output_order,");
            sqlCountyHead.append(" ' ' AS  bank_no ,     ' '   AS  BANK_NAME,"); 
            sqlCountyHead.append(" COUNT(*)  AS  COUNT_SEQ,");
            sqlCountyHead.append(" 'A90'  as  field_SEQ,");
                       
            StringBuffer sqlCountyTail =  new StringBuffer();
            sqlCountyTail.append(" from  (select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y') cd01");
            //99.12.01
            sqlCountyTail.append(" left join(select bn01.bank_type,bn01.bank_no,bn01.bank_name,wlx01.hsien_id ");                  
            sqlCountyTail.append(" from (select * from bn01 where m_year=? and bank_type=?)bn01,(select * from wlx01 where m_year=?)wlx01 ");
            sqlCountyTail.append(" where bn01.bank_no = wlx01.bank_no)bn01 on bn01.hsien_id=cd01.hsien_id ");              
            //sqlCountyTail.append(" left join (select * from a01 where m_year =? and m_month = ?) a01 on  bn01.bank_no = a01.bank_code");
            sqlCountyTail.append("           left join (select (CASE WHEN (a01.m_year <= 102) THEN '102' ");
            sqlCountyTail.append("                                                   WHEN (a01.m_year > 102) THEN '103' ");
            sqlCountyTail.append("                                                   ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01 "); 
            sqlCountyTail.append("                               where a01.m_year = ? and a01.m_month = ? ) a01 on  bn01.bank_no = a01.bank_code ");
            sqlCountyTail.append(" group by  a01.YEAR_TYPE,nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no ,bn01.BANK_NAME ) a01 ");
            sqlCountyTail.append(" where a01.bank_no <> ' '");
            sqlCountyTail.append(" group by a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order");
            for(int k=1;k<=31;k++){
                paramList.add(unit);
            }
            paramList.add(wlx01_m_year);
            paramList.add(bank_type);
            paramList.add(wlx01_m_year);           
            paramList.add(s_year);
            paramList.add(s_month);
            
            StringBuffer sqlCounty =  new StringBuffer();
            sqlCounty.append(sqlCountyHead);
            sqlCounty.append(sqlCom);
            sqlCounty.append(sqlCountyTail);
            System.out.println("sqlCounty="+sqlCounty.toString());
            //SQL_臺灣省_小計
            StringBuffer sqlTaiwanHead =  new StringBuffer();
            sqlTaiwanHead.append(" select  ' '  AS  hsien_id ,  '臺灣省'   AS hsien_name,  '025'  AS FR001W_output_order,");
            sqlTaiwanHead.append(" ' ' AS  bank_no ,     ' '   AS  BANK_NAME,"); 
            sqlTaiwanHead.append(" COUNT(*)  AS  COUNT_SEQ,");
            sqlTaiwanHead.append(" 'A92'  as  field_SEQ,");
                       
            StringBuffer sqlTaiwanTail =  new StringBuffer();
            sqlTaiwanTail.append(" from  (select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y' and cd01.Hsien_div = '2') cd01");
            //99.12.01
            sqlTaiwanTail.append(" left join(select bn01.bank_type,bn01.bank_no,bn01.bank_name,wlx01.hsien_id ");                  
            sqlTaiwanTail.append(" from (select * from bn01 where m_year=? and bank_type=?)bn01,(select * from wlx01 where m_year=?)wlx01 ");
            sqlTaiwanTail.append(" where bn01.bank_no = wlx01.bank_no)bn01 on bn01.hsien_id=cd01.hsien_id ");   
            //sqlTaiwanTail.append(" left join (select * from a01 where m_year = ? and m_month = ? ) a01 on  bn01.bank_no = a01.bank_code");
            sqlTaiwanTail.append("           left join (select (CASE WHEN (a01.m_year <= 102) THEN '102' ");
            sqlTaiwanTail.append("                                                   WHEN (a01.m_year > 102) THEN '103' ");
            sqlTaiwanTail.append("                                                   ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01 "); 
            sqlTaiwanTail.append("                               where a01.m_year = ? and a01.m_month = ? ) a01 on  bn01.bank_no = a01.bank_code ");
            sqlTaiwanTail.append(" group by  a01.YEAR_TYPE,nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no ,  bn01.BANK_NAME");
            sqlTaiwanTail.append(" ) a01 ");
            sqlTaiwanTail.append(" where a01.bank_no <> ' '");
            for(int k=1;k<=31;k++){
                paramList.add(unit);
            }
            paramList.add(wlx01_m_year);
            paramList.add(bank_type);
            paramList.add(wlx01_m_year);            
            paramList.add(s_year);
            paramList.add(s_month);
            StringBuffer sqlTaiwan = new StringBuffer();
            sqlTaiwan.append(sqlTaiwanHead);
            sqlTaiwan.append(sqlCom);
            sqlTaiwan.append(sqlTaiwanTail);
        
            //SQL_福建省_小計
            StringBuffer sqlFukienHead = new StringBuffer();
            sqlFukienHead.append(" select  ' '  AS  hsien_id ,  '福建省'   AS hsien_name,  '235'  AS FR001W_output_order,"); 
            sqlFukienHead.append(" ' ' AS  bank_no ,     ' '   AS  BANK_NAME,");  
            sqlFukienHead.append(" COUNT(*)  AS  COUNT_SEQ,"); 
            sqlFukienHead.append(" 'A93'  as  field_SEQ,"); 
                       
            StringBuffer sqlFukienTail = new StringBuffer();
            sqlFukienTail.append(" from  (select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y' and cd01.Hsien_div = '3') cd01"); 
            //99.12.01
            sqlFukienTail.append(" left join(select bn01.bank_type,bn01.bank_no,bn01.bank_name,wlx01.hsien_id ");                  
            sqlFukienTail.append(" from (select * from bn01 where m_year=? and bank_type=?)bn01,(select * from wlx01 where m_year=?)wlx01 ");
            sqlFukienTail.append(" where bn01.bank_no = wlx01.bank_no)bn01 on bn01.hsien_id=cd01.hsien_id ");   
            //sqlFukienTail.append(" left join (select * from a01 where m_year = ? and m_month = ?) a01 on  bn01.bank_no = a01.bank_code");
            sqlFukienTail.append("           left join (select (CASE WHEN (a01.m_year <= 102) THEN '102' ");
            sqlFukienTail.append("                                                   WHEN (a01.m_year > 102) THEN '103' ");
            sqlFukienTail.append("                                                   ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01 "); 
            sqlFukienTail.append("                               where a01.m_year = ? and a01.m_month = ? ) a01 on  bn01.bank_no = a01.bank_code ");
            sqlFukienTail.append(" group by  a01.YEAR_TYPE,nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no ,  bn01.BANK_NAME");
            sqlFukienTail.append(" ) a01 ");
            sqlFukienTail.append(" where a01.bank_no <> ' '");
            for(int k=1;k<=31;k++){
                paramList.add(unit);
            }
            paramList.add(wlx01_m_year);
            paramList.add(bank_type);
            paramList.add(wlx01_m_year);            
            paramList.add(s_year);
            paramList.add(s_month);
             
            StringBuffer sqlFukien = new StringBuffer();
            sqlFukien.append(sqlFukienHead);
            sqlFukien.append(sqlCom);
            sqlFukien.append(sqlFukienTail);
      
            //抓總表資料 
            StringBuffer total_sqlCmd = new StringBuffer();
            total_sqlCmd.append(" select  hsien_id , hsien_name, FR001W_output_order,bank_no , BANK_NAME,COUNT_SEQ,field_SEQ, ");
            total_sqlCmd.append(" field_220100, field_220200, field_220300, field_220400, field_220500, ");
            total_sqlCmd.append(" field_220600, field_220700, field_220800, field_220900, field_221000, ");  
            total_sqlCmd.append(" field_220000, field_110321, field_110322, field_110323, field_110324, ");
            total_sqlCmd.append(" field_110325, field_110326, field_110327, field_110320, field_110301, ");
            total_sqlCmd.append(" field_110302, field_110303, field_110304, field_110305, field_110306, ");
            total_sqlCmd.append(" field_110307, field_110310, field_110311, field_110312, field_110313, ");
            total_sqlCmd.append(" field_110300 ");
            total_sqlCmd.append(" from ( ");
            total_sqlCmd.append(sqlSubTotal);   
            total_sqlCmd.append(" UNION ALL");
            total_sqlCmd.append(sqlCounty);             
            total_sqlCmd.append(" UNION ALL");
            total_sqlCmd.append(sqlTaiwan);             
            total_sqlCmd.append(" UNION ALL");
            total_sqlCmd.append(sqlFukien);
            total_sqlCmd.append(" )  a01 group by a01.hsien_id ,a01.hsien_name, a01.FR001W_output_order, bank_no , BANK_NAME  ");
            total_sqlCmd.append(" ,COUNT_SEQ,field_SEQ,");  
            total_sqlCmd.append("field_220100, field_220200, field_220300, field_220400, field_220500,");  
            total_sqlCmd.append("field_220600, field_220700, field_220800, field_220900, field_221000,");  
            total_sqlCmd.append("field_220000, field_110321, field_110322, field_110323, field_110324,");  
            total_sqlCmd.append("field_110325, field_110326, field_110327, field_110320, field_110301,");  
            total_sqlCmd.append("field_110302, field_110303, field_110304, field_110305, field_110306,");  
            total_sqlCmd.append("field_110307, field_110310, field_110311, field_110312, field_110313, ");
            total_sqlCmd.append(" field_110300 ");
            total_sqlCmd.append("ORDER by    FR001W_output_order, field_SEQ,  hsien_id ,  bank_no");
                         
            List dbData = DBManager.QueryDB_SQLParam(total_sqlCmd.toString(),paramList,"field_220100,field_220200,field_220300,field_220400,field_220500,field_220600,field_220700,field_220800,field_220900,field_221000,field_220000,field_110321,field_110322,field_110323,field_110324,field_110325,field_110326,field_110327,field_110320,field_110301,field_110302,field_110303,field_110304,field_110305,field_110306,field_110307,field_110310,field_110311,field_110312,field_110313,field_110300");       
            System.out.println("dbData.size() ="+dbData.size());
            System.out.println("總表----------------------------");
			
			int rowNum=4;
			int cellcount=0;
			DataObject bean = null;
			for(int rowcount=0;rowcount<dbData.size();rowcount++){			    		       
		    	bean = (DataObject)dbData.get(rowcount);
				for(cellcount=0;cellcount<32;cellcount++){
					insertValue = (bean.getValue(field_name[cellcount]) == null)?"":(bean.getValue(field_name[cellcount])).toString();
					//System.out.println(rowcount+":"+cellcount+insertValue);
					insertCell(insertValue,rowNum,cellcount,wb,row,sheet,cell,s_year,bank_type);				
         	    }//cellcount
				rowNum++;                             
			}//end of rowcount
			

	 		HSSFCellStyle cs = wb.createCellStyle();	
	 		cs.setAlignment(HSSFCellStyle.ALIGN_LEFT);
	 		rowNum =rowNum+2;
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
			System.out.println("RptFR040WW.createRpt Error:"+e+e.getMessage());
		}
		return errMsg;
	}
	
	public static void insertCell(String insertValue,int rowNum,int cellcount,HSSFWorkbook wb,HSSFRow row,HSSFSheet sheet,HSSFCell cell,String s_year,String bank_type)
	{
		row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
			
		HSSFCellStyle cs1 = cell.getCellStyle();//會套用原本excel所設定的格式
		HSSFCellStyle cs = wb.createCellStyle();	
		HSSFCellStyle cs2 = wb.createCellStyle();		
		
		cell=row.createCell((short)cellcount);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);										
		cs1.setBorderTop(HSSFCellStyle.BORDER_THIN);
		cs1.setBorderBottom(HSSFCellStyle.BORDER_THIN); 
		cs1.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		cs1.setBorderRight(HSSFCellStyle.BORDER_THIN);
		cs.setBorderTop(HSSFCellStyle.BORDER_THIN);
		cs.setBorderBottom(HSSFCellStyle.BORDER_THIN); 
		cs.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		cs.setBorderRight(HSSFCellStyle.BORDER_THIN); 
		if(bank_type.equals("6")){//農會
			if(((Integer.parseInt(s_year) <= 99) && ((rowNum==5||rowNum==6||rowNum==7||rowNum==30)&&cellcount==0))
		    || ((Integer.parseInt(s_year) >= 100) && ((rowNum==5||rowNum==6||rowNum==7||rowNum==8||rowNum==9||rowNum==10||rowNum==11||rowNum==27)&&cellcount==0)))	
			{	
				HSSFFont ft = wb.createFont();
				ft.setFontName("標楷體");
				ft.setFontHeightInPoints((short)12);	
				cs.setFont(ft);		
				cs.setAlignment(HSSFCellStyle.ALIGN_CENTER);
				cell.setCellStyle(cs);	
				cell.setCellValue(insertValue);		
			}
			if(((Integer.parseInt(s_year) <= 99) && ((rowNum>=8 && rowNum <=29||rowNum==31)&&cellcount==0)
		    || ((Integer.parseInt(s_year) >= 100) &&  ((rowNum>=12 && rowNum <=26||rowNum==28)&&cellcount==0))))
			{	
				cs1.setAlignment(HSSFCellStyle.ALIGN_RIGHT);			
				cell.setCellStyle(cs1);	
				cell.setCellValue(insertValue);		
			}
		}else{//漁會
			if(((Integer.parseInt(s_year) <= 99) && ((rowNum==5||rowNum==6||rowNum==22)&&cellcount==0))
		    || ((Integer.parseInt(s_year) >= 100) && ((rowNum==5||rowNum==6||rowNum==7||rowNum==8||rowNum==9||rowNum==20)&&cellcount==0)))	
			{	
				HSSFFont ft = wb.createFont();
				ft.setFontName("標楷體");
				ft.setFontHeightInPoints((short)12);	
				cs.setFont(ft);		
				cs.setAlignment(HSSFCellStyle.ALIGN_CENTER);
				cell.setCellStyle(cs);	
				cell.setCellValue(insertValue);		
			}	
			if(((Integer.parseInt(s_year) <= 99) && ((rowNum>=7 && rowNum <=21)&&cellcount==0)
			|| ((Integer.parseInt(s_year) >= 100) &&  ((rowNum>=10 && rowNum <=19)&&cellcount==0))))
			{	
				cs1.setAlignment(HSSFCellStyle.ALIGN_RIGHT);			
				cell.setCellStyle(cs1);	
				cell.setCellValue(insertValue);		
			}
		}
		
		if(((Integer.parseInt(s_year) <= 99) && (rowNum>=4 && rowNum<=31))
        ||  ((Integer.parseInt(s_year) >= 100) && (rowNum>=4 && rowNum<=28)))  				
		{
			if(cellcount>=1 && cellcount<=31)
			{
			cs1.setAlignment(HSSFCellStyle.ALIGN_RIGHT);					
			cell.setCellStyle(cs1);
			cell.setCellValue(Utility.setCommaFormat(insertValue));
		  }
		}
		
		if(rowNum==4&&cellcount==0)
		{	
			HSSFFont ft1 = wb.createFont();
			ft1.setFontName("標楷體");
			ft1.setFontHeightInPoints((short)12);	
			cs2.setFont(ft1);	
			cs2.setAlignment(HSSFCellStyle.ALIGN_LEFT);					
			cell.setCellStyle(cs2);	
			cell.setCellValue(insertValue);		
		}
		
		if(rowNum==1&&cellcount==23)
		{				
			cs2.setBorderTop(HSSFCellStyle.BORDER_NONE);
			cs2.setBorderBottom(HSSFCellStyle.BORDER_THIN); 
			cs2.setBorderLeft(HSSFCellStyle.BORDER_NONE);
			cs2.setBorderRight(HSSFCellStyle.BORDER_NONE);  
			cs2.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
			cell.setCellStyle(cs2);	
			cell.setCellValue(insertValue);						
			cell=row.createCell((short)24);			
			cell.setCellStyle(cs2);		
			cell=row.createCell((short)25);			
			cell.setCellStyle(cs2);	
			cell=row.createCell((short)26);			
			cell.setCellStyle(cs2);											
		}    	
	}  
    
}
