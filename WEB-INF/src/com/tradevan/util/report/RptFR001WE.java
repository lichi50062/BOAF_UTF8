/*
 * 2012.08.08 create 臺閩地區農漁會信用部經營指標 by 2295 
 * 2012.08.09 add 利率若分母為0時,顯示N/A.其餘顯示至小數點2位,不足補0 by 2295
 * 2012.11.13 fix 報表格式 by 2295
 * 2013.06.13 add 103/01以後.A01漁會.套用新科目代號/計算公式 by 2295  
 * 2013.09.17 fix 台灣省農會.歸屬於其他.放在嘉義市後面 by 2295
 * 2013.11.19 add 原QueryDB改套用QueryDB_SQLParam by 2295
 * 2014.12.23 fix 調整桃園市位置 by 2295
 * 2017.08.24 add 農會總表的其他.顯示成中華民國農會(黃稽核來電.暫停辦理) by 2295 
 */
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import java.io.*;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR001WE {
	 static String field_seq = "";
	 static String hsien_name = "";         
	 static String bank_no = "";            
	 static String bank_name = "";          
	 static String count_seq = "";          
	 static String field_debit = ""; //存款總額       
	 static String field_credit = ""; //放款總額      
	 static String field_dc_rate = "";//存放比率      
	 static String field_120700= "";//內部融資餘額        
	 static String field_over = "";//逾放金額         
	 static String field_over_rate = ""; //逾放比率    
	 static String field_320300 = "";//本期損益
	 static String field_310000 = "";//事業資金及公積  
	 static String field_320000 = ""; //盈虧及損益
	 static String field_net = ""; //淨值         
	 static String field_fixnet_rate = "";//固定資產佔淨值比  
	 static String field_150200 = "";//催收款項       
	 static String field_backup = ""; //備抵呆帳
	 static String field_backup_over_rate = "";//備呆占狹義逾期放款比率(備抵呆帳/狹義逾放)
	 static String field_backup_credit_rate = "";//備呆占放總總額比率(備抵呆帳/放款總額)	
	 static DecimalFormat df_md = new DecimalFormat("############0.00");//顯示小數點至第2位,不足者補0
     
	 static File logfile;
     static FileOutputStream logos=null;      
     static BufferedOutputStream logbos = null;
     static PrintStream logps = null;
     static Date nowlog = new Date();
     static SimpleDateFormat logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");        
     static SimpleDateFormat logfileformat = new SimpleDateFormat("yyyyMMddHHmmss");
     static Calendar logcalendar;
     static File logDir = null;
    public static String createRpt(String s_year,String s_month,String unit,String bank_type,HSSFWorkbook wb){
          String errMsg = "";         
          StringBuffer sqlSub = new StringBuffer();
	      StringBuffer sqlSubhead = new StringBuffer();  			//各縣市、臺灣省、福建省(金馬地區)小計之共同開頭
	      StringBuffer sqlSubtail  = new StringBuffer(); 		    //各縣市、臺灣省、福建省(金馬地區)小計之共同結尾
	      StringBuffer sqlSubtail2 = new StringBuffer();
	      StringBuffer sqlSubtail3 = new StringBuffer();
	      StringBuffer sqlTaiwanDiv = new StringBuffer();
	      StringBuffer sqlFukienDiv = new StringBuffer();
	      StringBuffer sqlDivSum = new StringBuffer();
	      StringBuffer sqlTaiwan = new StringBuffer();
	      StringBuffer sqlFukien = new StringBuffer();
	      StringBuffer sqlTotal  = new StringBuffer();
	      StringBuffer sqlDetail = new StringBuffer();
	      StringBuffer sqlDiv =  new StringBuffer();
	      StringBuffer sqlDivtemp = new StringBuffer();
	      StringBuffer sqlCombine = new StringBuffer();
	      StringBuffer sqlSubdiv1 = new StringBuffer();
	      StringBuffer sqlSubdiv2 = new StringBuffer();
	      StringBuffer sqlSubdiv3 = new StringBuffer();
	      StringBuffer sqlSubdiv4 = new StringBuffer();
	      StringBuffer sqlSubdiv5 = new StringBuffer();
	      StringBuffer sqlSubdiv6 = new StringBuffer();
	      StringBuffer sqlSubdiv7 = new StringBuffer();
          List dbData_All = null;
          List dbData_Part = null;
          List dbData_Taiwan = null;          
          String cd01_table = "";
          String wlx01_m_year = "";
  		  List paramList = new ArrayList(); 
  		  List sqlTotal_paramList = new ArrayList();   
  		  List sqlSub_paramList = new ArrayList();     
  		  List sqlTaiwan_paramList = new ArrayList();//台灣省用參數  
  		  List sqlFukien_paramList = new ArrayList();//福健省用參數  
  		  List sqlDetail_paramList = new ArrayList();//明細表用參數  
  		  List sqlDivtemp_paramList = new ArrayList();
  		  List sqlSubtail_paramList   = new ArrayList(); //各縣市、臺灣省、福建省(金馬地區)小計之共同結尾
  		  List sqlSubtail2_paramList  = new ArrayList();
  		  List sqlSubtail3_paramList  = new ArrayList();
  		  List sqlTaiwanDiv_paramList = new ArrayList();
  		  List sqlFukienDiv_paramList = new ArrayList();
  		  List sqlDiv_paramList = new ArrayList();
  		  List sqlCombine_paramList = new ArrayList();
  		  List sqlSubhead_paramList = new ArrayList();
          String bank_type_name="";
          String unit_name = "";
          int rowNum=0;
          reportUtil reportUtil = new reportUtil();
          System.out.println("RptFR001WE.bank_type="+bank_type);
          if(bank_type.equals("ALL")){//95.06.14 add 總表?加農漁會
             bank_type_name = "農漁會";
          }else{
             bank_type_name = (bank_type.equals("6"))?"農會":"漁會";
          }          
          unit_name = Utility.getUnitName(unit);//取得單位名稱
          
      try{
          logDir  = new File(Utility.getProperties("logDir"));
          if(!logDir.exists()){
              if(!Utility.mkdirs(Utility.getProperties("logDir"))){
                 System.out.println("目錄新增失敗");
              }    
          }
          logfile = new File(logDir + System.getProperty("file.separator") + "RptFR001WE.log");
          
          System.out.println("logfile filename="+logDir + System.getProperty("file.separator") +"RptFR001WE.log");
          logos = new FileOutputStream(logfile,true);                         
          logbos = new BufferedOutputStream(logos);
          logps = new PrintStream(logbos);   
         
          
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
            String openfile="臺閩地區"+bank_type_name+"信用部經營指標.xls";
            
            FileInputStream finput = null;
            System.out.println("開啟檔:" + openfile);
            finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+ openfile );
            //設定FileINputStream讀取Excel檔
            POIFSFileSystem fs = new POIFSFileSystem( finput );
            if(fs==null){System.out.println("open 範本檔失敗");} else System.out.println("open 範本檔成功");
            wb = new HSSFWorkbook(fs);
            if(wb==null){System.out.println("open工作表失敗");}else System.out.println("open 工作表 成功");
          
	  		HSSFSheet sheet =null;	  		
	  		sheet = wb.getSheetAt(0);//讀取第一個工作表，宣告其為sheet	  		
	  		if(sheet==null){System.out.println("open sheet 失敗");}else System.out.println("open sheet 成功");
	  		HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
	        //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
	        //sheet.setAutobreaks(true); //自動分頁

	        //設定頁面符合列印大小
	        //sheet.setAutobreaks( false );
	        //ps.setScale( ( short )85 ); //列印縮放百分比	        
			
	        //ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
	  		//wb.setSheetName(0,"test");
	  		//設定表頭 為固定 先設欄的起始再設列的起始
	        //wb.setRepeatingRowsAndColumns(0, 1, 21, 2, 3);
	        finput.close();

            HSSFRow row=null;//宣告一列
	  		HSSFCell cell=null;//宣告一個儲存格

		    //共同sql開頭
	  		sqlSubhead.append( " select a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order, ");
	  		sqlSubhead.append( "        a01.bank_no ,   a01.BANK_NAME,  ");
	  		sqlSubhead.append( "        COUNT_SEQ, field_SEQ,   ");
	  		sqlSubhead.append( "	 	round(field_DEBIT /?,0)  as field_DEBIT,   ");    //2006/01/20 fixed by 4180 除以單位
	  		sqlSubhead.append( " 		round(field_CREDIT /?,0)  as field_CREDIT, ");
	  		sqlSubhead.append( "    	decode(a01.fieldI_Y,0,'N/A',  ");
	  		sqlSubhead.append( "    round( ");
	  		sqlSubhead.append( "          (a01.fieldI_XA                                          + ");
	  		sqlSubhead.append( "           decode(sign(a01.fieldI_XB1 - a01.fieldI_XB2),-1,0,     ");
	  		sqlSubhead.append( "                      (a01.fieldI_XB1 - a01.fieldI_XB2))          + ");
	  		sqlSubhead.append( "           decode(sign(a01.fieldI_XC1 - a01.fieldI_XC2),-1,0,       ");
	  		sqlSubhead.append( "                      (a01.fieldI_XC1 - a01.fieldI_XC2))          + ");
	  		sqlSubhead.append( "           decode(sign(a01.fieldI_XD1 - a01.fieldI_XD2),-1,0,        ");
	  		sqlSubhead.append( "                      (a01.fieldI_XD1 - a01.fieldI_XD2))           + ");
	  		sqlSubhead.append( "           decode(sign(a01.fieldI_XE1 - a01.fieldI_XE2),-1,0,        ");
	  		sqlSubhead.append( "                      (a01.fieldI_XE1 - a01.fieldI_XE2))           - ");
	  		sqlSubhead.append( "           decode(sign(a01.fieldI_XF1 - a01.fieldI_XF3 -  a01.fieldI_XF2),-1,0,"); //2006/01/27 公式修改後新增field_XF3
	  		sqlSubhead.append( "                      (a01.fieldI_XF1 - a01.fieldI_XF3 - a01.fieldI_XF2)) "); //2006/01/27 公式修改後新增field_XF3
	  		sqlSubhead.append( "          )  ");
	  		sqlSubhead.append( "         /    a01.fieldI_Y * 100,2))        as     field_DC_RATE , ");
	  		sqlSubhead.append( "	round(field_120700 /?,0)  as field_120700, ");   //2006/01/20 fixed by 4180 除以單位
	  		sqlSubhead.append( "	round(field_OVER /?,0)  as field_OVER, ");
	  		sqlSubhead.append( "    decode(a01.field_CREDIT,0,'N/A',round(a01.field_OVER /  a01.field_CREDIT *100 ,2))  as   field_OVER_RATE,");
	  		sqlSubhead.append( "	round(field_320300 /?,0)  as field_320300, "); //2006/01/25 fixed by 4180 除以單位
	  		sqlSubhead.append( "	round(field_310000 /?,0)  as field_310000, ");         //2006/01/20 fixed by 4180 除以單位
	  		sqlSubhead.append( "    round(field_320000 /?,0)  as field_320000, ");//盈虧及損益 
	  		sqlSubhead.append( "	round(field_NET /?,0)  as field_NET, ");
	  		sqlSubhead.append( "    decode(field_NET,0,'N/A',round(a01.field_140000 /  a01.field_NET *100 ,2))  as   field_FIXNET_RATE,  ");
	  		sqlSubhead.append( "	round(field_150200 /?,0)  as field_150200, ");
	  		sqlSubhead.append( "	round(field_BACKUP /?,0)  as field_BACKUP, ");
	  		sqlSubhead.append( "    decode(FIELD_OVER,0,'N/A',round(FIELD_BACKUP /  a01.FIELD_OVER *100 ,2))  as   field_BACKUP_OVER_RATE,  ");//101.08.08 add 備呆占狹義逾期放款比率(備抵呆帳/狹義逾放)
	        sqlSubhead.append( "    decode(a01.field_CREDIT,0,'N/A',round(FIELD_BACKUP /  a01.field_CREDIT *100 ,2))  as   field_BACKUP_CREDIT_RATE  ");//101.08.08 add 備呆占放款總額比率(備抵呆帳/放款總額)
	        sqlSubhead_paramList.add(unit);
	        sqlSubhead_paramList.add(unit);
	        sqlSubhead_paramList.add(unit);
	        sqlSubhead_paramList.add(unit);
	        sqlSubhead_paramList.add(unit);
	        sqlSubhead_paramList.add(unit);
	        sqlSubhead_paramList.add(unit);
	        sqlSubhead_paramList.add(unit);
	        sqlSubhead_paramList.add(unit);
	        sqlSubhead_paramList.add(unit);
            //各縣市小計用 
            sqlSubdiv3.append( " from ( select a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order,");
            sqlSubdiv3.append( "               ' ' AS  bank_no ,     ' '   AS  BANK_NAME, ");
            sqlSubdiv3.append( "               COUNT(*)  AS  COUNT_SEQ, ");
            sqlSubdiv3.append( "		       'A90'  as  field_SEQ,  ");
           
            //總計用             
            sqlSubdiv4.append( " from ( select ' '  AS  hsien_id ,  '總計'   AS hsien_name,  '001'  AS FR001W_output_order, ");
            sqlSubdiv4.append( "               ' ' AS  bank_no ,     ' '   AS  BANK_NAME, ");
            sqlSubdiv4.append( "               COUNT(*)  AS  COUNT_SEQ, ");
            sqlSubdiv4.append( "               'A99'  as  field_SEQ,    ");
           
            //臺灣省用
            sqlSubdiv5.append( " from ( select ' '  AS  hsien_id ,  '臺灣省'   AS hsien_name,  '025'  AS FR001W_output_order, ");
            sqlSubdiv5.append( "               ' ' AS  bank_no ,     ' '   AS  BANK_NAME, ");
            sqlSubdiv5.append( "               COUNT(*)  AS  COUNT_SEQ, ");
            sqlSubdiv5.append( " 		        'A92'  as  field_SEQ,   ");
           
            //金門地區用 
            sqlSubdiv6.append( "  from( select ' '  AS  hsien_id ,  '金馬地區'   AS hsien_name,  '235'  AS FR001W_output_order,  ");
            sqlSubdiv6.append( " 		       ' ' AS  bank_no ,     ' '   AS  BANK_NAME, ");
            sqlSubdiv6.append( "  		       COUNT(*)  AS  COUNT_SEQ, "); 
            sqlSubdiv6.append( "  		       'A93'  as  field_SEQ,    ");
           
            //明細用
            sqlSubdiv7.append( " from ( select a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order, ");
            sqlSubdiv7.append( "      	       a01.bank_no ,   a01.BANK_NAME,");
            sqlSubdiv7.append( "               1  AS  COUNT_SEQ,      ");
            sqlSubdiv7.append( " 		       'A01'  as  field_SEQ,   ");			
           
            //SUM共同區段
            sqlDivSum.append( " SUM(field_120700)  field_120700 , ");
            sqlDivSum.append( " SUM(field_320300)  field_320300 , ");
            sqlDivSum.append( " SUM(field_310000)  field_310000,  ");
            sqlDivSum.append( " SUM(field_320000)  field_320000,  ");//101.08.08 add 盈虧及損益
            sqlDivSum.append( " SUM(field_NET)     field_NET,     ");
            sqlDivSum.append( " SUM(field_140000)  field_140000,  ");
            sqlDivSum.append( " SUM(field_150200)  field_150200,  ");
            sqlDivSum.append( " SUM(field_BACKUP)  field_BACKUP,  ");
            sqlDivSum.append( " SUM(field_OVER)    field_OVER,    ");
            sqlDivSum.append( " SUM(field_DEBIT)   field_DEBIT,   ");
            sqlDivSum.append( " SUM(field_CREDIT)  field_CREDIT,  ");
            sqlDivSum.append( " SUM(fieldI_XA)     fieldI_XA,     ");
            sqlDivSum.append( " SUM(fieldI_XB1)    fieldI_XB1,    ");
            sqlDivSum.append( " SUM(fieldI_XB2)    fieldI_XB2,    ");
            sqlDivSum.append( " SUM(fieldI_XC1)    fieldI_XC1,    ");
            sqlDivSum.append( " SUM(fieldI_XC2)    fieldI_XC2,    ");
            sqlDivSum.append( " SUM(fieldI_XD1)    fieldI_XD1,    ");
            sqlDivSum.append( " SUM(fieldI_XD2)    fieldI_XD2,    ");
            sqlDivSum.append( " SUM(fieldI_XE1)    fieldI_XE1,    ");
            sqlDivSum.append( " SUM(fieldI_XE2)    fieldI_XE2,    ");
            sqlDivSum.append( " SUM(fieldI_XF1)    fieldI_XF1,    ");
            sqlDivSum.append( " SUM(fieldI_XF3)    fieldI_XF3,    ");  //2006/01/27 公式修改後新增field_XF3
            sqlDivSum.append( " SUM(fieldI_XF2)    fieldI_XF2,    ");
            sqlDivSum.append( " SUM(fieldI_Y)      fieldI_Y       ");                        
             
            //共有sql
            sqlDiv.append( " from ( select nvl(cd01.hsien_id,' ')       as  hsien_id ,");
            sqlDiv.append( "               nvl(cd01.hsien_name,'OTHER') as  hsien_name,");
            sqlDiv.append( "               cd01.FR001W_output_order     as  FR001W_output_order,");
            sqlDiv.append( "               bn01.bank_no ,  bn01.BANK_NAME,");
            sqlDiv.append( "               sum(decode(a01.acc_code,'120700',amt,0)) as field_120700,");
            sqlDiv.append( "               sum(decode(a01.acc_code,'320300',amt,0)) as field_320300,");
            sqlDiv.append( " sum(decode(a01.acc_code,'310000',amt,0)) as field_310000, ");
		    sqlDiv.append( " sum(decode(bank_type,'6',decode(a01.acc_code,'310000',amt,'320000',amt,0),'7',decode(a01.acc_code,'300000',amt,0),0)) as field_NET,  ");    	   
		    sqlDiv.append("  sum(decode(a01.acc_code,'320000',amt,0)) as field_320000, ");//101.08.08 add 盈虧及損益
            sqlDiv.append( " sum(decode(a01.acc_code,'140000',amt,0)) as field_140000, ");
            sqlDiv.append( " sum(decode(a01.acc_code,'150200',amt,0)) as field_150200, ");
            sqlDiv.append( " sum(decode(a01.acc_code, '120800',amt,'150300',amt,0)) as  field_BACKUP,");
            sqlDiv.append( " sum(decode(a01.acc_code,'990000',amt,0))  as field_OVER,  ");
            sqlDiv.append( " sum(decode(a01.acc_code,'220000',amt,0))  as field_DEBIT, ");
            
            //95.06.14 add 農漁會//95.06.21 fix 存放比率sql
            sqlDiv.append( " sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) as  field_CREDIT, ");
            sqlDiv.append( " sum(decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'120101',amt,'120102',amt, ");                                           
            sqlDiv.append( "                           				                             '120200',amt,'120301',amt,  ");
            sqlDiv.append( "          					                                         '120302',amt,'120700',amt,  ");
            sqlDiv.append( "                               					                     '150200',amt,0) ");
            sqlDiv.append( "                                            ,'7',decode(a01.acc_code,'120101',amt,'120102',amt,  ");
            sqlDiv.append( " 		                       					                     '120300',amt,'120401',amt,  ");
            sqlDiv.append( "      	                     					                     '120402',amt,'120700',amt,  ");
            sqlDiv.append( "          	                 					                     '150200',amt,0),0),");
            sqlDiv.append( "                      '103',decode(a01.acc_code,'120101',amt,'120102',amt, ");                                           
            sqlDiv.append( "                                                '120200',amt,'120301',amt,  ");
            sqlDiv.append( "                                                '120302',amt,'120700',amt,  ");
            sqlDiv.append( "                                                '150200',amt,0),0 ))   as fieldI_XA, ");
            sqlDiv.append( " sum(decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'120401',amt,'120402',amt,0),'7',decode(a01.acc_code,'120201',amt,'120202',amt,0)),"); 
            sqlDiv.append( "		              '103',decode(a01.acc_code,'120401',amt,'120402',amt,0),0))     as fieldI_XB1, ");
            sqlDiv.append( " sum(decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'240305',amt,'251300',amt,'310800',amt, 0),'7',decode(a01.acc_code,'240205',amt, '310800',amt,0)),");
            sqlDiv.append( "                      '103',decode(bank_type,'6',decode(a01.acc_code,'240305',amt,'251300',amt,'310800',amt, 0),'7',decode(a01.acc_code,'240305',amt, '251200',amt,0)),0))  as fieldI_XB2, ");              
            sqlDiv.append( "  sum(decode(a01.acc_code,'120501',amt,'120502',amt,0)) as fieldI_XC1,  ");
            
            //95.06.14 add 農漁會//95.06.21 fix 存放比率sql
            sqlDiv.append( " sum(decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'240301',amt,'240302',amt,'240303',amt,'240304',amt,0),");
            sqlDiv.append( "                                             '7',decode(a01.acc_code,'240201',amt,'240202',amt,'240203',amt,'240204',amt,0)), ");
            sqlDiv.append( "                      '103',decode(a01.acc_code,'240301',amt,'240302',amt,'240303',amt,'240304',amt,0),0))    as fieldI_XC2,  ");                        
            sqlDiv.append( " sum(decode(a01.acc_code,'120600',amt,0)) as fieldI_XD1,  ");
            sqlDiv.append( " sum(decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'240200',amt,0),'7',decode(a01.acc_code,'240300',amt,0),0),");
            sqlDiv.append( "                      '103',decode(a01.acc_code,'240200',amt,0),0))                  as fieldI_XD2,  ");
            sqlDiv.append( " sum(decode(a01.acc_code,'150100',amt,0)) as fieldI_XE1,  ");
            sqlDiv.append( " sum(decode(a01.acc_code,'250100',amt,0)) as fieldI_XE2,  ");
            sqlDiv.append( " sum(decode(a01.acc_code,'310000',amt,'320000',amt,0)) as fieldI_XF1,  ");
            sqlDiv.append( " sum(decode(YEAR_TYPE,'102',decode(a01.acc_code,'310800',amt,0),'103',decode(bank_type,'6',decode(a01.acc_code,'310800',amt,0),'7',0,0),0))    			   as fieldI_XF3,  "); //2006/01/27 公式修改後新增field_XF3
            sqlDiv.append( " sum(decode(a01.acc_code,'140000',amt,0))  as fieldI_XF2,  ");
            sqlDiv.append( " round((sum(decode(a01.acc_code,'220100',amt,'220200',amt,                             ");
            sqlDiv.append( "                                '220300',amt,'220400',amt,                             ");
            sqlDiv.append( "                                '220500',amt,'220600',amt,                             ");
            sqlDiv.append( "                                '220700',amt,'220800',amt,                             ");
            sqlDiv.append( "                                '220900',amt,'221000',amt,0))-                         ");
            sqlDiv.append( "        round(sum(decode(a01.acc_code,'220900',amt,0))/2,0)) /1,0)   as fieldI_Y       ");                  
            sqlDiv.append( " from  (select * from "+cd01_table+" cd01 where cd01.hsien_id <> 'Y' ");
           
  
            sqlTaiwanDiv.append(sqlDiv); //臺灣省和福建省(金馬地區)共用區段只差中間一小部分，故在這邊assign給sqlTaiwanDiv \uFFFDBsqlFukienDiv
            sqlTaiwanDiv.append(" and cd01.Hsien_div = '2'");
            sqlFukienDiv.append(sqlDiv);
            sqlFukienDiv.append(" and cd01.Hsien_div = '3'");
                
            //sqlDivtemp設定參數
            sqlDivtemp.append( " ) cd01 left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = "+wlx01_m_year+ " and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL)");
            if(bank_type.equals("ALL")){
               sqlDivtemp.append( "  left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type in ('6','7') and bn01.m_year = "+wlx01_m_year+" and bn_type <> '2' and wlx01.m_year="+wlx01_m_year);
            }else{
               sqlDivtemp.append( "  left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=? and bn01.m_year = "+wlx01_m_year+" and bn_type <> '2' and wlx01.m_year="+wlx01_m_year);     
               sqlDivtemp_paramList.add(bank_type);
            }
            sqlDivtemp.append( "  left join (select  (CASE WHEN (a01.m_year <= 102) THEN '102'");
            sqlDivtemp.append( "                            WHEN (a01.m_year > 102) THEN '103'");
            sqlDivtemp.append( "                            ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01 where a01.m_year  = ? and a01.m_month  = ?) a01  on  bn01.bank_no = a01.bank_code ");
            sqlDivtemp.append( " group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME ");
            sqlDivtemp_paramList.add(s_year);
            sqlDivtemp_paramList.add(s_month);
           
            sqlDiv.append(sqlDivtemp);
            sqlTaiwanDiv.append(sqlDivtemp);    //臺灣省和福建省(金馬地區)共用區段只差中間一小部分，故在這邊assign給sqlTaiwanDiv \uFFFDBsqlFukienDiv
            sqlFukienDiv.append(sqlDivtemp);
            for(int sqlDivtempi=0;sqlDivtempi<sqlDivtemp_paramList.size();sqlDivtempi++){
            	   sqlDiv_paramList.add(sqlDivtemp_paramList.get(sqlDivtempi));
            	   sqlTaiwanDiv_paramList.add(sqlDivtemp_paramList.get(sqlDivtempi));
            	   sqlFukienDiv_paramList.add(sqlDivtemp_paramList.get(sqlDivtempi));
            }  
           
            //共同的結尾 總計、臺灣省、福建省(金馬地區)                            
		    sqlSubtail.append( " ) a01");		 
		    sqlSubtail.append( " where a01.bank_no <> ' ' ");
		    sqlSubtail.append( " ) a01");
		   
			//各縣市小計的結尾				
			sqlSubtail2.append( " ) a01 ");			
	 		sqlSubtail2.append( " where a01.bank_no <> ' ' ");
			sqlSubtail2.append( " GROUP BY a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order ");
			sqlSubtail2.append( " ) a01  ");
			
			//明細的結尾				
			sqlSubtail3.append( " ) a01,   ");  
 	 	 	sqlSubtail3.append( " (select * from a05 where a05.m_year = ? and a05.m_month = ? and  a05.ACC_code = '91060P') a05 ");
 	 	 	sqlSubtail3_paramList.add(s_year);
 	 	 	sqlSubtail3_paramList.add(s_month);
 			sqlSubtail3.append( " where a01.bank_no=a05.bank_code(+) and a01.bank_no <> ' ' ");
			sqlSubtail3.append( " GROUP BY a01.hsien_id,a01.hsien_name,a01.FR001W_output_order,a01.bank_no,a01.BANK_NAME  ");
			sqlSubtail3.append( " ) a01  ");

			/*end共同的sql段落
			  共有以下共用段落
			  sqlDiv      總表、明細、縣市小計 用
			  sqlTaiwanDiv  臺灣省小計用
			  sqlFukienDiv  福建省(金馬地區)小計用
			  sqlDivSum
			  sqlSubhead
			  sqlSubtail
			*/

            //各縣市、臺灣省、福建省(金馬地區)小計sql------------------------------
            sqlSub.append(sqlSubhead);//參數sqlSubhead
            for(int sqlSubheadi=0;sqlSubheadi<sqlSubhead_paramList.size();sqlSubheadi++){
                sqlSub_paramList.add(sqlSubhead_paramList.get(sqlSubheadi));
            }
            sqlSub.append(sqlSubdiv1);//no param
            sqlSub.append(sqlSubdiv3);//各縣市小計用no param
            sqlSub.append(sqlDivSum);//no param
            //sqlSub.append(" decode(SUM(field_910400),0,0,round(SUM(field_910400) * 100000 /SUM(field_910500),0)) as field_CAPTIAL ");
                   //2007.11.07 fix 淨值佔風險性資產比率總計.縣市別是以SUM(910400)/SUM(910500)去計算.不是以91060P/單位數 by 2295                
                                               
            sqlSub = sqlSub.append(sqlDiv);//參數 sqlDivtemp;
            for(int sqlDivi=0;sqlDivi<sqlDiv_paramList.size();sqlDivi++){
            	sqlSub_paramList.add(sqlDiv_paramList.get(sqlDivi));
            }
            sqlSub = sqlSub.append(sqlSubtail2);//參數sqlSubtail2
            for(int sqlSubtail2i=0;sqlSubtail2i<sqlSubtail2_paramList.size();sqlSubtail2i++){
            	sqlSub_paramList.add(sqlSubtail2_paramList.get(sqlSubtail2i));
            }
            sqlTaiwan.append(sqlSubhead);//參數sqlSubhead
            for(int sqlSubheadi=0;sqlSubheadi<sqlSubhead_paramList.size();sqlSubheadi++){
                sqlTaiwan_paramList.add(sqlSubhead_paramList.get(sqlSubheadi));
            }
            sqlTaiwan.append(sqlSubdiv1);//no param
            sqlTaiwan.append(sqlSubdiv5);//臺灣省用no param
            sqlTaiwan.append(sqlDivSum);//no param
            //sqlTaiwan.append(" decode(SUM(field_910400),0,0,round(SUM(field_910400) * 100000 /SUM(field_910500),0)) as field_CAPTIAL ");//96.11.07 add
            		  //2007.11.07 fix 淨值佔風險性資產比率總計.縣市別是以SUM(910400)/SUM(910500)去計算.不是以91060P/單位數 by 2295
            sqlTaiwan.append(sqlTaiwanDiv);//參數sqlDivtemp
            for(int sqlTaiwanDivi=0;sqlTaiwanDivi<sqlTaiwanDiv_paramList.size();sqlTaiwanDivi++){
            	sqlTaiwan_paramList.add(sqlTaiwanDiv_paramList.get(sqlTaiwanDivi));
            }
            sqlTaiwan.append(sqlSubtail);//參數sqlSubtail
            for(int sqlSubtaili=0;sqlSubtaili<sqlSubtail_paramList.size();sqlSubtaili++){
            	sqlTaiwan_paramList.add(sqlSubtail_paramList.get(sqlSubtaili));
            }
            sqlFukien.append(sqlSubhead);//參數sqlSubhead
            for(int sqlSubheadi=0;sqlSubheadi<sqlSubhead_paramList.size();sqlSubheadi++){
                sqlFukien_paramList.add(sqlSubhead_paramList.get(sqlSubheadi));
            }
            sqlFukien.append(sqlSubdiv1);//no param
            sqlFukien.append(sqlSubdiv6);//福建省(金馬地區)用no param
            sqlFukien.append(sqlDivSum);//no param
            sqlFukien.append(sqlFukienDiv);//參數 sqlDivtemp;
            for(int sqlFukienDivi=0;sqlFukienDivi<sqlFukienDiv_paramList.size();sqlFukienDivi++){
            	sqlFukien_paramList.add(sqlFukienDiv_paramList.get(sqlFukienDivi));
            }
            sqlFukien.append(sqlSubtail);//參數sqlSubtail
            for(int sqlSubtaili=0;sqlSubtaili<sqlSubtail_paramList.size();sqlSubtaili++){
            	sqlFukien_paramList.add(sqlSubtail_paramList.get(sqlSubtaili));
            }
            //總計sql，總表和明細表共用
            sqlTotal.append(sqlSubhead);//參數sqlSubhead
            for(int sqlSubheadi=0;sqlSubheadi<sqlSubhead_paramList.size();sqlSubheadi++){
                sqlTotal_paramList.add(sqlSubhead_paramList.get(sqlSubheadi));
            }
            sqlTotal.append(sqlSubdiv1);//no param
            sqlTotal.append(sqlSubdiv4);//總計用no param
            sqlTotal.append(sqlDivSum);//no param
            sqlTotal.append(sqlDiv);//參數 sqlDivtemp;
            for(int sqlDivi=0;sqlDivi<sqlDiv_paramList.size();sqlDivi++){
            	sqlTotal_paramList.add(sqlDiv_paramList.get(sqlDivi));
            }
            sqlTotal.append(sqlSubtail);//參數sqlSubtail
            for(int sqlSubtaili=0;sqlSubtaili<sqlSubtail_paramList.size();sqlSubtaili++){
            	sqlTotal_paramList.add(sqlSubtail_paramList.get(sqlSubtaili));
            }
            //end 總計sql--------------------------------
            //明細sql-------------------------------------
            sqlDetail.append(sqlSubhead);//參數sqlSubhead
            for(int sqlSubheadi=0;sqlSubheadi<sqlSubhead_paramList.size();sqlSubheadi++){
                sqlDetail_paramList.add(sqlSubhead_paramList.get(sqlSubheadi));
            }
            sqlDetail.append(sqlSubdiv2);//no param
            sqlDetail.append(sqlSubdiv7);//明細用no param
            sqlDetail.append(sqlDivSum);//no param
            //sqlDetail.append(" SUM(nvl(a05.amt,0))  as field_CAPTIAL ");
            sqlDetail.append(sqlDiv);//參數 sqlDivtemp
            for(int sqlDivi=0;sqlDivi<sqlDiv_paramList.size();sqlDivi++){
            	sqlDetail_paramList.add(sqlDiv_paramList.get(sqlDivi));
            }
            sqlDetail.append(sqlSubtail3);//參數sqlSubtail3
            for(int sqlSubtail3i=0;sqlSubtail3i<sqlSubtail3_paramList.size();sqlSubtail3i++){
            	sqlDetail_paramList.add(sqlSubtail3_paramList.get(sqlSubtail3i));
            }
            //end  明細sql----------------------------------
                
            //組合sql---------------------------------------
            sqlCombine.append( " select  ");
            sqlCombine.append( " hsien_id , hsien_name, FR001W_output_order, ");
            sqlCombine.append( " bank_no ,  BANK_NAME, COUNT_SEQ, field_SEQ, ");
            sqlCombine.append( " field_DEBIT, field_CREDIT, ");
            sqlCombine.append( " decode(sign(field_DC_RATE - 0),-1,0,field_DC_RATE) as field_DC_RATE,");//96.04.30存放比率若為負數,以0顯示
            sqlCombine.append( " field_OVER,field_OVER_RATE,field_150200,field_BACKUP, field_backup_over_rate,field_backup_credit_rate,");
            sqlCombine.append( " field_120700,field_FIXNET_RATE,field_NET, field_310000,");
            sqlCombine.append( " field_320000,field_320300 ");//101.08.08
            sqlCombine.append( " from (             ");
                               
         	//總表
         	sqlCombine.append(sqlTotal);
         	sqlCombine.append(" UNION ALL");
         	sqlCombine.append(sqlSub);
         	sqlCombine.append(" UNION ALL ");
         	sqlCombine.append(sqlTaiwan); 
         	sqlCombine.append(" UNION ALL ");
            sqlCombine.append(sqlFukien);
            sqlCombine.append(" )  a01  ORDER by FR001W_output_order, field_SEQ, hsien_id, bank_no");
            //add sqlTotal參數
            for(int sqlTotali=0;sqlTotali<sqlTotal_paramList.size();sqlTotali++){
            	sqlCombine_paramList.add(sqlTotal_paramList.get(sqlTotali));
            }
            //add sqlSub參數
            for(int sqlSubi=0;sqlSubi<sqlSub_paramList.size();sqlSubi++){
            	sqlCombine_paramList.add(sqlSub_paramList.get(sqlSubi));
            }
            //add sqlTaiwan參數
            for(int sqlTaiwani=0;sqlTaiwani<sqlTaiwan_paramList.size();sqlTaiwani++){
            	sqlCombine_paramList.add(sqlTaiwan_paramList.get(sqlTaiwani));
            }
            //sqlFukien參數
            for(int sqlFukieni=0;sqlFukieni<sqlFukien_paramList.size();sqlFukieni++){
            	sqlCombine_paramList.add(sqlFukien_paramList.get(sqlFukieni));
            }
            
                    
            //end組合sql------------------------------------
			//建表開始--------------------------------------
			//總表			
     		System.out.println("總表sql="+sqlCombine);
     		printLog(logps,"總表sql="+sqlCombine);
	 		dbData_All = DBManager.QueryDB_SQLParam(sqlCombine.toString(),sqlCombine_paramList,
	 		                           "hsien_name,count_seq,field_debit,field_credit,field_dc_rate,field_120700,field_over,"+
		                               "field_over_rate,field_320300,field_310000,field_net,field_fixnet_rate,field_150200,"+
		                               "field_backup,field_backup_over_rate,field_backup_credit_rate,field_320000");

	 		System.out.print("總表資料 共"+dbData_All.size()+"筆");
	
	 		HSSFFont ft = wb.createFont();
     		HSSFCellStyle cs = wb.createCellStyle();
	 		ft.setFontHeightInPoints((short)18);
	 		ft.setFontName("標楷體");
	 		cs.setFont(ft);
	 		cs.setAlignment(HSSFCellStyle.ALIGN_CENTER);

	   	   	row = sheet.getRow(2);
          	cell = row.getCell( (short) 0);
	   	   	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	   	   	cell.setCellValue( "臺閩地區"+bank_type_name +"信用部經營指標");
	   	   	cell.setCellStyle(cs);
                     
	   	   	row = sheet.getRow(3);             

	   	   	cell = row.getCell( (short)0);
	   	   	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	   	   	cell.setCellValue("中華民國"+s_year+"年"+s_month+"月");
                    
	   	   	cell = row.getCell( (short)15);
	   	   	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	   	   	cell.setCellValue("單位:新臺幣" + unit_name + "、%");
                     
            rowNum=6;
            DataObject bean = null;
			String insertValue = "";
			double sampleValue;
       		for(int i=0;i<dbData_All.size();i++){
	   		 	bean = (DataObject)dbData_All.get(i);	
	   		    getBeanData(bean);//取得各欄位內容
	   		    for(int rowcount=6;rowcount <= 32;rowcount++){
	   		        //System.out.println("hsien_name="+hsien_name);
	   		        row = sheet.getRow(rowcount);
	   		        cell = row.getCell( (short)0);
	   		        if("其他".equals(hsien_name)){//106.08.24 農會總表的其他.顯示成中華民國農會
                       hsien_name = "中華民國農會";
                    }
	   		        if(cell.getStringCellValue().indexOf(hsien_name) != -1 && Integer.parseInt(count_seq) > 0){//找到對應的縣市別
	   		 	       for(int cellcount=1;cellcount<16;cellcount++){ 		    	       	
      			       	cell = row.getCell( (short)cellcount);
     			       	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				       	insertValue = "";
				       	insertValue = setInsertValue(cellcount);
				       	insertValue = (insertValue.equals("null"))?"":insertValue;
				       	cell.setCellValue(insertValue);	    		           
				       }//end of cellcount
	   		        }//end of hsien_name
	   		    }//end of rowcount
			}//end of dbData_All
			
 		    //建表結束--------------------------------------
		   
		    //HSSFFooter footer = sheet.getFooter();
		    //footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
		    //footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
		    FileOutputStream fout=new FileOutputStream(reportDir + System.getProperty("file.separator")+"臺閩地區農漁會信用部經營指標.xls");
		    wb.write(fout);//儲存
		    fout.close();
		    System.out.println("儲存完成");
		    
      }catch(Exception e){
                System.out.println("createRpt Error:"+e+e.getMessage());
      }
      return errMsg;
    }//end of createRpt
    
    //98.07.23 取得各欄位data
    private static void getBeanData(DataObject bean){
    	try{
    		
        	hsien_name = bean.getValue("hsien_name") == null?"":String.valueOf(bean.getValue("hsien_name"));//單位名稱
            bank_no = bean.getValue("bank_no") == null?"":String.valueOf(bean.getValue("bank_no"));
            bank_name = bean.getValue("bank_name") == null?"":String.valueOf(bean.getValue("bank_name"));
            count_seq = bean.getValue("count_seq") == null?"":String.valueOf(bean.getValue("count_seq"));
            field_seq = bean.getValue("field_seq") == null?"":String.valueOf(bean.getValue("field_seq"));
            field_debit = bean.getValue("field_debit") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_debit")));//存款總額
            field_credit = bean.getValue("field_credit") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_credit")));//放款總額
            field_dc_rate = bean.getValue("field_dc_rate") == null?"":String.valueOf(bean.getValue("field_dc_rate"));//存放比率
            field_120700 = bean.getValue("field_120700") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_120700")));//內部融資餘額
            field_over = bean.getValue("field_over") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_over")));//逾放金額
            field_over_rate = bean.getValue("field_over_rate") == null?"":String.valueOf(bean.getValue("field_over_rate"));//逾放比率
            field_320300 = bean.getValue("field_320300") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_320300")));//本期損益
            field_310000 = bean.getValue("field_310000") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_310000")));//事業資金及公積
            field_320000 = bean.getValue("field_320000") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_320000")));//盈虧及損益
            field_net = bean.getValue("field_net") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_net")));//淨值
            field_fixnet_rate = bean.getValue("field_fixnet_rate") == null?"":String.valueOf(bean.getValue("field_fixnet_rate"));//固定資產佔淨值比
            field_150200 = bean.getValue("field_150200") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_150200")));//催收款項
            field_backup = bean.getValue("field_backup") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_backup")));//備抵呆帳
            field_backup_over_rate = bean.getValue("field_backup_over_rate") == null?"":String.valueOf(bean.getValue("field_backup_over_rate"));//備呆占狹義逾期放款比率(備抵呆帳/狹義逾放)
            field_backup_credit_rate = bean.getValue("field_backup_credit_rate") == null?"":String.valueOf(bean.getValue("field_backup_credit_rate"));//備呆占放總總額比率(備抵呆帳/放款總額)
            //利率顯示小數點至第2位,不足者補0
            field_dc_rate = field_dc_rate.equals("N/A") ? "N/A" : Utility.setCommaFormat(df_md.format(Double.parseDouble(field_dc_rate)));
            field_over_rate = field_over_rate.equals("N/A") ? "N/A" : Utility.setCommaFormat(df_md.format(Double.parseDouble(field_over_rate)));
            field_fixnet_rate = field_fixnet_rate.equals("N/A") ? "N/A" : Utility.setCommaFormat(df_md.format(Double.parseDouble(field_fixnet_rate)));
            field_backup_over_rate = field_backup_over_rate.equals("N/A") ? "N/A" : Utility.setCommaFormat(df_md.format(Double.parseDouble(field_backup_over_rate)));
            field_backup_credit_rate = field_backup_credit_rate.equals("N/A") ? "N/A" : Utility.setCommaFormat(df_md.format(Double.parseDouble(field_backup_credit_rate)));
              
     	}catch(Exception e){
    		System.out.println("getBeanData Error:"+e+e.getMessage());
    	}
    }
    //98.07.23 add set insertValue
   	private static String setInsertValue(int cellcount){
    		 String insertValue = "";
    		    if( cellcount==1 ) insertValue =field_debit;//存款總額
                else if( cellcount==2 ) insertValue =field_credit;//放款總額
                else if( cellcount==3 ) insertValue =field_dc_rate;//存放比率
                else if( cellcount==4 ) insertValue =field_over;//逾放金額
                else if( cellcount==5 ) insertValue =field_over_rate;//逾放比率               
                else if( cellcount==6 ) insertValue =field_150200;//催收款項
                else if( cellcount==7 ) insertValue =field_backup;//備抵呆帳
                else if( cellcount==8 ) insertValue =field_backup_over_rate;//備呆占狹義逾期放款比率(備抵呆帳/狹義逾放)
                else if( cellcount==9 ) insertValue =field_backup_credit_rate;//備呆占放總總額比率(備抵呆帳/放款總額)
                else if( cellcount==10 ) insertValue =field_120700;//內部融資餘額
                else if( cellcount==11 )insertValue =field_fixnet_rate;//固定資產佔淨值比
                else if( cellcount==12 )insertValue =field_net;//淨值
                else if( cellcount==13 )insertValue =field_310000;//事業資金及公積
                else if( cellcount==14 )insertValue =field_320000;//盈虧及損益
                else if( cellcount==15 ) insertValue =field_320300;//本期損益
    		    
             return insertValue;
    }
   	public static void printLog(PrintStream logps,String errRptMsg){
        if(!errRptMsg.equals("")){
           logcalendar = Calendar.getInstance(); 
           nowlog = logcalendar.getTime();
           logps.println(logformat.format(nowlog)+errRptMsg);
           logps.flush();
        }
   }
}
