/*
 * Created on 2011/10/18(農漁會信用部經營概況統計表) by 2295 
 * 2012.08.08 fix 增加顯示縣市/機構英文名稱 by 2295
 * 2012.08.09 add 利率若分母為0時,顯示N/A.其餘顯示至小數點2位,不足補0 by 2295
 * 2012.10.17 add 機構.縣市英文使用cellStyle by 2295
 * 2012.11.14 fix 資料年月英文顯示 by 2295
 * 2012.11.26 fix 調整報表文字大小/最後小計.合計文字跳行 by 2295
 * 2012.12.21 fix 漁會不顯示福建省 by 2295
 * 2013.05.24 fix 機構.英文名稱字型大小更改為8 by 2295
 * 2013.11.19 add 原QueryDB改套用QueryDB_SQLParam by 2295
 * 2014.01.16 add 臺灣省改其他,並增加說明 by 2295
 * 2014.12.23 add 增加顯示桃園市於最下方直轄市統計 by 2295
 * 2017.03.07 fix 調整原台灣省改為其他(包含台灣省及福建省.中華民國農會),福建省合併至其他 by 2295  
 * 2019.04.16 add 調整原台灣省改其他(增加包含其他(中華民國農會)) by 2295
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

public class RptFR001WW {
	static String field_seq = "";
	static String hsien_name = "";  
	static String hsien_english = "";//縣市英文名稱
	static String bank_no = "";            
	static String bank_name = "";   
	static String bank_english = "";//機構英文名稱
	static String count_seq = "";          
	static String field_debit = ""; //存款總額      
	static String field_credit = "";//放款總額
	static String field_net = "";//淨值
	static String field_320300 = "";//本期損益
	
	static String field_over = ""; //狹義逾期放款   
	static String field_840740 = "";//廣義逾期放款     
	static String field_over_rate = "";//狹義逾放比率(狹義逾期放款/放款總額)    
	static String field_840740_rate = "";//廣義逾放比率(廣義逾期放款/放款總額)	
	static String field_backup = "";//備抵呆帳
	static String field_backup_100_rate = "";//備呆占狹義逾期放款比率(備抵呆帳/狹義逾放)
	static String field_840740_100_rate = "";//備呆占廣義逾期放款比率(備抵呆帳/廣義逾放)
	static String field_captial_rate = "";//淨值佔風險性資產比率
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
    public static String createRpt(String s_year,String s_month,String unit,String bank_type,String rptStyle,String febxlsFlag,HSSFWorkbook wb,String showEng){
          String errMsg = "";
          String sqlCmd = "";
          StringBuffer sqlSub = new StringBuffer();
	      StringBuffer sqlSubhead = new StringBuffer();  			//各縣市、臺灣省、福建省小計之共同開頭
	      StringBuffer sqlSubtail  = new StringBuffer(); 		    //各縣市、臺灣省、福建省小計之共同結尾
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
          List dbData_Fukien = null;
          String cd01_table = "";
          String wlx01_m_year = "";
  		  List paramList = new ArrayList(); 
  		  List sqlTotal_paramList = new ArrayList();   
  		  List sqlSub_paramList = new ArrayList();     
  		  List sqlTaiwan_paramList = new ArrayList();//台灣省用參數  
  		  List sqlFukien_paramList = new ArrayList();//福健省用參數  
  		  List sqlDetail_paramList = new ArrayList();//明細表用參數  
  		  List sqlDivtemp_paramList = new ArrayList();
  		  List sqlSubtail_paramList   = new ArrayList(); //各縣市、臺灣省、福建省小計之共同結尾
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
          String data_year="";
          String data_month="";
          reportUtil reportUtil = new reportUtil();
          System.out.println("RptFR001WW.bank_type="+bank_type);
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
          logfile = new File(logDir + System.getProperty("file.separator") + "RptFR001WW.log");
          
          System.out.println("logfile filename="+logDir + System.getProperty("file.separator") +"RptFR001WW.log");
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
            String openfile="農漁會信用部經營概況統計表";
            
            FileInputStream finput = null;
            if(febxlsFlag.equals("")){//98.06.16 原主要經營指標
               //openfile+=(rptStyle.equals("0"))?"總表":"明細表_100";
               openfile+=".xls";
               System.out.println("開啟檔:" + openfile);
               finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+ openfile );
               //設定FileINputStream讀取Excel檔
               POIFSFileSystem fs = new POIFSFileSystem( finput );
               if(fs==null){System.out.println("open 範本檔失敗");} else System.out.println("open 範本檔成功");
               wb = new HSSFWorkbook(fs);
               if(wb==null){System.out.println("open工作表失敗");}else System.out.println("open 工作表 成功");
            }
	  		HSSFSheet sheet =null;
	  		if(febxlsFlag.equals("")){//98.06.16 原主要經營指標
	  		   sheet = wb.getSheetAt(0);//讀取第一個工作表，宣告其為sheet
	  		}else{//98.06.16檢查局格式	  			
	  		   sheet = wb.getSheetAt((bank_type.equals("6")?0:1));
	  		}
	  		if(sheet==null){System.out.println("open sheet 失敗");}else System.out.println("open sheet 成功");
	  		HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
	        //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
	        //sheet.setAutobreaks(true); //自動分頁

	        //設定頁面符合列印大小
	        sheet.setAutobreaks( false );
	        if(rptStyle.equals("0"))       ps.setScale( ( short )70 ); //列印縮放百分比
	        else  ps.setScale( ( short )65 ); //列印縮放百分比
			
	        ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
	  		//wb.setSheetName(0,"test");
	  		//設定表頭 為固定 先設欄的起始再設列的起始
	        wb.setRepeatingRowsAndColumns(0, 1, 21, 2, 3);

	        if(febxlsFlag.equals("")) finput.close();

            HSSFRow row=null;//宣告一列
	  		HSSFCell cell=null;//宣告一個儲存格

		    //共同sql開頭
	  		sqlSubhead.append( " select a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order, ");
	  		sqlSubhead.append( "        a01.bank_no ,   a01.BANK_NAME,  a01.english, ");
	  		sqlSubhead.append( "        COUNT_SEQ, field_SEQ,   ");
	  		sqlSubhead.append( "	 	round(field_DEBIT /?,0)  as field_DEBIT,   ");//存款總額
	  		sqlSubhead.append( " 		round(field_CREDIT /?,0)  as field_CREDIT, ");//放款總額	  		
	  		sqlSubhead.append( "		round(field_OVER /?,0)  as field_OVER, ");//狹義逾期放款
	  		sqlSubhead.append( "    	decode(a01.field_CREDIT,0,'N/A',round(a01.field_OVER /  a01.field_CREDIT *100 ,2))  as   field_OVER_RATE,");//狹義逾放比率
	  		sqlSubhead.append( "		round(field_320300 /?,0)  as field_320300, "); //本期損益	  		     //2006/01/20 fixed by 4180 除以單位
	  		sqlSubhead.append( "		round(field_NET /?,0)  as field_NET, ");//淨值
	  		sqlSubhead.append( "		round(field_BACKUP /?,0)  as field_BACKUP, ");//備抵呆帳
	  		sqlSubhead.append( "    	round(field_840740 /?,0)  as field_840740, ");//100.06.22 add 廣義逾放
	  		sqlSubhead.append( "    	decode(field_CREDIT,0,'N/A',round(field_840740 /  a01.field_CREDIT *100 ,2))  as   field_840740_RATE,  ");//100.06.22 add 廣義逾放比(廣義逾放/放款總額)
	  		sqlSubhead.append( "    	decode(FIELD_OVER,0,'N/A',round(FIELD_BACKUP /  a01.FIELD_OVER *100 ,2))  as   field_BACKUP_100_RATE,  ");//100.06.22 add 備呆占狹義逾期放款比率(備抵呆帳/狹義逾放)
	  		sqlSubhead.append( "    	decode(field_840740,0,'N/A',round(FIELD_BACKUP /  field_840740 *100 ,2))  as   field_840740_100_RATE,  ");//100.06.22 add 備呆占廣義逾期放款比率(備抵呆帳/廣義逾放)
	  		sqlSubhead_paramList.add(unit);
	  		sqlSubhead_paramList.add(unit);
	  		sqlSubhead_paramList.add(unit);
	  		sqlSubhead_paramList.add(unit);
	  		sqlSubhead_paramList.add(unit);
	  		sqlSubhead_paramList.add(unit);
	  		sqlSubhead_paramList.add(unit);
            //各縣市、臺灣省、福建省、總計才加
            //2007.11.07 fix 淨值佔風險性資產比率總計.縣市別是以SUM(910400)/SUM(910500)去計算.不是以91060P/單位數 by 2295
            sqlSubdiv1.append( " round(field_CAPTIAL / 1000  ,2)  as   field_CAPTIAL_RATE ");//淨值佔風險性資產比率            
            //明細才加		
            sqlSubdiv2.append( " round(field_CAPTIAL /  1000 ,2)  as   field_CAPTIAL_RATE ");//淨值佔風險性資產比率
            
            //各縣市小計用 
            sqlSubdiv3.append( " from ( select a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order,");
            sqlSubdiv3.append( "               ' ' AS  bank_no ,     ' '   AS  BANK_NAME, '' as english, ");
            sqlSubdiv3.append( "               COUNT(*)  AS  COUNT_SEQ, ");
            sqlSubdiv3.append( "		       'A90'  as  field_SEQ,  ");
           
            //總計用             
            sqlSubdiv4.append( " from ( select ' '  AS  hsien_id ,  ' 總   計 '   AS hsien_name,  '001'  AS FR001W_output_order, ");
            sqlSubdiv4.append( "               ' ' AS  bank_no ,     ' '   AS  BANK_NAME, 'Total' AS english,  ");
            sqlSubdiv4.append( "               COUNT(*)  AS  COUNT_SEQ, ");
            sqlSubdiv4.append( "               'A99'  as  field_SEQ,    ");
           
            //臺灣省用
            //106.03.07 add 臺灣省改為其他縣市(包含台灣省及福建省) 
            sqlSubdiv5.append( " from ( select ' '  AS  hsien_id ,  '其他'   AS hsien_name,  '025'  AS FR001W_output_order, ");
            //sqlSubdiv5.append( " from ( select ' '  AS  hsien_id ,  '臺灣省'   AS hsien_name,  '025'  AS FR001W_output_order, ");
            sqlSubdiv5.append( "               ' ' AS  bank_no ,     ' '   AS  BANK_NAME, 'Taiwan Province' AS english, ");
            sqlSubdiv5.append( "               COUNT(*)  AS  COUNT_SEQ, ");
            sqlSubdiv5.append( " 		        'A92'  as  field_SEQ,   ");
           
            //福建省用 
            sqlSubdiv6.append( "  from( select ' '  AS  hsien_id ,  '福建省'   AS hsien_name,  '235'  AS FR001W_output_order,  ");
            sqlSubdiv6.append( " 		       ' ' AS  bank_no ,     ' '   AS  BANK_NAME, 'Fuchien Province' AS english, ");
            sqlSubdiv6.append( "  		       COUNT(*)  AS  COUNT_SEQ, "); 
            sqlSubdiv6.append( "  		       'A93'  as  field_SEQ,    ");
           
            //明細用
            sqlSubdiv7.append( " from ( select a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order, ");
            sqlSubdiv7.append( "      	       a01.bank_no ,   a01.BANK_NAME, a01.english, ");
            sqlSubdiv7.append( "               1  AS  COUNT_SEQ,      ");
            sqlSubdiv7.append( " 		       'A01'  as  field_SEQ,   ");			
           
            //SUM共同區段
           
            sqlDivSum.append( " SUM(field_320300)  field_320300 ,  ");//本期損益           
            sqlDivSum.append( " SUM(field_NET)     field_NET,       ");//淨值            
            sqlDivSum.append( " SUM(field_BACKUP)  field_BACKUP,    ");//備抵呆帳         
            sqlDivSum.append( " SUM(field_OVER)    field_OVER,      ");//狹義逾期放款
            sqlDivSum.append( " SUM(field_DEBIT)   field_DEBIT,     ");//存款總額
            sqlDivSum.append( " SUM(field_CREDIT)  field_CREDIT,    ");//放款總額                      
            sqlDivSum.append( " SUM(field_840740)  field_840740,   ");//100.06.22 add //廣義逾期放款          
            //共有sql
            sqlDiv.append( " from ( select nvl(cd01.hsien_id,' ')       as  hsien_id ,");
            sqlDiv.append( "               nvl(cd01.hsien_name,'OTHER') as  hsien_name,");
            sqlDiv.append( "               cd01.FR001W_output_order     as  FR001W_output_order,");
            sqlDiv.append( "               bn01.bank_no ,  bn01.BANK_NAME, wlx01.english, ");
            sqlDiv.append( "               round(sum(decode(a01.acc_code,'320300',amt,0)) /1,0) as field_320300,");//本期損益
            if(bank_type.equals("6")){//農會
			  sqlDiv.append( " round(sum(decode(a01.acc_code,'310000',amt,'320000',amt,0)) /1,0)     as field_NET,  ");//淨值
		    }else if(bank_type.equals("7")){//漁會
		   	  sqlDiv.append( " round(sum(decode(a01.acc_code,'300000',amt,0)) /1,0)              as field_NET, ");//淨值
		    }else if(bank_type.equals("ALL")){//95.06.14 add 農漁會
			  sqlDiv.append( " round(sum(decode(bank_type,'6',decode(a01.acc_code,'310000',amt,'320000',amt,0),'7',decode(a01.acc_code,'300000',amt,0),0)) /1,0)     as field_NET,  ");//淨值	   
		    }		    
            if(bank_type.equals("6")){//農會
               sqlDiv.append( " round(sum(decode(a01.acc_code, '120800',amt,'150300',amt,0)) /1,0) as  field_BACKUP,");//備抵呆帳               
            }else if(bank_type.equals("7")){//漁會
               sqlDiv.append( " round(sum(decode(a01.acc_code, '120800',amt,'150300',amt,0)) /1,0) as  field_BACKUP, ");//備抵呆帳               
            }else if(bank_type.equals("ALL")){//95.06.14 add 農漁會//95.06.21 fix 存放比率sql
               sqlDiv.append( " round(sum(decode(a01.acc_code, '120800',amt,'150300',amt,0)) /1,0) as  field_BACKUP, ");//備抵呆帳              
            }
           
            sqlDiv.append( " round(sum(decode(a01.acc_code,'990000',amt,0)) /1,0) as field_OVER,  ");//狹義逾期放款
            sqlDiv.append( " round(sum(decode(a01.acc_code,'220000',amt,0)) /1,0) as field_DEBIT, ");//存款總額
                
            if(bank_type.equals("6")){//農會
               sqlDiv.append( " round(sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /1,0) as  field_CREDIT ");//放款總額              
            }else if(bank_type.equals("7")){//漁會
               sqlDiv.append( " round(sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /1,0) as  field_CREDIT ");//放款總額              
            }else if(bank_type.equals("ALL")){//95.06.14 add 農漁會//95.06.21 fix 存放比率sql
               sqlDiv.append( " round(sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /1,0) as  field_CREDIT ");//放款總額               
            }         
            sqlDiv.append( " from  (select * from "+cd01_table+" cd01 where cd01.hsien_id <> 'Y' ");
           
            //106.03.07 add 臺灣省改為其他縣市(包含台灣省及福建省)
            //108.04.16 add 其他縣市(其他(中華民國農會))
            sqlTaiwanDiv.append(sqlDiv); //臺灣省和福建省共用區段只差中間一小部分，故在這邊assign給sqlTaiwanDiv \uFFFDBsqlFukienDiv
            sqlTaiwanDiv.append(" and cd01.Hsien_div in ('2','3','4')");//2:台灣省 3:福建省 4:其他(中華民國農會)
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
            sqlDivtemp.append( "  left join (select * from a01 where a01.m_year  = ? and a01.m_month  = ?) a01  on  bn01.bank_no = a01.bank_code ");
            sqlDivtemp.append( " group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME, wlx01.english ");
            sqlDivtemp_paramList.add(s_year);
            sqlDivtemp_paramList.add(s_month);
           
            sqlDiv.append(sqlDivtemp);
            sqlTaiwanDiv.append(sqlDivtemp);    //臺灣省和福建省共用區段只差中間一小部分，故在這邊assign給sqlTaiwanDiv \uFFFDBsqlFukienDiv
            sqlFukienDiv.append(sqlDivtemp);
            for(int sqlDivtempi=0;sqlDivtempi<sqlDivtemp_paramList.size();sqlDivtempi++){
            	   sqlDiv_paramList.add(sqlDivtemp_paramList.get(sqlDivtempi));
            	   sqlTaiwanDiv_paramList.add(sqlDivtemp_paramList.get(sqlDivtempi));
            	   sqlFukienDiv_paramList.add(sqlDivtemp_paramList.get(sqlDivtempi));
            }  
           
            //共同的結尾 總計、臺灣省、福建省                            
		    sqlSubtail.append( " ) a01,");
			  //2007.11.07 fix 淨值佔風險性資產比率總計.縣市別是以SUM(910400)/SUM(910500)去計算.不是以91060P/單位數 by 2295       
		    sqlSubtail.append( " (  select nvl(cd01.hsien_id,' ')       as  hsien_id ,");  
		    sqlSubtail.append( "  		    nvl(cd01.hsien_name,'OTHER')  as  hsien_name,");  
		    sqlSubtail.append( "			cd01.FR001W_output_order     as  FR001W_output_order,");  
		    sqlSubtail.append( " 			bn01.bank_no as bank_code,  bn01.bank_name, wlx01.english,");  
		    sqlSubtail.append( "		    round(sum(decode(a05.acc_code,'910400',amt,0)) /1,0) as field_910400,");   
		    sqlSubtail.append( "			round(sum(decode(a05.acc_code,'910500',amt,0)) /1,0) as field_910500 ");                 	  
		    sqlSubtail.append( "	  from  (select * from "+cd01_table+" cd01 where cd01.hsien_id <> 'Y') cd01 ");  
		    sqlSubtail.append( "	  left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = "+wlx01_m_year+ " and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL)");
		    if(bank_type.equals("ALL")){
		       sqlSubtail.append( "  left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type in ('6','7') and bn01.m_year = "+wlx01_m_year+" and bn_type <> '2' and wlx01.m_year = "+wlx01_m_year);
		    }else{
		       sqlSubtail.append( "  left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=? and bn01.m_year = "+wlx01_m_year+" and bn_type <> '2' and wlx01.m_year = "+wlx01_m_year);
		       sqlSubtail_paramList.add(bank_type);
		    }	
		    sqlSubtail.append( "     left join (select * from a05 where a05.m_year  = ? and a05.m_month  = ? and a05.ACC_code in ('910400','910500','91060P') ) a05 ");
		    sqlSubtail_paramList.add(s_year);
		    sqlSubtail_paramList.add(s_month);
            
		    sqlSubtail.append( "               on  bn01.bank_no = a05.bank_code ");
		    sqlSubtail.append( "     group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME, wlx01.english ");
		    sqlSubtail.append( " ) a05 ");  
		    //====================================================================================================
		    //2009.07.23 add A02.990611 by 2295       
		    sqlSubtail.append( " ,(  select nvl(cd01.hsien_id,' ')       as  hsien_id ,");  
		    sqlSubtail.append( "  		    nvl(cd01.hsien_name,'OTHER')  as  hsien_name,");  
		    sqlSubtail.append( "			 cd01.FR001W_output_order     as  FR001W_output_order,");  
		    sqlSubtail.append( " 			 bn01.bank_no as bank_code,  bn01.bank_name, wlx01.english,");  
		    sqlSubtail.append( "		     round(sum(decode(a02.acc_code,'990611',amt,0)) /1,0) as field_990611");
		    sqlSubtail.append( "	  from  (select * from "+cd01_table+" cd01 where cd01.hsien_id <> 'Y') cd01 ");  
		    sqlSubtail.append( "	  left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = "+wlx01_m_year+ " and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL)");		   
		    if(bank_type.equals("ALL")){
			   sqlSubtail.append( "  left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type in ('6','7') and bn01.m_year = "+wlx01_m_year+" and bn_type <> '2' and wlx01.m_year = "+wlx01_m_year);
		    }else{
			   sqlSubtail.append( "  left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=? and bn01.m_year = "+wlx01_m_year+" and bn_type <> '2' and wlx01.m_year = "+wlx01_m_year);     
			   sqlSubtail_paramList.add(bank_type);
		    }
		    sqlSubtail.append( "    left join (select * from a02 where a02.m_year  = ? and a02.m_month  = ? and a02.ACC_code in ('990611') ) a02 ");
		    sqlSubtail_paramList.add(s_year);
		    sqlSubtail_paramList.add(s_month);
            
		    sqlSubtail.append( "               on  bn01.bank_no = a02.bank_code ");
		    sqlSubtail.append( "    group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME, wlx01.english ");
		    sqlSubtail.append( " ) a02 ");
            //2011.06.22 add A04.840740+840760 by 2295       
		    sqlSubtail.append( " ,(  select nvl(cd01.hsien_id,' ')       as  hsien_id ,");  
		    sqlSubtail.append( "  		    nvl(cd01.hsien_name,'OTHER')  as  hsien_name,");  
		    sqlSubtail.append( "			 cd01.FR001W_output_order     as  FR001W_output_order,");  
		    sqlSubtail.append( " 			 bn01.bank_no as bank_code,  bn01.bank_name, wlx01.english,");  
		    sqlSubtail.append( "		     round(sum(decode(a04.acc_code,'840740',amt,'840760',amt,0)) /1,0) as field_840740");//廣義逾期放款
		    sqlSubtail.append( "	  from  (select * from "+cd01_table+" cd01 where cd01.hsien_id <> 'Y') cd01 ");  
		    sqlSubtail.append( "	  left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = "+wlx01_m_year+ " and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL)");		   
		    if(bank_type.equals("ALL")){
			   sqlSubtail.append( "  left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type in ('6','7') and bn01.m_year = "+wlx01_m_year+" and bn_type <> '2' and wlx01.m_year = "+wlx01_m_year);
		    }else{
			   sqlSubtail.append( "  left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=? and bn01.m_year = "+wlx01_m_year+" and bn_type <> '2' and wlx01.m_year = "+wlx01_m_year);     
			   sqlSubtail_paramList.add(bank_type);
		    }
		    sqlSubtail.append( "    left join (select * from a04 where a04.m_year  = ? and a04.m_month  = ? and a04.ACC_code in ('840740','840760') ) a04 ");
		    sqlSubtail_paramList.add(s_year);
		    sqlSubtail_paramList.add(s_month);
            
		    sqlSubtail.append( "               on  bn01.bank_no = a04.bank_code ");
		    sqlSubtail.append( "    group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME, wlx01.english ");
		    sqlSubtail.append( " ) a04 ");
		    sqlSubtail.append( " where  a01.bank_no=a05.bank_code(+) and a01.bank_no=a02.bank_code(+) and a01.bank_no=a04.bank_code(+) and a01.bank_no <> ' ' ");
		    sqlSubtail.append( " ) a01");
		   
			//各縣市小計的結尾				
			sqlSubtail2.append( " ) a01,    ");
                            //2007.11.07 fix 淨值佔風險性資產比率總計.縣市別是以SUM(910400)/SUM(910500)去計算.不是以91060P/單位數 by 2295
			sqlSubtail2.append( " ( select nvl(cd01.hsien_id,' ')       as  hsien_id ,");  
			sqlSubtail2.append( "  		   nvl(cd01.hsien_name,'OTHER')  as  hsien_name,");  
			sqlSubtail2.append( "		   cd01.FR001W_output_order     as  FR001W_output_order,");  
			sqlSubtail2.append( " 		   bn01.bank_no as bank_code,  bn01.bank_name, wlx01.english,");  
			sqlSubtail2.append( "		   round(sum(decode(a05.acc_code,'910400',amt,0)) /1,0) as field_910400,");   
			sqlSubtail2.append( "		   round(sum(decode(a05.acc_code,'910500',amt,0)) /1,0) as field_910500 ");                 	  
			sqlSubtail2.append( "	  from  (select * from "+cd01_table+" cd01 where cd01.hsien_id <> 'Y') cd01 ");  
			sqlSubtail2.append( "	  left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = "+wlx01_m_year+ " and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL)");
			if(bank_type.equals("ALL")){
			   sqlSubtail2.append( "  left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type in ('6','7') and bn01.m_year = "+wlx01_m_year+" and bn_type <> '2' and wlx01.m_year = "+wlx01_m_year);
			}else{
			   sqlSubtail2.append( "  left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=? and bn01.m_year = "+wlx01_m_year+" and bn_type <> '2' and wlx01.m_year = "+wlx01_m_year);    
			   sqlSubtail2_paramList.add(bank_type);
			}
		
			sqlSubtail2.append( "      left join (select * from a05 where a05.m_year = ? and a05.m_month  = ? and a05.ACC_code in ('910400','910500','91060P') ) a05 ");
			sqlSubtail2_paramList.add(s_year);
			sqlSubtail2_paramList.add(s_month);
	        sqlSubtail2.append( "             on  bn01.bank_no = a05.bank_code ");
	        sqlSubtail2.append( "   group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order, bn01.bank_no,bn01.BANK_NAME, wlx01.english ");
	        sqlSubtail2.append( " ) a05 "); 
			//=====================================================================================================
	        //2009.07.23 add A02.990611 by 2295       
			sqlSubtail2.append( " ,(  select nvl(cd01.hsien_id,' ')       as  hsien_id ,");  
			sqlSubtail2.append( "  		     nvl(cd01.hsien_name,'OTHER')  as  hsien_name,");  
			sqlSubtail2.append( "			 cd01.FR001W_output_order     as  FR001W_output_order,");  
			sqlSubtail2.append( " 		     bn01.bank_no as bank_code,  bn01.bank_name, wlx01.english,");  
			sqlSubtail2.append( "		     round(sum(decode(a02.acc_code,'990611',amt,0)) /1,0) as field_990611");
			sqlSubtail2.append( "	  from  (select * from "+cd01_table+" cd01 where cd01.hsien_id <> 'Y') cd01 ");  
			sqlSubtail2.append( "	  left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = "+wlx01_m_year+ " and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL)");			
			if(bank_type.equals("ALL")){
			   sqlSubtail2.append( "  left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type in ('6','7') and bn01.m_year = "+wlx01_m_year+" and bn_type <> '2' and wlx01.m_year = "+wlx01_m_year);
			}else{
			   sqlSubtail2.append( "  left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=? and bn01.m_year = "+wlx01_m_year+" and bn_type <> '2' and wlx01.m_year = "+wlx01_m_year);    
			   sqlSubtail2_paramList.add(bank_type);
			}		
			sqlSubtail2.append( "   left join (select * from a02 where a02.m_year = ? and a02.m_month = ? and a02.ACC_code in ('990611') ) a02 on  bn01.bank_no = a02.bank_code ");
			sqlSubtail2_paramList.add(s_year);
			sqlSubtail2_paramList.add(s_month);
			sqlSubtail2.append( "    group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME, wlx01.english ");
			sqlSubtail2.append( " ) a02 ");
			
			//2011.06.22 add A04.84740+840760 by 2295       
			sqlSubtail2.append( " ,(  select nvl(cd01.hsien_id,' ')       as  hsien_id ,");  
			sqlSubtail2.append( "  		     nvl(cd01.hsien_name,'OTHER')  as  hsien_name,");  
			sqlSubtail2.append( "			 cd01.FR001W_output_order     as  FR001W_output_order,");  
			sqlSubtail2.append( " 		     bn01.bank_no as bank_code,  bn01.bank_name, wlx01.english,");  
			sqlSubtail2.append( "		     round(sum(decode(a04.acc_code,'840740',amt,'840760',amt,0)) /1,0) as field_840740");//廣義逾期放款
			sqlSubtail2.append( "	  from  (select * from "+cd01_table+" cd01 where cd01.hsien_id <> 'Y') cd01 ");  
			sqlSubtail2.append( "	  left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = "+wlx01_m_year+ " and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL)");			
			if(bank_type.equals("ALL")){
			   sqlSubtail2.append( "  left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type in ('6','7') and bn01.m_year = "+wlx01_m_year+" and bn_type <> '2' and wlx01.m_year = "+wlx01_m_year);
			}else{
			   sqlSubtail2.append( "  left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=? and bn01.m_year = "+wlx01_m_year+" and bn_type <> '2' and wlx01.m_year = "+wlx01_m_year);    
			   sqlSubtail2_paramList.add(bank_type);
			}		
			sqlSubtail2.append( "   left join (select * from a04 where a04.m_year = ? and a04.m_month = ? and a04.ACC_code in ('840740','840760') ) a04 on  bn01.bank_no = a04.bank_code ");
			sqlSubtail2_paramList.add(s_year);
			sqlSubtail2_paramList.add(s_month);
			sqlSubtail2.append( "    group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME, wlx01.english ");
			sqlSubtail2.append( " ) a04 ");
			
	 		sqlSubtail2.append( " where a01.bank_no=a05.bank_code(+) and a01.bank_no=a02.bank_code(+) and a01.bank_no=a04.bank_code(+) and a01.bank_no <> ' ' ");
			sqlSubtail2.append( " GROUP BY a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order ");
			sqlSubtail2.append( " ) a01  ");
			
			//明細的結尾				
			sqlSubtail3.append( " ) a01,   ");  
 	 	 	sqlSubtail3.append( " (select * from a05 where a05.m_year = ? and a05.m_month = ? and  a05.ACC_code = '91060P') a05 ");
 	 	 	sqlSubtail3_paramList.add(s_year);
 	 	 	sqlSubtail3_paramList.add(s_month);
 						    //2009.07.23 add A02.990611 by 2295       
 			sqlSubtail3.append( " ,( select nvl(cd01.hsien_id,' ')       as  hsien_id ,");  
 			sqlSubtail3.append( "  		    nvl(cd01.hsien_name,'OTHER')  as  hsien_name,");  
 			sqlSubtail3.append( "			cd01.FR001W_output_order     as  FR001W_output_order,");  
 			sqlSubtail3.append( " 		    bn01.bank_no as bank_code,  bn01.bank_name, wlx01.english,");  
 			sqlSubtail3.append( "		    round(sum(decode(a02.acc_code,'990611',amt,0)) /1,0) as field_990611");
 			sqlSubtail3.append( "	   from  (select * from "+cd01_table+" cd01 where cd01.hsien_id <> 'Y') cd01 ");  
 			sqlSubtail3.append( "	   left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = "+wlx01_m_year+ " and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL)"); 			
 			if(bank_type.equals("ALL")){
 				sqlSubtail3.append( "  left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type in ('6','7') and bn01.m_year = "+wlx01_m_year+" and bn_type <> '2' and wlx01.m_year = "+wlx01_m_year);
 			}else{
 				sqlSubtail3.append( "  left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=? and bn01.m_year = "+wlx01_m_year+" and bn_type <> '2' and wlx01.m_year = "+wlx01_m_year);     
 				sqlSubtail3_paramList.add(bank_type);
 			}
 			sqlSubtail3.append( "    left join (select * from a02 where a02.m_year = ? and a02.m_month = ? and a02.ACC_code in ('990611') ) a02 ");
 			sqlSubtail3_paramList.add(s_year);
 	 	 	sqlSubtail3_paramList.add(s_month);
 			sqlSubtail3.append( "               on  bn01.bank_no = a02.bank_code ");
 			sqlSubtail3.append( "    group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME, wlx01.english ");
 			sqlSubtail3.append( " ) a02 ");
 			
		    //2011.06.22 add A04.84740+840760 by 2295       
 			sqlSubtail3.append( " ,( select nvl(cd01.hsien_id,' ')       as  hsien_id ,");  
 			sqlSubtail3.append( "  		    nvl(cd01.hsien_name,'OTHER')  as  hsien_name,");  
 			sqlSubtail3.append( "			cd01.FR001W_output_order     as  FR001W_output_order,");  
 			sqlSubtail3.append( " 		    bn01.bank_no as bank_code,  bn01.bank_name, wlx01.english,");  
 			sqlSubtail3.append( "		    round(sum(decode(a04.acc_code,'840740',amt,'840760',amt,0)) /1,0) as field_840740");//廣義逾期放款
 			sqlSubtail3.append( "	   from  (select * from "+cd01_table+" cd01 where cd01.hsien_id <> 'Y') cd01 ");  
 			sqlSubtail3.append( "	   left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = "+wlx01_m_year+ " and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL)"); 			
 			if(bank_type.equals("ALL")){
 				sqlSubtail3.append( "  left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type in ('6','7') and bn01.m_year = "+wlx01_m_year+" and bn_type <> '2' and wlx01.m_year = "+wlx01_m_year);
 			}else{
 				sqlSubtail3.append( "  left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=? and bn01.m_year = "+wlx01_m_year+" and bn_type <> '2' and wlx01.m_year = "+wlx01_m_year);     
 				sqlSubtail3_paramList.add(bank_type);
 			}
 			sqlSubtail3.append( "    left join (select * from a04 where a04.m_year = ? and a04.m_month = ? and a04.ACC_code in ('840740','840760') ) a04 ");
 			sqlSubtail3_paramList.add(s_year);
 	 	 	sqlSubtail3_paramList.add(s_month);
 			sqlSubtail3.append( "               on  bn01.bank_no = a04.bank_code ");
 			sqlSubtail3.append( "    group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME, wlx01.english ");
 			sqlSubtail3.append( " ) a04 ");
 			
 			sqlSubtail3.append( " where a01.bank_no=a05.bank_code(+) and a01.bank_no=a02.bank_code(+) and a01.bank_no = a04.bank_code(+) and a01.bank_no <> ' ' ");
			sqlSubtail3.append( " GROUP BY a01.hsien_id,a01.hsien_name,a01.FR001W_output_order,a01.bank_no,a01.BANK_NAME,a01.english  ");
			sqlSubtail3.append( " ) a01  ");

			/*end共同的sql段落
			  共有以下共用段落
			  sqlDiv      總表、明細、縣市小計 用
			  sqlTaiwanDiv  臺灣省小計用
			  sqlFukienDiv  福建省小計用
			  sqlDivSum
			  sqlSubhead
			  sqlSubtail
			*/

            //各縣市、臺灣省、福建省小計sql------------------------------
            sqlSub.append(sqlSubhead);//參數 sqlSubhead;            
            for(int sqlSubheadi=0;sqlSubheadi<sqlSubhead_paramList.size();sqlSubheadi++){
                sqlSub_paramList.add(sqlSubhead_paramList.get(sqlSubheadi));
            }
            sqlSub.append(sqlSubdiv1);//no param
            sqlSub.append(sqlSubdiv3);//各縣市小計用no param
            sqlSub.append(sqlDivSum);//no param
            sqlSub.append(" decode(SUM(field_910400),0,0,round(SUM(field_910400) * 100000 /SUM(field_910500),0)) as field_CAPTIAL ");
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
            sqlTaiwan.append(" decode(SUM(field_910400),0,0,round(SUM(field_910400) * 100000 /SUM(field_910500),0)) as field_CAPTIAL ");//96.11.07 add
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
            sqlFukien.append(sqlSubdiv6);//福建省用no param
            sqlFukien.append(sqlDivSum);//no param
            sqlFukien.append(" decode(SUM(field_910400),0,0,round(SUM(field_910400) * 100000 /SUM(field_910500),0)) as field_CAPTIAL ");//96.11.07 add
                      //2007.11.07 fix 淨值佔風險性資產比率總計.縣市別是以SUM(910400)/SUM(910500)去計算.不是以91060P/單位數 by 2295	   			 
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
            sqlTotal.append(" decode(SUM(field_910400),0,0,round(SUM(field_910400) * 100000 /SUM(field_910500),0)) as field_CAPTIAL ");//96.11.07 add
             		 //2007.11.07 fix 淨值佔風險性資產比率總計.縣市別是以SUM(910400)/SUM(910500)去計算.不是以91060P/單位數 by 2295
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
            sqlDetail.append(" SUM(nvl(a05.amt,0))  as field_CAPTIAL ");
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
            sqlCombine.append( " select                                      ");
            sqlCombine.append( " a01.hsien_id , a01.hsien_name, ");
            sqlCombine.append( " cd01.english as hsien_english, ");//--縣市英文名稱
            sqlCombine.append( " a01.FR001W_output_order, ");
            sqlCombine.append( " a01.bank_no ,  a01.BANK_NAME, ");
            sqlCombine.append( " a01.english as bank_english, ");//--機構英文名稱
            sqlCombine.append( " a01.COUNT_SEQ, a01.field_SEQ, ");
            sqlCombine.append( " field_DEBIT,          field_CREDIT,         ");            
            sqlCombine.append( " field_OVER, field_OVER_RATE,               ");//狹義逾放比率
            sqlCombine.append( " field_320300, field_NET, field_BACKUP,  ");           
            sqlCombine.append( " field_CAPTIAL_RATE,");//淨值佔風險性資產比率
            sqlCombine.append( " field_840740,field_840740_rate,field_backup_100_rate,field_840740_100_rate       ");//98.07.22
            sqlCombine.append( " from (             ");
                               
         	if(rptStyle.equals("0")){//總表	
         		sqlCombine.append(sqlTotal);//總表sql
         		sqlCombine.append(" UNION ALL");
         		sqlCombine.append(sqlSub);
         		sqlCombine.append(" UNION ALL ");
         		sqlCombine.append(sqlTaiwan); //臺灣省sql
         		sqlCombine.append(" UNION ALL ");
                sqlCombine.append(sqlFukien);//福建省sql
                sqlCombine.append(" )  a01 ");
                sqlCombine.append(" left join  (select * from  cd01 where cd01.hsien_id <> 'Y') cd01  on a01.hsien_id = cd01.hsien_id");
                sqlCombine.append(" ORDER by    FR001W_output_order, field_SEQ,  hsien_id ,  bank_no ");
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
            }else{//明細表                	
            	sqlCombine.append(sqlTotal);
            	sqlCombine.append(" UNION ALL");
            	sqlCombine.append(sqlDetail);
            	sqlCombine.append(" UNION ALL");
            	sqlCombine.append(sqlSub);
            	sqlCombine.append(" )  a01 ");
            	sqlCombine.append(" left join  (select * from  cd01 where cd01.hsien_id <> 'Y') cd01  on a01.hsien_id = cd01.hsien_id");
            	sqlCombine.append(" ORDER by    FR001W_output_order, field_SEQ,  hsien_id ,  bank_no ");
            	//add sqlTotal參數
            	for(int sqlTotali=0;sqlTotali<sqlTotal_paramList.size();sqlTotali++){
                	sqlCombine_paramList.add(sqlTotal_paramList.get(sqlTotali));
                }
            	//add sqlDetail參數
            	for(int sqlDetaili=0;sqlDetaili<sqlDetail_paramList.size();sqlDetaili++){
                	sqlCombine_paramList.add(sqlDetail_paramList.get(sqlDetaili));
                }
            	//add sqlSub參數
            	for(int sqlSubi=0;sqlSubi<sqlSub_paramList.size();sqlSubi++){
                	sqlCombine_paramList.add(sqlSub_paramList.get(sqlSubi));
                }
            }
                    
            //end組合sql------------------------------------
			//建表開始--------------------------------------
			//總表
			if(rptStyle.equals("0")){
			   /* 
     		   System.out.println("總表sql="+sqlCombine);
     		   printLog(logps,"總表sql="+sqlCombine);
	 		   dbData_All = DBManager.QueryDB_SQLParam(sqlCombine.toString(),sqlCombine_paramList,"hsien_name,hsien_english,count_seq,field_debit,"+
		                                  "field_credit,field_over,field_over_rate,field_320300,field_net,field_backup,field_captial_rate,field_840740,field_840740_rate,field_backup_100_rate,field_840740_100_rate");

	 		   System.out.print("總表資料 共"+dbData_All.size()+"筆");
	
	 		   HSSFFont ft = wb.createFont();
     		   HSSFCellStyle cs = wb.createCellStyle();
	 		   ft.setFontHeightInPoints((short)18);
	 		   ft.setFontName("標楷體");
	 		   cs.setFont(ft);
	 		   cs.setAlignment(HSSFCellStyle.ALIGN_CENTER);

	   	   	   row = sheet.getRow(0);
                        
	   	   	   for(int v=0;v<21;v++){
	   	   	   	row.createCell((short)v);
	   	   	   }
                        
	   	   	   cell = row.getCell( (short) 0);
	   	   	   cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	   	   	   cell.setCellValue(""+s_year+"年"+s_month+"月"+ bank_type_name +"信用部主要經營指標總表");
	   	   	   cell.setCellStyle(cs);
                 
	   	   	   row = sheet.getRow(1);
	   	       cell = row.getCell( (short)10);
               cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	   	   	   cell.setCellValue(getMonthName(Integer.parseInt(s_month))+String.valueOf(Integer.parseInt(s_year+1911)));
	   	   	   	   	   	   
	   	   	   row = sheet.getRow(2);
                        
	   	   	   cell = row.getCell( (short)14);
	   	   	   cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	   	   	   cell.setCellValue("單位:新臺幣 " + unit_name + ",%");
                        
               rowNum=3;
               DataObject bean = null;
			   String insertValue = "";
			   double sampleValue;
       		   for(int i=0;i<dbData_All.size();i++){
	   		 	 	bean = (DataObject)dbData_All.get(i);	
	   		 	    getBeanData(bean,showEng);//取得各欄位內容
	   		 	 	for(int cellcount=0;cellcount<19;cellcount++){
 		       			row = sheet.createRow(rowNum+i);
      		   			cell = row.createCell( (short)cellcount);
     		   			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			   			insertValue = "";
			   			insertValue = setInsertValue(rptStyle,cellcount);
						HSSFFont f = wb.createFont();
    					HSSFCellStyle cs2 = wb.createCellStyle();
    					f.setFontHeightInPoints((short)10);
    					cs2.setFont(f);
				
						insertValue = (insertValue.equals("null"))?"":insertValue;
						
						cell.setCellValue(insertValue);	
    			
    			        if(cellcount == 0 ){
   				           //96.04.23農會總表多一個台灣省.調整福建省位置 by 2295
    			       	if(i==0){
    			       		cs2.setAlignment(HSSFCellStyle.ALIGN_LEFT);
    			       	}else if(Integer.parseInt(s_year) <= 99 && (i==1 || i==2 || i==3  || i==26) && (bank_type.equals("6") || bank_type.equals("ALL"))){//99.09.23 fix 99年度以前.農漁會
    			       		cs2.setAlignment(HSSFCellStyle.ALIGN_CENTER);  	
    			       	}else if(Integer.parseInt(s_year) >= 100 && (i==1 || i==2 || i==3  || i==4 || i==5 || i==6 || i==23) && (bank_type.equals("6") || bank_type.equals("ALL"))){//99.09.23 fix100年度以後.農漁會
    			       		cs2.setAlignment(HSSFCellStyle.ALIGN_CENTER);  	
    			       	}else if(Integer.parseInt(s_year) <= 99 && (i==1 || i==2 || i==18) && bank_type.equals("7")){//99.09.23 fix 99年度以前.漁會
    			       		cs2.setAlignment(HSSFCellStyle.ALIGN_CENTER);  	
    			    	}else if(Integer.parseInt(s_year) >= 100 && (i==1 || i==2 || i==3 || i==4 || i==16) && bank_type.equals("7")){//99.09.23 fix 100年度以後.漁會
    			       		cs2.setAlignment(HSSFCellStyle.ALIGN_CENTER);  		
    			       	}else{
    			       		cs2.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
    			       	}
    			       }else{
	 			          cs2.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
	 			       }
	 			       cs2.setBorderBottom(HSSFCellStyle.BORDER_THIN);
                       cs2.setBorderLeft(HSSFCellStyle.BORDER_THIN);
                       cs2.setBorderRight(HSSFCellStyle.BORDER_THIN);
	 			       cell.setCellStyle(cs2);
					}//end of cellcount
			   }//end of dbData_All
		
	    	   rowNum = rowNum + dbData_All.size()+1;
	 
	 		   HSSFFont ft4 = wb.createFont();
     		   HSSFCellStyle cs4 = wb.createCellStyle();
	 		   ft4.setFontHeightInPoints((short)12);
	 		   ft4.setFontName("標楷體");
	 		   cs4.setFont(ft4);
	 		   cs4.setAlignment(HSSFCellStyle.ALIGN_LEFT);
	    	   row = sheet.createRow(rowNum);
      	       cell = row.createCell( (short)0);
    	       cell.setEncoding(HSSFCell.ENCODING_UTF_16);
    	       cell.setCellValue("1.資料來源:依各農(漁)會信用部由自有電腦設備或委由相關資訊中心以網際網路傳送之資料彙編");
			   cell.setCellStyle(cs4);
			   rowNum++;
    	       row = sheet.createRow(rowNum);
      	       cell = row.createCell( (short)0);
    	       cell.setEncoding(HSSFCell.ENCODING_UTF_16);
    	       cell.setCellValue("2.備抵呆帳金額(農會)部份:係放款及催收之備抵呆帳合計數,(漁會)部份係放款及透支與催收之備抵呆帳合計數");
			   cell.setCellStyle(cs4);
			   */
			}else{//明細表
			    HSSFCellStyle cellStyle = wb.createCellStyle();    
			      //明細內容用
			      cellStyle = HssfStyle.setStyle( cellStyle, wb.createFont(),
			                                   new String[] {
			                                   "BORDER", "PHR", "PVC", "F09",
			                                   "WRAP"} );
			      //單位代號.機構名稱用
			      HSSFCellStyle cellStyle_name = wb.createCellStyle();    
			      cellStyle_name = HssfStyle.setStyle( cellStyle_name, wb.createFont(),
			                                   new String[] {
			                                   "BORDER", "PHL", "PVC", "F08",
			                                   "WRAP"} );
			   dbData_Part= DBManager.QueryDB_SQLParam(sqlCombine.toString(),sqlCombine_paramList,"hsien_name,hsien_english,bank_no,bank_name,bank_english,count_seq,field_seq,field_debit,"+
		                                  "field_credit,field_over,field_over_rate,field_320300,field_net,field_backup,field_captial_rate,field_840740,field_840740_rate,field_backup_100_rate,field_840740_100_rate");
			   System.out.println("明細表資料 共"+dbData_Part.size()+"筆");
			   printLog(logps,"明細表sql="+sqlCombine);
			   HSSFFont ft = wb.createFont();
    		   HSSFCellStyle cs = wb.createCellStyle();
	 		   ft.setFontHeightInPoints((short)18);
	 		   ft.setFontName("標楷體");
	 		   cs.setFont(ft);
	 		   cs.setAlignment(HSSFCellStyle.ALIGN_CENTER);	 		  
	 		    
	 		   row = sheet.getRow((short)0);    
	           cell = row.getCell( (short) 0);
	           cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	           cell.setCellValue(""+s_year+"年"+s_month+"月"+ bank_type_name +"信用部經營概況統計表");
	           cell.setCellStyle(cs);
               
	           row = sheet.getRow((short)1);    
	           cell = row.getCell( (short) 10);
               cell.setEncoding(HSSFCell.ENCODING_UTF_16);
               
               cell.setCellValue(getMonthName(Integer.parseInt(s_month))+String.valueOf(Integer.parseInt(s_year)+1911));
               
	           row = sheet.getRow(2);
	           cell = row.getCell( (short)9);
	           cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	           cell.setCellValue("單位:新臺幣 " + unit_name + ",%\n"+"Unit：NT$"+Utility.setCommaFormat(unit)+" percent" );
	          
               rowNum=4;
               DataObject bean = null;
               //共同資料的格式
               HSSFFont ft2 = wb.createFont();
               HSSFCellStyle cs3 = wb.createCellStyle();
               ft2.setFontHeightInPoints((short)9);
	           cs3.setFont(ft2);
	           
	           //設定給各地方農漁會信用部name部分 左靠
	           HSSFFont fl = wb.createFont();
               HSSFCellStyle cl = wb.createCellStyle();
               fl.setFontHeightInPoints((short)9);
	           cl.setFont(fl);
	           
	           cl.setAlignment(HSSFCellStyle.ALIGN_LEFT);
	           cl.setBorderTop(HSSFCellStyle.BORDER_THIN);
	           cl.setBorderBottom(HSSFCellStyle.BORDER_THIN);
               cl.setBorderLeft(HSSFCellStyle.BORDER_THIN);
               cl.setBorderRight(HSSFCellStyle.BORDER_THIN);
               String insertValue = ""; 
               double sampleValue;           
       		   for(int i=1;i<dbData_Part.size();i++){
	   		       bean = (DataObject)dbData_Part.get(i);
	   		       getBeanData(bean,showEng);	   		           
	               row = sheet.createRow(rowNum+i);
				   for(int cellcount=0;cellcount<14;cellcount++){ 		         
      		   		   cell = row.createCell( (short)cellcount);
				       cell = row.getCell( (short)cellcount);
     		           cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			           cs3.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
			           insertValue = "";
			           insertValue = setInsertValue(rptStyle,cellcount);
			           //System.out.println("insertValue="+insertValue);
     		           if( cellcount==0 && field_seq.equals("A01") )insertValue =bank_no;
     		           else if( cellcount==1 && field_seq.equals("A01") )insertValue =bank_name;	
     		           else if( cellcount==1 && ( field_seq.equals("A90") ||field_seq.equals("A99") ) ) insertValue =hsien_name;
     		           
			           insertValue = (insertValue.equals("null"))?"":insertValue;
				       cell.setCellValue(insertValue);
                       
    			       if(cellcount == 0 || cellcount ==1 )
    			       cs3.setAlignment(HSSFCellStyle.ALIGN_CENTER);
    			       else
	 			       cs3.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
				       
				       cs3.setBorderTop(HSSFCellStyle.BORDER_THIN);
	 			       cs3.setBorderBottom(HSSFCellStyle.BORDER_THIN);
                       cs3.setBorderLeft(HSSFCellStyle.BORDER_THIN);
                       cs3.setBorderRight(HSSFCellStyle.BORDER_THIN);
                       
				       if(cellcount==1){
				           //cell.setCellStyle(cl);
				           cell.setCellStyle(cellStyle_name);//2012.10.17 add 機構.縣市英文使用
				       } else cell.setCellStyle(cs3);
				       
			       }//end of cellcount
				   if( field_seq.equals("A90") ||field_seq.equals("A99") ){
				  	  rowNum++;row = sheet.createRow(rowNum+i); 		
			  		  for(int cellcount=0;cellcount<14;cellcount++){
			          	 cell = row.createCell( (short)cellcount);
			          	 cell.setCellValue(""); 
			          	 cell.setCellStyle(cs3);
			          }
			          //小計和總計後要空一行
			       }
			   }//end of dbData_Part
			   //總計再取出來一次，放在最底下
	    	   bean = (DataObject)dbData_Part.get(0);
	    	   getBeanData(bean,showEng);
	    	   
			   for(int cellcount=0;cellcount<14;cellcount++){
 		           row = sheet.createRow(dbData_Part.size()+rowNum);
      		       cell = row.createCell( (short)cellcount);
     		       cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			       cs3.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
			       cs3.setWrapText(true);//設定文字內容可跳行
			       insertValue = "";
			       insertValue = setInsertValue(rptStyle,cellcount);			      
     		       if( cellcount==1 && field_seq.equals("A99") ){
     		          if("true".equals(showEng)){//需顯示機構英文名稱
     		              insertValue ="總計\nTotal";     		   
     		          }
     		       }
     		       
			       insertValue = (insertValue.equals("null"))?"0":insertValue;
                   cell.setCellValue(insertValue);
	 		       cell.setCellStyle(cs3);
			   }//end of cellcount
			   //畫最下面的小計與總計的部分
			   rowNum = 3+rowNum+dbData_Part.size();
			   //台北市.高雄市==================================================================
			   for(int i=0;i<dbData_Part.size();i++){
	   			   bean = (DataObject)dbData_Part.get(i);
	   			   field_seq = String.valueOf(bean.getValue("field_seq"));
	   			   hsien_name = String.valueOf(bean.getValue("hsien_name"));
	   			   //System.out.println("hsien_name="+hsien_name);	
	   			   if(Integer.parseInt(s_year) <= 99){	
	   			   	 if( (field_seq.equals("A90") || field_seq.equals("A99")) &&
	   				     (hsien_name.equals("台北市") || hsien_name.equals("高雄市") ) )
	   			     { 		
	   			   	 	  getBeanData(bean,showEng);	   				       
	   			   	 	  row = sheet.createRow(rowNum++);
	   			   	 	  for(int cellcount=0;cellcount<14;cellcount++){
	   			   	 	  	  cell = row.createCell( (short)cellcount);
	   			   	 	  	  cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	   			   	 	  	  insertValue = "";
	   			   	 	  	  insertValue = setInsertValue(rptStyle,cellcount);
	   			   	 	  	  insertValue = (insertValue.equals("null"))?"":insertValue;
	   			   	 	  	  cell.setCellValue(insertValue);
	   			   	 	  	  cell.setCellStyle(cs3);
	   			   	 	  }//end of cellcount
	   			     }//end of 台北市.高雄市
	   			   }else if(Integer.parseInt(s_year) >= 100){//99.09.23 100年(含)以後.明細表加印新北市/台中市/台南市
	   			    if( field_seq.equals("A90") || field_seq.equals("A99")){
	   			    	System.out.println("hsien_name="+hsien_name);	
	   			    }
	   			   	
	   			      if( (field_seq.equals("A90") || field_seq.equals("A99")) &&
	   				      (hsien_name.equals("新北市") || hsien_name.equals("臺北市") 
	   				    || hsien_name.equals("臺中市") || hsien_name.equals("臺南市") 
					    || hsien_name.equals("高雄市") || hsien_name.equals("桃園市") ) ) //103.12.23 add 桃園市
	   			      { 
	   			         getBeanData(bean,showEng);	   				       
					     row = sheet.createRow(rowNum++);
					     for(int cellcount=0;cellcount<14;cellcount++){
		                     cell = row.createCell( (short)cellcount);
		                     cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		                     insertValue = "";
		                     insertValue = setInsertValue(rptStyle,cellcount);  
		                     insertValue = (insertValue.equals("null"))?"":insertValue;
		                     cell.setCellValue(insertValue);
		                     cell.setCellStyle(cs3);
					     }//end of cellcount
	   			      }
	   			   }//end of 新北市.台北市.台中市.台南市.高雄市
			   }//end of dbData_Part
 		       /*106.03.07 原福建省合併至台灣省
 			   dbData_Fukien= DBManager.QueryDB_SQLParam(sqlFukien.toString(),sqlFukien_paramList,"hsien_name,hsien_english,bank_no,bank_name,bank_english,count_seq,field_seq,field_debit,"+
                    "field_credit,field_over,field_over_rate,field_320300,field_net,field_backup,field_captial_rate,field_840740,field_840740_rate,field_backup_100_rate,field_840740_100_rate");
 			   printLog(logps,"福建省sql="+sqlFukien);
 			   //福建省============================================================================
 			   if(bank_type.equals("6")){//101.12.21 漁會不顯示福建省 by 2295
                  for(int i=0;i<dbData_Fukien.size();i++){
               		  bean = (DataObject)dbData_Fukien.get(i);
               		  field_seq = String.valueOf(bean.getValue("field_seq"));
               		  hsien_name = String.valueOf(bean.getValue("hsien_name"));	
               		  if(hsien_name.equals("福建省")  ){                		    
               		      getBeanData(bean,showEng);	
               		      if("true".equals(showEng)){//需顯示機構英文名稱
               		          hsien_english = String.valueOf(bean.getValue("english"));
               		          System.out.println("hsien_name="+hsien_name+":hsien_english="+hsien_english); 
               		          hsien_name += "\n" + hsien_english;
               		      }
               		      row = sheet.createRow(rowNum++);
               		      for(int cellcount=0;cellcount<14;cellcount++){
               		          cell = row.createCell( (short)cellcount);
               		          cell.setEncoding(HSSFCell.ENCODING_UTF_16);
               		          insertValue = "";
               		          insertValue = setInsertValue(rptStyle,cellcount); 
               		          insertValue = (insertValue.equals("null"))?"":insertValue;
               		          cell.setCellValue(insertValue);
               		          cell.setCellStyle(cs3);
               		      }//end of cellcount
               		  }//end of 福建省
                  }//end of dbData_Fukien	
 			   }	
 			   */
			   dbData_Taiwan= DBManager.QueryDB_SQLParam(sqlTaiwan.toString(),sqlTaiwan_paramList,"hsien_name,hsien_english,bank_no,bank_name,bank_english,count_seq,field_seq,field_debit,"+
		                                  "field_credit,field_over,field_over_rate,field_320300,field_net,field_backup,field_captial_rate,field_840740,field_840740_rate,field_backup_100_rate,field_840740_100_rate");
			   printLog(logps,"臺灣省sql="+dbData_Taiwan);
			   //臺灣省============================================================================
			   for(int i=0;i<dbData_Taiwan.size();i++){
	   			  bean = (DataObject)dbData_Taiwan.get(i);
	   			  field_seq = String.valueOf(bean.getValue("field_seq"));
	   			  hsien_name = String.valueOf(bean.getValue("hsien_name"));	
	   			 	   			 
	   			  //System.out.println("hsien_name="+hsien_name);	
	   			  if( field_seq.equals("A92") && hsien_name.equals("其他")  )//106.03.07 原台灣省改為其他(含台灣省及福建省.中華民國農會)
	   			  { 		
	   			      getBeanData(bean,showEng);
	   			      if("true".equals(showEng)){//需顯示機構英文名稱
	   			          hsien_english = String.valueOf(bean.getValue("english"));   			      
	   			          System.out.println("hsien_name="+hsien_name+":hsien_english="+hsien_english); 
	   			          hsien_name += "\n" + hsien_english;
	   			      }
					  row = sheet.createRow(rowNum++);
					  for(int cellcount=0;cellcount<14;cellcount++){
 		                  cell = row.createCell( (short)cellcount);
     		              cell.setEncoding(HSSFCell.ENCODING_UTF_16);
     		              insertValue = "";
     		              insertValue = setInsertValue(rptStyle,cellcount);
				          insertValue = (insertValue.equals("null"))?"":insertValue;
				          cell.setCellValue(insertValue);
	 			          cell.setCellStyle(cs3);
 		              }//end of cellcount
				  }//end of 臺灣省
			   }//end of dbData_Taiwan
			  			  
			  
			  
			  bean = (DataObject)dbData_Part.get(0);
			  getBeanData(bean,showEng);	
			  if("true".equals(showEng)){//需顯示機構英文名稱
			      bank_english = String.valueOf(bean.getValue("bank_english"));
			      hsien_name += "\n"+ bank_english;
			  }
			  row = sheet.createRow(rowNum++);
			  for(int cellcount=0;cellcount<14;cellcount++){
 		          cell = row.createCell( (short)cellcount);
     		      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
     		      insertValue = "";
     		      insertValue = setInsertValue(rptStyle,cellcount);
				  insertValue = (insertValue.equals("null"))?"":insertValue;
				  cell.setCellValue(insertValue);
	 			  cell.setCellStyle(cs3);
 		      }//end of cellcount
			  /*
			  sheet.setDefaultRowHeight((short)41);
			  
			  //設定列高============================================================            
	          for ( int i = 5; i <= rowNum; i++ ) {     
	                  
	                  //sheet.setDefaultRowHeightInPoints(height)	                         
	          }
	          */
	          //======================================================================================
			  
			  HSSFFont ft4 = wb.createFont();
    		  HSSFCellStyle cs4 = wb.createCellStyle();
	          ft4.setFontHeightInPoints((short)12);
	          ft4.setFontName("標楷體");
	          cs4.setFont(ft4);
	          cs4.setAlignment(HSSFCellStyle.ALIGN_LEFT);
	 
	          rowNum =rowNum+2;
	          row = sheet.createRow(rowNum);
      	      cell = row.createCell( (short)0);
    	      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
    	      				
    	      data_year = s_year;
    	      data_month = String.valueOf(Integer.parseInt(s_month)+1);
    	      if(data_month.equals("13")){
    	      	data_year = String.valueOf(Integer.parseInt(s_year)+1);
    	      	data_month = "1";
    	      }
    	      
    	      cell.setCellValue("1.資料來源:依各農(漁)會信用部"+data_year+"年"+((Integer.parseInt(data_month) < 10)?"0":"")+data_month+"月20日前，由自有電腦設備或委由相關資訊中心以網際網路傳送之資料彙編。");
		      cell.setCellStyle(cs4);
		      rowNum++;
    	      row = sheet.createRow(rowNum);
      	      cell = row.createCell( (short)0);
    	      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
    	      cell.setCellValue("2.備抵呆帳金額(農會)部份:係放款及催收之備抵呆帳合計數,(漁會)部份係放款及透支與催收之備抵呆帳合計數。");
		      cell.setCellStyle(cs4);
		      rowNum++;
		      /*106.08.15 取消顯示
              row = sheet.createRow(rowNum);
              cell = row.createCell( (short)0);
              cell.setEncoding(HSSFCell.ENCODING_UTF_16);
              cell.setCellValue("3.地區別欄之其他(農會)係指原臺灣省農會，該農會於102年5月22日更名為中華民國農會。");
              cell.setCellStyle(cs4);
		      */
            }//end of 明細表
 		    //建表結束--------------------------------------
		    if(febxlsFlag.equals("")){	
		    	HSSFFooter footer = sheet.getFooter();
		    	footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
		    	footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
		    	FileOutputStream fout=new FileOutputStream(reportDir + System.getProperty("file.separator")+openfile);
		    	wb.write(fout);//儲存
		    	fout.close();
		    	System.out.println("儲存完成");
		    }
      }catch(Exception e){
                System.out.println("createRpt Error:"+e+e.getMessage());
      }
      return errMsg;
    }//end of createRpt
    
    //98.07.23 取得各欄位data
    private static void getBeanData(DataObject bean,String showEng){
    	try{
    		hsien_name = bean.getValue("hsien_name") == null?"":String.valueOf(bean.getValue("hsien_name"));//單位名稱
    		hsien_english = bean.getValue("hsien_english") == null?"":String.valueOf(bean.getValue("hsien_english"));//101.08.08 add縣市英文名稱
    		if("true".equals(showEng)){//需顯示機構英文名稱
    		    hsien_name += (hsien_english.equals("")?"":"\n"+ hsien_english);
    		}
            bank_no = bean.getValue("bank_no") == null?"":String.valueOf(bean.getValue("bank_no"));
            bank_name = bean.getValue("bank_name") == null?"":String.valueOf(bean.getValue("bank_name"));
            bank_english = bean.getValue("bank_english") == null?"":String.valueOf(bean.getValue("bank_english"));//101.08.08 add機構英文名稱
            if("true".equals(showEng)){//需顯示機構英文名稱
                bank_name += (bank_english.equals("")?"":"\n"+bank_english);
            }
            count_seq = bean.getValue("count_seq") == null?"":String.valueOf(bean.getValue("count_seq"));
            field_seq = bean.getValue("field_seq") == null?"":String.valueOf(bean.getValue("field_seq"));
            field_debit = bean.getValue("field_debit") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_debit")));//存款總額
            field_credit = bean.getValue("field_credit") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_credit")));//放款總額
            field_over = bean.getValue("field_over") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_over")));//逾放金額
            field_over_rate = bean.getValue("field_over_rate") == null?"":String.valueOf(bean.getValue("field_over_rate"));//逾放比率
            
            field_320300 = bean.getValue("field_320300") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_320300")));//本期損益
            field_net = bean.getValue("field_net") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_net")));//淨值
            field_backup = bean.getValue("field_backup") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_backup")));//備抵呆帳
            field_captial_rate = bean.getValue("field_captial_rate") == null?"":String.valueOf(bean.getValue("field_captial_rate"));//淨值佔風險性資產比率           
 	 	  	field_840740 = bean.getValue("field_840740") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_840740")));//field_840740
 	 	  	field_840740_rate = bean.getValue("field_840740_rate") == null?"":String.valueOf(bean.getValue("field_840740_rate"));//field_840740_rate廣義逾放比
 	 	  	field_backup_100_rate = bean.getValue("field_backup_100_rate") == null?"":String.valueOf(bean.getValue("field_backup_100_rate"));//備呆占狹義逾期放款比率(備抵呆帳/狹義逾放)
 	 	  	field_840740_100_rate = bean.getValue("field_840740_100_rate") == null?"":String.valueOf(bean.getValue("field_840740_100_rate"));//備呆占廣義逾期放款比率(備抵呆帳/廣義逾放)
 	 	    //利率顯示小數點至第2位,不足者補0
            field_over_rate = field_over_rate.equals("N/A") ? "N/A" : Utility.setCommaFormat(df_md.format(Double.parseDouble(field_over_rate)));
            field_captial_rate = field_captial_rate.equals("N/A") ? "N/A" : Utility.setCommaFormat(df_md.format(Double.parseDouble(field_captial_rate)));
            field_840740_rate = field_840740_rate.equals("N/A") ? "N/A" : Utility.setCommaFormat(df_md.format(Double.parseDouble(field_840740_rate)));
            field_backup_100_rate = field_backup_100_rate.equals("N/A") ? "N/A" : Utility.setCommaFormat(df_md.format(Double.parseDouble(field_backup_100_rate)));
            field_840740_100_rate = field_840740_100_rate.equals("N/A") ? "N/A" : Utility.setCommaFormat(df_md.format(Double.parseDouble(field_840740_100_rate)));
            
    	}catch(Exception e){
    		System.out.println("getBeanData Error:"+e+e.getMessage());
    	}
    }
    //98.07.23 add set insertValue
   	private static String setInsertValue(String rptStyle,int cellcount){
    		 String insertValue = "";
    		 if(rptStyle.equals("0")){//總表
    		    if ( cellcount==0 )     insertValue =hsien_name;//單位名稱
                else if( cellcount==1 ) insertValue =field_debit;//存款總額
                else if( cellcount==2 ) insertValue =field_credit;//放款總額
                else if( cellcount==3 )insertValue =field_net;//淨值
                else if( cellcount==4 ) insertValue =field_320300;//本期損益                
                else if( cellcount==5 ) insertValue =field_over;//狹義逾期放款
                else if( cellcount==6 )insertValue =field_840740;//廣義逾期放款
                else if( cellcount==7 ) insertValue =field_over_rate;//狹義逾放比率
                else if( cellcount==8 )insertValue =field_840740_rate;//廣義逾放比率
                else if( cellcount==9 )insertValue =field_backup;//備抵呆帳
                else if( cellcount==10 )insertValue =field_backup_100_rate;//備呆占狹義逾期放款比率(備抵呆帳/狹義逾放)
                else if( cellcount==11 )insertValue =field_840740_100_rate;//備呆占廣義逾期放款比率(備抵呆帳/廣義逾放)
                else if( cellcount==12 )insertValue =field_captial_rate;//淨值佔風險性資產比率
    		 }else{
    		    if( cellcount==1 )insertValue =hsien_name;
    		    else if( cellcount==2 ) insertValue =field_debit;//存款總額
                else if( cellcount==3 ) insertValue =field_credit;//放款總額
                else if( cellcount==4 )insertValue =field_net;//淨值
                else if( cellcount==5 ) insertValue =field_320300;//本期損益                
                else if( cellcount==6 ) insertValue =field_over;//狹義逾期放款
                else if( cellcount==7 )insertValue =field_840740;//廣義逾期放款
                else if( cellcount==8 ) insertValue =field_over_rate;//狹義逾放比率
                else if( cellcount==9 )insertValue =field_840740_rate;//廣義逾放比率
                else if( cellcount==10 )insertValue =field_backup;//備抵呆帳
                else if( cellcount==11 )insertValue =field_backup_100_rate;//備呆占狹義逾期放款比率(備抵呆帳/狹義逾放)
                else if( cellcount==12 )insertValue =field_840740_100_rate;//備呆占廣義逾期放款比率(備抵呆帳/廣義逾放)
                else if( cellcount==13 )insertValue =field_captial_rate;//淨值佔風險性資產比率
    		 }
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
   public static String getMonthName(int month){
        String unit_name="";
        try{
            if (month==1){
                unit_name="Jan-";
            }else if (month==2){                 
                unit_name="Feb-";
            }else if (month==3){
                unit_name="Mar-";
            }else if (month==4){
                unit_name="Apr-";
            }else if (month==5){
                unit_name="May-";
            }else if (month==6){
                unit_name="Jun-";
            }else if (month==7){                 
                unit_name="Jul-";
            }else if (month==8){
                unit_name="Aug-";
            }else if (month==9){
                unit_name="Sep-";
            }else if (month==10){
                unit_name="Oct-";
            }else if (month==11){
                unit_name="Nov-";
            }else if (month==12){
                unit_name="Dec-";
            }
        }catch(Exception e){
            System.out.println("getMonthName Error:"+e.getMessage());
        }
        System.out.println("unit_name="+unit_name);
        return unit_name;   
    }
}

