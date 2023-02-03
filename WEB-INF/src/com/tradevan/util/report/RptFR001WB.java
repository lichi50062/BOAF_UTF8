/*
 * Created on 2005/12/04 by 4180
 * fixed on 2006/01/20 by 4180
 * SQL中的除以金額單位(unit)移除，改於資料送出前處理
 * fixed on 2006/01/24 by 4180
 * 存放比率公式修正
 * 2006.06.14 add 農漁會總表 by 2295
 * 2006.06.21 fix 農漁會總表存放比率sql by 2295
 * 2007.04.24 add 農會總表多一個台灣省.調整福建省位置 by 2295
 * 2007.04.30 存放比率若為負數,以0顯示 by 2295
 * 2007.11.07 fix 淨值佔風險性資產比率總計.縣市別是以910400/910500去計算.不是以91060P/單位數 by 2295
 * 2009.06.16 add 檢查局傳送格式 by 2295
 * 2009.07.23 add A02.990611對直轄市、縣(市)政府、離島地區鄉(鎮、市)公所辦理之授信總額 by 2295
 * 2010.04.12 fix 根據查詢年度.100年以後取得新縣市別.100年以前取得舊縣市別 
 * 				  使用PreparedStatement;並列印轉換後的SQL;套用QueryDB_SQLParam by 2295
 * 2010.09.23 fix 調整100年度.縣市排列順序 by 2295
 *                100年(含)以後.明細表最後加印新北市/台中市/台南市 by 2295
 * 2010.10.26 add 只顯示未被裁撤的機構 by 2295
 * 2011.04.25 fix 因直轄市由台北市/台中市/台南市.更名成臺北市/臺中市/臺南市,修改最底下合計顯示 by 2295
 * 2013.06.07 add 103/01以後.A01漁會.套用新科目代號/計算公式 by 2295  
 * 2013.11.19 add 原QueryDB改套用QueryDB_SQLParam by 2295
 * 2013.12.04 add 上傳至檢查局報表增加A99.992710建築放款餘額 by 2295
 * 2014.01.09 add 累積虧損占信用部上年度淨算淨值之比率 by 2295
 * 2014.01.15 add 政策性農業專案貸款餘額A01.120600 by 2295
 * 2014.01.16 add 臺灣省改其他,並增加說明 by 2295
 * 2014.02.11 fix 檢查局需求.政策性農業專案貸款餘額移至建築放款餘額之後 by 2295
 * 2014.12.23 add 總表調整桃園市位置及明細表增加顯示桃園市於最下方直轄市統計 by 2295
 * 2016.03.16 add 縣市政府列印總表/明細表,只顯示其所屬縣市別或轄區下的農漁會信用部明細資料 by 2295  
 * 2016.11.04 add 明細表.調整原台灣省改為其他(包含台灣省及福建省.中華民國農會) by 2295   
 * 2017.08.04 add 總表.其他縣市.顯示成中華民國農會 by 2295         
 * 2018.08.13 add 調整縣市政府列印總表可看到所有農漁會信用部  by 2295
 * 2019.04.17 add 明細表.其他縣市(增加包含其他(中華民國農會)) by 2295
 */
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;


import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

import javax.servlet.http.HttpSession;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR001WB {
	static String field_seq = "";
	static String hsien_name = "";         
	static String bank_no = "";            
	static String bank_name = "";          
	static String count_seq = "";          
	static String field_debit = "";        
	static String field_credit = "";       
	static String field_dc_rate = "";      
	static String field_120700= "";        
	static String field_over = "";         
	static String field_over_rate = "";    
	static String field_320300 = "";       
	static String field_transfer = "";     
	static String field_transfer_rate = "";
	static String field_310000 = "";       
	static String field_net = "";          
	static String field_fixnet_rate = "";  
	static String field_check_rate = "";   
	static String field_150200 = "";       
	static String field_backup = "";       
	static String field_noassure = "";     
	static String field_modifynet= "";     
	static String field_captial_rate = "";
	static String field_990611 = "";//98.07.22 add 對直轄市、縣(市)政府、離島地區鄉(鎮、市)公所辦理之授信總額
	static String field_992710 = "";//102.12.04 add A99.992710建築放款餘額
	static String field_320100_rate = "";//103.01.09 add 累積虧損占信用部上年度決算淨值之比率 
	static String field_120600 = "";//103.01.14 add 政策性農業專案貸款餘額
	static File logfile;
    static FileOutputStream logos=null;      
    static BufferedOutputStream logbos = null;
    static PrintStream logps = null;
    static Date nowlog = new Date();
    static SimpleDateFormat logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");        
    static SimpleDateFormat logfileformat = new SimpleDateFormat("yyyyMMddHHmmss");
    static Calendar logcalendar;
    static File logDir = null;
    public static String createRpt(String s_year,String s_month,String unit,String bank_type,String rptStyle,String febxlsFlag,HSSFWorkbook wb){
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
          String hsien_id_b="";//縣市政府所屬縣市別 105.03.16 add
          int rowNum=0;
          reportUtil reportUtil = new reportUtil();
          System.out.println("RptFR001WB.bank_type="+bank_type);
          if(bank_type.equals("ALL") || bank_type.startsWith("B")){//95.06.14 add 總表?加農漁會             
             bank_type_name = "農漁會";            
             if(bank_type.startsWith("B")){//105.03.16 add 縣市政府,則為全部轄區下的農漁會信用部              
                 hsien_id_b=bank_type.substring(2,3);
                 bank_type="ALL";
                 System.out.println("FR001WB.hsien_id="+hsien_id_b);
             }            
          }else{
             bank_type_name = (bank_type.equals("6"))?"農會":"漁會";
          }          
          unit_name = Utility.getUnitName(unit);//取得單位名稱
          
      try{
          /*
          logDir  = new File(Utility.getProperties("logDir"));
          if(!logDir.exists()){
              if(!Utility.mkdirs(Utility.getProperties("logDir"))){
                 System.out.println("目錄新增失敗");
              }    
          }
          logfile = new File(logDir + System.getProperty("file.separator") + "RptFR001WB.log");
          
          System.out.println("logfile filename="+logDir + System.getProperty("file.separator") +"RptFR001WB.log");
          logos = new FileOutputStream(logfile,true);                         
          logbos = new BufferedOutputStream(logos);
          logps = new PrintStream(logbos);   
          */
          
          
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
            
            String openfile="全體農漁會信用部主要經營指標";
            
            FileInputStream finput = null;
            if(febxlsFlag.equals("")){//98.06.16 原主要經營指標
               openfile+=(rptStyle.equals("0"))?"總表":"明細表";
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
	        if(rptStyle.equals("0"))       ps.setScale( ( short )60 ); //列印縮放百分比
	        else  ps.setScale( ( short )55 ); //列印縮放百分比
			
	        ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
	  		//wb.setSheetName(0,"test");
	  		//設定表頭 為固定 先設欄的起始再設列的起始
	        wb.setRepeatingRowsAndColumns(0, 1, 21, 2, 3);

	        if(febxlsFlag.equals("")) finput.close();

            HSSFRow row=null;//宣告一列
	  		HSSFCell cell=null;//宣告一個儲存格

		    //共同sql開頭
	  		sqlSubhead.append( " select a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order, ");
	  		sqlSubhead.append( "         a01.bank_no ,   a01.BANK_NAME,  ");
	  		sqlSubhead.append( "         COUNT_SEQ, field_SEQ,   ");
	  		sqlSubhead.append( "	 	  round(field_DEBIT /?,0)  as field_DEBIT,   ");    //2006/01/20 fixed by 4180 除以單位
	  		sqlSubhead.append( " 		  round(field_CREDIT /?,0)  as field_CREDIT, ");
	  		sqlSubhead.append( "    	  decode(a01.fieldI_Y,0,0,  ");
	  		sqlSubhead.append( "    round( ");
	  		sqlSubhead.append( "          (a01.fieldI_XA                                          + ");
	  		sqlSubhead.append( "             decode(sign(a01.fieldI_XB1 - a01.fieldI_XB2),-1,0,     ");
	  		sqlSubhead.append( "                      (a01.fieldI_XB1 - a01.fieldI_XB2))          + ");
	  		sqlSubhead.append( "           decode(sign(a01.fieldI_XC1 - a01.fieldI_XC2),-1,0,       ");
	  		sqlSubhead.append( "                      (a01.fieldI_XC1 - a01.fieldI_XC2))          + ");
	  		sqlSubhead.append( "          decode(sign(a01.fieldI_XD1 - a01.fieldI_XD2),-1,0,        ");
	  		sqlSubhead.append( "                     (a01.fieldI_XD1 - a01.fieldI_XD2))           + ");
	  		sqlSubhead.append( "          decode(sign(a01.fieldI_XE1 - a01.fieldI_XE2),-1,0,        ");
	  		sqlSubhead.append( "                     (a01.fieldI_XE1 - a01.fieldI_XE2))           - ");
	  		sqlSubhead.append( "          decode(sign(a01.fieldI_XF1 - a01.fieldI_XF3 -  a01.fieldI_XF2),-1,0,"); //2006/01/27 公式修改後新增field_XF3
	  		sqlSubhead.append( "                     (a01.fieldI_XF1 - a01.fieldI_XF3 - a01.fieldI_XF2)) "); //2006/01/27 公式修改後新增field_XF3
	  		sqlSubhead.append( "          )  ");
	  		sqlSubhead.append( "         /    a01.fieldI_Y * 100,2))        as     field_DC_RATE , ");
	  		sqlSubhead.append( "	round(field_120700 /?,0)  as field_120700, ");   //2006/01/20 fixed by 4180 除以單位
	  		sqlSubhead.append( "	round(field_OVER /?,0)  as field_OVER, ");
	  		sqlSubhead.append( "    decode(a01.field_CREDIT,0,0,round(a01.field_OVER /  a01.field_CREDIT *100 ,2))  as   field_OVER_RATE,");
	  		sqlSubhead.append( "	round(field_320300 /?,0)  as field_320300, "); //2006/01/25 fixed by 4180 除以單位
	  		sqlSubhead.append( "	round(field_TRANSFER /?,0)  as field_TRANSFER, "); //2006/01/25 fixed by 4180 除以單位
	  		sqlSubhead.append( "  	decode(field_DEPOSITBANK,0,0,round(a01.field_DEPOSITBANK_AA / a01.field_DEPOSITBANK *100 ,2))  as   field_TRANSFER_RATE, ");
	  		sqlSubhead.append( "	round(field_310000 /?,0)  as field_310000, ");         //2006/01/20 fixed by 4180 除以單位
	  		sqlSubhead.append( "	round(field_NET /?,0)  as field_NET, ");
	  		sqlSubhead.append( "    decode(field_NET,0,0,round(a01.field_140000 /  a01.field_NET *100 ,2))  as   field_FIXNET_RATE,  ");
	  		sqlSubhead.append( "    decode(field_DEBIT,0,0,round(a01.field_CHECK /  a01.field_DEBIT *100 ,2))  as   field_CHECK_RATE, ");
	  		sqlSubhead.append( "	round(field_150200 /?,0)  as field_150200, ");
	  		sqlSubhead.append( "	round(field_BACKUP /?,0)  as field_BACKUP, ");
	  		sqlSubhead.append( "	round(field_NOASSURE /?,0)  as field_NOASSURE, ");
	  		sqlSubhead.append( "    round(field_990611 /?,0)  as field_990611, ");//98.07.22 add
	  		sqlSubhead.append( "    round(field_992710 /?,0)  as field_992710, ");//102.12.04 add 檢查局用.建築放款餘額
	  		sqlSubhead.append( "    decode(field_990230-field_990240-field_992810,0,0,round(field_320100 / (field_990230-field_990240-field_992810) *100 ,2)) as field_320100_rate, ");//103.01.09 add 累積虧損占信用部上年度決算淨值之比率
	  		sqlSubhead.append( "    round(field_120600 /?,0)  as field_120600, ");//103.01.14 add A01.120600政策性農業專案貸款餘額
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
	  		sqlSubhead_paramList.add(unit);
	  		sqlSubhead_paramList.add(unit);
	  		sqlSubhead_paramList.add(unit);
	  		sqlSubhead_paramList.add(unit);
            //各縣市、臺灣省、福建省、總計才加
            //2007.11.07 fix 淨值佔風險性資產比率總計.縣市別是以SUM(910400)/SUM(910500)去計算.不是以91060P/單位數 by 2295
            sqlSubdiv1.append( " round(a01.field_MODIFYNET / COUNT_SEQ,0)  as   field_MODIFYNET,");
            //96.11.05" round(field_CAPTIAL / COUNT_SEQ /  1000  ,2)  as   field_CAPTIAL_RATE         ";
            sqlSubdiv1.append( " round(field_CAPTIAL / 1000  ,2)  as   field_CAPTIAL_RATE ");
            		   //" round(field_990611 /"+unit+",0)  as field_990611 ";//98.07.22 add
            //明細才加
            sqlSubdiv2.append( " round(field_MODIFYNET /"+unit+" ,0)  as field_MODIFYNET,  ");  //2006/01/20 fixed by 4180 除以單位			
            sqlSubdiv2.append( " round(field_CAPTIAL /  1000 ,2)  as   field_CAPTIAL_RATE  ");
               		    //" round(field_990611 /"+unit+",0)  as field_990611 ";//98.07.22 add
            //各縣市小計用 
            sqlSubdiv3.append( " from ( select a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order,");
            sqlSubdiv3.append( "               ' ' AS  bank_no ,     ' '   AS  BANK_NAME, ");
            sqlSubdiv3.append( "               COUNT(*)  AS  COUNT_SEQ, ");
            sqlSubdiv3.append( "		       'A90'  as  field_SEQ,  ");
           
            //總計用             
            sqlSubdiv4.append( " from ( select ' '  AS  hsien_id ,  ' 總   計 '   AS hsien_name,  '001'  AS FR001W_output_order, ");
            sqlSubdiv4.append( "               ' ' AS  bank_no ,     ' '   AS  BANK_NAME, ");
            sqlSubdiv4.append( "               COUNT(*)  AS  COUNT_SEQ, ");
            sqlSubdiv4.append( "               'A99'  as  field_SEQ,    ");
           
            //臺灣省用
            if(rptStyle.equals("1")){//明細表.105.11.04 add 明細表的臺灣省改為其他縣市(包含台灣省及福建省)            
                sqlSubdiv5.append( " from ( select ' '  AS  hsien_id ,  '其他'   AS hsien_name,  '025'  AS FR001W_output_order, ");
            }else{
                sqlSubdiv5.append( " from ( select ' '  AS  hsien_id ,  '臺灣省'   AS hsien_name,  '025'  AS FR001W_output_order, ");
            }
            //sqlSubdiv5.append( " from ( select ' '  AS  hsien_id ,  '臺灣省'   AS hsien_name,  '025'  AS FR001W_output_order, ");
            sqlSubdiv5.append( "               ' ' AS  bank_no ,     ' '   AS  BANK_NAME, ");
            sqlSubdiv5.append( "               COUNT(*)  AS  COUNT_SEQ, ");
            sqlSubdiv5.append( " 		        'A92'  as  field_SEQ,   ");
           
            //福建省用 
            sqlSubdiv6.append( "  from( select ' '  AS  hsien_id ,  '福建省'   AS hsien_name,  '235'  AS FR001W_output_order,  ");
            sqlSubdiv6.append( " 		       ' ' AS  bank_no ,     ' '   AS  BANK_NAME, ");
            sqlSubdiv6.append( "  		       COUNT(*)  AS  COUNT_SEQ, "); 
            sqlSubdiv6.append( "  		       'A93'  as  field_SEQ,    ");
           
            //明細用
            sqlSubdiv7.append( " from ( select a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order, ");
            sqlSubdiv7.append( "      	       a01.bank_no ,   a01.BANK_NAME,");
            sqlSubdiv7.append( "               1  AS  COUNT_SEQ,      ");
            sqlSubdiv7.append( " 		       'A01'  as  field_SEQ,   ");			
           
            //SUM共同區段
            sqlDivSum.append( " SUM(field_120600)  field_120600 ,");//103.01.14 add
            sqlDivSum.append( " SUM(field_120700)  field_120700 ,");
            sqlDivSum.append( " SUM(field_320100)  field_320100 ,");//103.01.09 add
            sqlDivSum.append( " SUM(field_320300)  field_320300 , ");
            sqlDivSum.append( " SUM(field_TRANSFER) field_TRANSFER ,");
            sqlDivSum.append( " SUM(field_DEPOSITBANK_AA) field_DEPOSITBANK_AA,");
            sqlDivSum.append( " SUM(field_DEPOSITBANK)    field_DEPOSITBANK,");
            sqlDivSum.append( " SUM(field_310000)  field_310000, ");
            sqlDivSum.append( " SUM(field_NET)     field_NET,    ");
            sqlDivSum.append( " SUM(field_140000)  field_140000, ");
            sqlDivSum.append( " SUM(field_CHECK)   field_CHECK,  ");
            sqlDivSum.append( " SUM(field_150200)  field_150200, ");
            sqlDivSum.append( " SUM(field_BACKUP)    field_BACKUP,  ");
            sqlDivSum.append( " SUM(field_NOASSURE)  field_NOASSURE, ");
            sqlDivSum.append( " SUM(field_MODIFYNET) field_MODIFYNET, ");
            sqlDivSum.append( " SUM(field_OVER)      field_OVER, ");
            sqlDivSum.append( " SUM(field_DEBIT)     field_DEBIT, ");
            sqlDivSum.append( " SUM(field_CREDIT)    field_CREDIT, ");
            sqlDivSum.append( " SUM(fieldI_XA)       fieldI_XA, ");
            sqlDivSum.append( " SUM(fieldI_XB1)      fieldI_XB1, ");
            sqlDivSum.append( " SUM(fieldI_XB2)      fieldI_XB2, ");
            sqlDivSum.append( " SUM(fieldI_XC1)      fieldI_XC1, ");
            sqlDivSum.append( " SUM(fieldI_XC2)      fieldI_XC2, ");
            sqlDivSum.append( " SUM(fieldI_XD1)      fieldI_XD1, ");
            sqlDivSum.append( " SUM(fieldI_XD2)      fieldI_XD2, ");
            sqlDivSum.append( " SUM(fieldI_XE1)      fieldI_XE1, ");
            sqlDivSum.append( " SUM(fieldI_XE2)      fieldI_XE2, ");
            sqlDivSum.append( " SUM(fieldI_XF1)      fieldI_XF1, ");
            sqlDivSum.append( " SUM(fieldI_XF3)      fieldI_XF3, ");  //2006/01/27 公式修改後新增field_XF3
            sqlDivSum.append( " SUM(fieldI_XF2)      fieldI_XF2, ");
            sqlDivSum.append( " SUM(fieldI_Y)        fieldI_Y,   ");                        
                        //" SUM(nvl(a05.amt,0))       field_CAPTIAL,       ");
            sqlDivSum.append( " SUM(a02.field_990611) field_990611,");//103.01.09
            sqlDivSum.append( " SUM(a02.field_990230) field_990230,");//103.01.09
            sqlDivSum.append( " SUM(a02.field_990240) field_990240,");//103.01.09
            sqlDivSum.append( " SUM(a99.field_992710) field_992710,");//103.01.09
            sqlDivSum.append( " SUM(a99.field_992810) field_992810,");//103.01.09
            //共有sql
            sqlDiv.append( " from ( select nvl(cd01.hsien_id,' ')       as  hsien_id ,");
            sqlDiv.append( "               nvl(cd01.hsien_name,'OTHER') as  hsien_name,");
            sqlDiv.append( "               cd01.FR001W_output_order     as  FR001W_output_order,");
            sqlDiv.append( "               bn01.bank_no ,  bn01.BANK_NAME,");
            sqlDiv.append( "               round(sum(decode(a01.acc_code,'120600',amt,0)) /1,0) as field_120600,");//103.01.14 add
            sqlDiv.append( "               round(sum(decode(a01.acc_code,'120700',amt,0)) /1,0) as field_120700,");
            sqlDiv.append( "               round(sum(decode(a01.acc_code,'320100',amt,0)) /1,0) as field_320100,");//103.01.09 add
            sqlDiv.append( "               round(sum(decode(a01.acc_code,'320300',amt,0)) /1,0) as field_320300,");
            sqlDiv.append( "               round(sum( decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'110324',amt,'110325',amt,0),'7',decode(a01.acc_code,'110254',amt,'110255',amt,0)),");
            sqlDiv.append( "                                           '103',decode(a01.acc_code,'110324',amt,'110325',amt,0),0) ");
            sqlDiv.append( "                     ) /1,0)     as field_TRANSFER, "); 
            sqlDiv.append( "               round(sum( decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'110320',amt,0),'7',decode(a01.acc_code,'110250',amt,0)),");
            sqlDiv.append( "                                           '103',decode(a01.acc_code,'110320',amt,0),0)");
            sqlDiv.append( "                     ) /1,0)     as field_DEPOSITBANK_AA, "); 
            sqlDiv.append( "               round(sum( decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'110300',amt,0),'7',decode(a01.acc_code,'110200',amt,0)),");
            sqlDiv.append( "                                           '103',decode(a01.acc_code,'110300',amt,0),0) ");
            sqlDiv.append( "                    ) /1,0)     as field_DEPOSITBANK, "); 
            sqlDiv.append( " round(sum(decode(a01.acc_code,'310000',amt,0)) /1,0)     as field_310000, ");
			sqlDiv.append( " round(sum(decode(bank_type,'6',decode(a01.acc_code,'310000',amt,'320000',amt,0),'7',decode(a01.acc_code,'300000',amt,0),0)) /1,0)     as field_NET,  ");	   
		    sqlDiv.append( " round(sum(decode(a01.acc_code,'140000',amt,0)) /1,0) as field_140000, ");
		    sqlDiv.append( " round(sum(decode(a01.acc_code,'220100',amt, '220200',amt, '220300',amt, '220400',amt, '220500',amt,0)) /1,0)  as field_CHECK, ");
		    sqlDiv.append( " round(sum(decode(a01.acc_code,'150200',amt,0)) /1,0) as field_150200, ");
            sqlDiv.append( " round(sum(decode(a01.acc_code, '120800',amt,'150300',amt,0)) /1,0) as  field_BACKUP, ");
            sqlDiv.append( " round(sum( decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code, '120101',amt,'120301',amt, '120401',amt, '120501',amt,0),'7',decode(a01.acc_code, '120101',amt,'120401', amt, '120201',amt, '120501',amt,0)), ");
            sqlDiv.append( "                             '103',decode(a01.acc_code, '120101',amt,'120301',amt, '120401',amt, '120501',amt,0),0) ");
            sqlDiv.append( " ) /1,0) as  field_NOASSURE, ");  
            sqlDiv.append( " round((sum(decode(bank_type,'6',decode(a01.acc_code, '310000',amt,'320000',amt, '120800',amt,'150300',amt,0),'7',decode(a01.acc_code, '300000',amt, '120800',amt, '150300',amt,0))) -  "); 
            sqlDiv.append( " round(sum(decode(a01.acc_code, '990000',amt,0)) * 1.25 * 0.7,0))/1,0) as  field_MODIFYNET, ");  
            sqlDiv.append( " round(sum(decode(a01.acc_code,'990000',amt,0)) /1,0) as field_OVER,  ");
            sqlDiv.append( " round(sum(decode(a01.acc_code,'220000',amt,0)) /1,0) as field_DEBIT, ");
            sqlDiv.append( " round(sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /1,0) as  field_CREDIT, ");
            sqlDiv.append( " round(sum(decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'120101',amt,'120102',amt, "); 
            sqlDiv.append( "                                                      '120200',amt,'120301',amt,'120302',amt,'120700',amt,'150200',amt,0) "); 
            sqlDiv.append( "                                                   ,'7',decode(a01.acc_code,'120101',amt,'120102',amt, "); 
            sqlDiv.append( "                                                      '120300',amt,'120401',amt,'120402',amt,'120700',amt,'150200',amt,0)), ");
            sqlDiv.append( "                            '103',decode(a01.acc_code,'120101',amt,'120102',amt,'120200',amt,'120301',amt,");  
            sqlDiv.append( "                                 '120302',amt,'120700',amt,'150200',amt,0),0)");
            sqlDiv.append( " ) /1,0)     as fieldI_XA,");  
            sqlDiv.append( " round(sum(decode(year_type,'102',decode(bank_type,'6',decode(a01.acc_code,'120401',amt,'120402',amt,0),'7',decode(a01.acc_code,'120201',amt,'120202',amt,0)),");
            sqlDiv.append( "                            '103',decode(a01.acc_code,'120401',amt,'120402',amt,0),0)");
            sqlDiv.append( " ) /1,0)     as fieldI_XB1,");  
            sqlDiv.append( " sum(decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'240305',amt,'251300',amt,'310800',amt, 0),'7',decode(a01.acc_code,'240205',amt, '310800',amt,0)),");
            sqlDiv.append( "                      '103',decode(bank_type,'6',decode(a01.acc_code,'240305',amt,'251300',amt,'310800',amt, 0),'7',decode(a01.acc_code,'240305',amt, '251200',amt,0)),0)");
            sqlDiv.append( " )  as fieldI_XB2,"); 
            sqlDiv.append( "  round(sum(decode(a01.acc_code,'120501',amt,'120502',amt,0)) /1,0)     as fieldI_XC1,  ");
            sqlDiv.append( "  round(sum( decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'240301',amt,'240302',amt,'240303',amt,'240304',amt,0),");  
            sqlDiv.append( "                                                     '7',decode(a01.acc_code,'240201',amt,'240202',amt,'240203',amt,'240204',amt,0)),");
            sqlDiv.append( "                              '103', decode(a01.acc_code,'240301',amt,'240302',amt,'240303',amt,'240304',amt,0),0)  ");            
            sqlDiv.append( ") /1,0)     as fieldI_XC2, ");
            sqlDiv.append( " round(sum(decode(a01.acc_code,'120600',amt,0)) /1,0)                  as fieldI_XD1,  ");
            sqlDiv.append( " round(sum( decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'240200',amt,0),");
            sqlDiv.append( "                                                    '7',decode(a01.acc_code,'240300',amt,0)),");
            sqlDiv.append( "                             '103',decode(a01.acc_code,'240200',amt,0),0)");                                  
            sqlDiv.append( " ) /1,0)   as fieldI_XD2, ");
            sqlDiv.append( " round(sum(decode(a01.acc_code,'150100',amt,0)) /1,0)                  as fieldI_XE1,  ");
            sqlDiv.append( " round(sum(decode(a01.acc_code,'250100',amt,0)) /1,0)                  as fieldI_XE2,  ");
            sqlDiv.append( " round(sum(decode(a01.acc_code,'310000',amt,'320000',amt,0)) /1,0)     as fieldI_XF1,  ");
            sqlDiv.append( " round(sum( decode(YEAR_TYPE,'102',decode(a01.acc_code,'310800',amt,0),");
            sqlDiv.append( "                             '103',decode(bank_type,'6',decode(a01.acc_code,'310800',amt,0),'7',0,0),0)");//2006/01/27 公式修改後新增field_XF3
            sqlDiv.append( " ) /1,0)  as fieldI_XF3, ");
            sqlDiv.append( " round(sum(decode(a01.acc_code,'140000',amt,0)) /1,0)                  as fieldI_XF2,  ");
            sqlDiv.append( " round((sum(decode(a01.acc_code,'220100',amt,'220200',amt,                             ");
            sqlDiv.append( "                                '220300',amt,'220400',amt,                             ");
            sqlDiv.append( "                                '220500',amt,'220600',amt,                             ");
            sqlDiv.append( "                                '220700',amt,'220800',amt,                             ");
            sqlDiv.append( "                                '220900',amt,'221000',amt,0))-                         ");
            sqlDiv.append( " round(sum(decode(a01.acc_code,'220900',amt,0))/2,0)) /1,0)   as fieldI_Y       ");                       
            sqlDiv.append( " from  (select * from "+cd01_table+" cd01 where cd01.hsien_id <> 'Y' ");           
  
            sqlTaiwanDiv.append(sqlDiv); //臺灣省和福建省共用區段只差中間一小部分，故在這邊assign給sqlTaiwanDiv \uFFFDBsqlFukienDiv
            if(rptStyle.equals("1")){//明細表
                //105.11.04 add 明細表的臺灣省改為其他縣市(包含台灣省及福建省)
                //108.04.17 add 其他縣市(增加包含其他(中華民國農會))
                sqlTaiwanDiv.append(" and cd01.Hsien_div  in ('2','3','4')");//2:台灣省 3:福建省 4:其他(中華民國農會)
            }else{
                sqlTaiwanDiv.append(" and cd01.Hsien_div = '2'");
            }
            sqlFukienDiv.append(sqlDiv);
            sqlFukienDiv.append(" and cd01.Hsien_div = '3'");
                
            //sqlDivtemp設定參數
            sqlDivtemp.append( " ) cd01 left join (select * from wlx01 where m_year="+wlx01_m_year+ " and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL))wlx01 on wlx01.hsien_id=cd01.hsien_id ");
            if(bank_type.equals("ALL")){
               sqlDivtemp.append( "  left join (select * from bn01 where m_year="+wlx01_m_year+" and bn_type <> '2')bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type in ('6','7') ");
            }else{
               sqlDivtemp.append( "  left join (select * from bn01 where m_year="+wlx01_m_year+" and bn_type <> '2')bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=? ");     
               sqlDivtemp_paramList.add(bank_type);
            }
            sqlDivtemp.append( "  left join (select (CASE WHEN (a01.m_year <= 102) THEN '102'");
            sqlDivtemp.append( "   WHEN (a01.m_year > 102) THEN '103'");
            sqlDivtemp.append( "  ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01 where a01.m_year  = ? and a01.m_month  = ?) a01  on  bn01.bank_no = a01.bank_code ");
            sqlDivtemp.append( " group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME ");
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
		    sqlSubtail.append( "  		   nvl(cd01.hsien_name,'OTHER')  as  hsien_name,");  
		    sqlSubtail.append( "		   cd01.FR001W_output_order     as  FR001W_output_order,");  
		    sqlSubtail.append( " 	       bn01.bank_no as bank_code,  bn01.bank_name,");  
		    sqlSubtail.append( "		   round(sum(decode(a05.acc_code,'910400',amt,0)) /1,0) as field_910400,");   
		    sqlSubtail.append( "		   round(sum(decode(a05.acc_code,'910500',amt,0)) /1,0) as field_910500 ");                 	  
		    sqlSubtail.append( "	from  (select * from "+cd01_table+" cd01 where cd01.hsien_id <> 'Y') cd01 ");  
		    sqlSubtail.append( "	left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = "+wlx01_m_year+ " and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL)");
		    if(bank_type.equals("ALL")){
		       sqlSubtail.append( " left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type in ('6','7') and bn01.m_year = "+wlx01_m_year+" and bn_type <> '2' and wlx01.m_year = "+wlx01_m_year);
		    }else{
		       sqlSubtail.append( "  left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=? and bn01.m_year = "+wlx01_m_year+" and bn_type <> '2' and wlx01.m_year = "+wlx01_m_year);
		       sqlSubtail_paramList.add(bank_type);
		    }	
		    sqlSubtail.append( "     left join (select * from a05 where a05.m_year  = ? and a05.m_month  = ? and a05.ACC_code in ('910400','910500','91060P') ) a05 ");
		    sqlSubtail_paramList.add(s_year);
		    sqlSubtail_paramList.add(s_month);
            
		    sqlSubtail.append( "               on  bn01.bank_no = a05.bank_code ");
		    sqlSubtail.append( "     group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME ");
		    sqlSubtail.append( " ) a05 ");  
		    //====================================================================================================
		    //2009.07.23 add A02.990611 by 2295       
		    sqlSubtail.append( " ,(  select nvl(cd01.hsien_id,' ')       as  hsien_id ,");  
		    sqlSubtail.append( "  		    nvl(cd01.hsien_name,'OTHER')  as  hsien_name,");  
		    sqlSubtail.append( "			 cd01.FR001W_output_order     as  FR001W_output_order,");  
		    sqlSubtail.append( " 			 bn01.bank_no as bank_code,  bn01.bank_name,");  
		    sqlSubtail.append( "		     round(sum(decode(a02.acc_code,'990611',amt,0)) /1,0) as field_990611,");
		    sqlSubtail.append( "             round(sum(decode(a02.acc_code,'990230',amt,0)) /1,0) as field_990230,");
		    sqlSubtail.append( "             round(sum(decode(a02.acc_code,'990240',amt,0)) /1,0) as field_990240");
		    sqlSubtail.append( "	  from  (select * from "+cd01_table+" cd01 where cd01.hsien_id <> 'Y') cd01 ");  
		    sqlSubtail.append( "	  left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = "+wlx01_m_year+ " and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL)");		   
		    if(bank_type.equals("ALL")){
			   sqlSubtail.append( "  left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type in ('6','7') and bn01.m_year = "+wlx01_m_year+" and bn_type <> '2' and wlx01.m_year = "+wlx01_m_year);
		    }else{
			   sqlSubtail.append( "  left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=? and bn01.m_year = "+wlx01_m_year+" and bn_type <> '2' and wlx01.m_year = "+wlx01_m_year);     
			   sqlSubtail_paramList.add(bank_type);
		    }
		    sqlSubtail.append( "    left join (select * from a02 where a02.m_year  = ? and a02.m_month  = ? and a02.ACC_code in ('990611','990230','990240') ) a02 ");
		    sqlSubtail_paramList.add(s_year);
		    sqlSubtail_paramList.add(s_month);
            
		    sqlSubtail.append( "               on  bn01.bank_no = a02.bank_code ");
		    sqlSubtail.append( "    group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME ");
		    sqlSubtail.append( " ) a02 ");
		    //2013.12.04 add 檢查局格式增加A99.992710建築放款餘額
		    sqlSubtail.append( " ,(  select nvl(cd01.hsien_id,' ')       as  hsien_id ,");  
            sqlSubtail.append( "            nvl(cd01.hsien_name,'OTHER')  as  hsien_name,");  
            sqlSubtail.append( "             cd01.FR001W_output_order     as  FR001W_output_order,");  
            sqlSubtail.append( "             bn01.bank_no as bank_code,  bn01.bank_name,");  
            sqlSubtail.append( "             round(sum(decode(a99.acc_code,'992710',amt,0)) /1,0) as field_992710,");
            sqlSubtail.append( "             round(sum(decode(a99.acc_code,'992810',amt,0)) /1,0) as field_992810");
            sqlSubtail.append( "      from  (select * from "+cd01_table+" cd01 where cd01.hsien_id <> 'Y') cd01 ");  
            sqlSubtail.append( "      left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = "+wlx01_m_year+ " and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL)");          
            if(bank_type.equals("ALL")){
               sqlSubtail.append( "  left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type in ('6','7') and bn01.m_year = "+wlx01_m_year+" and bn_type <> '2' and wlx01.m_year = "+wlx01_m_year);
            }else{
               sqlSubtail.append( "  left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=? and bn01.m_year = "+wlx01_m_year+" and bn_type <> '2' and wlx01.m_year = "+wlx01_m_year);     
               sqlSubtail_paramList.add(bank_type);
            }
            sqlSubtail.append( "    left join (select * from a99 where a99.m_year  = ? and a99.m_month  = ? and a99.ACC_code in ('992710','992810') ) a99 ");
            sqlSubtail_paramList.add(s_year);
            sqlSubtail_paramList.add(s_month);
            
            sqlSubtail.append( "               on  bn01.bank_no = a99.bank_code ");
            sqlSubtail.append( "    group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME ");
            sqlSubtail.append( " ) a99 ");
		    sqlSubtail.append( " where  a01.bank_no=a05.bank_code(+) and a01.bank_no=a02.bank_code(+) and a01.bank_no=a99.bank_code(+) and a01.bank_no <> ' ' ");
		    sqlSubtail.append( " ) a01");
		   
			//各縣市小計的結尾				
			sqlSubtail2.append( " ) a01,    ");
                            //2007.11.07 fix 淨值佔風險性資產比率總計.縣市別是以SUM(910400)/SUM(910500)去計算.不是以91060P/單位數 by 2295
			sqlSubtail2.append( " ( select nvl(cd01.hsien_id,' ')       as  hsien_id ,");  
			sqlSubtail2.append( "  		   nvl(cd01.hsien_name,'OTHER')  as  hsien_name,");  
			sqlSubtail2.append( "		   cd01.FR001W_output_order     as  FR001W_output_order,");  
			sqlSubtail2.append( " 		   bn01.bank_no as bank_code,  bn01.bank_name,");  
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
	        sqlSubtail2.append( "   group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order, bn01.bank_no,bn01.BANK_NAME ");
	        sqlSubtail2.append( " ) a05 "); 
			//=====================================================================================================
	        //2009.07.23 add A02.990611 by 2295       
			sqlSubtail2.append( " ,(  select nvl(cd01.hsien_id,' ')       as  hsien_id ,");  
			sqlSubtail2.append( "  		     nvl(cd01.hsien_name,'OTHER')  as  hsien_name,");  
			sqlSubtail2.append( "			 cd01.FR001W_output_order     as  FR001W_output_order,");  
			sqlSubtail2.append( " 		     bn01.bank_no as bank_code,  bn01.bank_name,");  
			sqlSubtail2.append( "		     round(sum(decode(a02.acc_code,'990611',amt,0)) /1,0) as field_990611,");
			sqlSubtail2.append( "            round(sum(decode(a02.acc_code,'990230',amt,0)) /1,0) as field_990230,");
			sqlSubtail2.append( "            round(sum(decode(a02.acc_code,'990240',amt,0)) /1,0) as field_990240");
			sqlSubtail2.append( "	  from  (select * from "+cd01_table+" cd01 where cd01.hsien_id <> 'Y') cd01 ");  
			sqlSubtail2.append( "	  left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = "+wlx01_m_year+ " and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL)");			
			if(bank_type.equals("ALL")){
			   sqlSubtail2.append( "  left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type in ('6','7') and bn01.m_year = "+wlx01_m_year+" and bn_type <> '2' and wlx01.m_year = "+wlx01_m_year);
			}else{
			   sqlSubtail2.append( "  left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=? and bn01.m_year = "+wlx01_m_year+" and bn_type <> '2' and wlx01.m_year = "+wlx01_m_year);    
			   sqlSubtail2_paramList.add(bank_type);
			}		
			sqlSubtail2.append( "   left join (select * from a02 where a02.m_year = ? and a02.m_month = ? and a02.ACC_code in ('990611','990230','990240') ) a02 on  bn01.bank_no = a02.bank_code ");
			sqlSubtail2_paramList.add(s_year);
			sqlSubtail2_paramList.add(s_month);
			sqlSubtail2.append( "    group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME ");
			sqlSubtail2.append( " ) a02 ");
			//2013.12.04 add 檢查局格式增加A99.992710建築放款餘額
			sqlSubtail2.append( " ,(  select nvl(cd01.hsien_id,' ')       as  hsien_id ,");  
            sqlSubtail2.append( "            nvl(cd01.hsien_name,'OTHER')  as  hsien_name,");  
            sqlSubtail2.append( "            cd01.FR001W_output_order     as  FR001W_output_order,");  
            sqlSubtail2.append( "            bn01.bank_no as bank_code,  bn01.bank_name,");  
            sqlSubtail2.append( "            round(sum(decode(a99.acc_code,'992710',amt,0)) /1,0) as field_992710,");
            sqlSubtail2.append( "            round(sum(decode(a99.acc_code,'992810',amt,0)) /1,0) as field_992810");
            sqlSubtail2.append( "     from  (select * from "+cd01_table+" cd01 where cd01.hsien_id <> 'Y') cd01 ");  
            sqlSubtail2.append( "     left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = "+wlx01_m_year+ " and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL)");           
            if(bank_type.equals("ALL")){
               sqlSubtail2.append( "  left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type in ('6','7') and bn01.m_year = "+wlx01_m_year+" and bn_type <> '2' and wlx01.m_year = "+wlx01_m_year);
            }else{
               sqlSubtail2.append( "  left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=? and bn01.m_year = "+wlx01_m_year+" and bn_type <> '2' and wlx01.m_year = "+wlx01_m_year);    
               sqlSubtail2_paramList.add(bank_type);
            }       
            sqlSubtail2.append( "   left join (select * from a99 where a99.m_year = ? and a99.m_month = ? and a99.ACC_code in ('992710','992810') ) a99 on  bn01.bank_no = a99.bank_code ");
            sqlSubtail2_paramList.add(s_year);
            sqlSubtail2_paramList.add(s_month);
            sqlSubtail2.append( "    group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME ");
            sqlSubtail2.append( " ) a99 ");
	 		sqlSubtail2.append( " where a01.bank_no=a05.bank_code(+) and a01.bank_no=a02.bank_code(+) and a01.bank_no=a99.bank_code(+) and a01.bank_no <> ' ' ");
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
 			sqlSubtail3.append( " 		    bn01.bank_no as bank_code,  bn01.bank_name,");  
 			sqlSubtail3.append( "		    round(sum(decode(a02.acc_code,'990611',amt,0)) /1,0) as field_990611,");
 			sqlSubtail3.append( "           round(sum(decode(a02.acc_code,'990230',amt,0)) /1,0) as field_990230,");
 			sqlSubtail3.append( "           round(sum(decode(a02.acc_code,'990240',amt,0)) /1,0) as field_990240");
 			sqlSubtail3.append( "	   from  (select * from "+cd01_table+" cd01 where cd01.hsien_id <> 'Y') cd01 ");  
 			sqlSubtail3.append( "	   left join (select * from wlx01 where m_year="+wlx01_m_year+" and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL))wlx01 on wlx01.hsien_id=cd01.hsien_id "); 			
 			if(bank_type.equals("ALL")){
 				sqlSubtail3.append( "  left join (select * from bn01 where m_year="+wlx01_m_year+" and bn_type <> '2')bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type in ('6','7') ");
 			}else{
 				sqlSubtail3.append( "  left join (select * from bn01 where m_year="+wlx01_m_year+" and bn_type <> '2')bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=? ");     
 				sqlSubtail3_paramList.add(bank_type);
 			}
 			sqlSubtail3.append( "    left join (select * from a02 where a02.m_year = ? and a02.m_month = ? and a02.ACC_code in ('990611','990230','990240') ) a02 ");
 			sqlSubtail3_paramList.add(s_year);
 	 	 	sqlSubtail3_paramList.add(s_month);
 			sqlSubtail3.append( "               on  bn01.bank_no = a02.bank_code ");
 			sqlSubtail3.append( "    group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME ");
 			sqlSubtail3.append( " ) a02 ");
 			//2013.12.04 add 檢查局格式增加A99.992710建築放款餘額     
            sqlSubtail3.append( " ,( select nvl(cd01.hsien_id,' ')       as  hsien_id ,");  
            sqlSubtail3.append( "           nvl(cd01.hsien_name,'OTHER')  as  hsien_name,");  
            sqlSubtail3.append( "           cd01.FR001W_output_order     as  FR001W_output_order,");  
            sqlSubtail3.append( "           bn01.bank_no as bank_code,  bn01.bank_name,");  
            sqlSubtail3.append( "           round(sum(decode(a99.acc_code,'992710',amt,0)) /1,0) as field_992710,");
            sqlSubtail3.append( "           round(sum(decode(a99.acc_code,'992810',amt,0)) /1,0) as field_992810");
            sqlSubtail3.append( "      from  (select * from "+cd01_table+" cd01 where cd01.hsien_id <> 'Y') cd01 ");  
            sqlSubtail3.append( "      left join (select * from wlx01 where m_year="+wlx01_m_year+" and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL))wlx01 on wlx01.hsien_id=cd01.hsien_id ");           
            if(bank_type.equals("ALL")){
                sqlSubtail3.append( "  left join (select * from bn01 where m_year="+wlx01_m_year+" and bn_type <> '2')bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type in ('6','7') ");
            }else{
                sqlSubtail3.append( "  left join (select * from bn01 where m_year="+wlx01_m_year+" and bn_type <> '2')bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=? ");     
                sqlSubtail3_paramList.add(bank_type);
            }
            sqlSubtail3.append( "    left join (select * from a99 where a99.m_year = ? and a99.m_month = ? and a99.ACC_code in ('992710','992810') ) a99 ");
            sqlSubtail3_paramList.add(s_year);
            sqlSubtail3_paramList.add(s_month);
            sqlSubtail3.append( "               on  bn01.bank_no = a99.bank_code ");
            sqlSubtail3.append( "    group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME ");
            sqlSubtail3.append( " ) a99 ");
 			sqlSubtail3.append( " where a01.bank_no=a05.bank_code(+) and a01.bank_no=a02.bank_code(+) and a01.bank_no=a99.bank_code(+) and a01.bank_no <> ' ' ");
			sqlSubtail3.append( " GROUP BY a01.hsien_id,a01.hsien_name,a01.FR001W_output_order,a01.bank_no,a01.BANK_NAME  ");
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
            sqlSub.append(sqlSubhead);//參數sqlSubhead
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
            sqlCombine.append( " hsien_id , hsien_name, FR001W_output_order, ");
            sqlCombine.append( " bank_no ,  BANK_NAME, COUNT_SEQ, field_SEQ, ");
            sqlCombine.append( " field_DEBIT,          field_CREDIT,         ");
            sqlCombine.append( " decode(sign(field_DC_RATE - 0),-1,0,field_DC_RATE) as field_DC_RATE,");//96.04.30存放比率若為負數,以0顯示
            sqlCombine.append( " field_120700,  field_OVER,  ");
            sqlCombine.append( " field_OVER_RATE,               ");
            sqlCombine.append( " field_320300,  field_TRANSFER, ");
            sqlCombine.append( " field_TRANSFER_RATE,           ");
            sqlCombine.append( " field_310000,  field_NET,      ");
            sqlCombine.append( " field_FIXNET_RATE,             ");
            sqlCombine.append( " field_CHECK_RATE,              ");
            sqlCombine.append( " field_150200,   field_BACKUP,   field_NOASSURE,   ");
            sqlCombine.append( " field_MODIFYNET,   ");
            sqlCombine.append( " field_CAPTIAL_RATE,");
            sqlCombine.append( " field_990611,       ");//98.07.22
            sqlCombine.append( " field_992710, ");//--102.12.04
            sqlCombine.append( " field_320100_rate,");//103.01.09
            sqlCombine.append( " field_120600 ");//--103.01.14
            sqlCombine.append( " from (             ");
                               
         	if(rptStyle.equals("0")){//總表	
         		sqlCombine.append(sqlTotal);
         		sqlCombine.append(" UNION ALL");
         		sqlCombine.append(sqlSub);
         		sqlCombine.append(" UNION ALL ");
         		sqlCombine.append(sqlTaiwan); 
         		sqlCombine.append(" UNION ALL ");
                sqlCombine.append(sqlFukien);
                sqlCombine.append(" )  a01 ");
                /*107.08.13 fix 縣市政府總表可看到所有農漁會信用部
                if(!hsien_id_b.equals("")){
                    sqlCombine.append(" where hsien_id=?");//105.03.16 add 縣市政府所屬縣市別                    
                }
                */
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
                /*107.08.13 fix 縣市政府總表可看到所有農漁會信用部
                if(!hsien_id_b.equals("")){
                   sqlCombine_paramList.add(hsien_id_b);//105.03.16 add 縣市政府所屬縣市別              
                }
                */
            }else{//明細表                	
            	sqlCombine.append(sqlTotal);
            	sqlCombine.append(" UNION ALL");
            	sqlCombine.append(sqlDetail);
            	sqlCombine.append(" UNION ALL");
            	sqlCombine.append(sqlSub);
            	sqlCombine.append(" )  a01 ");            	
            	if(!hsien_id_b.equals("")){
                    sqlCombine.append(" where hsien_id=?");  //105.03.16 add 縣市政府所屬縣市別                                
                }                
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
            	if(!hsien_id_b.equals("")){
                    sqlCombine_paramList.add(hsien_id_b);//105.03.16 add 縣市政府所屬縣市別              
                }                
            }
                    
            //end組合sql------------------------------------
			//建表開始--------------------------------------
			//總表
			if(rptStyle.equals("0")){
     		   System.out.println("總表sql="+sqlCombine);
     		   //printLog(logps,"總表sql="+sqlCombine);
	 		   dbData_All = DBManager.QueryDB_SQLParam(sqlCombine.toString(),sqlCombine_paramList,"hsien_name,count_seq,field_debit,"+
		                                  "field_credit,field_dc_rate,field_120700,field_over,"+
		                                  "field_over_rate,field_320300,field_transfer,field_transfer_rate"+
		                                  "field_310000,field_net,field_fixnet_rate,field_check_rate,field_150200"+
		                                  "field_backup,field_noassure,field_modifynet,field_captial_rate,field_990611,field_992710,field_320100_rate,field_120600");

	 		   System.out.print("總表資料 共"+dbData_All.size()+"筆");
	 		   //DBManager.closeQryConnection();//109.07.09 add		 
	 		   HSSFFont ft = wb.createFont();
     		   HSSFCellStyle cs = wb.createCellStyle();
	 		   ft.setFontHeightInPoints((short)18);
	 		   ft.setFontName("標楷體");
	 		   cs.setFont(ft);
	 		   cs.setAlignment(HSSFCellStyle.ALIGN_CENTER);

	   	   	   row = sheet.getRow(0);
                        
	   	   	   for(int v=0;v<23;v++){
	   	   	   	row.createCell((short)v);
	   	   	   }
                        
	   	   	   cell = row.getCell( (short) 0);
	   	   	   cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	   	   	   cell.setCellValue(""+s_year+"年"+s_month+"月"+ bank_type_name +"信用部主要經營指標總表");
	   	   	   cell.setCellStyle(cs);
                        
	   	   	   row = sheet.getRow(1);
                        
	   	   	   cell = row.getCell( (short)14);
	   	   	   cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	   	   	   cell.setCellValue("單位:新臺幣 " + unit_name + ",%");
                        
               rowNum=3;
               DataObject bean = null;
			   String insertValue = "";
			   double sampleValue;
			   HSSFFont f = wb.createFont();
               HSSFCellStyle cs2 = wb.createCellStyle();
               f.setFontHeightInPoints((short)10);
               cs2.setFont(f);
               HSSFCellStyle cs2_left = wb.createCellStyle();               
               cs2_left.setFont(f);
               cs2_left.setAlignment(HSSFCellStyle.ALIGN_LEFT);
               cs2_left.setBorderBottom(HSSFCellStyle.BORDER_THIN);
               cs2_left.setBorderLeft(HSSFCellStyle.BORDER_THIN);
               cs2_left.setBorderRight(HSSFCellStyle.BORDER_THIN);
               HSSFCellStyle cs2_center = wb.createCellStyle();               
               cs2_center.setFont(f);
               cs2_center.setAlignment(HSSFCellStyle.ALIGN_CENTER); 
               cs2_center.setBorderBottom(HSSFCellStyle.BORDER_THIN);
               cs2_center.setBorderLeft(HSSFCellStyle.BORDER_THIN);
               cs2_center.setBorderRight(HSSFCellStyle.BORDER_THIN);
               HSSFCellStyle cs2_right = wb.createCellStyle();               
               cs2_right.setFont(f);
               cs2_right.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
               cs2_right.setBorderBottom(HSSFCellStyle.BORDER_THIN);
               cs2_right.setBorderLeft(HSSFCellStyle.BORDER_THIN);
               cs2_right.setBorderRight(HSSFCellStyle.BORDER_THIN);
       		   for(int i=0;i<dbData_All.size();i++){
	   		 	 	bean = (DataObject)dbData_All.get(i);	
	   		 	    getBeanData(bean);//取得各欄位內容
	   		 	 	for(int cellcount=0;cellcount<22;cellcount++){
 		       			row = sheet.createRow(rowNum+i);
      		   			cell = row.createCell( (short)cellcount);
     		   			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			   			insertValue = "";
			   			insertValue = setInsertValue(rptStyle,cellcount);	
						insertValue = (insertValue.equals("null"))?"":insertValue;
						cell.setCellValue(insertValue);	
    			
    			        if(cellcount == 0 ){
   				           //96.04.23農會總表多一個台灣省.調整福建省位置 by 2295
    			       	   if(i==0){    			       		  
    			       		  cell.setCellStyle(cs2_left);
    			       	   }else if(Integer.parseInt(s_year) <= 99 && (i==1 || i==2 || i==3  || i==26) && (bank_type.equals("6") || bank_type.equals("ALL"))){//99.09.23 fix 99年度以前.農漁會
    			       		  cell.setCellStyle(cs2_center);
    			       	   }else if(Integer.parseInt(s_year) >= 100 && (i==1 || i==2 || i==3  || i==4 || i==5 || i==6 || i==7 || i==23) && (bank_type.equals("6") || bank_type.equals("ALL"))){//99.09.23 fix100年度以後.農漁會//103.12.23 add 桃園市
    			       	      cell.setCellStyle(cs2_center);	
    			       	   }else if(Integer.parseInt(s_year) <= 99 && (i==1 || i==2 || i==18) && bank_type.equals("7")){//99.09.23 fix 99年度以前.漁會
    			       	      cell.setCellStyle(cs2_center);
    			    	   }else if(Integer.parseInt(s_year) >= 100 && (i==1 || i==2 || i==3 || i==4 || i==5 || i==16) && bank_type.equals("7")){//99.09.23 fix 100年度以後.漁會//103.12.23 add 桃園市
    			    	      cell.setCellStyle(cs2_center);		
    			       	   }else{
    			       		  cell.setCellStyle(cs2_right);
    			       	   }
    			       }else{	 			          
	 			          cell.setCellStyle(cs2_right);
	 			       }	 			       
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
			   /*106.08.04 取消顯示
			   if(bank_type.equals("6") || bank_type.equals("ALL")){//農會/農漁會才加說明
			   rowNum++;
	           row = sheet.createRow(rowNum);
	           cell = row.createCell( (short)0);
	           cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	           cell.setCellValue("3.地區別欄之其他(農會)係指原臺灣省農會，該農會於102年5月22日更名為中華民國農會");
	           cell.setCellStyle(cs4);
			   }	
			   */  		   
			}else{//明細表
			   //printLog(logps,"明細表sql="+sqlCombine);
			   dbData_Part= DBManager.QueryDB_SQLParam(sqlCombine.toString(),sqlCombine_paramList,"hsien_name,bank_no,bank_name,count_seq,field_seq,field_debit,"+
		                                  "field_credit,field_dc_rate,field_120700,field_over,"+
		                                  "field_over_rate,field_320300,field_transfer,field_transfer_rate"+
		                                  "field_310000,field_net,field_fixnet_rate,field_check_rate,field_150200"+
		                                  "field_backup,field_noassure,field_modifynet,field_captial_rate,field_990611,field_992710,field_320100_rate,field_120600");
			   System.out.println("明細表資料 共"+dbData_Part.size()+"筆");
			   
			   HSSFFont ft = wb.createFont();
    		   HSSFCellStyle cs = wb.createCellStyle();
	 		   ft.setFontHeightInPoints((short)18);
	 		   ft.setFontName("標楷體");
	 		   cs.setFont(ft);
	 		   cs.setAlignment(HSSFCellStyle.ALIGN_CENTER);

	  		   row = sheet.createRow((short)0);
	  		   
	   		   for(int v=0;v<23;v++){//create 空白列
	   				row.createCell((short)v);
	   		   }
	   		   
	           cell = row.getCell( (short) 0);
	           cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	           cell.setCellValue(""+s_year+"年"+s_month+"月"+ bank_type_name +"信用部主要經營指標明細表");
	           cell.setCellStyle(cs);
               
	           row = sheet.getRow(1);
	           cell = row.getCell( (short)15);
	           cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	           cell.setCellValue("單位:新臺幣 " + unit_name + ",%");
               
               rowNum=2;
               DataObject bean = null;
               //共同資料的格式
               HSSFFont ft2 = wb.createFont();
               HSSFCellStyle cs3 = wb.createCellStyle();
               ft2.setFontHeightInPoints((short)10);
	           cs3.setFont(ft2);
	           
	           //設定給各地方農漁會信用部name部分 左靠
	           HSSFFont fl = wb.createFont();
               HSSFCellStyle cl = wb.createCellStyle();
               fl.setFontHeightInPoints((short)10);
	           cl.setFont(fl);
	           cl.setAlignment(HSSFCellStyle.ALIGN_LEFT);
	           cl.setBorderTop(HSSFCellStyle.BORDER_THIN);
	           cl.setBorderBottom(HSSFCellStyle.BORDER_THIN);
               cl.setBorderLeft(HSSFCellStyle.BORDER_THIN);
               cl.setBorderRight(HSSFCellStyle.BORDER_THIN);
               String insertValue = ""; 
               double sampleValue; 
               int detail_cellcount = 23;    
               
               if(!hsien_id_b.equals("")) rowNum=3; //105.03.16 縣市政府列印時,從第3列開始
       		   for(int i=(hsien_id_b.equals("")?1:0);i<dbData_Part.size();i++){               
	   		       bean = (DataObject)dbData_Part.get(i);
	   		       getBeanData(bean);	   		           
	               row = sheet.createRow(rowNum+i);
				   for(int cellcount=0;cellcount<detail_cellcount;cellcount++){ 
      		   		   cell = row.createCell( (short)cellcount);
     		           cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			           cs3.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
			           insertValue = "";
			           insertValue = setInsertValue(rptStyle,cellcount);
     		           if( cellcount==0 && field_seq.equals("A01") )insertValue =bank_no;
     		           else if( cellcount==1 && field_seq.equals("A01") )insertValue =bank_name;	
     		           else if( cellcount==1 && ( field_seq.equals("A90") ||field_seq.equals("A99") ) ) insertValue =hsien_name;
     		           
			           insertValue = (insertValue.equals("null"))?"":insertValue;
				       cell.setCellValue(insertValue);
				       //System.out.println("insertValue="+insertValue);
    			       if(cellcount == 0 || cellcount ==1 )
    			       cs3.setAlignment(HSSFCellStyle.ALIGN_CENTER);
    			       else
	 			       cs3.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
				       
				       cs3.setBorderTop(HSSFCellStyle.BORDER_THIN);
	 			       cs3.setBorderBottom(HSSFCellStyle.BORDER_THIN);
                       cs3.setBorderLeft(HSSFCellStyle.BORDER_THIN);
                       cs3.setBorderRight(HSSFCellStyle.BORDER_THIN);
                       
				       if(cellcount==1)cell.setCellStyle(cl);
				       else cell.setCellStyle(cs3);
			       }//end of cellcount
				   if(hsien_id_b.equals("")){//105.03.16 add 非縣市政府列印時,才多空一行
				      if( field_seq.equals("A90") ||field_seq.equals("A99") ){
				  	     rowNum++;row = sheet.createRow(rowNum+i); 		
			  		     for(int cellcount=0;cellcount<detail_cellcount;cellcount++){
			             	 cell = row.createCell( (short)cellcount);
			             	 cell.setCellValue(""); 
			             	 cell.setCellStyle(cs3);
			             }
			             //小計和總計後要空一行
			          }
				   }
			   }//end of dbData_Part
       		   if(hsien_id_b.equals("")){//105.03.16 add 非縣市政府列印時,才顯示總計
			      //總計再取出來一次，放在最底下
	    	      bean = (DataObject)dbData_Part.get(0);
	    	      getBeanData(bean);	    	   
			      for(int cellcount=0;cellcount<detail_cellcount;cellcount++){
 		              row = sheet.createRow(dbData_Part.size()+rowNum);
      		          cell = row.createCell( (short)cellcount);
     		          cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			          cs3.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
			          insertValue = "";
			          
			          insertValue = setInsertValue(rptStyle,cellcount);
     		          if( cellcount==1 && field_seq.equals("A99") )insertValue ="總計";     		          
     		          
			          insertValue = (insertValue.equals("null"))?"0":insertValue;
                      cell.setCellValue(insertValue);
	 		          cell.setCellStyle(cs3);
			      }//end of cellcount
       		   }//end of hsien_id_b
			   //畫最下面的小計與總計的部分
			   rowNum = 3+rowNum+dbData_Part.size();
			   if(hsien_id_b.equals("")){//105.03.16 add 非縣市政府列印時,才顯示直轄市/臺灣省
			      //列印直轄市==================================================================
			      for(int i=0;i<dbData_Part.size();i++){
	   			      bean = (DataObject)dbData_Part.get(i);
	   			      field_seq = String.valueOf(bean.getValue("field_seq"));
	   			      hsien_name = String.valueOf(bean.getValue("hsien_name"));
	   			      if(Integer.parseInt(s_year) <= 99){	
	   			      	 if( (field_seq.equals("A90") || field_seq.equals("A99")) &&
	   			   	     (hsien_name.equals("台北市") || hsien_name.equals("高雄市") ) )
	   			        { 		
	   			      	 	  getBeanData(bean);	   				       
	   			      	 	  row = sheet.createRow(rowNum++);
	   			      	 	  for(int cellcount=0;cellcount<detail_cellcount;cellcount++){
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
	   			            getBeanData(bean);	   				       
				   	     row = sheet.createRow(rowNum++);
				   	     for(int cellcount=0;cellcount<detail_cellcount;cellcount++){
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
			   
			      dbData_Taiwan= DBManager.QueryDB_SQLParam(sqlTaiwan.toString(),sqlTaiwan_paramList,"hsien_name,bank_no,bank_name,count_seq,field_seq,field_debit,"+
		                                     "field_credit,field_dc_rate,field_120700,field_over,"+
		                                     "field_over_rate,field_320300,field_transfer,field_transfer_rate"+
		                                     "field_310000,field_net,field_fixnet_rate,field_check_rate,field_150200"+
		                                     "field_backup,field_noassure,field_modifynet,field_captial_rate,field_990611,field_992710,field_320100_rate,field_120600");
			      //DBManager.closeQryConnection();//109.07.09 add		  
                 //臺灣省============================================================================			 
			     for(int i=0;i<dbData_Taiwan.size();i++){
	   			     bean = (DataObject)dbData_Taiwan.get(i);
	   			     field_seq = String.valueOf(bean.getValue("field_seq"));
	   			     hsien_name = String.valueOf(bean.getValue("hsien_name"));	   		
	   		
	   			     if( field_seq.equals("A92") && hsien_name.equals("其他")  )//105.11.04 明細表,原台灣省改為其他(含台灣省及福建省.中華民國農會)
	   			     { 		
	   			      getBeanData(bean);	   				      
					  row = sheet.createRow(rowNum++);
					  for(int cellcount=0;cellcount<detail_cellcount;cellcount++){
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
			     getBeanData(bean);				  		 
			     row = sheet.createRow(rowNum++);
			     for(int cellcount=0;cellcount<detail_cellcount;cellcount++){
 		             cell = row.createCell( (short)cellcount);
     		         cell.setEncoding(HSSFCell.ENCODING_UTF_16);
     		         insertValue = "";
     		         insertValue = setInsertValue(rptStyle,cellcount);
				     insertValue = (insertValue.equals("null"))?"":insertValue;
				     cell.setCellValue(insertValue);
	 			     cell.setCellStyle(cs3);
 		         }//end of cellcount
			  }//end of hsien_id_b 非縣市政府列印
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
    	      cell.setCellValue("1.資料來源:依各農(漁)會信用部由自有電腦設備或委由相關資訊中心以網際網路傳送之資料彙編");
		      cell.setCellStyle(cs4);
		      rowNum++;
    	      row = sheet.createRow(rowNum);
      	      cell = row.createCell( (short)0);
    	      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
    	      cell.setCellValue("2.備抵呆帳金額(農會)部份:係放款及催收之備抵呆帳合計數,(漁會)部份係放款及透支與催收之備抵呆帳合計數");
		      cell.setCellStyle(cs4);
		      if(bank_type.equals("6") || bank_type.equals("ALL")){//農會/農漁會.才加說明
		      rowNum++;
		      /*108.04.17移除中華民國農會備註
              row = sheet.createRow(rowNum);
              cell = row.createCell( (short)0);
              cell.setEncoding(HSSFCell.ENCODING_UTF_16);
              cell.setCellValue("3.地區別欄之其他(農會)係指原臺灣省農會，該農會於102年5月22日更名為中華民國農會");
              cell.setCellStyle(cs4);
              */
		      }		      
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
    private static void getBeanData(DataObject bean){
    	try{
    		
        	hsien_name = bean.getValue("hsien_name") == null?"":String.valueOf(bean.getValue("hsien_name"));//單位名稱
            bank_no = bean.getValue("bank_no") == null?"":String.valueOf(bean.getValue("bank_no"));
            bank_name = bean.getValue("bank_name") == null?"":String.valueOf(bean.getValue("bank_name"));
            count_seq = bean.getValue("count_seq") == null?"":String.valueOf(bean.getValue("count_seq"));
            field_seq = bean.getValue("field_seq") == null?"":String.valueOf(bean.getValue("field_seq"));
            field_debit = bean.getValue("field_debit") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_debit")));//存款總額
            field_credit = bean.getValue("field_credit") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_credit")));//放款總額
            field_dc_rate = bean.getValue("field_dc_rate") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_dc_rate")));//存放比率
            field_120700 = bean.getValue("field_120700") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_120700")));//內部融資餘額
            field_over = bean.getValue("field_over") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_over")));//逾放金額
            field_over_rate = bean.getValue("field_over_rate") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_over_rate")));//逾放比率
            field_320300 = bean.getValue("field_320300") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_320300")));//本期損益
            field_transfer = bean.getValue("field_transfer") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_transfer")));//轉存全國農業金庫
            field_transfer_rate = bean.getValue("field_transfer_rate") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_transfer_rate")));//轉存全國農業金庫比率
            field_310000 = bean.getValue("field_310000") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_310000")));//事業資金及公積
            field_net = bean.getValue("field_net") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_net")));//淨值
            field_fixnet_rate = bean.getValue("field_fixnet_rate") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_fixnet_rate")));//固定資產佔淨值比
            field_check_rate = bean.getValue("field_check_rate") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_check_rate")));//活存比率
            field_150200 = bean.getValue("field_150200") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_150200")));//催收款項
            field_backup = bean.getValue("field_backup") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_backup")));//備抵呆帳
            field_noassure = bean.getValue("field_noassure") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_noassure")));//無擔保放款
	        field_modifynet = bean.getValue("field_modifynet") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_modifynet")));
            field_captial_rate = bean.getValue("field_captial_rate") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_captial_rate")));//淨值佔風險性資產比率
            //98.07.22 add 990611 對直轄市、縣(市)政府、離島地區鄉(鎮、市)公所辦理之授信總額
 	 	    field_990611 = bean.getValue("field_990611") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_990611")));//990611
 	 	    //102.12.04 add A99.992710建築放款餘額
 	 	    field_992710 = bean.getValue("field_992710") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_992710")));//A99.992710
 	 	    //103.01.09 add累積虧損占信用部上年度淨算淨值之比率
 	 	    field_320100_rate = bean.getValue("field_320100_rate") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_320100_rate")));//累積虧損占信用部上年度淨算淨值之比率
 	 	    //103.01.14 add A01.120600政策性農業專案貸款餘額
            field_120600 = bean.getValue("field_120600") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_120600")));//A01.120600政策性農業專案貸款餘額            
    	}catch(Exception e){
    		System.out.println("getBeanData Error:"+e+e.getMessage());
    	}
    }
    //98.07.23 add set insertValue
   	private static String setInsertValue(String rptStyle,int cellcount){
    		 String insertValue = "";
    		 if(rptStyle.equals("0")){//總表
    		    if("其他".equals(hsien_name)){//106.08.04 總表的其他.顯示成中華民國農會
    		        hsien_name = "中華民國農會";
    		    }
    		    if ( cellcount==0 )     insertValue =hsien_name;//單位名稱
                else if( cellcount==1 ) insertValue =field_debit;//存款總額
                else if( cellcount==2 ) insertValue =field_credit;//放款總額               
                else if( cellcount==3 ) insertValue =field_dc_rate;//存放比率
                else if( cellcount==4 ) insertValue =field_120700;//內部融資餘額
                else if( cellcount==5 ) insertValue =field_over;//逾放金額
                else if( cellcount==6 ) insertValue =field_over_rate;//逾放比率
                else if( cellcount==7 ) insertValue =field_320300;//本期損益
                else if( cellcount==8 ) insertValue =field_transfer;//轉存全國農業金庫
                else if( cellcount==9 ) insertValue =field_transfer_rate;//轉存全國農業金庫比率
                else if( cellcount==10 )insertValue =field_310000;//事業資金及公積
                else if( cellcount==11 )insertValue =field_net;//淨值
                else if( cellcount==12 )insertValue =field_fixnet_rate;//固定資產佔淨值比
                else if( cellcount==13 )insertValue =field_check_rate;//活存比率
                else if( cellcount==14 )insertValue =field_150200;//催收款項
                else if( cellcount==15 )insertValue =field_backup;//備抵呆帳
                else if( cellcount==16 )insertValue =field_noassure;//無擔保放款
                //else if( cellcount==17 )insertValue =field_modifynet;
                else if( cellcount==17 )insertValue =field_990611;//對直轄市、縣(市)政府、離島地區鄉(鎮、市)公所辦理之授信總額   
                else if( cellcount==18 )insertValue =field_captial_rate;//淨值佔風險性資產比率         
                else if( cellcount==19 )insertValue =field_992710;//102.12.04 add A99.992710建築放款餘額
                else if( cellcount==20) insertValue =field_120600;//103.02.11 調整位置.政策性農業專案貸款餘額
                else if( cellcount==21 )insertValue =field_320100_rate;//103.01.09 add 累積虧損占信用部上年度決算淨值之比率 
    		 }else{
    		    if( cellcount==1 )insertValue =hsien_name;
    		    else if( cellcount==2 )insertValue =field_debit;
    		    else if( cellcount==3 )insertValue =field_credit;    		    
    		    else if( cellcount==4 )insertValue =field_dc_rate;
    		    else if( cellcount==5 )insertValue =field_120700;
    		    else if( cellcount==6 )insertValue =field_over;
    		    else if( cellcount==7 )insertValue =field_over_rate;
    		    else if( cellcount==8 )insertValue =field_320300;
    		    else if( cellcount==9 )insertValue =field_transfer;
    		    else if( cellcount==10 )insertValue =field_transfer_rate;
    		    else if( cellcount==11 )insertValue =field_310000;
    		    else if( cellcount==12 )insertValue =field_net;
    		    else if( cellcount==13 )insertValue =field_fixnet_rate;
    		    else if( cellcount==14 )insertValue =field_check_rate;
    		    else if( cellcount==15 )insertValue =field_150200;
    		    else if( cellcount==16 )insertValue =field_backup;
    		    else if( cellcount==17 )insertValue =field_noassure;
    		    else if( cellcount==18 )insertValue =field_990611;
    		    else if( cellcount==19 )insertValue =field_captial_rate;
    		    else if( cellcount==20 )insertValue =field_992710;//102.12.04 add A99.992710建築放款餘額
    		    else if( cellcount==21 )insertValue =field_120600;//103.02.11 調整位置.政策性農業專案貸款餘額
    		    else if( cellcount==22 )insertValue =field_320100_rate;//103.01.09 add 累積虧損占信用部上年度決算淨值之比率 
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
}
