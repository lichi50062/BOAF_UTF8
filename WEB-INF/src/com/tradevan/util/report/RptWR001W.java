/*
 * Created 農漁會信用部財業務資料_靜態 by 2968
 * 102.08.26 各項法定比率有違反的增加顯示***
 * 103.04.09 1. 取消顯示一、財務狀況(三)資本適足率 
             2. 二、主要業務-放款業務,增加合計欄為各上開儲存格加總,原合計欄修改為各類總額 
                                        原無申報資料,調整為空白顯示 
             3. 二、主要業務增加(三)專案農貸業務;原(三)投資業務,調整為(四)投資業務
             4.(六)列入警示項目:原與1.上月比較/2.與上季比較/3.與上年度同期比較調整文字 by 2968
 * 103.06.16 add 投資業務,增加專案基金/原A99.910401/910403/910402調整至A05取得 by 2295 
 * 103.08.28 add (七)其他,顯示違反農金法及其子法而遭處分/限制或解釋函令/舞弊案件 by2968
 * 103.09.24 add (七)其他,調整文字 by 2295
 * 103.09.25 add 調整格線顯示 by 2295
 * 103.10.01 add 2.限制或核准業務函令.欄高依限制函號or限制內容.調整高度 by 2295
 * 108.09.16 add 108年10月以後,調整為 7.購置住宅放款及房屋修繕放款餘額占存款總餘額 by 2295
 * 108.03.08 add 110年4月以後 -->改110年5月以後
 *               原4.贊助會員授信總額占贊助會員存款總額->3.贊助會員授信總額占贊助會員存款總額   
 *               原6.非會員授信總額占非會員存款總額 ->4.非會員授信總額占非會員存款總額
 *               新增
 *               4.1非會員擔保授信得徵提之擔保品種類
 *               4.2非會員擔保授信得徵提之擔保品坐落地
 *               原5.非會員無擔保消費性貸款總額占農(漁)會上年度決算淨值 ->5.非會員無擔保消費性貸款總額占農(漁)會上年度決算淨值
 *               原3.非會員存款總額占農(漁)會上年度決算淨值->6.非會員存款總額占農(漁)會上年度決算淨值
 * 110.09.11 fix 調整9.外幣資產與外幣負債差額絕對值逾新台幣100萬元且占信用部上年度決算->9.外幣風險上限占信用部前一年度決算淨值,原5%調整為10% by 2295             
 */
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.*;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

import com.tradevan.util.dao.DataObject;
import com.tradevan.util.*; 

public class RptWR001W {
    
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
    public static String createRpt(String s_year,String s_month,String unit,String bank_no){
          String errMsg = "";
          String bank_type_name="";
          String unit_name = Utility.getUnitName(unit);
          String data_year="";
          String data_month="";
          String last_year = "";
          String last_month = "";
          String lastSeason_year = "";
          String lastSeason_month = "";
          DataObject bean_Title = null;
          DataObject bean_Sum = null;
          DataObject bean_Fukien = null;
          DataObject bean_Taiwan = null;
          reportUtil reportUtil = new reportUtil();
          FileOutputStream fileOut = null;          
          HSSFRow row=null;//宣告一列
          HSSFCell cell=null;//宣告一個儲存格
          //99.09.16 add 查詢年度100年以前.縣市別不同===============================
          String cd01_table = (Integer.parseInt(s_year) < 100)?"cd01_99":"cd01"; 
          String wlx01_m_year = (Integer.parseInt(s_year) < 100)?"99":"100"; 
          String wlx01_last_m_year = ((Integer.parseInt(s_year)-1) < 100)?"99":"100"; 
          //===================================================================== 
          int idx = 0;//110.03.05 add
          last_year  = (Integer.parseInt(s_month) == 1) ? String.valueOf(Integer.parseInt(s_year) - 1) : s_year; 
          last_month = (Integer.parseInt(s_month) == 1) ? "12" : String.valueOf(Integer.parseInt(s_month) - 1);
          lastSeason_year  = (Integer.parseInt(s_month) == 3) ? String.valueOf(Integer.parseInt(s_year) - 1) : s_year; 
          if(Integer.parseInt(s_month) <= 3 && Integer.parseInt(s_month)>=1){
              lastSeason_month = "12";
          }else if(Integer.parseInt(s_month) <= 6 && Integer.parseInt(s_month)>=4){
              lastSeason_month = "3";
          }else if(Integer.parseInt(s_month) <= 9 && Integer.parseInt(s_month)>=7){
              lastSeason_month = "6";
          }else if(Integer.parseInt(s_month) <= 12 && Integer.parseInt(s_month)>=10){
              lastSeason_month = "9";
          }
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
            String openfile="農漁會信用部財業務資料_靜態"+(((Integer.parseInt(s_year) * 100 + Integer.parseInt(s_month)) >= 11005)?"_11004":"")+".xls";//要去開啟的範本檔
            
            System.out.println("開啟檔:" + openfile);
            FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+ openfile );         
            //設定FileINputStream讀取Excel檔
            
            POIFSFileSystem fs = new POIFSFileSystem( finput );//新增一個xls unit
            if(fs==null){
                System.out.println("open 範本檔失敗");
            } else{ 
                System.out.println("open 範本檔成功");
            }
            
            HSSFWorkbook wb = new HSSFWorkbook(fs);//新增一個sheet
            if(wb==null){
                System.out.println("open工作表失敗");
            } else {
                System.out.println("open 工作表成功");
            }
            for(int s=0;s<2;s++){
	            //對第一個sheet工作
	            HSSFSheet sheet = wb.getSheetAt(s);//讀取第一個工作表，宣告其為sheet 
	            if(sheet==null){
	                System.out.println("open sheet 失敗");
	            }else {
	                System.out.println("open sheet 成功");
	            }
	            
	            //做屬性設定
	            HSSFPrintSetup ps = sheet.getPrintSetup();  //取得設定
	            //sheet.setZoom(80, 100);                   //螢幕上看到的縮放大小
	            //sheet.setAutobreaks(true);                //自動分頁
	            
	            //設定頁面符合列印大小
	            sheet.setAutobreaks( false );
	            ps.setScale( ( short )67 );                 //列印縮放百分比
	            ps.setPaperSize( ( short )9 );              //設定紙張大小 A4
	            
	            //設定表頭 為固定 先設欄的起始再設列的起始
	            //wb.setRepeatingRowsAndColumns(0, 1, 17, 2, 3);
	            finput.close();
	            HSSFFooter footer = sheet.getFooter();            
		  		if(sheet==null){System.out.println("open sheet 失敗");}else System.out.println("open sheet 成功");
		  		footer.setCenter( "Page:" + HSSFFooter.page() + " of " +
	                    HSSFFooter.numPages() );                                        
		  		footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa")); 
		  		//設定表頭===============================================================================
		  		HSSFCellStyle style = wb.createCellStyle();
		        HSSFFont font = wb.createFont();
		        font.setFontHeightInPoints((short) 14);
		        style.setFont(font);
		        HSSFCellStyle style1 = wb.createCellStyle();
	            HSSFFont font1 = wb.createFont();
	            font1.setFontHeightInPoints((short) 12);
	            font1.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);//加粗  
	            style1.setFont(font1);
	            
		  		if(s==0){
			        row=(sheet.getRow((short)0)==null)? sheet.createRow((short)0) : sheet.getRow((short)0);
		            cell = (row.getCell((short)0)==null)? row.getCell((short)0) : row.getCell((short)0);
		            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		            List bank_info = Utility.getBN01(bank_no);
		            String bank_name = (String)((DataObject)bank_info.get(0)).getValue("bank_name");
		            String bank_type = (String)((DataObject)bank_info.get(0)).getValue("bank_type");//103.06.17 add
		            //String bank_name = getBank_name(bank_no,wlx01_m_year);
		            cell.setCellValue(s_year+"年度"+s_month+"月份"+bank_name+"財業務資料");
		            cell.setCellStyle(style);
		            Calendar now = Calendar.getInstance();
		            String nowYear  = String.valueOf(now.get(Calendar.YEAR)-1911); //回覆值為西元年故需-1911取得民國年;
		            String nowMonth = String.valueOf(now.get(Calendar.MONTH)+1);   //月份以0開始故加1取得實際月份;
		            String nowDay   = String.valueOf(now.get(Calendar.DATE));
		            insertCell("列印日期："+nowYear+"年"+nowMonth+"月"+nowDay+"日",wb,row,(short)8);
		            row=(sheet.getRow((short)1)==null)? sheet.createRow((short)1) : sheet.getRow((short)1);
		            insertCell("單位：新台幣"+Utility.getUnitName(unit)+"、％",wb,row,(short)8);
		            
		            //取得資料===============================================================================
		            //--一、財務狀況/二、主要業務.查詢
		            List mainInfo = mainInfo(s_year,s_month,unit,bank_no,wlx01_m_year,wlx01_last_m_year);
		            //--(一)所屬之資訊共用中心
		            List centerInfo = getInfoCenterInfo(bank_no,wlx01_m_year);
		            //--(二)稽核人員
		            List auditInfo = getAuditInfo(bank_no,wlx01_m_year);
		            //--(三)最近3次金融檢查(報告編號及基準日)—只顯示最近3筆資料.中間用空白隔開 ex: 102C102(101.7.31)  100C136(100.8.31)  099C067(99.4.30)
		            List baseDate = getBaseDate(bank_no);
		            //--(四)主要負責人一覽表(由左到右)
		            List chargePerson = getChargePerson(bank_no);
		            //--(五)組織及人員配置
		            List orgInfo = getOrgInfo(s_year,s_month,bank_no,wlx01_m_year);
		            //--(六)列入警示項目
		            List remarkInfo = getRemarkInfo(s_year,s_month,unit,bank_no);
		            //--4.2縣市別
		            List cityInfo = getCityInfo(s_year,s_month,bank_no);
		            //寫入資料=======================================================================================            
		            if(mainInfo.size()>0){
		                DataObject mainInfoBean = (DataObject)mainInfo.get(0);
		                String field_190000_cal= (mainInfoBean.getValue("field_190000_cal")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_190000_cal").toString());
		                String field_debit= (mainInfoBean.getValue("field_debit")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_debit").toString());
		                String field_credit= (mainInfoBean.getValue("field_credit")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_credit").toString());
		                String field_320300= (mainInfoBean.getValue("field_320300")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_320300").toString());
		                String field_320300_last= (mainInfoBean.getValue("field_320300_last")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_320300_last").toString());
		                String field_net= (mainInfoBean.getValue("field_net")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_net").toString());
		                String field_over_rate= (mainInfoBean.getValue("field_over_rate")==null)?"0.00":Utility.setCommaFormat(mainInfoBean.getValue("field_over_rate").toString());
		                String field_over= (mainInfoBean.getValue("field_over")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_over").toString());
		                String field_backup= (mainInfoBean.getValue("field_backup")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_backup").toString());
		                String field_backup_credit_rate= (mainInfoBean.getValue("field_backup_credit_rate")==null)?"0.00":Utility.setCommaFormat(mainInfoBean.getValue("field_backup_credit_rate").toString());
		                String field_backup_over_rate= (mainInfoBean.getValue("field_backup_over_rate")==null)?"0.00":Utility.setCommaFormat(mainInfoBean.getValue("field_backup_over_rate").toString());
		                String field_captial_rate= (mainInfoBean.getValue("field_captial_rate")==null)?"0.00":Utility.setCommaFormat(mainInfoBean.getValue("field_captial_rate").toString());
		                String field_1= (mainInfoBean.getValue("field_1")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_1").toString());
		                String field_2_1= (mainInfoBean.getValue("field_2_1")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_2_1").toString());
		                String field_2_2= (mainInfoBean.getValue("field_2_2")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_2_2").toString());
		                String field_3= (mainInfoBean.getValue("field_3")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_3").toString());
		                String field_4= (mainInfoBean.getValue("field_4")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_4").toString());
		                String field_5= (mainInfoBean.getValue("field_5")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_5").toString());
		                String field_6= (mainInfoBean.getValue("field_6")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_6").toString());
		                String field_7= (mainInfoBean.getValue("field_7")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_7").toString());
		                String field_8= (mainInfoBean.getValue("field_8")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_8").toString());
		                String field_9= (mainInfoBean.getValue("field_9")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_9").toString());
		                String field_10= (mainInfoBean.getValue("field_10")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_10").toString());
		                String field_11= (mainInfoBean.getValue("field_11")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_11").toString());
		                String field_1_violate= (mainInfoBean.getValue("field_1_violate")==null)?"":String.valueOf(mainInfoBean.getValue("field_1_violate"));
		                String field_2_1_violate= (mainInfoBean.getValue("field_2_1_violate")==null)?"":String.valueOf(mainInfoBean.getValue("field_2_1_violate").toString());
		                String field_2_2_violate= (mainInfoBean.getValue("field_2_2_violate")==null)?"":String.valueOf(mainInfoBean.getValue("field_2_2_violate").toString());
		                String field_3_violate= (mainInfoBean.getValue("field_3_violate")==null)?"":String.valueOf(mainInfoBean.getValue("field_3_violate").toString());
		                String field_4_violate= (mainInfoBean.getValue("field_4_violate")==null)?"":String.valueOf(mainInfoBean.getValue("field_4_violate").toString());
		                String field_990624= (mainInfoBean.getValue("field_990624")==null)?"":String.valueOf(mainInfoBean.getValue("field_990624").toString());//110.03.05 add
		                String field_4_1_violate= (mainInfoBean.getValue("field_4_1_violate")==null)?"":String.valueOf(mainInfoBean.getValue("field_4_1_violate").toString());//110.03.05 add		                
		                String field_4_2_violate= (mainInfoBean.getValue("field_4_2_violate")==null)?"":String.valueOf(mainInfoBean.getValue("field_4_2_violate").toString());//110.03.05 add
		                
		                String field_4_2 =  "";
		                if(cityInfo.size()>0){ 
		                	DataObject cityInfoBean = (DataObject)cityInfo.get(0);
		                	field_4_2 = (cityInfoBean.getValue("amt_name")==null)?"":String.valueOf(cityInfoBean.getValue("amt_name").toString());//110.03.05 add
		                }
		                //System.out.println("field_990624="+field_990624);
		                //System.out.println("field_4_1_violate="+field_4_1_violate);
		                //System.out.println("field_4_2_violate="+field_4_2_violate);
		                String field_5_violate= (mainInfoBean.getValue("field_5_violate")==null)?"":String.valueOf(mainInfoBean.getValue("field_5_violate").toString());
		                String field_6_violate= (mainInfoBean.getValue("field_6_violate")==null)?"":String.valueOf(mainInfoBean.getValue("field_6_violate").toString());
		                String field_7_violate= (mainInfoBean.getValue("field_7_violate")==null)?"":String.valueOf(mainInfoBean.getValue("field_7_violate").toString());
		                String field_8_violate= (mainInfoBean.getValue("field_8_violate")==null)?"":String.valueOf(mainInfoBean.getValue("field_8_violate").toString());
		                String field_9_violate= (mainInfoBean.getValue("field_9_violate")==null)?"":String.valueOf(mainInfoBean.getValue("field_9_violate").toString());
		                String field_10_violate= (mainInfoBean.getValue("field_10_violate")==null)?"":String.valueOf(mainInfoBean.getValue("field_10_violate").toString());
		                String field_debit_1= (mainInfoBean.getValue("field_debit_1")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_debit_1").toString());
		                String field_992130= (mainInfoBean.getValue("field_992130")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_992130").toString());
		                String field_990420= (mainInfoBean.getValue("field_990420")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_990420").toString());
		                String field_990310= (mainInfoBean.getValue("field_990310")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_990310").toString());
		                String field_992140= (mainInfoBean.getValue("field_992140")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_992140").toString());
		                String field_990410= (mainInfoBean.getValue("field_990410")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_990410").toString());
		                String field_990610_990611= (mainInfoBean.getValue("field_990610_990611")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_990610_990611").toString());
		                String field_120700= (mainInfoBean.getValue("field_120700")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_120700").toString());
		                String field_990611= (mainInfoBean.getValue("field_990611")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_990611").toString());
		                String field_2_1_sum= (mainInfoBean.getValue("field_2_1_sum")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_2_1_sum").toString());
		                String field_credit_1= (mainInfoBean.getValue("field_credit_1")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_credit_1").toString());
		                String field_990510= (mainInfoBean.getValue("field_990510")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_990510").toString());
		                String field_2_2_sum= (mainInfoBean.getValue("field_2_2_sum")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_2_2_sum").toString());
		                String field_noassure= (mainInfoBean.getValue("field_noassure")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_noassure").toString());
		                String field_992710= (mainInfoBean.getValue("field_992710")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_992710").toString());
		                String field_992510= (mainInfoBean.getValue("field_992510")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_992510").toString());
		                String field_992520= (mainInfoBean.getValue("field_992520")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_992520").toString());
		                String field_992530= (mainInfoBean.getValue("field_992530")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_992530").toString());
		                String field_992540= (mainInfoBean.getValue("field_992540")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_992540").toString());
		                String field_2_3_sum= (mainInfoBean.getValue("field_2_3_sum")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_2_3_sum").toString());
		                String field_over_1= (mainInfoBean.getValue("field_over_1")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_over_1").toString());
		                String field_992510_992140_rate= (mainInfoBean.getValue("field_992510_992140_rate")==null)?"0.00":Utility.setCommaFormat(mainInfoBean.getValue("field_992510_992140_rate").toString());
		                String field_992520_990410_rate= (mainInfoBean.getValue("field_992520_990410_rate")==null)?"0.00":Utility.setCommaFormat(mainInfoBean.getValue("field_992520_990410_rate").toString());
		                String field_992530_990610_cal_rate= (mainInfoBean.getValue("field_992530_990610_cal_rate")==null)?"0.00":Utility.setCommaFormat(mainInfoBean.getValue("field_992530_990610_cal_rate").toString());
		                String field_992540_120700_rate= (mainInfoBean.getValue("field_992540_120700_rate")==null)?"0.00":Utility.setCommaFormat(mainInfoBean.getValue("field_992540_120700_rate").toString());
		                String field_2_4_sum= (mainInfoBean.getValue("field_2_4_sum")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_2_4_sum").toString());
		                String field_over_rate_1= (mainInfoBean.getValue("field_over_rate_1")==null)?"0.00":Utility.setCommaFormat(mainInfoBean.getValue("field_over_rate_1").toString());
		                String field_992610= (mainInfoBean.getValue("field_992610")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_992610").toString());
		                String field_992620= (mainInfoBean.getValue("field_992620")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_992620").toString());
		                String field_992630= (mainInfoBean.getValue("field_992630")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_992630").toString());
		                String field_992640= (mainInfoBean.getValue("field_992640")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_992640").toString());
		                String field_2_5_sum= (mainInfoBean.getValue("field_2_5_sum")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_2_5_sum").toString());
		                String field_992610_cal= (mainInfoBean.getValue("field_992610_cal")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_992610_cal").toString());
		                String field_992510_992610= (mainInfoBean.getValue("field_992510_992610")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_992510_992610").toString());
		                String field_992520_992620= (mainInfoBean.getValue("field_992520_992620")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_992520_992620").toString());
		                String field_992530_992630= (mainInfoBean.getValue("field_992530_992630")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_992530_992630").toString());
		                String field_992540_992640= (mainInfoBean.getValue("field_992540_992640")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_992540_992640").toString());
		                String field_2_6_sum= (mainInfoBean.getValue("field_2_6_sum")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_2_6_sum").toString());
		                String field_over_992610_cal= (mainInfoBean.getValue("field_over_992610_cal")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_over_992610_cal").toString());
		                String field_ab_992140_rate= (mainInfoBean.getValue("field_ab_992140_rate")==null)?"0.00":Utility.setCommaFormat(mainInfoBean.getValue("field_ab_992140_rate").toString());
		                String field_ab_990410_rate= (mainInfoBean.getValue("field_ab_990410_rate")==null)?"0.00":Utility.setCommaFormat(mainInfoBean.getValue("field_ab_990410_rate").toString());
		                String field_ab_990610_rate= (mainInfoBean.getValue("field_ab_990610_rate")==null)?"0.00":Utility.setCommaFormat(mainInfoBean.getValue("field_ab_990610_rate").toString());
		                String field_ab_120700_rate= (mainInfoBean.getValue("field_ab_120700_rate")==null)?"0.00":Utility.setCommaFormat(mainInfoBean.getValue("field_ab_120700_rate").toString());
		                String field_2_7_sum= (mainInfoBean.getValue("field_2_7_sum")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_2_7_sum").toString());
		                String field_ab_credit_rate= (mainInfoBean.getValue("field_ab_credit_rate")==null)?"0.00":Utility.setCommaFormat(mainInfoBean.getValue("field_ab_credit_rate").toString());
		                String field_110600= (mainInfoBean.getValue("field_110600")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_110600").toString());
		                String field_130200= (mainInfoBean.getValue("field_130200")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_130200").toString());
		                String field_130100= (mainInfoBean.getValue("field_130100")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_130100").toString());//103.06.16 add 專案基金
		                String field_910401= (mainInfoBean.getValue("field_910401")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_910401").toString());
		                String field_910403= (mainInfoBean.getValue("field_910403")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_910403").toString());
		                String field_910402= (mainInfoBean.getValue("field_910402")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_910402").toString());
		                String field_130200_other= (mainInfoBean.getValue("field_130200_other")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_130200_other").toString());
		                String field_loan_bal_amt= (mainInfoBean.getValue("field_loan_bal_amt")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_loan_bal_amt").toString());
		                String field_over6m_loan_bal_amt= (mainInfoBean.getValue("field_over6m_loan_bal_amt")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_over6m_loan_bal_amt").toString());
		                String field_over_rate_argi= (mainInfoBean.getValue("field_over_rate_argi")==null)?"0.00":Utility.setCommaFormat(mainInfoBean.getValue("field_over_rate_argi").toString());
		                String field_delay_loan_cnt_sum= (mainInfoBean.getValue("field_delay_loan_cnt_sum")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_delay_loan_cnt_sum").toString());
		                String field_delay_loan_amt_sum= (mainInfoBean.getValue("field_delay_loan_amt_sum")==null)?"0":Utility.setCommaFormat(mainInfoBean.getValue("field_delay_loan_amt_sum").toString());
		                
		                //一、財務狀況 (一)主要資產、負債與經營績效
		                row=(sheet.getRow((short)3)==null)? sheet.createRow((short)3) : sheet.getRow((short)3);
		                insertCell(field_190000_cal,wb,row,(short)2);//資產總額
		                insertCell(field_debit,wb,row,(short)7);//存款   
		                row=(sheet.getRow((short)4)==null)? sheet.createRow((short)4) : sheet.getRow((short)4);
		                insertCell(field_credit,wb,row,(short)2);//放款
		                insertCell(field_320300,wb,row,(short)7);//本期損益 
		                row=(sheet.getRow((short)5)==null)? sheet.createRow((short)5) : sheet.getRow((short)5);
		                insertCell(field_320300_last,wb,row,(short)7);//去年度損益    
		                //一、財務狀況 (二)資產品質
		                row=(sheet.getRow((short)7)==null)? sheet.createRow((short)7) : sheet.getRow((short)7);
		                insertCell(field_net,wb,row,(short)2);//淨值
		                insertCell(field_over_rate,wb,row,(short)7);//逾放比率   
		                row=(sheet.getRow((short)8)==null)? sheet.createRow((short)8) : sheet.getRow((short)8);
		                insertCell(field_over,wb,row,(short)2);//逾期放款
		                insertCell(field_backup,wb,row,(short)7);//備抵呆帳
		                row=(sheet.getRow((short)9)==null)? sheet.createRow((short)9) : sheet.getRow((short)9);
		                insertCell(field_backup_credit_rate,wb,row,(short)2);//放款覆蓋率
		                insertCell(field_backup_over_rate,wb,row,(short)7);//備抵呆帳覆蓋率  
		                /*//一、財務狀況 (三)資本適足率
		                row=(sheet.getRow((short)10)==null)? sheet.createRow((short)10) : sheet.getRow((short)10);
		                insertCell(field_captial_rate,wb,row,(short)2,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_LEFT,true);*/
		                //一、財務狀況 (四)各項法定比率
		                row=(sheet.getRow((short)11)==null)? sheet.createRow((short)11) : sheet.getRow((short)11);
		                insertCell(field_1_violate,wb,row,(short)0);
		                insertCell("1.存放比率："+field_1,wb,row,(short)1);
		                row=(sheet.getRow((short)12)==null)? sheet.createRow((short)12) : sheet.getRow((short)12);
		                insertCell(field_2_1_violate,wb,row,(short)0);
		                insertCell("2.1內部融資占上年度信用部決算淨值："+field_2_1,wb,row,(short)1);
		                row=(sheet.getRow((short)13)==null)? sheet.createRow((short)13) : sheet.getRow((short)13);
		                insertCell(field_2_2_violate,wb,row,(short)0);
		                insertCell("2.2內部融資(中、長期)占上年度信用部決算淨值："+field_2_2,wb,row,(short)1);
		                
		                //110.03.05 add 110年4月(含)以後套用新格式
		                if((Integer.parseInt(s_year) * 100 + Integer.parseInt(s_month)) >= 11005){
		                	row=(sheet.getRow((short)14)==null)? sheet.createRow((short)14) : sheet.getRow((short)14);
		                	insertCell(field_4_violate,wb,row,(short)0);
		                    insertCell("3.贊助會員授信總額占贊助會員存款總額："+field_4,wb,row,(short)1);
		                    
		                    
		                    row=(sheet.getRow((short)15)==null)? sheet.createRow((short)15) : sheet.getRow((short)15);
			                insertCell(field_6_violate,wb,row,(short)0);
			                insertCell("4.非會員授信總額占非會員存款總額："+field_6,wb,row,(short)1);
			                
			                row=(sheet.getRow((short)16)==null)? sheet.createRow((short)16) : sheet.getRow((short)16);
			                insertCell(field_4_1_violate,wb,row,(short)0);
		                    insertCell("4.1非會員擔保授信得徵提之擔保品種類："+field_990624,wb,row,(short)1);
		                    
		                    
		                    row=(sheet.getRow((short)17)==null)? sheet.createRow((short)17) : sheet.getRow((short)17);
		                    insertCell(field_4_2_violate,wb,row,(short)0);
		                    insertCell("4.2非會員擔保授信得徵提之擔保品坐落地："+field_4_2,wb,row,(short)1);
		                    
		                    row=(sheet.getRow((short)18)==null)? sheet.createRow((short)18) : sheet.getRow((short)18);		                    
		                    insertCell(field_5_violate,wb,row,(short)0);
		                    insertCell("5.非會員無擔保消費性貸款總額占農(漁)會上年度決算淨值："+field_5,wb,row,(short)1);
		                    
		                    row=(sheet.getRow((short)19)==null)? sheet.createRow((short)19) : sheet.getRow((short)19);
			                insertCell(field_3_violate,wb,row,(short)0);
			                insertCell("6.非會員存款總額占農(漁)會上年度決算淨值："+field_3,wb,row,(short)1);
			                
			                row=(sheet.getRow((short)20)==null)? sheet.createRow((short)20) : sheet.getRow((short)20);
			                insertCell(field_7_violate,wb,row,(short)0);
			                insertCell("7.購置住宅放款及房屋修繕放款餘額占存款總餘額："+field_7,wb,row,(short)1);
			                
			                row=(sheet.getRow((short)21)==null)? sheet.createRow((short)21) : sheet.getRow((short)21);
			                insertCell(field_8_violate,wb,row,(short)0);
			                insertCell("8.固定資產淨額占農(漁)會信用部上年度決算淨值："+field_8,wb,row,(short)1);
			                
			                row=(sheet.getRow((short)22)==null)? sheet.createRow((short)22) : sheet.getRow((short)22);
			                insertCell(field_9_violate,wb,row,(short)0);
			                insertCell("9.外幣風險上限占信用部前一年度決算淨值："+field_9,wb,row,(short)1);//110.09.11調整名稱
			                
			                row=(sheet.getRow((short)23)==null)? sheet.createRow((short)23) : sheet.getRow((short)23);
			                insertCell(field_10_violate,wb,row,(short)0);
			                insertCell("10.理監事職員及利害關係人擔保授信餘額占農(漁)會上年度決算淨值："+field_10,wb,row,(short)1);
			                
			                row=(sheet.getRow((short)24)==null)? sheet.createRow((short)24) : sheet.getRow((short)24);
			                insertCell("11.合格淨值占風險性資產總額："+field_11,wb,row,(short)1);
			    
		                }else{
		                
		                
		                	row=(sheet.getRow((short)14)==null)? sheet.createRow((short)14) : sheet.getRow((short)14);
		                	insertCell(field_3_violate,wb,row,(short)0);
		                	insertCell("3.非會員存款總額占農(漁)會上年度決算淨值："+field_3,wb,row,(short)1);
		                	row=(sheet.getRow((short)15)==null)? sheet.createRow((short)15) : sheet.getRow((short)15);
		                	insertCell(field_4_violate,wb,row,(short)0);
		                	insertCell("4.贊助會員授信總額占贊助會員存款總額："+field_4,wb,row,(short)1);
		                	row=(sheet.getRow((short)16)==null)? sheet.createRow((short)16) : sheet.getRow((short)16);
		                	insertCell(field_5_violate,wb,row,(short)0);
		                	insertCell("5.非會員無擔保消費性貸款總額占農(漁)會上年度決算淨值："+field_5,wb,row,(short)1);
		                	row=(sheet.getRow((short)17)==null)? sheet.createRow((short)17) : sheet.getRow((short)17);
		                	insertCell(field_6_violate,wb,row,(short)0);
		                	insertCell("6.非會員授信總額占非會員存款總額："+field_6,wb,row,(short)1);
		                	row=(sheet.getRow((short)18)==null)? sheet.createRow((short)18) : sheet.getRow((short)18);
		                	insertCell(field_7_violate,wb,row,(short)0);
		                	//108.09.16 add 108年10月以後套用新格式
		                	if((Integer.parseInt(s_year) * 100 + Integer.parseInt(s_month)) >= 10810){
		                		insertCell("7.購置住宅放款及房屋修繕放款餘額占存款總餘額："+field_7,wb,row,(short)1);
		                	}else{
		                		insertCell("7.自用住宅放款總額占定期性存款總額："+field_7,wb,row,(short)1);
		                	}
		                	row=(sheet.getRow((short)19)==null)? sheet.createRow((short)19) : sheet.getRow((short)19);
		                	insertCell(field_8_violate,wb,row,(short)0);
		                	insertCell("8.固定資產淨額占農(漁)會信用部上年度決算淨值："+field_8,wb,row,(short)1);
		                	row=(sheet.getRow((short)20)==null)? sheet.createRow((short)20) : sheet.getRow((short)20);
		                	insertCell(field_9_violate,wb,row,(short)0);
		                	insertCell("9.外幣風險上限占信用部前一年度決算淨值："+field_9,wb,row,(short)1);//110.09.11 調整名稱
		                	row=(sheet.getRow((short)21)==null)? sheet.createRow((short)21) : sheet.getRow((short)21);
		                	insertCell(field_10_violate,wb,row,(short)0);
		                	insertCell("10.理監事職員及利害關係人擔保授信餘額占農(漁)會上年度決算淨值："+field_10,wb,row,(short)1);
		                	row=(sheet.getRow((short)22)==null)? sheet.createRow((short)22) : sheet.getRow((short)22);
		                	insertCell("11.合格淨值占風險性資產總額："+field_11,wb,row,(short)1);
		                }
		                
		                //二、主要業務(一)存款業務
		                idx = ((Integer.parseInt(s_year) * 100 + Integer.parseInt(s_month)) >= 11005) ? 28:26;
		                row=(sheet.getRow((short)idx)==null)? sheet.createRow((short)idx) : sheet.getRow((short)idx);
		                insertCell(field_debit_1,wb,row,(short)1);
		                insertCell(field_992130,wb,row,(short)3);
		                insertCell(field_990420,wb,row,(short)5);
		                insertCell(field_990310,wb,row,(short)7);
		                //二、主要業務(二)放款業務
		                idx = ((Integer.parseInt(s_year) * 100 + Integer.parseInt(s_month)) >= 11005) ? 31:29;
		                row=(sheet.getRow((short)idx)==null)? sheet.createRow((short)idx) : sheet.getRow((short)idx);
		                insertCell(field_992140,wb,row,(short)2);//放款-正會員
		                insertCell(field_992510,wb,row,(short)5);//逾期放款(A)-正會員
		                insertCell(field_992510_992140_rate,wb,row,(short)6);//逾放比率-正會員
		                insertCell(field_992610,wb,row,(short)7);//應予觀察放款(B)-正會員
		                insertCell(field_992510_992610,wb,row,(short)8);//廣義逾放(A+B)-正會員
		                insertCell(field_ab_992140_rate,wb,row,(short)9);//廣義逾放比-正會員
		                idx++;
		                row=(sheet.getRow((short)idx)==null)? sheet.createRow((short)idx) : sheet.getRow((short)idx);
		                insertCell(field_990410,wb,row,(short)2);//放款-贊助會員
		                insertCell(field_992520,wb,row,(short)5);//逾期放款(A)-贊助會員
		                insertCell(field_992520_990410_rate,wb,row,(short)6);//逾放比率-贊助會員
		                insertCell(field_992620,wb,row,(short)7);//應予觀察放款(B)-贊助會員
		                insertCell(field_992520_992620,wb,row,(short)8);//廣義逾放(A+B)-贊助會員
		                insertCell(field_ab_990410_rate,wb,row,(short)9);//廣義逾放比-贊助會員
		                idx++;
		                row=(sheet.getRow((short)idx)==null)? sheet.createRow((short)idx) : sheet.getRow((short)idx);
		                insertCell(field_990610_990611,wb,row,(short)2);//放款-非會員
		                insertCell(field_990510,wb,row,(short)3);//無擔保放款-非會員
		                insertCell(field_992530,wb,row,(short)5);//逾期放款(A)-非會員
		                insertCell(field_992530_990610_cal_rate,wb,row,(short)6);//逾放比率-非會員
		                insertCell(field_992630,wb,row,(short)7);//應予觀察放款(B)-非會員
		                insertCell(field_992530_992630,wb,row,(short)8);//廣義逾放(A+B)-非會員
		                insertCell(field_ab_990610_rate,wb,row,(short)9);//廣義逾放比-非會員
		                idx++;
		                row=(sheet.getRow((short)idx)==null)? sheet.createRow((short)idx) : sheet.getRow((short)idx);
		                insertCell(field_120700,wb,row,(short)2);//放款-內部融資
		                insertCell(field_992540,wb,row,(short)5);//逾期放款(A)-內部融資
		                insertCell(field_992540_120700_rate,wb,row,(short)6);//逾放比率-內部融資
		                insertCell(field_992640,wb,row,(short)7);//應予觀察放款(B)-內部融資
		                insertCell(field_992540_992640,wb,row,(short)8);//廣義逾放(A+B)-內部融資
		                insertCell(field_ab_120700_rate,wb,row,(short)9);//廣義逾放比-內部融資
		                
		                //縣市政府貸款
		                idx++;
		                row=(sheet.getRow((short)idx)==null)? sheet.createRow((short)idx) : sheet.getRow((short)idx);
		                insertCell(field_990611,wb,row,(short)2);//放款-縣市政府貸款
		                //合計
		                idx++;
		                row=(sheet.getRow((short)idx)==null)? sheet.createRow((short)idx) : sheet.getRow((short)idx);
		                insertCell(field_2_1_sum,wb,row,(short)2);//放款-合計
		                insertCell(field_2_2_sum,wb,row,(short)3);//無擔保放款-合計
		                insertCell(field_992710,wb,row,(short)4);//建築放款-合計
		                insertCell(field_2_3_sum,wb,row,(short)5);//逾期放款(A)-合計
		                insertCell(field_2_4_sum,wb,row,(short)6);//逾放比率-合計
		                insertCell(field_2_5_sum,wb,row,(short)7);//應予觀察放款(B)-合計
		                insertCell(field_2_6_sum,wb,row,(short)8);//廣義逾放(A+B)-合計
		                insertCell(field_2_7_sum,wb,row,(short)9);//廣義逾放比-合計
		                //各類總額
		                idx++;
		                row=(sheet.getRow((short)idx)==null)? sheet.createRow((short)idx) : sheet.getRow((short)idx);
		                insertCell(field_credit_1,wb,row,(short)2);//放款-各類總額
		                insertCell(field_noassure,wb,row,(short)3);//無擔保放款-各類總額
		                insertCell(field_992710,wb,row,(short)4);//建築放款-合計
		                insertCell(field_over_1,wb,row,(short)5);//逾期放款-各類總額
		                insertCell(field_over_rate_1,wb,row,(short)6);//逾放比率-各類總額
		                insertCell(field_992610_cal,wb,row,(short)7);//應予觀察放款(B)-各類總額
		                insertCell(field_over_992610_cal,wb,row,(short)8);//廣義逾放(A+B)-各類總額
		                insertCell(field_ab_credit_rate,wb,row,(short)9);//廣義逾放比-各類總額
		                
		                //二、主要業務(三)專案農貸業務
		                idx = ((Integer.parseInt(s_year) * 100 + Integer.parseInt(s_month)) >= 11005) ? 40:38;
		                row=(sheet.getRow((short)idx)==null)? sheet.createRow((short)idx) : sheet.getRow((short)idx);		               
		                insertCell(field_loan_bal_amt,wb,row,(short)1);//專案農貸業務.專案農貸餘額[(7)貸放餘額]
		                insertCell(field_over6m_loan_bal_amt,wb,row,(short)3);//專案農貸業務.逾期放款[(9)逾期六個月放款餘額-金額]
		                insertCell(field_over_rate_argi,wb,row,(short)4);//專案農貸業務.逾放比率[(9)逾期六個月放款餘額-金額/(7)貸放餘額-金額]
		                insertCell(field_delay_loan_cnt_sum,wb,row,(short)5);//專案農貸業務.當年度累計核准延期還款件數[(10)當月核准延期還款-件數]
		                insertCell(field_delay_loan_amt_sum,wb,row,(short)7);//專案農貸業務.當年度累計核准延期還款金額[(11)當月核准延期還款-金額]
		                //二、主要業務(四)投資業務
		                idx = ((Integer.parseInt(s_year) * 100 + Integer.parseInt(s_month)) >= 11005) ? 42:40;
		                row=(sheet.getRow((short)idx)==null)? sheet.createRow((short)idx) : sheet.getRow((short)idx);
		                idx++;
		                insertCell(field_110600,wb,row,(short)2);//有價證券
		                row=(sheet.getRow((short)idx)==null)? sheet.createRow((short)idx) : sheet.getRow((short)idx);
		                String field_130200_tmp = "2.基金及"+(bank_type.equals("6")?"出資":"投資")+"：長期投資 "+field_130200+"；專案基金 "+field_130100;//103.06.16 add
		                insertCell(field_130200_tmp,wb,row,(short)1);//長期投資 //103.06.16 add專案基金
		                idx = ((Integer.parseInt(s_year) * 100 + Integer.parseInt(s_month)) >= 11005) ? 45:43;
		                row=(sheet.getRow((short)idx)==null)? sheet.createRow((short)idx) : sheet.getRow((short)idx);
		                insertCell(field_910401,wb,row,(short)1);//長期投資.全國農業金庫股票
		                insertCell(field_910403,wb,row,(short)3);//長期投資.合作金庫股票
		                insertCell(field_910402,wb,row,(short)5);//長期投資.財金資訊(股)公司股票
		                insertCell(field_130200_other,wb,row,(short)7);//長期投資.其他
		            }else{
		                row=(sheet.getRow((short)1)==null)? sheet.createRow((short)1) : sheet.getRow((short)1);
		                insertCell("一、財務狀況 (無資料)",wb,row,(short)0);
		                idx = ((Integer.parseInt(s_year) * 100 + Integer.parseInt(s_month)) >= 11005) ? 25:23;
		                row=(sheet.getRow((short)23)==null)? sheet.createRow((short)idx) : sheet.getRow((short)idx);
		                insertCell("二、主要業務 (無資料)",wb,row,(short)0);
		            }
		           if(centerInfo.size()>0){
		               DataObject centerInfoBean = (DataObject)centerInfo.get(0);
		               //三、其他事項(一)所屬之資訊共用中心
		               idx = ((Integer.parseInt(s_year) * 100 + Integer.parseInt(s_month)) >= 11005) ? 47:45;
		               row=(sheet.getRow((short)idx)==null)? sheet.createRow((short)idx) : sheet.getRow((short)idx);
		               insertCell(Utility.getTrimString(centerInfoBean.getValue("bank_name")),wb,row,(short)3);
		           }
		           //三、其他事項(二)稽核人員：
		           if(auditInfo.size()>0){
		        	   String audit_name = "";
		        	   idx = ((Integer.parseInt(s_year) * 100 + Integer.parseInt(s_month)) >= 11005) ? 48:46;
		               row=(sheet.getRow((short)idx)==null)? sheet.createRow((short)idx) : sheet.getRow((short)idx);
		               for(int i=0;i<auditInfo.size();i++){
		            	   DataObject b = (DataObject)auditInfo.get(i);
		            	   if(!"".equals(audit_name))audit_name+="/";
		            	   audit_name+=Utility.getTrimString(b.getValue("audit_name"));
		               }
		               insertCell(audit_name,wb,row,(short)2);
		           }
		           if(baseDate.size()>0){
		               //三、其他事項(三)最近3次金融檢查(報告編號及基準日)：
		        	   idx = ((Integer.parseInt(s_year) * 100 + Integer.parseInt(s_month)) >= 11005) ? 49:47;
		               row=(sheet.getRow((short)idx)==null)? sheet.createRow((short)idx) : sheet.getRow((short)idx);
		               int cellNum = 4;
		               int size = (baseDate.size()>=3)?3:baseDate.size();
		               for(int i=0;i<size;i++){
		                   DataObject baseDateBean = (DataObject)baseDate.get(i);
		                   String base_date= (baseDateBean.getValue("base_date")==null)?"":baseDateBean.getValue("base_date").toString();
		                   insertCell(base_date,wb,row,cellNum);
		                   cellNum+=2;
		               }
		           }
		           if(chargePerson.size()>0){
		               //三、其他事項(四)主要負責人一覽表(由左到右)
		               int rowNum = ((Integer.parseInt(s_year) * 100 + Integer.parseInt(s_month)) >= 11005) ? 52:50;
		               int cellNum = 1;
		               for(int i=0;i<chargePerson.size();i++){
		                   DataObject chargePersonBean = (DataObject)chargePerson.get(i);
		                   String cmuse_name= (chargePersonBean.getValue("cmuse_name")==null)?"":chargePersonBean.getValue("cmuse_name").toString();
		                   String name= (chargePersonBean.getValue("name")==null)?"":chargePersonBean.getValue("name").toString();
		                   if(i%2==0 && i!=0){
		                       rowNum += 1;
		                       cellNum = 1;
		                   }
		                   row=(sheet.getRow((short)rowNum)==null)? sheet.createRow((short)rowNum) : sheet.getRow((short)rowNum);
		                   insertCell1(cmuse_name,wb,row,cellNum,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_LEFT,true);
		                   cellNum += 1;
		                   insertCell1(name,wb,row,cellNum,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_LEFT,true);
		                   cellNum += 3;
		               }
		           }
		           
		           if(orgInfo.size()>0){
		               //三、其他事項(五)組織及人員配置：信用部員工○人
		        	   idx=((Integer.parseInt(s_year) * 100 + Integer.parseInt(s_month)) >= 11005) ? 54:52;
		               row=(sheet.getRow((short)idx)==null)? sheet.createRow((short)idx) : sheet.getRow((short)idx);
		               String credit_staff_num= (((DataObject)orgInfo.get(0)).getValue("credit_staff_num")==null)?"0":Utility.setCommaFormat(((DataObject)orgInfo.get(0)).getValue("credit_staff_num").toString());
		               insertCell1("(五)組織及人員配置：信用部員工"+credit_staff_num+"人",wb,row,1,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_LEFT,false);
		               int rowNum = ((Integer.parseInt(s_year) * 100 + Integer.parseInt(s_month)) >= 11005) ? 56:54;
		               for(int i=0;i<orgInfo.size();i++){
		                   DataObject orgInfoBean = (DataObject)orgInfo.get(i);
		                   String bankName= (orgInfoBean.getValue("bank_name")==null)?"":orgInfoBean.getValue("bank_name").toString();
		                   String staff_num= (orgInfoBean.getValue("staff_num")==null)?"":Utility.setCommaFormat(orgInfoBean.getValue("staff_num").toString());
		                   String addr= (orgInfoBean.getValue("addr")==null)?"":orgInfoBean.getValue("addr").toString();
		                   String telno= (orgInfoBean.getValue("telno")==null)?"":orgInfoBean.getValue("telno").toString();
		                   String cmuse_name= (orgInfoBean.getValue("cmuse_name")==null)?"":orgInfoBean.getValue("cmuse_name").toString();
		                   String name= (orgInfoBean.getValue("name")==null)?"":orgInfoBean.getValue("name").toString();
		                   row=(sheet.getRow((short)rowNum)==null)? sheet.createRow((short)rowNum) : sheet.getRow((short)rowNum);
		                   insertCell1(bankName,wb,row,1,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
		                   insertCell1(staff_num,wb,row,2,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT,true);
		                   insertCell1(addr,wb,row,3,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_LEFT,true);
		                   insertCell1("",wb,row,4,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_LEFT,true);
		                   insertCell1("",wb,row,5,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
		                   sheet.addMergedRegion(new Region((short)rowNum,(short)3,(short)rowNum,(short)5));
		                   insertCell1(telno,wb,row,6,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_LEFT,true);
		                   insertCell1("",wb,row,7,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
		                   sheet.addMergedRegion(new Region((short)rowNum,(short)6,(short)rowNum,(short)7));
		                   insertCell1(cmuse_name,wb,row,8,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
		                   insertCell1(name,wb,row,9,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
		                   rowNum += 1;
		               }
		           }else{
		               insertCell1("(五)組織及人員配置：信用部員工0人",wb,sheet.getRow((short)idx),1,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_LEFT,false);
		           }
		           
		           //(六)列入警示項目：
		           int rowNum =((Integer.parseInt(s_year) * 100 + Integer.parseInt(s_month)) >= 11005) ? 56:54;
		           if(orgInfo.size()>0){
		               rowNum += orgInfo.size();
		           }else{
		               rowNum = ((Integer.parseInt(s_year) * 100 + Integer.parseInt(s_month)) >= 11005) ? 60:58;
		           }
		           String tmpTitle = s_year+"年"+s_month+"月農漁會信用部營運狀況警訊報表";
		           row=(sheet.getRow((short)rowNum)==null)? sheet.createRow((short)rowNum) : sheet.getRow((short)rowNum);
		           insertCell1("(六)列入警示項目：",wb,row,1,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_LEFT,false); 
		           boolean isWriteRpt0 = false;
		           boolean isWriteRpt1 = false;
		           boolean isWriteRpt2 = false;
		           if(remarkInfo.size()>0){
		               int listN=1;
		               String wr_rpt ="";
		               String type ="";
		               String amt ="";
		               String remark ="";
		               for(int i=0;i<remarkInfo.size();i++){
		                   DataObject remarkInfoBean = (DataObject)remarkInfo.get(i);
		                   String cellVal = "";
		                   wr_rpt= (remarkInfoBean.getValue("wr_rpt")==null)?"":remarkInfoBean.getValue("wr_rpt").toString();
		                   type= (remarkInfoBean.getValue("type")==null)?"":remarkInfoBean.getValue("type").toString();
		                   if("4".equals(type)){
		                       amt=(remarkInfoBean.getValue("amt")==null)?"0.00":Utility.setCommaFormat(remarkInfoBean.getValue("amt").toString());
		                   }else{
		                       amt=(remarkInfoBean.getValue("amt")==null)?"0":Utility.setCommaFormat((new BigDecimal(String.valueOf(Double.parseDouble(remarkInfoBean.getValue("amt").toString())/Integer.parseInt(unit))).setScale(0, BigDecimal.ROUND_HALF_UP)).toString());
		                   }
		                   remark= (remarkInfoBean.getValue("remark")==null)?"":remarkInfoBean.getValue("remark").toString();
		                   if(i==0){
		                       rowNum++;
		                       row=(sheet.getRow((short)rowNum)==null)? sheet.createRow((short)rowNum) : sheet.getRow((short)rowNum);
		                       if("0".equals(wr_rpt)){
		                           for(int k=1;k<=9;k++){
		                               cell = (row.getCell((short)k)==null)? row.createCell((short)k) : row.getCell((short)k);
		                           }
		                           cell = row.getCell((short)1);
		                           cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		                           cell.setCellValue("1."+tmpTitle+"與上月份("+last_year+"年"+last_month+"月)比較");
		                           cell.setCellStyle(style1);
		                           sheet.addMergedRegion(new Region((short)rowNum,(short)1,(short)rowNum,(short)9));
		                           isWriteRpt0 = true;
		                       }
		                       if("1".equals(wr_rpt)){
		                           for(int k=1;k<=9;k++){
		                               cell = (row.getCell((short)k)==null)? row.createCell((short)k) : row.getCell((short)k);
		                           }
		                           cell = row.getCell((short)1);
		                           cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		                           cell.setCellValue("1."+tmpTitle+"與上月份("+last_year+"年"+last_month+"月)比較：無");
		                           cell.setCellStyle(style1);
		                           sheet.addMergedRegion(new Region((short)rowNum,(short)1,(short)rowNum,(short)9));
		                           rowNum++;
		                           row=(sheet.getRow((short)rowNum)==null)? sheet.createRow((short)rowNum) : sheet.getRow((short)rowNum);
		                           for(int k=1;k<=9;k++){
		                               cell = (row.getCell((short)k)==null)? row.createCell((short)k) : row.getCell((short)k);
		                           }
		                           cell = row.getCell((short)1);
		                           cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		                           cell.setCellValue("2."+tmpTitle+"與上一季("+lastSeason_year+"年"+lastSeason_month+"月)比較");
		                           cell.setCellStyle(style1);
		                           sheet.addMergedRegion(new Region((short)rowNum,(short)1,(short)rowNum,(short)9));
		                           isWriteRpt1 = true;
		                       }
		                       if("2".equals(wr_rpt)){
		                           for(int k=1;k<=9;k++){
		                               cell = (row.getCell((short)k)==null)? row.createCell((short)k) : row.getCell((short)k);
		                           }
		                           cell = row.getCell((short)1);
		                           cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		                           cell.setCellValue("1."+tmpTitle+"與上月份("+last_year+"年"+last_month+"月)比較：無");
		                           cell.setCellStyle(style1);
		                           sheet.addMergedRegion(new Region((short)rowNum,(short)1,(short)rowNum,(short)9));
		                           rowNum++;
		                           row=(sheet.getRow((short)rowNum)==null)? sheet.createRow((short)rowNum) : sheet.getRow((short)rowNum);
		                           for(int k=1;k<=9;k++){
		                               cell = (row.getCell((short)k)==null)? row.createCell((short)k) : row.getCell((short)k);
		                           }
		                           cell = row.getCell((short)1);
		                           cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		                           cell.setCellValue("2."+tmpTitle+"與上一季("+lastSeason_year+"年"+lastSeason_month+"月)比較：無");
		                           cell.setCellStyle(style1);
		                           sheet.addMergedRegion(new Region((short)rowNum,(short)1,(short)rowNum,(short)9));
		                           rowNum++;
		                           row=(sheet.getRow((short)rowNum)==null)? sheet.createRow((short)rowNum) : sheet.getRow((short)rowNum);
		                           for(int k=1;k<=9;k++){
		                               cell = (row.getCell((short)k)==null)? row.createCell((short)k) : row.getCell((short)k);
		                           }
		                           cell = row.getCell((short)1);
		                           cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		                           cell.setCellValue("3."+tmpTitle+"與上一年度同期("+String.valueOf(Integer.parseInt(s_year)-1)+"年"+String.valueOf(Integer.parseInt(s_month))+"月)比較");
		                           cell.setCellStyle(style1);
		                           sheet.addMergedRegion(new Region((short)rowNum,(short)1,(short)rowNum,(short)9));
		                           isWriteRpt2 = true;
		                       }
		                   }else{ 
		                       if(!wr_rpt.equals(((DataObject)remarkInfo.get(i-1)).getValue("wr_rpt"))){
		                           rowNum++;
		                           row=(sheet.getRow((short)rowNum)==null)? sheet.createRow((short)rowNum) : sheet.getRow((short)rowNum);
		                           if("1".equals(wr_rpt)){
		                               cellVal = "2."+tmpTitle+"與上一季("+lastSeason_year+"年"+lastSeason_month+"月)比較";
		                           }else if("2".equals(wr_rpt)){
		                               if(!isWriteRpt1){
		                                   cellVal = "2."+tmpTitle+"與上一季("+lastSeason_year+"年"+lastSeason_month+"月)比較：無";
		                                   for(int k=1;k<=9;k++){
		                                       cell = (row.getCell((short)k)==null)? row.createCell((short)k) : row.getCell((short)k);
		                                   }
		                                   cell = row.getCell((short)1);
		                                   cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		                                   cell.setCellValue(cellVal);
		                                   cell.setCellStyle(style1);
		                                   sheet.addMergedRegion(new Region((short)rowNum,(short)1,(short)rowNum,(short)9));
		                                   rowNum++;
		                                   row=(sheet.getRow((short)rowNum)==null)? sheet.createRow((short)rowNum) : sheet.getRow((short)rowNum);
		                               }
		                               cellVal = "3."+tmpTitle+"與上一年度同期("+String.valueOf(Integer.parseInt(s_year)-1)+"年"+String.valueOf(Integer.parseInt(s_month))+"月)比較";
		                           }
		                           for(int k=1;k<=9;k++){
		                               cell = (row.getCell((short)k)==null)? row.createCell((short)k) : row.getCell((short)k);
		                           }
		                           cell = row.getCell((short)1);
		                           cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		                           cell.setCellValue(cellVal);
		                           cell.setCellStyle(style1);
		                           sheet.addMergedRegion(new Region((short)rowNum,(short)1,(short)rowNum,(short)9));
		                           listN = 1;
		                       }
		                   }
		                   if(!"".equals(remark)){
		                       rowNum++;
		                       row=(sheet.getRow((short)rowNum)==null)? sheet.createRow((short)rowNum) : sheet.getRow((short)rowNum);
		                       if("0".equals(wr_rpt)){
		                           cellVal = "1."+String.valueOf(listN)+"."+remark+"："+amt;
		                       }
		                       if("1".equals(wr_rpt)){
		                           cellVal = "2."+String.valueOf(listN)+"."+remark+"："+amt;
		                       }
		                       if("2".equals(wr_rpt)){
		                           cellVal = "3."+String.valueOf(listN)+"."+remark+"："+amt;
		                       }
		                       insertCell1(cellVal,wb,row,1,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_LEFT,true);
		                       sheet.addMergedRegion(new Region((short)rowNum,(short)1,(short)rowNum,(short)9));
		                       listN++;
		                   }
		                   if(i==remarkInfo.size()-1){
		                       if("0".equals(wr_rpt)){
		                           rowNum++;
		                           row=(sheet.getRow((short)rowNum)==null)? sheet.createRow((short)rowNum) : sheet.getRow((short)rowNum);
		                           for(int k=1;k<=9;k++){
		                               cell = (row.getCell((short)k)==null)? row.createCell((short)k) : row.getCell((short)k);
		                           }
		                           cell = row.getCell((short)1);
		                           cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		                           cell.setCellValue("2."+tmpTitle+"與上一季("+lastSeason_year+"年"+lastSeason_month+"月)比較：無");
		                           cell.setCellStyle(style1);
		                           sheet.addMergedRegion(new Region((short)rowNum,(short)1,(short)rowNum,(short)9));
		                           rowNum++;
		                           row=(sheet.getRow((short)rowNum)==null)? sheet.createRow((short)rowNum) : sheet.getRow((short)rowNum);
		                           for(int k=1;k<=9;k++){
		                               cell = (row.getCell((short)k)==null)? row.createCell((short)k) : row.getCell((short)k);
		                           }
		                           cell = row.getCell((short)1);
		                           cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		                           cell.setCellValue("3."+tmpTitle+"與上一年度同期("+String.valueOf(Integer.parseInt(s_year)-1)+"年"+String.valueOf(Integer.parseInt(s_month))+"月)比較：無");
		                           cell.setCellStyle(style1);
		                           sheet.addMergedRegion(new Region((short)rowNum,(short)1,(short)rowNum,(short)9));
		                       }
		                       if("1".equals(wr_rpt)){
		                           rowNum++;
		                           row=(sheet.getRow((short)rowNum)==null)? sheet.createRow((short)rowNum) : sheet.getRow((short)rowNum);
		                           for(int k=1;k<=9;k++){
		                               cell = (row.getCell((short)k)==null)? row.createCell((short)k) : row.getCell((short)k);
		                           }
		                           cell = row.getCell((short)1);
		                           cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		                           cell.setCellValue("3."+tmpTitle+"與上一年度同期("+String.valueOf(Integer.parseInt(s_year)-1)+"年"+String.valueOf(Integer.parseInt(s_month))+"月)比較：無");
		                           cell.setCellStyle(style1);
		                           sheet.addMergedRegion(new Region((short)rowNum,(short)1,(short)rowNum,(short)9));
		                           
		                       }
		                   }
		               }
		           }else{
		               
		               rowNum++;
		               row=(sheet.getRow((short)rowNum)==null)? sheet.createRow((short)rowNum) : sheet.getRow((short)rowNum);
		               for(int k=1;k<=9;k++){
		                   cell = (row.getCell((short)k)==null)? row.createCell((short)k) : row.getCell((short)k);
		               }
		               cell = row.getCell((short)1);
		               cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		               cell.setCellValue("1."+tmpTitle+"與上月份("+last_year+"年"+last_month+"月)比較：無");
		               cell.setCellStyle(style1);
		               sheet.addMergedRegion(new Region((short)rowNum,(short)1,(short)rowNum,(short)9));
		               rowNum++;
		               row=(sheet.getRow((short)rowNum)==null)? sheet.createRow((short)rowNum) : sheet.getRow((short)rowNum);
		               for(int k=1;k<=9;k++){
		                   cell = (row.getCell((short)k)==null)? row.createCell((short)k) : row.getCell((short)k);
		               }
		               cell = row.getCell((short)1);
		               cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		               cell.setCellValue("2."+tmpTitle+"與上一季("+lastSeason_year+"年"+lastSeason_month+"月)比較：無");
		               cell.setCellStyle(style1);
		               sheet.addMergedRegion(new Region((short)rowNum,(short)1,(short)rowNum,(short)9));
		               rowNum++;
		               row=(sheet.getRow((short)rowNum)==null)? sheet.createRow((short)rowNum) : sheet.getRow((short)rowNum);
		               for(int k=1;k<=9;k++){
		                   cell = (row.getCell((short)k)==null)? row.createCell((short)k) : row.getCell((short)k);
		               }
		               cell = row.getCell((short)1);
		               cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		               cell.setCellValue("3."+tmpTitle+"與上一年度同期("+String.valueOf(Integer.parseInt(s_year)-1)+"年"+String.valueOf(Integer.parseInt(s_month))+"月)比較：無");
		               cell.setCellStyle(style1);
		               sheet.addMergedRegion(new Region((short)rowNum,(short)1,(short)rowNum,(short)9));
		            }
		  		}else if(s==1){
		  			//違反農金法及子法需遭處分
		            List violateInfo = getViolateInfo(bank_no,wlx01_m_year);
		            List violate_type_dbdate= getViolate_typeList();
		            List violate_type_list= new LinkedList();
		            DataObject violate_type_bean = null;
		        	int rowNum = 1;
		        	sheet.setColumnWidth((short)3,(short)(256 * 30));     //設定欄位寬度 
		        	if(violateInfo.size()>0){
		        		row=(sheet.getRow((short)rowNum)==null)? sheet.createRow((short)rowNum) : sheet.getRow((short)rowNum);
			        	insertCell1("1.違反農金法及其子法而遭處分",wb,row,1,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_LEFT,true);
			        	insertCell1("",wb,row,2,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_LEFT,true);
			        	insertCell1("",wb,row,3,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_LEFT,true);
			        	sheet.addMergedRegion(new Region((short)rowNum,(short)1,(short)rowNum,(short)3));
			        	rowNum++;
			            row=(sheet.getRow((short)rowNum)==null)? sheet.createRow((short)rowNum) : sheet.getRow((short)rowNum);
			        	insertCell1("受處分日期",wb,row,1,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
			        	insertCell1("處分方式",wb,row,2,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
			        	insertCell1("主旨",wb,row,3,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
			        	insertCell1("說明",wb,row,4,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
			        	insertCell1("",wb,row,5,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
			        	insertCell1("法令依據",wb,row,6,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
			        	sheet.addMergedRegion(new Region((short)rowNum,(short)4,(short)rowNum,(short)5));
			        	rowNum++;
			        	for(int i=0;i<violateInfo.size();i++){
			        		DataObject b = (DataObject)violateInfo.get(i);
			        		String violate_date= Utility.getTrimString(b.getValue("violate_date"));
			        		String violate_type_name= Utility.getTrimString(b.getValue("violate_type"));
			        		if(!"".equals(violate_type_name)){
				            	violate_type_list = Utility.getStringTokenizerData(violate_type_name,":");
				            	violate_type_name = "";
				        		for(int j=0;j<violate_type_list.size();j++){//處分方式
				        			System.out.println("j="+(String)violate_type_list.get(j));
				        			violate_type_loop:
				        			for(int k=0;k<violate_type_dbdate.size();k++){//代碼檔
				        				violate_type_bean = (DataObject)violate_type_dbdate.get(k);  	
				        				if(((String)violate_type_list.get(j)).equals(violate_type_bean.getValue("cmuse_id"))){
				        					violate_type_name += violate_type_bean.getValue("cmuse_name");
				        					break violate_type_loop;
				        				}
				        			}
				        		}
				        	}
			        		String title= Utility.getTrimString(b.getValue("title"));
			        		String content= Utility.getTrimString(b.getValue("content"));
			        		String law_content= Utility.getTrimString(b.getValue("law_content"));
			        		row=(sheet.getRow((short)rowNum)==null)? sheet.createRow((short)rowNum) : sheet.getRow((short)rowNum);
				        	insertCell1(violate_date,wb,row,1,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
				        	insertCell1(violate_type_name,wb,row,2,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
				        	insertCell1(title,wb,row,3,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
				        	insertCell1(content,wb,row,4,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
				        	insertCell1("",wb,row,5,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
				        	insertCell1(law_content,wb,row,6,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
				        	sheet.addMergedRegion(new Region((short)rowNum,(short)4,(short)rowNum,(short)5));
				        	int size = (content.length() / 12)+1 ;
				        	row.setHeight((short)(256 * size));
				        	rowNum++;
			        	}
			        	 
		            }else{
		            	row=(sheet.getRow((short)rowNum)==null)? sheet.createRow((short)rowNum) : sheet.getRow((short)rowNum);
		            	insertCell1("1.違反農金法及其子法而遭處分:無",wb,row,1,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_LEFT,true);
			        	insertCell1("",wb,row,2,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_LEFT,true);
			        	insertCell1("",wb,row,3,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_LEFT,true);
			        	sheet.addMergedRegion(new Region((short)rowNum,(short)1,(short)rowNum,(short)3));
			        	rowNum++;
			        }
		        	//限制或核准業務函令
		        	List loalInfo = getLoalInfo(bank_no,wlx01_m_year);
		        	sheet.setColumnWidth((short)5,(short)(256 * 40));     //設定欄位寬度 
		        	if(loalInfo.size()>0){
		        		row=(sheet.getRow((short)rowNum)==null)? sheet.createRow((short)rowNum) : sheet.getRow((short)rowNum);
		        		if(violateInfo.size()>0){
		        			insertCell1("2.限制或核准業務函令",wb,row,1,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_LEFT,true);
				        	insertCell1("",wb,row,2,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_LEFT,true);
		            	}else{
		            		insertCell1("2.限制或核准業務函令",wb,row,1,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_LEFT,true);
				        	insertCell1("",wb,row,2,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_LEFT,true);
		            	}
		        		
			        	
			        	sheet.addMergedRegion(new Region((short)rowNum,(short)1,(short)rowNum,(short)2));
			        	
			        	rowNum++;
			            row=(sheet.getRow((short)rowNum)==null)? sheet.createRow((short)rowNum) : sheet.getRow((short)rowNum);
			        	insertCell1("限制函號",wb,row,1,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
			        	insertCell1("限制內容",wb,row,2,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
			        	insertCell1("",wb,row,3,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
			        	sheet.addMergedRegion(new Region((short)rowNum,(short)2,(short)rowNum,(short)3));
			        	insertCell1("狀態",wb,row,4,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
			        	insertCell1("備註",wb,row,5,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
			        	insertCell1("",wb,row,6,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
			        	sheet.addMergedRegion(new Region((short)rowNum,(short)5,(short)rowNum,(short)6));
			        	rowNum++;
			        	for(int i=0;i<loalInfo.size();i++){
			        		DataObject b = (DataObject)loalInfo.get(i);
			        		String loal_number= Utility.getTrimString(b.getValue("loal_number"));
			        		String loal_content= Utility.getTrimString(b.getValue("loal_content"));
			        		String select_name= Utility.getTrimString(b.getValue("select_name"));
			        		String loal_ps= Utility.getTrimString(b.getValue("loal_ps"));
			        		row=(sheet.getRow((short)rowNum)==null)? sheet.createRow((short)rowNum) : sheet.getRow((short)rowNum);
			        		insertCell1(loal_number,wb,row,1,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT,true);
				        	insertCell1(loal_content,wb,row,2,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
				        	insertCell1("",wb,row,3,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
				        	sheet.addMergedRegion(new Region((short)rowNum,(short)2,(short)rowNum,(short)3));
				        	insertCell1(select_name,wb,row,4,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
				        	insertCell1(loal_ps,wb,row,5,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
				        	insertCell1("",wb,row,6,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
				        	sheet.addMergedRegion(new Region((short)rowNum,(short)5,(short)rowNum,(short)6));
				        	int size = (loal_number.length() > loal_content.length())?(loal_number.length() / 4 )+1:(loal_content.length() / 10 )+1;//103.10.01 add				        	
				        	row.setHeight((short)(256 * size));
				        	rowNum++;
			        	}
			        	
		            }else{
		            	row=(sheet.getRow((short)rowNum)==null)? sheet.createRow((short)rowNum) : sheet.getRow((short)rowNum);
		            	//if(violateInfo.size()>0){
				        //	insertCell("2.限制或核准業務函令:無",wb,row,1,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_LEFT,true);
				        //	insertCell("",wb,row,2,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_LEFT,true);
				        //	insertCell("",wb,row,3,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_LEFT,true);
		            	//}else{
		            		insertCell1("2.限制或核准業務函令:無",wb,row,1,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_LEFT,true);
				        	insertCell1("",wb,row,2,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_LEFT,true);
				        	insertCell1("",wb,row,3,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_LEFT,true);
		            	//}
			        	sheet.addMergedRegion(new Region((short)rowNum,(short)1,(short)rowNum,(short)3));
			        	rowNum++;
			        }
		        	
		        	//舞幣案件
		        	List emInfo = getEmInfo(bank_no,wlx01_m_year);
		        	sheet.setColumnWidth((short)6,(short)(256 * 30));     //設定欄位寬度  
		        	if(emInfo.size()>0){
		        		row=(sheet.getRow((short)rowNum)==null)? sheet.createRow((short)rowNum) : sheet.getRow((short)rowNum);
		        		//if(loalInfo.size()>0){
		            	//	insertCell("3.舞弊案件",wb,row,1,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_LEFT,false);
		            	//}else{
		            		insertCell1("3.舞弊案件",wb,row,1,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_LEFT,false);
		            	//}
			        	rowNum++;
			            row=(sheet.getRow((short)rowNum)==null)? sheet.createRow((short)rowNum) : sheet.getRow((short)rowNum);
			        	insertCell1("來文文號",wb,row,1,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
			        	insertCell1("來文日期",wb,row,2,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
			        	insertCell1("舞弊人",wb,row,3,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
			        	insertCell1("舞弊內容",wb,row,4,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
			        	insertCell1("",wb,row,5,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
			        	insertCell1("",wb,row,6,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
			        	sheet.addMergedRegion(new Region((short)rowNum,(short)4,(short)rowNum,(short)6));
			        	rowNum++;
			        	 
			        	for(int i=0;i<emInfo.size();i++){
			        		DataObject b = (DataObject)emInfo.get(i);
			        		String em_number= Utility.getTrimString(b.getValue("em_number"));
			        		String em_date= Utility.getTrimString(b.getValue("em_date"));
			        		String em_emer= Utility.getTrimString(b.getValue("em_emer"));
			        		String em_content= Utility.getTrimString(b.getValue("em_content"));
			        		row=(sheet.getRow((short)rowNum)==null)? sheet.createRow((short)rowNum) : sheet.getRow((short)rowNum);
			        		insertCell1(em_number,wb,row,1,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
				        	insertCell1(em_date,wb,row,2,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
				        	insertCell1(em_emer,wb,row,3,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
				        	insertCell1(em_content,wb,row,4,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
				        	insertCell1("",wb,row,5,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
				        	insertCell1("",wb,row,6,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_LEFT,true);
				        	sheet.addMergedRegion(new Region((short)rowNum,(short)4,(short)rowNum,(short)6));
				        	int size = (em_content.length() / 20)+1 ;				        
				        	row.setHeight((short)(256 * size));
				        	rowNum++;
			        	}
			        	
		            }else{
		            	row=(sheet.getRow((short)rowNum)==null)? sheet.createRow((short)rowNum) : sheet.getRow((short)rowNum);
		            	//if(loalInfo.size()>0){
		            	//	insertCell("3.舞弊案件:無",wb,row,1,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_LEFT,false);
		            	//}else{
		            		insertCell1("3.舞弊案件:無",wb,row,1,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_LEFT,false);
		            	//}
		            	rowNum++;
			        }
		  		}
		  		
           }
           // Write the output to a file============================   
           fileOut = new FileOutputStream( Utility.getProperties("reportDir")+System.getProperty("file.separator")+"農漁會信用部財業務資料_靜態.xls" );
           wb.write( fileOut );
           fileOut.close();            
           System.out.println("儲存完成");
        } catch ( Exception e ) {            
           e.printStackTrace();
           
        } finally {
           try {
               if ( fileOut != null ) {
                   fileOut.close();
               }
           } catch ( Exception e ) {
                 System.out.println(e.getMessage() );
           }
        }
      return errMsg;
    }
    
   	public static void printLog(PrintStream logps,String errRptMsg){
        if(!errRptMsg.equals("")){
           logcalendar = Calendar.getInstance(); 
           nowlog = logcalendar.getTime();
           logps.println(logformat.format(nowlog)+errRptMsg);
           logps.flush();
        }
   }
   //一、財務狀況/二、主要業務.查詢
   public static List mainInfo(String s_year,String s_month,String unit,String bank_no,String wlx01_m_year,String wlx01_last_m_year){
       StringBuffer sqlCmd = new StringBuffer(); 
       List paramList = new ArrayList();
       sqlCmd.append(" select a01.bank_no,");
       sqlCmd.append("        a01.bank_name,"); //--機構名稱
       sqlCmd.append("        round(field_190000_cal/?,0) as field_190000_cal,");//-- 一、(一).資產總額
       sqlCmd.append("        round(field_DEBIT/?,0) as field_DEBIT,");//-- 一、(一).存款
       sqlCmd.append("        round(field_CREDIT/?,0) as field_CREDIT,");//-- 一、(一).放款
       sqlCmd.append("        round(field_320300/?,0) as field_320300,");//-- 一、(一).本期損益
       sqlCmd.append("        round(field_320300_last/?,0) as field_320300_last,");//-- 一、(一).去年度損益
       sqlCmd.append("        round(field_NET/?,0) as field_NET,");//--(二)淨值
       sqlCmd.append("        decode(a01.field_CREDIT,0,0,round(a01.field_OVER /  a01.field_CREDIT *100 ,2))  as   field_OVER_RATE,");//-- (二)逾放比率
       sqlCmd.append("        round(field_OVER/?,0) as field_OVER,");//-- (二)逾期放款
       sqlCmd.append("        round(field_BACKUP/?,0) as field_BACKUP,");//-- (二)備抵呆帳
       sqlCmd.append("        decode(a01.field_CREDIT,0,0,round(a01.field_BACKUP /  a01.field_CREDIT *100 ,2)) as   field_BACKUP_CREDIT_RATE,");//-- (二)放款覆蓋率=備抵呆帳/放款
       sqlCmd.append("        decode(a01.field_OVER,0,0,round(a01.field_BACKUP /  a01.field_OVER *100 ,2)) as   field_BACKUP_OVER_RATE,");//-- (二)備抵呆帳覆蓋率=備抵呆帳/逾期放款
       sqlCmd.append("        round(field_CAPTIAL /  1000 ,2)  as   field_CAPTIAL_RATE,");  //--(三)資本適足率
       sqlCmd.append("        field_1,decode(field_1_violate,1,'***','') as field_1_violate,");//--(四)1.存放比率  102.08.26各項法定比率有違反的增加顯示***
       sqlCmd.append("        field_2_1,decode(field_2_1_violate,1,'***','') as field_2_1_violate,");//-- (四)2.1內部融資占上年度信用部決算淨值
       sqlCmd.append("        field_2_2,decode(field_2_2_violate,1,'***','') as field_2_2_violate,");//-- (四)2.2內部融資(中、長期)占上年度信用部決算淨值
       sqlCmd.append("        field_3,decode(field_3_violate,1,'***','') as field_3_violate,");//-- (四)3.非會員存款總額占農(漁)會上年度決算淨值
       sqlCmd.append("        field_4,decode(field_4_violate,1,'***','') as field_4_violate,");//-- (四)4.贊助會員授信總額占贊助會員存款總額
       sqlCmd.append("        field_5,decode(field_5_violate,1,'***','') as field_5_violate,");//-- (四)5.非會員無擔保消費性貸款總額占農(漁)會上年度決算淨值
       sqlCmd.append("        field_6,decode(field_6_violate,1,'***','') as field_6_violate,");//-- (四)6.非會員授信總額占非會員存款總額
       sqlCmd.append("        field_7,decode(field_7_violate,1,'***','') as field_7_violate,");//-- (四)7.自用住宅放款總額占定期性存款總額
       sqlCmd.append("        field_8,decode(field_8_violate,1,'***','') as field_8_violate,");//-- (四)8.固定資產淨額占農(漁)會信用部上年度決算淨值
       sqlCmd.append("        field_9,decode(field_9_violate,1,'***','') as field_9_violate,");//-- (四)9.外幣風險上限占信用部前一年度決算淨值 //110.09.11調整名稱
       sqlCmd.append("        field_10,decode(field_10_violate,1,'***','') as field_10_violate,");//-- (四)10.理監事職員及利害關係人擔保授信餘額占農(漁)會上年度決算淨值
       sqlCmd.append("        decode(field_990624,'1','逾放比率低於5%,已申請經主管機關同意者,得以土地或建築物等不動產、動產為擔保品','2','逾放比率高於5%未達10%,已申請經主管機關同意者, 得以住宅、已取得建築執照或雜項執照之建築基地為擔保品','3','逾放比率高10%,得以住宅為擔保品','') as field_990624,");//--4.1非會員擔保授信得徵提之擔保品種類 //110.03.04 add
       sqlCmd.append("        decode((field_990624_1_rule+field_990624_2_rule),1,'***','') as field_4_1_violate,");//4.1 顯示違反 
       sqlCmd.append("        decode(field_990626_rule,1,'***','') as field_4_2_violate,");//4.2 顯示違反
       sqlCmd.append("        round(field_CAPTIAL /  1000 ,2)  as   field_11,");  //--(四)11.合格淨值占風險性資產總額
       sqlCmd.append("        round(field_DEBIT/?,0) as field_DEBIT_1,");//--二、(一)存款業務.存款總額
       sqlCmd.append("        round(field_992130/?,0) as field_992130,");//--二、(一)存款業務.正會員
       sqlCmd.append("        round(field_990420/?,0) as field_990420,");//--二、(一)存款業務.贊助會員
       sqlCmd.append("        round(field_990310/?,0) as field_990310,");//--二、(一)存款業務.非會員
       sqlCmd.append("        round(field_992140/?,0) as field_992140,");//--(二)放款業務.放款-正會員
       sqlCmd.append("        round(field_990410/?,0) as field_990410,");//-- (二)放款業務.放款-贊助會員
       sqlCmd.append("        round((field_990610-field_990611)/?,0) as field_990610_990611,");//-- (二)放款業務.放款-非會員
       sqlCmd.append("        round(field_120700/?,0) as field_120700,");//-- (二)放款業務.放款-內部融資
       sqlCmd.append("        round(field_990611/?,0) as field_990611,");//-- (二)放款業務.放款-縣市政府貸款
       sqlCmd.append("        round((field_992140+field_990410+((field_990610-field_990611))+field_120700+field_990611)/?,0) as  field_2_1_sum,");//--放款業務.放款-合計
       sqlCmd.append("        round(field_CREDIT/?,0) as field_CREDIT_1,");//-- (二)放款業務.放款-各類總額
       sqlCmd.append("        round(field_990510/?,0) as field_990510,");//-- (二)無擔保放款-非會員
       sqlCmd.append("        round(field_990510/?,0) as field_2_2_sum,");//--無擔保放款-合計
       sqlCmd.append("        round(field_NOASSURE/?,0) as field_NOASSURE,");//-- (二)無擔保放款--各類總額-
       sqlCmd.append("        round(field_992710/?,0) as field_992710,");//-- (二)建築放款-合計
       sqlCmd.append("        round(field_992510/?,0) as field_992510,");//-- (二)逾期放款(A)-正會員
       sqlCmd.append("        round(field_992520/?,0) as field_992520,");//-- (二)逾期放款(A)-贊助會員
       sqlCmd.append("        round(field_992530/?,0) as field_992530,");//-- (二)逾期放款(A)-非會員
       sqlCmd.append("        round(field_992540/?,0) as field_992540,");//-- (二)逾期放款(A)-內部融資
       sqlCmd.append("        round((field_992510+field_992520+field_992530+field_992540)/?,0) as field_2_3_sum,");// --逾期放款-合計
       sqlCmd.append("        round(field_OVER/?,0) as field_OVER_1,");//-- (二)逾期放款--各類總額
       sqlCmd.append("        decode(field_992140,0,0,round(field_992510 / field_992140 *100 ,2)) as   field_992510_992140_RATE,");//-- (二)逾放比率-正會員
       sqlCmd.append("        decode(field_990410,0,0,round(field_992520 / field_990410 *100 ,2)) as   field_992520_990410_RATE,");//-- (二)逾放比率-贊助會員
       sqlCmd.append("        decode((field_990610-field_990611),0,0,round(field_992530 / (field_990610-field_990611) *100 ,2)) as   field_992530_990610_cal_RATE,");//-- (二)逾放比率-非會員
       sqlCmd.append("        decode(field_120700,0,0,round(field_992540 / field_120700 *100 ,2)) as   field_992540_120700_RATE,");//-- (二)逾放比率-內部融資
       sqlCmd.append("        decode((field_992140+field_990410+(field_990610-field_990611)+field_120700),0,0,round((field_992510+field_992520+field_992530+field_992540)/ (field_992140+field_990410+(field_990610-field_990611)+field_120700) *100 ,2))  as   field_2_4_sum,");//--逾放比率-合計
       sqlCmd.append("        decode(a01.field_CREDIT,0,0,round(a01.field_OVER /  a01.field_CREDIT *100 ,2))  as   field_OVER_RATE_1,");//-- (二)逾放比率--各類總額
       sqlCmd.append("        round(field_992610/?,0) as field_992610,");//-- (二)應予觀察放款(B)-正會員
       sqlCmd.append("        round(field_992620/?,0) as field_992620,");//-- (二)應予觀察放款(B)-贊助會員
       sqlCmd.append("        round(field_992630/?,0) as field_992630,");//-- (二)應予觀察放款(B)-非會員
       sqlCmd.append("        round(field_992640/?,0) as field_992640,");//-- (二)應予觀察放款(B)-內部融資
       sqlCmd.append("        round((field_992610+field_992620+field_992630+field_992640)/ ?,0) as field_2_5_sum,");//--應予觀察放款(B)-合計
       sqlCmd.append("        round((field_992610+field_992620+field_992630+field_992640+field_992650)/ ?,0) as field_992610_cal,");//-- (二)應予觀察放款(B)-各類總額
       sqlCmd.append("        round((field_992510+field_992610)/?,0) as field_992510_992610,");//-- (二)廣義逾放(A+B)-正會員
       sqlCmd.append("        round((field_992520+field_992620)/?,0) as field_992520_992620,");//-- (二)廣義逾放(A+B)-贊助會員
       sqlCmd.append("        round((field_992530+field_992630)/?,0) as field_992530_992630,");//-- (二)廣義逾放(A+B)-非會員
       sqlCmd.append("        round((field_992540+field_992640)/?,0) as field_992540_992640,");//-- (二)廣義逾放(A+B)-內部融資
       sqlCmd.append("        round(((field_992510+field_992610)+(field_992520+field_992620)+(field_992530+field_992630)+(field_992540+field_992640))/ ?,0) as field_2_6_sum,");//--廣義逾放(A+B)-合計
       sqlCmd.append("        round((field_OVER+field_992610+field_992620+field_992630+field_992640+field_992650)/?,0) as field_OVER_992610_cal,");//-- (二)廣義逾放(A+B)--各類總額
       sqlCmd.append("        decode(field_992140,0,0,round((field_992510+field_992610) /  field_992140 *100 ,2)) as   field_AB_992140_RATE,");//-- (二)廣義逾放比-正會員
       sqlCmd.append("        decode(field_990410,0,0,round((field_992520+field_992620) /  field_990410 *100 ,2)) as   field_AB_990410_RATE,");//-- (二)廣義逾放比-贊助會員
       sqlCmd.append("        decode((field_990610-field_990611),0,0,round((field_992530+field_992630) /  (field_990610-field_990611) *100 ,2)) as   field_AB_990610_RATE,");//-- (二)廣義逾放比-非會員
       sqlCmd.append("        decode(field_120700,0,0,round((field_992540+field_992640) /  field_120700 *100 ,2)) as   field_AB_120700_RATE,");//-- (二)廣義逾放比-內部融資
       sqlCmd.append("        decode((field_992140+field_990410+ (field_990610-field_990611)+field_120700),0,0,round(((field_992510+field_992610)+(field_992520+field_992620)+(field_992530+field_992630)+(field_992540+field_992640) ) / (field_992140+field_990410+ (field_990610-field_990611)+field_120700) *100 ,2)) as  field_2_7_sum,");//--廣義逾放比-合計
       sqlCmd.append("        decode(field_CREDIT,0,0,round((field_OVER+field_992610+field_992620+field_992630+field_992640+field_992650) / field_CREDIT *100 ,2)) as field_AB_CREDIT_RATE,");//-- (二)廣義逾放比-各類總額
       sqlCmd.append("        round(field_110600/?,0) as field_110600,");//-- (三)投資業務.有價證券
       sqlCmd.append("        round(field_130200/?,0) as field_130200,");//-- (三)投資業務.長期投資
       sqlCmd.append("        round(field_130100/?,0) as field_130100,");//-- (三)投資業務.專案基金 103.06.16 add
       sqlCmd.append("        round(field_910401/?,0) as field_910401,");//-- (三)長期投資.全國農業金庫股票
       sqlCmd.append("        round(field_910403/?,0) as field_910403,");//-- (三)長期投資.合作金庫股票
       sqlCmd.append("        round(field_910402/?,0) as field_910402,");//-- (三)長期投資.財金資訊(股)公司股票
       sqlCmd.append("        round((field_130200-field_910401-field_910403-field_910402+field_130100)/?,0) as field_130200_other,");//--(三)長期投資.其他,增加(含基金) 103.06.16 add
       sqlCmd.append("        round(field_loan_bal_amt/?,0) as field_loan_bal_amt,");//--專案農貸業務.專案農貸餘額[(7)貸放餘額]
       sqlCmd.append("        round(field_over6m_loan_bal_amt/?,0) as field_over6m_loan_bal_amt,");//--專案農貸業務.逾期放款[(9)逾期六個月放款餘額-金額]-->103.03.18 add
       sqlCmd.append("        decode(field_loan_bal_amt,0,0,round(field_over6m_loan_bal_amt /  field_loan_bal_amt *100 ,2))  as   field_OVER_RATE_argi,");//--專案農貸業務.逾放比率[(9)逾期六個月放款餘額-金額/(7)貸放餘額-金額]-->103.03.18 add
       sqlCmd.append("        round(field_delay_loan_cnt_sum/1,0) as field_delay_loan_cnt_sum,");//--專案農貸業務.當年度累計核准延期還款件數[(10)當月核准延期還款-件數]-->103.03.18 add
       sqlCmd.append("        round(field_delay_loan_amt_sum/?,0) as field_delay_loan_amt_sum");//--專案農貸業務.當年度累計核准延期還款金額[(11)當月核准延期還款-金額]-->103.03.18 add
       for(int i=1;i<=51;i++){
           paramList.add(unit);
       }
       sqlCmd.append(" from  ");
       sqlCmd.append(" ( ");
       sqlCmd.append("   select a01.bank_no,  a01.BANK_NAME,");
       sqlCmd.append("          SUM(field_190000_cal) field_190000_cal ,");
       sqlCmd.append("          SUM(field_DEBIT)  field_DEBIT ,");
       sqlCmd.append("          SUM(field_CREDIT) field_CREDIT,");        
       sqlCmd.append("          SUM(field_320300) field_320300,");        
       sqlCmd.append("          SUM(field_320300_last) field_320300_last,");  
       sqlCmd.append("          SUM(field_NET) field_NET,");
       sqlCmd.append("          SUM(field_OVER) field_OVER,");
       sqlCmd.append("          SUM(field_BACKUP) field_BACKUP,");       
       sqlCmd.append("          SUM(field_120700) field_120700,");  
       sqlCmd.append("          SUM(field_NOASSURE) field_NOASSURE,");
       sqlCmd.append("          SUM(field_110600) field_110600,");
       sqlCmd.append("          SUM(field_130200) field_130200,");
       sqlCmd.append("          SUM(field_130100) field_130100,");//103.06.16 add
       sqlCmd.append("          SUM(field_990310) field_990310,"); 
       sqlCmd.append("          SUM(field_990410) field_990410,");
       sqlCmd.append("          SUM(field_990420) field_990420,");
       sqlCmd.append("          SUM(field_990510) field_990510,");
       sqlCmd.append("          SUM(field_990610) field_990610,");
       sqlCmd.append("          SUM(field_990611) field_990611,");
       sqlCmd.append("          SUM(field_990624) field_990624,");//110.03.05 add
       sqlCmd.append("          SUM(field_990624_1_rule) field_990624_1_rule,");//110.03.05 add
       sqlCmd.append("          SUM(field_990624_2_rule) field_990624_2_rule,");//110.03.05 add
       sqlCmd.append("          SUM(field_990626) field_990626,");//110.03.05 add
       sqlCmd.append("          SUM(field_990626_rule) field_990626_rule,");//110.03.05 add
       sqlCmd.append("          SUM(field_CAPTIAL) field_CAPTIAL,");
       sqlCmd.append("          SUM(field_910401) field_910401,");
       sqlCmd.append("          SUM(field_910402) field_910402,");
       sqlCmd.append("          SUM(field_910403) field_910403,");
       sqlCmd.append("          SUM(field_992130) field_992130,");
       sqlCmd.append("          SUM(field_992140) field_992140,");
       sqlCmd.append("          SUM(field_992510) field_992510,");
       sqlCmd.append("          SUM(field_992520) field_992520,");
       sqlCmd.append("          SUM(field_992530) field_992530,");
       sqlCmd.append("          SUM(field_992540) field_992540,");
       sqlCmd.append("          SUM(field_992610) field_992610,");
       sqlCmd.append("          SUM(field_992620) field_992620,");
       sqlCmd.append("          SUM(field_992630) field_992630,");
       sqlCmd.append("          SUM(field_992640) field_992640,");
       sqlCmd.append("          SUM(field_992650) field_992650,");
       sqlCmd.append("          SUM(field_992710) field_992710,");
       sqlCmd.append("          SUM(field_1) field_1,");
       sqlCmd.append("          SUM(field_2_1) field_2_1,");
       sqlCmd.append("          SUM(field_2_2) field_2_2,");
       sqlCmd.append("          SUM(field_3) field_3,");
       sqlCmd.append("          SUM(field_4) field_4,");
       sqlCmd.append("          SUM(field_5) field_5,");
       sqlCmd.append("          SUM(field_6) field_6,");
       sqlCmd.append("          SUM(field_7) field_7,");
       sqlCmd.append("          SUM(field_8) field_8,");
       sqlCmd.append("          SUM(field_9) field_9,");
       sqlCmd.append("          SUM(field_10) field_10,");
       sqlCmd.append("          SUM(field_1_violate) as field_1_violate,"); //102.08.26各項法定比率有違反的增加顯示***
       sqlCmd.append("          SUM(field_2_1_violate) as field_2_1_violate,");
       sqlCmd.append("          SUM(field_2_2_violate) as field_2_2_violate,");
       sqlCmd.append("          SUM(field_3_violate) as field_3_violate,");
       sqlCmd.append("          SUM(field_4_violate) as field_4_violate,");
       sqlCmd.append("          SUM(field_5_violate) as field_5_violate,");
       sqlCmd.append("          SUM(field_6_violate) as field_6_violate,");
       sqlCmd.append("          SUM(field_7_violate) as field_7_violate,");
       sqlCmd.append("          SUM(field_8_violate) as field_8_violate,");
       sqlCmd.append("          SUM(field_9_violate) as field_9_violate,");
       sqlCmd.append("          SUM(field_10_violate) as field_10_violate,");
       sqlCmd.append("          SUM(field_loan_bal_amt) as field_loan_bal_amt,");//--專案農貸餘額:(7)貸放餘額 
       sqlCmd.append("          SUM(field_over6m_loan_bal_amt) as field_over6m_loan_bal_amt,");//--逾期放款:(9)逾期六個月放款餘額-金額     
       sqlCmd.append("          SUM(field_delay_loan_cnt_sum) as field_delay_loan_cnt_sum,");//--當年度累計核准延期還款件數:(10)當月核准延期還款-件數         
       sqlCmd.append("          SUM(field_delay_loan_amt_sum) as field_delay_loan_amt_sum ");//--當年度累計核准延期還款金額:(11)當月核准延期還款-金額
       sqlCmd.append(" from ");         
       sqlCmd.append(" (  select  bn01.bank_no , bn01.BANK_NAME,");           
       sqlCmd.append("            round(sum(decode(bn01.bank_type,'6',decode(a01.acc_code,'190000',amt,0),'7',decode(a01.acc_code,'100000',amt,0),0)) /1,0)     as field_190000_cal,");//--資產總額
       sqlCmd.append("            round(sum(decode(a01.acc_code,'220000',amt,0)) /1,0) as field_DEBIT,");//--存款
       sqlCmd.append("            round(sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /1,0) as  field_CREDIT,");//--放款
       sqlCmd.append("            round(sum(decode(a01.acc_code,'320300',amt,0)) /1,0) as field_320300,");//--本期損益
       sqlCmd.append("            round(sum(decode(bn01.bank_type,'6',decode(a01.acc_code,'310000',amt,'320000',amt,0),'7',decode(a01.acc_code,'300000',amt,0),0)) /1,0)     as field_NET,");//--淨值
       sqlCmd.append("            round(sum(decode(a01.acc_code,'990000',amt,0)) /1,0) as field_OVER,");//--逾期放款
       sqlCmd.append("            round(sum(decode(a01.acc_code, '120800',amt,'150300',amt,0)) /1,0) as  field_BACKUP,");//--備抵呆帳
       sqlCmd.append("            round(sum(decode(a01.acc_code,'120700',amt,0)) /1,0) as field_120700,");//--內部融資   
       sqlCmd.append("            round(sum( decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code, '120101',amt,'120301',amt, '120401',amt, '120501',amt,0),'7',decode(a01.acc_code, '120101',amt,'120401', amt, '120201',amt, '120501',amt,0)),");  
       sqlCmd.append("                               -                           '103',decode(a01.acc_code, '120101',amt,'120301',amt, '120401',amt, '120501',amt,0),0) ) /1,0) as  field_NOASSURE,");  
       sqlCmd.append("            round(sum(decode(a01.acc_code,'110600',amt,0)) /1,0) as field_110600,");//--有價證券 
       sqlCmd.append("            round(sum(decode(a01.acc_code,'130200',amt,0)) /1,0) as field_130200,");//--長期投資
       sqlCmd.append("            round(sum(decode(a01.acc_code,'130100',amt,0)) /1,0) as field_130100 ");//--專案基金 103.06.16 add
       sqlCmd.append("     from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
       sqlCmd.append("     left join (select (CASE WHEN (a01.m_year <= 102) THEN '102' ");                             
       sqlCmd.append("                             WHEN (a01.m_year > 102) THEN '103'  ");                            
       sqlCmd.append("                             ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01 ");
       sqlCmd.append("                 where m_year=? and m_month=? ");               
       sqlCmd.append("                ) a01  on  bn01.bank_no = a01.bank_code ");
       sqlCmd.append("    where a01.bank_code=? ");
       paramList.add(wlx01_m_year);
       paramList.add(s_year);
       paramList.add(s_month);
       paramList.add(bank_no);
       sqlCmd.append("    group by a01.m_year,a01.m_month,bn01.bank_no,bn01.BANK_NAME ");
       sqlCmd.append(" )a01,");
       sqlCmd.append("  (  select bn01.bank_no as bank_code, bn01.BANK_NAME,");
       sqlCmd.append("            round(sum(decode(acc_code,'990310',amt,0)) /1,0) as field_990310,");
       sqlCmd.append("            round(sum(decode(acc_code,'990410',amt,0)) /1,0) as field_990410,");
       sqlCmd.append("            round(sum(decode(acc_code,'990420',amt,0)) /1,0) as field_990420,");
       sqlCmd.append("            round(sum(decode(acc_code,'990510',amt,0)) /1,0) as field_990510,");
       sqlCmd.append("            round(sum(decode(acc_code,'990610',amt,0)) /1,0) as field_990610,");
       sqlCmd.append("            round(sum(decode(acc_code,'990611',amt,0)) /1,0) as field_990611,");
       sqlCmd.append("            round(sum(decode(acc_code,'990624',amt,0)) /1,0) as field_990624, ");//110.03.05 add
       sqlCmd.append("            round(sum(decode(acc_code,'990626',amt,0)) /1,0) as field_990626 ");//110.03.05 add
       sqlCmd.append("       from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
       sqlCmd.append("       left join (select * from a02 where m_year=? and m_month=? )a02 on bn01.bank_no = a02.bank_code ");
       sqlCmd.append("      where bank_code=? ");
       paramList.add(wlx01_m_year);
       paramList.add(s_year);
       paramList.add(s_month);
       paramList.add(bank_no);
       sqlCmd.append("      group by bn01.bank_no,bn01.BANK_NAME  ");
       sqlCmd.append("   ) a02,");
       sqlCmd.append(" (  select bn01.bank_no as bank_code,  bn01.bank_name,");
       sqlCmd.append("           round(sum(decode(a05.acc_code,'91060P',amt,0)) /1,0) as field_CAPTIAL, ");
       sqlCmd.append("           round(sum(decode(a05.acc_code,'910401',amt,0)) /1,0) as field_910401,");//103.06.16 fix
       sqlCmd.append("           round(sum(decode(a05.acc_code,'910402',amt,0)) /1,0) as field_910402,");//103.06.16 fix
       sqlCmd.append("           round(sum(decode(a05.acc_code,'910403',amt,0)) /1,0) as field_910403 ");//103.06.16 fix  
       sqlCmd.append("      from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
       sqlCmd.append("      left join (select * from a05 where m_year=? and m_month=? ) a05 on  bn01.bank_no = a05.bank_code ");
       sqlCmd.append("     where a05.bank_code=? ");
       paramList.add(wlx01_m_year);
       paramList.add(s_year);
       paramList.add(s_month);
       paramList.add(bank_no);
       sqlCmd.append("     group by bn01.bank_no,bn01.BANK_NAME ");
       sqlCmd.append(" ) a05,");
       sqlCmd.append("  (   select bn01.bank_no as bank_code, bn01.BANK_NAME,");
         sqlCmd.append("           round(sum(decode(a99.acc_code,'992130',amt,0)) /1,0) as field_992130,");
       sqlCmd.append("             round(sum(decode(a99.acc_code,'992140',amt,0)) /1,0) as field_992140,");
       sqlCmd.append("             round(sum(decode(a99.acc_code,'992510',amt,0)) /1,0) as field_992510,");
       sqlCmd.append("             round(sum(decode(a99.acc_code,'992520',amt,0)) /1,0) as field_992520,");
       sqlCmd.append("             round(sum(decode(a99.acc_code,'992530',amt,0)) /1,0) as field_992530,");
       sqlCmd.append("             round(sum(decode(a99.acc_code,'992540',amt,0)) /1,0) as field_992540,");
       sqlCmd.append("             round(sum(decode(a99.acc_code,'992610',amt,0)) /1,0) as field_992610,");
       sqlCmd.append("             round(sum(decode(a99.acc_code,'992620',amt,0)) /1,0) as field_992620,");
       sqlCmd.append("             round(sum(decode(a99.acc_code,'992630',amt,0)) /1,0) as field_992630,");
       sqlCmd.append("             round(sum(decode(a99.acc_code,'992640',amt,0)) /1,0) as field_992640,");
       sqlCmd.append("             round(sum(decode(a99.acc_code,'992650',amt,0)) /1,0) as field_992650,");
       sqlCmd.append("             round(sum(decode(a99.acc_code,'992710',amt,0)) /1,0) as field_992710 ");
       sqlCmd.append("       from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
       sqlCmd.append("       left join (select * from a99 where m_year = ?  and m_month = ? )a99 on bn01.bank_no = a99.bank_code ");
       sqlCmd.append("      where bank_code=? ");
       paramList.add(wlx01_m_year);
       paramList.add(s_year);
       paramList.add(s_month);
       paramList.add(bank_no);
       sqlCmd.append("      group by bn01.bank_no,bn01.BANK_NAME ");
       sqlCmd.append("   ) a99,");
       sqlCmd.append(" (   select bank_code,");
       sqlCmd.append("            sum(decode(a02_operation.acc_code,'field_month_dc_rate',amt,0)) as field_1,");
       sqlCmd.append("            sum(decode(a02_operation.acc_code,'field_month_dc_rate',decode(violate,'Y',1,0),0)) as field_1_violate,");
       sqlCmd.append("            sum(decode(a02_operation.acc_code,'field_990210/(990230-990240)',amt,0)) as field_2_1,");
       sqlCmd.append("            sum(decode(a02_operation.acc_code,'field_990210/(990230-990240)',decode(violate,'Y',1,0),0)) as field_2_1_violate,");
       sqlCmd.append("            sum(decode(a02_operation.acc_code,'field_990220/(990230-990240)',amt,0)) as field_2_2,");
       sqlCmd.append("            sum(decode(a02_operation.acc_code,'field_990220/(990230-990240)',decode(violate,'Y',1,0),0)) as field_2_2_violate,");
       sqlCmd.append("            sum(decode(a02_operation.acc_code,'field_f/g',amt,0)) as field_3,");
       sqlCmd.append("            sum(decode(a02_operation.acc_code,'field_f/g',decode(violate,'Y',1,0),0)) as field_3_violate,");
       sqlCmd.append("            sum(decode(a02_operation.acc_code,'field_990410/990420',amt,0)) as field_4,");
       sqlCmd.append("            sum(decode(a02_operation.acc_code,'field_990410/990420',decode(violate,'Y',1,0),0))  as field_4_violate,");
       sqlCmd.append("            sum(decode(a02_operation.acc_code,'field_990512/990320',amt,0)) as field_5,");
       sqlCmd.append("            sum(decode(a02_operation.acc_code,'field_990512/990320',decode(violate,'Y',1,0),0))  as field_5_violate,");
       sqlCmd.append("            sum(decode(a02_operation.acc_code,'field_k/990620',amt,0)) as field_6,");
       sqlCmd.append("            sum(decode(a02_operation.acc_code,'field_k/990620',decode(violate,'Y',1,0),0))  as field_6_violate,");
       //sqlCmd.append("            sum(decode(a02_operation.acc_code,'field_990710/990720',amt,0)) as field_7,");
       //sqlCmd.append("            sum(decode(a02_operation.acc_code,'field_990710/990720',decode(violate,'Y',1,0),0))  as field_7_violate,");
       sqlCmd.append("            CASE WHEN (m_year * 100 + m_month >= 10810) THEN sum(decode(a02_operation.acc_code,'field_990711_990712/fieldi_y',amt,0))");//108.09.16 fix 
       sqlCmd.append("            ELSE sum(decode(a02_operation.acc_code,'field_990710/990720',amt,0)) END  as field_7,");//108.09.16 fix
       sqlCmd.append("            CASE WHEN (m_year * 100 + m_month >= 10810) THEN sum(decode(a02_operation.acc_code,'field_990711_990712/fieldi_y',decode(violate,'Y',1,0),0))");//108.09.16 fix
       sqlCmd.append("            ELSE sum(decode(a02_operation.acc_code,'field_990710/990720',decode(violate,'Y',1,0),0)) END  as field_7_violate,");//108.09.16 fix       
       sqlCmd.append("            sum(decode(a02_operation.acc_code,'field_990810/(990230-990240)',amt,0)) as field_8,");
       sqlCmd.append("            sum(decode(a02_operation.acc_code,'field_990810/(990230-990240)',decode(violate,'Y',1,0),0))  as field_8_violate,");
       sqlCmd.append("            sum(decode(a02_operation.acc_code,'field_|990910-990920|/990230',amt,0)) as field_9,");
       sqlCmd.append("            sum(decode(a02_operation.acc_code,'field_|990910-990920|/990230',decode(violate,'Y',1,0),0))  as field_9_violate,");
       sqlCmd.append("            sum(decode(a02_operation.acc_code,'field_991020/990320',amt,0)) as field_10,");
       sqlCmd.append("            sum(decode(a02_operation.acc_code,'field_991020/990320',decode(violate,'Y',1,0),0))  as field_10_violate, ");
       sqlCmd.append("            sum(decode(a02_operation.acc_code,'990624_1_rule',decode(violate,'Y',1,0),0))  as field_990624_1_rule, ");
       sqlCmd.append("            sum(decode(a02_operation.acc_code,'990624_2_rule',decode(violate,'Y',1,0),0))  as field_990624_2_rule, ");
       sqlCmd.append("            sum(decode(a02_operation.acc_code,'990626_rule',decode(violate,'Y',1,0),0))  as field_990626_rule ");
       sqlCmd.append("       from  a02_operation ");
       sqlCmd.append("      where m_year = ?  and m_month=? "); 
       sqlCmd.append("        and bank_code= ? ");
       paramList.add(s_year);
       paramList.add(s_month);
       paramList.add(bank_no);
       sqlCmd.append("      group by m_year,m_month,bank_code ");//108.09.16 fix         
       sqlCmd.append("  ) a02_operation,");
       sqlCmd.append("  (select  bn01.bank_no as bank_code , bn01.BANK_NAME,");           
       sqlCmd.append("           round(sum(decode(a01.acc_code,'320300',amt,0)) /1,0) as field_320300_last ");//--上年度損益          
       sqlCmd.append("     from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
       sqlCmd.append("     left join (select (CASE WHEN (a01.m_year <= 102) THEN '102' ");                              
       sqlCmd.append("                             WHEN (a01.m_year > 102) THEN '103'  ");                             
       sqlCmd.append("                             ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01 ");
       sqlCmd.append("                 where m_year= ? and m_month=? "); 
       sqlCmd.append("             ) a01  on  bn01.bank_no = a01.bank_code ");
       sqlCmd.append("    where a01.bank_code= ? "); 
       paramList.add(wlx01_last_m_year);
       paramList.add(Integer.parseInt(s_year)-1);
       paramList.add("12");
       paramList.add(bank_no);
       sqlCmd.append("   group by a01.m_year,a01.m_month,bn01.bank_no,bn01.BANK_NAME ");
       sqlCmd.append(" )a01_last,");
       sqlCmd.append(" (select agri_loan.bank_no ,   agri_loan.BANK_NAME,");             
       sqlCmd.append("         SUM(field_loan_bal_amt)   as field_loan_bal_amt,");//--專案農貸餘額:(7)貸放餘額 
       sqlCmd.append("         SUM(field_over6m_loan_bal_amt) as field_over6m_loan_bal_amt,");//--逾期放款:(9)逾期六個月放款餘額-金額     
       sqlCmd.append("         SUM(field_delay_loan_cnt_sum)   as field_delay_loan_cnt_sum,");//--當年度累計核准延期還款件數:(10)當月核准延期還款-件數         
       sqlCmd.append("         SUM(field_delay_loan_amt_sum)   as field_delay_loan_amt_sum ");//--當年度累計核准延期還款金額:(11)當月核准延期還款-金額                
       sqlCmd.append("    from (select bn01.bank_no , bn01.BANK_NAME,");
       sqlCmd.append("                 round(sum(agri_loan.loan_bal_amt) /1,0) as field_loan_bal_amt,");//--(7)貸放餘額
       sqlCmd.append("                 round(sum(agri_loan.over6m_loan_bal_amt) /1,0) as field_over6m_loan_bal_amt ");//--(9)逾期六個月放款餘額-金額          
       sqlCmd.append("            from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
       sqlCmd.append("            left join (select * from agri_loan  where m_year=? and m_month=?) agri_loan  on  bn01.bank_no = agri_loan.bank_code ");
       sqlCmd.append("           where bank_code=? ");
       paramList.add(wlx01_last_m_year);
       paramList.add(s_year);
       paramList.add(s_month);
       paramList.add(bank_no);
       sqlCmd.append("           group by bn01.bank_no,bn01.BANK_NAME ");
       sqlCmd.append("          ) agri_loan,");
       sqlCmd.append("          (select bn01.bank_no , bn01.BANK_NAME,");        
       sqlCmd.append("                  sum(agri_loan.delay_loan_cnt) as field_delay_loan_cnt_sum,");//--(10)當月核准延期還款-件數        
       sqlCmd.append("                  sum(agri_loan.delay_loan_amt) as field_delay_loan_amt_sum ");//--(11)當月核准延期還款-金額        
       sqlCmd.append("            from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
       sqlCmd.append("            left join (select * from agri_loan  where (to_char(m_year * 100 + m_month) >= ? and to_char(m_year * 100 + m_month) <= ?) ");       
       sqlCmd.append("                      ) agri_loan  on  bn01.bank_no = agri_loan.bank_code ");
       sqlCmd.append("           where bank_code=? ");
       paramList.add(wlx01_last_m_year);
       paramList.add(s_year+"01");
       paramList.add(s_year+s_month);
       paramList.add(bank_no);
       sqlCmd.append("           group by bn01.bank_no,bn01.BANK_NAME ");//--.1月~當月累計 
       sqlCmd.append("          )agri_loan_sum ");
       sqlCmd.append("    where agri_loan.bank_no=agri_loan_sum.bank_no(+) ");
       sqlCmd.append("      and agri_loan.bank_no <> ' ' "); 
       sqlCmd.append("    GROUP BY agri_loan.bank_no,agri_loan.BANK_NAME ");
       sqlCmd.append(" )agri_loan ");
       sqlCmd.append(" where a01.bank_no = a02.bank_code(+) and a01.bank_no = a05.bank_code(+)  and a01.bank_no=a99.bank_code(+) ");
       sqlCmd.append("   and a01.bank_no =a02_operation.bank_code(+) and a01.bank_no = a01_last.bank_code(+) ");
       sqlCmd.append("   and a01.bank_no=agri_loan.bank_no(+) ");
       sqlCmd.append("   and a01.bank_no=? ");
       paramList.add(bank_no);
       sqlCmd.append(" GROUP BY a01.bank_no,a01.BANK_NAME ");  
       sqlCmd.append(" )a01 ");
       List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"bank_no,bank_name,field_190000_cal,field_debit,field_credit,field_320300,field_320300_last,"
                                                                            +"field_net,field_over_rate,field_over,field_backup,field_backup_credit_rate,field_backup_over_rate,field_captial_rate,"
                                                                            +"field_1,field_1_violate,field_2_1,field_2_1_violate,field_2_2,field_2_2_violate,field_3,field_3_violate,field_4,field_4_violate,field_5,field_5_violate,"
                                                                            +"field_6,field_6_violate,field_7,field_7_violate,field_8,field_8_violate,field_9,field_9_violate,field_10,field_10_violate,field_11,"
                                                                            +"field_debit_1,field_992130,field_990420,field_990310,field_992140,field_990410,field_990610_990611,"
                                                                            +"field_120700,field_990611,field_2_1_sum,field_credit_1,field_990510,field_2_2_sum,field_noassure,field_992710,field_992510,"
                                                                            +"field_992520,field_992530,field_992540,field_2_3_sum,field_over_1,field_992510_992140_rate,field_992520_990410_rate,"
                                                                            +"field_992530_990610_cal_rate,field_992540_120700_rate,field_2_4_sum,field_over_rate_1,field_992610,field_992620,"
                                                                            +"field_992630,field_992640,field_2_5_sum,field_992610_cal,field_992510_992610,field_992520_992620,field_992530_992630,"
                                                                            +"field_992540_992640,field_2_6_sum,field_over_992610_cal,field_ab_992140_rate,field_ab_990410_rate,field_ab_990610_rate,"
                                                                            +"field_ab_120700_rate,field_2_7_sum,field_ab_credit_rate,field_110600,field_130200,field_130100,field_910401,field_910403,field_910402,field_130200_other,"
                                                                            +"field_loan_bal_amt,field_over6m_loan_bal_amt,field_over_rate_argi,field_delay_loan_cnt_sum,field_delay_loan_amt_sum,"
                                                                            +"field_990624,field_4_1_violate,field_4_2_violate");
       System.out.println("dbData_mainInfo.size()="+dbData.size()); 
       return dbData;
   }
   //--(一)所屬之資訊共用中心
   public static List getInfoCenterInfo(String tbank_no,String wlx01_m_year){
       StringBuffer sqlCmd = new StringBuffer();  
       List paramList = new ArrayList();
       sqlCmd.append(" select wlx01.bank_no,");
       sqlCmd.append("        ba01.bank_name ");//--所屬之資訊共用中心
       sqlCmd.append("   from (select * from wlx01 where m_year=?)wlx01 ");
       sqlCmd.append("   left join (select * from ba01 where m_year=? and bank_type=?)ba01 ");
       sqlCmd.append("     on wlx01.center_no = ba01.bank_no ");
       sqlCmd.append("  where wlx01.bank_no=? ");
       paramList.add(wlx01_m_year);
       paramList.add(wlx01_m_year);
       paramList.add("8");
       paramList.add(tbank_no);
       List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"bank_no,bank_name");
       System.out.println("dbData_getInfoCenterInfo.size()="+dbData.size()); 
       return dbData;
   }
   //--(二)稽核人員
   public static List getAuditInfo(String tbank_no,String wlx01_m_year){
       StringBuffer sqlCmd = new StringBuffer();  
       List paramList = new ArrayList();
       sqlCmd.append(" select wlx01.bank_no,");
       sqlCmd.append("        name as audit_name");//--稽核人員
       sqlCmd.append("   from (select * from wlx01 where m_year=?)wlx01 ");
       sqlCmd.append("   left join wlx01_audit on wlx01.bank_no = wlx01_audit.bank_no ");
       sqlCmd.append("  where wlx01.bank_no=? order by seq_no ");
       paramList.add(wlx01_m_year);
       paramList.add(tbank_no);
       List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"bank_no,audit_name");
       System.out.println("dbData_getAuditInfo.size()="+dbData.size()); 
       return dbData;
   }
   //--(三)最近3次金融檢查(報告編號及基準日)—只顯示最近3筆資料.中間用空白隔開 ex: 102C102(101.7.31)  100C136(100.8.31)  099C067(99.4.30)
   public static List getBaseDate(String tbank_no){
       StringBuffer sqlCmd  = new StringBuffer(); 
       List paramList = new ArrayList();
       sqlCmd.append(" select bank_no,reportno || '(' || ((TO_CHAR(base_date,'yyyy')-1911)||'/'|| TO_CHAR(base_date,'mm/dd')) || ')'  base_date ");//--報告編號及基準日
       sqlCmd.append("   from exreportf ");
       sqlCmd.append("  where bank_no=? ");
       sqlCmd.append("  order by base_date desc ");
       paramList.add(tbank_no);
       List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"bank_no,base_date");
       System.out.println("dbData_getBaseDate.size()="+dbData.size()); 
       return dbData;
   }
   //--(四)主要負責人一覽表(由左到右)
   public static List getChargePerson(String tbank_no){
       StringBuffer sqlCmd = new StringBuffer(); 
       List paramList = new ArrayList();
       sqlCmd.append(" select cmuse_name,");//--職別
       sqlCmd.append("        name ");//--姓名
       sqlCmd.append("   from WLX01_M,cdshareno where bank_no=? ");
       sqlCmd.append("    and WLX01_M.POSITION_CODE = cdshareno.CMUSE_ID and cdshareno.CMUSE_DIV='005' ");
       sqlCmd.append("    and abdicate_code !='Y' ");
       sqlCmd.append("  order by position_code,abdicate_date ");
       paramList.add(tbank_no);
       List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"cmuse_name,name");
       System.out.println("dbData_getChargePerson.size()="+dbData.size()); 
       return dbData;
   }
   //--(五)組織及人員配置
   public static List getOrgInfo(String s_year,String s_month,String tbank_no,String wlx01_m_year){
       StringBuffer sqlCmd = new StringBuffer(); 
       List paramList = new ArrayList();
       sqlCmd.append(" select ba01.bank_no,");
       sqlCmd.append("        ba01.bank_name,");//--名稱
       sqlCmd.append("        credit_staff_num,");//--信用部員工x人
       sqlCmd.append("        staff_num,");//--員工人數
       sqlCmd.append("        addr,");//--地址
       sqlCmd.append("        trim(telno) telno,"); //--電話
       sqlCmd.append("        cmuse_name,"); //--職稱
       sqlCmd.append("        name ");//--姓名
       sqlCmd.append("   from (select * from wlx01 where m_year=?)wlx01 ");
       sqlCmd.append("   left join (select * from ba01 where m_year=? and bank_type in ('6','7') and bank_kind='0')ba01 on wlx01.bank_no = ba01.bank_no ");
       sqlCmd.append("   left join ( select bank_no,cmuse_name,name from WLX01_M,cdshareno where WLX01_M.POSITION_CODE = cdshareno.CMUSE_ID and cdshareno.CMUSE_DIV='005' ");
       sqlCmd.append("    and abdicate_code !='Y' and position_code='4')wlx01_m on wlx01.bank_no = wlx01_m.bank_no ");
       sqlCmd.append("  where wlx01.bank_no=? ");
       sqlCmd.append("    and cancel_no != 'Y' ");
       paramList.add(wlx01_m_year);
       paramList.add(wlx01_m_year);
       paramList.add(tbank_no);
       sqlCmd.append(" union ");
       sqlCmd.append(" select ba01.bank_no,ba01.bank_name,");
       sqlCmd.append("        0,");
       sqlCmd.append("        staff_num,");
       sqlCmd.append("        addr,");
       sqlCmd.append("        trim(telno) telno,");
       sqlCmd.append("        cmuse_name,name ");
       sqlCmd.append("   from (select * from wlx02 where m_year=?)wlx02  ");
       sqlCmd.append("   left join (select * from ba01 where m_year=? and bank_type in ('6','7') and bank_kind='1')ba01 on wlx02.tbank_no = ba01.pbank_no and wlx02.bank_no=ba01.bank_no ");
       sqlCmd.append("   left join ( select bank_no,cmuse_name,name from WLX02_M,cdshareno where WLX02_M.POSITION_CODE = cdshareno.CMUSE_ID and cdshareno.CMUSE_DIV='007' ");
       sqlCmd.append("    and abdicate_code !='Y' and position_code='1')wlx02_m on wlx02.bank_no = wlx02_m.bank_no  ");
       sqlCmd.append("  where wlx02.tbank_no=?  ");
       sqlCmd.append("    and cancel_no != 'Y' ");
       paramList.add(wlx01_m_year);
       paramList.add(wlx01_m_year);
       paramList.add(tbank_no);
       sqlCmd.append("  order by bank_no asc ");
       List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"bank_no,bank_name,credit_staff_num,staff_num,addr,telno,cmuse_name,name");
       System.out.println("dbData_getOrgInfo.size()="+dbData.size()); 
       return dbData;
   }
   //--(六)列入警示項目
   public static List getRemarkInfo(String s_year,String s_month,String unit,String bank_no){
       StringBuffer sqlCmd = new StringBuffer(); 
       List paramList = new ArrayList();
       //sqlCmd.append(" select bank_code,wr_range.remark || ':' || decode(type,4,amt,round(amt/?,0)) as remark ");//--警示項目
       sqlCmd.append(" select t1.bank_code,t1.wr_rpt,t1.type,t1.amt,t2.remark ");
       sqlCmd.append("   from wr_operation t1 ");
       sqlCmd.append("   left join wr_range t2 on t1.wr_rpt = t2.wr_rpt and t1.wr_range_serial = t2.serial ");
       sqlCmd.append("  where t1.m_year=? ");
       sqlCmd.append("    and t1.m_month=? ");
       sqlCmd.append("    and t1.bank_code=? ");
       sqlCmd.append("    and t1.warn_type='Y' ");
       sqlCmd.append("    and t1.wr_range_serial is not null");
       sqlCmd.append("  order by t1.wr_rpt,t1.wr_range_serial "); 
       paramList.add(s_year);
       paramList.add(s_month);
       paramList.add(bank_no);
       List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"bank_code,wr_rpt,type,amt,remark");
       System.out.println("dbData_getRemarkInfo.size()="+dbData.size()); 
       return dbData;
   }
   //套用原excel欄位設定
   private static void insertCell(String value,HSSFWorkbook wb,HSSFRow row,int i){
           HSSFCell cell=(row.getCell((short)i)==null)? row.createCell((short)i) : row.getCell((short)i);
           //HSSFCellStyle cs1 = wb.createCellStyle();110.03.05 fix
           HSSFCellStyle cs1 = cell.getCellStyle();//會套用原本excel所設定的格式
           //設置邊框
           /*
           cs1.setBorderTop(topBorder); //上邊框
           cs1.setBorderBottom(bottomBorder); //下邊框
           cs1.setBorderLeft(leftBorder); //左邊框
           cs1.setBorderRight(rightBorder); //右邊框
           cs1.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);//垂直置中
           cs1.setAlignment(alignment);//水平置中
           cs1.setWrapText(warptext);//自動換行
           cell.setCellStyle(cs1);
           cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
           */
           cell.setCellValue(value);
   }
   //設定欄位格式
   private static void insertCell1(String value,HSSFWorkbook wb,HSSFRow row,int i,short topBorder,short bottomBorder,short leftBorder,short rightBorder, short alignment,boolean warptext){
       HSSFCell cell=(row.getCell((short)i)==null)? row.createCell((short)i) : row.getCell((short)i);
       HSSFCellStyle cs1 = wb.createCellStyle();
       //HSSFCellStyle cs1 = cell.getCellStyle();//會套用原本excel所設定的格式
       //設置邊框       
       cs1.setBorderTop(topBorder); //上邊框
       cs1.setBorderBottom(bottomBorder); //下邊框
       cs1.setBorderLeft(leftBorder); //左邊框
       cs1.setBorderRight(rightBorder); //右邊框
       cs1.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);//垂直置中
       cs1.setAlignment(alignment);//水平置中
       cs1.setWrapText(warptext);//自動換行
       cell.setCellStyle(cs1);
       cell.setEncoding( HSSFCell.ENCODING_UTF_16 );       
       cell.setCellValue(value);
}
   //取得總機構代碼
   private static String getBank_name(String bank_no,String m_year){
       String bank_name = "";
       StringBuffer sqlCmd = new StringBuffer(); 
       List paramList = new ArrayList();
       sqlCmd.append(" select bank_name from bn01 where bank_no = ? and m_year = ? ");
       paramList.add(bank_no);
       paramList.add(m_year);
       List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"bank_name");
       if(dbData.size()>0){
           bank_name = Utility.getTrimString(((DataObject)dbData.get(0)).getValue("bank_name"));
       }
       return bank_name;
   } 
   	//違反農金法及子法需遭處分
   	public static List getViolateInfo(String tbank_no,String wlx01_m_year){
	   	StringBuffer sqlCmd = new StringBuffer();  
	   	List paramList = new ArrayList();
	    sqlCmd.append(" select bn01.bank_no,bn01.bank_name,"); 
		sqlCmd.append(" ((TO_CHAR(violate_date,'yyyy')-1911)||'/'|| TO_CHAR(violate_date,'mm/dd')) as violate_date,");//--受處分日期
		sqlCmd.append(" TO_CHAR(violate_date,'yyyy/mm/dd') as violate_date_1,");
		sqlCmd.append(" violate_type,");//--處分方式  
		sqlCmd.append(" title,");//--主旨
		sqlCmd.append(" content,");//--說明
		sqlCmd.append(" law_content ");//--法令依據
		sqlCmd.append(" from mis_violatelaw "); 
		sqlCmd.append(" left join (select * from bn01 where m_year= ? )bn01 on mis_violatelaw.bank_no = bn01.bank_no "); 
		sqlCmd.append(" where  mis_violatelaw.bank_no= ? ");
		sqlCmd.append(" ORDER BY violate_date ");
		paramList.add(wlx01_m_year);
		paramList.add(tbank_no);
		List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"bank_no,bank_name,violate_date,violate_date_1,violate_type,title,content,law_content");
		System.out.println("dbData_getViolateInfo.size()="+dbData.size()); 
		return dbData;
   	}
   
   	//限制或核准業務函令
	public static List getLoalInfo(String tbank_no,String wlx01_m_year){
		StringBuffer sqlCmd = new StringBuffer();  
		List paramList = new ArrayList();
		sqlCmd.append(" select mis_loal.bank_no,ba01.bank_name,");
		sqlCmd.append(" loal_number,");//--限制函號
		sqlCmd.append(" loal_content,");//--限制內容
		sqlCmd.append(" loal_states,");
		sqlCmd.append(" select_name,");//--狀態
		sqlCmd.append(" loal_ps,");//--備註
		sqlCmd.append(" loal_no "); 
		sqlCmd.append(" from mis_loal ");
		sqlCmd.append(" left join (select * from ba01 where m_year=?)ba01 on mis_loal.bank_no=ba01.bank_no ");
		sqlCmd.append(" left join (select * from mis_select where select_id='LOAL_STATES')mis_select on mis_loal.loal_states = mis_select.select_num ");
		sqlCmd.append(" where mis_loal.bank_no=? "); 
		sqlCmd.append(" order by loal_add_date desc ");
		paramList.add(wlx01_m_year);
		paramList.add(tbank_no);
		List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"bank_no,bank_name,loal_number,loal_content,loal_states,select_name,loal_ps,loal_no");
		System.out.println("dbData_getLoalInfo.size()="+dbData.size()); 
		return dbData;
	}
	//舞幣案件
	public static List getEmInfo(String tbank_no,String wlx01_m_year){
		StringBuffer sqlCmd = new StringBuffer();  
	    List paramList = new ArrayList();
		sqlCmd.append(" select em_number,");//--來文文號
		sqlCmd.append(" ((TO_CHAR(em_date,'yyyy')-1911)||'/'|| TO_CHAR(em_date,'mm/dd')) as em_date,");//--來文日期 
		sqlCmd.append(" em_emer,");//--舞弊人
		sqlCmd.append(" mis_em.bank_no,ba01.bank_name,");
		sqlCmd.append(" em_content,");//--舞幣內容
		sqlCmd.append(" em_no ");
		sqlCmd.append(" from mis_em left join (select * from ba01 where m_year=? )ba01 on mis_em.bank_no=ba01.bank_no ");
		sqlCmd.append(" where mis_em.bank_no = ? "); 
		sqlCmd.append(" order by em_date desc ");
		paramList.add(wlx01_m_year);
		paramList.add(tbank_no);
		List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"em_number,em_date,em_emer,bank_no,bank_name,em_content,em_no");
		System.out.println("dbData_getEmInfo.size()="+dbData.size()); 
		return dbData;
	}
	
	
	public static List getViolate_typeList(){
		StringBuffer sqlCmd = new StringBuffer();  
	    List paramList = new ArrayList();
	    sqlCmd.append(" select cmuse_id,cmuse_name from cdshareno where cmuse_div=? order by input_order"); 
		paramList.add("038");		
		List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"cmuse_id,cmuse_name");
		System.out.println("dbData_getViolate_typeList.size()="+dbData.size()); 
		return dbData;
	}
	
	//取得4.2-縣市別
	public static List getCityInfo(String s_year,String s_month,String bank_no){
			StringBuffer sqlCmd = new StringBuffer();  
		    List paramList = new ArrayList();
		    
		    sqlCmd.append(" select a1.amt_name ||  decode(NVL(a2.amt_name,''),'','','、'||a2.amt_name ) || decode(NVL(a3.amt_name,''),'','','、'||a3.amt_name ) as amt_name");
		    sqlCmd.append(" from");
		    sqlCmd.append(" (");
		    sqlCmd.append(" select  NVL(cd01.hsien_name||cd02.area_name,'') as amt_name");
		    sqlCmd.append(" from a02"); 
		    sqlCmd.append(" left join (select hsien_id,hsien_name");
		    sqlCmd.append(" from cd01");
		    sqlCmd.append(" order by fr001w_output_order)cd01 on cd01.hsien_id =  substr(amt_name1,1,1)");
		    sqlCmd.append(" left join (select area_id,area_name,hsien_id from cd02)cd02 on cd02.hsien_id=substr(amt_name1,1,1) and cd02.area_id = substr(amt_name1,3,3)"); 
		    sqlCmd.append(" where m_year=? and m_month=? and bank_code=?"); 
		    sqlCmd.append(" and acc_code in (?)"); 
		    sqlCmd.append(" and amt in (2,4)");
		    paramList.add(s_year);
			paramList.add(s_month);
			paramList.add(bank_no);
			paramList.add("990626");
		    sqlCmd.append(" )a1,");
		    sqlCmd.append(" (");
		    sqlCmd.append(" select  NVL(cd01.hsien_name||cd02.area_name,'') as amt_name");
		    sqlCmd.append(" from a02"); 
		    sqlCmd.append(" left join (select hsien_id,hsien_name");
		    sqlCmd.append(" from cd01");
		    sqlCmd.append(" order by fr001w_output_order)cd01 on cd01.hsien_id =  substr(amt_name2,1,1)");
		    sqlCmd.append(" left join (select area_id,area_name,hsien_id from cd02)cd02 on cd02.hsien_id=substr(amt_name2,1,1) and cd02.area_id = substr(amt_name2,3,3)"); 
		    sqlCmd.append(" where m_year=? and m_month=? and bank_code=?"); 
		    sqlCmd.append(" and acc_code in (?)"); 
		    sqlCmd.append(" and amt in (2,4)");
		    paramList.add(s_year);
			paramList.add(s_month);
			paramList.add(bank_no);
			paramList.add("990626");
		    sqlCmd.append(" )a2,");
		    sqlCmd.append(" (");
		    sqlCmd.append(" select NVL(hsien_name,'') as amt_name");
		    sqlCmd.append(" from a02 left join (select hsien_id,hsien_name");
		    sqlCmd.append(" from cd01");
		    sqlCmd.append(" order by fr001w_output_order)cd01 on cd01.hsien_id = amt_name");
		    sqlCmd.append(" where m_year=? and m_month=? and bank_code=?"); 
		    sqlCmd.append(" and acc_code in (?)");
		    sqlCmd.append(" and amt in (2)");
		    sqlCmd.append(" )a3");
		    paramList.add(s_year);
			paramList.add(s_month);
			paramList.add(bank_no);
			paramList.add("990626");
			List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
			System.out.println("dbData_getCityInfo.size()="+dbData.size()); 
			return dbData;
		}
 }

