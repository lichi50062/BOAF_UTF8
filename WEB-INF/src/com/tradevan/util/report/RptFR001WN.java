/*
 * Created N年內及目前月份經營概況趨勢統計分析 by 2968
 * 103.12.24 fix 桃園縣升格調整 by 2968
 * 106.03.07 fix 調整原台灣省改為其他(包含台灣省及福建省.中華民國農會),福建省合併至其他 by 2295  
 */
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.Region;

import java.awt.Font;
import java.io.*;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

import com.sun.corba.se.impl.javax.rmi.CORBA.Util;
import com.tradevan.util.dao.DataObject;
import com.tradevan.util.*; 

public class RptFR001WN {
	static String hsien_id = "";
	static String hsien_name = "";  
	static String hsien_english = "";//縣市英文名稱
	static String fr001w_output_order = "";
	static String bank_no = "";
	static String bank_code = "";
	static String bank_name = "";   
	static String bank_english = "";//機構英文名稱
	static String count_seq = ""; 
	static String field_seq = "";
	static String field_debit = ""; //存款總額      
	static String field_credit = "";//放款總額
	static String field_net = "";//淨值
	static String field_320300 = "";//本期損益
	static String field_over = ""; //狹義逾期放款   
	static String field_840740 = "";//廣義逾期放款     
	static String field_over_rate = "";//狹義逾放比率(狹義逾期放款/放款總額)    
	static String field_840740_rate = "";//廣義逾放比率(廣義逾期放款/放款總額)	
	static String field_backup = "";//備抵呆帳
	static String field_backup_over_rate = "";//備呆占狹義逾期放款比率(備抵呆帳/狹義逾放)
	static String field_backup_840740_rate = "";//備呆占廣義逾期放款比率(備抵呆帳/廣義逾放)
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
    public static String createRpt(String s_year,String s_month,String unit,String bank_type,String s_year0){
          String errMsg = "";
          String bank_type_name="";
          String unit_name = Utility.getUnitName(unit);
          String data_year="";
          String data_month="";
          DataObject bean_Title = null;
          DataObject bean_Sum = null;
          DataObject bean_Fukien = null;
          DataObject bean_Taiwan = null;
          List qList_Title =  new ArrayList();
          List qList_Sum =  new ArrayList();
          List qList_Fukien =  new ArrayList();
          List qList_Taiwan =  new ArrayList();
          reportUtil reportUtil = new reportUtil();
          FileOutputStream fileOut = null;          
          HSSFRow row=null;//宣告一列
          HSSFCell cell=null;//宣告一個儲存格
          System.out.println("RptFR001WN.bank_type="+bank_type);
          if(bank_type.equals("ALL")){
             bank_type_name = "農漁會";
          }else{
             bank_type_name = (bank_type.equals("6"))?"農會":"漁會";
          }          
          //99.09.16 add 查詢年度100年以前.縣市別不同===============================
          String cd01_table = (Integer.parseInt(s_year) < 100)?"cd01_99":"cd01"; 
          String wlx01_m_year = (Integer.parseInt(s_year) < 100)?"99":"100"; 
          //===================================================================== 
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
            String titleName=s_year0+"年內及目前月份經營概況趨勢統計分析";
            //Creating Cells
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.createSheet( "report" ); //建立sheet，及名稱
            HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定  
            //設定頁面符合列印大小          
            sheet.setAutobreaks( false );
            ps.setScale( ( short )100 ); //列印縮放百分比
            ps.setPaperSize( ( short )8 ); //設定紙張大小 A3            
            ps.setLandscape( true ); // 設定橫印           
            HSSFFooter footer = sheet.getFooter();            
            //設定樣式和位置(請精減style物件的使用量，以免style物件太多excel報表無法開啟)
            HSSFCellStyle leftStyle = reportUtil.getLeftStyle(wb);
            HSSFCellStyle rightStyle = reportUtil.getRightStyle(wb);
            /*HSSFCellStyle defaultStyle = reportUtil.getDefaultStyle(wb);//有框內文置中
            HSSFCellStyle noBorderDefaultStyle = reportUtil.getNoBorderDefaultStyle(wb);//無框內文置中
            reportUtil.setDefaultStyle(defaultStyle);
            reportUtil.setNoBorderDefaultStyle(noBorderDefaultStyle);*/           
	  		if(sheet==null){System.out.println("open sheet 失敗");}else System.out.println("open sheet 成功");
	  		footer.setCenter( "Page:" + HSSFFooter.page() + " of " +
                    HSSFFooter.numPages() );                                        
	  		footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa")); 
           //報表欄位=======================================================================
           int cellNum = 0;
           insertCell("單位代號\r\n Code",wb,sheet.createRow( ( short )3 ),cellNum,HSSFCellStyle.BORDER_THIN);
           insertCell("",wb,sheet.createRow( ( short )4 ),cellNum,HSSFCellStyle.BORDER_THIN);
           insertCell("",wb,sheet.createRow( ( short )5 ),cellNum,HSSFCellStyle.BORDER_THIN);
           sheet.addMergedRegion( new Region( ( short )3, ( short )cellNum, ( short )5, ( short )cellNum ) );
           cellNum++;
           
           insertCell("單位名稱\r\n Name",wb,sheet.createRow( ( short )3 ),cellNum,HSSFCellStyle.BORDER_THIN);
           insertCell("",wb,sheet.createRow( ( short )4 ),cellNum,HSSFCellStyle.BORDER_THIN);
           insertCell("",wb,sheet.createRow( ( short )5 ),cellNum,HSSFCellStyle.BORDER_THIN);
           sheet.addMergedRegion( new Region( ( short )3, ( short )cellNum, ( short )5, ( short )cellNum ) );
           cellNum++;
           
           insertCell("存款總額\r\n Deposits",wb,sheet.createRow( ( short )3 ),cellNum,HSSFCellStyle.BORDER_THIN);
           insertCell("",wb,sheet.createRow( ( short )4 ),cellNum,HSSFCellStyle.BORDER_THIN);
           int startCell = cellNum;
           for(int n=Integer.parseInt(s_year)-Integer.parseInt(s_year0);n<=Integer.parseInt(s_year);n++){
               if(n!=Integer.parseInt(s_year)-Integer.parseInt(s_year0)){
                   row = sheet.createRow( ( short )3 );
                   insertCell("",wb,row,cellNum,HSSFCellStyle.BORDER_THIN);
               }
               row = sheet.createRow( ( short )5 );
               if(n==Integer.parseInt(s_year)){
                   insertCell(s_year+"/"+s_month,wb,row,cellNum,HSSFCellStyle.BORDER_THIN);
               }else{
                   insertCell(n+"年底",wb,row,cellNum,HSSFCellStyle.BORDER_THIN);
               }
               cellNum++;
           }
           int endCell = cellNum-1;
           sheet.addMergedRegion( new Region( ( short )3, ( short )startCell, ( short )4, ( short )endCell ) );
           
           insertCell("放款總額\r\n Loans",wb,sheet.createRow( ( short )3 ),cellNum,HSSFCellStyle.BORDER_THIN);
           insertCell("",wb,sheet.createRow( ( short )4 ),cellNum,HSSFCellStyle.BORDER_THIN);
           startCell = cellNum;
           for(int n=Integer.parseInt(s_year)-Integer.parseInt(s_year0);n<=Integer.parseInt(s_year);n++){
               if(n!=Integer.parseInt(s_year)-Integer.parseInt(s_year0)){
                   row = sheet.createRow( ( short )3 );
                   insertCell("",wb,row,cellNum,HSSFCellStyle.BORDER_THIN);
               }
               row = sheet.createRow( ( short )5 );
               if(n==Integer.parseInt(s_year)){
                   insertCell(s_year+"/"+s_month,wb,row,cellNum,HSSFCellStyle.BORDER_THIN);
               }else{
                   insertCell(n+"年底",wb,row,cellNum,HSSFCellStyle.BORDER_THIN);
               }
               cellNum++;
           }
           endCell = cellNum-1;
           sheet.addMergedRegion( new Region( ( short )3, ( short )startCell, ( short )4, ( short )endCell ) );
           
           insertCell("淨值\r\n Net Worth",wb,sheet.createRow( ( short )3 ),cellNum,HSSFCellStyle.BORDER_THIN);
           insertCell("",wb,sheet.createRow( ( short )4 ),cellNum,HSSFCellStyle.BORDER_THIN);
           startCell = cellNum;
           for(int n=Integer.parseInt(s_year)-Integer.parseInt(s_year0);n<=Integer.parseInt(s_year);n++){
               if(n!=Integer.parseInt(s_year)-Integer.parseInt(s_year0)){
                   row = sheet.createRow( ( short )3 );
                   insertCell("",wb,row,cellNum,HSSFCellStyle.BORDER_THIN);
               }
               row = sheet.createRow( ( short )5 );
               if(n==Integer.parseInt(s_year)){
                   insertCell(s_year+"/"+s_month,wb,row,cellNum,HSSFCellStyle.BORDER_THIN);
               }else{
                   insertCell(n+"年底",wb,row,cellNum,HSSFCellStyle.BORDER_THIN);
               }
               cellNum++;
           }
           endCell = cellNum-1;
           sheet.addMergedRegion( new Region( ( short )3, ( short )startCell, ( short )4, ( short )endCell ) );
           
           insertCell("本期損益\r\n Profit Before Tax",wb,sheet.createRow( ( short )3 ),cellNum,HSSFCellStyle.BORDER_THIN);
           insertCell("",wb,sheet.createRow( ( short )4 ),cellNum,HSSFCellStyle.BORDER_THIN);
           startCell = cellNum;
           for(int n=Integer.parseInt(s_year)-Integer.parseInt(s_year0);n<=Integer.parseInt(s_year);n++){
               if(n!=Integer.parseInt(s_year)-Integer.parseInt(s_year0)){
                   row = sheet.createRow( ( short )3 );
                   insertCell("",wb,row,cellNum,HSSFCellStyle.BORDER_THIN);
               }
               row = sheet.createRow( ( short )5 );
               if(n==Integer.parseInt(s_year)){
                   insertCell(s_year+"/"+s_month,wb,row,cellNum,HSSFCellStyle.BORDER_THIN);
               }else{
                   insertCell(n+"年底",wb,row,cellNum,HSSFCellStyle.BORDER_THIN);
               }
               cellNum++;
           }
           endCell = cellNum-1;
           sheet.addMergedRegion( new Region( ( short )3, ( short )startCell, ( short )4, ( short )endCell ) );
           
           row = sheet.createRow( ( short )4 );
           startCell = cellNum;
           endCell = startCell+Integer.parseInt(s_year0);
           insertCell("狹義逾期放款\r\n NPL(Old)",wb,row,startCell,HSSFCellStyle.BORDER_THIN);
           sheet.addMergedRegion( new Region( ( short )4, ( short )startCell, ( short )4, ( short )endCell ) );
           startCell = endCell+1;
           endCell = startCell+Integer.parseInt(s_year0);
           insertCell("廣義逾期放款\r\n NPL(New)",wb,row,startCell,HSSFCellStyle.BORDER_THIN);
           sheet.addMergedRegion( new Region( ( short )4, ( short )startCell, ( short )4, ( short )endCell ) );
           startCell = endCell+1;
           endCell = startCell+Integer.parseInt(s_year0);
           insertCell("狹義逾期比率\r\n NPL Ratio(Old)",wb,row,startCell,HSSFCellStyle.BORDER_THIN);
           sheet.addMergedRegion( new Region( ( short )4, ( short )startCell, ( short )4, ( short )endCell ) );
           startCell = endCell+1;
           endCell = startCell+Integer.parseInt(s_year0);
           insertCell("廣義逾期比率\r\n NPL Ratio(New)",wb,row,startCell,HSSFCellStyle.BORDER_THIN);
           sheet.addMergedRegion( new Region( ( short )4, ( short )startCell, ( short )4, ( short )endCell ) );
           row = sheet.createRow( ( short )3 );
           insertCell("逾期放款\r\n Non-performing Loan(NPL)",wb,row,cellNum,HSSFCellStyle.BORDER_THIN);
           sheet.addMergedRegion( new Region( ( short )3, ( short )cellNum, ( short )3, ( short )endCell ) );
           for(int j=0;j<4;j++){
               for(int n=Integer.parseInt(s_year)-Integer.parseInt(s_year0);n<=Integer.parseInt(s_year);n++){
                   if(j!=0|| (j==0 && n!=Integer.parseInt(s_year)-Integer.parseInt(s_year0))){
                       row = sheet.createRow( ( short )3 );
                       insertCell("",wb,row,cellNum,HSSFCellStyle.BORDER_THIN);
                   }
                   row = sheet.createRow( ( short )5 );
                   if(n==Integer.parseInt(s_year)){
                       insertCell(s_year+"/"+s_month,wb,row,cellNum,HSSFCellStyle.BORDER_THIN);
                   }else{
                       insertCell(n+"年底",wb,row,cellNum,HSSFCellStyle.BORDER_THIN);
                   }
                   cellNum++;
               }
           }
           row = sheet.createRow( ( short )4 );
           startCell = cellNum;
           endCell = startCell+Integer.parseInt(s_year0);
           insertCell("金額\r\n Amount",wb,row,startCell,HSSFCellStyle.BORDER_THIN);
           sheet.addMergedRegion( new Region( ( short )4, ( short )startCell, ( short )4, ( short )endCell ) );
           startCell = endCell+1;
           endCell = startCell+Integer.parseInt(s_year0);
           insertCell("占狹義逾期放款比率\r\n Coverage Rate",wb,row,startCell,HSSFCellStyle.BORDER_THIN);
           sheet.addMergedRegion( new Region( ( short )4, ( short )startCell, ( short )4, ( short )endCell ) );
           startCell = endCell+1;
           endCell = startCell+Integer.parseInt(s_year0);
           insertCell("占廣義逾期放款比率\r\n Coverage Rate(New)",wb,row,startCell,HSSFCellStyle.BORDER_THIN);
           sheet.addMergedRegion( new Region( ( short )4, ( short )startCell, ( short )4, ( short )endCell ) );
           row = sheet.createRow( ( short )3 );
           insertCell("備抵呆帳\r\n Loan Loss Deposit",wb,row,cellNum,HSSFCellStyle.BORDER_THIN);
           sheet.addMergedRegion( new Region( ( short )3, ( short )cellNum, ( short )3, ( short )endCell ) );
           for(int j=0;j<3;j++){
               for(int n=Integer.parseInt(s_year)-Integer.parseInt(s_year0);n<=Integer.parseInt(s_year);n++){
                   if(j!=0 || (j==0 && n!=Integer.parseInt(s_year)-Integer.parseInt(s_year0))){
                       row = sheet.createRow( ( short )3 );
                       insertCell("",wb,row,cellNum,HSSFCellStyle.BORDER_THIN);
                   }
                   row = sheet.createRow( ( short )5 );
                   if(n==Integer.parseInt(s_year)){
                       insertCell(s_year+"/"+s_month,wb,row,cellNum,HSSFCellStyle.BORDER_THIN);
                   }else{
                       insertCell(n+"年底",wb,row,cellNum,HSSFCellStyle.BORDER_THIN);
                   }
                   cellNum++;
               }
           }
           insertCell("淨值佔風險性資產比率\r\n BIS",wb,sheet.createRow( ( short )3 ),cellNum,HSSFCellStyle.BORDER_THIN);
           insertCell("",wb,sheet.createRow( ( short )4 ),cellNum,HSSFCellStyle.BORDER_THIN);
           startCell = cellNum;
           for(int n=Integer.parseInt(s_year)-Integer.parseInt(s_year0);n<=Integer.parseInt(s_year);n++){
               if(n!=Integer.parseInt(s_year)-Integer.parseInt(s_year0)){
                   insertCell("",wb,sheet.createRow( ( short )3 ),cellNum,HSSFCellStyle.BORDER_THIN);
                   insertCell("",wb,sheet.createRow( ( short )4 ),cellNum,HSSFCellStyle.BORDER_THIN);
               }
               row = sheet.createRow( ( short )5 );
               if(n==Integer.parseInt(s_year)){
                   insertCell(s_year+"/"+s_month,wb,row,cellNum,HSSFCellStyle.BORDER_THIN);
               }else{
                   insertCell(n+"年底",wb,row,cellNum,HSSFCellStyle.BORDER_THIN);
               }
               cellNum++;
           }
           endCell = cellNum-1;
           sheet.addMergedRegion( new Region( ( short )3, ( short )startCell, ( short )4, ( short )endCell ) );
           
           //設定表頭===============================================================================
           row = sheet.createRow( ( short )0 );
           //insertCell("",wb,row,(short)0,HSSFCellStyle.BORDER_NONE ,(short)18);
           cell = row.createCell((short)0);
           HSSFCellStyle style = wb.createCellStyle();
           HSSFFont font = wb.createFont();
           font.setFontHeightInPoints((short) 18);
           style.setAlignment(HSSFCellStyle.ALIGN_CENTER);//水平置中
           style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);//垂直置中
           style.setFont(font);
           cell.setCellStyle(style);
           cell.setEncoding(HSSFCell.ENCODING_UTF_16);
           cell.setCellValue(bank_type_name+titleName);
           sheet.addMergedRegion( new Region( ( short )0, ( short )0,
                                              ( short )0,
                                              ( short )endCell) );
           //列印單位=======================================================================================            
           row = sheet.createRow( ( short )2 );
           int unitLengthstr = endCell-2;
           insertCell("單位：新台幣"+Utility.getUnitName(unit)+"、％",wb,row,unitLengthstr,HSSFCellStyle.BORDER_NONE);
           sheet.addMergedRegion( new Region( ( short )2, ( short )unitLengthstr,
                                              ( short )2,
                                              ( short )endCell) );
           //寫入dbData
           cellNum = 2; //資料起始cell
           for(int n=Integer.parseInt(s_year)-Integer.parseInt(s_year0);n<=Integer.parseInt(s_year);n++){
               if(n == Integer.parseInt(s_year)){
                   //qList_Title = qList_Title(s_year,s_month,unit,bank_type,cd01_table,wlx01_m_year);
                   qList_Sum = qList_Sum(s_year,s_month,unit,bank_type,cd01_table,wlx01_m_year);
                   //qList_Fukien = qList_Fukien(s_year,s_month,unit,bank_type,cd01_table,wlx01_m_year);//106.03.07 合併至台灣省
                   qList_Taiwan = qList_Taiwan(s_year,s_month,unit,bank_type,cd01_table,wlx01_m_year);
               }else{
                   //qList_Title = qList_Title(String.valueOf(n),"12",unit,bank_type,cd01_table,wlx01_m_year);
                   qList_Sum = qList_Sum(String.valueOf(n),"12",unit,bank_type,cd01_table,wlx01_m_year);
                   //qList_Fukien = qList_Fukien(String.valueOf(n),"12",unit,bank_type,cd01_table,wlx01_m_year);//106.03.07 合併至台灣省
                   qList_Taiwan = qList_Taiwan(String.valueOf(n),"12",unit,bank_type,cd01_table,wlx01_m_year);
               }
               /*if(n==Integer.parseInt(s_year)-Integer.parseInt(s_year0)){
                   if(qList_Title!=null && qList_Title.size()>0){
                       for(int q=0;q<qList_Title.size();q++){
                           row = sheet.createRow( ( short )5+q );
                           bean_Title = (DataObject)qList_Title.get(q);
                           if(q!=0){
                               if("".equals(Utility.getTrimString(bean_Title.getValue("bank_code")))){
                                   insertCell("",wb,row,(short)0,HSSFCellStyle.BORDER_THIN ,(short)10);
                                   insertCell(bean_Title.getValue("hsien_name").toString(),wb,row,(short)1,HSSFCellStyle.BORDER_THIN ,(short)10);
                               }else{
                                   insertCell(bean_Title.getValue("bank_code").toString(),wb,row,(short)0,HSSFCellStyle.BORDER_THIN ,(short)10);
                                   insertCell(bean_Title.getValue("bank_name").toString(),wb,row,(short)1,HSSFCellStyle.BORDER_THIN ,(short)10);
                               }
                           }
                           if(q==qList_Title.size()-1){
                               bean_Title = (DataObject)qList_Title.get(0);
                               row = sheet.createRow( ( short )5+q+1 );
                               insertCell("",wb,row,(short)0,HSSFCellStyle.BORDER_THIN ,(short)10);
                               insertCell(bean_Title.getValue("hsien_name").toString(),wb,row,(short)1,HSSFCellStyle.BORDER_THIN ,(short)10);
                           }
                       }
                   }
               }*/
               int tmpRow = 5+qList_Sum.size()+1+2;
               if(qList_Sum!=null && qList_Sum.size()>0){
                   for(int q=0;q<qList_Sum.size();q++){
                       //N年前的第一年，印bank_name;
                       if(n==Integer.parseInt(s_year)-Integer.parseInt(s_year0)){
                           if(q!=0){
                               row = sheet.createRow( ( short )5+q );
                               bean_Sum = (DataObject)qList_Sum.get(q);
                               if("".equals(Utility.getTrimString(bean_Sum.getValue("bank_no")))){
                                   insertCell("",wb,row,(short)0,HSSFCellStyle.BORDER_THIN);
                                   insertCell("",wb,row,(short)1,HSSFCellStyle.BORDER_THIN);
                                   cell = row.getCell((short)1);
                                   cell.setCellValue(bean_Sum.getValue("hsien_name").toString());
                                   cell.setCellStyle(leftStyle);
                               }else{
                                   insertCell(bean_Sum.getValue("bank_no").toString(),wb,row,(short)0,HSSFCellStyle.BORDER_THIN);
                                   insertCell("",wb,row,(short)1,HSSFCellStyle.BORDER_THIN);
                                   cell = row.getCell((short)1);
                                   cell.setCellValue(bean_Sum.getValue("bank_name").toString());
                                   cell.setCellStyle(leftStyle);
                               }
                           }
                           if(q==qList_Sum.size()-1){
                               bean_Sum = (DataObject)qList_Sum.get(0);
                               row = sheet.createRow( ( short )5+q+1 );
                               insertCell("",wb,row,(short)0,HSSFCellStyle.BORDER_THIN);
                               insertCell("",wb,row,(short)1,HSSFCellStyle.BORDER_THIN);
                               cell = row.getCell((short)1);
                               cell.setCellValue(bean_Sum.getValue("hsien_name").toString());
                               cell.setCellStyle(leftStyle);
                           }
                       }
                       //所有年度========================
                       int prtTimes = 1;
                       if(q!=0){
                           bean_Sum = (DataObject)qList_Sum.get(q);
                           if(Integer.parseInt(s_year) <= 99){  
                               if( ("A90".equals(bean_Sum.getValue("field_seq").toString()) || "A99".equals(bean_Sum.getValue("field_seq").toString())) 
                                       && ("台北市".equals(bean_Sum.getValue("hsien_name").toString()) || "高雄市".equals(bean_Sum.getValue("hsien_name").toString()) ) ){      
                                   prtTimes = 2;       
                               }//end of 台北市.高雄市
                           }else if(Integer.parseInt(s_year) >= 100){//99.09.23 100年(含)以後.明細表加印新北市/台中市/台南市
                               if( ("A90".equals(bean_Sum.getValue("field_seq").toString()) || "A99".equals(bean_Sum.getValue("field_seq").toString())) 
                                    &&("新北市".equals(bean_Sum.getValue("hsien_name").toString()) || "臺北市".equals(bean_Sum.getValue("hsien_name").toString())
                                    || "桃園市".equals(bean_Sum.getValue("hsien_name").toString()) || "臺中市".equals(bean_Sum.getValue("hsien_name").toString()) 
                                    || "臺南市".equals(bean_Sum.getValue("hsien_name").toString()) || "高雄市".equals(bean_Sum.getValue("hsien_name").toString()) ) ){ 
                                   prtTimes = 2;       
                               }//end of 新北市.台北市.桃園市.台中市.台南市.高雄市
                           }
                           for(int i=1;i<=prtTimes;i++){
                               if(i==1){
                                   row = sheet.createRow( ( short )5+q );
                               }else{
                                   row = sheet.createRow( ( short )tmpRow );
                                   tmpRow++;
                                   insertCell("",wb,row,(short)0,HSSFCellStyle.BORDER_THIN);
                                   insertCell("",wb,row,(short)1,HSSFCellStyle.BORDER_THIN);
                                   cell = row.getCell((short)1);
                                   cell.setCellValue(bean_Sum.getValue("hsien_name").toString());
                                   cell.setCellStyle(leftStyle);
                               }
                               int tmpC = cellNum;
                               //System.out.println("***qList_Sum");
                               setInsertValue(wb,row,cell,s_year0,bean_Sum,tmpC,rightStyle);
                               tmpC = 2;
                           }
                           
                       }
                       if(q==qList_Sum.size()-1){
                           bean_Sum = (DataObject)qList_Sum.get(0);
                           for(int i=1;i<=2;i++){
                               if(i==1){ //第一個總計
                                   row = sheet.createRow( ( short )5+q+1 );
                               }else{ //最下面的總計
                                   row = sheet.createRow( ( short )tmpRow+1 );//106.03.07
                                   insertCell("",wb,row,(short)0,HSSFCellStyle.BORDER_THIN);
                                   insertCell("",wb,row,(short)1,HSSFCellStyle.BORDER_THIN);
                                   cell = row.getCell((short)1);
                                   cell.setCellValue(bean_Sum.getValue("hsien_name").toString());
                                   cell.setCellStyle(leftStyle);
                               }
                               int tmpC = cellNum;
                               //System.out.println("***qList_Sum");
                               setInsertValue(wb,row,cell,s_year0,bean_Sum,tmpC,rightStyle);
                               tmpC = 2;
                           }
                       }
                   }
                   /*106.03.07 福建省合併至台灣省
                   if(qList_Fukien!=null && qList_Fukien.size()>0){
                      // for(int q=0;q<qList_Fukien.size();q++){
                           bean_Fukien = (DataObject)qList_Fukien.get(0);
                           row = sheet.createRow( ( short )tmpRow );
                           insertCell("",wb,row,(short)0,HSSFCellStyle.BORDER_THIN);
                           insertCell("",wb,row,(short)1,HSSFCellStyle.BORDER_THIN);
                           cell = row.getCell((short)1);
                           cell.setCellValue(bean_Fukien.getValue("hsien_name").toString());
                           cell.setCellStyle(leftStyle);
                           int tmpC = cellNum;
                           //System.out.println("***qList_Fukien");
                           setInsertValue(wb,row,cell,s_year0,bean_Fukien,tmpC,rightStyle);
                           tmpC = 2;
                       //}
                   }
                   */
                   if(qList_Taiwan!=null && qList_Taiwan.size()>0){
                           bean_Taiwan = (DataObject)qList_Taiwan.get(0);
                           row = sheet.createRow( ( short )tmpRow );//106.03.07
                           insertCell("",wb,row,(short)0,HSSFCellStyle.BORDER_THIN);
                           insertCell("",wb,row,(short)1,HSSFCellStyle.BORDER_THIN);
                           cell = row.getCell((short)1);
                           cell.setCellValue(bean_Taiwan.getValue("hsien_name").toString());
                           cell.setCellStyle(leftStyle);
                           int tmpC = cellNum;
                           //System.out.println("***qList_Taiwan");
                           setInsertValue(wb,row,cell,s_year0,bean_Taiwan,tmpC,rightStyle);
                           tmpC = 2;
                       //}
                   }
               }
               
               cellNum++;
               
           }
           //設定欄位寬高============================
           row = sheet.createRow( ( short )0 );
           row.setHeightInPoints((short)40);
           row = sheet.createRow( ( short )3 );
           row.setHeightInPoints((short)40);
           row = sheet.createRow( ( short )4 );
           row.setHeightInPoints((short)40);
           sheet.setColumnWidth((short)1,(short)7000);
           for(int k=2;k<=(Integer.parseInt(s_year0)+1)*12+1;k++){
               sheet.setColumnWidth((short)k,(short)4000);
           }
           
           // Write the output to a file============================   
           fileOut = new FileOutputStream( Utility.getProperties("reportDir")+System.getProperty("file.separator")+"N年內及目前月份經營概況趨勢統計分析.xls" );
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
   //顯示各縣市別/總計/機構名稱title
   public static List qList_Title(String s_year,String s_month,String unit,String bank_type,String cd01_table,String wlx01_m_year){
       StringBuffer sqlCmd = new StringBuffer(); 
       List paramList = new ArrayList();
       sqlCmd.append("select ' ' AS hsien_id, ' 總   計 ' AS hsien_name, 'A99' as field_SEQ, '001' AS FR001W_output_order, '' as bank_code, '' as bank_name from dual ");
       sqlCmd.append("union all ");
       sqlCmd.append("select nvl(cd01.hsien_id,' ') as hsien_id, nvl(cd01.hsien_name,'OTHER') as hsien_name, 'A90' as field_SEQ, ");           
       sqlCmd.append("       cd01.FR001W_output_order as FR001W_output_order, '' as bank_code, '' as bank_name ");
       sqlCmd.append("  from (select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y') cd01 ");            
       sqlCmd.append("union all ");     
       sqlCmd.append("select nvl(cd01.hsien_id,' ') as hsien_id, nvl(cd01.hsien_name,'OTHER') as hsien_name, 'A01' as field_SEQ, ");           
       sqlCmd.append("       cd01.FR001W_output_order as FR001W_output_order, bn01.bank_no as bank_code, bn01.bank_name ");
       sqlCmd.append("  from (select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y') cd01 ");       
       sqlCmd.append("  left join (select * from wlx01 where m_year=? and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL)  )wlx01 on wlx01.hsien_id=cd01.hsien_id ");
       paramList.add(wlx01_m_year);
       sqlCmd.append("  left join (select * from bn01 where m_year=? "); 
       paramList.add(wlx01_m_year);
       if("ALL".equals(bank_type)){
           sqlCmd.append("            and bank_type in ('6','7') ");
       }else{
           sqlCmd.append("            and bank_type in (?) ");
           paramList.add(bank_type);
       }
       sqlCmd.append("                and bn_type <> '2')bn01 on wlx01.bank_no=bn01.bank_no ");    
       sqlCmd.append("where bn01.bank_no <> ' ' ");
       sqlCmd.append("ORDER by FR001W_output_order, field_SEQ, hsien_id, bank_code ");  
       List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
       System.out.println("dbData_Title.size()="+dbData.size()); 
       return dbData;
   }
   //各機構明細/縣市小計/總計資料
   public static List qList_Sum(String s_year,String s_month,String unit,String bank_type,String cd01_table,String wlx01_m_year){
       StringBuffer sqlCmd = new StringBuffer();  
       List paramList = new ArrayList();
       sqlCmd.append("select  a01.hsien_id , a01.hsien_name,  cd01.english as hsien_english,  a01.FR001W_output_order, ");  
       sqlCmd.append("        a01.bank_no ,  a01.BANK_NAME,  a01.english as bank_english,  a01.COUNT_SEQ, a01.field_SEQ, ");  
       sqlCmd.append("        field_DEBIT, ");  //--存款總額        
       sqlCmd.append("        field_CREDIT, "); //--放款總額         
       sqlCmd.append("        field_OVER, "); //--狹義逾期放款
       sqlCmd.append("        field_OVER_RATE, ");//--狹義逾放比率(狹義逾期放款/放款總額)                
       sqlCmd.append("        field_320300, "); //--本期損益
       sqlCmd.append("        field_NET, "); //--淨值
       sqlCmd.append("        field_BACKUP, ");//--備抵呆帳   
       sqlCmd.append("        field_CAPTIAL_RATE, ");//--淨值佔風險性資產比率 
       sqlCmd.append("        field_840740, ");//--廣義逾期放款
       sqlCmd.append("        field_840740_rate, ");//--廣義逾放比率(廣義逾期放款/放款總額)
       sqlCmd.append("        field_backup_over_rate, ");//--備呆占狹義逾期放款比率(備抵呆帳/狹義逾放)
       sqlCmd.append("        field_backup_840740_rate,  ");//--備呆占廣義逾期放款比率(備抵呆帳/廣義逾放) 
       sqlCmd.append("        field_captial_rate ");//淨值佔風險性資產比率
       sqlCmd.append("from (  ");  
       //--總計           
       sqlCmd.append("      select a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order, a01.bank_no ,   a01.BANK_NAME,  a01.english, ");         
       sqlCmd.append("             COUNT_SEQ, field_SEQ,  round(field_DEBIT /?,0)  as field_DEBIT, ");            
       sqlCmd.append("             round(field_CREDIT /?,0)  as field_CREDIT, round(field_OVER /?,0)  as field_OVER, ");         
       sqlCmd.append("             decode(a01.field_CREDIT,0,'N/A',round(a01.field_OVER /  a01.field_CREDIT *100 ,2))  as   field_OVER_RATE, ");        
       sqlCmd.append("             round(field_320300 /?,0)  as field_320300, round(field_NET /?,0)  as field_NET, ");         
       sqlCmd.append("             round(field_BACKUP /?,0)  as field_BACKUP, round(field_840740 /?,0)  as field_840740, ");         
       sqlCmd.append("             decode(field_CREDIT,0,'N/A',round(field_840740 /  a01.field_CREDIT *100 ,2))  as   field_840740_RATE, ");          
       sqlCmd.append("             decode(FIELD_OVER,0,'N/A',round(FIELD_BACKUP /  a01.FIELD_OVER *100 ,2))  as   field_backup_over_rate, ");          
       sqlCmd.append("             decode(field_840740,0,'N/A',round(FIELD_BACKUP /  field_840740 *100 ,2))  as   field_backup_840740_rate, ");   
       sqlCmd.append("             round(field_CAPTIAL / 1000  ,2)  as   field_CAPTIAL_RATE  "); 
       paramList.add(unit);
       paramList.add(unit);
       paramList.add(unit);
       paramList.add(unit);
       paramList.add(unit);
       paramList.add(unit);
       paramList.add(unit);
       sqlCmd.append("      from ( select ' '  AS  hsien_id ,  ' 總   計 '   AS hsien_name,  '001'  AS FR001W_output_order, ");                
       sqlCmd.append("               ' ' AS  bank_no ,     ' '   AS  BANK_NAME, 'Total' AS english,                 COUNT(*)  AS  COUNT_SEQ, ");                
       sqlCmd.append("               'A99'  as  field_SEQ,     SUM(field_320300)  field_320300 ,   SUM(field_NET)     field_NET, ");        
       sqlCmd.append("               SUM(field_BACKUP)  field_BACKUP, ");     
       sqlCmd.append("               SUM(field_OVER)    field_OVER,       SUM(field_DEBIT)   field_DEBIT,      SUM(field_CREDIT)  field_CREDIT, ");     
       sqlCmd.append("               SUM(field_840740)  field_840740, ");    
       sqlCmd.append("               decode(SUM(field_910400),0,0,round(SUM(field_910400) * 100000 /SUM(field_910500),0)) as field_CAPTIAL  "); 
       sqlCmd.append("             from  ");
       sqlCmd.append("             ( select nvl(cd01.hsien_id,' ')  as  hsien_id , nvl(cd01.hsien_name,'OTHER') as  hsien_name, ");               
       sqlCmd.append("                 cd01.FR001W_output_order   as  FR001W_output_order,  bn01.bank_no ,  bn01.BANK_NAME, wlx01.english, ");                
       sqlCmd.append("                 round(sum(decode(a01.acc_code,'320300',amt,0)) /1,0) as field_320300, "); 
       sqlCmd.append("                 round(sum(decode(a01.acc_code,'310000',amt,'320000',amt,0)) /1,0)     as field_NET, ");   
       sqlCmd.append("                 round(sum(decode(a01.acc_code, '120800',amt,'150300',amt,0)) /1,0) as  field_BACKUP, "); 
       sqlCmd.append("                 round(sum(decode(a01.acc_code,'990000',amt,0)) /1,0) as field_OVER,  ");  
       sqlCmd.append("                 round(sum(decode(a01.acc_code,'220000',amt,0)) /1,0) as field_DEBIT, ");  
       sqlCmd.append("                 round(sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /1,0) as  field_CREDIT ");  
       sqlCmd.append("               from  (select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y'  ) cd01  ");
       sqlCmd.append("               left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = ? and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL) ");
       paramList.add(wlx01_m_year);
       sqlCmd.append("               left join bn01 on wlx01.bank_no=bn01.bank_no ");
       if("ALL".equals(bank_type)){
           sqlCmd.append("            and bank_type in ('6','7') ");
       }else{
           sqlCmd.append("            and bank_type in (?) ");
           paramList.add(bank_type);
       }
       sqlCmd.append("                    and bn01.m_year = ? and bn_type <> '2' and wlx01.m_year=? ");
       paramList.add(wlx01_m_year);
       paramList.add(wlx01_m_year);
       sqlCmd.append("               left join (select * from a01 where a01.m_year  = ? and a01.m_month  = ?) a01 on  bn01.bank_no = a01.bank_code ");  
       paramList.add(s_year);
       paramList.add(s_month);
       sqlCmd.append("               group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME, wlx01.english ");  
       sqlCmd.append("            ) a01, "); 
       sqlCmd.append("            ( select nvl(cd01.hsien_id,' ')       as  hsien_id , nvl(cd01.hsien_name,'OTHER')  as  hsien_name, ");            
       sqlCmd.append("                cd01.FR001W_output_order     as  FR001W_output_order, bn01.bank_no as bank_code,  bn01.bank_name, wlx01.english, ");            
       sqlCmd.append("                round(sum(decode(a05.acc_code,'910400',amt,0)) /1,0) as field_910400, ");            
       sqlCmd.append("                round(sum(decode(a05.acc_code,'910500',amt,0)) /1,0) as field_910500  ");      
       sqlCmd.append("              from  (select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y') cd01     ");   
       sqlCmd.append("              left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = ? and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL)  ");
       paramList.add(wlx01_m_year);
       sqlCmd.append("              left join bn01 on wlx01.bank_no=bn01.bank_no ");
       if("ALL".equals(bank_type)){
           sqlCmd.append("            and bank_type in ('6','7') ");
       }else{
           sqlCmd.append("            and bank_type in (?) ");
           paramList.add(bank_type);
       }
       sqlCmd.append("                and bn01.m_year = ? and bn_type <> '2' and wlx01.m_year=? ");
       paramList.add(wlx01_m_year);
       paramList.add(wlx01_m_year);
       sqlCmd.append("              left join (select * from a05 where a05.m_year  = ?  and a05.m_month  =? and a05.ACC_code in ('910400','910500','91060P') ");
       paramList.add(s_year);
       paramList.add(s_month);
       sqlCmd.append("            ) a05  on  bn01.bank_no = a05.bank_code ");      
       sqlCmd.append("            group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME, wlx01.english ");  
       sqlCmd.append("            ) a05, ");
       sqlCmd.append("            ( select nvl(cd01.hsien_id,' ')  as  hsien_id ,  nvl(cd01.hsien_name,'OTHER')  as  hsien_name, ");             
       sqlCmd.append("                cd01.FR001W_output_order     as  FR001W_output_order, bn01.bank_no as bank_code,  bn01.bank_name, wlx01.english, ");             
       sqlCmd.append("                round(sum(decode(a02.acc_code,'990611',amt,0)) /1,0) as field_990611 ");      
       sqlCmd.append("              from  (select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y') cd01    ");    
       sqlCmd.append("              left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = ? and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL) ");  
       paramList.add(wlx01_m_year);
       sqlCmd.append("              left join bn01 on wlx01.bank_no=bn01.bank_no ");
       if("ALL".equals(bank_type)){
           sqlCmd.append("            and bank_type in ('6','7') ");
       }else{
           sqlCmd.append("            and bank_type in (?) ");
           paramList.add(bank_type);
       }
       sqlCmd.append("                and bn01.m_year = ? and bn_type <> '2' and wlx01.m_year = ? ");
       paramList.add(wlx01_m_year);
       paramList.add(wlx01_m_year);
       sqlCmd.append("              left join (select * from a02 where a02.m_year  = ? and a02.m_month  = ? and a02.ACC_code in ('990611') ) a02  on  bn01.bank_no = a02.bank_code ");
       paramList.add(s_year);
       paramList.add(s_month);
       sqlCmd.append("              group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME, wlx01.english ");  
       sqlCmd.append("            ) a02  , ");
       sqlCmd.append("            ( select nvl(cd01.hsien_id,' ')  as  hsien_id ,   nvl(cd01.hsien_name,'OTHER')  as  hsien_name, ");             
       sqlCmd.append("                cd01.FR001W_output_order     as  FR001W_output_order,  bn01.bank_no as bank_code,  bn01.bank_name, wlx01.english, ");             
       sqlCmd.append("                round(sum(decode(a04.acc_code,'840740',amt,'840760',amt,0)) /1,0) as field_840740   ");    
       sqlCmd.append("              from  (select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y') cd01 ");       
       sqlCmd.append("              left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = ? and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL) ");
       paramList.add(wlx01_m_year);
       sqlCmd.append("              left join bn01 on wlx01.bank_no=bn01.bank_no ");
       if("ALL".equals(bank_type)){
           sqlCmd.append("            and bank_type in ('6','7') ");
       }else{
           sqlCmd.append("            and bank_type in (?) ");
           paramList.add(bank_type);
       }
       sqlCmd.append("                        and bn01.m_year = ? and bn_type <> '2' and wlx01.m_year = ?  ");
       paramList.add(wlx01_m_year);
       paramList.add(wlx01_m_year);
       sqlCmd.append("              left join (select * from a04 where a04.m_year = ? and a04.m_month = ? and a04.ACC_code in ('840740','840760')) a04 on  bn01.bank_no = a04.bank_code "); 
       paramList.add(s_year);
       paramList.add(s_month);
       sqlCmd.append("              group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME, wlx01.english  "); 
       sqlCmd.append("            ) a04 ");  
       sqlCmd.append("            where  a01.bank_no=a05.bank_code(+) and a01.bank_no=a02.bank_code(+) and a01.bank_no=a04.bank_code(+) and a01.bank_no <> ' ' ");  
       sqlCmd.append("      ) a01 "); 
       sqlCmd.append("      UNION ALL  ");
       //--各別機構明細
       sqlCmd.append("      select a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order,         a01.bank_no ,   a01.BANK_NAME,  a01.english, ");         
       sqlCmd.append("             COUNT_SEQ, field_SEQ,            round(field_DEBIT /?,0)  as field_DEBIT, ");            
       sqlCmd.append("             round(field_CREDIT /?,0)  as field_CREDIT, ");         
       sqlCmd.append("             round(field_OVER /?,0)  as field_OVER, ");         
       sqlCmd.append("             decode(a01.field_CREDIT,0,'N/A',round(a01.field_OVER /  a01.field_CREDIT *100 ,2))  as   field_OVER_RATE, ");        
       sqlCmd.append("             round(field_320300 /?,0)  as field_320300, round(field_NET /?,0)  as field_NET, ");         
       sqlCmd.append("             round(field_BACKUP /?,0)  as field_BACKUP, round(field_840740 /?,0)  as field_840740, ");         
       sqlCmd.append("             decode(field_CREDIT,0,'N/A',round(field_840740 /  a01.field_CREDIT *100 ,2))  as   field_840740_RATE,   ");        
       sqlCmd.append("             decode(FIELD_OVER,0,'N/A',round(FIELD_BACKUP /  a01.FIELD_OVER *100 ,2))  as   field_backup_over_rate,  ");         
       sqlCmd.append("             decode(field_840740,0,'N/A',round(FIELD_BACKUP /  field_840740 *100 ,2))  as   field_backup_840740_rate, ");   
       sqlCmd.append("             round(field_CAPTIAL /  1000 ,2)  as   field_CAPTIAL_RATE   ");
       paramList.add(unit);
       paramList.add(unit);
       paramList.add(unit);
       paramList.add(unit);
       paramList.add(unit);
       paramList.add(unit);
       paramList.add(unit);
       sqlCmd.append("      from ( select a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order, ");                  
       sqlCmd.append("               a01.bank_no ,   a01.BANK_NAME, a01.english,  1  AS  COUNT_SEQ,  'A01'  as  field_SEQ, ");    
       sqlCmd.append("               SUM(field_320300)  field_320300 ,SUM(field_NET)     field_NET,   SUM(field_BACKUP)  field_BACKUP, ");     
       sqlCmd.append("               SUM(field_OVER)    field_OVER,   SUM(field_DEBIT)   field_DEBIT, SUM(field_CREDIT)  field_CREDIT, ");     
       sqlCmd.append("               SUM(field_840740)  field_840740, SUM(nvl(a05.amt,0)) as field_CAPTIAL ");  
       sqlCmd.append("             from  ");
       sqlCmd.append("             ( select nvl(cd01.hsien_id,' ')  as  hsien_id ,  nvl(cd01.hsien_name,'OTHER') as  hsien_name, ");               
       sqlCmd.append("                 cd01.FR001W_output_order     as  FR001W_output_order,  bn01.bank_no ,  bn01.BANK_NAME, wlx01.english, ");          
       sqlCmd.append("                 round(sum(decode(a01.acc_code,'320300',amt,0)) /1,0) as field_320300, "); 
       sqlCmd.append("                 round(sum(decode(a01.acc_code,'310000',amt,'320000',amt,0)) /1,0)     as field_NET, ");   
       sqlCmd.append("                 round(sum(decode(a01.acc_code, '120800',amt,'150300',amt,0)) /1,0) as  field_BACKUP, "); 
       sqlCmd.append("                 round(sum(decode(a01.acc_code,'990000',amt,0)) /1,0) as field_OVER, ");   
       sqlCmd.append("                 round(sum(decode(a01.acc_code,'220000',amt,0)) /1,0) as field_DEBIT, ");  
       sqlCmd.append("                 round(sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /1,0) as  field_CREDIT ");  
       sqlCmd.append("               from  (select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y'  ) cd01 "); 
       sqlCmd.append("               left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = ? and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL) ");
       paramList.add(wlx01_m_year);
       sqlCmd.append("               left join bn01 on wlx01.bank_no=bn01.bank_no ");
       if("ALL".equals(bank_type)){
           sqlCmd.append("            and bank_type in ('6','7') ");
       }else{
           sqlCmd.append("            and bank_type in (?) ");
           paramList.add(bank_type);
       }
       sqlCmd.append("               and bn01.m_year = ? and bn_type <> '2' and wlx01.m_year=? ");
       paramList.add(wlx01_m_year);
       paramList.add(wlx01_m_year);
       sqlCmd.append("               left join (select * from a01 where a01.m_year  = ? and a01.m_month  = ?) a01 on  bn01.bank_no = a01.bank_code ");
       paramList.add(s_year);
       paramList.add(s_month);
       sqlCmd.append("               group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME, wlx01.english ");  
       sqlCmd.append("             ) a01, ");    
       sqlCmd.append("             (select * from a05 where a05.m_year = ? and a05.m_month = ? and  a05.ACC_code = '91060P') a05, ");
       paramList.add(s_year);
       paramList.add(s_month);
       sqlCmd.append("             (select nvl(cd01.hsien_id,' ')  as  hsien_id ,  nvl(cd01.hsien_name,'OTHER')  as  hsien_name, ");            
       sqlCmd.append("                cd01.FR001W_output_order     as  FR001W_output_order, bn01.bank_no as bank_code,  bn01.bank_name, wlx01.english, ");            
       sqlCmd.append("                round(sum(decode(a02.acc_code,'990611',amt,0)) /1,0) as field_990611 ");       
       sqlCmd.append("              from  (select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y') cd01    ");     
       sqlCmd.append("              left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = ? and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL) ");
       paramList.add(wlx01_m_year);
       sqlCmd.append("              left join bn01 on wlx01.bank_no=bn01.bank_no ");
       if("ALL".equals(bank_type)){
           sqlCmd.append("            and bank_type in ('6','7') ");
       }else{
           sqlCmd.append("            and bank_type in (?) ");
           paramList.add(bank_type);
       }
       sqlCmd.append("              and bn01.m_year = ? and bn_type <> '2' and wlx01.m_year = ? ");
       paramList.add(wlx01_m_year);
       paramList.add(wlx01_m_year);
       sqlCmd.append("              left join (select * from a02 where a02.m_year = ? and a02.m_month = ? and a02.ACC_code in ('990611') ) a02 on  bn01.bank_no = a02.bank_code ");
       paramList.add(s_year);
       paramList.add(s_month);
       sqlCmd.append("              group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME, wlx01.english  "); 
       sqlCmd.append("             ) a02  , ");
       sqlCmd.append("             (select nvl(cd01.hsien_id,' ') as  hsien_id , nvl(cd01.hsien_name,'OTHER')  as  hsien_name, ");            
       sqlCmd.append("                cd01.FR001W_output_order     as  FR001W_output_order, bn01.bank_no as bank_code,  bn01.bank_name, wlx01.english, ");            
       sqlCmd.append("                round(sum(decode(a04.acc_code,'840740',amt,'840760',amt,0)) /1,0) as field_840740  ");      
       sqlCmd.append("             from (select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y') cd01 ");        
       sqlCmd.append("             left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = ? and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL) ");
       paramList.add(wlx01_m_year);
       sqlCmd.append("             left join bn01 on wlx01.bank_no=bn01.bank_no ");
       if("ALL".equals(bank_type)){
           sqlCmd.append("            and bank_type in ('6','7') ");
       }else{
           sqlCmd.append("            and bank_type in (?) ");
           paramList.add(bank_type);
       }
       sqlCmd.append("             and bn01.m_year = ? and bn_type <> '2' and wlx01.m_year = ?  ");
       paramList.add(wlx01_m_year);
       paramList.add(wlx01_m_year);
       sqlCmd.append("             left join (select * from a04 where a04.m_year = ? and a04.m_month = ? ");
       paramList.add(s_year);
       paramList.add(s_month);
       sqlCmd.append("               and a04.ACC_code in ('840740','840760') ) a04 on  bn01.bank_no = a04.bank_code  ");    
       sqlCmd.append("             group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME, wlx01.english   ");
       sqlCmd.append("             ) a04 "); 
       sqlCmd.append("             where a01.bank_no=a05.bank_code(+) and a01.bank_no=a02.bank_code(+) and a01.bank_no = a04.bank_code(+) and a01.bank_no <> ' ' ");  
       sqlCmd.append("             GROUP BY a01.hsien_id,a01.hsien_name,a01.FR001W_output_order,a01.bank_no,a01.BANK_NAME,a01.english  ");  
       sqlCmd.append("      ) a01  ");  
       sqlCmd.append("      UNION ALL "); 
       //--縣市小計
       sqlCmd.append("      select a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order,         a01.bank_no ,   a01.BANK_NAME,  a01.english, ");         
       sqlCmd.append("            COUNT_SEQ, field_SEQ, round(field_DEBIT /?,0)  as field_DEBIT, ");            
       sqlCmd.append("            round(field_CREDIT /?,0)  as field_CREDIT, round(field_OVER /?,0)  as field_OVER, ");         
       sqlCmd.append("            decode(a01.field_CREDIT,0,'N/A',round(a01.field_OVER /  a01.field_CREDIT *100 ,2))  as   field_OVER_RATE, ");        
       sqlCmd.append("            round(field_320300 /?,0)  as field_320300,  round(field_NET /?,0)  as field_NET, ");         
       sqlCmd.append("            round(field_BACKUP /?,0)  as field_BACKUP, round(field_840740 /?,0)  as field_840740, ");         
       sqlCmd.append("            decode(field_CREDIT,0,'N/A',round(field_840740 /  a01.field_CREDIT *100 ,2))  as   field_840740_RATE, ");    
       sqlCmd.append("            decode(FIELD_OVER,0,'N/A',round(FIELD_BACKUP /  a01.FIELD_OVER *100 ,2))  as   field_backup_over_rate, ");      
       sqlCmd.append("            decode(field_840740,0,'N/A',round(FIELD_BACKUP /  field_840740 *100 ,2))  as   field_backup_840740_rate, ");   
       sqlCmd.append("            round(field_CAPTIAL / 1000  ,2)  as   field_CAPTIAL_RATE  "); 
       paramList.add(unit);
       paramList.add(unit);
       paramList.add(unit);
       paramList.add(unit);
       paramList.add(unit);
       paramList.add(unit);
       paramList.add(unit);
       sqlCmd.append("      from ( select a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order, ");               
       sqlCmd.append("               ' ' AS  bank_no ,   ' ' AS  BANK_NAME, '' as english,  COUNT(*)  AS  COUNT_SEQ, ");                
       sqlCmd.append("               'A90'  as  field_SEQ,   SUM(field_320300)  field_320300 ,   SUM(field_NET)     field_NET, ");        
       sqlCmd.append("               SUM(field_BACKUP)  field_BACKUP,     SUM(field_OVER)    field_OVER,       SUM(field_DEBIT)   field_DEBIT, ");      
       sqlCmd.append("               SUM(field_CREDIT)  field_CREDIT,     SUM(field_840740)  field_840740, ");    
       sqlCmd.append("               decode(SUM(field_910400),0,0,round(SUM(field_910400) * 100000 /SUM(field_910500),0)) as field_CAPTIAL  "); 
       sqlCmd.append("             from  ");
       sqlCmd.append("             (select nvl(cd01.hsien_id,' ')  as  hsien_id ,  nvl(cd01.hsien_name,'OTHER') as  hsien_name, ");               
       sqlCmd.append("                cd01.FR001W_output_order  as  FR001W_output_order, bn01.bank_no ,  bn01.BANK_NAME, wlx01.english, ");                
       sqlCmd.append("                round(sum(decode(a01.acc_code,'320300',amt,0)) /1,0) as field_320300, "); 
       sqlCmd.append("                round(sum(decode(a01.acc_code,'310000',amt,'320000',amt,0)) /1,0)     as field_NET,  ");  
       sqlCmd.append("                round(sum(decode(a01.acc_code, '120800',amt,'150300',amt,0)) /1,0) as  field_BACKUP, "); 
       sqlCmd.append("                round(sum(decode(a01.acc_code,'990000',amt,0)) /1,0) as field_OVER, ");   
       sqlCmd.append("                round(sum(decode(a01.acc_code,'220000',amt,0)) /1,0) as field_DEBIT, ");  
       sqlCmd.append("                round(sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /1,0) as  field_CREDIT ");  
       sqlCmd.append("             from  (select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y'  ) cd01 "); 
       sqlCmd.append("             left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = ? and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL) ");
       paramList.add(wlx01_m_year);
       sqlCmd.append("             left join bn01 on wlx01.bank_no=bn01.bank_no ");
       if("ALL".equals(bank_type)){
           sqlCmd.append("            and bank_type in ('6','7') ");
       }else{
           sqlCmd.append("            and bank_type in (?) ");
           paramList.add(bank_type);
       }
       sqlCmd.append("             and bn01.m_year = ? and bn_type <> '2' and wlx01.m_year=? ");
       paramList.add(wlx01_m_year);
       paramList.add(wlx01_m_year);
       sqlCmd.append("             left join (select * from a01 where a01.m_year  = ? and a01.m_month  = ?) a01 on  bn01.bank_no = a01.bank_code "); 
       paramList.add(s_year);
       paramList.add(s_month);
       sqlCmd.append("             group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME, wlx01.english ");  
       sqlCmd.append("             ) a01, ");     
       sqlCmd.append("             (select nvl(cd01.hsien_id,' ') as  hsien_id ,  nvl(cd01.hsien_name,'OTHER')  as  hsien_name, ");           
       sqlCmd.append("               cd01.FR001W_output_order     as  FR001W_output_order,            bn01.bank_no as bank_code,  bn01.bank_name, wlx01.english, ");           
       sqlCmd.append("               round(sum(decode(a05.acc_code,'910400',amt,0)) /1,0) as field_910400, ");           
       sqlCmd.append("               round(sum(decode(a05.acc_code,'910500',amt,0)) /1,0) as field_910500  ");      
       sqlCmd.append("             from  (select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y') cd01     ");   
       sqlCmd.append("             left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = ? and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL) ");
       paramList.add(wlx01_m_year);
       sqlCmd.append("             left join bn01 on wlx01.bank_no=bn01.bank_no ");
       if("ALL".equals(bank_type)){
           sqlCmd.append("            and bank_type in ('6','7') ");
       }else{
           sqlCmd.append("            and bank_type in (?) ");
           paramList.add(bank_type);
       }
       sqlCmd.append("             and bn01.m_year = ? and bn_type <> '2' and wlx01.m_year = ?  ");
       paramList.add(wlx01_m_year);
       paramList.add(wlx01_m_year);
       sqlCmd.append("             left join (select * from a05 where a05.m_year = ? and a05.m_month  = ? "); 
       paramList.add(s_year);
       paramList.add(s_month);
       sqlCmd.append("                and a05.ACC_code in ('910400','910500','91060P') ) a05 on  bn01.bank_no = a05.bank_code   "); 
       sqlCmd.append("             group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order, bn01.bank_no,bn01.BANK_NAME, wlx01.english ");  
       sqlCmd.append("             ) a05, ");
       sqlCmd.append("             (select nvl(cd01.hsien_id,' ')  as  hsien_id ,  nvl(cd01.hsien_name,'OTHER')  as  hsien_name, ");             
       sqlCmd.append("                cd01.FR001W_output_order     as  FR001W_output_order,  bn01.bank_no as bank_code,  bn01.bank_name, wlx01.english, ");             
       sqlCmd.append("                round(sum(decode(a02.acc_code,'990611',amt,0)) /1,0) as field_990611 ");      
       sqlCmd.append("             from  (select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y') cd01      ");  
       sqlCmd.append("             left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = ? and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL)  ");
       paramList.add(wlx01_m_year);
       sqlCmd.append("             left join bn01 on wlx01.bank_no=bn01.bank_no ");
       if("ALL".equals(bank_type)){
           sqlCmd.append("            and bank_type in ('6','7') ");
       }else{
           sqlCmd.append("            and bank_type in (?) ");
           paramList.add(bank_type);
       }
       sqlCmd.append("                    and bn01.m_year = ? and bn_type <> '2' and wlx01.m_year = ?  ");
       paramList.add(wlx01_m_year);
       paramList.add(wlx01_m_year);
       sqlCmd.append("             left join (select * from a02 where a02.m_year = ? and a02.m_month = ? "); 
       paramList.add(s_year);
       paramList.add(s_month);
       sqlCmd.append("                and a02.ACC_code in ('990611') ) a02 on  bn01.bank_no = a02.bank_code   ");   
       sqlCmd.append("             group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME, wlx01.english ");  
       sqlCmd.append("             )a02 , ");
       sqlCmd.append("             (select nvl(cd01.hsien_id,' ')  as  hsien_id ,   nvl(cd01.hsien_name,'OTHER')  as  hsien_name, ");  
       sqlCmd.append("                cd01.FR001W_output_order     as  FR001W_output_order, bn01.bank_no as bank_code,  bn01.bank_name, wlx01.english, ");             
       sqlCmd.append("                round(sum(decode(a04.acc_code,'840740',amt,'840760',amt,0)) /1,0) as field_840740 ");      
       sqlCmd.append("              from  (select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y') cd01 ");       
       sqlCmd.append("              left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = ? and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL) ");
       paramList.add(wlx01_m_year);
       sqlCmd.append("              left join bn01 on wlx01.bank_no=bn01.bank_no ");
       if("ALL".equals(bank_type)){
           sqlCmd.append("            and bank_type in ('6','7') ");
       }else{
           sqlCmd.append("            and bank_type in (?) ");
           paramList.add(bank_type);
       }
       sqlCmd.append("                        and bn01.m_year = ? and bn_type <> '2' and wlx01.m_year = ? "); 
       paramList.add(wlx01_m_year);
       paramList.add(wlx01_m_year);
       sqlCmd.append("              left join (select * from a04 where a04.m_year = ? and a04.m_month = ? "); 
       paramList.add(s_year);
       paramList.add(s_month);
       sqlCmd.append("                 and a04.ACC_code in ('840740','840760') ) a04 on  bn01.bank_no = a04.bank_code ");     
       sqlCmd.append("              group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME, wlx01.english  "); 
       sqlCmd.append("             ) a04 ");  
       sqlCmd.append("             where a01.bank_no=a05.bank_code(+) and a01.bank_no=a02.bank_code(+) and a01.bank_no=a04.bank_code(+) and a01.bank_no <> ' ' ");  
       sqlCmd.append("             GROUP BY a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order ");  
       sqlCmd.append("      ) a01 ");   
       sqlCmd.append(")  a01 ");  
       sqlCmd.append("left join  (select * from  cd01 where cd01.hsien_id <> 'Y') cd01  on a01.hsien_id = cd01.hsien_id "); 
       sqlCmd.append("ORDER by    FR001W_output_order, field_SEQ,  hsien_id ,  bank_no ");
       List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"hsien_id,hsien_name,hsien_english,fr001w_output_order,bank_no,bank_name,bank_english,"+
                                                           "count_seq,field_seq,field_debit,field_credit,field_over,field_over_rate,field_320300,"+
                                                           "field_net,field_backup,field_captial_rate,field_840740,field_840740_rate,field_backup_over_rate,field_backup_840740_rate");
       System.out.println("dbData_Sum.size()="+dbData.size()); 
       return dbData;
   }
   //福建省.單一年度資料
   public static List qList_Fukien(String s_year,String s_month,String unit,String bank_type,String cd01_table,String wlx01_m_year){
       StringBuffer sqlCmd_Fukien  = new StringBuffer(); 
       List paramList_Fukien = new ArrayList();
       sqlCmd_Fukien.append("select a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order, a01.bank_no ,   a01.BANK_NAME,  a01.english, ");         
       sqlCmd_Fukien.append("       COUNT_SEQ, field_SEQ, round(field_DEBIT /?,0)  as field_DEBIT, ");            
       sqlCmd_Fukien.append("       round(field_CREDIT /?,0)  as field_CREDIT, round(field_OVER /?,0)  as field_OVER, ");         
       sqlCmd_Fukien.append("       decode(a01.field_CREDIT,0,'N/A',round(a01.field_OVER /  a01.field_CREDIT *100 ,2))  as   field_OVER_RATE, ");        
       sqlCmd_Fukien.append("       round(field_320300 /?,0)  as field_320300, round(field_NET /?,0)  as field_NET, ");         
       sqlCmd_Fukien.append("       round(field_BACKUP /?,0)  as field_BACKUP, round(field_840740 /?,0)  as field_840740, ");         
       sqlCmd_Fukien.append("       decode(field_CREDIT,0,'N/A',round(field_840740 /  a01.field_CREDIT *100 ,2))  as   field_840740_RATE, ");          
       sqlCmd_Fukien.append("       decode(FIELD_OVER,0,'N/A',round(FIELD_BACKUP /  a01.FIELD_OVER *100 ,2))  as   field_backup_over_rate,  ");         
       sqlCmd_Fukien.append("       decode(field_840740,0,'N/A',round(FIELD_BACKUP /  field_840740 *100 ,2))  as   field_backup_840740_rate, ");   
       sqlCmd_Fukien.append("       round(field_CAPTIAL / 1000,2)  as   field_CAPTIAL_RATE ");   
       paramList_Fukien.add(unit);
       paramList_Fukien.add(unit);
       paramList_Fukien.add(unit);
       paramList_Fukien.add(unit);
       paramList_Fukien.add(unit);
       paramList_Fukien.add(unit);
       paramList_Fukien.add(unit);
       sqlCmd_Fukien.append("from( select ' '  AS  hsien_id ,  '福建省'   AS hsien_name,  '235'  AS FR001W_output_order, ");                  
       sqlCmd_Fukien.append("             ' ' AS  bank_no ,     ' '   AS  BANK_NAME, 'Fuchien Province' AS english, ");                  
       sqlCmd_Fukien.append("             COUNT(*)  AS  COUNT_SEQ,                  'A93'  as  field_SEQ, ");     
       sqlCmd_Fukien.append("             SUM(field_320300)  field_320300 ,   SUM(field_NET)     field_NET, ");        
       sqlCmd_Fukien.append("             SUM(field_BACKUP)  field_BACKUP,     SUM(field_OVER)    field_OVER, ");       
       sqlCmd_Fukien.append("             SUM(field_DEBIT)   field_DEBIT,      SUM(field_CREDIT)  field_CREDIT, ");     
       sqlCmd_Fukien.append("             SUM(field_840740)  field_840740, ");    
       sqlCmd_Fukien.append("             decode(SUM(field_910400),0,0,round(SUM(field_910400) * 100000 /SUM(field_910500),0)) as field_CAPTIAL ");  
       sqlCmd_Fukien.append("       from  ");
       sqlCmd_Fukien.append("         (select nvl(cd01.hsien_id,' ')  as  hsien_id , nvl(cd01.hsien_name,'OTHER') as  hsien_name, ");               
       sqlCmd_Fukien.append("                 cd01.FR001W_output_order     as  FR001W_output_order, bn01.bank_no ,  bn01.BANK_NAME, wlx01.english, ");                
       sqlCmd_Fukien.append("                 round(sum(decode(a01.acc_code,'320300',amt,0)) /1,0) as field_320300, "); 
       sqlCmd_Fukien.append("                 round(sum(decode(a01.acc_code,'310000',amt,'320000',amt,0)) /1,0)     as field_NET, ");   
       sqlCmd_Fukien.append("                 round(sum(decode(a01.acc_code, '120800',amt,'150300',amt,0)) /1,0) as  field_BACKUP, "); 
       sqlCmd_Fukien.append("                 round(sum(decode(a01.acc_code,'990000',amt,0)) /1,0) as field_OVER,  ");  
       sqlCmd_Fukien.append("                 round(sum(decode(a01.acc_code,'220000',amt,0)) /1,0) as field_DEBIT, ");  
       sqlCmd_Fukien.append("                 round(sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /1,0) as  field_CREDIT ");  
       sqlCmd_Fukien.append("          from  (select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y'  and cd01.Hsien_div = '3' ) cd01  ");
       sqlCmd_Fukien.append("          left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = ? and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL) ");
       paramList_Fukien.add(wlx01_m_year);
       sqlCmd_Fukien.append("          left join bn01 on wlx01.bank_no=bn01.bank_no ");
       if("ALL".equals(bank_type)){
           sqlCmd_Fukien.append("            and bn01.bank_type in ('6','7') ");
       }else{
           sqlCmd_Fukien.append("            and bn01.bank_type in (?) ");
           paramList_Fukien.add(bank_type);
       }
       sqlCmd_Fukien.append("                    and bn01.m_year = ? and bn_type <> '2' and wlx01.m_year=? ");
       paramList_Fukien.add(wlx01_m_year);
       paramList_Fukien.add(wlx01_m_year);
       sqlCmd_Fukien.append("          left join (select * from a01 where a01.m_year  = ? and a01.m_month  = ?) a01 on  bn01.bank_no = a01.bank_code ");
       paramList_Fukien.add(s_year);
       paramList_Fukien.add(s_month);
       sqlCmd_Fukien.append("          group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME, wlx01.english ");  
       sqlCmd_Fukien.append("       ) a01, "); 
       sqlCmd_Fukien.append("       (  select nvl(cd01.hsien_id,' ')       as  hsien_id , ");              
       sqlCmd_Fukien.append("                 nvl(cd01.hsien_name,'OTHER')  as  hsien_name,            cd01.FR001W_output_order     as  FR001W_output_order, ");             
       sqlCmd_Fukien.append("                 bn01.bank_no as bank_code,  bn01.bank_name, wlx01.english, ");            
       sqlCmd_Fukien.append("                 round(sum(decode(a05.acc_code,'910400',amt,0)) /1,0) as field_910400, ");            
       sqlCmd_Fukien.append("                 round(sum(decode(a05.acc_code,'910500',amt,0)) /1,0) as field_910500  ");      
       sqlCmd_Fukien.append("          from  (select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y' and cd01.Hsien_div = '3') cd01 ");       
       sqlCmd_Fukien.append("          left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = ? and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL) "); 
       paramList_Fukien.add(wlx01_m_year);
       sqlCmd_Fukien.append("          left join bn01 on wlx01.bank_no=bn01.bank_no  ");
       if("ALL".equals(bank_type)){
           sqlCmd_Fukien.append("            and bn01.bank_type in ('6','7') ");
       }else{
           sqlCmd_Fukien.append("            and bn01.bank_type in (?) ");
           paramList_Fukien.add(bank_type);
       }
       sqlCmd_Fukien.append("                 and bn01.m_year = ? and bn_type <> '2' and wlx01.m_year = ?  ");  
       paramList_Fukien.add(wlx01_m_year);
       paramList_Fukien.add(wlx01_m_year);
       sqlCmd_Fukien.append("          left join (select * from a05 where a05.m_year  = ? and a05.m_month  = ? "); 
       paramList_Fukien.add(s_year);
       paramList_Fukien.add(s_month);
       sqlCmd_Fukien.append("          and a05.ACC_code in ('910400','910500','91060P') ) a05 on  bn01.bank_no = a05.bank_code ");      
       sqlCmd_Fukien.append("          group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME, wlx01.english  "); 
       sqlCmd_Fukien.append("       ) a05  , ");
       sqlCmd_Fukien.append("       (  select nvl(cd01.hsien_id,' ')       as  hsien_id , ");              
       sqlCmd_Fukien.append("                 nvl(cd01.hsien_name,'OTHER')  as  hsien_name,cd01.FR001W_output_order     as  FR001W_output_order, ");              
       sqlCmd_Fukien.append("                 bn01.bank_no as bank_code,  bn01.bank_name, wlx01.english, ");             
       sqlCmd_Fukien.append("                 round(sum(decode(a02.acc_code,'990611',amt,0)) /1,0) as field_990611 ");      
       sqlCmd_Fukien.append("          from  (select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y' and cd01.Hsien_div = '3') cd01 ");       
       sqlCmd_Fukien.append("          left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = ? and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL) ");
       paramList_Fukien.add(wlx01_m_year);
       sqlCmd_Fukien.append("          left join bn01 on wlx01.bank_no=bn01.bank_no  ");
       if("ALL".equals(bank_type)){
           sqlCmd_Fukien.append("            and bn01.bank_type in ('6','7') ");
       }else{
           sqlCmd_Fukien.append("            and bn01.bank_type in (?) ");
           paramList_Fukien.add(bank_type);
       }
       sqlCmd_Fukien.append("               and bn01.m_year = ? and bn_type <> '2' and wlx01.m_year = ? ");  
       paramList_Fukien.add(wlx01_m_year);
       paramList_Fukien.add(wlx01_m_year);
       sqlCmd_Fukien.append("          left join (select * from a02 where a02.m_year  = ? and a02.m_month  = ? ");
       paramList_Fukien.add(s_year);
       paramList_Fukien.add(s_month);
       sqlCmd_Fukien.append("          and a02.ACC_code in ('990611') ) a02  on  bn01.bank_no = a02.bank_code ");     
       sqlCmd_Fukien.append("          group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME, wlx01.english ");  
       sqlCmd_Fukien.append("        ) a02 , "); 
       sqlCmd_Fukien.append("        ( select nvl(cd01.hsien_id,' ')       as  hsien_id ,              nvl(cd01.hsien_name,'OTHER')  as  hsien_name, ");             
       sqlCmd_Fukien.append("                 cd01.FR001W_output_order     as  FR001W_output_order,              bn01.bank_no as bank_code,  bn01.bank_name, wlx01.english, ");             
       sqlCmd_Fukien.append("                 round(sum(decode(a04.acc_code,'840740',amt,'840760',amt,0)) /1,0) as field_840740 ");      
       sqlCmd_Fukien.append("           from  (select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y' and cd01.Hsien_div = '3') cd01 ");       
       sqlCmd_Fukien.append("           left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = ? and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL) "); 
       paramList_Fukien.add(wlx01_m_year);
       sqlCmd_Fukien.append("           left join bn01 on wlx01.bank_no=bn01.bank_no ");
       if("ALL".equals(bank_type)){
           sqlCmd_Fukien.append("            and bn01.bank_type in ('6','7') ");
       }else{
           sqlCmd_Fukien.append("            and bn01.bank_type in (?) ");
           paramList_Fukien.add(bank_type);
       }
       sqlCmd_Fukien.append("                   and bn01.m_year = ? and bn_type <> '2' and wlx01.m_year = ? ");
       paramList_Fukien.add(wlx01_m_year);
       paramList_Fukien.add(wlx01_m_year);
       sqlCmd_Fukien.append("           left join (select * from a04 where a04.m_year = ? and a04.m_month  = ? "); 
       paramList_Fukien.add(s_year);
       paramList_Fukien.add(s_month);
       sqlCmd_Fukien.append("           and a04.ACC_code in ('840740','840760') ) a04 on  bn01.bank_no = a04.bank_code ");     
       sqlCmd_Fukien.append("           group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME, wlx01.english ");  
       sqlCmd_Fukien.append("        ) a04 ");  
       sqlCmd_Fukien.append("where  a01.bank_no=a05.bank_code(+) and a01.bank_no=a02.bank_code(+) and a01.bank_no=a04.bank_code(+) and a01.bank_no <> ' ' ");  
       sqlCmd_Fukien.append(") a01 "); 
       List dbData = DBManager.QueryDB_SQLParam(sqlCmd_Fukien.toString(),paramList_Fukien,"hsien_id,hsien_name,fr001w_output_order,bank_no,bank_name,english," +
                                                               		"count_seq,field_seq,field_debit,field_credit,field_over,field_over_rate,field_320300,field_net," +
                                                               		"field_backup,field_840740,field_840740_rate,field_backup_over_rate,field_backup_840740_rate,field_captial_rate");
       System.out.println("dbData_Fukien.size()="+dbData.size()); 
       return dbData;
   }
   //臺灣省.單一年度資料
   public static List qList_Taiwan(String s_year,String s_month,String unit,String bank_type,String cd01_table,String wlx01_m_year){
       StringBuffer sqlCmd_Taiwan = new StringBuffer(); 
       List paramList_Taiwan = new ArrayList();
       sqlCmd_Taiwan.append("select a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order, a01.bank_no ,   a01.BANK_NAME,  a01.english, ");         
       sqlCmd_Taiwan.append("       COUNT_SEQ, field_SEQ, round(field_DEBIT /?,0)  as field_DEBIT, ");            
       sqlCmd_Taiwan.append("       round(field_CREDIT /?,0)  as field_CREDIT, round(field_OVER /?,0)  as field_OVER, ");         
       sqlCmd_Taiwan.append("       decode(a01.field_CREDIT,0,'N/A',round(a01.field_OVER /  a01.field_CREDIT *100 ,2))  as   field_OVER_RATE, ");        
       sqlCmd_Taiwan.append("       round(field_320300 /?,0)  as field_320300, round(field_NET /?,0)  as field_NET, ");         
       sqlCmd_Taiwan.append("       round(field_BACKUP /?,0)  as field_BACKUP, round(field_840740 /?,0)  as field_840740, ");         
       sqlCmd_Taiwan.append("       decode(field_CREDIT,0,'N/A',round(field_840740 /  a01.field_CREDIT *100 ,2))  as   field_840740_RATE, ");          
       sqlCmd_Taiwan.append("       decode(FIELD_OVER,0,'N/A',round(FIELD_BACKUP /  a01.FIELD_OVER *100 ,2))  as   field_backup_over_rate, ");          
       sqlCmd_Taiwan.append("       decode(field_840740,0,'N/A',round(FIELD_BACKUP /  field_840740 *100 ,2))  as   field_backup_840740_rate, ");   
       sqlCmd_Taiwan.append("       round(field_CAPTIAL / 1000  ,2)  as   field_CAPTIAL_RATE ");   
       paramList_Taiwan.add(unit);
       paramList_Taiwan.add(unit);
       paramList_Taiwan.add(unit);
       paramList_Taiwan.add(unit);
       paramList_Taiwan.add(unit);
       paramList_Taiwan.add(unit);
       paramList_Taiwan.add(unit);
       //sqlCmd_Taiwan.append("from( select ' '  AS  hsien_id ,  '臺灣省'    AS hsien_name,  '025'  AS FR001W_output_order, ");
       //106.03.07 原台灣省改為其他(含台灣省及福建省.中華民國農會)
       sqlCmd_Taiwan.append("from( select ' '  AS  hsien_id ,  '其他'    AS hsien_name,  '025'  AS FR001W_output_order, ");
       sqlCmd_Taiwan.append("             ' ' AS  bank_no ,     ' '   AS  BANK_NAME, 'Taiwan Province' AS english, ");                  
       sqlCmd_Taiwan.append("             COUNT(*)  AS  COUNT_SEQ,                  'A92'  as  field_SEQ, ");     
       sqlCmd_Taiwan.append("             SUM(field_320300)  field_320300 ,   SUM(field_NET)     field_NET, ");        
       sqlCmd_Taiwan.append("             SUM(field_BACKUP)  field_BACKUP,     SUM(field_OVER)    field_OVER, ");       
       sqlCmd_Taiwan.append("             SUM(field_DEBIT)   field_DEBIT,      SUM(field_CREDIT)  field_CREDIT, ");     
       sqlCmd_Taiwan.append("             SUM(field_840740)  field_840740, ");    
       sqlCmd_Taiwan.append("             decode(SUM(field_910400),0,0,round(SUM(field_910400) * 100000 /SUM(field_910500),0)) as field_CAPTIAL ");  
       sqlCmd_Taiwan.append("       from  ");
       sqlCmd_Taiwan.append("       ( select nvl(cd01.hsien_id,' ')  as  hsien_id , nvl(cd01.hsien_name,'OTHER') as  hsien_name, ");               
       sqlCmd_Taiwan.append("       cd01.FR001W_output_order     as  FR001W_output_order, bn01.bank_no ,  bn01.BANK_NAME, wlx01.english, ");                
       sqlCmd_Taiwan.append("       round(sum(decode(a01.acc_code,'320300',amt,0)) /1,0) as field_320300, "); 
       sqlCmd_Taiwan.append("       round(sum(decode(a01.acc_code,'310000',amt,'320000',amt,0)) /1,0)     as field_NET,  ");  
       sqlCmd_Taiwan.append("       round(sum(decode(a01.acc_code, '120800',amt,'150300',amt,0)) /1,0) as  field_BACKUP, "); 
       sqlCmd_Taiwan.append("       round(sum(decode(a01.acc_code,'990000',amt,0)) /1,0) as field_OVER, ");   
       sqlCmd_Taiwan.append("       round(sum(decode(a01.acc_code,'220000',amt,0)) /1,0) as field_DEBIT, ");  
       sqlCmd_Taiwan.append("       round(sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /1,0) as  field_CREDIT ");  
       sqlCmd_Taiwan.append("       from  (select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y'  and cd01.Hsien_div in ('2','3') ) cd01  ");//106.03.07 fix
       sqlCmd_Taiwan.append("       left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = ? and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL)    ");
       paramList_Taiwan.add(wlx01_m_year);
       sqlCmd_Taiwan.append("       left join bn01 on wlx01.bank_no=bn01.bank_no ");
       if("ALL".equals(bank_type)){
           sqlCmd_Taiwan.append("            and bn01.bank_type in ('6','7') ");
       }else{
           sqlCmd_Taiwan.append("            and bn01.bank_type in (?) ");
           paramList_Taiwan.add(bank_type);
       }
       sqlCmd_Taiwan.append("                 and bn01.m_year =? and bn_type <> '2' and wlx01.m_year=? ");  
       paramList_Taiwan.add(wlx01_m_year);
       paramList_Taiwan.add(wlx01_m_year);
       sqlCmd_Taiwan.append("       left join (select * from a01 where a01.m_year  = ? and a01.m_month  = ? ) a01  ");
       paramList_Taiwan.add(s_year);
       paramList_Taiwan.add(s_month);
       sqlCmd_Taiwan.append("on  bn01.bank_no = a01.bank_code  "); 
       sqlCmd_Taiwan.append("       group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME, wlx01.english ");  
       sqlCmd_Taiwan.append("       ) a01, "); 
       sqlCmd_Taiwan.append("       ( select nvl(cd01.hsien_id,' ')       as  hsien_id , ");              
       sqlCmd_Taiwan.append("       nvl(cd01.hsien_name,'OTHER')  as  hsien_name,            cd01.FR001W_output_order     as  FR001W_output_order, ");             
       sqlCmd_Taiwan.append("       bn01.bank_no as bank_code,  bn01.bank_name, wlx01.english, ");            
       sqlCmd_Taiwan.append("       round(sum(decode(a05.acc_code,'910400',amt,0)) /1,0) as field_910400, ");            
       sqlCmd_Taiwan.append("       round(sum(decode(a05.acc_code,'910500',amt,0)) /1,0) as field_910500  ");      
       sqlCmd_Taiwan.append("       from  (select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y' and cd01.Hsien_div in ('2','3')) cd01 ");//106.03.07 fix       
       sqlCmd_Taiwan.append("       left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = ? and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL)      ");
       paramList_Taiwan.add(wlx01_m_year);
       sqlCmd_Taiwan.append("       left join bn01 on wlx01.bank_no=bn01.bank_no ");
       if("ALL".equals(bank_type)){
           sqlCmd_Taiwan.append("            and bn01.bank_type in ('6','7') ");
       }else{
           sqlCmd_Taiwan.append("            and bn01.bank_type in (?) ");
           paramList_Taiwan.add(bank_type);
       }
       sqlCmd_Taiwan.append("                 and bn01.m_year = ? and bn_type <> '2' and wlx01.m_year = ? ");
       paramList_Taiwan.add(wlx01_m_year);
       paramList_Taiwan.add(wlx01_m_year);
       sqlCmd_Taiwan.append("       left join (select * from a05 where a05.m_year  = ? and a05.m_month  = ?  ");
       paramList_Taiwan.add(s_year);
       paramList_Taiwan.add(s_month);
       sqlCmd_Taiwan.append("and a05.ACC_code in ('910400','910500','91060P') ) a05 on  bn01.bank_no = a05.bank_code  ");     
       sqlCmd_Taiwan.append("       group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME, wlx01.english ");  
       sqlCmd_Taiwan.append("       ) a05  , ");
       sqlCmd_Taiwan.append("       (  select nvl(cd01.hsien_id,' ')       as  hsien_id , ");              
       sqlCmd_Taiwan.append("        nvl(cd01.hsien_name,'OTHER')  as  hsien_name,cd01.FR001W_output_order     as  FR001W_output_order, ");              
       sqlCmd_Taiwan.append("        bn01.bank_no as bank_code,  bn01.bank_name, wlx01.english, ");             
       sqlCmd_Taiwan.append("        round(sum(decode(a02.acc_code,'990611',amt,0)) /1,0) as field_990611  ");     
       sqlCmd_Taiwan.append("        from  (select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y' and cd01.Hsien_div in ('2','3')) cd01 ");//106.03.07 fix       
       sqlCmd_Taiwan.append("        left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = ? and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL) ");
       paramList_Taiwan.add(wlx01_m_year);
       sqlCmd_Taiwan.append("        left join bn01 on wlx01.bank_no=bn01.bank_no "); 
       if("ALL".equals(bank_type)){
           sqlCmd_Taiwan.append("            and bn01.bank_type in ('6','7') ");
       }else{
           sqlCmd_Taiwan.append("            and bn01.bank_type in (?) ");
           paramList_Taiwan.add(bank_type);
       }
       sqlCmd_Taiwan.append("        and bn01.m_year = ? and bn_type <> '2' and wlx01.m_year = ?  ");
       paramList_Taiwan.add(wlx01_m_year);
       paramList_Taiwan.add(wlx01_m_year);
       sqlCmd_Taiwan.append("        left join (select * from a02 where a02.m_year = ? and a02.m_month = ? "); 
       paramList_Taiwan.add(s_year);
       paramList_Taiwan.add(s_month);
       sqlCmd_Taiwan.append("and a02.ACC_code in ('990611') ) a02  on  bn01.bank_no = a02.bank_code ");     
       sqlCmd_Taiwan.append("        group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME, wlx01.english  "); 
       sqlCmd_Taiwan.append("        ) a02 , "); 
       sqlCmd_Taiwan.append("        (  select nvl(cd01.hsien_id,' ') as hsien_id, nvl(cd01.hsien_name,'OTHER') as hsien_name, ");             
       sqlCmd_Taiwan.append("           cd01.FR001W_output_order as FR001W_output_order, bn01.bank_no as bank_code, bn01.bank_name, wlx01.english, ");             
       sqlCmd_Taiwan.append("           round(sum(decode(a04.acc_code,'840740',amt,'840760',amt,0)) /1,0) as field_840740 ");      
       sqlCmd_Taiwan.append("           from  (select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y' and cd01.Hsien_div in ('2','3')) cd01 ");//106.03.07 fix       
       sqlCmd_Taiwan.append("           left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = ? and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL) ");
       paramList_Taiwan.add(wlx01_m_year);
       sqlCmd_Taiwan.append("           left join bn01 on wlx01.bank_no=bn01.bank_no ");
       if("ALL".equals(bank_type)){
           sqlCmd_Taiwan.append("            and bn01.bank_type in ('6','7') ");
       }else{
           sqlCmd_Taiwan.append("            and bn01.bank_type in (?) ");
           paramList_Taiwan.add(bank_type);
       }
       sqlCmd_Taiwan.append("                and bn01.m_year = ? and bn_type <> '2' and wlx01.m_year = ? ");
       paramList_Taiwan.add(wlx01_m_year);
       paramList_Taiwan.add(wlx01_m_year);
       sqlCmd_Taiwan.append("           left join (select * from a04 where a04.m_year  = ? and a04.m_month  = ? ");
       paramList_Taiwan.add(s_year);
       paramList_Taiwan.add(s_month);
       sqlCmd_Taiwan.append("and a04.ACC_code in ('840740','840760') ) a04 on  bn01.bank_no = a04.bank_code ");     
       sqlCmd_Taiwan.append("           group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME, wlx01.english ");  
       sqlCmd_Taiwan.append("        ) a04 ");  
       sqlCmd_Taiwan.append("where  a01.bank_no=a05.bank_code(+) and a01.bank_no=a02.bank_code(+) and a01.bank_no=a04.bank_code(+) and a01.bank_no <> ' ' ");  
       sqlCmd_Taiwan.append(") a01 ");
       List dbData = DBManager.QueryDB_SQLParam(sqlCmd_Taiwan.toString(),paramList_Taiwan,"hsien_id,hsien_name,fr001w_output_order,bank_no,bank_name,english,count_seq,field_seq,field_debit," +
                                                                   		"field_credit,field_over,field_over_rate,field_320300,field_net,field_backup,field_840740,field_840740_rate," +
                                                                   		"field_backup_over_rate,field_backup_840740_rate,field_captial_rate");
       //List dbData = DBManager.QueryDB_SQLParam(sqlCmd_Taiwan.toString(),paramList_Taiwan,"");
       System.out.println("dbData_Taiwan.size()="+dbData.size()); 
       return dbData;
   }
   private static void insertCell(String Item,HSSFWorkbook wb,HSSFRow row,int i,short border){
           HSSFCell cell=(row.getCell((short)i)==null)? row.createCell((short)i) : row.getCell((short)i);
           HSSFCellStyle cs1 = wb.createCellStyle();
           //HSSFCellStyle cs1 = cell.getCellStyle();//會套用原本excel所設定的格式
           //設置邊框
           cs1.setBorderTop(border); //上邊框
           cs1.setBorderBottom(border); //下邊框
           cs1.setBorderLeft(border); //左邊框
           cs1.setBorderRight(border); //右邊框
           cs1.setWrapText(true);//自動換行
           cs1.setAlignment(HSSFCellStyle.ALIGN_CENTER);//水平置中
           cs1.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);//垂直置中
           cell.setCellStyle(cs1);
           cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
           cell.setCellValue(Item);
   }
   private static void setInsertValue(HSSFWorkbook wb,HSSFRow row,HSSFCell cell,String s_year0,DataObject bean,int tmpC,HSSFCellStyle style){
       insertCell("",wb,row,(short)tmpC,HSSFCellStyle.BORDER_THIN);
       cell = row.getCell((short)(tmpC));
       field_debit = bean.getValue("field_debit")==null?"0":bean.getValue("field_debit").toString();
       cell.setCellValue(Utility.setCommaFormat(field_debit));
       cell.setCellStyle(style);
       //放款總額
       tmpC=tmpC+Integer.parseInt(s_year0)+1 ;
       insertCell("",wb,row,(short)tmpC,HSSFCellStyle.BORDER_THIN);
       cell = row.getCell((short)(tmpC));
       field_credit = bean.getValue("field_credit")==null?"0":bean.getValue("field_credit").toString();
       cell.setCellValue(Utility.setCommaFormat(field_credit));
       cell.setCellStyle(style);
       //淨值 
       tmpC=tmpC+Integer.parseInt(s_year0)+1 ;
       insertCell("",wb,row,(short)tmpC,HSSFCellStyle.BORDER_THIN);
       cell = row.getCell((short)(tmpC));
       field_net = bean.getValue("field_net")==null?"0":bean.getValue("field_net").toString();
       cell.setCellValue(Utility.setCommaFormat(field_net));
       cell.setCellStyle(style);
       //本期損益
       tmpC=tmpC+Integer.parseInt(s_year0)+1 ;
       insertCell("",wb,row,(short)tmpC,HSSFCellStyle.BORDER_THIN);
       cell = row.getCell((short)(tmpC));
       field_320300 = bean.getValue("field_320300")==null?"0":bean.getValue("field_320300").toString();
       cell.setCellValue(Utility.setCommaFormat(field_320300));
       cell.setCellStyle(style);
       //狹義逾期放款
       tmpC=tmpC+Integer.parseInt(s_year0)+1 ;
       insertCell("",wb,row,(short)tmpC,HSSFCellStyle.BORDER_THIN);
       cell = row.getCell((short)(tmpC));
       field_over = bean.getValue("field_over")==null?"0":bean.getValue("field_over").toString();
       cell.setCellValue(Utility.setCommaFormat(field_over));
       cell.setCellStyle(style);
       //廣義逾期放款
       tmpC=tmpC+Integer.parseInt(s_year0)+1 ;
       insertCell("",wb,row,(short)tmpC,HSSFCellStyle.BORDER_THIN);
       cell = row.getCell((short)(tmpC));
       field_840740 = bean.getValue("field_840740")==null?"0":bean.getValue("field_840740").toString();
       cell.setCellValue(Utility.setCommaFormat(field_840740));
       cell.setCellStyle(style);
       //狹義逾放比率(狹義逾期放款/放款總額)
       tmpC=tmpC+Integer.parseInt(s_year0)+1 ;
       insertCell("",wb,row,(short)tmpC,HSSFCellStyle.BORDER_THIN);
       cell = row.getCell((short)(tmpC));
       field_over_rate = bean.getValue("field_over_rate") == null?"N/A":String.valueOf(bean.getValue("field_over_rate"));
       field_over_rate = field_over_rate.equals("N/A") ? "N/A" : Utility.setCommaFormat(df_md.format(Double.parseDouble(field_over_rate)));//利率顯示小數點至第2位,不足者補0
       cell.setCellValue(field_over_rate);
       cell.setCellStyle(style);
       //廣義逾放比率(廣義逾期放款/放款總額)
       tmpC=tmpC+Integer.parseInt(s_year0)+1 ;
       insertCell("",wb,row,(short)tmpC,HSSFCellStyle.BORDER_THIN);
       cell = row.getCell((short)(tmpC));
       field_840740_rate = bean.getValue("field_840740_rate") == null?"N/A":String.valueOf(bean.getValue("field_840740_rate"));
       field_840740_rate = field_840740_rate.equals("N/A") ? "N/A" : Utility.setCommaFormat(df_md.format(Double.parseDouble(field_840740_rate)));//利率顯示小數點至第2位,不足者補0
       cell.setCellValue(field_840740_rate);
       cell.setCellStyle(style);
       //備抵呆帳
       tmpC=tmpC+Integer.parseInt(s_year0)+1 ;
       insertCell("",wb,row,(short)tmpC,HSSFCellStyle.BORDER_THIN);
       cell = row.getCell((short)(tmpC));
       field_backup = bean.getValue("field_backup")==null?"0":bean.getValue("field_backup").toString();
       cell.setCellValue(Utility.setCommaFormat(field_backup));
       cell.setCellStyle(style);
       //備呆占狹義逾期放款比率(備抵呆帳/狹義逾放)
       tmpC=tmpC+Integer.parseInt(s_year0)+1 ;
       insertCell("",wb,row,(short)tmpC,HSSFCellStyle.BORDER_THIN);
       cell = row.getCell((short)(tmpC));
       field_backup_over_rate = bean.getValue("field_backup_over_rate") == null?"N/A":String.valueOf(bean.getValue("field_backup_over_rate"));
       field_backup_over_rate = field_backup_over_rate.equals("N/A") ? "N/A" : Utility.setCommaFormat(df_md.format(Double.parseDouble(field_backup_over_rate)));//利率顯示小數點至第2位,不足者補0
       cell.setCellValue(field_backup_over_rate);
       cell.setCellStyle(style);
       //備呆占廣義逾期放款比率(備抵呆帳/廣義逾放)     
       tmpC=tmpC+Integer.parseInt(s_year0)+1 ;
       insertCell("",wb,row,(short)tmpC,HSSFCellStyle.BORDER_THIN);
       cell = row.getCell((short)(tmpC));
       field_backup_840740_rate = bean.getValue("field_backup_840740_rate") == null?"N/A":String.valueOf(bean.getValue("field_backup_840740_rate"));
       field_backup_840740_rate = field_backup_840740_rate.equals("N/A") ? "N/A" : Utility.setCommaFormat(df_md.format(Double.parseDouble(field_backup_840740_rate)));//利率顯示小數點至第2位,不足者補0
       cell.setCellValue(field_backup_840740_rate);
       cell.setCellStyle(style);
       //淨值佔風險性資產比率     
       tmpC=tmpC+Integer.parseInt(s_year0)+1 ;
       insertCell("",wb,row,(short)tmpC,HSSFCellStyle.BORDER_THIN);
       cell = row.getCell((short)(tmpC));
       field_captial_rate = bean.getValue("field_captial_rate") == null?"0.00":String.valueOf(bean.getValue("field_captial_rate"));
       field_captial_rate = field_captial_rate.equals("N/A") ? "N/A" : Utility.setCommaFormat(df_md.format(Double.parseDouble(field_captial_rate)));//利率顯示小數點至第2位,不足者補0
       cell.setCellValue(field_captial_rate);
       cell.setCellStyle(style);
   }
   
   }

