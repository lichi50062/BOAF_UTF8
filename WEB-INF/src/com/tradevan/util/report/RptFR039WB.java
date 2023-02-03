/*
 * Created on 2006/07/26-28 by 2295
 * 99.09.13-14 fix 根據查詢年度.100年以後取得新縣市別.100年以前取得舊縣市別 
  			        使用PreparedStatement;並列印轉換後的SQL;套用QueryDB_SQLParam by 2295
 * 102.06.21 1.add 103年後SQL  2.總表與明細表程式分開 by 2968	
 * 102.09.02 fix 報表最下方顯示所有直轄市小計      	
 * 103.01.27 add 臺灣省改其他,並增加說明 by 2295    	        
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

public class RptFR039WB {
    public static String createRpt(String s_year,String s_month,String unit,String bank_type){
        String errMsg = "";
        StringBuffer sqlCmd_rptStyle1 = new StringBuffer();//明細表組合sql
        StringBuffer sqlCmd_Taiwan = new StringBuffer();//台灣省小計
        List dbData_detail = null;//明細表.detail
        List dbData_Taiwan = null;//明細表.台灣省小計
        List dbData_detail_hsien = new LinkedList();//存款餘額表.新北市.台北市.桃園市.台中市.台南市.高雄市小計
        List paramList_rptStyle1 = new ArrayList();//明細表參數
        List paramList_Taiwan = new ArrayList();
        String field_seq = "";
        String hsien_name = "";
        String bank_no = "";
        String bank_name = "";
        String count_seq = "";
        String field_sumtotal = "";//存款餘額
        String field_sum1= "";//小計
        String field_220100 = "";//支票存款(含本會支票)
        String field_220300 = "";//活期存款
        String field_220400 = "";//活期儲蓄存款(含員工活儲)
        String field_sum2 = "";//小計
        String field_220600 = "";//定期存款
        String field_220700 = "";//定期儲蓄存款
        String field_220900 = "";//公庫存款
        String rpt2field_sumtotal = "";//放款總額
        String field_120200 = "";//貼現
        String field_120101 = "";//般放款及透支
        String field_120401 = "";//統一漁貸
        String field_120501 = ""; //專案放款
        String field_120601 = "";//農業發展基金放款
        String field_120700 = "";//內部融資
        String field_150200 = "";//催收款
        String[] dataIdx0 = {"field_sumtotal","field_sum1","field_220100","field_220300","field_220400","field_sum2","field_220600","field_220700","field_220900"};
        					//存款餘額           //小計        //支票存款     //活期存款        //活期儲蓄存款  //小計         //定期存款        //定期儲蓄存款   //公庫存款
        String[] dataIdx1 = {"rpt2field_sumtotal","field_120200","field_120101","field_120401","field_120501","field_120601","field_120700","field_150200"};
        					//放款總額              //貼現           //一般放款及透支//統一漁貸        //專案放款       //農業發展基金放款//內部融資     //催收款
        int[] sheetCell = {11,10};
        String bank_type_name="";
        String cd01_table = "";
        String wlx01_m_year = "";
        //99.09.13 add 查詢年度100年以前.縣市別不同===============================
	    cd01_table = (Integer.parseInt(s_year) < 100)?"cd01_99":"";
	    wlx01_m_year = (Integer.parseInt(s_year) < 100)?"99":"100";
	    //=====================================================================
        System.out.println("RptFR039WB.bank_type="+bank_type);
        if(bank_type.equals("ALL")){
           bank_type_name = "農漁會";
        }else{
           bank_type_name = (bank_type.equals("6"))?"農會":"漁會";
        }
        String unit_name = Utility.getUnitName(unit);

      int rowNum=0;
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
                String openfile="";
                String filename="";
                
                openfile = "農漁會信用部存放款餘額表_明細表.xls";
                filename = "農漁會信用部存放款餘額表_明細表.xls";

                System.out.println("開啟檔:" + openfile);
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
                ps.setScale( ( short )76 ); //列印縮放百分比

                ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
                //wb.setSheetName(0,"test");
                //設定表頭 為固定 先設欄的起始再設列的起始
                wb.setRepeatingRowsAndColumns(0, 0, 10, 0, 4);
                
                finput.close();

                HSSFRow row=null;//宣告一列
                HSSFCell cell=null;//宣告一個儲存格

                //明細表組合sql
                sqlCmd_rptStyle1.append(" select  hsien_id , hsien_name, FR001W_output_order,"); 
                sqlCmd_rptStyle1.append("         bank_no ,  BANK_NAME,"); 
                sqlCmd_rptStyle1.append("         COUNT_SEQ,");
                sqlCmd_rptStyle1.append("         field_SEQ,");
                sqlCmd_rptStyle1.append("         field_sumtotal ,");//--存款餘額
                sqlCmd_rptStyle1.append("         field_sum1 ,");//--小計
                sqlCmd_rptStyle1.append("         field_220100 ,");//--支票存款(含本會支票)
                sqlCmd_rptStyle1.append("         field_220300 ,");//--活期存款
                sqlCmd_rptStyle1.append("         field_220400 ,");//--活期儲蓄存款(含員工活儲)
                sqlCmd_rptStyle1.append("         field_sum2 ,");//--小計
                sqlCmd_rptStyle1.append("         field_220600 ,");//--定期存款
                sqlCmd_rptStyle1.append("         field_220700 ,");//--定期儲蓄存款
                sqlCmd_rptStyle1.append("         field_220900 ,");//--公庫存款
                sqlCmd_rptStyle1.append("         rpt2field_sumtotal ,");//--放款總額
                sqlCmd_rptStyle1.append("         field_120200 ,");//--貼現
                sqlCmd_rptStyle1.append("         field_120101 ,");//--一般放款及透支
                sqlCmd_rptStyle1.append("         field_120401 ,");//--統一漁貸
                sqlCmd_rptStyle1.append("         field_120501 ,");//--專案放款
                sqlCmd_rptStyle1.append("         field_120601 ,");//--農業發展基金放款
                sqlCmd_rptStyle1.append("         field_120700 ,");//--內部融資
                sqlCmd_rptStyle1.append("         field_150200 ");//--催收款
                sqlCmd_rptStyle1.append(" from ( ");
                                                  //明細表用.總計
                sqlCmd_rptStyle1.append("         select  ' ' as hsien_id ,'總 計 ' as hsien_name,");
                sqlCmd_rptStyle1.append("                 '001' as FR001W_output_order,");
                sqlCmd_rptStyle1.append("                 ' ' as bank_no ,  ' ' as BANK_NAME,");
                sqlCmd_rptStyle1.append("                 sum(count_seq) as count_seq,");
                sqlCmd_rptStyle1.append("                 'A99' as field_seq,");
                sqlCmd_rptStyle1.append("                 round(sum(field_sumtotal)/?,0) field_sumtotal,");
                sqlCmd_rptStyle1.append("                 round(sum(field_sum1)/?,0) field_sum1,"); 
                sqlCmd_rptStyle1.append("                 round(sum(field_220100)/?,0) field_220100,"); 
                sqlCmd_rptStyle1.append("                 round(sum(field_220300)/?,0) field_220300,"); 
                sqlCmd_rptStyle1.append("                 round(sum(field_220400)/?,0) field_220400,"); 
                sqlCmd_rptStyle1.append("                 round(sum(field_sum2)/?,0) field_sum2,"); 
                sqlCmd_rptStyle1.append("                 round(sum(field_220600)/?,0) field_220600,"); 
                sqlCmd_rptStyle1.append("                 round(sum(field_220700)/?,0) field_220700,"); 
                sqlCmd_rptStyle1.append("                 round(sum(field_220900)/?,0) field_220900,"); 
                sqlCmd_rptStyle1.append("                 round(sum(rpt2field_sumtotal)/?,0) rpt2field_sumtotal,"); 
                sqlCmd_rptStyle1.append("                 round(sum(field_120200)/?,0) field_120200,"); 
                sqlCmd_rptStyle1.append("                 round(sum(field_120101)/?,0) field_120101,"); 
                sqlCmd_rptStyle1.append("                 round(sum(field_120401)/?,0) field_120401,"); 
                sqlCmd_rptStyle1.append("                 round(sum(field_120501)/?,0) field_120501,"); 
                sqlCmd_rptStyle1.append("                 round(sum(field_120601)/?,0) field_120601,"); 
                sqlCmd_rptStyle1.append("                 round(sum(field_120700)/?,0) field_120700,"); 
                sqlCmd_rptStyle1.append("                 round(sum(field_150200)/?,0) field_150200 "); 
                for(int k=1;k<=17;k++){
                    paramList_rptStyle1.add(unit);
                }
                sqlCmd_rptStyle1.append("         from ");
                sqlCmd_rptStyle1.append("             ( ");
                sqlCmd_rptStyle1.append("                 select hsien_id , ");
                sqlCmd_rptStyle1.append("                        hsien_name, ");
                sqlCmd_rptStyle1.append("                        FR001W_output_order,");
                sqlCmd_rptStyle1.append("                        count(*) as count_seq,");
                sqlCmd_rptStyle1.append("                        bank_no ,  BANK_NAME,");
                sqlCmd_rptStyle1.append("                        sum(field_sumtotal) as field_sumtotal,");
                sqlCmd_rptStyle1.append("                        sum(field_sum1)     as field_sum1, ");
                sqlCmd_rptStyle1.append("                        sum(field_220100)   as field_220100,");
                sqlCmd_rptStyle1.append("                        sum(field_220300)     as field_220300,");
                sqlCmd_rptStyle1.append("                        sum(field_220400)     as field_220400,");
                sqlCmd_rptStyle1.append("                        sum(field_sum2)  as field_sum2,");
                sqlCmd_rptStyle1.append("                        sum(field_220600)     as field_220600,");
                sqlCmd_rptStyle1.append("                        sum(field_220700)     as field_220700,");
                sqlCmd_rptStyle1.append("                        sum(field_220900)     as field_220900,");
                sqlCmd_rptStyle1.append("                        sum(rpt2field_sumtotal)     as rpt2field_sumtotal,");
                sqlCmd_rptStyle1.append("                        sum(field_120200)     as field_120200,");
                sqlCmd_rptStyle1.append("                        sum(field_120101)     as field_120101,");
                sqlCmd_rptStyle1.append("                        sum(field_120401)     as field_120401,");
                sqlCmd_rptStyle1.append("                        sum(field_120501)     as field_120501,");
                sqlCmd_rptStyle1.append("                        sum(field_120601)     as field_120601,");
                sqlCmd_rptStyle1.append("                        sum(field_120700)     as field_120700,");
                sqlCmd_rptStyle1.append("                        sum(field_150200)     as field_150200 ");
                sqlCmd_rptStyle1.append("                 from ( ");
                sqlCmd_rptStyle1.append("                         select nvl(cd01.hsien_id,' ') as  hsien_id , ");
                sqlCmd_rptStyle1.append("                                nvl(cd01.hsien_name,'OTHER') as  hsien_name, ");
                sqlCmd_rptStyle1.append("                                cd01.FR001W_output_order     as  FR001W_output_order,");
                sqlCmd_rptStyle1.append("                                bn01.bank_no ,  bn01.BANK_NAME,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'220100',amt,'221000',amt,'220300',amt,'220400',amt,'220500',amt,'220600',amt,'220700',amt,'220900',amt,0))     as field_sumtotal,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'220100',amt,'221000',amt,'220300',amt,'220400',amt,'220500',amt,0))     as field_sum1, ");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'220100',amt,'221000',amt,0))     as field_220100,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'220300',amt,0))     as field_220300,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'220400',amt,'220500',amt,0))     as field_220400,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'220600',amt,'220700',amt,'220900',amt,0))  as field_sum2,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'220600',amt,0))     as field_220600,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'220700',amt,0))     as field_220700,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'220900',amt,0))     as field_220900,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0))     as rpt2field_sumtotal,");
                sqlCmd_rptStyle1.append("                                decode(YEAR_TYPE,'102',decode(bank_type,'6',sum(decode(a01.acc_code,'120200',amt,0)),'7',sum(decode(a01.acc_code,'120300',amt,0)),0),'103',sum(decode(a01.acc_code,'120200',amt,0)),0) as  field_120200, ");
                sqlCmd_rptStyle1.append("                                decode(YEAR_TYPE,'102',decode(bank_type,'6',sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120301',amt,'120302',amt,0)),'7',sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120401',amt,'120402',amt,0)),0),'103',sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120301',amt,'120302',amt,0)),0 ) as  field_120101,");   
                sqlCmd_rptStyle1.append("                                decode(YEAR_TYPE,'102',decode(bank_type,'6',sum(decode(a01.acc_code,'120401',amt,'120402',amt,0)),'7',sum(decode(a01.acc_code,'120201',amt,'120202',amt,0)),0),'103',sum(decode(a01.acc_code,'120401',amt,'120402',amt,0)),0 ) as  field_120401, ");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'120501',amt,'120502',amt,0))     as field_120501,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'120601',amt,'120602',amt,'120603',amt,'120604',amt,0))     as field_120601,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'120700',amt,0))     as field_120700,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'150200',amt,0))     as field_150200 ");
                sqlCmd_rptStyle1.append("                         from( select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y') cd01 ");
                sqlCmd_rptStyle1.append("                         left join ( select bank_type,bn01.bank_no,bank_name,hsien_id  ");
                sqlCmd_rptStyle1.append("                                     from (select * from bn01 where m_year=?   ");
                paramList_rptStyle1.add(wlx01_m_year);
                if(bank_type.equals("ALL")){
                    sqlCmd_rptStyle1.append("                                     and bank_type in('6','7') )bn01    ");
                }else{
                    sqlCmd_rptStyle1.append("                                     and bank_type in(?) )bn01    ");
                    paramList_rptStyle1.add(bank_type);
                }
                sqlCmd_rptStyle1.append("                                     left join (select * from wlx01 where m_year=?)wlx01 on bn01.bank_no = wlx01.bank_no");
                sqlCmd_rptStyle1.append("                                    )bn01 on bn01.hsien_id=cd01.hsien_id ");
                sqlCmd_rptStyle1.append("                         left join (select (CASE WHEN (a01.m_year <= 102) THEN '102'");
                sqlCmd_rptStyle1.append("                                                 WHEN (a01.m_year > 102) THEN '103'");
                sqlCmd_rptStyle1.append("                                            ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01 ");
                sqlCmd_rptStyle1.append("                                    where m_year=? and m_month=?  ");
                paramList_rptStyle1.add(wlx01_m_year);
                paramList_rptStyle1.add(s_year);
                paramList_rptStyle1.add(s_month);
                sqlCmd_rptStyle1.append("                                    and a01.acc_code in ('220100','221000','220300','220400','220500','220600','220700','220900','120000','120800','150300','120200','120300','120101','120102','120301','120302','120401','120402','120201','120202','120501','120502','120601','120602','120603','120604','120700','150200') ");
                sqlCmd_rptStyle1.append("                                    )a01  on  bn01.bank_no = a01.bank_code ");
                sqlCmd_rptStyle1.append("                         group by YEAR_TYPE,bank_type,nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no ,  bn01.BANK_NAME ");
                sqlCmd_rptStyle1.append("                      )a01 ");
                sqlCmd_rptStyle1.append("                 where a01.bank_no <> ' ' ");
                sqlCmd_rptStyle1.append("                group by  hsien_id ,   hsien_name,  FR001W_output_order,bank_no,bank_name ");
                sqlCmd_rptStyle1.append("             ) ");
                sqlCmd_rptStyle1.append("         UNION ALL");
                                                  //明細
                sqlCmd_rptStyle1.append("         select a01.hsien_id ,a01.hsien_name,");
                sqlCmd_rptStyle1.append("                a01.FR001W_output_order,");
                sqlCmd_rptStyle1.append("                a01.bank_no ,  a01.BANK_NAME,");
                sqlCmd_rptStyle1.append("                1  AS  COUNT_SEQ,");
                sqlCmd_rptStyle1.append("                'A01'  as  field_SEQ,");
                sqlCmd_rptStyle1.append("                round(field_sumtotal/?,0) field_sumtotal, ");
                sqlCmd_rptStyle1.append("                round(field_sum1/?,0) field_sum1, ");
                sqlCmd_rptStyle1.append("                round(field_220100/?,0) field_220100, ");
                sqlCmd_rptStyle1.append("                round(field_220300/?,0) field_220300, ");
                sqlCmd_rptStyle1.append("                round(field_220400/?,0) field_220400, ");
                sqlCmd_rptStyle1.append("                round(field_sum2/?,0) field_sum2, ");
                sqlCmd_rptStyle1.append("                round(field_220600/?,0) field_220600, ");
                sqlCmd_rptStyle1.append("                round(field_220700/?,0) field_220700, ");
                sqlCmd_rptStyle1.append("                round(field_220900/?,0) field_220900, ");
                sqlCmd_rptStyle1.append("                round(rpt2field_sumtotal/?,0) rpt2field_sumtotal, ");
                sqlCmd_rptStyle1.append("                round(field_120200/?,0) field_120200, ");
                sqlCmd_rptStyle1.append("                round(field_120101/?,0) field_120101, ");
                sqlCmd_rptStyle1.append("                round(field_120401/?,0) field_120401, ");
                sqlCmd_rptStyle1.append("                round(field_120501/?,0) field_120501, ");
                sqlCmd_rptStyle1.append("                round(field_120601/?,0) field_120601, ");
                sqlCmd_rptStyle1.append("                round(field_120700/?,0) field_120700, ");
                sqlCmd_rptStyle1.append("                round(field_150200/?,0) field_150200  ");
                for(int k=1;k<=17;k++){
                    paramList_rptStyle1.add(unit);
                }
                sqlCmd_rptStyle1.append("         from ");
                sqlCmd_rptStyle1.append("             ( ");
                sqlCmd_rptStyle1.append("                 select  hsien_id , ");
                sqlCmd_rptStyle1.append("                         hsien_name, ");
                sqlCmd_rptStyle1.append("                         FR001W_output_order,");
                sqlCmd_rptStyle1.append("                         bank_no ,  BANK_NAME,");
                sqlCmd_rptStyle1.append("                         sum(field_sumtotal) as field_sumtotal,");
                sqlCmd_rptStyle1.append("                         sum(field_sum1)     as field_sum1,    ");
                sqlCmd_rptStyle1.append("                         sum(field_220100)   as field_220100,  ");
                sqlCmd_rptStyle1.append("                         sum(field_220300)     as field_220300,");
                sqlCmd_rptStyle1.append("                         sum(field_220400)     as field_220400,");
                sqlCmd_rptStyle1.append("                         sum(field_sum2)  as field_sum2,       ");
                sqlCmd_rptStyle1.append("                         sum(field_220600)     as field_220600,");
                sqlCmd_rptStyle1.append("                         sum(field_220700)     as field_220700,");
                sqlCmd_rptStyle1.append("                         sum(field_220900)     as field_220900,");
                sqlCmd_rptStyle1.append("                         sum(rpt2field_sumtotal)     as rpt2field_sumtotal,");
                sqlCmd_rptStyle1.append("                         sum(field_120200)     as field_120200,");
                sqlCmd_rptStyle1.append("                         sum(field_120101)     as field_120101,");
                sqlCmd_rptStyle1.append("                         sum(field_120401)     as field_120401,");
                sqlCmd_rptStyle1.append("                         sum(field_120501)     as field_120501,");
                sqlCmd_rptStyle1.append("                         sum(field_120601)     as field_120601,");
                sqlCmd_rptStyle1.append("                         sum(field_120700)     as field_120700,");
                sqlCmd_rptStyle1.append("                         sum(field_150200)     as field_150200 ");   
                sqlCmd_rptStyle1.append("                 from ( ");
                sqlCmd_rptStyle1.append("                         select nvl(cd01.hsien_id,' ') as  hsien_id ,"); 
                sqlCmd_rptStyle1.append("                                nvl(cd01.hsien_name,'OTHER') as  hsien_name,");
                sqlCmd_rptStyle1.append("                                cd01.FR001W_output_order     as  FR001W_output_order,");
                sqlCmd_rptStyle1.append("                                bn01.bank_no ,  bn01.BANK_NAME,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'220100',amt,'221000',amt,'220300',amt,'220400',amt,'220500',amt,'220600',amt,'220700',amt,'220900',amt,0))     as field_sumtotal,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'220100',amt,'221000',amt,'220300',amt,'220400',amt,'220500',amt,0))     as field_sum1, ");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'220100',amt,'221000',amt,0))     as field_220100,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'220300',amt,0))     as field_220300,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'220400',amt,'220500',amt,0))     as field_220400,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'220600',amt,'220700',amt,'220900',amt,0))  as field_sum2,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'220600',amt,0))     as field_220600,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'220700',amt,0))     as field_220700,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'220900',amt,0))     as field_220900,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0))     as rpt2field_sumtotal,");
                sqlCmd_rptStyle1.append("                                decode(YEAR_TYPE,'102',decode(bank_type,'6',sum(decode(a01.acc_code,'120200',amt,0)),'7',sum(decode(a01.acc_code,'120300',amt,0)),0),'103',sum(decode(a01.acc_code,'120200',amt,0)),0) as  field_120200, ");
                sqlCmd_rptStyle1.append("                                decode(YEAR_TYPE,'102',decode(bank_type,'6',sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120301',amt,'120302',amt,0)),'7',sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120401',amt,'120402',amt,0)),0),'103',sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120301',amt,'120302',amt,0)),0 ) as  field_120101, ");  
                sqlCmd_rptStyle1.append("                                decode(YEAR_TYPE,'102',decode(bank_type,'6',sum(decode(a01.acc_code,'120401',amt,'120402',amt,0)),'7',sum(decode(a01.acc_code,'120201',amt,'120202',amt,0)),0),'103',sum(decode(a01.acc_code,'120401',amt,'120402',amt,0)),0 ) as  field_120401, ");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'120501',amt,'120502',amt,0))     as field_120501,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'120601',amt,'120602',amt,'120603',amt,'120604',amt,0))     as field_120601,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'120700',amt,0))     as field_120700,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'150200',amt,0))     as field_150200");
                sqlCmd_rptStyle1.append("                         from( select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y') cd01 ");
                sqlCmd_rptStyle1.append("                         left join ( select bank_type,bn01.bank_no,bank_name,hsien_id  ");
                sqlCmd_rptStyle1.append("                                     from (select * from bn01 where m_year=?  ");
                paramList_rptStyle1.add(wlx01_m_year);
                if(bank_type.equals("ALL")){
                    sqlCmd_rptStyle1.append("                                     and bank_type in('6','7') )bn01    ");
                }else{
                    sqlCmd_rptStyle1.append("                                     and bank_type in(?) )bn01    ");
                    paramList_rptStyle1.add(bank_type);
                }
                sqlCmd_rptStyle1.append("                                     left join (select * from wlx01 where m_year=?)wlx01 on bn01.bank_no = wlx01.bank_no ");
                sqlCmd_rptStyle1.append("                                    )bn01 on bn01.hsien_id=cd01.hsien_id ");
                sqlCmd_rptStyle1.append("                         left join (select (CASE WHEN (a01.m_year <= 102) THEN '102' ");
                sqlCmd_rptStyle1.append("                                                 WHEN (a01.m_year > 102) THEN '103' ");
                sqlCmd_rptStyle1.append("                                            ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01 ");
                sqlCmd_rptStyle1.append("                                    where m_year=? and m_month=? "); 
                paramList_rptStyle1.add(wlx01_m_year);
                paramList_rptStyle1.add(s_year);
                paramList_rptStyle1.add(s_month);
                sqlCmd_rptStyle1.append("                                    and a01.acc_code in ('220100','221000','220300','220400','220500','220600','220700','220900','120000','120800','150300','120200','120300','120101','120102','120301','120302','120401','120402','120201','120202','120501','120502','120601','120602','120603','120604','120700','150200') ");               
                sqlCmd_rptStyle1.append("                                    )a01  on  bn01.bank_no = a01.bank_code ");
                sqlCmd_rptStyle1.append("                         group by YEAR_TYPE,bank_type,nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no ,  bn01.BANK_NAME ");
                sqlCmd_rptStyle1.append("                     )a01 ");
                sqlCmd_rptStyle1.append("         where a01.bank_no <> ' ' ");
                sqlCmd_rptStyle1.append("         group by a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order,a01.bank_no ,a01.BANK_NAME ");
                sqlCmd_rptStyle1.append("        )a01 ");
                sqlCmd_rptStyle1.append("         UNION ALL ");
                                                  //各縣市小計
                sqlCmd_rptStyle1.append("         select  a01.hsien_id, a01.hsien_name,");
                sqlCmd_rptStyle1.append("                 a01.FR001W_output_order,");
                sqlCmd_rptStyle1.append("                 ' ' AS bank_no , ' ' AS BANK_NAME,");
                sqlCmd_rptStyle1.append("                 COUNT_SEQ,");
                sqlCmd_rptStyle1.append("                 'A90'  as  field_SEQ,");
                sqlCmd_rptStyle1.append("                 round(field_sumtotal/?,0) field_sumtotal,");
                sqlCmd_rptStyle1.append("                 round(field_sum1/?,0) field_sum1,");
                sqlCmd_rptStyle1.append("                 round(field_220100/?,0) field_220100,");
                sqlCmd_rptStyle1.append("                 round(field_220300/?,0) field_220300,");
                sqlCmd_rptStyle1.append("                 round(field_220400/?,0) field_220400,");
                sqlCmd_rptStyle1.append("                 round(field_sum2/?,0) field_sum2,");
                sqlCmd_rptStyle1.append("                 round(field_220600/?,0) field_220600,"); 
                sqlCmd_rptStyle1.append("                 round(field_220700/?,0) field_220700,"); 
                sqlCmd_rptStyle1.append("                 round(field_220900/?,0) field_220900,"); 
                sqlCmd_rptStyle1.append("                 round(rpt2field_sumtotal/?,0) rpt2field_sumtotal,"); 
                sqlCmd_rptStyle1.append("                 round(field_120200/?,0) field_120200,"); 
                sqlCmd_rptStyle1.append("                 round(field_120101/?,0) field_120101,"); 
                sqlCmd_rptStyle1.append("                 round(field_120401/?,0) field_120401,"); 
                sqlCmd_rptStyle1.append("                 round(field_120501/?,0) field_120501,"); 
                sqlCmd_rptStyle1.append("                 round(field_120601/?,0) field_120601,"); 
                sqlCmd_rptStyle1.append("                 round(field_120700/?,0) field_120700,"); 
                sqlCmd_rptStyle1.append("                 round(field_150200/?,0) field_150200 "); 
                for(int k=1;k<=17;k++){
                    paramList_rptStyle1.add(unit);
                }
                sqlCmd_rptStyle1.append("         from ");
                sqlCmd_rptStyle1.append("             ("); 
                sqlCmd_rptStyle1.append("                 select hsien_id ,"); 
                sqlCmd_rptStyle1.append("                        hsien_name,"); 
                sqlCmd_rptStyle1.append("                        FR001W_output_order,");
                sqlCmd_rptStyle1.append("                        COUNT(*)  AS  COUNT_SEQ,");
                sqlCmd_rptStyle1.append("                        sum(field_sumtotal) as field_sumtotal,");
                sqlCmd_rptStyle1.append("                        sum(field_sum1)     as field_sum1,    ");
                sqlCmd_rptStyle1.append("                        sum(field_220100)   as field_220100,  ");
                sqlCmd_rptStyle1.append("                        sum(field_220300)     as field_220300,");
                sqlCmd_rptStyle1.append("                        sum(field_220400)     as field_220400,");
                sqlCmd_rptStyle1.append("                        sum(field_sum2)  as field_sum2,       ");
                sqlCmd_rptStyle1.append("                        sum(field_220600)     as field_220600,");
                sqlCmd_rptStyle1.append("                        sum(field_220700)     as field_220700,");
                sqlCmd_rptStyle1.append("                        sum(field_220900)     as field_220900,");
                sqlCmd_rptStyle1.append("                        sum(rpt2field_sumtotal)     as rpt2field_sumtotal,");
                sqlCmd_rptStyle1.append("                        sum(field_120200)     as field_120200,");
                sqlCmd_rptStyle1.append("                        sum(field_120101)     as field_120101,");
                sqlCmd_rptStyle1.append("                        sum(field_120401)     as field_120401,");
                sqlCmd_rptStyle1.append("                        sum(field_120501)     as field_120501,");
                sqlCmd_rptStyle1.append("                        sum(field_120601)     as field_120601,");
                sqlCmd_rptStyle1.append("                        sum(field_120700)     as field_120700,");
                sqlCmd_rptStyle1.append("                        sum(field_150200)     as field_150200 ");   
                sqlCmd_rptStyle1.append("                 from (");
                sqlCmd_rptStyle1.append("                         select nvl(cd01.hsien_id,' ') as  hsien_id , ");
                sqlCmd_rptStyle1.append("                                nvl(cd01.hsien_name,'OTHER') as  hsien_name, ");
                sqlCmd_rptStyle1.append("                                cd01.FR001W_output_order     as  FR001W_output_order,");
                sqlCmd_rptStyle1.append("                                bn01.bank_no ,  bn01.BANK_NAME,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'220100',amt,'221000',amt,'220300',amt,'220400',amt,'220500',amt,'220600',amt,'220700',amt,'220900',amt,0))     as field_sumtotal,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'220100',amt,'221000',amt,'220300',amt,'220400',amt,'220500',amt,0))     as field_sum1, ");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'220100',amt,'221000',amt,0))     as field_220100,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'220300',amt,0))     as field_220300,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'220400',amt,'220500',amt,0))     as field_220400,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'220600',amt,'220700',amt,'220900',amt,0))  as field_sum2,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'220600',amt,0))     as field_220600,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'220700',amt,0))     as field_220700,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'220900',amt,0))     as field_220900,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0))     as rpt2field_sumtotal,");
                sqlCmd_rptStyle1.append("                                decode(YEAR_TYPE,'102',decode(bank_type,'6',sum(decode(a01.acc_code,'120200',amt,0)),'7',sum(decode(a01.acc_code,'120300',amt,0)),0),'103',sum(decode(a01.acc_code,'120200',amt,0)),0) as  field_120200, ");
                sqlCmd_rptStyle1.append("                                decode(YEAR_TYPE,'102',decode(bank_type,'6',sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120301',amt,'120302',amt,0)),'7',sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120401',amt,'120402',amt,0)),0),'103',sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120301',amt,'120302',amt,0)),0 ) as  field_120101, ");  
                sqlCmd_rptStyle1.append("                                decode(YEAR_TYPE,'102',decode(bank_type,'6',sum(decode(a01.acc_code,'120401',amt,'120402',amt,0)),'7',sum(decode(a01.acc_code,'120201',amt,'120202',amt,0)),0),'103',sum(decode(a01.acc_code,'120401',amt,'120402',amt,0)),0 ) as  field_120401, ");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'120501',amt,'120502',amt,0))     as field_120501,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'120601',amt,'120602',amt,'120603',amt,'120604',amt,0))     as field_120601,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'120700',amt,0))     as field_120700,");
                sqlCmd_rptStyle1.append("                                sum(decode(a01.acc_code,'150200',amt,0))     as field_150200 ");
                sqlCmd_rptStyle1.append("                         from( select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y') cd01 ");
                sqlCmd_rptStyle1.append("                         left join ( select bank_type,bn01.bank_no,bank_name,hsien_id  ");
                sqlCmd_rptStyle1.append("                                     from (select * from bn01 where m_year=?   ");
                paramList_rptStyle1.add(wlx01_m_year);
                if(bank_type.equals("ALL")){
                    sqlCmd_rptStyle1.append("                                     and bank_type in('6','7') )bn01    ");
                }else{
                    sqlCmd_rptStyle1.append("                                     and bank_type in(?) )bn01    ");
                    paramList_rptStyle1.add(bank_type);
                }
                sqlCmd_rptStyle1.append("                                     left join (select * from wlx01 where m_year=?)wlx01 on bn01.bank_no = wlx01.bank_no");
                sqlCmd_rptStyle1.append("                                    )bn01 on bn01.hsien_id=cd01.hsien_id ");
                sqlCmd_rptStyle1.append("                         left join (select (CASE WHEN (a01.m_year <= 102) THEN '102' ");
                sqlCmd_rptStyle1.append("                                                 WHEN (a01.m_year > 102) THEN '103'  ");
                sqlCmd_rptStyle1.append("                                            ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01 ");
                sqlCmd_rptStyle1.append("                                    where m_year=? and m_month=? ");
                paramList_rptStyle1.add(wlx01_m_year);
                paramList_rptStyle1.add(s_year);
                paramList_rptStyle1.add(s_month);
                sqlCmd_rptStyle1.append("                                    and a01.acc_code in ('220100','221000','220300','220400','220500','220600','220700','220900','120000','120800','150300','120200','120300','120101','120102','120301','120302','120401','120402','120201','120202','120501','120502','120601','120602','120603','120604','120700','150200') ");
                sqlCmd_rptStyle1.append("                                    )a01  on  bn01.bank_no = a01.bank_code ");
                sqlCmd_rptStyle1.append("                         group by YEAR_TYPE,bank_type,nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no ,  bn01.BANK_NAME ");
                sqlCmd_rptStyle1.append("                     )a01 ");
                sqlCmd_rptStyle1.append("         where a01.bank_no <> ' ' ");
                sqlCmd_rptStyle1.append("         group by a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order ");
                sqlCmd_rptStyle1.append("         )a01 ");
                sqlCmd_rptStyle1.append("     )  a01  ORDER by    FR001W_output_order, field_SEQ,  hsien_id ,  bank_no ");
                
                //台灣省資料
                sqlCmd_Taiwan.append(" select ' '  AS  hsien_id ,  '臺灣省'   AS hsien_name,  '025'  AS FR001W_output_order, ");
                sqlCmd_Taiwan.append("        ' ' AS  bank_no ,     ' '   AS  BANK_NAME, ");
                sqlCmd_Taiwan.append("        sum(COUNT_SEQ) as count_seq, ");
                sqlCmd_Taiwan.append("        'A92'  as  field_SEQ, ");
                sqlCmd_Taiwan.append("        round(sum(field_sumtotal)/?,0) field_sumtotal, ");
                sqlCmd_Taiwan.append("        round(sum(field_sum1)/?,0) field_sum1,         ");
                sqlCmd_Taiwan.append("        round(sum(field_220100)/?,0) field_220100,     ");
                sqlCmd_Taiwan.append("        round(sum(field_220300)/?,0) field_220300,     ");
                sqlCmd_Taiwan.append("        round(sum(field_220400)/?,0) field_220400,     ");
                sqlCmd_Taiwan.append("        round(sum(field_sum2)/?,0) field_sum2,         ");
                sqlCmd_Taiwan.append("        round(sum(field_220600)/?,0) field_220600,     ");
                sqlCmd_Taiwan.append("        round(sum(field_220700)/?,0) field_220700,     ");
                sqlCmd_Taiwan.append("        round(sum(field_220900)/?,0) field_220900,     ");
                sqlCmd_Taiwan.append("        round(sum(rpt2field_sumtotal)/?,0) rpt2field_sumtotal, "); 
                sqlCmd_Taiwan.append("        round(sum(field_120200)/?,0) field_120200,     ");
                sqlCmd_Taiwan.append("        round(sum(field_120101)/?,0) field_120101,     ");
                sqlCmd_Taiwan.append("        round(sum(field_120401)/?,0) field_120401,     ");
                sqlCmd_Taiwan.append("        round(sum(field_120501)/?,0) field_120501,     ");
                sqlCmd_Taiwan.append("        round(sum(field_120601)/?,0) field_120601,     ");
                sqlCmd_Taiwan.append("        round(sum(field_120700)/?,0) field_120700,     ");
                sqlCmd_Taiwan.append("        round(sum(field_150200)/?,0) field_150200      ");
                for(int k=1;k<=17;k++){
                    paramList_Taiwan.add(unit);
                }
                sqlCmd_Taiwan.append("   from (    ");
                sqlCmd_Taiwan.append("         select hsien_id ,  ");
                sqlCmd_Taiwan.append("                hsien_name, "); 
                sqlCmd_Taiwan.append("                FR001W_output_order,   ");
                sqlCmd_Taiwan.append("                count(*) as count_seq, ");
                sqlCmd_Taiwan.append("                bank_no ,  BANK_NAME,  ");
                sqlCmd_Taiwan.append("                sum(field_sumtotal) as field_sumtotal, ");
                sqlCmd_Taiwan.append("                sum(field_sum1)     as field_sum1,     ");
                sqlCmd_Taiwan.append("                sum(field_220100)   as field_220100,   ");
                sqlCmd_Taiwan.append("                sum(field_220300)     as field_220300, ");
                sqlCmd_Taiwan.append("                sum(field_220400)     as field_220400, ");
                sqlCmd_Taiwan.append("                sum(field_sum2)  as field_sum2,        ");
                sqlCmd_Taiwan.append("                sum(field_220600)     as field_220600, ");
                sqlCmd_Taiwan.append("                sum(field_220700)     as field_220700, ");
                sqlCmd_Taiwan.append("                sum(field_220900)     as field_220900, ");
                sqlCmd_Taiwan.append("                sum(rpt2field_sumtotal)     as rpt2field_sumtotal, ");
                sqlCmd_Taiwan.append("                sum(field_120200)     as field_120200, ");
                sqlCmd_Taiwan.append("                sum(field_120101)     as field_120101, ");
                sqlCmd_Taiwan.append("                sum(field_120401)     as field_120401, ");
                sqlCmd_Taiwan.append("                sum(field_120501)     as field_120501, ");
                sqlCmd_Taiwan.append("                sum(field_120601)     as field_120601, ");
                sqlCmd_Taiwan.append("                sum(field_120700)     as field_120700, ");
                sqlCmd_Taiwan.append("                sum(field_150200)     as field_150200  ");   
                sqlCmd_Taiwan.append("           from (  ");
                sqlCmd_Taiwan.append("                 select nvl(cd01.hsien_id,' ')       as  hsien_id ,  "); 
                sqlCmd_Taiwan.append("                        nvl(cd01.hsien_name,'OTHER') as  hsien_name, "); 
                sqlCmd_Taiwan.append("                        cd01.FR001W_output_order     as  FR001W_output_order, ");
                sqlCmd_Taiwan.append("                        COUNT(*)  AS  COUNT_SEQ, ");
                sqlCmd_Taiwan.append("                        bn01.bank_no ,  bn01.BANK_NAME, ");
                sqlCmd_Taiwan.append("                        sum(decode(a01.acc_code,'220100',amt,'221000',amt,'220300',amt,'220400',amt,'220500',amt,'220600',amt,'220700',amt,'220900',amt,0))     as field_sumtotal, ");
                sqlCmd_Taiwan.append("                        sum(decode(a01.acc_code,'220100',amt,'221000',amt,'220300',amt,'220400',amt,'220500',amt,0))     as field_sum1, "); 
                sqlCmd_Taiwan.append("                        sum(decode(a01.acc_code,'220100',amt,'221000',amt,0))     as field_220100, ");
                sqlCmd_Taiwan.append("                        sum(decode(a01.acc_code,'220300',amt,0))     as field_220300, ");
                sqlCmd_Taiwan.append("                        sum(decode(a01.acc_code,'220400',amt,'220500',amt,0))     as field_220400, ");
                sqlCmd_Taiwan.append("                        sum(decode(a01.acc_code,'220600',amt,'220700',amt,'220900',amt,0))  as field_sum2, ");
                sqlCmd_Taiwan.append("                        sum(decode(a01.acc_code,'220600',amt,0))     as field_220600, ");
                sqlCmd_Taiwan.append("                        sum(decode(a01.acc_code,'220700',amt,0))     as field_220700, ");
                sqlCmd_Taiwan.append("                        sum(decode(a01.acc_code,'220900',amt,0))     as field_220900, ");
                sqlCmd_Taiwan.append("                        sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0))     as rpt2field_sumtotal, ");
                sqlCmd_Taiwan.append("                        decode(YEAR_TYPE,'102',decode(bank_type,'6',sum(decode(a01.acc_code,'120200',amt,0)),'7',sum(decode(a01.acc_code,'120300',amt,0)),0),'103',sum(decode(a01.acc_code,'120200',amt,0)),0) as  field_120200, "); 
                sqlCmd_Taiwan.append("                        decode(YEAR_TYPE,'102',decode(bank_type,'6',sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120301',amt,'120302',amt,0)),'7',sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120401',amt,'120402',amt,0)),0),'103',sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120301',amt,'120302',amt,0)),0 ) as  field_120101, ");   
                sqlCmd_Taiwan.append("                        decode(YEAR_TYPE,'102',decode(bank_type,'6',sum(decode(a01.acc_code,'120401',amt,'120402',amt,0)),'7',sum(decode(a01.acc_code,'120201',amt,'120202',amt,0)),0),'103',sum(decode(a01.acc_code,'120401',amt,'120402',amt,0)),0 ) as  field_120401, "); 
                sqlCmd_Taiwan.append("                        sum(decode(a01.acc_code,'120501',amt,'120502',amt,0))     as field_120501, ");
                sqlCmd_Taiwan.append("                        sum(decode(a01.acc_code,'120601',amt,'120602',amt,'120603',amt,'120604',amt,0))     as field_120601, ");
                sqlCmd_Taiwan.append("                        sum(decode(a01.acc_code,'120700',amt,0))     as field_120700, ");
                sqlCmd_Taiwan.append("                        sum(decode(a01.acc_code,'150200',amt,0))     as field_150200  ");
                sqlCmd_Taiwan.append("                  from  (select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y' and cd01.Hsien_div = '2') cd01 "); 
                sqlCmd_Taiwan.append("                  left join (select bank_type,bn01.bank_no,bank_name,hsien_id "); 
                sqlCmd_Taiwan.append("                             from (select * from bn01 where m_year=? ");
                paramList_Taiwan.add(wlx01_m_year);
                if(bank_type.equals("ALL")){
                    sqlCmd_Taiwan.append("                                     and bank_type in('6','7') )bn01    ");
                }else{
                    sqlCmd_Taiwan.append("                                     and bank_type in(?) )bn01    ");
                    paramList_Taiwan.add(bank_type);
                }
                sqlCmd_Taiwan.append("                             left join (select * from wlx01 where m_year=?)wlx01 on bn01.bank_no = wlx01.bank_no ");
                sqlCmd_Taiwan.append("                            )bn01 on bn01.hsien_id=cd01.hsien_id  ");
                sqlCmd_Taiwan.append("                  left join (select (CASE WHEN (a01.m_year <= 102) THEN '102' ");
                sqlCmd_Taiwan.append("                                          WHEN (a01.m_year > 102) THEN '103'  ");
                sqlCmd_Taiwan.append("                                     ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01 "); 
                sqlCmd_Taiwan.append("                             where m_year=? and m_month=?  "); 
                paramList_Taiwan.add(wlx01_m_year);
                paramList_Taiwan.add(s_year);
                paramList_Taiwan.add(s_month);
                sqlCmd_Taiwan.append("                             and a01.acc_code in ('220100','221000','220300','220400','220500','220600','220700','220900','120000','120800','150300','120200','120300','120101','120102','120301','120302','120401','120402','120201','120202','120501','120502','120601','120602','120603','120604','120700','150200') ");               
                sqlCmd_Taiwan.append("                            )a01  on  bn01.bank_no = a01.bank_code "); 
                sqlCmd_Taiwan.append("                  group by YEAR_TYPE,bank_type,nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no ,  bn01.BANK_NAME ");
                sqlCmd_Taiwan.append("                ) a01 "); 
                sqlCmd_Taiwan.append("          where a01.bank_no <> ' ' "); 
                sqlCmd_Taiwan.append("          group by  hsien_id ,   hsien_name,   FR001W_output_order,bank_no,bank_name ");
                sqlCmd_Taiwan.append("        )  ");


              
              
              //建表開始--------------------------------------
              //明細表
                  dbData_detail= DBManager.QueryDB_SQLParam(sqlCmd_rptStyle1.toString(),paramList_rptStyle1,"field_sumtotal,field_sum1,field_220100,field_220300,field_220400,field_sum2,"
                          									 	  + "field_220600,field_220700,field_220900,rpt2field_sumtotal,field_120200,field_120101,"
						  									      + "field_120401,field_120501,field_120601,field_120700,field_150200");
                  System.out.println("明細表資料 共"+dbData_detail.size()+"筆 ");
                  dbData_Taiwan = DBManager.QueryDB_SQLParam(sqlCmd_Taiwan.toString(),paramList_Taiwan,"field_sumtotal,field_sum1,field_220100,field_220300,field_220400,field_sum2,"
                          										+ "field_220600,field_220700,field_220900,rpt2field_sumtotal,field_120200,field_120101,"
                          										+ "field_120401,field_120501,field_120601,field_120700,field_150200");
                  System.out.println("台灣省資料 共"+dbData_Taiwan.size()+"筆 ");
                  HSSFFont ft = wb.createFont();
                  HSSFCellStyle cs = wb.createCellStyle();
                  ft.setFontHeightInPoints((short)16);
                  ft.setFontName("標楷體");
                  cs.setFont(ft);
                  cs.setAlignment(HSSFCellStyle.ALIGN_CENTER);

                  //共同資料的格式
                  HSSFFont ft2 = wb.createFont();
                  HSSFCellStyle cs3 = wb.createCellStyle();
                  ft2.setFontHeightInPoints((short)10);
                  cs3.setFont(ft2);
                  cs3.setBorderTop(HSSFCellStyle.BORDER_THIN);
                  cs3.setBorderBottom(HSSFCellStyle.BORDER_THIN);
                  cs3.setBorderLeft(HSSFCellStyle.BORDER_THIN);
                  cs3.setBorderRight(HSSFCellStyle.BORDER_THIN);
                  //結尾說明使用
                  HSSFFont ft4 = wb.createFont();
                  HSSFCellStyle cs4 = wb.createCellStyle();
                  ft4.setFontHeightInPoints((short)12);
                  ft4.setFontName("標楷體");
                  cs4.setFont(ft4);
                  cs4.setAlignment(HSSFCellStyle.ALIGN_LEFT);
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
                  DataObject bean = null;
                  String insertValue = "";
                  HSSFFooter footer = sheet.getFooter();
                  //sheetidx=0存款餘額表,sheetidx=1放款總額表
                  for(int sheetidx=0;sheetidx<2;sheetidx++){
                      if(sheetidx==1){//放款總額表
                         sheet = wb.getSheetAt(1);//讀取第二個工作表，宣告其為sheet
                         if(sheet==null){System.out.println("open sheet 失敗");}else System.out.println("open sheet 成功");
                         ps = sheet.getPrintSetup(); //取得設定
                         //設定頁面符合列印大小
                         sheet.setAutobreaks( false );
                         ps.setScale( ( short )76 ); //列印縮放百分比
                         ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
                         finput.close();
                      }//end of 放款總額表
                      row = sheet.createRow((short)1);
                      if(sheetidx==0){//存款餘額表
                          for(int v=0;v<11;v++){
                              row.createCell((short)v);
                          }
                      }else if(sheetidx==1){//放款總額表
                          for(int v=0;v<10;v++){
                              row.createCell((short)v);
                          }
                      }
                      cell = row.getCell( (short) 4);
                      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                      cell.setCellValue(bank_type_name +"信用部"+(sheetidx==0?"存":"放")+"款餘額明細表");
                      cell.setCellStyle(cs);

                      row = sheet.createRow((short)2);
                      if(sheetidx==0){//存款餘額表
                          for(int v=0;v<11;v++){
                              row.createCell((short)v);
                          }
                      }else if(sheetidx==1){//放款總額表
                          for(int v=0;v<10;v++){
                              row.createCell((short)v);
                          }
                      }
                      cell = row.getCell( (short) 4);
                      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                      cell.setCellValue(""+s_year+"年"+s_month+"月");
                      cell.setCellStyle(cs);
                      row = sheet.getRow(2);
                      if(sheetidx==0){//存款餘額表
                         cell = row.getCell( (short)9);
                      }else{//放款總額表
                         cell = row.getCell( (short)8);
                      }
                      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                      cell.setCellValue("單位:新臺幣 " + unit_name + ",%");

                      rowNum=4;

                      //存/放款餘額表========================================================================
                      for(int i=1;i<dbData_detail.size();i++){
                          bean = (DataObject)dbData_detail.get(i);
                          hsien_name = String.valueOf(bean.getValue("hsien_name"));
                          bank_no = String.valueOf(bean.getValue("bank_no"));
                          bank_name = String.valueOf(bean.getValue("bank_name"));
                          count_seq = String.valueOf(bean.getValue("count_seq"));
                          field_seq = String.valueOf(bean.getValue("field_seq"));

                          row = sheet.createRow(rowNum+i);
                          for(int cellcount=0;cellcount< sheetCell[sheetidx];cellcount++){
                              cell = row.createCell( (short)cellcount);
                              cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                              cs3.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
                              insertValue = "";
                              if( cellcount==0 && field_seq.equals("A01") )insertValue =bank_no;
                              else if( cellcount==1 && field_seq.equals("A01") )insertValue =bank_name;
                              else if( cellcount==1 && ( field_seq.equals("A90") ||field_seq.equals("A99") ) ) insertValue =hsien_name;
                              else if( cellcount >=2){
                                  if(sheetidx == 0){//存款餘額表
                                      insertValue = getBeanData(bean,dataIdx0,cellcount-2);//存款餘額~公庫存款
                                  }else{//放款總額表
                                      insertValue = getBeanData(bean,dataIdx1,cellcount-2);//放款總額~催收款
                                  }
                              }
                              cell.setCellValue(insertValue);
                              if(cellcount == 0 || cellcount ==1 ){
                                  cs3.setAlignment(HSSFCellStyle.ALIGN_CENTER);
                              }else{
                                  cs3.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
                              }

                              if(cellcount==1)cell.setCellStyle(cl);
                              else cell.setCellStyle(cs3);
                          }//each of detail
                          if( field_seq.equals("A90") || field_seq.equals("A99") ){//A90->小計.A99總計
                          	  if(Integer.parseInt(s_year) <= 99){
                          	     if(hsien_name.equals("台北市") || hsien_name.equals("高雄市")){//將台北市.高雄市先儲存起來
                          	     	if((bank_type.equals("7") && dbData_detail_hsien.size() < 1)
                          	     	||(!bank_type.equals("7") && dbData_detail_hsien.size() < 2)){	
                          	     	   dbData_detail_hsien.add((DataObject)dbData_detail.get(i));
                          	     	}	
                          	     }	
                          	  }else{
                          	     if(hsien_name.equals("新北市") || hsien_name.equals("臺北市") || hsien_name.equals("桃園市") || hsien_name.equals("臺中市") || hsien_name.equals("臺南市") || hsien_name.equals("高雄市")){//將新北市.台北市.桃園市.台中市.台南市.高雄市先儲存起來
                          	     	if((bank_type.equals("7") && dbData_detail_hsien.size() < 4)
                                    ||(!bank_type.equals("7") && dbData_detail_hsien.size() < 6)){		
                          	     	   dbData_detail_hsien.add((DataObject)dbData_detail.get(i));
                          	     	}
                      	         }	
                          	  }                          	  
                              rowNum++;
                              row = sheet.createRow(rowNum+i);
                              for(int cellcount=0;cellcount<sheetCell[sheetidx];cellcount++){
                                  cell = row.createCell( (short)cellcount);cell.setCellValue(""); cell.setCellStyle(cs3);
                              }
                          }//小計和總計後要空一行
                      }//end of 存/放款餘額表
                      //總計==============================================================================
                      bean = (DataObject)dbData_detail.get(0);
                      if(sheetidx == 0){//存款餘額表.總計
                          System.out.println("put 總計 ="+rowNum);
                         putCell(bean,sheet,row,cell,cs3,dataIdx0,dbData_detail.size()+rowNum,"false",sheetCell[sheetidx]);//放總計
                      }else{//放款總額表.總計
                          System.out.println("put 總計 ="+rowNum);
                         putCell(bean,sheet,row,cell,cs3,dataIdx1,dbData_detail.size()+rowNum,"false",sheetCell[sheetidx]);//放總計
                      }
                      //=================================================================================
                      //畫最下面台北市.高雄市.台灣省的小計與總計的部分
                      rowNum = 3+rowNum+dbData_detail.size();
                      System.out.println("dbData_detail_hsien.size()="+dbData_detail_hsien.size());
                      if(sheetidx == 0){//存款餘額表
                          //System.out.println("sheetidx="+sheetidx);
                          //System.out.println("put 台北="+rowNum);
                      	  for(int i=0;i<dbData_detail_hsien.size();i++){
                      	  	  bean = (DataObject)dbData_detail_hsien.get(i);
                      	      rowNum = putCell(bean,sheet,row,cell,cs3,dataIdx0,rowNum,"true",sheetCell[sheetidx]);
                      	  }
                      }else{//放款總額表
                          //System.out.println("sheetidx="+sheetidx);
                          //System.out.println("put 台北="+rowNum);
                      	  for(int i=0;i<dbData_detail_hsien.size();i++){
                    	  	  bean = (DataObject)dbData_detail_hsien.get(i);
                    	  	  rowNum = putCell(bean,sheet,row,cell,cs3,dataIdx1,rowNum,"true",sheetCell[sheetidx]);
                    	  }                      	  
                      }
                      //台灣省小計======================================================================
                      bean = (DataObject)dbData_Taiwan.get(0);
                      field_seq = String.valueOf(bean.getValue("field_seq"));
                      hsien_name = String.valueOf(bean.getValue("hsien_name"));
                      if( field_seq.equals("A92") && hsien_name.equals("臺灣省")  ){
                          //System.out.println("sheetidx="+sheetidx);
                          if(sheetidx == 0){//存款餘額表
                             rowNum = putCell(bean,sheet,row,cell,cs3,dataIdx0,rowNum,"true",sheetCell[sheetidx]);
                             //System.out.println("put 台灣省 ="+rowNum);
                          }else{//放款餘額表
                             rowNum = putCell(bean,sheet,row,cell,cs3,dataIdx1,rowNum,"true",sheetCell[sheetidx]);
                             //System.out.println("put 台灣省 ="+rowNum);
                          }
                      }
                      //===============================================================================
                      //總計再取出來一次，放在最底下=======================================================
                      bean = (DataObject)dbData_detail.get(0);
                      if(sheetidx == 0){//存款餘額表
                          //System.out.println("put 總計 ="+rowNum);
                         rowNum = putCell(bean,sheet,row,cell,cs3,dataIdx0,rowNum,"true",sheetCell[sheetidx]);
                      }else{//放款總額表
                          //System.out.println("put 總計 ="+rowNum);
                         rowNum = putCell(bean,sheet,row,cell,cs3,dataIdx1,rowNum,"true",sheetCell[sheetidx]);
                      }
                      //===============================================================================
                      rowNum =rowNum+2;
                      row = sheet.createRow(rowNum);
                      cell = row.createCell( (short)0);
                      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                      cell.setCellValue("資料來源:依各農會信用部由自有電腦設備或委由相關資訊中心以網際網路傳送之資料彙編");
                      cell.setCellStyle(cs4);
                      rowNum++;
                      if(bank_type.equals("6") || bank_type.equals("ALL")){//103.01.27 add
                          row = sheet.createRow(rowNum);
                          cell = row.createCell( (short)0);
                          cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                          cell.setCellValue("原臺灣省農會，該農會於102年5月22日更名為中華民國農會");
                          cell.setCellStyle(cs4);
                          rowNum++;
                      }
                      footer = sheet.getFooter();
                      footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
                      footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
                  }//end of sheetidx
              //建表結束--------------------------------------
              FileOutputStream fout=new FileOutputStream(reportDir + System.getProperty("file.separator")+filename);
              wb.write(fout);//儲存
              fout.close();
              System.out.println("儲存完成");
      }catch(Exception e){
                System.out.println("createRpt Error:"+e+e.getMessage());
      }
      return errMsg;
    }
    /*
     * 取得bean裡的data
     * String[] dataIdx:column name
     * int idx :column index
     */
   private static String getBeanData(DataObject bean,String[] dataIdx,int idx){
           String insertValue = "";
           try{
               insertValue = String.valueOf(bean.getValue(dataIdx[idx]));
               if(insertValue != null && !insertValue.equals("") && !insertValue.equals("0")){
                   insertValue = Utility.setCommaFormat(insertValue);
               }
           }catch(Exception e){
               System.out.println("getBeanData Error["+dataIdx[idx]+"]:"+e.getMessage());
           }
           return insertValue;
   }

   /* 放台北市.高雄市.台灣省小計.總計
    * String[] dataIdx:column name
    * int rowNum: row number
    * String hasAddRowNum:是否要累加row rumber
    */
   private static int putCell(DataObject bean,HSSFSheet sheet,
                              HSSFRow row,HSSFCell cell,HSSFCellStyle cs3,
                              String[] dataIdx,int rowNum,String hasAddRowNum,int sheetlength)
   {
       int returnRowNum = 0;
       try{
           String hsien_name = String.valueOf(bean.getValue("hsien_name"));
           String bank_no = String.valueOf(bean.getValue("bank_no"));
           String bank_name = String.valueOf(bean.getValue("bank_name"));
           String count_seq = String.valueOf(bean.getValue("count_seq"));
           String field_seq = String.valueOf(bean.getValue("field_seq"));
           String insertValue = "";
           if(hasAddRowNum.equals("false")){
              row = sheet.createRow(rowNum);
           }else{
              row = sheet.createRow(rowNum++);
           }
           for(int cellcount=0;cellcount< sheetlength;cellcount++){
               cell = row.createCell( (short)cellcount);
               cell.setEncoding(HSSFCell.ENCODING_UTF_16);
               cs3.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
               insertValue = "";
               if( cellcount==1){
                   insertValue = hsien_name;
               }else if( cellcount >=2){
                   insertValue = getBeanData(bean,dataIdx,cellcount-2);//存款餘額~公庫存款
               }
               cell.setCellValue(insertValue);
               cell.setCellStyle(cs3);
           }//放總計
           returnRowNum = rowNum;
       }catch(Exception e){
           System.out.println();
       }
       return returnRowNum;
   }
}
