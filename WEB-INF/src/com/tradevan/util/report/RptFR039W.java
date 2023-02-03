/*
 * Created on 2006/07/26-28 by 2295
 * 99.09.13-14 fix 根據查詢年度.100年以後取得新縣市別.100年以前取得舊縣市別 
  			        使用PreparedStatement;並列印轉換後的SQL;套用QueryDB_SQLParam by 2295
 * 102.06.21 1.add 103年後SQL  2.總表與明細表程式分開 by 2968 
 * 103.01.27 add 臺灣省改其他,並增加說明 by 2295
 * 109.07.27 fix 漁會/農漁會報表無法下載,因104年增加金門縣金門區漁會信用部 by 2295        
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

public class RptFR039W {
    public static String createRpt(String s_year,String s_month,String unit,String bank_type){
        String errMsg = "";
        StringBuffer sqlCmd_rptStyle0 = new StringBuffer();//總表組合sql
        StringBuffer sqlCmd_sum = new StringBuffer();//全体農漁會
        List dbData_All = null;//總表
        List dbData_sum = null;//明細表.總計
        List paramList_rptStyle0 = new ArrayList();
        List paramList_sum = new ArrayList();
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
        System.out.println("RptFR039W.bank_type="+bank_type);
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
                openfile = bank_type_name+"信用部存放款餘額表_總表"+(Integer.parseInt(s_year)<=99?"_99":"")+".xls";
                filename = bank_type_name+"信用部存放款餘額表_總表.xls";

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
                ps.setScale( ( short )80 ); //列印縮放百分比
                

                ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
                //wb.setSheetName(0,"test");
                finput.close();

                HSSFRow row=null;//宣告一列
                HSSFCell cell=null;//宣告一個儲存格


                //總表組合sql
                sqlCmd_rptStyle0.append("  select hsien_id , hsien_name,  FR001W_output_order, ");
                sqlCmd_rptStyle0.append("         round(field_sumtotal/?,0) field_sumtotal, ");// --存款餘額
                sqlCmd_rptStyle0.append("         round(field_sum1/?,0) field_sum1,  ");//--小計 
                sqlCmd_rptStyle0.append("         round(field_220100/?,0) field_220100, ");//--支票存款(含本會支票)
                sqlCmd_rptStyle0.append("         round(field_220300/?,0) field_220300, ");//--活期存款
                sqlCmd_rptStyle0.append("         round(field_220400/?,0) field_220400, ");//--活期儲蓄存款(含員工活儲)
                sqlCmd_rptStyle0.append("         round(field_sum2/?,0) field_sum2, ");//--小計
                sqlCmd_rptStyle0.append("         round(field_220600/?,0) field_220600, ");//定期存款
                sqlCmd_rptStyle0.append("         round(field_220700/?,0) field_220700, ");//定期儲蓄存款
                sqlCmd_rptStyle0.append("         round(field_220900/?,0) field_220900, ");//公庫存款
                sqlCmd_rptStyle0.append("         round(rpt2field_sumtotal/?,0) rpt2field_sumtotal, ");//放款總額
                sqlCmd_rptStyle0.append("         round(field_120200/?,0) field_120200, ");//貼現
                sqlCmd_rptStyle0.append("         round(field_120101/?,0) field_120101, ");//一般放款及透支
                sqlCmd_rptStyle0.append("         round(field_120401/?,0) field_120401, ");//統一漁貸
                sqlCmd_rptStyle0.append("         round(field_120501/?,0) field_120501, ");//專案放款
                sqlCmd_rptStyle0.append("         round(field_120601/?,0) field_120601, ");//農業發展基金放款
                sqlCmd_rptStyle0.append("         round(field_120700/?,0) field_120700, ");//內部融資
                sqlCmd_rptStyle0.append("         round(field_150200/?,0) field_150200  ");//催收款
                for(int k=1;k<=17;k++){
                    paramList_rptStyle0.add(unit);
                }
                sqlCmd_rptStyle0.append(" from   ");
                sqlCmd_rptStyle0.append("     (  ");
                sqlCmd_rptStyle0.append("       select hsien_id , "); 
                sqlCmd_rptStyle0.append("              hsien_name,  ");
                sqlCmd_rptStyle0.append("              FR001W_output_order, ");
                sqlCmd_rptStyle0.append("              sum(field_sumtotal) as field_sumtotal, ");
                sqlCmd_rptStyle0.append("              sum(field_sum1)     as field_sum1, "); 
                sqlCmd_rptStyle0.append("              sum(field_220100)   as field_220100, ");
                sqlCmd_rptStyle0.append("              sum(field_220300)     as field_220300, ");
                sqlCmd_rptStyle0.append("              sum(field_220400)     as field_220400, ");
                sqlCmd_rptStyle0.append("              sum(field_sum2)  as field_sum2, ");
                sqlCmd_rptStyle0.append("              sum(field_220600)     as field_220600, ");
                sqlCmd_rptStyle0.append("              sum(field_220700)     as field_220700, ");
                sqlCmd_rptStyle0.append("              sum(field_220900)     as field_220900, ");
                sqlCmd_rptStyle0.append("              sum(rpt2field_sumtotal)     as rpt2field_sumtotal, ");
                sqlCmd_rptStyle0.append("              sum(field_120200)     as field_120200, ");
                sqlCmd_rptStyle0.append("              sum(field_120101)     as field_120101, ");
                sqlCmd_rptStyle0.append("              sum(field_120401)     as field_120401, ");
                sqlCmd_rptStyle0.append("              sum(field_120501)     as field_120501, ");
                sqlCmd_rptStyle0.append("              sum(field_120601)     as field_120601, ");
                sqlCmd_rptStyle0.append("              sum(field_120700)     as field_120700, ");
                sqlCmd_rptStyle0.append("              sum(field_150200)     as field_150200  ");   
                sqlCmd_rptStyle0.append("       from (  ");
                sqlCmd_rptStyle0.append("             select nvl(cd01.hsien_id,' ') as  hsien_id , "); 
                sqlCmd_rptStyle0.append("                    nvl(cd01.hsien_name,'OTHER') as  hsien_name, "); 
                sqlCmd_rptStyle0.append("                    cd01.FR001W_output_order     as  FR001W_output_order, ");
                sqlCmd_rptStyle0.append("                    bn01.bank_no ,  bn01.BANK_NAME, ");
                sqlCmd_rptStyle0.append("                    sum(decode(a01.acc_code,'220100',amt,'221000',amt,'220300',amt,'220400',amt,'220500',amt,'220600',amt,'220700',amt,'220900',amt,0))     as field_sumtotal, ");
                sqlCmd_rptStyle0.append("                    sum(decode(a01.acc_code,'220100',amt,'221000',amt,'220300',amt,'220400',amt,'220500',amt,0))     as field_sum1, "); 
                sqlCmd_rptStyle0.append("                    sum(decode(a01.acc_code,'220100',amt,'221000',amt,0))     as field_220100, ");
                sqlCmd_rptStyle0.append("                    sum(decode(a01.acc_code,'220300',amt,0))     as field_220300, ");
                sqlCmd_rptStyle0.append("                    sum(decode(a01.acc_code,'220400',amt,'220500',amt,0))     as field_220400, ");
                sqlCmd_rptStyle0.append("                    sum(decode(a01.acc_code,'220600',amt,'220700',amt,'220900',amt,0))  as field_sum2, ");
                sqlCmd_rptStyle0.append("                    sum(decode(a01.acc_code,'220600',amt,0))     as field_220600, ");
                sqlCmd_rptStyle0.append("                    sum(decode(a01.acc_code,'220700',amt,0))     as field_220700, ");
                sqlCmd_rptStyle0.append("                    sum(decode(a01.acc_code,'220900',amt,0))     as field_220900, ");
                sqlCmd_rptStyle0.append("                    sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0))     as rpt2field_sumtotal, ");
                sqlCmd_rptStyle0.append("                    decode(YEAR_TYPE,'102',decode(bank_type,'6',sum(decode(a01.acc_code,'120200',amt,0)),'7',sum(decode(a01.acc_code,'120300',amt,0)),0),'103',sum(decode(a01.acc_code,'120200',amt,0)),0) as  field_120200, "); 
                sqlCmd_rptStyle0.append("                    decode(YEAR_TYPE,'102',decode(bank_type,'6',sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120301',amt,'120302',amt,0)),'7',sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120401',amt,'120402',amt,0)),0),'103',sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120301',amt,'120302',amt,0)),0 ) as  field_120101, ");   
                sqlCmd_rptStyle0.append("                    decode(YEAR_TYPE,'102',decode(bank_type,'6',sum(decode(a01.acc_code,'120401',amt,'120402',amt,0)),'7',sum(decode(a01.acc_code,'120201',amt,'120202',amt,0)),0),'103',sum(decode(a01.acc_code,'120401',amt,'120402',amt,0)),0 ) as  field_120401, "); 
                sqlCmd_rptStyle0.append("                    sum(decode(a01.acc_code,'120501',amt,'120502',amt,0))     as field_120501, ");
                sqlCmd_rptStyle0.append("                    sum(decode(a01.acc_code,'120601',amt,'120602',amt,'120603',amt,'120604',amt,0))     as field_120601, ");
                sqlCmd_rptStyle0.append("                    sum(decode(a01.acc_code,'120700',amt,0))     as field_120700, ");
                sqlCmd_rptStyle0.append("                    sum(decode(a01.acc_code,'150200',amt,0))     as field_150200  ");
                sqlCmd_rptStyle0.append("             from( select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y') cd01 "); 
                sqlCmd_rptStyle0.append("             left join ( select bank_type,bn01.bank_no,bank_name,hsien_id  ");
                sqlCmd_rptStyle0.append("                           from (select * from bn01 where m_year=? ");
                paramList_rptStyle0.add(wlx01_m_year);
                if(bank_type.equals("ALL")){
                    sqlCmd_rptStyle0.append("                            and bank_type in('6','7') )bn01 ");    
                }else{
                    sqlCmd_rptStyle0.append("                            and bank_type in(?) )bn01 ");
                    paramList_rptStyle0.add(bank_type);
                }
                sqlCmd_rptStyle0.append("                           left join (select * from wlx01 where m_year=? )wlx01 on bn01.bank_no = wlx01.bank_no ");
                sqlCmd_rptStyle0.append("                         )bn01 on bn01.hsien_id=cd01.hsien_id   ");
                sqlCmd_rptStyle0.append("             left join (select (CASE WHEN (a01.m_year <= 102) THEN '102' ");
                sqlCmd_rptStyle0.append("                                      WHEN (a01.m_year > 102) THEN '103' ");
                sqlCmd_rptStyle0.append("                                 ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01 "); 
                sqlCmd_rptStyle0.append("                               where m_year=? and m_month=? ");
                paramList_rptStyle0.add(wlx01_m_year);
                paramList_rptStyle0.add(s_year);
                paramList_rptStyle0.add(s_month);
                sqlCmd_rptStyle0.append("                               and a01.acc_code in ('220100','221000','220300','220400','220500','220600','220700','220900','120000','120800','150300','120200','120300','120101','120102','120301','120302','120401','120402','120201','120202','120501','120502','120601','120602','120603','120604','120700','150200') ");               
                sqlCmd_rptStyle0.append("                               )a01  on  bn01.bank_no = a01.bank_code  ");
                sqlCmd_rptStyle0.append("             group by YEAR_TYPE,bank_type,nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no ,  bn01.BANK_NAME ");
                sqlCmd_rptStyle0.append("            )a01 "); 
                sqlCmd_rptStyle0.append(" where a01.bank_no <> ' ' ");
                sqlCmd_rptStyle0.append(" GROUP  BY a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order ");
                sqlCmd_rptStyle0.append(" order by a01.FR001W_output_order ");
                sqlCmd_rptStyle0.append(" )  ");

                //存放款餘額表-全体農/漁會
                sqlCmd_sum.append("  select  ");
                sqlCmd_sum.append("         round(field_sumtotal/?,0) field_sumtotal, "); 
                sqlCmd_sum.append("         round(field_sum1/?,0) field_sum1, "); 
                sqlCmd_sum.append("         round(field_220100/?,0) field_220100, "); 
                sqlCmd_sum.append("         round(field_220300/?,0) field_220300, "); 
                sqlCmd_sum.append("         round(field_220400/?,0) field_220400, "); 
                sqlCmd_sum.append("         round(field_sum2/?,0) field_sum2, "); 
                sqlCmd_sum.append("         round(field_220600/?,0) field_220600, "); 
                sqlCmd_sum.append("         round(field_220700/?,0) field_220700, "); 
                sqlCmd_sum.append("         round(field_220900/?,0) field_220900, "); 
                sqlCmd_sum.append("         round(rpt2field_sumtotal/?,0) rpt2field_sumtotal, "); 
                sqlCmd_sum.append("         round(field_120200/?,0) field_120200, "); 
                sqlCmd_sum.append("         round(field_120101/?,0) field_120101, "); 
                sqlCmd_sum.append("         round(field_120401/?,0) field_120401, "); 
                sqlCmd_sum.append("         round(field_120501/?,0) field_120501, "); 
                sqlCmd_sum.append("         round(field_120601/?,0) field_120601, "); 
                sqlCmd_sum.append("         round(field_120700/?,0) field_120700, "); 
                sqlCmd_sum.append("         round(field_150200/?,0) field_150200  "); 
                for(int k=1;k<=17;k++){
                    paramList_sum.add(unit);
                }
                sqlCmd_sum.append(" from   ");
                sqlCmd_sum.append("     (  ");
                sqlCmd_sum.append("       select  ");
                sqlCmd_sum.append("              sum(field_sumtotal) as field_sumtotal,  ");
                sqlCmd_sum.append("              sum(field_sum1)     as field_sum1,      ");
                sqlCmd_sum.append("              sum(field_220100)   as field_220100,    ");
                sqlCmd_sum.append("              sum(field_220300)     as field_220300,  ");
                sqlCmd_sum.append("              sum(field_220400)     as field_220400,  ");
                sqlCmd_sum.append("              sum(field_sum2)  as field_sum2,         ");
                sqlCmd_sum.append("              sum(field_220600)     as field_220600,  ");
                sqlCmd_sum.append("              sum(field_220700)     as field_220700,  ");
                sqlCmd_sum.append("              sum(field_220900)     as field_220900,  ");
                sqlCmd_sum.append("              sum(rpt2field_sumtotal)     as rpt2field_sumtotal,  ");
                sqlCmd_sum.append("              sum(field_120200)     as field_120200,  ");
                sqlCmd_sum.append("              sum(field_120101)     as field_120101,  ");
                sqlCmd_sum.append("              sum(field_120401)     as field_120401,  ");
                sqlCmd_sum.append("              sum(field_120501)     as field_120501,  ");
                sqlCmd_sum.append("              sum(field_120601)     as field_120601,  ");
                sqlCmd_sum.append("              sum(field_120700)     as field_120700,  ");
                sqlCmd_sum.append("              sum(field_150200)     as field_150200   ");  
                sqlCmd_sum.append("       from (  ");
                sqlCmd_sum.append("             select nvl(cd01.hsien_id,' ') as  hsien_id , "); 
                sqlCmd_sum.append("                    nvl(cd01.hsien_name,'OTHER') as  hsien_name, "); 
                sqlCmd_sum.append("                    cd01.FR001W_output_order     as  FR001W_output_order,  ");
                sqlCmd_sum.append("                    bn01.bank_no ,  bn01.BANK_NAME,  ");
                sqlCmd_sum.append("                    sum(decode(a01.acc_code,'220100',amt,'221000',amt,'220300',amt,'220400',amt,'220500',amt,'220600',amt,'220700',amt,'220900',amt,0))     as field_sumtotal,  ");
                sqlCmd_sum.append("                    sum(decode(a01.acc_code,'220100',amt,'221000',amt,'220300',amt,'220400',amt,'220500',amt,0))     as field_sum1,  ");
                sqlCmd_sum.append("                    sum(decode(a01.acc_code,'220100',amt,'221000',amt,0))     as field_220100, ");
                sqlCmd_sum.append("                    sum(decode(a01.acc_code,'220300',amt,0))     as field_220300, ");
                sqlCmd_sum.append("                    sum(decode(a01.acc_code,'220400',amt,'220500',amt,0))     as field_220400, ");
                sqlCmd_sum.append("                    sum(decode(a01.acc_code,'220600',amt,'220700',amt,'220900',amt,0))  as field_sum2, ");
                sqlCmd_sum.append("                    sum(decode(a01.acc_code,'220600',amt,0))     as field_220600, ");
                sqlCmd_sum.append("                    sum(decode(a01.acc_code,'220700',amt,0))     as field_220700, ");
                sqlCmd_sum.append("                    sum(decode(a01.acc_code,'220900',amt,0))     as field_220900, ");
                sqlCmd_sum.append("                    sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0))     as rpt2field_sumtotal, ");
                sqlCmd_sum.append("                    decode(YEAR_TYPE,'102',decode(bank_type,'6',sum(decode(a01.acc_code,'120200',amt,0)),'7',sum(decode(a01.acc_code,'120300',amt,0)),0),'103',sum(decode(a01.acc_code,'120200',amt,0)),0) as  field_120200, "); 
                sqlCmd_sum.append("                    decode(YEAR_TYPE,'102',decode(bank_type,'6',sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120301',amt,'120302',amt,0)),'7',sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120401',amt,'120402',amt,0)),0),'103',sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120301',amt,'120302',amt,0)),0 ) as  field_120101, ");
                sqlCmd_sum.append("                    decode(YEAR_TYPE,'102',decode(bank_type,'6',sum(decode(a01.acc_code,'120401',amt,'120402',amt,0)),'7',sum(decode(a01.acc_code,'120201',amt,'120202',amt,0)),0),'103',sum(decode(a01.acc_code,'120401',amt,'120402',amt,0)),0 ) as  field_120401, "); 
                sqlCmd_sum.append("                    sum(decode(a01.acc_code,'120501',amt,'120502',amt,0))     as field_120501, ");
                sqlCmd_sum.append("                    sum(decode(a01.acc_code,'120601',amt,'120602',amt,'120603',amt,'120604',amt,0))     as field_120601, ");
                sqlCmd_sum.append("                    sum(decode(a01.acc_code,'120700',amt,0))     as field_120700, ");
                sqlCmd_sum.append("                    sum(decode(a01.acc_code,'150200',amt,0))     as field_150200  ");
                sqlCmd_sum.append("             from( select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y') cd01    ");
                sqlCmd_sum.append("             left join ( select bank_type,bn01.bank_no,bank_name,hsien_id     ");
                sqlCmd_sum.append("                            from (select * from bn01 where m_year=? ");
                paramList_sum.add(wlx01_m_year);
                if(bank_type.equals("ALL")){
                    sqlCmd_sum.append("                            and bank_type in('6','7') )bn01 ");
                }else{
                    sqlCmd_sum.append("                            and bank_type in(?) )bn01 ");
                    paramList_sum.add(bank_type);
                }
                sqlCmd_sum.append("                           left join (select * from wlx01 where m_year=?)wlx01 on bn01.bank_no = wlx01.bank_no ");
                sqlCmd_sum.append("                         )bn01 on bn01.hsien_id=cd01.hsien_id   ");
                sqlCmd_sum.append("             left join (select (CASE WHEN (a01.m_year <= 102) THEN '102' ");
                sqlCmd_sum.append("                                      WHEN (a01.m_year > 102) THEN '103' ");
                sqlCmd_sum.append("                                 ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01 "); 
                sqlCmd_sum.append("                               where m_year=? and m_month=? ");        
                paramList_sum.add(wlx01_m_year);
                paramList_sum.add(s_year);
                paramList_sum.add(s_month);
                sqlCmd_sum.append("                               and a01.acc_code in ('220100','221000','220300','220400','220500','220600','220700','220900','120000','120800','150300','120200','120300','120101','120102','120301','120302','120401','120402','120201','120202','120501','120502','120601','120602','120603','120604','120700','150200') ");               
                sqlCmd_sum.append("                               )a01  on  bn01.bank_no = a01.bank_code "); 
                sqlCmd_sum.append("             group by YEAR_TYPE,bank_type,nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no ,  bn01.BANK_NAME ");
                sqlCmd_sum.append("            )a01 "); 
                sqlCmd_sum.append(" where a01.bank_no <> ' ' ");
                sqlCmd_sum.append(" )  ");

              
              
                //建表開始--------------------------------------
                //總表
              
                  System.out.println("總表sql="+sqlCmd_rptStyle0);
                  dbData_All = DBManager.QueryDB_SQLParam(sqlCmd_rptStyle0.toString(),paramList_rptStyle0,"field_sumtotal,field_sum1,field_220100,field_220300,field_220400,field_sum2,"
                   						  + "field_220600,field_220700,field_220900,rpt2field_sumtotal,field_120200,field_120101,"
                   						  + "field_120401,field_120501,field_120601,field_120700,field_150200");
                  System.out.print("總表資料 共"+dbData_All.size()+"筆 ");
                  dbData_sum = DBManager.QueryDB_SQLParam(sqlCmd_sum.toString(),paramList_sum,"field_sumtotal,field_sum1,field_220100,field_220300,field_220400,field_sum2,"
							  + "field_220600,field_220700,field_220900,rpt2field_sumtotal,field_120200,field_120101,"
							  + "field_120401,field_120501,field_120601,field_120700,field_150200");
                  System.out.println("存放款餘額表-全体農/漁會.size="+dbData_sum.size());
                  if(dbData_All.size() == 0){
                      row=sheet.getRow(1);
                      cell=row.getCell((short)3);
                      //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
                      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                      cell.setCellValue(s_year +"年" +s_month +"月無資料存在");
                  }else{
                      //存款餘額表表頭
                      row=sheet.getRow(1);
                      cell=row.getCell((short)4);
                      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                      cell.setCellValue("中華民國" +s_year +"年" +s_month+"月");
                      //列印單位
                      row=sheet.getRow(1);
                      cell=(row.getCell((short)9)==null)? row.createCell((short)9) : row.getCell((short)9);
                      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                      cell.setCellValue("單位:新臺幣 " + unit_name + ",%");
                      //=================================================================================
                      //放款總額表表頭======================================================================
                      if(bank_type.equals("6") || bank_type.equals("ALL")){//農會.農漁會
                      	if(Integer.parseInt(s_year) <= 99){
                      		row=sheet.getRow(38);//農會.農漁會
                      	}else{
                      		if(bank_type.equals("6")){
                      		   row=sheet.getRow(35);//農會
                      		}else if(bank_type.equals("ALL")){
                      		   row=sheet.getRow(36);//農漁會//109.07.27 add 金門縣
                      		}
                      	}
                      }else{//漁會
                      	 if(Integer.parseInt(s_year) <= 99){
                            row=sheet.getRow(26);
                      	 }else if(Integer.parseInt(s_year) >= 100){
                      	 	row=sheet.getRow(25);//農漁會//109.07.27 add 金門縣
                      	 }
                      }
                      cell=row.getCell((short)4);
                      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                      cell.setCellValue("中華民國" +s_year +"年" +s_month+"月");
                      cell=(row.getCell((short)7)==null)? row.createCell((short)7) : row.getCell((short)7);
                      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                      cell.setCellValue("單位:新臺幣 " + unit_name + ",%");
                      //==================================================================================
                      rowNum=4;
                      DataObject bean = null;
                      String insertValue = "";

                      for(int i=0;i<dbData_All.size()+1;i++){
                          if( i == dbData_All.size() && dbData_sum.size() != 0 && dbData_sum != null){
           	  	 		      bean = (DataObject)dbData_sum.get(0);//全体農/漁會
           		  	 	  }else{
           		  	 	      bean = (DataObject)dbData_All.get(i);
           		  	 	      hsien_name = String.valueOf(bean.getValue("hsien_name"));//單位名稱
           		  	 	      count_seq = String.valueOf(bean.getValue("count_seq"));
           		  	 	  }
                          //存款餘額表
                          for(int cellcount=2;cellcount<11;cellcount++){
                              row = sheet.getRow(rowNum);
                              cell= row.getCell((short)cellcount);
          	  	 		      cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
          		          	  insertValue = "";
                              insertValue = getBeanData(bean,dataIdx0,cellcount-2);//存款餘額~公庫存款
                              System.out.println("rowNum="+rowNum+":"+insertValue);
                              if(!insertValue.equals("0")) cell.setCellValue(insertValue);
                          }//end of 存款餘額表
                          //放款總額表
                          for(int cellcount=2;cellcount<10;cellcount++){
                              if(bank_type.equals("6") || bank_type.equals("ALL")){//農會.農漁會
                              	 if(Integer.parseInt(s_year) <= 99){
                                    row = sheet.getRow(rowNum+37);//農會.農漁會
                              	 }else{
                              		if(bank_type.equals("6")){
                              		   row = sheet.getRow(rowNum+34);//農會
                              		}else if(bank_type.equals("ALL")){
                              		   row = sheet.getRow(rowNum+35);//農漁會//109.07.27 add 金門縣
                              		}
                              	 }
                              }else{
                              	 if(Integer.parseInt(s_year) <= 99){
                                    row = sheet.getRow(rowNum+25);//漁會
                              	 }else{
                              	 	row = sheet.getRow(rowNum+24);//漁會//109.07.27 add 金門縣
                              	 }
                              }
                              cell=row.getCell((short)cellcount);
                              cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                              insertValue = "";
                              insertValue = getBeanData(bean,dataIdx1,cellcount-2);//放款總額~催收款
                              System.out.println("rowNum="+rowNum+":"+insertValue);
                              if(!insertValue.equals("0")) cell.setCellValue(insertValue);
                          }//end of 放款總額表
                          ++rowNum;
                      }//end of dbData_All
                  }//end of 總表有資料
              
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

  
}
