// 94.03.10 增加機構類別判斷及更改檔名FR006W為FR006W處理 add by egg
// 94.03.15 增加輸出報表產生日期資料及fix int金額長度不足改為long處理
// 94.07.07 將"o" 能以實際數值表示 fix by 2495
// 97.01.30 add 990612非會員政策性農業專案貸款;
//		         非會員授信總額占非會員存款總額比率=990610-990511-990611-990612 by 2295
// 99.01.25 fix 計算上年度信用部決算淨值/全体農(漁)會決算淨值時,需扣除992810投資全國農業金庫尚未攤提之損失[99.1月起] by 2295
//          fix 94.10以後.10~13公式.漁會同農會
//100.06.24 fix 根據查詢年度.100年以後取得新縣市別.100年以前取得舊縣市別 
// 			           使用PreparedStatement;並列印轉換後的SQL;套用QueryDB_SQLParam by 2295
//101.08.27 add 違反信用部固定資產淨額不得超過上年度信用部決算淨值不在此限的原因項目：  by 2295
//              990812一、因購置或汰換安全維護或營業相關設備，經中央主管機關核准
//              990813二、因固定資產重估增值
//              990814三、因淨值降低
//102.01.16 add (四)1.信用部逾放比< 2%  且 BIS > 8%者 得不超過 150%
//                 2.信用部逾放比< 1%  且 BIS > 10%且備抵呆帳覆蓋率高100%,已申請經主管機關同意者,得不超過200%
//          add (六)逾放比低於2%且資本適足率高8%,已申請經主管機關同意者,得不超過150% by 2295
//103.01.06 add (二)2.1/2.2及(八)上年度信用部決算淨值為負數時,要顯示違反  by 2295
//          add (九)逾新台幣100萬且超過前一年度信用部決算淨值5%,才算違反 by 2295
//103.01.13 add %顯示至小數點2位,不足者補0 by 2295 
//104.02.24 add (一)月底存放比不得超過80%之規定,取消(三)限制
//          add (四)信用部逾放比< 1%  且 BIS > 10%且備抵呆帳覆蓋率高100%且 > 全體信用部備抵呆帳覆蓋率平均值且備呆占放款比率> 2%,已申請經主管機關同意者,不受限制) by 2295          
//          add (六)逾放比< 1% 且 BIS > 10% 且備抵呆帳覆蓋率高100%且 > 全體信用部備抵呆帳覆蓋率平均值且備呆占放款比率> 2%,已申請經主管機關同意者,得不超過 200%) by 2295       
//          add (七)調整為50% by 2295
//          add (十四)對鄉(鎮、市)公所授信未經其所隸屬之縣政府保證之限額 by 2295
//104.02.25 add 所有比率調整跟DS002W一樣.小數點取至2位.四捨5入 by 2295
//105.03.24 add 縣市政府列印只顯示其轄區下的有違反的農漁會信用部明細資料 by 2295              
//106.01.19 fix 原3.非會員存款之額度限制.取消顯示;調整4.6title顯示名稱 by 2295 
//106.05.19 add 若逾期放款為0,備抵呆帳覆蓋率(field_backup_over_rate=(120800+150300)/99000,因分母為0,則該比率顯示為N/A,調整(四)符合為不受限制及200%/(六)符合200% by 2295
//106.10.06 add (五)無擔保消費性貸款調整為990510非會員無擔保消費性貸款 -990511非會員無擔保消費性政策貸款   by 2295
//          add (六)非會員授信總額.移除扣除990511非會員無擔保消費性政策貸款=非會員授信總額(990610) - (990611) - (990612)] by 2295
//107.04.09 fix (四)不受限制,取消 備抵呆帳覆蓋率 > 全體信用部備抵呆帳覆蓋率平均值及不<100%條件檢核 by 2295
//              (六)200%,取消 備抵呆帳覆蓋率 > 全體信用部備抵呆帳覆蓋率平均值及不<100%條件檢核  by 2295
//          add (四)/(六)增加檢核條件不符合時,顯示違反 by 2295
//108.09.09 add 108年10月以後,調整為 6.違反購置住宅放款及房屋修繕放款之餘額不得超過存款總餘額55%之規定 by 2295
//110.03.03 add 調整 3.贊助會員授信總額占贊助會員存款總額之比率
//              原5.非會員授信總額占非會員存款總額之比率,調整為4.非會員授信總額占非會員存款總額之比率
//              原4.辦理非會員無擔保消費性貸款（1,000千元以下）之限制,調整為7.辦理非會員無擔保消費性貸款（1,000千元以下）之限制
//              其餘項目往後順延 
//110.05.13 fix 調整110/5套用新的報表格式 by 2295
//110.09.03 fix 10.外幣風險之限制-逾新台幣100萬且超過前一年度信用部決算淨值10%才算違反 by 2295
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*;
import java.io.*;
import java.text.DecimalFormat;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;
import com.tradevan.util.report.HssfStyle;

public class FR006W_Excel {
    public static String createRpt(String S_YEAR,String S_MONTH,String bank_type){
        System.out.println("FR006W_Excel Start ...");
        String errMsg = "";
        List dbData = null;        
        StringBuffer sqlCmd = new StringBuffer();
        List paramList = new ArrayList();
        String cd01_table = "";
        String wlx01_m_year = "";
        Properties A02Data = new Properties();
        String acc_code = "";
        String amt = "";
        DecimalFormat df_md = new DecimalFormat("############0.00");//顯示小數點至第2位,不足者補0
        double tmp_A = 0.0;
        double tmp_B = 0.0;
        String hsien_id_b="";//縣市政府所屬縣市別 105.03.24 add
        //String bank_type="";
        try {
          //100.06.24 add 查詢年度100年以前.縣市別不同===============================
  	      cd01_table = (Integer.parseInt(S_YEAR) < 100)?"cd01_99":""; 
  	      wlx01_m_year = (Integer.parseInt(S_YEAR) < 100)?"99":"100"; 
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
          
          if(bank_type.startsWith("B")){//105.03.24 add 縣市政府,則為全部轄區下的農漁會信用部                   
              hsien_id_b=bank_type.substring(2,bank_type.length());
              bank_type="ALL";
              System.out.println("FR006W.hsien_id="+hsien_id_b);
          }
        
          FileInputStream finput = null;
          if (bank_type.equals("6") || bank_type.equals("ALL")) {
             if(Integer.parseInt(S_YEAR+S_MONTH) <= 10809){
    		    System.out.println("信用部違反法定比率規定分析表-農會.xls");
    		    finput = new FileInputStream(xlsDir + System.getProperty("file.separator") + "信用部違反法定比率規定分析表-農會.xls");
             }else if(Integer.parseInt(S_YEAR+S_MONTH) >= 10810 && Integer.parseInt(S_YEAR+S_MONTH) <= 11003){
       		    System.out.println("信用部違反法定比率規定分析表-農會_10810.xls");
       		    finput = new FileInputStream(xlsDir + System.getProperty("file.separator") + "信用部違反法定比率規定分析表-農會_10810.xls");
       	    }else if(Integer.parseInt(S_YEAR+S_MONTH) >= 11005){
       		    System.out.println("信用部違反法定比率規定分析表-農會_11004.xls");
       		    finput = new FileInputStream(xlsDir + System.getProperty("file.separator") + "信用部違反法定比率規定分析表-農會_11004.xls");
       	    }
          }else if(bank_type.equals("7")) {
        	  if(Integer.parseInt(S_YEAR+S_MONTH) < 9410){
        		  System.out.println("信用部違反法定比率規定分析表-漁會.xls");
                  finput = new FileInputStream(xlsDir + System.getProperty("file.separator") + "信用部違反法定比率規定分析表-漁會.xls");
          	  }else if(Integer.parseInt(S_YEAR+S_MONTH) >= 9410 && Integer.parseInt(S_YEAR+S_MONTH) <= 10809){
          		  System.out.println("信用部違反法定比率規定分析表-漁會_9410.xls");
          		  finput = new FileInputStream(xlsDir + System.getProperty("file.separator") + "信用部違反法定比率規定分析表-漁會_9410.xls");
        	  }else if(Integer.parseInt(S_YEAR+S_MONTH) >= 10810 && Integer.parseInt(S_YEAR+S_MONTH) <= 11003){
        		  System.out.println("信用部違反法定比率規定分析表-漁會_10810.xls");
        		  finput = new FileInputStream(xlsDir + System.getProperty("file.separator") + "信用部違反法定比率規定分析表-漁會_10810.xls");
        	  }else if(Integer.parseInt(S_YEAR+S_MONTH) >= 11005){
        		  System.out.println("信用部違反法定比率規定分析表-漁會_11004.xls");
        		  finput = new FileInputStream(xlsDir + System.getProperty("file.separator") + "信用部違反法定比率規定分析表-漁會_11004.xls");
        	  }	  
          }
          
               
          //設定FileINputStream讀取Excel檔
          POIFSFileSystem fs = new POIFSFileSystem(finput);
          HSSFWorkbook wb = new HSSFWorkbook(fs);
          HSSFSheet sheet = wb.getSheetAt(0); //讀取第一個工作表，宣告其為sheet
          HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定

          HSSFCellStyle leftStyle_unitName = wb.createCellStyle(); //有框內文置左
          leftStyle_unitName = HssfStyle.setStyle( leftStyle_unitName,wb.createFont() ,
                                                        new String[] {
                                                        "BORDER", "PHL", "PVC", "C09",
                                                        "WRAP"} );

          HSSFCellStyle centerStyle_value = wb.createCellStyle(); //有框內文置左
          centerStyle_value = HssfStyle.setStyle( centerStyle_value,wb.createFont() ,
                                                   new String[] {
                                                   "BORDER", "PHC", "PVC", "C09",
                                                   "WRAP"} );

          //設定頁面符合列印大小
          sheet.setAutobreaks(false);
          ps.setScale( (short) 70); //列印縮放百分比

          ps.setPaperSize( (short) 9); //設定紙張大小 A4
          //wb.setSheetName(0,"test");
          finput.close();

          HSSFRow row = null; //宣告一列
          HSSFCell cell = null; //宣告一個儲存格

          short i = 0;
          short y = 0;
          
          sqlCmd.append(" select a02.HSIEN_ID, nvl(a02.hsien_div_1,' ') as hsien_div_1,a02.bank_type,a02.bank_code as bank_no,bank_name,");
          sqlCmd.append("      amt990110,amt990120,amt990130,amt990140,amt990150,amt990210,amt990230,amt990240,amt990220,");
          sqlCmd.append("      amt990310,amt990630,amt990320,amt990410,amt990420,amt990421,amt990422,amt990422_limit,amt990510,amt990512,amt990610,amt990511,amt990611,amt990612,");
          sqlCmd.append("      amt990620,amt990621,amt990622,amt990623,amt990623_limit,amt990710,amt990720,amt990810,amt990811,amt990812,amt990813,amt990814,amt990910,amt990920,amt991010,amt991020,amt991030,amt991110,");
          sqlCmd.append("      amt991120,amt991210,amt991220,amt991310,amt991320,amt992810,amt996114,amt996115,amt990711,amt990712,");
          sqlCmd.append("      a05.amt as a05bis,");
          sqlCmd.append("      a01_operation.field_over_rate as a01field_over_rate,");
          sqlCmd.append("      a01_operation.field_backup_over_rate as a01field_backup_over_rate,");//備呆占狹義逾期放款比率
          sqlCmd.append("      a01_operation.field_over as a01field_over,");
          sqlCmd.append("      a01_operation.field_backup as a01field_backup,");
          sqlCmd.append("      a01_operation.field_backup_credit_rate as a01field_backup_credit_rate,");//104.02.24 add 備呆占放款比率
          sqlCmd.append("      a01_operation_sum.field_backup_over_rate as a01field_backup_over_rate_avg,");//104.02.24 add全體信用部備呆占狹義逾期放款比率平均值
          sqlCmd.append("	   a01_operation.fieldi_y as a01fieldi_y"); //108.09.09 add 存放比率-存款總餘額	   
          sqlCmd.append(" from  ");
          sqlCmd.append("      (select wlx01.HSIEN_ID,wlx01.hsien_div_1,bn01.bank_type,a02.bank_code,bn01.bank_name, ");
          sqlCmd.append("              sum(decode(a02.acc_code,'990110',amt,0)) as  amt990110,");
          sqlCmd.append("              sum(decode(a02.acc_code,'990120',amt,0)) as  amt990120,");
          sqlCmd.append("              sum(decode(a02.acc_code,'990130',amt,0)) as  amt990130,");
          sqlCmd.append("              sum(decode(a02.acc_code,'990140',amt,0)) as  amt990140,");
          sqlCmd.append("              sum(decode(a02.acc_code,'990150',amt,0)) as  amt990150,");
          sqlCmd.append("              sum(decode(a02.acc_code,'990210',amt,0)) as  amt990210,");
          sqlCmd.append("              sum(decode(a02.acc_code,'990230',amt,0)) as  amt990230,");
          sqlCmd.append("              sum(decode(a02.acc_code,'990240',amt,0)) as  amt990240,");
          sqlCmd.append("              sum(decode(a02.acc_code,'990220',amt,0)) as  amt990220,");
          sqlCmd.append("              sum(decode(a02.acc_code,'990310',amt,0)) as  amt990310,");
          sqlCmd.append("              sum(decode(a02.acc_code,'990630',amt,0)) as  amt990630,");
          sqlCmd.append("              sum(decode(a02.acc_code,'990320',amt,0)) as  amt990320,");
          sqlCmd.append("              sum(decode(a02.acc_code,'990410',amt,0)) as  amt990410,");
          sqlCmd.append("              sum(decode(a02.acc_code,'990420',amt,0)) as  amt990420,");
          sqlCmd.append("              sum(decode(a02.acc_code,'990421',amt,0)) as  amt990421,");//102.01.15 add
          sqlCmd.append("              sum(decode(a02.acc_code,'990422',amt,0)) as  amt990422,");//104.02.24 add
          sqlCmd.append("              sum(decode(a02.acc_code,'990422',decode(NVL(amt_name2,''),'',NVL(amt_name1,0),999))) as amt990422_limit,");//110.03.03 add
          sqlCmd.append("              sum(decode(a02.acc_code,'990510',amt,0)) as  amt990510,");//106.10.06 add
          sqlCmd.append("              sum(decode(a02.acc_code,'990512',amt,0)) as  amt990512,");
          sqlCmd.append("              sum(decode(a02.acc_code,'990610',amt,0)) as  amt990610,");
          sqlCmd.append("              sum(decode(a02.acc_code,'990511',amt,0)) as  amt990511,");
          sqlCmd.append("              sum(decode(a02.acc_code,'990611',amt,0)) as  amt990611,");
          sqlCmd.append("              sum(decode(a02.acc_code,'990612',amt,0)) as  amt990612,");//97.01.29 add
          sqlCmd.append("              sum(decode(a02.acc_code,'990620',amt,0)) as  amt990620,");
          sqlCmd.append("              sum(decode(a02.acc_code,'990621',amt,0)) as  amt990621,");//102.01.15 add
          sqlCmd.append("              sum(decode(a02.acc_code,'990622',amt,0)) as  amt990622,");//104.02.24 add
          sqlCmd.append("              sum(decode(a02.acc_code,'990623',amt,0)) as  amt990623,");//110.03.03 add
   	      sqlCmd.append("              sum(decode(a02.acc_code,'990623',decode(NVL(amt_name2,''),'',NVL(amt_name1,0),999))) as amt990623_limit,");//110.03.04 add
          sqlCmd.append("              sum(decode(a02.acc_code,'990710',amt,0)) as  amt990710,");
          sqlCmd.append("              sum(decode(a02.acc_code,'990720',amt,0)) as  amt990720,");
          sqlCmd.append("              sum(decode(a02.acc_code,'990810',amt,0)) as  amt990810,");
          sqlCmd.append("              sum(decode(a02.acc_code,'990811',amt,0)) as  amt990811,");//101.08.23 add
          sqlCmd.append("              sum(decode(a02.acc_code,'990812',amt,0)) as  amt990812,");//101.08.23 add
          sqlCmd.append("              sum(decode(a02.acc_code,'990813',amt,0)) as  amt990813,");//101.08.23 add
          sqlCmd.append("              sum(decode(a02.acc_code,'990814',amt,0)) as  amt990814,");//101.08.23 add
          sqlCmd.append("              sum(decode(a02.acc_code,'990910',amt,0)) as  amt990910,");
          sqlCmd.append("              sum(decode(a02.acc_code,'990920',amt,0)) as  amt990920,");
          sqlCmd.append("              sum(decode(a02.acc_code,'991010',amt,0)) as  amt991010,");//102.01.16 add
          sqlCmd.append("              sum(decode(a02.acc_code,'991020',amt,0)) as  amt991020,");
          sqlCmd.append("              sum(decode(a02.acc_code,'991030',amt,0)) as  amt991030,");//102.01.16 add          
          sqlCmd.append("              sum(decode(a02.acc_code,'991110',amt,0)) as  amt991110,");
          sqlCmd.append("              sum(decode(a02.acc_code,'991120',amt,0)) as  amt991120,");
          sqlCmd.append("              sum(decode(a02.acc_code,'991210',amt,0)) as  amt991210,");
          sqlCmd.append("              sum(decode(a02.acc_code,'991220',amt,0)) as  amt991220,");
          sqlCmd.append("              sum(decode(a02.acc_code,'991310',amt,0)) as  amt991310,");
          sqlCmd.append("              sum(decode(a02.acc_code,'991320',amt,0)) as  amt991320,");
          sqlCmd.append("              sum(decode(a02.acc_code,'992810',amt,0)) as  amt992810,");
          sqlCmd.append("              sum(decode(a02.acc_code,'996114',amt,0)) as  amt996114,");//104.02.24 add
          sqlCmd.append("              sum(decode(a02.acc_code,'996115',amt,0)) as  amt996115,");//104.02.24 add
          sqlCmd.append("              sum(decode(a02.acc_code,'990711',amt,0)) as  amt990711,");//108.09.09 add
          sqlCmd.append("              sum(decode(a02.acc_code,'990712',amt,0)) as  amt990712 ");//108.09.09 add
          sqlCmd.append("      from (select * from bn01 where m_year=? and bank_type in (?,?))bn01 left join (select * from a02 where a02.m_year=? and a02.m_month=?");
          paramList.add(wlx01_m_year);
          if(bank_type.equals("ALL")){
              paramList.add("6");//105.03.24 add 
              paramList.add("7");//105.03.24 add 
          }else{
             paramList.add(bank_type);
             paramList.add(bank_type);//105.03.24 add
          }
          paramList.add(S_YEAR);
          paramList.add(S_MONTH);
          sqlCmd.append("                             union");
          sqlCmd.append("                             select  m_year,m_month,bank_code,acc_code,amt,'','','' from a99 where a99.m_year=? and a99.m_month=?");
          paramList.add(S_YEAR);
          paramList.add(S_MONTH);
          sqlCmd.append("                            )a02 on a02.bank_code = bn01.bank_no");
          if(bank_type.equals("ALL")){//105.03.24 add
             sqlCmd.append("      left join (select * from wlx01 where m_year=? and m2_name=?)wlx01 on bn01.bank_no = wlx01.BANK_NO");
             paramList.add(wlx01_m_year);
             paramList.add(hsien_id_b);
          }else{
             sqlCmd.append("      left join (select * from wlx01 where m_year=?)wlx01 on bn01.bank_no = wlx01.BANK_NO");
             paramList.add(wlx01_m_year);
          }
          
          sqlCmd.append("      where bn01.bank_type in (?,?)");// and bn01.bank_no in('6220077')"
          if(bank_type.equals("ALL")){
             paramList.add("6");//105.03.24 add 
             paramList.add("7");//105.03.24 add 
          }else{
             paramList.add(bank_type);
             paramList.add(bank_type);//105.03.24 add
          }
          sqlCmd.append("      and wlx01.HSIEN_ID is not null ");
          sqlCmd.append("      group by wlx01.HSIEN_ID, wlx01.hsien_div_1,bn01.bank_type,a02.bank_code,bn01.bank_name");
          sqlCmd.append("      order by wlx01.HSIEN_ID, wlx01.hsien_div_1,bn01.bank_type,a02.bank_code,bn01.bank_name)a02");
          sqlCmd.append("   ,(select m_year,m_month,bank_code,acc_code,'4' as type,round(amt/1000,3) as amt,'' as violate from a05 ");
          sqlCmd.append("    where  a05.m_year =? and a05.m_month = ?");
          paramList.add(S_YEAR);
          paramList.add(S_MONTH);
          sqlCmd.append("    and a05.acc_code in ('91060P'))a05");
          sqlCmd.append("    ,(select m_year,m_month,bank_code,");
          sqlCmd.append("     sum(decode(acc_code,'field_over_rate',amt,0)) as  field_over_rate,");
          sqlCmd.append("     sum(decode(acc_code,'field_backup_over_rate',amt,0)) as  field_backup_over_rate, ");
          sqlCmd.append("     sum(decode(acc_code,'field_over',amt,0)) as  field_over, ");//逾放金額  102.09.23 add
          sqlCmd.append("     sum(decode(acc_code,'field_backup',amt,0)) as  field_backup, "); //備抵呆帳 102.09.23 add
          sqlCmd.append("     sum(decode(acc_code,'field_backup_credit_rate',amt,0)) as  field_backup_credit_rate, "); //備呆占放款比率 104.02.24 add 
          sqlCmd.append("     sum(decode(acc_code,'fieldi_y',amt,0)) as  fieldi_y "); //存放比率-存款總餘額 108.09.09 add
          sqlCmd.append("     from a01_operation ");
          sqlCmd.append("     where m_year=? and m_month=?");
          paramList.add(S_YEAR);
          paramList.add(S_MONTH);
          sqlCmd.append("     and acc_code in('field_over_rate','field_backup_over_rate','field_over','field_backup','field_backup_credit_rate','fieldi_y') ");
          sqlCmd.append("     group by m_year,m_month,bank_code  ) a01_operation, ");
          sqlCmd.append("    (select m_year,m_month,");      
          sqlCmd.append("     sum(decode(acc_code,'field_backup_over_rate',amt,0)) as  field_backup_over_rate ");//104.02.24 add備呆占狹義逾期放款比率         
          sqlCmd.append("     from a01_operation ");
          sqlCmd.append("     where m_year=? and m_month=?");
          paramList.add(S_YEAR);
          paramList.add(S_MONTH);
          sqlCmd.append("     and acc_code in('field_backup_over_rate') ");
          sqlCmd.append("     and bank_type='ALL' ");
          sqlCmd.append("     and bank_code='ALL' ");
          sqlCmd.append("     and hsien_id=' '");
          sqlCmd.append("     group by m_year,m_month ) a01_operation_sum ");         
          sqlCmd.append("  where a02.bank_code = a05.bank_code ");
          sqlCmd.append("    and a02.bank_code=a01_operation.bank_code ");

       dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt990110,amt990120,amt990130,amt990140,amt990150,amt990210,amt990230,amt990240,amt990220,"
                   + "amt990310,amt990630,amt990320,amt990410,amt990420,amt990421,amt990422,amt990422_limit,amt990510,amt990512,amt990610,amt990511,amt990611,amt990612,"
                   + "amt990620,amt990621,amt990622,amt990623,amt990623_limit,amt990710,amt990720,amt990810,amt990811,amt990812,amt990813,amt990814,amt990910,amt990920,amt991010,amt991020,amt991030,amt991110,"
                   + "amt991120,amt991210,amt991220,amt991310,amt991320,amt992810,amt996114,amt996115,amt990711,amt990712,a05bis,a01field_over_rate,a01field_backup_over_rate,a01field_over,a01field_backup,a01field_backup_credit_rate,a01field_backup_over_rate_avg,a01fieldi_y");
       System.out.println("dbData.size()="+dbData.size());
          
          
          /*102.01.16 舊的SQL
          sqlCmd.append("select a.bank_code,c.bank_name,b.hsien_div_1 ,");
          sqlCmd.append("       sum(decode(a.acc_code,'990110',amt,0)) \"990110\", ");
          sqlCmd.append("       sum(decode(a.acc_code,'990120',amt,0)) \"990120\", ");
          sqlCmd.append("       sum(decode(a.acc_code,'990130',amt,0)) \"990130\", ");
          sqlCmd.append("       sum(decode(a.acc_code,'990140',amt,0)) \"990140\", ");
          sqlCmd.append("       sum(decode(a.acc_code,'990150',amt,0)) \"990150\", ");
          sqlCmd.append("       sum(decode(a.acc_code,'990210',amt,0)) \"990210\", ");
          sqlCmd.append("       sum(decode(a.acc_code,'990220',amt,0)) \"990220\", ");
          sqlCmd.append("       sum(decode(a.acc_code,'990230',amt,0)) \"990230\", ");
          sqlCmd.append("       sum(decode(a.acc_code,'990240',amt,0)) \"990240\", ");
          sqlCmd.append("       sum(decode(a.acc_code,'990310',amt,0)) \"990310\", ");
          sqlCmd.append("       sum(decode(a.acc_code,'990320',amt,0)) \"990320\", ");
          sqlCmd.append("       sum(decode(a.acc_code,'990410',amt,0)) \"990410\", ");
          sqlCmd.append("       sum(decode(a.acc_code,'990420',amt,0)) \"990420\", ");
          sqlCmd.append("       sum(decode(a.acc_code,'990421',amt,0)) \"990421\", ");//102.01.16 add
          sqlCmd.append("       sum(decode(a.acc_code,'990510',amt,0)) \"990510\", ");
          sqlCmd.append("       sum(decode(a.acc_code,'990511',amt,0)) \"990511\", ");//97.01.30 add
          sqlCmd.append("       sum(decode(a.acc_code,'990610',amt,0)) \"990610\", ");
          sqlCmd.append("       sum(decode(a.acc_code,'990611',amt,0)) \"990611\", ");//97.01.30 add
          sqlCmd.append("       sum(decode(a.acc_code,'990612',amt,0)) \"990612\", ");//97.01.30 add
          sqlCmd.append("       sum(decode(a.acc_code,'990620',amt,0)) \"990620\", ");
          sqlCmd.append("       sum(decode(a.acc_code,'990621',amt,0)) \"990621\", ");//102.01.16 add
          sqlCmd.append("       sum(decode(a.acc_code,'990630',amt,0)) \"990630\", ");
          sqlCmd.append("       sum(decode(a.acc_code,'990710',amt,0)) \"990710\", ");
          sqlCmd.append("       sum(decode(a.acc_code,'990720',amt,0)) \"990720\", ");
          sqlCmd.append("       sum(decode(a.acc_code,'990810',amt,0)) \"990810\", ");
          sqlCmd.append("       sum(decode(a.acc_code,'990812',amt,0)) \"990812\", ");//101.08.27 add
          sqlCmd.append("       sum(decode(a.acc_code,'990813',amt,0)) \"990813\", ");//101.08.27 add
          sqlCmd.append("       sum(decode(a.acc_code,'990814',amt,0)) \"990814\", ");//101.08.27 add      
          sqlCmd.append("       sum(decode(a.acc_code,'990910',amt,0)) \"990910\", ");
          sqlCmd.append("       sum(decode(a.acc_code,'990920',amt,0)) \"990920\", ");
          sqlCmd.append("       sum(decode(a.acc_code,'990930',amt,0)) \"990930\", ");
          sqlCmd.append("       sum(decode(a.acc_code,'991010',amt,0)) \"991010\", ");
          sqlCmd.append("       sum(decode(a.acc_code,'991020',amt,0)) \"991020\", ");
          sqlCmd.append("       sum(decode(a.acc_code,'991030',amt,0)) \"991030\", ");
          sqlCmd.append("       sum(decode(a.acc_code,'991110',amt,0)) \"991110\", ");
          sqlCmd.append("       sum(decode(a.acc_code,'991120',amt,0)) \"991120\", ");
          sqlCmd.append("       sum(decode(a.acc_code,'991210',amt,0)) \"991210\", ");
          sqlCmd.append("       sum(decode(a.acc_code,'991220',amt,0)) \"991220\", ");
          sqlCmd.append("       sum(decode(a.acc_code,'991310',amt,0)) \"991310\", ");
          sqlCmd.append("       sum(decode(a.acc_code,'991320',amt,0)) \"991320\", ");
          sqlCmd.append("       sum(decode(a.acc_code,'992810',amt,0)) \"992810\", ");//99.01.22 992810投資全國農業金庫尚未攤提之損失[99.1月起]
          sqlCmd.append("       sum(decode(a.acc_code,null,0,'99141Y',amt,0)) \"99141Y\"  "); 
          sqlCmd.append(" from (select m_year,m_month,bank_code,acc_code,amt from A02 where m_year=? and m_month=? ");              
          sqlCmd.append("       union ");
          sqlCmd.append("       select * from A99 where m_year=? and m_month=?");            
          sqlCmd.append("       )a,(select * from wlx01 where m_year=?)b,(select * from bn01 where m_year=?)c ");
          sqlCmd.append("       ,(select * from v_bank_location where m_year=?)e");
          sqlCmd.append("   ,(select m_year,m_month,bank_code,acc_code,'4' as type,round(amt/1000,3) as amt,'' as violate from a05 ");
          sqlCmd.append("    where  a05.m_year =? and a05.m_month = ?");         
          sqlCmd.append("    and a05.acc_code in ('91060P'))a05");
          sqlCmd.append("    ,(select m_year,m_month,bank_code,");
          sqlCmd.append("     sum(decode(acc_code,'field_over_rate',amt,0)) as  field_over_rate,");
          sqlCmd.append("     sum(decode(acc_code,'field_backup_over_rate',amt,0)) as  field_backup_over_rate ");  
          sqlCmd.append("     from a01_operation ");
          sqlCmd.append("     where m_year=? and m_month=?");         
          sqlCmd.append("     and acc_code in('field_over_rate','field_backup_over_rate') ");
          sqlCmd.append("     group by m_year,m_month,bank_code  ) a01_operation ");
          sqlCmd.append(" where a.bank_code = b.bank_no ");
          sqlCmd.append("  and  a.bank_code = c.bank_no ");
          sqlCmd.append("  and  a.bank_code = e.bank_no ");
          sqlCmd.append("  and  a.bank_code = a05.bank_code ");
          sqlCmd.append("  and  a.bank_code = a01_operation.bank_code ");
          sqlCmd.append("  and  c.bank_type = ? ");
          sqlCmd.append(" group by a.bank_code,c.bank_name,b.hsien_div_1 ,e.fr001w_output_order");
          sqlCmd.append(" order by e.fr001w_output_order,bank_code ");
          
          paramList.add(S_YEAR);
          paramList.add(S_MONTH);
          paramList.add(S_YEAR);
          paramList.add(S_MONTH);
          paramList.add(wlx01_m_year);
          paramList.add(wlx01_m_year);
          paramList.add(wlx01_m_year);
          paramList.add(S_YEAR);
          paramList.add(S_MONTH);
          paramList.add(S_YEAR);
          paramList.add(S_MONTH);
          paramList.add(bank_type);
                     
          dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList, "m_year,m_month,amt,990110,990120,990130,990140,990150,990210,990220,990230,990240,990310,990320,990410,990420,990421,990510,990511,990610,990611,990612,990620,990621,990630,990710,990720,990810,990812,990813,990814,990910,990920,990930,991010,991020,991030,991110,991120,991210,991220,991310,991320,992810,99141Y");
          System.out.println("dbData.size=" + dbData.size());
          */
          //取得當前日期
          Calendar rightNow = Calendar.getInstance();
          String year = String.valueOf(rightNow.get(Calendar.YEAR) - 1911);
          String month = String.valueOf(rightNow.get(Calendar.MONTH) + 1);
          String day = String.valueOf(rightNow.get(Calendar.DAY_OF_MONTH));
          if (dbData.size() == 0) {
            row = sheet.getRow(0);
            cell = row.getCell( (short) 2);
            //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            //設定無資料
            cell.setCellValue(S_YEAR + "年度" + S_MONTH + "月份無資料存在");
            row = sheet.getRow(2);
            cell = row.getCell( (short) 0);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellValue("列印日期：" + year + "年" + month + "月" + day + "日");
          }else {
            //設定報表表頭資料
            row = sheet.getRow(0);
            cell = row.getCell( (short) 2);
           
            //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            //bank_type = (String)((DataObject)dbData.get(0)).getValue("bank_type");
            if (bank_type.equals("6")) {
              cell.setCellValue(S_YEAR + "年度" + S_MONTH + "月份" +"信用部違反法定比率規定分析表-農會");
            } else if (bank_type.equals("7")) {
              cell.setCellValue(S_YEAR + "年度" + S_MONTH + "月份" +"信用部違反法定比率規定分析表-漁會");
            } else {
                cell.setCellValue(S_YEAR + "年度" + S_MONTH + "月份" +"信用部違反法定比率規定分析表-農漁會");
            }

            row = sheet.getRow(2);
            cell = row.getCell( (short) 0);
            //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellValue("列印日期：" + year + "年" + month + "月" + day + "日");

            String Bank_type = "", Bank_no = "", bank_name = "", areaNo = "";

			//94.07.07  fix by 2495						
            double[][]  flag = {{0.0, 0.0}, {0.0, 0.0}, {0.0, 0.0}, {0.0, 0.0}, {0.0, 0.0}, {0.0, 0.0}, {0.0, 0.0},{0.0, 0.0}, {0.0, 0.0}, {0.0, 0.0}, {0.0, 0.0}, 
            		            {0.0, 0.0}, {0.0, 0.0}, {0.0, 0.0}, {0.0, 0.0}, {0.0, 0.0}, {0.0, 0.0}, {0.0, 0.0}};
            int j = 0;
            int getRow = 0;
            double a = 0, b = 0;
            
						
            //設定違反事項狀態資料


            //以巢狀迴圈處理所有儲存格資料
            System.out.println("total row =" + sheet.getLastRowNum());
            System.out.println("dbData.size=" + dbData.size());
            int countAllItem=0;
            //94.07.07  fix by 2495
            long amt_990110=0;
            long amt_990120=0;
            long amt_990130=0;
            long amt_990140=0;
            long amt_990150=0;
            long amt_990210=0;
            long amt_990230=0;
            long amt_990240=0;
            long amt_990220=0;
            long amt_990310=0;
            long amt_990630=0;
            long amt_990320=0;
            long amt_990410=0;
            long amt_990420=0;
            long amt_990421=0;//102.01.15 add
            long amt_990422=0;//104.02.24 add
            long amt_990422_limit=0;//110.03.03 add
            long amt_990510=0;//106.10.06 add
            long amt_990512=0;
            long amt_990610=0;
            long amt_990511=0;
            long amt_990611=0;
            long amt_990612=0;//97.01.29 add
            long amt_990620=0;
            long amt_990621=0;//102.01.15 add
            long amt_990622=0;//104.02.24 add
            long amt_990623=0;//110.03.03 add
            long amt_990623_limit=0;//110.03.03 add
            long amt_990710=0;
            long amt_990720=0;
            long amt_990810=0;
            long amt_990811=0;//101.08.23 add
            long amt_990812=0;//101.08.23 add
            long amt_990813=0;//101.08.23 add
            long amt_990814=0;//101.08.23 add
            long amt_990910=0;
            long amt_990920=0;
            long amt_991020=0;
            long amt_991010=0;//102.01.16 add
            long amt_991030=0;//102.01.16 add
            long amt_991110=0;
            long amt_991120=0;
            long amt_991210=0;
            long amt_991220=0;
            long amt_991310=0;
            long amt_991320=0;
            long amt_992810=0;//99.01.25 add
            long amt_996114=0;//104.02.24 add
            long amt_996115=0;//104.02.24 add
            long amt_990711=0;//108.09.09 add
    	    long amt_990712=0;//108.09.09 add 
            long a01field_over=0;//102.09.23 add
            long a01field_backup=0;//102.09.23 add
            double a05bis=0;
            double a01field_over_rate=0;
            double a01field_backup_over_rate=0;//102.01.15 add備呆占狹義逾期放款比率(備抵呆帳覆蓋率)
            double a01field_backup_credit_rate=0;//104.02.24 add備呆占放款比率
            double a01field_backup_over_rate_avg=0;//104.02.24 add備呆占狹義逾期放款比率平均值
            double a01fieldi_y=0;//108.09.09 add 存放比率-存款總餘額
            DataObject bean = null;
            for (i = 0; i < dbData.size(); i++) {
                bean=(DataObject)dbData.get(i);
            	for (j = 0; j <= 17; j++) {
                	flag[j][0] = 0.0;
                	flag[j][1] = 0.0;
                }		
           
              bank_name = (String) ( (DataObject) dbData.get(i)).getValue("bank_name");
              areaNo = ( (DataObject) dbData.get(i)).getValue("hsien_div_1") == null ? "" : (String) ( (DataObject) dbData.get(i)).getValue("hsien_div_1");             
              amt_990110 = Long.parseLong( ( bean.getValue("amt990110")).toString());
              amt_990120 = Long.parseLong( ( bean.getValue("amt990120")).toString());
              amt_990130 = Long.parseLong( ( bean.getValue("amt990130")).toString());
              amt_990140 = Long.parseLong( ( bean.getValue("amt990140")).toString());
              amt_990150 = Long.parseLong( ( bean.getValue("amt990150")).toString());              
              amt_990210 = Long.parseLong( ( bean.getValue("amt990210")).toString());
              amt_990230 = Long.parseLong( ( bean.getValue("amt990230")).toString());
              amt_990240 = Long.parseLong( ( bean.getValue("amt990240")).toString());
              amt_990220 = Long.parseLong( ( bean.getValue("amt990220")).toString());
              amt_990310 = Long.parseLong( ( bean.getValue("amt990310")).toString());
              amt_990630 = Long.parseLong( ( bean.getValue("amt990630")).toString());
              amt_990320 = Long.parseLong( ( bean.getValue("amt990320")).toString());
              amt_990410 = Long.parseLong( ( bean.getValue("amt990410")).toString());
              amt_990420 = Long.parseLong( ( bean.getValue("amt990420")).toString());
              amt_990421 = Long.parseLong( ( bean.getValue("amt990421")).toString());//102.01.15 add                
              amt_990422 = Long.parseLong( ( bean.getValue("amt990422")).toString());//104.02.24 add
              amt_990422_limit = bean.getValue("amt990422_limit") == null?0:Long.parseLong( ( bean.getValue("amt990422_limit")).toString());//110.03.03 add
              amt_990512 = Long.parseLong( ( bean.getValue("amt990510")).toString());//106.10.06 add
              amt_990512 = Long.parseLong( ( bean.getValue("amt990512")).toString());
              amt_990610 = Long.parseLong( ( bean.getValue("amt990610")).toString());
              amt_990511 = Long.parseLong( ( bean.getValue("amt990511")).toString());
              amt_990611 = Long.parseLong( ( bean.getValue("amt990611")).toString());
              amt_990612 = Long.parseLong( ( bean.getValue("amt990612")).toString());//97.01.29 add
              amt_990620 = Long.parseLong( ( bean.getValue("amt990620")).toString());
              amt_990621 = Long.parseLong( ( bean.getValue("amt990621")).toString());//102.01.15 add              
              amt_990622 = Long.parseLong( ( bean.getValue("amt990622")).toString());//104.02.24 add
              amt_990623 = Long.parseLong( ( bean.getValue("amt990623")).toString());//110.03.03 add              
              amt_990623_limit = bean.getValue("amt990623_limit") == null?0:Long.parseLong( ( bean.getValue("amt990623_limit")).toString());//110.03.03 add
              amt_990710 = Long.parseLong( ( bean.getValue("amt990710")).toString());
              amt_990720 = Long.parseLong( ( bean.getValue("amt990720")).toString());
              amt_990810 = Long.parseLong( ( bean.getValue("amt990810")).toString());
              amt_990811 = Long.parseLong( ( bean.getValue("amt990811")).toString());//101.08.23 add
              amt_990812 = Long.parseLong( ( bean.getValue("amt990812")).toString());//101.08.23 add
              amt_990813 = Long.parseLong( ( bean.getValue("amt990813")).toString());//101.08.23 add
              amt_990814 = Long.parseLong( ( bean.getValue("amt990814")).toString());//101.08.23 add
              amt_990910 = Long.parseLong( ( bean.getValue("amt990910")).toString());              
              amt_990920 = Long.parseLong( ( bean.getValue("amt990920")).toString());              
              amt_991020 = Long.parseLong( ( bean.getValue("amt991020")).toString());             
              amt_991010 = Long.parseLong( ( bean.getValue("amt991010")).toString());             
              amt_991030 = Long.parseLong( ( bean.getValue("amt991030")).toString());              
              amt_991110 = Long.parseLong( ( bean.getValue("amt991110")).toString());             
              amt_991120 = Long.parseLong( ( bean.getValue("amt991120")).toString());              
              amt_991210 = Long.parseLong( ( bean.getValue("amt991210")).toString());             
              amt_991220 = Long.parseLong( ( bean.getValue("amt991220")).toString());              
              amt_991310 = Long.parseLong( ( bean.getValue("amt991310")).toString());              
              amt_991320 = Long.parseLong( ( bean.getValue("amt991320")).toString());             
              amt_992810 = Long.parseLong( ( bean.getValue("amt992810")).toString());//99.01.25     
              amt_996114 = Long.parseLong( ( bean.getValue("amt996114")).toString());//104.02.24              
              amt_996115 = Long.parseLong( ( bean.getValue("amt996115")).toString());//104.02.24      
              amt_990711 = Long.parseLong( ( bean.getValue("amt990711")).toString());//108.09.09
              amt_990712 = Long.parseLong( ( bean.getValue("amt990712")).toString());//108.09.09
              a01field_over = Long.parseLong( ( bean.getValue("a01field_over")).toString());
              a01field_backup = Long.parseLong( ( bean.getValue("a01field_backup")).toString());              
              //long amt_99141Y = ( (DataObject) dbData.get(i)).getValue("99141Y") == null ? 0 : Long.parseLong( ( ( (DataObject) dbData.get(i)).getValue("99141Y")).toString());
              a05bis=0;
              a01field_over_rate =0;
              a01field_backup_over_rate = 0;
              a05bis = Double.parseDouble( ( ( (DataObject) dbData.get(i)).getValue("a05bis")).toString());
              a01field_over_rate = Double.parseDouble( ( ( (DataObject) dbData.get(i)).getValue("a01field_over_rate")).toString());
              a01field_backup_over_rate = Double.parseDouble( ( ( (DataObject) dbData.get(i)).getValue("a01field_backup_over_rate")).toString());
              a01field_backup_credit_rate = Double.parseDouble( ( bean.getValue("a01field_backup_credit_rate")).toString());//104.02.24 add 
              a01field_backup_over_rate_avg = Double.parseDouble( ( bean.getValue("a01field_backup_over_rate_avg")).toString());//104.02.24 add
          	  a01fieldi_y = Double.parseDouble( ( bean.getValue("a01fieldi_y")).toString());//108.09.09 add	
              //1.最近一年平均存放比率-->比率用『月底存放比率(A01)_合庫版』
              //104.02.24 一律調整為80%
              if ( (amt_990110 != 0) || (amt_990120 != 0) || (amt_990130 != 0) ||
                  (amt_990140 != 0) || (amt_990150 != 0)) {
                 //System.out.println("990110 ~ 990150不為0");
                if ( (amt_990140 - ( amt_990150  / 2)) != 0) {
                  if (amt_990120 >= amt_990130) {
                    a = (amt_990110 - amt_990120 + amt_990130) /
                        (amt_990140 - ( (amt_990150 * 1) / 2));
                  }else {
                    a = amt_990110 / (amt_990140 - ( (amt_990150 * 1) / 2));
                  }
                }else {
                  a = 0;
                }
                //104.02.24 一律調整為80%
                //if (areaNo.equals("1") && (!areaNo.equals(""))) {
                //  if (a >= 0.78) { //直轄市及省、縣轄市農會不得超過78%                    	
                //      flag[0][0] = 1; 
                //      flag[0][1] = a;                      
                //  }
                //}else {
                  //System.out.println("鄉鎮地區");
                  if (a >= 0.80) { //違反鄉鎮地區農會不得超過80%                    
                    flag[0][0] = 1; 
                    flag[0][1] = a;
                    System.out.println("debug---a="+a);
                    System.out.println("debug---flag[0][0]="+flag[0][0]);
                    System.out.println("debug---flag[0][1]="+flag[0][1]);
                  }
                //}
              }
              //2.違反農會信用部對農會經濟事業門融通資金之限制
              //103.01.06 add 上年度信用部決算淨值為負數時,要顯示違反
              a = 0;
              a = amt_990230 - amt_990240 - amt_992810; //上年度信用部決算淨值(扣除檢查應補提未提足之備抵呆帳)
              //99.01.25 fix 99.01開始.上年度信用部決算淨值扣除992810投資全國農業金庫尚未攤提之損失
              if (a != 0) {
                //System.out.println("990230-990240不為0");
                
                if ( (amt_990210 / a) > 0.6) { //違反內部融資不得超過信用部上年度決算淨值60%之規定                  
                  flag[1][0] = 1; 
                  //flag[1][1] = (amt_990210 /a);
                  tmp_A=((double)amt_990210)/((double)amt_990230 - (double)amt_990240 - (double)amt_992810);
                  tmp_B =Math.round( tmp_A * 10000);
                  tmp_A=((double)tmp_B)/100;                      
                  flag[1][1] = tmp_A;
                }
                if(a < 0){
                    flag[1][0] = 1; //103.01.06 add 上年度信用部決算淨值為負數時,要顯示違反
                    //flag[1][1] = (amt_990210 /a);
                    tmp_A=((double)amt_990210)/((double)amt_990230 - (double)amt_990240 - (double)amt_992810);
                    tmp_B =Math.round( tmp_A * 10000);
                    tmp_A=((double)tmp_B)/100; 
                    flag[1][1] = tmp_A;
                }
                if ( (amt_990220 / a) > 0.3) { //違反內部融資(中長期)不得超過信用部上年度決算淨值30%之規定                 
                  flag[2][0] = 1; 
                  //flag[2][1] = (amt_990220 /a);
                  tmp_A=((double)amt_990220)/((double)amt_990230 - (double)amt_990240 - (double)amt_992810);
                  tmp_B =Math.round( tmp_A * 10000);
                  tmp_A=((double)tmp_B)/100; 
                  flag[2][1] = tmp_A;
                }
                if(a < 0){
                    flag[2][0] = 1;//103.01.06 add 上年度信用部決算淨值為負數時,要顯示違反 
                    //flag[2][1] = (amt_990220 /a);
                    tmp_A=((double)amt_990220)/((double)amt_990230 - (double)amt_990240 - (double)amt_992810);
                    tmp_B =Math.round( tmp_A * 10000);
                    tmp_A=((double)tmp_B)/100; 
                    flag[2][1] = tmp_A;
                }
              }
              //3.違反非會員存款總額不得超過上年度農漁會全體決算淨值10倍之規定
              //99.01.25 fix 99.01開始.全体農/漁會上年度決算淨值扣除992810投資全國農業金庫尚未攤提之損失
              //104.02.24 add 取消(三)限制
              //106.01.19 原3.非會員存款之額度限制.取消顯示
              /*
              if((amt_990320-amt_992810)!=0){	
                //System.out.println("990320 not zero");
                if (((amt_990310-(amt_990630  / 2)) / (amt_990320-amt_992810)) > 10) {//102.01.16 add 990630/2
                  //flag[3] = 1;
                  //flag[3][0] = 1; //104.02.24 add 取消(三)限制
                  //flag[3][1] = ((amt_990310-(amt_990630  / 2)) / (amt_990320-amt_992810));
                }
              }
              */
              //3.贊助會員授信總額占贊助會員存款總額之比率
              //違反其對全部贊助會員授信總額占贊助會員存款總額之比率不得超過150%之規定
              //97.01.30 add 990612
              //102.01.16 add (信用部逾放比< 2%  且 BIS > 8%者 得不超過 150%)
              //102.01.16 add 信用部逾放比< 1%  且 BIS > 10%且備抵呆帳覆蓋率高100%,已申請經主管機關同意者,得不超過200%
              //104.02.24 add 信用部逾放比< 1%  且 BIS > 10%且備抵呆帳覆蓋率高100%且 > 全體信用部備抵呆帳覆蓋率平均值且備呆占放款比率> 2%,已申請經主管機關同意者,不受限制)
              //106.05.19 add 若逾期放款=0,備抵呆帳覆蓋率因分母為0,則符合不受限制及200%範圍 
              //107.04.09 fix 取消不受限制-備抵呆帳覆蓋率 > 全體信用部備抵呆帳覆蓋率平均值及不<100% 檢核條件
              //107.04.09 add 若不符合比率區間時,顯示違反
              //110.03.03 add 調整比率範圍預設為150%
              //               信用部逾放比< 2% 且 BIS > 8%者 得不超過 150%)-->取消
              //              (信用部逾放比< 1% 且 BIS > 10%且備抵呆帳覆蓋率 > 100%已申請經主管機關同意者：得不超過 200%)-->990421
              //              (信用部逾放比< 1% 且放款覆蓋率> 2%且BIS > 10%,已申請經主管機關同意者：得逾200%)-->990422
              
              if (amt_990420 != 0) {
                  if(amt_990422 > 0){//得逾200%
                      //信用部逾放比< 1% 且放款覆蓋率> 2%且BIS > 10%,已申請經主管機關同意者,得逾200%
                      System.out.println(bank_name+"-1.990422 > 0"); 
                      System.out.println(bank_name+"-1.核准上限-"+amt_990422_limit);
                      //110.03.03 add 檢核超過核定上限
                      if((amt_990422_limit != 999) && ( ((double)amt_990410)/((double)amt_990420) > (amt_990422_limit/100))){//超過核定上限
                    	  flag[3][0] = 1;
                          tmp_A=((double)amt_990410)/((double)amt_990420);
                          tmp_B =Math.round( tmp_A * 10000);
                          tmp_A=((double)tmp_B)/100;                      
                          flag[3][1] = tmp_A;         
                          System.out.println(bank_name+"-1.超過-核定上限-"+amt_990422_limit);
                      }
                      //107.04.09增加檢核條件不符合時,顯示違反
                      if(!(a01field_over_rate < 1 && a05bis > 10 && a01field_backup_credit_rate > 2 )){//條件不符合
                    	  //(信用部逾放比< 1% 且BIS > 10% 且放款覆蓋率> 2%,已申請經主管機關同意者：得逾200%)
                          flag[3][0] = 1; 
                          //flag[4][1] = (amt_990410 / amt_990420);
                          tmp_A=((double)amt_990410)/((double)amt_990420);
                          tmp_B =Math.round( tmp_A * 10000);
                          tmp_A=((double)tmp_B)/100;                      
                          flag[3][1] = tmp_A;         
                          System.out.println(bank_name+"-1.得逾200%-條件不符合");
                      }
                  }else if(amt_990421 > 0){/*200%*/
                      //信用部逾放比< 1%  且 BIS > 10%且備抵呆帳覆蓋率 > 100%已申請經主管機關同意者,得不超過 200%
                      System.out.println(bank_name+"-2.990421 > 0"); 
                      if ( ((double)amt_990410)/((double)amt_990420) > 2) {/*200%*/                              
                          flag[3][0] = 1; 
                          //flag[4][1] = (amt_990410 / amt_990420);
                          tmp_A=((double)amt_990410)/((double)amt_990420);
                          tmp_B =Math.round( tmp_A * 10000);
                          tmp_A=((double)tmp_B)/100;                      
                          flag[3][1] = tmp_A;
                          System.out.println(bank_name+"-2.200%");
                      }
                      //107.04.09增加檢核條件不符合時,顯示違反
                      if(a01field_over == 0 && a01field_backup_over_rate == 0.0 && a01field_backup > 0){//102.09.23 add 若備呆占狹義逾期放款比率(備抵呆帳/狹義逾放)=0,但備抵呆帳 > 0 且狹義逾放=0時,可>200
                          //備抵呆帳覆蓋率為0時
                          System.out.println("a01field_over == 0 && a01field_backup_over_rate == 0.0-->無限大 && a01field_backup > 0");
                          if(!(a01field_over_rate < 1 && a05bis > 10)){//條件不符合
                              flag[3][0] = 1; 
                              //flag[4][1] = (amt_990410 / amt_990420);
                              tmp_A=((double)amt_990410)/((double)amt_990420);
                              tmp_B =Math.round( tmp_A * 10000);
                              tmp_A=((double)tmp_B)/100;                      
                              flag[3][1] = tmp_A;
                              System.out.println(bank_name+"-2.200%-條件不符合(field_over=0)");
                          }
                      }else{
                          //備抵呆帳覆蓋率不為0時
                          if(!(a01field_over_rate < 1 && a05bis > 10 && a01field_backup_over_rate > 100)){//條件不符合
                             flag[3][0] = 1; 
                             //flag[4][1] = (amt_990410 / amt_990420);
                             tmp_A=((double)amt_990410)/((double)amt_990420);
                             tmp_B =Math.round( tmp_A * 10000);
                             tmp_A=((double)tmp_B)/100;                      
                             flag[3][1] = tmp_A;
                             System.out.println(bank_name+"-2.200%-條件不符合(field_over!=0)");
                          }
                      }
                  }else{
                      System.out.println(bank_name+"-3.990422=0 && 990421=0 ");
                      //信用部逾放比< 2%  且 BIS > 8%者 得不超過 150% //110.03.03取消限制
                      /*
                      if(a01field_over_rate < 2 && a05bis > 8){
                          if ( ((double)amt_990410)/((double)amt_990420) > 1.5) { //150%                             
                              flag[4][0] = 1; 
                              //flag[4][1] = (amt_990410 / amt_990420);
                              tmp_A=((double)amt_990410)/((double)amt_990420);
                              tmp_B =Math.round( tmp_A * 10000);
                              tmp_A=((double)tmp_B)/100;                      
                              flag[4][1] = tmp_A;
                              System.out.println(bank_name+"-3.150%");
                          }   
                          
                      }else */
                      if ( ((double)amt_990410)/((double)amt_990420) > 1.5) {//150%
                          //不超過 150% //110.03.03預設為150%
                          flag[3][0] = 1; 
                          //flag[4][1] = (amt_990410 / amt_990420);
                          tmp_A=((double)amt_990410)/((double)amt_990420);
                          tmp_B =Math.round( tmp_A * 10000);
                          tmp_A=((double)tmp_B)/100;                      
                          flag[3][1] = tmp_A;
                          System.out.println(bank_name+"-3.150%");
                          
                      }
                      
                  }
              }
              if(Integer.parseInt(S_YEAR+S_MONTH) >= 11005){                  
              //4.非會員授信總額占非會員存款總額之比率
              //違反非會員授信總額占非會員存款總額比率不得超過150%之規定
              //97.01.30 add 990612非會員政策性農業專案貸款 
              //102.01.16 add 逾放比低於2%且資本適足率高8%,已申請經主管機關同意者,得不超過150%
              //104.02.24 add 逾放比< 1% 且 BIS > 10% 且備抵呆帳覆蓋率高100%且 > 全體信用部備抵呆帳覆蓋率平均值且備呆占放款比率> 2%,已申請經主管機關同意者,得不超過 200%)
              //106.05.19 add 若逾期放款=0,備抵呆帳覆蓋率因分母為0,則符合200%範圍 
              //106.10.06 add 非會員授信總額.移除扣除990511非會員無擔保消費性政策貸款
              //          fix 非會員授信總額(990610) - (990611) - (990612)
              //107.04.09 fix 取消200%-備抵呆帳覆蓋率 > 全體信用部備抵呆帳覆蓋率平均值及不<100% 檢核條件
              //107.04.09 add 若不符合比率區間時,顯示違反
              //110.02.26 add 調整適用範圍,預設為150%
              //              (信用部逾放比< 2% 且 BIS > 8%已申請經主管機關同意者,得不超過 150%)-->取消限制
      		  //              (信用部逾放比< 1% 且BIS > 10% 且備抵呆帳覆蓋率> 100%,已申請經主管機關同意者,得不超過 200%)-->990622
              //              (信用部逾放比< 1% 且放款覆蓋率>2% 且BIS > 12%,已申請經主管機關同意者,得逾 200%)-->990623

              a = 0;
              a = amt_990610 - amt_990611 - amt_990612;//106.10.06 移除扣除990511
              System.out.println("990610-990611-990612="+a);
              
              if (amt_990620 != 0) {
            	  if(amt_990623 > 0){//得逾200%
                      //(信用部逾放比< 1% 且放款覆蓋率> 2%且BIS > 12%,已申請經主管機關同意者,得超過 200%)
                      System.out.println(bank_name+"-1.990623 > 0"); 
                      System.out.println(bank_name+"-1.核准上限-"+amt_990623_limit);
                      //110.03.03 add 檢核超過核定上限
                      if((amt_990623_limit != 999) && ( (a / amt_990620) > (amt_990623_limit/100))){//超過核定上限
                    	  flag[4][0] = 1;
                    	  tmp_A=((double)amt_990610 - (double)amt_990611 - (double)amt_990612)/((double)amt_990620);//106.10.06 移除扣除990511
                          tmp_B =Math.round( tmp_A * 10000);
                          tmp_A=((double)tmp_B)/100;                  
                          flag[4][1] = tmp_A;         
                          System.out.println(bank_name+"-1.超過-核定上限-"+amt_990623_limit);
                      }
                      //110.03.03檢核條件不符合時,顯示違反
                      if(!(a01field_over_rate < 1 && a05bis > 12 && a01field_backup_credit_rate > 2 )){//條件不符合                    	  
                    	  //(信用部逾放比< 1% 且BIS > 12% 且放款覆蓋率> 2%,已申請經主管機關同意者,得超過 200%)
                    	  flag[4][0] = 1;
                    	  tmp_A=((double)amt_990610 - (double)amt_990611 - (double)amt_990612)/((double)amt_990620);//106.10.06 移除扣除990511
                          tmp_B =Math.round( tmp_A * 10000);
                          tmp_A=((double)tmp_B)/100;                  
                          flag[4][1] = tmp_A;         
                          System.out.println(bank_name+"-1.得逾200%-條件不符合");
                      }
            	  }else if(amt_990622 > 0){/*200%*/
                      //信用部逾放比< 1%  且 BIS > 10%且備抵呆帳覆蓋率 > 100%已申請經主管機關同意者,得不超過 200%
                      System.out.println(bank_name+"-2.990622 > 0"); 
                      if ( (a / amt_990620) > 2) {/*200%*/                              
                    	  flag[4][0] = 1;
                    	  tmp_A=((double)amt_990610 - (double)amt_990611 - (double)amt_990612)/((double)amt_990620);//106.10.06 移除扣除990511
                          tmp_B =Math.round( tmp_A * 10000);
                          tmp_A=((double)tmp_B)/100;                  
                          flag[4][1] = tmp_A;      
                          System.out.println(bank_name+"-2.200%");
                      }
                      //107.04.09增加檢核條件不符合時,顯示違反
                      if(a01field_over == 0 && a01field_backup_over_rate == 0.0 && a01field_backup > 0){//102.09.23 add 若備呆占狹義逾期放款比率(備抵呆帳/狹義逾放)=0,但備抵呆帳 > 0 且狹義逾放=0時,可>200
                          //備抵呆帳覆蓋率為0時
                          System.out.println("a01field_over == 0 && a01field_backup_over_rate == 0.0-->無限大 && a01field_backup > 0");
                          if(!(a01field_over_rate < 1 && a05bis > 10)){//條件不符合
                        	  flag[4][0] = 1;
                        	  tmp_A=((double)amt_990610 - (double)amt_990611 - (double)amt_990612)/((double)amt_990620);//106.10.06 移除扣除990511
                              tmp_B =Math.round( tmp_A * 10000);
                              tmp_A=((double)tmp_B)/100;                  
                              flag[4][1] = tmp_A;      
                              System.out.println(bank_name+"-2.200%-條件不符合(field_over=0)");
                          }
                      }else{
                          //備抵呆帳覆蓋率不為0時
                          if(!(a01field_over_rate < 1 && a05bis > 10 && a01field_backup_over_rate > 100)){//條件不符合
                        	  flag[4][0] = 1;
                        	  tmp_A=((double)amt_990610 - (double)amt_990611 - (double)amt_990612)/((double)amt_990620);//106.10.06 移除扣除990511
                              tmp_B =Math.round( tmp_A * 10000);
                              tmp_A=((double)tmp_B)/100;                  
                              flag[4][1] = tmp_A;      
                              System.out.println(bank_name+"-2.200%-條件不符合(field_over!=0)");
                          }
                      }
                  }else{
                      System.out.println(bank_name+"-3.990623=0 && 990622=0 ");                      
                      if ( (a / amt_990620) > 1.5) {//150%
                          //不超過 150% //110.03.03預設為150%
                    	  flag[4][0] = 1;
                    	  tmp_A=((double)amt_990610 - (double)amt_990611 - (double)amt_990612)/((double)amt_990620);//106.10.06 移除扣除990511
                          tmp_B =Math.round( tmp_A * 10000);
                          tmp_A=((double)tmp_B)/100;                  
                          flag[4][1] = tmp_A;     
                          System.out.println(bank_name+"-3.150%");
                          
                      }
                      
                  }
              }
              
              //7.辦理非會員無擔保消費性貸款（1,000千元以下）之限制
              //違反非會員無擔保消費性貸款占農漁會上年度決算淨值之比率不得超過100%之規定
              // 99.01.25 fix 99.01開始.全体農/漁會上年度決算淨值扣除992810投資全國農業金庫尚未攤提之損失
              //106.10.06 add 無擔保消費性貸款調整為990510非會員無擔保消費性貸款 -990511非會員無擔保消費性政策貸款
              if((amt_990320-amt_992810)!=0){	
                //System.out.println("990320 not zero");
                if ( ((amt_990510-amt_990511) / (amt_990320-amt_992810)) > 1) {                  
                	  flag[5][0] = 1; 
                      //flag[5][1] = (amt_990512 / (amt_990320-amt_992810));
                      tmp_A=((double)amt_990510-(double)amt_990511)/((double)amt_990320-(double)amt_992810);
                      tmp_B =Math.round( tmp_A * 10000);
                      tmp_A=((double)tmp_B)/100;                      
                      flag[5][1] = tmp_A;
                }
              }
              
              }else{//110.04以前.
            	  //4.辦理非會員無擔保消費性貸款（1,000千元以下）之限制
            	  //違反非會員無擔保消費性貸款占農漁會上年度決算淨值之比率不得超過100%之規定
                  // 99.01.25 fix 99.01開始.全体農/漁會上年度決算淨值扣除992810投資全國農業金庫尚未攤提之損失
                  //106.10.06 add 無擔保消費性貸款調整為990510非會員無擔保消費性貸款 -990511非會員無擔保消費性政策貸款
                  if((amt_990320-amt_992810)!=0){	
                    //System.out.println("990320 not zero");
                    if ( ((amt_990510-amt_990511) / (amt_990320-amt_992810)) > 1) {                  
                    	  flag[4][0] = 1;
                          tmp_A=((double)amt_990510-(double)amt_990511)/((double)amt_990320-(double)amt_992810);
                          tmp_B =Math.round( tmp_A * 10000);
                          tmp_A=((double)tmp_B)/100;                      
                          flag[4][1] = tmp_A;
                    }
                  }
                  //5.非會員授信總額占非會員存款總額之比率
                  //違反非會員授信總額占非會員存款總額比率不得超過100%之規定
                  //97.01.30 add 990612非會員政策性農業專案貸款 
                  //102.01.16 add 逾放比低於2%且資本適足率高8%,已申請經主管機關同意者,得不超過150%
                  //104.02.24 add 逾放比< 1% 且 BIS > 10% 且備抵呆帳覆蓋率高100%且 > 全體信用部備抵呆帳覆蓋率平均值且備呆占放款比率> 2%,已申請經主管機關同意者,得不超過 200%)
                  //106.05.19 add 若逾期放款=0,備抵呆帳覆蓋率因分母為0,則符合200%範圍 
                  //106.10.06 add 非會員授信總額.移除扣除990511非會員無擔保消費性政策貸款
                  //          fix 非會員授信總額(990610) - (990611) - (990612)
                  //107.04.09 fix 取消200%-備抵呆帳覆蓋率 > 全體信用部備抵呆帳覆蓋率平均值及不<100% 檢核條件
                  //107.04.09 add 若不符合比率區間時,顯示違反
                  a = 0;
                  a = amt_990610 - amt_990611 - amt_990612;//106.10.06 移除扣除990511
                  System.out.println("990610-990611-990612="+a);
                  
                  if (amt_990620 != 0) {
                      if(amt_990622 > 0){//不超過 200%
                         //信用部逾放比< 1% 且放款覆蓋率> 2%且BIS > 10%,已申請經主管機關同意者,得不超過 200%
                         System.out.println(bank_name+"-amt_990622 > 0"); 
                         if( (a / amt_990620) > 2) { //104.02.24 add 990622 //逾放比< 1% 且 BIS > 10% 且備呆占放款比率> 2%,已申請經主管機關同意者,得不超過 200%
                             flag[5][0] = 1;  /*200%*/
                             //flag[6][1] = (a / amt_990620);
                             tmp_A=((double)amt_990610 - (double)amt_990611 - (double)amt_990612)/((double)amt_990620);//106.10.06 移除扣除990511
                             tmp_B =Math.round( tmp_A * 10000);
                             tmp_A=((double)tmp_B)/100;                      
                             flag[5][1] = tmp_A;
                             System.out.println(bank_name+"-1.200%");
                         }
                         //107.04.09增加檢核條件不符合時,顯示違反
                         if(!(a01field_over_rate < 1 && a05bis > 10 && a01field_backup_credit_rate > 2)){//條件不符合   
                             flag[5][0] = 1;  /*200%*/
                             //flag[6][1] = (a / amt_990620);
                             tmp_A=((double)amt_990610 - (double)amt_990611 - (double)amt_990612)/((double)amt_990620);//106.10.06 移除扣除990511
                             tmp_B =Math.round( tmp_A * 10000);
                             tmp_A=((double)tmp_B)/100;                      
                             flag[5][1] = tmp_A;
                             System.out.println(bank_name+"-1.200%-條件不符合");
                         }
                      }else if(amt_990621 > 0){//信用部逾放比< 2%  且 BIS > 8%已申請經主管機關同意者,得不超過 150%                     
                          System.out.println(bank_name+"-amt_990621 > 0");
                          //信用部逾放比< 2%  且 BIS > 8%已申請經主管機關同意者,得不超過 150%
                          if ( (a / amt_990620) > 1.5) { /*150%*/                
                              flag[5][0] = 1; 
                              //flag[6][1] = (a / amt_990620);
                              tmp_A=((double)amt_990610 - (double)amt_990611 - (double)amt_990612)/((double)amt_990620);//106.10.06 移除扣除990511
                              tmp_B =Math.round( tmp_A * 10000);
                              tmp_A=((double)tmp_B)/100;                      
                              flag[5][1] = tmp_A;
                              System.out.println(bank_name+"-2.150%");
                          }
                          //107.04.09增加檢核條件不符合時,顯示違反
                          if(!(a01field_over_rate < 2 && a05bis > 8)){//條件不符合                                   
                              flag[5][0] = 1; 
                              //flag[6][1] = (a / amt_990620);
                              tmp_A=((double)amt_990610 - (double)amt_990611 - (double)amt_990612)/((double)amt_990620);//106.10.06 移除扣除990511
                              tmp_B =Math.round( tmp_A * 10000);
                              tmp_A=((double)tmp_B)/100;                      
                              flag[5][1] = tmp_A;
                              System.out.println(bank_name+"-2.150%-條件不符合");
                          }
                      }else{//100%                     
                          System.out.println(bank_name+"-3.amt_990622=0 && amt_990621=0");
                          if ( (a / amt_990620) > 1) {/*100%*/                 
                              flag[5][0] = 1; 
                              //flag[6][1] = (a / amt_990620);
                              tmp_A=((double)amt_990610 - (double)amt_990611 - (double)amt_990612)/((double)amt_990620);//106.10.06 移除扣除990511
                              tmp_B =Math.round( tmp_A * 10000);
                              tmp_A=((double)tmp_B)/100;                      
                              flag[5][1] = tmp_A;
                              System.out.println(bank_name+"-3.100%");
                          }
                      } 
                  }
              }
              
              //8.購置住宅放款及房屋修繕放款限額 
              //違反購置住宅放款及房屋修繕放款之餘額不得超過存款總餘額55%之規定
              //108.09.09 add 108年10月以後套用新格式
              if(Integer.parseInt(S_YEAR+S_MONTH) >= 10810){            	
            	//field_990711+field_990712/a01fieldi_y=a+b/c
            	if (a01fieldi_y != 0) {                     
                      if ( ((amt_990711+amt_990712) / a01fieldi_y) > 0.55) {             
                        flag[6][0] = 1;
                        tmp_A=((double)amt_990711+(double)amt_990712)/(double)a01fieldi_y;
                        tmp_B=Math.round( tmp_A * 10000);
                        tmp_A=((double)tmp_B)/100;                      
                        flag[6][1] = tmp_A;
                      }
                 }            	
              }else{
                 //7.違反自用住宅放款總額占定期性存款總額比率不得超過40%之規定
                 //104.02.24 調整為50%
                 if (amt_990720 != 0) {
                   //System.out.println("990720 not zero");
                   if ( (amt_990710 / amt_990720) > 0.5) {//104.02.24 調整為50%                  
                     flag[6][0] = 1; 
                     //flag[7][1] = (amt_990710 / amt_990720);
                     tmp_A=((double)amt_990710)/((double)amt_990720);
                     tmp_B =Math.round( tmp_A * 10000);
                     tmp_A=((double)tmp_B)/100;                      
                     flag[6][1] = tmp_A;
                   }
                 }
              }
              //9.固定資產淨額限制
              //違反信用部固定資產淨額不得超過上年度信用部決算淨值              
              //103.01.06 add 上年度信用部決算淨值為負數時,要顯示違反 
              a = 0;
              a = amt_990230 - amt_990240 - amt_992810; //上年度信用部決算淨值（扣除檢查應補提未提足之備抵呆帳）
              //99.01.25 fix 99.01開始.上年度信用部決算淨值扣除992810投資全國農業金庫尚未攤提之損失
              if (a != 0) {
                //System.out.println("990230-990240 not zero");
                //101.08.27 add
                //違反信用部固定資產淨額不得超過上年度信用部決算淨值不在此限的原因項目：
                //990812一、因購置或汰換安全維護或營業相關設備，經中央主管機關核准
                //990813二、因固定資產重估增值
                //990814三、因淨值降低
                if ( (amt_990810 / a) > 1 && amt_990812 == 0 && amt_990813 == 0 & amt_990814 == 0) {                 
                  flag[7][0] = 1; 
                  flag[7][1] = (amt_990810 / a);                  
                }
                if(a < 0){
                   flag[7][0] = 1;//103.01.06 add 上年度信用部決算淨值為負數時,要顯示違反 
                   flag[7][1] = (amt_990810 / a);                    
                }
              }
              //10.外幣風險之限制
              //違反外幣資產與外幣負債差額絕對值不得超過上年底信用部帳面價值5%之限制
              //99.01.25 fix 99.01開始.上年度信用部決算淨值扣除992810投資全國農業金庫尚未攤提之損失
              //103.01.06 add 逾新台幣100萬且超過前一年度信用部決算淨值5%才算違反
              //110.09.03 fix 逾新台幣100萬且超過前一年度信用部決算淨值10%才算違反
              
              a = 0;
              if ((amt_990230-amt_992810) != 0) {
                a = Math.abs(amt_990910 - amt_990920);
                System.out.println("(990910-990920)/(990230-992810)"+(a / (amt_990230-amt_992810)));
                if ( (a / (amt_990230-amt_992810)) > 0.1) {//110.09.03原為5%調整為10%    
                    if(a > 1000000){//103.01.06 add 超過100萬者,才算違反
                       flag[8][0] = 1; 
                       //flag[9][1] = (a / (amt_990230-amt_992810));                       
                       tmp_A=(Math.abs((double)amt_990910 - (double)amt_990920))/((double)amt_990230-(double)amt_992810);
                       tmp_B =Math.round( tmp_A * 10000);
                       tmp_A=((double)tmp_B)/100;                      
                       flag[8][1] = tmp_A;
                    }
                }
              }
              //99.01.25 fix 94.10以後.10~13公式.漁會同農會=====================================
              //11.對負責人、各部門員工或與其負責人或辦理授信之職員有利害關係者為擔保授信限制
              if(bank_type.equals("6") || (Integer.parseInt(S_YEAR+S_MONTH) >= 9410)){	
                //11.1違反擔保放款最高限額不得超過(上一年度農會信用部決算淨值加上月底存款總額)25%之限制
                a = 0;
                b = 0; //最高限額
                a = amt_990230 - amt_990240 - amt_992810; //上年度信用部決算淨值（扣除檢查應補提未提足之備抵呆帳）
                //99.01.25 fix 99.01開始.上年度信用部決算淨值扣除992810投資全國農業金庫尚未攤提之損失
                if (6000000 >= (a * 0.25)) {
                  b = 6000000;
                }else if (9000000 >= (a * 0.25)) {
                  b = 9000000;
                }else {
                  b = (a * 0.25);
                }
                if (amt_991010 > b) {                 
                  flag[9][0] = 1; 
                  flag[9][1] =amt_991010;
                  
                }
              }else {
                //11.1違反擔保放款最高限額不得超過(上一年度漁會信用部決算淨值加上月底存款總額)2%之限制
                a = 0;
                b = 0; //最高限額
                a = amt_990230 - amt_990240; //上年度信用部決算淨值（扣除檢查應補提未提足之備抵呆帳）
                if (6000000 >= ( ( (a + amt_991030)) * 0.02)) {
                  b = 6000000;
                }else {
                  b = ( (a + amt_991030) * 0.02);
                }
                if (amt_991010 > b) {                 
                  flag[9][0] = 1; 
                  flag[9][1] = (amt_991010);
                }
              }

              //11.2違反擔保授信總額不得超過上一年度農/漁會決算淨值1.5倍之限制
              //99.01.25 fix 99.01開始.上年度全体農/漁會決算淨值扣除992810投資全國農業金庫尚未攤提之損失
              if((amt_990320-amt_992810)!=0){	
                //System.out.println("990320 not zero");
                if ( (amt_991020 / (amt_990320-amt_992810)) > 1.5) {
                  flag[10][0] = 1; 
                  flag[10][1] = (amt_991020 / (amt_990320-amt_992810));                  
                }
              }
              
              //12.對每一會員(含同戶家屬)及同一關係人放款最高限額              
              //99.01.25 fix 94.10以後.10~13公式.漁會同農會=====================================
              if(bank_type.equals("6") || (Integer.parseInt(S_YEAR+S_MONTH) >= 9410)){	
                //12.1違反放款最高限額不得超過上一年度農會信用部決算淨值25%之限制
                a = 0;
                b = 0; //最高限額
                a = amt_990230 - amt_990240 - amt_992810; //上年度信用部決算淨值（扣除檢查應補提未提足之備抵呆帳）
                //99.01.25 fix 99.01開始.上年度信用部決算淨值扣除992810投資全國農業金庫尚未攤提之損失
                if (6000000 >= (a * 0.25)) {
                  b = 6000000;
                }else if (9000000 >= (a * 0.25)) {
                  b = 9000000;
                }else {
                  b = (a * 0.25);
                }
                if (amt_991110 > b) {                 
                  flag[11][0] = 1; 
                  flag[11][1] = (amt_991110);
                }
                //12.2違反無擔保放款最高限額不得超過上一年度農會信用部決算淨值5%之限制
                a = 0;
                b = 0; //最高限額
                a = amt_990230 - amt_990240 - amt_992810; //上年度信用部決算淨值（扣除檢查應補提未提足之備抵呆帳）
                //99.01.25 fix 99.01開始.上年度信用部決算淨值扣除992810投資全國農業金庫尚未攤提之損失
                if (2000000 >= (a * 0.05)) {
                  b = 2000000;
                }else {
                  b = (a * 0.05);
                }
                if (amt_991120 > b) {                 
                  flag[12][0] = 1; 
                  flag[12][1] = (amt_991120);
                }
              }else {//漁會.12.對每一會員(含同戶家屬)及同一關係人放款最高限額
                //12.1違反放款最高限額不得超過上一年度漁會信用部決算淨值2%之限制
                a = 0;
                b = 0; //最高限額
                a = amt_990230 - amt_990240; //上年度信用部決算淨值（扣除檢查應補提未提足之備抵呆帳）
                if (6000000 >= (a + amt_991030) * 0.02) {
                  b = 6000000;
                }else {
                  b = (a * amt_991030) * 0.02;
                }
                if (amt_991110 > b) {                 
                  flag[11][0] = 1; 
                  flag[11][1] = (amt_991110);
                }
                //12.2違反無擔保放款最高限額不得超過上一年度漁會信用部決算淨值1%之限制
                a = 0;
                b = 0; //最高限額
                a = amt_990230 - amt_990240; //上年度信用部決算淨值（扣除檢查應補提未提足之備抵呆帳）
                if (2000000 >= (a + amt_991030) * 0.01) {
                  b = 2000000;
                }else {
                  b = (a + amt_991030) * 0.01;
                }
                if (amt_991120 > b) {                 
                  flag[12][0] = 1; 
                  flag[12][1] = (amt_991120);
                }
              }
              //13.對每一贊助會員及同一關係人之授信限額
              //99.01.25 fix 94.10以後.10~13公式.漁會同農會=====================================
              if(bank_type.equals("6") || (Integer.parseInt(S_YEAR+S_MONTH) >= 9410)){              
                //13.1違反放款最高限額不得超過上一年度農會信用部決算淨值25%之限制
                a = 0;
                b = 0; //最高限額
                a = amt_990230 - amt_990240 - amt_992810; //上年度信用部決算淨值（扣除檢查應補提未提足之備抵呆帳）
                //99.01.25 fix 99.01開始.上年度信用部決算淨值扣除992810投資全國農業金庫尚未攤提之損失
                if (6000000 >= (a * 0.25)) {
                  b = 6000000;
                }else if (9000000 >= (a * 0.25)) {
                  b = 9000000;
                }else {
                  b = (a * 0.25);
                }
                if (amt_991210 > b) {                 
                  flag[13][0] = 1; 
                  flag[13][1] = (amt_991210);
                }
                //13.2違反無擔保放款最高限額不得超過上一年度農會信用部決算淨值5%之限制
                a = 0;
                b = 0; //最高限額
                a = amt_990230 - amt_990240 - amt_992810; //上年度信用部決算淨值（扣除檢查應補提未提足之備抵呆帳）
                //99.01.25 fix 99.01開始.上年度信用部決算淨值扣除992810投資全國農業金庫尚未攤提之損失
                if (2000000 >= (a * 0.05)) {
                  b = 2000000;
                }else {
                  b = (a * 0.05);
                }
                if (amt_991220 > b) {                 
                  flag[14][0] = 1; 
                  flag[14][1] = (amt_991220);
                }
              }else {//漁會.13.對每一贊助會員及同一關係人之授信限額
                //13.1違反放款最高限額不得超過上一年度漁會信用部決算淨值2%之限制
                a = 0;
                b = 0; //最高限額
                a = amt_990230 - amt_990240; //上年度信用部決算淨值（扣除檢查應補提未提足之備抵呆帳）
                if (6000000 >= ( (a + amt_991030) * 0.02)) {
                  b = 6000000;
                }else {
                  b = (a + amt_991030) * 0.02;
                }
                if (amt_991210 > b) {                 
                  flag[13][0] = 1; 
                  flag[13][1] = (amt_991210);
                }
                //13.2違反無擔保放款最高限額不得超過上一年度漁會信用部決算淨值1%之限制
                a = 0;
                b = 0; //最高限額
                a = amt_990230 - amt_990240; //上年度信用部決算淨值（扣除檢查應補提未提足之備抵呆帳）
                if (2000000 >= ( (a + amt_991030) * 0.01)) {
                  b = 2000000;
                }else {
                  b = (a + amt_991030) * 0.01;
                }
                if (amt_991220 > b) {                 
                  flag[14][0] = 1; 
                  flag[14][1] = (amt_991220);
                }
              }
              //14.對同一非會員及同一關係人之授信限額
              //99.01.25 fix 94.10以後.10~13公式.漁會同農會=====================================              
              if(bank_type.equals("6") || (Integer.parseInt(S_YEAR+S_MONTH) >= 9410)){   
                //14.1違反放款最高限額不得超過上一年度農會信用部決算淨值12.5%之限制
                a = 0;
                b = 0; //最高限額
                a = amt_990230 - amt_990240 - amt_992810; //上年度信用部決算淨值（扣除檢查應補提未提足之備抵呆帳）
                //99.01.25 fix 99.01開始.上年度信用部決算淨值扣除992810投資全國農業金庫尚未攤提之損失
                if (6000000 >= (a * 0.125)) {
                  b = 6000000;
                }else if (9000000 >= (a * 0.125)) {
                  b = 9000000;
                }else {
                  b = (a * 0.125);
                }
                if (amt_991310 > b) {                 
                  flag[15][0] = 1; 
                  flag[15][1] = (amt_991310);
                }

                //14.2違反無擔保放款最高限額不得超過上一年度農會信用部決算淨值2.5%之限制
                a = 0;
                b = 0; //最高限額
                a = amt_990230 - amt_990240 - amt_992810; //上年度信用部決算淨值（扣除檢查應補提未提足之備抵呆帳）
                //99.01.25 fix 99.01開始.上年度信用部決算淨值扣除992810投資全國農業金庫尚未攤提之損失
                if (2000000 >= (a * 0.025)) {
                  b = 2000000;
                }else {
                  b = (a * 0.025);
                }
                if (amt_991320 > b) {                 
                  flag[16][0] = 1; 
                  flag[16][1] = (amt_991320);
                }
              }else {//漁會.14.對同一非會員及同一關係人之授信限額
                //14.1違反放款最高限額不得超過上一年度漁會信用部決算淨值1%之限制
                a = 0;
                b = 0; //最高限額
                a = amt_990230 - amt_990240; //上年度信用部決算淨值（扣除檢查應補提未提足之備抵呆帳）
                if (6000000 >= ( (a + amt_991030) * 0.01)) {
                  b = 6000000;
                }else {
                  b = (a + amt_991030) * 0.01;
                }
                if (amt_991310 > b) {                 
                  flag[15][0] = 1; 
                  flag[15][1] = (amt_991310);
                }

                //14.2違反無擔保放款最高限額不得超過上一年度農會信用部決算淨值0.5%之限制
                a = 0;
                b = 0; //最高限額
                a = amt_990230 - amt_990240; //上年度信用部決算淨值（扣除檢查應補提未提足之備抵呆帳）
                if (2000000 >= ( (a + amt_991030) * 0.005)) {
                  b = 2000000;
                }else {
                  b = (a + amt_991030) * 0.005;
                }
                if (amt_991320 > b) {                 
                  flag[16][0] = 1; 
                  flag[16][1] = (amt_991320);
                }
              }
              //104.02.24 add
              //15.對鄉(鎮、市)公所授信未經其所隸屬之縣政府保證之限額-違反對鄉(鎮、市)公所授信未經其所隸屬之縣政府保證,及對直轄市、縣(市)政府投資經營之公營事業,其授信經該直轄市、縣(市)政府保證,兩者合計不得超過上一年度信用部決算淨值       
              //field_996114_996115/(990230-990240)=(i+j)/D
              a = 0;
              b = 0; 
              a = amt_996114 + amt_996115;
              b = amt_990230 - amt_990240 - amt_992810; //上年度信用部決算淨值（扣除檢查應補提未提足之備抵呆帳）
              tmp_A=((double)amt_996114+(double)amt_996115)/((double)amt_990230-(double)amt_990240-(double)amt_992810);
              tmp_B =Math.round( tmp_A * 10000);
              tmp_A=((double)tmp_B)/100;
              //System.out.println("14.field_996114_996115/(990230-990240)="+tmp_A);
              
              if ( tmp_A > 100) {
                  
                  flag[17][0] = 1; 
                  flag[17][1] = (tmp_A);
              }
                
              //設定cell:違反事項符號
              int count=0;
              for (j = 0; j <= 17; j++) {
                if(flag[j][0]==1){
                  count++;
                }
              }
              System.out.println("count = "+count);
              if(count>0){
                 countAllItem++;
              }
              if (count > 0) {
                row = (sheet.getRow( (countAllItem + 4)) == null) ? sheet.createRow( (countAllItem + 4)) :
                sheet.getRow( (countAllItem + 4));
                cell = row.getCell( (short) 0) == null ? row.createCell( (short) 0) :
                row.getCell( (short) 0);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
                cell.setCellStyle(leftStyle_unitName);
                cell.setCellValue(bank_name);

                for (j = 0; j < 18; j++) {                   
                  //System.out.println("flag[" + j + "]=" + flag[j]);
                  row = (sheet.getRow( (countAllItem + 4)) == null) ? sheet.createRow( (countAllItem + 4)) :
                  sheet.getRow( (countAllItem + 4));                  
                  cell = (row.getCell( (short) (j + 1)) == null) ? row.createCell( (short) (j + 1)) : row.getCell( (short) (j + 1));
                  cell.setCellStyle(centerStyle_value);
                  cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                  if (flag[j][0] == 1) {
                      //System.out.println("(double)flag[j][1]="+(double)flag[j][1]);
    				  if(j==7 || j==10){//j==7-->9.固定資產淨額限制 j==10-->11.2違反擔保授信總額不得超過上一年度農會決算淨值1.5倍之限制
    				  	cell.setCellValue(flag[j][1]+"倍");
    				  	if(j==7)
    				  	{
    				  		double temp = (double)flag[j][1]*100;
    				  		temp = Math.round(temp);
    				  		temp = temp/100;
    				  		cell.setCellValue(df_md.format(temp).toString()+"倍");
    				  	}
    				  }
    				  
    				  //103.01.13 add 顯示至小數點2位,不足者補0 
    				  if(j==0 || j==1 || j==2 || j==3 || j==4 || j==5 || j==6 || j==8|| j==17){    				      
    				    //double temp = (double)flag[j][1]*100;
    				    //temp = Math.round(temp);
    				    cell.setCellValue(df_md.format(flag[j][1]).toString()+"%");
    				  }
    				  if(j==9 || j==11 || j==12 || j==13 || j==14 || j==15 || j==16 ){
                    	 cell.setCellValue(flag[j][1]);
                      }    				  
                    //System.out.println("set cell flag value = "+cell.getStringCellValue());
                  }
                  else {
                    //System.out.println("flag[" + j + "]=0");
                    cell.setCellValue("");
                    //System.out.println("set cell flag value = "+cell.getStringCellValue());
                  }
                }
              }
            }
            System.out.println("countAllItem = "+countAllItem);
            if(countAllItem==0){
               row = (sheet.getRow( (i + 2)) == null) ? sheet.createRow( (i + 2)) :
               sheet.getRow( (i + 2));
               cell = (row.getCell( (short) (0)) == null) ?
               row.createCell( (short) (0)) :
               row.getCell( (short) (0));
               cell.setCellStyle(centerStyle_value);
               cell.setEncoding(HSSFCell.ENCODING_UTF_16);
               cell.setCellValue("本頁無相關資料");
            }
          }
          File old_Report = new File(reportDir +System.getProperty("file.separator") +"信用部違反法定比率規定分析表-"+(bank_type.equals("6")?"農會":bank_type.equals("7")?"漁會":"農漁會")+".xls");
          if (old_Report.exists()) {
            old_Report.delete();
          }
          FileOutputStream fout = null;
          if (bank_type.equals("6")) {
            fout = new FileOutputStream(reportDir +System.getProperty("file.separator") +"信用部違反法定比率規定分析表-農會.xls");
          }else if (bank_type.equals("7")) {
            fout = new FileOutputStream(reportDir +System.getProperty("file.separator") +"信用部違反法定比率規定分析表-漁會.xls");
          }else if (bank_type.equals("ALL")) {
            fout = new FileOutputStream(reportDir +System.getProperty("file.separator") +"信用部違反法定比率規定分析表-農漁會.xls");           
          }
          wb.write(fout);
          //儲存
          fout.close();
        }catch (Exception e) {
          System.out.println("createRpt Error:" + e + e.getMessage());
        }
        System.out.println("FR006W_Excel End ...");
        return errMsg;
      }
}
