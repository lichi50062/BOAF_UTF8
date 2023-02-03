/*
 * 105.11.09 add 查核案件數統計表(農漁會別) by 2295
 */
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.Region;

import java.io.*;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptDS068W {	 
	 public static String createRpt(String rptKind,String S_YEAR,String S_MONTH,String S_DAY,String E_YEAR,String E_MONTH,String E_DAY,String begDate,String endDate,String selectitem,String hasBankListALL) {    

	    String errMsg = "";	    
	    List dbData = null;	    
	    int rowNum=0;
	    int cellNum=0;
	    DataObject bean = null;
	    reportUtil reportUtil = new reportUtil();
		HSSFCellStyle cs_right = null; 
		HSSFCellStyle cs_center = null;
	    HSSFCellStyle cs_left = null;
	    HSSFCellStyle cs_noborderleft = null;	      
	   
	    try {
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
	      
	      System.out.println("rptKind="+rptKind);
	      System.out.println("begDate="+begDate);
	      System.out.println("endDate="+endDate);
	      System.out.println("selectitem="+selectitem);
	      System.out.println("hasBankListALL="+hasBankListALL);
	      
	      FileInputStream finput = null;
	      String fileName = "";
	      if(rptKind.equals("1")){//統計分類:貸款種類
              fileName = "查核案件數統計表_貸款種類別.xls";
          }else if(rptKind.equals("2")){//統計分類:缺失態樣
              fileName = "查核案件數統計表_缺失態樣別.xls";
          }else if(rptKind.equals("3")){  
              if(hasBankListALL.equals("true")){//受檢單位選全部時
                 fileName = "查核案件數統計表_縣市別.xls";                                
              }else{  
                 fileName = "查核案件數統計表_農漁會別.xls";                                 
              } 
          }   
	      
	      finput = new FileInputStream(xlsDir +System.getProperty("file.separator") +fileName);      
	      
	      //設定FileINputStream讀取Excel檔
          POIFSFileSystem fs = new POIFSFileSystem(finput);
          HSSFWorkbook wb = new HSSFWorkbook(fs);
          HSSFSheet sheet = wb.getSheetAt(0);//讀取第一個工作表，宣告其為sheet
          HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
         
               
	      //設定頁面符合列印大小
	      sheet.setAutobreaks(false);
	    
	      ps.setScale( (short) (rptKind.equals("2")?80:100)); //列印縮放百分比
	      
	      ps.setPaperSize( (short) 9); //設定紙張大小 A4
	      
	      finput.close();

	      HSSFRow row = null; //宣告一列
	      HSSFCell cell = null; //宣告一個儲存格	      
	      cs_right = reportUtil.getRightStyle(wb);
	      cs_center = reportUtil.getDefaultStyle(wb);
	      cs_left = reportUtil.getLeftStyle(wb);	      
	      cs_noborderleft = reportUtil.getNoBorderLeftStyle(wb); 
	      
	      //轉換查核季別
          //第1季=0101~0331; 第2季=0401~0631; 第3季=0701~0931; 年第4季=1001~1231
          String beg_season = "",beg_season_tmp,end_season = "",end_season_tmp="";
          if(!begDate.equals("") && !endDate.equals("")){
              beg_season_tmp = begDate.substring(4,begDate.length());
              end_season_tmp = endDate.substring(4,endDate.length());
              System.out.println("beg_season_tmp="+beg_season_tmp+";end_season_tmp="+end_season_tmp);
              if(Integer.parseInt(beg_season_tmp) >= Integer.parseInt("0101") && Integer.parseInt(beg_season_tmp) <= Integer.parseInt("0331")){
                  beg_season="01";
              }else if(Integer.parseInt(beg_season_tmp) >= Integer.parseInt("0401") && Integer.parseInt(beg_season_tmp) <= Integer.parseInt("0631")){
                  beg_season="02";
              }else if(Integer.parseInt(beg_season_tmp) >= Integer.parseInt("0701") && Integer.parseInt(beg_season_tmp) <= Integer.parseInt("0931")){
                  beg_season="03";
              }else if(Integer.parseInt(beg_season_tmp) >= Integer.parseInt("1001") && Integer.parseInt(beg_season_tmp) <= Integer.parseInt("1231")){
                  beg_season="04";
              }   
              if(Integer.parseInt(end_season_tmp) >= Integer.parseInt("0101") && Integer.parseInt(end_season_tmp) <= Integer.parseInt("0331")){
                  end_season="01";
              }else if(Integer.parseInt(end_season_tmp) >= Integer.parseInt("0401") && Integer.parseInt(end_season_tmp) <= Integer.parseInt("0631")){
                  end_season="02";
              }else if(Integer.parseInt(end_season_tmp) >= Integer.parseInt("0701") && Integer.parseInt(end_season_tmp) <= Integer.parseInt("0931")){
                  end_season="03";
              }else if(Integer.parseInt(end_season_tmp) >= Integer.parseInt("1001") && Integer.parseInt(end_season_tmp) <= Integer.parseInt("1231")){
                  end_season="04";
              } 
              beg_season = S_YEAR + beg_season;
              end_season = E_YEAR + end_season;
              System.out.println("beg_season="+beg_season+";end_season="+end_season);
          }
	      
          
	      StringBuffer sql = new StringBuffer();
	      List paramList = new ArrayList();//傳入參數
	      if(rptKind.equals("1")){//統計分類:貸款種類
	         //依所選取貸款種類統計.報表查詢SQL
	         sql.append(" select loan_item,loan_item_name,");//--貸款種類
	         sql.append("        sum(bankcount_feb) as bankcount_feb,");//--金管會檢查報告.農漁會家數
	         sql.append("        sum(ex_type_feb) as ex_type_feb,");//--金管會檢查報告.案件數
	         sql.append("        sum(bankcount_agri) as bankcount_agri,");//--農業金庫查核.農漁會家數
	         sql.append("        sum(ex_type_agri) as ex_type_agri,");//--農業金庫查核.案件數
	         sql.append("        sum(bankcount_boaf) as bankcount_boaf,");//--農金局訪查.農漁會家數
	         sql.append("        sum(ex_type_boaf) as ex_type_boaf,");//--農金局訪查.案件數
	         sql.append("        sum(bankcount_sum) as bankcount_sum,");//--合計.農漁會家數
	         sql.append("        sum(ex_type_sum) as ex_type_sum ");//--案件數合計
	         sql.append(" from(");
	         sql.append(" select loan_item,loan_item_name,bank_no,");
	         sql.append("        decode(sum(ex_type_feb),0,0,null,0,1) as bankcount_feb,");
	         sql.append("        decode(sum(ex_type_feb),null,0,sum(ex_type_feb)) as ex_type_feb,");
	         sql.append("        decode(sum(ex_type_agri),0,0,null,0,1) as bankcount_agri,");
	         sql.append("        decode(sum(ex_type_agri),null,0,sum(ex_type_agri))  as ex_type_agri,");
	         sql.append("        decode(sum(ex_type_boaf),0,0,null,0,1) as bankcount_boaf,");
	         sql.append("        decode(sum(ex_type_boaf),null,0,sum(ex_type_boaf)) as ex_type_boaf,");
	         sql.append("        decode(sum(ex_type_feb)+sum(ex_type_agri)+sum(ex_type_boaf),0,0,null,0,1) as bankcount_sum,");
	         sql.append("        decode(sum(ex_type_feb)+sum(ex_type_agri)+sum(ex_type_boaf),null,0,sum(ex_type_feb)+sum(ex_type_agri)+sum(ex_type_boaf)) as ex_type_sum");
	         sql.append(" from (");
	         sql.append("       select frm_loan_item.loan_item,loan_item_name,frm_exmaster.bank_no,frm_exmaster.ex_no,");
	         sql.append("              decode(ex_type,'FEB',1,0) as ex_type_feb,");
	         sql.append("              decode(ex_type,'AGRI',1,0) as ex_type_agri,");
	         sql.append("              decode(ex_type,'BOAF',1,0) as ex_type_boaf");
	         sql.append("       from frm_exdef");
	         sql.append("       left join frm_snrtdoc on frm_exdef.ex_no = frm_snrtdoc.ex_no and frm_exdef.bank_no=frm_snrtdoc.bank_no and frm_exdef.def_seq=frm_snrtdoc.def_seq");
	         sql.append("       left join frm_exmaster  on frm_exdef.ex_no = frm_exmaster.ex_no and frm_exmaster.bank_no=frm_snrtdoc.bank_no");
	         sql.append("       left join frm_loan_item on frm_exdef.loan_item = frm_loan_item.loan_item");
	         sql.append("       where frm_exdef.loan_item in ("+selectitem+")");//--所挑選的貸款種類別
	         sql.append("         and ((ex_type='FEB' and TO_CHAR(doc_date ,'yyyymmdd') BETWEEN '"+begDate+"' AND '"+endDate+"')");
	         sql.append("         or (ex_type='AGRI' and frm_exmaster.ex_no >= '"+beg_season+"' and frm_exmaster.ex_no <= '"+end_season+"')");//105年第1季=1050101~1050331; 105年第2季=1050401~1050631所挑選的受檢單位代碼; 105年第3季=1050701~1050931; 105年第4季=1051001~1051231
	         sql.append("         or (ex_type='BOAF' and frm_exmaster.ex_no >= '"+begDate+"' and frm_exmaster.ex_no <= '"+endDate+"'))");	                
	         sql.append("       group by frm_loan_item.loan_item,loan_item_name,frm_exmaster.bank_no,frm_exmaster.ex_no,ex_type");
	         sql.append("       )group by  loan_item , loan_item_name, bank_no");
	         sql.append(" )group by  loan_item , loan_item_name");
	         sql.append(" union");
	         sql.append(" select '99' as loan_item,'合計' as loan_item_name,");
	         sql.append("        sum(bankcount_feb) as bankcount_feb,");//--農漁會家數
	         sql.append("        sum(ex_type_feb) as ex_type_feb,");//--金管會檢查報告案件數
	         sql.append("        sum(bankcount_agri) as bankcount_agri,");//--農漁會家數
	         sql.append("        sum(ex_type_agri) as ex_type_agri,");//--農業金庫查核案件數
	         sql.append("        sum(bankcount_boaf) as bankcount_boaf,");//--農漁會家數
	         sql.append("        sum(ex_type_boaf) as ex_type_boaf,");//--農金局查核案件數
	         sql.append("        sum(bankcount_sum) as bankcount_sum,");//--農漁會家數.
	         sql.append("        sum(ex_type_sum) as ex_type_sum ");//--案件數合計
	         sql.append(" from(");
	         sql.append(" select loan_item,loan_item_name,bank_no,");
	         sql.append("        decode(sum(ex_type_feb),0,0,null,0,1) as bankcount_feb,");
	         sql.append("        decode(sum(ex_type_feb),null,0,sum(ex_type_feb)) as ex_type_feb,");
	         sql.append("        decode(sum(ex_type_agri),0,0,null,0,1) as bankcount_agri,");
	         sql.append("        decode(sum(ex_type_agri),null,0,sum(ex_type_agri))  as ex_type_agri,");
	         sql.append("        decode(sum(ex_type_boaf),0,0,null,0,1) as bankcount_boaf,");
	         sql.append("        decode(sum(ex_type_boaf),null,0,sum(ex_type_boaf)) as ex_type_boaf,");
	         sql.append("        decode(sum(ex_type_feb)+sum(ex_type_agri)+sum(ex_type_boaf),0,0,null,0,1) as bankcount_sum,");
	         sql.append("        decode(sum(ex_type_feb)+sum(ex_type_agri)+sum(ex_type_boaf),null,0,sum(ex_type_feb)+sum(ex_type_agri)+sum(ex_type_boaf)) as ex_type_sum");
	         sql.append(" from (");
	         sql.append("       select frm_loan_item.loan_item,loan_item_name,frm_exmaster.bank_no,frm_exmaster.ex_no,");
	         sql.append("              decode(ex_type,'FEB',1,0) as ex_type_feb,");
	         sql.append("              decode(ex_type,'AGRI',1,0) as ex_type_agri,");
	         sql.append("              decode(ex_type,'BOAF',1,0) as ex_type_boaf");
	         sql.append("       from frm_exdef");
	         sql.append("       left join frm_snrtdoc on frm_exdef.ex_no = frm_snrtdoc.ex_no and frm_exdef.bank_no=frm_snrtdoc.bank_no and frm_exdef.def_seq=frm_snrtdoc.def_seq");
	         sql.append("       left join frm_exmaster  on frm_exdef.ex_no = frm_exmaster.ex_no and frm_exmaster.bank_no=frm_snrtdoc.bank_no");
	         sql.append("       left join frm_loan_item on frm_exdef.loan_item = frm_loan_item.loan_item");
	         sql.append("       where frm_exdef.loan_item in ("+selectitem+") ");//--所挑選的貸款種類別
	         sql.append("         and ((ex_type='FEB' and TO_CHAR(doc_date ,'yyyymmdd') BETWEEN '"+begDate+"' AND '"+endDate+"')");
             sql.append("         or (ex_type='AGRI' and frm_exmaster.ex_no >= '"+beg_season+"' and frm_exmaster.ex_no <= '"+end_season+"')");//105年第1季=1050101~1050331; 105年第2季=1050401~1050631所挑選的受檢單位代碼; 105年第3季=1050701~1050931; 105年第4季=1051001~1051231
             sql.append("         or (ex_type='BOAF' and frm_exmaster.ex_no >= '"+begDate+"' and frm_exmaster.ex_no <= '"+endDate+"'))");
	         sql.append("       group by frm_loan_item.loan_item,loan_item_name,frm_exmaster.bank_no,frm_exmaster.ex_no,ex_type");
	         sql.append("      )group by  loan_item , loan_item_name, bank_no");
	         sql.append(" )");
	         sql.append(" order by loan_item");
	         dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "bankcount_feb,ex_type_feb,bankcount_agri,ex_type_agri,bankcount_boaf,ex_type_boaf,bankcount_sum,ex_type_sum");
          }else if(rptKind.equals("2")){//統計分類:缺失態樣
             //依所選取缺失態樣統計.報表查詢SQL      
             sql.append(" select * ");
             sql.append(" from ( ");
             sql.append(" select def_type , def_name,");// --缺失態樣
             sql.append("        def_case,");
             sql.append("        def_type||'-'||def_case as def_case_no,");// --缺失情節代碼
             sql.append("        case_name,");//--缺失情節名稱
             sql.append("        sum(bankcount_feb) as bankcount_feb,");//--金管會檢查報告.農漁會家數
             sql.append("        sum(ex_type_feb) as ex_type_feb,");//--金管會檢查報告.案件數
             sql.append("        sum(bankcount_agri) as bankcount_agri,");//--農業金庫查核.農漁會家數
             sql.append("        sum(ex_type_agri) as ex_type_agri,");//--農業金庫查核.案件數
             sql.append("        sum(bankcount_boaf) as bankcount_boaf,");//--農金局訪查.農漁會家數
             sql.append("        sum(ex_type_boaf) as ex_type_boaf,");//--農金局訪查.案件數
             sql.append("        sum(bankcount_sum) as bankcount_sum,");//--合計.農漁會家數
             sql.append("        sum(ex_type_sum) as ex_type_sum ");//--案件數合計
             sql.append(" from(");
             sql.append(" select def_type , def_name,def_case,case_name,bank_no,");
             sql.append("        decode(sum(ex_type_feb),0,0,null,0,1) as bankcount_feb,");
             sql.append("        decode(sum(ex_type_feb),null,0,sum(ex_type_feb)) as ex_type_feb,");
             sql.append("        decode(sum(ex_type_agri),0,0,null,0,1) as bankcount_agri,");
             sql.append("        decode(sum(ex_type_agri),null,0,sum(ex_type_agri))  as ex_type_agri,");
             sql.append("        decode(sum(ex_type_boaf),0,0,null,0,1) as bankcount_boaf,");
             sql.append("        decode(sum(ex_type_boaf),null,0,sum(ex_type_boaf)) as ex_type_boaf,");
             sql.append("        decode(sum(ex_type_feb)+sum(ex_type_agri)+sum(ex_type_boaf),0,0,null,0,1) as bankcount_sum,");
             sql.append("        decode(sum(ex_type_feb)+sum(ex_type_agri)+sum(ex_type_boaf),null,0,sum(ex_type_feb)+sum(ex_type_agri)+sum(ex_type_boaf)) as ex_type_sum");
             sql.append(" from (");
             sql.append("       select frm_exdef.def_type,cmuse_name as def_name,frm_exdef.def_case,case_name,");
             sql.append("              frm_exdef.bank_no,frm_exdef.ex_no,frm_exdef.def_seq,");//--docno,(同一個缺失不同文號,算一個案件)
             sql.append("              decode(ex_type,'FEB',1,0) as ex_type_feb,");
             sql.append("              decode(ex_type,'AGRI',1,0) as ex_type_agri,");
             sql.append("              decode(ex_type,'BOAF',1,0) as ex_type_boaf");
             sql.append("       from (select frm_exdef.*,ex_type from frm_exdef  left join frm_exmaster  on frm_exdef.ex_no = frm_exmaster.ex_no and frm_exdef.bank_no=frm_exmaster.bank_no)frm_exdef");
             sql.append("       left join frm_snrtdoc on frm_exdef.ex_no = frm_snrtdoc.ex_no and frm_exdef.bank_no=frm_snrtdoc.bank_no and frm_exdef.def_seq=frm_snrtdoc.def_seq");
             sql.append("       left join frm_def_item on frm_exdef.def_type = frm_def_item.def_type and frm_exdef.def_case = frm_def_item.def_case");
             sql.append("       left join (select * from cdshareno where cmuse_div='047')cdshareno on frm_exdef.def_type=cdshareno.cmuse_id");
             sql.append("       where frm_exdef.def_type in ("+selectitem+") ");//--所挑選的缺失態樣別
             sql.append("         and ((ex_type='FEB' and TO_CHAR(doc_date ,'yyyymmdd') BETWEEN '"+begDate+"' AND '"+endDate+"')");
             sql.append("         or (ex_type='AGRI' and frm_exdef.ex_no >= '"+beg_season+"' and frm_exdef.ex_no <= '"+end_season+"')");//105年第1季=1050101~1050331; 105年第2季=1050401~1050631所挑選的受檢單位代碼; 105年第3季=1050701~1050931; 105年第4季=1051001~1051231
             sql.append("         or (ex_type='BOAF' and frm_exdef.ex_no >= '"+begDate+"' and frm_exdef.ex_no <= '"+endDate+"'))");
             sql.append("       group by frm_exdef.def_type,cmuse_name,frm_exdef.def_case,case_name,frm_exdef.bank_no,frm_exdef.ex_no,frm_exdef.def_seq,ex_type");
             sql.append("      order by ex_no,frm_exdef.def_seq");
             sql.append("       )group by  def_type , def_name,def_case,case_name, bank_no");
             sql.append(" )group by def_type , def_name,def_case,case_name");
             sql.append(" union");
             sql.append(" select '9' as def_type , ' ' as def_name, ");//--缺失態樣
             sql.append("         '','9' as def_case, '合計' as case_name,");//--缺失情節
             sql.append("        sum(bankcount_feb) as bankcount_feb,");//--農漁會家數
             sql.append("        sum(ex_type_feb) as ex_type_feb,");//--金管會檢查報告案件數
             sql.append("        sum(bankcount_agri) as bankcount_agri,");//--農漁會家數
             sql.append("        sum(ex_type_agri) as ex_type_agri,");//--農業金庫查核案件數
             sql.append("        sum(bankcount_boaf) as bankcount_boaf,");//--農漁會家數
             sql.append("        sum(ex_type_boaf) as ex_type_boaf,");//--農金局查核案件數
             sql.append("        sum(bankcount_sum) as bankcount_sum,");//--農漁會家數.
             sql.append("        sum(ex_type_sum) as ex_type_sum ");//--案件數合計
             sql.append(" from(");
             sql.append(" select def_type , def_name,def_case,case_name,bank_no,");
             sql.append("        decode(sum(ex_type_feb),0,0,null,0,1) as bankcount_feb,");
             sql.append("        decode(sum(ex_type_feb),null,0,sum(ex_type_feb)) as ex_type_feb,");
             sql.append("        decode(sum(ex_type_agri),0,0,null,0,1) as bankcount_agri,");
             sql.append("        decode(sum(ex_type_agri),null,0,sum(ex_type_agri))  as ex_type_agri,");
             sql.append("        decode(sum(ex_type_boaf),0,0,null,0,1) as bankcount_boaf,");
             sql.append("        decode(sum(ex_type_boaf),null,0,sum(ex_type_boaf)) as ex_type_boaf,");
             sql.append("        decode(sum(ex_type_feb)+sum(ex_type_agri)+sum(ex_type_boaf),0,0,null,0,1) as bankcount_sum,");
             sql.append("        decode(sum(ex_type_feb)+sum(ex_type_agri)+sum(ex_type_boaf),null,0,sum(ex_type_feb)+sum(ex_type_agri)+sum(ex_type_boaf)) as ex_type_sum");
             sql.append(" from (");
             sql.append("        select frm_exdef.def_type,cmuse_name as def_name,frm_exdef.def_case,case_name,");
             sql.append("              frm_exdef.bank_no,frm_exdef.ex_no,frm_exdef.def_seq,");//--docno, (同一個缺失不同文號,算一個案件)
             sql.append("              decode(ex_type,'FEB',1,0) as ex_type_feb,");
             sql.append("              decode(ex_type,'AGRI',1,0) as ex_type_agri,");
             sql.append("              decode(ex_type,'BOAF',1,0) as ex_type_boaf");
             sql.append("       from (select frm_exdef.*,ex_type from frm_exdef  left join frm_exmaster  on frm_exdef.ex_no = frm_exmaster.ex_no and frm_exdef.bank_no=frm_exmaster.bank_no)frm_exdef");
             sql.append("       left join frm_snrtdoc on frm_exdef.ex_no = frm_snrtdoc.ex_no and frm_exdef.bank_no=frm_snrtdoc.bank_no and frm_exdef.def_seq=frm_snrtdoc.def_seq");
             sql.append("       left join frm_def_item on frm_exdef.def_type = frm_def_item.def_type and frm_exdef.def_case = frm_def_item.def_case");
             sql.append("       left join (select * from cdshareno where cmuse_div='047')cdshareno on frm_exdef.def_type=cdshareno.cmuse_id");
             sql.append("       where frm_exdef.def_type in ("+selectitem+")");// --所挑選的缺失態樣別
             sql.append("         and ((ex_type='FEB' and TO_CHAR(doc_date ,'yyyymmdd') BETWEEN '"+begDate+"' AND '"+endDate+"')");
             sql.append("         or (ex_type='AGRI' and frm_exdef.ex_no >= '"+beg_season+"' and frm_exdef.ex_no <= '"+end_season+"')");//105年第1季=1050101~1050331; 105年第2季=1050401~1050631所挑選的受檢單位代碼; 105年第3季=1050701~1050931; 105年第4季=1051001~1051231
             sql.append("         or (ex_type='BOAF' and frm_exdef.ex_no >= '"+begDate+"' and frm_exdef.ex_no <= '"+endDate+"'))");                 
             sql.append("       group by frm_exdef.def_type,cmuse_name,frm_exdef.def_case,case_name,frm_exdef.bank_no,frm_exdef.ex_no,frm_exdef.def_seq,ex_type");
             sql.append("      order by ex_no,frm_exdef.def_seq");
             sql.append("      )group by  def_type , def_name,def_case,case_name,bank_no");
             sql.append(" )");
             sql.append(" )a order by def_type,to_number(def_case)");             
             dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "bankcount_feb,ex_type_feb,bankcount_agri,ex_type_agri,bankcount_boaf,ex_type_boaf,bankcount_sum,ex_type_sum");  
          }else if(rptKind.equals("3")){
	          if(hasBankListALL.equals("true")){//受檢單位選全部時    
	             //縣市別統計.報表查詢SQL
	             sql.append(" select hsien_id,hsien_name,FR001W_output_order,");
	             sql.append("        sum(bankcount_feb) as bankcount_feb,");//--金管會檢查報告.農漁會家數
	             sql.append("        sum(ex_type_feb) as ex_type_feb,");//--金管會檢查報告.案件數
	             sql.append("        sum(bankcount_agri) as bankcount_agri,");//--農業金庫查核.農漁會家數
	             sql.append("        sum(ex_type_agri) as ex_type_agri,");//--農業金庫查核.案件數
	             sql.append("        sum(bankcount_boaf) as bankcount_boaf,");//--農金局訪查.農漁會家數
	             sql.append("        sum(ex_type_boaf) as ex_type_boaf,");//--農金局訪查.案件數
	             sql.append("        sum(bankcount_sum) as bankcount_sum,");//--合計.農漁會家數.
	             sql.append("        sum(ex_type_sum) as ex_type_sum ");//--案件數合計      
	             sql.append(" from ( ");//--縣市小計
	             sql.append("        select nvl(cd01.hsien_id,' ')       as  hsien_id ,");               
	             sql.append("               nvl(cd01.hsien_name,'OTHER') as  hsien_name,");               
	             sql.append("               cd01.FR001W_output_order     as  FR001W_output_order,");               
	             sql.append("               bn01.bank_no ,  bn01.BANK_NAME,"); 
	             sql.append("               decode(sum(ex_type_feb),0,0,null,0,1) as bankcount_feb,");
	             sql.append("               decode(sum(ex_type_feb),null,0,sum(ex_type_feb)) as ex_type_feb,");
	             sql.append("               decode(sum(ex_type_agri),0,0,null,0,1) as bankcount_agri,");
	             sql.append("               decode(sum(ex_type_agri),null,0,sum(ex_type_agri))  as ex_type_agri,");
	             sql.append("               decode(sum(ex_type_boaf),0,0,null,0,1) as bankcount_boaf,");
	             sql.append("               decode(sum(ex_type_boaf),null,0,sum(ex_type_boaf)) as ex_type_boaf,");
	             sql.append("               decode(sum(ex_type_feb)+sum(ex_type_agri)+sum(ex_type_boaf),0,0,null,0,1) as bankcount_sum,");
	             sql.append("              decode(sum(ex_type_feb)+sum(ex_type_agri)+sum(ex_type_boaf),null,0,sum(ex_type_feb)+sum(ex_type_agri)+sum(ex_type_boaf)) as ex_type_sum"); 
	             sql.append("        from  (select * from  cd01 where cd01.hsien_id <> 'Y'  ) cd01"); 
	             sql.append("        left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = 100 and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL)");  
	             sql.append("        left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type in ('6','7') and bn01.m_year = 100 and bn_type <> '2' and wlx01.m_year=100");  
	             sql.append("        left join (select frm_exmaster.bank_no,frm_exmaster.ex_no,");
	             sql.append("                          decode(ex_type,'FEB',1,0) as ex_type_feb,");
	             sql.append("                          decode(ex_type,'AGRI',1,0) as ex_type_agri,");
	             sql.append("                          decode(ex_type,'BOAF',1,0) as ex_type_boaf");
	             sql.append("                   from frm_exmaster");
	             sql.append("                   left join frm_snrtdoc on frm_exmaster.ex_no = frm_snrtdoc.ex_no and frm_exmaster.bank_no=frm_snrtdoc.bank_no");
	             sql.append("                   where (ex_type='FEB' and TO_CHAR(doc_date ,'yyyymmdd') BETWEEN '"+begDate+"' AND '"+endDate+"')");
	             sql.append("                   or (ex_type='AGRI' and frm_exmaster.ex_no >= '"+beg_season+"' and frm_exmaster.ex_no <= '"+end_season+"')");//105年第1季=1050101~1050331; 105年第2季=1050401~1050631所挑選的受檢單位代碼; 105年第3季=1050701~1050931; 105年第4季=1051001~1051231
	             sql.append("                   or (ex_type='BOAF' and frm_exmaster.ex_no >= '"+begDate+"' and frm_exmaster.ex_no <= '"+endDate+"')");	           
	             sql.append("                   group by frm_exmaster.bank_no,frm_exmaster.ex_no,ex_type");
	             sql.append("                  ) frm_exmaster  on  bn01.bank_no = frm_exmaster.bank_no"); 
	             sql.append("        group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME");
	             sql.append("      ) frm_exmaster where  frm_exmaster.bank_no <> ' '");  
	             sql.append(" GROUP BY frm_exmaster.hsien_id ,frm_exmaster.hsien_name,frm_exmaster.FR001W_output_order"); 
	             sql.append(" union all");
	             sql.append(" select '  ','合計' as  hsien_name,'999' asFR001W_output_order,");
	             sql.append("        sum(bankcount_feb) as bankcount_feb,");//--農漁會家數
	             sql.append("        sum(ex_type_feb) as ex_type_feb,");//--金管會檢查報告案件數
	             sql.append("        sum(bankcount_agri) as bankcount_agri,");//--農漁會家數
	             sql.append("        sum(ex_type_agri) as ex_type_agri,");//--農業金庫查核案件數
	             sql.append("        sum(bankcount_boaf) as bankcount_boaf,");//--農漁會家數
	             sql.append("        sum(ex_type_boaf) as ex_type_boaf,");//--農金局查核案件數
	             sql.append("        sum(bankcount_sum) as bankcount_sum,");//--農漁會家數.
	             sql.append("        sum(ex_type_sum) as ex_type_sum ");//--案件數合計      
	             sql.append(" from ( ");//--合計
	             sql.append("        select nvl(cd01.hsien_id,' ')       as  hsien_id,");               
	             sql.append("               nvl(cd01.hsien_name,'OTHER') as  hsien_name,");               
	             sql.append("               cd01.FR001W_output_order     as  FR001W_output_order,");               
	             sql.append("               bn01.bank_no ,  bn01.BANK_NAME,"); 
	             sql.append("               decode(sum(ex_type_feb),0,0,null,0,1) as bankcount_feb,");
	             sql.append("               decode(sum(ex_type_feb),null,0,sum(ex_type_feb)) as ex_type_feb,");
	             sql.append("               decode(sum(ex_type_agri),0,0,null,0,1) as bankcount_agri,");
	             sql.append("               decode(sum(ex_type_agri),null,0,sum(ex_type_agri))  as ex_type_agri,");
	             sql.append("               decode(sum(ex_type_boaf),0,0,null,0,1) as bankcount_boaf,");
	             sql.append("               decode(sum(ex_type_boaf),null,0,sum(ex_type_boaf)) as ex_type_boaf,");
	             sql.append("               decode(sum(ex_type_feb)+sum(ex_type_agri)+sum(ex_type_boaf),0,0,null,0,1) as bankcount_sum,");
	             sql.append("               decode(sum(ex_type_feb)+sum(ex_type_agri)+sum(ex_type_boaf),null,0,sum(ex_type_feb)+sum(ex_type_agri)+sum(ex_type_boaf)) as ex_type_sum"); 
	             sql.append("        from  (select * from  cd01 where cd01.hsien_id <> 'Y'  ) cd01"); 
	             sql.append("        left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = 100 and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL)");  
	             sql.append("        left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type in ('6','7') and bn01.m_year = 100 and bn_type <> '2' and wlx01.m_year=100");  
	             sql.append("        left join (select frm_exmaster.bank_no,frm_exmaster.ex_no,");
	             sql.append("                          decode(ex_type,'FEB',1,0) as ex_type_feb,");
	             sql.append("                          decode(ex_type,'AGRI',1,0) as ex_type_agri,");
	             sql.append("                          decode(ex_type,'BOAF',1,0) as ex_type_boaf");
	             sql.append("                    from frm_exmaster");
	             sql.append("                    left join frm_snrtdoc on frm_exmaster.ex_no = frm_snrtdoc.ex_no and frm_exmaster.bank_no=frm_snrtdoc.bank_no");
	             sql.append("                   where (ex_type='FEB' and TO_CHAR(doc_date ,'yyyymmdd') BETWEEN '"+begDate+"' AND '"+endDate+"')");
                 sql.append("                   or (ex_type='AGRI' and frm_exmaster.ex_no >= '"+beg_season+"' and frm_exmaster.ex_no <= '"+end_season+"')");//105年第1季=1050101~1050331; 105年第2季=1050401~1050631所挑選的受檢單位代碼; 105年第3季=1050701~1050931; 105年第4季=1051001~1051231
                 sql.append("                   or (ex_type='BOAF' and frm_exmaster.ex_no >= '"+begDate+"' and frm_exmaster.ex_no <= '"+endDate+"')");
	             sql.append("                    group by frm_exmaster.bank_no,frm_exmaster.ex_no,ex_type");
	             sql.append("        ) frm_exmaster  on  bn01.bank_no = frm_exmaster.bank_no ");
	             sql.append("        group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME");
	             sql.append("      ) frm_exmaster where  frm_exmaster.bank_no <> ' '");
	             sql.append(" order by fr001w_output_order");	    
	             dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "bankcount_feb,ex_type_feb,bankcount_agri,ex_type_agri,bankcount_boaf,ex_type_boaf,bankcount_sum,ex_type_sum");
	          }else{          
	             //所選農漁會統計.報表查詢SQL          
                 sql.append(" select a.bank_no,");//農漁會別.機構代碼
                 sql.append(" bank_name,");//農漁會別.機構名稱
                 sql.append(" fr001w_output_order,");//縣市別排序欄位
                 sql.append(" sum(ex_type_feb) as ex_type_feb,");//金管會檢查報告案件數
                 sql.append(" sum(ex_type_agri) as ex_type_agri,");//農業金庫查核案件數
                 sql.append(" sum(ex_type_boaf) as ex_type_boaf,");//農金局訪查案件數
                 sql.append(" sum(ex_type_feb)+sum(ex_type_agri)+sum(ex_type_boaf) as ex_type_sum"); //案件數合計
                 sql.append(" from (");
                 sql.append(" select frm_exmaster.bank_no,frm_exmaster.ex_no, fr001w_output_order");
                 sql.append(" ,decode(ex_type,'FEB',1,0) as ex_type_feb");
                 sql.append(" ,decode(ex_type,'AGRI',1,0) as ex_type_agri");
                 sql.append(" ,decode(ex_type,'BOAF',1,0) as ex_type_boaf");
                 sql.append(" from frm_exmaster");
                 sql.append(" left join frm_snrtdoc on frm_exmaster.ex_no = frm_snrtdoc.ex_no and frm_exmaster.bank_no=frm_snrtdoc.bank_no");
                 sql.append(" left join (select wlx01.*,cd01.fr001w_output_order from wlx01 left join cd01 on  wlx01.hsien_id=cd01.hsien_id where m_year =100)wlx01 on frm_exmaster.bank_no=wlx01.bank_no");
                 sql.append(" where frm_exmaster.bank_no in ("+selectitem+")");// --所挑選的受檢單位代碼
                 sql.append(" and ((ex_type='FEB' and TO_CHAR(doc_date ,'yyyymmdd') BETWEEN '"+begDate+"' AND '"+endDate+"')");
                 sql.append(" or (ex_type='AGRI' and frm_exmaster.ex_no >= '"+beg_season+"' and frm_exmaster.ex_no <= '"+end_season+"')");//105年第1季=1050101~1050331; 105年第2季=1050401~1050631所挑選的受檢單位代碼; 105年第3季=1050701~1050931; 105年第4季=1051001~1051231
                 sql.append(" or (ex_type='BOAF' and frm_exmaster.ex_no >= '"+begDate+"' and frm_exmaster.ex_no <= '"+endDate+"'))");
                 sql.append(" group by frm_exmaster.bank_no,frm_exmaster.ex_no,ex_type,fr001w_output_order)a");
                 sql.append(" left join (select * from bn01 where m_year=100)bn01 on a.bank_no=bn01.bank_no");
                 sql.append(" group by a.bank_no,bank_name,fr001w_output_order");
                 sql.append(" union");
                 sql.append(" select '9999999','合計','999',sum(ex_type_feb) as ex_type_feb,sum(ex_type_agri) as ex_type_agri,sum(ex_type_boaf) as ex_type_boaf,");
                 sql.append(" sum(ex_type_feb)+sum(ex_type_agri)+sum(ex_type_boaf) as ex_type_sum");
                 sql.append(" from (");
                 sql.append(" select frm_exmaster.bank_no,frm_exmaster.ex_no,decode(ex_type,'FEB',1,0) as ex_type_feb");
                 sql.append(" ,decode(ex_type,'AGRI',1,0) as ex_type_agri");
                 sql.append(" ,decode(ex_type,'BOAF',1,0) as ex_type_boaf");
                 sql.append(" from frm_exmaster");
                 sql.append(" left join frm_snrtdoc on frm_exmaster.ex_no = frm_snrtdoc.ex_no and frm_exmaster.bank_no=frm_snrtdoc.bank_no");
                 sql.append(" where frm_exmaster.bank_no in ("+selectitem+")");// --所挑選的受檢單位代碼
                 sql.append(" and ((ex_type='FEB' and TO_CHAR(doc_date ,'yyyymmdd') BETWEEN '"+begDate+"' AND '"+endDate+"')");
                 sql.append(" or (ex_type='AGRI' and frm_exmaster.ex_no >= '"+beg_season+"' and frm_exmaster.ex_no <= '"+end_season+"')");//105年第1季=1050101~1050331; 105年第2季=1050401~1050631所挑選的受檢單位代碼; 105年第3季=1050701~1050931; 105年第4季=1051001~1051231
                 sql.append(" or (ex_type='BOAF' and frm_exmaster.ex_no >= '"+begDate+"' and frm_exmaster.ex_no <= '"+endDate+"'))");
                 sql.append(" group by frm_exmaster.bank_no,frm_exmaster.ex_no,ex_type)a");
                 sql.append(" order by fr001w_output_order ");
                 dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "ex_type_feb,ex_type_agri,ex_type_boaf,ex_type_sum");
	          }
	      }    
           
          System.out.println("dbData.size=" + dbData.size());	      
	      //設定報表表頭資料============================================
          row = sheet.getRow(1);
          cell = row.getCell((short) 0);
          //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
          cell.setEncoding(HSSFCell.ENCODING_UTF_16);        
          if (dbData.size() == 0) {
              cell.setCellValue("查核期間:"+S_YEAR + "年" + S_MONTH + "月"+S_DAY+"日至"+E_YEAR+"年"+E_MONTH+"月"+E_DAY+"日無資料存在");
          } else {
              cell.setCellValue("查核期間:"+S_YEAR + "年" + S_MONTH + "月"+S_DAY+"日至"+E_YEAR+"年"+E_MONTH+"月"+E_DAY+"日");
               
              
              if(rptKind.equals("1")){//統計分類:貸款種類
                  rowNum=4;         
                  cellNum=9;
              }else if(rptKind.equals("2")){//統計分類:缺失態樣
                  rowNum=4;         
                  cellNum=11;
              }else if(rptKind.equals("3")){  
                  if(hasBankListALL.equals("true")){//受檢單位選全部時(縣市別統計)
                      rowNum=4;         
                      cellNum=9;
                  }else{//所選農漁會統計  
                      rowNum=3;   
                      cellNum=5;
                  } 
              } 
              
              for(int i=0;i<dbData.size();i++){
                  row = sheet.createRow(rowNum);
                  bean = (DataObject)dbData.get(i);                       
                  for(int cellcount=0;cellcount<cellNum;cellcount++){                                       
                    cell = (row.getCell((short) cellcount) == null) ? row.createCell((short) cellcount) : row.getCell((short) cellcount);         
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    if(rptKind.equals("1")){//統計分類:貸款種類
                        cell.setCellStyle(cs_center);
                        if(cellcount == 0) cell.setCellValue(((bean.getValue("loan_item_name")==null)?"":(bean.getValue("loan_item_name")).toString()));//貸款種類
                        if(cellcount == 1) cell.setCellValue((bean.getValue("bankcount_feb") == null)?"0":(bean.getValue("bankcount_feb")).toString());//--金管會檢查報告.農漁會家數
                        if(cellcount == 2) cell.setCellValue((bean.getValue("ex_type_feb") == null)?"0":(bean.getValue("ex_type_feb")).toString());//--金管會檢查報告.案件數
                        if(cellcount == 3) cell.setCellValue((bean.getValue("bankcount_agri") == null)?"0":(bean.getValue("bankcount_agri")).toString());//--農業金庫查核.農漁會家數                         
                        if(cellcount == 4) cell.setCellValue((bean.getValue("ex_type_agri") == null)?"0":(bean.getValue("ex_type_agri")).toString());//--農業金庫查核.案件數                           
                        if(cellcount == 5) cell.setCellValue((bean.getValue("bankcount_boaf") == null)?"0":(bean.getValue("bankcount_boaf")).toString());//--農金局訪查.農漁會家數                          
                        if(cellcount == 6) cell.setCellValue((bean.getValue("ex_type_boaf") == null)?"0":(bean.getValue("ex_type_boaf")).toString());//--農金局訪查.案件數                           
                        if(cellcount == 7) cell.setCellValue((bean.getValue("bankcount_sum") == null)?"0":(bean.getValue("bankcount_sum")).toString());//--合計.農漁會家數                           
                        if(cellcount == 8) cell.setCellValue((bean.getValue("ex_type_sum") == null)?"0":(bean.getValue("ex_type_sum")).toString());//--案件數合計
                    }else if(rptKind.equals("2")){//統計分類:缺失態樣
                        if(cellcount<=2){
                            cell.setCellStyle(cs_left);
                        }else{
                            cell.setCellStyle(cs_center);
                        }
                        if(cellcount == 0) cell.setCellValue(((bean.getValue("def_name")==null)?"":(bean.getValue("def_name")).toString()));//缺失態樣
                        if(cellcount == 1) cell.setCellValue(((bean.getValue("def_case_no")==null)?"":(bean.getValue("def_case_no")).toString()));//缺失情節代碼
                        if(cellcount == 2) cell.setCellValue(((bean.getValue("case_name")==null)?"":(bean.getValue("case_name")).toString()));//缺失情節名稱
                        if(cellcount == 3) cell.setCellValue((bean.getValue("bankcount_feb") == null)?"0":(bean.getValue("bankcount_feb")).toString());//--金管會檢查報告.農漁會家數
                        if(cellcount == 4) cell.setCellValue((bean.getValue("ex_type_feb") == null)?"0":(bean.getValue("ex_type_feb")).toString());//--金管會檢查報告.案件數
                        if(cellcount == 5) cell.setCellValue((bean.getValue("bankcount_agri") == null)?"0":(bean.getValue("bankcount_agri")).toString());//--農業金庫查核.農漁會家數                         
                        if(cellcount == 6) cell.setCellValue((bean.getValue("ex_type_agri") == null)?"0":(bean.getValue("ex_type_agri")).toString());//--農業金庫查核.案件數                           
                        if(cellcount == 7) cell.setCellValue((bean.getValue("bankcount_boaf") == null)?"0":(bean.getValue("bankcount_boaf")).toString());//--農金局訪查.農漁會家數                          
                        if(cellcount == 8) cell.setCellValue((bean.getValue("ex_type_boaf") == null)?"0":(bean.getValue("ex_type_boaf")).toString());//--農金局訪查.案件數                           
                        if(cellcount == 9) cell.setCellValue((bean.getValue("bankcount_sum") == null)?"0":(bean.getValue("bankcount_sum")).toString());//--合計.農漁會家數                           
                        if(cellcount == 10) cell.setCellValue((bean.getValue("ex_type_sum") == null)?"0":(bean.getValue("ex_type_sum")).toString());//--案件數合計
                    }else if(rptKind.equals("3")){//統計分類:農漁會別  
                        if(hasBankListALL.equals("true")){//受檢單位選全部時(縣市別統計)                          
                            cell.setCellStyle(cs_center);
                            if(cellcount == 0) cell.setCellValue(((bean.getValue("hsien_name")==null)?"":(bean.getValue("hsien_name")).toString()));//縣市別
                            if(cellcount == 1) cell.setCellValue((bean.getValue("bankcount_feb") == null)?"0":(bean.getValue("bankcount_feb")).toString());//--金管會檢查報告.農漁會家數
                            if(cellcount == 2) cell.setCellValue((bean.getValue("ex_type_feb") == null)?"0":(bean.getValue("ex_type_feb")).toString());//--金管會檢查報告.案件數
                            if(cellcount == 3) cell.setCellValue((bean.getValue("bankcount_agri") == null)?"0":(bean.getValue("bankcount_agri")).toString());//--農業金庫查核.農漁會家數                         
                            if(cellcount == 4) cell.setCellValue((bean.getValue("ex_type_agri") == null)?"0":(bean.getValue("ex_type_agri")).toString());//--農業金庫查核.案件數                           
                            if(cellcount == 5) cell.setCellValue((bean.getValue("bankcount_boaf") == null)?"0":(bean.getValue("bankcount_boaf")).toString());//--農金局訪查.農漁會家數                          
                            if(cellcount == 6) cell.setCellValue((bean.getValue("ex_type_boaf") == null)?"0":(bean.getValue("ex_type_boaf")).toString());//--農金局訪查.案件數                           
                            if(cellcount == 7) cell.setCellValue((bean.getValue("bankcount_sum") == null)?"0":(bean.getValue("bankcount_sum")).toString());//--合計.農漁會家數                           
                            if(cellcount == 8) cell.setCellValue((bean.getValue("ex_type_sum") == null)?"0":(bean.getValue("ex_type_sum")).toString());//--案件數合計                           
                        }else{//所選農漁會統計
                            if(cellcount==0){
                                cell.setCellStyle(cs_left);
                            }else{
                                cell.setCellStyle(cs_center);
                            }
                            if(cellcount == 0){
                                if(!((bean.getValue("bank_no")).toString()).equals("9999999")){
                                    cell.setCellValue(((bean.getValue("bank_no")==null)?"":(bean.getValue("bank_no")).toString())+((bean.getValue("bank_name")==null)?"":(bean.getValue("bank_name")).toString()));
                                }else{
                                    cell.setCellValue(((bean.getValue("bank_name")==null)?"":(bean.getValue("bank_name")).toString()));
                                }
                            }       
                            if(cellcount == 1) cell.setCellValue((bean.getValue("ex_type_feb") == null)?"0":(bean.getValue("ex_type_feb")).toString());
                            if(cellcount == 2) cell.setCellValue((bean.getValue("ex_type_agri") == null)?"0":(bean.getValue("ex_type_agri")).toString());
                            if(cellcount == 3) cell.setCellValue((bean.getValue("ex_type_boaf") == null)?"0":(bean.getValue("ex_type_boaf")).toString());
                            if(cellcount == 4) cell.setCellValue((bean.getValue("ex_type_sum") == null)?"0":(bean.getValue("ex_type_sum")).toString());                           
                        }
                     }//統計分類:農漁會別
                  }
                  rowNum++;
              } 
              
              if(rptKind.equals("2")){//統計分類:缺失態樣
                  sheet.addMergedRegion( new Region( ( short )rowNum-1, ( short )0,( short )rowNum-1,( short )2) );
                  row = sheet.getRow(rowNum-1);
                  cell = row.getCell((short) 0);
                  cell.setCellStyle(cs_center);
                  cell.setCellValue("合計");
              }
              
              row = sheet.createRow(rowNum);
              cell = (row.getCell((short) 0) == null) ? row.createCell((short) 0) : row.getCell((short) 0);            
              cell.setCellValue("");
              
              rowNum++;              
              row = sheet.createRow(rowNum);
              cell = row.createCell( (short)0);
              cell.setEncoding(HSSFCell.ENCODING_UTF_16);
              //cell.setCellStyle(cs_noborderleft);     
              if(rptKind.equals("3")){//統計分類:農漁會別  
                  if(hasBankListALL.equals("true")){//受檢單位選全部時(縣市別統計) 
                      cell.setCellValue("註：金管會檢查報告統計該段期間發文之缺失案件，農業金庫查核及農金局訪查則統計該段期間所有抽核件數");
                  }else{
                      cell.setCellValue("註1：金管會檢查報告統計該段期間發文之缺失案件，農業金庫查核及農金局訪查則統計該段期間所有抽核件數");
                      rowNum++;
                      row = sheet.createRow(rowNum);
                      cell = row.createCell( (short)0);
                      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                      //cell.setCellStyle(cs_noborderleft);
                      cell.setCellValue("註2：依農漁會所屬縣市別排序"); 
                  }  
              }
	      }
	     
	     
	      FileOutputStream fout = null;     
	      fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + "DS068W.xls");
	     
	      HSSFFooter footer = sheet.getFooter();
	      footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
	      footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
	      wb.write(fout);
	      //儲存
	      fout.close();
	      System.out.println("儲存成功!");
	      
	    }catch (Exception e) {
	      System.out.println("RptDS068W.createRpt Error:" + e + e.getMessage());
	    }
	    
	    return errMsg;
	  }
}
