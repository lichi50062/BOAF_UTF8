/*
	94.03.11 Fix 金額資料為零是不輸出0，改為輸出空白及欄位資料右靠處理
	94.03.15 ADD 月份輸入條件查詢 	by EGG
	94.08.04 fix a01 改成  ((select * from a01) union (select * from a02)) a01
	94.08.23 add 明細表 by 2295
	94.08.30 fix 放款總額（fieldB）小括號位置	jwang
	94.08.30 add 變數「fromtable2」	jwang
	99.09.16 fix 根據查詢年度.100年以後取得新縣市別.100年以前取得舊縣市別 
  			      使用PreparedStatement;並列印轉換後的SQL;套用QueryDB_SQLParam by 2295
   100.10.13 fix 101年.顯示更改後的組織名稱(行政院農業委員會->行政院農業部) by 2295
   102.06.21 fix SQL調整   by2968	
   102.08.27 fix 總表100年後已無高雄縣  by2968
   106.08.07 add 農會總表的其他.顯示成中華民國農會 by 2295		      
*/
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import java.io.*;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR001W {
    public static String createRpt(String s_year,String s_month,String unit,String datestate,String bank_type,String rptStyle){
    	System.out.println("inpute s_month = "+s_month);
		String errMsg = "";
		StringBuffer sqlCmd = new StringBuffer();
		StringBuffer sqlCmd_sum = new StringBuffer();
		StringBuffer sqlCmd_sumtaiwan = new StringBuffer(); 
		StringBuffer sqlCmd_sumfuchien = new StringBuffer(); 
		StringBuffer sqlCmd_each = new StringBuffer(); 
		StringBuffer sqlCmd_detail = new StringBuffer(); 
		List paramList = new ArrayList();
		List paramList_detail = new ArrayList();
		int rowNum=0;
		int i=0;
		int j=0;
		String s_year_last="";
		String s_month_last="";
		String hsien_id_sum[]={"","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p"};
		String bank_type_name=(bank_type.equals("6"))?"農會":"漁會";
		String unit_name="";
		String filename="";
		filename=(bank_type.equals("6"))?"台閩地區農會信用部經營指標.xls":"台閩地區漁會信用部經營指標.xls";
		List dbData_sum = null;
		List dbData_sumtaiwan = null;
		List dbData_sumfuchien = null;
		List dbData_each = null;
		List dbData_detail = null;
		reportUtil reportUtil = new reportUtil();
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
    		
    		String openfile="台閩地區農漁會信用部經營指標"+(rptStyle.equals("0")?((Integer.parseInt(s_year) <= 99)?"_99":(bank_type.equals("6"))?"_農會":"_漁會"):"_明細表")+".xls";
    		System.out.println("open file "+openfile);
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
	        if(rptStyle.equals("0")){
                ps.setScale( ( short )85 ); //列印縮放百分比
            }else{
                ps.setScale( ( short )90 ); //列印縮放百分比 
            }

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
	  		
	  		//總表共用SQL
	  		sqlCmd.append("         round(sum(a01.fieldB)/ ?,0)            as fieldB , "); //--存款總額
	  		sqlCmd.append("         round(sum(a01.fieldC)/ ?,0)            as fieldC , "); //--放款總額 
	  		sqlCmd.append("         decode(sum(a01.fieldB),0,0,round(sum(a01.fieldC) / sum(a01.fieldB) *100 ,2))  as   fieldD, ");  //--存放比
	  		sqlCmd.append("         round(sum(a01.fieldE)/ ?,0)            as fieldE , "); //--逾期放款餘額
	  		sqlCmd.append("         decode(sum(a01.field_120000),0,0,round(sum(a01.fieldE) / sum(a01.field_120000) *100 ,2))  as   fieldF, "); //--逾期放款比率         
	  		sqlCmd.append("         round(sum(a01.fieldG)/ ?,0)            as fieldG , "); //--催收款項
	  		sqlCmd.append("         round(sum(a01.fieldH)/ ?,0)            as fieldH , "); //--備抵呆帳
	  		sqlCmd.append("         decode(sum(a01.fieldE),0,0,round(sum(a01.fieldH) / sum(a01.fieldE) *100 ,2))  as   fieldI, "); //--備抵呆帳占逾期放款比  
	  		sqlCmd.append("         decode(sum(a01.field_120000),0,0,round(sum(a01.fieldH) / sum(a01.field_120000) *100 ,2))  as   fieldJ, "); //--備抵呆帳占放款總額比        
	  		sqlCmd.append("         round(sum(a01.fieldK)/ ?,0)            as fieldK , "); //--內部透支及內部融資
	  		sqlCmd.append("         round(sum(a01.fieldL)/ ?,0)            as fieldL , "); //--固定資產淨額
	  		sqlCmd.append("         round(sum(a01.fieldM)/ ?,0)            as fieldM , "); //--本期損益
	  		sqlCmd.append("         round(sum(a01.fieldN)/ ?,0)            as fieldN , "); //--事業資金及公積
	  		sqlCmd.append("         round(sum(a01.fieldO)/ ?,0)            as fieldO , "); //--盈虧及損益
	  		sqlCmd.append("         round(sum(a01.fieldP)/ ?,0)            as fieldP   "); //--淨值
	  		for(int k=1;k<=11;k++){
	  		  paramList.add(unit);
	  		}
	  		sqlCmd.append(" from (select * from ").append(cd01_table).append(")cd01 "); 
	  		sqlCmd.append(" left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id "); 
	  		paramList.add(wlx01_m_year);
	  		sqlCmd.append(" left join "); 
	  		sqlCmd.append(" ( select bank_code, "); 
	  		sqlCmd.append("   sum(fieldB) as fieldB , sum(fieldC) as fieldC , sum(fieldD) as fieldD ,sum(fieldE) as fieldE , ");           
	  		sqlCmd.append("   sum(fieldF) as fieldF , sum(fieldG) as fieldG , sum(fieldH) as fieldH ,sum(fieldI) as fieldI ,sum(fieldJ) as fieldJ , ");   
	  		sqlCmd.append("   sum(fieldK) as fieldK , sum(fieldL) as fieldL , sum(fieldM) as fieldM ,sum(fieldN) as fieldN , ");   
	  		sqlCmd.append("   sum(fieldO) as fieldO , sum(fieldP) as fieldP , sum(field_120000) as field_120000 ");    
	  		sqlCmd.append("   from ( "); 
	  		sqlCmd.append("         ( select bank_code, fieldB, fieldC, 0 as fieldD, ");             
	  		sqlCmd.append("                   0 as fieldE, 0 as fieldF, 0 as fieldG, 0 as fieldH, 0 as fieldI, ");             
	  		sqlCmd.append("                   0 as fieldJ, 0 as fieldK, 0 as fieldL, 0 as fieldM, 0 as fieldN, ");            
	  		sqlCmd.append("                   0 as fieldO, 0 as fieldP, 0 as field_120000 ");      
	  		sqlCmd.append("            from (select  bank_code, ");                 
	  		sqlCmd.append("                  round((sum(decode(a01.acc_code,'220000',amt,0))-round(sum(decode(a01.acc_code,'220900',amt,0))/2,0))/ ?,0) as fieldB, ");
	  		sqlCmd.append("                  round(sum(decode(a01.acc_code,'120000',amt,0))/ ?,0)         as fieldC   ");
	  		paramList.add(div);
	  		paramList.add(div);
	  		sqlCmd.append("                  from a01,(select ba01.bank_type,wlx01.hsien_id,ba01.bank_no ");                    
	  		sqlCmd.append("                  from (select * from ba01 where m_year=? and bank_type=?)ba01,(select * from wlx01 where m_year=?)wlx01 ");                    
	  		paramList.add(wlx01_m_year);
            paramList.add(bank_type);
            paramList.add(wlx01_m_year);
	  		sqlCmd.append("                  where ba01.bank_no = wlx01.bank_no)ba01 ");          
	  		sqlCmd.append("                  where ((a01.m_year =? and a01.m_month=? ) "); 
	  		sqlCmd.append(" or (a01.m_year =? and a01.m_month=? )) ");
	  		paramList.add(s_year);
            paramList.add(s_month);
            paramList.add(s_year_last);
            paramList.add(s_month_last);
	  		sqlCmd.append("                  and a01.bank_code=ba01.bank_no ");          
	  		sqlCmd.append("                  GROUP BY a01.bank_code ");         
	  		sqlCmd.append("                 ) a01 "); 
	  		sqlCmd.append("           ) ");          
	  		sqlCmd.append("          union all "); 
	  		sqlCmd.append("         (SELECT bank_code,  0  as fieldB, 0  as fieldC,  0  as fieldD,sum(decode(a01.acc_code,'990000',amt,0)) AS fieldE, ");            
	  		sqlCmd.append("                 round(decode(sum(decode(a01.acc_code,'120000',amt,0)),0,0,sum(decode(a01.acc_code,'990000',amt,0))/sum(decode(a01.acc_code,'120000',amt,0))),4) AS fieldF, ");            
	  		sqlCmd.append("                 sum(decode(a01.acc_code,'150200',amt,0)) AS fieldG,    sum(decode(a01.acc_code,'120800',amt,'150300',amt,0)) AS fieldH, ");    
	  		sqlCmd.append("                 0  as fieldI,0 as fieldJ, ");  
	  		sqlCmd.append("                  decode(bank_type,'6',sum(decode(a01.acc_code,'120700',amt,0)) ,'7',sum(decode(a01.acc_code,'120900',amt,0))) AS fieldK, ");
	  		sqlCmd.append("                  sum(decode(a01.acc_code,'140000',amt,0)) AS fieldL, ");             
	  		sqlCmd.append("                  sum(decode(a01.acc_code,'320300',amt,0)) AS fieldM, ");            
	  		sqlCmd.append("                  sum(decode(a01.acc_code,'310000',amt,0)) AS fieldN, ");            
	  		sqlCmd.append("                  sum(decode(a01.acc_code,'320000',amt,0)) AS fieldO, ");            
	  		sqlCmd.append("                  decode(bank_type,'6',sum(decode(a01.acc_code,'310000',amt,'320000',amt,0)),'7',sum(decode(a01.acc_code,'300000',amt,0))) AS fieldP, ");
	  		sqlCmd.append("                  sum(decode(a01.acc_code,'120000',amt,0)) as field_120000 ");    
	  		sqlCmd.append("           from a01,(select ba01.bank_type,wlx01.hsien_id,ba01.bank_no ");               
	  		sqlCmd.append("           from (select * from ba01 where m_year=? and bank_type=?)ba01,(select * from wlx01 where m_year=?)wlx01 ");              
	  		sqlCmd.append("           where ba01.bank_no = wlx01.bank_no)ba01 WHERE (a01.m_year= ? and a01.m_month = ? ) "); 
	  		paramList.add(wlx01_m_year);
            paramList.add(bank_type);
            paramList.add(wlx01_m_year);
            paramList.add(s_year);
            paramList.add(s_month);
	  		sqlCmd.append("           and a01.bank_code=ba01.bank_no  group by ba01.bank_type,a01.bank_code ");
	  		sqlCmd.append("          ) "); 
	  		sqlCmd.append("         )  a01 "); 
	  		sqlCmd.append("   group by a01.bank_code "); 
	  		sqlCmd.append(" ) a01 on a01.bank_code=wlx01.bank_no "); 
	  		//總表-總計 
            sqlCmd_sum.append("select "+ sqlCmd);
	  		sqlCmd_sum.append(" where cd01.hsien_id <> 'Y' ");
	  		//總表-台灣省小計
            sqlCmd_sumtaiwan.append("select "+ sqlCmd);
            sqlCmd_sumtaiwan.append(" where cd01.Hsien_div = '2' ");
	  		//總表-福建省小計
            sqlCmd_sumfuchien.append("select "+ sqlCmd);
            sqlCmd_sumfuchien.append(" where cd01.Hsien_div = '3' ");
	  		//總表-各縣市小計
            sqlCmd_each.append("select  nvl(cd01.hsien_id,' ') as hsien_id, nvl(cd01.hsien_name,'其他') as hsien_name, "+ sqlCmd );
            sqlCmd_each.append(" where  cd01.hsien_id <> 'Y' ");  
            sqlCmd_each.append(" group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'其他'),cd01.FR001W_output_order "); 
            sqlCmd_each.append(" order by cd01.FR001W_output_order ");
            
            //明細表
            sqlCmd_detail.append(" select  nvl(cd01.hsien_id,' ') as hsien_id,  nvl(cd01.hsien_name,'其他') as hsien_name,cd01.FR001W_output_order, ");
            sqlCmd_detail.append("         ba01.bank_no  as bank_no ,  ba01.BANK_NAME  as BANK_NAME, ");
            sqlCmd_detail.append("         round(sum(a01.fieldB)/ ?,0)            as fieldB , "); //--存款總額
            sqlCmd_detail.append("         round(sum(a01.fieldC)/ ?,0)            as fieldC , "); //--放款總額 
            sqlCmd_detail.append("         decode(sum(a01.fieldB),0,0,round(sum(a01.fieldC) / sum(a01.fieldB) *100 ,2))  as   fieldD, ");  //--存放比
            sqlCmd_detail.append("         round(sum(a01.fieldE)/ ?,0)            as fieldE , "); //--逾期放款餘額     
            sqlCmd_detail.append("         decode(sum(a01.field_120000),0,0,round(sum(a01.fieldE) / sum(a01.field_120000) *100 ,2))  as   fieldF, ");  //--逾期放款比率        
            sqlCmd_detail.append("         round(sum(a01.fieldG)/ ?,0)            as fieldG , "); //--催收款項
            sqlCmd_detail.append("         round(sum(a01.fieldH)/ ?,0)            as fieldH , "); //--備抵呆帳
            sqlCmd_detail.append("         decode(sum(a01.fieldE),0,0,round(sum(a01.fieldH) / sum(a01.fieldE) *100 ,2))  as   fieldI, ");  //--備抵呆帳占逾期放款比  
            sqlCmd_detail.append("         decode(sum(a01.field_120000),0,0,round(sum(a01.fieldH) / sum(a01.field_120000) *100 ,2))  as   fieldJ, ");  //--備抵呆帳占放款總額比
            sqlCmd_detail.append("         round(sum(a01.fieldK)/ ?,0)            as fieldK , "); //--內部透支及內部融資
            sqlCmd_detail.append("         round(sum(a01.fieldL)/ ?,0)            as fieldL , "); //--固定資產淨額
            sqlCmd_detail.append("         round(sum(a01.fieldM)/ ?,0)            as fieldM , "); //--本期損益
            sqlCmd_detail.append("         round(sum(a01.fieldN)/ ?,0)            as fieldN , "); //--事業資金及公積
            sqlCmd_detail.append("         round(sum(a01.fieldO)/ ?,0)            as fieldO , "); //--盈虧及損益
            sqlCmd_detail.append("         round(sum(a01.fieldP)/ ?,0)            as fieldP   "); //--淨值
            for(int k=1;k<=11;k++){
                paramList_detail.add(unit);
            }
            sqlCmd_detail.append(" from (select * from ").append(cd01_table).append(")cd01, ");
            sqlCmd_detail.append("      (select ba01.bank_type,ba01.bank_name,wlx01.hsien_id,ba01.bank_no "); 
            sqlCmd_detail.append("       from (select * from ba01 where m_year=? and bank_type=?)ba01,(select * from wlx01 where m_year=?)wlx01 ");
            paramList_detail.add(wlx01_m_year);
            paramList_detail.add(bank_type);
            paramList_detail.add(wlx01_m_year);
            sqlCmd_detail.append("             where ba01.bank_no = wlx01.bank_no)ba01, ");
            sqlCmd_detail.append(" ( select bank_code, "); 
            sqlCmd_detail.append("   sum(fieldB) as fieldB , sum(fieldC) as fieldC , sum(fieldD) as fieldD ,sum(fieldE) as fieldE , ");           
            sqlCmd_detail.append("   sum(fieldF) as fieldF , sum(fieldG) as fieldG , sum(fieldH) as fieldH ,sum(fieldI) as fieldI ,sum(fieldJ) as fieldJ , ");   
            sqlCmd_detail.append("   sum(fieldK) as fieldK , sum(fieldL) as fieldL , sum(fieldM) as fieldM ,sum(fieldN) as fieldN , ");   
            sqlCmd_detail.append("   sum(fieldO) as fieldO , sum(fieldP) as fieldP , sum(field_120000) as field_120000 ");    
            sqlCmd_detail.append("   from ( "); 
            sqlCmd_detail.append("         ( select bank_code, fieldB, fieldC, 0 as fieldD, ");        
            sqlCmd_detail.append("                   0 as fieldE, 0 as fieldF, 0 as fieldG, 0 as fieldH, 0 as fieldI, ");             
            sqlCmd_detail.append("                   0 as fieldJ, 0 as fieldK, 0 as fieldL, 0 as fieldM, 0 as fieldN, ");             
            sqlCmd_detail.append("                   0 as fieldO, 0 as fieldP, 0 as field_120000 ");      
            sqlCmd_detail.append("            from (select  bank_code, ");                 
            sqlCmd_detail.append("                  round((sum(decode(a01.acc_code,'220000',amt,0))-round(sum(decode(a01.acc_code,'220900',amt,0))/2,0))/?,0) as fieldB, "); 
            sqlCmd_detail.append("                  round(sum(decode(a01.acc_code,'120000',amt,0))/?,0)    as fieldC ");
            paramList_detail.add(div);
            paramList_detail.add(div);
            sqlCmd_detail.append("                  from a01,(select ba01.bank_type,wlx01.hsien_id,ba01.bank_no ");                    
            sqlCmd_detail.append("                  from (select * from ba01 where m_year=? and bank_type=? )ba01,(select * from wlx01 where m_year=? )wlx01 ");                    
            paramList_detail.add(wlx01_m_year);
            paramList_detail.add(bank_type);
            paramList_detail.add(wlx01_m_year);
            sqlCmd_detail.append("                  where ba01.bank_no = wlx01.bank_no)ba01 ");          
            sqlCmd_detail.append("                  where ((a01.m_year = ? and a01.m_month= ? ) "); 
            sqlCmd_detail.append(" or (a01.m_year = ? and a01.m_month= ? )) ");      
            paramList_detail.add(s_year);
            paramList_detail.add(s_month);
            paramList_detail.add(s_year_last);
            paramList_detail.add(s_month_last);
            sqlCmd_detail.append("                  and a01.bank_code=ba01.bank_no ");          
            sqlCmd_detail.append("                  GROUP BY a01.bank_code ");         
            sqlCmd_detail.append("                 ) a01 ");    
            sqlCmd_detail.append("           )  ");          
            sqlCmd_detail.append("          union all "); 
            sqlCmd_detail.append("         (SELECT bank_code,  0  as fieldB, 0  as fieldC,  0  as fieldD,sum(decode(a01.acc_code,'990000',amt,0)) AS fieldE, ");            
            sqlCmd_detail.append("                 round(decode(sum(decode(a01.acc_code,'120000',amt,0)),0,0,sum(decode(a01.acc_code,'990000',amt,0))/sum(decode(a01.acc_code,'120000',amt,0))),4) AS fieldF, ");                            
            sqlCmd_detail.append("                 sum(decode(a01.acc_code,'150200',amt,0)) AS fieldG,    sum(decode(a01.acc_code,'120800',amt,'150300',amt,0)) AS fieldH, ");            
            sqlCmd_detail.append("                  0  as fieldI,0 as fieldJ, ");
            sqlCmd_detail.append("                  decode(bank_type,'6',sum(decode(a01.acc_code,'120700',amt,0)) ,'7',sum(decode(a01.acc_code,'120900',amt,0))) AS fieldK, ");
            sqlCmd_detail.append("                  sum(decode(a01.acc_code,'140000',amt,0)) AS fieldL, ");             
            sqlCmd_detail.append("                  sum(decode(a01.acc_code,'320300',amt,0)) AS fieldM, ");            
            sqlCmd_detail.append("                  sum(decode(a01.acc_code,'310000',amt,0)) AS fieldN, ");            
            sqlCmd_detail.append("                  sum(decode(a01.acc_code,'320000',amt,0)) AS fieldO, ");            
            sqlCmd_detail.append("                  decode(bank_type,'6',sum(decode(a01.acc_code,'310000',amt,'320000',amt,0)),'7',sum(decode(a01.acc_code,'300000',amt,0))) AS fieldP, ");
            sqlCmd_detail.append("                  sum(decode(a01.acc_code,'120000',amt,0)) as field_120000 ");    
            sqlCmd_detail.append("           from a01,(select ba01.bank_type,wlx01.hsien_id,ba01.bank_no ");               
            sqlCmd_detail.append("           from (select * from ba01 where m_year=? and bank_type=? )ba01,(select * from wlx01 where m_year=? )wlx01 ");              
            sqlCmd_detail.append("           where ba01.bank_no = wlx01.bank_no)ba01 WHERE (a01.m_year= ? and a01.m_month = ? ) ");
            paramList_detail.add(wlx01_m_year);
            paramList_detail.add(bank_type);
            paramList_detail.add(wlx01_m_year);
            paramList_detail.add(s_year);
            paramList_detail.add(s_month);
            sqlCmd_detail.append("           and a01.bank_code=ba01.bank_no  group by ba01.bank_type,a01.bank_code ");
            sqlCmd_detail.append("          ) "); 
            sqlCmd_detail.append("         )  a01 "); 
            sqlCmd_detail.append("   group by a01.bank_code "); 
            sqlCmd_detail.append("  ) a01 "); 
            sqlCmd_detail.append(" where ba01.hsien_id=cd01.hsien_id  and ba01.bank_no=a01.bank_code ");                 
            sqlCmd_detail.append(" group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'其他'),cd01.FR001W_output_order, ");
            sqlCmd_detail.append(" ba01.bank_type,ba01.bank_no ,  ba01.BANK_NAME ");
            sqlCmd_detail.append(" order by cd01.FR001W_output_order,  ba01.bank_no ");
			
            if("0".equals(rptStyle)){//總表
			   dbData_sum = DBManager.QueryDB_SQLParam(sqlCmd_sum.toString(),paramList,"fieldb,fieldc,fieldd,fielde,fieldf,fieldg,fieldh,fieldi,fieldj,fieldk,fieldl,fieldm,fieldn,fieldo,fieldp");
			   System.out.println("總計的dbData_sum.size()="+dbData_sum.size());
			   dbData_sumtaiwan = DBManager.QueryDB_SQLParam(sqlCmd_sumtaiwan.toString(),paramList,"fieldb,fieldc,fieldd,fielde,fieldf,fieldg,fieldh,fieldi,fieldj,fieldk,fieldl,fieldm,fieldn,fieldo,fieldp");
			   System.out.println("台灣省小計的dbData_sumtaiwan.size()="+dbData_sumtaiwan.size());
			   dbData_sumfuchien = DBManager.QueryDB_SQLParam(sqlCmd_sumfuchien.toString(),paramList,"fieldb,fieldc,fieldd,fielde,fieldf,fieldg,fieldh,fieldi,fieldj,fieldk,fieldl,fieldm,fieldn,fieldo,fieldp");
			   System.out.println("福建省小計的dbData_sumfuchien.size()="+dbData_sumfuchien.size());
			   dbData_each = DBManager.QueryDB_SQLParam(sqlCmd_each.toString(),paramList,"fieldb,fieldc,fieldd,fielde,fieldf,fieldg,fieldh,fieldi,fieldj,fieldk,fieldl,fieldm,fieldn,fieldo,fieldp");
			   System.out.println("各別的dbData_each.size()="+dbData_each.size());
			}else{//明細表
				dbData_detail = DBManager.QueryDB_SQLParam(sqlCmd_detail.toString(),paramList_detail,"fieldb,fieldc,fieldd,fielde,fieldf,fieldg,fieldh,fieldi,fieldj,fieldk,fieldl,fieldm,fieldn,fieldo,fieldp");
				System.out.println("明細表的dbData_detail.size()="+dbData_detail.size());
			}
			//列印表頭
			row=(sheet.getRow(2)==null)? sheet.createRow(2) : sheet.getRow(2);
			cell=row.getCell((short)0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("台閩地區"+(bank_type.equals("6")?"農會":"漁會")+"信用部經營指標");
			//列印年度
			row=(sheet.getRow(3)==null)? sheet.createRow(3) : sheet.getRow(3);
			Calendar rightNow = Calendar.getInstance();
			String year = String.valueOf(rightNow.get(Calendar.YEAR)-1911);
			String month = String.valueOf(rightNow.get(Calendar.MONTH)+1);
			String day = String.valueOf(rightNow.get(Calendar.DAY_OF_MONTH));
			cell=row.getCell((short)0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			if(s_month.equals("0")){
				cell.setCellValue("中華民國　"+s_year+"　年度");
			}else {
				cell.setCellValue("中華民國　"+s_year+"　年 " + s_month + "月");
			}

			//列印單位
			cell=row.getCell((short)15);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			unit_name = Utility.getUnitName(unit);//取得單位名稱
			cell.setCellValue("單位：新台幣"+unit_name+"、％");

			short top=0,down=0;
			if(rptStyle.equals("0")){//總表
			    //先印總計
			    row=(sheet.getRow(6)==null)? sheet.createRow(6) : sheet.getRow(6);
			    if(dbData_sum != null && dbData_sum.size() != 0){
  			    	//row.setHeightInPoints((float)30);
			    	for(i=1;i<=15;i++){
			    		if(i==15){
			    		   insertCell(dbData_sum,true,0,"field"+hsien_id_sum[i],wb,row,(short)i, (short)26,top,down,HSSFCellStyle.BORDER_THIN,(short)0);
			    		}else{
  			    	       insertCell(dbData_sum,true,0,"field"+hsien_id_sum[i],wb,row,(short)i, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);
			    		}
			    	}
	       	    }
			    //台北市.高雄市
			    rowNum=7;
			    if(dbData_each != null && dbData_each.size() != 0){
			       //台北市(hsien_id=A);dbData_each.get(0)-->台北市
			       for(j=0;j<dbData_each.size();j++){
			           String hsienId = (((DataObject)dbData_each.get(j)).getValue("hsien_id")).toString();
    			       if( Integer.parseInt(s_year) >=100 && "W".equals(hsienId)){ //因100年表格金門縣連江縣的順序與SQL撈出來的順序不同，99年無此問題，故在此調整
    			           rowNum--;
                       }
    			       
    			       if(Integer.parseInt(s_year) <=99 && "h".equals(hsienId)){ //因99年表格其他的順序與SQL撈出來的順序不同，100年無此問題，故在此調整
    			           row=(sheet.getRow(31)==null)? sheet.createRow(31) : sheet.getRow(31);
    			       }else{
                           row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
    			       }
    			       
    			       if((j-1)==dbData_each.size()) down = HSSFCellStyle.BORDER_MEDIUM; //最後一筆時加上底線
                       else down = 0;
                       
                       for(i=1;i<=15;i++){
                            if(i==15){
                               insertCell(dbData_each,true,j,"field"+hsien_id_sum[i],wb,row,(short)i, (short)26,top,down,HSSFCellStyle.BORDER_THIN,(short)0);
                            }else{
                               insertCell(dbData_each,true,j,"field"+hsien_id_sum[i],wb,row,(short)i, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);
                            }
                       }
                       
			           //高雄市後.列印臺灣省小計
			           if((Integer.parseInt(s_year) <=99 && "E".equals(hsienId)) 
			           || (Integer.parseInt(s_year) >=100 && "e".equals(hsienId))){//-->高雄市
			           	  rowNum++;
			          	  //印臺灣省小計
			    		  row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
			    		  if(dbData_sumtaiwan != null && dbData_sumtaiwan.size() != 0){
		  	    			//row.setHeightInPoints((float)30);
			    			for(i=1;i<=15;i++){
			    				if(i==15){
			    				   insertCell(dbData_sumtaiwan,true,0,"field"+hsien_id_sum[i],wb,row,(short)i, (short)26,top,down,HSSFCellStyle.BORDER_THIN,(short)0);
			    				}else{
		  	    			       insertCell(dbData_sumtaiwan,true,0,"field"+hsien_id_sum[i],wb,row,(short)i, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);
			    				}
			    			}
			           	  }
			           }//end of 高雄市
			           
			           //列印福建省小計
			           if((Integer.parseInt(s_year) <=99 && "D".equals(hsienId)) //-->台南市後
			    	   || (Integer.parseInt(s_year) >=100 && "h".equals(hsienId)))//-->其他後
			           {
			              if(Integer.parseInt(s_year) >=100 && "h".equals(hsienId)){
			                  rowNum++;
			              }else if(Integer.parseInt(s_year) <=99 && "D".equals(hsienId)){
			                  rowNum+=2;
			              }
		          	      //印福建省小計
			    		  row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
			    		  if(dbData_sumfuchien != null && dbData_sumfuchien.size() != 0){
		  	    			//row.setHeightInPoints((float)30);
			    			for(i=1;i<=15;i++){
			    				if(i==15){
			    				   insertCell(dbData_sumfuchien,true,0,"field"+hsien_id_sum[i],wb,row,(short)i, (short)26,top,down,HSSFCellStyle.BORDER_THIN,(short)0);
			    				}else{
		  	    			       insertCell(dbData_sumfuchien,true,0,"field"+hsien_id_sum[i],wb,row,(short)i, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);
			    				}
			    			}
			           	  }
			           }//end of 福建省小計
			           if(!(Integer.parseInt(s_year) <=99 && "h".equals(hsienId))){
			               rowNum++;
			           }
			       }
			    }
			    int tailNum = 0;
			    if("100".equals(wlx01_m_year)){
			        tailNum = 36;
			    }else if("99".equals(wlx01_m_year)){
			        tailNum = 39;
			    }
			    row=(sheet.getRow(tailNum)==null)? sheet.createRow(tailNum) : sheet.getRow(tailNum);
                insertCell_tail("填表說明：",wb,row,(short)0,(short)0,(short)0);
                //insertCell_tail("1.本表編製一式三份，一份送行政院"+((Integer.parseInt(s_year)>= 101)?"農業部":"農業委員會")+"統計室，一份送本局會計室，一份自存。",wb,row,(short)1,(short)0,(short)0);
                insertCell_tail("1.本表編製一式三份，一份送行政院農業委員會統計室，一份送本局會計室，一份自存。",wb,row,(short)1,(short)0,(short)0);
                /*106.08.04 取消顯示
                if("6".equals(bank_type)){
                    tailNum++;
                    row=(sheet.getRow(tailNum)==null)? sheet.createRow(tailNum) : sheet.getRow(tailNum);
                    insertCell_tail("2.其他係指原臺灣省農會，該農會於102年5月22日更名為中華民國農會。",wb,row,(short)1,(short)0,(short)0);
                }
                */
			}else{//明細表
				System.out.println("明細表begin");
				rowNum=6;
				if(dbData_detail != null && dbData_detail.size() != 0){
				    for(j=0;j<dbData_detail.size();j++){
					    row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
					    if(j==dbData_detail.size()-1){//最後一筆時加上底線
					       	  down = HSSFCellStyle.BORDER_MEDIUM;
					       	  System.out.println("j.lastdata="+j);
					    }else{
					       	  down = 0;
					    }
					    insertCell(dbData_detail,true,j,"bank_name",wb,row,(short)0, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);
					    for(i=1;i<=15;i++){
					    	if(i==15){
					    		insertCell_detail(dbData_detail,true,j,"field"+hsien_id_sum[i],wb,row,(short)i, (short)26,top,down,HSSFCellStyle.BORDER_THIN,(short)0);
					    	}else{
					    		insertCell_detail(dbData_detail,true,j,"field"+hsien_id_sum[i],wb,row,(short)i, (short)26,top,down,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN);
					    	}
			            }
					    rowNum++;
				    }
			   }
				row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
				insertCell_tail("填表",wb,row,(short)0,HSSFCellStyle.BORDER_MEDIUM,(short)0);
				insertCell_tail("審核",wb,row,(short)2,HSSFCellStyle.BORDER_MEDIUM,(short)0);
				insertCell_tail("主辦業務人員",wb,row,(short)7,HSSFCellStyle.BORDER_MEDIUM,(short)0);
				insertCell_tail("機關長官",wb,row,(short)12,HSSFCellStyle.BORDER_MEDIUM,(short)0);

				rowNum++;
				row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
				insertCell_tail("主辦統計人員",wb,row,(short)7,(short)0,(short)0);

				rowNum++;
				row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
				insertBlank(15,wb,(short)0,(short)0,row);

				rowNum++;
				row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
				insertCell_tail("資料來源：根據各"+bank_type_name+"信用部資料彙編。",wb,row,(short)0,(short)0,(short)0);
				
				//100.10.13 fix 更改組織名稱       
				rowNum++;
				row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
				//insertCell_tail("填表說明：本表編製一式三份，一份送行政院"+((Integer.parseInt(s_year)>= 101)?"農業部":"農業委員會")+"統計室，一份送本局會計室，一份自存。",wb,row,(short)0,(short)0,(short)0);
				insertCell_tail("填表說明：本表編製一式三份，一份送行政院農業委員會統計室，一份送本局會計室，一份自存。",wb,row,(short)0,(short)0,(short)0);
				System.out.println("明細表end");
			}

	        HSSFFooter footer = sheet.getFooter();
	        footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));

	        FileOutputStream fout=new FileOutputStream(reportDir + System.getProperty("file.separator")+filename);
	        wb.write(fout);
	        //儲存
	        fout.close();
	        System.out.println("儲存完成");
		}catch(Exception e){
			System.out.println("createRpt Error:"+e+e.getMessage());
		}
		return errMsg;
	}

	public static void insertCell(List dbData,boolean getstate,int index,String Item,HSSFWorkbook wb,HSSFRow row,short j,
	                              short bg,
							      short bordertop,short borderbottom,short borderleft,short borderright)
	{
		try{
		String insertValue="";
  		if(getstate) insertValue= (((DataObject)dbData.get(index)).getValue(Item)).toString();
  		else         insertValue= Item;
	    HSSFCell cell=(row.getCell(j)==null)? row.createCell(j) : row.getCell(j);
	    HSSFCellStyle cs1 = cell.getCellStyle();//會套用原本excel所設定的格式

	    cs1.setBorderTop(bordertop);
	    cs1.setBorderBottom(borderbottom);
	    cs1.setBorderLeft(borderleft);
	    cs1.setBorderRight(borderright);

	    cell.setCellStyle(cs1);
	    cell.setEncoding( HSSFCell.ENCODING_UTF_16 );	    
	    //double value=0;
	    if(!insertValue.equals("0") && !insertValue.equals("") && insertValue != null){
	    	if(insertValue.indexOf("信用部") != -1){
	    		cell.setCellValue(insertValue);
	    	}else{
	    	   cell.setCellValue(Utility.setCommaFormat(insertValue));
	    	}
	    }
		}catch(Exception e){
			System.out.println("insertCell Error:"+e+e.getMessage());
		}
	}

	public static void insertCell_detail(List dbData,boolean getstate,int index,String Item,HSSFWorkbook wb,HSSFRow row,short j,
            short bg,
		      short bordertop,short borderbottom,short borderleft,short borderright)
	{
		try{
			String insertValue="";
			if(getstate) insertValue= (((DataObject)dbData.get(index)).getValue(Item)).toString();
			else         insertValue= Item;
			
			HSSFCell cell=(row.getCell(j)==null)? row.createCell(j) : row.getCell(j);
			HSSFCellStyle cs1 = wb.createCellStyle();

			cs1.setBorderTop(bordertop);
			cs1.setBorderBottom(borderbottom);
			cs1.setBorderLeft(borderleft);
			cs1.setBorderRight(borderright);
			cs1.setAlignment(HSSFCellStyle.ALIGN_RIGHT);

			cell.setCellStyle(cs1);
			cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
			
			if(!insertValue.equals("0") && !insertValue.equals("") && insertValue != null){
				if(insertValue.indexOf("信用部") != -1){
					cell.setCellValue(insertValue);
				}else{				    
					cell.setCellValue(Utility.setCommaFormat(insertValue));
				}				
			}

		}catch(Exception e){
			System.out.println("insertCell Error:"+e+e.getMessage());
		}
	}
	public static void insertCell_tail(String Item,HSSFWorkbook wb,HSSFRow row,short j,short top,short down)
	{
		try{
			//System.out.println(j+"Item="+Item);
			HSSFCell cell=(row.getCell(j)==null)? row.createCell(j) : row.getCell(j);
			HSSFCellStyle cs1 = wb.createCellStyle();
			//HSSFCellStyle cs1 = cell.getCellStyle();//會套用原本excel所設定的格式
			HSSFFont font = wb.createFont();   
	        font.setFontHeightInPoints((short) 10);   
	        font.setFontName("標楷體");
	        cs1.setFont(font);  
			cs1.setBorderTop(top);
		    cs1.setBorderBottom(down);
		    cs1.setBorderLeft((short)0);
		    cs1.setBorderRight((short)0);

		    cs1.setAlignment(HSSFCellStyle.ALIGN_GENERAL);
		    cell.setCellStyle(cs1);
			cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
			cell.setCellValue(Item);

		}catch(Exception e){
			System.out.println("insertCell_tail Error:"+e+e.getMessage());
		}
	}
	public static void insertBlank(int maxlength,HSSFWorkbook wb,short top,short down,HSSFRow row){
	    for(int k=0;k<=maxlength;k++){
	        //System.out.println("k="+k);
   	        insertCell(null,false,0," ",wb,row,(short)k, (short)64,top,(short)0,(short)0,(short)0);
        }//end of insert 空值表格
	}
}
