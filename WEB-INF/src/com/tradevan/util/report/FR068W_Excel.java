/* 
   104.01.08 add by 2968  
   104.01.26 fix 設定為100%,橫印 by 2295
   104.01.26 add 調整聯絡人資料顯示為A12申報者資料 by 2295  
   104.02.04 fix 調整備抵呆帳-本月提列金額(A01.520800),若為1月份資料時,則不用扣掉上月份資料 by 2295
   104.04.27 add A12_extra中文存保用檔案下載 by 2295
*/
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*;
import java.io.*;
import java.math.BigDecimal;
import java.util.*;

import com.tradevan.util.DownLoad;
import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class FR068W_Excel {
    public static String createRpt(String S_YEAR,String S_MONTH,String bank_code,String unit,String bank_type,String muser_id){
    	String bank_type_name ="農漁會";
    	if(bank_type.equals("6")){
    		bank_type_name ="農會"; 
    	}else if(bank_type.equals("7")){
    		bank_type_name ="漁會"; 
    	}
    	
    	String errMsg = "";
		StringBuffer sql = new StringBuffer () ;
		ArrayList paramList = new ArrayList() ;
		List dbData = null;
		List dbData_other = null;
		String bank_name="全體"+bank_type_name+"信用部";
		String unit_name=Utility.getUnitName(unit);
		FileInputStream finput = null;
		String u_year = "100" ;
		if(S_YEAR!=null && Integer.parseInt(S_YEAR) <= 99) {
			u_year = "99" ;
		}
		String lastu_year = "100";
		String lastYear = S_YEAR;
		String lastMonth = String.valueOf(Integer.parseInt(S_MONTH)-1);
		if("0".equals(lastMonth)){
			lastMonth = "12";
			lastYear = String.valueOf(Integer.parseInt(S_YEAR)-1);
			if(lastYear!=null && Integer.parseInt(lastYear) <= 99) {
				lastu_year = "99" ;
			}
		}
		if(lastMonth.length()<2)lastMonth="0"+lastMonth;
		
		Utility.printLogTime("FR068W_Excel begin time");	
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
    		
            finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+"逾期放款及轉銷呆帳及存款準備率降低所增盈餘月報表.xls" );
			
			System.out.println(xlsDir + System.getProperty("file.separator")+"逾期放款及轉銷呆帳及存款準備率降低所增盈餘月報表.xls");
	  	    //設定FileINputStream讀取Excel檔
	  		POIFSFileSystem fs = new POIFSFileSystem( finput );
	  		HSSFWorkbook wb = new HSSFWorkbook(fs);
	  		HSSFSheet sheet = wb.getSheetAt(0);//讀取第一個工作表，宣告其為sheet 
	  		HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
	        //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
	        //sheet.setAutobreaks(true); //自動分頁
			
	        //設定頁面符合列印大小
	        sheet.setAutobreaks( false );
	        ps.setScale( ( short )100 ); //列印縮放百分比
	        ps.setLandscape(true);//設定橫印
	        HSSFFooter footer = sheet.getFooter();
	        ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
	  		//wb.setSheetName(0,"test");
	  		finput.close();
	  		
	  		HSSFRow row=null;//宣告一列 
	  		HSSFCell cell=null;//宣告一個儲存格  		
	  		sql.append(" select bank_no,bank_name from bn01 where bank_no= ? and m_year =? ") ;
	  		paramList.add(bank_code) ;
            paramList.add(u_year) ;
	  		if(!"".equals(bank_type)){
	  			sql.append(" and bank_type= ? ") ;
	  			paramList.add(bank_type) ;
	  		}
            
            dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "bank_name");  
            if(dbData != null && dbData.size()!=0 ){
                bank_name=(String)((DataObject)dbData.get(0)).getValue("bank_name");
            }
	  		
            sql.setLength(0) ;
            paramList.clear() ;
            sql.append("select a01.bank_code,");
            sql.append("  round(a01.field_over /?,0)  as field_over ,");//--逾期放款-月底餘額
            sql.append("  round((a01.field_over-a01_lastmonth.field_over)/ ?,0) as field_over_diff,");//--逾期放款-較上月底增減金額,
            sql.append("  round(a01.field_credit /?,0)  as field_credit,");//--放款總額
            sql.append("  decode(a01.field_CREDIT,0,0,round(a01.field_OVER /  a01.field_CREDIT *100 ,2))  as   field_OVER_RATE,");//--逾放比率
            sql.append("  round(a01.field_BACKUP/?,0) as field_BACKUP,");//--備低呆帳-月底餘額
            if(Integer.parseInt(S_MONTH) != 1){
            sql.append("  round((a01.field_520800-a01_lastmonth.field_520800)/?,0) as field_520800_diff,");//--備抵呆帳-本月提列金額
            }else{
            sql.append("  round(a01.field_520800/?,0) as field_520800_diff,");//--備抵呆帳-本月提列金額    
            }
            sql.append("  round(a12.baddebt_amt /?,0)  as baddebt_amt ,");// --本月轉銷呆帳金額-減少備抵呆帳金額B1
            sql.append("  round(a12.loss_amt /?,0)  as loss_amt ,");// --本月轉銷呆帳金額-直接認列損失金額B2
            sql.append("  round((a12_sum.baddebt_amt+a12_sum.loss_amt)/ ?,0) as baddebt_loss_sum, ");//--C1累計轉銷金額B1+B2(當年度)
            sql.append("  round(a12.profit_amt/?,0) as profit_amt ,");// --存款準備率降低所增加盈餘-本月增加金額 B3
            sql.append("  round(a12_sum.profit_amt/?,0)  as profit_sum ");//--C2存款準備率降低所增加盈餘-累計增加金額(當年度)
            paramList.add(unit) ;
            paramList.add(unit) ;
            paramList.add(unit) ;
            paramList.add(unit) ;
            paramList.add(unit) ;
            paramList.add(unit) ;
            paramList.add(unit) ;
            paramList.add(unit) ;
            paramList.add(unit) ;
            paramList.add(unit) ;
            sql.append("from ");
            sql.append("(select ");
            if("ALL".equals(bank_code)){
            	sql.append(" 'ALL' as bank_code,");
            }else{
            	sql.append(" bank_code,");
            }
            sql.append("round(sum(decode(a01.acc_code,'990000',amt,0)) /1,0)  as field_OVER,");//--逾放金額
            sql.append("round(sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /1,0) as  field_CREDIT,");//--放款總額
            sql.append("round(sum(decode(a01.acc_code, '120800',amt,'150300',amt,0)) /1,0) as  field_BACKUP,");//--備低呆帳
            sql.append("round(sum(decode(a01.acc_code, '520800',amt,0)) /1,0) as  field_520800 ");//--呆帳
            sql.append("from a01 left join (select * from bn01 where m_year=?)bn01 on a01.bank_code=bn01.bank_no ");
            sql.append("where a01.m_year=? ");
            sql.append("and m_month=? ");
            paramList.add(u_year) ;
            paramList.add(S_YEAR) ;
            paramList.add(S_MONTH) ;
            if("ALL".equals(bank_code)){
            	sql.append("  and bank_type=? ");
            	paramList.add(bank_type) ;
            }else{
            	sql.append(" and bank_code=? ");
            	sql.append("group by bank_code ");
 	            paramList.add(bank_code) ;
            }
            sql.append(")a01,");
            sql.append("(select ");
            if("ALL".equals(bank_code)){
            	sql.append(" 'ALL' as bank_code,");
            }else{
            	sql.append(" bank_code,");
            }
            sql.append("round(sum(decode(a01.acc_code,'990000',amt,0)) /1,0)  as field_OVER,");//--逾放金額
            sql.append("round(sum(decode(a01.acc_code, '120800',amt,'150300',amt,0)) /1,0) as  field_BACKUP,");//--備低呆帳
            sql.append("round(sum(decode(a01.acc_code, '520800',amt,0)) /1,0) as  field_520800 ");//--呆帳
            sql.append("from a01  left join (select * from bn01 where m_year=?)bn01 on a01.bank_code=bn01.bank_no ");
            sql.append("where  a01.m_year=? ");
            sql.append("and m_month=? ");
            paramList.add(lastu_year) ;
            paramList.add(lastYear) ;
            paramList.add(lastMonth) ;
            if("ALL".equals(bank_code)){
            	sql.append("  and bank_type=? ");
            	paramList.add(bank_type) ;
            }else{
            	sql.append(" and bank_code=? ");
            	sql.append("group by bank_code ");
 	            paramList.add(bank_code) ;
            }
            sql.append(")a01_lastmonth,");
            sql.append("(select ");
            if("ALL".equals(bank_code)){
            	sql.append(" 'ALL' as bank_code,");
            }else{
            	sql.append(" bank_code,");
            }
            sql.append("       round(sum(baddebt_amt) /1,0) as baddebt_amt, ");//--本月轉銷呆帳金額-減少備抵呆帳金額B1
            sql.append("       round(sum(loss_amt) /1,0) as loss_amt, ");//--本月轉銷呆帳金額-直接認列損失金額B2
            sql.append("       round(sum(profit_amt) /1,0) as profit_amt ");//--存款準備率降低所增加盈餘-本月增加金額 B3                                  
            sql.append(" from a12  left join (select * from bn01 where  m_year=?)bn01 on a12.bank_code=bn01.bank_no ");
            sql.append("where a12.m_year=? ");
            sql.append("and m_month=? ");
            paramList.add(u_year) ;
            paramList.add(S_YEAR) ;
            paramList.add(S_MONTH) ;
            if("ALL".equals(bank_code)){
            	sql.append("  and bank_type=? ");
            	paramList.add(bank_type) ;
            }else{
            	sql.append(" and bank_code=? ");
            	sql.append("group by bank_code ");
 	            paramList.add(bank_code) ;
            }
            sql.append(" )a12  left join (select * from bn01 where m_year=?)bn01 on a12.bank_code=bn01.bank_no,");
            sql.append(" (select ");
            paramList.add(u_year) ;
            if("ALL".equals(bank_code)){
            	sql.append(" 'ALL' as bank_code,");
            }else{
            	sql.append(" bank_code,");
            }
            sql.append("       round(sum(baddebt_amt) /1,0) as baddebt_amt,");//--本月轉銷呆帳金額-減少備抵呆帳金額B1
            sql.append("       round(sum(loss_amt) /1,0) as loss_amt,");//--本月轉銷呆帳金額-直接認列損失金額B2
            sql.append("       round(sum(profit_amt) /1,0) as profit_amt ");//--存款準備率降低所增加盈餘-本月增加金額 B3                        
            sql.append("  from a12   left join (select * from bn01 where m_year=?)bn01 on a12.bank_code=bn01.bank_no ");
            sql.append("  where (to_char(a12.m_year * 100 + m_month) >= ? and to_char(a12.m_year * 100 + m_month) <= ?) ");//--年度資料
            paramList.add(u_year) ;
            paramList.add(S_YEAR+"01") ;
            paramList.add(S_YEAR+S_MONTH) ;
            if("ALL".equals(bank_code)){
            	sql.append("  and bank_type=? ");
            	paramList.add(bank_type) ;
            }else{
            	sql.append(" and bank_code=? ");
            	sql.append("group by bank_code ");
 	            paramList.add(bank_code) ;
            }
            sql.append(" )a12_sum ");//--當年度累計      
            sql.append("where a01.bank_code=a01_lastmonth.bank_code ");
            sql.append("and a01.bank_code=a12.bank_code ");
            sql.append("and a01.bank_code=a12_sum.bank_code ");
            
            dbData =DBManager.QueryDB_SQLParam(sql.toString(),paramList, "field_over,field_over_diff,field_credit,field_over_rate,field_backup,"
                                                                        +"field_520800_diff,baddebt_amt,loss_amt,baddebt_loss_sum,profit_amt,profit_sum");	 
        	System.out.println("dbData.size()="+dbData.size());
        	
        	sql.setLength(0) ;
            paramList.clear() ;
        	sql.append(" select muser_name || m_telno as muser_data,");//--聯絡人姓名及電話
            sql.append(" director_name ");//--主管姓名            
            sql.append(" from wtt01 ");
            sql.append(" left join muser_data on wtt01.muser_id=muser_data.muser_id ");
            sql.append(" left join (select * from bn01 where m_year=?)bn01 on wtt01.tbank_no = bn01.bank_no ");
            if("ALL".equals(bank_code)){
                sql.append(" where wtt01.muser_id=? ");
                paramList.add(u_year) ;
                paramList.add(muser_id) ;
            }else{//調整聯絡人資料顯示為A12申報者資料 by 2295            
                sql.append(" ,(select * from wml01 where report_no='A12' and m_year=? and m_month=? and bank_code=?)wml01");
                sql.append(" where wtt01.muser_id=wml01.add_user ");
                paramList.add(u_year) ;
                paramList.add(S_YEAR) ;
                paramList.add(S_MONTH) ;
                paramList.add(bank_code) ;
            }
            
           
            dbData_other =DBManager.QueryDB_SQLParam(sql.toString(),paramList, "muser_data,director_name,bank_name");	 
            System.out.println("dbData_other.size()="+dbData_other.size());
            
            row=sheet.getRow(0);
	  		cell=row.getCell((short)0);	       	
	  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	   	    cell.setCellValue("基層金融機構逾期放款及轉銷呆帳及存款準備率降低所增盈餘月報表");
  		  	row=sheet.getRow(2);
  		  	cell=row.getCell((short)6);	       	
  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
  		  	cell.setCellValue("單位：新台幣"+unit_name+"、％");
  		  	row=sheet.getRow(2);
  		  	cell=row.getCell((short)0);	       	
  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16); 
  		  	cell.setCellValue("金融機構名稱："+bank_name);
  		 
  	        if(dbData.size() == 0){	
		  		row=sheet.getRow(1);
		  		cell=row.getCell((short)0);	       	
		  		//設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
		  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
   	   	       	cell.setCellValue("基準日："+S_YEAR +"年" +S_MONTH +"月底無資料存在");
	  	    }else{
	  	      
	  		  	row=sheet.getRow(1);
	  		  	cell=row.getCell((short)0);	       	
	  		  	//設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16); 
	  		  	cell.setCellValue("基準日："+S_YEAR +"年" +S_MONTH +"月底");
	  		  	row=sheet.getRow(5);
	  		  	cell=row.getCell((short)2);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		  	cell.setCellValue(Utility.setCommaFormat((((DataObject) dbData.get(0)).getValue("field_over")).toString()));
	  		  	row=sheet.getRow(5);
	  		  	cell=row.getCell((short)3);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		  	cell.setCellValue(Utility.setCommaFormat((((DataObject) dbData.get(0)).getValue("field_over_diff")).toString()));
	  		  	row=sheet.getRow(5);
	  		  	cell=row.getCell((short)4);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellValue(Utility.setCommaFormat((((DataObject) dbData.get(0)).getValue("field_credit")).toString()));
	  		  	row=sheet.getRow(5);
	  		  	cell=row.getCell((short)5);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		  	cell.setCellValue(Utility.setCommaFormat((((DataObject) dbData.get(0)).getValue("field_over_rate")).toString()));
	  		  	row=sheet.getRow(5);
	  		  	cell=row.getCell((short)6);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellValue(Utility.setCommaFormat((((DataObject) dbData.get(0)).getValue("field_backup")).toString()));
	  		    row=sheet.getRow(5);
	  		  	cell=row.getCell((short)7);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellValue(Utility.setCommaFormat((((DataObject) dbData.get(0)).getValue("field_520800_diff")).toString()));
	  		    row=sheet.getRow(10);
	  		  	cell=row.getCell((short)2);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellValue(Utility.setCommaFormat((((DataObject) dbData.get(0)).getValue("baddebt_amt")).toString()));
	  		    row=sheet.getRow(10);
	  		  	cell=row.getCell((short)3);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellValue(Utility.setCommaFormat((((DataObject) dbData.get(0)).getValue("loss_amt")).toString()));
	  		  	row=sheet.getRow(10);
	  		  	cell=row.getCell((short)4);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellValue(Utility.setCommaFormat((((DataObject) dbData.get(0)).getValue("baddebt_loss_sum")).toString()));
	  		  	row=sheet.getRow(10);
	  		  	cell=row.getCell((short)5);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellValue(Utility.setCommaFormat((((DataObject) dbData.get(0)).getValue("profit_amt")).toString()));
	  		  	row=sheet.getRow(10);
	  		  	cell=row.getCell((short)6);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellValue(Utility.setCommaFormat((((DataObject) dbData.get(0)).getValue("profit_sum")).toString()));
	  		 
	  	    }
  	       
	  	      if(dbData_other.size() > 0){
	  	    	  row=sheet.getRow(17);
		  		  cell=row.getCell((short)0);	       	
		  		  cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	 	       	  cell.setCellValue("聯絡人姓名及電話："+Utility.getTrimString(((DataObject) dbData_other.get(0)).getValue("muser_data")));	       	
		   	      row=sheet.getRow(17);
		  		  cell=row.getCell((short)4);	       	
		  		  cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		       	  cell.setCellValue("主管姓名："+Utility.getTrimString(((DataObject) dbData_other.get(0)).getValue("director_name")));
	          }
	  		
  	        //設定涷結欄位
            //sheet.createFreezePane(0,1,0,1);
            footer.setCenter( "Page:" + HSSFFooter.page() + " of " +
                             HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));			
	       
  	        
  	        FileOutputStream fout=new FileOutputStream(reportDir + System.getProperty("file.separator")+"逾期放款及轉銷呆帳及存款準備率降低所增盈餘月報表.xls");
	        wb.write(fout);
	        //儲存 
	        fout.close();
	        Utility.printLogTime("FR068W_Excel end time");	
		}catch(Exception e){
			System.out.println("createRpt Error:"+e+e.getMessage());
		}
		return errMsg;
	}	
    
    //104.04.27 add A12_extra中央存保格式檔案下載改用回傳List 
    public static List printData(String S_YEAR, String S_MONTH, String selectBank_no, String bank_type,String unit) {                                  
        String errMsg = "";
        List dbData = null;
        StringBuffer sql = new StringBuffer () ;
        String printResult="";
        List return_List = new LinkedList();
        String cd01_table = "";
        String wlx01_m_year = "";
        List paramList = new ArrayList();
        
        
        try {
            //100.02.22 add 查詢年度100年以前.縣市別不同===============================
            cd01_table = (Integer.parseInt(S_YEAR) < 100)?"cd01_99":""; 
            wlx01_m_year = (Integer.parseInt(S_YEAR) < 100)?"99":"100";
            //=====================================================================   
            String u_year = "100" ;
            if(S_YEAR!=null && Integer.parseInt(S_YEAR) <= 99) {
                u_year = "99" ;
            }
            String lastu_year = "100";
            String lastYear = S_YEAR;
            String lastMonth = String.valueOf(Integer.parseInt(S_MONTH)-1);
            if("0".equals(lastMonth)){
                lastMonth = "12";
                lastYear = String.valueOf(Integer.parseInt(S_YEAR)-1);
                if(lastYear!=null && Integer.parseInt(lastYear) <= 99) {
                    lastu_year = "99" ;
                }
            }
            if(lastMonth.length()<2)lastMonth="0"+lastMonth;
            sql.setLength(0) ;
            paramList.clear() ;
            sql.append("select a01.bank_code,");
            sql.append("  round(a01.field_over /?,0)  as field_over ,");//--逾期放款-月底餘額
            sql.append("  round((a01.field_over-a01_lastmonth.field_over)/ ?,0) as field_over_diff,");//--逾期放款-較上月底增減金額,
            sql.append("  round(a01.field_credit /?,0)  as field_credit,");//--放款總額
            sql.append("  decode(a01.field_CREDIT,0,0,round(a01.field_OVER /  a01.field_CREDIT *100 ,2))  as   field_OVER_RATE,");//--逾放比率
            sql.append("  round(a01.field_BACKUP/?,0) as field_BACKUP,");//--備低呆帳-月底餘額
            if(Integer.parseInt(S_MONTH) != 1){
            sql.append("  round((a01.field_520800-a01_lastmonth.field_520800)/?,0) as field_520800_diff,");//--備抵呆帳-本月提列金額
            }else{
            sql.append("  round(a01.field_520800/?,0) as field_520800_diff,");//--備抵呆帳-本月提列金額    
            }
            sql.append("  round(a12.baddebt_amt /?,0)  as baddebt_amt ,");// --本月轉銷呆帳金額-減少備抵呆帳金額B1
            sql.append("  round(a12.loss_amt /?,0)  as loss_amt ,");// --本月轉銷呆帳金額-直接認列損失金額B2
            sql.append("  round((a12_sum.baddebt_amt+a12_sum.loss_amt)/ ?,0) as baddebt_loss_sum, ");//--C1累計轉銷金額B1+B2(當年度)
            sql.append("  round(a12.profit_amt/?,0) as profit_amt ,");// --存款準備率降低所增加盈餘-本月增加金額 B3
            sql.append("  round(a12_sum.profit_amt/?,0)  as profit_sum ");//--C2存款準備率降低所增加盈餘-累計增加金額(當年度)
            paramList.add(unit) ;
            paramList.add(unit) ;
            paramList.add(unit) ;
            paramList.add(unit) ;
            paramList.add(unit) ;
            paramList.add(unit) ;
            paramList.add(unit) ;
            paramList.add(unit) ;
            paramList.add(unit) ;
            paramList.add(unit) ;
            sql.append("from ");
            sql.append("(select bank_code,");
            sql.append("round(sum(decode(a01.acc_code,'990000',amt,0)) /1,0)  as field_OVER,");//--逾放金額
            sql.append("round(sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /1,0) as  field_CREDIT,");//--放款總額
            sql.append("round(sum(decode(a01.acc_code, '120800',amt,'150300',amt,0)) /1,0) as  field_BACKUP,");//--備低呆帳
            sql.append("round(sum(decode(a01.acc_code, '520800',amt,0)) /1,0) as  field_520800 ");//--呆帳
            sql.append("from a01 left join (select * from bn01 where m_year=?)bn01 on a01.bank_code=bn01.bank_no ");
            sql.append("where a01.m_year=? ");
            sql.append("and m_month=? ");
            paramList.add(u_year) ;
            paramList.add(S_YEAR) ;
            paramList.add(S_MONTH) ;
            sql.append(" and "+selectBank_no);
            sql.append("group by bank_code ");   
            sql.append(")a01,");
            sql.append("(select bank_code,");
            sql.append("round(sum(decode(a01.acc_code,'990000',amt,0)) /1,0)  as field_OVER,");//--逾放金額
            sql.append("round(sum(decode(a01.acc_code, '120800',amt,'150300',amt,0)) /1,0) as  field_BACKUP,");//--備低呆帳
            sql.append("round(sum(decode(a01.acc_code, '520800',amt,0)) /1,0) as  field_520800 ");//--呆帳
            sql.append("from a01  left join (select * from bn01 where m_year=?)bn01 on a01.bank_code=bn01.bank_no ");
            sql.append("where  a01.m_year=? ");
            sql.append("and m_month=? ");
            paramList.add(lastu_year) ;
            paramList.add(lastYear) ;
            paramList.add(lastMonth) ;           
            sql.append(" and "+selectBank_no);
            sql.append("group by bank_code ");           
            sql.append(")a01_lastmonth,");
            sql.append("(select bank_code,");
            sql.append("       round(sum(baddebt_amt) /1,0) as baddebt_amt, ");//--本月轉銷呆帳金額-減少備抵呆帳金額B1
            sql.append("       round(sum(loss_amt) /1,0) as loss_amt, ");//--本月轉銷呆帳金額-直接認列損失金額B2
            sql.append("       round(sum(profit_amt) /1,0) as profit_amt ");//--存款準備率降低所增加盈餘-本月增加金額 B3                                  
            sql.append(" from a12  left join (select * from bn01 where  m_year=?)bn01 on a12.bank_code=bn01.bank_no ");
            sql.append("where a12.m_year=? ");
            sql.append("and m_month=? ");
            paramList.add(u_year) ;
            paramList.add(S_YEAR) ;
            paramList.add(S_MONTH) ;
            sql.append(" and "+selectBank_no);
            sql.append("group by bank_code ");
              
            sql.append(" )a12  left join (select * from bn01 where m_year=?)bn01 on a12.bank_code=bn01.bank_no,");
            paramList.add(u_year) ;
            sql.append(" (select bank_code,");            
            sql.append("       round(sum(baddebt_amt) /1,0) as baddebt_amt,");//--本月轉銷呆帳金額-減少備抵呆帳金額B1
            sql.append("       round(sum(loss_amt) /1,0) as loss_amt,");//--本月轉銷呆帳金額-直接認列損失金額B2
            sql.append("       round(sum(profit_amt) /1,0) as profit_amt ");//--存款準備率降低所增加盈餘-本月增加金額 B3                        
            sql.append("  from a12   left join (select * from bn01 where m_year=?)bn01 on a12.bank_code=bn01.bank_no ");
            sql.append("  where (to_char(a12.m_year * 100 + m_month) >= ? and to_char(a12.m_year * 100 + m_month) <= ?) ");//--年度資料
            paramList.add(u_year) ;
            paramList.add(S_YEAR+"01") ;
            paramList.add(S_YEAR+S_MONTH) ;
            sql.append(" and "+selectBank_no);
            sql.append("group by bank_code ");
            sql.append(" )a12_sum ");//--當年度累計      
            sql.append("where a01.bank_code=a01_lastmonth.bank_code ");
            sql.append("and a01.bank_code=a12.bank_code ");
            sql.append("and a01.bank_code=a12_sum.bank_code ");
            
            dbData =DBManager.QueryDB_SQLParam(sql.toString(),paramList, "field_over,field_over_diff,field_credit,field_over_rate,field_backup,"
                                                                        +"field_520800_diff,baddebt_amt,loss_amt,baddebt_loss_sum,profit_amt,profit_sum");   
            System.out.println("dbData.size()="+dbData.size()); 
            DataObject bean = null;
            double  field_over_rate = 0.0;
            String tmp_over_rate = "";
            for (int k = 0; k < dbData.size(); k++) {       
                bean =(DataObject) dbData.get(k);
                //System.out.print("field_over_rate1="+(bean.getValue("field_over_rate")).toString());
                field_over_rate = Double.parseDouble((bean.getValue("field_over_rate")).toString());
                field_over_rate = field_over_rate * 100;
                tmp_over_rate = String.valueOf(field_over_rate);
                if(tmp_over_rate.indexOf(".") != -1){
                    tmp_over_rate = tmp_over_rate.substring(0,tmp_over_rate.indexOf("."));
                    //System.out.print("field_over_rate2="+tmp_over_rate);
                }
                 
                return_List.add(DownLoad.fillStuff(S_YEAR, "L", "0", 3)//年                                 
                        +  DownLoad.fillStuff(S_MONTH, "L", "0", 2)//月
                        +  DownLoad.fillStuff((String)bean.getValue("bank_code"), "R", "0", 7)//機構代號                        
                        +  DownLoad.fillStuff((bean.getValue("field_over")).toString(), "L", "0", 0, 14)//逾期放款-月底餘額
                        +  DownLoad.fillStuff((bean.getValue("field_over_diff")).toString(), "L", "0", 0, 14)//逾期放款較上月底增減金額
                        +  DownLoad.fillStuff((bean.getValue("field_credit")).toString(), "L", "0", 0, 14)//放款總額
                        +  DownLoad.fillStuff(tmp_over_rate, "L", "0", 0, 14)//逾放比率
                        +  DownLoad.fillStuff((bean.getValue("field_backup")).toString(), "L", "0", 0, 14)//備抵呆帳-月底餘額
                        +  DownLoad.fillStuff((bean.getValue("field_520800_diff")).toString(), "L", "0", 0, 14)//備抵呆帳-本月提列金額
                        +  DownLoad.fillStuff((bean.getValue("baddebt_amt")).toString(), "L", "0", 0, 14)//本月轉銷呆帳-減少備抵呆帳金額
                        +  DownLoad.fillStuff((bean.getValue("loss_amt")).toString(), "L", "0", 0, 14)//本月轉銷呆帳-直接認列損失金額
                        +  DownLoad.fillStuff((bean.getValue("baddebt_loss_sum")).toString(), "L", "0", 0, 14)//累計轉銷金額
                        +  DownLoad.fillStuff((bean.getValue("profit_amt")).toString(), "L", "0", 0, 14)//存款準備率降低所增加盈餘-本月增加金額
                        +  DownLoad.fillStuff((bean.getValue("profit_sum")).toString(), "L", "0", 0, 14)//存款準備率降低所增加盈餘-累計增加金額
                        ); 
            }
        } catch (Exception e) {
            System.out.println("printData Error:" + e + e.getMessage());
        }
        //return printResult;
        return return_List;
    }	 
}
