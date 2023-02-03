/*
	94.03.11 fix 金額資料為零是不輸出0，改為輸出空白及欄位資料右靠處理	
	94.08.12 fix 增加頁尾 by 2295
	94.11.17 add 增加全体農會資料負債表/金額單位 by 2295
	94.12.06 add 逾期放款金額/逾放比率/存放比率 by 2295
	94.12.12 fix 更改公式 by 2295
              1.round(sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /1,0) as  field_CREDIT_FISH,//漁會
              2.round(sum(decode(a01.acc_code,'240305',amt,'251300',amt,'310800',amt, 0)) /1,0) as fieldI_XB2,//農會
                round(sum(decode(a01.acc_code,'240205',amt, '310800',amt,0)) /1,0) as fieldI_XB2_FISH,//漁會
    95.01.24 fix 更改公式 by 2295
 			  1. round(sum(decode(a01.acc_code,'310800',amt,0)) /1,0)     as fieldI_XF3,
 			  2.decode(sign(a01.fieldI_XF1 - a01.fieldI_XF3 - a01.fieldI_XF2),-1,0,
                     (a01.fieldI_XF1 - a01.fieldI_XF3 - a01.fieldI_XF2))
    95.01.26 fix 全体農會加上只抓bank_type='6' by 2295
    95.02.10 fix 全体農會信用部/bank_name會變成亂碼的問題 by 2295
    95.10.03 增加檢核結果與最後異動日期  by 2495    
    96.12.19 add 97/01以後,套用新表格(增加/異動科目代號) by 2295
    99.04.27 fix 縣市合併問題 and sql injection by 2808
   100.02.16 fix 無法顯示逾期放款金額/逾放比率/存放比率 by 2295 
   102.11.19 add 原QueryDB改套用QueryDB_SQLParam by 2295    
   103.05.28 add 全体總表,若有檢核有誤資料,顯示農漁會信用部名稱  by 2295
*/
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.poifs.filesystem.*;
import java.io.*;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class FR003W_Excel {
    public static String createRpt(String S_YEAR,String S_MONTH,String bank_code,String BANK_NAME,String unit){
		String errMsg = "";
		List dbData = null;
		List dbData_other = null;
		String sqlCmd = "", sqlCmd_other = "";
		StringBuffer sql = new StringBuffer () ;
		ArrayList paramList = new ArrayList() ;
		Properties A01Data = new Properties();		
		String acc_code = "";
		String amt = "";
		String unit_name="";
		FileInputStream finput = null;
		String ncacno="ncacno";
		int rowNum=0;
		String u_year = "100" ;
		DataObject bean = null;
		if(S_YEAR!=null && Integer.parseInt(S_YEAR) <= 99) {
			u_year = "99" ;
		}
		Utility.printLogTime("FR003W_Excel begin time");	
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
    		//95.02.10 fix bank_name變成亂碼
    		sql.setLength(0) ;
    		sql.append(" select bank_no,bank_name from bn01 where bank_type= ? ");
    		sql.append(" and bank_no = ? ");
    		sql.append(" and m_year = ? ");
    		paramList.add("6") ;
    		paramList.add(bank_code);
    		paramList.add(u_year) ;
    		
            dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "");	
            	//DBManager.QueryDB("select bank_no,bank_name from bn01 where bank_type='6' and bank_no='"+bank_code+"'",""); 
            if(dbData != null && dbData.size()!=0 ){
                BANK_NAME=(String)((DataObject)dbData.get(0)).getValue("bank_name");
            }
            paramList.clear() ;
            //96.12.19 add 97/01以後,套用新表格(增加/異動科目代號) 
			if(Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH) >= 9701){
		    	ncacno = "ncacno_rule";
				finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+"農業信用部資產負債表_9701.xls" );
				System.out.println("農業信用部資產負債表_9701.xls");
		    }else{
		    	finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+"農業信用部資產負債表.xls" );
				System.out.println("農業信用部資產負債表.xls");
		    }
			
			
	  	    //設定FileINputStream讀取Excel檔
	  		POIFSFileSystem fs = new POIFSFileSystem( finput );
	  		HSSFWorkbook wb = new HSSFWorkbook(fs);
	  		HSSFSheet sheet = wb.getSheetAt(0);//讀取第一個工作表，宣告其為sheet 
	  		HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
	        //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
	        //sheet.setAutobreaks(true); //自動分頁
			
	        //設定頁面符合列印大小
	        sheet.setAutobreaks( false );
	        ps.setScale( ( short )70 ); //列印縮放百分比
	        HSSFFooter footer = sheet.getFooter();
	        ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
	  		//wb.setSheetName(0,"test");
	  		finput.close();
	  		
	  		HSSFRow row=null;//宣告一列 
	  		HSSFCell cell=null;//宣告一個儲存格  		
	  		
	  		short i=0;
	  		short y=0;  
	  		System.out.println("start sql 2 ===============================================") ;
	  		sql.setLength(0) ;
	  		sql.append("  select A01.m_year, A01.m_month, "+ncacno+".acc_range, a01.acc_code, ") ;
	  		//sqlCmd = " select A01.m_year, A01.m_month, "+ncacno+".acc_range, a01.acc_code, ";
	  		if(bank_code.equals("ALL")){//全体農會	  		    
	  		   //sqlCmd += " round(sum(amt)/"+unit+",0) as amt";
	  			sql.append(" round(sum(amt)/ ? ,0) as amt ");
	  			paramList.add(unit) ;
	  		}else{	  		   
	  		   //sqlCmd += " round(amt/"+unit+",0) as amt";
	  			sql.append(" round(amt/? ,0) as amt ");
	  			paramList.add(unit) ;
	  		}
	  		//sqlCmd += " from A01 LEFT JOIN "+ncacno+" ON A01.acc_code = "+ncacno+".acc_code";
	  		sql.append(" from A01 LEFT JOIN "+ncacno+" ON A01.acc_code = "+ncacno+".acc_code ") ;
	  		if(bank_code.equals("ALL")){//全体農會 fix 95.01.26
	  		   //sqlCmd += ",bn01 ";
	  			sql.append(",(select * from bn01 where bn01.m_year= ?  ) bn01 ");
	  			paramList.add(u_year) ;
	  		}
	  		//sqlCmd += " where A01.m_year="+S_YEAR 
			//	   + "   and A01.m_month="+String.valueOf(Integer.parseInt(S_MONTH));
	  		sql.append(" where A01.m_year= ? ");
	  		sql.append(" and A01.m_month=  ? ");
	  		paramList.add(S_YEAR) ;
	  		paramList.add(S_MONTH) ;
	  		if(!bank_code.equals("ALL")){
			    //sqlCmd += "   and A01.bank_code='"+bank_code+"'";
	  			sql.append(" and A01.bank_code=? ");
	  			paramList.add(bank_code) ;
	  		}	   
			//sqlCmd += "   and "+ncacno+".acc_div='01'";
	  		sql.append(" and "+ncacno+".acc_div= ? ") ;
	  		paramList.add("01") ;
	  		if(bank_code.equals("ALL")){//全体農會 fix 95.01.26
			   //sqlCmd += " and A01.bank_code = bn01.bank_no" 
			   //    	   + " and bn01.bank_type='6'"
			   //        + " group by A01.m_year, A01.m_month, "+ncacno+".acc_range, A01.acc_code"
	           //        +  " order by A01.m_year, A01.m_month, "+ncacno+".acc_range, A01.acc_code";
	  			sql.append("  and A01.bank_code = bn01.bank_no ") ;
	  			sql.append("  and bn01.bank_type= ?  ");
	  			sql.append("  group by A01.m_year, A01.m_month, "+ncacno+".acc_range, A01.acc_code ");
	  			sql.append("  order by A01.m_year, A01.m_month, "+ncacno+".acc_range, A01.acc_code ");
	  			paramList.add("6") ;
	  		}else{			
			    //sqlCmd += " order by "+ncacno+".acc_range";			    	
			    sql.append(" order by "+ncacno+".acc_range ");
			}
	  		dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "m_year,m_month,amt");	
	  			//DBManager.QueryDB(sqlCmd,"m_year,m_month,amt");	
	  		int dbDatasize=dbData.size();
            System.out.println("a01.size()="+dbData.size());
	  		//94.12.06 add 逾期放款金額/逾放比率/存放比率===============================================
            System.out.println("start sql 3==========================================================") ;
            sql.setLength(0) ;
            paramList.clear() ;
            sql.append(" select round(field_OVER / ?,0)    as field_OVER, ") ; //逾期放款金額            
            sql.append(" decode(a01.field_CREDIT,0,0,round(a01.field_OVER /  a01.field_CREDIT *100 ,2))  as   field_OVER_RATE,"); //逾放比率
            sql.append(" decode(a01.fieldI_Y,0,0, round((a01.fieldI_XA + "              );
            sql.append(" decode(sign(a01.fieldI_XB1 - a01.fieldI_XB2),-1,0,"            );
            sql.append("       (a01.fieldI_XB1 - a01.fieldI_XB2))        +"             );
            sql.append(" decode(sign(a01.fieldI_XC1 - a01.fieldI_XC2),-1,0,"            );
            sql.append("	   (a01.fieldI_XC1 - a01.fieldI_XC2))           +"      );
            sql.append(" decode(sign(a01.fieldI_XD1 - a01.fieldI_XD2),-1,0,"            );
            sql.append("	   (a01.fieldI_XD1 - a01.fieldI_XD2))           +"      );
            sql.append(" decode(sign(a01.fieldI_XE1 - a01.fieldI_XE2),-1,0,"            );
            sql.append("	    (a01.fieldI_XE1 - a01.fieldI_XE2))           -"     );
            sql.append(" decode(sign(a01.fieldI_XF1 - a01.fieldI_XF3 - a01.fieldI_XF2),-1,0," );//95.01.24 fix 
            sql.append("	    (a01.fieldI_XF1 - a01.fieldI_XF3 - a01.fieldI_XF2))"      );
            sql.append(" )"                                                                   );
            sql.append(" /    a01.fieldI_Y * 100,2))        as     field_DC_RATE"); //存放比率
            sql.append(" from ("                                                 );
            sql.append(" select"                                                 );
            sql.append(" round(sum(decode(a01.acc_code,'990000',amt,0)) /1,0)        as field_OVER,"   );
            sql.append(" round(sum(decode(a01.acc_code,'220000',amt,0)) /1,0)        as field_DEBIT,"  );                  
            sql.append(" round(sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /1,0) as  field_CREDIT," );                   
            sql.append(" round(sum(decode(a01.acc_code,'120101',amt,'120102',amt,"                                         );
            sql.append("				     '120200',amt,'120301',amt,"                                   );
            sql.append(" 				 '120302',amt,'120700',amt,"                                       );
            sql.append(" 				 '150200',amt,0)) /1,0) as fieldI_XA,");//農會
            sql.append(" round(sum(decode(a01.acc_code,'120401',amt,'120402',amt,0))/1,0) as fieldI_XB1,");//農會
            sql.append(" round(sum(decode(a01.acc_code,'240305',amt,'251300',amt,'310800',amt, 0)) /1,0) as fieldI_XB2,");//農會                    
            sql.append(" round(sum(decode(a01.acc_code,'120501',amt,'120502',amt,0)) /1,0) as fieldI_XC1,"                    );
            sql.append(" round(sum(decode(a01.acc_code,'240301',amt,'240302',amt,"                                            );
            sql.append(" '240303',amt,'240304',amt,0)) /1,0) as fieldI_XC2," );//農會                                           
            sql.append(" round(sum(decode(a01.acc_code,'120600',amt,0)) /1,0) as fieldI_XD1,"                                 );
            sql.append(" round(sum(decode(a01.acc_code,'240200',amt,0)) /1,0) as fieldI_XD2,"                                 );
            sql.append(" round(sum(decode(a01.acc_code,'150100',amt,0)) /1,0) as fieldI_XE1,"                                 );
            sql.append(" round(sum(decode(a01.acc_code,'250100',amt,0)) /1,0) as fieldI_XE2,"                                 );
            sql.append(" round(sum(decode(a01.acc_code,'310000',amt,'320000',amt,0)) /1,0) as fieldI_XF1,"                    );
            sql.append(" round(sum(decode(a01.acc_code,'310800',amt,0)) /1,0)     as fieldI_XF3,");//95.01.24 add 
            sql.append(" round(sum(decode(a01.acc_code,'140000',amt,0)) /1,0) as fieldI_XF2,");
            sql.append(" round((sum(decode(a01.acc_code,'220100',amt,'220200',amt,");
            sql.append("           					   '220300',amt,'220400',amt,"                        );
            sql.append("					               '220500',amt,'220600',amt,"                    );
            sql.append("				                   '220700',amt,'220800',amt,"                        );
            sql.append("				                   '220900',amt,'221000',amt,0))-"                    );
            sql.append(" round(sum(decode(a01.acc_code,'220900',amt,0))/2,0)) /1,0)   as fieldI_Y"                            );
            sql.append(" from  a01, (select * from bn01 where bn01.m_year= ? ) bn01 ");
            sql.append(" where (a01.m_year  = ? ");
            sql.append("        and    a01.m_month = ? ) ");
            paramList.add(unit) ;
            paramList.add(u_year) ;
            paramList.add(S_YEAR) ;
            paramList.add(S_MONTH) ;
	  		
	  		
	  		if (!bank_code.equals("ALL")) {
	  		    //sqlCmd_other += " and   (a01.bank_code  =  '" + bank_code + "')";
	  			sql.append(" and   (a01.bank_code  = ? ) ");
	  			paramList.add(bank_code) ;
	  		}
        	//sqlCmd_other += " and   (a01.bank_code=bn01.bank_no  and bn01.bank_type='6')"
        	//    		 + " ) a01";
	  		sql.append("  and   (a01.bank_code=bn01.bank_no  and bn01.bank_type= ? ) ");//農會
	  		sql.append(" ) a01 ");
	  		paramList.add("6") ;
        	dbData_other =DBManager.QueryDB_SQLParam(sql.toString(),paramList, "field_over,field_over_rate,field_dc_rate");	 
        		     ///DBManager.QueryDB(sqlCmd_other,"field_over,field_over_rate,field_dc_rate");
        	System.out.println("dbData_other.size()="+dbData.size());
	  		//=================================================================================================
  			for(i=0;i<dbData.size();i++){
  				bean = (DataObject) dbData.get(i);
  				acc_code = (String)bean.getValue("acc_code");
  				amt = bean.getValue("amt").toString();
	       	    A01Data.setProperty(acc_code,amt);
	       	    //System.out.println("acc_code="+acc_code+":amt="+amt);
	        }
  	        if(dbDatasize == 0){	
		  		row=sheet.getRow(0);
		  		cell=row.getCell((short)3);	       	
		  		//設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
		  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
   	   	       	cell.setCellValue(S_YEAR +"年" +S_MONTH +"月無資料存在");
	  	    }else{
	  		  	row=sheet.getRow(0);
	  		  	cell=row.getCell((short)3);	       	
	  		  	//設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);   	       	   
                if(bank_code.equals("ALL")){//全体農會//95.02.10 fix by 2295
	  		      cell.setCellValue("全體農會信用部資產負債表");
	  		    }else{
   	       	      cell.setCellValue(BANK_NAME+"資產負債表");
	  		    }  
                
	  		  	row=sheet.getRow(1);
	  		  	cell=row.getCell((short)2);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		  	
	  		  	if(bank_code.equals("ALL")){//全体農會//94.11.17 add by 2295
	  		  	   cell.setCellValue("全體");	
	  		  	}else{
   	       	       cell.setCellValue(bank_code);
	  		  	}
	  		  	
   	       	            	  		
	  		  	row=sheet.getRow(1);
	  		  	cell=row.getCell((short)3);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
   	       	    cell.setCellValue("中華民國" +S_YEAR +"年" +S_MONTH +"月底");	       	
				
   	       	    //列印單位//94.11.17 add by 2295
   	       	    row=sheet.getRow(1);
   	       	    cell=(row.getCell((short)9)==null)? row.createCell((short)9) : row.getCell((short)9);   	 		
   	 	   	 	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
   	 	   	 	
   	 	   	 	unit_name = Utility.getUnitName(unit) ;
   	 	   	 	cell.setCellValue("單位：新台幣"+unit_name+"、％");
   	 	   	 	
	  		  	//以巢狀迴圈讀取所有儲存格資料 
	  		  	System.out.println("total row ="+sheet.getLastRowNum());
	  		    //96.12.19 add 97/01以後,套用新表格(增加/異動科目代號) 
	  		  	rowNum=ncacno.equals("ncacno")?106:110;
	  		  	
	  		  	for(i=4;i<rowNum;i++){    		
	    	    	row=sheet.getRow(i);	    		
	    	    	cell=row.getCell((short)3);
	    	    	//System.out.print((int)cell.getNumericCellValue()+"=");
	    	    	amt = Utility.setCommaFormat(A01Data.getProperty(String.valueOf((int)cell.getNumericCellValue())));
	    	    	cell=row.getCell((short)4);
	    	    	cell.setEncoding( HSSFCell.ENCODING_UTF_16 );	       			       		
	    	    	if(!amt.equals("0"))cell.setCellValue(amt);	//94.03.11 add if condition by egg
	    	    	cell=row.getCell((short)8);
	    	    	//System.out.print((int)cell.getNumericCellValue()+"=");
	    	    	amt = Utility.setCommaFormat(A01Data.getProperty(String.valueOf((int)cell.getNumericCellValue())));	       		
	    	    	cell=row.getCell((short)9);
	    	    	if(!amt.equals("0"))cell.setCellValue(amt);	//94.03.11 add if condition by egg	    	    	
	  		  	}
	  		  	
	  		    //94.12.06 add 逾期放款金額/逾放比率/存放比率
                if(dbData_other != null && dbData_other.size() != 0){
                    //逾期放款金額
                    row = sheet.getRow(rowNum++);
                    cell = row.getCell((short) 9);                   
                    amt = Utility.setCommaFormat((((DataObject) dbData_other.get(0)).getValue("field_over")).toString());                    
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    if (!amt.equals("0")) cell.setCellValue(amt);
                    //逾放比率
                    row = sheet.getRow(rowNum++);
                    cell = row.getCell((short) 9);                   
                    amt = Utility.setCommaFormat((((DataObject) dbData_other.get(0)).getValue("field_over_rate")).toString());                    
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    if (!amt.equals("0")) cell.setCellValue(amt);
                    //存放比率
                    row = sheet.getRow(rowNum++);
                    cell = row.getCell((short) 9);                   
                    amt = Utility.setCommaFormat((((DataObject) dbData_other.get(0)).getValue("field_dc_rate")).toString());                    
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    if (!amt.equals("0")) cell.setCellValue(amt);
                }else{
                   System.out.println("dbData_other is null"); 
                }
                //103.05.26 add 全体總表,若有檢核有誤資料,顯示農漁會信用部名稱
                if("ALL".equals(bank_code)){
                   //103.05.26 add 若有檢核有誤資料,顯示農漁會信用部名稱
                   String wml01Error = Utility.getWML01_Error(S_YEAR,S_MONTH,"6","A01");
                   if(!"".equals(wml01Error)){
                       row=sheet.getRow(ncacno.equals("ncacno")?113:117);
                       rowNum = ncacno.equals("ncacno")?113:117;
                       cell=row.getCell((short)0);
                       cell.setEncoding(HSSFCell.ENCODING_UTF_16);   
                       HSSFFont f = wb.createFont();
                       //set font 1 to 12 point type
                       f.setFontHeightInPoints((short) 12);
                       f.setFontName("標楷體");
                       //make it single (normal) underline 
                       f.setUnderline(HSSFFont.U_SINGLE); 
                       //make it red
                       f.setColor( HSSFFont.COLOR_RED );
                       
                       HSSFCellStyle columnStyle = wb.createCellStyle(); 
                       columnStyle = HssfStyle.setStyle( columnStyle, f,
                                                new String[] {
                                                "PHL", "PVC", "F12",
                                                "WRAP"} );
                       cell.setCellStyle(columnStyle);
                       cell.setCellValue(wml01Error+"檢核有誤");
                       sheet.addMergedRegion( new Region( ( short )rowNum, ( short )0,( short )rowNum,( short ) 9 ) );      
                   }
                }else{
                   //95.10.03 增加檢核結果與最後異動日期
                   List paramList1 = new ArrayList();
                   sqlCmd = " select UPD_CODE, to_char(UPDATE_DATE,'yyyymmdd') as UPDATE_DATE"
                          + " from WML01"                
                          + " where M_YEAR=?"                 
                          + " and M_MONTH=?"
                          + " and BANK_CODE=?"
                          + " and REPORT_NO='A01'";
                   paramList1.add(S_YEAR);
                   paramList1.add(S_MONTH);
                   paramList1.add(bank_code);               
                   System.out.println("sqlCmd="+sqlCmd); 	   
		   		      
				   dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList1,"");	  
				   String UPD_CODE="";
				   String UPDATE_DATE="";
				   String M_YEAR="";   
				   String M_MONTH="";    
				   String M_DATE=""; 
				   System.out.println("dbData.size()="+dbData.size()); 	   
				   if(dbData.size()>0){
				   	   System.out.println("dbData.size()="+dbData.size()); 
				   	   UPD_CODE = (String)((DataObject)dbData.get(0)).getValue("upd_code");  
				   	   UPDATE_DATE = (String)((DataObject)dbData.get(0)).getValue("update_date");       		   
				   	   System.out.println("UPD_CODE="+UPD_CODE); 
				   	   System.out.println("UPDATE_DATE="+UPDATE_DATE); 
				   	   if(UPD_CODE.equals("N")) UPD_CODE="待檢核";
				   	   else if(UPD_CODE.equals("E")) UPD_CODE="檢核錯誤";
				   	   else if(UPD_CODE.equals("U")) UPD_CODE="檢核成功";
				   	   else UPD_CODE="待檢核";
				   	   System.out.println("UPD_CODE="+UPD_CODE);
				   	   M_YEAR  = Integer.toString(Integer.parseInt(UPDATE_DATE.substring(0,4))-1911);	
				   	   M_MONTH  = Integer.toString(Integer.parseInt(UPDATE_DATE.substring(4,6))-0);	
				   	   M_DATE  = Integer.toString(Integer.parseInt(UPDATE_DATE.substring(6,8))-0);	
				   	   UPDATE_DATE=M_YEAR+"年"+M_MONTH+"月"+M_DATE+"日";	
				   	   System.out.println("UPDATE_DATE="+UPDATE_DATE); 	
				   }   
				   row=sheet.getRow(ncacno.equals("ncacno")?113:117);
		    	   cell=row.getCell((short)0);
		  		   cell.setEncoding(HSSFCell.ENCODING_UTF_16);	  		
		  		   cell.setCellValue("檢核結果:"+(dbData.size()>0?UPD_CODE:"待檢核"));	 
		  		   	   
		  		   row=sheet.getRow(ncacno.equals("ncacno")?114:118);		  			
				   cell=row.getCell((short)0);
				   cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				   cell.setCellValue("最後異動日期:"+(dbData.size()>0?UPDATE_DATE:"無"));				  
                }//end of 總表列印檢核有信用部名稱
	  	    }//end of 有資料

	  		
  	        //設定涷結欄位
            //sheet.createFreezePane(0,1,0,1);
            footer.setCenter( "Page:" + HSSFFooter.page() + " of " +
                             HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));			
	       
  	        
  	        FileOutputStream fout=new FileOutputStream(reportDir + System.getProperty("file.separator")+"農業信用部資產負債表.xls");
	        wb.write(fout);
	        //儲存 
	        fout.close();
	        Utility.printLogTime("FR003W_Excel end time");	
		}catch(Exception e){
			System.out.println("createRpt Error:"+e+e.getMessage());
		}
		return errMsg;
	}	

		 
}
