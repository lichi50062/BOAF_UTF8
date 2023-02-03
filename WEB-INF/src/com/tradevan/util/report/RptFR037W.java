/*	 
    95.04.27 add 農漁會信用部逾期放款統計表 by 2295            
    95.07.04 fix 更改單一信用部.總表sql by 2295
    95.08.11 fix 拿掉減項:960500,調查合計位置 by 2295
    95.10.03 add 增加檢核結果與最後異動日期 by 2495
    95.11.13 fix 不顯示逾放比率 by 2295  
    99.05.12 fix fix sql injection by 2808
   102.11.19 add 原QueryDB改套用QueryDB_SQLParam by 2295    
   103.02.10 add 103/01以後,漁會套用新表格(增加/異動科目代號) by 2295
*/
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*;
import java.io.*;
import java.util.*;
import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR037W{
    public static String createRpt(String S_YEAR,String S_MONTH,String bank_code,String bank_type,String BANK_NAME,String unit){
		String errMsg = "";
		List dbData = null;
		String sqlCmd = "";
		List A06Data = new LinkedList();		
		String acc_code = "";
		String amt = "";		
		String unit_name="";		
		StringBuffer sql = new StringBuffer() ;
		List paramList = new ArrayList ();
		DataObject bean = null ;
		String u_year = "100" ;
		if(S_YEAR==null || Integer.parseInt(S_YEAR)<=99) {
			u_year = "99" ;
		}
		try{			
		    System.out.print("S_YEAR="+S_YEAR);
		    System.out.print(":S_MONTH="+S_MONTH);
		    System.out.print(":bank_code="+bank_code);
		    System.out.print(":bank_type="+bank_type);
		    System.out.print(":BANK_NAME="+BANK_NAME);
		    System.out.println(":unit="+unit);
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
			//FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+"農業信用部損益表"+((hasMonth.equals("true"))?"_單月":"")+".xls" );
    		FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+(bank_type.equals("6")?"農":"漁")+"會信用部逾期放款統計表.xls" );
			
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

	        ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
	  		//wb.setSheetName(0,"test");
	  		finput.close();
	  		
	  		HSSFRow row=null;//宣告一列 
	  		HSSFCell cell=null;//宣告一個儲存格  		
	  		
	  		short i=0;
	  		String ncacno = (bank_type.equals("6"))?"ncacno":"ncacno_7";
	  	    //103.02.10 add 103/01以後,漁會套用新表格(增加/異動科目代號)
            if( bank_type.equals("7") && (Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH) >= 10301) ){
                ncacno = "ncacno_7_rule";
            }          
	  		//String a01table="";
	  		StringBuffer a01table =  new StringBuffer() ;
	  		//String a06rule="";
	  		StringBuffer a06rule = new StringBuffer() ;
	  		//String a01rule="";
	  		StringBuffer a01rule = new StringBuffer() ;
	  		if(!bank_code.equals("ALL")){//單一信用部
	  		    if(bank_type.equals("6")){//農會
		  		   a01table.append("a01");
		  		}else{//漁會
		  		   if(Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH) >= 10301){
		                a01table.append(" a01 ") ;
		  		   }else{     
		  			a01table.append(" (select "+S_YEAR+" as m_year,"+String.valueOf(Integer.parseInt(S_MONTH))+" as m_month,'"+bank_code+"' as bank_code,acc_code,amt ");
		  			a01table.append("  from a01 ");
		  			a01table.append("  where acc_code not in (?,?) ");
		  			a01table.append("  and m_year=? and m_month= ? and bank_code=? ");
		  			a01table.append("  union ");
		  			a01table.append("  select "+S_YEAR+"  as m_year,"+String.valueOf(Integer.parseInt(S_MONTH))+" as m_month,'"+bank_code+"' as bank_code,'120700',sum(amt) amt from a01 ");  
		  			a01table.append("  where acc_code in(?,?)" );
		  			a01table.append("  and m_year=? and m_month=? and bank_code=? ) a01 ");
		  		   }	
		  		}
	  		    /*
	  		    sqlCmd = " select A06.m_year, A06.m_month, "
	  		           + " "+ncacno+".acc_range,A06.acc_code, "+ncacno+".acc_name,"
	  		           + " round(amt_3month/"+unit+",0) as amt_3month, " //未滿3個月
	  		           + " round(amt_6month/"+unit+",0) as amt_6month, " //3個月~6個月
	  		           + " round(amt_1year/"+unit+",0) as amt_1year, " //6個月~1年
	  		           + " round(amt_2year/"+unit+",0) as amt_2year, " //1年~2年
	  		           + " round(amt_over2year/"+unit+",0) as amt_over2year, " //2年以上
	  		           + " round(amt_total/"+unit+",0) as amt_total, " //逾放合計
	  		           + " round(A01.amt/"+unit+",0) as amt, " //a01放款合計
	  		           + " round(round(amt_total/"+unit+",0)/ decode(round(A01.amt/"+unit+",0),0,1,round(A01.amt/"+unit+",0)) * 100,3) as a01_per " //逾放比率 
	  		           + " from A06 LEFT JOIN  "+ncacno+" ON A06.acc_code =  "+ncacno+".acc_code and  "+ncacno+".acc_div='08' " 
	  		           + " LEFT JOIN  (select m_year,m_month,bank_code, "
				       + "        			  decode(acc_code,'990000','970000',acc_code) as acc_code,amt "
                       + "			   from "+a01table.toString()                       
					   + "			   )A01 ON A01.bank_code = A06.bank_code " 
                       + "   			    and A01.m_year = A06.m_year and A01.m_month = A06.m_month " 
					   + "			    	and A01.acc_code = A06.acc_code "              	              
					   + " where  A06.m_year="+S_YEAR+ " and A06.m_month="+String.valueOf(Integer.parseInt(S_MONTH))
					   + " and A06.bank_code = '"+bank_code+"'"
					   + " order by  "+ncacno+".acc_range ";
	  		    */
	  		    sql.append(" select A06.m_year, A06.m_month, ");
	  		    sql.append(ncacno).append(".acc_range,A06.acc_code, ").append(ncacno).append(".acc_name,");
	  		    sql.append(" round(amt_3month/ ? ,0) as amt_3month,");
		  		sql.append(" round(amt_6month/ ? ,0) as amt_6month,");
		  		sql.append(" round(amt_1year/  ?,0) as amt_1year,");
		  		sql.append(" round(amt_2year/  ?,0) as amt_2year,");
		  		sql.append(" round(amt_over2year/?,0) as amt_over2year,");
		  		sql.append(" round(amt_total/?,0) as amt_total,");
		  		sql.append(" round(A01.amt/?,0) as amt, ");
		  		sql.append(" round(round(amt_total/?,0)/ decode(round(A01.amt/?,0),0,1,round(A01.amt/?,0)) * 100,3) as a01_per");
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
		  		sql.append(" from A06 LEFT JOIN  ").append(ncacno).append(" ON A06.acc_code =  ").append(ncacno).append(".acc_code and  ").append(ncacno).append(".acc_div= ? ");
		  		paramList.add("08") ;
		  		sql.append(" LEFT JOIN  (select m_year,m_month,bank_code,");
		  		sql.append("                    decode(acc_code,'990000','970000',acc_code) as acc_code,amt ");
		  		sql.append(" 			 from ").append(a01table.toString());
		  		//=====a01table 條件===
		  		if(!"6".equals(bank_type)){//漁會
		  		    if(Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH) < 10301){
			  		   paramList.add("120700") ;
		  			   paramList.add("120900") ;
		  			   paramList.add(S_YEAR) ;
		  			   paramList.add(String.valueOf(Integer.parseInt(S_MONTH))) ;
		  			   paramList.add(bank_code) ;
		  			   paramList.add("120700") ;
		  			   paramList.add("120900") ;
		  			   paramList.add(S_YEAR) ;
		  			   paramList.add(String.valueOf(Integer.parseInt(S_MONTH))) ;
		  			   paramList.add(bank_code) ;
		  		    }
		  		}
		  		sql.append("    		 )A01 ON A01.bank_code = A06.bank_code");
		  		sql.append(" 			 and A01.m_year = A06.m_year and A01.m_month = A06.m_month");
		  		sql.append("			 and A01.acc_code = A06.acc_code");
		  		sql.append(" where  A06.m_year= ? and A06.m_month= ? ");
		  		paramList.add(S_YEAR) ;
		  		paramList.add(String.valueOf(Integer.parseInt(S_MONTH))) ;
		  		sql.append(" and A06.bank_code = ? ");
		  		paramList.add(bank_code) ;
		  		sql.append(" order by  ").append(ncacno).append(".acc_range ");
		  		
	  		}else{//全體信用部
	  		   System.out.println("全體信用部==================") ;
	  		   if(bank_type.equals("6")){//農會
		  		   a01table.append("a01");
		  	   }else{//漁會		
		  	       if(Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH) >= 10301){
                      a01table.append(" a01 ") ;
                   }else{     
		  		      a01table.append(" (select m_year,m_month,bank_code,acc_code,amt ");
		  		      a01table.append("  from a01 ");
		  		      a01table.append(" where acc_code not in (?,?) ");
		  		      a01table.append(" and m_year= ?  and m_month= ? ");
		  		      a01table.append(" union ");
		  		      a01table.append(" select  m_year,m_month,bank_code,'120700' as acc_code,sum(amt) amt from a01 ");  
		  		      a01table.append(" where acc_code in(?,?)" );
		  		      a01table.append(" and m_year=? and m_month= ?");
		  		      a01table.append(" group by m_year,m_month,bank_code ");
		  	          a01table.append(" order by m_year,m_month,bank_code ) a01 "); 
                   }
		  	   } 
	  		 a06rule.append(" select A06.m_year, A06.m_month, "+ncacno+".acc_range, a06.acc_code,"+ncacno+".acc_name" 
	       	   + " 		  ,round(sum(amt_3month)/ ?,0) as amt_3month " //未滿3個月
	       	   + " 		  ,round(sum(amt_6month)/ ? ,0) as amt_6month " //3個月~6個月
	       	   + " 		  ,round(sum(amt_1year)/ ?,0) as amt_1year " //6個月~1年
	       	   + "		  ,round(sum(amt_2year)/ ? ,0) as amt_2year " //1年~2年
	       	   + "		  ,round(sum(amt_over2year)/?,0) as amt_over2year " //2年以上
	       	   + " 		  ,round(sum(amt_total)/?,0) as amt_total " //逾放合計
	       	   + " from A06 LEFT JOIN "+ncacno+" ON A06.acc_code = "+ncacno+".acc_code "
	       	   + " ,(select * from bn01 where m_year=? and bank_type=? )bn01 " 
	       	   + " where A06.m_year= ? and A06.m_month= ? "
	       	   + " and "+ncacno+".acc_div= ? "
	       	   + " and A06.bank_code = bn01.bank_no" 
	       	   //+ " and bn01.bank_type= ?  "
	       	   + " group by A06.m_year, A06.m_month, "+ncacno+".acc_range, A06.acc_code, "+ncacno+".acc_name"
	       	   + " order by A06.m_year, A06.m_month, "+ncacno+".acc_range, A06.acc_code, "+ncacno+".acc_name");
	  		   paramList.add(unit) ;
	  		   paramList.add(unit) ;
		  	   paramList.add(unit) ;
		  	   paramList.add(unit) ;
		  	   paramList.add(unit) ;
		  	   paramList.add(unit) ;
	  		   paramList.add(u_year) ;
	  		   paramList.add(bank_type) ;
	  		   paramList.add(S_YEAR) ;
	  		   paramList.add(String.valueOf(Integer.parseInt(S_MONTH))) ;
	  		   paramList.add("08") ;
	  		  
	  		a01rule.append(" select A01.m_year, A01.m_month, "+ncacno+".acc_range," 
	       	   + "		  decode(a01.acc_code,'990000','970000',a01.acc_code) as acc_code,"+ncacno+".acc_name," 
	       	   + " 		  round(sum(amt)/?,0) as amt " //a01放款合計
	       	   + " from "+a01table.toString()+" LEFT JOIN "+ncacno+" ON A01.acc_code = "+ncacno+".acc_code "
	       	   + "		,(select * from bn01 where m_year=? and bank_type=? ) bn01 " 
	       	   + " where A01.m_year= ? and A01.m_month= ? "
	       	   + " and ("+ncacno+".acc_div= ?  or "+ncacno+".acc_code= ? )"
	       	   + " and A01.bank_code = bn01.bank_no" 
	       	   //+ " and bn01.bank_type= ? "
	       	   + " group by A01.m_year, A01.m_month, "+ncacno+".acc_range, A01.acc_code,"+ncacno+".acc_name"
	       	   + " order by A01.m_year, A01.m_month, "+ncacno+".acc_range, A01.acc_code,"+ncacno+".acc_name");  
	  		   paramList.add(unit) ;
	  		   paramList.add(u_year) ;
	  		   paramList.add(bank_type) ;
	  		   if(!"6".equals(bank_type)) {//漁會
	  		      if(Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH) < 10301){
	  			     paramList.add("120700") ;
		  	         paramList.add("120900") ;
		  	         paramList.add(S_YEAR) ;
		  	         paramList.add(String.valueOf(Integer.parseInt(S_MONTH))) ;
		  	         paramList.add("120700") ;
		  	         paramList.add("120900") ;
		  	         paramList.add(S_YEAR) ;
		  	         paramList.add(String.valueOf(Integer.parseInt(S_MONTH))) ;
	  		      }
	  		   }
	  		   paramList.add(S_YEAR) ;
	  		   paramList.add(String.valueOf(Integer.parseInt(S_MONTH))) ;
	  		   paramList.add("08") ;
	  		   paramList.add("990000") ;
	  		   //paramList.add(bank_type) ;
	  		    /*
	  		    sqlCmd = " select a06.m_year,a06.m_month,a06.acc_range,a06.acc_code,a06.acc_name,a06.amt_3month,a06.AMT_6MONTH,"
              		   + "	      a06.AMT_1YEAR,a06.AMT_2YEAR,a06.AMT_OVER2YEAR,a06.AMT_TOTAL,a01.amt "
              		   + "		  ,round(a06.AMT_TOTAL/ decode(a01.amt,0,1,a01.amt) * 100,3) as a01_per " //逾放比率
              		   + " from "	   
              		   + " ("+a06rule+")a06 left join ("+a01rule+")a01" 
              		   + "                             ON A06.m_year = A01.m_year " 
                       + "                            and A06.m_month = A01.m_month " 
					   + "                            and A06.acc_code = A01.acc_code "  
                       + " order by a06.m_year,a06.m_month,a06.acc_range ";
                       */
	  		   sql.append(" select a06.m_year,a06.m_month,a06.acc_range,a06.acc_code,a06.acc_name,a06.amt_3month,a06.AMT_6MONTH, ") ;
		  	   sql.append("        a06.AMT_1YEAR,a06.AMT_2YEAR,a06.AMT_OVER2YEAR,a06.AMT_TOTAL,a01.amt  ") ;
		  	   sql.append("        ,round(a06.AMT_TOTAL/ decode(a01.amt,0,1,a01.amt) * 100,3) as a01_per ") ;
		  	   sql.append(" from  ") ;
		  	   sql.append(" ("+a06rule.toString()+")a06 left join ("+a01rule.toString()+")a01") ;
		  	   sql.append(" ON A06.m_year = A01.m_year ") ;
		  	   sql.append(" and A06.m_month = A01.m_month ") ;
		  	   sql.append(" and A06.acc_code = A01.acc_code  ") ;
		  	   sql.append(" order by a06.m_year,a06.m_month,a06.acc_range ") ;
	  		}
	  		dbData = DBManager.QueryDB_SQLParam(sql.toString(), paramList, "m_year,m_month,amt_3month,amt_6month,amt_1year,amt_2year,amt_over2year,amt_total,amt,a01_per" );
			//dbData = DBManager.QueryDB(sqlCmd,"m_year,m_month,amt_3month,amt_6month,amt_1year,amt_2year,amt_over2year,amt_total,amt,a01_per");	  	         
			List amt_List =null;
  			for(i=0;i<dbData.size();i++){  		
  				bean = (DataObject)dbData.get(i) ;
  			    amt_List = new LinkedList();
  				acc_code = (String)bean.getValue("acc_code");
  				amt_List.add(acc_code);
  				amt_List.add((String)bean.getValue("acc_name"));  				
  				amt_List.add(bean.getValue("amt_3month").toString());
  				amt_List.add(bean.getValue("amt_6month").toString());  				
  				amt_List.add(bean.getValue("amt_1year").toString());  				
  				amt_List.add(bean.getValue("amt_2year").toString());  				
  				amt_List.add(bean.getValue("amt_over2year").toString());  				
  				amt_List.add(bean.getValue("amt_total").toString());
  				if(((DataObject)dbData.get(i)).getValue("amt") != null){
  				   amt_List.add(bean.getValue("amt").toString());
  				}else{
  				   amt_List.add("0"); 
  				}
  				if(bean.getValue("a01_per") != null){  				 
  				    if(bean.getValue("a01_per").toString().indexOf(".") != -1){
  				       //取到小數第3位
  				       amt_List.add(bean.getValue("a01_per").toString());  				       
  				    }else{
  				       amt_List.add("0");				       
  				    }   
  				}else{
  				   amt_List.add("0");
  				}  				
  				System.out.println(amt_List);
  				A06Data.add(amt_List);	         		       	    	       	    
	       	}
  			
  			row=sheet.getRow(0);
  	 		cell=row.getCell((short)0);	       	
  	 		//設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
  	 		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
  	 		if(bank_code.equals("ALL")){//全体農漁會	  	 		   
	  		   cell.setCellValue("全體"+(bank_type.equals("6")?"農會":"漁會")+"信用部逾期放款統計表");	  	 		   
	  		}else{		  		   
	          	   cell.setCellValue(BANK_NAME+"逾期放款統計表");		  		   
	  		}        
  	        if(dbData.size() == 0){	
  	            row=sheet.getRow(1);
	  	 		cell=row.getCell((short)3);	       	
	  			//設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
	  			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
   	       	    cell.setCellValue(S_YEAR +"年" +S_MONTH +"月無資料存在");
	  	 	}else{	   	
	  	 		row=sheet.getRow(1);
	  	 		cell=row.getCell((short)3);	       	
	  	 		cell.setEncoding(HSSFCell.ENCODING_UTF_16);	  	 		
   	          	cell.setCellValue("中華民國" +S_YEAR +"年" +S_MONTH+"月");
   	          	
   	            //列印單位
	          	row=sheet.getRow(1);	          	
	            cell=(row.getCell((short)14)==null)? row.createCell((short)14) : row.getCell((short)14);
	            	          	   	 		
	 		   	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	 		   	//取得金額單位中文名稱
	 		   	unit_name = Utility.getUnitName(unit) ;
	 		   	
	 		   	cell.setCellValue("單位：新台幣"+unit_name+"、％");
	  	 		//以巢狀迴圈讀取所有儲存格資料 
	  	 		//System.out.println("total row ="+sheet.getLastRowNum());
	  	 		
	  	 		for(i=4;i<20;i++){
	  	 		    if(i==19){//合計
	  	 		       row=sheet.getRow(i+4);
		          	}else{
	  	 		       row=sheet.getRow(i);
		          	}   
	  	 		    //System.out.println("i="+i);
	  	 		    amt_List = (List)A06Data.get(i-4);
	  	 		    acc_code = ((String)amt_List.get(0)).trim();//科目代號
	  	 		    System.out.println(amt_List);
	  	 		    //System.out.println("acc_code="+acc_code);
	  	 		    cell=row.getCell((short)0);
	  	 		    cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
	  	 		    if(acc_code.equals("120700") && bank_type.equals("7")){
	  	 		       cell.setCellValue((String)amt_List.get(1)+"(內含120900)");//項目
	  	 		    }else{    
	  	 		       cell.setCellValue((String)amt_List.get(1));//項目
	  	 		    }
	  	 		    //System.out.print(":"+(String)amt_List.get(1));
	  	 		    cell=row.getCell((short)2);//未滿3個月
	  	 		    cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
		          	amt = Utility.setCommaFormat((String)amt_List.get(2));		          	
		          	if(!amt.equals("0"))cell.setCellValue(amt);
		          	//System.out.print(":"+(String)amt_List.get(2));
		          	cell=row.getCell((short)4);//3個月~6個月
	  	 		    cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
		          	amt = Utility.setCommaFormat((String)amt_List.get(3));		          	
		          	if(!amt.equals("0"))cell.setCellValue(amt);
		          	//System.out.print(":"+(String)amt_List.get(3));
		          	cell=row.getCell((short)6);//6個月~1年
	  	 		    cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
		          	amt = Utility.setCommaFormat((String)amt_List.get(4));		          	
		          	if(!amt.equals("0"))cell.setCellValue(amt);
		          	//System.out.print(":"+(String)amt_List.get(4));
		          	cell=row.getCell((short)8);//1年~2年
	  	 		    cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
		          	amt = Utility.setCommaFormat((String)amt_List.get(5));		          	
		          	if(!amt.equals("0"))cell.setCellValue(amt);
		          	//System.out.print(":"+(String)amt_List.get(5));
		          	cell=row.getCell((short)10);//2年以上
	  	 		    cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
		          	amt = Utility.setCommaFormat((String)amt_List.get(6));		          	
		          	if(!amt.equals("0"))cell.setCellValue(amt);
		          	//System.out.print(":"+(String)amt_List.get(6));
		          	cell=row.getCell((short)12);//逾放合計
	  	 		    cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
		          	amt = Utility.setCommaFormat((String)amt_List.get(7));		          	
		          	if(!amt.equals("0"))cell.setCellValue(amt);
		          	//System.out.println(":"+(String)amt_List.get(7));
		          	
		          	if(!acc_code.equals("960500")){
		          	   cell=row.getCell((short)14);//A01放款合計
	  	 		       cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
		          	   amt = Utility.setCommaFormat((String)amt_List.get(8));		          	
		          	   if(!amt.equals("0"))cell.setCellValue(amt);
		          	   //System.out.print(":"+(String)amt_List.get(8));
		          	}  
		          	/*95.11.13 fix 不顯示逾放比率 
		          	if(!acc_code.equals("960500") && !acc_code.equals("970000")){		          	   
		          	   cell=row.getCell((short)16);//逾放比率
	  	 		       cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
		          	   amt = Utility.setCommaFormat((String)amt_List.get(9));		          	
		          	   if(!amt.equals("0"))cell.setCellValue(amt);
		          	   //System.out.println(":"+(String)amt_List.get(9));
		          	} 
		          	*/
	  	 		}
	  	 	if(!bank_code.equals("ALL")){	
	  	 	  paramList = new ArrayList();
	  	 	 //95.10.03 增加檢核結果與最後異動日期 by 詹雅惠	  			  		
	  		sqlCmd = " select UPD_CODE, to_char(UPDATE_DATE,'yyyymmdd') as UPDATE_DATE"
				+ " from WML01"		   		  
				+ " where M_YEAR=?"		   		   
				+ " and M_MONTH=?"
				+ " and BANK_CODE=?"
 				+ " and REPORT_NO=?";
	  		paramList.add(S_YEAR);
	  		paramList.add(S_MONTH);
	  		paramList.add(bank_code);
	  		paramList.add("A06");
		   	//System.out.println("sqlCmd="+sqlCmd); 	   
		   		   
				dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");	  
    		String UPD_CODE="";
    		String UPDATE_DATE="";
    		String M_YEAR="";   
    		String M_MONTH="";    
    		String M_DATE=""; 
    		if(dbData != null){
       		   System.out.println("dbData.size()="+dbData.size()); 
       		    UPD_CODE = (String)((DataObject)dbData.get(0)).getValue("upd_code");  
       		    UPDATE_DATE = (String)((DataObject)dbData.get(0)).getValue("update_date");       		   
					   System.out.println("UPD_CODE="+UPD_CODE); 
					   System.out.println("UPDATE_DATE="+UPDATE_DATE); 
					   if(UPD_CODE.equals("N")) UPD_CODE="待檢核";
					   else if(UPD_CODE.equals("E")) UPD_CODE="檢核錯誤";
					   else if(UPD_CODE.equals("U")) UPD_CODE="檢核通過";
					   else UPD_CODE="";
					   System.out.println("UPD_CODE="+UPD_CODE);
					   M_YEAR  = Integer.toString(Integer.parseInt(UPDATE_DATE.substring(0,4))-1911);	
					   M_MONTH  = Integer.toString(Integer.parseInt(UPDATE_DATE.substring(4,6))-0);	
					   M_DATE  = Integer.toString(Integer.parseInt(UPDATE_DATE.substring(6,8))-0);	
					   UPDATE_DATE=M_YEAR+"年"+M_MONTH+"月"+M_DATE+"日";	
					   System.out.println("UPDATE_DATE="+UPDATE_DATE); 			   
       		}	
	  			  		
	  		row=sheet.getRow(27);	    		
	    	cell=row.getCell((short)0);
	  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);	  		
	  		cell.setCellValue("檢核結果:"+UPD_CODE);	  
	  	  		
        row=sheet.getRow(28);	    		
	    	cell=row.getCell((short)0);
	  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		cell.setCellValue("最後異動日期:"+UPDATE_DATE);	  
	  	 }
	  	 	
	  	 	}//end of monthExist
	    	FileOutputStream fout=new FileOutputStream(reportDir + System.getProperty("file.separator")+"農漁會信用部逾期放款統計表.xls");
	    	wb.write(fout);
	    	//儲存 
	    	fout.close();
	        
		}catch(Exception e){
			System.out.println("createRpt Error:"+e+e.getMessage());
		}
		return errMsg;
	}		 
}
