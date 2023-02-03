/*
	 94.03.11 fix 金額資料為零是不輸出0，改為輸出空白及欄位資料右靠處理
	 94.08.15 fix 報表title日期格式 by 2295
	 94.11.18 add 增加全体漁會損益表/金額單位 by 2295
	 95.01.26 fix 全体漁會加上只抓bank_type='7' by 2295
	 95.02.10 fix 全体漁會信用部/bank_name會變成亂碼的問題 by 2295
	 95.03.16 add 單月金額 by 2295
			      1月份時.當月金額與累計金額相同
     95.03.28 add 損益表.單月損益表分開二份報表 by 2295
     95.04.07 fix 單月損益表.只放單月金額.不放累計金額 by 2295
     95.04.11 add 單月損益表.判斷本月跟上月有無申報資料 by 2295	  	
     95.05.08 fix 單月損益表.上月or本月無資料時,顯示成中華民國xx年xx月(xx年xx月無資料) by 2295
                                                    本月無資料時,也要顯示信用部名稱 by 2295	
     96.11.06 add 增加檢核結果與最後異動日期 by 2295      
     99.04.27 fix sql injection & 縣市合併 by 2808 
    102.05.31 add 103年度後報表 by 2968    
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

public class FR004F_Excel {
    public static String createRpt(String S_YEAR,String S_MONTH,String bank_code,String BANK_NAME,String unit,String hasMonth){
		String errMsg = "";
		List dbData = null;
		String sqlCmd = "";
		Properties A01Data = new Properties();		
		Properties A01_monthData = new Properties();
		String acc_code = "";
		String amt = "";
		String month_amt = "";
		String YEAR = "";
		String unit_name="";
		boolean monthExist = true;
		StringBuffer sql = new StringBuffer() ;
		ArrayList paramList = new ArrayList() ;
		DataObject bean = null ;
		String fileName = "漁會信用部損益表.xls";
        String ncacno_7 = "ncacno_7" ;
		String u_year = "100" ;
		if(S_YEAR!=null && Integer.parseInt(S_YEAR)<= 99) {
			u_year = "99" ;
		}
		if(Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH) >= 10301){
            ncacno_7 = "ncacno_7_rule";
            fileName = "漁會信用部損益表_10301.xls";
        }
		String lastMonth=(Integer.parseInt(S_MONTH)==1)?"12":String.valueOf(Integer.parseInt(S_MONTH)-1);
		String lastYear=(lastMonth=="12")?String.valueOf(Integer.parseInt(S_YEAR)-1):S_YEAR;
		try{	
			System.out.println(fileName);
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
			//FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+"漁會信用部損益表"+((hasMonth.equals("true"))?"_單月":"")+".xls" );
			FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+fileName );
			//95.02.10 fix bank_name變成亂碼
			sql.setLength(0) ;
            paramList.clear();
            sql.append("select bank_no,bank_name from bn01 where bank_type='7' and bank_no=? and m_year=? ") ;
            paramList.add(bank_code) ;
            paramList.add(u_year) ;
            dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"") ;
            if(dbData != null && dbData.size()!=0 ){
                BANK_NAME=(String)((DataObject)dbData.get(0)).getValue("bank_name");
            }
            sql.setLength(0) ;
            paramList.clear();
	  	    //設定FileINputStream讀取Excel檔
	  		POIFSFileSystem fs = new POIFSFileSystem( finput );
	  		HSSFWorkbook wb = new HSSFWorkbook(fs);
	  		HSSFSheet sheet = wb.getSheetAt(0);//讀取第一個工作表，宣告其為sheet 
	  		HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
	        //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
	        //sheet.setAutobreaks(true); //自動分頁
			
	        //設定頁面符合列印大小
	        sheet.setAutobreaks( false );
	        ps.setScale( ( short )76 ); //列印縮放百分比

	        ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
	  		//wb.setSheetName(0,"test");
	  		finput.close();
	  		
	  		HSSFRow row=null;//宣告一列 
	  		HSSFCell cell=null;//宣告一個儲存格  		
	  		
	  		short i=0;
	  		short y=0;	  
	  		if(!hasMonth.equals("true")){//原損益表
	  		      //sqlCmd = " select A01.m_year, A01.m_month, ncacno_7.acc_range, a01.acc_code,";
	  			  sql.setLength(0) ;
	  			  sql.append(" select A01.m_year, A01.m_month, ").append(ncacno_7).append(".acc_range, a01.acc_code, ") ;
		          if(bank_code.equals("ALL")){//全体漁會	  		    
		              //sqlCmd += " round(sum(amt)/"+unit+",0) as amt";
		        	  sql.append(" round(sum(amt)/ ? ,0) as amt ");
		        	  paramList.add(unit) ;
		          }else{	  		   
		              //sqlCmd += " round(amt/"+unit+",0) as amt";
		        	  sql.append(" round(amt/? ,0) as amt ");
		        	  paramList.add(unit) ;
		          }
		          //sqlCmd += " from A01 LEFT JOIN ncacno_7 ON A01.acc_code = ncacno_7.acc_code";
		          sql.append(" from A01 LEFT JOIN ").append(ncacno_7).append(" ON A01.acc_code = ").append(ncacno_7).append(".acc_code ") ;
		          if(bank_code.equals("ALL")){//全体漁會 fix 95.01.26
		              //sqlCmd += ",bn01 ";
		        	  sql.append(" ,(select * from bn01 where m_year = ? and bank_type=? ) bn01 ");
		        	  paramList.add(u_year) ;
		        	  paramList.add("7") ;
		          }
		          //sqlCmd += " where A01.m_year="+S_YEAR
		          sql.append(" where A01.m_year=? ");
		          //+ "   and A01.m_month="+String.valueOf(Integer.parseInt(S_MONTH));
		          sql.append(" and A01.m_month= ? ");
		          paramList.add(S_YEAR) ;
		          paramList.add(S_MONTH) ;
		          if(!bank_code.equals("ALL")){
		              //sqlCmd += "   and A01.bank_code='"+bank_code+"'";
		        	  sql.append(" and A01.bank_code= ? ");
		        	  paramList.add(bank_code) ;
		          }	   
		          //sqlCmd += "   and ncacno_7.acc_div='02'";
		          sql.append(" and ").append(ncacno_7).append(".acc_div= ?  ");
		          paramList.add("02") ;
		          if(bank_code.equals("ALL")){//全体漁會 fix 95.01.26
		              //sqlCmd += " and A01.bank_code = bn01.bank_no" 
		              //       + " and bn01.bank_type='7'"
		              //       + " group by A01.m_year, A01.m_month, ncacno_7.acc_range, A01.acc_code"
		              //      + " order by A01.m_year, A01.m_month, ncacno_7.acc_range, A01.acc_code";
		        	  sql.append(" and A01.bank_code = bn01.bank_no ");
		        	  sql.append(" and bn01.bank_type= ? ");
		        	  sql.append(" group by A01.m_year, A01.m_month, ").append(ncacno_7).append(".acc_range, A01.acc_code ");
		        	  sql.append(" order by A01.m_year, A01.m_month, ").append(ncacno_7).append(".acc_range, A01.acc_code ");
		        	  paramList.add("7") ;
		        	}else{
		              //sqlCmd += " order by ncacno_7.acc_range";
		        	  sql.append(" order by ").append(ncacno_7).append(".acc_range ") ;
		          }
	  		}else{//單月損益表
	  		       //sqlCmd = " select a.m_year , a.m_month , a.acc_range , a.acc_code , nvl(a.amt,0) as amt,nvl(b.amt,0) as b_amt, nvl((a.amt  - b.amt),0) as month_amt"
	  		       //+ " from ("
	  		       //+ "        select A01.m_year, A01.m_month, ncacno_7.acc_range, a01.acc_code,";
	  			   sql.setLength(0) ;
	  			   sql.append(" select a.m_year , a.m_month , a.acc_range , a.acc_code , nvl(a.amt,0) as amt,nvl(b.amt,0) as b_amt, nvl((a.amt  - b.amt),0) as month_amt ");
	  			   sql.append(" from ( ");
	  			   sql.append("       select A01.m_year, A01.m_month, ").append(ncacno_7).append(".acc_range, a01.acc_code, ");
	  			   if(bank_code.equals("ALL")){//全体漁會	  		    
	  			      //sqlCmd += " round(sum(amt)/"+unit+",0) as amt";
	  			      sql.append("  round(sum(amt)/ ? ,0) as amt ");
	  			      paramList.add(unit) ;
	  			   }else{	  		   
	  			      //sqlCmd += " round(amt/"+unit+",0) as amt";
	  			      sql.append("  round(amt/? ,0) as amt ") ;
	  			      paramList.add(unit) ;
	  			   }
	  			   //sqlCmd += " from A01 LEFT JOIN ncacno_7 ON A01.acc_code = ncacno_7.acc_code";
	  			   sql.append(" from A01 LEFT JOIN ").append(ncacno_7).append(" ON A01.acc_code = ").append(ncacno_7).append(".acc_code ") ;
	  			   if(bank_code.equals("ALL")){//全体漁會 fix 95.01.26
	  			      //sqlCmd += ",bn01 ";
	  			      sql.append(" ,(select * from bn01 where m_year=? and bank_type = ? ) bn01 ") ;
	  			      paramList.add(u_year) ;
	  			      paramList.add("7") ;
	  			   }
	  			   //sqlCmd += " where A01.m_year="+S_YEAR 
	  			   //+ "   and A01.m_month="+String.valueOf(Integer.parseInt(S_MONTH));
	  			   sql.append(" where A01.m_year= ? ") ;
	  			   sql.append(" and A01.m_month= ?  ");
	  			   paramList.add(S_YEAR) ;
	  			   paramList.add(String.valueOf(Integer.parseInt(S_MONTH))) ;
	  			   if(!bank_code.equals("ALL")){
	  			      //sqlCmd += "   and A01.bank_code='"+bank_code+"'";
	  			      sql.append(" and A01.bank_code= ? ");
	  			      paramList.add(bank_code) ;
	  			   }	   
	  			   //sqlCmd += "   and ncacno_7.acc_div='02'";
	  			   sql.append(" and ").append(ncacno_7).append(".acc_div= ? ");
	  			   paramList.add("02") ;
	  			   if(bank_code.equals("ALL")){//全体漁會 fix 95.01.26
	  			      //sqlCmd += " and A01.bank_code = bn01.bank_no" 
	  			      //       + " and bn01.bank_type='7'"
	  			      //       + " group by A01.m_year, A01.m_month, ncacno_7.acc_range, A01.acc_code"
	  			      //       + " order by A01.m_year, A01.m_month, ncacno_7.acc_range, A01.acc_code";
	  			      sql.append(" and A01.bank_code = bn01.bank_no ") ;
	  			      sql.append(" and bn01.bank_type= ? ") ;
	  			      sql.append(" group by A01.m_year, A01.m_month, ").append(ncacno_7).append(".acc_range, A01.acc_code ") ;
	  			      sql.append(" order by A01.m_year, A01.m_month, ").append(ncacno_7).append(".acc_range, A01.acc_code ") ;
	  			      paramList.add("7") ;
	  			   }else{
	  			      //sqlCmd += " order by ncacno_7.acc_range";
	  			      sql.append(" order by ").append(ncacno_7).append(".acc_range ") ;
	  			   }
	  		       //sqlCmd += " ) a"
	  		       //       +  " left join "
	  		       //       +  " ( "
	  		       //       + "   select A01.m_year, A01.m_month, ncacno_7.acc_range, a01.acc_code,";
	  		       sql.append(" ) a ");
	  		       sql.append(" left join ");
	  		       sql.append(" ( ") ;
	  		       sql.append(" select A01.m_year, A01.m_month, ").append(ncacno_7).append(".acc_range, a01.acc_code,  ");
				   if(bank_code.equals("ALL")){//全体漁會	  		    
				      //sqlCmd += " round(sum(amt)/"+unit+",0) as amt";
				      sql.append(" round(sum(amt)/ ? ,0) as amt ");
				      paramList.add(unit) ;
				   }else{	  	
				      sql.append(" round(amt/ ? ,0) as amt ") ;
				      paramList.add(unit);
				      //sqlCmd += " round(amt/"+unit+",0) as amt";
				   }
				   //sqlCmd += " from A01 LEFT JOIN ncacno_7 ON A01.acc_code = ncacno_7.acc_code";
				   sql.append(" from A01 LEFT JOIN ").append(ncacno_7).append(" ON A01.acc_code = ").append(ncacno_7).append(".acc_code ");
				   if(bank_code.equals("ALL")){//全体漁會 fix 95.01.26
				      //sqlCmd += ",bn01 ";
				      sql.append(",(select * from bn01 where u_year=? and bank_type=? ) bn01 ") ;
				      paramList.add(u_year) ;
				      paramList.add("7") ;
				   }
				   //sqlCmd += " where A01.m_year="+S_YEAR 
				   //+ "   and A01.m_month="+String.valueOf(Integer.parseInt(S_MONTH) - 1);
				   sql.append(" where A01.m_year= ? ");
				   sql.append(" and A01.m_month= ? ");
				   paramList.add(S_YEAR);
				   paramList.add(String.valueOf(Integer.parseInt(S_MONTH) - 1)) ;
 				   if(!bank_code.equals("ALL")){
				      //sqlCmd += "   and A01.bank_code='"+bank_code+"'";
 				      sql.append(" and A01.bank_code= ? ") ;
 				      paramList.add(bank_code) ;
				   }	   
				   //sqlCmd += "   and ncacno_7.acc_div='02'";
 				   sql.append("  and ").append(ncacno_7).append(".acc_div= ? ");
 				   paramList.add("02") ;
				   if(bank_code.equals("ALL")){//全体漁會 fix 95.01.26
				      //sqlCmd += " and A01.bank_code = bn01.bank_no" 
				      //       + " and bn01.bank_type='7'"
				      //       + " group by A01.m_year, A01.m_month, ncacno_7.acc_range, A01.acc_code"
				      //       + " order by A01.m_year, A01.m_month, ncacno_7.acc_range, A01.acc_code";
				      sql.append(" and A01.bank_code = bn01.bank_no ");
				      sql.append(" and bn01.bank_type= ? ");
				      sql.append(" group by A01.m_year, A01.m_month, ").append(ncacno_7).append(".acc_range, A01.acc_code ");
				      sql.append(" order by A01.m_year, A01.m_month, ").append(ncacno_7).append(".acc_range, A01.acc_code ");
				      paramList.add("7") ;
				   }else{
				       //sqlCmd += " order by ncacno_7.acc_range";
				       sql.append(" order by ").append(ncacno_7).append(".acc_range ");
				   }
		           //sqlCmd += " ) b "
			       //	   +  " on a.m_year = b.m_year and a.acc_code = b.acc_code ";
		           sql.append(" ) b ");
		           sql.append(" on a.m_year = b.m_year and a.acc_code = b.acc_code ");
	  		}//end of 單月損益表
	  		dbData = DBManager.QueryDB_SQLParam(sql.toString(), paramList, "m_year,m_month,amt,b_amt,month_amt" ) ;
	 
  			for(i=0;i<dbData.size();i++){
  				bean = (DataObject)dbData.get(i) ;
  				acc_code = (String)bean.getValue("acc_code");
  				amt = bean.getValue("amt").toString();
  				if(hasMonth.equals("true")){
  				   if(Integer.parseInt(S_MONTH) == 1){//95.03.16 1月份時.當月金額與累計金額相同
  				      month_amt = bean.getValue("amt").toString();
  				   }else{
  				      month_amt = bean.getValue("month_amt").toString();  				
  				   }
  				   A01_monthData.setProperty(acc_code,month_amt);
  				}
	         	A01Data.setProperty(acc_code,amt);
	         	
	       	    //System.out.println("acc_code="+acc_code+":amt="+amt);
	       	    //System.out.println("acc_code="+acc_code+":month_amt="+month_amt);
	        }
  			
  			List monthData = null;
  			List month_1Data = null;
  			sql.setLength(0) ;
  			paramList.clear() ;
  			if(hasMonth.equals("true")){//95.04.11 add 單月損益表.判斷本月跟上月有無申報資料 
  			    //sqlCmd = " select * from wml01 where bank_code='"+bank_code+"'"
  				//      + " and m_year="+S_YEAR 
	  		 	//      + " and m_month="+String.valueOf(Integer.parseInt(S_MONTH))
	  		 	//      + " and report_no='A01'";  			
  				sql.append(" select * from wml01 where bank_code= ? ");
  				sql.append(" and m_year= ? ") ;
  				sql.append(" and m_month= ? ") ;
  				sql.append(" and report_no= ? ") ;
  				paramList.add(bank_code) ;
  				paramList.add(S_YEAR) ;
  				paramList.add(S_MONTH) ;
  				paramList.add("A01") ;
  			   monthData = DBManager.QueryDB_SQLParam(sql.toString(), paramList, "")  ;
  				          //DBManager.QueryDB(sqlCmd,"");  	  	       
  			   sql.setLength(0) ;
  			   paramList.clear();
  			   //sqlCmd = " select * from wml01 where bank_code='"+bank_code+"'"
			   //   + " and m_year="+S_YEAR 
		 	   //   + " and m_month="+String.valueOf(Integer.parseInt(S_MONTH)-1)
		 	   //   + " and report_no='A01'";  			
  			   sql.append(" select * from wml01 where bank_code= ? ") ;
  			   sql.append(" and m_year= ? ") ;
  			   sql.append(" and m_month= ? ") ;
  			   sql.append(" and report_no= ? ") ;
  			   paramList.add(bank_code) ;
  			   paramList.add(lastYear) ;
               paramList.add(lastMonth) ;
  			   paramList.add("A01") ;
		       month_1Data = DBManager.QueryDB_SQLParam(sql.toString(), paramList, "") ; 
		    	            //DBManager.QueryDB(sqlCmd,"");
  	  	    }
  			
  			if(dbData.size() == 0){	
	  	  		row=sheet.getRow(0);
	  	  		cell=row.getCell((short)3);	       	
	  	  		//設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
	  	  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  	        if(bank_code.equals("ALL")){//全体漁會//95.02.10 fix by 2295
	  	  	       if(!hasMonth.equals("true")){//原損益表
	  		           cell.setCellValue("全體漁會信用部損益表");
	  	  	       }else{
	  	  	           cell.setCellValue("全體漁會信用部單月損益表");
	  	  	       }
	  		    }else{
	  		       if(!hasMonth.equals("true")){//原損益表
	           	      cell.setCellValue(BANK_NAME+"損益表");
	  		       }else{
	  		          cell.setCellValue(BANK_NAME+"單月損益表");
	  		       }
	  		    }   
	  	  	    row=sheet.getRow(1);
	  	  	    cell=row.getCell((short)3);
	  	  	    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
   	           	cell.setCellValue(S_YEAR +"年" +S_MONTH +"月無資料存在");
	  	  	}else{
	  	  		row=sheet.getRow(0);
	  	  		cell=row.getCell((short)3);	       	
	  	  		//設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
	  	  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  	  	    if(bank_code.equals("ALL")){//全体漁會//95.02.10 fix by 2295
	  	  	       if(!hasMonth.equals("true")){//原損益表
	  		           cell.setCellValue("全體漁會信用部損益表");
	  	  	       }else{
	  	  	           cell.setCellValue("全體漁會信用部單月損益表");
	  	  	       }
	  		    }else{
	  		       if(!hasMonth.equals("true")){//原損益表
   	           	      cell.setCellValue(BANK_NAME+"損益表");
	  		       }else{
	  		          cell.setCellValue(BANK_NAME+"單月損益表");
	  		       }
	  		    }   	  	  		
	  	  			       	
	  	  		YEAR = String.valueOf(Integer.parseInt(S_YEAR) + 1911);
                String s = Utility.getLastDay(YEAR +S_MONTH,"YYYYMM");
                String L_DAY = (Utility.getCHTdate(s.substring(0,10),1));
                            
                row=sheet.getRow(1);
	  	  		cell=row.getCell((short)3);
	  	  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  	  	    if(!hasMonth.equals("true")){//原損益表
   	           	   cell.setCellValue("中華民國" +S_YEAR +"年" +"1月1日至" +L_DAY);
	  	  	    }else{//單月損益表//95.04.11 add 單月損益表.判斷本月跟上月有無申報資料
	  	  	       //95.05.08 fix 單月損益表.上月or本月無資料時,顯示成中華民國xx年xx月(xx年xx月無資料) by 2295
		  	 	   if(!bank_code.equals("ALL")){//單一機構 
		  	 	       if(monthData.size() == 0){//目前這個月份無申報
		  	 	          monthExist = false;		    	     	  
		    	     	  cell.setCellValue("中華民國" +S_YEAR +"年" +S_MONTH+"月("+S_YEAR +"年" +S_MONTH+"月無資料)");
		  	 		   }else if(month_1Data.size() == 0){//上個月份無申報
		  	 		      monthExist = false;		  	 		      
		  	 		      cell.setCellValue("中華民國" +S_YEAR +"年" +S_MONTH+"月("+S_YEAR +"年"+String.valueOf(Integer.parseInt(S_MONTH)-1)+"月無資料)");
		  	 		   }else{    
		  	 		      cell.setCellValue("中華民國" +S_YEAR +"年" +S_MONTH+"月");
		  	 		   }
		  	 	   }else{
		  	 	       cell.setCellValue("中華民國" +S_YEAR +"年" +S_MONTH+"月");
		  	 	   }	
		  	 	   System.out.println("monthExist="+monthExist);
	  	  	    }
   	           	//列印單位//94.11.17 add by 2295
	            row=sheet.getRow(1);
	            //if(!hasMonth.equals("true")){//原損益表
	                cell=(row.getCell((short)9)==null)? row.createCell((short)9) : row.getCell((short)9);
	            //}else{
	            //    cell=(row.getCell((short)10)==null)? row.createCell((short)10) : row.getCell((short)10);
	            //}
	 	 		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	 	 		unit_name = Utility.getUnitName(unit) ;
	 	 		
	 	 		cell.setCellValue("單位：新台幣"+unit_name+"、％");
	  	  		//以巢狀迴圈讀取所有儲存格資料 
	  	  		System.out.println("total row ="+sheet.getLastRowNum());
	  	  		if(monthExist){//單月損益表.有一個月份無資料,則不顯示
	  	  		   for(i=4;i<31;i++){    	  	  		   
	  	  		       row=sheet.getRow(i);	    		
           		       cell=row.getCell((short)3);
           		       //System.out.print((int)cell.getNumericCellValue()+"=");
           		       if(!hasMonth.equals("true")){//原損益表
           		          amt = Utility.setCommaFormat(A01Data.getProperty(String.valueOf((int)cell.getNumericCellValue())));
           		       }else{//單月損益表
           		          amt = Utility.setCommaFormat(A01_monthData.getProperty(String.valueOf((int)cell.getNumericCellValue())));
           		       }
           		       cell=row.getCell((short)4);
           		       cell.setEncoding( HSSFCell.ENCODING_UTF_16 );	       			       		
           		       if(!amt.equals("0"))cell.setCellValue(amt);	//94.03.11 add if condition by egg
           		       cell=row.getCell((short)8);
           		       System.out.print((int)cell.getNumericCellValue()+"=");
           		       if(!hasMonth.equals("true")){//原損益表
           		          amt = Utility.setCommaFormat(A01Data.getProperty(String.valueOf((int)cell.getNumericCellValue())));
           		       }else{//單月損益表
           		          amt = Utility.setCommaFormat(A01_monthData.getProperty(String.valueOf((int)cell.getNumericCellValue())));
           		       }
           		       cell=row.getCell((short)9);
           		       if(!amt.equals("0"))cell.setCellValue(amt);	//94.03.11 add if condition by egg	  	  		    
	  	  		   }
	  	  	    }
	  	  		
	  	        //96.11.06 增加檢核結果與最後異動日期
	  	        sql.setLength(0) ;
	  	        paramList.clear() ;
		        //103.05.28 add 全体總表,若有檢核有誤資料,顯示農漁會信用部名稱
                if("ALL".equals(bank_code)){
                   //103.05.28 add 若有檢核有誤資料,顯示農漁會信用部名稱
                   String wml01Error = Utility.getWML01_Error(S_YEAR,S_MONTH,"7","A01");
                   if(!"".equals(wml01Error)){
                       row=sheet.getRow(34);                   
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
                       sheet.addMergedRegion( new Region( ( short )34, ( short )0,( short )34,( short ) 9 ) );      
                   }
                }else{
	  	           sql.append(" select UPD_CODE, to_char(UPDATE_DATE,'yyyymmdd') as UPDATE_DATE ") ;
	  	           sql.append(" from WML01 ") ;
	  	           sql.append(" where M_YEAR= ? ") ;
		  	       sql.append(" and M_MONTH=  ?") ;
		  	       sql.append(" and BANK_CODE= ? ") ;
		  	       sql.append(" and REPORT_NO= ? ") ;
		  	       paramList.add(S_YEAR);
		  	       paramList.add(S_MONTH);
		  	       paramList.add(bank_code);
		  	       paramList.add("A01");
   		           
		           dbData = DBManager.QueryDB_SQLParam(sql.toString(), paramList, "") ;		           	   
		           sql.setLength(0) ;
		           paramList.clear();
		           String UPD_CODE="";
		           String UPDATE_DATE="";
		           String M_YEAR="";   
		           String M_MONTH="";    
		           String M_DATE=""; 
		           if(dbData.size()>0){
		           	  //System.out.println("dbData.size()="+dbData.size()); 
		           	  bean = (DataObject) dbData.get(0) ;
		           	  UPD_CODE = (String)bean.getValue("upd_code");  
		           	  UPDATE_DATE = (String)bean.getValue("update_date");       		   
			          //System.out.println("UPD_CODE="+UPD_CODE); 
			          //System.out.println("UPDATE_DATE="+UPDATE_DATE); 
			          if(UPD_CODE.equals("N")) UPD_CODE="待檢核";
			          else if(UPD_CODE.equals("E")) UPD_CODE="檢核錯誤";
			          else if(UPD_CODE.equals("U")) UPD_CODE="檢核成功";
			          else UPD_CODE="";
			          //System.out.println("UPD_CODE="+UPD_CODE);
			          M_YEAR  = Integer.toString(Integer.parseInt(UPDATE_DATE.substring(0,4))-1911);	
			          M_MONTH  = Integer.toString(Integer.parseInt(UPDATE_DATE.substring(4,6))-0);	
			          M_DATE  = Integer.toString(Integer.parseInt(UPDATE_DATE.substring(6,8))-0);	
			          UPDATE_DATE=M_YEAR+"年"+M_MONTH+"月"+M_DATE+"日";	
			          //System.out.println("UPDATE_DATE="+UPDATE_DATE); 
		           }
		           row=sheet.getRow(34);
	               cell=row.getCell((short)0);
	               cell.setEncoding(HSSFCell.ENCODING_UTF_16);          
	               cell.setCellValue("檢核結果:"+(dbData.size()>0?UPD_CODE:"待檢核"));  
	                      
	               row=sheet.getRow(35);                   
	               cell=row.getCell((short)0);
	               cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	               cell.setCellValue("最後異動日期:"+(dbData.size()>0?UPDATE_DATE:"無")); 
                }//end of 總表列印檢核有信用部名稱	
	  	  	}//end of monthExist = true
	        FileOutputStream fout=new FileOutputStream(reportDir + System.getProperty("file.separator")+fileName);
	        wb.write(fout);
	        //儲存 
	        fout.close();
	        
		}catch(Exception e){
			System.out.println("createRpt Error:"+e+e.getMessage());
		}
		return errMsg;
	}		 
}       
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        