/*
	94.03.11 fix 金額資料為零是不輸出0，改為輸出空白及欄位資料右靠處理
	95.03.13 fix 年月區間sql by 2295 
	95.10.13 增加檢核結果與最後異動日期 BY 2495
   102.11.19 add 原QueryDB改套用QueryDB_SQLParam by 2295    
*/
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.poifs.filesystem.*;
import java.io.*;
import java.text.*;
import java.util.*;
import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;
import java.util.Calendar;

public class FR007W_Excel {
    public static String createRpt(String S_YEAR,String S_MONTH,String E_YEAR,String E_MONTH,String bank_type,String datestate,String Unit){
		String errMsg = "";
		List dbData = null;
		String sqlCmd = "";
		try{	
			System.out.println("農(漁)會資產品質分析彙總表.xls");
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
			FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+"農(漁)會資產品質分析彙總表.xls" );
			
	  	    //設定FileINputStream讀取Excel檔
	  		POIFSFileSystem fs = new POIFSFileSystem( finput );
	  		HSSFWorkbook wb = new HSSFWorkbook(fs);
	  		HSSFSheet sheet = wb.getSheetAt(0);//讀取第一個工作表，宣告其為sheet 
	  		HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
	        //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
	        //sheet.setAutobreaks(true); //自動分頁
			
	        //設定頁面符合列印大小
	        sheet.setAutobreaks( false );
	        ps.setScale( ( short )65 ); //列印縮放百分比
                ps.setLandscape( true ); // 設定橫式
	        ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
	  		//wb.setSheetName(0,"test");
	  		finput.close();
	  		
	  		HSSFRow row=null;//宣告一列 
	  		HSSFCell cell=null;//宣告一個儲存格  		
	  		
	  		short i=0;
	  		List paramList = new ArrayList();	  
	  		if (bank_type.equals("ALL")){
	  		   sqlCmd = "select m_year,m_month,"
                                  + "sum(decode(a04.acc_code,'840740',amt,0))/? as Field1,"
                                  + "sum(decode(a04.acc_code,'840750',amt,0))/? as Field2,"
                                  //+ "round(decode(sum(decode(a04.acc_code,'840750',amt,0)),0,0,sum(decode(a04.acc_code,'840740',amt,0))/sum(decode(a04.acc_code,'840750',amt,0))),4)"
                                  + "round(decode(sum(decode(a04.acc_code,'840750',amt,0)),0,0,sum(decode(a04.acc_code,'840740',amt,0))/sum(decode(a04.acc_code,'840750',amt,0)))*100,2)"
                                  + "as Field3,"
                                  + "sum(decode(a04.acc_code,'840760',amt,0))/? as Field4,"
                                  + "round(decode(sum(decode(a04.acc_code,'840750',amt,0)),0,0,"
                                  //+ "sum(decode(a04.acc_code,'840710',amt,'840720',amt,'840731',amt,'840732',amt,'840733',amt,'840734',amt,'840735',amt,0))/sum(decode(a04.acc_code,'840750',amt,0))),4)"
                                  + "sum(decode(a04.acc_code,'840760',amt,0))/sum(decode(a04.acc_code,'840750',amt,0)))*100,2)"
                                  + "as Field5,"
                                  + "round(decode(sum(decode(a04.acc_code,'840750',amt,0)),0,0,"
                                  //+ "sum(decode(a04.acc_code,'840740',amt,'840710',amt,'840720',amt,'840731',amt,'840732',amt,'840733',amt,'840734',amt,'840735',amt,0))/sum(decode(a04.acc_code,'840750',amt,0))),4)"
                                  + "sum(decode(a04.acc_code,'840740',amt,'840760',amt,0))/sum(decode(a04.acc_code,'840750',amt,0)))*100,2)"
                                  + "as Field6"
                                  //+ " from a04,bn01"
                                  + " from a04,(select distinct bn01.BANK_NO from cd01, wlx01, bn01 where cd01.hsien_id <> 'Y'  and wlx01.hsien_id=cd01.hsien_id  and wlx01.bank_no=bn01.bank_no   order by bn01.BANK_NO) bn01"
                                  + " where a04.bank_code = bn01.bank_no"
                                  + " and   to_char(m_year * 100 + m_month) >= ?" 
           					      + " and   to_char(m_year * 100 + m_month) <= ?"
                                  + " group by m_year,m_month";
	  		       paramList.add(Unit); 
	  		       paramList.add(Unit); 
	  		       paramList.add(Unit); 
	  		       paramList.add(S_YEAR+S_MONTH);
	  		       paramList.add(E_YEAR+E_MONTH);
	  		       dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"m_year,m_month,field1,field2,field3,field4,field5,field6");
	  	        }else{
	  		   sqlCmd = "select m_year,m_month,"
                                  + "sum(decode(a04.acc_code,'840740',amt,0))/? as Field1,"
                                  + "sum(decode(a04.acc_code,'840750',amt,0))/? as Field2,"
                                  //+ "round(decode(sum(decode(a04.acc_code,'840750',amt,0)),0,0,sum(decode(a04.acc_code,'840740',amt,0))/sum(decode(a04.acc_code,'840750',amt,0))),4)"
                                  + "round(decode(sum(decode(a04.acc_code,'840750',amt,0)),0,0,sum(decode(a04.acc_code,'840740',amt,0))/sum(decode(a04.acc_code,'840750',amt,0)))*100,2)"
                                  + "as Field3,"
                                  + "sum(decode(a04.acc_code,'840760',amt,0))/? as Field4,"
                                  + "round(decode(sum(decode(a04.acc_code,'840750',amt,0)),0,0,"
                                  //+ "sum(decode(a04.acc_code,'840710',amt,'840720',amt,'840731',amt,'840732',amt,'840733',amt,'840734',amt,'840735',amt,0))/sum(decode(a04.acc_code,'840750',amt,0))),4)"
                                  + "sum(decode(a04.acc_code,'840760',amt,0))/sum(decode(a04.acc_code,'840750',amt,0)))*100,2)"
                                  + "as Field5,"
                                  + "round(decode(sum(decode(a04.acc_code,'840750',amt,0)),0,0,"
                                  //+ "sum(decode(a04.acc_code,'840740',amt,'840710',amt,'840720',amt,'840731',amt,'840732',amt,'840733',amt,'840734',amt,'840735',amt,0))/sum(decode(a04.acc_code,'840750',amt,0))),4)"
                                  + "sum(decode(a04.acc_code,'840740',amt,'840760',amt,0))/sum(decode(a04.acc_code,'840750',amt,0)))*100,2)"
                                  + "as Field6"
                                  //+ " from a04,bn01"
                                  + " from a04,(select distinct bn01.BANK_NO from cd01, wlx01, bn01 where cd01.hsien_id <> 'Y'  and wlx01.hsien_id=cd01.hsien_id  and wlx01.bank_no=bn01.bank_no and bn01.bank_type = '" +bank_type+ "' order by bn01.BANK_NO) bn01"
                                  + " where a04.bank_code = bn01.bank_no"                                                                                             //全部就不要加這列,農會='6',漁會='7'
                                  + " and   to_char(m_year * 100 + m_month) >= ?" 
           					      + " and   to_char(m_year * 100 + m_month) <= ?"
                                  + " group by m_year,m_month";
	  		   paramList.add(Unit); 
	  		   paramList.add(Unit); 
	  		   paramList.add(Unit); 
	  		   paramList.add(S_YEAR+S_MONTH);
	  		   paramList.add(E_YEAR+E_MONTH);
	  		   dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"m_year,m_month,field1,field2,field3,field4,field5,field6");
	  	        }  	         

  	        if(dbData.size() == 0){	
	  		row=sheet.getRow(0);
	  		cell=row.getCell((short)5);	       	
	  		//設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
	  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
   	       	        cell.setCellValue("選擇之期間無資料存在");
	  	}else{
	  		row=sheet.getRow(0);
	  		cell=row.getCell((short)5);	       	
	  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
   	       	        System.out.println("bank_type ="+bank_type);
   	       	        if (bank_type.equals("6")){
   	       	           cell.setCellValue("農會資產品質分析彙總表");
   	       	        }else if (bank_type.equals("7")){
   	       	           cell.setCellValue("漁會資產品質分析彙總表");
   	       	        }else if (bank_type.equals("ALL")){
   	       	           cell.setCellValue("農漁會資產品質分析彙總表");
   	       	        }

	  		System.out.println("datestate ="+datestate);
	  		if (datestate.equals("1")){
	  		   row=sheet.getRow(1);
	  		   cell=row.getCell((short)0);
	  		   Calendar rightNow = Calendar.getInstance();
   	                   String year  = String.valueOf(rightNow.get(Calendar.YEAR)-1911);      //回覆值為西元年故需-1911取得民國年;
   	                   String month = String.valueOf(rightNow.get(Calendar.MONTH)+1);        //月份以0開始故加1取得實際月份;	       	
	  		   String day   = String.valueOf(rightNow.get(Calendar.DAY_OF_MONTH)); //日期以0開始故加1取得實際日期;
	  		   cell.setEncoding(HSSFCell.ENCODING_UTF_16);
   	       	           cell.setCellValue("列印日期:" +year +"年" +month +"月" +day +"日");
   	       	        }
	  		
	  		row=sheet.getRow(1);
	  		cell=row.getCell((short)12);	       	
	  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
   	       	        System.out.println("Unit ="+Unit);
   	       	        if (Unit.equals("1")){
   	       	           cell.setCellValue("元 ,%");
   	       	        }else if (Unit.equals("1000")){
   	       	           cell.setCellValue("千元 ,%");
   	       	        }else if (Unit.equals("10000")){
   	       	           cell.setCellValue("萬元 ,%");
   	       	        }else if (Unit.equals("1000000")){
   	       	           cell.setCellValue("百萬元 ,%");
   	       	        }else if (Unit.equals("10000000")){
   	       	           cell.setCellValue("千萬元 ,%");
   	       	        }else if (Unit.equals("100000000")){
   	       	           cell.setCellValue("億元 ,%");
   	       	        }
			
	  		//以巢狀迴圈讀取所有儲存格資料 
	  		System.out.println("dbData.size ="+dbData.size());
	  		for(i=0;i<dbData.size();i++){    		
	       		   row=sheet.getRow(2);
	       		   cell=row.getCell((short)(i+1));
	       		   String m_year = (((DataObject)dbData.get(i)).getValue("m_year")).toString();
	       		   String m_month = (((DataObject)dbData.get(i)).getValue("m_month")).toString();
	       		   cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	       		   cell.setCellValue(m_year+"年"+m_month+"月");
	       		   
	       		   System.out.println("I ="+i);
	       		   
	       		   row=sheet.getRow(3);
	       		   cell=row.getCell((short)(i+1));
	       		   String Field1 = Utility.setCommaFormat((((DataObject)dbData.get(i)).getValue("field1")).toString());
	       		   System.out.println("Field1 ="+Field1);
	       		   if(!Field1.equals("0"))cell.setCellValue(Field1);	//94.03.11 add if condition by egg
	       		   
	       		   row=sheet.getRow(4);
	       		   cell=row.getCell((short)(i+1));
	       		   String Field2 = Utility.setCommaFormat((((DataObject)dbData.get(i)).getValue("field2")).toString());
	       		   System.out.println("Field2 ="+Field2);
	       		   if(!Field2.equals("0"))cell.setCellValue(Field2);	//94.03.11 add if condition by egg
	       		   
	       		   row=sheet.getRow(5);
	       		   cell=row.getCell((short)(i+1));
	       		   String Field3 = (((DataObject)dbData.get(i)).getValue("field3")).toString();
	       		   System.out.println("Field3 ="+Field3);
	       		   if(!Field3.equals("0"))cell.setCellValue(Field3);	//94.03.11 add if condition by egg
	       		   
	       		   row=sheet.getRow(6);
	       		   cell=row.getCell((short)(i+1));
	       		   String Field4 = Utility.setCommaFormat((((DataObject)dbData.get(i)).getValue("field4")).toString());
	       		   System.out.println("Field4 ="+Field4);
	       		   if(!Field4.equals("0"))cell.setCellValue(Field4);	//94.03.11 add if condition by egg
	       		   
	       		   row=sheet.getRow(7);
	       		   cell=row.getCell((short)(i+1));
	       		   String Field5 = (((DataObject)dbData.get(i)).getValue("field5")).toString();
	       		   System.out.println("Field5 ="+Field5);
	       		   if(!Field5.equals("0"))cell.setCellValue(Field5);	//94.03.11 add if condition by egg
	       		   
	       		   row=sheet.getRow(8);
	       		   cell=row.getCell((short)(i+1));
	       		   String Field6 = (((DataObject)dbData.get(i)).getValue("field6")).toString();
	       		   System.out.println("Field6 ="+Field6);
	       		   if(!Field6.equals("0"))cell.setCellValue(Field6);	//94.03.11 add if condition by egg
	       		   
	  		}
	  		
	  		
	  	}

	        FileOutputStream fout=new FileOutputStream(reportDir + System.getProperty("file.separator")+"農(漁)會資產品質分析彙總表.xls");
	        wb.write(fout);
	        //儲存 
	        fout.close();
	        
		}catch(Exception e){
			System.out.println("createRpt Error:"+e+e.getMessage());
		}
		return errMsg;
	}		 
}
