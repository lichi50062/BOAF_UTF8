//94.06.24 add part 一.二若為"0"時.顯示"...
//             part三若為"0"時.顯示-"
//94.06.24 add 全國農業金庫目前還沒有值先insert一筆空的
//94.07.22 fix 更改農業信用保證基金業務統計(二).xls的樣式 by 2295
//02.11.19 add 原QueryDB改套用QueryDB_SQLParam by 2295    
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.poifs.filesystem.*;
import org.apache.poi.hssf.record.ExtendedFormatRecord;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import java.io.*;
import java.text.*;
import java.util.*;
import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR032W {
    public static String createRpt(String s_year,String s_month,String e_year,String e_month,String datestate){
		String errMsg = "";
		List dbData = null;
		String sqlCmd = "";
		Properties A01Data = new Properties();		
		String acc_code = "";
		String amt = "";
		reportUtil reportUtil = new reportUtil();
		try{	
			System.out.println("農業信用保證基金業務統計(二).xls");
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
			FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+"農業信用保證基金業務統計(二).xls" );
			System.out.println("Open excel 完成");
			
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
	        ps.setScale( ( short )100 ); //列印縮放百分比

	        ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
	  		//wb.setSheetName(0,"test");
	  		finput.close();
	  		
	  		HSSFRow row=null;//宣告一列 
	  		HSSFCell cell=null;//宣告一個儲存格  		
	  		
	  		short i=0;
	  		short y=0;	  
			row=(sheet.getRow(3)==null)? sheet.createRow(3) : sheet.getRow(3);
			insertCell(dbData,false,0,"中華民國　"+s_year+"年　"+s_month+"月",wb,row,(short)0,0);
			//加上列印日期			
			/*if(datestate.equals("1")){
				row=(sheet.getRow(2)==null)? sheet.createRow(2) : sheet.getRow(2);
				Calendar rightNow = Calendar.getInstance();
				String year = String.valueOf(rightNow.get(Calendar.YEAR)-1911);
				String month = String.valueOf(rightNow.get(Calendar.MONTH)+1);
				String day = String.valueOf(rightNow.get(Calendar.DAY_OF_MONTH));
				insertCell(dbData,false,0,"列印日期："+year+"年"+month+"月"+day+"日",wb,row,(short)0,0);
            }*/
			List paramList = new ArrayList();
			sqlCmd = "select m01.guarantee_item_no as item, "
			       + "       decode(substr(data_range,4,2),'YT','BLOCK1_0','ET','BLOCK1_0','YY','BLOCK1_1','TT','BLOCK1_1') as area_block, "
			       + "       output_order, "
			       + "       sum(decode(substr(data_range,4,2),'YT',guarantee_cnt,'YY',guarantee_cnt,0)) as guarantee_cnt_year, "
			       + "       round(sum(decode(substr(data_range,4,2),'YT',loan_amt,'YY',loan_amt,0))/1000000) as loan_amt_year, "
			       + "       round(sum(decode(substr(data_range,4,2),'YT',guarantee_amt,'YY',guarantee_amt,0))/1000000) as guarantee_amt_year, "
			       //+ "       round(sum(decode(substr(data_range,4,2),'YT',guarantee_bal,'YY',guarantee_bal,0))/1000000) as guarantee_bal_year, "
			       + "       round(sum(decode(substr(data_range,4,2),'TT',guarantee_bal,'ET',guarantee_bal,0))/1000000) as guarantee_bal_year," 
			       + "       round(sum(decode( "
			       + "                        (select sum(guarantee_bal) "
			       + "                           from m01  "
			       + "                          where lpad(m_year,3,'0')||lpad(m_month,2,'0') = lpad(?,3,'0') || lpad(?,2,'0')  "
			       //+ "                            and substr(data_range,4,2)='YT'),0,0, "
			       + "                            and substr(data_range,4,2)='ET'),0,0, " 
			       //+ "                        decode(substr(data_range,4,2),'YT',guarantee_bal,'YY',guarantee_bal,0) /  "
			       + "                          decode(substr(data_range,4,2),'ET',guarantee_bal,'TT',guarantee_bal,0) / "  
			       + "                        (select sum(guarantee_bal)  "
			       + "                           from m01  "
			       + "                          where lpad(m_year,3,'0')||lpad(m_month,2,'0') = lpad(?,3,'0') || lpad(?,2,'0')  "
			       //+ "                            and substr(data_range,4,2)='YT')) "
                               + "                            and substr(data_range,4,2)='ET')) " 
			       + "                       )*100,2) as guarantee_bal_stu_year, "
			       + "       sum(decode(substr(data_range,4,2),'ET',guarantee_cnt,'TT',guarantee_cnt,0)) as guarantee_cnt_totacc, "
			       + "       round(sum(decode(substr(data_range,4,2),'ET',loan_amt,'TT',loan_amt,0))/1000000) as loan_amt_totacc, "
			       + "       round(sum(decode(substr(data_range,4,2),'ET',guarantee_amt,'TT',guarantee_amt,0))/1000000) as guarantee_amt_totacc "
			       + "  from m01 join m00_guarantee_item on m01.guarantee_item_no=m00_guarantee_item.guarantee_item_no "
			       + " where lpad(m_year,3,'0')||lpad(m_month,2,'0') = lpad(?,3,'0') || lpad(?,2,'0')  "
			       + "   and substr(data_range,4,2) in ('YT','ET','YY','TT') "
			       + " group by m01.guarantee_item_no,decode(substr(data_range,4,2),'YT','BLOCK1_0','ET','BLOCK1_0','YY','BLOCK1_1','TT','BLOCK1_1'),output_order "
			       + " order by decode(substr(data_range,4,2),'YT','BLOCK1_0','ET','BLOCK1_0','YY','BLOCK1_1','TT','BLOCK1_1'),output_order ";
			paramList.add(s_year);
			paramList.add(s_month);
			paramList.add(s_year);
			paramList.add(s_month);
			paramList.add(s_year);
			paramList.add(s_month);
			dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"guarantee_cnt_year,loan_amt_year,guarantee_amt_year,guarantee_bal_year,guarantee_bal_stu_year,guarantee_cnt_totacc,loan_amt_totacc,guarantee_amt_totacc");	  	         
			System.out.println("layer 1 dbData.size()="+dbData.size());
			int j=0;
			short top=0,down=0;
  			for(i=0;i<dbData.size();i++){
  				j=1;
  				row=(sheet.getRow(i+8)==null)? sheet.createRow(i+8) : sheet.getRow(i+8);
  				insertCell(dbData,true,i,"guarantee_cnt_year",wb,row,(short)(j+1),1);j++;
  				insertCell(dbData,true,i,"loan_amt_year",wb,row,(short)(j+1),1);j++;
  				insertCell(dbData,true,i,"guarantee_amt_year",wb,row,(short)(j+1),1);j++;
  				insertCell(dbData,true,i,"guarantee_bal_year",wb,row,(short)(j+1),1);j++;
  				insertCell(dbData,true,i,"guarantee_bal_stu_year",wb,row,(short)(j+1),1);j++;
  				insertCell(dbData,true,i,"guarantee_cnt_totacc",wb,row,(short)(j+1),1);j++;
  				insertCell(dbData,true,i,"loan_amt_totacc",wb,row,(short)(j+1),1);j++;
  				insertCell(dbData,true,i,"guarantee_amt_totacc",wb,row,(short)(j+1),1);j++;
	       	}

  			paramList = new ArrayList();
			sqlCmd = "select m04.loan_use_no as item, "
			       + "       decode(m04.loan_use_no,'0','BLOCK2_1','BLOCK2_2') as area_block, "
			       + "       input_order, "
			       + "       guarantee_no_year as guarantee_cnt_year, "
			       + "       round(loan_amt_year/1000000) as loan_amt_year, "
			       + "       round(guarantee_amt_year/1000000) as guarantee_amt_year, "
			       + "       0 as guarantee_bal_year, "
			       + "       0 as guarantee_bal_stu_year, "
			       + "       guarantee_no_totacc as guarantee_cnt_totacc, "
			       + "       round(loan_amt_totacc/1000000) as loan_amt_totacc, "
			       + "       round(guarantee_amt_totacc/1000000) as guarantee_amt_totacc "
			       + "  from m04 join m00_loan_use on m04.loan_use_no=m00_loan_use.loan_use_no "
			       + " where lpad(m_year,3,'0')||lpad(m_month,2,'0') = lpad(?,3,'0') || lpad(?,2,'0')  "
			       + " order by decode(m04.loan_use_no,'0','BLOCK2_1','BLOCK2_2'),input_order ";
			paramList.add(s_year);
			paramList.add(s_month);
			dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"guarantee_cnt_year,loan_amt_year,guarantee_amt_year,guarantee_bal_year,guarantee_bal_stu_year,guarantee_cnt_totacc,loan_amt_totacc,guarantee_amt_totacc");
			System.out.println("layer 1 dbData.size()="+dbData.size());
			j=0;
			top=0;
			down=0;
  			for(i=0;i<dbData.size();i++){
  				j=1;
  				row=(sheet.getRow(i+17)==null)? sheet.createRow(i+17) : sheet.getRow(i+17);
  				insertCell(dbData,true,i,"guarantee_cnt_year",wb,row,(short)(j+1),2);j++;
  				insertCell(dbData,true,i,"loan_amt_year",wb,row,(short)(j+1),2);j++;
  				insertCell(dbData,true,i,"guarantee_amt_year",wb,row,(short)(j+1),2);j++;
  				insertCell(dbData,true,i,"guarantee_bal_year",wb,row,(short)(j+1),2);j++;
  				insertCell(dbData,true,i,"guarantee_bal_stu_year",wb,row,(short)(j+1),2);j++;
  				insertCell(dbData,true,i,"guarantee_cnt_totacc",wb,row,(short)(j+1),2);j++;
  				insertCell(dbData,true,i,"loan_amt_totacc",wb,row,(short)(j+1),2);j++;
  				insertCell(dbData,true,i,"guarantee_amt_totacc",wb,row,(short)(j+1),2);j++;
	       	}

  			paramList = new ArrayList();
			sqlCmd = "select m02.loan_unit_no as item, "
			       + "       decode(substr(data_range,4,2),'YT','BLOCK3_0','ET','BLOCK3_0','YY','BLOCK3_1','TT','BLOCK3_1') as area_block, "
			       //+ "       input_order, "
			       + "       output_order, "
			       + "       sum(decode(substr(data_range,4,2),'YT',guarantee_cnt,'YY',guarantee_cnt,0)) as guarantee_cnt_year, "
			       + "       round(sum(decode(substr(data_range,4,2),'YT',loan_amt,'YY',loan_amt,0))/1000000) as loan_amt_year, "
			       + "       round(sum(decode(substr(data_range,4,2),'YT',guarantee_amt,'YY',guarantee_amt,0))/1000000) as guarantee_amt_year, "
			       //+ "       round(sum(decode(substr(data_range,4,2),'YT',guarantee_bal,'YY',guarantee_bal,0))/1000000) as guarantee_bal_year, "
			       + "       round(sum(decode(substr(data_range,4,2),'TT',guarantee_bal,'ET',guarantee_bal,0))/1000000) as guarantee_bal_year, "
			       + "       round(sum(decode( "
			       + "                        (select sum(guarantee_bal)  "
			       + "                           from m02  "
			       + "                          where lpad(m_year,3,'0')||lpad(m_month,2,'0') = lpad(?,3,'0') || lpad(?,2,'0')  "
			       //+ "                            and substr(data_range,4,2)='YT'),0,0, "
			       + "                            and substr(data_range,4,2)='ET'),0,0, "
			       //+ "                        decode(substr(data_range,4,2),'YT',guarantee_bal,'YY',guarantee_bal,0) /  "
			       + "                        decode(substr(data_range,4,2),'ET',guarantee_bal,'TT',guarantee_bal,0) /  "
			       + "                        (select sum(guarantee_bal)  "
			       + "                           from m02  "
			       + "                          where lpad(m_year,3,'0')||lpad(m_month,2,'0') = lpad(?,3,'0') || lpad(?,2,'0')  "
			       //+ "                            and substr(data_range,4,2)='YT')) "
			       + "                            and substr(data_range,4,2)='ET')) "
			       + "                       )*100,2) as guarantee_bal_stu_year, "
			       + "       sum(decode(substr(data_range,4,2),'ET',guarantee_cnt,'TT',guarantee_cnt,0)) as guarantee_cnt_totacc, "
			       + "       round(sum(decode(substr(data_range,4,2),'ET',loan_amt,'TT',loan_amt,0))/1000000) as loan_amt_totacc, "
			       + "       round(sum(decode(substr(data_range,4,2),'ET',guarantee_amt,'TT',guarantee_amt,0))/1000000) as guarantee_amt_totacc "
			       + "  from m02 join m00_loan_unit on m02.loan_unit_no=m00_loan_unit.loan_unit_no "
			       + " where lpad(m_year,3,'0')||lpad(m_month,2,'0') = lpad(?,3,'0') || lpad(?,2,'0')  "
			       + "   and substr(data_range,4,2) in ('YT','ET','YY','TT') "
			       //+ " group by m02.loan_unit_no,decode(substr(data_range,4,2),'YT','BLOCK3_0','ET','BLOCK3_0','YY','BLOCK3_1','TT','BLOCK3_1'),input_order "
			       + " group by m02.loan_unit_no,decode(substr(data_range,4,2),'YT','BLOCK3_0','ET','BLOCK3_0','YY','BLOCK3_1','TT','BLOCK3_1'),output_order "
			       //+ " order by decode(substr(data_range,4,2),'YT','BLOCK3_0','ET','BLOCK3_0','YY','BLOCK3_1','TT','BLOCK3_1'),input_order ";
			       + " order by decode(substr(data_range,4,2),'YT','BLOCK3_0','ET','BLOCK3_0','YY','BLOCK3_1','TT','BLOCK3_1'),output_order ";
			paramList.add(s_year);
			paramList.add(s_month);
			paramList.add(s_year);
			paramList.add(s_month);
			paramList.add(s_year);
			paramList.add(s_month);
			
			dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"guarantee_cnt_year,loan_amt_year,guarantee_amt_year,guarantee_bal_year,guarantee_bal_stu_year,guarantee_cnt_totacc,loan_amt_totacc,guarantee_amt_totacc");
			System.out.println("layer 1 dbData.size()="+dbData.size());
			j=0;
			top=0;
			down=0;
  			for(i=0;i<dbData.size();i++){
  				System.out.println("i="+i);
  				j=1;  				
  				if(i>3){//94.06.24 add 一筆空的for全國農業金庫
  				    row=(sheet.getRow(i+1+40)==null)? sheet.createRow(i+1+40) : sheet.getRow(i+1+40);
  				}else{
  					row=(sheet.getRow(i+40)==null)? sheet.createRow(i+40) : sheet.getRow(i+40);	
  				}
				insertCell(dbData,true,i,"guarantee_cnt_year",wb,row,(short)(j+1),3);j++;
				insertCell(dbData,true,i,"loan_amt_year",wb,row,(short)(j+1),3);j++;
				insertCell(dbData,true,i,"guarantee_amt_year",wb,row,(short)(j+1),3);j++;
				insertCell(dbData,true,i,"guarantee_bal_year",wb,row,(short)(j+1),3);j++;
				insertCell(dbData,true,i,"guarantee_bal_stu_year",wb,row,(short)(j+1),3);j++;
				insertCell(dbData,true,i,"guarantee_cnt_totacc",wb,row,(short)(j+1),3);j++;
				insertCell(dbData,true,i,"loan_amt_totacc",wb,row,(short)(j+1),3);j++;
				insertCell(dbData,true,i,"guarantee_amt_totacc",wb,row,(short)(j+1),3);j++;				   
  				
  				//94.06.24 add 全國農業金庫目前還沒有值先insert一筆空的=========================
  				if(i==3){
  				   j=1;
  				   row=(sheet.getRow(i+1+40)==null)? sheet.createRow(i+1+40) : sheet.getRow(i+1+40);
  				   insertCell(dbData,false,i+1,"0",wb,row,(short)(j+1),3);j++;
  				   insertCell(dbData,false,i+1,"0",wb,row,(short)(j+1),3);j++;
  				   insertCell(dbData,false,i+1,"0",wb,row,(short)(j+1),3);j++;
				   insertCell(dbData,false,i+1,"0",wb,row,(short)(j+1),3);j++;
				   insertCell(dbData,false,i+1,"0",wb,row,(short)(j+1),3);j++;
  				   insertCell(dbData,false,i+1,"0",wb,row,(short)(j+1),3);j++;
  				   insertCell(dbData,false,i+1,"0",wb,row,(short)(j+1),3);j++;
				   insertCell(dbData,false,i+1,"0",wb,row,(short)(j+1),3);j++;
				   row=(sheet.getRow(i+1+40)==null)? sheet.createRow(i+1+40) : sheet.getRow(i+1+40);
				   continue;
  				}
  				//===========================================================================
  				row=(sheet.getRow(i+40)==null)? sheet.createRow(i+40) : sheet.getRow(i+40);
	       	}
  			paramList = new ArrayList();
			sqlCmd = "select m08.id_no as item, "
			       + "       decode(substr(data_range,4,2),'YT','BLOCK4_0','ET','BLOCK4_0','YY','BLOCK4_1','TT','BLOCK4_1') as area_block, "
			       + "       input_order, "
			       + "       sum(decode(substr(data_range,4,2),'YT',guarantee_no_month,'YY',guarantee_no_month,0)) as guarantee_cnt_year, "
			       + "       round(sum(decode(substr(data_range,4,2),'YT',loan_amt_month,'YY',loan_amt_month,0))/1000000) as loan_amt_year, "
			       + "       round(sum(decode(substr(data_range,4,2),'YT',guarantee_amt_month,'YY',guarantee_amt_month,0))/1000000) as guarantee_amt_year, "
			       + "       round(sum(decode(substr(data_range,4,2),'YT',guarantee_bal_month,'YY',guarantee_bal_month,0))/1000000) as guarantee_bal_year, "
			       + "       round(sum(decode(substr(data_range,4,2),'YT',guarantee_bal_p,'YY',guarantee_bal_p))/1000,2) as guarantee_bal_stu_year, "
			       + "       sum(decode(substr(data_range,4,2),'ET',guarantee_no_month,'TT',guarantee_no_month,0)) as guarantee_cnt_totacc, "
			       + "       round(sum(decode(substr(data_range,4,2),'ET',loan_amt_month,'TT',loan_amt_month,0))/1000000) as loan_amt_totacc, "
			       + "       round(sum(decode(substr(data_range,4,2),'ET',guarantee_amt_month,'TT',guarantee_amt_month,0))/1000000) as guarantee_amt_totacc "
			       + "  from m08 join m00_id_item on m08.id_no=m00_id_item.id_no "
			       + " where lpad(m_year,3,'0')||lpad(m_month,2,'0') = lpad(?,3,'0') || lpad(?,2,'0')  "
			       + "   and substr(data_range,4,2) in ('YT','ET','YY','TT') "
			       + " group by m08.id_no,decode(substr(data_range,4,2),'YT','BLOCK4_0','ET','BLOCK4_0','YY','BLOCK4_1','TT','BLOCK4_1'),input_order "
			       + " order by decode(substr(data_range,4,2),'YT','BLOCK4_0','ET','BLOCK4_0','YY','BLOCK4_1','TT','BLOCK4_1'),input_order ";
			//System.out.println("sqlCmd="+sqlCmd);
			paramList.add(s_year);
			paramList.add(s_month);
			dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"guarantee_cnt_year,loan_amt_year,guarantee_amt_year,guarantee_bal_year,guarantee_bal_stu_year,guarantee_cnt_totacc,loan_amt_totacc,guarantee_amt_totacc");
			System.out.println("m08 1 dbData.size()="+dbData.size());
			j=0;
			top=0;
			down=0;
  			for(i=0;i<dbData.size();i++){
  				j=1;
  				row=(sheet.getRow(i+53)==null)? sheet.createRow(i+53) : sheet.getRow(i+53);
  				insertCell(dbData,true,i,"guarantee_cnt_year",wb,row,(short)(j+1),4);j++;
  				insertCell(dbData,true,i,"loan_amt_year",wb,row,(short)(j+1),4);j++;
  				insertCell(dbData,true,i,"guarantee_amt_year",wb,row,(short)(j+1),4);j++;
  				insertCell(dbData,true,i,"guarantee_bal_year",wb,row,(short)(j+1),4);j++;
  				insertCell(dbData,true,i,"guarantee_bal_stu_year",wb,row,(short)(j+1),4);j++;
  				insertCell(dbData,true,i,"guarantee_cnt_totacc",wb,row,(short)(j+1),4);j++;
  				insertCell(dbData,true,i,"loan_amt_totacc",wb,row,(short)(j+1),4);j++;
  				insertCell(dbData,true,i,"guarantee_amt_totacc",wb,row,(short)(j+1),4);j++;
	       	}

	       	System.out.println(reportDir + System.getProperty("file.separator")+"農業信用保證基金業務統計(二).xls");
	        FileOutputStream fout=new FileOutputStream(reportDir + System.getProperty("file.separator")+"農業信用保證基金業務統計(二).xls");
	        wb.write(fout);
	        //儲存 
	        fout.close();
	        System.out.println("儲存完成");
		}catch(Exception e){
			System.out.println("createRpt Error:"+e+e.getMessage());
		}
		return errMsg;
	}		 
	
	public static void insertCell(List dbData,boolean getstate,int index,String Item,HSSFWorkbook wb,HSSFRow row,short j,int part)
	{
		String insertValue="";
  		if(getstate) insertValue= (((DataObject)dbData.get(index)).getValue(Item)).toString();
  		else         insertValue= Item;
		System.out.println("insertValue="+insertValue);
	    HSSFCell cell=(row.getCell(j)==null)? row.createCell(j) : row.getCell(j); 
	    cell.setEncoding( HSSFCell.ENCODING_UTF_16 );	       			       		
	    double value=0;
	    try{
	    	//94.06.24 add 一.二若為"0"時.顯示"..."
	    	//             三若為"0"時.顯示-" 
	    	if((part == 1 || part ==2) && insertValue.equals("0")){
	    		cell.setCellValue("             ...");
	    	}else if(part == 3 && insertValue.equals("0")){
	    		cell.setCellValue("           -");	
	    	}else{
	    	    cell.setCellValue(Double.parseDouble(insertValue));
	    	    System.out.println("double="+Double.parseDouble(insertValue));
	    	}
	    }catch(NumberFormatException e){
	    	cell.setCellValue(insertValue);
	    }
	}
}
