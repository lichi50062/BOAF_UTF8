/*
    94.09.27 add 明細表 by 2495
 * 102.11.19 add 原QueryDB改套用QueryDB_SQLParam by 2295    
*/
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.poifs.filesystem.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.hssf.record.ExtendedFormatRecord;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import java.io.*;
import java.text.*;
import java.util.*;
import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR030W {
  	public static String createRpt(String s_year,String s_month,String unit,String datestate,String bank_type)
	{
    	System.out.println("inpute s_month = "+s_month);
		String errMsg = "";
		String unit_name="";		
		int i=0;
		int j=0;
		String s_year_last="";
		String s_month_last="";
		String hsien_id_sum[]={"a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p"};		 
		String bank_type_name=(bank_type.equals("6"))?"農會":"漁會";
		
		
		String filename="";
		filename=(bank_type.equals("6"))?"全體農會信用部逾放比率超逾5%明細表.xls":"全體漁會信用部逾放比率超逾5%明細表.xls";
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
    		
    		
			FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+ filename );			
			
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
	        ps.setScale( ( short )90 ); //列印縮放百分比

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

			
			//列印年度
			row=(sheet.getRow(1)==null)? sheet.createRow(1) : sheet.getRow(1);													

			cell=row.getCell((short)0);			
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);			
			if(s_month.equals("0")){
				cell.setCellValue("  					中華民國　"+s_year+"　年度");
			}else {				
				
				cell.setCellValue("           				中華民國　"+s_year+"　年 " + s_month + "　月");							
			}
			
			//列印單位						
			cell=row.getCell((short)5);			
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);									
			if (unit.equals("1")){
			unit_name="元";
	   	    	}else if (unit.equals("1000")){
	   	    	unit_name="千元";
  	       		}else if (unit.equals("10000")){   	        	
			unit_name="萬元";
	   	    	}else if (unit.equals("1000000")){
	   	    	unit_name="百萬元";
	   	   	}else if (unit.equals("10000000")){
	   	    	unit_name="千萬元";
   	        	}else if (unit.equals("100000000")){
   	        	unit_name="億元";
	   	   	}		
			cell.setCellValue("單位：新台幣"+unit_name+"、％");

String combination =" select CNT_COUNT,CNT_TYPE_DIV,hsien_id ,bank_no ,BANK_NAME,field_DEBIT,field_CREDIT, field_DC_RATE ,field_OVER, field_OVER_RATE "		     
	            +" from (";
String summary	    =" select 0 as  CNT_COUNT, CNT_TYPE AS CNT_TYPE_DIV, CNT_TYPE, hsien_id, hsien_name,  FR001W_output_order,bank_no, BANK_NAME, field_DEBIT, field_CREDIT, field_DC_RATE,  field_OVER,  field_OVER_RATE "
                    +" from (";
List paramList_common = new ArrayList();
String common	    =" select (CASE WHEN (a01.OVER_RATE >= 0.05 and a01.OVER_RATE <= 0.1) THEN '10' "   
                    +" WHEN (a01.OVER_RATE > 0.1  and a01.OVER_RATE <= 0.15)    THEN '20' "    
		    +" WHEN (a01.OVER_RATE > 0.15 AND   a01.OVER_RATE <= 0.25)    THEN '30' "    
		    +" WHEN (a01.OVER_RATE > 0.25 AND   a01.OVER_RATE <=0.5)    THEN '40' "    
		    +" WHEN (a01.OVER_RATE  >  0.5)            THEN '50' "    
		    +" ELSE '00' END)  as CNT_TYPE,"  
      		    +" a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order, a01.bank_no, a01.BANK_NAME, " 
      		    +" round(a01.field_DEBIT/?,0)    as    field_DEBIT,"
   		    +" round(a01.field_CREDIT/?,0)   as    field_CREDIT,"
   		    +" decode(a01.fieldI_Y,0,0,"
    		    +" round((round((a01.fieldI_XA+"
	  	    +" decode(sign(a01.fieldI_XB1 - a01.fieldI_XB2),-1,0,(a01.fieldI_XB1 - a01.fieldI_XB2))+ "
          	    +" decode(sign(a01.fieldI_XC1 - a01.fieldI_XC2),-1,0,(a01.fieldI_XC1 - a01.fieldI_XC2))+ "
          	    +" decode(sign(a01.fieldI_XD1 - a01.fieldI_XD2),-1,0,(a01.fieldI_XD1 - a01.fieldI_XD2))+ "
          	    +" decode(sign(a01.fieldI_XE1 - a01.fieldI_XE2),-1,0,(a01.fieldI_XE1 - a01.fieldI_XE2))- "
          	    +" decode(sign(a01.fieldI_XF1 - a01.fieldI_XF2),-1,0,(a01.fieldI_XF1 - a01.fieldI_XF2))) "
		    +" /1,0))/a01.fieldI_Y * 100,2))        as    field_DC_RATE , "
   		    +" round(a01.field_OVER/?,0)            as    field_OVER, " 
   		    +" decode(a01.field_CREDIT,0,0, round(a01.field_OVER /  a01.field_CREDIT *100 ,2))  as   field_OVER_RATE "
		    +" from ("
		    +" select nvl(cd01.hsien_id,' ')       		       as  hsien_id , "
       		    +" nvl(cd01.hsien_name,'OTHER')                            as  hsien_name, "
       		    +" cd01.FR001W_output_order                                as  FR001W_output_order, "
                    +" bn01.bank_no ,  bn01.BANK_NAME,"
  		    +" sum(decode(a01.acc_code,'320300',amt,0))                as fieldL, " 
        	    +" sum(decode(a01.acc_code,'310000',amt,'320000',amt,0))   as fieldM, " 
                    +" sum(decode(a01.acc_code,'990000',amt,0))                as field_OVER, "
                    +" sum(decode(a01.acc_code,'220000',amt,0))                as field_DEBIT, ";
paramList_common.add(unit);
paramList_common.add(unit);
paramList_common.add(unit);
                    if(bank_type.equals("6")){                    
        	    	common +=" sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) as  field_CREDIT,";
		    }else{
  		        common +=" sum(decode(a01.acc_code,'120000',amt,'120800',amt,'121000',amt,'150300',amt,0)) as  field_CREDIT,";
                    }

  		    common +=" decode(sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)),0,0,"
                    	    +" round(sum(decode(a01.acc_code,'990000',amt,0))/sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)),4))  as OVER_RATE, ";

		    if(bank_type.equals("6")){
        	    	common +=" sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120200',amt,'120301',amt,'120302',amt,'120700',amt,'150200',amt,0)) as fieldI_XA, "
			        +" sum(decode(a01.acc_code,'120401',amt,'120402',amt,0))     as fieldI_XB1, "
				+" sum(decode(a01.acc_code,'240305',amt,'251300',amt,0))     as fieldI_XB2, "
				+" sum(decode(a01.acc_code,'120501',amt,'120502',amt,0))     as fieldI_XC1, "
				+" sum(decode(a01.acc_code,'240301',amt,'240302',amt,'240303',amt,'240304',amt,0))     as fieldI_XC2, "
				+" sum(decode(a01.acc_code,'120600',amt,0))                  as fieldI_XD1, "
				+" sum(decode(a01.acc_code,'240200',amt,0))                  as fieldI_XD2, "
				+" sum(decode(a01.acc_code,'150100',amt,0))                  as fieldI_XE1, "
				+" sum(decode(a01.acc_code,'250100',amt,0))                  as fieldI_XE2, "
				+" sum(decode(a01.acc_code,'310000',amt,'320000',amt,0))     as fieldI_XF1, "
				+" sum(decode(a01.acc_code,'140000',amt,0))                  as fieldI_XF2, ";		
		    }else{
  		    	common +=" sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120300',amt,'120401',amt,'120402',amt,'120700',amt,'150200',amt,0)) as fieldI_XA, "
				+" sum(decode(a01.acc_code,'120201',amt,'120202',amt,0))     as fieldI_XB1, "
				+" sum(decode(a01.acc_code,'240205',amt,0))                  as fieldI_XB2, "                 
                    		+" sum(decode(a01.acc_code,'120501',amt,'120502',amt,0))     as fieldI_XC1, "		    
                    		+" sum(decode(a01.acc_code,'240201',amt,'240202',amt,'240203',amt,'240204',amt,0))     as fieldI_XC2, "                  
                    		+" sum(decode(a01.acc_code,'120600',amt,0))                  as fieldI_XD1, "
  		    		+" sum(decode(a01.acc_code,'240200',amt,0))                  as fieldI_XD2, "
                    		+" sum(decode(a01.acc_code,'150100',amt,0))                  as fieldI_XE1, "
                   		+" sum(decode(a01.acc_code,'250100',amt,0))                  as fieldI_XE2, "
          	    		+" sum(decode(a01.acc_code,'310000',amt,'320000',amt,0))     as fieldI_XF1, "
  		    		+" sum(decode(a01.acc_code,'140000',amt,0))                  as fieldI_XF2, "; 
		    }
 
                    common +=" sum(decode(a01.acc_code,'220100',amt,'220200',amt,'220300',amt,'220400',amt,'220500',amt,'220600',amt,'220700',amt,'220800',amt,'220900',amt,'221000',amt,0))- "
      		    	    +" round(sum(decode(a01.acc_code,'220900',amt,0))/2,0)   as fieldI_Y "
		            +" from  a01,  bn01, wlx01 , cd01 " 
		            +" where (a01.m_year=  ? and a01.m_month  = ?) and ";
                    paramList_common.add(s_year);
                    paramList_common.add(s_month);
                     if(bank_type.equals("6")){          
      		    common +=" (wlx01.bank_no=bn01.bank_no  and bn01.bank_type='6') and ";
		    }else{
		    common +=" (wlx01.bank_no=bn01.bank_no  and bn01.bank_type='7') and ";
                    }
       		    common +=" wlx01.hsien_id=cd01.hsien_id                         and "
                            +"(a01.bank_code=bn01.bank_no ) "         
                            +" group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,  bn01.BANK_NAME) a01 "   
		            +" ) a01 where a01.CNT_TYPE <>  '00' ";
		
String subtotal  =" select COUNT(*) as  CNT_COUNT, "    
       		 +" decode(CNT_TYPE,'10','15','20','25','30','35','40','45', '50','55','65') AS CNT_TYPE_DIV,"
                 +" CNT_TYPE, "
		 +" ''  as  hsien_id, "
		 +" ''  as  hsien_name, "
		 +" ''  as  FR001W_output_order, "
		 +" ''  as  bank_no , "
                 +" ''  as  BANK_NAME, "  
	         +" sum(field_DEBIT)  as field_DEBIT, "
		 +" sum(field_CREDIT) as field_CREDIT, " 
	   	 +" 0  as field_DC_RATE , " 
	   	 +" sum(field_OVER) as field_OVER , "   
	         +" decode(sum(field_CREDIT),0,0,round(sum(field_OVER) /  sum(field_CREDIT) *100 ,2))  as   field_OVER_RATE "
		 +" from ( ";

String total   =" select COUNT(*) as  CNT_COUNT, "   
       		+" '99'  as CNT_TYPE_DIV, " 
                +" '99'  as  CNT_TYPE, "  
       		+" ''   as hsien_id , "
		+" ''   as  hsien_name, "
		+" ''   as  FR001W_output_order, "
       		+" ''   as bank_no , "
		+" ''   as  BANK_NAME, "  
	   	+" sum(field_DEBIT)  as field_DEBIT, "
		+" sum(field_CREDIT) as field_CREDIT, " 
	   	+" 0 as field_DC_RATE , "
	   	+" sum(field_OVER) as field_OVER , "   
	   	+" decode(sum(field_CREDIT),0,0,round(sum(field_OVER) /  sum(field_CREDIT) *100 ,2))  as   field_OVER_RATE "
 		+" from ( ";
List paramList = new ArrayList();
String sqlCmd= combination+summary+common
               +" UNION ALL "
               +  subtotal+common
               +" GROUP BY  CNT_TYPE "
               +" UNION ALL"
               +  total+common
               +" ) a01 where CNT_TYPE <>  '00' " 
	       +" order by  CNT_TYPE_DIV, field_OVER_RATE,  FR001W_output_order,  bank_no ";
for(int k=0;k<paramList_common.size();k++){
    paramList.add(paramList_common.get(k));
}
for(int k=0;k<paramList_common.size();k++){
    paramList.add(paramList_common.get(k));
}
for(int k=0;k<paramList_common.size();k++){
    paramList.add(paramList_common.get(k));
}
List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"cnt_count,cnt_type_div,hsien_id ,bank_no ,bank_name,field_debit,field_credit,field_dc_rate,field_over,field_over_rate");		
	 	
		 System.out.println("明細表----------------------------");
		 
		 System.out.println("dbData.size() ="+dbData.size());	
		 String  insertValue="",Befear_Precent="",After_Precent="";
		 int	id=0,Befear=0,After=5;
		 int rownum=4;
                 boolean summary_flag=true,subtotal_flag=true;
		 for(int rowcount=0;rowcount<dbData.size()-1;rowcount++)
		 {
			int cnt_type_div = Integer.parseInt((((DataObject)dbData.get(rowcount)).getValue("cnt_type_div")).toString());
			id = cnt_type_div%10;
			if(id==0)
			{
			    
			    for(int cellcount=0;cellcount<7;cellcount++)
			    {			    
				if(cellcount==0)
				{
					if(summary_flag)
			    		{
						Befear=Befear+5;
						After=After+5;
			       			After_Precent = Integer.toString(Befear);
			       			//cnt_type_div=cnt_type_div-5;
			       			Befear_Precent= Integer.toString(After);
			       			insertValue ="**"+After_Precent+"%~"+Befear_Precent+"%";
						insertCell(insertValue,rownum,0,wb,row,sheet,cell,0);
						for(int x=1;x<7;x++)
			    			{
							insertValue =" ";
							insertCell(insertValue,rownum,x,wb,row,sheet,cell,0);
						}	
                                		summary_flag=false;
						rownum++;
                            		}
					insertValue = (((DataObject)dbData.get(rowcount)).getValue("bank_no")).toString();
					insertCell(insertValue,rownum,cellcount,wb,row,sheet,cell,0);
			    	}
				if(cellcount==1)
				{
					insertValue = (((DataObject)dbData.get(rowcount)).getValue("bank_name")).toString();
					insertCell(insertValue,rownum,cellcount,wb,row,sheet,cell,0);
			    	}
				if(cellcount==2)
				{
					insertValue = (((DataObject)dbData.get(rowcount)).getValue("field_debit")).toString();
					insertCell(insertValue,rownum,cellcount,wb,row,sheet,cell,1);
				}
				if(cellcount==3)
				{
					insertValue = (((DataObject)dbData.get(rowcount)).getValue("field_credit")).toString();
					insertCell(insertValue,rownum,cellcount,wb,row,sheet,cell,1);
				}
				if(cellcount==4)
				{
					insertValue = (((DataObject)dbData.get(rowcount)).getValue("field_dc_rate")).toString();
					insertCell(insertValue,rownum,cellcount,wb,row,sheet,cell,2);
				}
				if(cellcount==5)
				{
					insertValue = (((DataObject)dbData.get(rowcount)).getValue("field_over")).toString();
					insertCell(insertValue,rownum,cellcount,wb,row,sheet,cell,1);
				}
				if(cellcount==6)
				{
					insertValue = (((DataObject)dbData.get(rowcount)).getValue("field_over_rate")).toString();
					insertCell(insertValue,rownum,cellcount,wb,row,sheet,cell,2);
				}
				subtotal_flag=true;
			    }
			}
			
			if(id==5)
			{
			    for(int cellcount=0;cellcount<7;cellcount++)
			    {
                                if(subtotal_flag)
				{
					if(cellcount==0)
					{
						insertValue = "小  計";
						insertCell(insertValue,rownum,cellcount,wb,row,sheet,cell,0);				
			    		}
					if(cellcount==1 || cellcount==4 )
					{
						insertValue = " ";
						insertCell(insertValue,rownum,cellcount,wb,row,sheet,cell,0);
			    		}
					if(cellcount==2)
					{
						insertValue = (((DataObject)dbData.get(rowcount)).getValue("field_debit")).toString();
						System.out.println("field_debit="+insertValue);
						insertCell(insertValue,rownum,cellcount,wb,row,sheet,cell,1);
					}
					if(cellcount==3)
					{
						insertValue = (((DataObject)dbData.get(rowcount)).getValue("field_credit")).toString();
						System.out.println("field_credit="+insertValue);
						insertCell(insertValue,rownum,cellcount,wb,row,sheet,cell,1);
					}
					if(cellcount==4 )
					{
						insertValue = " ";
						insertCell(insertValue,rownum,cellcount,wb,row,sheet,cell,2);
			    		}
					if(cellcount==5)
					{
						insertValue = (((DataObject)dbData.get(rowcount)).getValue("field_over")).toString();
						insertCell(insertValue,rownum,cellcount,wb,row,sheet,cell,1);
					}
					if(cellcount==6)
					{
						insertValue = (((DataObject)dbData.get(rowcount)).getValue("field_over_rate")).toString();
						insertCell(insertValue,rownum,cellcount,wb,row,sheet,cell,2);
						insertValue = "家 數";
						rownum++;
						cellcount=0;
						insertCell(insertValue,rownum,cellcount,wb,row,sheet,cell,0);
						cellcount=1;
						insertValue = (((DataObject)dbData.get(rowcount)).getValue("cnt_count")).toString();
						insertCell(insertValue,rownum,cellcount,wb,row,sheet,cell,0);
						for(cellcount=2;cellcount<7;cellcount++)
						{
							insertValue = " ";
							insertCell(insertValue,rownum,cellcount,wb,row,sheet,cell,1);
						}
						summary_flag=true;
					}
						
				}
			    }	
				subtotal_flag=false;			
			}
			
			rownum++;	
		  }
		
		//總 計
		for(int cellcount=0;cellcount<7;cellcount++)
	        {
			insertValue = " ";
			insertCell(insertValue,rownum,cellcount,wb,row,sheet,cell,0);	
		}		
	        rownum++;
		insertValue = "總 計";
		int cellcount=0;
		insertCell(insertValue,rownum,cellcount,wb,row,sheet,cell,0);
		cellcount=1;
		insertValue = " ";
		insertCell(insertValue,rownum,cellcount,wb,row,sheet,cell,0);
		cellcount=2;
		int x=dbData.size()-1;
		insertValue = (((DataObject)dbData.get(x)).getValue("field_debit")).toString();
		System.out.println("(((DataObject)dbData.get(dbData.size())).getValue(field_debit)).toString()="+insertValue);
		insertCell(insertValue,rownum,cellcount,wb,row,sheet,cell,1);
		cellcount=3;
		insertValue = (((DataObject)dbData.get(x)).getValue("field_credit")).toString();
		insertCell(insertValue,rownum,cellcount,wb,row,sheet,cell,1);
		cellcount=4;
		insertValue = " ";
		insertCell(insertValue,rownum,cellcount,wb,row,sheet,cell,0);
		cellcount=5;
		insertValue = (((DataObject)dbData.get(x)).getValue("field_over")).toString();
		insertCell(insertValue,rownum,cellcount,wb,row,sheet,cell,1);
		cellcount=6;
		insertValue = " ";
		insertCell(insertValue,rownum,cellcount,wb,row,sheet,cell,0);
		rownum++;
		insertValue = "家 數";
		cellcount=0;
		insertCell(insertValue,rownum,cellcount,wb,row,sheet,cell,0);
		cellcount=1;
		insertValue = (((DataObject)dbData.get(x)).getValue("cnt_count")).toString();
		insertCell(insertValue,rownum,cellcount,wb,row,sheet,cell,0);
		for(cellcount=2;cellcount<7;cellcount++)
	        {
			insertValue = " ";
			insertCell(insertValue,rownum,cellcount,wb,row,sheet,cell,1);	
		}
		
			
		HSSFFooter footer=sheet.getFooter();
	        footer.setCenter( "Page:" +HSSFFooter.page() +" of " +HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
		FileOutputStream fout=new FileOutputStream(reportDir+ System.getProperty("file.separator")+ filename);
		wb.write(fout);
	        //儲存 
	        fout.close();
	        System.out.println("儲存完成");
		}catch(Exception e){
			System.out.println("createRpt Error:"+e+e.getMessage());
		}
		return errMsg;
	}


	public static void insertCell(String insertValue,int rownum,int cellcount,HSSFWorkbook wb,HSSFRow row,HSSFSheet sheet,HSSFCell cell,int getstate)
	{
		System.out.println("insertValue="+insertValue);
		row=(sheet.getRow(rownum)==null)? sheet.createRow(rownum) : sheet.getRow(rownum);																
		cell=row.createCell((short)cellcount);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);			
		HSSFCellStyle cs1 = wb.createCellStyle();					
		cs1.setBorderTop(HSSFCellStyle.BORDER_THIN);
		cs1.setBorderBottom(HSSFCellStyle.BORDER_THIN); 
	       	cs1.setBorderLeft(HSSFCellStyle.BORDER_THIN);
	    	cs1.setBorderRight(HSSFCellStyle.BORDER_THIN);
		/*
		HSSFFont f = wb.createFont();
	    	f.setFontHeightInPoints((short)10);
                cs1.setFont(f);	
		f.setFontName("標楷體");
		*/
		if(getstate==1)
		{	
			cs1.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
			cell.setCellStyle(cs1);	
			cell.setCellValue(Utility.setCommaFormat(insertValue));
		}
		if(getstate==0)
		{
			cell.setCellValue(insertValue);
			cell.setCellStyle(cs1);	
		}
		if(getstate==2)
		{
			cs1.setAlignment(HSSFCellStyle.ALIGN_CENTER);
			cell.setCellStyle(cs1);	
			cell.setCellValue(insertValue);
		}
	}
	
}



