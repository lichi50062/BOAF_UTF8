/*
	2006/3/10  add 明細表 by 2495
	2011/02/11 fix 縣市合併 && sql injection by 2295
	2013.11.19 add 原QueryDB改套用QueryDB_SQLParam by 2295    
*/
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import java.io.*;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR007WA_BOAF{
  	public static String createRpt(String s_year,String s_month,String unit,String datestate,String bank_type,String bank_no,String bank_name)
	{
    System.out.println("inpute bank_no = "+bank_no);
		String errMsg = "";
		String unit_name="";		
		int i=0;
		int j=0;
		String s_year_last="";
		String s_month_last="";
		String hsien_id_sum[]={"a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p"};		 
		String bank_type_name=(bank_type.equals("6"))?"農會":"漁會";
		String cd01_table = "";
        String wlx01_m_year = "";
 		List paramList = new ArrayList(); 
		
		String filename="";
		filename="農漁會信用部資產品質分析明細表.xls";		
		
		
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
    		
    		//100.02.11 add 查詢年度100年以前.縣市別不同===============================
  	    	cd01_table = (Integer.parseInt(s_year) < 100)?"cd01_99":""; 
  	    	wlx01_m_year = (Integer.parseInt(s_year) < 100)?"99":"100"; 
  	    	//=====================================================================    
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
	        ps.setScale( ( short )70 ); //列印縮放百分比

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

			
			//列印表頭
			row=(sheet.getRow(0)==null)? sheet.createRow(0) : sheet.getRow(0);													
			cell=row.getCell((short)7);			
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);																		
			cell.setCellValue("          "+bank_name+"資產品質分析明細表 ");							
			
			
			//列印年度
			row=(sheet.getRow(1)==null)? sheet.createRow(1) : sheet.getRow(1);													
			cell=row.getCell((short)8);			
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);			
			if(s_month.equals("0")){
				cell.setCellValue("  中華民國　"+s_year+"　年度");
			}else {													
				cell.setCellValue("中華民國　"+s_year+"　年 " + s_month + "月");							
			}
			
			//列印單位						
			cell=row.getCell((short)18);			
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			unit_name = Utility.getUnitName(unit);//取得單位名稱					
			cell.setCellValue(" 單位：新台幣"+unit_name+"、％");		

		//抓農漁會信用部各類對象存放款比率資料	

String  sub_sqlCmd = " select    a04.hsien_id,a04.hsien_name,a04.FR001W_output_order,a04.bank_no ,a04.BANK_NAME, (CASE WHEN (field_OVER_A01 = field_OVER_A04 ) THEN '   ' ELSE '***'  END)  as OVER_TYPE_1, "
				   +" (CASE WHEN (field_CREDIT_A01 = field_CREDIT_A04 ) THEN '   '"                                
				   +" ELSE '***'  END)  as CREDIT_TYPE_1,"
				   +" round(field_OVER_A01 /?,0)    as field_OVER_A01,"
				   +" round(field_CREDIT_A01 /?,0)  as field_CREDIT_A01,"
				   +" decode(field_CREDIT_A01,0,0,round(field_OVER_A01 /  field_CREDIT_A01 *100 ,2))  as   field_OVER_RATE_A01,"
				   +" round(field_OVER_A04 /?,0)    as field_OVER_A04,"
				   +" round(field_CREDIT_A04 /?,0)  as field_CREDIT_A04,"
				   +" decode(field_CREDIT_A04,0,0,round(field_OVER_A04 /  field_CREDIT_A04 *100 ,2))  as   field_OVER_RATE_A04,"
				   +" round(field_840760_A04 /?,0)  as field_840760_A04, "
				   +" decode(field_CREDIT_A04,0,0,round(field_840760_A04 /  field_CREDIT_A04 *100 ,2))  as   field_RATE_840760_A04,"
				   +" decode(field_CREDIT_A04,0,0,round((field_OVER_A04 + field_840760_A04) /  field_CREDIT_A04 *100 ,2))  as   field_RATE_840760_OVER_A04,"
				   +" round(field_840710_A04_A /?,0)  as field_840710_A04_A,"                                        
				   +" round(field_840720_A04_B /?,0)  as field_840720_A04_B," 
				   +" round(field_840731_A04_a /?,0)  as field_840731_A04_a," 
				   +" round(field_840732_A04_b /?,0)  as field_840732_A04_b," 
				   +" round(field_840733_A04_c /?,0)  as field_840733_A04_c," 
				   +" round(field_840734_A04_d /?,0)  as field_840734_A04_d," 
				   +" round(field_840735_A04_e /?,0)  as field_840735_A04_e " 	  
				   +" from ("
				   +" select nvl(cd01.hsien_id,' ') as  hsien_id ,"
				   +" nvl(cd01.hsien_name,'OTHER')  as  hsien_name,"
				   +" cd01.FR001W_output_order      as  FR001W_output_order,"
				   +" bn01.bank_no , bn01.BANK_NAME, sum(decode(acc_code,'990000',amt,0))  as field_OVER_A01,"
				   +" sum(decode(acc_code,'120000',amt,'120800',amt,'150300',amt,0))  	as  field_CREDIT_A01,"
				   +" sum(decode(acc_code,'840740',amt,0))         as field_OVER_A04,"
	   			   +" sum(decode(acc_code,'840750',amt,0))         as field_CREDIT_A04,"
				   +" sum(decode(acc_code,'840760',amt,0))         as field_840760_A04,"
				   +" sum(decode(acc_code,'840710',amt,0))         as field_840710_A04_A,"
				   +" sum(decode(acc_code,'840720',amt,0))         as field_840720_A04_B,"
				   +" sum(decode(acc_code,'840731',amt,0))         as field_840731_A04_a,"
				   +" sum(decode(acc_code,'840732',amt,0))         as field_840732_A04_b,"
				   +" sum(decode(acc_code,'840733',amt,0))         as field_840733_A04_c,"
				   +" sum(decode(acc_code,'840734',amt,0))         as field_840734_A04_d,"
				   +" sum(decode(acc_code,'840735',amt,0))         as field_840735_A04_e "
				   +" from  (select * from cd01 where cd01.hsien_id <> 'Y') cd01"
				   +" left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id"
				   +" left join (select * from bn01 where m_year=?)bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=?"                
       		       +" left join ((select * from a04 where m_year  = ? and bank_code = ? and m_month  = ?) "
				   +" union" 
				   +" (select * from a01" 
				   +" where m_year  =  ? and bank_code = ?  and m_month  = ?    and" 
				   +" acc_code in('990000', '120000','120800','150300'))) a04"  
				   +" on  bn01.bank_no = a04.bank_code"
				   +" group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,"
				   +" bn01.bank_no ,  bn01.BANK_NAME) a04   where a04.bank_no <> ' '"
				   +" order by a04.hsien_id ,  a04.hsien_name,  a04.FR001W_output_order, a04.bank_no, a04.BANK_NAME";
        paramList.add(unit);
        paramList.add(unit);
        paramList.add(unit);
        paramList.add(unit);
        paramList.add(unit);
        paramList.add(unit);
        paramList.add(unit);
        paramList.add(unit);
        paramList.add(unit);
        paramList.add(unit);
        paramList.add(unit);
        paramList.add(unit);        
		paramList.add(wlx01_m_year);
		paramList.add(wlx01_m_year);
		paramList.add(bank_type);
		paramList.add(s_year);
		paramList.add(bank_no);
		paramList.add(s_month);
		paramList.add(s_year);
		paramList.add(bank_no);
		paramList.add(s_month);
		
		List dbData = DBManager.QueryDB_SQLParam(sub_sqlCmd,paramList,"field_over_a01,field_credit_a01,field_over_rate_a01,field_over_a04,field_credit_a04,field_over_rate_a04,field_840760_a04,field_rate_840760_a04,field_rate_840760_over_a04,field_840710_a04_a,field_840720_a04_b,field_840731_a04_a,field_840732_a04_b,field_840733_a04_c,field_840734_a04_d,field_840735_a04_e");		
		System.out.println("dbData.size() ="+dbData.size());
			System.out.println("明細表----------------------------");

			int count_over_type_1=0, count_credit_type_1=0;
			int rowNum=4;
			String  insertValue="";
			for(int rowcount=0;rowcount<dbData.size();rowcount++)
			{		
		    insertValue = (((DataObject)dbData.get(rowcount)).getValue("bank_no")).toString();
		    if(insertValue.equals(bank_no))
		    {
		    //System.out.println("rowcount="+rowcount);			       
		    for(int cellcount=0;cellcount<20;cellcount++)
				{	
					//System.out.println("cellcount="+cellcount);			
					if(cellcount==0)
					{					
						insertValue = (((DataObject)dbData.get(rowcount)).getValue("over_type_1")).toString();											
					  	//System.out.println("over_type_1="+insertValue);
						//2006_03_27 add 計算逾放不符家數 by 2495 
						if(insertValue.equals("***"))
						{
							count_over_type_1++; 
						}	
					}
					if(cellcount==1)	
					{			
						 insertValue = (((DataObject)dbData.get(rowcount)).getValue("credit_type_1")).toString();
						 //System.out.println("credit_type_1="+insertValue);
						//2006_03_27 add 計算逾放不符家數 by 2495 
						if(insertValue.equals("***"))
						{
							count_credit_type_1++; 
						}	
					}
					if(cellcount==2)
					{				
						insertValue = (((DataObject)dbData.get(rowcount)).getValue("bank_no")).toString();
						//System.out.println("bank_no="+insertValue);
					}
					if(cellcount==3)	
					{			
						insertValue = (((DataObject)dbData.get(rowcount)).getValue("bank_name")).toString();
						//System.out.println("bank_name="+insertValue);
					}
					if(cellcount==4)	
					{			
						insertValue = (((DataObject)dbData.get(rowcount)).getValue("field_over_a01")).toString();						
						//System.out.println("field_OVER_A01="+insertValue);
					}
					if(cellcount==5)
					{				
						insertValue = (((DataObject)dbData.get(rowcount)).getValue("field_credit_a01")).toString();
						//System.out.println("field_CREDIT_A01="+insertValue);
					}
					if(cellcount==6)
					{				
						insertValue = (((DataObject)dbData.get(rowcount)).getValue("field_over_rate_a01")).toString();
						//System.out.println("field_OVER_RATE_A01="+insertValue);
					}
					if(cellcount==7)
					{				
						insertValue = (((DataObject)dbData.get(rowcount)).getValue("field_over_a04")).toString();
						System.out.println("field_OVER_A04="+insertValue);
					}
					if(cellcount==8)
					{				
						insertValue = (((DataObject)dbData.get(rowcount)).getValue("field_credit_a04")).toString();
					  //System.out.println("field_CREDIT_A04="+insertValue);
					}
					if(cellcount==9)
					{				
						insertValue = (((DataObject)dbData.get(rowcount)).getValue("field_over_rate_a04")).toString();
					  //System.out.println("field_OVER_RATE_A04="+insertValue);
					}
					if(cellcount==10)
					{				
					  insertValue = (((DataObject)dbData.get(rowcount)).getValue("field_840760_a04")).toString();					  
                                          
                                          //System.out.println("field_840760_A04="+insertValue);
					}
					if(cellcount==11)
					{				
						insertValue = (((DataObject)dbData.get(rowcount)).getValue("field_rate_840760_a04")).toString();
						//System.out.println("field_RATE_840760_A04="+insertValue);
					}
					if(cellcount==12)
					{		
						insertValue = (((DataObject)dbData.get(rowcount)).getValue("field_rate_840760_over_a04")).toString();
						//System.out.println("field_RATE_840760_OVER_A04="+insertValue);
					}
					if(cellcount==13)	
					{	
						insertValue = (((DataObject)dbData.get(rowcount)).getValue("field_840710_a04_a")).toString();						
					  //System.out.println("field_840710_A04_A="+insertValue);
					}
					if(cellcount==14)		
					{
						insertValue = (((DataObject)dbData.get(rowcount)).getValue("field_840720_a04_b")).toString();
						//System.out.println("field_840720_A04_B="+insertValue);
					}
					if(cellcount==15)		
					{
						insertValue = (((DataObject)dbData.get(rowcount)).getValue("field_840731_a04_a")).toString();
					  	
						//System.out.println("field_840731_A04_a="+insertValue);
					}
					if(cellcount==16)
					{				
						insertValue = (((DataObject)dbData.get(rowcount)).getValue("field_840732_a04_b")).toString();
						//insertValue = (String)((DataObject)dbData.get(0)).getValue("field_840732_a04_b");
						//insertValue="1";
						//System.out.println("field_840732_A04_b="+insertValue);
					}
					if(cellcount==17)	
					{			
						insertValue = (((DataObject)dbData.get(rowcount)).getValue("field_840733_a04_c")).toString();
						//System.out.println("field_840733_A04_c="+insertValue);
					}
					if(cellcount==18)	
					{			
						insertValue = (((DataObject)dbData.get(rowcount)).getValue("field_840734_a04_d")).toString();
						//System.out.println("field_840734_A04_d="+insertValue);
					}
					if(cellcount==19)
					{				
						insertValue = (((DataObject)dbData.get(rowcount)).getValue("field_840735_a04_e")).toString();
						//System.out.println("field_840735_A04_e="+insertValue);
					}
					insertCell(insertValue,rowNum,cellcount,wb,row,sheet,cell);				
         }
        }				  		                       
			}
			//95.10.13 增加檢核結果與最後異動日期  BY 2495
			paramList.clear();
	  		String sqlCmd = " select UPD_CODE, to_char(UPDATE_DATE,'yyyymmdd') as UPDATE_DATE"
	  					  + " from WML01"		   		  
						  + " where M_YEAR=?"		   		   
						  + " and M_MONTH=?"
						  + " and BANK_CODE=?"
						  + " and REPORT_NO=?";
	  		paramList.add(s_year);
	  		paramList.add(s_month);
	  		paramList.add(bank_no);
	  		paramList.add("A04");
	  	   		   
			dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");	  
    		String UPD_CODE="";
    		String UPDATE_DATE="";
    		String M_YEAR="";   
    		String M_MONTH="";    
    		String M_DATE=""; 
    		if(dbData.size()>0){
       		   System.out.println("dbData.size()="+dbData.size()); 
       		    UPD_CODE = (String)((DataObject)dbData.get(0)).getValue("upd_code");  
       		    UPDATE_DATE = (String)((DataObject)dbData.get(0)).getValue("update_date");       		   
					   System.out.println("UPD_CODE="+UPD_CODE); 
					   System.out.println("UPDATE_DATE="+UPDATE_DATE); 
					   if(UPD_CODE.equals("N")) UPD_CODE="待檢核";
					   else if(UPD_CODE.equals("E")) UPD_CODE="檢核錯誤";
					   else if(UPD_CODE.equals("U")) UPD_CODE="檢核成功";
					   else UPD_CODE="";
					   System.out.println("UPD_CODE="+UPD_CODE);
					   M_YEAR  = Integer.toString(Integer.parseInt(UPDATE_DATE.substring(0,4))-1911);	
					   M_MONTH  = Integer.toString(Integer.parseInt(UPDATE_DATE.substring(4,6))-0);	
					   M_DATE  = Integer.toString(Integer.parseInt(UPDATE_DATE.substring(6,8))-0);	
					   UPDATE_DATE=M_YEAR+"年"+M_MONTH+"月"+M_DATE+"日";	
					   System.out.println("UPDATE_DATE="+UPDATE_DATE); 	
					   row=sheet.getRow(6);	    		
				    	cell=row.getCell((short)0);
				  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);	  		
				  		cell.setCellValue("檢核結果:"+UPD_CODE);	  
				  	  		
			        row=sheet.getRow(7);	    		
				    	cell=row.getCell((short)0);
				  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				  		cell.setCellValue("最後異動日期:"+UPDATE_DATE);	  		   
       		}else{
       			  row=sheet.getRow(6);	    		
				    	cell=row.getCell((short)0);
				  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);	  		
				  		cell.setCellValue("檢核結果:待檢核");	  
				  	  		
			        row=sheet.getRow(7);	    		
				    	cell=row.getCell((short)0);
				  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				  		cell.setCellValue("最後異動日期:無");	     
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
	
	public static void insertCell(String insertValue,int rowNum,int cellcount,HSSFWorkbook wb,HSSFRow row,HSSFSheet sheet,HSSFCell cell)
	{
		row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
			
		HSSFCellStyle cs1 = cell.getCellStyle();//會套用原本excel所設定的格式
		//HSSFCellStyle cs1 = wb.createCellStyle();															
		cell=row.createCell((short)cellcount);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);										
		cs1.setBorderTop(HSSFCellStyle.BORDER_THIN);
		cs1.setBorderBottom(HSSFCellStyle.BORDER_THIN); 
	       	cs1.setBorderLeft(HSSFCellStyle.BORDER_THIN);
	    	cs1.setBorderRight(HSSFCellStyle.BORDER_THIN);
		
		
		if(cellcount==3)
	        {
			//System.out.println("靠左cellcount=="+cellcount);
			cs1.setAlignment(HSSFCellStyle.ALIGN_LEFT);
			//System.out.println("insertValue=="+insertValue);
			cell.setCellStyle(cs1);
			cell.setCellValue(insertValue);						       
                }
		if(cellcount==0||cellcount==1||cellcount==2||cellcount==6||cellcount==9||cellcount==11||cellcount==12)
		{
			
			cs1.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
			//System.out.println("置中cellcount=="+cellcount);
			cell.setCellStyle(cs1);	
			cell.setCellValue(insertValue);		
		}
		if(cellcount==4||cellcount==5||cellcount==7||cellcount==8||cellcount==10||cellcount==13||cellcount==14||cellcount==15||cellcount==16||cellcount==17||cellcount==18||cellcount==19)
		{
			//System.out.println("靠右cellcount=="+cellcount);
			cs1.setAlignment(HSSFCellStyle.ALIGN_RIGHT);	
			cell.setCellStyle(cs1);
			cell.setCellValue(Utility.setCommaFormat(insertValue));
		}
		         	
	}
}



