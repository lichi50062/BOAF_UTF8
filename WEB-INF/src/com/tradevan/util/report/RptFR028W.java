/*	
	94.11.18 add 台灣地區農(漁)會信用部放款餘額表 by 2295
	94.11.21 add 台灣地區農會信用部放款餘額表 by 2295
	94.11.22 add 台灣地區漁會信用部放款餘額表 by 2295
	94.11.23 add remove 換頁 by 2295
	94.11.24 fix 查詢年月的sql by 2295
	95.03.14 fix 區分農漁會 by 2295
	99.11.03 fix 根據查詢年度.100年以後取得新縣市別.100年以前取得舊縣市別 
 		           使用PreparedStatement;並列印轉換後的SQL;套用QueryDB_SQLParam by 2295
   102.11.19 add 原QueryDB改套用QueryDB_SQLParam by 2295    
*/
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*;
import java.io.*;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR028W {
    public static String createRpt(String S_YEAR,String S_MONTH,String E_YEAR,String E_MONTH,String bank_type,String unit){
		String errMsg = "";
		List dbData = null;
		StringBuffer sqlCmd = new StringBuffer(); 
		//99.11.03 add 查詢年度100年以前.縣市別不同===============================
	    String cd01_table = (Integer.parseInt(S_YEAR) < 100)?"cd01_99":""; 
	    String wlx01_m_year = (Integer.parseInt(S_YEAR) < 100)?"99":"100"; 
	    //=====================================================================  
	    List paramList = new ArrayList();
		List A01List = new LinkedList();
		List YMList = new LinkedList();//儲存年月
		Properties A01Data = new Properties();		
		String acc_code = "";
		String amt = "";
		String unit_name="";
		String bank_type_name=(bank_type.equals("6"))?"農會":"漁會";
		String filename="";
		String ncacno="";
		DataObject bean = null;
		filename=(bank_type.equals("6"))?"台灣地區農會信用部放款餘額表.xls":"台灣地區漁會信用部放款餘額表.xls";		
		
		try{	
			System.out.println("台灣地區農會信用部放款餘額表.xls");
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
			FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+filename );
			
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
	  		int rowNum=2;//列印資料起始列
	  		ncacno = bank_type.equals("6")?"ncacno":"ncacno_7";
	  		sqlCmd.append(" select A01.m_year, A01.m_month, "+ncacno+".acc_range, a01.acc_code,");
			sqlCmd.append(" round(sum(amt)/?,0) as amt");	  			                       
			sqlCmd.append(" from A01 LEFT JOIN "+ncacno+" ON A01.acc_code = "+ncacno+".acc_code ");
			sqlCmd.append(",(select * from bn01 where m_year=?)bn01 ");
			sqlCmd.append(" where to_char(A01.m_year * 100 + A01.m_month) >= ?"); 
			sqlCmd.append(" and to_char(A01.m_year * 100 + A01.m_month) <= ?");
		    sqlCmd.append(" and (a01.acc_code like '12%' or a01.acc_code like '15%')");
		    sqlCmd.append(" and "+ncacno+".acc_div='01'");
		    sqlCmd.append(" and A01.bank_code = bn01.bank_no");
		    sqlCmd.append(" and bn01.bank_type=?");
		    sqlCmd.append(" group by A01.m_year, A01.m_month, "+ncacno+".acc_range, a01.acc_code");
		    sqlCmd.append(" order by A01.m_year, A01.m_month, "+ncacno+".acc_range, a01.acc_code");
		    paramList.add(unit);
		    paramList.add(wlx01_m_year);
		    paramList.add(S_YEAR+S_MONTH);
		    paramList.add(E_YEAR+E_MONTH);
		    paramList.add(bank_type);
	  		
	  		dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt"); 	         
              		    
	 	    row=sheet.getRow(1);
	  		cell=row.getCell((short)5);	       	
	  		//設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
	  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
  	        if(dbData.size() == 0){		  		
   	   	       	cell.setCellValue(S_YEAR +"年" +S_MONTH +"月至"+E_YEAR+"年"+E_MONTH+"月無資料存在");
	  	    }else{
	  	    	String t_year="",t_month="";
	  	    	bean = (DataObject)dbData.get(0);
	            t_year=(bean.getValue("m_year")).toString();
	            t_month=(bean.getValue("m_month")).toString();
	            
	  			for(i=0;i<dbData.size();i++){
	  				bean = (DataObject)dbData.get(i);
	  				if((!(bean.getValue("m_year")).toString().equals(t_year)) ||
	  				   (!(bean.getValue("m_month")).toString().equals(t_month))){
	  					A01List.add(A01Data);
	  					YMList.add(t_year+"/"+t_month);
	  					//System.out.println(t_year+"/"+t_month);
	  					A01Data = new Properties();		
	  					t_year=(bean.getValue("m_year")).toString();
	  					t_month=(bean.getValue("m_month")).toString();	  					
	  				}
	  				acc_code = (String)bean.getValue("acc_code");
	  				amt = (bean.getValue("amt")).toString();
		       	    A01Data.setProperty(acc_code,amt);
		       	    //System.out.println((bean.getValue("m_year")).toString()+"/"+(((DataObject)dbData.get(i)).getValue("m_month")).toString()+":acc_code="+acc_code+":amt="+amt);
		       	    if(i+1==dbData.size()){
		       	       t_year=(bean.getValue("m_year")).toString();
	  				   t_month=(bean.getValue("m_month")).toString();
		       	       A01List.add(A01Data);
		       	       YMList.add(t_year+"/"+t_month);
		       	       //System.out.println(t_year+"/"+t_month);
		       	    }
		        }
	  			
	  	    	System.out.println("A01List.size()="+A01List.size());
	  	    	System.out.println("YMList.size()="+YMList.size());
	  	    	cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
	  	    	cell.setCellValue(S_YEAR +"年" +S_MONTH +"月至"+E_YEAR+"年"+E_MONTH+"月");
	  	    	
	  	    	//列印單位		       	
		       	cell=(row.getCell((short)12)==null)? row.createCell((short)12) : row.getCell((short)12);   	 		
		 	   	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		 	   	
		 	   	unit_name=Utility.getUnitName(unit);
		 	   	   	    	
		 	   	cell.setCellValue("單位：新臺幣"+unit_name+"、％");
	  	    	
	  	    	
	  		  	//以巢狀迴圈讀取所有儲存格資料
	  	    	int oddidx=13;//單數column
	  	    	int evenidx=6;//雙數column
	  	    	int columnidx=0;	  	    	
	  	    	YMLoop:
	  	    	for(i=0;i<YMList.size();i++){//年月(i=單數在右邊 i=雙數在左邊)
	  	    		if(bank_type.equals("6")){
	  	    		   rowNum=2+20*(i/2);//農會	  	  
	  	    		}else{
	  	    		   rowNum=2+22*(i/2);//漁會
	  	    		}
	  	    		if(i % 2 == 0){//雙數
	  	    		   columnidx=evenidx;	  	    		   
	  	    		}else{
	  	    		   columnidx=oddidx;
	  	    		}
	  	    		//System.out.print("columnidx="+columnidx);
	  	    		//System.out.println(":rowNum="+rowNum);
	  	    	    //列印年月
	  	    		row=sheet.getRow(rowNum+1);	  	    		
	  	    		cell=row.getCell((short)(columnidx-5));	  	    		
	  	    		//System.out.println((String)YMList.get(i));	  	    		
	  	    		//System.out.println( ((String)YMList.get(i)).substring(0,((String)YMList.get(i)).indexOf("/") ) );
	  	    		//System.out.println(((String)YMList.get(i)).substring(3,((String)YMList.get(i)).length()));
	  	    		cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
	  	    		cell.setCellValue(((String)YMList.get(i)).substring(0,((String)YMList.get(i)).indexOf("/"))+"年" +((String)YMList.get(i)).substring(3,((String)YMList.get(i)).length())+"月");	  	    		
	  	    		cell=row.getCell((short)columnidx);
	  	    		A01Data = (Properties)A01List.get(i);
	  	    		//System.out.println((String)YMList.get(i)+":A01Data.size()="+A01Data.size());	  	    		
	  	    		
	  	    		for(int j=0;j<A01Data.size();j++){
	  	    			if((bank_type.equals("6")/*農會*/ && j==19) || (bank_type.equals("7")/*漁會*/ && j==21)){	  	    		    	
	  	    		    	continue YMLoop;
	  	    		    }
	  	    			row=sheet.getRow(rowNum);
	  	    			cell=row.getCell((short)(columnidx-1));	  	    			
	  	    		    //System.out.print((int)cell.getNumericCellValue()+"=");	  	    		    
	    	    	    amt = Utility.setCommaFormat(A01Data.getProperty(String.valueOf((int)cell.getNumericCellValue())));
	    	    	    cell=row.getCell((short)columnidx);
	    	    	    cell.setEncoding( HSSFCell.ENCODING_UTF_16 );	       			       		
	    	    	    if(!amt.equals("0"))cell.setCellValue(amt);
	    	    	    rowNum++;
	    	    	    row=sheet.getRow(rowNum);
	  	    		}//end of A01Data
	  	    	}//end of 年月  	        
	  	        //System.out.println("total row ="+sheet.getLastRowNum());
	  	    }//有資料
  	        
  	        //刪除無資料的表格
  	    	deleteCell(sheet,wb,row,cell,bank_type,YMList,unit_name);  	    	
  	        //設定涷結欄位
            //sheet.createFreezePane(0,1,0,1);
            footer.setCenter( "Page:" + HSSFFooter.page() + " of " +
                             HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa")); 
  	        
  	        FileOutputStream fout=new FileOutputStream(reportDir + System.getProperty("file.separator")+filename);
	        wb.write(fout);
	        //儲存 
	        fout.close();
	        
		}catch(Exception e){
			System.out.println("createRpt Error:"+e+e.getMessage());
		}
		return errMsg;
	}//end of createRpt
 
    public static void deleteCell(HSSFSheet sheet,HSSFWorkbook wb,HSSFRow row,HSSFCell cell,String bank_type,List YMList,String unit_name){
    	//刪除無資料的表格
    	try{
	        System.out.println("total row ="+sheet.getLastRowNum());	  	        
	        int oddidx=8;//單數column
	    	int evenidx=1;//雙數column
	    	int rowcount=0;
	    	int columnidx=0;
	    	int rowNum=0;
	    	boolean hideRow_left=false;
	    	boolean hideRow_right=false;
	    	if(YMList.size() <= 6){//只有一頁時,remove 換頁
	    		if(bank_type.equals("6")){
	    			sheet.removeRowBreak(61); 
	    			sheet.removeRowBreak(121);
	    		}
	    		if(bank_type.equals("7")){
	    			sheet.removeRowBreak(67);
	    			sheet.removeRowBreak(133);
	    		}
	    	}
	    	
	    	if(YMList.size() == 0 || YMList.size() == 1){
	    	   row=sheet.getRow(1);
	    	   //去除金額單位的底線
	    	   insertCell(null,false,0,"單位：新臺幣"+unit_name+"、％",wb,row,(short)12, (short)64,(short)0,(short)0,(short)0,(short)0);	  	    	   
	    	   insertCell(null,false,0,"單位：新臺幣"+unit_name+"、％",wb,row,(short)13, (short)64,(short)0,(short)0,(short)0,(short)0);	    	   
	    	}   
	        for(int i=0;i<12;i++){
	        	if(i % 2 == 0){//雙數
	        		hideRow_left=false;
	    	    	hideRow_right=false;	
  	   		        columnidx=evenidx;	  	    		   
  	   		    }else{
  	   		        columnidx=oddidx;
  	   		    }		  	   		
	        	if(bank_type.equals("6")){
	        	   rowNum=2+20*(i/2);//農會
	        	   rowcount=18;
	        	}else{
	        	   rowNum=2+22*(i/2);//漁會
	        	   rowcount=20;
	        	}
	        	
	        	//System.out.print("now.columnidx="+columnidx);
  	   			//System.out.println(":now.rowNum="+rowNum);
	        	row=sheet.getRow(rowNum+1);	        	
	        	
	        	cell=row.getCell((short)columnidx);//年月
	        	if(cell.getStringCellValue().equals("")){	        		
	        		if(i % 2 == 0){//雙數
	   	   		       hideRow_left=true;	  	    		   
	   	   		    }else{
	   	   		       hideRow_right=true;
	   	   		    }	
	        		//System.out.print("delete.columnidx="+columnidx);
	  	   			//System.out.println(":detete.rowNum="+rowNum);
	        		for(int j=rowNum;j<=rowNum+rowcount;j++){
	        			row=sheet.getRow(j);	        			
		    		    //delete cell	        			
			            for(int k=columnidx;k<=columnidx+5;k++){			            	
			            	cell=row.getCell((short)k);
			            	//insertCell(null,false,0," ",wb,row,(short)k, (short)64,(short)0,(short)0,(short)0,(short)0);			            	
			            	row.removeCell(cell);			            	
		                }//end of insert 空值表格
			            //若2個都為空值時,才隱藏row 
			            if(hideRow_left && hideRow_right){
			            	row.setHeight((short)0);		                
			            }
		  	        }//end of row
	        	}
	        }//end of 刪除空白資料
    	}catch(Exception e){
    		System.out.println("deleteCell Error:"+e+e.getMessage());
    	}
    }
    public static void insertCell(List dbData,boolean getstate,int index,String Item,HSSFWorkbook wb,HSSFRow row,short j, 
            						short bg,short bordertop,short borderbottom,short borderleft,short borderright)
    {
    	try{
    		String insertValue="";
    		if(getstate) insertValue= (((DataObject)dbData.get(index)).getValue(Item)).toString();
    		else         insertValue= Item;
    		/*
    		if(insertValue == null){
    			System.out.println(Item+" == null");
    		}else{
    			System.out.println(j+"insertValue="+insertValue);
    		}
    		*/
    		HSSFCell cell=(row.getCell(j)==null)? row.createCell(j) : row.getCell(j);
    		HSSFCellStyle cs1 = cell.getCellStyle();//會套用原本excel所設定的格式

    		cs1.setBorderTop(bordertop);
    		cs1.setBorderBottom(borderbottom);
    		cs1.setBorderLeft(borderleft);
    		cs1.setBorderRight(borderright);

    		cell.setCellStyle(cs1);
    		cell.setEncoding( HSSFCell.ENCODING_UTF_16 );    		
    		cell.setCellValue(insertValue);	
    	}catch(Exception e){
    		System.out.println("insertCell Error:"+e+e.getMessage());
    	}
    }//end of insertCell    
}
