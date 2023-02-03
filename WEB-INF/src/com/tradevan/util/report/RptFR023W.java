/*
 * 2005.09.27 add 明細表 by 2495
 * 2007.01.10 fix 列印單位 by 2295	
 * 2007.03.15 fix 逾期(含催收款)金額改為逾放金額(990000) by 2295
 * 2008.08.26 add 存放比率公式修正310800統一農貸公積 by 2295   
 *            add 將逾放金額區分成(會員/贊助會員/非會員/總額)
 *                [逾放金額-會員=逾期放款-正會員]
 * 				  [逾放金額-贊助會員=逾期放款-贊助會員]
 * 				  [逾放金額-非會員=逾期放款-非會員] by 2295
 * 2010.04.12 fix 1.因應縣市合併調整SQL 2.修改SQL查詢以pareparedStatement方式 by 2808 
 * 2012.12.28 fix 調整.機構名稱靠左 by 2295
 * 2013.01.14 fix 不顯示,已被裁撤的機構 by 2295
 * 2013.02.19 fix a02增加amt_name by 2295
 * 2013.06.03 fix sql by 2968
 * 2014.01.16 add 臺灣省農會更名為中華民國農會增加說明 by 2295
 * 2016.03.24 add 縣市政府列印只顯示其轄區下的有違反的農漁會信用部明細資料 by 2295    
*/
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import java.io.*;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR023W {
  	public static String createRpt(String s_year,String s_month,String unit,String datestate,String bank_type)
	{
    	System.out.println("inpute s_month = "+s_month);
    	int u_year = 99 ;
    	if(s_year==null || Integer.parseInt(s_year)<100) {
    		u_year = 99 ;
    	}else {
    		u_year = 100 ;
    	}
    	List paramList = new ArrayList();
    	List dbData = null;
    	DataObject bean = null;
		String errMsg = "";
		String unit_name="";		
		int i=0;
		int j=0;
		String s_year_last="";
		String s_month_last="";
		//String hsien_id_sum[]={"a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p"};		 
		String bank_type_name=(bank_type.equals("6") || bank_type.startsWith("B"))?"農會":"漁會";
		String hsien_id_b="";//縣市政府所屬縣市別 105.03.24 add
		
		String filename="";
		filename=(bank_type.equals("6") || bank_type.startsWith("B"))?"農會信用部各類對象存放款比率表.xls":"漁會信用部各類對象存放款比率表.xls";		
		filename="農會信用部各類對象存放款比率表.xls";
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
    		
    		if(bank_type.startsWith("B")){//105.03.24 add 縣市政府,則為全部轄區下的農漁會信用部                   
                hsien_id_b=bank_type.substring(2,bank_type.length());
                bank_type="ALL";
                System.out.println("RptFR023W.hsien_id="+hsien_id_b);
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
	        ps.setScale( ( short )65 ); //列印縮放百分比

	        ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
	  		//wb.setSheetName(0,"test");
	  		finput.close();
	  		
	  		
	  		HSSFCellStyle cs_left = wb.createCellStyle();
	        HSSFCellStyle cs2_right = wb.createCellStyle();
	        cs_left.setBorderTop((short)0);
	        cs_left.setBorderBottom(HSSFCellStyle.BORDER_THIN);
	        cs_left.setBorderRight(HSSFCellStyle.BORDER_THIN);
	        cs_left.setBorderLeft(HSSFCellStyle.BORDER_THIN);
	        cs_left.setAlignment(HSSFCellStyle.ALIGN_LEFT);
	        cs2_right.setBorderTop((short)0);
	        cs2_right.setBorderBottom(HSSFCellStyle.BORDER_THIN);
	        cs2_right.setBorderRight(HSSFCellStyle.BORDER_THIN);
	        cs2_right.setBorderLeft(HSSFCellStyle.BORDER_THIN);
	        cs2_right.setAlignment(HSSFCellStyle.ALIGN_RIGHT);    
	  		
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

	  	    StringBuffer sql = new StringBuffer () ;
	  	    sql.append(getReportSql(u_year,bank_type)) ;
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
		  	paramList.add(String.valueOf(u_year));
		  	if(bank_type.equals("ALL")){//105.03.24 add
		  	  paramList.add(hsien_id_b);
		  	}
		  	paramList.add(s_year);
		  	paramList.add(s_month);
		  	paramList.add(s_year);
		  	paramList.add(s_month);
		  	paramList.add(s_year);
		  	paramList.add(s_month);
		  	paramList.add(s_year);
		  	paramList.add(s_month);
		  	dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "field_A,field_B,field_C,field_D,field_E,field_F,field_G,field_H,field_I,field_j,field_K,field_L,field_992510,field_992520,field_992530"); 
		  	
		   //add 處理沒有資料 by 2495
		   System.out.println("dbData.size()="+dbData.size());
	       if(dbData.size()==0){	
			  row=(sheet.getRow(1)==null)? sheet.createRow(1) : sheet.getRow(1);
			  cell=row.getCell((short)4);			
			  cell.setEncoding(HSSFCell.ENCODING_UTF_16);	
   	       	  cell.setCellValue(s_year +"年" +s_month +"月無資料存在");
			  System.out.println(s_year +"年" +s_month +"月無資料存在");
           }else{
			  //列印年度
			  row=(sheet.getRow(1)==null)? sheet.createRow(1) : sheet.getRow(1);
			  cell=row.getCell((short)4);			
			  cell.setEncoding(HSSFCell.ENCODING_UTF_16);			
			  if(s_month.equals("0")){
			  	cell.setCellValue("  中華民國　"+s_year+"　年度");
			  }else{								
				cell.setCellValue("                                      中華民國　"+s_year+"　年 " + s_month + "月");							
			  }
			  System.out.println("unit="+unit);
			  //列印單位.96.01.10 fix 列印單位 by 2295						
			  cell=row.getCell((short)11);			
			  cell.setEncoding(HSSFCell.ENCODING_UTF_16);	
			  unit_name = Utility.getUnitName(unit) ;
			  
   	          System.out.println("unit_name="+unit_name);
			  cell.setCellValue("     單位：新台幣"+unit_name+"、％");	

			  System.out.println("明細表----------------------------");

			  int rowNum=4;
			  String  insertValue="";
			  for(int rowcount=0;rowcount<dbData.size();rowcount++){		
		          //System.out.println("rowcount="+rowcount);			       
		          for(int cellcount=0;cellcount<18;cellcount++){	
		        	bean = (DataObject) dbData.get(rowcount);
			  		if(cellcount==0)					
			  			insertValue = bean.getValue("hsien_name").toString();											
			  		if(cellcount==1)				
			  			insertValue = bean.getValue("bank_no").toString();
			  		if(cellcount==2)				
			  			insertValue = bean.getValue("bank_name").toString();
			  		if(cellcount==3)				
			  			insertValue = bean.getValue("field_a").toString();
			  		if(cellcount==4)				
			  			insertValue = bean.getValue("field_b").toString();
			  		if(cellcount==5)				
			  			insertValue = bean.getValue("field_c").toString();
			  		if(cellcount==6)				
			  			insertValue = bean.getValue("field_d").toString();
			  		if(cellcount==7)				
			  			insertValue = bean.getValue("field_e").toString();
			  		if(cellcount==8)				
			  			insertValue = bean.getValue("field_f").toString();
			  		if(cellcount==9)				
			  			insertValue = bean.getValue("field_g").toString();
			  		if(cellcount==10)				
			  			insertValue = bean.getValue("field_h").toString();
			  		if(cellcount==11)		
			  			insertValue = bean.getValue("field_i").toString();//存放比率
			  		if(cellcount==12)		
			  			insertValue = bean.getValue("field_992510") == null ? "" : bean.getValue("field_992510").toString();//逾期放款-正會員
			  		if(cellcount==13)		
			  			insertValue = bean.getValue("field_992520") == null ? "" : bean.getValue("field_992520").toString();//逾期放款-贊助會員
			  		if(cellcount==14)		
			  			insertValue = bean.getValue("field_992530") == null ? "" : bean.getValue("field_992530").toString();//逾期放款-非會員
			  		if(cellcount==15){
			  			//System.out.println((((DataObject)dbData.get(rowcount)).getValue("field_j")).toString());
			  			insertValue = bean.getValue("field_j").toString();//逾放金額
			  		}	
			  		if(cellcount==16)		
			  			insertValue = bean.getValue("field_k").toString();
			  		if(cellcount==17)		
			  			insertValue = bean.getValue("field_l").toString();
			  		
			  		insertCell(insertValue,rowNum,cellcount,wb,row,sheet,cell,rowcount,dbData,cs_left,cs2_right);				
                  }//end of cellcount
			  	rowNum++;                             
			  }//end of rowcount
			  //add 處理沒有資料的情況 by 2495
			  rowNum++;
              row = sheet.createRow(rowNum);
              cell = row.createCell( (short)0);
              cell.setEncoding(HSSFCell.ENCODING_UTF_16);
              cell.setCellValue("縣市別欄之其他(農會)係指原臺灣省農會，該農會於102年5月22日更名為中華民國農會");              
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
	
	public static void insertCell(String insertValue,int rowNum,int cellcount,HSSFWorkbook wb,HSSFRow row,HSSFSheet sheet,HSSFCell cell,int rowcount,List dbData,HSSFCellStyle cs_left,HSSFCellStyle cs_right)
	{
		row = sheet.createRow(rowNum);
		cell=row.createCell((short)cellcount);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		
		if(cellcount <= 2){
			//System.out.println("靠左cellcount=="+cellcount);
			cell.setCellValue(insertValue);			
			//cs1.setAlignment(HSSFCellStyle.ALIGN_LEFT);
			cell.setCellStyle(cs_left);
		}
		if(cellcount >= 3){
			//System.out.println("靠右cellcount=="+cellcount);
			cell.setCellValue(Utility.setCommaFormat(insertValue));
			//cs1.setAlignment(HSSFCellStyle.ALIGN_RIGHT);	
			cell.setCellStyle(cs_right);
		}
	}		
	/**
	 * 
	 * @param u_year
	 * @return
	 */
	private static String getReportSql(int u_year,String bank_type ) {
		StringBuffer sql = new StringBuffer () ;
		sql.append("select bank_type, hsien_id, hsien_name, FR001W_output_order, bank_no, BANK_NAME, ");      
		sql.append("        (field_DEBIT - (field_990420 +  field_990620))  as  field_A, ");          
		sql.append("        field_990420  as   field_B, ");                    
		sql.append("        field_990620  as   field_C, ");                    
		sql.append("        field_DEBIT   as   field_D, ");                    
		sql.append("        (field_CREDIT - (field_990410 +  field_990610))  as  field_E, ");         
		sql.append("        field_990410  as   field_F, ");                    
		sql.append("        field_990610  as   field_G, ");                    
		sql.append("        field_CREDIT  as   field_H, ");                    
		sql.append("        field_992510,field_992520,field_992530, ");        
		sql.append("        decode(a01.fieldI_Y,0,0,round((a01.fieldI_XA+ ");  
		sql.append("        decode(sign(a01.fieldI_XB1 - a01.fieldI_XB2),-1,0,(a01.fieldI_XB1 - a01.fieldI_XB2))+ ");        
		sql.append("        decode(sign(a01.fieldI_XC1 - a01.fieldI_XC2),-1,0,(a01.fieldI_XC1 - a01.fieldI_XC2))+ ");        
		sql.append("        decode(sign(a01.fieldI_XD1 - a01.fieldI_XD2),-1,0,(a01.fieldI_XD1 - a01.fieldI_XD2))+ ");        
		sql.append("        decode(sign(a01.fieldI_XE1 - a01.fieldI_XE2),-1,0,(a01.fieldI_XE1 - a01.fieldI_XE2))- ");        
		sql.append("        decode(sign(a01.fieldI_XF1 - a01.fieldI_XF3 - a01.fieldI_XF2),-1,0,(a01.fieldI_XF1 - a01.fieldI_XF3 - a01.fieldI_XF2))) ");  //--2008/08/26 存放比率公式修正310800統一農貸公積
		sql.append("               /a01.fieldI_Y * 100,2))     as     field_I, ");    
		sql.append("        field_OVER_T as   field_J, ");                     
		sql.append("        decode(a01.field_CREDIT,0,0,round(a01.field_OVER_T /  a01.field_CREDIT *100 ,2))  as   field_K, ");                       
		sql.append("        round(a01.field_91060P /  1000  ,2)     as   field_L ");                  
		sql.append(" from (select bn01.bank_type,   nvl(cd01.hsien_id,' ')  hsien_id, ");             
		sql.append("              nvl(cd01.hsien_name,'OTHER')  as  hsien_name, ");                   
		sql.append("              cd01.FR001W_output_order,bn01.bank_no ,  bn01.BANK_NAME, ");        
		sql.append("              round(sum(decode(a01.acc_code,'990420',amt,0)) /?,0) as  field_990420, ");                 
		sql.append("              round(sum(decode(a01.acc_code,'990620',amt,0)) /?,0) as  field_990620, ");                 
		sql.append("              round(sum(decode(a01.acc_code,'990410',amt,0)) /?,0) as  field_990410, ");                 
		sql.append("              round(sum(decode(a01.acc_code,'990610',amt,0)) /?,0) as  field_990610, ");                 
		sql.append("              round(sum(decode(a01.acc_code,'992510',amt,0)) /?,0) as  field_992510, ");//97.08.26 add 逾期放款-正會員   
		sql.append("              round(sum(decode(a01.acc_code,'992520',amt,0)) /?,0) as  field_992520, ");//97.08.26 add 逾期放款-贊助會員 
		sql.append("              round(sum(decode(a01.acc_code,'992530',amt,0)) /?,0) as  field_992530, ");//97.08.26 add 逾期放款-非會員
		sql.append("              sum(decode(a01.acc_code,'91060P',amt,0))         as  field_91060P, ");   
		sql.append("              round(sum(decode(a01.acc_code,'990000',amt,0)) /?,0) as  field_OVER_T, ");//96.03.15 fix 逾期(含催收款)金額改為逾放金額    
		sql.append("              round(sum(decode(a01.acc_code,'220000',amt,0)) /?,0)                  as field_DEBIT, ");  
		sql.append("              round(sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /?,0) as  field_CREDIT, ");
		sql.append("              decode(YEAR_TYPE,'102',decode(bank_type,'6',round(sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120200',amt,'120301',amt,'120302',amt,'120700',amt,'150200',amt,0)) /?,0),'7',round(sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120300',amt,'120401',amt,'120402',amt,'120700',amt,'150200',amt,0)) /?,0)), ");
		sql.append("                               '103', round(sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120200',amt,'120301',amt,'120302',amt,'120700',amt,'150200',amt,0)) /?,0) ,0 ) as fieldI_XA, ");           
		sql.append("              decode(YEAR_TYPE,'102',decode(bank_type,'6', round(sum(decode(a01.acc_code,'120401',amt,'120402',amt,0)) /?,0) ,'7',round(sum(decode(a01.acc_code,'120201',amt,'120202',amt,0)) /?,0)), ");
		sql.append("                               '103',round(sum(decode(a01.acc_code,'120401',amt,'120402',amt,0)) /?,0),0) as fieldI_XB1, ");   
		sql.append("              decode(YEAR_TYPE,'102',decode(bank_type,'6', round(sum(decode(a01.acc_code,'240305',amt,'251300',amt,'310800',amt, 0)) /?,0),'7',round(sum(decode(a01.acc_code,'240205',amt,'310800',amt, 0)) /?,0)), ");     
		sql.append("                               '103',decode(bank_type,'6',round(sum(decode(a01.acc_code,'240305',amt,'251300',amt,'310800',amt, 0)) /?,0),'7',round(sum(decode(a01.acc_code,'240305',amt,'251200',amt, 0)) /?,0)),0) as fieldI_XB2, ");            
		sql.append("              round(sum(decode(a01.acc_code,'120501',amt,'120502',amt,0)) /?,0)     as fieldI_XC1, ");   
		sql.append("              decode(YEAR_TYPE,'102',decode(bank_type,'6', round(sum(decode(a01.acc_code,'240301',amt,'240302',amt,'240303',amt,'240304',amt,0)) /?,0),'7',  round(sum(decode(a01.acc_code,'240201',amt,'240202',amt,'240203',amt,'240204',amt,0)) /?,0)), ");
		sql.append("                               '103',round(sum(decode(a01.acc_code,'240301',amt,'240302',amt,'240303',amt,'240304',amt,0)) /?,0) ,0)    as fieldI_XC2, ");                           
		sql.append("              round(sum(decode(a01.acc_code,'120600',amt,0)) /?,0)                  as fieldI_XD1, ");   
		sql.append("              decode(YEAR_TYPE,'102',decode(bank_type,'6', round(sum(decode(a01.acc_code,'240200',amt,0)) /?,0),'7',round(sum(decode(a01.acc_code,'240300',amt,0)) /?,0) ), ");
		sql.append("                               '103',round(sum(decode(a01.acc_code,'240200',amt,0)) /?,0),0)     as fieldI_XD2, ");   
		sql.append("              round(sum(decode(a01.acc_code,'150100',amt,0)) /?,0)                  as fieldI_XE1, ");   
		sql.append("              round(sum(decode(a01.acc_code,'250100',amt,0)) /?,0)                  as fieldI_XE2, ");   
		sql.append("              round(sum(decode(a01.acc_code,'310000',amt,'320000',amt,0)) /?,0)     as fieldI_XF1, ");   
		sql.append("              decode(YEAR_TYPE,'102', round(sum(decode(a01.acc_code,'310800',amt,0)) /?,0) , ");
		sql.append("                               '103',decode(bank_type,'6',round(sum(decode(a01.acc_code,'310800',amt,0)) /?,0),'7',0),0) as fieldI_XF3, ");//2008/08/26 存放比率公式修正310800統一農貸公積(修改後新增field_XF3)
		sql.append("              round(sum(decode(a01.acc_code,'140000',amt,0)) /?,0)                  as fieldI_XF2, ");   
		sql.append("              round((sum(decode(a01.acc_code,'220100',amt,'220200',amt,'220300',amt,'220400',amt,'220500',amt,'220600',amt,'220700',amt,'220800',amt,'220900',amt,'221000',amt,0))-  ");             
		sql.append("              round(sum(decode(a01.acc_code,'220900',amt,0))/2,0)) /?,0)   as fieldI_Y  ");              
		sql.append("       from  cd01 left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = ? ");
		if(bank_type.equals("ALL")){//105.03.24 add
		     sql.append(" and m2_name = ?");
		}
		sql.append("                  left join bn01 on wlx01.bank_no=bn01.bank_no   and wlx01.m_year = bn01.m_year and  bn01.bank_type in ('6', '7')  and bn01.bn_type <> '2' ");   
		sql.append("                  left join ((select  (CASE WHEN (a01.m_year <= 102) THEN '102' ");
		sql.append("                                             WHEN (a01.m_year > 102) THEN '103' ");
		sql.append("                                        ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt ");
		sql.append("                              from a01  where a01.m_year = ?   and a01.m_month= ? ) ");       
		sql.append("                              union all  ");       
		sql.append("                              (select (CASE WHEN (m_year <= 102) THEN '102' ");
		sql.append("                                            WHEN (m_year > 102) THEN '103' ");
		sql.append("                                       ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a02 where a02.m_year = ?   and a02.m_month= ?  and a02.acc_code in('990420','990620','990410', '990610')) ");                   
		sql.append("                              union all  ");                 
		sql.append("                              (select (CASE WHEN (m_year <= 102) THEN '102' ");
		sql.append("                                            WHEN (m_year > 102) THEN '103' ");
		sql.append("                                       ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a99 where a99.m_year = ?   and a99.m_month= ?  and a99.acc_code in('992510','992520','992530')) ");//97.08.26 add 逾期放款-正會員.贊助會員.非會員 
		sql.append("                              union all ");                  
		sql.append("                              (select  (CASE WHEN (m_year <= 102) THEN '102' ");
		sql.append("                                             WHEN (m_year > 102) THEN '103' ");
		sql.append("                                        ELSE '00' END) as YEAR_TYPE,m_year, m_month, bank_code, acc_code, amt  from a05 ");              
		sql.append("       where a05.m_year = ?  and a05.m_month= ?  and  a05.acc_code = '91060P')) a01 on a01.bank_code=bn01.bank_no ");          
		sql.append("       group by year_type,bn01.bank_type,nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order, bn01.bank_no ,bn01.BANK_NAME ");             
		sql.append("       )  a01  ");                     
		sql.append(" where  bank_type  in  ('6', '7') ");                      
		sql.append(" order by  bank_type,FR001W_output_order,hsien_id,bank_no ");   
		/*if(u_year > 99) {
			sql.append(" select bank_type, hsien_id, hsien_name, FR001W_output_order, bank_no, BANK_NAME,     ");
		  	sql.append("        (field_DEBIT - (field_990420 +  field_990620))  as  field_A,         ");
		  	sql.append("        field_990420  as   field_B,                   ");
		  	sql.append("        field_990620  as   field_C,                   ");
		  	sql.append("        field_DEBIT   as   field_D,                   ");
		  	sql.append("        (field_CREDIT - (field_990410 +  field_990610))  as  field_E,        ");
		  	sql.append("        field_990410  as   field_F,                   ");
		  	sql.append("        field_990610  as   field_G,                   ");
		  	sql.append("        field_CREDIT  as   field_H,                   ");
		  	sql.append("        field_992510,field_992520,field_992530,       ");
		  	sql.append("        decode(a01.fieldI_Y,0,0,round((a01.fieldI_XA+ ");
		  	sql.append("        decode(sign(a01.fieldI_XB1 - a01.fieldI_XB2),-1,0,(a01.fieldI_XB1 - a01.fieldI_XB2))+       ");
		  	sql.append("        decode(sign(a01.fieldI_XC1 - a01.fieldI_XC2),-1,0,(a01.fieldI_XC1 - a01.fieldI_XC2))+       ");
		  	sql.append("        decode(sign(a01.fieldI_XD1 - a01.fieldI_XD2),-1,0,(a01.fieldI_XD1 - a01.fieldI_XD2))+       ");
		  	sql.append("        decode(sign(a01.fieldI_XE1 - a01.fieldI_XE2),-1,0,(a01.fieldI_XE1 - a01.fieldI_XE2))-       ");
		  	sql.append("        decode(sign(a01.fieldI_XF1 - a01.fieldI_XF3 - a01.fieldI_XF2),-1,0,(a01.fieldI_XF1 - a01.fieldI_XF3 - a01.fieldI_XF2)))");// --2008/08/26 存放比率公式修正310800統一農貸公積
		  	sql.append(" 			  /a01.fieldI_Y * 100,2))     as     field_I,   ");
		  	sql.append("        field_OVER_T as   field_J,                    ");
		  	sql.append("        decode(a01.field_CREDIT,0,0,round(a01.field_OVER_T /  a01.field_CREDIT *100 ,2))  as   field_K, 	                 ");
		  	sql.append("        round(a01.field_91060P /  1000  ,2)     as   field_L                 ");
		  	sql.append(" from (select bn01.bank_type,   nvl(cd01.hsien_id,' ')  hsien_id,            ");
		  	sql.append("              nvl(cd01.hsien_name,'OTHER')  as  hsien_name,                  ");
		  	sql.append("              cd01.FR001W_output_order,bn01.bank_no ,  bn01.BANK_NAME,       ");
		  	sql.append("              round(sum(decode(a01.acc_code,'990420',amt,0)) /?,0) as  field_990420,                ");
		  	sql.append("              round(sum(decode(a01.acc_code,'990620',amt,0)) /?,0) as  field_990620,                ");
		  	sql.append("              round(sum(decode(a01.acc_code,'990410',amt,0)) /?,0) as  field_990410,                ");
		  	sql.append("              round(sum(decode(a01.acc_code,'990610',amt,0)) /?,0) as  field_990610,                ");
		  	sql.append("              round(sum(decode(a01.acc_code,'992510',amt,0)) /?,0) as  field_992510,      ");//--97.08.26 add 逾期放款-正會員   
		  	sql.append("              round(sum(decode(a01.acc_code,'992520',amt,0)) /?,0) as  field_992520,      ");//--97.08.26 add 逾期放款-贊助會員 
		  	sql.append("              round(sum(decode(a01.acc_code,'992530',amt,0)) /?,0) as  field_992530,         ");//--97.08.26 add 逾期放款-非會員
		  	sql.append("              sum(decode(a01.acc_code,'91060P',amt,0)) 		as  field_91060P,  ");
		  	sql.append("              round(sum(decode(a01.acc_code,'990000',amt,0)) /?,0) as  field_OVER_T, ");//--96.03.15 fix 逾期(含催收款)金額改為逾放金額	
		  	sql.append("              round(sum(decode(a01.acc_code,'220000',amt,0)) /?,0)                  as field_DEBIT, ");
		  	sql.append("              round(sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /?,0) as  field_CREDIT, 		         ");
		  	sql.append("              round(sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120200',amt,'120301',amt,'120302',amt,'120700',amt,'150200',amt,0)) /?,0)  as fieldI_XA,          ");
		  	sql.append("              round(sum(decode(a01.acc_code,'120401',amt,'120402',amt,0)) /?,0)     as fieldI_XB1,  ");
		  	sql.append("              round(sum(decode(a01.acc_code,'240305',amt,'251300',amt,'310800',amt, 0)) /?,0)     as fieldI_XB2,           ");
		  	sql.append("              round(sum(decode(a01.acc_code,'120501',amt,'120502',amt,0)) /?,0)     as fieldI_XC1,  ");
		  	sql.append("              round(sum(decode(a01.acc_code,'240301',amt,'240302',amt,'240303',amt,'240304',amt,0)) /?,0)     as fieldI_XC2,		                  ");
		  	sql.append("              round(sum(decode(a01.acc_code,'120600',amt,0)) /?,0)                  as fieldI_XD1,  ");
		  	sql.append("              round(sum(decode(a01.acc_code,'240200',amt,0)) /?,0)                  as fieldI_XD2,  ");
		  	sql.append("              round(sum(decode(a01.acc_code,'150100',amt,0)) /?,0)                  as fieldI_XE1,  ");
		  	sql.append("              round(sum(decode(a01.acc_code,'250100',amt,0)) /?,0)                  as fieldI_XE2,  ");
		  	sql.append("              round(sum(decode(a01.acc_code,'310000',amt,'320000',amt,0)) /?,0)     as fieldI_XF1,  ");
		  	sql.append("              round(sum(decode(a01.acc_code,'310800',amt,0)) /?,0)     			    as fieldI_XF3,       ");//--2008/08/26 存放比率公式修正310800統一農貸公積(修改後新增field_XF3)
		  	sql.append("              round(sum(decode(a01.acc_code,'140000',amt,0)) /?,0)                  as fieldI_XF2,  ");
		  	sql.append("              round((sum(decode(a01.acc_code,'220100',amt,'220200',amt,'220300',amt,'220400',amt,'220500',amt,'220600',amt,'220700',amt,'220800',amt,'220900',amt,'221000',amt,0))-             ");
		  	sql.append("              round(sum(decode(a01.acc_code,'220900',amt,0))/2,0)) /?,0)   as fieldI_Y              ");
		  	sql.append(" 	  from  cd01 left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = ?               ");
		  	sql.append("                 left join bn01 on wlx01.bank_no=bn01.bank_no   and wlx01.m_year = bn01.m_year and  bn01.bank_type in ('6', '7')  and bn01.bn_type <> '2'  ");
		  	sql.append("                 left join ((select * from a01  where a01.m_year = ?   and a01.m_month=  ? )      ");
		  	sql.append(" 						   union all       ");
		  	sql.append("     (select m_year,m_month,bank_code,acc_code,amt from a02 where a02.m_year = ?   and a02.m_month=  ?  and a02.acc_code in('990420','990620','990410', '990610'))                  ");
		  	sql.append("     union all                 ");
		  	sql.append("     (select * from a99 where a99.m_year = ?   and a99.m_month=  ?  and a99.acc_code in('992510','992520','992530'))    ");//--97.08.26 add 逾期放款-正會員.贊助會員.非會員 
		  	sql.append("     union all                 ");
		  	sql.append("     (select m_year, m_month, bank_code, acc_code, amt  from a05             ");
		  	sql.append(" 	  where a05.m_year =  ?   and a05.m_month=  ?  and  a05.acc_code = '91060P')) a01 on a01.bank_code=bn01.bank_no         ");
		  	sql.append(" 	  group by bn01.bank_type,nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order, bn01.bank_no ,bn01.BANK_NAME            ");
		  	sql.append(" 	  )  a01                     ");
		  	sql.append(" where  bank_type  in  ('6', '7')                     ");
		  	sql.append(" order by  bank_type,FR001W_output_order,hsien_id,bank_no                    ");
		}else {
			sql.append(" select bank_type, hsien_id, hsien_name, FR001W_output_order, bank_no, BANK_NAME,             ");
			sql.append("        (field_DEBIT - (field_990420 +  field_990620))  as  field_A,                          ");
			sql.append("        field_990420  as   field_B,                                                           ");
			sql.append("        field_990620  as   field_C,                                                           ");
			sql.append("        field_DEBIT   as   field_D,                                                           ");
			sql.append("        (field_CREDIT - (field_990410 +  field_990610))  as  field_E,                         ");
			sql.append("        field_990410  as   field_F,                                                           ");
			sql.append("        field_990610  as   field_G,                                                           ");
			sql.append("        field_CREDIT  as   field_H,                                                           ");
			sql.append("        field_992510,field_992520,field_992530,                                               ");
			sql.append("        decode(a01.fieldI_Y,0,0,round((a01.fieldI_XA+                                         ");
			sql.append("        decode(sign(a01.fieldI_XB1 - a01.fieldI_XB2),-1,0,(a01.fieldI_XB1 - a01.fieldI_XB2))+  ");
			sql.append("        decode(sign(a01.fieldI_XC1 - a01.fieldI_XC2),-1,0,(a01.fieldI_XC1 - a01.fieldI_XC2))+  ");
			sql.append("        decode(sign(a01.fieldI_XD1 - a01.fieldI_XD2),-1,0,(a01.fieldI_XD1 - a01.fieldI_XD2))+  ");
			sql.append("        decode(sign(a01.fieldI_XE1 - a01.fieldI_XE2),-1,0,(a01.fieldI_XE1 - a01.fieldI_XE2))-  ");
			sql.append("        decode(sign(a01.fieldI_XF1 - a01.fieldI_XF3 - a01.fieldI_XF2),-1,0,(a01.fieldI_XF1 - a01.fieldI_XF3 - a01.fieldI_XF2)))");// --2008/08/26 存放比率公式修正310800統一農貸公積   
			sql.append(" 			  /a01.fieldI_Y * 100,2))     as     field_I,                                                              ");
			sql.append("        field_OVER_T as   field_J,                                                                                             ");
			sql.append("        decode(a01.field_CREDIT,0,0,round(a01.field_OVER_T /  a01.field_CREDIT *100 ,2))  as   field_K, 	                   ");
			sql.append("        round(a01.field_91060P /  1000  ,2)     as   field_L                                                                   ");
			sql.append(" from (select bn01.bank_type,   nvl(cd01.hsien_id,' ')  hsien_id,                                                              ");
			sql.append("              nvl(cd01.hsien_name,'OTHER')  as  hsien_name,                                                                    ");
			sql.append("              cd01.FR001W_output_order,bn01.bank_no ,  bn01.BANK_NAME,                                                         ");
			sql.append("              round(sum(decode(a01.acc_code,'990420',amt,0)) /?,0) as  field_990420,                                           ");
			sql.append("              round(sum(decode(a01.acc_code,'990620',amt,0)) /?,0) as  field_990620,                                           ");
			sql.append("              round(sum(decode(a01.acc_code,'990410',amt,0)) /?,0) as  field_990410,                                           ");
			sql.append("              round(sum(decode(a01.acc_code,'990610',amt,0)) /?,0) as  field_990610,                                           ");
			sql.append("              round(sum(decode(a01.acc_code,'992510',amt,0)) /?,0) as  field_992510,             ");//--97.08.26 add 逾期放款-正會員
			sql.append("              round(sum(decode(a01.acc_code,'992520',amt,0)) /?,0) as  field_992520,           ");//--97.08.26 add 逾期放款-贊助會員
			sql.append("              round(sum(decode(a01.acc_code,'992530',amt,0)) /?,0) as  field_992530,             ");//--97.08.26 add 逾期放款-非會員
			sql.append("              sum(decode(a01.acc_code,'91060P',amt,0)) 		as  field_91060P,                                          ");
			sql.append("              round(sum(decode(a01.acc_code,'990000',amt,0)) /?,0) as  field_OVER_T,	");	 // --96.03.15 fix 逾期(含催收款)金額改為逾放金額          		          
			sql.append("              round(sum(decode(a01.acc_code,'220000',amt,0)) /?,0)                  as field_DEBIT,                                 ");
			sql.append("              round(sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /?,0) as  field_CREDIT, 		        ");           
			sql.append("              round(sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120200',amt,'120301',amt,'120302',amt,'120700',amt,'150200',amt,0)) /?,0)  as fieldI_XA,  ");
			sql.append("              round(sum(decode(a01.acc_code,'120401',amt,'120402',amt,0)) /?,0)     as fieldI_XB1,                                                                  ");
			sql.append("              round(sum(decode(a01.acc_code,'240305',amt,'251300',amt,'310800',amt, 0)) /?,0)     as fieldI_XB2,                                                    ");
			sql.append("              round(sum(decode(a01.acc_code,'120501',amt,'120502',amt,0)) /?,0)     as fieldI_XC1,                                                                  ");
			sql.append("              round(sum(decode(a01.acc_code,'240301',amt,'240302',amt,'240303',amt,'240304',amt,0)) /?,0)     as fieldI_XC2,		                        ");
			sql.append("              round(sum(decode(a01.acc_code,'120600',amt,0)) /?,0)                  as fieldI_XD1,                                                                  ");
			sql.append("              round(sum(decode(a01.acc_code,'240200',amt,0)) /?,0)                  as fieldI_XD2,                                                                  ");
			sql.append("              round(sum(decode(a01.acc_code,'150100',amt,0)) /?,0)                  as fieldI_XE1,                                                                  ");
			sql.append("              round(sum(decode(a01.acc_code,'250100',amt,0)) /?,0)                  as fieldI_XE2,                                                                  ");
			sql.append("              round(sum(decode(a01.acc_code,'310000',amt,'320000',amt,0)) /?,0)     as fieldI_XF1,                                                                  ");
			sql.append("              round(sum(decode(a01.acc_code,'310800',amt,0)) /?,0)     			    as fieldI_XF3,   ");//  --2008/08/26 存放比率公式修正310800統一農貸公積(修改後新增field_XF3)
			sql.append("              round(sum(decode(a01.acc_code,'140000',amt,0)) /?,0)                  as fieldI_XF2,                                                                                     ");
			sql.append("              round((sum(decode(a01.acc_code,'220100',amt,'220200',amt,'220300',amt,'220400',amt,'220500',amt,'220600',amt,'220700',amt,'220800',amt,'220900',amt,'221000',amt,0))-    ");
			sql.append("              round(sum(decode(a01.acc_code,'220900',amt,0))/2,0)) /?,0)   as fieldI_Y                                                                                                 ");
			sql.append(" 	  from cd01_99  cd01 left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = ?                                                                                         ");
			sql.append("                 left join bn01 on wlx01.bank_no=bn01.bank_no   and wlx01.m_year = bn01.m_year and  bn01.bank_type in ('6', '7')  and bn01.bn_type <> '2'   ");
			sql.append("                 left join ((select * from a01  where a01.m_year =  ?   and a01.m_month=  ? )                                                                                         ");
			sql.append(" 						   union all                                                                                                                               ");
			sql.append("                            (select m_year,m_month,bank_code,acc_code,amt from a02 where a02.m_year =   ?   and a02.m_month=  ?  and a02.acc_code in('990420','990620','990410', '990610'))                                ");
			sql.append("                            union all                                                                                                                                                  ");
			sql.append("                            (select * from a99 where a99.m_year =   ?   and a99.m_month=  ?  and a99.acc_code in('992510','992520','992530'))  ");//--97.08.26 add 逾期放款-正會員.贊助會員.非會員
			sql.append("                            union all                                                                                                                                                        ");
			sql.append("                            (select m_year, m_month, bank_code, acc_code, amt  from a05                                                                                                      ");
			sql.append(" 	  where a05.m_year =  ?   and a05.m_month=  ?  and  a05.acc_code = '91060P')) a01 on a01.bank_code=bn01.bank_no                                                                         ");
			sql.append(" 	  group by bn01.bank_type,nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order, bn01.bank_no ,bn01.BANK_NAME                                                     ");
			sql.append(" 	  )  a01                                                                                                                                                                                 ");
			sql.append(" where  bank_type  in  ('6', '7') ");
			sql.append(" order by  bank_type,FR001W_output_order,hsien_id,bank_no     ");
		}*/
		return sql.toString() ;
 	}
}



