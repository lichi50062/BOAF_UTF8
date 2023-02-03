/*
	2006/3/10  add 明細表 by 2495
    2010.04.27 fix 縣市合併 && sql injection by 2808
    2010.11.09 fix 縣市合併排序 by 2808 
    		       add v_bank_location 關聯排序
*/
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import java.io.*;
import java.util.*;
import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR007WA {
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
		StringBuffer sql = new StringBuffer() ;
		ArrayList paramList = new ArrayList() ;
		DataObject bean = null ;
		String u_year = "100" ;
		String cd01 = "cd01" ; //Table Name 
		String orderbyFile = "FR001W_OUTPUT_ORDER" ;
		if(s_year!=null && Integer.parseInt(s_year) <=99 ) {
        	u_year = "99" ;
        	cd01 = "cd01_99" ;
        	orderbyFile = "OUTPUT_ORDER" ;
        }
		String filename="";
		filename=(bank_type.equals("6"))?"農會資產品質分析明細表.xls":"漁會資產品質分析明細表.xls";		
		
		
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
	  		if(fs==null){System.out.println("open 範本檔失敗");}
	  		HSSFWorkbook wb = new HSSFWorkbook(fs);
	  		if(wb==null){System.out.println("open工作表失敗");}
	  		HSSFSheet sheet = wb.getSheetAt(0);//讀取第一個工作表，宣告其為sheet 
	  		if(sheet==null){System.out.println("open sheet 失敗");}
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
			unit_name = Utility.getUnitName(unit) ;
			
			cell.setCellValue(" 單位：新台幣"+unit_name+"、％");						
			

		

		//抓農漁會信用部各類對象存放款比率資料	
		sql.append(" select t1.* from ( ");
        sql.append(" select    a04.hsien_id");
        sql.append(" ,a04.hsien_name");
        sql.append(" ,a04.FR001W_output_order");
        sql.append(" ,a04.bank_no ");
        sql.append(" ,a04.BANK_NAME");
        sql.append(" ,(CASE WHEN (field_OVER_A01 = field_OVER_A04 ) THEN '   ' ELSE '***'  END)  as OVER_TYPE_1 ") ;
        sql.append(" ,(CASE WHEN (field_CREDIT_A01 = field_CREDIT_A04 ) THEN '   ' ");
        sql.append("   ELSE '***'  END)  as CREDIT_TYPE_1 ") ;
        sql.append(" ,round(field_OVER_A01 /? ,0)    as field_OVER_A01 ") ;
        paramList.add(unit) ;
        sql.append(" ,round(field_CREDIT_A01 /? ,0)  as field_CREDIT_A01 ") ;
        paramList.add(unit) ;
        sql.append(" ,decode(field_CREDIT_A01,0,0,round(field_OVER_A01 /  field_CREDIT_A01 *100 ,2))  as   field_OVER_RATE_A01 ");
        sql.append(" ,round(field_OVER_A04 / ? ,0)    as field_OVER_A04 ") ;
        paramList.add(unit) ;
        sql.append(" ,round(field_CREDIT_A04 / ? ,0)  as field_CREDIT_A04 ") ;
        paramList.add(unit) ;
        sql.append(" ,decode(field_CREDIT_A04,0,0,round(field_OVER_A04 /  field_CREDIT_A04 *100 ,2))  as   field_OVER_RATE_A04 ");
        sql.append(" ,round(field_840760_A04 /? ,0)  as field_840760_A04  ") ;
        paramList.add(unit) ;
        sql.append(" ,decode(field_CREDIT_A04,0,0,round(field_840760_A04 /  field_CREDIT_A04 *100 ,2))  as   field_RATE_840760_A04 ") ;
        sql.append(" ,decode(field_CREDIT_A04,0,0,round((field_OVER_A04 + field_840760_A04) /  field_CREDIT_A04 *100 ,2))  as   field_RATE_840760_OVER_A04");
        sql.append(" ,round(field_840710_A04_A /?,0)  as field_840710_A04_A ");
        sql.append(" ,round(field_840720_A04_B /?,0)  as field_840720_A04_B ");
        sql.append(" ,round(field_840731_A04_a /?,0)  as field_840731_A04_a ");
        sql.append(" ,round(field_840732_A04_b /?,0)  as field_840732_A04_b ");
        sql.append(" ,round(field_840733_A04_c /?,0)  as field_840733_A04_c ");
        sql.append(" ,round(field_840734_A04_d /?,0)  as field_840734_A04_d ");
        sql.append(" ,round(field_840735_A04_e /?,0)  as field_840735_A04_e  ") ;
        paramList.add(unit) ;
        paramList.add(unit) ;
        paramList.add(unit) ;
        paramList.add(unit) ;
        paramList.add(unit) ;
        paramList.add(unit) ;
        paramList.add(unit) ;
        
		sql.append(" from ( ") ;
		sql.append(" select nvl(cd01.hsien_id,' ') as  hsien_id , ");
		sql.append(" nvl(cd01.hsien_name,'OTHER')  as  hsien_name, ");
		sql.append(" cd01.").append(orderbyFile).append("      as  FR001W_output_order, ");
		sql.append(" bn01.bank_no , bn01.BANK_NAME, sum(decode(acc_code,'990000',amt,0))  as field_OVER_A01, ");
		sql.append(" sum(decode(acc_code,'120000',amt,'120800',amt,'150300',amt,0))  	as  field_CREDIT_A01, ");
		sql.append(" sum(decode(acc_code,'840740',amt,0))         as field_OVER_A04, ");
		sql.append(" sum(decode(acc_code,'840750',amt,0))         as field_CREDIT_A04, ");
		sql.append(" sum(decode(acc_code,'840760',amt,0))         as field_840760_A04, ");
		sql.append(" sum(decode(acc_code,'840710',amt,0))         as field_840710_A04_A, ");
		sql.append(" sum(decode(acc_code,'840720',amt,0))         as field_840720_A04_B, ");
		sql.append(" sum(decode(acc_code,'840731',amt,0))         as field_840731_A04_a, ");
		sql.append(" sum(decode(acc_code,'840732',amt,0))         as field_840732_A04_b, ");
		sql.append(" sum(decode(acc_code,'840733',amt,0))         as field_840733_A04_c, ") ;
		sql.append(" sum(decode(acc_code,'840734',amt,0))         as field_840734_A04_d, ") ;
		sql.append(" sum(decode(acc_code,'840735',amt,0))         as field_840735_A04_e") ;
        
		sql.append(" from  (select * from ").append(cd01).append(" where hsien_id <> 'Y') cd01 ");
		sql.append(" left join (select * from wlx01 where wlx01.m_year= ? ) wlx01  ");
		paramList.add(u_year) ;
		sql.append(" on wlx01.hsien_id=cd01.hsien_id ") ;
		sql.append(" left join (select * from bn01 where bn01.m_year= ? and bn01.bank_type =? ) bn01 ");
		paramList.add(u_year) ;
		paramList.add(bank_type) ;
		sql.append(" on wlx01.bank_no=bn01.bank_no ") ;
        sql.append(" left join ((select * from a04 where m_year  =  ?  and m_month= ? ) ");
        paramList.add(s_year) ;
        paramList.add(s_month) ;
        sql.append(" union ") ;
        sql.append(" (select * from a01 ") ;
        sql.append("  where    m_year  = ?  and m_month = ? and ") ;
        paramList.add(s_year) ;
        paramList.add(s_month) ;
        sql.append(" acc_code in('990000', '120000','120800','150300'))) a04 ") ;
        sql.append(" on  bn01.bank_no = a04.bank_code ") ;
        sql.append(" group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.").append(orderbyFile).append(", ") ;
        sql.append(" bn01.bank_no ,  bn01.BANK_NAME) a04   ");
        sql.append(" where a04.bank_no <> ' ' ") ;
        sql.append(" order by a04.hsien_id ,  a04.hsien_name,  a04.FR001W_output_order, a04.bank_no, a04.BANK_NAME ") ;
        sql.append(" ) T1 , v_bank_location T2 ");
        sql.append(" where t2.m_year = ? and t1.bank_no = t2.bank_no");
        paramList.add(s_year) ;
        sql.append(" order by t2.fr001w_output_order ") ;
	    List dbData = DBManager.QueryDB_SQLParam(sql.toString(), paramList, "field_over_a01,field_credit_a01,field_over_rate_a01,field_over_a04,field_credit_a04,field_over_rate_a04,field_840760_a04,field_rate_840760_a04,field_rate_840760_over_a04,field_840710_a04_a,field_840720_a04_b,field_840731_a04_a,field_840732_a04_b,field_840733_a04_c,field_840734_a04_d,field_840735_a04_e" ) ;
		
        sql.setLength(0) ;
        paramList.clear() ;
        System.out.println("dbData.size() ="+dbData.size());
	    System.out.println("明細表----------------------------");

	    int count_over_type_1=0, count_credit_type_1=0;
		int rowNum=4;
		String  insertValue="";
		for(int rowcount=0;rowcount<dbData.size();rowcount++)
		{		
		    //System.out.println("rowcount="+rowcount);			       
		    for(int cellcount=0;cellcount<20;cellcount++)
				{	
					//System.out.println("cellcount="+cellcount);			
					if(cellcount==0)
					{		
						bean = (DataObject) dbData.get(rowcount) ;
						insertValue = bean.getValue("over_type_1").toString();											
					  	//System.out.println("over_type_1="+insertValue);
						//2006_03_27 add 計算逾放不符家數 by 2495 
						if(insertValue.equals("***"))
						{
							count_over_type_1++; 
						}	
					}
					if(cellcount==1)	
					{			
						 insertValue = bean.getValue("credit_type_1").toString();
						 //System.out.println("credit_type_1="+insertValue);
						//2006_03_27 add 計算逾放不符家數 by 2495 
						if(insertValue.equals("***"))
						{
							count_credit_type_1++; 
						}	
					}
					if(cellcount==2)
					{				
						insertValue = bean.getValue("bank_no").toString();
						//System.out.println("bank_no="+insertValue);
					}
					if(cellcount==3)	
					{			
						insertValue = bean.getValue("bank_name").toString();
						//System.out.println("bank_name="+insertValue);
					}
					if(cellcount==4)	
					{			
						insertValue = bean.getValue("field_over_a01").toString();						
						//System.out.println("field_OVER_A01="+insertValue);
					}
					if(cellcount==5)
					{				
						insertValue = bean.getValue("field_credit_a01").toString();
						//System.out.println("field_CREDIT_A01="+insertValue);
					}
					if(cellcount==6)
					{				
						insertValue = bean.getValue("field_over_rate_a01").toString();
						//System.out.println("field_OVER_RATE_A01="+insertValue);
					}
					if(cellcount==7)
					{				
						insertValue = bean.getValue("field_over_a04").toString();
						//System.out.println("field_OVER_A04="+insertValue);
					}
					if(cellcount==8)
					{				
						insertValue = bean.getValue("field_credit_a04").toString();
					  //System.out.println("field_CREDIT_A04="+insertValue);
					}
					if(cellcount==9)
					{				
						insertValue = bean.getValue("field_over_rate_a04").toString();
					  //System.out.println("field_OVER_RATE_A04="+insertValue);
					}
					if(cellcount==10)
					{				
					  insertValue = bean.getValue("field_840760_a04").toString();					  
                      //System.out.println("field_840760_A04="+insertValue);
					}
					if(cellcount==11)
					{				
						insertValue = bean.getValue("field_rate_840760_a04").toString();
						//System.out.println("field_RATE_840760_A04="+insertValue);
					}
					if(cellcount==12)
					{		
						insertValue = bean.getValue("field_rate_840760_over_a04").toString();
						//System.out.println("field_RATE_840760_OVER_A04="+insertValue);
					}
					if(cellcount==13)	
					{	
						insertValue = bean.getValue("field_840710_a04_a").toString();						
					  //System.out.println("field_840710_A04_A="+insertValue);
					}
					if(cellcount==14)		
					{
						insertValue = bean.getValue("field_840720_a04_b").toString();
						//System.out.println("field_840720_A04_B="+insertValue);
					}
					if(cellcount==15)		
					{
						insertValue = bean.getValue("field_840731_a04_a").toString();
					  	
						//System.out.println("field_840731_A04_a="+insertValue);
					}
					if(cellcount==16)
					{				
						insertValue = bean.getValue("field_840732_a04_b").toString();
						//insertValue = (String)((DataObject)dbData.get(0)).getValue("field_840732_a04_b");
						//insertValue="1";
						//System.out.println("field_840732_A04_b="+insertValue);
					}
					if(cellcount==17)	
					{			
						insertValue = bean.getValue("field_840733_a04_c").toString();
						//System.out.println("field_840733_A04_c="+insertValue);
					}
					if(cellcount==18)	
					{			
						insertValue = bean.getValue("field_840734_a04_d").toString();
						//System.out.println("field_840734_A04_d="+insertValue);
					}
					if(cellcount==19)
					{				
						insertValue = bean.getValue("field_840735_a04_e").toString();
						//System.out.println("field_840735_A04_e="+insertValue);
					}
					insertCell(insertValue,rowNum,cellcount,wb,row,sheet,cell);				
         }
				rowNum++;                             
			}
	



        //合計
		System.out.println(" start 合計 =================");
        sql.append(" select    a04.hsien_id");
        sql.append(" ,a04.hsien_name");
        sql.append(" ,a04.FR001W_output_order");
        sql.append(" ,a04.bank_no ");
        sql.append(" ,a04.BANK_NAME");
        sql.append(" ,(CASE WHEN (field_OVER_A01 = field_OVER_A04 ) THEN '   ' ELSE '***'  END)  as OVER_TYPE_1 ") ;
        sql.append(" ,(CASE WHEN (field_CREDIT_A01 = field_CREDIT_A04 ) THEN '   ' ");                               
	    sql.append("        ELSE '***'  END)  as CREDIT_TYPE_1 ");
	    sql.append(" ,round(field_OVER_A01 /?,0)  as field_OVER_A01");
	    sql.append(" ,round(field_CREDIT_A01 /?,0)  as field_CREDIT_A01");
		sql.append(" ,decode(field_CREDIT_A01,0,0,round(field_OVER_A01 /  field_CREDIT_A01 *100 ,2))  as field_OVER_RATE_A01");
 	    sql.append(" ,round(field_OVER_A04 /?,0)    as field_OVER_A04");
	    sql.append(" ,round(field_CREDIT_A04 /?,0)  as field_CREDIT_A04");
        sql.append(" ,decode(field_CREDIT_A04,0,0,round(field_OVER_A04 /  field_CREDIT_A04 *100 ,2))  as   field_OVER_RATE_A04");
	    sql.append(" ,round(field_840760_A04 /?,0)  as field_840760_A04");
	    sql.append(" ,decode(field_CREDIT_A04,0,0,round(field_840760_A04 /  field_CREDIT_A04 *100 ,2))  as   field_RATE_840760_A04");
	    sql.append(" ,decode(field_CREDIT_A04,0,0,round((field_OVER_A04 + field_840760_A04) /  field_CREDIT_A04 *100 ,2))  as   field_RATE_840760_OVER_A04");
	    sql.append(" ,round(field_840710_A04_A /?,0)  as field_840710_A04_A");
        sql.append(" ,round(field_840720_A04_B /?,0)  as field_840720_A04_B");
        sql.append(" ,round(field_840731_A04_a /?,0)  as field_840731_A04_a");
	    sql.append(" ,round(field_840732_A04_b /?,0)  as field_840732_A04_b");
	    sql.append(" ,round(field_840733_A04_c /?,0)  as field_840733_A04_c");
	    sql.append(" ,round(field_840734_A04_d /?,0)  as field_840734_A04_d");
	    sql.append(" ,round(field_840735_A04_e /?,0)  as field_840735_A04_e ");
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
	    paramList.add(unit) ;
	    paramList.add(unit) ;
	    sql.append(" from ("                                                                                                                               );
        sql.append(" select ' '       as  hsien_id ,"                                                                                                      );
	    sql.append("        ' '       as  hsien_name,"                                                                                                     );
 		sql.append("        ' '       as  FR001W_output_order,"                                                                                            );
 		sql.append("        ' '   bank_no ,   '合計'    as  BANK_NAME,"                                                                                    );
 		sql.append(" sum(decode(acc_code,'990000',amt,0))         as field_OVER_A01,"                                                                      );
        sql.append(" sum(decode(acc_code,'120000',amt,'120800',amt,'150300',amt,0))  as  field_CREDIT_A01,"                                                );
 	    sql.append(" sum(decode(acc_code,'840740',amt,0))         as field_OVER_A04,"                                                                      );
 	    sql.append(" sum(decode(acc_code,'840750',amt,0))         as field_CREDIT_A04,"                                                                    );
        sql.append(" sum(decode(acc_code,'840760',amt,0))         as field_840760_A04,"                                                                    );
 	    sql.append(" sum(decode(acc_code,'840710',amt,0))         as field_840710_A04_A,"                                                                  );
 	    sql.append(" sum(decode(acc_code,'840720',amt,0))         as field_840720_A04_B,"                                                                  );
 	    sql.append(" sum(decode(acc_code,'840731',amt,0))         as field_840731_A04_a,"                                                                  );
 		sql.append(" sum(decode(acc_code,'840732',amt,0))         as field_840732_A04_b,"                                                                  );
 	    sql.append(" sum(decode(acc_code,'840733',amt,0))         as field_840733_A04_c,"                                                                  );
 	    sql.append(" sum(decode(acc_code,'840734',amt,0))         as field_840734_A04_d,"                                                                  );
 	    sql.append(" sum(decode(acc_code,'840735',amt,0))         as field_840735_A04_e "                                                                  );
	    sql.append(" from  (select * from ").append(cd01).append(" where hsien_id <> 'Y') cd01"                                                                           );
 		sql.append(" left join (select * from wlx01 where wlx01.m_year = ?) wlx01  ");
 		paramList.add(u_year) ;
 		sql.append(" on wlx01.hsien_id=cd01.hsien_id");
		sql.append(" left join (select * from bn01 where bn01.m_year=? and bn01.bank_type = ? ) bn01  ");
		paramList.add(u_year) ;
		paramList.add(bank_type) ;
		sql.append(" on wlx01.bank_no=bn01.bank_no  ");
		sql.append(" left join ((select * from a04 where m_year  =  ?  and m_month  =  ? ) ") ;	
       	paramList.add(s_year) ;
       	paramList.add(s_month) ;
       	//	total_sqlCmd +=" left join ((select * from a04 where m_year  =  "+s_year+" and m_month  = "+s_month+")" 
	    //          +" union"
       	sql.append(" union ");
       	sql.append(" (select * from a01 ");
       	sql.append("  where m_year  =  ?  and m_month  = ?   and ");
       	paramList.add(s_year) ;
       	paramList.add(s_month) ;
       	sql.append(" acc_code in('990000', '120000','120800','150300'))) a04 ");
       	sql.append(" on  bn01.bank_no = a04.bank_code) a04 ");
		
       	dbData = DBManager.QueryDB_SQLParam(sql.toString(), paramList, "field_over_a01,field_credit_a01,field_over_rate_a01,field_over_a04,field_credit_a04,field_over_rate_a04,field_840760_a04,field_rate_840760_a04,field_rate_840760_over_a04,field_840710_a04_a,field_840720_a04_b,field_840731_a04_a,field_840732_a04_b,field_840733_a04_c,field_840734_a04_d,field_840735_a04_e" ) ;
       			

       	rowNum++;

       	for( int rowcount=0;rowcount<dbData.size();rowcount++)
			{		
		    //System.out.println("rowcount="+rowcount);			       
		    for( int cellcount=0;cellcount<20;cellcount++)
				{	
					//System.out.println("cellcount="+cellcount);			
					if(cellcount==0)
					{					
						bean = (DataObject)dbData.get(rowcount) ;
						insertValue = bean.getValue("over_type_1").toString();											
					  	//2006_03_27 add 計算逾放不符家數 by 2495 
						insertValue = Integer.toString(count_over_type_1);
						//System.out.println("count_over_type_1="+insertValue);	
					}
					if(cellcount==1)	
					{			
						 insertValue = bean.getValue("credit_type_1").toString();
						 //2006_03_27 add 計算逾放不符家數 by 2495 
						 insertValue = Integer.toString(count_credit_type_1);
						 //System.out.println("count_credit_type_1="+insertValue);	
					}
					if(cellcount==2)
					{				
						insertValue = bean.getValue("bank_no").toString();
						//System.out.println("bank_no="+insertValue);
					}
					if(cellcount==3)	
					{			
						insertValue = bean.getValue("bank_name").toString();
						//System.out.println("bank_name="+insertValue);
					}
					if(cellcount==4)	
					{			
						insertValue = bean.getValue("field_over_a01").toString();						
						//System.out.println("field_OVER_A01="+insertValue);
					}
					if(cellcount==5)
					{				
						insertValue = bean.getValue("field_credit_a01").toString();
						//System.out.println("field_CREDIT_A01="+insertValue);
					}
					if(cellcount==6)
					{				
						insertValue = bean.getValue("field_over_rate_a01").toString();
						//System.out.println("field_OVER_RATE_A01="+insertValue);
					}
					if(cellcount==7)
					{				
						insertValue = bean.getValue("field_over_a04").toString();
						//System.out.println("field_OVER_A04="+insertValue);
					}
					if(cellcount==8)
					{				
						insertValue = bean.getValue("field_credit_a04").toString();
					  //System.out.println("field_CREDIT_A04="+insertValue);
					}
					if(cellcount==9)
					{				
						insertValue = bean.getValue("field_over_rate_a04").toString();
					  //System.out.println("field_OVER_RATE_A04="+insertValue);
					}
					if(cellcount==10)
					{				
					  insertValue = bean.getValue("field_840760_a04").toString();
                                          
					  //System.out.println("field_840760_A04="+insertValue);
					}
					if(cellcount==11)
					{				
						insertValue = bean.getValue("field_rate_840760_a04").toString();
						//System.out.println("field_RATE_840760_A04="+insertValue);
					}
					if(cellcount==12)
					{		
						insertValue = bean.getValue("field_rate_840760_over_a04").toString();
						//System.out.println("field_RATE_840760_OVER_A04="+insertValue);
					}
					if(cellcount==13)	
					{	
						insertValue = bean.getValue("field_840710_a04_a").toString();						
					  //System.out.println("field_840710_A04_A="+insertValue);
					}
					if(cellcount==14)		
					{
						insertValue = bean.getValue("field_840720_a04_b").toString();
						//System.out.println("field_840720_A04_B="+insertValue);
					}
					if(cellcount==15)		
					{
						insertValue = bean.getValue("field_840731_a04_a").toString();
					  	
						System.out.println("field_840731_A04_a="+insertValue);
					}
					if(cellcount==16)
					{				
						insertValue = bean.getValue("field_840732_a04_b").toString();
						//insertValue = (String)((DataObject)dbData.get(0)).getValue("field_840732_a04_b");
						//insertValue="1";
						//System.out.println("field_840732_A04_b="+insertValue);
					}
					if(cellcount==17)	
					{			
						insertValue = bean.getValue("field_840733_a04_c").toString();
						//System.out.println("field_840733_A04_c="+insertValue);
					}
					if(cellcount==18)	
					{			
						insertValue = bean.getValue("field_840734_a04_d").toString();
						//System.out.println("field_840734_A04_d="+insertValue);
					}
					if(cellcount==19)
					{				
						insertValue = bean.getValue("field_840735_a04_e").toString();
						//System.out.println("field_840735_A04_e="+insertValue);
					}
					insertCell(insertValue,rowNum,cellcount,wb,row,sheet,cell);				
         }
rowNum++;
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



