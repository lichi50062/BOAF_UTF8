/*
 * 108.04.15 create 全體農漁會信用部各身分別存款金額一覽表  by 2295
*/
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.util.Region;

import java.io.*;
import java.text.DecimalFormat;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR081W {

  	public static String createRpt(String s_year,String s_month,String unit)
	{
    	
		String errMsg = "";
		String filename="";
        int rowNum = 4;
        String hsien_name = "";         
        String bank_no = "";            
        String bank_name = "";
        String field_mem = "";        
        String field_mem_rate = "";       
        String field_990420 = "";  
        String field_990420_rate = "";
        String field_990620 = "";
        String field_990620_rate = "";
        String field_debit = "";
        DecimalFormat formatter = new DecimalFormat("#.##");
        DataObject bean = null;
        
		filename="全體農漁會信用部各身分別存款金額一覽表.xls";
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
	        ps.setScale( ( short )76 ); //列印縮放百分比

	        ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
	  		
	  		finput.close();
	  		
	  		HSSFRow row=null;//宣告一列 
	  		HSSFCell cell=null;//宣告一個儲存格
	  		
	  	    HSSFCellStyle cs_center = reportUtil.getDefaultStyle(wb);//有框內文置中          
            HSSFCellStyle cs_left = reportUtil.getLeftStyle(wb);//有框內文置左
            HSSFCellStyle cs_right = reportUtil.getRightStyle(wb);//有框內文置右
          
	  	    List dbData = getData1(s_year,s_month,unit);		
	 	    System.out.println("dbData.size() ="+dbData.size());
	 	    
	 	    //列印年度
            row=sheet.getRow(1); 
            cell=row.getCell((short)0);         
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);         
            if(dbData.size() == 0){
                cell.setCellValue("                       中華民國　"+s_year+"　年 " + s_month + "　月無資料存在");                          
            }else{
                cell.setCellValue("                       中華民國　"+s_year+"　年 " + s_month + "　月");
            }
            
            //列印單位               
            HSSFFont ft = wb.createFont();
            HSSFCellStyle cs = wb.createCellStyle();
            ft.setFontHeightInPoints((short)12);
            ft.setFontName("標楷體");
            cs.setFont(ft);
            cs.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
            row=sheet.getRow(1); 
            cell=row.getCell((short)7);         
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellStyle(cs);
            cell.setCellValue("單位：新台幣"+Utility.getUnitName(unit)+"、％");
            
            rowNum = 4;
            
            if (dbData !=null && dbData.size() !=0) {
                for(int i=0;i<dbData.size();i++){
                    bean = (DataObject)dbData.get(i);                    
                    hsien_name = bean.getValue("hsien_name") == null?"":String.valueOf(bean.getValue("hsien_name"));//縣市別
                    bank_no = bean.getValue("bank_no") == null?"":String.valueOf(bean.getValue("bank_no"));//農(漁)會代號
                    bank_name = bean.getValue("bank_name") == null?"":String.valueOf(bean.getValue("bank_name"));//農(漁)會名稱                
                    field_mem = bean.getValue("field_mem") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_mem")));//會員-金額
                    field_mem_rate = bean.getValue("field_mem_rate") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_mem_rate")));//會員-比例
                    field_990420 = bean.getValue("field_990420") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_990420")));//贊助會員-金額
                    field_990420_rate = bean.getValue("field_990420_rate") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_990420_rate")));//贊助會員-比例
                    field_990620 = bean.getValue("field_990620") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_990620")));//非會員-金額
                    field_990620_rate = bean.getValue("field_990620_rate") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_990620_rate")));//非會員-比例
                    field_debit = bean.getValue("field_debit") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_debit")));//存款總額
                        
                    row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
                    for(int cellcount=0;cellcount<10;cellcount++){
                        cell = row.createCell( (short)cellcount);
                        cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                        if(cellcount == 0 || cellcount ==1 ){
                          cell.setCellStyle(cs_center);
                        }
                        if(cellcount == 2) cell.setCellStyle(cs_left);
                        if(cellcount > 2) cell.setCellStyle(cs_right);
                    }
                    //縣市別
                    cell=row.getCell((short)0);               
                    cell.setCellValue(hsien_name);
                    //農(漁)會代號
                    cell=row.getCell((short)1);                
                    cell.setCellValue(bank_no);
                    //農(漁)會名稱
                    cell=row.getCell((short)2);                
                    cell.setCellValue(bank_name);
                    //會員-金額
                    cell=row.getCell((short)3);                
                    cell.setCellValue(field_mem);
                    //會員-比例
                    cell=row.getCell((short)4);                
                    cell.setCellValue(field_mem_rate);
                    //贊助會員-金額
                    cell=row.getCell((short)5);                
                    cell.setCellValue(field_990420);
                    //贊助會員-比例
                    cell=row.getCell((short)6);                
                    cell.setCellValue(field_990420_rate);
                    //非會員-金額
                    cell=row.getCell((short)7);
                    cell.setCellValue(field_990620);
                    //非會員-比例
                    cell=row.getCell((short)8);            
                    cell.setCellValue(field_990620_rate);
                    //存款總額
                    cell=row.getCell((short)9);          
                    cell.setCellValue(field_debit);    
                    if("合計".equals(hsien_name)){
                      sheet.addMergedRegion( new Region( ( short )rowNum, ( short )0,             
                                                         ( short )rowNum, ( short )2) );       
                    }
                    
                    rowNum++;
                    
                }
            }//end of dbData is not null
		 
			
		    HSSFFooter footer=sheet.getFooter();
	        footer.setCenter( "Page:" +HSSFFooter.page() +" of " +HSSFFooter.numPages() );
		    footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
		    FileOutputStream fout=new FileOutputStream(reportDir+ System.getProperty("file.separator")+ filename);
		    wb.write(fout);
	        //儲存 
	        fout.close();
	        System.out.println("儲存完成");
		    }catch(Exception e){
		 		System.out.println("RptFR081W.createRpt Error:"+e+e.getMessage());
		 	}
		 	return errMsg;
	}

  	public static List getData1(String s_year,String s_month,String unit){
   	    //查詢年度100年以前.縣市別不同=================================================
        String cd01_table = (Integer.parseInt(s_year) < 100)?"cd01_99":"cd01"; 
        String wlx01_m_year = (Integer.parseInt(s_year) < 100)?"99":"100"; 
        //=====================================================================    
        StringBuffer sql = new StringBuffer();
        List paramList = new ArrayList();//共同參數
        
        sql.append("select bank_type, hsien_id, hsien_name, FR001W_output_order, bank_no, BANK_NAME,");        
        sql.append(" field_DEBIT - field_990420 - field_990620  as  field_mem,");//--會員-金額      
        sql.append(" decode(field_DEBIT,0,0,round((field_DEBIT - field_990420 - field_990620) / field_DEBIT *100 ,2)) as field_mem_rate,");//--會員-比例 
        sql.append(" field_990420, ");//--贊助會員-金額                     
        sql.append(" decode(field_DEBIT,0,0,round(field_990420 / field_DEBIT *100 ,2)) as field_990420_rate,");//--贊助會員-比例 
        sql.append(" field_990620, ");//--非會員-金額                 
        sql.append(" decode(field_DEBIT,0,0,round(field_990620 / field_DEBIT *100 ,2)) as field_990620_rate,");//--非會員-比例
        sql.append(" field_DEBIT ");//--總額                      
        sql.append(" from "); 
        sql.append(" (select bn01.bank_type,   nvl(cd01.hsien_id,' ')  hsien_id,");               
        sql.append("         nvl(cd01.hsien_name,'OTHER')  as  hsien_name,");                     
        sql.append("         cd01.FR001W_output_order,bn01.bank_no ,  bn01.BANK_NAME,");          
        sql.append("         round(sum(decode(a01.acc_code,'990420',amt,0)) /?,0) as  field_990420,");//  --贊助會員-存款總額                 
        sql.append("         round(sum(decode(a01.acc_code,'990620',amt,0)) /?,0) as  field_990620,");//  --非會員-存款總額                 
        sql.append("         round(sum(decode(a01.acc_code,'220000',amt,0)) /?,0) as field_DEBIT");// --存款總額
        paramList.add(unit);  
        paramList.add(unit);
        paramList.add(unit);
        sql.append("  from  "+cd01_table+" left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = ?");
        paramList.add(wlx01_m_year);        
        sql.append("  left join bn01 on wlx01.bank_no=bn01.bank_no and wlx01.m_year = bn01.m_year and  bn01.bank_type in ('6', '7')  and bn01.bn_type <> '2'");     
        sql.append("  left join ((select  (CASE WHEN (a01.m_year <= 102) THEN '102'");  
        sql.append("                            WHEN (a01.m_year > 102) THEN '103'");  
        sql.append("                            ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt");  
        sql.append("             from a01  where a01.m_year =? and a01.m_month= ? )");
        paramList.add(s_year); 
        paramList.add(s_month); 
        sql.append("             union all");          
        sql.append("            (select (CASE WHEN (m_year <= 102) THEN '102'");  
        sql.append("                          WHEN (m_year > 102) THEN '103'");  
        sql.append("                          ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt ");
        sql.append("             from a02 where a02.m_year = ? and a02.m_month=?");
        paramList.add(s_year); 
        paramList.add(s_month);
        sql.append("             and a02.acc_code in('990420','990620','990410', '990610'))");                     
        sql.append("             ) a01 on a01.bank_code=bn01.bank_no");            
        sql.append("   group by year_type,bn01.bank_type,nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order, bn01.bank_no ,bn01.BANK_NAME");               
        sql.append("   union"); 
        sql.append("   select 'ALL' as bank_type, 'ALL' as hsien_id,");               
        sql.append("          '合計'  as  hsien_name,");                     
        sql.append("          '999' as FR001W_output_order, '' as bank_no , '' as BANK_NAME,");          
        sql.append("          round(sum(decode(a01.acc_code,'990420',amt,0)) /?,0) as field_990420,");//--贊助會員-存款總額                 
        sql.append("          round(sum(decode(a01.acc_code,'990620',amt,0)) /?,0) as  field_990620,");//--非會員-存款總額                 
        sql.append("          round(sum(decode(a01.acc_code,'220000',amt,0)) /?,0) as field_DEBIT ");//--存款總額
        paramList.add(unit);  
        paramList.add(unit);
        paramList.add(unit);
        sql.append("   from  "+cd01_table+" left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = ?");
        paramList.add(wlx01_m_year); 
        sql.append("   left join bn01 on wlx01.bank_no=bn01.bank_no and wlx01.m_year = bn01.m_year and  bn01.bank_type in ('6', '7')  and bn01.bn_type <> '2'");     
        sql.append("   left join ((select  (CASE WHEN (a01.m_year <= 102) THEN '102'");  
        sql.append("                             WHEN (a01.m_year > 102) THEN '103'");  
        sql.append("                             ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt");  
        sql.append("               from a01  where a01.m_year = ? and a01.m_month= ? )");
        paramList.add(s_year); 
        paramList.add(s_month);
        sql.append("               union all");          
        sql.append("              (select (CASE WHEN (m_year <= 102) THEN '102'");  
        sql.append("                            WHEN (m_year > 102) THEN '103'");  
        sql.append("                            ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt");
        sql.append("               from a02 where a02.m_year = ? and a02.m_month= ?");
        paramList.add(s_year); 
        paramList.add(s_month);
        sql.append("               and a02.acc_code in('990420','990620','990410', '990610'))");                     
        sql.append("             ) a01 on a01.bank_code=bn01.bank_no");            
        sql.append(" )  a01");                      
        sql.append(" where  bank_type  in  ('6', '7','ALL')");                        
        sql.append(" order by  bank_type,FR001W_output_order,hsien_id,bank_no");
        
        
        List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"field_mem,field_mem_rate,field_990420,field_990420_rate,field_990620,field_990620_rate,field_debit");
        System.out.println("dbData1.size()="+dbData.size());
        return dbData;
    }
}



