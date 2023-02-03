/*
	94.09.27 add 明細表 by 2495
    95.06.26 fix data is null by 2295 
    99.04.12 fix 縣市合併問題與SQL 修改為preparedstatement 方式查詢 by 2808
   102.05.24 fix 漁會淨值都顯示為0,修改查詢SQL by 2295
*/
package com.tradevan.util.report;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.poifs.filesystem.*;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import java.io.*;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR029W {
  	public static String createRpt(String s_year,String s_month,String unit,String datestate,String bank_type){
  		System.out.println("start........RptFR029W.java") ;
  		List paramList = new ArrayList();
		DataObject bean = null;
		int u_year = 99 ; //配合縣市合併調整SQL用的參數
		String cd01Table = "cd01" ;
		s_year = s_year==null?"" : s_year.trim() ;
		if(!"".equals(s_year) && Integer.parseInt(s_year) >= 100) {
			u_year = 100  ;
		}else {
			u_year = 99 ;
			cd01Table = "cd01_99" ; 
		}
		System.out.println("u_year:"+u_year) ;
		String totalCount = "0" ; //家數
		String errMsg = "";
		String unit_name="";		
		int i=0;
		int j=0;
		String s_year_last="";
		String s_month_last="";
		String hsien_id_sum[]={"a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p"};		 
		String bank_type_name=(bank_type.equals("6"))?"農會":"漁會";
		
		
		String filename="";
		filename=(bank_type.equals("6"))?"全體農會信用部本期損益或淨值為負單位明細表.xls":"全體漁會信用部本期損益或淨值為負單位明細表.xls";		
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
    		String openfile="全體農會信用部本期損益或淨值為負單位明細表"+".xls";
    		System.out.println("open file "+filename);
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
	        ps.setScale( ( short )75 ); //列印縮放百分比

	        ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
	  		//wb.setSheetName(0,"test");
	  		finput.close();
               	  		
	  		HSSFRow row=null;//宣告一列 
	  		HSSFCell cell=null;//宣告一個儲存格
		
	        if("6".equals(bank_type)) {
	        	System.out.println("農會報表");
	        }else {
	        	System.out.println("漁會報表");
	        }
	        StringBuffer sql = new StringBuffer() ;
	        sql.append(" select hsien_id, hsien_name, 1 as cnt,FR001W_output_order, bank_no , BANK_NAME,");
	        sql.append("        round(fieldL/?,0)  as  fieldL,   ");
	        sql.append("        round(fieldM/?,0)  as  fieldM,   ");
	        sql.append("        round(fieldN/?,0)  as  fieldN,   ");
	        sql.append("        round(fieldNO/?,0)  as  fieldNO, ");
	        sql.append("        decode(fieldNO,0,0, round(( fieldN/fieldNO )*100 ,2) ) as  fieldO ");
	        sql.append(" from  ( ");
	        sql.append("        select nvl(cd01.hsien_id,' ')       	as  hsien_id ,  ");
	        sql.append("               nvl(cd01.hsien_name,'OTHER') 	as  hsien_name, ");
	        sql.append("               cd01.FR001W_output_order     	as  FR001W_output_order, ");
	        sql.append("               bn01.bank_no , bn01.BANK_NAME,   ");
	        sql.append("               sum(decode(a01.acc_code,'320300',amt,0))   as fieldL, ");
	        sql.append("               sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) as  fieldNO,");
	        sql.append("               sum(decode(a01.acc_code,'990000',amt,0))   as fieldN, ");
	        sql.append(" 		       decode(bank_type,'6',sum(decode(a01.acc_code,'310000',amt,'320000',amt,0)) ");
	        sql.append(" 			                    ,'7',sum(decode(a01.acc_code,'300000',amt,0)))  AS fieldM ");
	        sql.append("        from  (select * from ").append(cd01Table).append(" where hsien_id <> 'Y') cd01 ");
	        sql.append("              left join (select * from wlx01 where m_year=?) wlx01 on wlx01.hsien_id=cd01.hsien_id ");
	        sql.append("              left join (select * from bn01 where bank_type=? and m_year=? ) bn01 on wlx01.bank_no=bn01.bank_no ");
	        sql.append("              left join (select * from a01  ");
	        sql.append(" 			 	  	    where a01.m_year  =  ? and a01.m_month  = ?  ");
	        sql.append(" 			 	  	    ) a01 on  bn01.bank_no = a01.bank_code ");
	        sql.append("        GROUP BY  nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_type,");
	        sql.append(" 	   		 	 bn01.bank_no,bn01.BANK_NAME ");
	        sql.append("        )tmp ");
	        sql.append(" where (tmp.fieldL < 0)  or (tmp.fieldM < 0) ");
	        sql.append(" union  ");
	        //sql.append(" --農/漁會合計                                                                                 ");
	        sql.append(" select  'ZZ' as hsien_id ,'總計' as hsien_name ,count(*) as cnt,'999',' ' as bank_no,'' as BANK_NAME, "); 				  
	        sql.append(" 		round(sum(fieldL)/?,0)  as  fieldL, ");
	        sql.append(" 		round(sum(fieldM)/?,0)  as  fieldM, ");
	        sql.append(" 		round(sum(fieldN)/?,0)  as  fieldN, ");
	        sql.append(" 		round(sum(fieldNO)/?,0)  as  fieldNO,");
	        sql.append("  		decode(sum(fieldNO),0,0,round((sum(fieldN)/sum(fieldNO))*100,2)) as  fieldO ");				                     
	        sql.append(" from  (  ");
	        sql.append(" 		select nvl(cd01.hsien_id,' ')        as  hsien_id , ");
	        sql.append("   			   nvl(cd01.hsien_name,'OTHER')  as  hsien_name,");
	        sql.append("   			   cd01.FR001W_output_order      as  FR001W_output_order, ");
	        sql.append("   			   bn01.bank_no , bn01.BANK_NAME, ");
	        sql.append("               sum(decode(a01.acc_code,'320300',amt,0))  as fieldL,");
	        sql.append("               decode(bank_type,'6',sum(decode(a01.acc_code,'310000',amt,'320000',amt,0))");
	        sql.append(" 			                   ,'7',sum(decode(a01.acc_code,'300000',amt,0)))  AS fieldM, ");
	        sql.append("   			   sum(decode(a01.acc_code,'990000',amt,0))  as fieldN,                     ");
	        sql.append("               sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) as  fieldNO ");
	        sql.append("          from  (select * from ").append(cd01Table).append(" where hsien_id <> 'Y') cd01  ");
	        sql.append("          left join (select * from wlx01 where m_year=? )wlx01 on wlx01.hsien_id=cd01.hsien_id ");
	        sql.append(" 		  left join (select * from bn01 where bank_type=? and m_year=?)bn01 on wlx01.bank_no=bn01.bank_no");
	        sql.append("   		  left join (select * from a01 where a01.m_year  =  ? and a01.m_month  = ? ) a01 on  bn01.bank_no = a01.bank_code");
	        sql.append("    	  group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_type,");
	        sql.append("    		 		  bn01.bank_no ,  bn01.BANK_NAME  ");
	        sql.append("       ) totalsum   ");
	        sql.append(" where fieldL < 0  or fieldM < 0        ");
	        sql.append(" order by FR001W_output_order,  bank_no ");
	        
	        paramList.add(unit) ; 
	        paramList.add(unit) ; 
	        paramList.add(unit) ; 
	        paramList.add(unit) ; 
	        paramList.add(String.valueOf(u_year)) ;
	        paramList.add(bank_type) ;
	        paramList.add(String.valueOf(u_year)) ;
	        paramList.add(s_year) ;
	        paramList.add(s_month) ;
	        paramList.add(unit) ; 
	        paramList.add(unit) ; 
	        paramList.add(unit) ; 
	        paramList.add(unit) ; 
	        paramList.add(String.valueOf(u_year));
	        paramList.add(bank_type) ; 
	        paramList.add(String.valueOf(u_year)) ;
	        paramList.add(s_year) ;
	        paramList.add(s_month) ;
	        //抓全體農會信用部本期損益或淨值為負單位明細資料	
		
			List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"fieldL,fieldM,fieldN,fieldNO,fieldO,hsien_id,hsien_name,fr001w_output_order,cnt");
			int dbsize = dbData.size() ;
			System.out.println("dbsize:"+dbsize) ;
			//94.03.12 add 月份無資料 by 2495
			if(dbData == null || dbData.size() == 0)  {//95.06.26 add dbData == null by 2295
				row=sheet.getRow(1);
	  			cell=row.getCell((short)0);	       	
	  			//設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
	  			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
   	       	    cell.setCellValue(        s_year +"年" +s_month +"月無資料存在");
	  	 	}else{
	  	 		String div=(Integer.parseInt(s_year)==94 && Integer.parseInt(s_month)==6)?"1":"2";
	  			if (Integer.parseInt(s_month) == 1) {
					s_year_last = String.valueOf(Integer.parseInt(s_year) - 1);
					s_month_last = "12";
				} else {
					s_year_last = s_year;
					s_month_last = String.valueOf(Integer.parseInt(s_month) - 1);
				}
	  	   
		  	    unit_name = Utility.getUnitName(unit);//取得單位名稱
		  	    //列印年度
				row=(sheet.getRow(1)==null)? sheet.createRow(1) : sheet.getRow(1);													
				cell=row.getCell((short)0);			
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				if(s_month.equals("0")){
					cell.setCellValue("  中華民國　"+s_year+"　年度");
				}else {				
					cell.setCellValue("                      			 中華民國　"+s_year+"　年 " + s_month + "月                 單位：新台幣"+unit_name+"、％");							
				}
			
				System.out.println("明細表----------------------------");
			  
				int rowNum=3;
				String  insertValue="";
				for(int rowcount=0;rowcount<dbData.size();rowcount++){
					for (int cellcount = 0; cellcount < 7; cellcount++) {
						bean = (DataObject) dbData.get(rowcount);
						String hsien_id = bean.getValue("hsien_id")==null ? "" : bean.getValue("hsien_id").toString() ;
						if("ZZ".equals(hsien_id)) {
							totalCount = bean.getValue("cnt")==null? "0" : bean.getValue("cnt").toString() ;
						}
						if (cellcount == 0)
							insertValue = bean.getValue("hsien_name").toString();
						if (cellcount == 1)
							insertValue = bean.getValue("bank_no") == null ? "": bean.getValue("bank_no").toString();
						if (cellcount == 2)
							insertValue = bean.getValue("bank_name") == null ? "": bean.getValue("bank_name").toString();
						if (cellcount == 3)
							insertValue = bean.getValue("fieldl") == null ? "0": bean.getValue("fieldl").toString();
						if (cellcount == 4)
							insertValue = bean.getValue("fieldm") == null ? "0": bean.getValue("fieldm").toString();
						if (cellcount == 5)
							insertValue = bean.getValue("fieldn") == null ? "0": bean.getValue("fieldn").toString();
						if (cellcount == 6)
							insertValue = bean.getValue("fieldo") == null ? "0": bean.getValue("fieldo").toString();

						// 儲存CELL值
						row = sheet.createRow(rowNum);
						cell = row.createCell((short) cellcount);
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						if (cellcount == 0 || cellcount == 1 || cellcount == 2)
							cell.setCellValue(insertValue);
						else
							cell.setCellValue(Utility
									.setCommaFormat(insertValue));
						// 設定CELL格式
						HSSFCellStyle cs1 = wb.createCellStyle();
						// HSSFCellStyle cs1 =
						// cell.getCellStyle();//會套用原本excel所設定的格式
						cs1.setBorderTop((short) 0);
						int countline = rowcount;
						countline++;
						if (countline == dbData.size())
							cs1.setBorderBottom(HSSFCellStyle.BORDER_THIN);
						else
							cs1.setBorderBottom(HSSFCellStyle.BORDER_THIN);

						cs1.setBorderLeft(HSSFCellStyle.BORDER_THIN);
						cs1.setBorderRight(HSSFCellStyle.BORDER_THIN);
						HSSFFont f = wb.createFont();
						f.setFontHeightInPoints((short) 10);
						f.setFontName("標楷體");
						cs1.setFont(f);

						if (cellcount == 0 || cellcount == 2)
							cs1.setAlignment(HSSFCellStyle.ALIGN_GENERAL);
						else if (cellcount == 1)
							cs1.setAlignment(HSSFCellStyle.ALIGN_CENTER);
						else
							cs1.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
						cell.setCellStyle(cs1);
					}			
				rowNum++;                             
			}
		System.out.println("總計----------------------------");
		row = sheet.createRow(rowNum);
		cell=row.createCell((short)0);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellValue("  總       計  ");
		
		sheet.addMergedRegion(new Region(rowNum,(short)0,rowNum,(short)1));
		//設定CELL格式
		HSSFCellStyle cs1 = wb.createCellStyle();		
		cs1.setBorderTop(HSSFCellStyle.BORDER_THIN);               
		cs1.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		cs1.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		cs1.setBorderRight(HSSFCellStyle.BORDER_THIN);                			
		HSSFFont f = wb.createFont();
		f.setFontHeightInPoints((short)10);
		f.setFontName("標楷體");		    
		cs1.setFont(f);               
		cs1.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		cell.setCellStyle(cs1);
		
        //補總計之其他格式
		for(int cellcount=1;cellcount<3;cellcount++){
			cell=row.createCell((short)cellcount);
 			cs1.setBorderTop(HSSFCellStyle.BORDER_THIN);               
			cs1.setBorderBottom(HSSFCellStyle.BORDER_THIN);		
			cs1.setBorderRight(HSSFCellStyle.BORDER_THIN);                			
			cell.setCellStyle(cs1);
		}
		
			//95.06.26 add cnt != 0 才顯示資料 by 2295
		    
			//bean = (DataObject) dbData.get(dbsize-2);
			for(int cellcount=3;cellcount<7;cellcount++){									
				if(cellcount==3)										
					insertValue = bean.getValue("fieldl")==null? "0" : bean.getValue("fieldl").toString();
				if(cellcount==4)
					insertValue = bean.getValue("fieldm")==null? "0" : bean.getValue("fieldm").toString();
                if(cellcount==5)
					insertValue = bean.getValue("fieldn")==null? "0" : bean.getValue("fieldn").toString();
				if(cellcount==6)
					insertValue = bean.getValue("fieldo")==null? "0" : bean.getValue("fieldo").toString();
				
				System.out.println("insertValue = "+insertValue );											
				                                                    						
                //儲存CELL值
                row = sheet.createRow(rowNum);
				cell=row.createCell((short)cellcount);	
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellValue(Utility.setCommaFormat(insertValue));
						
                //設定CELL格式           		
				cs1 = wb.createCellStyle();		
				cs1.setBorderTop(HSSFCellStyle.BORDER_THIN);               
				cs1.setBorderBottom(HSSFCellStyle.BORDER_THIN);
				cs1.setBorderLeft(HSSFCellStyle.BORDER_THIN);
				cs1.setBorderRight(HSSFCellStyle.BORDER_THIN);                	
				f = wb.createFont();
				f.setFontHeightInPoints((short)10);
				f.setFontName("標楷體");		    
				cs1.setFont(f);               
				cs1.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
				cell.setCellStyle(cs1);                     	
            }
			rowNum++;
			System.out.println("家數----------------------------");
			row = sheet.createRow(rowNum);
			cell=row.createCell((short)0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("  家       數  ");
			sheet.addMergedRegion(new Region(rowNum,(short)0,rowNum,(short)1));
			//設定CELL格式
			cs1 = wb.createCellStyle();		
			cs1.setBorderTop(HSSFCellStyle.BORDER_THIN);               
			cs1.setBorderBottom(HSSFCellStyle.BORDER_THIN);
			cs1.setBorderLeft(HSSFCellStyle.BORDER_THIN);
			cs1.setBorderRight(HSSFCellStyle.BORDER_THIN);                			
			f = wb.createFont();
			f.setFontHeightInPoints((short)10);
			f.setFontName("標楷體");		    
			cs1.setFont(f);               
			cs1.setAlignment(HSSFCellStyle.ALIGN_CENTER);
			cell.setCellStyle(cs1);
		
			cell=row.createCell((short)1);
	 		cs1.setBorderTop(HSSFCellStyle.BORDER_THIN);               
			cs1.setBorderBottom(HSSFCellStyle.BORDER_THIN);		
			cs1.setBorderRight(HSSFCellStyle.BORDER_THIN);                			
			cell.setCellStyle(cs1);
			//儲存CELL值
			//insertValue = bean.getValue("cnt")==null? "0" : (String)bean.getValue("cnt");
            row = sheet.createRow(rowNum);
			cell=row.createCell((short)2);	
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue(Utility.setCommaFormat(totalCount));
            //設定CELL格式		           		
			cs1 = wb.createCellStyle();		
			cs1.setBorderTop(HSSFCellStyle.BORDER_THIN);               
			cs1.setBorderBottom(HSSFCellStyle.BORDER_THIN);
			cs1.setBorderLeft(HSSFCellStyle.BORDER_THIN);
			cs1.setBorderRight(HSSFCellStyle.BORDER_THIN); 
			f = wb.createFont();
			f.setFontHeightInPoints((short)10);
			f.setFontName("標楷體");		    
			cs1.setFont(f);               
			cs1.setAlignment(HSSFCellStyle.ALIGN_CENTER);
			cell.setCellStyle(cs1);
			//補家數之其他格式
			for(int cellcount=3;cellcount<7;cellcount++){
				cell=row.createCell((short)cellcount);
	 			cs1.setBorderTop(HSSFCellStyle.BORDER_THIN);               
				cs1.setBorderBottom(HSSFCellStyle.BORDER_THIN);		
				cs1.setBorderRight(HSSFCellStyle.BORDER_THIN);                			
				cell.setCellStyle(cs1);
			}
			
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
}



