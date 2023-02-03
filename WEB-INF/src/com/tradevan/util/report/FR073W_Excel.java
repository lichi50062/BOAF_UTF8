//104.11.11 add by 2968  
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*;
import java.io.*;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class FR073W_Excel {
    public static String createRpt(String S_YEAR,String S_MONTH,String bank_code,String unit,String bank_type,String muser_id){
    	
    	String errMsg = "";
		StringBuffer sql = new StringBuffer () ;
		ArrayList paramList = new ArrayList() ;
		String unit_name=Utility.getUnitName(unit);
		FileInputStream finput = null;
		List dbData = new ArrayList();
    	String bank_name="";
    	String u_year = "100" ;
		if(S_YEAR!=null && Integer.parseInt(S_YEAR) <= 99) {
			u_year = "99" ;
		}
		
		Utility.printLogTime("FR073W_Excel begin time");	
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
    		
            finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+"信用部建築貸款統計表.xls" );
			
			System.out.println(xlsDir + System.getProperty("file.separator")+"信用部建築貸款統計表.xls");
	  	    //設定FileINputStream讀取Excel檔
	  		POIFSFileSystem fs = new POIFSFileSystem( finput );
	  		HSSFWorkbook wb = new HSSFWorkbook(fs);
	  		HSSFSheet sheet = wb.getSheetAt(0);//讀取第一個工作表，宣告其為sheet 
	  		HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
	        //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
	        //sheet.setAutobreaks(true); //自動分頁
			
	        //設定頁面符合列印大小
	        //sheet.setAutobreaks( false );
	        ps.setScale( ( short )85 ); //列印縮放百分比
	        ps.setLandscape(true);//設定橫印
	       // HSSFFooter footer = sheet.getFooter();
	        ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
	  		//wb.setSheetName(0,"test");
	  		finput.close();
	  		
	  		HSSFRow row=null;//宣告一列 
	  		HSSFCell cell=null;//宣告一個儲存格  		
	  		//如果你需要使用換行符,你需要設置  
  		    //單元格的樣式wrap=true,代碼如下:  
	  		HSSFCellStyle cs = wb.createCellStyle();  
  		  	cs.setWrapText(true);  
  		  	cs.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
  		  	cs.setBorderTop(HSSFCellStyle.BORDER_THIN);
  		  	cs.setBorderLeft(HSSFCellStyle.BORDER_THIN);
  		  	cs.setBorderRight(HSSFCellStyle.BORDER_THIN);
  		  	cs.setBorderBottom(HSSFCellStyle.BORDER_THIN);
  		  	HSSFFont font = wb.createFont();
  		  	font.setFontHeightInPoints((short) 12);
  		  	cs.setFont(font);
  		  	
	  		if("ALL".equals(bank_code)){
	    		if(bank_type.equals("6")){
	        		bank_name ="全體農會信用部"; 
	        	}else if(bank_type.equals("7")){
	        		bank_name ="全體漁會信用部"; 
	        	}
	    	}else if("ALL67".equals(bank_code)){
	    		bank_name ="全體農(漁)會信用部";
	    	}else{
	    		dbData = Utility.getBN01(bank_code); 
	    		if(dbData != null && dbData.size()!=0 ){
	                bank_name=(String)((DataObject)dbData.get(0)).getValue("bank_name");
	            }
	    	}
	  		
            sql.append(" select field_995210, ");//--申報當月新增.件數 
            sql.append("        round(field_995300/?,0) field_995300,");//--申報當月新增.餘額   
            sql.append("        round(field_995361/?,0) field_995361,");//--申報當月新增.建築貸款備抵呆帳餘額
            sql.append("        round(field_995400/1000/count_seq,2) field_995400,");//--申報當月新增建築貸款案件最高貸放成數 
            sql.append("        field_995200,");//--當月底建築貸款.件數 
            sql.append("        round(field_992710/?,0) field_992710,");//--當月底建築貸款.餘額
            sql.append("        round(field_992920/?,0) field_992920,");//--當月底建築貸款.建築貸款備抵呆帳餘額
            sql.append("        decode(field_992710,0,0,round(field_992920 / field_992710 *100 ,2)) as field_992920_rate,");//--當月底建築貸款.建築貸款放款覆蓋率
            sql.append("        field_995211,");//--都市計畫區.案件數
            sql.append("        field_995212,");//--非都市計畫區.案件數
            sql.append("        field_995221,");//--擔保放款.案件數
            sql.append("        field_995222,");//--無擔保放款.案件數
            sql.append("        field_995231,");//--已動工興建.案件數
            sql.append("        field_995232,");//--未動工興建.案件數
            sql.append("        field_995241,");//--金庫主辦聯貸案.案件數
            sql.append("        field_995242,");//--農(漁)會主辦聯貸案.案件數
            sql.append("        field_995243,");//--農(漁)會自貸案件.案件數
            sql.append("        field_995251,");//--建地.案件數
            sql.append("        field_995252,");//--農地.案件數
            sql.append("        field_995253,");//--其他.案件數
            sql.append("        round(field_995311/?,0) field_995311,");//--都市計畫區.餘額               
            sql.append("        round(field_995312/?,0) field_995312,");//--非都市計畫區.餘額             
            sql.append("        round(field_995321/?,0) field_995321,");//--擔保放款.餘額                 
            sql.append("        round(field_995322/?,0) field_995322,");//--無擔保放款.餘額              
            sql.append("        round(field_995331/?,0) field_995331,");//--已動工興建.餘額               
            sql.append("        round(field_995332/?,0) field_995332,");//--未動工興建.餘額               
            sql.append("        round(field_995341/?,0) field_995341,");//--金庫主辦聯貸案.餘額         
            sql.append("        round(field_995342/?,0) field_995342,");//--農(漁)會主辦聯貸案.餘額      
            sql.append("        round(field_995343/?,0) field_995343,");//--農(漁)會自貸案件.餘額        
            sql.append("        round(field_995351/?,0) field_995351,");//--建地.餘額                     
            sql.append("        round(field_995352/?,0) field_995352,");//--農地.餘額                     
            sql.append("        round(field_995353/?,0) field_995353 ");//--其他.餘額              
            sql.append(" from ");
            for(int i=0;i<16;i++){
            	paramList.add(unit) ;
            }
            sql.append(" (    ");
            sql.append("         select       COUNT(*) AS COUNT_SEQ, "); 
            sql.append("                     SUM(field_995210)  field_995210,");
            sql.append("                     SUM(field_995300)  field_995300,");
            sql.append("                     SUM(field_995361)  field_995361,"); 
            sql.append("                     SUM(field_995400)  field_995400,");
            sql.append("                     SUM(field_995200)  field_995200,");
            sql.append("                     SUM(field_995211)  field_995211,");
            sql.append("                     SUM(field_995212)  field_995212,");
            sql.append("                     SUM(field_995221)  field_995221,");
            sql.append("                     SUM(field_995222)  field_995222,");
            sql.append("                     SUM(field_995231)  field_995231,");
            sql.append("                     SUM(field_995232)  field_995232,");
            sql.append("                     SUM(field_995241)  field_995241,");
            sql.append("                     SUM(field_995242)  field_995242,");
            sql.append("                     SUM(field_995243)  field_995243,");
            sql.append("                     SUM(field_995251)  field_995251,");
            sql.append("                     SUM(field_995252)  field_995252,");
            sql.append("                     SUM(field_995253)  field_995253,");
            sql.append("                     SUM(field_995311)  field_995311,");
            sql.append("                     SUM(field_995312)  field_995312,");
            sql.append("                     SUM(field_995321)  field_995321,");
            sql.append("                     SUM(field_995322)  field_995322,"); 
            sql.append("                     SUM(field_995331)  field_995331,");
            sql.append("                     SUM(field_995332)  field_995332,");
            sql.append("                     SUM(field_995341)  field_995341,");
            sql.append("                     SUM(field_995342)  field_995342,");
            sql.append("                     SUM(field_995343)  field_995343,");
            sql.append("                     SUM(field_995351)  field_995351,");
            sql.append("                     SUM(field_995352)  field_995352,");
            sql.append("                     SUM(field_995353)  field_995353,");
            sql.append("                     SUM(field_992710)  field_992710,"); 
            sql.append("                     SUM(field_992920)  field_992920 "); 
            sql.append("              from (  select nvl(cd01.hsien_id,' ') as  hsien_id ,");  
            sql.append("                           nvl(cd01.hsien_name,'OTHER') as hsien_name,");  
            sql.append("                           cd01.FR001W_output_order  as FR001W_output_order,");                            
            sql.append("                           bn01.bank_no,bn01.BANK_NAME,");  
            sql.append("                           round(sum(decode(a13.acc_code,'995210',amt,0)) /1,0) as field_995210,");
            sql.append("                           round(sum(decode(a13.acc_code,'995300',amt,0)) /1,0) as field_995300,");
            sql.append("                           round(sum(decode(a13.acc_code,'995361',amt,0)) /1,0) as field_995361,");
            sql.append("                           round(sum(decode(a13.acc_code,'995400',amt,0)) /1,0) as field_995400,");
            sql.append("                           round(sum(decode(a13.acc_code,'995200',amt,0)) /1,0) as field_995200,");
            sql.append("                           round(sum(decode(a13.acc_code,'995211',amt,0)) /1,0) as field_995211,");
            sql.append("                           round(sum(decode(a13.acc_code,'995212',amt,0)) /1,0) as field_995212,");
            sql.append("                           round(sum(decode(a13.acc_code,'995221',amt,0)) /1,0) as field_995221,");
            sql.append("                           round(sum(decode(a13.acc_code,'995222',amt,0)) /1,0) as field_995222,");
            sql.append("                           round(sum(decode(a13.acc_code,'995231',amt,0)) /1,0) as field_995231,");
            sql.append("                           round(sum(decode(a13.acc_code,'995232',amt,0)) /1,0) as field_995232,");
            sql.append("                           round(sum(decode(a13.acc_code,'995241',amt,0)) /1,0) as field_995241,");
            sql.append("                           round(sum(decode(a13.acc_code,'995242',amt,0)) /1,0) as field_995242,");
            sql.append("                           round(sum(decode(a13.acc_code,'995243',amt,0)) /1,0) as field_995243,");
            sql.append("                           round(sum(decode(a13.acc_code,'995251',amt,0)) /1,0) as field_995251,");
            sql.append("                           round(sum(decode(a13.acc_code,'995252',amt,0)) /1,0) as field_995252,");
            sql.append("                           round(sum(decode(a13.acc_code,'995253',amt,0)) /1,0) as field_995253,");
            sql.append("                           round(sum(decode(a13.acc_code,'995311',amt,0)) /1,0) as field_995311,");
            sql.append("                           round(sum(decode(a13.acc_code,'995312',amt,0)) /1,0) as field_995312,");
            sql.append("                           round(sum(decode(a13.acc_code,'995321',amt,0)) /1,0) as field_995321,");
            sql.append("                           round(sum(decode(a13.acc_code,'995322',amt,0)) /1,0) as field_995322,");
            sql.append("                           round(sum(decode(a13.acc_code,'995331',amt,0)) /1,0) as field_995331,");
            sql.append("                           round(sum(decode(a13.acc_code,'995332',amt,0)) /1,0) as field_995332,");
            sql.append("                           round(sum(decode(a13.acc_code,'995341',amt,0)) /1,0) as field_995341,");
            sql.append("                           round(sum(decode(a13.acc_code,'995342',amt,0)) /1,0) as field_995342,");
            sql.append("                           round(sum(decode(a13.acc_code,'995343',amt,0)) /1,0) as field_995343,");
            sql.append("                           round(sum(decode(a13.acc_code,'995351',amt,0)) /1,0) as field_995351,");
            sql.append("                           round(sum(decode(a13.acc_code,'995352',amt,0)) /1,0) as field_995352,");
            sql.append("                           round(sum(decode(a13.acc_code,'995353',amt,0)) /1,0) as field_995353 ");
            sql.append("                     from (select * from cd01 where cd01.hsien_id <> 'Y') cd01  "); 
            sql.append("                     left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id  "); 
            sql.append("                     left join (select * from bn01 where m_year=?)bn01 on wlx01.bank_no=bn01.bank_no  ");
            paramList.add(u_year) ;
            paramList.add(u_year) ;
            if("ALL67".equals(bank_code)){
            	sql.append(" and bn01.bank_type in('6','7')  ");//全体農漁會信用部,才加入
            }else if("ALL".equals(bank_code)){
            	sql.append("  and bn01.bank_type in(?) ");
            	paramList.add(bank_type) ;
            }else{
            	sql.append(" and bn01.bank_no=? ");
 	            paramList.add(bank_code) ;
            }
            sql.append("                     left join (select * from a13 where a13.m_year = ? and a13.m_month = ? ) a13 on  bn01.bank_no = a13.bank_code ");
            paramList.add(S_YEAR) ;
            paramList.add(S_MONTH) ;
            sql.append("                     group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no ,  bn01.BANK_NAME  "); 
            sql.append("                  ) a13,(select bank_code,  ");
            sql.append("                                     sum(decode(a99.acc_code,'992710',amt,0)) as  field_992710, ");
            sql.append("                                     sum(decode(a99.acc_code,'992920',amt,0)) as  field_992920  ");
            sql.append("                         from a99 where a99.m_year= ? and a99.m_month  = ? "); 
            paramList.add(S_YEAR) ;
            paramList.add(S_MONTH) ;
            sql.append("                         group by bank_code ");
            sql.append("                        ) a99 ");       
            sql.append("             where   a13.bank_no=a99.bank_code(+) ");           
            sql.append("              and a13.bank_no <> ' '  ");
            sql.append(" )a13 ");
            
            dbData =DBManager.QueryDB_SQLParam(sql.toString(),paramList, "field_995210,field_995300,field_995361,field_995400,field_995200,"
													            		+"field_992710,field_992920,field_992920_rate,field_995211,"
													            		+"field_995212,field_995221,field_995222,field_995231,"
													            		+"field_995232,field_995241,field_995242,field_995243,"
													            		+"field_995251,field_995252,field_995253,field_995311,field_995312,"
													            		+"field_995321,field_995322,field_995331,field_995332,field_995341,"        
													            		+"field_995342,field_995343,field_995351,field_995352,field_995353");	 
        	System.out.println("dbData.size()="+dbData.size());
        	
        	
            
            row=sheet.getRow(0);
	  		cell=row.getCell((short)0);	       	
	  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	   	    cell.setCellValue(bank_name+"信用部建築貸款統計表");
  		  	row=sheet.getRow(1);
  		  	cell=row.getCell((short)5);	       	
  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
  		  	cell.setCellValue("單位：新台幣"+unit_name+"、％");
  		    String s_day = "";
		  	String lastDay = Utility.getLastDay(String.valueOf(Integer.parseInt(S_YEAR)+1911)+S_MONTH+"01","yyyymmdd");
		  	//System.out.println("******* lastDay="+lastDay);
		  	if(!"".equals(lastDay)){
	  		  	String[] aArray = lastDay.split(" ");
	  		    String[] bArray = aArray[0].split("-");
	  		    s_day = bArray[2];
		  	}
  		  	
  	        if(dbData.size() == 0){	
		  		row=sheet.getRow(1);
		  		cell=row.getCell((short)3);	       	
		  		//設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
		  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
   	   	       	cell.setCellValue("基準日："+S_YEAR +"年" +S_MONTH +"月"+s_day+"日無資料存在");
   	   	       	
	  	    }else{
	  	      
	  		  	row=sheet.getRow(1);
	  		  	cell=row.getCell((short)3);	       	
	  		  	//設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16); 
	  		  	cell.setCellValue("基準日："+S_YEAR +"年" +S_MONTH +"月"+s_day+"日");
	  	    	DataObject obj = (DataObject) dbData.get(0);
	  		  	
	  	    	row=sheet.getRow(3);
	  		  	cell=row.getCell((short)2);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		  	cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_995210")==null?"0":obj.getValue("field_995210").toString()));
	  		  	row=sheet.getRow(3);
	  		  	cell=row.getCell((short)3);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		  	cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_995300")==null?"0":obj.getValue("field_995300").toString()));
	  		  	row=sheet.getRow(3);
	  		  	cell=row.getCell((short)4);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_995361")==null?"0":obj.getValue("field_995361").toString()));
	  		  	row=sheet.getRow(3);
	  		  	cell=row.getCell((short)6);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		  	cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_995400")==null?"0":obj.getValue("field_995400").toString()));
	  		  	row=sheet.getRow(4);
	  		  	cell=row.getCell((short)2);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		  	cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_995200")==null?"0":obj.getValue("field_995200").toString()));
	  		      

	  		  	row=sheet.getRow(4);
	  		  	cell=row.getCell((short)3);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellStyle(cs);
	  		  	cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_992710")==null?"0":obj.getValue("field_992710").toString())+"\n (A)");
	  		  	row=sheet.getRow(4);
	  		  	cell=row.getCell((short)4);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellStyle(cs);
	  		    cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_992920")==null?"0":obj.getValue("field_992920").toString())+"\n (B)");
	  		  	row=sheet.getRow(4);
	  		  	cell=row.getCell((short)5);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		  	cell.setCellValue(obj.getValue("field_992920_rate")==null?0.00:Double.parseDouble(obj.getValue("field_992920_rate").toString()));
	  		  	
	  		  	
	  		  	row=sheet.getRow(8);
	  		  	cell=row.getCell((short)1);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_995211")==null?"0":obj.getValue("field_995211").toString()));
	  		    row=sheet.getRow(8);
	  		  	cell=row.getCell((short)2);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_995212")==null?"0":obj.getValue("field_995212").toString()));
	  		    row=sheet.getRow(8);
	  		  	cell=row.getCell((short)3);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_995221")==null?"0":obj.getValue("field_995221").toString()));
	  		    row=sheet.getRow(8);
	  		  	cell=row.getCell((short)4);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_995222")==null?"0":obj.getValue("field_995222").toString()));
	  		  	row=sheet.getRow(8);
	  		  	cell=row.getCell((short)5);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_995231")==null?"0":obj.getValue("field_995231").toString()));
	  		  	row=sheet.getRow(8);
	  		  	cell=row.getCell((short)6);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_995232")==null?"0":obj.getValue("field_995232").toString()));
	  		  	row=sheet.getRow(8);
	  		  	cell=row.getCell((short)7);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_995241")==null?"0":obj.getValue("field_995241").toString()));
	  		    row=sheet.getRow(8);
	  		  	cell=row.getCell((short)8);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_995242")==null?"0":obj.getValue("field_995242").toString()));
	  		    row=sheet.getRow(8);
	  		  	cell=row.getCell((short)9);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_995243")==null?"0":obj.getValue("field_995243").toString()));
	  		    row=sheet.getRow(8);
	  		  	cell=row.getCell((short)10);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_995251")==null?"0":obj.getValue("field_995251").toString()));
	  		    row=sheet.getRow(8);
	  		  	cell=row.getCell((short)11);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_995252")==null?"0":obj.getValue("field_995252").toString()));
	  		    row=sheet.getRow(8);
	  		  	cell=row.getCell((short)12);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_995253")==null?"0":obj.getValue("field_995253").toString()));
	  		    
	  		    row=sheet.getRow(9);
	  		  	cell=row.getCell((short)1);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_995311")==null?"0":obj.getValue("field_995311").toString()));
	  		    row=sheet.getRow(9);
	  		  	cell=row.getCell((short)2);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_995312")==null?"0":obj.getValue("field_995312").toString()));
	  		    row=sheet.getRow(9);
	  		  	cell=row.getCell((short)3);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_995321")==null?"0":obj.getValue("field_995321").toString()));
	  		    row=sheet.getRow(9);
	  		  	cell=row.getCell((short)4);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_995322")==null?"0":obj.getValue("field_995322").toString()));
	  		  	row=sheet.getRow(9);
	  		  	cell=row.getCell((short)5);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_995331")==null?"0":obj.getValue("field_995331").toString()));
	  		  	row=sheet.getRow(9);
	  		  	cell=row.getCell((short)6);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_995332")==null?"0":obj.getValue("field_995332").toString()));
	  		  	row=sheet.getRow(9);
	  		  	cell=row.getCell((short)7);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_995341")==null?"0":obj.getValue("field_995341").toString()));
	  		    row=sheet.getRow(9);
	  		  	cell=row.getCell((short)8);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_995342")==null?"0":obj.getValue("field_995342").toString()));
	  		    row=sheet.getRow(9);
	  		  	cell=row.getCell((short)9);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_995343")==null?"0":obj.getValue("field_995343").toString()));
	  		    row=sheet.getRow(9);
	  		  	cell=row.getCell((short)10);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_995351")==null?"0":obj.getValue("field_995351").toString()));
	  		    row=sheet.getRow(9);
	  		  	cell=row.getCell((short)11);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_995352")==null?"0":obj.getValue("field_995352").toString()));
	  		    row=sheet.getRow(9);
	  		  	cell=row.getCell((short)12);	       	
	  		  	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  		    cell.setCellValue(Utility.setCommaFormat(obj.getValue("field_995353")==null?"0":obj.getValue("field_995353").toString()));
	  	    
	  	    }
  	        
  	        HSSFFooter footer=sheet.getFooter();
			footer.setCenter( "Page:"+HSSFFooter.page()+" of "+HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
			
	        FileOutputStream fout=new FileOutputStream(reportDir + System.getProperty("file.separator")+"信用部建築貸款統計表.xls");
	        wb.write(fout);
	        //儲存 
	        fout.close();
	        Utility.printLogTime("FR073W_Excel end time");	
		}catch(Exception e){
			System.out.println("createRpt Error:"+e+e.getMessage());
		}
		return errMsg;
	}	
    
}
