//104.11.13 add by 2968
//104.12.07 fix SQL by 2295
package com.tradevan.util.report;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import java.io.*;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR074W {
	public static String createRpt(String s_year,String s_month,String unit) {    
		String errMsg="";
		int rowNum=0;
		int i=0;
		int j=0;
		String filename="農舍及農地放款資料彙整表.xls";
        //99.09.10 add 查詢年度100年以前.縣市別不同===============================
        String cd01_table = (Integer.parseInt(s_year) < 100)?"cd01_99":""; 
        String wlx01_m_year = (Integer.parseInt(s_year) < 100)?"99":"100"; 
	    //=====================================================================    

		try {
			File xlsDir=new File(Utility.getProperties("xlsDir"));
			File reportDir=new File(Utility.getProperties("reportDir"));
	
			if(!xlsDir.exists()){
				if(!Utility.mkdirs(Utility.getProperties("xlsDir"))){
			   		errMsg+=Utility.getProperties("xlsDir")+"目錄新增失敗";
				}
			}
			if(!reportDir.exists()){
				if(!Utility.mkdirs(Utility.getProperties("reportDir"))){
			   		errMsg+=Utility.getProperties("reportDir")+"目錄新增失敗";
				}
			}
			
			//String openfile="建築貸款占信用部決算淨值逾100%明細表.xls";
			System.out.println("open file "+filename);
	
			FileInputStream finput=new FileInputStream(xlsDir+System.getProperty("file.separator")+filename );
	
		    //設定FileINputStream讀取Excel檔
			POIFSFileSystem fs=new POIFSFileSystem( finput );
			if(fs==null){System.out.println("open 範本檔失敗");} else System.out.println("open 範本檔成功");
			HSSFWorkbook wb=new HSSFWorkbook(fs);
			if(wb==null){System.out.println("open工作表失敗");}else System.out.println("open 工作表 成功");
			HSSFSheet sheet=wb.getSheetAt(0);//讀取第一個工作表，宣告其為sheet
			if(sheet==null){System.out.println("open sheet 失敗");}else System.out.println("open sheet 成功");
			HSSFPrintSetup ps=sheet.getPrintSetup(); //取得設定
		    //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
		    //sheet.setAutobreaks(true); //自動分頁
	
		    //設定頁面符合列印大小
		    sheet.setAutobreaks( false );
		    ps.setScale( ( short )55 ); //列印縮放百分比
		    ps.setLandscape( true ); // 設定橫印
		    //ps.setPaperSize( ( short )8 ); //設定紙張大小 A3
		    ps.setPaperSize( (short) 8); //設定紙張大小 A4 (A3:8/A4:9)
			
			finput.close();
			
			
			HSSFFont ft = wb.createFont();
			ft.setFontHeightInPoints((short)12);
            ft.setFontName("標楷體");
            
            //單間 cs cs1
            HSSFCellStyle cs = wb.createCellStyle();
            cs.setFont(ft);
            cs.setBorderLeft(HSSFCellStyle.BORDER_THIN);  
            cs.setBorderTop(HSSFCellStyle.BORDER_THIN);   
            cs.setBorderRight(HSSFCellStyle.BORDER_THIN); 
            cs.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            cs.setAlignment(HSSFCellStyle.ALIGN_LEFT);
            HSSFCellStyle cs1 = wb.createCellStyle();
            cs1.setFont(ft);
            cs1.setBorderLeft(HSSFCellStyle.BORDER_THIN);  
            cs1.setBorderTop(HSSFCellStyle.BORDER_THIN);   
            cs1.setBorderRight(HSSFCellStyle.BORDER_THIN); 
            cs1.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            cs1.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
            
            //小計cs2 cs3
            HSSFCellStyle cs2 = wb.createCellStyle();
            cs2.setFont(ft);
            cs2.setBorderLeft(HSSFCellStyle.BORDER_THIN);  
            cs2.setBorderTop(HSSFCellStyle.BORDER_THIN);   
            cs2.setBorderRight(HSSFCellStyle.BORDER_THIN); 
            cs2.setBorderBottom(HSSFCellStyle.BORDER_THICK);
            cs2.setAlignment(HSSFCellStyle.ALIGN_LEFT);
            HSSFCellStyle cs3 = wb.createCellStyle();
            cs3.setFont(ft);
            cs3.setBorderLeft(HSSFCellStyle.BORDER_THIN);  
            cs3.setBorderTop(HSSFCellStyle.BORDER_THIN);   
            cs3.setBorderRight(HSSFCellStyle.BORDER_THIN); 
            cs3.setBorderBottom(HSSFCellStyle.BORDER_THICK);
            cs3.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
            
            //總計cs4 cs5
            HSSFCellStyle cs4 = wb.createCellStyle();
            cs4.setFont(ft);
            cs4.setBorderLeft(HSSFCellStyle.BORDER_THIN);  
            cs4.setBorderTop(HSSFCellStyle.BORDER_THICK);   
            cs4.setBorderRight(HSSFCellStyle.BORDER_THIN); 
            cs4.setBorderBottom(HSSFCellStyle.BORDER_THICK);
            cs4.setAlignment(HSSFCellStyle.ALIGN_LEFT);
            HSSFCellStyle cs5 = wb.createCellStyle();
            cs5.setFont(ft);
            cs5.setBorderLeft(HSSFCellStyle.BORDER_THIN);  
            cs5.setBorderTop(HSSFCellStyle.BORDER_THICK);   
            cs5.setBorderRight(HSSFCellStyle.BORDER_THIN); 
            cs5.setBorderBottom(HSSFCellStyle.BORDER_THICK);
            cs5.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
           
			
            HSSFRow row=null;//宣告一列
			HSSFCell cell=null;//宣告一個儲存格
			
			
		    List dbData=getRPT_Data(s_year,s_month,wlx01_m_year,unit);
			System.out.println("清單.size()="+dbData.size());
			
			//列印報表名稱
			row=sheet.getRow(1);
			cell=row.getCell((short)1);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("基準日："+s_year+"年"+s_month+"月底");
			row=sheet.getRow(2);
			cell=row.getCell((short)1);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("單位："+Utility.getUnitName(unit));
			
			rowNum = 6;
			if (dbData !=null && dbData.size() !=0) {
				for(j=1;j<dbData.size();j++){
					DataObject obj = (DataObject)dbData.get(j);
					String field_seq = obj.getValue("field_seq")==null?"":obj.getValue("field_seq").toString();
					if(j>0){
						if("A90".equals(field_seq)){//A01單筆;A90小計;A99總計
							setRptVal(sheet,row,cell,cs2,cs3,obj,rowNum);
						}else{
							setRptVal(sheet,row,cell,cs,cs1,obj,rowNum);
						}
						rowNum++;
						if(j==dbData.size()-1){
							obj = (DataObject)dbData.get(0);
							setRptVal(sheet,row,cell,cs4,cs5,obj,6+dbData.size()-1);
						}
					}
				}
			}
			
			HSSFFooter footer=sheet.getFooter();
			footer.setCenter( "Page:"+HSSFFooter.page()+" of "+HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
	
			FileOutputStream fout=new FileOutputStream(reportDir+System.getProperty("file.separator")+filename);
			wb.write(fout);
			//儲存
			fout.close();
			System.out.println("儲存完成");
		}catch(Exception e) {
			System.out.println("createRpt Error:"+e+e.getMessage());
		}
		return errMsg;
	}
	
	public static void setRptVal(HSSFSheet sheet,HSSFRow row,HSSFCell cell,HSSFCellStyle leftCs,HSSFCellStyle uCs,DataObject obj,int rowNum){
		String hsien_id = obj.getValue("hsien_id")==null?"":obj.getValue("hsien_id").toString();
		String hsien_name = obj.getValue("hsien_name")==null?"":obj.getValue("hsien_name").toString();
		String fr001w_output_order = obj.getValue("fr001w_output_order")==null?"":obj.getValue("fr001w_output_order").toString();
		String bank_name = obj.getValue("bank_name")==null?"":obj.getValue("bank_name").toString();
		String count_seq = obj.getValue("count_seq")==null?"":obj.getValue("count_seq").toString();
		String field_seq = obj.getValue("field_seq")==null?"":obj.getValue("field_seq").toString();
		String field_credit = obj.getValue("field_credit")==null?"0":obj.getValue("field_credit").toString();
		String field_993310 = obj.getValue("field_993310")==null?"0":obj.getValue("field_993310").toString();
		Double field_993310_rate = obj.getValue("field_993310_rate")==null?0.00:Double.parseDouble(obj.getValue("field_993310_rate").toString());
		String field_993410 = obj.getValue("field_993410")==null?"0":obj.getValue("field_993410").toString();
		Double field_993410_rate = obj.getValue("field_993410_rate")==null?0.00:Double.parseDouble(obj.getValue("field_993410_rate").toString());
		String field_993510 = obj.getValue("field_993510")==null?"0":obj.getValue("field_993510").toString();
		Double field_993510_rate = obj.getValue("field_993510_rate")==null?0.00:Double.parseDouble(obj.getValue("field_993510_rate").toString());
		String field_993610 = obj.getValue("field_993610")==null?"0":obj.getValue("field_993610").toString();
		Double field_993610_rate = obj.getValue("field_993610_rate")==null?0.00:Double.parseDouble(obj.getValue("field_993610_rate").toString());
		String field_993310_993610 = obj.getValue("field_993310_993610")==null?"0":obj.getValue("field_993310_993610").toString();
		Double field_993310_993610_rate = obj.getValue("field_993310_993610_rate")==null?0.00:Double.parseDouble(obj.getValue("field_993310_993610_rate").toString());
		Double field_993910 = obj.getValue("field_993910")==null?0.00:Double.parseDouble(obj.getValue("field_993910").toString());
		String field_993710 = obj.getValue("field_993710")==null?"0":obj.getValue("field_993710").toString();
		Double field_993810 = obj.getValue("field_993810")==null?0.00:Double.parseDouble(obj.getValue("field_993810").toString());
		String field_994010 = obj.getValue("field_994010")==null?"0":obj.getValue("field_994010").toString();
		String field_994110 = obj.getValue("field_994110")==null?"0":obj.getValue("field_994110").toString();
		String field_994210 = obj.getValue("field_994210")==null?"0":obj.getValue("field_994210").toString();
		String field_994310 = obj.getValue("field_994310")==null?"0":obj.getValue("field_994310").toString();
		String field_994410 = obj.getValue("field_994410")==null?"0":obj.getValue("field_994410").toString();
		String field_994510 = obj.getValue("field_994510")==null?"0":obj.getValue("field_994510").toString();
		String field_994610 = obj.getValue("field_994610")==null?"0":obj.getValue("field_994610").toString();
		String field_994810 = obj.getValue("field_994810")==null?"0":obj.getValue("field_994810").toString();
		Double field_994810_rate = obj.getValue("field_994810_rate")==null?0.00:Double.parseDouble(obj.getValue("field_994810_rate").toString());
		String field_994910 = obj.getValue("field_994910")==null?"0":obj.getValue("field_994910").toString();
		Double field_994910_rate = obj.getValue("field_994910_rate")==null?0.00:Double.parseDouble(obj.getValue("field_994910_rate").toString());
		String field_995010 = obj.getValue("field_995010")==null?"0":obj.getValue("field_995010").toString();
		Double field_995010_rate = obj.getValue("field_995010_rate")==null?0.00:Double.parseDouble(obj.getValue("field_995010_rate").toString());
		String field_995110 = obj.getValue("field_995110")==null?"0":obj.getValue("field_995110").toString();
		Double field_995110_rate = obj.getValue("field_995110_rate")==null?0.00:Double.parseDouble(obj.getValue("field_995110_rate").toString());
		String field_994810_995110 = obj.getValue("field_994810_995110")==null?"0":obj.getValue("field_994810_995110").toString();
		Double field_sum_rate = obj.getValue("field_sum_rate")==null?0.00:Double.parseDouble(obj.getValue("field_sum_rate").toString()); 
		row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
		for(int l=1;l<34;l++){
            cell = row.createCell( (short)l);
        }
		
		cell=row.getCell((short)1);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(leftCs);
		if("A90".equals(field_seq)){//A01單筆;A90小計;A99總計
			cell.setCellValue(hsien_name+"("+count_seq+")小計");
		}else if("A99".equals(field_seq)){
			cell.setCellValue(hsien_name+"("+count_seq+")合計");
		}else{
			cell.setCellValue(bank_name);
		}
		
		
		cell=row.getCell((short)2);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(uCs);
		cell.setCellValue(Utility.setCommaFormat(field_credit));
		
		cell=row.getCell((short)3);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(uCs);
		cell.setCellValue(Utility.setCommaFormat(field_993310));
		
		cell=row.getCell((short)4);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(uCs);
		cell.setCellValue(field_993310_rate);
		
		cell=row.getCell((short)5);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(uCs);
		cell.setCellValue(Utility.setCommaFormat(field_993410));
		
		cell=row.getCell((short)6);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(uCs);
		cell.setCellValue(field_993410_rate);
		
		cell=row.getCell((short)7);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(uCs);
		cell.setCellValue(Utility.setCommaFormat(field_993510));
		
		cell=row.getCell((short)8);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(uCs);
		cell.setCellValue(field_993510_rate);
		
		cell=row.getCell((short)9);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(uCs);
		cell.setCellValue(Utility.setCommaFormat(field_993610));
		
		cell=row.getCell((short)10);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(uCs);
		cell.setCellValue(field_993610_rate);
	    
		cell=row.getCell((short)11);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(uCs);
		cell.setCellValue(Utility.setCommaFormat(field_993310_993610));
		
		cell=row.getCell((short)12);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(uCs);
		cell.setCellValue(field_993310_993610_rate);
		
		cell=row.getCell((short)13);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(uCs);
		cell.setCellValue(field_993910);
		
		cell=row.getCell((short)14);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(uCs);
		cell.setCellValue(Utility.setCommaFormat(field_993710));
		
		cell=row.getCell((short)15);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(uCs);
		cell.setCellValue(Utility.setCommaFormat(field_993410));
		
		cell=row.getCell((short)16);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(uCs);
		cell.setCellValue(field_993810);
		
		cell=row.getCell((short)17);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(uCs);
		cell.setCellValue(Utility.setCommaFormat(field_994010));
		
		cell=row.getCell((short)18);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(uCs);
		cell.setCellValue(Utility.setCommaFormat(field_994110));
		
		cell=row.getCell((short)19);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(uCs);
		cell.setCellValue(Utility.setCommaFormat(field_994210));
		
		cell=row.getCell((short)20);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(uCs);
		cell.setCellValue(Utility.setCommaFormat(field_994310));
		
		cell=row.getCell((short)21);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(uCs);
		cell.setCellValue(Utility.setCommaFormat(field_994410));
		
		cell=row.getCell((short)22);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(uCs);
		cell.setCellValue(Utility.setCommaFormat(field_994510));
		
		cell=row.getCell((short)23);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(uCs);
		cell.setCellValue(Utility.setCommaFormat(field_994610));
		
		cell=row.getCell((short)24);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(uCs);
		cell.setCellValue(Utility.setCommaFormat(field_994810));
		
		cell=row.getCell((short)25);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(uCs);
		cell.setCellValue(field_994810_rate);
		
		cell=row.getCell((short)26);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(uCs);
		cell.setCellValue(Utility.setCommaFormat(field_994910));
		
		cell=row.getCell((short)27);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(uCs);
		cell.setCellValue(field_994910_rate);
		
		cell=row.getCell((short)28);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(uCs);
		cell.setCellValue(Utility.setCommaFormat(field_995010));
		
		cell=row.getCell((short)29);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(uCs);
		cell.setCellValue(field_995010_rate);
		
		cell=row.getCell((short)30);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(uCs);
		cell.setCellValue(Utility.setCommaFormat(field_995110));

		cell=row.getCell((short)31);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(uCs);
		cell.setCellValue(field_995110_rate);
		
		cell=row.getCell((short)32);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(uCs);
		cell.setCellValue(Utility.setCommaFormat(field_994810_995110));
		
		cell=row.getCell((short)33);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(uCs);
		cell.setCellValue(field_sum_rate);
	}
	public static List getRPT_Data(String syear,String smonth,String wlx01_m_year,String unit){
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();//共同參數
		sqlCmd.append("select hsien_id , hsien_name,"); 
		sqlCmd.append("       FR001W_output_order,");  
		sqlCmd.append("       decode(bank_no,' ','ALL',bank_no) as bank_no ,"); 
		sqlCmd.append("       bank_name,");//--單位名稱
		sqlCmd.append("       COUNT_SEQ,");//--縣市家數小計/全体信用部家數小計  
		sqlCmd.append("       field_SEQ,");  
		sqlCmd.append("       round(field_CREDIT/?,0) as field_CREDIT,");//--放款總餘額
		sqlCmd.append("       round(field_993310/?,0) as field_993310,field_993310_rate,");//--(A)農地放款總餘額.1.以農地(空地)為擔保品.金額及比率
		sqlCmd.append("       round(field_993410/?,0) as field_993410,");//--(A)農地放款總餘額.2.以農地上有農舍之不動產為擔保品.金額及(B)以農舍為擔保品之放款案件.本金餘額小計
		sqlCmd.append("       field_993410_rate,");//--(A)農地放款總餘額.2.以農地上有農舍之不動產為擔保品.比率
		sqlCmd.append("       round(field_993510/?,0) as field_993510,field_993510_rate,");//--(A)農地放款總餘額.3.以農地上有工廠之不動產為擔保品.金額及比率
		sqlCmd.append("       round(field_993610/?,0) as field_993610,field_993610_rate,");//--(A)農地放款總餘額.4.其他.金額及比率
		sqlCmd.append("       round(field_993310_993610/?,0) as field_993310_993610,field_993310_993610_rate, ");//--(A)農地放款總餘額-合計.金額及比率
		sqlCmd.append("       field_993910, ");//--(B)以農舍為擔保品之放款案件.件數
		sqlCmd.append("       round(field_993710/?,0) as field_993710,");//--(B)以農舍為擔保品之放款案件.核准金額小計
		sqlCmd.append("       round(field_993810/1000,2) as field_993810,");//--(B)以農舍為擔保品之放款案件.平均貸放成數(%).固定除以1000
		sqlCmd.append("       field_994010,");//--(B)以農舍為擔保品之放款案件.借款用途(件).1.自用型
		sqlCmd.append("       field_994110,"); //--(B)以農舍為擔保品之放款案件.借款用途(件).2.投資型
		sqlCmd.append("       field_994210,"); //--(B)以農舍為擔保品之放款案件.所有權人(件).1.農保或健保第三類  
		sqlCmd.append("       field_994310,"); //--(B)以農舍為擔保品之放款案件.所有權人(件) .2.不具上開資格者
		sqlCmd.append("       field_994410,"); //--(B)以農舍為擔保品之放款案件.會員別(件).1.會員
		sqlCmd.append("       field_994510,"); //--(B)以農舍為擔保品之放款案件.會員別(件).2.贊助會員  
		sqlCmd.append("       field_994610,"); //--(B)以農舍為擔保品之放款案件.會員別(件).3.非會員 
		sqlCmd.append("       round(field_994810/?,0) as field_994810,field_994810_rate,"); //--(C)農地放款逾期情形.1.以農地(空地)為擔保品.金額及比率
		sqlCmd.append("       round(field_994910/?,0) as field_994910,field_994910_rate,"); //--(C)農地放款逾期情形.2.以農地上有農舍之不動產為擔保品.金額及比率
		sqlCmd.append("       round(field_995010/?,0) as field_995010,field_995010_rate,"); //--(C)農地放款逾期情形.3.以農地上有工廠之不動產為擔保品.金額及比率
		sqlCmd.append("       round(field_995110/?,0) as field_995110,field_995110_rate,"); //--(C)農地放款逾期情形.4.其他.金額及比率
		sqlCmd.append("       round(field_994810_995110/?,0) as field_994810_995110,field_sum_rate ");//--(C)農地放款逾期情形.合計.金額及比率
		for(int i=0;i<12;i++){
			paramList.add(unit);
		}
		sqlCmd.append(" from (select a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order,");//a01.bank_type,
		sqlCmd.append("              a01.bank_no , a01.bank_name, COUNT_SEQ,");  
		sqlCmd.append("              field_SEQ,");  
		sqlCmd.append("              round(field_CREDIT /1,0) as field_CREDIT,");
		sqlCmd.append("              round(field_993310 /1,0) as field_993310,");
		sqlCmd.append("              decode(field_CREDIT,0,0,round(field_993310 / field_CREDIT *100 ,2)) as field_993310_rate,");
		sqlCmd.append("              round(field_993410 /1,0) as field_993410,");
		sqlCmd.append("              decode(field_CREDIT,0,0,round(field_993410 / field_CREDIT *100 ,2)) as field_993410_rate,");
		sqlCmd.append("              round(field_993510 /1,0) as field_993510,");
		sqlCmd.append("              decode(field_CREDIT,0,0,round(field_993510 / field_CREDIT *100 ,2)) as field_993510_rate,");          
		sqlCmd.append("              round(field_993610 /1,0) as field_993610,");
		sqlCmd.append("              decode(field_CREDIT,0,0,round(field_993610 / field_CREDIT *100 ,2)) as field_993610_rate,");          
		sqlCmd.append("              round((field_993310+field_993410+field_993510+field_993610) /1,0) as field_993310_993610,");
		sqlCmd.append("              decode(field_CREDIT,0,0,round((field_993310+field_993410+field_993510+field_993610) / field_CREDIT *100 ,2)) as  field_993310_993610_rate,");
		sqlCmd.append("              round(field_993910 /1,0) as field_993910,");
		sqlCmd.append("              round(field_993710 /1,0) as field_993710,");
		sqlCmd.append("              round(field_993810 /1,0) as field_993810,");
		sqlCmd.append("              round(field_994010 /1,0) as field_994010,");
		sqlCmd.append("              round(field_994110 /1,0) as field_994110,");
		sqlCmd.append("              round(field_994210 /1,0) as field_994210,");
		sqlCmd.append("              round(field_994310 /1,0) as field_994310,");
		sqlCmd.append("              round(field_994410 /1,0) as field_994410,");
		sqlCmd.append("              round(field_994510 /1,0) as field_994510,");
		sqlCmd.append("              round(field_994610 /1,0) as field_994610,");
		sqlCmd.append("              round(field_994810 /1,0) as field_994810,");
		sqlCmd.append("              decode(field_993310,0,0,round(field_994810 / field_993310 *100 ,2)) as field_994810_rate,");          
		sqlCmd.append("              round(field_994910 /1,0) as field_994910,");
		sqlCmd.append("              decode(field_993410,0,0,round(field_994910 / field_993410 *100 ,2)) as field_994910_rate,");
		sqlCmd.append("              round(field_995010 /1,0) as field_995010,");
		sqlCmd.append("              decode(field_993510,0,0,round(field_995010 / field_993510 *100 ,2)) as field_995010_rate,");
		sqlCmd.append("              round(field_995110 /1,0) as field_995110,");
		sqlCmd.append("              decode(field_993610,0,0,round(field_995110 / field_993610 *100 ,2)) as field_995110_rate,");        
		sqlCmd.append("              round((field_994810+field_994910+field_995010+field_995110) /1,0)  as field_994810_995110,");
		sqlCmd.append("              decode((field_993310+field_993410+field_993510+field_993610) ,0,0,round((field_994810+field_994910+field_995010+field_995110) / (field_993310+field_993410+field_993510+field_993610)  *100 ,2))  as   field_sum_rate ");
		sqlCmd.append("      from (select  ' '  AS  hsien_id,'全體信用部'AS hsien_name,'001'AS FR001W_output_order,'ALL' as bank_type,");  
		sqlCmd.append("                    ' ' AS  bank_no ,' '   AS  bank_name,");  
		sqlCmd.append("                    COUNT(*) AS COUNT_SEQ,");  
		sqlCmd.append("                    'A99'  as  field_SEQ,");  
		sqlCmd.append("                    SUM(field_993310)  field_993310,");
		sqlCmd.append("                    SUM(field_993410)  field_993410,");
		sqlCmd.append("                    SUM(field_993510)  field_993510,"); 
		sqlCmd.append("                    SUM(field_993610)  field_993610,");
		sqlCmd.append("                    SUM(field_993710)  field_993710,");
		sqlCmd.append("                    SUM(field_993810)  field_993810,");
		sqlCmd.append("                    SUM(field_993910)  field_993910,");
		sqlCmd.append("                    SUM(field_994010)  field_994010,");
		sqlCmd.append("                    SUM(field_994110)  field_994110,");
		sqlCmd.append("                    SUM(field_994210)  field_994210,");
		sqlCmd.append("                    SUM(field_994310)  field_994310,");
		sqlCmd.append("                    SUM(field_994410)  field_994410,");
		sqlCmd.append("                    SUM(field_994510)  field_994510,");
		sqlCmd.append("                    SUM(field_994610)  field_994610,");
		sqlCmd.append("                    SUM(field_994810)  field_994810,");
		sqlCmd.append("                    SUM(field_994910)  field_994910,");
		sqlCmd.append("                    SUM(field_995010)  field_995010,");
		sqlCmd.append("                    SUM(field_995110)  field_995110,");
		sqlCmd.append("                    SUM(a01.field_CREDIT) field_CREDIT "); 
		sqlCmd.append("             from (  select nvl(cd01.hsien_id,' ') as  hsien_id ,");  
		sqlCmd.append("                          nvl(cd01.hsien_name,'OTHER') as hsien_name,");  
		sqlCmd.append("                          cd01.FR001W_output_order  as FR001W_output_order,");  
		sqlCmd.append("                          bn01.bank_no,bn01.BANK_NAME,");  
		sqlCmd.append("                          round(sum(decode(a13.acc_code,'993310',amt,0)) /1,0) as field_993310,");
		sqlCmd.append("                          round(sum(decode(a13.acc_code,'993410',amt,0)) /1,0) as field_993410,");
		sqlCmd.append("                          round(sum(decode(a13.acc_code,'993510',amt,0)) /1,0) as field_993510,");
		sqlCmd.append("                          round(sum(decode(a13.acc_code,'993610',amt,0)) /1,0) as field_993610,");
		sqlCmd.append("                          round(sum(decode(a13.acc_code,'993710',amt,0)) /1,0) as field_993710,");
		sqlCmd.append("                          round(sum(decode(a13.acc_code,'993810',amt,0)) /1,0) as field_993810,");
		sqlCmd.append("                          round(sum(decode(a13.acc_code,'993910',amt,0)) /1,0) as field_993910,");
		sqlCmd.append("                          round(sum(decode(a13.acc_code,'994010',amt,0)) /1,0) as field_994010,");
		sqlCmd.append("                          round(sum(decode(a13.acc_code,'994110',amt,0)) /1,0) as field_994110,");
		sqlCmd.append("                          round(sum(decode(a13.acc_code,'994210',amt,0)) /1,0) as field_994210,");
		sqlCmd.append("                          round(sum(decode(a13.acc_code,'994310',amt,0)) /1,0) as field_994310,");
		sqlCmd.append("                          round(sum(decode(a13.acc_code,'994410',amt,0)) /1,0) as field_994410,");
		sqlCmd.append("                          round(sum(decode(a13.acc_code,'994510',amt,0)) /1,0) as field_994510,");
		sqlCmd.append("                          round(sum(decode(a13.acc_code,'994610',amt,0)) /1,0) as field_994610,");
		sqlCmd.append("                          round(sum(decode(a13.acc_code,'994810',amt,0)) /1,0) as field_994810,");
		sqlCmd.append("                          round(sum(decode(a13.acc_code,'994910',amt,0)) /1,0) as field_994910,");
		sqlCmd.append("                          round(sum(decode(a13.acc_code,'995010',amt,0)) /1,0) as field_995010,");
		sqlCmd.append("                          round(sum(decode(a13.acc_code,'995110',amt,0)) /1,0) as field_995110 ");
		sqlCmd.append("                    from (select * from cd01 where cd01.hsien_id <> 'Y') cd01 "); 
		sqlCmd.append("                    left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id "); 
		sqlCmd.append("                    left join (select * from bn01 where m_year=?)bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type in('6','7') "); 
		sqlCmd.append("                    left join (select * from a13  where  a13.m_year  = ? and a13.m_month = ?) a13 on  bn01.bank_no = a13.bank_code ");
		paramList.add(wlx01_m_year);
		paramList.add(wlx01_m_year);
		paramList.add(syear);
		paramList.add(smonth);
		sqlCmd.append("                    group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no ,  bn01.BANK_NAME "); 
		sqlCmd.append("                 ) a13,(select bank_code, sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) as  field_CREDIT ");//--放款總額
		sqlCmd.append("                        from a01 where a01.m_year = ? and a01.m_month  = ? "); 
		paramList.add(syear);
		paramList.add(smonth);
		sqlCmd.append("                        group by bank_code ");
		sqlCmd.append("                       ) a01 ");      
		sqlCmd.append("            where   a13.bank_no=a01.bank_code(+) ");           
		sqlCmd.append("             and a13.bank_no <> ' ' ");
		sqlCmd.append("            ) a01 ");  
		sqlCmd.append("       UNION ALL ");  
		//--各別明細
		sqlCmd.append("       select a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order,");
		sqlCmd.append("            a01.bank_no , a01.BANK_NAME,COUNT_SEQ,");
		sqlCmd.append("            field_SEQ,");
		sqlCmd.append("            round(field_CREDIT /1,0) as field_CREDIT,");
		sqlCmd.append("            round(field_993310 /1,0) as field_993310,");
		sqlCmd.append("            decode(field_CREDIT,0,0,round(field_993310 /  field_CREDIT *100 ,2))  as field_993310_rate,");
		sqlCmd.append("            round(field_993410 /1,0) as field_993410,");
		sqlCmd.append("            decode(field_CREDIT,0,0,round(field_993410 /  field_CREDIT *100 ,2))  as field_993410_rate,");
		sqlCmd.append("            round(field_993510 /1,0) as field_993510,");
		sqlCmd.append("            decode(field_CREDIT,0,0,round(field_993510 /  field_CREDIT *100 ,2))  as field_993510_rate,");          
		sqlCmd.append("            round(field_993610 /1,0) as field_993610,");
		sqlCmd.append("            decode(field_CREDIT,0,0,round(field_993610 /  field_CREDIT *100 ,2))  as field_993610_rate,");          
		sqlCmd.append("            round((field_993310+field_993410+field_993510+field_993610) /1,0)  as field_993310_993610,");
		sqlCmd.append("            decode(field_CREDIT,0,0,round((field_993310+field_993410+field_993510+field_993610) /  field_CREDIT *100 ,2))  as  field_993310_993610_rate,");
		sqlCmd.append("            round(field_993910 /1,0) as field_993910,");
		sqlCmd.append("            round(field_993710 /1,0) as field_993710,");
		sqlCmd.append("            round(field_993810 /1,0) as field_993810,");
		sqlCmd.append("            round(field_994010 /1,0) as field_994010,");
		sqlCmd.append("            round(field_994110 /1,0) as field_994110,");
		sqlCmd.append("            round(field_994210 /1,0) as field_994210,");
		sqlCmd.append("            round(field_994310 /1,0) as field_994310,");
		sqlCmd.append("            round(field_994410 /1,0) as field_994410,");
		sqlCmd.append("            round(field_994510 /1,0) as field_994510,");
		sqlCmd.append("            round(field_994610 /1,0) as field_994610,");
		sqlCmd.append("            round(field_994810 /1,0) as field_994810,");
		sqlCmd.append("            decode(field_993310,0,0,round(field_994810 /  field_993310 *100 ,2))  as field_994810_rate,");          
		sqlCmd.append("            round(field_994910 /1,0) as field_994910,");
		sqlCmd.append("            decode(field_993410,0,0,round(field_994910 /  field_993410 *100 ,2))  as field_994910_rate,");
		sqlCmd.append("            round(field_995010 /1,0) as field_995010,");
		sqlCmd.append("            decode(field_993510,0,0,round(field_995010 /  field_993510 *100 ,2))  as field_995010_rate,");
		sqlCmd.append("            round(field_995110 /1,0) as field_995110,");
		sqlCmd.append("            decode(field_993610,0,0,round(field_995110 /  field_993610 *100 ,2))  as field_995110_rate,");        
		sqlCmd.append("            round((field_994810+field_994910+field_995010+field_995110) /1,0)  as field_994810_995110,");
		sqlCmd.append("            decode((field_993310+field_993410+field_993510+field_993610) ,0,0,round((field_994810+field_994910+field_995010+field_995110) / (field_993310+field_993410+field_993510+field_993610)  *100 ,2))  as   field_sum_rate ");    
		sqlCmd.append("       from (select a13.hsien_id ,  a13.hsien_name,  a13.FR001W_output_order,a13.bank_type,");
		sqlCmd.append("                    a13.bank_no,a13.bank_name,");
		sqlCmd.append("                     1  AS  COUNT_SEQ,");
		sqlCmd.append("                    'A01'  as  field_SEQ,");
		sqlCmd.append("                    SUM(field_993310) field_993310,");
		sqlCmd.append("                    SUM(field_993410) field_993410,");
		sqlCmd.append("                    SUM(field_993510) field_993510,"); 
		sqlCmd.append("                    SUM(field_993610) field_993610,");
		sqlCmd.append("                    SUM(field_993710) field_993710,");
		sqlCmd.append("                    SUM(field_993810) field_993810,");
		sqlCmd.append("                    SUM(field_993910) field_993910,");
		sqlCmd.append("                    SUM(field_994010) field_994010,");
		sqlCmd.append("                    SUM(field_994110) field_994110,");
		sqlCmd.append("                    SUM(field_994210) field_994210,");
		sqlCmd.append("                    SUM(field_994310) field_994310,");
		sqlCmd.append("                    SUM(field_994410) field_994410,");
		sqlCmd.append("                    SUM(field_994510) field_994510,");
		sqlCmd.append("                    SUM(field_994610) field_994610,");
		sqlCmd.append("                    SUM(field_994810) field_994810,");
		sqlCmd.append("                    SUM(field_994910) field_994910,");
		sqlCmd.append("                    SUM(field_995010) field_995010,");
		sqlCmd.append("                    SUM(field_995110) field_995110,");
		sqlCmd.append("                    SUM(a01.field_CREDIT) field_CREDIT ");
		sqlCmd.append("              from ( select nvl(cd01.hsien_id,' ')       as  hsien_id ,");
		sqlCmd.append("                            nvl(cd01.hsien_name,'OTHER')  as  hsien_name,");
		sqlCmd.append("                            cd01.FR001W_output_order     as  FR001W_output_order,bn01.bank_type,");
		sqlCmd.append("                            bn01.bank_no ,  bn01.bank_name,");
		sqlCmd.append("                            round(sum(decode(a13.acc_code,'993310',amt,0)) /1,0) as field_993310,");
		sqlCmd.append("                            round(sum(decode(a13.acc_code,'993410',amt,0)) /1,0) as field_993410,");
		sqlCmd.append("                            round(sum(decode(a13.acc_code,'993510',amt,0)) /1,0) as field_993510,");
		sqlCmd.append("                            round(sum(decode(a13.acc_code,'993610',amt,0)) /1,0) as field_993610,");
		sqlCmd.append("                            round(sum(decode(a13.acc_code,'993710',amt,0)) /1,0) as field_993710,");
		sqlCmd.append("                            round(sum(decode(a13.acc_code,'993810',amt,0)) /1,0) as field_993810,");
		sqlCmd.append("                            round(sum(decode(a13.acc_code,'993910',amt,0)) /1,0) as field_993910,");
		sqlCmd.append("                            round(sum(decode(a13.acc_code,'994010',amt,0)) /1,0) as field_994010,");
		sqlCmd.append("                            round(sum(decode(a13.acc_code,'994110',amt,0)) /1,0) as field_994110,");
		sqlCmd.append("                            round(sum(decode(a13.acc_code,'994210',amt,0)) /1,0) as field_994210,");
		sqlCmd.append("                            round(sum(decode(a13.acc_code,'994310',amt,0)) /1,0) as field_994310,");
		sqlCmd.append("                            round(sum(decode(a13.acc_code,'994410',amt,0)) /1,0) as field_994410,");
		sqlCmd.append("                            round(sum(decode(a13.acc_code,'994510',amt,0)) /1,0) as field_994510,");
		sqlCmd.append("                            round(sum(decode(a13.acc_code,'994610',amt,0)) /1,0) as field_994610,");
		sqlCmd.append("                            round(sum(decode(a13.acc_code,'994810',amt,0)) /1,0) as field_994810,");
		sqlCmd.append("                            round(sum(decode(a13.acc_code,'994910',amt,0)) /1,0) as field_994910,");
		sqlCmd.append("                            round(sum(decode(a13.acc_code,'995010',amt,0)) /1,0) as field_995010,");
		sqlCmd.append("                            round(sum(decode(a13.acc_code,'995110',amt,0)) /1,0) as field_995110 ");
		sqlCmd.append("                     from  (select * from cd01 where cd01.hsien_id <> 'Y') cd01 ");
		sqlCmd.append("                     left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id ");
		sqlCmd.append("                     left join (select * from bn01 where m_year=?)bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type in ('6','7') ");
		sqlCmd.append("                     left join (select * from a13  where  a13.m_year  = ? and a13.m_month  = ?) a13 on  bn01.bank_no = a13.bank_code ");
		paramList.add(wlx01_m_year);
		paramList.add(wlx01_m_year);
		paramList.add(syear);
		paramList.add(smonth);
		sqlCmd.append("                     group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_type,bn01.bank_no ,  bn01.BANK_NAME ");
		sqlCmd.append("                  ) a13,(select bank_code, sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) as  field_CREDIT ");//--放款總額
		sqlCmd.append("                         from a01 where a01.m_year= ? and a01.m_month  =?"); 
		paramList.add(syear);
		paramList.add(smonth);
		sqlCmd.append("                         group by bank_code ");
		sqlCmd.append("                        ) a01 ");             
		sqlCmd.append("               where   a13.bank_no=a01.bank_code(+) ");          
		sqlCmd.append("               and a13.bank_no <> ' ' ");
		sqlCmd.append("               GROUP  BY a13.hsien_id,a13.hsien_name,a13.FR001W_output_order,a13.bank_type,a13.bank_no,a13.BANK_NAME ");
		sqlCmd.append("          ) a01 ");
		sqlCmd.append("      UNION ALL ");  
		//--縣市小計
		sqlCmd.append("      select a01.hsien_id,a01.hsien_name,a01.FR001W_output_order,"); 
		sqlCmd.append("             a01.bank_no,a01.bank_name,");  
		sqlCmd.append("             COUNT_SEQ,");  
		sqlCmd.append("             field_SEQ,");  
		sqlCmd.append("             round(field_CREDIT /1,0) as field_CREDIT,");
		sqlCmd.append("             round(field_993310 /1,0) as field_993310,");
		sqlCmd.append("             decode(field_CREDIT,0,0,round(field_993310 / field_CREDIT *100 ,2)) as field_993310_rate,");
		sqlCmd.append("             round(field_993410 /1,0) as field_993410,");
		sqlCmd.append("             decode(field_CREDIT,0,0,round(field_993410 / field_CREDIT *100 ,2)) as field_993410_rate,");
		sqlCmd.append("             round(field_993510 /1,0) as field_993510,");
		sqlCmd.append("             decode(field_CREDIT,0,0,round(field_993510 / field_CREDIT *100 ,2)) as field_993510_rate,");          
		sqlCmd.append("             round(field_993610 /1,0) as field_993610,");
		sqlCmd.append("             decode(field_CREDIT,0,0,round(field_993610 / field_CREDIT *100 ,2)) as field_993610_rate,");          
		sqlCmd.append("             round((field_993310+field_993410+field_993510+field_993610) /1,0) as field_993310_993610,");
		sqlCmd.append("             decode(field_CREDIT,0,0,round((field_993310+field_993410+field_993510+field_993610) /  field_CREDIT *100 ,2)) as  field_993310_993610_rate,");
		sqlCmd.append("             round(field_993910 /1,0) as field_993910,");
		sqlCmd.append("             round(field_993710 /1,0) as field_993710,");
		sqlCmd.append("             round(field_993810 /1,0) as field_993810,");
		sqlCmd.append("             round(field_994010 /1,0) as field_994010,");
		sqlCmd.append("             round(field_994110 /1,0) as field_994110,");
		sqlCmd.append("             round(field_994210 /1,0) as field_994210,");
		sqlCmd.append("             round(field_994310 /1,0) as field_994310,");
		sqlCmd.append("             round(field_994410 /1,0) as field_994410,");
		sqlCmd.append("             round(field_994510 /1,0) as field_994510,");
		sqlCmd.append("             round(field_994610 /1,0) as field_994610,");
		sqlCmd.append("             round(field_994810 /1,0) as field_994810,");
		sqlCmd.append("             decode(field_993310,0,0,round(field_994810 / field_993310 *100 ,2)) as field_994810_rate,");          
		sqlCmd.append("             round(field_994910 /1,0) as field_994910,");
		sqlCmd.append("             decode(field_993410,0,0,round(field_994910 / field_993410 *100 ,2)) as field_994910_rate,");
		sqlCmd.append("             round(field_995010 /1,0) as field_995010,");
		sqlCmd.append("             decode(field_993510,0,0,round(field_995010 / field_993510 *100 ,2)) as field_995010_rate,");
		sqlCmd.append("             round(field_995110 /1,0) as field_995110,");
		sqlCmd.append("             decode(field_993610,0,0,round(field_995110 / field_993610 *100 ,2)) as field_995110_rate,");        
		sqlCmd.append("             round((field_994810+field_994910+field_995010+field_995110) /1,0) as field_994810_995110,");
		sqlCmd.append("             decode((field_993310+field_993410+field_993510+field_993610) ,0,0,round((field_994810+field_994910+field_995010+field_995110) / (field_993310+field_993410+field_993510+field_993610)  *100 ,2))  as   field_sum_rate ");     
		sqlCmd.append("       from ( "); 
		sqlCmd.append("             select a13.hsien_id,a13.hsien_name,a13.FR001W_output_order,");  
		sqlCmd.append("                    ' ' AS  bank_no ,' ' AS  bank_name,");  
		sqlCmd.append("                    COUNT(*) AS COUNT_SEQ,");  
		sqlCmd.append("                    'A90' as field_SEQ,");  
		sqlCmd.append("                    SUM(field_993310) field_993310 ,");
		sqlCmd.append("                    SUM(field_993410) field_993410 ,");
		sqlCmd.append("                    SUM(field_993510) field_993510 ,"); 
		sqlCmd.append("                    SUM(field_993610) field_993610 ,");
		sqlCmd.append("                    SUM(field_993710) field_993710 ,");
		sqlCmd.append("                    SUM(field_993810) field_993810,");
		sqlCmd.append("                    SUM(field_993910) field_993910,");
		sqlCmd.append("                    SUM(field_994010) field_994010,");
		sqlCmd.append("                    SUM(field_994110) field_994110,");
		sqlCmd.append("                    SUM(field_994210) field_994210,");
		sqlCmd.append("                    SUM(field_994310) field_994310,");
		sqlCmd.append("                    SUM(field_994410) field_994410,");
		sqlCmd.append("                    SUM(field_994510) field_994510,");
		sqlCmd.append("                    SUM(field_994610) field_994610,");
		sqlCmd.append("                    SUM(field_994810) field_994810,");
		sqlCmd.append("                    SUM(field_994910) field_994910,");
		sqlCmd.append("                    SUM(field_995010) field_995010,");
		sqlCmd.append("                    SUM(field_995110) field_995110,");
		sqlCmd.append("                    SUM(a01.field_CREDIT) field_CREDIT "); 
		sqlCmd.append("             from ( select nvl(cd01.hsien_id,' ') as hsien_id ,");  
		sqlCmd.append("                           nvl(cd01.hsien_name,'OTHER')as hsien_name,");  
		sqlCmd.append("                           cd01.FR001W_output_order as FR001W_output_order,bn01.bank_type,");  
		sqlCmd.append("                           bn01.bank_no,bn01.bank_name, "); 
		sqlCmd.append("                           round(sum(decode(a13.acc_code,'993310',amt,0)) /1,0) as field_993310,");
		sqlCmd.append("                           round(sum(decode(a13.acc_code,'993410',amt,0)) /1,0) as field_993410,");
		sqlCmd.append("                           round(sum(decode(a13.acc_code,'993510',amt,0)) /1,0) as field_993510,");
		sqlCmd.append("                           round(sum(decode(a13.acc_code,'993610',amt,0)) /1,0) as field_993610,");
		sqlCmd.append("                           round(sum(decode(a13.acc_code,'993710',amt,0)) /1,0) as field_993710,");
		sqlCmd.append("                           round(sum(decode(a13.acc_code,'993810',amt,0)) /1,0) as field_993810,");
		sqlCmd.append("                           round(sum(decode(a13.acc_code,'993910',amt,0)) /1,0) as field_993910,");
		sqlCmd.append("                           round(sum(decode(a13.acc_code,'994010',amt,0)) /1,0) as field_994010,");
		sqlCmd.append("                           round(sum(decode(a13.acc_code,'994110',amt,0)) /1,0) as field_994110,");
		sqlCmd.append("                           round(sum(decode(a13.acc_code,'994210',amt,0)) /1,0) as field_994210,");
		sqlCmd.append("                           round(sum(decode(a13.acc_code,'994310',amt,0)) /1,0) as field_994310,");
		sqlCmd.append("                           round(sum(decode(a13.acc_code,'994410',amt,0)) /1,0) as field_994410,");
		sqlCmd.append("                           round(sum(decode(a13.acc_code,'994510',amt,0)) /1,0) as field_994510,");
		sqlCmd.append("                           round(sum(decode(a13.acc_code,'994610',amt,0)) /1,0) as field_994610,");
		sqlCmd.append("                           round(sum(decode(a13.acc_code,'994810',amt,0)) /1,0) as field_994810,");
		sqlCmd.append("                           round(sum(decode(a13.acc_code,'994910',amt,0)) /1,0) as field_994910,");
		sqlCmd.append("                           round(sum(decode(a13.acc_code,'995010',amt,0)) /1,0) as field_995010,");
		sqlCmd.append("                           round(sum(decode(a13.acc_code,'995110',amt,0)) /1,0) as field_995110 ");
		sqlCmd.append("                    from (select * from cd01 where cd01.hsien_id <> 'Y') cd01 "); 
		sqlCmd.append("                    left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id "); 
		sqlCmd.append("                    left join (select * from bn01 where m_year=?)bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type in ('6','7') "); 
		sqlCmd.append("                    left join (select * from a13  where  a13.m_year  = ? and a13.m_month = ?) a13 on  bn01.bank_no = a13.bank_code ");
		paramList.add(wlx01_m_year);
		paramList.add(wlx01_m_year);
		paramList.add(syear);
		paramList.add(smonth);
		sqlCmd.append("                    group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_type,bn01.bank_no ,  bn01.BANK_NAME "); 
		sqlCmd.append("                   ) a13,(select bank_code, sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) as  field_CREDIT ");//--放款總額
		sqlCmd.append("                          from a01 where a01.m_year = ? and a01.m_month = ? "); 
		paramList.add(syear);
		paramList.add(smonth);
		sqlCmd.append("                          group by bank_code ");
		sqlCmd.append("                         ) a01 ");             
		sqlCmd.append("             where   a13.bank_no=a01.bank_code(+) ");          
		sqlCmd.append("             and a13.bank_no <> ' ' ");
		sqlCmd.append("             GROUP  BY a13.hsien_id,a13.hsien_name,a13.FR001W_output_order "); 
		sqlCmd.append("            ) a01 ");  
		sqlCmd.append("    )  a01  ORDER by FR001W_output_order,field_SEQ,hsien_id,bank_no ");
		
	    List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"hsien_id,hsien_name,fr001w_output_order,bank_no,bank_name,"
															    		+"count_seq,field_seq,field_credit,field_993310,field_993310_rate,"
															    		+"field_993410,field_993410_rate,field_993510,field_993510_rate,"
															    		+"field_993610,field_993610_rate,field_993310_993610,field_993310_993610_rate,"
															    		+"field_993910,field_993710,field_993810,field_994010,field_994110,field_994210,field_994310,"
															    		+"field_994410,field_994510,field_994610,field_994810,field_994810_rate,field_994910,field_994910_rate,"
															    		+"field_995010,field_995010_rate,field_995110,field_995110_rate,field_994810_995110,field_sum_rate" );
		System.out.println("dbData.size()="+dbData.size());
		return dbData;
	}
	
}