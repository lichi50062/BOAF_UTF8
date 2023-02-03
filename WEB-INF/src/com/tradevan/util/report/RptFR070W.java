/*
104.06.03	add by 2968       
*/
package com.tradevan.util.report;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.Region;
import java.io.*;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR070W {
	public static String createRpt(String s_year,String s_month,String unit) {    
		String errMsg="";
		String unit_name=Utility.getUnitName(unit);
		int rowNum=0;
		int i=0;
		int j=0;
		String filename="全體農漁會信用部建築貸款占信用部決算淨值逾100%明細表.xls";
		List dbData1=null;
		List dbData2=null;
        //99.09.10 add 查詢年度100年以前.縣市別不同===============================
        String cd01_table = (Integer.parseInt(s_year) < 100)?"cd01_99":""; 
        String wlx01_m_year = (Integer.parseInt(s_year) < 100)?"99":"100"; 
	    //=====================================================================    
		String lyear = s_year;
		String lmonth = String.valueOf(Integer.parseInt(s_month)-1);
		if("0".equals(lmonth)){
			lmonth = "12";
			lyear = String.valueOf(Integer.parseInt(s_year)-1);
		}
		String l2year = s_year;
		String l2month = s_month;
		if("0".equals(l2month)){
			lmonth = "12";
			lyear = String.valueOf(Integer.parseInt(s_year)-1);
		}
		if((Integer.parseInt(s_month) % 2) == 0){
			l2month = String.valueOf(Integer.parseInt(s_month)-2);
			if("0".equals(l2month)){
				l2month = "12";
				l2year = String.valueOf(Integer.parseInt(s_year)-1);
			}
		}else{
			l2year = lyear;
			l2month = lmonth;
		}


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
		    //ps.setScale( ( short )65 ); //列印縮放百分比
		    ps.setLandscape( true ); // 設定橫印
		    //ps.setPaperSize( ( short )8 ); //設定紙張大小 A3
		    ps.setPaperSize( (short) 9); //設定紙張大小 A4 (A3:8/A4:9)
			
			finput.close();
			
			HSSFFont ft = wb.createFont();
            HSSFCellStyle cs = wb.createCellStyle();
            ft.setFontHeightInPoints((short)12);
            ft.setFontName("標楷體");
            cs.setFont(ft);
            cs.setBorderLeft(HSSFCellStyle.BORDER_THIN);  
            cs.setBorderTop(HSSFCellStyle.BORDER_THIN);   
            cs.setBorderRight(HSSFCellStyle.BORDER_THIN); 
            cs.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            cs.setAlignment(HSSFCellStyle.ALIGN_CENTER);
			HSSFFont ft1 = wb.createFont();
            HSSFCellStyle cs1 = wb.createCellStyle();
            ft1.setFontHeightInPoints((short)12);
            ft1.setFontName("標楷體");
            cs1.setFont(ft1);
            cs1.setBorderLeft(HSSFCellStyle.BORDER_THIN);  
            cs1.setBorderTop(HSSFCellStyle.BORDER_THIN);   
            cs1.setBorderRight(HSSFCellStyle.BORDER_THIN); 
            cs1.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            cs1.setAlignment(HSSFCellStyle.ALIGN_LEFT);
            HSSFFont ft2 = wb.createFont();
            HSSFCellStyle cs2 = wb.createCellStyle();
            ft2.setFontHeightInPoints((short)12);
            ft2.setFontName("標楷體");
            cs2.setFont(ft2);
            cs2.setBorderLeft(HSSFCellStyle.BORDER_THIN);  
            cs2.setBorderTop(HSSFCellStyle.BORDER_THIN);   
            cs2.setBorderRight(HSSFCellStyle.BORDER_THIN); 
            cs2.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            cs2.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
            HSSFCellStyle cs3 = wb.createCellStyle();
            cs3.setFont(ft1);
            cs3.setBorderLeft(HSSFCellStyle.BORDER_NONE);  
            cs3.setBorderTop(HSSFCellStyle.BORDER_THIN);   
            cs3.setBorderRight(HSSFCellStyle.BORDER_NONE); 
            cs3.setBorderBottom(HSSFCellStyle.BORDER_NONE);
            cs3.setAlignment(HSSFCellStyle.ALIGN_LEFT);
            HSSFCellStyle cs4 = wb.createCellStyle();
            cs4.setFont(ft1);
            cs4.setBorderLeft(HSSFCellStyle.BORDER_NONE);  
            cs4.setBorderTop(HSSFCellStyle.BORDER_NONE);   
            cs4.setBorderRight(HSSFCellStyle.BORDER_NONE); 
            cs4.setBorderBottom(HSSFCellStyle.BORDER_NONE);
            cs4.setAlignment(HSSFCellStyle.ALIGN_LEFT);
			HSSFRow row=null;//宣告一列
			HSSFCell cell=null;//宣告一個儲存格
	        
			//paramList.add(wlx01_m_year);
			//paramList.add(bank_type);
            
		    dbData1=getData1(s_year,s_month,unit,wlx01_m_year,l2year,l2month);
			System.out.println("清單.size()="+dbData1.size());
			dbData2=getData2(s_year,s_month,wlx01_m_year,lyear,lmonth,l2year,l2month);
			System.out.println("其餘說明文字.size()="+dbData2.size());
			String bankCnt = getBankCount(wlx01_m_year);
			System.out.println("農漁會信用部家數="+bankCnt);
			
			//列印報表名稱
			row=(sheet.getRow(1)==null)? sheet.createRow(1) : sheet.getRow(1);
			cell=row.getCell((short)0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue(s_year+"年"+s_month+"月全體農漁會信用部建築貸款占信用部決算淨值逾100%明細表");
			row=(sheet.getRow(2)==null)? sheet.createRow(2) : sheet.getRow(2);
			cell=row.getCell((short)9);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("單位：新台幣　"+unit_name+",%");
			rowNum = 4;
			
			if (dbData1 !=null && dbData1.size() !=0) {
				for(j=0;j<dbData1.size();j++){
					DataObject obj = (DataObject)dbData1.get(j);
					String bank_name = obj.getValue("bank_name")==null?"":obj.getValue("bank_name").toString();
					String field_992710 = obj.getValue("field_992710")==null?"0":obj.getValue("field_992710").toString();
					Double field_9902710_rate = obj.getValue("field_9902710_rate")==null?0.00:Double.parseDouble(obj.getValue("field_9902710_rate").toString());
					String field_992920_need = obj.getValue("field_992920_need")==null?"0":obj.getValue("field_992920_need").toString();
					String field_992910 = obj.getValue("field_992910")==null?"0":obj.getValue("field_992910").toString();
					String field_992920 = obj.getValue("field_992920")==null?"0":obj.getValue("field_992920").toString();
					String field_backup = obj.getValue("field_backup")==null?"0":obj.getValue("field_backup").toString();
					Double field_backup_over_rate = obj.getValue("field_backup_over_rate")==null?0.00:Double.parseDouble(obj.getValue("field_backup_over_rate").toString());
					Double field_backup_credit_rate = obj.getValue("field_backup_credit_rate")==null?0.00:Double.parseDouble(obj.getValue("field_backup_credit_rate").toString());
					String field_993010 = obj.getValue("field_993010")==null?"0":obj.getValue("field_993010").toString();
					String field_993110 = obj.getValue("field_993110")==null?"0":obj.getValue("field_993110").toString();
					
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
					for(int l=0;l<11;l++){
		                cell = row.createCell( (short)l);
		            }
					cell=row.getCell((short)0);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs1);
					cell.setCellValue(bank_name);
					
					cell=row.getCell((short)1);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(field_992710));
					
					cell=row.getCell((short)2);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(field_9902710_rate);
					
					cell=row.getCell((short)3);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					if(Integer.parseInt(s_month)%2==0){
						cell.setCellValue(Utility.setCommaFormat(field_992920_need));
					}else{
						cell.setCellValue("");
					}
					
					cell=row.getCell((short)4);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(field_992910));
					
					cell=row.getCell((short)5);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(field_992920));
					
					cell=row.getCell((short)6);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(field_backup));
					
					cell=row.getCell((short)7);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(field_backup_over_rate);
					
					cell=row.getCell((short)8);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(field_backup_credit_rate);
					
					cell=row.getCell((short)9);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(field_993010));
					
					cell=row.getCell((short)10);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(field_993110));
					
					rowNum++;
					
				}
			}
			String field_992710 = "0";
			String field_credit = "0";					
			Double field_992710_credit_rate = 0.00;
			String field_992710_diff = "0";
			Double range_rate = 0.00;
			String count_sum_1 = "0";
			String field_992710_range_1 = "0";
			Double field_992710_rate_1 = 0.00;
			String count_sum_2 = "0";
			String field_992710_range_2 = "0";
			Double field_992710_rate_2 = 0.00;
			String count_sum_3 = "0";
			String field_992710_range_3 = "0";
			Double field_992710_rate_3 = 0.00;
			int tmp_992920need = 0;
			String field_992910 = "0";
			String field_992920 = "0";
			if (dbData2 !=null && dbData2.size() !=0) {
				for(j=0;j<dbData2.size();j++){
					DataObject obj = (DataObject)dbData2.get(j);
					if(j==0){
						field_992710 = obj.getValue("field_992710")==null?"0":obj.getValue("field_992710").toString();
						field_credit = obj.getValue("field_credit")==null?"0":obj.getValue("field_credit").toString();					
						field_992710_credit_rate = (obj.getValue("field_992710_credit_rate")==null)?0.00:Double.parseDouble(obj.getValue("field_992710_credit_rate").toString());
						field_992710_diff = obj.getValue("field_992710_diff")==null?"0":obj.getValue("field_992710_diff").toString();
						field_992910 = obj.getValue("field_992910")==null?"0":obj.getValue("field_992910").toString();
						field_992920 = obj.getValue("field_992920")==null?"0":obj.getValue("field_992920").toString();
					}
					range_rate = (obj.getValue("range_rate")==null)?0.00:Double.parseDouble(obj.getValue("range_rate").toString());
					if(range_rate==0.01){
						count_sum_1 = obj.getValue("count_sum")==null?"0":obj.getValue("count_sum").toString();
						field_992710_range_1 = obj.getValue("field_992710_range")==null?"0":obj.getValue("field_992710_range").toString();
						field_992710_rate_1 = (obj.getValue("field_992710_rate")==null)?0.00:Double.parseDouble(obj.getValue("field_992710_rate").toString());
					}
					if(range_rate==0.02){
						count_sum_2 = obj.getValue("count_sum")==null?"0":obj.getValue("count_sum").toString();
						field_992710_range_2 = obj.getValue("field_992710_range")==null?"0":obj.getValue("field_992710_range").toString();
						field_992710_rate_2 = (obj.getValue("field_992710_rate")==null)?0.00:Double.parseDouble(obj.getValue("field_992710_rate").toString());			
					}
					if(range_rate==0.03){
						count_sum_3 = obj.getValue("count_sum")==null?"0":obj.getValue("count_sum").toString();
						field_992710_range_3 = obj.getValue("field_992710_range")==null?"0":obj.getValue("field_992710_range").toString();
						field_992710_rate_3 = (obj.getValue("field_992710_rate")==null)?0.00:Double.parseDouble(obj.getValue("field_992710_rate").toString());
					}
					if(obj.getValue("field_992920_need")!=null){
						tmp_992920need += (int)Math.round(Integer.parseInt(obj.getValue("field_992920_need").toString()));
					}
					
				}
			}else{
				if(dbData1 ==null && dbData1.size()==0){
					rowNum = 8;
				}
			}
			row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
			cell = (row.getCell((short)0)==null)? row.createCell((short)0) : row.getCell((short)0); 
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellStyle(cs3);
			cell.setCellValue("一、全體"+bankCnt+"家農漁會信用部建築貸款餘額為"+Utility.setCommaFormat(field_992710)+"億元，占放款總餘額"+Utility.setCommaFormat(field_credit)+"億元之"+field_992710_credit_rate+"％，較前月增(減)"+Utility.setCommaFormat(field_992710_diff)+"億元。");
			rowNum++;
			row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
			cell = (row.getCell((short)0)==null)? row.createCell((short)0) : row.getCell((short)0); 
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellStyle(cs4);
			cell.setCellValue("二、建築貸款占信用部上年度決算淨值比率分佈情形：");
			rowNum++;
			row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
			cell = (row.getCell((short)0)==null)? row.createCell((short)0) : row.getCell((short)0); 
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellStyle(cs4);
			cell.setCellValue("(一)未達100%計"+Utility.setCommaFormat(count_sum_1)+"家，建築貸款餘額為"+Utility.setCommaFormat(field_992710_range_1)+"億元，占全體農漁會信用部建築貸款餘額"+Utility.setCommaFormat(field_992710)+"億元之"+field_992710_rate_1+"%。");
			rowNum++;
			row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
			cell = (row.getCell((short)0)==null)? row.createCell((short)0) : row.getCell((short)0); 
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellStyle(cs4);
			cell.setCellValue("(二)100%(含)以上至未達200%計"+Utility.setCommaFormat(count_sum_2)+"家，建築貸款餘額為"+Utility.setCommaFormat(field_992710_range_2)+"億元，占全體農漁會信用部建築貸款餘額"+Utility.setCommaFormat(field_992710)+"億元之"+field_992710_rate_2+"%。");
			rowNum++;
			row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
			cell = (row.getCell((short)0)==null)? row.createCell((short)0) : row.getCell((short)0); 
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellStyle(cs4);
			cell.setCellValue("(三)200%(含)以上至未達300%計"+Utility.setCommaFormat(count_sum_3)+"家，建築貸款餘額為"+Utility.setCommaFormat(field_992710_range_3)+"億元，占全體農漁會信用部建築貸款餘額"+Utility.setCommaFormat(field_992710)+"億元之"+field_992710_rate_3+"%。");
			rowNum++;
			row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
			cell = (row.getCell((short)0)==null)? row.createCell((short)0) : row.getCell((short)0); 
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellStyle(cs4);
			cell.setCellValue("三、全體農漁會信用部建築貸款備底呆帳提列情形：");
			rowNum++;
			row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
			cell = (row.getCell((short)0)==null)? row.createCell((short)0) : row.getCell((short)0); 
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellStyle(cs4);
			
			int field992920need = (int)Math.round((double)tmp_992920need/10000);//備抵呆帳
			int field992910 = (int)Math.round((double)Integer.parseInt(field_992910)/10000);//實際提撥
			int b_a = (int)Math.round((double)(Integer.parseInt(field_992910) - tmp_992920need)/10000);//寬提
			System.out.println("**** tmp_992920need/10000="+(double)tmp_992920need/10000+" ,Math.round="+field992920need);
			System.out.println("**** field_992910/10000="+(double) Integer.parseInt(field_992910)/10000+" ,Math.round="+field992910);
			System.out.println("**** b-a="+(double) (Integer.parseInt(field_992910) - tmp_992920need)/10000+" ,Math.round="+b_a);
			String tmpStr1 = Utility.setCommaFormat(String.valueOf(field992920need));
			String tmpStr2 = Utility.setCommaFormat(String.valueOf(b_a));
			if(Integer.parseInt(s_month)%2!=0){
				tmpStr1 = "  ";
				tmpStr2 = "  ";
			}
			
			cell.setCellValue("(一)"+s_year+"年"+s_month+"月應提建築貸款備底呆帳"+tmpStr1+"萬元"
								+"，實際提撥"+Utility.setCommaFormat(String.valueOf(field992910))+"萬元"
								+"，寬提"+tmpStr2+"萬元。");
			rowNum++;
			row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
			cell = (row.getCell((short)0)==null)? row.createCell((short)0) : row.getCell((short)0); 
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellStyle(cs4);
			cell.setCellValue("(二)截至"+s_year+"年"+s_month+"月已提撥之建築貸款備底呆帳合計"+Utility.setCommaFormat(field_992920)+"萬元。");
			
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
	public static List getData1(String syear,String smonth,String unit,String wlx01_m_year,String l2year,String l2month){
		StringBuffer sql = new StringBuffer();
		List paramList = new ArrayList();//共同參數
		sql.append("select bank_no,bank_name,");
		sql.append("       round(field_992710/?,0) as field_992710,");//--本月底建築貸款餘額
		sql.append("       field_9902710_rate,");//--建築貸款占上年度決算淨值比率
		sql.append("       round(((field_992710 - field_992710_even)*backup_rate_need)/?,0) as  field_992920_need,");//--應提之建築貸款備抵呆帳(A)
		sql.append("       round(field_992910/?,0) as field_992910 ,");//--B-實際提撥之建築貸款備抵呆帳
		sql.append("       round(field_992920/?,0) as field_992920,");//--建築貸款備抵呆帳餘額
		sql.append("       round(field_BACKUP/?,0) as field_BACKUP,");//--備抵呆帳
		sql.append("       field_BACKUP_OVER_RATE,");//--備抵呆帳覆蓋率
		sql.append("       field_BACKUP_CREDIT_RATE,");//--放款覆蓋率
		sql.append("       round(field_993010 /?,0)  as field_993010,");//--專案核准之建築貸款金額
		sql.append("       round(field_993110 /?,0)  as field_993110 ");//--專案核准建築貸款提撥備抵呆帳比率      
		paramList.add(unit);
		paramList.add(unit);
		paramList.add(unit);
		paramList.add(unit);
		paramList.add(unit);
		paramList.add(unit);
		paramList.add(unit);
		sql.append("from ");
		sql.append("(    ");
		sql.append("  select a99.*,");
		sql.append("    (CASE WHEN (field_9902710_rate < 100) THEN 0.01 ");                              
		sql.append("          WHEN (field_9902710_rate >=  100 and field_9902710_rate < 200) THEN 0.02 ");                                 
		sql.append("          ELSE 0.03 END) as backup_rate_need ");
		sql.append("  from ");
		sql.append("  (select bank_no,bank_name,");
		sql.append("          field_992710,");//--本月底建築貸款餘額
		sql.append("          field_992710_even,");//--上一個雙數月.建築貸款餘額
		sql.append("          field_990230_990240,");//--信用部上年度決算後淨值
		sql.append("          decode(field_990230_990240,0,0,round(field_992710 /  field_990230_990240 *100 ,2))  as  field_9902710_rate,");//--建築貸款占上年度決算淨值比率                      
		//應提之建築貸款備抵呆帳(A)
		sql.append("          field_992910,");//B-實際提撥之建築貸款備抵呆帳
		//建築貸款備抵呆帳寬提金額(B-A)
		sql.append("          field_992920,");//--建築貸款備抵呆帳餘額      
		sql.append("          field_BACKUP,");//--備抵呆帳       
		sql.append("          decode(field_over,0,0,round(field_BACKUP / field_over *100 ,2))  as   field_BACKUP_OVER_RATE,");//--備抵呆帳覆蓋率
		sql.append("          decode(field_CREDIT,0,0,round(field_BACKUP /  field_CREDIT *100 ,2))  as   field_BACKUP_CREDIT_RATE,");//--放款覆蓋率
		sql.append("          field_993010,field_993110 ");
		sql.append("  from (select a99.m_year,a99.m_month,a99.bank_no , a99.BANK_NAME,");              
		sql.append("               a99_even.field_992710 as field_992710_even,");
		sql.append("               a99.field_992710,");              
		sql.append("               a99.field_990230 - a99.field_990240 - a99.field_992810 as field_990230_990240,");               
		sql.append("               a99.field_CREDIT,a99.field_992720,a99.field_990000 as field_over,");
		sql.append("               a99.field_992910,a99.field_992920,a99.field_BACKUP,a99.field_993010,a99.field_993110 ");                   
		sql.append("        from ");
		sql.append("        ( select a99.m_year,a99.m_month, a99.bank_no ,   a99.BANK_NAME,");
		sql.append("                 SUM(a99.field_992710) as field_992710 ,");
		sql.append("                 SUM(a99.field_992720) as field_992720 ,");
		sql.append("                 SUM(a99.field_992910) as field_992910 ,");
		sql.append("                 SUM(a99.field_992920) as field_992920 ,");
		sql.append("                 SUM(a99.field_993010) as field_993010 ,");
		sql.append("                 SUM(a99.field_993110) as field_993110 ,");
		sql.append("                 SUM(a02.field_990230) as field_990230, ");
		sql.append("                 SUM(a02.field_990240) as field_990240, ");
		sql.append("                 SUM(a99.field_992810) as field_992810, ");               
		sql.append("                 SUM(a01.field_120800) + SUM(a01.field_150300) as field_BACKUP ,");         
		sql.append("                 SUM(a01.field_120000) + SUM(a01.field_120800) + SUM(a01.field_150300) as field_CREDIT,");                
		sql.append("                 SUM(a01.field_990000) as field_990000 ");              
		sql.append("          from ");
		sql.append("               ( select a99.m_year,a99.m_month, bn01.bank_no , bn01.BANK_NAME,");           
		sql.append("                        round(sum(decode(a99.acc_code,'992710',amt,0)) /1,0) as field_992710,");
		sql.append("                        round(sum(decode(a99.acc_code,'992720',amt,0)) /1,0) as field_992720,");
		sql.append("                        round(sum(decode(a99.acc_code,'992810',amt,0)) /1,0) as field_992810,");
		sql.append("                        round(sum(decode(a99.acc_code,'992910',amt,0)) /1,0) as field_992910,");
		sql.append("                        round(sum(decode(a99.acc_code,'992920',amt,0)) /1,0) as field_992920,");
		sql.append("                        round(sum(decode(a99.acc_code,'993010',amt,0)) /1,0) as field_993010,");
		sql.append("                        round(sum(decode(a99.acc_code,'993110',amt,0)) /1,0) as field_993110 ");
		sql.append("                 from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
		sql.append("                 left join (select * from a99 ");    
		sql.append("                            where m_year=? and m_month=? ");
		paramList.add(wlx01_m_year);
		paramList.add(syear);
		paramList.add(smonth);
		sql.append("                           ) a99  on  bn01.bank_no = a99.bank_code ");
		sql.append("                 group by a99.m_year,a99.m_month,bn01.bank_no,bn01.BANK_NAME ");
		sql.append("               ) a99,");    
		sql.append("               ( select a02.m_year,a02.m_month, bn01.bank_no , bn01.BANK_NAME,");           
		sql.append("                        round(sum(decode(a02.acc_code,'990230',amt,0)) /1,0) as field_990230,");
		sql.append("                        round(sum(decode(a02.acc_code,'990240',amt,0)) /1,0) as field_990240 ");
		sql.append("                 from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
		sql.append("                 left join (select * from a02 ");
		sql.append("                            where m_year=? and m_month=? "); 
		paramList.add(wlx01_m_year);
		paramList.add(syear);
		paramList.add(smonth);
		sql.append("                           ) a02  on  bn01.bank_no = a02.bank_code ");
		sql.append("                  group by a02.m_year,a02.m_month,bn01.bank_no,bn01.BANK_NAME ");
		sql.append("               ) a02,");
		sql.append("               ( select a01.m_year,a01.m_month, bn01.bank_no , bn01.BANK_NAME,");           
		sql.append("                        round(sum(decode(a01.acc_code,'120000',amt,0)) /1,0) as field_120000,");
		sql.append("                        round(sum(decode(a01.acc_code,'120800',amt,0)) /1,0) as field_120800,");
		sql.append("                        round(sum(decode(a01.acc_code,'150300',amt,0)) /1,0) as field_150300,");                        
		sql.append("                        round(sum(decode(a01.acc_code,'990000',amt,0)) /1,0) as field_990000 ");
		sql.append("                 from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
		sql.append("                 left join (select * from a01 ");
		sql.append("                            where m_year=? and m_month=? "); 
		paramList.add(wlx01_m_year);
		paramList.add(syear);
		paramList.add(smonth);
		sql.append("                           ) a01  on  bn01.bank_no = a01.bank_code ");
		sql.append("                 group by a01.m_year,a01.m_month,bn01.bank_no,bn01.BANK_NAME ");
		sql.append("               ) a01 ");
		sql.append("               where a99.bank_no=a02.bank_no(+) and a99.bank_no = a01.bank_no(+) ");             
		sql.append("               and a99.bank_no <> ' ' "); 
		sql.append("               GROUP BY a99.m_year,a99.m_month,a99.bank_no,a99.BANK_NAME ");
		sql.append("        ) a99,");//--本月份資料       
		sql.append("        ( select a99.m_year,a99.m_month, bn01.bank_no , bn01.BANK_NAME,");           
		sql.append("                 round(sum(decode(a99.acc_code,'992710',amt,0)) /1,0) as field_992710 ");
		sql.append("          from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
		sql.append("          left join (select * from a99 ");
		sql.append("                     where m_year=? and m_month=? ");//申報年月的上個雙數月 ex:若現在為104/02,則上一個雙數月為103/12               
		paramList.add(wlx01_m_year);
		paramList.add(l2year);
		paramList.add(l2month);
		sql.append("                    ) a99  on  bn01.bank_no = a99.bank_code ");
		sql.append("          group by a99.m_year,a99.m_month,bn01.bank_no,bn01.BANK_NAME ");
		sql.append("        ) a99_even ");//--上個雙數月       
		sql.append("         where  a99.bank_no=a99_even.bank_no(+) ");        
		sql.append("       )a99 ");
		sql.append("   )a99 ");
		sql.append(")a99 ");
	    List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"bank_no,bank_name,field_992710,field_9902710_rate,field_992920_need,field_992910,field_992920,"+
	    																"field_backup,field_backup_over_rate,field_backup_credit_rate,field_993010,field_993110");
		System.out.println("dbData1.size()="+dbData.size());
		return dbData;
	}
	//農漁會信用部家數
	public static String getBankCount(String wlx01_m_year){
		String count = "";
		StringBuffer sql = new StringBuffer();
		List paramList = new ArrayList();
		sql.append("select count(*) as cnt from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2' ");
		paramList.add(wlx01_m_year);
		List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"cnt");
		if (dbData !=null && dbData.size() >0) {
			DataObject obj = (DataObject)dbData.get(0);
			count = obj.getValue("cnt")==null?"0":(obj.getValue("cnt")).toString();				   
		}
		return count;
	}
	//其餘說明文字
	public static List getData2(String syear,String smonth,String wlx01_m_year,String lyear,String lmonth,String l2year,String l2month){
		StringBuffer sql = new StringBuffer();
		List paramList = new ArrayList();//共同參數
		//一、全體
		sql.append("select round(a99.field_992710 /100000000,0)  as field_992710,");//--全体農漁會.建築貸款餘額(億元)
		sql.append("       round(a01.field_CREDIT /100000000,0)  as field_CREDIT,");//--全体農漁會.放款總餘額(億元)
		sql.append("       decode(a01.field_CREDIT,0,0,round(a99.field_992710 /  a01.field_CREDIT *100 ,2))  as  field_992710_credit_rate,");//--全体農漁會.建築貸款占放款比率
		sql.append("       round((a99.field_992710 - a99_last.field_992710) /100000000,0)  as field_992710_diff,");//--全体農漁會.建築貸款餘額較前月增減金額(億元) 
		//二、建築貸款占信用部上年度決算淨值比率分佈情形：
		sql.append("       a99_range.range_rate,");//--0.01:未達100%,0.02:100%(含)以上至未達200%,0.03:200%(含)以上未達300%
		sql.append("       a99_range.count_sum,");//--家數
		sql.append("       round(a99_range.field_992710 /100000000,0)  as field_992710_range,");//--該家數.建築貸款餘額加總(億元)        
		sql.append("       decode(a99.field_992710,0,0,round(a99_range.field_992710 / a99.field_992710 *100 ,2)) as field_992710_rate,");//--該家數占全体農漁會比率
		//三、全體農漁會信用部建築貸款備底呆帳提列情形：
		sql.append("       round(((a99_range.field_992710 - a99_range.field_992710_even)*range_rate),0) as field_992920_need,");//--應提建築貸款備底呆帳A元
		sql.append("       a99.field_992910,");//--實際提撥之建築貸款備抵呆帳(元)
		sql.append("       round(a99.field_992920 /10000,0)  as field_992920 ");//--已提撥之建築貸款備底呆帳(萬元)
		sql.append("from ");         
		sql.append("( select a99.m_year,a99.m_month,");
		sql.append("         round(sum(decode(a99.acc_code,'992710',amt,0)) /1,0) as field_992710,");
		sql.append("         round(sum(decode(a99.acc_code,'992910',amt,0)) /1,0) as field_992910,");
		sql.append("         round(sum(decode(a99.acc_code,'992920',amt,0)) /1,0) as field_992920 ");
		sql.append("from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
		sql.append("left join (select * from a99 ");
		sql.append("           where m_year=? and m_month=? "); 
		paramList.add(wlx01_m_year);
		paramList.add(syear);
		paramList.add(smonth);
		sql.append("          ) a99  on  bn01.bank_no = a99.bank_code ");
		sql.append("where a99.m_year is not null ");             
		sql.append("group by a99.m_year,a99.m_month ");
		sql.append(")a99,");
		sql.append("( select a99.m_year,a99.m_month,");
		sql.append("         round(sum(decode(a99.acc_code,'992710',amt,0)) /1,0) as field_992710 ");
		sql.append("from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
		sql.append("left join (select * from a99 ");
		sql.append("           where m_year=? and m_month=? "); 
		paramList.add(wlx01_m_year);
		paramList.add(lyear);
		paramList.add(lmonth);
		sql.append("          ) a99  on  bn01.bank_no = a99.bank_code ");
		sql.append("where a99.m_year is not null ");            
		sql.append("group by a99.m_year,a99.m_month ");
		sql.append(")a99_last,");//--上月份資料   
		sql.append("( select a01.m_year,a01.m_month,"); 
		sql.append("         round(sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /1,0) as field_CREDIT ");       
		sql.append("  from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
		sql.append("  left join (select * from a01 ");
		sql.append("             where m_year=? and m_month=? ");
		paramList.add(wlx01_m_year);
		paramList.add(syear);
		paramList.add(smonth);
		sql.append("            ) a01  on  bn01.bank_no = a01.bank_code ");
		sql.append("  where a01.m_year is not null ");              
		sql.append("  group by a01.m_year,a01.m_month ");
		sql.append(") a01,");
		sql.append("(  ");
		sql.append("select range_rate,sum(field_992710) as field_992710,sum(field_992710_even) as field_992710_even,count(*) as count_sum ");
		sql.append("from ");
		sql.append("( ");
		sql.append("  select bank_no,bank_name,field_992710,field_992710_even,field_9902710_rate,");
		sql.append("        (CASE WHEN (a99.field_9902710_rate < 100) THEN 0.01 ");                              
		sql.append("              WHEN (a99.field_9902710_rate >= 100 and field_9902710_rate < 200) THEN 0.02 ");          
		sql.append("              WHEN (a99.field_9902710_rate >= 200 and field_9902710_rate < 300) THEN 0.03 ");                       
		sql.append("              ELSE 0.00 END) as range_rate ");
		sql.append("  from (  ");             
		sql.append("        select a99.bank_no , a99.BANK_NAME,");
		sql.append("               round(a99.field_992710 /1,0)  as field_992710,");//--建築貸款餘額
		sql.append("               round(a99_even.field_992710 /1,0)  as field_992710_even,");//--上個雙數月建築貸款餘額
		sql.append("               decode(field_990230_990240,0,0,round(a99.field_992710 /  field_990230_990240 *100 ,2))  as  field_9902710_rate ");//--建築貸款占上年度決算淨值比率
		sql.append("        from ");
		sql.append("        (  ");
		sql.append("         select a99.m_year,a99.m_month, a99.bank_no ,   a99.BANK_NAME,");
		sql.append("                 SUM(a99.field_992710)    as field_992710 ,");
		sql.append("                 SUM(a02.field_990230) - SUM(a02.field_990240) - SUM(a99.field_992810) as field_990230_990240  ");
		sql.append("         from ");         
		sql.append("             ( select a99.m_year,a99.m_month, bn01.bank_no , bn01.BANK_NAME,");           
		sql.append("                      round(sum(decode(a99.acc_code,'992710',amt,0)) /1,0) as field_992710,");               
		sql.append("                      round(sum(decode(a99.acc_code,'992810',amt,0)) /1,0) as field_992810 ");
		sql.append("               from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
		sql.append("               left join (select * from a99 ");
		sql.append("                          where m_year=? and m_month=? ");  
		paramList.add(wlx01_m_year);
		paramList.add(syear);
		paramList.add(smonth);
		sql.append("                         ) a99  on  bn01.bank_no = a99.bank_code ");
		sql.append("               group by a99.m_year,a99.m_month,bn01.bank_no,bn01.BANK_NAME ");
		sql.append("             ) a99,"); 
		sql.append("             ( select a02.m_year,a02.m_month, bn01.bank_no , bn01.BANK_NAME,");           
		sql.append("                      round(sum(decode(a02.acc_code,'990230',amt,0)) /1,0) as field_990230,");
		sql.append("                      round(sum(decode(a02.acc_code,'990240',amt,0)) /1,0) as field_990240 ");
		sql.append("               from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
		sql.append("               left join (select * from a02 ");
		sql.append("                          where m_year=? and m_month=? "); 
		paramList.add(wlx01_m_year);
		paramList.add(syear);
		paramList.add(smonth);
		sql.append("                         ) a02  on  bn01.bank_no = a02.bank_code ");
		sql.append("               group by a02.m_year,a02.m_month,bn01.bank_no,bn01.BANK_NAME ");
		sql.append("             ) a02 ");
		sql.append("         where a99.bank_no=a02.bank_no(+) ");
		sql.append("         and a99.bank_no <> ' ' ");  
		sql.append("         GROUP BY a99.m_year,a99.m_month,a99.bank_no,a99.BANK_NAME ");
		sql.append("         )a99,");
		sql.append("         ( select a99.m_year,a99.m_month, bn01.bank_no , bn01.BANK_NAME,");           
		sql.append("                  round(sum(decode(a99.acc_code,'992710',amt,0)) /1,0) as field_992710 ");
		sql.append("           from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
		sql.append("           left join (select * from a99 where m_year=? and m_month=? ) a99 on bn01.bank_no = a99.bank_code ");//申報年月的上個雙數月 ex:若現在為104/02,則上一個雙數月為103/12
		paramList.add(wlx01_m_year);
		paramList.add(l2year);
		paramList.add(l2month);
		sql.append("           group by a99.m_year,a99.m_month,bn01.bank_no,bn01.BANK_NAME ");
		sql.append("         ) a99_even ");//--上個雙數月   
		sql.append("         where a99.bank_no = a99_even.bank_no(+) ");
		sql.append("   )a99 ");
		sql.append(" )a99 ");
		sql.append(" group by range_rate ");
		sql.append(")a99_range ");//--比率分佈
	    List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"field_992710,field_credit,field_992710_credit_rate,field_992710_diff,range_rate,count_sum,"+
	    																  "field_992710_range,field_992710_rate,field_992920_need,field_992910,field_992920");
		System.out.println("dbData2.size()="+dbData.size());
		return dbData;
	}
}