/*
104.06.11 add by 2968
111.07.19 add 建築貸款餘額-購地/興建房屋/週轉金  by 2295
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

public class RptFR071W {
	public static String createRpt(String s_year,String s_month,String unit) {    
		String errMsg="";
		String unit_name=Utility.getUnitName(unit);
		int rowNum=0;
		int j=0;
		String filename="全體農漁會信用部建築貸款明細表.xls";
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


	    reportUtil reportUtil=new reportUtil();
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
		    ps.setPaperSize( (short) 8); //設定紙張大小 A4 (A3:8/A4:9)
			
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
            HSSFCellStyle cs5 = wb.createCellStyle();
            cs5.setFont(ft1);
            cs5.setBorderLeft(HSSFCellStyle.BORDER_NONE);  
            cs5.setBorderTop(HSSFCellStyle.BORDER_NONE);   
            cs5.setBorderRight(HSSFCellStyle.BORDER_NONE); 
            cs5.setBorderBottom(HSSFCellStyle.BORDER_NONE);
            cs5.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
            HSSFCellStyle cs6 = wb.createCellStyle();
            cs6.setFont(ft1);
            cs6.setBorderLeft(HSSFCellStyle.BORDER_THIN);  
            cs6.setBorderTop(HSSFCellStyle.BORDER_THIN);   
            cs6.setBorderRight(HSSFCellStyle.BORDER_THIN); 
            cs6.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            cs6.setAlignment(HSSFCellStyle.ALIGN_LEFT);
            cs6.setWrapText(true); 
			HSSFRow row=null;//宣告一列
			HSSFCell cell=null;//宣告一個儲存格
	        
			//paramList.add(wlx01_m_year);
			//paramList.add(bank_type);
            
			dbData1=getData1(s_year,s_month,unit,wlx01_m_year,lyear,lmonth,l2year,l2month);
			System.out.println("清單.size()="+dbData1.size());
			
			//列印報表名稱
			row=(sheet.getRow(0)==null)? sheet.createRow(0) : sheet.getRow(0);
			cell=row.getCell((short)0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue(s_year+"年"+s_month+"月全體農漁會信用部建築貸款明細表");
			row=(sheet.getRow(1)==null)? sheet.createRow(1) : sheet.getRow(1);
			cell=row.getCell((short)0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("單位：新台幣　"+unit_name+",%");
			row=(sheet.getRow(2)==null)? sheet.createRow(2) : sheet.getRow(2);
			cell=row.getCell((short)16);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue(s_year+"年"+s_month+"月建築貸款逾期放款");
			row=(sheet.getRow(3)==null)? sheet.createRow(3) : sheet.getRow(3);
			cell=row.getCell((short)3);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue(lyear+"年"+lmonth+"月底建築貸款餘額(上1月)");
			cell=row.getCell((short)4);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue(s_year+"年"+s_month+"月底建築貸款餘額");
			//111.07.19 add 建築貸款餘額-購地/興建房屋/週轉金
			cell=row.getCell((short)5);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue(s_year+"年"+s_month+"月底建築貸款餘額-購地");
			cell=row.getCell((short)6);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue(s_year+"年"+s_month+"月底建築貸款餘額-興建房屋");
			cell=row.getCell((short)7);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue(s_year+"年"+s_month+"月底建築貸款餘額-週轉金");
			
			cell=row.getCell((short)8);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue(s_year+"年"+s_month+"月底較控管前(104年6月)建築貸款增減情形");
			cell=row.getCell((short)9);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue(s_year+"年"+s_month+"月底較"+lmonth+"月(上1月)建築貸款增減情形");
			cell=row.getCell((short)10);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("信用部"+(Integer.parseInt(s_year)-1)+"年決算後淨值(上1年度)");
			cell=row.getCell((short)11);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue(s_year+"年"+s_month+"月建築貸款占信用部上年度決算淨值比率(%)");
			cell=row.getCell((short)12);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue(s_year+"年"+s_month+"月建築貸款占放款比率(%)");
			cell=row.getCell((short)13);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue(s_year+"年"+s_month+"月建築貸款佔存款比率(%)");
			cell=row.getCell((short)19);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue(s_year+"年"+s_month+"月增提之建築貸款備抵呆帳");
			cell=row.getCell((short)20);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("截至"+s_year+"年"+s_month+"月已提撥之建築貸款備抵呆帳合計(建築貸款備抵呆帳餘額)");
			
			
			if (dbData1 !=null && dbData1.size() !=0) {
				rowNum = 4;
				String bank_no = "";
				String bank_name = "";
				String field_992710_10308 ="0";
				String field_992710_last ="0";
				String field_992710 ="0";
				String field_992711 ="0";//111.07.19 add
				String field_992712 ="0";//111.07.19 add
				String field_992713 ="0";//111.07.19 add
				String field_992710_diff_10308 ="0";
				String field_992710_diff_last ="0";
				String field_990230_990240 ="0";
				Double field_9902710_rate =0.00;
				Double field_9902710_credit_rate =0.00;
				Double field_9902710_220000_rate =0.00;
				//String field_992720 ="0";
				String field_over ="0";
				String field_over_rate = "";
				String field_992720 = "0";
				String field_992730 = "0";
				String field_992720_rate = "";
				//String field_992920_need ="0";
				String field_992910 ="0";
				//String b_a ="0";
				String build_baddebt_sum ="0";
				String field_baddebt_rate ="";
				String field_backup ="0";
				String field_backup_10308 = "0";
				String field_backup_diff ="0";
				String field_backup_over_rate ="";
				String field_backup_credit_rate ="";
				Double field_captial_rate =0.00;
				
				for(j=0;j<dbData1.size();j++){
					DataObject obj = (DataObject)dbData1.get(j);
					bank_no = obj.getValue("bank_no")==null?"":obj.getValue("bank_no").toString();
					bank_name = obj.getValue("bank_name")==null?"":obj.getValue("bank_name").toString();
					//System.out.println("bank_no="+bank_no);
					field_992710_10308 = obj.getValue("field_992710_10308")==null?"0":obj.getValue("field_992710_10308").toString();
					field_992710_last = obj.getValue("field_992710_last")==null?"0":obj.getValue("field_992710_last").toString();
					field_992710 = obj.getValue("field_992710")==null?"0":obj.getValue("field_992710").toString();
					field_992711 = obj.getValue("field_992711")==null?"0":obj.getValue("field_992711").toString();//111.07.19 add
					field_992712 = obj.getValue("field_992712")==null?"0":obj.getValue("field_992712").toString();//111.07.19 add
					field_992713 = obj.getValue("field_992713")==null?"0":obj.getValue("field_992713").toString();//111.07.19 add
					//System.out.println("field_992710="+field_992710+":field_992711="+field_992711+":field_992712="+field_992712+":field_992713="+field_992713);
					field_992710_diff_10308 = obj.getValue("field_992710_diff_10308")==null?"0":obj.getValue("field_992710_diff_10308").toString();
					field_992710_diff_last = obj.getValue("field_992710_diff_last")==null?"0":obj.getValue("field_992710_diff_last").toString();
					field_990230_990240 = obj.getValue("field_990230_990240")==null?"0":obj.getValue("field_990230_990240").toString();
					field_9902710_rate = obj.getValue("field_9902710_rate")==null?0.00:Double.parseDouble(obj.getValue("field_9902710_rate").toString());
					field_9902710_credit_rate = obj.getValue("field_9902710_credit_rate")==null?0.00:Double.parseDouble(obj.getValue("field_9902710_credit_rate").toString());
					field_9902710_220000_rate = obj.getValue("field_9902710_220000_rate")==null?0.00:Double.parseDouble(obj.getValue("field_9902710_220000_rate").toString());
					//field_992720 = obj.getValue("field_992720")==null?"0":obj.getValue("field_992720").toString();
					field_over = obj.getValue("field_over")==null?"0":obj.getValue("field_over").toString();
					field_over_rate = (obj.getValue("field_over_rate")==null)?"": obj.getValue("field_over_rate").toString();
					field_992720 = obj.getValue("field_992720")==null?"0":obj.getValue("field_992720").toString();
					field_992730 = obj.getValue("field_992730")==null?"0":obj.getValue("field_992730").toString();
					field_992720_rate = (obj.getValue("field_992720_rate")==null)?"": obj.getValue("field_992720_rate").toString();
					//field_992920_need = obj.getValue("field_992920_need")==null?"0":obj.getValue("field_992920_need").toString();
					field_992910 = obj.getValue("field_992910")==null?"0":obj.getValue("field_992910").toString();
					//b_a =obj.getValue("b_a")==null?"0":obj.getValue("b_a").toString();
					build_baddebt_sum = obj.getValue("build_baddebt_sum")==null?"0":obj.getValue("build_baddebt_sum").toString();
					field_baddebt_rate = (obj.getValue("field_baddebt_rate")==null)?"0.00":obj.getValue("field_baddebt_rate").toString();
					field_backup = obj.getValue("field_backup")==null?"0":obj.getValue("field_backup").toString();
					field_backup_10308 = obj.getValue("field_backup_10308")==null?"0":obj.getValue("field_backup_10308").toString();
					field_backup_diff = obj.getValue("field_backup_diff")==null?"0":obj.getValue("field_backup_diff").toString();
					field_backup_over_rate = (obj.getValue("field_backup_over_rate")==null)?"":obj.getValue("field_backup_over_rate").toString();
					field_backup_credit_rate = (obj.getValue("field_backup_credit_rate")==null)?"":obj.getValue("field_backup_credit_rate").toString();
					field_captial_rate = obj.getValue("field_captial_rate")==null?0.00:Double.parseDouble(obj.getValue("field_captial_rate").toString());
					//System.out.println("******j="+j);
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
					for(int l=0;l<28;l++){
		                cell = row.createCell( (short)l);
		            }
					cell=row.getCell((short)0);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs1);
					cell.setCellValue(bank_no);
					cell=row.getCell((short)1);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs1);
					cell.setCellValue(bank_name);
					cell=row.getCell((short)2);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(field_992710_10308));//控管前(104年6月)建築貸款餘額
					cell=row.getCell((short)3);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(field_992710_last));//○年○-1月底建築貸款餘額(上1月)
					cell=row.getCell((short)4);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(field_992710));//○年○月底建築貸款餘額
					//111.07.19 add 建築貸款餘額-購地/興建房屋/週轉金
					cell=row.getCell((short)5);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(field_992711));//○年○月底建築貸款餘額-購地
					
					cell=row.getCell((short)6);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(field_992712));//○年○月底建築貸款餘額-興建房屋
					
					cell=row.getCell((short)7);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(field_992713));//○年○月底建築貸款餘額-週轉金
					
					cell=row.getCell((short)8);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(field_992710_diff_10308));//○年○月底較控管前(104年6月)建築貸款增減情形
					cell=row.getCell((short)9);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(field_992710_diff_last));//建築貸款餘額較上月底增減情形
					cell=row.getCell((short)10);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(field_990230_990240));//信用部○年決算後淨值(上1年度)
					cell=row.getCell((short)11);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(field_9902710_rate);//○年○月建築貸款占信用部上年度決算淨值比率(%)
					cell=row.getCell((short)12);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(field_9902710_credit_rate);//○年○月建築貸款占放款比率(%)
					cell=row.getCell((short)13);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(field_9902710_220000_rate);//○年○月建築貸款佔存款比率(%)
					//cell=row.getCell((short)10);
					//cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					//cell.setCellStyle(cs2);
					//cell.setCellValue(Utility.setCommaFormat(field_992720));//建築貸款逾放金額
					cell=row.getCell((short)14);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(field_over));//逾放金額
					cell=row.getCell((short)15);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					if(!"NA".equals(field_over_rate)){
						cell.setCellValue(Double.parseDouble(field_over_rate));
					}else{
						cell.setCellValue(field_over_rate);//逾放比率
					}
					cell=row.getCell((short)16);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(field_992720));//建築貸款逾放金額
					cell=row.getCell((short)17);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(field_992730));//建築貸款應予觀察放款
					cell=row.getCell((short)18);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					if(!"NA".equals(field_992720_rate)){
						cell.setCellValue(Double.parseDouble(field_992720_rate));
					}else{
						cell.setCellValue(field_992720_rate);//建築貸款逾放比率
					}
					cell=row.getCell((short)19);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(field_992910));//○年○月增提之建築貸款備抵呆帳(實際提撥之建築貸款備抵呆帳)
					/*cell=row.getCell((short)15);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					if(Integer.parseInt(s_month)%2==0){
						cell.setCellValue(Utility.setCommaFormat(b_a));//截至○年○月已提撥之建築貸款備抵呆帳合計(建築貸款備抵呆帳餘額)
					}else{
						cell.setCellValue("");
					}*/
					cell=row.getCell((short)20);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(build_baddebt_sum));//截至○年○月已提撥之建築貸款備抵呆帳合計(建築貸款備抵呆帳餘額)
					cell=row.getCell((short)21);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					if(!"NA".equals(field_baddebt_rate)){
						cell.setCellValue(Double.parseDouble(field_baddebt_rate));
					}else{
						cell.setCellValue(field_baddebt_rate);//已提撥之建築貸款備抵呆帳占建築貸款比率
					}
					cell=row.getCell((short)22);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(field_backup));//備抵呆帳
					cell=row.getCell((short)23);
					cell.setCellStyle(cs2);;
					cell.setCellValue(Utility.setCommaFormat(field_backup_10308));//控管前備抵呆帳
					cell=row.getCell((short)24);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(Utility.setCommaFormat(field_backup_diff));//較控管前備抵呆帳增加金額
					cell=row.getCell((short)25);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					if(!"NA".equals(field_backup_over_rate)){
						cell.setCellValue(Double.parseDouble(field_backup_over_rate));
					}else{
						cell.setCellValue(field_backup_over_rate);//備抵呆帳覆蓋率
					}
					cell=row.getCell((short)26);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					if(!"NA".equals(field_backup_credit_rate)){
						cell.setCellValue(Double.parseDouble(field_backup_credit_rate));
					}else{
						cell.setCellValue(field_backup_credit_rate);//放款覆蓋率
					}
					cell=row.getCell((short)27);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs2);
					cell.setCellValue(field_captial_rate);//淨值占風險性資產比率(%)
					rowNum++;
				}
				rowNum++;
			}else{
				rowNum = 7;
			}
			row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
			cell=row.createCell((short)0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellStyle(cs4);
			cell.setCellValue("備註:(A)應提之建築貸款備抵呆帳=(雙數月建築貸款餘額-上1雙數月建築貸款餘額)*所對應之比率");
			rowNum++;
			row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
			cell=row.createCell((short)0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellStyle(cs5);
			cell.setCellValue("比率如下:");
			cell=row.createCell((short)1);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellStyle(cs6);
			cell.setCellValue("建築貸款占信用部上年度淨算淨值");
			cell=row.createCell((short)2);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellStyle(cs1);
			cell.setCellValue("");
			sheet.addMergedRegion( new Region( ( short )rowNum, ( short )1, ( short )rowNum, ( short )2 ) );
			//do 1,2 merge
			cell=row.createCell((short)3);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellStyle(cs6);
			cell.setCellValue("提列備抵呆帳比率");
			rowNum++;
			row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
			cell=row.createCell((short)1);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellStyle(cs1);
			cell.setCellValue("未逾100%");
			cell=row.createCell((short)2);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellStyle(cs1);
			cell.setCellValue("");
			sheet.addMergedRegion( new Region( ( short )rowNum, ( short )1, ( short )rowNum, ( short )2 ) );
			cell=row.createCell((short)3);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellStyle(cs2);
			cell.setCellValue("1.5%");
			rowNum++;
			row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
			cell=row.createCell((short)1);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellStyle(cs1);
			cell.setCellValue("100%(含)至未逾200%");
			cell=row.createCell((short)2);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellStyle(cs1);
			cell.setCellValue("");
			sheet.addMergedRegion( new Region( ( short )rowNum, ( short )1, ( short )rowNum, ( short )2 ) );
			cell=row.createCell((short)3);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellStyle(cs2);
			cell.setCellValue("2%");
			rowNum++;
			row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
			cell=row.createCell((short)1);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellStyle(cs1);
			cell.setCellValue("200%(含)以上");
			cell=row.createCell((short)2);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellStyle(cs1);
			cell.setCellValue("");
			sheet.addMergedRegion( new Region( ( short )rowNum, ( short )1, ( short )rowNum, ( short )2 ) );
			cell=row.createCell((short)3);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellStyle(cs2);
			cell.setCellValue("3%");
			rowNum++;

			
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
	public static List getData1(String syear,String smonth,String unit,String wlx01_m_year,String lyear,String lmonth,String l2year,String l2month){
		StringBuffer sql = new StringBuffer();
		List paramList = new ArrayList();//共同參數
		sql.append("select a99.bank_no,a99.bank_name,");
		sql.append("       round(field_992710_10308/?,0) as field_992710_10308,");//--控管前(104年6月)建築貸款餘額
		sql.append("       round(field_992710_last/?,0) as field_992710_last,");//--上月底建築貸款餘額
		sql.append("       round(field_992710/?,0) as field_992710,");//--本月底建築貸款餘額
		sql.append("       round(field_992711/?,0) as field_992711,");//--本月底建築貸款餘額-購地 111.07.19 add
		sql.append("       round(field_992712/?,0) as field_992712,");//--本月底建築貸款餘額-興建房屋 111.07.19 add
		sql.append("       round(field_992713/?,0) as field_992713,");//--本月底建築貸款餘額-週轉金 111.07.19 add
		sql.append("       round(field_992710_diff_10308/?,0) as field_992710_diff_10308,");//--建築貸款餘額較104年6月增減情形 
		sql.append("       round(field_992710_diff_last/?,0) as field_992710_diff_last,");//--建築貸款餘額較上月底增減情形
		sql.append("       round(field_990230_990240/?,0) as field_990230_990240,");//--信用部上年度決算後淨值
		sql.append("       field_9902710_rate,");//--建築貸款占上年度決算淨值比率
		sql.append("       field_9902710_credit_rate,");//--建築貸款占放款比率
		sql.append("       field_9902710_220000_rate,");//--建築貸款占存款比率
		//sql.append("       round(field_992720/?,0) as field_992720,");//--建築貸款逾放金額105.03.07 此欄位取消
		sql.append("       round(field_over/?,0) as field_over,");//--信用部逾期放款
		sql.append("       field_OVER_RATE,");//--逾放比率
		sql.append("       round(field_992720/?,0) as field_992720,");//--建築貸款逾放金額 105.06.20 增加顯示
		sql.append("       round(field_992730/?,0) as field_992730,");//--建築貸款應予觀察放款 105.06.20 增加顯示
		sql.append("       decode(field_992710,0,'NA',round(field_992720 / field_992710 *100 ,2))  as  field_992720_RATE,");//--建築貸款逾放比率105.06.20 add
		//sql.append("       round(((field_992710 - field_992710_even)*backup_rate_need)/?,0) as  field_992920_need,");//--應提之建築貸款備抵呆帳(A)105.03.07 此欄位取消
		sql.append("       round(field_992910/?) as field_992910 ,");//--B-實際提撥之建築貸款備抵呆帳
		sql.append("       round((field_992910 -round(((field_992710 - field_992710_even)*backup_rate_need)/1,0))/?,0) as b_a,");// --建築貸款備抵呆帳寬提金額(B-A)
		//sql.append("       round(field_992920/?,0) as field_992920,");//--建築貸款備抵呆帳餘額 //105.07.05 add 取消此欄位,改另一個公式
		//sql.append("       field_992920_RATE,");//--已提撥之建築貸款備抵呆帳占建築貸款比率 //105.07.05 add 取消此欄位,改另一個公式
		sql.append("       round(build_baddebt_sum/?,0) as build_baddebt_sum,");//--建築貸款備抵呆帳餘額A10.990025 105.07.05 add
		sql.append("       field_baddebt_RATE,");//--已提撥之建築貸款備抵呆帳占建築貸款比率 105.07.05 add
		sql.append("       round(field_BACKUP/?,0) as field_BACKUP,");//--備抵呆帳
		sql.append("       round(field_BACKUP_10308/?,0) as field_BACKUP_10308,");//--控管前備抵呆帳備抵呆帳 105.03.12 add
		sql.append("       round(field_BACKUP_diff/?,0) as field_BACKUP_diff,");//--備抵呆帳較控管前增加金額
		sql.append("       field_BACKUP_OVER_RATE,");//--備抵呆帳覆蓋率
		sql.append("       field_BACKUP_CREDIT_RATE,");//--放款覆蓋率
		sql.append("       field_CAPTIAL_RATE ");//--淨值占風險性資產比率
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
		//paramList.add(unit);
		sql.append("from ");
		sql.append("( ");
		sql.append("  select a99.*,");
		sql.append("    (CASE WHEN (field_9902710_rate < 100) THEN 0.015 ");//105.03.07 利率調整為1.5%                              
		sql.append("          WHEN (field_9902710_rate >=  100 and field_9902710_rate < 200) THEN 0.02 ");                                 
		sql.append("          ELSE 0.03 END) as backup_rate_need ");
		sql.append("  from ");
		sql.append("  (select bank_no,bank_name,");
		sql.append("          field_992710_10308,");//--控管前(104年6月)建築貸款餘額
		sql.append("          field_992710_last,");//--上月底建築貸款餘額
		sql.append("          field_992710,");//--本月底建築貸款餘額
		sql.append("          field_992711,");//--本月底建築貸款餘額-購地 111.07.19 add
		sql.append("          field_992712,");//--本月底建築貸款餘額-興建房屋 111.07.19 add
		sql.append("          field_992713,");//--本月底建築貸款餘額-週轉金 111.07.19 add
		sql.append("          field_992710_even,");//--上一個雙數月.建築貸款餘額
		sql.append("          field_992710_diff_10308,");//--建築貸款餘額較104年6月增減情形 
		sql.append("          field_992710_diff_last,");//--建築貸款餘額較上月底增減情形
		sql.append("          field_990230_990240,");//--信用部上年度決算後淨值
		sql.append("          decode(field_990230_990240,0,0,round(field_992710 /  field_990230_990240 *100 ,2))  as  field_9902710_rate,");//--建築貸款占上年度決算淨值比率
		sql.append("          decode(field_CREDIT,0,0,round(field_992710 /  field_CREDIT *100 ,2))  as  field_9902710_credit_rate,");//--建築貸款占放款比率
		sql.append("          decode(field_220000,0,0,round(field_992710 /  field_220000 *100 ,2))  as  field_9902710_220000_rate,");//--建築貸款占存款比率
		sql.append("          field_992720,");//--建築貸款逾放金額
		sql.append("          field_992730,");//--建築貸款應予觀察放款 105.06.20 add
		sql.append("          field_over,");//--信用部逾期放款
		sql.append("          decode(field_CREDIT,0,'NA',to_char(round(field_over /  field_CREDIT *100 ,2),'fm9999990.00'))  as  field_OVER_RATE,");//--逾放比率
		//--應提之建築貸款備抵呆帳(A)
		sql.append("          field_992910,");//--B-實際提撥之建築貸款備抵呆帳
		//--建築貸款備抵呆帳寬提金額(B-A)
		//sql.append("          field_992920,");//--建築貸款備抵呆帳餘額 //105.07.05 add 取消此欄位,改公式
		//sql.append("          decode(field_992710,0,'NA',to_char(round(field_992920 /  field_992710 *100 ,2),'fm9999990.00'))  as  field_992920_RATE,");//--已提撥之建築貸款備抵呆帳占建築貸款比率 //105.07.05 add 取消此欄位,改公式
		sql.append("          (build1_baddebt+build2_baddebt+build3_baddebt+build4_baddebt) as build_baddebt_sum,");//--建築貸款備抵呆帳餘額A10.990025 105.07.05 add
		sql.append("          decode(field_992710,0,'NA',to_char(round((build1_baddebt+build2_baddebt+build3_baddebt+build4_baddebt) / field_992710 *100 ,2),'fm9999990.00'))  as  field_baddebt_RATE,");//--已提撥之建築貸款備抵呆帳占建築貸款比率105.06.08 add 105.07.05 調整公式
		sql.append("          field_BACKUP,");//--備抵呆帳
		sql.append("          field_BACKUP_10308,");//--控管前備抵呆帳備抵呆帳 105.03.14 add
		sql.append("          field_BACKUP_diff,");//--備抵呆帳較控管前增加金額
		sql.append("          decode(field_over,0,'NA',to_char(round(field_BACKUP / field_over *100 ,2),'fm9999990.00'))  as   field_BACKUP_OVER_RATE,");//--備抵呆帳覆蓋率
		sql.append("          decode(field_CREDIT,0,'NA',to_char(round(field_BACKUP /  field_CREDIT *100 ,2),'fm9999990.00'))  as   field_BACKUP_CREDIT_RATE,");//--放款覆蓋率
		sql.append("          round(decode(field_910400,0,0,round(field_910400 * 100000 /field_910500,0)) /  1000 ,2)  as   field_CAPTIAL_RATE ");//--淨值占風險性資產比率
		sql.append("  from (select a99.m_year,a99.m_month,a99.bank_no , a99.BANK_NAME,");
		sql.append("               a99_10308.field_992710 as field_992710_10308,");
		sql.append("               a99_last.field_992710 as field_992710_last,");
		sql.append("               a99_even.field_992710 as field_992710_even,");
		sql.append("               a99.field_992710,");
		sql.append("               a99.field_992711,a99.field_992712,a99.field_992713,");//--111.07.19 add
		sql.append("               a99.field_992730,");//--105.06.20 add
		sql.append("               a99.field_992710 - a99_10308.field_992710  as field_992710_diff_10308,");
		sql.append("               a99.field_992710 - a99_last.field_992710 as field_992710_diff_last,");
		sql.append("               a99.field_990230 - a99.field_990240 - a99.field_992810 as field_990230_990240,");
		//--建築貸款占上年度決算淨值比率
		sql.append("               a99.field_CREDIT,a99.field_220000,a99.field_992720,a99.field_990000 as field_over,");
		sql.append("               a99.field_992910,a99.field_992920,a99.field_BACKUP,");
		sql.append("               a99_10308.field_BACKUP as field_BACKUP_10308,");//--105.03.14 add
		sql.append("               a99.field_BACKUP - a99_10308.field_BACKUP as  field_BACKUP_diff,");
		sql.append("               a99.field_910400,field_910500, ");
		sql.append("               a99.build1_baddebt,a99.build2_baddebt,a99.build3_baddebt,a99.build4_baddebt ");//105.07.05 add
		sql.append("        from ");
		sql.append("        ( select a99.m_year,a99.m_month, a99.bank_no ,   a99.BANK_NAME,");
		sql.append("                 SUM(a99.field_992710) as field_992710 ,");
		sql.append("                 SUM(a99.field_992711) as field_992711 ,");//111.07.19 add
		sql.append("                 SUM(a99.field_992712) as field_992712 ,");//111.07.19 add
		sql.append("                 SUM(a99.field_992713) as field_992713 ,");//111.07.19 add
		sql.append("                 SUM(a99.field_992720) as field_992720 ,");
		sql.append("                 SUM(a99.field_992730) as field_992730 ,");//105.06.20 add
		sql.append("                 SUM(a99.field_992910) as field_992910 ,");
		sql.append("                 SUM(a99.field_992920) as field_992920 ,");
		sql.append("                 SUM(a02.field_990230) as field_990230, ");
		sql.append("                 SUM(a02.field_990240) as field_990240, ");
		sql.append("                 SUM(a99.field_992810) as field_992810, ");
		sql.append("                 SUM(a01.field_120000) as field_120000, ");
		sql.append("                 SUM(a01.field_120800) + SUM(a01.field_150300) as field_BACKUP ,");         
		sql.append("                 SUM(a01.field_120000) + SUM(a01.field_120800) + SUM(a01.field_150300) as field_CREDIT,");
		sql.append("                 SUM(a01.field_220000) as field_220000,");
		sql.append("                 SUM(a01.field_990000) as field_990000,");
		sql.append("                 SUM(a05.field_910400) as field_910400,");
		sql.append("                 SUM(a05.field_910500) as field_910500,"); 
		sql.append("                 SUM(a10.build1_baddebt) as build1_baddebt,");//105.07.05 add
		sql.append("                 SUM(a10.build2_baddebt) as build2_baddebt,");//105.07.05 add
		sql.append("                 SUM(a10.build3_baddebt) as build3_baddebt,");//105.07.05 add
		sql.append("                 SUM(a10.build4_baddebt) as build4_baddebt ");//105.07.05 add 
		sql.append("          from ");
		sql.append("               ( select a99.m_year,a99.m_month, bn01.bank_no , bn01.BANK_NAME,");           
		sql.append("                        round(sum(decode(a99.acc_code,'992710',amt,0)) /1,0) as field_992710,");
		sql.append("                        round(sum(decode(a99.acc_code,'992711',amt,0)) /1,0) as field_992711,");//111.07.19 add
		sql.append("                        round(sum(decode(a99.acc_code,'992712',amt,0)) /1,0) as field_992712,");//111.07.19 add
		sql.append("                        round(sum(decode(a99.acc_code,'992713',amt,0)) /1,0) as field_992713,");//111.07.19 add
		sql.append("                        round(sum(decode(a99.acc_code,'992720',amt,0)) /1,0) as field_992720,");
		sql.append("                        round(sum(decode(a99.acc_code,'992730',amt,0)) /1,0) as field_992730,");//105.06.20 add
		sql.append("                        round(sum(decode(a99.acc_code,'992810',amt,0)) /1,0) as field_992810,");
		sql.append("                        round(sum(decode(a99.acc_code,'992910',amt,0)) /1,0) as field_992910,");
		sql.append("                        round(sum(decode(a99.acc_code,'992920',amt,0)) /1,0) as field_992920 ");
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
		sql.append("                        round(sum(decode(a01.acc_code,'220000',amt,0)) /1,0) as field_220000,");
		sql.append("                        round(sum(decode(a01.acc_code,'990000',amt,0)) /1,0) as field_990000 ");
		sql.append("                 from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
		sql.append("                 left join (select * from a01 ");
		sql.append("                            where m_year=? and m_month=? ");
		paramList.add(wlx01_m_year);
		paramList.add(syear);
		paramList.add(smonth);
		sql.append("                           ) a01  on  bn01.bank_no = a01.bank_code ");
		sql.append("                 group by a01.m_year,a01.m_month,bn01.bank_no,bn01.BANK_NAME ");
		sql.append("               ) a01,");
		sql.append("               ( select a05.m_year,a05.m_month, bn01.bank_no , bn01.BANK_NAME,");           
		sql.append("                        round(sum(decode(a05.acc_code,'910400',amt,0)) /1,0) as field_910400,");
		sql.append("                        round(sum(decode(a05.acc_code,'910500',amt,0)) /1,0) as field_910500 ");            
		sql.append("                 from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
		sql.append("                 left join (select * from a05 ");
		sql.append("                            where m_year=? and m_month=? ");  
		paramList.add(wlx01_m_year);
		paramList.add(syear);
		paramList.add(smonth);
		sql.append("                           ) a05  on  bn01.bank_no = a05.bank_code ");
		sql.append("                 group by a05.m_year,a05.m_month,bn01.bank_no,bn01.BANK_NAME ");
		sql.append("               ) a05, ");
		sql.append("               ( select a10.m_year,a10.m_month, bn01.bank_no , bn01.BANK_NAME,");
		sql.append("                        a10.build1_baddebt,a10.build2_baddebt,a10.build3_baddebt,a10.build4_baddebt ");
		sql.append("               from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
		sql.append("               left join (select * from a10 ");    
		sql.append("                        where m_year=? and m_month=? ");  
		paramList.add(wlx01_m_year);
		paramList.add(syear);
		paramList.add(smonth);
		sql.append("                       ) a10  on  bn01.bank_no = a10.bank_code ");
		sql.append("               ) a10 ");//105.07.05 add 
		sql.append("               where a99.bank_no=a02.bank_no(+) and a99.bank_no = a01.bank_no(+) ");
		sql.append("               and a99.bank_no = a05.bank_no(+) ");
		sql.append("               and a99.bank_no = a10.bank_no(+) ");//105.07.05 add
		sql.append("               and a99.bank_no <> ' ' "); 
		sql.append("               GROUP BY a99.m_year,a99.m_month,a99.bank_no,a99.BANK_NAME ");
		sql.append("        ) a99,");//--本月份資料
		sql.append("        (  ");
		sql.append("         select a99_10308.m_year,a99_10308.m_month, a99_10308.bank_no , a99_10308.BANK_NAME,");
		sql.append("                SUM(a99_10308.field_992710) as field_992710,");
		sql.append("                SUM(a01_10308.field_120800) + SUM(a01_10308.field_150300) as field_BACKUP ");
		sql.append("         from ( select a99.m_year,a99.m_month, bn01.bank_no , bn01.BANK_NAME,");           
		sql.append("                       round(sum(decode(a99.acc_code,'992710',amt,0)) /1,0) as field_992710 ");
		sql.append("               from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
		paramList.add(wlx01_m_year);
		sql.append("               left join (select * from a99 ");
		sql.append("                          where m_year=104 and m_month=6 ");//固定為104年6月份為控管前資料               
		sql.append("                         ) a99  on  bn01.bank_no = a99.bank_code ");
		sql.append("               group by a99.m_year,a99.m_month,bn01.bank_no,bn01.BANK_NAME ");
		sql.append("              ) a99_10308,");
		sql.append("              ( select a01.m_year,a01.m_month, bn01.bank_no , bn01.BANK_NAME,");           
		sql.append("                       round(sum(decode(a01.acc_code,'120800',amt,0)) /1,0) as field_120800,");
		sql.append("                       round(sum(decode(a01.acc_code,'150300',amt,0)) /1,0) as field_150300 ");
		sql.append("               from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
		paramList.add(wlx01_m_year);
		sql.append("               left join (select * from a01 "); 
		sql.append("                          where m_year=104 and m_month=6 ");//固定為104年6月份為控管前資料              
		sql.append("                          ) a01  on  bn01.bank_no = a01.bank_code ");
		sql.append("               group by a01.m_year,a01.m_month,bn01.bank_no,bn01.BANK_NAME ");
		sql.append("              ) a01_10308 ");
		sql.append("         where a99_10308.bank_no=a01_10308.bank_no(+) ");
		sql.append("         and a99_10308.bank_no <> ' ' "); 
		sql.append("         GROUP BY a99_10308.m_year,a99_10308.m_month,a99_10308.bank_no,a99_10308.BANK_NAME ");
		sql.append("        )a99_10308, ");//--控管前資料104年6月份
		sql.append("        ( select a99.m_year,a99.m_month, bn01.bank_no , bn01.BANK_NAME,");           
		sql.append("                 round(sum(decode(a99.acc_code,'992710',amt,0)) /1,0) as field_992710 ");
		sql.append("          from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
		sql.append("          left join (select * from a99 ");
		sql.append("                     where m_year=? and m_month=? ");//申報年月的上個雙數月 ex:若現在為104/02,則上一個雙數月為103/12                  
		paramList.add(wlx01_m_year);
		paramList.add(l2year);
		paramList.add(l2month);
		sql.append("                    ) a99  on  bn01.bank_no = a99.bank_code "); 
		sql.append("          group by a99.m_year,a99.m_month,bn01.bank_no,bn01.BANK_NAME ");
		sql.append("        ) a99_even,");//--上個雙數月
		sql.append("        ( select a99.m_year,a99.m_month, bn01.bank_no , bn01.BANK_NAME,");           
		sql.append("                 round(sum(decode(a99.acc_code,'992710',amt,0)) /1,0) as field_992710 ");
		sql.append("          from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
		sql.append("          left join (select * from a99 ");
		sql.append("                     where m_year=? and m_month=? "); 
		paramList.add(wlx01_m_year);
		paramList.add(lyear);
		paramList.add(lmonth);
		sql.append("                    ) a99  on  bn01.bank_no = a99.bank_code ");
		sql.append("          group by a99.m_year,a99.m_month,bn01.bank_no,bn01.BANK_NAME ");
		sql.append("         ) a99_last ");//--上月份
		sql.append("         where a99.bank_no=a99_10308.bank_no(+) ");  
		sql.append("         and a99.bank_no=a99_even.bank_no(+) ");
		sql.append("         and a99.bank_no=a99_last.bank_no(+) ");        
		sql.append("       )a99 ");
		sql.append("   )a99 ");
		sql.append(")a99 ");
		sql.append(" left join (select * from v_bank_location where m_year=?)v_bank_location on a99.bank_no=v_bank_location.bank_no ");
		paramList.add(wlx01_m_year);
		sql.append(" order by v_bank_location.FR001W_output_order,bank_no ");

	    List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"field_992710_10308,field_992710_last,field_992710,field_992711,field_992712,field_992713,field_992710_diff_10308,field_992710_diff_last,"+
																		"field_990230_990240,field_9902710_rate,field_9902710_credit_rate,field_9902710_220000_rate,"+
																		"field_over,field_over_rate,field_992720,field_992730,field_992720_rate,field_992910,build_baddebt_sum,"+
																		"field_baddebt_rate,field_backup,field_backup_10308,field_backup_diff,field_backup_over_rate,field_backup_credit_rate,field_captial_rate");
		System.out.println("dbData1.size()="+dbData.size());
		return dbData;
	}
	
}