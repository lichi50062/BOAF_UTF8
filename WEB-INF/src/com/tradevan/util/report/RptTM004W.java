/*
105.11.07add by 2968       
*/
package com.tradevan.util.report;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.Region;

import java.io.*;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptTM004W {
	public static String createRpt(String acc_tr_type,String acc_div,String unit) {    
		String errMsg="";
		String unit_name=Utility.getUnitName(unit);
		String acc_tr_name=getAaa_Tr_Name(acc_tr_type);
		String amtTitle = "貸款餘額";
		String rptTitle = "("+acc_tr_name+")舊貸展延需求預估彙總表";
		if("02".equals(acc_div)){
			amtTitle = "貸款金額";
			rptTitle = "("+acc_tr_name+")新貸需求預估彙總表";
		}
		String filename="舊貸展延需求預估彙總表.xls";

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
            HSSFCellStyle ls = wb.createCellStyle();
            ft1.setFontHeightInPoints((short)12);
            ft1.setFontName("標楷體");
            ls.setFont(ft1);
            ls.setBorderLeft(HSSFCellStyle.BORDER_THIN);  
            ls.setBorderTop(HSSFCellStyle.BORDER_THIN);   
            ls.setBorderRight(HSSFCellStyle.BORDER_THIN); 
            ls.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            ls.setAlignment(HSSFCellStyle.ALIGN_LEFT);
            HSSFFont ft2 = wb.createFont();
            HSSFCellStyle rs = wb.createCellStyle();
            ft2.setFontHeightInPoints((short)12);
            ft2.setFontName("標楷體");
            rs.setFont(ft2);
            rs.setBorderLeft(HSSFCellStyle.BORDER_THIN);  
            rs.setBorderTop(HSSFCellStyle.BORDER_THIN);   
            rs.setBorderRight(HSSFCellStyle.BORDER_THIN); 
            rs.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            rs.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
            HSSFCellStyle ls_top = wb.createCellStyle();
            ls_top.setFont(ft1);
            ls_top.setBorderLeft(HSSFCellStyle.BORDER_NONE);  
            ls_top.setBorderTop(HSSFCellStyle.BORDER_THIN);   
            ls_top.setBorderRight(HSSFCellStyle.BORDER_NONE); 
            ls_top.setBorderBottom(HSSFCellStyle.BORDER_NONE);
            ls_top.setAlignment(HSSFCellStyle.ALIGN_LEFT);
            HSSFCellStyle ls_none = wb.createCellStyle();
            ls_none.setFont(ft1);
            ls_none.setBorderLeft(HSSFCellStyle.BORDER_NONE);  
            ls_none.setBorderTop(HSSFCellStyle.BORDER_NONE);   
            ls_none.setBorderRight(HSSFCellStyle.BORDER_NONE); 
            ls_none.setBorderBottom(HSSFCellStyle.BORDER_NONE);
            ls_none.setAlignment(HSSFCellStyle.ALIGN_LEFT);
            HSSFCellStyle cs_none = wb.createCellStyle();
            cs_none.setFont(ft1);
            cs_none.setBorderLeft(HSSFCellStyle.BORDER_NONE);  
            cs_none.setBorderTop(HSSFCellStyle.BORDER_NONE);   
            cs_none.setBorderRight(HSSFCellStyle.BORDER_NONE); 
            cs_none.setBorderBottom(HSSFCellStyle.BORDER_NONE);
            cs_none.setAlignment(HSSFCellStyle.ALIGN_CENTER);
            HSSFCellStyle cs6 = wb.createCellStyle();
            cs6.setFont(ft1);
            cs6.setBorderLeft(HSSFCellStyle.BORDER_THIN);  
            cs6.setBorderTop(HSSFCellStyle.BORDER_THIN);   
            cs6.setBorderRight(HSSFCellStyle.BORDER_THIN); 
            cs6.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            cs6.setAlignment(HSSFCellStyle.ALIGN_LEFT);
            cs6.setWrapText(true);//自動換行 
            
            
			HSSFRow row=null;//宣告一列
			HSSFCell cell=null;//宣告一個儲存格
	        
			//getSubItemList
			List subItemList=getSubItemList(acc_tr_type,acc_div);
			String subItems = ""; String itemNames = "";String periods="";
			if(subItemList!=null && subItemList.size()>0){
				for(int i=0;i<subItemList.size();i++){
					DataObject obj = (DataObject)subItemList.get(i);
					if(!"".equals(subItems)){
						subItems+=";";
						itemNames+=";";
						periods+=";";
					}
					subItems+=(String)obj.getValue("subitem");
					itemNames+=(String)obj.getValue("subitem_name");
					periods+=obj.getValue("rate_period").toString();
				}
			}
			row=(sheet.getRow(5)==null)? sheet.createRow(5) : sheet.getRow(5);
			cell=row.getCell((short)3);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellStyle(cs);
			cell.setCellValue(amtTitle);
			String[] nameStr = itemNames.split(";");
			int titleCel = 4;
			for (int i = 0; i < nameStr.length; i++){
				System.out.println("itemName="+nameStr[i]); 
				row=(sheet.getRow(4)==null)? sheet.createRow(4) : sheet.getRow(4);
				cell=row.getCell((short)titleCel)==null?row.createCell((short)titleCel):row.getCell((short)titleCel);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs);
				cell.setCellValue((String)nameStr[i]);
				row=(sheet.getRow(5)==null)? sheet.createRow(5) : sheet.getRow(5);
				cell=row.getCell((short)titleCel)==null?row.createCell((short)titleCel):row.getCell((short)titleCel);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs);
				cell.setCellValue("戶數");
				sheet.setColumnWidth((short)titleCel, (short)(19*256));  
				titleCel++;
				row=(sheet.getRow(4)==null)? sheet.createRow(4) : sheet.getRow(4);
				//合併儲存格，參數分別為起始行、起始列、結束行、結束列 
				sheet.addMergedRegion(new Region((short)4, (short) (titleCel-1), (short)4, (short)titleCel));
				cell=row.getCell((short)titleCel)==null?row.createCell((short)titleCel):row.getCell((short)titleCel);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs);
				cell.setCellValue("");
				row=(sheet.getRow(5)==null)? sheet.createRow(5) : sheet.getRow(5);
				cell=row.getCell((short)titleCel)==null?row.createCell((short)titleCel):row.getCell((short)titleCel);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs);
				cell.setCellValue(amtTitle);
				sheet.setColumnWidth((short)titleCel, (short)(19*256)); 
				if(i<nameStr.length-1)titleCel++;
			}
			
			//列印報表名稱
			row=(sheet.getRow(1)==null)? sheet.createRow(1) : sheet.getRow(1);
			for(int l=1;l<=titleCel;l++){
                cell = row.getCell((short)l)==null?row.createCell((short)l):row.getCell((short)l);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                cell.setCellValue("");
            }
			sheet.addMergedRegion(new Region((short)1, (short)1, (short)1, (short)titleCel));
			cell=row.getCell((short)1);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue(rptTitle);
			
			row=(sheet.getRow(3)==null)? sheet.createRow(3) : sheet.getRow(3);
			cell=row.getCell((short)1)==null? row.createCell((short)1) : row.getCell((short)1);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("單位：新台幣"+unit_name+",%");
			sheet.addMergedRegion(new Region((short)3, (short)(titleCel-1), (short)3, (short)titleCel));
			cell=row.getCell((short)(titleCel-1))==null? row.createCell((short)(titleCel-1)) : row.getCell((short)(titleCel-1));
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellStyle(ls_none);
			cell.setCellValue("報表產生日:"+getNow());
			
			
			
			
			List dataList = getDataList(acc_tr_type,acc_div,unit,subItems);
			int rowNum = 6;
			if(dataList!=null && dataList.size()>0){
				for(int i=0;i<dataList.size();i++){
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
					for(int l=0;l<=titleCel;l++){
		                cell = row.createCell( (short)l);
		            }
					rowNum++;
				}
				rowNum = 6;
				for(int i=0;i<dataList.size();i++){
					DataObject obj = (DataObject)dataList.get(i);
					//String bank_code = obj.getValue("bank_code")==null?"":(String)obj.getValue("bank_code");
					String bank_name = obj.getValue("bank_name")==null?"":(String)obj.getValue("bank_name");
					String field_seq = obj.getValue("field_seq")==null?"":(String)obj.getValue("field_seq");
					if("A99".equals(field_seq)){//總計
						bank_name = (String)obj.getValue("hsien_name");
					}else if("A90".equals(field_seq)){//縣市小計
						bank_name = (String)obj.getValue("hsien_name")+"小計";
					}
					String count_seq = obj.getValue("count_seq")==null?"":obj.getValue("count_seq").toString();
					String loan_cnt_sum = obj.getValue("loan_cnt_sum")==null?"":(String)obj.getValue("loan_cnt_sum");
					String loan_amt_sum = obj.getValue("loan_amt_sum")==null?"":obj.getValue("loan_amt_sum").toString();
					if("A99".equals(field_seq)){//因為sql撈出的第一筆就是總計，固直接寫到總計位置
						row=sheet.getRow(6+(dataList.size()-1));
					}else{
						row=sheet.getRow(rowNum);
					}
					cell=row.getCell((short)1);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(ls);
					cell.setCellValue(bank_name);
					cell=row.getCell((short)2);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(rs);
					cell.setCellValue(Utility.setCommaFormat(loan_cnt_sum));//合計-戶數
					cell=row.getCell((short)3);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(rs);
					cell.setCellValue(Utility.setCommaFormat(loan_amt_sum));//合計-貸款餘金額
					
					int celNum = 4;
					String[] itemStr = subItems.split(";"); 
					for (int s = 0; s < itemStr.length; s++){
						String tmpCnt = obj.getValue("loan_cnt_"+itemStr[s])==null?"":obj.getValue("loan_cnt_"+itemStr[s]).toString();
						String tmpAmt = obj.getValue("loan_amt_"+itemStr[s])==null?"":obj.getValue("loan_amt_"+itemStr[s]).toString();
						cell=row.getCell((short)celNum);
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell.setCellStyle(rs);
						cell.setCellValue(Utility.setCommaFormat(tmpCnt));//子貸款別-戶數
						celNum++;
						cell=row.getCell((short)celNum);
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell.setCellStyle(rs);
						cell.setCellValue(Utility.setCommaFormat(tmpAmt));//子貸款別-貸款餘金額//Utility.setCommaFormat
						celNum++;
					}
					if(!"A99".equals(field_seq)){
						rowNum++;
					}
						
				}
				rowNum++;//總計那行要加回rowNum
			}
			//預估貸款利率/利息
			row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
			for(int l=1;l<=titleCel;l++){
                cell = row.createCell( (short)l);
    			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
    			cell.setCellStyle(rs);
    			cell.setCellValue("");
            }
			cell=row.getCell((short)1);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellStyle(ls);
			cell.setCellValue("預估貸款利率/利息");
			//預估補貼利率/補貼息
			row=(sheet.getRow(rowNum+1)==null)? sheet.createRow(rowNum+1) : sheet.getRow(rowNum+1);
			for(int l=1;l<=titleCel;l++){
                cell = row.createCell( (short)l);
    			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
    			cell.setCellStyle(rs);
    			cell.setCellValue("");
            }
			cell=row.getCell((short)1);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellStyle(ls);
			cell.setCellValue("預估補貼利率/補貼息");
			//利息經費合計
			row=(sheet.getRow(rowNum+2)==null)? sheet.createRow(rowNum+2) : sheet.getRow(rowNum+2);
			for(int l=1;l<=titleCel;l++){
                cell = row.createCell( (short)l);
    			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
    			cell.setCellStyle(rs);
    			cell.setCellValue("");
            }
			cell=row.getCell((short)1);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellStyle(ls);
			cell.setCellValue("利息經費合計");
			
			if(!"".equals(subItems)){
				List rateList = getRateList(acc_tr_type,acc_div,unit,subItems);//計算預估利息、補貼利息
				int celNum = 4;
				if(rateList!=null && rateList.size()>0){
					int loan_amt_sum = 0;
					int base_amt_sum = 0;
					int interest_sum = 0;
					DataObject obj = (DataObject)rateList.get(0);
					String[] itemStr = subItems.split(";");
					for (int i = 0; i < itemStr.length; i++){
						int loan_amt = Integer.parseInt(obj.getValue("loan_amt_"+itemStr[i]).toString());
						int base_amt = Integer.parseInt(obj.getValue("base_amt_"+itemStr[i]).toString());
						int interest = Integer.parseInt(obj.getValue("interest_"+itemStr[i]).toString());
						//預估貸款利率/利息
						row=sheet.getRow(rowNum);
						cell=row.getCell((short)celNum);
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell.setCellStyle(rs);
						cell.setCellValue(obj.getValue("loan_rate_"+itemStr[i]).toString()+"%");
						cell=row.getCell((short)(celNum+1));
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell.setCellStyle(rs);
						cell.setCellValue(Utility.setCommaFormat(String.valueOf(loan_amt)));
						
						//預估補貼利率/補貼息
						row=sheet.getRow(rowNum+1);
						cell=row.getCell((short)celNum);
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell.setCellStyle(rs);
						cell.setCellValue(obj.getValue("base_loan_rate_"+itemStr[i]).toString()+"%");
						cell=row.getCell((short)(celNum+1));
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell.setCellStyle(rs);
						cell.setCellValue(Utility.setCommaFormat(String.valueOf(base_amt)));
						
						//利息經費合計
						row=sheet.getRow(rowNum+2);
						cell=row.getCell((short)(celNum+1));
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell.setCellStyle(rs);
						cell.setCellValue(Utility.setCommaFormat(String.valueOf(interest)));
						celNum = celNum+2;
						
						loan_amt_sum +=loan_amt;
						base_amt_sum +=base_amt;
						interest_sum +=interest;
						
					}
					row=sheet.getRow(rowNum);
					cell=row.getCell((short)3);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(rs);
					cell.setCellValue(Utility.setCommaFormat(String.valueOf(loan_amt_sum)));
					row=sheet.getRow(rowNum+1);
					cell=row.getCell((short)3);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(rs);
					cell.setCellValue(Utility.setCommaFormat(String.valueOf(base_amt_sum)));
					row=sheet.getRow(rowNum+2);
					cell=row.getCell((short)3);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(rs);
					cell.setCellValue(Utility.setCommaFormat(String.valueOf(interest_sum)));
					
				}
			}
			//註1
			row=(sheet.getRow(rowNum+3)==null)? sheet.createRow(rowNum+3) : sheet.getRow(rowNum+3);
			cell = row.createCell( (short)1);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			if(dataList!=null && dataList.size()>0){
				cell.setCellStyle(ls_top);
			}else{
				cell.setCellStyle(ls_none);
			}
			cell.setCellValue("註1：系統自動帶出各該貸款之利率，並自動計算利息金額【以免息期間半年為例：預估貸款利息=貸款餘額總計*貸款利率÷2】");
			//註2
			row=(sheet.getRow(rowNum+4)==null)? sheet.createRow(rowNum+4) : sheet.getRow(rowNum+4);
			cell = row.createCell( (short)1);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellStyle(ls_none);
			cell.setCellValue("註2：系統自動帶出各該貸款之補貼基準扣除貸款利率，並自動計算補貼息金額【以免息期間半年為例：預估補貼息=貸款餘額總計*(補貼基準-貸款利率)÷2】");
			
			
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
	
	//依所挑選的協助措施名稱的報表貸款名稱
	public static List getSubItemList(String acc_tr_type,String acc_div){
		StringBuffer sql = new StringBuffer();
		List paramList = new ArrayList();//共同參數
		sql.append("select subitem,");//--貸款子目別代碼
		sql.append("  	   subitem_name,");//--貸款名稱
		sql.append("  	   loan_rate,");//--預估貸款利率
		sql.append("  	   base_rate,");//--補貼基準
		sql.append("  	   (base_rate - loan_rate) as loan_base_rate,");//--預估補貼利率
		sql.append("  	   to_number(decode(rate_period,'0','0.5',rate_period)) as rate_period ");//--免息期間
		sql.append("  from loan_subitem,(select * from loan_ncacno where acc_tr_type=? and acc_div=? )loan_ncacno ");
		sql.append(" where loan_subitem.subitem = loan_ncacno.acc_code ");
		paramList.add(acc_tr_type);
		paramList.add(acc_div);
	    List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"loan_rate,base_rate,loan_base_rate,rate_period");
		System.out.println("getSubItemList.size()="+dbData.size());
		return dbData;
	}
	//依貸款子目別代碼組合下列SQL
	public static List getDataList(String acc_tr_type,String acc_div,String unit,String subItems){
		String[] itemStr = subItems.split(";"); 
		for (int i = 0; i < itemStr.length; i++){
			System.out.println("subItem="+itemStr[i]); 
		}
		StringBuffer sql = new StringBuffer();
		List paramList = new ArrayList();//共同參數
		sql.append("select loan_rpt.hsien_id ,"); 
		sql.append("       loan_rpt.hsien_name,"); 
		sql.append("       loan_rpt.FR001W_output_order,");
		sql.append("       loan_rpt.bank_code ,");//--機構代號 
		sql.append("       loan_rpt.BANK_NAME,");//--貸款經辦機構名稱
		sql.append("       loan_rpt.COUNT_SEQ,");
		sql.append("       loan_rpt.field_SEQ,"); 
		sql.append("       decode(loan_cnt_sum,null,'未申報',loan_cnt_sum) as loan_cnt_sum,");//--戶數.合計 105.09.06 fix 若該貸款機構未申報時，顯示未申報字串
		sql.append("       loan_amt_sum ");//--貨款金額或餘額.合計
		if(!"".equals(subItems)){
			for (int i = 0; i < itemStr.length; i++){
				sql.append("      ,loan_cnt_"+itemStr[i]);//--貸款子類別.戶數
				sql.append("      ,loan_amt_"+itemStr[i]);//--貸款子類別.金額  
			}
		}
		sql.append(" from ( "); 
		sql.append("      select loan_rpt.hsien_id , loan_rpt.hsien_name, loan_rpt.FR001W_output_order,");
		sql.append("             loan_rpt.bank_code , loan_rpt.BANK_NAME, COUNT_SEQ, field_SEQ,");      
		sql.append("             round(loan_cnt_sum/1,0) as loan_cnt_sum,round(loan_amt_sum/?,0) as loan_amt_sum ");
		paramList.add(unit);
		if(!"".equals(subItems)){
			for (int i = 0; i < itemStr.length; i++){
				sql.append("            ,round(loan_cnt_"+itemStr[i]+"/1,0) as loan_cnt_"+itemStr[i]+",round(loan_amt_"+itemStr[i]+"/?,0) as loan_amt_"+itemStr[i]);
				paramList.add(unit);
			}
		}
		sql.append("        from ");//--總計begin 
		sql.append("        ( ");
		sql.append("          select ' ' AS  hsien_id ,' 總   計 ' AS hsien_name, '001' AS FR001W_output_order,");                
		sql.append("                 ' ' AS  bank_code,' ' AS  BANK_NAME,COUNT(*) AS COUNT_SEQ,'A99' as  field_SEQ ");     
		String tmpStr = "";
		if(!"".equals(subItems)){
			for (int i = 0; i < itemStr.length; i++){
				tmpStr+="SUM(loan_cnt_"+itemStr[i]+")";
				if(i!=itemStr.length-1){
					tmpStr+="+";
				}
			}
		}else{
			tmpStr = " null ";
		}
		sql.append(","+tmpStr+" as loan_cnt_sum ");
		
		tmpStr = "";
		if(!"".equals(subItems)){
			for (int i = 0; i < itemStr.length; i++){
				tmpStr+="SUM(loan_amt_"+itemStr[i]+")";
				if(i!=itemStr.length-1){
					tmpStr+="+";
				}
			}
		}else{
			tmpStr = " null ";
		}
		sql.append(","+tmpStr+" as loan_amt_sum ");
		
		
		if(!"".equals(subItems)){
			for (int i = 0; i < itemStr.length; i++){
				sql.append("                ,SUM(loan_cnt_"+itemStr[i]+") loan_cnt_"+itemStr[i]+",SUM(loan_amt_"+itemStr[i]+") loan_amt_"+itemStr[i]);
			}
		}
		sql.append("          from ( ");
		sql.append("            select hsien_id,hsien_name,FR001W_output_order,bank_code,BANK_NAME ");
		if(!"".equals(subItems)){
			for (int i = 0; i < itemStr.length; i++){
				sql.append("                ,SUM(loan_cnt_"+itemStr[i]+") loan_cnt_"+itemStr[i]+",SUM(loan_amt_"+itemStr[i]+") loan_amt_"+itemStr[i]);
			}
		}
		sql.append("            from(  ");                            
		sql.append("              select nvl(cd01.hsien_id,' ')  as  hsien_id ,");               
		sql.append("                     nvl(cd01.hsien_name,'OTHER') as  hsien_name,");               
		sql.append("                     cd01.FR001W_output_order     as  FR001W_output_order,");               
		sql.append("                     loan_bn01.bank_code ,loan_bn01.BANK_NAME "); 
		if(!"".equals(subItems)){
			for (int i = 0; i < itemStr.length; i++){
				sql.append("                    ,round(sum(decode(acc_code,?,loan_cnt,'')) /1,0) as loan_cnt_"+itemStr[i]);//--貸款子類別.戶數
				sql.append("                    ,round(sum(decode(acc_code,?,loan_amt,'')) /1,0) as loan_amt_"+itemStr[i]);//--貸款子類別.金額           
				paramList.add(itemStr[i]);
				paramList.add(itemStr[i]);
			}
		}
		sql.append("              from  (select * from  cd01 where cd01.hsien_id <> 'Y'  ) cd01 ");
		sql.append("              left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = 100 and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL) "); 
		sql.append("              left join (select loan_bn01.*,bank_name from loan_bn01 left join bn01 on loan_bn01.bank_code=bn01.bank_no and bn01.m_year=100  where acc_tr_type=?)loan_bn01 on wlx01.bank_no=loan_bn01.bank_code and wlx01.m_year=100  ");
		paramList.add(acc_tr_type);
		sql.append("              left join (select * from loan_rpt where acc_tr_type=? and acc_div=?) loan_rpt  on  loan_bn01.bank_code = loan_rpt.bank_code ");
		paramList.add(acc_tr_type);
		paramList.add(acc_div);
		sql.append("              group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,loan_bn01.bank_code,loan_bn01.BANK_NAME,acc_code ");                   
		sql.append("            )group by hsien_id,hsien_name,FR001W_output_order,bank_code,BANK_NAME ");
		sql.append("          ) loan_rpt where loan_rpt.bank_code <> ' '  ");
		sql.append("        ) loan_rpt ");//--總計end   
		sql.append("        UNION ALL ");
		sql.append("        select loan_rpt.hsien_id , loan_rpt.hsien_name, loan_rpt.FR001W_output_order,");
		sql.append("               loan_rpt.bank_code , loan_rpt.BANK_NAME, COUNT_SEQ, field_SEQ,");
		sql.append("               round(loan_cnt_sum/1,0) as loan_cnt_sum,round(loan_amt_sum/?,0) as loan_amt_sum ");  
		paramList.add(unit);
		if(!"".equals(subItems)){
			for (int i = 0; i < itemStr.length; i++){
				sql.append("              ,round(loan_cnt_"+itemStr[i]+"/1,0) as loan_cnt_"+itemStr[i]+",round(loan_amt_"+itemStr[i]+"/?,0) as loan_amt_"+itemStr[i]);
				paramList.add(unit);
			}
		}
		sql.append("        from ");//--各別機構明細begin
		sql.append("        ( ");
		sql.append("          select loan_rpt.hsien_id ,loan_rpt.hsien_name,loan_rpt.FR001W_output_order,");                  
		sql.append("                 loan_rpt.bank_code , loan_rpt.BANK_NAME,count(*)  AS  COUNT_SEQ, 'A01'  as  field_SEQ "); 
		
		tmpStr = "";
		if(!"".equals(subItems)){
			for (int i = 0; i < itemStr.length; i++){
				tmpStr+="SUM(loan_cnt_"+itemStr[i]+")";
				if(i!=itemStr.length-1){
					tmpStr+="+";
				}
			}
		}else{
			tmpStr = " null ";
		}
		sql.append(","+tmpStr+" as loan_cnt_sum ");
		
		tmpStr = "";
		if(!"".equals(subItems)){
			for (int i = 0; i < itemStr.length; i++){
				tmpStr+="SUM(loan_amt_"+itemStr[i]+")";
				if(i!=itemStr.length-1){
					tmpStr+="+";
				}
			}
		}else{
			tmpStr = " null ";
		}
		sql.append(","+tmpStr+" as loan_amt_sum ");
		
		if(!"".equals(subItems)){
			for (int i = 0; i < itemStr.length; i++){
				sql.append("                ,SUM(loan_cnt_"+itemStr[i]+") loan_cnt_"+itemStr[i]+",SUM(loan_amt_"+itemStr[i]+") loan_amt_"+itemStr[i]); 
			}
		}
		sql.append("          from (  ");
		sql.append("            select hsien_id,hsien_name,FR001W_output_order,bank_code,BANK_NAME ");
		if(!"".equals(subItems)){
			for (int i = 0; i < itemStr.length; i++){
				sql.append("                ,SUM(loan_cnt_"+itemStr[i]+") loan_cnt_"+itemStr[i]+",SUM(loan_amt_"+itemStr[i]+") loan_amt_"+itemStr[i]); 
			}
		}
		sql.append("            from(  ");                         
		sql.append("              select nvl(cd01.hsien_id,' ')       as  hsien_id ,");               
		sql.append("                     nvl(cd01.hsien_name,'OTHER') as  hsien_name,");               
		sql.append("                     cd01.FR001W_output_order     as  FR001W_output_order,");               
		sql.append("                     loan_bn01.bank_code ,  loan_bn01.BANK_NAME "); 
		if(!"".equals(subItems)){
			for (int i = 0; i < itemStr.length; i++){
				sql.append("                    ,round(sum(decode(acc_code,?,loan_cnt,'')) /1,0) as loan_cnt_"+itemStr[i]);//--貸款子類別.戶數
				sql.append("                    ,round(sum(decode(acc_code,?,loan_amt,'')) /1,0) as loan_amt_"+itemStr[i]);//--貸款子類別.金額           
				paramList.add(itemStr[i]);
				paramList.add(itemStr[i]);
			}
		}
		sql.append("              from  (select * from  cd01 where cd01.hsien_id <> 'Y'  ) cd01 ");
		sql.append("              left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = 100 and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL) "); 
		sql.append("              left join (select loan_bn01.*,bank_name from loan_bn01 left join bn01 on loan_bn01.bank_code=bn01.bank_no and bn01.m_year=100  where acc_tr_type=?)loan_bn01 on wlx01.bank_no=loan_bn01.bank_code and wlx01.m_year=100 "); 
		paramList.add(acc_tr_type);
		sql.append("              left join (select * from loan_rpt where acc_tr_type=? and acc_div=?) loan_rpt  on  loan_bn01.bank_code = loan_rpt.bank_code ");
		paramList.add(acc_tr_type);
		paramList.add(acc_div);
		sql.append("              group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,loan_bn01.bank_code,loan_bn01.BANK_NAME,acc_code ");                   
		sql.append("            )group by hsien_id,hsien_name,FR001W_output_order,bank_code,BANK_NAME ");
		sql.append("          )loan_rpt where loan_rpt.bank_code <> ' ' "); 
		sql.append("          GROUP BY loan_rpt.hsien_id,loan_rpt.hsien_name,loan_rpt.FR001W_output_order,loan_rpt.bank_code,loan_rpt.BANK_NAME ");
		sql.append("         ) loan_rpt ");//--各別機構明細  
		sql.append("         UNION ALL ");
		sql.append("         select loan_rpt.hsien_id ,loan_rpt.hsien_name,loan_rpt.FR001W_output_order,");
		sql.append("                loan_rpt.bank_code ,loan_rpt.BANK_NAME,COUNT_SEQ, field_SEQ,");  
		sql.append("                round(loan_cnt_sum/1,0) as loan_cnt_sum,round(loan_amt_sum/?,0) as loan_amt_sum ");
		paramList.add(unit);
		if(!"".equals(subItems)){
			for (int i = 0; i < itemStr.length; i++){
				sql.append("              ,round(loan_cnt_"+itemStr[i]+"/1,0) as loan_cnt_"+itemStr[i]+",round(loan_amt_"+itemStr[i]+"/?,0) as loan_amt_"+itemStr[i]);
				paramList.add(unit);
			}
		}
		sql.append("         from ");
		sql.append("         ( ");//--縣市小計
		sql.append("           select loan_rpt.hsien_id , loan_rpt.hsien_name,  loan_rpt.FR001W_output_order,"); 
		sql.append("                  ' ' AS  bank_code , ' ' AS  BANK_NAME, count(*)  AS  COUNT_SEQ,'A90'  as  field_SEQ ");    
		
		tmpStr = "";
		if(!"".equals(subItems)){
			for (int i = 0; i < itemStr.length; i++){
				tmpStr+="SUM(loan_cnt_"+itemStr[i]+")";
				if(i!=itemStr.length-1){
					tmpStr+="+";
				}
			}
		}else{
			tmpStr = " null ";
		}
		sql.append(","+tmpStr+" as loan_cnt_sum ");
		
		tmpStr = "";
		if(!"".equals(subItems)){
			for (int i = 0; i < itemStr.length; i++){
				tmpStr+="SUM(loan_amt_"+itemStr[i]+")";
				if(i!=itemStr.length-1){
					tmpStr+="+";
				}
			}
		}else{
			tmpStr = " null ";
		}
		sql.append(","+tmpStr+" as loan_amt_sum ");
		
		if(!"".equals(subItems)){
			for (int i = 0; i < itemStr.length; i++){
				sql.append("                ,SUM(loan_cnt_"+itemStr[i]+") loan_cnt_"+itemStr[i]+",SUM(loan_amt_"+itemStr[i]+") loan_amt_"+itemStr[i]); 
			}
		}
		
		sql.append("           from (  ");      
		sql.append("             select hsien_id,hsien_name,FR001W_output_order,bank_code,BANK_NAME ");
		if(!"".equals(subItems)){
			for (int i = 0; i < itemStr.length; i++){
				sql.append("                ,SUM(loan_cnt_"+itemStr[i]+") loan_cnt_"+itemStr[i]+",SUM(loan_amt_"+itemStr[i]+") loan_amt_"+itemStr[i]);
			}
		}
		sql.append("             from( ");                      
		sql.append("               select nvl(cd01.hsien_id,' ')       as  hsien_id ,");               
		sql.append("                      nvl(cd01.hsien_name,'OTHER') as  hsien_name,");               
		sql.append("                      cd01.FR001W_output_order     as  FR001W_output_order,");               
		sql.append("                      loan_bn01.bank_code ,  loan_bn01.BANK_NAME "); 
		if(!"".equals(subItems)){
			for (int i = 0; i < itemStr.length; i++){
				sql.append("                    ,round(sum(decode(acc_code,?,loan_cnt,'')) /1,0) as loan_cnt_"+itemStr[i]);//--貸款子類別.戶數
				sql.append("                    ,round(sum(decode(acc_code,?,loan_amt,'')) /1,0) as loan_amt_"+itemStr[i]);//--貸款子類別.金額           
				paramList.add(itemStr[i]);
				paramList.add(itemStr[i]);
			}
		}
		sql.append("               from  (select * from  cd01 where cd01.hsien_id <> 'Y'  ) cd01 ");
		sql.append("               left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = 100 and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL) "); 
		sql.append("               left join (select loan_bn01.*,bank_name from loan_bn01 left join bn01 on loan_bn01.bank_code=bn01.bank_no and bn01.m_year=100  where acc_tr_type=?)loan_bn01 on wlx01.bank_no=loan_bn01.bank_code and wlx01.m_year=100 "); 
		paramList.add(acc_tr_type);
		sql.append("               left join (select * from loan_rpt where acc_tr_type=? and acc_div=?) loan_rpt  on  loan_bn01.bank_code = loan_rpt.bank_code ");
		paramList.add(acc_tr_type);
		paramList.add(acc_div);
		sql.append("               group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,loan_bn01.bank_code,loan_bn01.BANK_NAME,acc_code ");            
		sql.append("             )group by hsien_id,hsien_name,FR001W_output_order,bank_code,BANK_NAME ");
		sql.append("           )loan_rpt where loan_rpt.bank_code <> ' ' "); 
		sql.append("           GROUP BY loan_rpt.hsien_id ,loan_rpt.hsien_name,loan_rpt.FR001W_output_order ");
		sql.append("          ) loan_rpt ");//--縣市小計end
		sql.append("  )  loan_rpt ");
		sql.append("  left join (select * from  cd01 where cd01.hsien_id <> 'Y') cd01 on loan_rpt.hsien_id = cd01.hsien_id ");
		sql.append("  ORDER by FR001W_output_order,field_SEQ,hsien_id,bank_code "); 
		
		String tStr = "count_seq,loan_amt_sum";
		if(!"".equals(subItems)){
			for (int i = 0; i < itemStr.length; i++){
				tStr +=(",loan_cnt_"+itemStr[i]+",loan_amt_"+itemStr[i]);
			}
		}
		System.out.println("****getDataList tStr="+tStr);
	    List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,tStr);
		System.out.println("getDataList.size()="+dbData.size());
		return dbData;
	}
		//計算預估利息、補貼利息
		public static List getRateList(String acc_tr_type,String acc_div,String unit,String subItems){
			String[] itemStr = subItems.split(";"); 
			for (int i = 0; i < itemStr.length; i++){
				System.out.println("subItem="+itemStr[i]); 
			}
			StringBuffer sql = new StringBuffer();
			List paramList = new ArrayList();//共同參數
			sql.append("select ");
			for (int i = 0; i < itemStr.length; i++){
				sql.append("       loan_rate_"+itemStr[i]+",");//--貸款子類別.預估貸款利率
				sql.append("       round(( (loan_amt_"+itemStr[i]+" * loan_rate_"+itemStr[i]+"/100)*rate_period_"+itemStr[i]+")/ ?,0) as loan_amt_"+itemStr[i]+",");//--貸款子類別.預估利息
				sql.append("       base_rate_"+itemStr[i]+",base_rate_"+itemStr[i]+"-loan_rate_"+itemStr[i]+" as base_loan_rate_"+itemStr[i]+",");//--貸款子類別.預估補貼利率
				sql.append("       round(( (loan_amt_"+itemStr[i]+" * (base_rate_"+itemStr[i]+"-loan_rate_"+itemStr[i]+")/100*rate_period_"+itemStr[i]+"))/ ?,0) as base_amt_"+itemStr[i]+",");//--貸款子類別.預估補貼息
				sql.append("       round(( (loan_amt_"+itemStr[i]+" * loan_rate_"+itemStr[i]+"/100)*rate_period_"+itemStr[i]+")/ ?,0) + ");
				sql.append("       round(( (loan_amt_"+itemStr[i]+" * (base_rate_"+itemStr[i]+"-loan_rate_"+itemStr[i]+")/100*rate_period_"+itemStr[i]+"))/ ?,0)  as interest_"+itemStr[i]);//--貸款子類別.利息經費合計
				paramList.add(unit);
				paramList.add(unit);
				paramList.add(unit);
				paramList.add(unit);
				if(i!=itemStr.length-1){
					sql.append(",");
				}
			}
			sql.append(" from ");
			sql.append(" ( ");
			sql.append("   Select "); 
			for (int i = 0; i < itemStr.length; i++){
				sql.append("           sum(loan_rate_"+itemStr[i]+") as loan_rate_"+itemStr[i]+", sum(base_rate_"+itemStr[i]+") as base_rate_"+itemStr[i]+","); 
				sql.append("           sum(rate_period_"+itemStr[i]+") as rate_period_"+itemStr[i]+",round(sum(loan_amt_"+itemStr[i]+")/1,0) as loan_amt_"+itemStr[i]);
				if(i!=itemStr.length-1){
					sql.append(",");
				}
			}
			sql.append("   from ");//--總計begin 
			sql.append("   ( ");
			sql.append("          select ' ' AS  hsien_id ,' 總   計 ' AS hsien_name, '001' AS FR001W_output_order,");                
			sql.append("                 ' ' AS  bank_code,' ' AS  BANK_NAME,COUNT(*) AS COUNT_SEQ,'A99' as  field_SEQ,");  
			for (int i = 0; i < itemStr.length; i++){
				sql.append("                 0 as loan_rate_"+itemStr[i]+",0 as base_rate_"+itemStr[i]+",0 as rate_period_"+itemStr[i]+",SUM(loan_amt_"+itemStr[i]+") loan_amt_"+itemStr[i]);
				if(i!=itemStr.length-1){
					sql.append(",");
				}
			}
			sql.append("          from ( ");
			sql.append("            select hsien_id,hsien_name,FR001W_output_order,bank_code,BANK_NAME,");
			for (int i = 0; i < itemStr.length; i++){
				sql.append("                   SUM(loan_amt_"+itemStr[i]+") loan_amt_"+itemStr[i]);
				if(i!=itemStr.length-1){
					sql.append(",");
				}
			}
			sql.append("            from( ");                            
			sql.append("              select nvl(cd01.hsien_id,' ')  as  hsien_id ,");               
			sql.append("                     nvl(cd01.hsien_name,'OTHER') as  hsien_name,");               
			sql.append("                     cd01.FR001W_output_order     as  FR001W_output_order,");               
			sql.append("                     loan_bn01.bank_code ,loan_bn01.BANK_NAME,"); 
			for (int i = 0; i < itemStr.length; i++){
				sql.append("                     round(sum(decode(acc_code,?,loan_amt,'')) /1,0) as loan_amt_"+itemStr[i]);//--貸款子類別.金額 
				paramList.add(itemStr[i]);
				if(i!=itemStr.length-1){
					sql.append(",");
				}
			}
			sql.append("              from  (select * from  cd01 where cd01.hsien_id <>  'Y' ) cd01 ");
			sql.append("              left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = 100 and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL) "); 
			sql.append("              left join (select loan_bn01.*,bank_name from loan_bn01 left join bn01 on loan_bn01.bank_code=bn01.bank_no and bn01.m_year=100  where acc_tr_type=?)loan_bn01 on wlx01.bank_no=loan_bn01.bank_code and wlx01.m_year=100 "); 
			sql.append("              left join (select * from loan_rpt where acc_tr_type=? and acc_div=?) loan_rpt  on  loan_bn01.bank_code = loan_rpt.bank_code ");
			paramList.add(acc_tr_type);
			paramList.add(acc_tr_type);
			paramList.add(acc_div);
			sql.append("              group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,");
			sql.append("                       loan_bn01.bank_code,loan_bn01.BANK_NAME,acc_code ");                  
			sql.append("            )group by hsien_id,hsien_name,FR001W_output_order,bank_code,BANK_NAME ");
			sql.append("          ) loan_rpt where loan_rpt.bank_code <> ' '  ");
			sql.append("          union all ");
			sql.append("          select '','','','','',0,'', ");
			for (int i = 0; i < itemStr.length; i++){
				sql.append("              decode(subitem,?,loan_rate) as loan_rate_"+itemStr[i]+",");
				sql.append("              decode(subitem,?,base_rate) as base_rate_"+itemStr[i]+",");
				sql.append("              decode(loan_ncacno.acc_code,?,to_number(decode(rate_period,'0','0.5',rate_period)),0) as rate_period_"+itemStr[i]+",0");
				paramList.add(itemStr[i]);
				paramList.add(itemStr[i]);
				paramList.add(itemStr[i]);
				if(i!=itemStr.length-1){
					sql.append(",");
				}
			}
			sql.append("          from loan_subitem,(select * from loan_ncacno where acc_tr_type=? ");
			sql.append("               and acc_div=? )loan_ncacno ");
			paramList.add(acc_tr_type);
			paramList.add(acc_div);
			sql.append("          where loan_subitem.subitem = loan_ncacno.acc_code  ");   
			sql.append("        ) loan_rpt ");//--總計end;              
			sql.append("     )a ");
			String tStr = "";
			for (int i = 0; i < itemStr.length; i++){
				tStr += ("loan_rate_"+itemStr[i]+",loan_amt_"+itemStr[i]+",base_loan_rate_"+itemStr[i]+",base_amt_"+itemStr[i]+",interest_"+itemStr[i]);
				if(i!=itemStr.length-1){
					tStr +=",";
				}
			}
			System.out.println("****getRateList  tStr="+tStr);
		    List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, tStr);
			System.out.println("getRateList.size()="+dbData.size());
			return dbData;
		}	
	public static String getAaa_Tr_Name(String acc_tr_type){
		String rtnVal = "";
        StringBuffer sqlCmd = new StringBuffer();
        List paramList = new ArrayList();//傳內的參數List   
        sqlCmd.append(" select distinct acc_tr_name from loan_ncacno ");
        sqlCmd.append("  where acc_tr_type=? "); 
        paramList.add(acc_tr_type); 
        
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
        if(dbData!=null && dbData.size()>0){
        	DataObject bean = (DataObject)dbData.get(0);
        	rtnVal = (String)bean.getValue("acc_tr_name");
        }
        return  rtnVal; 
    }
	public static String getNow() {
		Calendar rightNow = Calendar.getInstance();
		String year = formatNumber(String.valueOf(new Integer(rightNow.get(Calendar.YEAR)) - 1911 ), 3);
		String month = formatNumber((new Integer(rightNow.get(Calendar.MONTH) + 1)).toString(), 2);
		String day = formatNumber(String.valueOf(rightNow.get(Calendar.DAY_OF_MONTH)), 2);
		String hour = (new Integer(rightNow.get(Calendar.HOUR_OF_DAY))).toString();
		String minute = (new Integer(rightNow.get(Calendar.MINUTE))).toString();
		String second = (new Integer(rightNow.get(Calendar.SECOND))).toString();
		if (hour.length() == 1) hour = "0" + hour;
		if (minute.length() == 1) minute = "0" + minute;
		if (second.length() == 1) second = "0" + second;
		return (year+"年"+month+"月"+day+"日 "+hour+":"+minute+":"+second);
	}
		/**
		 * 將傳入的數字格式化 aDigits 位數的字串，不滿位數則補零
		 * 
		 * @param aNumber
		 * @param aDigits
		 *            長度
		 * @return
		 */
		public static String formatNumber(String num, int digits) {
			StringBuffer sbFmt = new StringBuffer();
			NumberFormat formatter;

			for (int i = 0; i < digits; i++) {
				sbFmt.append("0");
			}
			formatter = new DecimalFormat(sbFmt.toString());
			return formatter.format(Integer.parseInt(num));
		}
}