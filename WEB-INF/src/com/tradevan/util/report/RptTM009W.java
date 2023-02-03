/*
105.11.09 add by 2968       
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

public class RptTM009W {
	public static String createRpt(String acc_tr_type,String acc_div,String unit,String applydate) {    
		String errMsg="";
		String unit_name=Utility.getUnitName(unit);
		String acc_tr_name=getAaa_Tr_Name(acc_tr_type);
		String rptTitle = acc_tr_name+"案件總表";
		
		String filename="協助措施案件明細表_舊貸.xls";
		if("02".equals(acc_div)){
			filename="協助措施案件明細表_新貸.xls";
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
		    ps.setPaperSize( (short) 9); //設定紙張大小 A4 (A3:8/A4:9)
			
			finput.close();
			HSSFCellStyle leftStyle = wb.createCellStyle(); 
			HSSFCellStyle ls = HssfStyle.setStyle( leftStyle, wb.createFont(),
                    new String[] {
                    "BORDER", "PHL", "PVC", "F10",
                    "WRAP"} );//有框內文置左
			 
            HSSFCellStyle cs = reportUtil.getDefaultStyle(wb);//有框內文置中
            HSSFCellStyle rs = reportUtil.getRightStyle(wb);//有框內文置右
            
			HSSFRow row=null;//宣告一列
			HSSFCell cell=null;//宣告一個儲存格
			String applyTypeName = "";
			String applydate_Period = "";
			String sumperiod = "";
			List dataList = getDataList(acc_tr_type,acc_div,unit,applydate);
			int rowNum = 5;
			
			if(dataList!=null && dataList.size()>0){
				String lBank = "";
				int bank_start = 5;
				for(int i=0;i<dataList.size();i++){
					int celNum = 0;
					DataObject obj = (DataObject)dataList.get(i);
					String bank_code = obj.getValue("bank_code")==null?"":obj.getValue("bank_code").toString();
					String bank_name = obj.getValue("bank_name")==null?"":obj.getValue("bank_name").toString();
					if(i==0){
						String applytype = obj.getValue("applytype")==null?"2":obj.getValue("applytype").toString();
						if("2".equals(applytype)){//申報頻率 2:雙週報 1:週報 4:月報
							applyTypeName="本2週";
						}else if("1".equals(applytype)){
							applyTypeName="本週";
						}else if("4".equals(applytype)){
							applyTypeName="本月";
						}
						sumperiod = obj.getValue("sumperiod")==null?"":obj.getValue("sumperiod").toString();
						applydate_Period = obj.getValue("applydate_period")==null?"":obj.getValue("applydate_period").toString();
					}
					String acc_name = obj.getValue("acc_name")==null?"":obj.getValue("acc_name").toString();
					String apply_cnt = obj.getValue("apply_cnt")==null?"0":obj.getValue("apply_cnt").toString();
					String apply_amt = obj.getValue("apply_amt")==null?"0":Utility.setCommaFormat(obj.getValue("apply_amt").toString());
					String apply_bal = obj.getValue("apply_bal")==null?"0":Utility.setCommaFormat(obj.getValue("apply_bal").toString());
					String apply_cnt_sum = obj.getValue("apply_cnt_sum")==null?"0":Utility.setCommaFormat(obj.getValue("apply_cnt_sum").toString());
					String apply_amt_sum = obj.getValue("apply_amt_sum")==null?"0":Utility.setCommaFormat(obj.getValue("apply_amt_sum").toString());
					String apply_bal_sum = obj.getValue("apply_bal_sum")==null?"0":Utility.setCommaFormat(obj.getValue("apply_bal_sum").toString());
					String appr_cnt = obj.getValue("appr_cnt")==null?"0":Utility.setCommaFormat(obj.getValue("appr_cnt").toString());
					String appr_amt = obj.getValue("appr_amt")==null?"0":Utility.setCommaFormat(obj.getValue("appr_amt").toString());
					String appr_bal = obj.getValue("appr_bal")==null?"0":Utility.setCommaFormat(obj.getValue("appr_bal").toString());
					String appr_cnt_sum = obj.getValue("appr_cnt_sum")==null?"0":Utility.setCommaFormat(obj.getValue("appr_cnt_sum").toString());
					String appr_amt_sum = obj.getValue("appr_amt_sum")==null?"0":Utility.setCommaFormat(obj.getValue("appr_amt_sum").toString());
					String appr_bal_sum = obj.getValue("appr_bal_sum")==null?"0":Utility.setCommaFormat(obj.getValue("appr_bal_sum").toString());
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
					if("01".equals(acc_div)){
						for(int l=0;l<14;l++){
			                cell = row.getCell((short)l)==null?row.createCell((short)l):row.getCell((short)l);
			                cell=row.getCell((short)l);
							cell.setEncoding(HSSFCell.ENCODING_UTF_16);
							cell.setCellValue("");
			            }
					}else{
						for(int l=0;l<12;l++){
			                cell = row.getCell((short)l)==null?row.createCell((short)l):row.getCell((short)l);
			                cell=row.getCell((short)l);
							cell.setEncoding(HSSFCell.ENCODING_UTF_16);
							cell.setCellValue("");
			            }
					}
					if("9999999".equals(bank_code)){
						cell=row.getCell((short)celNum);
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell.setCellStyle(cs);
						cell.setCellValue(bank_name);
						celNum++;
						cell=row.getCell((short)celNum);
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell.setCellStyle(ls);
						cell.setCellValue("");
						celNum++;
						sheet.addMergedRegion(new Region((short)rowNum, (short)0, (short)rowNum, (short)1));
					}else{
						cell=row.getCell((short)celNum);
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell.setCellStyle(ls);
						cell.setCellValue(bank_code+bank_name);
						if(lBank.equals(bank_code)){
							sheet.addMergedRegion(new Region((short)bank_start, (short)celNum, (short)rowNum, (short)celNum));
						}else{
							bank_start =rowNum;
						}
						celNum++;
						cell=row.getCell((short)celNum);
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell.setCellStyle(cs);
						cell.setCellValue(acc_name);
						celNum++;
					}
					
					//本週申請-件數
					row=sheet.getRow(2);
					cell=row.getCell((short)celNum);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellValue(applyTypeName+"申請");
					row=sheet.getRow(rowNum);
					cell=row.getCell((short)celNum);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(rs);
					cell.setCellValue(apply_cnt);
					celNum++;
					
					//本週申請-貸款金額
					row=sheet.getRow(rowNum);
					cell=row.getCell((short)celNum);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(rs);
					cell.setCellValue(apply_amt);
					celNum++;
					
					//本週申請-貸款餘額
					if("01".equals(acc_div)){
						row=sheet.getRow(rowNum);
						cell=row.getCell((short)celNum);
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell.setCellStyle(rs);
						cell.setCellValue(apply_bal);
						celNum++;
					}
					
					//申請累計-件數
					row=sheet.getRow(3);
					cell=row.getCell((short)celNum);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellValue(sumperiod);
					row=sheet.getRow(rowNum);
					cell=row.getCell((short)celNum);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(rs);
					cell.setCellValue(apply_cnt_sum);
					celNum++;
					
					//申請累計-貸款金額
					row=sheet.getRow(rowNum);
					cell=row.getCell((short)celNum);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(rs);
					cell.setCellValue(apply_amt_sum);
					celNum++;
					
					//申請累計-貸款餘額
					if("01".equals(acc_div)){
						row=sheet.getRow(rowNum);
						cell=row.getCell((short)celNum);
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell.setCellStyle(rs);
						cell.setCellValue(apply_bal_sum);
						celNum++;
					}
					
					//本週核准-件數
					row=sheet.getRow(2);
					cell=row.getCell((short)celNum);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellValue(applyTypeName+"核准");
					row=sheet.getRow(rowNum);
					cell=row.getCell((short)celNum);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(rs);
					cell.setCellValue(appr_cnt);
					celNum++;
					
					//本週核准-貸款金額
					row=sheet.getRow(rowNum);
					cell=row.getCell((short)celNum);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(rs);
					cell.setCellValue(appr_amt);
					celNum++;
					
					//本週核准-貸款餘額
					row=sheet.getRow(rowNum);
					cell=row.getCell((short)celNum);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(rs);
					cell.setCellValue(appr_bal);
					celNum++;
					
					//核准累計-件數
					row=sheet.getRow(3);
					cell=row.getCell((short)celNum);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellValue(sumperiod);
					row=sheet.getRow(rowNum);
					cell=row.getCell((short)celNum);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(rs);
					cell.setCellValue(appr_cnt_sum);
					celNum++;
					
					//核准累計-貸款金額
					row=sheet.getRow(rowNum);
					cell=row.getCell((short)celNum);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(rs);
					cell.setCellValue(appr_amt_sum);
					celNum++;
					
					//核准累計-貸款餘額
					row=sheet.getRow(rowNum);
					cell=row.getCell((short)celNum);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(rs);
					cell.setCellValue(appr_bal_sum);
					
					rowNum++;
					lBank = bank_code;	
				}
			}
			
			//列印報表名稱
			row=(sheet.getRow(0)==null)? sheet.createRow(0) : sheet.getRow(0);
			cell=row.getCell((short)0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue(rptTitle);
			row=(sheet.getRow(1)==null)? sheet.createRow(1) : sheet.getRow(1);
			cell=row.getCell((short)0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue(applydate_Period);//上一次申報基準日至所選取的申報基準日-1
			if("02".equals(acc_div)){
				cell=row.getCell((short)9);
			}else{
				cell=row.getCell((short)11);
			}
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("單位：新台幣"+unit_name);
			
			
			
			
			HSSFFooter footer=sheet.getFooter();
			footer.setCenter( "Page:"+HSSFFooter.page()+" of "+HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
	
			FileOutputStream fout=new FileOutputStream(reportDir+System.getProperty("file.separator")+"協助措施案件明細表.xls");
			wb.write(fout);
			//儲存
			fout.close();
			System.out.println("儲存完成");
		}catch(Exception e) {
			System.out.println("createRpt Error:"+e+e.getMessage());
		}
		return errMsg;
	}
	
	
	//依貸款子目別代碼組合下列SQL
	public static List getDataList(String acc_tr_type,String acc_div,String unit,String applydate){
		StringBuffer sql = new StringBuffer();
		List paramList = new ArrayList();//共同參數
		sql.append(" select bank_code,");
		sql.append(" 		bank_name,");//--貸款經辦機構名稱
		sql.append(" 		applytype,");//--申報頻率
		sql.append(" 		'(自'||F_TRANSCHINESEDATE(begindate)||'起至'|| F_TRANSCHINESEDATE(loanapply_rpt.applydate)||')'  as sumperiod  ,");//--累計期間
		sql.append(" 		F_TRANSCHINESEDATE(applydate_b) || '至' || F_TRANSCHINESEDATE(applydate_e) applydate_period,");//--申報期間
		sql.append(" 		loanapply_rpt.acc_code,");
		sql.append(" 		acc_name,");//--貸款種類
		sql.append(" 		apply_cnt,");//--申請件數
		sql.append(" 		round(apply_amt/?,0) as apply_amt,");//--申請.貸款金額
		sql.append(" 		round(apply_bal/?,0) as apply_bal,");//--申請.貸款餘額
		sql.append(" 		apply_cnt_sum,");//--申請累計件數
		sql.append(" 		round(apply_amt_sum/?,0) as apply_amt_sum,");//--申請累計.貸款金額
		sql.append(" 		round(apply_bal_sum/?,0) as apply_bal_sum,");// --申請累計.貸款餘額
		sql.append(" 		appr_cnt,");//--核准件數
		sql.append(" 		round(appr_amt/?,0) as appr_amt,");//--核准.貸款金額
		sql.append(" 		round(appr_bal/?,0) as appr_bal,");//--核准.貸款餘額
		sql.append(" 		appr_cnt_sum,");//--核准累計件數
		sql.append(" 		round(appr_amt_sum/?,0) as appr_amt_sum,");//--核准累計.貸款金額
		sql.append(" 		round(appr_bal_sum/?,0) as appr_bal_sum ");//--核准累計.貸款餘額
		for(int i=0;i<8;i++){
			paramList.add(unit);
		}
		sql.append("   from loanapply_rpt ");
		sql.append("   left join loanapply_ncacno ");
		sql.append("     on loanapply_rpt.acc_tr_type=loanapply_ncacno.acc_tr_type and loanapply_rpt.acc_div=loanapply_ncacno.acc_div and loanapply_rpt.acc_code=loanapply_ncacno.acc_code ");
		sql.append("   left join loanapply_period on loanapply_rpt.acc_tr_type = loanapply_period.acc_tr_type and loanapply_rpt.applydate = loanapply_period.applydate ");
		sql.append("   left join (select * from bn01 where m_year=100)bn01 on loanapply_rpt.bank_code=bn01.bank_no ");
		sql.append("  where loanapply_rpt.acc_tr_type=? ");
		sql.append("    and loanapply_rpt.acc_div=? ");
		sql.append("    and loanapply_rpt.applydate = TO_DATE(?, 'YYYY/MM/DD') ");
		sql.append("    and (apply_cnt != 0 or apply_amt != 0 ) ");
		paramList.add(acc_tr_type);
		paramList.add(acc_div);
		paramList.add(applydate);
		sql.append(" union ");
		sql.append(" select "); 
		sql.append(" 		'9999999' as bank_code,");
		sql.append(" 		'合 計' as bank_name,");//--貸款經辦機構名稱
		sql.append(" 		'',null,'','','', ");
		sql.append(" 		sum(apply_cnt) as apply_cnt,");//--申請件數
		sql.append(" 		round(sum(apply_amt)/?,0) as apply_amt,");//--申請.貸款金額
		sql.append(" 		round(sum(apply_bal)/?,0) as apply_bal,");//--申請.貸款餘額
		sql.append(" 		sum(apply_cnt_sum) as apply_cnt_sum,");//--申請累計件數
		sql.append(" 		round(sum(apply_amt_sum)/?,0) as apply_amt_sum,");//--申請累計.貸款金額
		sql.append(" 		round(sum(apply_bal_sum)/?,0) as apply_bal_sum,");//--申請累計.貸款餘額
		sql.append(" 		sum(appr_cnt) as appr_cnt,");//--核准件數
		sql.append(" 		round(sum(appr_amt)/?,0) as appr_amt,");//--核准.貸款金額
		sql.append(" 		round(sum(appr_bal)/?,0) as appr_bal,");//--核准.貸款餘額
		sql.append(" 		sum(appr_cnt_sum) as appr_cnt,");//--核准累計件數
		sql.append(" 		round(sum(appr_amt_sum)/?,0) as appr_amt_sum,");//--核准累計.貸款金額
		sql.append(" 		round(sum(appr_bal_sum)/?,0) as appr_bal_sum ");//--核准累計.貸款餘額
		for(int i=0;i<8;i++){
			paramList.add(unit);
		}
		sql.append("   from loanapply_rpt ");
		sql.append("   left join loanapply_ncacno ");
		sql.append("     on loanapply_rpt.acc_tr_type=loanapply_ncacno.acc_tr_type and loanapply_rpt.acc_div=loanapply_ncacno.acc_div and loanapply_rpt.acc_code=loanapply_ncacno.acc_code ");
		sql.append("   left join loanapply_period on loanapply_rpt.acc_tr_type = loanapply_period.acc_tr_type and loanapply_rpt.applydate = loanapply_period.applydate ");
		sql.append("   left join (select * from bn01 where m_year=100)bn01 on loanapply_rpt.bank_code=bn01.bank_no  ");
		sql.append("  where loanapply_rpt.acc_tr_type=? ");
		sql.append("    and loanapply_rpt.acc_div=? "); 
		sql.append("    and loanapply_rpt.applydate = TO_DATE(?, 'YYYY/MM/DD') ");
		sql.append("    and (apply_cnt != 0 or apply_amt != 0 ) ");
		sql.append("  order by bank_code ");
		paramList.add(acc_tr_type);
		paramList.add(acc_div);
		paramList.add(applydate);
			
		List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "apply_cnt,apply_amt,apply_bal,apply_cnt_sum,apply_amt_sum,apply_bal_sum"
		                                                                  +",appr_cnt,appr_amt,appr_bal,appr_cnt_sum,appr_amt_sum,appr_bal_sum");
		System.out.println("getDataList.size()="+dbData.size());
		return dbData;
	}	
	public static String getAaa_Tr_Name(String acc_tr_type){
		String rtnVal = "";
        StringBuffer sqlCmd = new StringBuffer();
        List paramList = new ArrayList();//傳內的參數List   
        sqlCmd.append(" select distinct acc_tr_name from loanapply_ncacno ");
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