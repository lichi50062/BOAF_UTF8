/*
//105.11.09 add by 2968     
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

import javax.servlet.http.HttpSession;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptTM007W {
	public static String createRpt(String acc_tr_type,String lguser_name) {    
		String errMsg="";
		String acc_tr_name=getAaa_Tr_Name(acc_tr_type);
		String rptTitle = acc_tr_name;
		String filename="逾期未填報辦理情形農漁會明細表.xls";
		/*
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
*/

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
		    ps.setLandscape( false ); // 設定橫印:true
		    //ps.setPaperSize( ( short )8 ); //設定紙張大小 A3
		    ps.setPaperSize( (short) 9); //設定紙張大小 A4 (A3:8/A4:9)
			
			finput.close();
			
			
            HSSFCellStyle defaultStyle = reportUtil.getDefaultStyle(wb);//有框內文置中
            HSSFCellStyle noBoderStyle = reportUtil.getNoBoderStyle(wb);//無框置右	
            HSSFCellStyle leftStyle = reportUtil.getLeftStyle(wb);//有框內文置左
			HSSFRow row=null;//宣告一列
			HSSFCell cell=null;//宣告一個儲存格
	        
			
			//列印報表名稱
			row=(sheet.getRow(1)==null)? sheet.createRow(1) : sheet.getRow(1);
			cell=row.getCell((short)1);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue(rptTitle);//協助措施名稱
			row=(sheet.getRow(3)==null)? sheet.createRow(3) : sheet.getRow(3);
			cell=row.getCell((short)1);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			String today = Utility.getCHTYYMMDD("yy")+"年"+Utility.getCHTYYMMDD("mm")+"月"+Utility.getCHTYYMMDD("dd")+"日";
			cell.setCellValue("截至"+today);//截至○年○月○日
			row=(sheet.getRow(4)==null)? sheet.createRow(4) : sheet.getRow(4);
			cell=row.getCell((short)1);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellStyle(noBoderStyle);
			cell.setCellValue("列印日期："+getNow());
			row=(sheet.getRow(5)==null)? sheet.createRow(5) : sheet.getRow(5);
			cell=row.getCell((short)1);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellStyle(noBoderStyle);
			cell.setCellValue("列印人員："+lguser_name);
			
			
			List dataList = getDataList(acc_tr_type);
			int rowNum = 7;
			if(dataList!=null && dataList.size()>0){
				for(int i=0;i<dataList.size();i++){
					DataObject obj = (DataObject)dataList.get(i);
					String bank_code = obj.getValue("bank_code")==null?"":obj.getValue("bank_code").toString();
					String bank_name = obj.getValue("bank_name")==null?"":obj.getValue("bank_name").toString();
					String applydate = obj.getValue("applydate")==null?"":Utility.getCHTdate(obj.getValue("applydate").toString(), 0);
					String cnt_name = obj.getValue("cnt_name")==null?"":obj.getValue("cnt_name").toString();
					row=(sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
					System.out.println("date="+applydate);
					cell=row.getCell((short)1)==null?row.createCell((short)1):row.getCell((short)1);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(defaultStyle);
					cell.setCellValue(bank_code);//機構代號
					
					cell=row.getCell((short)2)==null?row.createCell((short)2):row.getCell((short)2);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(leftStyle);
					cell.setCellValue(bank_name);//機構名稱
					
					cell=row.getCell((short)3)==null?row.createCell((short)3):row.getCell((short)3);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(defaultStyle);
					cell.setCellValue(applydate);//申報基準日
					
					cell=row.getCell((short)4)==null?row.createCell((short)4):row.getCell((short)4);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(defaultStyle);
					cell.setCellValue(cnt_name);//聯絡人資料
					
					rowNum++;
						
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
	
	
	//報表查詢SQL
	public static List getDataList(String acc_tr_type){
		StringBuffer sql = new StringBuffer();
		List paramList = new ArrayList();//共同參數
		sql.append(" select a.bank_code,");//--機構代號
		sql.append(" bank_name,");//--機構名稱
		sql.append(" to_char(applydate,'YYYY/MM/DD') applydate,");//--申報基準日
		sql.append(" decode(cnt_name,null,wlx01.telno,cnt_name||' '||cnt_tel) as cnt_name  ");//--聯絡人資料
		sql.append(" from ( ");
		sql.append(" select loanapply_bn01.bank_code,");
		sql.append(" CASE WHEN loanapply_bn01.bank_code = loanapply_wml01.bank_code THEN 'V' ELSE '' END as wml01_ok,");
		sql.append(" loanapply_period.applydate,");//--基準日
		sql.append(" CASE WHEN loanapply_bn01.bank_code = loanapply_wml01.bank_code THEN loanapply_wml01.add_date ELSE null END as wml01_add_date ");//--申報日期
		sql.append(" from loanapply_period ");
		sql.append(" left join loanapply_wml01 on  loanapply_period.acc_tr_type=loanapply_wml01.acc_tr_type and  loanapply_period.applydate=loanapply_wml01.applydate ");
		sql.append(" left join loanapply_bn01 on loanapply_period.acc_tr_type=loanapply_bn01.acc_tr_type ");
		sql.append(" where loanapply_period.acc_tr_type=? ");
		sql.append(" and loanapply_period.applydate < sysdate ");
		sql.append(" )a ");
		sql.append(" left join (select bank_code,max(cnt_name) as cnt_name,max(cnt_tel) as cnt_tel from loanapply_wml01 where acc_tr_type=? group by bank_code)loanapply_wml_cnt on a.bank_code = loanapply_wml_cnt.bank_code ");
		sql.append(" left join wlx01 on a.bank_code=wlx01.bank_no ");
		sql.append(" left join (select * from bn01 where m_year=100)bn01 on a.bank_code = bn01.bank_no "); 
		sql.append(" where  wml01_add_date is null ");
		sql.append(" group by a.bank_code,bank_name,applydate,decode(cnt_name,null,wlx01.telno,cnt_name||' '||cnt_tel) ");
		sql.append(" order by a.bank_code,applydate ");
		paramList.add(acc_tr_type);
		paramList.add(acc_tr_type);
	    List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"bank_code,bank_name,applydate,cnt_name");
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