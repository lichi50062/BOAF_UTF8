/*
 *102.01.08 create  by 2968
 */
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import java.io.*;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.*;

import com.tradevan.util.DownLoad;
import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RtpFR063W_Print {
 	
	 public static String createRpt(String s_year, String s_month, String bank_no, String bank_name, String unit) {    

	    String errMsg = "";
	    List dbData = null;
	    String sqlCmd = "";    
	    int rowNum=0;
	    DataObject bean = null;
	    reportUtil reportUtil = new reportUtil();
		HSSFCellStyle cs_right = null; 
		HSSFCellStyle cs_center = null;
		HSSFCellStyle cs_left = null;
		HSSFCellStyle nb_left = null;
		HSSFCellStyle nb_right = null;
	   
	  	StringBuffer sql = new StringBuffer() ;
	  	ArrayList paramList = new ArrayList() ;
	  	//String bank_type_name =( bank_type.equals("6") )?"農會":"漁會";
	  	/*String u_year = "100" ;
	  	if(s_year==null || Integer.parseInt(s_year)<=99 ) {
	  		u_year ="99" ;
	  	}*/
	    try {

	      File xlsDir = new File(Utility.getProperties("xlsDir"));
	      File reportDir = new File(Utility.getProperties("reportDir"));

	      if (!xlsDir.exists()) {
	        if (!Utility.mkdirs(Utility.getProperties("xlsDir"))) {
	          errMsg += Utility.getProperties("xlsDir") + "目錄新增失敗";
	        }
	      }
	      if (!reportDir.exists()) {
	        if (!Utility.mkdirs(Utility.getProperties("reportDir"))) {
	          errMsg += Utility.getProperties("reportDir") + "目錄新增失敗";
	        }
	      }
	      FileInputStream finput = null;
	      String filename="農漁會信用部聯合貸款案件明細表.xls";
	      //input the standard report form      
	      finput = new FileInputStream(xlsDir +System.getProperty("file.separator") +filename);
	      //設定FileINputStream讀取Excel檔
	      POIFSFileSystem fs = new POIFSFileSystem(finput);
	      HSSFWorkbook wb = new HSSFWorkbook(fs);
	      HSSFSheet sheet = wb.getSheetAt(0); //讀取第一個工作表，宣告其為sheet
	      HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
	      //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
	      //sheet.setAutobreaks(true); //自動分頁

	      //設定頁面符合列印大小
	      sheet.setAutobreaks(false);
	      ps.setScale( (short) 66); //列印縮放百分比

	      ps.setPaperSize( (short) 9); //設定紙張大小 A4
	      //wb.setSheetName(0,"test");
	      finput.close();

	      HSSFRow row = null; //宣告一列
	      HSSFCell cell = null; //宣告一個儲存格

	      short i = 0;
	      short y = 0;
	      cs_right = reportUtil.getRightStyle(wb);
	      cs_center = reportUtil.getDefaultStyle(wb);
	      cs_left = reportUtil.getLeftStyle(wb);
	      nb_left = reportUtil.getNoBorderLeftStyle(wb);
	      nb_right = reportUtil.getNoBoderStyle(wb);
	      String unit_name = Utility.getUnitName(unit); 
	   	  if(s_month.length()==1){
	   	      s_month="0"+s_month;
	   	  }
	                paramList =new ArrayList() ;
	                sqlCmd = "select case_no,seq_no,m_year,m_month," 
	                        +"      loan_idn," //--借款人統一編號
	                        +"      loan_name," //--借款人名稱
	                        +"      round(loan_amt_sum/?,0) as loan_amt_sum," //--授信案總金額
	                        +"      case_begin_year," //--授信案期間
	                        +"      case_begin_month,"
	                        +"      case_end_year,"
	                        +"      case_end_month,"
	                        +"      F_TRANSCODE('041',bank_no_max) as bank_no_max," //--主辦行
	                        +"      manabank_name," //--管理行
	                        +"      F_TRANSCODE('042',loan_kind) as loan_kind," //--參貸型式
	                        +"      round(loan_amt/?,0) as loan_amt,"//--參貸額度
	                        +"      round(loan_bal_amt/?,0) as loan_bal_amt," //--實際授信餘額
	                        +"      F_TRANSCODE('043',loan_type) as loan_type," //--信用部參貸部分之授信用途
	                        +"      F_TRANSCODE('044',pay_state) as pay_state," //--目前放款繳息情形
	                        +"      F_TRANSCODE('045',violate_type) as violate_type," //--有無違反契約承諾條款
	                        +"      loan_rate," //--目前放款利率
	                        +"      decode(new_case,'1','是','2','否','否') as new_case " //--是否本月新增案件
	                        +"from wlx10_m_loan "
	                        +"where bank_no=? "
	                        +"and to_char(m_year * 100 + m_month) = ? "
	                        +"union "
	                        +"select '',999,999,999,'合計','',"
	                        +"       null as loan_amt_sum," //--授信案總金額
	                        +"       null,null,null,null,'','','',"
	                        +"       sum(round(loan_amt/?,0)) as loan_amt," //--參貸額度
	                        +"       sum(round(loan_bal_amt/?,0)) as loan_bal_amt," //--實際授信餘額
	                        +"       '','','',null,'' "
	                        +"from wlx10_m_loan "
	                        +"where bank_no=? "
	                        +"and to_char(m_year * 100 + m_month) = ? "
	                        +"order by m_year,m_month,case_no,seq_no ";
	        paramList.add(unit) ;
	        paramList.add(unit) ;
	        paramList.add(unit) ;
	        paramList.add(bank_no) ;
	        paramList.add(s_year+s_month) ;
	        //paramList.add(unit) ;
            paramList.add(unit) ;
            paramList.add(unit) ;
	        paramList.add(bank_no) ;
	        paramList.add(s_year+s_month) ;
	        dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"case_no,seq_no,m_year,m_month,loan_idn,loan_name,loan_amt_sum," +
	                "case_begin_year,case_begin_month,case_end_year,case_end_month," +
	                "bank_no_max,manabank_name,loan_kind,loan_amt,loan_bal_amt,loan_type," +
	                "pay_state,violate_type,loan_rate,new_case"); 
	        System.out.println("dbData.size=" + dbData.size());
	     
	      //設定報表表頭資料============================================
	        row = sheet.getRow(2);                                                                        
            cell = row.getCell((short)0);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellValue(bank_name);
            row = sheet.getRow(3);                          
            cell = row.getCell((short)0);                  
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellValue(s_year+"年"+s_month+"月聯合貸款案件資料表"+((dbData == null || dbData.size() ==1)?"無資料存在":""));
            row = sheet.getRow(4);
            cell = row.getCell((short)0);                  
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            //cell.setCellStyle(cs_right);
            cell.setCellValue("單位:新臺幣" + unit_name+"、％"); 
            rowNum = 7;
            if (dbData != null) { 
                if(dbData.size()>1){
                    for (int k=0; k<dbData.size();k++){
        	            bean = (DataObject)dbData.get(k);
        	            String case_no=(bean.getValue("case_no") == null)?"":(bean.getValue("case_no")).toString();
        	            String loan_idn=(bean.getValue("loan_idn") == null)?"":(bean.getValue("loan_idn")).toString();
        	            String loan_name=(bean.getValue("loan_name") == null)?"":(bean.getValue("loan_name")).toString();
        	            String loan_amt_sum=(bean.getValue("loan_amt_sum") == null)?"":(bean.getValue("loan_amt_sum")).toString();
        	            String case_begin_year=(bean.getValue("case_begin_year") == null)?"":(bean.getValue("case_begin_year")).toString();
        	            String case_begin_month=(bean.getValue("case_begin_month") == null)?"":(bean.getValue("case_begin_month")).toString();
        	            String case_end_year=(bean.getValue("case_end_year") == null)?"":(bean.getValue("case_end_year")).toString();
        	            String case_end_month=(bean.getValue("case_end_month") == null)?"":(bean.getValue("case_end_month")).toString();
        	            String bank_no_max=(bean.getValue("bank_no_max") == null)?"":(bean.getValue("bank_no_max")).toString();
        	            String manabank_name=(bean.getValue("manabank_name") == null)?"":(bean.getValue("manabank_name")).toString();
        	            String loan_kind=(bean.getValue("loan_kind") == null)?"":(bean.getValue("loan_kind")).toString();
        	            String loan_type=(bean.getValue("loan_type") == null)?"":(bean.getValue("loan_type")).toString();
        	            String loan_amt=(bean.getValue("loan_amt") == null)?"":(bean.getValue("loan_amt")).toString();
        	            String loan_bal_amt=(bean.getValue("loan_bal_amt") == null)?"":(bean.getValue("loan_bal_amt")).toString();
        	            String pay_state=(bean.getValue("pay_state") == null)?"":(bean.getValue("pay_state")).toString();
        	            String violate_type=(bean.getValue("violate_type") == null)?"":(bean.getValue("violate_type")).toString();    
        	            String loan_rate=(bean.getValue("loan_rate") == null)?"":String.valueOf(Double.parseDouble((bean.getValue("loan_rate")).toString()));
        	            String new_case=(bean.getValue("new_case") == null)?"":(bean.getValue("new_case")).toString();
                        rowNum++;
                        row = sheet.getRow(rowNum);
                        System.out.println("rowNum= "+rowNum);
                        //申報編號
                        cell = row.getCell((short)0); 
                        cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                        cell.setCellStyle(cs_center);
                        cell.setCellValue(case_no);
        	            //借款人統一編號
                        cell = row.getCell((short)1); 
                        cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                        cell.setCellStyle(cs_center);
                        cell.setCellValue(loan_idn);
                        //借款人名稱
                         cell = row.getCell((short)2);
                        cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                        cell.setCellStyle(cs_center); 
                        cell.setCellValue(loan_name);
                        //授信案總金額
                        cell = row.getCell((short)3);
                        cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                        cell.setCellStyle(cs_right);
                        cell.setCellValue(Utility.setCommaFormat(loan_amt_sum));
                         //授信案期間
                        cell = row.getCell((short)4);
                        cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                        cell.setCellStyle(cs_right); 
                        cell.setCellValue(case_begin_year);
                        cell = row.getCell((short)5);
                        cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                        cell.setCellStyle(cs_right); 
                        cell.setCellValue(case_begin_month);
                        cell = row.getCell((short)6);
                        cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                        cell.setCellStyle(cs_right); 
                        cell.setCellValue(case_end_year);
                        cell = row.getCell((short)7);
                        cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                        cell.setCellStyle(cs_right); 
                        cell.setCellValue(case_end_month);
                        //主辦行
                        cell = row.getCell((short)8);
                        cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                        cell.setCellStyle(cs_center);
                        cell.setCellValue(bank_no_max);
                        //管理行
                        cell = row.getCell((short)9);
                        cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                        cell.setCellStyle(cs_center);
                        cell.setCellValue(manabank_name);
                        //參貸型式
                        cell = row.getCell((short)10);
                        cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                        cell.setCellStyle(cs_center);
                        cell.setCellValue(loan_kind);
                        //參貸額度
                        cell = row.getCell((short)11);
                        cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                        cell.setCellStyle(cs_right); 
                        cell.setCellValue(Utility.setCommaFormat(loan_amt));
                        //實際授信餘額
                        cell = row.getCell((short)12);
                        cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                        cell.setCellStyle(cs_right); 
                        cell.setCellValue(Utility.setCommaFormat(loan_bal_amt));
                        //信用部參貸部分之授信用途
                        cell = row.getCell((short)13);
                        cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                        cell.setCellStyle(cs_center);
                        cell.setCellValue(loan_type);
                        //目前放款繳息情形
                        cell = row.getCell((short)14);
                        cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                        cell.setCellStyle(cs_center);
                        cell.setCellValue(pay_state);
                        //有無違反契約承諾條款
                        cell = row.getCell((short)15);
                        cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                        cell.setCellStyle(cs_center);
                        cell.setCellValue(violate_type);
                        //目前放款利率
                        cell = row.getCell((short)16);
                        cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                        cell.setCellStyle(cs_right);
                        //cell.setCellValue("");
                        //建構時決定資料輸出格式  
                        //#字號當為小數後面為0時會自動去除  
                        DecimalFormat formatter = new DecimalFormat("#.###");  
                        //ex:輸出0.35  System.out.println(formatter.format(0.350));
                        //透過applyPattern改變格式  
                        //Pattern 裡0的用處為，當需要自動補0時可以遞補  
                        //ex:輸出1.30000  System.out.println(formatter.format(1.3));
                        formatter.applyPattern("0.0000");
                        cell.setCellValue((loan_rate=="")?"":formatter.format(Double.parseDouble(loan_rate)));
                        //是否本月新增案件
                        cell = row.getCell((short)17);
                        cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                        cell.setCellStyle(cs_center);
                        cell.setCellValue(new_case);
        	      	} // end of for
                } //end of if(dbData.size()>1)
	      } //end of if(dbData != null)
	      
	      
	      FileOutputStream fout = null;     
	      fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + filename);
	      System.out.println(reportDir + System.getProperty("file.separator") + filename);
	      HSSFFooter footer = sheet.getFooter();
	      footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
	      footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
	      wb.write(fout);
	      //儲存
	      fout.close();
	      System.out.println("儲存成功!");
	    }
	    catch (Exception e) {
	    	System.out.println("RptFR063W.createRpt Error:" + e + e.getMessage());
	    }
	    
	    return errMsg;
	  }
	 
	/*  ===============================================
     *  取得顯示的金額單位
     *  ===============================================
     */ 
    public static String getUnitName(String unit){
        
        String unit_name ="";
        
        //設定顯示的金額單位
        if(unit.equals("1")){
            unit_name="元";
        }else if (unit.equals("1000")){
            unit_name="千元";
        }else if (unit.equals("10000")){
            unit_name="萬元";
        }else if (unit.equals("1000000")){
            unit_name="百萬元";
        }else if (unit.equals("10000000")){
            unit_name="千萬元";
        }else if (unit.equals("100000000")){
            unit_name ="億元";
        }
        return unit_name;
    }//end of getUnitName
    
    /*  ===============================================
     *  補零
     *  ===============================================
     */ 
   /*public static String addZeroForNum(String str, int strLength) {
       int strLen = str.length();
       if (strLen < strLength) {
           while (strLen < strLength) {
               StringBuffer sb = new StringBuffer();
               //sb.append("0").append(str);//左補零
               sb.append(str).append("0");//右補零
               str = sb.toString();
               strLen = str.length();
           }
       }
       return str;
   }*/

 
}
