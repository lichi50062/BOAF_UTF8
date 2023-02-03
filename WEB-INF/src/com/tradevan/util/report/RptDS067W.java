/*
 * 105.11.17 add by 2968
 */
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.Region;

import java.io.*;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptDS067W {
	 public static String createRpt(String ACC_TR_TYPE,String Unit,String ACC_DIV,String APPLYDATE,String selectAccCode,String selectBank_no,int selBankCnt,String hasBankListALL) {    
	    String errMsg = "";	    
	    short rowNum=0;
	    short celNum=0;
	    DataObject bean = null;
	    reportUtil reportUtil = new reportUtil();
		HSSFCellStyle cs_right = null; 
		HSSFCellStyle cs_center = null;
	    HSSFCellStyle cs_left = null;
	    HSSFCellStyle cs_noborderleft = null;	      
	    String unit_name=Utility.getUnitName(Unit);
		String acc_tr_name=getAcc_Tr_Name(ACC_TR_TYPE);
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
	      System.out.println("ACC_TR_TYPE="+ACC_TR_TYPE);
	      System.out.println("ACC_DIV="+ACC_DIV);
	      System.out.println("APPLYDATE="+APPLYDATE);
	      System.out.println("Unit="+Unit);
	      System.out.println("selectAccCode="+selectAccCode);
	      System.out.println("selectBank_no="+selectBank_no);
	      System.out.println("selBankCnt="+selBankCnt);
	      System.out.println("hasBankListALL="+hasBankListALL);
	      
	      FileInputStream finput = null;
	      String fileName = "";
	      if("02".equals(ACC_DIV)){
              fileName = "DS067W_個別_新貸.xls";
          }else{
        	  fileName = "DS067W_個別_舊貸.xls";
          }   
	      
	      finput = new FileInputStream(xlsDir +System.getProperty("file.separator") +fileName);      
	      
	      //設定FileINputStream讀取Excel檔
          POIFSFileSystem fs = new POIFSFileSystem(finput);
          HSSFWorkbook wb = new HSSFWorkbook(fs);
          HSSFSheet sheet = wb.getSheetAt(0);//讀取第一個工作表，宣告其為sheet
          HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
         
               
	      //設定頁面符合列印大小
	      sheet.setAutobreaks(false);
	      ps.setLandscape( true ); // 設定橫印
	      ps.setScale( (short) (ACC_DIV.equals("01")?66:82)); //列印縮放百分比
	      ps.setPaperSize( (short) 9); //設定紙張大小 A4
	      
	      finput.close();

	      HSSFRow row = null; //宣告一列
	      HSSFCell cell = null; //宣告一個儲存格	      
	      cs_right = reportUtil.getRightStyle(wb);
	      cs_center = reportUtil.getDefaultStyle(wb);
	      cs_left = reportUtil.getLeftStyle(wb);	      
	      cs_noborderleft = reportUtil.getNoBorderLeftStyle(wb); 
	      
	      List dataList = new ArrayList();
	      if("true".equals(hasBankListALL)){//若選取全部，則依縣市別統計資料
	    	  dataList = getDataList3(ACC_TR_TYPE,ACC_DIV,Unit,APPLYDATE,selectAccCode,selectBank_no);
	      }else{
	    	  if(selBankCnt>1){//若選取多家貸款經辦機構，則統計所取選的貸款經辦機構合計資料
	    		  dataList = getDataList2(ACC_TR_TYPE,ACC_DIV,Unit,APPLYDATE,selectAccCode,selectBank_no);
	    	  }else{//若選取單一家貸款經辦機構，則統計單一家貸款經辦機構明細申報資料
	    		  dataList = getDataList1(ACC_TR_TYPE,ACC_DIV,Unit,APPLYDATE,selectAccCode,selectBank_no);
	    	  }
	      }
	      String applydate_period="";String tmpTitle="本週";String bank_code  ="";String bank_name ="";
	      String title_bank="";String cnt_name ="";String cnt_tel ="";
	      if(dataList!=null && dataList.size()>0){
	    	  rowNum=6;
	    	  for(int i=0;i<dataList.size();i++){
        		  bean = (DataObject)dataList.get(i);
        		  
        		  String apply_cnt = (bean.getValue("apply_cnt") == null)?"":Utility.setCommaFormat((bean.getValue("apply_cnt")).toString());
        		  String apply_amt = (bean.getValue("apply_amt") == null)?"":Utility.setCommaFormat((bean.getValue("apply_amt")).toString());
        		  String apply_bal = (bean.getValue("apply_bal") == null)?"":Utility.setCommaFormat((bean.getValue("apply_bal")).toString());
        		  String apply_cnt_sum = (bean.getValue("apply_cnt_sum") == null)?"":Utility.setCommaFormat((bean.getValue("apply_cnt_sum")).toString());
        		  String apply_amt_sum = (bean.getValue("apply_amt_sum") == null)?"":Utility.setCommaFormat((bean.getValue("apply_amt_sum")).toString());
        		  String apply_bal_sum = (bean.getValue("apply_bal_sum") == null)?"":Utility.setCommaFormat((bean.getValue("apply_bal_sum")).toString());
        		  String appr_cnt = (bean.getValue("appr_cnt") == null)?"":Utility.setCommaFormat((bean.getValue("appr_cnt")).toString());
        		  String appr_amt = (bean.getValue("appr_amt") == null)?"":Utility.setCommaFormat((bean.getValue("appr_amt")).toString());
        		  String appr_bal = (bean.getValue("appr_bal") == null)?"":Utility.setCommaFormat((bean.getValue("appr_bal")).toString());
        		  String appr_cnt_sum = (bean.getValue("appr_cnt_sum") == null)?"":Utility.setCommaFormat((bean.getValue("appr_cnt_sum")).toString());
        		  String appr_amt_sum = (bean.getValue("appr_amt_sum") == null)?"":Utility.setCommaFormat((bean.getValue("appr_amt_sum")).toString());
        		  String appr_bal_sum = (bean.getValue("appr_bal_sum") == null)?"":Utility.setCommaFormat((bean.getValue("appr_bal_sum")).toString());
        		  String nonappr_cnt = (bean.getValue("nonappr_cnt") == null)?"":Utility.setCommaFormat((bean.getValue("nonappr_cnt")).toString());
        		  String applytype="";String sumperiod="";String acc_code="";String acc_name="";
        		  String hsien_id="";String hsien_name ="";
        		  String fr001w_output_order ="";String count_seq ="";String field_seq ="";
        		  String nonappr_reason ="";
        		  if("true".equals(hasBankListALL)){//縣市別統計資料
        			  hsien_id = (bean.getValue("hsien_id") == null)?"":(bean.getValue("hsien_id")).toString();
        			  acc_name = (bean.getValue("hsien_name") == null)?"":(bean.getValue("hsien_name")).toString();
        			  fr001w_output_order = (bean.getValue("fr001w_output_order") == null)?"":(bean.getValue("fr001w_output_order")).toString();
        			  count_seq = (bean.getValue("count_seq") == null)?"":(bean.getValue("count_seq")).toString();
        			  field_seq = (bean.getValue("field_seq") == null)?"":(bean.getValue("field_seq")).toString();
        			  if(i==0){
        				  applydate_period = getApplydate_period(ACC_TR_TYPE,APPLYDATE);
        			  }
        			  List reasonList = getNonappr_reason3(ACC_TR_TYPE,ACC_DIV,Unit,APPLYDATE,hsien_id);
        			  for(int r=0;r<reasonList.size();r++){
        				  DataObject rBean = (DataObject)reasonList.get(r);
        				  if(!"".equals(nonappr_reason)){
        					  nonappr_reason += "\n";
        				  }
        				  nonappr_reason += String.valueOf(rBean.getValue("nonappr_reason"));
        			  }
        			  
        		  }else{
        			  if(i==0){
	        			  applytype  = (bean.getValue("applytype") == null)?"":(bean.getValue("applytype")).toString();
	        			  if("2".equals(applytype))tmpTitle="本2週";
	            		  if("4".equals(applytype))tmpTitle="本月";
	        			  sumperiod  = (bean.getValue("sumperiod") == null)?"":(bean.getValue("sumperiod")).toString();
	            		  applydate_period = (bean.getValue("applydate_period") == null)?"":(bean.getValue("applydate_period")).toString();
        			  }
            		  acc_code = (bean.getValue("acc_code") == null)?"":(bean.getValue("acc_code")).toString();
            		  acc_name = (bean.getValue("acc_name") == null)?"":(bean.getValue("acc_name")).toString();
        			  if(selBankCnt>1){
    	        		  List reasonList = getNonappr_reason2(ACC_TR_TYPE,ACC_DIV,Unit,APPLYDATE,acc_code,selectBank_no);
    	        		  for(int r=0;r<reasonList.size();r++){
            				  DataObject rBean = (DataObject)reasonList.get(r);
            				  if(!"".equals(nonappr_reason)){
            					  nonappr_reason += "\n";
            				  }
            				  nonappr_reason += String.valueOf(rBean.getValue("nonappr_reason"));
            			  }
    	        	  }else{
    	        		  hsien_name = (bean.getValue("hsien_name") == null)?"":(bean.getValue("hsien_name")).toString();
    	        		  bank_code  = (bean.getValue("bank_code") == null)?"":(bean.getValue("bank_code")).toString();
    	        		  bank_name  = (bean.getValue("bank_name") == null)?"":(bean.getValue("bank_name")).toString();
    	        		  nonappr_reason = (bean.getValue("nonappr_reason") == null)?"":(bean.getValue("nonappr_reason")).toString();
    	        		  if(i==0){
	    	        		  cnt_name = (bean.getValue("cnt_name") == null)?"":(bean.getValue("cnt_name")).toString();
	    	        		  cnt_tel  = (bean.getValue("cnt_tel") == null)?"":(bean.getValue("cnt_tel")).toString();
	    	        		  title_bank = hsien_name+bank_name.replace("信用部", "");
    	        		  }
    	        	  }
        			  
        		  }
        		  
	        	  celNum = 0;
	        	  setCelVal(sheet,row,cell,cs_left,rowNum,celNum,acc_name);//貸款種類or縣市別
	        	  
	        	  celNum++;
	        	  if(i==0){
        			  row = sheet.getRow((short)3);
        	 		  cell = row.getCell(celNum);
        	          cell.setEncoding(HSSFCell.ENCODING_UTF_16);        
        	          cell.setCellValue(tmpTitle+"申請");
        		  }
        		  setCelVal(sheet,row,cell,cs_right,rowNum,celNum,apply_cnt);//本週申請-件數
        		  celNum++;
        		  setCelVal(sheet,row,cell,cs_right,rowNum,celNum,apply_amt);//本週申請-貸款金額
        		  if("01".equals(ACC_DIV)){
        			  celNum++;
	        		  setCelVal(sheet,row,cell,cs_right,rowNum,celNum,apply_bal);//本週申請-貸款餘額
        		  }
        		  
        		  celNum++;
        		  if(i==0){
        			  row = sheet.getRow((short)4);
        	 		  cell = row.getCell(celNum);
        	          cell.setEncoding(HSSFCell.ENCODING_UTF_16);        
        	          cell.setCellValue(sumperiod);//(自發佈日期起至所選取的申報基準日)
        		  }
        		  setCelVal(sheet,row,cell,cs_right,rowNum,celNum,apply_cnt_sum);//申請累計-件數
        		  celNum++;
        		  setCelVal(sheet,row,cell,cs_right,rowNum,celNum,apply_amt_sum);//申請累計-貸款金額
        		  if("01".equals(ACC_DIV)){
        			  celNum++;
	        		  setCelVal(sheet,row,cell,cs_right,rowNum,celNum,apply_bal_sum);//申請累計-貸款餘額
        		  }
        		  
        		  celNum++;
        		  if(i==0){
        			  row = sheet.getRow((short)3);
        	 		  cell = row.getCell(celNum);
        	          cell.setEncoding(HSSFCell.ENCODING_UTF_16);        
        	          cell.setCellValue(tmpTitle+"核准");
        		  }
        		  setCelVal(sheet,row,cell,cs_right,rowNum,celNum,appr_cnt);//本週核准-件數
        		  
        		  celNum++;
        		  setCelVal(sheet,row,cell,cs_right,rowNum,celNum,appr_amt);//本週核准-貸款金額
        		  if("01".equals(ACC_DIV)){
        			  celNum++;
	        		  setCelVal(sheet,row,cell,cs_right,rowNum,celNum,appr_bal);//本週核准-貸款餘額
        		  }
        		  celNum++;
        		  if(i==0){
        			  row = sheet.getRow((short)4);
        	 		  cell = row.getCell(celNum);
        	          cell.setEncoding(HSSFCell.ENCODING_UTF_16);        
        	          cell.setCellValue(sumperiod);//(自發佈日期起至所選取的申報基準日)
        		  }
        		  setCelVal(sheet,row,cell,cs_right,rowNum,celNum,appr_cnt_sum);//核准累計-件數
        		  
        		  celNum++;
        		  setCelVal(sheet,row,cell,cs_right,rowNum,celNum,appr_amt_sum);//核准累計-貸款金額
        		  celNum++;
        		  if(i==dataList.size()-1){//跑到最後一筆時
        			  if(!"true".equals(hasBankListALL) && selBankCnt==1){//選取單一家貸款經辦機構時,才顯示填報人資訊
        				  setCelVal(sheet,row,cell,cs_noborderleft,(short)(rowNum+1),celNum,"填報人資訊：");
        				  setCelVal(sheet,row,cell,cs_noborderleft,(short)(rowNum+2),celNum,"  姓名："+cnt_name);
        				  setCelVal(sheet,row,cell,cs_noborderleft,(short)(rowNum+2),(short)(celNum+1),"");
        				  setCelVal(sheet,row,cell,cs_noborderleft,(short)(rowNum+2),(short)(celNum+2),"");
        				  setCelVal(sheet,row,cell,cs_noborderleft,(short)(rowNum+3),celNum,"  電話："+cnt_tel);
        				  setCelVal(sheet,row,cell,cs_noborderleft,(short)(rowNum+3),(short)(celNum+1),"");
        				  setCelVal(sheet,row,cell,cs_noborderleft,(short)(rowNum+3),(short)(celNum+2),"");
        				  sheet.addMergedRegion(new Region((short)(rowNum+2), (short)celNum, (short)(rowNum+2), (short)(celNum+2)));
        				  sheet.addMergedRegion(new Region((short)(rowNum+3), (short)celNum, (short)(rowNum+3), (short)(celNum+2)));
        	          }
        		  }
        		  setCelVal(sheet,row,cell,cs_right,rowNum,celNum,appr_bal_sum);//核准累計-貸款餘額
        		  
        		  celNum++;
        		  if(i==0){
        			  row = sheet.getRow((short)3);
        	 		  cell = row.getCell(celNum);
        	          cell.setEncoding(HSSFCell.ENCODING_UTF_16);        
        	          cell.setCellValue(tmpTitle+"不予核貸案件");
        		  }
        		  setCelVal(sheet,row,cell,cs_right,rowNum,celNum,nonappr_cnt);//本週不予核貸案件-件數
        		  celNum++;
        		  setCelVal(sheet,row,cell,cs_left,rowNum,celNum,nonappr_reason);//本週不予核貸案件-原因
        		  
        		  rowNum++;
        		  
        		  
        	  }
	    	  
	    	  //設定報表表頭資料============================================
	    	  //(協助措施名稱)辦理情形統計表
	    	  //1.若為單一家貸款經辦機構,則顯示所選取的金融協助措施名稱+辦理情形統計表
	    	  //2.若選取多家貸款經辦機構或[全部]時,則顯示所選取的金融協助措施名稱+辦理情形彙總表
	    	  String titleStr=acc_tr_name+"辦理情形統計表";
	    	  if("true".equals(hasBankListALL) || selBankCnt>1){
	    		  titleStr=acc_tr_name+"辦理情形彙總表";
	    	  }
	    	  row = sheet.getRow((short)0);
	 		  cell = row.getCell((short)0);
	          cell.setEncoding(HSSFCell.ENCODING_UTF_16);        
	          cell.setCellValue(titleStr);
	    	  //______縣(市)________農(漁)會
	    	  row = sheet.getRow((short)1);
	 		  cell = row.getCell((short)0);
	          cell.setEncoding(HSSFCell.ENCODING_UTF_16); 
	          cell.setCellValue(title_bank);
	          //上一次申報基準日至所選取的申報基準日-1
	    	  row = sheet.getRow((short)1);
	 		  cell = row.getCell((short)1);
	          cell.setEncoding(HSSFCell.ENCODING_UTF_16);        
	          cell.setCellValue(applydate_period);
	          
	          //單位：新台幣千元
	    	  row = sheet.getRow((short)1);
	    	  if("02".equals(ACC_DIV)){
	    		  cell = row.getCell((short)10);
	    	  }else{
	    		  cell = row.getCell((short)13); 
	    	  }
	 		  
	          cell.setEncoding(HSSFCell.ENCODING_UTF_16);        
	          cell.setCellValue("單位：新台幣"+unit_name);
	          if("true".equals(hasBankListALL)){//貸款種類/縣市別
	        	  row = sheet.getRow((short)3);
		 		  cell = row.getCell((short)0);
		          cell.setEncoding(HSSFCell.ENCODING_UTF_16);        
		          cell.setCellValue("縣市別");
	          }
	          
          }
	      
	      
	     
	      FileOutputStream fout = null;     
	      fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + "DS067W.xls");
	     
	      HSSFFooter footer = sheet.getFooter();
	      footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
	      footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
	      wb.write(fout);
	      //儲存
	      fout.close();
	      System.out.println("儲存成功!");
	      
	    }catch (Exception e) {
	      System.out.println("RptDS067W.createRpt Error:" + e + e.getMessage());
	    }
	    
	    return errMsg;
	  }
      
	 public static void setCelVal(HSSFSheet sheet,HSSFRow row,HSSFCell cell,HSSFCellStyle style,short rowNum,short celNum,String val){
		 row = (sheet.getRow(rowNum)==null)? sheet.createRow(rowNum) : sheet.getRow(rowNum);
		 cell = (row.getCell(celNum)==null)? row.createCell(celNum) : row.getCell(celNum);
         cell.setEncoding(HSSFCell.ENCODING_UTF_16);
         cell.setCellStyle(style);
         cell.setCellValue(val);
	 }
	 
	 //單一家機構明細資料.彈性報表查詢
	 public static List getDataList1(String acc_tr_type,String acc_div,String unit,String applydate,String selectAccCode,String selectBank_no){
			StringBuffer sql = new StringBuffer();
			List paramList = new ArrayList();//共同參數
			sql.append(" select wlx01.hsien_name,");//--縣市名稱
			sql.append(" loanapply_rpt.bank_code,"); 
			sql.append(" bank_name,");//--貸款經辦機構名稱
			sql.append(" applytype,");//--申報頻率 
			sql.append(" '(自'||F_TRANSCHINESEDATE(begindate)||'起至'|| F_TRANSCHINESEDATE(loanapply_rpt.applydate)||')'  as sumperiod  ,");//--累計期間
			sql.append(" F_TRANSCHINESEDATE(applydate_b) || '至' || F_TRANSCHINESEDATE(applydate_e) applydate_period,");//--申報期間
			sql.append(" loanapply_rpt.acc_code,");
			sql.append(" acc_name,");//--貸款種類
			sql.append(" apply_cnt,");//--申請件數
			sql.append(" round(apply_amt/?,0) as apply_amt,");//--申請.貸款金額
			sql.append(" round(apply_bal/?,0) as apply_bal,");//--申請.貸款餘額
			sql.append(" apply_cnt_sum,");//--申請累計件數
			sql.append(" round(apply_amt_sum/?,0) as apply_amt_sum,");//--申請累計.貸款金額
			sql.append(" round(apply_bal_sum/?,0) as apply_bal_sum,");//--申請累計.貸款餘額
			sql.append(" appr_cnt,");//--核准件數
			sql.append(" round(appr_amt/?,0) as appr_amt,");//--核准.貸款金額
			sql.append(" round(appr_bal/?,0) as appr_bal,");//--核准.貸款餘額
			sql.append(" appr_cnt_sum,");//--核准累計件數
			sql.append(" round(appr_amt_sum/?,0) as appr_amt_sum,");//--核准累計.貸款金額
			sql.append(" round(appr_bal_sum/?,0) as appr_bal_sum,");//--核准累計.貸款餘額
			for(int i=0;i<8;i++){
				paramList.add(unit);
			}
			sql.append(" nonappr_cnt,");//--不予核貸件數
			sql.append(" nonappr_reason,");//--不予核貸原因
			sql.append(" cnt_name,");//--填報人資訊.姓名
			sql.append(" cnt_tel ");//--填報人資訊.電話
			sql.append(" from loanapply_rpt ");
			sql.append(" left join loanapply_ncacno "); 
			sql.append(" on loanapply_rpt.acc_tr_type=loanapply_ncacno.acc_tr_type and loanapply_rpt.acc_div=loanapply_ncacno.acc_div and loanapply_rpt.acc_code=loanapply_ncacno.acc_code ");
			sql.append(" left join loanapply_period on loanapply_rpt.acc_tr_type = loanapply_period.acc_tr_type and loanapply_rpt.applydate = loanapply_period.applydate ");
			sql.append(" left join (select * from bn01 where m_year=100)bn01 on loanapply_rpt.bank_code=bn01.bank_no ");
			sql.append(" left join loanapply_wml01 on loanapply_rpt.acc_tr_type = loanapply_wml01.acc_tr_type and loanapply_rpt.bank_code=loanapply_wml01.bank_code and loanapply_rpt.applydate=loanapply_wml01.applydate ");
			sql.append(" left join (select wlx01.*,hsien_name from wlx01 left join cd01 on wlx01.hsien_id=cd01.hsien_id where m_year=100)wlx01 on loanapply_rpt.bank_code=wlx01.bank_no ");
			sql.append(" where loanapply_rpt.acc_tr_type=? ");
			sql.append(" and loanapply_rpt.acc_div=? ");
			sql.append(" and loanapply_rpt.applydate = TO_DATE(?, 'YYYY/MM/DD') ");
			sql.append(" and loanapply_rpt.bank_code in (").append(selectBank_no).append(") ");
			sql.append(" and loanapply_rpt.acc_code in (").append(selectAccCode).append(") ");
			paramList.add(acc_tr_type);
			paramList.add(acc_div);
			paramList.add(applydate);
			sql.append(" union ");
			sql.append(" select '',");
			sql.append(" '9999999' as bank_code,");
			sql.append(" '' as bank_name,");//--貸款經辦機構名稱
			sql.append(" '',null,'','','合 計',");
			sql.append(" sum(apply_cnt) as apply_cnt,");//--申請件數
			sql.append(" round(sum(apply_amt)/?,0) as apply_amt,");//--申請.貸款金額
			sql.append(" round(sum(apply_bal)/?,0) as apply_bal,");//--申請.貸款餘額
			sql.append(" sum(apply_cnt_sum) as apply_cnt_sum,");//--申請累計件數
			sql.append(" round(sum(apply_amt_sum)/?,0) as apply_amt_sum,");//--申請累計.貸款金額
			sql.append(" round(sum(apply_bal_sum)/?,0) as apply_bal_sum,");//--申請累計.貸款餘額
			sql.append(" sum(appr_cnt) as appr_cnt,");//--核准件數
			sql.append(" round(sum(appr_amt)/?,0) as appr_amt,");//--核准.貸款金額
			sql.append(" round(sum(appr_bal)/?,0) as appr_bal,");//--核准.貸款餘額
			sql.append(" sum(appr_cnt_sum) as appr_cnt,");//--核准累計件數
			sql.append(" round(sum(appr_amt_sum)/?,0) as appr_amt_sum,");//--核准累計.貸款金額
			sql.append(" round(sum(appr_bal_sum)/?,0) as appr_bal_sum,");//--核准累計.貸款餘額
			for(int i=0;i<8;i++){
				paramList.add(unit);
			}
			sql.append(" null,'','','' ");
			sql.append(" from loanapply_rpt ");
			sql.append(" left join loanapply_ncacno ");
			sql.append(" on loanapply_rpt.acc_tr_type=loanapply_ncacno.acc_tr_type and loanapply_rpt.acc_div=loanapply_ncacno.acc_div and loanapply_rpt.acc_code=loanapply_ncacno.acc_code ");
			sql.append(" left join loanapply_period on loanapply_rpt.acc_tr_type = loanapply_period.acc_tr_type and loanapply_rpt.applydate = loanapply_period.applydate ");
			sql.append(" left join (select * from bn01 where m_year=100)bn01 on loanapply_rpt.bank_code=bn01.bank_no ");
			sql.append(" where loanapply_rpt.acc_tr_type=? ");
			sql.append(" and loanapply_rpt.acc_div=? ");
			sql.append(" and loanapply_rpt.applydate = TO_DATE(?, 'YYYY/MM/DD') ");
			sql.append(" and loanapply_rpt.bank_code in (").append(selectBank_no).append(") ");
			sql.append(" and loanapply_rpt.acc_code in (").append(selectAccCode).append(") ");
			paramList.add(acc_tr_type);
			paramList.add(acc_div);
			paramList.add(applydate);
			List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "hsien_name,bank_code,bank_name,applytype,sumperiod,applydate_period,acc_code,acc_name,"+
													"apply_cnt,apply_amt,apply_bal,apply_cnt_sum,apply_amt_sum,apply_bal_sum,"+
													"appr_cnt,appr_amt,appr_bal,appr_cnt_sum,appr_amt_sum,appr_bal_sum,nonappr_cnt,nonappr_reason,cnt_name,cnt_tel");
			System.out.println("getDataList1.size()="+dbData.size());
			return dbData;
		}
	 //多家機構明細資料.彈性報表查詢
	 public static List getDataList2(String acc_tr_type,String acc_div,String unit,String applydate,String selectAccCode,String selectBank_no){
			StringBuffer sql = new StringBuffer();
			List paramList = new ArrayList();//共同參數
			sql.append(" select applytype,");//--申報頻率 
			sql.append(" '(自'||F_TRANSCHINESEDATE(begindate)||'起至'|| F_TRANSCHINESEDATE(loanapply_rpt.applydate)||')'  as sumperiod ,");//--累計期間
			sql.append(" F_TRANSCHINESEDATE(applydate_b) || '至' || F_TRANSCHINESEDATE(applydate_e) applydate_period,");//--申報期間
			sql.append(" loanapply_rpt.acc_code,");
			sql.append(" acc_name,");//--貸款種類
			sql.append(" sum(apply_cnt) as apply_cnt,");//--申請件數
			sql.append(" round(sum(apply_amt)/?,0) as apply_amt,");//--申請.貸款金額
			sql.append(" round(sum(apply_bal)/?,0) as apply_bal,");//--申請.貸款餘額
			sql.append(" sum(apply_cnt_sum) as apply_cnt_sum,");//--申請累計件數
			sql.append(" round(sum(apply_amt_sum)/?,0) as apply_amt_sum,");//--申請累計.貸款金額
			sql.append(" round(sum(apply_bal_sum)/?,0) as apply_bal_sum,");//--申請累計.貸款餘額
			sql.append(" sum(appr_cnt) as appr_cnt,");//--核准件數
			sql.append(" round(sum(appr_amt)/?,0) as appr_amt,");//--核准.貸款金額
			sql.append(" round(sum(appr_bal)/?,0) as appr_bal,");//--核准.貸款餘額
			sql.append(" sum(appr_cnt_sum) as appr_cnt_sum,");//--核准累計件數
			sql.append(" round(sum(appr_amt_sum)/?,0) as appr_amt_sum,");//--核准累計.貸款金額
			sql.append(" round(sum(appr_bal_sum)/?,0) as appr_bal_sum,");//--核准累計.貸款餘額
			for(int i=0;i<8;i++){
				paramList.add(unit);
			}
			sql.append(" sum(nonappr_cnt) as nonappr_cnt ");//--不予核貸件數
			sql.append(" from loanapply_rpt ");
			sql.append(" left join loanapply_ncacno ");
			sql.append(" on loanapply_rpt.acc_tr_type=loanapply_ncacno.acc_tr_type and loanapply_rpt.acc_div=loanapply_ncacno.acc_div and loanapply_rpt.acc_code=loanapply_ncacno.acc_code ");
			sql.append(" left join loanapply_period on loanapply_rpt.acc_tr_type = loanapply_period.acc_tr_type and loanapply_rpt.applydate = loanapply_period.applydate ");
			sql.append(" left join (select * from bn01 where m_year=100)bn01 on loanapply_rpt.bank_code=bn01.bank_no ");
			sql.append(" where loanapply_rpt.acc_tr_type=? ");
			sql.append(" and loanapply_rpt.acc_div=? ");
			sql.append(" and loanapply_rpt.applydate = TO_DATE(?, 'YYYY/MM/DD') ");
			sql.append(" and loanapply_rpt.bank_code in (").append(selectBank_no).append(") ");//--UI.所選取多家貸款經辦機構
			sql.append(" and loanapply_rpt.acc_code in (").append(selectAccCode).append(") ");
			paramList.add(acc_tr_type);
			paramList.add(acc_div);
			paramList.add(applydate);
			sql.append(" group by applytype,");//--申報頻率 
			sql.append(" '(自'||F_TRANSCHINESEDATE(begindate)||'起至'|| F_TRANSCHINESEDATE(loanapply_rpt.applydate)||')'  ,");//--累計期間
			sql.append(" F_TRANSCHINESEDATE(applydate_b) || '至' || F_TRANSCHINESEDATE(applydate_e) ,");//--申報期間
			sql.append(" loanapply_rpt.acc_code,acc_name ");//--貸款種類
			sql.append(" union  ");
			sql.append(" select '','','','','合 計',");
			sql.append(" sum(apply_cnt) as apply_cnt,");//--申請件數
			sql.append(" round(sum(apply_amt)/?,0) as apply_amt,");//--申請.貸款金額
			sql.append(" round(sum(apply_bal)/?,0) as apply_bal,");//--申請.貸款餘額
			sql.append(" sum(apply_cnt_sum) as apply_cnt_sum,");//--申請累計件數
			sql.append(" round(sum(apply_amt_sum)/?,0) as apply_amt_sum,");//--申請累計.貸款金額
			sql.append(" round(sum(apply_bal_sum)/?,0) as apply_bal_sum,");//--申請累計.貸款餘額
			sql.append(" sum(appr_cnt) as appr_cnt,");//--核准件數
			sql.append(" round(sum(appr_amt)/?,0) as appr_amt,");//--核准.貸款金額
			sql.append(" round(sum(appr_bal)/?,0) as appr_bal,");//--核准.貸款餘額
			sql.append(" sum(appr_cnt_sum) as appr_cnt,");//--核准累計件數
			sql.append(" round(sum(appr_amt_sum)/?,0) as appr_amt_sum,");//--核准累計.貸款金額
			sql.append(" round(sum(appr_bal_sum)/?,0) as appr_bal_sum,");//--核准累計.貸款餘額
			for(int i=0;i<8;i++){
				paramList.add(unit);
			}
			sql.append(" sum(nonappr_cnt) as nonappr_cnt ");//--不予核貸件數
			sql.append(" from loanapply_rpt ");
			sql.append(" left join loanapply_ncacno ");
			sql.append(" on loanapply_rpt.acc_tr_type=loanapply_ncacno.acc_tr_type and loanapply_rpt.acc_div=loanapply_ncacno.acc_div and loanapply_rpt.acc_code=loanapply_ncacno.acc_code ");
			sql.append(" left join loanapply_period on loanapply_rpt.acc_tr_type = loanapply_period.acc_tr_type and loanapply_rpt.applydate = loanapply_period.applydate ");
			sql.append(" left join (select * from bn01 where m_year=100)bn01 on loanapply_rpt.bank_code=bn01.bank_no ");
			sql.append(" where loanapply_rpt.acc_tr_type=? ");
			sql.append(" and loanapply_rpt.acc_div=? ");
			sql.append(" and loanapply_rpt.applydate = TO_DATE(?, 'YYYY/MM/DD')  ");
			sql.append(" and loanapply_rpt.bank_code in (").append(selectBank_no).append(") ");//--UI.所選取多家貸款經辦機構
			sql.append(" and loanapply_rpt.acc_code in (").append(selectAccCode).append(") ");
			paramList.add(acc_tr_type);
			paramList.add(acc_div);
			paramList.add(applydate);
			
			List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "applytype,sumperiod,applydate_period,acc_code,acc_name,"+
																	"apply_cnt,apply_amt,apply_bal,apply_cnt_sum,apply_amt_sum,apply_bal_sum,"+
																	"appr_cnt,appr_amt,appr_bal,appr_cnt_sum,appr_amt_sum,appr_bal_sum,nonappr_cnt");
			System.out.println("getDataList2.size()="+dbData.size());
			return dbData;
	}
	 //取得多家機構不予核貸原因查詢SQL:
	 public static List getNonappr_reason2(String acc_tr_type,String acc_div,String unit,String applydate,String acc_Code,String selectBank_no){
			StringBuffer sql = new StringBuffer();
			List paramList = new ArrayList();//共同參數
			sql.append(" select loanapply_rpt.acc_code,");
			sql.append(" nonappr_reason ");//--不予核貸原因
			sql.append(" from loanapply_rpt ");
			sql.append(" where acc_tr_type=? ");
			sql.append(" and acc_div=?");
			sql.append(" and applydate = TO_DATE(?, 'YYYY/MM/DD') ");
			sql.append(" and bank_code in (").append(selectBank_no).append(") ");//--UI.所選取多家貸款經辦機構
			sql.append(" and loanapply_rpt.acc_code =? ");
			sql.append(" and length(nonappr_reason) !=0 ");
			paramList.add(acc_tr_type);
			paramList.add(acc_div);
			paramList.add(applydate);
			paramList.add(acc_Code);
			List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "acc_code,nonappr_reason");
			System.out.println("getNonappr_reason2.size()="+dbData.size());
			return dbData;
	}
	//縣市統計資料.彈性報表查詢
	 public static List getDataList3(String acc_tr_type,String acc_div,String unit,String applydate,String selectAccCode,String selectBank_no){
			StringBuffer sql = new StringBuffer();
			List paramList = new ArrayList();//共同參數
			sql.append(" select loanapply_rpt.hsien_id ,loanapply_rpt.hsien_name,");//--縣市別
			sql.append("        loanapply_rpt.FR001W_output_order,");      
			sql.append("        COUNT_SEQ, field_SEQ,");  
			sql.append("        apply_cnt,");//--申請件數
			sql.append("        apply_amt,");//--申請.貸款金額
			sql.append("        apply_bal,");//--申請.貸款餘額
			sql.append("        apply_cnt_sum,");//--申請累計件數
			sql.append("        apply_amt_sum,");//--申請累計.貸款金額
			sql.append("        apply_bal_sum,");//--申請累計.貸款餘額
			sql.append("        appr_cnt,");//--核准件數
			sql.append("        appr_amt,");//--核准.貸款金額
			sql.append("        appr_bal,");//--核准.貸款餘額
			sql.append("        appr_cnt_sum,");//--核准累計件數
			sql.append("        appr_amt_sum,");//--核准累計.貸款金額
			sql.append("        appr_bal_sum,");//--核准累計.貸款餘額
			sql.append("        nonappr_cnt ");//--不予核貸件數
			sql.append(" from ");
			sql.append(" ( ");//--縣市小計
			sql.append("  select loanapply_rpt.hsien_id , loanapply_rpt.hsien_name,  loanapply_rpt.FR001W_output_order,"); 
			sql.append("         ' ' AS  bank_code , ' ' AS  BANK_NAME, count(*)  AS  COUNT_SEQ,'A90'  as  field_SEQ,");   
			sql.append("         sum(apply_cnt) as apply_cnt,");//--申請件數
			sql.append("         round(sum(apply_amt)/?,0) as apply_amt,");//--申請.貸款金額
			sql.append("         round(sum(apply_bal)/?,0) as apply_bal,");//--申請.貸款餘額
			sql.append("         sum(apply_cnt_sum) as apply_cnt_sum,");//--申請累計件數
			sql.append("         round(sum(apply_amt_sum)/?,0) as apply_amt_sum,");//--申請累計.貸款金額
			sql.append("         round(sum(apply_bal_sum)/?,0) as apply_bal_sum,");//--申請累計.貸款餘額
			sql.append("         sum(appr_cnt) as appr_cnt,");//--核准件數
			sql.append("         round(sum(appr_amt)/?,0) as appr_amt,");//--核准.貸款金額
			sql.append("         round(sum(appr_bal)/?,0) as appr_bal,");//--核准.貸款餘額
			sql.append("         sum(appr_cnt_sum) as appr_cnt_sum,");//--核准累計件數
			sql.append("         round(sum(appr_amt_sum)/?,0) as appr_amt_sum,");//--核准累計.貸款金額
			sql.append("         round(sum(appr_bal_sum)/?,0) as appr_bal_sum,");//--核准累計.貸款餘額
			sql.append("         sum(nonappr_cnt) as nonappr_cnt ");//--不予核貸件數   
			for(int i=0;i<8;i++){
				paramList.add(unit);
			}
			sql.append("  from (  ");      
			sql.append("    select hsien_id,hsien_name,FR001W_output_order,bank_code,BANK_NAME,");
			sql.append("           sum(apply_cnt) as apply_cnt,");//--申請件數
			sql.append("           sum(apply_amt) as apply_amt,");//--申請.貸款金額
			sql.append("           sum(apply_bal) as apply_bal,");//--申請.貸款餘額
			sql.append("           sum(apply_cnt_sum) as apply_cnt_sum,");//--申請累計件數
			sql.append("           sum(apply_amt_sum) as apply_amt_sum,");//--申請累計.貸款金額
			sql.append("           sum(apply_bal_sum) as apply_bal_sum,");//--申請累計.貸款餘額
			sql.append("           sum(appr_cnt) as appr_cnt,");//--核准件數
			sql.append("           sum(appr_amt) as appr_amt,");//--核准.貸款金額
			sql.append("           sum(appr_bal) as appr_bal,");//--核准.貸款餘額
			sql.append("           sum(appr_cnt_sum) as appr_cnt_sum,");//--核准累計件數
			sql.append("           sum(appr_amt_sum) as appr_amt_sum,");//--核准累計.貸款金額
			sql.append("           sum(appr_bal_sum) as appr_bal_sum,");//--核准累計.貸款餘額
			sql.append("           sum(nonappr_cnt) as nonappr_cnt ");//--不予核貸件數
			sql.append("    from(  ");                     
			sql.append("      select nvl(cd01.hsien_id,' ')       as  hsien_id ,");               
			sql.append("             nvl(cd01.hsien_name,'OTHER') as  hsien_name,");               
			sql.append("             cd01.FR001W_output_order     as  FR001W_output_order,");               
			sql.append("             loanapply_bn01.bank_code ,  loanapply_bn01.BANK_NAME,"); 
			sql.append("             sum(apply_cnt) as apply_cnt,");//--申請件數
			sql.append("             round(sum(apply_amt)/'1',0) as apply_amt,");//--申請.貸款金額
			sql.append("             round(sum(apply_bal)/'1',0) as apply_bal,");//--申請.貸款餘額
			sql.append("             sum(apply_cnt_sum) as apply_cnt_sum,");//--申請累計件數
			sql.append("             round(sum(apply_amt_sum)/'1',0) as apply_amt_sum,");//--申請累計.貸款金額
			sql.append("             round(sum(apply_bal_sum)/'1',0) as apply_bal_sum,");//--申請累計.貸款餘額
			sql.append("             sum(appr_cnt) as appr_cnt,");//--核准件數
			sql.append("             round(sum(appr_amt)/'1',0) as appr_amt,");//--核准.貸款金額
			sql.append("             round(sum(appr_bal)/'1',0) as appr_bal,");//--核准.貸款餘額
			sql.append("             sum(appr_cnt_sum) as appr_cnt_sum,");//--核准累計件數
			sql.append("             round(sum(appr_amt_sum)/'1',0) as appr_amt_sum,");//--核准累計.貸款金額
			sql.append("             round(sum(appr_bal_sum)/'1',0) as appr_bal_sum,");//--核准累計.貸款餘額
			sql.append("             sum(nonappr_cnt) as nonappr_cnt ");//--不予核貸件數
			sql.append("      from  (select * from  cd01 where cd01.hsien_id <> 'Y'  ) cd01 ");
			sql.append("      left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = 100 and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL) "); 
			sql.append("      left join (select loanapply_bn01.*,bank_name from loanapply_bn01 left join bn01 on loanapply_bn01.bank_code=bn01.bank_no and bn01.m_year=100 where acc_tr_type=?)loanapply_bn01 ");
			paramList.add(acc_tr_type);
			sql.append(" on wlx01.bank_no=loanapply_bn01.bank_code and wlx01.m_year=100 "); 
			sql.append("      left join (select * from loanapply_rpt where acc_tr_type=? and acc_div=? ");
			sql.append(" and applydate = TO_DATE(?, 'YYYY/MM/DD')) loanapply_rpt "); 
			paramList.add(acc_tr_type);
			paramList.add(acc_div);
			paramList.add(applydate);
			sql.append(" on  loanapply_bn01.bank_code = loanapply_rpt.bank_code ");
			sql.append("      group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,loanapply_bn01.bank_code,loanapply_bn01.BANK_NAME,acc_code  ");           
			sql.append("    )group by hsien_id,hsien_name,FR001W_output_order,bank_code,BANK_NAME ");
			sql.append("  )loanapply_rpt where loanapply_rpt.bank_code <> ' ' ");
			sql.append("  GROUP BY loanapply_rpt.hsien_id ,loanapply_rpt.hsien_name,loanapply_rpt.FR001W_output_order ");
			sql.append("  union ");
			sql.append("  select ' ' AS  hsien_id ,'合計' AS hsien_name, '999' AS FR001W_output_order,");                
			sql.append("         ' ' AS  bank_code,' ' AS  BANK_NAME,COUNT(*) AS COUNT_SEQ,'A99' as  field_SEQ,");     
			sql.append("         sum(apply_cnt) as apply_cnt,");//--申請件數
			sql.append("         round(sum(apply_amt)/?,0) as apply_amt,");//--申請.貸款金額
			sql.append("         round(sum(apply_bal)/?,0) as apply_bal,");//--申請.貸款餘額
			sql.append("         sum(apply_cnt_sum) as apply_cnt_sum,");//--申請累計件數
			sql.append("         round(sum(apply_amt_sum)/?,0) as apply_amt_sum,");//--申請累計.貸款金額
			sql.append("         round(sum(apply_bal_sum)/?,0) as apply_bal_sum,");//--申請累計.貸款餘額
			sql.append("         sum(appr_cnt) as appr_cnt,");//--核准件數
			sql.append("         round(sum(appr_amt)/?,0) as appr_amt,");//--核准.貸款金額
			sql.append("         round(sum(appr_bal)/?,0) as appr_bal,");//--核准.貸款餘額
			sql.append("         sum(appr_cnt_sum) as appr_cnt_sum,");//--核准累計件數
			sql.append("         round(sum(appr_amt_sum)/?,0) as appr_amt_sum,");//--核准累計.貸款金額
			sql.append("         round(sum(appr_bal_sum)/?,0) as appr_bal_sum,");//--核准累計.貸款餘額
			sql.append("         sum(nonappr_cnt) as nonappr_cnt ");//--不予核貸件數
			for(int i=0;i<8;i++){
				paramList.add(unit);
			}
			sql.append("   from ( ");
			sql.append("     select hsien_id,hsien_name,FR001W_output_order,bank_code,BANK_NAME,");
			sql.append("            sum(apply_cnt) as apply_cnt,");//--申請件數
			sql.append("            sum(apply_amt) as apply_amt,");//--申請.貸款金額
			sql.append("            sum(apply_bal) as apply_bal,");//--申請.貸款餘額
			sql.append("            sum(apply_cnt_sum) as apply_cnt_sum,");//--申請累計件數
			sql.append("            sum(apply_amt_sum) as apply_amt_sum,");//--申請累計.貸款金額
			sql.append("            sum(apply_bal_sum) as apply_bal_sum,");//--申請累計.貸款餘額
			sql.append("            sum(appr_cnt) as appr_cnt,");//--核准件數
			sql.append("            sum(appr_amt) as appr_amt,");//--核准.貸款金額
			sql.append("            sum(appr_bal) as appr_bal,");//--核准.貸款餘額
			sql.append("            sum(appr_cnt_sum) as appr_cnt_sum,");//--核准累計件數
			sql.append("            sum(appr_amt_sum) as appr_amt_sum,");//--核准累計.貸款金額
			sql.append("            sum(appr_bal_sum) as appr_bal_sum,");//--核准累計.貸款餘額
			sql.append("            sum(nonappr_cnt) as nonappr_cnt ");//--不予核貸件數
			sql.append("     from( ");                            
			sql.append("       select nvl(cd01.hsien_id,' ')  as  hsien_id ,");               
			sql.append("              nvl(cd01.hsien_name,'OTHER') as  hsien_name,");               
			sql.append("              cd01.FR001W_output_order  as  FR001W_output_order,");               
			sql.append("              loanapply_bn01.bank_code ,loanapply_bn01.BANK_NAME,"); 
			sql.append("              sum(apply_cnt) as apply_cnt,");//--申請件數
			sql.append("              round(sum(apply_amt)/'1',0) as apply_amt,");//--申請.貸款金額
			sql.append("              round(sum(apply_bal)/'1',0) as apply_bal,");//--申請.貸款餘額
			sql.append("              sum(apply_cnt_sum) as apply_cnt_sum,");//--申請累計件數
			sql.append("              round(sum(apply_amt_sum)/'1',0) as apply_amt_sum,");//--申請累計.貸款金額
			sql.append("              round(sum(apply_bal_sum)/'1',0) as apply_bal_sum,");//--申請累計.貸款餘額
			sql.append("              sum(appr_cnt) as appr_cnt,");//--核准件數
			sql.append("              round(sum(appr_amt)/'1',0) as appr_amt,");//--核准.貸款金額
			sql.append("              round(sum(appr_bal)/'1',0) as appr_bal,");//--核准.貸款餘額
			sql.append("              sum(appr_cnt_sum) as appr_cnt_sum,");//--核准累計件數
			sql.append("              round(sum(appr_amt_sum)/'1',0) as appr_amt_sum,");//--核准累計.貸款金額
			sql.append("              round(sum(appr_bal_sum)/'1',0) as appr_bal_sum,");//--核准累計.貸款餘額
			sql.append("              sum(nonappr_cnt) as nonappr_cnt ");//--不予核貸件數
			sql.append("       from  (select * from  cd01 where cd01.hsien_id <> 'Y'  ) cd01 ");
			sql.append("       left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = 100 and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL)  ");
			sql.append("       left join (select loanapply_bn01.*,bank_name from loanapply_bn01 left join bn01 on loanapply_bn01.bank_code=bn01.bank_no and bn01.m_year=100  where acc_tr_type=?)loanapply_bn01 ");
			paramList.add(acc_tr_type);
			sql.append(" on wlx01.bank_no=loanapply_bn01.bank_code and wlx01.m_year=100  ");
			sql.append("       left join (select * from loanapply_rpt where acc_tr_type=? and acc_div=? ");
			sql.append(" and applydate = TO_DATE(?, 'YYYY/MM/DD')) loanapply_rpt  ");
			sql.append(" on  loanapply_bn01.bank_code = loanapply_rpt.bank_code ");
			sql.append("       group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,loanapply_bn01.bank_code,loanapply_bn01.BANK_NAME,acc_code ");                  
			sql.append("     )group by hsien_id,hsien_name,FR001W_output_order,bank_code,BANK_NAME ");
			sql.append("   ) loanapply_rpt where loanapply_rpt.bank_code <> ' ' "); 
			sql.append(" ) loanapply_rpt ");
			sql.append(" ORDER by FR001W_output_order,field_SEQ,hsien_id,bank_code ");
			paramList.add(acc_tr_type);
			paramList.add(acc_div);
			paramList.add(applydate);
			List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "hsien_id,hsien_name,fr001w_output_order,count_seq, field_seq,"+
																		"apply_cnt,apply_amt,apply_bal,apply_cnt_sum,apply_amt_sum,apply_bal_sum,"+
																		"appr_cnt,appr_amt,appr_bal,appr_cnt_sum,appr_amt_sum,appr_bal_sum,nonappr_cnt");
			System.out.println("getDataList3.size()="+dbData.size());
			return dbData;
	}
	//取得該縣市不予核貸原因查詢
	 public static List getNonappr_reason3(String acc_tr_type,String acc_div,String unit,String applydate,String hsien_id){
			StringBuffer sql = new StringBuffer();
			List paramList = new ArrayList();//共同參數
			sql.append(" select hsien_id,loanapply_rpt.acc_code,");
			sql.append(" nonappr_reason ");//--不予核貸原因
			sql.append(" from loanapply_rpt ");
			sql.append(" left join (select wlx01.*,hsien_name from wlx01 left join cd01 on wlx01.hsien_id=cd01.hsien_id where m_year=100)wlx01 on loanapply_rpt.bank_code=wlx01.bank_no ");
			sql.append(" where acc_tr_type=? ");
			sql.append(" and acc_div=? ");
			sql.append(" and applydate = TO_DATE(?, 'YYYY/MM/DD') ");
			sql.append(" and hsien_id = ? ");//查詢清單上的縣市別代碼
			sql.append(" and length(nonappr_reason) !=0 ");
			paramList.add(acc_tr_type);
			paramList.add(acc_div);
			paramList.add(applydate);
			paramList.add(hsien_id);	
			List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "hsien_id,acc_code,nonappr_reason");
			System.out.println("getNonappr_reason3.size()="+dbData.size());
			return dbData;
	}
	 public static String getAcc_Tr_Name(String acc_tr_type){
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
	 public static String getApplydate_period(String acc_tr_type,String applydate){
			String rtnVal = "";
	        StringBuffer sqlCmd = new StringBuffer();
	        List paramList = new ArrayList();//傳內的參數List   
	        sqlCmd.append(" select F_TRANSCHINESEDATE(applydate_b) || '至' || F_TRANSCHINESEDATE(applydate_e) applydate_period ");
	        sqlCmd.append("   from loanapply_period ");
	        sqlCmd.append("  where acc_tr_type=? and applydate = TO_DATE(?, 'YYYY/MM/DD') ");
	        paramList.add(acc_tr_type); 
	        paramList.add(applydate);
	        List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
	        if(dbData!=null && dbData.size()>0){
	        	DataObject bean = (DataObject)dbData.get(0);
	        	rtnVal = (String)bean.getValue("applydate_period");
	        }
	        return  rtnVal; 
	 }
}
