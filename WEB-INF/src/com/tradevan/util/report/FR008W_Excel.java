/*
 * Created on 2005/1/19
 *
 * TODO To change the template for this generated file go to
 * Window - Preferences - Java - Code Style - Code Templates
 */
/*
  94.03.11 fix 金額資料為零是不輸出0，改為輸出空白及欄位資料右靠處理
  95.05.15 fix 報表格式加入910203/910404.備註 by 2295
  95.08.17 add 910202-910203 < 0 ,則910202-910203等於0
 		   add 註2:備抵呆帳、損失準備及營業準備(第二類資本)-910202 by 2295
  95.08.21 fix 報表格式/列印該月份最後一天 by 2295	
           add 金額單位 by 2295	  
  95.09.25 fix 若單位為1000時,利率會為0 by 2295
           add 若910202超過910500 * 1.25% ,以910500 * 1.25%顯示 
  95.10.03 增加檢核結果與最後異動日期 by 2495              
  96.01.22 fix 備註加上單位別 by 2295     
  96.07.05 if【910204】< (【910500】* 1.25% )       by 2295
		    Temp_910202 = 【910202】-【910203】
		   Else (為910204 = 910500 * 1.25%)
		    Temp_910202 = 【910204】            
  99.04.28 fix sql injection by 2808
 108.01.08 fix 調整報表格式 by 2295
 */
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*;
import java.io.*;
import java.math.BigDecimal;
import java.util.*;
import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;
/**
 * @author 2295
 * 
 * TODO To change the template for this generated type comment go to Window -
 * Preferences - Java - Code Style - Code Templates
 */
public class FR008W_Excel {
	final private static String[] table = {"910101", "910102", "910103", "910104", "910105", "910106", "910108", "910109", "910110", "910199", "910201", "910202", "910299", "910300", "910401", "910402", "910403", "910404", "910400", "910500", "91060P"};
	public static String createRpt(String S_YEAR, String S_MONTH, String bank_code, String BANK_NAME,String bank_type,String unit) {
		String errMsg = "";
		List dbData = null;
		String sqlCmd = "";
		String ncacno="";
		String unit_name="";
		BigDecimal tmpZero=new BigDecimal("0");	
		StringBuffer sql = new StringBuffer () ;
		ArrayList paramList = new ArrayList() ;
		DataObject bean = null ;
		String u_year = "100" ;
		if(S_YEAR==null || Integer.parseInt(S_YEAR)<=99) {
			u_year = "99" ;
		}
		BANK_NAME = BANK_NAME==null ? "" : BANK_NAME.substring(7) ; //取中文不顯示機構代號
		try {
			System.out.println("信用部淨值占風險性資產比率.xls");
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
			FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator") + (bank_type.equals("6")?"農會":"漁會")+"信用部淨值占風險性資產比率.xls");
			ncacno = bank_type.equals("6")?"ncacno":"ncacno_7";	
			// 設定FileINputStream讀取Excel檔
			POIFSFileSystem fs = new POIFSFileSystem(finput);
			HSSFWorkbook wb = new HSSFWorkbook(fs);
			HSSFSheet sheet = wb.getSheetAt(0);// 讀取第一個工作表，宣告其為sheet
			HSSFPrintSetup ps = sheet.getPrintSetup(); // 取得設定
			// sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
			// sheet.setAutobreaks(true); //自動分頁

			// 設定頁面符合列印大小
			sheet.setAutobreaks(false);
			ps.setScale((short) 120); // 列印縮放百分比

			ps.setPaperSize((short) 9); // 設定紙張大小 A4
			// wb.setSheetName(0,"test");
			finput.close();

			HSSFRow row = null;// 宣告一列
			HSSFCell cell = null;// 宣告一個儲存格

			short rowNo = 0;
			// short y = 0;
			//95.09.25 fix 若單位為1000時,利率會為0 by 2295
			//sqlCmd = " select m_year,m_month,"+ncacno+".acc_code,"+ncacno+".acc_name,"
   			//        //+ " round(decode(substr("+ncacno+".acc_code,length("+ncacno+".acc_code)),'P',nvl(a05.amt,0)/1000,nvl(a05.amt,0))/"+unit+",0) as amt, "
   			//        + " decode(substr("+ncacno+".acc_code,length("+ncacno+".acc_code)),'P',nvl(a05.amt,0)/1000,round(nvl(a05.amt,0)/"+unit+",0)) as amt," 
   			//        + " nvl(a05.amt_name,'') as amt_name "
			//		+ " from "+ncacno+" left join a05 on "+ncacno+".acc_code=a05.acc_code "
			//		+ " and "+ncacno+".acc_code like '91%' " + " and   m_year = " + S_YEAR + " and   m_month = " + String.valueOf(Integer.parseInt(S_MONTH)) + " and   a05.bank_code = '" + bank_code+"'"
			//		+ " where acc_tr_type='A05'"
			//		+ " order by "+ncacno+".acc_range";
			sql.append(" select m_year,m_month,"+ncacno+".acc_code,"+ncacno+".acc_name, ") ;
			sql.append(" decode(substr("+ncacno+".acc_code,length("+ncacno+".acc_code)),'P',nvl(a05.amt,0)/1000,round(nvl(a05.amt,0)/? ,0)) as amt,");
			sql.append(" nvl(a05.amt_name,'') as amt_name ");
			sql.append(" from "+ncacno+" left join a05 on "+ncacno+".acc_code=a05.acc_code ");
			sql.append(" and "+ncacno+".acc_code like ? " + " and   m_year = ?  and   m_month = ?  and   a05.bank_code = ? ");
			sql.append(" where acc_tr_type= ?  ");
			sql.append(" order by "+ncacno+".acc_range ");
			paramList.add(unit) ;
			paramList.add("91%");
			paramList.add(S_YEAR) ;
			paramList.add(S_MONTH) ;
			paramList.add(bank_code) ;
			paramList.add("A05") ;
			dbData = DBManager.QueryDB_SQLParam(sql.toString(), paramList, "m_year,m_month,amt") ;
				     //DBManager.QueryDB(sqlCmd, "m_year,m_month,amt");
			sql.setLength(0) ;
			paramList.clear();
			// 取出資料存入MAP
			HashMap h = new HashMap();
			for (int k = 0; k < dbData.size(); k++) {
				DataObject obj = (DataObject) dbData.get(k);
				h.put(obj.getValue("acc_code"), obj.getValue("amt").toString());
			}
			
			rowNo = 0;			
			row = sheet.getRow(rowNo);
			cell = row.getCell((short) 0);
			// 設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue(BANK_NAME + "淨值占風險性資產比率計算表");
			String lastday = Utility.getCHTdate(Utility.getLastDay(String.valueOf(Integer.parseInt(S_YEAR)+1911)+"/"+S_MONTH,"yyyy/mm"),1);
			
            rowNo=1;
			if (dbData.size() == 0) {
				row = sheet.getRow(rowNo);
				cell = row.getCell((short) 0);
				// 設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);			
				cell.setCellValue(S_YEAR + "年" + S_MONTH + "月"+lastday.substring(0,10).substring(8,10)+"日無資料存在");
			} else {
				row = sheet.getRow(rowNo);
				cell = row.getCell((short) 0);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellValue(S_YEAR + "年" + S_MONTH + "月"+lastday.substring(0,10).substring(8,10)+"日");				
				cell = row.getCell((short) 1);
				unit_name = Utility.getUnitName(unit) ;
				
   	       	    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
   	 	   	 	cell.setCellValue("單位：新台幣"+unit_name+"、％");
				
				rowNo = 3;
				BigDecimal tmpAmt = null;
				BigDecimal tmpAmt_910500 = null;
				BigDecimal tmpAmt_910500_910203 = null;
				BigDecimal tmpAmt_910202 = null;
				BigDecimal tmpAmt_910204 = null;//96.07.05 add 
				for (int i = 0; i < table.length; i++) {
					System.out.print("Row =" + rowNo+":"+table[i]);
					row = sheet.getRow(rowNo++);
					cell = row.getCell((short) 1);
					String amt = "0";
					if(table[i].equals("910108")){					    
					   tmpAmt = new BigDecimal(((String) h.get(table[i])));//910108
					   tmpAmt = tmpAmt.add( new BigDecimal(((String) h.get(table[i+1]))).negate());//910108-910109
					   amt = Utility.setCommaFormat(tmpAmt.toString());
					   System.out.println("910108-910109="+tmpAmt.toString());
					   i++;
					}else if(table[i].equals("910202")){
					    tmpAmt_910202 = new BigDecimal(((String) h.get(table[i])));//910202	
					   /*95.09.25 fix  
					   System.out.println("910203="+(String) h.get("910203"));
					   tmpAmt = new BigDecimal(((String) h.get(table[i])));//910202
					   tmpAmt = tmpAmt.add( new BigDecimal(((String) h.get("910203"))).negate());//910202-910203
					   //95.08.17 add 910202-910203 < 0 ,則910202-910203等於0
					   if(tmpAmt.compareTo(tmpZero) == -1){//910202-910203 < 0 
					       tmpAmt = new BigDecimal("0");
					   }
					   */
					   amt = Utility.setCommaFormat(tmpAmt_910202.toString());//108.01.08 fix 
				    }else{
				       if(table[i].equals("910500")){
				          tmpAmt_910500 = new BigDecimal(((String) h.get(table[i])));//910500  
				          tmpAmt_910500_910203 = tmpAmt_910500.add( new BigDecimal(((String) h.get("910203"))));//910500+910203
				       }
					   amt = Utility.setCommaFormat((String) h.get(table[i]));    
					}
					//System.out.println(":amt =" + amt);
					if (!amt.equals("0")) {
						cell.setCellValue(amt);
					}
				}
				
				/*96.07.05 if【910204】< (【910500】* 1.25% )
				              Temp_910202 = 【910202】-【910203】
				           Else (為910204 = 910500 * 1.25%)
				              Temp_910202 = 【910204】*/   
				/*108.01.08 取消因910204已為min([910202-910203],910500*1.75%)
				row = sheet.getRow((short)14);//910202
				cell = row.getCell((short) 1);				
				String amt = "0";
				BigDecimal tmp125=new BigDecimal("0.0125");	
				tmpAmt_910204 = h.get("910204") == null ? new BigDecimal("0") : new BigDecimal(((String) h.get("910204")));//910204
				//910500*1.25%->取到整數,小數點下以無條件捨去
				tmpAmt_910500 = new BigDecimal(tmpAmt_910500.multiply(tmp125).intValue());
			    //System.out.println("910500*1.25%="+tmpAmt_910500);	
			    
			    if(tmpAmt_910204.compareTo(tmpAmt_910500) == -1){//if【910204】< (【910500】* 1.25%
			       //System.out.println("910203="+(String) h.get("910203"));
				   tmpAmt = new BigDecimal(((String) h.get("910202")));//910202
				   tmpAmt = tmpAmt.add( new BigDecimal(((String) h.get("910203"))).negate());//910202-910203				   
			    }else{//為910204 = 910500 * 1.25%
			       tmpAmt = new BigDecimal(((String) h.get("910204")));//910204
			    }
			    //add 910202-910203 or 910204 < 0 ,則設為0
				if(tmpAmt.compareTo(tmpZero) == -1){//910202-910203 or 910204 < 0 
				   tmpAmt = new BigDecimal("0");
				}
				amt = Utility.setCommaFormat(tmpAmt.toString());				
				cell.setCellValue(amt);
				*/
				//==============================================================================
				String amt320200="0";
				//sqlCmd = " select m_year,m_month,acc_code,round(nvl(amt,0)/"+unit+",0)) as amt"  
				//	   + " from a01  "
				//	   + " where m_year = "+ S_YEAR 
				//	   + " and   m_month = " + String.valueOf(Integer.parseInt(S_MONTH)) 
				//	   + " and   bank_code = '" + bank_code+"'"
				//	   + " and   acc_code='320200'";
				sql.append(" select m_year,m_month,acc_code,round(nvl(amt,0)/?,0) as amt ");
				sql.append(" from a01 ");
				sql.append(" where m_year = ?");
				sql.append(" and   m_month = ? ");
				sql.append(" and   bank_code = ? ");
				sql.append(" and   acc_code= ? ");
				paramList.add(unit) ;
				paramList.add(S_YEAR) ;
				paramList.add(S_MONTH) ;
				paramList.add(bank_code) ;
				paramList.add("320200") ;
				dbData = DBManager.QueryDB_SQLParam(sql.toString(), paramList, "m_year,m_month,amt") ;
					     //DBManager.QueryDB(sqlCmd, "m_year,m_month,amt");
				sql.setLength(0) ;
				paramList.clear();
				if(dbData != null && dbData.size() != 0){
				    amt320200 = ((DataObject)dbData.get(0)).getValue("amt").toString();
				}
				//System.out.println("Row =" + rowNo);
				row = sheet.getRow(rowNo++);
				cell = row.getCell((short) 0);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				//96.01.22 fix 備註加上單位別 
				cell.setCellValue("註1:「(7)累積盈虧」有加計「上期損益: "+Utility.setCommaFormat(amt320200)+unit_name+"」及扣除「備抵呆帳、損失準備");
				row = sheet.getRow(rowNo++);
				cell = row.getCell((short) 0);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellValue("    及營業準備提列不足  "+Utility.setCommaFormat((String) h.get("910109"))+unit_name+"」");
				//95.08.17 add 2:備抵呆帳、損失準備及營業準備(第一類資本)-910202 ==============================================================
				row = sheet.getRow(rowNo++);
				cell = row.getCell((short) 0);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellValue("  2:帳列備抵呆帳、損失準備及營業準備 "+Utility.setCommaFormat((String) h.get("910202"))+unit_name);
				//========================================================================================================================
				row = sheet.getRow(rowNo++);
				cell = row.getCell((short) 0);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellValue("  3:特定損失所提列備抵呆帳、損失準備及營業準備  "+Utility.setCommaFormat((String) h.get("910203"))+unit_name);
				row = sheet.getRow(rowNo++);
				cell = row.getCell((short) 0);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellValue("  4:原風險性資產總額(未扣除特定損失)   "+Utility.setCommaFormat(tmpAmt_910500_910203.toString())+unit_name);//910500+910203
				
				//95.10.03 增加檢核結果與最後異動日期
				//sqlCmd = " select UPD_CODE, to_char(UPDATE_DATE,'yyyymmdd') as UPDATE_DATE"
				//	   + " from WML01"		   		  
				//	   + " where M_YEAR='"+S_YEAR+"'"		   		   
				//	   + " and M_MONTH='"+S_MONTH+"'"
				//	   + " and BANK_CODE='"+bank_code+"'"
				//	   + " and REPORT_NO='A05'";
				sql.append("  select UPD_CODE, to_char(UPDATE_DATE,'yyyymmdd') as UPDATE_DATE ") ;
				sql.append("  from WML01 ") ;
				sql.append("  where M_YEAR= ? ") ;
				sql.append("  and M_MONTH= ?  ") ;
				sql.append("  and BANK_CODE= ? ") ;
				sql.append("  and REPORT_NO= ? ") ;
				paramList.add(S_YEAR) ;
				paramList.add(S_MONTH) ;
				paramList.add(bank_code) ;
				paramList.add("A05") ;
				//System.out.println("sqlCmd="+sqlCmd); 	   
   		   
				dbData = DBManager.QueryDB_SQLParam(sql.toString(), paramList, "") ;
					     //DBManager.QueryDB(sqlCmd,"");	
				sql.setLength(0) ;
				paramList.clear();
				String UPD_CODE="";
				String UPDATE_DATE="";
				String M_YEAR="";   
				String M_MONTH="";    
				String M_DATE=""; 
				if(dbData.size()>0){
					//System.out.println("dbData.size()="+dbData.size()); 
					bean = (DataObject)dbData.get(0) ;
					UPD_CODE = (String)bean.getValue("upd_code");  
					UPDATE_DATE = (String)bean.getValue("update_date");       		   
					//System.out.println("UPD_CODE="+UPD_CODE); 
					//System.out.println("UPDATE_DATE="+UPDATE_DATE); 
					if(UPD_CODE.equals("N")) UPD_CODE="待檢核";
					else if(UPD_CODE.equals("E")) UPD_CODE="檢核錯誤";
					else if(UPD_CODE.equals("U")) UPD_CODE="檢核成功";
					else UPD_CODE="";
					//System.out.println("UPD_CODE="+UPD_CODE);
					M_YEAR  = Integer.toString(Integer.parseInt(UPDATE_DATE.substring(0,4))-1911);	
					M_MONTH  = Integer.toString(Integer.parseInt(UPDATE_DATE.substring(4,6))-0);	
					M_DATE  = Integer.toString(Integer.parseInt(UPDATE_DATE.substring(6,8))-0);	
					UPDATE_DATE=M_YEAR+"年"+M_MONTH+"月"+M_DATE+"日";	
					//System.out.println("UPDATE_DATE="+UPDATE_DATE); 	
					rowNo++;	
					rowNo++;    		
					row=sheet.getRow((short)rowNo);	    		
					cell=row.getCell((short)0);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);	  		
					cell.setCellValue("檢核結果:"+UPD_CODE);	  
					rowNo++;	 		
					row=sheet.getRow((short)rowNo);	    		
					cell=row.getCell((short)0);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellValue("最後異動日期:"+UPDATE_DATE);	 		   
				}else{
					rowNo++;	
					row=sheet.getRow(rowNo);	    		
					cell=row.getCell((short)0);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);	  		
					cell.setCellValue("檢核結果:待檢核");	  
					rowNo++;	 		
					row=sheet.getRow(rowNo);	    		
					cell=row.getCell((short)0);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellValue("最後異動日期:無");	     
				}//end of 檢核結果與最後異動日期
			}//end of 有資料存在
			
			FileOutputStream fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + "信用部淨值占風險性資產比率.xls");
			wb.write(fout);
			// 儲存
			fout.close();
		} catch (Exception e) {
			System.out.println("createRpt Error:" + e + e.getMessage());
		}
		return errMsg;
	}
}
