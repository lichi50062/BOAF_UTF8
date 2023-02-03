/*
99.08.06 add  全體農漁會信用部從業人員參加金融相關業務進修情形統計表 by 2660
99.08.20 fix 100年度區分縣市別才要加m_year條件  by 2295
99.08.31 fix 不管有無其參訓資訊,所有機構資料都要顯示 by 2295
99.10.15 fix 修改SQL只要有年累計都顯示 by 2295
99.10.19 add 增加合計 by 2295
*/
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import java.io.*;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR057W {

	public static String createRpt(String s_year, String s_month, String bank_type)	{

		String errMsg = "";
		String filename = "";
		filename = "全體農漁會信用部從業人員參加金融相關業務進修情形統計表";
		reportUtil reportUtil = new reportUtil();
		StringBuffer sql = new StringBuffer () ;
		ArrayList paramList = new ArrayList() ;
		//99.10.15 add 查詢年度100年以前.縣市別不同===============================
	    String cd01_table = (Integer.parseInt(s_year) < 100)?"cd01_99":""; 
	    String wlx01_m_year = (Integer.parseInt(s_year) < 100)?"99":"100"; 
	
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

			String openfile = "全體農漁會信用部從業人員參加金融相關業務進修情形統計表.xls";	
			System.out.println("open file " + openfile);
			FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+ openfile );			

			//設定FileINputStream讀取Excel檔
			POIFSFileSystem fs = new POIFSFileSystem( finput );
			if (fs == null) {
				System.out.println("Open 範本檔失敗");
			} else {
				System.out.println("Open 範本檔成功");
			}
			HSSFWorkbook wb = new HSSFWorkbook(fs);
			if (wb == null) {
				System.out.println("Open 工作表失敗");
			} else {
				System.out.println("Open 工作表成功");
			}
			HSSFSheet sheet = wb.getSheetAt(0); //讀取第一個工作表，宣告其為sheet
			if (sheet == null) {
				System.out.println("open sheet 失敗");
			} else {
				System.out.println("open sheet 成功");
			}
			HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
			// sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
			// sheet.setAutobreaks(true); //自動分頁

			// 設定頁面符合列印大小
			sheet.setAutobreaks(false);
			ps.setScale((short)90); //列印縮放百分比
			ps.setPaperSize((short)9); //設定紙張大小 A4

			//設定表頭 為固定 先設欄的起始再設列的起始
			wb.setRepeatingRowsAndColumns(0, 1, 16, 2, 4);
			finput.close();
	  		
			HSSFRow row = null;  //宣告一列
			HSSFCell cell = null;//宣告一個儲存格
			
			HSSFFont ft1 = wb.createFont();			
			HSSFCellStyle cs1 = wb.createCellStyle();
			HSSFCellStyle cs_center = reportUtil.getDefaultStyle(wb);
			HSSFCellStyle cs_left = reportUtil.getNoBorderLeftStyle(wb);

			ft1.setFontHeightInPoints((short)12);
			ft1.setFontName("細明體");
			cs1.setFont(ft1);
			cs1.setBorderTop((short)0);
			cs1.setBorderLeft((short)0);
			cs1.setBorderLeft((short)0);
			cs1.setBorderBottom((short)0);

			//套用DAO.preparestatment			
			//總計明細			
			sql.append(" select * ");
			sql.append(" from (  ");
			sql.append("    select nvl(cd01.hsien_id,' ')     as  hsien_id ,");
			sql.append("       	   nvl(cd01.hsien_name,'其他') as  hsien_name,");
			sql.append("       	   cd01.FR001W_output_order   as  FR001W_output_order,");
			sql.append("       	   bn01.bank_no,bn01.BANK_NAME,");
			sql.append("       	   wlx_trainning.wlx02count,credit_staff_num,person_mm, hour_mm, person_yy,hour_yy ");
			sql.append("    from (select * from "+cd01_table+" cd01 where cd01.hsien_id <> 'Y' ) cd01   ");
			sql.append("    left join (select * from wlx01 where m_year = ? AND cancel_no <> 'Y' OR cancel_no IS NULL)wlx01 on wlx01.hsien_id=cd01.hsien_id ");
			sql.append("    left join (select * from bn01 where  m_year=? and bn_type <> '2')bn01 on wlx01.bank_no=bn01.bank_no and bn01.bank_type= ? ");
			paramList.add(wlx01_m_year);
			paramList.add(wlx01_m_year);
			paramList.add(bank_type) ;
			sql.append("    left join ( ");
			sql.append("               select wlx02.bank_no,");//金融機構代碼
			sql.append("                      wlx02count,");//國內營業分支機構家數
			sql.append("                      total_hour,");  
			sql.append("                      person_mm,");//本月人次
			sql.append("                      hour_mm,");//本月時數 
			sql.append("                      person_yy,");//本年人次
			sql.append("                      hour_yy,");//本年時數
			sql.append("                      decode(hour_yy,null,0,hour_yy) - total_hour AS less_hour");
			sql.append("               from "); 								   
			sql.append("                  (select bn01.bank_no,decode(wlx02count,null,0,wlx02count) as wlx02count,");
			sql.append("                          decode(total_hour,null,16,total_hour) as total_hour");
			sql.append("                   from (select * from bn01 where m_year=? and bn_type <> '2' and bank_type in ('6','7'))bn01");
			paramList.add(wlx01_m_year);
			sql.append("                       ,(");
			sql.append("                         SELECT   tbank_no,");
			sql.append("                                  COUNT(*) AS wlx02count,");
			sql.append("                                  (COUNT (*) + 1) * 16 AS total_hour");
			sql.append("                         FROM wlx02");
			sql.append("                         WHERE  m_year = ? AND cancel_no <> 'Y' OR cancel_no IS NULL");
			paramList.add(wlx01_m_year);
			sql.append("				 		 GROUP BY tbank_no)wlx02");
			sql.append("				   where bn01.bank_no = wlx02.tbank_no(+)");
			sql.append("                   )wlx02 ,");
			sql.append("                  (SELECT yeardata.tbank_no,person_mm, hour_mm, person_yy, hour_yy");                                   
			sql.append("				   FROM(SELECT wlx_trainning.m_year,");
			sql.append("				  		       wlx_trainning.tbank_no,");
			sql.append("						       bank_name, COUNT (*) AS person_yy,");//本年人次
			sql.append("						       SUM (course_hour) AS hour_yy");//本年時數 
			sql.append("					    FROM wlx_trainning"); 
			sql.append("						LEFT JOIN (SELECT * FROM bn01 WHERE m_year = ? and bn_type <> '2') bn01 ON wlx_trainning.tbank_no = bn01.bank_no"); 
			sql.append("					    WHERE wlx_trainning.m_year = ?  and wlx_trainning.m_month <= ?");           
			sql.append("				        GROUP BY wlx_trainning.m_year,wlx_trainning.tbank_no,bank_name");
			sql.append("				  )yeardata"); 
			paramList.add(wlx01_m_year);
			paramList.add(s_year);
			paramList.add(s_month) ;
			sql.append("				  left join"); 
			sql.append("				 (SELECT wlx_trainning.m_year,");
			sql.append("					     wlx_trainning.m_month,");
			sql.append("					     wlx_trainning.tbank_no,");
			sql.append("					     bank_name, COUNT (*) AS person_mm,");//本月人次
			sql.append("					     SUM (course_hour) AS hour_mm");//本月時數
			sql.append("				  FROM wlx_trainning LEFT JOIN (SELECT * FROM bn01 WHERE m_year = ? and bn_type <> '2') bn01 ON wlx_trainning.tbank_no = bn01.bank_no");
			sql.append("    			  WHERE wlx_trainning.m_year = ? AND wlx_trainning.m_month = ?");            
			sql.append("				  GROUP BY wlx_trainning.m_year,wlx_trainning.m_month,wlx_trainning.tbank_no,bank_name");
			sql.append("				 ) monthdata on  monthdata.tbank_no = yeardata.tbank_no");
			paramList.add(wlx01_m_year);
			paramList.add(s_year);
			paramList.add(s_month) ;
			sql.append("				)wlx_trainning");								                                 
			sql.append("			where wlx02.bank_no = wlx_trainning.tbank_no(+)");                               
			sql.append("		    order by wlx02.bank_no");			
			sql.append("           ) wlx_trainning  on wlx_trainning.bank_no=bn01.bank_no ");
			sql.append("    ) Temp1 where bank_name is not null ");
			sql.append(" order by  Temp1.FR001W_output_order, Temp1.bank_no ");
			List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "wlx02count,credit_staff_num,person_mm,hour_mm,person_yy,hour_yy");
				
			//總表合計
			sql.delete(0,sql.length());
			paramList = new ArrayList() ;
			sql.append(" select '合計' as bank_no,");
			sql.append("         count(*) || '家' as bank_name,");
			sql.append(" 		 sum(wlx02count) as wlx02count,");
			sql.append("         sum(credit_staff_num) as credit_staff_num,");
			sql.append("         sum(person_mm) as person_mm,");
			sql.append("         sum(hour_mm) as hour_mm,");
			sql.append("         sum(person_yy) as person_yy,");
			sql.append("         sum(hour_yy) as hour_yy");
			sql.append(" from (");	
			sql.append(" select * ");
			sql.append(" from (  ");
			sql.append("    select nvl(cd01.hsien_id,' ')     as  hsien_id ,");
			sql.append("       	   nvl(cd01.hsien_name,'其他') as  hsien_name,");
			sql.append("       	   cd01.FR001W_output_order   as  FR001W_output_order,");
			sql.append("       	   bn01.bank_no,bn01.BANK_NAME,");
			sql.append("       	   wlx_trainning.wlx02count,credit_staff_num,person_mm, hour_mm, person_yy,hour_yy ");
			sql.append("    from (select * from "+cd01_table+" cd01 where cd01.hsien_id <> 'Y' ) cd01   ");
			sql.append("    left join (select * from wlx01 where m_year = ? AND cancel_no <> 'Y' OR cancel_no IS NULL)wlx01 on wlx01.hsien_id=cd01.hsien_id ");
			sql.append("    left join (select * from bn01 where  m_year=? and bn_type <> '2')bn01 on wlx01.bank_no=bn01.bank_no and bn01.bank_type= ? ");
			paramList.add(wlx01_m_year);
			paramList.add(wlx01_m_year);
			paramList.add(bank_type) ;
			sql.append("    left join ( ");
			sql.append("               select wlx02.bank_no,");//金融機構代碼
			sql.append("                      wlx02count,");//國內營業分支機構家數
			sql.append("                      total_hour,");  
			sql.append("                      person_mm,");//本月人次
			sql.append("                      hour_mm,");//本月時數 
			sql.append("                      person_yy,");//本年人次
			sql.append("                      hour_yy,");//本年時數
			sql.append("                      decode(hour_yy,null,0,hour_yy) - total_hour AS less_hour");
			sql.append("               from "); 								   
			sql.append("                  (select bn01.bank_no,decode(wlx02count,null,0,wlx02count) as wlx02count,");
			sql.append("                          decode(total_hour,null,16,total_hour) as total_hour");
			sql.append("                   from (select * from bn01 where m_year=? and bn_type <> '2' and bank_type in ('6','7'))bn01");
			paramList.add(wlx01_m_year);
			sql.append("                       ,(");
			sql.append("                         SELECT   tbank_no,");
			sql.append("                                  COUNT(*) AS wlx02count,");
			sql.append("                                  (COUNT (*) + 1) * 16 AS total_hour");
			sql.append("                         FROM wlx02");
			sql.append("                         WHERE  m_year = ? AND cancel_no <> 'Y' OR cancel_no IS NULL");
			paramList.add(wlx01_m_year);
			sql.append("				 		 GROUP BY tbank_no)wlx02");
			sql.append("				   where bn01.bank_no = wlx02.tbank_no(+)");
			sql.append("                   )wlx02 ,");
			sql.append("                  (SELECT yeardata.tbank_no,person_mm, hour_mm, person_yy, hour_yy");                                   
			sql.append("				   FROM(SELECT wlx_trainning.m_year,");
			sql.append("				  		       wlx_trainning.tbank_no,");
			sql.append("						       bank_name, COUNT (*) AS person_yy,");//本年人次
			sql.append("						       SUM (course_hour) AS hour_yy");//本年時數 
			sql.append("					    FROM wlx_trainning"); 
			sql.append("						LEFT JOIN (SELECT * FROM bn01 WHERE m_year = ? and bn_type <> '2') bn01 ON wlx_trainning.tbank_no = bn01.bank_no"); 
			sql.append("					    WHERE wlx_trainning.m_year = ?  and wlx_trainning.m_month <= ?");           
			sql.append("				        GROUP BY wlx_trainning.m_year,wlx_trainning.tbank_no,bank_name");
			sql.append("				  )yeardata"); 
			paramList.add(wlx01_m_year);
			paramList.add(s_year);
			paramList.add(s_month) ;
			sql.append("				  left join"); 
			sql.append("				 (SELECT wlx_trainning.m_year,");
			sql.append("					     wlx_trainning.m_month,");
			sql.append("					     wlx_trainning.tbank_no,");
			sql.append("					     bank_name, COUNT (*) AS person_mm,");//本月人次
			sql.append("					     SUM (course_hour) AS hour_mm");//本月時數
			sql.append("				  FROM wlx_trainning LEFT JOIN (SELECT * FROM bn01 WHERE m_year = ? and bn_type <> '2') bn01 ON wlx_trainning.tbank_no = bn01.bank_no");
			sql.append("    			  WHERE wlx_trainning.m_year = ? AND wlx_trainning.m_month = ?");            
			sql.append("				  GROUP BY wlx_trainning.m_year,wlx_trainning.m_month,wlx_trainning.tbank_no,bank_name");
			sql.append("				 ) monthdata on  monthdata.tbank_no = yeardata.tbank_no");
			paramList.add(wlx01_m_year);
			paramList.add(s_year);
			paramList.add(s_month) ;
			sql.append("				)wlx_trainning");								                                 
			sql.append("			where wlx02.bank_no = wlx_trainning.tbank_no(+)");                               
			sql.append("		    order by wlx02.bank_no");			
			sql.append("           ) wlx_trainning  on wlx_trainning.bank_no=bn01.bank_no ");
			sql.append("    ) Temp1 where bank_name is not null ");
			sql.append(" order by  Temp1.FR001W_output_order, Temp1.bank_no ");
			sql.append(" )");
			List dbData_sum = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "bank_name,wlx02count,credit_staff_num,person_mm,hour_mm,person_yy,hour_yy");
			System.out.println("dbData_sum.size()="+dbData_sum.size());	
			
			int rowNum = 5;
			//add 處理沒有資料 by 2495
			if(dbData.size() == 0){
				row = sheet.getRow(1);
				cell = row.getCell((short)0);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellValue("民國" + s_year + "年" + s_month + "月無資料存在");				
			} else {
				row = sheet.getRow(1);
				cell = row.getCell((short)0);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellValue("民國" + s_year + "年" + s_month + "月");			

				
				String sValue = "";
				DataObject bean = null;
				List dbData_count = null;
				for(int tablecount=0;tablecount <2;tablecount++){
					if(tablecount == 0) dbData_count = dbData;//列印明細
					if(tablecount == 1) dbData_count = dbData_sum;//列印合計
				    for (int rowcount = 0; rowcount < dbData_count.size(); rowcount++) {
				    	bean = (DataObject) dbData_count.get(rowcount);
				    	for (int cellcount = 0; cellcount < 8; cellcount++) {
				    		if (cellcount == 0)
				    			sValue = (bean.getValue("bank_no") == null)?"":String.valueOf(bean.getValue("bank_no"));
				    		else if (cellcount == 1)
				    			sValue = (bean.getValue("bank_name") == null)?"":String.valueOf(bean.getValue("bank_name"));
				    		else if (cellcount == 2)
				    			sValue = (bean.getValue("wlx02count") == null)?"":bean.getValue("wlx02count").toString();
				    		else if (cellcount == 3)
				    			sValue = (bean.getValue("credit_staff_num") == null)?"":bean.getValue("credit_staff_num").toString();
				    		else if (cellcount == 4)
				    			sValue = (bean.getValue("person_mm") == null)?"":bean.getValue("person_mm").toString();
				    		else if (cellcount == 5)
				    			sValue = (bean.getValue("hour_mm") == null)?"":bean.getValue("hour_mm").toString();
				    		else if (cellcount == 6)
				    			sValue = (bean.getValue("person_yy") == null)?"":bean.getValue("person_yy").toString();
				    		else if (cellcount == 7)
				    			sValue = (bean.getValue("hour_yy") == null)?"":bean.getValue("hour_yy").toString();
                    
				    		row = sheet.createRow(rowNum);
				    		cell = row.createCell((short)cellcount);
				    		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				    		cell.setCellStyle(cs_center);
				    		 if (cellcount == 0 || cellcount == 1){							
				    			cell.setCellValue(sValue);
				    		}else{
				    			if(!sValue.equals("")){
				    			   cell.setCellValue(Double.parseDouble(sValue));
				    			}
				    		}
				    		
				    		int countline = rowcount;
				    		countline++;
				    	} // end of cellcount
		    	    	rowNum++;
				    } // end of rowcount
				}//end of tablecout
			}

			row = sheet.createRow(rowNum + 1);
			cell = row.createCell((short)0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cs1.setAlignment(HSSFCellStyle.ALIGN_LEFT);
			cell.setCellStyle(cs1);
			cell.setCellValue("資料來源：根據各農漁會信用部於農業金融局網際網路申報系統上傳資料彙編。");

			HSSFFooter footer = sheet.getFooter();
			footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
			FileOutputStream fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + openfile);
			wb.write(fout);
			// 儲存
			fout.close();
			System.out.println("儲存完成");
		} catch(Exception e) {
			System.out.println("createRpt Error:"+e+e.getMessage());
		}
		return errMsg;
	}
}