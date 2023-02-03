/*
106.05.23 create 農漁會信用部超逾法定比率一覽表 by George
108.09.17 fix 108年10月以後為購置住宅放款及房屋修繕放款餘額占存款總餘額大於55%
              108年9月(含)以前為自用住宅放款總額占定期性存款總額大於50% by 2295
110.03.04 add 110年4月(含)以後套用新格式->改為110月5月
                         贊助會員授信總餘額占贊助會員存款總餘額 > 100% -->移除
                         贊助會員授信總餘額占贊助會員存款總餘額 > 核定限額 -->新增
                         非會員授信總餘額占非會員存款總餘額  > 100% -->移除
                         非會員授信總餘額占非會員存款總餘額    > 核定限額 -->新增
                         非會員擔保品種類-->新增
                         擔保品坐落地選擇毗鄰之直轄市或縣市，而財務指標未符規定-->新增
 110.07.08 fix field_990710/990720取值 by 2295
 110.09.15 add 非由政府發行之債券及票券餘額占存款總額大於15%
                           單一銀行所發行之金融債券及可轉讓定期存單之原始取得成本總餘額，占前一年度信用部決算淨值大於15%
                           單一企業所發行之短期票券及公司債之原始取得成本總餘額，占前一年度信用部決算淨值大於10%
           fix 外幣資產與外幣負債差額絕對值逾新台幣100萬元且占信用部上年度決算淨值大於10%,原5%調整為10%   by 2295
 111.06.07 fix 非會員擔保品種類-->移除
                           擔保品坐落地選擇毗鄰之直轄市或縣市，而財務指標未符規定-->移除
                           非由政府發行之債券及票券餘額占存款總額大於15%-->調整公式為(A02)990860÷(A01)220000>15%
                           單一銀行所發行之金融債券及可轉讓定期存單之原始取得成本總餘額，占前一年度信用部決算淨值大於15%-->調整公式為(A02)990870÷[(A02)990230-(A02)990240]-(A99)992810]>15%
                           單一企業所發行之短期票券及公司債之原始取得成本總餘額，占前一年度信用部決算淨值大於10%-->調整公式為(A02)990880÷[(A02)990230-(A02)990240]-(A99)992810]>10%       
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

public class RptFR078W {

	public static String createRpt(String s_year, String s_month, String bank_type)	{

		String errMsg = "";
		String m_year = (Integer.parseInt(s_year) < 100)?"99":"100"; 
		
		
		
		reportUtil reportUtil = new reportUtil();
		StringBuffer sql = new StringBuffer () ;
		ArrayList paramList = new ArrayList() ;
		
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
			
			
			String openfile = "";
			if((Integer.parseInt(s_year) * 100 + Integer.parseInt(s_month)) >= 10810 
			&& (Integer.parseInt(s_year) * 100 + Integer.parseInt(s_month)) <= 11003
			){
			     openfile = "超逾法定比率一覽表_10810.xls";
		    }else if((Integer.parseInt(s_year) * 100 + Integer.parseInt(s_month)) >= 11005){
		    	 openfile = "超逾法定比率一覽表_11004.xls";
		    }else{
		    	 openfile = "超逾法定比率一覽表.xls";
		    }
		    	
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
			ps.setScale((short)(((Integer.parseInt(s_year) * 100 + Integer.parseInt(s_month)) >= 11005)?55:70)); //列印縮放百分比			
			ps.setPaperSize((short)9); //設定紙張大小 A4
			ps.setLandscape(true);//設定橫印

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
			
			sql.setLength(0) ;
			
			//資料SQL
			sql.append("select bank_no,bank_name, ");//農(漁)會名稱
			sql.append("sum(field_month_dc_rate) as field_month_dc_rate, ");//存放比率大於80%
			sql.append("sum(field_990210_990230) as field_990210_990230, ");//內部融資餘餘額占上年度信用部決算淨值大於60%
			sql.append("sum(field_990410_990420_100) as field_990410_990420_100, ");//贊助會員授信總餘額占贊助會員存款總餘額 > 100%
			sql.append("sum(field_990410_990420_150) as field_990410_990420_150, ");//贊助會員授信總餘額占贊助會員存款總餘額 > 150%
			sql.append("sum(field_990410_990420_200) as field_990410_990420_200, ");//贊助會員授信總餘額占贊助會員存款總餘額 > 200%
			sql.append("sum(field_990410_990420_over_200) as field_990410_990420_over_200,");//贊助會員授信總餘額占贊助會員存款總餘額 > 核定限額 ->110.03.04 add
			sql.append("sum(field_990512_990320) as field_990512_990320, ");//非會員無擔保消費性貸款總額占農(漁)會上年度全體決算淨值大於100%
			sql.append("sum(field_k_990620_100) as field_k_990620_100, ");//非會員授信總餘額占非會員存款總餘額  > 100%
			sql.append("sum(field_k_990620_150) as field_k_990620_150, ");//非會員授信總餘額占非會員存款總餘額  > 150%
			sql.append("sum(field_k_990620_200) as field_k_990620_200, ");//非會員授信總餘額占非會員存款總餘額  > 200%
			sql.append("sum(field_k_990620_over_200) as field_k_990620_over_200,");//非會員授信總餘額占非會員存款總餘額    > 核定限額 ->110.03.04 add
			//111.06.07 取消顯示非會員擔保品種類
			//sql.append("sum(field_990624_1) as field_990624_1,");//非會員擔保品種類-勾選土地或建築物等不動產、動產，而逾放比率大於等於5% ->110.03.04 add
			//sql.append("sum(field_990624_2) as field_990624_2,");//非會員擔保品種類-勾選住宅、已取得建築執照或雜項執照之建築基地，而逾放比率大於等於10% ->110.03.04 add
			//111.06.07 取消顯示擔保品坐落地選擇毗鄰之直轄市或縣市，而財務指標未符規定
		    //sql.append("sum(field_990626_1) as field_990626_1,");//擔保品坐落地選擇毗鄰之直轄市或縣市，而財務指標未符規定-逾放比率大於等於1% ->110.03.04 add
		    //sql.append("sum(field_990626_2) as field_990626_2,");//擔保品坐落地選擇毗鄰之直轄市或縣市，而財務指標未符規定-放款覆蓋率小於等於2% ->110.03.04 add
		    //sql.append("sum(field_990626_3) as field_990626_3,");//擔保品坐落地選擇毗鄰之直轄市或縣市，而財務指標未符規定-BIS小於等於10% ->110.03.04 add
			sql.append("sum(field_990710_990720) as field_990710_990720, ");//自用住宅放款總額占定期性存款總額大於50%[108年9月(含)以前使用]		
			sql.append("sum(field_990810_990230) as field_990810_990230, ");//固定資產淨額占農(漁)會信用部上年度信用部決算淨值大於100%
			sql.append("sum(field_990910_990230) as field_990910_990230, ");//外幣資產與外幣負債差額絕對值逾新台幣100萬元且占信用部上年度決算淨值大於10% //110.09.15 fix
			sql.append("sum(field_991020_990320) as field_991020_990320, ");//理監事職員及利害關係人擔保授信餘額占農(漁)會上年度全體決算淨值超過150%
			sql.append("sum(field_996114_990230) as field_996114_990230, ");//對鄉(鎮、市)公所授信未經其所隸屬之縣政府保證,及對直轄市、縣(市)政府投資經營之公營事業,其授信經該直轄市、縣(市)政府保證,兩者合計超
			sql.append("sum(field_captial_rate) as field_captial_rate, ");//合格淨值占風險性資產總額小於8%
			sql.append("sum(field_990860_220000) as field_990860_220000, ");//非由政府發行之債券及票券餘額占存款總額大於15%->111.06.07 fix
			sql.append("sum(field_990870_990230) as field_990870_990230, ");//單一銀行所發行之金融債券及可轉讓定期存單之原始取得成本總餘額，占前一年度信用部決算淨值大於15%->111.06.07 fix 
			sql.append("sum(field_990880_990230) as field_990880_990230 ");//單一企業所發行之短期票券及公司債之原始取得成本總餘額，占前一年度信用部決算淨值大於10%->111.06.07 fix

			sql.append("from ( ");
			sql.append("select bank_no,bank_name, ");//信用部名稱
			sql.append("decode(acc_code,'field_month_dc_rate',amt,0) as field_month_dc_rate, ");
			sql.append("decode(acc_code,'field_990210/(990230-990240)',amt,0) as field_990210_990230, ");
			sql.append("decode(acc_code,'field_990410/990420',amt,0) as field_990410_990420, ");
			sql.append("CASE ");
			sql.append("  WHEN acc_code='field_990410/990420' and range = '100' THEN amt ");
			sql.append("END as field_990410_990420_100, ");
			sql.append("CASE ");
			sql.append("  WHEN acc_code='field_990410/990420' and range = '150' THEN amt ");
			sql.append("END as field_990410_990420_150, ");
			sql.append("CASE ");
			sql.append("  WHEN acc_code='field_990410/990420' and range = '200' THEN amt ");
			sql.append("END as field_990410_990420_200, ");
			
			sql.append("CASE ");   
		    sql.append("  WHEN acc_code='field_990410/990420' and (range <> '150' and range <> '200')  THEN amt ");   
			sql.append("END as field_990410_990420_over_200, ");
			sql.append("decode(acc_code,'field_990512/990320',amt,0) as field_990512_990320, ");
			sql.append("decode(acc_code,'field_k/990620',amt,0) as field_k_990620, ");
			sql.append("CASE ");
			sql.append("  WHEN acc_code='field_k/990620' and range = '100' THEN amt ");
			sql.append("END as field_k_990620_100, ");
			sql.append("CASE ");
			sql.append("  WHEN acc_code='field_k/990620' and range = '150' THEN amt ");
			sql.append("END as field_k_990620_150, ");
			sql.append("CASE ");
			sql.append("  WHEN acc_code='field_k/990620' and range = '200' THEN amt ");
			sql.append("END as field_k_990620_200, ");
			sql.append("CASE ");   
			sql.append("  WHEN acc_code='field_k/990620'  and (range <> '150' and range <> '200') THEN amt ");   
			sql.append("END as field_k_990620_over_200, "); 
			//111.06.07 取消顯示 非會員擔保品種類/擔保品坐落地選擇毗鄰之直轄市或縣市，而財務指標未符規定
			//sql.append("decode(acc_code,'990624_1_rule',amt,0) as field_990624_1, ");   
			//sql.append("decode(acc_code,'990624_2_rule',amt,0) as field_990624_2, ");  
			//sql.append("decode(acc_code,'990626_1_rule',amt,0) as field_990626_1, ");   
			//sql.append("decode(acc_code,'990626_2_rule',amt,0) as field_990626_2, ");  
			//sql.append("decode(acc_code,'990626_3_rule',amt,0) as field_990626_3, ");
			//sql.append("decode(acc_code,'field_990710/990720',amt,0) as field_990710_990720, ");
			sql.append("CASE WHEN (a02.m_year * 100 + a02.m_month >= 10810) THEN decode(acc_code,'field_990711_990712/fieldi_y',amt,0)");
			sql.append(" ELSE decode(acc_code,'field_990710/990720',amt,0) END as field_990710_990720,");//108.09.17 add
			sql.append("decode(acc_code,'field_990810/(990230-990240)',amt,0) as field_990810_990230, ");
			sql.append("decode(acc_code,'field_|990910-990920|/990230',amt,0) as field_990910_990230, ");
			sql.append("decode(acc_code,'field_991020/990320',amt,0) as field_991020_990320, ");
			sql.append("decode(acc_code,'field_996114_996115/(990230-990240)',amt,0) as field_996114_990230,0 as field_captial_rate, ");
			sql.append("decode(acc_code,'field_990860/220000',amt,0) as field_990860_220000, ");//111.06.07 fix 
			sql.append("decode(acc_code,'field_990870/(990230-990240-992810)',amt,0) as field_990870_990230, ");//111.06.07 fix
			sql.append("decode(acc_code,'field_990880/(990230-990240-992810)',amt,0) as field_990880_990230 ");//111.06.07 fix 

			sql.append("from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
			paramList.add(m_year) ;
			sql.append("left join (select * from a02_operation where m_year=? and m_month=? )a02 on  bn01.bank_no = a02.bank_code ");
			paramList.add(s_year) ;
			paramList.add(s_month) ;
			if((Integer.parseInt(s_year) * 100 + Integer.parseInt(s_month)) >= 10810){
			//108.10(含)以後使用
			sql.append("where acc_code in ('field_month_dc_rate','field_990210/(990230-990240)','field_990410/990420','field_990512/990320','field_k/990620','field_990711_990712/fieldi_y','field_990810/(990230-990240)','field_|990910-990920|/990230','field_991020/990320','field_996114_996115/(990230-990240)','field_990860/220000','field_990870/(990230-990240-992810)','field_990880/(990230-990240-992810)')");                                                                               
			}else{
			//108.09(含)以前使用
			sql.append("where acc_code in ('field_month_dc_rate','field_990210/(990230-990240)','field_990410/990420','field_990512/990320','field_k/990620','field_990710/990720','field_990810/(990230-990240)','field_|990910-990920|/990230','field_991020/990320','field_996114_996115/(990230-990240)') ");
			}
			sql.append("and violate='Y' ");//有違反
			sql.append("union ");
			sql.append("select bank_no,bank_name, ");//信用部名稱        
			sql.append("0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0, ");
			sql.append("amt as  field_captial_rate,0,0,0 ");//BIS%
			sql.append("from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
			paramList.add(m_year) ;
			sql.append("left join (select * from a01_operation where m_year=? and m_month=? and bank_type in ('6','7') and bank_code !='ALL')a01 on  bn01.bank_no = a01.bank_code ");
			paramList.add(s_year) ;
			paramList.add(s_month) ;
			sql.append("where acc_code='field_captial_rate' "); //資本適足率
			sql.append("and amt < 8  ");
			sql.append(")a    ");
			sql.append("group by bank_no,bank_name ");
			sql.append("order by bank_no ");
			
			List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "bank_no,field_month_dc_rate,field_990210_990230,field_990410_990420_100,field_990410_990420_150,field_990410_990420_200,field_990512_990320,field_k_990620_100,field_k_990620_150,field_k_990620_200,field_990710_990720,field_990810_990230,field_990910_990230,field_991020_990320,field_996114_990230,field_captial_rate,field_990711_990712_y,field_990410_990420_over_200,field_k_990620_over_200,field_990860_220000,field_990870_990230,field_990880_990230");
			System.out.println("dbData.size()="+dbData.size());
			
			//開始寫入Excel的內容
			int rowNum = 4;
			
			row = sheet.getRow(0);
			cell = row.getCell((short) 0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("農漁會信用部"+s_year+"年"+s_month+"月底超逾法定比率一覽表");
			
			row = sheet.getRow(1);
			cell = row.getCell((short) 0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("列印日期：民國" + (Integer.parseInt(Utility.getDateFormat("yyyy"))-1911)+ "年" + Utility.getDateFormat("MM")  + "月" + Utility.getDateFormat("dd")  + "日");

			String sValue = "";
			DataObject bean = null;
			
			// 列印明細/合計
			for (int tablecount = 0; tablecount < 2; tablecount++) {
				if (tablecount == 0)
				for (int rowcount = 0; rowcount < dbData.size(); rowcount++) {
					bean = (DataObject) dbData.get(rowcount);
					for (int cellcount = 0; cellcount < (((Integer.parseInt(s_year) * 100 + Integer.parseInt(s_month)) >= 11005)?19:16); cellcount++) {
						if((Integer.parseInt(s_year) * 100 + Integer.parseInt(s_month)) >= 11005){
							if (cellcount == 0)
								sValue = (bean.getValue("bank_name") == null) ? "" : String.valueOf(bean.getValue("bank_name"));
							else if (cellcount == 1)
								sValue = (bean.getValue("field_month_dc_rate") == null) ? "" : String.valueOf(bean.getValue("field_month_dc_rate"));
							else if (cellcount == 2)
								sValue = (bean.getValue("field_990210_990230") == null) ? "" : String.valueOf(bean.getValue("field_990210_990230"));
							else if (cellcount == 3)
								sValue = (bean.getValue("field_990410_990420_150") == null) ? "" : String.valueOf(bean.getValue("field_990410_990420_150"));
							else if (cellcount == 4)
								sValue = (bean.getValue("field_990410_990420_200") == null) ? "" : String.valueOf(bean.getValue("field_990410_990420_200"));
							else if (cellcount == 5)
								sValue = (bean.getValue("field_990410_990420_over_200") == null) ? "" : String.valueOf(bean.getValue("field_990410_990420_over_200"));
							else if (cellcount == 6)
								sValue = (bean.getValue("field_k_990620_150") == null) ? "" : String.valueOf(bean.getValue("field_k_990620_150"));
							else if (cellcount == 7)
								sValue = (bean.getValue("field_k_990620_200") == null) ? "" : String.valueOf(bean.getValue("field_k_990620_200"));
							else if (cellcount == 8)
								sValue = (bean.getValue("field_k_990620_over_200") == null) ? "" : String.valueOf(bean.getValue("field_k_990620_over_200"));
							/*111.06.07取消顯示							
							else if (cellcount == 9)
								sValue = (bean.getValue("field_990624_1") == null) ? "" : String.valueOf(bean.getValue("field_990624_1"));
							else if (cellcount == 10)
								sValue = (bean.getValue("field_990624_2") == null) ? "" : String.valueOf(bean.getValue("field_990624_2"));
							else if (cellcount == 11)
								sValue = (bean.getValue("field_990626_1") == null) ? "" : String.valueOf(bean.getValue("field_990626_1"));
							else if (cellcount == 12)
								sValue = (bean.getValue("field_990626_2") == null) ? "" : String.valueOf(bean.getValue("field_990626_2"));
							else if (cellcount == 13)
								sValue = (bean.getValue("field_990626_3") == null) ? "" : String.valueOf(bean.getValue("field_990626_3"));
							*/
							else if (cellcount == 9)
								sValue = (bean.getValue("field_990512_990320") == null) ? "" : String.valueOf(bean.getValue("field_990512_990320"));							
							else if (cellcount == 10)							
								sValue = (bean.getValue("field_990710_990720") == null) ? "" : String.valueOf(bean.getValue("field_990710_990720"));
							else if (cellcount == 11)
								sValue = (bean.getValue("field_990810_990230") == null) ? "" : String.valueOf(bean.getValue("field_990810_990230"));
							else if (cellcount == 12)
								sValue = (bean.getValue("field_990860_220000") == null) ? "" : String.valueOf(bean.getValue("field_990860_220000"));//111.06.07 fix
							else if (cellcount == 13)
								sValue = (bean.getValue("field_990870_990230") == null) ? "" : String.valueOf(bean.getValue("field_990870_990230"));//111.06.07 fix
							else if (cellcount == 14)
								sValue = (bean.getValue("field_990880_990230") == null) ? "" : String.valueOf(bean.getValue("field_990880_990230"));//111.06.07 fix							
							else if (cellcount == 15)
								sValue = (bean.getValue("field_990910_990230") == null) ? "" : String.valueOf(bean.getValue("field_990910_990230"));
							else if (cellcount == 16)
								sValue = (bean.getValue("field_991020_990320") == null) ? "" : String.valueOf(bean.getValue("field_991020_990320"));
							else if (cellcount == 17)
								sValue = (bean.getValue("field_996114_990230") == null) ? "" : String.valueOf(bean.getValue("field_996114_990230"));
							else if (cellcount == 18)
								sValue = (bean.getValue("field_captial_rate") == null) ? "" : String.valueOf(bean.getValue("field_captial_rate"));		
						}else{	
						if (cellcount == 0)
							sValue = (bean.getValue("bank_name") == null) ? "" : String.valueOf(bean.getValue("bank_name"));
						else if (cellcount == 1)
							sValue = (bean.getValue("field_month_dc_rate") == null) ? "" : String.valueOf(bean.getValue("field_month_dc_rate"));
						else if (cellcount == 2)
							sValue = (bean.getValue("field_990210_990230") == null) ? "" : String.valueOf(bean.getValue("field_990210_990230"));
						else if (cellcount == 3)
							sValue = (bean.getValue("field_990410_990420_100") == null) ? "" : String.valueOf(bean.getValue("field_990410_990420_100"));
						else if (cellcount == 4)
							sValue = (bean.getValue("field_990410_990420_150") == null) ? "" : String.valueOf(bean.getValue("field_990410_990420_150"));
						else if (cellcount == 5)
							sValue = (bean.getValue("field_990410_990420_200") == null) ? "" : String.valueOf(bean.getValue("field_990410_990420_200"));
						else if (cellcount == 6)
							sValue = (bean.getValue("field_990512_990320") == null) ? "" : String.valueOf(bean.getValue("field_990512_990320"));
						else if (cellcount == 7)
							sValue = (bean.getValue("field_k_990620_100") == null) ? "" : String.valueOf(bean.getValue("field_k_990620_100"));
						else if (cellcount == 8)
							sValue = (bean.getValue("field_k_990620_150") == null) ? "" : String.valueOf(bean.getValue("field_k_990620_150"));
						else if (cellcount == 9)
							sValue = (bean.getValue("field_k_990620_200") == null) ? "" : String.valueOf(bean.getValue("field_k_990620_200"));
						else if (cellcount == 10)							
							sValue = (bean.getValue("field_990710_990720") == null) ? "" : String.valueOf(bean.getValue("field_990710_990720"));
						else if (cellcount == 11)
							sValue = (bean.getValue("field_990810_990230") == null) ? "" : String.valueOf(bean.getValue("field_990810_990230"));
						else if (cellcount == 12)
							sValue = (bean.getValue("field_990910_990230") == null) ? "" : String.valueOf(bean.getValue("field_990910_990230"));
						else if (cellcount == 13)
							sValue = (bean.getValue("field_991020_990320") == null) ? "" : String.valueOf(bean.getValue("field_991020_990320"));
						else if (cellcount == 14)
							sValue = (bean.getValue("field_996114_990230") == null) ? "" : String.valueOf(bean.getValue("field_996114_990230"));
						else if (cellcount == 15)
							sValue = (bean.getValue("field_captial_rate") == null) ? "" : String.valueOf(bean.getValue("field_captial_rate"));						
						}
					    
						if(sValue.equals("0")) 
							sValue = "";
						
						row = sheet.createRow(rowNum);
						cell = row.createCell((short) cellcount);
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell.setCellStyle(cs_center);
						cell.setCellValue(sValue);
						
					} // end of cellcount
					rowNum++;
				} // end of rowcount
			} // end of tablecout

			HSSFFooter footer = sheet.getFooter();
			footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
			FileOutputStream fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + "超逾法定比率一覽表.xls");
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