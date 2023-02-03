/*
 94.11.21 add 明細表 by 4180
 94.12.21 fix 明細表 by 2386
 96.10.03 add 增加牌告利率欄位(定期存款-二年/三年,定期儲蓄存款-一年/二年/三年,活期存款機動利率,活期儲蓄存款機動利率) by 2295
 96.11.28 fix 若利率為0時,不算入最低利率.也不算入平均值 by 2295 
 98.03.31 add 基準利率-指標利率(月調)/基準利率(月調)/指數型房貸指標利率(月調) by 2295
 99.11.03 fix 根據查詢年度.100年以後取得新縣市別.100年以前取得舊縣市別 
 		       使用PreparedStatement;並列印轉換後的SQL;套用QueryDB_SQLParam by 2295
100.08.26 fix 顯示小數點至第3位,不足者補0 by 2295
101.07.26 fix 查詢頁修改+機構中英文 by 2968
101.11.26 fix 調整報表格式,當利率為0.000時,顯示成"-" by 2295 
101.12.21 add 說明文字.機構名稱靠左 by 2295
*/
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import java.io.*;
import java.text.DecimalFormat;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class FR027W_Excel {
  	public static String createRpt(String s_year,String s_quarter,String bank_type,String showEng)	{

		String errMsg = "";
		DataObject bean = null;
		String bank_type_name=(bank_type.equals("6"))?"農會":"漁會";
		//99.11.03 add 查詢年度100年以前.縣市別不同===============================
	    String cd01_table = (Integer.parseInt(s_year) < 100)?"cd01_99":""; 
	    String wlx01_m_year = (Integer.parseInt(s_year) < 100)?"99":"100"; 
	    //=====================================================================  
	    List paramList = new ArrayList();
		String filename="";
		filename=(bank_type.equals("6"))?"全體農會信用部牌告利率彙總表(季報)":"全體漁會信用部牌告利率彙總表(季報)";		
		reportUtil reportUtil = new reportUtil();
		DecimalFormat df_md = new DecimalFormat("############0.000");//顯示小數點至第3位,不足者補0
    	  
		try{				
			File xlsDir = new File(Utility.getProperties("xlsDir"));        
        	File reportDir = new File(Utility.getProperties("reportDir"));       
	    	        
    		if(!xlsDir.exists()){
     			if(!Utility.mkdirs(Utility.getProperties("xlsDir"))){
     		   		errMsg +=Utility.getProperties("xlsDir")+"目錄新增失敗";
     			}    
    		}
    		if(!reportDir.exists()){
     			if(!Utility.mkdirs(Utility.getProperties("reportDir"))){
     		   		errMsg +=Utility.getProperties("reportDir")+"目錄新增失敗";
     			}    
    		}
    		//String openfile="農漁會信用部牌告利率彙總表(季報).xls";
    		String openfile=(bank_type.equals("6"))?"農會信用部牌告利率彙總表(季報).xls":"漁會信用部牌告利率彙總表(季報).xls";	
    		System.out.println("open file "+openfile);
			FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+ openfile );			
			
	  	    //設定FileINputStream讀取Excel檔
	  		POIFSFileSystem fs = new POIFSFileSystem( finput );
	  		if(fs==null){System.out.println("open 範本檔失敗");} else System.out.println("open 範本檔成功");
	  		HSSFWorkbook wb = new HSSFWorkbook(fs);
	  		if(wb==null){System.out.println("open工作表失敗");}else System.out.println("open 工作表 成功");
	  		HSSFSheet sheet = wb.getSheetAt(0);//讀取第一個工作表，宣告其為sheet 
	  		if(sheet==null){System.out.println("open sheet 失敗");}else System.out.println("open sheet 成功");
	  		HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
	        //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
	        //sheet.setAutobreaks(true); //自動分頁
			
	        //設定頁面符合列印大小
	        sheet.setAutobreaks( false );
	        ps.setScale( ( short )90 ); //列印縮放百分比
	        ps.setPaperSize( ( short )9 ); //設定紙張大小 A4

	  		//設定表頭 為固定 先設欄的起始再設列的起始
	        wb.setRepeatingRowsAndColumns(0, 1, 16, 2, 4);	  		
	  		finput.close();
	  		
	  		HSSFRow row=null;//宣告一列 
	  		HSSFCell cell=null;//宣告一個儲存格
			HSSFFont ft = wb.createFont();
			HSSFCellStyle cs = wb.createCellStyle();		   
		    ft.setFontHeightInPoints((short)20);
		    ft.setFontName("標楷體");
		    cs.setFont(ft);
		    cs.setAlignment(HSSFCellStyle.ALIGN_CENTER);
			cs.setBorderTop((short)0);
            cs.setBorderLeft((short)0);
            cs.setBorderLeft((short)0);
            cs.setBorderBottom((short)0);
			/*
			row = sheet.getRow(0);
      		cell = row.getCell( (short) 0);
      		cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
            
		    ft.setFontHeightInPoints((short)20);
		    ft.setFontName("標楷體");
		    cs.setFont(ft);
		    cs.setAlignment(HSSFCellStyle.ALIGN_CENTER);
            cell.setCellStyle(cs); 	
            */
            //有框內文置中
            HSSFCellStyle cellStyle = wb.createCellStyle();
        	cellStyle = HssfStyle.setStyle( cellStyle, wb.createFont(),
                    					   new String[] {
                    					   "BORDER", "PHC", "PVC", "F08",
                    				  	   "WRAP"} );        	
        	//有框內文置左
        	HSSFCellStyle leftStyle = wb.createCellStyle(); 
            leftStyle = HssfStyle.setStyle( leftStyle, wb.createFont(),
                                    new String[] {
                                    "BORDER", "PHL", "PVC", "F08",
                                    "WRAP"} );
        	StringBuffer sqlCmd = new StringBuffer(); 
        	sqlCmd.append(" select * ");
        	sqlCmd.append(" from (  ");
        	sqlCmd.append("	    select nvl(cd01.hsien_id,' ')       as  hsien_id ,");
       		sqlCmd.append("			   nvl(cd01.hsien_name,'其他')  as  hsien_name,");
       		sqlCmd.append("			   cd01.FR001W_output_order     as  FR001W_output_order,");
       		sqlCmd.append("			   bn01.BANK_NAME,");
       		sqlCmd.append("            wlx01.english,");
    		sqlCmd.append("			   WLX_S_RATE.*");
    		sqlCmd.append("     from (select * from "+cd01_table+" cd01 where cd01.hsien_id <> 'Y' ) cd01 ");
			sqlCmd.append("   	left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id          ");
			sqlCmd.append("   	left join (select * from bn01 where m_year=?)bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=? ");
    		sqlCmd.append("   	left join (select * from WLX_S_RATE  where m_year=? and M_Quarter =?) WLX_S_RATE  ");
			sqlCmd.append("               on WLX_S_RATE.BANK_NO=bn01.bank_no  ");
			sqlCmd.append("	    ) Temp1 where Temp1.m_year > 0 ");
			sqlCmd.append(" order by  Temp1.FR001W_output_order, Temp1.bank_no");
			paramList.add(wlx01_m_year);
			paramList.add(wlx01_m_year);
			paramList.add(bank_type);
			paramList.add(s_year);
			paramList.add(s_quarter);
			List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"bank_name,english,period_1_fix_rate,period_1_var_rate,"+
												   "period_3_fix_rate,period_3_var_rate,"+
 												   "period_6_fix_rate,period_6_var_rate, "+
 												   "period_9_fix_rate,period_9_var_rate,"+
 												   "period_12_fix_rate,period_12_var_rate,"+
 												   "period_24_fix_rate,period_24_var_rate,"+
                                                   "period_36_fix_rate,period_36_var_rate,"+
                                                   "deposit_12_fix_rate,deposit_12_var_rate,"+
                                                   "deposit_24_fix_rate,deposit_24_var_rate,"+
                                                   "deposit_36_fix_rate,deposit_36_var_rate,"+
                                                   "deposit_var_rate,save_var_rate,"+
												   "basic_pay_var_rate,period_house_var_rate,"+
 												   "base_mark_rate,base_fix_rate,"+ "base_base_rate,"+
												   "base_mark_rate_month,base_base_rate_month,period_house_var_rate_month");	
			
			//add 處理沒有資料 by 2495
		    if(dbData.size()==0){
		    	row=sheet.getRow(1);
		    	cell=row.getCell((short)0);	       	
		    	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		    	//96.11.05 96/10月以前為季報.96/10月以後為月報================================
		    	if((Integer.parseInt(s_year) * 100 + Integer.parseInt(s_quarter)) < 9610){
   	       	        cell.setCellValue(s_year +"年 第 " +s_quarter +" 季無資料存在");
		    	}else{
		    		cell.setCellValue(s_year +"年" +s_quarter +"月無資料存在");
		    	}
   	       	    //ft.setFontHeightInPoints((short)12);
		    	//cs.setFont(ft);
		    	//cs.setAlignment(HSSFCellStyle.ALIGN_CENTER);		    
		    	//cell.setCellStyle(cs);
		    }else{		    	
		    	row= sheet.getRow(1);													
		    	cell=row.getCell((short)0);			
		    	cell.setEncoding(HSSFCell.ENCODING_UTF_16);		
		    	if((Integer.parseInt(s_year) * 100 + Integer.parseInt(s_quarter)) < 9610){
		    	   cell.setCellValue(" 中華民國"+s_year+"年 第 " + s_quarter + " 季");		
		    	}else{
		    	   cell.setCellValue(" 中華民國"+s_year+"年 " + s_quarter + " 月");
		    	}
		    	/*
		    	ft.setFontHeightInPoints((short)12);
		    	cs.setFont(ft);
		    	cs.setAlignment(HSSFCellStyle.ALIGN_CENTER);		    
		    	cell.setCellStyle(cs);
		    	*/
		    	System.out.println("明細表----------------------------");
		    	int rowNum=5;
		    	String  bank_name_Value="";
		    	Double insertValue=Double.valueOf("0");;
		    	
		    	
		    	for(int rowcount=0;rowcount<dbData.size();rowcount++){
		    		for(int cellcount=0;cellcount<31;cellcount++){	
		    			bean = (DataObject) dbData.get(rowcount);
		    			if(cellcount==0){
		    			    if("true".equals(showEng)){
		    			        bank_name_Value= String.valueOf(bean.getValue("bank_name")+"\n"+bean.getValue("english"));
		    			    }else{
		    			        bank_name_Value= String.valueOf(bean.getValue("bank_name"));
		    			    }
		    				
		    			}else if(cellcount==1)	
		    				insertValue= Double.valueOf(bean.getValue("deposit_var_rate").toString());//活期存款機動利率
		    			else if(cellcount==2)	
		    				insertValue= Double.valueOf(bean.getValue("save_var_rate").toString());//活期儲蓄存款機動利率
		    			else if(cellcount==3)	
		    				insertValue= Double.valueOf(bean.getValue("period_1_fix_rate").toString());//定期存款-一個月-固定
		    			else if(cellcount==4)	
		    				insertValue= Double.valueOf(bean.getValue("period_1_var_rate").toString());//定期存款-一個月-變動
		    			else if(cellcount==5)	
		    				insertValue= Double.valueOf(bean.getValue("period_3_fix_rate").toString());//定期存款-三個月-固定
		    			else if(cellcount==6)	
		    				insertValue= Double.valueOf(bean.getValue("period_3_var_rate").toString());//定期存款-三個月-變動
		    			else if(cellcount==7)	
		    				insertValue= Double.valueOf(bean.getValue("period_6_fix_rate").toString());//定期存款-六個月-固定
		    			else if(cellcount==8)	
		    				insertValue= Double.valueOf(bean.getValue("period_6_var_rate").toString());//定期存款-六個月-變動
		    			else if(cellcount==9)	
		    				insertValue= Double.valueOf(bean.getValue("period_9_fix_rate").toString());//定期存款-九個月-固定
		    			else if(cellcount==10)	
		    				insertValue= Double.valueOf(bean.getValue("period_9_var_rate").toString());//定期存款-九個月-變動
		    			else if(cellcount==11)	
		    				insertValue= Double.valueOf(bean.getValue("period_12_fix_rate").toString());//定期存款-一年-固定
		    			else if(cellcount==12)	
		    				insertValue= Double.valueOf(bean.getValue("period_12_var_rate").toString());//定期存款-一年-變動
		    			else if(cellcount==13)	
		    				insertValue= Double.valueOf(bean.getValue("period_24_fix_rate").toString());//定期存款-二年-固定
		    			else if(cellcount==14)	
		    				insertValue= Double.valueOf(bean.getValue("period_24_var_rate").toString());//定期存款-二年-變動		    			
		    			else if(cellcount==15)	
		    				insertValue= Double.valueOf(bean.getValue("period_36_fix_rate").toString());//定期存款-三年-固定
		    			else if(cellcount==16)	
		    				insertValue= Double.valueOf(bean.getValue("period_36_var_rate").toString());//定期存款-三年-變動
		    			else if(cellcount==17)	
		    				insertValue= Double.valueOf(bean.getValue("deposit_12_fix_rate").toString());//定期儲蓄存款-一年-固定
		    			else if(cellcount==18)	
		    				insertValue= Double.valueOf(bean.getValue("deposit_12_var_rate").toString());//定期儲蓄存款-一年-變動
		    			else if(cellcount==19)	
		    				insertValue= Double.valueOf(bean.getValue("deposit_24_fix_rate").toString());//定期儲蓄存款-二年-固定
		    			else if(cellcount==20)	
		    				insertValue= Double.valueOf(bean.getValue("deposit_24_var_rate").toString());//定期儲蓄存款-二年-變動		    			
		    			else if(cellcount==21)	
		    				insertValue= Double.valueOf(bean.getValue("deposit_36_fix_rate").toString());//定期儲蓄存款-三年-固定
		    			else if(cellcount==22)	
		    				insertValue= Double.valueOf(bean.getValue("deposit_36_var_rate").toString());//定期儲蓄存款-三年-變動		    			
		    			else if(cellcount==23)	
		    				insertValue= Double.valueOf(bean.getValue("basic_pay_var_rate").toString());//基本放款利率(機動)
		    			else if(cellcount==24)	
		    				insertValue= Double.valueOf(bean.getValue("period_house_var_rate").toString());//指數型房貸指標利率(季調)
		    			else if(cellcount==25)	
		    				insertValue= Double.valueOf(bean.getValue("period_house_var_rate_month") == null?"0":bean.getValue("period_house_var_rate_month").toString());//指數型房貸指標利率(月調)
		    			else if(cellcount==26)	
		    				insertValue= Double.valueOf(bean.getValue("base_mark_rate").toString());//基準利率-指標利率(1)(季調)
		    			else if(cellcount==27)	
		    				insertValue= Double.valueOf(bean.getValue("base_mark_rate_month") == null?"0":bean.getValue("base_mark_rate_month").toString());//基準利率-指標利率(1)(月調)
		    			else if(cellcount==28)	
		    				insertValue= Double.valueOf(bean.getValue("base_fix_rate").toString());//基準利率-一定比率(2)
		    			else if(cellcount==29)	
		    				insertValue= Double.valueOf(bean.getValue("base_base_rate").toString());//基準利率-基準利率 (3)=(1)+(2)(季調)
		    			else if(cellcount==30)	
		    				insertValue= Double.valueOf(bean.getValue("base_base_rate_month") == null?"0":bean.getValue("base_base_rate_month").toString());//基準利率-基準利率 (3)=(1)+(2)(月調)
		    			else insertValue=Double.valueOf("0");	
		    					    			
		    			row = sheet.createRow(rowNum);
		    			cell=row.createCell((short)cellcount);	
		    			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		    			cell.setCellStyle(cellStyle);
		    			
		    			if(cellcount==0){		    				
		    				cell.setCellValue(bank_name_Value);	
		    				cell.setCellStyle(leftStyle);//101.12.21 fix 單位名稱改置左 by 2295
		    			}else{		 
		    			    if((df_md.format(insertValue.doubleValue()).toString()).equals("0.000")){
		    			        cell.setCellValue("-");
		    			    }else{
		    			        cell.setCellValue(df_md.format(insertValue.doubleValue()).toString());//顯示小數點至第3位,不足者補0
		    			    }
		    			}
		    			
				        int countline=rowcount;
                        countline++;                        
		    		}//end of cellcount         
		    		rowNum++;               
		    	}//end of rowcount
		        	
		    	//最高
		    	sqlCmd.delete(0,sqlCmd.length());		    	
		    	sqlCmd.append("select  Max(deposit_var_rate)  as  AA_1,  Max(save_var_rate)  as  AA_2,");
		    	sqlCmd.append("		   Max(Period_1_FIX_Rate)  as  AA_3,  Max(Period_1_Var_Rate)  as  AA_4, ");
        		sqlCmd.append("		   Max(Period_3_FIX_Rate)  as  AA_5,  Max(Period_3_Var_Rate)  as  AA_6, ");
  				sqlCmd.append("		   Max(Period_6_FIX_Rate)  as  AA_7,  Max(Period_6_Var_Rate)  as  AA_8, ");
  				sqlCmd.append("		   Max(Period_9_FIX_Rate)  as  AA_9,  Max(Period_9_Var_Rate)  as  AA_10, ");
                sqlCmd.append("		   Max(Period_12_FIX_Rate)  as  AA_11, Max(Period_12_Var_Rate)  as  AA_12, ");
                sqlCmd.append("		   Max(Period_24_FIX_Rate)  as  AA_13, Max(Period_24_Var_Rate)  as  AA_14, ");
                sqlCmd.append("		   Max(Period_36_FIX_Rate)  as  AA_15, Max(Period_36_Var_Rate)  as  AA_16, ");
                sqlCmd.append("		   Max(deposit_12_fix_rate)  as  AA_17, Max(deposit_12_var_rate)  as  AA_18, ");
                sqlCmd.append("		   Max(deposit_24_fix_rate)  as  AA_19, Max(deposit_24_var_rate)  as  AA_20, ");
                sqlCmd.append("		   Max(deposit_36_fix_rate)  as  AA_21, Max(deposit_36_var_rate)  as  AA_22, ");
 				sqlCmd.append("		   Max(Basic_Pay_Var_Rate)  as  AA_23,");
				sqlCmd.append("	       Max(Period_House_Var_Rate)  as  AA_24, Max(Period_House_Var_Rate_month)  as  AA_25,");
  				sqlCmd.append("		   Max(Base_Mark_Rate)      as  AA_26, Max(Base_Mark_Rate_month)      as  AA_27,");
				sqlCmd.append("        Max(Base_Fix_Rate)  as  AA_28, ");
                sqlCmd.append("		  Max(Base_Base_Rate)      as  AA_29,Max(Base_Base_Rate_month)      as  AA_30 ");
				sqlCmd.append(" from ( ");
				sqlCmd.append("       select WLX_S_RATE.*  ");
				sqlCmd.append(" 	   from (select * from "+cd01_table+" cd01 where cd01.hsien_id <> 'Y' ) cd01 ");
				sqlCmd.append("   	   left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id          ");
				sqlCmd.append("   	   left join (select * from bn01 where m_year=?)bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=?");
    			sqlCmd.append("   	   left join (select * from WLX_S_RATE  where m_year=? and M_Quarter =?) WLX_S_RATE  ");
				sqlCmd.append("             on WLX_S_RATE.BANK_NO=bn01.bank_no  ");
				sqlCmd.append("	  ) Temp1 where Temp1.m_year > 0 ");
							   
		    	List dbMax = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"aa_1,aa_2,aa_3,aa_4,aa_5,aa_6,aa_7,aa_8,aa_9,aa_10,aa_11,aa_12,aa_13,aa_14,aa_15,aa_16,aa_17,aa_18,aa_19,aa_20,aa_21,aa_22,aa_23,aa_24,aa_25,aa_26,aa_27,aa_28,aa_29,aa_30");
		    	System.out.println("dbMax.size()="+dbMax.size());
		    	//最低//96.11.28 fix 若利率為0時,不算入最低利率中
		    	sqlCmd.delete(0,sqlCmd.length());		
		    	sqlCmd.append("select  Min(decode(deposit_var_rate,0,999,deposit_var_rate))  as  AA_1,  ");
				sqlCmd.append(" 	   Min(decode(save_var_rate,0,999,save_var_rate))  as  AA_2,");
		    	sqlCmd.append("        Min(decode(Period_1_FIX_Rate,0,999,Period_1_FIX_Rate))  as  AA_3,");
				sqlCmd.append("		   Min(decode(Period_1_Var_Rate,0,999,Period_1_Var_Rate))  as  AA_4, ");
        		sqlCmd.append("		   Min(decode(Period_3_FIX_Rate,0,999,Period_3_FIX_Rate))  as  AA_5, ");
				sqlCmd.append("		   Min(decode(Period_3_Var_Rate,0,999,Period_3_Var_Rate))  as  AA_6, ");
  				sqlCmd.append("		   Min(decode(Period_6_FIX_Rate,0,999,Period_6_FIX_Rate))  as  AA_7, ");
				sqlCmd.append("		   Min(decode(Period_6_Var_Rate,0,999,Period_6_Var_Rate))  as  AA_8, ");
  				sqlCmd.append("		   Min(decode(Period_9_FIX_Rate,0,999,Period_9_FIX_Rate))  as  AA_9, ");
				sqlCmd.append("		   Min(decode(Period_9_Var_Rate,0,999,Period_9_Var_Rate))  as  AA_10,");
                sqlCmd.append("		   Min(decode(Period_12_FIX_Rate,0,999,Period_12_FIX_Rate))  as  AA_11,");
				sqlCmd.append("		   Min(decode(Period_12_Var_Rate,0,999,Period_12_Var_Rate))  as  AA_12,");
                sqlCmd.append("		   Min(decode(Period_24_FIX_Rate,0,999,Period_24_FIX_Rate))  as  AA_13,");
				sqlCmd.append("		   Min(decode(Period_24_Var_Rate,0,999,Period_24_Var_Rate))  as  AA_14,");
                sqlCmd.append("		   Min(decode(Period_36_FIX_Rate,0,999,Period_36_FIX_Rate))  as  AA_15,");
				sqlCmd.append("		   Min(decode(Period_36_Var_Rate,0,999,Period_36_Var_Rate))  as  AA_16,");
                sqlCmd.append("		   Min(decode(deposit_12_fix_rate,0,999,deposit_12_fix_rate)) as AA_17, ");
				sqlCmd.append("		   Min(decode(deposit_12_var_rate,0,999,deposit_12_var_rate)) as AA_18, ");
                sqlCmd.append("		   Min(decode(deposit_24_fix_rate,0,999,deposit_24_fix_rate)) as AA_19, ");
				sqlCmd.append("		   Min(decode(deposit_24_var_rate,0,999,deposit_24_var_rate)) as AA_20, ");
                sqlCmd.append("		   Min(decode(deposit_36_fix_rate,0,999,deposit_36_fix_rate)) as AA_21, ");
				sqlCmd.append("		   Min(decode(deposit_36_var_rate,0,999,deposit_36_var_rate)) as AA_22, ");
 				sqlCmd.append("		   Min(decode(Basic_Pay_Var_Rate,0,999,Basic_Pay_Var_Rate))  as  AA_23, ");
				sqlCmd.append("		   Min(decode(Period_House_Var_Rate,0,999,Period_House_Var_Rate))  as  AA_24, ");
				sqlCmd.append("		   Min(decode(Period_House_Var_Rate_month,0,999,Period_House_Var_Rate_month))  as  AA_25, ");
  				sqlCmd.append("		   Min(decode(Base_Mark_Rate,0,999,Base_Mark_Rate))     as  AA_26,     ");
  				sqlCmd.append("		   Min(decode(Base_Mark_Rate_month,0,999,Base_Mark_Rate_month))     as  AA_27,     ");
				sqlCmd.append("		   Min(decode(Base_Fix_Rate,0,999,Base_Fix_Rate))  as  AA_28, ");
                sqlCmd.append("		   Min(decode(Base_Base_Rate,0,999,Base_Base_Rate))      as  AA_29, ");
                sqlCmd.append("		   Min(decode(Base_Base_Rate_month,0,999,Base_Base_Rate_month))      as  AA_30 ");
				sqlCmd.append(" from ( ");
				sqlCmd.append(" 	   select WLX_S_RATE.*  ");
				sqlCmd.append("  	   from (select * from "+cd01_table+" cd01 where cd01.hsien_id <> 'Y' ) cd01 ");
				sqlCmd.append("   	   left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id          ");
				sqlCmd.append("   	   left join (select * from bn01 where m_year=?)bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=? ");
    			sqlCmd.append("   	   left join (select * from WLX_S_RATE  where m_year=? and M_Quarter =?) WLX_S_RATE  ");  
				sqlCmd.append("             on WLX_S_RATE.BANK_NO=bn01.bank_no  ");
				sqlCmd.append(" ) Temp1 where Temp1.m_year > 0 ");
							   
		    	List dbMin = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"aa_1,aa_2,aa_3,aa_4,aa_5,aa_6,aa_7,aa_8,aa_9,aa_10,aa_11,aa_12,aa_13,aa_14,aa_15,aa_16,aa_17,aa_18,aa_19,aa_20,aa_21,aa_22,aa_23,aa_24,aa_25,aa_26,aa_27,aa_28,aa_29,aa_30");
		    	System.out.println("dbMin.size()="+dbMin.size());
		    	//平均//96.11.28 fix 若利率為0時,不算入平均利率
		    	sqlCmd.delete(0,sqlCmd.length());	
		    	sqlCmd.append("select COUNT(*) a_cnt,");
		    	sqlCmd.append("	      ROUND((SUM(deposit_var_rate)/sum(decode(deposit_var_rate,0,0,1))),4)    as  AA_1, ");
				sqlCmd.append("		  ROUND((SUM(save_var_rate)/sum(decode(save_var_rate,0,0,1))),4)          as  AA_2, ");					
       			sqlCmd.append("	      ROUND((SUM(Period_1_FIX_Rate)/sum(decode(Period_1_FIX_Rate,0,0,1))),4)  as  AA_3, ");
  				sqlCmd.append("		  ROUND((SUM(Period_1_Var_Rate)/sum(decode(Period_1_Var_Rate,0,0,1))),4)  as  AA_4, ");
  				sqlCmd.append("		  ROUND((SUM(Period_3_FIX_Rate)/sum(decode(Period_3_FIX_Rate,0,0,1))),4)  as  AA_5, ");
  				sqlCmd.append("		  ROUND((SUM(Period_3_Var_Rate)/sum(decode(Period_3_Var_Rate,0,0,1))),4)  as  AA_6, ");
  				sqlCmd.append("		  ROUND((SUM(Period_6_FIX_Rate)/sum(decode(Period_6_FIX_Rate,0,0,1))),4)  as  AA_7, ");
  				sqlCmd.append("		  ROUND((SUM(Period_6_Var_Rate)/sum(decode(Period_6_Var_Rate,0,0,1))),4)  as  AA_8, ");
  				sqlCmd.append("		  ROUND((SUM(Period_9_FIX_Rate)/sum(decode(Period_9_FIX_Rate,0,0,1))),4)  as  AA_9, ");
  				sqlCmd.append("		  ROUND((SUM(Period_9_Var_Rate)/sum(decode(Period_9_Var_Rate,0,0,1))),4)  as  AA_10, ");
  				sqlCmd.append("		  ROUND((SUM(Period_12_FIX_Rate)/sum(decode(Period_12_FIX_Rate,0,0,1))),4)    as  AA_11, ");
  				sqlCmd.append("		  ROUND((SUM(Period_12_Var_Rate)/sum(decode(Period_12_Var_Rate,0,0,1))),4)    as  AA_12, ");
  				sqlCmd.append("		  ROUND((SUM(Period_24_FIX_Rate)/sum(decode(Period_24_FIX_Rate,0,0,1))),4)    as  AA_13, ");
  				sqlCmd.append("		  ROUND((SUM(Period_24_Var_Rate)/sum(decode(Period_24_Var_Rate,0,0,1))),4)    as  AA_14, ");
  				sqlCmd.append("		  ROUND((SUM(Period_36_FIX_Rate)/sum(decode(Period_36_FIX_Rate,0,0,1))),4)    as  AA_15, ");
  				sqlCmd.append("		  ROUND((SUM(Period_36_Var_Rate)/sum(decode(Period_36_Var_Rate,0,0,1))),4)    as  AA_16, ");
  				sqlCmd.append("		  ROUND((SUM(deposit_12_FIX_Rate)/sum(decode(deposit_12_FIX_Rate,0,0,1))),4)  as  AA_17, ");
  				sqlCmd.append("		  ROUND((SUM(deposit_12_Var_Rate)/sum(decode(deposit_12_Var_Rate,0,0,1))),4)  as  AA_18, ");
  				sqlCmd.append("		  ROUND((SUM(deposit_24_FIX_Rate)/sum(decode(deposit_24_FIX_Rate,0,0,1))),4)  as  AA_19, ");
  				sqlCmd.append("		  ROUND((SUM(deposit_24_Var_Rate)/sum(decode(deposit_24_Var_Rate,0,0,1))),4)  as  AA_20, ");
  				sqlCmd.append("		  ROUND((SUM(deposit_36_FIX_Rate)/sum(decode(deposit_36_FIX_Rate,0,0,1))),4)  as  AA_21, ");
  				sqlCmd.append("		  ROUND((SUM(deposit_36_Var_Rate)/sum(decode(deposit_36_Var_Rate,0,0,1))),4)  as  AA_22, ");
  				sqlCmd.append("		  ROUND((SUM(Basic_Pay_Var_Rate)/sum(decode(Basic_Pay_Var_Rate,0,0,1))),4)    as  AA_23, ");
  				sqlCmd.append("		  ROUND((SUM(Period_House_Var_Rate)/sum(decode(Period_House_Var_Rate,0,0,1))),4)  as  AA_24, ");
  				sqlCmd.append("		  decode(sum(decode(Period_House_Var_Rate_month,0,0,1)),0,0, ROUND((SUM(Period_House_Var_Rate_month)/sum(decode(Period_House_Var_Rate_month,0,0,1))),4))  as  AA_25, ");
  				sqlCmd.append("		  ROUND((SUM(Base_Mark_Rate)/sum(decode(Base_Mark_Rate,0,0,1))),4)  as  AA_26, ");
  				sqlCmd.append("		  decode(sum(decode(Base_Mark_Rate_month,0,0,1)),0,0,ROUND((SUM(Base_Mark_Rate_month)/sum(decode(Base_Mark_Rate_month,0,0,1))),4))  as  AA_27, ");
  				sqlCmd.append("		  ROUND((SUM(Base_Fix_Rate)/sum(decode(Base_Fix_Rate,0,0,1))),4)    as  AA_28, ");
  				sqlCmd.append("		  ROUND((SUM(Base_Base_Rate)/sum(decode(Base_Base_Rate,0,0,1))),4)  as  AA_29, ");
  				sqlCmd.append("		  decode(sum(decode(Base_Base_Rate_month,0,0,1)),0,0,ROUND((SUM(Base_Base_Rate_month)/sum(decode(Base_Base_Rate_month,0,0,1))),4))  as  AA_30 ");
				sqlCmd.append(" from ( ");
				sqlCmd.append(" 		select WLX_S_RATE.*  ");
				sqlCmd.append("  	    from (select * from "+cd01_table+" cd01 where cd01.hsien_id <> 'Y' ) cd01 ");
				sqlCmd.append("   	    left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id          ");
				sqlCmd.append("   	    left join (select * from bn01 where m_year=?)bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=?");
    			sqlCmd.append("   	    left join (select * from WLX_S_RATE  where m_year=? and M_Quarter =?) WLX_S_RATE  ");  
				sqlCmd.append("             on WLX_S_RATE.BANK_NO=bn01.bank_no  ");
				sqlCmd.append(" ) Temp1 where Temp1.m_year > 0 ");
		    	List dbSum = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"a_cnt,aa_1,aa_2,aa_3,aa_4,aa_5,aa_6,aa_7,aa_8,aa_9,aa_10,aa_11,aa_12,aa_13,aa_14,aa_15,aa_16,aa_17,aa_18,aa_19,aa_20,aa_21,aa_22,aa_23,aa_24,aa_25,aa_26,aa_27,aa_28,aa_29,aa_30");
		        System.out.println("dbSum.size()="+dbSum.size());		    	
		    	rowNum = rowNum+2;
		    	String totalValue = String.valueOf(((DataObject) dbSum.get(0)).getValue("a_cnt"));
		    	row = sheet.createRow(rowNum);
		    	cell=row.createCell((short)0);	
		    	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		    	cell.setCellStyle(cellStyle);
		    	cell.setCellValue("統計家數");
		    	
		    	cell=row.createCell((short)1);	
		    	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		    	cell.setCellStyle(cellStyle);
		    	cell.setCellValue(totalValue);        		
		    	
		    	
		    	rowNum = rowNum+1;
		    	row = sheet.createRow(rowNum);
		    	cell=row.createCell((short)0);	
		    	cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		    	cell.setCellStyle(cellStyle);
		    	cell.setCellValue("最高");
		    	
		    	String index="";
		    	
				for(int y=1;y<31;y++){
					index="aa_"+String.valueOf(y);					
					insertValue = Double.valueOf(((DataObject) dbMax.get(0)).getValue(index).toString());
					cell=row.createCell((short)y);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);					
					cell.setCellStyle(cellStyle);
		  			cell.setCellValue(df_md.format(insertValue.doubleValue()).toString());//顯示小數點至第3位,不足者補0 
				}
				
				rowNum = rowNum+1;
				row = sheet.createRow(rowNum);
				cell=row.createCell((short)0);	
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cellStyle);
				cell.setCellValue("最低");
				
				for(int y=1;y<31;y++){
					index="aa_"+String.valueOf(y);
					insertValue = Double.valueOf(((DataObject) dbMin.get(0)).getValue(index).toString());
					cell=row.createCell((short)y);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);					
					cell.setCellStyle(cellStyle);
					cell.setCellValue(df_md.format(insertValue.doubleValue()).toString());//顯示小數點至第3位,不足者補0  
				}		         
				rowNum = rowNum+1;
				row = sheet.createRow(rowNum);
				cell=row.createCell((short)0);	
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cellStyle);
				cell.setCellValue("平均");
				
				for(int y=1;y<31;y++){					
					index="aa_"+String.valueOf(y);
					insertValue = Double.valueOf(((DataObject) dbSum.get(0)).getValue(index).toString());
					cell=row.createCell((short)y);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
 					cell.setCellStyle(cellStyle);
 					cell.setCellValue(df_md.format(insertValue.doubleValue()).toString());//顯示小數點至第3位,不足者補0		  			
				}
				//101.12.21 add 說明文字 by 2295
				rowNum = rowNum+3;
                row = sheet.createRow(rowNum);
                cell=row.createCell((short)0);  
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);               
                cell.setCellValue("1.資料來源:全國農業金庫於每月底依各農(漁)會信用部由自有電腦設備或委由相關資訊中心該日以網際網路傳送之資料，彙整傳送至本局農業金融機構網際網路申報系統。");
                rowNum = rowNum+1;
                row = sheet.createRow(rowNum);
                cell=row.createCell((short)0);  
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);                
                cell.setCellValue("2.基本放款利率：係始自民國74年，至92年起依主管機關規定，全部改為基準利率，唯部分舊約未更換計價方式，仍使用原基本放款利率計價。");

		    }			
						
		    HSSFFooter footer=sheet.getFooter();
	        footer.setCenter( "Page:" +HSSFFooter.page() +" of " +HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
			FileOutputStream fout=new FileOutputStream(reportDir+ System.getProperty("file.separator")+ openfile);
			wb.write(fout);
	        //儲存 
	        fout.close();
	        System.out.println("儲存完成");
		}catch(Exception e){
			System.out.println("createRpt Error:"+e+e.getMessage());
		}
		return errMsg;
	}	
}



