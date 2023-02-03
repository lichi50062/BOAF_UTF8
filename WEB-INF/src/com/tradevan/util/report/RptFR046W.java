/*
 * 97.05.09 fix 基本放款利率計價之舊貸案件總表/明細表..利率取到小數點第2位
 * 99.03.18 fix 根據查詢年度.100年以後取得新縣市別.100年以前取得舊縣市別 by 2295
 * 99.04.12 fix 根據查詢年度.100年以後取得新縣市別.100年以前取得舊縣市別 
 * 				使用PreparedStatement;並列印轉換後的SQL;套用QueryDB_SQLParam by 2295
 * 99.09.13 fix 調整100年度.縣市排列順序 by 2295
 *              100年(含)以後.明細表最後加印新北市/台中市/台南市 by 2295
 *102.11.19 add 原QueryDB改套用QueryDB_SQLParam by 2295
 *102.11.22 add 100年後農會報表備註 by 2968
 *106.08.07 add 總表.其他縣市.顯示成中華民國農會 by 2295 
 *108.04.19 add 調整總表.台灣省不包含中華民國農會 by 2295
 *              調整明細表.台灣省改其他.含台灣省/福建省/中華民國農會 by 2295
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



public class RptFR046W {
	
    public static String createRpt(String s_year,String s_month,String unit,String bank_type,String rptStyle){
          String errMsg = "";
          String sqlCmd = "";          
          StringBuffer  sqlSub = new StringBuffer();          
          StringBuffer  sqlSubtail = new StringBuffer(); 		    //各縣市、臺灣省、福建省小計之共同結尾
          StringBuffer  sqlSubtail2 = new StringBuffer();
          StringBuffer  sqlSubtail3 = new StringBuffer();
          StringBuffer  sqlTaiwanDiv = new StringBuffer();
          StringBuffer  sqlFukienDiv = new StringBuffer();
          StringBuffer  sqlDivSum = new StringBuffer();
          StringBuffer  sqlTaiwan = new StringBuffer();
          StringBuffer  sqlFukien = new StringBuffer();
	      StringBuffer  sqlTotal = new StringBuffer();
	      StringBuffer  sqlDetail = new StringBuffer();
	      StringBuffer  sqlDiv = new StringBuffer();
	      StringBuffer  sqlDivtemp = new StringBuffer();	
	      StringBuffer  sqlCombine = new StringBuffer();
	      StringBuffer  sqlSub_hsien = new StringBuffer();//各縣市小計用
	      StringBuffer  sqlSub_total = new StringBuffer();//總計用
	      StringBuffer  sqlSub_Taiwan = new StringBuffer();//台灣省用
	      StringBuffer  sqlSub_Fukien = new StringBuffer();//福建省用
	      StringBuffer  sqlSub_detail = new StringBuffer();//明細用
	      List sqlTotal_paramList = new ArrayList();   
  		  List sqlSub_paramList = new ArrayList();     
  		  List sqlTaiwan_paramList = new ArrayList();//台灣省用參數  
  		  List sqlFukien_paramList = new ArrayList();//福健省用參數  
  		  List sqlDetail_paramList = new ArrayList();//明細表用參數  
  		  List sqlDivtemp_paramList = new ArrayList();
  		  List sqlSubtail_paramList   = new ArrayList(); //各縣市、臺灣省、福建省小計之共同結尾
  		  List sqlSubtail2_paramList  = new ArrayList();
  		  List sqlSubtail3_paramList  = new ArrayList();
  		  List sqlTaiwanDiv_paramList = new ArrayList();
  		  List sqlFukienDiv_paramList = new ArrayList();
  		  List sqlDiv_paramList = new ArrayList();
  		  List sqlCombine_paramList = new ArrayList();
          List dbData_All = null;
          List dbData_Part = null;
          List dbData_Taiwan = null;          
             
          String field_seq = ""; 
          String hsien_id = ""; 
          String hsien_name = "";           
          String bank_no = "";              
          String bank_name = "";            
          String count_seq = "";            
          String over_cnt = "";             
          String over_amt = "";             
          String push_over_amt = "";        
          String totalamt= "";              
          String push_totalamt = "";        
          String over_total_rate = "";   
          
          String bank_type_name="";         
          String unit_name = "";            

          int rowNum=0;
          String cd01_table = "";
          String wlx01_m_year = "";
          StringBuffer sql = new StringBuffer();  
          List paramList = new ArrayList();  

          reportUtil reportUtil = new reportUtil();
          System.out.println("RptFR046W.bank_type="+bank_type);
          if(bank_type.equals("ALL")){//95.06.14 add 總表?加農漁會
             bank_type_name = "農漁會";
          }else{
             bank_type_name = (bank_type.equals("6"))?"農會":"漁會";
          } 
      try{
      	    unit_name = Utility.getUnitName(unit); 
      	    //99.03.18 add 查詢年度100年以前.縣市別不同===============================
      	    cd01_table = (Integer.parseInt(s_year) < 100)?"cd01_99":""; 
      	    wlx01_m_year = (Integer.parseInt(s_year) < 100)?"99":"100"; 
      	    //=====================================================================
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
            String openfile="全體農漁會信用部基本放款利率計價之舊貸案件";
            
            openfile+=(rptStyle.equals("0"))?"總表":"明細表";
            openfile+=".xls";

            System.out.println("開啟檔:" + openfile);
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
	        if(rptStyle.equals("0"))       ps.setScale( ( short )70 ); //列印縮放百分比
	        else  ps.setScale( ( short )65 ); //列印縮放百分比
			
	        ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
	  		//wb.setSheetName(0,"test");
	  		//設定表頭 為固定 先設欄的起始再設列的起始
	        wb.setRepeatingRowsAndColumns(0, 1, 21, 2, 3);

	  		finput.close();

            HSSFRow row=null;//宣告一列
	  		HSSFCell cell=null;//宣告一個儲存格

		    //各縣市、臺灣省、福建省、總計才加
	  		
            //各縣市小計用
	  		sqlSub_hsien.append(" select a01.hsien_id,a01.hsien_name,a01.FR001W_output_order,");
	  		sqlSub_hsien.append("        ' ' as bank_no,' ' as BANK_NAME,");
	  		sqlSub_hsien.append("        COUNT(*) as COUNT_SEQ,");
	  		sqlSub_hsien.append(" 		 'A90'    as field_SEQ,");
            //總計用
            sqlSub_total.append(" select ' ' as hsien_id,' 總   計 ' as hsien_name,'001' as FR001W_output_order,");
            sqlSub_total.append("        ' ' as bank_no,' ' as BANK_NAME,");
            sqlSub_total.append("        COUNT(*) as COUNT_SEQ,");
            sqlSub_total.append("        'A99'  as field_SEQ,");
            //臺灣省用
            if(rptStyle.equals("1")){//明細表.108.04.08 add 明細表的臺灣省改為其他縣市(包含台灣省及福建省)
                sqlSub_Taiwan.append(" select ' ' as hsien_id,'其他' as hsien_name,'025' as FR001W_output_order,");
            }else{
                sqlSub_Taiwan.append(" select ' ' as hsien_id,'臺灣省' as hsien_name,'025' as FR001W_output_order,");                
            }
            sqlSub_Taiwan.append("        ' ' as bank_no ,' ' as BANK_NAME,");
            sqlSub_Taiwan.append("        COUNT(*) as COUNT_SEQ, ");
            sqlSub_Taiwan.append(" 		  'A92'as field_SEQ,");
            //福建省用
            sqlSub_Fukien.append(" select ' ' as hsien_id,'福建省' as hsien_name,'235' as FR001W_output_order,");
            sqlSub_Fukien.append(" 		  ' ' as bank_no,' ' as BANK_NAME,");
            sqlSub_Fukien.append(" 		  COUNT(*) as COUNT_SEQ,");
            sqlSub_Fukien.append(" 		  'A93' as field_SEQ,");
            //明細用
            sqlSub_detail.append(" select a01.hsien_id,a01.hsien_name,a01.FR001W_output_order, ");
            sqlSub_detail.append("        a01.bank_no,a01.BANK_NAME,");
            sqlSub_detail.append("        1 as COUNT_SEQ,");
            sqlSub_detail.append(" 		  'A01' as field_SEQ,");
            					
            //SUM共同區段
            sqlDivSum.append(" SUM(over_cnt)   	  over_cnt ,");
            sqlDivSum.append(" SUM(over_amt)   	  over_amt ,");
            sqlDivSum.append(" SUM(push_over_amt)  push_over_amt ,");
            sqlDivSum.append(" SUM(totalamt)       totalamt,");
            sqlDivSum.append(" SUM(push_totalamt)  push_totalamt,");
            sqlDivSum.append(" decode((SUM(totalamt) - SUM(push_totalamt)),0,0,Round((SUM(over_amt)-SUM(push_over_amt))*10000/(SUM(totalamt) - SUM(push_totalamt)),0)) as over_total_rate");
            
            //共有sql
            sqlDiv.append(" from ( select nvl(cd01.hsien_id,' ')  as  hsien_id,");
            sqlDiv.append("               nvl(cd01.hsien_name,'OTHER')  as hsien_name,");
            sqlDiv.append("               cd01.FR001W_output_order     as  FR001W_output_order,");
            sqlDiv.append("			   	  bn01.bank_no,bn01.BANK_NAME,");//97.01.11 for 明細表		
            sqlDiv.append("			      round(sum(over_cnt) /1,0) as over_cnt,");
            sqlDiv.append("			      round(sum(over_amt) /1,0) as over_amt,"); 				
            sqlDiv.append("			      round(sum(push_over_amt) /1,0) as push_over_amt,");
            sqlDiv.append("			      round(sum(totalamt) /1,0) as totalamt,");		
            sqlDiv.append("			      round(sum(push_totalamt) /1,0) as push_totalamt,");     
            sqlDiv.append("			      round(sum(over_total_rate) /1,0) as over_total_rate");     
            sqlDiv.append("        from  (select * from "+cd01_table+" cd01 where cd01.hsien_id <> 'Y'");
                
            sqlTaiwanDiv.append(sqlDiv);    //臺灣省和福建省共用區段只差中間一小部分，故在這邊assign給sqlTaiwanDiv \uFFFDBsqlFukienDiv
            if(rptStyle.equals("1")){//明細表                
                //108.04.08 add 其他縣市(包含台灣省.福建省.其他(中華民國農會))
                sqlTaiwanDiv.append(" and cd01.Hsien_div in ('2','3','4')");//2:台灣省 3:福建省 4:其他(中華民國農會)
            }else{
                sqlTaiwanDiv.append(" and cd01.Hsien_div = '2'"); 
            }
            sqlFukienDiv.append(sqlDiv);
            sqlFukienDiv.append(" and cd01.Hsien_div = '3'");
            
            sqlDivtemp.append(" ) cd01 left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = "+wlx01_m_year);
            if(bank_type.equals("ALL")){//農漁會
            	sqlDivtemp.append("  left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type in ('6','7') and bn01.m_year = wlx01.m_year and bn01.m_year = "+wlx01_m_year+" and wlx01.m_year="+wlx01_m_year);
            }else{
            	sqlDivtemp.append("  left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type = ? and bn01.m_year = wlx01.m_year and bn01.m_year = "+wlx01_m_year+" and wlx01.m_year="+wlx01_m_year);
            	sqlDivtemp_paramList.add(bank_type);
            }   
            sqlDivtemp.append(" left join (select * from a09 where a09.m_year  = ? and a09.m_month  = ?) a01  ");
            sqlDivtemp_paramList.add(s_year);
            sqlDivtemp_paramList.add(s_month);
            sqlDivtemp.append(" on  bn01.bank_no = a01.bank_code ");
            sqlDivtemp.append(" group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,");
            sqlDivtemp.append(" bn01.bank_no ,  bn01.BANK_NAME ");
                        
            sqlDiv.append(sqlDivtemp);
            sqlTaiwanDiv.append(sqlDivtemp);    //臺灣省和福建省共用區段只差中間一小部分，故在這邊assign給sqlTaiwanDiv \uFFFDBsqlFukienDiv
            sqlFukienDiv.append(sqlDivtemp);
            
            for(int sqlDivtempi=0;sqlDivtempi<sqlDivtemp_paramList.size();sqlDivtempi++){
         	   sqlDiv_paramList.add(sqlDivtemp_paramList.get(sqlDivtempi));
         	   sqlTaiwanDiv_paramList.add(sqlDivtemp_paramList.get(sqlDivtempi));
         	   sqlFukienDiv_paramList.add(sqlDivtemp_paramList.get(sqlDivtempi));
            }  
            
            //共同的結尾 總計、臺灣省、福建省
			sqlSubtail.append(" ) a01 ");
			sqlSubtail.append(" where a01.bank_no <> ' '");//97.01.15 add
            
			//各縣市小計的結尾
			sqlSubtail2.append(" ) a01  ");
			sqlSubtail2.append(" where a01.bank_no <> ' '");//97.01.15 add
			sqlSubtail2.append(" GROUP  BY a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order ");					    
			

			/*end共同的sql段落
			  共有以下共用段落
			  sqlDiv      總表、明細、縣市小計 用
			  sqlTaiwanDiv  臺灣省小計用
			  sqlFukienDiv  福建省小計用
			  sqlDivSum
			  sqlSubhead
			  sqlSubtail
			*/

            //各縣市、臺灣省、福建省小計sql------------------------------
            sqlSub.append(sqlSub_hsien);//各縣市小計用no param
            sqlSub.append(sqlDivSum);//SUM共同區段no param
            //sqlSub += " over_amt,push_over_amt,totalamt,push_totalamt";
            sqlSub.append(sqlDiv);//共有sql,參數sqlDivtemp
            for(int sqlDivi=0;sqlDivi<sqlDiv_paramList.size();sqlDivi++){
            	sqlSub_paramList.add(sqlDiv_paramList.get(sqlDivi));
            }
            sqlSub.append(sqlSubtail2);//各縣市小計的結尾 no param
            
            sqlTaiwan.append(sqlSub_Taiwan);//臺灣省用 no param
            if(rptStyle.equals("1")){//108.04.19 add 明細表的下方臺灣省改其他.增加金額單位 
                sqlTaiwan.append(" SUM(over_cnt) over_cnt ,");
                sqlTaiwan.append(" round(SUM(over_amt)/?,0) as over_amt,");
                sqlTaiwan.append(" round(SUM(push_over_amt)/?,0) push_over_amt,");
                sqlTaiwan.append(" round(SUM(totalamt)/?,0) totalamt,");
                sqlTaiwan.append(" round(SUM(push_totalamt)/?,0) push_totalamt,");
                sqlTaiwan.append(" decode((SUM(totalamt) - SUM(push_totalamt)),0,0,Round((SUM(over_amt)-SUM(push_over_amt))*10000/(SUM(totalamt) - SUM(push_totalamt)),0)) as over_total_rate");   
                sqlTaiwan_paramList.add(unit);
                sqlTaiwan_paramList.add(unit);
                sqlTaiwan_paramList.add(unit);
                sqlTaiwan_paramList.add(unit);
            }else{//總表
                sqlTaiwan.append(" SUM(over_cnt) over_cnt ,");
                sqlTaiwan.append(" SUM(over_amt) over_amt,");
                sqlTaiwan.append(" SUM(push_over_amt) push_over_amt,");
                sqlTaiwan.append(" SUM(totalamt) totalamt,");
                sqlTaiwan.append(" SUM(push_totalamt) push_totalamt,");
                sqlTaiwan.append(" decode((SUM(totalamt) - SUM(push_totalamt)),0,0,Round((SUM(over_amt)-SUM(push_over_amt))*10000/(SUM(totalamt) - SUM(push_totalamt)),0)) as over_total_rate");   
            }
            sqlTaiwan.append(sqlTaiwanDiv);//參數sqlDivtemp
            
            for(int sqlTaiwanDivi=0;sqlTaiwanDivi<sqlTaiwanDiv_paramList.size();sqlTaiwanDivi++){
            	sqlTaiwan_paramList.add(sqlTaiwanDiv_paramList.get(sqlTaiwanDivi));
            }
            sqlTaiwan.append(sqlSubtail);//共同的結尾 總計、臺灣省、福建省 no param  
            
            sqlFukien.append(sqlSub_Fukien);//福建省用 no param
            sqlFukien.append(sqlDivSum);//SUM共同區段no param                		  	   			 
            sqlFukien.append(sqlFukienDiv);//參數sqlDivtemp
            for(int sqlFukienDivi=0;sqlFukienDivi<sqlFukienDiv_paramList.size();sqlFukienDivi++){
            	sqlFukien_paramList.add(sqlFukienDiv_paramList.get(sqlFukienDivi));
            }
            sqlFukien.append(sqlSubtail);//共同的結尾 總計、臺灣省、福建省 no param
            
            //總計sql，總表和明細表共用
            sqlTotal.append(sqlSub_total);//總計用no param
            sqlTotal.append(sqlDivSum);  //SUM共同區段no param                      
            sqlTotal.append(sqlDiv);//共有sql,參數sqlDivtemp
            for(int sqlDivi=0;sqlDivi<sqlDiv_paramList.size();sqlDivi++){
            	sqlTotal_paramList.add(sqlDiv_paramList.get(sqlDivi));
            }
            sqlTotal.append(sqlSubtail);//共同的結尾 總計、臺灣省、福建省 no param 
            //end 總計sql--------------------------------
            
            //明細sql-------------------------------------
            sqlDetail.append(sqlSub_detail);//明細用 no param
            sqlDetail.append(" over_cnt ,over_amt , push_over_amt ,totalamt, push_totalamt,over_total_rate");//sqlDivSum;                		   
            sqlDetail.append(sqlDiv);//共有sql,參數sqlDivtemp
            for(int sqlDivi=0;sqlDivi<sqlDiv_paramList.size();sqlDivi++){
            	sqlDetail_paramList.add(sqlDiv_paramList.get(sqlDivi));
            }
            sqlDetail.append(" ) a01 ");
            sqlDetail.append(" where bank_no is not null");
            //end  明細sql----------------------------------
            
            //組合sql---------------------------------------
            sqlCombine.append(" select hsien_id ,hsien_name,FR001W_output_order,");
            sqlCombine.append(" 	   bank_no,BANK_NAME,COUNT_SEQ,field_SEQ,");
            sqlCombine.append("		   over_cnt,");
            sqlCombine.append("		   round(over_amt /?,0) as over_amt,");
            sqlCombine.append("		   round(push_over_amt /?,0) as push_over_amt,");
            sqlCombine.append("		   round(totalamt /?,0) as totalamt,");
            sqlCombine.append("		   round(push_totalamt /?,0) as push_totalamt,");
            sqlCombine.append(" 	   over_total_rate");            		
            sqlCombine.append(" from ( ");

            if(!rptStyle.equals("0")){   	//明細表組合	
                sqlCombine_paramList.add(unit);
                sqlCombine_paramList.add(unit);
                sqlCombine_paramList.add(unit);
                sqlCombine_paramList.add(unit);
            	sqlCombine.append(sqlTotal); 
            	sqlCombine.append(" UNION ALL ");
            	sqlCombine.append(sqlDetail); 
            	sqlCombine.append(" UNION ALL ");
            	sqlCombine.append(sqlSub); 
            	sqlCombine.append(" )a01  ORDER by FR001W_output_order,field_SEQ,hsien_id,bank_no ");            	
            	//add sqlTotal參數
            	for(int sqlTotali=0;sqlTotali<sqlTotal_paramList.size();sqlTotali++){
                	sqlCombine_paramList.add(sqlTotal_paramList.get(sqlTotali));
                }
            	//add sqlDetail參數
            	for(int sqlDetaili=0;sqlDetaili<sqlDetail_paramList.size();sqlDetaili++){
                	sqlCombine_paramList.add(sqlDetail_paramList.get(sqlDetaili));
                }
            	//add sqlSub參數
            	for(int sqlSubi=0;sqlSubi<sqlSub_paramList.size();sqlSubi++){
                	sqlCombine_paramList.add(sqlSub_paramList.get(sqlSubi));
                }
            	
            }else{ //總表      
                sqlCombine_paramList.add(unit);
                sqlCombine_paramList.add(unit);
                sqlCombine_paramList.add(unit);
                sqlCombine_paramList.add(unit);
            	sqlCombine.append(sqlTotal);            	
            	sqlCombine.append(" UNION ALL ");
            	sqlCombine.append(sqlSub); 
            	sqlCombine.append(" UNION ALL ");
            	sqlCombine.append(sqlTaiwan);             	
            	sqlCombine.append(" UNION ALL ");
            	sqlCombine.append(sqlFukien);
            	sqlCombine.append(" )  a01  ORDER by FR001W_output_order,field_SEQ,hsien_id,bank_no ");
            	//add sqlTotal參數
                for(int sqlTotali=0;sqlTotali<sqlTotal_paramList.size();sqlTotali++){
                	sqlCombine_paramList.add(sqlTotal_paramList.get(sqlTotali));
                }
                //add sqlSub參數
                for(int sqlSubi=0;sqlSubi<sqlSub_paramList.size();sqlSubi++){
                	sqlCombine_paramList.add(sqlSub_paramList.get(sqlSubi));
                }
                //add sqlTaiwan參數
                for(int sqlTaiwani=0;sqlTaiwani<sqlTaiwan_paramList.size();sqlTaiwani++){
                	sqlCombine_paramList.add(sqlTaiwan_paramList.get(sqlTaiwani));
                }
                //sqlFukien參數
                for(int sqlFukieni=0;sqlFukieni<sqlFukien_paramList.size();sqlFukieni++){
                	sqlCombine_paramList.add(sqlFukien_paramList.get(sqlFukieni));
                }
            }
                    
            //end組合sql------------------------------------
			//建表開始--------------------------------------
            HSSFFont ft = wb.createFont();
  		    HSSFCellStyle cs = wb.createCellStyle();
	 		ft.setFontHeightInPoints((short)18);
	 		ft.setFontName("標楷體");
	 		cs.setFont(ft);
	 		cs.setAlignment(HSSFCellStyle.ALIGN_CENTER);
	   	   	row = sheet.createRow(0);//報表.標題
                     
			//總表
			if(rptStyle.equals("0")){
     		   System.out.println("總表sql="+sqlCombine);
	 		   dbData_All = DBManager.QueryDB_SQLParam(sqlCombine.toString(),sqlCombine_paramList,"hsien_id,hsien_name,count_seq,over_cnt,over_amt,push_over_amt,totalamt,push_totalamt,over_total_rate");
	 		   System.out.print("總表資料 共"+dbData_All.size()+"筆");
	 		   
	   	   	   for(int v=0;v<7;v++){
	   	   	   	row.createCell((short)v);
	   	   	   }
                        
	   	   	   cell = row.getCell( (short) 0);
	   	   	   cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	   	   	   cell.setCellValue(""+s_year+"年"+s_month+"月"+ bank_type_name +"信用部基本放款利率計價之舊貸案件總表");
	   	   	   cell.setCellStyle(cs);
                        
	   	   	   row = sheet.getRow(1);                        
	   	   	   cell = row.getCell( (short)5);
	   	   	   cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	   	   	   cell.setCellValue("單位:新臺幣 " + unit_name + ",%");
                        
               rowNum=4;
               DataObject bean = null;
			   String insertValue = "";
			   double sampleValue;
       		   for(int i=0;i<dbData_All.size();i++){
	   		 	 	bean = (DataObject)dbData_All.get(i);
	   		 	 	//getBeanData(bean);//取得各欄位data
	   		 	 	hsien_id = String.valueOf(bean.getValue("hsien_id"));
	   		 	 	hsien_name = String.valueOf(bean.getValue("hsien_name"));//單位名稱
	   		 	    if("其他".equals(hsien_name)){//106.08.07 總表的其他.顯示成中華民國農會
	                    hsien_name = "中華民國農會";
	                }
	   		 	 	count_seq = String.valueOf(bean.getValue("count_seq"));
	   		 	 	over_cnt = bean.getValue("over_cnt") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("over_cnt")));//剩餘件數
	   		 	 	over_amt = bean.getValue("over_amt") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("over_amt")));//剩餘金額(A)
	   		 	 	push_over_amt = bean.getValue("push_over_amt") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("push_over_amt")));//剩餘金額-催收款(B)
	   		 	 	totalamt = bean.getValue("totalamt") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("totalamt")));//全會放出總金額(C)
	   		 	 	push_totalamt = bean.getValue("push_totalamt") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("push_totalamt")));//全會放出總金額-催收款(D)
	   		 	    over_total_rate=(bean.getValue("over_total_rate") == null)?"":String.valueOf(Double.parseDouble(String.valueOf(bean.getValue("over_total_rate")))/100);//佔放款總額的比率(A-B)/(C-D)  		
				    System.out.println("*** i="+i+",hsien_id="+hsien_id+",hsien_name="+hsien_name);
				    for(int cellcount=0;cellcount<7;cellcount++){
 		       			row = sheet.createRow(rowNum+i);
      		   			cell = row.createCell( (short)cellcount);
     		   			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			   			insertValue = "";
			   
     		            if( cellcount==0 )     insertValue =hsien_name;//
     		            else if( cellcount==1 )insertValue =over_cnt;//剩餘件數
     		            else if( cellcount==2 )insertValue =over_amt;//剩餘金額(A)
     		            else if( cellcount==3 )insertValue =push_over_amt;//剩餘金額-催收款(B)
     		            else if( cellcount==4 )insertValue =totalamt;//全會放出總金額(C)
     		            else if( cellcount==5 )insertValue =push_totalamt;//全會放出總金額-催收款(D)
     		            else if( cellcount==6 )insertValue =over_total_rate;//佔放款總額的比率(A-B)/(C-D)
     		                           
     		            //System.out.println("insertValue="+insertValue);
						HSSFFont f = wb.createFont();
    					HSSFCellStyle cs2 = wb.createCellStyle();
    					f.setFontHeightInPoints((short)10);
    					cs2.setFont(f);
				
						insertValue = (insertValue.equals("null"))?"":insertValue;
						cell.setCellValue(insertValue);	
						
    			        if(cellcount == 0 ){
   				           //96.04.23農會總表多一個台灣省.調整福建省位置 by 2295
    			        	if(i==0){
    			        		cs2.setAlignment(HSSFCellStyle.ALIGN_LEFT);
    			        	}else if(Integer.parseInt(s_year) <= 99){	
    			        	         if((i==1 || i==2 || i==3  || i==26) && bank_type.equals("6")){
    			        		        cs2.setAlignment(HSSFCellStyle.ALIGN_CENTER);  					
    			        	         }else if((i==1 || i==2 || i==18) && bank_type.equals("7")){
    			        		        cs2.setAlignment(HSSFCellStyle.ALIGN_CENTER);
    			        	         }else cs2.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
    			        	}else if (Integer.parseInt(s_year) >= 100){//99.09.13調整100年度.縣市排列順序	
			        	         if((i<=7  || i==23) && bank_type.equals("6")){
			        		        cs2.setAlignment(HSSFCellStyle.ALIGN_CENTER);  					
			        	         }else if((i<=5 || i==16) && bank_type.equals("7")){
			        		        cs2.setAlignment(HSSFCellStyle.ALIGN_CENTER);
			        	         }else cs2.setAlignment(HSSFCellStyle.ALIGN_RIGHT);        
    			        	}else{
    			        		cs2.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
    			            }    			           
    			        }else{
	 			          cs2.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
	 			        }
	 			        cs2.setBorderBottom(HSSFCellStyle.BORDER_THIN);
                        cs2.setBorderLeft(HSSFCellStyle.BORDER_THIN);
                        cs2.setBorderRight(HSSFCellStyle.BORDER_THIN);
	 			        cell.setCellStyle(cs2);
					}//end of cellcount
			   }//end of dbData_All
		
	    	   rowNum = rowNum + dbData_All.size()+1;
	    	   rowNum = printReMark(wb,sheet,row,cell,rowNum,rptStyle,wlx01_m_year,bank_type);//列印總表結尾的備註
	    	             
			}else{//明細表
			   dbData_Part= DBManager.QueryDB_SQLParam(sqlCombine.toString(),sqlCombine_paramList,"hsien_name,bank_no,bank_name,count_seq,over_cnt,over_amt,push_over_amt,totalamt,push_totalamt,over_total_rate");
			   System.out.println("明細表資料 共"+dbData_Part.size()+"筆");
			   
	   		   for(int v=0;v<8;v++){
	   				row.createCell((short)v);
	   		   }

	           cell = row.getCell( (short) 0);
	           cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	           cell.setCellValue(""+s_year+"年"+s_month+"月"+ bank_type_name +"信用部基本放款利率計價之舊貸案件明細表");
	           cell.setCellStyle(cs);
               
	           row = sheet.getRow(1);
	           cell = row.getCell( (short)6);
	           cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	           cell.setCellValue("單位:新臺幣 " + unit_name + ",%");
               
               rowNum=3;
               DataObject bean = null;
               //共同資料的格式
               HSSFFont ft2 = wb.createFont();
               HSSFCellStyle cs3 = wb.createCellStyle();
               ft2.setFontHeightInPoints((short)10);
	           cs3.setFont(ft2);
	           
	           //設定給各地方農漁會信用部name部分 左靠
	           HSSFFont fl = wb.createFont();
               HSSFCellStyle cl = wb.createCellStyle();
               fl.setFontHeightInPoints((short)10);
	           cl.setFont(fl);
	           cl.setAlignment(HSSFCellStyle.ALIGN_LEFT);
	           cl.setBorderTop(HSSFCellStyle.BORDER_THIN);
	           cl.setBorderBottom(HSSFCellStyle.BORDER_THIN);
               cl.setBorderLeft(HSSFCellStyle.BORDER_THIN);
               cl.setBorderRight(HSSFCellStyle.BORDER_THIN);
               cs3.setBorderTop(HSSFCellStyle.BORDER_THIN);
			   cs3.setBorderBottom(HSSFCellStyle.BORDER_THIN);
               cs3.setBorderLeft(HSSFCellStyle.BORDER_THIN);
               cs3.setBorderRight(HSSFCellStyle.BORDER_THIN);
               String insertValue = ""; 
               double sampleValue;           
       		   for(int i=1;i<dbData_Part.size();i++){
	   		       bean = (DataObject)dbData_Part.get(i);
	   		       //getBeanData(bean);//取得各欄位data
	   		       
	   		       hsien_name = String.valueOf(bean.getValue("hsien_name"));
	   		       bank_no = String.valueOf(bean.getValue("bank_no"));
	   		       bank_name = String.valueOf(bean.getValue("bank_name"));
	   		       count_seq = String.valueOf(bean.getValue("count_seq"));
	   		       field_seq = String.valueOf(bean.getValue("field_seq"));
	   		       over_cnt = bean.getValue("over_cnt") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("over_cnt")));//剩餘件數
   		 	 	   over_amt = bean.getValue("over_amt") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("over_amt")));//剩餘金額(A)
   		 	 	   push_over_amt = bean.getValue("push_over_amt") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("push_over_amt")));//剩餘金額-催收款(B)
   		 	 	   totalamt = bean.getValue("totalamt") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("totalamt")));//全會放出總金額(C)
   		 	 	   push_totalamt = bean.getValue("push_totalamt") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("push_totalamt")));//全會放出總金額-催收款(D)
   		 	 	   //over_total_rate=(bean.getValue("over_total_rate") == null)?"":Utility.getPercentNumber((bean.getValue("over_total_rate")).toString());//佔放款總額的比率(A-B)/(C-D)
   		 	 	   over_total_rate=(bean.getValue("over_total_rate") == null)?"":String.valueOf(Double.parseDouble(String.valueOf(bean.getValue("over_total_rate")))/100);//佔放款總額的比率(A-B)/(C-D)  		
   			       
   			       row = sheet.createRow(rowNum+i);
				   for(int cellcount=0;cellcount<8;cellcount++){ 		         
      		   		   cell = row.createCell( (short)cellcount);
     		           cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			           cs3.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
			           insertValue = "";
			           
     		           if( cellcount==0 && field_seq.equals("A01") )insertValue =bank_no;
     		           else if( cellcount==1 && field_seq.equals("A01") )insertValue =bank_name;	
     		           else if( cellcount==1 && ( field_seq.equals("A90") ||field_seq.equals("A99") ) ) insertValue =hsien_name;      			    
     		           else if( cellcount==2 )insertValue =over_cnt;
     		           else if( cellcount==3 )insertValue =over_amt;
     		           else if( cellcount==4 )insertValue =push_over_amt;
     		           else if( cellcount==5 )insertValue =totalamt;
     		           else if( cellcount==6 )insertValue =push_totalamt;
     		           else if( cellcount==7 )insertValue =over_total_rate;
     		          
			
			           insertValue = (insertValue.equals("null"))?"":insertValue;
				       cell.setCellValue(insertValue);
                       
    			       if(cellcount == 0 || cellcount ==1 ){
    			          cs3.setAlignment(HSSFCellStyle.ALIGN_CENTER);
    			       }else{
	 			          cs3.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
    			       }				       
                       
				       if(cellcount==1){
				       	  cell.setCellStyle(cl);
				       }else{
				       	  cell.setCellStyle(cs3);
				       }
			       }//end of cellcount
				   if( field_seq.equals("A90") ||field_seq.equals("A99") ){
				  	  rowNum++;row = sheet.createRow(rowNum+i); 		
			  		  for(int cellcount=0;cellcount<8;cellcount++){
			          	 cell = row.createCell( (short)cellcount);
			          	 cell.setCellValue(""); 
			          	 cell.setCellStyle(cs3);
			          }
			          //小計和總計後要空一行
			       }
			   }//end of dbData_Part
			   //總計再取出來一次，放在最底下
	    	   bean = (DataObject)dbData_Part.get(0);
	    	   //getBeanData(bean);//取得各欄位data
	    	   
	   		   hsien_name = String.valueOf(bean.getValue("hsien_name"));
	   		   bank_no = String.valueOf(bean.getValue("bank_no"));
	   		   bank_name = String.valueOf(bean.getValue("bank_name"));
	   		   count_seq = String.valueOf(bean.getValue("count_seq"));
	   		   field_seq = String.valueOf(bean.getValue("field_seq"));
	   		   over_cnt = bean.getValue("over_cnt") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("over_cnt")));//剩餘件數
	 	 	   over_amt = bean.getValue("over_amt") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("over_amt")));//剩餘金額(A)
	 	 	   push_over_amt = bean.getValue("push_over_amt") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("push_over_amt")));//剩餘金額-催收款(B)
	 	 	   totalamt = bean.getValue("totalamt") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("totalamt")));//全會放出總金額(C)
	 	 	   push_totalamt = bean.getValue("push_totalamt") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("push_totalamt")));//全會放出總金額-催收款(D)
	 	 	   //over_total_rate=(bean.getValue("over_total_rate") == null)?"":Utility.getPercentNumber((bean.getValue("over_total_rate")).toString());//佔放款總額的比率(A-B)/(C-D)
	 	 	   over_total_rate=(bean.getValue("over_total_rate") == null)?"":String.valueOf(Double.parseDouble(String.valueOf(bean.getValue("over_total_rate")))/100);//佔放款總額的比率(A-B)/(C-D)  		
			   
			   for(int cellcount=0;cellcount<8;cellcount++){
 		           row = sheet.createRow(dbData_Part.size()+rowNum);
      		       cell = row.createCell( (short)cellcount);
     		       cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			       cs3.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
			       insertValue = "";
                   
     		       if( cellcount==1 && field_seq.equals("A99") )insertValue ="總計";
     		       else if( cellcount==2 )insertValue =over_cnt;
     		       else if( cellcount==3 )insertValue =over_amt;
     		       else if( cellcount==4 )insertValue =push_over_amt;
     		       else if( cellcount==5 )insertValue =totalamt;
     		       else if( cellcount==6 )insertValue =push_totalamt;
     		       else if( cellcount==7 )insertValue =over_total_rate;
     		     
                   
			       insertValue = (insertValue.equals("null"))?"0":insertValue;
                   cell.setCellValue(insertValue);
	 		       cell.setCellStyle(cs3);
			   }//end of cellcount
			   
			   //畫最下面的小計與總計的部分
			   rowNum = 3+rowNum+dbData_Part.size();
			
			   //列印台北市/高雄市
			   for(int i=0;i<dbData_Part.size();i++){
	   			   bean = (DataObject)dbData_Part.get(i);
	   			   field_seq = String.valueOf(bean.getValue("field_seq"));
	   			   hsien_name = String.valueOf(bean.getValue("hsien_name"));
	   			   if(Integer.parseInt(s_year) <= 99){	
	   			      if( (field_seq.equals("A90") || field_seq.equals("A99")) &&
	   				    (hsien_name.equals("台北市") || hsien_name.equals("高雄市") ) )
	   			      { 
	   			         rowNum = printCity(bean,sheet,row,cell,rowNum,hsien_name,cs3);	   			      
				      }//end of 台北市.高雄市
	   			   }else if(Integer.parseInt(s_year) >= 100){//99.09.13 100年(含)以後.明細表加印新北市/台中市/台南市	
	   			      if( (field_seq.equals("A90") || field_seq.equals("A99")) &&
		   				    (hsien_name.equals("新北市") || hsien_name.equals("臺北市") || hsien_name.equals("桃園市") 
		   				  || hsien_name.equals("臺中市") || hsien_name.equals("臺南市") 
						  || hsien_name.equals("高雄市") ) )
		   			      { 	   			        
	   			            rowNum = printCity(bean,sheet,row,cell,rowNum,hsien_name,cs3);		   			            
					      }//end of 新北市.台北市.桃園市.台中市.台南市.高雄市
	   			   }  
			   }//end of dbData_Part
 		
			   dbData_Taiwan= DBManager.QueryDB_SQLParam(sqlTaiwan.toString(),sqlTaiwan_paramList,"hsien_name,bank_no,bank_name,count_seq,field_seq,over_cnt,over_amt,push_over_amt,totalamt,push_totalamt,over_total_rate");
			  //列印臺灣省 		
			  for(int i=0;i<dbData_Taiwan.size();i++){
	   			  bean = (DataObject)dbData_Taiwan.get(i);
	   			  field_seq = String.valueOf(bean.getValue("field_seq"));
	   			  hsien_name = String.valueOf(bean.getValue("hsien_name"));	
	   			  if( field_seq.equals("A92") && hsien_name.equals("其他")){ 	
	   			      //hsien_name = "其他";//108.04.08 原下方合計欄的臺灣省改成其他
	   			  	  rowNum = printCity(bean,sheet,row,cell,rowNum,hsien_name,cs3);
				  }//end of 臺灣省
			  }//end of dbData_Taiwan
			  
			  bean = (DataObject)dbData_Part.get(0);
	   		  hsien_name = String.valueOf(bean.getValue("hsien_name"));
	   		  field_seq =String.valueOf(bean.getValue("field_seq"));	   		  
	   		  rowNum = printCity(bean,sheet,row,cell,rowNum,hsien_name,cs3);	   		  
	   		  rowNum = printReMark(wb,sheet,row,cell,rowNum,rptStyle,wlx01_m_year,bank_type);//列印明細表結尾的備註	   		       
            }
 		    //建表結束--------------------------------------
            HSSFFooter footer = sheet.getFooter();
            footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
            footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
            FileOutputStream fout=new FileOutputStream(reportDir + System.getProperty("file.separator")+openfile);
            wb.write(fout);//儲存
            fout.close();
            System.out.println("儲存完成");
      }catch(Exception e){
         System.out.println("createRpt Error:"+e+e.getMessage());
      }      
      return errMsg;
    }//end of createRpt
    private static int printCity(DataObject bean,HSSFSheet sheet, HSSFRow row,HSSFCell cell
    							 ,int rowNum,String hsien_name,HSSFCellStyle cs3){
    	String bank_no = "";            
        String bank_name = "";          
        String count_seq = "";          
        String over_cnt = "";        
        String over_amt = "";       
        String push_over_amt = "";      
        String totalamt= "";        
        String push_totalamt = "";
        String over_total_rate = "";
        String insertValue = "";     
        int rowNum_return = 0;
    	try{
    		 bank_no = String.valueOf(bean.getValue("bank_no"));                                                                                                                              
		     bank_name = String.valueOf(bean.getValue("bank_name"));                                                                                                    
		     count_seq = String.valueOf(bean.getValue("count_seq"));                                                                                                    
		     over_cnt = bean.getValue("over_cnt") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("over_cnt")));//剩餘件數                               
	 	     over_amt = bean.getValue("over_amt") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("over_amt")));//剩餘金額(A)                            
	 	     push_over_amt = bean.getValue("push_over_amt") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("push_over_amt")));//剩餘金額-催收款(B)      
	 	     totalamt = bean.getValue("totalamt") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("totalamt")));//全會放出總金額(C)                      
	 	     push_totalamt = bean.getValue("push_totalamt") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("push_totalamt")));//全會放出總金額-催收款(D)
	 	     //over_total_rate=(bean.getValue("over_total_rate") == null)?"":Utility.getPercentNumber((bean.getValue("over_total_rate")).toString());//佔放款總額的比率(A-B)/(C-D)
	 	     over_total_rate=(bean.getValue("over_total_rate") == null)?"":String.valueOf(Double.parseDouble(String.valueOf(bean.getValue("over_total_rate")))/100);//佔放款總額的比率(A-B)/(C-D)  		
			 row = sheet.createRow(rowNum++);   
			 rowNum_return = rowNum;
			 for(int cellcount=0;cellcount<8;cellcount++){                                                                                                             
                  cell = row.createCell( (short)cellcount);                                                                                                              
	              cell.setEncoding(HSSFCell.ENCODING_UTF_16);                                                                                                            
	              insertValue = "";                                                                                                                                      
	              if( cellcount==1 )insertValue =hsien_name;                                                                                                             
	              else if( cellcount==2 )insertValue =over_cnt;                                                                                                          
	              else if( cellcount==3 )insertValue =over_amt;                                                                                                          
	              else if( cellcount==4 )insertValue =push_over_amt;                                                                                                     
	              else if( cellcount==5 )insertValue =totalamt;                                                                                                          
	              else if( cellcount==6 )insertValue =push_totalamt;                                                                                                     
	              else if( cellcount==7 )insertValue =over_total_rate;                                                                                                 
	                                                                                                                                                                     
		          insertValue = (insertValue.equals("null"))?"":insertValue;                                                                                             
		          cell.setCellValue(insertValue);                                                                                                                        
		          cell.setCellStyle(cs3);                                                                                                                                
              }//end of cellcount                                    
    	}catch(Exception e){
    		System.out.println("printCity Error:"+e+e.getMessage());
    	}
    	return rowNum_return;
    }
    
    private static int printReMark(HSSFWorkbook wb,HSSFSheet sheet,HSSFRow row,HSSFCell cell,int rowNum,String rptStyle,String u_year,String bank_type){
    	int rowNum_return=0;
    	try{
    		HSSFFont ft4 = wb.createFont();
  		    HSSFCellStyle cs4 = wb.createCellStyle();
	        ft4.setFontHeightInPoints((short)12);
	        ft4.setFontName("標楷體");
	        cs4.setFont(ft4);
	        cs4.setAlignment(HSSFCellStyle.ALIGN_LEFT);
	        if(rptStyle.equals("1")){//明細表
	           rowNum = rowNum+2;
	        }
	 		row = sheet.createRow(rowNum);
    	    cell = row.createCell( (short)0);
  	        cell.setEncoding(HSSFCell.ENCODING_UTF_16);
  	        cell.setCellValue("資料來源:依各農(漁)會信用部由自有電腦設備或委由相關資訊中心以網際網路傳送之資料彙編");
		    cell.setCellStyle(cs4);
		    if(rptStyle.equals("0")){//總表
		        if("100".equals(u_year) && "6".equals(bank_type)){
		            /*106.08.07 取消顯示
    		        rowNum++;
    		        row = sheet.createRow(rowNum);
    	            cell = row.createCell( (short)0);
    	            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
    	            cell.setCellValue("填表說明:地區別欄之其他係指原臺灣省農會，該農會於102年5月22日更名為中華民國農會。");
    	            cell.setCellStyle(cs4);
    	            */
		        }
		    	rowNum++;
		    	rowNum_return = rowNum;
		    }else{//明細表
		    	rowNum_return = rowNum;
		    }
    	}catch(Exception e){
    		System.out.println("printReMark Error:"+e+e.getMessage());
    	}
    	return rowNum_return;
    }
    /*
    //99.04.13 取得各欄位data
    public static void getBeanData(DataObject bean){
    	try{
    			hsien_name = bean.getValue("hsien_name") == null?"":String.valueOf(bean.getValue("hsien_name"));//單位名稱
    			bank_no = bean.getValue("bank_no") == null?"":String.valueOf(bean.getValue("bank_no"));
  		        bank_name = bean.getValue("bank_name") == null?"":String.valueOf(bean.getValue("bank_name"));
    			count_seq = bean.getValue("count_seq") == null?"":String.valueOf(bean.getValue("count_seq"));
    			field_seq = bean.getValue("field_seq") == null?"":String.valueOf(bean.getValue("field_seq"));
    			over_cnt = bean.getValue("over_cnt") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("over_cnt")));//剩餘件數
   		 	 	over_amt = bean.getValue("over_amt") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("over_amt")));//剩餘金額(A)
   		 	 	push_over_amt = bean.getValue("push_over_amt") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("push_over_amt")));//剩餘金額-催收款(B)
   		 	 	totalamt = bean.getValue("totalamt") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("totalamt")));//全會放出總金額(C)
   		 	 	push_totalamt = bean.getValue("push_totalamt") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("push_totalamt")));//全會放出總金額-催收款(D)
   		 	    over_total_rate=(bean.getValue("over_total_rate") == null)?"":String.valueOf(Double.parseDouble(String.valueOf(bean.getValue("over_total_rate")))/100);//佔放款總額的比率(A-B)/(C-D)   		 	    
   
    	}catch(Exception e){
    		System.out.println("getBeanData Error:"+e+e.getMessage());
    	}
    }
    */
}
