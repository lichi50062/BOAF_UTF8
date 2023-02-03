<%
//101.08.07 created 農信保--金融機構別及地區別保證案件分析表(M206) 
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.util.*,java.io.*" %>
<%@ page import="org.apache.poi.poifs.filesystem.*,org.apache.poi.hssf.usermodel.*" %>
<%@ page import="org.apache.poi.hssf.util.Region" %>
<%@ page import="com.tradevan.util.Utility" %>								          
<%@ page import="com.tradevan.util.report.Report01" %>								          
<%@ page import="com.tradevan.util.report.HssfStyle" %>								          
<%@ page import="com.tradevan.util.report.reportUtil" %>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="java.text.*" %>
<%@ page import="java.lang.*" %>
<%
   response.setContentType("application/msexcel;charset=UTF-8");//以上這行設定本網頁為excel格式的網頁   
   String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");			  
   System.out.println("act="+act);
   if(act.equals("view")){
      //以上這行設定傳送到前端瀏覽器時的檔名為test1.xls
      //就是靠這一行，讓前端瀏覽器以為接收到一個excel檔 
      response.setHeader("Content-disposition","inline; filename=view.xls");
   }else if (act.equals("download")){   
      response.setHeader("Content-Disposition","attachment; filename=download.xls");
   }   
%>
<%
	DecimalFormat dft = new DecimalFormat("#.##");
	String actMsg = "";
	FileOutputStream fileOut = null;      	
    HSSFCellStyle defaultStyle;
    HSSFCellStyle rightStyle;
    HSSFCellStyle noBorderDefaultStyle;
    HSSFCellStyle noBorderLeftStyle;
    HSSFCellStyle titleStyle;
	HSSFCellStyle columnStyle;
	HSSFCellStyle noBoderStyle;
	HSSFRow row;
	HSSFRow acc_code_row;//讀取acc_code的row
	HSSFCell cell = null;//宣告一個儲存格
	String titleName = "";
    reportUtil reportUtil = new reportUtil();
    String BankList = "";//儲存bank_code/bank_name
    String btnFieldList = "";//儲存所選取的大類acc_code/名稱 
    String Unit = "";//列印單位
    String S_YEAR = "";//年
    String E_YEAR = "";//年
    String S_MONTH = "";//月
    String E_MONTH = "";//月    
    List btnFieldList_data = null;
    List SortList_data = null;    
	int i = 0;
	int j= 0;
	String lguser_name = ( session.getAttribute("muser_name")==null ) ? "" : (String)session.getAttribute("muser_name");
	String report_no="M206";
			
	try{	      	
			//儲存報表的目錄================================================================
        	File reportDir = new File(Utility.getProperties("reportDir"));       
    		if(!reportDir.exists()){
     			if(!Utility.mkdirs(Utility.getProperties("reportDir"))){
     	   			actMsg +=Utility.getProperties("reportDir")+"目錄新增失敗";
     			}    
    		}
    		//==============================================================================    		
    		
			//報表欄位
			if(session.getAttribute("btnFieldList") != null && !((String)session.getAttribute("btnFieldList")).equals("")){
		   		btnFieldList = (String)session.getAttribute("btnFieldList");
		   		btnFieldList_data = Utility.getReportData(btnFieldList);
		   		System.out.println("btnFieldList_data.size()="+btnFieldList_data.size());		   
		   		System.out.println("btnFieldList_data="+btnFieldList_data);		   
			}
						
			titleName += "金融機構別及地區別保證案件分析表(M206)";		  		
			//年
			if(session.getAttribute("S_YEAR") != null && !((String)session.getAttribute("S_YEAR")).equals("")){
		  		S_YEAR = (String)session.getAttribute("S_YEAR");		  				   
			}
			//年
			if(session.getAttribute("E_YEAR") != null && !((String)session.getAttribute("E_YEAR")).equals("")){
		  		E_YEAR = (String)session.getAttribute("E_YEAR");		  				   
			}			
			//月
			if(session.getAttribute("S_MONTH") != null && !((String)session.getAttribute("S_MONTH")).equals("")){
		  		S_MONTH = (String)session.getAttribute("S_MONTH");		  				   
			}
			//月
			if(session.getAttribute("E_MONTH") != null && !((String)session.getAttribute("E_MONTH")).equals("")){
		  		E_MONTH = (String)session.getAttribute("E_MONTH");		  				   
			}
			//金額單位
			if(session.getAttribute("Unit") != null && !((String)session.getAttribute("Unit")).equals("")){
		  		Unit = (String)session.getAttribute("Unit");		  				   
			}  
		
			//讀取欄位大類所包含的細項===================================================================================        	
        	Properties prop_column = new Properties();
			StringBuffer sql = new StringBuffer();		
	        List paramList = new ArrayList();
	        DataObject bean = null;
            sql.append(" select cano,rulename from rptrule_acgf where acc_tr_type=? and type='1' order by cano ");
     		paramList.add(report_no);
     		List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,""); 
			for(i=0;i<dbData.size();i++){
			    bean = (DataObject)dbData.get(i);
			    prop_column.setProperty((String)bean.getValue("cano"),(String)bean.getValue("rulename"));  
	        }		
			//=======================================================================================================================			
			//取出欄位細項將資料存入MAP-->key=大類acc_code,value=細項acc_code=============================================================
			HashMap h_column = new HashMap();//儲存column大類,及其細項的acc_code			
			List detail_column = new LinkedList();
			String column_tmp = "";			
			String selectacc_code = "";//選取的detail科目代號		
			String field_sum="";			
			String ori_field="";
			String trans_field="";
			int columnLength=0;//column個數
			for(i=0;i<btnFieldList_data.size();i++){			    
			    column_tmp = "";
			    column_tmp = (String)prop_column.get((String)((List)btnFieldList_data.get(i)).get(0));
			    //System.out.println("column_tmp="+column_tmp);
			    if(!column_tmp.equals("")){
			        detail_column = Utility.getStringTokenizerData(column_tmp,"+");
			        System.out.println(detail_column);
			        if(detail_column != null && detail_column.size() != 0){               
			           columnLength += detail_column.size();//累加總欄位個數
              		   for(j=0;j<detail_column.size();j++){
            	 		   selectacc_code +="'"+(String)detail_column.get(j)+"'";           	
            	 		   trans_field = (String)detail_column.get(j);
            	 		   ori_field += (String)detail_column.get(j);
            	 		   //合計欄位,不加入field_sum,獨立選取
            	 		  if(!trans_field.equals("guarantee_cnt_year_0") && !trans_field.equals("guarantee_cnt_year0")){
           	 		      	if(field_sum.indexOf(trans_field) == -1){//95.10.03 fix 若不存在時,才加入sum的acc_code裡
           	 		        	if(!trans_field.equals("guarantee_amt_year0") &&!trans_field.equals("loan_amt_year0") &&!trans_field.equals("guarantee_bal_year0") 
           	 		        	        &&!trans_field.equals("guarantee_cnt_sum0") &&!trans_field.equals("guarantee_amt_sum0") &&!trans_field.equals("loan_amt_sum0") ){
           	 		        	    field_sum += " ,sum(decode(loan_unit,'"+trans_field.substring(trans_field.length()-1,trans_field.length()).toUpperCase()+"',"+trans_field.substring(0,trans_field.length()-2)+",0)) as "+trans_field;
           	 		        	}else{
           	 		        		field_sum += " ,sum(decode(loan_unit,'"+trans_field.substring(trans_field.length()-1,trans_field.length()).toUpperCase()+"',"+trans_field.substring(0,trans_field.length()-1)+",0)) as "+trans_field;
           	 		        	}
						  	}              
           	 		   	  } 
            	 		  
            	 		   //====================================================================================================================================
            	 		   if(j < detail_column.size()-1){
            	 		      selectacc_code +=",";            	 		      
            	 		      ori_field +=",";            	 		      
            	 		   }  
               		   }                              		   
               		   //System.out.println("select acc_code="+selectacc_code);
            	   }else{
            	       
            	   }   			      
			       h_column.put((String)((List)btnFieldList_data.get(i)).get(0),detail_column);			       
			       if(i < btnFieldList_data.size()-1){
			          selectacc_code +=",";
			          ori_field +=",";            	
			       }   
			    }
			}
			System.out.println("h_column.size()="+h_column.size());
			System.out.println("select acc_code="+selectacc_code);        		
			 
  			/*
  			 select m_year || '/' || decode(length(m_month),'1','0','') || m_month as m_yearmonth, --資料年月
				    guarantee_cnt_year_a,guarantee_cnt_year_b,guarantee_cnt_year_g,guarantee_cnt_year_t, --本年度保證件數
				    guarantee_cnt_year_l,guarantee_cnt_year_c,guarantee_cnt_year_h,guarantee_cnt_year_j,
				    guarantee_cnt_year_d,guarantee_cnt_year_w,guarantee_cnt_year_k,guarantee_cnt_year_e,
				    guarantee_cnt_year_y,guarantee_cnt_year_n,guarantee_cnt_year_p,guarantee_cnt_year_s,
				    guarantee_cnt_year_u,guarantee_cnt_year_0,
				    guarantee_cnt_year0 --本年度保證件數-合計
				from (
				      select m_year,m_month,
				             sum(decode(loan_unit,'A',guarantee_cnt_year,0)) as guarantee_cnt_year_a,--本年度保證件數-農會
				             sum(decode(loan_unit,'B',guarantee_cnt_year,0)) as guarantee_cnt_year_b,--本年度保證件數-漁會
				             sum(decode(loan_unit,'G',guarantee_cnt_year,0)) as guarantee_cnt_year_g,--本年度保證件數-全國農業金庫
				             sum(decode(loan_unit,'T',guarantee_cnt_year,0)) as guarantee_cnt_year_t,--本年度保證件數-臺灣銀行
				             sum(decode(loan_unit,'L',guarantee_cnt_year,0)) as guarantee_cnt_year_l,--本年度保證件數-臺灣土地銀行
				             sum(decode(loan_unit,'C',guarantee_cnt_year,0)) as guarantee_cnt_year_c,--本年度保證件數-合作金庫銀行
				             sum(decode(loan_unit,'H',guarantee_cnt_year,0)) as guarantee_cnt_year_h,--本年度保證件數-華南商業銀行
				             sum(decode(loan_unit,'J',guarantee_cnt_year,0)) as guarantee_cnt_year_j,--本年度保證件數-彰化商業銀行
				             sum(decode(loan_unit,'D',guarantee_cnt_year,0)) as guarantee_cnt_year_d,--本年度保證件數-第一商業銀行
				             sum(decode(loan_unit,'W',guarantee_cnt_year,0)) as guarantee_cnt_year_w,--本年度保證件數-臺灣中小企銀
				             sum(decode(loan_unit,'K',guarantee_cnt_year,0)) as guarantee_cnt_year_k,--本年度保證件數-慶豐商業銀行
				             sum(decode(loan_unit,'E',guarantee_cnt_year,0)) as guarantee_cnt_year_e,--本年度保證件數-遠東商業銀行
				             sum(decode(loan_unit,'Y',guarantee_cnt_year,0)) as guarantee_cnt_year_y,--本年度保證件數-元大商業銀行
				             sum(decode(loan_unit,'N',guarantee_cnt_year,0)) as guarantee_cnt_year_n,--本年度保證件數-京城商業銀行
				             sum(decode(loan_unit,'P',guarantee_cnt_year,0)) as guarantee_cnt_year_p,--本年度保證件數-永豐商業銀行
				             sum(decode(loan_unit,'S',guarantee_cnt_year,0)) as guarantee_cnt_year_s,--本年度保證件數-臺灣新光銀行
				             sum(decode(loan_unit,'U',guarantee_cnt_year,0)) as guarantee_cnt_year_u,--本年度保證件數-國泰世華銀行             
				             sum(decode(loan_unit,'0',guarantee_cnt_year,0)) as guarantee_cnt_year_0, --本年度保證件數-合計 
				             sum(decode(loan_unit,'0',guarantee_cnt_year,0)) as guarantee_cnt_year0 --本年度保證件數-合計
				      from m206 where (to_char(m_year * 100 + m_month) >= UI.查詢年度月份(起) ex:10001 
				and to_char(m_year * 100 + m_month) <= UI.查詢年度月份(迄) ex:10106)
				      group by m_year,m_month
				     )m206
				group by m_year,m_month, 
				guarantee_cnt_year_a,guarantee_cnt_year_b,guarantee_cnt_year_g,guarantee_cnt_year_t,
				    guarantee_cnt_year_l,guarantee_cnt_year_c,guarantee_cnt_year_h,guarantee_cnt_year_j,
				    guarantee_cnt_year_d,guarantee_cnt_year_w,guarantee_cnt_year_k,guarantee_cnt_year_e,
				    guarantee_cnt_year_y,guarantee_cnt_year_n,guarantee_cnt_year_p,guarantee_cnt_year_s,
				    guarantee_cnt_year_u,guarantee_cnt_year_0,guarantee_cnt_year0
				order by m_year,m_month


  			*/
  			
            String column = "";//選取欄位           
			//======================================================
            StringBuffer sqlCmd =new StringBuffer();            
            int parami = 0;
            List sqlCmd_paramList = new ArrayList();
            
            List main_table_paramList = new ArrayList();
            StringBuffer main_table=new StringBuffer(); 
		          
			
			
			sqlCmd.append(" select m_year || '/' || decode(length(m_month),'1','0','') || m_month as m_yearmonth,");
			sqlCmd.append(ori_field);
            sqlCmd.append(" from (");
            sqlCmd.append("      select m_year,m_month");
            sqlCmd.append(field_sum +",");    
            sqlCmd.append(" sum(decode(loan_unit,'0',guarantee_cnt_year,0)) as guarantee_cnt_year_0,");
            sqlCmd.append(" sum(decode(loan_unit,'0',guarantee_cnt_year,0)) as guarantee_cnt_year0 ");
            sqlCmd.append("      from m206");
            sqlCmd.append("      where (to_char(m_year * 100 + m_month) >= ? and to_char(m_year * 100 + m_month) <= ?)");
            sqlCmd.append("      group by m_year,m_month");
            sqlCmd.append("     )m206");
            sqlCmd.append(" group by m_year,m_month,"+ori_field);
            sqlCmd.append(" order by m_year,m_month");
			
			sqlCmd_paramList.add(S_YEAR+S_MONTH);	
			sqlCmd_paramList.add(E_YEAR+E_MONTH);
			
						      
            System.out.println("sqlCmd="+sqlCmd.toString());
            //讀取報表欄位名稱===================================================================================        	
        	Properties prop_column_name = new Properties();
			sql = new StringBuffer();		
	        paramList = new ArrayList();
	        bean = null;
	        sql.append(" select cano,rulename from rptrule_acgf where acc_tr_type=? and type='2' order by cano ");
     		paramList.add(report_no);
     		dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,""); 
			for(i=0;i<dbData.size();i++){
			    bean = (DataObject)dbData.get(i);
			    prop_column_name.setProperty((String)bean.getValue("cano"),(String)bean.getValue("rulename"));  
	        }	
			//====================================================================================================
			
            //Creating Cells
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.createSheet( "report" ); //建立sheet，及名稱
            wb.setSheetName(0, titleName, HSSFWorkbook.ENCODING_UTF_16 );
            HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定            
            //設定頁面符合列印大小          
            sheet.setAutobreaks( false );
            ps.setScale( ( short )100 ); //列印縮放百分比
            ps.setPaperSize( ( short )9 ); //設定紙張大小 A4            
            ps.setLandscape( true ); // 設定橫印           
            HSSFFooter footer = sheet.getFooter();            
            //設定樣式和位置(請精減style物件的使用量，以免style物件太多excel報表無法開啟)
			defaultStyle = reportUtil.getDefaultStyle(wb);//有框內文置中
			rightStyle = reportUtil.getRightStyle(wb);//有框內文置右
    		noBorderDefaultStyle = reportUtil.getNoBorderDefaultStyle(wb);//無框內文置中
    		noBorderLeftStyle = reportUtil.getNoBorderLeftStyle(wb);//無框內文置左
			reportUtil.setDefaultStyle(defaultStyle);
			reportUtil.setNoBorderDefaultStyle(noBorderDefaultStyle);			
    		titleStyle = reportUtil.getTitleStyle(wb); //標題用
    		columnStyle = reportUtil.getColumnStyle(wb);//報表欄位名稱用--有框內文置中			                                               
    		noBoderStyle = reportUtil.getNoBoderStyle(wb);//無框置右			                                               
    		//============================================================================                        
            //設定表頭===============================================================================
            row = sheet.createRow( ( short )1 );
            reportUtil.createCell( wb, row, ( short )1, titleName, titleStyle );
            
            for(i=2;i<columnLength+2;i++){
              reportUtil.createCell( wb, row, ( short )i, "", noBorderDefaultStyle );
            }            
            sheet.addMergedRegion( new Region( ( short )1, ( short )1,
                                               ( short )1,
                                               ( short )(columnLength+1)) );
                                                           
            //列印年月=======================================================================================
            row = sheet.createRow( ( short )2 );            
            row.setHeight((short) 0x200);
            //95.12.01 add 查詢年月區間
            reportUtil.createCell( wb, row, ( short )1, S_YEAR + "年" + S_MONTH + "月~"+E_YEAR + "年" + E_MONTH + "月", titleStyle );
            for(i=2;i<columnLength+2;i++){
               reportUtil.createCell( wb, row, ( short )i, "", noBorderDefaultStyle );
            }            
            sheet.addMergedRegion( new Region( ( short )2, ( short )1,
                                               ( short )2,
                                               ( short )(columnLength+1) ) );
            //======================================================================================================                                                                     
            row = sheet.createRow( ( short )3 );            
            String printTime = Utility.getDateFormat("  HH:mm:ss");
            String printDate = Utility.getDateFormat("yyyy/MM/dd");                                    
            reportUtil.createCell( wb, row, ( short )1, "列印日期："+Utility.getCHTdate(printDate, 1)+printTime, noBoderStyle );
            sheet.addMergedRegion( new Region( ( short )3, ( short )1,
                                               ( short )3,
                                               ( short )(columnLength+1) ) );            
            row = sheet.createRow( ( short )4 );                        
            //列印單位=======================================================================================            
            //System.out.println("unit_name="+Utility.getUnitName(Unit));                                               
            //System.out.println("columnLength="+columnLength);
            reportUtil.createCell( wb, row, ( short )1, "單位：新台幣"+Utility.getUnitName(Unit)+"、％", noBorderLeftStyle );                                                
            /*
            sheet.addMergedRegion( new Region( ( short )4, ( short )1,
                                               ( short )4,
                                               ( short )2) );          
            */                                   
            //設定列印人員==========================================================            
           /* reportUtil.createCell( wb, row, ( short )2, Utility.ISOtoBig5("列印人員：")+lguser_name, noBoderStyle );
            sheet.addMergedRegion( new Region( ( short )4, ( short )2,
                                               ( short )4,
                                               ( short )(columnLength+1) ) );     */       
            //報表欄位=======================================================================
            //列印單位代號+機構名稱            
            for(i=5;i<8;i++){
                row = sheet.createRow( ( short )i );                
                reportUtil.createCell( wb, row, ( short )1, "查詢年月", columnStyle );
            }           
            sheet.addMergedRegion( new Region( ( short )5, ( short )1,
                                               ( short )7,
                                               ( short )1) );                                                                     
                                                                                                       
            row = sheet.createRow( ( short )5 );//大類表頭
            int columnIdx = 2;
            for(i=0;i<btnFieldList_data.size();i++){
               //System.out.println("["+i+"]i="+(String)((List)btnFieldList_data.get(i)).get(1));
               //System.out.println("columnIdx="+columnIdx);
               //設定表頭欄位
               for(j=columnIdx;j<((List)h_column.get(((List)btnFieldList_data.get(i)).get(0))).size() + columnIdx;j++){
                  reportUtil.createCell( wb, row, ( short )j, (String)((List)btnFieldList_data.get(i)).get(1), columnStyle );               
               }
               sheet.addMergedRegion( new Region( ( short )5, ( short )columnIdx,
                                               ( short )5,
                                               ( short )(((List)h_column.get(((List)btnFieldList_data.get(i)).get(0))).size() + columnIdx - 1)) );                                              
               columnIdx +=  ((List)h_column.get(((List)btnFieldList_data.get(i)).get(0))).size();                                             
            }
            
            row = sheet.createRow( ( short ) 6);//細項表頭
            columnIdx = 2;          
            for(i=0;i<btnFieldList_data.size();i++){
               //System.out.println("["+i+"]i="+(String)((List)btnFieldList_data.get(i)).get(1));
               //System.out.println("columnIdx="+columnIdx);
               detail_column = (List)h_column.get(((List)btnFieldList_data.get(i)).get(0));//取出該大項的細類
               //System.out.println("detail_column="+detail_column);
               //設定細項表頭欄位
               for(j=0 ;j<detail_column.size();j++){                    
                  //System.out.println((String)detail_column.get(j)+"="+((String)prop_column_name.get((String)detail_column.get(j)));                                           
                  reportUtil.createCell( wb, row, ( short )columnIdx, (String)prop_column_name.get((String)detail_column.get(j)), columnStyle );               
                  columnIdx ++;
               }                              
            }
            row = sheet.createRow( ( short ) 7);//細項-科目代號
            columnIdx = 2;          
            for(i=0;i<btnFieldList_data.size();i++){
               //System.out.println("["+i+"]i="+(String)((List)btnFieldList_data.get(i)).get(1));
               //System.out.println("columnIdx="+columnIdx);
               detail_column = (List)h_column.get(((List)btnFieldList_data.get(i)).get(0));//取出該大項的細類
               //設定細項表頭欄位
               for(j=0 ;j<detail_column.size();j++){ 
                  //System.out.println((String)detail_column.get(j)+"="+(String)prop_column_name.get((String)detail_column.get(j)));                                           
                  reportUtil.createCell( wb, row, ( short )columnIdx, (String)detail_column.get(j), columnStyle );               
                  columnIdx ++;
               }                              
            }
            
            //wb.setRepeatingRowsAndColumns( 0, 1, 8, 1, 3 ); //設定表頭 為固定 先設欄的起始再設列的起始              
            wb.setRepeatingRowsAndColumns(0, 1, columnLength+3, 1, 7); //設定表頭 為固定 先設欄的起始再設列的起始
  			  			
  			//System.out.println("DS001W_Excel.sqlCmd="+sqlCmd);	   
  			System.out.println("ori_field="+ori_field);   				
			dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),sqlCmd_paramList,ori_field+",m_yearmonth");
			System.out.println("dbData.size()="+dbData.size());
			short rowNo = ( short )8;//資料起始列     
			//無資料時,顯示訊息========================================================================			
			if(dbData == null || dbData.size() == 0){
			   	row = sheet.createRow( rowNo );     
                row.setHeight((short) 0x120); 
                reportUtil.createCell( wb, row, ( short )1,"無資料" ,noBorderDefaultStyle ); 
                sheet.addMergedRegion( new Region( ( short )8, ( short )1,
                                               ( short )8,
                                               ( short )(columnLength+1)) );
			}else{
			    //將DBData寫入===============================================================================================      
                acc_code_row = sheet.getRow(7);
                short lastCellNum = acc_code_row.getLastCellNum();
                //System.out.println("lastCellNum="+lastCellNum);                        
                columnIdx = 1; 
                String amt="";               
                for(i=0;i<dbData.size();i++){            
                    row = sheet.createRow( rowNo );  
                    bean = (DataObject)dbData.get(i);
                    //System.out.println("rowNo="+rowNo);   
                    //System.out.println("m_yearmonth="+(bean.getValue("m_yearmonth").toString());
                    row.setHeight((short) 0x120); 
                    
                    //95.12.01 add 查詢年月
                    reportUtil.createCell( wb, row, ( short )columnIdx, bean.getValue("m_yearmonth").toString(), defaultStyle );//查詢年月
                    columnIdx++;
                    for(int cellIdx =2;cellIdx < (new Short(lastCellNum)).intValue();cellIdx++){                    
                         amt="";                        
                         cell = acc_code_row.getCell((short)cellIdx);                    
                         //System.out.println("acc_code="+cell.getStringCellValue()); 
                         amt = bean.getValue(cell.getStringCellValue()).toString();
                         if(cell.getStringCellValue().indexOf("cnt") == -1 && cell.getStringCellValue().indexOf("cnt") == -1){//cnt為件數,不除以金額單位
                         	amt = Utility.getRound(amt,Unit);                         
                         }
                         //System.out.println("amt="+amt);
                        
                         amt = Utility.setCommaFormat(amt);
                         reportUtil.createCell( wb, row, ( short )columnIdx, amt, rightStyle );               
                         columnIdx ++;
                    }
                    columnIdx = 1;
                    rowNo++;                    
                }            
            }//end of 有data   
          
            //把中間值的acc_code合併成一個欄位=======================================================================     
                 
            row = sheet.getRow(7);
            columnIdx = 2;
            short lastCellNum = row.getLastCellNum();
            for(i=2;i<(new Short(lastCellNum)).intValue();i++){
                //System.out.println("["+i+"]i="+(String)((List)btnFieldList_data.get(i)).get(0));                
                        
                //System.out.println((((List)h_column.get(((List)btnFieldList_data.get(i)).get(0))).size() + columnIdx - 1));
                cell = row.getCell((short)i);     
                //System.out.println("i="+i+":"+cell.getStringCellValue());      
                if(cell.getStringCellValue().indexOf("year0") != -1 || cell.getStringCellValue().indexOf("sum0") != -1){//合併單一科目代號                									 
                	sheet.addMergedRegion( new Region( ( short )5, ( short )columnIdx,
                                               ( short )7,
                                               ( short )columnIdx) );      
                }else{//合併中間值科目代號
            	 	sheet.addMergedRegion( new Region( ( short )6, ( short )columnIdx,
                                               ( short )7,
                                               ( short )columnIdx) );      
            	}	
                    
                                                            
                   columnIdx++;                            
                //columnIdx +=  ((List)h_column.get(((List)btnFieldList_data.get(i)).get(0))).size();                                                             
            }
           
            //設定寬度============================================================            
            for ( i = 1; i <= columnLength+3; i++ ) {                                
                  sheet.setColumnWidth( ( short )i,
                                        ( short ) ( 256 * ( 15 + 4 ) ) );                
            }
			//======================================================================================
            //設定涷結欄位
            //sheet.createFreezePane(0,1,0,1);
            footer.setCenter( "Page:" + HSSFFooter.page() + " of " +
                             HSSFFooter.numPages() );		                                 
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));			
			
            // Write the output to a file            
            fileOut = new FileOutputStream( Utility.getProperties("reportDir")+System.getProperty("file.separator")+ titleName+".xls" );
            wb.write( fileOut );
            fileOut.close();            
            
            FileInputStream fin = new FileInputStream(Utility.getProperties("reportDir")+System.getProperty("file.separator")+ titleName+".xls");  		 
			ServletOutputStream out1 = response.getOutputStream();           
			byte[] line = new byte[8196];
			int getBytes=0;
			while( ((getBytes=fin.read(line,0,8196)))!=-1 ){		    		
				out1.write(line,0,getBytes);
				out1.flush();
	    	}
		
			fin.close();
			out1.close();            		      
        } catch ( Exception e ) {            
            e.printStackTrace();
            
        } finally {
            try {
                if ( fileOut != null ) {
                    fileOut.close();
                }
            } catch ( Exception e ) {
                  System.out.println(e.getMessage() );
            }
        }   
%>	    		