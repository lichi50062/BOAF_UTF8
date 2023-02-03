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
    reportUtil reportUtil = new reportUtil();
    String report_no = "DS065W";
    String titleName = Utility.getPgName(report_no);
    String BankList = "";//儲存bank_code/bank_name
    String btnFieldList = "";//儲存所選取的大類acc_code/名稱
    String SortList = "";//排序的acc_code
    String CANCEL_NO = "";//裁撤別
    String Unit = "";//列印單位
    String loan_item = "";//貸款項目別
    String itemName ="";
    String S_YEAR = "";//年
    String S_MONTH = "";//月
    String lastYY = "";
    String lastMM = "";
    List BankList_data = null;//儲存bank_code/bank_name的集合
    List btnFieldList_data = null;
    List SortList_data = null;    
	int i = 0;
	int j = 0;
	String lguser_name = ( session.getAttribute("muser_name")==null ) ? "" : (String)session.getAttribute("muser_name");
	String nowbank_type="";
	String hasBankListALL="false";
	String str9701="";
	//99.04.29 add==================================================================
	String cd01_table = "";
    String wlx01_m_year = "";    
	//============================================================================
	try{
	      	
	    	nowbank_type = ((String)session.getAttribute("nowbank_type")).equals("")?"6":(String)session.getAttribute("nowbank_type");	
			System.out.println("nowbank_type="+nowbank_type);
			//儲存報表的目錄================================================================
        	File reportDir = new File(Utility.getProperties("reportDir"));       
    		if(!reportDir.exists()){
     			if(!Utility.mkdirs(Utility.getProperties("reportDir"))){
     	   			actMsg +=Utility.getProperties("reportDir")+"目錄新增失敗";
     			}    
    		}
    		//==============================================================================
    		//營運中/已裁撤
			if(session.getAttribute("CANCEL_NO") != null && !((String)session.getAttribute("CANCEL_NO")).equals("")){
		  		CANCEL_NO = (String)session.getAttribute("CANCEL_NO");		  				   
			}
    		//金融機構
			if(session.getAttribute("BankList") != null && !((String)session.getAttribute("BankList")).equals("")){
		   		BankList = (String)session.getAttribute("BankList");
		   		BankList_data = Utility.getReportData(BankList);
		   		System.out.println("BankList_data.size()="+BankList_data.size());		   
		   		System.out.println("BankList_data="+BankList_data);		   
			}
			//報表欄位
			if(session.getAttribute("btnFieldList") != null && !((String)session.getAttribute("btnFieldList")).equals("")){
		   		btnFieldList = (String)session.getAttribute("btnFieldList");
		   		btnFieldList_data = Utility.getReportData(btnFieldList);
		   		System.out.println("btnFieldList_data.size()="+btnFieldList_data.size());		   
		   		System.out.println("btnFieldList_data="+btnFieldList_data);		   
			}
			//排序欄位
			if(session.getAttribute("SortList") != null && !((String)session.getAttribute("SortList")).equals("")){
		  		SortList = (String)session.getAttribute("SortList");
		  		SortList_data = Utility.getReportData(SortList);
		   		System.out.println("SortList_data.size()="+SortList_data.size());		   
		   		System.out.println("SortList_data="+SortList_data);		   
			}
        	
			//年
			if(session.getAttribute("S_YEAR") != null && !((String)session.getAttribute("S_YEAR")).equals("")){
		  		S_YEAR = (String)session.getAttribute("S_YEAR");		  				   
			}
			//月
			if(session.getAttribute("S_MONTH") != null && !((String)session.getAttribute("S_MONTH")).equals("")){
		  		S_MONTH = (String)session.getAttribute("S_MONTH");		  				   
			}
			lastMM = (Integer.parseInt(S_MONTH)==1)?"12":String.valueOf(Integer.parseInt(S_MONTH)-1) ;
			lastYY = (Integer.parseInt(lastMM)==12)?String.valueOf(Integer.parseInt(S_YEAR)-1):S_YEAR;
			
			//金額單位
			if(session.getAttribute("Unit") != null && !((String)session.getAttribute("Unit")).equals("")){
		  		Unit = (String)session.getAttribute("Unit");		  				   
			}
			//貸款項目別
			if(session.getAttribute("loan_item") != null && !((String)session.getAttribute("loan_item")).equals("")){
			    loan_item = (String)session.getAttribute("loan_item");		  				   
			}
			//96.12.20 add 97/01以後,農會套用新表格(增加/異動科目代號) 
			if( nowbank_type.equals("6") && (Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH) >= 9701) ){
    			str9701 = "_9701";
    		}
    		//102.11.11 add 103/01以後,漁會套用新表格(增加/異動科目代號) 
    		if( nowbank_type.equals("7") && (Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH) >= 10301) ){
    			str9701 = "_10301";
    		}
			
			//99.04.29 add 查詢年度100年以前.縣市別不同===============================
  	    	cd01_table = (Integer.parseInt(S_YEAR) < 100)?"cd01_99":"cd01"; 
  	    	wlx01_m_year = (Integer.parseInt(S_YEAR) < 100)?"99":"100"; 
  	    	//=====================================================================         
		
			//讀取欄位大類所包含的細項===================================================================================        	
        	Properties prop_column = new Properties();
			prop_column.load(new FileInputStream(Utility.getProperties("schemaDir")+System.getProperty("file.separator")+"A01_"+nowbank_type+"_detail"+str9701+".TXT"));			
			//=======================================================================================================================			
			//取出欄位細項將資料存入MAP-->key=大類acc_code,value=細項acc_code=============================================================
			HashMap h_column = new HashMap();//儲存column大類,及其細項的acc_code			
			List detail_column = new LinkedList();
			String column_tmp = "";			
			String selectacc_code = "";//選取的detail科目代號
			String ori_field="";
			int columnLength=0;//column個數
			System.out.println("btnFieldList_data.size()="+btnFieldList_data.size());
			for(i=0;i<btnFieldList_data.size();i++){			    
			    column_tmp = (String)((List)btnFieldList_data.get(i)).get(0);
			    if(!"".equals(column_tmp)){
			        detail_column = Utility.getStringTokenizerData(column_tmp,"+");
			        if(detail_column != null && detail_column.size() != 0){               
			           columnLength += detail_column.size();//累加總欄位個數
			           String tmpStr = "";
              		   for(j=0;j<detail_column.size();j++){
              		     	String str=(String)detail_column.get(j);
            	 		   	selectacc_code +="'"+str+"'";  
            	 		   	if(!"".equals(str)){
            	 		      	String[] strarray=str.split("_",2);//使用limit，最多分割成2個字符串
            	 		      	tmpStr = strarray[1];
            	 		   	}
            	 		    ori_field += ",field_"+tmpStr+",type_"+tmpStr;            	 		                
            	 		    if(j < detail_column.size()-1){
            	 		      selectacc_code +=",";            	 		      
            	 		    }  
               		   }                              		   
            	   } 			      
			       h_column.put((String)((List)btnFieldList_data.get(i)).get(0),detail_column);			       
			       if(i < btnFieldList_data.size()-1){
			          selectacc_code +=",";
			       }   
			    }
			}
			System.out.println("h_column.size()="+h_column.size());
			System.out.println("select acc_code="+selectacc_code);
        	//讀取報表欄位名稱===================================================================================        	
        	Properties prop_column_name = new Properties();
			prop_column_name.load(new FileInputStream(Utility.getProperties("schemaDir")+System.getProperty("file.separator")+"A01_"+nowbank_type+"_column"+str9701+".TXT"));			
			//====================================================================================================
			String selectBank_no = "";//選取的金融機構代號
            //金融機構代號=============================================================
            if(BankList_data != null && BankList_data.size() != 0){               
               for(i=0;i<BankList_data.size();i++){
                   //95.09.04 判斷機構代號是否為ALL:全部===============================================
			       if(((String)((List)BankList_data.get(i)).get(0)).equals("ALL")){
			          hasBankListALL="true";
			       
			       }    			       			   
            	   selectBank_no +="'"+(String)((List)BankList_data.get(i)).get(0)+"'";            	
            	   if(i < BankList_data.size()-1) selectBank_no +=",";
               }               
               System.out.println("select bank_no="+selectBank_no);
            }   
            //==============================================================================
            String order = "";//排序欄位
            //排序欄位=========================================================================
            if(SortList_data != null && SortList_data.size() != 0){
            	for(i=0;i<SortList_data.size();i++){
            		order += (String)((List)SortList_data.get(i)).get(0);
            		if(i < SortList_data.size() -1 ) order +=",";            
	            }
	            System.out.println("order="+order);
            }
            //====================================================================================           
            StringBuffer sqlCmd =new StringBuffer();
            int parami = 0;
            List sqlCmd_paramList = new ArrayList();
			sqlCmd.append(" select a01.bank_code,bn01.bank_name,a01.m_year || '/' || m_month as m_yearmonth,wlx01.hsien_id,"+cd01_table+".hsien_name").append(ori_field);
			sqlCmd.append("   from ( ");
			sqlCmd.append("      select m_year,m_month,bank_code,");
			sqlCmd.append("        		field_loan_bal_cnt,"); 
			sqlCmd.append("        		field_loan_bal_amt,");
			sqlCmd.append("        		field_over6m_loan_bal_cnt,");  
			sqlCmd.append("        		field_over6m_loan_bal_amt,");      
			sqlCmd.append("        		field_delay_loan_cnt,");  
			sqlCmd.append("        		field_delay_loan_amt,");
			sqlCmd.append("        		field_over3p_loan_cnt,");  
			sqlCmd.append("        		field_over3p_loan_amt,");
			sqlCmd.append("        		field_again_loan_cnt,");
			sqlCmd.append("        		field_again_loan_bal_amt,");       
			sqlCmd.append("        		field_again_loan_bal_amt_1th,");
			sqlCmd.append("        		'0' as type_loan_bal_cnt,"); 
			sqlCmd.append("        		'0' as type_loan_bal_amt,");
			sqlCmd.append("        		'0' as type_over6m_loan_bal_cnt,"); 
			sqlCmd.append("        		'0' as type_over6m_loan_bal_amt,");
			sqlCmd.append("        		'0' as type_delay_loan_cnt,");
			sqlCmd.append("        		'0' as type_delay_loan_amt,");
			sqlCmd.append("        		'0' as type_over3p_loan_cnt,");
			sqlCmd.append("        		'0' as type_over3p_loan_amt,");
			sqlCmd.append("        		'0' as type_again_loan_cnt,");
			sqlCmd.append("        		'0' as type_again_loan_bal_amt,");
			sqlCmd.append("        		'0' as type_again_loan_bal_amt_1th ");
			sqlCmd.append(" 	   from (select agri_loan.m_year,agri_loan.m_month, agri_loan.bank_code,");
			sqlCmd.append(" 	                round(sum(loan_bal_cnt) /1,0) as field_loan_bal_cnt,");  
			sqlCmd.append(" 	                round(sum(loan_bal_amt) /1,0) as field_loan_bal_amt,");   
			sqlCmd.append(" 	                round(sum(over6m_loan_bal_cnt) /1,0) as field_over6m_loan_bal_cnt,"); 
			sqlCmd.append(" 	                round(sum(over6m_loan_bal_amt) /1,0) as field_over6m_loan_bal_amt,"); 
			sqlCmd.append(" 	                round(sum(delay_loan_cnt) /1,0) as field_delay_loan_cnt,"); 
			sqlCmd.append(" 	                round(sum(delay_loan_amt) /1,0) as field_delay_loan_amt,");
			sqlCmd.append(" 	                round(sum(over3p_loan_cnt) /1,0) as field_over3p_loan_cnt,");
			sqlCmd.append(" 	                round(sum(over3p_loan_amt) /1,0) as field_over3p_loan_amt,");                    
			sqlCmd.append(" 	                round(sum(again_loan_cnt) /1,0) as field_again_loan_cnt,");
			sqlCmd.append(" 	                round(sum(again_loan_bal_amt) /1,0) as field_again_loan_bal_amt,");
			sqlCmd.append(" 	                round(sum(again_loan_bal_amt_1th) /1,0) as field_again_loan_bal_amt_1th ");
			sqlCmd.append("        		   from agri_loan");
			sqlCmd.append("               where m_year=? and m_month=? ");
			sqlCmd_paramList.add(S_YEAR);
			sqlCmd_paramList.add(S_MONTH);
			if(!"ALL".equals(loan_item)){
				sqlCmd.append("             and loan_item=? ");//--貸款項目別,若為全部,則不加入此條件    
				sqlCmd_paramList.add(loan_item);
			}
			sqlCmd.append("         	   group by agri_loan.m_year,agri_loan.m_month,agri_loan.bank_code ");
			sqlCmd.append("         	)agri_loan ");
			sqlCmd.append(" 	)a01 "); 
			sqlCmd.append("		,(select * from bn01 where bn01.m_year=?)bn01 ");                
			sqlCmd.append("	left join (select * from wlx01 where m_year=?)wlx01 on bn01.bank_no = wlx01.bank_no ");
			sqlCmd_paramList.add(wlx01_m_year);
			sqlCmd_paramList.add(wlx01_m_year);
			sqlCmd.append(" left join "+cd01_table+" on wlx01.HSIEN_ID = "+cd01_table+".HSIEN_ID ");  
			sqlCmd.append("where a01.bank_code = bn01.bank_no ");
			sqlCmd.append("  and bn01.bank_no in ("+selectBank_no+") ");
			sqlCmd.append("  and bn01.bn_type <> '2' and a01.bank_code = bn01.bank_no ");
			sqlCmd.append("group by a01.bank_code,bn01.bank_name,a01.m_year,a01.m_month,wlx01.hsien_id,"+cd01_table+".hsien_name").append(ori_field);
			if(!"".equals(order)){	  
			    //100.05.12 fix 有挑選排序欄位年月時,查詢SQL error by 2295 						    	    
			    sqlCmd.append(" order by "+order); //各別機構
				//SoryBy=asc/desc
				if( session.getAttribute("SortBy") != null && !((String)session.getAttribute("SortBy")).equals("")){	   
			        sqlCmd.append(" "+(String)session.getAttribute("SortBy"));	  		            
			    }
			}else{  		       
			       sqlCmd.append(" order by a01.bank_code,a01.m_year,a01.m_month ");
			}
			
			
			List paramList = new ArrayList();
			StringBuffer sql = new StringBuffer();
			sql.append("select loan_item_name from agri_loan_item where loan_item = ?");		
			paramList.add(loan_item);
			List qList = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"loan_item_name");
			if(qList.size() > 0){
			    itemName = (String)((DataObject)qList.get(0)).getValue("loan_item_name");
			}else{
			    itemName = "全部"; 
			}
			
			
            //讀取報表欄位名稱===================================================================================        	
        	Properties p = new Properties();
			p.load(new FileInputStream(Utility.getProperties("schemaDir")+System.getProperty("file.separator")+"A01_"+nowbank_type+"_column"+str9701+".TXT"));			
			//====================================================================================================
			
            //Creating Cells
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.createSheet( "report" ); //建立sheet，及名稱
            wb.setSheetName(0, titleName, HSSFWorkbook.ENCODING_UTF_16 );
            HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定            
            //設定頁面符合列印大小
            //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
            //sheet.setAutobreaks(true); //自動分頁            
            sheet.setAutobreaks( false );
            ps.setScale( ( short )100 ); //列印縮放百分比
            ps.setPaperSize( ( short )9 ); //設定紙張大小 A4            
            ps.setLandscape( true ); // 設定橫印
            //ps.setFitWidth((short)14);
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
            //設定表頭(資產負債表/損益表)===============================================================================
            row = sheet.createRow( ( short )1 );
            reportUtil.createCell( wb, row, ( short )1, titleName, titleStyle );
            
            for(i=2;i<columnLength+5;i++){
              reportUtil.createCell( wb, row, ( short )i, "", noBorderDefaultStyle );
            }            
            sheet.addMergedRegion( new Region( ( short )1, ( short )1,
                                               ( short )1,
                                               ( short )(columnLength+3)) );
            //列印年月=======================================================================================
            row = sheet.createRow( ( short )2 );            
            row.setHeight((short) 0x200);
            //95.12.01 add 查詢年月區間
            reportUtil.createCell( wb, row, ( short )1, S_YEAR + "年" + S_MONTH + "月", titleStyle );
            for(i=2;i<columnLength+5;i++){
               reportUtil.createCell( wb, row, ( short )i, "", noBorderDefaultStyle );
            }            
            sheet.addMergedRegion( new Region( ( short )2, ( short )1,
                                               ( short )2,
                                               ( short )(columnLength+3) ) );
            //======================================================================================================                                                                     
            row = sheet.createRow( ( short )3 );            
            String printTime = Utility.getDateFormat("  HH:mm:ss");
            String printDate = Utility.getDateFormat("yyyy/MM/dd");                                    
            reportUtil.createCell( wb, row, ( short )1, "列印日期："+Utility.getCHTdate(printDate, 1)+printTime, noBoderStyle );
            sheet.addMergedRegion( new Region( ( short )3, ( short )1,
                                               ( short )3,
                                               ( short )(columnLength+3) ) );            
            row = sheet.createRow( ( short )4 );                        
            //列印單位=======================================================================================            
            //System.out.println("unit_name="+Utility.getUnitName(Unit));                                               
            //System.out.println("columnLength="+columnLength);
            reportUtil.createCell( wb, row, ( short )1, "貸款項目："+itemName+",單位：新台幣"+Utility.getUnitName(Unit)+"、％", noBorderLeftStyle );                                                
            sheet.addMergedRegion( new Region( ( short )4, ( short )1,
                                               ( short )4,
                                               ( short )3) );          
            //設定列印人員==========================================================            
            reportUtil.createCell( wb, row, ( short )4, "列印人員："+lguser_name, noBoderStyle );
            sheet.addMergedRegion( new Region( ( short )4, ( short )4,
                                               ( short )4,
                                               ( short )(columnLength+3) ) );            
            //報表欄位=======================================================================
            //列印單位代號+機構名稱            
            for(i=5;i<8;i++){
                row = sheet.createRow( ( short )i );
                reportUtil.createCell( wb, row, ( short )1, "單位代號", columnStyle );               
                reportUtil.createCell( wb, row, ( short )2, "單位名稱", columnStyle );   
                reportUtil.createCell( wb, row, ( short )3, "查詢年月", columnStyle );//95.12.01 add 查詢年月
            }           
            sheet.addMergedRegion( new Region( ( short )5, ( short )1,
                                               ( short )7, ( short )1) );                                                                     
            sheet.addMergedRegion( new Region( ( short )5, ( short )2,
                                               ( short )7, ( short )2) );                                              
            sheet.addMergedRegion( new Region( ( short )5, ( short )3,
                                               ( short )7, ( short )3) );                                                                                             
            row = sheet.createRow( ( short )5 );//大類表頭
            int columnIdx = 4;
            for(i=0;i<btnFieldList_data.size();i++){
               System.out.println("["+i+"]i="+(String)((List)btnFieldList_data.get(i)).get(1));
               System.out.println("columnIdx="+columnIdx);
               //設定表頭欄位
               for(j=columnIdx;j<((List)h_column.get(((List)btnFieldList_data.get(i)).get(0))).size() + columnIdx;j++){
                  reportUtil.createCell( wb, row, ( short )j, (String)((List)btnFieldList_data.get(i)).get(1), columnStyle );               
               }
               sheet.addMergedRegion( new Region( ( short )5, ( short )columnIdx,
                                               ( short )7,
                                               ( short )columnIdx));                                              
               columnIdx +=  ((List)h_column.get(((List)btnFieldList_data.get(i)).get(0))).size();                                             
            }
            
            row = sheet.createRow( ( short ) 6);//細項表頭
            columnIdx = 4;          
            for(i=0;i<btnFieldList_data.size();i++){
               //System.out.println("["+i+"]i="+(String)((List)btnFieldList_data.get(i)).get(1));
               //System.out.println("columnIdx="+columnIdx);
               detail_column = (List)h_column.get(((List)btnFieldList_data.get(i)).get(0));//取出該大項的細類
               //設定細項表頭欄位
               for(j=0 ;j<detail_column.size();j++){                    
                  //System.out.println((String)detail_column.get(j)+"="+Utility.ISOtoBig5((String)prop_column_name.get((String)detail_column.get(j))));                                           
                  reportUtil.createCell( wb, row, ( short )columnIdx,(String)detail_column.get(j), columnStyle );               
                  columnIdx ++;
               }                              
            }
            row = sheet.createRow( ( short ) 7);//細項-科目代號
            columnIdx = 4;          
            for(i=0;i<btnFieldList_data.size();i++){
               //System.out.println("["+i+"]i="+(String)((List)btnFieldList_data.get(i)).get(1));
               //System.out.println("columnIdx="+columnIdx);
               detail_column = (List)h_column.get(((List)btnFieldList_data.get(i)).get(0));//取出該大項的細類
               //設定細項表頭欄位
               for(j=0 ;j<detail_column.size();j++){ 
                  //System.out.println((String)detail_column.get(j)+"="+Utility.ISOtoBig5((String)prop_column_name.get((String)detail_column.get(j))));                                           
                  reportUtil.createCell( wb, row, ( short )columnIdx, (String)detail_column.get(j), columnStyle );               
                  columnIdx ++;
               }                              
            }
            
            //wb.setRepeatingRowsAndColumns( 0, 1, 8, 1, 3 ); //設定表頭 為固定 先設欄的起始再設列的起始              
            wb.setRepeatingRowsAndColumns(0, 1, columnLength+3, 1, 7); //設定表頭 為固定 先設欄的起始再設列的起始
  			  			
  			//System.out.println(report_no+"_Excel.sqlCmd="+sqlCmd);	   
  			System.out.println("ori_field="+ori_field); 
  			List dbData = null;
  			if(hasBankListALL.equals("false")){  
  					  
			  dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),sqlCmd_paramList,ori_field+"m_yearmonth");            
			  //System.out.println(report_no+"_Excel.sqlCmd="+sqlCmd);	          
			}else{//95.09.04 add 金融機構代號=ALL全部  
						 
			  //dbData = DBManager.QueryDB_SQLParam(sqlCmd_sum.toString(),sqlCmd_sum_paramList,ori_field+"m_yearmonth");     
			  //System.out.println(report_no+"_Excel.sqlCmd_sum="+sqlCmd_sum);	          
			}
			
			short rowNo = ( short )8;//資料起始列     
			//無資料時,顯示訊息========================================================================			
			if(dbData == null || dbData.size() == 0){
			   	row = sheet.createRow( rowNo );     
                row.setHeight((short) 0x120); 
                reportUtil.createCell( wb, row, ( short )1,"無資料" ,noBorderDefaultStyle ); 
                sheet.addMergedRegion( new Region( ( short )8, ( short )1,
                                               ( short )8,
                                               ( short )(columnLength+3)) );
			}else{
			//將DBData寫入===============================================================================================      
            acc_code_row = sheet.getRow(7);
            short lastCellNum = acc_code_row.getLastCellNum();
            //System.out.println("lastCellNum="+lastCellNum);                        
            columnIdx = 1;       
            double amt_d = 0.0;                 
            float amt_f = 0; 
            String amt_type="";
            String amt="";
            String prtbank_code="";
            for(i=0;i<dbData.size();i++){            
                row = sheet.createRow( rowNo );  
                //System.out.println("rowNo="+rowNo);   
                //System.out.println("m_yearmonth="+(((DataObject) dbData.get(i)).getValue("m_yearmonth")).toString());
                row.setHeight((short) 0x120);    
                //System.out.println("bank_code="+(String) ((DataObject) dbData.get(i)).getValue("bank_code"));                
                /*95.12.05 fix 讓User可根據Excel做其他運用.單位代號.機構名稱repeat
                if(prtbank_code.equals((String) ((DataObject) dbData.get(i)).getValue("bank_code"))){
                	reportUtil.createCell( wb, row, ( short )columnIdx, " ", defaultStyle );//單位代號
                	columnIdx++;
                	reportUtil.createCell( wb, row, ( short )columnIdx, " ", defaultStyle );//機構名稱
                	columnIdx++;                
                }else{*/
                	reportUtil.createCell( wb, row, ( short )columnIdx, (String) ((DataObject) dbData.get(i)).getValue("bank_code"), defaultStyle );//單位代號
                	columnIdx++;
                	reportUtil.createCell( wb, row, ( short )columnIdx, (String) ((DataObject) dbData.get(i)).getValue("bank_name"), defaultStyle );//機構名稱
                	columnIdx++;
                //}
                //95.12.01 add 查詢年月
                reportUtil.createCell( wb, row, ( short )columnIdx, (((DataObject) dbData.get(i)).getValue("m_yearmonth")).toString(), defaultStyle );//查詢年月
                columnIdx++;
                for(int cellIdx =4;cellIdx < (new Short(lastCellNum)).intValue();cellIdx++){                    
                     amt="";
                     amt_type = "";
                     cell = acc_code_row.getCell((short)cellIdx); 
                     String str = cell.getStringCellValue();
                     String tmpStr = "";
     	 		   	 if(!"".equals(str)){
     	 		      	String[] strarray=str.split("_",2);//使用limit，最多分割成2個字符串
     	 		      	tmpStr = strarray[1];
     	 		   	 }
                     amt = (((DataObject) dbData.get(i)).getValue("field_"+tmpStr)).toString();                     
                     amt_type = (((DataObject) dbData.get(i)).getValue("type_"+tmpStr)).toString();                                                             
                     if(!(amt_type.equals("4") || "loan_bal_cnt".equals(tmpStr) || "over6m_loan_bal_cnt".equals(tmpStr)
                             || "over3p_loan_cnt".equals(tmpStr) || "again_loan_cnt".equals(tmpStr))){//該值不為利率或戶數時.再除以單位4捨五入
                        amt = Utility.getRound(amt,Unit);                        
                     }
                     amt = Utility.setCommaFormat(amt);
                     if(amt_type.equals("4")) amt+="%";
                     reportUtil.createCell( wb, row, ( short )columnIdx, amt, rightStyle );               
                     columnIdx ++;
                }
                columnIdx = 1;
                rowNo++;
            }
            
            }//end of 有data   
           
            //95.09.04 add 把中間值的acc_code合併成一個欄位=======================================================================
            row = sheet.getRow(5);
            columnIdx = 4;
            for(i=0;i<btnFieldList_data.size();i++){
                //System.out.println("["+i+"]i="+(String)((List)btnFieldList_data.get(i)).get(0));                
                //System.out.println("columnIdx="+columnIdx);
                if(((String)((List)btnFieldList_data.get(i)).get(0)).length() > 6 ){//科目代號:6碼
                     //設定表頭欄位.把中間值的acc_code合併成一個欄位只顯示中文名稱=====================================
                     for(j=columnIdx;j<((List)h_column.get(((List)btnFieldList_data.get(i)).get(0))).size() + columnIdx;j++){
                         reportUtil.createCell( wb, row, ( short )j, (String)((List)btnFieldList_data.get(i)).get(1), columnStyle );               
                     }
                     sheet.addMergedRegion( new Region( ( short )5, ( short )columnIdx,
                                               ( short )7,
                                               ( short )(((List)h_column.get(((List)btnFieldList_data.get(i)).get(0))).size() + columnIdx - 1)) );                                              
                }                               
                columnIdx +=  ((List)h_column.get(((List)btnFieldList_data.get(i)).get(0))).size();                                                             
            }
            //設定寬度============================================================            
            for ( i = 1; i <= columnLength+3; i++ ) {                
                if(i==2){
                  sheet.setColumnWidth( ( short )i,
                                        ( short ) ( 256 * ( 25 + 4 ) ) );//機構名稱
                }else{
                  sheet.setColumnWidth( ( short )i,
                                        ( short ) ( 256 * ( 15 + 4 ) ) );
                }                        
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