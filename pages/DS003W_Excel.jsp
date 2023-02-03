<%
// 94.09.29 create by 2295
// 95.12.01 add 增加年月區間,可根據年月做sort,拿掉可選欄位做sort by 2295
// 95.12.05 fix 讓User可根據Excel做其他運用.單位代號.機構名稱repeat by 2295
// 99.05.13 fix 縣市合併&sql injection by 2808
//100.05.12 fix 有挑選排序欄位年月時,查詢SQL error by 2295 	
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
<%@ page import="java.util.List" %>
<%@ page import="java.text.SimpleDateFormat" %>
<%@ page import="java.util.StringTokenizer" %>
<%@ page import="java.lang.StringBuffer" %>
<%@ page import="java.lang.Short" %>
<%@ page import="java.lang.Math" %>
<%@ page import="java.text.DecimalFormat" %>

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
	String titleName = "信用部";
    reportUtil reportUtil = new reportUtil();
    String BankList = "";//儲存bank_code/bank_name
    String btnFieldList = "";//儲存所選取的大類acc_code/名稱
    String SortList = "";//排序的acc_code
    String CANCEL_NO = "";//裁撤別    
    String Unit = "";//列印單位
    String S_YEAR = "";//年
    String E_YEAR = "";//年
    String S_MONTH = "";//月
    String E_MONTH = "";//月
    List BankList_data = null;//儲存bank_code/bank_name的集合
    List btnFieldList_data = null;
    List SortList_data = null;    
	int i = 0;
	int j= 0;
	String lguser_name = "測試使用者";
	String bank_type="";
	String hasBankListALL="false";
	DataObject bean = null ;
	String u_year = "99" ; //縣市合併用
	String cd01Table = "cd01_99" ;//縣市合併用 99年以前 cd01_99 以後 cd01 
	try{
			bank_type = ((String)session.getAttribute("nowbank_type")).equals("")?"6":(String)session.getAttribute("nowbank_type");	
			System.out.println("bank_type="+bank_type);
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
        	
        	//機構類別
			if(session.getAttribute("nowbank_type") != null && !((String)session.getAttribute("nowbank_type")).equals("")){
			    if(((String)session.getAttribute("nowbank_type")).equals("6")){
			       titleName = "農會" + titleName;
			    }else if(((String)session.getAttribute("nowbank_type")).equals("7")){
			       titleName = "漁會" + titleName;
			    }else{
			       titleName = "農漁會" + titleName;
			    }			    
			}			
			
			titleName += "平均利率資料";		  		
			
			//年
			if(session.getAttribute("S_YEAR") != null && !((String)session.getAttribute("S_YEAR")).equals("")){
		  		S_YEAR = (String)session.getAttribute("S_YEAR");	
		  		if(Integer.parseInt(S_YEAR) > 99 ) {
		  			u_year = "100" ;
		  			cd01Table = "cd01" ;
		  		}
			}
			//年
			if(session.getAttribute("E_YEAR") != null && !((String)session.getAttribute("E_YEAR")).equals("")){
		  		E_YEAR = (String)session.getAttribute("E_YEAR");		  				   
			}
			//95.12.01 增加年月區間
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
			prop_column.load(new FileInputStream(Utility.getProperties("schemaDir")+System.getProperty("file.separator")+"A03_"+bank_type+"_detail.TXT"));			
			//=======================================================================================================================			
			//取出欄位細項將資料存入MAP-->key=大類acc_code,value=細項acc_code=============================================================
			HashMap h_column = new HashMap();//儲存column大類,及其細項的acc_code			
			List detail_column = new LinkedList();
			String column_tmp = "";			
			String selectacc_code = "";//選取的detail科目代號
			String a03_operation_field_sum="";
			String a03_operation_field="";
			String a03_field_sum="";
			String a03_field="";
			String ori_field="";
			int columnLength=0;//column個數
			for(i=0;i<btnFieldList_data.size();i++){			    
			    column_tmp = "";
			    column_tmp = (String)prop_column.get((String)((List)btnFieldList_data.get(i)).get(0));
			    //System.out.println("column_tmp="+column_tmp);
			    if(!column_tmp.equals("")){
			        detail_column = Utility.getStringTokenizerData(column_tmp,"+");
			        //System.out.println(detail_column);
			        if(detail_column != null && detail_column.size() != 0){               
			           columnLength += detail_column.size();//累加總欄位個數
              		   for(j=0;j<detail_column.size();j++){
            	 		   selectacc_code +="'"+(String)detail_column.get(j)+"'";            	
            	 		   //95.09.26 add================================================================================================================= 
            	 		   if(((String)detail_column.get(j)).length() <= 6){
            	 		      ori_field+= "amt"+(String)detail_column.get(j);        
            	 		      a03_field += ",amt"+(String)detail_column.get(j)+",a03.type_"+(String)detail_column.get(j);            	 		      
            	 		      a03_field_sum += " ,sum(decode(acc_code,'"+(String)detail_column.get(j)+"',amt,0)) as amt"+(String)detail_column.get(j)+", "
            	 		                +  " '0' as type_"+(String)detail_column.get(j);							
            	 		   }else{
            	 		      ori_field += (String)detail_column.get(j);            	 		                
            	 		      a03_field += ","+(String)detail_column.get(j)+",a03_operation.type_"+(String)detail_column.get(j);
            	 		      a03_field_sum += " ,sum(decode(acc_code,'"+(String)detail_column.get(j)+"',amt,0)) as "+(String)detail_column.get(j)+"," 
							                + "  sum(decode(acc_code,'"+(String)detail_column.get(j)+"',type,'0')) as type_"+(String)detail_column.get(j);
            	 		   }
            	 		   ori_field += ",type_"+(String)detail_column.get(j);							
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
			System.out.println("select acc_code="+selectacc_code);
			System.out.println("h_column.size()="+h_column.size());
        	//讀取報表欄位名稱===================================================================================        	
        	Properties prop_column_name = new Properties();
			prop_column_name.load(new FileInputStream(Utility.getProperties("schemaDir")+System.getProperty("file.separator")+"A03_"+bank_type+"_column.TXT"));			
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
            		//order += (String)((List)SortList_data.get(i)).get(0);            		            		
            		//95.09.26 add====================================================
            		if(hasBankListALL.equals("true")){//選全部時  
            		   order += (String)((List)SortList_data.get(i)).get(0);
            		}else{
            		   if(((String)((List)SortList_data.get(i)).get(0)).length() <= 6){
            		      order += "amt"+(String)((List)SortList_data.get(i)).get(0);            		      
            		   }else{
            		      order += (String)((List)SortList_data.get(i)).get(0);
            		   }
            		}               		       		
            		//===============================================================            		
            		if(i < SortList_data.size() -1 ) order +=",";            
	            }
	            System.out.println("order="+order);
            }
            //====================================================================================            
  			/*
  			//各別機構,按bank_code,m_year,m_month排序
            select a03.bank_code,bn01.bank_name,m_year || '/' || m_month as m_yearmonth,wlx01.hsien_id,cd01.hsien_name,a03.field_1200_1502_1503
            from (
            	  select m_year,m_month,bank_code,
            	         sum(decode(acc_code,'field_1200_1502_1503',amt,0)) as field_1200_1502_1503, 
            			 sum(decode(acc_code,'110000',amt,0)) as amt110000,
            			 sum(decode(acc_code,'110100',amt,0)) as amt110100, 
            			 '0' as type_110000,
            			 '0' as type_110100
            	  from (select * 
            		    from 
            		         ( select m_year,m_month,bank_code,acc_code,type,amt from a03_operation
            				   where (to_char(m_year * 100 + m_month) >= 9406 and to_char(m_year * 100 + m_month) <= 9408)  					  		   
            		  		   and bank_code in  ('6230012','6230023')
            		  		   and acc_code in('field_1200_1502_1503')
            				 ) a03_operation union
            	    		( select m_year,m_month,bank_code,acc_code,'0' as type,amt from a03    
            		  		  where (to_char(m_year * 100 + m_month) >= 9406 and to_char(m_year * 100 + m_month) <= 9408)
            				  and a03.bank_code in  ('6230012','6230023')
            		  		  and a03.acc_code in ('110000','110100','110200')
            				)
            		   ) 
            	  group by m_year,m_month,bank_code
            	 )a03 
            ,bn01 left join wlx01                 
            on bn01.bank_no = wlx01.bank_no
            left join cd01 on wlx01.HSIEN_ID = cd01.HSIEN_ID  
            where a03.bank_code = bn01.bank_no
            and bn01.bank_no in ('6230012','6230023') and bn01.bn_type <> '2' 
            and a03.bank_code = bn01.bank_no			
            group by a03.bank_code,bn01.bank_name,a03.m_year,a03.m_month,wlx01.hsien_id,cd01.hsien_name,a03.field_1200_1502_1503
            order by a03.bank_code,a03.m_year,a03.m_month
			//縣市別.按縣市.年月sort
            select  cd01.FR001W_OUTPUT_ORDER,
            	    a03_operation.HSIEN_ID as bank_code,cd01.hsien_name as bank_name,
            		m_year || '/' || m_month as m_yearmonth,
            	    sum(decode(acc_code,'field_220400',amt,0)) as field_220400,  
            	    sum(decode(acc_code,'field_220400',type,'')) as type_field_220400
            from a03_operation
            left join cd01 on a03_operation.HSIEN_ID = cd01.HSIEN_ID
            where (to_char(m_year * 100 + m_month) >= 9406 and to_char(m_year * 100 + m_month) <= 9408)
            and bank_code='ALL' and (type='1' or type='2' or type='4')
            and bank_type='6'
            group by cd01.FR001W_OUTPUT_ORDER,a03_operation.HSIEN_ID,cd01.hsien_name,m_year,m_month
            order by cd01.FR001W_OUTPUT_ORDER,a03_operation.HSIEN_ID,m_year,m_month
  			*/
  			
            String column = "";//選取欄位           
            StringBuffer condition = new StringBuffer() ;//其他條件
            List conditionList = new ArrayList() ;
            //add 營運中/已裁撤============================             
			condition.append(" and bn01.bn_type "+(CANCEL_NO.equals("N")?"<>":"=")+" ?  and a03.bank_code = bn01.bank_no ");
            conditionList.add("2") ;
			//======================================================
            //String sqlCmd="";
			StringBuffer sqlCmd = new StringBuffer() ;
			List sqlCmdList = new ArrayList() ;
			//String sqlCmd_sum="";//縣市別小計
            StringBuffer sqlCmd_sum = new StringBuffer() ;
            List sqlCmd_sumList = new ArrayList() ;
            //String a03_table="";//a03
            StringBuffer a03_table = new StringBuffer() ;
            List a03_tableList = new ArrayList() ;
            
            sqlCmd_sum.append(" select  cd01.FR001W_OUTPUT_ORDER,a03_operation.HSIEN_ID as bank_code,"
            		   + "		   cd01.hsien_name as bank_name,m_year || '/' || m_month as m_yearmonth"
            		   + a03_field_sum
					   + " from a03_operation "
				       + " left join ").append(cd01Table).append(" cd01 on a03_operation.hsien_id = cd01.hsien_id "
				       + " where (to_char(m_year * 100 + m_month) >= ? and to_char(m_year * 100 + m_month) <= ? )"					   					   
				       + " and   bank_type=? "
					   + " and   bank_code=? and (type= ?  or type=? or type=?)"//type=1-->平均,2-->sum小計,4-->利率					   
					   + " and   acc_code in ("+selectacc_code+")"
					   + " group by cd01.FR001W_OUTPUT_ORDER,a03_operation.HSIEN_ID,cd01.hsien_name,m_year,m_month" );					   
            sqlCmd_sumList.add(S_YEAR+S_MONTH) ;
            sqlCmd_sumList.add(E_YEAR+E_MONTH) ;
            sqlCmd_sumList.add(bank_type) ;
            sqlCmd_sumList.add("ALL") ;
            sqlCmd_sumList.add("1") ;
            sqlCmd_sumList.add("2") ;
            sqlCmd_sumList.add("4") ;
            
			a03_table.append(" ( "
				  	  + "  select m_year,m_month,bank_code"+ a03_field_sum
 				      + "  from (select * "
					  + "        from "
					  + "        ( select m_year,m_month,bank_code,acc_code,type,amt from a03_operation "
					  + "          where (to_char(m_year * 100 + m_month) >= ? and to_char(m_year * 100 + m_month) <= ? )"					     					  
			          + "          and   bank_code in ("+selectBank_no+")"				      
				      + "          and   acc_code in ("+selectacc_code+")"    
					  + "		 ) a03_operation union "
				      + "		 ( select m_year,m_month,bank_code,acc_code,'0' as type,amt from a03 "
  					  + "          where (to_char(m_year * 100 + m_month) >= ?  and to_char(m_year * 100 + m_month) <= ? )"					   
			          + "          and   bank_code in ("+selectBank_no+")"				      
				      + "          and   acc_code in ("+selectacc_code+")"    
					  + " 		 ) "
					  + "       ) "
				      + "  group by m_year,m_month,bank_code "
				      + " )a03 ");
			a03_tableList.add(S_YEAR+S_MONTH);		 
			a03_tableList.add(E_YEAR+E_MONTH) ;
			a03_tableList.add(S_YEAR+S_MONTH) ;
			a03_tableList.add(E_YEAR+E_MONTH) ;
			
			sqlCmd.append(" select a03.bank_code,bn01.bank_name,a03.m_year || '/' || a03.m_month as m_yearmonth,wlx01.hsien_id,cd01.hsien_name,"
			       + ori_field			
		           + " from "+a03_table+",(select * from bn01 where m_year= ? )bn01 left join (select * from wlx01 where m_year= ? )wlx01 on bn01.bank_no = wlx01.bank_no "
			       + "                         left join ").append(cd01Table).append(" cd01 on wlx01.HSIEN_ID = cd01.HSIEN_ID " 
			       + " where a03.bank_code = bn01.bank_no "
			       + " and bn01.bank_no in ("+selectBank_no+")"
			       + condition.toString()				     
			       + " group by a03.bank_code,bn01.bank_name,a03.m_year,a03.m_month,wlx01.hsien_id,cd01.hsien_name,"+ori_field);								       
			for(Object obj : a03_tableList) {
				sqlCmdList.add(obj) ;
			}
			sqlCmdList.add(u_year) ;
			sqlCmdList.add(u_year) ;
			for(Object obj : conditionList) {
				sqlCmdList.add(obj);
			}
			if(!order.equals("")){	 
			    //100.05.12 fix 有挑選排序欄位年月時,查詢SQL error by 2295 		  			
			    //各別機構
			    sqlCmd.append(" order by "+order.replaceAll("order_YYMM", "a03.m_year,m_month")); 
			    sqlCmd_sum.append(" order by " + order.replaceAll("order_YYMM", "m_year,m_month") );			    			    
  				if( session.getAttribute("SortBy") != null && !((String)session.getAttribute("SortBy")).equals("")){	   
  		            sqlCmd.append( " " + ((String)session.getAttribute("SortBy")));	  		            
  		            sqlCmd_sum.append( " " + ((String)session.getAttribute("SortBy")));	  		            
  		         }
  		    }else{
  		       sqlCmd.append(" order by a03.bank_code,a03.m_year,a03.m_month ");
  		       sqlCmd_sum.append(" order by cd01.FR001W_OUTPUT_ORDER,a03_operation.HSIEN_ID,m_year,m_month ");
  		    }	       			      
            //System.out.println("sqlCmd="+sqlCmd);
            //讀取報表欄位名稱===================================================================================        	
        	Properties p = new Properties();
			p.load(new FileInputStream(Utility.getProperties("schemaDir")+System.getProperty("file.separator")+"A03_"+bank_type+"_column.TXT"));			
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
            //設定表頭===============================================================================
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
            reportUtil.createCell( wb, row, ( short )1, S_YEAR + "年" + S_MONTH + "月~"+E_YEAR + "年" + E_MONTH + "月", titleStyle );
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
            reportUtil.createCell( wb, row, ( short )1, "單位：新台幣"+Utility.getUnitName(Unit)+"、％", noBorderLeftStyle );                                                
            sheet.addMergedRegion( new Region( ( short )4, ( short )1,
                                               ( short )4,
                                               ( short )2) );          
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
                                               ( short )7,
                                               ( short )1) );                                                                     
            sheet.addMergedRegion( new Region( ( short )5, ( short )2,
                                               ( short )7,
                                               ( short )2) );                                              
            sheet.addMergedRegion( new Region( ( short )5, ( short )3,
                                               ( short )7,
                                               ( short )3) );                                                                                 
            row = sheet.createRow( ( short )5 );//大類表頭
            int columnIdx = 4;
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
            columnIdx = 4;          
            for(i=0;i<btnFieldList_data.size();i++){
               //System.out.println("["+i+"]i="+(String)((List)btnFieldList_data.get(i)).get(1));
               //System.out.println("columnIdx="+columnIdx);
               detail_column = (List)h_column.get(((List)btnFieldList_data.get(i)).get(0));//取出該大項的細類
               //設定細項表頭欄位
               for(j=0 ;j<detail_column.size();j++){                    
                  //System.out.println((String)detail_column.get(j)+"="+Utility.ISOtoBig5((String)prop_column_name.get((String)detail_column.get(j))));                                           
                  reportUtil.createCell( wb, row, ( short )columnIdx, Utility.ISOtoUTF8((String)prop_column_name.get((String)detail_column.get(j))), columnStyle );               
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
            wb.setRepeatingRowsAndColumns(0, 1, columnLength+2, 1, 7); //設定表頭 為固定 先設欄的起始再設列的起始
  			  			
  			//System.out.println("DS003W_Excel.sqlCmd="+sqlCmd);	   
  			System.out.println("ori_field="+ori_field); 
  			List dbData = null;
  			if(hasBankListALL.equals("false")){  
  				System.out.println("print......hasBankListALL = false") ;
  			  //System.out.println("tt1="+sqlCmd);			  
  			  //System.out.println("tt1="+ori_field);			  
			  //dbData = DBManager.QueryDB(sqlCmd,ori_field+"m_yearmonth");   
			  dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),sqlCmdList,ori_field+"m_yearmonth") ;
			  //System.out.println("DS003W_Excel.sqlCmd="+sqlCmd);	          
			}else{//95.09.04 add 金融機構代號=ALL全部  		
				System.out.println("print......hasBankListALL = true") ;
			    //dbData = DBManager.QueryDB(sqlCmd_sum,ori_field+"m_yearmonth");
			    dbData = DBManager.QueryDB_SQLParam(sqlCmd_sum.toString(),sqlCmd_sumList,ori_field+"m_yearmonth") ;
			    
			  //System.out.println("DS003W_Excel.sqlCmd_sum="+sqlCmd_sum);	          
			}
			
			short rowNo = ( short )8;//資料起始列     
			//無資料時,顯示訊息========================================================================			
			if(dbData == null || dbData.size() == 0){
			   	row = sheet.createRow( rowNo );     
                row.setHeight((short) 0x120); 
                reportUtil.createCell( wb, row, ( short )1,"無資料",noBorderDefaultStyle ); 
                sheet.addMergedRegion( new Region( ( short )8, ( short )1,
                                               ( short )8,
                                               ( short )(columnLength+5)) );
			}else{
			//有Data時,將DBData寫入===============================================================================================      
            acc_code_row = sheet.getRow(7);
            short lastCellNum = acc_code_row.getLastCellNum();
            //System.out.println("lastCellNum="+lastCellNum);                        
            columnIdx = 1;       
            double amt_d = 0.0;                 
            float amt_f = 0; 
            String amt_type="";
            String amt="";
            String acc_code = "";
            String prtbank_code="";
            for(i=0;i<dbData.size();i++){            
            	bean = (DataObject)dbData.get(i) ;
                acc_code = "";
                row = sheet.createRow( rowNo );  
                //System.out.println("rowNo="+rowNo);   
                row.setHeight((short) 0x120);    
                //System.out.println("bank_code="+(String) ((DataObject) dbData.get(i)).getValue("bank_code"));                
                /*95.12.05 fix 讓User可根據Excel做其他運用.單位代號.機構名稱repeat
                if(prtbank_code.equals((String) ((DataObject) dbData.get(i)).getValue("bank_code"))){
                   reportUtil.createCell( wb, row, ( short )columnIdx, " ", defaultStyle );//單位代號
                   columnIdx++;
                   reportUtil.createCell( wb, row, ( short )columnIdx, " ", defaultStyle );//機構名稱
                   columnIdx++;
                }else{*/
                   reportUtil.createCell( wb, row, ( short )columnIdx, (String) bean.getValue("bank_code"), defaultStyle );//單位代號
                   columnIdx++;
                   reportUtil.createCell( wb, row, ( short )columnIdx, (String) bean.getValue("bank_name"), defaultStyle );//機構名稱
                   columnIdx++;
                //}
                //95.12.01 add 查詢年月
                reportUtil.createCell( wb, row, ( short )columnIdx, ((bean).getValue("m_yearmonth")).toString(), defaultStyle );//查詢年月
                columnIdx++;
                for(int cellIdx =4;cellIdx < (new Short(lastCellNum)).intValue();cellIdx++){                    
                     amt="";
                     amt_type = "";
                     cell = acc_code_row.getCell((short)cellIdx);                    
                     acc_code = cell.getStringCellValue().toLowerCase();                     
                     //System.out.println("acc_code="+acc_code);                      
                     if(hasBankListALL.equals("true")){  		
                        amt =(bean.getValue(acc_code)).toString();
                     }else{                     
                        if((cell.getStringCellValue()).indexOf("field") == -1){                                                                          
                           amt = (bean.getValue("amt"+acc_code)).toString();                                                
                        }else{                  
                           amt = (bean.getValue(acc_code)).toString();
                        }
                     }                     
                     amt_type = (bean.getValue("type_"+acc_code)).toString();                                                             
                     //System.out.println("amt="+amt);
                     //System.out.println("amt_type="+amt_type);
                     
                     if(acc_code.substring(acc_code.length()-1).equals("p")){//95.10.02 add acc_code若為P結尾為利率
                        amt = Utility.getPercentNumber(amt);
                     }else if(!amt_type.equals("4")){//該值不為利率時.再除以單位4捨五入
                        amt = Utility.getRound(amt,Unit);                        
                     }                     
                    
                     amt = Utility.setCommaFormat(amt);
                     reportUtil.createCell( wb, row, ( short )columnIdx, amt, rightStyle );               
                     columnIdx ++;
                }
                columnIdx = 1;
                rowNo++;
                //prtbank_code = (String) ((DataObject) dbData.get(i)).getValue("bank_code");
            }
            
            }//end of 有data  
            
            //95.10.02 add 合併acc_code的欄位名稱=================================================================           
            columnIdx = 4;
            for(i=0;i<btnFieldList_data.size();i++){
                //System.out.println("["+i+"]i="+(String)((List)btnFieldList_data.get(i)).get(0));                
                //System.out.println("columnIdx="+columnIdx);                
                //設定表頭欄位.把中間值的acc_code合併成一個欄位只顯示中文名稱=====================================
                 detail_column = (List)h_column.get(((List)btnFieldList_data.get(i)).get(0));//取出該大項的細類
               	//設定細項表頭欄位
               	for(j=0 ;j<detail_column.size();j++){ 
                    //System.out.println((String)detail_column.get(j)+"="+Utility.ISOtoBig5((String)prop_column_name.get((String)detail_column.get(j))));                                           
                  	 //95.10.02 add 合併acc_code長度為大於6的.欄位名稱
                  	if(((String)detail_column.get(j)).length() > 6 ){//科目代號:6碼
                  	    row = sheet.getRow(6);
                  	    reportUtil.createCell( wb, row, ( short )columnIdx,Utility.ISOtoUTF8((String)prop_column_name.get((String)detail_column.get(j))), columnStyle );               
                  	    sheet.addMergedRegion( new Region( ( short )6, ( short )columnIdx,
                                                           ( short )7,
                                               			   ( short )columnIdx) );                                              
                  		columnIdx ++;
               		}else{ //95.10.02 add 合併acc_code長度為=6的.欄位名稱
               		    row = sheet.getRow(5);
               		    reportUtil.createCell( wb, row, ( short )columnIdx,Utility.ISOtoUTF8((String)prop_column_name.get((String)detail_column.get(j))), columnStyle );               
               		    sheet.addMergedRegion( new Region( ( short )5, ( short )columnIdx,
                                                           ( short )6,
                                               			   ( short )columnIdx) );                                              
                  		columnIdx ++;
               		}    
                }                                              
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