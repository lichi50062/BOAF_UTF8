<%
// 94.09.19 create by 2295
// 96.03.22 fix 取得lasthsien_div / lasthsien_id by 2295
// 98.07.20 add 990611對政府部門授信金額/990610-990611非會員放款扣除對政府部門授信金額/990140月平均存款總額 by 2295
// 99.01.25 fix 計算上年度信用部決算淨值/全体農(漁)會決算淨值時,需扣除992810投資全國農業金庫尚未攤提之損失[99.1月起] by 2295
//              套用新的A02_6_detail_column_9901.TXT/A02_7_column_9901.TXT             
// 99.05.13 fix 縣市合併 & sql injection by 2808
//101.08.23 add 違反信用部固定資產淨額不得超過上年度信用部決算淨值不在此限的原因項目(990812/990813/990814合併顯示) by 2295
//				field_990812_990813_990814的值為123時,表示(一)990812/(二)990813/(三)990814都有勾選 by 2295
//102.01.16 add (四)信用部逾放比< 1%  且 BIS > 10%且備抵呆帳覆蓋率高100%,已申請經主管機關同意者,得不超過200%
//          add (六)逾放比低於2%且資本適足率高8%,已申請經主管機關同意者,得不超過150% by 2295
//104.02.13 add 備呆占放款比率 by 2295
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
	String actMsg = "";
	FileOutputStream fileOut = null;      	
    HSSFCellStyle defaultStyle;
    HSSFCellStyle rightStyle;
    HSSFCellStyle noBorderDefaultStyle;
    HSSFCellStyle noBorderLeftStyle;
    HSSFCellStyle noBorderCenter_Red_UnderlineStyle;    
    HSSFCellStyle noBorderRight_Red_UnderlineStyle;    
    HSSFCellStyle titleStyle;
	HSSFCellStyle columnStyle;
	HSSFCellStyle column_LeftStyle;
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
    String acc_div = "";//1.資產負債表:2.損益表
    String Unit = "";//列印單位
    String S_YEAR = "";//年
    String E_YEAR = "";//年
    String S_MONTH = "";//月
    String u_year = "99" ;//判斷縣市合併用
    String cdTable = "cd01_99" ;//判斷縣市合併用
    List BankList_data = null;//儲存bank_code/bank_name的集合
    List btnFieldList_data = null;
    List SortList_data = null;    
	int i = 0;
	int j= 0;
	int detailidx= 0;
	String lguser_name = "測試使用者";
	String bank_type="";
	String hasBankListALL="false";
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
			titleName += "法定比率資料";
			//年
			if(session.getAttribute("S_YEAR") != null && !((String)session.getAttribute("S_YEAR")).equals("")){
		  		S_YEAR = (String)session.getAttribute("S_YEAR");	
		  	}
			if(S_YEAR!=null && Integer.parseInt(S_YEAR)>99) {
				u_year = "100" ;
				cdTable = "cd01" ;
			}
			//月
			if(session.getAttribute("S_MONTH") != null && !((String)session.getAttribute("S_MONTH")).equals("")){
		  		S_MONTH = (String)session.getAttribute("S_MONTH");		  				   
			}
			//金額單位
			if(session.getAttribute("Unit") != null && !((String)session.getAttribute("Unit")).equals("")){
		  		Unit = (String)session.getAttribute("Unit");		  				   
			}
			
			//讀取欄位大類所包含的中類===================================================================================        	
        	Properties prop_detail_column = new Properties();
			prop_detail_column.load(new FileInputStream(Utility.getProperties("schemaDir")+System.getProperty("file.separator")+"A02_"+bank_type+"_detail.TXT"));									
			//取出欄位中類將資料存入MAP-->key=大類acc_code,value=中類acc_code=============================================================
			HashMap h_column_detail = new HashMap();//儲存column大類,及其中類的acc_code			
			List detail_column_detail = new LinkedList();
			String column_tmp_detail = "";						
			for(i=0;i<btnFieldList_data.size();i++){			    
			    column_tmp_detail = "";
			    column_tmp_detail = (String)prop_detail_column.get((String)((List)btnFieldList_data.get(i)).get(0));
			    System.out.println("column_tmp_detail="+column_tmp_detail);
			    if(!column_tmp_detail.equals("")){
			        detail_column_detail = Utility.getStringTokenizerData(column_tmp_detail,"+");
			        //System.out.println((String)((List)btnFieldList_data.get(i)).get(0)+"="+detail_column_detail);			        		      
			        h_column_detail.put((String)((List)btnFieldList_data.get(i)).get(0),detail_column_detail);//大類->中類			       			      
			    }
			}
			System.out.println(Utility.ISOtoUTF8("大類->中類")+".h_column_detail.size()="+h_column_detail.size());
			//System.out.println("view dtat");
			//System.out.println(h_column_detail.get("A02_1"));
			//System.out.println(h_column_detail.get("A02_2"));
			
			//讀取欄位中類所包含的細項===================================================================================        	
			Properties prop_column = new Properties();
			prop_column.load(new FileInputStream(Utility.getProperties("schemaDir")+System.getProperty("file.separator")+"A02_"+bank_type+"_detail_column.TXT"));									
			//取出欄位細項將資料存入MAP-->key=中類acc_code,value=細項acc_code=============================================================
			HashMap h_column = new HashMap();//儲存column中類,及其細項的acc_code			
			List detail_column = new LinkedList();
			String column_tmp = "";			
			String selectacc_code = "";//選取的detail科目代號
			int columnLength=0;//column個數			
			for(i=0;i<btnFieldList_data.size();i++){			    
			    column_tmp = "";
			    //System.out.println("test111="+(String)((List)btnFieldList_data.get(i)).get(0));
			    detail_column_detail = (List)h_column_detail.get((String)((List)btnFieldList_data.get(i)).get(0));//取得大類.包含的中類
			    for(detailidx=0;detailidx<detail_column_detail.size();detailidx++){
			        column_tmp = (String)prop_column.get((String)detail_column_detail.get(detailidx));//取得中類.包含的細項
			        //System.out.println("column_tmp="+column_tmp);
			        if(!column_tmp.equals("")){
			           detail_column = Utility.getStringTokenizerData(column_tmp,"+");
			           //System.out.println(detail_column);
			           if(detail_column != null && detail_column.size() != 0){               
			              columnLength += detail_column.size();//累加總欄位個數
              		      for(j=0;j<detail_column.size();j++){
            	 		      selectacc_code +="'"+(String)detail_column.get(j)+"'";            	
            	 		      if(j < detail_column.size()-1){            	 		         
            	 		         selectacc_code +=",";
            	 		      }   
               		      }                              		   
               		      //System.out.println("select acc_code="+selectacc_code);
            	       }			      			        
			       }
			       h_column.put((String)detail_column_detail.get(detailidx),detail_column);			       			       
			       if(detailidx < detail_column_detail.size()-1){
			          selectacc_code +=",";			       
			       }   
			    }//end of 中類			     
			    if(i < btnFieldList_data.size()-1){			       
			       selectacc_code +=",";			       
			    }   
			}//end of 大類
			System.out.println("select acc_code="+selectacc_code);
			System.out.println(Utility.ISOtoUTF8("中類->細項")+".h_column.size()="+h_column.size());
        	//讀取報表欄位名稱===================================================================================        	
        	Properties prop_column_name = new Properties();        	
			prop_column_name.load(new FileInputStream(Utility.getProperties("schemaDir")+System.getProperty("file.separator")+"A02_"+bank_type+"_column"+((Integer.parseInt(S_YEAR+S_MONTH) >= 9901)?"_9901":"")+".TXT"));			
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
  			/*
  			//各別機構  			
  			select wlx01.hsien_id,cd01.hsien_name,
  			       decode(wlx01.hsien_div_1,'1','A2','2','A1','---') as hsien_div,
  			       decode(a02.violate,'Y','***','---') as violoate,
  			       cd01.hsien_name,a02.bank_code,bn01.bank_name,a02.acc_code,
  			decode(a02.type,'4',amt,round(amt/1,0)) as amt
			from (
					 select * from 
					( 
					select m_year,m_month,bank_code,acc_code,type,amt,violate from a02_operation
  					  where m_year=94 and m_month=8
  					  and bank_code in  ('6030016')
  					  and acc_code in('field_990210')					
					) a02_operation union					
				    (
					 select m_year,m_month,bank_code,acc_code,
					        decode(acc_code,'99141Y','4','0') as type,amt,
					        '' as violate from a02
  					  where   a02.bank_code in  ('6030016') 
   					  and a02.m_year = 94 and a02.m_month = 8 
   					  and a02.acc_code in ('990110','990310','990510','990610','991010','99141Y')					
					) union
				    (
					 select m_year,m_month,bank_code,acc_code,'4' as type,amt,'' as violate from a05
  					  where   a05.bank_code in  ('6030016') 
   					  and a05.m_year = 94 and a05.m_month = 8 
   					  and a05.acc_code in ('91060P')
					) union
					(
					select m_year,m_month,bank_code,acc_code,type,amt,'' as violate from a01_operation
  					  where m_year=94 and m_month=8
  					  and bank_code in  ('6030016')
  					  and acc_code in('field_over_rate')					
					) 
				)a02				 
			,bn01 left join wlx01                 
			on bn01.bank_no = wlx01.bank_no
			left join cd01 on wlx01.HSIEN_ID = cd01.HSIEN_ID  
			where a02.bank_code = bn01.bank_no
			and bn01.bank_no in ('6030016') and bn01.bn_type <> '2' 
			and a02.bank_code = bn01.bank_no
			order by wlx01.hsien_id,a02.bank_code asc	
			//縣市別小計.無縣市別統計			
  			*/
  			
            String column = "";//選取欄位           
            String condition = "";//其他條件            
			condition += " and bn01.bn_type "+(CANCEL_NO.equals("N")?"<>":"=")+" '2' and a02.bank_code = bn01.bank_no ";		  	 
			//======================================================
            //String sqlCmd="";
			StringBuffer sqlCmd = new StringBuffer() ;
			List sqlCmdList = new ArrayList() ;
            String sqlCmd_sum="";//縣市別小計
            //String a02_table="";//a02
            StringBuffer a02_table = new StringBuffer() ;
            List a02_tableList = new ArrayList() ;
            String haveViolate="false";//有1~13項時,才要加91060P.field_over_rate
            for(i=0;i<btnFieldList_data.size();i++){
                if(((String)((List)btnFieldList_data.get(i)).get(0)).indexOf("A02_") != -1){
                   haveViolate="true";
                }
            }    
            System.out.println("haveViolate="+haveViolate);                      				   
			a02_table.append(" ( "
					  + "   select * from "
					  + "            ( select m_year,m_month,bank_code,acc_code,type,amt,violate from a02_operation "
					  + "              where m_year = ?"
				      + "              and   m_month = ?"
			          + "              and   bank_code in ("+selectBank_no+")"				      
				      + "              and   acc_code in ("+selectacc_code+")"      					  
					  + "             )a02_operation union "					
				      + "            ( select m_year,m_month,bank_code,acc_code,"
				      + "              decode(acc_code,'99141Y','4','0') as type,amt,'' as violate  from a02 "
				      + "              where m_year = ? "
				      + "              and   m_month = ? "
			          + "              and   bank_code in ("+selectBank_no+")"				      
				      + "              and   acc_code in ("+selectacc_code+") and acc_code not in (?, ? )"      
				      + "			  )union "
				      //G.上年度農會決算淨值(扣除未攤銷全國農業金庫股票之損失)(990320)-(992810)
				      + "			  ( select m_year,m_month,bank_code,'990320','0',"		   		  
				      + "				       sum(decode(acc_code,'990320',amt,'0'))  - sum(decode(acc_code,'992810',amt,'0')) as amt ,'' "
				      + "			    from (select m_year,m_month,bank_code,acc_code,amt from a02 " 
				      + "                     where m_year =  ? "
				      + "                     and   m_month = ? "
			          + "                     and   bank_code in ("+selectBank_no+")"				      
				      + "                     and   acc_code in (?)"    
				      + "					  union "
				      + "			          select * from a99 "
				      + "                     where m_year = ?"
				      + "                     and   m_month = ?"
			          + "                     and   bank_code in ("+selectBank_no+")"				      
				      + "                     and   acc_code in (?)"    
				      + "				     )a02"  
				      + "              group by m_year,m_month,bank_code " 		
				      + "			  )union "
				      //S.上年度信用部決算淨值(扣除未攤銷全國農業金庫股票之損失)(990230)-(992810)
				      + "			  ( select m_year,m_month,bank_code,'990230','0',"		   		  
				      + "				       sum(decode(acc_code,'990230',amt,'0'))  - sum(decode(acc_code,'992810',amt,'0')) as amt ,'' "
				      + "			    from (select m_year,m_month,bank_code,acc_code,amt from a02 " 
				      + "                     where m_year = ?"
				      + "                     and   m_month = ?"
			          + "                     and   bank_code in ("+selectBank_no+")"				      
				      + "                     and   acc_code in (?)"    
				      + "					  union "
				      + "			          select * from a99 "
				      + "                     where m_year = ?"
				      + "                     and   m_month = ?"
			          + "                     and   bank_code in ("+selectBank_no+")"				      
				      + "                     and   acc_code in (?)"    
				      + "				     )a02"  
				      + "              group by m_year,m_month,bank_code " 		
					  + "             )union " //非會員放款扣除對政府部門授信金額 990610-990611
					  + "			  ( select m_year,m_month,bank_code,'990610-990611','0',"		   		  
				      + "				       sum(decode(acc_code,'990610',amt,'0'))  - sum(decode(acc_code,'990611',amt,'0')) as amt ,''"
					  + " 				from a02 "      
					  + "              where m_year = ?"
				      + "              and   m_month = ?"
			          + "              and   bank_code in ("+selectBank_no+")"				      
				      + "              and acc_code in (?,?)"      					              
					  + "			   group by m_year,m_month,bank_code"      
					  + "			 ) ");
			a02_tableList.add(S_YEAR) ;
			a02_tableList.add(String.valueOf(Integer.parseInt(S_MONTH))) ;
			a02_tableList.add(S_YEAR) ;
			a02_tableList.add(String.valueOf(Integer.parseInt(S_MONTH))) ;
			a02_tableList.add("990320") ;
			a02_tableList.add("990230") ;
			a02_tableList.add(S_YEAR) ;
			a02_tableList.add(String.valueOf(Integer.parseInt(S_MONTH))) ;
			a02_tableList.add("990320") ;
			a02_tableList.add(S_YEAR) ;
			a02_tableList.add(String.valueOf(Integer.parseInt(S_MONTH))) ;
			a02_tableList.add("992810") ;
			a02_tableList.add(S_YEAR) ;
			a02_tableList.add(String.valueOf(Integer.parseInt(S_MONTH))) ;
			a02_tableList.add("990230") ;
			a02_tableList.add(S_YEAR) ;
			a02_tableList.add(String.valueOf(Integer.parseInt(S_MONTH))) ;
			a02_tableList.add("992810") ;
			a02_tableList.add(S_YEAR) ;
			a02_tableList.add(String.valueOf(Integer.parseInt(S_MONTH))) ;
			a02_tableList.add("990610") ;
			a02_tableList.add("990611") ;			
           if(haveViolate.equals("true")){					  
              a02_table.append("         union "
				      + "            ( select m_year,m_month,bank_code,acc_code,'4' as type,amt,'' as violate from a05 "
				      + "              where m_year = ?"
				      + "              and   m_month = ?"
			          + "              and   bank_code in ("+selectBank_no+")"				      
				      + "              and   acc_code in (?)"      					  					    					 					    					  
					  + "            ) union "
					  + "            ( select m_year,m_month,bank_code,acc_code,type,amt,'' as violate  from a01_operation "
					  + "              where m_year = ?"
				      + "              and   m_month = ?"
			          + "              and   bank_code in ("+selectBank_no+")"				      
				      + "              and   acc_code in ('field_over_rate','field_backup_over_rate','field_backup_credit_rate' )"      					  					    					 					    					  					  
					  + "            ) ");
              a02_tableList.add(S_YEAR) ;
  			  a02_tableList.add(String.valueOf(Integer.parseInt(S_MONTH))) ;
  			  a02_tableList.add("91060P") ;
  			  a02_tableList.add(S_YEAR) ;
			  a02_tableList.add(String.valueOf(Integer.parseInt(S_MONTH))) ;
			  //a02_tableList.add("field_over_rate,field_backup_over_rate") ;
              
		   }			  
		   a02_table.append(" )a02 ");								   
           sqlCmd.append(" select wlx01.hsien_id,cd01.hsien_name, "
                  + "        decode(wlx01.hsien_div_1,'1','A2','2','A1','---') as hsien_div,"
                  + "        decode(a02.violate,'Y','***','---') as violate,"
                  + "        cd01.hsien_name,a02.bank_code,bn01.bank_name,a02.acc_code,"
                  + "	     decode(a02.type,'4',amt,round(amt/?,0)) as amt " //95.09.05 add 除以單位  				
			      + " from ").append(a02_table.toString()).append(",");
           sqlCmd.append("(select * from bn01 where m_year=? )bn01 ");
           sqlCmd.append("left join (select * from wlx01 where m_year=? )wlx01 ");
           sqlCmd.append("on bn01.bank_no = wlx01.bank_no "
			      + " left join ").append(cdTable).append(" cd01 ");
           sqlCmd.append("on wlx01.HSIEN_ID = cd01.HSIEN_ID " 			      
			      + " where a02.bank_code = bn01.bank_no "
			      + " and bn01.bn_type "+(CANCEL_NO.equals("N")?"<>":"=")+" ? "
			      + " and bn01.bank_no in ("+selectBank_no+")");		
           sqlCmdList.add(Unit) ;
           
           for(Object obj : a02_tableList) {
        	   sqlCmdList.add(obj) ;
           }
           sqlCmdList.add(u_year) ;
           sqlCmdList.add(u_year) ;
           sqlCmdList.add("2") ;
			      //order by wlx01.hsien_id,a02.bank_code asc	
		   if(!order.equals("")){	   			
			    //各別機構
			    if(order.indexOf("wlx01.hsien_id") == -1 && order.indexOf("a02.bank_code") == -1){	  
			       //sqlCmd += " order by amt"; 
			    }else{			    
			       sqlCmd.append(" order by ").append(order) ;
			       if(order.equals("wlx01.hsien_id")){//95.09.06 add 只有縣市別sort時,再加上機構代號 by 2295	 
			          sqlCmd.append(",a02.bank_code");
			       }
			    }			    
  				//SoryBy=asc/desc
  				if( session.getAttribute("SortBy") != null && !((String)session.getAttribute("SortBy")).equals("")){	   
  		            sqlCmd.append(" " + ((String)session.getAttribute("SortBy")));	  		            
  		            //sqlCmd_sum += " " + ((String)session.getAttribute("SortBy"));	  		            
  		         }
  		    }else{
  		       sqlCmd.append(" order by wlx01.hsien_id,a02.bank_code");
  		       //sqlCmd_sum += " order by cd01.hsien_id";
  		    }	       			      
            //System.out.println("sqlCmd="+sqlCmd);
            //讀取報表欄位名稱===================================================================================        	
        	Properties p = new Properties();
			p.load(new FileInputStream(Utility.getProperties("schemaDir")+System.getProperty("file.separator")+"A02_"+bank_type+"_column"+((Integer.parseInt(S_YEAR+S_MONTH) >= 9901)?"_9901":"")+".TXT"));			
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
    		noBorderCenter_Red_UnderlineStyle = reportUtil.getCenter_Red_UnderlineStyle(wb);//95.09.21 add 無框內文置中.紅字+底線
    		noBorderRight_Red_UnderlineStyle = reportUtil.getRight_Red_UnderlineStyle(wb);//95.09.21 add 無框內文置右.紅字+底線
			reportUtil.setDefaultStyle(defaultStyle);
			reportUtil.setNoBorderDefaultStyle(noBorderDefaultStyle);			
    		titleStyle = reportUtil.getTitleStyle(wb); //標題用
    		columnStyle = reportUtil.getColumnStyle(wb);//報表欄位名稱用--有框內文置中			                                               
    		column_LeftStyle = reportUtil.getColumn_LeftStyle(wb);//報表欄位名稱用--有框內文置左	
    		noBoderStyle = reportUtil.getNoBoderStyle(wb);//無框置右			                                               
    		//============================================================================                        
            //設定表頭(資產負債表/損益表)===============================================================================
            row = sheet.createRow( ( short )1 );
            reportUtil.createCell( wb, row, ( short )1, titleName, titleStyle );
            
            for(i=2;i<columnLength+4;i++){
              reportUtil.createCell( wb, row, ( short )i, "", noBorderDefaultStyle );
            }            
            sheet.addMergedRegion( new Region( ( short )1, ( short )1,
                                               ( short )1,
                                               ( short )(columnLength+4)) );
            //列印年月=======================================================================================
            row = sheet.createRow( ( short )2 );            
            row.setHeight((short) 0x200);
            reportUtil.createCell( wb, row, ( short )1, S_YEAR + "年" + S_MONTH + "月", titleStyle );
            for(i=2;i<columnLength+4;i++){
               reportUtil.createCell( wb, row, ( short )i, "", noBorderDefaultStyle );
            }            
            sheet.addMergedRegion( new Region( ( short )2, ( short )1,
                                               ( short )2,
                                               ( short )(columnLength+4) ) );
            //======================================================================================================                                                                     
            row = sheet.createRow( ( short )3 );            
            String printTime = Utility.getDateFormat("  HH:mm:ss");
            String printDate = Utility.getDateFormat("yyyy/MM/dd");                                    
            reportUtil.createCell( wb, row, ( short )1, "列印日期："+Utility.getCHTdate(printDate, 1)+printTime, noBoderStyle );
            sheet.addMergedRegion( new Region( ( short )3, ( short )1,
                                               ( short )3,
                                               ( short )(columnLength+4) ) );            
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
                                               ( short )(columnLength+4) ) );            
            //報表欄位=======================================================================
            //列印單位代號+機構名稱            
            
            for(i=5;i<9;i++){
                row = sheet.createRow( ( short )i );
                reportUtil.createCell( wb, row, ( short )1, "違反(***)", columnStyle );  
                reportUtil.createCell( wb, row, ( short )2, "縣市別", columnStyle );  
                reportUtil.createCell( wb, row, ( short )3, "單位代號", columnStyle );               
                reportUtil.createCell( wb, row, ( short )4, "單位名稱", columnStyle );   
            } 
                    
            sheet.addMergedRegion( new Region( ( short )5, ( short )1,
                                               ( short )8,
                                               ( short )1) );                                                                     
            sheet.addMergedRegion( new Region( ( short )5, ( short )2,
                                               ( short )8,
                                               ( short )2) );                                              
            sheet.addMergedRegion( new Region( ( short )5, ( short )3,
                                               ( short )8,
                                               ( short )3) );                                              
            sheet.addMergedRegion( new Region( ( short )5, ( short )4,
                                               ( short )8,
                                               ( short )4) );                                                                                             
                                                         
            row = sheet.createRow( ( short )5 );//大類表頭
            int columnIdx = 5;
            int detailsize = 0;
            for(i=0;i<btnFieldList_data.size();i++){
                //System.out.println("["+i+"]i="+(String)((List)btnFieldList_data.get(i)).get(1));
                //System.out.println("columnIdx="+columnIdx);
                //設定表頭大類欄位
                detailsize=0;
                detail_column_detail = (List)h_column_detail.get((String)((List)btnFieldList_data.get(i)).get(0));//取得大類.包含的中類
			    for(detailidx=0;detailidx<detail_column_detail.size();detailidx++){
                    detailsize += ((List)h_column.get(((String)detail_column_detail.get(detailidx)))).size();//取得中類的細項			
                }
                
                for(j=columnIdx;j<detailsize + columnIdx;j++){                  
                  reportUtil.createCell( wb, row, ( short )j, (String)((List)btnFieldList_data.get(i)).get(1), columnStyle );               
                }
                sheet.addMergedRegion( new Region( ( short )5, ( short )columnIdx,
                                                   ( short )5,
                                                   ( short )(detailsize + columnIdx - 1)) );                                              
                columnIdx += detailsize;                                             
            }
            row = sheet.createRow( ( short ) 6);//中類表頭            
            row.setHeightInPoints(68);//設定大類表頭高度
            columnIdx = 5; 
            for(i=0;i<btnFieldList_data.size();i++){
               //System.out.println("["+i+"]i="+(String)((List)btnFieldList_data.get(i)).get(1));
               //System.out.println("columnIdx="+columnIdx);
               //設定表頭中類欄位               
               detail_column_detail = (List)h_column_detail.get((String)((List)btnFieldList_data.get(i)).get(0));//取得大類.包含的中類
			   for(detailidx=0;detailidx<detail_column_detail.size();detailidx++){
                   //System.out.println("detail.size="+((List)h_column.get(((String)detail_column_detail.get(detailidx)))).size());//取得中類的細項			
                   for(j=columnIdx;j<((List)h_column.get(((String)detail_column_detail.get(detailidx)))).size() + columnIdx;j++){
                      //System.out.println("detail.column="+(String)detail_column_detail.get(detailidx));
                      //System.out.println("detail.name="+Utility.ISOtoBig5((String)prop_column_name.get((String)detail_column_detail.get(detailidx))));
                      reportUtil.createCell( wb, row, ( short )j, Utility.ISOtoUTF8((String)prop_column_name.get((String)detail_column_detail.get(detailidx))), columnStyle );                              
                   }
                   sheet.addMergedRegion( new Region( ( short )6, ( short )columnIdx,
                                                      ( short )6,
                                                      ( short )(((List)h_column.get(((String)detail_column_detail.get(detailidx)))).size() + columnIdx - 1)) );                                              
                   columnIdx +=  ((List)h_column.get(((String)detail_column_detail.get(detailidx)))).size(); 
               }                                           
            }
            
            row = sheet.createRow( ( short ) 7);//細項表頭
            row.setHeightInPoints(100);//設定細項表頭高度
            columnIdx = 5;          
            for(i=0;i<btnFieldList_data.size();i++){
               //System.out.println("["+i+"]i="+(String)((List)btnFieldList_data.get(i)).get(1));
               //System.out.println("columnIdx="+columnIdx);
               detail_column_detail = (List)h_column_detail.get((String)((List)btnFieldList_data.get(i)).get(0));//取得大類.包含的中類
               for(detailidx=0;detailidx<detail_column_detail.size();detailidx++){//中類
                   //System.out.println("detail.size="+((List)h_column.get(((String)detail_column_detail.get(detailidx)))).size());//取得中類.包含細項的size	
                   detail_column = (List)h_column.get(((String)detail_column_detail.get(detailidx)));//取得中類的細項		
                   //設定細項表頭欄位
                   for(j=0 ;j<detail_column.size();j++){                    
                       //System.out.println((String)detail_column.get(j)+"="+Utility.ISOtoBig5((String)prop_column_name.get((String)detail_column.get(j))));                                           
                       reportUtil.createCell( wb, row, ( short )columnIdx, Utility.ISOtoUTF8((String)prop_column_name.get((String)detail_column.get(j))), columnStyle );               
                       columnIdx ++;
                   }//end of 細項                        
               }//end of中類                        
            }//end of 大類
            
            row = sheet.createRow( ( short ) 8);//細項-科目代號
            columnIdx = 5;  
            for(i=0;i<btnFieldList_data.size();i++){
               //System.out.println("["+i+"]i="+(String)((List)btnFieldList_data.get(i)).get(1));
               //System.out.println("columnIdx="+columnIdx);
               detail_column_detail = (List)h_column_detail.get((String)((List)btnFieldList_data.get(i)).get(0));//取得大類.包含的中類
               for(detailidx=0;detailidx<detail_column_detail.size();detailidx++){//中類
                   //System.out.println("detail.size="+((List)h_column.get(((String)detail_column_detail.get(detailidx)))).size());//取得中類.包含細項的size	
                   detail_column = (List)h_column.get(((String)detail_column_detail.get(detailidx)));//取得中類的細項		
                   //設定細項表頭欄位
                   for(j=0 ;j<detail_column.size();j++){                    
                       //System.out.println((String)detail_column.get(j)+"="+Utility.ISOtoBig5((String)prop_column_name.get((String)detail_column.get(j))));                                           
                       reportUtil.createCell( wb, row, ( short )columnIdx, (String)detail_column.get(j), columnStyle );               
                       columnIdx ++;
                   }//end of 細項                        
               }//end of中類                        
            }//end of 大類                    
            
            
            //wb.setRepeatingRowsAndColumns( 0, 1, 9, 1, 3 ); //設定表頭 為固定 先設欄的起始再設列的起始              
            wb.setRepeatingRowsAndColumns(0, 1, columnLength+2, 1, 8); //設定表頭 為固定 先設欄的起始再設列的起始
              						
  			//System.out.println("DS001W_Excel.sqlCmd="+sqlCmd);	   
  			
  			List dbData = null;
  			if(hasBankListALL.equals("false")){
			  //dbData = DBManager.QueryDB(sqlCmd,"amt");
			  System.out.println("Query====================================") ;
			  dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),sqlCmdList,"amt") ;
			  //System.out.println("DS001W_Excel.sqlCmd="+sqlCmd);	          
			}else{//95.09.04 add 金融機構代號=ALL全部  
			  dbData = DBManager.QueryDB_SQLParam(sqlCmd_sum.toString(),null,"amt");     
			  //System.out.println("DS001W_Excel.sqlCmd_sum="+sqlCmd_sum);	          
			}
			
			short rowNo = ( short )9;//資料起始列     
			//無資料時,顯示訊息========================================================================			
			if(dbData == null || dbData.size() == 0){
			   	row = sheet.createRow( rowNo );     
                row.setHeight((short) 0x120); 
                reportUtil.createCell( wb, row, ( short )1,"無資料" ,noBorderDefaultStyle ); 
                sheet.addMergedRegion( new Region( ( short )9, ( short )1,
                                               ( short )9,
                                               ( short )(columnLength+4)) );
			}else{
			System.out.println("dbData.size()="+dbData.size());
			//有Data時=========================================================================================           						
			//將dbData存入A01Map===============================================================================
            HashMap A01Map = new HashMap();
            Properties A01Data = new Properties(); 			
            String bank_code="";
            String bank_name="";
            String hsien_div="";
            String hsien_id="";
            String lastbank_code="";
            String lastbank_name="";
            String lasthsien_div="";
            String lasthsien_id="";
            String acc_code="";
            String amt="";
            String field_990812_990813_990814="";
            String violate="";            
            String violate_mark="---";         
            List bank_code_List = new LinkedList();
            List bank_List = new LinkedList();            
            bank_code = (String) ((DataObject) dbData.get(0)).getValue("bank_code");
            lastbank_code = (String) ((DataObject) dbData.get(0)).getValue("bank_code");
            lastbank_name = (String) ((DataObject) dbData.get(0)).getValue("bank_name");
            lasthsien_div = (String) ((DataObject) dbData.get(0)).getValue("hsien_div");//96.03.22 fix 取得lasthsien_div              
            lasthsien_id = (String) ((DataObject) dbData.get(0)).getValue("hsien_name");//96.03.22 fix 取得lasthsien_id              
            for (i = 0; i < dbData.size(); i++) {
                bank_code = (String) ((DataObject) dbData.get(i)).getValue("bank_code");
                bank_name = (String) ((DataObject) dbData.get(i)).getValue("bank_name");
                hsien_div = (String) ((DataObject) dbData.get(i)).getValue("hsien_div");              
                hsien_id = (String) ((DataObject) dbData.get(i)).getValue("hsien_name");                             
                if(bank_code.equals(lastbank_code)){//與前一個bank_code相同時,儲存acc_code,amt			          			       			                          
                   acc_code = (String) ((DataObject) dbData.get(i)).getValue("acc_code");                   
                   amt = (((DataObject) dbData.get(i)).getValue("amt")).toString();
                   //101.08.23 add 990812/990813/990814.合併顯示 by 2295
                   //field_990812_990813_990814的值為123時,表示(一)990812/(二)990813/(三)990814都有勾選 by 2295
                   if("field_990812_990813_990814".equals(acc_code)){
                   	 //System.out.println("field_990812_990813_990814.amt="+amt);
                	 if(!amt.equals("")){
                	 	 field_990812_990813_990814 = "";
                	 	 if(amt.indexOf("1") != -1) field_990812_990813_990814 = "1";
                	 	 if(amt.indexOf("2") != -1){
                	 	 	field_990812_990813_990814 += field_990812_990813_990814.length() > 0 ?"、":"";
                	 	 	field_990812_990813_990814 += "2";
                	 	 }	 
                	 	 if(amt.indexOf("3") != -1){
                	 	 	field_990812_990813_990814 += field_990812_990813_990814.length() > 0 ?"、":"";
                	 	 	field_990812_990813_990814 += "3";
                	 	 }
                	 	
                	 	 amt = 	field_990812_990813_990814;
                	 	 //System.out.println("field_990812_990813_990814.amt="+amt);
                	 }	
                   }	
                   violate = (String) ((DataObject) dbData.get(i)).getValue("violate");  
                   System.out.println("acc_code="+acc_code);
                   System.out.println("violate='"+violate+"'");
                   if( acc_code.equals("field_V_X") || acc_code.equals("field_W_x") ||	acc_code.equals("field_a_X") ||                  
                       	acc_code.equals("field_b_x") || acc_code.equals("field_c_d") || acc_code.equals("field_f_g") )
                   {                
                      if(violate.equals("***")) violate_mark="***";                    
                      A01Data.setProperty(acc_code, violate);
                      //System.out.println("A02 set acc_code="+acc_code+":violate="+violate);   
                   }else{
                      if(violate.equals("***")) violate_mark="***";  
                      A01Data.setProperty(acc_code, amt);
                      A01Data.setProperty(acc_code+"_violate", violate);//95.09.21
                      //System.out.println("A02 set violate="+violate);
                      //System.out.println("A02 set acc_code="+acc_code+":amt="+amt);   
                   }   
                }else{                                   
                    //System.out.println("put "+lastbank_code);
                    bank_code_List.add(lastbank_code);
                    bank_code_List.add(lastbank_name);
                    bank_List.add(bank_code_List);
                    A01Data.setProperty("hsien_div", lasthsien_div);
                    A01Data.setProperty("hsien_id", lasthsien_id);
                    A01Data.setProperty("violate_mark", violate_mark);
                    A01Map.put(lastbank_code,A01Data);  
                    lastbank_code = bank_code;                                     
                    lastbank_name = bank_name;  
                    lasthsien_div = hsien_div;  
                    lasthsien_id = hsien_id;  
                    violate_mark="---";
                    A01Data = new Properties(); 
                    bank_code_List = new LinkedList();			
                    acc_code = (String) ((DataObject) dbData.get(i)).getValue("acc_code");
                    amt = (((DataObject) dbData.get(i)).getValue("amt")).toString();
                    violate = (String) ((DataObject) dbData.get(i)).getValue("violate");                                      
                    if( acc_code.equals("field_V_X") || acc_code.equals("field_W_x") ||	acc_code.equals("field_a_X") ||                  
                       	acc_code.equals("field_b_x") || acc_code.equals("field_c_d") || acc_code.equals("field_f_g") )
                    {                  
                       if(violate.equals("***")) violate_mark="***";      
                       A01Data.setProperty(acc_code, violate);
                       //System.out.println("A02 set acc_code="+acc_code+":violate="+violate);   
                    }else{
                       if(violate.equals("***")) violate_mark="***";  
                       A01Data.setProperty(acc_code, amt);
                       A01Data.setProperty(acc_code+"_violate", violate);//95.09.21
                       //System.out.println("A02 set violate="+violate);
                       //System.out.println("A02 set acc_code="+acc_code+":amt="+amt);   
                    }                       
                }                
            }
            //System.out.println("put last bank_code="+lastbank_code);
            A01Data.setProperty("hsien_div", lasthsien_div);
            A01Data.setProperty("hsien_id", lasthsien_id);
            A01Data.setProperty("violate_mark", violate_mark);
            A01Map.put(lastbank_code,A01Data);//將最後一筆寫入              
            bank_code_List.add(lastbank_code);
            bank_code_List.add(lastbank_name);
            bank_List.add(bank_code_List);
            System.out.println("A02Map.size="+A01Map.size());
            System.out.println("bank_List.size()="+bank_List.size());                                         			
            
            //===========================================================================================================
            //將DBData寫入===============================================================================================      
            acc_code_row = sheet.getRow(8);
            short lastCellNum = acc_code_row.getLastCellNum();
            //System.out.println("lastCellNum="+lastCellNum);                        
            columnIdx = 1;       
            double amt_d = 0.0;                 
            float amt_f = 0; 
            hsien_id="";
            for(i=0;i<bank_List.size();i++){            
                row = sheet.createRow( rowNo );  
                //System.out.println("rowNo="+rowNo);   
                row.setHeight((short) 0x120);    
                bank_code_List = (List)bank_List.get(i);             
                A01Data = (Properties)A01Map.get(bank_code_List.get(0));                               
                //System.out.println("get bank_code="+bank_code_List.get(0));                
                if(((String)A01Data.get("violate_mark")).equals("---")){                   
                   reportUtil.createCell( wb, row, ( short )columnIdx,"", defaultStyle );//違反(***)                
                   columnIdx++;
                }else{
                  reportUtil.createCell( wb, row, ( short )columnIdx,"***", defaultStyle );//違反(***)                
                   columnIdx++;
                }
                if(!hsien_id.equals((String)A01Data.get("hsien_id"))){                   
                   hsien_id = (String)A01Data.get("hsien_id");
                   reportUtil.createCell( wb, row, ( short )columnIdx,hsien_id, defaultStyle );//縣市別                   
                   columnIdx++;
                }else{
                   reportUtil.createCell( wb, row, ( short )columnIdx,"", defaultStyle );//縣市別                   
                   columnIdx++;
                }
                
                reportUtil.createCell( wb, row, ( short )columnIdx,(String)bank_code_List.get(0), defaultStyle );//單位代號
                columnIdx++;
                reportUtil.createCell( wb, row, ( short )columnIdx, (String)bank_code_List.get(1), defaultStyle );//機構名稱
                columnIdx++;
                for(int cellIdx =5;cellIdx < (new Short(lastCellNum)).intValue();cellIdx++){ //測試用                   
				//for(int cellIdx =5;cellIdx < (new Short(lastCellNum)).intValue();cellIdx++){ //農金局.正式機用                                   	
                     amt="";
                     cell = acc_code_row.getCell((short)cellIdx);                    
                     if((String)A01Data.get(cell.getStringCellValue()) != null){
                        amt = (String)A01Data.get(cell.getStringCellValue());
                     }   
                     //System.out.print("cell_acc_code="+cell.getStringCellValue() );
                     //System.out.println(":amt="+amt);
                     //if(amt.indexOf(".") == -1){//95.09.05 add 該值不為利率時.再除以單位4捨五入
                     //   amt = Utility.getRound(amt,Unit);                        
                     //}
                    
                     if(!amt.equals("***") && !amt.equals("---") && amt.indexOf("A") == -1 && amt.indexOf(".") == -1
                     && !(cell.getStringCellValue().trim().equals("91060P")) 
                     ){
                        amt = Utility.setCommaFormat(amt);                     
                     }
                     if(amt.equals("---")){
                        amt = "";
                     }   
                     if(((String)A01Data.get(cell.getStringCellValue()) != null) 
                     && (cell.getStringCellValue().trim().equals("91060P"))){//A05.91060P
                        //System.out.println("get Percent");
                        amt = Utility.getPercentNumber(amt);
                     }   
                     if(amt.equals("***")){
                        //System.out.println("set center red under");
                        reportUtil.createCell( wb, row, ( short )columnIdx, amt, noBorderCenter_Red_UnderlineStyle );               
                     }if(((String)A01Data.get(cell.getStringCellValue()+"_violate") != null) 
                      && (((String)A01Data.get(cell.getStringCellValue()+"_violate")).equals("***"))){
                        reportUtil.createCell( wb, row, ( short )columnIdx, amt, noBorderRight_Red_UnderlineStyle );               
                     }else{
                        reportUtil.createCell( wb, row, ( short )columnIdx, amt, rightStyle );               
                     }
                     columnIdx ++;
                }
                columnIdx = 1;
                rowNo++;
            }
            
            }//end of 有data   
            
            //95.09.21 設定表頭欄位.把中間值的acc_code合併成一個欄位只顯示中文名稱=======================================================================                        
            row = sheet.getRow(7);
            columnIdx = 5;    
                    
             for(i=0;i<btnFieldList_data.size();i++){
               //System.out.println("["+i+"]i="+(String)((List)btnFieldList_data.get(i)).get(1));
               //System.out.println("columnIdx="+columnIdx);
               detail_column_detail = (List)h_column_detail.get((String)((List)btnFieldList_data.get(i)).get(0));//取得大類.包含的中類
               for(detailidx=0;detailidx<detail_column_detail.size();detailidx++){//中類
                   //System.out.println("detail.size="+((List)h_column.get(((String)detail_column_detail.get(detailidx)))).size());//取得中類.包含細項的size	
                   detail_column = (List)h_column.get(((String)detail_column_detail.get(detailidx)));//取得中類的細項		
                   //設定細項表頭欄位                   
                   for(j=0 ;j<detail_column.size();j++){                    
                       //System.out.println((String)detail_column.get(j)+"="+Utility.ISOtoBig5((String)prop_column_name.get((String)detail_column.get(j))));                                                                  
                       if( ((String)detail_column.get(j)).equals("field_X") || ((String)detail_column.get(j)).equals("field_x") ||	
                           ((String)detail_column.get(j)).equals("field_d") || ((String)detail_column.get(j)).equals("field_g") || ((String)detail_column.get(j)).equals("A02_8_a")) 
                       {
                          reportUtil.createCell( wb, row, ( short )columnIdx, Utility.ISOtoUTF8((String)prop_column_name.get((String)detail_column.get(j))), column_LeftStyle );                                      
                       }else{                     
                          reportUtil.createCell( wb, row, ( short )columnIdx, Utility.ISOtoUTF8((String)prop_column_name.get((String)detail_column.get(j))), columnStyle );                                      
                       }   
                       System.out.println((String)prop_column_name.get((String)detail_column.get(j)));
                      
                       //設定細項表頭欄位.把中間值的acc_code合併成一個欄位只顯示中文名稱=====================================
                       //101.08.23 add 違反信用部固定資產淨額不得超過上年度信用部決算淨值不在此限的原因項目 
                       if(((String)prop_column_name.get((String)detail_column.get(j))).equals("違反信用部固定資產淨額不得超過上年度信用部決算淨值不在此限的原因項目：\n一、因購置或汰換安全維護或營業相關設備，經中央主管機關核准\n二、因固定資產重估增值\n三、因淨值降低") 
                       ){
                           row = sheet.getRow(5);
                           reportUtil.createCell( wb, row, ( short )columnIdx, Utility.ISOtoUTF8((String)prop_column_name.get((String)detail_column.get(j))), column_LeftStyle );                                      
                           sheet.addMergedRegion( new Region( ( short )5, ( short )columnIdx,
                           				                      ( short )8,
                                        			          ( short ) columnIdx ) ); 
						   row = sheet.getRow(7);                                        			               
                       }else if(((String)prop_column_name.get((String)detail_column.get(j))).equals("月平均放款總額")
                       || ((String)prop_column_name.get((String)detail_column.get(j))).equals("月平均存款總額")
                       || ((String)prop_column_name.get((String)detail_column.get(j))).equals("非會員存款總額")
                       || ((String)prop_column_name.get((String)detail_column.get(j))).equals("非會員無擔保消費性貸款")
                       || ((String)prop_column_name.get((String)detail_column.get(j))).equals("非會員授信總額")
                       || ((String)prop_column_name.get((String)detail_column.get(j))).equals("對負責人、各部門員工或與其負責人或辦理授信之職員有利害關係者擔保放款最高限額")
                       || ((String)prop_column_name.get((String)detail_column.get(j))).equals("最近決算年度")
                       || ((String)prop_column_name.get((String)detail_column.get(j))).equals("對政府部門授信金額") 
                       || ((String)prop_column_name.get((String)detail_column.get(j))).equals("非會員放款扣除對政府部門授信金額")                        
                       ){
                           sheet.addMergedRegion( new Region( ( short )5, ( short )columnIdx,
                           				                      ( short )7,
                                        			          ( short ) columnIdx ) );      
                       }else{                       
                            sheet.addMergedRegion( new Region( ( short )7, ( short )columnIdx,
                           				                       ( short )8,
                                        			           ( short ) columnIdx ) );                             
                       }                 			           
                       columnIdx ++;
                   }//end of 細項                                                                                                     
               }//end of中類                        
            }//end of 大類           
            //設定寬度============================================================                                   
            for ( i = 1; i <= columnLength+4; i++ ) {                
                if(i==4){
                  sheet.setColumnWidth( ( short )i,
                                        ( short ) ( 256 * ( 25 + 4 ) ) );//機構名稱
                }else{
                  sheet.setColumnWidth( ( short )i,
                                        ( short ) ( 256 * ( 15 + 4 ) ) );
                }                        
            }
			//======================================================================================
			
            //設定涷結欄位
            if(haveViolate.equals("false")){//隱藏違反(***)
                sheet.setColumnWidth( ( short )1,(short)0);//違反(***)
            }
            
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
        	System.out.println("DS002W_Report have erros:"+e.getMessage()) ; 
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