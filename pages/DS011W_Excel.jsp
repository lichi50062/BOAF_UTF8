<%
//95.11.02 create by 2295
//95.11.02 農.漁會.都使用同一個設定檔 by 2295
//95.12.06 add 增加年月區間,可根據年月做sort,拿掉可選欄位做sort by 2295
//99.05.24 fixed by 2808 縣市合併&sql injection 
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
	String acc_code="";
	String u_year = "99" ;
	String cd01Table = "cd01_99" ;
	StringBuffer bn01= new StringBuffer() ;
	StringBuffer wlx01 = new StringBuffer() ;
	try{
		
		bn01.append("(select BANK_NO,BANK_NAME,BN_TYPE,BANK_TYPE,ADD_USER,ADD_NAME,ADD_DATE,BANK_B_NAME,KIND_1,KIND_2,BN_TYPE2,EXCHANGE_NO from bn01 where m_year=? )") ;
		wlx01.append("(select BANK_NO ,ENGLISH,SETUP_APPROVAL_UNT,SETUP_DATE,SETUP_NO,CHG_LICENSE_DATE,");
		wlx01.append(" CHG_LICENSE_NO,CHG_LICENSE_REASON,START_DATE,BUSINESS_ID,HSIEN_ID,AREA_ID,ADDR,");
		wlx01.append(" TELNO,FAX,EMAIL,WEB_SITE,CENTER_FLAG,CENTER_NO,STAFF_NUM,IT_HSIEN_ID,IT_AREA_ID,");
		wlx01.append(" IT_ADDR,IT_NAME,IT_TELNO,AUDIT_HSIEN_ID,AUDIT_AREA_ID,AUDIT_ADDR,AUDIT_NAME,AUDIT_TELNO,FLAG,OPEN_DATE,M2_NAME,");
		wlx01.append(" HSIEN_DIV_1,CANCEL_NO,CANCEL_DATE from wlx01 where m_year =? )");
		
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

			titleName += "支票存款資料";
			//年
			if(session.getAttribute("S_YEAR") != null && !((String)session.getAttribute("S_YEAR")).equals("")){
		  		S_YEAR = (String)session.getAttribute("S_YEAR");
		  		if(Integer.parseInt(S_YEAR)>99) {
		  			u_year = "100" ;
		  			cd01Table = "cd01" ;
		  		}
			}
			//年
			if(session.getAttribute("E_YEAR") != null && !((String)session.getAttribute("E_YEAR")).equals("")){
		  		E_YEAR = (String)session.getAttribute("E_YEAR");
			}
			//95.12.05 增加年月區間
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
			prop_column.load(new FileInputStream(Utility.getProperties("schemaDir")+System.getProperty("file.separator")+"WLX07_M_CHECKBANK_detail.TXT"));
			//=======================================================================================================================
			//取出欄位細項將資料存入MAP-->key=大類acc_code,value=細項acc_code=============================================================
			HashMap h_column = new HashMap();//儲存column大類,及其細項的acc_code
			List detail_column = new LinkedList();
			String column_tmp = "";
			String selectacc_code = "";//選取的detail科目代號
			//String wlx07_m_checkbank_operation_field_sum="";
			//String wlx07_m_checkbank_operation_field="";
			//String wlx07_m_checkbank_field_sum="";
			String ori_field="";
			String wlx07_m_checkbankacc_code="";
			//String wlx07_m_checkbank_field_sum_hsien_id = "";//縣市別用
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
              		   	   //wlx07_m_checkbankacc_code = (String)detail_column.get(j);
            	 		   selectacc_code +="'"+wlx07_m_checkbankacc_code+"'";
            	 		   //95.09.26 add=================================================================================================================
            	 		   ori_field += wlx07_m_checkbankacc_code;
            	 		   //sum(d.ATM_CNT) as ATM_CNT,
            	 		   //if(wlx07_m_checkbank_field_sum_hsien_id.indexOf(" ,sum(d."+wlx07_m_checkbankacc_code+") as "+wlx07_m_checkbankacc_code) == -1){
            	 		   //    wlx07_m_checkbank_field_sum_hsien_id += " ,sum(d."+wlx07_m_checkbankacc_code+") as "+wlx07_m_checkbankacc_code;
            	 		   //}
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
			prop_column_name.load(new FileInputStream(Utility.getProperties("schemaDir")+System.getProperty("file.separator")+"WLX07_M_CHECKBANK_column.TXT"));
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
            select bank_no,bank_name,m_yearmonth,
                   hsien_id,hsien_name,fr001w_output_order,
            	   checkbank_cnt_tot,
            	   checkbank_bal_tot,
            	   checkbank_bal,
            	   checkbank_bal_s,
            	   checkbank_bal_n
            from ( select  bn01.bank_no,bn01.bank_name,
            			   m_year || '/' || decode(substr(to_char(m_month),2),'','0','') || decode(substr(to_char(m_month),1),'1','1',substr(to_char(m_month),1)) as m_yearmonth,
              		 	   m_year,m_month,
              		 	   nvl(cd01.hsien_id,' ') as hsien_id ,
              		 	   nvl(cd01.hsien_name,'OTHER')  as  hsien_name,  
            		 	   cd01.FR001W_output_order as FR001W_output_order, 
            		 	   (CheckBank_Cnt +  CheckBank_Cnt_S + CheckBank_Cnt_N) as CheckBank_cnt_TOT, 
            		 	   Round( (CheckBank_Bal + CheckBank_Bal_S +  CheckBank_Bal_N)/1,0) as CheckBank_Bal_TOT, 
            		 	   CheckBank_Cnt, Round(CheckBank_Bal  /1,0) As CheckBank_Bal,  
            		 	   CheckBank_Cnt_S, Round( CheckBank_Bal_S  /1,0) as CheckBank_Bal_S, 
            		 	   CheckBank_Cnt_N, Round(CheckBank_Bal_N   /1,0) as CheckBank_Bal_N 
              		from  (select * from cd01 where cd01.hsien_id <> 'Y') cd01  
              		left join wlx01 on wlx01.hsien_id=cd01.hsien_id 
              	    left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type='6'  
              	    left join (select * from WLX07_M_CHECKBANK
              			       where (to_char(m_year * 100 + m_month) >= 9501 and to_char(m_year * 100 + m_month) <= 9510)
              			       and bank_no in ('6030016','6230023','6230034')
                               ) WLX07_M_CHECKBANK on  bn01.bank_no = WLX07_M_CHECKBANK.bank_no
            	    where WLX07_M_CHECKBANK.bank_no  <>  ' '  and  WLX07_M_CHECKBANK.CheckBank_Cnt > 0  
                    group by bn01.bank_no,bn01.bank_name,m_year || '/' || decode(substr(to_char(m_month),2),'','0','') || decode(substr(to_char(m_month),1),'1','1',substr(to_char(m_month),1)),
                     		  m_year,m_month,nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,CheckBank_Cnt,
            				  (CheckBank_Cnt +  CheckBank_Cnt_S + CheckBank_Cnt_N) , 
            		 	 	   Round( (CheckBank_Bal + CheckBank_Bal_S +  CheckBank_Bal_N)/1,0), 
            	   		 	   CheckBank_Cnt, Round(CheckBank_Bal  /1,0) ,  
            	   		 	   CheckBank_Cnt_S, Round( CheckBank_Bal_S  /1,0), 
            	   		 	   CheckBank_Cnt_N, Round(CheckBank_Bal_N   /1,0) 				           		  
             	   ) WLX07_M_CHECKBANK
            order by fr001w_output_order,WLX07_M_CHECKBANK.bank_no,WLX07_M_CHECKBANK.m_year,WLX07_M_CHECKBANK.m_month

            //縣市別.按縣市.年月sort(拿掉總計)
            select fr001w_output_order,hsien_id,hsien_name,m_yearmonth,m_year,m_month,
            	   checkbank_cnt_tot,checkbank_bal_tot,checkbank_bal,checkbank_bal_s,checkbank_bal_n 
            from 
            ( select cd01.FR001W_output_order as FR001W_output_order,
              		 nvl(cd01.hsien_id,' ') as hsien_id , 
              		 nvl(cd01.hsien_name,'OTHER')  as  hsien_name,
            		 m_year || '/' || decode(substr(to_char(m_month),2),'','0','') || decode(substr(to_char(m_month),1),'1','1',substr(to_char(m_month),1)) as m_yearmonth,
            		 m_year,m_month,
            		 Sum(CheckBank_Cnt + CheckBank_Cnt_S + CheckBank_Cnt_N) as CheckBank_cnt_TOT,  
            		 Round(Sum(CheckBank_Bal + CheckBank_Bal_S + CheckBank_Bal_N)/1,0) as CheckBank_Bal_TOT,  
            		 Sum(CheckBank_Cnt) as CheckBank_Cnt, Round(Sum(CheckBank_Bal) /1,0) As CheckBank_Bal,  
            		 Sum(CheckBank_Cnt_S) as CheckBank_Cnt_S,Round(Sum(CheckBank_Bal_S) /1,0) as CheckBank_Bal_S,  
            		 Sum(CheckBank_Cnt_N) as CheckBank_Cnt_N,Round(Sum(CheckBank_Bal_N) /1,0) as CheckBank_Bal_N  
  			 from  (select * from cd01 where cd01.hsien_id <> 'Y') cd01  
  			 left join wlx01 on wlx01.hsien_id=cd01.hsien_id  
  			 left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type='6' 
  			 left join (select * from WLX07_M_CHECKBANK
  	     	 			where (to_char(m_year * 100 + m_month) >= 9501 and to_char(m_year * 100 + m_month) <= 9510) 
  	   				   ) WLX07_M_CHECKBANK on  bn01.bank_no = WLX07_M_CHECKBANK.bank_no 
  			 group by  cd01.FR001W_output_order,nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),m_year || '/' || m_month,m_year,m_month 
		    ) WLX07_M_CHECKBANK			 
		    group by FR001W_output_order,hsien_id,hsien_name,m_yearmonth,m_year,m_month,checkbank_cnt_tot,checkbank_bal_tot,checkbank_bal,checkbank_bal_s,checkbank_bal_n 
			order by fr001w_output_order,m_year,m_month
  			*/

            String column = "";//選取欄位
            StringBuffer condition = new StringBuffer() ;//其他條件
            List conditionList =new ArrayList() ;
            //add 營運中/已裁撤============================
			condition.append(" and bn01.bn_type "+(CANCEL_NO.equals("N")?"<>":"=")+" ? and wlx07_m_checkbank.bank_no = bn01.bank_no ");
            conditionList.add("2") ;
			//======================================================
            StringBuffer sqlCmd= new StringBuffer();
			List sqlCmdList = new ArrayList() ;
            StringBuffer sqlCmd_sum= new StringBuffer();//縣市別小計
            List sqlCmd_sumList =new ArrayList() ;
            //String wlx07_m_checkbank_table="";//wlx07_m_checkbank                       
            sqlCmd.append(" select bank_no,bank_name,m_yearmonth,"
                   + " 		  hsien_id,hsien_name,fr001w_output_order,"
            	   + "        checkbank_cnt_tot,checkbank_bal_tot,checkbank_cnt,checkbank_bal,checkbank_cnt_s,checkbank_bal_s,checkbank_cnt_n,checkbank_bal_n"
            	   + " from ( "
            	   + "	select  bn01.bank_no,bn01.bank_name,"
            	   + "			m_year || '/' || decode(substr(to_char(m_month),2),'','0','') || decode(substr(to_char(m_month),1),'1','1',substr(to_char(m_month),1)) as m_yearmonth,"
              	   + "	 		m_year,m_month,"
              	   + "	 		cd01.FR001W_output_order as FR001W_output_order,"
            	   + " 			nvl(cd01.hsien_id,' ') as hsien_id ,"
              	   + "	 	    nvl(cd01.hsien_name,'OTHER')  as  hsien_name," 
            	   + "		    (CheckBank_Cnt +  CheckBank_Cnt_S + CheckBank_Cnt_N) as CheckBank_cnt_TOT," 
            	   + "   	    Round( (CheckBank_Bal + CheckBank_Bal_S +  CheckBank_Bal_N)/1,0) as CheckBank_Bal_TOT," 
            	   + "	 		CheckBank_Cnt, Round(CheckBank_Bal  /1,0) As CheckBank_Bal,"  
            	   + "		    CheckBank_Cnt_S, Round( CheckBank_Bal_S  /1,0) as CheckBank_Bal_S," 
            	   + "		    CheckBank_Cnt_N, Round(CheckBank_Bal_N   /1,0) as CheckBank_Bal_N " 
              	   + "	 from  (select * from ").append(cd01Table).append(" cd01 where cd01.hsien_id <> 'Y') cd01 " 
              	   + "	 left join ").append(wlx01.toString()).append(" wlx01 on wlx01.hsien_id=cd01.hsien_id "
                   + "   left join ").append(bn01.toString()).append(" bn01 on wlx01.bank_no = bn01.bank_no and bn01.bank_type=?"
              	   + "	 left join (select * from WLX07_M_CHECKBANK "
              	   + "  			where (to_char(m_year * 100 + m_month) >= ? and to_char(m_year * 100 + m_month) <= ? )"              	   
				   + "               and  bank_no in ("+selectBank_no+")"	
              	   + "			   )WLX07_M_CHECKBANK on  bn01.bank_no = WLX07_M_CHECKBANK.bank_no "
              	   + "   where WLX07_M_CHECKBANK.bank_no  <>  ' ' and WLX07_M_CHECKBANK.CheckBank_Cnt > 0" 
              	   + "	 group by bn01.bank_no,bn01.bank_name,m_year || '/' || decode(substr(to_char(m_month),2),'','0','') || decode(substr(to_char(m_month),1),'1','1',substr(to_char(m_month),1)),"
         		   + "			  m_year,m_month,nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,CheckBank_Cnt,"
				   + "			  (CheckBank_Cnt +  CheckBank_Cnt_S + CheckBank_Cnt_N) , "
		 	 	   + "			  Round( (CheckBank_Bal + CheckBank_Bal_S +  CheckBank_Bal_N)/1,0), "
	   		 	   + "			  CheckBank_Cnt, Round(CheckBank_Bal  /1,0) , " 
	   		 	   + "			  CheckBank_Cnt_S, Round( CheckBank_Bal_S  /1,0), "
	   		 	   + "			  CheckBank_Cnt_N, Round(CheckBank_Bal_N   /1,0)" 			              	   
              	   + " )WLX07_M_CHECKBANK");   
            	   //+ " order by fr001w_output_order,WLX07_M_CHECKBANK.bank_no,WLX07_M_CHECKBANK.m_year,WLX07_M_CHECKBANK.m_month";            	   
		    sqlCmdList.add(u_year) ;
		    sqlCmdList.add(u_year) ;
		    sqlCmdList.add(bank_type) ;
		    sqlCmdList.add(S_YEAR+S_MONTH) ;
		    sqlCmdList.add(E_YEAR+E_MONTH) ;
			sqlCmd_sum.append(" select fr001w_output_order,hsien_id,hsien_name,m_yearmonth,m_year,m_month,"
					   + "        checkbank_cnt_tot,checkbank_bal_tot,checkbank_cnt,checkbank_bal,checkbank_cnt_s,checkbank_bal_s,checkbank_cnt_n,checkbank_bal_n"
			 		   + " from "
            		   + " ( select cd01.FR001W_output_order as FR001W_output_order," 
            		   + "			nvl(cd01.hsien_id,' ') as hsien_id ," 
              		   + " 			nvl(cd01.hsien_name,'OTHER')  as  hsien_name,"
              		   + "			m_year || '/' || decode(substr(to_char(m_month),2),'','0','') || decode(substr(to_char(m_month),1),'1','1',substr(to_char(m_month),1)) as m_yearmonth,"            		   
		 			   + "			m_year,m_month,"
            		   + "		    Sum(CheckBank_Cnt + CheckBank_Cnt_S + CheckBank_Cnt_N) as CheckBank_cnt_TOT,"  
            		   + "			Round(Sum(CheckBank_Bal + CheckBank_Bal_S + CheckBank_Bal_N)/1,0) as CheckBank_Bal_TOT,"  
            		   + "			Sum(CheckBank_Cnt) as CheckBank_Cnt, Round(Sum(CheckBank_Bal) /1,0) As CheckBank_Bal,"  
            		   + "			Sum(CheckBank_Cnt_S) as CheckBank_Cnt_S,Round(Sum(CheckBank_Bal_S) /1,0) as CheckBank_Bal_S,"  
            		   + "		    Sum(CheckBank_Cnt_N) as CheckBank_Cnt_N,Round(Sum(CheckBank_Bal_N) /1,0) as CheckBank_Bal_N"  
              		   + "  from  (select * from ").append(cd01Table).append(" cd01 where cd01.hsien_id <> 'Y') cd01 " 
                	   + "	left join ").append(wlx01.toString()).append(" wlx01 on wlx01.hsien_id=cd01.hsien_id " 
              		   + "	left join ").append(bn01.toString()).append(" bn01 on wlx01.bank_no=bn01.bank_no and bn01.bank_type in (");
            sqlCmd_sum.append("ALL".equals(bank_type)?"? ,? " : " ? ");
            sqlCmd_sum.append(")"
              		   + "  left join (select * from WLX07_M_CHECKBANK " 
              		   + "  			where (to_char(m_year * 100 + m_month) >= ? and to_char(m_year * 100 + m_month) <= ? )"              	   
              	       + "			  ) WLX07_M_CHECKBANK  on  bn01.bank_no = WLX07_M_CHECKBANK.bank_no" 
              	       + "	group by  cd01.FR001W_output_order,nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),"
              	       + "            m_year || '/' || decode(substr(to_char(m_month),2),'','0','') || decode(substr(to_char(m_month),1),'1','1',substr(to_char(m_month),1)),m_year,m_month "              		   
              	       + " ) WLX07_M_CHECKBANK "
              	       + " where m_year is not null "
					   + " and m_month is not null "
              	       + " group by FR001W_output_order,hsien_id,hsien_name,m_yearmonth,m_year,m_month,"
              	       + "          checkbank_cnt_tot,checkbank_bal_tot,checkbank_cnt,checkbank_bal,checkbank_cnt_s,checkbank_bal_s,checkbank_cnt_n,checkbank_bal_n");
				       //+ " order by fr001w_output_order,m_year,m_month";  					   
		    sqlCmd_sumList.add(u_year) ;
		    sqlCmd_sumList.add(u_year) ;
		    if("ALL".equals(bank_type)) {
		    	sqlCmd_sumList.add("6") ;
		    	sqlCmd_sumList.add("7") ;
		    }else {
		    	sqlCmd_sumList.add(bank_type) ;
		    }
		    sqlCmd_sumList.add(S_YEAR+S_MONTH) ;
		    sqlCmd_sumList.add(E_YEAR+E_MONTH) ;
			if(!order.equals("")){
			    //各別機構
			    sqlCmd.append(" order by "+order );
			    sqlCmd_sum.append(" order by " + order);
  				if( session.getAttribute("SortBy") != null && !((String)session.getAttribute("SortBy")).equals("")){
  		            sqlCmd.append(" " + ((String)session.getAttribute("SortBy")));
  		            sqlCmd_sum.append(" " + ((String)session.getAttribute("SortBy")));
  		        }
  		    }else{
  		        sqlCmd.append(" order by fr001w_output_order,WLX07_M_CHECKBANK.bank_no,WLX07_M_CHECKBANK.m_year,WLX07_M_CHECKBANK.m_month");
  		        sqlCmd_sum.append(" order by fr001w_output_order,m_year,m_month");  		        			   
  		    }			
            //System.out.println("sqlCmd="+sqlCmd);

            //讀取報表欄位名稱===================================================================================
        	Properties p = new Properties();
			p.load(new FileInputStream(Utility.getProperties("schemaDir")+System.getProperty("file.separator")+"WLX07_M_CHECKBANK_column.TXT"));
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

            for(i=2;i<columnLength+4;i++){
              reportUtil.createCell( wb, row, ( short )i, "", noBorderDefaultStyle );
            }
            sheet.addMergedRegion( new Region( ( short )1, ( short )1,
                                               ( short )1,
                                               ( short )(columnLength+3)) );
            //列印年月=======================================================================================            
            row = sheet.createRow( ( short )2 );
            row.setHeight((short) 0x200);
            //95.12.06 add 查詢年月區間
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
                reportUtil.createCell( wb, row, ( short )3, "查詢年月", columnStyle );//95.12.06 add 查詢年月
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
            //row.setHeightInPoints(90);//設定細項表頭高度
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
            wb.setRepeatingRowsAndColumns(0, 1, columnLength+3, 1, 7); //設定表頭 為固定 先設欄的起始再設列的起始

  			System.out.println("DS011W_Excel.sqlCmd="+sqlCmd);
  			System.out.println("ori_field="+ori_field);
  			List dbData = null;
  			if(hasBankListALL.equals("false")){
			  //dbData = DBManager.QueryDB(sqlCmd,"checkbank_cnt_tot,checkbank_bal_tot,checkbank_cnt,checkbank_bal,checkbank_cnt_s,checkbank_bal_s,checkbank_cnt_n,checkbank_bal_n");
			  dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),sqlCmdList,"checkbank_cnt_tot,checkbank_bal_tot,checkbank_cnt,checkbank_bal,checkbank_cnt_s,checkbank_bal_s,checkbank_cnt_n,checkbank_bal_n") ;
			  //System.out.println("DS011W_Excel.sqlCmd="+sqlCmd);
			}else{//95.09.04 add 金融機構代號=ALL全部
			  //dbData = DBManager.QueryDB(sqlCmd_sum,"checkbank_cnt_tot,checkbank_bal_tot,checkbank_cnt,checkbank_bal,checkbank_cnt_s,checkbank_bal_s,checkbank_cnt_n,checkbank_bal_n");
			  dbData = DBManager.QueryDB_SQLParam(sqlCmd_sum.toString(),sqlCmd_sumList,"checkbank_cnt_tot,checkbank_bal_tot,checkbank_cnt,checkbank_bal,checkbank_cnt_s,checkbank_bal_s,checkbank_cnt_n,checkbank_bal_n") ;
			  //System.out.println("DS011W_Excel.sqlCmd_sum="+sqlCmd_sum);
			}

			short rowNo = ( short )8;//資料起始列
			//無資料時,顯示訊息========================================================================
			if(dbData == null || dbData.size() == 0){
			   	row = sheet.createRow( rowNo );
                row.setHeight((short) 0x120);
                reportUtil.createCell( wb, row, ( short )1,"無資料" ,noBorderDefaultStyle );
                sheet.addMergedRegion( new Region( ( short )8, ( short )1,
                                               ( short )8,
                                               ( short )(columnLength+5)) );
			}else{
			//有Data時,將DBData寫入===============================================================================================
            acc_code_row = sheet.getRow(7);
            short lastCellNum = acc_code_row.getLastCellNum();
            System.out.println("lastCellNum="+lastCellNum);
            columnIdx = 1;
            double amt_d = 0.0;
            float amt_f = 0;            
            String amt="";
            String prtbank_code = "";
            DataObject bean = null ;
            for(i=0;i<dbData.size();i++){
            	bean = (DataObject) dbData.get(i) ;
                acc_code = "";                
                row = sheet.createRow( rowNo );
                //System.out.println("rowNo="+rowNo);
                row.setHeight((short) 0x120);
                if(hasBankListALL.equals("false")){
                   //System.out.println("bank_code="+(String) ((DataObject) dbData.get(i)).getValue("bank_no"));                                      
                   reportUtil.createCell( wb, row, ( short )columnIdx, (String) bean.getValue("bank_no"), defaultStyle );//單位代號                   
                   columnIdx++;
                   reportUtil.createCell( wb, row, ( short )columnIdx, (String) bean.getValue("bank_name"), defaultStyle );//機構名稱
                   columnIdx++;
                   prtbank_code = (String) bean.getValue("bank_no");                   
                }else{//縣市別小計
                   //System.out.println("hsien_id="+(String) ((DataObject) dbData.get(i)).getValue("hsien_id"));
                   reportUtil.createCell( wb, row, ( short )columnIdx, (String) bean.getValue("hsien_id"), defaultStyle );//單位代號
                   columnIdx++;
                   reportUtil.createCell( wb, row, ( short )columnIdx, (String) bean.getValue("hsien_name"), defaultStyle );//機構名稱
                   columnIdx++;
                }
                //95.12.06 add 查詢年月
                reportUtil.createCell( wb, row, ( short )columnIdx, bean.getValue("m_yearmonth").toString(), defaultStyle );//查詢年月
                columnIdx++;
                for(int cellIdx =4;cellIdx < (new Short(lastCellNum)).intValue();cellIdx++){
                     amt="";
                     cell = acc_code_row.getCell((short)cellIdx);
                     acc_code = cell.getStringCellValue().toLowerCase();                                          
                     //System.out.println("acc_code="+acc_code);
                     if(bean.getValue(acc_code) != null){
                        amt =bean.getValue(acc_code).toString();
                     }
                     if(acc_code.indexOf("bal") != -1){//餘額類
                        amt = Utility.getRound(amt,Unit);                        
                     }
                     amt = Utility.setCommaFormat(amt);
                     reportUtil.createCell( wb, row, ( short )columnIdx, amt, rightStyle );
                     columnIdx ++;
                }
                columnIdx = 1;
                rowNo++;
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
               	     acc_code = (String)detail_column.get(j);
                     //System.out.println((String)detail_column.get(j)+"="+Utility.ISOtoBig5((String)prop_column_name.get((String)detail_column.get(j))));                     
               		 row = sheet.getRow(6);
                  	 reportUtil.createCell( wb, row, ( short )columnIdx,Utility.ISOtoUTF8((String)prop_column_name.get(acc_code)), columnStyle );
                  	    sheet.addMergedRegion( new Region( ( short )6, ( short )columnIdx,
                                                           ( short )7,
                                               			   ( short )columnIdx) );
                  		columnIdx ++;               		
                }
            }            
            //設定寬度============================================================
            for ( i = 1; i <= columnLength+3; i++ ) {
                if(i==2){
                  sheet.setColumnWidth( ( short )i,
                                        ( short ) ( 256 * ( 25 + 4 ) ) );//機構名稱
                }else{
                  sheet.setColumnWidth( ( short )i,
                                        ( short ) ( 256 * ( 21 + 4 ) ) );
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