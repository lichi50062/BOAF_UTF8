<%
//94.02.16 add 申報資料批次下載 by 2295
//94.06.21 fix 無資料時顯示無資料可提供下載 by 2295
//95.11.20 fix 取不到月份的問題 by 2295
//         add A06.A99批次下載 by 2295
//96.04.16 add 中央存保格式A02.A01.A99檔案下載 by 2295
//96.04.24 fix A99_extra中央存保格式A02.A01.A99檔案下載改用回傳List by 2295
//96.07.11 add A08批次下載 by 2295
//97.01.03 add A09批次下載 by 2295
//97.06.13 add A10批次下載 by 2295
//97.06.19 fix A10_extra中央存保格式A10檔案下載 by 2295
//99.10.07 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
//103.01.13 add A11_extra中央存保用檔案下載 by 2295
//104.01.21 add A12 by 2295
//104.03.14 add A10_extra中文存保用檔案下載增加欄位 by 2968
//104.04.27 add A12_extra中文存保用檔案下載 by 2295
//104.10.13 add A13 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DownLoad" %>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.report.*" %>
<%@ page import="java.util.*" %>
<%@include file="./include/Header.include" %>
<%
	
	String S_YEAR = Utility.getTrimString(dataMap.get("S_YEAR"));
	String S_MONTH = Utility.getTrimString(dataMap.get("S_MONTH"));
	String M_MONTH = Utility.getTrimString(dataMap.get("M_MONTH"));
	String Report_no = Utility.getTrimString(dataMap.get("Report_no"));
		
	//fix 93.12.18 若有已點選的tbank_no,則以已點選的tbank_no為主============================================================
	String bank_code = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");				
	String nowtbank_no = Utility.getTrimString(dataMap.get("tbank_no"));
	if(!nowtbank_no.equals("")){
	   session.setAttribute("nowtbank_no",nowtbank_no);//將已點選的tbank_no寫入session	   
	}   
	bank_code = ( session.getAttribute("nowtbank_no")==null ) ? bank_code : (String)session.getAttribute("nowtbank_no");			
	//=======================================================================================================================
    //fix 93.12.20 若有已點選的bank_type,則以已點選的bank_type為主============================================================	
	String bank_type = Utility.getTrimString(dataMap.get("bank_type"));
	bank_type = ( session.getAttribute("nowbank_type")==null ) ? bank_type : (String)session.getAttribute("nowbank_type");			
	//=======================================================================================================================	
	
	
	StringBuffer sqlCmd = new StringBuffer();
	List paramList = new ArrayList();
	String BankList = "";    
    List BankList_data = null;
    String selectBank_no = "";//選取的金融機構代號
    
	//S_MONTH = String.valueOf(Integer.parseInt(S_MONTH));//95.11.20 fix by 2295	
    if(!Utility.CheckPermission(request,report_no)){//無權限時,導向到LoginError.jsp
        rd = application.getRequestDispatcher( LoginErrorPgName );        
    }else{            
    	//set next jsp 	
    	//將選取的金融機構代碼.縣市別寫到session=======================================    
    	if(request.getParameter("BankList")	!= null && !((String)request.getParameter("BankList")).equals("")){
    	   session.setAttribute("BankList",(String)request.getParameter("BankList"));    	
    	}  
    	if(request.getParameter("HSIEN_ID")	!= null && !((String)request.getParameter("HSIEN_ID")).equals("")){
           session.setAttribute("HSIEN_ID",(String)request.getParameter("HSIEN_ID"));   
        } 
    	if(act.equals("new")){    	    
        	rd = application.getRequestDispatcher( QryPgName+"?bank_type="+bank_type );            	
    	}else if(act.equals("Download")){   
    	    //金融機構
			if(session.getAttribute("BankList") != null && !((String)session.getAttribute("BankList")).equals("")){
		   		BankList = (String)session.getAttribute("BankList");
		   		BankList_data = Utility.getReportData(BankList);
		   		System.out.println("BankList_data.size()="+BankList_data.size());		   
			}
    		List data=null; 	
    		//============modify by 2354 12.22 begin ---
    		if(Report_no.equals("A01") || Report_no.equals("A02") || Report_no.equals("A03") || Report_no.equals("A04") || Report_no.equals("A05") || Report_no.equals("A06") || Report_no.equals("A99") || Report_no.equals("A99_extra") || Report_no.equals("A08") || Report_no.equals("A09") || Report_no.equals("A10") || Report_no.equals("A10_extra") || Report_no.equals("A11_extra") || Report_no.equals("A12") || Report_no.equals("A12_extra") || Report_no.equals("A13")){ //95.11.20 add A06.A99 批次下載 by 2295 //96.07.11 add A08 //97.01.03 add A09 //97.06.13 add A10 //103.01.13 add A11_extra //104.01.21 add A12				
    			//金融機構代號=============================================================
                if(BankList_data != null && BankList_data.size() != 0){
                   if(Report_no.equals("A11_extra")){
                   	  selectBank_no += " bank_no IN (";
                   }else{	
                      selectBank_no += " bank_code IN (";
                   }
                   for(int i=0;i<BankList_data.size();i++){
                	 selectBank_no +="'"+(String)((List)BankList_data.get(i)).get(0)+"'";            	
                	 if(i < BankList_data.size()-1) selectBank_no +=",";
                   }
                   selectBank_no += ")";
                }                   
                //==============================================================================
    			if(!Report_no.equals("A11_extra")){
    			   sqlCmd.append("select * from "+Report_no+" where m_year=? and m_month=?");
    			   paramList.add(S_YEAR);
    			   paramList.add(S_MONTH);    			   
                   sqlCmd.append(" and "+selectBank_no);
            	}
                if(!Report_no.equals("A08") && !Report_no.equals("A09") && !Report_no.equals("A10") && !Report_no.equals("A11_extra") && !Report_no.equals("A12")) sqlCmd.append(" order by acc_code ");
               
                if(Report_no.equals("A10_extra")){//97.06.19 add 中央存保格式A10檔案下載	    	       
 	               data = RptFR047W.printData(S_YEAR, S_MONTH,selectBank_no,"","1");
 	            }else if(Report_no.equals("A99_extra")){//96.04.16 add 中央存保格式A02.A01.A99檔案下載	
 	               data = RptFR0066WB.printData(S_YEAR, S_MONTH,selectBank_no,"","1");//96.04.24 fix 改用回傳List
 	            }else if(Report_no.equals("A12_extra")){//104.04.27 add 中央存保格式A12檔案下載	
 	               data = FR068W_Excel.printData(S_YEAR, S_MONTH,selectBank_no,"","1"); 	                                  
 	            }else if(Report_no.equals("A08")){ //96.07.11 add A08 
    		       data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,warnaccount_cnt,limitaccount_cnt,erroraccount_cnt,otheraccount_cnt,depositaccount_tcnt");
    		    }else if(Report_no.equals("A09")){ //97.01.03 add A09 
    		       data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,over_cnt,over_amt,push_over_amt,totalamt,push_totalamt,over_total_rate");    		      
    		    }else if(Report_no.equals("A10")){ //97.06.13 add A10 
    		       data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,loan1_amt,loan2_amt,loan3_amt,loan4_amt,invest1_amt,invest2_amt,invest3_amt,invest4_amt,other1_amt,other2_amt,other3_amt,other4_amt,"
    		    		   +"loan1_baddebt,loan2_baddebt,loan3_baddebt,loan4_baddebt,build1_baddebt,build2_baddebt,build3_baddebt,build4_baddebt,baddebt_noenough,baddebt_104,baddebt_105,baddebt_106,baddebt_107,baddebt_108");    		      
    		    }else if(Report_no.equals("A11_extra")){ //103.01.13 add A11 
    		       sqlCmd.append(" select m_year,m_month,bank_no as bank_code,");
                   sqlCmd.append(" case_no,");//--申報編號
                   sqlCmd.append(" loan_amt_sum,");//--授信案總金額
                   sqlCmd.append(" bank_no_max,");//--主辦行 1全國農業金庫2其他金融機構
                   sqlCmd.append(" loan_kind,");//--參貸型式1主辦行2參貸行
                   sqlCmd.append(" loan_amt,");//--參貸額度
                   sqlCmd.append(" loan_bal_amt,");//--實際授信餘額
                   sqlCmd.append(" loan_type,");//--信用部參貸部分之授信用途 1購地2建築3其他
                   sqlCmd.append(" pay_state,");//--目前放款繳息情形1正常2逾放3其他有欠正常放款4申報基準日無放款
                   sqlCmd.append(" new_case,");//--是否本月新增案件
                   sqlCmd.append(" manabank_name");//--管理行
                   sqlCmd.append(" from wlx10_m_loan");
                   sqlCmd.append(" where m_year=? and m_month=?");
                   paramList.add(S_YEAR);
    			   paramList.add(S_MONTH);    			   
                   sqlCmd.append(" and "+selectBank_no);                   
                   sqlCmd.append(" union"); 
                   sqlCmd.append(" select m_year,m_month,bank_no as bank_code,'NODATA',0,'','',0,0,'','','',''");
                   sqlCmd.append(" from wlx10_m_loan_apply");
                   sqlCmd.append(" where m_year=? and m_month=?");
                   paramList.add(S_YEAR);
    			   paramList.add(S_MONTH);    			   
                   sqlCmd.append(" and "+selectBank_no); 
                   sqlCmd.append(" order by m_year,m_month,bank_code,case_no");
    		       data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,loan_amt_sum,loan_amt,loan_bal_amt");    		      
    		    }else{    		      
    		       data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt"); 	 
    		    }   
	    	}else if(Report_no.equals("M01") || Report_no.equals("M02") || Report_no.equals("M03") || Report_no.equals("M04") || Report_no.equals("M05") || Report_no.equals("M06") || Report_no.equals("M07")){
	    	    sqlCmd.append("select * from "+Report_no+" where m_year=? and m_month=?");
	    	    paramList.add(S_YEAR);
	    	    paramList.add(S_MONTH);
	    		data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt"); 	    
	    	}
	    	//============modify by 2354 12.22 end ---
    		if(data == null || data.size() == 0){
    		   actMsg = actMsg + "無資料可提供下載";    		   
    		   request.setAttribute("actMsg",actMsg);
    		   rd = application.getRequestDispatcher( nextPgName );
    		}else{    		   
    		   rd = application.getRequestDispatcher( PrintDataPgName +"?Report_no="+Report_no+"&M_YEAR="+S_YEAR+"&M_MONTH="+S_MONTH+"&bank_code="+bank_code+"&test=nothing");            		   
    	    }
    	}
    }        
     
%>

<%@include file="./include/Tail.include" %>

<%!
    private final static String report_no = "WMFileDownloadBatch";
    private final static String nextPgName = "/pages/ActMsg.jsp";
    private final static String QryPgName = "/pages/"+report_no+"_Qry.jsp";
    private final static String PrintDataPgName = "/pages/"+report_no+"_PrintData.jsp";    
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
%>    