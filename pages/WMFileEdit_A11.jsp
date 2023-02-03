<%
//102.1.8 created by 2968
//102.12.26 add 本月無申報資料時,也要寫入申報記錄檔WML01 by 2295
//104.05.12 add 增加A111111111可調整案件編號及分項項數 by 2295
//104.05.12 fix 補舊案件資料時,新增下一筆參貸項目時,該案件編號會是空值 by 2295
%>

<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.ArrayList" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Date" %>
<%
	System.out.println("WMFileEdit_A11 start ..........");
	
	RequestDispatcher rd = null;
	String actMsg = "";	
	String alertMsg = "";	
	String webURL = "";	
	boolean doProcess = false;	
	String userid = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");
	String username = ( session.getAttribute("muser_name")==null ) ? "" : (String)session.getAttribute("muser_name");
	//取得session資料,取得成功時,才繼續往下執行===================================================
	if(userid == null){//session timeout	
      System.out.println("WMFileEdit_A11 login timeout");   
	   rd = application.getRequestDispatcher( "/pages/reLogin.jsp?url=LoginError.jsp?timeout=true" );         	   
	   try{
          rd.forward(request,response);
       }catch(Exception e){
          System.out.println("forward Error:"+e+e.getMessage());
       }
    }else{
      doProcess = true;
    }    
    
	if(doProcess){//若userid有資料時,表示登入成功====================================================================	
	
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");			
	String bank_type = ( request.getParameter("bank_type")==null ) ? "" : (String)request.getParameter("bank_type");							
	String list_type = ( request.getParameter("list_type")==null ) ? "" : (String)request.getParameter("list_type");				
	System.out.println("act="+act);		
	System.out.println("list_type="+list_type);	
	
	//登入者資訊
	String lguser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		
	String lguser_name = ( session.getAttribute("muser_name")==null ) ? "" : (String)session.getAttribute("muser_name");		
	String lguser_type = ( session.getAttribute("muser_type")==null ) ? "" : (String)session.getAttribute("muser_type");							
	String lguser_bank_type = ( session.getAttribute("bank_type")==null ) ? "" : (String)session.getAttribute("bank_type");
	
	//======================================================================================================================											
	//fix 94.01.07 若有已點選的tbank_no,則以已點選的tbank_no為主
	String bank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");				
	String nowtbank_no =  ( request.getParameter("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");			
	if(!nowtbank_no.equals("")){
	   session.setAttribute("nowtbank_no",nowtbank_no);//將已點選的tbank_no寫入session	   
	}   
	bank_no = ( session.getAttribute("nowtbank_no")==null ) ? bank_no : (String)session.getAttribute("nowtbank_no");		
	//=======================================================================================================================
	if(!Utility.CheckPermission(request,report_no)){//無權限時,導向到LoginError.jsp
        rd = application.getRequestDispatcher( LoginErrorPgName );
    }else{
	//set next jsp 	
    if(act.equals("New")){
        String m_year  =  ( request.getParameter("s_year")==null ) ? "" : (String)request.getParameter("s_year");
        String m_month =  ( request.getParameter("s_month")==null ) ? "" : (String)request.getParameter("s_month");
    	request.setAttribute("act",act);
    	request.setAttribute("A11_S_Edit",getLoan_Data("","","",""));
    	request.setAttribute("A11_Lock",isLock(m_year,m_month,bank_no));
    	request.setAttribute("isLastMonthData",isLastMonthData(m_year,m_month,bank_no));
    	request.setAttribute("isHaveApplyData",isHaveApplyData(m_year,m_month,bank_no));
    	request.setAttribute("isHaveNoApplyData",isHaveNoApplyData(m_year,m_month,bank_no));
    	request.setAttribute("isHaveNoApplyLoanData",isHaveNoApplyLoanData(m_year,m_month,bank_no));
    	//編輯頁下拉選單
    	List bankNoMaxList = getBankNoMaxList("041");
    	request.setAttribute("bankNoMaxList",bankNoMaxList);
    	List loanKindList = getLoanKindList("042");
    	request.setAttribute("loanKindList",loanKindList);
    	List loanTypeList = getLoanTypeList("043");
    	request.setAttribute("loanTypeList",loanTypeList);
    	List payStateList = getPayStateList("044");
    	request.setAttribute("payStateList",payStateList);
    	List violateTypeList = getViolateTypeList("045");
    	request.setAttribute("violateTypeList",violateTypeList);
    	if(m_month.length()==1){
    	        m_month = "0"+m_month;
    	} 
    	if(m_year.length()==1){
    	        m_year = "00"+m_year;
    	}else if(m_year.length()==2){
    	        m_year = "0"+m_year;
    	} 
    	String NextCase_no = m_year+m_month+"0101";
    	rd = application.getRequestDispatcher( getCaseCntPgName +"?bank_no="+bank_no+"&case_no="+NextCase_no+"&s_year="+m_year+"&s_month="+m_month);
    }else if(act.equals("continueEditCnt")){
        String m_year		=((String)request.getParameter("m_year")==null)?"":(String)request.getParameter("m_year");
  		String m_month 		=((String)request.getParameter("m_month")==null)?"":(String)request.getParameter("m_month");
  		bank_no 		=((String)request.getParameter("bank_no")==null)?"":(String)request.getParameter("bank_no");
  		String case_no 		=((String)request.getParameter("case_no")==null)?"":(String)request.getParameter("case_no");
    	request.setAttribute("act",act);
    	request.setAttribute("A11_S_Edit",getLoan_Data("","","",""));
    	request.setAttribute("A11_Lock",isLock(m_year,m_month,bank_no));
    	request.setAttribute("isLastMonthData",isLastMonthData(m_year,m_month,bank_no));
    	request.setAttribute("isHaveApplyData",isHaveApplyData(m_year,m_month,bank_no));
    	request.setAttribute("isHaveNoApplyData",isHaveNoApplyData(m_year,m_month,bank_no));
    	request.setAttribute("isHaveNoApplyLoanData",isHaveNoApplyLoanData(m_year,m_month,bank_no));
    	//編輯頁下拉選單
    	List bankNoMaxList = getBankNoMaxList("041");
    	request.setAttribute("bankNoMaxList",bankNoMaxList);
    	List loanKindList = getLoanKindList("042");
    	request.setAttribute("loanKindList",loanKindList);
    	List loanTypeList = getLoanTypeList("043");
    	request.setAttribute("loanTypeList",loanTypeList);
    	List payStateList = getPayStateList("044");
    	request.setAttribute("payStateList",payStateList);
    	List violateTypeList = getViolateTypeList("045");
    	request.setAttribute("violateTypeList",violateTypeList);
    	String NextCase_no = "";
    	String Mcase_no = getMaxCase_No(m_year,m_month,bank_no,case_no);
    	String Mcase_cnt = "";
    	String firstLoan_amt_sum = "";
    	if(!"".equals(Mcase_no)){
    	    Mcase_cnt = getCase_Cnt(m_year,m_month,bank_no,Mcase_no);
    	    firstLoan_amt_sum = getFirstLoan_amt_sum(m_year,m_month,bank_no,Mcase_no.substring(0, 7)+"01");
    	    if((String.valueOf(Integer.parseInt(Mcase_no.substring(7, 9)))).equals(Mcase_cnt)){
    	        NextCase_no = String.valueOf(Integer.parseInt(Mcase_no.substring(0, 7))+1)+"01";
    	    }else{
    	        NextCase_no = String.valueOf((Integer.parseInt(Mcase_no)+1));
    	    }
    	}
    	request.setAttribute("firstLoan_amt_sum",firstLoan_amt_sum);
    	rd = application.getRequestDispatcher( EditPgName +"?bank_no="+bank_no+"&case_no="+NextCase_no+"&case_cnt="+Mcase_cnt+"&s_year="+m_year+"&s_month="+m_month);
    }else if(act.equals("nextPg")){
        System.out.println("act=nextPg start..........");
        String m_year  =  ( request.getParameter("s_year")==null ) ? "" : (String)request.getParameter("s_year");
        String m_month =  ( request.getParameter("s_month")==null ) ? "" : (String)request.getParameter("s_month");
        String case_cnt =  ( request.getParameter("case_cnt")==null ) ? "" : (String)request.getParameter("case_cnt");
        List dbData = getLoan_Data("","","","");
        request.setAttribute("act",act);
        request.setAttribute("A11_S_Edit",dbData);
        request.setAttribute("A11_Lock",isLock(m_year,m_month,bank_no));
        request.setAttribute("isLastMonthData",isLastMonthData(m_year,m_month,bank_no));
        request.setAttribute("isHaveApplyData",isHaveApplyData(m_year,m_month,bank_no));
        request.setAttribute("isHaveNoApplyData",isHaveNoApplyData(m_year,m_month,bank_no));
        request.setAttribute("isHaveNoApplyLoanData",isHaveNoApplyLoanData(m_year,m_month,bank_no));
        //編輯頁下拉選單
        List bankNoMaxList = getBankNoMaxList("041");
        request.setAttribute("bankNoMaxList",bankNoMaxList);
        List loanKindList = getLoanKindList("042");
        request.setAttribute("loanKindList",loanKindList);
        List loanTypeList = getLoanTypeList("043");
        request.setAttribute("loanTypeList",loanTypeList);
        List payStateList = getPayStateList("044");
        request.setAttribute("payStateList",payStateList);
        List violateTypeList = getViolateTypeList("045");
        request.setAttribute("violateTypeList",violateTypeList);
        String NextCase_no = "";
        String Mcase_no = getMaxCase_No(m_year,m_month,bank_no,"");
        String Mcase_cnt = "";
        String firstLoan_amt_sum = "";
        if(!"".equals(Mcase_no)){
        	Mcase_cnt = getCase_Cnt(m_year,m_month,bank_no,Mcase_no);
        	if(String.valueOf(Integer.parseInt(Mcase_no.substring(7, 9))).equals(Mcase_cnt)){
        	    NextCase_no = String.valueOf(Integer.parseInt(Mcase_no.substring(0, 7))+1)+"01";
        	}else{
        	    NextCase_no = String.valueOf((Integer.parseInt(Mcase_no)+1));
        	}
        }else{
        	if(m_month.length()==1){
        	    m_month = "0"+m_month;
        	} 
        	if(m_year.length()==1){
        	    m_year = "00"+m_year;
        	}else if(m_year.length()==2){
        	    m_year = "0"+m_year;
        	}
        	NextCase_no = m_year+m_month+"0101";
        }
        request.setAttribute("firstLoan_amt_sum",firstLoan_amt_sum);
        rd = application.getRequestDispatcher( EditPgName +"?bank_no="+bank_no+"&case_no="+NextCase_no+"&case_cnt="+case_cnt);
        System.out.println("act=nextPg end..........");
    //Edit page==========================================================
    }else if(act.equals("Edit")){
    	//取得sequnce number
    	String seq_no =  ( request.getParameter("seq_no")==null ) ? "" : (String)request.getParameter("seq_no");
    	String m_year    =  ( request.getParameter("m_year")==null ) ? "" : (String)request.getParameter("m_year");
    	String m_month =  ( request.getParameter("m_month")==null ) ? "" : (String)request.getParameter("m_month");
    	String case_no =  ( request.getParameter("case_no")==null ) ? "" : (String)request.getParameter("case_no");
    	request.setAttribute("act",act);		
    	List dbData = getLoan_Data(m_year,m_month,bank_no,seq_no);
    	request.setAttribute("A11_S_Edit",dbData);
    	List qNoApplyData = getNoApply_Data(m_year,m_month,bank_no);
    	request.setAttribute("A11_NoApply_Edit",qNoApplyData);
    	request.setAttribute("A11_Lock",isLock(m_year,m_month,bank_no));
    	request.setAttribute("isLastMonthData",isLastMonthData(m_year,m_month,bank_no));
    	request.setAttribute("isHaveApplyData",isHaveApplyData(m_year,m_month,bank_no));
    	request.setAttribute("isHaveNoApplyData",isHaveNoApplyData(m_year,m_month,bank_no));
    	request.setAttribute("isHaveNoApplyLoanData",isHaveNoApplyLoanData(m_year,m_month,bank_no));
    	request.setAttribute("maxCase_No",getMaxCase_No(m_year,m_month,bank_no,case_no));
    	//編輯頁下拉選單
    	List bankNoMaxList = getBankNoMaxList("041");
    	request.setAttribute("bankNoMaxList",bankNoMaxList);
    	List loanKindList = getLoanKindList("042");
    	request.setAttribute("loanKindList",loanKindList);
    	List loanTypeList = getLoanTypeList("043");
    	request.setAttribute("loanTypeList",loanTypeList);
    	List payStateList = getPayStateList("044");
    	request.setAttribute("payStateList",payStateList);
    	List violateTypeList = getViolateTypeList("045");
    	request.setAttribute("violateTypeList",violateTypeList);
    	String firstLoan_amt_sum = "";
    	if(!"".equals(case_no)){
    		firstLoan_amt_sum = getFirstLoan_amt_sum(m_year,m_month,bank_no,case_no.substring(0, 7)+"01");
    	}
    	request.setAttribute("firstLoan_amt_sum",firstLoan_amt_sum);
   	    rd = application.getRequestDispatcher( EditPgName +"?bank_no="+bank_no);
   	//List all the bank============================================
   	}else if(act.equals("List")){    	    	    
   	    //所有資料 
   	    List dbData1 = getBank_Data(bank_no);
   	    request.setAttribute("A11_S",dbData1);
   	    //所有資料之筆數
   	    String m_year    =  ( request.getParameter("m_year")==null ) ? "" : (String)request.getParameter("m_year");
   	    String m_month    =  ( request.getParameter("m_month")==null ) ? "" : (String)request.getParameter("m_month");
   	    List dbData2 = getBank_Sum(bank_no,m_year,m_month);
   	    request.setAttribute("A11_S_Sum",dbData2);
       	rd = application.getRequestDispatcher(ListPgName +"?bank_no="+bank_no+"&bank_type="+bank_type);
   //Insert  data ==================================================
   }else if(act.equals("Insert")){                            	        
        actMsg = InsertDB(request,lguser_id,lguser_name); 
        String case_no      =((String)request.getParameter("case_no")==null)?"":(String)request.getParameter("case_no");
        System.out.println("Insert.actMsg="+actMsg);
        if("EditNextCnt".equals(actMsg)){
            String m_year		=((String)request.getParameter("hyear")==null)?"":(String)request.getParameter("hyear");
      		String m_month 		=((String)request.getParameter("hmonth")==null)?"":(String)request.getParameter("hmonth");
 	    	rd = application.getRequestDispatcher( "/pages/WMFileEdit_A11_EditNextCnt.jsp?act=New&s_year="+m_year+"&s_month="+m_month+"&bank_no="+bank_no+"&case_no="+case_no ); 
        }else{
            rd = application.getRequestDispatcher( nextPgName+"?FX=WMFileEdit_A11" ); 
        }
   //No Apply  data ==================================================
   }else if(act.equals("No_Apply")){                            	        
        actMsg = No_ApplyDB(request,lguser_id,lguser_name);      	
    	rd = application.getRequestDispatcher( nextPgName+"?FX=WMFileEdit_A11" ); 
   //Update  data ==================================================
   }else if(act.equals("Update")){
   	    actMsg = UpdateDB(request,lguser_id,lguser_name);
        rd = application.getRequestDispatcher( nextPgName+"?FX=WMFileEdit_A11" ); 
    	
   //Delete  data ==================================================
   }else if(act.equals("Delete")){
   	    actMsg = DeleteDB(request,lguser_id,lguser_name);
        rd = application.getRequestDispatcher( nextPgName+"?FX=WMFileEdit_A11"); 
   }else if(act.equals("updateCase_NO")){//104.05.12 add
   	    actMsg = updCase_No(request,lguser_id,lguser_name);
        rd = application.getRequestDispatcher( nextPgName+"?FX=WMFileEdit_A11"); 
   }	   		
   request.setAttribute("actMsg",actMsg);    
    } 
   System.out.println("WMFileEdit_A11 End ..........");
   	// ------導向---------------- 
   try{
      //forward to next present jsp
      rd.forward(request, response);
   }catch(NullPointerException npe){
      System.out.println(npe);		
   }
  
 }//end of doProcess
//=================================================================================================
%>

<%!
//=================================================================================================
    private final static String report_no = "WMFileEdit";
	private final static String nextPgName = "/pages/ActMsg.jsp";    
    private final static String EditPgName = "/pages/WMFileEdit_A11_Edit.jsp";    
    private final static String ListPgName = "/pages/WMFileEdit_A11_List.jsp"; 
    private final static String LoadhistoryPgName = "/pages/WMFileEdit_A11_Load.jsp";        
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    private final static String getCaseCntPgName = "/pages/WMFileEdit_A11_CaseCnt.jsp";
  	//獲得點選bank資料==================================================
    private List getBank_Data(String bank_no){
    		List paramList =new ArrayList() ;
    		String sqlCmd = "select m_year || '/' || decode(length(m_month),'1','0','') || m_month as m_yearmonth,"
    		        +"m_year,m_month,sum(apply_cnt) apply_cnt,"
    		        +"sum(loan_amt_sum) loan_amt_sum,"
    		        +"sum(loan_amt) loan_amt,"
    		        +"sum(loan_bal_amt) loan_bal_amt "
		    		+"from "
		    		+"("
		    		+"select m_year,m_month,count(*) as apply_cnt,"
		    		+"sum(LOAN_AMT_SUM) as loan_amt_sum, sum(LOAN_AMT) as loan_amt,"
		    		+"sum(LOAN_BAL_AMT) as loan_bal_amt "
		    		+"from wlx10_m_loan where bank_no=? "
		    		+"group by m_year,m_month "
		    		+"union "
		    		+"select m_year,m_month,1,0,0,0 "
		    		+"from  wlx10_m_loan_apply where bank_no=? "
		    		+"group by m_year,m_month "
		    		+")"
		    		+"group by m_year,m_month "
		    		+"order by m_year desc,m_month desc ";
    	    paramList.add(bank_no) ;
    	    paramList.add(bank_no) ;
        	List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"M_YEARMONTH,M_YEAR,M_MONTH,APPLY_CNT,LOAN_AMT_SUM,LOAN_AMT,LOAN_BAL_AMT");
        	return dbData;  
    }
    
  //點選申報年月後,展開的明細=============
    private List getBank_Sum(String bank_no,String m_year,String m_month){
		List paramList =new ArrayList() ;
    	String sqlCmd = "select m_year || '/' || decode(length(m_month),'1','0','') || m_month as m_yearmonth," //申報年月
    		            +"m_year,m_month,seq_no,case_cnt," //明細.序號.總項數
    		            +"case_no," //--申報編號
    		            +"loan_amt_sum," //授信案總金額
    		            +"loan_amt," //參貸額度
    		            +"loan_bal_amt," //實際授信餘額
    		            +"user_id," //異動者帳號
    		            +"user_name," //異動者姓名
    		            +"to_char(update_date,'yyyy/mm/dd hh:mi') as update_date " //異動日期
    		            +"from wlx10_m_loan "
    		            +"where bank_no = ?  "
				    	//+"and m_year=? "
				    	//+"and m_month=? "
		    	        +"union  "
		    	        +"select m_year || '/' || decode(length(m_month),'1','0','') || m_month as m_yearmonth, "//--申報年月,
		    	        +"m_year,m_month,999,0,'',0,0,0,user_id,user_name, "
		    	        +"to_char(update_date,'yyyy/mm/dd hh:mi') as update_date " //--異動日期
		    	        +"from  wlx10_m_loan_apply where bank_no= ? ";
		    	       // +"and m_year=? "
		    	       // +"and m_month=? "
    	paramList.add(bank_no) ;
    	//paramList.add(m_year) ;
    	//paramList.add(m_month) ;
    	paramList.add(bank_no) ;
    	//paramList.add(m_year) ;
    	//paramList.add(m_month) ;
    	sqlCmd +="order by m_year desc,m_month desc,seq_no ";
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"M_YEARMONTH,M_YEAR,M_MONTH,SEQ_NO,CASE_CNT,CASE_NO,LOAN_AMT_SUM,LOAN_AMT,LOAN_BAL_AMT,USER_ID,USER_NAME,UPDATE_DATE");
        return dbData;
    }
    
   //編輯頁資料
    private List getLoan_Data(String m_year,String m_month,String bank_no,String seq_no){
		List paramList =new ArrayList() ;
		String sqlCmd = "select M_YEAR,M_MONTH,BANK_NO,SEQ_NO,LOAN_IDN,LOAN_NAME,LOAN_AMT_SUM," 
			        + "CASE_BEGIN_YEAR,CASE_BEGIN_MONTH,CASE_END_YEAR,CASE_END_MONTH," 
			        + "BANK_NO_MAX,MANABANK_NAME,LOAN_KIND,LOAN_AMT,LOAN_BAL_AMT,LOAN_TYPE," 
			        + "PAY_STATE,VIOLATE_TYPE,LOAN_RATE,NEW_CASE,USER_ID,USER_NAME,UPDATE_DATE,CASE_NO,CASE_CNT " 
			        + "from wlx10_m_loan "
			        + "where bank_no = ? ";
		paramList.add(bank_no) ;
		if(!"".equals(m_year)&&!"".equals(m_month)){
			sqlCmd += "and m_year = ? and m_month = ?  ";
			paramList.add(m_year) ;
			paramList.add(m_month) ;
		}
		if(!"".equals(seq_no)){
			sqlCmd += "and seq_no = ? ";
			paramList.add(seq_no) ;
		}
    	List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"M_YEAR,M_MONTH,BANK_NO,SEQ_NO,LOAN_IDN,LOAN_NAME,LOAN_AMT_SUM,"
													    	        +"CASE_BEGIN_YEAR,CASE_BEGIN_MONTH,CASE_END_YEAR,CASE_END_MONTH," 
													    	        +"BANK_NO_MAX,MANABANK_NAME,LOAN_KIND,LOAN_AMT,LOAN_BAL_AMT,LOAN_TYPE,"
													    	        +"PAY_STATE,VIOLATE_TYPE,LOAN_RATE,NEW_CASE,USER_ID,USER_NAME,UPDATE_DATE,CASE_NO,CASE_CNT ");
    	return dbData;  
	}//end of getLoan_Data 
	
	
	private List getNoApply_Data(String m_year,String m_month,String bank_no){
		List paramList =new ArrayList() ;
		String sqlCmd = "select M_YEAR,M_MONTH,BANK_NO,APPLY_CNT " 
			        + "from wlx10_m_loan_apply "
			        + "where bank_no = ? and m_year = ? and m_month = ? ";
		paramList.add(bank_no) ;
		paramList.add(m_year) ;
		paramList.add(m_month) ;
    	List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"M_YEAR,M_MONTH,BANK_NO,APPLY_CNT");
    	return dbData;  
	}
	
    //主辦行下拉選項 ================================================
    private List getBankNoMaxList(String cmuse_div){
    		List paramList =new ArrayList() ;
    		String sqlCmd 	= "select cmuse_id,cmuse_name from cdshareno where cmuse_div=? order by input_order";
    		paramList.add(cmuse_div) ;
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"CMUSE_ID,CMUSE_NAME");
        return dbData;
    }
    
  	//參貸型式下拉選項 ================================================
    private List getLoanKindList(String cmuse_div){
        List paramList =new ArrayList() ;
		String sqlCmd 	= "select cmuse_id,cmuse_name from cdshareno where cmuse_div=? order by input_order";
		paramList.add(cmuse_div) ;
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"CMUSE_ID,CMUSE_NAME");
    	return dbData;
	}
    
  	//信用部參貸部分之授信用途下拉選項 ================================================
    private List getLoanTypeList(String cmuse_div){
        List paramList =new ArrayList() ;
		String sqlCmd 	= "select cmuse_id,cmuse_name from cdshareno where cmuse_div=? order by input_order";
		paramList.add(cmuse_div) ;
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"CMUSE_ID,CMUSE_NAME");
    	return dbData;
	}
  	
  	//目前放款繳息情形下拉選項 ================================================
    private List getPayStateList(String cmuse_div){
        List paramList =new ArrayList() ;
		String sqlCmd 	= "select cmuse_id,cmuse_name from cdshareno where cmuse_div=? order by input_order";
		paramList.add(cmuse_div) ;
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"CMUSE_ID,CMUSE_NAME");
    	return dbData;
	}
  	
  	//有無違反契約承諾條款下拉選項 ================================================
    private List getViolateTypeList(String cmuse_div){
        List paramList =new ArrayList() ;
		String sqlCmd 	= "select cmuse_id,cmuse_name from cdshareno where cmuse_div=? order by input_order";
		paramList.add(cmuse_div) ;
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"CMUSE_ID,CMUSE_NAME");
    	return dbData;
	}
    //申報本月無資料前檢查是否該月已有申報資料
    private String isHaveApplyData(String m_year,String m_month,String bank_no){
		List paramList =new ArrayList() ;
		StringBuffer sqlCmd = new StringBuffer() ;
		sqlCmd.append("select count(*) cnt from WLX10_M_LOAN "
						+ " where m_month = ? "
						+ " and m_year = ? "
						+ " and bank_no = ? "
						+ " and (loan_idn is not null or loan_name is not null) ");
		paramList.add(String.valueOf(m_month)) ;
		paramList.add(String.valueOf(m_year)) ;
		paramList.add(bank_no) ;
        List qList = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"cnt");
		String count =( ((DataObject)qList.get(0)).getValue("cnt")==null ) ? "null" :((DataObject)qList.get(0)).getValue("cnt").toString();
		if("null".equals(count)||"0".equals(count)){
		    return "N";
		}else{
		    return "Y";
		}
	}
    //申報本月無資料前檢查是否該月已申報過無資料
    private String isHaveNoApplyData(String m_year,String m_month,String bank_no){
		List paramList =new ArrayList() ;
		StringBuffer sqlCmd = new StringBuffer() ;
		sqlCmd.append("select count(*) cnt from WLX10_M_LOAN_APPLY "
						+ " where m_month = ? "
						+ " and m_year = ? "
						+ " and bank_no = ? ");
		paramList.add(String.valueOf(m_month)) ;
		paramList.add(String.valueOf(m_year)) ;
		paramList.add(bank_no) ;
        List qList = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"cnt");
		String count =( ((DataObject)qList.get(0)).getValue("cnt")==null ) ? "null" :((DataObject)qList.get(0)).getValue("cnt").toString();
		if("null".equals(count)||"0".equals(count)){
		    return "N";
		}else{
		    return "Y";
		}
	}
  	//申報本月無資料前檢查是否該月已申報過無資料但有實際授信餘額
    private String isHaveNoApplyLoanData(String m_year,String m_month,String bank_no){
		List paramList =new ArrayList() ;
		StringBuffer sqlCmd = new StringBuffer() ;
		sqlCmd.append("select count(*) cnt from WLX10_M_LOAN "
						+ " where m_month = ? "
						+ " and m_year = ? "
						+ " and bank_no = ? "
						+ " and (loan_idn is null or loan_name is null) ");
		paramList.add(String.valueOf(m_month)) ;
		paramList.add(String.valueOf(m_year)) ;
		paramList.add(bank_no) ;
        List qList = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"cnt");
		String count =( ((DataObject)qList.get(0)).getValue("cnt")==null ) ? "null" :((DataObject)qList.get(0)).getValue("cnt").toString();
		if("null".equals(count)||"0".equals(count)){
		    return "N";
		}else{
		    return "Y";
		}
	}
  	//檢查有無鎖定===========================
    private String isLock(String m_year,String m_month,String bank_no){
    		List paramList =new ArrayList() ;
    		StringBuffer sqlCmd = new StringBuffer() ;
    		sqlCmd.append("select lock_status from WML01 "
    						+ " where m_year = ? "
    						+ " and m_month =? "
    						+ " and bank_code = ? "
    						+ " and report_no= ? ");
    		paramList.add(m_year) ;
    		paramList.add(m_month) ;
    		paramList.add(bank_no) ;
    		paramList.add("A11") ;
            List qList1 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"lock_status");
            System.out.println("qList1.size()="+ qList1.size());
            sqlCmd.setLength(0) ;
			paramList.clear() ;
			sqlCmd.append("select lock_status from WML01_LOCK "
						+ " where m_year = ? "
						+ " and m_month =? "
						+ " and bank_code = ? "
						+ " and report_no= ? ");
			paramList.add(m_year) ;
			paramList.add(m_month) ;
			paramList.add(bank_no) ;
			paramList.add("A11") ;
		    List qList2 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
		    System.out.println("qList2.size()="+ qList2.size());
		    
		    String lock_status1 ="";
		    String lock_status2 ="";
		    if(qList1.size()>0){
		    	lock_status1 =( ((DataObject)qList1.get(0)).getValue("lock_status")==null ) ? "null" :(String)((DataObject)qList1.get(0)).getValue("lock_status");
		    }
		    if(qList2.size()>0){
		    	lock_status2 =( ((DataObject)qList2.get(0)).getValue("lock_status")==null ) ? "null" :(String)((DataObject)qList2.get(0)).getValue("lock_status");
		    }
		    if("Y".equals(lock_status1) || "Y".equals(lock_status2)){
		        return "Y";
		    }else{
		        return "N";
		    }
    }
    //是否上個月有無申報資料 (BY m_year,m_quarter,bank_no)===========================
    private String isLastMonthData(String m_year,String m_month,String bank_no){
		List paramList =new ArrayList() ;
		StringBuffer sqlCmd = new StringBuffer() ;
		sqlCmd.append("select count(*) cnt from WLX10_M_LOAN "
						+ " where m_month = ? "
						+ " and m_year = ? "
						+ " and bank_no = ? ");
		int last_year = 0;
		int last_month = Integer.parseInt(m_month)-1 ;
		if(last_month==0){
		    last_month = 12;
		    last_year = Integer.parseInt(m_year)-1;
		}else{
		    last_year = Integer.parseInt(m_year);
		}
		paramList.add(String.valueOf(last_month)) ;
		paramList.add(String.valueOf(last_year)) ;
		paramList.add(bank_no) ;
        List qList = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"cnt");
		String count =( ((DataObject)qList.get(0)).getValue("cnt")==null ) ? "null" :((DataObject)qList.get(0)).getValue("cnt").toString();
		if("null".equals(count)||"0".equals(count)){
		    return "N";
		}else{
		    return "Y";
		}
	}
   
    
    //插入新增的資料 ()=================================================================
    public String InsertDB(HttpServletRequest request,String lguser_id,String lguser_name) throws Exception{    	
  		StringBuffer sqlCmd = new StringBuffer() ;		
  		String errMsg = "";
  		String case_no      =((String)request.getParameter("case_no")==null)?"":(String)request.getParameter("case_no");
  		String case_cnt     =((String)request.getParameter("case_cnt")==null)?"":(String)request.getParameter("case_cnt");
  		String m_year		=((String)request.getParameter("hyear")==null)?"":(String)request.getParameter("hyear");
  		String m_month 		=((String)request.getParameter("hmonth")==null)?"":(String)request.getParameter("hmonth");
  		String bank_no 		=((String)request.getParameter("hbank_no")==null)?"":(String)request.getParameter("hbank_no");
  		String loan_idn 	=((String)request.getParameter("loan_idn")==null)?"":(String)request.getParameter("loan_idn");
  		String loan_name 	=((String)request.getParameter("loan_name")==null)?"":(String)request.getParameter("loan_name");
  		String loan_amt_sum 	=((String)request.getParameter("loan_amt_sum")==null)?"":(String)request.getParameter("loan_amt_sum");
  		String case_begin_year 	=((String)request.getParameter("case_begin_year")==null)?"":(String)request.getParameter("case_begin_year");
  		String case_begin_month =((String)request.getParameter("case_begin_month")==null)?"":(String)request.getParameter("case_begin_month");
  		String case_end_year 	=((String)request.getParameter("case_end_year")==null)?"":(String)request.getParameter("case_end_year");
  		String case_end_month 	=((String)request.getParameter("case_end_month")==null)?"":(String)request.getParameter("case_end_month");
  		String bank_no_max 	=((String)request.getParameter("bank_no_max")==null)?"":(String)request.getParameter("bank_no_max");
  		String manabank_name =((String)request.getParameter("manabank_name")==null)?"":(String)request.getParameter("manabank_name");
  		String loan_kind 	=((String)request.getParameter("loan_kind")==null)?"":(String)request.getParameter("loan_kind");
  		String loan_amt 	=((String)request.getParameter("loan_amt")==null)?"":(String)request.getParameter("loan_amt");
  		String loan_bal_amt =((String)request.getParameter("loan_bal_amt")==null)?"":(String)request.getParameter("loan_bal_amt");
  		String loan_type 	=((String)request.getParameter("loan_type")==null)?"":(String)request.getParameter("loan_type");
  		String pay_state 	=((String)request.getParameter("pay_state")==null)?"":(String)request.getParameter("pay_state");
  		String violate_type =((String)request.getParameter("violate_type")==null)?"":(String)request.getParameter("violate_type");
  		String loan_rate 	=((String)request.getParameter("loan_rate")==null)?"":(String)request.getParameter("loan_rate");
  		String new_case 	=((String)request.getParameter("new_case")==null)?"":(String)request.getParameter("new_case");
  		try {
  			List max_seq_number = new LinkedList();
  			List<List> updateDBList = new ArrayList<List>();//0:sql 1:data		
  			List<String> paramList =  new ArrayList<String>();//儲存參數的data
  			
  			//取得seq_no
  			sqlCmd.delete(0, sqlCmd.length());
  			paramList.clear() ;
  			sqlCmd.append("select to_char(max(SEQ_NO)+1) as maxq from wlx10_m_loan where m_year=? "
  						 + "and m_month=? "
  						 + "and bank_no=? ");			    
  			paramList.add(m_year) ;
  			paramList.add(m_month) ;
  			paramList.add(bank_no) ;
  			max_seq_number = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");	
  		    System.out.println("max_seq_number="+ max_seq_number.size());
  			String max_seq = (((DataObject)max_seq_number.get(0)).getValue("maxq")==null ) ? "1" :(String)((DataObject)max_seq_number.get(0)).getValue("maxq");
  		    
     		sqlCmd.delete(0, sqlCmd.length());
  			paramList.clear() ;
     		sqlCmd.append(" INSERT INTO WLX10_M_LOAN (") ;
     		sqlCmd.append(" M_YEAR,M_MONTH,BANK_NO,SEQ_NO,LOAN_IDN,LOAN_NAME,LOAN_AMT_SUM,");
     		sqlCmd.append(" CASE_BEGIN_YEAR,CASE_BEGIN_MONTH,CASE_END_YEAR,CASE_END_MONTH,");
     		sqlCmd.append(" BANK_NO_MAX,MANABANK_NAME,LOAN_KIND,LOAN_AMT,LOAN_BAL_AMT,LOAN_TYPE,");
     		sqlCmd.append(" PAY_STATE,VIOLATE_TYPE,LOAN_RATE,NEW_CASE,USER_ID,USER_NAME,UPDATE_DATE,CASE_NO,CASE_CNT) VALUES(");
     		sqlCmd.append(" ?,?,?,?,?,?,?, ?,?,?,?, ?,?,?,?,?,?, ?,?,?,?,?,?,sysdate,?,? ) ");
     		paramList.add(m_year) ;
     		paramList.add(m_month) ;
     		paramList.add(bank_no) ;
     		paramList.add(max_seq) ; 
     		paramList.add(loan_idn) ;
     		paramList.add(loan_name) ;
     		paramList.add(loan_amt_sum) ;
     		paramList.add(case_begin_year) ;
     		paramList.add(case_begin_month) ;
     		paramList.add(case_end_year) ;
     		paramList.add(case_end_month) ;
     		paramList.add(bank_no_max) ;
     		paramList.add(manabank_name) ;
     		paramList.add(loan_kind) ;
     		paramList.add(loan_amt) ;
     		paramList.add(loan_bal_amt) ;
     		paramList.add(loan_type) ;
     		paramList.add(pay_state) ;
     		paramList.add(violate_type) ;
     		paramList.add(loan_rate) ;
     		paramList.add(new_case) ;
     		paramList.add(lguser_id) ;
     		paramList.add(lguser_name) ;
     		paramList.add(case_no) ;
     		paramList.add(case_cnt) ;
     		if(!this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)) {   			
     			errMsg = errMsg + "相關資料寫入WLX10_M_LOAN資料庫失敗 <br>[DBManager.getErrMsg()]:<br>" 
  							+ DBManager.getErrMsg();
     		}
     			
     		sqlCmd.delete(0, sqlCmd.length());
  			paramList.clear() ;
  			sqlCmd.append("SELECT count(*) as data from WML01 "
  							+ " WHERE m_year=? AND m_month=? "
  							+ " AND bank_code=? AND report_no=? " ) ;
  			paramList.add(m_year) ;
  			paramList.add(m_month);
  			paramList.add(bank_no) ;
  			paramList.add("A11") ;
  			List qList = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"data");
  			if(qList != null && qList.size() > 0){
  				int data = Integer.parseInt(((DataObject)qList.get(0)).getValue("data").toString());
  				if( data > 0 ){
  				    sqlCmd.delete(0, sqlCmd.length());
  					paramList.clear() ;
  					sqlCmd.append(" insert into wml01_log ") ;
  					sqlCmd.append(" (m_year,m_month,bank_code,report_no,input_method,add_user,add_name,add_date,upd_code,user_id,user_name,update_date,user_id_c,user_name_c,update_date_c,update_type_c) ") ;
  					sqlCmd.append(" select m_year,m_month,bank_code,report_no,input_method,add_user,add_name,add_date,upd_code,user_id,user_name,update_date, ? , ? ,sysdate, ? ") ;
  					sqlCmd.append(" from wml01 ") ;
  					sqlCmd.append(" where m_year= ? and m_month= ? and bank_code= ? and report_no= ? ") ;
  					paramList.add(lguser_id) ;
  		   			paramList.add(lguser_name) ;
  		   			paramList.add("U") ;
  					paramList.add(m_year) ;
  		   			paramList.add(m_month) ;
  		   			paramList.add(bank_no) ;
  		   			paramList.add("A11") ;
  		   			if(!this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)) {   			
  		   				errMsg = errMsg + "相關資料寫入wml01_log資料庫失敗 <br>[DBManager.getErrMsg()]:<br>" 
  									+ DBManager.getErrMsg();
  		   			}
  		   			sqlCmd.delete(0, sqlCmd.length());
  					paramList.clear() ;
  		   			sqlCmd.append(" update wml01 ") ;
  		   			sqlCmd.append(" set user_id=?, user_name=?, update_date=sysdate ");
  		   			sqlCmd.append(" where m_year=? and m_month=? and bank_code=? and report_no=? ");
  		   			paramList.add(lguser_id) ;
  		   			paramList.add(lguser_name) ;
  					paramList.add(m_year) ;
  		   			paramList.add(m_month) ;
  		   			paramList.add(bank_no) ;
  		   			paramList.add("A11") ;
  				} else {
  				    sqlCmd.setLength(0) ;
  					paramList.clear() ;
  					sqlCmd.append(" insert into wml01 ") ;
  					sqlCmd.append(" (m_year,m_month,bank_code,report_no,input_method,add_user,add_name,add_date,upd_code,user_id,user_name,update_date) ") ;
  					sqlCmd.append(" values ( ?,?,?,?,?,?,?,sysdate,?,?,?,sysdate) ") ;
  					paramList.add(m_year) ;
  		   			paramList.add(m_month) ;
  		   			paramList.add(bank_no) ;
  		   			paramList.add("A11") ;
  		   			paramList.add("W") ;
  		   			paramList.add(lguser_id) ;
  		   			paramList.add(lguser_name) ;
  		   			paramList.add("U") ;
  		   			paramList.add(lguser_id) ;
  		   			paramList.add(lguser_name) ;
  				}
  				if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){
		   			if(!(case_cnt).equals(String.valueOf(Integer.parseInt(case_no.substring(7, 9))))){
				        errMsg = errMsg + "EditNextCnt";
				    }else{
   						errMsg = errMsg + "相關資料寫入資料庫成功";
				    }
	   				}else{
	   					errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" 
	   									+ DBManager.getErrMsg();
	   				}
  			}
     			
  		}catch (Exception e){
  			System.out.println(e+":"+e.getMessage());
  			errMsg = errMsg + "相關資料寫入資料庫失敗";					
  		}//end of catch	
  		
  		return errMsg;
  	} //end of insert
	
    //記錄本月沒有資料需要申報 ()=================================================================
    //102.12.26 add 本月無申報資料時,也要寫入申報記錄檔WML01
  	public String No_ApplyDB(HttpServletRequest request,String lguser_id,String lguser_name) throws Exception{    	
		StringBuffer sqlCmd = new StringBuffer() ;		
		String errMsg = "";
		List paramList = new ArrayList() ;
		String m_year	 =((String)request.getParameter("s_year")==null)?"":(String)request.getParameter("s_year");
		String m_month   =((String)request.getParameter("s_month")==null)?"":(String)request.getParameter("s_month");
		String bank_no 	 =((String)request.getParameter("bank_no")==null)?"":(String)request.getParameter("bank_no");
		try {
			String isHaveNoApplyData = isHaveNoApplyData(m_year,m_month,bank_no);
			if("Y".equals(isHaveNoApplyData)){
				errMsg = errMsg + "此筆資料已存在無法新增<br>";
			} else {
				sqlCmd.setLength(0) ;
				paramList.clear() ;
				sqlCmd.append("INSERT INTO WLX10_M_LOAN_APPLY (M_YEAR,M_MONTH,BANK_NO,APPLY_CNT,"
	  						+ "USER_ID,USER_NAME,UPDATE_DATE) VALUES("                   
							+ "?"        
							+ ",?"			            	      
							+ ",?"
							+ ",'0'"
							+ ",?"
							+ ",?"
							+ ",sysdate)");         		 						
				paramList.add(m_year) ;
				paramList.add(m_month) ;
				paramList.add(bank_no) ;
				paramList.add(lguser_id) ;
				paramList.add(lguser_name) ;
	   			if(!this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){
	   				errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" 
	   									+ DBManager.getErrMsg();
	   			}
	   			//102.12.26 add 本月無申報資料時,也要寫入申報記錄檔WML01
	   			sqlCmd.delete(0, sqlCmd.length());
  			    paramList.clear() ;
  			    sqlCmd.append("SELECT count(*) as data from WML01 "
  			    				+ " WHERE m_year=? AND m_month=? "
  			    				+ " AND bank_code=? AND report_no=? " ) ;
  			    paramList.add(m_year) ;
  			    paramList.add(m_month);
  			    paramList.add(bank_no) ;
  			    paramList.add("A11") ;
  			    List qList = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"data");
  			    if(qList != null && qList.size() > 0){
  			    	int data = Integer.parseInt(((DataObject)qList.get(0)).getValue("data").toString());
  			    	if( data > 0 ){
  			    	    sqlCmd.delete(0, sqlCmd.length());
  			    		paramList.clear() ;
  			    		sqlCmd.append(" insert into wml01_log ") ;
  			    		sqlCmd.append(" (m_year,m_month,bank_code,report_no,input_method,add_user,add_name,add_date,upd_code,user_id,user_name,update_date,user_id_c,user_name_c,update_date_c,update_type_c) ") ;
  			    		sqlCmd.append(" select m_year,m_month,bank_code,report_no,input_method,add_user,add_name,add_date,upd_code,user_id,user_name,update_date, ? , ? ,sysdate, ? ") ;
  			    		sqlCmd.append(" from wml01 ") ;
  			    		sqlCmd.append(" where m_year= ? and m_month= ? and bank_code= ? and report_no= ? ") ;
  			    		paramList.add(lguser_id) ;
  		   	    		paramList.add(lguser_name) ;
  		   	    		paramList.add("U") ;
  			    		paramList.add(m_year) ;
  		   	    		paramList.add(m_month) ;
  		   	    		paramList.add(bank_no) ;
  		   	    		paramList.add("A11") ;
  		   	    		if(!this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)) {   			
  		   	    			errMsg = errMsg + "相關資料寫入wml01_log資料庫失敗 <br>[DBManager.getErrMsg()]:<br>" 
  			    						+ DBManager.getErrMsg();
  		   	    		}
  		   	    		sqlCmd.delete(0, sqlCmd.length());
  			    		paramList.clear() ;
  		   	    		sqlCmd.append(" update wml01 ") ;
  		   	    		sqlCmd.append(" set user_id=?, user_name=?, update_date=sysdate ");
  		   	    		sqlCmd.append(" where m_year=? and m_month=? and bank_code=? and report_no=? ");
  		   	    		paramList.add(lguser_id) ;
  		   	    		paramList.add(lguser_name) ;
  			    		paramList.add(m_year) ;
  		   	    		paramList.add(m_month) ;
  		   	    		paramList.add(bank_no) ;
  		   	    		paramList.add("A11") ;
  			    	} else {
  			    	    sqlCmd.setLength(0) ;
  			    		paramList.clear() ;
  			    		sqlCmd.append(" insert into wml01 ") ;
  			    		sqlCmd.append(" (m_year,m_month,bank_code,report_no,input_method,add_user,add_name,add_date,upd_code,user_id,user_name,update_date) ") ;
  			    		sqlCmd.append(" values ( ?,?,?,?,?,?,?,sysdate,?,?,?,sysdate) ") ;
  			    		paramList.add(m_year) ;
  		   	    		paramList.add(m_month) ;
  		   	    		paramList.add(bank_no) ;
  		   	    		paramList.add("A11") ;
  		   	    		paramList.add("W") ;
  		   	    		paramList.add(lguser_id) ;
  		   	    		paramList.add(lguser_name) ;
  		   	    		paramList.add("U") ;
  		   	    		paramList.add(lguser_id) ;
  		   	    		paramList.add(lguser_name) ;
  			    	}
  			    	if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){
   			    		errMsg = errMsg + "相關資料寫入資料庫成功";			    	   
	   		    	}else{
	   		    		errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" 
	   		    						+ DBManager.getErrMsg();
	   		    	}
  			    }
	   			
	   			
   			}
		}catch (Exception e){
			System.out.println(e+":"+e.getMessage());
			errMsg = errMsg + "相關資料寫入資料庫失敗";						
		}//end of catch	
		return errMsg;
	} //end of insert
  	
	//修改所點選的資料 ()==================================================
	  public String UpdateDB(HttpServletRequest request,String lguser_id,String lguser_name) throws Exception{    	
		StringBuffer sqlCmd = new StringBuffer() ;		
		String errMsg = "";
		String m_year		=((String)request.getParameter("hyear")==null)?"":(String)request.getParameter("hyear");
		String m_month 		=((String)request.getParameter("hmonth")==null)?"":(String)request.getParameter("hmonth");
		String bank_no 		=((String)request.getParameter("hbank_no")==null)?"":(String)request.getParameter("hbank_no");
		String seq_no		=((String)request.getParameter("hseq_no")==null)?"":(String)request.getParameter("hseq_no");
		String loan_idn 	=((String)request.getParameter("loan_idn")==null)?"":(String)request.getParameter("loan_idn");
		String loan_name 	=((String)request.getParameter("loan_name")==null)?"":(String)request.getParameter("loan_name");
		String loan_amt_sum 	=((String)request.getParameter("loan_amt_sum")==null)?"":(String)request.getParameter("loan_amt_sum");
		String case_begin_year 	=((String)request.getParameter("case_begin_year")==null)?"":(String)request.getParameter("case_begin_year");
		String case_begin_month =((String)request.getParameter("case_begin_month")==null)?"":(String)request.getParameter("case_begin_month");
		String case_end_year 	=((String)request.getParameter("case_end_year")==null)?"":(String)request.getParameter("case_end_year");
		String case_end_month 	=((String)request.getParameter("case_end_month")==null)?"":(String)request.getParameter("case_end_month");
		String bank_no_max 	=((String)request.getParameter("bank_no_max")==null)?"":(String)request.getParameter("bank_no_max");
		String manabank_name =((String)request.getParameter("manabank_name")==null)?"":(String)request.getParameter("manabank_name");
		String loan_kind 	=((String)request.getParameter("loan_kind")==null)?"":(String)request.getParameter("loan_kind");
		String loan_amt 	=((String)request.getParameter("loan_amt")==null)?"":(String)request.getParameter("loan_amt");
		String loan_bal_amt =((String)request.getParameter("loan_bal_amt")==null)?"":(String)request.getParameter("loan_bal_amt");
		String loan_type 	=((String)request.getParameter("loan_type")==null)?"":(String)request.getParameter("loan_type");
		String pay_state 	=((String)request.getParameter("pay_state")==null)?"":(String)request.getParameter("pay_state");
		String violate_type =((String)request.getParameter("violate_type")==null)?"":(String)request.getParameter("violate_type");
		String loan_rate 	=((String)request.getParameter("loan_rate")==null)?"":(String)request.getParameter("loan_rate");
		String new_case 	=((String)request.getParameter("new_case")==null)?"":(String)request.getParameter("new_case");
		try {
			List paramList =new ArrayList() ;
			sqlCmd.append("select * from WLX10_M_LOAN where m_year = ?"
							 + "and m_month = ?"
							 + "and bank_no = ?"
							 + "and seq_no =?");			
			paramList.add(m_year) ;
			paramList.add(m_month) ;
			paramList.add(bank_no);
			paramList.add(seq_no);
			List qList = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");		
				
			 if (qList.size() == 0){
				errMsg = errMsg + "此筆資料不存在無法修改<br>";
			 }else{
				sqlCmd.setLength(0) ;
				paramList.clear() ;
				//Insert to log ---------------
				sqlCmd.append(" insert into wlx10_m_loan_log") ;
				sqlCmd.append(" select m_year,m_month,bank_no,seq_no,loan_idn,loan_name,loan_amt_sum,case_begin_year,case_begin_month,case_end_year,") ;
				sqlCmd.append(" case_end_month,bank_no_max,manabank_name,loan_kind,loan_amt,loan_bal_amt,loan_type,pay_state,violate_type,loan_rate,") ;
				sqlCmd.append(" new_case,user_id,user_name,update_date,?,?,sysdate,?,case_no,case_cnt ") ;
				sqlCmd.append(" from wlx10_m_loan ") ;
				sqlCmd.append(" where m_year=? and m_month=? and bank_no=? and seq_no=? ") ;
				paramList.add(lguser_id) ;
	   			paramList.add(lguser_name) ;
	   			paramList.add("U") ;
				paramList.add(m_year) ;
	   			paramList.add(m_month) ;
	   			paramList.add(bank_no) ;
	   			paramList.add(seq_no) ;
	   			if(!this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){
	   				errMsg = errMsg + "相關資料 wlx10_m_loan_log寫入失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg()+"\n";
	   			}	
			    
			  	//Update ---------------  
			    paramList.clear() ;
			    sqlCmd.setLength(0) ;
				sqlCmd.append("UPDATE wlx10_m_loan SET ") ;        
				sqlCmd.append( "LOAN_IDN=?"); paramList.add(loan_idn) ;       
				sqlCmd.append( ",LOAN_NAME=?"); paramList.add(loan_name);      
				sqlCmd.append( ",LOAN_AMT_SUM=?");paramList.add(loan_amt_sum) ; 
				sqlCmd.append( ",CASE_BEGIN_YEAR=?"); paramList.add(case_begin_year);      
				sqlCmd.append( ",CASE_BEGIN_MONTH=?"); paramList.add(case_begin_month);       
				sqlCmd.append( ",CASE_END_YEAR=?"); paramList.add(case_end_year);    
				sqlCmd.append( ",CASE_END_MONTH=?"); paramList.add(case_end_month);   
				sqlCmd.append( ",BANK_NO_MAX=?"); paramList.add(bank_no_max);  
				sqlCmd.append( ",MANABANK_NAME=?"); paramList.add(manabank_name);     
				sqlCmd.append( ",LOAN_KIND=?"); paramList.add(loan_kind); 
				sqlCmd.append( ",LOAN_AMT=?"); paramList.add(loan_amt);					       
				sqlCmd.append( ",LOAN_BAL_AMT=?"); paramList.add(loan_bal_amt) ;     
				sqlCmd.append( ",LOAN_TYPE=?"); paramList.add(loan_type);
				sqlCmd.append( ",PAY_STATE=?");paramList.add(pay_state);
				sqlCmd.append( ",VIOLATE_TYPE=?");paramList.add(violate_type);
				sqlCmd.append( ",LOAN_RATE=?");paramList.add(loan_rate);
				sqlCmd.append( ",NEW_CASE=?");paramList.add(new_case);
				sqlCmd.append( ",USER_ID=?");paramList.add(lguser_id);
				sqlCmd.append( ",USER_NAME=?");paramList.add(lguser_name);
				sqlCmd.append( ",UPDATE_DATE = sysdate");
				sqlCmd.append( " WHERE M_YEAR =?");paramList.add(m_year) ;
				sqlCmd.append( " and M_MONTH = ?");paramList.add(m_month);
				sqlCmd.append( " and BANK_NO = ?");paramList.add(bank_no);
				sqlCmd.append( " and SEQ_NO =?");paramList.add(seq_no);
	   			if(!this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){
	   				errMsg = errMsg + "相關資料wlx10_m_loan寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
	   			}
	   			
	   			sqlCmd.delete(0, sqlCmd.length());
				paramList.clear() ;
				sqlCmd.append("SELECT count(*) as data from WML01 "
								+ " WHERE m_year=? AND m_month=? "
								+ " AND bank_code=? AND report_no=? " ) ;
				paramList.add(m_year) ;
				paramList.add(m_month);
				paramList.add(bank_no) ;
				paramList.add("A11") ;
				List qData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"data");
				if(qData != null || qData.size() > 0){
					int data = Integer.parseInt(((DataObject)qData.get(0)).getValue("data").toString());
					if( data > 0 ){
					    sqlCmd.delete(0, sqlCmd.length());
						paramList.clear() ;
						sqlCmd.append(" insert into wml01_log ") ;
						sqlCmd.append(" (m_year,m_month,bank_code,report_no,input_method,add_user,add_name,add_date,upd_code,user_id,user_name,update_date,user_id_c,user_name_c,update_date_c,update_type_c) ") ;
						sqlCmd.append(" select m_year,m_month,bank_code,report_no,input_method,add_user,add_name,add_date,upd_code,user_id,user_name,update_date, ? , ? ,sysdate, ? ") ;
						sqlCmd.append(" from wml01 ") ;
						sqlCmd.append(" where m_year= ? and m_month= ? and bank_code= ? and report_no= ? ") ;
						paramList.add(lguser_id) ;
			   			paramList.add(lguser_name) ;
			   			paramList.add("U") ;
						paramList.add(m_year) ;
			   			paramList.add(m_month) ;
			   			paramList.add(bank_no) ;
			   			paramList.add("A11") ;
			   			if(!this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)) {   			
			   				errMsg = errMsg + "相關資料寫入wml01_log資料庫失敗 <br>[DBManager.getErrMsg()]:<br>" 
										+ DBManager.getErrMsg();
			   			}
			   			sqlCmd.delete(0, sqlCmd.length());
						paramList.clear() ;
			   			sqlCmd.append(" update wml01 ") ;
			   			sqlCmd.append(" set user_id=?, user_name=?, update_date=sysdate ");
			   			sqlCmd.append(" where m_year=? and m_month=? and bank_code=? and report_no=? ");
			   			paramList.add(lguser_id) ;
			   			paramList.add(lguser_name) ;
						paramList.add(m_year) ;
			   			paramList.add(m_month) ;
			   			paramList.add(bank_no) ;
			   			paramList.add("A11") ;
			   			if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){
	   						errMsg = errMsg + "相關資料寫入資料庫成功";
		   				}else{
		   					errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" 
		   									+ DBManager.getErrMsg();
		   				}
					} else {
					    sqlCmd.setLength(0) ;
						paramList.clear() ;
						sqlCmd.append(" insert into wml01 ") ;
						sqlCmd.append(" (m_year,m_month,bank_code,report_no,input_method,add_user,add_name,add_date,upd_code,user_id,user_name,update_date) ") ;
						sqlCmd.append(" values ( ?,?,?,?,?,?,?,sysdate,?,?,?,sysdate) ") ;
						paramList.add(m_year) ;
			   			paramList.add(m_month) ;
			   			paramList.add(bank_no) ;
			   			paramList.add("A11") ;
			   			paramList.add("W") ;
			   			paramList.add(lguser_id) ;
			   			paramList.add(lguser_name) ;
			   			paramList.add("U") ;
			   			paramList.add(lguser_id) ;
			   			paramList.add(lguser_name) ;
						if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){
		   					errMsg = errMsg + "相關資料寫入資料庫成功";
		   				}else{
		   					errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" 
		   									+ DBManager.getErrMsg();
		   				}
					}
				}
	   			
	   		}//END OF IF-ELSE
	   					    
		}catch (Exception e){
			System.out.println(e+":"+e.getMessage());
			errMsg = errMsg + "相關資料寫入資料庫失敗";							
		}//end of catch	
				
		return errMsg;
	} //end of update
	
	//刪除所點選的資料 ()====================================================
	public String DeleteDB(HttpServletRequest request,String lguser_id,String lguser_name) throws Exception{    	
		StringBuffer sqlCmd = new StringBuffer();		
		String errMsg = "";
		String m_year		=((String)request.getParameter("hyear")==null)?"":(String)request.getParameter("hyear");
		String m_month 		=((String)request.getParameter("hmonth")==null)?"":(String)request.getParameter("hmonth");
		String bank_no 		=((String)request.getParameter("hbank_no")==null)?"":(String)request.getParameter("hbank_no");	
		String seq_no		=((String)request.getParameter("hseq_no")==null)?"":(String)request.getParameter("hseq_no");
		
			try {
				List paramList =new ArrayList() ;
				sqlCmd.append("select * from WLX10_M_LOAN where m_year = ?"
							 + "and m_month = ?"
							 + "and bank_no = ?"
							 + "and seq_no =? ");			
				paramList.add(m_year) ;
				paramList.add(m_month) ;
				paramList.add(bank_no) ;
				paramList.add(seq_no) ;
				List qList = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");		
				
				paramList.clear() ;
			    sqlCmd.setLength(0) ;
			    sqlCmd.append("select * from WLX10_M_LOAN_APPLY where m_year = ?"
						 + "and m_month = ?"
						 + "and bank_no = ?");			
				paramList.add(m_year) ;
				paramList.add(m_month) ;
				paramList.add(bank_no) ;
				List qList2 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
			
			 if (qList.size() > 0 || qList2.size() > 0 ){
			 	
			    if(qList.size()>0){
			      	//Insert to log ==================================================
					paramList.clear() ;
				    sqlCmd.setLength(0) ;
					sqlCmd.append(" insert into wlx10_m_loan_log") ;
					sqlCmd.append(" select m_year,m_month,bank_no,seq_no,loan_idn,loan_name,loan_amt_sum,case_begin_year,case_begin_month,case_end_year,") ;
					sqlCmd.append(" case_end_month,bank_no_max,manabank_name,loan_kind,loan_amt,loan_bal_amt,loan_type,pay_state,violate_type,loan_rate,") ;
					sqlCmd.append(" new_case,user_id,user_name,update_date,?,?,sysdate,?,case_no,case_cnt ") ;
					sqlCmd.append(" from wlx10_m_loan ") ;
					sqlCmd.append(" where m_year=? and m_month=? and bank_no=? and seq_no=? ") ;
					paramList.add(lguser_id) ;
		   			paramList.add(lguser_name) ;
		   			paramList.add("D") ;
					paramList.add(m_year) ;
		   			paramList.add(m_month) ;
		   			paramList.add(bank_no) ;
		   			paramList.add(seq_no) ;
		   			if(!this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){
		   				errMsg = errMsg + "相關資料 wlx10_m_loan_log寫入失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg()+"\n";
		   			}	
				 	//Delete  =========================================================
					sqlCmd.setLength(0) ;
				    paramList.clear() ;
					sqlCmd.append("DELETE wlx10_m_loan "                
							+ " WHERE M_YEAR =?"
							+ " and M_MONTH = ?"
							+ " and BANK_NO = ?"
							+ " and SEQ_NO = ? ");
		        	paramList.add(m_year) ;
		        	paramList.add(m_month) ;
		        	paramList.add(bank_no) ;
		        	paramList.add(seq_no) ;
		   			
		   			if(!this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)) {
		   				errMsg = errMsg + "相關資料wlx10_m_loan刪除失敗";
		   			}
			    }
			    
			    if(qList2.size()>0){
				 	//Delete  =========================================================
					sqlCmd.setLength(0) ;
				    paramList.clear() ;
					sqlCmd.append("DELETE wlx10_m_loan_apply "                
							+ " WHERE M_YEAR =?"
							+ " and M_MONTH = ?"
							+ " and BANK_NO = ?");
		        	paramList.add(m_year) ;
		        	paramList.add(m_month) ;
		        	paramList.add(bank_no) ;
		   			if(!this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)) {
		   				errMsg = errMsg + "相關資料wlx10_m_loan_apply刪除失敗";
		   			}
			    }
			    
	   			sqlCmd.delete(0, sqlCmd.length());
				paramList.clear() ;
				sqlCmd.append("SELECT count(*) as data from WML01 "
								+ " WHERE m_year=? AND m_month=? "
								+ " AND bank_code=? AND report_no=? " ) ;
				paramList.add(m_year) ;
				paramList.add(m_month);
				paramList.add(bank_no) ;
				paramList.add("A11") ;
				List qData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"data");
				if(qData != null && qData.size() > 0){
					int data = Integer.parseInt(((DataObject)qData.get(0)).getValue("data").toString());
					if( data > 0 ){
					    sqlCmd.delete(0, sqlCmd.length());
						paramList.clear() ;
						sqlCmd.append(" insert into wml01_log ") ;
						sqlCmd.append(" (m_year,m_month,bank_code,report_no,input_method,add_user,add_name,add_date,upd_code,user_id,user_name,update_date,user_id_c,user_name_c,update_date_c,update_type_c) ") ;
						sqlCmd.append(" select m_year,m_month,bank_code,report_no,input_method,add_user,add_name,add_date,upd_code,user_id,user_name,update_date, ? , ? ,sysdate, ? ") ;
						sqlCmd.append(" from wml01 ") ;
						sqlCmd.append(" where m_year= ? and m_month= ? and bank_code= ? and report_no= ? ") ;
						paramList.add(lguser_id) ;
			   			paramList.add(lguser_name) ;
			   			paramList.add("D") ;
						paramList.add(m_year) ;
			   			paramList.add(m_month) ;
			   			paramList.add(bank_no) ;
			   			paramList.add("A11") ;
			   			if(!this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)) {   			
			   				errMsg = errMsg + "相關資料wml01_log寫入資料庫失敗 <br>[DBManager.getErrMsg()]:<br>" 
										+ DBManager.getErrMsg();
			   			}
			   			sqlCmd.delete(0, sqlCmd.length());
						paramList.clear() ;
			   			sqlCmd.append(" update wml01 ") ;
			   			sqlCmd.append(" set user_id=?, user_name=?, update_date=sysdate, upd_code=? ");
			   			sqlCmd.append(" where m_year=? and m_month=? and bank_code=? and report_no=? ");
			   			paramList.add(lguser_id) ;
			   			paramList.add(lguser_name) ;
			   			paramList.add("U") ;
						paramList.add(m_year) ;
			   			paramList.add(m_month) ;
			   			paramList.add(bank_no) ;
			   			paramList.add("A11") ;
			   			if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){
	   						errMsg = errMsg + "相關資料寫入資料庫成功";
		   				}else{
		   					errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" 
		   									+ DBManager.getErrMsg();
		   				}
			   			
					} else {
					    sqlCmd.setLength(0) ;
						paramList.clear() ;
						sqlCmd.append(" insert into wml01 ") ;
						sqlCmd.append(" (m_year,m_month,bank_code,report_no,input_method,add_user,add_name,add_date,upd_code,user_id,user_name,update_date) ") ;
						sqlCmd.append(" values ( ?,?,?,?,?,?,?,sysdate,?,?,?,sysdate) ") ;
						paramList.add(m_year) ;
			   			paramList.add(m_month) ;
			   			paramList.add(bank_no) ;
			   			paramList.add("A11") ;
			   			paramList.add("W") ;
			   			paramList.add(lguser_id) ;
			   			paramList.add(lguser_name) ;
			   			paramList.add("U") ;
			   			paramList.add(lguser_id) ;
			   			paramList.add(lguser_name) ;
						if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){
		   					errMsg = errMsg + "相關資料寫入資料庫成功";
		   				}else{
		   					errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" 
		   									+ DBManager.getErrMsg();
		   				}
					}
				}
	   		}			    
		}catch (Exception e){
			System.out.println(e+":"+e.getMessage());
			errMsg = errMsg + "相關資料刪除失敗";						
		}//end of catch		
			
		return errMsg;
	} //end of Delete
	
	
	//104.05.12 add 修改案件編號/分項項數=================================================
	public String updCase_No(HttpServletRequest request,String lguser_id,String lguser_name) throws Exception{    	
		StringBuffer sqlCmd = new StringBuffer();		
		String errMsg = "";
		String m_year		=((String)request.getParameter("m_year")==null)?"":(String)request.getParameter("m_year");
		String m_month 		=((String)request.getParameter("m_month")==null)?"":(String)request.getParameter("m_month");
		String bank_no 		=((String)request.getParameter("bank_no")==null)?"":(String)request.getParameter("bank_no");	
		String seq_no		=((String)request.getParameter("seq_no")==null)?"":(String)request.getParameter("seq_no");
		String case_no		=((String)request.getParameter("upd_case_no")==null)?"":(String)request.getParameter("upd_case_no");
		String case_cnt		=((String)request.getParameter("upd_case_cnt")==null)?"":(String)request.getParameter("upd_case_cnt");
		
			try {
				System.out.println("test1.case_no="+case_no);
				System.out.println("test1.case_cnt="+case_cnt);
				List paramList =new ArrayList() ;
				sqlCmd.append("select * from WLX10_M_LOAN where m_year = ?"
							 + "and m_month = ?"
							 + "and bank_no = ?"
							 + "and seq_no =? ");			
				paramList.add(m_year) ;
				paramList.add(m_month) ;
				paramList.add(bank_no) ;
				paramList.add(seq_no) ;
				List qList = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");		
				
			    if(qList.size()>0){
					sqlCmd.setLength(0) ;
				    paramList.clear() ;
					sqlCmd.append("Update wlx10_m_loan "  
					        + " set case_no= ?,case_cnt=?"
							+ " WHERE M_YEAR =?"
							+ " and M_MONTH = ?"
							+ " and BANK_NO = ?"
							+ " and SEQ_NO = ? ");
					paramList.add(case_no) ;
		        	paramList.add(case_cnt) ;		        			
		        	paramList.add(m_year) ;
		        	paramList.add(m_month) ;
		        	paramList.add(bank_no) ;
		        	paramList.add(seq_no) ;
		   			
		   			if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){
		   			   errMsg = errMsg + "案件編號更新成功";
		   			}else{
		   			  errMsg = errMsg + "案件編號更新失敗<br>[DBManager.getErrMsg()]:<br>" 
		   							  + DBManager.getErrMsg();
		   			}
			    }
			    	    
		}catch (Exception e){
			System.out.println(e+":"+e.getMessage());
			errMsg = errMsg + "案件編號更新失敗";						
		}//end of catch		
			
		return errMsg;
	} //end of Delete

	private String getMaxCase_No(String m_year,String m_month,String bank_no,String case_no){
		List paramList =new ArrayList() ;
		StringBuffer sqlCmd = new StringBuffer() ;
		sqlCmd.append("select max(case_no) mcase_no from WLX10_M_LOAN "
						+ " where m_month = ? "
						+ " and m_year = ? "
						+ " and bank_no = ? "
						+ " and case_no like ? ");
		paramList.add(m_month) ;
		paramList.add(m_year) ;
		paramList.add(bank_no) ;
		if(m_month.length()==1){
	        m_month = "0"+m_month;
	    } 
	    if(m_year.length()==1){
	        m_year = "00"+m_year;
	    }else if(m_year.length()==2){
	        m_year = "0"+m_year;
	    }
	    if(!"".equals(case_no)){
	        paramList.add(case_no.substring(0, 7)+"%") ;
	    }else{
	        paramList.add(m_year+m_month+"%") ;
	    }
        List qList = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"mcase_no");
        String mCase_No = (((DataObject)qList.get(0)).getValue("mcase_no")==null ) ? "" :(String)((DataObject)qList.get(0)).getValue("mcase_no");
        return mCase_No;
	}
	private String getCase_Cnt(String m_year,String m_month,String bank_no,String mCase_No){
		List paramList =new ArrayList() ;
		StringBuffer sqlCmd = new StringBuffer() ;
		sqlCmd.append("select case_cnt from WLX10_M_LOAN "
						+ " where m_month = ? "
						+ " and m_year = ? "
						+ " and bank_no = ? "
						+ " and case_no = ? ");
		paramList.add(m_month) ;
		paramList.add(m_year) ;
		paramList.add(bank_no) ;
		paramList.add(mCase_No) ;
        List qList = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"case_cnt");
        String case_cnt =((DataObject)qList.get(0)).getValue("case_cnt").toString();
        return case_cnt;
	}
	private String getFirstLoan_amt_sum(String m_year,String m_month,String bank_no,String mCase_No){
		List paramList =new ArrayList() ;
		StringBuffer sqlCmd = new StringBuffer() ;
		sqlCmd.append("select loan_amt_sum from WLX10_M_LOAN "
						+ " where m_month = ? "
						+ " and m_year = ? "
						+ " and bank_no = ? "
						+ " and case_no = ? ");
		paramList.add(m_month) ;
		paramList.add(m_year) ;
		paramList.add(bank_no) ;
		paramList.add(mCase_No) ;
        List qList = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"loan_amt_sum");
        String loan_amt_sum ="";
        if(qList.size()>0){
            if(((DataObject)qList.get(0)).getValue("loan_amt_sum")!=null){
                loan_amt_sum = ((DataObject)qList.get(0)).getValue("loan_amt_sum").toString();                
            }else{
                loan_amt_sum = "";
            }
            //loan_amt_sum =((DataObject)qList.get(0)).getValue("loan_amt_sum").toString();
        }
        return loan_amt_sum;
	}
private boolean updDbUsesPreparedStatement(String sql ,List paramList) throws Exception{
	List updateDBList = new ArrayList();//0:sql 1:data
    List updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
	List updateDBDataList = new ArrayList();//儲存參數的List
	
	updateDBDataList.add(paramList);
	updateDBSqlList.add(sql);
	updateDBSqlList.add(updateDBDataList);
	updateDBList.add(updateDBSqlList);
	return DBManager.updateDB_ps(updateDBList) ;
}
  
  	
%>      
