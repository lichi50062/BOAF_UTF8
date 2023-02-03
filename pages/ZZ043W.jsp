<%
// 98.08.24 add 線上產生(A01~A04)Axx_operation中間值 by 2295
// 99.12.12 fix sqlInjection by 2808
//102.07.03 add bank_no_hsien_id_ALL','農.漁會-各別機構/農.漁會-縣市別小計/全体農漁會縣市別小計/農.漁會總計/全体農漁會總計 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.*" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Enumeration" %>
<%@ page import="java.util.HashMap" %>
<%@ page import="java.util.Hashtable" %>
<%@ page import="java.util.StringTokenizer" %>
<%@ page import="java.util.*,java.io.*" %>
<%
	RequestDispatcher rd = null;
	String ftpMsg = "";
	String actMsg = "";	
	String alertMsg = "";	
	String webURL = "";	
	boolean doProcess = false;	
	
	//取得session資料,取得成功時,才繼續往下執行===================================================
	if(session.getAttribute("muser_id") == null){//session timeout	
       System.out.println("ZZ043W login timeout");   
	   rd = application.getRequestDispatcher( "/pages/reLogin.jsp?url=LoginError.jsp?timeout=true" );         	   
	   try{
          rd.forward(request,response);
       }catch(Exception e){
          System.out.println("forward Error:"+e+e.getMessage());
       }
    }else{
      doProcess = true;
    }    
	if(doProcess){//若muser_id資料時,表示登入成功====================================================================	
	
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");				
	
	
	System.out.println("act="+act);	
   
	//登入者資訊
	String lguser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		
	String lguser_name = ( session.getAttribute("muser_name")==null ) ? "" : (String)session.getAttribute("muser_name");		
	String lguser_type = ( session.getAttribute("muser_type")==null ) ? "" : (String)session.getAttribute("muser_type");				
	String lguser_tbank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");			
	session.setAttribute("nowtbank_no",null);//94.01.05 fix 沒有Bank_List,把所點選的Bank_no清除======	
	String szReport_no = ( request.getParameter("Report_no")==null ) ? "" : (String)request.getParameter("Report_no");		
	String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? "" : (String)request.getParameter("S_YEAR");				
    String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? "" : (String)request.getParameter("S_MONTH");				
    String E_YEAR = ( request.getParameter("E_YEAR")==null ) ? "" : (String)request.getParameter("E_YEAR");				
    String E_MONTH = ( request.getParameter("E_MONTH")==null ) ? "" : (String)request.getParameter("E_MONTH");				
    String yyymm = ( request.getParameter("yyymm")==null ) ? "" : (String)request.getParameter("yyymm");	
    String firstStatus = ( request.getParameter("firstStatus")==null ) ? "" : (String)request.getParameter("firstStatus");			    	
   	System.out.println("ZZ043W.lguser_id="+lguser_id);	    	
    System.out.println("ZZ043W.S_YEAR="+S_YEAR);		    
	System.out.println("ZZ043W.S_MONTH="+S_MONTH);		    
	System.out.println("ZZ043W.E_YEAR="+E_YEAR);		    
	System.out.println("ZZ043W.E_MONTH="+E_MONTH);
	
    if(!Utility.CheckPermission(request,report_no)){//無權限時,導向到LoginError.jsp
        rd = application.getRequestDispatcher( LoginErrorPgName );        
    }else{            
    	//set next jsp 	    	
    	if(act.equals("Qry")){                    	        	    
    	    rd = application.getRequestDispatcher( ListPgName +"?act=Qry&firstStatus="+firstStatus);            	        	        	    	    
    	}else if(act.equals("List")){                    	    
    	    List dbData = getQryResult(szReport_no,S_YEAR,S_MONTH,E_YEAR,E_MONTH);    	       
    	    request.setAttribute("reportList",dbData);    	     	        	        	     	    
    	    rd = application.getRequestDispatcher( ListPgName +"?act=Qry&Report_no="+szReport_no+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&E_YEAR="+E_YEAR+"&E_MONTH="+E_MONTH);            	        	        	    	    
    	}else if(act.equals("Generate")){
    	    GenerateOperation a = new GenerateOperation();
            a.exeGenerateOperation(szReport_no,S_YEAR,S_MONTH,"true",lguser_id);
            List paramList = new ArrayList();
            paramList.add(S_YEAR);
            paramList.add(S_MONTH);
            List dbData = DBManager.QueryDB_SQLParam(" select count(*) as datacount from "+szReport_no+"_operation where m_year= ? and m_month= ? ",paramList,"datacount");
            if(dbData != null && dbData.size() > 0){
               if(Integer.parseInt(((DataObject)dbData.get(0)).getValue("datacount").toString()) > 0){
                  actMsg = szReport_no+"彈性報表資料已產生完成!!";                  
               }
            }else{
               actMsg =  szReport_no+"彈性報表資料產生失敗!!";
            }
            out.print(actMsg);
    	    rd = application.getRequestDispatcher( nextPgName);      				            
    	}	     	
    	request.setAttribute("actMsg",actMsg);    
    }        
     
%>

<%
	try {
        //forward to next present jsp
        rd.forward(request, response);
    } catch (NullPointerException npe) {
    }
    }//end of doProcess
%>


<%!
    private final static String report_no = "ZZ043W";
    private final static String nextPgName = "/pages/ActMsg.jsp";        
    private final static String ListPgName = "/pages/"+report_no+"_List.jsp";        
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    
    //取得查詢結果
    private List getQryResult(String szReport_no,String S_YEAR,String S_MONTH,String E_YEAR,String E_MONTH){    	       
    		//查詢條件        		
    		StringBuffer sqlCmd = new StringBuffer() ;
    		List paramList = new ArrayList() ;
            sqlCmd.append(" select decode(report_no,'A01','DS001W(A01營運明細資料)','A02','DS002W(A02法定比率資料)',"
	   			   + "		         'A03','DS003W(A03平均利率資料)','A04','DS004W(A04資產品質分析資料)',report_no) as s_report_name,"
                   + "        decode(kind_type,'bank_no_ALL','農.漁會各別機構','hsien_id_6','農會.縣市別小計',"
                   + "                         'hsien_id_7','漁會.縣市別小計','hsien_id_ALL','全体農漁會.縣市別小計',"
				   + "		                   'bank_no','農.漁會各別機構','bank_no_6','農會.各別機構',"
				   + " 					       'bank_no_7','漁會.各別機構','bank_no_67','農.漁會.各別機構','bank_no_hsien_id_ALL','農.漁會-各別機構/農.漁會-縣市別小計/全体農漁會縣市別小計/農.漁會總計/全体農漁會總計',kind_type) as kink_type,"
				   + "		  (to_char(m_year)|| LPAD(to_char(m_month),2,'0')) as s_yymm,total,user_id,"
				   + "		  ((to_char(UPDATE_DATE,'yyyymmdd')-19110000)  ||  to_char(UPDATE_DATE,' hh24:mi'))  as S_UpdateDate "							   
				   + " from operation_log "
				   + " where report_no=? "
				   + " and to_char(m_year * 100 + m_month) >= ?" 
				   + " and to_char(m_year * 100 + m_month) <= ?"				   
				   + " order by m_year,m_month,update_date desc " );
		    paramList.add(szReport_no) ;
		    paramList.add(S_YEAR+S_MONTH)  ;
		    paramList.add(E_YEAR+E_MONTH)  ;
    		//System.out.println(sqlCmd);
    		List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"total");
    		if(dbData != null && dbData.size() != 0){
    		   System.out.println("dbData.size="+dbData.size());  
    		}else{
    		   System.out.println("dbData is null or size = 0");  
    		}
    		
			return dbData;
    }
    
	
%>    