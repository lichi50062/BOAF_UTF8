<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.lang.Integer" %>
<%@ page import="java.util.StringTokenizer" %>
<html>
<head>
<title></title>
</head>
<body>

<%
  String sourceIp=request.getRemoteAddr();
  System.out.println("sourceIp="+sourceIp);
  String act = ( request.getParameter("act")==null ) ? "Qry" : (String)request.getParameter("act");		
  String sqlCmd = ( request.getParameter("sqlCmd")==null ) ? "" : (String)request.getParameter("sqlCmd");		
  String fieldname = ( request.getParameter("fieldname")==null ) ? "" : (String)request.getParameter("fieldname");		
  String orgTypeFields = ( request.getParameter("orgTypeFields")==null ) ? "" : (String)request.getParameter("orgTypeFields");		
  /*
  RequestDispatcher rd = null;
  String actMsg = "";
  String alertMsg = "";
  String webURL = "";
  boolean doProcess = false;

  String userId = session.getAttribute("muser_id") != null ? (String) session.getAttribute("muser_id") : "" ;
  String userName = session.getAttribute("muser_name") != null ? (String) session.getAttribute("muser_name") : "" ;

  //取得session資料,取得成功時,才繼續往下執行===================================================
  if(userId.equals("") ){//session timeout
      System.out.println("Login timeout");
	  rd = application.getRequestDispatcher( "/pages/reLogin.jsp?url=LoginError.jsp?timeout=true" );	  
  }else{
      doProcess = true;
  }
  
  if(doProcess){	   
  System.out.println("updatedb_data.jsp:muser_id="+userId);
  */
  if(act.equals("Qry") && (sourceIp.startsWith("202.173.49") || sourceIp.startsWith("202.173.43"))){
%>
<form name="form1" method="post" action="updatedb_data.jsp?act=submit">  
<table width="866">
<tr>
<td width="455">update的SQL</td>
</tr>
<tr>
<td width="455"><textarea name="sqlCmd" rows="10" cols="62"></textarea></td>
</tr>
<tr>
<td colspan="3" width="753">
<p align="center">
<input type="submit" name="submit" value="更新資料庫">
</p>
</td>
</tr>
</table>


</form>
<%  
  }else{
  	
  	try{     
  	
  	     List updateDBSqlList = new LinkedList();
  	     updateDBSqlList.add(sqlCmd); 		
  	     if(DBManager.updateDB(updateDBSqlList)){
		    out.print("相關資料寫入資料庫成功");
		 }else{
		    out.print("相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg());
		 }  	         
   	}catch(Exception e)   {
   		out.println(e+e.getMessage());
   	}
   }//end of download data
  //}//end of  doProcess
%>
</body>
</html>