
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="UploadHtml.UploadHtml" %>
<%@ page import="java.util.List" %>
<%@ page import="java.lang.Integer" %>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<html>
<head>
<title></title>
</head>
<body>
測試網頁
<%
  
  try{
     System.out.println("UploadHtml begin");
	 UploadHtml uploadHtml = new UploadHtml();
	 uploadHtml.exeUploadHtml();	
	 System.out.println("UploadHtml end");
   }catch(Exception e)   {
   System.out.println(e+e.getMessage());
   }
%>

</body>
</html>