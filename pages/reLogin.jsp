<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<% 
System.out.println("reLogin begin");
String url = ( request.getParameter("url")==null ) ? "" : (String)request.getParameter("url");					
System.out.println("url="+url);
%>
<html>
<head>
<title>網際網路申報系統</title>
<link href="css/b51.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
</script>
</head>

<body background="images/bg_1.gif" leftmargin="0" topmargin="0" onLoad="">
<form method=post>
</form>
</body>
<script language="JavaScript">
 parent.top.location.replace("<%=url%>"); 
</script>
</html>
