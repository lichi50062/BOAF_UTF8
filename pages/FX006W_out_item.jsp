<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="java.util.*" %>
<%@ page import="java.util.LinkedList" %>
<%
  String sqlCmd = " select cmuse_name "
			    + " from cdshareno "
			    + " where cmuse_div='031' "
			    + " order by cmuse_id ";
    	
  List dbData = DBManager.QueryDB_SQLParam(sqlCmd,null,"cmuse_name");  
%>
<html>

<head>
<title>委外項目說明</title>
<link href="css/b51.css" rel="stylesheet" type="text/css">
</head>

<body bgColor="#d8efee">
<form method="post" name="out_item_data">
 <table width="460" border="0" align="center" cellpadding="0" cellspacing="0" height="2">
  <%DataObject bean = null;
    String out_item = "";
    for(int i=0;i<dbData.size();i++){
        bean = (DataObject)dbData.get(i);
        out_item = (String)bean.getValue("cmuse_name");
  %>
    <tr class="sbody">
     <td width=40 valign="top"><%=out_item.substring(0,out_item.indexOf("、"))%>、</td>
     <td><%=out_item.substring(out_item.indexOf("、")+1,out_item.length())%></td>   
    </tr>
  <%}%>
  </table>
</form>     
</body>
</html>
