<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@include file="common.jsp"%>
<%
	String lguser_id = (session.getAttribute("muser_id") == null)?"":(String)session.getAttribute("muser_id"); 
	String lguser_name = (session.getAttribute("muser_name") == null)?"":(String)session.getAttribute("muser_name"); 
	String lguser_telno = (session.getAttribute("muser_telno") == null)?"":(String)session.getAttribute("muser_telno"); 
	String lguser_email = (session.getAttribute("muser_email") == null)?"":(String)session.getAttribute("muser_email"); 
	String maintain_user_id = "";
	String maintain_user_name = "";
	String maintain_update_date = "";
	
	String sqlCmd = "";
	if(request.getAttribute("maintainInfo") != null){
	   sqlCmd = (String)request.getAttribute("maintainInfo");
	}   
	System.out.println("maintainInfo.sqlCmd = "+sqlCmd);
	List dbData = null;
	if(!sqlCmd.equals("")){	   	  
	   dbData = DBManager.QueryDB_SQLParam(sqlCmd,null,"update_date");
	   if(dbData != null && dbData.size()!= 0){
	      maintain_user_id = (((DataObject)dbData.get(0)).getValue("user_id") == null )?"":(String)((DataObject)dbData.get(0)).getValue("user_id");
	      maintain_user_name = (((DataObject)dbData.get(0)).getValue("user_name") == null)?"":(String)((DataObject)dbData.get(0)).getValue("user_name");
	      maintain_update_date = (((DataObject)dbData.get(0)).getValue("update_date") == null)?"":Utility.getCHTdate((((DataObject)dbData.get(0)).getValue("update_date")).toString().substring(0, 10), 0);	      
	   }
	}
	maintain_user_name = (maintain_user_name.length()<3)?maintain_user_name:Utility.maskChar(maintain_user_name, 2, maintain_user_name.length()-2, "＊");
	lguser_name = (lguser_name.length()<3)?lguser_name:Utility.maskChar(lguser_name, 2, lguser_name.length()-2, "＊");
	String width=(request.getParameter("width") == null)?"600":(String)request.getParameter("width"); 	
%>

<table width=<%=width%> border='1' align='center' cellpadding="1" cellspacing="1" bordercolor="#3A9D99" class="sbody">
  <%if(dbData != null && dbData.size()!= 0){%>
  <tr bgcolor="#9AD3D0"> 
    <td colspan='2' class="sbody" align="center"><font color='#000000'>異動者資訊</font></td>
  </tr>
  <tr> 
    <td width='15%' bgcolor='#D8EFEE' align='left'>帳號</td>
	<td width='85%' bgcolor='e7e7e7'>&nbsp;<%=maintain_user_id%></td>				    	
  </tr>
  <tr> 
    <td width='15%' bgcolor='#D8EFEE' align='left'>姓名</td>
	<td width='85%' bgcolor='e7e7e7'>&nbsp;<%=maintain_user_name%></td>			    
  </tr>
  <tr> 
    <td width='15%' bgcolor='#D8EFEE' align='left'>日期</td>
	<td width='85%' bgcolor='e7e7e7'>&nbsp;<%=maintain_update_date%></td>			    
  </tr>
  <%}%>
  <tr bgcolor="#9AD3D0"> 
    <td colspan='2' class="sbody" align="center"><font color='#000000'>登入者資訊</font></td>
  </tr>
  <tr> 
    <td width='15%' bgcolor='#D8EFEE' align='left'>姓名</td>
	<td width='85%' bgcolor='e7e7e7'>&nbsp;<%=lguser_name%></td>			    
  </tr>
  <tr> 
    <td width='15%' bgcolor='#D8EFEE' align='left'>電話</td>
	<td width='85%' bgcolor='e7e7e7'>&nbsp;<%=lguser_telno%></td>			    
  </tr>
  <tr> 
    <td width='15%' bgcolor='#D8EFEE' align='left'>E-MAIL</td>
	<td width='85%' bgcolor='e7e7e7'>&nbsp;<%=lguser_email%></td>			    
  </tr>
</table> 

                    