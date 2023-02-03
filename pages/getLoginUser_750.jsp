<%
//100.05.10 fix 取得機構名稱區分99/100年度 by 2295
//102.10.07 fix 作業人員姓名部份遮蔽 by 2968
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@include file="common.jsp"%>

<%
	String s_year = Utility.getYear();
	String wlx01_m_year = (Integer.parseInt(s_year) < 100)?"99":"100"; 
	List paramList = new ArrayList(); 
	
	String lguser_name = (session.getAttribute("muser_name") == null)?"":Utility.maskChar((String)session.getAttribute("muser_name"),2,((String)session.getAttribute("muser_name")).length()-2,"＊"); 
	//登入者的tbank_no
	String tbank_no = (session.getAttribute("tbank_no") == null)?"":(String)session.getAttribute("tbank_no"); 
	//所點選的nowtbank_no
	//若有點選的tbank_no,則顯示所點選的tbank_no
	tbank_no = ( session.getAttribute("nowtbank_no")==null ) ? tbank_no : (String)session.getAttribute("nowtbank_no");		
	String bank_name = "";
	String sqlCmd = "select bank_name from BN01  where bank_no=? and m_year=?";
	
	paramList.add(tbank_no);
	paramList.add(wlx01_m_year);
	
	List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");
	if(dbData != null && dbData.size() != 0){
	   bank_name = (String)((DataObject)dbData.get(0)).getValue("bank_name");
	} 	
%>
 <td><table width=750 border=1 align='center' cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
     <tr bgcolor="#9AD3D0" class="sbody"> 
     <td width=112 bgcolor=#9AD3D0><font face=細明體 color=#000000>機構名稱</font></td>
     <td width=175><font face=細明體 color=#000000><%=bank_name%></font></td>
     <td width=113><font face=細明體 color=#000000>作業人員</font></td>
     <td width=175><font face=細明體 color=#000000><%=lguser_name%></font></td>
     </tr>
     </table>
 </td>