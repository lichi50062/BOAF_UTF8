<%
/*
 *102.1.8 create  by 2968
 */
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

	Properties permission = ( session.getAttribute("WMFileEdit")==null ) ? new Properties() : (Properties)session.getAttribute("WMFileEdit"); 
	if(permission == null){
	   System.out.println("WMFileEdit_A11.permission == null");
	}else{
	   System.out.println("WMFileEdit_A11.permission.size ="+permission.size());
	           
	}
	//---------------------------------------取得參數---------------------------------------------
	List A11_list1 = (List)request.getAttribute("A11_S");   	//內含每月的資料
	List A11_list2 = (List)request.getAttribute("A11_S_Sum");   //內含年月以及筆數
	String bank_no =  ( request.getParameter("bank_no")==null ) ? " " : (String)request.getParameter("bank_no");
	String case_no =  ( request.getParameter("case_no")==null ) ? " " : (String)request.getParameter("case_no");
	String s_year = (  request.getParameter("s_year")==null ) ? " " : (String) request.getParameter("s_year");
	String s_month = (  request.getParameter("s_month")==null ) ? " " : (String) request.getParameter("s_month");
	String A11_Lock = (String)request.getAttribute("A11_Lock");    //檢查有無鎖定
	String isLastMonthData = (String)request.getAttribute("isLastMonthData"); //上個月是否有申報資料
	String isHaveApplyData = (String)request.getAttribute("isHaveApplyData"); //本月是否已有申報資料
	String isHaveNoApplyData = (String)request.getAttribute("isHaveNoApplyData"); //本月是否已有申報過無資料
	String isHaveNoApplyLoanData = (String)request.getAttribute("isHaveNoApplyLoanData"); //本月是否已有申報過無資料但有實際授信餘額
	DataObject Bean = new DataObject(); 
	int width = 500;//1300 pixel

%>
<!-- ------------------------------------------------------------------------------------------ -->
<script language="JavaScript" type="text/JavaScript">
<!-- 
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
-->
</script>
<!-- ------------------------------------------------------------------------------------------- -->
<html>
<link href="css/b51.css" rel="stylesheet" type="text/css">
<script language="javascript" src="js/Common.js"> </script>
<script language="javascript" src="js/FileEdit_A11.js"> </script>
<script language="javascript" event="onresize" for="window"> </script>
<head>
<title>聯合貸款案件資料表申報情形</title>
</head>
<body marginwidth="0" marginheight="0" leftmargin="0" topmargin="0" leftmargin="0">
<form method=post>
  <table border="0" cellspacing="0" width="<%=width %>" cellpadding="0" bgcolor="#FFFFFF">
    <tr>
      <td height="18" width="<%=width %>"></td>
    </tr>
    <tr>
	<table border="1" cellspacing="1" width="<%=width %>" bordercolor="#3A9D99">
 		<tr class="sbody" bgcolor="#D8EFEE">
 		<td>
<!-- -----------------------------hide type---------------------------- -->
			<input type='hidden' name='s_year' value="<%=s_year%>" >
			<input type='hidden' name='s_month' value="<%=s_month%>">
<!-- -----------------------------hide type---------------------------- -->
 			<br>
 			本月有新增聯貸案件申報資料？
 			<br>&nbsp&nbsp&nbsp&nbsp&nbsp
 			<input type="Radio" name="radio1" value="Y" onclick="javascript:document.getElementById('BlockY').style.display='inline';" checked>是
 			<input type="Radio" name="radio1" value="N" onclick="javascript:document.getElementById('BlockY').style.display='none' ;document.getElementById('case_cnt').value=''; ">否
 			
 			<div id="BlockY" style="display:inline" >
 				<br><br>
	 			新增聯貸案件有無甲、乙、丙等授信項目?
	 			<br><br>&nbsp&nbsp&nbsp&nbsp
	 			<input type="Radio" name="radio2" value="Y" checked onclick="document.getElementById('case_cnt').value='';" >是&nbsp
	 			請輸入項數：
	 			<input type='text' name='case_cnt' size='2' maxlength='2' >
	 			<br><br>&nbsp&nbsp&nbsp&nbsp
	 			<input type="Radio" name="radio2" value="N" onclick="javascript:document.getElementById('case_cnt').value='';">否
	 			<br>
 			</div>
 			<br>
 			<center><input type="button" name="button" value="確定" onclick="javascript:doSubmit(this.document.forms[0],'nextPg','<%=bank_no%>','<%=A11_Lock%>','<%=isLastMonthData%>','<%=isHaveApplyData%>','<%=isHaveNoApplyData%>','<%=isHaveNoApplyLoanData%>');"></center>
 			<br>
 		</td>
        </tr>
    </table>
	</tr>

  </table>
</form>
</body>

</html>

<!-- --------------------------------------------------- -->
<script language='JavaScript'>

</script>
<!-- -------------------------------------------------	-->
