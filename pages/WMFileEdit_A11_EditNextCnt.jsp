<%
//新增下一筆參貸項目 by 2968
//104.05.12 fix 補舊案件資料時,新增下一筆參貸項目時,該案件編號會是空值 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%
	String s_year = ( request.getParameter("s_year")==null ) ? "" : (String)request.getParameter("s_year");
	String s_month = ( request.getParameter("s_month")==null ) ? "" : (String)request.getParameter("s_month");
	String bank_no = ( request.getParameter("bank_no")==null ) ? "" : (String)request.getParameter("bank_no");
	String case_no = ( request.getParameter("case_no")==null ) ? "" : (String)request.getParameter("case_no");
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");
	String actMsg = ( request.getAttribute("actMsg")==null ) ? "" : (String)request.getAttribute("actMsg");
	System.out.println("WMFileEdit_A11_EditNextCnt.actMsg="+actMsg);
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/FileEdit_A11.js"> </script>
<script language="javascript" event="onresize" for="window"></script>
<HTML>
<HEAD>
<link href="css/b51.css" rel="stylesheet" type="text/css">
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
//-->
</script>
</HEAD>

<BODY>
<form method=post name="Msgform">
<center>
<tr class="sbody">
相關資料寫入資料庫成功
<input type='hidden' name='hyear' value="<%=s_year%>">
<input type='hidden' name='hmonth' value="<%=s_month%>">
<input type="button" value="新增下一筆參貸項目" onclick="javascript:continueEditCnt(this.document.forms[0],'<%=s_year%>','<%=s_month%>','<%=bank_no%>','<%=case_no%>');">
<tr>
</form>
</BODY>
</HTML>