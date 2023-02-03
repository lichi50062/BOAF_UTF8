<%
//94.02.15 fix 若actMsg屬於檔案上傳成功時,則回到上傳時的那一個頁面 by 2295
//94.10.31 fix 若actMsg在農漁會新增申報完成後，則顯示回到查詢頁 by 4180
//94.11.15 add 回查詢頁所要link的jsp名稱 by 2295
//95.01.17 fix 查詢頁所要link加bank_type by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@include file="common.jsp"%>
<%
	String Report_no = ( request.getParameter("Report_no")==null ) ? "" : (String)request.getParameter("Report_no");
	String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? "" : (String)request.getParameter("S_YEAR");
	String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? "" : (String)request.getParameter("S_MONTH");
	//94.11.15 add 回查詢頁所要link的jsp名稱 
	String goPages = ( request.getParameter("goPages")==null ) ? "" : (String)request.getParameter("goPages");
	String bank_code = ( request.getParameter("bank_code")==null ) ? "" : (String)request.getParameter("bank_code");
	String bank_type = ( request.getParameter("bank_type")==null ) ? "" : (String)request.getParameter("bank_type");
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");
	String acc_Tr_Type = ( request.getParameter("acc_Tr_Type")==null ) ? "" : (String)request.getParameter("acc_Tr_Type");//TM006W
    //取得是否為農漁會新增資料程式
    String FX = ( request.getParameter("FX")==null ) ? "" : (String)request.getParameter("FX");
	String actMsg = ( request.getAttribute("actMsg")==null ) ? "" : (String)request.getAttribute("actMsg");
	String alertMsg = ( request.getAttribute("alertMsg")==null ) ? "" : (String)request.getAttribute("alertMsg");
	String webURL_Y = ( request.getAttribute("webURL_Y")==null ) ? "" : (String)request.getAttribute("webURL_Y");
	String webURL_N = ( request.getAttribute("webURL_N")==null ) ? "" : (String)request.getAttribute("webURL_N");
	System.out.println("FX="+FX);
	System.out.println("actMsg="+actMsg);
	System.out.println("alertMsg="+alertMsg);
	System.out.println("webURL_Y="+webURL_Y);
	System.out.println("webURL_N="+webURL_N);
%>
<script language="javascript" src="js/Common.js"></script>
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
<form method=post name="Msgform">
<%if(!webURL_N.equals("")){%>
<BODY onLoad="javascript:ConfirmMsg(this.document.Msgform,'<%=alertMsg%>','<%=webURL_Y%>','<%=webURL_N%>')">
<%}else{%>
<BODY onLoad="javascript:AlertMsg(this.document.Msgform,'<%=alertMsg%>','<%=webURL_Y%>')">
<%}%>
<%= actMsg %>
<% if(alertMsg.equals("")){ %>
<center>
<tr class="sbody">
<%if(actMsg.indexOf("檔案上傳成功") != -1){%>
<div><a href='/pages/WMFileUpload.jsp?act=new&Report_no=<%=Report_no%>&S_YEAR=<%=S_YEAR%>&S_MONTH=<%=S_MONTH%>&test=nothing' onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image8','','images/bt_backb.gif',1)"><img src="images/bt_back.gif" name="Image8" width="80" height="25" border="0"></a></div>
<%//若農漁會新增資料或聯貸案件申報作業，則可回查詢頁
}else if(FX.indexOf("FX") != -1 || FX.equals("WMFileEdit_A11")){	
%>
<div><a href='/pages/<%=FX%>.jsp?act=List' onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image8','','images/bt_05b.gif',1)"><img src="images/bt_05.gif" name="Image8" width="80" height="25" border="0"></a></div>
<%}else if(!goPages.equals("")){ //94.11.15 add 回查詢頁的link%>
<div><a href='/pages/<%=goPages%>?act=<%=act%>&Report_no=<%=Report_no%>&bank_code=<%=bank_code%>&S_YEAR=<%=S_YEAR%>&S_MONTH=<%=S_MONTH%>&bank_type=<%=bank_type%>&acc_Tr_Type=<%=acc_Tr_Type%>&test=nothing' onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image8','','images/bt_05b.gif',1)"><img src="images/bt_05.gif" name="Image8" width="80" height="25" border="0"></a></div>
<%}else{%>
<div><a href="javascript:history.back();" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image8','','images/bt_backb.gif',1)"><img src="images/bt_back.gif" name="Image8" width="80" height="25" border="0"></a></div>
<%}%>
<tr>
<%}%>
</form>
</BODY>

</HTML>




