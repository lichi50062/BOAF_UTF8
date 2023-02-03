<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%
String actMsg = ( request.getAttribute("actMsg")==null ) ? "" : (String)request.getAttribute("actMsg");
System.out.println("JSP.actMsg="+actMsg);
%>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<title>信保基金靜態網頁檔案上傳</title>
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

function message() {
<% if(!actMsg.equals("")) {    
%>
   alert('<%=actMsg%>');
<%}%>
return ;
}

function doSubmit(form) {	
	if (form.UpFileName.value == '') {
		alert('上傳檔案位置為空值');
		form.UpFileName.focus();
		return;
	}	
	form.action="FileUpload.jsp?UserID="+form.UserID.value+"&UserPWD="+form.UserPWD.value;
	//alert(form.action);
	form.submit();
}

//-->
</script>
</head>

<body  leftmargin="0" topmargin="0" >
<form method=post ENCTYPE="multipart/form-data" action='FileUpload.jsp' onload='javascript:<%if(!actMsg.equals("")){%>alert('<%=actMsg%>');<%}%>'>
<table width="640" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#297A76">
  <tr>
    <td bordercolor="#FFFFFF"><table width="640" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td bgcolor="#FFFFFF">&nbsp;</td>
        </tr>
        <tr> 
          <td bgcolor="#FFFFFF">
<table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="150"><img src="image/banner_bg1.gif" width="150" height="17"></td>
                      <td width="300"><font color='#000000' size=4><b> 
                        <center>
                          信保基金靜態網頁檔案上傳 
                        </center>
                        </b></font> </td>
                      <td width="150"><img src="image/banner_bg1.gif" width="150" height="17"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><img src="image/space_1.gif" width="12" height="12"></td>
              </tr>
              <tr> 
                <td><table width="600" border="0" align="center" cellpadding="0" cellspacing="0">               
                    <tr> 
                      <td><table width=600 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">                          	  					
                        <tr class="sbody">
						<td bgcolor='#D8EFEE' align='left'>使用者帳號</td>
						<td colspan=2 bgcolor='e7e7e7'>
        				<input type=text size=20 name=UserID >        				     				
        			    </td>
        			    </tr>	
        			    <tr class="sbody">
						<td bgcolor='#D8EFEE' align='left'>使用者密碼</td>
						<td colspan=2 bgcolor='e7e7e7'>
        				<input type=password size=10 name=UserPWD >        				     				
        			    </td>
        			    </tr>	
        				<tr class="sbody">
						<td bgcolor='#D8EFEE' align='left'>上傳檔案位置</td>
						<td colspan=2 bgcolor='e7e7e7'>
        				<input type=file size=40 name=UpFileName >
        				<input type=button value="上傳" onclick="doSubmit(this.document.forms[0]);">        				
        			    </td>
        			    </tr>		
                        </Table></td>
                    </tr>                 
                    <tr> 
              </tr>   
              <tr> 
                <td>&nbsp;</td>
              </tr> 
              <tr>
              <td><font color='red'><%=actMsg%></font></td>
              </tr>
      </table></td>
  </tr>
  <tr> 
    <td><table width="600" border="0" cellpadding="1" cellspacing="1" class="sbody">
        <tr> 
          <td colspan="2"><font color='#990000'><img src="image/arrow_1.gif" width="28" height="23" align="absmiddle"><font color="#007D7D" size="3">使用說明 
            : </font></font></td>
        </tr>
        <tr> 
          <td width="16">&nbsp;</td>
          <td width="577"> <ul>
              <li>本網頁提供上傳信保基金靜態網頁。</li>                                               
            </ul></td>
        </tr>
      </table>
    </td>
  </tr>             
</table>
</form>
</body>
</html>

