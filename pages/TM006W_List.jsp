<%
//105.10.03 create by2968
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.Date" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Calendar" %>
<%
    List dbData= (request.getAttribute("dbData")==null)?null:(List)request.getAttribute("dbData");
    String bank_no = ( request.getParameter("bank_no")==null ) ? "" : (String)request.getParameter("bank_no");
    Properties permission = ( session.getAttribute("TM006W")==null ) ? new Properties() : (Properties)session.getAttribute("TM006W");
    if (permission == null) {
        System.out.println("TM006W_List.permission == null");
    }else {
        System.out.println("TM006W_List.permission.size ="+permission.size());
    }
   
%>
<script language="javascript" src="js/TM006W.js"></script>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>

<head>
<link href="css/b51.css" rel="stylesheet" type="text/css">
<title>適用協助措施之經辦機構申報作業</title>
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
</head>
<body topmargin="0" leftmargin="15">
<form name="date" method="post">
<table width="770" border="0" align="left" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" >
        <tr>
         	<td width="770" height="16"><img src="images/space_1.gif" width="12" height="12"></td>
        </tr>       
        <tr>        
          <td bgcolor="#FFFFFF" width="770" height="310">        
          <table width="770" border="0" align="center" cellpadding="0" cellspacing="0" >
              <tr>
                <td width="770">
                  <table width="770" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="190"><img src="images/banner_bg1.gif" width="190" height="17" align="left"></td>
                      <td width="365" align="center"><b><center><font color="#000000" size="4">適用協助措施之經辦機構申報作業</font></center></b> </td>
                      <td width="190" align="right"><img src="images/banner_bg1.gif" width="190" height="17"></td>
                    </tr>
                  </table>
                <table width="770" border="0" align="center" cellpadding="0" cellspacing="0" >  
                <tr><div align="right"><jsp:include page="getLoginUser.jsp?width=770" flush="true" /></div></tr>                  
                    <tr class="sbody">
                      <td  class="sbody"  width="770">
                      <table class="sbody" width="770" border=1 cellpadding="0" cellspacing="1" bordercolor="#3A9D99" height="20" >
                          
                          <tr>
                            <td width="762" bgcolor="#D2F0FF" align="center" >協助措施名稱</td>
                          </tr>
                          <%if(dbData.size()==0){%>
                          <tr>
                            <td width="762" align="center" colspan="" bgcolor="#FFFFE6">無適用的協助措施</td>
                          </tr>
                          <%}else{                          	  
                              for(int i=0;i<dbData.size();i++){
                          %>
                           <tr class="sbody" bgcolor="<%out.print((i%2==0)?"#FFFFE6":"#F2F2F2");%>">
                          	   <td class="sbody" width="762" align="center">                            
	                           <a href='TM006W.jsp?act=ApplyList&acc_Tr_Type=<%=String.valueOf(((DataObject)dbData.get(i)).getValue("acc_tr_type")) %>'>
	                           <%=String.valueOf(((DataObject)dbData.get(i)).getValue("acc_tr_name"))%></a></td>             
                           </tr>                 
                          <%  }//end for
                            }//end else %>
                         </table>                     
                      </td>
                    </tr>


      			</table>
      			</td>
  			  </tr>
          </table>
          </td>
       </tr>   
</table>
</form>
</html>
