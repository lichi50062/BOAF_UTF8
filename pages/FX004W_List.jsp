<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.LinkedList" %>
<%
	
	List BA01 = (List)request.getAttribute("BA01");
	String list_type = ( request.getParameter("list_type")==null ) ? "" : (String)request.getParameter("list_type");		
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/FX004W.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<%/*if(list_type.equals("1")){%>
<!--title>地方主管機關基本資料維護</title>
<%}else if(list_type.equals("2")){%>
<title>共用中心基本資料維護</title>
<%}else if(list_type.equals("3")){%>
<title>農業行庫基本資料維護</title-->
<%}*/%>
<title>機構基本資料維護</title>
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
</head>

<body marginwidth="0" marginheight="0" leftmargin="0" topmargin="0" leftmargin="0">
<form method=post>
<input type="hidden" name="act" value="List">  
<table width="640" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  		<tr> 
   		 <td><img src="images/space_1.gif" width="12" height="12"></td>
  		</tr>
        <tr> 
          <td bgcolor="#FFFFFF">
		  <table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="220"><img src="images/banner_bg1.gif" width="220" height="17"></td>
                      <td width="160"><font color='#000000' size=4><b> 
                        <center>
                        機構基本資料維護                 
                        </center>
                        </b></font> </td>
                      <td width="220"><img src="images/banner_bg1.gif" width="220" height="17"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><img src="images/space_1.gif" width="12" height="12"></td>
              </tr>
              <tr> 
                <td><table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
               
                    <tr> 
                      <div align="right"><jsp:include page="getLoginUser.jsp" flush="true" /></div> 
                    </tr>                      
                    <tr> 
                      <td><table width=600 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
                          <tr class="sbody" bgcolor="#D8EFEE">
                          <td width='30%' align=center>機構代號</td>
      					  <td width='70%' align=center>機構名稱</td>
      					  
                          </tr>    
                          <%if(BA01.size() == 0){%>
                          <tr class="sbody" bgcolor="#e7e7e7">
                          <td colspan="2" align=center>查無資料</td>
                          </tr>
                          <%}else{
                             String bgcolor="#D3EBE0";
                             for(int i=0;i<BA01.size();i++){
                                 bgcolor = (i % 2 == 0)?"#e7e7e7":"#D3EBE0";
                          %>
                          <tr class="sbody" bgcolor='<%=bgcolor%>'>
                          <td class="sbody" width='30%' align=center>
                             <a href='FX004W.jsp?act=Edit&bank_no=<%=(String)((DataObject)BA01.get(i)).getValue("bank_no")%>'>
                             <%=(String)((DataObject)BA01.get(i)).getValue("bank_no")%>
                             </a>
                          </td>
      					  <td class="sbody"  width='70%' align=center>
      					     <a href='FX004W.jsp?act=Edit&bank_no=<%=(String)((DataObject)BA01.get(i)).getValue("bank_no")%>'>
      					     <%=(String)((DataObject)BA01.get(i)).getValue("bank_name")%>
      					     </a>
      					  </td>      					  						  
                          </tr>  
                          
                          <%   
                             }//end for
                            }//end of if 
                          %>   
                          
                         </table>
                      </td>
                    </tr>  
                    
                    
                    <td><div align="right"><jsp:include page="getMaintainUser.jsp" flush="true" /></div></td>                                              
      </table></td>
  </tr>
  <tr> 
                <td><table width="600" border="0" cellpadding="1" cellspacing="1" class="sbody">
                    <tr> 
                      <td colspan="2"><font color='#990000'><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle"><font color="#007D7D" size="3">使用說明 
                        : </font></font></td>
                    </tr>
                    <tr> 
                      <td width="16">&nbsp;</td>
                      <td width="577"> <ul>                          
                          <li>點選所列之[機構代號或機構名稱]可變更該機構通訊錄。</li>                          
                        </ul></td>
                    </tr>
                  </table></td>
              </tr>
              <!--tr> 
                <td><div align="center"><img src="images/line_1.gif" width="600" height="12"></div></td>
              </tr-->
</table>
</form>
</body>
</html>
