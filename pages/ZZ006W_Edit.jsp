<%
//102.06.07 created by 2968
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%  
	String title="各項報表維護";
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");				
	String program_id =( request.getParameter("program_id")==null ) ? "" : (String)request.getParameter("program_id");
	String rpt_no =( request.getParameter("rpt_no")==null ) ? "" : (String)request.getParameter("rpt_no");	
	List mtnInfo = (List)request.getAttribute("maintainInfo");	
	if(mtnInfo == null){
		System.out.println("mtnInfo == null");
	}else{
		System.out.println("mtnInfo.size()="+mtnInfo.size());
	}
	
	//取得ZZ006W的權限
	Properties permission = ( session.getAttribute("ZZ006W")==null ) ? new Properties() : (Properties)session.getAttribute("ZZ006W"); 
	if(permission == null){
       System.out.println("ZZ006W_Edit.permission == null");
    }else{
       System.out.println("ZZ006W_Edit.permission.size ="+permission.size());
               
    }
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/ZZ006W.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<title><%=title%></title>
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
<form method=post action='#'>
<input type="hidden" name="act" value="">  
<table width="800" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  		<tr> 
   		 <td><img src="images/space_1.gif" width="12" height="12"></td>
  		</tr>
        <tr> 
          <td bgcolor="#FFFFFF">
			<table width="800" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="800" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="250"><img src="images/banner_bg1.gif" width="250" height="17"></td>
                      <td width="300"><font color='#000000' size=4><b> 
                        <center>
                          <%=title%> 
                        </center>
                        </b></font> </td>
                      <td width="250"><img src="images/banner_bg1.gif" width="250" height="17"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><img src="images/space_1.gif" width="12" height="12"></td>
              </tr>
              <tr> 
                <td><table width="800" border="0" align="center" cellpadding="0" cellspacing="0">
               
                    <tr> 
                      <div align="right"><jsp:include page="getLoginUser.jsp?width=800" flush="true" /></div> 
                    </tr>          
                    <%
                      String nameColor="nameColor_sbody";
                      String textColor="textColor_sbody";
                      String bordercolor="#3A9D99";
                    %>         
                    <tr> 
                      <td><table width=800 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="<%=bordercolor%>">
                      	  <tr class="sbody">
	                      	  <td width='20%' class="<%=nameColor%>" align=center>功能代碼</td>
							  <td width='80%' colspan=2 class="<%=textColor%>">
							  	<input type="hidden" name="rpt_no" value="<%if(mtnInfo != null && mtnInfo.size() != 0 && (((DataObject)mtnInfo.get(0)).getValue("rpt_no")).toString() != null ) out.print((((DataObject)mtnInfo.get(0)).getValue("rpt_no")).toString());%>"> 
							  	<input type="hidden" name="rpt_idOri" value="<%if(mtnInfo != null && mtnInfo.size() != 0 && ((DataObject)mtnInfo.get(0)).getValue("rpt_id") != null ) out.print((String)((DataObject)mtnInfo.get(0)).getValue("rpt_id"));%>"> 
							  	<input type="text" name="rpt_id" size="30" maxlength="30" value="<%if(mtnInfo != null && mtnInfo.size() != 0 && ((DataObject)mtnInfo.get(0)).getValue("rpt_id") != null ) out.print((String)((DataObject)mtnInfo.get(0)).getValue("rpt_id"));%>">
	                            </td>                           
                          </tr>
					      <tr class="sbody">
							  <td width='20%' class="<%=nameColor%>" align=center>報表抬頭名稱</td>						  
							  <td width='80%' colspan=2 class="<%=textColor%>">
	  		                    <input type="text" name="rpt_name" size="80" maxlength="100" value="<%if(mtnInfo != null && mtnInfo.size() != 0 && ((DataObject)mtnInfo.get(0)).getValue("rpt_name") != null ) out.print((String)((DataObject)mtnInfo.get(0)).getValue("rpt_name"));%>" size='50' maxlength='100'>  		                    
	                          </td>
                          </tr>   		      
					      <tr class="sbody">
							  <td width='20%' class="<%=nameColor%>" align=center>報表功能路徑</td>
							  <td width='80%' colspan=2 class="<%=textColor%>"> 
	  		                    <input type="text" name="rpt_path" size="80" maxlength="200" value="<%if(mtnInfo != null && mtnInfo.size() != 0 && ((DataObject)mtnInfo.get(0)).getValue("rpt_path") != null ) out.print((String)((DataObject)mtnInfo.get(0)).getValue("rpt_path"));%>" size='50' maxlength='100'>  		                    
	                          </td>
                          </tr>
					      <tr class="sbody">
							  <td width='20%' class="<%=nameColor%>" align=center>報表內欄位名稱</td>
							  <td width='80%' colspan=2 class="<%=textColor%>">
							  <%for(int i=0;i<mtnInfo.size();i++){ 
							  	out.println((i+1)+".");
							  	out.println((String)((DataObject)mtnInfo.get(i)).getValue("col_name"));
							  	out.println("<br>");
	                          } %>
	                          </td>
                          </tr>
                          
                        </Table>
                        <table width="800" border="0" cellpadding="1" cellspacing="1" class="sbody" align="center">
						  <tr>
						    <td colspan='2' align='center'>
						    
						 <%if(permission != null && permission.get("U") != null && permission.get("U").equals("Y")){ %>
						  <a href="javascript:doSubmit(this.document.forms[0],'Update');"onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image102','','images/bt_updateb.gif',1)"><img src="images/bt_update.gif" name="Image102" width="66" height="25" border="0" id="Image102"></a>
						  <%} %>
						 <%if(permission != null && permission.get("D") != null && permission.get("D").equals("Y")){ %>
						  <a href="javascript:doSubmit(this.document.forms[0],'Delete');"onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_deleteb.gif',1)"><img src="images/bt_delete.gif" name="Image103" width="66" height="25" border="0" id="Image103"></a>
						 <%} %>
						<%if(permission != null && permission.get("Q") != null && permission.get("Q").equals("Y")){ %>
						   <a href="javascript:location.replace('ZZ006W.jsp?act=List');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_05b.gif',1)"><img src="images/bt_05.gif" name="Image105" width="80" height="25" border="0" id="Image105"></a>
						<%} %>
						  </td>
						  </tr>
						  <tr> 
						    <td colspan="2">
						      <font color='#990000'>
						        <img src="images/arrow_1.gif" width="28" height="23" align="absmiddle">
						        <font color="#007D7D" size="3">使用說明  : </font></font></td>
						  </tr>
						  <tr> 
						    <td width="3%">&nbsp;</td>
						    <td width="82%"> 
						      <ul>                                                
						        
						            <li>本網頁提供報表資料維護功能。</li>
						            <li>按<font color="#666666">【修改】</font>即將修改資料寫入資料庫。</li> 
								    <li>按<font color="#666666">【刪除】</font>刪除報表資料。</li>
						         
						
						        <li>按<font color="#666666">【回查詢頁】</font>則回至查詢畫面。</li>
						      </ul>
						    </td>
						 </tr>
						</table>
                        </td></tr>                 
			      </table></td>
			  </tr>

		</table>
		</td>
		</tr>
</table>
</form>
</body>

<script language="JavaScript" >
<!--
-->
</script>
</html>
