<%
//105.9.21 create by 2968
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="java.util.*" %>
<%@ page import="java.util.LinkedList" %>

<%
	String sAcc_Tr_Name = ( request.getAttribute("sAcc_Tr_Name")==null ) ? "" : (String)request.getAttribute("sAcc_Tr_Name");
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");
	List dbData = (List)request.getAttribute("dbData");		   
	String[][] Report_List = new String[99][7];
    if(dbData != null && dbData.size() != 0){
    	Report_List = new String[dbData.size()][7];
       for(int i=0;i<dbData.size();i++){
           Report_List[i][0]=(String)((Map)dbData.get(i)).get("acc_tr_type");
	       Report_List[i][1]=(String)((Map)dbData.get(i)).get("acc_tr_name");
	       Report_List[i][2]=(String)((Map)dbData.get(i)).get("acc_name01");
	       Report_List[i][3]=(String)((Map)dbData.get(i)).get("acc_name02");
	       Report_List[i][4]=(String)((Map)dbData.get(i)).get("applytype_name");
	       Report_List[i][5]=(String)((Map)dbData.get(i)).get("applydate");
       }
	}	   		       
    String list1Color="list1Color_sbody";
    String list2Color="list2Color_sbody";
    String listColor="list1Color_sbody"; 
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/TM005W.js"></script>
<script language="javascript" event="onresize" for="window"></script>

<html>
<head>
<title>協助措施申報項目維護作業</title>
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
<form name=form method=post>
<input type="hidden" name="act" value="qList">  
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
                      <td width="130"><img src="images/banner_bg1.gif" width="130" height="17"></td>
                      <td width="340"><font color='#000000' size=4><b> 
                        <center>協助措施申報項目維護作業 </center>
                        </b></font> </td>
                      <td width="130"><img src="images/banner_bg1.gif" width="130" height="17"></td>
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
                          <tr align=left bgcolor="#9AD3D0" class="sbody"  >
                            <td width=130 >協助措施名稱</td>
                            <td class="<%=list2Color%>" >
                            <input type='text' name='acc_Tr_Name' size=35 value='<%=sAcc_Tr_Name%>'>
                            	&nbsp;&nbsp;
                            	<a href="javascript:doSubmit(this.document.forms[0],'qList');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_queryb.gif',1)"><img src="images/bt_query.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>
                            	<a href="javascript:doSubmit(this.document.forms[0],'new');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image102','','images/bt_addb.gif',1)"><img src="images/bt_add.gif" name="Image102" width="66" height="25" border="0" id="Image102"></a>
							</td>
                          </tr>
                        </Table></td>
                    </tr>
				    <tr><td>&nbsp;</td></tr>
				    
                    <tr>
                      <td><table width=600 border=1 align='left' cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
                          <tr  bgcolor="#9AD3D0" class="sbody" align='center'>
                            <td width=30>序號</td>
                            <td width=170>協助措施名稱</td>
                            <td width=120>舊貸展延需求</td>
                            <td width=120>新貸需求</td>
                            <td width=60>申報頻率</td>
                            <td width=100>申報日期</td>
                          </tr>
                        </Table>
                      </td>
                    </tr>
                    <tr> 
                      <td>
                      <table width='600' border=1 align='left' cellpadding="1" cellspacing="1" bordercolor="#3A9D99">                  
                          <%
                          if(dbData != null && dbData.size() != 0){
	                           for(int i = 0; i < Report_List.length; i++){
	                        	   listColor = (i % 2 == 0)?list2Color:list1Color;
	                           %>
							   <tr class="<%=listColor%>" >
							   		<td width=30><%=i+1 %></td>  
									<td width=170><a href="javascript:goEdit(this.document.forms[0],'<%=Report_List[i][0]%>','<%=Report_List[i][1]%>');"><%=Report_List[i][1]%></a></td>
									<td width=120><%=Report_List[i][2]%></td>
									<td width=120><%=Report_List[i][3]%></td>
									<td width=60><%=Report_List[i][4]%></td>
									<td width=100><%=Report_List[i][5]%></td>	   	
	    						</tr>    
	    					  <%}%>	 
    					   <%}else{%>
    					   	<tr class="<%=list1Color%>">                   			   
                  			   <td colspan=10 align=center >無資料可供查詢</td><tr>
                  			<tr> 
    					   <%} %>
            			</table>
            		 </td>
        			</tr>  
        		     
		      </table></td>
		  </tr>
</table>
</form>
</body>
</html>
