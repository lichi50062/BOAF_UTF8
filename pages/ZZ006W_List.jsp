<%
// 102.06.07 created by 2968
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
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");		
	String sn_docno0 = ( request.getParameter("sn_docno0")==null ) ? "" : (String)request.getParameter("sn_docno0");
	String sn_docno2 = ( request.getParameter("sn_docno2")==null ) ? "" : (String)request.getParameter("sn_docno2");
	String sn_docno1 = ( request.getParameter("sn_docno1")==null ) ? "" : (String)request.getParameter("sn_docno1");
	String pgId = ( request.getParameter("pgId")==null ) ? "" : (String)request.getParameter("pgId");
	List qList = (List)request.getAttribute("qList");		
	if(qList == null){
	   System.out.println("qList == null");
	}else{
	   System.out.println("qList.size()="+qList.size());
	}
	//取得ZZ006W的權限
	Properties permission = ( session.getAttribute("ZZ006W")==null ) ? new Properties() : (Properties)session.getAttribute("ZZ006W"); 
	if(permission == null){
       System.out.println("ZZ006W_List.permission == null");
    }else{
       System.out.println("ZZ006W_List.permission.size ="+permission.size());               
    }
    
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/ZZ006W.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<title>各項報表查詢 </title>
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
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  		<tr> 
   		 <td><img src="images/space_1.gif" width="12" height="12"></td>
  		</tr>

        <tr> 
          <td bgcolor="#FFFFFF">
		  <table width="800" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="250"><img src="images/banner_bg1.gif" width="250" height="17"></td>
                      <td width="300"><font color='#000000' size=4><b> 
                        <center>各項報表查詢 </center>
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
                      <td>
                      <table width="800" border=1 align=center cellpadding="1" cellspacing="1" bordercolor="<%=bordercolor%>">
                      		<tr class="sbody">
							<td width="93" class="<%=nameColor%>" height="1">
							<p align="left">功能代碼</td>
							<td width="586" class="<%=textColor%>" height="1">
							   <p align="left">&nbsp;<input type="text" name="pgId" size="50" value='<%=pgId %>'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							   <%if(permission != null && permission.get("Q") != null && permission.get("Q").equals("Y")){ %>                   	        	                                   		     
	                            <a href="javascript:doSubmit(this.document.forms[0],'Qry');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_queryb.gif',1)"><img src="images/bt_query.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>
	                            <%}%>
							  </td>
							</tr>
                            <tr class="sbody">
							<td width="93" class="<%=nameColor%>" height="1">
							<p align="left">報表名稱</td>
							<td width="586" class="<%=textColor%>" height="1">
							   <p align="left">&nbsp;<input type="text" name="sn_docno0" size="50" value='<%=sn_docno0 %>'>
							  </td>
							</tr>
							<tr class="sbody">
							<td width="93" class="<%=nameColor%>" height="1">
							<p align="left">報表欄位名稱</td>
							<td width="586" class="<%=textColor%>" height="1">
							   <p align="left">
							   <input type="text" name="sn_docno2" size="51" value='<%=sn_docno2 %>'></td>
							</tr>
							<tr class="sbody">
								<td width="93" class="<%=nameColor%>" height="1">
								<p align="left">科目代碼</td>
								<td width="586" class="<%=textColor%>" height="1">
									<p align="left">
									<input type="text" name="sn_docno1" size="20" value='<%=sn_docno1 %>'></td>
							</tr>
                          </table></td></tr> </table>     
                        
                      
                      <tr> 
                      <td><table width="800" border=1 align=center cellpadding="1" cellspacing="1" bordercolor="<%=bordercolor%>">                      
                      <%
                      String tmpbank_type="";                      
                      String tmptbank_no="";                      
                      int i = 0;      
                      String listTitle="listTitleColor_sbody"; 
                      String list1Color="list1Color_sbody";
                      String list2Color="list2Color_sbody";
                      String listColor="list1Color_sbody";      
                      if(qList != null){ %>
                      	<tr class="sbody" bgcolor="#BFDFAE" >
                      	    <%if(!"".equals(sn_docno2)||!"".equals(sn_docno1)){ %>
						        <td class="<%=listTitle%>"  width="8%" >功能代碼</td>
						        <td class="<%=listTitle%>"  width="24%">報表名稱</td>
						        <td class="<%=listTitle%>"  width="20%">報表欄位</td>
						        <td class="<%=listTitle%>"  width="69%">報表功能路徑</td>
					        <%}else{ %>
						        <td class="<%=listTitle%>"  width="10%">功能代碼</td>
						        <td class="<%=listTitle%>"  width="24%">報表名稱</td>
						        <td class="<%=listTitle%>"  width="66%">報表功能路徑</td>
					        <%} %>
					    </tr>
                   		    <% if(qList.size() == 0){%>
                   			   <tr class="<%=list1Color%>" >
                   			   <td colspan=11 align=center>無資料可供查詢</td><tr>
                   			   <tr>                   			   
                   			<% }
                    		   while(i < qList.size()){ 
                    		      listColor = (i % 2 == 0)?list2Color:list1Color; %>                         	  
		                          <tr>
		                          	<%if(!"".equals(sn_docno2)||!"".equals(sn_docno1)){ %>
								        <td class="<%=listColor%>" align="center" width="8%" >
								        	<a href="javascript:doSubmit(this.document.forms[0],'Edit','<%=(String)((DataObject)qList.get(i)).getValue("rpt_id")%>',<%=(((DataObject)qList.get(i)).getValue("rpt_no")).toString() %>)"><%=(String)((DataObject)qList.get(i)).getValue("rpt_id") %></a>
								        </td>
								        <td class="<%=listColor%>" align="center" width="24%" ><%=(String)((DataObject)qList.get(i)).getValue("rpt_name")%></td>
								        <td class="<%=listColor%>" align="left" width="20%" ><%=(String)((DataObject)qList.get(i)).getValue("col_name")%></td>
								        <td class="<%=listColor%>" align="left" width="69%" ><%=(String)((DataObject)qList.get(i)).getValue("rpt_path")%></td>
							        <%}else{ %>
								        <td class="<%=listColor%>" align="center" width="10%" >
								        	<a href="javascript:doSubmit(this.document.forms[0],'Edit','<%=(String)((DataObject)qList.get(i)).getValue("rpt_id")%>',<%=(((DataObject)qList.get(i)).getValue("rpt_no")).toString() %>)"><%=(String)((DataObject)qList.get(i)).getValue("rpt_id") %></a>
								        </td>
								        <td class="<%=listColor%>" align="left" width="24%" ><%=(String)((DataObject)qList.get(i)).getValue("rpt_name")%></td>
								        <td class="<%=listColor%>" align="left" width="66%" ><%=(String)((DataObject)qList.get(i)).getValue("rpt_path")%></td>
							        <%} %>
							      </tr> 
					      <%
                  			   i++;
	                  		   }//end of while
	                  		}//end of if
                  			%>            
					      </table> 
                                  
		      </table></td>
		</tr>
  
	  <table width="800" border="0" align="center" cellpadding="0" cellspacing="0">
	  <tr> 
	    <td colspan="2">
	      <p align="left">
	      <font color='#990000'>
	        <img src="images/arrow_1.gif" width="28" height="23" align="absmiddle">
	        <font color="#007D7D" size="3">使用說明  : </font></font></td>
	  </tr>
	  <tr> 
	    <td width="3%">&nbsp;</td>
	    <td width="88%"> 
	      <ul>                                                
	          
	      <%if("Qry".equals(act)){ %>
	      		<li>本網頁提供基本資料查詢功能。</li>
	            <li>按<font color="#666666">【查詢】</font>則依報表名稱<font color="#FF0000">關鍵字</font>模糊查詢資料。</li> 
			    <li>按<font color="#666666">【功能代碼】連結</font>則可查看此筆資料。</li>
	      <%}else{ %>
	      		<li><p align="left">按<font color="#666666">【查詢】</font>則依報表名稱<font color="#FF0000">關鍵字</font>模糊查詢資料。</li> 
	      <%} %>
	      </ul>
	    </td>
	  </tr>
	  </table>
 
</table>

</form>
</body>
<script language="JavaScript" >
<!--

-->
</script>

</html>
