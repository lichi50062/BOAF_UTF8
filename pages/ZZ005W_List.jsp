<%
//93.12.29 create by 2295
//94.02.03 fix 區分網際網路申報/MIS管理系統的配色 by 2295
//94.03.07 fix 增加order by program_id,input_order,sub_type,order_type by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");							
	
	//取得ZZ005W的權限
	Properties permission = ( session.getAttribute("ZZ005W")==null ) ? new Properties() : (Properties)session.getAttribute("ZZ005W"); 
	if(permission == null){
       System.out.println("ZZ005W_List.permission == null");
    }else{
       System.out.println("ZZ005W_List.permission.size ="+permission.size());               
    }	
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/ZZ005W.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<title>程式功能維護管理</title>
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
<%List WTT03_1List = DBManager.QueryDB_SQLParam("select * from wtt03_1 order by program_id,input_order,sub_type,order_type",null,""); %>
<%if(WTT03_1List != null && WTT03_1List.size() != 0){%>
<input type="hidden" name="row" value="<%=WTT03_1List.size()+1%>">   
<%}%>
<table width="600" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  		<tr> 
   		 <td><img src="images/space_1.gif" width="12" height="12"></td>
  		</tr>

        <tr> 
          <td bgcolor="#FFFFFF">
		  <table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="170"><img src="images/banner_bg1.gif" width="170" height="17"></td>
                      <td width="300"><font color='#000000' size=4><b> 
                        <center>
                          「程式功能維護管理」 
                        </center>
                        </b></font> </td>
                      <td width="170"><img src="images/banner_bg1.gif" width="170" height="17"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><img src="images/space_1.gif" width="12" height="12"></td>
              </tr>
              <tr> 
                <td><table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
               
                    <tr> 
                      <div align="right"><jsp:include page="getLoginUser.jsp?width=600" flush="true" /></div> 
                    </tr>        
                     <%
                      String nameColor="nameColor_sbody";
                      String textColor="textColor_sbody";
                      String bordercolor="#3A9D99";
                     %>            
                    <tr> 
                      <td><table width=600 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="<%=bordercolor%>">                          
                          <tr class="<%=textColor%>">
                          <td width='100%' colspan=2>	
                          <%if(act.equals("List") || act.equals("Qry")){%>			
                            <%if(permission != null && permission.get("Q") != null && permission.get("Q").equals("Y")){ %>                   	        	                                   		     			 						
 							<a href="javascript:doSubmit(this.document.forms[0],'Qry');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image106','','images/bt_queryb.gif',1)"><img src="images/bt_query.gif" name="Image106" width="66" height="25" border="0" id="Image106"></a>                       
 							<%}%>
 							<%if(permission != null && permission.get("A") != null && permission.get("A").equals("Y")){ %>                   	        	                                   		      							
 							<a href="javascript:doSubmit(this.document.forms[0],'new');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image107','','images/bt_addb.gif',1)"><img src="images/bt_add.gif" name="Image107" width="66" height="25" border="0" id="Image107"></a>                       
 							<%}%>
 							<%if(permission != null && permission.get("D") != null && permission.get("D").equals("Y")){ %>                   	        	                                   		     
 							<a href="javascript:doSubmit(this.document.forms[0],'del');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image108','','images/bt_deleteb.gif',1)"><img src="images/bt_delete.gif" name="Image108" width="66" height="25" border="0" id="Image108"></a>                        						  
 							<%}%>
 						  <%}%>	
 						  <%if(act.equals("del")){%>
 						    <%if(permission != null && permission.get("D") != null && permission.get("D").equals("Y")){ %>                   	        	                                   		     
 						    <a href="javascript:selectAll(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_selectallb.gif',1)"><img src="images/bt_selectall.gif" name="Image101" width="80" height="25" border="0" id="Image101"></a>                       
  						    <a href="javascript:selectNo(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image102','','images/bt_selectnob.gif',1)"><img src="images/bt_selectno.gif" name="Image102" width="80" height="25" border="0" id="Image102"></a>                        							  						    
  						    <a href="javascript:doSubmit(this.document.forms[0],'Delete');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image108','','images/bt_deleteb.gif',1)"><img src="images/bt_delete.gif" name="Image108" width="66" height="25" border="0" id="Image108"></a>                        						   						    
  						    <%}%>
 						  <%}%>
                          </td>
                          </tr>     					      					      
                          </table>      
                      </td>    
                      </tr>
                      
                      
                      <tr> 
                      <td><table width=600 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="<%=bordercolor%>">                      
                      <%        
                      int i = 0;                            
                      String listTitle="listTitleColor_sbody"; 
                      String list1Color="list1Color_sbody";
                      String list2Color="list2Color_sbody";
                      String listColor="list1Color_sbody";
                      if(WTT03_1List != null){ %>
                       	  <tr class="<%=listTitle%>">                       	  
                            <td width="30">序號</td>     
                            <%if(act.equals("del")){%>		
                            <td width="25">選項</td>                            
                            <%}%>
                            <td width="100">程式代碼</td>
            				<td width="200">程式名稱</td>            				            				
            				<td width="25">新增</td>          
            				<td width="25">刪除</td>          
            				<td width="25">修改</td>          
            				<td width="25">查詢</td>          
            				<td width="25">印表</td>          
            				<td width="25">上傳</td>          
            				<td width="25">下載</td>          
            				<td width="25">鎖定</td>          
            				<td width="25">申報</td>          
					      </tr>   
                   		    <% if(WTT03_1List.size() == 0){%>
                   			   <tr class="<%=list1Color%>">
                   			   <td colspan=10 align=center>無資料可供查詢</td><tr>
                   			   <tr>                   			   
                   			<% }
                    		   while(i < WTT03_1List.size()){ 
                    		      listColor = (i % 2 == 0)?list2Color:list1Color;	                      		      
                      %>                         	  
                          <tr class="<%=listColor%>">
                            <td width="30"><%=i+1%></td>                       				
            				<%if(act.equals("del")){%>		
            				<td width="25">
            				<input type="checkbox" name="isModify_<%=(i+1)%>" value="<%=(String)((DataObject)WTT03_1List.get(i)).getValue("program_id")%>">
            				</td>
            				<%}%>                     		
            				<td width="100"><a href="javascript:doSubmit(this.document.forms[0],'Edit','<%=(String)((DataObject)WTT03_1List.get(i)).getValue("program_id")%>');"><%if( ((DataObject)WTT03_1List.get(i)).getValue("program_id") != null ) out.print((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")); else out.print("&nbsp;");%></a></td>
            				<td width="200"><a href="javascript:doSubmit(this.document.forms[0],'Edit','<%=(String)((DataObject)WTT03_1List.get(i)).getValue("program_id")%>');"><%if( ((DataObject)WTT03_1List.get(i)).getValue("program_name") != null ) out.print((String)((DataObject)WTT03_1List.get(i)).getValue("program_name")); else out.print("&nbsp;");%></a></td>            				
            				<td width="25"><%if( ((DataObject)WTT03_1List.get(i)).getValue("p_add") != null ) out.print((String)((DataObject)WTT03_1List.get(i)).getValue("p_add")); else out.print("&nbsp;");%></td>
            				<td width="25"><%if( ((DataObject)WTT03_1List.get(i)).getValue("p_delete") != null ) out.print((String)((DataObject)WTT03_1List.get(i)).getValue("p_delete")); else out.print("&nbsp;");%></td>
            				<td width="25"><%if( ((DataObject)WTT03_1List.get(i)).getValue("p_update") != null ) out.print((String)((DataObject)WTT03_1List.get(i)).getValue("p_update")); else out.print("&nbsp;");%></td>
            				<td width="25"><%if( ((DataObject)WTT03_1List.get(i)).getValue("p_query") != null ) out.print((String)((DataObject)WTT03_1List.get(i)).getValue("p_query")); else out.print("&nbsp;");%></td>
            				<td width="25"><%if( ((DataObject)WTT03_1List.get(i)).getValue("p_print") != null ) out.print((String)((DataObject)WTT03_1List.get(i)).getValue("p_print")); else out.print("&nbsp;");%></td>
            				<td width="25"><%if( ((DataObject)WTT03_1List.get(i)).getValue("p_upload") != null ) out.print((String)((DataObject)WTT03_1List.get(i)).getValue("p_upload")); else out.print("&nbsp;");%></td>
            				<td width="25"><%if( ((DataObject)WTT03_1List.get(i)).getValue("p_download") != null ) out.print((String)((DataObject)WTT03_1List.get(i)).getValue("p_download")); else out.print("&nbsp;");%></td>
            				<td width="25"><%if( ((DataObject)WTT03_1List.get(i)).getValue("p_lock") != null ) out.print((String)((DataObject)WTT03_1List.get(i)).getValue("p_lock")); else out.print("&nbsp;");%></td>            				
            				<td width="25"><%if( ((DataObject)WTT03_1List.get(i)).getValue("p_other") != null ) out.print((String)((DataObject)WTT03_1List.get(i)).getValue("p_other")); else out.print("&nbsp;");%></td>            				
					      </tr> 					      
					      <%
                  			   i++;
	                  		   }//end of while
	                  		}//end of if
    			          %>  
                  		  
					      </table>      
                      </td>    
                      </tr>
                                  
      </table></td>
  </tr> 
</table>
</form>
</body>
</html>
