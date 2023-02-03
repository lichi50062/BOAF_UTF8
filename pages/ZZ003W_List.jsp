<%
//93.12.28 create by 2295
//94.02.03 fix 區分網際網路申報/MIS管理系統的配色 by 2295
//94.02.15 fix 區分網際網路申報/MIS by 2295
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
	String muser_type = (session.getAttribute("muser_type") == null)?"":(String)session.getAttribute("muser_type");			
	String szweb_type =  ( request.getParameter("web_type")==null ) ? "" : (String)request.getParameter("web_type");				
	String szsys_type =  ( request.getParameter("sys_type")==null ) ? "" : (String)request.getParameter("sys_type");				
	String szprogram_type =  ( request.getParameter("program_type")==null ) ? "" : (String)request.getParameter("program_type");					
	String szMenu = ( request.getParameter("menu")==null ) ? "" : (String)request.getParameter("menu");					 	 
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");				
	
	System.out.println("szMenu="+szMenu);
	List WTT03_2List = (List)request.getAttribute("WTT03_2List");		
	if(WTT03_2List == null){
	   System.out.println("WTT03_2List == null");
	}else{
	   System.out.println("WTT03_2List.size()="+WTT03_2List.size());
	}
	
	//取得ZZ003W的權限
	Properties permission = ( session.getAttribute("ZZ003W")==null ) ? new Properties() : (Properties)session.getAttribute("ZZ003W"); 
	if(permission == null){
       System.out.println("ZZ003W_List.permission == null");
    }else{
       System.out.println("ZZ003W_List.permission.size ="+permission.size());               
    }	
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/ZZ003W.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<title>程式_歸屬系統類別維護</title>
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
<%if(WTT03_2List != null && WTT03_2List.size() != 0){%>
<input type="hidden" name="row" value="<%=WTT03_2List.size()+1%>">   
<%}%>
<table width="600" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  		<tr> 
   		 <td><img src="images/space_1.gif" width="12" height="12"></td>
  		</tr>

        <tr> 
          <td bgcolor="#FFFFFF">
		  <table width="640" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="640" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="170"><img src="images/banner_bg1.gif" width="170" height="17"></td>
                      <td width="300"><font color='#000000' size=4><b> 
                        <center>
                          「程式_歸屬系統類別維護」 
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
                <td><table width="640" border="0" align="center" cellpadding="0" cellspacing="0">
               
                    <tr> 
                      <div align="right"><jsp:include page="getLoginUser.jsp?width=640" flush="true" /></div> 
                    </tr>    
                    <%
                      String nameColor="nameColor_sbody";
                      String textColor="textColor_sbody";
                      String bordercolor="#3A9D99";
                     %>                
                    <tr> 
                      <td><table width=640 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="<%=bordercolor%>">
                          
                          <%if(act.equals("del")){%>
                          <% List web_type = DBManager.QueryDB_SQLParam("select cmuse_id,cmuse_name from cdshareno where cmuse_div='017' order by input_order",null,"");%>
                          <tr class="sbody">						  
						  <td width='15%' class="<%=nameColor%>">網站</td>
                          <td width='85%' class="<%=textColor%>"'>	
                            <select name='WEB_TYPE'>                                                        
                            <%for(int i=0;i<web_type.size();i++){%>
                            <option value="<%=(String)((DataObject)web_type.get(i)).getValue("cmuse_id")%>"                                                        
                            <%if(szweb_type.equals((String)((DataObject)web_type.get(i)).getValue("cmuse_id"))) out.print("selected");%>
                            ><%=(String)((DataObject)web_type.get(i)).getValue("cmuse_name")%></option>                            
                            <%}%>
                            </select>                                 
                          </td>             
                          </tr>            
                          <% List sys_type = DBManager.QueryDB_SQLParam("select cmuse_id,cmuse_name from cdshareno where cmuse_div='018' order by input_order",null,"");%>
                          <tr class="sbody">                          
                          <td width='15%' class="<%=nameColor%>">系統類別</td>
                          <td width='85%' class="<%=textColor%>">
                            <select name='SYS_TYPE'>                                                        
                            <%for(int i=0;i<sys_type.size();i++){%>
                            <option value="<%=(String)((DataObject)sys_type.get(i)).getValue("cmuse_id")%>"                                                        
                            <%if(szsys_type.equals((String)((DataObject)sys_type.get(i)).getValue("cmuse_id"))) out.print("selected");%>
                            ><%=(String)((DataObject)sys_type.get(i)).getValue("cmuse_name")%></option>                            
                            <%}%>
                            </select>                                 
                          </td>         
                          </tr>
                          <% List program_type = DBManager.QueryDB_SQLParam("select cmuse_id,cmuse_name from cdshareno where cmuse_div='016' order by input_order",null,"");%>
                          <tr class="sbody">                          
						  <td width='15%' class="<%=nameColor%>">程式歸屬機構</td>
                          <td width='85%' class="<%=textColor%>">	
                            <select name='PROGRAM_TYPE'>                                                        
                            <%for(int i=0;i<program_type.size();i++){%>
                            <option value="<%=(String)((DataObject)program_type.get(i)).getValue("cmuse_id")%>"                                                        
                            <%if(szprogram_type.equals((String)((DataObject)program_type.get(i)).getValue("cmuse_id"))) out.print("selected");%>
                            ><%=(String)((DataObject)program_type.get(i)).getValue("cmuse_name")%></option>                            
                            <%}%>
                            </select>                                 
                            <a href="javascript:selectAll(this.document.forms[0],'del');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_selectallb.gif',1)"><img src="images/bt_selectall.gif" name="Image101" width="80" height="25" border="0" id="Image101"></a>                       
 						    <a href="javascript:selectNo(this.document.forms[0],'del');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image102','','images/bt_selectnob.gif',1)"><img src="images/bt_selectno.gif" name="Image102" width="80" height="25" border="0" id="Image102"></a>                        							
 						    <a href="javascript:doSubmit(this.document.forms[0],'DelQry');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image106','','images/bt_queryb.gif',1)"><img src="images/bt_query.gif" name="Image106" width="66" height="25" border="0" id="Image106"></a>                       
 						    <a href="javascript:doSubmit(this.document.forms[0],'Delete');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image104','','images/bt_deleteb.gif',1)"><img src="images/bt_delete.gif" name="Image104" width="66" height="25" border="0" id="Image104"></a>                       
 						    <a href="javascript:history.back();"  onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_backb.gif',1)"><img src="images/bt_back.gif" name="Image105" width="80" height="25" border="0" id="Image105"></a>
                          </td>                                   
                          </tr>
                          <%}%>
                          <tr class="sbody">
                          <td width='100%' colspan=2 class="<%=textColor%>">							
 						  <%if(act.equals("List") || act.equals("Qry")){%>
 							<a href="javascript:doSubmit(this.document.forms[0],'Qry');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image106','','images/bt_queryb.gif',1)"><img src="images/bt_query.gif" name="Image106" width="66" height="25" border="0" id="Image106"></a>                       
 							<a href="javascript:doSubmit(this.document.forms[0],'new');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image107','','images/bt_addb.gif',1)"><img src="images/bt_add.gif" name="Image107" width="66" height="25" border="0" id="Image107"></a>                       
 							<a href="javascript:doSubmit(this.document.forms[0],'del');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image108','','images/bt_deleteb.gif',1)"><img src="images/bt_delete.gif" name="Image108" width="66" height="25" border="0" id="Image108"></a>                       
 						  <%}%> 							
                          </td>
                          </tr>     					      					      
                          </table>      
                      </td>    
                      </tr>
                      
                      
                      <tr> 
                      <td><table width=640 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="<%=bordercolor%>">                      
                      <%
                      String tmpbank_type="";                      
                      int i = 0;      
                      String listTitle="listTitleColor_sbody"; 
                      String list1Color="list1Color_sbody";
                      String list2Color="list2Color_sbody";
                      String listColor="list1Color_sbody";    
                      String usertype="";  
                      if(WTT03_2List != null){ %>
                       	  <tr class="<%=listTitle%>">                       	  
                            <td width="25">序號</td>       
                            <%if(act.equals("del")){%>    
                            <td width="25">選項</td>
                            <%}%>
                            <td width="100">程式代碼</td>
            				<td width="180">程式名稱</td>            				
            				<td width="100">系統類別</td>
            				<td width="100">歸屬機構類別</td>            				
            				<td width="100">網站</td>            				            				
					      </tr>   
                   		    <% if(WTT03_2List.size() == 0){%>
                   			   <tr class="<%=list1Color%>">
                   			   <td colspan=10 align=center>無資料可供查詢</td><tr>
                   			   <tr>                   			   
                   			<% }
                    		   while(i < WTT03_2List.size()){ 
                    		     if(( szMenu.equals("MIS") && 
                    		        ( ((String)((DataObject)WTT03_2List.get(i)).getValue("web_type")).equals("2") || 
                    		          ((String)((DataObject)WTT03_2List.get(i)).getValue("program_id")).startsWith("ZZ") ) 
                    		        )/*MIS*/
                    		        ||
                    		        ( szMenu.equals("") && 
                    		        ( ((String)((DataObject)WTT03_2List.get(i)).getValue("web_type")).equals("1") || 
                    		          ((String)((DataObject)WTT03_2List.get(i)).getValue("program_id")).startsWith("ZZ") )
                    		        )/*網際網路申報*/
                    		       )
                    		       {
                    		     listColor = (i % 2 == 0)?list2Color:list1Color;	                         		      
                      %>                         	  
                          <tr class="<%=listColor%>">
                            <td width="25"><%=i+1%></td>                       				
            				<%if(act.equals("del")){%>		
            				<td width="25">
            				<input type="checkbox" name="isModify_<%=(i+1)%>" value="<%=(String)((DataObject)WTT03_2List.get(i)).getValue("sys_type")%>:<%=(String)((DataObject)WTT03_2List.get(i)).getValue("program_type")%>:<%=(String)((DataObject)WTT03_2List.get(i)).getValue("program_id")%>">
            				</td>
            				<%}%>                     		
            				<td width="100"><%if( ((DataObject)WTT03_2List.get(i)).getValue("program_id") != null ) out.print((String)((DataObject)WTT03_2List.get(i)).getValue("program_id")); else out.print("&nbsp;");%></a></td>
            				<td width="180"><%if( ((DataObject)WTT03_2List.get(i)).getValue("program_name") != null ) out.print((String)((DataObject)WTT03_2List.get(i)).getValue("program_name")); else out.print("&nbsp;");%></a></td>
            				<td width="100"><%if( ((DataObject)WTT03_2List.get(i)).getValue("sys_type_name") != null ) out.print((String)((DataObject)WTT03_2List.get(i)).getValue("sys_type_name")); else out.print("&nbsp;");%></a></td>
            				<td width="100"><%if( ((DataObject)WTT03_2List.get(i)).getValue("program_type_name") != null ) out.print((String)((DataObject)WTT03_2List.get(i)).getValue("program_type_name")); else out.print("&nbsp;");%></a></td>
            				<td width="100"><%if( ((DataObject)WTT03_2List.get(i)).getValue("web_type_name") != null ) out.print((String)((DataObject)WTT03_2List.get(i)).getValue("web_type_name")); else out.print("&nbsp;");%></a></td>            				
					      </tr> 					      
					      <%     }//end of BOAF/MIS
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
