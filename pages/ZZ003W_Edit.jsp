<%
//93.12.28 create by 2295
//94.02.03 fix 區分網際網路申報/MIS管理系統的配色 by 2295
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
	
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");				
	
	
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
<title>「程式_歸屬系統類別維護」新增建檔 </title>
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
		  <table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="140"><img src="images/banner_bg1.gif" width="140" height="17"></td>
                      <td width="320"><font color='#000000' size=4><b> 
                        <center>
                          「程式_歸屬系統類別維護」新增建檔  
                        </center>
                        </b></font> </td>
                      <td width="140"><img src="images/banner_bg1.gif" width="140" height="17"></td>
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
                          <% List program_id = DBManager.QueryDB_SQLParam("select program_id,program_name from WTT03_1 order by program_id",null,"");%>                          
                          <tr class="sbody">						  
						  <td width='15%' class="<%=nameColor%>">程式代碼</td>
                          <td width='85%' class="<%=textColor%>"'>	
                            <select name='PROGRAM_ID'>                                                        
                            <%for(int i=0;i<program_id.size();i++){%>
                            <option value="<%=(String)((DataObject)program_id.get(i)).getValue("program_id")%>"                                                                                    
                            ><%=(String)((DataObject)program_id.get(i)).getValue("program_id")%>&nbsp;<%=(String)((DataObject)program_id.get(i)).getValue("program_name")%></option>                            
                            <%}%>
                            </select>                                 
                          </td>             
                          </tr>            
                          <% List web_type = DBManager.QueryDB_SQLParam("select cmuse_id,cmuse_name from cdshareno where cmuse_div='017' order by input_order",null,"");%>
                          <tr class="sbody">						  
						  <td width='15%' class="<%=nameColor%>">網站</td>
                          <td width='85%' class="<%=textColor%>">	
                            <select name='WEB_TYPE'>                                                        
                            <%for(int i=0;i<web_type.size();i++){%>
                            <option value="<%=(String)((DataObject)web_type.get(i)).getValue("cmuse_id")%>"                                                        
                            ><%=(String)((DataObject)web_type.get(i)).getValue("cmuse_name")%></option>                            
                            <%}%>
                            </select>  
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                               
                            <a href="javascript:doSubmit(this.document.forms[0],'Insert');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_confirmb.gif',1)"><img src="images/bt_confirm.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>                       
                            <a href="javascript:history.back();"  onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image102','','images/bt_backb.gif',1)"><img src="images/bt_back.gif" name="Image102" width="80" height="25" border="0" id="Image102"></a>
                          </td>             
                          </tr>                                                                
                          </table>      
                      </td>    
                      </tr>
                      <tr><td><table><tr><br></tr></table></td></tr>  
                      
                      <tr> 
                      <td><table width=600 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="<%=bordercolor%>">                      
                        <tr class="sbody">
                          <td width='100%' colspan=4 bgcolor='D2F0FF'><b>系統類別：</b> 						 						  
 							<a href="javascript:selectAll(this.document.forms[0],'sys_type');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_selectallb.gif',1)"><img src="images/bt_selectall.gif" name="Image103" width="80" height="25" border="0" id="Image103"></a>                       
 						    <a href="javascript:selectNo(this.document.forms[0],'sys_type');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image104','','images/bt_selectnob.gif',1)"><img src="images/bt_selectno.gif" name="Image104" width="80" height="25" border="0" id="Image104"></a>                        							 						    
                          </td>
                        </tr>     					      					      
                        
                      <%                                            
                      String bgcolor="#D3EBE0";                       
                      List sys_type_List = DBManager.QueryDB_SQLParam("select cmuse_id,cmuse_name from cdshareno where cmuse_div='018' order by input_order",null,"");
                      String listTitle="listTitleColor_sbody"; 
                      String list1Color="list1Color_sbody";
                      String list2Color="list2Color_sbody";
                      String listColor="list1Color_sbody";
                      if(sys_type_List != null){ %>
                          <input type="hidden" name="row_systype" value="<%=sys_type_List.size()+1%>">   
                       	  <tr class="<%=listTitle%>">                       	  
                            <td width="30">序號</td>                                   
                            <td width="30">選項</td>                                                                    				
            				<td width="100">系統類別</td>
            				<td width="400">系統類別名稱</td>            				            				
					      </tr>                      		    
                      <%
                      	 String sys_type="";
                      	 String sys_type_id="";
                      	 String sys_type_name="";
                         for(int i=0;i<sys_type_List.size();i++){    
                             listColor = (i % 2 == 0)?list2Color:list1Color;	  	                       
                             sys_type = (String)((DataObject)sys_type_List.get(i)).getValue("cmuse_name");
                             sys_type_id = sys_type.substring(0,sys_type.indexOf("-"));
                             sys_type_name = sys_type.substring(sys_type.indexOf("-")+1,sys_type.length());
                      %>    
                          <tr class="<%=listColor%>">
                            <td width="30"><%=i+1%></td>                       				            				
            				<td width="30">
            				<input type="checkbox" name="Systype_isModify_<%=(i+1)%>" value="<%=(String)((DataObject)sys_type_List.get(i)).getValue("cmuse_id")%>">
            				</td>            				          		
            				<td width="100"><%=sys_type_id%></td>
            				<td width="400"><%=sys_type_name%></td>            				            				            				
					      </tr> 					      
					   <%
					     }//end of for
	                  	 }//end of if
    			       %>  
                  		 
                  		 
                  	  <tr class="sbody">
                         <td width='100%' colspan=4 bgcolor='D2F0FF'><b>程式歸屬機構：</b> 						 						  
 						  <a href="javascript:selectAll(this.document.forms[0],'program_type');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_selectallb.gif',1)"><img src="images/bt_selectall.gif" name="Image105" width="80" height="25" border="0" id="Image105"></a>                       
 						  <a href="javascript:selectNo(this.document.forms[0],'program_type');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image106','','images/bt_selectnob.gif',1)"><img src="images/bt_selectno.gif" name="Image106" width="80" height="25" border="0" id="Image106"></a>                        							 						     						 
                         </td>
                      </tr>     					      					      
                      <%                                                                  
                      List program_type_List = DBManager.QueryDB_SQLParam("select cmuse_id,cmuse_name from cdshareno where cmuse_div='016' order by input_order",null,"");
                      if(program_type_List != null){ %>
                          <input type="hidden" name="row_programtype" value="<%=program_type_List.size()+1%>">
                       	  <tr class="<%=listTitle%>">                       	  
                            <td width="30">序號</td>                                   
                            <td width="30">選項</td>                                                                    				
            				<td width="120">程式歸屬機構代碼</td>
            				<td width="400">程式歸屬機構名稱</td>            				            				
					      </tr>                      		    
                      <%
                         for(int i=0;i<program_type_List.size();i++){    
                             listColor = (i % 2 == 0)?list2Color:list1Color;		                                                    
                      %>    
                          <tr class="<%=listColor%>">
                            <td width="30"><%=i+1%></td>                       				            				
            				<td width="30">
            				<input type="checkbox" name="Programtype_isModify_<%=(i+1)%>" value="<%=(String)((DataObject)program_type_List.get(i)).getValue("cmuse_id")%>">
            				</td>            				          		
            				<td width="120"><%=(String)((DataObject)program_type_List.get(i)).getValue("cmuse_id")%></td>
            				<td width="400"><%=(String)((DataObject)program_type_List.get(i)).getValue("cmuse_name")%></td>            				            				            				
					      </tr> 					      
					   <%
					     }//end of for
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
