<%
//94.01.06 fix 增加機構類別bank_type,機構代號tbank_no為查詢條件 by 2295
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
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");				
    String szbank_type = ( request.getParameter("bank_type")==null ) ? "" : (String)request.getParameter("bank_type");		
	String sztbank_no = ( request.getParameter("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");		
    		
	System.out.println("ZZ002W_List.act="+act);
	List WTT02List = (List)request.getAttribute("WTT02List");		
	if(WTT02List == null){
	   System.out.println("WTT02List == null");
	}else{
	   System.out.println("WTT02List.size()="+WTT02List.size());
	}
	
	
	String title="「金融機構類別維護權限設定」";		
	title =(muser_type.equals("S"))?"系統管理者"+title:title;
	title =(muser_type.equals("A"))?"使用者"+title:title;
	title =(act.equals("del"))?title+"異動建檔":title;
	title =(act.equals("new"))?title+"新增建檔":title;
	System.out.println("muser_type="+muser_type);
	
	
	//取得ZZ002W的權限
	Properties permission = ( session.getAttribute("ZZ002W")==null ) ? new Properties() : (Properties)session.getAttribute("ZZ002W"); 
	if(permission == null){
       System.out.println("ZZ002W_List.permission == null");
    }else{
       System.out.println("ZZ002W_List.permission.size ="+permission.size());               
    }	
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/ZZ002W.js"></script>
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
<%if(WTT02List != null && WTT02List.size() != 0){%>
<input type="hidden" name="row" value="<%=WTT02List.size()+1%>">   
<%}%>
<table width="640" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  		<tr> 
   		 <td><img src="images/space_1.gif" width="12" height="12"></td>
  		</tr>

        <tr> 
          <td bgcolor="#FFFFFF">
		  <table width="640" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="640" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="100"><img src="images/banner_bg1.gif" width="100" height="17"></td>
                      <td width="440"><font color='#000000' size=4><b> 
                        <center>
                          <%=title%>
                        </center>
                        </b></font> </td>
                      <td width="100"><img src="images/banner_bg1.gif" width="100" height="17"></td>
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
                      <td><table width=640 border=1 align=center cellpadding="1" cellspacing="1"  bordercolor="<%=bordercolor%>">
                          <tr class="sbody">
                          <% List bank_type = DBManager.QueryDB_SQLParam("select cmuse_id,cmuse_name from cdshareno where cmuse_div='001' and cmuse_id <> 'Z' order by input_order",null,"");%>						  
						  <td width='12%' class="<%=nameColor%>">機構類別</td>
						  <td width='88%' colspan=2 class="<%=textColor%>">														  
                            <select name='BANK_TYPE' onchange="javascript:getData(this.document.forms[0],'<%=act%>');"
                             <%if(!muser_type.equals("S")) out.print("disabled");%>                                                          
                            >                            
                            <%for(int i=0;i<bank_type.size();i++){%>
                            <option value="<%=(String)((DataObject)bank_type.get(i)).getValue("cmuse_id")%>"                            
                            <%if(szbank_type.equals((String)((DataObject)bank_type.get(i)).getValue("cmuse_id"))) out.print("selected");%>
                            ><%=(String)((DataObject)bank_type.get(i)).getValue("cmuse_name")%></option>                            
                            <%}%>
                            <%if(muser_type.equals("S")) out.print("<option value='ALL'>全部</option>");%>                            
                            </select>                                                                       
                          </td>
                          </tr>  
                          
                          <tr class="sbody">
						  <td width='12%' class="<%=nameColor%>">總機構代碼</td>
						  <td width='88%' colspan=2 class="<%=textColor%>">														  
						  <% List tbank_no = (List)request.getAttribute("tbank_no");%>
						
                            <select name='TBANK_NO'
                            <%if(!muser_type.equals("S")) out.print("disabled");%>                                                                                 
                            >                            
                            <%
                            if(muser_type.equals("S")) out.print("<option value='ALL'>全部</option>");
                            if(tbank_no != null && tbank_no.size() != 0){
                            for(int i=0;i<tbank_no.size();i++){%>
                            <option value="<%=(String)((DataObject)tbank_no.get(i)).getValue("bank_no")%>"                            
                            <%if( ((DataObject)tbank_no.get(i)).getValue("bank_no") != null &&  (sztbank_no.equals((String)((DataObject)tbank_no.get(i)).getValue("bank_no")))) out.print("selected");%>
                            >
                            <%//if(((DataObject)tbank_no.get(i)).getValue("bank_no") != null) out.print((String)((DataObject)tbank_no.get(i)).getValue("bank_no"));%>
                            
                            <%if(((DataObject)tbank_no.get(i)).getValue("bank_name") != null) out.print((String)((DataObject)tbank_no.get(i)).getValue("bank_name"));%>                            
                            </option>                            
                            <%}
                            }%>
                            </select>                                                        
                            <%if(act.equals("new") || act.equals("del")){%>
                            <a href="javascript:doSubmit(this.document.forms[0],'<%=act%>Qry');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image106','','images/bt_queryb.gif',1)"><img src="images/bt_query.gif" name="Image106" width="66" height="25" border="0" id="Image106"></a>                       
                               <%if(act.equals("new")){%>
                               <a href="javascript:doSubmit(this.document.forms[0],'Insert');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_addb.gif',1)"><img src="images/bt_add.gif" name="Image103" width="66" height="25" border="0" id="Image103"></a>                       
                               <%}%>
                               <%if(act.equals("del")){%>
                               <a href="javascript:doSubmit(this.document.forms[0],'Delete');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image104','','images/bt_deleteb.gif',1)"><img src="images/bt_delete.gif" name="Image104" width="66" height="25" border="0" id="Image104"></a>                        							
 							   <%}%>
 							   <a href="javascript:history.back();"  onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_backb.gif',1)"><img src="images/bt_back.gif" name="Image105" width="80" height="25" border="0" id="Image105"></a> 							 							   
                            <%}%>
 							<%if(act.equals("List") || act.equals("Qry")){%>
 							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 							
 							<a href="javascript:doSubmit(this.document.forms[0],'Qry');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image106','','images/bt_queryb.gif',1)"><img src="images/bt_query.gif" name="Image106" width="66" height="25" border="0" id="Image106"></a>                       
 							<a href="javascript:doSubmit(this.document.forms[0],'new');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image107','','images/bt_addb.gif',1)"><img src="images/bt_add.gif" name="Image107" width="66" height="25" border="0" id="Image107"></a>                       
 							<a href="javascript:doSubmit(this.document.forms[0],'del');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image108','','images/bt_deleteb.gif',1)"><img src="images/bt_delete.gif" name="Image108" width="66" height="25" border="0" id="Image108"></a>                       
 							<%}%> 							
                          </td>                          
                          </tr>
                          </table>      
                      </td>    
                      </tr>
                      <tr><td><table><tr><br></tr></table></td></tr>  
                      
                      <%if(WTT02List != null){ %>
                      <tr> 
                      <td><table width=640 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="<%=bordercolor%>">                      
                        <%if(act.equals("del")){%>
                          <tr>
                          <td width='100%' colspan=3 bgcolor="#D2F0FF">	                            
 							<a href="javascript:selectAll(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_selectallb.gif',1)"><img src="images/bt_selectall.gif" name="Image101" width="80" height="25" border="0" id="Image101"></a>                       
 							<a href="javascript:selectNo(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image102','','images/bt_selectnob.gif',1)"><img src="images/bt_selectno.gif" name="Image102" width="80" height="25" border="0" id="Image102"></a>                        							 							 							 							
 					      </td>		      					      					      
                          </tr>
                        <%}%>
                          
                        <%if(act.equals("new")){%>
                          <tr class="sbody">                          
                          <% List bank_typeList = (List)request.getAttribute("bank_typeList");%>		
						  <td width='100%' colspan=3 bgcolor="#D2F0FF"><b>欲維護的金融機構類別</b>						  
                            <select name="maintainBANK_TYPE">                            
                            <%for(int i=0;i<bank_typeList.size();i++){%>
                            <option value="<%=(String)((DataObject)bank_typeList.get(i)).getValue("bank_type")%>">
                            <%=(String)((DataObject)bank_typeList.get(i)).getValue("cmuse_name")%>
                            </option>                            
                            <%}%>                            
                            </select>      
                            <a href="javascript:selectAll(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_selectallb.gif',1)"><img src="images/bt_selectall.gif" name="Image101" width="80" height="25" border="0" id="Image101"></a>                       
 							<a href="javascript:selectNo(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image102','','images/bt_selectnob.gif',1)"><img src="images/bt_selectno.gif" name="Image102" width="80" height="25" border="0" id="Image102"></a>                        							 							 							 							
                          </td>
                          </tr>  
                        <%}%>
                       </table>   
                       </td>
                      </tr> 
                      <%}%>
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
                      if(WTT02List != null){ %>
                       	  <tr class="<%=listTitle%>">                       	  
                            <td width="25">序號</td>       
                            <%if(act.equals("new") || act.equals("del")){%>    
                            <td width="25">選項</td>
                            <%}%>
                            <td width="60">機構<br>類別</td>
            				<td width="160">總機構代碼</td>
            				<td width="80">組室</td>
            				<td width="50">科別</td>
            				<td width="70">使用者<br>帳號</td>            				
            				<td width="100">名稱</td>            				
            				<td width="80">類型</td>            				
            				<%if(act.equals("Qry") || act.equals("del")){%>
            				<td width="130">維護的<br>機構類別</td>           
            				<%}%> 				
					      </tr>   
                   		    <% if(WTT02List.size() == 0){%>
                   			   <tr class="<%=list1Color%>"">
                   			   <td colspan=10 align=center>無資料可供查詢</td><tr>
                   			   <tr>                   			   
                   			<% }
                    		   while(i < WTT02List.size()){ 
                    		     listColor = (i % 2 == 0)?list2Color:list1Color;	                    		      
                      %>                         	  
                          <tr class="<%=listColor%>"">
                            <td width="25"><%=i+1%></td>                       				
                          
                            <%if(act.equals("new")){%>            				
                            <td width="25">
            				<input type="checkbox" name="isModify_<%=(i+1)%>" value="<%=(String)((DataObject)WTT02List.get(i)).getValue("muser_id")%>">            				     
            				</td>
            				<%}%>            		            				
            				<%if(act.equals("del")){%>		
            				<td width="25">
            				<input type="checkbox" name="isModify_<%=(i+1)%>" value="<%=(String)((DataObject)WTT02List.get(i)).getValue("pg_bank_type")%>:<%=(String)((DataObject)WTT02List.get(i)).getValue("muser_id")%>">
            				</td>
            				<%}%>                     		
            				<td width="60"><%if( ((DataObject)WTT02List.get(i)).getValue("bank_type_name") != null ) out.print((String)((DataObject)WTT02List.get(i)).getValue("bank_type_name")); else out.print("&nbsp;");%></a></td>
            				<td width="160"><%if( ((DataObject)WTT02List.get(i)).getValue("tbank_no_name") != null ) out.print((String)((DataObject)WTT02List.get(i)).getValue("tbank_no_name")); else out.print("&nbsp;");%></a></td>
            				<td width="80"><%if( ((DataObject)WTT02List.get(i)).getValue("bank_no_name") != null ) out.print((String)((DataObject)WTT02List.get(i)).getValue("bank_no_name")); else out.print("&nbsp;");%></a></td>
            				<td width="50"><%if( ((DataObject)WTT02List.get(i)).getValue("subdep_id_name") != null ) out.print((String)((DataObject)WTT02List.get(i)).getValue("subdep_id_name")); else out.print("&nbsp;");%></a></td>
            				<td width="70"><%if( ((DataObject)WTT02List.get(i)).getValue("muser_id") != null ) out.print((String)((DataObject)WTT02List.get(i)).getValue("muser_id")); else out.print("&nbsp;");%></a></td>
            				<td width="100"><%if( ((DataObject)WTT02List.get(i)).getValue("muser_name") != null ) out.print((String)((DataObject)WTT02List.get(i)).getValue("muser_name")); else out.print("&nbsp;");%></a></td>            				
            				<td width="80">
            				<%if( ((DataObject)WTT02List.get(i)).getValue("muser_type") != null ){
                    		     if(((String)((DataObject)WTT02List.get(i)).getValue("muser_type")).equals("A")){
                    		        usertype="管理者";
                    		     }else{
                    		        usertype="一般<br>使用者";
                    		     }
                    		  }
                    		  out.print(usertype);
                    		%>
            				</a></td>
            				<%if(act.equals("Qry") || act.equals("del")){%>
            				<td width="130"><%if( ((DataObject)WTT02List.get(i)).getValue("pg_type_name") != null ) out.print((String)((DataObject)WTT02List.get(i)).getValue("pg_type_name")); else out.print("&nbsp;");%></a></td>            				            				
            				<%}%>
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
