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
	List WTT01List = null;
	if(act.equals("Qry")){
	   DBManager.QueryDB_SQLParam("select * from wtt01",null,"");
	   if(WTT01List == null){
	      System.out.println("WTT01List == null");
	   }else{
	      System.out.println("WTT01List.size()="+WTT01List.size());
	   }
	}   
	String title="「使用者帳號」管理作業";			
%>
<script language="javascript" src="js/Common.js"></script>
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
<script language="JavaScript" type="text/JavaScript">
function doSubmit(form,cnd){	     	    	    
	    form.action="/pages/ZZ001W_BOAF.jsp?act="+cnd+"&test=nothing";	    
	    form.submit();
}	

function getData(form,cnd){	     	    	    
	    form.action="/pages/ZZ001W_BOAF.jsp?act=getData&nowact="+cnd+"&test=nothing";
	    form.submit();	    
}
</script>
</head>

<body marginwidth="0" marginheight="0" leftmargin="0" topmargin="0" leftmargin="0">
<form method=post action='#'>
<input type="hidden" name="act" value="">   
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
                      <td width="75"><img src="images/banner_bg1.gif" width="75" height="17"></td>
                      <td width="350"><font color='#000000' size=4><b> 
                        <center>
                          <%=title%>
                        </center>
                        </b></font> </td>
                      <td width="75"><img src="images/banner_bg1.gif" width="75" height="17"></td>
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
                          <tr class="sbody">
						  <td width='15%' align='left' bgcolor='#D8EFEE'>機構類別</td>
						  <% List bank_type = DBManager.QueryDB_SQLParam("select cmuse_id,cmuse_name from cdshareno where cmuse_div='001' and cmuse_id <> 'Z'order by input_order",null,"");%>
						  <td width='85%' colspan=2 bgcolor='e7e7e7'>
                            <select name='BANK_TYPE' onchange="javascript:getData(this.document.forms[0],'<%=act%>');">                            
                            <%for(int i=0;i<bank_type.size();i++){%>
                            <option value="<%=(String)((DataObject)bank_type.get(i)).getValue("cmuse_id")%>"                            
                            <%if(szbank_type.equals((String)((DataObject)bank_type.get(i)).getValue("cmuse_id"))) out.print("selected");%>
                            ><%=(String)((DataObject)bank_type.get(i)).getValue("cmuse_name")%></option>                            
                            <%}%>
                            </select>
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                            
                            <a href="javascript:doSubmit(this.document.forms[0],'Qry');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_queryb.gif',1)"><img src="images/bt_query.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>                            
                           </td>
                          </td>
                           
                          </tr>     
					      
					      <tr class="sbody">
						  <td width='15%' align='left' bgcolor='#D8EFEE'>總機構代碼</td>
						  <% List tbank_no = (List)request.getAttribute("tbank_no");%>
						  <td width='85%' colspan=2 bgcolor='e7e7e7'>
                            <select name='TBANK_NO'>        
                            
                            <%                            
                            for(int i=0;i<tbank_no.size();i++){%>
                            <option value="<%=(String)((DataObject)tbank_no.get(i)).getValue("bank_no")%>"                            
                            <%if( ((DataObject)tbank_no.get(i)).getValue("bank_no") != null &&  (sztbank_no.equals((String)((DataObject)tbank_no.get(i)).getValue("bank_no")))) out.print("selected");%>
                            >
                            <%if(((DataObject)tbank_no.get(i)).getValue("bank_no") != null) out.print((String)((DataObject)tbank_no.get(i)).getValue("bank_no"));%>                            
                            <%if(((DataObject)tbank_no.get(i)).getValue("bank_name") != null) out.print((String)((DataObject)tbank_no.get(i)).getValue("bank_name"));%>                            
                            </option>                            
                            <%}%>
                            </select>
                          </td>
                          </tr>   
                          </table>      
                      </td>    
                      </tr>
                      <tr><td><table><tr><br></tr></table></td></tr>  
                      
                      <tr> 
                      <td><table width=600 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">                      
                      <%
                      String tmpbank_name="";
                      String usertype="&nbsp;";
                      int i = 0;      
                      String bgcolor="#D3EBE0";       
                      if(WTT01List != null){ %>
                       	  <tr class="sbody" bgcolor="#9AD3D0">
                            <td width="25">&nbsp;</td>           
            				<td width="100">總機構代碼名稱</td>
            				<td width="100">組室名稱</td>
            				<td width="100">帳號</td>
            				<td width="100">密碼</td>
            				<td width="100">類別</td>                           				
					      </tr>   
                   		    <% if(WTT01List.size() == 0){%>
                   			   <tr class="sbody" bgcolor="<%=bgcolor%>">
                   			   <td colspan=11 align=center>無資料可供查詢</td><tr>
                   			   <tr>                   			   
                   			<% }
                    		   while(i < WTT01List.size()){ 
                    		      bgcolor = (i % 2 == 0)?"#e7e7e7":"#D3EBE0";	
                    		      usertype="一般使用者";
                      %>                         	  
                          <tr class="sbody" bgcolor="<%=bgcolor%>">
                            <td width="25"><%=i+1%></td>           
            				<td width="100">
            				<%if( ((DataObject)WTT01List.get(i)).getValue("tbank_no") != null ){                      		
                       			  if(tmpbank_name.equals((String)((DataObject)WTT01List.get(i)).getValue("tbank_no"))){   
                       			     out.print("&nbsp;");
                    		      }else{
                    		        out.print((String)((DataObject)WTT01List.get(i)).getValue("tbank_no"));
                    		        out.print("<br>");
                    		        out.print((String)((DataObject)WTT01List.get(i)).getValue("tbank_name"));
                    		        tmpbank_name = (String)((DataObject)WTT01List.get(i)).getValue("tbank_no");
                    		      }  
                    		  }    
                    		%>    
                    		</td>
            				<td width="100"><%if( ((DataObject)WTT01List.get(i)).getValue("bank_name") != null ) out.print((String)((DataObject)WTT01List.get(i)).getValue("bank_name")); else out.print("&nbsp;");%></td>            				
            				<td width="10"><%if( ((DataObject)WTT01List.get(i)).getValue("muser_id") != null ) out.print((String)((DataObject)WTT01List.get(i)).getValue("muser_id"));%></td>
            				<td width="100"><%if( ((DataObject)WTT01List.get(i)).getValue("muser_password") != null ) out.print((String)((DataObject)WTT01List.get(i)).getValue("muser_password")); else out.print("&nbsp;");%></td>           
            				<td width="100">            				
                    		<%if( ((DataObject)WTT01List.get(i)).getValue("muser_type") != null ){
                    		     if(((String)((DataObject)WTT01List.get(i)).getValue("muser_type")).equals("S")){
                    		        usertype="super user";
                    		     }else{
                    		        usertype="管理者";
                    		     }
                    		  }
                    		  out.print(usertype);
                    		%>                    		
            				</td>               
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
