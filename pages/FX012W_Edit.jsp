<%
//97.09.22 create 簽証會計師申報 by 2295
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
	List BOAF_ACCOUNT= (request.getAttribute("BOAF_ACCOUNT")==null)?null:(List)request.getAttribute("BOAF_ACCOUNT");
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");
	 																										
	Properties permission = ( session.getAttribute("FX012W")==null ) ? new Properties() : (Properties)session.getAttribute("FX012W");
	if(permission == null){
       System.out.println("FX012W_Edit.permission == null");
    }else{
       System.out.println("FX012W_Edit.permission.size ="+permission.size());
    }

	String bank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");
	String nowtbank_no =  ( request.getParameter("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");
	if(!nowtbank_no.equals("")){
	   session.setAttribute("nowtbank_no",nowtbank_no);//將已點選的tbank_no寫入session
	}
	bank_no = ( session.getAttribute("nowtbank_no")==null ) ? bank_no : (String)session.getAttribute("nowtbank_no");
	String m_year="",bank_name="",name="",title="",addr="",telno="";    
        
    String sqlCmd = "";
	DataObject bean = null; 
    if(BOAF_ACCOUNT!=null){
       bean = (DataObject)BOAF_ACCOUNT.get(0);       
       bank_name = bean.getValue("bank_name") == null?"":(String)bean.getValue("bank_name");       
       name = bean.getValue("name") == null?"":(String)bean.getValue("name");
       title = bean.getValue("title") == null?"":(String)bean.getValue("title");
       addr = bean.getValue("addr") == null?"":(String)bean.getValue("addr");
       telno = bean.getValue("telno") == null?"":(String)bean.getValue("telno");
       m_year = bean.getValue("m_year") == null?"":bean.getValue("m_year").toString();
       
   }
	
%>

<script language="javascript" src="js/FX012W.js"></script>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<head>
<title>簽証會計師申報作業</title>
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
<body leftmargin="15" topmargin="0">
<table width="321" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  		<tr>
   		 <td width="618"><img src="images/space_1.gif" width="12" height="12"></td>
  		</tr>
      <tr>
        <td bgcolor="#FFFFFF" width="600">  
           <tr>
                <td width="600" height="18">
                 <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="150"><img src="images/banner_bg1.gif" width="150" height="17" align="left"></td>
                      <td width="300" align="center"><b><font size="4">簽証會計師申報作業</font></b></td>
                      <td width="150" align="right"><img src="images/banner_bg1.gif" width="150" height="17"></td>
                    </tr>
                  </table>
               </td>
              </tr>
              <tr>
              <td width="600" height="50">
                <div align="left">
                <table width="600" border="0"  cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="18" width="600" >               
                      </td>
                    </tr>
                    <tr>
                      
                       <div>
                       <jsp:include page="getLoginUser.jsp?width=600" flush="true" />             			
                       </div>			 							
			 		   
                    </tr>
                <tr>
                <td width="600" height="200">
                <form method="post" name="outpushdata">
                <table width="600" border="0" align="center" cellpadding="0" cellspacing="0" height="2">
                    <tr>
                      <td  class="sbody" width="600" height="76">
                        <table height="129"  bordercolor="#3A9D99" border="1" width="600">                          
                            <tr class="sbody">
							<td align="left" width="87" bgColor="#d8efee" height="23">年度</td>
							<td width="497" bgcolor="#e7e7e7" height="23">
							<%if(act.equals("Edit")){%>
							 <%=m_year%>
							  <input type="hidden" name="m_year" value="<%=m_year%>">
							<%}else{%>
  							  <input type="text" name="m_year" value='<%=m_year%>' size='3'>
  							<%}%>	
  							年   							 
  							</td>  								  							
							</tr>
														
							
                            
                            <tr class="sbody">
                            <td align="left" width="87" bgColor="#d8efee" height="23">會計師姓名
                            </td>
                            <td width="497" bgColor="#e7e7e7" height="23">                            
                             <input type="text" name="name" value='<%=name%>' size='12'>
                            </td>   
                            </tr>                          
                            
                            <tr class="sbody">
                            <td align="left" width="87" bgColor="#d8efee" height="23">事務所名稱
                            </td>
                            <td width="497" bgColor="#e7e7e7" height="23">                            
                            <input type="text" name="title" value='<%=title%>' size='50'>
                            </td>   
                            </tr>     
                            
                            <tr class="sbody">
                            <td align="left" width="87" bgColor="#d8efee" height="23" >地址
                            </td>
                            <td width="497" bgColor="#e7e7e7" height="23">                            
                            <input type="text" name="addr" value='<%=addr%>' size='70' maxlength='80'>
                            </td>   
                            </tr>    
                            
                            <tr class="sbody">
                            <td align="left" width="87" bgColor="#d8efee" height="23" >電話
                            </td>
                            <td width="497" bgColor="#e7e7e7" height="23">                            
                            <input type="text" name="telno" value='<%=telno%>' size='30'>
                            </td>   
                            </tr>    
                        </table>
                      </td>
                    </tr>
                    <td width="600" height="21"><div align="right"><div align="right">
                    <div align="right"><jsp:include page="getMaintainUser.jsp" flush="true" /></div>
                 </td>
             </table>
             </td></tr>
             <tr>
                <td width="660" height="123"><table width="591" border="0" cellpadding="1" cellspacing="1" class="sbody" height="176">
                    <tr>
                      <td colspan="2" width="583" height="41">
                      <div align="center">
                    <table width="243" border="0" cellpadding="1" cellspacing="1">
                      <tr>
                       <%if(act.equals("new")){
                       if(permission != null && permission.get("A") != null && permission.get("A").equals("Y")){//Add
                      %>
				        <td width="66"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'add','<%=bank_no%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_confirmb.gif',1)"><img src="images/bt_confirm.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a></div></td>
				        <% } }else{
				        if(permission != null && permission.get("U") != null && permission.get("U").equals("Y")){//Update %>
				        <td width="66"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'modify','<%=bank_no%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_updateb.gif',1)"><img src="images/bt_update.gif" name="Image103" width="66" height="25" border="0" id="Image103"></a></div></td>
				        <td width="66"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'delete');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image104','','images/bt_deleteb.gif',1)"><img src="images/bt_delete.gif" name="Image104" width="66" height="25" border="0" id="Image104"></a></div></td>
				         <% }}%>
                        <td width="66"><div align="center"><a href="javascript:AskReset(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_cancelb.gif',1)"><img src="images/bt_cancel.gif" name="Image105" width="66" height="25" border="0" id="Image105"></a></div></td>
                        <td width="93"><div align="center"><a href="javascript:history.back();"  onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image106','','images/bt_backb.gif',1)"><img src="images/bt_back.gif" name="Image106" width="80" height="25" border="0" id="Image106"></a></div></td>                        
                      </tr>
                    </table>
                  </div>
                    </td>
                    </tr>
                    <tr>
                      <td colspan="2" width="583" height="41"><font color='#990000'><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle"><font color="#007D7D" size="3">使用說明      
                        :</font></font> </td>
                    </tr>
                    <tr>
                      <td width="16" height="127">&nbsp;</td>
                      <td width="561" height="127">
                      <ul>
                      	  <li class="sbody" >確認輸入資料無誤後, 按<font color="#666666">【確定】即將本表上的資料, 於資料庫中建檔。</li>      
                      	  <li class="sbody" >按<font color="#666666">【修改】即修改的資料,寫入資料庫料庫中。</li>
                          <li class="sbody" >欲重新輸入資料, 按<font color="#666666">【取消】即將本表上的資料清空</li>      
                          <li class="sbody" >如放棄修改或無修改之資料需輸入, 按【回上一頁】]即離開本程式。</li>                                
                          
                        </ul>
                            </font>
                            </font>
                          </font>
                       </font>
                      </td>
                    </tr>
                  </table></td>
              </tr>              
</table>
</form>
</table>
</body>
