<%
//94.01.05 fix 農金局以外使用者,只能輸入3碼...其帳號為機構代號+3碼
//94.02.03 fix 區分網際網路申報/MIS管理系統的配色 by 2295
//94.03.24 fix 鎖定改成暫停 by 2295
//94.04.07 fix 姓名改12 byte by 2295
//94.04.08 add 農金局系統管理者前7碼必須與總機構代碼相同 by 2295
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
	
	//String muser_type = ( request.getParameter("muser_type")==null ) ? "S" : (String)request.getParameter("muser_type");		
	String szbank_type = ( request.getParameter("bank_type")==null ) ? "" : (String)request.getParameter("bank_type");		
	String sztbank_no = ( request.getParameter("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");		
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");				
	String muser_type = (session.getAttribute("muser_type") == null)?"":(String)session.getAttribute("muser_type");			
		
	System.out.println("ZZ001W_Edit.sztbank_no="+sztbank_no);
	
	List WTT01 = (List)request.getAttribute("WTT01");		
	String title="「使用者帳號維護」";		
	title =(muser_type.equals("S"))?"系統管理者"+title:title;
	title =(act.equals("Edit"))?title+"異動維護建檔":title;
	title =(act.equals("new"))?title+"新增建檔":title;
	if(muser_type.equals("A")){
	   szbank_type = (session.getAttribute("bank_type") == null)?"":(String)session.getAttribute("bank_type");		   
	   sztbank_no = (session.getAttribute("tbank_no") == null)?"":(String)session.getAttribute("tbank_no");		   
	   System.out.println("szbank_type="+szbank_type);
	   System.out.println("szbank_no="+sztbank_no);
	}
	
	//取得ZZ001W的權限
	Properties permission = ( session.getAttribute("ZZ001W")==null ) ? new Properties() : (Properties)session.getAttribute("ZZ001W"); 
	if(permission == null){
       System.out.println("ZZ001W_Edit.permission == null");
    }else{
       System.out.println("ZZ001W_Edit.permission.size ="+permission.size());
               
    }	
	System.out.println("ZZ001W_Edit.szbank_no="+sztbank_no);
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/ZZ001W.js"></script>
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
<input type="hidden" name="lguser_id" value="<%=(String)session.getAttribute("muser_id")%>">  
<input type="hidden" name="lguser_type" value="<%=(String)session.getAttribute("muser_type")%>">  
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
                      <td width="110"><img src="images/banner_bg1.gif" width="110" height="17"></td>
                      <td width="380"><font color='#000000' size=4><b> 
                        <center>
                          <%=title%> 
                        </center>
                        </b></font> </td>
                      <td width="110"><img src="images/banner_bg1.gif" width="110" height="17"></td>
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
                    <%
                      String nameColor="nameColor_sbody";
                      String textColor="textColor_sbody";
                      String bordercolor="#3A9D99";
                    %>                    
                    <tr> 
                      <td><table width=600 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="<%=bordercolor%>">
                          <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">機構類別</td>
						  <% List bank_type = DBManager.QueryDB_SQLParam("select cmuse_id,cmuse_name from cdshareno where cmuse_div='001' and cmuse_id <> 'Z'order by input_order",null,"");%>
						  <td width='85%' colspan=2 class="<%=textColor%>">
                            <select name='BANK_TYPE' onchange="javascript:getData(this.document.forms[0],'<%=act%>','bank_type');"
                             <%if((!muser_type.equals("S")) || act.equals("Edit")) out.print("disabled");%>                             
                            >                            
                            <%for(int i=0;i<bank_type.size();i++){%>
                            <option value="<%=(String)((DataObject)bank_type.get(i)).getValue("cmuse_id")%>"
                            <%if((WTT01 != null && WTT01.size() != 0) && ( ((DataObject)WTT01.get(0)).getValue("bank_type") != null && ((String)((DataObject)WTT01.get(0)).getValue("bank_type")).equals((String)((DataObject)bank_type.get(i)).getValue("cmuse_id")))) out.print("selected");%>
                            <%if(szbank_type.equals((String)((DataObject)bank_type.get(i)).getValue("cmuse_id"))) out.print("selected");%>
                            ><%=(String)((DataObject)bank_type.get(i)).getValue("cmuse_name")%></option>                            
                            <%}%>
                            </select>
                          </td>
                          </tr>     
					      
					      <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">總機構代碼</td>
						  <% List tbank_no = (List)request.getAttribute("tbank_no");%>
						  <td width='85%' colspan=2 class="<%=textColor%>">
                            <select name='TBANK_NO' <%if(szbank_type.equals("2")){%>onchange="javascript:getData(this.document.forms[0],'<%=act%>','tbank_no');"<%}%>                           
                            <%if((!muser_type.equals("S")) || act.equals("Edit")) out.print("disabled");%>                             
                            >                            
                            <%for(int i=0;i<tbank_no.size();i++){%>
                            <option value="<%=(String)((DataObject)tbank_no.get(i)).getValue("bank_no")%>"
                            <%if((WTT01 != null && WTT01.size() != 0) && ( ((DataObject)WTT01.get(0)).getValue("tbank_no") != null && ((String)((DataObject)WTT01.get(0)).getValue("tbank_no")).equals((String)((DataObject)tbank_no.get(i)).getValue("bank_no")))) out.print("selected");%>
                            <%if(sztbank_no.equals((String)((DataObject)tbank_no.get(i)).getValue("bank_no"))) out.print(" selected");%>
                            ><%=(String)((DataObject)tbank_no.get(i)).getValue("bank_no")%>&nbsp;<%=(String)((DataObject)tbank_no.get(i)).getValue("bank_name")%></option>                            
                            <%}%>
                            </select>
                          </td>
                          </tr>   
                          
                          <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">組室代碼</td>
						  <% List bank_no = (List)request.getAttribute("bank_no");%>
						  <td width='85%' colspan=2 class="<%=textColor%>">
                            <select name='BANK_NO'>  
                            <%if(szbank_type.equals("2")){ //農金局局內專用                                                    
                                for(int i=0;i<bank_no.size();i++){%>
                                <option value="<%=(String)((DataObject)bank_no.get(i)).getValue("bank_no")%>"
                                <%if((WTT01 != null && WTT01.size() != 0) && ( ((DataObject)WTT01.get(0)).getValue("bank_no") != null && ((String)((DataObject)WTT01.get(0)).getValue("bank_no")).equals((String)((DataObject)bank_no.get(i)).getValue("bank_no")))) out.print("selected");%>
                                ><%=(String)((DataObject)bank_no.get(i)).getValue("bank_name")%></option>                            
                            <%  }//end of for
                              }else{//非農金局局內 
                            %>  
                              <option value="">&nbsp;&nbsp;</option> 
                            <%}//end of if%>                            
                            </select>
                             (農金局局內專用)
                          </td>
                          </tr>					      
					      
					      <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">科別</td>
						  <% List subdep_id = DBManager.QueryDB_SQLParam("select cmuse_id,cmuse_name from cdshareno where cmuse_div='010' order by input_order",null,"");%>
						  <td width='85%' colspan=2 class="<%=textColor%>">
                            <select name='SUBDEP_ID'>                            
                            <%if(szbank_type.equals("2")){ //農金局局內專用                                                                                
                                for(int i=0;i<subdep_id.size();i++){%>
                                <option value="<%=(String)((DataObject)subdep_id.get(i)).getValue("cmuse_id")%>"
                                <%if((WTT01 != null && WTT01.size() != 0) && ( ((DataObject)WTT01.get(0)).getValue("subdep_id") != null && ((String)((DataObject)WTT01.get(0)).getValue("subdep_id")).equals((String)((DataObject)subdep_id.get(i)).getValue("cmuse_id")))) out.print("selected");%>
                                ><%=(String)((DataObject)subdep_id.get(i)).getValue("cmuse_name")%></option>                            
                            <%  }//end of for
                              }else{//非農金局局內
                            %>
                                <option value="">&nbsp;&nbsp;</option> 
                            <%}//end of if%>                                
                            </select>
                             (農金局局內專用)
                          </td>
                          </tr>     					      
					     
					      <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">使用者帳號</td>
						  <td width='85%' colspan=2 class="<%=textColor%>">                            
  		                    <%if(act.equals("Edit")){%>
  		                     <input type='hidden' name='MUSER_ID' value="<%if(WTT01 != null && WTT01.size() != 0 && ((DataObject)WTT01.get(0)).getValue("muser_id") != null ) out.print((String)((DataObject)WTT01.get(0)).getValue("muser_id"));%>">
  		                     <%if(WTT01 != null && WTT01.size() != 0 && ((DataObject)WTT01.get(0)).getValue("muser_id") != null ) out.print((String)((DataObject)WTT01.get(0)).getValue("muser_id"));%>
  		                    <%}else{%>
  		                    <input type='text' name='MUSER_ID' value="<%if(WTT01 != null && WTT01.size() != 0 && ((DataObject)WTT01.get(0)).getValue("muser_id") != null ) out.print((String)((DataObject)WTT01.get(0)).getValue("muser_id"));%>" size='12' maxlength='12'>
  		                    <%}%>           
  		                    <font color='red' size=4>*</font>                                           
                          </td>
                          </tr>
					     
					      <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">使用者姓名</td>
						  <td width='85%' colspan=2 class="<%=textColor%>">
                            <input type='text' name='MUSER_NAME' value="<%if(WTT01 != null && WTT01.size() != 0 && ((DataObject)WTT01.get(0)).getValue("muser_name") != null ) out.print((String)((DataObject)WTT01.get(0)).getValue("muser_name"));%>" size='12' maxlength='12' >                          
                            <font color='red' size=4>*</font>
                          </td>
                          </tr>
                          <%if(act.equals("Edit")){%>
                          <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">暫停與否</td>
						  <td width='85%' colspan=2 class="<%=textColor%>">
						    <select name=LOCK_MARK >
                            <option value="Y" <%if((WTT01 != null && WTT01.size() != 0) &&  (((DataObject)WTT01.get(0)).getValue("lock_mark") != null && ((String)((DataObject)WTT01.get(0)).getValue("lock_mark")).equals("Y"))  ) out.print("selected");%>>Y</option>
                            <option value="N" <%if((WTT01 == null) || ((WTT01 != null && WTT01.size() != 0) &&  ((((DataObject)WTT01.get(0)).getValue("lock_mark") == null) || ((String)((DataObject)WTT01.get(0)).getValue("lock_mark")).equals("N")))   ) out.print("selected");%>>N</option>                           
                            </select>
                          </td>
                          </tr>                    
                          <%}%>      
                        </Table></td>
                    </tr>                 
                    <tr>                  
                <td><div align="right"><jsp:include page="getMaintainUser.jsp" flush="true" /></div></td>                                              
              </tr>
              
              <tr> 
                <td><div align="center"> 
                    <table width="243" border="0" cellpadding="1" cellspacing="1">
                      <tr>                             
				        <%if(act.equals("new")){%>       
				            <%if(permission != null && permission.get("A") != null && permission.get("A").equals("Y")){ %>                   	        	                                   		     
                        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Insert','');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_confirmb.gif',1)"><img src="images/bt_confirm.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a></div></td>
                            <%}%>
         				<%}%>
         				<%if(act.equals("Edit")){%>         				
         				    <%if(permission != null && permission.get("U") != null && permission.get("U").equals("Y")){ %>                   	        	                                   		     
         				<td width="80"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'ResetPwd','');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image106','','images/bt_resetpwdb.gif',1)"><img src="images/bt_resetpwd.gif" name="Image106" width="80" height="25" border="0" id="Image106"></a></div></td>
				        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Update','');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image102','','images/bt_updateb.gif',1)"><img src="images/bt_update.gif" name="Image102" width="66" height="25" border="0" id="Image102"></a></div></td>
				            <%}%>
				            <%if(permission != null && permission.get("D") != null && permission.get("D").equals("Y")){ %>                   	        	                                   		     
				        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Delete','');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_deleteb.gif',1)"><img src="images/bt_delete.gif" name="Image103" width="66" height="25" border="0" id="Image103"></a></div></td>						
				            <%}%>
						<%}%>				
         				<%if(!act.equals("Query")){%>       
                        <td width="66"> <div align="center"><a href="javascript:AskReset(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image104','','images/bt_cancelb.gif',1)"><img src="images/bt_cancel.gif" name="Image104" width="66" height="25" border="0" id="Image104"></a></div></td>
		                <%}%>        
                        <td width="80"><div align="center"><a href="javascript:history.back();"  onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_backb.gif',1)"><img src="images/bt_back.gif" name="Image105" width="80" height="25" border="0" id="Image105"></a></div></td>
                      </tr>
                      </tr>
                    </table>
                  </div></td>
              </tr>                               
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
                          <li>本網頁提供新增使用者帳號功能。</li>
                          <li>新增的帳號其密碼將和帳號相同。</li>                          
                          <li>農金局局內使用者帳號，須填入10碼國民身份證號，並填入使用者姓名。</li>      
                          <li>農金局<font color="red">局外使用者帳號</font>，須填入<font color="red">任意3碼</font>，並填入使用者姓名。</li>
                          <li><font color="red">實際核發的使用者帳號為總機構代碼+任意3碼</font>。</li>      
                          <li>新增時,可直接於空格內更改資料，資料更改完畢後，按"確定"即將本表上的資料於資料庫中建檔。</li>
                          <li>修改時,資料更改完畢後，按<font color="#666666">【修改】</font>即將本表上的資料於資料庫中建檔。</li>                        
                          <li>按<font color="#666666">【取消】</font>即重新輸入資料。</li>                         
                          <li>按<font color="#666666">【回上一頁】則放棄新增帳號, 回至上個畫面</font>。</li>
                          <li>【<font color="red" size=4>*</font>】為必填欄位。</li> 
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
