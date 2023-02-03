<%
//93.12.29 create by 2295
//94.02.01 fix sub_type/order_type為MIS系統專用
//94.02.03 fix 區分網際網路申報/MIS管理系統的配色 by 2295
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
	String title = (act.equals("new"))?"新增建檔":"異動維護建檔";
	
	
	List WTT03_1 = (List)request.getAttribute("WTT03_1");		
	if(WTT03_1 == null){
	   System.out.println("WTT03_1 == null");
	}else{
	   System.out.println("WTT03_1.size()="+WTT03_1.size());
	}
	
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
<title>「程式功能維護」<%=title%></title>
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
<table width="640" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  		<tr> 
   		 <td><img src="images/space_1.gif" width="12" height="12"></td>
  		</tr>
          <td bgcolor="#FFFFFF">
			<table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="110"><img src="images/banner_bg1.gif" width="110" height="17"></td>
                      <td width="380"><font color='#000000' size=4><b> 
                        <center>
                          「程式功能維護」<%=title%> 
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
						  <td width='15%' class="<%=nameColor%>">程式代碼</td>						  
						  <td width='85%' colspan=3 class="<%=textColor%>">
						    <input type='text' name='PROGRAM_ID' value="<%if(WTT03_1 != null && WTT03_1.size() != 0 && ((DataObject)WTT03_1.get(0)).getValue("program_id") != null ) out.print((String)((DataObject)WTT03_1.get(0)).getValue("program_id"));%>" size='20' maxlength='20' <%if(act.equals("Edit")){out.print("disabled");}%>   >                            
						    <%if(act.equals("Edit")){%>
						    <input type='hidden' name='program_id' value="<%if(WTT03_1 != null && WTT03_1.size() != 0 && ((DataObject)WTT03_1.get(0)).getValue("program_id") != null ) out.print((String)((DataObject)WTT03_1.get(0)).getValue("program_id"));%>" size='20'%>                             
						    <%}%>
						    <font color='red'>*</font>    
                          </td>
                         </tr>                              
                         
                         <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">程式名稱</td>						  
						  <td width='85%' colspan=3 class="<%=textColor%>">
						    <input type='text' name='PROGRAM_NAME' value="<%if(WTT03_1 != null && WTT03_1.size() != 0 && ((DataObject)WTT03_1.get(0)).getValue("program_name") != null ) out.print((String)((DataObject)WTT03_1.get(0)).getValue("program_name"));%>" size='60' maxlength="100">                            
						    <font color='red'>*</font>    
                          </td>
                         </tr>                              
                         
                         <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">程式URL</td>						  
						  <td width='85%' colspan=3 class="<%=textColor%>"'>
						    <input type='text' name='URL_ID' value="<%if(WTT03_1 != null && WTT03_1.size() != 0 && ((DataObject)WTT03_1.get(0)).getValue("url_id") != null ) out.print((String)((DataObject)WTT03_1.get(0)).getValue("url_id"));%>" size='60' maxlength="80">                            
						    <font color='red'>*</font>    
                          </td>
                         </tr> 
                         
                         <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">新增</td>						  
						  <td width='85%' colspan=3 class="<%=textColor%>">
						    <input type='radio' name='P_ADD' value="Y" <%if(WTT03_1 != null && WTT03_1.size() != 0 && (((DataObject)WTT03_1.get(0)).getValue("p_add") != null  && ((String)((DataObject)WTT03_1.get(0)).getValue("p_add")).equals("Y") )) out.print("checked");%>>Y                            
						    <input type='radio' name='P_ADD' value="N" <%if ((WTT03_1 == null) || (WTT03_1 != null && WTT03_1.size() != 0 && (((DataObject)WTT03_1.get(0)).getValue("p_add") != null  && ((String)((DataObject)WTT03_1.get(0)).getValue("p_add")).equals("N") ))) out.print("checked");%>>N                            
                          </td>
                         </tr>     
                         
                         
                         <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">刪除</td>						  
						  <td width='85%' colspan=3 class="<%=textColor%>">
						    <input type='radio' name='P_DELETE' value="Y" <%if(WTT03_1 != null && WTT03_1.size() != 0 && (((DataObject)WTT03_1.get(0)).getValue("p_delete") != null  && ((String)((DataObject)WTT03_1.get(0)).getValue("p_delete")).equals("Y") )) out.print("checked");%>>Y                            
						    <input type='radio' name='P_DELETE' value="N" <%if ((WTT03_1 == null) || (WTT03_1 != null && WTT03_1.size() != 0 && (((DataObject)WTT03_1.get(0)).getValue("p_delete") != null  && ((String)((DataObject)WTT03_1.get(0)).getValue("p_delete")).equals("N") ))) out.print("checked");%>>N                                                     
                          </td>
                         </tr>     
                         
                         <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">修改</td>						  
						  <td width='85%' colspan=3 class="<%=textColor%>">
						    <input type='radio' name='P_UPDATE' value="Y" <%if(WTT03_1 != null && WTT03_1.size() != 0 && (((DataObject)WTT03_1.get(0)).getValue("p_update") != null  && ((String)((DataObject)WTT03_1.get(0)).getValue("p_update")).equals("Y") )) out.print("checked");%>>Y                            
						    <input type='radio' name='P_UPDATE' value="N" <%if ((WTT03_1 == null) || (WTT03_1 != null && WTT03_1.size() != 0 && (((DataObject)WTT03_1.get(0)).getValue("p_update") != null  && ((String)((DataObject)WTT03_1.get(0)).getValue("p_update")).equals("N") ))) out.print("checked");%>>N                                                      
                          </td>
                         </tr>     
                         
                         <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">查詢</td>						  
						  <td width='85%' colspan=3 class="<%=textColor%>">
						    <input type='radio' name='P_QUERY' value="Y" <%if(WTT03_1 != null && WTT03_1.size() != 0 && (((DataObject)WTT03_1.get(0)).getValue("p_query") != null  && ((String)((DataObject)WTT03_1.get(0)).getValue("p_query")).equals("Y") )) out.print("checked");%>>Y                            
						    <input type='radio' name='P_QUERY' value="N" <%if ((WTT03_1 == null) || (WTT03_1 != null && WTT03_1.size() != 0 && (((DataObject)WTT03_1.get(0)).getValue("p_query") != null  && ((String)((DataObject)WTT03_1.get(0)).getValue("p_query")).equals("N") ))) out.print("checked");%>>N                                                   
                          </td>
                         </tr>     
                         
                         <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">印表</td>						  
						  <td width='85%' colspan=3 class="<%=textColor%>">
						    <input type='radio' name='P_PRINT' value="Y" <%if(WTT03_1 != null && WTT03_1.size() != 0 && (((DataObject)WTT03_1.get(0)).getValue("p_print") != null  && ((String)((DataObject)WTT03_1.get(0)).getValue("p_print")).equals("Y") )) out.print("checked");%>>Y                            
						    <input type='radio' name='P_PRINT' value="N" <%if ((WTT03_1 == null) || (WTT03_1 != null && WTT03_1.size() != 0 && (((DataObject)WTT03_1.get(0)).getValue("p_print") != null  && ((String)((DataObject)WTT03_1.get(0)).getValue("p_print")).equals("N") ))) out.print("checked");%>>N                                                      
                          </td>
                         </tr>     
                         
                         <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">上傳</td>						  
						  <td width='85%' colspan=3 class="<%=textColor%>"'>
						    <input type='radio' name='P_UPLOAD' value="Y" <%if(WTT03_1 != null && WTT03_1.size() != 0 && (((DataObject)WTT03_1.get(0)).getValue("p_upload") != null  && ((String)((DataObject)WTT03_1.get(0)).getValue("p_upload")).equals("Y") )) out.print("checked");%>>Y                           
						    <input type='radio' name='P_UPLOAD' value="N" <%if ((WTT03_1 == null) || (WTT03_1 != null && WTT03_1.size() != 0 && (((DataObject)WTT03_1.get(0)).getValue("p_upload") != null  && ((String)((DataObject)WTT03_1.get(0)).getValue("p_upload")).equals("N") ))) out.print("checked");%>>N                                                      
                          </td>
                         </tr>     
						
						 <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">下載</td>						  
						  <td width='85%' colspan=3 class="<%=textColor%>">
						    <input type='radio' name='P_DOWNLOAD' value="Y" <%if(WTT03_1 != null && WTT03_1.size() != 0 && (((DataObject)WTT03_1.get(0)).getValue("p_download") != null  && ((String)((DataObject)WTT03_1.get(0)).getValue("p_download")).equals("Y") )) out.print("checked");%>>Y                            
						    <input type='radio' name='P_DOWNLOAD' value="N" <%if ((WTT03_1 == null) || (WTT03_1 != null && WTT03_1.size() != 0 && (((DataObject)WTT03_1.get(0)).getValue("p_download") != null  && ((String)((DataObject)WTT03_1.get(0)).getValue("p_download")).equals("N") ))) out.print("checked");%>>N                                                      
                          </td>
                         </tr>     
                         
                         <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">鎖定</td>						  
						  <td width='85%' colspan=3 class="<%=textColor%>">
						    <input type='radio' name='P_LOCK' value="Y" <%if(WTT03_1 != null && WTT03_1.size() != 0 && (((DataObject)WTT03_1.get(0)).getValue("p_lock") != null  && ((String)((DataObject)WTT03_1.get(0)).getValue("p_lock")).equals("Y") )) out.print("checked");%>>Y                            
						    <input type='radio' name='P_LOCK' value="N" <%if ((WTT03_1 == null) || (WTT03_1 != null && WTT03_1.size() != 0 && (((DataObject)WTT03_1.get(0)).getValue("p_lock") != null  && ((String)((DataObject)WTT03_1.get(0)).getValue("p_lock")).equals("N") ))) out.print("checked");%>>N                                                      
                          </td>
                         </tr>     
                         
                         <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>"'>申報</td>						  
						  <td width='85%' colspan=3 class="<%=textColor%>">
						    <input type='radio' name='P_OTHER' value="Y" <%if(WTT03_1 != null && WTT03_1.size() != 0 && (((DataObject)WTT03_1.get(0)).getValue("p_other") != null  && ((String)((DataObject)WTT03_1.get(0)).getValue("p_other")).equals("Y") )) out.print("checked");%>>Y                           
						    <input type='radio' name='P_OTHER' value="N" <%if ((WTT03_1 == null) || (WTT03_1 != null && WTT03_1.size() != 0 && (((DataObject)WTT03_1.get(0)).getValue("p_other") != null  && ((String)((DataObject)WTT03_1.get(0)).getValue("p_other")).equals("N") ))) out.print("checked");%>>N                                                       
                          </td>
                         </tr>                              
                         
                         
                         <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">資料順序</td>						  
						  <td width='85%' colspan=3 class="<%=textColor%>">
						    <input type='text' name='INPUT_ORDER' value="<%if(WTT03_1 != null && WTT03_1.size() != 0 && ((DataObject)WTT03_1.get(0)).getValue("input_order") != null ) out.print((String)((DataObject)WTT03_1.get(0)).getValue("input_order"));%>" size='3' maxlength='3'>                            
						    <font color='red'>*</font>    
                          </td>
                         </tr>     
                         
                         <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">次分類</td>						  
						  <td width='85%' colspan=3 class="<%=textColor%>">
						    <input type='text' name='SUB_TYPE' value="<%if(WTT03_1 != null && WTT03_1.size() != 0 && ((DataObject)WTT03_1.get(0)).getValue("sub_type") != null ) out.print((String)((DataObject)WTT03_1.get(0)).getValue("sub_type"));%>" size='3' maxlength='3'>                            						    
						    (檢查追蹤管理系統專用)
                          </td>
                         </tr>     
                         
                         <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">排序</td>						  
						  <td width='85%' colspan=3 class="<%=textColor%>">
						    <input type='text' name='ORDER_TYPE' value="<%if(WTT03_1 != null && WTT03_1.size() != 0 && ((DataObject)WTT03_1.get(0)).getValue("order_type") != null ) out.print((String)((DataObject)WTT03_1.get(0)).getValue("order_type"));%>" size='3' maxlength='3'>                            
						    (檢查追蹤管理系統專用)
                          </td>
                         </tr>     
                          
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
                        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Insert');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_confirmb.gif',1)"><img src="images/bt_confirm.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a></div></td>
                        <td width="66"> <div align="center"><a href="javascript:AskReset(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_cancelb.gif',1)"><img src="images/bt_cancel.gif" name="Image103" width="66" height="25" border="0" id="Image103"></a></div></td>		      
                            <%}%>
         				<%}%>
         				<%if(act.equals("Edit")){%>         				
         				    <%if(permission != null && permission.get("U") != null && permission.get("U").equals("Y")){ %>                   	        	                                   		              				
				        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Update');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image102','','images/bt_updateb.gif',1)"><img src="images/bt_update.gif" name="Image102" width="66" height="25" border="0" id="Image102"></a></div></td>
				        <td width="66"> <div align="center"><a href="javascript:AskReset(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_cancelb.gif',1)"><img src="images/bt_cancel.gif" name="Image103" width="66" height="25" border="0" id="Image103"></a></div></td>		      
				            <%}%>				           
						<%}%>			                    
                       
                       <td width="80"><div align="center"><a href="javascript:history.back();"  onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image104','','images/bt_backb.gif',1)"><img src="images/bt_back.gif" name="Image104" width="80" height="25" border="0" id="Image104"></a></div></td>
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
                          <li>本網頁提供程式功能維護功能。</li> 
					      <li>新增時,按<font color="#666666">【確定】</font>即將資料寫入資料庫。</li> 
					      <li>修改時,按<font color="#666666">【修改】</font>即將資料寫入資料庫。</li> 
					      <li>按<font color="#666666">【取消】</font>放棄資料修改。</li>
                          <li>按<font color="#666666">【回上一頁】則放棄資料維護, 回至上個畫面</font>。</li>
                          <li>【<font color="red">*</font>】為必填欄位。</li>
                        </ul></td>
                    </tr>
                  </table></td>
              </tr>             
</form>
</body>
</html>
