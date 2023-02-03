<%
// 94.01.12 fix 機構名稱欄位加長至100 by 2295
// 94.02.03 fix 區分網際網路申報/MIS管理系統的配色 by 2295
// 99.10.14 fix 取得該功能項目權限移至Utility.getPermission by 2295
//102.11.05 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%
	Map dataMap =Utility.saveSearchParameter(request);
	String report_no = "AS002W";
	String act = Utility.getTrimString(dataMap.get("act"));	
	String szbank_type = Utility.getTrimString(dataMap.get("bank_type"));
	String sztbank_no = Utility.getTrimString(dataMap.get("tbank_no"));	
	
	List BN02 = (List)request.getAttribute("BN02");		
	String title="「分支機構代碼」";			
	title =(act.equals("Edit"))?title+"異動維護建檔":title;
	title =(act.equals("new"))?title+"新增建檔":title;	
	
    
    List tbankList = (List)request.getAttribute("tbank_no");
	System.out.println("test="+tbankList.size());
    // XML Ducument for 總機構代碼 begin
    out.println("<xml version=\"1.0\" encoding=\"UTF-8\" ID=\"TBankXML\">");
    out.println("<datalist>");
    for(int i=0;i< tbankList.size(); i++) {
        DataObject bean =(DataObject)tbankList.get(i);
        out.println("<data>");
        out.println("<bankType>"+bean.getValue("bank_type")+"</bankType>");
        out.println("<bankValue>"+bean.getValue("bank_no")+"</bankValue>");
        out.println("<bankName>"+bean.getValue("bank_no")+"  "+bean.getValue("bank_name")+"</bankName>");
        out.println("</data>");
    }
    out.println("</datalist>\n</xml>");
    // XML Ducument for 總機構代碼 end 		
	
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/AS002W.js"></script>
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
<table width="600" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  		<tr> 
   		 <td><img src="images/space_1.gif" width="12" height="12"></td>
  		</tr>
        <tr> 
          <td bgcolor="#FFFFFF">
			<table width="560" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="560" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="140"><img src="images/banner_bg1.gif" width="140" height="17"></td>
                      <td width="320"><font color='#000000' size=4><b> 
                        <center>
                          <%=title%> 
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
                <td><table width="560" border="0" align="center" cellpadding="0" cellspacing="0">
               
                    <tr> 
                      <div align="right"><jsp:include page="getLoginUser.jsp?width=560" flush="true" /></div> 
                    </tr>          
                    <%
                      String nameColor="nameColor_sbody";
                      String textColor="textColor_sbody";
                      String bordercolor="#3A9D99";
                    %>         
                    <tr> 
                      <td><table width=560 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="<%=bordercolor%>">
                      	  <td width='15%' class="<%=nameColor%>">機構類別</td>
						  <% List bank_type = DBManager.QueryDB_SQLParam("select cmuse_id,cmuse_name from cdshareno where cmuse_div='001' and cmuse_id <> 'Z' order by input_order",null,"");%>
						  <td width='85%' colspan=2 class="<%=textColor%>">
                            <select name='BANK_TYPE' onchange="javascript:changeOption(this.document.forms[0]);"
                            <%if(act.equals("Edit")) out.print("disabled");%>  
                            >                                                        
                            <%for(int i=0;i<bank_type.size();i++){%>
                            <option value="<%=(String)((DataObject)bank_type.get(i)).getValue("cmuse_id")%>" ><%=(String)((DataObject)bank_type.get(i)).getValue("cmuse_name")%></option>                            
                            <%}%>
                            </select>                            
                           </td>
                          </td>
                           
                          </tr>     
					      
					      <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">總機構代碼</td>
						  <td width='85%' colspan=2 class="<%=textColor%>">
                            <select name='TBANK_NO' <%if(act.equals("Edit")) out.print("disabled");%> >                            
                            
                            </select>
                          </td>
                          </tr>   
					      
					      <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">分支機構代碼</td>						  
						  <td width='85%' colspan=2 class="<%=textColor%>">
						    <%if(act.equals("Edit")){%>
  		                     <input type='hidden' name='BANK_NO' value="<%if(BN02 != null && BN02.size() != 0 && ((DataObject)BN02.get(0)).getValue("bank_no") != null ) out.print((String)((DataObject)BN02.get(0)).getValue("bank_no"));%>">
  		                     <%if(BN02 != null && BN02.size() != 0 && ((DataObject)BN02.get(0)).getValue("bank_no") != null ) out.print((String)((DataObject)BN02.get(0)).getValue("bank_no"));%>
  		                    <%}else{%>
  		                    <input type='text' name='BANK_NO' value="<%if(BN02 != null && BN02.size() != 0 && ((DataObject)BN02.get(0)).getValue("bank_no") != null ) out.print((String)((DataObject)BN02.get(0)).getValue("bank_no"));%>" size='7' maxlength='7'>
  		                    <%}%>  
                            </select>
                            <font color='red'>*</font>  
                          </td>
                          </tr>   		      
					     
					      <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">機構名稱</td>
						  <td width='85%' colspan=2 class="<%=textColor%>">                              		                    
  		                    <input type='text' name='BANK_NAME' value="<%if(BN02 != null && BN02.size() != 0 && ((DataObject)BN02.get(0)).getValue("bank_name") != null ) out.print((String)((DataObject)BN02.get(0)).getValue("bank_name"));%>" size='50' maxlength='100'>  		                    
  		                    <font color='red'>*</font>  
                          </td>
                          </tr>
					     
					      <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">機構簡稱</td>
						  <td width='85%' colspan=2 class="<%=textColor%>">
                            <input type='text' name='BANK_B_NAME' value="<%if(BN02 != null && BN02.size() != 0 && ((DataObject)BN02.get(0)).getValue("bank_b_name") != null ) out.print((String)((DataObject)BN02.get(0)).getValue("bank_b_name"));%>" size='50' maxlength='100' >                          
                            <input type='button' value='同上' onClick="javascript:sameBank_Name(form);">
                            <font color='red'>*</font>  
                          </td>
                          </tr>
                          
                          <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">通匯代碼</td>
						  <td width='85%' colspan=2 class="<%=textColor%>">
                            <input type='text' name='EXCHANGE_NO' value="<%if(BN02 != null && BN02.size() != 0 && ((DataObject)BN02.get(0)).getValue("exchange_no") != null ) out.print((String)((DataObject)BN02.get(0)).getValue("exchange_no"));%>" size='7' maxlength='7' >                          
                          </td>
                          </tr>                          
                          
                        </Table></td>
                    </tr>                 
                    <tr>                  
                <td><div align="right"><jsp:include page="getMaintainUser.jsp?width=560" flush="true" /></div></td>                                              
              </tr>
              
              <tr> 
                <td><div align="center"> 
                    <table width="243" border="0" cellpadding="1" cellspacing="1">
                      <tr>                             
				        
				        <%if(act.equals("new") && Utility.getPermission(request,report_no,"A")){ %>      				        
                        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Insert');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_confirmb.gif',1)"><img src="images/bt_confirm.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a></div></td>                        
         				<%}%>
         				<%if(act.equals("Edit") && Utility.getPermission(request,report_no,"U")){ %>                   	        	                                   		              				                     	        	                                   		              				
				        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Update');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image102','','images/bt_updateb.gif',1)"><img src="images/bt_update.gif" name="Image102" width="66" height="25" border="0" id="Image102"></a></div></td>
				        <%}%>
				        <%if(act.equals("Edit") &&Utility.getPermission(request,report_no,"D")){ %>
				        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Delete');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_deleteb.gif',1)"><img src="images/bt_delete.gif" name="Image103" width="66" height="25" border="0" id="Image103"></a></div></td>										        
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
                <td><table width="560" border="0" cellpadding="1" cellspacing="1" class="sbody">
                    <tr> 
                      <td colspan="2"><font color='#990000'><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle"><font color="#007D7D" size="3">使用說明 
                        : </font></font></td>
                    </tr>
                    <tr> 
                      <td width="16">&nbsp;</td>
                      <td width="500"> <ul>                                            
                          <li>本網頁提供新增分支機構代碼功能。</li>                          
                          <li>農(漁)會分支機構代碼，須填入7碼，且其前三碼須與總機構前三碼相同，其他至少是3碼。</li>      
                          <li>農(漁)會的通匯代碼，須填入7碼。</li>      
                          <li>新增時,可直接於空格內輸入資料，按<font color="#666666">【確定】</font>即將本表上的資料於資料庫中建檔。</li>
                          <li>修改時,資料更改完畢後，按<font color="#666666">【修改】</font>即將本表上的資料於資料庫中建檔。</li>                        
                          <li>按<font color="#666666">【取消】</font>即重新輸入資料。</li>                         
                          <li>按<font color="#666666">【回上一頁】則放棄新增總機構代碼, 回至上個畫面</font>。</li>
                          <li>【<font color="red">*</font>】為必填欄位。</li> 
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
<script language="JavaScript" >
<!--
setSelect(this.document.forms[0].BANK_TYPE,"<%=szbank_type%>");
changeOption(this.document.forms[0]);
setSelect(this.document.forms[0].TBANK_NO,"<%=sztbank_no%>");
-->
</script>
</html>
