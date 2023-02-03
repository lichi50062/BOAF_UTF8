<%
//96.07.10 create by 2295
//99.10.05 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295 
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="java.util.*" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%
	String YEAR  = Utility.getYear();
   	String MONTH = Utility.getMonth();
   	
	System.out.println("WMFileEdit_A08.jsp");	
	//String Report = ( request.getParameter("Report")==null ) ? "" : (String)request.getParameter("Report");		
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");		
	String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? YEAR : (String)request.getParameter("S_YEAR");		
	String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? MONTH : (String)request.getParameter("S_MONTH");		
	System.out.println("S_MONTH="+S_MONTH);
	//String Acc_Div = ( request.getParameter("Acc_Div")==null ) ? "" : (String)request.getParameter("Acc_Div");		
	//System.out.println("Report="+Report);
	List data_div01 = null;
	String WarnAccount_Cnt="";//警示帳戶」戶數（A）
    String LimitAccount_Cnt="";//「衍生管制帳戶」戶數（B）
  	String ErrorAccount_Cnt="";//「自行篩選有異常並已採資金流出管制措施之存款帳戶」戶數（不包括警示帳戶及衍生管制帳戶）（C）
  	String OtherAccount_Cnt="";//其他帳戶戶數（D）
  	String DepositAccount_TCnt="";  
  	DataObject bean = null;
  	data_div01 = (List)request.getAttribute("data_div01");
	if(data_div01 != null && data_div01.size() != 0){	 		
		System.out.println("data_div01.size="+data_div01.size());
		bean = (DataObject)data_div01.get(0);
		WarnAccount_Cnt=(bean.getValue("warnaccount_cnt") == null)?"0":(bean.getValue("warnaccount_cnt")).toString();
		LimitAccount_Cnt=(bean.getValue("limitaccount_cnt") == null)?"0":(bean.getValue("limitaccount_cnt")).toString();
		ErrorAccount_Cnt=(bean.getValue("erroraccount_cnt") == null)?"0":(bean.getValue("erroraccount_cnt")).toString();
		OtherAccount_Cnt=(bean.getValue("otheraccount_cnt") == null)?"0":(bean.getValue("otheraccount_cnt")).toString();
		DepositAccount_TCnt=(bean.getValue("depositaccount_tcnt") == null)?"0":(bean.getValue("depositaccount_tcnt")).toString();		
	}
	
	Properties permission = ( session.getAttribute("WMFileEdit")==null ) ? new Properties() : (Properties)session.getAttribute("WMFileEdit"); 
	if(permission == null){
       System.out.println("WMFileEdit_A08.permission == null");
    }else{
       System.out.println("WMFileEdit_A08.permission.size ="+permission.size());
               
    }
%>
<html>
<head>
<style>
all.clsMenuItemNS{font: x-small Verdana; color: white; text-decoration: none;}
.clsMenuItemIE{text-decoration: none; font: x-small Verdana; color: white; cursor: hand;}
A:hover {color: white;}
</style>
<%if(act.equals("Query")){%>
<title>申報資料查詢</title>
<%}else{%>
<title>線上編輯申報資料</title>
<%}%>

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
<script language="javascript" event="onresize" for="window"></script>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/WMFileEdit.js"></script>
</head>

<body marginwidth="0" marginheight="0" leftmargin="0" topmargin="0" leftmargin="0">
<script language="JavaScript" src="js/menu.js"></script>
<!--
不需使用浮動視窗
<script language="JavaScript" src="js/menucontext_A04.js"></script> 
<script language="JavaScript">
showToolbar();
</script>
-->
<script language="JavaScript">
function UpdateIt(){
if (document.all){
document.all["MainTable"].style.top = document.body.scrollTop;
setTimeout("UpdateIt()", 200);
}
}
UpdateIt();
</script>
<form name='frmWMFileEdit' method=post action='/pages/WMFileEdit.jsp'>
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
                      <td width="100"><img src="images/banner_bg1.gif" width="100" height="17"></td>
                      <td width="*"><font color='#000000' size=4><b> 
                        <center>
                          <b> 
                          <center>
                          <%if(act.equals("Query")){%>
                            <font color='#000000' size=4>申報資料查詢</font> 
                          <%}else{%>
                            <font color='#000000' size=4>線上編輯</font><font color="#CC0000">【<font size=4><%=ListArray.getDLIdName("1", "A08")%>】</font></font><font color='#000000' size=4></font> 
                          <%}%>  
                          </center>
                          </b> 
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
                <td><table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                       <div align="right"><jsp:include page="getLoginUser.jsp" flush="true" /></div> 
                    </tr>
                    <tr> 
                      <td><Table width=600 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
                          <%if(act.equals("Query")){%>
                      	  <tr class="sbody"> 
                            <td width="112" bgcolor='#D8EFEE'> <div align=left>申報資料</div></td>
                            <td colspan=2 bgcolor='e7e7e7'>A08&nbsp;&nbsp;&nbsp;<%=ListArray.getDLIdName("1", "A08")%></td>
                          </tr>  
                          <%}%>
                          <tr class="sbody"> 
                            <td bgcolor="#D2F0FF" width="175"> <div align=left>申報年月</div></td>
                            <td bgcolor='e7e7e7'>
                            <input type='text' name='S_YEAR' value="<%=S_YEAR%>" <%if(act.equals("Edit")) out.print("disabled");%> size='3' maxlength='3' onblur='CheckYear(this)'>
        						<font color='#000000'>年
        						<select id="hide1" name=S_MONTH <%if(act.equals("Edit")) out.print("disabled");%>>
        						<option></option>
        						<%
        							for (int j = 1; j <= 12; j++) {
        							if (j < 10){%>        	
        							<option value=0<%=j%> <%if(S_MONTH.equals(String.valueOf(j))) out.print("selected");%>>0<%=j%></option>        		
            						<%}else{%>
            						<option value=<%=j%> <%if(S_MONTH.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
            						<%}%>
        						<%}%>
        						</select><font color='#000000'>月</font>
                            </td>
                          </tr>
                          
						  
						 <tr class="sbody">	
						    <td bgcolor="#D2F0FF" width="175">							
							   <div align=left>「警示帳戶」戶數（A）</div>
			                </td>
			                <td bgcolor="#e7e7e7">							
			                   <input type='text' name='WarnAccount_Cnt' value="<%=Utility.setCommaFormat(WarnAccount_Cnt)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);addA08(this.form);' style='text-align: right;<%if(Integer.parseInt(WarnAccount_Cnt.equals("")?"0":WarnAccount_Cnt) < 0) out.print(" color:red;");%>'>			                </td>						
			            </tr>
			            <tr class="sbody">	
						    <td bgcolor="#D2F0FF" width="175">							
							   <div align=left>「衍生管制帳戶」戶數（B）</div>
			                </td>
			                <td bgcolor="#e7e7e7">							
			                   <input type='text' name='LimitAccount_Cnt' value="<%=Utility.setCommaFormat(LimitAccount_Cnt)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);addA08(this.form);' style='text-align: right;<%if(Integer.parseInt(LimitAccount_Cnt.equals("")?"0":LimitAccount_Cnt) < 0) out.print(" color:red;");%>'>
			                </td>						
			            </tr>
			           
			            <tr class="sbody">	
						    <td bgcolor="#D2F0FF" width="175">							
							   <div align=left>「自行篩選有異常並已採資金流出管制措施之存款帳戶」戶數（不包括警示帳戶及衍生管制帳戶）（C）</div>
			                </td>
			                <td bgcolor="#e7e7e7">							
			                   <input type='text' name='ErrorAccount_Cnt' value="<%=Utility.setCommaFormat(ErrorAccount_Cnt)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);addA08(this.form);' style='text-align: right;<%if(Integer.parseInt(ErrorAccount_Cnt.equals("")?"0":ErrorAccount_Cnt) < 0) out.print(" color:red;");%>'>
			                </td>						
			            </tr>
			            <tr class="sbody">	
						    <td bgcolor="#D2F0FF" width="175">							
							   <div align=left>其他帳戶戶數（D）</div>
			                </td>
			                <td bgcolor="#e7e7e7">							
			                   <input type='text' name='OtherAccount_Cnt' value="<%=Utility.setCommaFormat(OtherAccount_Cnt)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);addA08(this.form);' style='text-align: right;<%if(Integer.parseInt(OtherAccount_Cnt.equals("")?"0":OtherAccount_Cnt) < 0) out.print(" color:red;");%>'>
			                </td>						
			            </tr>
			             <tr class="sbody">	
						    <td bgcolor="#D2F0FF" width="175">							
							   <div align=left>存款帳戶總戶數 （E）</div>
			                </td>
			                <td bgcolor="#e7e7e7">							
			                   <input type='text' name='DepositAccount_TCnt' readonly value="<%=Utility.setCommaFormat(DepositAccount_TCnt)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this)' style='text-align: right; color:#808080; background-color:#FFFFE6;'>
			                </td>						
			            </tr>
						
	          </Table></td>
                    </tr>
                    <tr> 
                      <td>&nbsp;</td>
                    </tr>
                  </table></td>
              </tr>
              <tr>                  
                <td><div align="right"><jsp:include page="getMaintainUser.jsp" flush="true" /></div></td>              
              </tr>
              <tr> 
                <td>&nbsp;</td>
              </tr>
              <tr> 
                <td><div align="center"> 
                    <table width="243" border="0" cellpadding="1" cellspacing="1">
                      <tr>     
			 	<% //如果.有權限做update,且程科目代號不為空值時才顯示確定跟取消%> 
				<%if(act.equals("new")){%>     
				     <%if(permission != null && permission.get("A") != null && permission.get("A").equals("Y")){ //add%>                   	        	                                   		       
                        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Insert','A08','','','');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_confirmb.gif',1)"><img src="images/bt_confirm.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a></div></td>                        
                     <%}%>   
         		<%}%>
         		<%if(act.equals("Edit")){%>
         		     <%if(permission != null && permission.get("U") != null && permission.get("U").equals("Y")){ //update%>                   	        	                                   		     
				        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Update','A08','<%=S_YEAR%>','<%=S_MONTH%>','');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image102','','images/bt_updateb.gif',1)"><img src="images/bt_update.gif" name="Image102" width="66" height="25" border="0" id="Image102"></a></div></td>			            
				     <%}%>   
				     <%if(permission != null && permission.get("D") != null && permission.get("D").equals("Y")){ //delete%>                   	        	                                   		     
				        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Delete','A08','<%=S_YEAR%>','<%=S_MONTH%>','');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_deleteb.gif',1)"><img src="images/bt_delete.gif" name="Image103" width="66" height="25" border="0" id="Image103"></a></div></td>										               
				     <%}%>   
				<%}%>				
         		<%if(!act.equals("Query")){%>       
         		     <%if( (permission != null && permission.get("A") != null && permission.get("A").equals("Y"))                  	        	                                   		        
         		         ||(permission != null && permission.get("U") != null && permission.get("U").equals("Y"))                  	        	                                   		     
         		         ||(permission != null && permission.get("D") != null && permission.get("D").equals("Y"))){ //Add/Update/delete%>                   	        	                                   		     
                        <td width="66"> <div align="center"><a href="javascript:AskReset(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image104','','images/bt_cancelb.gif',1)"><img src="images/bt_cancel.gif" name="Image104" width="66" height="25" border="0" id="Image104"></a></div></td>
                      <%}%>  
                <%}%>        
                        <td width="93"><div align="center"><a href="javascript:history.back();"  onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image81','','images/bt_backb.gif',1)"><img src="images/bt_back.gif" name="Image81" width="80" height="25" border="0" id="Image81"></a></div></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
              <tr> 
                <td>&nbsp;</td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF"><table width="600" border="0" align="center" cellpadding="1" cellspacing="1">
              <tr> 
                <td><div align="center"><img src="images/line_1.gif" width="600" height="12"></div></td>
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
                          <li>本網頁提供新增<%=ListArray.getDLIdName("1", "A08")%>。</li>
                          <li>承辦員E_MAIL請勿填寫外部免費電子信箱以免無法收到更新結果通知。</li>
                          <li>確認資料無誤後，按<font color="#666666">【確定】</font>即將本網頁上的資料，於資料庫中新增。</li>
                          <li>按<font color="#666666">【取消】</font>即重新輸入資料。</li>
                          <li>點選所列之<font color="#666666">【回上一頁】</font>則放棄資料， 回至前一畫面。</li>
                          <font color='red'>
                          <li>存款帳戶總戶數為(A)+(B)+(C)+(D)合計數。</li>
                          <li>存款帳戶總戶數指:支票存款、定期存款及活期性存款帳戶,含外匯活期存款及行員存款。</li>	                          
                          </font>
                        </ul></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><div align="center"><img src="images/line_1.gif" width="600" height="12"></div></td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
</form>
</body>

</html>
