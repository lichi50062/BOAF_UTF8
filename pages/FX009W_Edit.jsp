<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%
	List WLX09_S_WARNING= (request.getAttribute("WLX09_S_WARNING")==null)?null:(List)request.getAttribute("WLX09_S_WARNING");
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");
        String loaddata = ( request.getParameter("loaddata")==null ) ? "" : (String)request.getParameter("loaddata");
        String syear = ( request.getParameter("S_YEAR")==null ) ? "" : (String)request.getParameter("S_YEAR");
        String squarter = ( request.getParameter("S_QUARTER")==null ) ? "" : (String)request.getParameter("S_QUARTER");
        if(WLX09_S_WARNING!=null && !loaddata.equals("loaded"))//若不是要載入上月資料並且是修改資料時
        {
        syear =String.valueOf(((DataObject)WLX09_S_WARNING.get(0)).getValue("m_year"));
        squarter = String.valueOf(((DataObject)WLX09_S_WARNING.get(0)).getValue("m_quarter"));
        }

	Properties permission = ( session.getAttribute("FX009W")==null ) ? new Properties() : (Properties)session.getAttribute("FX009W");
	if(permission == null){
       System.out.println("FX009W_List.permission == null");
       }else{
       System.out.println("FX009W_List.permission.size ="+permission.size());
       }

	String bank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");
	String nowtbank_no =  ( request.getParameter("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");
	if(!nowtbank_no.equals("")){
	   session.setAttribute("nowtbank_no",nowtbank_no);//將已點選的tbank_no寫入session
	}
	bank_no = ( session.getAttribute("nowtbank_no")==null ) ? bank_no : (String)session.getAttribute("nowtbank_no");

String warnaccount_tcnt="";
String warnaccount_tbal="";
String warnaccount_remit_tcnt="";
String warnaccount_refund_apply_cnt="";
String warnaccount_refund_apply_amt="";
String warnaccount_refund_cnt="";
String warnaccount_refund_amt="";

if(WLX09_S_WARNING!=null)
{
warnaccount_tcnt = String.valueOf(((DataObject)WLX09_S_WARNING.get(0)).getValue("warnaccount_tcnt"));
warnaccount_tbal = String.valueOf(((DataObject)WLX09_S_WARNING.get(0)).getValue("warnaccount_tbal"));
warnaccount_remit_tcnt = String.valueOf(((DataObject)WLX09_S_WARNING.get(0)).getValue("warnaccount_remit_tcnt"));
warnaccount_refund_apply_cnt = String.valueOf(((DataObject)WLX09_S_WARNING.get(0)).getValue("warnaccount_refund_apply_cnt"));
warnaccount_refund_apply_amt = String.valueOf(((DataObject)WLX09_S_WARNING.get(0)).getValue("warnaccount_refund_apply_amt"));
warnaccount_refund_cnt = String.valueOf(((DataObject)WLX09_S_WARNING.get(0)).getValue("warnaccount_refund_cnt"));
warnaccount_refund_amt = String.valueOf(((DataObject)WLX09_S_WARNING.get(0)).getValue("warnaccount_refund_amt"));
}
%>

<script language="javascript" src="js/FX009W.js"></script>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<head>
<title>金融機構警示帳戶調查資料維護</title>
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
<%if(act.equals("modify")){%>
<script language="javascript" type="text/JavaScript">
alert("已有該年季資料，將進行修改動作!");
</script>
<%}%>
<table width="321" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
        <td width="618"><img src="images/space_1.gif" width="12" height="12"></td>
      </tr>
        <tr>
          <td bgcolor="#FFFFFF" width="618">
      <table width="569" border="0" align="center" cellpadding="0" cellspacing="0" height="328">
      <form method="post" name="atmdata">
              <tr>
                <td width="660" height="18"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="154" align="center"><img src="images/banner_bg1.gif" width="150" height="17" align="left"></td>
                      <td width="288" align="center"><b>
                        <center><font color="#000000" size="4">金融機構警示帳戶調查資料維護</font>
                        </center>
                        </b> </td>
                      <td width="154" valign="middle"><img src="images/banner_bg1.gif" width="150" height="17" align="right"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr>
                <td width="660" height="300">
                <table width="579" border="0" align="center" cellpadding="0" cellspacing="0" height="8">
                    <tr>
                    <td width="600" height="10">
                    </tr>
                    <tr>
                      
     <div align="right"><jsp:include page="getLoginUser.jsp?" flush="true" /></div>
                    </tr>
                    <tr>
            <td width="600" height="82">
              <table width="603" border="1" align="center" cellpadding="1" cellspacing="1" bordercolor="#3A9D99" height="274">
                <tr>
                  <td class="sbody" width="196" bgcolor="#D8EFEE" align="left" height="27" colspan="3">金融機構代號</td>
                  <td class="sbody" width="388" bgcolor="e7e7e7" height="27">
                  <%=bank_no%>
                  </td>
                </tr>
                <tr class="sbody">
                  <td width="196" bgcolor="#D8EFEE" align="left" height="1" colspan="3">申報年季<font color="red">*</font></td>
                  <td width="388" bgcolor="e7e7e7" height="1">
                    <p style="margin-left: 0; margin-top: 2; margin-bottom: 2">
                    <input type="text" name="nyear" value="<%=syear%>" size="3" maxlength="3" onblur="CheckYear(this)"
                     <%
                             if(!syear.equals("")){
                              out.print("disabled");
                              }
                            %>> 
                      <font color="#000000">年     
                      <select name="nquarter" size="1"
                   		 <%
                    		if(!squarter.equals("")){
                    		out.print("disabled");
                    		}
                    		%>>
                      <option><%=squarter%></option>
                    
                    </select>季&nbsp;     
                    <%if(!squarter.equals("") && !syear.equals("")){
                    %>
                    <input type="hidden" name="hyear" value="<%=syear%>">             
                    <input type="hidden" name="hquarter" value="<%=squarter%>"><%}%>    
                    <%if(loaddata.equals("ok")){%>          
                    <input type="button" name="load" value="載入上季申報資料" onClick="javascript:doSubmit(this.document.forms[0],'load');">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;     
                   
                    </p><%}%>            
                    </font></td>
                </tr>
                <tr class="sbody">
                  <td width="75" bgcolor="#D8EFEE" height="77" rowspan="3">
                    <p align="center" style="margin-top: 2">警示帳戶(91.11.1~該季末日)</p>
                  </td>
                  <td width="112" bgcolor="#D8EFEE" align="left" height="34" colspan="2">總戶數<font color="red">*</font></td>
                  <td width="388" bgcolor="e7e7e7" height="34">
                  <input type='text' name='warnaccount_tcnt' value="<%=Utility.setCommaFormat(warnaccount_tcnt)%>" size='10' maxlength='20' onFocus='this.value=changeVal(this)'
            				 onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;'>
                  </td>
                </tr>
                <tr class="sbody">
                  <td width="112" bgcolor="#D8EFEE" align="left" height="34" colspan="2">總餘額<font color="red">*</font></td>
                  <td width="388" bgcolor="e7e7e7" height="34">
                  <input type='text' name='warnaccount_tbal' value="<%=Utility.setCommaFormat(warnaccount_tbal)%>" size='20' maxlength='20' onFocus='this.value=changeVal(this)'
            				 onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;'>
            				<font color='red' size="1" face="標楷體">(單位:新台幣元)</font>    </td>
                </tr>
                <tr class="sbody">
                  <td width="112" bgcolor="#D8EFEE" align="left" height="34" colspan="2">警示帳戶內所匯（轉）入總筆數<font color="red">*</font></td>
                  <td width="388" bgcolor="e7e7e7" height="34">
                  <input type='text' name='warnaccount_remit_tcnt' value="<%=Utility.setCommaFormat(warnaccount_remit_tcnt)%>" size='10' maxlength='20' onFocus='this.value=changeVal(this)'
            				 onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;'>
                  </td>
                </tr>
                <tr class="sbody">
                  <td width="75" align="left" bgcolor="#D8EFEE" height="125" rowspan="4">
                    <p align="center">警示帳戶內剩餘款項之返還情形(91.11.1~該季末日)</p>
                  </td>
                  <td width="53" align="left" bgcolor="#D8EFEE" height="58" rowspan="2">申請退還</td>
                  <td width="59" align="left" bgcolor="#D8EFEE" height="34">戶數<font color="red">*</font></td>
                  <td width="388" bgcolor="e7e7e7" height="34">
									<input type='text' name='warnaccount_refund_apply_cnt' value="<%=Utility.setCommaFormat(warnaccount_refund_apply_cnt)%>" size='10' maxlength='20' onFocus='this.value=changeVal(this)'
            				 onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;'>                  
                  </td>
                </tr>
                <tr class="sbody">
                  <td width="59" align="left" bgcolor="#D8EFEE" height="34">金額<font color="red">*</font></td>
                  <td width="388" bgcolor="e7e7e7" height="34">
                  <input type='text' name='warnaccount_refund_apply_amt' value="<%=Utility.setCommaFormat(warnaccount_refund_apply_amt)%>" size='20' maxlength='20' onFocus='this.value=changeVal(this)'
            				 onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;'>
            				 <font color='red' size="1" face="標楷體">(單位:新台幣元)</font>    </td>
                </tr>
                <tr class="sbody">
                  <td width="53" align="left" bgcolor="#D8EFEE" height="61" rowspan="2">巳辦理退還</td>
                  <td width="59" align="left" bgcolor="#D8EFEE" height="35">戶數<font color="red">*</font></td>
                  <td width="388" bgcolor="e7e7e7" height="35">
                  <input type='text' name='warnaccount_refund_cnt' value="<%=Utility.setCommaFormat(warnaccount_refund_cnt)%>" size='10' maxlength='20' onFocus='this.value=changeVal(this)'
            				 onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;'></td>
                </tr>
                
                <tr class="sbody">
                  <td width="59" align="left" bgcolor="#D8EFEE" height="35">金額<font color="red">*</font></td>
                  <td width="388" bgcolor="e7e7e7" height="35">
                  <input type='text' name='warnaccount_refund_amt' value="<%=Utility.setCommaFormat(warnaccount_refund_amt)%>" size='20' maxlength='20' onFocus='this.value=changeVal(this)'
            				 onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;'>
            				 <font color='red' size="1" face="標楷體">(單位:新台幣元)</font></td>
                    </tr>
                </form>
              </table>
              <p>　
   </td>
 </tr>
                    <td width="600" height="21"><div align="right"><div align="right">

<div align="right"><jsp:include page="getMaintainUser.jsp" flush="true" /></div>

                    </div>
                      </div></td>
      </table></td>
  </tr>
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
                <td width="66"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'add');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_confirmb.gif',1)"><img src="images/bt_confirm.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a></div></td>
               <% } }else{
               if(permission != null && permission.get("U") != null && permission.get("U").equals("Y")){//Update %>
                <td width="66"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'modify');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_updateb.gif',1)"><img src="images/bt_update.gif" name="Image103" width="66" height="25" border="0" id="Image103"></a></div></td>
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
 <ul>                 	  <li class="sbody" >確認輸入資料無誤後, 按<font color="#666666">【確定】即將本表上的資料, 於資料庫中建檔。</li>     
                      	  <li class="sbody" >按<font color="#666666">【修改】即修改的資料,寫入資料庫料庫中。</li>
                          <li class="sbody" >欲重新輸入資料, 按<font color="#666666">【取消】即將本表上的資料清空</li>     
                          <li class="sbody" >如放棄修改或無修改之資料需輸入, 按【回上一頁】]即離開本程式。</li>     
                          <li class="sbody" >【<font color="red">*</font>】為必填欄位。</li>
                          <li class="sbody" >【<font color="red">*</font>】如果沒有申報仍請填「0」</li>
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
