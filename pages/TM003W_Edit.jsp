<%
//105.09.29 create by2968
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
	List dbData= (request.getAttribute("dbData")==null)?null:(List)request.getAttribute("dbData");
	List accDivList01= (request.getAttribute("accDivList01")==null)?null:(List)request.getAttribute("accDivList01");
	List accDivList02= (request.getAttribute("accDivList02")==null)?null:(List)request.getAttribute("accDivList02");
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");
	String bank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");
	String nowtbank_no =  ( request.getParameter("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");
	if(!nowtbank_no.equals("")){
	   session.setAttribute("nowtbank_no",nowtbank_no);//將已點選的tbank_no寫入session
	}
	bank_no = ( session.getAttribute("nowtbank_no")==null ) ? bank_no : (String)session.getAttribute("nowtbank_no");
	String acc_tr_type = "";
	String acc_tr_name = "";
	if(accDivList01!=null && accDivList01.size()>0){
		acc_tr_type = String.valueOf(((DataObject)accDivList01.get(0)).getValue("acc_tr_type"));
		acc_tr_name = String.valueOf(((DataObject)accDivList01.get(0)).getValue("acc_tr_name"));
	}else if(accDivList02!=null && accDivList02.size()>0){
		acc_tr_type = String.valueOf(((DataObject)accDivList02.get(0)).getValue("acc_tr_type"));
		acc_tr_name = String.valueOf(((DataObject)accDivList02.get(0)).getValue("acc_tr_name"));
	}
	//取得TM003W的權限
  	Properties permission = ( session.getAttribute("TM003W")==null ) ? new Properties() : (Properties)session.getAttribute("TM003W"); 
  	if(permission == null){
         System.out.println("TM003W_Edit.permission == null");
      }else{
         System.out.println("TM003W_Edit.permission.size ="+permission.size());
                 
      }
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/TM003W.js"></script>

<script language="javascript" event="onresize" for="window"></script>
<head>
<title>適用協助措施之經辦機構預估需求填報作業</title>
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
<Form name='form' method=post action='/pages/TM003W.jsp' >
<input type='hidden' name='loan_Sbm'>
<input type='hidden' name='acc_Tr_Type' value='<%=acc_tr_type%>'>
<table width="321" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
        <td width="618"><img src="images/space_1.gif" width="12" height="12"></td>
      </tr>
      <tr>
          <td bgcolor="#FFFFFF" width="618">
      	  <table width="569" border="0" align="center" cellpadding="0" cellspacing="0" height="328">      	  
              <tr>
                <td width="660" height="18">
                <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="100" align="center"><img src="images/banner_bg1.gif" width="100" height="17" align="left"></td>
                      <td width="400" align="center"><b>
                        <center><font color="#000000" size="4">適用協助措施之經辦機構預估需求填報作業</font>
                        </center>
                        </b> </td>
                      <td width="100" valign="middle"><img src="images/banner_bg1.gif" width="100" height="17" align="right"></td>
                    </tr>
                </table>
                </td>
              </tr>
              <tr>
                <td width="660" height="150">
                <table width="638" border="0" align="center" cellpadding="0" cellspacing="0" >
                    <tr><td width="638" height="10"></tr>
                    <tr> 
                      <div align="right"><jsp:include page="getLoginUser.jsp?width=638" flush="true" /></div> 
                    </tr>   
                    <tr>
            		<td width="645" height="82">
					<table width=638 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99" >
            		<tr>
            			<td class="sbody" width='240' bgcolor='#D8EFEE' align='left' height="28" colspan="2">
							協助措施名稱</td>
            			<td class="sbody" width='385' bgcolor='e7e7e7' height="33" colspan="2">
							<%=acc_tr_name %></td>
            		</tr>
            	    
					<%if(accDivList01!=null && accDivList01.size()>0){ %>
					<tr class="sbody">
            			<td width='89' bgcolor='#D8EFEE' align='left' rowspan="<%=accDivList01.size()+1%>">
							舊貸展延需求</td>
            			<td width='146' bgcolor='#D8EFEE' align='left' >
							貸款種類</td>
            			<td width='93' bgcolor='e7e7e7' >
							<p style="margin-left: 0; margin-top: 2; margin-bottom: 2">戶數</td>
            			<td width='263' bgcolor='e7e7e7' height="28">貸款餘額</td>
            	    </tr>
						<%for(int i=0;i<accDivList01.size();i++){ 
							String acc_code01 = String.valueOf(((DataObject)accDivList01.get(i)).getValue("acc_code"));
							%>
							<tr class="sbody">
		            			<td width='146' bgcolor='#D8EFEE' align='left' >
		            				<input type='hidden' name='acc_Code01' id='acc_Code01' value='<%=acc_code01 %>'>
									<%=String.valueOf(((DataObject)accDivList01.get(i)).getValue("acc_name")) %></td>
									<%
									String loan_cnt = "0";
									String loan_amt = "0";
									if(dbData!=null && dbData.size()>0){
										for(int s=0;s<dbData.size();s++){
											if("01".equals(String.valueOf(((DataObject)dbData.get(s)).getValue("acc_div")))){
												if(acc_code01.equals(String.valueOf(((DataObject)dbData.get(s)).getValue("acc_code")))){
													loan_cnt = ((DataObject)dbData.get(s)).getValue("loan_cnt")==null?"0":String.valueOf(((DataObject)dbData.get(s)).getValue("loan_cnt"));
													loan_amt = ((DataObject)dbData.get(s)).getValue("loan_amt")==null?"0":String.valueOf(((DataObject)dbData.get(s)).getValue("loan_amt"));
												}
											}
										}
									} 
									%>
		            			<td width='93' bgcolor='e7e7e7' >
		                            <p style="margin-left: 0; margin-top: 2; margin-bottom: 2">
		            				<input type='text' name='loan_Cnt01' id='loan_Cnt01' value="<%=loan_cnt %>" size='10' style='text-align: right;' onkeyup="this.value=this.value.replace(/[^0-9]/g,'')"></td>
		            			<td width='263' bgcolor='e7e7e7' >
									<input type='text' name='loan_Amt01' id='loan_Amt01' value="<%=loan_amt %>" size='15' maxlength='14' style='text-align: right;' onkeyup="this.value=this.value.replace(/[^0-9]/g,'')"></td>
		            	    </tr>
	            	    <%} %>
            	    <%} %>
            	   

					<%if(accDivList02!=null && accDivList02.size()>0){ %>
					<tr class="sbody">
            			<td width='89' bgcolor='#BDDE9C' align='left'  rowspan="<%=accDivList02.size()+1%>">
							新貸需求</td>
            			<td width='146' bgcolor='#BDDE9C' align='left' >
							貸款種類</td>
            			<td width='93' bgcolor='#EBF4E1' >
                            <p style="margin-left: 0; margin-top: 2; margin-bottom: 2">戶數</td>
            			<td width='263' bgcolor='#EBF4E1' >貸款金額</td>
            	    </tr>
						<%for(int i=0;i<accDivList02.size();i++){ 
							String acc_code02 = String.valueOf(((DataObject)accDivList02.get(i)).getValue("acc_code"));
							%>
							<tr class="sbody">
		            			<td width='146' bgcolor='#BDDE9C' align='left' >
		            				<input type='hidden' name='acc_Code02' id='acc_Code02' value='<%=acc_code02 %>'>
									<%=String.valueOf(((DataObject)accDivList02.get(i)).getValue("acc_name")) %></td>
									<%
									String loan_cnt = "0";
									String loan_amt = "0";
									if(dbData!=null && dbData.size()>0){
										for(int s=0;s<dbData.size();s++){
											if("02".equals(String.valueOf(((DataObject)dbData.get(s)).getValue("acc_div")))){
												if(acc_code02.equals(String.valueOf(((DataObject)dbData.get(s)).getValue("acc_code")))){
													loan_cnt = ((DataObject)dbData.get(s)).getValue("loan_cnt")==null?"0":String.valueOf(((DataObject)dbData.get(s)).getValue("loan_cnt"));
													loan_amt = ((DataObject)dbData.get(s)).getValue("loan_amt")==null?"0":String.valueOf(((DataObject)dbData.get(s)).getValue("loan_amt"));
												}
											}
										}
									} 
									%>
		            			<td width='93' bgcolor='#EBF4E1' >
		                            <p style="margin-left: 0; margin-top: 2; margin-bottom: 2">
		            				<input type='text' name='loan_Cnt02' id='loan_Cnt02' value="<%=loan_cnt %>" size='10' style='text-align: right;' onkeyup="this.value=this.value.replace(/[^0-9]/g,'')" ></td>
		            			<td width='263' bgcolor='#EBF4E1' >
									<input type='text' name='loan_Amt02' id='loan_Amt02' value="<%=loan_amt %>" size='15' maxlength='14' style='text-align: right;' onkeyup="this.value=this.value.replace(/[^0-9]/g,'')"></td>
		            	    </tr>
	            	    <%} %>
            	    <%} %>
            	        	
            	</Table>
   			</td>
 		</tr>


     	<td width="638" height="21"><div align="right"><div align="right">
		<div align="right"><jsp:include page="getMaintainUser.jsp?width=638" flush="true" /></div>
    	</div>
    	</div></td>
      </table></td>
  </tr>
   <tr>
                <td width="688" height="123">
				<table width="688" border="0" cellpadding="1" cellspacing="1" class="sbody" height="176">
                    <tr >
			<td colspan="13"><div align="center">
			<%if(act.equals("new")){%>
									<%if(permission != null && permission.get("A") != null && permission.get("A").equals("Y")){ %> 
										<a href="javascript:doSubmit(this.document.forms[0],'Update');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_confirmb.gif',1)"><img src="images/bt_confirm.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>
										<a href="javascript:AskReset(form);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_cancelb.gif',1)"><img src="images/bt_cancel.gif" name="Image105" width="66" height="25" border="0" id="Image105"></a>
									<%}%>
								<%}%>
								<%if(act.equals("Edit")){%>
									<%if(permission != null && permission.get("U") != null && permission.get("U").equals("Y")){ %>
										 <a href="javascript:doSubmit(this.document.forms[0],'Update');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image102','','images/bt_updateb.gif',1)"><img src="images/bt_update.gif" name="Image102" width="66" height="25" border="0" id="Image102"></a>
											        		<!-- <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Update');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image102','','images/bt_updateb.gif',1)"><img src="images/bt_update.gif" name="Image102" width="66" height="25" border="0" id="Image102"></a></div></td> -->
									<%}%>
									<%if(permission != null && permission.get("D") != null && permission.get("D").equals("Y")){ %>
										 <a href="javascript:doSubmit(this.document.forms[0],'Delete');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_deleteb.gif',1)"><img src="images/bt_delete.gif" name="Image103" width="66" height="25" border="0" id="Image103"></a>
									<%}%>
								<%}%>
							    <a href="javascript:doSubmit(form,'List');"  onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image106','','images/bt_backb.gif',1)"><img src="images/bt_back.gif" name="Image106" width="80" height="25" border="0" id="Image106"></a>
							    
	    		</div></td> 
	    </tr>
                    <tr>
                      <td colspan="2" width="684" height="41"><font color='#990000'><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle"><font color="#007D7D" size="3">使用說明        
                        :</font></font> </td>
                    </tr>
                    <tr>
                      <td width="16" height="127">&nbsp;</td>
                      <td width="664" height="127">
 <ul>
                      	 
                      	 
                      	  <li class="sbody" >確認輸入資料無誤後, 按<font color="#666666">【確定】即將本表上的資料, 於資料庫中建檔。</li>              
                      	  <li class="sbody" >按<font color="#666666">【修改】即修改的資料,寫入資料庫料庫中。</li>
                          <li class="sbody" >欲重新輸入資料, 按<font color="#666666">【取消】即將本表上的資料清空</li>              
                          <li class="sbody" >如放棄修改或無修改之資料需輸入, 按【回上一頁】]即離開本程式。</li>              
                          <li class="sbody" >【<font color="red">*</font>】為必填欄位。</li>
                          <li class="sbody" >【<font color="red">*</font>】如果沒有申報仍請填「0」</font></font></font><font color="red"> </li>   
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
</table>
</form>
</body>
