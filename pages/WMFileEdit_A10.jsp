<%
//97.06.12 create by 2295
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
   	request.setAttribute("table_width","729");
	System.out.println("WMFileEdit_A10.jsp");	
	//String Report = ( request.getParameter("Report")==null ) ? "" : (String)request.getParameter("Report");		
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");		
	String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? YEAR : (String)request.getParameter("S_YEAR");		
	String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? MONTH : (String)request.getParameter("S_MONTH");		
	System.out.println("act="+act);
	
	List data_div01 = null;	
	String loan1_amt="";//放款-第一類
	String loan2_amt="";//放款-第二類
	String loan3_amt="";//放款-第三類
	String loan4_amt="";//放款-第四類
	String loan_sum="";//放款-合計  
	String invest1_amt="";//投資-第一類
	String invest2_amt="";//投資-第二類
	String invest3_amt="";//投資-第三類
	String invest4_amt="";//投資-第四類
	String invest_sum="";//投資-合計
	String other1_amt="";//其他-第一類
	String other2_amt="";//其他-第二類
	String other3_amt="";//其他-第三類
	String other4_amt="";//其他-第四類
	String other_sum="";//其他-合計
	
	String type1_sum="";//註1-第一類合計
  	String type2_sum="";//註1-第二類合計
  	String type3_sum="";//註1-第三類合計
  	String type4_sum="";//註1-第四類合計
  	String type_sum="";//註1-合計
  	
	String loan1_baddebt="";//放款-帳列備抵呆帳-第一類
	String loan2_baddebt="";//放款-帳列備抵呆帳-第二類
	String loan3_baddebt="";//放款-帳列備抵呆帳-第三類
	String loan4_baddebt="";//放款-帳列備抵呆帳-第四類
	String loan_baddebt="";//放款-帳列備抵呆帳-合計
	String build1_baddebt="";//建築貸款-帳列備抵呆帳-第一類
	String build2_baddebt="";//建築貸款-帳列備抵呆帳-第二類
	String build3_baddebt="";//建築貸款-帳列備抵呆帳-第三類
	String build4_baddebt="";//建築貸款-帳列備抵呆帳-第四類
	String build_baddebt="";//建築貸款-帳列備抵呆帳-合計
	
	String type2_sum1="";//註3-第一類合計
  	String type2_sum2="";//註3-第二類合計
  	String type2_sum3="";//註3-第三類合計
  	String type2_sum4="";//註3-第四類合計
  	String type2_sum5="";//註3-合計
  	
	String baddebt_flag="";//備呆等於或大於應提最低標準
	String baddebt_noenough="";//備呆不足額
	String baddebt_delay="";//申請展延
	String baddebt_104="";//104年底前提足備呆
	String baddebt_105="";//105年底前提足備呆
	String baddebt_106="";//106年底前提足備呆
	String baddebt_107="";//107年底前提足備呆
	String baddebt_108="";//108年底前提足備呆
	
	
  	
  	
  	DataObject bean = null;
  	data_div01 = (List)request.getAttribute("data_div01");
	if(data_div01 != null && data_div01.size() != 0){	 		
		System.out.println("data_div01.size="+data_div01.size());
		bean = (DataObject)data_div01.get(0);
		loan1_amt=(bean.getValue("loan1_amt") == null)?"0":(bean.getValue("loan1_amt")).toString();
		loan2_amt=(bean.getValue("loan2_amt") == null)?"0":(bean.getValue("loan2_amt")).toString();//放款-列二類
		loan3_amt=(bean.getValue("loan3_amt") == null)?"0":(bean.getValue("loan3_amt")).toString();//放款-列三類
		loan4_amt=(bean.getValue("loan4_amt") == null)?"0":(bean.getValue("loan4_amt")).toString();//放款-列四類
		invest1_amt=(bean.getValue("invest1_amt") == null)?"0":(bean.getValue("invest1_amt")).toString();
		invest2_amt=(bean.getValue("invest2_amt") == null)?"0":(bean.getValue("invest2_amt")).toString();//投資-列二類
		invest3_amt=(bean.getValue("invest3_amt") == null)?"0":(bean.getValue("invest3_amt")).toString();//投資-列三類		
		invest4_amt=(bean.getValue("invest4_amt") == null)?"0":(bean.getValue("invest4_amt")).toString();//投資-列四類  
		other1_amt=(bean.getValue("other1_amt") == null)?"0":(bean.getValue("other1_amt")).toString();
		other2_amt=(bean.getValue("other2_amt") == null)?"0":(bean.getValue("other2_amt")).toString();//其他-列二類
		other3_amt=(bean.getValue("other3_amt") == null)?"0":(bean.getValue("other3_amt")).toString();//其他-列三類		
		other4_amt=(bean.getValue("other4_amt") == null)?"0":(bean.getValue("other4_amt")).toString();//其他-列四類 
		loan1_baddebt=(bean.getValue("loan1_baddebt")==null ) ? "0" : (bean.getValue("loan1_baddebt")).toString();//放款-帳列備抵呆帳-第一類
		loan2_baddebt=(bean.getValue("loan2_baddebt")==null ) ? "0" : (bean.getValue("loan2_baddebt")).toString();//放款-帳列備抵呆帳-第二類
		loan3_baddebt=(bean.getValue("loan3_baddebt")==null ) ? "0" : (bean.getValue("loan3_baddebt")).toString();//放款-帳列備抵呆帳-第三類
		loan4_baddebt=(bean.getValue("loan4_baddebt")==null ) ? "0" : (bean.getValue("loan4_baddebt")).toString();//放款-帳列備抵呆帳-第四類
		build1_baddebt=(bean.getValue("build1_baddebt")==null ) ? "0" : (bean.getValue("build1_baddebt")).toString();//建築貸款-帳列備抵呆帳-第一類
		build2_baddebt=(bean.getValue("build2_baddebt")==null ) ? "0" : (bean.getValue("build2_baddebt")).toString();//建築貸款-帳列備抵呆帳-第二類
		build3_baddebt=(bean.getValue("build3_baddebt")==null ) ? "0" : (bean.getValue("build3_baddebt")).toString();//建築貸款-帳列備抵呆帳-第三類
		build4_baddebt=(bean.getValue("build4_baddebt")==null ) ? "0" : (bean.getValue("build4_baddebt")).toString();//建築貸款-帳列備抵呆帳-第四類
		baddebt_flag=Utility.getTrimString(bean.getValue("baddebt_flag"));//備呆等於或大於應提最低標準
  		baddebt_noenough=(bean.getValue("baddebt_noenough")==null ) ? "" : Utility.setCommaFormat((bean.getValue("baddebt_noenough")).toString());//備呆不足額
  		baddebt_delay=Utility.getTrimString(bean.getValue("baddebt_delay"));//申請展延
  		baddebt_104=(bean.getValue("baddebt_104")==null ) ? "" : Utility.setCommaFormat((bean.getValue("baddebt_104")).toString());//104年底前提足備呆
  		baddebt_105=(bean.getValue("baddebt_105")==null ) ? "" : Utility.setCommaFormat((bean.getValue("baddebt_105")).toString());//105年底前提足備呆
  		baddebt_106=(bean.getValue("baddebt_106")==null ) ? "" : Utility.setCommaFormat((bean.getValue("baddebt_106")).toString());//106年底前提足備呆
  		baddebt_107=(bean.getValue("baddebt_107")==null ) ? "" : Utility.setCommaFormat((bean.getValue("baddebt_107")).toString());//107年底前提足備呆
  		baddebt_108=(bean.getValue("baddebt_108")==null ) ? "" : Utility.setCommaFormat((bean.getValue("baddebt_108")).toString());//108年底前提足備呆
	}
	
	Properties permission = ( session.getAttribute("WMFileEdit")==null ) ? new Properties() : (Properties)session.getAttribute("WMFileEdit"); 
	if(permission == null){
       System.out.println("WMFileEdit_A10.permission == null");
    }else{
       System.out.println("WMFileEdit_A10.permission.size ="+permission.size());
               
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
function clearFalse(form){
	form.baddebt_delay[0].checked=false;
	form.baddebt_noenough.value='0';
	form.baddebt_delay[1].checked=false;
	form.baddebt_104.value='0';
	form.c5.checked=false; 
	form.c6.checked=false;
	form.c7.checked=false;
	form.c8.checked=false;
	form.baddebt_105.value='0';
	form.baddebt_106.value='0';
	form.baddebt_107.value='0';
	form.baddebt_108.value='0';
}
function clearDelayY(form){
	/*104.03.27客戶通知.104年時可能同時並存.先不清空
	form.baddebt_delay[1].checked=false;
	form.c5.checked=false; 
	form.c6.checked=false;
	form.c7.checked=false;
	form.c8.checked=false;
	form.baddebt_105.value='0';
	form.baddebt_106.value='0';
	form.baddebt_107.value='0';
	form.baddebt_108.value='0';
	*/
}
function clearDelayN(form){
	/*104.03.27客戶通知.104年時可能同時並存.先不清空
	form.baddebt_delay[0].checked=false;
	form.baddebt_104.value='0';
	*/
}
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
			<table width="746" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width=<%=request.getAttribute("table_width") %> border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="90"><img src="images/banner_bg1.gif" width="90" height="17"></td>
                      <td width="*"><font color='#000000' size=4><b> 
                        <center>
                          <b> 
                          <center>
                          <%if(act.equals("Query")){%>
                            <font color='#000000' size=4>申報資料查詢</font> 
                          <%}else{%>
                            <font color='#000000' size=4>線上編輯</font><font color="#CC0000">【<font size=4><%=ListArray.getDLIdName("1", "A10")%>】</font></font><font color='#000000' size=4></font> 
                          <%}%>  
                          </center>
                          </b> 
                        </center>
                        </b></font> </td>
                      <td width="90"><img src="images/banner_bg1.gif" width="90" height="17"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><img src="images/space_1.gif" width="12" height="12"></td>
              </tr>
              <tr> 
                <td><table width="739" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                       <div align="right"><jsp:include page="getLoginUser.jsp" flush="true" /></div> 
                    </tr>
                    <tr> 
                      <td><Table width=<%=request.getAttribute("table_width") %> border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
                          <%if(act.equals("Query")){%>
                      	  <tr class="sbody"> 
                            <td width="313" colspan=2 bgcolor='#D8EFEE'> <div align=left>申報資料</div></td>
                            <td colspan=4 bgcolor='e7e7e7'>A10&nbsp;&nbsp;&nbsp;<%=ListArray.getDLIdName("1", "A10")%></td>
                          </tr>  
                          <%}%>
                          <tr class="sbody"> 
                            <td bgcolor="#D8EFEE" colspan=2 width="313"> <div align=left>申報年月</div></td>
                            <td bgcolor='e7e7e7' colspan="4">
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
        						<input type="button" name="LastMonthDataBtn" value="上月資料匯入" onclick="javascript:doSubmit(this.document.forms[0],'getLastMonthData','A10','','');">&nbsp;&nbsp; 
        						是否確定
        						<select name="LastMonthDataYN" size="1">
                              		<option selected>N</option>
                              		<option>Y</option>
                              	</select-->	
                            </td>
                          </tr>
                          
						  <tr class="sbody">
                  			<td width="210" bgcolor="#D8EFEE" height="51" rowspan="2">
                    		<p align="center" style="margin-top: 2">項目</p>
                  			</td>
                  			<td width="506" bgcolor="#D8EFEE" align="center" height="25" colspan="5">
								應予評估資產金額<font color="#FF0000">【註1】</font> </td>
                		   </tr>
                		   <tr class="sbody">
                		    <td width="98" bgcolor="#D8EFEE" align="center" height="25">第一類</td>
                  			<td width="97" bgcolor="#D8EFEE" align="center" height="25">第二類</td>
                  			<td width="97" bgcolor="#D8EFEE" height="25" align="center">第三類</td>
                  			<td width="98" bgcolor="#D8EFEE" height="25" align="center">第四類</td>
                  			<td width="96" bgcolor="#D8EFEE" height="25" align="center">合計</td>
                		   </tr>
 							
 						   <tr class="sbody">
                  			 <td width="210" bgcolor="#D8EFEE" height="23" align="center">放款</td>
                  			 <td width="98" bgcolor="#E7E7E7" align="left" height="23">                            
                  			 	<input type='text' name='loan1_amt' value="<%=Utility.setCommaFormat(loan1_amt)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);addA10(this.form);' style='text-align: right;'>
			                 </td>
                  			 <td width="97" bgcolor="#E7E7E7" align="left" >                            
                  			 	<input type='text' name='loan2_amt' value="<%=Utility.setCommaFormat(loan2_amt)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);addA10(this.form);' style='text-align: right;'>
			                 </td>
                  			 <td width="97" bgcolor="e7e7e7" height="23">
                            	<input type='text' name='loan3_amt' value="<%=Utility.setCommaFormat(loan3_amt)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);addA10(this.form);' style='text-align: right;'>
			                 </td>
			                 <td width="98" bgcolor="e7e7e7" height="23">
                            	<input type='text' name='loan4_amt' value="<%=Utility.setCommaFormat(loan4_amt)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);addA10(this.form);' style='text-align: right;'>
			                 </td>
			                 <td width="96" bgcolor="e7e7e7" height="23">
                            	<input type='text' name='loan_sum' value="<%=Utility.setCommaFormat(loan_sum)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;  color:#808080; background-color:#FFFFE6;' readonly>
			                 </td>
                		   </tr>
							
						   <tr class="sbody">
                  			 <td width="210" bgcolor="#D8EFEE" height="23" align="center">投資</td>
                  			 <td width="98" bgcolor="#E7E7E7" align="left" height="23">                            
                  			 	<input type='text' name='invest1_amt' value="<%=Utility.setCommaFormat(invest1_amt)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);addA10(this.form);' style='text-align: right;'>
			                 </td>
                  			 <td width="97" bgcolor="#E7E7E7" align="left" height="23">                            
                  			 	<input type='text' name='invest2_amt' value="<%=Utility.setCommaFormat(invest2_amt)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);addA10(this.form);' style='text-align: right;'>
			                 </td>
                  			 <td width="97" bgcolor="e7e7e7" height="23">
                            	<input type='text' name='invest3_amt' value="<%=Utility.setCommaFormat(invest3_amt)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);addA10(this.form);' style='text-align: right;'>
			                 </td>
			                 <td width="98" bgcolor="e7e7e7" height="23">
                            	<input type='text' name='invest4_amt' value="<%=Utility.setCommaFormat(invest4_amt)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);addA10(this.form);' style='text-align: right;'>
			                 </td>
			                 <td width="96" bgcolor="e7e7e7" height="23">
                            	<input type='text' name='invest_sum' value="<%=Utility.setCommaFormat(invest_sum)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;  color:#808080; background-color:#FFFFE6;' readonly>
			                 </td>
                		   </tr>
						 
			           	   <tr class="sbody">
                  			 <td width="210" bgcolor="#D8EFEE" height="23" align="center">其他</td>
                  			 <td width="98" bgcolor="#E7E7E7" align="left" height="23">                            
                  			 	<input type='text' name='other1_amt' value="<%=Utility.setCommaFormat(other1_amt)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);addA10(this.form);' style='text-align: right;'>
			                 </td>
                  			 <td width="97" bgcolor="#E7E7E7" align="left" height="23">                            
                  			 	<input type='text' name='other2_amt' value="<%=Utility.setCommaFormat(other2_amt)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);addA10(this.form);' style='text-align: right;'>
			                 </td>
                  			 <td width="97" bgcolor="e7e7e7" height="23">
                            	<input type='text' name='other3_amt' value="<%=Utility.setCommaFormat(other3_amt)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);addA10(this.form);' style='text-align: right;'>
			                 </td>
			                 <td width="98" bgcolor="e7e7e7" height="23">
                            	<input type='text' name='other4_amt' value="<%=Utility.setCommaFormat(other4_amt)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);addA10(this.form);' style='text-align: right;'>
			                 </td>
			                 <td width="96" bgcolor="e7e7e7" height="23">
                            	<input type='text' name='other_sum' value="<%=Utility.setCommaFormat(other_sum)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;  color:#808080; background-color:#FFFFE6;' readonly>
			                 </td>
                		   </tr>
                		   
                		   <tr class="sbody">
                  			 <td width="210" bgcolor="#D8EFEE" height="23" align="center">合計</td>
                  			 <td width="98" bgcolor="#E7E7E7" align="left" height="23">                            
                  			 	<input type='text' name='type1_sum' value="<%=Utility.setCommaFormat(type1_sum)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;  color:#808080; background-color:#FFFFE6;' readonly >
			                 </td>
                  			 <td width="97" bgcolor="#E7E7E7" align="left" height="23">                            
                  			 	<input type='text' name='type2_sum' value="<%=Utility.setCommaFormat(type2_sum)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;  color:#808080; background-color:#FFFFE6;' readonly >
			                 </td>
                  			 <td width="97" bgcolor="e7e7e7" height="23">
                            	<input type='text' name='type3_sum' value="<%=Utility.setCommaFormat(type3_sum)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;  color:#808080; background-color:#FFFFE6;' readonly >
			                 </td>
			                 <td width="98" bgcolor="e7e7e7" height="23">
                            	<input type='text' name='type4_sum' value="<%=Utility.setCommaFormat(type4_sum)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;  color:#808080; background-color:#FFFFE6;' readonly >
			                 </td>
			                 <td width="96" bgcolor="e7e7e7" height="23">
                            	<input type='text' name='type_sum' value="<%=Utility.setCommaFormat(type_sum)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;  color:#808080; background-color:#FFFFE6;' readonly >
			                 </td>
                		   </tr>
			               <tr class="sbody">
			                  <td width="210" bgcolor="#D8EFEE" height="51" rowspan="2">
                    		<p align="center" style="margin-top: 2">項目</p>
                  			</td>
                  			<td width="506" bgcolor="#D8EFEE" align="center" height="25" colspan="5">
								帳列備抵呆帳<font color="#FF0000">【註2】</font></td>
                		   </tr>
                		   <tr class="sbody">
                  			<td width="98" bgcolor="#D8EFEE" align="center" height="25">
							第一類</td>
                  			<td width="97" bgcolor="#D8EFEE" align="center" height="25">
							第二類</td>
                  			<td width="97" bgcolor="#D8EFEE" height="25" align="center">
							第三類</td>
                  			<td width="98" bgcolor="#D8EFEE" height="25" align="center">
							第四類</td>
                  			<td width="96" bgcolor="#D8EFEE" height="25" align="center">合計</td>
                		   </tr>
 							
 						   <tr class="sbody">
                  			 <td width="210" bgcolor="#D8EFEE" height="23" align="center">放款</td>
                  			 <td width="98" bgcolor="#E7E7E7" align="left" height="23">                            
                  			 	<input type='text' name='loan1_baddebt' value="<%=Utility.setCommaFormat(loan1_baddebt)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);addA10(this.form);this.document.forms[0].baddebt_flag[0].checked=true;' style='text-align: right;'>
			                 </td>
                  			 <td width="97" bgcolor="#E7E7E7" align="left" height="23">                            
                  			 	<input type='text' name='loan2_baddebt' value="<%=Utility.setCommaFormat(loan2_baddebt)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);addA10(this.form);this.document.forms[0].baddebt_flag[1].checked=true;' style='text-align: right;'>
			                 </td>
                  			 <td width="97" bgcolor="e7e7e7" height="23">
                            	<input type='text' name='loan3_baddebt' value="<%=Utility.setCommaFormat(loan3_baddebt)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);addA10(this.form);this.document.forms[0].baddebt_flag[1].checked=true;' style='text-align: right;'>
			                 </td>
			                 <td width="98" bgcolor="e7e7e7" height="23">
                            	<input type='text' name='loan4_baddebt' value="<%=Utility.setCommaFormat(loan4_baddebt)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);addA10(this.form);this.document.forms[0].baddebt_flag[1].checked=true;' style='text-align: right;'>
			                 </td>
			                 <td width="96" bgcolor="e7e7e7" height="23">
                            	<input type='text' name='loan_baddebt' value="<%=Utility.setCommaFormat(loan_baddebt)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;  color:#808080; background-color:#FFFFE6;' readonly>
			                 </td>
                		   </tr>
							
						   <tr class="sbody">
                  			 <td width="210" bgcolor="#D8EFEE" height="25" align="center">
								建築貸款</td>
                  			 <td width="98" bgcolor="#E7E7E7" align="left" height="25">                            
                  			 	<input type='text' name='build1_baddebt' value="<%=Utility.setCommaFormat(build1_baddebt)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);addA10(this.form);' style='text-align: right;'>
			                 </td>
                  			 <td width="97" bgcolor="#E7E7E7" align="left" height="25">                            
                  			 	<input type='text' name='build2_baddebt' value="<%=Utility.setCommaFormat(build2_baddebt)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);addA10(this.form);' style='text-align: right;'>
			                 </td>
                  			 <td width="97" bgcolor="e7e7e7" height="25">
                            	<input type='text' name='build3_baddebt' value="<%=Utility.setCommaFormat(build3_baddebt)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);addA10(this.form);' style='text-align: right;'>
			                 </td>
			                 <td width="98" bgcolor="e7e7e7" height="25">
                            	<input type='text' name='build4_baddebt' value="<%=Utility.setCommaFormat(build4_baddebt)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);addA10(this.form);' style='text-align: right;'>
			                 </td>
			                 <td width="96" bgcolor="e7e7e7" height="25">
                            	<input type='text' name='build_baddebt' value="<%=Utility.setCommaFormat(build_baddebt)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;  color:#808080; background-color:#FFFFE6;' readonly>
			                 </td>
                		   </tr>
						 
						 
				        <tr class="sbody">
				        <td width="210" bgcolor="#D8EFEE" height="49" rowspan="2">
                    		<p align="center" style="margin-top: 2">項目</p>
                  			</td>
                  			<td width="506" bgcolor="#D8EFEE" align="center" height="23" colspan="5">
								依規定應提列最低標準之備抵呆帳<font color="#FF0000">【註3】</font></td>
                		   </tr>
                		   <tr class="sbody">
                  			<td width="98" bgcolor="#D8EFEE" height="25" align="center">
							第一類( 1%)</td>
                  			<td width="97" bgcolor="#D8EFEE" height="25" align="center">
							第二類( 2%)</td>
                  			<td width="97" bgcolor="#D8EFEE" height="25" align="center">
							第三類(50%)</td>
                  			<td width="98" bgcolor="#D8EFEE" height="25" align="center">
							第四類</td>
                  			<td width="96" bgcolor="#D8EFEE" height="25" align="center">合計</td>
                		   </tr>
 							
 						   <tr class="sbody">
                  			 <td width="210" bgcolor="#D8EFEE" height="23" align="center">放款</td>
                  			 <td width="98" bgcolor="#E7E7E7" align="left" height="23">                            
                  			 	<input type='text' name='type2_sum1' value="<%=Utility.setCommaFormat(type2_sum1)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;  color:#808080; background-color:#FFFFE6;' readonly >
			                 </td>
                  			 <td width="97" bgcolor="#E7E7E7" align="left" height="23">                            
                  			 	<input type='text' name='type2_sum2' value="<%=Utility.setCommaFormat(type2_sum2)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;  color:#808080; background-color:#FFFFE6;' readonly >
			                 </td>
                  			 <td width="97" bgcolor="e7e7e7" height="23">
                            	<input type='text' name='type2_sum3' value="<%=Utility.setCommaFormat(type2_sum3)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;  color:#808080; background-color:#FFFFE6;' readonly >
			                 </td>
			                 <td width="98" bgcolor="e7e7e7" height="23">
                            	<input type='text' name='type2_sum4' value="<%=Utility.setCommaFormat(type2_sum4)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;  color:#808080; background-color:#FFFFE6;' readonly >
			                 </td>
			                 <td width="96" bgcolor="e7e7e7" height="23">
                            	<input type='text' name='type2_sum5' value="<%=Utility.setCommaFormat(type2_sum5)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;  color:#808080; background-color:#FFFFE6;' readonly >
			                 </td>
                		   </tr>
                		   <tr class="sbody">
				        <td width="721" bgcolor="#D8EFEE" height="26" colspan="6">
                    		<p align="center">備抵呆帳是否依規定提足<font color="#FF0000">【註4】</font></td>
                		   </tr>
                		   <tr class="sbody">
				        <td width="721" bgcolor="#D8EFEE" height="25" colspan="6">
                    		&nbsp;<input type="radio" name="baddebt_flag" value="Y" <%if("Y".equals(baddebt_flag)){%>checked<%}%> onClick='clearFalse(this.form);'>是：帳列備抵呆帳等於或大於依規定應提列最低標準之備抵呆帳 </td>
                		   </tr>
                		   <tr class="sbody">
				        		<td width="721" bgcolor="#D8EFEE" height="25" colspan="6">
                    			&nbsp;<input type="radio" name="baddebt_flag" value="N" <%if("N".equals(baddebt_flag)){%>checked<%}%>>否：帳列備抵呆帳小於依規定應提列最低標準之備抵呆帳<br>
                    			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    			備抵呆帳提列不足額部分
                    		 	<input type='text' name='baddebt_noenough' value="<%=Utility.setCommaFormat(baddebt_noenough)%>" size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;'>
                    		 	<br>
                    		 	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
								不足額分年提撥情形：<br>
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								<input type="checkbox" name="baddebt_delay" value="N" <%if("N".equals(baddebt_delay)){%>checked<%}%> onClick='clearDelayY(form);'>可於1年內提足(即於104年底前提足)，預計提撥金額
							 	<input type='text' name='baddebt_104'  size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;'><br>
							 	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							 	<input type="checkbox" name="baddebt_delay" value="Y" <%if("Y".equals(baddebt_delay)){%>checked<%}%> onClick='clearDelayN(form);'>無法於1年內提足者，申請展延：<br>
							 	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							 	<input type="checkbox" name="c5" <%if(!"".equals(baddebt_105) && !"0".equals(baddebt_105)){%>checked<%}%> >105年底前提撥金額
							 	<input type='text' name='baddebt_105' size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;'><br>
							 	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							 	<input type="checkbox" name="c6" <%if(!"".equals(baddebt_106) && !"0".equals(baddebt_106)){%>checked<%}%>>106年底前提撥金額
							 	<input type='text' name='baddebt_106' size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;'><br>
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								<input type="checkbox" name="c7" <%if(!"".equals(baddebt_107) && !"0".equals(baddebt_107)){%>checked<%}%>>107年底前提撥金額
								<input type='text' name='baddebt_107' size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;'><br>
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								<input type="checkbox" name="c8" <%if(!"".equals(baddebt_108) && !"0".equals(baddebt_108)){%>checked<%}%>>108年底前提撥金額
								<input type='text' name='baddebt_108' size=16 maxlength=16 onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;'><br>

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
                        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Insert','A10','','','');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_confirmb.gif',1)"><img src="images/bt_confirm.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a></div></td>                        
                     <%}%>   
         		<%}%>
         		<%if(act.equals("Edit")){%>
         		     <%if(permission != null && permission.get("U") != null && permission.get("U").equals("Y")){ //update%>                   	        	                                   		     
				        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Update','A10','<%=S_YEAR%>','<%=S_MONTH%>','');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image102','','images/bt_updateb.gif',1)"><img src="images/bt_update.gif" name="Image102" width="66" height="25" border="0" id="Image102"></a></div></td>			            
				     <%}%>   
				     <%if(permission != null && permission.get("D") != null && permission.get("D").equals("Y")){ //delete%>                   	        	                                   		     
				        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Delete','A10','<%=S_YEAR%>','<%=S_MONTH%>','');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_deleteb.gif',1)"><img src="images/bt_delete.gif" name="Image103" width="66" height="25" border="0" id="Image103"></a></div></td>										               
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
          <td bgcolor="#FFFFFF"><table width=<%=request.getAttribute("table_width") %> border="0" align="center" cellpadding="1" cellspacing="1">
              <tr> 
                <td><div align="center"><img src="images/line_1.gif" width=<%=request.getAttribute("table_width") %> height="12"></div></td>
              </tr>
              <tr> 
                <td><table width=<%=request.getAttribute("table_width") %> border="0" cellpadding="1" cellspacing="1" class="sbody">
                    <tr> 
                      <td colspan="2"><font color='#990000'><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle"><font color="#007D7D" size="3">使用說明 
                        : </font></font></td>
                    </tr>
                    <tr> 
                      <td width="16">&nbsp;</td>
                      <td width="729"> <ul>
                          <li>本網頁提供新增<%=ListArray.getDLIdName("1", "A10")%>。</li>
                          <li class="sbody" >
                           <font color="red">【註1】</font><font color="#FF0000">
                          				 應予評估資產，指應予評估之放款、投資(有價證券)及其他資產，經評估列為第一類&nbsp;&nbsp;&nbsp; (正常)、第二類(可望全數收回)、第三類(收回有困難)及第四類(收回無望)者；應予評估放款請依「農會漁會信用部資產評估損失準備提列及逾期放款催收款呆帳處理辦法」第3條規定予以評估分類。</font></font></li>
							<li class="sbody" >
                           <font color="red">【註2】帳列備抵呆帳，應與資產負債表之「備抵呆帳-放款」、「 備抵呆帳-催收款項」一致；「建築貸款備抵呆帳」係依主管機關之控管措施專款提撥，包含於前開帳列備抵呆帳餘額。</font></font></li>
							<li class="sbody" >
                           <font color="red">【註3】依規定應提列最低標準備之備抵呆帳，係依「農會漁會信用部資產評估損失準備提列及逾期放款催收款呆帳處理辦法」第4條第1項規定計算，由應予評估放款之申報資料自動代入產生之。</font></font></li>
							<li class="sbody" >
                           <font color="red">【註4】備抵呆帳是否依規定提足，係依「農會漁會信用部資產評估損失準備提列及逾期放款催收款呆帳處理辦法」第4條第1項規定計算；提列不足者，依同條第2項規定，應於1年內提足(即於104年底前提足)，或有正當理由，得申請展延，展延期限不得超過4年，請申報展延期限及分年提撥金額。</font></font></li>
						  <li class="sbody" ><font color="#FF0000">如果沒有申報仍請填「0」</font></li>   
						  <li class="sbody" ><font color="#FF0000">若上月資料已有申報時,自動代入上月資料以供修改後,儲存</font></li>                                                 
                          <li>確認資料無誤後，按<font color="#666666">【確定】</font>即將本網頁上的資料，於資料庫中新增。</li>
                          <li>按<font color="#666666">【取消】</font>即重新輸入資料。</li>
                          <li>點選所列之<font color="#666666">【回上一頁】</font>則放棄資料， 回至前一畫面。</li>                          
                        </ul></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><div align="center"><img src="images/line_1.gif" width=<%=request.getAttribute("table_width") %> height="12"></div></td>
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
<script language="JavaScript" >
<!--
addA10(this.document.forms[0]);
if('<%=baddebt_flag%>'=='N'){
	document.forms[0].baddebt_flag[1].checked=true;
	document.forms[0].baddebt_noenough.value='<%=baddebt_noenough%>';
	if('<%=baddebt_delay%>'=='N'){
		document.forms[0].baddebt_delay[0].checked=true;
		document.forms[0].baddebt_104.value='<%=baddebt_104%>';
	}
	if('<%=baddebt_delay%>'=='Y'){
		document.forms[0].baddebt_delay[1].checked=true;
		if('<%=baddebt_105%>'!=''  && '<%=baddebt_105%>'!='0'){
			document.forms[0].c5.checked=true; 
			document.forms[0].baddebt_105.value='<%=baddebt_105%>';
		}
		if('<%=baddebt_106%>'!=''  && '<%=baddebt_106%>'!='0'){
			document.forms[0].c6.checked=true;
			document.forms[0].baddebt_106.value='<%=baddebt_106%>';
		}
		if('<%=baddebt_107%>'!=''  && '<%=baddebt_107%>'!='0'){
			document.forms[0].c7.checked=true;
			document.forms[0].baddebt_107.value='<%=baddebt_107%>';
		}
		if('<%=baddebt_108%>'!=''  && '<%=baddebt_108%>'!='0'){
			document.forms[0].c8.checked=true;
			document.forms[0].baddebt_108.value='<%=baddebt_108%>';
		}
	}
	
}
-->
</script>

</body>

</html>
