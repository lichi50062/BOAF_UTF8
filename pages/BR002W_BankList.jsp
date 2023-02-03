<%
// 94.01.20 create by 2295
// 94.01.31 add 權限 by 2295
// 94.03.23 add 營運中/已裁撤 by 2295
// 94.09.05 fix 拿掉檢視報表 by 2295
// 99.12.14 fix SQLInjection by 2479
// 99.12.21 fix 根據查詢年度.100年以後取得新縣市別.100年以前取得舊縣市別 by 2295
//          fix 查詢年月/可選擇項目 套用共用include by 2295
// 101.06   add 報表欄位 by 2968
//103.01.21 add BOAF/MIS共用畫面 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%
    Map dataMap =Utility.saveSearchParameter(request);
	String report_no = "BR002W";	
	String act = Utility.getTrimString(dataMap.get("act"));			
	
	//94.03.23 add 營運中/已裁撤===================================================================================
	String cancel_no = ( session.getAttribute("CANCEL_NO")==null ) ? "N" : (String)session.getAttribute("CANCEL_NO");				
	//========================================================================================================
	String bank_type = (session.getAttribute("nowbank_type")==null)?"":(String)session.getAttribute("nowbank_type");	
	String DS_bank_type = (session.getAttribute("DS_bank_type")==null)?"6":(String)session.getAttribute("DS_bank_type");	
	String szExcelAction = (session.getAttribute("excelaction")==null)?"download":(String)session.getAttribute("excelaction");
	//System.out.println("BR002W_BankList.szExcelAction="+szExcelAction);
	String hsien_id = ( session.getAttribute("HSIEN_ID")==null ) ? "ALL" : (String)session.getAttribute("HSIEN_ID");				
	System.out.println("BR002W_BankList.hsien_id="+hsien_id);
	String YEAR  = Utility.getYear();
   	String MONTH = Utility.getMonth();
	String S_YEAR = (session.getAttribute("S_YEAR")==null)?YEAR:(String)session.getAttribute("S_YEAR");	
	String S_MONTH = (session.getAttribute("S_MONTH")==null)?MONTH:(String)session.getAttribute("S_MONTH");
	String bank_name = (bank_type.equals("6"))?"農會":"漁會";
	if(DS_bank_type.equals("B")) bank_name="農漁會";
%>

<%@include file="./include/BR_bank_no_hsien_id.include" %>

<script language="javascript" src="js/Common.js"></script>
<!--script language="javascript" src="js/BR002W.js"></script-->
<!--script language="javascript" src="js/BRUtil.js"></script-->
<script language="javascript" src="js/DSUtil.js"></script>
<script language="javascript" src="js/movesels.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<script language="JavaScript" type="text/JavaScript">
<!--
function doSubmit(report_no,cnd){
   if(cnd == 'createRpt'){      
      if(this.document.forms[0].BankListDst.length == 0){      	 
      	 alert('金融機構代碼必須選擇');
      	 return;
      }
      if(this.document.forms[0].btnFieldList.value == ''){
         alert('報表欄位必須選擇');
         return;
      }
   }   
   
   MoveSelectToBtn(this.document.forms[0].BankList, this.document.forms[0].BankListDst);	
   fn_ShowPanel(report_no,cnd);      
}
//-->
</script>
<link href="css/b51.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0">
<form method=post action='#' name='BankListfrm'>
<INPUT type="hidden" name=bank_type value=<%=bank_type%>>
<table width="750" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
     <td>&nbsp;</td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF">
	<table width="750" border="0" align="center" cellpadding="1" cellspacing="1">        
        <tr> 
          <td><table width="750" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="225"><img src="images/banner_bg1.gif" width="225" height="17"></td>
                <td width="300" class="title_font"><%=bank_name+Utility.getPgName(report_no)%></td>
                <td width="225"><img src="images/banner_bg1.gif" width="225" height="17"></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td><img src="images/space_1.gif" width="8" height="8"></td>
        </tr>
        <tr> 
          <td><table width="750" border="1" align="center" cellpadding="0" cellspacing="0" class="bordercolor">
              <tr> 
                <td bordercolor="#E9F4E3" bgcolor="#E9F4E3"><table width="750" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#E9F4E3">
                    <tr> 
                      <td class="bt_bgcolor"> 
                          <div align="right">                          
                      	  <input type='radio' name="excelaction" value='download' <%if(szExcelAction.equals("download")){out.print("checked");}%> >下載報表
                      	  <%if(Utility.getPermission(request,report_no,"P")){//Print--有列印權限時 %> 
                      	  <a href="javascript:doSubmit('<%=report_no%>','createRpt');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image41','','images/bt_execb.gif',1)"><img src="images/bt_exec.gif" name="Image41" width="66" height="25" border="0" id="Image41"></a> 
                      	  <%}%>
                          <a href="javascript:ResetAllData_BR('BankList');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image51','','images/bt_cancelb.gif',0)"><img src="images/bt_cancel.gif" name="Image51" width="66" height="25" border="0" id="Image51"></a> 
                          <a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image61','','images/bt_reporthelpb.gif',1)"><img src="images/bt_reporthelp.gif" name="Image61" width="80" height="25" border="0" id="Image61"></a> 
                        </div></td>
                    </tr>
                    <tr> 
                      <td class="menu_bgcolor"> 
                      <table width="750" border="0" align="center" cellpadding="1" cellspacing="1" class="sbody">
                          <tr class="sbody"> 
                            <td width="100"><img src="images/2_icon_01.gif" width="16" height="16" align="absmiddle"> 
                              <a href="#"><font color="#CC6600">1.金融機構</font></a></td>
                            <td width="100"><a href="javascript:doSubmit('<%=report_no%>','RptColumn')"><font color='black'>2.報表欄位</font></a></td>                            
                            <td width="100"><a href="javascript:doSubmit('<%=report_no%>','RptOrder')"><font color='black'>3.排序欄位</font></a></td>
                            <td width="100"><a href="javascript:doSubmit('<%=report_no%>','RptStyle')"><font color='black'>4.報表格式</font></a></td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr> 
                      <td class="body_bgcolor"> 
                        <table width="750" border="0" align="center" cellpadding="0" cellspacing="0">
                          <tr class="sbody">
                            <td><img src="images/2_icon_01.gif" width="16" height="16" align="absmiddle"><span class="mtext">查詢年月 :</span> 						  						
                               <input type='text' name='S_YEAR' value="<%=S_YEAR%>" size='3' maxlength='3' onblur='CheckYear(this)' onchange="javascript:changeCity('CityXML', this.document.forms[0].HSIEN_ID, this.document.forms[0].S_YEAR, this.document.forms[0]);changeOption(document.forms[0],'change');"><font color='#000000'>年                             
                          		<select id="hide1" name=S_MONTH>        						
                          		<%
                          			for (int j = 1; j <= 12; j++) {			
                          			if (j < 10){%>        	
                          			<option value=0<%=j%> <%if(String.valueOf(Integer.parseInt(S_MONTH)).equals(String.valueOf(j))) out.print("selected");%>>0<%=j%></option>        		
                          			<%}else{%>
                          			<option value=<%=j%> <%if(String.valueOf(Integer.parseInt(S_MONTH)).equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
                          			<%}%>
                          		<%}%>
                          		</select><font color='#000000'>月</font>
                          		<input type=hidden name=S_DATE value=''>
                          		
                              </td>
                          </tr> 
                          <%@include file="./include/DS_Cancel_No_Hsien_ID.include" %><!-- 1.營運中/裁撤別 2.縣市別-->                              
                        </table>
                      </td>
                    </tr>
                    
                    <%@include file="./include/DS_BankList.include" %><!-- 可選擇項目-->
                    
                    
                  </table></td>
              </tr>
            </table></td>
        </tr>
        
      </table>
    </td>
  </tr>
</table>
<INPUT type="hidden" name=BankList><!--//BankList儲存已勾選的金融機構代碼-->
<INPUT type="hidden" name=btnFieldList value='<%if(session.getAttribute("btnFieldList") != null) out.print((String)session.getAttribute("btnFieldList"));%>'><!--//btnFieldList儲存已勾選的報表欄位名稱-->
</form>
<script language="JavaScript" >
<!--

<%
//從session裡把勾選的金融機構代碼讀出來.放在BankListDst
if(session.getAttribute("BankList") != null && !((String)session.getAttribute("BankList")).equals("")){ 
   System.out.println("BR002W_BankList.BankList="+(String)session.getAttribute("BankList"));
%>
var bnlist;
bnlist = '<%=(String)session.getAttribute("BankList")%>';
var a = bnlist.split(',');
for (var i =0; i < a.length; i ++){
	var j = a[i].split('+');
	this.document.forms[0].BankListDst.options[i] = new Option(j[1], j[0]);
}
<%}%>

setSelect(this.document.forms[0].HSIEN_ID,"<%=hsien_id%>");
setSelect(this.document.forms[0].CANCEL_NO,"<%=cancel_no%>");
changeCity('CityXML', this.document.forms[0].HSIEN_ID, this.document.forms[0].S_YEAR, this.document.forms[0]);
//changeOption(this.document.forms[0],'');
function clearBankList(){
 <%
	session.setAttribute("BankList",null);//清除已勾選的BankList
 %>
}
-->
</script>

</body>
</html>
