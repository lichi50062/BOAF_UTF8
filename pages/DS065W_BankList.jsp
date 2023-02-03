<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>

<%
	// 查詢條件值 
    Map dataMap =Utility.saveSearchParameter(request);
	String report_no = "DS065W";
	String title = Utility.getPgName(report_no);
	String act = Utility.getTrimString(dataMap.get("act"));		
	
	//營運中/已裁撤===================================================================================
	String cancel_no = ( session.getAttribute("CANCEL_NO")==null ) ? "N" : (String)session.getAttribute("CANCEL_NO");				
	//========================================================================================================	
	String szExcelAction = (session.getAttribute("excelaction")==null)?"download":(String)session.getAttribute("excelaction");		
	
	String hsien_id = ( session.getAttribute("HSIEN_ID")==null ) ? "ALL" : (String)session.getAttribute("HSIEN_ID");				
	System.out.println(report_no+"_BankList.hsien_id="+hsien_id);
	String YEAR  = Utility.getYear();
   	String MONTH = Utility.getMonth();
	String S_YEAR = (session.getAttribute("S_YEAR")==null)?YEAR:(String)session.getAttribute("S_YEAR");
	String S_MONTH = (session.getAttribute("S_MONTH")==null)?MONTH:(String)session.getAttribute("S_MONTH");            
   	String Unit = (session.getAttribute("Unit")==null)?"":(String)session.getAttribute("Unit");   	   	
	
	//95.11.10 取得登入者資訊=================================================================================================
	String muser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		
    String muser_bank_type = ( session.getAttribute("bank_type")==null ) ? "" : (String)session.getAttribute("bank_type");			
    //==============================================================================================================    	    
    String bank_type = (session.getAttribute("nowbank_type")==null)?"6":(String)session.getAttribute("nowbank_type");	    
	String DS_bank_type = (session.getAttribute("DS_bank_type")==null)?"6":(String)session.getAttribute("DS_bank_type");	
	System.out.print("nowbank_type="+(String)session.getAttribute("nowbank_type"));
	System.out.print(report_no+"_BankList.szExcelAction="+szExcelAction);
    System.out.print(":S_YEAR="+S_YEAR+":S_MONTH="+S_MONTH);
	System.out.println(":bank_type="+bank_type);
	/*
	農(漁)會別:農會/漁會/農漁會
	可選擇項目:顯示各農漁會信用部名稱;農會顯示所有農會信用部/漁會顯示所有漁會信用部/農漁會顯示所有農漁會信用部名稱
	*/
	//控制農漁會清單1 
	session.setAttribute("nowbank_type","");
	session.setAttribute("muser_id","");
	session.setAttribute("bank_type","3");
%>
<%@include file="./include/DS_bank_no_hsien_id.include" %>
<%  //控制農漁會清單2 
	session.setAttribute("nowbank_type", bank_type); 
	session.setAttribute("muser_id", muser_id);
	session.setAttribute("bank_type", bank_type);
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/DSUtil.js"></script>
<script language="javascript" src="js/movesels.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<link href="css/b51.css" rel="stylesheet" type="text/css">
</head>
<script language="JavaScript" type="text/JavaScript">
<!--
//102.11.11 add 漁會科目代號新舊無法同時列印
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
      if(!confirm("本項報表會執行10-15秒，是否確定執行？")){
         return;
      }   
   }   
   
   MoveSelectToBtn(this.document.forms[0].BankList, this.document.forms[0].BankListDst);	
   fn_ShowPanel(report_no,cnd);      
}

//-->
</script>
<body leftmargin="0" topmargin="0">
<form method=post action='#' name='BankListfrm'>
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
                <td width="240"><img src="images/banner_bg1.gif" width="240" height="17"></td>
                <td width="*" class="title_font"><%=title %></td>
                <td width="240"><img src="images/banner_bg1.gif" width="240" height="17"></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td><img src="images/space_1.gif" width="8" height="8"></td>
        </tr>
        <tr>          
          <td><table width="750" border="1" align="center" cellpadding="0" cellspacing="0" class="bordercolor">
              <tr> 
                <td bordercolor="#E9F4E3" bgcolor="#E9F4E3">
                <table width="750" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#E9F4E3">
                    <tr> 
                      <td class="bt_bgcolor"> 
                       <div align="right">
                          <!--input type='radio' name="excelaction" value='view' <%if(szExcelAction.equals("view")){out.print("checked");}%> >檢視報表-->
                      	  <input type='radio' name="excelaction" value='download' <%if(szExcelAction.equals("download")){out.print("checked");}%> >下載報表
                      	  <%if(Utility.getPermission(request,report_no,"P")){//Print--有列印權限時 %> 
                      	  <a href="javascript:doSubmit('<%=report_no%>','createRpt');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image41','','images/bt_execb.gif',1)"><img src="images/bt_exec.gif" name="Image41" width="66" height="25" border="0" id="Image41"></a> 
                      	  <%}%>
                          <a href="javascript:ResetAllData('BankList');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image51','','images/bt_cancelb.gif',0)"><img src="images/bt_cancel.gif" name="Image51" width="66" height="25" border="0" id="Image51"></a> 
                          <a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image61','','images/bt_reporthelpb.gif',1)"><img src="images/bt_reporthelp.gif" name="Image61" width="80" height="25" border="0" id="Image61"></a> 
                        </div>
                       </td>
                    </tr>
                    <tr> 
                      <td class="menu_bgcolor"> 
                        <table width="700" border="0" align="center" cellpadding="1" cellspacing="1" class="sbody">
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
                         
                    	 <!--查詢年月-->
                         <tr class="sbody">
                           <td><img src="images/2_icon_01.gif" width="16" height="16" align="absmiddle"><span class="mtext">查詢年月 :</span> 						  						
                              <input type='text' name='S_YEAR' value="<%=S_YEAR%>" size='3' maxlength='3' onblur='CheckYear(this)' onchange="javascript:changeCity('CityXML', this.document.forms[0].HSIEN_ID, this.document.forms[0].S_YEAR, this.document.forms[0]);changeOption(document.forms[0],'change');"><font color='#000000'>年                             
                         		<select id="hide1" name=S_MONTH>        						
                         		<%for(int j = 1; j <= 12; j++) {
                         			if (j < 10){%>        	
                         				<option value=0<%=j%> <%if(String.valueOf(Integer.parseInt(S_MONTH)).equals(String.valueOf(j))) out.print("selected");%>>0<%=j%></option>        		
                         			<%}else{%>
                         				<option value=<%=j%> <%if(String.valueOf(Integer.parseInt(S_MONTH)).equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
                         			<%}%>
                         		<%}%>
                         		</select><font color='#000000'><%out.print("月");%></font>
                         </tr> 
                         <!--農(漁)會別-->
                         <tr class="sbody">
                         	<td><img src="images/2_icon_01.gif" width="16" height="16" align="absmiddle"><span class="mtext">農(漁)會別 :</span>
                             	<select size="1" name="bank_type" onchange="javascript:changeOption(document.forms[0],'change');">                                                                                        
                             	<%//if(DS_bank_type.equals("6")){//95.11.10 有農會的menu時,才可顯示農會%>
                             		<option value ='6' <%if((!bank_type.equals("")) && bank_type.equals("6")) out.print("selected");%>>農會</option>                                                            
                             	<%//}%>
                             	<%//if(DS_bank_type.equals("7")){//95.11.10 有漁會的menu時,才可顯示漁會%>
                             		<option value ='7' <%if((!bank_type.equals("")) && bank_type.equals("7")) out.print("selected");%>>漁會</option>                              
                             	<%//}%>
                             	<%//if(!bank_type.equals("") && (muser_bank_type.equals("2") || muser_id.equals("A111111111"))){
                             		//95.11.10 登入者為A111111111 or 農金局時,才可顯示農漁會%>                              
                             		<option value ='ALL' <%if((!bank_type.equals("")) && bank_type.equals("ALL")) out.print("selected");%>>農漁會</option>                              
                             	<%//}%>
                             </select>
                         	</td>
                         </tr>  
                         <%@include file="./include/DS_Unit.include" %><!-- 金額單位-->
                         <%@include file="./include/DS_Cancel_No_Hsien_ID.include" %><!-- 1.營運中/裁撤別 2.縣市別-->    
                          
                        </table>
                       </td>
                    </tr>
                    
                    <%@include file="./include/DS_BankList.include" %><!-- 可選擇項目-->
                    
                    <tr> 
                 	  <td class="body_bgcolor">
                  	    <table width="750" border="0" cellpadding="1" cellspacing="1">
                    	<tr>                           
                          <td width="750" align=left><font color="red" size=2>註：『已選擇項目』清單選取之項目說明</font></td>                              
                    	</tr>                              
                        <tr>                           
                          <td width="750" align=left><font color="red" size=2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(1)選取[各信用部]：按「<%=title %>」的報表內容列印</font></td>                              
                        </tr>                                                  
                        </table>
                      </td>
                    </tr>
                  </table></td>
              </tr>              
            </table></td>
        </tr>
        
      </table>
    </td>
  </tr>
</table>
<INPUT type="hidden" name=BankList ><!--//BankList儲存已勾選的金融機構代碼-->
<INPUT type="hidden" name=btnFieldList value='<%if(session.getAttribute("btnFieldList") != null) out.print((String)session.getAttribute("btnFieldList"));%>'><!--//btnFieldList儲存已勾選的報表欄位名稱-->
<INPUT type="hidden" name=loan_item value='<%if(session.getAttribute("loan_item") != null) out.print((String)session.getAttribute("loan_item"));%>'>
<INPUT type="hidden" name=clearbtnFieldList><!--//儲存是否清除btnFieldList-->
<INPUT type="hidden" name=agri_loan value="true"><!--//專案農貸註記-->
</form>
<script language="JavaScript" >
<!--

<%
//從session裡把勾選的金融機構代碼讀出來.放在BankListDst
if(session.getAttribute("BankList") != null && !((String)session.getAttribute("BankList")).equals("")){ 
   System.out.println(report_no+"_BankList.BankList="+(String)session.getAttribute("BankList"));
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


-->
</script>

</body>
</html>
