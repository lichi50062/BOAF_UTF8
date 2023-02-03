<%
//99.07.30 create 信用部主任及分部主任參訓情形彈性報表 by 2660
//99.10.19 fix 更名為信用部從業人員參加金融相關業務進修情形明細資料 by 2295
//102.11.05 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Calendar" %>
<%
	String act = ( request.getParameter("act") == null ) ? "" : (String)request.getParameter("act");
	//營運中/已裁撤===================================================================================
	String cancel_no = ( session.getAttribute("CANCEL_NO") == null ) ? "N" : (String)session.getAttribute("CANCEL_NO");
	//========================================================================================================
	String bank_type = (session.getAttribute("nowbank_type") == null)?"6":(String)session.getAttribute("nowbank_type");
	//String title=(bank_type.equals("6"))?"農會":"漁會";
	String szExcelAction = (session.getAttribute("excelaction") == null)?"download":(String)session.getAttribute("excelaction");
	System.out.println("DS020W_BankList.szExcelAction=" + szExcelAction);
	String hsien_id = ( session.getAttribute("HSIEN_ID") == null ) ? "ALL" : (String)session.getAttribute("HSIEN_ID");
	System.out.println("DS020W_BankList.hsien_id=" + hsien_id);
	String YEAR = Utility.getYear();
  String MONTH = Utility.getMonth();
	String S_YEAR = (session.getAttribute("S_YEAR")==null)?YEAR:(String)session.getAttribute("S_YEAR");
	String E_YEAR = (session.getAttribute("E_YEAR")==null)?YEAR:(String)session.getAttribute("E_YEAR");
	String S_MONTH = (session.getAttribute("S_MONTH")==null)?MONTH:(String)session.getAttribute("S_MONTH");            
	String E_MONTH = (session.getAttribute("E_MONTH")==null)?MONTH:(String)session.getAttribute("E_MONTH");                              
  String Unit = (session.getAttribute("Unit")==null)?"":(String)session.getAttribute("Unit");
  //95.11.14 取得登入者資訊=================================================================================================
  String muser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		
  String muser_bank_type = ( session.getAttribute("bank_type")==null ) ? "" : (String)session.getAttribute("bank_type");			
  String DS_bank_type = (session.getAttribute("DS_bank_type")==null)?"6":(String)session.getAttribute("DS_bank_type");	
  //==============================================================================================================    	    
  	 	
	System.out.println("S_MONTH=" + S_MONTH);
	System.out.println("bank_type=" + bank_type);
	System.out.println("DS_bank_type=" + DS_bank_type);
			
  List tbankList = Utility.getBankList(request);//95.11.14 add 可選機構代號權限設定(農漁會)
    
    
  //取得DS020W的權限
	Properties permission = ( session.getAttribute("DS020W")==null ) ? new Properties() : (Properties)session.getAttribute("DS020W"); 
	if (permission == null) {
    System.out.println("DS020W_BankList.permission == null");
  } else {
    System.out.println("DS020W_BankList.permission.size ="+permission.size());
  }	
    
  // XML Ducument for 總機構代碼 begin
  out.println("<xml version=\"1.0\" encoding=\"big5\" ID=\"TBankXML\">");
  out.println("<datalist>"); 
  /*95.10.03 add A02違反法定比率.沒有全部的資料
  if (muser_id.equals("A111111111") || muser_bank_type.equals("2")) { //95.11.10 登入者為A111111111 or 農金局時,才可選全部
    out.println("<data>");        
    out.println("<BnType>1</BnType>");
    out.println("<BankType>6</BankType>");
    out.println("<HsienId>ALL</HsienId>");
    out.println("<bankValue>ALL</bankValue>");
    out.println("<bankName>全部</bankName>");
    out.println("</data>");      
    out.println("<data>");        
    out.println("<BnType>1</BnType>");
    out.println("<BankType>7</BankType>");
    out.println("<HsienId>ALL</HsienId>");
    out.println("<bankValue>ALL</bankValue>");
    out.println("<bankName>全部</bankName>");
    out.println("</data>");      
    out.println("<data>");        
    out.println("<BnType>1</BnType>");
    out.println("<BankType>ALL</BankType>");
    out.println("<HsienId>ALL</HsienId>");
    out.println("<bankValue>ALL</bankValue>");
    out.println("<bankName>全部</bankName>");
    out.println("</data>");      
  }
  */

  if (tbankList != null) {
    for(int i=0;i< tbankList.size(); i++) {
      DataObject bean =(DataObject)tbankList.get(i);
      out.println("<data>");        
      out.println("<BnType>" + bean.getValue("bn_type") + "</BnType>");
      out.println("<BankType>" + bean.getValue("bank_type") + "</BankType>");
      out.println("<HsienId>" + bean.getValue("hsien_id") + "</HsienId>");
      out.println("<bankValue>" + bean.getValue("bank_no") + "</bankValue>");
      out.println("<bankName>" + bean.getValue("bank_no") + bean.getValue("bank_name") + "</bankName>");
      out.println("</data>");
      //System.out.println("<option>" + bean.getValue("bank_no") + "&nbsp;"+bean.getValue("bank_name") + "</option>");
    }
  }
  out.println("</datalist>\n</xml>");
  // XML Ducument for 總機構代碼 end 	
%>

<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/DSUtil.js"></script>
<script language="javascript" src="js/movesels.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
  <head>
    <script language="JavaScript" type="text/JavaScript">
<!--
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
  var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
  if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
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

function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}

function doSubmit(cnd){
  if (cnd == 'createRpt') {
    if (this.document.forms[0].BankListDst.length == 0) {      	 
      alert('金融機構代碼必須選擇');
      return;
    }
    if (!chInput(this.document.forms[0],'BankList')) return;//95.12.06 add 起始日期不可大於結束日期
    if (this.document.forms[0].btnFieldList.value == '') {
      alert('報表欄位必須選擇');
      return;
    }
    if (!confirm("本項報表會執行10-15秒，是否確定執行？")) {
      return;
    }   
  }
  MoveSelectToBtn(this.document.forms[0].BankList, this.document.forms[0].BankListDst);	
  fn_ShowPanel(cnd);      
}

function fn_ShowPanel(cnd) {   
	//act=BankList/RptColumn/RptOrder/RptType	
	this.document.forms[0].action = "/pages/DS020W.jsp?act=" + cnd;
  this.document.forms[0].target = '_self';
	this.document.forms[0].submit();
}

function ResetAllData(){
  if (confirm("確定要清除已選定的資料嗎？")) {  	
    this.document.forms[0].BankListDst.length = 0;
    this.document.forms[0].HSIEN_ID[0].selected = true;
    changeOption(this.document.forms[0],'');
    //clearBankList();95.12.07
  }
	return;	
}
//-->
    </script>
    <link href="css/b51.css" rel="stylesheet" type="text/css">
  </head>

  <body leftmargin="0" topmargin="20">
    <form method=post action='#' name='BankListfrm'>
      <center>
      	<div style="vertical-align:middle;display:inline"><img src="images/banner_bg1.gif" width="160" height="17"></div>
      	<div style="margin:10px;vertical-align:middle;display:inline"><font color="#336600" size=4>信用部從業人員參加金融相關業務進修情形明細資料</font></div>
      	<div style="vertical-align:middle;display:inline"><img src="images/banner_bg1.gif" width="160" height="17"></div>
      	<div><img src="images/space_1.gif" width="8" height="8"></div>
        <div style="width=750;border-width:1px;border-color:#5DA525;border-style:solid;padding:1px;">
          <table width="750" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#E9F4E3">
            <tr>
              <td class="bt_bgcolor" align="right">
              	<div style="display:inline">
                  <!--input type='radio' name="excelaction" value='view' <%if(szExcelAction.equals("view")){out.print("checked");}%> >檢視報表-->
                  <input type='radio' name="excelaction" value='download' <% if (szExcelAction.equals("download")) { out.print("checked"); } %> />下載報表
                </div>
                <div style="vertical-align:middle;display:inline">
<%
  if (permission != null && permission.get("P") != null && permission.get("P").equals("Y")) { //Print
%>                   	        	                                   		     			        
                  <a href="javascript:doSubmit('createRpt');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image41','','images/bt_execb.gif',1)"><img src="images/bt_exec.gif" name="Image41" width="66" height="25" border="0" id="Image41"></a> 
<%}%>
                  <a href="javascript:ResetAllData();" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image51','','images/bt_cancelb.gif',0)"><img src="images/bt_cancel.gif" name="Image51" width="66" height="25" border="0" id="Image51"></a> 
                  <a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image61','','images/bt_reporthelpb.gif',1)"><img src="images/bt_reporthelp.gif" name="Image61" width="80" height="25" border="0" id="Image61"></a> 
                </div>
              <td>
            </tr>
            <tr>
              <td align="center" class="menu_bgcolor">
                <div style="vertical-align:middle;padding:3px 0px 3px 0px"">
                  <div style="margin:50px;vertical-align:middle;display:inline"><img src="images/2_icon_01.gif" width="16" height="16" align="absmiddle"><a href="#"><font color="#CC6600">1.金融機構</font></a></div>
                  <div style="margin:50px;vertical-align:middle;display:inline"><a href="javascript:doSubmit('RptColumn')"><font color='black'>2.報表欄位</font></div>
                  <div style="margin:50px;vertical-align:middle;display:inline"><a href="javascript:doSubmit('RptOrder')"><font color='black'>3.排序欄位</font></a></div>
                  <div style="margin:50px;vertical-align:middle;display:inline"><a href="javascript:doSubmit('RptStyle')"><font color='black'>4.報表格式</font></a></div>
                </div>
              </td>
            </tr>
            <tr>
              <td>
              	<div style="vertical-align:middle;padding-top:2px">
                  <img src="images/2_icon_01.gif" width="16" height="16" align="absmiddle"><span class="mtext">查詢年月 :</span> 						  						
                  <input type='text' name='S_YEAR' value="<%=S_YEAR%>" size='3' maxlength='3' onblur='CheckYear(this)'><font color='#000000'>年                            
                  <select id="hide1" name=S_MONTH>        						
<%
  for (int j = 1; j <= 12; j++) {
    if (j < 10) {%>        	
                    <option value=0<%=j%> <%if(String.valueOf(Integer.parseInt(S_MONTH)).equals(String.valueOf(j))) out.print("selected");%>>0<%=j%></option>        		
<%  } else {%>
                    <option value=<%=j%> <%if(String.valueOf(Integer.parseInt(S_MONTH)).equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
<%  }
  }%>
                  </select><font color='#000000'>月</font>~
                  <input type='text' name='E_YEAR' value="<%=E_YEAR%>" size='3' maxlength='3' onblur='CheckYear(this)'><font color='#000000'>年	
                  <select id="hide1" name=E_MONTH>
<%
  for (int j = 1; j <= 12; j++) {
    if (j < 10) {%>        	
                    <option value=0<%=j%> <%if(String.valueOf(Integer.parseInt(E_MONTH)).equals(String.valueOf(j))) out.print("selected");%>>0<%=j%></option>        		
<%  } else {%>
                    <option value=<%=j%> <%if(String.valueOf(Integer.parseInt(E_MONTH)).equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
<%  }
  }%>
                  </select><font color='#000000'>月</font>                            
                  <input type=hidden name=S_DATE value=''>
                  <input type=hidden name=E_DATE value=''>
                </div>
              </td>
            </tr>
            <tr>
              <td>
              	<div style="vertical-align:middle;padding-top:2px">
                  <img src="images/2_icon_01.gif" width="16" height="16" align="absmiddle"><span class="mtext">農(漁)會別 :</span>
                  <select size="1" name="bank_type" onchange="javascript:changeOption(document.forms[0],'change');">
<%
  if (DS_bank_type.equals("6")) { //95.11.10 有農會的menu時,才可顯示農會
%>
                    <option value ='6' <% if ((!bank_type.equals("")) && bank_type.equals("6")) out.print("selected"); %> >農會</option>
<%}
  if (DS_bank_type.equals("7")) { //95.11.10 有漁會的menu時,才可顯示漁會
%>
                    <option value ='7' <% if ((!bank_type.equals("")) && bank_type.equals("7")) out.print("selected"); %> >漁會</option>
<%}
  if (!bank_type.equals("") && (muser_bank_type.equals("2") || muser_id.equals("A111111111"))) {
//95.11.10 登入者為A111111111 or 農金局時,才可顯示農漁會 %>
                    <option value ='ALL' <%if((!bank_type.equals("")) && bank_type.equals("ALL")) out.print("selected");%>>農漁會</option>
<%}%>
                  </select>
                </div>
              </td>
            </tr>
            <tr>
              <td>
              	<div style="vertical-align:middle;padding-top:2px">
                  <img src="images/2_icon_01.gif" width="16" height="16" align="absmiddle"><span class="mtext">營運中/裁撤別 :</span>
                  <select name='CANCEL_NO' onchange="javascript:changeOption(document.forms[0],'change');">
                    <option  value="N" <%if((!cancel_no.equals("")) && cancel_no.equals("N")) out.print("selected");%>>營運中</option>
                    <option  value="Y" <%if((!cancel_no.equals("")) && cancel_no.equals("Y")) out.print("selected");%>>已裁撤</option>
                  </select>
                </div>
              </td>
            </tr>
<%
  List hsien_id_data = DBManager.QueryDB_SQLParam("select distinct hsien_id,hsien_name from cd01",null,""); 
%>
            <tr>
              <td>
              	<div style="vertical-align:middle;padding-top:2px">
                  <img src="images/2_icon_01.gif" width="16" height="16" align="absmiddle"><span class="mtext">縣市別 :</span>                                
                  <select name="HSIEN_ID" onchange="javascript:changeOption(document.forms[0],'');">                               
                    <option value="ALL">全部</option>
<%
  for(int i=0; i < hsien_id_data.size(); i++) {
%>                                
                    <option value="<%=(String)((DataObject)hsien_id_data.get(i)).getValue("hsien_id")%>" <% if (((String)((DataObject)hsien_id_data.get(i)).getValue("hsien_id")).equals(hsien_id)) out.print("selected"); %> >
<%=(String)((DataObject)hsien_id_data.get(i)).getValue("hsien_name")%>
                    </option>                            
<%}%>
                </div>
              </td>
            </tr>
            <tr>
              <td>
              	<div style="vertical-align:middle;padding-top:2px">
              	  <table>
              	  	<tr>
              	  	  <td align="center" class="chooseitem_bgcolor">可選擇項目</td>
              	  	  <td></td>
              	  	  <td align="center" class="chooseitem_bgcolor">已選擇項目</td>
              	    </tr>
              	    <tr>
              	      <td>
                        <select multiple size=10 name="BankListSrc" ondblclick="javascript:movesel(this.document.forms[0].BankListSrc,this.document.forms[0].BankListDst);fn_changeBankListSrc(document.forms[0]);" style="width:294;height:190"></select>
              	      </td>
              	      <td width='100px' align="center">
                        <div style="padding:5px 0px 5px 0px">                                 
                          <a href="javascript:movesel(this.document.forms[0].BankListSrc,this.document.forms[0].BankListDst);fn_changeBankListSrc(document.forms[0]);"><img src="images/arrow_right.gif" width="24" height="22" border="0"></a>
                        </div>
                        <div style="padding:5px 0px 5px 0px">                                 
                          <a href="javascript:moveallsel(this.document.forms[0].BankListSrc,this.document.forms[0].BankListDst);"><img src="images/arrow_rightall.gif" width="24" height="22" border="0"></a>
                        </div>
                        <div style="padding:5px 0px 5px 0px">                                 
                          <a href="javascript:movesel(this.document.forms[0].BankListDst,this.document.forms[0].BankListSrc);fn_changeBankListSrc(document.forms[0]);""><img src="images/arrow_left.gif" width="24" height="22" border="0"></a>
                        </div>
                        <div style="padding:5px 0px 5px 0px">                                 
                          <a href="javascript:moveallsel(this.document.forms[0].BankListDst,this.document.forms[0].BankListSrc);"><img src="images/arrow_leftall.gif" width="24" height="22" border="0"></a>
                        </div>
              	      </td>
              	      <td>
                        <select multiple size=10  name="BankListDst" ondblclick="javascript:movesel(this.document.forms[0].BankListDst,this.document.forms[0].BankListSrc);fn_changeBankListSrc(document.forms[0]);"" style="width:294;height:190"></select>
              	      </td>
              	    </tr>
              	  </table>
                </div>
              </td>
            </tr>
            <tr>
              <td>
              	<div style="vertical-align:middle">
              	  <font color="red" size=2>
              	    <div>註：『已選擇項目』清單選取之項目說明</div>
              	    <div>　　(1)選取[各信用部]：按「信用部主任及分部主任參訓情形」的報表內容列印</div>
                  </font>
                </div>
              </td>
            </tr>
          </table>
        </div>
    	</center>
      <INPUT type="hidden" name=BankList><!--//BankList儲存已勾選的金融機構代碼-->
      <INPUT type="hidden" name=btnFieldList value='<%if(session.getAttribute("btnFieldList") != null) out.print((String)session.getAttribute("btnFieldList"));%>'><!--//btnFieldList儲存已勾選的報表欄位名稱-->
      <INPUT type="hidden" name=clearbtnFieldList><!--//儲存是否清除btnFieldList-->
    </form>
    <script language="JavaScript" >
<!--
<%
  //從session裡把勾選的金融機構代碼讀出來.放在BankListDst
  if (session.getAttribute("BankList") != null && !((String)session.getAttribute("BankList")).equals("")) { 
    System.out.println("DS020W_BankList.BankList="+(String)session.getAttribute("BankList"));
%>
    var bnlist;
    bnlist = '<%=(String)session.getAttribute("BankList")%>';
    var a = bnlist.split(',');
    for (var i =0; i < a.length; i ++) {
	    var j = a[i].split('+');
	    this.document.forms[0].BankListDst.options[i] = new Option(j[1], j[0]);
    }
<%}%>

setSelect(this.document.forms[0].HSIEN_ID,"<%=hsien_id%>");
setSelect(this.document.forms[0].CANCEL_NO,"<%=cancel_no%>");
changeOption(this.document.forms[0],'');
    </script>
  </body>
</html>