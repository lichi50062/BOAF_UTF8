<%
//99.07.30 create 信用部從業人員參加金融相關業務進修情形明細資料彈性報表 by 2660
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.HashMap" %>
<%
  String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");					
  String szExcelAction = (session.getAttribute("excelaction")==null)?"download":(String)session.getAttribute("excelaction");
  System.out.println("DS020W_RptStyle.szExcelAction="+szExcelAction);
  String bank_type = (session.getAttribute("nowbank_type")==null)?"":(String)session.getAttribute("nowbank_type");
  //String title=(bank_type.equals("6"))?"農會":"漁會";
  //取得DS020W的權限
  Properties permission = ( session.getAttribute("DS020W")==null ) ? new Properties() : (Properties)session.getAttribute("DS020W"); 
  if (permission == null) {
    System.out.println("DS020W_RptStyle.permission == null");
  } else {
    System.out.println("DS020W_RptStyle.permission.size ="+permission.size());               
  }
  List templateList = null;
  if (request.getAttribute("templateList") != null) {
    templateList = (List)request.getAttribute("templateList");
    System.out.println("templateList.size()="+templateList.size());
  } else {
    System.out.println("templateList is null");
  }
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

function fn_ShowPanel(cnd) {   
  //act=BankList/RptColumn/RptOrder/RptType	
  this.document.forms[0].action = "/pages/DS020W.jsp?act="+cnd;	
  this.document.forms[0].target = '_self';
  this.document.forms[0].submit(); 	  
}
//-->
    </script>
    <link href="css/b51.css" rel="stylesheet" type="text/css">
  </head>

  <body leftmargin="0" topmargin="20">
    <form method=post action='#' name='RptOrderfrm'>
      <center>
      	<div style="vertical-align:middle;display:inline"><img src="images/banner_bg1.gif" width="180" height="17"></div>
      	<div style="margin:10px;vertical-align:middle;display:inline"><font color="#336600" size=4>信用部從業人員參加金融相關業務進修情形明細資料</font></div>
      	<div style="vertical-align:middle;display:inline"><img src="images/banner_bg1.gif" width="180" height="17"></div>
      	<div><img src="images/space_1.gif" width="8" height="8"></div>
        <div style="width=790;border-width:1px;border-color:#5DA525;border-style:solid;padding:1px;">
          <table width="790" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#E9F4E3">
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
                  <div style="margin:50px;vertical-align:middle;display:inline"><a href="javascript:doSubmit('BankList')"><font color='black'>1.金融機構</font></a></div>
                  <div style="margin:50px;vertical-align:middle;display:inline"><a href="javascript:doSubmit('RptColumn')"><font color='black'>2.報表欄位</font></a></div>
                  <div style="margin:50px;vertical-align:middle;display:inline"><a href="javascript:doSubmit('RptOrder')"><font color='black'>3.排序欄位</font></a></div>
                  <div style="margin:50px;vertical-align:middle;display:inline"><img src="images/2_icon_01.gif" width="16" height="16" align="absmiddle"><a href="#"><font color="#CC6600">4.報表格式</font></a></div>
                </div>
              </td>
            </tr>
            <tr>
              <td>
                <div style="vertical-align:middle;padding-top:2px">
                  <table>
                    <tr>
                      <td class="sbody">　</td>
                      <td class="sbody">範本名稱</td>
                      <td class="sbody">建置者</td>
                      <td class="sbody">建置日期</td>
                      <td class="sbody">　</td>
                    </tr>
                    <tr>
                      <td class="sbody" width="245px" style="vertical-align:top">
                      	<div class="report_sbody" style="vertical-align:middle;"><img src="images/2_icon_01.gif" width="16" height="16" align="absmiddle">請點選範本名稱</div>
                      	<div class="report_sbody" style="padding-left:16px">1.按【讀取格式檔】按鈕:<font color="#800000">載入格式檔</font></div>
                      	<div class="report_sbody" style="padding-left:16px">2.按【刪除】按鈕:<font color="red">刪除格式檔</font></div>
                      </td>
                      <td>
                        <select multiple size=6 name="template_list" style="width: 252; height: 114">							
<%if(templateList == null || templateList.size() == 0) { %>
                          <option></option>
<%} else {
    for(int i=0;i<templateList.size();i++) { %>
                          <option value='<%=(String)((HashMap)templateList.get(i)).get("template")%>:<%=(String)((HashMap)templateList.get(i)).get("createUser")%>:<%=(String)((HashMap)templateList.get(i)).get("updateDate")%>'><%=(String)((HashMap)templateList.get(i)).get("showTemplate")%></option>
<%  }//end of for
  }//end of templateList != null %> 
                        </select>
                      </td>
                      <td>
                        <select multiple  size=6  name="CreateUser_list" style="width: 80; height: 114" disabled>							
<%if(templateList == null || templateList.size() == 0) { %>
                          <option>&nbsp;</option>
<%} else {
    for (int i=0;i<templateList.size();i++) { %>
                          <option value='<%=(String)((HashMap)templateList.get(i)).get("template")%>:<%=(String)((HashMap)templateList.get(i)).get("createUser")%>:<%=(String)((HashMap)templateList.get(i)).get("updateDate")%>'><%=(String)((HashMap)templateList.get(i)).get("showCreateUser")%></option>
<%  }//end of for
  }//end of templateList != null
%>
                        </select>      
                      </td>
                      <td>
                        <select multiple  size=6  name="UpdateDate_list" style="width: 84; height: 114" disabled>			
<%if(templateList == null || templateList.size() == 0) { %>
                          <option>&nbsp;</option>
<%} else {
    for(int i=0;i<templateList.size();i++) { %>
                          <option value='<%=(String)((HashMap)templateList.get(i)).get("template")%>:<%=(String)((HashMap)templateList.get(i)).get("createUser")%>:<%=(String)((HashMap)templateList.get(i)).get("updateDate")%>'><%=(String)((HashMap)templateList.get(i)).get("showUpdateDate")%></option>
<%  }//end of for
  }//end of templateList != null
%> 
                        </select>
                      </td>
                      <td style="vertical-align:top">
                        <div><a href="javascript:doSubmit('ReadRpt');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image9','','images/bt_readreportb.gif',1)"><img src="images/bt_readreport.gif" name="Image9" width="86" height="25" border="0"></a></div>
                        <div style="padding-top:25px"><a href="javascript:doSubmit('DeleteRpt');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image91','','images/bt_deleteb.gif',1)"><img src="images/bt_delete.gif" name="Image91" width="66" height="25" border="0"></a></div>
                      </td>
                    </tr>
                    <tr>
                      <td class="sbody">
                      	<div class="report_sbody" style="vertical-align:middle;"><img src="images/2_icon_01.gif" width="16" height="16" align="absmiddle">請輸入欲儲存格式之名稱</div>
                      	<div class="report_sbody" style="padding-left:16px">(再按【儲存格式檔】按鈕)</div>
                      </td>
                      <td colspan="3"><input type="text" name="template" size="40"></td>
                      <td><a href="javascript:doSubmit('SaveRpt');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image92','','images/bt_savereportb.gif',1)"><img src="images/bt_savereport.gif" name="Image92" width="86" height="25" border="0"></a></td>
                    </td>
                  </table>
                </div>
              </td>
            </tr>
          </table>
        </div>
      </center>
      <INPUT type="hidden" name=SortList><!--//btnSortList儲存已勾選的報表排序欄位名稱-->
      <INPUT type="hidden" name=btnFieldList value='<%if(session.getAttribute("btnFieldList") != null) out.print((String)session.getAttribute("btnFieldList"));%>'><!--//btnFieldList儲存已勾選的報表欄位名稱-->
      <INPUT type="hidden" name=BankList value='<%if(session.getAttribute("BankList") != null) out.print((String)session.getAttribute("BankList"));%>'><!--//BankList儲存已勾選的金融機構-->
      <INPUT type="hidden" name=S_YEAR value='<%if(session.getAttribute("S_YEAR") != null) out.print((String)session.getAttribute("S_YEAR"));%>'><!--//95.12.06儲存已勾選的起始日期-年-->
      <INPUT type="hidden" name=E_YEAR value='<%if(session.getAttribute("E_YEAR") != null) out.print((String)session.getAttribute("E_YEAR"));%>'><!--//95.12.06儲存已勾選的結束日期-年-->
      <INPUT type="hidden" name=S_MONTH value='<%if(session.getAttribute("S_MONTH") != null) out.print((String)session.getAttribute("S_MONTH"));%>'><!--//95.12.06儲存已勾選的起始日期-月-->
      <INPUT type="hidden" name=E_MONTH value='<%if(session.getAttribute("E_MONTH") != null) out.print((String)session.getAttribute("E_MONTH"));%>'><!--//95.12.06儲存已勾選的結束日期-月-->
    </form>
  </body>
</html>
