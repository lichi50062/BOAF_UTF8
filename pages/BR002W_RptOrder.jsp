<%
// 94.01.20 create by 2295
// 94.01.31 add 權限 by 2295
// 99.12.21 add 排序欄位.可挑選項目套用共用include by 2295
// 99.12.21 add 排序欄位套用共用include by 2295
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
	String bank_type = (session.getAttribute("nowbank_type")==null)?"":(String)session.getAttribute("nowbank_type");
	String DS_bank_type = (session.getAttribute("DS_bank_type")==null)?"6":(String)session.getAttribute("DS_bank_type");	
	String szSortBy = ( session.getAttribute("SortBy")==null ) ? "asc" : (String)session.getAttribute("SortBy");				
	System.out.println("szSortBy="+szSortBy);
	String szExcelAction = (session.getAttribute("excelaction")==null)?"view":(String)session.getAttribute("excelaction");
	//System.out.println("BR002W_RptOrder.szExcelAction="+szExcelAction);	
	String bank_name = (bank_type.equals("6"))?"農會":"漁會";
	if(DS_bank_type.equals("B")) bank_name="農漁會";
%>



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
      if(this.document.forms[0].BankList.value == ''){        
         alert('金融機構代碼必須選擇');
         return;
      }
    
      if(this.document.forms[0].btnFieldList.value == ''){
         alert('報表欄位必須選擇');
         return;
      }
   }   
   MoveSelectToBtn(this.document.forms[0].SortList, this.document.forms[0].SortListDst);	
   fn_ShowPanel(report_no,cnd);      
}

function ResetAllData_BR(){
    if(confirm("確定要清除已選定的資料嗎？")){  	
        this.document.forms[0].SortListDst.length = 0;
        getbtnFieldList();    	
	}
	return;	
}
//-->
</script>
<link href="css/b51.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0">
<form method=post action='#' name='RptOrderfrm'>
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
                          <!--input type='radio' name="excelaction" value='view' <%if(szExcelAction.equals("view")){out.print("checked");}%> >檢視報表-->
                      	  <input type='radio' name="excelaction" value='download' <%if(szExcelAction.equals("download")){out.print("checked");}%> >下載報表
                          <%if(Utility.getPermission(request,report_no,"P")){//Print--有列印權限時 %> 
                      	  <a href="javascript:doSubmit('<%=report_no%>','createRpt');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image411','','images/bt_execb.gif',1)"><img src="images/bt_exec.gif" name="Image411" width="66" height="25" border="0" id="Image41"></a> 
                      	  <%}%>
                          <a href="javascript:ResetAllData_BR('RptOrder');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image511','','images/bt_cancelb.gif',0)"><img src="images/bt_cancel.gif" name="Image511" width="66" height="25" border="0" id="Image51"></a> 
                          <a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image611','','images/bt_reporthelpb.gif',1)"><img src="images/bt_reporthelp.gif" name="Image611" width="80" height="25" border="0" id="Image61"></a> 
                        </div></td>
                    </tr>
                    <tr> 
                      <td class="menu_bgcolor"><table width="750" border="0" align="center" cellpadding="1" cellspacing="1" class="sbody">
                          <tr class="sbody"> 
                            <td width="100"><a href="javascript:doSubmit('<%=report_no%>','BankList')"><font color='black'>1.金融機構</font></a></td>                            
                            <td width="100"><a href="javascript:doSubmit('<%=report_no%>','RptColumn')"><font color='black'>2.報表欄位</font></a></td>
                            <td width="100"><img src="images/2_icon_01.gif" width="16" height="16" align="absmiddle"> 
                              <a href="#"><font color="#CC6600">3.排序欄位</font></a></td>                                                       
                            <td width="100"><a href="javascript:doSubmit('<%=report_no%>','RptStyle')"><font color='black'>4.報表格式</font></a></td>
                          </tr>
                        </table></td>
                    </tr>
                    <%@include file="./include/DS_RptOrder.include" %><!-- 排序欄位.可挑選項目-->                    
                  </table></td>
              </tr>
            </table></td>
        </tr>        
      </table>
    </td>
  </tr>
</table>
<INPUT type="hidden" name=SortList><!--//btnSortList儲存已勾選的報表排序欄位名稱-->
<INPUT type="hidden" name=btnFieldList value='<%if(session.getAttribute("btnFieldList") != null) out.print((String)session.getAttribute("btnFieldList"));%>'><!--//btnFieldList儲存已勾選的報表欄位名稱-->
<INPUT type="hidden" name=BankList value='<%if(session.getAttribute("BankList") != null) out.print((String)session.getAttribute("BankList"));%>'><!--//BankList儲存已勾選的金融機構-->
</form>
<script language="JavaScript" >
<!--  

<%//從session裡把勾選的欲sort的報表欄位讀出來.放在SortListDst 	
if(session.getAttribute("SortList") != null && !((String)session.getAttribute("SortList")).equals("")){ 
%>
	var bnlist;
	bnlist = '<%=(String)session.getAttribute("SortList")%>';
	var a = bnlist.split(',');
	for (var i =0; i < a.length; i ++){
		var j = a[i].split('+');
		this.document.forms[0].SortListDst.options[i] = new Option(j[1], j[0]);
	}
<%}%>

<%//從session裡把已勾選的報表欄位讀出來.放在FieldListDst 	
if(session.getAttribute("btnFieldList") != null && !((String)session.getAttribute("btnFieldList")).equals("")){ 
   System.out.println("BR002W_RptOrder.btnFieldList="+(String)session.getAttribute("btnFieldList"));
%>
	getbtnFieldList();
<%}%>
function getbtnFieldList(){
	var bnlist;
	var checkAdd = false;
	var addCount = 0;		
	bnlist = '<%=(String)session.getAttribute("btnFieldList")%>';	
	<%
	session.setAttribute("SortList",null);//清除已勾選的欲sort的報表欄位
	%>
	var a = bnlist.split(',');
	for (var i =0; i < a.length; i ++){
		var j = a[i].split('+');
		checkAdd=false;
		for(var k =0;k<this.document.forms[0].SortListDst.length;k++){			
			if(this.document.forms[0].SortListDst.options[k].text == j[1]){		    
		   		checkAdd = true;			       
	    	}   
    	}
    	if(checkAdd == false){	       	       
       		this.document.forms[0].SortListSrc.options[addCount] = new Option(j[1], j[0]);
       		addCount++;
    	}	
	
	}
}
-->
</script>
</body>
</html>
