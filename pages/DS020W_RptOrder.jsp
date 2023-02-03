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
<%
  String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");
  String bank_type = (session.getAttribute("nowbank_type")==null)?"":(String)session.getAttribute("nowbank_type");
  //String title=(bank_type.equals("6"))?"農會":"漁會";
  String szSortBy = ( session.getAttribute("SortBy")==null ) ? "asc" : (String)session.getAttribute("SortBy");
  System.out.println("szSortBy="+szSortBy);
  String szExcelAction = (session.getAttribute("excelaction")==null)?"download":(String)session.getAttribute("excelaction");
  System.out.println("DS020W_RptOrder.szExcelAction="+szExcelAction);
  //取得DS020W的權限
  Properties permission = ( session.getAttribute("DS020W")==null ) ? new Properties() : (Properties)session.getAttribute("DS020W");
  if (permission == null) {
    System.out.println("DS020W_RptOrder.permission == null");
  } else {
    System.out.println("DS020W_RptOrder.permission.size ="+permission.size());               
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

function doSubmit(cnd) {
  if(cnd == 'createRpt') {
    if(this.document.forms[0].BankList.value == '') {        
      alert('金融機構代碼必須選擇');
      return;
    }
    if (!chInput(this.document.forms[0],'')) return;//95.12.06 add 起始日期不可大於結束日期
    if (this.document.forms[0].btnFieldList.value == '') {
      alert('報表欄位必須選擇');
      return;
    }
    if(!confirm("本項報表會執行10-15秒，是否確定執行？")) {
      return;
    }
  }
  MoveSelectToBtn(this.document.forms[0].SortList, this.document.forms[0].SortListDst);
  fn_ShowPanel(cnd);
}

function fn_ShowPanel(cnd) {
  //act=BankList/RptColumn/RptOrder/RptType
  this.document.forms[0].action = "/pages/DS020W.jsp?act="+cnd;
  this.document.forms[0].target = '_self';
  this.document.forms[0].submit();
}

function ResetAllData() {
  if(confirm("確定要清除已選定的資料嗎？")) {
    this.document.forms[0].SortListDst.length = 0;
    getbtnFieldList();
  }
  return;
}
//-->
    </script>
    <link href="css/b51.css" rel="stylesheet" type="text/css">
  </head>
  <body leftmargin="0" topmargin="20">
    <form method=post action='#' name='RptOrderfrm'>
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
                  <div style="margin:50px;vertical-align:middle;display:inline"><a href="javascript:doSubmit('BankList')"><font color='black'>1.金融機構</font></a></div>
                  <div style="margin:50px;vertical-align:middle;display:inline"><a href="javascript:doSubmit('RptColumn')"><font color='black'>2.報表欄位</font></a></div>
                  <div style="margin:50px;vertical-align:middle;display:inline"><img src="images/2_icon_01.gif" width="16" height="16" align="absmiddle"><a href="#"><font color="#CC6600">3.排序欄位</font></a></div>
                  <div style="margin:50px;vertical-align:middle;display:inline"><a href="javascript:doSubmit('RptStyle')"><font color='black'>4.報表格式</font></a></div>
                </div>
              </td>
            </tr>
            <tr>
              <td>
              	<div style="vertical-align:middle;padding-top:2px">
              	  <table>
              	    <tr>
              	      <td>
                        <select multiple size=10 name="SortListSrc" ondblclick="javascript:movesel(this.document.forms[0].SortListSrc,this.document.forms[0].SortListDst);fn_changeBankListSrc(document.forms[0]);" style="width:292;height:190"></select>
              	      </td>
              	      <td width='40px' align="center">
                        <div style="padding:5px 0px 5px 0px">                                 
                          <a href="javascript:movesel(this.document.forms[0].SortListSrc,this.document.forms[0].SortListDst);"><img src="images/arrow_right.gif" width="24" height="22" border="0"></a>
                        </div>
                        <div style="padding:5px 0px 5px 0px">                                 
                          <a href="javascript:moveallsel(this.document.forms[0].SortListSrc,this.document.forms[0].SortListDst);"><img src="images/arrow_rightall.gif" width="24" height="22" border="0"></a>
                        </div>
                        <div style="padding:5px 0px 5px 0px">                                 
                          <a href="javascript:movesel(this.document.forms[0].SortListDst,this.document.forms[0].SortListSrc);"><img src="images/arrow_left.gif" width="24" height="22" border="0"></a>
                        </div>
                        <div style="padding:5px 0px 5px 0px">                                 
                          <a href="javascript:moveallsel(this.document.forms[0].SortListDst,this.document.forms[0].SortListSrc);"><img src="images/arrow_leftall.gif" width="24" height="22" border="0"></a>
                        </div>
              	      </td>
              	      <td>
                        <select multiple size=10 name="SortListDst" ondblclick="javascript:movesel(this.document.forms[0].SortListDst,this.document.forms[0].SortListSrc);fn_changeBankListSrc(document.forms[0]);"" style="width:292;height:190"></select>
              	      </td>
              	      <td class="sbody">
              	      	<div style="padding:5px 0px 5px 0px">
                          <div style="display:inline;vertical-align:middle;padding-left:5px"><a href="javascript:moveup(this.document.forms[0].SortListDst);"><img src="images/arrow_up.gif" width="24" height="22" border="0"></a></div>
                          <div class="sbody"  style="display:inline;padding-left:8px">欄位上移</div>
                        </div>
                        <div style="padding:5px 0px 5px 0px">
                          <div style="display:inline;vertical-align:middle;padding-left:5px"><a href="javascript:movedown(this.document.forms[0].SortListDst);"><img src="images/arrow_down.gif" width="24" height="22" border="0"></a></div>
                          <div class="sbody"  style="display:inline;padding-left:8px">欄位下移</div>
                        </div>
                        <div style="padding:5px 0px 5px 0px">
                          <div style="display:inline;vertical-align:middle;padding-left:5px"><input type="radio" name="SortBy" value="asc" <% if(szSortBy.equals("asc")) out.print("checked"); %>>遞增排序</div>                          
                        </div>
                        <div style="padding:5px 0px 5px 0px">
                          <div style="display:inline;vertical-align:middle;padding-left:5px"><input type="radio" name="SortBy" value="desc" <% if(szSortBy.equals("desc")) out.print("checked"); %>>遞減排序</div>                         
                        </div>
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
              	    <div>註：「信用部主任及分部主任參訓情形」本排序欄位若未選取，則按「縣市別」、「金融機構代號名稱」及「年月」等欄位排序印出。</div>
                  </font>
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
    <script language="JavaScript" >
<!--  
<% //從session裡把勾選的欲sort的報表欄位讀出來.放在SortListDst 	
  if (session.getAttribute("SortList") != null && !((String)session.getAttribute("SortList")).equals("")) { 
%>
    var bnlist;
    bnlist = "<%=(String)session.getAttribute("SortList")%>";//95.12.07 add 
    var a = bnlist.split(',');
    for (var i =0; i < a.length; i ++) {
      var j = a[i].split('+');
      this.document.forms[0].SortListDst.options[i] = new Option(j[1], j[0]);
    }
<%}%>

<% //從session裡把已勾選的報表欄位讀出來.放在FieldListDst 	
  if (session.getAttribute("btnFieldList") != null && !((String)session.getAttribute("btnFieldList")).equals("")) { 
    System.out.println("DS020W_RptOrder.btnFieldList="+(String)session.getAttribute("btnFieldList"));
  }
%>
getbtnFieldList();
function getbtnFieldList(){
  //95.08.23 add 縣市別/金融機構代號名稱放在已選的sort報表欄位===================================================    
  //this.document.forms[0].SortListDst.options[0] = new Option("縣市別", "hsien_id");
  //this.document.forms[0].SortListDst.options[1] = new Option("金融機構代號名稱", "bank_no");        
  //=============================================================================================================
  //根據所選加入欲sort的報表欄位跟已選sort的報表欄位
  var bnlist;
  var checkAdd = false;
  var addCount = 0;		
  bnlist = "<%=(String)session.getAttribute("btnFieldList")%>";//95.12.07 add 	
<%
  session.setAttribute("SortList",null);//清除已勾選的欲sort的報表欄位
%>
  var a = bnlist.split(',');
  //95.08.23 add 若已選排序欄位無"縣市別"."金融機構代號名稱"時,才加入到可選排序欄位 by 2295 ==============================================
  //95.08.30 add 若已選排序欄位無"年"時,才加入到可選排序欄位 by 2295 		
  //alert(this.document.forms[0].BankList.value);
  //95.09.04 add 若已選的機構代號不為全部時，才可根據縣市別．金融機構代號sort	   
  //                               =全部時，根據年，縣市別小計sort
  //95.12.06 add 各別機構可根據縣市別.金融機構代號.年月sort
  //             全部可選根據縣市別小計.年月sort	
  if (this.document.forms[0].BankList.value.indexOf("ALL") == -1) {
    //縣市別======================================================================================================
    checkAdd = false;
    for (var k =0;k<this.document.forms[0].SortListDst.length;k++) {			
      if(this.document.forms[0].SortListDst.options[k].text == "縣市別") {
        checkAdd = true;
	    }  
    }  
    if(checkAdd == false) {	  
      this.document.forms[0].SortListSrc.options[addCount] = new Option("縣市別", "FR001W_output_order");
      addCount++;
    }
    //金融機構代號名稱============================================================================================
    checkAdd = false;
    for (var k =0;k<this.document.forms[0].SortListDst.length;k++) {			
      if(this.document.forms[0].SortListDst.options[k].text == "金融機構代號名稱") {		    
        checkAdd = true;			       
      }  
    }    
    if(checkAdd == false) {	  
      this.document.forms[0].SortListSrc.options[addCount] = new Option("金融機構代號名稱", "bank_no");
      addCount++;
    }
    //年.月===================================================================================================
    checkAdd = false;
    for (var k =0;k<this.document.forms[0].SortListDst.length;k++) {			
      if (this.document.forms[0].SortListDst.options[k].text == "年月") {		       
        checkAdd = true;			       
      }  
    } 	    
    if (checkAdd == false) {	  
      this.document.forms[0].SortListSrc.options[addCount] = new Option("年月", "m_yearmonth");
      addCount++;
    }	
  } else { //已選的機構代號=ALL全部
    //縣市別小計============================================================================================
    checkAdd = false;
    for (var k =0;k<this.document.forms[0].SortListDst.length;k++) {
      if (this.document.forms[0].SortListDst.options[k].text == "縣市別小計") {
        checkAdd = true;
      }
    }
    if (checkAdd == false) {
      this.document.forms[0].SortListSrc.options[addCount] = new Option("縣市別小計", "FR001W_OUTPUT_ORDER");
      addCount++;
    }
    //年.月===================================================================================================
    checkAdd = false;
    for (var k =0; k<this.document.forms[0].SortListDst.length; k++) {			
      if (this.document.forms[0].SortListDst.options[k].text == "年月") {		    	       
        checkAdd = true;			       
      }  
    } 	    
    if (checkAdd == false) {	  
      this.document.forms[0].SortListSrc.options[addCount] = new Option("年月", "m_yearmonth");
      addCount++;
    }	
  }//end of 已選的機構代號=ALL全部
    
    //======================================================================================================
    //已選的報表欄位.95.12.06 不可根據已選欄位做sort
    /*
	for (var i =0; i < a.length; i ++){
		var j = a[i].split('+');
		checkAdd=false;		
		for(var k =0;k<this.document.forms[0].SortListDst.length;k++){			
			if(this.document.forms[0].SortListDst.options[k].text == j[1]){		    			    
		   		checkAdd = true;			       
	    	}   
    	}    	
    	if(checkAdd == false){	       	       
    	    //if(j[1] == '單位數'){//95.09.05 add 單位數不可做為sort欄位
    	    //   continue;
    	    //}    	        	    
       		this.document.forms[0].SortListSrc.options[addCount] = new Option(j[1], j[0]);       		
       		addCount++;
    	}	
	}*/	
}
-->
    </script>
  </body>
</html>