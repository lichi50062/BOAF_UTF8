<%
//Copy from FR003W
//102.01.24 add by 2968
//105.03.23 add for 地方主管機關使用 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.util.Calendar" %>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="java.util.List" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%@include file="./include/bank_no_hsien_id.include" %>
<%
	showCancel_No=true;//顯示營運中/裁撤別
	showBankType=false;//顯示金融機構類別
	showCityType=false;//顯示縣市別
	showUnit=false;//顯示金額單位
	showPageSetting=true;//顯示報表列印格式
	setLandscape=false;//true:橫印
	
	String report_no = "FR063W";
    Map dataMap =Utility.saveSearchParameter(request);
	String title="信用部聯合貸款案件明細表";
    String S_YEAR = Utility.getTrimString(dataMap.get("S_YEAR"));
	String S_MONTH = Utility.getTrimString(dataMap.get("S_MONTH"));
	String bankType = Utility.getTrimString(dataMap.get("bankType"));
	String hasMonth = Utility.getTrimString(dataMap.get("hasMonth"));
	String cancel_no = Utility.getTrimString(dataMap.get("CANCEL_NO"));
    String bank_type = Utility.getTrimString(dataMap.get("bank_type"));
    if (tBankList == null) {
		//tBankList = Utility.getBankList(request) ;//Utility.getALLTBank(bank_type) ;
		//tBankList = Utility.getALLTBank(bank_type);//100.05.04
		session.setAttribute("nowbank_type", bank_type);//100.06.24
		tBankList = Utility.getBankList(request);//100.06.24按照直轄市在前.其他縣市在後排序.
		// XML Ducument for 總機構代碼 begin
		out.println("<xml version=\"1.0\" encoding=\"UTF-8\" ID=\"TBankXML\">");
		out.println("<datalist>");
		for (int i = 0; i < tBankList.size(); i++) {
			bean = (DataObject) tBankList.get(i);
			out.println("<data>");
			out.println("<bankType>" + bean.getValue("bank_type")
					+ "</bankType>");
			out.println("<BnType>" + bean.getValue("bn_type")
					+ "</BnType>");
			out.println("<bankYear>"
					+ bean.getValue("m_year").toString()
					+ "</bankYear>");
			out.println("<bankCity>" + bean.getValue("hsien_id")
					+ "</bankCity>");
			out.println("<bankValue>" + bean.getValue("bank_no")
					+ "</bankValue>");
			out.println("<bankName>" + bean.getValue("bank_no") + "  "
					+ bean.getValue("bank_name") + "</bankName>");
			out.println("</data>");
		}
		out.println("</datalist>\n</xml>");
		// XML Ducument for 總機構代碼 en
	}
%>
<script language="javascript" src="js/Common.js"></script>
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

function doSubmit(){ 
		var bankInfo = this.document.forms[0].BANK_NO.value;
		var strAry = new Array();
		var strAry = bankInfo.split("/");
		if(strAry[0] == "ALL"){
			this.document.forms[0].action = "/pages/FR063W_Excel.jsp?report=FR064W&bank_type="+<%=bankType%>+
											"&s_year="+this.document.forms[0].S_YEAR.value +
											"&s_month="+this.document.forms[0].S_MONTH.value;	
			this.document.forms[0].target = '_self';
	   	}else{
	   		this.document.forms[0].action = "/pages/FR063W_Excel.jsp?report=FR063W&bank_name="+strAry[1]+
	   										"&bank_no="+strAry[0]+
							   				"&s_year="+this.document.forms[0].S_YEAR.value+
							   				"&s_month="+this.document.forms[0].S_MONTH.value+
							   				"&Unit="+this.document.forms[0].Unit.value;
	   	}
		this.document.forms[0].submit(); 
}
//94.03.10 add by egg
//重導FR063W_Qry.jsp頁面以致輸出對應之金融機構資料
function getData(form,item){
	if(item == 'cancel_no'){
		//初始BANK_NO值
	   	form.BANK_NO.value="";
	}
	form.action="/pages/FR063W_Qry.jsp?act=getData&test=nothing";
	form.submit();
}
function changeTbank(xml) {
	var myXML,nodeType,nodeValue, nodeName,nodeCity,nodeYear;
    var oOption;    
    var target = document.getElementById("BANK_NO");
    var type = '<%=bankType%>';
   	var m_year = form.S_YEAR.value;
   	target.length = 0;
    if(m_year >= 100){
       m_year = 100;
    }else{
       m_year = 99;
    }
   	
    myXML = document.all(xml).XMLDocument;
    nodeType = myXML.getElementsByTagName("bankType");//bank_type農/漁會
    nodeCity = myXML.getElementsByTagName("bankCity");//hsien_id縣市別
	nodeValue = myXML.getElementsByTagName("bankValue");//bank_no機構代號
	nodeName = myXML.getElementsByTagName("bankName");//bank_no+bank_name
	nodeYear = myXML.getElementsByTagName("bankYear");//m_year所屬年度
	BnType = myXML.getElementsByTagName("BnType");//bn_type營運中/已裁撤

	/*
	var typeName="";
	if(type=="6"){
		typeName="農會"
	}else{
		typeName="漁會"
	}
	oOption = document.createElement("OPTION");
	oOption.text="全體"+typeName+"信用部";
	oOption.value="ALL/全體"+typeName+"信用部";
	target.add(oOption);*/
	
	for(var i=0;i<nodeType.length ;i++)	{
		//無區分縣市別==============
	    //有區分營運中/已裁撤
	    if((nodeYear.item(i).firstChild.nodeValue == m_year)) {  
		    if(form.CANCEL_NO.value == 'N'){//營運中				
		         if(BnType.item(i).firstChild.nodeValue != '2'){
		        	 oOption = document.createElement("OPTION");
		        	 oOption.text=nodeName.item(i).firstChild.nodeValue;
			         oOption.value=nodeValue.item(i).firstChild.nodeValue+"/"+nodeName.item(i).firstChild.nodeValue;  
			         target.add(oOption);
		         }		
	        }else{//已裁撤		    
	             if(BnType.item(i).firstChild.nodeValue == '2'){
	            	 oOption = document.createElement("OPTION");
	            	 oOption.text=nodeName.item(i).firstChild.nodeValue;
			         oOption.value=nodeValue.item(i).firstChild.nodeValue+"/"+nodeName.item(i).firstChild.nodeValue;
			         target.add(oOption);
		         }
	        }
		    
	    }
	}
	target[0].selected=true;
}
function changeCity(xml, target, source, form) {} //共用畫面會呼叫到的物件
//-->
</script>
<link href="css/b51.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0">
<form name='form' method=post action='#'>
<input type='hidden' name="showTbank" value='<%=showBankType %>'>
<input type='hidden' name="showCityType" value='<%=showCityType%>'>
<input type="hidden" name="showCancel_No" value='<%=showCancel_No %>'>
<table width='600' border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="25%"><img src="images/banner_bg1.gif" width="100%" height="17"></td>
    <td width="50%"><font color='#000000' size=4><center><%=title %></center></font></td>
    <td width="25%"><img src="images/banner_bg1.gif" width="100%" height="17"></td>
  </tr>
</table>
<%
  String nameColor="nameColor_sbody";
  String textColor="textColor_sbody";
  String bordercolor="#3A9D99";
%> 
<Table border=1 width='600' align=center height="65" bgcolor="#FFFFF" bordercolor="<%=bordercolor%>">
<tr class="sbody">
    <td class="<%=nameColor%>" colspan="2" height="1">
      <div align="right">
       <input type='radio' name="excelaction" value='download' checked> 下載報表
       <a href="javascript:doSubmit()" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image411','','images/bt_execb.gif',1)"><img src="images/bt_exec.gif" name="Image411" width="66" height="25" border="0" id="Image41"></a>
       <a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image511','','images/bt_cancelb.gif',0)"><img src="images/bt_cancel.gif" name="Image511" width="66" height="25" border="0" id="Image51"></a>
       <a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image611','','images/bt_reporthelpb.gif',1)"><img src="images/bt_reporthelp.gif" name="Image611" width="80" height="25" border="0" id="Image61"></a>
      </div>
    </td>
</tr>
<%//105.6.6 控制點不同，無法共用  ym_hsien_id_unit_m2.include  by 2968%>
<%@include file="./include/ym_hsien_id_unit_m2_1.include"%><!--顯示查詢日期/金融機構類別/縣市別/機構單位-->
<!-- 金額單位-->
<tr class="sbody">
	<td width="118" class="<%=nameColor%>" height="1">金額單位 :</td>
 	<td width="416" class="<%=textColor%>" height="1">
	   <select size="1" name="Unit">
	     <option value ='1' <%if((!Unit.equals("")) && Unit.equals("1")) out.print("selected");%>>元</option>
	     <option value ='1000' <%if((!Unit.equals("")) && Unit.equals("1000")) out.print("selected");%>>千元</option>
	     <option value ='10000' <%if((!Unit.equals("")) && Unit.equals("10000")) out.print("selected");%>>萬元</option>
	     <option value ='1000000' <%if((!Unit.equals("")) && Unit.equals("1000000")) out.print("selected");%>>百萬元</option>
	     <option value ='10000000' <%if((!Unit.equals("")) && Unit.equals("10000000")) out.print("selected");%>>千萬元</option>
	     <option value ='100000000' <%if((!Unit.equals("")) && Unit.equals("100000000")) out.print("selected");%>>億元</option>
	   </select>
 	</td>
</tr> 
<table border="1" width="600" align="center" height="54" bgcolor="#FFFFF" bordercolor="<%=bordercolor%>">
  <tr>
  	<td class="<%=nameColor%>" colspan="2">
       <div align="center">
  		<table width="574" border="0" cellpadding="1" cellspacing="1">
  		<tr><td width="34"><img src="/pages/images/print_1.gif" width="34" height="34"></td>
            <td width="492"><font color="#CC6600">本報表採用A4紙張直印</font></td>                              
        </tr>
        </table>
        </div>
    </td>
  </tr>  
</table> 
</Table>
<!-- ======================================================== -->

</form>

</body>
<script language="JavaScript" type="text/JavaScript">

changeTbank("TBankXML") ;

//removeBankOption();

</script>
</html>
