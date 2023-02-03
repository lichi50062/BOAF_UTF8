<%
// 94.03.10	增加營運中/裁撤別下拉選單處理 add by egg
// 94.08.12  add Office 2003版Excel現僅提供「下載報表」功能 
// 94.09.05  add 拿掉檢視報表 by 2295
// 94.11.16  add 全體漁會信用部.金額單位 by 2295
// 99.04.26  fix 縣市合併調整 by 2808
//100.05.03 fix 無法顯示漁會機構名稱 by 2295
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
	showUnit=true;//顯示金額單位
	showPageSetting=true;//顯示報表列印格式
	setLandscape=false;//true:橫印
	Map dataMap =Utility.saveSearchParameter(request);
	String S_YEAR = Utility.getTrimString(dataMap.get("S_YEAR")) ;		
	String S_MONTH = Utility.getTrimString(dataMap.get("S_MONTH")) ;		
	String report_no = "FR003F";
	String bankType = Utility.getTrimString(dataMap.get("bankType"));
	//94.03.10 add by egg
   	String sqlCmd = "";
	String cancel_no = "";
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/jquery-1.2.6.js"></script>
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

   this.document.forms[0].action = "/pages/FR003F_Excel.jsp?act="+cnd;	
   this.document.forms[0].target = '_self';
   this.document.forms[0].submit();   
}
//94.03.10 add by egg
//重導FR003F_Qry.jsp頁面以致輸出對應之金融機構資料
function getData(form,item){
	if(item == 'cancel_no'){
		//初始BANK_NO值
	   	form.BANK_NO.value="";
	}
	form.action="/pages/FR003F_Qry.jsp?act=getData&test=nothing";
	form.submit();
}
//added  縣市合併調整 by 2808
function changeTbank(xml) {
	var myXML,nodeType,nodeValue, nodeName,nodeCity,nodeYear;
    var oOption;    
    
   	
    myXML = document.all(xml).XMLDocument;
    nodeType = myXML.getElementsByTagName("bankType");//bank_type農/漁會
    nodeCity = myXML.getElementsByTagName("bankCity");//hsien_id縣市別
	nodeValue = myXML.getElementsByTagName("bankValue");//bank_no機構代號
	nodeName = myXML.getElementsByTagName("bankName");//bank_no+bank_name
	nodeYear = myXML.getElementsByTagName("bankYear");//m_year所屬年度
	BnType = myXML.getElementsByTagName("BnType");//bn_type營運中/已裁撤
	for(var i=0;i<nodeType.length ;i++)	{
		//無區分縣市別==============
	    //有區分營運中/已裁撤
	    if($('select[@name=CANCEL_NO]').val() == 'N'){//營運中				
	         if(BnType.item(i).firstChild.nodeValue != '2'){
		         $('select[@name=BANK_NO]')
	        	 .append("<option value='"+nodeValue.item(i).firstChild.nodeValue+"/"+nodeName.item(i).firstChild.nodeValue+"' year='"+nodeYear.item(i).firstChild.nodeValue+"'>"+nodeName.item(i).firstChild.nodeValue+"</option>") ;
	         }		
        }else{//已裁撤		    
             if(BnType.item(i).firstChild.nodeValue == '2'){
            	 $('select[@name=BANK_NO]')
            	 .append("<option value='"+nodeValue.item(i).firstChild.nodeValue+"/"+nodeName.item(i).firstChild.nodeValue+"' year='"+nodeYear.item(i).firstChild.nodeValue+"'>"+nodeName.item(i).firstChild.nodeValue+"</option>") ;	  
	         }
          }
	}
	removeBankOption() ;
	$('select[@name=BANK_NO]').prepend("<option value='ALL/全體漁會信用部'>全體漁會信用部</option>") ;
	$('select[@name=BANK_NO] > option[@value=ALL/全體漁會信用部]').attr('selected',true) ;
}
function removeBankOption() {
	var year = $('input[@name=S_YEAR]').val()==''? 99 : eval($('input[@name=S_YEAR]').val() ) ;
	syear = '' ;
	if(year > 99) {
		syear = '100' ;
	}else {
		syear = '99' ;
	}
	$('select[@name=BANK_NO] > option').each(function(){
		if($(this).attr('year')!=syear) {
			$(this).remove() ;
		}
	});

}
function changeCity(xml, target, source, form) {} //共用畫面會呼叫到的物件
//-->
</script>
<link href="css/b51.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0">
<form method=post action='#'>
<input type='hidden' name="showTbank" value='<%=showBankType %>'>
<input type='hidden' name="showCityType" value='<%=showCityType%>'>
<input type="hidden" name="showCancel_No" value='<%=showCancel_No %>'>
<table width='600' border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="25%"><img src="images/banner_bg1.gif" width="100%" height="17"></td>
    <td width="50%"><font color='#000000' size=4><b><center>漁會信用部資產負債表</center></b></font></td>
    <td width="25%"><img src="images/banner_bg1.gif" width="100%" height="17"></td>
  </tr>
</table>
<Table border=1 width='600' align=center height="65" bgcolor="#FFFFF" bordercolor="#76C657">
<tr class="sbody" bgcolor="#BDDE9C">
    <td colspan="2" height="1">
      <div align="right">
       <input type='radio' name="excelaction" value='download' checked> 下載報表
       <a href="javascript:doSubmit('createRpt')" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image411','','images/bt_execb.gif',1)"><img src="images/bt_exec.gif" name="Image411" width="66" height="25" border="0" id="Image41"></a>
       <a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image511','','images/bt_cancelb.gif',0)"><img src="images/bt_cancel.gif" name="Image511" width="66" height="25" border="0" id="Image51"></a>
       <a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image611','','images/bt_reporthelpb.gif',1)"><img src="images/bt_reporthelp.gif" name="Image611" width="80" height="25" border="0" id="Image61"></a>
      </div>
    </td>
</tr>
<%@include file="./include/ym_hsien_id_unit.include" %><!--顯示查詢日期/金融機構類別/縣市別-->
<tr class="sbody">
	<td width="118" bgcolor="#BDDE9C" height="1">機構單位 :</td>
    <td width="416" bgcolor="#EBF4E1" height="1">
		<select name='BANK_NO'/>
    </td>
<table border="1" width="600" align="center" height="54" bgcolor="#FFFFF" bordercolor="#76C657">
  <tr>
  	<td bgcolor="#E9F4E3" colspan="2">
       <div align="center">
  		<table width="574" border="0" cellpadding="1" cellspacing="1">
  		<tr><td width="34"><img src="/pages/images/print_1.gif" width="34" height="34"></td>
            <td width="492"><font color="#CC6600">本報表採用A4紙張橫印</font></td>                              
        </tr>
        </table>
        </div>
    </td>
  </tr>  
</table>    
</table>

<!-- -------------------------------------------------------------- -->

</form>
<script language="JavaScript" type="text/JavaScript">

changeTbank("TBankXML") ;
$('select[@name=CANCEL_NO]').bind('change',function(){
	changeTbank("TBankXML") ;
 }) ;
$('input[@name=S_YEAR]').bind('change',function(){
	changeTbank("TBankXML") ;
 }) ;
</script>
</body>
</html>
