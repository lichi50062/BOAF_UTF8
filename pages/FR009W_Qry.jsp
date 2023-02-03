<%
// 94.03.10	增加營運中/裁撤別下拉選單處理 add by egg
// 94.09.05 fix 拿掉檢視報表 by 2295
//100.05.09 fix 套用畫面模組/縣市合併問題 by 2295
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
	
	String report_no = "FR009W";
    Map dataMap =Utility.saveSearchParameter(request);
    
	String S_YEAR = Utility.getTrimString(dataMap.get("S_YEAR")) ;		
	String S_MONTH = Utility.getTrimString(dataMap.get("S_MONTH")) ;				
	String bank_type = Utility.getTrimString(dataMap.get("bank_type"));	 
	String bankType = bank_type;//ym_hsien_id_unit.include用 
    String cancel_no  = "" ;
 	String title=(bank_type.equals("6"))?"農會":"漁會";
 	if(MONTH.equals("1")){//若本月為1月份是..則是申報上個年度的12月份
       YEAR = String.valueOf(Integer.parseInt(YEAR) - 1);
       MONTH = "12";
    }   
 
    if(tBankList==null) {
        System.out.println("tBankList==null");
    	tBankList = Utility.getBankList(request) ;
    	 // XML Ducument for 總機構代碼 begin
	    out.println("<xml version=\"1.0\" encoding=\"UTF-8\" ID=\"TBankXML\">");
	    out.println("<datalist>");
	    for(int i=0;i< tBankList.size(); i++) {
	        bean =(DataObject)tBankList.get(i);
	        out.println("<data>");
	        out.println("<bankType>"+bean.getValue("bank_type")+"</bankType>");
	        out.println("<BnType>"+bean.getValue("bn_type")+"</BnType>");
	        out.println("<bankYear>"+bean.getValue("m_year").toString()+"</bankYear>");
	        out.println("<bankCity>"+bean.getValue("hsien_id")+"</bankCity>");
	        out.println("<bankValue>"+bean.getValue("bank_no")+"</bankValue>");
	        out.println("<bankName>"+bean.getValue("bank_no")+"  "+bean.getValue("bank_name")+"</bankName>");
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

function doSubmit(cnd){ 
   var tmp_year,tmp_month;
   tmp_year = parseInt(this.document.forms[0].S_YEAR.value);
   tmp_month = parseInt(this.document.forms[0].S_MONTH.value);
   if( tmp_year * 100 + tmp_month < 9406){
      alert('94年6月起開始受理申報資料');
      return;
   } 

   this.document.forms[0].action = "/pages/FR009W_Excel.jsp";	
   this.document.forms[0].target = '_self';
   this.document.forms[0].submit();   
}

function changeTbank(xml) {
	var myXML,nodeType,nodeValue, nodeName,nodeCity,nodeYear;
    var oOption;    
    var target = document.getElementById("BANK_NO");
    //var type = form.bankType.value;
   	var m_year = form.S_YEAR.value;
   	target.length = 0;
    if(m_year >= 100){
       m_year = 100;
    }else{
       m_year = 99;
    }
   	//this.document.forms[0].BANK_NO.length = 0;
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
	    if(nodeYear.item(i).firstChild.nodeValue == m_year){
		    //if(nodeType.item(i).firstChild.nodeValue == type)  {//相同年度.農/漁會  
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
			    
		    //}
	    }
	}
	/*for(var i=0;i<nodeType.length ;i++)	{
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
	}*/
	//removeBankOption() ;	
}
/*
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

}*/
function changeCity(xml, target, source, form) {} //共用畫面會呼叫到的物件

//-->
</script>
<link href="css/b51.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0">
<form name='form' method=post action='#'>
<input type=hidden name="bank_type" value=<%=bank_type%>>
<input type='hidden' name="showTbank" value='<%=showBankType %>'>
<input type='hidden' name="showCityType" value='<%=showCityType%>'>
<input type="hidden" name="showCancel_No" value='<%=showCancel_No %>'>
<table width='600' border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="20%"><img src="images/banner_bg1.gif" width="100%" height="17"></td>
    <td width="60%"><font color='#000000' size=4><center>信用部淨值占風險性資產比率計算表</center></font></td>
    <td width="20%"><img src="images/banner_bg1.gif" width="100%" height="17"></td>
  </tr>
</table>
<%
  String nameColor="nameColor_sbody";
  String textColor="textColor_sbody";
  String bordercolor="#3A9D99";
%>
<Table border=1 width='600' align=center height="65" bgcolor="#FFFFF" bordercolor="<%=bordercolor%>">
<tr class="sbody" >
    <td colspan="2" height="1" class="<%=nameColor%>">
      <div align="right">
       <input type='radio' name="act" value='download' checked> 下載報表
       <%if(Utility.getPermission(request,report_no,"P")){//Print %> 
       <a href="javascript:doSubmit('createRpt')" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image411','','images/bt_execb.gif',1)"><img src="images/bt_exec.gif" name="Image411" width="66" height="25" border="0" id="Image41"></a>
       <%}%>
       <a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image511','','images/bt_cancelb.gif',0)"><img src="images/bt_cancel.gif" name="Image511" width="66" height="25" border="0" id="Image51"></a>
       <a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image611','','images/bt_reporthelpb.gif',1)"><img src="images/bt_reporthelp.gif" name="Image611" width="80" height="25" border="0" id="Image61"></a>
      </div>
    </td>
</tr>
<%//105.6.6 控制點不同，無法共用  ym_hsien_id_unit_m2.include  by 2968%>
<%@include file="./include/ym_hsien_id_unit_m2_1.include" %><!--顯示查詢日期/金融機構類別/縣市別/機構單位-->
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
/*$('select[@name=CANCEL_NO]').bind('change',function(){
	changeTbank("TBankXML") ;
 }) ;
$('input[@name=S_YEAR]').bind('change',function(){
	changeTbank("TBankXML") ;
 }) ;*/


//removeBankOption();
</script>
</html>
