<%
//95.10.17 create by 2295
//95.11.14 add 區分BOAF.MIS配色 by 2295
//95.12.04 add 起始日期不可大於結束日期 by 2295
//99.05.13 fix 套用共用畫面DS_RptStyle.include by 2808
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%
	//查詢條件值 
	Map dataMap =Utility.saveSearchParameter(request);
	String report_no = "DS004W";	
	String act = Utility.getTrimString(dataMap.get("act"));					
	String szExcelAction = (session.getAttribute("excelaction")==null)?"download":(String)session.getAttribute("excelaction");
	
	String bank_type = (session.getAttribute("nowbank_type")==null)?"":(String)session.getAttribute("nowbank_type");
	//String title=(bank_type.equals("6"))?"農會":"漁會";
	//取得DS004W的權限
	Properties permission = ( session.getAttribute("DS004W")==null ) ? new Properties() : (Properties)session.getAttribute("DS004W"); 
	if(permission == null){
       System.out.println("DS004W_RptStyle.permission == null");
    }else{
       System.out.println("DS004W_RptStyle.permission.size ="+permission.size());               
    }	
    List templateList = null;
    if(request.getAttribute("templateList") != null){
       templateList = (List)request.getAttribute("templateList");
       System.out.println("templateList.size()="+templateList.size());
    }else{
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

//-->
</script>
<link href="css/b51.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0">
<form method=post action='#' name='RptStylefrm'>
<table width="750" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
     <td>&nbsp;</td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF">
	<table width="791" border="0" align="center" cellpadding="1" cellspacing="1">        
        <tr> 
          <td width="824"><table width="797" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="257"><img src="images/banner_bg1.gif" width="256" height="17"></td>
                <td width="280" class="title_font">A04資產品質分析資料 </td>
                <td width="254"><img src="images/banner_bg1.gif" width="254" height="17"></td>
              </tr>
            </table></td>
        </tr>
        
        <%@include file="./include/DS_RptStyle.include" %><!-- 報表格式(範本檔) -->               
                    
      </table>
 	  </td>
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
<INPUT type="hidden" name=S_YEAR value='<%if(session.getAttribute("S_YEAR") != null) out.print((String)session.getAttribute("S_YEAR"));%>'><!--//95.12.04儲存已勾選的起始日期-年-->
<INPUT type="hidden" name=E_YEAR value='<%if(session.getAttribute("E_YEAR") != null) out.print((String)session.getAttribute("E_YEAR"));%>'><!--//95.12.04儲存已勾選的結束日期-年-->
<INPUT type="hidden" name=S_MONTH value='<%if(session.getAttribute("S_MONTH") != null) out.print((String)session.getAttribute("S_MONTH"));%>'><!--//95.12.04儲存已勾選的起始日期-月-->
<INPUT type="hidden" name=E_MONTH value='<%if(session.getAttribute("E_MONTH") != null) out.print((String)session.getAttribute("E_MONTH"));%>'><!--//95.12.04儲存已勾選的結束日期-月-->
</form>

</body>
</html>
