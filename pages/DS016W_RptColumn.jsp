<%
//98.07.16 create A99 法定比率延申資料 by 2295
//99.05.25 fix 欄位選取改共用./include/DS_RptColumn.include by 2808
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
	String report_no = "DS016W";	
	String act = Utility.getTrimString(dataMap.get("act"));				
	String bank_type = (session.getAttribute("nowbank_type")==null)?"":(String)session.getAttribute("nowbank_type");	
	/*String title=(bank_type.equals("6"))?"農會":"漁會";*/
	String szExcelAction = (session.getAttribute("excelaction")==null)?"download":(String)session.getAttribute("excelaction");
	System.out.println("DS016W_RptColumn.szExcelAction="+szExcelAction);
	
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
   if(cnd == 'createRpt'){
      if(this.document.forms[0].BankList.value == ''){        
         alert('金融機構代碼必須選擇');
         return;
      }
      if(!chInput(this.document.forms[0],'')) return;//95.12.04 add 起始日期不可大於結束日期
      if(this.document.forms[0].FieldListDst.length == 0){        
      	 alert('報表欄位必須選擇');
      	 return;
      }
      if(!confirm("本項報表會執行10-15秒，是否確定執行？")){
         return;
      }   
   }   
   
   MoveSelectToBtn(this.document.forms[0].btnFieldList, this.document.forms[0].FieldListDst);		         
   fn_ShowPanel('<%=report_no%>',cnd);      
}

function ResetAllData(){
    if(confirm("確定要清除已選定的資料嗎？")){  	
        this.document.forms[0].FieldListDst.length = 0;        
        fn_loadFieldList(this.document.forms[0]);//顯示所有的報表欄位名稱
	}
	return;	
}

//-->
</script>
<link href="css/b51.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0">
<form method=post action='#'>
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
                <td width="*" class="title_font">A99法定比率延申資料 </td>                
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
                <td bordercolor="#E9F4E3" bgcolor="#E9F4E3"><table width="750" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#E9F4E3">
                    <tr> 
                      <td class="bt_bgcolor"> <div align="right">
                          <!--input type='radio' name="excelaction" value='view' <%if(szExcelAction.equals("view")){out.print("checked");}%> >檢視報表-->
                      	  <input type='radio' name="excelaction" value='download' <%if(szExcelAction.equals("download")){out.print("checked");}%> >下載報表
                          <%if(Utility.getPermission(request,report_no,"P")){//Print--有列印權限時 %>                   	        	                                   		     			        
                      	  <a href="javascript:doSubmit('createRpt');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image411','','images/bt_execb.gif',1)"><img src="images/bt_exec.gif" name="Image411" width="66" height="25" border="0" id="Image41"></a> 
                      	  <%}%>
                          <a href="javascript:ResetAllData();" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image511','','images/bt_cancelb.gif',0)"><img src="images/bt_cancel.gif" name="Image511" width="66" height="25" border="0" id="Image51"></a> 
                          <a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image611','','images/bt_reporthelpb.gif',1)"><img src="images/bt_reporthelp.gif" name="Image611" width="80" height="25" border="0" id="Image61"></a> 
                        </div></td>
                    </tr>
                    <tr> 
                      <td class="menu_bgcolor"> <table width="700" border="0" align="center" cellpadding="1" cellspacing="1" class="sbody">
                          <tr class="sbody"> 
                            <td width="100"><a href="javascript:doSubmit('BankList')"><font color='black'>1.金融機構</font></a></td>                            
                            <td width="100"><img src="images/2_icon_01.gif" width="16" height="16" align="absmiddle"> 
                              <a href="#"><font color="#CC6600">2.報表欄位</font></a></td>                           
                            <td width="100"><a href="javascript:doSubmit('RptOrder')"><font color='black'>3.排序欄位</font></a></td>
                            <td width="100"><a href="javascript:doSubmit('RptStyle')"><font color='black'>4.報表格式</font></a></td>
                          </tr>
                        </table></td>
                    </tr>
                    <%@include file="./include/DS_RptColumn.include" %><!-- 報表欄位.可挑選項目-->
                  </table></td>
              </tr>
            </table></td>
        </tr>
        
      </table>
    </td>
  </tr>
</table>
<INPUT type="hidden" name=FieldList value=<%=request.getAttribute("FieldList")%>><!--//FieldList儲存所有的報表欄位名稱-->
<INPUT type="hidden" name=btnFieldList><!--//btnFieldList儲存已勾選的報表欄位名稱-->
<INPUT type="hidden" name=BankList value='<%if(session.getAttribute("BankList") != null) out.print((String)session.getAttribute("BankList"));%>'><!--//BankList儲存已勾選的金融機構代碼-->
<INPUT type="hidden" name=S_YEAR value='<%if(session.getAttribute("S_YEAR") != null) out.print((String)session.getAttribute("S_YEAR"));%>'><!--//95.12.04儲存已勾選的起始日期-年-->
<INPUT type="hidden" name=E_YEAR value='<%if(session.getAttribute("E_YEAR") != null) out.print((String)session.getAttribute("E_YEAR"));%>'><!--//95.12.04儲存已勾選的結束日期-年-->
<INPUT type="hidden" name=S_MONTH value='<%if(session.getAttribute("S_MONTH") != null) out.print((String)session.getAttribute("S_MONTH"));%>'><!--//95.12.04儲存已勾選的起始日期-月-->
<INPUT type="hidden" name=E_MONTH value='<%if(session.getAttribute("E_MONTH") != null) out.print((String)session.getAttribute("E_MONTH"));%>'><!--//95.12.04儲存已勾選的結束日期-月-->
</form>
<script language="JavaScript" >
<!-- 
    
<%//從session裡把勾選的報表欄位讀出來.放在FieldListDst 	
if(session.getAttribute("btnFieldList") != null && !((String)session.getAttribute("btnFieldList")).equals("")){ 
   System.out.println("DS016W_RptColumn.btnFieldList="+(String)session.getAttribute("btnFieldList"));
%>
	var bnlist;
	bnlist = '<%=(String)session.getAttribute("btnFieldList")%>';
	<%
	session.setAttribute("btnFieldList",null);//清除已勾選的勾選的報表欄位
    %>
	var a = bnlist.split(',');
	for (var i =0; i < a.length; i ++){
		var j = a[i].split('+');
		this.document.forms[0].FieldListDst.options[i] = new Option(j[1], j[0]);
	}
<%}%>

fn_loadFieldList(this.document.forms[0]);//顯示所有的報表欄位名稱

-->
</script>

</body>
</html>
