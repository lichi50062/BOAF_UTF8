<%
// 96.01.04 fix 調整畫面 by 2295
// 96.01.16 add 是否需重新至金庫取檔 by 2295
//103.04.02 add 專案農貸資料 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Calendar" %>

<%
	Calendar now = Calendar.getInstance();
   	String S_YEAR  = String.valueOf(now.get(Calendar.YEAR)-1911); //回覆值為西元年故需-1911取得民國年;
   	String S_MONTH = String.valueOf(now.get(Calendar.MONTH)+1);   //月份以0開始故加1取得實際月份;
    if(S_MONTH.equals("1")){//若本月為1月份是..則是申報上個年度的12月份
       S_YEAR = String.valueOf(Integer.parseInt(S_YEAR) - 1);
       S_MONTH = "12";
    }else{    
      S_MONTH = String.valueOf(Integer.parseInt(S_MONTH) - 1);//申報上個月份的
    }
%>    

<script language="javascript" src="js/Common.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<title>農業金庫傳輸檔案</title>
<link href="css/b51.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
function doSubmit(form,cnd){
	if(!(form.candidate[0].checked ||form.candidate[1].checked ||form.candidate[2].checked ||form.candidate[3].checked ||form.candidate[4].checked ||form.candidate[5].checked)){
		alert("請選擇傳輸檔案")  
	    return;
	}				
	if(!confirm("本項傳輸檔案會執行5~10分鐘，是否確定執行？"))	return;			
	form.action="/pages/ZZ092W.jsp?act="+cnd+"&year="+form.S_YEAR.value+"&month="+form.S_MONTH.value+"&getFiles="+form.getFiles.checked;   			
	form.submit();	    	  	    
}	

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
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
//-->
</script>
</head>

<body marginwidth="0" marginheight="0" leftmargin="0" topmargin="0" leftmargin="0">
<form method=post action='#'>
<input type="hidden" name="act" value="">   
<table width="640" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  		<tr> 
   		 <td><img src="images/space_1.gif" width="12" height="12"></td>
  		</tr>

        <tr> 
          <td bgcolor="#FFFFFF">
		  <table width="640" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="640" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="170"><img src="images/banner_bg1.gif" width="170" height="17"></td>
                      <td width="300"><font color='#000000' size=4><b> 
                        <center>農業金庫傳輸檔案</center>
                        </b></font> </td>
                      <td width="170"><img src="images/banner_bg1.gif" width="170" height="17"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><img src="images/space_1.gif" width="12" height="12"></td>
              </tr>
              <tr> 
                <td><table width="640" border="0" align="center" cellpadding="0" cellspacing="0">               
                    <tr> 
                      <div align="right"><jsp:include page="getLoginUser.jsp?width=640" flush="true" /></div> 
                    </tr>     
                     <%
                      String nameColor="nameColor_sbody";
                      String textColor="textColor_sbody";
                      String bordercolor="#3A9D99";
                     %>               
                    <tr> 
                      <td><table width=640 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="<%=bordercolor%>">                          
                          <tr class="sbody">
                          <td width='10%' class="<%=nameColor%>">年月</td>	
                          <td width='90%' bgcolor='e7e7e7'>	                          	  
                            <input type='text' name='S_YEAR' value="<%=S_YEAR%>" size='3' maxlength='3' onblur='CheckYear(this)'>
        						<font color='#000000'>年</font>			
        						<select id="hide1" name=S_MONTH>        						
        						<%
        							for (int j = 1; j <= 12; j++) {
        							if (j < 10){%>        	
        							<option value=0<%=j%> <%if(S_MONTH.equals(String.valueOf(j))) out.print("selected");%>>0<%=j%></option>        		
            						<%}else{%>
            						<option value=<%=j%> <%if(S_MONTH.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
            						<%}%>
        						<%}%>
        						</select><font color='#000000'>月</font>			
        						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        						<input type="checkbox" name="getFiles" value="">需重新至金庫取檔
                          </td>                									                           
						  </tr>						  
                          </table>      
                      </td>                          
                      </tr>
                      <tr> 
                      	<td><table width=640 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="<%=bordercolor%>">                          
                          <tr class="sbody">
                          <td width='10%' class="<%=nameColor%>">檔案</td>	
                          <td width='90%' bgcolor='e7e7e7'>	
                    			<input type="checkbox" name="candidate" value="ATM">金融卡發卡及ATM裝設
       							<input type="checkbox" name="candidate" value="CHK">支票存款辦理
       							<input type="checkbox" name="candidate" value="FGN">外國人存款
       							<input type="checkbox" name="candidate" value="INT">牌告利率<br>     
       							<input type="checkbox" name="candidate" value="FRM">專案農貸   
       							<input type="checkbox" name="candidate" value="ITEM">貸款項目代碼 
       							&nbsp;&nbsp;&nbsp;&nbsp;
                                <a href="javascript:doSubmit(this.document.forms[0],'download');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_queryb.gif',1)"><img src="images/bt_query.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>                                    							
               			  </td>                
						  </tr>						  
                          </table>      
                      	</td>                          
                      </tr>
</table>
</form>
</body>
</html>
