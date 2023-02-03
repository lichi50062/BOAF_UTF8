<% 
// 94.01.11 fix 更改密碼至少為 6碼 by 2295
// 94.01.13 fix 更改密碼不可以帳號相同 by 2295
// 94.10.28 fix 新增公告區 by 2495
// 96.08.15 add 密碼必須為8碼的文數字組合 by 2295
// 98.09.17 fix 公告維護區 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>

<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.dao.DAOFactory" %>
<%@ page import="com.tradevan.util.dao.RdbCommonDao" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.Date" %>
<%@ page import="java.util.*" %>
<%@ page import="com.oreilly.servlet.MultipartRequest"%>
<%@ page import="java.util.List" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.math.BigDecimal" %>
<%@ page import="java.io.*" %>
<%@ page import="java.sql.*" %>

<%!
    //取得有效的公告資料
    private List Get_WLX_Notify(){    	
    	//查詢條件    
    	List paramList = new ArrayList () ;
    	String sqlCmd = "select seq_no,headmark,to_char(notify_date,'yyyy/mm/dd hh:mi') as notify_date,"
    				  + " to_char(notify_end_date,'yyyy/mm/dd') as notify_end_date,append_file,"
    				  + " user_id,user_name,to_char(update_date,'mm/dd/yyyy hh:mi:ss') as update_date,"
    				  + " appfile_link,notify_url from WLX_Notify "
    				  + " where to_char(notify_date,'yyyy/mm/dd hh') <= to_char(sysdate,'yyyy/mm/dd hh') "
 				      + " and to_char(notify_end_date,'yyyy/mm/dd hh') >= to_char(sysdate,'yyyy/mm/dd hh') "
    				  + " order by notify_date desc";  
        //System.out.println("sqlCmd="+sqlCmd);    				  
    	List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"seq_no");      	
        return dbData;
    }
%>


<html>
<head>
<title>網際網路申報系統</title>
<link href="css/b51.css" rel="stylesheet" type="text/css">
<script language="javascript" src="js/Common.js"></script>
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
function doSubmit(form){
    if(trimString(form.muser_id.value) =="" ){
       alert("用戶帳號不可為空白");
       form.muser_id.focus();
       return;
    }
    if(trimString(form.muser_password.value) =="" ){
       alert("用戶密碼不可為空白");
       form.muser_password.focus();
       return;
    }else{            
       if((trimString(form.ChangePwd.value) !="" ) && (form.muser_password.value == form.ChangePwd.value)){
          alert("卻更改之密碼不可與舊密碼相同");
          form.ChangePwd.focus();
          return;
       }       
    }
    
    if((trimString(form.ChangePwd.value) !="" ) || (trimString(form.ConfirmPwd.value) !="" )){
       if(form.ChangePwd.value != form.ConfirmPwd.value){
          alert("欲更改之密碼與確認密碼不符合");
          form.ConfirmPwd.focus();
          return;
       }
    }
    
    if(trimString(form.ChangePwd.value) !=""  && form.ChangePwd.value.length < 8 ){
       alert("欲更改之密碼至少為8碼的文數字組合");
       form.ChangePwd.focus();
       return;
    }
    if(trimString(form.ChangePwd.value) !="" && !checkPwd(form.ChangePwd.value)){
       alert("欲更改之密碼必須為文數字組合");
       form.ChangePwd.focus();
       return;
    }
    if(trimString(form.ChangePwd.value) !="" && (trimString(form.ChangePwd.value) == trimString(form.muser_id.value))){
       alert("欲更改之密碼不可與帳號相同");     
       form.ChangePwd.focus();
       return;   
    }
    form.submit();
}

function checkPwd(pwd){
   letter = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
   number = "0123456789";
   var count_letter=0;
   var count_num=0;
   //alert(pwd.length);
   for(i =0;i<pwd.length;i++){
       j = letter.indexOf( pwd.charAt(i).toUpperCase() );
       if ( j != -1 ) count_letter=count_letter+1;
   }
   //alert('count_letter='+count_letter);
   if ( count_letter <= 0 ){        
        return false;
   }
   
   for(i =0;i<pwd.length;i++){
       j = number.indexOf( pwd.charAt(i) );
       if ( j != -1 ) count_num=count_num+1;
   }
   if ( count_num <= 0 ){        
        return false;
   }
   //alert('count_num='+count_num);
   return true;
}
function doLink(url){
	
	window.open(url); 
}

var keyPressed;

function chkKey(e){  
    keyPressed = String.fromCharCode(window.event.keyCode);
    if (keyPressed == "\x0D") {
      doSubmit(window.document.loginfrm);
    }  
}

if (window.document.captureEvents!=null) 
  window.document.captureEvents(Event.KEYPRESS)
window.document.onkeypress = chkKey;


//-->

</script>
</head>

<body background="images/bg_1.gif" leftmargin="0" topmargin="0" onLoad="">
<form name="loginfrm" method=post action='/pages/Login.jsp'>
<table width="764" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#006666">
  <tr>
    <td height="20" bordercolor="#FFFFFF" bgcolor="#FFFFFF"><table width="764" border="0" align="left" cellpadding="0" cellspacing="0" bgcolor="#70BEBB">
          <tr> 
          <td width="68"><img src="images/Login_Image_01.gif" width="68" height="23"></td>
          <td width="261" bgcolor="#4B9C99">&nbsp;</td>
          <td width="34" bgcolor="#4B9C99">&nbsp;</td>
          <td width="396" bgcolor="#4B9C99">&nbsp;</td>
          <td width="21"><img src="images/Login_Image_02.gif" width="21" height="23"></td>
        </tr>
        <tr> 
          <td><img src="images/Login_Image_03.gif" width="68" height="233"></td>
          <td><img src="images/Login_Image_04.gif" width="261" height="233"></td>
          <td><img src="images/Login_Image_05.gif" width="34" height="233"></td>
          <td><img src="images/Login_Image_06.gif" width="380" height="233"></td>
          <td><img src="images/Login_Image_07.gif" width="21" height="233"></td>
        </tr>
        <tr> 
          <td><img src="images/Login_Image_08.gif" width="68" height="69"></td>
          <td><img src="images/Login_Image_09.gif" width="261" height="69"></td>
          <td><img src="images/Login_Image_10.gif" width="34" height="69"></td>
          <td rowspan="2">
            <table width="330" border="0" align="center" cellpadding="0" cellspacing="0">													              
              <tr>
                <td><table width="330" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#006666">
                    <tr>
                    	
        <td bordercolor="#63B1AE"><table width="330" border="0" align="center" cellpadding="1" cellspacing="0" bordercolor="#FFFFFF">                                           
		<%
		 List WLX_Notify = Get_WLX_Notify();
		 //List WLX_Notify = null;//關閉公告
		 
         if(WLX_Notify != null && WLX_Notify.size() != 0 && WLX_Notify.size() > 0){
              System.out.println("有效公告:"+WLX_Notify.size());
        %>
	<tr colspan="0" bordercolor="#63B1AE" bgcolor="#408C87"><img src="images/hotnews.gif" width="70" height="18"></tr>	
           <script language=JavaScript> 
           document.write("<marquee  bgcolor=#70BEBB scrollamount='1' scrolldelay='6' direction= 'up' width='380' id=xiaoqing height='110' onmouseover=xiaoqing.stop() onmouseout=xiaoqing.start()>");
           <%
            DataObject bean = null;
            String notify_date = "";
            String fontColor="black";
            String headmark="";
            for(int i=0;i<WLX_Notify.size();i++){
                bean = (DataObject)WLX_Notify.get(i);
                notify_date = (String)bean.getValue("notify_date");
                headmark = (String)bean.getValue("headmark");
                notify_date = Utility.getCHTdate(notify_date.substring(0,10),0)+notify_date.substring(10,13)+":00";
                fontColor=((i == 0)?"white":"black");//最新一筆.以藍色顯示
                if(headmark.indexOf("異常") != -1) fontColor="yellow";//系統異常訊息.以黃色顯示
                      
		    	if(bean.getValue("append_file") != null && !((String)bean.getValue("append_file")).trim().equals("")){//有附加檔案%>
				   document.write("&nbsp;<a href=\"javascript:doLink('/pages/DomloadNotify.jsp?seq_no=<%=bean.getValue("seq_no").toString()%>');\"><img src=images/download.gif border=0></a>&nbsp;<font size='2' "); 		 									   
				   <%if(bean.getValue("notify_url") != null){ //有URL%>
				   	   document.write("color=<%=fontColor%>>"); 
				       document.write("<%=notify_date%>&nbsp;<a href=\"javascript:doLink('<%=(String)bean.getValue("notify_url")%>');\"><%=headmark%></a></font><br>"); 	
				   <%}else{//無URL%>
				       document.write("color=<%=fontColor%>><%=notify_date%>&nbsp;<%=headmark%></font><br>");	
				   <%}
				}else{//無附加檔案%>
				   document.write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font size='2'");
				   <%if(bean.getValue("notify_url") != null){ //有URL%>
				       document.write(" color=<%=fontColor%>><%=notify_date%>&nbsp;<a href=\"javascript:doLink('<%=(String)bean.getValue("notify_url")%>');\"><%=headmark%></a></font><br>"); 
				   <%}else{//無URL%>
				   	document.write(" color=<%=fontColor%>><%=notify_date%>&nbsp;<%=headmark%></font><br>");							 									
				   <%}
				}//end of 無附加檔案	
		    }//end of WLX_Notify		    
            %>		
			document.write("</marquee>");				
			</script>
		<%}else{%>
		<tr> 
                <td width="12"><img src="images/arrow_01.gif" width="9" height="9" align="absmiddle"></td>
                <td width="355" class="sbody">網際網路申報系統線上申請作業說明。</td>
              </tr>
              <tr> 
                <td valign="top"><img src="images/arrow_01.gif" width="9" height="9" align="absmiddle"></td>
                <td class="sbody">本系統全年全天候開放，惟須暫停連線服務時，將事先於本網站首頁公佈。</td>
              </tr>
		<%	}%>
                          </tr>	
			                    </table>
                     </td>
                    </tr>
                    
                  </table></td>
              </tr>
              
            </table>
	  </td>
          <td bgcolor="#70BEBB">&nbsp;</td>
        </tr>
        <tr> 
          <td background="images/Login_Image_11.gif"><img src="images/Login_Image_11.gif" width="68" height="100"></td>
          <td rowspan="2" valign="top" bgcolor="#78B5B3">
			<table width="250" border="0" align="center" cellpadding="0" cellspacing="1" class="sbody">
              <tr> 
                <td>用戶帳號:</td>
                <td><input type="text" maxlength=12 name=muser_id size=12></td>                
              </tr>
              <tr> 
                <td width="65">用戶密碼 :</td>
                <td width="182"><input type="password" maxlength=20 name=muser_password size=20></td>
              </tr>              
              <tr> 
                <td width="65">更改密碼 :</td>
                <td width="182"><input type="password" maxlength=20 name=ChangePwd size=20></td>
              </tr>
              <tr> 
                <td width="65">確認密碼 :</td>
                <td width="182"><input type="password" maxlength=20 name=ConfirmPwd size=20></td>
              </tr>
              
              <tr>
                <td>&nbsp;</td>
                <td><a href="javascript:doSubmit(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image23','','images/Login_bt_3b.gif',1)"><img src="images/Login_bt_3.gif" name="Image23" width="58" height="22" border="0"></a> 
                  <a href="javascript:AskReset(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image20','','images/Login_bt_1b.gif',1)"><img src="images/Login_bt_1.gif" name="Image20" width="58" height="22" border="0"></a></td>
              </tr>
            </table>
          </td>
          <td background="images/Login_Image_12.gif"><img src="images/Login_Image_12.gif" width="34" height="100"></td>
          <td bgcolor="#70BEBB">&nbsp;</td>
        </tr>
        <tr> 
          <td background="images/Login_Image_11.gif"><img src="images/Login_Image_11.gif" width="68" height="42"></td>
          <td background="images/Login_Image_11b.gif"><img src="images/Login_Image_11b.gif" width="34" height="42"></td>
          <td bgcolor="#4B9C99"><table width="366" border="0" align="center" cellpadding="0" cellspacing="0" class="sbody">
              <tr> 
                <td>建議使用IE 5.0以上版本之瀏覽器螢幕解析度800X600以上瀏覽</td>
              </tr>
              <tr>
                <td>版權所有 翻版必究 Copyringht 2004<a href="#"> BOAF </a>All Rights 
                  Reserved .</td>
              </tr>
            </table></td>
          <td bgcolor="#4B9C99">&nbsp;</td>
        </tr>
        <tr> 
          <td><img src="images/Login_Image_13.gif" width="68" height="18"></td>
          <td><img src="images/Login_Image_14.gif" width="261" height="18"></td>
          <td><img src="images/Login_Image_15.gif" width="34" height="18"></td>
          <td bgcolor="#4B9C99"><img src="images/space_1.gif" width="5" height="12"></td>
          <td><img src="images/Login_Image_16.gif" width="21" height="18"></td>
        </tr><a name="start">        
      </table></td>
  </tr>
</table>
</form>
</body>
<script language="javascript">
  window.scrollTo(window.pageXOffset,1200);
</script>

</html>
