<% 
//94.01.11 fix 更改密碼至少為 6碼 by 2295
//94.01.13 fix 更改密碼不可以帳號相同 by 2295
%>
<%@ page contentType="text/html;charset=Big5" %>
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
    
    if(trimString(form.ChangePwd.value) !=""  && form.ChangePwd.value.length < 6 ){
       alert("欲更改之密碼至少為6碼");
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
    <td height="20" bordercolor="#FFFFFF" bgcolor="#FFFFFF"><table width="764" border="0" align="left" cellpadding="0" cellspacing="0">
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
          <td rowspan="2" bgcolor="#70BEBB"><table width="370" border="0" align="center" cellpadding="0" cellspacing="1">
              <tr> 
                <td width="12"><img src="images/arrow_01.gif" width="9" height="9" align="absmiddle"></td>
                <td width="355" class="sbody">網際網路申報系統線上申請作業說明。</td>
              </tr>
              <tr> 
                <td valign="top"><img src="images/arrow_01.gif" width="9" height="9" align="absmiddle"></td>
                <td class="sbody">本系統全年全天候開放，惟須暫停連線服務時，將事先於本網站首頁公佈。</td>
              </tr>
            </table></td>
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
