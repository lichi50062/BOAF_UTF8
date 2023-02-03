<%
//94/10/17 first design by 4180
//95/04/04 依增修功能案會議記錄fix by 2495

%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.Date" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%
		List WLX05_ATM_SETUP= (request.getAttribute("WLX05_ATM_SETUP")==null)?null:(List)request.getAttribute("WLX05_ATM_SETUP");
    String bank_no = ( request.getParameter("bank_no")==null ) ? "" : (String)request.getParameter("bank_no");
    Properties permission = ( session.getAttribute("FX005AW")==null ) ? new Properties() : (Properties)session.getAttribute("FX005AW");
    if (permission == null) {
        System.out.println("FX005AW_List.permission == null");
    }
    else {
        System.out.println("FX005AW_List.permission.size ="+permission.size());
    }
String seq_no="";
String site_name="";
String property_no="";
String machine_name="";

String hsien_name="";
String area_name="";
String addr="";
String setup_date="";
String cancel_type="";
String cancel_date="";
String comment_m="";
String maintain_id="" ;
String maintain_name="" ;
String maintain_date="" ;
String sqlcmd="";
%>
<script language="javascript" src="js/FX005AW.js"></script>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>

<head>
<link href="css/b51.css" rel="stylesheet" type="text/css">
<title>各農漁會ATM裝設紀錄</title>
<script language="JavaScript" type="text/JavaScript">
<!--
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
<body topmargin="0" leftmargin="15">
<table width="815" border="0" align="left" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
        <tr>
            <td width="815"><img src="images/space_1.gif" width="12" height="12"></td>
            </tr>
        <tr>
          <td bgcolor="#FFFFFF" width="815">
        
              <tr>
                <td width="815" height="18">
                 <table width="815" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="200"><img src="images/banner_bg1.gif" width="200" height="17" align="left"></td>
                      <td width="400" align="center"><b><font size="4">各農漁會ATM裝設紀錄</font></b></td>
                      <td width="200" align="right"><img src="images/banner_bg1.gif" width="200" height="17"></td>
                    </tr>
                  </table>
               </td>
              </tr>
              <tr>
              <td width="815" height="50">
                <div align="left">
                <table width="815" border="0"  cellpadding="0" cellspacing="0">
                   <tr> 
                      <div align="right"><jsp:include page="getLoginUser.jsp?width=815" flush="true" /></div> 
                    </tr>
                    
        <%if(permission != null && permission.get("A") != null && permission.get("A").equals("Y")){//Add %>
                    <tr ><td width="815" align="right" valign="bottom">
          <table border="1" cellspacing="1" bordercolor="#3A9D99" width="815"  class="sbody" cellpadding="0" height="20">
            <tr>
            <form name="date" method="post">
                 <td class="sbody" bgcolor="#E7E7E7" width="815"  valign="middle" height="20">
                  <a href="/pages/FX005AW.jsp?act=new" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_addb.gif',1)">
                  <img src="images/bt_add.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>
                  </td>
            </form>
            </tr>
         </table>
         </td>
         </tr>
         <%}%>

                    <tr>
                      <td  align="center" width="815" valign="top">
						     <table class="sbody" width="815" border=1 cellpadding="0" cellspacing="1" bordercolor="#3A9D99" height="1" >
                          <tr  bgcolor="#9AD3D0">
                          <td  width="30" align="center" height="1" rowspan="2">
                          序號
                          </td>
                            <td   align="center" height="2" width="90" rowspan="2">
                              裝設地點
                          </td>
                            <td  align="center" height="2" width="200" rowspan="2">
                              裝設地址
                          </td>
                            <td  align="center" height="2" width="60" rowspan="2">
                              裝設日期
                          </td>
                            <td  align="center" height="2" width="60" rowspan="2">
                          機器編號
                          </td>
                            <td  align="center" height="2" width="60" rowspan="2">
                          機器品名
                          </td>
                           <td    align="center" height="1" width="100" colspan="2">遷移/裁撒
                          </td>
                            <td    align="center" height="1" width="80" rowspan="2">
                              備註
                          </td>
                          <td width="75" align="center" height="44" rowspan="2">
                          <p style="margin-top: 0; margin-bottom: 0">異動者</p>
                          <p style="margin-top: 0; margin-bottom: 0">帳號/</p>
                          <p style="margin-top: 0; margin-bottom: 0">姓名</p>
                          </td>
                          <td width="55" align="center" height="44" rowspan="2">異動日期</td>
                          </tr>
                          <tr  bgcolor="#9AD3D0">
                            <td  align="center" height="2" width="50">異動別
                          </td>
                            <td  align="center" height="2" width="50">日期
                          </td>
                          </tr>
										     <% if(WLX05_ATM_SETUP.size()==0){%>
                           <tr class="sbody" bgcolor="#D8EFEE"><td align="center" colspan="11" height="19" width="815" >
                           <font class="sbody">尚無資料</font></td></tr>
													<%
													}else{
													     for(int i=0;i<WLX05_ATM_SETUP.size();i++){
											     
seq_no = String.valueOf(((DataObject)WLX05_ATM_SETUP.get(i)).getValue("seq_no")); 
site_name=String.valueOf(((DataObject)WLX05_ATM_SETUP.get(i)).getValue("site_name"));
//property_no=String.valueOf(((DataObject)WLX05_ATM_SETUP.get(i)).getValue("property_no"))==null ? "" :String.valueOf(((DataObject)WLX05_ATM_SETUP.get(i)).getValue("property_no"));
//System.out.println("property_no ="+property_no );
property_no=String.valueOf(((DataObject)WLX05_ATM_SETUP.get(i)).getValue("property_no"));
machine_name=String.valueOf(((DataObject)WLX05_ATM_SETUP.get(i)).getValue("machine_name"));
hsien_name=String.valueOf(((DataObject)WLX05_ATM_SETUP.get(i)).getValue("hsien_name"));
area_name=String.valueOf(((DataObject)WLX05_ATM_SETUP.get(i)).getValue("area_name"));
addr=String.valueOf(((DataObject)WLX05_ATM_SETUP.get(i)).getValue("addr"));
setup_date=Utility.getCHTdate((((DataObject)WLX05_ATM_SETUP.get(i)).getValue("setup_date")).toString().substring(0, 10), 0);   
cancel_type=String.valueOf(((DataObject)WLX05_ATM_SETUP.get(i)).getValue("cancel_type"));
cancel_date=String.valueOf(((DataObject)WLX05_ATM_SETUP.get(i)).getValue("cancel_date"));   
comment_m=String.valueOf(((DataObject)WLX05_ATM_SETUP.get(i)).getValue("comment_m"));											
maintain_id=String.valueOf(((DataObject)WLX05_ATM_SETUP.get(i)).getValue("user_id"));
maintain_name=String.valueOf(((DataObject)WLX05_ATM_SETUP.get(i)).getValue("user_name"));
maintain_date=Utility.getCHTdate((((DataObject)WLX05_ATM_SETUP.get(i)).getValue("update_date")).toString().substring(0, 10), 0);   

cancel_date = (cancel_date.equals("null"))?"-":Utility.getCHTdate(cancel_date.substring(0, 10),0);     	
comment_m = (comment_m.equals("null"))?"-":comment_m;  
													%>
                          <tr class="sbody" bgcolor="<%out.print((i%2==0)?"#e7e7e7":"#D3EBE0");%>" width="815">
                          <td class="sbody"  align=center height="19" width="30">
                          <%out.print(String.valueOf(i+1));%>
                      	  </td>
                      	  <td class="sbody"  align=center height="19" width="90">
                      	   <u><a href=FX005AW.jsp?act=Edit&seq_no=<%=seq_no%>><%=site_name%></a></u>
                      	   </td>
                      	  <td class="sbody"  align=center height="19" width="200">
                      	  <%
                      	
                     	  		out.print(hsien_name+area_name+addr);
                      	  		
                      	  %>
                      	 
                      	  </td>
                      	  <td class="sbody"  align=center height="19" width="60">
                      	  <%=setup_date%>
                      	  </td>
                      	  <td class="sbody"  align=center height="19" width="60">
                      	  <%if(property_no.equals("null")){out.print("-");}
                      	    else{out.print(property_no);}
                      	  %>
                      	  </td>
                      	  <td class="sbody"  align=center height="19" width="60">
                      	  <%if(machine_name.equals("null")){out.print("-");}
                      	    else{out.print(machine_name);}
                      	  %>
                      	  </td>
                      	  <td class="sbody"  align=center height="19" width="50">
                      	  <%if(cancel_type.equals("1")){out.print("遷移");}
                      	    else if(cancel_type.equals("2")){out.print("裁撒");}
                      	    else{out.print("-");}
                      	  %>
                      	  </td>
                      	  <td class="sbody"  align=center height="19" width="50">
                      	  <%=cancel_date%>
                      	  </td>
                      	  <td class="sbody"  align=center height="19" width="80">
                      	  <%=comment_m%>
                      	  </td>
                      	  <td class="sbody"  align=center height="19" width="70">
                      	  <%=maintain_id%>/<%=maintain_name%>
                      	  </td>
                      	  <td class="sbody"  align=center height="19" width="55">
                      	  <%=maintain_date%>
                      	  </td>
                          </tr>
                          <% } }%>
                         </table>                 
                      </td>
                    </tr>


                    <td width="815" height="56"><div align="right"><div align="right">
                    <div align="center"><jsp:include page="getMaintainUser.jsp?width=815" flush="true" /></div>
                    </div>
                    <p align="left">　</div></td>
      </table>
                </div>
              </td>
  </tr>
  <tr>
                <td width="825" height="123"><table width="591" border="0" cellpadding="1" cellspacing="1" class="sbody">
                    <tr>
                      <td colspan="2" width="583"><font color='#990000'><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle"><font color="#007D7D" size="3">使用說明     
                        : </font></font></td>
                    </tr>
                    <tr>
                      <td width="16">&nbsp;</td>
                      <td class="sbody" width="561">
                        <ul>
                          <li class="sbody">點選<FONT color=#666666>【</FONT>新增】按鈕可新增「各農漁會ATM裝設紀錄」之資料。</li>
                          <li class="sbody">點選所列之[裝設地點]可變更該申報之資料。 </li>
                          <li class="sbody"><font color="red">本項作業須依據裝設異動建立歷史性資料，如有遷移/裁撤時請辦理遷移/裁撤異動，
											請勿以刪除方式處理。</font></li>
                        </ul>
                      </td>
                    </tr>
                  </table>
                  </td>
              </tr>

</table>
</form>
