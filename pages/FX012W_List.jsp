<%
//97.09.18 create 簽証會計師申報 by 2295
//99.03.01 fix 會計師姓名為edit Link.若未填時.顯示-.以供後續補資料用 by 2295
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
	List BOAF_ACCOUNT= (request.getAttribute("BOAF_ACCOUNT")==null)?null:(List)request.getAttribute("BOAF_ACCOUNT");
    String bank_no = ( request.getParameter("bank_no")==null ) ? "" : (String)request.getParameter("bank_no");
    Properties permission = ( session.getAttribute("FX012W")==null ) ? new Properties() : (Properties)session.getAttribute("FX012W");
    if (permission == null) {
        System.out.println("FX012W_List.permission == null");
    }else {
        System.out.println("FX012W_List.permission.size ="+permission.size());
    }
    String m_year="",bank_name="",name="",title="",addr="",telno="";    
	DataObject bean = null; 

%>
<script language="javascript" src="js/FX012W.js"></script>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>

<head>
<link href="css/b51.css" rel="stylesheet" type="text/css">
<title>簽證會計師申報</title>
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
                      <td width="300"><img src="images/banner_bg1.gif" width="300" height="17" align="left"></td>
                      <td width="215" align="center"><b><font size="4">簽證會計師申報作業</font></b></td>
                      <td width="300" align="right"><img src="images/banner_bg1.gif" width="300" height="17"></td>
                    </tr>
                  </table>
               </td>
              </tr>
              <tr>
              <td width="815" height="50">
                <div align="left">
                <table width="815" border="0"  cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="18" width="815" >               
                      </td>
                    </tr>
                    <tr>                       
                    <div><jsp:include page="getLoginUser.jsp?width=815" flush="true" /></div>			 							
			 		               
                    </tr>
                    <%if(permission != null && permission.get("A") != null && permission.get("A").equals("Y")){//Add %>
                    <tr>
                    <td width="815" align="right" valign="bottom">
                    <table border="1" cellspacing="1" bordercolor="#3A9D99" width="815"  class="sbody" cellpadding="0" height="20">
                      <tr>
                      <form name="date" method="post">                          				  		
                           <tr>
                           <td class="sbody" bgcolor="#E7E7E7" width="815"  valign="middle" height="20">
						  	 簽証會計師歷年明細表 
                     		 <a href="javascript:doSubmit(this.document.forms[0],'Print','<%=bank_no%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_print.gif',1);">
                    			<img src="images/bt_print.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>  
                    			&nbsp;&nbsp;
                    	      <a href="/pages/FX012W.jsp?act=new" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_addb.gif',1)">
                            <img src="images/bt_add.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>		
						   </td>
						   </tr>
                      </form>
                      </tr>
                    </table>
                    </td>
                    </tr>
                    <%}%>
					
					 <tr>
                      <td  align="center" width="815" valign="top">
					  <table class="sbody" width="815" border=1 cellpadding="0" cellspacing="1" bordercolor="#3A9D99" height="45">
                         <tr bgcolor="#9AD3D0">                          
                          <td width="32" align="center" height="1">年度</td>
                          <td width="87" align="center" height="2">會計師姓名</td>
                          <td width="208" align="center" height="2">事務所名稱</td>
                          <td width="345" align="center" height="2">地址</td>
                          <td width="120" align="center" height="2">電話</td>               
                         </tr>
                          
                         <% if(BOAF_ACCOUNT == null || BOAF_ACCOUNT.size()==0){%>
                         <tr class="sbody" bgcolor="#D8EFEE"><td colspan="5" align="center"  height="19" width="835" >
                           <font class="sbody">尚無資料</font></td></tr>
						 <%}else{
						      
						      for(int i=0;i<BOAF_ACCOUNT.size();i++){
						          m_year="";bank_no="";bank_name="";name="";title="";addr="";telno="";
						          bean = (DataObject)BOAF_ACCOUNT.get(i);
						          bank_no = bean.getValue("bank_no") == null?"":(String)bean.getValue("bank_no");
						          bank_name = bean.getValue("bank_name") == null?"":(String)bean.getValue("bank_name");
                                  name = bean.getValue("name") == null?"-":(String)bean.getValue("name");
                                  title = bean.getValue("title") == null?"":(String)bean.getValue("title");
                                  addr = bean.getValue("addr") == null?"":(String)bean.getValue("addr");
                                  telno = bean.getValue("telno") == null?"":(String)bean.getValue("telno");                                  
                                  m_year = bean.getValue("m_year") == null?"":bean.getValue("m_year").toString();                                
                                
						  %>					  
						  
						  
                          <tr class="sbody" bgcolor="<%out.print((i%2==0)?"#e7e7e7":"#D3EBE0");%>" width="835">                           
                             <td class="sbody"  align=center height="1" width="32"><%=m_year%>&nbsp;</td>
                             <td class="sbody"  align=center height="2" width="87"><a href=FX012W.jsp?act=Edit&m_year=<%=m_year%>&bank_no=<%=bank_no%>><%=name%></a>&nbsp;</td>
                             <td class="sbody"  align=center height="2" width="208"><%=title%>&nbsp;</td>
                             <td class="sbody"  align=center height="2" width="345"><%=addr%>&nbsp</td>
                             <td class="sbody"  align=center height="2" width="120"><%=telno%>&nbsp;</td>                            
                          </tr>
                          <%  }//end of for 
                           }%>
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
                          <li class="sbody">點選<FONT color=#666666>【</FONT>新增】按鈕可新增「簽証會計師」之資料。</li>
                          <li class="sbody">點選所列之[會計師姓名]可變更該申報之資料。 </li>                         
                        </ul>
                      </td>
                    </tr>
                  </table>
                  </td>
              </tr>

</table>
</form>
