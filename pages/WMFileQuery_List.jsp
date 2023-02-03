<%
//93.12.20 add 若有已點選的tbank_no,則以已點選的tbank_no為主
//94.04.15 fix 檢核為0時.顯示"內容值皆為零" by 2295
%>

<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="java.util.List" %>

<%
	List dbData = (List)request.getAttribute("dbData");
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");		
	//fix 93.12.20 若有已點選的tbank_no,則以已點選的tbank_no為主============================================================
	String bank_code = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");				
	String nowtbank_no =  ( request.getParameter("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");			
	if(!nowtbank_no.equals("")){
	   session.setAttribute("nowtbank_no",nowtbank_no);//將已點選的tbank_no寫入session	   
	}   
	bank_code = ( session.getAttribute("nowtbank_no")==null ) ? bank_code : (String)session.getAttribute("nowtbank_no");			
	//=======================================================================================================================	
	
	if(dbData != null){
	   System.out.println("dbData.size()="+dbData.size());
	}else{
		System.out.println("dbData  == null");
	}
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/WMFileEdit.js"></script>
<script language="javascript" event="onresize" for="window"></script>

<html>
<head>
<title>資料申報狀況查詢</title>
<link href="css/b51.css" rel="stylesheet" type="text/css">
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


<body marginwidth="0" marginheight="0" leftmargin="0" topmargin="0" leftmargin="0">
<form name=frmWMFileEdit method=post>
<input type="hidden" name="act" value=""> 
<table width="640" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  		<tr> 
   		 <td><img src="images/space_1.gif" width="12" height="12"></td>
  		</tr>

        <tr> 
          <td bgcolor="#FFFFFF">
			<table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="200"><img src="images/banner_bg1.gif" width="200" height="17"></td>
                      <td width="244"><font color='#000000' size=4><b> 
                        <center>
                          資料申報狀況查詢 
                        </center>
                        </b></font> </td>
                      <td width="200"><img src="images/banner_bg1.gif" width="200" height="17"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><img src="images/space_1.gif" width="12" height="12"></td>
              </tr>
              <tr> 
                <td><table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
               
                    <tr> 
                      <div align="right"><jsp:include page="getLoginUser.jsp" flush="true" /></div> 
                    </tr>                    
                  
                    <tr> 
                      <td><table width='600' border=1 align='center' cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
                          <tr bgcolor='#D2F0FF' class="sbody"> 
                            <td width=10% align='center' bordercolor="#3A9D99"><font color=#000000>報表代號</font></td>
                            <td width=10% align='center' bordercolor="#3A9D99"><font color=#000000>基準日</font></td>
                            <td width=20% align='center' bordercolor="#3A9D99"><font color=#000000>申報日期</font></td>
                            <td width=10% align='center' bordercolor="#3A9D99"><font color=#000000>申報方式</font></td>
                            <td width=10% align='center' bordercolor="#3A9D99"><font color=#000000>申報人員</font></td>
                            <td width=20% align='center' bordercolor="#3A9D99"><font color=#000000>檢核狀態</font></td>
                          </tr>
<%				
				if(dbData.size() == 0){%>
				 <tr bgcolor='#B1DEDC' class="sbody"> <td align=center colspan=3>無已申報過的資料</td></tr>
				<%
				}
				
				int i = 0;
				String input_method = "";				
				String upd_code="";
				String bgcolor="#FFFFE6";
				DataObject bean = null;
				while(i < dbData.size()){		
				      bgcolor = (i % 2 == 0)?"#FFFFE6":"#F2F2F2";		
				      bean = (DataObject)dbData.get(i);
				      if (((String)bean.getValue("input_method")) == null){
							input_method = "";
					  }else if (((String)bean.getValue("input_method")).equals("F")){
							input_method = "檔案上傳";
					  }else if (((String)bean.getValue("input_method")).equals("W")){
							input_method = "線上編輯";
					  }else{
							input_method = "";
					  }		
					  
					  if (((String)bean.getValue("upd_code")) == null){
					      upd_code = "待檢核";
					  }else if (((String)bean.getValue("upd_code")).equals("E")){
					      upd_code = "檢核有誤";
					  }else if (((String)bean.getValue("upd_code")).equals("U")){
						  upd_code = "檢核成功";
					  }else if (((String)bean.getValue("upd_code")).equals("F")){
					      upd_code = "上傳檔案不存在";
					  }else if (((String)bean.getValue("upd_code")).equals("Z")){
					      upd_code = "內容值皆為零";//94.04.15檢核為0時.顯示內容值皆為零   					  
					  }else{
						   upd_code = (String)bean.getValue("upd_code");
					  }
					  	
%>			


                            <tr bgcolor='<%=bgcolor%>' class="sbody"> 
                            <td align='center' bordercolor="#3A9D99"><%=(String)bean.getValue("report_no")%></td>                            
                            <td align='center' bordercolor="#3A9D99"><%=bean.getValue("inputdate").toString()%></td>                            
                       		<td align='center' bordercolor="#3A9D99"><%=bean.getValue("add_date").toString()%></td>
                            <td align='center' bordercolor="#3A9D99"><%=input_method%></td>
                            <td align='center' bordercolor="#3A9D99"><%=(String)bean.getValue("user_name")%></td>
                            <td align='center' bordercolor="#3A9D99">
                            <% if (upd_code != null && upd_code.equals("檢核有誤")){ %>                            
                            <a href='WMFileQuery_ShowError.jsp?Report_no=<%=(String)bean.getValue("report_no")%>&S_YEAR=<%=(bean.getValue("m_year")).toString()%>&S_MONTH=<%=(bean.getValue("m_month")).toString()%>&bank_code=<%=bank_code%>'><font color=red><%=upd_code%></font></a><img src="images/icon_1.gif" width="12" height="12" align="absmiddle">
                            <%}else{%>
                            <font color=black><%=upd_code%></font>
                            <%}%>
                            </td>
                            
                          </tr>
                          <%
					i++;
					}
				%>	                        
                        </table></td>
                    </tr>
                  </table></td>
              </tr>              
              
              <tr>
                <td><div align="center"><a href="javascript:history.back();" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image8','','images/bt_backb.gif',1)"><img src="images/bt_back.gif" name="Image8" width="80" height="25" border="0"></a></div></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
              </tr>
            </table></td>
        </tr>
       
      </table></td>
  </tr>
</table>
</form>
</body>
</html>
