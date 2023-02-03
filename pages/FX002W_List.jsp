<%
//94.03.08 fix 把快速查詢拿掉 by 2295
//94.04.06 fix 顯示裁撤日期 by 2295
//99.12.03 fix sqlInjection by 2808
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%
	
	List ba01Data = (List)request.getAttribute("ba01Data");
	List constTypeData = (List)request.getAttribute("constTypeData");
	List HsiendIdData = (List)request.getAttribute("HsiendIdData");
	List bn02RevokeData = (List)request.getAttribute("bn02RevokeData");
	String bank_type = ( request.getParameter("bank_type")==null ) ? "" : (String)request.getParameter("bank_type");			
	System.out.println("FX002W_List.bank_type="+bank_type);
	if(ba01Data != null){
	   System.out.println("ba01Data.size="+ba01Data.size());
	}else{
	   System.out.println("ba01Data == null");
	}
	if(constTypeData != null){
	   System.out.println("constTypeData.size="+constTypeData.size());
	}else{
	   System.out.println("constTypeData == null");
	}
	if(HsiendIdData != null){
	   System.out.println("HsiendIdData.size="+HsiendIdData.size());
	}else{
	   System.out.println("HsiendIdData == null");
	}
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/FX002W.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<title>國內營業分支機構基本資料維護</title>
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
<form method=post action='/pages/FX002W.jsp'>
<input type="hidden" name="act" value="List">  
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
                      <td width="150"><img src="images/banner_bg1.gif" width="150" height="17"></td>
                      <td width="300"><font color='#000000' size=4><b> 
                        <center>
                          國內營業分支機構基本資料維護 
                        </center>
                        </b></font> </td>
                      <td width="150"><img src="images/banner_bg1.gif" width="150" height="17"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><img src="images/space_1.gif" width="12" height="12"></td>
              </tr>
              <tr> 
                <td><table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
               		
                    <!--tr> 
                      <div align="right"><jsp:include page="getLoginUser.jsp" flush="true" /></div> 
                    </tr>
                    <tr> 
                      <td><table width=600 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
                          <tr class="sbody" bgcolor='#D2F0FF'>
						  <td width='30%' align='left'>快速查詢</td>
						  <td width='50%'>輸入分支機構代號:
      						<input type='text' name='QRY_BANK_NO' size=7 maxlength='7'>      						
      					  </td>	
      					  <td width='20%'>	
      						<div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Query','');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_queryb.gif',1)"><img src="images/bt_query.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a><div>
      					  </td>      					       					  
                          </tr>
                          </table>
                      </td>    
                    </tr-->  
                    <tr> 
                      <td><table width=600 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
                          <tr class="sbody" bgcolor="#9AD3D0">
						  <td colspan=2 align='center' >營運中之分支機構</td>						  
                          </tr>
                            
                          <tr class="sbody" bgcolor="#D8EFEE">                          
      					  <td width='30%' align=center>分支機構代號</td>
      					  <td width='70%' align=center>分支機構名稱</td>      					  						  
                          </tr>    
                          <%if(ba01Data == null || ba01Data.size() == 0){%>
                          <tr class="sbody" bgcolor="#e7e7e7">
                          <td  colspan=2 align=center>無營運中之分支機構</td>      					  
                          </tr>  
                          <%}else{
                             String bgcolor="#D3EBE0";
                          	 for(int i=0;i<ba01Data.size();i++){
                          	 bgcolor = (i % 2 == 0)?"#e7e7e7":"#D3EBE0";	
                          %>
                          <tr class="sbody" bgcolor="<%=bgcolor%>">                          
      					  <td width='30%' align=center>
      					    <a href='FX002W.jsp?act=Edit&tbank_no=<%=(((DataObject)ba01Data.get(i)).getValue("pbank_no")).toString()%>&bank_no=<%=(((DataObject)ba01Data.get(i)).getValue("bank_no")).toString()%>&bank_type=<%=bank_type%>'>
      					    <%=(String)((DataObject)ba01Data.get(i)).getValue("bank_no")%>
      					    </a>
      					  </td>
      					  <td width='75%' align=center>
      					    <a href='FX002W.jsp?act=Edit&tbank_no=<%=(((DataObject)ba01Data.get(i)).getValue("pbank_no")).toString()%>&bank_no=<%=(((DataObject)ba01Data.get(i)).getValue("bank_no")).toString()%>&bank_type=<%=bank_type%>'>
      					    <%=(String)((DataObject)ba01Data.get(i)).getValue("bank_name")%>
      					    </a>     					        					   
      					  </td>      					  						        					  
                          </tr>  
                          <%  }//end of for
                            }//end of if
                          %>
                         </table>
                      </td>
                    </tr>  
                    
                    <tr> 
                      <td><table width=600 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
                          <tr class="sbody" bgcolor="#9AD3D0">
						  <td colspan=5 align='center'>營運中之家數統計表</td>						  
                          </tr>
      					  
      					  <tr class="sbody" bgcolor="#D8EFEE">
        					<td width='20%' align=center>&nbsp;</td>
        					<% List paramList =new ArrayList() ;
        					   List const_type = DBManager.QueryDB_SQLParam("select cmuse_id,cmuse_name from cdshareno where cmuse_div='006' order by input_order",paramList,"");%>
        					<% for(int i=0;i<const_type.size();i++){%>
						  	<td width='20%' align=center><%=(String)((DataObject)const_type.get(i)).getValue("cmuse_name")%></td>
                            <% }%>                            
                            <td width='20%' align=center>合計</td>
      					  </tr>	
						  <%if(constTypeData != null && constTypeData.size() != 0){%>
						  <tr class="sbody" align=left>
        					<td width='20%' bgcolor="#D8EFEE">合計</td>               					 					
      						<td width='20%' bgcolor="#e7e7e7">&nbsp;<%=(((DataObject)constTypeData.get(0)).getValue("const_type1") == null) ? "" :(((DataObject)constTypeData.get(0)).getValue("const_type1")).toString()%></td>
      						<td width='20%' bgcolor="#e7e7e7">&nbsp;<%=(((DataObject)constTypeData.get(0)).getValue("const_type2") == null) ? "" :(((DataObject)constTypeData.get(0)).getValue("const_type2")).toString()%></td>
      						<td width='20%' bgcolor="#e7e7e7">&nbsp;<%=(((DataObject)constTypeData.get(0)).getValue("const_type9") == null) ? "" :(((DataObject)constTypeData.get(0)).getValue("const_type9")).toString()%></td>      						
      						<td width='20%' bgcolor="#e7e7e7">&nbsp;<%=(((DataObject)constTypeData.get(0)).getValue("const_tpyecount") == null) ? "" :(((DataObject)constTypeData.get(0)).getValue("const_tpyecount")).toString()%></td>      						      						
      						
      					  </tr>	
      					  <%}%>
      					   
      					  <%if(HsiendIdData != null && HsiendIdData.size() != 0){%>
      					  <tr class="sbody" align=left>
        					<td width='20%' bgcolor="#D8EFEE"><%=(String)((DataObject)HsiendIdData.get(0)).getValue("hsien_name")%></td>
      						<td width='20%' bgcolor="#e7e7e7">&nbsp;<%=(((DataObject)HsiendIdData.get(0)).getValue("const_type1") == null) ? "" :(((DataObject)HsiendIdData.get(0)).getValue("const_type1")).toString()%></td>
      						<td width='20%' bgcolor="#e7e7e7">&nbsp;<%=(((DataObject)HsiendIdData.get(0)).getValue("const_type2") == null) ? "" :(((DataObject)HsiendIdData.get(0)).getValue("const_type2")).toString()%></td>
      						<td width='20%' bgcolor="#e7e7e7">&nbsp;<%=(((DataObject)HsiendIdData.get(0)).getValue("const_type9") == null) ? "" :(((DataObject)HsiendIdData.get(0)).getValue("const_type9")).toString()%></td>
      						<td width='20%' bgcolor="#e7e7e7">&nbsp;<%=(((DataObject)HsiendIdData.get(0)).getValue("const_typecount") == null) ? "" :(((DataObject)HsiendIdData.get(0)).getValue("const_typecount")).toString()%></td>
      					  </tr>
      					  <%}%>	 
      					  
                          </table>
                      </td>    
                    </tr>  
                    <tr> 
                      <td><table width=600 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
                          <tr class="sbody" bgcolor="#9AD3D0">
						  <td colspan=5 align='center'>已裁撤之分支機構</td>						  
                          </tr>
                            
                          <tr class="sbody" bgcolor="#D8EFEE">                          
      					  <td width='30%' align=center>分支機構代號</td>
      					  <td width='40%' align=center>分支機構名稱</td>    
      					  <td width='30%' align=center>裁撤日期</td>       					 			  						  
                          </tr>    
                           <%if(bn02RevokeData == null || bn02RevokeData.size() == 0){%>
                          <tr class="sbody" bgcolor="#e7e7e7">
                          <td  colspan=3 align=center>無已裁撤之分支機構</td>      					  
                          </tr>  
                          <%}else{
                             String bgcolor="#D3EBE0";
                          	 for(int i=0;i<bn02RevokeData.size();i++){
                          	 bgcolor = (i % 2 == 0)?"#e7e7e7":"#D3EBE0";	
                          %>
                          <tr class="sbody" bgcolor="<%=bgcolor%>">                          
      					  <td width='30%' align=center>
      					    <a href='FX002W.jsp?act=Edit&tbank_no=<%=(((DataObject)bn02RevokeData.get(i)).getValue("tbank_no")).toString()%>&bank_no=<%=(((DataObject)bn02RevokeData.get(i)).getValue("bank_no")).toString()%>&bank_type=<%=bank_type%>'>
      					    <%=(String)((DataObject)bn02RevokeData.get(i)).getValue("bank_no")%>
      					    </a>
      					  </td>
      					  <td width='40%' align=center>
      					    <a href='FX002W.jsp?act=Edit&tbank_no=<%=(((DataObject)bn02RevokeData.get(i)).getValue("tbank_no")).toString()%>&bank_no=<%=(((DataObject)bn02RevokeData.get(i)).getValue("bank_no")).toString()%>&bank_type=<%=bank_type%>'>
      					    <%=(String)((DataObject)bn02RevokeData.get(i)).getValue("bank_name")%>
      					    </a>     					        					   
      					  </td> 
      					  <td width='30%' align=center>
      					    <a href='FX002W.jsp?act=Edit&tbank_no=<%=(((DataObject)bn02RevokeData.get(i)).getValue("tbank_no")).toString()%>&bank_no=<%=(((DataObject)bn02RevokeData.get(i)).getValue("bank_no")).toString()%>&bank_type=<%=bank_type%>'>      					  
      					    <%=Utility.getCHTdate((((DataObject)bn02RevokeData.get(i)).getValue("cancel_date")).toString().substring(0, 10), 0)%>    					  						        					  
      					    </a>
      					  </td>  
                          </tr>  
                          <%  }//end of for
                            }//end of if
                          %>
                         </table>
                      </td>
                    </tr>  
                    <td><div align="right"><jsp:include page="getMaintainUser.jsp" flush="true" /></div></td>                                              
      </table></td>
  </tr>
  <tr> 
                <td><table width="600" border="0" cellpadding="1" cellspacing="1" class="sbody">
                    <tr> 
                      <td colspan="2"><font color='#990000'><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle"><font color="#007D7D" size="3">使用說明 
                        : </font></font></td>
                    </tr>
                    <tr> 
                      <td width="16">&nbsp;</td>
                      <td width="577"> <ul>
                          <li>輸入欲查詢之分支機構代號，按<font color="#666666">【查詢】</font>即列出該筆分支機構清單。</li>
                          <li>點選所列之[分支機構代號]或[分支機構名稱]可變更該分支機構之資料。</li>                          
                        </ul></td>
                    </tr>
                  </table></td>
              </tr>
              <!--tr> 
                <td><div align="center"><img src="images/line_1.gif" width="600" height="12"></div></td>
              </tr-->
</table>
</form>
</body>
</html>
