<%
//93.12.21 add 權限檢核 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%
	
	List WLX04 = (List)request.getAttribute("WLX04");
	String bank_no = ( request.getParameter("bank_no")==null ) ? "" : (String)request.getParameter("bank_no");		
	//取得FX003W的權限
	Properties permission = ( session.getAttribute("FX003W")==null ) ? new Properties() : (Properties)session.getAttribute("FX003W"); 
	if(permission == null){
       System.out.println("FX003W_Edit.permission == null");
    }else{
       System.out.println("FX003W_Edit.permission.size ="+permission.size());               
    }	
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/FX003W.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<title>理監事基本資料維護</title>
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
<form method=post>
<input type="hidden" name="act" value="Edit">  
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
                      <td width="165"><img src="images/banner_bg1.gif" width="165" height="17"></td>
                      <td width="270"><font color='#000000' size=4><b> 
                        <center>
                          理監事基本資料維護 
                        </center>
                        </b></font> </td>
                      <td width="165"><img src="images/banner_bg1.gif" width="165" height="17"></td>
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
                      <td><table width=600 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
                          <tr class="sbody" bgcolor='e7e7e7'>						
      					  <td width='100%'>
      					  <%if(permission != null && permission.get("A") != null && permission.get("A").equals("Y")){//Add %>                   	        	                                   		     
      						<div align="left"><a href="FX003W.jsp?act=new" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_addb.gif',1)"><img src="images/bt_add.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a><div>
      				      <%}%>		
      					  </td>      					       					  
                          </tr>
                          </table>
                      </td>    
                    </tr>  
                    <tr> 
                      <td><table width=600 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
                          <tr class="sbody" bgcolor="#9AD3D0">
                          <td width='15%' align=center>合計</td>
      					  <td width='20%' align=center>理監事</td>
      					  <td width='65%' align=center>自然人姓名/代表人姓名</td>      					  						  
                          </tr>    
                          <%if(WLX04.size() == 0){%>
                          <tr class="sbody" bgcolor="#e7e7e7">
                          <td colspan="3" align=center>查無資料</td>
                          </tr>
                          <%}else{
                             int count1=0,count2=0;
                             String bgcolor="#D3EBE0";
                             for(int i=0;i<WLX04.size();i++){
                                 if(((String)((DataObject)WLX04.get(i)).getValue("abdicate_code") == null) || ((String)((DataObject)WLX04.get(i)).getValue("abdicate_code")).equals("N")){        				  
                                     bgcolor = (i % 2 == 0)?"#e7e7e7":"#D3EBE0";
                                     if(((String)((DataObject)WLX04.get(i)).getValue("position_code")).equals("1")){
                                          count1++ ;
                          %>
                          <tr class="sbody" bgcolor="<%=bgcolor%>">
                          <td width='15%' align=center>&nbsp;</td>
      					  <td width='20%' align=center><%=(String)((DataObject)WLX04.get(i)).getValue("cmuse_name")%></td>
      					  <td width='65%' align=center>
      					  <a href='FX003W.jsp?act=Edit&bank_no=<%=bank_no%>&POSITION_CODE=<%=(((DataObject)WLX04.get(i)).getValue("position_code") == null) ?"&nbsp;":(String)((DataObject)WLX04.get(i)).getValue("position_code")%>&seq_no=<%=(((DataObject)WLX04.get(i)).getValue("seq_no")).toString()%>'>
      					  <%=(String)((DataObject)WLX04.get(i)).getValue("name")%>
      					  </td>      					  						  
                          </tr>  
                          
                          <%     
                                     }//end of 理事
                                }//end of 未卸任
                             }//end for
                          %>   
                          <%if(count1 > 0){%> 
                          <tr class="sbody" bgcolor="#D8EFEE">
                          <td width='15%' align=center>合計:<%=count1%></td>
      					  <td width='20%' align=center>&nbsp;</td>
      					  <td width='65%' align=center>&nbsp;</td>      					  						  
      					  </tr>
      					  <%}%>
      					  <%  
      					  	  bgcolor="#D3EBE0";	
      					   	  for(int i=0;i<WLX04.size();i++){
      					   	      if(((String)((DataObject)WLX04.get(i)).getValue("abdicate_code") == null) || ((String)((DataObject)WLX04.get(i)).getValue("abdicate_code")).equals("N")){        				  
      					   	  	       bgcolor = (i % 2 == 0)?"#e7e7e7":"#D3EBE0";
                                       if(((String)((DataObject)WLX04.get(i)).getValue("position_code")).equals("2")){
                                            count2++ ; 
                          %>
      					  <tr class="sbody" bgcolor="<%=bgcolor%>">
                          <td width='15%' align=center>&nbsp;</td>
      					  <td width='20%' align=center><%=(String)((DataObject)WLX04.get(i)).getValue("cmuse_name")%></td>
      					  <td width='65%' align=center>
      					  <a href='FX003W.jsp?act=Edit&bank_no=<%=bank_no%>&POSITION_CODE=<%=(((DataObject)WLX04.get(i)).getValue("position_code") == null) ?"&nbsp;":(String)((DataObject)WLX04.get(i)).getValue("position_code")%>&seq_no=<%=(((DataObject)WLX04.get(i)).getValue("seq_no")).toString()%>'>
      					  <%=(String)((DataObject)WLX04.get(i)).getValue("name")%>
      					  </a>
      					  </td>      					  						       					  						  
                          </tr> 
                          <%           }//end of 監事
                                 }//end of 未卸任
                              }//end for
                          %>
                          <%if(count2 > 0){%>
      					  <tr class="sbody" bgcolor="#D8EFEE">
                          <td width='15%' align=center>合計:<%=count2%></td>
      					  <td width='20%' align=center>&nbsp;</td>
      					  <td width='65%' align=center>&nbsp;</td>      					  						  
                          </tr>  
                          <%}%>
                          <%
                          }//end else
                          %>
                         </table>
                      </td>
                    </tr>  
                    <tr>
                <td><table width=600 border=1 align='center' cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
                    <tr bgcolor="#9AD3D0" class="sbody"> 
                        <td colspan=4 bgcolor=#9AD3D0><font face=細明體 color=#000000>卸任理監事基本資料</font></td>                    	
                    </tr>
                    <tr class="sbody">
						<td bgcolor='#D8EFEE' width='25%' bgcolor='#D8EFEE' align='center'>職稱</td>
						<td bgcolor='#D8EFEE' width='25%' bgcolor='#D8EFEE' align='center'>姓名</td>
						<td bgcolor='#D8EFEE' width='25%' bgcolor='#D8EFEE' align='center'>就任日期</td>
						<td bgcolor='#D8EFEE' width='25%' bgcolor='#D8EFEE' align='center'>卸任日期</td>						
        			</tr>
        			
        			<%if(WLX04.size() == 0){%>
        			    <tr class="sbody">
        				<td bgcolor='e7e7e7' width='100%' colspan=4 align='center'>無卸任之理監事基本資料</td>
        				</tr>
        			<%}else{
        			    boolean haveabdicate=false;
        				for(int i=0;i<WLX04.size();i++){
							if(((String)((DataObject)WLX04.get(i)).getValue("abdicate_code") != null) && ((String)((DataObject)WLX04.get(i)).getValue("abdicate_code")).equals("Y")){        				  
							    haveabdicate=true;
        			%>
        			    <tr class="sbody">
						<td bgcolor='e7e7e7' width='25%' bgcolor='#D8EFEE' align='center'><%=(String)((DataObject)WLX04.get(i)).getValue("cmuse_name")%></td>
						<td bgcolor='e7e7e7' width='25%' bgcolor='#D8EFEE' align='center'><a href='FX003W.jsp?act=Edit&bank_no=<%=bank_no%>&POSITION_CODE=<%=(((DataObject)WLX04.get(i)).getValue("position_code") == null) ?"&nbsp;":(String)((DataObject)WLX04.get(i)).getValue("position_code")%>&seq_no=<%=(((DataObject)WLX04.get(i)).getValue("seq_no")).toString()%>'><%=(String)((DataObject)WLX04.get(i)).getValue("name")%></td>
						<td bgcolor='e7e7e7' width='25%' bgcolor='#D8EFEE' align='center'><%=Utility.getCHTdate((((DataObject)WLX04.get(i)).getValue("induct_date")).toString().substring(0, 10), 0)%></td>						        										
						<td bgcolor='e7e7e7' width='25%' bgcolor='#D8EFEE' align='center'><%=Utility.getCHTdate((((DataObject)WLX04.get(i)).getValue("abdicate_date")).toString().substring(0, 10), 0)%></td>						        																
						</tr>
						
					<%      }//end if
						}//end for%>
					   	<%if(!haveabdicate){%>        				
        				<tr class="sbody">
        				<td bgcolor='e7e7e7' width='100%' colspan=4 align='center'>無卸任之高階主管相關資料</td>
        				</tr>
        			    <%}//haveabdicate=false
					  }//end of else
					%>	        			
                </table></td>
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
                          <li>點選所列之<font color="#666666">【新增】</font>可新增其他理監事之資料。</li>
                          <li>點選所列之[姓名或法人名稱]可變更該理監事之資料。</li>                          
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
