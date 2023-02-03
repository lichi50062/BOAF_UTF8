<%
//94/10/17 first design by 4180
//99.12.03 fix sqlInjection by 2808
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.Date" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.ArrayList" %>
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
String hsien_id="";
String area_id="";
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
List paramList = new ArrayList() ;
%>
<script language="javascript" src="js/FX005AW.js"></script>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>

<head>
<link href="css/b51.css" rel="stylesheet" type="text/css">
<title>各農漁會ATM裝設紀錄</title>
</head>
<body topmargin="15" leftmargin="15">
<form method="post" name="loaddata">
<input type="button" value="確認載入" onClick="javascript:doSubmit(this.document.forms[0],'checkload');">
<table width="815" border="0" align="left" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
        <tr>
            <td width="815"><img src="images/space_1.gif" width="12" height="12"></td>
            </tr>    
                 <tr>
                      <td  align="center" width="815" valign="top">
						     <table class="sbody" width="815" border=1 cellpadding="0" cellspacing="1" bordercolor="#3A9D99" height="1" >
                          <tr  bgcolor="#9AD3D0">
                          <td  width="30" align="center" height="1" rowspan="2">
                          確認
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
property_no=String.valueOf(((DataObject)WLX05_ATM_SETUP.get(i)).getValue("property_no"));
machine_name=String.valueOf(((DataObject)WLX05_ATM_SETUP.get(i)).getValue("machine_name"));
hsien_id=String.valueOf(((DataObject)WLX05_ATM_SETUP.get(i)).getValue("hsien_id"));
area_id=String.valueOf(((DataObject)WLX05_ATM_SETUP.get(i)).getValue("area_id"));
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
                           <input type="checkbox" name="CHECKSEQ" value="<%=seq_no%>">
                      	  </td>
                      	  <td class="sbody"  align=center height="19" width="90">
                      	   <%=site_name%>                	  
                      	   </td>
                      	  <td class="sbody"  align=center height="19" width="200">
                      	  <%
                      	  //99.12.08 fix 縣市合併問題 by 2808
                      	  String cd01Table = Integer.parseInt(Utility.getYear()) > 99 ?"cd01" : "cd01_99" ;
                      	  sqlcmd = "Select * From  "+cd01Table+" CD01, CD02 "+
							 	" Where  CD01.HSIEN_ID=? And CD02.AREA_ID=? ";
                      	  paramList.add(hsien_id) ;
                      	  paramList.add(area_id);
								List hsien_id_area_id = DBManager.QueryDB_SQLParam(sqlcmd,paramList,"hsien_name,area_name");
								System.out.print(sqlcmd);
							  if(hsien_id_area_id.size()!=0){
								hsien_name = String.valueOf(((DataObject)hsien_id_area_id.get(0)).getValue("hsien_name"));
                      	  		area_name = String.valueOf(((DataObject)hsien_id_area_id.get(0)).getValue("area_name"));
                      	  		out.print(hsien_name+area_name+addr);
                      	  		}
                      	  		else
                      	  		{
                      	  		out.print(addr);
                      	  		}
                      	  %>
                      	 
                      	  </td>
                      	  <td class="sbody"  align=center height="19" width="60">
                      	  <%=setup_date%>
                      	  </td>
                      	  <td class="sbody"  align=center height="19" width="60">
                      	  <%=property_no%>
                      	  </td>
                      	  <td class="sbody"  align=center height="19" width="60">
                      	  <%=machine_name%>
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
      </table>
             
              </td>
  </tr>


</table>
</form>
