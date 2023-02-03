<%
// 94.10.21 create by 2495
//100.01.13 fix 畫面格式 by 2295
//102.01.08 fix 無法顯示公告附加檔案資料 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Enumeration" %>
<%@ page import="java.util.HashMap" %>
<%@ page import="java.util.Hashtable" %>
<%@ page import="java.util.StringTokenizer" %>



<%
	List WLX_Notify = DBManager.QueryDB_SQLParam("select seq_no, headmark, append_file, user_id, user_name, update_date,appfile_link,notify_url,to_char(Notify_DATE,'yyyy/mm/dd') as Notify_DATE ,to_char(Notify_End_DATE,'yyyy/mm/dd') as Notify_End_DATE from WLX_Notify order by notify_date desc",null,"seq_no,headmark,notify_date,append_file,user_id,user_name,update_date,notify_end_date "); 				   
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/ZZ091W.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<title>總機構高階主管基本資料維護</title>
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
<style type="text/css">
<!--
.style1 {color: #666666}
-->
</style>
</head>

<body marginwidth="0" marginheight="0" leftmargin="0" topmargin="0" leftmargin="0"><body onLoad="MM_preloadImages('images/bt_confirmb.gif','images/bt_cancelb.gif','images/bt_backb.gif')">
<form method=post action='#'>
<table width="640" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  		  <tr> 
   		   <td><img src="images/space_1.gif" width="12" height="12"></td>
  		  </tr>
          <td bgcolor="#FFFFFF">
		  <table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
                   <tr> 
                      <td width="170"><img src="images/banner_bg1.gif" width="170" height="17"></td>
                      <td width="250"><font color='#000000' size=4><b> 
                        <center>首頁公告區標題清單</center>
                        </b></font> </td>
                      <td width="170"><img src="images/banner_bg1.gif" width="170" height="17"></td>
                    </tr>                   
                  </table></td>
              </tr>
              <tr> 
                <td><img src="images/space_1.gif" width="12" height="12"></td>
              </tr>
              <tr> 
                <td><table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
               
                    <tr> 
                      
                      <td><table width=600 border=1 align='center' cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
                        <tr bgcolor="#9AD3D0" class="sbody">
                          <td width=40 bgcolor=#9AD3D0><div align="center">項次</div></td>                          
                          <td width=220><div align="center"><font face=細明體 color=#000000>標題</font></div></td>
                          <td width=180><div align="center"><font color="#000000" face="細明體">公告期間</font></div></td>
                          <td width=100><div align="center"><font color="#000000" face="細明體">有附加檔案否</font></div></td>
                          <td width=100><div align="center"><font color="#000000" face="細明體">公告單位人員</font></div></td>
                        </tr>
										
  										<% if(WLX_Notify.size() == 0){%>                 			   
                   			   <td colspan=10 align=center>無資料可供查詢</td><tr>
                   			   <tr>                   			   
                   	  <% }
                    		   int i=0;
                    		   String list1Color="list1Color_sbody";
                      		   String list2Color="list2Color_sbody";
                      	       String listColor="list1Color_sbody";
                    		   while(i < WLX_Notify.size()){  
                    		    listColor = (i % 2 == 0)?list2Color:list1Color;
				   				String seq_str = ((DataObject)WLX_Notify.get(i)).getValue("seq_no").toString();
				  			    System.out.println("seq_str="+seq_str);                  		                      		                          		      
                      %>                         	  
                        <tr class="sbody">
                          <td class="<%=listColor%>" width="40"><%=i+1%><input type="checkbox" name="isDelete" value="<%=((DataObject)WLX_Notify.get(i)).getValue("seq_no")%>"></td>    				                      				                                                   
                          <td class="<%=listColor%>" width=170><div align="center"><font face=細明體 color=#000000><a href="javascript:doSubmit(this.document.forms[0],'Edit','<%=(String)((DataObject)WLX_Notify.get(i)).getValue("seq_no").toString()%>');"><%if( ((DataObject)WLX_Notify.get(i)).getValue("headmark").toString() != null ) out.print(((DataObject)WLX_Notify.get(i)).getValue("headmark").toString()); else out.print("&nbsp;");%></a></font></div></td>              
                          <td class="<%=listColor%>" width=220><div align="center"><font color="#000000" face="細明體"><%if( ((DataObject)WLX_Notify.get(i)).getValue("notify_date") != null ) out.print((String)((DataObject)WLX_Notify.get(i)).getValue("notify_date")); else out.print("&nbsp;");%> ~ <%if( ((DataObject)WLX_Notify.get(i)).getValue("notify_end_date") != null ) out.print((String)((DataObject)WLX_Notify.get(i)).getValue("notify_end_date")); else out.print("&nbsp;");%></font></div></td>
                          <td class="<%=listColor%>" width=100><div align="center"><font face=細明體 color=#000000><%if( ((DataObject)WLX_Notify.get(i)).getValue("append_file") == null || ((DataObject)WLX_Notify.get(i)).getValue("append_file").equals(" ")) out.print("N"); else out.print("Y");%></font></div></td>
                          <td class="<%=listColor%>" width=88><div align="center"><font color="#000000" face="細明體"><%if( ((DataObject)WLX_Notify.get(i)).getValue("user_name") != null ) out.print((String)((DataObject)WLX_Notify.get(i)).getValue("user_name")); else out.print("&nbsp;");%></font></div></td>
                        </tr> 					      
					      			<%
                  			    i++;
	                  		   }//end of while                     
                      %>

                      </table></td>
                    </tr>
                    <tr> 
                      <td>&nbsp;</td>
                    </tr>                 
                    <tr>                  
                <td><div align="right"></div></td>                                              
              </tr>
              
              <tr> 
                <td>&nbsp;</td>
              </tr>
              <tr> 
                <td><div align="center"> 
                    <table width="243" border="0" cellpadding="1" cellspacing="1">
                      <tr>   
                        <td width="66"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'new');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image107','','images/bt_addb.gif',1)"><img src="images/bt_add.gif" name="Image107" width="66" height="25" border="0" id="Image107"></a>
 												<td width="66"><div align="center"><a href="javascript:AskReset(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_cancelb.gif',1)"></a></div></td>
 												<td width="66"><div align="center"><a href="javascript:AskReset(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_cancelb.gif',1)"></a></div></td>
 												<td width="66"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'del');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image108','','images/bt_deleteb.gif',1)"><img src="images/bt_delete.gif" name="Image108" width="66" height="25" border="0" id="Image108"></a>                        						  
				        				                                
                        
                      </tr>
                    </table>
                  </div></td>
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
                      	  <li>點選標題內容即可讀取資料庫內現存公告標題內容。</li>     
                      	  <li>若要新增一筆公告標題請按<font color="#666666">【</font><span class="style1">新增</span><font color="#666666">】</font>以開始登打公告。</li>
                      	  <li>若要修改一筆公告標題請點選公告標題以後即可進入修改公告。</li>
                      	  <li>若要刪除公告標題請勾選小空格後點選<font color="#666666">【</font><span class="style1">刪除</span><font color="#666666">】</font>。</li>
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
