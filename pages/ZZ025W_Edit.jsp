<%
//94.01.06 fix 所有人都有使用者資料維護的權限 by 2295
//94.02.03 fix 區分網際網路申報/MIS管理系統的配色 by 2295
//94.04.07 fix 姓名改12 byte by 2295
//99.12.10 fix sqlInjection by 2808
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%@ page import="java.util.ArrayList" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%
	String muser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		
	String muser_name = ( session.getAttribute("muser_name")==null ) ? "" : (String)session.getAttribute("muser_name");		
	String muser_type = ( session.getAttribute("muser_type")==null ) ? "" : (String)session.getAttribute("muser_type");			
	String bank_type = ( session.getAttribute("bank_type")==null ) ? "" : (String)session.getAttribute("bank_type");		
	String bank_type_name = "";
	String tbank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");			
	String bankdata = "";
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");				
	//取得ZZ025的權限
	/*所有人都有使用者資料維護的權限
	Properties permission = ( session.getAttribute("ZZ025W")==null ) ? new Properties() : (Properties)session.getAttribute("ZZ025W"); 
	if(permission == null){
       System.out.println("ZZ025W_Edit.permission == null");
    }else{
       System.out.println("ZZ025W_Edit.permission.size ="+permission.size());
               
    }				 
    */
	List MUSER_DATA = (List)request.getAttribute("MUSER_DATA");		
	List WTT01 = (List)request.getAttribute("WTT01");		
	
	String title="基本資料維護」 ";		  
	title =(muser_type.equals(" "))?"「使用者"+title:title;
	title =(muser_type.equals("A"))?"「系統管理者"+title:title;	
	title =(act.equals("Edit"))?title+"異動維護建檔":title;
	title =(act.equals("new"))?title+"新增建檔":title;
	
	String sqlCmd = "";
	String yy = Integer.parseInt(Utility.getYear()) > 99 ?"100" :"99" ;
	List paramList = new ArrayList() ;
	//取得總機構代碼及名稱
	sqlCmd=" select bn01.bank_no || bn01.bank_name as bankdata from wtt01,(select * from bn01 where m_year=?)bn01"
		  +" where wtt01.muser_id = ? and wtt01.tbank_no = bn01.bank_no ";
	paramList.add(yy) ;
	paramList.add(muser_id) ;
	List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");	
	paramList.clear() ;
	if((dbData != null) && (dbData.size() !=0)){
	   bankdata = (String)((DataObject)dbData.get(0)).getValue("bankdata");
	}
    //取得機構類別的中文名稱
    paramList.add(bank_type) ;
    dbData = DBManager.QueryDB_SQLParam("select cmuse_id,cmuse_name from cdshareno where cmuse_div='001' and cmuse_id = ?",paramList,"");
    paramList.clear() ;
	if((dbData != null) && (dbData.size() !=0)){
	   bank_type_name = (String)((DataObject)dbData.get(0)).getValue("cmuse_name");
	}
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/ZZ025W.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<title><%=title%></title>
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
<form method=post action='#'>
<input type="hidden" name="act" value="">  
<table width="640" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  		<tr> 
   		 <td><img src="images/space_1.gif" width="12" height="12"></td>
  		</tr>
          <td bgcolor="#FFFFFF">
			<table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="110"><img src="images/banner_bg1.gif" width="110" height="17"></td>
                      <td width="380"><font color='#000000' size=4><b> 
                        <center>
                          <%=title%> 
                        </center>
                        </b></font> </td>
                      <td width="110"><img src="images/banner_bg1.gif" width="110" height="17"></td>
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
                    <%
                      String nameColor="nameColor_sbody";
                      String textColor="textColor_sbody";
                      String bordercolor="#3A9D99";                      
                    %>          
                    <tr> 
                      <td><table width=600 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="<%=bordercolor%>">
                          <tr bgcolor="#D2F0FF"><td colspan=4 class="sbody" align=center><b>個人基本資料</b></td></tr>
                          
                          <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">機構類別</td>						  
						  <td width='85%' colspan=3 class="<%=textColor%>">
						    <input type='text' name='BANK_TYPE' value="<%=bank_type%>&nbsp;&nbsp;<%=bank_type_name%>" size=50 readonly>                            
                          </td>
                          </tr>     
					      
					      <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">總機構代碼</td>						 
						  <td width='85%' colspan=3 class="<%=textColor%>">
						    <input type='text' name='BANKDATA' value="<%=bankdata%>" size=50 readonly>                                                       
                          </td>
                          </tr>   
                          
                          <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">帳號</td>						 
						  <td width='85%' colspan=3 class="<%=textColor%>">
						    <input type='text' name='MUSER_ID' value="<%=muser_id%>" size=50 readonly>                                                       
                          </td>
                          </tr>  
                          
                          <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">姓名</td>						 
						  <td width='85%' colspan=3 class="<%=textColor%>">
						    <input type='text' name='MUSER_NAME' value="<%=muser_name%>" size='12' maxlength='12' >                                                       
						    <font color='red' size=4>*</font>
                          </td>
                          </tr>  
                          
                          <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">組室代碼</td>
						  <% List bank_no = (List)request.getAttribute("bank_no");%>
						  <td width='85%' colspan=3 class="<%=textColor%>">
						    <input type="hidden" name="OLD_BANK_NO" value="<%if(WTT01 != null && WTT01.size() != 0 && ((DataObject)WTT01.get(0)).getValue("bank_no") != null ) out.print((String)((DataObject)WTT01.get(0)).getValue("bank_no"));%>">
                            <select name='BANK_NO'>  
                            <%if(bank_type.equals("2")){ //農金局局內專用                                                    
                                for(int i=0;i<bank_no.size();i++){%>
                                <option value="<%=(String)((DataObject)bank_no.get(i)).getValue("bank_no")%>"
                                <%if((WTT01 != null && WTT01.size() != 0) && ( ((DataObject)WTT01.get(0)).getValue("bank_no") != null && ((String)((DataObject)WTT01.get(0)).getValue("bank_no")).equals((String)((DataObject)bank_no.get(i)).getValue("bank_no")))) out.print("selected");%>
                                ><%=(String)((DataObject)bank_no.get(i)).getValue("bank_name")%></option>                            
                            <%  }//end of for
                              }else{//非農金局局內 
                            %>  
                              <option value="">&nbsp;&nbsp;</option> 
                            <%}//end of if%>                            
                            </select>
                             (農金局局內專用)
                          </td>
                          </tr>					      
					      
					      <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">科別</td>
						  <% List subdep_id = DBManager.QueryDB_SQLParam("select cmuse_id,cmuse_name from cdshareno where cmuse_div='010' order by input_order",null,"");%>
						  <td width='85%' colspan=3 class="<%=textColor%>">
						    <input type="hidden" name="OLD_SUBDEP_ID" value="<%if(WTT01 != null && WTT01.size() != 0 && ((DataObject)WTT01.get(0)).getValue("subdep_id") != null ) out.print((String)((DataObject)WTT01.get(0)).getValue("subdep_id"));%>">
                            <select name='SUBDEP_ID'>                            
                            <%if(bank_type.equals("2")){ //農金局局內專用                                                                                
                                for(int i=0;i<subdep_id.size();i++){%>
                                <option value="<%=(String)((DataObject)subdep_id.get(i)).getValue("cmuse_id")%>"
                                <%if((WTT01 != null && WTT01.size() != 0) && ( ((DataObject)WTT01.get(0)).getValue("subdep_id") != null && ((String)((DataObject)WTT01.get(0)).getValue("subdep_id")).equals((String)((DataObject)subdep_id.get(i)).getValue("cmuse_id")))) out.print("selected");%>
                                ><%=(String)((DataObject)subdep_id.get(i)).getValue("cmuse_name")%></option>                            
                            <%  }//end of for
                              }else{//非農金局局內
                            %>
                                <option value="">&nbsp;&nbsp;</option> 
                            <%}//end of if%>                                
                            </select>
                             (農金局局內專用)
                          </td>
                          </tr>     					      					     
					      
					      <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">職稱</td>
						  <td width='85%' colspan=3 class="<%=textColor%>">
                            <input type='text' name='M_POSITION' value="<%if(MUSER_DATA != null && MUSER_DATA.size() != 0 && ((DataObject)MUSER_DATA.get(0)).getValue("m_position") != null ) out.print((String)((DataObject)MUSER_DATA.get(0)).getValue("m_position"));%>" size='20' maxlength='20' >                          
                          </td>
                          </tr>                         
                          
                          <tr class="sbody">
                          <td width='15%' class="<%=nameColor%>">電話</td>
						  <td width='35%' align='left' class="<%=textColor%>">
                            <input type='text' name='M_TELNO' value="<%if(MUSER_DATA != null && MUSER_DATA.size() != 0 && ((DataObject)MUSER_DATA.get(0)).getValue("m_telno") != null ) out.print((String)((DataObject)MUSER_DATA.get(0)).getValue("m_telno"));%>" size='20' maxlength='30' >                          
                          </td>
						  <td width='15%' class="nameColor_sbody_right">行動電話</td>
						  <td width='35%' class="<%=textColor%>">
                            <input type='text' name='M_CELLINO' value="<%if(MUSER_DATA != null && MUSER_DATA.size() != 0 && ((DataObject)MUSER_DATA.get(0)).getValue("m_cellino") != null ) out.print((String)((DataObject)MUSER_DATA.get(0)).getValue("m_cellino"));%>" size='20' maxlength='30' >                          
                          </td>
                          </tr>
                          
                          <tr class="sbody">
                          <td width='15%' class="<%=nameColor%>">傳真號碼</td>
						  <td width='35%' class="<%=textColor%>">
                            <input type='text' name='M_FAX' value="<%if(MUSER_DATA != null && MUSER_DATA.size() != 0 && ((DataObject)MUSER_DATA.get(0)).getValue("m_fax") != null ) out.print((String)((DataObject)MUSER_DATA.get(0)).getValue("m_fax"));%>" size='20' maxlength='20' >                          
                          </td>
                          <td width='15%' class="nameColor_sbody_right">性別</td>
						  <td width='35%' class="<%=textColor%>">
						   <select name='M_SEX'>
						   <option value='M' <%if((MUSER_DATA == null) || ((MUSER_DATA != null && MUSER_DATA.size() != 0) &&  ( ((DataObject)MUSER_DATA.get(0)).getValue("m_sex") == null || ((String)((DataObject)MUSER_DATA.get(0)).getValue("m_sex")).equals("M") ) )) out.print("selected");%>>男</option>
        				   <option value='F' <%if((MUSER_DATA != null && MUSER_DATA.size() != 0) && ((String)((DataObject)MUSER_DATA.get(0)).getValue("m_sex")).equals("F")) out.print("selected");%>>女</option>						   
        				   </select>
        				  </td>    
                          </tr>
                          
                          <tr class="sbody">
                          <td width='15%' class="<%=nameColor%>">電子郵件帳號</td>
						  <td width='85%' colspan=3 class="<%=textColor%>">
                            <input type='text' name='M_EMAIL' value="<%if(MUSER_DATA != null && MUSER_DATA.size() != 0 && ((DataObject)MUSER_DATA.get(0)).getValue("m_email") != null ) out.print((String)((DataObject)MUSER_DATA.get(0)).getValue("m_email"));%>" size='60' maxlength='60' >                          
                            <font color='red' size=4>*</font>
                          </td>
                          </tr>
                          
                          <tr class="sbody">
                          <td width='15%' class="<%=nameColor%>">其他說明</td>
						  <td width='85%' colspan=3 class="<%=textColor%>">
                            <input type='text' name='M_NOTE' value="<%if(MUSER_DATA != null && MUSER_DATA.size() != 0 && ((DataObject)MUSER_DATA.get(0)).getValue("m_note") != null ) out.print((String)((DataObject)MUSER_DATA.get(0)).getValue("m_note"));%>" size='60' maxlength='60' >                          
                          </td>
                          </tr>
                          
                          <tr bgcolor="#D2F0FF"><td colspan=4 class="sbody" align=center><b>上一層主管基本資料</b></td></tr>
                          
                          <tr class="sbody">
						  <td width='15%' class="<%=nameColor%>">姓名</td>
						  <td width='85%' colspan=3 class="<%=textColor%>">
                            <input type='text' name='DIRECTOR_NAME' value="<%if(MUSER_DATA != null && MUSER_DATA.size() != 0 && ((DataObject)MUSER_DATA.get(0)).getValue("director_name") != null ) out.print((String)((DataObject)MUSER_DATA.get(0)).getValue("director_name"));%>" size='20' maxlength='30' >                          
                          </td>
                          </tr>
                          
                          <tr class="sbody">
                          <td width='15%' class="<%=nameColor%>">電話</td>
						  <td width='35%' class="<%=textColor%>">
                            <input type='text' name='DIRECTOR_TELNO' value="<%if(MUSER_DATA != null && MUSER_DATA.size() != 0 && ((DataObject)MUSER_DATA.get(0)).getValue("director_telno") != null ) out.print((String)((DataObject)MUSER_DATA.get(0)).getValue("director_telno"));%>" size='20' maxlength='30' >                          
                          </td>                     
                          <td width='15%' class="nameColor_sbody_right">行動電話</td>
						  <td width='35%' class="<%=textColor%>">
                            <input type='text' name='DIRECTOR_CELLINO' value="<%if(MUSER_DATA != null && MUSER_DATA.size() != 0 && ((DataObject)MUSER_DATA.get(0)).getValue("director_cellino") != null ) out.print((String)((DataObject)MUSER_DATA.get(0)).getValue("director_cellino"));%>" size='20' maxlength='30' >                          
                          </td>   
                          </tr>     
                          
                          <tr class="sbody">						                                                  
 						  <td width='15%' class="<%=nameColor%>">傳真號碼</td>
						  <td width='35%' class="<%=textColor%>">
                            <input type='text' name='DIRECTOR_FAX' value="<%if(MUSER_DATA != null && MUSER_DATA.size() != 0 && ((DataObject)MUSER_DATA.get(0)).getValue("director_fax") != null ) out.print((String)((DataObject)MUSER_DATA.get(0)).getValue("director_fax"));%>" size='20' maxlength='20' >                          
                          </td>                         
                          <td width='15%' class="nameColor_sbody_right">性別</td>
						  <td width='35%' class="<%=textColor%>">
						   <select name='DIRECTOR_SEX'>
						   <option value='M' <%if((MUSER_DATA == null) || ((MUSER_DATA != null && MUSER_DATA.size() != 0) && ((String)((DataObject)MUSER_DATA.get(0)).getValue("director_sex")).equals("M"))) out.print("selected");%>>男</option>
        				   <option value='F' <%if((MUSER_DATA != null && MUSER_DATA.size() != 0) && ((String)((DataObject)MUSER_DATA.get(0)).getValue("director_sex")).equals("F")) out.print("selected");%>>女</option>						   
        				   </select>
        				  </td>        				                            
                          </tr>
                          
                          <tr class="sbody">
                          <td width='15%' class="<%=nameColor%>">電子郵件帳號</td>
						  <td width='85%' colspan=3 class="<%=textColor%>">
                            <input type='text' name='DIRECTOR_EMAIL' value="<%if(MUSER_DATA != null && MUSER_DATA.size() != 0 && ((DataObject)MUSER_DATA.get(0)).getValue("director_email") != null ) out.print((String)((DataObject)MUSER_DATA.get(0)).getValue("director_email"));%>" size='60' maxlength='60' >                          
                          </td>
                          </tr>
                          
                        </Table></td>
                    </tr>                 
                    <tr>                  
                <td><div align="right"><jsp:include page="getMaintainUser.jsp" flush="true" /></div></td>                                              
              </tr>
              
              <tr> 
                <td><div align="center"> 
                    <table width="243" border="0" cellpadding="1" cellspacing="1">
                      <tr>              
                      <%//if(permission != null && permission.get("U") != null && permission.get("U").equals("Y")){ %>                   	        	                                   		     
				        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Update');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_confirmb.gif',1)"><img src="images/bt_confirm.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a></div></td>
         				<td width="66"> <div align="center"><a href="javascript:AskReset(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image104','','images/bt_cancelb.gif',1)"><img src="images/bt_cancel.gif" name="Image104" width="66" height="25" border="0" id="Image104"></a></div></td>		                
                        <td width="80"><div align="center"><a href="javascript:history.back();"  onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_backb.gif',1)"><img src="images/bt_back.gif" name="Image105" width="80" height="25" border="0" id="Image105"></a></div></td>
                      <%//}%>  
                      </tr>
                      </tr>
                    </table>
                  </div></td>
              </tr>                               
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
                          <li>本網頁提供基本資料維護功能。</li> 
					      <li>按<font color="#666666">【確定】</font>即將資料寫入資料庫。</li> 
					      <li>按<font color="#666666">【取消】</font>放棄資料修改。</li>
                          <li>按<font color="#666666">【回上一頁】則放棄基本資料維護, 回至上個畫面</font>。</li>
                          <li>【<font color="red" size=4>*</font>】為必填欄位。</li> 
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
