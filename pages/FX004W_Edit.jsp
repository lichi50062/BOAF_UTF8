<%
// 93.12.21 add 權限設定 by 2295
// 94.01.12 fix super user跟自己的機構才可以改機構基本資料,其他人只能查詢 by 2295
//100.01.26 fix cd02區域別.區分100年度/99年度 by 2295
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
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");		
	String list_type = ( request.getParameter("list_type")==null ) ? "" : (String)request.getParameter("list_type");		
	String muser_type = ( session.getAttribute("muser_type")==null ) ? "" : (String)session.getAttribute("muser_type");				    
	String tbank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");				        
		
	//取得FX004W的權限
	Properties permission = ( session.getAttribute("FX004W")==null ) ? new Properties() : (Properties)session.getAttribute("FX004W"); 
	if(permission == null){
       System.out.println("FX004W_Edit.permission == null");
    }else{
       System.out.println("FX004W_Edit.permission.size ="+permission.size());               
    }		
	List BANK_CMML = (List)request.getAttribute("BANK_CMML");
	System.out.println("FX004W_Edit.act="+act);			
	String cd01Table = Integer.parseInt(Utility.getYear())>99 ? "cd01" : "cd01_99" ;
	String cd02Table = Integer.parseInt(Utility.getYear())> 99 ?"cd02" : "cd02_99" ;
	String sqlcmd = "Select CD01.HSIEN_ID,CD01.HSIEN_NAME,CD02.AREA_ID,CD02.AREA_NAME From  "+cd01Table+" CD01, "+cd02Table+" CD02 "+
				 " Where  CD01.HSIEN_ID=CD02.HSIEN_ID Order by CD01.HSIEN_ID, CD02.AREA_ID ";
	List hsien_id_area_id = DBManager.QueryDB_SQLParam(sqlcmd,null,"");
	String HSIEN_ID_AREA_ID = "";
	if(BANK_CMML != null && BANK_CMML.size() != 0){
	     HSIEN_ID_AREA_ID =((((DataObject)BANK_CMML.get(0)).getValue("hsien_id") == null) ?"":(String)((DataObject)BANK_CMML.get(0)).getValue("hsien_id"))	   
						  +"/"+ ((((DataObject)BANK_CMML.get(0)).getValue("area_id") == null) ?"":(String)((DataObject)BANK_CMML.get(0)).getValue("area_id"));	   
	}	
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/FX004W.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<%//if(list_type.equals("1")){%>
<!--title>地方主管機關基本資料維護</title-->
<%//}else if(list_type.equals("2")){%>
<!--title>共用中心基本資料維護</title-->
<%//}else if(list_type.equals("3")){%>
<!--title>農業行庫基本資料維護</title-->
<%//}%>
<title>機構基本資料維護</title>
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
        <tr> 
          <td bgcolor="#FFFFFF">
			<table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="210"><img src="images/banner_bg1.gif" width="210" height="17"></td>
                      <td width="180"><font color='#000000' size=4><b> 
                        <center>
                        <%/*if(list_type.equals("1")){%>							
						<%}else if(list_type.equals("2")){%>							
						<%}else if(list_type.equals("3")){%>							
						<%}*/%>        
						 機構基本資料維護
                        </center>
                        </b></font> </td>
                      <td width="210"><img src="images/banner_bg1.gif" width="210" height="17"></td>
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
                          <tr class="sbody">
						  <td width='30%' align='left' bgcolor='#D8EFEE'>機構代號</td>						 
						  <td width='70%' colspan=2 bgcolor='e7e7e7'>
                            &nbsp;<%if(BANK_CMML != null && BANK_CMML.size() != 0) out.print((String)((DataObject)BANK_CMML.get(0)).getValue("bank_no"));%>
                            <input type='hidden' name='bank_no' value="<%if(BANK_CMML != null && BANK_CMML.size() != 0) out.print((String)((DataObject)BANK_CMML.get(0)).getValue("bank_no"));%>">
                          </td>
                          </tr>  
                            
                          <tr class="sbody">
						  <td width='30%' align='left' bgcolor='#D8EFEE'>機構中文名稱</td>
						  <td width='70%' colspan=2 bgcolor='e7e7e7'>
						    &nbsp;<%if(BANK_CMML != null && BANK_CMML.size() != 0) out.print((String)((DataObject)BANK_CMML.get(0)).getValue("bank_name"));%>						    
                          </td>
	  					  </tr>
                    
                          <tr class="sbody">
						  <td width='30%' align='left' bgcolor='#D8EFEE'>機構英文名稱</td>
						  <td width='70%' colspan=2 bgcolor='e7e7e7'>
                            <input type='text' name='BANK_ENGLISH' value="<%if(BANK_CMML != null && BANK_CMML.size() != 0) out.print( (((DataObject)BANK_CMML.get(0)).getValue("bank_english") == null) ? "":(String)((DataObject)BANK_CMML.get(0)).getValue("bank_english"));%>" size='30' maxlength='60'>                                                        
                          </td>
                          </tr>		  
                           
                          <tr class="sbody">
						  <td width='30%' bgcolor='#D8EFEE' align='left'>實際辦理農漁業金融<br>業務人員</td>
						  <td width='70%' colspan=2 bgcolor='e7e7e7'>
						    <input type='text' name='BUSINESS_PERSON' value="<%if(BANK_CMML != null && BANK_CMML.size() != 0) out.print( ((DataObject)BANK_CMML.get(0)).getValue("business_person") == null ?"0":(((DataObject)BANK_CMML.get(0)).getValue("business_person")).toString()); else out.print("0");%>" size='3' maxlength='3' >
						    (若為總機構請加上 「含分支機構人數」)
						    <font color="red" size=4>*</font>
						  </td>
        			      </tr>
        			      
        			      <tr class="sbody">
						  <td width='30%' bgcolor='#D8EFEE' align='left'>機構地址</td>
						  <td width='70%' colspan=2 bgcolor='e7e7e7'>						
						  <select name='HSIEN_ID_AREA_ID'>
						  <%for(int i=0;i<hsien_id_area_id.size();i++){%>
						  <option value="<%=(String)((DataObject)hsien_id_area_id.get(i)).getValue("hsien_id")%>/<%=(String)((DataObject)hsien_id_area_id.get(i)).getValue("area_id")%>"
						  <%if((HSIEN_ID_AREA_ID).equals(((String)((DataObject)hsien_id_area_id.get(i)).getValue("hsien_id"))+"/"+((String)((DataObject)hsien_id_area_id.get(i)).getValue("area_id")))) out.print("selected");%>
						  ><%=(String)((DataObject)hsien_id_area_id.get(i)).getValue("hsien_name")%>/<%=(String)((DataObject)hsien_id_area_id.get(i)).getValue("area_name")%></option>
						  <%}%>
						  </select>
						  <br>
        				  <input type='text' name='ADDR' value="<%if(BANK_CMML != null && BANK_CMML.size() != 0 && ((DataObject)BANK_CMML.get(0)).getValue("addr") != null ) out.print((String)((DataObject)BANK_CMML.get(0)).getValue("addr"));%>" size='50' maxlength='80' >                            
        			      </td>
        			      </tr>	                       
        			      
					      <tr class="sbody">
						  <td width='30%' bgcolor='#D8EFEE' align='left'>機構網址</td>
						  <td width='70%' colspan=2 bgcolor='e7e7e7'>
						    <input type='text' name='WEB_SITE' value="<%if(BANK_CMML != null && BANK_CMML.size() != 0 && ((DataObject)BANK_CMML.get(0)).getValue("web_site") != null ) out.print((String)((DataObject)BANK_CMML.get(0)).getValue("web_site"));%>" size='50' maxlength='60' >                            
        				   </td>
        				  </tr>	
                    	
                    	  <tr class="sbody">
						  <td colspan=3 bgcolor='#9AD3D0' align='center'>主要連絡人</td>						  
        				  </tr>	
        				  
        				  <tr class="sbody">
						  <td width='30%' bgcolor='#D8EFEE' align='left'>職稱</td>
						  <td width='70%' colspan=2 bgcolor='e7e7e7'>
						    <input type='text' name='M_POSITION' value="<%if(BANK_CMML != null && BANK_CMML.size() != 0 && ((DataObject)BANK_CMML.get(0)).getValue("m_position") != null ) out.print((String)((DataObject)BANK_CMML.get(0)).getValue("m_position"));%>" size='50' maxlength='80' >                            
        				   </td>
        				  </tr>	
        				  
        				  <tr class="sbody">
						  <td width='30%' bgcolor='#D8EFEE' align='left'>姓名</td>
						  <td width='70%' colspan=2 bgcolor='e7e7e7'>
						    <input type='text' name='M_NAME' value="<%if(BANK_CMML != null && BANK_CMML.size() != 0 && ((DataObject)BANK_CMML.get(0)).getValue("m_name") != null ) out.print((String)((DataObject)BANK_CMML.get(0)).getValue("m_name"));%>" size='50' maxlength='80' >                            
						    <font color="red" size=4>*</font>
        				   </td>
        				  </tr>	
        				  
        				  <tr class="sbody">
						  <td width='30%' bgcolor='#D8EFEE' align='left'>電話</td>
						  <td width='70%' colspan=2 bgcolor='e7e7e7'>
						    <input type='text' name='M_TELNO' value="<%if(BANK_CMML != null && BANK_CMML.size() != 0 && ((DataObject)BANK_CMML.get(0)).getValue("m_telno") != null ) out.print((String)((DataObject)BANK_CMML.get(0)).getValue("m_telno"));%>" size='50' maxlength='90' >                            
						    <font color="red" size=4>*</font>
        				   </td>
        				  </tr>	
        				  
        				  <tr class="sbody">
						  <td width='30%' bgcolor='#D8EFEE' align='left'>手機</td>
						  <td width='70%' colspan=2 bgcolor='e7e7e7'>
						    <input type='text' name='M_CELLINO' value="<%if(BANK_CMML != null && BANK_CMML.size() != 0 && ((DataObject)BANK_CMML.get(0)).getValue("m_cellino") != null ) out.print((String)((DataObject)BANK_CMML.get(0)).getValue("m_cellino"));%>" size='50' maxlength='90' >                            
        				   </td>
        				  </tr>	
        				  
        				  <tr class="sbody">
						  <td width='30%' bgcolor='#D8EFEE' align='left'>傳真號碼</td>
						  <td width='70%' colspan=2 bgcolor='e7e7e7'>
						    <input type='text' name='M_FAX' value="<%if(BANK_CMML != null && BANK_CMML.size() != 0 && ((DataObject)BANK_CMML.get(0)).getValue("m_fax") != null ) out.print((String)((DataObject)BANK_CMML.get(0)).getValue("m_fax"));%>" size='50' maxlength='60' >                            
        				   </td>
        				  </tr>	
						 
						  <tr class="sbody">
						  <td width='30%' bgcolor='#D8EFEE' align='left'>電子郵件帳號</td>
						  <td width='70%' colspan=2 bgcolor='e7e7e7'>
						    <input type='text' name='M_EMAIL' value="<%if(BANK_CMML != null && BANK_CMML.size() != 0 && ((DataObject)BANK_CMML.get(0)).getValue("m_email") != null ) out.print((String)((DataObject)BANK_CMML.get(0)).getValue("m_email"));%>" size='50' maxlength='180' >                            
						    <font color="red" size=4>*</font>
        				   </td>
        				  </tr>		      
        				  
        				  <tr class="sbody">
						  <td width='30%' bgcolor='#D8EFEE' align='left'>性別</td>
						  <td width='70%' colspan=2 bgcolor='e7e7e7'>
						    <select name='M_SEX'>
						     <option value='M' <%if((BANK_CMML == null) || ((BANK_CMML != null && BANK_CMML.size() != 0) && ( ((DataObject)BANK_CMML.get(0)).getValue("m_sex") != null &&  ((String)((DataObject)BANK_CMML.get(0)).getValue("m_sex")).equals("M")))) out.print("selected");%>>男</option>
        				     <option value='F' <%if((BANK_CMML != null && BANK_CMML.size() != 0) && ( ((DataObject)BANK_CMML.get(0)).getValue("m_sex") != null &&   ((String)((DataObject)BANK_CMML.get(0)).getValue("m_sex")).equals("F"))) out.print("selected");%>>女</option>
        				    </select>
        				   </td>
        				  </tr>
        				  
        				  <tr class="sbody">
						  <td colspan=3 bgcolor='#9AD3D0' align='center'>主要連絡人的主管</td>						  
        				  </tr>	
        				  
        				  <tr class="sbody">
						  <td width='30%' bgcolor='#D8EFEE' align='left'>職稱</td>
						  <td width='70%' colspan=2 bgcolor='e7e7e7'>
						    <input type='text' name='M_POSITION_OFFICER' value="<%if(BANK_CMML != null && BANK_CMML.size() != 0 && ((DataObject)BANK_CMML.get(0)).getValue("m_position_officer") != null ) out.print((String)((DataObject)BANK_CMML.get(0)).getValue("m_position_officer"));%>" size='20' maxlength='20' >                            
        				   </td>
        				  </tr>	
        				  
        				  <tr class="sbody">
						  <td width='30%' bgcolor='#D8EFEE' align='left'>姓名</td>
						  <td width='70%' colspan=2 bgcolor='e7e7e7'>
						    <input type='text' name='M_NAME_OFFICER' value="<%if(BANK_CMML != null && BANK_CMML.size() != 0 && ((DataObject)BANK_CMML.get(0)).getValue("m_name_officer") != null ) out.print((String)((DataObject)BANK_CMML.get(0)).getValue("m_name_officer"));%>" size='20' maxlength='20' >                            
        				   </td>
        				  </tr>	
        				  
        				  <tr class="sbody">
						  <td width='30%' bgcolor='#D8EFEE' align='left'>電話</td>
						  <td width='70%' colspan=2 bgcolor='e7e7e7'>
						    <input type='text' name='M_TELNO_OFFICER' value="<%if(BANK_CMML != null && BANK_CMML.size() != 0 && ((DataObject)BANK_CMML.get(0)).getValue("m_telno_officer") != null ) out.print((String)((DataObject)BANK_CMML.get(0)).getValue("m_telno_officer"));%>" size='30' maxlength='30' >                            
        				   </td>
        				  </tr>	
        				  
        				  <tr class="sbody">
						  <td width='30%' bgcolor='#D8EFEE' align='left'>手機</td>
						  <td width='70%' colspan=2 bgcolor='e7e7e7'>
						    <input type='text' name='M_CELLINO_OFFICER' value="<%if(BANK_CMML != null && BANK_CMML.size() != 0 && ((DataObject)BANK_CMML.get(0)).getValue("m_cellino_officer") != null ) out.print((String)((DataObject)BANK_CMML.get(0)).getValue("m_cellino_officer"));%>" size='30' maxlength='30' >                            
        				   </td>
        				  </tr>	
        				  
        				  <tr class="sbody">
						  <td width='30%' bgcolor='#D8EFEE' align='left'>傳真號碼</td>
						  <td width='70%' colspan=2 bgcolor='e7e7e7'>
						    <input type='text' name='M_FAX_OFFICER' value="<%if(BANK_CMML != null && BANK_CMML.size() != 0 && ((DataObject)BANK_CMML.get(0)).getValue("m_fax_officer") != null ) out.print((String)((DataObject)BANK_CMML.get(0)).getValue("m_fax_officer"));%>" size='20' maxlength='20' >                            
        				   </td>
        				  </tr>	
        				  
        				  <tr class="sbody">
						  <td width='30%' bgcolor='#D8EFEE' align='left'>電子郵件帳號</td>
						  <td width='70%' colspan=2 bgcolor='e7e7e7'>
						    <input type='text' name='M_EMAIL_OFFICER' value="<%if(BANK_CMML != null && BANK_CMML.size() != 0 && ((DataObject)BANK_CMML.get(0)).getValue("m_email_officer") != null ) out.print((String)((DataObject)BANK_CMML.get(0)).getValue("m_email_officer"));%>" size='50' maxlength='60' >                            
        				   </td>
        				  </tr>	
                    	   
                    	  <tr class="sbody">
						  <td width='30%' bgcolor='#D8EFEE' align='left'>性別</td>
						  <td width='70%' colspan=2 bgcolor='e7e7e7'>
						    <select name='M_SEX_OFFICER'>
						     <option value='M' <%if((BANK_CMML == null) || ((BANK_CMML != null && BANK_CMML.size() != 0) && ( ((DataObject)BANK_CMML.get(0)).getValue("m_sex_officer") != null &&  ((String)((DataObject)BANK_CMML.get(0)).getValue("m_sex_officer")).equals("M")))) out.print("selected");%>>男</option>
        				     <option value='F' <%if((BANK_CMML != null && BANK_CMML.size() != 0) && ( ((DataObject)BANK_CMML.get(0)).getValue("m_sex_officer") != null &&   ((String)((DataObject)BANK_CMML.get(0)).getValue("m_sex_officer")).equals("F"))) out.print("selected");%>>女</option>
        				    </select>
        				   </td>
        				  </tr> 
                    	   
        			    </Table></td>
                    </tr>                 
                    <tr>                  
                <td><div align="right"><jsp:include page="getMaintainUser.jsp" flush="true" /></div></td>                                              
              </tr>              
              <tr> 
                <td>&nbsp;</td>
              </tr>
              <tr> 
                <td><div align="center"> 
                    <table width="243" border="0" cellpadding="1" cellspacing="1">
                      <tr>     
                        <%
                        System.out.println("tbank_no="+tbank_no);
                        System.out.println("tbank_no_now="+(String)((DataObject)BANK_CMML.get(0)).getValue("bank_no"));
                        System.out.println("muser_type="+muser_type);                    
                        %>
                        <%if(permission != null && permission.get("U") != null && permission.get("U").equals("Y")){//Update %>  
                        <%
                        //94.01.12 fix super user跟自己的機構才可以改====
                        if(muser_type.equals("S") || ((String)((DataObject)BANK_CMML.get(0)).getValue("bank_no")).equals(tbank_no)){ 
                        %>                 	        	                                   		     
				        <td width="66"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Update');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_updateb.gif',1)"><img src="images/bt_update.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a></div></td>
				        <td width="66"><div align="center"><a href="javascript:AskReset(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_cancelb.gif',1)"><img src="images/bt_cancel.gif" name="Image105" width="66" height="25" border="0" id="Image105"></a></div></td>                        
				        <%}%>
				        <%}%>
                        <td width="93"><div align="center"><a href="javascript:history.back();"  onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image106','','images/bt_backb.gif',1)"><img src="images/bt_back.gif" name="Image106" width="80" height="25" border="0" id="Image106"></a></div></td>
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
                          <li>確認輸入資料無誤後, 按<font color="#666666">【修改】</font>即將本表上的資料, 於資料庫中建檔。</li>                          
                          <li>欲重新輸入資料, 按<font color="#666666">【取消】</font>即將本表上的資料清空。</li>                          
                          <li>如放棄, 按<font color="#666666">【回上一頁】</font>即離開本程式。</li>                         
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
