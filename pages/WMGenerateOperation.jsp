<%
//98.08.24 add 線上產生(A01~A04)Axx_operation中間值 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="java.util.*" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Calendar" %>
<%
	String actMsg = ( request.getAttribute("actMsg")==null ) ? "" : (String)request.getAttribute("actMsg");
	Calendar now = Calendar.getInstance();
   	String YEAR  = String.valueOf(now.get(Calendar.YEAR)-1911); //回覆值為西元年故需-1911取得民國年;
   	String MONTH = String.valueOf(now.get(Calendar.MONTH)+1);   //月份以0開始故加1取得實際月份;
    if(MONTH.equals("1")){//若本月為1月份是..則是申報上個年度的12月份
       YEAR = String.valueOf(Integer.parseInt(YEAR) - 1);
       MONTH = "12";
    }else{    
      MONTH = String.valueOf(Integer.parseInt(MONTH) - 1);//申報上個月份的
    }
    //94.05.20 add 取得request的bank_type=====================================================
	String bank_type = ( request.getParameter("bank_type")==null ) ? "" : (String)request.getParameter("bank_type");		
	
	String Report_no = ( request.getParameter("Report_no")==null ) ? "" : (String)request.getParameter("Report_no");					
	String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? YEAR : (String)request.getParameter("S_YEAR");		
	String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? MONTH : (String)request.getParameter("S_MONTH");		
		
	//======================================================================================================================
	//fix 93.12.18 若有已點選的tbank_no,則以已點選的tbank_no為主============================================================
	String bank_code = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");				
	String nowtbank_no =  ( request.getParameter("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");			
	if(!nowtbank_no.equals("")){
	   session.setAttribute("nowtbank_no",nowtbank_no);//將已點選的tbank_no寫入session	   
	}   
	bank_code = ( session.getAttribute("nowtbank_no")==null ) ? bank_code : (String)session.getAttribute("nowtbank_no");			
	//=======================================================================================================================
	//92.12.18取得可線上編輯的報表名稱===============================================================================================
	String sqlCmd = "";
	String muser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");
	List paramList = new ArrayList();
	sqlCmd = " select a.report_no,c.cmuse_name "               
 		   + " from  WTT04_1D a,  CDShareNO b,  CDShareNO c "
 		   + " where a.MUSER_ID=? " 
		   + " and a.PROGRAM_ID='WMFileUpload'"
		   + " and a.DETAIL_TYPE='1'"//1-->Upload
		   + " and b.CmUSE_Div = '011'"
		   + " and (a.TRANSFER_TYPE = b.CmUSE_id and b.CmUSE_Div = '011') "  
           + " and  c.CmUSE_Div in ('012',  '013', '014') "
           + " and b.Identify_no = c.Identify_no "
		   + " and c.cmuse_name like a.report_no||'%' order by a.report_no";
	paramList.add(muser_id);	   
    List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");		   
    String[][] Report_List = new String[dbData.size()][2];
    if(dbData != null && dbData.size() != 0){		   
       for(int i=0;i<dbData.size();i++){
           Report_List[i][0]=(String)((DataObject)dbData.get(i)).getValue("report_no");
	       Report_List[i][1]=(String)((DataObject)dbData.get(i)).getValue("cmuse_name");
	       if(Report_List[i][1].indexOf("_") != -1){
	          Report_List[i][1]=Report_List[i][1].substring(Report_List[i][1].indexOf("_")+1,Report_List[i][1].length());
	       }
       }
	}	   		
	//=======================================================================================================================
	//取得WMGenerateOperation的權限
	Properties permission = ( session.getAttribute("WMGenerateOperation")==null ) ? new Properties() : (Properties)session.getAttribute("WMGenerateOperation"); 
	if(permission == null){
       System.out.println("WMGenerateOperation.permission == null");
    }else{
       System.out.println("WMGenerateOperation.permission.size ="+permission.size());
               
    }		
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<title>線上產生彈性報表動態資料</title>
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
function message() {
<% if(!actMsg.equals("")) {%>
   alert("<%=actMsg%>");
<%}%>
return ;
}
//-->
</script>
</head>

<body  leftmargin="0" topmargin="0" >
<form method=post action='/pages/WMGenerateOperation.jsp' onload='javascript:message();'>
<input type="hidden" name="act" value="Status">
<table width="640" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#297A76">
  <tr>
    <td bordercolor="#FFFFFF"><table width="640" border="0" align="center" cellpadding="0" cellspacing="0">
        <!--tr> 
          <td><img src="images/topbanner_1.gif" width="780" height="103"></td>
        </tr-->
        <tr> 
          <td bgcolor="#FFFFFF">&nbsp;</td>
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
                          線上產生彈性報表動態資料
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
               
                    <tr> 
                      <div align="right"><jsp:include page="getLoginUser.jsp" flush="true" /></div> 
                    </tr>
                    <tr> 
                      <td><table width=600 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
                          <tr class="sbody">
						  <td align='left' bgcolor='#D8EFEE'>申報資料</td>
						  <td colspan=2 bgcolor='e7e7e7'><select name=Report_no>
							<%
							for (int i = 0; i < 4; i++) {//目前只有A01~A04
							%>
							<option value=<%=Report_List[i][0]%>  <%if(Report_no.equals(Report_List[i][0])) out.print("selected");%>>
								<%=Report_List[i][0]%>
								&nbsp;&nbsp;&nbsp;
								<%=Report_List[i][1]%>
							</option>
							<%}%>
						</td>
	  					</tr>

	  					<tr class="sbody">
						<td align='left' bgcolor='#D8EFEE'>基準日</td>
						<td colspan=2 bgcolor='e7e7e7'>
                            <input type='text' name='S_YEAR' value="<%=S_YEAR%>" size='3' maxlength='3' onblur='CheckYear(this)'>
        						<font color='#000000'>年
        						<select id="hide1" name=S_MONTH>
        						<option></option>
        						<%
        							for (int j = 1; j <= 12; j++) {
        							if (j < 10){%>        	
        							<option value=0<%=j%> <%if(S_MONTH.equals(String.valueOf(j))) out.print("selected");%>>0<%=j%></option>        		
            						<%}else{%>
            						<option value=<%=j%> <%if(S_MONTH.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
            						<%}%>
        						<%}%>
        						</select><font color='#000000'>月</font>
                            </td>
                         </tr>
        				
                        </Table>
                      </td>
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
                        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_uploadb.gif',1)"><img src="images/bt_upload.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a></div></td>                      
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
                          <li>本網頁提供線上產生彈性報表各類申報資料。</li>                          
                        </ul></td>
                    </tr>
                  </table></td>
              </tr>
              <!--tr> 
                <td><div align="center"><img src="images/line_1.gif" width="600" height="12"></div></td>
              </tr-->
</table>
</form>

<script language="javascript" >
<!--
//message();
//setSelect(this.document.forms[0].Report_no,"<%=Report_no%>");
//setSelect(this.document.forms[0].S_YEAR,"<%=S_YEAR%>");
//setSelect(this.document.forms[0].S_MONTH,"<%=S_MONTH%>");
-->
</script>
</body>
</html>
