<%
//92.12.18 add 取得可上傳檔案的報表名稱 by 2295
//94.11.04 add '030'-->F01_在台無住所之外國人新台幣存款表 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="java.util.*" %>
<%@ page import="java.util.LinkedList" %>

<%
    String bank_type = ( request.getParameter("bank_type")==null ) ? "" : (String)request.getParameter("bank_type");		
	
	System.out.println("bank_type="+bank_type);	
	
	//92.12.18取得可上傳檔案的報表名稱===============================================================================================
	String sqlCmd = "";
	List paramList = new ArrayList();
	String muser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		
	sqlCmd = " select a.report_no,c.cmuse_name "               
 		   + " from  WTT04_1D a,  CDShareNO b,  CDShareNO c "
 		   + " where a.MUSER_ID=?" 
		   + " and a.PROGRAM_ID='WMFileEdit'"
		   + " and a.DETAIL_TYPE='3'"//3-->Edit
		   + " and b.CmUSE_Div = '011'"
		   + " and (a.TRANSFER_TYPE = b.CmUSE_id and b.CmUSE_Div = '011') "  
           + " and  c.CmUSE_Div in ('012',  '013', '014','030') "//94.11.04 add 030:F01_在台無住所之外國人新台幣存款表
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
	
	//String bank_type = ( request.getParameter("bank_type")==null ) ? "1" : (String)request.getParameter("bank_type");			
	
	//String[][] DLId_List = ListArray.getUpListArray(bank_type);//可申報的檔案類型
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/WMFileEdit.js"></script>
<script language="javascript" event="onresize" for="window"></script>

<html>
<head>
<title>線上編輯申報資料</title>
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
<input type="hidden" name="act" value="Status">  
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
                          線上編輯申報資料 
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
                      <td><table width=600 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
                          <tr align=left bgcolor='#D2F0FF'>                           
                          <td align='center' width=40% class="sbody">申報檔案代號</td>
						  <td align='center' width=60% class="sbody">申報檔案名稱</td>
                          </tr>
                        </Table></td>
                    </tr>
                  
                    <tr> 
                      <td><table width='600' border=1 align='center' cellpadding="1" cellspacing="1" bordercolor="#3A9D99">                   
                          <%
                           String bgcolor="#FFFFE6";
                           for(int i = 0; i < Report_List.length; i++){
                            bgcolor = (i % 2 == 0)?"#FFFFE6":"#F2F2F2";
                           %>
						   <tr bgcolor='<%=bgcolor%>' class="sbody"> 
							<td align='center' width=40% bordercolor="#3A9D99"><a href="javascript:doSubmitStatus(this.document.forms[0],'<%=Report_List[i][0]%>','<%=Report_List[i][1]%>','<%=bank_type%>');"><%=Report_List[i][0]%></a></td>	   
							<td align='center' width=60% bordercolor="#3A9D99"><a href="javascript:doSubmitStatus(this.document.forms[0],'<%=Report_List[i][0]%>','<%=Report_List[i][1]%>','<%=bank_type%>');"><%=Report_List[i][1]%></a>   <% if((Report_List[i][0]).equals("A11")) out.print("<font color='red'>(本申報功能已修改版本，請詳閱使用說明)</font>");%></td>	   	
    						</tr>    
    					  <%}%>	                                       
            			</table>
            		 </td>
        </tr>        
      </table></td>
  </tr>
</table>
</form>
</body>
</html>
