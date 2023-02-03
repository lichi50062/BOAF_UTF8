<%
//93.12.17 add 權限檢核 by 2295
//93.12.18 add 若有已點選的tbank_no,則以已點選的tbank_no為主 by 2295
//92.12.18 add 取得可線上編輯的報表名稱 by 2295
//94.02.04 add 預設年月為上個月份,若本月為1月份時.則是申報上個年度的12月份
//94.05.20 fix 取得request的bank_type by 2295
//99.09.24 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
//101.11.27 fix 調整申報項目排列順序 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%
	String actMsg = ( request.getAttribute("actMsg")==null ) ? "" : (String)request.getAttribute("actMsg");
	String YEAR  = Utility.getYear();
   	String MONTH = Utility.getMonth();

    //94.05.20 add 取得request的bank_type=====================================================
	String bank_type = ( request.getParameter("bank_type")==null ) ? "" : (String)request.getParameter("bank_type");

	//String[][] DLId_List = ListArray.getUpListArray(bank_type);//可申報的檔案類型
	String Report_no = ( request.getParameter("Report_no")==null ) ? "" : (String)request.getParameter("Report_no");
	String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? YEAR : (String)request.getParameter("S_YEAR");
	String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? MONTH : (String)request.getParameter("S_MONTH");
	//String bank_code = ( session.getAttribute("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");
	//bank_code = ( request.getParameter("bank_code")==null ) ? bank_code : (String)request.getParameter("bank_code");
	//request.setAttribute("tbank_no",bank_code);

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
	StringBuffer sqlCmd = new StringBuffer();
	List paramList = new ArrayList();
	String muser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");
	sqlCmd.append(" select a.report_no,c.cmuse_name ");
 	sqlCmd.append(" from  WTT04_1D a,  CDShareNO b,  CDShareNO c ");
 	sqlCmd.append(" where a.MUSER_ID=?");
	sqlCmd.append(" and a.PROGRAM_ID='WMFileUpload'");
	sqlCmd.append(" and a.DETAIL_TYPE='1'");//1-->Upload
	sqlCmd.append(" and b.CmUSE_Div = '011'");
	sqlCmd.append(" and (a.TRANSFER_TYPE = b.CmUSE_id and b.CmUSE_Div = '011') ");
    sqlCmd.append(" and  c.CmUSE_Div in ('012',  '013', '014') ");
    sqlCmd.append(" and b.Identify_no = c.Identify_no ");
	//sqlCmd.append(" and c.cmuse_name like a.report_no||'%' order by a.report_no");
	sqlCmd.append(" and c.cmuse_name like a.report_no||'%' order by  c.CmUSE_Div asc, to_number(c.input_order) asc");
	 
    paramList.add(muser_id);
    List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
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
	//取得WMFileUpload的權限
	Properties permission = ( session.getAttribute("WMFileUpload")==null ) ? new Properties() : (Properties)session.getAttribute("WMFileUpload");
	if(permission == null){
       System.out.println("WMFileUpload_Edit.permission == null");
    }else{
       System.out.println("WMFileUpload_Edit.permission.size ="+permission.size());

    }
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/WMFileUpload.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<title>申報資料檔案上傳</title>
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
   alert("11111<%=actMsg%>");
<%}%>
return ;
}
//-->
</script>
</head>

<body  leftmargin="0" topmargin="0" >
<form method=post ENCTYPE="multipart/form-data" action='/pages/WMFileUpload.jsp' onload='javascript:message();'>
<input type="hidden" name="act" value="Status">
<input type="hidden" name="bank_code" value="<%=bank_code%>">
<input type="hidden" name="bank_type" value="<%=bank_type%>">
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
                      <td width="200"><img src="images/banner_bg1.gif" width="200" height="17"></td>
                      <td width="244"><font color='#000000' size=4><b>
                        <center>
                          申報資料檔案上傳
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
                          <tr class="sbody">
						  <td align='left' bgcolor='#D8EFEE'>申報資料</td>
						  <td colspan=2 bgcolor='e7e7e7'><select name=Report_no>
							<%
							for (int i = 0; i < Report_List.length; i++) {
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
        				<tr class="sbody">
						<td bgcolor='#D8EFEE' align='left'>上傳檔案位置</td>
						<td colspan=2 bgcolor='e7e7e7'>
        				<input type=file size=40 name=UpFileName >
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
                      <%if(permission != null && permission.get("up") != null && permission.get("up").equals("Y")){//Query %>
				         <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_uploadb.gif',1)"><img src="images/bt_upload.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a></div></td>
                         <td width="66"> <div align="center"><a href="javascript:AskReset(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image102','','images/bt_cancelb.gif',1)"><img src="images/bt_cancel.gif" name="Image102" width="66" height="25" border="0" id="Image102"></a></div></td>
                      <%}%>
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
                          <li>本網頁提供上傳各類申報資料。</li>
                          <li>選擇欲上傳的檔案，輸入年月，申報資料。</li>
                          <li>按上傳即可上傳該文字檔案。</li>
                          <li>按<font color="#666666">【取消】</font>即重新輸入資料。</li>
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
