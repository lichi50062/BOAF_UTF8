<%
//92.12.20 add 取得可查詢的報表名稱
//92.12.20 add 取得WMFileQuery的權限
//94.02.04 add 預設年月為上個月份,若本月為1月份時.則是申報上個年度的12月份
//94.11.15 add '030'-->F01_在台無住所之外國人新台幣存款表 by 2295
//99.09.27 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%
	String YEAR  = Utility.getYear();
   	String MONTH = Utility.getMonth();

	//92.12.20取得可查詢的報表名稱===============================================================================================
	StringBuffer sqlCmd = new StringBuffer();
	List paramList = new ArrayList();
	String muser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");
	sqlCmd.append(" select a.report_no,c.cmuse_name ");
 	sqlCmd.append(" from  WTT04_1D a,  CDShareNO b,  CDShareNO c ");
 	sqlCmd.append(" where a.MUSER_ID=?");
	sqlCmd.append(" and a.PROGRAM_ID='WMFileQuery'");
	sqlCmd.append(" and a.DETAIL_TYPE='4'");//4-->Query
	sqlCmd.append(" and b.CmUSE_Div = '011'");
	sqlCmd.append(" and (a.TRANSFER_TYPE = b.CmUSE_id and b.CmUSE_Div = '011') ");
    sqlCmd.append(" and  c.CmUSE_Div in ('012',  '013', '014','030') ");//94.11.15 add 030:F01_在台無住所之外國人新台幣存款表
    sqlCmd.append(" and b.Identify_no = c.Identify_no ");
	sqlCmd.append(" and c.cmuse_name like a.report_no||'%' order by a.report_no");
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
	//取得WMFileQuery的權限
	Properties permission = ( session.getAttribute("WMFileQuery")==null ) ? new Properties() : (Properties)session.getAttribute("WMFileQuery");
	if(permission == null){
       System.out.println("WMFileQuery_Qry.permission == null");
    }else{
       System.out.println("WMFileQuery_Qry.permission.size ="+permission.size());

    }

%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/WMFileQuery.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<title>申報資料查詢下載</title>
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
                      <td><table width=600 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
						<tr class="sbody">
							<td align='left' bgcolor='#D8EFEE'>基準日</td>
							<td colspan=2 bgcolor='e7e7e7'>
                                <input type='text' name='S_YEAR' value="<%=YEAR%>" size='3' maxlength='3' onblur='CheckYear(this)'>
        						<font color='#000000'>年
        						<select id="hide1" name=S_MONTH>
        						<option></option>
        						<%
        							for (int j = 1; j <= 12; j++) {
        							if (j < 10){%>
        							<option value=0<%=j%><%if(MONTH.equals(String.valueOf(j))){out.print(" selected");}%>>0<%=j%></option>
            						<%}else{%>
            						<option value=<%=j%><%if(MONTH.equals(String.valueOf(j))){out.print(" selected");}%>><%=j%></option>
            						<%}%>
        						<%}%>
        						</select><font color='#000000'>月</font>&nbsp;&nbsp;至
        						<input type='text' name='E_YEAR' value="<%=YEAR%>" size='3' maxlength='3' onblur='CheckYear(this)'>
        						<font color='#000000'>年
        						<select id="hide1" name=E_MONTH>
        						<option></option>
        						<%
        							for (int j = 1; j <= 12; j++) {
        							if (j < 10){%>
        							<option value=0<%=j%><%if(MONTH.equals(String.valueOf(j))){out.print(" selected");}%>>0<%=j%></option>
            						<%}else{%>
            						<option value=<%=j%><%if(MONTH.equals(String.valueOf(j))){out.print(" selected");}%>><%=j%></option>
            						<%}%>
        						<%}%>
        						</select><font color='#000000'>月</font>
        						 <input type=hidden name=S_DATE value=''>
        						 <input type=hidden name=E_DATE value=''>
                            </td>
                        </tr>

                        <tr class="sbody">
						  <td align='left' bgcolor='#D8EFEE'>申報資料</td>
						  <td colspan=2 bgcolor='e7e7e7'><select name=Report_no>
						  <option value='ALL'>全部</option>
							<%
							for (int i = 0; i < Report_List.length; i++) {
							%>
							<option value=<%=Report_List[i][0]%>>
								<%=Report_List[i][0]%>
								&nbsp;&nbsp;&nbsp;
								<%=Report_List[i][1]%>
							</option>
							<%}%>
						    </td>
	  					</tr>


                        </Table></td>
                    </tr>
                    <tr>
                <!--td><div align="right"><jsp:include page="getMaintainUser.jsp" flush="true" /></div></td-->
              </tr>
              <tr>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td><div align="center">
                    <table width="243" border="0" cellpadding="1" cellspacing="1">
                      <tr>
                      <%if(permission != null && permission.get("Q") != null && permission.get("Q").equals("Y")){//Query %>
                        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_queryb.gif',1)"><img src="images/bt_query.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a></div></td>
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
                      <td width="577">
                      <ul>
                      <li>本網頁提供查詢各類申報資料的狀況。</li>
					  <li>選擇欲查詢的檔案,按查詢即可查詢該檔案各月份的申報狀況。</li>
                      </ul>
                      </td>
                    </tr>
                  </table></td>
              </tr>
              <!--tr>
                <td><div align="center"><img src="images/line_1.gif" width="600" height="12"></div></td>
              </tr-->
</table>
</form>
<br><br><br><br><br>
</body>
</html>
