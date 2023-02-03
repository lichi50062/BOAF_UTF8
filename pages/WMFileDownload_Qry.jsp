<%
//93.12.17 add 權限檢核 by 2295
//92.12.18 add 取得可下載檔案的報表名稱 by 2295
//94.02.04 add 預設年月為上個月份,若本月為1月份時.則是申報上個年度的12月份
//94.05.25 fix 取得request的bank_type by 2295
//94.11.15 add '030'-->F01_在台無住所之外國人新台幣存款表 by 2295
//             若為F01時,只開放查詢 by 2295
//94.09.29 add 下載報表 by 2495
//94.12.13 拿掉 hiddenYM(form) by 2495
//96.04.16 add A99_extra中央存保格式A02.A01.A99檔案下載 by 2295
//97.01.10 add 地方主管機關.不能查詢.下載資料 by 2295
//99.09.27 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>

<%
	String YEAR  = Utility.getYear();
   	String MONTH = Utility.getMonth();
    //94.05.25 add 取得request的bank_type=====================================================
	String bank_type = ( request.getParameter("bank_type")==null ) ? "" : (String)request.getParameter("bank_type");
	List dbData = (List)request.getAttribute("dbData");
	//fix 93.12.18 若有已點選的tbank_no,則以已點選的tbank_no為主============================================================
	String bank_code = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");
	String nowtbank_no =  ( request.getParameter("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");
	if(!nowtbank_no.equals("")){
	   session.setAttribute("nowtbank_no",nowtbank_no);//將已點選的tbank_no寫入session
	}
	bank_code = ( session.getAttribute("nowtbank_no")==null ) ? bank_code : (String)session.getAttribute("nowtbank_no");
	//=======================================================================================================================
	//92.12.18取得可下載檔案的報表名稱===============================================================================================
	StringBuffer sqlCmd = new StringBuffer();
	List paramList = new ArrayList();
	//99.09.27 add 查詢年度100年以前.縣市別不同===============================
	String cd01_table = (Integer.parseInt(YEAR) < 100)?"cd01_99":"";
	String wlx01_m_year = (Integer.parseInt(YEAR) < 100)?"99":"100";

	sqlCmd.append("select * from BN01 where bank_no=? and m_year=?");
	paramList.add(bank_code);
	paramList.add(wlx01_m_year);
	List bank_nameData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
	String bank_name = "";
	if(bank_nameData != null && bank_nameData.size()!= 0){
	   bank_name = (String)((DataObject)bank_nameData.get(0)).getValue("bank_name");
	}


	String muser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");
	String muser_name = ( session.getAttribute("muser_name")==null ) ? "" : (String)session.getAttribute("muser_name");
	sqlCmd.delete(0,sqlCmd.length());
	sqlCmd.append(" select a.report_no,c.cmuse_name ");
 	sqlCmd.append(" from  WTT04_1D a,  CDShareNO b,  CDShareNO c ");
 	sqlCmd.append(" where a.MUSER_ID=?");
	sqlCmd.append(" and a.PROGRAM_ID='WMFileDownload'");
	sqlCmd.append(" and a.DETAIL_TYPE='2'");//2-->Download
	sqlCmd.append(" and b.CmUSE_Div = '011'");
	sqlCmd.append(" and (a.TRANSFER_TYPE = b.CmUSE_id and b.CmUSE_Div = '011') ");
    sqlCmd.append(" and  c.CmUSE_Div in ('012',  '013', '014','030') ");//94.11.15 add 030:F01_在台無住所之外國人新台幣存款表
    sqlCmd.append(" and b.Identify_no = c.Identify_no ");
	sqlCmd.append(" and c.cmuse_name like a.report_no||'%' order by a.report_no");
	paramList = new ArrayList();
	paramList.add(muser_id);

    List ReportData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
    String[][] Report_List = new String[ReportData.size()][2];
    if(ReportData != null && ReportData.size() != 0){
       for(int i=0;i<ReportData.size();i++){
           Report_List[i][0]=(String)((DataObject)ReportData.get(i)).getValue("report_no");
	       Report_List[i][1]=(String)((DataObject)ReportData.get(i)).getValue("cmuse_name");
	       if(Report_List[i][1].indexOf("_") != -1){
	          Report_List[i][1]=Report_List[i][1].substring(Report_List[i][1].indexOf("_")+1,Report_List[i][1].length());
	       }
       }
	}


	//取得WMFileDownload的權限
	Properties permission = ( session.getAttribute("WMFileDownload")==null ) ? new Properties() : (Properties)session.getAttribute("WMFileDownload");
	if(permission == null){
       System.out.println("WMFileDownload_Qry.permission == null");
    }else{
       System.out.println("WMFileDownload_Qrt.permission.size ="+permission.size());

    }

    //95.10.02 ADD BY 2495
   	String E_YEAR = YEAR;
	String E_MONTH = MONTH;
	if (E_MONTH.length()==1) E_MONTH="0"+E_MONTH;
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/WMFileDownload.js"></script>
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


<body marginwidth="0" marginheight="0" leftmargin="0" topmargin="0" leftmargin="0" onload="javascript:hiddenDownload(this.document.forms[0],'hidden');">
<form method=post>
 <input type="hidden" name="act" value="">
 <input type="hidden" name="bank_type" value="<%=bank_type%>">
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
                          申報資料查詢下載
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
						  <td align='left' bgcolor='#D8EFEE'>金融機構名稱</td>
						  <td colspan=2 bgcolor='e7e7e7'><select name=Bank_Code_S>
							<%
							int i = 0 ;
							if(dbData != null && dbData.size() > 0){
								while(i<dbData.size()){
							%>
							<option value=<%=(String)((DataObject)dbData.get(i)).getValue("bank_no")%>>
								<%=(String)((DataObject)dbData.get(i)).getValue("bank_no")%>
								&nbsp;&nbsp;&nbsp;
								<%=(String)((DataObject)dbData.get(i)).getValue("bank_name")%>
							</option>
							<%  }
							}%>
							<option value=<%=bank_code%>>
								<%=bank_code%>
								&nbsp;&nbsp;&nbsp;
								<%=bank_name%>
							</option>

						    </td>
	  					</tr>

                        <tr class="sbody">
						  <td align='left' bgcolor='#D8EFEE'>申報資料</td>
						  <%//94.11.15 add 若為F01時,只開放查詢 by 2295%>
						  <td colspan=2 bgcolor='e7e7e7'><select name=Report_no  onchange="javascript:hiddenDownload(this.document.forms[0],'hidden');">
							<%
							for (i = 0; i < Report_List.length; i++) {
							%>
							<option value=<%=Report_List[i][0]%>>
								<%=Report_List[i][0]%>
								&nbsp;&nbsp;&nbsp;
								<%=Report_List[i][1]%>
								<%if(Report_List[i][0].equals("A99")){%>
								  (中央存保用報表格式)
								<%}%>
							</option>
							<%//96.05.08 fix先不單掛if(Report_List[i][0].equals("A99")){%>
							<!--option value='A99_extra'>
								A99&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;A01.A02.A99中央存保用文字檔
							</option-->
							<%//}//end of if%>
							<%}//end of for%>
						    </td>
	  					</tr>

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
        						</select><font color='#000000'>月</font>

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
                      <%//Query //97.01.10地方主管機關.不能查詢.下載資料
                        if(permission != null && permission.get("Q") != null && permission.get("Q").equals("Y") && !bank_type.equals("B")){ %>
                        <td width="66"> <div id="Querybtn" style="POSITION:relative;display:block;width:100%;" align="center"><a href="javascript:doSubmit(this.document.forms[0],'Query','','','','');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_queryb.gif',1)"><img src="images/bt_query.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a></div></td>
                      <%}%>

                      <%//Download //97.01.10地方主管機關.不能查詢.下載資料
                        if(permission != null && permission.get("dl") != null && permission.get("dl").equals("Y")){%>
                        <td width="66"> <div id="Downloadbtn" style="POSITION:relative;display:block;width:100%;" align="center"><%if(!bank_type.equals("B")){%><a href="javascript:doSubmit(this.document.forms[0],'Download','','','','');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image102','','images/bt_downloadb.gif',1)"><img src="images/bt_download.gif" name="Image102" width="66" height="25" border="0" id="Image102"></a><%}%></div></td>
                      <%}%>
                      <%if(permission != null && permission.get("dl") != null && permission.get("dl").equals("Y")){//DownloadReport %>
                        <td width="66"> <div id="DownloadReport" style="POSITION:relative;display:block;width:100%;" align="center"><a href="javascript:doSubmit(this.document.forms[0],'DownloadReport','<%=bank_code%>','<%=bank_name%>','<%=muser_id%>','<%=muser_name%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_download_report.gif',1)"><img src="images/bt_download_report.gif" name="Image103" width="80" height="25" border="0" id="Image103"></a></div></td>
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
                      <li>本網頁提供查詢下載各類申報資料。</li>
					  <li>輸入年月按查詢後可查詢某年度某月份申報資料。</li>
					  <li>輸入年月按下載後可下載某年度某月份申報資料。</li>
					  <li><font color='red'>A99法定比率延申資料(中央存保用報表格式),點選【報表下載】按鈕後,可下載中央存保用的報表格式
					  ,其內容截取自A02法定比率資料及A99法定比率延申資料的申報資料組合而成,下載後可將此表直接申報給中央存保用,若需修改該報表的金額內容,須至A02或A99申報資料做修改。</font></li>
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
</body>
</html>
