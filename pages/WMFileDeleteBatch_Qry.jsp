<%
//94.04.14 add 申報資料批次刪除 by 2295
//99.10.08 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
//         add 根據查詢年度.縣市別.改變總機構名稱 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>

<%
	String bank_type = ( request.getParameter("bank_type")==null ) ? "" : (String)request.getParameter("bank_type");
	String lgbank_type = (session.getAttribute("bank_type")==null)?"":(String)session.getAttribute("bank_type");
	//若有已點選的tbank_no,則以已點選的tbank_no為主============================================================
	String bank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");
	String nowtbank_no =  ( request.getParameter("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");
	if(!nowtbank_no.equals("")){
	   session.setAttribute("nowtbank_no",nowtbank_no);//將已點選的tbank_no寫入session
	}
	bank_no = ( session.getAttribute("nowtbank_no")==null ) ? bank_no : (String)session.getAttribute("nowtbank_no");

	String hsien_id = ( session.getAttribute("HSIEN_ID")==null ) ? "ALL" : (String)session.getAttribute("HSIEN_ID");
	if(session.getAttribute("HSIEN_ID") == null){
	   System.out.println("HSIEN_ID == null");
	}else{
	   System.out.println("hsien_id"+(String)session.getAttribute("HSIEN_ID"));
	}
	String YEAR  = Utility.getYear();
   	String MONTH = Utility.getMonth();

    //99.10.08 add 查詢年度100年以前.縣市別不同===============================
	String cd01_table = (Integer.parseInt(YEAR) < 100)?"cd01_99":"";
	String wlx01_m_year = (Integer.parseInt(YEAR) < 100)?"99":"100";

    String firstStatus = ( request.getParameter("firstStatus")==null ) ? "" : (String)request.getParameter("firstStatus");
    System.out.println("firstStatus="+firstStatus);
    //清空已點選的BankList
    if(firstStatus.equals("true")){//若從Menu點選時,先清空session裡的資料
	   session.setAttribute("BankList",null);
	}
	//取得可下載檔案的報表名稱===============================================================================================
	StringBuffer sqlCmd = new StringBuffer();
	List paramList = new ArrayList();
	String muser_id = ( session.getAttribute("muser_id")==null ) ? "A111111111" : (String)session.getAttribute("muser_id");
	sqlCmd.append(" select a.report_no,c.cmuse_name ");
 	sqlCmd.append(" from  WTT04_1D a,  CDShareNO b,  CDShareNO c ");
 	sqlCmd.append(" where a.MUSER_ID=?");
	sqlCmd.append(" and a.PROGRAM_ID=?");
	sqlCmd.append(" and a.DETAIL_TYPE='2'");//2-->Download
	sqlCmd.append(" and b.CmUSE_Div = '011'");
	sqlCmd.append(" and (a.TRANSFER_TYPE = b.CmUSE_id and b.CmUSE_Div = '011') ");
    sqlCmd.append(" and  c.CmUSE_Div in ('012',  '013', '014') ");
    sqlCmd.append(" and b.Identify_no = c.Identify_no ");
	sqlCmd.append(" and c.cmuse_name like a.report_no||'%' order by a.report_no");
	paramList.add(muser_id);
	paramList.add("WMFileDownload");

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
	sqlCmd.delete(0,sqlCmd.length()); 	   
 	paramList = new ArrayList();
	sqlCmd.append(" select bn01.m_year,bn01.bn_type,wlx01.hsien_id,bn01.bank_no, bn01.bank_name from bn01,wlx01 ");
	if(!(lgbank_type.equals("8") || bank_type.equals("8"))){//登入者的共用中心代碼
	   sqlCmd.append(" where bank_type = ? and bn01.m_year = wlx01.m_year ");
	   paramList.add(bank_type);
	}else{
	  sqlCmd.append(" where bn01.m_year = wlx01.m_year and ");
	}
	sqlCmd.append(" bn01.bank_no = wlx01.bank_no ");
	sqlCmd.append(" and bn01.bn_type <> '2'");

	if(lgbank_type.equals("8") || bank_type.equals("8")){//登入者的共用中心代碼
	   sqlCmd.append(" and wlx01.bank_no in (select bank_no from wlx01 where center_flag='Y' and center_no = ?)");
	   paramList.add(bank_no);
	}

	sqlCmd.append(" order by wlx01.hsien_id,bn01.bank_no");
    List tbankList = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year");

	// XML Ducument for 總機構代碼 begin
    out.println("<xml version=\"1.0\" encoding=\"UTF-8\" ID=\"TBankXML\">");
    out.println("<datalist>");
    for(int i=0;i< tbankList.size(); i++) {
        DataObject bean =(DataObject)tbankList.get(i);
        out.println("<data>");
        out.println("<BankYear>"+bean.getValue("m_year").toString()+"</BankYear>");
        out.println("<BnType>"+bean.getValue("bn_type")+"</BnType>");
        out.println("<HsienId>"+bean.getValue("hsien_id")+"</HsienId>");
        out.println("<bankValue>"+bean.getValue("bank_no")+"</bankValue>");
        out.println("<bankName>"+bean.getValue("bank_no")+bean.getValue("bank_name")+"</bankName>");
        out.println("</data>");
    }
    out.println("</datalist>\n</xml>");
    // XML Ducument for 總機構代碼 end
    
    List cityList = Utility.getCity();
	// XML Ducument for 縣市別 begin
    out.println("<xml version=\"1.0\" encoding=\"UTF-8\" ID=\"CityXML\">");
    out.println("<datalist>");
    for(int i=0;i< cityList.size(); i++) {
        DataObject bean =(DataObject)cityList.get(i);
        out.println("<data>");
        out.println("<cityType>"+bean.getValue("hsien_id")+"</cityType>");
        out.println("<cityName>"+bean.getValue("hsien_name")+"</cityName>");
        out.println("<cityValue>"+bean.getValue("hsien_id")+"</cityValue>");
        out.println("<cityYear>"+bean.getValue("m_year").toString()+"</cityYear>");
        out.println("</data>");
    }
    out.println("</datalist>\n</xml>");
    // XML Ducument for 縣市別 end
    

	//取得WMFileDownload的權限
	Properties permission = ( session.getAttribute("WMFileDeleteBatch")==null ) ? new Properties() : (Properties)session.getAttribute("WMFileDeleteBatch");
	if(permission == null){
       System.out.println("WMFileDeleteBatch_Qry.permission == null");
    }else{
       System.out.println("WMFileDeleteBatch_Qry.permission.size ="+permission.size());
    }

%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/WMFileDeleteBatch.js"></script>
<script language="javascript" src="js/BRUtil.js"></script>
<script language="javascript" src="js/movesels.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<title>申報資料批次刪除</title>
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
                          申報資料批次刪除
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
						  <td width="15%" align='left' bgcolor='#D8EFEE'>申報資料</td>
						  <td colspan=2 bgcolor='e7e7e7'><select name=Report_no>
							<%
							for (int i = 0; i < Report_List.length; i++) {
							%>
							<option value=<%=Report_List[i][0]%>>
								<%=Report_List[i][0]%>
								&nbsp;&nbsp;&nbsp;
								<%=Report_List[i][1]%>
							</option>
							<%}%>
							</select>
							<%if(permission != null && permission.get("D") != null && permission.get("D").equals("Y")){//Delete %>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                            <a href="javascript:doSubmit(this.document.forms[0],'Delete');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image102','','images/bt_deleteb.gif',1)"><img src="images/bt_delete.gif" name="Image102" width="66" height="25" border="0" id="Image102"></a>
                            <%}%>
						    </td>
	  					</tr>

	  					<tr class="sbody">
							<td width="15%" align='left' bgcolor='#D8EFEE'>基準日</td>
							<td colspan=2 bgcolor='e7e7e7'>
                            <input type='text' name='S_YEAR' value="<%=YEAR%>" size='3' maxlength='3' onblur='CheckYear(this)' onchange="javascript:changeCity('CityXML', this.document.forms[0].HSIEN_ID, this.document.forms[0].S_YEAR, this.document.forms[0]);changeOption(document.forms[0],'');">
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

                        <tr class="sbody">
                         <td width="15%" align='left' bgcolor='#D8EFEE'>金融機構</td>
                        <td bgcolor="#e7e7e7"> <table width="200" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#e7e7e7">
                          <tr>
                            <td width="195"> <span class="mtext">縣市別 :</span>
                              <select name="HSIEN_ID" onchange="javascript:changeOption(document.forms[0],'');"></select>  	
                            <select multiple  size=10  name="BankListSrc" ondblclick="javascript:movesel(this.document.forms[0].BankListSrc,this.document.forms[0].BankListDst);" style="width: 17em">
							</select>
                            </td>
                            <td width="52"><table width="40" border="0" align="center" cellpadding="3" cellspacing="3">
                                <tr>
                                  <td>
                                  <div align="center">
                                  <a href="javascript:movesel(this.document.forms[0].BankListSrc,this.document.forms[0].BankListDst);"><img src="images/arrow_right.gif" width="24" height="22" border="0"></a>
                                  </div>
                                  </td>
                                </tr>
                                <tr>
                                  <td>
                                  <div align="center">
                                  <a href="javascript:moveallsel(this.document.forms[0].BankListSrc,this.document.forms[0].BankListDst);"><img src="images/arrow_rightall.gif" width="24" height="22" border="0"></a>
                                  </div>
                                  </td>
                                </tr>
                                <tr>
                                  <td>
                                  <div align="center">
                                  <a href="javascript:movesel(this.document.forms[0].BankListDst,this.document.forms[0].BankListSrc);"><img src="images/arrow_left.gif" width="24" height="22" border="0"></a>
                                  </div>
                                  </td>
                                </tr>
                                <tr>
                                  <td height="22">
                                  <div align="center">
                                  <a href="javascript:moveallsel(this.document.forms[0].BankListDst,this.document.forms[0].BankListSrc);"><img src="images/arrow_leftall.gif" width="24" height="22" border="0"></a>
                                  </div>
                                  </td>
                                </tr>
                              </table></td>
                            <td width="189"> <br>
                            <select multiple size=10  name="BankListDst" ondblclick="javascript:movesel(this.document.forms[0].BankListDst,this.document.forms[0].BankListSrc);" style="width: 17em">
							</select>
                          </tr>
                        </table></td>

                        </Table></td>
                    </tr>
                    <tr>
                <td><div align="right"><jsp:include page="getMaintainUser.jsp" flush="true" /></div></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
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
                      <li>本網頁提供批次刪除各類申報資料。</li>
					  <li>輸入年月按刪除後可刪除某年度某月份申報資料。</li>
					  <li><font color='red'>該申報資料未被鎖定時,方可刪除</font>。</li>
                      </ul>
                      </td>
                    </tr>
                  </table></td>
              </tr>
              <!--tr>
                <td><div align="center"><img src="images/line_1.gif" width="600" height="12"></div></td>
              </tr-->
</table>
<INPUT type="hidden" name=BankList><!--//BankList儲存已勾選的金融機構代碼-->
</form>
<script language="JavaScript" >
<!--

<%
//從session裡把勾選的金融機構代碼讀出來.放在BankListDst
if(session.getAttribute("BankList") != null && !((String)session.getAttribute("BankList")).equals("")){
   System.out.println("BR002W_BankList.BankList="+(String)session.getAttribute("BankList"));
%>
var bnlist;
bnlist = '<%=(String)session.getAttribute("BankList")%>';
var a = bnlist.split(',');
for (var i =0; i < a.length; i ++){
	var j = a[i].split('+');
	this.document.forms[0].BankListDst.options[i] = new Option(j[1], j[0]);
}
<%}%>

setSelect(this.document.forms[0].HSIEN_ID,"<%=hsien_id%>");
changeCity('CityXML', this.document.forms[0].HSIEN_ID, this.document.forms[0].S_YEAR, this.document.forms[0]);
//changeOption(this.document.forms[0],'');
function clearBankList(){
 <%
	session.setAttribute("BankList",null);//清除已勾選的BankList
 %>
}
-->
</script>
</body>
</html>
