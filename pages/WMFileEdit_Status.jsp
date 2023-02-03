<%
// 93.12.17 add 權限檢核 by 2295
// 94.02.21 add 加上lock_status from wml01_a_v,若lock_status為'Y'..則不能編輯 by 2295
// 94.03.15 fix ,若lock_status為'C'(共用中心鎖定)..則不能編輯 by 2295
// 94.04.15 fix 檢核為0時.顯示"內容值皆為零" by 2295
// 94.11.04 add '030'-->F01_在台無住所之外國人新台幣存款表 by 2295
// 		    add F01加申報年月 by 2295
// 		    add 若有已申報的資料,則顯示預設年月為上個月份,若本月為1月份時.則是申報上個年度的12月份
// 				若無則為WLX_APPLY_INI裡的預設年月 for F01 by 2295
// 94.12.23 add F01 加字 by 2295
// 95.04.10 add A06 加申報年月 by 2295
// 95.05.25 add A99 加申報年月 by 2295
// 95.09.08 add 不是F01時才能新增 by 2295
// 97.06.13 add A10 加申報年月 by 2295
// 99.10.01 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
//101.09.03 add M106/M201/M206只提供檔案上傳 by 2295
//102.04.11 add 調整顯示申報年月查詢結果清單.Axx.SQL by 2295
//103.12.22 add 調整顯示申報年月查詢結果清單.Axx.SQL by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="java.util.*" %>
<%

	String Report_no = ( request.getParameter("Report_no")==null ) ? "" : (String)request.getParameter("Report_no");
	String bank_code = ( request.getParameter("bank_code")==null ) ? "" : (String)request.getParameter("bank_code");
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");
	String bank_type = ( request.getParameter("bank_type")==null ) ? "" : (String)request.getParameter("bank_type");
	String S_YEAR="",S_MONTH="";//94.11.04 add by 2295
	//取出Report中文名稱=======================================================
	String Report_name = ( request.getAttribute("Report_name")==null ) ? "" : (String)request.getAttribute("Report_name");
	//92.12.18取得可上傳檔案的報表名稱===============================================================================================
	String muser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");
	StringBuffer sqlCmd = new StringBuffer();
	List paramList = new ArrayList();
	sqlCmd.append(" select a.report_no,c.cmuse_name ");
 	sqlCmd.append(" from  WTT04_1D a,  CDShareNO b,  CDShareNO c ");
 	sqlCmd.append(" where a.MUSER_ID=?");
	sqlCmd.append(" and a.PROGRAM_ID='WMFileEdit'");
	sqlCmd.append(" and a.DETAIL_TYPE='3'");//3-->Edit
	sqlCmd.append(" and b.CmUSE_Div = '011'");
	sqlCmd.append(" and (a.TRANSFER_TYPE = b.CmUSE_id and b.CmUSE_Div = '011') ");
    sqlCmd.append(" and  c.CmUSE_Div in ('012',  '013', '014','030') "); //94.11.04 add '030'-->F01_在台無住所之外國人新台幣存款表
    sqlCmd.append(" and b.Identify_no = c.Identify_no ");
	sqlCmd.append(" and c.cmuse_name like a.report_no||'%'");
	sqlCmd.append(" and a.report_no=?");
	paramList.add(muser_id);
	paramList.add(Report_no);	
	
    List ReportData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
    if(ReportData != null && ReportData.size() != 0){
       Report_name = (String)((DataObject)ReportData.get(0)).getValue("cmuse_name");
       if(Report_name.indexOf("_") != -1){
	      Report_name=Report_name.substring(Report_name.indexOf("_")+1,Report_name.length());
	   }
    }
	
	List dbData = getStatusData(bank_code,Report_no);
	if(dbData != null){
	   System.out.println("dbData.size()="+dbData.size());
	}else{
		System.out.println("dbData  == null");
	}

	//取得WMFileEdit的權限
	Properties permission = ( session.getAttribute("WMFileEdit")==null ) ? new Properties() : (Properties)session.getAttribute("WMFileEdit");
	if(permission == null){
       System.out.println("WMFileEdit_Status.permission == null");
    }else{
       System.out.println("WMFileEdit_Status.permission.size ="+permission.size());
       System.out.println("WMFileEdit_Status.bank_type="+bank_type);
    }


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
                            <td width=78 class="sbody">申報檔案</td>
                            <%if(Report_no.equals("F01") || Report_no.equals("A06") || Report_no.equals("A99") || Report_no.equals("A08") || Report_no.equals("A10")){%>
                            <td width=275 class="sbody"><%=Report_no%> &nbsp;&nbsp;<%=Report_name%></td>
                            <td width="250" class="sbody">
                            <%}else{%>
                            <td width=329 class="sbody"><%=Report_no%> &nbsp;&nbsp;<%=Report_name%></td>
                            <td width="175">
                            <%}%>
							<div align="left">&nbsp;
							<%if(Report_no.equals("F01") || Report_no.equals("A06") || Report_no.equals("A99") || Report_no.equals("A08") || Report_no.equals("A10")){//94.11.04 F01加申報年月//95.04.10 add A06加申報年月//95.05.25 add A99加申報年月 //97.06.13 add A10加申報年月
							    //若有已申報的資料,則顯示預設年月為上個月份,若本月為1月份時.則是申報上個年度的12月份
							    //若無則為WLX_APPLY_INI裡的預設年月
							    sqlCmd.delete(0,sqlCmd.length());
							    paramList = new ArrayList();
							    sqlCmd.append("select aa.M_YEAR as temp_m_year, aa.M_MONTH  as temp_m_month  from WLX_APPLY_INI  aa where  aa.REPORT_NO = ?");
							    paramList.add(Report_no);
							    List dbData1 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"temp_m_year,temp_m_month");
							    if(dbData1 != null){
							       sqlCmd.delete(0,sqlCmd.length());
							       paramList = new ArrayList();
							       sqlCmd.append(" select count(*)  as  cnt ");
        						   sqlCmd.append(" from WML01  aa");
								   sqlCmd.append(" where aa.M_YEAR  = ?");
								   sqlCmd.append("   and aa.M_MONTH = ?");
								   sqlCmd.append("   and aa.BANK_CODE  = ?");
								   sqlCmd.append("   and aa.REPORT_NO = ?");
								   paramList.add((((DataObject)dbData1.get(0)).getValue("temp_m_year")).toString());
								   paramList.add((((DataObject)dbData1.get(0)).getValue("temp_m_month")).toString());
								   paramList.add(bank_code);
								   paramList.add(Report_no);
								   List dbData2 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"cnt");
								   if(dbData2 != null && ((DataObject)dbData2.get(0)).getValue("cnt") != null){
								      if((((DataObject)dbData2.get(0)).getValue("cnt")).toString().equals("0")){
								         S_YEAR = (((DataObject)dbData1.get(0)).getValue("temp_m_year")).toString();
								         S_MONTH = (((DataObject)dbData1.get(0)).getValue("temp_m_month")).toString();
								      }else{								         
    									 S_YEAR=Utility.getYear();
    									 S_MONTH=Utility.getMonth();
								      }
								   }
							    }



							%>
							    申報年月
							    <input type='text' name='S_YEAR' value="<%=S_YEAR%>" size='3' maxlength='3' onblur='CheckYear(this)' >
        						<font color='#000000'>年
        						<select id="hide1" name=S_MONTH >
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
							<%}%>
							<%if((permission != null && permission.get("A") != null && permission.get("A").equals("Y"))){ //add.95.09.08 不是F01時才能新增%>
							<a href="javascript:doSubmit(this.document.forms[0],'new','<%=Report_no%>','','','<%=bank_type%>')" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image9','','images/bt_addb.gif',1)"><img src="images/bt_add.gif" name="Image9" width="66" height="25" border="0"></a></div></td>
							<%}%>
                          </tr>
                        </Table></td>
                    </tr>

                    <tr>
                      <td><table width='600' border=1 align='center' cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
                          <tr bgcolor='#B1DEDC' class="sbody">
                            <td width=30% align='center' bordercolor="#3A9D99"><font color=#000000>基準日</font></td>
                            <td width=40% align='center' bordercolor="#3A9D99"><font color=#000000>申報日期</font></td>
                            <td width=175 align='center' bordercolor="#3A9D99"><font color=#000000>檢核狀態</font></td>
                          </tr>
<%
				if(dbData.size() == 0){%>
				 <tr bgcolor='#F2F2F2' class="sbody"> <td align=center colspan=3>無已申報過的資料</td></tr>
				<%
				}

				int i = 0;
				String input_method = "";
				String upd_code="";
				String bgcolor="#FFFFE6";
				while(i < dbData.size()){
				      bgcolor = (i % 2 == 0)?"#FFFFE6":"#F2F2F2";
				      if (((String)((DataObject)dbData.get(i)).getValue("input_method")) == null){
							input_method = "";
					  }else if (((String)((DataObject)dbData.get(i)).getValue("input_method")).equals("F")){
							input_method = "檔案上傳";
					  }else if (((String)((DataObject)dbData.get(i)).getValue("input_method")).equals("W")){
							input_method = "線上編輯";
					  }else{
							input_method = "";
					  }
					  upd_code = (String)((DataObject)dbData.get(i)).getValue("upd_code");
					  if (upd_code == null){
					      input_method = input_method + "待檢核";
					  }else if (upd_code.equals("E")){
					      input_method = input_method + "檢核有誤";
					  }else if (upd_code.equals("U")){
						  input_method = input_method + "檢核成功";
					  }else if (upd_code.equals("F")){
					      input_method = "上傳檔案不存在";
					  }else if (upd_code.equals("Z")){
					      input_method = input_method + "內容值皆為零"; //94.04.15
					  }else{
						   input_method = upd_code;
					  }

%>
                            <tr bgcolor='<%=bgcolor%>' class="sbody">
                            <td align='center' bordercolor="#3A9D99">
                            <%/*94.02.21 add 加上lock_status from wml01_a_v,若lock_status為'Y'..則不能編輯 101.09.03 M106/M201/M206只有檔案上傳,不能編輯*/%>
                            <%if( (((DataObject)dbData.get(i)).getValue("lock_status") != null && ((String)((DataObject)dbData.get(i)).getValue("lock_status")).equals("Y")) 
                            || (((DataObject)dbData.get(i)).getValue("wml01_lock_status") != null  && (((String)((DataObject)dbData.get(i)).getValue("wml01_lock_status")).equals("Y") || ((String)((DataObject)dbData.get(i)).getValue("wml01_lock_status")).equals("C")) ) 
                            || ("M106".equals(Report_no) || "M201".equals(Report_no) || "M206".equals(Report_no))
                            ){%>
                                <%=(((DataObject)dbData.get(i)).getValue("inputdate")).toString()%>
                            <%}else{%>
                            <a href="javascript:doSubmit(this.document.forms[0],'Edit','<%=Report_no%>','<%=(((DataObject)dbData.get(i)).getValue("m_year")).toString()%>','<%=(((DataObject)dbData.get(i)).getValue("m_month")).toString()%>','<%=bank_type%>');" ><%=(((DataObject)dbData.get(i)).getValue("inputdate")).toString()%></a>
                            <%}%>
                            </td>
                            <td align='center' bordercolor="#3A9D99"><%=Utility.getCHTdate((((DataObject)dbData.get(i)).getValue("add_date")).toString().substring(0, 10), 0)%>&nbsp;&nbsp;<%=(((DataObject)dbData.get(i)).getValue("add_date")).toString().substring(11, 19)%></td>

                            <td align='center' bordercolor="#3A9D99">
                            <%//System.out.println("upd_code="+upd_code);%>
                            <% if (upd_code != null && upd_code.equals("E")){ %>
                            <a href='WMFileQuery_ShowError.jsp?Report_no=<%=Report_no%>&S_YEAR=<%=(((DataObject)dbData.get(i)).getValue("m_year")).toString()%>&S_MONTH=<%=(((DataObject)dbData.get(i)).getValue("m_month")).toString()%>&bank_code=<%=bank_code%>'><font color=red><%=input_method%></font></a><img src="images/icon_1.gif" width="12" height="12" align="absmiddle"></td>
                            <%}else{%>
                            <font color=black><%=input_method%></font></td>
                            <%}%>

                          </tr>
                          <%
					i++;
					}
				%>
                        </table></td>
                    </tr>
                  </table></td>
              </tr>

              <tr>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td><div align="right"><jsp:include page="getMaintainUser.jsp" flush="true" /></div></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td><div align="center"><a href="javascript:history.back();" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image8','','images/bt_backb.gif',1)"><img src="images/bt_back.gif" name="Image8" width="80" height="25" border="0"></a></div></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
              </tr>
            </table></td>
        </tr>

        <tr>
          <td bgcolor="#FFFFFF"><table width="600" border="0" align="center" cellpadding="1" cellspacing="1">
              <!--tr>
                <td><div align="center"><img src="images/line_1.gif" width="600" height="12"></div></td>
              </tr-->
              <tr>
                <td><table width="600" border="0" cellpadding="1" cellspacing="1" class="sbody">
                    <tr>
                      <td colspan="2"><font color='#990000'><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle"><font color="#007D7D" size="3">使用說明
                        : </font></font></td>
                    </tr>
                    <tr>
                      <td width="16">&nbsp;</td>
                      <td width="577"> <ul>
                          <li>本網頁提供線上編輯各類申報資料。</li>
                          <li>點選年月可修改該月份申報資料。</li>
                          <li>按新增可新增一筆新的月份資料。</li>
                          <%if(Report_no.equals("F01")){//94.11.04 F01加申報年月%>
                          <li><font color="#FF0000"><b>起始申報[第一個月]</b>的資料,務請確實輸入[上月底餘額],若輸入有誤則須刪除再重新『新增』鍵入(<=故第一次務請確實完成『上月底餘額』之輸入&nbsp;</font></li>
                          <li><font color="#FF0000">在下一個月的資料申報時,將會以上一個月的[本月底餘額]轉成本月的[上月底餘額]同時該欄將不會開放更動</font></li>
                          <li><font color="#FF0000">當月的資料有任何錯誤時,須先修正至完全正確,方會提供次一月份的資料申報</font></li>
                          <li><font color="#FF0000">資料自[<b>起始申報]月份起,不管有無資料均須逐月申報</b></font></li>
						  <%}%>
						  <%if(Report_no.equals("M106") || Report_no.equals("M201") || Report_no.equals("M206")){//101.09.03 M106/M201/M206只有檔案上傳%>
                          <li><font color="#FF0000"><b>僅提供檔案上傳</b></font></li>
                          <%}%>
                        </ul></td>
                    </tr>
                  </table></td>
              </tr>
              <!--tr>
                <td><div align="center"><img src="images/line_1.gif" width="600" height="12"></div></td>
              </tr-->
            </table></td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
</form>
</body>
</html>
<%!
	private List getStatusData(String bank_code,String Report_no){
    	StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
    	String BankType="";//-->93.12.17不使用
    	String tablename = "";//-->93.12.17不使用
    	//tablename = ListArray.getTempTableName(BankType, Report_no);-->不使用
    	//for A01
    	//sqlCmd = "SELECT distinct(a.m_year || '/' || substr(100 + a.m_month, 2)) as inputdate, add_date, upd_code, " +
		//		 "a.m_year, a.m_month, input_method FROM " + Report_no +
		//		 " a LEFT JOIN WML01 b ON a.m_year = b.m_year AND a.m_month = b.m_month AND " +
		//		 "a.bank_code = b.bank_code " +
		//		 "WHERE a.bank_code ='" + bank_code + "' AND report_no ='" + Report_no + "' " +
		//		 "ORDER BY a.m_year desc ,a.m_month desc";
		System.out.println(" in get data Report_no="+Report_no+"Report_no.substring(0,1)="+Report_no.substring(0,1));
    	if (Report_no.substring(0,1).equals("A")){
    		//102.04.11 調整SQL by 2295
    		//103.12.22 調整SQL by 2295
    		sqlCmd.append(" SELECT distinct(a.m_year || '/' || substr(100 + a.m_month, 2)) as inputdate, wml01.add_date, wml01.upd_code, ");
			sqlCmd.append(" a.m_year, a.m_month, wml01.input_method, c.wml01_lock_status as lock_status , c.wml01_lock_lock_status as wml01_lock_status FROM ");
			sqlCmd.append(" (select * from "+Report_no+" where bank_code=?)" );
			paramList.add(bank_code);
			sqlCmd.append(" a LEFT JOIN (select * from wml01 where bank_code=? and report_no=?)WML01");  
			paramList.add(bank_code);
			paramList.add(Report_no);
			sqlCmd.append(" on wml01.m_year = a.m_year AND wml01.m_month = a.m_month AND wml01.bank_code = a.bank_code ");
							 //94.02.21 add 加上lock_status by 2295
			sqlCmd.append("   LEFT JOIN ");
			sqlCmd.append(" ( ");
            sqlCmd.append("select * from (SELECT NVL (wml01_bank_code, wml01_lock_bank_code) AS bank_code,");
          	sqlCmd.append(" NVL (wml01_m_year, wml01_lock_m_year) AS m_year,");
          	sqlCmd.append(" NVL (wml01_m_month, wml01_lock_m_month) AS m_month,");
          	sqlCmd.append(" NVL (wml01_report_no, wml01_lock_report_no) AS report_no,");
          	sqlCmd.append(" wml01_lock_status,wml01_lock_lock_status,upd_code,input_method,common_center,upd_method");
     		sqlCmd.append(" FROM ( (SELECT a.m_year AS wml01_m_year,");
            sqlCmd.append("         a.m_month AS wml01_m_month,");
            sqlCmd.append("         a.bank_code AS wml01_bank_code,");
            sqlCmd.append("         a.report_no AS wml01_report_no,");
            sqlCmd.append("         a.lock_status AS wml01_lock_status,");
            sqlCmd.append("         b.m_year AS wml01_lock_m_year,");
            sqlCmd.append("         b.m_month AS wml01_lock_m_month,");
            sqlCmd.append("         b.bank_code AS wml01_lock_bank_code,");
            sqlCmd.append("         b.report_no AS wml01_lock_report_no,");
            sqlCmd.append("         b.lock_status AS wml01_lock_lock_status,");
            sqlCmd.append("         a.upd_code AS upd_code,");
            sqlCmd.append("         a.input_method AS input_method,");
            sqlCmd.append("         a.common_center AS common_center,");
            sqlCmd.append("         a.upd_method AS upd_method");
            sqlCmd.append("    FROM    (select * from wml01 where bank_code=? and report_no=?) a");
            paramList.add(bank_code);
			paramList.add(Report_no);
            sqlCmd.append("         LEFT JOIN (select * from wml01_lock where bank_code=? and report_no=?) b");
            paramList.add(bank_code);
			paramList.add(Report_no);
            sqlCmd.append("         ON  a.m_year = b.m_year AND a.m_month = b.m_month AND a.bank_code = b.bank_code AND a.report_no = b.report_no)");
            sqlCmd.append(" UNION ");
            sqlCmd.append(" (SELECT a.m_year AS wml01_m_year, ");
            sqlCmd.append("         a.m_month AS wml01_m_month, ");
            sqlCmd.append("         a.bank_code AS wml01_bank_code, ");
            sqlCmd.append("         a.report_no AS wml01_report_no, ");
            sqlCmd.append("         a.lock_status AS wml01_lock_status,");
            sqlCmd.append("         b.m_year AS wml01_lock_m_year,");
            sqlCmd.append("         b.m_month AS wml01_lock_m_month,");
            sqlCmd.append("         b.bank_code AS wml01_lock_bank_code,");
            sqlCmd.append("         b.report_no AS wml01_lock_report_no,");
            sqlCmd.append("         b.lock_status AS wml01_lock_lock_status,");
            sqlCmd.append("         a.upd_code AS upd_code,");
            sqlCmd.append("         a.input_method AS input_method,");
            sqlCmd.append("         a.common_center AS common_center,");
            sqlCmd.append("         a.upd_method AS upd_method");
            sqlCmd.append(" FROM     (select * from wml01_lock where bank_code=? and report_no=?) b");
            paramList.add(bank_code);
			paramList.add(Report_no);
            sqlCmd.append("        LEFT JOIN (select * from wml01 where bank_code=? and report_no=?)  a ");
            paramList.add(bank_code);
			paramList.add(Report_no);
            sqlCmd.append("        ON  a.m_year = b.m_year AND a.m_month = b.m_month AND a.bank_code = b.bank_code AND a.report_no = b.report_no ");
            sqlCmd.append("           )))wml01_a_v )c");
			sqlCmd.append(" on c.m_year  = a.m_year and c.m_month = a.m_month and c.bank_code = a.bank_code ");		
			sqlCmd.append(" where add_date is not null  ");	
			sqlCmd.append(" ORDER BY a.m_year desc ,a.m_month desc");			
		}else if (Report_no.substring(0,1).equals("B") && !Report_no.equals("B03")){
    		sqlCmd.append(" SELECT distinct(a.m_year || '/' || substr(100 + a.m_month, 2)) as inputdate, b.add_date, b.upd_code, ");
			sqlCmd.append(" a.m_year, a.m_month, b.input_method, c.wml01_lock_status as lock_status , c.wml01_lock_lock_status as wml01_lock_status FROM " + Report_no);
			sqlCmd.append(" a LEFT JOIN WML01 b ON a.m_year = b.m_year AND a.m_month = b.m_month AND b.bank_code =?");
							 //94.02.21 add 加上lock_status by 2295
			sqlCmd.append("   LEFT JOIN wml01_a_v c on b.bank_code = c.bank_code and ");
			sqlCmd.append("					 		  a.m_year  = c.m_year and ");
			sqlCmd.append("					  		  a.m_month = c.m_month and ");
			sqlCmd.append("							  b.report_no = c.report_no ");
			sqlCmd.append(" WHERE b.bank_code =? AND b.report_no =? ");
			sqlCmd.append(" ORDER BY a.m_year desc ,a.m_month desc");
			paramList.add(bank_code);
			paramList.add(bank_code);
			paramList.add(Report_no);
		}else if (Report_no.substring(0,1).equals("B") && Report_no.equals("B03")){
    		sqlCmd.append(" SELECT distinct(a.m_year || '/' || substr(100 + a.m_month, 2)) as inputdate, b.add_date, b.upd_code, ");
			sqlCmd.append(" a.m_year, a.m_month, b.input_method , c.wml01_lock_status as lock_status , c.wml01_lock_lock_status as wml01_lock_status FROM " + Report_no + "_1");
			sqlCmd.append(" a LEFT JOIN WML01 b ON a.m_year = b.m_year AND a.m_month = b.m_month AND b.bank_code =?");
							 //94.02.21 add 加上lock_status by 2295
			sqlCmd.append("   LEFT JOIN wml01_a_v c on b.bank_code = c.bank_code and ");
			sqlCmd.append("					 		  a.m_year  = c.m_year and ");
			sqlCmd.append("					  		  a.m_month = c.m_month and ");
			sqlCmd.append("							  b.report_no = c.report_no ");
			sqlCmd.append(" WHERE b.bank_code =? AND b.report_no =?");
			sqlCmd.append(" ORDER BY a.m_year desc ,a.m_month desc");
			paramList.add(bank_code);
			paramList.add(bank_code);
			paramList.add(Report_no);
		}else if (Report_no.substring(0,1).equals("M") || Report_no.substring(0,1).equals("F")){
    		sqlCmd.append(" SELECT distinct(a.m_year || '/' || substr(100 + a.m_month, 2)) as inputdate, b.add_date, b.upd_code, ");
			sqlCmd.append(" a.m_year, a.m_month, b.input_method, c.wml01_lock_status as lock_status , c.wml01_lock_lock_status as wml01_lock_status FROM " + Report_no);
			sqlCmd.append(" a LEFT JOIN WML01 b ON a.m_year = b.m_year AND a.m_month = b.m_month AND b.bank_code =?");
							 //94.02.21 add 加上lock_status by 2295
			sqlCmd.append("   LEFT JOIN wml01_a_v c on b.bank_code = c.bank_code and ");
			sqlCmd.append("					 		  a.m_year  = c.m_year and ");
			sqlCmd.append("					  		  a.m_month = c.m_month and ");
			sqlCmd.append("							  b.report_no = c.report_no ");
			sqlCmd.append(" WHERE b.bank_code =? AND b.report_no =?");
			sqlCmd.append(" ORDER BY a.m_year desc ,a.m_month desc");
			paramList.add(bank_code);
			paramList.add(bank_code);
			paramList.add(Report_no);
		}
		
		List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"add_date,m_year,m_month");
		return dbData;
    }

%>