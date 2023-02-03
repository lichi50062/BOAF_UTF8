<%
//
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Calendar" %>
<%

	String bank_type=(session.getAttribute("bank_type")==null)?"":(String)session.getAttribute("bank_type");
	String tbank_no=(session.getAttribute("tbank_no")==null)?"":(String)session.getAttribute("tbank_no");
	String act=(request.getParameter("act")==null) ? "" : (String)request.getParameter("act");
	String szreport_no=(request.getParameter("report_no")==null) ? "" : (String)request.getParameter("report_no");
	String szupd_code=(request.getParameter("upd_code")==null) ? "" : (String)request.getParameter("upd_code");

	String gp_itemcnt=(request.getParameter("itemcnt")==null) ? "" : (String)request.getParameter("itemcnt");
	String gp_tbankno=(request.getParameter("tbank_no")==null) ? "" : (String)request.getParameter("tbank_no");
	String gp_bankno=(request.getParameter("bank_no")==null) ? "" : (String)request.getParameter("bank_no");
	String gp_orgcate=(request.getParameter("org_cate")==null) ? "" : (String)request.getParameter("org_cate");

    String lguser_id=(session.getAttribute("muser_id")==null) ? "" : (String)session.getAttribute("muser_id");

	System.out.println("act="+act);
	System.out.println("bank_type="+bank_type);
	System.out.println("szreport_no="+szreport_no);
	System.out.println("szupd_code="+szupd_code);

	List bn01List=(List)request.getAttribute("bn01List");
	List bn02List=(List)request.getAttribute("bn02List");
	List ba01List=(List)request.getAttribute("ba01List");
	List orgList=(List)request.getAttribute("orgList");

	if(orgList==null){
	   System.out.println("orgList==null");
	}else{
	   System.out.println("orgList.size()="+orgList.size());
	}

	//取得ZZ099W的權限
	Properties permission=(session.getAttribute("ZZ099W")==null) ? new Properties() : (Properties)session.getAttribute("ZZ099W");
	if(permission==null){
       System.out.println("ZZ099W_List.permission==null");
    }else{
       System.out.println("ZZ099W_List.permission.size="+permission.size());
    }




%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/ZZ099W.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<title>「申報資料追蹤管理」</title>
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
// -->
</script>
</head>

<body marginwidth="0" marginheight="0" leftmargin="0" topmargin="0" leftmargin="0">
	<form method=post action='#'>
		<input type="hidden" name="act" value="">
<%
	if(orgList!=null && orgList.size()!=0) {
%>
		<input type="hidden" name="row" value="<%=orgList.size()+1%>">
<%
	}
%>
		<table width="640" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
			<tr>
				<td><img src="images/space_1.gif" width="12" height="12"></td>
			</tr>
			<tr>
				<td bgcolor="#FFFFFF">
					<table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
						<tr>
							<td>
							<table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
								<tr>
									<td width="110"><img src="images/banner_bg1.gif" width="110" height="17"></td>
									<td width="380"><font color='#000000' size=4><b><center>「總分支機構組織強制維護」</center></b></font></td>
									<td width="110"><img src="images/banner_bg1.gif" width="110" height="17"></td>
								</tr>
							</table>
							</td>
						</tr>
						<tr>
							<td><img src="images/space_1.gif" width="12" height="12"></td>
						</tr>
						<tr>
							<td>
							<table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
								<tr>
									<div align="right"><jsp:include page="getLoginUser.jsp?width=600" flush="true" /></div>
								</tr>
<%
	String nameColor="nameColor_sbody";
	String textColor="textColor_sbody";

%>

								<tr>
									<td>
									<table width="600" border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
										<tr class="sbody">
											<td width='15%' align='left' bgcolor='#D8EFEE'>機構種類</td>
											<td width='85%' colspan='3' bgcolor='e7e7e7'>
												<select name='orgcate' onChange="javascript:chgInputState(this.form);">
													<option></option>
													<option value="0" <% if (gp_orgcate.equals("0")) out.print("selected"); %>>總機構</option>
													<option value="1" <% if (gp_orgcate.equals("1")) out.print("selected"); %>>分支機構</option>
												</select>

											</td>
										</tr>
										<tr class="sbody">
											<td width='15%' class="<%=nameColor%>">總機構代碼</td>
											<td width='35%' align='left' class="<%=textColor%>">
<%
	if (act.equals("List")) {
%>
												<input type="text" name="TBANK_NO" value="" size=20 maxlength="100">
<%
	}
	if (act.equals("Qry")) {
		if (bn01List.size()==0) {
%>
												<input type="text" name="TBANK_NO" value="<%=gp_tbankno%>" size=20 maxlength="100">
<%
		}
			else {
%>
												<input type="text" name="TBANK_NO" value=<%=((String)((DataObject)bn01List.get(0)).getValue("pbank_no")==null) ? "" : (String)((DataObject)bn01List.get(0)).getValue("pbank_no")%>>
<%
			}
	}
%>
 											</td>
											<td width='15%' class="<%=nameColor%>">名稱</td>
											<td width='35%' class="<%=textColor%>">
<%
	if (act.equals("List")) {
%>
												<input type="text" name="TBANK_NAME" value="" size=30 maxlength="100" readonly>
<%
	}
	if (act.equals("Qry")) {
		if (bn01List.size()==0) {
%>
												<input type="text" name="TBANK_NAME" value="（總機構代碼未存於BN01）" size=30 readonly>
<%
		}
			else {
%>
												<input type="text" name="TBANK_NAME" value=<%=(String)((DataObject)bn01List.get(0)).getValue("pbank_name")%> size=30 readonly>
<%
			}
	}
%>
 											</td>
										</tr>
										<tr class="sbody">
											<td width='15%' class="<%=nameColor%>">分支機構代碼</td>
											<td width='35%' class="<%=textColor%>">
<%
	if (act.equals("List") || (act.equals("Qry") && gp_orgcate.equals("0"))) {
%>
												<input type="text" name="BANK_NO" value="" size=20 maxlength="100">
<%
	}
		else if (act.equals("Qry") && gp_orgcate.equals("1") && bn02List.size()!=0) {
%>
												<input type="text" name="BANK_NO" value=<%=(String)((DataObject)bn02List.get(0)).getValue("bank_no")%> size=20>
<%
		}
		else if (act.equals("Qry") && bn02List.size()==0) {
%>
												<input type="text" name="BANK_NO" value=<%=gp_bankno %> size=20>
<%
		}
%>
										</td>
											<td width='15%' class="<%=nameColor%>">名稱</td>
											<td width='35%' class="<%=textColor%>">
<%
	if (act.equals("List")) {
%>
												<input type="text" name="BANK_NAME" value="" size=30 maxlength="100" readonly>
<%
	}
		else if (act.equals("Qry") && gp_orgcate.equals("0")) {
%>
												<input type="text" name="BANK_NAME" value="" size=30 maxlength="100" readonly>
<%
	}
		else if (act.equals("Qry") && gp_orgcate.equals("1") && bn02List.size()!=0) {
%>
												<input type="text" name="BANK_NAME" value=<%=(String)((DataObject)bn02List.get(0)).getValue("bank_name") %> size=30 readonly>
<%
		}
		else if (act.equals("Qry") && bn02List.size()==0) {
%>
												<input type="text" name="BANK_NAME" value="（分支機構代碼未存於BN02）" size=30 readonly>
<%
		}
%>
 											</td>
										</tr>
										<tr class="sbody">
											<td width='15%' align='left' bgcolor='#D8EFEE'>密碼</td>
											<td width='85%' colspan='3' bgcolor='e7e7e7'>
												<input type="password" maxlength=20 name=user_pw size=20>
											</td>
										</tr>
										<tr class="sbody">
											<td width='15%' align='left' bgcolor='#D8EFEE'>檢核結果</td>
											<td width='85%' colspan='3' bgcolor='e7e7e7'>
												<select name='UPD_CODE'>
													<option value="ALL" <%if(szupd_code.equals("ALL")) out.print("selected");%>>全部</option>
												</select>
												&nbsp;&nbsp;&nbsp;&nbsp;
<%
	if (lguser_id.equals("A111111111")) {
%>
												<a href="javascript:doSubmit(this.document.forms[0],'Qry');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_queryb.gif',1)"><img src="images/bt_query.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>
<%
	}
	if (act.equals("Qry")) {
%>
												<a href="javascript:doSubmit(this.document.forms[0],'Delete');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image102','','images/bt_15b.gif',1)"><img src="images/bt_15.gif" name="Image102" width="66" height="25" border="0" id="Image102"></a>
<%
	}
%>											</td>
										</tr>
									</table>
									</td>
								</tr>
								<tr>
									<td>
										<table width=600 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
											<tr class="sbody">
												<td width='100%' colspan=10 bgcolor='D2F0FF'>
													<a href="javascript:selectAll(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image104','','images/bt_selectallb.gif',1)"><img src="images/bt_selectall.gif" name="Image104" width="80" height="25" border="0" id="Image104"></a>
													<a href="javascript:selectNo(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_selectnob.gif',1)"><img src="images/bt_selectno.gif" name="Image105" width="80" height="25" border="0" id="Image105"></a>
												</td>
											</tr>
<%
	int i=0;
	String bgcolor="#D3EBE0";
	String upd_code="";
	String input_method="";
	String upd_method="";
	if (orgList!=null) {
%>
											<tr class="sbody" bgcolor="#9AD3D0">
												<td width="30">序號</td>
												<td width="30">選項</td>
												<td width="80">檔案代號</td>
												<td width="240">名稱</td>
												<td width="60">查詢紀錄</td>
												<td width="60">檢核結果</td>
												<td width="100">備考</td>
											</tr>
<%
		if (orgList.size()==0) {
%>
											<tr class="sbody" bgcolor="<%=bgcolor%>">
												<td colspan=7 align=center>無資料可供查詢。</td>
											</tr>
											<tr>
<%
		}
		String upd_code_tmp="";
		while (i<orgList.size()) {
			upd_code_tmp="";
			bgcolor=(i % 2==0)?"#e7e7e7":"#D3EBE0";

			String t_itemname="";
			String t_itemcnt="";
			String t_itemcode="";
			t_itemname=(String)((DataObject)orgList.get(i)).getValue("itemname");
			t_itemcnt=(((DataObject)orgList.get(i)).getValue("itemcnt")).toString();
			t_itemcode=(String)((DataObject)orgList.get(i)).getValue("itemcode");

			upd_code_tmp=upd_code;
			if (upd_code.equals("E")) {
				upd_code="錯誤";
			}
				else if (upd_code.equals("U")) {
					upd_code="成功";
				}
%>
											<tr class="sbody" bgcolor="<%=bgcolor%>">
												<td width="30"><%=i+1%></td>
												<td width="30">
													<input type="checkbox" name="isModify_<%=(i+1)%>" value="<% if (((DataObject)orgList.get(i)).getValue("itemcode")!=null) out.print((String)((DataObject)orgList.get(i)).getValue("itemcode"));%>" <%
			if (gp_orgcate.equals("1") && (t_itemname.equals("總分支機構清單（總）") || t_itemcode.equals("BN01"))) out.print("disabled"); %>>
												</td>
												<td width="80">
<%
			if (t_itemcode!=null) out.print(t_itemcode); else out.print("&nbsp;");
%>
												</td>
												<td width="240">
<%
			if (t_itemname!=null) out.print(t_itemname); else out.print("&nbsp;");
%>
												</td>
												<td width="60">
<%
			if (t_itemcnt!=null) out.print(t_itemcnt); else out.print("&nbsp;");
%>
												</td>
												<td width="60">&nbsp;</td>
												<td width="100">
<%
			if (act.equals("List") || (act.equals("Qry") && (ba01List.size()==0))) out.print("&nbsp;");
				else if (t_itemcode.equals("BA01")) out.print("總機構代碼："+(String)((DataObject)ba01List.get(0)).getValue("pbank_no"));
				else out.print("&nbsp;");
%>
												</td>
											</tr>
<%
			i++;
		}//end of while
	}//end of if
%>
										</table>
									</td>
								</tr>
							</table>
							</td>
						</tr>
					</table>
				</td>
			</table>
		</form>
	</body>
</html>