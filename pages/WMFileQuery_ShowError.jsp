<%
//93.12.17 add 設定WML02.WML03異動者資訊
//94.03.02 fix 取得bank_no所屬的bank_type by 2295
//94.03.07 若小數點最後一位是"0"時,把"0"去掉 by 2295
//94.03.10 fix A04顯示WML02檢核公式的錯誤 by 2295
//94.03.11 fix Report_no="A05"時,公式檔不區別農漁會 by 2295
//94.03.11 fix Report_no="A04"時,公式檔不區別農漁會 by 2295
//94.11.15 add 030:F01_在台無住所之外國人新台幣存款表 by 2295
//95.05.17 add 顯示A02的檢核錯誤 by 2295
//99.09.27 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
//104.10.12 add A13 by 2295
%>

<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.ListArray_FC" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="java.util.*" %>
<%@ page import="java.math.BigDecimal" %>

<%
	//登入者的bank_type
	String bank_type = ( session.getAttribute("bank_type")==null ) ? "" : (String)session.getAttribute("bank_type");
	String Report_no = ( request.getParameter("Report_no")==null ) ? "" : (String)request.getParameter("Report_no");
	String bank_code = ( request.getParameter("bank_code")==null ) ? "" : (String)request.getParameter("bank_code");
	String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? "" : (String)request.getParameter("S_YEAR");
	String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? "" : (String)request.getParameter("S_MONTH");
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");
	//=======================================================================================================================
	//fix 93.12.20 若有已點選的bank_type,則以已點選的bank_type為主============================================================
	bank_type = ( request.getParameter("bank_type")==null ) ? bank_type : (String)request.getParameter("bank_type");
	bank_type = ( session.getAttribute("nowbank_type")==null ) ? bank_type : (String)session.getAttribute("nowbank_type");
	//=======================================================================================================================
    //99.09.27 add 查詢年度100年以前.縣市別不同===============================
	String cd01_table = (Integer.parseInt(S_YEAR) < 100)?"cd01_99":""; 
	String wlx01_m_year = (Integer.parseInt(S_YEAR) < 100)?"99":"100"; 
	//取出Report中文名稱=======================================================
	String Report_name = ( request.getAttribute("Report_name")==null ) ? "" : (String)request.getAttribute("Report_name");
	//92.12.18取得可上傳檔案的報表名稱===============================================================================================
	StringBuffer sqlCmd = new StringBuffer();
	List paramList = new ArrayList();
	String muser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");
	sqlCmd.append(" select a.report_no,c.cmuse_name ");
 	sqlCmd.append(" from  WTT04_1D a,  CDShareNO b,  CDShareNO c ");
 	sqlCmd.append(" where a.MUSER_ID=?");
	sqlCmd.append(" and a.PROGRAM_ID='WMFileEdit'");
	sqlCmd.append(" and a.DETAIL_TYPE='3'");//3-->Edit
	sqlCmd.append(" and b.CmUSE_Div = '011'");
	sqlCmd.append(" and (a.TRANSFER_TYPE = b.CmUSE_id and b.CmUSE_Div = '011') ");
    sqlCmd.append(" and  c.CmUSE_Div in ('012',  '013', '014','030') "); //94.11.15 add 030:F01_在台無住所之外國人新台幣存款表
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

%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/WMFileQuery.js"></script>
<script language="javascript" event="onresize" for="window"></script>

<html>
<head>
<title>錯誤資料檢視</title>
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
                          錯誤資料檢視
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
                            <td width=329 class="sbody"><%=Report_no%> &nbsp;&nbsp;<%=Report_name%></td>
                          </tr>
                        </Table></td>
                    </tr>

                    <tr>
                      <td><table width='600' border=1 align='center' cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
                      <%
                        String bgcolor="#FFFFE6";
						//show WML02 error
						//94.03.10 add 顯示A04的檢核錯誤
						//95.05.17 add 顯示A02的檢核錯誤
						if (Report_no.equals("A01") || Report_no.equals("A02") || Report_no.equals("A04") || Report_no.equals("A05") || Report_no.equals("A13") || Report_no.equals("B01") || Report_no.equals("B05")){
				      %>
                          <tr bgcolor='#B1DEDC' class="sbody">
                            <td width=70% align='center' bordercolor="#3A9D99"><font color=#000000>公式</font></td>
                            <td width=15% align='center' bordercolor="#3A9D99"><font color=#000000>左式金額</font></td>
                            <td width=15% align='center' bordercolor="#3A9D99"><font color=#000000>右式金額</font></td>
                          </tr>
                      <%
                          	//94.03.02 fix 取得bank_no所屬的bank_type
                          	sqlCmd.delete(0,sqlCmd.length());
                          	paramList = new ArrayList();
                            sqlCmd.append(" select * from ba01 where bank_no=? and m_year=?");
                            paramList.add(bank_code);
                            paramList.add(wlx01_m_year);
                            List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
                            System.out.println("ba01.dbData.size="+dbData.size());
                            if(dbData != null && dbData.size() != 0){
                               bank_type=(String)((DataObject)dbData.get(0)).getValue("bank_type");
                            }
                          	//show WML02 error
                          	sqlCmd.delete(0,sqlCmd.length());
                          	paramList = new ArrayList();
							sqlCmd.append("SELECT WML02.cano, l_amt, r_amt ,quop FROM WML02 LEFT JOIN ruleno1 ON  WML02.cano = ruleno1.cano ");
							//94.03.11 fix 若Report_no != 'A05'時,才區分bank_type
							//94.03.14 fix 若Report_no != 'A04'時,才區分bank_type				
							//104.10.12 add 若Report_no != 'A13'時,才區分bank_type						
							if (!Report_no.equals("A04") && !Report_no.equals("A05")){
								 sqlCmd.append(" AND ruleno1.acc_type=?");
								 paramList.add(bank_type);
							}
							sqlCmd.append(" WHERE m_year=?");
							sqlCmd.append(" AND m_month=? AND bank_code=? AND report_no=? ORDER BY WML02.cano");
							paramList.add(S_YEAR);
							paramList.add(S_MONTH);
							paramList.add(bank_code);
							paramList.add(Report_no);        

							dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"l_amt,r_amt");
							System.out.println("dbData.size="+dbData.size());

							int i = 0 ;
							String lstr = "";
							String rstr = "";
							String lamt = "";
							String ramt = "";
							String acc_code = "";
							while (i < dbData.size()) {
							        bgcolor = (i % 2 == 0)?"#FFFFE6":"#F2F2F2";
							        //93.12.17設定異動者資訊
							        request.setAttribute("maintainInfo","select * from WML02 WHERE m_year=" + S_YEAR +
											   	         " AND m_month=" + S_MONTH + " AND bank_code='" + bank_code + "' AND report_no='" + Report_no + "' ORDER BY WML02.cano");

									lstr = ""; rstr = "";
									lamt = ""; ramt = "";
									sqlCmd.delete(0,sqlCmd.length());
                          	        paramList = new ArrayList();
									sqlCmd.append("SELECT acc_code, noop, left_flag, nserial FROM ruleno2 WHERE");
									//94.03.11 fix Report_no="A05"時,公式檔不區別農漁會
									//94.03.14 fix Report_no="A04"時,公式檔不區別農漁會									
									if(Report_no.equals("A05") || Report_no.equals("A04")){//不分bank_type
									   sqlCmd.append(" cano=? ORDER BY 3, 4");
									   paramList.add((String)((DataObject)dbData.get(i)).getValue("cano"));
									}else{
									   sqlCmd.append(" acc_type=? ");
							 		   sqlCmd.append(" AND cano=? ORDER BY 3, 4");
							 		   paramList.add(bank_type); 
							 		   paramList.add((String)((DataObject)dbData.get(i)).getValue("cano"));         
							 		}
				    				List dbData2 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"nserial");
									System.out.println("dbData2.size="+dbData2.size());
									int j = 0;
									while (j < dbData2.size()) {
										if (((String)((DataObject)dbData2.get(j)).getValue("acc_code")).substring(0, 3).equals("FIX"))	//如果是常數，就將常數值取出
											acc_code = "(" + ListArray_FC.getCONST_VALUE((String)((DataObject)dbData2.get(j)).getValue("acc_code")) + ")";
										else
											acc_code = (String)((DataObject)dbData2.get(j)).getValue("acc_code");

										if (((String)((DataObject)dbData2.get(j)).getValue("left_flag")).equals("0"))
											lstr += " " + acc_code + " " + (((String)((DataObject)dbData2.get(j)).getValue("noop")) == null ? "" : ((String)((DataObject)dbData2.get(j)).getValue("noop")).trim());
										if (((String)((DataObject)dbData2.get(j)).getValue("left_flag")).equals("1"))
											rstr += " " + acc_code + " " + (((String)((DataObject)dbData2.get(j)).getValue("noop")) == null ? "" : ((String)((DataObject)dbData2.get(j)).getValue("noop")).trim());
										//System.out.println("lstr="+lstr);
										//System.out.println("rstr="+rstr);
										j ++ ;
									} //end of while(dbData2--ruleno2)
									lamt = (((((DataObject)dbData.get(i)).getValue("l_amt")).toString()) == null) ? "0" : ((((DataObject)dbData.get(i)).getValue("l_amt")).toString());
									ramt = (((((DataObject)dbData.get(i)).getValue("r_amt")).toString()) == null) ? "0" : ((((DataObject)dbData.get(i)).getValue("r_amt")).toString());
									//94.03.07若小數點最後一位是"0"時,把"0"去掉===========
									if(lamt.indexOf(".") != -1){
									   if(lamt.lastIndexOf("0") != -1){
									       if(lamt.indexOf(".") < lamt.lastIndexOf("0")){
									          lamt = lamt.substring(0,lamt.lastIndexOf("0"));
									       }
									   }
									}
									if(ramt.indexOf(".") != -1){
									   if(ramt.lastIndexOf("0") != -1){
									       if(ramt.indexOf(".") < ramt.lastIndexOf("0")){
									          ramt = ramt.substring(0,ramt.lastIndexOf("0"));
									       }
									   }
									}
									//====================================================
									%>

								<tr bgcolor='<%=bgcolor%>' class="sbody">
					    			<td align=left><font size=2><%=lstr + (((String)((DataObject)dbData.get(i)).getValue("quop")) == null ? " " : ((String)((DataObject)dbData.get(i)).getValue("quop"))) + rstr %></font></td>
					    			<td align=right><font size=2><%=Utility.setCommaFormat(lamt)%></font></td>
					   				<td align=right><font size=2><%=Utility.setCommaFormat(ramt)%></font></td>
								</tr>
						<%		i ++;
							} //while()
                      	  }//end if%>
                     </table></td>
                     <tr><td><table border=0><tr><br></tr></table></td></tr>

                     <tr><td><table width='600' border=1 align='center' cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
                     	     <tr bgcolor='#e7e7e7' class="sbody">
                                <td colspan=2 align='left' bordercolor="#3A9D99"><font color=#000000>其它錯誤原因</font></td>
                             </tr>
                             <tr bgcolor='#B1DEDC' class="sbody">
                            	<td width=10% align='left' bordercolor="#3A9D99"><font color=#000000>序號</font></td>
                            	<td align='left' bordercolor="#3A9D99"><font color=#000000>原因</font></td>
                             </tr>
                          	<%
         					//show WML03 error
         					sqlCmd.delete(0,sqlCmd.length());
                          	paramList = new ArrayList();
							sqlCmd.append(" SELECT serial_no, remark FROM WML03 WHERE m_year=? AND m_month=?");
					 		sqlCmd.append(" AND bank_code=? AND report_no=?");
					 		sqlCmd.append(" ORDER BY serial_no");
					 		paramList.add(S_YEAR);		
					 		paramList.add(S_MONTH);		
					 		paramList.add(bank_code);	
					 		paramList.add(Report_no);	
							List WML03Data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"serial_no");
							int i = 0;
							bgcolor="#FFFFE6";
							while (i < WML03Data.size()) {
							    bgcolor = (i % 2 == 0)?"#FFFFE6":"#F2F2F2";
							    //93.12.17設定異動者資訊
								request.setAttribute("maintainInfo","SELECT * from WML03 WHERE m_year=" + S_YEAR + " AND m_month=" +
					 								 S_MONTH + " AND bank_code='" + bank_code + "' AND report_no='" + Report_no + "' " +
					 								 "ORDER BY serial_no");
							%>
								<tr bgcolor='<%=bgcolor%>' class="sbody">
					    			<td align=left><font size=2><%=(((((DataObject)WML03Data.get(i)).getValue("serial_no")).toString()) == null) ? "" : ((((DataObject)WML03Data.get(i)).getValue("serial_no")).toString())%></font></td>
					    			<td align=left><font size=2><%=(((String)((DataObject)WML03Data.get(i)).getValue("remark")) == null ? " " : ((String)((DataObject)WML03Data.get(i)).getValue("remark")))%></font></td>
								</tr>
							<%	i ++;
							}%>
                             </table>
                         </td>
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
            </table></td>
        </tr>
      </table></td>
  </tr>
</table>
</form>
</body>
</html>
