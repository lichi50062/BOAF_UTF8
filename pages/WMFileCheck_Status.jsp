<%
//94.04.15 fix 檢核為0時.顯示"內容值皆為零" by 2295
//99.09.24 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>


<%
    System.out.println("WMFileCheck_Status");
	List dbData = (List)request.getAttribute("dbData");
	String Report_no = ( request.getParameter("Report_no")==null ) ? "" : (String)request.getParameter("Report_no");			
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");		
	
	String muser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		
	StringBuffer sqlCmd = new StringBuffer();
	List paramList = new ArrayList();
	String cmuse_div = "";
	//92.12.18取得可上傳檔案的報表名稱===============================================================================================	
	if(Report_no.indexOf("A") != -1){	
	   cmuse_div = "012";	   
	}else if(Report_no.indexOf("M") != -1){
	   cmuse_div = "013";	   
	}else if(Report_no.indexOf("B") != -1){
	   cmuse_div = "014";	   	   
	}   
		
	sqlCmd.append(" select * from cdshareno where cmuse_div = ? and cmuse_name like ?");	   
	paramList.add(cmuse_div);
	paramList.add(Report_no+"%");
    List Report_nameData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");		   
    String tmpReport_Name = "";
    String[][] Report_List = new String[Report_nameData.size()][2];
    if(Report_nameData != null && Report_nameData.size() != 0){		   
       for(int i=0;i<Report_nameData.size();i++){           
	       tmpReport_Name=(String)((DataObject)Report_nameData.get(i)).getValue("cmuse_name");
	       System.out.println("tmpReport_Name="+tmpReport_Name);
	       if(tmpReport_Name.indexOf("_") != -1){
	          Report_List[i][1]=tmpReport_Name.substring(tmpReport_Name.indexOf("_")+1,tmpReport_Name.length());
	          Report_List[i][0]=tmpReport_Name.substring(0,tmpReport_Name.indexOf("_"));
	       }
       }
	}	   		       
	System.out.println("Report_List.size="+Report_List.length);
	System.out.println("Report_List[0][0]="+Report_List[0][0]);
	System.out.println("Report_List[0][1]="+Report_List[0][1]);	
	
	if(dbData != null){
	   System.out.println("dbData.size()="+dbData.size());
	   System.out.println("Report_no="+Report_no);	   
	}else{
		System.out.println("dbData  == null");
	}
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/WMFileCheck.js"></script>
<script language="javascript" event="onresize" for="window"></script>

<html>
<head>
<title>人工更新作業</title>
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

<form name=frmWMFileCheck method=post>
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
                          人工更新作業 
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
                            <td width=78 class="sbody">申報資料</td>
                            <td width=329 class="sbody"><%=Report_List[0][0]%> &nbsp;&nbsp;<%=Report_List[0][1]%></td>                            
                          </tr>
                        </Table></td>
                    </tr>
                  
                    <tr> 
                      <td><table width='600' border=1 align='center' cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
                          <tr bgcolor='#B1DEDC' class="sbody"> 
                            <td width=13% align='center' bordercolor="#3A9D99"><font color=#000000>機構代碼</font></td>
                            <td width=42% align='center' bordercolor="#3A9D99"><font color=#000000>機構名稱</font></td>
                            <td width=10% align='center' bordercolor="#3A9D99"><font color=#000000>基準日</font></td>
                            <td width=13% align='center' bordercolor="#3A9D99"><font color=#000000>申報方式</font></td>
                            <td width=22% align='center' bordercolor="#3A9D99"><font color=#000000>檢核狀態</font></td>
                          </tr>
<%				
				if(dbData.size() == 0){%>
				 <tr bgcolor='#B1DEDC' class="sbody"> <td align=center colspan=3>無資料更新</td></tr>
				<%
				}
				
				int i = 0;
				String input_method = "";
				String upd_code="";
				String bgcolor="#FFFFE6";
				DataObject bean = null;
				while(i < dbData.size()){
				      bgcolor = (i % 2 == 0)?"#FFFFE6":"#F2F2F2";
				      bean = (DataObject)dbData.get(i);
				      if (((String)bean.getValue("input_method")) == null){
							input_method = "尚未申報";
					  }else if (((String)bean.getValue("input_method")).equals("F")){
							input_method = "檔案上傳";
					  }else if (((String)bean.getValue("input_method")).equals("W")){
							input_method = "線上編輯";
					  }else{
							input_method = "";
					  }		
					  upd_code = (String)bean.getValue("upd_code");
					  if (upd_code == null){
					      upd_code = "待檢核";
					  }else if (upd_code.equals("E")){
					      upd_code = "檢核有誤";
					  }else if (upd_code.equals("U")){
						  upd_code = "檢核成功";
					  }else if (upd_code.equals("F")){
					      upd_code = "上傳檔案不存在";
					  }else if (upd_code.equals("Z")){//94.03.25 add "Z"檢核為0
					      upd_code = "內容值皆為零"; //94.04.15   
					  }
					  	
%>			
                            <tr bgcolor='<%=bgcolor%>' class="sbody"> 
                            <td align='center' bordercolor="#3A9D99"><%=(String)bean.getValue("bank_code")%></td>
                            <td align='center' bordercolor="#3A9D99"><%=(String)bean.getValue("bank_name")%></td>
							<td align='center' bordercolor="#3A9D99"><%=(String)bean.getValue("mixdate")%></td>
							<td align='center' bordercolor="#3A9D99"><%=input_method%></td>                           
                            <td align='center' bordercolor="#3A9D99">                           
                            <% if (upd_code.equals("檢核有誤")){ %>                            
                            <a href='WMFileQuery_ShowError.jsp?Report_no=<%=Report_no%>&S_YEAR=<%=(bean.getValue("m_year")).toString()%>&S_MONTH=<%=(bean.getValue("m_month")).toString()%>&bank_code=<%=(String)bean.getValue("bank_code")%>'><font color=red><%=upd_code%></font></a><img src="images/icon_1.gif" width="12" height="12" align="absmiddle"></td>
                            <%}else{%>
                            <font color=black><%=upd_code%></font></td>
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
  
      </table></td>
  </tr>
</table>
</form>
</body>
</html>
