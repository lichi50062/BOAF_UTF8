<%
//93.12.17 add 權限檢核 by 2295
//92.12.22 fix 取得可檢核檔案的報表名稱根據傳輸類別 by 2295
//94.02.04 add 預設年月為上個月份,若本月為1月份時.則是申報上個年度的12月份
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
   
	List bn01Data = (List)request.getAttribute("bn01Data");	
	String bank_name = ( request.getAttribute("bank_name")==null ) ? "xxx農業信用部" : (String)request.getAttribute("bank_name");		

	
	//fix 93.12.20 若有已點選的bank_type,則以已點選的bank_type為主============================================================
	String bank_type = ( request.getParameter("bank_type")==null ) ? "" : (String)request.getParameter("bank_type");				   
	//bank_type = ( session.getAttribute("nowbank_type")==null ) ? bank_type : (String)session.getAttribute("nowbank_type");			
	//=======================================================================================================================	
	String transfer_type="1";
	if(bank_type.equals("6") || bank_type.equals("7")){//農會.漁會
	   transfer_type="1";
	}
	
	if(bank_type.equals("4")){//農業信用保証基金
	   transfer_type="2";
	}
	
	if(bank_type.equals("2")){//農業金融局
	   transfer_type="3";
	}	
	//92.12.20取得可檢核檔案的報表名稱===============================================================================================
	String sqlCmd = "";
	String muser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		
	sqlCmd = "select c.cmuse_name " 
		   + " from  CDShareNO b,  CDShareNO c ";
    if(bank_type.equals("Z")){//全部類別		   
       sqlCmd += " where b.CmUSE_Div = '011' "; 
    }else{
	   sqlCmd = sqlCmd + " where b.CmUSE_id = '"+transfer_type+"'"
	   		           + " and b.CmUSE_Div = '011' "; 	   
	}	   
	sqlCmd = sqlCmd 
	       + " and  c.CmUSE_Div in ('012',  '013', '014') "
		   + " and b.Identify_no = c.Identify_no " 
 		   + " order by cmuse_name";
 		   
    List dbData = DBManager.QueryDB_SQLParam(sqlCmd,null,"");		   
    String report_data="";
    String[][] Report_List = new String[dbData.size()][2];
    if(dbData != null && dbData.size() != 0){		   
       for(int i=0;i<dbData.size();i++){
           report_data = (String)((DataObject)dbData.get(i)).getValue("cmuse_name");
           if(report_data.indexOf("_") != -1){
              Report_List[i][0]=report_data.substring(0,report_data.indexOf("_"));
	          Report_List[i][1]=report_data.substring(report_data.indexOf("_")+1,report_data.length());
	       }           
       }
	}
	
	
	//取得WMFileCheck的權限
	Properties permission = ( session.getAttribute("WMFileCheck")==null ) ? new Properties() : (Properties)session.getAttribute("WMFileCheck"); 
	if(permission == null){
       System.out.println("WMFileCheck_Qry.permission == null");
    }else{
       System.out.println("WMFileCheck_Qry.permission.size ="+permission.size());
               
    }	
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/WMFileCheck.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<title>人工檢核轉檔作業</title>
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
                          人工檢核轉檔作業 
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
						  <td align='left' bgcolor='#D8EFEE'>金融機構類別</td>
						  <td colspan=2 bgcolor='e7e7e7'><%=bank_name%></td>
	  					</tr>                 
                        <tr class="sbody">
						  <td align='left' bgcolor='#D8EFEE'>申報資料</td>
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
						    </td>
	  					</tr>
	  					
						<tr class="sbody">
						  <td align='left' bgcolor='#D8EFEE'>金融機構名稱</td>
						  <td align='left' colspan=2 bgcolor='e7e7e7'><select name=bank_code>		
						  <option value="all">全部</option>				  
							<%
							int i = 0 ;							
							if(bn01Data != null && bn01Data.size() > 0){							
								while(i<bn01Data.size()){													
							%>
							<option value=<%=(String)((DataObject)bn01Data.get(i)).getValue("bank_no")%>>
								<%=((String)((DataObject)bn01Data.get(i)).getValue("bank_no")).trim()%>
								&nbsp;
								<%=((String)((DataObject)bn01Data.get(i)).getValue("bank_name")).trim()%>
							</option>												
							<%  i ++;
							   }
							}%>															
						    </td>
	  					</tr>
	  					<tr class="sbody">
							<td align='left' bgcolor='#D8EFEE'>基準日</td>							
							<td colspan=2 bgcolor='e7e7e7'>
							<input type='radio' name='YM' value='0' checked onclick='checkYM(this.document.forms[0])'>							
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
        					<input type='radio' name='YM' value='1' onclick='checkYM(this.document.forms[0])'>全部	
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
                      <%if(permission != null && permission.get("U") != null && permission.get("U").equals("Y")){ %>                   	        	                                   		     
                        <td width="66"> <div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Check','<%=bank_type%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_execb.gif',1)"><img src="images/bt_exec.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a></div></td>
                        <td width="93"><div align="center"><a href="javascript:history.back();"  onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image81','','images/bt_backb.gif',1)"><img src="images/bt_back.gif" name="Image81" width="80" height="25" border="0" id="Image81"></a></div></td>                        
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
                      <li>本網頁提供人工檢核轉檔作業的功能。</li>
					  <li>點選執行開始更新作業。</li>
					  <li>請盡量縮小條件的範圍。</li>
	    			  <li>此更新作業須較長的時間,<font color=red>請勿重複點選[執行]。</font></li> 					                    
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
<br><br>
</body>
</html>
