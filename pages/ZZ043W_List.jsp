<%
// 98.08.24 add 線上產生(A01~A04)Axx_operation中間值 by 2295
// 99.12.12 fix sqlInjection by 2808
//100.01.05 fix 日期顯示 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.ArrayList" %>
<%@ page import="java.util.ArrayList" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Calendar" %>
<%
	Calendar now = Calendar.getInstance();
   	String YEAR  = String.valueOf(now.get(Calendar.YEAR)-1911); //回覆值為西元年故需-1911取得民國年;
   	String MONTH = String.valueOf(now.get(Calendar.MONTH)+1);   //月份以0開始故加1取得實際月份;
    if(MONTH.equals("1")){//若本月為1月份是..則是申報上個年度的12月份
       YEAR = String.valueOf(Integer.parseInt(YEAR) - 1);
       MONTH = "12";	
    }else{    
      MONTH = String.valueOf(Integer.parseInt(MONTH) - 1);//申報上個月份的
    }
    
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");				
	String Report_no = ( request.getParameter("Report_no")==null ) ? "" : (String)request.getParameter("Report_no");					
	
	String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? YEAR : (String)request.getParameter("S_YEAR");				
	String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? MONTH : (String)request.getParameter("S_MONTH");				
	String E_YEAR = ( request.getParameter("E_YEAR")==null ) ? YEAR : (String)request.getParameter("E_YEAR");				
	String E_MONTH = ( request.getParameter("E_MONTH")==null ) ? MONTH : (String)request.getParameter("E_MONTH");				
	String lguser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		    
    String firstStatus = ( request.getParameter("firstStatus")==null ) ? "" : (String)request.getParameter("firstStatus");			    	
	System.out.println("ZZ043W_List.act="+act);	
	System.out.println("ZZ043W_List.S_YEAR="+S_YEAR);		    
	System.out.println("ZZ043W_List.S_MONTH="+S_MONTH);		    
	System.out.println("ZZ043W_List.E_YEAR="+E_YEAR);		    
	System.out.println("ZZ043W_List.E_MONTH="+E_MONTH);		    
	System.out.println("ZZ043W_List.firstStatus="+firstStatus);		    
    List reportList = (List)request.getAttribute("reportList");	
	String sqlCmd = "";
	List paramList =new ArrayList() ;
    String muser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		
	sqlCmd = " select a.report_no,c.cmuse_name "               
 		   + " from  WTT04_1D a,  CDShareNO b,  CDShareNO c "
 		   + " where a.MUSER_ID=? " 
		   + " and a.PROGRAM_ID='WMFileUpload'"
		   + " and a.DETAIL_TYPE='1'"//1-->Upload
		   + " and b.CmUSE_Div = '011'"
		   + " and (a.TRANSFER_TYPE = b.CmUSE_id and b.CmUSE_Div = '011') "  
           + " and  c.CmUSE_Div in ('012',  '013', '014') "
           + " and b.Identify_no = c.Identify_no "
		   + " and c.cmuse_name like a.report_no||'%' order by a.report_no";
	paramList.add(muser_id) ;	   
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
	//=======================================================================================================================
	
	//取得ZZ043W的權限
	Properties permission = ( session.getAttribute("ZZ043W")==null ) ? new Properties() : (Properties)session.getAttribute("ZZ043W"); 
	if(permission == null){
       System.out.println("ZZ043W_List.permission == null");
    }else{
       System.out.println("ZZ043W_List.permission.size ="+permission.size());               
    }	    
    
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<title>「線上產生彈性報表動態資料」</title>
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
function doSubmit(form,cnd){	   
	     if(!CheckQueryDate2(form.S_YEAR,form.S_MONTH,form.S_DATE,"起始日期")) return false;
         if(!CheckQueryDate2(form.E_YEAR,form.E_MONTH,form.E_DATE,"結束日期")) return false;     
	     if(trimString(form.S_DATE.value)!="" && trimString(form.E_DATE.value)!=""){
		    if(Math.abs(form.S_DATE.value) > Math.abs(form.E_DATE.value)){
    		   alert("起始查詢年月不可大於結束查詢年月");
    		   return false;
    	    }
    	 }   
    	 	  	     
	    form.action="/pages/ZZ043W.jsp?act="+cnd+"&test=nothing";	    
	    
	    form.submit();	    		    
}	
//-->
</script>
</head>

<body marginwidth="0" marginheight="0" leftmargin="0" topmargin="0" leftmargin="0">
<form method=post action='#'>
<input type="hidden" name="act" value="">  
<%if(reportList != null && reportList.size() != 0){%>
<input type="hidden" name="row" value="<%=reportList.size()+1%>">   
<%}%>
<table width="680" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  		<tr> 
   		 <td><img src="images/space_1.gif" width="12" height="12"></td>
  		</tr>

        <tr> 
          <td bgcolor="#FFFFFF">
		  <table width="680" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="680" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="190"><img src="images/banner_bg1.gif" width="190" height="17"></td>
                      <td width="300"><font color='#000000' size=4><b> 
                        <center>
                          「線上產生彈性報表動態資料」 
                        </center>
                        </b></font> </td>
                      <td width="190"><img src="images/banner_bg1.gif" width="190" height="17"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><img src="images/space_1.gif" width="12" height="12"></td>
              </tr>
              <tr> 
                <td><table width="680" border="0" align="center" cellpadding="0" cellspacing="0">
               
                    <tr> 
                      <div align="right"><jsp:include page="getLoginUser.jsp?width=680" flush="true" /></div> 
                    </tr>                    
                    <tr> 
                      <td><table width=680 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">                                                    
                          <tr class="sbody">						  
						  <td width='10%' align='left' bgcolor='#D8EFEE'>申報資料</td>
                          <td width='90%' bgcolor='e7e7e7'>	
                            <select name=Report_no>
							<%
							for (int i = 0; i < 4; i++) {//目前只有A01~A04
							%>
							<option value=<%=Report_List[i][0]%>  <%if(Report_no.equals(Report_List[i][0])) out.print("selected");%>>
								<%=Report_List[i][0]%>
								&nbsp;&nbsp;&nbsp;
								<%=Report_List[i][1]%>
							</option>
							<%}%>
						    </select>
                            &nbsp;&nbsp;&nbsp;&nbsp;                               
                            <input type="button" value="查詢清單" onclick="javascript:doSubmit(this.document.forms[0],'List');" name="QryList">                       
                            &nbsp; &nbsp; &nbsp; 
                            <input type="button" value="產生資料" onclick="javascript:doSubmit(this.document.forms[0],'Generate');" name="Generate">  &nbsp;  
                          </td>             
                          </tr>
                          
                          <tr class="sbody">                          
						  <td width='10%' align='left' bgcolor='#D8EFEE'>查詢年月</td>
                          <td width='90%' bgcolor='e7e7e7'>	
                            <input type='text' name='S_YEAR' value="<%=S_YEAR%>" size='3' maxlength='3' onblur='CheckYear(this)'>
        						<font color='#000000'>年
        						<select id="hide1" name=S_MONTH>
        						<option></option>
        						<%
        							for (int j = 1; j <= 12; j++) {
        							if (j < 10){%>        	
        							<option value=0<%=j%><%if(String.valueOf(Integer.parseInt(S_MONTH)).equals(String.valueOf(j))){out.print(" selected");}%>>0<%=j%></option>        		
            						<%}else{%>
            						<option value=<%=j%><%if(String.valueOf(Integer.parseInt(S_MONTH)).equals(String.valueOf(j))){out.print(" selected");}%>><%=j%></option>
            						<%}%>
        						<%}%>
        						</select><font color='#000000'>月</font>&nbsp;&nbsp;至
        						<input type='text' name='E_YEAR' value="<%=E_YEAR%>" size='3' maxlength='3' onblur='CheckYear(this)'>
        						<font color='#000000'>年
        						<select id="hide1" name=E_MONTH>
        						<option></option>
        						<%
        							for (int j = 1; j <= 12; j++) {
        							if (j < 10){%>        	
        							<option value=0<%=j%><%if(String.valueOf(Integer.parseInt(E_MONTH)).equals(String.valueOf(j))){out.print(" selected");}%>>0<%=j%></option>        		
            						<%}else{%>
            						<option value=<%=j%><%if(String.valueOf(Integer.parseInt(E_MONTH)).equals(String.valueOf(j))){out.print(" selected");}%>><%=j%></option>
            						<%}%>
        						<%}%>
        						</select><font color='#000000'>月</font>
        						 <input type=hidden name=S_DATE value=''>
        						 <input type=hidden name=E_DATE value=''>
                          </td>                                   
                          </tr>
                                                              		      					      
                          </table>      
                      </td>    
                      </tr>
                      
                      
                      <tr> 
                      <td><table width=680 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">                                              	                                              
                      <%                     
                      int i = 0;      
                      String bgcolor="#D3EBE0";                          
                      if(reportList != null){ %>
                       	  <tr class="sbody" bgcolor="#9AD3D0">                       	  
                            <td width="30">序號</td> 
                            <td width="190">彈性報表名稱</td>
            				<td width="149">資料類別</td>            				
            				<td width="60">申報年月</td>       
            				<td width="60">資料筆數</td>    
            				<td width="60">建立者</td>         				
            				<td width="80">產生日期</td>    				            				
					      </tr>   
                   		    <% if(reportList.size() == 0){%>
                   			   <tr class="sbody" bgcolor="<%=bgcolor%>">
                   			   <td colspan=7 align=center>無資料可供查詢</td><tr>
                   			   <tr>                   			   
                   			<% }else{ 
                   			    String update_date="";
                   			    String yyymm="";
                   			    DataObject bean = null;
                   			    while(i < reportList.size()){                     		      
                    		      bgcolor = (i % 2 == 0)?"#e7e7e7":"#D3EBE0";
                    		      bean = (DataObject)reportList.get(i);
                    		      if((String)bean.getValue("s_updatedate") != null){
                    		         update_date = (String)bean.getValue("s_updatedate");
                    		         //100.01.05 fix 日期顯示 by 2295
                    		         if(update_date.startsWith("1")){
                                		update_date = update_date.substring(0,3)+"/"+update_date.substring(3,5)+"/"+update_date.substring(5,7)+ " " + update_date.substring(8,update_date.length());	                    		
                                	 }else{
                                		update_date = update_date.substring(0,2)+"/"+update_date.substring(2,4)+"/"+update_date.substring(4,6)+ " " + update_date.substring(7,update_date.length());	                    		
                                	 }     
                    		         System.out.println("update_date"+update_date);
                    		      }
                    		      if( bean.getValue("s_yymm") != null){
                    		         yyymm = (String)bean.getValue("s_yymm");
                    		         if(yyymm.length() == 4){
		            					 yyymm = "0"+yyymm;
		          					 } 
                    		      }
                    		%>                         	       
                    		<tr class="sbody" bgcolor="<%=bgcolor%>">                       	  
                               <td width="30"><%=i+1%></td>                                  
                               <td width="190">
                               <%if( bean.getValue("s_report_name") != null ) out.print((String)bean.getValue("s_report_name")); else out.print("&nbsp;");%>            				   
                               </td>
            				   <td width="149">
            				   <%if( bean.getValue("kink_type") != null ) out.print((String)bean.getValue("kink_type")); else out.print("&nbsp;");%>            				   
            				   </td>            				
            				   <td width="60">
            				   <%if( bean.getValue("s_yymm") != null ) out.print((String)bean.getValue("s_yymm")); else out.print("&nbsp;");%>
            				   </td>
            				   <td width="60">
            				   <%if( bean.getValue("total") != null ) out.print(bean.getValue("total").toString()); else out.print("&nbsp;");%>
            				   </td>
            				   <td width="60">
            				   <%if( bean.getValue("user_id") != null ) out.print((String)bean.getValue("user_id")); else out.print("&nbsp;");%>
            				   </td>
            				   <td width="80">            				  
            				   <%if(!update_date.equals("")) out.print(update_date); else out.print("&nbsp;");%>
            				   </td>            				            				
					        </tr>   
                   			<%  i++; 
                   			    }//end of while
                   			   }//end of if%>
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
