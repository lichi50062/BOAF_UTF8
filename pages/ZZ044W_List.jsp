<%
//99.12.17 add 金庫BIS資料轉檔作業 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%
	String YEAR  = Utility.getYear();
   	String MONTH = Utility.getMonth();   	
    
    Map dataMap =Utility.saveSearchParameter(request);
	String report_no = "ZZ044W";
	String act = Utility.getTrimString(dataMap.get("act"));		
	String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? YEAR : (String)request.getParameter("S_YEAR");				
	String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? MONTH : (String)request.getParameter("S_MONTH");				
	String E_YEAR = ( request.getParameter("E_YEAR")==null ) ? YEAR : (String)request.getParameter("E_YEAR");				
	String E_MONTH = ( request.getParameter("E_MONTH")==null ) ? MONTH : (String)request.getParameter("E_MONTH");				
	String lguser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		    
	String firstStatus = Utility.getTrimString(dataMap.get("firstStatus"));		
	String acc_tr_type = Utility.getTrimString(dataMap.get("acc_tr_type"));		
    System.out.println(report_no+"_List.act="+act);	
	System.out.println(report_no+"_List.S_YEAR="+S_YEAR);		    
	System.out.println(report_no+"_List.S_MONTH="+S_MONTH);		    
	System.out.println(report_no+"_List.E_YEAR="+E_YEAR);		    
	System.out.println(report_no+"_List.E_MONTH="+E_MONTH);		    
	System.out.println(report_no+"_List.firstStatus="+firstStatus);		    
    List reportList = (List)request.getAttribute("reportList");	
    List acc_tr_type_List = (List)request.getAttribute("BISList");
	DataObject bean = null;
    String muser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");
      		
	//=======================================================================================================================	 
    
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<title>「金庫BIS報表轉檔作業」</title>
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
    	if(cnd=='parseAgriBankRpt' && form.acc_tr_type.value == 'ALL'){
    	   alert("轉檔程式只能挑選單一報表名稱!!");
    	   return false;
    	}	
    	if(cnd=='parseALLAgriBankRpt' && form.acc_tr_type.value != 'ALL'){
    	   alert("全部轉檔程式只能挑選全部報表名稱!!");
    	   return false;
    	}  	     
	    form.action="/pages/ZZ044W.jsp?act="+cnd+"&test=nothing";	    
	    
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
                          「金庫BIS報表轉檔作業」 
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
                     <%
                      String nameColor="nameColor_sbody";
                      String textColor="textColor_sbody";
                      String bordercolor="#3A9D99";
                     %>   
                    <tr> 
                      <div align="right"><jsp:include page="getLoginUser.jsp?width=680" flush="true" /></div> 
                    </tr>                    
                    <tr> 
                      <td><table width=680 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="<%=bordercolor%>">                                                    
                          <tr class="sbody">						  
						  <td width='10%' align='left' class="<%=nameColor%>">報表名稱</td>
                          <td width='90%' class="<%=textColor%>">	
                            <select name=acc_tr_type>
                            <option value='ALL'>全部</option>
							<%
							for (int i = 0; i < acc_tr_type_List.size(); i++) {
							     bean = (DataObject)acc_tr_type_List.get(i);							    
							%>
							<option value=<%=(String)bean.getValue("identify_no")%>  <%if(acc_tr_type.equals((String)bean.getValue("identify_no"))) out.print("selected");%>>
								<%=(String)bean.getValue("cmuse_name")%>								
							</option>
							<%}%>
						    </select>
                            &nbsp;                               
                            <input type="button" value="查詢" onclick="javascript:doSubmit(this.document.forms[0],'List');" name="QryList">&nbsp;                                                  
                            <input type="button" value="轉檔" onclick="javascript:doSubmit(this.document.forms[0],'parseAgriBankRpt');" name="parseAgriBankRpt">  &nbsp;                              
                            <input type="button" value="全部轉檔" onclick="javascript:doSubmit(this.document.forms[0],'parseALLAgriBankRpt');" name="parseAgriBankRpt">  &nbsp;  
                          </td>             
                          </tr>
                          
                          <tr class="sbody">                          
						  <td width='10%' align='left' class="<%=nameColor%>">查詢年月</td>
                          <td width='90%' class="<%=textColor%>">	
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
                      <td><table width=680 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="<%=bordercolor%>">                                              	                                              
                      <%                     
                      int i = 0;      
                      String bgcolor="#D3EBE0"; 
                      String listTitle="listTitleColor_sbody"; 
                      String list1Color="list1Color_sbody";
                      String list2Color="list2Color_sbody";
                      String listColor="list1Color_sbody";                          
                      if(reportList != null){ %>
                       	  <tr class="<%=listTitle%>">                       	  
                            <td width="30">序號</td> 
                            <td width="400">報表名稱</td>            				
            				<td width="60">申報年月</td>       
            				<td width="60">資料筆數</td>                								            				
					      </tr>   
                   		    <% if(reportList.size() == 0){%>
                   			   <tr class="<%=list1Color%>">
                   			   <td colspan=4 align=center>無資料可供查詢</td><tr>
                   			   <tr>                   			   
                   			<% }else{                    			    
                   			    while(i < reportList.size()){                     		                          		      
                    		      listColor = (i % 2 == 0)?list2Color:list1Color;	      
                    		      bean = (DataObject)reportList.get(i);                    		   
                    		%>                         	       
                    		<tr class="<%=listColor%>">                       	  
                               <td width="30"><%=i+1%></td>                                  
                               <td width="400">
                               <%if( bean.getValue("s_report_name") != null ) out.print((String)bean.getValue("s_report_name")); else out.print("&nbsp;");%>            				   
                               </td>            				            				
            				   <td width="60">
            				   <%if( bean.getValue("s_yymm") != null ) out.print((String)bean.getValue("s_yymm")); else out.print("&nbsp;");%>
            				   </td>
            				   <td width="60">
            				   <%if( bean.getValue("total") != null ) out.print(bean.getValue("total").toString()); else out.print("&nbsp;");%>
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
