<%
//94.10.21 create by 2495
//99.12.12 fix sqlInjection by 2808
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Enumeration" %>
<%@ page import="java.util.HashMap" %>
<%@ page import="java.util.Hashtable" %>
<%@ page import="java.util.StringTokenizer" %>

<%
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");
	String seq_no = ( request.getParameter("seq_no")==null ) ? "" : (String)request.getParameter("seq_no");															 
	String title = (act.equals("new"))?"首頁公告區新增標題":"首頁公告區維護";
  String head_mark ="";
 	String muser_id=" ",pre_muser_id=" ";
 	String muser_name=" ",pre_muser_name=" ",pre_m_telno=" ",notify_url="",pre_m_email=" ",pre_notify_date=" ",pre_notify_end_date=" ",pre_append_file=" ",pre_start_year="",pre_start_month=" ",pre_start_day=" ",pre_start_hour=" ",pre_end_year=" ",pre_end_month=" ",pre_end_day=" ";
 	String m_telno=" ";
 	String m_email=" ";
 	String m_systime=" ";

 	
	List WLX_Notify = (List)request.getAttribute("WLX_Notify");		
	if(WLX_Notify == null || act.equals("new")){
	   System.out.println("WLX_Notify == null");
	   head_mark ="";
	   muser_id = (String)session.getAttribute("muser_id");
	   List User_Data = TakeUserData(muser_id);	   	   
     muser_name = (String)((DataObject)User_Data.get(0)).getValue("user_name");	
	   m_telno = (String)((DataObject)User_Data.get(0)).getValue("m_telno");
	   m_email = (String)((DataObject)User_Data.get(0)).getValue("m_email");
	}
	else if(WLX_Notify != null ||act.equals("Edit")){
	   System.out.println("WLX_Notify.size()="+WLX_Notify.size());
	   head_mark = (String)((DataObject)WLX_Notify.get(0)).getValue("headmark");
	   notify_url = (((DataObject)WLX_Notify.get(0)).getValue("notify_url")==null ) ? "" : (String)((DataObject)WLX_Notify.get(0)).getValue("notify_url");		   
	   pre_notify_date = (String)((DataObject)WLX_Notify.get(0)).getValue("notify_date");
	   pre_notify_end_date = (String)((DataObject)WLX_Notify.get(0)).getValue("notify_end_date");
	   pre_muser_id = (String)((DataObject)WLX_Notify.get(0)).getValue("user_id");
	   pre_muser_name = (String)((DataObject)WLX_Notify.get(0)).getValue("user_name");
	   pre_append_file = (((DataObject)WLX_Notify.get(0)).getValue("append_file")==null ) ? "null" : (String)((DataObject)WLX_Notify.get(0)).getValue("append_file");		   
	   List User_Data = TakeUserData(pre_muser_id);
	   pre_m_telno = (String)((DataObject)User_Data.get(0)).getValue("m_telno");
	   pre_m_email = (String)((DataObject)User_Data.get(0)).getValue("m_email");
	    System.out.println("pre_notify_date="+pre_notify_date);
	   pre_start_year = pre_notify_date.substring(6,10);	    
	   int iStartYear = Integer.parseInt(pre_start_year)-1911;
	   System.out.println("iStartYear="+iStartYear);
		 pre_start_year = Integer.toString(iStartYear);
	   pre_start_month = pre_notify_date.substring(0,2); 
	   	 System.out.println("pre_start_month="+pre_start_month);
	   pre_start_day = pre_notify_date.substring(3,5); 
	   	 System.out.println("pre_start_day="+pre_start_day);
	   pre_start_hour = pre_notify_date.substring(11,13);
	     System.out.println("pre_start_hour="+pre_start_hour);	 	   
	   
	   
	    System.out.println("pre_notify_end_date="+pre_notify_date);
	   pre_end_year = pre_notify_end_date.substring(6,10);	    
	   int iEndYear = Integer.parseInt(pre_end_year)-1911;
	   System.out.println("iEndYear="+iEndYear);
	   pre_end_year = Integer.toString(iEndYear);
	   pre_end_month = pre_notify_end_date.substring(0,2); 
	   	 System.out.println("pre_end_month="+pre_end_month);
	   pre_end_day = pre_notify_end_date.substring(3,5); 
	   	 System.out.println("pre_end_day="+pre_end_day);
	   
	   
	   
	   
	   
	   muser_id = (String)session.getAttribute("muser_id");
	   User_Data = TakeUserData(muser_id);	   	   
     muser_name = (String)((DataObject)User_Data.get(0)).getValue("user_name");	
	   m_telno = (String)((DataObject)User_Data.get(0)).getValue("m_telno");
	   m_email = (String)((DataObject)User_Data.get(0)).getValue("m_email");
	   Calendar calendar=new GregorianCalendar();	   
		 int year=calendar.get(Calendar.YEAR)-1911;		 
		 int month=calendar.get(Calendar.MONTH);
     int day=calendar.get(Calendar.DATE);     
	   String myear = Integer.toString(year);		
	   String mmonth = Integer.toString(month);	
	   String mday = Integer.toString(day);
	   m_systime = myear+"//"+month+"//"+mday;
	    	     
	}
	
	//取得WLX_Notify的權限
	Properties permission = ( session.getAttribute("WLX_Notify")==null ) ? new Properties() : (Properties)session.getAttribute("WLX_Notify"); 
	if(permission == null){
       System.out.println("WLX_Notify_List.permission == null");
    }else{
       System.out.println("WLX_Notify_List.permission.size ="+permission.size());               
    }
    
%>
<%!
	  //取得該user_id之相關資料
    private List TakeUserData(String muser_id){
    		//查詢條件    
    		List paramList =new ArrayList();
    		String sqlCmd = "select * from MUSER_DATA where  muser_id =? ";
    		paramList.add(muser_id) ;
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
        return dbData;
    }
    
%>


<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/ZZ091W.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<title>總機構高階主管基本資料維護</title>
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
<style type="text/css">
<!--
.style1 {color: #666666}
-->
</style>
</head>

<body marginwidth="0" marginheight="0" leftmargin="0" topmargin="0" leftmargin="0">
<form method=post action='#' enctype="multipart/form-data">
<table width="640" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  		  <tr> 
   		   <td><img src="images/space_1.gif" width="12" height="12"></td>
  		  </tr>
          <td bgcolor="#FFFFFF">
		  <table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="170"><img src="images/banner_bg1.gif" width="170" height="17"></td>
                      <td width="250"><font color='#000000' size=4><b> 
                        <center><%=title%></center>
                        </b></font> </td>
                      <td width="170"><img src="images/banner_bg1.gif" width="170" height="17"></td>
                    </tr>  
                   
                  </table></td>
              </tr>
              <tr> 
                <td><img src="images/space_1.gif" width="12" height="12"></td>
              </tr>
              <tr> 
                
                    <tr> 
                      <div align="right"><jsp:include page="getLoginUser.jsp" flush="true" /></div> 
                    </tr>
                    
                    </tr>
                    <tr> 
                      <td><table width=600 border=1 align=center cellpadding="1" cellspacing="0" bordercolor="#3A9D99" height="208">                          						
                        <tr>
						<td width='138' align='left' bgcolor='#D8EFEE' height="1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;    
                          <div align="center"><font size="2">公告標題</font><font color='red' size=4>*</font></div></td>
						<td width='452' bgcolor='e7e7e7' height="1">
                                                <input type='text' name='head_title' value="<%=head_mark%>" size='68' maxlength='40' ></td>                               		
                        </tr>
	  					
                        <tr>
						<td width='138' align='left' bgcolor='#D8EFEE' height="8">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;    
                          <div align="center"><font size="2">公告起始時間</font><font color='red' size=4>*</font></div></td>
						<td width='452' height="8" bgcolor='e7e7e7' class="sbody">
						<%if(act.equals("new"))
						{ 
   						 Calendar calendar=new GregorianCalendar();
  						 int current_year= calendar.get(Calendar.YEAR);	
  						 current_year=current_year-1911;						 
   						 String current_month=Integer.toString(calendar.get(Calendar.MONTH)+1);							 
   						 System.out.println("current_month="+current_month);					 
   						 String current_day=Integer.toString(calendar.get(Calendar.DATE));
						 System.out.println("current_day="+current_day);						 
					         String current_hour=Integer.toString(calendar.get(Calendar.HOUR_OF_DAY));
						 System.out.println("current_hour="+current_hour);								 				 						 
						%>
						<input name="StartYear" type="text" value="<%=current_year%>" size="3" maxlength="3">年		                          
		                          
		        <select name=StartMonth>
   								<option></option>
   					   <%
     						for(int j = 1; j <= 12; j++) 
    						{
      						if (j < 10)
        					{%>        	
      	   	 				<option value=0<%=j%> <%if(current_month.equals(String.valueOf(j))) out.print("selected");%>>0<%=j%></option>        		
      		 				<%}
        	 				else
        					{%>
       	    				<option value=<%=j%> <%if(current_month.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
        					<%}%>
  						<%}%>
				    </select></font><font color='#000000'>月
				    	<select name=StartDay>
   								<option></option>
   					   <%
     						for(int j = 1; j <= 31; j++) 
    						{
      						if (j < 10)
        					{%>        	
      	   	 				<option value=0<%=j%> <%if(current_day.equals(String.valueOf(j))) out.print("selected");%>>0<%=j%></option>        		
      		 				<%}
        	 				else
        					{%>
       	    				<option value=<%=j%> <%if(current_day.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
        					<%}%>
  						<%}%>
				    </select></font><font color='#000000'>日
		                          
		                          
		                          <select name=StartHour>
   								<option></option>
   					   <%
     						for(int j = 1; j <= 24; j++) 
    						{
      						if (j < 10)
        					{%>        	
      	   	 				<option value=0<%=j%> <%if(current_hour.equals(String.valueOf(j))) out.print("selected");%>>0<%=j%></option>        		
      		 				<%}
        	 				else
        					{%>
       	    				<option value=<%=j%> <%if(current_hour.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
        					<%}%>
  						<%}%>
				    </select></font><font color='#000000'>時
		                          
														</td>

                        </tr>
	  				<%}else if(act.equals("Edit")){%>
	  				<input name="StartYear" type="text" value="<%=pre_start_year%>" size="3" maxlength="3">年		                          
		                          
		        <select name=StartMonth>
   						<option></option>
   					   <%
     						for(int j = 1; j <= 12; j++) 
    						{
      						if (j < 10)
        					{%>        	
      	   	 				<option value=0<%=j%> <%if(pre_start_month.equals(String.valueOf("0"+j))) out.print("selected");%>>0<%=j%></option>        		
      		 				<%}
        	 				else
        					{%>
       	    				<option value=<%=j%> <%if(pre_start_month.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
        					<%}%>
  						<%}%>				    
				    </select></font><font color='#000000'>月
				    	<select name=StartDay>
   								<option></option>
   					   <%
     						for(int j = 1; j <= 31; j++) 
    						{
      						if (j < 10)
        					{%>        	
      	   	 				<option value=0<%=j%> <%if(pre_start_day.equals(String.valueOf("0"+j))) out.print("selected");%>>0<%=j%></option>        		
      		 				<%}
        	 				else
        					{%>
       	    				<option value=<%=j%> <%if(pre_start_day.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
        					<%}%>
  						<%}%>
				    </select></font><font color='#000000'>日
				    	<select name=StartHour>
   								<option></option>
   					   <%
     						for(int j = 1; j <= 24; j++) 
    						{
      						if (j < 10)
        					{%>        	
      	   	 				<option value=0<%=j%> <%if(pre_start_hour.equals(String.valueOf(j))) out.print("selected");%>>0<%=j%></option>        		
      		 				<%}
        	 				else
        					{%>
       	    				<option value=<%=j%> <%if(pre_start_hour.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
        					<%}%>
  						<%}%>
				    </select></font><font color='#000000'>時	
	  					
	  				<%}%>	  				 
	  				
							
	  					<tr class="sbody">
						<td width='138' align='left' bgcolor='#D8EFEE' height="1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;    
                          <div align="center"><font size="2">公告結束日期</font><font color='red' size=4>*</font></div></td>
						<td width='452' bgcolor='e7e7e7' height="1">
							<%if(act.equals("new")){%>
							                  <input name="EndYear" type="text" value="95" size="3" maxlength="3">
																年
																<select name="EndMonth">
		                            <option value='01'>01
		                            <option value='02'>02
		                            <option value='03'>03
		                            <option value='04'>04
		                            <option value='05'>05
		                            <option value='06'>06
		                            <option value='07'>07
		                            <option value='08'>08
		                            <option value='09'>09
		                            <option value='10'>10
		                            <option value='11'>11
		                            <option selected value='12'>12	
					   </select>
					   月
					    <select name="EndDay">
					    <option value='01'>01
		                            <option value='02'>02
		                            <option value='03'>03
		                            <option value='04'>04
		                            <option value='05'>05
		                            <option value='06'>06
		                            <option value='07'>07
		                            <option value='08'>08
		                            <option value='09'>09
		                            <option value='10'>10
		                            <option value='11'>11
		                            <option value='12'>12
		                            <option value='13'>13
		                            <option value='14'>14
		                            <option value='15'>15
		                            <option value='16'>16
		                            <option value='17'>17
		                            <option value='18'>18
		                            <option value='19'>19
		                            <option value='20'>20
		                            <option value='21'>21
		                            <option value='22'>22
		                            <option value='23'>23
		                            <option value='24'>24
		                            <option value='25'>25
		                            <option value='26'>26
		                            <option value='27'>27
		                            <option value='28'>28
		                            <option value='29'>29
		                            <option selected value='30'>30
		                            <option value='31'>31
																</select>
																日</td>             
                        </tr>
             <%}else if(act.equals("Edit")){%>
             	  				<input name="EndYear" type="text" value="<%=pre_end_year%>" size="3" maxlength="3">年		                          
		                          
		        <select name=EndMonth>
   						<option></option>
   					   <%
     						for(int j = 1; j <= 12; j++) 
    						{
      						if (j < 10)
        					{%>        	
      	   	 				<option value=0<%=j%> <%if(pre_end_month.equals(String.valueOf("0"+j))) out.print("selected");%>>0<%=j%></option>        		
      		 				<%}
        	 				else
        					{%>
       	    				<option value=<%=j%> <%if(pre_end_month.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
        					<%}%>
  						<%}%>				    
				    </select></font><font color='#000000'>月
				    	<select name=EndDay>
   								<option></option>
   					   <%
     						for(int j = 1; j <= 31; j++) 
    						{
      						if (j < 10)
        					{%>        	
      	   	 				<option value=0<%=j%> <%if(pre_end_day.equals(String.valueOf("0"+j))) out.print("selected");%>>0<%=j%></option>        		
      		 				<%}
        	 				else
        					{%>
       	    				<option value=<%=j%> <%if(pre_end_day.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
        					<%}%>
  						<%}%>
				    </select></font><font color='#000000'>日
	  				<%}%>	  				           	
             <%if(act.equals("new")){%>
			    <tr class="sbody">
						<td width='138' align='left' bgcolor='#D8EFEE' height="26">
                            
                          <div align="center">
                            <font size="2">附加網址</font>                           
                          </div></td>
						<td width='452' bgcolor='e7e7e7' height="26">
														<p> </p> 
                            <p>                           
                              <input type="text" name="notifyUrl" value="" SIZE=68>
                              </font><font color='#000000'>                                                                                         
                              </font></p>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              </font></p>
            </td>
            </tr>		
             	
	  					<tr class="sbody">
						<td width='138' align='left' bgcolor='#D8EFEE' height="26">
                            
                          <div align="center">
                            <font size="2">附加檔案</font>                           
                          </div></td>
						<td width='452' bgcolor='e7e7e7' height="26">
														<p> </p> 
                            <p>                           
                              <input type="file" name="FileType3" value="瀏覽" SIZE=50>
                              </font><font color='#000000'>                                                                                         
                              </font></p>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              </font></p>
            </td>
            </tr>



            <%}%>
            <%if(act.equals("Edit")){%> 
		    <tr class="sbody">
		    <td width='138' align='left' bgcolor='#D8EFEE' height="26">        
                          <div align="center">
                            <font size="2">附加網址</font>                           
                          </div></td>
			<td width='452' bgcolor='e7e7e7' height="26">
			<p> </p> 
                            <p>                           
                              <input type="text" name="notifyUrl" value="<%=notify_url%>" SIZE=68>
                              </font><font color='#000000'>                                                                                         
                              </font></p>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              </font></p>
            </td>
            </tr>
                       	
	  					<tr class="sbody">
						<td width='138' align='left' bgcolor='#D8EFEE' height="26">
                            
                          <div align="center">
                            <font size="2">附加檔案</font>                           
                          </div></td>
						<td width='452' bgcolor='e7e7e7' height="26">
														<p> </p>
														<%if(!pre_append_file.equals(" ")){%> 										 
                            <p>                           
                              <input type="file" name="FileType3" value="瀏覽" SIZE=50  onClick="if(true){ this.form.FileType3.disabled = true} alert('若是要上傳新檔案,請先刪除原始檔案');">
                              </font><font color='#000000'>                                                                                         
                              </font>
                            </p> 
                            <%}else{%>
                            <p>                           
                              <input type="file" name="FileType3" value="瀏覽" SIZE=50>
                              </font><font color='#000000'>                                                                                         
                              </font>
                            </p>
                            <%}%>                              
            </td>                                                                                                                                                                                                                                                                                                                                                                                                         </font></p>            
            </tr>
            <%}%>
            
            <%if(act.equals("Edit")){%>                       
            <tr class="sbody">
						<td width='138' align='left' bgcolor='#D8EFEE' height="26">
                          
                          <div align="center">
                            <font size="2">已上傳的檔案</font>                           
                          </div></td>
                     
						<td width='452' bgcolor='e7e7e7' height="26">
														<p> </p> 
						<%if(!pre_append_file.equals(" ")){%> 								
                            <p>                           
                              <%=pre_append_file%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" name="ResetLoad" value="刪除"  onClick="javascript:doSubmit(this.document.forms[0],'Clear','<%=seq_no%>');">&nbsp;&nbsp;&nbsp;<input type="submit" name="ResetLoad" value="開啟檔案"  onClick="javascript:doSubmit(this.document.forms[0],'Open','<%=seq_no%>');">
                              </font><font color='#000000'>                                                                                         
                              </font></p>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              </font></p>
            <%}%>
            </td>           
            </tr>
            <%}%>		
                        	
                        </Table></td>
                    </tr>                 
                    <tr>                  
                <td><div align="right">

			  <tr>                  
                <td><div align="right"><jsp:include page="getMaintainUser.jsp" flush="true" /></div></td>              
              </tr>
                    </div></td>                                              
              </tr>
              
              <tr> 
                <td>&nbsp;</td>
              </tr>
              <tr> 
                <td><div align="center"> 
                    <table width="243" border="0" cellpadding="1" cellspacing="1">
                      <tr>
                      	<%if(act.equals("new")){%>							                      	        	                                   		     
				                <td width="66"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'insert','<%=head_mark%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_confirmb.gif',1)"><img src="images/bt_confirm.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a></div></td>			      	
                        <%}%>
                      	<%if(act.equals("Edit")){%>							                      	        	                                   		     
				                <td width="66"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Update','<%=seq_no%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_updateb.gif',1)"><img src="images/bt_update.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a></div></td>			      	
                        <%}%>
                        <td width="93"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'Back','');"  onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image106','','images/bt_backb.gif',1)"><img src="images/bt_back.gif" name="Image106" width="80" height="25" border="0" id="Image106"></a></div></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>   
              
  <tr> 
                <td><table width="600" border="0" cellpadding="1" cellspacing="1" class="sbody">
                    <tr> 
                      <td colspan="2"><font color='#990000'><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle"><font color="#007D7D" size="3">使用說明      
                        : </font></font></td>
                    </tr>
                    <tr> 
                      <td width="16">&nbsp;</td>
                      <td width="577"> <ul>
			  <li>公告標題以40個字以內為上限。</li>
                      	  <li>確認輸入資料無誤後, 按<font color="#666666">【確定】</font>即將本表上的資料, 於資料庫中建檔。</li>     
                      	  <li><font color="#666666">【</font><span class="style1">瀏覽</span><font color="#666666">】</font>-選取鍵入者欲附加之檔案,&nbsp;</li>                              
        				  				<li>如放棄修改或無修改之資料需輸入, 按<font color="#666666">【回上一頁】</font>]即離開本程式。</li>                              
                          <li>上傳檔案型態可為<font color="#666666">「Word檔;Excel檔;Zip壓縮檔;PowerPoint檔」。</li>  
        				  				<li>【<font color="red" size=4>*</font>】</font>為必填欄位。</li>
        				  				<li>「公告結束日期」係管制當登入系統網址的日期超出本欄時,首頁不會顯示本公告標題。</li>
                      </ul></td>
                    </tr>
                  </table></td>
              </tr>
              <!--tr> 
                <td><div align="center"><img src="images/line_1.gif" width="600" height="12"></div></td>
              </tr-->
</table>
</form>
</body>
