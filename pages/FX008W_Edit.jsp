<%
//94.10.23 standard edit and insert ok by 4183
//94.10.20 first design by 4183
//95.09.05 fix by 2495 座落地點從60字到250字
//99.04.02 fix 修正.顯示本季無承受擔保品延長處分申報資料 by 2295
//99.12.06 fix sqlInjection by 2808
%>
<script language="javascript" src="js/FX008W.js"> </script>
<script language="javascript" src="js/Common.js"></script>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<link href="css/b51.css" rel="stylesheet" type="text/css">

<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.ArrayList" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Date" %>


<html>
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
-->
</script>
<%
	//接收資料
	List WLX08_S_Edit = (List)request.getAttribute("WLX08_S_Edit");
	List FX008_Lock = (List)session.getAttribute("WLX08_Lock");    //內含被LOCK的年季
	String title = (WLX08_S_Edit == null)?"新增":"維護";
	String bank_no =  ( request.getParameter("bank_no")==null ) ? " " : (String)request.getParameter("bank_no");
	String s_year = (  request.getParameter("s_year")==null ) ? " " : (String) request.getParameter("s_year");
	String s_quarter = (  request.getParameter("s_quarter")==null ) ? " " : (String) request.getParameter("s_quarter");
	String hquarter=" ";
	String hyear=" ";
	DataObject Bean = new DataObject(); //使用dataobject type 來做資料傳遞
	
	if(WLX08_S_Edit == null){
	   System.out.println("WLX08_S_Edit == null");
	}else{
	   //使用dataobject type 來做資料傳遞
	   Bean = (DataObject)WLX08_S_Edit.get(0);//把選取的資料取出來
	   System.out.println("WLX08_S_Edit.size()="+WLX08_S_Edit.size());
	}

	//取得FX008W的權限
	Properties permission = ( session.getAttribute("FX008W")==null ) ? new Properties() : (Properties)session.getAttribute("ZZ005W");
	if(permission == null){
		System.out.println("FX008W_List.permission == null");
  	}
  else{
  		System.out.println("FX008W_List.permission.size ="+permission.size());
  	}
  
  //判斷申報年季不可是巳鎖定，若已鎖定將返回到list畫面
 	if(WLX08_S_Edit == null){
 			if(FX008_Lock.contains(s_year+s_quarter) || FX008_Lock.contains(s_year+"01"))
 				out.print("<script>alert(\"本季資料已鎖定，將返回清單畫面\");history.go(-1);</script>");
 	}

%>
<head>
<title> 各農漁會承受擔保品延長處分申報情形<%=title%> </title>
</head>


<body marginwidth="0" marginheight="0" leftmargin="0" topmargin="0" leftmargin="0">
<form method="post" >

<table border="0" cellspacing="1" width="600">
  
  <tr>
    <td width="600"></td>
  </tr>
  
  <tr>
  <center>
    <td width="600">
      <table border="0" cellspacing="1">
        <tr>
          <td width="110"><img src="images/banner_bg1.gif" width="110" height="17"></td>
          <td width="380"><p align="center"><font color='#000000' size=4><b>各農漁會承受擔保品延長處分申報情形<%=title%></b></font></p></td>
          <td width="110"><img src="images/banner_bg1.gif" width="110" height="17"></td>
        </tr>
      </table> 
    </td>
  </center>
  </tr>
  
  <tr> 
     <div align="right"><jsp:include page="getLoginUser.jsp" flush="true" /></div> 
  </tr>

  <!-- ---------------------------------------------------------------------------------------------------------------------------- -->
  <tr>
    <td width="600">
    <center>    
    <table class="sbody"  border="1" cellspacing="1" bordercolor="#3A9D99" width="600" >
    	<tr>
    		<td width='239' bgcolor="#D8EFEE" colspan="3">金融機構代號</td>
			<% 

				if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0){
					bank_no = Bean.getValue("bank_no").toString();
				}
			%>
			<td width='398' bgcolor='e7e7e7' height="27"><%=bank_no%></td>
        </tr>
        
        <tr>
        	<td width='239' bgcolor='#D8EFEE' colspan="3">申報年季</td>
			<td width='398' bgcolor='e7e7e7' height="10">
               	<input type='text' name='m_year'
               			value="<%
               			 	if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0)
               			 	{
               			 		out.print(Bean.getValue("m_year"));
               			 		hyear=Bean.getValue("m_year").toString();
               			 	}
               			 	else{
               			 		out.print(s_year);
               			 		hyear = s_year;
               			 	}
               			 	%>"
               			size='3' maxlength='3' onblur='CheckYear(this)' disabled >
               	<font color='#000000'>年第

             	<select id="hide2" name='m_quarter' disabled >
        			<%
        				if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0){
        			 		hquarter = Bean.getValue("m_quarter").toString();
        			 		out.print(hquarter);
        			 	}
        			 	else{
               				out.print(s_quarter);
               				hquarter = s_quarter;
               			}
               			
               			String sel1=" ";
               			String sel2=" ";
               			String sel3=" ";
               			String sel4=" ";

               			if(hquarter.equals("1"))
               				sel1="selected";
            			else if(hquarter.equals("2"))
               				sel2="selected";
                  		else if(hquarter.equals("3"))
               				sel3="selected";
               	  		else if(hquarter.equals("4"))
               				sel4="selected";

        			%>
        			<option value="1" <%=sel1%> >01</option>
        			<option value="2" <%=sel2%> >02</option>
        			<option value="3" <%=sel3%> >03</option>
        			<option value="4" <%=sel4%> >04</option>
        		</select>季&nbsp;&nbsp;</font>

<!-- -----------------------------hide type---------------------------- -->

<input type='hidden' name='hyear' value="<%=hyear%>" >
<input type='hidden' name='hquarter' value="<%=hquarter%>">
<input type='hidden' name='hbank_no' value="<%=bank_no%>">

<%
		String seq_no = " ";
		if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0){
			seq_no = Bean.getValue("seq_no").toString();
%>
			<input type='hidden' name='hseq_no' value="<%=seq_no%>" >
<%
	 }
%>
<!-- -----------------------------end of hide type---------------------------- -->
        		<% //判斷本季是否已經有資料，如果沒有則增加(本季無承受擔保品延長處分申報資料)按鈕
        			if(WLX08_S_Edit == null ){
        				List paramList =new ArrayList() ;
        				String sqlCmd = "select count(*) as Cnt from WLX08_S_GAGE "
        							+ "where BANK_NO  = ?"  
        							+ "and M_YEAR  = ?"
        							+ "and M_Quarter = ?";
						paramList.add(bank_no) ;
						paramList.add(s_year) ;
						paramList.add(s_quarter) ;
						
						List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"cnt");	
						System.out.println("dbData.size()="+dbData.size());
						String count =( ((DataObject)dbData.get(0)).getValue("cnt")==null ) ? "null" :((DataObject)dbData.get(0)).getValue("cnt").toString();
						System.out.println("count="+count);
						if(count.equals("null") || count.equals("0")){
        		%>
        		
        		<br><font color='red' size=2>「本季無承受擔保品延長處分申報資料」</font>
        		<input type="button" name="no_apply" value="確定" 
        		onClick="javascript:doSubmit(this.document.forms[0],'No_Apply','123');">
            	
            	<%	
            			}//end of if(count == "null" || count == "0")
            		}//end of if(WLX08_S_Edit == null )
            	%>
			</td>
        </tr>
        
        <tr>
			<td width='239' bgcolor='#D8EFEE' colspan="3">承受擔保品編號</td>
			<td width='398' bgcolor='e7e7e7' >
				<input type='text' name='dure_no'
             			value="<%
             				if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0)
               				{
               					out.print(Bean.getValue("dureassure_no"));
               			 	}
               		       	%>"
               			size='10' maxlength='10' onBlur='this.value=checknan(this);'
               			style='text-align: right;'> 
               			<font color='red' size=2 >* (限10位數字以內)</font>
			</td>
        </tr>
        
        <tr>
			<td width='239' bgcolor='#D8EFEE' colspan="3" >機構名稱(借款人)</td>
			<td width='398' bgcolor='e7e7e7' >
              	<input type='text' name='debtname'
                		value="<%
                			if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0)
               			 	{
               			 		out.print(Bean.getValue("debtname"));
               			 	}
               			 	%>"
               			size='20' maxlength='20' >
               	<font color='red' size=4>*</font><font color='red' size=4>*</font><input type="button" name="Check_Borrowe_Summary" value="機構名稱(借款人)申辦之歷史明細" onClick="javascript:doSubmit(this.document.forms[0],'load_history','');">
			</td>
		</tr>
        
        <tr>
			<td width='239' bgcolor='#D8EFEE' colspan="3">承受日期<font color='blue' size=2>(A)</font></td>
			<td width='398' bgcolor='e7e7e7' >
			
			<% //----------------get duredate and format it---------------------------------
				Date duDate = new Date();
				int month = 1;
				int day = 1;
      
        		if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0){
        			
        			duDate = (Date)Bean.getValue("duredate");               			 	 
        			month = duDate.getMonth()+1;
            	  	day = duDate.getDate();
        		}
        	%>
            	<input type='text' name='accept_year'
            		value="<%
            			if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0 && duDate != null)
               			{
               			 	out.print(duDate.getYear()-11);

               			}
               			%>"
              		size='3' maxlength='3' onblur='CheckYear(this)'>
            	年
            	<select  name='accept_month'>
            		<%
            		if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0 && duDate != null){
               		%>
            			<option value="<%=month%>"><%=month%></option>
            		<%} else {%>
            			<option value=" ">  </option>
            		<%}%>
            		<option value="01" >01</option>
			        <option value="02" >02</option>
			        <option value="03" >03</option>
			        <option value="04" >04</option>
			        <option value="05" >05</option>
			        <option value="06" >06</option>
			        <option value="07" >07</option>
			        <option value="08" >08</option>
			        <option value="09" >09</option>
			        <option value="10" >10</option>
			        <option value="11" >11</option>
			        <option value="12" >12</option>
			    </select>
			          
		        月
		        <select id="hide4" name='accept_day'>
		            <%
            			if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0 && duDate != null){
               		%>
            			<option value="<%=day%>" ><%=day%></option>
            		<%} else {%>
            			<option value=" ">  </option>
            		<%}%>
		          	<option value="01" >01</option>
		          	<option value="02" >02</option>
		          	<option value="03" >03</option>
		          	<option value="04" >04</option>
		          	<option value="05" >05</option>
		          	<option value="06" >06</option>
		          	<option value="07" >07</option>
		          	<option value="08" >08</option>
		          	<option value="09" >09</option>
		          	<option value="10" >10</option>
		          	<option value="11" >11</option>
		          	<option value="12" >12</option>
		          	<option value="13" >13</option>
		          	<option value="14" >14</option>
		          	<option value="15" >15</option>
		          	<option value="16" >16</option>
		          	<option value="17" >17</option>
			        <option value="18" >18</option>
			        <option value="19" >19</option>
			        <option value="20" >20</option>
			        <option value="21" >21</option>
			        <option value="22" >22</option>
			        <option value="23" >23</option>
			        <option value="24" >24</option>
			        <option value="25" >25</option>
			        <option value="26" >26</option>
			        <option value="27" >27</option>
			        <option value="28" >28</option>
			        <option value="29" >29</option>
			        <option value="30" >30</option>
			        <option value="31" >31</option>
			    </select>
			   	日
			   	<font color='red' size=4>*</font>
			</td>
		</tr>
		
		
		
		
		
		
		<tr>
        	<td width='239' bgcolor='#D8EFEE' colspan="3" height="17">承受擔保品座落</td>
			<td width='398' bgcolor='e7e7e7' height="39">
               	<textarea rows="5" cols="50" name='duresite'><%
                			if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0){	out.print( Bean.getValue("dureassuresite") );
               			 	}
               	%></textarea>
               	<font color='red' size=4>*</font> 
               	<br><font color="#FF0000" size="2">(土地:段、地號、面積、持分；建物:建號)</font>
			</td>
		</tr>
         
        <tr>
			<td width='239' bgcolor='#D8EFEE' colspan="3" height="17">帳列金額</td>
			<td width='398' bgcolor='e7e7e7' height="36">
				<input type='text' name='account'
					value="<%
							if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0){
								out.print( Utility.setCommaFormat( Bean.getValue("accountamt").toString())  );	
							}
          					%>"size='15' maxlength='12' onFocus='this.value=changeVal(this)' onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;'>
               	<font color='red' size=4>*&nbsp;&nbsp;</font>
               	<font size="2" color="#FF0000">(單位:新台幣元)</font>
			</td>
        </tr>
        
        <tr>
          	<td width='239' bgcolor='#D8EFEE' colspan="3" height="17">申請延長期間</td>
			<td width='398' bgcolor='e7e7e7' height="24">
				<input type='text' name='apply_year'
					value="<%
                 			if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0){
                 				out.print( Bean.getValue("applydelayyear") );
                 			}
                 			%>"
                     size='2' maxlength='2' onblur='CheckYear(this)'>
                 年
                
		            <select id="text" name='apply_month'>
		            <%
		            	if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0){
		            		out.print("<option value=\""+Bean.getValue("applydelaymonth")+"\">"+Bean.getValue("applydelaymonth")+"</option>");
		            	 }
		            %>
		            	<option value="0" > 0</option>
						<option value="01" >01</option>
						<option value="02" >02</option>
						<option value="03" >03</option>
						<option value="04" >04</option>
						<option value="05" >05</option>
						<option value="06" >06</option>
						<option value="07" >07</option>
						<option value="08" >08</option>
						<option value="09" >09</option>
						<option value="10" >10</option>
						<option value="11" >11</option>
						<option value="12" >12</option>
					</select>個月</font>                           
		            <font color='red' size=4>*</font>
			</td>
		</tr>
        
        <tr>
        	<td width='239' bgcolor='#D8EFEE' colspan="3" height="17">申請延長理由</td>
			<td width='398' bgcolor='e7e7e7' height="39">
               	<textarea rows="4" cols="50" name='apply_reason'><%
                			if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0){
                				out.print( Bean.getValue("applydelayreason") );
               			 	}
               	%></textarea>
               	<font color='red' size=4>*</font> 
			</td>
		</tr>
          
        <tr>
			<td width='70'  bgcolor='#D8EFEE' rowspan ="2" > <font color="#0000FF">縣市政府</font></td>
			<td width='150' bgcolor='#D8EFEE' colspan ="2" height="17">核准申請延長期間<font color='blue' size=2>(B)</font></td>
			<td width='380' bgcolor='e7e7e7'>
				<input type='text' name='Audit_Delay_y' 
						value="<%
								if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0 && Bean.getValue("audit_applydelayyear") != null){
										out.print(Bean.getValue("audit_applydelayyear"));
									}
								%>" 
						size='2' maxlength='2' disabled>&nbsp;年
						
		            <select id="text" name='Audit_Delay_m' disabled >
		            <%
		            	if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0){
		            		out.print("<option value=\""+Bean.getValue("audit_applydelaymonth")+"\">"+Bean.getValue("audit_applydelaymonth")+"</option>");
		            	 }
		            	 else{
		            	 	out.print("<option value=\" \"> </option>");
		            	 }
		            %>
						<option value="01" >01</option>
						<option value="02" >02</option>
						<option value="03" >03</option>
						<option value="04" >04</option>
						<option value="05" >05</option>
						<option value="06" >06</option>
						<option value="07" >07</option>
						<option value="08" >08</option>
						<option value="09" >09</option>
						<option value="10" >10</option>
						<option value="11" >11</option>
						<option value="12" >12</option>
					</select>個月</font>                           
		            <font color='red' size=4>*</font>      
			</td>
		</tr>
			
		<tr>
			<td width='150' bgcolor='#D8EFEE' colspan ="2" height="17">核准申請延長期限<font color='blue' size=2>(A)+(B)</font></td>
			<td width='380' bgcolor='#e7e7e7'>
					<%
						Date get_Cnt_date;
						String Cnt_year = " ";
						String Cnt_month = " ";
						String Cnt_date  = " ";
						
						if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0 && Bean.getValue("audit_duredate")!= null){	
							get_Cnt_date = (Date)Bean.getValue("audit_duredate");
							Cnt_year  = String.valueOf(get_Cnt_date.getYear()-11);
							Cnt_month = String.valueOf(get_Cnt_date.getMonth()+1);
							Cnt_date  = String.valueOf(get_Cnt_date.getDate());
						}

					%>
                    <input type='text' name='Cnt_year'  value="<%=Cnt_year %>" size='3' maxlength='3' disabled >&nbsp;年      
					<input type='text' name='Cnt_month' value="<%=Cnt_month %>" size='2' maxlength='2' disabled >&nbsp;月      
					<input type='text' name='Cnt_date'  value="<%=Cnt_date %>" size='2' maxlength='2' disabled >&nbsp;日                  
		            <font color='red' size=4>*&nbsp;</font>                   
			</td>
		</tr>
        
        <tr>  
			<td width='30' align='left' bgcolor='#D8EFEE' rowspan="4" height="124"><font color="#FF0000">審查總結</font><font color="#0000FF">(縣/市政府評註)</font></td>
			<td width='67' align='left' bgcolor='#D8EFEE' rowspan="4" height="124"><font color="#0000FF">是否符合行政院農業委員會93.10.15農授金字第0935080230號令規定</font></td>
			<td width='130' align='left' bgcolor='#D8EFEE' height="24">是否提足備抵趺價損失</td>
			<td width='337' bgcolor='e7e7e7' height="24">
				<select name='damage_yn' size="1" disabled >
				<%
					String selected1 = " ";
					String selected2 = " ";
					String selected3 = " ";
					
					
					if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0){
						
						String YN1 = String.valueOf(Bean.getValue("damage_yn"));

           				if(YN1 != null){	
            				if( YN1.equals("Y") )
              					selected1 = "selected";
            				else if(YN1.equals("N"))
            					selected2 = "selected";
		            		else 
		            			selected3 = "selected";
            			}//end of inner if
					}//end of outer if
			%>
    	  			<option value=" " <%=selected3 %> >  </option>
    	  			<option value="Y" <%=selected1 %> >是</option>
    	  			<option value="N" <%=selected2 %> >否</option>
			</td>
		</tr>
        
        <tr>
			<td width='130' align='left' bgcolor='#D8EFEE' height="24">是否有積極處分之事實</td>
			<td width='337' bgcolor='e7e7e7' height="24">
				<select name='disposal_fact_yn' size="1" disabled >
				<%
					selected1 = " ";
					selected2 = " ";
					selected3 = " ";

        			if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0){
						
						String YN2 = String.valueOf(Bean.getValue("disposal_fact_yn"));
						
           				if(YN2 != null){  
                                          
            				if( YN2.equals("Y") )
              					selected1 = "selected";
            				else if(YN2.equals("N"))
            					selected2 = "selected";
		            		else 
		            			selected3 = "selected";
		            			
           				}//end of inner if
        			}//end of inner if
      			%>
    	  			<option value=" " <%=selected3 %> >  </option>
    	  			<option value="Y" <%=selected1 %> >是</option>
    	  			<option value="N" <%=selected2 %> >否</option>
      	</tr>
      
      	<tr>
			<td width='130' align='left' bgcolor='#D8EFEE' height="34">未來處分計劃是否合理可行</td>
			<td width='337' bgcolor='e7e7e7' height="34">
				<select name='disposal_plan_yn' size="1" disabled >
				<%
					selected1 = " ";
					selected2 = " ";
					selected3 = " ";
                	
        			if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0){
        				
						String YN3 = String.valueOf(Bean.getValue("disposal_plan_yn"));
                	
           				if(YN3 != null){
            				if( YN3.equals("Y") )
              					selected1 = "selected";
            				else if(YN3.equals("N"))
            					selected2 = "selected";
		            		else 
		            			selected3 = "selected";
            			}
         			}//end of if
       			%>
    	  			<option value=" " <%=selected3 %> >  </option>
    	  			<option value="Y" <%=selected1 %> >是</option>
    	  			<option value="N" <%=selected2 %> >否</option>
			</td>
		</tr>
	
		<tr>
			<td width='130' align='left' bgcolor='#D8EFEE' height="24">審核結果</td>
			<td width='337' bgcolor='e7e7e7' height="24">
				<select name='auditresult_yn' size="1" disabled >
				<%
					selected1 = " ";
					selected2 = " ";
					selected3 = " ";
    	
    	     		if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0){
    	     			
						String YN4 = String.valueOf(Bean.getValue("auditresult_yn"));
    	
    	       			if(YN4 != null){
    	        			if( YN4.equals("Y") )
    	          				selected1 = "selected";
    	        			else if(YN4.equals("N"))
    	        				selected2 = "selected";
		            		else 
		            			selected3 = "selected";
    	        		}
    	      		}
    	  		%>
    	  			<option value=" " <%=selected3 %> >  </option>
    	  			<option value="Y" <%=selected1 %> >合格</option>
    	  			<option value="N" <%=selected2 %> >不合格</option>
    	  			
    	  	</td>
   		</tr>
          
   		<tr>
			<tr>
				<td width="70" bgcolor='#D8EFEE' rowspan="4"> <font color="#0000FF">縣市政府</font></td>
				<td width="60" bgcolor='#D8EFEE' rowspan="2">審核核准</td>
				<td width="90" bgcolor='#D8EFEE' >核准文號</td>
				<td width="380" bgcolor='#e7e7e7'>
					<input type='text' name='DocNo' 
						value="<%
									if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0 && Bean.getValue("applyok_docno")!= null){
										out.print( Bean.getValue("applyok_docno"));
									}
								%>"  
						size='30' maxlength='30' disabled >
					<font color='red' size=4>*</font>         
				</td>
			</tr>
	
			<tr>
				<td width="90" bgcolor='#D8EFEE'>核准日期</td>
				<td width="380" bgcolor='#e7e7e7'>
					<%
						Date get_ApplyOK_date;
						
						String ApplyOK_y  = " ";
						String ApplyOK_m = " ";
						String ApplyOK_d  = " ";
						
						if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0 && Bean.getValue("applyok_date")!= null){	
							get_ApplyOK_date = (Date)Bean.getValue("applyok_date");
							ApplyOK_y = String.valueOf(get_ApplyOK_date.getYear()-11);       
							ApplyOK_m = String.valueOf(get_ApplyOK_date.getMonth()+1);
							ApplyOK_d = String.valueOf(get_ApplyOK_date.getDate());
						}
					%>
					<input type='text' name='ApplyOK_y' value="<%=ApplyOK_y%>" size='4' maxlength='4' disabled >&nbsp;年      
					<input type='text' name='ApplyOK_m' value="<%=ApplyOK_m%>" size='2' maxlength='2' disabled >&nbsp;月      
					<input type='text' name='ApplyOK_d' value="<%=ApplyOK_d%>" size='2' maxlength='2' disabled >&nbsp;日      
					<font color='red' size=4>*</font>   				
				</td>
			</tr>
			
			<tr>
				<td width="70" bgcolor='#D8EFEE' rowspan="2" >函報農金局</td>
				<td width="90" bgcolor='#D8EFEE'>函報文號</td>
				<td width="380" bgcolor='#e7e7e7'>
					<input type='text' name='BOAF_DocNo' 
						value="<%
									if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0 && Bean.getValue("report_boaf_docno")!= null){
										out.print( Bean.getValue("report_boaf_docno"));
									}
								%>"  
						size='30' maxlength='30' disabled >
					<font color='red' size=4>*</font>         
				</td>          
			</tr>              
			                   
			<tr>
				<td width="75" bgcolor='#D8EFEE'>函報日期</td>
				<td width="380" bgcolor='#e7e7e7'>
					<%
						Date get_BOAF_date;
						
						String BOAF_y  = " ";
						String BOAF_m = " ";
						String BOAF_d  = " ";
						
						if(WLX08_S_Edit != null && WLX08_S_Edit.size() != 0 && Bean.getValue("report_boaf_date") != null){	
							get_BOAF_date = (Date)Bean.getValue("report_boaf_date");
							BOAF_y = String.valueOf(get_BOAF_date.getYear()-11);       
							BOAF_m = String.valueOf(get_BOAF_date.getMonth()+1);
							BOAF_d = String.valueOf(get_BOAF_date.getDate());
						}
					%>
					<input type='text' name='BOAF_y' value="<%=BOAF_y%>" size='4' maxlength='4' disabled >&nbsp;年      
					<input type='text' name='BOAF_m' value="<%=BOAF_m%>" size='2' maxlength='2' disabled >&nbsp;月      
					<input type='text' name='BOAF_d' value="<%=BOAF_d%>" size='2' maxlength='2' disabled >&nbsp;日      
					<font color='red' size=4>*</font>  
				</td>
			</tr>
				
      </table>
      </center>
    </td>
  </tr>
  <!-- ---------------------------------------------------------------------------------------------------------------------------- -->
  <tr>
    <td width="600">
       <center>
          <jsp:include page="getMaintainUser.jsp" flush="true" /><!--載入維護者資訊 -->
       </center>
    </td>
  </tr>

  <tr>
    <td width="600">
      <center>
        <table border="0" cellspacing="1">
          <tr>
          	<%
          		if(WLX08_S_Edit == null){
          	%>
          	<td width="66">
          		<a href="javascript:doSubmit(this.document.forms[0],'Insert','123');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_confirmb.gif',1)">
          		<img src="images/bt_confirm.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>
          	</td>
          	<%}else {%>
						<td width="66">
							<a href="javascript:doSubmit(this.document.forms[0],'Update','123');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_updateb.gif',1)">
							<img src="images/bt_update.gif" name="Image105" width="66" height="25" border="0" id="Image105"></a>
						</td>
						<td width="66">
							<a href="javascript:doSubmit(this.document.forms[0],'Delete','123');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_deleteb.gif',1)">
							<img src="images/bt_delete.gif" name="Image103" width="66" height="25" border="0" id="Image103"></a>
						</td>
						<%}%>
           	 <td width="93">
            		<a href="javascript:history.back();"  onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image106','','images/bt_backb.gif',1)">
            		<img src="images/bt_back.gif" name="Image106" width="80" height="25" border="0" id="Image106"></a>
            	</td>
          </tr>
        </table>
      </center>
   </td>
  </tr>

  <tr>
    <td width="600">
    
      <table class="sbody" border="0" cellspacing="1">
        <tr>
        	<td colspan="2">
				<font color='#990000'><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle">
				<font color="#007D7D">使用說明 : </font></font>
          	</td>
        </tr>
        
        <tr>
        	<td width="16">&nbsp;</td>
          	<td width="577"> 
          		<ul>
          			<li>確認輸入資料無誤後, 按【確定】即將本表上的資料, 於資料庫中建檔,建檔完成按[繼續新增下一筆]可在本畫面繼續新增建檔,或按[回查詢主畫面]回至查詢清單畫面。</li>
            		<li>如果本季無須申報資料,仍須選確定(Y)並按[本季無承受擔保品延長處分申報資料]按鈕,以示確定</li>
            		<li>按<font color="#666666">【修改】</font>即修改的資料,寫入資料庫料庫中。</font></li>
            		<li>欲重新輸入資料, 按<font color="#666666">【取消】即將本表上的資料清空</font></li>
        			<li>如放棄修改或無修改之資料需輸入, 按【回上一頁】]即離開本程式。</li>
        			<li>【<font color="red" size=4>*</font>】為必填欄位。</li>
        			<li><font color="red" size="4">&nbsp;</font>如果審查總結中,有乙項填「否」,而審核結果填「合格」會先出現「審查結果欄填註與細項欄內有不一致,請確定?」警訊,若你確定仍會依你所建檔結果存檔</li>
            	</ul>
          	</td>
        </tr>
      </table>
      
    </td>
  </tr>

</table>

</body>

</html>
