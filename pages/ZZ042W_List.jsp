<%
// 94.12.13 create by 2295
// 95.03.21 add A111111111比照農金局查詢權限 by 2295
// 99.12.10 fix sqlInjection by 2808
//100.01.05 fix 日期顯示 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
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
    //DataObject rptData = null;
	String bank_type = (session.getAttribute("bank_type") == null)?"":(String)session.getAttribute("bank_type");				
	String tbank_no = (session.getAttribute("tbank_no") == null)?"":(String)session.getAttribute("tbank_no");				
	String sztrans_type = ( request.getParameter("TRANS_TYPE")==null ) ? "" : (String)request.getParameter("TRANS_TYPE");				
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");				
	String szrpt_code = ( request.getParameter("RPT_CODE")==null ) ? "" : (String)request.getParameter("RPT_CODE");					
	String szrpt_output_type = ( request.getParameter("RPT_OUTPUT_TYPE")==null ) ? "" : (String)request.getParameter("RPT_OUTPUT_TYPE");					
	String szrpt_sort = ( request.getParameter("RPT_SORT")==null ) ? "" : (String)request.getParameter("RPT_SORT");					
	String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? YEAR : (String)request.getParameter("S_YEAR");				
	String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? MONTH : (String)request.getParameter("S_MONTH");				
	String E_YEAR = ( request.getParameter("E_YEAR")==null ) ? YEAR : (String)request.getParameter("E_YEAR");				
	String E_MONTH = ( request.getParameter("E_MONTH")==null ) ? MONTH : (String)request.getParameter("E_MONTH");				
	String szbank_code = ( request.getParameter("bank_code")==null ) ? "" : (String)request.getParameter("bank_code");				
	String szhsien_id = ( request.getParameter("HSIEN_ID")==null ) ? "ALL" : (String)request.getParameter("HSIEN_ID");
    String lguser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		    
    String firstStatus = ( request.getParameter("firstStatus")==null ) ? "" : (String)request.getParameter("firstStatus");			    	
	System.out.println("ZZ042W_List.act="+act);	
	System.out.println("ZZ042W_List.bank_type="+bank_type);		    
	System.out.println("ZZ042W_List.szrpt_code="+szrpt_code);		    
	System.out.println("ZZ042W_List.szrpt_output_type="+szrpt_output_type);		    
	System.out.println("ZZ042W_List.szhsien_id="+szhsien_id);		    
	System.out.println("ZZ042W_List.szrpt_sort="+szrpt_sort);		    	
	System.out.println("ZZ042W_List.S_YEAR="+S_YEAR);		    
	System.out.println("ZZ042W_List.S_MONTH="+S_MONTH);		    
	System.out.println("ZZ042W_List.E_YEAR="+E_YEAR);		    
	System.out.println("ZZ042W_List.E_MONTH="+E_MONTH);		    
	System.out.println("ZZ042W_List.firstStatus="+firstStatus);		    
    List reportList = (List)request.getAttribute("reportList");	
	String sqlCmd = "";
	List paramList =new ArrayList () ;
	if(lguser_id.equals("A111111111")){
	   bank_type="2";//95.03.21 add A111111111比照農金局查詢權限
	}
	//根據使用者權限帶出可查詢到的報表名稱
	sqlCmd = " select  distinct  rpt_nof.rpt_code,rpt_nof.rpt_name,rpt_nof.rpt_output_type,rpt_nof.rpt_include "
		   + " from rpt_nof,rpt_notypef "
		   + " where rpt_nof.rpt_code = rpt_notypef.rpt_code "
	       + " and   rpt_notypef.rpt_cate  =  ? "
	       + " order by  rpt_nof.rpt_code";
	paramList.add(bank_type) ;
    List rptname_list = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");//報表名稱    
    paramList.clear() ;
    if(szrpt_code.equals("")){
       szrpt_code = (String)((DataObject)rptname_list.get(0)).getValue("rpt_code");
    }
    sqlCmd = " select  distinct  rpt_nof.rpt_code,rpt_nof.rpt_name,rpt_nof.rpt_output_type,rpt_nof.rpt_include "
		   + " from rpt_nof,rpt_notypef "
		   + " where rpt_nof.rpt_code = ? "
		   + " and   rpt_nof.rpt_code = rpt_notypef.rpt_code "
	       + " and   rpt_notypef.rpt_cate  = ? "
	       + " order by  rpt_nof.rpt_code";
	paramList.add(szrpt_code) ;
	paramList.add(bank_type) ;
    List rptData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");//報表data    
    System.out.println("rptData.size()="+rptData.size());
	//取得ZZ042W的權限
	Properties permission = ( session.getAttribute("ZZ042W")==null ) ? new Properties() : (Properties)session.getAttribute("ZZ042W"); 
	if(permission == null){
       System.out.println("ZZ042W_List.permission == null");
    }else{
       System.out.println("ZZ042W_List.permission.size ="+permission.size());               
    }	    
    
    // XML Ducument for 報表名稱 begin
    out.println("<xml version=\"1.0\" encoding=\"UTF-8\" ID=\"ReportListXML\">");
    out.println("<datalist>");
    for(int i=0;i< rptname_list.size(); i++) {
        DataObject bean =(DataObject)rptname_list.get(i);
        out.println("<data>");        
        out.println("<rptCode>"+bean.getValue("rpt_code")+"</rptCode>");        
        out.println("<rptName>"+bean.getValue("rpt_name")+"</rptName>");                
        out.println("<rptOutputType>"+bean.getValue("rpt_output_type")+"</rptOutputType>");                
        out.println("<rptInclude>"+bean.getValue("rpt_include")+"</rptInclude>");                
		out.println("</data>");        
    }
    out.println("</datalist>\n</xml>");
    // XML Ducument for 報表名稱 end 	
    //縣市別
	List hsien_id_data = DBManager.QueryDB_SQLParam("select hsien_id,hsien_name from cd01 order by fr001w_output_order",null,"");     
    // XML Ducument for 縣市別 begin
    out.println("<xml version=\"1.0\" encoding=\"UTF-8\" ID=\"hsienListXML\">");
    out.println("<datalist>");
    for(int i=0;i< hsien_id_data.size(); i++) {
        DataObject bean =(DataObject)hsien_id_data.get(i);
        out.println("<data>");        
        out.println("<hsien_id>"+bean.getValue("hsien_id")+"</hsien_id>");        
        out.println("<hsien_name>"+bean.getValue("hsien_name")+"</hsien_name>");                        
		out.println("</data>");        
    }
    out.println("</datalist>\n</xml>");
    // XML Ducument for 縣市別 end 	
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/ZZ042W.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<title>「管理報表查詢下載」</title>
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
<form method=post action='#'>
<input type="hidden" name="act" value="">  
<%if(reportList != null && reportList.size() != 0){%>
<input type="hidden" name="row" value="<%=reportList.size()+1%>">   
<%}%>
<table width="680" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  		<tr> 
   		 <td><img src="images/space_1.gif" width="12" height="12"></td>
  		</tr>
		<%
          String nameColor="nameColor_sbody";
          String textColor="textColor_sbody";
          String bordercolor="#3A9D99";
        %>  
        <tr> 
          <td bgcolor="#FFFFFF">
		  <table width="680" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="680" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="220"><img src="images/banner_bg1.gif" width="220" height="17"></td>
                      <td width="240"><font color='#000000' size=4><b> 
                        <center>
                          「管理報表查詢下載」 
                        </center>
                        </b></font> </td>
                      <td width="220"><img src="images/banner_bg1.gif" width="220" height="17"></td>
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
                      <td><table width=680 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="<%=bordercolor%>">                                                    
                          <tr class="sbody">						  
						  <td width='10%' align='left' class="<%=nameColor%>">報表名稱</td>
                          <td width='90%' class="<%=textColor%>">	
                            <select name='RPT_CODE' onchange="javascript:changeOption(this.document.forms[0]);">                                                                                    
                            <%
                            if(rptname_list != null){
                             for(int i=0;i<rptname_list.size();i++){%>
                            <option value="<%=(String)((DataObject)rptname_list.get(i)).getValue("rpt_code")%>"                                                        
                            <%if(szrpt_code.equals((String)((DataObject)rptname_list.get(i)).getValue("rpt_code"))){ 
                                 out.print("selected");                                
                              }   
                            %>
                            ><%=(String)((DataObject)rptname_list.get(i)).getValue("rpt_code")%>&nbsp;<%=(String)((DataObject)rptname_list.get(i)).getValue("rpt_name")%></option>                            
                            <%}
                            }//end of if
                            %>
                            </select>    
                            &nbsp;                            
                            <input type="button" value="查詢清單" onclick="javascript:doSubmit(this.document.forms[0],'List');" name="QryList">                       
                            <input type="button" value="報表下載" onclick="javascript:doSubmit(this.document.forms[0],'MultiFiles');" name="RptDownload">  &nbsp;  
                          </td>             
                          </tr>       
                                                         
                          <tr class="sbody">                          
                          <td width='10%' align='left' class="<%=nameColor%>">農(漁)會別</td>
                          <td width='90%' class="<%=textColor%>">
                            <select name='RPT_OUTPUT_TYPE'>                                                                                    
                            <%
                             if(rptData != null){
                                if ( ((String)((DataObject)rptData.get(0)).getValue("rpt_output_type")).equals("X")){//顯示農會/漁會/農漁會  				  				 
                                   if(szrpt_output_type.equals("6")){
                                      out.print("<option value='6' selected>農會</option>");
                                   }else{
                                      out.print("<option value='6' >農會</option>");
                                   }
                                   if(szrpt_output_type.equals("7")){
  				                      out.print("<option value='7' selected>漁會</option>");
  				                   }else{
  				                      out.print("<option value='7'>漁會</option>");
  				                   }
  				                   if(szrpt_output_type.equals("ALL")){
  				                      out.print("<option value='ALL' selected>農/漁會</option>");  			       
  				                   }else{
  				                      out.print("<option value='ALL'>農/漁會</option>");  			       
  				                   }
  		    				    }else if ( ((String)((DataObject)rptData.get(0)).getValue("rpt_output_type")).equals("T")){//顯示農漁會  				  		    	 
  		    				       if(szrpt_output_type.equals("ALL")){
  		    				          out.print("<option value='ALL' selected>農/漁會</option>");  			         		    				       
  		    				       }else{
  		    				          out.print("<option value='ALL'>農/漁會</option>");  			       
  		    				       }
  		    				    }   
  		    				 }
                            %>
                            </select>                                 
                          </td>         
                          </tr>
                          
                          
                          <tr class="sbody"> 
                          
                          <td width='10%' align='left' class="<%=nameColor%>">縣市別</td> 
                          <td width='90%' class="<%=textColor%>">                                                      
                               <select name="HSIEN_ID">            
                                <%if(rptData != null){
                                   if ( ((String)((DataObject)rptData.get(0)).getValue("rpt_output_type")).equals("T")){//顯示全部縣市%>                                                     
                                         <option value="ALL">全部縣市</option>
                                <% }else if (!((String)((DataObject)rptData.get(0)).getValue("rpt_include")).equals("X")){//顯示全部縣市%>                                                               
                                         <option value="ALL">全部縣市</option>
                                <% }else{%>
                                         <option value="ALL">全部縣市</option>
                                         <%for(int i=0;i<hsien_id_data.size();i++){%>                                
                                		   <option value="<%=(String)((DataObject)hsien_id_data.get(i)).getValue("hsien_id")%>"
                                           <%if(((String)((DataObject)hsien_id_data.get(i)).getValue("hsien_id")).equals(szhsien_id)) out.print("selected");%>
                                           ><%=(String)((DataObject)hsien_id_data.get(i)).getValue("hsien_name")%></option>                            
                                        <%}%>
                                <% }
                                  }//end of rptData != null
                                %>
                                
                              </select>
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
                          
                          
                          
                          
                          <tr class="sbody">                          
                          <td width='10%' align='left' class="<%=nameColor%>">列表順序</td>
                          <td width='90%' class="<%=textColor%>">                          
                            <select name='RPT_SORT'>                                                                                                                              
                            <%if(rptData != null){
                                 if ( ((String)((DataObject)rptData.get(0)).getValue("rpt_include")).equals("T")){//顯示按報表年月、農(漁)會別%>                                                     
                                 <option value="4" <%if(szrpt_sort.equals("4"))out.print("selected");%> >按報表年月、農(漁)會別</option>  			     
                            <%   }else{%>
								 <option value="1" <%if(szrpt_sort.equals("1"))out.print("selected");%> >按縣市別、機構別、農(漁)會別、報表年月</option>
  			     				 <option value="2" <%if(szrpt_sort.equals("2"))out.print("selected");%> >按縣市別、機構別、報表年月、農(漁)會別</option>
  			     				 <option value="3" <%if(szrpt_sort.equals("3"))out.print("selected");%> >按報表年月、縣市別、機構別、農(漁)會別</option>  			     
                            <%   }
                              }
                            %>    
                            </select>                                 
                            
                          </td>         
                          </tr>                                                    		      					      
                          </table>      
                      </td>    
                      </tr>
                      
                      
                      <tr> 
                      <td><table width=680 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="<%=bordercolor%>">                                              	
                        <tr class="sbody">
                          <td width='100%' colspan=10 class="<%=nameColor%>">		                         	 						  
 							<a href="javascript:selectAll(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image104','','images/bt_selectallb.gif',1)"><img src="images/bt_selectall.gif" name="Image104" width="80" height="25" border="0" id="Image104"></a>                       
 						    <a href="javascript:selectNo(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_selectnob.gif',1)"><img src="images/bt_selectno.gif" name="Image105" width="80" height="25" border="0" id="Image105"></a>                        							 						     						  
                          </td>
                        </tr>     		                      
                      <%                     
                      int i = 0;                            
                      String bgcolor="#D3EBE0";     
                      String listTitle="listTitleColor_sbody"; 
                      String list1Color="list1Color_sbody";
                      String list2Color="list2Color_sbody";
                      String listColor="list1Color_sbody";                          
                      if(reportList != null){ %>
                       	  <tr class="sbody" >                       	  
                            <td class="<%=listTitle%>" width="30">序號</td>                                  
                            <td class="<%=listTitle%>" width="30">選項</td>                            
                            <td class="<%=listTitle%>" width="220">機構(報表)名稱</td>
            				<td class="<%=listTitle%>" width="60">農(漁)會別</td>            				
            				<td class="<%=listTitle%>" width="60">報表年月</td>
            				<td class="<%=listTitle%>" width="70">檔名</td>            				
            				<td class="<%=listTitle%>" width="80">檔案產生日期</td>            				            				
					      </tr>   
                   		    <% if(reportList.size() == 0){%>
                   			   <tr class="sbody" class="<%=list1Color%>">
                   			   <td colspan=10 align=center>無資料可供查詢</td><tr>
                   			   <tr>                   			   
                   			<% }else{ 
                   			    String update_date="";
                   			    String yyymm="";
                   			    
                   			    while(i < reportList.size()){                     		                          		      
                    		      listColor = (i % 2 == 0)?list2Color:list1Color;	
                    		      //bgcolor = (i % 2 == 0)?"#e7e7e7":"#D3EBE0";                    		                         		      
                    		      if((String)((DataObject)reportList.get(i)).getValue("s_updatedate") != null){
                    		         update_date = (String)((DataObject)reportList.get(i)).getValue("s_updatedate");
                    		         //100.01.05 fix 日期顯示 by 2295
                    		         if(update_date.startsWith("1")){
                                		update_date = update_date.substring(0,3)+"/"+update_date.substring(3,5)+"/"+update_date.substring(5,7)+ " " + update_date.substring(8,update_date.length());	                    		
                                	 }else{
                                		update_date = update_date.substring(0,2)+"/"+update_date.substring(2,4)+"/"+update_date.substring(4,6)+ " " + update_date.substring(7,update_date.length());	                    		
                                	 }        
                    		         System.out.println("update_date"+update_date);
                    		      }
                    		      if( ((DataObject)reportList.get(i)).getValue("s_yymm") != null){
                    		         yyymm = (String)((DataObject)reportList.get(i)).getValue("s_yymm");
                    		         if(yyymm.length() == 4){
		            					 yyymm = "0"+yyymm;
		          					 } 
                    		      }
                    		%>                         	       
                    		<tr class="<%=listColor%>">                       	  
                               <td width="30"><%=i+1%></td>                                  
                               <td width="30"><input type="checkbox" name="isModify_<%=(i+1)%>" value="<%=yyymm%>:<%if( ((DataObject)reportList.get(i)).getValue("rpt_fname") != null ) out.print((String)((DataObject)reportList.get(i)).getValue("rpt_fname"));%>"</td>                            
                               <td width="220">
                               <%if( ((DataObject)reportList.get(i)).getValue("s_report_name") != null ) out.print("<a href='ZZ042W.jsp?act=Download&filename="+(String)((DataObject)reportList.get(i)).getValue("rpt_fname")+"&yyymm="+yyymm+"&RPT_CODE="+szrpt_code+"'>"+(String)((DataObject)reportList.get(i)).getValue("s_report_name")+"</a>"); else out.print("&nbsp;");%>
                               </td>
            				   <td width="60">
            				   <%if( ((DataObject)reportList.get(i)).getValue("s_banktype_name") != null ) out.print((String)((DataObject)reportList.get(i)).getValue("s_banktype_name")); else out.print("&nbsp;");%>            				   
            				   </td>            				
            				   <td width="60">
            				   <%if( ((DataObject)reportList.get(i)).getValue("s_yymm") != null ) out.print((String)((DataObject)reportList.get(i)).getValue("s_yymm")); else out.print("&nbsp;");%>
            				   </td>
            				   <td width="70">            				  
            				   <%if( ((DataObject)reportList.get(i)).getValue("rpt_fname") != null ) out.print((String)((DataObject)reportList.get(i)).getValue("rpt_fname")); else out.print("&nbsp;");%>
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
<%if(firstStatus.equals("true")){%>
<script language="JavaScript" >
<!--
changeOption(this.document.forms[0]);
-->
</script>
<%}%>
</html>
