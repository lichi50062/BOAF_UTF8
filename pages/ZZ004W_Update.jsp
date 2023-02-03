<%
//93.12.29 create by 2295
//94.01.07 fix super user不可指派ZZ003W.ZZ005W給系統管理者 by 2295
//             系統管理者不可指派ZZ001W.ZZ002W.ZZ004W給一般使用者  by 2295
//94.02.03 fix 區分網際網路申報/MIS管理系統的配色 by 2295
//94.03.08 fix act=del把有的權限勾選,並disabled by 2295
//99.12.09 fix sqlInjection by 2808
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.ArrayList" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.lang.Boolean" %>

<%
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");					 	
	String szid_Create = ( request.getParameter("id_Create")==null ) ? "" : (String)request.getParameter("id_Create");					 
	String szid_notCreate = ( request.getParameter("id_notCreate")==null ) ? "" : (String)request.getParameter("id_notCreate");					 	
	String szMenu = ( request.getParameter("menu")==null ) ? "" : (String)request.getParameter("menu");					 	 
		 
	String muser_type = ( session.getAttribute("muser_type")==null ) ? "" : (String)session.getAttribute("muser_type");					
	String muser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");					
	String title="";
	String title_act="";
	String sqlCmd = "";
	List paramList = new ArrayList();
	title = (muser_type.equals("S"))?"系統管理者_":"使用者_";		
	title_act = (act.equals("del"))?"刪除建檔":"異動建檔";			
	List WTT01_hasNotCreateList = (List)request.getAttribute("WTT01_hasNotCreateList");		
	List WTT01_hasCreateList = (List)request.getAttribute("WTT01_hasCreateList");		
	List WTT01_notCreate = (List)request.getAttribute("WTT01_notCreate");		
	List WTT01_Create = (List)request.getAttribute("WTT01_Create");		
	List WTT04 = (List)request.getAttribute("WTT04");		
    	
	List WTT03_1 = (List)request.getAttribute("WTT03_1");		
	if(WTT03_1 == null){
	   System.out.println("WTT03_1 == null");
	}else{
	   System.out.println("WTT03_1.size()="+WTT03_1.size());
	}
	
    if(muser_type.equals("S")){	
	sqlCmd = " select a.cmuse_id as trans_type,a.cmuse_name as trans_type_name"
		   + " ,b.cmuse_id as report_no,b.cmuse_name as report_name"
		   + " ,a.input_order,b.input_order"
		   + " from cdshareno a,cdshareno b"
		   + " where a.cmuse_div='011' and b.cmuse_div in ('012','013','014') "
		   + " and a.identify_no = b.identify_no"
		   + " order by a.input_order,b.input_order";
	}
	if(muser_type.equals("A")){
	sqlCmd = " select distinct wtt04_1d.report_no,"
           +      			   "a.cmuse_id as trans_type,"
           +                   "a.cmuse_name as trans_type_name ,"
           +                   "b.cmuse_id as report_no,"
           +                   "b.cmuse_name as report_name ,"
           +                   "a.input_order,b.input_order "
  		   + " from wtt04_1d,cdshareno a,cdshareno b "
 		   + " where wtt04_1d.muser_id=? "  
   		   + " and a.cmuse_div='011' and a.cmuse_id = wtt04_1d.transfer_type"
   		   + " and b.cmuse_div in ('012','013','014')  "
   		   + " and a.identify_no = b.identify_no"
   		   + " and b.cmuse_name like wtt04_1d.report_no || '%' "
 		   + " order by a.input_order,b.input_order";
	paramList.add(muser_id) ;
	}
	List trans_typeList = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");
	
	List Report_upload = (List)request.getAttribute("Report_upload");		
	List Report_download = (List)request.getAttribute("Report_download");		
	List Report_edit = (List)request.getAttribute("Report_edit");		
	List Report_query = (List)request.getAttribute("Report_query");		  
	
	//取得ZZ004W的權限
	Properties permission = ( session.getAttribute("ZZ004W")==null ) ? new Properties() : (Properties)session.getAttribute("ZZ004W"); 
	if(permission == null){
       System.out.println("ZZ004W_List.permission == null");
    }else{
       System.out.println("ZZ004W_List.permission.size ="+permission.size());               
    }
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/ZZ004W.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<title>「<%=title%>使用者權限作業」<%=title_act%></title>
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
<table width="620" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  		<tr> 
   		 <td><img src="images/space_1.gif" width="12" height="12"></td>
  		</tr>
          <td bgcolor="#FFFFFF">
			<table width="620" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="620" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="110"><img src="images/banner_bg1.gif" width="110" height="17"></td>
                      <td width="380"><font color='#000000' size=4><b> 
                        <center>
                          「<%=title%>使用者權限作業」<%=title_act%>
                        </center>
                        </b></font> </td>
                      <td width="110"><img src="images/banner_bg1.gif" width="110" height="17"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><img src="images/space_1.gif" width="12" height="12"></td>
              </tr>
              <tr> 
                <td><table width="620" border="0" align="center" cellpadding="0" cellspacing="0">
               
                    <tr> 
                      <div align="right"><jsp:include page="getLoginUser.jsp?width=620" flush="true" /></div> 
                    </tr> 
                     <%
                      String nameColor="nameColor_sbody";
                      String textColor="textColor_sbody";
                      String bordercolor="#3A9D99";
                    %>                   
                    <tr> 
                      <td><table width=620 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="<%=bordercolor%>">
                         <tr class="sbody">                         
                          <td width='100%' colspan=4 class="list1Color_sbody_center"><b>[使用者資料]</b></td>						 						   							                          
                         </tr>  
                         
                         <tr class="sbody">
						  <td width='20%' class="<%=nameColor%>">使用者帳號/姓名</td>						  
						  <td width='80%' colspan=2 class="<%=textColor%>">						    
							<select name='MUSER_ID_Create' onchange="javascript:getData(this.document.forms[0],'<%=act%>','MUSER_ID_Create');">
						 	 <%for(int i=0;i<WTT01_hasCreateList.size();i++){%>
                             <option value="<%=(String)((DataObject)WTT01_hasCreateList.get(i)).getValue("muser_id")%>"
                             <% if(szid_Create.equals((String)((DataObject)WTT01_hasCreateList.get(i)).getValue("muser_id"))) out.print("selected");%>
                             >
                             <%=(String)((DataObject)WTT01_hasCreateList.get(i)).getValue("muser_id")%>/
                             <%=(String)((DataObject)WTT01_hasCreateList.get(i)).getValue("muser_name")%></option>                            
	                         <%}%>						
							</select>
                          </td>
                         
                         </tr>                              
                         
                         <tr class="sbody">
						  <td width='20%' class="<%=nameColor%>">機構類別</td>						  
						  <td width='80%' colspan=2 class="<%=textColor%>">						    							
						 	<input type='text' name='BANK_TYPE' value="<%if(WTT01_Create != null && WTT01_Create.size() != 0 && ((DataObject)WTT01_Create.get(0)).getValue("bank_type_name") != null ) out.print((String)((DataObject)WTT01_Create.get(0)).getValue("bank_type_name"));%>" readonly>                            	 
                          </td>                          
                         </tr>     
                         
                         <tr class="sbody">
						 
                          <td width='20%' class="<%=nameColor%>">總機構代碼</td>						  
						  <td width='80%' colspan=2 class="<%=textColor%>">						    							
						 	<input type='text' name='TBANK_NO' value="<%if(WTT01_Create != null && WTT01_Create.size() != 0 && ((DataObject)WTT01_Create.get(0)).getValue("tbank_no_name") != null ) out.print((String)((DataObject)WTT01_Create.get(0)).getValue("tbank_no_name"));%>" size="30" readonly>                            	 
                          </td>
                         </tr>  
                         
                         <tr class="sbody">						  
                          <td width='20%' class="<%=nameColor%>">組室代碼</td>						  
						  <td width='80%' colspan=2 class="<%=textColor%>">						    							
						 	<input type='text' name='BANK_NO' value="<%if(WTT01_Create != null && WTT01_Create.size() != 0 && ((DataObject)WTT01_Create.get(0)).getValue("bank_no_name") != null ) out.print((String)((DataObject)WTT01_Create.get(0)).getValue("bank_no_name"));%>" size="30" readonly>                            	 
                          </td>
                         </tr>  
                         
                         <tr class="sbody">						 
                          <td width='20%' class="<%=nameColor%>">科別</td>						  
						  <td width='80%' colspan=2 class="<%=textColor%>">						    							
						 	<input type='text' name='SUBDEP_ID' value="<%if(WTT01_Create != null && WTT01_Create.size() != 0 && ((DataObject)WTT01_Create.get(0)).getValue("subdep_id_name") != null ) out.print((String)((DataObject)WTT01_Create.get(0)).getValue("subdep_id_name"));%>" readonly>                            	 
                          </td>
                         </tr>  
                          
                         <tr class="sbody">
						  <td width='50%'colspan=4 class="textColor_sbody_center">
						    <%if(act.equals("del")){%>
						      <a href="javascript:doSubmit(this.document.forms[0],'Delete');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_deleteb.gif',1)"><img src="images/bt_delete.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>                       
						    <%}%>
						    <%if(act.equals("upd")){%>
						      <a href="javascript:doSubmit(this.document.forms[0],'Update');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_updateb.gif',1)"><img src="images/bt_update.gif" name="Image103" width="66" height="25" border="0" id="Image103"></a>                       
						    <%}%>
                            <a href="javascript:history.back();"  onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image102','','images/bt_backb.gif',1)"><img src="images/bt_back.gif" name="Image102" width="80" height="25" border="0" id="Image102"></a>
						  </td>						  						  				  						  
                         </tr>   
                         
                       </table></td>                  
     				   </tr>   
     				   
                       <tr> 
                       <td><table width=620 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="<%=bordercolor%>">                      
                           <tr class="sbody">
                            <td width='100%' colspan=12 class="list1Color_sbody"><b>基本功能</b>					 						  
 						    <a href="javascript:selectAll(this.document.forms[0],'program_id');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image104','','images/bt_selectallb.gif',1)"><img src="images/bt_selectall.gif" name="Image104" width="80" height="25" border="0" id="Image104"></a>                       
 						    <a href="javascript:selectNo(this.document.forms[0],'program_id');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_selectnob.gif',1)"><img src="images/bt_selectno.gif" name="Image105" width="80" height="25" border="0" id="Image105"></a>                        							 						     						 
                            </td>
                           </tr>     
                           <%                                            
                            String listTitle="listTitleColor_sbody"; 
                      		String list1Color="list1Color_sbody";
                      		String list2Color="list2Color_sbody";
                      		String listColor="list1Color_sbody";                         
                            List WTT03_1List = (List)request.getAttribute("WTT03_1List");
                            if(WTT03_1List != null){ %>
                          <input type="hidden" name="row_program" value="<%=WTT03_1List.size()%>">   
                       	  <tr class="<%=listTitle%>">                       	  
                            <td width="30">序號</td>                                   
                            <td width="100">程式代號</td>                                                                    				
            				<td width="200">程式名稱</td>
            				<td width="25">新增</td>            				            				
            				<td width="25">刪除</td>            				            				
            				<td width="25">修改</td>            				            				
            				<td width="25">查詢</td>            				            				
            				<td width="25">印表</td>            				            				
            				<td width="25">上傳</td>            				            				
            				<td width="25">下載</td>            				            				
            				<td width="25">鎖定</td>            				            				
            				<td width="25">申報</td>            				            				
					      </tr>                      		    
                      <%
                      	  for(int i=0;i<WTT03_1List.size();i++){    
                             listColor = (i % 2 == 0)?list2Color:list1Color;	                                                         
                      %>    
                          <tr class="<%=listColor%>">
                            <input type="hidden" name="program_id_<%=(i+1)%>" value="<%if( ((DataObject)WTT03_1List.get(i)).getValue("program_id") != null ) out.print((String)((DataObject)WTT03_1List.get(i)).getValue("program_id"));%>">
                            <td width="30"><%=i+1%></td>                       				            				            				
            				<td width="100"><%if( ((DataObject)WTT03_1List.get(i)).getValue("program_id") != null ) out.print((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")); else out.print("&nbsp;");%></td>
            				<td width="200"><%if( ((DataObject)WTT03_1List.get(i)).getValue("program_name") != null ) out.print((String)((DataObject)WTT03_1List.get(i)).getValue("program_name")); else out.print("&nbsp;");%></td>            				            				            				
            				<td width="25">            				
            				<input type="checkbox" name="P_ADD_<%=(i+1)%>" 
            				<%if(act.equals("del")) {            				            				     
            				     if( ((DataObject)WTT03_1List.get(i)).getValue("p_add") != null && (!((String)((DataObject)WTT03_1List.get(i)).getValue("p_add")).equals("Y")) ){
            				         out.print("disabled");
            				     }else{   
            				         out.print("checked disabled");//94.03.08 fix act=del把有的權限勾選,並disabled
            				     }
            				  }   
            				  if(act.equals("upd")) {            				
            				     if( ((DataObject)WTT03_1List.get(i)).getValue("p_add_o") != null && (!((String)((DataObject)WTT03_1List.get(i)).getValue("p_add_o")).equals("Y")) ) out.print("disabled");
            				     if( ((DataObject)WTT03_1List.get(i)).getValue("p_add") != null && ((String)((DataObject)WTT03_1List.get(i)).getValue("p_add")).equals("Y") ) out.print("checked");            				     
            				     //super user不可指派ZZ003W.ZZ005W給系統管理者
            				     //系統管理者不可指派ZZ001W.ZZ002W.ZZ004W給一般使用者            				     
            				     if( muser_type.equals("S") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ003W")) out.print(" disabled");
            					 if( muser_type.equals("S") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ005W")) out.print(" disabled");
            					 if( muser_type.equals("A") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ001W")) out.print(" disabled");
            					 if( muser_type.equals("A") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ002W")) out.print(" disabled");
            					 if( muser_type.equals("A") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ004W")) out.print(" disabled");            				            		
            				  }     
            				%>            				
            				>
            				</td>  
            				<td width="25">
            				<input type="checkbox" name="P_DELETE_<%=(i+1)%>" 
            				<%if(act.equals("del")) {            				            				     
            				     if( ((DataObject)WTT03_1List.get(i)).getValue("p_delete") != null && (!((String)((DataObject)WTT03_1List.get(i)).getValue("p_delete")).equals("Y")) ){
            				         out.print("disabled");
            				     }else{   
            				         out.print("checked disabled");//94.03.08 fix act=del把有的權限勾選,並disabled
            				     }   
            				  }   
            				  if(act.equals("upd")) {            				
            				     if( ((DataObject)WTT03_1List.get(i)).getValue("p_delete_o") != null && (!((String)((DataObject)WTT03_1List.get(i)).getValue("p_delete_o")).equals("Y")) ) out.print("disabled");
            				     if( ((DataObject)WTT03_1List.get(i)).getValue("p_delete") != null && ((String)((DataObject)WTT03_1List.get(i)).getValue("p_delete")).equals("Y") ) out.print("checked");
            				     //super user不可指派ZZ003W.ZZ005W給系統管理者
            				     //系統管理者不可指派ZZ001W.ZZ002W.ZZ004W給一般使用者            				     
            				     if( muser_type.equals("S") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ003W")) out.print(" disabled");
            					 if( muser_type.equals("S") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ005W")) out.print(" disabled");
            					 if( muser_type.equals("A") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ001W")) out.print(" disabled");
            					 if( muser_type.equals("A") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ002W")) out.print(" disabled");
            					 if( muser_type.equals("A") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ004W")) out.print(" disabled");            				            		
            				  }     
            				%>              				
            				>
            				
            				</td>  
            				<td width="25">
            				<input type="checkbox" name="P_UPDATE_<%=(i+1)%>" 
            				<%if(act.equals("del")) {            				            				     
            				     if( ((DataObject)WTT03_1List.get(i)).getValue("p_update") != null && (!((String)((DataObject)WTT03_1List.get(i)).getValue("p_update")).equals("Y")) ){
            				         out.print("disabled");
            				     }else{   
            				         out.print("checked disabled");//94.03.08 fix act=del把有的權限勾選,並disabled
            				     }
            				  }   
            				  if(act.equals("upd")) {            				
            				     if( ((DataObject)WTT03_1List.get(i)).getValue("p_update_o") != null && (!((String)((DataObject)WTT03_1List.get(i)).getValue("p_update_o")).equals("Y")) ) out.print("disabled");
            				     if( ((DataObject)WTT03_1List.get(i)).getValue("p_update") != null && ((String)((DataObject)WTT03_1List.get(i)).getValue("p_update")).equals("Y") ) out.print("checked");
            				     //super user不可指派ZZ003W.ZZ005W給系統管理者
            				     //系統管理者不可指派ZZ001W.ZZ002W.ZZ004W給一般使用者            				     
            				     if( muser_type.equals("S") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ003W")) out.print(" disabled");
            					 if( muser_type.equals("S") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ005W")) out.print(" disabled");
            					 if( muser_type.equals("A") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ001W")) out.print(" disabled");
            					 if( muser_type.equals("A") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ002W")) out.print(" disabled");
            					 if( muser_type.equals("A") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ004W")) out.print(" disabled");            				            		
            				  }     
            				%>           				
            				>
            				</td>  
            				<td width="25">
            				<input type="checkbox" name="P_QUERY_<%=(i+1)%>" 
            				<%if(act.equals("del")) {            				            				     
            				     if( ((DataObject)WTT03_1List.get(i)).getValue("p_query") != null && (!((String)((DataObject)WTT03_1List.get(i)).getValue("p_query")).equals("Y")) ){
            				         out.print("disabled");
            				     }else{   
            				         out.print("checked disabled");//94.03.08 fix act=del把有的權限勾選,並disabled
            				     }
            				  }   
            				  if(act.equals("upd")) {            				
            				     if( ((DataObject)WTT03_1List.get(i)).getValue("p_query_o") != null && (!((String)((DataObject)WTT03_1List.get(i)).getValue("p_query_o")).equals("Y")) ) out.print("disabled");
            				     if( ((DataObject)WTT03_1List.get(i)).getValue("p_query") != null && ((String)((DataObject)WTT03_1List.get(i)).getValue("p_query")).equals("Y") ) out.print("checked");
            				     //super user不可指派ZZ003W.ZZ005W給系統管理者
            				     //系統管理者不可指派ZZ001W.ZZ002W.ZZ004W給一般使用者            				     
            				     if( muser_type.equals("S") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ003W")) out.print(" disabled");
            					 if( muser_type.equals("S") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ005W")) out.print(" disabled");
            					 if( muser_type.equals("A") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ001W")) out.print(" disabled");
            					 if( muser_type.equals("A") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ002W")) out.print(" disabled");
            					 if( muser_type.equals("A") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ004W")) out.print(" disabled");            				            		
            				  }     
            				%>              				
            				>
            				</td>  
            				<td width="25">
            				<input type="checkbox" name="P_PRINT_<%=(i+1)%>" 
            				<%if(act.equals("del")) {            				            				     
            				     if( ((DataObject)WTT03_1List.get(i)).getValue("p_print") != null && (!((String)((DataObject)WTT03_1List.get(i)).getValue("p_print")).equals("Y")) ){
            				         out.print("disabled");
            				     }else{   
            				         out.print("checked disabled");//94.03.08 fix act=del把有的權限勾選,並disabled
            				     }   
            				  }   
            				  if(act.equals("upd")) {            				
            				     if( ((DataObject)WTT03_1List.get(i)).getValue("p_print_o") != null && (!((String)((DataObject)WTT03_1List.get(i)).getValue("p_print_o")).equals("Y")) ) out.print("disabled");
            				     if( ((DataObject)WTT03_1List.get(i)).getValue("p_print") != null && ((String)((DataObject)WTT03_1List.get(i)).getValue("p_print")).equals("Y") ) out.print("checked");
            				     //super user不可指派ZZ003W.ZZ005W給系統管理者
            				     //系統管理者不可指派ZZ001W.ZZ002W.ZZ004W給一般使用者            				     
            				     if( muser_type.equals("S") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ003W")) out.print(" disabled");
            					 if( muser_type.equals("S") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ005W")) out.print(" disabled");
            					 if( muser_type.equals("A") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ001W")) out.print(" disabled");
            					 if( muser_type.equals("A") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ002W")) out.print(" disabled");
            					 if( muser_type.equals("A") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ004W")) out.print(" disabled");            				            		
            				  }     
            				%>            				
            				>
            				</td>  
            				<td width="25">
            				<input type="checkbox" name="P_UPLOAD_<%=(i+1)%>" 
            				<%if(act.equals("del")) {            				            				     
            				     if( ((DataObject)WTT03_1List.get(i)).getValue("p_upload") != null && (!((String)((DataObject)WTT03_1List.get(i)).getValue("p_upload")).equals("Y")) ){
            				        out.print("disabled");
            				     }else{   
            				         out.print("checked disabled");//94.03.08 fix act=del把有的權限勾選,並disabled
            				     }   
            				  }   
            				  if(act.equals("upd")) {            				
            				     if( ((DataObject)WTT03_1List.get(i)).getValue("p_upload_o") != null && (!((String)((DataObject)WTT03_1List.get(i)).getValue("p_upload_o")).equals("Y")) ) out.print("disabled");
            				     if( ((DataObject)WTT03_1List.get(i)).getValue("p_upload") != null && ((String)((DataObject)WTT03_1List.get(i)).getValue("p_upload")).equals("Y") ) out.print("checked");
            				     //super user不可指派ZZ003W.ZZ005W給系統管理者
            				     //系統管理者不可指派ZZ001W.ZZ002W.ZZ004W給一般使用者            				     
            				     if( muser_type.equals("S") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ003W")) out.print(" disabled");
            					 if( muser_type.equals("S") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ005W")) out.print(" disabled");
            					 if( muser_type.equals("A") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ001W")) out.print(" disabled");
            					 if( muser_type.equals("A") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ002W")) out.print(" disabled");
            					 if( muser_type.equals("A") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ004W")) out.print(" disabled");            				            		
            				  }     
            				%>             				
            				>
            				</td>  
            				<td width="25">
            				<input type="checkbox" name="P_DOWNLOAD_<%=(i+1)%>" 
            				<%if(act.equals("del")) {            				            				     
            				     if( ((DataObject)WTT03_1List.get(i)).getValue("p_download") != null && (!((String)((DataObject)WTT03_1List.get(i)).getValue("p_download")).equals("Y")) ){
            				         out.print("disabled");
            				     }else{   
            				         out.print("checked disabled");//94.03.08 fix act=del把有的權限勾選,並disabled
            				     }
            				  }   
            				  if(act.equals("upd")) {            				
            				     if( ((DataObject)WTT03_1List.get(i)).getValue("p_download_o") != null && (!((String)((DataObject)WTT03_1List.get(i)).getValue("p_download_o")).equals("Y")) ) out.print("disabled");
            				     if( ((DataObject)WTT03_1List.get(i)).getValue("p_download") != null && ((String)((DataObject)WTT03_1List.get(i)).getValue("p_download")).equals("Y") ) out.print("checked");
            				     //super user不可指派ZZ003W.ZZ005W給系統管理者
            				     //系統管理者不可指派ZZ001W.ZZ002W.ZZ004W給一般使用者            				     
            				     if( muser_type.equals("S") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ003W")) out.print(" disabled");
            					 if( muser_type.equals("S") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ005W")) out.print(" disabled");
            					 if( muser_type.equals("A") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ001W")) out.print(" disabled");
            					 if( muser_type.equals("A") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ002W")) out.print(" disabled");
            					 if( muser_type.equals("A") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ004W")) out.print(" disabled");            				            		
            				  }     
            				%>             				          				
            				>
            				</td>  
            				<td width="25">
            				<input type="checkbox" name="P_LOCK_<%=(i+1)%>" 
            				<%if(act.equals("del")) {            				            				     
            				     if( ((DataObject)WTT03_1List.get(i)).getValue("p_lock") != null && (!((String)((DataObject)WTT03_1List.get(i)).getValue("p_lock")).equals("Y")) ){
            				         out.print("disabled");
            				     }else{   
            				         out.print("checked disabled");//94.03.08 fix act=del把有的權限勾選,並disabled
            				     }
            				  }   
            				  if(act.equals("upd")) {            				
            				     if( ((DataObject)WTT03_1List.get(i)).getValue("p_lock_o") != null && (!((String)((DataObject)WTT03_1List.get(i)).getValue("p_lock_o")).equals("Y")) ) out.print("disabled");
            				     if( ((DataObject)WTT03_1List.get(i)).getValue("p_lock") != null && ((String)((DataObject)WTT03_1List.get(i)).getValue("p_lock")).equals("Y") ) out.print("checked");
            				     //super user不可指派ZZ003W.ZZ005W給系統管理者
            				     //系統管理者不可指派ZZ001W.ZZ002W.ZZ004W給一般使用者            				     
            				     if( muser_type.equals("S") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ003W")) out.print(" disabled");
            					 if( muser_type.equals("S") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ005W")) out.print(" disabled");
            					 if( muser_type.equals("A") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ001W")) out.print(" disabled");
            					 if( muser_type.equals("A") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ002W")) out.print(" disabled");
            					 if( muser_type.equals("A") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ004W")) out.print(" disabled");            				            		
            				  }     
            				%>             				
            				>
            				</td>  
            				<td width="25">
            				<input type="checkbox" name="P_OTHER_<%=(i+1)%>" 
            				<%if(act.equals("del")) {            				            				     
            				     if( ((DataObject)WTT03_1List.get(i)).getValue("p_other") != null && (!((String)((DataObject)WTT03_1List.get(i)).getValue("p_other")).equals("Y")) ){
            				          out.print("disabled");
            				     }else{   
            				         out.print("checked disabled");//94.03.08 fix act=del把有的權限勾選,並disabled
            				     }    
            				  }   
            				  if(act.equals("upd")) {            				
            				     if( ((DataObject)WTT03_1List.get(i)).getValue("p_other_o") != null && (!((String)((DataObject)WTT03_1List.get(i)).getValue("p_other_o")).equals("Y")) ) out.print("disabled");
            				     if( ((DataObject)WTT03_1List.get(i)).getValue("p_other") != null && ((String)((DataObject)WTT03_1List.get(i)).getValue("p_other")).equals("Y") ) out.print("checked");
            				     //super user不可指派ZZ003W.ZZ005W給系統管理者
            				     //系統管理者不可指派ZZ001W.ZZ002W.ZZ004W給一般使用者            				     
            				     if( muser_type.equals("S") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ003W")) out.print(" disabled");
            					 if( muser_type.equals("S") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ005W")) out.print(" disabled");
            					 if( muser_type.equals("A") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ001W")) out.print(" disabled");
            					 if( muser_type.equals("A") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ002W")) out.print(" disabled");
            					 if( muser_type.equals("A") && ((String)((DataObject)WTT03_1List.get(i)).getValue("program_id")).equals("ZZ004W")) out.print(" disabled");            				            		
            				  }     
            				%>             				
            				>
            				</td>  
					      </tr> 					      
					   <%					     
					     }//end of for WTT03_1List
	                  	 }//end of if
    			       %>  
    			       </table></td>                  
     				   </tr> 
					   <%if(!szMenu.equals("MIS")){%>	     				   
     				   <tr> 
                       <td><table width=620 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="<%=bordercolor%>">                                                 
                         <%List uploadList = (List)request.getAttribute("uploadList");%>
    			         <tr class="sbody">
                            <td width='100%' colspan=12 class="list1Color_sbody"><b>申報(上傳)</b>					 						  
 						    <a href="javascript:selectAll(this.document.forms[0],'upload');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image106','','images/bt_selectallb.gif',1)"><img src="images/bt_selectall.gif" name="Image106" width="80" height="25" border="0" id="Image106"></a>                       
 						    <a href="javascript:selectNo(this.document.forms[0],'upload');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image107','','images/bt_selectnob.gif',1)"><img src="images/bt_selectno.gif" name="Image107" width="80" height="25" border="0" id="Image107"></a>                        							 						     						 
                            </td>
                         </tr>     
                           <%                                            
                                                            
                      if( Report_upload != null){ %>                          
                       	  <tr class="<%=listTitle%>">                       	  
                            <td width="30">序號</td>                                   
                            <td width="30">選項</td>                                   
                            <td width="60">傳輸類別</td>                                                                    				
            				<td width="220">傳輸類別名稱</td>
            				<td width="60">項目代號</td>            				            				
            				<td width="220">項目名稱</td>            				            				            				
					      </tr>                      		    
                      <% for(int i=0;i<Report_upload.size();i++){    
                           listColor = (i % 2 == 0)?list2Color:list1Color;
                      %>	    
                      	  <input type="hidden" name="row_upload" value="<%=Report_upload.size()%>">   
                      	  <tr class="<%=listColor%>">
                            <td width="30"><%=i+1%></td>                       				            				            				
                            <td widht="30">
                            <input type="checkbox" name="upload_isModify_<%=(i+1)%>" value="" 
                            <%if(act.equals("upd") && ((DataObject)Report_upload.get(i)).getValue("report_no") != null){ 
                                 out.print("checked");
                              }
                              //94.03.08 fix act=del把有的權限勾選,並disabled
                              if(act.equals("del") && ((DataObject)Report_upload.get(i)).getValue("report_no") != null){ 
                                 out.print("checked disabled");
                              }
                            %>                            
                            >
                            </td>
            				<td width="80">
            				<input type="hidden" name="uploadData_trans_type_<%=(i+1)%>" value="<%if( ((DataObject)Report_upload.get(i)).getValue("trans_type") != null ) out.print((String)((DataObject)Report_upload.get(i)).getValue("trans_type"));%>">
            				<%if( ((DataObject)Report_upload.get(i)).getValue("trans_type") != null ) out.print((String)((DataObject)Report_upload.get(i)).getValue("trans_type")); else out.print("&nbsp;");%>
            				</td>
            				<td width="220">            				
            				<%if( ((DataObject)Report_upload.get(i)).getValue("trans_type_name") != null ) out.print((String)((DataObject)Report_upload.get(i)).getValue("trans_type_name")); else out.print("&nbsp;");%>
            				</td>            				            				            				
            				<td width="80">
            				<input type="hidden" name="uploadData_report_no_<%=(i+1)%>" value="<%if( ((DataObject)Report_upload.get(i)).getValue("report_no_o") != null ) out.print((String)((DataObject)Report_upload.get(i)).getValue("report_no_o"));%>">
            				<%if( ((DataObject)Report_upload.get(i)).getValue("report_no_o") != null ) out.print((String)((DataObject)Report_upload.get(i)).getValue("report_no_o")); else out.print("&nbsp;");%>
            				</td>
            				<td width="220">            				
            				<%if( ((DataObject)Report_upload.get(i)).getValue("report_name") != null ) out.print((String)((DataObject)Report_upload.get(i)).getValue("report_name")); else out.print("&nbsp;");%>
            				</td>            				            				            				            				
					      </tr> 						      
					  <%    	
                      	    }//end of Report_upload                      	  
	                  	 }//end of if
    			       %>  
    			       
    			       </table></td>                  
     				   </tr>  
     				   
     				   <td><table width=620 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="<%=bordercolor%>">                                                 
    			       <tr class="sbody">
    			       <%List downloadList = (List)request.getAttribute("downloadList");%>
                           <td width='100%' colspan=12 bgcolor='D2F0FF'><b>申報(下載)</b>					 						  
 						    <a href="javascript:selectAll(this.document.forms[0],'download');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image107_1','','images/bt_selectallb.gif',1)"><img src="images/bt_selectall.gif" name="Image107_1" width="80" height="25" border="0" id="Image107_1"></a>                       
 						    <a href="javascript:selectNo(this.document.forms[0],'download');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image108','','images/bt_selectnob.gif',1)"><img src="images/bt_selectno.gif" name="Image108" width="80" height="25" border="0" id="Image108"></a>                        							 						     						 
                           </td>
                       </tr>     
                       <%if(Report_download != null){ %>                          
                       	  <tr class="<%=listColor%>">                       	  
                            <td width="30">序號</td>                                   
                            <td width="30">選項</td>                                   
                            <td width="60">傳輸類別</td>                                                                    				
            				<td width="220">傳輸類別名稱</td>
            				<td width="60">項目代號</td>            				            				
            				<td width="220">項目名稱</td>            				            				            				
					      </tr>                      		    
					      <% for(int i=0;i<Report_download.size();i++){    
                      	         listColor = (i % 2 == 0)?list2Color:list1Color;	                                                      
                          %>    
                          <input type="hidden" name="row_download" value="<%=Report_download.size()%>">   
                          <tr class="<%=listColor%>">
                            <td width="30"><%=i+1%></td>                       				            				            				
                            <td widht="30">
                            <input type="checkbox" name="download_isModify_<%=(i+1)%>" value="" 
                            <%if(act.equals("upd") && ((DataObject)Report_download.get(i)).getValue("report_no") != null){ 
                                 out.print("checked");
                              }
                              //94.03.08 fix act=del把有的權限勾選,並disabled
                              if(act.equals("del") && ((DataObject)Report_download.get(i)).getValue("report_no") != null){ 
                                 out.print("checked disabled");
                              }
                            %>     
                            >
                            </td>
            				<td width="80">
            				<input type="hidden" name="downloadData_trans_type_<%=(i+1)%>" value="<%if( ((DataObject)Report_download.get(i)).getValue("trans_type") != null ) out.print((String)((DataObject)Report_download.get(i)).getValue("trans_type"));%>">
            				<%if( ((DataObject)Report_download.get(i)).getValue("trans_type") != null ) out.print((String)((DataObject)Report_download.get(i)).getValue("trans_type")); else out.print("&nbsp;");%>
            				</td>
            				<td width="220"><%if( ((DataObject)Report_download.get(i)).getValue("trans_type_name") != null ) out.print((String)((DataObject)Report_download.get(i)).getValue("trans_type_name")); else out.print("&nbsp;");%></td>            				            				            				
            				<td width="80">
            				<input type="hidden" name="downloadData_report_no_<%=(i+1)%>" value="<%if( ((DataObject)Report_download.get(i)).getValue("report_no_o") != null ) out.print((String)((DataObject)Report_download.get(i)).getValue("report_no_o"));%>">
            				<%if( ((DataObject)Report_download.get(i)).getValue("report_no_o") != null ) out.print((String)((DataObject)Report_download.get(i)).getValue("report_no_o")); else out.print("&nbsp;");%>
            				</td>
            				<td width="220"><%if( ((DataObject)Report_download.get(i)).getValue("report_name") != null ) out.print((String)((DataObject)Report_download.get(i)).getValue("report_name")); else out.print("&nbsp;");%></td>            				            				            				
            				
					      </tr> 		
                      	  <%
							 }//end ofReport_download                       	                        	  					        	                  	  
    			          }//enf of have date%>  
    			        </table></td>                  
     				   </tr>  
    			       
    			       <tr> 
                       <td><table width=620 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="<%=bordercolor%>">                                                 
                         <%List editList = (List)request.getAttribute("editList");%>
    			         <tr class="sbody">
                            <td width='100%' colspan=12 bgcolor='D2F0FF'><b>申報(編輯)</b>					 						  
 						    <a href="javascript:selectAll(this.document.forms[0],'edit');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image109','','images/bt_selectallb.gif',1)"><img src="images/bt_selectall.gif" name="Image109" width="80" height="25" border="0" id="Image109"></a>                       
 						    <a href="javascript:selectNo(this.document.forms[0],'edit');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image110','','images/bt_selectnob.gif',1)"><img src="images/bt_selectno.gif" name="Image110" width="80" height="25" border="0" id="Image110"></a>                        							 						     						 
                            </td>
                         </tr>     
                      <%if(Report_edit != null){ %>                                                    
                       	  <tr class="<%=listTitle%>">                       	  
                            <td width="30">序號</td>                                   
                            <td width="30">選項</td>                                   
                            <td width="60">傳輸類別</td>                                                                    				
            				<td width="220">傳輸類別名稱</td>
            				<td width="60">項目代號</td>            				            				
            				<td width="220">項目名稱</td>            				            				            				
					      </tr> 
					      <%                      
                      	     for(int i=0;i<Report_edit.size();i++){    
                      	         listColor = (i % 2 == 0)?list2Color:list1Color;
                      	  %>
                      	  <input type="hidden" name="row_edit" value="<%=Report_edit.size()%>">          
                      	  <tr class="<%=listColor%>">
                            <td width="30"><%=i+1%></td>                       				            				            				
                            <td widht="30">
                            <input type="checkbox" name="edit_isModify_<%=(i+1)%>" value="" 
                            <%if(act.equals("upd") && ((DataObject)Report_edit.get(i)).getValue("report_no") != null){ 
                                 out.print("checked");
                              }
                              //94.03.08 fix act=del把有的權限勾選,並disabled
                              if(act.equals("del") && ((DataObject)Report_edit.get(i)).getValue("report_no") != null){ 
                                 out.print("checked disabled");
                              }
                            %>     
                            >
                            </td>
            				<td width="80">
            				<input type="hidden" name="editData_trans_type_<%=(i+1)%>" value="<%if( ((DataObject)Report_edit.get(i)).getValue("trans_type") != null ) out.print((String)((DataObject)Report_edit.get(i)).getValue("trans_type"));%>">
            				<%if( ((DataObject)Report_edit.get(i)).getValue("trans_type") != null ) out.print((String)((DataObject)Report_edit.get(i)).getValue("trans_type")); else out.print("&nbsp;");%>
            				</td>
            				<td width="220"><%if( ((DataObject)Report_edit.get(i)).getValue("trans_type_name") != null ) out.print((String)((DataObject)Report_edit.get(i)).getValue("trans_type_name")); else out.print("&nbsp;");%></td>            				            				            				
            				<td width="80">
            				<input type="hidden" name="editData_report_no_<%=(i+1)%>" value="<%if( ((DataObject)Report_edit.get(i)).getValue("report_no_o") != null ) out.print((String)((DataObject)Report_edit.get(i)).getValue("report_no_o"));%>">
            				<%if( ((DataObject)Report_edit.get(i)).getValue("report_no_o") != null ) out.print((String)((DataObject)Report_edit.get(i)).getValue("report_no_o")); else out.print("&nbsp;");%>
            				</td>
            				<td width="220"><%if( ((DataObject)Report_edit.get(i)).getValue("report_name") != null ) out.print((String)((DataObject)Report_edit.get(i)).getValue("report_name")); else out.print("&nbsp;");%></td>            				            				            				            				
					      </tr> 		                                                   
                          <% }//end of Report_edit
                           }//end of have data	                  	  
    			          %>  
    			        </table></td>                  
     				   </tr>  
    			       <tr> 
                       <td><table width=620 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="<%=bordercolor%>">                                                 
                         <%List queryList = (List)request.getAttribute("queryList");%>
    			         <tr class="sbody">
                            <td width='100%' colspan=12 class="list1Color_sbody"><b>申報(查詢)</b>					 						  
 						    <a href="javascript:selectAll(this.document.forms[0],'query');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image111','','images/bt_selectallb.gif',1)"><img src="images/bt_selectall.gif" name="Image111" width="80" height="25" border="0" id="Image111"></a>                       
 						    <a href="javascript:selectNo(this.document.forms[0],'query');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image112','','images/bt_selectnob.gif',1)"><img src="images/bt_selectno.gif" name="Image112" width="80" height="25" border="0" id="Image112"></a>                        							 						     						 
                            </td>
                         </tr>     
                         <%if(Report_query != null){ %>                                                                              
                       	  <tr class="<%=listTitle%>">                       	  
                            <td width="30">序號</td>                                   
                            <td width="30">選項</td>                                   
                            <td width="60">傳輸類別</td>                                                                    				
            				<td width="220">傳輸類別名稱</td>
            				<td width="60">項目代號</td>            				            				
            				<td width="220">項目名稱</td>            				            				            				
					      </tr>   
					      <%                      	 
                      	     for(int i=0;i<Report_query.size();i++){    
                      	        listColor = (i % 2 == 0)?list2Color:list1Color;	
                      	  %>                   		    
                      	  <input type="hidden" name="row_query" value="<%=Report_query.size()%>">       
                      	  <tr class="<%=listColor%>">
                            <td width="30"><%=i+1%></td>                       				            				            				
                            <td widht="30">
                            <input type="checkbox" name="query_isModify_<%=(i+1)%>" value="" 
                            <%if(act.equals("upd") && ((DataObject)Report_query.get(i)).getValue("report_no") != null){ 
                                 out.print("checked");
                              }
                              //94.03.08 fix act=del把有的權限勾選,並disabled
                              if(act.equals("del") && ((DataObject)Report_query.get(i)).getValue("report_no") != null){ 
                                 out.print("checked disabled");
                              }
                            %>  
                            >
                            </td>
            				<td width="80">
            				<input type="hidden" name="queryData_trans_type_<%=(i+1)%>" value="<%if( ((DataObject)Report_query.get(i)).getValue("trans_type") != null ) out.print((String)((DataObject)Report_query.get(i)).getValue("trans_type"));%>">
            				<%if( ((DataObject)Report_query.get(i)).getValue("trans_type") != null ) out.print((String)((DataObject)Report_query.get(i)).getValue("trans_type")); else out.print("&nbsp;");%>
            				</td>
            				<td width="220"><%if( ((DataObject)Report_query.get(i)).getValue("trans_type_name") != null ) out.print((String)((DataObject)Report_query.get(i)).getValue("trans_type_name")); else out.print("&nbsp;");%></td>            				            				            				
            				<td width="80">
            				<input type="hidden" name="queryData_report_no_<%=(i+1)%>" value="<%if( ((DataObject)Report_query.get(i)).getValue("report_no_o") != null ) out.print((String)((DataObject)Report_query.get(i)).getValue("report_no_o"));%>">
            				<%if( ((DataObject)Report_query.get(i)).getValue("report_no_o") != null ) out.print((String)((DataObject)Report_query.get(i)).getValue("report_no_o")); else out.print("&nbsp;");%>
            				</td>
            				<td width="220"><%if( ((DataObject)Report_query.get(i)).getValue("report_name") != null ) out.print((String)((DataObject)Report_query.get(i)).getValue("report_name")); else out.print("&nbsp;");%></td>            				            				            				
            				
					      </tr> 	
                      	         
                          <% }//end of Report_query                          
    			          }//end of have data
    			          %>  
    			        </table></td>                  
     				   </tr>  
    			      <%}//end of szMenu%>
    			       
    			       
     					  </table></td>                  
     					 </tr>  
                         
                          
                        </Table></td>
                    </tr>                 
                    <tr>                  
                <td><div align="right"><jsp:include page="getMaintainUser.jsp?width=620" flush="true" /></div></td>                                              
              </tr>
              
              <tr> 
                <td><div align="center"> 
                    <table width="243" border="0" cellpadding="1" cellspacing="1">
                      <tr>              
                       
                      </tr>
                      </tr>
                    </table>
                  </div></td>
              </tr>                               
      </table></td>
  </tr>
 
</form>
</body>
</html>
