<%
//94.12.02 fix 農漁會地方主管第二層list by lilic0c0 4183
//93.12.22 fix 農會/漁會無已裁撤之相關機構
//94.01.12 fix 若登入者為農漁會共用中心,則農會.漁會只顯示參加該共用中心的機構代碼 by 2295
//94.03.30 fix 農會/漁會顯示已裁撤之相關機構 by 2295
//         fix 修正FX004W url by 2295
//94.04.12 fix 按bank_no sort by 2295
//95.08.18 add 警示帳戶.地方主管機關進來.可看到其所屬的農漁會.農漁會進來只能看到自己的 by 2295
//95.08.24 add 承受擔保品.警示帳戶.地方主管機關進來.可看到其所屬的農漁會.並出現列印按鈕 by 2495
//99.09.23 fix 根據查詢年度.100年以後取得新機構名稱.100年以前取得舊機構名稱
// 			   使用PreparedStatement;並列印轉換後的SQL;套用QueryDB_SQLParam by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@include file="common.jsp"%>
<%
   RequestDispatcher rd = null;
   String actMsg="";
   String muser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		
   String muser_name = ( session.getAttribute("muser_name")==null ) ? "" : (String)session.getAttribute("muser_name");		
   String muser_type = ( session.getAttribute("muser_type")==null ) ? "" : (String)session.getAttribute("muser_type");			
   String muser_bank_type = ( session.getAttribute("bank_type")==null ) ? "" : (String)session.getAttribute("bank_type");			
   String muser_tbank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");			
   
   String program_name = "";   
   String program_url = "";   
   StringBuffer sqlCmd = new StringBuffer();
   StringBuffer sqlCmd_normal = new StringBuffer();     
   StringBuffer sqlCmd_revoke = new StringBuffer();
   List paramList = new ArrayList();
   List paramList_normal = new ArrayList();
   List paramList_revoke = new ArrayList();
   boolean Banklist_flag =false;//判斷這是進入第幾層 bank_list( true為第二層 false 為第一層)
   Map dataMap =Utility.saveSearchParameter(request);
   String bank_type = Utility.getTrimString(dataMap.get("bank_type"));
   String link_name = Utility.getTrimString(dataMap.get("link_name"));
   String list_type = Utility.getTrimString(dataMap.get("list_type"));
   String s_year = Utility.getCHTYYMMDD("yy");
   //99.09.23 add 查詢年度100年以前.縣市別不同===============================
   String cd01_table = (Integer.parseInt(s_year) < 100)?"cd01_99":"";
   String wlx01_m_year = (Integer.parseInt(s_year) < 100)?"99":"100";
  //=====================================================================
   DataObject bean = null;
   
   if(!bank_type.equals("")){
	   session.setAttribute("nowbank_type",bank_type);//將已點選的bank_type寫入session	   
   }   
   
   //94.01.12 fix 農會.漁會只顯示參加該共用中心的機構代碼============================
   //94.01.14 fix 農會.漁會只顯示參加該地方主管機關的機構代碼============================
   sqlCmd.append("select bn01.bank_no,bn01.bank_name,bn01.update_date,wlx01.cancel_date from (select * from bn01 where m_year=?)bn01 ");
   sqlCmd.append("left join (select * from wlx01 where m_year=?)wlx01 on bn01.bank_no = WLX01.bank_no");      
   sqlCmd_normal.append(sqlCmd);
   sqlCmd_revoke.append(sqlCmd);
   
   paramList_normal.add(wlx01_m_year);
   paramList_normal.add(wlx01_m_year);
   
   paramList_revoke.add(wlx01_m_year);
   paramList_revoke.add(wlx01_m_year);
   
   
   sqlCmd_normal.append(" where bn01.bn_type <> '2'"); 
   sqlCmd_revoke.append(" where bn01.bn_type = '2'"); 
   
   //94.01.12 fix 農會.漁會只顯示參加該共用中心的機構代碼============================	       
   if(muser_bank_type.equals("8") && (bank_type.equals("6") || bank_type.equals("7"))){//農漁會共用中心	       
      sqlCmd_normal.append(" and WLX01.center_no = ?"); 
      sqlCmd_normal.append(" and bn01.bank_type=?");
      paramList_normal.add(muser_tbank_no);
      paramList_normal.add(bank_type);
      sqlCmd_revoke.append(" and WLX01.center_no = ?");
      sqlCmd_revoke.append(" and bn01.bank_type=?"); 
      paramList_revoke.add(muser_tbank_no);
      paramList_revoke.add(bank_type);
   }   
   
   //94.01.14 fix 農會.漁會只顯示參加該地方主管機關的機構代碼============================	       
   else if(muser_bank_type.equals("B") && (bank_type.equals("6") || bank_type.equals("7"))){//如果是地方主管機關  
      sqlCmd_normal.append(" and WLX01.m2_name = ?"); 
      paramList_normal.add(muser_tbank_no);
      sqlCmd_revoke.append(" and WLX01.m2_name = ?"); 
      paramList_revoke.add(muser_tbank_no);
   }
   //94.12.02 fix 農會.漁會只顯示參加該地方主管機關的機構代碼============================
   //如果是地方主管機關 從地方主管機關的選項來 
   //或是特權使用者	從地方主管機關的選項來        
   else if(bank_type.equals("B")){
      //如果進入第二層bank_list則去取第一層所選地方主管機關的bank_no
      //如果沒有進入第二層 則去取使用者的bank_no來當now_bank_no
      
      String now_bank_no = ( request.getParameter("tbank_no")==null ) ? muser_tbank_no : (String)request.getParameter("tbank_no");
      
      //判斷是地方主管還是上級主管
      //Banklist_flag表示這是第二次進入banklist
      //如果可以第二次進入表示登入者是一級主管
      if(muser_bank_type.equals("B")||(request.getParameter("tbank_no")!=null)){
      	sqlCmd_normal.append(" and WLX01.m2_name = ?"); 
      	paramList_normal.add(now_bank_no);
      	sqlCmd_revoke.append(" and WLX01.m2_name = ?"); 
      	paramList_revoke.add(now_bank_no);
   	  }
   	  else{//登入者是一級主管 則列出所有的地方主管機關   
   	  	sqlCmd_normal.append(" and bn01.bank_type=?");
   	  	paramList_normal.add(bank_type);
      	sqlCmd_revoke.append(" and bn01.bank_type=?");
      	paramList_revoke.add(bank_type);
   	  }
   }
   //login是A11111111的話 直接去抓在這分類(農會或漁會)所有的機構
   else{
   	    sqlCmd_normal.append(" and bn01.bank_type=?");
   	    paramList_normal.add(bank_type);
   	    sqlCmd_revoke.append(" and bn01.bank_type=?");
   	    paramList_revoke.add(bank_type); 
   }

  
   sqlCmd_normal.append(" order by bank_no");
   sqlCmd_revoke.append(" order by bank_no");
   //=================================================================================
   List dbData_online = DBManager.QueryDB_SQLParam(sqlCmd_normal.toString(),paramList_normal,"");		  
   //sqlCmd = "select bank_no,bank_name,update_date from bn01 where bank_type='"+bank_type+"' and bn_type = '2'";
   
   List dbData_revoke = DBManager.QueryDB_SQLParam(sqlCmd_revoke.toString(),paramList_revoke,"update_date,cancel_date");		  
   sqlCmd = new StringBuffer();
   sqlCmd.append("select * from WTT03_1 where program_id=?");
   paramList.add(link_name);
   List dbData_pgname = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
   
   if(dbData_pgname != null && dbData_pgname.size() != 0){
      bean = (DataObject)dbData_pgname.get(0);
      program_name = (String)bean.getValue("program_name");
      program_url = (String)bean.getValue("url_id");
      
      if(program_url.startsWith("FX004W.jsp")){
         program_url = "FX004W.jsp?act=Edit";
      }
      else if( program_url.startsWith("FX008AW.jsp")&&
      		   (request.getParameter("tbank_no")==null )&&
      		   (!muser_bank_type.equals("B")) ){
      	//如果不是地方主管機關且又可以進來fx008aw則為特權帳號，則列出第一層banklist
      	 Banklist_flag = true ;//表示這是第一層banklist
      	 program_url = "Bank_List.jsp?link_name=FX008AW";      	 
      }else if( program_url.startsWith("FX009W.jsp") && request.getParameter("tbank_no") == null && bank_type.equals("B") && !muser_bank_type.equals("B") ){
        //95.08.18 add 警示帳戶.地方主管機關進來.可看到其所屬的農漁會.農漁會進來只能看到自己的 by 2295
      	//如果不是地方主管機關且又可以進來fx009w則為特權帳號，則列出第一層banklist
      	 Banklist_flag = true; //表示這是第一層banklist
      	 program_url = "Bank_List.jsp?link_name=FX009W";
      }
      
   }
   
   System.out.println("dbData_online.size()="+dbData_online.size());
   System.out.println("dbData_revoke.size()="+dbData_revoke.size());
   session.removeAttribute("nowtbank_no");//清空已點選的tbank_no
   
   if(list_type.equals("1")){
      program_name = "地方主管機關"+program_name;
   }else if(list_type.equals("2")){
     program_name = "共用中心"+program_name;
   }else if(list_type.equals("3")){
     program_name = "農業行庫"+program_name;
   }  
   
   if(session.getAttribute("muser_id") == null){	
      System.out.println("Bank_List login timeout");   
	  rd = application.getRequestDispatcher( "/pages/reLogin.jsp?url=LoginError.jsp?timeout=true" );         	   
      rd.forward(request,response);
   }  
  
%>


<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/FX009W.js"></script>
<script language="javascript" event="onresize" for="window"></script>

<html>
<head>
<title>
<%=program_name%>
</title>
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

<!--body marginwidth="0" marginheight="0" leftmargin="0" topmargin="0" background="images/bg_1.gif" leftmargin="0"-->
<body marginwidth="0" marginheight="0" leftmargin="0" topmargin="0" leftmargin="0">
<form name=frmWMFileEdit method=post>
<input type="hidden" name="act" value="Status">  	
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
                      <td width="150"><img src="images/banner_bg1.gif" width="150" height="17"></td>
                      <td width="300"><font color='#000000' size=4><b> 
                        <center>
                          <%=program_name%>
                        </center>
                        </b></font> </td>
                      <td width="150"><img src="images/banner_bg1.gif" width="150" height="17"></td>
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
                <%//95.08.24 add 承受擔保品.警示帳戶.地方主管機關進來.可看到其所屬的農漁會.並出現列印按鈕 by 2495%>
                <tr class="sbody">                        	
						  	<td width='100%' colspan=3  bgcolor='#9AD3D0'> 
						  		<%
						  		
						  		if(!Banklist_flag&&!bank_type.equals("6")&&!bank_type.equals("7")&&(muser_bank_type.equals("B")||muser_bank_type.equals("2")||muser_id.equals("A111111111"))&&(link_name.equals("FX008AW")||link_name.equals("FX009W"))){%>
						  		<input type="text" name="S_YEAR" size="3" value="<% 
                      Date today = new Date();
                     out.print(today.getYear()-11);
                  %>">年 第<%
                    String select1 ="";
                    String select2 ="";
                    String select3 ="";
                    String select4 ="";

                   	if(today.getMonth() < 3)
                   		select4 ="selected";
                   	else if(today.getMonth() < 6)
                   		select1 ="selected";
                   	else if(today.getMonth() < 9)
                   		select2 ="selected";
                   	else if(today.getMonth() < 12)
                   		select3 ="selected";	

                %>
               		<select type="text" name='S_QUARTER' size="1"> 
                		<option value="1" <%= select1 %> >01</option>
                		<option value="2" <%= select2 %> >02</option>
                		<option value="3" <%= select3 %> >03</option>
                		<option value="4" <%= select4 %> >04</option>
                	 </select>
                    季 
                    <%if(link_name.equals("FX008AW")){
                    	String tbank_no = ( request.getParameter("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");		
                    %>
                     <a href="javascript:printSubmitB(this.document.forms[0],'print','<%=session.getAttribute("bank_type")%>','<%=tbank_no%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_print.gif',1);">
                    <img src="images/bt_print.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>  

                    <%}
                    if(link_name.equals("FX009W")&&(bank_type.equals("B")||bank_type.equals("2"))){
                    String tbank_no = ( request.getParameter("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");		
                    %>
                    <a href="javascript:printSubmit(this.document.forms[0],'print','<%=session.getAttribute("bank_type")%>','<%=tbank_no%>','<%=muser_id%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_print.gif',1);">
                    <img src="images/bt_print.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>  
                    <%}%>
                    
                                     
                    
	  					  <%}%>

	  					  </tr>
                          <tr class="sbody">
						  	<td width='100%' colspan=3 align='center' bgcolor='#9AD3D0'>營運中之機構</td>
	  					  </tr>
                          
                          <tr class="sbody" bgcolor='#D8EFEE'>
						  	<td width='30%' align='left'>機構代碼</td>
						  	<td width='70%' colspan=2 align='left'>機構名稱</td>
	  					  </tr>
	  					  <%
	  					  	int i = 0;
	  					    String bgcolor="#D3EBE0";
	  					    if(dbData_online != null && dbData_online.size() != 0){	  					       	
	  					       while(i<dbData_online.size()){
	  					      	 bgcolor = (i % 2 == 0)?"#e7e7e7":"#D3EBE0";	
	  					  %>
						  <tr class="sbody" bgcolor='<%=bgcolor%>'>
						  	<td width='30%'  align='left'>
						  	<% if(!list_type.equals("")){%>
						      <%if(program_url.startsWith("FX004W")){%>
						       <a href='<%=program_url%>&bank_no=<%=(String)((DataObject)dbData_online.get(i)).getValue("bank_no")%>'>
						      <%}else{%>
						       <a href='<%=program_url%>&tbank_no=<%=(String)((DataObject)dbData_online.get(i)).getValue("bank_no")%>&bank_type=<%=bank_type%>&list_type=<%=list_type%>'>                          
						      <%}%>
                          <%}else{%>
                              <%if(program_url.startsWith("FX004W")){%>
						        <a href='<%=program_url%>&bank_no=<%=(String)((DataObject)dbData_online.get(i)).getValue("bank_no")%>'>
						      <%}else{%>
                                <a href='<%=program_url%>&tbank_no=<%=(String)((DataObject)dbData_online.get(i)).getValue("bank_no")%>&bank_type=<%=bank_type%>'>
                              <%}%>  
                          <%}%>
						  <%=(String)((DataObject)dbData_online.get(i)).getValue("bank_no")%>
						  </a>
						  </td>
						  <td width='70%' colspan=2 >
						  <%if(!list_type.equals("")){%>
						        <%if(program_url.startsWith("FX004W")){%>
						           <a href='<%=program_url%>&bank_no=<%=(String)((DataObject)dbData_online.get(i)).getValue("bank_no")%>'>
						        <%}else{%>
						           <a href='<%=program_url%>&tbank_no=<%=(String)((DataObject)dbData_online.get(i)).getValue("bank_no")%>&bank_type=<%=bank_type%>&list_type=<%=list_type%>'>
						        <%}%>   
						  <%}else{%>
						        <%if(program_url.startsWith("FX004W")){%>
						           <a href='<%=program_url%>&bank_no=<%=(String)((DataObject)dbData_online.get(i)).getValue("bank_no")%>'>
						        <%}else{%>
						           <a href='<%=program_url%>&tbank_no=<%=(String)((DataObject)dbData_online.get(i)).getValue("bank_no")%>&bank_type=<%=bank_type%>'>
						        <%}%>   
						  <%}%>
						  <%=(String)((DataObject)dbData_online.get(i)).getValue("bank_name")%>
						  </a>
						  </td>
        			      </tr>
        			      <%     i++;
        			           }
        			        }else{        			        
        			      %>   
        			      <tr><td width='100%'  class="sbody" colspan=3 align='center' bgcolor='e7e7e7'>無機構相關資料</td></tr>
        			      <%}%>
        			      <%//if(!(bank_type.equals("6")/*農會*/ || bank_type.equals("7")/*漁會*/)){%>
        			      <tr class="sbody">
						  <td width='100%' colspan=3 align='center' bgcolor='#9AD3D0'>已裁撤之機構</td>
	  					  </tr>
        			      <tr class="sbody" bgcolor='#D8EFEE'>
						  <td width='30%' align='left' >總機構代碼</td>
						  <td width='40%' align='left' >總機構名稱</td>
						  <td width='30%' align='left' >裁撤日期</td>
	  					  </tr>
	  					  <%i = 0;
	  					    if(dbData_revoke != null && dbData_revoke.size() != 0){
	  					       bgcolor="#D3EBE0";
	  					       while(i<dbData_revoke.size()){
	  					       bgcolor = (i % 2 == 0)?"#e7e7e7":"#D3EBE0";		
	  					  %>
						  <tr class="sbody" bgcolor='<%=bgcolor%>'>
						  <td width='30%' align='left'>
						  <%if(program_url.startsWith("FX004W")){%>
						  <a href='<%=program_url%>&bank_no=<%=(String)((DataObject)dbData_revoke.get(i)).getValue("bank_no")%>'>
						  <%}else{%>
						  <a href='<%=program_url%>&tbank_no=<%=(String)((DataObject)dbData_revoke.get(i)).getValue("bank_no")%>'>
						  <%}%>
						  <%=(String)((DataObject)dbData_revoke.get(i)).getValue("bank_no")%>
						  </a>
						  </td>
						  <td width='40%' >
						  <%if(program_url.startsWith("FX004W")){%>
						  <a href='<%=program_url%>&bank_no=<%=(String)((DataObject)dbData_revoke.get(i)).getValue("bank_no")%>'>
						  <%}else{%>
						  <a href='<%=program_url%>&tbank_no=<%=(String)((DataObject)dbData_revoke.get(i)).getValue("bank_no")%>'>
						  <%}%>
						  <%=(String)((DataObject)dbData_revoke.get(i)).getValue("bank_name")%>
						  </a>
						  </td>
						  <td width='40%' >
						  <%if(((DataObject)dbData_revoke.get(i)).getValue("cancel_date") != null && !(((DataObject)dbData_revoke.get(i)).getValue("cancel_date")).equals("")){%>
						  <a href='<%=program_url%>&tbank_no=<%=(String)((DataObject)dbData_revoke.get(i)).getValue("bank_no")%>'>
						  <%if(((DataObject)dbData_revoke.get(i)).getValue("cancel_date") != null && !(((DataObject)dbData_revoke.get(i)).getValue("cancel_date")).equals("")){
						        out.print(Utility.getCHTdate((((DataObject)dbData_revoke.get(i)).getValue("cancel_date")).toString().substring(0, 10), 0));						  
						    }    
						  %>
						  </a>
						  <%}else{
						        out.print("&nbsp;");
						    }    
						  %>  
						  </td>
        			      </tr>
        			      <%     i++;
        			           }
        			        }else{
        			      %>   
        			      <tr><td width='100%'  class="sbody" colspan=3 align='center' bgcolor='e7e7e7'>無已裁撤之機構相關資料</td></tr>        			      
        			      <%}
        			      //}//end of bank_type != 6/7
        			      %>
        			      
                        </Table></td>
                    </tr>
      </table></td>
  </tr>
</table>
</form>
</body>
</html>
