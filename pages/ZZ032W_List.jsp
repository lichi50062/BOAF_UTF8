<%
// 95.01.03 create by 2295
//102.11.14 fix 日期顯示/QueryDB改preparestatement by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
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
    String tmpbank_type= ( request.getParameter("tmpbank_type")==null ) ? "" : (String)request.getParameter("tmpbank_type");				
	String bank_type = (session.getAttribute("bank_type") == null)?"":(String)session.getAttribute("bank_type");				
	String tbank_no = (session.getAttribute("tbank_no") == null)?"":(String)session.getAttribute("tbank_no");				
	String sztrans_type = ( request.getParameter("TRANS_TYPE")==null ) ? "" : (String)request.getParameter("TRANS_TYPE");				
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");				
	String szreport_no = ( request.getParameter("report_no")==null ) ? "" : (String)request.getParameter("report_no");					
	String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? YEAR : (String)request.getParameter("S_YEAR");				
	String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? MONTH : (String)request.getParameter("S_MONTH");				
	String szupd_code = ( request.getParameter("upd_code")==null ) ? "" : (String)request.getParameter("upd_code");				
	String szlock_status = ( request.getParameter("lock_status")==null ) ? "" : (String)request.getParameter("lock_status");				
    String szbank_code = ( request.getParameter("bank_code")==null ) ? "" : (String)request.getParameter("bank_code");				
    String lguser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		
    String firstStatus = ( request.getParameter("firstStatus")==null ) ? "" : (String)request.getParameter("firstStatus");			    	
	System.out.println("act="+act);
	
	System.out.println("ZZ032W_List.bank_type="+bank_type);
	System.out.println("ZZ032W_List.sztrans_type="+sztrans_type);
	System.out.println("ZZ032W_List.szreport_no="+szreport_no);
	System.out.println("ZZ032W_List.S_YEAR="+S_YEAR);
	System.out.println("ZZ032W_List.S_MONTH="+S_MONTH);
	System.out.println("ZZ032W_List.szupd_code="+szupd_code);
	System.out.println("ZZ032W_List.szlock_status="+szlock_status);
	System.out.println("ZZ032W_List.firstStatus="+firstStatus);
	System.out.println("ZZ032W_List.tmpbank_type="+tmpbank_type);	
	String sqlCmd = "";
		
	List lockList = (List)request.getAttribute("lockList");		
		
	if(lockList == null){
	   System.out.println("lockList == null");
	}else{
	   System.out.println("lockList.size()="+lockList.size());
	}
	
	//取得ZZ032W的權限
	Properties permission = ( session.getAttribute("ZZ032W")==null ) ? new Properties() : (Properties)session.getAttribute("ZZ032W"); 
	if(permission == null){
       System.out.println("ZZ032W_List.permission == null");
    }else{
       System.out.println("ZZ032W_List.permission.size ="+permission.size());               
    }	
    
   
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/ZZ032W.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<title>「每月(季)基本資料申報追蹤管理」</title>
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
<%if(act.equals("List")){%>
<input type="hidden" name="TRANS_TYPE" value="<%=sztrans_type%>">  
<input type="hidden" name="S_YEAR" value="<%=S_YEAR%>">  
<input type="hidden" name="S_MONTH" value="<%=S_MONTH%>">  
<input type="hidden" name="UPD_CODE" value="<%=szupd_code%>">  
<input type="hidden" name="LOCK_STATUS" value="<%=szlock_status%>">  
<%}%>
<%if(lockList != null && lockList.size() != 0){%>
<input type="hidden" name="row" value="<%=lockList.size()+1%>">   
<%}%>
<table width="665" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  		<tr> 
   		 <td><img src="images/space_1.gif" width="12" height="12"></td>
  		</tr>

        <tr> 
          <td bgcolor="#FFFFFF">
		  <table width="665" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="665" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="182"><img src="images/banner_bg1.gif" width="182" height="17"></td>
                      <td width="300"><font color='#000000' size=4><b> 
                        <center>
                          「每月(季)基本資料申報追蹤管理」 
                        </center>
                        </b></font> </td>
                      <td width="183"><img src="images/banner_bg1.gif" width="183" height="17"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><img src="images/space_1.gif" width="12" height="12"></td>
              </tr>
              <tr> 
                <td><table width="665" border="0" align="center" cellpadding="0" cellspacing="0">
               
                    <tr> 
                      <div align="right"><jsp:include page="getLoginUser.jsp?width=665" flush="true" /></div> 
                    </tr>                    
                    <tr> 
                       <td ><table width=665 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">  
                          <tr class="sbody">						  
						  <td width='15%' align='left' bgcolor='#D8EFEE'>申報項目</td>
                          <td width='85%' bgcolor='e7e7e7'>	
                            <select name='TRANS_TYPE' onchange="javascript:changeOption(document.forms[0]);" <%if(act.equals("List")) out.print("disabled");%>>                                                        
                            <!--select name='TRANS_TYPE' onchange="javascript:getData(this.document.forms[0],'<%=act%>');"-->                                                        
                            <%
                            List trans_typeList = null;
							trans_typeList = DBManager.QueryDB_SQLParam("select * from cdshareno where cmuse_div='029' and IDENTIFY_NO IN ('M', 'Q') order by cmuse_id",null,"");
	
                            if(trans_typeList != null){
                             for(int i=0;i<trans_typeList.size();i++){%>
                            <option value="<%=((String)((DataObject)trans_typeList.get(i)).getValue("cmuse_name")).substring(0,3)%>:<%=(String)((DataObject)trans_typeList.get(i)).getValue("identify_no")%>"                                                        
                            <%//if(sztrans_type.equals((String)((DataObject)trans_typeList.get(i)).getValue("cmuse_id"))) out.print("selected");%>
                            <%if(firstStatus.equals("true") && i==0) out.print("selected");%>
                            <%
                            if((((String)((DataObject)trans_typeList.get(i)).getValue("cmuse_name")).substring(0,3)+":"+((String)((DataObject)trans_typeList.get(i)).getValue("identify_no"))).equals(sztrans_type)){
                               out.print("selected");
                            }                            
                            %>
                            ><%=((String)((DataObject)trans_typeList.get(i)).getValue("cmuse_name")).substring(0,((String)((DataObject)trans_typeList.get(i)).getValue("cmuse_name")).length())%></option>                            
                            <%}
                            }//end of if
                            %>
                            </select>    
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                            <%if(act.equals("Qry")){%>
                            <input type="button" value="查詢清單" onclick="javascript:doSubmit(this.document.forms[0],'List');" name="QryList">                       
                            <%}%>
                            <%if(!act.equals("Qry")){%>
                            <input type="button" value="回查詢條件區" onclick="javascript:doSubmit(this.document.forms[0],'goQry');">  &nbsp;                                                        
                            <%}%>
                          </td>             
                          </tr>       
                          
                          
                          <tr class="sbody">                          
						  <td width='15%' align='left' bgcolor='#D8EFEE'>總機構代號</td>
                          <td width='85%' bgcolor='e7e7e7'>	
                            <input type="text" name="TBANK_NO" value="<% if(!szbank_code.equals("")){ out.print(szbank_code);}%>" size="7" maxlength="7" <%if(act.equals("List")) out.print("disabled");%>>                            
                            <input type="hidden" name="TBANK_NO" value="<% if(!szbank_code.equals("")){ out.print(szbank_code);}%>" >                            
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                            <%if(!act.equals("Qry")){%>
                            <a href="javascript:doSubmit(this.document.forms[0],'Lock');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image102','','images/bt_lockb.gif',1)"><img src="images/bt_lock.gif" name="Image102" width="66" height="25" border="0" id="Image102"></a>                       
 						  	<a href="javascript:doSubmit(this.document.forms[0],'unLock');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_nolockb.gif',1)"><img src="images/bt_nolock.gif" name="Image103" width="80" height="25" border="0" id="Image103"></a>                                                                             
 						  	<%}%>
                          </td>                                   
                          </tr>
                          
                          <tr class="sbody">                          
						  <td width='15%' align='left' bgcolor='#D8EFEE'>申報年月(季)</td>
                          <td width='85%' bgcolor='e7e7e7'>	
                            <input type='text' name='S_YEAR' value="<%=S_YEAR%>" size='3' maxlength='3' onblur='CheckYear(this)' <%if(act.equals("List")) out.print("disabled");%>>
        						<font color='#000000'>年
        						<select id="hide1" name=S_MONTH <%if(act.equals("List")) out.print("disabled");%>>        						
        						<%
        						    YMLoop:
        							for (int j = 1; j <= 12; j++) {
        							if( j > 4 && (!sztrans_type.equals("")) && (sztrans_type.substring(sztrans_type.indexOf(":")+1,sztrans_type.length())).equals("Q")){
        							    break YMLoop;
        							}
        							if (j < 10){%>        	
        							<option value=0<%=j%> <%if(S_MONTH.equals(String.valueOf(j))) out.print("selected");%>>0<%=j%></option>        		
            						<%}else{%>
            						<option value=<%=j%> <%if(S_MONTH.equals(String.valueOf(j))) out.print("selected");%>><%=j%></option>
            						<%}%>
        						<%  }%>
        						</select><font color='#000000'>月/季</font>
                          </td>                                   
                          </tr>
                          
                          <tr class="sbody">                          
                          <td width='15%' align='left' bgcolor='#D8EFEE'>檢核結果</td>
                          <td width='85%' bgcolor='e7e7e7'>                          
                            <select name='UPD_CODE' <%if(act.equals("List")) out.print("disabled");%>>                                                                                                                
                            <option value="ALL" <%if(szupd_code.equals("ALL")) out.print("selected");%>>全部</option>                            
                            <option value="0" <%if(szupd_code.equals("0")) out.print("selected");%>>已申報</option>                            
                            <option value="1" <%if(szupd_code.equals("1")) out.print("selected");%>>未申報</option>    
                            <%if((!sztrans_type.equals("")) && ((sztrans_type.substring(0,sztrans_type.indexOf(":"))).equals("C04") || (sztrans_type.substring(0,sztrans_type.indexOf(":"))).equals("C05")) ){%>                                                                       
                            <option value="2" <%if(szupd_code.equals("2")) out.print("selected");%>>未審核</option>                                                                                    
                            <%}%>
                            </select>                                 
                          </td>         
                          </tr>
                          
                          <tr class="sbody">                          
                          <td width='15%' align='left' bgcolor='#D8EFEE'>鎖定與否</td>
                          <td width='85%' bgcolor='e7e7e7'>                          
                            <select name='LOCK_STATUS' <%if(act.equals("List")) out.print("disabled");%>>                                                                                    
                            <option value="ALL" <%if(szlock_status.equals("ALL")) out.print("selected");%>>全部</option>                            
                            <option value="Y" <%if(szlock_status.equals("Y")) out.print("selected");%>>鎖定</option>                            
                            <option value="N" <%if(szlock_status.equals("N")) out.print("selected");%>>未鎖定</option>                            
                            </select>                                 
                            
                          </td>         
                          </tr>                                                    		      					      
                          </table>      
                      </td>    
                      </tr>
                      
                      
                      <tr> 
                      <td><table width=665 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99">  
                        <tr class="sbody">
                          <td width='665' colspan=19 bgcolor='D2F0FF'>	                         	 						  
 							<a href="javascript:selectAll(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image104','','images/bt_selectallb.gif',1)"><img src="images/bt_selectall.gif" name="Image104" width="80" height="25" border="0" id="Image104"></a>                       
 						    <a href="javascript:selectNo(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_selectnob.gif',1)"><img src="images/bt_selectno.gif" name="Image105" width="80" height="25" border="0" id="Image105"></a>                        							 						     						  
                          </td>
                        </tr>     		
                      
                     
                       	   <tr class="sbody" bgcolor="#9AD3D0">                       	  
                            <td width="20">序號</td>                                  
                            <td width="22">選項</td>                            
                            <td width="207">總機構代號及名稱</td>            				       				            				
        				    <td width="33">申報年月</td>            				            				            				
        				    <td width="76">檢核結果</td>
        				    <td width="31">申報筆數</td>            				            				            				
					        <td width="27">審核筆數</td>
            				<td width="71">申報(審核)單位鎖定</td>            				            				            				
            				<td width="35">局內鎖定</td>
					        <td width="110">鎖定最後異動日期</td>
					      </tr> 

					       <%
		                      int i = 0;      
        		              String bgcolor="#D3EBE0";                     		      
                   		      if((lockList == null) || (lockList.size() == 0)){%>
                   			   <tr class="sbody" bgcolor="<%=bgcolor%>">
                   			   <td colspan=19 align=center>無資料可供查詢</td><tr>
                   			   <tr>                   			   
                   			<%}else{                   			   
                   			   String update_date="";
                   			   String have_mk="";
                   			   String lock_mgr="";
                   			   String tot_cnt="";
                   			   String agree_cnt="";
                    		   while(i < lockList.size()){ 
                    		      update_date = "";
                    		      have_mk="";
                    		      lock_mgr="";
                    		      tot_cnt="";
                    		      agree_cnt="";
                    		      bgcolor = (i % 2 == 0)?"#e7e7e7":"#D3EBE0";	
                    		      if((String)((DataObject)lockList.get(i)).getValue("have_mk") != null){
                    		          have_mk = (String)((DataObject)lockList.get(i)).getValue("have_mk");
                    		          if(!have_mk.equals("*")){
                    		              have_mk = "未申報";
                    		          }                    		          
                    		      }
                    		      
                    		      if(((String)((DataObject)lockList.get(i)).getValue("s_updatedate") != null) && (!(((String)((DataObject)lockList.get(i)).getValue("s_updatedate")).trim()).equals(""))){ 	
                    		         update_date = ((String)((DataObject)lockList.get(i)).getValue("s_updatedate")).trim();                    		          
                    		         //102.11.14 fix 日期顯示 by 2295
                    		         if(update_date.startsWith("1")){
                                		update_date = update_date.substring(0,3)+"/"+update_date.substring(3,5)+"/"+update_date.substring(5,7)+ " " + update_date.substring(8,update_date.length());	                    		
                                	 }else{
                                		update_date = update_date.substring(0,2)+"/"+update_date.substring(2,4)+"/"+update_date.substring(4,6)+ " " + update_date.substring(7,update_date.length());	                    		
                                	 }        
                                	 
                    		         //System.out.println("update_date"+update_date);
                    		      }
                    		      
                    		      if( ((DataObject)lockList.get(i)).getValue("lock_mgr") != null  ){
                    		          lock_mgr = (String)((DataObject)lockList.get(i)).getValue("lock_mgr");
                    		      } 
                    		      
                    		      if( ((DataObject)lockList.get(i)).getValue("tot_cnt") != null ){
                    		        tot_cnt = (((DataObject)lockList.get(i)).getValue("tot_cnt")).toString();
                    		      }
                    		      if( ((DataObject)lockList.get(i)).getValue("agree_cnt") != null ){
                    		        agree_cnt = (((DataObject)lockList.get(i)).getValue("agree_cnt")).toString();
                    		      }
                    		      			            				
                    		      
                    		   
                      %>                         	  
                          <tr class="sbody" bgcolor="<%=bgcolor%>">
                            <td width="20"><%=i+1%></td>                       				            				
            				<td width="22">
            				<input type="checkbox" name="isModify_<%=(i+1)%>" value="<%if( ((DataObject)lockList.get(i)).getValue("bank_no") != null ) out.print((String)((DataObject)lockList.get(i)).getValue("bank_no"));%>"
            				
            				<%
            				  if(tmpbank_type.equals("Z")){
            				     //System.out.println("enabled1");            				     
            				  }else if((!lock_mgr.equals("Y")) && (tmpbank_type.equals("6") || tmpbank_type.equals("7")) && tot_cnt.equals(agree_cnt) &&  (Integer.parseInt(tot_cnt) > 0) ){
            				     //System.out.println("enabled2");
            				  }else{
            				     out.print(" disabled");
            				  }
            				%>
            				>
            				</td>            				
            				<td width="207">
            				<%if( ((DataObject)lockList.get(i)).getValue("s_report_name") != null ) out.print((String)((DataObject)lockList.get(i)).getValue("s_report_name")); else out.print("&nbsp;");%>            				
            				</td>
            				<td width="33">            				
            				<%=S_YEAR%><%if(S_MONTH.length() ==1) out.print("0");%><%=S_MONTH%>
            				</td>
            				<td width="76"><%=have_mk%></td>
            				<td width="31">
            				<%if( ((DataObject)lockList.get(i)).getValue("tot_cnt") != null && (!tot_cnt.equals("0")) ) out.print((((DataObject)lockList.get(i)).getValue("tot_cnt")).toString()); else out.print("&nbsp;");%>            				            				
            				</td>
            				<td width="27">
            				<%if( ((DataObject)lockList.get(i)).getValue("agree_cnt") != null && (!agree_cnt.equals("0")) )  out.print((((DataObject)lockList.get(i)).getValue("agree_cnt")).toString()); else out.print("&nbsp;");%>            				            				
            				</td>
            				
            				<td width="71">
            				<%if( ((DataObject)lockList.get(i)).getValue("lock_own") != null && ((String)((DataObject)lockList.get(i)).getValue("lock_own")).equals("Y")) out.print((String)((DataObject)lockList.get(i)).getValue("lock_own")); else out.print("&nbsp;");%>            				            				
            				</td>
            				<td width="35">
            				<%if( ((DataObject)lockList.get(i)).getValue("lock_mgr") != null && ((String)((DataObject)lockList.get(i)).getValue("lock_mgr")).equals("Y")) out.print((String)((DataObject)lockList.get(i)).getValue("lock_mgr")); else out.print("&nbsp;");%>            				            				
            				</td>            				
            				<td width="110">
            				<%if(!update_date.equals("")) out.print(update_date); else out.print("&nbsp;");%>            				
            				</td>
            			  </tr> 					      
					      <%
                  			   i++;
	                  		   }//end of while
	                  		}//end of if
    			          %>  
                  		  
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
