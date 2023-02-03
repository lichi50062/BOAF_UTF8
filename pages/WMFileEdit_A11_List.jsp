<%
/*
 *102.01.08 create  by 2968
 *104.05.12 add A111111111增加顯示案件分項.項數 by 2295
 */
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
<%@ page import="java.util.Date" %>
<%

	Properties permission = ( session.getAttribute("WMFileEdit")==null ) ? new Properties() : (Properties)session.getAttribute("WMFileEdit"); 
	if(permission == null){
	   System.out.println("WMFileEdit_A11.permission == null");
	}else{
	   System.out.println("WMFileEdit_A11.permission.size ="+permission.size());
	           
	}
	//---------------------------------------取得參數---------------------------------------------
	List A11_list1 = (List)request.getAttribute("A11_S");   	//內含每月的資料
	List A11_list2 = (List)request.getAttribute("A11_S_Sum");   //內含年月以及筆數
	String bank_no   = ( request.getParameter("bank_no")==null ) ? "" :request.getParameter("bank_no");
	String muser_tbank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");
	String bank_type   = ( request.getParameter("bank_type")==null ) ? "" :request.getParameter("bank_type");
	String tbank_no   = ( request.getParameter("tbank_no")==null ) ? "" :request.getParameter("tbank_no");
	DataObject Bean = new DataObject(); 
	/***
    * 99.12.08 fix 縣市合併年度區分
    **/
	List paramList =new ArrayList() ;
	String t_year = "99" ;
    if(Integer.parseInt(Utility.getYear()) > 99) {
    	t_year = "100" ;
    }
    String bank_name = "";
	String sqlCmd = "select bank_name from BA01 where bank_no=? and m_year=?";
	paramList.add(bank_no) ;
	paramList.add(t_year) ;
	List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
	if(dbData != null && dbData.size() != 0){
	   bank_name = (String)((DataObject)dbData.get(0)).getValue("bank_name");
	   System.out.println("bank_no="+bank_no);
	   System.out.println("bank_name="+bank_name);
	} 
	
	if(A11_list1 == null){
		   System.out.println("A11_list1 == null");
		}else{
		    if(A11_list1.size() != 0){
			   //使用dataobject type 來做資料傳遞
			   Bean = (DataObject)A11_list1.get(0);//把選取的資料取出來
		    }
		   System.out.println("A11_list1.size()="+A11_list1.size());
		}
	if(A11_list2 == null){
		   System.out.println("A11_list2 == null");
		}else{
		    if(A11_list2.size() != 0){
			   //使用dataobject type 來做資料傳遞
			   Bean = (DataObject)A11_list2.get(0);//把選取的資料取出來
		    }
		   System.out.println("A11_list2.size()="+A11_list2.size());
		}
	
	
	//table width
	int width = 1024;//1300 pixel

%>
<!-- ------------------------------------------------------------------------------------------ -->
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
<script language="javascript" src="js/Common.js"> </script>
<script language="javascript" src="js/FileEdit_A11.js"> </script>
<script language="javascript" event="onresize" for="window"> </script>
<!-- ------------------------------------------------------------------------------------------- -->
<html>
<link href="css/b51.css" rel="stylesheet" type="text/css">
<head>
<title>聯合貸款案件資料表申報情形</title>
</head>
<body marginwidth="0" marginheight="0" leftmargin="0" topmargin="0" leftmargin="0">
<form method=post>
  <table border="0" cellspacing="0" width="<%=width %>" cellpadding="0" bgcolor="#FFFFFF">
    <tr>
      <td height="18" width="<%=width %>"></td>
    </tr>
    <tr>
      <td width="<%=width %>">
     	 <center>
      		<table border="0" cellspacing="1">
          		<tr>
          			<td><img src="images/banner_bg1.gif" width="<%=(width/2-130) %>" height="17"></td>
          			<td><b><font color='#000000' size="4">聯合貸款案件資料表申報情形</font></b></td>
          			<td><img src="images/banner_bg1.gif" width="<%=(width/2-130) %>" height="17"></td>
          		</tr>
      		</table>
     	</center>
     </td>
    </tr>

    <tr>
     <div align="right"><jsp:include page="getLoginUser.jsp?width=1024" flush="true" /></div>
    </tr>

    <tr>
      <td width="<%=width %>">
         <center>
          <table border="1" cellspacing="1" bordercolor="#3A9D99" width="<%=width %>" height="25">
            <tr>
            	<td class="sbody" bgcolor="#D8EFEE" width="110"><p align="center">申報年月</p></td>
             	<td class="sbody" bgcolor="#E7E7E7" width="150" valign="middle" align="center" >
             		<input type="text" name='s_year' value="<% 
                     out.print(Utility.getYear());
                     %>" size="3" maxlength="4" onblur="CheckYear(this)">年 
               		<select name="s_month" size="1">
                    <%
                   
                    for(int i=1;i<=12;i++)
                    {
                        if(i == Integer.parseInt(Utility.getMonth()))
                           out.print("<option value="+i+" selected>"+i+"</option>");
                        else
                           out.print("<option value="+i+">"+i+"</option>");
                    }
                    %>

                    </select> 月
                </td>
                
             	<td class="sbody" bgcolor="#E7E7E7" valign="middle">
             		<div align="left">
             		<% // 若為A111111111或農金局帳號，都可以選金額單位
             		if("8888888".equals(muser_tbank_no) || "BOAF000".equals(muser_tbank_no)){ %>
             		<!--金額單位-->
             			<span class="mtext">金額單位 :</span>
                        <select size="1" name="Unit">
                            <option value ='1' >元</option>
                            <option value ='1000' selected>千元</option>
                            <option value ='10000' >萬元</option>
                            <option value ='1000000' >百萬元</option>
                            <option value ='10000000' >千萬元</option>
                            <option value ='100000000' >億元</option>
                        </select>
             		<%}else{
             		//if(!"BOAF000".equals(muser_tbank_no)){ 
             			//if(!"8888888".equals(muser_tbank_no)){%>
             			<%//} %>
             			<input type='hidden' name='Unit' value="1000" >
             		<%} %>
             		<%if(permission != null && permission.get("A") != null && permission.get("A").equals("Y")){ //add%> 
             			<!-- Insert a data-->
             			<a href="javascript:doSubmit(this.document.forms[0],'New','<%=bank_no%>','','','','','');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_addb.gif',1)"><img src="images/bt_add.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>
             		<%} %>
             		<!-- Print a data-->
             		<%if(permission != null && permission.get("P") != null && permission.get("P").equals("Y")){//Print %>
             			<a href="javascript:printSubmit(this.document.forms[0],'<%=bank_no%>','<%=bank_name%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_print.gif',1);"><img src="images/bt_print.gif" name="Image101" width="60" height="25" border="0" id="Image101"></a>
             		<%} %>
             		</div>
              	</td>
            </tr>
         </table>
        </center>
     </td>
    </tr>

	<center>
    <tr>
      <td width="<%=width %>">
        <table border="1" cellspacing="1" width="<%=width %>" bordercolor="#3A9D99">
          <tr class="sbody">
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width='10%'>申報年月</td>
            <td bgcolor="#9AD3D0" colspan="2" valign="middle" align="center" width='6%'>總筆數</td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width='7%'>明細</td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width='9%'>申報編號</td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width='15%'>授信案總金額</td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width='15%'>參貸額度</td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width='15%'>實際授信餘額</td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width='15%'>異動者 帳號/姓名</td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width='8%'>異動 日期</td>
          </tr>
         </table>
      </td>
    </tr>
 <!-- -----------start to parse the form ----------------------------------------------------------- -->
 <%
 		if(A11_list1.size()==0)
 		{
 %>			
 		<center>
 			<tr>
 				<td>   
 					<table border="1" cellspacing="1" width="<%=width %>" bordercolor="#3A9D99">
 						<tr class="sbody" bgcolor="#D8EFEE">
 							<td colspan="15" align=center width="<%=width %>">尚無資料</td>
         				</tr>
         			</table>
	     		</td>
			</tr>
      	</center>
 				 				   
 <%
		}else{
				
				//備註 使用div圖層與javascript去達成展開的動作
				
				int count=0;				//紀錄在哪一年哪一月之下有多少筆資料
				int displacement=0;	//紀錄位移(因為所有的資料是放在同一個list中,所以要紀錄之前已經印過多少筆資料)
				String m_yearmonth  = ""; 
				String m_year  = ""; 
				String m_month = ""; 
				String case_no = "";
				String apply_cnt    = ""; //總筆數
				String loan_amt_sum = ""; //授信案總金額
				String loan_amt 	= ""; //參貸額度
				String loan_bal_amt = ""; //實際授信餘額
				String seq_no       = "";
				String user_id      = "";
				String user_name    = "";
				String update_date  = "";
				String bgcolor      = "";
				String case_cnt     = ""; //總項數
				boolean hyper_yn = true; //判斷要不要加超連結 (如曾經審核不管過不過就不加超連結)
				
				for(int i=0;i<A11_list1.size();i++){
						count = Integer.parseInt(((DataObject)A11_list1.get(i)).getValue("apply_cnt").toString()); 
 %> 
 			<center>
 			<tr>
 			<td>      
 			<table border="1" cellspacing="1" width="<%=width %>" bordercolor="#3A9D99" bgcolor="#D8EFEE">       
          	<tr class="sbody">
            <td valign="middle" align="center" width="10%" > 
            	  <!--	<DIV onmousedown= "fmenu0() style="cursor:hand;"> -->
            	    <% 
            			if(count > 0 ){
										out.print("<u>");
            		    out.print("<DIV onmousedown=fmenu"+i+"() style=\"cursor:hand;\">");
            		  }
            	    if(count==0){
                        out.print("<a href='WMFileEdit_A11.jsp?act=Edit&seqNo=&m_year="+((DataObject)A11_list1.get(i)).getValue("m_year")+"&m_month="+((DataObject)A11_list1.get(i)).getValue("m_month")+"&bank_no="+bank_no+"'>"+((DataObject)A11_list1.get(i)).getValue("m_yearmonth")); 
                   	}else{
            			out.print( ((DataObject)A11_list1.get(i)).getValue("m_yearmonth"));
                   	}
            	  	if(count > 0){
						out.print("</u>");
            		    out.print("</div>");
            		  }
            	  	
            		%>

            </td>
            <td valign="middle" align="center" width='6%'><%= count %></td>
            <td valign="middle" align="center" width='7%'>&nbsp;</td>
            <td valign="middle" align="center" width='9%'>&nbsp;</td>
            <td valign="middle" align="center" width='15%'>&nbsp;</td>
            <td valign="middle" align="center" width='15%'>&nbsp;</td>
            <td valign="middle" align="center" width='15%'>&nbsp;</td>
            <td valign="middle" align="center" width='15%'>&nbsp;</td>
            <td valign="middle" align="center" width='8%'>&nbsp;</td>
         	 </tr>
          </table>
         </td>
        </tr>
      </center>
         
      <center> 
         <tr>
          <td>
            <table id=menu<%=i %> style="display:none" border="1" cellspacing="1" width="<%=width %>" bordercolor="#3A9D99" >
<%         					
		  //在這筆年月之下有多少筆資料 列出來
          for(int j=0;j < count;j++){
          	
          	DataObject bean = (DataObject)A11_list2.get(j+displacement);
			m_yearmonth = bean.getValue("m_yearmonth").toString();
			m_year 		= bean.getValue("m_year").toString();
			m_month 	= bean.getValue("m_month").toString();
			seq_no		= bean.getValue("seq_no").toString();
			case_no		= (String.valueOf( bean.getValue("case_no"))=="null")?"":String.valueOf( bean.getValue("case_no"));
			case_cnt		= (String.valueOf( bean.getValue("case_cnt"))=="null")?"":String.valueOf( bean.getValue("case_cnt"));
			loan_amt_sum = (String.valueOf( bean.getValue("loan_amt_sum"))=="null")?"0":String.valueOf( bean.getValue("loan_amt_sum")); 			
			loan_amt = (String.valueOf( bean.getValue("loan_amt"))=="null")?"0":String.valueOf( bean.getValue("loan_amt")); 
			loan_bal_amt = (String.valueOf( bean.getValue("loan_bal_amt"))=="null")?"0":String.valueOf( bean.getValue("loan_bal_amt")); 
			user_id		 = bean.getValue("user_id").toString();
			user_name	 = bean.getValue("user_name").toString();
			update_date	 = bean.getValue("update_date").toString();
					
			bgcolor = (j % 2 == 0)?"#E7E7E7":"#D3EBE0";
%>
          		<tr class="sbody" bgcolor="<%= bgcolor %>">
            		<td valign="middle" align="center" width='10%'>&nbsp;</td>
            		<td valign="middle" align="center" width='6%'>
            		     <%if(((String)session.getAttribute("muser_id")).equals("A111111111")){%>            
            		     	分項:<%=case_cnt %>
	            		 <%}else{%>
            		      	&nbsp;
            		     <%}%>	
            		</td>
            		<td valign="middle" align="center" width='7%'>
            				<!-- -------做超連結------ -->
            				<% if(hyper_yn){%>
            					<a href="WMFileEdit_A11.jsp?act=Edit&seq_no=<%=seq_no%>&m_year=<%=m_year%>&m_month=<%=m_month%>&bank_no=<%=bank_no %>&case_no=<%=case_no %>"> 
										<%}%>
								<%if("999".equals(seq_no)){
									seq_no="0";
								}%>
            				<%=  seq_no %>
							<% if(hyper_yn){%>
            					</a>	
										<%}%>
            		</td>
            		<%if(!"0".equals(seq_no)){ %>
	            		<td valign="middle" align="center" width='9%'><%=case_no %></td>
	            		<td valign="middle" align="center" width='15%'> <%= Utility.setCommaFormat(loan_amt_sum)%></td>
	            		<td valign="middle" align="center" width='15%'> <%= Utility.setCommaFormat(loan_amt) %></td>
	            		<td valign="middle" align="center" width='15%'>	<%= Utility.setCommaFormat(loan_bal_amt)%> </td>
            		<%}else{ %>
            			<td valign="middle" align="center" colspan="4">本月無新增聯貸案件申報資料</td>
            		<%} %>
            		<td valign="middle" align="center" width='15%'> <%= user_id %>  /<br><%= user_name %></td>
            		<td valign="middle" align="center" width='8%'> <%= update_date %> </td>            		
            		</td>
            		
            		
          		</tr>
          		</form>
 <%
        }//end of for j 
        displacement+=count;
 %>
 					</table>
 				</td>
 			</tr>
 		</center>
 <%
			 }//end of for i
		}//end of else
 %>
    </center>
<!-- ------------------------------------------------------------------------------------------------------------------------------------------- -->
    <tr>
      <td width="1000">
          <center>
          <jsp:include page="getMaintainUser.jsp?width=1024" flush="true" /><!--載入維護者資訊 -->
          </center>
      </td>
    </tr>
    <tr class="sbody">
      <td height="145" width="<%=width %>">
          <table border="0" cellspacing="1" width="<%=width %>" height="150">
            <tr>
              <td colspan="2" width="500" height="32"><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle"><font color="#007D7D" size="3">使用說明:</font>
              </td>
            </tr>
            <tr class="sbody">
              <td width="10" height="111"></td>
              <td width="<%=width %>" height="111">
              	<ul>
                    <li><font color="#0000FF">[新增]:</font>輸入申報之年月再點選【新增】按鈕可新增該年月「聯合貸款案件」之資料。</li>
                    <li><b>每月於申報<font color="#FF0000">新增</font>聯貸案件時，所<font color="#FF0000">參貸</font>之<font color="#FF0000">所有授信項目</font>請<font color="#FF0000">一併</font>填列。
                    	<br>例：一聯貸案件共有甲、乙、丙3項授信項目，倘參貸甲、乙2項，於申報時，該甲、乙項請一併申報。</b></li>
                    <li><b>每月均已將上月資料全數載入申報當月，以方便維護資料；另申報當月若無新增聯貸案件，請直接點選「申報年月」當月資料，並確認修正<font color="#FF0000">實際授信餘額</font>等相關資料。</b></li>
                    <li><font color="#0000FF">[修正]</font>:點選該[申報年月]之[總筆數]即會自動展開(或收合)明細記錄,再點選[明細]欄即可變更該申報之資料。</font></li>
                    <li>本表係按最近的[申報年月]先排序,依此類推..........</li>
                    <li><font color="#FF0000">若當月份無資料時，也需申報</font></li>
                </ul>
              </td>
            </tr>
         </table>
      </td>
    </tr>
    <tr>
      <td height="18" width="1024"></td>
    </tr>

  </table>
  </form>

</body>

</html>

<!-- --------------------------------------------------- -->
<script language='JavaScript'>
<%
	DataObject tempBean;
	
	for(int k=0;k < A11_list1.size(); k++){
		tempBean = ((DataObject)A11_list1.get(k));
%>
		function fmenu<%= k %>(){
			if( menu<%= k %>.style.display == "none")
				menu<%= k %>.style.display = "block";
			else
				menu<%= k %>.style.display = "none";
		}
		arrpush('<%=tempBean.getValue("m_year").toString()+tempBean.getValue("m_month").toString()%>');
<%
	}
%>
</script>
<!-- -------------------------------------------------	-->
