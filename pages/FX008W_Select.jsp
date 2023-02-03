<%
//94.10.27 first designed by 4183
%>

<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<link href="css/b51.css" rel="stylesheet" type="text/css">
<script language="javascript" src="js/Common.js"> </script>
<script language="javascript" src="js/FX008W.js"> </script>
<script language="javascript" event="onresize" for="window"> </script>

<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Date" %>
<%
  
  //內含本季的資料	
	List FX008_Load = (List)session.getAttribute("FX008_Load");
	List FX008_Lock = (List)session.getAttribute("WLX08_Lock");    //內含被LOCK的年季   		
	String last_year = ( (String)session.getAttribute("last_year")==null ) ? "" :(String)session.getAttribute("last_year");
	String last_quarter = ( (String)session.getAttribute("last_quarter")==null ) ? "" :(String)session.getAttribute("last_quarter");
	String load_year    =  ( session.getAttribute("load_year")==null ) ? " " : (String)session.getAttribute("load_year");
	String load_quarter =  ( session.getAttribute("load_quarter")==null ) ? " " : (String)session.getAttribute("load_quarter");
	String bank_no = ( request.getParameter("bank_no")==null ) ? "" :request.getParameter("bank_no");
    			
	if( FX008_Load == null){
		System.out.println("The  FX008_Load list is null");
	}
	else{
		System.out.println("The  size of FX008_Load list "+ FX008_Load.size());
	}
	
	//判斷申報年季不可是巳鎖定，若已鎖定將返回到list畫面
 		if(FX008_Lock.contains(load_year+load_quarter))
 			out.print("<script>alert(\"本季資料已鎖定，將返回清單畫面\");history.go(-1);</script>");

%>

<html>
<head>
<title>上季資料列表</title>
</head>

<body>
<form method="POST">

<table class= "sbody" border="0" cellspacing="1" width="900">
  <tr>
    <td width="900" height="15"></td>
  </tr>
  <tr>
    <td width="900" height="30">
        <center>
          <table border="0" cellspacing="1">
						<tr>
              <td><img src="images/banner_bg1.gif" width="165" height="17"></td>
              <td><font size="4"><b><%=last_year %>年 第<%=last_quarter%>季資料列表</b></font></td>
              <td><img src="images/banner_bg1.gif" width="165" height="17"></td>
            </tr>
          </table>
        </center>
    </td>
  </tr>
  <tr>
    <td width="900">
      <table border="1"  bordercolor="#3A9D99">
        <tr>
					<jsp:include page="getLoginUser.jsp?width=900" flush="true" /> <!-- list the login user in form-->
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td width="900">
        <center>
        <table class= "sbody" border="1" cellspacing="1" bordercolor="#3A9D99" width="900">
          <tr>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="40" >選取</td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="40" >編號</td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="80" >機構名稱<br>(借款人)</td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="60" >承受<br>日期</td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="120">承受擔保品<br>座落</td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="80" >帳列金額<br><font color="#ff0000">(新台幣元)</font></td>            
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="60" >申請延長<br>期間</td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="120">申請延長<br>理由</td>
            <td bgcolor="#9AD3D0" colspan="4" valign="middle" align="center" width="120">縣/市政府評註</td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="80" >異動者<br>帳號/姓名</td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="70" >異動日期</td>
          </tr>
          <tr>
            <td bgcolor="#9AD3D0" valign="middle" align="middle" width="30">提足備抵趺價損失</td>
            <td bgcolor="#9AD3D0" valign="middle" align="middle" width="30">有積極處分之事實</td>
            <td bgcolor="#9AD3D0" valign="middle" align="middle" width="30">未來處分計劃合理可行</td>
            <td bgcolor="#9AD3D0" valign="middle" align="middle" width="30">審核結果</td>
          </tr>
          <%
  					if(FX008_Load.size()==0){
  				%>	
  				 <tr>
            <td bgcolor="#D8EFEE" valign="middle" align="middle" width="900" colspan="14"  >本季無資料</td>
           </tr>
  				<%
  					}
  					else{
  					String bgcolor ="";
  					Date duDate;
						int year = 0; 
						int month = 0;
						int day = 0;
  					
          		for(int i=0;i<FX008_Load.size();i++){
          			duDate = (Date)((DataObject)FX008_Load.get(i)).getValue("duredate"); 
								year   = duDate.getYear()-11;                                                      
								month  = duDate.getMonth()+1;                                                      
								day    = duDate.getDate();                                                         
	
          		
 							if(i%2 == 1)
 								bgcolor="#D3EBE0";
 							else
 								bgcolor="#E7E7E7";
          %>
          <tr bgcolor="<%= bgcolor %>">
            <td valign="middle" align="center" width="40"  ><input type="checkbox" name="<%="C"+i%>" value="true">           </td>
            <td valign="middle" align="center" width="40"  ><%=((DataObject)FX008_Load.get(i)).getValue("dureassure_no")%>   </td>
            <td valign="middle" align="center" width="80"  ><%=((DataObject)FX008_Load.get(i)).getValue("debtname")%>        </td>
            <td valign="middle" align="center" width="60"  ><%= year %><%="/"+month+"/"%><%= day %>                           </td>
            <td valign="middle" align="center" width="120" ><%=((DataObject)FX008_Load.get(i)).getValue("dureassuresite")%>  </td>
            <td valign="middle" align="center" width="80"  ><%=((DataObject)FX008_Load.get(i)).getValue("accountamt")%>      </td>
            <td valign="middle" align="center" width="60"  ><%=((DataObject)FX008_Load.get(i)).getValue("applydelayyear")+" 年"+((DataObject)FX008_Load.get(i)).getValue("applydelaymonth")+" 個月"%></td>
            <td valign="middle" align="center" width="120" ><%=((DataObject)FX008_Load.get(i)).getValue("applydelayreason")%></td>
            <td valign="middle" align="center" width="30"  >&nbsp;<%=((DataObject)FX008_Load.get(i)).getValue("damage_yn")%>       </td>
            <td valign="middle" align="center" width="30"  >&nbsp;<%=((DataObject)FX008_Load.get(i)).getValue("disposal_fact_yn")%></td>
            <td valign="middle" align="center" width="30"  >&nbsp;<%=((DataObject)FX008_Load.get(i)).getValue("disposal_plan_yn")%></td>
            <td valign="middle" align="center" width="30"  >&nbsp;<%=((DataObject)FX008_Load.get(i)).getValue("auditresult_yn")%>  </td>
            <td valign="middle" align="center" width="80"  ><%=((DataObject)FX008_Load.get(i)).getValue("user_id")+"/<br>"+((DataObject)FX008_Load.get(i)).getValue("user_name")%></td>
            <td valign="middle" align="center" width="70"  ><%=((DataObject)FX008_Load.get(i)).getValue("update_date")%>     </td>
          </tr>
        <%
       		}//end of for
				}//end of else
  			%>  
       </table>
     </center>
    </td>
  </tr>
  
  <tr>
  <center>
  <table>
  <!-- -------輸出 按鈕-------- -->
		<td width="70">
			<a href="javascript:doSubmit(this.document.forms[0],'Load_To','123');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_confirmb.gif',1)">
			<img src="images/bt_confirm.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>
		</td>
		<td width="70">
			<a href="javascript:doSubmit(this.document.forms[0],'123','123');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_cancelb.gif',1)">
			<img src="images/bt_cancel.gif" name="Image105" width="66" height="25" border="0" id="Image105"></a>
		</td>
		<td width="90">
			<a href="javascript:history.back();"  onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image106','','images/bt_backb.gif',1)">
			<img src="images/bt_back.gif" name="Image106" width="80" height="25" border="0" id="Image106"></a>
		</td>
	</table>
	</center>
	</tr>
  
  <tr>
    <td width="900" >
      
      <table class= "sbody">
        <tr>
          <td width="600" colSpan="2" height="40"><img src="/images/arrow_1.gif" align="absMiddle" width="28" height="23">
          	<font color="#007d7d" size="3">使用說明 :</font>
          </td>
        </tr>
        <tr>
          <td width="20" >&nbsp;</td>
          <td width="600">
            	<ul>
               	<li class="sbody">勾選想要載入之上季資料無誤後, 按
               		<font color="#666666">【確定】即將本表所勾選的資料, 於資料庫中建檔。</font>
            		<li class="sbody">欲重新輸入時, 按【取消】即將本表上的選擇全部取消
            		<li class="sbody">如放棄修改或無修改之資料需輸入, 按【回上一頁】]即離開本程式。
            	</ul>
          </td>
        </tr>
      </table>
    
    </td>
  </tr>
  
  </form>
</table>


</body>

</html>
