<%
// 94.10.14 first designed by 4183 lilic0c0
//100.01.27 fix title顯示寬度 by 2295
%>
<link href="css/b51.css" rel="stylesheet" type="text/css">
<script language="javascript" src="js/Common.js"> </script>
<script language="javascript" src="js/FX008AW.js"> </script>
<script language="javascript" event="onresize" for="window"> </script>

<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Date" %>
<%	//---------------------------------------取得參數---------------------------------------------
	List FX008_list1 = (List)request.getAttribute("WLX08_AS");		//內含每季的資料
	List FX008_list2 = (List)request.getAttribute("WLX08_AS_Sum");	//內含年季以及筆數
	List FX008_Lock  = (List)session.getAttribute("WLX08_ALock");	//內含被LOCK的年季
	String bank_no   = ( request.getParameter("bank_no")==null ) ? "" :request.getParameter("bank_no");
	
	//table width
	int width = 1340;//1300 pixel
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
<!-- ------------------------------------------------------------------------------------------- -->
<html>

<head>
<title>承受擔保品延長處分申報情形清單</title>
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
          			<td><img src="images/banner_bg1.gif" width="<%=(width/2-215) %>" height="17"></td>
          			<td><b><font color='#000000' size="4">縣市政府對「承受擔保品延長處分申報情形」審核</font></b></td>
          			<td><img src="images/banner_bg1.gif" width="<%=(width/2-215) %>" height="17"></td>
          		</tr>
      		</table>
     	</center>
     </td>
    </tr>
    
    <tr>
     <div align="right"><jsp:include page="getLoginUser.jsp?width=1340" flush="true" /></div>
    </tr>
    
    <tr>
      <td width="<%=width %>">
        <table border="1" cellspacing="1" width="<%=width %>" bordercolor="#3A9D99">
          <tr class="sbody">
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="60">申報<br><br>年季</td>
            <td bgcolor="#9AD3D0" colspan="2" valign="middle" align="center" width="80">承受擔保品</td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="80">機構名稱<p>(借款人)</p></td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="50">承受<p>日期</p></td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="120">承受擔保品<p>座落</p></td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="80">帳列金額<p><font color="#ff0000">(新台幣元)</font></td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="70">申請延長<p>期間</p></td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="120">申請延長<p>理由</p></td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="100">(核准申請)<p>延長期間/<br>延長期限</p></td>
            <td bgcolor="#9AD3D0" colspan="4" valign="middle" align="center" width="160">縣/市政府評註</td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="100">(縣市政府)<p>核准文號/<br>核准日期</p></td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="100">(函報農金局)<p>文號/日期</p></td>
            <td bgcolor="#9AD3D0" valign="middle" align="center" rowspan="2" width="80">異動者<p>帳號/姓名</p></td>
            <td bgcolor="#9AD3D0" valign="middle" align="center" rowspan="2" width="70">異動<p>日期</p></td>
          </tr>
          <tr class="sbody">
            <td bgcolor="#9AD3D0" valign="middle" align="center" width="40">總筆數</td>
            <td bgcolor="#9AD3D0" valign="middle" align="center" width="40">編號</td>
            <td bgcolor="#9AD3D0" valign="middle" align="center" width="40">提足備抵趺價損失</td>
            <td bgcolor="#9AD3D0" valign="middle" align="center" width="40">有積極處分之事實</td>
            <td bgcolor="#9AD3D0" valign="middle" align="center" width="40">未來處分計劃合理可行</td>
            <td bgcolor="#9AD3D0" valign="middle" align="center" width="40">審核結果</td>
          </tr>
         </table>
      </td>
    </tr>
 <!-- -----------start to parse the form ----------------------------------------------------------- -->
 <%
 		if(FX008_list2.size()==0)
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
		}
		else{
				
				//備註 使用div圖層與javascript去達成展開的動作
				
				int count=0;		//紀錄在哪一年哪一季之下有多少筆資料
				int displacement=0;	//紀錄位移(因為所有的資料是放在同一個list中,所以要紀錄之前已經印過多少筆資料)
				String seq_no =	"";
				String m_year =	"";
				String m_quarter = "";
				String bgcolor = "";
				String damage_yn = ""; 
				String disposal_fact_yn = "";
				String disposal_plan_yn = "";
				String auditresult_yn   = "";
				boolean hyper_yn = true; //判斷要不要加超連結 (如被lock就不加超連結)
				
				//---------------------
				Date duDate;
				int year = 0;
				int month = 0;
				int day = 0;
				
				//---------------------
				Date get_Cnt_date;
						
				int Cnt_year  = 0;
				int Cnt_month = 0;
				int Cnt_date  = 0;
				//---------------------
				Date get_ApplyOK_date;						
				int ApplyOK_y  = 0;
				int ApplyOK_m = 0;
				int ApplyOK_d  = 0;
				
				//--------------------
				Date get_BOAF_date;	
				int BOAF_y  = 0;
				int BOAF_m = 0;
				int BOAF_d  = 0;
	
				
				for(int i=0;i<FX008_list2.size();i++){
					count = Integer.parseInt(((DataObject)FX008_list2.get(i)).getValue("cnt").toString());                                                
 %> 
 	<center>
 	<tr>
 		<td>      
 		<table border="1" cellspacing="1" width="<%=width %>" bordercolor="#3A9D99" bgcolor="#D8EFEE">       
          <tr class="sbody">
            <td valign="middle" align="center" width="60" > 
            	  <!--	<DIV onmousedown= "fmenu0() style="cursor:hand;"> -->
            		<% 
            			if(count > 0 ){
							out.print("<u>");
            		    	out.print("<DIV onmousedown=fmenu"+i+"() style=\"cursor:hand;\">");
            		 	}
            	
            			out.print( ((DataObject)FX008_list2.get(i)).getValue("m_year")+"年/"+((DataObject)FX008_list2.get(i)).getValue("m_quarter")+"季");
            	   
            	  		if(count > 0){
							out.print("</u>");
            		    	out.print("</div>");
            		  	}
            		%>
            </td>
            <td valign="middle" align="center" width="40" > <%= count %> </td>
            <td valign="middle" align="center" width="40" >&nbsp;</td>
            <td valign="middle" align="center" width="80" >&nbsp;</td>
            <td valign="middle" align="center" width="50" >&nbsp;</td>
            <td valign="middle" align="center" width="120">&nbsp;</td>
            <td valign="middle" align="center" width="80" >&nbsp;</td>
            <td valign="middle" align="center" width="70" >&nbsp;</td>
            <td valign="middle" align="center" width="120">&nbsp;</td>
            <td valign="middle" align="center" width="100">&nbsp;</td>
            <td valign="middle" align="center" width="40" >&nbsp;</td>
            <td valign="middle" align="center" width="40" >&nbsp;</td>
            <td valign="middle" align="center" width="40" >&nbsp;</td>
            <td valign="middle" align="center" width="40" >&nbsp;</td>
            <td valign="middle" align="center" width="100">&nbsp;</td>
            <td valign="middle" align="center" width="100">&nbsp;</td>
            <td valign="middle" align="center" width="80" >&nbsp;</td>
            <td valign="middle" align="center" width="70" >&nbsp;</td>
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
		//在這筆年季之下有多少筆資料 列出來
          for(int j=0;j < count;j++){
          	
          	//取得資料
          	DataObject bean = (DataObject)FX008_list1.get(j+displacement);
          	
          	seq_no			 = bean.getValue("seq_no").toString();
          	m_year			 = bean.getValue("m_year").toString();
          	m_quarter		 = bean.getValue("m_quarter").toString();
          	damage_yn		 = String.valueOf(bean.getValue("damage_yn")); 			
			disposal_fact_yn = String.valueOf(bean.getValue("disposal_fact_yn")); 
			disposal_plan_yn = String.valueOf(bean.getValue("disposal_plan_yn"));  
			auditresult_yn   = String.valueOf(bean.getValue("auditresult_yn")); 	

			//date 從 1900年開始記 月份從0月開始記
			duDate   = (Date)bean.getValue("duredate");
			year     = duDate.getYear()-11;  
			month    = duDate.getMonth()+1; 
			day      = duDate.getDate();
			
			//---------------------
			if(bean.getValue("audit_duredate")!= null){
				get_Cnt_date = (Date)bean.getValue("audit_duredate");			
				Cnt_year  = get_Cnt_date.getYear()-11;
				Cnt_month = get_Cnt_date.getMonth()+1;
				Cnt_date  = get_Cnt_date.getDate();   
				
				//---------------------
				get_ApplyOK_date = (Date)bean.getValue("applyok_date");						
				ApplyOK_y  = get_ApplyOK_date.getYear()-11;
				ApplyOK_m  = get_ApplyOK_date.getMonth()+1;
				ApplyOK_d  = get_ApplyOK_date.getDate();   
				
				//--------------------
				get_BOAF_date = (Date)bean.getValue("report_boaf_date");
				BOAF_y  = get_BOAF_date.getYear()-11;
				BOAF_m  = get_BOAF_date.getMonth()+1;
				BOAF_d  = get_BOAF_date.getDate();   
			}  

		
			//判斷要不要加超連結
			if(FX008_Lock.contains(m_year+m_quarter)){ 
				 hyper_yn = false;
			}
			else{
				 hyper_yn = true;
			}
		
			bgcolor = (j % 2 == 0)?"#E7E7E7":"#D3EBE0";
%>
          		<tr class="sbody" bgcolor="<%= bgcolor %>">
            		<td valign="middle" align="center" width="60" >&nbsp;</td>
            		<td valign="middle" align="center" width="40" >&nbsp;</td>
            		<td valign="middle" align="center" width="40" >
            				<!-- -------做超連結------ -->
            				<% if(hyper_yn){%>
            					<a href="FX008AW.jsp?act=Edit&seq_no=<%=seq_no%>&m_year=<%=m_year%>&m_quarter=<%=m_quarter%>&bank_no=<%=bank_no %>"> 
							<%}%>
            				<%= bean.getValue("dureassure_no") %>
							<% if(hyper_yn){%>
            					</a>	
							<%}%>
            		</td>
            		<td valign="middle" align="center" width="80" height="9"> 
            				<!-- -------做超連結------ -->
							<% if(hyper_yn){%>
            					<a href="FX008AW.jsp?act=Edit&seq_no=<%=seq_no%>&m_year=<%=m_year%>&m_quarter=<%=m_quarter%>&bank_no=<%=bank_no %>"> 
										<%}%>
            				<%=  bean.getValue("debtname") %> 
							<% if(hyper_yn){%>
            					</a>
										<%}%>	
            		</td>
            		<td valign="middle" align="center" width="50" > <%= year %><%="/"+month+"/"%><%= day%></td>
            		<td valign="middle" align="left"   width="120"> <%= bean.getValue("dureassuresite") %></td>
            		<td valign="middle" align="right"  width="80" >	<%= "$"+Utility.setCommaFormat( bean.getValue("accountamt").toString())%> </td>
            		<td valign="middle" align="center" width="70" > <%= bean.getValue("applydelayyear")+"年"+bean.getValue("applydelaymonth")+"個月" %> </td>
            		<td valign="middle" align="left"   width="120"> <%= bean.getValue("applydelayreason") %></td>
            		<td valign="middle" align="center" width="100">
            		<%
            			if(bean.getValue("audit_duredate")!= null){ 
            				out.print(bean.getValue("audit_applydelayyear")+"年"+bean.getValue("audit_applydelaymonth")+"個月/<br>");
            				out.print(Cnt_year+"年"+Cnt_month+"月"+Cnt_date+"日");
            			}else{
            				out.print("尚未審核");
            			}
            		%>
            		</td>
            		<td valign="middle" align="center" width="40" > &nbsp;<%= damage_yn	%> </td>
            		<td valign="middle" align="center" width="40" > &nbsp;<%= disposal_fact_yn %> </td>
            		<td valign="middle" align="center" width="40" > &nbsp;<%= disposal_plan_yn %> </td>
            		<td valign="middle" align="center" width="40" > &nbsp;<%= auditresult_yn   %> </td>
            		<td valign="middle" align="center" width="100">
            		<%
            			if(bean.getValue("audit_duredate")!= null){ 
            				out.print(bean.getValue("applyok_docno")+"/<br>");
            				out.print(ApplyOK_y+"年"+ApplyOK_m+"月"+ApplyOK_d+"日");
            			}else{
            				out.print("尚未審核");
            			}
            		%>
            		</td>
            		<td valign="middle" align="center" width="100"> 
            		<%
            			if(bean.getValue("audit_duredate")!= null){ 
            				out.print(bean.getValue("report_boaf_docno")+"/<br>");
            				out.print(BOAF_y+"年"+BOAF_m+"月"+BOAF_d+"日");
            			}else{
            				out.print("尚未審核");
            			}
            		%>
            		</td>
            		<td valign="middle" align="center" width="80" > <%= bean.getValue("user_id") %>  /<br><%= bean.getValue("user_name") %></td>
            		<td valign="middle" align="center" width="70" > <%= bean.getValue("update_date") %> </td>
          		</tr>
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
      <td width="<%=width %>">
          <center>
          <jsp:include page="getMaintainUser.jsp?width=1340" flush="true" /><!--載入維護者資訊 -->
          </center>
      </td>
    </tr>
    <tr class="sbody">
      <td height="145" width="<%=width %>" >
          <table border="0" cellspacing="1" width="500" height="150">
            <tr>
              <td colspan="2" width="500" height="32"><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle"><font color="#007D7D" size="3">使用說明:</font>
              </td>
            </tr>
            <tr class="sbody">
              <td width="10" height="111"></td>
              <td width="490" height="111">
              	<ul>
                    <li><font color="#0000FF">申報時程</font>:每年1、4、7及10月之15日前分別完成上一季之資料申報</li>
                    <li><font color="#0000FF">[轉錄上季申報資料</font>]:(1)輸入申報之年季(2)按[載入上季申報資料]按鈕(註:載入後再逐筆辦理本季資料異動事宜</li>
                    <li><font color="#0000FF">[新增]:</font>輸入申報之年季再點選【新增】按鈕可新增該年季「承受擔保品延長處分申報情形」之資料。</li>
                    <li><font color="#0000FF">[修正]</font>:點選該[申報年季]之[總筆數]即會自動展開(或收合)明細記錄,再點選[編號]欄即可變更該申報之資料。</font></li>
                    <li>本表係按最近的[申報年季]先排序,依此類推..........</font></li>
                    <li><font color="#FF0000">如果在[承受擔保品編號]欄位沒有出現底線,表示巳辦理[鎖定]或縣市政府巳完成評註,僅提供查詢不可再異動</font></li>
                </ul>
              </td>
            </tr>
         </table>
      </td>
    </tr>
    <tr>
      <td height="18" width="<%=width %>"></td>
    </tr>

  </table>
  </form>

</body>

</html>

<!-- --------------------------------------------------- -->
<script language='JavaScript'>
<%
	for(int k=0;k < FX008_list2.size(); k++){
%>

		function fmenu<%= k %>(){
			if( menu<%= k %>.style.display == "none")
				menu<%= k %>.style.display = "block";
			else
				menu<%= k %>.style.display = "none";
		}
<%
	}
%>
</script>
<!-- -------------------------------------------------	-->
