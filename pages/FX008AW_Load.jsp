<%
// 94.10.14 first designed by 4180 lilic0c0
// 95.04.25 add 機構名稱(借款人)歸戶明細 by 2495 
// 99.12.06 fix sqlInjection by 2808
//100.01.27 fix bug by 2295
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
<%@ page import="java.util.ArrayList" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Date" %>
<%	//---------------------------------------取得參數---------------------------------------------
	
	String bank_no   = ( request.getParameter("bank_no")==null ) ? "" :request.getParameter("bank_no");
	System.out.println("bank_no ..........="+bank_no);
	String debt_name   = ( request.getParameter("debt_name")==null ) ? "" :request.getParameter("debt_name");
	System.out.println("debt_name ..........="+debt_name);
	List paramList =new ArrayList() ;
	//所有資料 (order by year quarter )
  List dbData = getdebt_name_Data(debt_name,bank_no);
 
	
	//table width
	int width = 1340;//1300 pixel
%>
<%!
   //獲得點選借款人於該bank資料 (BY bank number)==================================================
    private List getdebt_name_Data(String debt_name,String bank_no){
    	
    	//查詢條件    
    	  List paramList = new ArrayList() ;
    	  String sqlCmd = "select M_YEAR,M_QUARTER,BANK_NO,SEQ_NO,DUREASSURE_NO,DEBTNAME,to_char(DUREDATE,'mm/dd/yyyy') as DUREDATE,DUREASSURESITE,ACCOUNTAMT,APPLYDELAYYEAR,APPLYDELAYMONTH,APPLYDELAYREASON,AUDIT_APPLYDELAYYEAR,AUDIT_APPLYDELAYMONTH,to_char(AUDIT_DUREDATE,'mm/dd/yyyy') as AUDIT_DUREDATE,damage_yn,disposal_fact_yn,disposal_plan_yn,auditresult_yn,to_char(APPLYOK_DATE,'mm/dd/yyyy') as APPLYOK_DATE,to_char(REPORT_BOAF_DATE,'mm/dd/yyyy') as REPORT_BOAF_DATE,REPORT_BOAF_DOCNO,user_id,user_name,to_char(update_date,'mm/dd/yyyy') as update_date,applyok_docno ";
    	  sqlCmd +=" from wlx08_s_gage where debtname=? and bank_no=? Order BY  bank_no  desc,M_YEAR desc,M_Quarter  desc,DUREASSURE_NO";
    	  paramList.add(debt_name);
		  paramList.add(bank_no);
    	  //System.out.println("獲得點選借款人於該bank資料 sqlCmd ="+sqlCmd);  
          List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"M_YEAR,M_QUARTER,BANK_NO,SEQ_NO,DUREASSURE_NO,DEBTNAME,DUREDATE,"+
        										   "DUREASSURESITE,ACCOUNTAMT,APPLYDELAYYEAR,APPLYDELAYMONTH,APPLYDELAYREASON,"+
        										   "AUDIT_APPLYDELAYYEAR,AUDIT_APPLYDELAYMONTH,AUDIT_DUREDATE,"+
        										   "damage_yn,disposal_fact_yn,disposal_plan_yn,auditresult_yn,"+
        										   "APPLYOK_DATE,applyok_docno,REPORT_BOAF_DOCNO,REPORT_BOAF_DATE,UPDATE_DATE");
        	return dbData;  
    }//end of getBank_Data  

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

  
    <tr>
      <td height="18" width="<%=width %>"></td>
    </tr>
    <tr>
      <td width="<%=width %>">
     	 <center>
      		<table border="0" cellspacing="1">
          		<tr>
          			<td><img src="images/banner_bg1.gif" width="200" height="17"></td>
          			<td><b><font color='#000000' size="4" width="100">機構名稱(借款人)申辦之歷史明細</font></b></td>
          			<td><img src="images/banner_bg1.gif" width="200" height="17"></td>
          		</tr>
      		</table>
     	</center>
     </td>
    </tr>
    
    <tr>
      <td width="<%=width %>">
          <center>
          
          </center>
      </td>
    </tr>
    
    <tr>
      <td width="<%=width %>">
        <table border="1" cellspacing="1" width="<%=width %>" bordercolor="#3A9D99" >
          <tr class="sbody">
          	<td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="70">信用部</td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="60">申報<br><br>年季</td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="120">承受擔保品<br><br>編號</td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="80">機構名稱<p>(借款人)</p></td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="60">承受日期</td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="120">承受擔保品<p>座落</p></td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="80">帳列金額<p><font color="#ff0000">(新台幣元)</font></td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="70">申請延長<p>期間</p></td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="120">申請延長<p>理由</p></td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="100">(核准申請)<p>延長期間/<br>延長期限</p></td>
            <td bgcolor="#9AD3D0" colspan="4" valign="middle" align="center" width="160">縣/市政府評註</td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="100">(縣市政府)<p>核准文號/<br>核准日期</p></td>
            <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="100">(函報農金局)<p>文號/<br>日期</p></td>
            <td bgcolor="#9AD3D0" valign="middle" align="center" rowspan="2" width="80">異動者<p>帳號/姓名</p></td>
            <td width="70" rowspan="2" align="center" valign="middle" bgcolor="#9AD3D0">異動
                <p>日期</p></td>
            
          </tr>
          <tr class="sbody">
            <td bgcolor="#9AD3D0" valign="middle" align="center" width="40">提足備抵趺價損失</td>
            <td bgcolor="#9AD3D0" valign="middle" align="center" width="40">有積極處分之事實</td>
            <td bgcolor="#9AD3D0" valign="middle" align="center" width="40">未來處分計劃合理可行</td>
            <td bgcolor="#9AD3D0" valign="middle" align="center" width="40">審核結果</td>
            
          </tr>

 <%
 		if(dbData.size()==0)
 		{
 %>		
 
		 	<center>
		 		   
		 		    	<table border="1" cellspacing="1" width="<%=width %>" bordercolor="#3A9D99">
		 		     		<tr class="sbody" bgcolor="#D8EFEE">
		 		      			<td colspan="15" align=center width="<%=width %>">尚無資料</td>
		 		        </tr>
						</table>  	   
					
			</center>	
			   
 <%
		}
		else
		{
					for(int i=0;i<dbData.size();i++){	
					      String bgcolor = (i % 2 == 0)?"#A5E1F0":"#D3EBE0";						     
						  //System.out.println("dbData.SIZE()="+dbData.size());
	      				  String BANK_NO = (String)((DataObject)dbData.get(i)).getValue("bank_no");
	     				  //System.out.println("BANK_NO="+BANK_NO);
	     				  String m_year = Integer.parseInt(Utility.getYear())>99 ? "100" : "99" ; 
						  String sqlCmd1 = "select bank_name from BA01 where bank_no=? and m_year=? ";	
						  paramList.clear();
								paramList.add(BANK_NO) ;
								paramList.add(m_year);
								List dbData1 = DBManager.QueryDB_SQLParam(sqlCmd1,paramList,"");
								System.out.println("dbData2.size() ="+dbData1.size());	
								String	BANK_NAME = (((DataObject)dbData1.get(0)).getValue("bank_name")).toString();
								System.out.println("BANK_NAME="+BANK_NAME);
								
	     					String M_YEAR = (String)((DataObject)dbData.get(i)).getValue("m_year").toString();
	     					System.out.println("M_YEAR="+M_YEAR);
					      String M_QUARTER = (String)((DataObject)dbData.get(i)).getValue("m_quarter").toString();
					     	System.out.println("M_QUARTER="+M_QUARTER); 
					      String DUREASSURE_NO = (((DataObject)dbData.get(i)).getValue("dureassure_no")).toString();
					     	System.out.println("DUREASSURE_NO="+DUREASSURE_NO); 
					      String DUREDATE = (String)((DataObject)dbData.get(i)).getValue("duredate").toString();
					     	System.out.println("DUREDATE="+DUREDATE); 
					      String DUREASSURE_SITE = (String)((DataObject)dbData.get(i)).getValue("dureassuresite").toString();
					    	System.out.println("DUREASSURE_SITE="+DUREASSURE_SITE); 
					      String ACCOUNTAMT= (((DataObject)dbData.get(i)).getValue("accountamt")).toString();
					     	System.out.println("ACCOUNTAMT="+ACCOUNTAMT); 
					      String APPLYDELAYYEAR = (String)((DataObject)dbData.get(i)).getValue("applydelayyear").toString();
					     	System.out.println("APPLYDELAYYEAR="+APPLYDELAYYEAR); 					     
					      String APPLYDELAYMONTH = (String)((DataObject)dbData.get(i)).getValue("applydelaymonth").toString();
					     	System.out.println("APPLYDELAYMONTH="+APPLYDELAYMONTH); 
					      String APPLYDELAYREASON = (String)((DataObject)dbData.get(i)).getValue("applydelayreason").toString();
					      System.out.println("APPLYDELAYREASON="+APPLYDELAYREASON);       
					      String DAMAGE_YN = (((DataObject)dbData.get(i)).getValue("damage_yn")==null ) ? "" : (String)((DataObject)dbData.get(i)).getValue("damage_yn").toString();	
					     	System.out.println("DAMAGE_YN="+DAMAGE_YN); 					     
					      String DISPOSAL_FACT_YN= (((DataObject)dbData.get(i)).getValue("disposal_fact_yn")).toString();
					     	System.out.println("DISPOSAL_FACT_YN="+DISPOSAL_FACT_YN); 
					      String DISPOSAL_PLAN_YN = (String)((DataObject)dbData.get(i)).getValue("disposal_plan_yn");
					     	System.out.println("DISPOSAL_PLAN_YN="+DISPOSAL_PLAN_YN);
					      String AUDITRESULT_YN = (String)((DataObject)dbData.get(i)).getValue("auditresult_yn").toString();
					     	System.out.println("AUDITRESULT_YN="+AUDITRESULT_YN);   
					      String AUDIT_APPLYDELAYYEAR = (String)((DataObject)dbData.get(i)).getValue("audit_applydelayyear").toString();
					     	System.out.println("AUDIT_APPLYDELAYYEAR="+AUDIT_APPLYDELAYYEAR); 
					      String AUDIT_APPLYDELAYMONTH = (String)((DataObject)dbData.get(i)).getValue("audit_applydelaymonth").toString();
					     	System.out.println("AUDIT_APPLYDELAYMONTH="+AUDIT_APPLYDELAYMONTH); 	     
					      String AUDIT_DUREDATE = (((DataObject)dbData.get(i)).getValue("audit_duredate")==null ) ? "null" : (String)((DataObject)dbData.get(i)).getValue("audit_duredate").toString();
					     	System.out.println("AUDIT_DUREDATE="+AUDIT_DUREDATE); 	      
					      String APPLYOK_DOCNO = (((DataObject)dbData.get(i)).getValue("applyok_docno")==null ) ? "" : (String)((DataObject)dbData.get(i)).getValue("applyok_docno").toString();					     	
					      System.out.println("APPLYYOK_DOCNO="+APPLYOK_DOCNO);	     
					      String APPLYOK_DATE = (((DataObject)dbData.get(i)).getValue("applyok_date")==null ) ? "" : (String)((DataObject)dbData.get(i)).getValue("applyok_date").toString();
							  System.out.println("APPLYOK_DATE="+APPLYOK_DATE);				  
					      String REPORT_BOAF_DOCNO = (String)((DataObject)dbData.get(i)).getValue("report_boaf_docno").toString();
					      System.out.println("REPORT_BOAF_DOCNO="+REPORT_BOAF_DOCNO); 
					     	String REPORT_BOAF_DATE = (((DataObject)dbData.get(i)).getValue("report_boaf_date")==null ) ? "" : (String)((DataObject)dbData.get(i)).getValue("report_boaf_date").toString();
				        System.out.println("REPORT_BOAF_DATE="+REPORT_BOAF_DATE); 
				        String USER_ID = (((DataObject)dbData.get(i)).getValue("user_id")==null ) ? "" : (String)((DataObject)dbData.get(i)).getValue("user_id").toString();
				        System.out.println("USER_ID="+USER_ID); 
				        String USER_NAME = (((DataObject)dbData.get(i)).getValue("user_name")==null ) ? "" : (String)((DataObject)dbData.get(i)).getValue("user_name").toString();
				        System.out.println("USER_NAME="+USER_NAME); 
				        String UPDATEDATE = (((DataObject)dbData.get(i)).getValue("update_date")==null ) ? "" : (String)((DataObject)dbData.get(i)).getValue("update_date").toString();
				        System.out.println("USER_NAME="+USER_NAME); 
%>
    
     <tr>
		 		 <td>  
          <tr class="sbody">
          	<td bgcolor='<%=bgcolor%>' rowspan="2" valign="middle" align="center" width="70"><%=BANK_NAME%></td>
            <td bgcolor='<%=bgcolor%>' rowspan="2" valign="middle" align="center" width="60"><%=M_YEAR%>/<%=M_QUARTER%></td>
            <td bgcolor='<%=bgcolor%>' rowspan="2" valign="middle" align="center" width="120"><%=DUREASSURE_NO%></td>
            <td bgcolor='<%=bgcolor%>' rowspan="2" valign="middle" align="center" width="80"><%=debt_name%></td>
            <td bgcolor='<%=bgcolor%>' rowspan="2" valign="middle" align="center" width="50"><%=DUREDATE%></td>
            <td bgcolor='<%=bgcolor%>' rowspan="2" valign="middle" align="center" width="120"><%=DUREASSURE_SITE%></td>
            <td bgcolor='<%=bgcolor%>' rowspan="2" valign="middle" align="center" width="80"><%="$"+Utility.setCommaFormat(ACCOUNTAMT)%></td>
            <td bgcolor='<%=bgcolor%>' rowspan="2" valign="middle" align="center" width="70"><%=APPLYDELAYYEAR%>年<%=APPLYDELAYMONTH%>月</td>
            <td bgcolor='<%=bgcolor%>' rowspan="2" valign="middle" align="center" width="120"><%=APPLYDELAYREASON%></td>
            <td bgcolor='<%=bgcolor%>' rowspan="2" valign="middle" align="center" width="100">
            <%
                if(!AUDIT_APPLYDELAYYEAR.equals("null")){ 
            				out.print(AUDIT_APPLYDELAYYEAR+"年/<br>"+APPLYDELAYMONTH+"個月<br>");
            			}else{
            				out.print("尚未審核");
            			}    
            %></td>
            
            <td bgcolor='<%=bgcolor%>' rowspan="2" valign="middle" align="center" width="40">
            	<%
                if(DAMAGE_YN.equals("Y")||DAMAGE_YN.equals("N")){ 
            				out.print(DAMAGE_YN);
            			}else{
            				out.print("尚未審核");
            			}    
            %></td>
            <td bgcolor='<%=bgcolor%>' rowspan="2" valign="middle" align="center" width="40">
            	<%
                if(DISPOSAL_FACT_YN.equals("Y")||DISPOSAL_FACT_YN.equals("N")){ 
            				out.print(DISPOSAL_FACT_YN);
            			}else{
            				out.print("尚未審核");
            			}    
            %></td>
            <td bgcolor='<%=bgcolor%>' rowspan="2" valign="middle" align="center" width="40">
            	<%
                if(DISPOSAL_PLAN_YN.equals("Y")||DISPOSAL_PLAN_YN.equals("N")){ 
            				out.print(DISPOSAL_PLAN_YN);
            			}else{
            				out.print("尚未審核");
            			}    
            %></td>
            <td bgcolor='<%=bgcolor%>' rowspan="2" valign="middle" align="center" width="40">
            	<%
                if(AUDITRESULT_YN.equals("Y")||AUDITRESULT_YN.equals("N")){ 
            				out.print(AUDITRESULT_YN);
            			}else{
            				out.print("尚未審核");
            			}    
            %></td>
            <td bgcolor='<%=bgcolor%>' rowspan="2" valign="middle" align="center" width="100">
            <%
            			if(!APPLYOK_DOCNO.equals(" ")){ 
            				out.print(APPLYOK_DOCNO+"/<br>");
             				out.print(APPLYOK_DATE);
            				
            			}else{
            				out.print("尚未審核");
            			}
            %></td>
            <td bgcolor='<%=bgcolor%>' rowspan="2" valign="middle" align="center" width="100">
            <%
            			if(!REPORT_BOAF_DOCNO.equals(" ")){ 
            				out.print(REPORT_BOAF_DOCNO+"/<br>");
             				out.print(REPORT_BOAF_DATE);
            				
            			}else{
            				out.print("尚未審核");
            			}
            %></td>
            <td bgcolor='<%=bgcolor%>' valign="middle" align="center" rowspan="2" width="80"><%=USER_ID%>/<%=USER_NAME%></p></td>
            <td width="70" rowspan="2" align="center" valign="middle" bgcolor='<%=bgcolor%>'><%=UPDATEDATE%></td>            
          </tr>
         </td>
 </tr>
   
<%				         			          
					}			
		}
 %>		              
         </table>
      </td>
    </tr>
     <!-- -----------start to parse the form ----------------------------------------------------------- -->
	
  
  </form>
</body>

</html>

