<%
// 95.05.24 fix 獨立成統一農(漁)貸資料辦理情形維護 by 2495
// 99.12.06 fix sqlInjection by 2808
//102.12.30 add 本月平均利率/本年平均利率 by 2295
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
<%
  System.out.println("2 TEST act=new&loaddata=false");
  
	List WLX07_M_Credit= (request.getAttribute("WLX07_M_Credit")==null)?null:(List)request.getAttribute("WLX07_M_Credit");
	List inidate= (request.getAttribute("WLX07_INI")==null)?null:(List)request.getAttribute("WLX07_INI");  //取得初始年月
	List hisdata= (request.getAttribute("WLX07_HIS")==null)?null:(List)request.getAttribute("WLX07_HIS");  //取得已申報資料
	int flage=1,flag_house=0,flag_cash=0;    
	List paramList =new ArrayList () ;
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");
    String loaddata = ( request.getParameter("loaddata")==null ) ? "" : (String)request.getParameter("loaddata");
    String syear = ( request.getParameter("S_YEAR")==null ) ? "" : (String)request.getParameter("S_YEAR");
    String smonth = ( request.getParameter("S_MONTH")==null ) ? "" : (String)request.getParameter("S_MONTH");
    String editable = ( request.getParameter("editable")==null ) ? "" : (String)request.getParameter("editable");
    
    if(loaddata.equals("false")){
    flage=0;
    }
        
    if(WLX07_M_Credit!=null && !loaddata.equals("loaded")){//若不是要載入上月資料並且是修改資料時
       syear =String.valueOf(((DataObject)WLX07_M_Credit.get(0)).getValue("m_year"));
       smonth = String.valueOf(((DataObject)WLX07_M_Credit.get(0)).getValue("m_month"));
       
    }
		
	//初始年月
	String iniyear="", inimonth="";
    if(inidate!=null){
       iniyear = String.valueOf(((DataObject)inidate.get(0)).getValue("m_year"));
       inimonth =String.valueOf(((DataObject)inidate.get(0)).getValue("m_month"));
       
	}
		
	Properties permission = ( session.getAttribute("FX007WB")==null ) ? new Properties() : (Properties)session.getAttribute("FX007WB");
	if(permission == null){
       System.out.println("FX007WB_List.permission == null");
       
    }else{
       System.out.println("FX007WB_List.permission.size ="+permission.size());
      
    }

	String bank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");
	String nowtbank_no =  ( request.getParameter("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");
	if(!nowtbank_no.equals("")){
	   session.setAttribute("nowtbank_no",nowtbank_no);//將已點選的tbank_no寫入session
	   
	}
	bank_no = ( session.getAttribute("nowtbank_no")==null ) ? bank_no : (String)session.getAttribute("nowtbank_no");
	  String creditmonth_cnt="";
		String creditmonth_amt="";
		String credityear_cnt_acc="";
		String credityear_amt_acc="";
		String credit_cnt="";
		String credit_bal="";				
		String overcreditmonth_cnt="";
		String overcreditmonth_amt="";
		String overcredit_cnt="";
		String overcredit_bal="";
		String creditmonth_avgrate="";//102.12.30 add 本月平均利率/本年平均利率
		String credityear_avgrate="";
	
	if(WLX07_M_Credit!=null){	
    	creditmonth_cnt = String.valueOf(((DataObject)WLX07_M_Credit.get(0)).getValue("creditmonth_cnt")); 
		creditmonth_amt = String.valueOf(((DataObject)WLX07_M_Credit.get(0)).getValue("creditmonth_amt")); 
		credit_cnt = String.valueOf(((DataObject)WLX07_M_Credit.get(0)).getValue("credit_cnt")); 
		credit_bal = String.valueOf(((DataObject)WLX07_M_Credit.get(0)).getValue("credit_bal")); 	
		credityear_cnt_acc = String.valueOf(((DataObject)WLX07_M_Credit.get(0)).getValue("credityear_cnt_acc")); 
		credityear_amt_acc = String.valueOf(((DataObject)WLX07_M_Credit.get(0)).getValue("credityear_amt_acc")); 
		
		overcreditmonth_cnt = String.valueOf(((DataObject)WLX07_M_Credit.get(0)).getValue("overcreditmonth_cnt"));	
		overcreditmonth_amt = String.valueOf(((DataObject)WLX07_M_Credit.get(0)).getValue("overcreditmonth_amt")); 
		overcredit_cnt = String.valueOf(((DataObject)WLX07_M_Credit.get(0)).getValue("overcredit_cnt")); 
		overcredit_bal = String.valueOf(((DataObject)WLX07_M_Credit.get(0)).getValue("overcredit_bal")); 
		creditmonth_avgrate = String.valueOf(((DataObject)WLX07_M_Credit.get(0)).getValue("creditmonth_avgrate")); 
		credityear_avgrate = String.valueOf(((DataObject)WLX07_M_Credit.get(0)).getValue("credityear_avgrate")); 
	}

	//取得上個月的資料 本月 及 上月累計
	String lastyear="";
	String lastmonth="";	
	String creditcnt ="0"; 
	String creditamt ="0";
	String overcreditcnt ="0"; 
	String overcreditamt ="0";
	String sqlTemp="";
	List WLX07_M_IMPORTANT_LAST=null;
	
	if(smonth.equals("1")){ 
	   lastyear=String.valueOf(Integer.parseInt(syear)-1); 
	   lastmonth ="12";
    }else { 
       lastyear = syear;
       lastmonth = String.valueOf(Integer.parseInt(smonth)-1); 
    } 
   
    if(flage==1 && act.equals("new"))
    {      
	    if(!smonth.equals("1"))
	    {
	        System.out.println("ENTER.................");
				  // add by 2495
					int tyear=Integer.parseInt(syear);
					int tmonth=Integer.parseInt(smonth);
					int totalmonth=tyear*12+tmonth;
			    //  93*12+1=1117 
					if(totalmonth>1117)
					{
						tmonth=tmonth-1;			
						smonth=Integer.toString(tmonth);
					}
				
			List dbData = getWLX07_M_IMPORTANT(bank_no,syear,smonth);
    		System.out.println("flage==1 dbData="+dbData.size());
			credityear_cnt_acc = String.valueOf(((DataObject)dbData.get(0)).getValue("credityear_cnt_acc")); 
			credityear_amt_acc = String.valueOf(((DataObject)dbData.get(0)).getValue("credityear_amt_acc")); 
	  		System.out.println("credityear_cnt_acc="+credityear_cnt_acc);
	  		System.out.println("credityear_amt_acc="+credityear_amt_acc);
	  		tmonth=Integer.parseInt(smonth);	  
	  		tmonth++; 
	  		smonth=Integer.toString(tmonth);
	  	}	
	  }
	  int ini_credityear_cnt_acc=0,ini_credityear_amt_acc=0;
	  if(!credityear_cnt_acc.equals(""))
	  {
	    ini_credityear_cnt_acc=Integer.parseInt(credityear_cnt_acc);
      }
              
    if(!credityear_amt_acc.equals(""))
	  {
	    ini_credityear_amt_acc=Integer.parseInt(credityear_amt_acc);
    } 
    //95.09.07 add by 2495
    if(nowtbank_no.equals(""))
    		nowtbank_no = bank_no;
    int flag0=0;		
    paramList.clear() ;
    paramList.add(nowtbank_no) ;
    List dbDataWLX07_M_Credit = DBManager.QueryDB_SQLParam("select count(*) as test0 from WLX07_M_Credit where bank_no=? ",paramList,"test0");              
    String strcount = (((DataObject)dbDataWLX07_M_Credit.get(0)).getValue("test0")==null ) ? "0" : (((DataObject)dbDataWLX07_M_Credit.get(0)).getValue("test0")).toString();
	System.out.println(".......size()="+dbDataWLX07_M_Credit.size());
    System.out.println("act="+act);
	System.out.println("strcount="+strcount);    
    int count = Integer.parseInt(strcount); 
    if((act.equals("new") && count==0 && !smonth.equals("1"))||(act.equals("Edit") && count==1 && !smonth.equals("1"))){flag_house=1;flag_cash=1;}
    
%>
<%!
private List getWLX07_M_IMPORTANT(String bank_no, String myear, String mmonth){
    		//程序為顯示畫面，查詢條件
    		List paramList = new ArrayList() ;
    		String sqlCmd = "select * from WLX07_M_Credit where bank_no=? and m_year=? and m_month= ? ";
    		paramList.add(bank_no) ;
    		paramList.add(myear) ;
    		paramList.add(mmonth) ;
         	//System.out.println(" sqlCmd="+sqlCmd);
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"m_year,m_month,creditmonth_cnt,creditmonth_amt,credityear_cnt_acc,"+
            									   "credityear_cnt_acc,credityear_amt_acc,credit_cnt,credit_bal,overcreditmonth_cnt,overcreditmonth_amt,overcredit_cnt,overcredit_bal,update_date");     
            return dbData;
    }
%>

<script language="javascript" src="js/FX007WB.js"></script>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<head>
<title>統一農(漁)貸資料辦理情形維護</title>
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

<!--pushData('<%=creditcnt%>','<%=creditamt%>','<%=overcreditcnt%>','<%=overcreditamt%>','<%=loaddata%>');-->
//-->
</script>
</head>
<body leftmargin="15" topmargin="0">
<%if(act.equals("modify")){%>
<script language="javascript" type="text/JavaScript">
alert("已有該年月資料，將進行修改動作!");
</script>
<%}%>

<!--95/01/09 fix by 4180 若上月份無資料，則無法新增資料-->
<%if(loaddata.equals("false") && hisdata.size()!=0){%>
<script language="javascript" type="text/JavaScript">
alert("請完成前一個月的資料申報後，再辦理本月份資料申報");
history.go(-1);
</script>
<%}%>
<!--95/01/09 fix by 4180 若下月份已申報資料，則無法修改資料-->
<%if(editable.equals("false")){%>
<script language="javascript" type="text/JavaScript">
alert("次一個月的申報資料巳建檔，僅提供查詢功能，該月份不可再辦理修改");
history.go(-1);
</script>
<%}%>
<!--95/01/09 fix by 4180 若不為初始申報年月，則提醒使用者-->
<%if((!syear.equals(iniyear) || !smonth.equals(inimonth)) && inidate!=null 
	  && hisdata.size()==0 ){%>
<script language="javascript" type="text/JavaScript">  
  if(confirm("1.農金局規定本案起始的申報年月為<%=iniyear%>年<%=inimonth%>月，\n"+
            "   與你第一次將申報年月(畫面.申報年月) 不一致!\n"+
      		"2.如果貴單位是<%=iniyear%>年<%=inimonth%>月前成立的單位，\n"+
            "   請務必從<%=iniyear%>年<%=inimonth%>月的資料開始申報。\n"+
            "3.如果<%=iniyear%>年<%=inimonth%>月後才成立的單位，則從成立之\n"+
            "   當月起之資料申報。\n"+
            "4.資料自[起始申報]月份起，不管有無異動均須逐月申報。")
            )
	 history.go(1);
  else
    history.go(-1);     

</script>
<%}%>

<form method="post" name="atmdata">
<table width="321" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
        <td width="618"><img src="images/space_1.gif" width="12" height="12"></td>
      </tr>
      <tr>
          <td bgcolor="#FFFFFF" width="618">
      	  <table width="569" border="0" align="center" cellpadding="0" cellspacing="0" height="328">      	  
              <tr>
                <td width="734" height="18">
                <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="223" align="center"><img src="images/banner_bg1.gif" width="150" height="17" align="left"></td>
                      <td width="288" align="center"><b>
                        <center><font color="#000000" size="4">統一農(漁)貸資料辦理情形維護</font>
                        </center>
                        </b> </td>
                      <td width="223" valign="middle"><img src="images/banner_bg1.gif" width="150" height="17" align="right"></td>
                    </tr>
                </table>
                </td>
              </tr>
              <tr>
                <td width="734" height="150">
                <table width="734" border="0" align="center" cellpadding="0" cellspacing="0" height="8">
                    <tr><td width="734" height="10"></tr>
                    <tr> 
                      <div align="right"><jsp:include page="getLoginUser.jsp?width=734" flush="true" /></div> 
                    </tr>   
                    <tr>
            		<td width="734" height="82"><table width=734 border=1 align=center cellpadding="1" cellspacing="1" bordercolor="#3A9D99" height="257">
            		<tr>
            			<td class="sbody" width='218' bgcolor='#D8EFEE' align='left' height="27" colspan="4">金融機構代號</td>
            			<td class="sbody" width='491' bgcolor='e7e7e7' height="33"><%=bank_no%></td>
            		</tr>
            		<tr class="sbody">
            			<td width='218' bgcolor='#D8EFEE' align='left' height="28" colspan="4">申報年月<font color="red">*</font></td>
            			<td width='491' bgcolor='e7e7e7' height="28">
                            <p style="margin-left: 0; margin-top: 2; margin-bottom: 2">
                            <input type='text' name='nyear' value="<%=syear%>" size='3' maxlength='3' onblur='CheckYear(this)'
                            <%
                             if(!syear.equals("")){
                              out.print("disabled");
                              }
                            %>>	
                    		<font color='#000000'>年               	
                    		<select name=nmonth
                    		<%
                    		if(!smonth.equals("")){
                    		out.print("disabled");
                    		}
                    		%> size="1">
                    		<option></option>
                    		<%
                     		for(int i=0;i<=12;i++){
                      			if(!smonth.equals(String.valueOf(i))){                      			    
                       		  		out.print("<option value="+i+" >"+i+"</option>");
                      			}else{                      				  
                        	  		out.print("<option value="+i+" selected >"+i+"</option>");
                        		}  
                     		}%>
                    		</select>月       
                    		<%        
                    		if(!smonth.equals("") && !syear.equals("")){
                    		%>
                    		<input type="hidden" name="hyear" value="<%=syear%>">               
                    		<input type="hidden" name="hmonth" value="<%=smonth%>"><%}%>               
                    		<%if(loaddata.equals("ok")){%>              
                    		<input type="button" name="load" value="載入上月申報資料" onClick="javascript:doSubmit(this.document.forms[0],'load');">              
                    		</p>
                    		<%}%>
                    		</font>
                    	</td>
            	    </tr>
            	<tr class="sbody">
            <td width='22' align='left' bgcolor='#D8EFEE' height="256" rowspan="12">
            <p align="center">&nbsp;統一農(漁)貸資料</p>
            </td>
            <td width='14' align='left' bgcolor='#D8EFEE' height="135" rowspan="8"><font color="#0000FF">貸放資料</font>
            </td>
            <td width='74' align='center' bgcolor='#D8EFEE' height="39" rowspan="3"><font color="#0000FF">本月新增</font>
            </td>
            <td width='86' align='left' bgcolor='#D8EFEE' height="18"><font color="#0000FF">貸放戶數</font><font color="#FF0000">*</font>
            </td>
                       

            <td width='491' bgcolor='e7e7e7' height="18">
            <input type='text' name='creditmonth_cnt' value="<%=Utility.setCommaFormat(creditmonth_cnt)%>" size='10' maxlength='10' onFocus='this.value=changeVal(this)' 
            
            onBlur='checkTotal_house(this.document.forms[0],<%=ini_credityear_cnt_acc%>,<%=flag_house%>);this.value=changeStr(this);' style='text-align: right;'  >
            </td>
	    </tr>
            <tr class="sbody">
            <td width='86' align='left' bgcolor='#D8EFEE' height="15"><font color="#0000FF">貸放金額</font><font color="#FF0000">*</font>
            </td>
             
            <td width='491' bgcolor='e7e7e7' height="15">
             <input type='text' name='creditmonth_amt' value="<%=Utility.setCommaFormat(creditmonth_amt)%>" size='20' maxlength='20' onFocus='this.value=changeVal(this)' 
             onBlur='checkTotal_cash(this.document.forms[0],<%=ini_credityear_amt_acc%>,<%=flag_cash%>);this.value=changeStr(this);' style='text-align: right;'>
             <font face="標楷體">                
            <font color='red' size="1">(單位:新台幣元)</font>       
            </font>
            </td>
	    </tr>
 		<tr class="sbody">
            <td width='86' align='left' bgcolor='#D8EFEE' height="15"><font color="#0000FF">平均利率</font><font color="#FF0000">*</font>
            </td>
             
            <td width='491' bgcolor='e7e7e7' height="15">
             <input type='text' name='creditmonth_avgrate' value="<% out.print((creditmonth_avgrate == null)? "":creditmonth_avgrate.toString()); %>" size='20' maxlength='20' onFocus='this.value=changeVal(this)' 
             onBlur='this.value=changeStr(this);' style='text-align: right;'>%
             <font color='red' size="1" face="標楷體">(單位：x.xxx%)</font><br><font color='red' size="2" face="標楷體">            
            平均利率=本月貸放案件核定利率之合計數/本月貸放戶數</font></td>
	    </tr>

            <tr class="sbody">
            <td width='74' align='center' bgcolor='#D8EFEE' height="56" rowspan="3"><font color="#0000FF">&nbsp;&nbsp;&nbsp;&nbsp;本年累計</font>
            </td>
            
           
            <td width='86' align='left' bgcolor='#D8EFEE' height="28"><font color="#0000FF">貸放戶數</font>
            </td>
            <td width='491' bgcolor='e7e7e7' height="28">
            <% 
               if(nowtbank_no.equals(""))
               			nowtbank_no = bank_no;
               paramList.clear() ;
               paramList.add(nowtbank_no) ;
               dbDataWLX07_M_Credit = DBManager.QueryDB_SQLParam("select count(*) as test0 from WLX07_M_Credit where bank_no=?",paramList,"test0");                             
               strcount = (((DataObject)dbDataWLX07_M_Credit.get(0)).getValue("test0")==null ) ? "0" : (((DataObject)dbDataWLX07_M_Credit.get(0)).getValue("test0")).toString();
							 System.out.println("strcount="+strcount);  
               System.out.println(".......size()="+dbDataWLX07_M_Credit.size());
               count = Integer.parseInt(strcount); 
             if((act.equals("new") && count==0 && !smonth.equals("1"))||(act.equals("Edit") && count==1 && !smonth.equals("1"))){flag_house=1;
            %>	
            
            <input type='text' name='credityear_cnt_acc' value="<%=Utility.setCommaFormat(credityear_cnt_acc)%>" size='10' maxlength='10' onFocus='this.value=changeVal(this)' 
            onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;'  >
            <%}else{ %>
            	<input type='text' name='credityear_cnt_acc'   value="<%=Utility.setCommaFormat(credityear_cnt_acc)%>" size='10' maxlength='10' onFocus='this.value=changeVal(this)' 
            onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='color:#808080;text-align: right;'  readonly>
            <%}%>
            </td>
                    
	    </tr>
            <tr class="sbody">
           
           
            <td width='86' align='left' bgcolor='#D8EFEE' height="22"><font color="#0000FF">貸放金額</font>
            </td>
            <td width='491' bgcolor='e7e7e7' height="22">
            	
             <% 
               if(nowtbank_no.equals(""))
               			nowtbank_no = bank_no;
               paramList.clear() ;
               paramList.add(nowtbank_no) ;
               dbDataWLX07_M_Credit = DBManager.QueryDB_SQLParam("select count(*) as test0 from WLX07_M_Credit where bank_no=?",paramList,"test0");                             
               strcount = (((DataObject)dbDataWLX07_M_Credit.get(0)).getValue("test0")==null ) ? "0" : (((DataObject)dbDataWLX07_M_Credit.get(0)).getValue("test0")).toString();
							 System.out.println("strcount="+strcount);  
               System.out.println(".......size()="+dbDataWLX07_M_Credit.size());
               count = Integer.parseInt(strcount); 
             if((act.equals("new") && count==0 && !smonth.equals("1"))||(act.equals("Edit") && count==1 && !smonth.equals("1"))){flag_cash=1;
            %>	
            <input type='text' name='credityear_amt_acc' value="<%=Utility.setCommaFormat(credityear_amt_acc)%>" size='20' maxlength='20' onFocus='this.value=changeVal(this)' 
             onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;'  >
             <%}else{%>
             	 <input type='text' name='credityear_amt_acc' value="<%=Utility.setCommaFormat(credityear_amt_acc)%>" size='20' maxlength='20' onFocus='this.value=changeVal(this)' 
             onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='color:#808080; text-align: right;'  readonly>
             <%}%>
             <font color='red' size="1" face="標楷體">(單位:新台幣元)</font>               
             </td>            
	     </tr>
	   
	      <tr class="sbody">           
           
            <td width='86' align='left' bgcolor='#D8EFEE' height="22"><font color="#0000FF">平均利率</font><font color="#FF0000">*</font>
            </td>
           <td width='491' bgcolor='e7e7e7' height="15">
             <input type='text' name='credityear_avgrate' value="<% out.print((credityear_avgrate == null)? "":credityear_avgrate.toString()); %>" size='20' maxlength='20' onFocus='this.value=changeVal(this)' 
             onBlur='this.value=changeStr(this);' style='text-align: right;'>%<font color='red' size="1" face="標楷體">(單位：x.xxx%)</font>    <br>             
             	<font color='red' size="2" face="標楷體">
					平均利率=本年貸放案件核定利率之合計數/本年貸放戶數 </font></td>
	     </tr>

	   
            <tr class="sbody">
            <td width='70' align='left' bgcolor='#D8EFEE' height="28" rowspan="2">
              <p align="center"><font color="#0000FF">貸放</font><font color="#0000FF">餘額</font></p>
            </td>
            <td width='86' align='left' bgcolor='#D8EFEE' height="14">
              <font color="#0000FF">戶數<font color="red">*</font></font>
            </td>
            <td width='491' bgcolor='e7e7e7' height="14">
            <input type='text' name='credit_cnt' value="<%=Utility.setCommaFormat(credit_cnt)%>" size='10' maxlength='10' onFocus='this.value=changeVal(this)' 
            onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;' >
             </td>
	     </tr>
	   
            <tr class="sbody">
            <td width='86' align='left' bgcolor='#D8EFEE' height="14">
              <font color="#0000FF">餘額<font color="red">*</font></font>
            </td>
            <td width='491' bgcolor='e7e7e7' height="14">
             <input type='text' name='credit_bal' value="<%=Utility.setCommaFormat(credit_bal)%>" size='20' maxlength='20' onFocus='this.value=changeVal(this)' 
             onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;'>
             <font color='red' size="1" face="標楷體">(單位:新台幣元)</font>              
             </td>
	     </tr>
            <tr class="sbody">
            <td width='14' align='left' bgcolor='#D8EFEE' height="115" rowspan="4"><font color="#FF0000">逾放資料</font>
            </td>
            <td width='74' align='left' bgcolor='#D8EFEE' height="19" rowspan="2">
              <p align="center"><font color="#FF0000">本月新增</font>
              </p>
            </td>
            <td width='86' align='left' bgcolor='#D8EFEE' height="1"><font color="#FF0000">逾放戶數*</font>
            </td>
            <td width='491' bgcolor='#D3EBE0' height="1">
            <input type='text' name='overcreditmonth_cnt' value="<%=Utility.setCommaFormat(overcreditmonth_cnt)%>" size='10' maxlength='10' onFocus='this.value=changeVal(this)' 
             onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;'>　 
            　   
            
             </td>
	     </tr>
            <tr class="sbody">
            <td width='86' align='left' bgcolor='#D8EFEE' height="13"><font color="#FF0000">逾放金額*</font>
            </td>
            <td width='491' bgcolor='#D3EBE0' height="13">
             <input type='text' name='overcreditmonth_amt' value="<%=Utility.setCommaFormat(overcreditmonth_amt)%>" size='20' maxlength='20' onFocus='this.value=changeVal(this)' 
             onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;'><font color='red' size="1" face="標楷體">(單位:新台幣元)</font>     
             </td>
	     </tr>
            <tr class="sbody">
            <td width='70' align='center' bgcolor='#D8EFEE' height="51" rowspan="2"><font color="#FF0000">&nbsp;&nbsp;&nbsp;&nbsp;</font><font color="#FF0000">逾放</font><font color="#FF0000">餘額</font>
            </td>
            <td width='86' align='center' bgcolor='#D8EFEE' height="33">
              <p align="left"><font color="#FF0000">戶數*</font>
            </td>
            <td width='491' bgcolor='#D3EBE0' height="26">
             <input type='text' name='overcredit_cnt' value="<%=Utility.setCommaFormat(overcredit_cnt)%>" size='10' maxlength='20' onFocus='this.value=changeVal(this)' 
             onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;'>　 
             　 　

             
             </td>
	     </tr>
            <tr class="sbody">
            <td width='86' align='left' bgcolor='#D8EFEE' height="18">
              <font color="#FF0000">餘額<font color="red">*</font></font>
            </td>
            <td width='491' bgcolor='#D3EBE0' height="18">
              <input type='text' name='overcredit_bal' value="<%=Utility.setCommaFormat(overcredit_bal)%>" size='20' maxlength='20' onFocus='this.value=changeVal(this)' 
             onBlur='checkPoint_focus(this);this.value=changeStr(this);' style='text-align: right;'>
             <font color='red' size="1" face="標楷體">(單位:新台幣元)</font>          
              　  
            
             </td>
	     </tr>
     </Table>
   </td>
 </tr>


     <td width="734" height="21"><div align="right"><div align="right">
	<div align="right"><jsp:include page="getMaintainUser.jsp?width=734" flush="true" /></div>
    </div>
    </div></td>
      </table></td>
  </tr>
  <tr>
                <td width="734" height="123"><table width="734" border="0" cellpadding="1" cellspacing="1" class="sbody" height="176">
                    <tr>
                      <td colspan="2" width="734" height="41">
                      <div align="center">
                    <table width="243" border="0" cellpadding="1" cellspacing="1">
                      <tr>
                       <%if(act.equals("new")){
                       if(permission != null && permission.get("A") != null && permission.get("A").equals("Y")){//Add
                      %>
                <td width="66"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'add',<%=ini_credityear_cnt_acc%>,<%=ini_credityear_amt_acc%>,<%=flag_house%>,<%=flag_cash%>);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_confirmb.gif',1)"><img src="images/bt_confirm.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a></div></td>
                <% } }else{

                if(permission != null && permission.get("U") != null && permission.get("U").equals("Y")){//Update %>
                <td width="66"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'modify');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_updateb.gif',1)"><img src="images/bt_update.gif" name="Image103" width="66" height="25" border="0" id="Image103"></a></div></td>
                <td width="66"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'delete');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image104','','images/bt_deleteb.gif',1)"><img src="images/bt_delete.gif" name="Image104" width="66" height="25" border="0" id="Image104"></a></div></td>
              <% }}%>
                        <td width="66"><div align="center"><a href="javascript:AskReset(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_cancelb.gif',1)"><img src="images/bt_cancel.gif" name="Image105" width="66" height="25" border="0" id="Image105"></a></div></td>
                        <td width="93"><div align="center"><a href="javascript:doSubmit(this.document.forms[0],'returnList');"  onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image106','','images/bt_05b.gif',1)"><img src="images/bt_05.gif" name="Image106" width="80" height="25" border="0" id="Image106"></a></div></td>
                      </tr>
                    </table>
                  </div>
                     </td>
                    </tr>
                    <tr>
                      <td colspan="2" width="583" height="41"><font color='#990000'><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle"><font color="#007D7D" size="3">使用說明        
                        :</font></font> </td>
                    </tr>
                    <tr>
                      <td width="16" height="127">&nbsp;</td>
                      <td width="561" height="127">
 <ul>
                      	 
                      	 
                      	  <li class="sbody" >確認輸入資料無誤後, 按<font color="#666666">【確定】即將本表上的資料, 於資料庫中建檔。</li>              
                      	  <li class="sbody" >按<font color="#666666">【修改】即修改的資料,寫入資料庫料庫中。</li>
                          <li class="sbody" >欲重新輸入資料, 按<font color="#666666">【取消】即將本表上的資料清空</li>              
                          <li class="sbody" >如放棄修改或無修改之資料需輸入, 按【回上一頁】]即離開本程式。</li>              
                          <li class="sbody" >【<font color="red">*</font>】為必填欄位。</li>
                          <li class="sbody" >【<font color="red">*</font>】如果沒有申報仍請填「0」</li>
                          <li><font color="red">「統一農(漁)貸資料」係以申報月底日的資料來申報</font></li>	
                          <li><font color="red">「本年累計貸放戶數」、「本年累計貸放金額」<br>
                            (1)「本年累計貸放戶數」、「本年累計貸放金額」係以每年作一累計；故每年元月時，「本月新增  
                            </font>
                            </font>
                          </font>
                            </font><font color="red">
                            貸放戶數」與「本年累計貸放戶數」、「本月新增貸放金額」與「本年累計貸放金額」的數字應相同<br> 
                            (2) 其他2-12月份的「本年累計貸放戶數」應為「本月新增貸放戶數+上月的本年累計貸放戶數」「本年累計貸放金額」應為「本月新增貸放金額+上月的本年累計貸放金額」<br>
                            (3) 起始申報年月時，由農(漁)會信用部負責輸入，後續則由系統提供自動累加功能。<br> 
                            (4) 本月平均利率=本月貸放案件核定利率之合計數/本月貸放戶數。<br> 
                            (5) 平均利率=本年貸放案件核定利率之合計數/本年貸放戶數。
                            </li>   
                        </ul>
                            </font>
                            </font>
                          </font>
 </font>
                      </td>
                    </tr>
                  </table></td>
      </tr>         
	</table>
</table>
</form>
</body>
