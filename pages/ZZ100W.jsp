<%
//94.11.08 create by 2495
%>

<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.dao.DAOFactory" %>
<%@ page import="com.tradevan.util.dao.RdbCommonDao" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.io.*" %>
<%@ page import="com.tradevan.util.ManualHeadPerson" %>
<%@ page import="com.tradevan.util.AutoHeadPerson" %>		
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/ZZ100W.js"></script>
<script language="javascript" event="onresize" for="window"></script>

<%
	RequestDispatcher rd = null;	
	String actMsg = "";	
	String alertMsg = "";	
	String webURL = "";	
	boolean doProcess = false;
		
	//取得session資料,取得成功時,才繼續往下執行===================================================
  if(session.getAttribute("muser_id") == null)
  {//session timeout	
		System.out.println("ZZ100W.jsp login timeout");   
	   	rd = application.getRequestDispatcher( "/pages/reLogin.jsp?url=LoginError.jsp?timeout=true" );         	   
	   	try{
          rd.forward(request,response);
                   }catch(Exception e){
                       System.out.println("forward Error:"+e+e.getMessage());
                   }
  }
  else
  {
      doProcess = true;

  }
      
  if(doProcess)
  {//若muser_id資料時,表示登入成功====================================================================
	  	String muser_id = (String)session.getAttribute("muser_id");	
		String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");
                String update_muser_id = ( request.getParameter("user_id")==null ) ? "" : (String)request.getParameter("user_id");		
                String update_muser_pwd = ( request.getParameter("user_pwd")==null ) ? "" : (String)request.getParameter("user_pwd");
    		String update_ip_Address = ( request.getParameter("ip_Address")==null ) ? "" : (String)request.getParameter("ip_Address");
        //95.02.06 拿掉 by 2495
        /*
                String StartYear = ( request.getParameter("StartYear")==null ) ? "" : (String)request.getParameter("StartYear");		
                String StartMonth = ( request.getParameter("StartMonth")==null ) ? "" : (String)request.getParameter("StartMonth");
    						String StartDay = ( request.getParameter("StartDay")==null ) ? "" : (String)request.getParameter("StartDay");
                String EndYear = ( request.getParameter("EndYear")==null ) ? "" : (String)request.getParameter("EndYear");		
                String EndMonth = ( request.getParameter("EndMonth")==null ) ? "" : (String)request.getParameter("EndMonth");
    						String EndDay = ( request.getParameter("EndDay")==null ) ? "" : (String)request.getParameter("EndDay");
				*/

		List User_Data = TakeUserData(muser_id);	   	   
     		String muser_name = (String)((DataObject)User_Data.get(0)).getValue("user_name");
                System.out.println("muser_name:"+muser_name);									
    if(!CheckPermission(request))
    {//無權限時,導向到LoginError.jsp
       rd = application.getRequestDispatcher( LoginErrorPgName );        
    }   
    else
    {          
         if(act.equals("Clear"))
         {             	  	
            //String Path = "C:\\Sun\\WebServer6.1\\BOAF\\WLX_WTT07\\HeadPerson.txt";	
	    String Path = Utility.getProperties("headPersonDir");
	    File delFile = new File(Path);   
    	    if(delFile.exists())                                     
 		delFile.delete();	                      	    
    	}
        if(act.equals("New"))
        {             	  	
            actMsg =UpdateDB(update_muser_id,update_muser_pwd,update_ip_Address,muser_id,muser_name);                		                                 	    
    	}
        if(act.equals("Delete")){           
            String Star[] = request.getParameterValues("isDelete");
	    System.out.println("Star[]="+Star);
	    for(int i=0;i<Star.length;i++)
	    {						      
		System.out.println("Star.length="+Star.length); 
		System.out.println("Star[i]="+Star[i]);       	   		
        	DelWTT07(Star[i]);
            }                     
       }
      /* 
       if(act.equals("CreateHeadText"))
       {
	    int west_year=Integer.parseInt(StartYear);				
	    west_year = west_year+1911;
	    StartYear= Integer.toString(west_year);
	    String StartDate=StartYear+StartMonth+StartDay;

	    west_year=Integer.parseInt(EndYear);				
	    west_year = west_year+1911;
	    EndYear= Integer.toString(west_year);
	    String  EndDate=EndYear+EndMonth+EndDay;
            //actMsg = AutoHeadPerson.exeHeadPerson();  		
	    actMsg = ManualHeadPerson.exeHeadPerson(StartDate,EndDate);    	  	            		                                 	    
       }
       */
       //95.02.06 修改 by 2495      
       if(act.equals("CreateHeadText"))
       {
	    	//int west_year=Integer.parseInt(StartYear);				
	    	//west_year = west_year+1911;
	    	//StartYear= Integer.toString(west_year);
	    	//String StartDate=StartYear+StartMonth+StartDay;
				  String StartDate="";
	    	//west_year=Integer.parseInt(EndYear);				
	    	//west_year = west_year+1911;
	    	//EndYear= Integer.toString(west_year);
	    	//String  EndDate=EndYear+EndMonth+EndDay;
     	       //actMsg = AutoHeadPerson.exeHeadPerson();
     	    String EndDate="";     		
	    	actMsg = ManualHeadPerson.exeHeadPerson(StartDate,EndDate);    	  	            		                                 	    
       } 
    }
  }   
%>

<%
        String Path = "C:\\Sun\\WebServer6.1\\BOAF\\WLX_WTT07\\HeadPerson.txt";	
	File objFile = new File(Path);          
        List WTT07_ELM = DBManager.QueryDB_SQLParam("select MUSER_ID,MUSER_PASSWORD,IP_ADDRESS,to_char(UPDATE_DATE,'yyyy/mm/dd') as UPDATE_DATE  from WTT07_ELM",null,"");    	                
%>

<%!
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    private boolean CheckPermission(HttpServletRequest request)
    {//檢核權限 
            	    
    	   boolean CheckOK=false;
    	    HttpSession session = request.getSession();            
            Properties permission = ( session.getAttribute("ZZ100W")==null ) ? new Properties() : (Properties)session.getAttribute("ZZ100W");				                
            if(permission == null){
              System.out.println("WLX_Notify.permission == null");
            }else{
               System.out.println("WLX_Notify.permission.size ="+permission.size());
               
            }
           
        return true;
    }
   
    //取得該user_id之相關資料
    private List TakeUserData(String muser_id){
    		//查詢條件    
    		List paramList =new ArrayList() ;
    		String sqlCmd = "select * from MUSER_DATA where  muser_id =? ";
    		paramList.add(muser_id);
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
        return dbData;
    }

    private String UpdateDB(String update_muser_id,String update_muser_pwd,String update_ip_Address,String muser_id,String muser_name)
    throws Exception{    	
		String sqlCmd = "";		
		String errMsg="";		
		
		List paramList =new ArrayList() ;
		//List updateDBSqlList = new LinkedList();								   				   
		//insert WTT07_ELM_LOG===================================================	
                	    
		sqlCmd = " INSERT INTO WTT07_ELM_LOG(muser_id,muser_password,ip_address,user_id,user_name,update_date,user_id_c,user_name_c,update_date_c,update_type_c) VALUES("
			   + "?,?,?,?,?,sysdate,?,?,sysdate,'U')";
		paramList.add(update_muser_id) ;    
		paramList.add(update_muser_pwd) ;
		paramList.add(update_ip_Address) ;
		paramList.add(muser_id) ;
		paramList.add(muser_name) ;
		paramList.add(muser_id) ;
		paramList.add(muser_name);
		//updateDBSqlList.add(sqlCmd);
        this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
        paramList.clear() ;
		//=========================================================================
		
                sqlCmd = "Delete from WTT07_ELM";
		//updateDBSqlList.add(sqlCmd);
		this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
                

                sqlCmd = " INSERT INTO WTT07_ELM(muser_id,muser_password,ip_address,user_id,user_name,update_date) VALUES("
			   + "?,?,?,?,?,sysdate)";
		paramList.add(update_muser_id) ;
		paramList.add(update_muser_pwd) ;
		paramList.add(update_ip_Address);
		paramList.add(muser_id) ;
		paramList.add(muser_name) ;
		this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
	//	updateDBSqlList.add(sqlCmd);               		            		            		
		
     //           DBManager.updateDB(updateDBSqlList,"ZZ100W.jsp"); 
                
		return errMsg;     
  }
  
  private String DelWTT07(String input_date)throws Exception{
     String sqlCmd = "";		
     String errMsg="";
     
     //List updateDBSqlList = new LinkedList();
     List paramList = new ArrayList() ;
     sqlCmd = "Delete from WTT07 where input_date=TO_TIMESTAMP(?,'mm/dd/yyyy hh24:mi:ss')";
     paramList.add(input_date) ;
     //updateDBSqlList.add(sqlCmd);
     //DBManager.updateDB(updateDBSqlList,"ZZ100W.jsp");
     this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
     return errMsg;  
  }
  private boolean updDbUsesPreparedStatement(String sql ,List paramList) throws Exception{
		List updateDBList = new ArrayList();//0:sql 1:data
	    List updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
		List updateDBDataList = new ArrayList();//儲存參數的List
		
		updateDBDataList.add(paramList);
		updateDBSqlList.add(sql);
		updateDBSqlList.add(updateDBDataList);
		updateDBList.add(updateDBSqlList);
		return DBManager.updateDB_ps(updateDBList) ;
	}         
%>

<html>
<head>
<title>總機構高階主管基本資料維護</title>
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
<input type="hidden" name="seq_no" value="">
<input type="hidden" name="nowDay" value="2005/10/07">
<table width="640" border="0" align="left" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  		  <tr> 
   		   <td><img src="images/space_1.gif" width="12" height="12"></td>
  		  </tr>
          <td bgcolor="#FFFFFF">
		  <table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td><table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="170"><img src="images/banner_bg1.gif" width="170" height="17"></td>
                      <td width="250"><center> 
                        </center>
                        <center> 
                        </center>
                        <font color='#000000' size=4><b><center>農漁會「理監事及負責人」傳輸資料建立維護</center></b></font> </td>
                      <td width="170"><img src="images/banner_bg1.gif" width="170" height="17"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td><img src="images/space_1.gif" width="12" height="12"></td>
              </tr>
              <tr> 
                <td><table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
               
                    <tr> 
                      <div align="right"><jsp:include page="getLoginUser.jsp" flush="true"/></div> 
                    </tr>
                    <table width=600 border='1' align='center' cellpadding="1" cellspacing="0" bordercolor="#3A9D99" class="sbody" height="125">
  
  <tr bgcolor="#9AD3D0"> 
    <td colspan='2' class="sbody" align="center" height="17" width="594"><span style="font-size: 12.0pt; mso-bidi-font-size: 10.0pt; font-family: 標楷體; mso-bidi-font-family: Times New Roman; mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW; mso-bidi-language: AR-SA; mso-ascii-font-family: Times New Roman; mso-hansi-font-family: Times New Roman"><font color="#0000FF">(1)巳產生之資料檔資訊</font></span></td>
  </tr>


  <%if(objFile.exists()){%>     

  <tr> 
    <td width='100' bgcolor='#D8EFEE' align='left' height="28"><span style="font-size: 12.0pt; mso-bidi-font-size: 10.0pt; font-family: 標楷體; mso-bidi-font-family: Times New Roman; mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW; mso-bidi-language: AR-SA; mso-ascii-font-family: Times New Roman; mso-hansi-font-family: Times New Roman">路徑/檔名</span></td>
	<td width='490' bgcolor='e7e7e7' height="28"><%=Path%></td>			    
  </tr>
  <tr> 
    <td width='100' bgcolor='#D8EFEE' align='left' height="28"><span style="font-size: 12.0pt; mso-bidi-font-size: 10.0pt; font-family: 標楷體; mso-bidi-font-family: Times New Roman; mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW; mso-bidi-language: AR-SA; mso-ascii-font-family: Times New Roman; mso-hansi-font-family: Times New Roman">大小</span></td>
	<td width='490' bgcolor='e7e7e7' height="28"><%=objFile.length()%>位元組</td>			    
  </tr>
  <tr> 
    <td width='100' bgcolor='#D8EFEE' align='left' height="28"><span style="font-size: 12.0pt; mso-bidi-font-size: 10.0pt; font-family: 標楷體; mso-bidi-font-family: Times New Roman; mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW; mso-bidi-language: AR-SA; mso-ascii-font-family: Times New Roman; mso-hansi-font-family: Times New Roman">產生日期</span></td>
	<td width='490' bgcolor='e7e7e7' height="28"><%if(((DataObject)WTT07_ELM.get(0)).getValue("update_date") != null ) out.print((String)((DataObject)WTT07_ELM.get(0)).getValue("update_date")); else out.print("&nbsp;");%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      
                            <a href="javascript:doSubmit(this.document.forms[0],'Qry');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_queryb.gif',1)"><img src="images/bt_query.gif" name="Image101" width="66" height="25" border="0" id="Image107"></a>&nbsp;    
                            <a href="javascript:doSubmit(this.document.forms[0],'Clear');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_delete.gif',1)"><img src="images/bt_delete.gif"name="Image101" width="66" height="25" border="0" id="Image107"></a>&nbsp;   
    </td>			    
  </tr>
<%}else{%>
  <tr> 
    <td width='100' bgcolor='#D8EFEE' align='left' height="28"><span style="font-size: 12.0pt; mso-bidi-font-size: 10.0pt; font-family: 標楷體; mso-bidi-font-family: Times New Roman; mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW; mso-bidi-language: AR-SA; mso-ascii-font-family: Times New Roman; mso-hansi-font-family: Times New Roman">路徑/檔名</span></td>
	<td width='490' bgcolor='e7e7e7' height="28"></td>			    
  </tr>
  <tr> 
    <td width='100' bgcolor='#D8EFEE' align='left' height="28"><span style="font-size: 12.0pt; mso-bidi-font-size: 10.0pt; font-family: 標楷體; mso-bidi-font-family: Times New Roman; mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW; mso-bidi-language: AR-SA; mso-ascii-font-family: Times New Roman; mso-hansi-font-family: Times New Roman">大小</span></td>
	<td width='490' bgcolor='e7e7e7' height="28">&nbsp;&nbsp;</td>			    
  </tr>
  <tr> 
    <td width='100' bgcolor='#D8EFEE' align='left' height="28"><span style="font-size: 12.0pt; mso-bidi-font-size: 10.0pt; font-family: 標楷體; mso-bidi-font-family: Times New Roman; mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW; mso-bidi-language: AR-SA; mso-ascii-font-family: Times New Roman; mso-hansi-font-family: Times New Roman">產生日期</span></td>
	<td width='490' bgcolor='e7e7e7' height="28">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;     
    </td>
    		    
  </tr>
<%}%>   
</table>
<tr>
<td> 
<table width=605 border=1 align=center cellpadding="1" cellspacing="0" bordercolor="#3A9D99" height="84">                          						
        			    <tr class="sbody">
			<td width='589' bgcolor='#D8EFEE' align='left' height="27" colspan="5">
                        <span style="font-size: 12.0pt; mso-bidi-font-size: 10.0pt; font-family: 標楷體; mso-bidi-font-family: Times New Roman; mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW; mso-bidi-language: AR-SA; mso-ascii-font-family: Times New Roman; mso-hansi-font-family: Times New Roman">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                        <font color="#0000FF">(2)建立擷取者身份識別資訊</font></span>
                        </td>
        			    </tr>
						
        			   
						
        			    <tr class="sbody">
						<td width='53' bgcolor='#D8EFEE' align='left' height="27" colspan="2">
                        <span style="font-size: 12.0pt; mso-bidi-font-size: 10.0pt; font-family: 標楷體; mso-bidi-font-family: Times New Roman; mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW; mso-bidi-language: AR-SA; mso-ascii-font-family: Times New Roman; mso-hansi-font-family: Times New Roman">帳號</span>
                        </td>
						<td width='536' bgcolor='e7e7e7' height="27" colspan="3">
                        <input type='text' name='user_id' value="<%if(((DataObject)WTT07_ELM.get(0)).getValue("muser_id") != null ) out.print((String)((DataObject)WTT07_ELM.get(0)).getValue("muser_id")); else out.print("&nbsp;");%>" size='20' maxlength='12'>
						</td>
        			    </tr>
							
        			    <tr class="sbody">
						<td width='53' bgcolor='#D8EFEE' align='left' height="26" colspan="2">
                        <span style="font-size: 12.0pt; mso-bidi-font-size: 10.0pt; font-family: 標楷體; mso-bidi-font-family: Times New Roman; mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW; mso-bidi-language: AR-SA; mso-ascii-font-family: Times New Roman; mso-hansi-font-family: Times New Roman">密碼</span>
                        </td>
						<td width='536' bgcolor='e7e7e7' height="28" colspan="3">
                            <input type='text' name='user_pwd' value="<%if(((DataObject)WTT07_ELM.get(0)).getValue("muser_password") != null ) out.print((String)((DataObject)WTT07_ELM.get(0)).getValue("muser_password")); else out.print("&nbsp;");%>" size='20' maxlength='20'>
						</td>
        			    </tr>
						        			   						
        			    <tr class="sbody">
						<td width='53' bgcolor='#D8EFEE' align='left' height="24" colspan="2">
                        <span style="font-size: 12.0pt; mso-bidi-font-size: 10.0pt; font-family: 標楷體; mso-bidi-font-family: Times New Roman; mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW; mso-bidi-language: AR-SA; mso-ascii-font-family: Times New Roman; mso-hansi-font-family: Times New Roman">來源IP</span></td>
						<td width='536' bgcolor='e7e7e7' height="19" colspan="3">
                            <input type='text' name='ip_Address' value="<%if(((DataObject)WTT07_ELM.get(0)).getValue("ip_address") != null ) out.print((String)((DataObject)WTT07_ELM.get(0)).getValue("ip_address")); else out.print("&nbsp;");%>" size='20' maxlength='20'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" name="load" value="建立身份識別資料檔" onClick="doSubmit(this.document.forms[0],'New');">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;   
						
        			    </tr>
						
        			   						
        			    <tr class="sbody">
						<td width='583' bgcolor='#D8EFEE' align='left' height="30" colspan="5">
                          <span style="font-size: 12.0pt; mso-bidi-font-size: 10.0pt; font-family: 標楷體; mso-bidi-font-family: Times New Roman; mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW; mso-bidi-language: AR-SA; mso-ascii-font-family: Times New Roman; mso-hansi-font-family: Times New Roman"><font color="#0000FF">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                          (3)</font></span><font color="#0000FF"><span style="font-size: 12.0pt; mso-bidi-font-size: 10.0pt; font-family: 標楷體; mso-bidi-font-family: Times New Roman; mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW; mso-bidi-language: AR-SA; mso-ascii-font-family: Times New Roman; mso-hansi-font-family: Times New Roman">產生最新「農漁會理監事及負責人」檔案</span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="javascript:doSubmit(this.document.forms[0],'CreateHeadText');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_confirmb.gif',1)"><img src="images/bt_confirm.gif" name="Image101" width="66" height="25" border="0" id="Image108"></a>     
</font>
                        </td>
        			    </tr>
						
        			   
						
        			    
						
        			   
						
        			    <tr class="sbody">
						<td width='583' bgcolor='#D8EFEE' align='left' height="30" colspan="5"></span><span style="font-size: 12.0pt; mso-bidi-font-size: 10.0pt; font-family: 標楷體; mso-bidi-font-family: Times New Roman; mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW; mso-bidi-language: AR-SA; mso-ascii-font-family: Times New Roman; mso-hansi-font-family: Times New Roman">
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#0000FF">(4)</font></span><font color="#0000FF"><span style="font-size: 12.0pt; mso-bidi-font-size: 10.0pt; font-family: 標楷體; mso-bidi-font-family: Times New Roman; mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW; mso-bidi-language: AR-SA; mso-ascii-font-family: Times New Roman; mso-hansi-font-family: Times New Roman">歷程記錄檔查詢維護</span></font><a href="javascript:selectAll(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image104','','images/bt_selectallb.gif',1)"><img src="images/bt_selectall.gif" name="Image104" width="80" height="25" border="0" id="Image104"></a>                       
 						    						<a href="javascript:selectNo(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_selectnob.gif',1)"><img src="images/bt_selectno.gif" name="Image105" width="80" height="25" border="0" id="Image105"></a>                        							 						     						                          </span><a href="javascript:doSubmit(this.document.forms[0],'Delete');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_confirmb.gif',1)"><img src="images/bt_confirm.gif" name="Image101" width="66" height="25" border="0" id="Image109"></a><span style="font-size: 12.0pt; mso-bidi-font-size: 10.0pt; font-family: 標楷體; mso-bidi-font-family: Times New Roman; mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW; mso-bidi-language: AR-SA; mso-ascii-font-family: Times New Roman; mso-hansi-font-family: Times New Roman"> 
                                                </td>
        			    </tr>
						
        			   
						
        			    <tr class="sbody">
						
						<td width='50' bgcolor='e7e7e7' height="23">
                        刪除
						</td>
                        			<td width='220' bgcolor='e7e7e7' height="23">
                        登入日期時間
						</td> 
						<td width='110' bgcolor='e7e7e7' height="23">
                        登入者帳號/密碼
						</td>
						<td width='50' bgcolor='e7e7e7' height="23">
                        登入者IP
						</td>
						<td width='150' bgcolor='e7e7e7' height="23">
                        擷取資料結果狀態
						</td>
        			    </tr>
<% List WTT07 = DBManager.QueryDB_SQLParam("select to_char(input_date,'mm/dd/yyyy hh24:mi:ss') as input_date,muser_id,muser_password,ip_address,result_p from WTT07 ",null,""); %>
<% if(WTT07.size() == 0){%>
                        <tr>                  			   
                   	<td colspan=10 align=center>無資料可供查詢</td>		   
                   	</tr>
			<% } 
                        int i=0;
                    	while(i<WTT07.size()){                                             		                      		                          		      
                      %>                         	  
                        <tr >
                          <td width=50><input type="checkbox" name="isDelete" value="<%=((DataObject)WTT07.get(i)).getValue("input_date")%>"></td>    				                      				                          
                          <td width=220><div align="center"><font face=細明體 color=#000000><%if(((DataObject)WTT07.get(i)).getValue("input_date") != null ) out.print((String)((DataObject)WTT07.get(i)).getValue("input_date")); else out.print("&nbsp;");%></font></div></td>              
                          <td width=110><div align="center"><font face=細明體 color=#000000><%if(((DataObject)WTT07.get(i)).getValue("muser_id") != null ) out.print((String)((DataObject)WTT07.get(i)).getValue("muser_id")); else out.print("&nbsp;");%>/<%if( ((DataObject)WTT07.get(i)).getValue("muser_password") != null ) out.print((String)((DataObject)WTT07.get(i)).getValue("muser_password")); else out.print("&nbsp;");%></font></div></td>              
                          <td width=50><div align="center"><font face=細明體 color=#000000><%if(((DataObject)WTT07.get(i)).getValue("ip_address") != null ) out.print((String)((DataObject)WTT07.get(i)).getValue("ip_address")); else out.print("&nbsp;");%></font></div></td>
                          <td width=150><div align="center"><font face=細明體 color=#000000><%if(((DataObject)WTT07.get(i)).getValue("result_p") != null ) out.print((String)((DataObject)WTT07.get(i)).getValue("result_p")); else out.print("&nbsp;");%></font></div></td>
                        </tr> 					      
		      <%
                  			   i++;
	                  		   }//end of while                     
                      %>
        			   
						
        			    
						
        			   
						
                        </Table></td>
                    </tr>                 
                    <tr>                  
                <td><div align="right">







<table width=600 border='1' align='center' cellpadding="1" cellspacing="1" bordercolor="#3A9D99" class="sbody" height="96">
  
  <tr bgcolor="#9AD3D0"> 
    <td colspan='2' class="sbody" align="center" height="17"><font color='#000000'>維護者資訊</font></td>
  </tr>
  <tr> 
    <td width='15%' bgcolor='#D8EFEE' align='left' height="17">姓名</td>
	<td width='85%' bgcolor='e7e7e7' height="17">&nbsp;測試使用者</td>			    
  </tr>
  <tr> 
    <td width='15%' bgcolor='#D8EFEE' align='left' height="21">電話</td>
	<td width='85%' bgcolor='e7e7e7' height="21">&nbsp;02-26551188</td>			    
  </tr>
  <tr> 
    <td width='15%' bgcolor='#D8EFEE' align='left' height="17">E-MAIL</td>
	<td width='85%' bgcolor='e7e7e7' height="17">&nbsp;yanjaneshi@yahoo.com.tw</td>			    
  </tr>
</table> 

                    </div></td>                                              
              </tr>
              
              <tr> 
                <td>&nbsp;</td>
              </tr>
              <tr> 
                <td><div align="center"> 
                    <table width="243" border="0" cellpadding="1" cellspacing="1">
                      
                    </table>
                  </div></td>
              </tr>   
              
  <tr> 
                <td><table width="600" border="0" cellpadding="1" cellspacing="1" class="sbody">
                    <tr> 
                      <td colspan="2"><font color='#990000'><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle"><font color="#007D7D" size="3">使用說明         
                        : </font></font></td>
                    </tr>
                    <tr> 
                      <td width="16">&nbsp;</td>
                      <td width="577"> <ul>
                      	  <li><font color="#0000FF">產生新檔案:</font>確認輸入之異動期間無誤後,    
                            按【確定】即產生該期間有異動之「農漁會理監事及負責人」檔案,供擷取至農金局全球資訊網發佈。</li> 
                      	  <li>按<font color="#666666">【查詢】</font>即可查閱巳產生之資料檔內容。</li>                         
                      	  <li>按<font color="#666666">【刪除】</font>即會刪除巳產生之資料檔。</li>                         
                          <li>欲重新輸入資料, 按<font color="#666666">【取消】</font>即將本表上的資料清空。</li>    
                        </ul></td>
                    </tr>
                  </table></td>
              </tr>
              <!--tr> 
                <td><div align="center"><img src="images/line_1.gif" width="600" height="12"></div></td>
              </tr-->
</table>
</form>
</body>
