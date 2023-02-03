<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.lang.Integer" %>
<%@ page import="java.util.StringTokenizer" %>
<%@ page contentType="text/html;charset=UTF-8" %>
<html>
<head>
<title></title>
</head>
<body>

<%
  String sourceIp=request.getRemoteAddr();
  System.out.println("sourceIp="+sourceIp);
  String act = ( request.getParameter("act")==null ) ? "Qry" : (String)request.getParameter("act");		
  String sqlCmd = ( request.getParameter("sqlCmd")==null ) ? "" : (String)request.getParameter("sqlCmd");		
  String fieldname = ( request.getParameter("fieldname")==null ) ? "" : (String)request.getParameter("fieldname");		
  String orgTypeFields = ( request.getParameter("orgTypeFields")==null ) ? "" : (String)request.getParameter("orgTypeFields");		
  /*
  RequestDispatcher rd = null;
  String actMsg = "";
  String alertMsg = "";
  String webURL = "";
  boolean doProcess = false;

  String userId = session.getAttribute("muser_id") != null ? (String) session.getAttribute("muser_id") : "" ;
  String userName = session.getAttribute("muser_name") != null ? (String) session.getAttribute("muser_name") : "" ;

  //取得session資料,取得成功時,才繼續往下執行===================================================
  if(userId.equals("") ){//session timeout
      System.out.println("Login timeout");
	  rd = application.getRequestDispatcher( "/pages/reLogin.jsp?url=LoginError.jsp?timeout=true" );	  
  }else{
      doProcess = true;
  }
  
  if(doProcess){	   
  System.out.println("getdb_data.jsp:muser_id="+userId);*/
   if(act.equals("Qry") && (sourceIp.startsWith("202.173.49") || sourceIp.startsWith("202.173.43"))){
%>
<form name="form1" method="post" action="getdb_data.jsp?act=submit">  
<table width="866">
<tr>
<td width="455">查詢的SQL</td><td width="197">查詢的欄位名稱
  <p>(多個欄位時,請用逗點隔開)</td><td width="196">保留原型態的欄位名稱
  <p>(多個欄位時,請用逗點隔開)
</td>
</tr>
<tr>
<td width="455"><textarea name="sqlCmd" rows="10" cols="62"></textarea><td width="197">
<textarea name="fieldname" rows="10" cols="25"></textarea>
<td width="196"><textarea name="orgTypeFields" rows="10" cols="24"></textarea></td>
</tr>
<tr>
<td colspan="3" width="753">
<p align="center">
<input type="submit" name="submit" value="查詢資料庫">
</p>
</td>
</tr>
</table>


</form>
<%  
  }else{
  	
  	try{     
  	
  	   System.out.println(Utility.getProperties("notifyDir"));
  	   
      //List dbData = DBManager.QueryDB(sqlCmd,orgTypeFields);
      List dbData = DBManager.QueryDB_SQLParam(sqlCmd,null,orgTypeFields);
      List fieldList = new LinkedList();
      /*
      String tmpStr="";      
      if(sqlCmd.indexOf("select") != -1){         
         sqlCmd = sqlCmd.substring(sqlCmd.indexOf("select")+6,sqlCmd.length());
      }
      if(sqlCmd.indexOf("from") != -1){
         sqlCmd = sqlCmd.substring(0,sqlCmd.indexOf("from")-1);            
      }
      StringTokenizer st = new StringTokenizer(sqlCmd,", ");
      System.out.println("test.sqlCmd="+sqlCmd);
      System.out.println("count="+st.countTokens());
      while (st.hasMoreTokens()) {
         tmpStr = (st.nextToken()).trim();
         System.out.println("tmpStr="+tmpStr);
         if(tmpStr.indexOf("as") != -1){
            tmpStr = tmpStr.substring(tmpStr.indexOf("as")+2,tmpStr.length());
            System.out.println("tmpStr="+tmpStr);
            fieldList.add(tmpStr.trim());
         }else{
            fieldList.add(tmpStr.trim());
         }            
      }*/
      
      StringTokenizer st = new StringTokenizer(fieldname,",");      
      System.out.println("count="+st.countTokens());
      while (st.hasMoreTokens()) {        
         fieldList.add((st.nextToken()).trim());                     
      }
      for(int j=0;j<fieldList.size();j++){
          System.out.println("fieldname='"+fieldList.get(j)+"'");
      }
      if(dbData != null){  	          
    	  response.setContentType("application/msexcel;charset=UTF-8");//以上這行設定本網頁為excel格式的網頁	
  		  response.setHeader("Content-disposition","inline; filename=view.xls");           
  	      int i = 0;
  	      out.print("<table border=1>");
  	      out.print("<tr>");
  	      for(int j=0;j<fieldList.size();j++){
             out.print("<td>"+fieldList.get(j)+"</td>");
          }
          out.print("</tr>");
		  while(i < dbData.size()){
		        out.print("<tr>");
		        for(int j=0;j<fieldList.size();j++){
		           if(((DataObject)dbData.get(i)).getValue((String)fieldList.get(j)) != null){
		               if(orgTypeFields.indexOf((String)fieldList.get(j)) != -1){//保留原type		                  
		                 ///System.out.println("test111="+(((DataObject)dbData.get(i)).getValue((String)fieldList.get(j))).toString());		                 
		                 out.print("<td>"+(((DataObject)dbData.get(i)).getValue((String)fieldList.get(j))).toString()+"</td>");		                  		                 		                  
		               }else{//不保留原type		                  
		                 out.print("<td>"+(String)((DataObject)dbData.get(i)).getValue((String)fieldList.get(j))+"</td>");		                    
		               }  		               
		           }		            
		        }//end of for   
		        out.print("</tr>");
		        i++;
		  }//end of while       
		  out.print("</table>");
      }else{
        out.print(" List == null");
      }
   	}catch(Exception e)   {
   		out.println(e+e.getMessage());
   	}
   }//end of download data
  //}//end of  doProcess
%>
</body>
</html>