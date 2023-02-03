<%
//101.08.09 add 農信保原始檔下載 by 2295							
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.util.*,java.io.*" %>
<%@ page import="com.tradevan.util.Utility" %>
<%
   response.setContentType("application/msexcel;charset=UTF-8");//以上這行設定本網頁為excel格式的網頁
   String Report_no = ( request.getParameter("Report_no")==null ) ? "" : (String)request.getParameter("Report_no");
   String filename = ( request.getParameter("filename")==null ) ? "" : (String)request.getParameter("filename");
  
   System.out.println("Report_no="+Report_no);
   System.out.println("filename="+filename);   
   if(filename.toLowerCase().endsWith(new String("csv"))){ 
    response.setHeader("Content-Disposition","attachment; filename="+filename);
%>
<%
	try{	
		
	    System.out.println("filename="+Utility.getProperties("ClientRptDir")+System.getProperty("file.separator")+filename);
		FileInputStream fin = new FileInputStream(Utility.getProperties("ClientRptDir")+System.getProperty("file.separator")+filename);  		 
		ServletOutputStream out1 = response.getOutputStream();           
		byte[] line = new byte[8196];
		int getBytes=0;
		while( ((getBytes=fin.read(line,0,8196)))!=-1 ){		    		
			out1.write(line,0,getBytes);
			out1.flush();
	    }
		
		fin.close();
		out1.close();    
		java.io.File tmpFile = new java.io.File(Utility.getProperties("ClientRptDir")+System.getProperty("file.separator")+filename);
        System.out.println("tmpFile.exists()??"+tmpFile.exists());
		if(tmpFile.exists()) tmpFile.delete();        		      
		
	}catch(Exception e){
	   System.out.println(e.getMessage());
	}		  
 }else{
 	RequestDispatcher rd = null;
    String actMsg = "can not download file";
    request.setAttribute("actMsg",actMsg);
    rd = application.getRequestDispatcher( nextPgName );     
    try{
      rd.forward(request,response);
    }catch(Exception e){
      System.out.println("forward Error:"+e+e.getMessage());
    }
  }    
%>
<%!
    private final static String nextPgName = "/pages/ActMsg.jsp";
%>    