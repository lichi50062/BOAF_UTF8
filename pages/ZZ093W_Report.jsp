<%
//102.07.01 created 基本資料歷程記錄查詢列印 by 2968 
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.util.*,java.io.*" %>
<%@ page import="org.apache.poi.poifs.filesystem.*,org.apache.poi.hssf.usermodel.*" %>
<%@ page import="com.tradevan.util.Utility" %>								          
<%@ page import="com.tradevan.util.report.RptZZ093W" %>								          
<%
   response.setContentType("application/msexcel;charset=UTF-8");//以上這行設定本網頁為excel格式的網頁	     
   String act = ( request.getParameter("act")==null ) ? "view" : (String)request.getParameter("act");
   String excelaction = ( request.getParameter("excelaction")==null ) ? "view" : (String)request.getParameter("excelaction");
   List body =  (List) request.getAttribute("ReportBody");
   String begDate = ( request.getParameter("begDate")==null ) ? "" : (String)request.getParameter("begDate");
   String endDate = ( request.getParameter("endDate")==null ) ? "" : (String)request.getParameter("endDate");
   int begY = Integer.parseInt(begDate.substring(0,4))-1911;
   String begM = begDate.substring(4,6);
   String begD = begDate.substring(6,8);
   begDate = begY+"/"+begM+"/"+begD;
   int endY = Integer.parseInt(endDate.substring(0,4))-1911;
   String endM = endDate.substring(4,6);
   String endD = endDate.substring(6,8);
   endDate = endY+"/"+endM+"/"+endD;
   String duringDate = begDate+"~"+endDate;
   System.out.println("ZZ093W_Excel.jsp ...Start");
   System.out.println("act="+act);
   System.out.println("body.size = " + body.size());
   System.out.println("excelaction="+excelaction);

   if(excelaction.equals("view")){
      //以上這行設定傳送到前端瀏覽器時的檔名為test1.xls
      //就是靠這一行，讓前端瀏覽器以為接收到一個excel檔 
      response.setHeader("Content-disposition","inline; filename=view.xls");
   }else if (excelaction.equals("download")){   
      response.setHeader("Content-Disposition","attachment; filename=download.xls");
   }

%>
<%
	try{
	    
	  String fileName = new RptZZ093W().createReport( body, duringDate);
	  System.out.println("ZZ093W檔案路徑位置 ="+Utility.getProperties("reportDir")+System.getProperty("file.separator")+fileName);
		FileInputStream fin = new FileInputStream(Utility.getProperties("reportDir")+System.getProperty("file.separator")+fileName);  		 
		ServletOutputStream out1 = response.getOutputStream();           
		byte[] line = new byte[8196];
		int getBytes=0;
		while( ((getBytes=fin.read(line,0,8196)))!=-1 ){
			System.out.println("getBytes ="+getBytes);		  	  	
			out1.write(line,0,getBytes);
			out1.flush();
	    }		
		fin.close();
		out1.close();            		      
		
	}catch(Exception e){
	   System.out.println(e.getMessage());
	}		       
%>	    		