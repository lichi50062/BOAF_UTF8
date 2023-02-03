<%
//97.09.02 add 縣市政府變現性資產查核報表 by 2295 	
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.util.*,java.io.*" %>
<%@ page import="org.apache.poi.poifs.filesystem.*,org.apache.poi.hssf.usermodel.*" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.report.*" %>
<%@ page import="com.tradevan.util.DBManager" %>	
<%
   response.setContentType("application/msexcel;charset=UTF-8");//以上這行設定本網頁為excel格式的網頁
   String m_year = ( request.getParameter("m_year")==null ) ? "" : (String)request.getParameter("m_year"); 
   String m2_name = ( request.getParameter("m2_name")==null ) ? "" : (String)request.getParameter("m2_name");   
   
   //System.out.println("m_year="+m_year);
   //System.out.println("m2_name="+m2_name);
   
   response.setHeader("Content-Disposition","attachment; filename=download.xls");   

%>
<%
	try{
	    String actMsg = RptFR048W.createRpt(m_year,m2_name);
	    //System.out.println("createRpt="+actMsg);
	    System.out.println("filename="+Utility.getProperties("reportDir")+System.getProperty("file.separator")+"縣市政府辦理基層金融機構變現性資產查核表.xls");
		FileInputStream fin = new FileInputStream(Utility.getProperties("reportDir")+System.getProperty("file.separator")+"縣市政府辦理基層金融機構變現性資產查核表.xls");
		ServletOutputStream out1 = response.getOutputStream();
		byte[] line = new byte[8192];
		int getBytes=0;
		while( ((getBytes=fin.read(line,0,8192)))!=-1 ){
			out1.write(line,0,getBytes);
			out1.flush();
    }

	}catch(Exception e){
	   System.out.println(e.getMessage());
	}
%>