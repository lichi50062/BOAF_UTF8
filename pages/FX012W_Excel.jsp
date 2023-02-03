<%
//97.09.26 add 簽証會計師歷年明細表 by 2295 	
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
   String bank_no = ( request.getParameter("bank_no")==null ) ? "" : (String)request.getParameter("bank_no");   
   
   System.out.println("m_year="+m_year);
   System.out.println("bank_no="+bank_no);
   
   response.setHeader("Content-Disposition","attachment; filename=download.xls");   

%>
<%
	try{	    
	    String actMsg = RptFR052W.createRpt("2",m_year,"",bank_no);
	    System.out.println("createRpt="+actMsg);
	    System.out.println("filename="+Utility.getProperties("reportDir")+System.getProperty("file.separator")+"信用部歷年簽証會計師明細表.xls");
		FileInputStream fin = new FileInputStream(Utility.getProperties("reportDir")+System.getProperty("file.separator")+"信用部歷年簽証會計師明細表.xls");
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