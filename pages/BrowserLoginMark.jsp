<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.LinkedList" %>

<%!
int FindCookie(Cookie cookies[],String cookie_name)
{
	for(int i=0;i<cookies.length;i++)
	{
		if(cookies[i].getName().equals(cookie_name))
		{
			return i;
		}
	}
	
	return -1;
}
%>

<% 			
  String sqlCmd="";
	List updateDBSqlList = new LinkedList();								   				   
 	
 	Cookie cookies[] = request.getCookies();
 	int index = FindCookie(cookies,"cmuser_id");  
	String muser_id = cookies[index].getValue();
  
	  
  if(!muser_id.equals(""))
  {
         sqlCmd = "UPDATE WTT01 SET "
		        + " login_mark='N'"
		        + ",update_date=sysdate" 		            		 						       
		        + " where muser_id='"+muser_id+"'";				    	   
					   
		        updateDBSqlList.add(sqlCmd); 		            		            	            	            		
		     		DBManager.updateDB(updateDBSqlList);					 						 
			      System.out.println("Session is exiting!!");	
			      session.invalidate();
	} 
	else
	{
			 		System.out.println("NO DATA!!");	
	}	
	
	 cookies[index].setMaxAge(0);
	 response.addCookie(cookies[index]);
	
%>