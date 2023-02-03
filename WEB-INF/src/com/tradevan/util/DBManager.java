/*
 * Created on 2004/10/30
 * 99.02.03 add 查詢使用preparestatment;QueryDB_SQLParam(szSQL,paramList,orgTypeFields)
 *          add 更新時;使用PreparedStatement;updateDB_ps(szSQL)
 */
package com.tradevan.util;



import java.util.LinkedList;
import java.util.List;
import java.sql.*;

//import javax.activation.DataSource;
//import javax.naming.Context;
//import com.iplanet.ias.server.ApplicationServer;
import com.tradevan.util.dao.RdbCommonDao;
import com.tradevan.util.dao.DAOFactory;


public class DBManager {
	private static String errMsg="";
	private static RdbCommonDao RdbCommonDao=null;
    public static String getErrMsg(){
    	return errMsg;
    }
    //原update
	public static boolean updateDB(List szSQL){          
		 try {
		   	   errMsg = "";
		 	   RdbCommonDao RdbCommonDao = DAOFactory.getRdbCommonDao("");          
			   if(RdbCommonDao.updateDB(szSQL)){				   	
			      return true;	          
			   }else{
			   	  errMsg = RdbCommonDao.getErrMsg();
			   	  return false;
			   }	
		 }catch(Exception e) {             	
			e.printStackTrace();
			return false;
		 }         
	}
	/*
	 * 使用PreparedStatement
	 * @szSQL 0:要執行的PreparedStatement 
	 *        1:參數List;若參數為null則不加入參數	 
	 */   
	public static boolean updateDB_ps(List szSQL){          
		 try {
		   	   errMsg = "";
		   	   RdbCommonDao RdbCommonDao = DAOFactory.getRdbCommonDao("");
		 	   if(RdbCommonDao.updateDB_ps(szSQL)){				   	
			      return true;	          
			   }else{
			   	  errMsg = RdbCommonDao.getErrMsg();
			   	  return false;
			   }	
		 }catch(Exception e) {             	
			e.printStackTrace();
			return false;
		 }         
	}
	
	public static boolean updateDB(List szSQL,String pgName){          
		 try {
		   	   errMsg = "";
		 	   RdbCommonDao RdbCommonDao = DAOFactory.getRdbCommonDao("");          
			   if(RdbCommonDao.updateDB(szSQL,pgName)){				   	
			      return true;	          
			   }else{
			   	  errMsg = RdbCommonDao.getErrMsg();
			   	  return false;
			   }	
		 }catch(Exception e) {             	
			e.printStackTrace();
			return false;
		 }         
	}
	
	//96.05.01 add QueryDB,使用preparestatment =========================
	public static List QueryDB_ps(List szSQL,String orgTypeFields){          
		 try {
		   	   errMsg = "";
		 	   RdbCommonDao RdbCommonDao = DAOFactory.getRdbCommonDao("");    
		 	   
		 	   List queryList = RdbCommonDao.QueryDB_ps(szSQL, null, orgTypeFields,null);
			   	
			   if(queryList != null){
			   	  int i = 0;
			      return queryList;	          
			   }else{
			   	  errMsg = RdbCommonDao.getErrMsg();
			   	  System.out.println("errMsg="+errMsg);
			   	  return null;
			   }	
			   
		 }catch(Exception e) {             	
			e.printStackTrace();
			return null;
		 }         
	}
	//原查詢功能
	public static List QueryDB(String szSQL,String orgTypeFields){          
		 try {
		 	   RdbCommonDao RdbCommonDao = DAOFactory.getRdbCommonDao("");          
			   List queryList = RdbCommonDao.QueryDB(szSQL, null, orgTypeFields,null);
			   	
			   if(queryList != null){
			   	  int i = 0;
			   	  
			      return queryList;	          
			   }else{
			   	  errMsg = RdbCommonDao.getErrMsg();
			   	  System.out.println("errMsg="+errMsg);
			   	  return null;
			   }	
			   
		 }catch(Exception e) {             	
			e.printStackTrace();
			return null;
		 }         
	}
    
	//查詢使用preparestatment
	public static List QueryDB_SQLParam(String szSQL,List paramList,String orgTypeFields){          
		 try {
		      
		 	   RdbCommonDao RdbCommonDao = DAOFactory.getRdbCommonDao("");   
		      
			   List queryList = RdbCommonDao.QueryDB_SQLParam(szSQL, null, orgTypeFields,null,paramList);
			   	
			   if(queryList != null){
			   	  int i = 0;			   	  
			      return queryList;	          
			   }else{
			   	  errMsg = RdbCommonDao.getErrMsg();
			   	  System.out.println("errMsg="+errMsg);
			   	  return null;
			   }	
			   
		 }catch(Exception e) {   
		    System.out.println("QueryDB_SQLParam Error"); 
			e.printStackTrace();
			return null;
		 }         
	}
   
	//查詢使用preparestatment
	public static List QueryDB_SQLParam_new(String szSQL,List paramList,String orgTypeFields){          
			 try {
			      
			 	   RdbCommonDao RdbCommonDao = DAOFactory.getRdbCommonDao("");   
			      
				   List queryList = RdbCommonDao.QueryDB_SQLParam_new(szSQL, null, orgTypeFields,null,paramList);
				   	
				   if(queryList != null){
				   	  int i = 0;			   	  
				      return queryList;	          
				   }else{
				   	  errMsg = RdbCommonDao.getErrMsg();
				   	  System.out.println("errMsg="+errMsg);
				   	  return null;
				   }	
				   
			 }catch(Exception e) {   
			    System.out.println("QueryDB_SQLParam_new Error"); 
				e.printStackTrace();
				return null;
			 }         
	}
	public static List QueryDB_DBPool(String szSQL,String orgTypeFields,String dbPool){          
		 try {
		 	   RdbCommonDao RdbCommonDao = DAOFactory.getRdbCommonDao(dbPool);          
			   List queryList = RdbCommonDao.QueryDB(szSQL, null, orgTypeFields,null);
			   	
			   if(queryList != null){
			   	  int i = 0;
			   	  
			      return queryList;	          
			   }else{
			   	  errMsg = RdbCommonDao.getErrMsg();
			   	  System.out.println("errMsg="+errMsg);
			   	  return null;
			   }	
			 
		 }catch(Exception e) {             	
			e.printStackTrace();
			return null;
		 }         
	}
	public static List QueryDB(String szSQL,String orgTypeFields,String pgName) {          
		 try {		 	  
		   	   RdbCommonDao RdbCommonDao = DAOFactory.getRdbCommonDao("");          
			   List queryList = RdbCommonDao.QueryDB(szSQL, null, orgTypeFields,null,pgName);
			   	
			   if(queryList != null){
			   	  int i = 0;			   	  
			      return queryList;	          
			   }else{
			   	  errMsg = RdbCommonDao.getErrMsg();
			   	  System.out.println("errMsg="+errMsg);
			   	  return null;
			   }	
			   
		 }catch(Exception e) {  
		 	errMsg = "error";
			e.printStackTrace();
			return null;
		 }         
	}
    

	public static Connection newQryConnection(){
		try{
			//RdbCommonDao = DAOFactory.getRdbCommonDao(""); 
			Connection conn = RdbCommonDao.newConnection();
			return conn;
		}catch(Exception e){
			System.out.println("newQryConnection Error:"+e.getMessage());
			return null;
		}
	}
	
	public static void closeQryConnection(){
		try{
			//RdbCommonDao RdbCommonDao = DAOFactory.getRdbCommonDao(""); 
			RdbCommonDao.closeQryConnection();
		}catch(Exception e){
			System.out.println("closeQryConnection Error:"+e+e.getMessage());
		}
	}
}
