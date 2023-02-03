// 95.03.14 close Context by 2295
// 98.05.04 add 使用cdao取得connection
//102.04.11 add 使用sun one connection pool by 2295
//102.10.02 add closeStatement增加PreparedStatement by 2295
//102.10.03 add closeStatement增加ResultSet by 2295
//102.11.05 fix executeQuery(String sql,String qryCount)-取消使用 by 2295
package com.tradevan.util.dao;


import java.sql.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

import javax.sql.DataSource;//102.04.09 add
//import javax.naming.InitialContext;

import javax.naming.Context;
import javax.naming.InitialContext;
import javax.naming.NamingException;
//import javax.transaction.UserTransaction;
import com.iplanet.ias.server.ApplicationServer;
import com.iplanet.ias.server.ServerContext;
import com.tradevan.util.sql.SQLParameter;
import com.tradevan.util.sql.SQLUtils;

//import com.sun.webserver.init.ServerContextImpl;
//import com.sun.webserver.init.ServerContext;
import javax.naming.*;
public class DaoUtil {

    //private InitialContext ctx;
    private Context ctx;
    public Connection connection;
    public Statement stat;
    public PreparedStatement pstmt;//102.10.03 add
    public ResultSet rs;//102.10.03 add
    private String jndi;
    public List paraList = new ArrayList();
    
    public DaoUtil(String jndi) throws NamingException {
        this.jndi = jndi;       
        
        try{
            //Context ctx = new InitialContext();
            ServerContext sc = ApplicationServer.getServerContext(); //106.11.02原本的取得方式
            //ServerContextImpl sci = new ServerContextImpl();//106.11.09
            //ctx = sci.getNamingContext();//106.11.09
            System.out.println("DaoUtil.ServerContext="+sc); //106.11.06 原本的取得方式
        //ctx=ApplicationServer.getServerContext().getNamingContext();//103.11.11
            ctx=sc.getNamingContext();//106.11.06 原本的取得方式
            System.out.println("ctx="+sc.getNamingContext());//106.11.06 原本的取得方式
            System.out.println("ctx.getNameInNamespace()="+ctx.getNameInNamespace());
                       
        }catch(Exception e){System.out.println("DaoUtil Error:"+e+e.getMessage());}    
    }
//106.11.06 end ======================================================================    
    /*
    public UserTransaction getTransaction() throws NamingException {
        // do NOT use 'java:comp/UserTransaction', this will cause exception
        // when invoked from timer service on weblogic server 6.1
        return (UserTransaction)ctx.lookup("javax.transaction.UserTransaction");
 

    }*/	
    public void openConnection() throws NamingException, SQLException {    	
        //DataSource ds = (DataSource)ctx.lookup("java:comp/env/jdbc/"+jndi);
        /*102.04.09 原程式.公司xdao
    	if(connection == null){
           connection = Factory.getConnection();
    	}
    	*/
    	//98.05.04//102.04.11 sun one connection pool
       
    	try{
    	System.out.println("jndi="+jndi);
    	System.out.println("open connection begin--");
    	
    	//DataSource ds =null;//108.10.28
    	DataSource ds = null;
    	try{    	   
    	    ds = (DataSource)ctx.lookup(jndi);//106.11.06 原本的   		
    	    System.out.println("ctx.lookup");
    	   /*106.11.06 add begin===============================================
    	    
    	    Properties properties = new Properties();
            //properties.put("java.naming.factory.initial" , "com.sun.jndi.fscontext.FSContextFactory");
            Context initContext = new InitialContext();
            System.out.println("test1");
            Context webContext = (Context)initContext.lookup("java:/comp/env");

            ds = (DataSource) webContext.lookup("jdbc/myJdbc");
             
            System.out.println("test2");
           // List the objects 
            //String target = "";
            //NamingEnumeration namingEnum = initContext.list(target);
            //while (namingEnum.hasMore()) {
            //     System.out.println("test1="+namingEnum.next());
            //}
            //ctx.close();
            
            //System.out.println("test1="+initContext.getNameInNamespace());
            //Context webContext = (Context)initContext.lookup("java:/comp/env/jdbc");
            //System.out.println("test2="+webContext.getNameInNamespace());
            //System.out.println("test1");
            //ds = (DataSource)ctx.lookup("jdbc/PBOAFPool");//106.11.06 原本的
            //System.out.println("test2");
            //System.out.println("test2="+webContext.getNameInNamespace());
            //ds = (DataSource) webContext.lookup("jdbc/myJdbc");
            //System.out.println("test3");
            //ds = (DataSource)webContext.lookup("jdbc/myJdbc");
            //System.out.println("test4");
            //ds = (DataSource)ctx.lookup("PBOAFPool");             
             */
    	}catch(Exception e1){
            e1.printStackTrace();
            System.out.println("(DataSource)ctx.lookup(jndi) Error:"+e1+e1.getMessage());
        }
    	
        connection = ds.getConnection();//108.10.28       
        System.out.println("open connection end");        
        
    	}catch(Exception e){
    		e.printStackTrace();
    		System.out.println("openConnection Error:"+e+e.getMessage());
    	}
    	
    }
    public void openConnection(String pgName) throws NamingException, SQLException {    	
        //DataSource ds = (DataSource)ctx.lookup("java:comp/env/jdbc/"+jndi);
        /*102.04.09 原程式.公司xdao
    	 if(connection == null){
            connection = Factory.getConnection();
    	 }  
    	*/ 
    	//98.05.04//102.04.11 sun one connection pool
    	System.out.println("jndi="+jndi);
    	System.out.print(pgName+":open connection begin--");
    	DataSource ds = (DataSource)ctx.lookup(jndi);
        connection = ds.getConnection();
        System.out.println(connection+":open connection end");
        
    }
    public Connection newConnection() throws NamingException, SQLException {    	
        //DataSource ds = (DataSource)ctx.lookup("java:comp/env/jdbc/"+jndi);
        /*102.04.09 原程式.公司xdao
    	return Factory.getConnection();
    	*/
    	//98.05.04//102.04.11 sun one connection pool
    	System.out.println("jndi="+jndi);
    	System.out.println("new connection begin--");
    	DataSource ds = (DataSource)ctx.lookup(jndi);
        Connection newConnection = ds.getConnection();
        System.out.println("new connection end");
        return newConnection;
        
    }
    public void closeConnection() throws SQLException {   
        if(connection != null && !connection.isClosed()){//104.10.06
        //if (connection != null) {        	
        	System.out.println("closeConnection begin--");
            connection.close();//108.10.30 fix
            //connection = null;//108.10.30 fix //109.06.04 fix
            //95.03.14 close Context//102.04.11 sun one connection pool
            /*108.10.30 fix
            try {
                ctx.close();
                ctx = null;
            } catch (NamingException e) {                
                e.printStackTrace();
            }
            */
            System.out.println("closeConnection end");
        }
    }
    public void closeConnection(String pgName) throws SQLException {    	
        //if (connection != null) {
        if(connection != null && !connection.isClosed()){//104.10.06    
        	System.out.print(pgName+":closeConnection begin--"+connection+":");
            connection.close(); //108.10.30 fix
            connection = null;//108.10.30 fix
            //95.03.14 close Context//102.04.11 sun one connection pool      
            /*108.10.30 fix
            try {
                ctx.close();
                ctx = null;
            } catch (NamingException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            } 
            */           
            System.out.println("closeConnection end");
        }
    }
    public void setConnection(Connection newConnection){    	
           this.connection = newConnection;
           System.out.println("set to this.connection="+this.connection);
    }
    public PreparedStatement getPreparedStatement(String sql) throws SQLException {
        System.out.println("getPreparedStatement begin--");
        //102.10.02PreparedStatement ps = connection.prepareStatement(sql);
        if(pstmt != null){
            pstmt.close();
            //pstmt = null;//104.11.03 fix //108.10.30 fix   
        }
        pstmt = connection.prepareStatement(sql);
        System.out.println("getPreparedStatement end");
        return pstmt;
    }
    
    public Statement getStatement() throws SQLException {
    	System.out.print("getStatement begin--");
        stat = connection.createStatement();
        System.out.println("getStatement end");
        return stat;
    }
    public void closeStatement() throws SQLException {
        System.out.println("closeStatement(rs/stat/pstmt) begin--");
        if(rs != null){
            rs.close();
            //rs = null;//102.10.03 //108.10.30 fix
        }
        
    	if(stat != null){
    	    stat.close();
    	    //stat = null;//102.04.09 //108.10.30 fix
    	}
    	
    	if(pstmt != null){
    	    pstmt.close();
    	    //pstmt = null;//102.10.02 //108.10.30 fix
    	}
    	
    	System.out.println("closeStatement(rs/stat/pstmt) end");
    }
    
    public ResultSet executeQuery(String sql,String qryCount) throws SQLException {
        /*
    	if(connection == null){
    	   System.out.println("connection is null");
    	}
        stat = connection.createStatement();
        
        if(qryCount != null && !qryCount.equals("")){  
           System.out.println("qryCount.exe="+qryCount);      
		   stat.setMaxRows(Integer.parseInt(qryCount));
        }
        
        rs = stat.executeQuery(sql);//102.10.03 fix 
        return rs;
        */
        return null;
    }
    
    public ResultSet executeQuery_SQLParam(String sql,List paramList,String qryCount) throws SQLException {
    	if(connection == null){
    	   System.out.println("connection is null");
    	}       
        
        
        //System.out.println("paramList.size()="+paramList.size());
        System.out.println(sql);       
        getPreparedStatement(sql);//102.10.02
        /*102.10.02
        PreparedStatement pstmt = null;
        pstmt = getPreparedStatement(sql);
        */
        paraList.add(0, "");
        if(paramList != null){        
            //System.out.println("setParameter");
        	for(int i = 0; i < paramList.size(); i++) {
        		//System.out.println("i="+i+":"+paramList.get(i));
        		//pstmt.setObject(i+1, paramList.get(i)); 
                setParameter(i + 1, paramList.get(i));
        	}
        }
        
        if (paraList.size() > 1) {
        	
        	SQLUtils.printQuerySQL(sql,paraList);
        	
            for (int i = 1; i < paraList.size(); i++) {
                SQLParameter sp = (SQLParameter) paraList.get(i);
                //System.out.println(i+"="+sp.getValue());
                if(sp.getSQLType() != 0) {
                    pstmt.setObject(i, sp.getValue(), sp.getSQLType());
                 } else {
                   pstmt.setObject(i, sp.getValue());  
                }                    
            }
        }
        
        if(qryCount != null && !qryCount.equals("")){  
           System.out.println("qryCount.exe="+qryCount);      
		   stat.setMaxRows(Integer.parseInt(qryCount));
        }
        rs = pstmt.executeQuery(); //102.10.03 fix       
        return rs;
    }
    public ResultSet executeQuery(String sql,String qryCount,Connection conn) throws SQLException {
    	if(conn != null){
    	   System.out.println("connection is not null");
    	   this.connection = conn;
    	}
        stat = connection.createStatement();
        
        if(qryCount != null && !qryCount.equals("")){  
           System.out.println("qryCount.exe="+qryCount);      
		   stat.setMaxRows(Integer.parseInt(qryCount));
        }
        
        ResultSet rs = stat.executeQuery(sql);     	
       
        return rs;
    }
    public void setAutoCommit(boolean commit) throws SQLException{
    	System.out.print("setAutoCommit begin--");
    	connection.setAutoCommit(commit);    	
    	System.out.println("setAutoCommit end");
    }
    public void commit() throws SQLException{
    	System.out.print("Commit begin--");
    	connection.commit();    	
    	System.out.println("Commit end");
    }
    public void rollback() throws SQLException{
    	System.out.print("rollback begin--");
    	connection.rollback();    	
    	System.out.println("rollback end");
    }    
    
    public void setParameter(int index, Object value) {
        SQLParameter sp = new SQLParameter(index, value);
        paraList.add(index, sp);
    }

    public void setParameter(int index, Object value, int SQLType) {
        SQLParameter sp = new SQLParameter(index, value, SQLType);
        paraList.add(index, sp);
    } 
}
