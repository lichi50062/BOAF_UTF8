package com.tradevan.util;

import java.util.Hashtable;
import javax.servlet.http.HttpSessionAttributeListener;
import javax.servlet.http.HttpSessionBindingEvent;
import javax.servlet.http.HttpSessionListener;
import javax.servlet.http.HttpSessionEvent;
import javax.servlet.http.HttpSession;
import com.tradevan.util.DBManager;

import java.util.*;


public class UpdateTimeOut implements HttpSessionListener,HttpSessionAttributeListener {
    
    private Hashtable usertable = new Hashtable();
    private String muser_id = null;
    public synchronized void sessionCreated(HttpSessionEvent se) {
        
    }

    public synchronized void sessionDestroyed(HttpSessionEvent se) {
	  
	muser_id = (String) usertable.get(se.getSession().getId());
	System.out.println("muser_id;="+muser_id);
	usertable.remove(se.getSession().getId());
	
	List paramList =new ArrayList() ;  
    String sqlCmd = "UPDATE WTT01 SET "
               + " login_mark='N'"
               + ",update_date=sysdate"    
               + " where muser_id=? ";            
    paramList.add(muser_id) ;       
    try{
    updDbUsesPreparedStatement(sqlCmd,paramList) ;
    }catch(Exception e){System.out.println("UpdateTimeOut error:"+e+e.getMessage());}	       
    }
      
    public synchronized void attributeAdded(HttpSessionBindingEvent se) {
		if (se.getSession().getAttribute("muser_id") != null) {
		usertable.put(se.getSession().getId(), se.getSession().getAttribute("muser_id"));
      }
    }
    public synchronized void attributeRemoved(HttpSessionBindingEvent se) {
    }
    public synchronized void attributeReplaced(HttpSessionBindingEvent se) {
    }
    public boolean updDbUsesPreparedStatement(String sql ,List paramList) throws Exception{
        List updateDBList = new ArrayList();//0:sql 1:data
        List updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
        List updateDBDataList = new ArrayList();//儲存參數的List
        
        updateDBDataList.add(paramList);
        updateDBSqlList.add(sql);
        updateDBSqlList.add(updateDBDataList);
        updateDBList.add(updateDBSqlList);
        return DBManager.updateDB_ps(updateDBList) ;
    }
}

