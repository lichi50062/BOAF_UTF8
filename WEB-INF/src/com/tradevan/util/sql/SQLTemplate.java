package com.tradevan.util.sql;

import java.sql.*;
import java.util.*;

public class SQLTemplate {
    
    private DBConnectFactory dbfactory = null;
    //private List pramlist = new ArrayList();




    public SQLTemplate(DBConnectFactory dbfactory) {
        this.dbfactory = dbfactory;
        
    }
    
    public StoredProcedure createProcedure(String sql) throws Exception {
        if(dbfactory == null) {
            throw new SQLTemplateException("The Connection is null.");
        }
        return new StoredProcedure(dbfactory,sql);
    }
 

    public Query createQuery(String sql) throws Exception{
        if(dbfactory == null) {
            throw new SQLTemplateException("The Connection is null.");
        }
        return new Query(dbfactory,sql);
    }
    
    public Query createQuery(String s, List list) throws Exception{
        if(dbfactory == null) 
           throw new SQLTemplateException("The Connection is null.");
        Query query = new Query(dbfactory, s);
        if(list != null){
           for(int i = 0; i < list.size(); i++)
               query.setParameter(i + 1, list.get(i));
        }
        return query;
    }
    
    public DMLStatement createStatement(String sql) throws Exception{
        if(dbfactory == null) {
            throw new SQLTemplateException("The Connection is null.");
        }
        return new DMLStatement(dbfactory,sql);
    }
    
    public DMLStatement createStatement(String s, List list) throws Exception{
        if(dbfactory == null)
           throw new SQLTemplateException("The Connection is null.");
        DMLStatement dmlstatement = new DMLStatement(dbfactory, s);
        if(list != null){
           for(int i = 0; i < list.size(); i++)
              dmlstatement.setParameter(i + 1, list.get(i));
        }
        return dmlstatement;
    }
    
    
    
    private void debug(String s) {
        // System.out.println("[Query] : " + s);
    }


}