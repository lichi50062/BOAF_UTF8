package com.tradevan.util.sql;

import java.sql.*;
import java.util.*;
import java.lang.reflect.*;

/**
 * @author Administrator
 *
 */
public class Query {
    private Class bean = null;
    private DBConnectFactory dbfactory = null;
    //private Connection conn = null;

    private Map methodMap = null;
    private List paraList = new ArrayList();

    private String sql = null;
    
   
   
    public Query(DBConnectFactory dbfactory, String sql) {
        this.sql = sql;
        this.dbfactory = dbfactory;
        paraList.add(0, "");
    }
    
    /**
     *  put query result to bean .
     * @param bean 
     * @return List
     * @throws Exception
     */
    public List beanList(Class bean) throws Exception {
        if (dbfactory == null) {
            throw new SQLTemplateException("The Connection is null.");
        }
        if (sql == null) {
            throw new SQLTemplateException("The SQL statement is null.");
        }
        this.bean = bean;
        return executeQuery("BEAN");
    }

    private void debug(String s) {
        // System.out.println("[Query] : " + s);
    }

    private List executeQuery(String outputType) {
        List resultList = new ArrayList();
        Connection conn = dbfactory.getDBConnection();
        PreparedStatement pstmt = null;
        ResultSet rs = null;
        try {
            pstmt = conn.prepareStatement(sql);
            if (paraList.size() > 1) {
                for (int i = 1; i < paraList.size(); i++) {
                    SQLParameter sp = (SQLParameter) paraList.get(i);
                    if(sp.getSQLType() != 0) {
                        pstmt.setObject(i, sp.getValue(), sp.getSQLType());
                    } else {
                        pstmt.setObject(i, sp.getValue());
                        
                    }                    
                }
            }
            SQLUtils.printQuerySQL(sql,paraList);

            rs = pstmt.executeQuery();
            while (rs.next()) {
                if ("LIST".equals(outputType)) {
                    resultList.add(saveToMap(rs));
                } else if ("ARRAY".equals(outputType)) {
                    resultList.add(saveToArray(rs));
                } else if ("BEAN".equals(outputType)) {
                    resultList.add(saveToBean(rs));
                }

            }
            

        } catch (Exception e) {
            e.printStackTrace();
        }  finally {
            SQLUtils.closeResultSet(rs);
            SQLUtils.closeStatement(pstmt);
            SQLUtils.closeConnection(conn);
            debug("free connection");
        }
        
        return resultList;

    }

    public Object[] getResult() throws Exception {
        if (dbfactory == null) {
            throw new SQLTemplateException("The Connection is null.");
        }
        if (sql == null) {
            throw new SQLTemplateException("The SQL statement is null.");
        }
        return executeQuery("ARRAY").toArray();

    }

    public List list() throws Exception {
        if (dbfactory == null) {
            throw new SQLTemplateException("The Connection is null.");
        }
        if (sql == null) {
            throw new SQLTemplateException("The SQL statement is null.");
        }

        return executeQuery("LIST");

    }

    private void parserBean() throws Exception {
        methodMap = new HashMap();

        // search all method's name
        Method[] methods = bean.getMethods();
        debug("Bean Name : " + bean.getName());
        for (int i = 0; i < methods.length; i++) {
            if ("set".equals(methods[i].getName().substring(0, 3))) {
                // get argument type
                Class[] para = methods[i].getParameterTypes();

                if (para.length == 1) {
                    if (para[0].getName().indexOf("String") != -1) {
                        methodMap.put(methods[i].getName().substring(3, methods[i].getName().length()).toLowerCase(), methods[i].getName());
                    }
                }

            }
        }
    }

    
    private Object saveToArray(ResultSet rs) throws Exception {
        // Map resultMap = new HashMap();
        ResultSetMetaData rsm = rs.getMetaData();
        int colCount = rsm.getColumnCount();
        Object[] obj = new Object[colCount];
        debug("column count : " + colCount);
        for (int i = 0; i < colCount; i++) {

            obj[i] = rs.getObject((i + 1));
            debug("column " + i + "  ==>  [" + obj[i] + "]   type : " + rsm.getColumnTypeName((i + 1)));

        }
        return obj;
    }

    private Object saveToBean(ResultSet rs) throws Exception {
        parserBean();

        Map resultMap = new HashMap();
        ResultSetMetaData rsm = rs.getMetaData();
        int colCount = rsm.getColumnCount() + 1;
        debug("column count : " + (colCount - 1));

        // create an object
        Object newBean = bean.newInstance();

        for (int i = 1; i < colCount; i++) {
            String colName = rsm.getColumnName(i).toLowerCase();

            // set arguments type
            Class[] cl = {java.lang.String.class};

            // set arguments value
            String temp = rs.getObject(i) != null ?  rs.getObject(i).toString() : "";
            Object[] obj = {temp};

            resultMap.put(colName, rs.getObject(i));
            String methodName = (String) methodMap.get(colName);

            // find method of the class
            Method method = bean.getMethod(methodName, cl);
            // invoke the method
            method.invoke(newBean, obj);
            debug(obj[0] + " ===>  " + methodName + " ");
        }

        return newBean;
    }

    private Map saveToMap(ResultSet rs) throws Exception {
        Map resultMap = new HashMap();
        ResultSetMetaData rsm = rs.getMetaData();
        int colCount = rsm.getColumnCount() + 1;
        debug("column count : " + (colCount - 1));
        for (int i = 1; i < colCount; i++) {
            String colName = rsm.getColumnName(i).toLowerCase();
            resultMap.put(colName, rs.getObject(i));

            debug(colName + "  [" + rs.getObject(i) + "]   type : " + rsm.getColumnTypeName(i));

        }
        return resultMap;
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