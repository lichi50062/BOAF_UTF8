package com.tradevan.util.sql;

import java.sql.*;
import java.util.*;

public class StoredProcedure {

    private DBConnectFactory dbfactory = null;
    
    private String callSql = null;

    private Map returnMap = new HashMap();
    private List paraList = new ArrayList();

    public StoredProcedure(DBConnectFactory dbfactory, String callSql) {
        this.dbfactory = dbfactory;
        this.callSql = callSql;
        paraList.add("");
    }

    public void execute() throws Exception {
        Connection conn = dbfactory.getDBConnection();

        CallableStatement cstmt = null;
        try {
            cstmt = conn.prepareCall(callSql);

            for (int i = 1; i < paraList.size(); i++) {
                SQLParameter sp = (SQLParameter) paraList.get(i);

                if (sp.getIndex() != i)
                    throw new Exception(" CallabeStatment index is not mapping");

                if (sp instanceof OutputParameter) {
                    cstmt.registerOutParameter(i, sp.getSQLType());

                    returnMap.put(Integer.toString(i), sp);
                } else {

                    cstmt.setObject(i, sp.getValue());
                }
            }

            SQLUtils.printQuerySQL(callSql, paraList);
            cstmt.execute();

            parserResultToMap(cstmt);

            SQLUtils.closeStatement(cstmt);

        } catch (Exception e) {
            e.printStackTrace();
            SQLUtils.closeStatement(cstmt);

        }

    }

    private void parserResultToMap(CallableStatement cstmt) throws Exception {
        Iterator it = returnMap.keySet().iterator();

        while (it.hasNext()) {
            String key = (String) it.next();
            SQLParameter sp = (SQLParameter) returnMap.get(key);
            debug("Put return code : " + sp.getIndex() + " value:" + cstmt.getObject(sp.getIndex()));
            sp.setValue(cstmt.getObject(sp.getIndex()));
            returnMap.put(key, sp);
            sp = null;
        }
    }

    public void setInParameter(int index, Object value) {
        SQLParameter sp = new SQLParameter(index, value);
        paraList.add(index, sp);
    }

    public void setOutParameter(int index, int type) {
        OutputParameter sp = new OutputParameter(index, type);
        paraList.add(index, sp);
    }

    public Object getReturnObject(int index) {
        SQLParameter sp = (SQLParameter) returnMap.get(Integer.toString(index));

        return sp.getValue();
    }

    public class OutputParameter extends SQLParameter {
        public OutputParameter(int index, int sqlType) {
            super(index, null, sqlType);
        }
    }

   

    private void debug(String s) {
        System.out.println("[StoredProcedure] : " + s);
    }

}
