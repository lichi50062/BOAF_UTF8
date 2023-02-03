package com.tradevan.util.sql;

import java.sql.*;
import java.util.*;

import com.tradevan.util.sql.StoredProcedure.OutputParameter;

public class DMLStatement {

    private DBConnectFactory dbfactory = null;

    private String sql = null;
    private List paraList = new ArrayList();

    DMLStatement(DBConnectFactory dbfactory, String sql) {
        this.sql = sql;
        this.dbfactory = dbfactory;
        paraList.add(0, "");
    }

    public boolean execute() throws Exception {
        Connection conn = dbfactory.getDBConnection();
        boolean flag = false;

        PreparedStatement pstmt = null;
        try {
            pstmt = conn.prepareCall(sql);

            for (int i = 1; i < paraList.size(); i++) {
                SQLParameter sp = (SQLParameter) paraList.get(i);

                if (sp.getIndex() != i)
                    throw new Exception(" PreparedStatement index is not mapping");

                if (sp.getSQLType() != 0) {
                    pstmt.setObject(i, sp.getValue(), sp.getSQLType());
                } else {
                    pstmt.setObject(i, sp.getValue());
                }
            }

            SQLUtils.printQuerySQL(sql, paraList);
            flag = pstmt.execute();

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            SQLUtils.closeStatement(pstmt);
            SQLUtils.closeConnection(conn);
        }
        return flag;
    }

    public int executeUpdate() throws Exception {
        Connection conn = dbfactory.getDBConnection();
        int flag = 0;

        PreparedStatement pstmt = null;
        try {
            pstmt = conn.prepareCall(sql);

            for (int i = 1; i < paraList.size(); i++) {
                SQLParameter sp = (SQLParameter) paraList.get(i);

                if (sp.getIndex() != i)
                    throw new Exception(" PreparedStatement index is not mapping");

                if (sp.getSQLType() != 0) {
                    pstmt.setObject(i, sp.getValue(), sp.getSQLType());
                } else {
                    pstmt.setObject(i, sp.getValue());
                }
            }

            SQLUtils.printQuerySQL(sql, paraList);
            flag = pstmt.executeUpdate();

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            SQLUtils.closeStatement(pstmt);
            SQLUtils.closeConnection(conn);
        }
        return flag;
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
