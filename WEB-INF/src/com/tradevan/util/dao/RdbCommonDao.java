//95.03.15 add close statement by 2295
//98.06.17 add type=CLOB/BLOB,增加null處理 by 2295
//99.01.29 add 列印PreparedStatement轉換後的SQL statement by 2295
//99.02.03 add 使用PreparedStatement;並列印轉換後的SQL;
//             QueryDB_SQLParam(sql,getFields,orgTypeFields,qryCount,paramList) by 2295
//99.02.03 add updateDB_ps(List updateSQL)使用PreparedStatement,並列印轉換後的SQL statement by 2295
//             updateSQL 0:要執行的PreparedStatement 1:參數List;若參數為null則不加入參數
//             列印PreparedStatement轉換後的SQL statement 
//102.06.03 add updateDB_ps.每個PreparedStatement執行完後.做close by 2295
//102.10.02 add updateDB_ps改成batch作業 by 2295
//          add 原本的updateDB增加 closeStatement by 2295
//104.10.07 add 部份close後,設成null by 2295
//111.02.15 fix updateDB_ps SQLUtil中的參數有?時,轉成大寫？ by 2295
package com.tradevan.util.dao;

import java.io.*;
import java.sql.*;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;

import javax.naming.NamingException;
import com.tradevan.util.sql.SQLParameter;
import com.tradevan.util.sql.SQLUtils;

//import javax.transaction.UserTransaction;
//import com.tradevan.util.dao.DAOFactory;

public class RdbCommonDao {
    private DaoUtil daoutil;
    private String table;
    private boolean encoding = false;
    private String errMsg = "";

    public RdbCommonDao(String poolname) {
        try {            
            if (!poolname.equals("")) {
                daoutil = new DaoUtil(poolname);              
            } else {
                daoutil = new DaoUtil("PBOAFPool");               
            }
        } catch (Exception e) {
            System.out.println("new RdbCommonDao Error~~");
            e.printStackTrace();
        }
    }

    public String getErrMsg() {
        return errMsg;
    }

    // 原update
    public boolean updateDB(List updateSQL) throws Exception {
        boolean updateOK = true;
        int i = 0;
        Statement DBstat = null;
        errMsg = "";
        try {
            daoutil.openConnection();
            daoutil.setAutoCommit(false);
            DBstat = daoutil.getStatement();
            for (int idx = 0; idx < updateSQL.size(); idx++) {
                DBstat.addBatch((String) updateSQL.get(idx));
                // System.out.println((String)updateSQL.get(idx));
            }

            int[] rowCount = DBstat.executeBatch();
            System.out.println("rowCount=" + rowCount.length);
            while (i < rowCount.length) {
                if (rowCount[i] <= 0) {
                    System.out.println("i=" + i + ":" + rowCount[i] + ":sql="
                            + (String) updateSQL.get(i));
                    updateOK = false;
                }
                i++;
            }

            if (!updateOK) {
                try {
                    daoutil.rollback();
                } catch (Exception ex) {
                    ex.printStackTrace();
                }
                System.out.println("Update Fail");
                errMsg = errMsg + "資料庫中的資料並未更動或新增";
                return false;
            } else {
                System.out.print("Update OK begin--");
                daoutil.commit();
                System.out.println("Update OK end");
                return true;
            }
        } catch (Exception e) {
            System.out.println("RdbCommonDAO.updateDB Error[" + e + "]:"
                    + e.getMessage());
            errMsg = errMsg + "RdbCommonDAO.updateDB Error";
            try {
                daoutil.rollback();
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        } finally {
            if (DBstat != null) {
                DBstat.close();
                //DBstat = null;// 104.10.06 add //108.10.30 fix
            }
            daoutil.closeStatement();// 102.10.02 add
            daoutil.closeConnection();
        }
        return false;
    }

    // 95.04.19 add 使用PreparedStatement by 2295
    // 99.01.29 add 列印PreparedStatement轉換後的SQL statement by 2295
    //102.06.03 add ps.close() by 2295
    //102.10.02 fix 改成batch作業 by 2295
    //111.02.15 fix SQLUtil中的參數有?時,轉成大寫？ by 2295
    public boolean updateDB_ps(List updateSQL) throws Exception {
        // updateSQL 0:要執行的PreparedStatement 1:參數List
        // 若參數為null則不加入參數
        boolean updateOK = true;
        int i = 0;
        int psidx = 0;
        // Statement DBstat = null;
        PreparedStatement ps = null;
        List ps_data = null;
        errMsg = "";
        int rowCount = 0;
        List paraList = new ArrayList();// 99.01.29列印PreparedStatement轉換成SQL需使用的參數List
        try {
            daoutil.openConnection();
            daoutil.setAutoCommit(false);
            // DBstat = daoutil.getStatement();

            paraList.add(0, "");// 99.01.29第0個.預設需塞空白

            System.out.println("new updateSQL.size()=" + updateSQL.size());
            for (int sqlidx = 0; sqlidx < updateSQL.size(); sqlidx++) {
                ps = daoutil.getPreparedStatement((String) ((List) updateSQL
                        .get(sqlidx)).get(0));
                if (((List) updateSQL.get(sqlidx)).get(1) == null) {// 無參數
                    rowCount = ps.executeUpdate();
                    System.out.println("[rowCount=]" + rowCount);
                    if (rowCount == 0) {
                        try {
                            daoutil.rollback();
                            if (ps != null) {// 自己建的pstmt
                                ps.close();
                                ps = null;// 104.10.06 add
                                System.out.println("ps.close");
                            }
                        } catch (Exception ex) {
                            ex.printStackTrace();
                        }
                        daoutil.closeStatement();// 102.10.02關共用的stat/pstmt
                        daoutil.closeConnection();
                        System.out.println("Update Fail");
                        errMsg = errMsg + "資料庫中的資料並未更動或新增";
                        return false;
                    }
                } else {// 有參數
                    ps_data = ((List) ((List) updateSQL.get(sqlidx)).get(1));// ps的data
                    System.out.print("ps_data.size()=" + ps_data.size() + ":");
                    System.out.println("data size="
                            + ((List) ps_data.get(0)).size());                    
                    ps_dataLoop: for (i = 0; i < ps_data.size(); i++) {// 逐筆加入參數
                        psidx = 1;
                        for (int dataidx = 0; dataidx < ((List) ps_data.get(i))
                                .size(); dataidx++) {
                            //System.out.print(dataidx+":"+(String)((List)ps_data.get(i)).get(dataidx));     
                                //111.02.15 fix SQLUtil中的參數有?時,轉成大寫？
                            	setParameter(psidx,
                            			((String)((List)ps_data.get(i)).get(dataidx)).replaceAll("\\?", "？"),
                                        paraList);// 99.01.29設定SQLUtil需使用的參數
                            ps.setObject(psidx++,
                                    ((List) ps_data.get(i)).get(dataidx));
                        }// end of ps_data.get(i)

                        //System.out.println("paraList.size()="+paraList.size());
                        
                        SQLUtils.printQuerySQL(
                                (String) ((List) updateSQL.get(sqlidx)).get(0),
                                paraList);// 99.01.29列印轉換後的SQL
                         
                        ps.addBatch();// 102.10.02
                        /*
                         * 102.10.02 rowCount = ps.executeUpdate();
                         * System.out.println("i="+i+"[rowCount=]"+rowCount);
                         * if(rowCount == 0){ try{ daoutil.rollback();
                         * //102.06.03 add }catch(Exception ex){
                         * ex.printStackTrace(); }
                         * System.out.println("Update Fail"); break ps_dataLoop;
                         * }
                         */
                    }// end of ps_data加入參數
                    int[] rowCount_ps = ps.executeBatch();
                    System.out.println("rowCount_ps=" + rowCount_ps.length);
                    i = 0;
                    updateOK = true;
                    while (i < rowCount_ps.length) {
                        // System.out.print("i="+i+":"+rowCount_ps[i]);
                        if (rowCount_ps[i] == PreparedStatement.SUCCESS_NO_INFO) {
                            // System.out.println("更新成功");
                        }
                        if (!(rowCount_ps[i] == PreparedStatement.SUCCESS_NO_INFO || rowCount_ps[i] >= 0)) {
                            updateOK = false;
                            System.out.print("i=" + i + ":" + rowCount_ps[i]
                                    + ":更新失敗");
                        }
                        /*
                         * if(rowCount_ps[i] == PreparedStatement.EXECUTE_FAILED
                         * ){ System.out.println("更新失敗"); } if(rowCount_ps[i] <=
                         * 0){ updateOK = false; }
                         */
                        i++;
                    }

                    if (!updateOK) {
                        try {
                            daoutil.rollback();
                            if (ps != null) {// 自己建的pstmt
                                ps.close();
                                //ps = null;// 104.10.06 add //108.10.30 fix
                                System.out.println("ps.close");
                            }
                        } catch (Exception ex) {
                            ex.printStackTrace();
                        }

                        daoutil.closeStatement();// 102.10.02關共用的stat/pstmt
                        daoutil.closeConnection();
                        System.out.println("Update Fail");
                        errMsg = errMsg + "資料庫中的資料並未更動或新增";
                        return false;
                    }// 更新失敗
                }// end of updateSQL.get(i).get(1).size() != 0有參數
                if (ps != null) {// 自己建的pstmt
                    ps.close();
                    //ps = null;// 104.10.07 add //108.10.30 fix
                    System.out.println("ps.close");
                }
            }// end of updateSQL
            /*
             * 102.10.02 if(rowCount == 0){ try{ daoutil.rollback();
             * }catch(Exception ex){ ex.printStackTrace(); }
             * System.out.println("Update Fail"); errMsg = errMsg +
             * "資料庫中的資料並未更動或新增"; return false; }else{
             */
            // System.out.print("Update OK begin--");
            daoutil.commit();
            System.out.println("Update OK end");
            return true;
            // }
            // return false;
        } catch (Exception e) {
            System.out.println("RdbCommonDAO.updateDB_ps Error[" + e + "]:"
                    + e.getMessage());
            errMsg = errMsg + "RdbCommonDAO.updateDB_ps Error"+e.getMessage();
            try {
                daoutil.rollback();
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        } finally {
            if (ps != null) {
                ps.close();
                //ps = null;// 104.10.06 add //108.10.30 fix
                System.out.println("ps.close");
            }
            daoutil.closeStatement();// 102.10.02關共用的stat/pstmt
            daoutil.closeConnection();
        }
        return false;
    }

    public boolean updateDB(List updateSQL, String pgName) throws Exception {
        boolean updateOK = true;
        int i = 0;
        Statement DBstat = null;
        errMsg = "";
        try {
            daoutil.openConnection(pgName);
            daoutil.setAutoCommit(false);
            DBstat = daoutil.getStatement();
            for (int idx = 0; idx < updateSQL.size(); idx++) {
                DBstat.addBatch((String) updateSQL.get(idx));
                System.out.println((String) updateSQL.get(idx));
            }

            int[] rowCount = DBstat.executeBatch();
            System.out.println("rowCount=" + rowCount.length);
            while (i < rowCount.length) {
                System.out.println("i=" + i + ":" + rowCount[i]);
                if (rowCount[i] <= 0) {
                    updateOK = false;
                }
                i++;
            }

            if (!updateOK) {
                try {
                    daoutil.rollback();
                } catch (Exception ex) {
                    ex.printStackTrace();
                }
                System.out.println("Update Fail");
                errMsg = errMsg + "資料庫中的資料並未更動或新增";
                return false;
            } else {
                System.out.print("Update OK begin--");
                daoutil.commit();
                System.out.println("Update OK end");
                return true;
            }
        } catch (Exception e) {
            System.out.println("RdbCommonDAO.updateDB Error[" + e + "]:"
                    + e.getMessage());
            errMsg = errMsg + "RdbCommonDAO.updateDB Error";
            try {
                daoutil.rollback();
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        } finally {
            if (DBstat != null)
                DBstat.close();
            daoutil.closeStatement();// 102.10.02關共用的stat/pstmt
            daoutil.closeConnection(pgName);
        }
        return false;
    }

    /*
     * 套用Transaction public boolean updateDB(List updateSQL) throws Exception{
     * //UserTransaction tx = null; try {
     * 
     * //tx = daoutil.getTransaction(); //tx.begin(); daoutil.openConnection();
     * Statement DBstat = daoutil.getStatement(); int rowCount = 0; for(int idx
     * = 0;idx < updateSQL.size();idx++){
     * DBstat.executeUpdate((String)updateSQL.get(idx)); } if (DBstat != null)
     * DBstat.close(); daoutil.closeConnection();
     * 
     * if(rowCount == 0){ return false; }else{ tx.commit(); return true; }
     * }catch (Exception e){ System.out.println(e+":"+e.getMessage()); errMsg =
     * errMsg + e.getMessage(); try{ tx.rollback(); }catch(Exception ex){
     * ex.printStackTrace(); } } return false; }
     */

    // 96.05.01 add 使用PreparedStatement by 2295
    // UpdateA01在使用
    public List QueryDB_ps(List querySQL, String getFields,
            String orgTypeFields, String qryCount) throws Exception {
        // querySQL 0:要執行的PreparedStatement 1:參數List
        // 若參數為null則不加入參數
        boolean updateOK = true;
        int i = 0;
        int psidx = 0;

        PreparedStatement ps = null;
        List ps_data = null;
        errMsg = "";
        int rowCount = 0;
        ResultSet rs = null;
        System.out.print("A01 Use ");
        List list = new LinkedList();
        try {
            daoutil.openConnection();
            // daoutil.setAutoCommit(false);
            // DBstat = daoutil.getStatement();
            System.out.println("querySQL.size()=" + querySQL.size());
            System.out.println(querySQL.get(0));
            // System.out.println("test1="+((List)querySQL.get(0)).get(0));
            // System.out.println("test2="+((List)querySQL.get(0)).get(1));
            for (int sqlidx = 0; sqlidx < querySQL.size(); sqlidx++) {
                ps = daoutil.getPreparedStatement((String) ((List) querySQL
                        .get(sqlidx)).get(0));
                System.out.println("SQL="
                        + (String) ((List) querySQL.get(sqlidx)).get(0));
                if (((List) querySQL.get(sqlidx)).get(1) == null) {// 沒有參數
                    rs = ps.executeQuery();
                } else {// 有參數
                    ps_data = ((List) ((List) querySQL.get(sqlidx)).get(1));// ps的data
                    System.out.print("ps_data.size()=" + ps_data.size() + ":");
                    System.out.println("data size="
                            + ((List) ps_data.get(0)).size());
                    ps_dataLoop: for (i = 0; i < ps_data.size(); i++) {
                        psidx = 1;
                        for (int dataidx = 0; dataidx < ((List) ps_data.get(i))
                                .size(); dataidx++) {
                            // System.out.println((String)((List)ps_data.get(i)).get(dataidx));
                            ps.setObject(psidx++,
                                    ((List) ps_data.get(i)).get(dataidx));
                        }// end of ps_data.get(i)
                        rs = ps.executeQuery();
                    }// end of ps_data
                }// end of querySQL(i).get(1).size() != 0
                if (ps != null) {// 104.10.07 add
                    ps.close();
                   // ps = null; //108.10.30 fix
                }
            }// end of updateSQL
            while (rs.next()) {
                // System.out.println("have data");
                list.add(RsToDataObject(rs, getFields, orgTypeFields));
            }
            // return false;
        } catch (Exception e) {
            System.out.println("RdbCommonDAO.QueryDB_ps Error[" + e + "]:"
                    + e.getMessage());
            errMsg = errMsg + "RdbCommonDAO.QueryDB_ps Error";
        } finally {
            if (ps != null) {
                ps.close();
                //ps = null; //108.10.30 fix
            }
            if (rs != null) {// 104.10.07 add
                rs.close();
                //rs = null; //108.10.30 fix
            }
            daoutil.closeStatement();
            daoutil.closeConnection();
        }
        return list;
    }

    public List QueryDB(String sql, String getFields, String orgTypeFields,
            String qryCount) throws Exception {
        System.out.println("QueryDB.SQL: " + sql);
        List list = new LinkedList();
        ResultSet rs = null;
        try {
             daoutil.openConnection();//108.11.26 fix
             rs = daoutil.executeQuery(sql, qryCount);
            while (rs.next()) {
                System.out.println("have data");
                list.add(RsToDataObject(rs, getFields, orgTypeFields));
            }
            if (rs != null)
                rs.close();
            //rs = null;// 102.04.09 add //108.10.30 fix
            if (daoutil.stat != null)
                daoutil.stat.close();
            //daoutil.stat = null;// 102.04.09 //108.10.30 fix
            daoutil.connection.close();
            //daoutil.connection = null; //108.10.30 fix
            System.out.println("QueryDB close end");
        } catch (Exception e) {
            System.out.println("RdbCommonDAO.QueryDB Error[" + e + "]:"
                    + e.getMessage());
            errMsg = errMsg + "RdbCommonDAO.QueryDB Error";
        } finally {
            if (rs != null) rs.close();
            //rs = null;// 102.04.09 add //108.10.30 fix
            if (daoutil.stat != null) daoutil.stat.close();
            //daoutil.stat = null;// 102.04.09 //108.10.30 fix
            if (daoutil.connection != null)  daoutil.connection.close();
            //daoutil.connection = null; //108.10.30 fix
            // daoutil.closeStatement();//95.03.15 add
            // daoutil.closeConnection();
        }
        return list;
    }

    // 99.02.03 add 使用PreparedStatement;並列印轉換後的SQL by 2295
    public List QueryDB_SQLParam(String sql, String getFields,
            String orgTypeFields, String qryCount, List paramList)
            throws Exception {
        System.out.println("QueryDB_SQLParam.SQL: " + sql);
        List list = new LinkedList();
        ResultSet rs = null;
        DateFormat dateFormat = new SimpleDateFormat("yyyyMMddHHmmssSSS");
        //Calendar cal = Calendar.getInstance();
        try {
            daoutil.openConnection(); //108.11.26 fix
            //cal = Calendar.getInstance();        
            //System.out.println("2="+dateFormat.format(cal.getTime()));
            //System.out.println("getConnection end");           
            
            rs = daoutil.executeQuery_SQLParam(sql, paramList, qryCount);
            //cal = Calendar.getInstance();     
            //System.out.println("3="+dateFormat.format(cal.getTime()));
            while (rs.next()) {
                list.add(RsToDataObject(rs, getFields, orgTypeFields));
            }
            //cal = Calendar.getInstance();     
            //System.out.println("4="+dateFormat.format(cal.getTime()));
            if (rs != null) {
                rs.close();
                //rs = null;// 104.10.06 //108.10.30 fix
            }
        } catch (Exception e) {
            System.out.println("RdbCommonDAO.QueryDB_SQLParam Error[" + e
                    + "]:" + e.getMessage());
            errMsg = errMsg + "RdbCommonDAO.QueryDB_SQLParam Error";
        } finally {
            daoutil.closeStatement();// 95.03.15 add
            //cal = Calendar.getInstance();     
            //System.out.println("5="+dateFormat.format(cal.getTime()));
            daoutil.closeConnection();
            //cal = Calendar.getInstance();     
            //System.out.println("6="+dateFormat.format(cal.getTime()));
        }
        return list;
    }
    
    public List QueryDB_SQLParam_new(String sql, String getFields,
            String orgTypeFields, String qryCount, List paramList)
            throws Exception {
        System.out.println("QueryDB_SQLParam_new.SQL: " + sql);
        List list = new LinkedList();
        ResultSet rs = null;
        DateFormat dateFormat = new SimpleDateFormat("yyyyMMddHHmmssSSS");
        Calendar cal = Calendar.getInstance();
        try {
            daoutil.openConnection(); //108.11.26 fix
            
            rs = daoutil.executeQuery_SQLParam(sql, paramList, qryCount);
            //cal = Calendar.getInstance();     
            //System.out.println("3="+dateFormat.format(cal.getTime()));
            while (rs.next()) {
                list.add(RsToDataObject(rs, getFields, orgTypeFields));
            }
            //cal = Calendar.getInstance();     
            //System.out.println("4="+dateFormat.format(cal.getTime()));
            if (rs != null) {
                rs.close();
                //rs = null;// 104.10.06 //108.10.30 fix
            }
        } catch (Exception e) {
            System.out.println("RdbCommonDAO.QueryDB_SQLParam Error[" + e
                    + "]:" + e.getMessage());
            errMsg = errMsg + "RdbCommonDAO.QueryDB_SQLParam Error";
        } finally {
            daoutil.closeStatement();// 95.03.15 add
            //cal = Calendar.getInstance();     
            //System.out.println("5="+dateFormat.format(cal.getTime()));
            //daoutil.closeConnection();//109.07.09 add
            //cal = Calendar.getInstance();//109.07.09 add     
            //System.out.println("6="+dateFormat.format(cal.getTime()));//109.07.09 add
        }
        return list;
    }

    public List QueryDB(String sql, String getFields, String orgTypeFields,
            String qryCount, String pgName) throws Exception {
        System.out.println("SQL: " + sql);
        List list = new LinkedList();
        ResultSet rs = null;
        try {
            daoutil.openConnection(pgName);
            rs = daoutil.executeQuery(sql, qryCount);
            while (rs.next()) {
                // System.out.println("have data");
                list.add(RsToDataObject(rs, getFields, orgTypeFields));
            }
            if (rs != null) {
                rs.close();
                //rs = null;// 104.10.06 //108.10.30 fix
            }
        } catch (Exception e) {
            System.out.println("RdbCommonDAO.QueryDB Error[" + e + "]:"
                    + e.getMessage());
            errMsg = errMsg + "RdbCommonDAO.QueryDB Error";
        } finally {
            daoutil.closeStatement();// 95.03.15 add
            daoutil.closeConnection(pgName);
        }
        return list;
    }

    public void newQryConnection() throws NamingException, SQLException {
        Connection conn = daoutil.newConnection();
        // return conn;
        daoutil.setConnection(conn);
        // daoutil.openConnection();
    }

    public void closeQryConnection() throws SQLException {
        daoutil.closeConnection();
    }

    public Connection newConnection() throws NamingException, SQLException {
        Connection conn = daoutil.newConnection();
        return conn;
    }

    private DataObject RsToDataObject(ResultSet rs, String getFields,
            String orgTypeFields) throws SQLException {

        DataObject dataobject = null;
        ResultSetMetaData rsMetadata = rs.getMetaData();

        /*
         * if(rsMetadata != null){ System.out.println("rsMetadata != null");
         * }else{ System.out.println("rsMetadata == null"); }
         */

        int col = rsMetadata.getColumnCount();
        // System.out.println("col="+col);
        Map values = new HashMap();

        if (getFields != null)
            getFields = getFields.toLowerCase();
        if (orgTypeFields != null)
            orgTypeFields = orgTypeFields.toLowerCase();
        for (int i = 1; i <= col; i++) {
            String colName = rsMetadata.getColumnName(i).toLowerCase();
            // System.out.println("colName="+colName);
            if (getFields != null && getFields.indexOf(colName) == -1) {
                continue;
            }

            if (orgTypeFields != null && orgTypeFields.indexOf(colName) != -1) {
                if (orgTypeFields != null) {
                    // System.out.println("orgTypeFields != null");
                }
                // if(orgTypeFields.indexOf(colName) !=
                // -1){System.out.println("orgTypeFields.indexOf(colName) != -1");}
                // System.out.println(rs.getObject(colName).toString());
                // System.out.println((rs.getObject(colName)).getClass());

                values.put(colName, rs.getObject(colName));
                // System.out.println("values.put="+colName+":"+rs.getObject(colName).toString());
            } else {
                String colType = rsMetadata.getColumnTypeName(i);
                if (colType.indexOf("CHAR") != -1) {
                    // System.out.println("char");
                    if (encoding) {
                        values.put(colName, rs.getString(colName));
                        // System.out.println("put colName="+colName+":"+rs.getString(colName));
                    } else {
                        values.put(colName, rs.getString(colName));
                        // System.out.println("put colName="+colName+":"+rs.getString(colName));
                    }
                }
                if (colType.equals("BLOB")) {// 98.06.17 add null處理
                    if (rs.getBlob(colName) != null) {
                        values.put(colName, inStream2String(rs.getBlob(colName)
                                .getBinaryStream()));
                    }
                } else if (colType.equals("CLOB")) {// 98.06.17 add null處理
                    if (rs.getClob(colName) != null) {
                        System.out.println(Reader2String((rs.getClob(colName)
                                .getCharacterStream())));
                        values.put(colName, Reader2String((rs.getClob(colName)
                                .getCharacterStream())));
                    }
                }
            }
        }
        dataobject = new DataObject(values);
        return dataobject;
    }

    private String inStream2String(InputStream inStream) {
        StringBuffer strBuffer = new StringBuffer();
        String str = null;
        try {
            if (inStream == null)
                return "";
            InputStreamReader isr = new InputStreamReader(inStream);
            BufferedReader br = new BufferedReader(isr);

            while ((str = br.readLine()) != null) {
                strBuffer.append(str);
                strBuffer.append("\r\n");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return strBuffer.toString();
    }

    private String Reader2String(Reader in) {
        StringBuffer strBuffer = new StringBuffer();
        String str = null;
        try {
            if (in == null)
                return "";
            BufferedReader br = new BufferedReader(in);

            while ((str = br.readLine()) != null) {
                strBuffer.append(str);
                strBuffer.append("\r\n");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return strBuffer.toString();
    }

    public void setParameter(int index, Object value, List paraList) {
        SQLParameter sp = new SQLParameter(index, value);
        paraList.add(index, sp);
    }

    public void setParameter(int index, Object value, int SQLType, List paraList) {
        SQLParameter sp = new SQLParameter(index, value, SQLType);
        paraList.add(index, sp);
    }
    // ==================================================================
}
