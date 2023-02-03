package com.tradevan.util.report;

import java.util.*;
import java.sql.*;

public class RptCG002W_E {

  /**
   * 備份DB--EF組
   * @param debug boolean 列印debug 訊息
   * @param isDelete boolean 是否刪除LOG檔的資料
   * @param sYear String 起始年份
   * @param eYear String 結束年份
   * @param sMonth String 起始月份
   * @param eMonth String 結束月份
   * @param conn1 Connection 原資料庫
   * @param conn2 Connection 備份資料庫
   * @return String	產生錯誤的TB_NAME
   */
  public String bakDB(boolean debug, boolean isDelete,
                            String sYear, String eYear,
                            String sMonth, String eMonth,
                            Connection conn1, Connection conn2) {
      Statement st = null;
      Statement st2 = null;
      Statement st3 = null;
      Statement st4 = null;
      ResultSet rs = null;
      String sqlCmd = null;
      String sqlCmd2 = null;
      String tbNameNow = "";

      try {
        conn1.setAutoCommit(false);     //原DB(TBOAF)
        conn2.setAutoCommit(false);     //備份DB(BBOAF)
        st = conn1.createStatement();   //查詢xxx_LOG
        st2 = conn2.createStatement();  //維護xxx_BAK
        st3 = conn1.createStatement();  //維護STATISTICS_BAK
        st4 = conn1.createStatement();  //刪除xxx_LOG

        //取得機構名稱及類別
        HashMap bankMap = new HashMap();
        sqlCmd = "SELECT BANK_NO,BANK_NAME FROM BN01";
        rs = st.executeQuery(sqlCmd);
        while (rs.next()) {
          bankMap.put(rs.getString("BANK_NO"),rs.getString("BANK_NAME"));
        }

        //ExDefGoodF_log=============================================================
        String tbName12[] = {"ExDefGoodF_log"};
        for(int i=0;i<tbName12.length;i++){
          tbNameNow = tbName12[i];
          //STATISTICS_BAK
          String mYear = "";
          String mMonth = "";
          String userId = "";
          String userName = "";
          String bankNo = "";
          String bankType = "";
          String bankName = "";
          int updateNum = 0;
          int deleteNum = 0;
          int downloadNum = 0;
          int loginNum = 0;

          //查詢符合條件的資料
          sqlCmd = "SELECT "
              + "A.REPORTNO,A.REPORTNO_SEQ,A.ITEM_NO,A.EX_CONTENT,"
              + "A.COMMENTT,A.FAULT_ID,A.AUDIT_OPPINION,A.ACT_ID,A.AUDIT_RESULT,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName12[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName12[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName12[i]+"'"
                  + " AND ((M_YEAR = " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_MONTH BETWEEN " + sMonth + " AND 12)"
                  + " OR (M_YEAR = " + (Integer.parseInt(eYear) - 1911)
                  + " AND M_MONTH BETWEEN 1 AND " + eMonth + ")"
                  + " OR (M_YEAR > " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_YEAR < " + (Integer.parseInt(eYear) - 1911) + "))";
              if (debug) System.out.println("sqlCmd=" + sqlCmd);
              st3.addBatch(sqlCmd);
            }

            //取得資料
            String dbReportNo = RptCG002W.getTrimString(rs.getString("REPORTNO"));
            String dbReportNoSeq = RptCG002W.getTrimString(rs.getString("REPORTNO_SEQ"));
            String dbItemNo = RptCG002W.getTrimString(rs.getString("ITEM_NO"));
            String dbExContent = RptCG002W.getTrimString(rs.getString("EX_CONTENT"));
            String dbCommentt = RptCG002W.getTrimString(rs.getString("COMMENTT"));
            String dbFaultId = RptCG002W.getTrimString(rs.getString("FAULT_ID"));
            String dbAuditOppinion = RptCG002W.getTrimString(rs.getString("AUDIT_OPPINION"));
            String dbActId = RptCG002W.getTrimString(rs.getString("ACT_ID"));
            String dbAuditResult = RptCG002W.getTrimString(rs.getString("AUDIT_RESULT"));
            String dbUserIdC = RptCG002W.getTrimString(rs.getString("USER_ID_C"));
            String dbUserNameC = RptCG002W.getTrimString(rs.getString("USER_NAME_C"));
            Timestamp dbUpdateTimestamp = rs.getTimestamp("UPDATE_DATE_C");
            String dbUpdateDateC = dbUpdateTimestamp.toString();
            dbUpdateDateC = dbUpdateDateC.substring(0, dbUpdateDateC.indexOf("."));
            String dbUpdateTypeC = RptCG002W.getTrimString(rs.getString("UPDATE_TYPE_C"));
            String dbUYear = RptCG002W.getTrimString(rs.getString("U_YEAR"));
            String dbUMonth = RptCG002W.getTrimString(rs.getString("U_MONTH"));
            String tmBankType =rs.getString("TM_BANK_TYPE");
            String tmTBankNo = rs.getString("TM_TBANK_NO");
            String tmBankName = RptCG002W.getTrimString((String)bankMap.get(tmTBankNo)," ");

            //新增_BAK的TABLE中
            sqlCmd2 = "INSERT INTO "+tbName12[i]+"_BAK("
                + "REPORTNO,REPORTNO_SEQ,ITEM_NO,EX_CONTENT,"
                + "COMMENTT,FAULT_ID,AUDIT_OPPINION,ACT_ID,AUDIT_RESULT,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + "'" + dbReportNo + "','" + dbReportNoSeq + "',"
                + "'" + dbItemNo + "','" + dbExContent + "',"
                + "'" + dbCommentt + "','" + dbFaultId + "',"
                + "'" + dbAuditOppinion + "','" + dbActId + "',"
                + "'" + dbAuditResult + "',"
                + "'" + dbUserIdC + "','" + dbUserNameC + "',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUpdateTypeC +"')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName12[i] + "'," + mYear + "," + mMonth
                    + ",'" + userId + "','" + userName + "','" + bankNo
                    + "','" + bankType + "','" + bankName + "'," + updateNum
                    + "," + deleteNum + "," + downloadNum + "," + loginNum + ")";
                if (debug) System.out.println("sqlCmd=" + sqlCmd);
                st3.addBatch(sqlCmd);

                updateNum = 0;
                deleteNum = 0;
                downloadNum = 0;
              }
            }

            mYear = dbUYear;
            mMonth = dbUMonth;
            userId = dbUserIdC;
            userName = dbUserNameC;
            bankNo = tmTBankNo;
            bankType = tmBankType;
            bankName = tmBankName;

            //統計資料
            if (dbUpdateTypeC.equals("U")) {
              updateNum++;
            }else if (dbUpdateTypeC.equals("D")) {
              deleteNum++;
            }else if (dbUpdateTypeC.equals("L")) {
              downloadNum++;
            }
          }
          int rows[] = st2.executeBatch();  //用rows.length就知道一共變更了幾筆資料

          if (!mYear.equals("")) {
            sqlCmd = "INSERT INTO STATISTICS_BAK("
                + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                + "DOWNLOAD_NUM,LOGIN_NUM) "
                + "VALUES('" + tbName12[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName12[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //ExDG_HistoryF_log=============================================================
        String tbName13[] = {"ExDG_HistoryF_log"};
        for(int i=0;i<tbName13.length;i++){
          tbNameNow = tbName13[i];
          //STATISTICS_BAK
          String mYear = "";
          String mMonth = "";
          String userId = "";
          String userName = "";
          String bankNo = "";
          String bankType = "";
          String bankName = "";
          int updateNum = 0;
          int deleteNum = 0;
          int downloadNum = 0;
          int loginNum = 0;

          //查詢符合條件的資料
          sqlCmd = "SELECT "
              + "A.REPORTNO,A.REPORTNO_SEQ,A.DIGEST,A.RT_DOCNO,"
              + "A.RT_DATE,A.AUDIT_RESULT,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName13[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName13[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName13[i]+"'"
                  + " AND ((M_YEAR = " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_MONTH BETWEEN " + sMonth + " AND 12)"
                  + " OR (M_YEAR = " + (Integer.parseInt(eYear) - 1911)
                  + " AND M_MONTH BETWEEN 1 AND " + eMonth + ")"
                  + " OR (M_YEAR > " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_YEAR < " + (Integer.parseInt(eYear) - 1911) + "))";
              if (debug) System.out.println("sqlCmd=" + sqlCmd);
              st3.addBatch(sqlCmd);
            }

            //取得資料
            String dbReportNo = RptCG002W.getTrimString(rs.getString("REPORTNO"));
            String dbReportNoSeq = RptCG002W.getTrimString(rs.getString("REPORTNO_SEQ"));
            String dbDigest = RptCG002W.getTrimString(rs.getString("DIGEST"));
            String dbRtDocNo = RptCG002W.getTrimString(rs.getString("RT_DOCNO"));
            Timestamp dbRtDateTimestamp = rs.getTimestamp("RT_DATE");
            String dbRtDate = "";
            if(dbRtDateTimestamp != null){
              dbRtDate = dbRtDateTimestamp.toString();
              dbRtDate = dbRtDate.substring(0, dbRtDate.indexOf("."));
            }
            String dbAuditResult = RptCG002W.getTrimString(rs.getString("AUDIT_RESULT"));

            String dbUserIdC = RptCG002W.getTrimString(rs.getString("USER_ID_C"));
            String dbUserNameC = RptCG002W.getTrimString(rs.getString("USER_NAME_C"));
            Timestamp dbUpdateTimestamp = rs.getTimestamp("UPDATE_DATE_C");
            String dbUpdateDateC = dbUpdateTimestamp.toString();
            dbUpdateDateC = dbUpdateDateC.substring(0, dbUpdateDateC.indexOf("."));
            String dbUpdateTypeC = RptCG002W.getTrimString(rs.getString("UPDATE_TYPE_C"));
            String dbUYear = RptCG002W.getTrimString(rs.getString("U_YEAR"));
            String dbUMonth = RptCG002W.getTrimString(rs.getString("U_MONTH"));
            String tmBankType =rs.getString("TM_BANK_TYPE");
            String tmTBankNo = rs.getString("TM_TBANK_NO");
            String tmBankName = RptCG002W.getTrimString((String)bankMap.get(tmTBankNo)," ");

            //新增_BAK的TABLE中
            sqlCmd2 = "INSERT INTO "+tbName13[i]+"_BAK("
                + "REPORTNO,REPORTNO_SEQ,DIGEST,RT_DOCNO,"
                + "RT_DATE,AUDIT_RESULT,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + "'" + dbReportNo + "','" + dbReportNoSeq + "',"
                + "'" + dbDigest + "','" + dbRtDocNo + "',"
                + "TO_DATE('" + dbRtDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbAuditResult + "',"
                + "'" + dbUserIdC + "','" + dbUserNameC + "',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUpdateTypeC +"')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName13[i] + "'," + mYear + "," + mMonth
                    + ",'" + userId + "','" + userName + "','" + bankNo
                    + "','" + bankType + "','" + bankName + "'," + updateNum
                    + "," + deleteNum + "," + downloadNum + "," + loginNum + ")";
                if (debug) System.out.println("sqlCmd=" + sqlCmd);
                st3.addBatch(sqlCmd);

                updateNum = 0;
                deleteNum = 0;
                downloadNum = 0;
              }
            }

            mYear = dbUYear;
            mMonth = dbUMonth;
            userId = dbUserIdC;
            userName = dbUserNameC;
            bankNo = tmTBankNo;
            bankType = tmBankType;
            bankName = tmBankName;

            //統計資料
            if (dbUpdateTypeC.equals("U")) {
              updateNum++;
            }else if (dbUpdateTypeC.equals("D")) {
              deleteNum++;
            }else if (dbUpdateTypeC.equals("L")) {
              downloadNum++;
            }
          }
          int rows[] = st2.executeBatch();  //用rows.length就知道一共變更了幾筆資料

          if (!mYear.equals("")) {
            sqlCmd = "INSERT INTO STATISTICS_BAK("
                + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                + "DOWNLOAD_NUM,LOGIN_NUM) "
                + "VALUES('" + tbName13[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName13[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //ExDisTripF_log=============================================================
        String tbName14[] = {"ExDisTripF_log"};
        for(int i=0;i<tbName14.length;i++){
          tbNameNow = tbName14[i];
          //STATISTICS_BAK
          String mYear = "";
          String mMonth = "";
          String userId = "";
          String userName = "";
          String bankNo = "";
          String bankType = "";
          String bankName = "";
          int updateNum = 0;
          int deleteNum = 0;
          int downloadNum = 0;
          int loginNum = 0;

          //查詢符合條件的資料
          sqlCmd = "SELECT "
              + "A.DISP_ID,A.EXAM_ID,A.EXAM_DIV,A.APPR_DATE,"
              + "A.CH_TYPE,A.BASIS,A.PRJ_ITEM,A.PRJ_REMARK,"
              + "A.BANK_TYPE,A.TBANK_NO,A.BANK_NO,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName14[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName14[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName14[i]+"'"
                  + " AND ((M_YEAR = " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_MONTH BETWEEN " + sMonth + " AND 12)"
                  + " OR (M_YEAR = " + (Integer.parseInt(eYear) - 1911)
                  + " AND M_MONTH BETWEEN 1 AND " + eMonth + ")"
                  + " OR (M_YEAR > " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_YEAR < " + (Integer.parseInt(eYear) - 1911) + "))";
              if (debug) System.out.println("sqlCmd=" + sqlCmd);
              st3.addBatch(sqlCmd);
            }

            //取得資料
            String dbDispId = RptCG002W.getTrimString(rs.getString("DISP_ID"));
            String dbExamId = RptCG002W.getTrimString(rs.getString("EXAM_ID"));
            String dbExamDiv = RptCG002W.getTrimString(rs.getString("EXAM_DIV"));
            Timestamp dbApprDateTimestamp = rs.getTimestamp("APPR_DATE");
            String dbApprDate = "";
            if(dbApprDateTimestamp != null){
              dbApprDate = dbApprDateTimestamp.toString();
              dbApprDate = dbApprDate.substring(0, dbApprDate.indexOf("."));
            }
            String dbChType = RptCG002W.getTrimString(rs.getString("CH_TYPE"));
            String dbBasis = RptCG002W.getTrimString(rs.getString("BASIS"));
            String dbPrjItem = RptCG002W.getTrimString(rs.getString("PRJ_ITEM"));
            String dbPrjRemark = RptCG002W.getTrimString(rs.getString("PRJ_REMARK"));
            String dbBankType = RptCG002W.getTrimString(rs.getString("BANK_TYPE"));
            String dbTbankNo = RptCG002W.getTrimString(rs.getString("TBANK_NO"));
            String dbBankNo = RptCG002W.getTrimString(rs.getString("BANK_NO"));

            String dbUserIdC = RptCG002W.getTrimString(rs.getString("USER_ID_C"));
            String dbUserNameC = RptCG002W.getTrimString(rs.getString("USER_NAME_C"));
            Timestamp dbUpdateTimestamp = rs.getTimestamp("UPDATE_DATE_C");
            String dbUpdateDateC = dbUpdateTimestamp.toString();
            dbUpdateDateC = dbUpdateDateC.substring(0, dbUpdateDateC.indexOf("."));
            String dbUpdateTypeC = RptCG002W.getTrimString(rs.getString("UPDATE_TYPE_C"));
            String dbUYear = RptCG002W.getTrimString(rs.getString("U_YEAR"));
            String dbUMonth = RptCG002W.getTrimString(rs.getString("U_MONTH"));
            String tmBankType =rs.getString("TM_BANK_TYPE");
            String tmTBankNo = rs.getString("TM_TBANK_NO");
            String tmBankName = RptCG002W.getTrimString((String)bankMap.get(tmTBankNo)," ");

            //新增_BAK的TABLE中
            sqlCmd2 = "INSERT INTO "+tbName14[i]+"_BAK("
                + "DISP_ID,EXAM_ID,EXAM_DIV,APPR_DATE,"
                + "CH_TYPE,BASIS,PRJ_ITEM,PRJ_REMARK,"
                + "BANK_TYPE,TBANK_NO,BANK_NO,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + "'" + dbDispId + "','" + dbExamId + "',"
                + "'" + dbExamDiv + "',"
                + "TO_DATE('" + dbApprDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbChType + "','" + dbBasis + "',"
                + "'" + dbPrjItem + "','" + dbPrjRemark + "',"
                + "'" + dbBankType + "','" + dbTbankNo + "',"
                + "'" + dbBankNo + "',"
                + "'" + dbUserIdC + "','" + dbUserNameC + "',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUpdateTypeC +"')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName14[i] + "'," + mYear + "," + mMonth
                    + ",'" + userId + "','" + userName + "','" + bankNo
                    + "','" + bankType + "','" + bankName + "'," + updateNum
                    + "," + deleteNum + "," + downloadNum + "," + loginNum + ")";
                if (debug) System.out.println("sqlCmd=" + sqlCmd);
                st3.addBatch(sqlCmd);

                updateNum = 0;
                deleteNum = 0;
                downloadNum = 0;
              }
            }

            mYear = dbUYear;
            mMonth = dbUMonth;
            userId = dbUserIdC;
            userName = dbUserNameC;
            bankNo = tmTBankNo;
            bankType = tmBankType;
            bankName = tmBankName;

            //統計資料
            if (dbUpdateTypeC.equals("U")) {
              updateNum++;
            }else if (dbUpdateTypeC.equals("D")) {
              deleteNum++;
            }else if (dbUpdateTypeC.equals("L")) {
              downloadNum++;
            }
          }
          int rows[] = st2.executeBatch();  //用rows.length就知道一共變更了幾筆資料

          if (!mYear.equals("")) {
            sqlCmd = "INSERT INTO STATISTICS_BAK("
                + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                + "DOWNLOAD_NUM,LOGIN_NUM) "
                + "VALUES('" + tbName14[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName14[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //ExHelpItemF_log=============================================================
        String tbName15[] = {"ExHelpItemF_log"};
        for(int i=0;i<tbName15.length;i++){
          tbNameNow = tbName15[i];
          //STATISTICS_BAK
          String mYear = "";
          String mMonth = "";
          String userId = "";
          String userName = "";
          String bankNo = "";
          String bankType = "";
          String bankName = "";
          int updateNum = 0;
          int deleteNum = 0;
          int downloadNum = 0;
          int loginNum = 0;

          //查詢符合條件的資料
          sqlCmd = "SELECT "
              + "A.DISP_ID,A.MUSER_ID,A.EXAM_ITEM,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName15[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName15[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName15[i]+"'"
                  + " AND ((M_YEAR = " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_MONTH BETWEEN " + sMonth + " AND 12)"
                  + " OR (M_YEAR = " + (Integer.parseInt(eYear) - 1911)
                  + " AND M_MONTH BETWEEN 1 AND " + eMonth + ")"
                  + " OR (M_YEAR > " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_YEAR < " + (Integer.parseInt(eYear) - 1911) + "))";
              if (debug) System.out.println("sqlCmd=" + sqlCmd);
              st3.addBatch(sqlCmd);
            }

            //取得資料
            String dbDispId = RptCG002W.getTrimString(rs.getString("DISP_ID"));
            String dbMuserId = RptCG002W.getTrimString(rs.getString("MUSER_ID"));
            String dbExamItem = RptCG002W.getTrimString(rs.getString("EXAM_ITEM"));

            String dbUserIdC = RptCG002W.getTrimString(rs.getString("USER_ID_C"));
            String dbUserNameC = RptCG002W.getTrimString(rs.getString("USER_NAME_C"));
            Timestamp dbUpdateTimestamp = rs.getTimestamp("UPDATE_DATE_C");
            String dbUpdateDateC = dbUpdateTimestamp.toString();
            dbUpdateDateC = dbUpdateDateC.substring(0, dbUpdateDateC.indexOf("."));
            String dbUpdateTypeC = RptCG002W.getTrimString(rs.getString("UPDATE_TYPE_C"));
            String dbUYear = RptCG002W.getTrimString(rs.getString("U_YEAR"));
            String dbUMonth = RptCG002W.getTrimString(rs.getString("U_MONTH"));
            String tmBankType =rs.getString("TM_BANK_TYPE");
            String tmTBankNo = rs.getString("TM_TBANK_NO");
            String tmBankName = RptCG002W.getTrimString((String)bankMap.get(tmTBankNo)," ");

            //新增_BAK的TABLE中
            sqlCmd2 = "INSERT INTO "+tbName15[i]+"_BAK("
                + "DISP_ID,MUSER_ID,EXAM_ITEM,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + "'" + dbDispId + "','" + dbMuserId + "',"
                + "'" + dbExamItem + "',"
                + "'" + dbUserIdC + "','" + dbUserNameC + "',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUpdateTypeC +"')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName15[i] + "'," + mYear + "," + mMonth
                    + ",'" + userId + "','" + userName + "','" + bankNo
                    + "','" + bankType + "','" + bankName + "'," + updateNum
                    + "," + deleteNum + "," + downloadNum + "," + loginNum + ")";
                if (debug) System.out.println("sqlCmd=" + sqlCmd);
                st3.addBatch(sqlCmd);

                updateNum = 0;
                deleteNum = 0;
                downloadNum = 0;
              }
            }

            mYear = dbUYear;
            mMonth = dbUMonth;
            userId = dbUserIdC;
            userName = dbUserNameC;
            bankNo = tmTBankNo;
            bankType = tmBankType;
            bankName = tmBankName;

            //統計資料
            if (dbUpdateTypeC.equals("U")) {
              updateNum++;
            }else if (dbUpdateTypeC.equals("D")) {
              deleteNum++;
            }else if (dbUpdateTypeC.equals("L")) {
              downloadNum++;
            }
          }
          int rows[] = st2.executeBatch();  //用rows.length就知道一共變更了幾筆資料

          if (!mYear.equals("")) {
            sqlCmd = "INSERT INTO STATISTICS_BAK("
                + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                + "DOWNLOAD_NUM,LOGIN_NUM) "
                + "VALUES('" + tbName15[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName15[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //ExReportF_log=============================================================
        String tbName16[] = {"ExReportF_log"};
        for(int i=0;i<tbName16.length;i++){
          tbNameNow = tbName16[i];
          //STATISTICS_BAK
          String mYear = "";
          String mMonth = "";
          String userId = "";
          String userName = "";
          String bankNo = "";
          String bankType = "";
          String bankName = "";
          int updateNum = 0;
          int deleteNum = 0;
          int downloadNum = 0;
          int loginNum = 0;

          //查詢符合條件的資料
          sqlCmd = "SELECT "
              + "A.REPORTNO,A.REPORT_IN_DATE,A.REPORT_EN_DATE,A.BASE_DATE,"
              + "A.DISP_ID,A.ORIGINUNT_ID,A.CH_TYPE,A.BANK_TYPE,A.TBANK_NO,"
              + "A.BANK_NO,A.REPORT_COME_DATE,A.REPORT_COME_DOCNO,"
              + "A.REPORT_RECEIVE_DOCNO,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName16[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName16[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName16[i]+"'"
                  + " AND ((M_YEAR = " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_MONTH BETWEEN " + sMonth + " AND 12)"
                  + " OR (M_YEAR = " + (Integer.parseInt(eYear) - 1911)
                  + " AND M_MONTH BETWEEN 1 AND " + eMonth + ")"
                  + " OR (M_YEAR > " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_YEAR < " + (Integer.parseInt(eYear) - 1911) + "))";
              if (debug) System.out.println("sqlCmd=" + sqlCmd);
              st3.addBatch(sqlCmd);
            }

            //取得資料
            String dbReportNo = RptCG002W.getTrimString(rs.getString("REPORTNO"));
            Timestamp dbReportInDateTimestamp = rs.getTimestamp("REPORT_IN_DATE");
            String dbReportInDate = "";
            if(dbReportInDateTimestamp != null){
              dbReportInDate = dbReportInDateTimestamp.toString();
              dbReportInDate = dbReportInDate.substring(0, dbReportInDate.indexOf("."));
            }
            Timestamp dbReportEnDateTimestamp = rs.getTimestamp("REPORT_EN_DATE");
            String dbReportEnDate = "";
            if(dbReportEnDateTimestamp != null){
              dbReportEnDate = dbReportEnDateTimestamp.toString();
              dbReportEnDate = dbReportEnDate.substring(0, dbReportEnDate.indexOf("."));
            }
            Timestamp dbBaseDateTimestamp = rs.getTimestamp("BASE_DATE");
            String dbBaseDate = "";
            if(dbBaseDateTimestamp != null){
              dbBaseDate = dbBaseDateTimestamp.toString();
              dbBaseDate = dbBaseDate.substring(0, dbBaseDate.indexOf("."));
            }
            String dbOriginuntId = RptCG002W.getTrimString(rs.getString("ORIGINUNT_ID"));
            String dbChType = RptCG002W.getTrimString(rs.getString("CH_TYPE"));
            String dbBankType = RptCG002W.getTrimString(rs.getString("BANK_TYPE"));
            String dbTbankNo = RptCG002W.getTrimString(rs.getString("TBANK_NO"));
            String dbBankNo = RptCG002W.getTrimString(rs.getString("BANK_NO"));
            Timestamp dbReportComeDateTimestamp = rs.getTimestamp("REPORT_COME_DATE");
            String dbReportComeDate = "";
            if(dbReportComeDateTimestamp != null){
              dbReportComeDate = dbReportComeDateTimestamp.toString();
              dbReportComeDate = dbReportComeDate.substring(0, dbReportComeDate.indexOf("."));
            }
            String dbReportComeDocNo = RptCG002W.getTrimString(rs.getString("REPORT_COME_DOCNO"));
            String dbReportReceiveDocNo = RptCG002W.getTrimString(rs.getString("REPORT_RECEIVE_DOCNO"));

            String dbUserIdC = RptCG002W.getTrimString(rs.getString("USER_ID_C"));
            String dbUserNameC = RptCG002W.getTrimString(rs.getString("USER_NAME_C"));
            Timestamp dbUpdateTimestamp = rs.getTimestamp("UPDATE_DATE_C");
            String dbUpdateDateC = dbUpdateTimestamp.toString();
            dbUpdateDateC = dbUpdateDateC.substring(0, dbUpdateDateC.indexOf("."));
            String dbUpdateTypeC = RptCG002W.getTrimString(rs.getString("UPDATE_TYPE_C"));
            String dbUYear = RptCG002W.getTrimString(rs.getString("U_YEAR"));
            String dbUMonth = RptCG002W.getTrimString(rs.getString("U_MONTH"));
            String tmBankType =rs.getString("TM_BANK_TYPE");
            String tmTBankNo = rs.getString("TM_TBANK_NO");
            String tmBankName = RptCG002W.getTrimString((String)bankMap.get(tmTBankNo)," ");

            //新增_BAK的TABLE中
            sqlCmd2 = "INSERT INTO "+tbName16[i]+"_BAK("
                + "REPORTNO,REPORT_IN_DATE,REPORT_EN_DATE,BASE_DATE,"
                + "DISP_ID,ORIGINUNT_ID,CH_TYPE,BANK_TYPE,TBANK_NO,"
                + "BANK_NO,REPORT_COME_DATE,REPORT_COME_DOCNO,"
                + "REPORT_RECEIVE_DOCNO,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + "'" + dbReportNo + "',"
                + "TO_DATE('" + dbReportInDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "TO_DATE('" + dbReportEnDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "TO_DATE('" + dbBaseDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbOriginuntId + "','" + dbChType + "',"
                + "'" + dbBankType + "','" + dbTbankNo + "',"
                + "'" + dbBankNo + "',"
                + "TO_DATE('" + dbReportComeDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbReportComeDocNo + "','" + dbReportReceiveDocNo + "',"
                + "'" + dbUserIdC + "','" + dbUserNameC + "',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUpdateTypeC +"')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName16[i] + "'," + mYear + "," + mMonth
                    + ",'" + userId + "','" + userName + "','" + bankNo
                    + "','" + bankType + "','" + bankName + "'," + updateNum
                    + "," + deleteNum + "," + downloadNum + "," + loginNum + ")";
                if (debug) System.out.println("sqlCmd=" + sqlCmd);
                st3.addBatch(sqlCmd);

                updateNum = 0;
                deleteNum = 0;
                downloadNum = 0;
              }
            }

            mYear = dbUYear;
            mMonth = dbUMonth;
            userId = dbUserIdC;
            userName = dbUserNameC;
            bankNo = tmTBankNo;
            bankType = tmBankType;
            bankName = tmBankName;

            //統計資料
            if (dbUpdateTypeC.equals("U")) {
              updateNum++;
            }else if (dbUpdateTypeC.equals("D")) {
              deleteNum++;
            }else if (dbUpdateTypeC.equals("L")) {
              downloadNum++;
            }
          }
          int rows[] = st2.executeBatch();  //用rows.length就知道一共變更了幾筆資料

          if (!mYear.equals("")) {
            sqlCmd = "INSERT INTO STATISTICS_BAK("
                + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                + "DOWNLOAD_NUM,LOGIN_NUM) "
                + "VALUES('" + tbName16[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName16[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //ExRtDocF_log=============================================================
        String tbName17[] = {"ExRtDocF_log"};
        for(int i=0;i<tbName17.length;i++){
          tbNameNow = tbName17[i];
          //STATISTICS_BAK
          String mYear = "";
          String mMonth = "";
          String userId = "";
          String userName = "";
          String bankNo = "";
          String bankType = "";
          String bankName = "";
          int updateNum = 0;
          int deleteNum = 0;
          int downloadNum = 0;
          int loginNum = 0;

          //查詢符合條件的資料
          sqlCmd = "SELECT "
              + "A.RT_DOCNO,A.RT_DATE,A.SN_DOCNO,A.RECEIVE_DOCNO,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName17[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName17[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName17[i]+"'"
                  + " AND ((M_YEAR = " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_MONTH BETWEEN " + sMonth + " AND 12)"
                  + " OR (M_YEAR = " + (Integer.parseInt(eYear) - 1911)
                  + " AND M_MONTH BETWEEN 1 AND " + eMonth + ")"
                  + " OR (M_YEAR > " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_YEAR < " + (Integer.parseInt(eYear) - 1911) + "))";
              if (debug) System.out.println("sqlCmd=" + sqlCmd);
              st3.addBatch(sqlCmd);
            }

            //取得資料
            String dbRtDocNo = RptCG002W.getTrimString(rs.getString("RT_DOCNO"));
            Timestamp dbRtDateTimestamp = rs.getTimestamp("RT_DATE");
            String dbRtDate = "";
            if(dbRtDateTimestamp != null){
              dbRtDate = dbRtDateTimestamp.toString();
              dbRtDate = dbRtDate.substring(0, dbRtDate.indexOf("."));
            }
            String dbSnDocNo = RptCG002W.getTrimString(rs.getString("SN_DOCNO"));
            String dbReceiveDocNo = RptCG002W.getTrimString(rs.getString("RECEIVE_DOCNO"));

            String dbUserIdC = RptCG002W.getTrimString(rs.getString("USER_ID_C"));
            String dbUserNameC = RptCG002W.getTrimString(rs.getString("USER_NAME_C"));
            Timestamp dbUpdateTimestamp = rs.getTimestamp("UPDATE_DATE_C");
            String dbUpdateDateC = dbUpdateTimestamp.toString();
            dbUpdateDateC = dbUpdateDateC.substring(0, dbUpdateDateC.indexOf("."));
            String dbUpdateTypeC = RptCG002W.getTrimString(rs.getString("UPDATE_TYPE_C"));
            String dbUYear = RptCG002W.getTrimString(rs.getString("U_YEAR"));
            String dbUMonth = RptCG002W.getTrimString(rs.getString("U_MONTH"));
            String tmBankType =rs.getString("TM_BANK_TYPE");
            String tmTBankNo = rs.getString("TM_TBANK_NO");
            String tmBankName = RptCG002W.getTrimString((String)bankMap.get(tmTBankNo)," ");

            //新增_BAK的TABLE中
            sqlCmd2 = "INSERT INTO "+tbName17[i]+"_BAK("
                + "RT_DOCNO,RT_DATE,SN_DOCNO,RECEIVE_DOCNO,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + "'" + dbRtDocNo + "',"
                + "TO_DATE('" + dbRtDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbSnDocNo + "','" + dbReceiveDocNo + "',"
                + "'" + dbUserIdC + "','" + dbUserNameC + "',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUpdateTypeC +"')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName17[i] + "'," + mYear + "," + mMonth
                    + ",'" + userId + "','" + userName + "','" + bankNo
                    + "','" + bankType + "','" + bankName + "'," + updateNum
                    + "," + deleteNum + "," + downloadNum + "," + loginNum + ")";
                if (debug) System.out.println("sqlCmd=" + sqlCmd);
                st3.addBatch(sqlCmd);

                updateNum = 0;
                deleteNum = 0;
                downloadNum = 0;
              }
            }

            mYear = dbUYear;
            mMonth = dbUMonth;
            userId = dbUserIdC;
            userName = dbUserNameC;
            bankNo = tmTBankNo;
            bankType = tmBankType;
            bankName = tmBankName;

            //統計資料
            if (dbUpdateTypeC.equals("U")) {
              updateNum++;
            }else if (dbUpdateTypeC.equals("D")) {
              deleteNum++;
            }else if (dbUpdateTypeC.equals("L")) {
              downloadNum++;
            }
          }
          int rows[] = st2.executeBatch();  //用rows.length就知道一共變更了幾筆資料

          if (!mYear.equals("")) {
            sqlCmd = "INSERT INTO STATISTICS_BAK("
                + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                + "DOWNLOAD_NUM,LOGIN_NUM) "
                + "VALUES('" + tbName17[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName17[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //ExScheduleF_log=============================================================
        String tbName18[] = {"ExScheduleF_log"};
        for(int i=0;i<tbName18.length;i++){
          tbNameNow = tbName18[i];
          //STATISTICS_BAK
          String mYear = "";
          String mMonth = "";
          String userId = "";
          String userName = "";
          String bankNo = "";
          String bankType = "";
          String bankName = "";
          int updateNum = 0;
          int deleteNum = 0;
          int downloadNum = 0;
          int loginNum = 0;

          //查詢符合條件的資料
          sqlCmd = "SELECT "
              + "A.DISP_ID,A.BASE_DATE,A.GO_DATE,A.GO_AMPM,A.BK_DATE,"
              + "A.BK_AMPM,A.WARE_DATE,A.ST_DATE,A.ST_AMPM,A.EN_DATE,"
              + "A.EN_AMPM,A.REPORT_DATE,A.WORKDAYS,A.WORKMDAYS,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName18[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName18[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName18[i]+"'"
                  + " AND ((M_YEAR = " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_MONTH BETWEEN " + sMonth + " AND 12)"
                  + " OR (M_YEAR = " + (Integer.parseInt(eYear) - 1911)
                  + " AND M_MONTH BETWEEN 1 AND " + eMonth + ")"
                  + " OR (M_YEAR > " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_YEAR < " + (Integer.parseInt(eYear) - 1911) + "))";
              if (debug) System.out.println("sqlCmd=" + sqlCmd);
              st3.addBatch(sqlCmd);
            }

            //取得資料
            String dbDispId = RptCG002W.getTrimString(rs.getString("DISP_ID"));
            Timestamp dbBaseDateTimestamp = rs.getTimestamp("BASE_DATE");
            String dbBaseDate = "";
            if(dbBaseDateTimestamp != null){
              dbBaseDate = dbBaseDateTimestamp.toString();
              dbBaseDate = dbBaseDate.substring(0, dbBaseDate.indexOf("."));
            }
            Timestamp dbGoDateTimestamp = rs.getTimestamp("GO_DATE");
            String dbGoDate = "";
            if(dbGoDateTimestamp != null){
              dbGoDate = dbGoDateTimestamp.toString();
              dbGoDate = dbGoDate.substring(0, dbGoDate.indexOf("."));
            }
            String dbGoAmPm = RptCG002W.getTrimString(rs.getString("GO_AMPM"));
            Timestamp dbBkDateTimestamp = rs.getTimestamp("BK_DATE");
            String dbBkDate = "";
            if(dbBkDateTimestamp != null){
              dbBkDate = dbBkDateTimestamp.toString();
              dbBkDate = dbBkDate.substring(0, dbBkDate.indexOf("."));
            }
            String dbBkAmPm = RptCG002W.getTrimString(rs.getString("BK_AMPM"));
            Timestamp dbWareDateTimestamp = rs.getTimestamp("WARE_DATE");
            String dbWareDate = "";
            if(dbWareDateTimestamp != null){
              dbWareDate = dbWareDateTimestamp.toString();
              dbWareDate = dbWareDate.substring(0, dbWareDate.indexOf("."));
            }
            Timestamp dbStDateTimestamp = rs.getTimestamp("ST_DATE");
            String dbStDate = "";
            if(dbStDateTimestamp != null){
              dbStDate = dbStDateTimestamp.toString();
              dbStDate = dbStDate.substring(0, dbStDate.indexOf("."));
            }
            String dbStAmPm = RptCG002W.getTrimString(rs.getString("ST_AMPM"));
            Timestamp dbEnDateTimestamp = rs.getTimestamp("EN_DATE");
            String dbEnDate = "";
            if(dbEnDateTimestamp != null){
              dbEnDate = dbEnDateTimestamp.toString();
              dbEnDate = dbEnDate.substring(0, dbEnDate.indexOf("."));
            }
            String dbEnAmPm = RptCG002W.getTrimString(rs.getString("EN_AMPM"));
            Timestamp dbReportDateTimestamp = rs.getTimestamp("REPORT_DATE");
            String dbReportDate = "";
            if(dbReportDateTimestamp != null){
              dbReportDate = dbReportDateTimestamp.toString();
              dbReportDate = dbReportDate.substring(0, dbReportDate.indexOf("."));
            }
            String dbWorkdays = RptCG002W.getTrimString(rs.getString("WORKDAYS"));
            String dbWorkmdays = RptCG002W.getTrimString(rs.getString("WORKMDAYS"));

            String dbUserIdC = RptCG002W.getTrimString(rs.getString("USER_ID_C"));
            String dbUserNameC = RptCG002W.getTrimString(rs.getString("USER_NAME_C"));
            Timestamp dbUpdateTimestamp = rs.getTimestamp("UPDATE_DATE_C");
            String dbUpdateDateC = dbUpdateTimestamp.toString();
            dbUpdateDateC = dbUpdateDateC.substring(0, dbUpdateDateC.indexOf("."));
            String dbUpdateTypeC = RptCG002W.getTrimString(rs.getString("UPDATE_TYPE_C"));
            String dbUYear = RptCG002W.getTrimString(rs.getString("U_YEAR"));
            String dbUMonth = RptCG002W.getTrimString(rs.getString("U_MONTH"));
            String tmBankType =rs.getString("TM_BANK_TYPE");
            String tmTBankNo = rs.getString("TM_TBANK_NO");
            String tmBankName = RptCG002W.getTrimString((String)bankMap.get(tmTBankNo)," ");

            //新增_BAK的TABLE中
            sqlCmd2 = "INSERT INTO "+tbName18[i]+"_BAK("
                + "DISP_ID,BASE_DATE,GO_DATE,GO_AMPM,BK_DATE,"
                + "BK_AMPM,WARE_DATE,ST_DATE,ST_AMPM,EN_DATE,"
                + "EN_AMPM,REPORT_DATE,WORKDAYS,WORKMDAYS,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + "'" + dbDispId + "',"
                + "TO_DATE('" + dbBaseDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "TO_DATE('" + dbGoDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbGoAmPm + "',"
                + "TO_DATE('" + dbBkDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbBkAmPm + "',"
                + "TO_DATE('" + dbWareDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "TO_DATE('" + dbStDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbStAmPm + "',"
                + "TO_DATE('" + dbEnDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbEnAmPm + "',"
                + "TO_DATE('" + dbReportDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + dbWorkdays + "," + dbWorkmdays + ","
                + "'" + dbUserIdC + "','" + dbUserNameC + "',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUpdateTypeC +"')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName18[i] + "'," + mYear + "," + mMonth
                    + ",'" + userId + "','" + userName + "','" + bankNo
                    + "','" + bankType + "','" + bankName + "'," + updateNum
                    + "," + deleteNum + "," + downloadNum + "," + loginNum + ")";
                if (debug) System.out.println("sqlCmd=" + sqlCmd);
                st3.addBatch(sqlCmd);

                updateNum = 0;
                deleteNum = 0;
                downloadNum = 0;
              }
            }

            mYear = dbUYear;
            mMonth = dbUMonth;
            userId = dbUserIdC;
            userName = dbUserNameC;
            bankNo = tmTBankNo;
            bankType = tmBankType;
            bankName = tmBankName;

            //統計資料
            if (dbUpdateTypeC.equals("U")) {
              updateNum++;
            }else if (dbUpdateTypeC.equals("D")) {
              deleteNum++;
            }else if (dbUpdateTypeC.equals("L")) {
              downloadNum++;
            }
          }
          int rows[] = st2.executeBatch();  //用rows.length就知道一共變更了幾筆資料

          if (!mYear.equals("")) {
            sqlCmd = "INSERT INTO STATISTICS_BAK("
                + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                + "DOWNLOAD_NUM,LOGIN_NUM) "
                + "VALUES('" + tbName18[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName18[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //ExSnDocF_log=============================================================
        String tbName18_2[] = {"ExSnDocF_log"};
        for(int i=0;i<tbName18_2.length;i++){
          tbNameNow = tbName18_2[i];
          //STATISTICS_BAK
          String mYear = "";
          String mMonth = "";
          String userId = "";
          String userName = "";
          String bankNo = "";
          String bankType = "";
          String bankName = "";
          int updateNum = 0;
          int deleteNum = 0;
          int downloadNum = 0;
          int loginNum = 0;

          //查詢符合條件的資料
          sqlCmd = "SELECT "
              + "A.SN_DOCNO,A.REPORTNO,A.SN_DATE,A.DOCTYPE,"
              + "A.DOCTYPE_CNT,A.LIMITDATE,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName18_2[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName18_2[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName18_2[i]+"'"
                  + " AND ((M_YEAR = " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_MONTH BETWEEN " + sMonth + " AND 12)"
                  + " OR (M_YEAR = " + (Integer.parseInt(eYear) - 1911)
                  + " AND M_MONTH BETWEEN 1 AND " + eMonth + ")"
                  + " OR (M_YEAR > " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_YEAR < " + (Integer.parseInt(eYear) - 1911) + "))";
              if (debug) System.out.println("sqlCmd=" + sqlCmd);
              st3.addBatch(sqlCmd);
            }

            //取得資料
            String dbSnDocNo = RptCG002W.getTrimString(rs.getString("SN_DOCNO"));
            String dbReportNo = RptCG002W.getTrimString(rs.getString("REPORTNO"));
            Timestamp dbSnDateTimestamp = rs.getTimestamp("SN_DATE");
            String dbSnDate = "";
            if(dbSnDateTimestamp != null){
              dbSnDate = dbSnDateTimestamp.toString();
              dbSnDate = dbSnDate.substring(0, dbSnDate.indexOf("."));
            }
            String dbDocType = RptCG002W.getTrimString(rs.getString("DOCTYPE"));
            String dbDocTypeCnt = RptCG002W.getTrimString(rs.getString("DOCTYPE_CNT"));
            Timestamp dbLimitDateTimestamp = rs.getTimestamp("LIMITDATE");
            String dbLimitDate = "";
            if(dbLimitDateTimestamp != null){
              dbLimitDate = dbLimitDateTimestamp.toString();
              dbLimitDate = dbLimitDate.substring(0, dbLimitDate.indexOf("."));
            }

            String dbUserIdC = RptCG002W.getTrimString(rs.getString("USER_ID_C"));
            String dbUserNameC = RptCG002W.getTrimString(rs.getString("USER_NAME_C"));
            Timestamp dbUpdateTimestamp = rs.getTimestamp("UPDATE_DATE_C");
            String dbUpdateDateC = dbUpdateTimestamp.toString();
            dbUpdateDateC = dbUpdateDateC.substring(0, dbUpdateDateC.indexOf("."));
            String dbUpdateTypeC = RptCG002W.getTrimString(rs.getString("UPDATE_TYPE_C"));
            String dbUYear = RptCG002W.getTrimString(rs.getString("U_YEAR"));
            String dbUMonth = RptCG002W.getTrimString(rs.getString("U_MONTH"));
            String tmBankType =rs.getString("TM_BANK_TYPE");
            String tmTBankNo = rs.getString("TM_TBANK_NO");
            String tmBankName = RptCG002W.getTrimString((String)bankMap.get(tmTBankNo)," ");

            //新增_BAK的TABLE中
            sqlCmd2 = "INSERT INTO "+tbName18_2[i]+"_BAK("
                + "SN_DOCNO,REPORTNO,SN_DATE,DOCTYPE,DOCTYPE_CNT,LIMITDATE,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + "'" + dbSnDocNo + "','" + dbReportNo + "',"
                + "TO_DATE('" + dbSnDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbDocType + "','" + dbDocTypeCnt + "',"
                + "TO_DATE('" + dbLimitDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUserIdC + "','" + dbUserNameC + "',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUpdateTypeC +"')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName18_2[i] + "'," + mYear + "," + mMonth
                    + ",'" + userId + "','" + userName + "','" + bankNo
                    + "','" + bankType + "','" + bankName + "'," + updateNum
                    + "," + deleteNum + "," + downloadNum + "," + loginNum + ")";
                if (debug) System.out.println("sqlCmd=" + sqlCmd);
                st3.addBatch(sqlCmd);

                updateNum = 0;
                deleteNum = 0;
                downloadNum = 0;
              }
            }

            mYear = dbUYear;
            mMonth = dbUMonth;
            userId = dbUserIdC;
            userName = dbUserNameC;
            bankNo = tmTBankNo;
            bankType = tmBankType;
            bankName = tmBankName;

            //統計資料
            if (dbUpdateTypeC.equals("U")) {
              updateNum++;
            }else if (dbUpdateTypeC.equals("D")) {
              deleteNum++;
            }else if (dbUpdateTypeC.equals("L")) {
              downloadNum++;
            }
          }
          int rows[] = st2.executeBatch();  //用rows.length就知道一共變更了幾筆資料

          if (!mYear.equals("")) {
            sqlCmd = "INSERT INTO STATISTICS_BAK("
                + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                + "DOWNLOAD_NUM,LOGIN_NUM) "
                + "VALUES('" + tbName18_2[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName18_2[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //ExWarningF_log=============================================================
        String tbName19[] = {"ExWarningF_log"};
        for(int i=0;i<tbName19.length;i++){
          tbNameNow = tbName19[i];
          //STATISTICS_BAK
          String mYear = "";
          String mMonth = "";
          String userId = "";
          String userName = "";
          String bankNo = "";
          String bankType = "";
          String bankName = "";
          int updateNum = 0;
          int deleteNum = 0;
          int downloadNum = 0;
          int loginNum = 0;

          //查詢符合條件的資料
          sqlCmd = "SELECT "
              + "A.SERIAL,A.BANK_NO,A.EVENTDATE,A.RT_DOCNO,RT_DATE,"
              + "A.ITEM_ID,A.TRACK,A.SUMMARY,A.REMARK,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName19[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName19[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName19[i]+"'"
                  + " AND ((M_YEAR = " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_MONTH BETWEEN " + sMonth + " AND 12)"
                  + " OR (M_YEAR = " + (Integer.parseInt(eYear) - 1911)
                  + " AND M_MONTH BETWEEN 1 AND " + eMonth + ")"
                  + " OR (M_YEAR > " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_YEAR < " + (Integer.parseInt(eYear) - 1911) + "))";
              if (debug) System.out.println("sqlCmd=" + sqlCmd);
              st3.addBatch(sqlCmd);
            }

            //取得資料
            String dbSerial = RptCG002W.getTrimString(rs.getString("SERIAL"));
            String dbBankNo = RptCG002W.getTrimString(rs.getString("BANK_NO"));
            Timestamp dbEventDateTimestamp = rs.getTimestamp("EVENTDATE");
            String dbEventDate = "";
            if(dbEventDateTimestamp != null){
              dbEventDate = dbEventDateTimestamp.toString();
              dbEventDate = dbEventDate.substring(0, dbEventDate.indexOf("."));
            }
            String dbRtDocNo = RptCG002W.getTrimString(rs.getString("RT_DOCNO"));
            Timestamp dbRtDateTimestamp = rs.getTimestamp("RT_DATE");
            String dbRtDate = "";
            if(dbRtDateTimestamp != null){
              dbRtDate = dbRtDateTimestamp.toString();
              dbRtDate = dbRtDate.substring(0, dbRtDate.indexOf("."));
            }
            String dbItemId = RptCG002W.getTrimString(rs.getString("ITEM_ID"));
            String dbTrack = RptCG002W.getTrimString(rs.getString("TRACK"));
            String dbSummary = RptCG002W.getTrimString(rs.getString("SUMMARY"));
            String dbRemark = RptCG002W.getTrimString(rs.getString("REMARK"));

            String dbUserIdC = RptCG002W.getTrimString(rs.getString("USER_ID_C"));
            String dbUserNameC = RptCG002W.getTrimString(rs.getString("USER_NAME_C"));
            Timestamp dbUpdateTimestamp = rs.getTimestamp("UPDATE_DATE_C");
            String dbUpdateDateC = dbUpdateTimestamp.toString();
            dbUpdateDateC = dbUpdateDateC.substring(0, dbUpdateDateC.indexOf("."));
            String dbUpdateTypeC = RptCG002W.getTrimString(rs.getString("UPDATE_TYPE_C"));
            String dbUYear = RptCG002W.getTrimString(rs.getString("U_YEAR"));
            String dbUMonth = RptCG002W.getTrimString(rs.getString("U_MONTH"));
            String tmBankType =rs.getString("TM_BANK_TYPE");
            String tmTBankNo = rs.getString("TM_TBANK_NO");
            String tmBankName = RptCG002W.getTrimString((String)bankMap.get(tmTBankNo)," ");

            //新增_BAK的TABLE中
            sqlCmd2 = "INSERT INTO "+tbName19[i]+"_BAK("
                + "SERIAL,BANK_NO,EVENTDATE,RT_DOCNO,RT_DATE,"
                + "ITEM_ID,TRACK,SUMMARY,REMARK,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + "'" + dbSerial + "','" + dbBankNo + "',"
                + "TO_DATE('" + dbEventDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbRtDocNo + "',"
                + "TO_DATE('" + dbRtDate + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbItemId + "','" + dbTrack + "',"
                + "'" + dbSummary + "','" + dbRemark + "',"
                + "'" + dbUserIdC + "','" + dbUserNameC + "',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUpdateTypeC +"')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName19[i] + "'," + mYear + "," + mMonth
                    + ",'" + userId + "','" + userName + "','" + bankNo
                    + "','" + bankType + "','" + bankName + "'," + updateNum
                    + "," + deleteNum + "," + downloadNum + "," + loginNum + ")";
                if (debug) System.out.println("sqlCmd=" + sqlCmd);
                st3.addBatch(sqlCmd);

                updateNum = 0;
                deleteNum = 0;
                downloadNum = 0;
              }
            }

            mYear = dbUYear;
            mMonth = dbUMonth;
            userId = dbUserIdC;
            userName = dbUserNameC;
            bankNo = tmTBankNo;
            bankType = tmBankType;
            bankName = tmBankName;

            //統計資料
            if (dbUpdateTypeC.equals("U")) {
              updateNum++;
            }else if (dbUpdateTypeC.equals("D")) {
              deleteNum++;
            }else if (dbUpdateTypeC.equals("L")) {
              downloadNum++;
            }
          }
          int rows[] = st2.executeBatch();  //用rows.length就知道一共變更了幾筆資料

          if (!mYear.equals("")) {
            sqlCmd = "INSERT INTO STATISTICS_BAK("
                + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                + "DOWNLOAD_NUM,LOGIN_NUM) "
                + "VALUES('" + tbName19[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName19[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        //F01_log=============================================================
        String tbName20[] = {"F01_log"};
        for(int i=0;i<tbName20.length;i++){
          tbNameNow = tbName20[i];
          //STATISTICS_BAK
          String mYear = "";
          String mMonth = "";
          String userId = "";
          String userName = "";
          String bankNo = "";
          String bankType = "";
          String bankName = "";
          int updateNum = 0;
          int deleteNum = 0;
          int downloadNum = 0;
          int loginNum = 0;

          //查詢符合條件的資料
          sqlCmd = "SELECT "
              + "A.M_YEAR,A.M_MONTH,A.BANK_CODE,A.DEP_TYPE,A.ACCT_TYPE,"
              + "A.ACCT_CNT_TM,A.BAL_LM,A.DEP_TM,A.WTD_TM,A.BAL_TM,"
              + "A.USER_ID_C,A.USER_NAME_C,A.UPDATE_DATE_C,A.UPDATE_TYPE_C,"
              + "TO_CHAR(A.UPDATE_DATE_C,'yyyy') - 1911 AS U_YEAR,"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm') AS U_MONTH,"
              + "W.BANK_TYPE AS TM_BANK_TYPE,W.TBANK_NO AS TM_TBANK_NO "
              + "FROM "+tbName20[i]+" A ,WTT01 W "
              + "WHERE A.USER_ID_C = W.MUSER_ID "
              + "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') "
              + "BETWEEN '" + (sYear + sMonth) + "' "
              + "AND '" + (eYear + eMonth) + "' "
              + "ORDER BY TO_CHAR(A.UPDATE_DATE_C,'yyyy'),"
              + "TO_CHAR(A.UPDATE_DATE_C,'mm'),"
              + "A.USER_ID_C,W.TBANK_NO";
          if (debug) System.out.println("sqlCmd=" + sqlCmd);
          rs = st.executeQuery(sqlCmd);

          //備份
          while (rs.next()) {
            //刪除原先備份的檔案
            if (mYear.equals("")) {
              sqlCmd2 = "DELETE "+tbName20[i]+"_BAK "
                  + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymm') "
                  + "BETWEEN '" + (sYear + sMonth) + "' "
                  + "AND '" + (eYear + eMonth) + "' ";
              if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
              st2.addBatch(sqlCmd2);

              sqlCmd = "DELETE STATISTICS_BAK "
                  + "WHERE TB_NAME = '"+tbName20[i]+"'"
                  + " AND ((M_YEAR = " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_MONTH BETWEEN " + sMonth + " AND 12)"
                  + " OR (M_YEAR = " + (Integer.parseInt(eYear) - 1911)
                  + " AND M_MONTH BETWEEN 1 AND " + eMonth + ")"
                  + " OR (M_YEAR > " + (Integer.parseInt(sYear) - 1911)
                  + " AND M_YEAR < " + (Integer.parseInt(eYear) - 1911) + "))";
              if (debug) System.out.println("sqlCmd=" + sqlCmd);
              st3.addBatch(sqlCmd);
            }

            //取得資料
            String dbMYear = RptCG002W.getTrimString(rs.getString("M_YEAR"));
            String dbMMonth = RptCG002W.getTrimString(rs.getString("M_MONTH"));
            String dbBankCode = RptCG002W.getTrimString(rs.getString("BANK_CODE"));
            String dbDepType = RptCG002W.getTrimString(rs.getString("DEP_TYPE"));
            String dbAcctType = RptCG002W.getTrimString(rs.getString("ACCT_TYPE"));
            String dbAcctCntTm = RptCG002W.getTrimString(rs.getString("ACCT_CNT_TM"));
            String dbBalLm = RptCG002W.getTrimString(rs.getString("BAL_LM"));
            String dbDepTm = RptCG002W.getTrimString(rs.getString("DEP_TM"));
            String dbWtdTm = RptCG002W.getTrimString(rs.getString("WTD_TM"));
            String dbBalTm = RptCG002W.getTrimString(rs.getString("BAL_TM"));

            String dbUserIdC = RptCG002W.getTrimString(rs.getString("USER_ID_C"));
            String dbUserNameC = RptCG002W.getTrimString(rs.getString("USER_NAME_C"));
            Timestamp dbUpdateTimestamp = rs.getTimestamp("UPDATE_DATE_C");
            String dbUpdateDateC = dbUpdateTimestamp.toString();
            dbUpdateDateC = dbUpdateDateC.substring(0, dbUpdateDateC.indexOf("."));
            String dbUpdateTypeC = RptCG002W.getTrimString(rs.getString("UPDATE_TYPE_C"));
            String dbUYear = RptCG002W.getTrimString(rs.getString("U_YEAR"));
            String dbUMonth = RptCG002W.getTrimString(rs.getString("U_MONTH"));
            String tmBankType =rs.getString("TM_BANK_TYPE");
            String tmTBankNo = rs.getString("TM_TBANK_NO");
            String tmBankName = RptCG002W.getTrimString((String)bankMap.get(tmTBankNo)," ");

            //新增_BAK的TABLE中
            sqlCmd2 = "INSERT INTO "+tbName20[i]+"_BAK("
                + "M_YEAR,M_MONTH,BANK_CODE,DEP_TYPE,ACCT_TYPE,"
                + "ACCT_CNT_TM,BAL_LM,DEP_TM,WTD_TM,BAL_TM,"
                + "USER_ID_C,USER_NAME_C,UPDATE_DATE_C,UPDATE_TYPE_C) "
                + "VALUES("
                + dbMYear + "," + dbMMonth + ","
                + "'" + dbBankCode + "','" + dbDepType + "',"
                + "'" + dbAcctType + "'," + dbAcctCntTm + ","
                + dbBalLm + "," + dbDepTm + ","
                + dbWtdTm + "," + dbBalTm + ","
                + "'" + dbUserIdC + "','" + dbUserNameC + "',"
                + "TO_DATE('" + dbUpdateDateC + "','yyyy-mm-dd hh24:mi:ss'),"
                + "'" + dbUpdateTypeC +"')";
             if (debug) System.out.println("sqlCmd2=" + sqlCmd2);
             st2.addBatch(sqlCmd2);

            //新增STATISTICS_BAK
            if (! (mYear + mMonth + userId + bankNo).equals(dbUYear + dbUMonth +
                dbUserIdC + tmTBankNo)) {
              if (! (mYear + mMonth + userId).equals("")) {
                sqlCmd = "INSERT INTO STATISTICS_BAK("
                    + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                    + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                    + "DOWNLOAD_NUM,LOGIN_NUM) "
                    + "VALUES('" + tbName20[i] + "'," + mYear + "," + mMonth
                    + ",'" + userId + "','" + userName + "','" + bankNo
                    + "','" + bankType + "','" + bankName + "'," + updateNum
                    + "," + deleteNum + "," + downloadNum + "," + loginNum + ")";
                if (debug) System.out.println("sqlCmd=" + sqlCmd);
                st3.addBatch(sqlCmd);

                updateNum = 0;
                deleteNum = 0;
                downloadNum = 0;
              }
            }

            mYear = dbUYear;
            mMonth = dbUMonth;
            userId = dbUserIdC;
            userName = dbUserNameC;
            bankNo = tmTBankNo;
            bankType = tmBankType;
            bankName = tmBankName;

            //統計資料
            if (dbUpdateTypeC.equals("U")) {
              updateNum++;
            }else if (dbUpdateTypeC.equals("D")) {
              deleteNum++;
            }else if (dbUpdateTypeC.equals("L")) {
              downloadNum++;
            }
          }
          int rows[] = st2.executeBatch();  //用rows.length就知道一共變更了幾筆資料

          if (!mYear.equals("")) {
            sqlCmd = "INSERT INTO STATISTICS_BAK("
                + "TB_NAME,M_YEAR,M_MONTH,USER_ID,USER_NAME,BANK_NO,"
                + "BANK_TYPE,BANK_NAME,UPDATE_NUM,DELETE_NUM,"
                + "DOWNLOAD_NUM,LOGIN_NUM) "
                + "VALUES('" + tbName20[i] + "'," + mYear + "," + mMonth
                + ",'" +userId + "','" + userName + "','" + bankNo + "','"
                + bankType + "','" + bankName + "'," + updateNum + ","
                + deleteNum + "," + downloadNum + "," + loginNum + ")";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            st3.addBatch(sqlCmd);
          }
          st3.executeBatch();

          //刪除
          if (rows != null && rows.length > 0) {
            sqlCmd = "DELETE " + tbName20[i] + " "
                + "WHERE TO_CHAR(UPDATE_DATE_C,'yyyymmd') "
                + "BETWEEN '" + (sYear + sMonth) + "' "
                + "AND '" + (eYear + eMonth) + "' ";
            if (debug) System.out.println("sqlCmd=" + sqlCmd);
            if (isDelete) st4.executeUpdate(sqlCmd);
          }
        }

        conn1.commit();
        conn2.commit();
        tbNameNow = "";
      }catch (Exception e) {
        System.out.println("tbNameNow = " + tbNameNow);
        System.out.println("//RptCG002W_E() Have Error.....");
        e.printStackTrace();
        System.out.println(e.toString());
        System.out.println("//-------------------------------------");
      }
      return tbNameNow;
    }

}
