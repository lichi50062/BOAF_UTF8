package com.tradevan.util.manager;

import java.sql.*;
import java.util.*;
import com.tradevan.util.bean.DetailRepBean;
import com.tradevan.util.Utility;
public class ReportManager
    extends BaseManager
    implements Manager {
  private ReportManager() {
  }

  public static Object getManager() {
    return new ReportManager();
  }

  /**
   * prepareRepList
   *
   * @param model Object
   * @param con Connection
   * @throws Exception
   * @return Object
   * @todo Implement this com.tradevan.util.manager.Manager method
   */
  public Object prepareRepList(Object model, Connection con) throws Exception {
    Map h = (Map)model;
    List returnList=null;
    String reportId = (String)h.get("reportId");
    try {
      if (reportId.equals("FR007WX")) {
        returnList = goFR007WX(h, con);
      }
      return returnList;
    }
    catch (Exception ex) {
      return null;
    }
  }
  private List goFR007WX(Map h, Connection con){
    List returnList = new ArrayList();
    CallableStatement cstmt = null;
    PreparedStatement pstmt = null;
    ResultSet rs = null;
    StringBuffer sql = new StringBuffer();
    try {
      cstmt = con.prepareCall("{call RptFR007WX(?,?,?,?,?,?,?)}");
      cstmt.setString(1, (String) h.get("reportId"));
      cstmt.setString(2, (String) h.get("userId"));
      cstmt.setString(3, (String) h.get("bank_type"));
      cstmt.setString(4, (String) h.get("spec"));
      cstmt.setString(5, (String) h.get("S_YEAR"));
      cstmt.setString(6, (String) h.get("S_MONTH"));
      cstmt.registerOutParameter(7, java.sql.Types.VARCHAR);
      cstmt.execute();
      int finorNot = cstmt.getInt(7);
      System.out.println(" finorNot= " + finorNot);
      if (finorNot == 0) {
        throw new SQLException();
      }
      else {
        sql.append("select seqno,h1,c1,c2,c3,c21,n1,n2,n3,n4,n5,n6,n7, ")
            .append("      n8,n9,n10,n11,n12,n13,n14,n15,n16 ")
            .append(" from detailRep ")
            .append("where prgno=? ")
            .append("  and userid=? ")
            .append("order by seqno,h1,c91 ");
        pstmt = con.prepareStatement(sql.toString());
        pstmt.setString(1, (String) h.get("reportId"));
        pstmt.setString(2, (String) h.get("userId"));
        rs = pstmt.executeQuery();
        DetailRepBean bean = null;
        while (rs.next()) {
          bean = new DetailRepBean();
          bean.setC1(Utility.getTrimString(rs.getString("c1")));
          bean.setC2(Utility.getTrimString(rs.getString("c2")));
          bean.setC3(Utility.getTrimString(rs.getString("c3")));
          bean.setC21(Utility.getTrimString(rs.getString("c21")));
          bean.setN1(Utility.getTrimString(rs.getString("n1")));
          bean.setN2(Utility.getTrimString(rs.getString("n2")));
          bean.setN3(Utility.getTrimString(rs.getString("n3")));
          bean.setN4(Utility.getTrimString(rs.getString("n4")));
          bean.setN5(Utility.getTrimString(rs.getString("n5")));
          bean.setN6(Utility.getTrimString(rs.getString("n6")));
          bean.setN7(Utility.getTrimString(rs.getString("n7")));
          bean.setN8(Utility.getTrimString(rs.getString("n8")));
          bean.setN9(Utility.getTrimString(rs.getString("n9")));
          bean.setN10(Utility.getTrimString(rs.getString("n10")));
          bean.setN11(Utility.getTrimString(rs.getString("n11")));
          bean.setN12(Utility.getTrimString(rs.getString("n12")));
          bean.setN13(Utility.getTrimString(rs.getString("n13")));
          bean.setN14(Utility.getTrimString(rs.getString("n14")));
          bean.setN15(Utility.getTrimString(rs.getString("n15")));
          bean.setN16(Utility.getTrimString(rs.getString("n16")));
          bean.setSeqNo(Utility.getTrimString(rs.getString("seqno")));
          bean.setH1(Utility.getTrimString(rs.getString("h1")));
          returnList.add(bean);
          bean = null;
        }
      }
    }
    catch (SQLException ex) {
     System.out.println("SQLException : "+ex.getMessage());
    }finally{
      try {
        if (rs != null) {
          rs.close();
        }
        if (pstmt != null) {
          pstmt.close();
        }
        if (cstmt != null) {
          cstmt.close();
        }
      }
      catch (SQLException ex1) {
      }
    }
    return returnList;
  }
}
