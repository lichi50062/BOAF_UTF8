<%
// 99.07.29 first designed by 2660
// 99.08.20 fix 100年度區分縣市別才要加m_year條件 by 2295
//100.01.27 fix sqlInjection by 2295
%>

<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%@ page import="java.io.*" %>

<%
  System.out.println("FX013W start ..........");

  RequestDispatcher rd = null;
  String actMsg = "";
  String alertMsg = "";
  String webURL = "";
  boolean doProcess = false;

  // 取得 session 資料，取得成功時，才繼續往下執行 ===================================================================
  if ( session.getAttribute("muser_id") == null ) { // session timeout
    System.out.println("FX013W login timeout");
    rd = application.getRequestDispatcher( "/pages/reLogin.jsp?url=LoginError.jsp?timeout=true" );
    try {
      rd.forward(request,response);
    } catch(Exception e) {
      System.out.println("forward Error:"+e+e.getMessage());
    }
  } else {
    doProcess = true;
  }

  if ( doProcess ) { // 若 muser_id 資料時，表示登入成功 =============================================================
    String act = ( request.getParameter("act") == null ) ? "" : (String)request.getParameter("act");
    String bank_type = ( request.getParameter("bank_type") == null ) ? "" : (String)request.getParameter("bank_type");
    String list_type = ( request.getParameter("list_type") == null ) ? "" : (String)request.getParameter("list_type");
    System.out.println("act="+act);
    System.out.println("list_type="+list_type);	

    // 登入者資訊
    String lguser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");
    String lguser_name = ( session.getAttribute("muser_name")==null ) ? "" : (String)session.getAttribute("muser_name");
    String lguser_type = ( session.getAttribute("muser_type")==null ) ? "" : (String)session.getAttribute("muser_type");
    String lguser_bank_type = ( session.getAttribute("bank_type")==null ) ? "" : (String)session.getAttribute("bank_type");
	
    // ===============================================================================================================
    String tbank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");
    String nowtbank_no = ( request.getParameter("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");
    if ( !nowtbank_no.equals("") ){
      session.setAttribute("nowtbank_no",nowtbank_no); // 將已點選的 tbank_no 寫入 session	   
    }   
    tbank_no = ( session.getAttribute("nowtbank_no") == null ) ? tbank_no : (String)session.getAttribute("nowtbank_no");		
    // ===============================================================================================================

    if(!Utility.CheckPermission(request,"FX013W")){//無權限時,導向到LoginError.jsp
      rd = application.getRequestDispatcher( LoginErrorPgName );        
    } else {
      // Set Next JSP
      if ( act.equals("New") ) {
        String m_year = ( request.getParameter("m_year")==null ) ? "" : (String)request.getParameter("m_year");
        String m_month = ( request.getParameter("m_month")==null ) ? "" : (String)request.getParameter("m_month");
        System.out.println("m_month="+m_month);

        request.setAttribute("BN01_Data",getBN01_Data(tbank_no, m_year));

        request.setAttribute("BA01_Data",getBA01_Data(tbank_no, m_year));

        request.setAttribute("CD034_Data",getCDSHARENO_Data("034"));

        request.setAttribute("CD035_Data",getCDSHARENO_Data("035"));

        rd = application.getRequestDispatcher( EditPgName +"?tbank_no="+tbank_no+"&m_year="+m_year+"&m_month="+m_month);        
      // Edit Page ====================================================================================================
      } else if ( act.equals("Edit") ) {
      // 取得 Sequnce Number
        String seq_no = ( request.getParameter("seq_no")==null ) ? "" : (String)request.getParameter("seq_no");
        String m_year = ( request.getParameter("m_year")==null ) ? "" : (String)request.getParameter("m_year");
        String m_month = ( request.getParameter("m_month")==null ) ? "" : (String)request.getParameter("m_month");
    			
        List  dbData = getGage_Data(seq_no, m_year, m_month, tbank_no);
        request.setAttribute("WLX13_S_Edit",dbData);

        request.setAttribute("BN01_Data",getBN01_Data(tbank_no, m_year));

        request.setAttribute("BA01_Data",getBA01_Data(tbank_no, m_year));

        request.setAttribute("CD034_Data",getCDSHARENO_Data("034"));

        request.setAttribute("CD035_Data",getCDSHARENO_Data("035"));

        rd = application.getRequestDispatcher( EditPgName +"?tbank_no="+tbank_no );
      // List all the bank ===========================================================================================
      } else if ( act.equals("List") ) {
        // 所有資料 (order by year quarter )
        List dbData1 = getBank_Data(tbank_no);
        request.setAttribute("WLX13_S",dbData1);

        // 所有資料之筆數 (order by year quarter )
        List dbData2 = getBank_Sum(tbank_no);
        request.setAttribute("WLX13_S_Sum",dbData2);

        rd = application.getRequestDispatcher(ListPgName +"?tbank_no="+tbank_no);
      // Insert  data ==================================================
      } else if ( act.equals("Insert") ) {
        actMsg = InsertDB(request,lguser_id,lguser_name);
        rd = application.getRequestDispatcher( nextPgName+"?FX=FX013W" ); 
      // Update  data ==================================================
      } else if ( act.equals("Update") ) {
        actMsg = UpdateDB(request,lguser_id,lguser_name);
        rd = application.getRequestDispatcher( nextPgName+"?FX=FX013W" ); 
      // Delete  data ==================================================
      } else if ( act.equals("Delete") ) {
        actMsg = DeleteDB(request,lguser_id,lguser_name);
        rd = application.getRequestDispatcher( nextPgName+"?FX=FX013W" ); 
      }
      request.setAttribute("actMsg",actMsg);
    }
    
    System.out.println("FX013W End ..........");
    // ------導向---------------- 
    try {
      // forward to next present jsp
      rd.forward(request, response);
    } 
    catch (NullPointerException npe) {
      System.out.println(npe);
    }
 } // end of doProcess
//=================================================================================================
%>

<%!
//=================================================================================================
  private final static String nextPgName = "/pages/ActMsg.jsp";
  private final static String EditPgName = "/pages/FX013W_Edit.jsp";
  private final static String ListPgName = "/pages/FX013W_List.jsp";
  private final static String LoginErrorPgName = "/pages/LoginError.jsp";
   
 
  //獲得點選 bank 每年每月的資料總筆數 (BY bank number)=============
  private List getBank_Sum(String tbank_no) {
    //查詢條件    
    List paramList = new ArrayList() ;
    String sqlCmd = "Select WT.M_YEAR, WT.M_MONTH, Count(*) as Cnt "
                  + "From WLX_TRAINNING WT "
                  + "Where TBANK_NO=? "
                  + "Group By M_YEAR, M_MONTH";
	paramList.add(tbank_no);    		   
    List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"M_YEAR, M_MONTH, Cnt");
        
    return dbData;
  } // end of get_bank_List     
    
  //獲得點選bank資料 (BY bank number)==================================================
  private List getBank_Data(String tbank_no) {
    List paramList = new ArrayList() ;	    
    //查詢條件    
    String sqlCmd = "Select WT.M_YEAR, WT.M_MONTH, WT.TBANK_NO, WT.BANK_NO, "
                  + "SEQ_NO, BA01.BANK_NAME, NAME, "
                  + "POSITION_CODE, F_TRANSCODE('034',POSITION_CODE) As POSITION_NAME, "
                  + "TRAIN_PLACE, F_TRANSCODE('035',TRAIN_PLACE) as TRAIN_PLACE_NAME, "
                  + "COURSE_NAME, COURSE_HOUR, WT.USER_ID, WT.USER_NAME, WT.UPDATE_DATE "
                  + "From WLX_TRAINNING WT "                  
                  + "Left Join BA01 On BA01.M_YEAR = WT.M_YEAR And WT.BANK_NO = BA01.BANK_NO "
                  + "Where TBANK_NO=?"
                  + "Order By WT.M_YEAR, WT.M_MONTH, WT.TBANK_NO, SEQ_NO";        	
    paramList.add(tbank_no);
    List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"M_YEAR, M_MONTH, SEQ_NO, COURSE_HOUR, UPDATE_DATE");
  
    return dbData;  
  }//end of getBank_Data 

  //獲得總機構代號及名稱列表 ==================================================
  private List getBN01_Data(String tbank_no, String m_year) {
    List paramList = new ArrayList() ;	
    
    //查詢條件
    String sqlCmd = " Select BANK_NO, BANK_NAME "
                  + " From BN01 "
                  + " Where BANK_NO=? And M_YEAR=?" ;
	 
	paramList.add(tbank_no);       
	paramList.add(m_year);            
    List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");
    return dbData;
  }//end of getBN01_Data
  
  //獲得總分支機構代號及名稱列表 ==================================================
  private List getBA01_Data(String tbank_no, String m_year) {
    List paramList = new ArrayList() ;	
    //查詢條件
    String sqlCmd = " Select BANK_NO, BANK_NAME, BANK_KIND "
                  + " From BA01 "
                  + " Where PBANK_NO=?"
                  + " And M_YEAR=? Order By BANK_NO";
    paramList.add(tbank_no);       
	paramList.add(m_year);              
    List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");
    return dbData;
  }//end of getBA01_Data
  
  //獲得職稱或訓練機構代碼及名稱 ==================================================
  private List getCDSHARENO_Data(String cmuse_div) {
    List paramList = new ArrayList() ;	
    //查詢條件
    String sqlCmd = " Select CMUSE_ID, CMUSE_NAME "
                  + " From CDSHARENO "
                  + " Where CMUSE_DIV=? Order By OUTPUT_ORDER";
	paramList.add(cmuse_div);                  
    List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");
    return dbData;
  }//end of getBA01_Data

  // 獲得要求的資料 (BY data seq_no)================================================
  private List getGage_Data(String seq_no,String m_year,String m_month,String tbank_no){
    List paramList = new ArrayList() ;	
    // 查詢條件
    String sqlCmd = " Select * From WLX_TRAINNING "
                  + " Where M_YEAR = ?" 
                  + " And M_MONTH = ?" 
                  + " And TBANK_NO = ?"
                  + " And SEQ_NO = ?"  ;
	paramList.add(m_year);
	paramList.add(m_month);
	paramList.add(tbank_no);
	paramList.add(seq_no);
    List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"M_YEAR,M_MONTH,SEQ_NO,COURSE_HOUR,BEGIN_DATE,END_DATE");
    return dbData;
  }//end of getGage_Data

  //插入新增的資料 ()=================================================================
  public String InsertDB(HttpServletRequest request,String lguser_id,String lguser_name) throws Exception{    	
    List paramList = new ArrayList() ;	
    String sqlCmd = "";
    String errMsg = "";

    String m_year = ((String)request.getParameter("m_year")==null)?"":(String)request.getParameter("m_year");
    String m_month = ((String)request.getParameter("m_month")==null)?"":(String)request.getParameter("m_month");
    String tbank_no = ((String)request.getParameter("tbank_no")==null)?"":(String)request.getParameter("tbank_no");
    String bank_no = ((String)request.getParameter("bank_no")==null)?"":(String)request.getParameter("bank_no");
    String name = ((String)request.getParameter("name")==null)?"":(String)request.getParameter("name");
    String position_code = ((String)request.getParameter("position_code")==null)?"":(String)request.getParameter("position_code");
    String train_place = ((String)request.getParameter("train_place")==null)?"":(String)request.getParameter("train_place");
    String course_name = ((String)request.getParameter("course_name")==null)?"":(String)request.getParameter("course_name");
    String course_hour = ((String)request.getParameter("course_hour")==null)?"":(String)request.getParameter("course_hour");
    String begin_date_y	= ((String)request.getParameter("begin_date_y")==null)?"":(String)request.getParameter("begin_date_y");
    String begin_date_m	= ((String)request.getParameter("begin_date_m")==null)?"":(String)request.getParameter("begin_date_m");
    String begin_date_d	= ((String)request.getParameter("begin_date_d")==null)?"":(String)request.getParameter("begin_date_d");
    String end_date_y	= ((String)request.getParameter("end_date_y")==null)?"":(String)request.getParameter("end_date_y");
    String end_date_m	= ((String)request.getParameter("end_date_m")==null)?" ":(String)request.getParameter("end_date_m");
    String end_date_d	= ((String)request.getParameter("end_date_d")==null)?" ":(String)request.getParameter("end_date_d");
    String licno = ((String)request.getParameter("licno")==null)?" ":(String)request.getParameter("licno");
    String other_trainplace	= ((String)request.getParameter("other_trainplace")==null)?" ":(String)request.getParameter("other_trainplace");

    try {
      List max_seq_number = new LinkedList();
      List updateDBSqlList = new LinkedList();
      int temp_year = 0;

      //民國轉西元-----------------
      if ( begin_date_y.length() < 4 ) {
        temp_year = Integer.parseInt(begin_date_y) + 1911;
        begin_date_y = Integer.toString(temp_year);
      }
      String begin_date = begin_date_y + "/" + begin_date_m + "/" + begin_date_d;

      //民國轉西元-----------------
      if ( end_date_y.length() < 4 ) {
        temp_year = Integer.parseInt(end_date_y) + 1911;
        end_date_y = Integer.toString(temp_year);
      }
      String end_date = end_date_y + "/" + end_date_m + "/" + end_date_d;

      //先去取得seq number
      sqlCmd = "Select To_Char(Max(SEQ_NO)+1) as MAXQ From WLX_TRAINNING Where M_YEAR = ?"
             + " And M_MONTH = ?"
             + " And TBANK_NO = ?";
	  paramList.add(m_year);
	  paramList.add(m_month);
	  paramList.add(tbank_no);
      max_seq_number = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");	
      
      System.out.println("max_seq_number="+ max_seq_number.size());
      
      String max_seq = (max_seq_number.size() == 0 || ((DataObject)max_seq_number.get(0)).getValue("maxq") == null) ? "1" :(String)((DataObject)max_seq_number.get(0)).getValue("maxq");
      
      paramList.clear() ;
      sqlCmd = "INSERT INTO WLX_TRAINNING (M_YEAR, M_MONTH, TBANK_NO, BANK_NO, SEQ_NO,"
             + "NAME, POSITION_CODE, TRAIN_PLACE, COURSE_NAME, COURSE_HOUR, BEGIN_DATE,"
             + "END_DATE, LICNO, OTHER_TRAINPLACE, USER_ID, USER_NAME, UPDATE_DATE) "
             + "VALUES(?,?,?,?,?,?,?,?,?,?,to_date(?,'YYYY/MM/DD'),to_date(?,'YYYY/MM/DD'),?,?,?,?,sysdate)";
             
      paramList.add(m_year);
      paramList.add(m_month);
      paramList.add(tbank_no);
      paramList.add(bank_no);
      paramList.add(max_seq);
      paramList.add(name);
      paramList.add(position_code);
      paramList.add(train_place);
      paramList.add(course_name);
      paramList.add(course_hour);
      paramList.add(begin_date);
      paramList.add(end_date);
      paramList.add(licno);
      paramList.add(other_trainplace);
      paramList.add(lguser_id);
      paramList.add(lguser_name);      
      
      if(this.updDbUsesPreparedStatement(sqlCmd,paramList)) {
	   	errMsg = errMsg + "相關資料寫入資料庫成功";
	  }else{
	   	errMsg = errMsg + "相關資料寫入資料庫失敗" ;
	  }

      		    
    }catch (Exception e){
      System.out.println(e+":"+e.getMessage());
      errMsg = errMsg + "相關資料寫入資料庫失敗";								
    }//end of catch	
    
    return errMsg;
  } //end of insert

  //修改所點選的資料 ()==================================================
  public String UpdateDB(HttpServletRequest request,String lguser_id,String lguser_name) throws Exception{    	
    List paramList = new ArrayList() ;
    String sqlCmd = "";
    String errMsg = "";

    String m_year = ((String)request.getParameter("m_year")==null)?"":(String)request.getParameter("m_year");
    String m_month = ((String)request.getParameter("m_month")==null)?"":(String)request.getParameter("m_month");
    String tbank_no = ((String)request.getParameter("tbank_no")==null)?"":(String)request.getParameter("tbank_no");
    String bank_no = ((String)request.getParameter("bank_no")==null)?"":(String)request.getParameter("bank_no");
    String seq_no = ((String)request.getParameter("seq_no")==null)?"":(String)request.getParameter("seq_no");
    String name = ((String)request.getParameter("name")==null)?"":(String)request.getParameter("name");
    String position_code = ((String)request.getParameter("position_code")==null)?"":(String)request.getParameter("position_code");
    String train_place = ((String)request.getParameter("train_place")==null)?"":(String)request.getParameter("train_place");
    String course_name = ((String)request.getParameter("course_name")==null)?"":(String)request.getParameter("course_name");
    String course_hour = ((String)request.getParameter("course_hour")==null)?"":(String)request.getParameter("course_hour");
    String begin_date_y	= ((String)request.getParameter("begin_date_y")==null)?"":(String)request.getParameter("begin_date_y");
    String begin_date_m	= ((String)request.getParameter("begin_date_m")==null)?"":(String)request.getParameter("begin_date_m");
    String begin_date_d	= ((String)request.getParameter("begin_date_d")==null)?"":(String)request.getParameter("begin_date_d");
    String end_date_y	= ((String)request.getParameter("end_date_y")==null)?"":(String)request.getParameter("end_date_y");
    String end_date_m	= ((String)request.getParameter("end_date_m")==null)?" ":(String)request.getParameter("end_date_m");
    String end_date_d	= ((String)request.getParameter("end_date_d")==null)?" ":(String)request.getParameter("end_date_d");
    String licno = ((String)request.getParameter("licno")==null)?" ":(String)request.getParameter("licno");
    String other_trainplace	= ((String)request.getParameter("other_trainplace")==null)?" ":(String)request.getParameter("other_trainplace");

    try {
      List data = new LinkedList();
      List updateDBSqlList = new LinkedList();

      sqlCmd = " Select * From WLX_TRAINNING "
             + " Where M_YEAR = ? "
             + " And M_MONTH = ? " 
             + " And TBANK_NO = ? "
             + " And SEQ_NO = ? ";
      paramList.add(m_year);             
      paramList.add(m_month);
      paramList.add(tbank_no);
      paramList.add(seq_no);
           
      data = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");		
			
      if (data.size() == 0){
        errMsg = errMsg + "此筆資料不存在無法修改<br>";
      }else{

        int temp_year = 0;

        //民國轉西元-----------------
        if ( begin_date_y.length() < 4 ) {
          temp_year = Integer.parseInt(begin_date_y) + 1911;
          begin_date_y = Integer.toString(temp_year);
        }
        String begin_date = begin_date_y + "/" + begin_date_m + "/" + begin_date_d;

        //民國轉西元-----------------
        if ( end_date_y.length() < 4 ) {
          temp_year = Integer.parseInt(end_date_y) + 1911;
          end_date_y = Integer.toString(temp_year);
        }
        String end_date = end_date_y + "/" + end_date_m + "/" + end_date_d;
        
		paramList.clear() ;
        //Update ---------------  
        sqlCmd = " UPDATE WLX_TRAINNING SET "
               + " M_YEAR=?,"
               + " M_MONTH=?,"
               + " TBANK_NO=?,"
               + " BANK_NO=?,"
               + " SEQ_NO=?,"
               + " NAME=?,"
               + " POSITION_CODE=?,"
               + " TRAIN_PLACE=?,"
               + " COURSE_NAME=?,"
               + " COURSE_HOUR=?,"
               + " BEGIN_DATE=to_date(?,'YYYY/MM/DD'),"
               + " END_DATE=to_date(?,'YYYY/MM/DD'),"
               + " LICNO=?,"
               + " OTHER_TRAINPLACE=?,"
               + " USER_ID=?,"
               + " USER_NAME=?,"
               + " UPDATE_DATE = sysdate"
               + " WHERE M_YEAR =?" 
               + " And M_MONTH = ?" 
               + " And TBANK_NO = ?"
               + " And SEQ_NO = ? ";
		paramList.add(m_year);
		paramList.add(m_month);
		paramList.add(tbank_no);
		paramList.add(bank_no);
		paramList.add(seq_no);
		paramList.add(name);
		paramList.add(position_code);
		paramList.add(train_place);
		paramList.add(course_name);
		paramList.add(course_hour);
		paramList.add(begin_date);
		paramList.add(end_date);
		paramList.add(licno);
		paramList.add(other_trainplace);
		paramList.add(lguser_id);
		paramList.add(lguser_name);
		paramList.add(m_year);
		paramList.add(m_month);
		paramList.add(tbank_no);
		paramList.add(seq_no);
		
        if(this.updDbUsesPreparedStatement(sqlCmd,paramList)) {
	   	   errMsg = errMsg + "相關資料寫入資料庫成功";
	  	}else{
	   		errMsg = errMsg + "相關資料寫入資料庫失敗" ;
	  	}
      }//END OF IF-ELSE

    }catch (Exception e){
      System.out.println(e+":"+e.getMessage());
      errMsg = errMsg + "相關資料寫入資料庫失敗";								
    }//end of catch	
    return errMsg;
} //end of update
	
//刪除所點選的資料 ()====================================================
public String DeleteDB(HttpServletRequest request,String lguser_id,String lguser_name) throws Exception{    	
  List paramList = new ArrayList() ;
  String sqlCmd = "";
  String errMsg = "";

  String m_year = ((String)request.getParameter("m_year")==null)?"":(String)request.getParameter("m_year");
  String m_month = ((String)request.getParameter("m_month")==null)?"":(String)request.getParameter("m_month");
  String tbank_no = ((String)request.getParameter("tbank_no")==null)?"":(String)request.getParameter("tbank_no");
  String seq_no = ((String)request.getParameter("seq_no")==null)?"":(String)request.getParameter("seq_no");

  try {
    List data = new LinkedList();
    List updateDBSqlList = new LinkedList();

       
    sqlCmd = " Select * From WLX_TRAINNING "
           + " Where M_YEAR = ? "
           + " And M_MONTH = ? " 
           + " And TBANK_NO = ? "
           + " And SEQ_NO = ? ";
    paramList.add(m_year);             
    paramList.add(m_month);
    paramList.add(tbank_no);
    paramList.add(seq_no);
         
    data = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");		
			
    if (data.size() == 0) {
      errMsg = errMsg + "此筆資料不存在無法刪除<br>";
    }
    else{
      //Delete  =========================================================
      sqlCmd  = " Delete WLX_TRAINNING "                
              + " Where M_YEAR = ? "
              + " And M_MONTH = ? "
              + " And TBANK_NO = ? "
              + " And SEQ_NO = ? ";
        
      if(this.updDbUsesPreparedStatement(sqlCmd,paramList)) {
	   	   errMsg = errMsg + "相關資料刪除成功";
	  }else{
	   		errMsg = errMsg + "相關資料刪除失敗" ;
	  } 		
      
    }			    
  }catch (Exception e){
    System.out.println(e+":"+e.getMessage());
    errMsg = errMsg + "相關資料刪除失敗";								
  }//end of catch		
  return errMsg;
} //end of Delete

private boolean updDbUsesPreparedStatement(String sql ,List paramList) throws Exception{
		List updateDBList = new ArrayList();//0:sql 1:data
	    List updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
		List updateDBDataList = new ArrayList();//儲存參數的List
		
		updateDBDataList.add(paramList);
		updateDBSqlList.add(sql);
		updateDBSqlList.add(updateDBDataList);
		updateDBList.add(updateDBSqlList);
		return DBManager.updateDB_ps(updateDBList) ;
}
%>