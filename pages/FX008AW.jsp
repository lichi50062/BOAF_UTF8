<%
// 94.10.14 first designed by lilic0c0 4183
// 99.12.06 fix sqlInjection by 2808 
//100.01.27 fix bug by 2295
%>

<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%@ page import="java.util.ArrayList" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Date" %>
<%
	System.out.println("FX008W start ..........");
  
	RequestDispatcher rd = null;
	String actMsg = "";	
	String alertMsg = "";	
	String webURL = "";	
	boolean doProcess = false;	
	
	//取得session資料,取得成功時,才繼續往下執行===================================================
	if(session.getAttribute("muser_id") == null){//session timeout	
      System.out.println("FX008AW login timeout");   
	   rd = application.getRequestDispatcher( "/pages/reLogin.jsp?url=LoginError.jsp?timeout=true" );         	   
	   try{
          rd.forward(request,response);
       }catch(Exception e){
          System.out.println("forward Error:"+e+e.getMessage());
       }
    }else{
      doProcess = true;
    }    
    
	if(doProcess){//若muser_id資料時,表示登入成功====================================================================	
	
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");			
	String bank_type = ( request.getParameter("bank_type")==null ) ? "" : (String)request.getParameter("bank_type");							
	String list_type = ( request.getParameter("list_type")==null ) ? "" : (String)request.getParameter("list_type");				
	String debt_name = ( request.getParameter("debt_name")==null ) ? "" : (String)request.getParameter("debt_name");
	String dure_no = ( request.getParameter("dure_no")==null ) ? "" : (String)request.getParameter("dure_no");

	System.out.println("act="+act);		
	System.out.println("list_type="+list_type);	
	System.out.println("debt_name="+debt_name);	
	System.out.println("dure_no="+dure_no);
	//登入者資訊
	String lguser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		
	String lguser_name = ( session.getAttribute("muser_name")==null ) ? "" : (String)session.getAttribute("muser_name");		
	String lguser_type = ( session.getAttribute("muser_type")==null ) ? "" : (String)session.getAttribute("muser_type");							
	String lguser_bank_type = ( session.getAttribute("bank_type")==null ) ? "" : (String)session.getAttribute("bank_type");
	
	//======================================================================================================================											
	//session.setAttribute("nowtbank_no",null);//94.01.05 fix 沒有Bank_List,把所點選的Bank_no清除 //94.11.3 關掉
	//fix 94.01.07 若有已點選的tbank_no,則以已點選的tbank_no為主
	String bank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");				
	String nowtbank_no =  ( request.getParameter("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");			
	if(!nowtbank_no.equals("")){
	   session.setAttribute("nowtbank_no",nowtbank_no);//將已點選的tbank_no寫入session	   
	}   
	bank_no = ( session.getAttribute("nowtbank_no")==null ) ? bank_no : (String)session.getAttribute("nowtbank_no");		
	//=======================================================================================================================
		
    	if(!Utility.CheckPermission(request,"FX008AW")){//無權限時,導向到LoginError.jsp
        	rd = application.getRequestDispatcher( LoginErrorPgName );        
    	}else{            
    	//set next jsp 	

    	if(act.equals("Edit")){
    		//取得sequnce number
    		String seq_no =  ( request.getParameter("seq_no")==null ) ? "" : (String)request.getParameter("seq_no");
    		String m_year    =  ( request.getParameter("m_year")==null ) ? "" : (String)request.getParameter("m_year");
    		String m_quarter =  ( request.getParameter("m_quarter")==null ) ? "" : (String)request.getParameter("m_quarter");
    			
    		//取得在list中所點選的那筆資料
    		List  dbData = getGage_Data(seq_no,m_year,m_quarter,bank_no);
    		session.setAttribute("WLX08_AS_Edit",dbData);
    		
    		//取得這年這季呈報出去的文號以及日期
    		List  dbData2 = getRptBOAF_Data(m_year,m_quarter,bank_no);
    		session.setAttribute("WLX08_RptBOAF_Data",dbData2);
    		
    		rd = application.getRequestDispatcher( EditPgName +"?bank_no="+bank_no);
      
      	//List all the bank============================================
    	}else if(act.equals("List")){    	    	    
    	    //所有資料 (order by year quarter )
    	    List dbData1 = getBank_Data(bank_no);
    	    request.setAttribute("WLX08_AS",dbData1);
    	    
    	    //所有資料之筆數 (order by year quarter )
    	    List dbData2 = getBank_Sum(bank_no);
    	    request.setAttribute("WLX08_AS_Sum",dbData2);
    	    
    	    
    	    //取得可申報日期的起始年月份
    	    List dbData3 =getBank_Lock(bank_no);
    	    session.setAttribute("WLX08_ALock",dbData3);
    	    
        	rd = application.getRequestDispatcher(ListPgName +"?bank_no="+bank_no);
  
    	//Update  data ==================================================
    	}else if(act.equals("Update")){
    		List WLX08_S_Edit = (List)session.getAttribute("WLX08_AS_Edit");	
    		
    		actMsg = UpdateDB(request,lguser_id,lguser_name,WLX08_S_Edit);
    		
    		rd = application.getRequestDispatcher( nextPgName+"?FX=FX008AW" ); 
    	//Load  data ==================================================    	  
    	}else if(act.equals("load")){
        	rd = application.getRequestDispatcher(LoadPgName +"?debt_name="+debt_name+"&bank_no="+bank_no);
    	}
    	
    	request.setAttribute("actMsg",actMsg);    
    }
    
    System.out.println("FX008W End ..........");
    
    // ------導向------- 
    try{
    	//forward to next present jsp
        rd.forward(request, response);
    }
    catch (NullPointerException npe){
    	System.out.println(npe);
    }
    
 }//end of doProcess
//=================================================================================================
%>

<%!
//=================================================================================================
    private final static String nextPgName = "/pages/ActMsg.jsp";    
    private final static String EditPgName = "/pages/FX008AW_Edit.jsp";    
    private final static String ListPgName = "/pages/FX008AW_List.jsp";
    private final static String LoadPgName = "/pages/FX008AW_Load.jsp";             
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
   
     
    //獲得點選bank已被Lock的資料 ()===========================
    private List getBank_Lock(String bank_no){
    List paramList = new ArrayList();	
 	String sqlCmd = " select M_YEAR,M_QUARTER from WLX_APPLY_LOCK "
 		      + " where BANK_CODE = ? "
 		      + " and REPORT_NO = 'C04'"          
 		      + " and((LOCK_OWN   =  'Y') or (LOCK_MGR = 'Y'))";
 	paramList.add(bank_no) ;	       				
 	List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"M_YEAR,M_QUARTER");
 			
 	List dbData_Lock = new LinkedList();
 	String inputValue = "";
 			 
 	for(int i=0;i<dbData.size();i++){
 		inputValue = ((DataObject)dbData.get(i)).getValue("m_year").toString()+((DataObject)dbData.get(i)).getValue("m_quarter").toString();
 		dbData_Lock.add(inputValue);
 	}
        
        return dbData_Lock;
    }//end of getBank_Lock
    
    
  //獲得該借款人點選bank已被Lock的資料 ()===========================
    private List getdebt_name_Lock(String bank_no,String debt_name){
  /*  	
 	String sqlCmd = " select M_YEAR,M_QUARTER from WLX_APPLY_LOCK "
 		      + " where BANK_CODE = "+bank_no
 		      + " and DEBTNAME = "+debt_name
 		      + " and REPORT_NO = 'C04'"          
 		      + " and((LOCK_OWN   =  'Y') or (LOCK_MGR = 'Y'))";
 	*/
 	String sqlCmd = " select M_YEAR,M_QUARTER from WLX_APPLY_LOCK "
 		      + " where  REPORT_NO = 'C04'"          
 		      + " and((LOCK_OWN   =  'Y') or (LOCK_MGR = 'Y'))";
 	
 	System.out.println("獲得該借款人點選bank已被Lock的資料 sqlCmd ="+sqlCmd); 	       				
 	List dbData = DBManager.QueryDB_SQLParam(sqlCmd,null,"M_YEAR,M_QUARTER");
 			
 	List dbData_Lock = new LinkedList();
 	String inputValue = "";
 			 
 	for(int i=0;i<dbData.size();i++){
 		inputValue = ((DataObject)dbData.get(i)).getValue("m_year").toString()+((DataObject)dbData.get(i)).getValue("m_quarter").toString();
 		dbData_Lock.add(inputValue);
 	}       
        return dbData_Lock;
    }//end of getBank_Lock
        
    
    //獲得點選bank 每年每季的資料總筆數 (BY bank number)=============
    private List getBank_Sum(String bank_no){
    	//查詢條件    
    	List paramList = new ArrayList() ;
    	String sqlCmd 	= " select  aa.bank_no, aa.M_YEAR,  aa.M_Quarter, sum(Cnt)  as  Cnt "
    					+ " from ( (select  aa.bank_no, aa.M_YEAR,  aa.M_Quarter, count(*)  as  Cnt"
    					+ " from  WLX08_S_GAGE  aa where  aa.bank_no = ? GROUP BY aa.bank_no, aa.M_YEAR,  aa.M_Quarter) Union all "
    					+ " (select  aa.bank_no, aa.M_YEAR,  aa.M_Quarter, 0  as  Cnt from"
    					+ " WLX08_S_GAGE_APPLY  aa where  aa.bank_no= ?)) aa GROUP BY aa.bank_no, aa.M_YEAR,  aa.M_Quarter "
    					+ " Order BY aa.bank_no, aa.M_YEAR desc, aa.M_Quarter  desc ";
    	paramList.add(bank_no) ;
    	paramList.add(bank_no) ;
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"aa.bank_no,aa.M_YEAR,aa.M_Quarter,aa.Cnt");
        
        return dbData;
    }//end of get_bank_List
    
    //獲得點選bank 每年每季的資料總筆數 (BY bank number)=============
    private List getdebt_name_Sum(String bank_no,String debt_name,String dure_no){
    	//查詢條件 
    	/*   
    	String sqlCmd 	= " select  aa.bank_no, aa.M_YEAR,  aa.M_Quarter, sum(Cnt)  as  Cnt "
    					+ " from ( (select  aa.bank_no, aa.M_YEAR,  aa.M_Quarter, count(*)  as  Cnt"
    					+ " from  WLX08_S_GAGE  aa where  aa.bank_no = '"
    					+ bank_no
    					+ "' and aa.debtname = '"
    					+ debt_name
    					+ "' GROUP BY aa.bank_no, aa.M_YEAR,  aa.M_Quarter) Union all "
    					+ " (select  aa.bank_no, aa.M_YEAR,  aa.M_Quarter, 0  as  Cnt from"
    					+ " WLX08_S_GAGE_APPLY  aa where  aa.bank_no= '"
    					+ bank_no
    					+ "' )) aa GROUP BY aa.bank_no, aa.M_YEAR,  aa.M_Quarter "
    					+ " Order BY aa.bank_no, aa.M_YEAR desc, aa.M_Quarter  desc ";
    		*/
    		List paramList = new ArrayList() ;
    		String sqlCmd 	= " select  aa.bank_no, aa.M_YEAR,  aa.M_Quarter, sum(Cnt)  as  Cnt "
    					+ " from ( (select  aa.bank_no, aa.M_YEAR,  aa.M_Quarter, count(*)  as  Cnt"
    					+ " from  WLX08_S_GAGE  aa where  aa.debtname =? and dureassure_no <>? GROUP BY aa.bank_no, aa.M_YEAR,  aa.M_Quarter) Union all "
    					+ " (select  aa.bank_no, aa.M_YEAR,  aa.M_Quarter, 0  as  Cnt from"
    					+ " WLX08_S_GAGE_APPLY  aa "
    					+ " )) aa GROUP BY aa.bank_no, aa.M_YEAR,  aa.M_Quarter "
    					+ " Order BY aa.bank_no, aa.M_YEAR desc, aa.M_Quarter  desc ";
    		paramList.add(debt_name) ;
    		paramList.add(dure_no) ;
    		//System.out.println("獲得點選bank 每年每季的資料總筆數 sqlCmd ="+sqlCmd); 	  
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"aa.bank_no,aa.M_YEAR,aa.M_Quarter,aa.Cnt");
        
        return dbData;
    }//end of get_bank_List     
    
    //獲得點選bank資料 (BY bank number)==================================================
    private List getBank_Data(String bank_no){
    		List paramList =new ArrayList() ;
    		//查詢條件    
    		String sqlCmd = "select * from wlx08_s_gage where bank_no=? Order BY M_YEAR desc,M_Quarter  desc,DUREASSURE_NO";
    	    paramList.add(bank_no) ;
        	List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"M_YEAR,M_QUARTER,BANK_NO,SEQ_NO,DUREASSURE_NO,DEBTNAME,DUREDATE,"+
        										   "DUREASSURESITE,ACCOUNTAMT,APPLYDELAYYEAR,APPLYDELAYMONTH,"+
        										   "AUDIT_APPLYDELAYYEAR,AUDIT_APPLYDELAYMONTH,AUDIT_DUREDATE,"+
        										   "damage_yn,disposal_fact_yn,disposal_plan_yn,auditresult_yn,"+
        										   "APPLYOK_DATE,REPORT_BOAF_DATE,UPDATE_DATE");
        	return dbData;  
    }//end of getBank_Data 
    
    //獲得點選借款人於該bank資料 (BY bank number)==================================================
    private List getdebt_name_Data(String bank_no,String debt_name){
    	
    		//查詢條件    
    		/*
    		String sqlCmd = "select * from wlx08_s_gage where bank_no='"+bank_no+"' and debtname='"+debt_name+"' Order BY M_YEAR desc,M_Quarter  desc,DUREASSURE_NO";
    	  */
    	  String sqlCmd = "select * from wlx08_s_gage where debtname=? Order BY M_YEAR desc,M_Quarter  desc,DUREASSURE_NO";
		  List paramList  = new ArrayList () ;
		  paramList.add(debt_name) ;
    	  //System.out.println("獲得點選借款人於該bank資料 sqlCmd ="+sqlCmd);  
        	List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"M_YEAR,M_QUARTER,BANK_NO,SEQ_NO,DUREASSURE_NO,DEBTNAME,DUREDATE,"+
        										   "DUREASSURESITE,ACCOUNTAMT,APPLYDELAYYEAR,APPLYDELAYMONTH,"+
        										   "AUDIT_APPLYDELAYYEAR,AUDIT_APPLYDELAYMONTH,AUDIT_DUREDATE,"+
        										   "damage_yn,disposal_fact_yn,disposal_plan_yn,auditresult_yn,"+
        										   "APPLYOK_DATE,REPORT_BOAF_DATE,UPDATE_DATE");
        	return dbData;  
    }//end of getBank_Data 
    
    
    
    
    
    
    //獲得要求的資料 (BY data seq_no)================================================
    private List getGage_Data(String seq_no,String m_year,String m_quarter,String bank_no){
    	//查詢條件    
    	String sqlCmd 	= " select * from wlx08_s_gage where m_year =?"
    			+ " and m_quarter =?"
    			+ " and bank_no = ?"
    			+ " and seq_no = ?";
    	List paramList = new ArrayList() ;
    	paramList.add(m_year) ;
    	paramList.add(m_quarter) ;
    	paramList.add(bank_no) ;
    	paramList.add(seq_no) ;
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"M_YEAR,M_QUARTER,BANK_NO,SEQ_NO,DUREASSURE_NO,DEBTNAME,DUREDATE,"+
        										   "DUREASSURESITE,ACCOUNTAMT,APPLYDELAYYEAR,APPLYDELAYMONTH,"+
        										   "AUDIT_APPLYDELAYYEAR,AUDIT_APPLYDELAYMONTH,AUDIT_DUREDATE,"+
        										   "damage_yn,disposal_fact_yn,disposal_plan_yn,auditresult_yn,"+
        										   "APPLYOK_DATE,REPORT_BOAF_DATE,UPDATE_DATE");
        
        return dbData;
    }//end of getGage_Data
    
    //取得呈報出去的文號以及日期 (BY year,quarter,bank_no)================================================
    private List getRptBOAF_Data(String m_year,String m_quarter,String bank_no){
    	//查詢條件    
    	String sqlCmd = " SELECT DISTINCT ApplyOK_DocNo, ApplyOK_Date, Report_BOAF_DocNo, Report_BOAF_Date,UPDATE_DATE"
    				+ " from WLX08_S_GAGE where M_YEAR =?"
    				+ " and m_quarter =?"
    				+ " and bank_no =? "
    				+ " and ApplyOK_DocNo <> ' '"
    				+ " ORDER BY UPDATE_DATE DESC";
    	List paramList = new ArrayList() ;
    	paramList.add(m_year) ;
    	paramList.add(m_quarter) ;
    	paramList.add(bank_no) ;
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"ApplyOK_Date,Report_BOAF_Date");
        
        return dbData;
    }//end of getGage_Data
  
  //更新要求的資料 (BY)================================================  
  public String UpdateDB(HttpServletRequest request,String lguser_id,String lguser_name,List WLX08_S_Edit) throws Exception{    	
	
	StringBuffer sqlCmd = new StringBuffer() ;		
	String errMsg = "";
	
	String m_year		=((DataObject)WLX08_S_Edit.get(0)).getValue("m_year").toString();
	String m_quarter 	=((DataObject)WLX08_S_Edit.get(0)).getValue("m_quarter").toString();
	String bank_no		=((DataObject)WLX08_S_Edit.get(0)).getValue("bank_no").toString();
	String seq_no		=((DataObject)WLX08_S_Edit.get(0)).getValue("seq_no").toString();
	String dure_no		=((DataObject)WLX08_S_Edit.get(0)).getValue("dureassure_no").toString();
	String debtname 	=((DataObject)WLX08_S_Edit.get(0)).getValue("debtname").toString();
	String duresite 	=((DataObject)WLX08_S_Edit.get(0)).getValue("dureassuresite").toString();
	String account		=((DataObject)WLX08_S_Edit.get(0)).getValue("accountamt").toString();
	String apply_year	=((DataObject)WLX08_S_Edit.get(0)).getValue("applydelayyear").toString();
	String apply_month	=((DataObject)WLX08_S_Edit.get(0)).getValue("applydelaymonth").toString();
	String apply_reason	=((DataObject)WLX08_S_Edit.get(0)).getValue("applydelayreason").toString();	
	Date duDate 		=(Date)((DataObject)WLX08_S_Edit.get(0)).getValue("duredate");
	int du_year   = duDate.getYear()+1900;                                                      
	int du_month  = duDate.getMonth()+1;                                                      
	int du_day    = duDate.getDate();    
	
	
	//取得參數
	String damage_yn		=(request.getParameter("damage_yn")==null)?" ":request.getParameter("damage_yn").toString();
	String disposal_fact_yn	=(request.getParameter("disposal_fact_yn")==null)?" ":request.getParameter("disposal_fact_yn").toString();
	String disposal_plan_yn =(request.getParameter("disposal_plan_yn")==null)?" ":request.getParameter("disposal_plan_yn").toString();
	String auditresult_yn	=(request.getParameter("auditresult_yn")==null)?" ":request.getParameter("auditresult_yn").toString();
	String audit_applydelayyear	 	=(request.getParameter("Audit_Delay_y")==null)?" ":request.getParameter("Audit_Delay_y").toString();
	String audit_applydelaymonth 	=(request.getParameter("Audit_Delay_m")==null)?" ":request.getParameter("Audit_Delay_m").toString();
	String applyok_docno 			=(request.getParameter("DocNo")==null)?" ":request.getParameter("DocNo").toString();
	String report_boaf_docno 		=(request.getParameter("BOAF_DocNo")==null)?" ":request.getParameter("BOAF_DocNo").toString();
	//-------
	String Cnt_year  =(request.getParameter("Cnt_year")==null)?" ":request.getParameter("Cnt_year").toString();
	String Cnt_month =(request.getParameter("Cnt_month")==null)?" ":request.getParameter("Cnt_month").toString();
	String Cnt_date  =(request.getParameter("Cnt_date")==null)?" ":request.getParameter("Cnt_date").toString();
	System.out.println("Cnt_year="+Cnt_year);
	int iCnt_year  = Integer.parseInt(Cnt_year)+1911;
	int iCnt_month = Integer.parseInt(Cnt_month);
	int iCnt_date  = Integer.parseInt(Cnt_date);
	//-------
	String ApplyOK_y =(request.getParameter("ApplyOK_y")==null)?" ":request.getParameter("ApplyOK_y").toString();
	String ApplyOK_m =(request.getParameter("ApplyOK_m")==null)?" ":request.getParameter("ApplyOK_m").toString();
	String ApplyOK_d =(request.getParameter("ApplyOK_d")==null)?" ":request.getParameter("ApplyOK_d").toString();
	int iApplyOK_y = Integer.parseInt(ApplyOK_y)+1911;
	int iApplyOK_m = Integer.parseInt(ApplyOK_m);
	int iApplyOK_d = Integer.parseInt(ApplyOK_d);
	//-------
	String BOAF_y 	=(request.getParameter("BOAF_y")==null)?" ":request.getParameter("BOAF_y").toString();
	String BOAF_m 	=(request.getParameter("BOAF_m")==null)?" ":request.getParameter("BOAF_m").toString();
	String BOAF_d 	=(request.getParameter("BOAF_d")==null)?" ":request.getParameter("BOAF_d").toString();
	int iBOAF_y = Integer.parseInt(BOAF_y)+1911;
	int iBOAF_m = Integer.parseInt(BOAF_m);
	int iBOAF_d = Integer.parseInt(BOAF_d);
	
	try {
			List data = new LinkedList();
			//List updateDBSqlList = new LinkedList();
			List paramList = new ArrayList() ;
			sqlCmd.append(" select * from WLX08_S_GAGE where m_year = ?"
						 + " and m_quarter = ?"
						 + " and bank_no = ?"
						 + " and seq_no =?");			
			paramList.add(m_year) ;
			paramList.add(m_quarter) ;
			paramList.add(bank_no) ;
			paramList.add(seq_no) ;
			data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");		
			paramList.clear() ;
			sqlCmd.setLength(0) ;
		 if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法修改<br>";
		 }
		 else{
			//Insert to log ---------------
		 	sqlCmd.append("Insert INTO WLX08_S_GAGE_log "
		 			+ "SELECT M_YEAR,M_QUARTER,BANK_NO,SEQ_NO,DUREASSURE_NO,"
  					+ "DEBTNAME,DUREDATE,DUREASSURESITE,ACCOUNTAMT,APPLYDELAYYEAR,APPLYDELAYMONTH,"
  					+ "APPLYDELAYREASON,DAMAGE_YN,DISPOSAL_FACT_YN,DISPOSAL_PLAN_YN,"
  					+ "AUDITRESULT_YN,USER_ID,USER_NAME,UPDATE_DATE,"
  					+ "?,?,sysdate,'U',"
  					+ " AUDIT_APPLYDELAYYEAR,AUDIT_APPLYDELAYMONTH,AUDIT_DUREDATE,APPLYOK_DOCNO,"
	   				+ " APPLYOK_DATE,REPORT_BOAF_DOCNO,REPORT_BOAF_DATE" 
		          	+ " from WLX08_S_GAGE"
		          	+ " WHERE bank_no= ?"
		          	+ " AND m_year = ?" 
		          	+ " AND m_quarter = ?"
		          	+ " AND SEQ_NO = ? ");
			paramList.add(lguser_id) ;
			paramList.add(lguser_name) ;
			paramList.add(bank_no) ;
			paramList.add(m_year) ;
			paramList.add(m_quarter) ;
			paramList.add(seq_no) ;
		 	//updateDBSqlList.add(sqlCmd);
		    
		    if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){
   				errMsg = errMsg + "相關資料log寫入成功\n";
   			}
   			else{
   				errMsg = errMsg + "相關資料log寫入失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg()+"\n";
   			}	
		  //Update ---------------  
		    sqlCmd.setLength(0) ;
		    paramList.clear() ;
			sqlCmd.append("UPDATE wlx08_s_gage SET " ) ;    
			sqlCmd.append("DUREASSURE_NO=?"); paramList.add( dure_no  ) ;       
			sqlCmd.append(",DEBTNAME=?") ;paramList.add( debtname ) ;      
			sqlCmd.append(",DUREDATE=to_date(?,'yyyy/mm/dd')" ) ;paramList.add(du_year+"/"+du_month+"/"+du_day) ; 
			sqlCmd.append(",DUREASSURESITE=?"); paramList.add( duresite)  ;     
			sqlCmd.append(",ACCOUNTAMT=?"); paramList.add( account);       
			sqlCmd.append(",APPLYDELAYYEAR=?"); paramList.add( apply_year );    
			sqlCmd.append(",APPLYDELAYMONTH=?"); paramList.add( apply_month ) ;  
			sqlCmd.append(",APPLYDELAYREASON=?"); paramList.add( apply_reason  ) ;  
			sqlCmd.append(",DAMAGE_YN =?"); paramList.add( damage_yn  ) ;     
			sqlCmd.append(",DISPOSAL_FACT_YN=?"); paramList.add( disposal_fact_yn  ) ; 
			sqlCmd.append(",DISPOSAL_PLAN_YN=?"); paramList.add( disposal_plan_yn  );					       
			sqlCmd.append(",AUDITRESULT_YN=?"); paramList.add( auditresult_yn  ) ;     
			sqlCmd.append(",USER_ID=?"); paramList.add( lguser_id ) ;
			sqlCmd.append(",USER_NAME=?");paramList.add(lguser_name  ) ;
			sqlCmd.append(",UPDATE_DATE = sysdate" ) ;
			sqlCmd.append(",AUDIT_APPLYDELAYYEAR = ?");paramList.add( audit_applydelayyear ) ;   
			sqlCmd.append(",AUDIT_APPLYDELAYMONTH = ?");paramList.add( audit_applydelaymonth ) ;
			sqlCmd.append(",AUDIT_DUREDATE = to_date(?,'yyyy/mm/dd')");paramList.add(iCnt_year+"/"+iCnt_month+"/"+iCnt_date) ;
			sqlCmd.append(",APPLYOK_DOCNO = ?"); paramList.add( applyok_docno) ;        
			sqlCmd.append(",APPLYOK_DATE  =  to_date(?,'yyyy/mm/dd')" ) ;paramList.add(iApplyOK_y+"/"+iApplyOK_m+"/"+iApplyOK_d) ;      
			sqlCmd.append(",REPORT_BOAF_DOCNO = ?"); paramList.add(report_boaf_docno ) ;
			sqlCmd.append(",REPORT_BOAF_DATE  = to_date(?,'yyyy/mm/dd')" ) ;paramList.add(iBOAF_y+"/"+iBOAF_m+"/"+iBOAF_d) ;
			sqlCmd.append(" WHERE M_YEAR =?");paramList.add(m_year ) ;
			sqlCmd.append(" and M_QUARTER = ?");paramList.add(m_quarter) ;
			sqlCmd.append(" and BANK_NO = ?");paramList.add(bank_no ) ;
			sqlCmd.append(" and SEQ_NO = ?") ;paramList.add(seq_no);
			/*		+ "DUREASSURE_NO='" + dure_no + "'"       
					+ ",DEBTNAME='" + debtname + "'"      
					+ ",DUREDATE=to_date('"+du_year+"/"+du_month+"/"+du_day+"','yyyy/mm/dd')" 
					+ ",DUREASSURESITE='" + duresite + "'"      
					+ ",ACCOUNTAMT='" + account + "'"       
					+ ",APPLYDELAYYEAR='" + apply_year + "'"    
					+ ",APPLYDELAYMONTH='" + apply_month + "'"   
					+ ",APPLYDELAYREASON='" + apply_reason + "'"  
					+ ",DAMAGE_YN ='" + damage_yn + "'"     
					+ ",DISPOSAL_FACT_YN='" + disposal_fact_yn + "'" 
					+ ",DISPOSAL_PLAN_YN='" + disposal_plan_yn +"'"					       
					+ ",AUDITRESULT_YN='" + auditresult_yn + "'"     
					+ ",USER_ID='" + lguser_id + "'"
					+ ",USER_NAME='"+lguser_name + "'"
					+ ",UPDATE_DATE = sysdate"
					+ ",AUDIT_APPLYDELAYYEAR = '"+ audit_applydelayyear+ "'"   
					+ ",AUDIT_APPLYDELAYMONTH = '"+ audit_applydelaymonth+ "'"
					+ ",AUDIT_DUREDATE = to_date('"+iCnt_year+"/"+iCnt_month+"/"+iCnt_date+"','yyyy/mm/dd')"
					+ ",APPLYOK_DOCNO = '" + applyok_docno+"'"        
  					+ ",APPLYOK_DATE  =  to_date('"+iApplyOK_y+"/"+iApplyOK_m+"/"+iApplyOK_d+"','yyyy/mm/dd')"      
  					+ ",REPORT_BOAF_DOCNO = '" +report_boaf_docno+"'"
  					+ ",REPORT_BOAF_DATE  = to_date('"+iBOAF_y+"/"+iBOAF_m+"/"+iBOAF_d+"','yyyy/mm/dd')"
					+ " WHERE M_YEAR ="+m_year
					+ " and M_QUARTER = "+m_quarter
					+ " and BANK_NO = "+bank_no
					+ " and SEQ_NO = "+seq_no;
         		 						
				updateDBSqlList.clear();
   				updateDBSqlList.add(sqlCmd);*/
   			
   			//if(DBManager.updateDB(updateDBSqlList)){
   			if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)) {
   				errMsg = errMsg + "相關資料寫入資料庫成功";
   			}
   			else{
   				errMsg = errMsg + "相關資料寫入資料庫失敗";
   			}	
   		}			    
		}catch (Exception e){
			System.out.println(e+":"+e.getMessage());
			errMsg = errMsg + "相關資料寫入資料庫失敗";						
		}//end of catch	
		
		return errMsg;
		
	} //end of update
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
