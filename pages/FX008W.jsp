<%
// 94.10.14 first designed by 4183
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
      System.out.println("FX008W login timeout");   
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
	System.out.println("act="+act);		
	System.out.println("list_type="+list_type);	
	
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
		
    	if(!Utility.CheckPermission(request,"FX008W")){//無權限時,導向到LoginError.jsp   
        	rd = application.getRequestDispatcher( LoginErrorPgName );        
    	}else{            
    	//set next jsp 	
    	if(act.equals("New")){
        	rd = application.getRequestDispatcher( EditPgName +"?bank_no="+bank_no);        
    	//Edit page==========================================================
    	}else if(act.equals("Edit")){
    			//取得sequnce number
    			String seq_no =  ( request.getParameter("seq_no")==null ) ? "" : (String)request.getParameter("seq_no");
    			String m_year    =  ( request.getParameter("m_year")==null ) ? "" : (String)request.getParameter("m_year");
    			String m_quarter =  ( request.getParameter("m_quarter")==null ) ? "" : (String)request.getParameter("m_quarter");
    			
    			List  dbData = getGage_Data(seq_no,m_year,m_quarter,bank_no);
    			request.setAttribute("WLX08_S_Edit",dbData);

    	    rd = application.getRequestDispatcher( EditPgName +"?bank_no="+bank_no);
      	//List all the bank============================================
    	}else if(act.equals("List")){    	    	    
    	    //所有資料 (order by year quarter )
    	    List dbData1 = getBank_Data(bank_no);
    	    request.setAttribute("WLX08_S",dbData1);
    	    
    	    //所有資料之筆數 (order by year quarter )
    	    List dbData2 = getBank_Sum(bank_no);
    	    request.setAttribute("WLX08_S_Sum",dbData2);
    	    
    	    //取得可申報日期的起始年月份
    	    int  apply_ini = getBank_Ini();
    	    request.setAttribute("apply_ini",Integer.toString(apply_ini) );
    	    
    	    //取得可申報日期的起始年月份
    	    List dbData3 =getBank_Lock(bank_no);
    	    session.setAttribute("WLX08_Lock",dbData3);
    	    
        	rd = application.getRequestDispatcher(ListPgName +"?bank_no="+bank_no);
      //Load last quarter data========================================
      }else if(act.equals("Load")){
      	String m_year    =  ( request.getParameter("s_year")==null ) ? "" : (String)request.getParameter("s_year");
		String m_quarter =  ( request.getParameter("s_quarter")==null ) ? "" : (String)request.getParameter("s_quarter");
        int temp_quarter = Integer.parseInt(m_quarter);
   	    int temp_year = Integer.parseInt(m_year);
    			
    	//如果是本季是第一季,則載入去年最後一季
    	if(	--temp_quarter == 0 ){
    			temp_quarter=4;
    			temp_year--;
    	}//end of if	
    	    
    	//所有資料 (order by year quarter )
    	List dbData = getLoad_Data(Integer.toString(temp_year),Integer.toString(temp_quarter),bank_no);
    	
    	//傳送參數
    	session.setAttribute("FX008_Load",dbData);
    	session.setAttribute("load_year",m_year);
    	session.setAttribute("load_quarter",m_quarter);
       	session.setAttribute("last_year",Integer.toString(temp_year));
        session.setAttribute("last_quarter",Integer.toString(temp_quarter));
        	
        rd = application.getRequestDispatcher(LoadPgName+"?bank_no="+bank_no);
        	
      //Insert  data ==================================================
      }else if(act.equals("Insert")){                            	        
          actMsg = InsertDB(request,lguser_id,lguser_name);      	
    	  rd = application.getRequestDispatcher( nextPgName+"?FX=FX008W" ); 
      
      //No Apply  data ==================================================
      }else if(act.equals("No_Apply")){                            	        
          actMsg = No_ApplyDB(request,lguser_id,lguser_name);      	
    	  rd = application.getRequestDispatcher( nextPgName+"?FX=FX008W" ); 
      //Update  data ==================================================
      }else if(act.equals("Update")){
    	  actMsg = UpdateDB(request,lguser_id,lguser_name);
          rd = application.getRequestDispatcher( nextPgName+"?FX=FX008W" ); 
    	
      //Delete  data ==================================================
      }else if(act.equals("Delete")){
    	  actMsg = DeleteDB(request,lguser_id,lguser_name);
          rd = application.getRequestDispatcher( nextPgName+"?FX=FX008W" ); 
          		
      //Load last quarter data ========================================
      }else if(act.equals("Load_To")){
    	  String m_year    =  ( session.getAttribute("load_year")==null ) ? " " : (String)session.getAttribute("load_year");
    	  String m_quarter =  ( session.getAttribute("load_quarter")==null ) ? " " : (String)session.getAttribute("load_quarter");
    	  String form_size =  ( request.getParameter("form_size")==null ) ? " " : (String)request.getParameter("form_size");
    	  List FX008_Load  =  (List)session.getAttribute("FX008_Load");
    		    			
    	  actMsg = Load_Data(request,lguser_id,lguser_name,m_year,m_quarter,FX008_Load,Integer.parseInt(form_size));
    	  rd = application.getRequestDispatcher( nextPgName+"?FX=FX008W" ); 
	  }else if(act.equals("load_history")){
	  	    String debt_name = ( request.getParameter("debt_name")==null ) ? "" : (String)request.getParameter("debt_name");
        	rd = application.getRequestDispatcher(LoadhistoryPgName +"?debt_name="+debt_name+"&bank_no="+bank_no);
    	}
    
    	request.setAttribute("actMsg",actMsg);    
    }        
    
    System.out.println("FX008W End ..........");
   	// ------導向---------------- 
	try {
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
    private final static String EditPgName = "/pages/FX008W_Edit.jsp";    
    private final static String ListPgName = "/pages/FX008W_List.jsp"; 
    private final static String LoadPgName = "/pages/FX008W_Select.jsp";
    private final static String LoadhistoryPgName = "/pages/FX008AW_Load.jsp";        
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
   
   	//獲得點選bank的起始申報年限資料 ()==================================================
    private int getBank_Ini(){
        List paramList = new ArrayList();
    		int apply_ini = 0;
    		//查詢條件    
    		String sqlCmd = "SELECT M_YEAR,M_MONTH FROM WLX_APPLY_INI WHERE REPORT_NO = ? ";
    		paramList.add("C04"); 
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"M_YEAR,M_MONTH");
        
        if(dbData!=null && dbData.size()>0){
        	apply_ini = Integer.parseInt(((DataObject)dbData.get(0)).getValue("m_year").toString())*12;
        	apply_ini += Integer.parseInt(((DataObject)dbData.get(0)).getValue("m_month").toString());
        }
        
        return apply_ini;
    }//end of getBank_Ini
    
    //獲得點選bank已被Lock的資料 ()===========================
    private List getBank_Lock(String bank_no){
    		List paramList = new ArrayList() ;
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
    
	//獲得點選bank 每年每季的資料總筆數 (BY bank number)=============
    private List getBank_Sum(String bank_no){
    		//查詢條件    
    		List paramList =new ArrayList() ;
    		String sqlCmd = "select  aa.bank_no, aa.M_YEAR,  aa.M_Quarter, sum(Cnt)  as  Cnt "
    			      +"from ( (select  aa.bank_no, aa.M_YEAR,  aa.M_Quarter, count(*)  as  Cnt"
    			      +" from  WLX08_S_GAGE  aa where  aa.bank_no = ?  GROUP BY aa.bank_no, aa.M_YEAR,  aa.M_Quarter) Union all "
    			      +"(select  aa.bank_no, aa.M_YEAR,  aa.M_Quarter, 0  as  Cnt from"
    			      +" WLX08_S_GAGE_APPLY  aa where  aa.bank_no= ? )) aa GROUP BY aa.bank_no, aa.M_YEAR,  aa.M_Quarter "
    			      +"Order BY aa.bank_no, aa.M_YEAR desc, aa.M_Quarter  desc ";
    		paramList.add(bank_no) ;
    		paramList.add(bank_no) ;
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"aa.bank_no,aa.M_YEAR,aa.M_Quarter,aa.Cnt");
        
        return dbData;
    }//end of get_bank_List     
    
    //獲得點選bank資料 (BY bank number)==================================================
    private List getBank_Data(String bank_no){
    	
    		//查詢條件    
    		List paramList =new ArrayList() ;
    		String sqlCmd = "select * from wlx08_s_gage where bank_no=? Order BY M_YEAR desc,M_Quarter  desc,DUREASSURE_NO";
    	    paramList.add(bank_no) ;
        	List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"M_YEAR,M_QUARTER,BANK_NO,SEQ_NO,DUREASSURE_NO,DEBTNAME,DUREDATE,"+
        										   "DUREASSURESITE,ACCOUNTAMT,APPLYDELAYYEAR,APPLYDELAYMONTH,"+
        										   "AUDIT_APPLYDELAYYEAR,AUDIT_APPLYDELAYMONTH,AUDIT_DUREDATE,"+
        										   "APPLYOK_DATE,REPORT_BOAF_DATE,UPDATE_DATE");
        	return dbData;  
    }//end of getBank_Data 
    
    //獲得要求的資料 (BY data seq_no)================================================
    private List getGage_Data(String seq_no,String m_year,String m_quarter,String bank_no){
    		//查詢條件    
    		List paramList =new ArrayList() ;
    		String sqlCmd 	= "select * from wlx08_s_gage where m_year =?"
    										+"and m_quarter =?"
    										+"and bank_no = ?"
    										+"and seq_no = ? ";
    		paramList.add(m_year) ;
    		paramList.add(m_quarter) ;
    		paramList.add(bank_no) ;
    		paramList.add(seq_no) ;
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"M_YEAR,M_QUARTER,BANK_NO,SEQ_NO,DUREASSURE_NO,DEBTNAME,DUREDATE,"+
        										   "DUREASSURESITE,ACCOUNTAMT,APPLYDELAYYEAR,APPLYDELAYMONTH,"+
        										   "AUDIT_APPLYDELAYYEAR,AUDIT_APPLYDELAYMONTH,AUDIT_DUREDATE,"+
        										   "APPLYOK_DATE,REPORT_BOAF_DATE,UPDATE_DATE");
        return dbData;
    }//end of getGage_Data
    
    //獲得點選bank的上季資料 (BY m_year,m_quarter,bank_no)===========================
    private List getLoad_Data(String m_year,String m_quarter,String bank_no){
    	
    		//查詢條件    
    		List paramList =new ArrayList() ;
    		String sqlCmd = "select * from wlx08_s_gage where"
    						+ " M_YEAR = ? "
    						+ " and M_QUARTER =? "
    						+ " and bank_no = ? "
    						+ " order by DUREASSURE_NO desc";
    		paramList.add(m_year) ;
    		paramList.add(m_quarter) ;
    		paramList.add(bank_no) ;
           List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"M_YEAR,M_QUARTER,BANK_NO,SEQ_NO,DUREASSURE_NO,DEBTNAME,DUREDATE,"+
        										   "DUREASSURESITE,ACCOUNTAMT,APPLYDELAYYEAR,APPLYDELAYMONTH,"+
        										   "AUDIT_APPLYDELAYYEAR,AUDIT_APPLYDELAYMONTH,AUDIT_DUREDATE,"+
        										   "APPLYOK_DATE,REPORT_BOAF_DATE,UPDATE_DATE");
        return dbData;
    }//end of getLoad_Data 
    
  //插入新增的資料 ()=================================================================
  public String InsertDB(HttpServletRequest request,String lguser_id,String lguser_name) throws Exception{    	
	
	StringBuffer sqlCmd = new StringBuffer() ;		
	String errMsg = "";
		
	String m_year		=((String)request.getParameter("hyear")==null)?"":(String)request.getParameter("hyear");
	String m_quarter 	=((String)request.getParameter("hquarter")==null)?"":(String)request.getParameter("hquarter");
	String bank_no 		=((String)request.getParameter("hbank_no")==null)?"":(String)request.getParameter("hbank_no");
	String dure_no		=((String)request.getParameter("dure_no")==null)?"":(String)request.getParameter("dure_no");
	String debtname 	=((String)request.getParameter("debtname")==null)?"":(String)request.getParameter("debtname");
	String accept_year 	=((String)request.getParameter("accept_year")==null)?"":(String)request.getParameter("accept_year");
	String accept_month	=((String)request.getParameter("accept_month")==null)?"":(String)request.getParameter("accept_month");
	String accept_day	=((String)request.getParameter("accept_day")==null)?"":(String)request.getParameter("accept_day");
	String duresite 	=((String)request.getParameter("duresite")==null)?"":(String)request.getParameter("duresite");
	String account		=((String)request.getParameter("account")==null)?"":(String)request.getParameter("account");
	String apply_year	=((String)request.getParameter("apply_year")==null)?"":(String)request.getParameter("apply_year");
	String apply_month	=((String)request.getParameter("apply_month")==null)?"":(String)request.getParameter("apply_month");
	String apply_reason	=((String)request.getParameter("apply_reason")==null)?"":(String)request.getParameter("apply_reason");
	String damage_yn	=((String)request.getParameter("damage_yn")==null)?" ":(String)request.getParameter("damage_yn");
	String disposal_fact_yn	=((String)request.getParameter("disposal_fact_yn")==null)?" ":(String)request.getParameter("disposal_fact_yn");
	String disposal_plan_yn =((String)request.getParameter("disposal_plan_yn")==null)?" ":(String)request.getParameter("disposal_plan_yn");
	String auditresult_yn	=((String)request.getParameter("auditresult_yn")==null)?" ":(String)request.getParameter("auditresult_yn");
	
		try {
			List max_seq_number = new LinkedList();
			//List updateDBSqlList = new LinkedList();
			List paramList = new ArrayList() ;
			//民國轉西元-----------------
			if(accept_year.length() < 4){
				int temp_year = Integer.parseInt(accept_year);
				temp_year+=1911;
				accept_year=Integer.toString(temp_year);
			}
			
			//先去取得seq number
			sqlCmd.append("select to_char(max(SEQ_NO)+1) as maxq from WLX08_S_GAGE where m_year = ?"
						 + "and m_quarter = ?"
						 + "and bank_no = ? ");			    
			paramList.add(m_year) ;
			paramList.add(m_quarter) ;
			paramList.add(bank_no) ;
			max_seq_number = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");	
			sqlCmd.setLength(0) ;
			paramList.clear() ;
		    System.out.println("max_seq_number="+ max_seq_number.size());
		  
			String max_seq = (((DataObject)max_seq_number.get(0)).getValue("maxq")==null ) ? "1" :(String)((DataObject)max_seq_number.get(0)).getValue("maxq");
		
			
			/*sqlCmd  = "INSERT INTO wlx08_s_gage (M_YEAR,M_QUARTER,BANK_NO,SEQ_NO,DUREASSURE_NO,"
  					+ "DEBTNAME,DUREDATE,DUREASSURESITE,ACCOUNTAMT,APPLYDELAYYEAR,APPLYDELAYMONTH,"
  					+ "APPLYDELAYREASON,DAMAGE_YN,DISPOSAL_FACT_YN,DISPOSAL_PLAN_YN,AUDITRESULT_YN,"
  					+ "USER_ID,USER_NAME,UPDATE_DATE) VALUES("                   
					+ "'" + m_year + "'"        
					+ ",'" + m_quarter + "'"			            	      
					+ ",'" + bank_no + "'"
					+ ",'" + max_seq+ "'"
					+ ",'" + dure_no + "'"       
					+ ",'" + debtname + "'"      
					+ ",to_date('"+accept_year+"/"+accept_month+"/"+accept_day+"','YYYY/MM/DD')"
					+ ",'" + duresite + "'"      
					+ ",'" + account + "'"       
					+ ",'" + apply_year + "'"    
					+ ",'" + apply_month + "'"   
					+ ",'" + apply_reason + "'"  
					+ ",'" + damage_yn + "'"     
					+ ",'" + disposal_fact_yn + "'" 
					+ ",'" + disposal_plan_yn +"'"					       
					+ ",'" + auditresult_yn + "'"     
					+ ",'" + lguser_id + "'"
					+ ",'" + lguser_name + "'"
					+ ",sysdate)";         		 						

   			updateDBSqlList.add(sqlCmd);*/
   			sqlCmd.append(" INSERT INTO wlx08_s_gage (M_YEAR,M_QUARTER,BANK_NO,SEQ_NO,DUREASSURE_NO, ") ;
   			sqlCmd.append("DEBTNAME,DUREDATE,DUREASSURESITE,ACCOUNTAMT,APPLYDELAYYEAR,APPLYDELAYMONTH,");
   			sqlCmd.append("APPLYDELAYREASON,DAMAGE_YN,DISPOSAL_FACT_YN,DISPOSAL_PLAN_YN,AUDITRESULT_YN,");
   			sqlCmd.append("USER_ID,USER_NAME,UPDATE_DATE) VALUES( ");
   			sqlCmd.append("?,?,?,?,?,?,to_date(?,'YYYY/MM/DD'),?,?,?,?,?,?,?,?,?,?,?,sysdate ) ");
   			paramList.add(m_year) ;
   			paramList.add(m_quarter) ;
   			paramList.add(bank_no) ;
   			paramList.add(max_seq) ;
   			paramList.add(dure_no) ;
   			paramList.add(debtname) ;
   			paramList.add(accept_year+"/"+accept_month+"/"+accept_day) ;
   			paramList.add(duresite) ;
   			paramList.add(account) ;
   			paramList.add(apply_year) ;
   			paramList.add(apply_month) ;
   			paramList.add(apply_reason) ;
   			paramList.add(damage_yn) ;
   			paramList.add(disposal_fact_yn) ;
   			paramList.add(disposal_plan_yn) ;
   			paramList.add(auditresult_yn) ;
   			paramList.add(lguser_id) ;
   			paramList.add(lguser_name) ;
   			if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)) {   			
   				errMsg = errMsg + "相關資料寫入資料庫成功";
   			}
   			else{
   				errMsg = errMsg + "相關資料寫入資料庫失敗";
   			}				    
		}catch (Exception e){
			System.out.println(e+":"+e.getMessage());
			errMsg = errMsg + "相關資料寫入資料庫失敗";					
		}//end of catch	
		
		return errMsg;
	} //end of insert
	
	//記錄本季沒有資料需要申報 ()=================================================================
  	public String No_ApplyDB(HttpServletRequest request,String lguser_id,String lguser_name) throws Exception{    	
	
	StringBuffer  sqlCmd = new StringBuffer() ;		
	String errMsg = "";
		
	String m_year	 =((String)request.getParameter("hyear")==null)?"":(String)request.getParameter("hyear");
	String m_quarter =((String)request.getParameter("hquarter")==null)?"":(String)request.getParameter("hquarter");
	String bank_no 	 =((String)request.getParameter("hbank_no")==null)?"":(String)request.getParameter("hbank_no");
	
		try {
			List updateDBSqlList = new LinkedList();
			List paramList = new ArrayList() ;
			sqlCmd.append("SELECT count(*) as cnt from WLX08_S_GAGE_APPLY "
							+ " WHERE M_YEAR = ?"
							+ " AND M_QUARTER =? "
							+ " AND BANK_NO = ?" ) ;
			paramList.add(m_year) ;
			paramList.add(m_quarter);
			paramList.add(bank_no) ;
			updateDBSqlList = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"cnt");
			
			if(updateDBSqlList != null && updateDBSqlList.size() > 0){
				
				int cnt = Integer.parseInt(((DataObject)updateDBSqlList.get(0)).getValue("cnt").toString());
			
				if( cnt > 0 ){
					errMsg = errMsg + "此筆資料已存在無法新增<br>";
				} else {
					//Insert into		
					sqlCmd.setLength(0) ;
					paramList.clear() ;
					sqlCmd.append("INSERT INTO WLX08_S_GAGE_APPLY (M_YEAR,M_QUARTER,BANK_NO,APPLY_CNT,"
  								+ "USER_ID,USER_NAME,UPDATE_DATE) VALUES("                   
								+ "?"        
								+ ",?"			            	      
								+ ",?"
								+ ",'0'"
								+ ",?"
								+ ",?"
								+ ",sysdate)");         		 						
					paramList.add(m_year) ;
					paramList.add(m_quarter) ;
					paramList.add(bank_no) ;
					paramList.add(lguser_id) ;
					paramList.add(lguser_name) ;
					//updateDBSqlList.clear();
   					//updateDBSqlList.add(sqlCmd);
   					if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){
   					//if(DBManager.updateDB(updateDBSqlList)){
   						errMsg = errMsg + "相關資料寫入資料庫成功";
   					}
   					else{
   						errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" 
   									+ DBManager.getErrMsg();
   					}	
   				}
			} else {
				errMsg = errMsg + "List沒有資料,Null pointer";	
			}
						    
		}catch (Exception e){
			System.out.println(e+":"+e.getMessage());
			errMsg = errMsg + "相關資料寫入資料庫失敗";						
		}//end of catch	
		
		return errMsg;
	} //end of insert
	
  //修改所點選的資料 ()==================================================
  public String UpdateDB(HttpServletRequest request,String lguser_id,String lguser_name) throws Exception{    	
	
	StringBuffer sqlCmd = new StringBuffer() ;		
	String errMsg = "";
		
	String m_year		=((String)request.getParameter("hyear")==null)?"":(String)request.getParameter("hyear");
	String m_quarter 	=((String)request.getParameter("hquarter")==null)?"":(String)request.getParameter("hquarter");
	String bank_no		=((String)request.getParameter("hbank_no")==null)?"":(String)request.getParameter("hbank_no");
	String seq_no		=((String)request.getParameter("hseq_no")==null)?"":(String)request.getParameter("hseq_no");
	String dure_no		=((String)request.getParameter("dure_no")==null)?"":(String)request.getParameter("dure_no");
	String debtname 	=((String)request.getParameter("debtname")==null)?"":(String)request.getParameter("debtname");
	String accept_year 	=((String)request.getParameter("accept_year")==null)?"":(String)request.getParameter("accept_year");
	String accept_month	=((String)request.getParameter("accept_month")==null)?"":(String)request.getParameter("accept_month");
	String accept_day	=((String)request.getParameter("accept_day")==null)?"":(String)request.getParameter("accept_day");
	String duresite 	=((String)request.getParameter("duresite")==null)?"":(String)request.getParameter("duresite");
	String account		=((String)request.getParameter("account")==null)?"":(String)request.getParameter("account");
	String apply_year	=((String)request.getParameter("apply_year")==null)?"":(String)request.getParameter("apply_year");
	String apply_month	=((String)request.getParameter("apply_month")==null)?"":(String)request.getParameter("apply_month");
	String apply_reason	=((String)request.getParameter("apply_reason")==null)?"":(String)request.getParameter("apply_reason");
	String damage_yn	=((String)request.getParameter("damage_yn")==null)?" ":(String)request.getParameter("damage_yn");
	String disposal_fact_yn	=((String)request.getParameter("disposal_fact_yn")==null)?" ":(String)request.getParameter("disposal_fact_yn");
	String disposal_plan_yn =((String)request.getParameter("disposal_plan_yn")==null)?" ":(String)request.getParameter("disposal_plan_yn");
	String auditresult_yn	=((String)request.getParameter("auditresult_yn")==null)?" ":(String)request.getParameter("auditresult_yn");
	
		try {
			List data = new LinkedList();
			//List updateDBSqlList = new LinkedList();
			List paramList =new ArrayList() ;
			sqlCmd.append("select * from WLX08_S_GAGE where m_year = ?"
						 + "and m_quarter = ?"
						 + "and bank_no = ?"
						 + "and seq_no =?");			
			paramList.add(m_year) ;
			paramList.add(m_quarter) ;
			paramList.add(bank_no);
			paramList.add(seq_no);
			data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");		
			
		 if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法修改<br>";
		 }else{
		 	
		 	//民國轉西元-----------------
			if(accept_year.length() < 4){
				int temp_year = Integer.parseInt(accept_year);
				temp_year+=1911;
				accept_year=Integer.toString(temp_year);
			}
			sqlCmd.setLength(0) ;
			paramList.clear() ;
			//Insert to log ---------------
		 	sqlCmd.append( "Insert INTO WLX08_S_GAGE_log "
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
		          	+ " AND SEQ_NO = ?");
		 		paramList.add(lguser_id) ;
		 		paramList.add(lguser_name);
		 		paramList.add(bank_no) ;
		 		paramList.add(m_year);
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
		    paramList.clear() ;
		    sqlCmd.setLength(0) ;
			sqlCmd.append("UPDATE wlx08_s_gage SET ") ;        
			sqlCmd.append( "DUREASSURE_NO=?"); paramList.add( dure_no  ) ;       
			sqlCmd.append( ",DEBTNAME=?"); paramList.add( debtname );      
			sqlCmd.append( ",DUREDATE=to_date(?,'YYYY/MM/DD')");paramList.add(accept_year+"/"+accept_month+"/"+accept_day) ; 
			sqlCmd.append( ",DUREASSURESITE=?"); paramList.add( duresite );      
			sqlCmd.append( ",ACCOUNTAMT=?"); paramList.add( account);       
			sqlCmd.append( ",APPLYDELAYYEAR=?"); paramList.add( apply_year );    
			sqlCmd.append( ",APPLYDELAYMONTH=?"); paramList.add( apply_month );   
			sqlCmd.append( ",APPLYDELAYREASON=?"); paramList.add( apply_reason );  
			sqlCmd.append( ",DAMAGE_YN =?"); paramList.add( damage_yn );     
			sqlCmd.append( ",DISPOSAL_FACT_YN=?"); paramList.add( disposal_fact_yn ); 
			sqlCmd.append( ",DISPOSAL_PLAN_YN=?"); paramList.add( disposal_plan_yn );					       
			sqlCmd.append( ",AUDITRESULT_YN=?"); paramList.add( auditresult_yn  ) ;     
			sqlCmd.append( ",USER_ID=?"); paramList.add( lguser_id  );
			sqlCmd.append( ",USER_NAME=?");paramList.add(lguser_name );
			sqlCmd.append( ",UPDATE_DATE = sysdate");
			sqlCmd.append( " WHERE M_YEAR =?");paramList.add(m_year) ;
			sqlCmd.append( " and M_QUARTER = ?");paramList.add(m_quarter);
			sqlCmd.append( " and BANK_NO = ?");paramList.add(bank_no);
			sqlCmd.append( " and SEQ_NO =?");paramList.add(seq_no);
			/*		+ "DUREASSURE_NO='" + dure_no + "'"       
					+ ",DEBTNAME='" + debtname + "'"      
					+ ",DUREDATE=to_date('"+accept_year+"/"+accept_month+"/"+accept_day+"','YYYY/MM/DD')" 
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
					+ " WHERE M_YEAR ="+m_year
					+ " and M_QUARTER = "+m_quarter
					+ " and BANK_NO = "+bank_no
					+ " and SEQ_NO = "+seq_no;*/
         		 						
				//updateDBSqlList.clear();
   				//updateDBSqlList.add(sqlCmd);
   			
   			//if(DBManager.updateDB(updateDBSqlList)){
   			if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){
   				errMsg = errMsg + "相關資料寫入資料庫成功";
   			}
   			else{
   				errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
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
	
	StringBuffer sqlCmd = new StringBuffer();		
	String errMsg = "";
		
	String m_year		=((String)request.getParameter("hyear")==null)?"":(String)request.getParameter("hyear");
	String m_quarter 	=((String)request.getParameter("hquarter")==null)?"":(String)request.getParameter("hquarter");
	String bank_no		=((String)request.getParameter("hbank_no")==null)?"":(String)request.getParameter("hbank_no");
	String seq_no		=((String)request.getParameter("hseq_no")==null)?"":(String)request.getParameter("hseq_no");
	
		try {
			List data = new LinkedList();
			//List updateDBSqlList = new LinkedList();
			List paramList =new ArrayList() ;
			sqlCmd.append("select * from WLX08_S_GAGE where m_year = ?"
						 + "and m_quarter = ?"
						 + "and bank_no = ?"
						 + "and seq_no =? ");			
			paramList.add(m_year) ;
			paramList.add(m_quarter) ;
			paramList.add(bank_no) ;
			paramList.add(seq_no) ;
			data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");		
			
		 if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法刪除<br>";
		 }
		 else{
		 //Insert to log ==================================================
			paramList.clear() ;
		    sqlCmd.setLength(0) ;
		 	sqlCmd.append("Insert INTO WLX08_S_GAGE_log "
		 			+ "SELECT M_YEAR,M_QUARTER,BANK_NO,SEQ_NO,DUREASSURE_NO,"
  					+ "DEBTNAME,DUREDATE,DUREASSURESITE,ACCOUNTAMT,APPLYDELAYYEAR,APPLYDELAYMONTH,"
  					+ "APPLYDELAYREASON,DAMAGE_YN,DISPOSAL_FACT_YN,DISPOSAL_PLAN_YN,AUDITRESULT_YN,"
  					+ "USER_ID,USER_NAME,UPDATE_DATE,"
		          	+ "?,?,sysdate,'D',"
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
		    
		    //if(DBManager.updateDB(updateDBSqlList)){
		    if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){
   				errMsg = errMsg + "相關資料log寫入成功\n";
   			}
   			else{
   				errMsg = errMsg + "相關資料log寫入失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg()+"\n";
   			}	
		 
		 //Delete  =========================================================
			sqlCmd.setLength(0) ;
		    paramList.clear() ;
			sqlCmd.append("DELETE wlx08_s_gage "                
					+ " WHERE M_YEAR =?"
					+ " and M_QUARTER = ?"
					+ " and BANK_NO = ?"
					+ " and SEQ_NO = ? ");
        	paramList.add(m_year) ;
        	paramList.add(m_quarter) ;
        	paramList.add(bank_no) ;
        	paramList.add(seq_no) ;
        	//updateDBSqlList.clear();
   			//updateDBSqlList.add(sqlCmd);
   			
   			//if(DBManager.updateDB(updateDBSqlList)){
   			if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)) {
   				errMsg = errMsg + "相關資料刪除成功";
   			}
   			else{
   				errMsg = errMsg + "相關資料刪除失敗";
   			}	
   		}			    
	}catch (Exception e){
		System.out.println(e+":"+e.getMessage());
		errMsg = errMsg + "相關資料刪除失敗";						
	}//end of catch		
		
	return errMsg;
} //end of Delete
	
//載入上季資料為本季資料 (BY m_year,m_quarter,bank_no)=====================
public String Load_Data(HttpServletRequest request,String lguser_id,String lguser_name,
  						String m_year,String m_quarter,List FX008_Load,int form_size) throws Exception{
  	
  	StringBuffer sqlCmd = new StringBuffer() ;
	String errMsg = "";							
  	List data = new LinkedList();
  	
  	try {								 
			String temp = "";				
			List paramList = new ArrayList() ;	
			sqlCmd.append("select * from WLX08_S_GAGE "
				 + "where m_year = ? "
				 + "and m_quarter = ? "
				 + "and bank_no = ?");			
			paramList.add(m_year) ;
			paramList.add(m_quarter) ;
			paramList.add(((DataObject)FX008_Load.get(0)).getValue("bank_no").toString()) ;
			data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");		
			
		 	if (data.size() != 0){
				    errMsg = errMsg + "本季已有資料<br>";
			}else{
				String check_yn 		= "";
				String bank_no			= "";
				String dure_no			= "";
				String debtname 		= "";
				Date   duredate;
				String duresite 		= "";
				String account			= "";
				String apply_year		= "";
				String apply_month		= "";
				String apply_reason		= "";
				String max_seq 			= "";
				int accept_year 		= 0;
				int accept_month 		= 0;
				int accept_date			= 0;
 

				for(int i=0;i<form_size;i++){
					//確認這筆資料是否要載入
					check_yn =(String)request.getParameter("C"+Integer.toString(i));
					
					//注意check box如果沒有選的話會傳null值回來 要小心去取到null
					//---------------------------------------------------------------------
					
					if(check_yn!=null){
						System.out.println("第"+i+"項-----------------------"+FX008_Load.size());
						//取得參數
						bank_no		=((DataObject)FX008_Load.get(i)).getValue("bank_no").toString();
						dure_no		=((DataObject)FX008_Load.get(i)).getValue("dureassure_no").toString();
						debtname 	=((DataObject)FX008_Load.get(i)).getValue("debtname").toString();
						duredate	=(Date)((DataObject)FX008_Load.get(i)).getValue("duredate");
						duresite 	=((DataObject)FX008_Load.get(i)).getValue("dureassuresite").toString();
						account		=((DataObject)FX008_Load.get(i)).getValue("accountamt").toString();
						apply_year	=((DataObject)FX008_Load.get(i)).getValue("applydelayyear").toString();
						apply_month	=((DataObject)FX008_Load.get(i)).getValue("applydelaymonth").toString();
						apply_reason=((DataObject)FX008_Load.get(i)).getValue("applydelayreason").toString();
						//把日期做轉換

						accept_year =duredate.getYear()+1900;  
						accept_month =duredate.getMonth()+1; 
						accept_date	=duredate.getDate();   
						
						//先去取得seq number
						sqlCmd.setLength(0) ;
						paramList.clear() ;
						sqlCmd.append("select to_char(max(SEQ_NO)+1) as maxq from WLX08_S_GAGE "
								+ "where m_year = ?"
						 		+ "and m_quarter = ?"
						 		+ "and bank_no = ?");			
						paramList.add( m_year) ;
						paramList.add(m_quarter) ;
						paramList.add(bank_no) ;
						
						data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");	
						max_seq = ((String)((DataObject)data.get(0)).getValue("maxq")==null)?"1":(String)((DataObject)data.get(0)).getValue("maxq");
						
						System.out.println("Max seq ="+ max_seq);
						sqlCmd.setLength(0) ;
						paramList.clear() ;
						//SQL 語句輸入
						sqlCmd.append("INSERT INTO wlx08_s_gage (M_YEAR,M_QUARTER,BANK_NO,SEQ_NO,DUREASSURE_NO,"
  								+ "DEBTNAME,DUREDATE,DUREASSURESITE,ACCOUNTAMT,APPLYDELAYYEAR,APPLYDELAYMONTH,"
  								+ "APPLYDELAYREASON,DAMAGE_YN,DISPOSAL_FACT_YN,DISPOSAL_PLAN_YN,AUDITRESULT_YN,"
  								+ "USER_ID,USER_NAME,UPDATE_DATE) VALUES("                   
								+ "?"        
								+ ",?"			            	      
								+ ",?"
								+ ",?"
								+ ",?"       
								+ ",?"      
								+ ",to_date(?,'YYYY/MM/DD')" 
								+ ",?"      
								+ ",?"       
								+ ",?"    
								+ ",?"   
								+ ",?"  
								+ ",' '"     
								+ ",' '" 
								+ ",' '"					       
								+ ",' '"     
								+ ",?"
								+ ",?"
								+ ",sysdate)");	   
						paramList.add(m_year) ;
						paramList.add(m_quarter) ;
						paramList.add(bank_no) ;
						paramList.add(max_seq) ;
						paramList.add(dure_no) ;
						paramList.add(debtname) ;
						paramList.add(accept_year+"/"+accept_month+"/"+accept_date) ;
						paramList.add(duresite) ;
						paramList.add(account);
						paramList.add(apply_year) ;
						paramList.add(apply_month) ;
						paramList.add(apply_reason) ;
						paramList.add(lguser_id) ;
						paramList.add(lguser_name);
						//System.out.println(sqlCmd);
						
						//用完要記得清掉
   						//data.clear();
   					
   						//data.add(sqlCmd);
						
						//if(DBManager.updateDB(data)){
						if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){	
   							errMsg = errMsg + "相關資料寫入資料庫成功";
   						}
   						else{
   							errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:" + DBManager.getErrMsg();
   							return errMsg;
   						}//end of DBManager.updateDB	
					}//end of if check_yn---------------------------------------------------
				}//end of for
			}//end of if-else(size)
		}catch (Exception e){
			System.out.println(e+":"+e.getMessage());
			errMsg = errMsg + "相關資料寫入資料庫失敗";							
		}//end of catch	
    
   return errMsg;
}//end of Load_Data   
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
