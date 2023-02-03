<%
// 93.12.20 add 權限檢核 by 2295
//          add 設定異動者資訊 by 2295
// 93.12.23 add 超過登入時間,請重新登入 by 2295
// 94.01.05 fix 沒有Bank_List,把所點選的Bank_no清除 by 2295
// 94.01.07 fix super user可以看該bank_type的基本資料維護,管理者只能看自己的 by 2295
// 94.01.12 fix 登入者的bank_type != bank_type時,不代入tbank_no by 2295
// 99.12.08 fix sqlInjection by 2808
//100.01.26 fix bank_cmml區分99/100年度 by 2295

%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.ArrayList" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%
	RequestDispatcher rd = null;
	String actMsg = "";	
	String alertMsg = "";	
	String webURL = "";	
	boolean doProcess = false;	
	
	//取得session資料,取得成功時,才繼續往下執行===================================================
	if(session.getAttribute("muser_id") == null){//session timeout	
      System.out.println("FX001W login timeout");   
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
	//String bank_no = ( request.getParameter("bank_no")==null ) ? "" : (String)request.getParameter("bank_no");				
	String list_type = ( request.getParameter("list_type")==null ) ? "" : (String)request.getParameter("list_type");				
	System.out.println("act="+act);		
	System.out.println("list_type="+list_type);	
	
	//登入者資訊
	String lguser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		
	String lguser_name = ( session.getAttribute("muser_name")==null ) ? "" : (String)session.getAttribute("muser_name");		
	String lguser_type = ( session.getAttribute("muser_type")==null ) ? "" : (String)session.getAttribute("muser_type");							
	String lguser_bank_type = ( session.getAttribute("bank_type")==null ) ? "" : (String)session.getAttribute("bank_type");							
	String tbank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");						
	session.setAttribute("nowtbank_no",null);//94.01.05 fix 沒有Bank_List,把所點選的Bank_no清除======
	//======================================================================================================================
	//fix 94.01.07 若有已點選的tbank_no,則以已點選的tbank_no為主============================================================
	String bank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");				
	String nowtbank_no =  ( request.getParameter("bank_no")==null ) ? "" : (String)request.getParameter("bank_no");			
	if(!nowtbank_no.equals("")){
	   session.setAttribute("nowtbank_no",nowtbank_no);//將已點選的tbank_no寫入session	   
	}   
	bank_no = ( session.getAttribute("nowtbank_no")==null ) ? bank_no : (String)session.getAttribute("nowtbank_no");			
	//=======================================================================================================================
		
    if(!Utility.CheckPermission(request,"FX004W")){//無權限時,導向到LoginError.jsp   
        rd = application.getRequestDispatcher( LoginErrorPgName );        
    }else{            
    	//set next jsp 	
    	if(act.equals("new")){
        	rd = application.getRequestDispatcher( EditPgName +"?act=new");        
    	}else if(act.equals("Edit")){    	    
    	    List dbData = getBANK_CMML(bank_no);
    	    request.setAttribute("BANK_CMML",dbData);
    	    //93.12.20設定異動者資訊======================================================================
			request.setAttribute("maintainInfo","select * from BANK_CMML WHERE bank_no='" + bank_no+"'");								       
			//=======================================================================================================================		
        	rd = application.getRequestDispatcher( EditPgName +"?act=Edit&list_type="+list_type);        
    	}else if(act.equals("List")){    	    	    
    	    //94.01.12 fix 登入者的bank_type != bank_type時,不代tbank_no
    	    if(!lguser_bank_type.equals(bank_type)){
    	       tbank_no = "";
    	    }
    	    List dbData = getBank_No(bank_type,lguser_type,tbank_no);
    	    request.setAttribute("BA01",dbData);
        	rd = application.getRequestDispatcher( ListPgName +"?act=List&list_type="+list_type );        
        	
    	}else if(act.equals("Update")){
    	    actMsg = UpdateDB(request,bank_no,lguser_id,lguser_name);
        	rd = application.getRequestDispatcher( nextPgName );        
    	}
    	request.setAttribute("actMsg",actMsg);    
    }        
     
%>

<%
	try {
        //forward to next present jsp
        rd.forward(request, response);
    } catch (NullPointerException npe) {
    }
    }//end of doProcess
%>


<%!
    private final static String nextPgName = "/pages/ActMsg.jsp";    
    private final static String EditPgName = "/pages/FX004W_Edit.jsp";    
    private final static String ListPgName = "/pages/FX004W_List.jsp";        
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
            
    /***
    *  99.12.08 fix 縣市合併 by 2808
    * 100.01.26 fix bank_cmml區分99/100年度 by 2295
    */
    private List getBANK_CMML(String bank_no){
    		//查詢條件    
    		List paramList =new ArrayList () ;
    		String yy = Integer.parseInt(Utility.getYear()) > 99? "100" : "99" ;
    		String sqlCmd = "select * from (select * from ba01 where m_year=?)ba01 LEFT JOIN (select * from bank_cmml where m_year=?)bank_cmml ON bank_cmml.bank_no = ba01.bank_no "
						  + " where ba01.bank_no=? ";		
			paramList.add(yy) ;			  
    		paramList.add(yy) ;
    		paramList.add(bank_no) ;
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"business_person,update_date");            
            return dbData;
    } 
    /***
     * 99.12.08 fix 縣市合併 by 2808
     */
    private List getBank_No(String bank_type,String lguser_type,String tbank_no){
    		//查詢條件    
    		List paramList =new ArrayList() ;
    		String yy = Integer.parseInt(Utility.getYear()) > 99? "100" : "99" ;
    		String sqlCmd = "select * from ba01 where bank_type=? and m_year=? ";
    		paramList.add(bank_type) ;
    		paramList.add(yy);
    		if(lguser_type.equals("A") && !tbank_no.equals("")){    		    
    		   sqlCmd += "  and bank_no=? ";
    		   paramList.add(tbank_no) ;
    		}
    		/*
    		if(list_type.equals("1")){
				   sqlCmd += " where bank_type='B'";//地方主管機關		
			}else if(list_type.equals("2")){
			       sqlCmd += " where bank_type='8'";//共用中心		
			}else if(list_type.equals("3")){
			       sqlCmd += " where bank_type  in ('1','2','3','4','5','8','9','A')";//農業行庫		
			}                       
    		*/
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
            return dbData;
    } 
    //100.01.26 fix bank_cmml區分99/100年度 by 2295
	public String UpdateDB(HttpServletRequest request,String bank_no,String lguser_id,String lguser_name) throws Exception{    	
		StringBuffer sqlCmd = new StringBuffer();	
		List paramList = new ArrayList(); 
		String errMsg="";		
		String bank_name=((String)request.getParameter("BANK_NAME")==null)?"":(String)request.getParameter("BANK_NAME");
		String bank_english=((String)request.getParameter("BANK_ENGLISH")==null)?"":(String)request.getParameter("BANK_ENGLISH");		
		String business_person=((String)request.getParameter("BUSINESS_PERSON") == null || ((String)request.getParameter("BUSINESS_PERSON")).equals(""))?"0" : (String)request.getParameter("BUSINESS_PERSON");
		String hsien_id_area_id=((String)request.getParameter("HSIEN_ID_AREA_ID")==null)?"":(String)request.getParameter("HSIEN_ID_AREA_ID");
		String hsien_id = hsien_id_area_id.substring(0,hsien_id_area_id.indexOf("/"));
		String area_id = hsien_id_area_id.substring(hsien_id_area_id.indexOf("/")+1,hsien_id_area_id.length());
		String addr=((String)request.getParameter("ADDR")==null)?"":(String)request.getParameter("ADDR");
		String web_site=((String)request.getParameter("WEB_SITE")==null)?"":(String)request.getParameter("WEB_SITE");
		String m_position=((String)request.getParameter("M_POSITION")==null)?"":(String)request.getParameter("M_POSITION");
		String m_name=((String)request.getParameter("M_NAME")==null)?"":(String)request.getParameter("M_NAME");		
		String m_telno=((String)request.getParameter("M_TELNO")==null)?"":(String)request.getParameter("M_TELNO");
		String m_cellino=((String)request.getParameter("M_CELLINO")==null)?"":(String)request.getParameter("M_CELLINO");
		String m_fax=((String)request.getParameter("M_FAX")==null)?"":(String)request.getParameter("M_FAX");		
		String m_email=((String)request.getParameter("M_EMAIL")==null)?"":(String)request.getParameter("M_EMAIL");
		String m_sex=((String)request.getParameter("M_SEX")==null)?"":(String)request.getParameter("M_SEX");
		String m_position_officer=((String)request.getParameter("M_POSITION_OFFICER")==null)?"":(String)request.getParameter("M_POSITION_OFFICER");
		String m_name_officer=((String)request.getParameter("M_NAME_OFFICER")==null)?"":(String)request.getParameter("M_NAME_OFFICER");		
		String m_telno_officer=((String)request.getParameter("M_TELNO_OFFICER")==null)?"":(String)request.getParameter("M_TELNO_OFFICER");
		String m_cellino_officer=((String)request.getParameter("M_CELLINO_OFFICER")==null)?"":(String)request.getParameter("M_CELLINO_OFFICER");
		String m_fax_officer=((String)request.getParameter("M_FAX_OFFICER")==null)?"":(String)request.getParameter("M_FAX_OFFICER");		
		String m_email_officer=((String)request.getParameter("M_EMAIL_OFFICER")==null)?"":(String)request.getParameter("M_EMAIL_OFFICER");
		String m_sex_officer=((String)request.getParameter("M_SEX_OFFICER")==null)?"":(String)request.getParameter("M_SEX_OFFICER");		
		String user_id=lguser_id;
	    String user_name=lguser_name;	
		
		try {
				String yy = Integer.parseInt(Utility.getYear()) > 99? "100" : "99" ;
				//List updateDBSqlList = new LinkedList();
				sqlCmd.append("SELECT * FROM BANK_CMML WHERE bank_no=? and m_year=?");					 
				paramList.add(bank_no) ;	 
				paramList.add(yy);//100.01.26
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"business_person");		 			    
				System.out.println("BANK_CMML.size="+data.size());
				sqlCmd.setLength(0) ;
			    paramList.clear() ;
				if (data.size() == 0){//無資料時,Insert
					sqlCmd.append(" insert into bank_cmml Values( ");
				    sqlCmd.append("?,?,?,?,?,?,?,?,?,?,");
				    sqlCmd.append("?,?,?,?,?,?,?,?,?,?,");
				    sqlCmd.append("?,?,?,?,sysdate,?)");
				    paramList.add(bank_no);
				    paramList.add(bank_name );
			        paramList.add(bank_english ); 					       
			        paramList.add(business_person); 
			        paramList.add(hsien_id );
			        paramList.add(area_id ); 
			        paramList.add(addr ); 
			        paramList.add(web_site ); 
			        paramList.add(m_position ); 
			        paramList.add(m_name );
			        paramList.add(m_telno );
			        paramList.add(m_cellino );
			        paramList.add(m_fax );
			        paramList.add(m_email );					       
			        paramList.add(m_sex );
			        paramList.add(m_position_officer ); 
			        paramList.add(m_name_officer );
			        paramList.add(m_telno_officer );
			        paramList.add(m_cellino_officer );
			        paramList.add(m_fax_officer );
			        paramList.add(m_email_officer );					       
			        paramList.add(m_sex_officer );					       
			        paramList.add(user_id );			       
			        paramList.add(user_name );
			        paramList.add(yy);//100.01.26
				}else{//有資料時,Update    
				    sqlCmd.append("UPDATE BANK_CMML SET ");
				    sqlCmd.append(" bank_name=?"); paramList.add( bank_name); 
			        sqlCmd.append(",bank_english=?"); paramList.add( bank_english );					       
			        sqlCmd.append(",business_person=?");paramList.add(business_person); 
			        sqlCmd.append(",hsien_id=?"); paramList.add( hsien_id );
			        sqlCmd.append(",area_id=?"); paramList.add( area_id ); 
			        sqlCmd.append(",addr=?"); paramList.add(addr); 
			        sqlCmd.append(",web_site=?"); paramList.add( web_site ); 
			        sqlCmd.append(",m_position=?"); paramList.add( m_position ); 
			        sqlCmd.append(",m_name=?"); paramList.add( m_name );
			        sqlCmd.append(",m_telno=?"); paramList.add( m_telno );
			        sqlCmd.append(",m_cellino=?"); paramList.add( m_cellino );
			        sqlCmd.append(",m_fax=?"); paramList.add( m_fax );
			        sqlCmd.append(",m_email=?");paramList.add( m_email );					       
			        sqlCmd.append(",m_sex=?"); paramList.add( m_sex );
			        sqlCmd.append(",m_position_officer=?"); paramList.add( m_position_officer ); 
			        sqlCmd.append(",m_name_officer=?"); paramList.add( m_name_officer );
			        sqlCmd.append(",m_telno_officer=?"); paramList.add( m_telno_officer);
			        sqlCmd.append(",m_cellino_officer=?"); paramList.add( m_cellino_officer);
			        sqlCmd.append(",m_fax_officer=?"); paramList.add( m_fax_officer );
			        sqlCmd.append(",m_email_officer=?"); paramList.add( m_email_officer);					       
			        sqlCmd.append(",m_sex_officer=?"); paramList.add( m_sex_officer );					       
			        sqlCmd.append(",user_id=?"); paramList.add( user_id);		       
			        sqlCmd.append(",user_name=?"); paramList.add( user_name);
			        sqlCmd.append(",update_date=sysdate"); 		            		 	
		    	    sqlCmd.append(" where bank_no=?");paramList.add(bank_no);
		    	    sqlCmd.append(" and  m_year=?");paramList.add(yy);
	    				    	   
    	   		}
    	   		
   				//updateDBSqlList.add(sqlCmd); 		            	
	            		
				if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){					 
					errMsg = errMsg + "相關資料寫入資料庫成功";					
				}else{
				 	errMsg = errMsg + "相關資料寫入資料庫失敗";
				}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";				
		}	

		return errMsg;
	} 
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