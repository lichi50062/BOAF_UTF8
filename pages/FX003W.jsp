<%
// 93.12.21 add 權限檢核 by 2295
//          add 若有已點選的tbank_no,則以已點選的tbank_no為主 by 2295
//          add 設定異動者資訊
// 93.12.23 add 超過登入時間,請重新登入 by 2295
// 94.04.01 add 主key更改為bank_no+seq_no by 2295
// 95.06.05 add 將異動資料寫入WLX04_LOG by 2295
// 99.12.03 fix sqlInjection by 2808
//102.03.27 add 增加順位排序 by 2295
//102.04.24 add idn加解密  by2968
//102.06.28 add 操作歷程寫入log by2968
//102.12.18 fix 更新理監事ID資料,若無修改ID時,存入DB會變成已mask過的資料 by 2295
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
      System.out.println("FX003W login timeout");   
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
	String seq_no = ( request.getParameter("seq_no")==null ) ? "" : (String)request.getParameter("seq_no");			
	//String bank_no = ( request.getParameter("bank_no")==null ) ? "1234567" : (String)request.getParameter("bank_no");				
	//String position_code = ( request.getParameter("position_code")==null ) ? "" : (String)request.getParameter("position_code");			
	//String id = ( request.getParameter("id")==null ) ? "" : (String)request.getParameter("id");			
	
	//登入者資訊
	String lguser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		
	String lguser_name = ( session.getAttribute("muser_name")==null ) ? "" : (String)session.getAttribute("muser_name");		
	String lguser_type = ( session.getAttribute("muser_type")==null ) ? "" : (String)session.getAttribute("muser_type");						
	//======================================================================================================================
	
	System.out.println("act="+act);		
	//fix 93.12.20 若有已點選的tbank_no,則以已點選的tbank_no為主============================================================
	String bank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");				
	String nowtbank_no =  ( request.getParameter("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");			
	if(!nowtbank_no.equals("")){
	   session.setAttribute("nowtbank_no",nowtbank_no);//將已點選的tbank_no寫入session	   
	}   
	bank_no = ( session.getAttribute("nowtbank_no")==null ) ? bank_no : (String)session.getAttribute("nowtbank_no");			
	//=======================================================================================================================
	
    if(!Utility.CheckPermission(request,"FX003W")){//無權限時,導向到LoginError.jsp   
        rd = application.getRequestDispatcher( LoginErrorPgName );        
    }else{            
    	//set next jsp 	
    	if(act.equals("new")){
        	rd = application.getRequestDispatcher( EditPgName +"?act=new");        
    	}else if(act.equals("Edit")){    	    
    	    List dbData = getWLX04(bank_no,seq_no);
    	    request.setAttribute("WLX04",dbData);
    	    //93.12.21設定異動者資訊======================================================================
			request.setAttribute("maintainInfo","select * from WLX04 WHERE bank_no='" + bank_no+"' and seq_no="+seq_no);								       
			//=======================================================================================================================		        	        	        	
        	//操作歷程寫入log
    		this.InsertWlXOPERATE_LOG(request,lguser_id,program_id,bank_no,"","Q");
			rd = application.getRequestDispatcher( EditPgName +"?act=Edit");        
    	}else if(act.equals("List")){    	    
    	    List dbData = getWLX04(bank_no,"");
    	    request.setAttribute("WLX04",dbData);
        	rd = application.getRequestDispatcher( ListPgName +"?bank_no="+bank_no);        
    	}else if(act.equals("Insert")){
    	    actMsg = InsertDB(request,bank_no,lguser_id,lguser_name);
    	    if("Y".equals(actMsg)){
    	        //操作歷程寫入log
	    		actMsg = this.InsertWlXOPERATE_LOG(request,lguser_id,program_id,bank_no,"","I");
	    		if("Y".equals(actMsg)){
	    		    actMsg = "相關資料寫入資料庫成功";
	    		}
    	    }
        	rd = application.getRequestDispatcher( nextPgName );        
    	}else if(act.equals("Update")){
    	    actMsg = UpdateDB(request,bank_no,seq_no,lguser_id,lguser_name);
    	    if("Y".equals(actMsg)){
    	        //操作歷程寫入log
    	        actMsg = this.InsertWlXOPERATE_LOG(request,lguser_id,program_id,bank_no,"","U");
    	        if("Y".equals(actMsg)){
	    		    actMsg = "相關資料寫入資料庫成功";
	    		}
    	    }
        	rd = application.getRequestDispatcher( nextPgName );        
    	}else if(act.equals("Delete")){
    	    actMsg = DeleteDB(request,bank_no,seq_no,lguser_id,lguser_name);
    	    if("Y".equals(actMsg)){
    	      	//操作歷程寫入log
        		this.InsertWlXOPERATE_LOG(request,lguser_id,program_id,bank_no,"","D");
    	        if("Y".equals(actMsg)){
	    		    actMsg = "相關資料刪除成功";
	    		}
    	    }
    	  	
        	rd = application.getRequestDispatcher( nextPgName );        
    	}else if(act.equals("Abdicate")){
    	    actMsg = AbdicateDB(request,bank_no,seq_no,lguser_id,lguser_name);
    	    if("Y".equals(actMsg)){
    	        //操作歷程寫入log
    	        actMsg = this.InsertWlXOPERATE_LOG(request,lguser_id,program_id,bank_no,"","U");
    	        if("Y".equals(actMsg)){
	    		    actMsg = "執行卸任成功";
	    		}
    	    }
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
	private final static String program_id = "FX003W";
    private final static String nextPgName = "/pages/ActMsg.jsp";    
    private final static String EditPgName = "/pages/"+program_id+"_Edit.jsp";    
    private final static String ListPgName = "/pages/"+program_id+"_List.jsp";        
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    
    //94.04.01 add 主key更改為bank_no+seq_no by 2295
    //102.03.27 add 增加順位排序 by 2295
    private List getWLX04(String bank_no,String seq_no){
    		//查詢條件    		
    		List paramList =new ArrayList() ;
    		String sqlCmd = "select * from WLX04,cdshareno where bank_no=? ";
    		paramList.add(bank_no) ;
    		if(!seq_no.equals("")){			
    			sqlCmd = sqlCmd + " and seq_no= ? ";
    			paramList.add(seq_no) ;
    		}
    		sqlCmd = sqlCmd + "and wlx04.POSITION_CODE = cdshareno.CMUSE_ID and cdshareno.CMUSE_DIV='008'";
    		sqlCmd =sqlCmd + " order by position_code,rank";	
    		
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"POSITION_CODE,birth_date,induct_date,abdicate_date,period_start,period_end,update_date,rank,appointed_num,seq_no");            
            return dbData;
    } 
    //94.04.01 add 主key更改為bank_no+seq_no by 2295
    public String InsertDB(HttpServletRequest request,String bank_no,String lguser_id,String lguser_name) throws Exception{    	
		StringBuffer sqlCmd = new StringBuffer();		
		String errMsg="";		
		String name=((String)request.getParameter("NAME")==null)?"":(String)request.getParameter("NAME");
		String birth_date=(String)request.getParameter("BIRTH_DATE");
		String rank=((String)request.getParameter("RANK") == null || ((String)request.getParameter("RANK")).equals(""))?"0" : (String)request.getParameter("RANK");
		String passport_area=((String)request.getParameter("PASSPORT_AREA")==null)?"":(String)request.getParameter("PASSPORT_AREA");
		String passport_no=((String)request.getParameter("PASSPORT_NO")==null)?"":(String)request.getParameter("PASSPORT_NO");
		String induct_date=(String)request.getParameter("INDUCT_DATE");
		String abdicate_code=(String)request.getParameter("ABDICATE_CODE");
		String abdicate_date=(String)request.getParameter("ABDICATE_DATE");
		String appointed_num=((String)request.getParameter("APPOINTED_NUM") == null || ((String)request.getParameter("APPOINTED_NUM")).equals(""))?"0":(String)request.getParameter("APPOINTED_NUM");
		String period_start=(String)request.getParameter("PERIOD_START");
		String period_end=(String)request.getParameter("PERIOD_END");
		String sex=((String)request.getParameter("SEX")==null)?"":(String)request.getParameter("SEX");
		String degree=((String)request.getParameter("DEGREE")==null)?"":(String)request.getParameter("DEGREE");
		String background=((String)request.getParameter("BACKGROUND")==null)?"":(String)request.getParameter("BACKGROUND");
		String professional="";
		String telno=((String)request.getParameter("TELNO")==null)?"":(String)request.getParameter("TELNO");;
		String finance_exp="";
		String email=((String)request.getParameter("EMAIL")==null)?"":(String)request.getParameter("EMAIL");;
		String user_id=lguser_id;
	    String user_name=lguser_name;	    
		String seq_no="";
	    String id_code=((String)request.getParameter("ID_CODE")==null)?"N":(String)request.getParameter("ID_CODE");		
	    String position_code = ( request.getParameter("POSITION_CODE")==null ) ? "" : (String)request.getParameter("POSITION_CODE");			
	    String id = ( request.getParameter("ID")==null ) ? "" : Utility.encode((String)request.getParameter("ID"));			

		
		try {
				//List updateDBSqlList = new LinkedList();
				List paramList =new ArrayList() ;
				List dbData = DBManager.QueryDB_SQLParam("SELECT to_char(wlx04_seqno.NEXTVAL) seq_no FROM DUAL",paramList,"");
                seq_no = (String)((DataObject)dbData.get(0)).getValue("seq_no");					 
				paramList.add(bank_no) ;
				paramList.add(position_code);
				paramList.add(id) ;
				paramList.add(name);
		        paramList.add(birth_date);  
		        paramList.add(rank );
		        paramList.add(passport_area ); 
		        paramList.add(passport_no ); 
		        paramList.add(induct_date); 
		        paramList.add(abdicate_code ); 
		        paramList.add(abdicate_date);  
		        paramList.add(appointed_num  ); 
		        paramList.add(period_start );  
		        paramList.add(period_end); 
		        paramList.add(sex );
		        paramList.add(degree ); 
		        paramList.add(background );
		        paramList.add(professional);
		        paramList.add(telno );
		        paramList.add(finance_exp);					       
		        paramList.add(email );
		        paramList.add(user_id );
		        paramList.add(user_name );
		        paramList.add(seq_no);
		        paramList.add(id_code); 
				sqlCmd.append("INSERT INTO WLX04 VALUES (?,?,?,?,to_date(?,'YYYY/MM/DD'),");
				sqlCmd.append("?,?,?,to_date(?,'YYYY/MM/DD'),?,");
				sqlCmd.append("to_date(?,'YYYY/MM/DD'),?,to_date(?,'YYYY/MM/DD'),to_date(?,'YYYY/MM/DD'),?,");
			    sqlCmd.append("?,?,?,?,?,?,?,?,sysdate,?,?)");
					       /*+ ",'" + name + "'" 
					       + ",to_date('"+birth_date+"','YYYY/MM/DD')"  
					       + ","+rank 
					       + ",'" + passport_area + "'" 
					       + ",'" + passport_no + "'" 
					       + ",to_date('"+induct_date+"','YYYY/MM/DD')" 
					       + ",'" + abdicate_code + "'" 
					       + ",to_date('"+abdicate_date+"','YYYY/MM/DD')"  
					       + ","+appointed_num   
					       + ",to_date('"+period_start+"','YYYY/MM/DD')"  
					       + ",to_date('"+period_end+"','YYYY/MM/DD')"  
					       + ",'" + sex + "'" 
					       + ",'" + degree + "'" 
					       + ",'" + background +"'"
					       + ",'" + professional + "'"
					       + ",'" + telno + "'"
					       + ",'" + finance_exp + "'"					       
					       + ",'" + email + "'"
					       + ",'" + user_id +"'"
					       + ",'" + user_name + "'"
					       + ",sysdate"
					       + ","+seq_no
					       + ",'"+id_code+"')"); 		           		 	
					       
		            updateDBSqlList.add(sqlCmd); 	*/
		            	  		            
					if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){					 
						errMsg = errMsg + "Y";					
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗";
					}
    	   		
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";					
		}	

		return errMsg;
	} 
	//94.04.01 add 主key更改為bank_no+seq_no by 2295
	//102.12.18 fix 更新理監事ID資料,若無修改ID時,存入DB會變成已mask過的資料 by 2295
	public String UpdateDB(HttpServletRequest request,String bank_no,String seq_no,String lguser_id,String lguser_name) throws Exception{    	
		StringBuffer sqlCmd = new StringBuffer();
		List paramList =new ArrayList() ;
		String errMsg="";		
		String name=((String)request.getParameter("NAME")==null)?"":(String)request.getParameter("NAME");
		String birth_date=(String)request.getParameter("BIRTH_DATE");
		String rank=((String)request.getParameter("RANK") == null || ((String)request.getParameter("RANK")).equals(""))?"0" : (String)request.getParameter("RANK");
		String passport_area=((String)request.getParameter("PASSPORT_AREA")==null)?"":(String)request.getParameter("PASSPORT_AREA");
		String passport_no=((String)request.getParameter("PASSPORT_NO")==null)?"":(String)request.getParameter("PASSPORT_NO");
		String induct_date=(String)request.getParameter("INDUCT_DATE");
		String abdicate_code=(String)request.getParameter("ABDICATE_CODE");
		String abdicate_date=(String)request.getParameter("ABDICATE_DATE");
		String appointed_num=((String)request.getParameter("APPOINTED_NUM") == null || ((String)request.getParameter("APPOINTED_NUM")).equals(""))?"0":(String)request.getParameter("APPOINTED_NUM");
		String period_start=(String)request.getParameter("PERIOD_START");
		String period_end=(String)request.getParameter("PERIOD_END");
		String sex=((String)request.getParameter("SEX")==null)?"":(String)request.getParameter("SEX");
		String degree=((String)request.getParameter("DEGREE")==null)?"":(String)request.getParameter("DEGREE");
		String background=((String)request.getParameter("BACKGROUND")==null)?"":(String)request.getParameter("BACKGROUND");
		String professional="";
		String telno=((String)request.getParameter("TELNO")==null)?"":(String)request.getParameter("TELNO");;
		String finance_exp="";
		String email=((String)request.getParameter("EMAIL")==null)?"":(String)request.getParameter("EMAIL");;
		String user_id=lguser_id;
	    String user_name=lguser_name;	    
		String id_code=((String)request.getParameter("ID_CODE")==null)?"N":(String)request.getParameter("ID_CODE");			   
	    String position_code = ( request.getParameter("POSITION_CODE")==null ) ? "" : (String)request.getParameter("POSITION_CODE");			
	    String id = ( request.getParameter("ID")==null ) ? "" : Utility.encode((String)request.getParameter("ID"));			
		//102.12.18 add ==============================================================================================================
	    String ui_ID = 	( request.getParameter("ID")==null ) ? "" : (String)request.getParameter("ID");//UI上所key的ID	
	    String encode_ID = ( request.getParameter("encode_ID")==null ) ? "" : (String)request.getParameter("encode_ID");//DB取得的ID	
	    
		System.out.println("UI.id="+(String)request.getParameter("ID"));
		System.out.println("decode.id="+Utility.decode(encode_ID));
		System.out.println("id.src="+id);
		if(ui_ID.indexOf("****") != -1){ //UI上的已經是mask過的ID,且這次無變更ID資料
			id = encode_ID;
		}	
		System.out.println("id.after="+id);
		//============================================================================================================================
		
		
		try {
				//List updateDBSqlList = new LinkedList();
				sqlCmd.append("SELECT * FROM WLX04 WHERE bank_no=? AND seq_no=?" );
				paramList.add(bank_no) ;
				paramList.add(seq_no);
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"rank,appointed_num,birth_date,induct_date,abdicate_date,period_start,period_end,seq_no");		 			    
				System.out.println("WLX04.size="+data.size());
				
				if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法修改<br>";
				}else{    
				    //95.06.05 add 寫入WLX04_LOG ===================================================================
				    sqlCmd.setLength(0) ;
				    paramList.clear() ;
                    sqlCmd.append(" INSERT INTO WLX04_LOG "
                           + " select BANK_NO,POSITION_CODE,ID,NAME,BIRTH_DATE,RANK,PASSPORT_AREA,PASSPORT_NO,INDUCT_DATE,"
                           + " 	      ABDICATE_CODE,ABDICATE_DATE,APPOINTED_NUM,PERIOD_START,PERIOD_END,SEX,DEGREE,BACKGROUND,"
                           + "		  PROFESSIONAL,TELNO,FINANCE_EXP,EMAIL,USER_ID,USER_NAME,UPDATE_DATE,SEQ_NO,ID_CODE,"
                           + "       ? ,?,sysdate,'U'"
						   + " from WLX04"
						   + " where bank_no=? and seq_no= ? ");
                    paramList.add(user_id) ;
                    paramList.add(user_name);
                    paramList.add(bank_no);
                    paramList.add(seq_no);
					//updateDBSqlList.add(sqlCmd);
					this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList) ;
					//==================================================================================================
				    sqlCmd.setLength(0) ;
					paramList.clear() ;
				    sqlCmd.append("UPDATE WLX04 SET ");
				    sqlCmd.append(" position_code=?");paramList.add(position_code );
			        sqlCmd.append(",id=?");paramList.add(id);
			        sqlCmd.append(",id_code=?");paramList.add(id_code);
			    	sqlCmd.append(",name=?");paramList.add(name);
			        sqlCmd.append(",birth_date=to_date(?,'YYYY/MM/DD')");paramList.add(birth_date);											   					    
			        sqlCmd.append(",rank=?");paramList.add(rank);
			        sqlCmd.append(",passport_area=?");paramList.add(passport_area);
	    	        sqlCmd.append(",passport_no=?");paramList.add(passport_no);
	    	        sqlCmd.append(",induct_date=to_date(?,'YYYY/MM/DD')");paramList.add(induct_date) ;											   					    
			        sqlCmd.append(",abdicate_code=?");paramList.add(abdicate_code);											   					    
			        sqlCmd.append(",abdicate_date=to_date(?,'YYYY/MM/DD')");paramList.add(abdicate_date) ;											   					    
			        sqlCmd.append(",appointed_num=?");paramList.add(appointed_num);
			        sqlCmd.append(",period_start=to_date(?,'YYYY/MM/DD')");paramList.add(period_start) ;											   					    
			        sqlCmd.append(",period_end=to_date(?,'YYYY/MM/DD')");paramList.add(period_end);											   					    
			        sqlCmd.append(",sex=?");paramList.add(sex);
			        sqlCmd.append(",degree=?");paramList.add(degree);
			        sqlCmd.append(",background=?");paramList.add(background);
			        sqlCmd.append(",professional='',telno=?");paramList.add(telno);						   
	    	        sqlCmd.append(",finance_exp=?");paramList.add(finance_exp);
	    	        sqlCmd.append(",email=?");paramList.add(email);
	    	        sqlCmd.append(",user_id=?");paramList.add(user_id);
			        sqlCmd.append(",user_name=?");paramList.add(user_name);
			        sqlCmd.append(",update_date=sysdate");											   					    
			        sqlCmd.append(" where bank_no=?");paramList.add(bank_no);
			        sqlCmd.append(" and seq_no=? ");paramList.add(seq_no);
				    /*       + " position_code='"+position_code+"'"
				           + ",id='"+id+"'"
				           + ",id_code='"+id_code+"'"
				    	   + ",name='"+name+"'"
						   + ",birth_date=to_date('"+birth_date+"','YYYY/MM/DD')"											   					    
						   + ",rank="+rank+",passport_area='"+passport_area+"'"
				    	   + ",passport_no='"+passport_no+"'"
				    	   + ",induct_date=to_date('"+induct_date+"','YYYY/MM/DD')"											   					    
						   + ",abdicate_code='"+abdicate_code+"'"											   					    
						   + ",abdicate_date=to_date('"+abdicate_date+"','YYYY/MM/DD')"											   					    
						   + ",appointed_num="+appointed_num
						   + ",period_start=to_date('"+period_start+"','YYYY/MM/DD')"											   					    
						   + ",period_end=to_date('"+period_end+"','YYYY/MM/DD')"											   					    
						   + ",sex='"+sex+"',degree='"+degree+"',background='"+background+"',professional='',telno='"+telno+"'"						   
				    	   + ",finance_exp='"+finance_exp+"',email='"+email+"',user_id='"+user_id+"'"
						   + ",user_name='"+user_name+"',update_date=sysdate"											   					    
						   + " where bank_no='"+bank_no+"' and seq_no="+seq_no);				            		 						       						    
						   
		            updateDBSqlList.add(sqlCmd); 		            		            			  
		            */	
					if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){					 
						errMsg = errMsg + "Y";					
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗";
					}
    	   		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";						
		}	

		return errMsg;
	} 
    //94.04.01 add 主key更改為bank_no+seq_no by 2295
    public String DeleteDB(HttpServletRequest request,String bank_no,String seq_no,String lguser_id,String lguser_name) throws Exception{    	
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList() ;
		String errMsg="";		
		String user_id=lguser_id;
	    String user_name=lguser_name;	    
		
		
		try {
				//List updateDBSqlList = new LinkedList();
				sqlCmd.append("SELECT * FROM WLX04 WHERE bank_no=? AND seq_no=?");			
				paramList.add(bank_no) ;
				paramList.add(seq_no);
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"rank,appointed_num,birth_date,induct_date,abdicate_date,period_start,period_end,seq_no");		 			    
				System.out.println("WLX04.size="+data.size());
				
				if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法刪除<br>";
				}else{    
				    //95.06.05 add 寫入WLX04_LOG ===================================================================
				    sqlCmd.setLength(0);
				    paramList.clear() ;
                    sqlCmd.append(" INSERT INTO WLX04_LOG "
                           + " select BANK_NO,POSITION_CODE,ID,NAME,BIRTH_DATE,RANK,PASSPORT_AREA,PASSPORT_NO,INDUCT_DATE,"
                           + " 	      ABDICATE_CODE,ABDICATE_DATE,APPOINTED_NUM,PERIOD_START,PERIOD_END,SEX,DEGREE,BACKGROUND,"
                           + "		  PROFESSIONAL,TELNO,FINANCE_EXP,EMAIL,USER_ID,USER_NAME,UPDATE_DATE,SEQ_NO,ID_CODE,"
                           + "        ?,?,sysdate,'D'"
						   + " from WLX04"
						   + " where bank_no=? and seq_no=?");
                    paramList.add(user_id) ;
                    paramList.add(user_name);
                    paramList.add(bank_no);
                    paramList.add(seq_no);
					//updateDBSqlList.add(sqlCmd); 		            	
					//==================================================================================================
				    sqlCmd.setLength(0) ;
					paramList.clear() ;
				    sqlCmd.append(" delete WLX04 where bank_no=? AND seq_no= ? ");
				    paramList.add(bank_no) ;
				    paramList.add(seq_no);
				    
		            //updateDBSqlList.add(sqlCmd); 		            		            				              
		            
					if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList) ){					 
						errMsg = errMsg + "Y";					
					}else{
				   		errMsg = errMsg + "相關資料刪除失敗";
					}
    	   		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料刪除失敗";							
		}	

		return errMsg;
	} 
	//94.04.01 add 主key更改為bank_no+seq_no by 2295
	public String AbdicateDB(HttpServletRequest request,String bank_no,String seq_no,String lguser_id,String lguser_name) throws Exception{    	
		StringBuffer sqlCmd = new StringBuffer();		
		String errMsg="";		
		String abdicate_code=(String)request.getParameter("ABDICATE_CODE");
		String abdicate_date=(String)request.getParameter("ABDICATE_DATE");
		String user_id=lguser_id;
	    String user_name=lguser_name;	    
		String position_code = ( request.getParameter("POSITION_CODE")==null ) ? "" : (String)request.getParameter("POSITION_CODE");			
	    String id = ( request.getParameter("encode_ID")==null ) ? "" : (String)request.getParameter("encode_ID");			
	  
		
		try {
				//List updateDBSqlList = new LinkedList();
				List paramList = new ArrayList();
				sqlCmd.append("SELECT * FROM WLX04 WHERE bank_no=? AND position_code=? and id =? "); 
				paramList.add(bank_no) ;
				paramList.add(position_code) ;
				paramList.add(id);
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"rank,appointed_num,birth_date,induct_date,abdicate_date,period_start,period_end");		 			    
				System.out.println("WLX04.size="+data.size());
				
				if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法卸任<br>";
				}else{    
				    //95.06.05 add 寫入WLX04_LOG ===================================================================
				    sqlCmd.setLength(0) ;
				    paramList.clear() ;
                    sqlCmd.append(" INSERT INTO WLX04_LOG "
                           + " select BANK_NO,POSITION_CODE,ID,NAME,BIRTH_DATE,RANK,PASSPORT_AREA,PASSPORT_NO,INDUCT_DATE,"
                           + " 	      ABDICATE_CODE,ABDICATE_DATE,APPOINTED_NUM,PERIOD_START,PERIOD_END,SEX,DEGREE,BACKGROUND,"
                           + "		  PROFESSIONAL,TELNO,FINANCE_EXP,EMAIL,USER_ID,USER_NAME,UPDATE_DATE,SEQ_NO,ID_CODE,"
                           + "        ?,?,sysdate,'A'"//abdicate
						   + " from WLX04"
						   + " where bank_no=? and position_code=? and id=? ");
                    paramList.add(user_id) ;
                    paramList.add(user_name);
                    paramList.add(bank_no);
                    paramList.add(position_code);
                    paramList.add(id);
					//updateDBSqlList.add(sqlCmd);
					this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList) ;
					//==================================================================================================
				    sqlCmd.setLength(0) ;
					paramList.clear() ;
				    sqlCmd.append(" UPDATE WLX04 SET abdicate_code=?,abdicate_date=to_date(?,'YYYY/MM/DD')"			
				    	   + " where bank_no=? and position_code=? and id=? ");
				    paramList.add(abdicate_code) ;
				    paramList.add(abdicate_date) ;
				    paramList.add(bank_no);
				    paramList.add(position_code);
				    paramList.add(id);
		            //updateDBSqlList.add(sqlCmd); 		            		            				            	  
		            
					if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){					 
						errMsg = errMsg + "Y";					
					}else{
				   		errMsg = errMsg + "執行卸任失敗";
					}
    	   		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "執行卸任失敗";								
		}	

		return errMsg;
	}
	public String InsertWlXOPERATE_LOG(HttpServletRequest request,String lguser_id,String program_id,String pbank_no,String bank_no,String update_type) throws Exception{    	
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList() ;
		String errMsg="";
		String position_code=( request.getParameter("POSITION_CODE")==null ) ? "" : (String)request.getParameter("POSITION_CODE");
		String seq_no=( request.getParameter("seq_no")==null ) ? "" : (String)request.getParameter("seq_no");
	    String upd_name=( request.getParameter("NAME")==null ) ? "" : (String)request.getParameter("NAME");
	    try {
	        sqlCmd.append("select name from WLX04 WHERE bank_no=? and seq_no=?  ");
			paramList.add(pbank_no) ;
			paramList.add(seq_no) ;
		    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"name");		 			    
			System.out.println("WLX01_M.size()="+data.size());
			if(data.size() > 0 && "".equals(upd_name)){
			    upd_name=Utility.getTrimString(((DataObject)data.get(0)).getValue("name"));
			}
			sqlCmd.setLength(0) ;
			paramList.clear() ;
	        sqlCmd.append(" INSERT INTO WlXOPERATE_LOG(muser_id,use_Date,program_id,ip_address,pbank_no,bank_no,position_code,upd_name,update_type)");
	        sqlCmd.append("                     VALUES(?,sysdate,?,?,?,?,?,?,?) ");
	        paramList.add(lguser_id);
	        paramList.add(program_id);
	        paramList.add(request.getRemoteAddr());//ipAddress
	        paramList.add(pbank_no);//總機構代號
	        paramList.add(bank_no);//分支機構代號
	        paramList.add(position_code);//異動職位(高階主管/負責人/理監事) 
	        paramList.add(upd_name);//異動姓名(高階主管/負責人/理監事) 
	        paramList.add(update_type);//操作類別 I-新增，U-異動，D-刪除，Q-明細，P-列印
	        if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){
				errMsg = errMsg + "Y";					
			}else{
			    errMsg = errMsg + "相關資料寫入log失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
			}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入log失敗";						
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