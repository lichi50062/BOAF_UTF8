<%
// 93.12.20 add 權限檢核 by 2295
// 93.12.20 fix 若有已點選的tbank_no,則以已點選的tbank_no為主 by 2295
//          add 設定異動者資訊
// 93.12.23 add 超過登入時間,請重新登入 by 2295
// 94.04.01 add 同一職務不可有一人以上擔任 by 2295
//          add 主key更改為bank_no+seq_no by 2295
// 94.04.01 fix 分支機構都已辦理裁撤,才能將總機構裁撤 by 2295
// 94.04.06 fix 理監事人數只統計未卸任者 by 2295
// 94.04.12 add abdicate_code is null 也是未裁撤 by 2295
// 94.04.12 fix 不為裁撤時,將bn01.BN_TYPE='1'
// 94.12.15 add 將異動的日期同時更新至『WLX01(總機構基本資料維護)』檔 by 2295         	
// 95.05.26 add 從事信用業務(存款、放款及農貸轉放等信用業務)之員工人數 by 2295
// 95.06.05 add 將異動資料寫入WLX01_LOG/WLX01_M_LOG/WLXC01_WM_LOG by 2295
// 95.08.25 add 分支機構總人數只統計未被裁撤的分支機構 by 2295
// 99.12.03 fix sqlInjection by 2808
//100.01.26 fix bug by 2295
//100.02.18 fix 無法卸任問題  by 2295
//100.03.22 fix 增加排序欄位abdicate_date  by 2479
//101.06.13 fix 增加機構英文欄位、郵遞區號 by2968
//102.02.05 fix 拿掉修改english欄位 by 2295
//102.04.24 add idn加解密  by2968
//102.06.27 add 操作歷程寫入log by2968
//102.12.18 fix 更新高階主管ID資料,若無修改ID時,存入DB會變成已mask過的資料 by 2295
//103.01.17 add 稽核人員基本資料頁
//103.01.22 add 農漁會信用部更改總機構相關資料後,回查詢頁 by 2295
//103.04.01 add 高階主管未卸任的資料,才檢核是否已建檔 by 2295
//103.05.30 fix 調整原稽核人員新增時,寫入空白 by 2295    
//103.06.24 fix 異動者資訊.增加區分99/100 by 2295
//104.06.24 add 人員配置情形 by 2968
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@include file="common.jsp"%>
<%
    
	RequestDispatcher rd = null;
	String actMsg = "";	
	String alertMsg = "";	
	String webURL = "";	
	boolean doProcess = false;	
	System.out.println("FX001W.ContentLength="+request.getContentLength());
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
	String nowact = ( request.getParameter("nowact")==null ) ? "" : (String)request.getParameter("nowact");		
	String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? "" : (String)request.getParameter("S_YEAR");		
	String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? "" : (String)request.getParameter("S_MONTH");		
	//String bank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");		
	//bank_no = ( request.getParameter("bank_code")==null ) ? bank_no : (String)request.getParameter("bank_code");			
	String seq_no = ( request.getParameter("seq_no")==null ) ? "" : (String)request.getParameter("seq_no");			
	//String position_code = ( request.getParameter("position_code")==null ) ? "" : (String)request.getParameter("position_code");			
	//String id = ( request.getParameter("id")==null ) ? "" : (String)request.getParameter("id");			
	System.out.println("act="+act);	
	System.out.println("nowact="+nowact);	
	//登入者資訊	
	String lguser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		
	String lguser_name = ( session.getAttribute("muser_name")==null ) ? "" : (String)session.getAttribute("muser_name");		
	String lguser_type = ( session.getAttribute("muser_type")==null ) ? "" : (String)session.getAttribute("muser_type");						
	//======================================================================================================================
	//fix 93.12.20 若有已點選的tbank_no,則以已點選的tbank_no為主============================================================
	String bank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");				
	String nowtbank_no =  ( request.getParameter("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");			
	if(!nowtbank_no.equals("")){
	   session.setAttribute("nowtbank_no",nowtbank_no);//將已點選的tbank_no寫入session	   
	}   
	bank_no = ( session.getAttribute("nowtbank_no")==null ) ? bank_no : (String)session.getAttribute("nowtbank_no");			
	//=======================================================================================================================
	S_YEAR = (S_YEAR.equals("")) ? "": String.valueOf(Integer.parseInt(S_YEAR));	
	S_MONTH = (S_MONTH.equals("")) ? "": String.valueOf(Integer.parseInt(S_MONTH));					
    //================================================================================================						
    String bank_type = ( session.getAttribute("bank_type")==null ) ? "" : (String)session.getAttribute("bank_type");   
    if(!Utility.CheckPermission(request,"FX001W")){//無權限時,導向到LoginError.jsp       
        rd = application.getRequestDispatcher( LoginErrorPgName );        
    }else{            
    	//set next jsp 	
    	if(act.equals("new")){    	    
    	    List dbData = getWLX01(bank_no);
    	    List dbDataWLX01_M = getWLX01_M(bank_no,"");    	    
    	    List dbDataWLX01_WM = getWLX01_WM(bank_no,"","");    	   
    	    List dbDataWLX01_Count = getWLX01_Count(bank_no);
    	    List dbDataWLX01_Audit = getWLX01_Audit(bank_no,"");
    	    //94.04.07 add 取得該分支機構的日期 ====================================================
    	    List paramList = new ArrayList() ;
    	    paramList.add(bank_no) ;
    	    List dbDataWLX02 = DBManager.QueryDB_SQLParam("select cancel_date from wlx02 where tbank_no=? and cancel_no='Y'",paramList,"cancel_date");
    	    String wlx02date = "";
    	    if(dbDataWLX02 != null && dbDataWLX02.size() != 0){
    	       for(int i=0;i<dbDataWLX02.size();i++){    	           
    	           wlx02date += (((DataObject)dbDataWLX02.get(i)).getValue("cancel_date")).toString().substring(0, 10);
    	           if( (i+1) < dbDataWLX02.size()){
    	              wlx02date += ",";
    	           }
    	       }    
    	       wlx02date = wlx02date.replace('-','/');	
    	       System.out.println("wlx02date="+wlx02date);       
    	       request.setAttribute("wlx02date",wlx02date);
    	    }
    	    //操作歷程寫入log
    	    this.InsertWlXOPERATE_LOG(request,lguser_id,program_id,bank_no,"","Q");
    	    //====================================================================================
    	    request.setAttribute("WLX01",dbData);
    	    request.setAttribute("WLX01_M",dbDataWLX01_M);
    	    request.setAttribute("WLX01_WM",dbDataWLX01_WM);
    	    request.setAttribute("WLX01_Count",dbDataWLX01_Count);
    	    request.setAttribute("WLX01_Audit",dbDataWLX01_Audit);
    	    //93.12.20設定異動者資訊======================================================================
    	    //103.06.24 fix 異動者資訊.增加區分99/100 by 2295
    	    String yy = Integer.parseInt(Utility.getYear())>99  ?"100" : "99" ;
			request.setAttribute("maintainInfo","select * from WLX01 WHERE bank_no='" + bank_no+"' and m_year="+yy);								       
			//=======================================================================================================================		
        	rd = application.getRequestDispatcher( EditWLX01PgName );        
    	/*}else if(act.equals("Edit")){//總機構    	    
    	    List dbData = getWLX01(bank_no);
    	    request.setAttribute("WLX01",dbData);
    	    rd = application.getRequestDispatcher( EditWLX01PgName );*/            		
    	}else if(act.equals("Update")){//總機構修改    	        	    
       	    actMsg = UpdateWLX01(request,bank_no,lguser_id,lguser_name);
	    	if("Y".equals(actMsg)){
	       		//操作歷程寫入log
	    		actMsg = this.InsertWlXOPERATE_LOG(request,lguser_id,program_id,bank_no,"","U");
	       		if("Y".equals(actMsg)){
	       		    actMsg = "相關資料寫入資料庫成功";
	       		}
	    	}
        	rd = application.getRequestDispatcher( nextPgName +((bank_type.equals("6") || bank_type.equals("7"))?"?goPages=FX001W.jsp&act=new&bank_type="+bank_type:""));        		
    	}else if(act.equals("Revoke")){//總機構裁撤    	    
    	    actMsg = RevokeWLX01(request,bank_no,lguser_id,lguser_name);
    		if("Y".equals(actMsg)){
    		  	//操作歷程寫入log
        		//actMsg = this.InsertWlXOPERATE_LOG(request,lguser_id,program_id,bank_no,"","U");
    		  	//if("Y".equals(actMsg)){
    		    	actMsg="執行裁撤總機構成功";
    		  	//}
    		}
        	rd = application.getRequestDispatcher( nextPgName +((bank_type.equals("6") || bank_type.equals("7"))?"?goPages=FX001W.jsp&act=new&bank_type="+bank_type:""));        		           		
    	}else if(act.equals("newM")){//高階主管新增    	    
        	rd = application.getRequestDispatcher( EditWLX01_MPgName +"?act=newM");        
    	}else if(act.equals("EditM")){//高階主管編輯    	    
    	    List dbData = getWLX01_M(bank_no,seq_no);
    	    request.setAttribute("WLX01_M",dbData);
    	    //93.12.20設定異動者資訊======================================================================
			request.setAttribute("maintainInfo","select * from WLX01_M WHERE bank_no='" + bank_no+"' and seq_no="+seq_no);								       
			//=======================================================================================================================		        	
        	//操作歷程寫入log
    		this.InsertWlXOPERATE_LOG(request,lguser_id,program_id,bank_no,"","Q");
			rd = application.getRequestDispatcher( EditWLX01_MPgName +"?act=EditM");        
        }else if(act.equals("loadM")){//代入WLX04該身份証的資料  	        	   	    
            String id = ( request.getParameter("ID")==null ) ? "" : (String)request.getParameter("ID");			
       	    List dbData = loadWLX01_M(bank_no,id);          
       	    System.out.println("loadWLX01_M="+dbData.size());
       	    if(dbData.size() == 0){       	       
       	       alertMsg = "無該身分證字號資料可載入";  
       	       rd = application.getRequestDispatcher( nextPgName );        		
       	       request.setAttribute("alertMsg",alertMsg);
       	    }else{
       	       request.setAttribute("WLX01_M",dbData);
       	       rd = application.getRequestDispatcher( EditWLX01_MPgName +"?act="+nowact);                	       	    
       	    }       	        		
    	}else if(act.equals("InsertM")){//高階主管新增    	        	    
       	    actMsg = InsertWLX01_M(request,bank_no,lguser_id,lguser_name);
	    	if("Y".equals(actMsg)){
	    	  	//操作歷程寫入log
	    		actMsg = this.InsertWlXOPERATE_LOG(request,lguser_id,program_id,bank_no,"","I");
	    	  	if("Y".equals(actMsg)){
	    	  	  actMsg = "相關資料寫入資料庫成功";
	    	  	}
	    	}
       		
        	rd = application.getRequestDispatcher( nextPgName +((bank_type.equals("6") || bank_type.equals("7"))?"?goPages=FX001W.jsp&act=new&bank_type="+bank_type:""));       		
    	}else if(act.equals("UpdateM")){//高階主管修改    	        	    
       	    actMsg = UpdateWLX01_M(request,bank_no,seq_no,lguser_id,lguser_name);
	       	if("Y".equals(actMsg)){
	 	        //操作歷程寫入log
	 	        actMsg = this.InsertWlXOPERATE_LOG(request,lguser_id,program_id,bank_no,"","U");
	 	        if("Y".equals(actMsg)){
	 	            actMsg = "相關資料寫入資料庫成功";
	 	        }
	 	    }
        	rd = application.getRequestDispatcher( nextPgName +((bank_type.equals("6") || bank_type.equals("7"))?"?goPages=FX001W.jsp&act=new&bank_type="+bank_type:""));
    	}else if(act.equals("DeleteM")){//高階主管刪除    	        	    
       	    actMsg = DeleteWLX01_M(request,bank_no,seq_no,lguser_id,lguser_name);
	    	if("Y".equals(actMsg)){
	       		//操作歷程寫入log
	    		this.InsertWlXOPERATE_LOG(request,lguser_id,program_id,bank_no,"","D");
	       		if("Y".equals(actMsg)){
	       		    actMsg= "相關資料刪除成功";
	       		}
	    	}
        	rd = application.getRequestDispatcher( nextPgName +((bank_type.equals("6") || bank_type.equals("7"))?"?goPages=FX001W.jsp&act=new&bank_type="+bank_type:""));        		
    	}else if(act.equals("AbdicateM")){//高階主管卸任    	    
    	    actMsg = AbdicateWLX01_M(request,bank_no,seq_no,lguser_id,lguser_name);
    		if("Y".equals(actMsg)){
    		  	//操作歷程寫入log
	 	        actMsg = this.InsertWlXOPERATE_LOG(request,lguser_id,program_id,bank_no,"","U");
    		  	if("Y".equals(actMsg)){
    		  	  actMsg = "執行卸任成功";
    		  	}
    		}
        	rd = application.getRequestDispatcher( nextPgName +((bank_type.equals("6") || bank_type.equals("7"))?"?goPages=FX001W.jsp&act=new&bank_type="+bank_type:""));            		
       	}else if(act.equals("newWM")){//每月申報資料新增    	    
         	rd = application.getRequestDispatcher( EditWLX01_WMPgName +"?act=newWM");            	
    	}else if(act.equals("EditWM")){//每月申報資料維護    	        	   
    	    List dbData = getWLX01_WM(bank_no,S_YEAR,S_MONTH);
    	    request.setAttribute("WLX01_WM",dbData);
    	    //93.12.20設定異動者資訊======================================================================
			request.setAttribute("maintainInfo","select * from WLX01_WM WHERE bank_no='" + bank_no+"' and m_year="+S_YEAR+" and m_month="+S_MONTH);								       
			//=======================================================================================================================		        	        	
        	rd = application.getRequestDispatcher( EditWLX01_WMPgName +"?act=EditWM&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH);                
    	}else if(act.equals("InsertWM")){//每月申報資料新增    	        	    
       	    actMsg = InsertWLX01_WM(request,bank_no,S_YEAR,S_MONTH,lguser_id,lguser_name);    	    	        	    
        	rd = application.getRequestDispatcher( nextPgName );        		
    	}else if(act.equals("UpdateWM")){//每月申報資料修改    	        	    
       	    actMsg = UpdateWLX01_WM(request,bank_no,S_YEAR,S_MONTH,lguser_id,lguser_name);    	    	        	    
        	rd = application.getRequestDispatcher( nextPgName );        		
    	}else if(act.equals("DeleteWM")){//每月申報資料刪除    	        	    
       	    actMsg = DeleteWLX01_WM(request,bank_no,S_YEAR,S_MONTH,lguser_id,lguser_name);    	    	        	    
        	rd = application.getRequestDispatcher( nextPgName );        		
    	}else if(act.equals("loadWM")){//代入上個月申報資料    	        	   	    
       	    List dbData = loadWLX01_WM(bank_no);          
       	    System.out.println("loadWLX01_WM="+dbData.size());
       	    if(dbData.size() == 0){       	       
       	       alertMsg = "無上個月申報資料可代入";  
       	       rd = application.getRequestDispatcher( nextPgName );        		
       	       request.setAttribute("alertMsg",alertMsg);
       	    }else{
       	       request.setAttribute("WLX01_WM",dbData);
       	       rd = application.getRequestDispatcher( EditWLX01_WMPgName +"?act=newWM");                	       	    
       	    }
    	}else if(act.equals("newAudit")){//稽核人員新增
    	    List dbData = getWLX01(bank_no);
    	    request.setAttribute("WLX01",dbData);
        	rd = application.getRequestDispatcher( EditWLX01_AuditPgName +"?act=newAudit");        
    	}else if(act.equals("EditAudit")){//稽核人員編輯    	 
    	    List dbData = getWLX01(bank_no);
    	    List dbData1 = getWLX01_Audit(bank_no,seq_no);
    	    request.setAttribute("WLX01",dbData);
    	    request.setAttribute("WLX01_Audit",dbData1);
    	    //93.12.20設定異動者資訊======================================================================
    		request.setAttribute("maintainInfo","select * from WLX01_Audit WHERE bank_no='" + bank_no+"' and seq_no="+seq_no);								       
    		//=======================================================================================================================		        	
        	//操作歷程寫入log
    		this.InsertWlXOPERATE_LOG(request,lguser_id,program_id,bank_no,"","Q");
    		rd = application.getRequestDispatcher( EditWLX01_AuditPgName +"?act=EditAudit");        
    	}else if(act.equals("InsertAudit")){  	        	    
       	    actMsg = InsertWLX01_Audit(request,bank_no,lguser_id,lguser_name);
        	if("Y".equals(actMsg)){
        	  	//操作歷程寫入log
        		actMsg = this.InsertWlXOPERATE_LOG(request,lguser_id,program_id,bank_no,"","I");
        	  	if("Y".equals(actMsg)){
        	  	  actMsg = "相關資料寫入資料庫成功";
        	  	}
        	}
        	rd = application.getRequestDispatcher( nextPgName +((bank_type.equals("6") || bank_type.equals("7"))?"?goPages=FX001W.jsp&act=new&bank_type="+bank_type:""));        		
    	}else if(act.equals("UpdateAudit")){   	        	    
       	    actMsg = UpdateWLX01_Audit(request,bank_no,seq_no,lguser_id,lguser_name);
           	if("Y".equals(actMsg)){
     	        //操作歷程寫入log
     	        actMsg = this.InsertWlXOPERATE_LOG(request,lguser_id,program_id,bank_no,"","U");
     	        if("Y".equals(actMsg)){
     	            actMsg = "相關資料寫入資料庫成功";
     	        }
     	    }
        	rd = application.getRequestDispatcher( nextPgName +((bank_type.equals("6") || bank_type.equals("7"))?"?goPages=FX001W.jsp&act=new&bank_type="+bank_type:""));        		
    	}else if(act.equals("DeleteAudit")){	        	    
       	    actMsg = DeleteWLX01_Audit(request,bank_no,seq_no,lguser_id,lguser_name);
        	if("Y".equals(actMsg)){
           		//操作歷程寫入log
        		this.InsertWlXOPERATE_LOG(request,lguser_id,program_id,bank_no,"","D");
           		if("Y".equals(actMsg)){
           		    actMsg= "相關資料刪除成功";
           		}
        	}
        	rd = application.getRequestDispatcher( nextPgName +((bank_type.equals("6") || bank_type.equals("7"))?"?goPages=FX001W.jsp&act=new&bank_type="+bank_type:""));
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
    
    }//end doProcess
%>

<%! 
	private final static String program_id = "FX001W";
    private final static String nextPgName = "/pages/ActMsg.jsp";    
    private final static String EditWLX01PgName = "/pages/"+program_id+"_Edit.jsp";
    private final static String EditWLX01_MPgName = "/pages/"+program_id+"_EditM.jsp";
    private final static String EditWLX01_WMPgName = "/pages/"+program_id+"_EditWM.jsp";
    private final static String EditWLX01_AuditPgName = "/pages/"+program_id+"_EditAudit.jsp";
    private final static String ListPgName = "/pages/"+program_id+"_List.jsp";        
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
   
    private List getWLX01(String bank_no){
    		//查詢條件    		
    		List paramList = new ArrayList() ;
    		String yy = Integer.parseInt(Utility.getYear())>99  ?"100" : "99" ;
    		String sqlCmd = "select * from WLX01 where bank_no=? and m_year=? ";   
    		paramList.add(bank_no) ;
    		paramList.add(yy) ;
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"setup_date,chg_license_date,start_date,staff_num,credit_staff_num,credit_staff,skill_staff,manual_staff,temp_staff,open_date,cancel_date,update_date");            
            return dbData;
    }
    //94.04.01 主key更改為bank_no+seq_no 
    //100.03.22 增加排序欄位abdicate_date
    private List getWLX01_M(String bank_no,String seq_no){
    		//查詢條件    		
    		List paramList = new ArrayList() ;
    		String sqlCmd = "select * from WLX01_M,cdshareno where bank_no=? ";
    		paramList.add(bank_no) ;
    		if(!seq_no.equals("")){			
    			sqlCmd = sqlCmd + " and seq_no= ? ";
    			paramList.add(seq_no) ;
    		}
    		
    		
    		sqlCmd = sqlCmd + " and WLX01_M.POSITION_CODE = cdshareno.CMUSE_ID and cdshareno.CMUSE_DIV='005'";
    		sqlCmd =sqlCmd + " order by position_code,abdicate_date";				
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"POSITION_CODE,birth_date,induct_date,abdicate_date,rank,seq_no");            
            return dbData;
    } 
     private List getWLX01_WM(String bank_no,String s_year,String s_month){
    		//查詢條件    		
    		List paramList = new ArrayList() ;
    		String sqlCmd = "select * from WLX01_WM where bank_no=? ";
    		paramList.add(bank_no) ;
    		if(!s_year.equals("")){			
    			sqlCmd = sqlCmd + " and m_year=?";
    			paramList.add(s_year) ;
    		}
    		if(!s_month.equals("")){			
    			sqlCmd = sqlCmd + " and m_month = ? ";  
    			paramList.add(s_month) ;
    		}
    				
    		sqlCmd = sqlCmd + " order by m_year,m_month desc";				
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"m_year,m_month,push_debitcard_cnt,tran_debitcard_cnt,atm_cnt,tran_cnt,tran_amt,check_deposit_cnt,check_deposit_amt");            
            return dbData;
    } 
    
    private List loadWLX01_WM(String bank_no){
    		//查詢條件    		
    		
    		String beforedate = Utility.CountDate("30");//取目前日期的前30天
    		System.out.println("beforedate="+beforedate);
    		String date = Utility.getCHTdate(beforedate.substring(0, 10), 0);
    		String date_Y="";
    		String date_M="";
    		
    		int i = 0;
		    if(date.length() == 9) i = 1; 
		    date_Y = date.substring(0,2+i);
		    date_M = date.substring(3+i,5+i);
    		List paramList = new ArrayList() ;
    		String sqlCmd = "select * from WLX01_WM where bank_no=?"
    					  + " and m_year = ? "
    					  + " and m_month =? ";     		
    		paramList.add(bank_no) ;
    		paramList.add(date_Y) ;
    		paramList.add(date_M) ;
    		sqlCmd = sqlCmd + " order by m_year,m_month desc";				
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"m_year,m_month,push_debitcard_cnt,tran_debitcard_cnt,atm_cnt,tran_cnt,tran_amt,check_deposit_cnt,check_deposit_amt");            
            return dbData;
    }
    
     private List loadWLX01_M(String bank_no,String id){
          //查詢條件    		
          List paramList = new ArrayList() ;
    		String sqlCmd = " select * from WLX04,cdshareno "
    		              + " where bank_no=? and id=? "
    		              + " and WLX04.POSITION_CODE = cdshareno.CMUSE_ID and cdshareno.CMUSE_DIV='005'"
    		              + " order by position_code";		
    		paramList.add(bank_no) ;
    		paramList.add(id) ;
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"birth_date,induct_date,abdicate_date,rank,seq_no");            
            return dbData;
     }
     //94.04.06 fix 理監事人數只統計未卸任者 
     private List getWLX01_Count(String bank_no){
    		//查詢條件    		
    		String yy = Integer.parseInt(Utility.getYear())>99  ?"100" : "99" ;
    		List paramList = new ArrayList() ;
    		String sqlCmd = " select sum(decode(ba01.bank_no,?,0,(wlx02.staff_num))) as wlx02staff_num,"
	   					  + " sum(decode(ba01.bank_no,?,0,1)) as bn02count,"
	   					  + " sum(decode(wlx04.position_code,'1',1,0)) as wlx04_1count,"
	   					  + " sum(decode(wlx04.position_code,'2',1,0)) as wlx04_2count"
					      + " from (select * from ba01 where m_year=? )ba01,(select * from wlx02 where m_year=?)wlx02,wlx04"
					      + " where (ba01.pbank_no=? )"
						  //+ " where (ba01.pbank_no='"+bank_no+"' or ba01.bank_no='"+bank_no+"')"
  						  //+ " and ba01.bank_no=wlx02.bank_no(+)"
  						  + " and (ba01.bank_no=wlx02.bank_no(+) and (wlx02.CANCEL_NO <> 'Y' OR wlx02.CANCEL_NO IS NULL)) "//95.08.25 add 分支機構總人數只統計未被裁撤的分支機構
  					      + " and ba01.bank_no=wlx04.bank_no(+)"
  					      + " and (wlx04.abdicate_code <> 'Y' OR wlx04.abdicate_code IS NULL)";  	
    		paramList.add(bank_no) ;
    		paramList.add(bank_no) ;
    		paramList.add(yy);
    		paramList.add(yy);
    		paramList.add(bank_no) ;
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"wlx02staff_num,bn02count,wlx04_1count,wlx04_2count");            
            return dbData;
    }
     private List getWLX01_Audit(String bank_no,String seq_no){
  		List paramList = new ArrayList() ;
  		String sqlCmd = "select BANK_NO,SEQ_NO,NAME,DEPARTMENT,SETUP_DATE,SETUP_NO,FULL_TIME,PART_TIME,PART_TIME_DATE,PART_TIME_NO,TELNO,HSIEN_ID,AREA_ID,ADDR from WLX01_AUDIT where bank_no=? ";
  		paramList.add(bank_no) ;
  		if(!"".equals(seq_no)){			
  			sqlCmd = sqlCmd + " and seq_no= ? ";
  			paramList.add(seq_no) ;
  		}
         List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"bank_no,seq_no,name,department,setup_date,setup_no,full_time,part_time,part_time_date,part_time_no,telno,hsiend_id,area_id,addr");            
         return dbData;
  	}
    //100.01.26 寫入/更新資料時,增加m_year by 2295
    //103.05.30 fix 調整原稽核人員新增時,寫入空白 by 2295
    public String UpdateWLX01(HttpServletRequest request,String bank_no,String lguser_id,String lguser_name) throws Exception{    	
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList() ;
		String errMsg="";		
		//String english=((String)request.getParameter("ENGLISH")==null)?"":(String)request.getParameter("ENGLISH");;
		String setup_approval_unt=((String)request.getParameter("SETUP_APPROVAL_UNT")==null)?"":(String)request.getParameter("SETUP_APPROVAL_UNT");//93.12.21 add		
		String setup_date=(String)request.getParameter("SETUP_DATE");
		String setup_no=((String)request.getParameter("SETUP_NO")==null)?"":(String)request.getParameter("SETUP_NO");		
		String chg_license_date=(String)request.getParameter("CHG_LICENSE_DATE");
		String chg_license_no=((String)request.getParameter("CHG_LICENSE_NO")==null)?"":(String)request.getParameter("CHG_LICENSE_NO");
		String chg_license_reason=((String)request.getParameter("CHG_LICENSE_REASON")==null)?"":(String)request.getParameter("CHG_LICENSE_REASON");
		String start_date=(String)request.getParameter("START_DATE");
		String business_id=((String)request.getParameter("BUSINESS_ID")==null)?"":(String)request.getParameter("BUSINESS_ID");		
		String hsien_id_area_id=((String)request.getParameter("HSIEN_ID_AREA_ID")==null)?"":(String)request.getParameter("HSIEN_ID_AREA_ID");
		String hsien_id = hsien_id_area_id.substring(0,hsien_id_area_id.indexOf("/"));
		String area_id = hsien_id_area_id.substring(hsien_id_area_id.indexOf("/")+1,hsien_id_area_id.length());
		String addr=((String)request.getParameter("ADDR")==null)?"":(String)request.getParameter("ADDR");
		String telno=((String)request.getParameter("TELNO")==null)?"":(String)request.getParameter("TELNO");
		String fax=((String)request.getParameter("FAX")==null)?"":(String)request.getParameter("FAX");
		String email=((String)request.getParameter("EMAIL")==null)?"":(String)request.getParameter("EMAIL");
		String web_site=((String)request.getParameter("WEB_SITE")==null)?"":(String)request.getParameter("WEB_SITE");
		String center_flag=((String)request.getParameter("CENTER_FLAG")==null)?"":(String)request.getParameter("CENTER_FLAG");
		String center_no=((String)request.getParameter("CENTER_NO")==null)?"":(String)request.getParameter("CENTER_NO");
		String staff_num=((String)request.getParameter("STAFF_NUM") == null || ((String)request.getParameter("STAFF_NUM")).equals(""))?"0" : (String)request.getParameter("STAFF_NUM");
		//95.05.26 add 從事信用業務之員工人數 by 2295==================================================================================================================================================================
		String credit_staff_num=((String)request.getParameter("CREDIT_STAFF_NUM") == null || ((String)request.getParameter("CREDIT_STAFF_NUM")).equals(""))?"0" : (String)request.getParameter("CREDIT_STAFF_NUM");
		//104.06.24 add 人員配置情形 by 2968 ==================================================================================================================================================================
		String credit_staff=((String)request.getParameter("CREDIT_STAFF") == null || ((String)request.getParameter("CREDIT_STAFF")).equals(""))?"0" : (String)request.getParameter("CREDIT_STAFF");
		String skill_staff=((String)request.getParameter("SKILL_STAFF") == null || ((String)request.getParameter("SKILL_STAFF")).equals(""))?"0" : (String)request.getParameter("SKILL_STAFF");
		String manual_staff=((String)request.getParameter("MANUAL_STAFF") == null || ((String)request.getParameter("MANUAL_STAFF")).equals(""))?"0" : (String)request.getParameter("MANUAL_STAFF");
		String temp_staff=((String)request.getParameter("TEMP_STAFF") == null || ((String)request.getParameter("TEMP_STAFF")).equals(""))?"0" : (String)request.getParameter("TEMP_STAFF");
		//=============================================================================================================================================================================================================
		String it_hsien_id_area_id=((String)request.getParameter("IT_HSIEN_ID_AREA_ID")==null)?"":(String)request.getParameter("IT_HSIEN_ID_AREA_ID");
		String it_hsien_id = it_hsien_id_area_id.substring(0,it_hsien_id_area_id.indexOf("/"));
		String it_area_id = it_hsien_id_area_id.substring(it_hsien_id_area_id.indexOf("/")+1,it_hsien_id_area_id.length());
		String it_addr=((String)request.getParameter("IT_ADDR")==null)?"":(String)request.getParameter("IT_ADDR");
		String it_name=((String)request.getParameter("IT_NAME")==null)?"":(String)request.getParameter("IT_NAME");
		String it_telno=((String)request.getParameter("IT_TELNO")==null)?"":(String)request.getParameter("IT_TELNO");
		//String audit_hsien_id_area_id=((String)request.getParameter("AUDIT_HSIEN_ID_AREA_ID")==null)?"":(String)request.getParameter("AUDIT_HSIEN_ID_AREA_ID");
		//String audit_hsien_id = audit_hsien_id_area_id.substring(0,audit_hsien_id_area_id.indexOf("/"));
		//String audit_area_id = audit_hsien_id_area_id.substring(audit_hsien_id_area_id.indexOf("/")+1,audit_hsien_id_area_id.length());
		//String audit_addr=((String)request.getParameter("AUDIT_ADDR")==null)?"":(String)request.getParameter("AUDIT_ADDR");		
		//String audit_name=((String)request.getParameter("AUDIT_NAME")==null)?"":(String)request.getParameter("AUDIT_NAME");				
		//String audit_telno=((String)request.getParameter("AUDIT_TELNO")==null)?"":(String)request.getParameter("AUDIT_TELNO");		
		String flag="";		
		String open_date=(String)request.getParameter("OPEN_DATE");		
		String m2_name=((String)request.getParameter("M2_NAME")==null)?"":(String)request.getParameter("M2_NAME");		
		String hsien_div_1=((String)request.getParameter("HSIEN_DIV")==null)?"":(String)request.getParameter("HSIEN_DIV");				
		String cancel_no=(String)request.getParameter("CANCEL_NO");
		String cancel_date=(String)request.getParameter("CANCEL_DATE");
		String user_id=lguser_id;
	    String user_name=lguser_name;	    
		
		try {
				List updateDBSqlList = new LinkedList();			
				String yy = Integer.parseInt(Utility.getYear())>99  ?"100" : "99" ;				
				sqlCmd.append("SELECT * FROM WLX01 WHERE bank_no=? and m_year=?");					 
				paramList.add(bank_no) ;	 
				paramList.add(yy) ;
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");		 			    
				System.out.println("WLX01.size="+data.size());
				sqlCmd.setLength(0) ;
			    paramList.clear() ;
				if (data.size() == 0){//無資料時,Insert
				     sqlCmd.append(" INSERT INTO WLX01 ( ");
				     sqlCmd.append("BANK_NO,ENGLISH,SETUP_APPROVAL_UNT,SETUP_DATE,SETUP_NO,");
				     sqlCmd.append("CHG_LICENSE_DATE,CHG_LICENSE_NO,CHG_LICENSE_REASON,START_DATE,BUSINESS_ID,");
				     sqlCmd.append("HSIEN_ID,AREA_ID,ADDR,TELNO,FAX,EMAIL,WEB_SITE,CENTER_FLAG,CENTER_NO,");
				     sqlCmd.append("STAFF_NUM,CREDIT_STAFF_NUM,CREDIT_STAFF,SKILL_STAFF,MANUAL_STAFF,TEMP_STAFF,");
				     sqlCmd.append("IT_HSIEN_ID, IT_AREA_ID,IT_ADDR,IT_NAME,IT_TELNO,AUDIT_HSIEN_ID,AUDIT_AREA_ID,");
				     sqlCmd.append("AUDIT_ADDR,AUDIT_NAME,AUDIT_TELNO,FLAG,OPEN_DATE,M2_NAME,HSIEN_DIV_1,");
				     sqlCmd.append("CANCEL_NO,CANCEL_DATE,USER_ID,USER_NAME,UPDATE_DATE,M_YEAR ) VALUES ( ");
				     sqlCmd.append(" ?,'',?,to_date(?,'YYYY/MM/DD'),?,");
				     sqlCmd.append("to_date(?,'YYYY/MM/DD'),?,?,to_date(?,'YYYY/MM/DD'),?,");
				     sqlCmd.append("?,?,?,?,?,?,?,?,?,");
				     sqlCmd.append("?,?,?,?,?,?,");
				     sqlCmd.append("?,?,?,?,?,'','','','','',");
				     sqlCmd.append("?,to_date(?,'YYYY/MM/DD'),?,?,");
				     sqlCmd.append("?,to_date(?,'YYYY/MM/DD'),?,?,sysdate,?)");
				     paramList.add(bank_no)  ;
				     paramList.add(setup_approval_unt) ;
				     paramList.add(setup_date) ;
				     paramList.add(setup_no) ;
				     
				     paramList.add(chg_license_date) ;
				     paramList.add(chg_license_no) ;
				     paramList.add(chg_license_reason) ;
				     paramList.add(start_date) ;
				     paramList.add(business_id) ;
				     
				     paramList.add(hsien_id);
				     paramList.add(area_id);
				     paramList.add(addr);
				     paramList.add(telno) ;
				     paramList.add(fax);
				     paramList.add(email);
				     paramList.add(web_site);
				     paramList.add(center_flag);
				     paramList.add(center_no);
				     
				     paramList.add(staff_num);
				     paramList.add(credit_staff_num) ;
				     paramList.add(credit_staff) ;
				     paramList.add(skill_staff) ;
				     paramList.add(manual_staff) ;
				     paramList.add(temp_staff) ;
				     
				     paramList.add(it_hsien_id);
				     paramList.add(it_area_id) ;
					 paramList.add(it_addr) ;
					 paramList.add(it_name) ;
					 paramList.add(it_telno) ;
					
					 paramList.add(flag) ;
					 paramList.add(open_date) ;
					 paramList.add(m2_name) ;
					 paramList.add(hsien_div_1) ;
					 
					 paramList.add(cancel_no) ;
					 paramList.add(cancel_date) ;
					 paramList.add(user_id) ;
					 paramList.add(user_name);
				     paramList.add(yy);//100.01.26 add
				     
				     if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){	   				      		 
						errMsg += "相關資料寫入資料庫成功";					
					 }else{
				 	    errMsg += "相關資料寫入資料庫失敗";
					 }
				     paramList.clear() ;
				     sqlCmd.setLength(0) ;
				}else{//有資料時,Update    
				    //94.04.01 fix 分支機構都已辦理裁撤,才能將總機構裁撤==============  
				    if(cancel_no.equals("Y")){//修改時,選裁撤
				       //取得分支機構未裁撤的數目
				        sqlCmd.append( " select count(*) as notcancel_cnt from (select * from wlx01 where m_year=?)wlx01,(select * from wlx02 where m_year=?)wlx02 "				    
						       + " where wlx01.bank_no=wlx02.tbank_no "
						       + " and wlx01.bank_no=? "
					           + " and (wlx02.cancel_no is null or wlx02.cancel_no = 'N')");
				    	paramList.add(yy) ;
				    	paramList.add(yy);
					    paramList.add(bank_no) ;       
				        data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"notcancel_cnt");	
						 			    
					    if(!((((DataObject)data.get(0)).getValue("notcancel_cnt")).toString()).equals("0")){
					       //分支機構未裁撤
					       errMsg = errMsg + "請先辦理分支機構裁撤事宜<br>";
					       return errMsg;
				        }
				        //將bn01.BN_TYPE='2'==================================
				        sqlCmd.setLength(0) ;
				        paramList.clear() ;
				        sqlCmd.append(" UPDATE bn01 SET " 
		             		   + " BN_TYPE='2' " 
		             		   + ",user_id=? "
					           + ",user_name=? "
                         	   + ",update_date=sysdate "
                               + " where bank_no=? and m_year=?");
				         paramList.add(user_id)  ;
				         paramList.add(user_name) ;
				         paramList.add(bank_no) ;
				         paramList.add(yy);//100.01.26 add
				         if(!this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){	   				      		 
							errMsg += "總機構裁撤失敗";
					 	 }
				         
				         paramList.clear() ;
				         sqlCmd.setLength(0) ;
                        //updateDBSqlList.add(sqlCmd); 		            		            		                    
                        //=======================================================
				    }else{//end of cancel_no='Y'				        
				        //94.04.12 fix 不為裁撤時,將bn01.BN_TYPE='1'==================================
				        sqlCmd.append(" UPDATE bn01 SET " 
		             		   + " BN_TYPE='1' " 
		             		   + ",user_id=? "
					           + ",user_name=?"
                         	   + ",update_date=sysdate "
                               + " where bank_no=? and m_year=?" );
				        paramList.add(user_id) ;
				        paramList.add(user_name) ;
				        paramList.add(bank_no) ;
				        paramList.add(yy);//100.01.26 add
                        //updateDBSqlList.add(sqlCmd);                        
                        if(!this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){	   				      		 
							errMsg += "bn01寫入失敗";
					 	}
                        paramList.clear() ;
                        sqlCmd.setLength(0) ;
                        //=======================================================
                    } 
                    //95.06.05 add 寫入WLX01_LOG ===================================================================
                    sqlCmd.append(" INSERT INTO WLX01_LOG "
                           + " select BANK_NO,ENGLISH,SETUP_APPROVAL_UNT,SETUP_DATE,SETUP_NO,CHG_LICENSE_DATE,CHG_LICENSE_NO,"
                           + " 		  CHG_LICENSE_REASON,START_DATE,BUSINESS_ID,HSIEN_ID,AREA_ID,ADDR,TELNO,FAX,EMAIL,WEB_SITE,"
                           + " 		  CENTER_FLAG,CENTER_NO,STAFF_NUM,CREDIT_STAFF_NUM,CREDIT_STAFF,SKILL_STAFF,MANUAL_STAFF,TEMP_STAFF,IT_HSIEN_ID,IT_AREA_ID,IT_ADDR,IT_NAME,IT_TELNO,"
                           + " 		  AUDIT_HSIEN_ID,AUDIT_AREA_ID,AUDIT_ADDR,AUDIT_NAME,AUDIT_TELNO,FLAG,OPEN_DATE,M2_NAME,"
                           + " 		  HSIEN_DIV_1,CANCEL_NO,CANCEL_DATE,USER_ID,USER_NAME,UPDATE_DATE,M_YEAR,"
                           + "        ?,?,sysdate,'U'"
						   + " from WLX01"
						   + " where bank_no=? and m_year=?");
                     paramList.add(user_id)  ;
                     paramList.add(user_name) ;
                     paramList.add(bank_no) ; 
                     paramList.add(yy);//100.01.26 add                   
                     if(!this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){	   				      		 
						errMsg += "WLX01_LOG寫入失敗";
					 }
                     paramList.clear() ;
                     sqlCmd.setLength(0) ;
					//updateDBSqlList.add(sqlCmd); 		            	
					//==================================================================================================
				    sqlCmd.append("UPDATE WLX01 SET ");				    
				    //sqlCmd.append("english=?"); paramList.add( english ) ;
				    sqlCmd.append("setup_approval_unt=?"); paramList.add( setup_approval_unt); 				    	    				    	   
					sqlCmd.append(",setup_date = to_date(?,'YYYY/MM/DD')");paramList.add(setup_date) ;  
					sqlCmd.append(",setup_no=?"); paramList.add( setup_no ); 					       
					sqlCmd.append(",chg_license_date = to_date(?,'YYYY/MM/DD')");paramList.add(chg_license_date) ;  
					sqlCmd.append(",chg_license_no=?");paramList.add( chg_license_no) ; 
					sqlCmd.append(",chg_license_reason=?");paramList.add(chg_license_reason) ; 
					sqlCmd.append(",start_date = to_date(?,'YYYY/MM/DD')" );paramList.add(start_date) ;
					sqlCmd.append(",business_id=?");paramList.add( business_id); 
					sqlCmd.append(",hsien_id=?"); paramList.add( hsien_id ); 
					sqlCmd.append(",area_id=?"); paramList.add( area_id); 
					sqlCmd.append(",addr=?"); paramList.add( addr); 
					sqlCmd.append(",telno=?");paramList.add(telno ); 					       
					sqlCmd.append(",fax=?"); paramList.add( fax );
					sqlCmd.append(",email=?"); paramList.add( email); 
					sqlCmd.append(",web_site=?"); paramList.add( web_site);
					sqlCmd.append(",center_flag=?"); paramList.add( center_flag ); 
					sqlCmd.append(",center_no=?"); paramList.add( center_no ); 
					sqlCmd.append(",staff_num=?");paramList.add(staff_num); 
					sqlCmd.append(",it_hsien_id=?"); paramList.add( it_hsien_id); 
					sqlCmd.append(",it_area_id=?"); paramList.add(it_area_id);
					sqlCmd.append(",it_addr=?"); paramList.add( it_addr) ; 
					sqlCmd.append(",it_name=?"); paramList.add( it_name); 					       
					sqlCmd.append(",it_telno=?"); paramList.add( it_telno); 					       
					//sqlCmd.append(",audit_hsien_id=?"); paramList.add( audit_hsien_id); 
					//sqlCmd.append(",audit_area_id=?"); paramList.add( audit_area_id); 
					//sqlCmd.append(",audit_addr=?"); paramList.add( audit_addr); 
					//sqlCmd.append(",audit_name=?"); paramList.add( audit_name);					       
					//sqlCmd.append(",audit_telno=?"); paramList.add( audit_telno ); 					       					       
					sqlCmd.append(",flag=?"); paramList.add( flag );
					sqlCmd.append(",open_date=to_date(?,'YYYY/MM/DD')");paramList.add(open_date) ;  
					sqlCmd.append(",m2_name=?"); paramList.add( m2_name );
					sqlCmd.append(",hsien_div_1=?"); paramList.add( hsien_div_1 ) ;
					sqlCmd.append(",cancel_no=?"); paramList.add( cancel_no );
					sqlCmd.append(",cancel_date=to_date(?,'YYYY/MM/DD')");paramList.add(cancel_date) ;  
					sqlCmd.append(",user_id=?"); paramList.add( user_id );
					sqlCmd.append(",user_name=?"); paramList.add( user_name );
					sqlCmd.append(",credit_staff_num=?");paramList.add(credit_staff_num) ; 
					sqlCmd.append(",credit_staff=?");paramList.add(credit_staff) ; 
					sqlCmd.append(",skill_staff=?");paramList.add(skill_staff) ; 
					sqlCmd.append(",manual_staff=?");paramList.add(manual_staff) ; 
					sqlCmd.append(",temp_staff=?");paramList.add(temp_staff) ; 
					sqlCmd.append(",update_date=sysdate"); 		            		 					    				    
				    sqlCmd.append(" where bank_no=? ");
				    sqlCmd.append(" and m_year=? "); //100.01.26 add
				    paramList.add(bank_no) ; 
				    paramList.add(yy);//100.01.26 add
				    if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){	   				      		 
					   errMsg += "Y";					
					}else{
				 	   errMsg += "相關資料寫入資料庫失敗";
					}
				    paramList.clear() ;
				    sqlCmd.setLength(0) ;
				}		   
		        //updateDBSqlList.add(sqlCmd);
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";						
		}	

		return errMsg;
	}     
    //100.01.26異動時,加上m_year及回傳訊息 by 2295
	public String RevokeWLX01(HttpServletRequest request,String bank_no,String lguser_id,String lguser_name) throws Exception{    	
		String sqlCmd = "";		
		String errMsg="";		
		String cancel_no=(String)request.getParameter("CANCEL_NO");
		String cancel_date=(String)request.getParameter("CANCEL_DATE");
		String user_id=lguser_id;
	    String user_name=lguser_name;
	    List paramList = new ArrayList() ;
		try {
				List updateDBSqlList = new LinkedList();
				String yy = Integer.parseInt(Utility.getYear())>99  ?"100" : "99" ;
				sqlCmd = "SELECT bank_no FROM WLX01 WHERE bank_no= ? and m_year=?";
				paramList.add(bank_no) ;
				paramList.add(yy) ;
			    List data = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");		 			    
				System.out.println("WLX01.size="+data.size());
				
				paramList.clear() ;
				if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法裁撤<br>";
				}else{  
				    //94.04.01 fix 分支機構都已辦理裁撤,才能將總機構裁撤==============  
				    //取得分支機構未裁撤的數目
				    sqlCmd = " select count(*) as notcancel_cnt from (select * from wlx01 where m_year=?)wlx01,(select * from wlx02 where m_year=?)wlx02 "				    
						   + " where wlx01.bank_no=wlx02.tbank_no "
						   + " and wlx01.bank_no=? "
					       + " and (wlx02.cancel_no is null or wlx02.cancel_no = 'N')";
				    paramList.add(yy);
				    paramList.add(yy);
				    paramList.add(bank_no) ;
				    data = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"notcancel_cnt");	
						 			    
					if(!((((DataObject)data.get(0)).getValue("notcancel_cnt")).toString()).equals("0")){
					   //分支機構未裁撤
					   errMsg = errMsg + "請先辦理分支機構裁撤事宜<br>";
				    }else{
				    	paramList.clear() ;
				       //95.06.05 add 寫入WLX01_LOG ===================================================================
                       sqlCmd = " INSERT INTO WLX01_LOG "
                          	  + " select BANK_NO,ENGLISH,SETUP_APPROVAL_UNT,SETUP_DATE,SETUP_NO,CHG_LICENSE_DATE,CHG_LICENSE_NO,"
                           	  + " 		  CHG_LICENSE_REASON,START_DATE,BUSINESS_ID,HSIEN_ID,AREA_ID,ADDR,TELNO,FAX,EMAIL,WEB_SITE,"
                           	  + " 		  CENTER_FLAG,CENTER_NO,STAFF_NUM,CREDIT_STAFF_NUM,CREDIT_STAFF,SKILL_STAFF,MANUAL_STAFF,TEMP_STAFF,IT_HSIEN_ID,IT_AREA_ID,IT_ADDR,IT_NAME,IT_TELNO,"
                           	  + " 		  AUDIT_HSIEN_ID,AUDIT_AREA_ID,AUDIT_ADDR,AUDIT_NAME,AUDIT_TELNO,FLAG,OPEN_DATE,M2_NAME,"
                           	  + " 		  HSIEN_DIV_1,CANCEL_NO,CANCEL_DATE,USER_ID,USER_NAME,UPDATE_DATE,M_YEAR"
                           	  + "        ?,?,sysdate,'R'"//revoke
						   	  + " from WLX01"
						   	  + " where bank_no=? and m_year=?";		
				       paramList.add(user_id) ;
				       paramList.add(user_name) ;
				       paramList.add(bank_no) ;
				       paramList.add(yy);//100.01.26 add
					   //updateDBSqlList.add(sqlCmd);
					   if(!this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){	   				      		 
							errMsg += "WLX01寫入失敗";
					   }
					   //==================================================================================================
					   paramList.clear() ;
				       sqlCmd = " UPDATE WLX01 SET cancel_no=?,cancel_date=to_date(?,'YYYY/MM/DD')"			
				    	      + " where bank_no=? and m_year=?";	
				       paramList.add(cancel_no) ; 
				       paramList.add(cancel_date) ;
				       paramList.add(bank_no) ;
				       paramList.add(yy);//100.01.26 add
				       if(!this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){	   				      		 
							errMsg += "WLX01寫入失敗";
					   }
		               //updateDBSqlList.add(sqlCmd); 		            		            		
		               //將bn01.BN_TYPE='2'==================================
		               paramList.clear() ;
		               sqlCmd = " UPDATE bn01 SET " 
		             		  + " BN_TYPE='2' " 
		             		  + ",user_id=? "
					          + ",user_name=? "
                         	  + ",update_date=sysdate "
                              + " where bank_no=? ";				
		               paramList.add(user_id) ;
		               paramList.add(user_name) ;
		               paramList.add(bank_no) ;
                       //updateDBSqlList.add(sqlCmd);                     
                       if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){					 
							errMsg = errMsg + "Y";					
					   }else{
				   			errMsg = errMsg + "執行裁撤總機構失敗";
					   }
                       //=====================================================  
					                  
					}//分支機構已裁撤  
    	   		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "執行裁撤總機構失敗" ;			
		}	

		return errMsg;
	}
    
    //94.04.01 主key更改為bank_no+seq_no 
    public String InsertWLX01_M(HttpServletRequest request,String bank_no,String lguser_id,String lguser_name) throws Exception{    	
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList() ;
		String errMsg="";		
		String name=((String)request.getParameter("NAME")==null)?"":(String)request.getParameter("NAME");
		String birth_date=(String)request.getParameter("BIRTH_DATE");
		String degree=((String)request.getParameter("DEGREE")==null)?"":(String)request.getParameter("DEGREE");
		String sex=((String)request.getParameter("SEX")==null)?"":(String)request.getParameter("SEX");
		String telno=((String)request.getParameter("TELNO")==null)?"":(String)request.getParameter("TELNO");;
		String fax=((String)request.getParameter("FAX")==null)?"":(String)request.getParameter("FAX");;
		String induct_date=(String)request.getParameter("INDUCT_DATE");
		String background=((String)request.getParameter("BACKGROUND")==null)?"":(String)request.getParameter("BACKGROUND");
		String choose_item="";
		String rank=((String)request.getParameter("RANK") == null || ((String)request.getParameter("RANK")).equals(""))?"0" : (String)request.getParameter("RANK");
		String speciality="";
		String incharge="";
		String abdicate_code=(String)request.getParameter("ABDICATE_CODE");
		String abdicate_date=(String)request.getParameter("ABDICATE_DATE");
		String email=((String)request.getParameter("EMAIL")==null)?"":(String)request.getParameter("EMAIL");;
		String user_id=lguser_id;
	    String user_name=lguser_name;
	    String seq_no="";
	    String id_code=((String)request.getParameter("ID_CODE")==null)?"N":(String)request.getParameter("ID_CODE");		
	    String position_code = ( request.getParameter("POSITION_CODE")==null ) ? "" : (String)request.getParameter("POSITION_CODE");			
	    String id = ( request.getParameter("ID")==null ) ? "" : Utility.encode((String)request.getParameter("ID"));			

	    try {
				//List updateDBSqlList = new LinkedList();
				sqlCmd.append(" SELECT count(*) as have_cnt FROM WLX01_M "
				       + " WHERE bank_no=? "
				       + " AND position_code=? "
					   + " AND (abdicate_code <> 'Y' or abdicate_code is null)");
				paramList.add(bank_no) ;
				paramList.add(position_code) ;
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"have_cnt");		 			    
				//System.out.println("WLX01_M.size="+data.size());
				//System.out.println("have_cnt="+(((DataObject)data.get(0)).getValue("have_cnt")).toString());
				
				if( Integer.parseInt((((DataObject)data.get(0)).getValue("have_cnt")).toString()) > 0 ){
				    errMsg = errMsg + "該職務已建檔,無法新增<br>";				    
				}else{    
					sqlCmd.setLength(0) ;
					paramList.clear() ;
					sqlCmd.append("SELECT to_char(wlx01_m_seqno.NEXTVAL) seq_no FROM DUAL") ;
				    List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
                    seq_no = (String)((DataObject)dbData.get(0)).getValue("seq_no");
                    sqlCmd.setLength(0) ;
                    paramList.add(bank_no) ;
                    paramList.add(position_code) ;
                    paramList.add(id) ;
                    paramList.add(name) ;
                    paramList.add(birth_date) ;
                    paramList.add(degree) ;
                    paramList.add(sex) ;
                    paramList.add(telno) ;
                    paramList.add(fax) ;
                    paramList.add(induct_date) ;
                    paramList.add(background) ;
                    paramList.add(choose_item) ;
                    paramList.add(rank) ;
                    paramList.add(speciality) ;
                    paramList.add(incharge) ;
                    paramList.add(abdicate_code) ;
                    paramList.add(abdicate_date) ;
                    paramList.add(email) ;
                    paramList.add(user_id) ;
                    paramList.add(user_name) ;
                    paramList.add(seq_no) ;
                    paramList.add(id_code) ;
					sqlCmd.append("INSERT INTO WLX01_M VALUES (?,?,?");
					sqlCmd.append(",?,to_date(?,'YYYY/MM/DD'),?,?,?,?,to_date(?,'YYYY/MM/DD')" 
					       + ",?,?,?,?,?,?,to_date(?,'YYYY/MM/DD'),?"
					       + ",?,?,sysdate,?,?)"); 		            		 	
					       
		            //updateDBSqlList.add(sqlCmd); 		            	
		            this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList) ;
	            	//94.12.15 add 將異動的日期同時更新至『WLX01(總機構基本資料維護)』檔       
	            	sqlCmd.setLength(0) ;
	            	paramList.clear() ;
	                sqlCmd.append(" update  WLX01 "
   				           + " set update_date  = sysdate"
       					   + " where  bank_no = ?");
	                paramList.add(bank_no) ;
		            //updateDBSqlList.add(sqlCmd); 	
		            if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){
					//if(DBManager.updateDB(updateDBSqlList)){					 
						errMsg = errMsg + "Y";					
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
					}
    	   		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";						
		}	

		return errMsg;
	} 
	
	//94.04.01 主key更改為bank_no+seq_no 
	//102.12.18 fix 更新高階主管ID資料,若無修改ID時,存入DB會變成已mask過的資料 by 2295
	public String UpdateWLX01_M(HttpServletRequest request,String bank_no,String seq_no,String lguser_id,String lguser_name) throws Exception{    	
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList() ;
		String errMsg="";		
		String name=((String)request.getParameter("NAME")==null)?"":(String)request.getParameter("NAME");
		String birth_date=(String)request.getParameter("BIRTH_DATE");
		String degree=((String)request.getParameter("DEGREE")==null)?"":(String)request.getParameter("DEGREE");
		String sex=((String)request.getParameter("SEX")==null)?"":(String)request.getParameter("SEX");
		String telno=((String)request.getParameter("TELNO")==null)?"":(String)request.getParameter("TELNO");;
		String fax=((String)request.getParameter("FAX")==null)?"":(String)request.getParameter("FAX");;
		String induct_date=(String)request.getParameter("INDUCT_DATE");
		String background=((String)request.getParameter("BACKGROUND")==null)?"":(String)request.getParameter("BACKGROUND");
		String choose_item="";
		String rank=((String)request.getParameter("RANK") == null || ((String)request.getParameter("RANK")).equals(""))?"0" : (String)request.getParameter("RANK");
		String speciality="";
		String incharge="";
		String abdicate_code=(String)request.getParameter("ABDICATE_CODE");
		String abdicate_date=(String)request.getParameter("ABDICATE_DATE");
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
					 
			    sqlCmd.append(" SELECT count(*) as have_cnt FROM WLX01_M "
				       + " WHERE bank_no=?"
				       + " AND position_code=?"
				       + " AND seq_no <> ? "
					   + " AND (abdicate_code <> 'Y' or abdicate_code is null)");
			    paramList.add(bank_no) ;
			    paramList.add(position_code) ;
			    paramList.add(seq_no) ;
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"have_cnt");		 			    
				System.out.println("WLX01_M.size="+data.size());
				
				if(!((((DataObject)data.get(0)).getValue("have_cnt")).toString()).equals("0") && !abdicate_code.equals("Y")){
				    errMsg = errMsg + "該職務已建檔<br>";//103.04.01 add 未卸任的資料,才檢核是否已建檔 				    				
				}else{    
					sqlCmd.setLength(0) ;
				    paramList.clear() ;
					//95.06.05 add 寫入WLX01_M_LOG ===================================================================
                    sqlCmd.append(" INSERT INTO WLX01_M_LOG "
                           + " select BANK_NO,POSITION_CODE,ID,NAME,BIRTH_DATE,DEGREE,SEX,TELNO,FAX,INDUCT_DATE,BACKGROUND,"
                           + "		  CHOOSE_ITEM,RANK,SPECIALITY,INCHARGE,ABDICATE_CODE,ABDICATE_DATE,EMAIL,USER_ID,USER_NAME,"
                           + " 	      UPDATE_DATE,SEQ_NO,ID_CODE,"
                           + "        ?,?,sysdate,'U'"
						   + " from WLX01_M"
						   + " where bank_no=? and seq_no=?");
					paramList.add(user_id) ;
					paramList.add(user_name) ;
					paramList.add(bank_no);
					paramList.add(seq_no);
					//updateDBSqlList.add(sqlCmd);
					this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList) ;
					//==================================================================================================				 
					sqlCmd.setLength(0) ;
					paramList.clear() ;
				    sqlCmd.append("UPDATE WLX01_M SET ");
				    sqlCmd.append(" position_code=?");paramList.add(position_code);
					sqlCmd.append(",id=?");paramList.add(id);
					sqlCmd.append(",id_code=?");paramList.add(id_code);
					sqlCmd.append(",name=?");paramList.add(name);				    	   
					sqlCmd.append(",birth_date=to_date(?,'YYYY/MM/DD')");paramList.add(birth_date) ;											   					    
					sqlCmd.append(",degree=?"); paramList.add( degree); 
			        sqlCmd.append(",sex=?"); paramList.add( sex) ; 
			        sqlCmd.append(",telno=?"); paramList.add( telno) ;
			        sqlCmd.append(",fax=?"); paramList.add( fax); 
			        sqlCmd.append(",induct_date=to_date(?,'YYYY/MM/DD')");paramList.add(induct_date) ; 
			        sqlCmd.append(",background=?"); paramList.add( background);
			        sqlCmd.append(",choose_item=?"); paramList.add( choose_item);
			        sqlCmd.append(",rank=?");paramList.add(rank); 					   
			        sqlCmd.append(",speciality=?"); paramList.add( speciality);
			        sqlCmd.append(",incharge=?"); paramList.add( incharge);    
			        sqlCmd.append(",abdicate_code=?" ); paramList.add(abdicate_code ); 
			        sqlCmd.append(",abdicate_date=to_date(?,'YYYY/MM/DD')");paramList.add(abdicate_date) ;  					       
			        sqlCmd.append(",email=?"); paramList.add( email);
			        sqlCmd.append(",user_id=?"); paramList.add( user_id );
			        sqlCmd.append(",user_name=?"); paramList.add( user_name );
			        sqlCmd.append(",update_date=sysdate"); 		            		 						       
					sqlCmd.append(" where bank_no=? and seq_no=?");
					paramList.add(bank_no) ;
					paramList.add(seq_no);
				    /*
				           + " position_code='"+position_code+"'"
				           + ",id='"+id+"'"
				           + ",id_code='"+id_code+"'"
				    	   + ",name='"+name+"'"				    	   
						   + ",birth_date=to_date('"+birth_date+"','YYYY/MM/DD')"											   					    
						   + ",degree='" + degree + "'" 
					       + ",sex='" + sex + "'" 
					       + ",telno='" + telno + "'"
					       + ",fax='" + fax + "'"
					       + ",induct_date=to_date('"+induct_date+"','YYYY/MM/DD')" 
					       + ",background='" + background +"'"
					       + ",choose_item='" + choose_item+"'"
					       + ",rank="+rank 					   
					       + ",speciality='" + speciality+"'"
					       + ",incharge='" + incharge+"'"    
					       + ",abdicate_code='" + abdicate_code + "'" 
					       + ",abdicate_date=to_date('"+abdicate_date+"','YYYY/MM/DD')"  					       
					       + ",email='" + email + "'"
					       + ",user_id='" + user_id +"'"
					       + ",user_name='" + user_name + "'"
					       + ",update_date=sysdate" 		            		 						       
						   + " where bank_no='"+bank_no+"' and seq_no="+seq_no;			*/	    	   
						   
		            //updateDBSqlList.add(sqlCmd); 		            	
		            this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList);
	            	//94.12.15 add 將異動的日期同時更新至『WLX01(總機構基本資料維護)』檔
	            	paramList.clear() ;
	            	sqlCmd.setLength(0) ;
	                sqlCmd.append(" update  WLX01 "
   				           + " set update_date  = sysdate"
       					   + " where  bank_no = ? ");
	                paramList.add(bank_no) ;
		            //updateDBSqlList.add(sqlCmd); 
	            		
					//if(DBManager.updateDB(updateDBSqlList)){
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
    //94.04.01 主key更改為bank_no+seq_no 
    public String DeleteWLX01_M(HttpServletRequest request,String bank_no,String seq_no,String lguser_id,String lguser_name) throws Exception{    	
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList() ; 
		String errMsg="";		
		String user_id=lguser_id;
	    String user_name=lguser_name;
	        
		try {
				//List updateDBSqlList = new LinkedList();
				sqlCmd.append("SELECT * FROM WLX01_M WHERE bank_no=? AND seq_no= ? ");
				paramList.add(bank_no) ;
				paramList.add(seq_no);
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"rank,birth_date,induct_date,abdicate_date");		 			    
				System.out.println("WLX01_M.size="+data.size());
				
				if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法刪除<br>";
				}else{    
				    //95.06.05 add 寫入WLX01_M_LOG ===================================================================
				    sqlCmd.setLength(0) ;
				    paramList.clear() ;
                    sqlCmd.append(" INSERT INTO WLX01_M_LOG "
                           + " select BANK_NO,POSITION_CODE,ID,NAME,BIRTH_DATE,DEGREE,SEX,TELNO,FAX,INDUCT_DATE,BACKGROUND,"
                           + "		  CHOOSE_ITEM,RANK,SPECIALITY,INCHARGE,ABDICATE_CODE,ABDICATE_DATE,EMAIL,USER_ID,USER_NAME,"
                           + " 	      UPDATE_DATE,SEQ_NO,ID_CODE,"
                           + "        ?,?,sysdate,'D'"
						   + " from WLX01_M"
						   + " where bank_no=? and seq_no=?");
                    paramList.add(user_id) ;
                    paramList.add(user_name) ;
                    paramList.add(bank_no) ;
                    paramList.add(seq_no) ;
					//updateDBSqlList.add(sqlCmd);
					this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList) ;
					//==================================================================================================				 
					sqlCmd.setLength(0);
					paramList.clear() ;
				    sqlCmd.append(" delete WLX01_M where bank_no=?  and seq_no= ? ");
				    paramList.add(bank_no) ;
				    paramList.add(seq_no) ;
		            //updateDBSqlList.add(sqlCmd); 		            		            		
		            this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList) ;
		            //94.12.15 add 將異動的日期同時更新至『WLX01(總機構基本資料維護)』檔
		            sqlCmd.setLength(0) ;
		            paramList.clear( );
	                sqlCmd.append(" update  WLX01 "
   				           + " set update_date  = sysdate"
       					   + " where  bank_no = ?");
	                paramList.add(bank_no) ;
		            //updateDBSqlList.add(sqlCmd); 
		            
					//if(DBManager.updateDB(updateDBSqlList)){
					if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)) {
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
	//94.04.01 主key更改為bank_no+seq_no 	
	public String AbdicateWLX01_M(HttpServletRequest request,String bank_no,String seq_no,String lguser_id,String lguser_name) throws Exception{    	
		StringBuffer sqlCmd = new StringBuffer();		
		String errMsg="";		
		String abdicate_code=(String)request.getParameter("ABDICATE_CODE");
		String abdicate_date=(String)request.getParameter("ABDICATE_DATE");
		String user_id=lguser_id;
	    String user_name=lguser_name;
        String position_code = ( request.getParameter("POSITION_CODE")==null ) ? "" : (String)request.getParameter("POSITION_CODE");			
	    String id = ( request.getParameter("encode_ID")==null ) ? "" : (String)request.getParameter("encode_ID");			
	    List paramList = new ArrayList() ;
		try {
				List updateDBSqlList = new LinkedList();
				sqlCmd.append("SELECT * FROM WLX01_M WHERE bank_no=?  AND position_code= ? AND id= ? ");					 
				paramList.add(bank_no) ;
				paramList.add(position_code) ;
				paramList.add(id) ;
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"rank,birth_date,induct_date,abdicate_date");		 			    
				System.out.println("WLX01_M.size="+data.size());
				
				if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法卸任<br>";
				}else{    
				    //95.06.05 add 寫入WLX01_M_LOG ===================================================================
				    sqlCmd.setLength(0) ;
				    paramList.clear() ;
                    sqlCmd.append( " INSERT INTO WLX01_M_LOG "
                           + " select BANK_NO,POSITION_CODE,ID,NAME,BIRTH_DATE,DEGREE,SEX,TELNO,FAX,INDUCT_DATE,BACKGROUND,"
                           + "		  CHOOSE_ITEM,RANK,SPECIALITY,INCHARGE,ABDICATE_CODE,ABDICATE_DATE,EMAIL,USER_ID,USER_NAME,"
                           + " 	      UPDATE_DATE,SEQ_NO,ID_CODE,"
                           + "        ?,?,sysdate,'A'"
						   + " from WLX01_M"						   
						   + " where bank_no=? and position_code=? and id=? ");
                    paramList.add(user_id) ;
                    paramList.add(user_name) ;
                    paramList.add(bank_no) ;
                    paramList.add(position_code) ;
                    paramList.add(id);
                    this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList) ;
					//updateDBSqlList.add(sqlCmd); 		            	
					//==================================================================================================				 
					sqlCmd.setLength(0) ;
					paramList.clear() ;
				    sqlCmd.append(" UPDATE WLX01_M SET abdicate_code=? ,abdicate_date=to_date(?,'YYYY/MM/DD')");			
				    sqlCmd.append(" where bank_no=? and position_code=? and id=?");
					paramList.add(abdicate_code) ;			
					paramList.add(abdicate_date);
					paramList.add(bank_no) ;
					paramList.add(position_code) ;
					paramList.add(id) ;
					this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList) ;
		            //updateDBSqlList.add(sqlCmd); 		            		            		
		            
		            //94.12.15 add 將異動的日期同時更新至『WLX01(總機構基本資料維護)』檔
		            paramList.clear() ;
		            sqlCmd.setLength(0) ;
	                sqlCmd.append(" update  WLX01 "
   				           + " set update_date  = sysdate"
       					   + " where  bank_no =? ");
	                paramList.add(bank_no) ;
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
    
    
    public String InsertWLX01_WM(HttpServletRequest request,String bank_no,String m_year,String m_month,String lguser_id,String lguser_name) throws Exception{    	
		StringBuffer sqlCmd = new StringBuffer() ;
		List paramList = new ArrayList() ;
		String errMsg="";		
		
		String push_debitcard_cnt=((String)request.getParameter("PUSH_DebitCard_CNT") == null || ((String)request.getParameter("PUSH_DebitCard_CNT")).equals(""))?"0" : (String)request.getParameter("PUSH_DebitCard_CNT");
		String tran_debitcard_cnt=((String)request.getParameter("TRAN_DebitCard_CNT") == null || ((String)request.getParameter("TRAN_DebitCard_CNT")).equals(""))?"0" : (String)request.getParameter("TRAN_DebitCard_CNT");
		String atm_cnt=((String)request.getParameter("ATM_CNT") == null || ((String)request.getParameter("ATM_CNT")).equals(""))?"0" : (String)request.getParameter("ATM_CNT");
		String tran_cnt=((String)request.getParameter("TRAN_CNT") == null || ((String)request.getParameter("TRAN_CNT")).equals(""))?"0" : (String)request.getParameter("TRAN_CNT");
		String tran_amt=((String)request.getParameter("TRAN_AMT") == null || ((String)request.getParameter("TRAN_AMT")).equals(""))?"0" : (String)request.getParameter("TRAN_AMT");
		String check_deposit_cnt=((String)request.getParameter("CHECK_DEPOSIT_CNT") == null || ((String)request.getParameter("CHECK_DEPOSIT_CNT")).equals(""))?"0" : (String)request.getParameter("CHECK_DEPOSIT_CNT");
		String check_deposit_amt=((String)request.getParameter("CHECK_DEPOSIT_AMT") == null || ((String)request.getParameter("CHECK_DEPOSIT_AMT")).equals(""))?"0" : (String)request.getParameter("CHECK_DEPOSIT_AMT");		
		String user_id=lguser_id;
	    String user_name=lguser_name;
	    
		
		try {
				//List updateDBSqlList = new LinkedList();
				sqlCmd.append("SELECT m_year,m_month FROM WLX01_WM WHERE bank_no= AND m_year= ? and m_month=? "); 
					 
				paramList.add(bank_no) ;
				paramList.add(String.valueOf(Integer.parseInt(m_year))) ;
				paramList.add(String.valueOf(Integer.parseInt(m_month))) ;
				
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month");		 			    
				System.out.println("WLX01_WM.size="+data.size());
				
				if (data.size() != 0){
				    errMsg = errMsg + "此筆資料已存在無法新增<br>";
				}else{    
					sqlCmd.setLength(0) ;
					paramList.clear() ;
					sqlCmd.append("INSERT INTO WLX01_WM VALUES ('" + bank_no + "',"+String.valueOf(Integer.parseInt(m_year))+"," +String.valueOf(Integer.parseInt(m_month))
					       + "," + push_debitcard_cnt 
					       + "," + tran_debitcard_cnt 
					       + "," + atm_cnt 
					       + "," + tran_cnt 
					       + "," + tran_amt 
					       + "," + check_deposit_cnt 
					       + "," + check_deposit_amt
					       + ",'" + user_id +"'"
					       + ",'" + user_name + "'"
					       + ",sysdate)"); 		            		 	
					 paramList.add(bank_no) ;
					 paramList.add(String.valueOf(Integer.parseInt(m_year))) ;
					 paramList.add(String.valueOf(Integer.parseInt(m_month))) ;
					 paramList.add(push_debitcard_cnt) ;
					 paramList.add(tran_debitcard_cnt) ;
					 paramList.add(atm_cnt) ;
					 paramList.add(tran_cnt) ;
		            //updateDBSqlList.add(sqlCmd); 		            	
	            		
					if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){					 
						errMsg = errMsg + "相關資料寫入資料庫成功";					
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
	
	public String UpdateWLX01_WM(HttpServletRequest request,String bank_no,String m_year,String m_month,String lguser_id,String lguser_name) throws Exception{    	
		StringBuffer sqlCmd = new StringBuffer();
		List paramList =new ArrayList() ;
		String errMsg="";		
		String push_debitcard_cnt=((String)request.getParameter("PUSH_DebitCard_CNT") == null || ((String)request.getParameter("PUSH_DebitCard_CNT")).equals(""))?"0" : (String)request.getParameter("PUSH_DebitCard_CNT");
		String tran_debitcard_cnt=((String)request.getParameter("TRAN_DebitCard_CNT") == null || ((String)request.getParameter("TRAN_DebitCard_CNT")).equals(""))?"0" : (String)request.getParameter("TRAN_DebitCard_CNT");
		String atm_cnt=((String)request.getParameter("ATM_CNT") == null || ((String)request.getParameter("ATM_CNT")).equals(""))?"0" : (String)request.getParameter("ATM_CNT");
		String tran_cnt=((String)request.getParameter("TRAN_CNT") == null || ((String)request.getParameter("TRAN_CNT")).equals(""))?"0" : (String)request.getParameter("TRAN_CNT");
		String tran_amt=((String)request.getParameter("TRAN_AMT") == null || ((String)request.getParameter("TRAN_AMT")).equals(""))?"0" : (String)request.getParameter("TRAN_AMT");
		String check_deposit_cnt=((String)request.getParameter("CHECK_DEPOSIT_CNT") == null || ((String)request.getParameter("CHECK_DEPOSIT_CNT")).equals(""))?"0" : (String)request.getParameter("CHECK_DEPOSIT_CNT");
		String check_deposit_amt=((String)request.getParameter("CHECK_DEPOSIT_AMT") == null || ((String)request.getParameter("CHECK_DEPOSIT_AMT")).equals(""))?"0" : (String)request.getParameter("CHECK_DEPOSIT_AMT");		
		String user_id=lguser_id;
	    String user_name=lguser_name;
	    
		
		try {
				//List updateDBSqlList = new LinkedList();
				sqlCmd.append( "SELECT m_year,m_month FROM WLX01_WM WHERE bank_no=? AND m_year= ? and m_month=?" );				 
				paramList.add(bank_no) ;
				paramList.add(String.valueOf(Integer.parseInt(m_year)) ) ;
				paramList.add(String.valueOf(Integer.parseInt(m_month))) ;
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month");		 			    
				System.out.println("WLX01_WM.size="+data.size());
				
				if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法修改<br>";
				}else{    
				 	//95.06.05 add 寫入WLX01_WM_LOG ===================================================================
				 	sqlCmd.setLength(0) ;
				 	paramList.clear() ;
                    sqlCmd.append(" INSERT INTO WLX01_WM_LOG "
                           + " select BANK_NO,M_YEAR,M_MONTH,PUSH_DEBITCARD_CNT,TRAN_DEBITCARD_CNT,ATM_CNT,TRAN_CNT,TRAN_AMT,"
                           + "	 	  CHECK_DEPOSIT_CNT,CHECK_DEPOSIT_AMT,USER_ID,USER_NAME,UPDATE_DATE, "
                           + "        ?,?,sysdate,'U'"
						   + " from WLX01_WM"						   
						   + " where bank_no=? and m_year=? and m_month= ?");
                    paramList.add(user_id) ;
                    paramList.add(user_name);
                    paramList.add(bank_no) ;
                    paramList.add(String.valueOf(Integer.parseInt(m_year))) ;
                    paramList.add(String.valueOf(Integer.parseInt(m_month))) ;
                    this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList) ;
					//updateDBSqlList.add(sqlCmd); 		            	
					//==================================================================================================				 				
					sqlCmd.setLength(0) ;
					paramList.clear() ;
				    sqlCmd.append( "UPDATE WLX01_WM SET "
					       + " push_debitcard_cnt= ?" 
					       + ",tran_debitcard_cnt= ?"  
					       + ",atm_cnt= ?"
					       + ",tran_cnt=? "  
					       + ",tran_amt=?"  
					       + ",check_deposit_cnt=?"
					       + ",check_deposit_amt=?"				       
				    	   + ",user_id=? "
					       + ",user_name=?"
					       + ",update_date=sysdate" 		            		 						       
						   + " where bank_no=? and m_year=? and m_month= ?");				    	   
					paramList.add(push_debitcard_cnt) ;	   
					paramList.add(tran_debitcard_cnt) ;
					paramList.add(atm_cnt) ;
					paramList.add(tran_cnt) ;
					paramList.add(tran_amt) ;
					paramList.add(check_deposit_cnt) ;
					paramList.add(check_deposit_amt) ;
					paramList.add(user_id) ;
					paramList.add(user_name) ;
					paramList.add(bank_no) ;
					paramList.add(String.valueOf(Integer.parseInt(m_year))) ;
					paramList.add(String.valueOf(Integer.parseInt(m_month))) ;
		            //updateDBSqlList.add(sqlCmd); 		            	
	            	
					if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){					 
						errMsg = errMsg + "相關資料寫入資料庫成功";					
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
    
    public String DeleteWLX01_WM(HttpServletRequest request,String bank_no,String m_year,String m_month,String lguser_id,String lguser_name) throws Exception{    	
		StringBuffer sqlCmd = new StringBuffer();	
		List paramList = new ArrayList() ;
		String errMsg="";		
		String user_id=lguser_id;
	    String user_name=lguser_name;
	    
		try {
				//List updateDBSqlList = new LinkedList();
				sqlCmd.append("SELECT m_year,m_month FROM WLX01_WM WHERE bank_no=? AND m_year=? and m_month=? ");
					 
				paramList.add(bank_no) ;
				paramList.add(String.valueOf(Integer.parseInt(m_year)) ) ;
				paramList.add(String.valueOf(Integer.parseInt(m_month))) ;
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month");		 			    
				System.out.println("WLX01_WM.size="+data.size());				
				
				if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法刪除<br>";
				}else{    
				    //95.06.05 add 寫入WLX01_WM_LOG ===================================================================
				    sqlCmd.setLength(0) ;
				    paramList.clear() ;
                    sqlCmd.append(" INSERT INTO WLX01_WM_LOG "
                           + " select BANK_NO,M_YEAR,M_MONTH,PUSH_DEBITCARD_CNT,TRAN_DEBITCARD_CNT,ATM_CNT,TRAN_CNT,TRAN_AMT,"
                           + "	 	  CHECK_DEPOSIT_CNT,CHECK_DEPOSIT_AMT,USER_ID,USER_NAME,UPDATE_DATE, "
                           + "        ?,?,sysdate,'D'"
						   + " from WLX01_WM"						   
						   + " where bank_no=? and m_year=? and m_month=?");
                    paramList.add(user_id) ;
                    paramList.add(user_name) ;
                    paramList.add(bank_no) ;
                    paramList.add(String.valueOf(Integer.parseInt(m_year))) ;
                    paramList.add(String.valueOf(Integer.parseInt(m_month))) ;
                    this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList) ;
					//updateDBSqlList.add(sqlCmd); 		            	
					//==================================================================================================				 				
					sqlCmd.setLength(0);
					paramList.clear() ;
				    sqlCmd.append(" delete WLX01_WM where bank_no=? and m_year=? and m_month= ? ");
				    paramList.add(bank_no) ;
				    paramList.add(String.valueOf(Integer.parseInt(m_year))) ;
				    paramList.add(String.valueOf(Integer.parseInt(m_month))) ;
		            //updateDBSqlList.add(sqlCmd); 		            		            		
					if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){					 
						errMsg = errMsg + "相關資料刪除成功";					
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
    public String InsertWlXOPERATE_LOG(HttpServletRequest request,String lguser_id,String program_id,String pbank_no,String bank_no,String update_type) throws Exception{    	
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList() ;
		String errMsg="";
		String position_code=( request.getParameter("POSITION_CODE")==null ) ? "" : (String)request.getParameter("POSITION_CODE");
		String seq_no=( request.getParameter("seq_no")==null ) ? "" : (String)request.getParameter("seq_no");
	    String upd_name=( request.getParameter("NAME")==null ) ? "" : (String)request.getParameter("NAME");
	    try {
	        sqlCmd.append("select name from WLX01_M WHERE bank_no=? and seq_no=?  ");
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
    public String InsertWLX01_Audit(HttpServletRequest request,String bank_no,String lguser_id,String lguser_name) throws Exception{    	
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList() ;
		String errMsg ="";		
		String name =((String)request.getParameter("NAME")==null)?"":(String)request.getParameter("NAME");
		String department = ((String)request.getParameter("DEPARTMENT")==null)?"":(String)request.getParameter("DEPARTMENT");
		String setup_date =(String)request.getParameter("SETUP_DATE");
		String setup_no =((String)request.getParameter("SETUP_NO")==null)?"":(String)request.getParameter("SETUP_NO");
		String full_time =((String)request.getParameter("FULL_TIME")==null)?"":(String)request.getParameter("FULL_TIME");
		String part_time =((String)request.getParameter("PART_TIME")==null)?"":(String)request.getParameter("PART_TIME");
		String part_time_date =(String)request.getParameter("PART_TIME_DATE");
		String part_time_no =((String)request.getParameter("PART_TIME_NO")==null)?"":(String)request.getParameter("PART_TIME_NO");
		String telno =((String)request.getParameter("TELNO")==null)?"":(String)request.getParameter("TELNO");
		String hsien_id_area_id=((String)request.getParameter("HSIEN_ID_AREA_ID")==null)?"":(String)request.getParameter("HSIEN_ID_AREA_ID");
		String hsien_id = hsien_id_area_id.substring(0,hsien_id_area_id.indexOf("/"));
		String area_id = hsien_id_area_id.substring(hsien_id_area_id.indexOf("/")+1,hsien_id_area_id.length());
		String addr =((String)request.getParameter("ADDR")==null)?"":(String)request.getParameter("ADDR");
		String user_id=lguser_id;
	    String user_name=lguser_name;
	    String seq_no="";
	    try {
			sqlCmd.append("select decode(max(seq_no),null,'1',max(seq_no)+1) as seq_no from WLX01_AUDIT where bank_no=? ") ;
			paramList.add(bank_no) ;
		    List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
            seq_no = (String)((DataObject)dbData.get(0)).getValue("seq_no");
            
            sqlCmd.setLength(0) ;
            paramList.clear() ;
            paramList.add(bank_no) ;
            paramList.add(seq_no) ;
            paramList.add(name) ;
            paramList.add(department) ;
            paramList.add(setup_date) ;
            paramList.add(setup_no) ;
            paramList.add(full_time) ;
            paramList.add(part_time) ;
            paramList.add(part_time_date) ;
            paramList.add(part_time_no) ;
            paramList.add(telno) ;
            paramList.add(hsien_id) ;
            paramList.add(area_id) ;
            paramList.add(addr) ;
            paramList.add(user_id) ;
            paramList.add(user_name) ;
			sqlCmd.append("INSERT INTO WLX01_AUDIT ");
			sqlCmd.append("     VALUES ( ?,?,?,?,to_date(?,'YYYY/MM/DD')");
			sqlCmd.append("     		,?,?,?,to_date(?,'YYYY/MM/DD'),?"); 
			sqlCmd.append("     		,?,?,?,?,?,?,sysdate) ");
            this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList) ;
        	//94.12.15 add 將異動的日期同時更新至『WLX01(總機構基本資料維護)』檔       
        	sqlCmd.setLength(0) ;
        	paramList.clear() ;
            sqlCmd.append(" update  WLX01 "
			           + " set update_date = sysdate"
					   + " where  bank_no = ? ");
            paramList.add(bank_no) ;
            if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){
				errMsg = errMsg + "Y";					
			}else{
		   		errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
			}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";						
		}	

		return errMsg;
	} 
	
	//94.04.01 主key更改為bank_no+seq_no 
	public String UpdateWLX01_Audit(HttpServletRequest request,String bank_no,String seq_no,String lguser_id,String lguser_name) throws Exception{    	
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList() ;
		String errMsg="";		
		String name =((String)request.getParameter("NAME")==null)?"":(String)request.getParameter("NAME");
		String department = ((String)request.getParameter("DEPARTMENT")==null)?"":(String)request.getParameter("DEPARTMENT");
		String setup_date =(String)request.getParameter("SETUP_DATE");
		String setup_no =((String)request.getParameter("SETUP_NO")==null)?"":(String)request.getParameter("SETUP_NO");
		String full_time =((String)request.getParameter("FULL_TIME")==null)?"":(String)request.getParameter("FULL_TIME");
		String part_time =((String)request.getParameter("PART_TIME")==null)?"":(String)request.getParameter("PART_TIME");
		String part_time_date =(String)request.getParameter("PART_TIME_DATE");
		String part_time_no =((String)request.getParameter("PART_TIME_NO")==null)?"":(String)request.getParameter("PART_TIME_NO");
		String telno =((String)request.getParameter("TELNO")==null)?"":(String)request.getParameter("TELNO");
		String hsien_id_area_id=((String)request.getParameter("HSIEN_ID_AREA_ID")==null)?"":(String)request.getParameter("HSIEN_ID_AREA_ID");
		String hsien_id = hsien_id_area_id.substring(0,hsien_id_area_id.indexOf("/"));
		String area_id = hsien_id_area_id.substring(hsien_id_area_id.indexOf("/")+1,hsien_id_area_id.length());
		String addr =((String)request.getParameter("ADDR")==null)?"":(String)request.getParameter("ADDR");
		String user_id=lguser_id;
	    String user_name=lguser_name;
		try {
			sqlCmd.setLength(0) ;
			paramList.clear() ;
            paramList.add(name) ;
            paramList.add(department) ;
            paramList.add(setup_date) ;
            paramList.add(setup_no) ;
            paramList.add(full_time) ;
            paramList.add(part_time) ;
            paramList.add(part_time_date) ;
            paramList.add(part_time_no) ;
            paramList.add(telno) ;
            paramList.add(hsien_id) ;
            paramList.add(area_id) ;
            paramList.add(addr) ;
            paramList.add(user_id) ;
            paramList.add(user_name) ;
		    sqlCmd.append("UPDATE WLX01_AUDIT ");
		    sqlCmd.append("SET name=?");
			sqlCmd.append(",department=?");
			sqlCmd.append(",setup_date=to_date(?,'YYYY/MM/DD')");											   					    
			sqlCmd.append(",setup_no=?"); 
	        sqlCmd.append(",full_time=?");
	        sqlCmd.append(",part_time=?");
	        sqlCmd.append(",part_time_date=to_date(?,'YYYY/MM/DD')");	
	        sqlCmd.append(",part_time_no=?");
	        sqlCmd.append(",telno=?"); 
	        sqlCmd.append(",hsien_id=?");  
	        sqlCmd.append(",area_id=?"); 
	        sqlCmd.append(",addr=?"); 
	        sqlCmd.append(",user_id=?");
	        sqlCmd.append(",user_name=?"); 
	        sqlCmd.append(",update_date=sysdate"); 		            		 						       
			sqlCmd.append(" where bank_no=? and seq_no=?");
			paramList.add(bank_no) ;
			paramList.add(seq_no);
            this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList);
        	//94.12.15 add 將異動的日期同時更新至『WLX01(總機構基本資料維護)』檔
        	sqlCmd.setLength(0) ;
        	paramList.clear() ;
            sqlCmd.append(" update  WLX01 "
			           + " set update_date  = sysdate"
					   + " where  bank_no = ? ");
            paramList.add(bank_no) ;
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
    //94.04.01 主key更改為bank_no+seq_no 
    public String DeleteWLX01_Audit(HttpServletRequest request,String bank_no,String seq_no,String lguser_id,String lguser_name) throws Exception{    	
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList() ; 
		String errMsg="";		
		String user_id=lguser_id;
	    String user_name=lguser_name;
	        
		try {
				sqlCmd.append("SELECT * FROM WLX01_AUDIT WHERE bank_no=? AND seq_no= ? ");
				paramList.add(bank_no) ;
				paramList.add(seq_no);
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");		 			    
				System.out.println("WLX01_Audit.size="+data.size());
				
				if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法刪除<br>";
				}else{    
					sqlCmd.setLength(0);
					paramList.clear() ;
				    sqlCmd.append(" delete WLX01_AUDIT where bank_no=?  and seq_no= ? ");
				    paramList.add(bank_no) ;
				    paramList.add(seq_no) ;
		            this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList) ;
		            //94.12.15 add 將異動的日期同時更新至『WLX01(總機構基本資料維護)』檔
		            sqlCmd.setLength(0) ;
		            paramList.clear( );
	                sqlCmd.append(" update  WLX01 "
   				           + " set update_date  = sysdate"
       					   + " where  bank_no = ?");
	                paramList.add(bank_no) ;
					if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)) {
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