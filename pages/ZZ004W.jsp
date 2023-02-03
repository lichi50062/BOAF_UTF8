<%
// 93.12.29 create by 2295
// 94.01.05 fix 沒有Bank_List,把所點選的Bank_no清除 by 2295
// 94.02.01 fix 區分網路申報/MIS管理系統 by 2295
// 94.03.08 fix 刪除時,整個muser_id的都刪除 by 2295
// 94.07.28 fix MIS在修改權限時,只能更改BR/FR/TC/ZZ/AS..BOAF只能修改FX/WM/ZZ
// 94.11.15 add 030:F01_在台無住所之外國人新台幣存款表 by 2295
// 95.10.16 fix MIS在修改權限時,可再更改DS/AN/CG by 2295
// 95.11.10 fix BOAF在修改權限時,可再更改DS by 2295
// 98.08.12 add MIS在修改權限時,可再更改MC/RptFileUpload(金庫報表上傳作業) by 2295
// 98.08.12 add BOAF在修改權限時,可再更改RptFileUpload(金庫報表上傳作業) by 2295
// 99.12.09	fix sqlInjection by 2808
//100.02.16 fix 修改時無法取得程式基本功能 by 2295
//102.10.16 add MIS.WR警示報表 by 2295
//103.01.20 add BOAF.BR權限 by 2295
//105.09.06 add BOAF.TM權限 by 2295
//105.10.19 add MIS.TM/FL權限 by 2295
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
<%@ page import="java.util.Enumeration" %>
<%@ page import="java.util.HashMap" %>
<%@ page import="java.util.Hashtable" %>
<%@ page import="java.util.StringTokenizer" %>

<%
	RequestDispatcher rd = null;
	String actMsg = "";	
	String alertMsg = "";	
	String webURL = "";	
	boolean doProcess = false;	

	//取得session資料,取得成功時,才繼續往下執行===================================================
	if(session.getAttribute("muser_id") == null){//session timeout	
      System.out.println("ZZ004W login timeout");   
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
	String menu = ( request.getParameter("menu")==null ) ? "" : (String)request.getParameter("menu");						
	String nowact = ( request.getParameter("nowact")==null ) ? "" : (String)request.getParameter("nowact");				
	String id_notCreate = ( request.getParameter("MUSER_ID_NotCreate")==null ) ? "" : (String)request.getParameter("MUSER_ID_NotCreate");						
	String id_Create = ( request.getParameter("MUSER_ID_Create")==null ) ? "" : (String)request.getParameter("MUSER_ID_Create");						
	String choose_type = ( request.getParameter("choose_type")==null ) ? "" : (String)request.getParameter("choose_type");						
	String inserUser = ( request.getParameter("inserUser")==null ) ? "" : (String)request.getParameter("inserUser");						
	String deleteUser = ( request.getParameter("deleteUser")==null ) ? "" : (String)request.getParameter("deleteUser");						
	String updateUser = ( request.getParameter("updateUser")==null ) ? "" : (String)request.getParameter("updateUser");						
	System.out.println("act="+act);	
	System.out.println("nowact="+nowact);	
	System.out.println("choose_type="+choose_type);	
	System.out.println("id_Create="+id_Create);	
	System.out.println("id_notCreate="+id_notCreate);	
   
	//登入者資訊
	String lguser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		
	String lguser_name = ( session.getAttribute("muser_name")==null ) ? "" : (String)session.getAttribute("muser_name");		
	String lguser_type = ( session.getAttribute("muser_type")==null ) ? "" : (String)session.getAttribute("muser_type");					
	String tbank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");					
	session.setAttribute("nowtbank_no",null);//94.01.05 fix 沒有Bank_List,把所點選的Bank_no清除======
	String program_id = "";
	
    if(!CheckPermission(request)){//無權限時,導向到LoginError.jsp
        rd = application.getRequestDispatcher( LoginErrorPgName );        
    }else{            
    	//set next jsp 	    	
    	if(!menu.equals("")){
    	   session.setAttribute("menu",menu);    	   
    	}
    	menu = (session.getAttribute("menu") == null)?"":(String)session.getAttribute("menu");
    	System.out.println("menu='"+menu+"'");
    	if(act.equals("new") || act.equals("del") || act.equals("upd") || act.equals("getData") || act.equals("getUser")){    	   
    	   List WTT01_hasNotCreateList = null;
    	   List WTT01_hasCreateList = null;
    	   if(act.equals("new") || act.equals("getUser")){
    	      WTT01_hasNotCreateList = getWTT01_hasNotCreate_New(lguser_type,tbank_no);    	  
    	      WTT01_hasCreateList = getWTT01_hasCreate_New(lguser_type,tbank_no);
    	   }else{
    	      WTT01_hasNotCreateList = getWTT01_hasNotCreate(lguser_type,tbank_no);    	   
    	      WTT01_hasCreateList = getWTT01_hasCreate(lguser_type,tbank_no);
    	   }   
    	   
    	   if((!act.equals("new")) && choose_type.equals("MUSER_ID_NotCreate")){
    	      List WTT01_notCreate = getWTT01(id_notCreate);
    	      request.setAttribute("WTT01_notCreate",WTT01_notCreate);  
    	   }
    	   
    	   if(choose_type.equals("MUSER_ID_Create") || act.equals("getUser") || act.equals("upd")){
    	      if(act.equals("upd") && WTT01_hasCreateList != null && WTT01_hasCreateList.size() != 0){    	         
    	         id_Create = (String)((DataObject)WTT01_hasCreateList.get(0)).getValue("muser_id");
    	         nowact="upd";
    	      }   
    	      List WTT01_Create = getWTT01(id_Create);
    	      request.setAttribute("WTT01_Create",WTT01_Create);  
    	   }
    	   List WTT03_1List = null;
    	   if(act.equals("getUser") || nowact.equals("del") || nowact.equals("upd") || nowact.equals("new")){
    	     System.out.println("getuser="+id_Create);
    	     if(nowact.equals("del") || nowact.equals("new")){
    	        WTT03_1List = getWTT03_1("",id_Create,"getUser",menu);  
    	     }
    	     
    	     if(nowact.equals("upd")){
    	        WTT03_1List = getWTT03_1(lguser_type,id_Create,"getUser",menu);  
    	     } 
    	     List Report_upload = null;
    	     List Report_download = null;
    	     List Report_edit = null;
    	     List Report_query = null;
    	     if(nowact.equals("upd")){
    	        Report_upload = getWTT04_Report_update(id_Create,"1");
    	        Report_download = getWTT04_Report_update(id_Create,"2");
    	        Report_edit = getWTT04_Report_update(id_Create,"3");
    	        Report_query= getWTT04_Report_update(id_Create,"4");
    	     }else{
    	        Report_upload = getWTT04_Report(id_Create,"1");
    	        Report_download = getWTT04_Report(id_Create,"2");
    	        Report_edit = getWTT04_Report(id_Create,"3");
    	        Report_query= getWTT04_Report(id_Create,"4");
    	     }
    	     request.setAttribute("Report_upload",Report_upload);  
    	     request.setAttribute("Report_download",Report_download);  
    	     request.setAttribute("Report_edit",Report_edit);  
    	     request.setAttribute("Report_query",Report_query);      	         
    	       	         	     
    	   }else{
    	     WTT03_1List = getWTT03_1(lguser_type,lguser_id,"",menu);    	         	     
    	   }
    	    
    	   List uploadList = getWTT03_1_upload(lguser_type,lguser_id);
    	   List downloadList = getWTT03_1_download(lguser_type,lguser_id);
    	   List editList = getWTT03_1_edit(lguser_type,lguser_id);
    	   List queryList = getWTT03_1_query(lguser_type,lguser_id);
    	   
    	   request.setAttribute("WTT01_hasNotCreateList",WTT01_hasNotCreateList);  
    	   request.setAttribute("WTT01_hasCreateList",WTT01_hasCreateList);  
    	   request.setAttribute("WTT03_1List",WTT03_1List);  
    	   request.setAttribute("uploadList",uploadList); 
    	   request.setAttribute("downloadList",downloadList); 
    	   request.setAttribute("editList",editList); 
    	   request.setAttribute("queryList",queryList); 
    	   if(act.equals("getUser")){
    	      System.out.println("1111111getUser");
    	      if(act.equals("del") || nowact.equals("del")){
    	         rd = application.getRequestDispatcher( UpdatePgName +"?act=del&id_Create="+id_Create+"&id_notCreate="+id_notCreate+"&getUser=true");                	
    	      }else if(act.equals("upd") || nowact.equals("upd")){
    	         rd = application.getRequestDispatcher( UpdatePgName +"?act=upd&id_Create="+id_Create+"&id_notCreate="+id_notCreate+"&getUser=true");                	
    	      }else{
    	         rd = application.getRequestDispatcher( EditPgName +"?act=new&id_Create="+id_Create+"&id_notCreate="+id_notCreate+"&getUser=true");                	
    	      }
    	   }else{    	
    	   	  System.out.println("not getUser");      
    	      if(act.equals("del") || nowact.equals("del")){   
    	         System.out.println("go to "+UpdatePgName); 	         
    	         rd = application.getRequestDispatcher( UpdatePgName +"?act=del&menu="+menu+"&id_Create="+id_Create+"&id_notCreate="+id_notCreate);                	
    	      }else if(act.equals("upd") || nowact.equals("upd")){
    	         System.out.println("go to "+UpdatePgName); 	         
    	         rd = application.getRequestDispatcher( UpdatePgName +"?act=upd&menu="+menu+"&id_Create="+id_Create+"&id_notCreate="+id_notCreate);                	
    	      }else{
    	         System.out.println("go to "+EditPgName); 	         
    	         rd = application.getRequestDispatcher( EditPgName +"?act=new&menu="+menu+"&id_Create="+id_Create+"&id_notCreate="+id_notCreate);                	
    	      }
    	   }   	   
        }else if(act.equals("List") || act.equals("Qry")){
            List WTT01_hasCreateList = getWTT01_hasCreate(lguser_type,tbank_no);
            request.setAttribute("WTT01_hasCreateList",WTT01_hasCreateList);  
            List WTT04List = null;
            if(!id_Create.equals("")){
               if(id_Create.equals("ALL")){
                  WTT04List = getQryResult(lguser_type,tbank_no,"");    	      
               }else{ 
                  WTT04List = getQryResult(lguser_type,tbank_no,id_Create);    	      
               }  
            }   
            request.setAttribute("WTT04List",WTT04List);  
            rd = application.getRequestDispatcher( ListPgName +"?act="+act+"&id_Create="+id_Create);                	              	        	        	    	            
    	}else if(act.equals("Insert")){
    	    actMsg = InsertDB(request,inserUser,lguser_id,lguser_name);    	    	        	    
        	rd = application.getRequestDispatcher( nextPgName );         
        }else if(act.equals("Update")){
    	    actMsg = UpdateDB(request,updateUser,lguser_id,lguser_name,menu);    	    	        	    
        	rd = application.getRequestDispatcher( nextPgName );                 
        }else if(act.equals("Delete")){
            actMsg = DeleteDB(request,deleteUser,lguser_id,lguser_name);  	        	    
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
    private final static String EditPgName = "/pages/ZZ004W_Edit.jsp";    
    private final static String UpdatePgName = "/pages/ZZ004W_Update.jsp";    
    private final static String ListPgName = "/pages/ZZ004W_List.jsp";        
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    private boolean CheckPermission(HttpServletRequest request){//檢核權限    	    
    	    boolean CheckOK=false;
    	    HttpSession session = request.getSession();            
            Properties permission = ( session.getAttribute("ZZ004W")==null ) ? new Properties() : (Properties)session.getAttribute("ZZ004W");				                
            if(permission == null){
              System.out.println("ZZ004W.permission == null");
            }else{
               System.out.println("ZZ004W.permission.size ="+permission.size());
               
            }
            //只要有Query的權限,就可以進入畫面
        	if(permission != null && permission.get("Q") != null && permission.get("Q").equals("Y")){            
        	   CheckOK = true;//Query
        	}
        	return CheckOK;
    }           
    
    //取得WTT01未建立權限的user
    private List getWTT01_hasNotCreate_New(String muser_type,String tbank_no){
    		//查詢條件       
    		String yy = Integer.parseInt(Utility.getYear())>99 ?"100" :"99" ;
    		List paramList =new ArrayList() ;
  			String sqlCmd = " select a.muser_id,a.muser_name,a.bank_type,c1.cmuse_name as bank_type_name"
    					  + " ,a.tbank_no,bn01.bank_name as tbank_no_name"
    					  + " ,a.bank_no,bn02.bank_name as bank_no_name"
    					  + " ,a.subdep_id,c.cmuse_name as subdep_id_name "
					      + " from wtt01 a"
						  + " LEFT JOIN cdshareno c1 on a.bank_type = c1.cmuse_id and c1.cmuse_div='001' "
						  + " LEFT JOIN (select * from bn01 where m_year=?)bn01 on a.tbank_no = bn01.bank_no"
						  + " LEFT JOIN (select * from bn02 where m_year=?)bn02 on a.bank_no = bn02.bank_no"
					      + " LEFT JOIN cdshareno c on a.subdep_id = c.cmuse_id and c.cmuse_div='010'"
						  + " where not exists"
					      + " (select muser_id "
   						  + " from wtt04"
  						  + " where a.muser_id=wtt04.muser_id "
  						  + ")";
  			paramList.add(yy) ;
  			paramList.add(yy) ;
  			if(muser_type.equals("S")){
  			   sqlCmd += " and a.muser_type='A'";
  			}			  
  			if(muser_type.equals("A")){
  			   sqlCmd += " and a.muser_type=' '"
  			           + " and a.tbank_no = ?";
  			   paramList.add(tbank_no) ;
  			}
            sqlCmd += " and a.delete_mark <> 'Y'"
					+ " order by a.muser_id,a.muser_name";
    		
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
            return dbData;
    } 
    
    //取得WTT01已建立權限的user
    private List getWTT01_hasCreate_New(String muser_type,String tbank_no){
    		//查詢條件        
    		String yy = Integer.parseInt(Utility.getYear())>99 ?"100" :"99" ;
    		List paramList =new ArrayList() ;
  			String sqlCmd = " select a.muser_id,a.muser_name,a.bank_type,c1.cmuse_name as bank_type_name"
    					  + " ,a.tbank_no,bn01.bank_name as tbank_no_name"
    					  + " ,a.bank_no,bn02.bank_name as bank_no_name"
    					  + " ,a.subdep_id,c.cmuse_name as subdep_id_name "
					      + " from wtt01 a"
						  + " LEFT JOIN cdshareno c1 on a.bank_type = c1.cmuse_id and c1.cmuse_div='001' "
						  + " LEFT JOIN (select * from bn01 where m_year=?)bn01 on a.tbank_no = bn01.bank_no"
						  + " LEFT JOIN (select * from bn02 where m_year=?)bn02 on a.bank_no = bn02.bank_no"
					      + " LEFT JOIN cdshareno c on a.subdep_id = c.cmuse_id and c.cmuse_div='010'"
					      + " where exists"
					      + " (select muser_id "
   						  + " from wtt04"
  						  + " where a.muser_id=wtt04.muser_id)";
  			paramList.add(yy) ;
  			paramList.add(yy);
  			if(muser_type.equals("S")){
  			   sqlCmd += " and a.muser_type='A'";
  			}			  
  			if(muser_type.equals("A")){
  			   sqlCmd += " and a.muser_type=' '"
  			           + " and a.tbank_no = ? ";
  			   paramList.add(tbank_no) ;
  			   
  			}
            sqlCmd += " and a.delete_mark <> 'Y'"
					+ " order by a.muser_id,a.muser_name";
    		
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
            return dbData;
    } 
    
    //取得WTT01未建立權限的user
    private List getWTT01_hasNotCreate(String muser_type,String tbank_no){
    		//查詢條件    
    		List paramList =new ArrayList() ;
    		String sqlCmd = " select wtt01.muser_id,wtt01.muser_name "
					      + " from wtt01 "
					      + " where not exists"
					      + " (select muser_id "
   						  + " from wtt04"
  						  + " where wtt01.muser_id=wtt04.muser_id)";
  			if(muser_type.equals("S")){
  			   sqlCmd += " and wtt01.muser_type='A'";
  			}			  
  			if(muser_type.equals("A")){
  			   sqlCmd += " and wtt01.muser_type=' '"
  			           + " and wtt01.tbank_no =? ";
  			   paramList.add(tbank_no) ;
  			   
  			}
            sqlCmd += " and wtt01.delete_mark <> 'Y'"
					+ " order by wtt01.muser_id,wtt01.muser_name";
    		
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
            return dbData;
    } 
    //取得WTT01已建立權限的user
    private List getWTT01_hasCreate(String muser_type,String tbank_no){
    		//查詢條件    
    		List paramList =new ArrayList() ;
    		String sqlCmd = " select wtt01.muser_id,wtt01.muser_name "
					      + " from wtt01 "
					      + " where exists"
					      + " (select muser_id "
   						  + " from wtt04"
  						  + " where wtt01.muser_id=wtt04.muser_id)";
  			if(muser_type.equals("S")){
  			   sqlCmd += " and wtt01.muser_type='A'";
  			}			  
  			if(muser_type.equals("A")){
  			   sqlCmd += " and wtt01.muser_type=' '"
  			           + " and wtt01.tbank_no = ? ";
  			   paramList.add(tbank_no) ;
  			   
  			}
            sqlCmd += " and wtt01.delete_mark <> 'Y'"
					+ " order by wtt01.muser_id,wtt01.muser_name";
    		
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
            return dbData;
    }
    
    //取得WTT01基本資料
    private List getWTT01(String muser_id){
    		//查詢條件 
    		String yy = Integer.parseInt(Utility.getYear())>99 ?"100" :"99" ;
    		List paramList =new ArrayList() ;
    		String sqlCmd = " select a.bank_type,c1.cmuse_name as bank_type_name"
    					  + " ,a.tbank_no,bn01.bank_name as tbank_no_name"
    					  + " ,a.bank_no,bn02.bank_name as bank_no_name"
    					  + " ,a.subdep_id,c.cmuse_name as subdep_id_name "
					      + " from wtt01 a"
						  + " LEFT JOIN cdshareno c1 on a.bank_type = c1.cmuse_id and c1.cmuse_div='001' "
						  + " LEFT JOIN (select * from bn01 where m_year=?)bn01 on a.tbank_no = bn01.bank_no"
						  + " LEFT JOIN (select * from bn02 where m_year=?)bn02 on a.bank_no = bn02.bank_no"
					      + " LEFT JOIN cdshareno c on a.subdep_id = c.cmuse_id and c.cmuse_div='010'"
						  + " where a.muser_id=? ";
    		paramList.add(yy) ;
    		paramList.add(yy) ;
    		paramList.add(muser_id) ;
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
            return dbData;
    } 
    
    
    //取得WTT03_1程式基本功能    
    private List getWTT03_1(String muser_type,String muser_id,String cnd,String menu){
    		//查詢條件    
    		System.out.println("muser_type="+muser_type);   		
    		System.out.println("muser_id="+muser_id);   		
    		System.out.println("cnd="+cnd);   		
    		String sqlCmd = "";
    		List paramList =new ArrayList() ;
    		if(muser_type.equals("S")){
    		  sqlCmd += " select * from WTT03_1 ";    		  
    		  if(menu.equals("MIS")){
    		     sqlCmd += " where program_id like 'FR%' or program_id like 'BR%'"
    		     	     + " or program_id like 'TC%' or program_id like 'ZZ%'"
    		     	     + " or program_id like 'AS%'"
    		     	     + " or program_id like 'DS%'" //95.10.16 add MIS.DS/AN/CG
    		     	     + " or program_id like 'AN%'"
    		     	     + " or program_id like 'CG%'"
    		     	     + " or program_id like 'MC%' or program_id = 'RptFileUpload'"//98.08.12 add MIS.MC
    		     	     + " or program_id like 'WR%'"//102.10.16 add 警示報表
    		     	     + " or program_id like 'TM%'"//105.09.06 add 協助措施
    		     	     + " or program_id like 'FL%'";//105.10.19 add 專案農貸檢查追蹤
    		  }else{
    		      sqlCmd += " where program_id like 'WM%' or program_id like 'ZZ%'"
    		             + " or program_id like 'FX%'"
    		             + " or program_id like 'DS%'"//95.11.10 add BOAF.DS    	
    		             + " or program_id like 'BR%'"//103.01.20 add BOAF.BR	     	     
    		             + " or program_id = 'RptFileUpload'"//98.08.12 add RptFileUpload    	
    		             + " or program_id like 'TM%'";//105.09.06 add 協助措施    		     	    
    		  }
    		  sqlCmd += " order by program_id,program_name";
    		}
    		if(muser_type.equals("A")){
    		  sqlCmd = " select a.program_id,b.program_name,a.p_add,a.p_delete,a.p_update"
    		  		 + " ,a.p_query,a.p_print,a.p_upload,a.p_download,a.p_lock,a.p_other"
					 + " from wtt04 a "
					 + " LEFT JOIN wtt03_1 b on b.program_id = a.program_id "
				     + " where a.muser_id=? ";
    		  //paramList.add(muser_id) ;
			  if(menu.equals("MIS")){
    		     sqlCmd += " and (a.program_id like 'FR%' or a.program_id like 'BR%'"
    		     	     + " or a.program_id like 'TC%' or a.program_id like 'ZZ%'"
    		     	     + " or a.program_id like 'AS%'"
    		     	     + " or a.program_id like 'DS%'" //95.10.16 add MIS.DS/AN/CG
    		     	     + " or a.program_id like 'AN%'"
    		     	     + " or a.program_id like 'CG%'"
    		     	     + " or a.program_id like 'MC%' or a.program_id = 'RptFileUpload'"//98.08.12 add MIS.MC
    		     	     + " or a.program_id like 'WR%'"//102.10.16 add 警示報表
    		     	     + " or a.program_id like 'TM%'"//105.09.06 add 協助措施
    		     	     + " or a.program_id like 'FL%'"//105.10.19 add 專案農貸檢查追蹤
    		     	     + ")";
    		  }else{
    		      sqlCmd += " and (a.program_id like 'WM%' or a.program_id like 'ZZ%'"
    		             + " or a.program_id like 'FX%'"
    		             + " or a.program_id like 'DS%'" //95.11.10 add BOAF.DS
    		             + " or a.program_id like 'BR%'"//103.01.20 add BOAF.BR
    		             + " or a.program_id = 'RptFileUpload'"//98.08.12 add RptFileUpload    		            
    		             + " or a.program_id like 'TM%'"//105.09.06 add 協助措施    		     	     
    		             + ")";    		     	     
    		  }	     
			  sqlCmd += " order by a.program_id";
    		}
    		
    		if(cnd.equals("getUser")){ 
    		   
    		  if(muser_type.equals("S") || muser_type.equals("A")){
    		    sqlCmd = " select wtt04_1.program_id,wtt03_1.program_name,"
       				   + " wtt04_1.p_add as p_add_o,wtt04_1.p_delete as p_delete_o,"
       				   + " wtt04_1.p_update as p_update_o,wtt04_1.p_query as p_query_o,"
       				   + " wtt04_1.p_print as p_print_o,wtt04_1.p_upload as p_upload_o,"
       				   + " wtt04_1.p_download as p_download_o,wtt04_1.p_lock as p_lock_o,"
       				   + " wtt04_1.p_other as p_other_o,"
       				   + " nvl(wtt04_2.p_add ,'N') as p_add,nvl(wtt04_2.p_delete ,'N') as p_delete,"
       				   + " nvl(wtt04_2.p_update ,'N') as p_update,nvl(wtt04_2.p_query ,'N') as p_query,"
       				   + " nvl(wtt04_2.p_print ,'N') as p_print,nvl(wtt04_2.p_upload ,'N') as p_upload,"
       				   + " nvl(wtt04_2.p_download ,'N') as p_download,nvl(wtt04_2.p_lock ,'N') as p_lock,"
       				   + " nvl(wtt04_2.p_other ,'N') as p_other "
				       + " from wtt01 join wtt04 wtt04_1 on wtt01.muser_id=? and wtt01.add_user=wtt04_1.muser_id";
    		     //paramList.add(muser_id) ;
			     if(menu.equals("MIS")){
    		         sqlCmd += " and ( wtt04_1.program_id like 'FR%' or wtt04_1.program_id like 'BR%'"
    		     	        + " or wtt04_1.program_id like 'TC%' or wtt04_1.program_id like 'ZZ%'"
    		     	        + " or wtt04_1.program_id like 'AS%'"
    		     	        + " or wtt04_1.program_id like 'DS%'"//95.10.16 add MIS.DS/AN/CG
    		     	        + " or wtt04_1.program_id like 'AN%'"
    		     	        + " or wtt04_1.program_id like 'CG%'"
    		     	        + " or wtt04_1.program_id like 'MC%' or wtt04_1.program_id = 'RptFileUpload'"//98.08.12 add MIS.MC
    		     	        + " or wtt04_1.program_id like 'WR%'"//102.10.16 add 警示報表
    		     	        + " or wtt04_1.program_id like 'TM%'"//105.09.06 add 協助措施
    		     	        + " or wtt04_1.program_id like 'FL%'"//105.10.19 add 專案農貸檢查追蹤
    		     	        + " )";
    		     	        
    		      }else{
    		         sqlCmd += " and ( wtt04_1.program_id like 'WM%' or wtt04_1.program_id like 'ZZ%'"
    		                + " or wtt04_1.program_id like 'FX%'"
    		                + " or wtt04_1.program_id like 'DS%'"//95.11.10 add BOAF.DS
    		                + " or wtt04_1.program_id like 'BR%'"//103.01.20 add BOAF.BR
    		                + " or wtt04_1.program_id = 'RptFileUpload'"//98.08.12 add RptFileUpload
    		                + " or wtt04_1.program_id like 'TM%'"//105.09.06 add 協助措施    		     	        
    		     	        + " )";
    		      }				       
     			  sqlCmd += " left join wtt04 wtt04_2 on wtt04_1.program_id=wtt04_2.program_id and wtt01.muser_id=wtt04_2.muser_id"
	 				     + " join wtt03_1 on wtt03_1.program_id=wtt04_1.program_id "	 			  	   
				         + " order by wtt03_1.program_id";
    		  }else{
    		     sqlCmd = " select a.program_id,b.program_name,a.p_add,a.p_delete,a.p_update,a.p_query,"
    		  		 + " a.p_print,a.p_upload,a.p_download,a.p_lock,a.p_other"
 				     + " from wtt04 a,wtt03_1 b"
 					 + " where a.muser_id=? "
 				     + " and a.program_id = b.program_id";
    		     //paramList.add(muser_id) ;
 				 if(menu.equals("MIS")){
    		        sqlCmd += " and (a.program_id like 'FR%' or a.program_id like 'BR%'"
    		     	       + " or a.program_id like 'TC%' or a.program_id like 'ZZ%'"
    		     	       + " or a.program_id like 'AS%'"
    		     	       + " or a.program_id like 'DS%'"//95.10.16 add MIS.DS/AN/CG
    		     	       + " or a.program_id like 'AN%'"
    		     	       + " or a.program_id like 'CG%'"
    		     	       + " or a.program_id like 'MC%' or a.program_id = 'RptFileUpload'"//98.08.12 add MIS.MC
    		     	       + " or a.program_id like 'WR%'"//102.10.16 add 警示報表
    		     	       + " or a.program_id like 'TM%'"//105.09.06 add 協助措施
    		     	       + " or a.program_id like 'FL%'"//105.10.19 add 專案農貸檢查追蹤
    		     	       + " )";
    		     }else{
    		      sqlCmd += " and (a.program_id like 'WM%' or a.program_id like 'ZZ%'"
    		      	     + " or a.program_id like 'FX%'"
    		     	     + " or a.program_id like 'DS%'"//95.11.10 add BOAF.DS
    		     	     + " or a.program_id like 'BR%'"//103.01.20 add BOAF.BR
    		     	     + " or a.program_id = 'RptFileUpload'"//98.08.12 add RptFileUpload
    		     	     + " or a.program_id like 'TM%'"//105.09.06 add 協助措施
    		     	     + " )";
    		     }	       
    		  }     
 			}//end of getUser	     
 			//100.02.16 add
 			List dbData = null;
 			if(muser_type.equals("S")){
 			   if(cnd.equals("getUser")){ 
               	  paramList.add(muser_id) ;
               	  dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");   
               }else{
                  dbData = DBManager.QueryDB_SQLParam(sqlCmd,null,"");            
               }
            }else{
               paramList.add(muser_id) ;
               dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");   
            }
            
            return dbData;
    }
    
    //取得WTT03_1 Edit權限
    private List getWTT03_1_edit(String muser_type,String muser_id){
    		//查詢條件    
    		String sqlCmd = "";
    		List paramList =new ArrayList() ;
    		if(muser_type.equals("S")){
    		  sqlCmd = " select count(*) as temp_edit "
			         + " from wtt03_1 b"
				     + " where b.p_other = 'Y'"
				     + " and  (b.p_add = 'Y' or p_delete = 'Y' or b.p_update='Y')";
    		}
    		if(muser_type.equals("A")){
    		  sqlCmd = " select count(*) as temp_edit"
					 + " from wtt04 a"
				     + " where  a.muser_id = ? "
				     + " and a.p_other = 'Y'"
				     + " and (a.p_add='Y' or a.p_delete='Y' or a.p_update='Y') ";
    		  paramList.add(muser_id) ;
    		}
    		
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
            return dbData;
    }
    
    
    //取得WTT03_1 upload權限
    private List getWTT03_1_upload(String muser_type,String muser_id){
    		//查詢條件    
    		String sqlCmd = "";
    		List paramList =new ArrayList() ;
    		if(muser_type.equals("S")){
    		  sqlCmd = " select count(*) as temp_upload "
			         + " from wtt03_1 b"
				     + " where b.p_other = 'Y'"
				     + " and  b.p_upload='Y'";
    		}
    		if(muser_type.equals("A")){
    		  sqlCmd = " select count(*) as temp_upload"
					 + " from wtt04 a"
				     + " where  a.muser_id = ? "
				     + " and a.p_other = 'Y'"
				     + " and a.p_upload='Y' ";
    		  paramList.add(muser_id) ;
    		}
    		
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
            
            return dbData;
    }
    
    //取得WTT03_1 download權限
    private List getWTT03_1_download(String muser_type,String muser_id){
    		//查詢條件    
    		String sqlCmd = "";
    		List paramList =new ArrayList() ;
    		if(muser_type.equals("S")){
    		  sqlCmd = " select count(*) as temp_download "
			         + " from wtt03_1 b"
				     + " where b.p_other = 'Y'"
				     + " and  b.p_download='Y'";
    		}
    		if(muser_type.equals("A")){
    		  sqlCmd = " select count(*) as temp_download"
					 + " from wtt04 a"
				     + " where  a.muser_id = ? "
				     + " and a.p_other = 'Y'"
				     + " and a.p_download='Y' ";
    		  paramList.add(muser_id) ;
    		}
    		
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
            return dbData;
    }
    
    //取得WTT03_1 Query權限
    private List getWTT03_1_query(String muser_type,String muser_id){
    		//查詢條件    
    		String sqlCmd = "";
    		List paramList = new ArrayList() ;
    		if(muser_type.equals("S")){
    		  sqlCmd = " select count(*) as temp_query "
			         + " from wtt03_1 b"
				     + " where b.p_other = 'Y'"
				     + " and  b.p_query = 'Y'";
    		}
    		if(muser_type.equals("A")){
    		  sqlCmd = " select count(*) as temp_query"
					 + " from wtt04 a"
				     + " where  a.muser_id = ?"
				     + " and a.p_other = 'Y'"
				     + " and a.p_query = 'Y' ";
    		  paramList.add(muser_id) ;
    		}
    		
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
            return dbData;
    }
    
    
    //取得WTT04該使用者程式權限
    private List getWTT04(String muser_id){
    		//查詢條件    
    		String sqlCmd = "";    		
    		List paramList =new ArrayList() ;
    		sqlCmd = " select * from WTT04 where muser_id =?  order by program_id";
    		paramList.add(muser_id) ;    		
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
            return dbData;
    }
    
    //取得WTT04該使用者報表權限
    private List getWTT04_Report(String muser_id,String detail_type){
    		//查詢條件    
    		String sqlCmd = "";    	
    		List paramList =new ArrayList() ;
    		sqlCmd = " select distinct wtt04_1d.report_no,"
           		   +			       "a.cmuse_id as trans_type,"
           		   +                   "a.cmuse_name as trans_type_name ,"
           		   +                   "b.cmuse_id as report_no,"
           		   +                   "b.cmuse_name as report_name ,"
           		   +                   "a.input_order,b.input_order "
  		   		   + " from wtt04_1d,cdshareno a,cdshareno b "
 		  		   + " where wtt04_1d.muser_id=?"  
 		  		   + " and wtt04_1d.detail_type = ? "
   		   		   + " and a.cmuse_div='011' and a.cmuse_id = wtt04_1d.transfer_type"
   		   		   + " and b.cmuse_div in ('012','013','014','030')  " //94.11.15 add 030:F01_在台無住所之外國人新台幣存款表
   		   		   + " and a.identify_no = b.identify_no"
   		   		   + " and b.cmuse_name like wtt04_1d.report_no || '%' "
 		           + " order by a.input_order,b.input_order"; 		
   		    paramList.add(muser_id) ;
   		    paramList.add(detail_type) ;
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
            return dbData;
    }
    
    //取得WTT04該使用者可修改的報表權限
    private List getWTT04_Report_update(String muser_id,String detail_type){
    		//detail --> 1:upload 2:download 3:edit 4:query
    		//查詢條件    
    		String sqlCmd = "";    	
    		List paramList =new ArrayList() ;
    		sqlCmd = " select distinct wtt04_1.report_no as report_no_o,wtt04_2.report_no,"
       			   + " c1.cmuse_id as trans_type,c1.cmuse_name as trans_type_name,"
	   			   + " c2.cmuse_id as report_no,c2.cmuse_name as report_name,"
       			   + " c1.input_order,c2.input_order "
				   + " from wtt01 join wtt04_1d wtt04_1 on wtt01.muser_id=?  and "
                   + "                 wtt01.add_user=wtt04_1.muser_id and "
				   + "				   wtt04_1.detail_type =? "
     		       + " left join wtt04_1d wtt04_2 on wtt04_1.program_id=wtt04_2.program_id and "
	        	   + "                    wtt01.muser_id=wtt04_2.muser_id and"
                   + "		              wtt04_1.detail_type=wtt04_1.detail_type and "
                   + "		              wtt04_1.transfer_type=wtt04_2.transfer_type and "
                   + "                    wtt04_1.report_no=wtt04_2.report_no"
     			   + " join cdshareno c1 on c1.cmuse_div='011' and c1.cmuse_id = wtt04_1.transfer_type"
	 			   + " join cdshareno c2 on c2.cmuse_div in ('012','013','014','030') and " //94.11.15 add 030:F01_在台無住所之外國人新台幣存款表
	               + " c1.identify_no = c2.identify_no and "
				   + " c2.cmuse_name like wtt04_1.report_no || '%'"
				   + " order by c1.input_order,c2.input_order ";
	 	    paramList.add(muser_id) ;
	 	    paramList.add(detail_type) ;
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
            return dbData;
    }
    //取得查詢結果
    private List getQryResult(String muser_type,String tbank_no,String id_Create){
    		//查詢條件        		
    		String sqlCmd = "";
    		String yy = Integer.parseInt(Utility.getYear())>99 ?"100" : "99" ;
    		List paramList = new ArrayList() ;
    		sqlCmd = " select a.tbank_no,bn01.bank_name as tbank_no_name,a.bank_no,bn02.bank_name as bank_no_name"
    			   + " ,a.subdep_id,cdshareno.cmuse_name as subdep_id_name"
    			   + " ,a.muser_id,a.muser_name,c.program_id,c.program_name"
				   + " ,b.p_add,b.p_delete,b.p_update,b.p_query,b.p_print"
				   + " ,b.p_upload,b.p_download,b.p_lock,b.p_other"
			       + " from WTT01 a "
			       + " LEFT JOIN (select * from bn01 where m_year=?)bn01 on a.tbank_no = bn01.bank_no"
				   + " LEFT JOIN (select * from bn02 where m_year=?)bn02 on a.bank_no = bn02.bank_no"
				   + " LEFT JOIN cdshareno on a.subdep_id = cdshareno.cmuse_id and cdshareno.cmuse_div='010'"
				   + " ,WTT04 b "
				   + " ,WTT03_1 c";
    		paramList.add(yy) ;
    		paramList.add(yy) ;
            if(id_Create.equals("")){				   
			   sqlCmd += " where a.muser_id = b.muser_id";
		    }else{
		       sqlCmd += " where a.muser_id = b.muser_id and b.muser_id=? ";
		       paramList.add(id_Create) ;
		    }		   
			if(muser_type.equals("S")){	   
			   sqlCmd += " and muser_type='A'";
            }				   
            if(muser_type.equals("A")){	   
			   sqlCmd += " and muser_type=' '";
			   sqlCmd += " and a.tbank_no=? ";
			   paramList.add(tbank_no) ;
            }            
            
			sqlCmd += " and a.delete_mark <> 'Y'"
				   + " and b.program_id = c.program_id"
				   + " order by a.bank_type,a.tbank_no,a.bank_no,a.subdep_id,a.muser_id,b.program_id";
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
            return dbData;
    }
    
    
   
    
	public String InsertDB(HttpServletRequest request,String muser_id,String lguser_id,String lguser_name) throws Exception{    	
		String sqlCmd = "";		
		String errMsg="";
	    String user_id=lguser_id;
	    String user_name=lguser_name;			 
	    String program_id="";
	    String p_add="";
	    String p_delete="";
	    String p_update="";
	    String p_query="";
	    String p_print="";
	    String p_upload="";
	    String p_download="";
	    String p_lock="";
	    String p_other="";   	    
	    String transfer_type="";
	    String report_no="";
	    String detail_type="";
		//List updateDBSqlList = new LinkedList();
		List paramList = new ArrayList() ;
		try {			    			    			    
			    System.out.println("insert begin");
			    //取出form裡的所有變數=================================== 
		  		Enumeration ep = request.getParameterNames();
		  		Enumeration ea = request.getAttributeNames();
		  		Hashtable t = new Hashtable();
		  		String name = "";
		  
		  		for ( ; ep.hasMoreElements() ; ) {
			   		name = (String)ep.nextElement();
			   		t.put( name, request.getParameter(name) );		
			   		System.out.println(name+"="+request.getParameter(name));	   
		  		}		
		  		System.out.println("insertuser="+muser_id);
		  		//取得程式基本功能權限=======================================================  		  		
		  		int row_program =Integer.parseInt((String)t.get("row_program"));		  				  		
		  		System.out.println("row_program="+row_program);
		  	    List pgData = new LinkedList();
		  	    List programidData = null;
		  		for ( int i = 0; i < row_program; i++) {		  	    		  	  			  
		  		    programidData = new LinkedList();
		  		    if(t.get("program_id_" + (i+1)) != null ) {					  
					   programidData.add((String)t.get("program_id_"+(i+1)));
					}		  		   
					if(t.get("P_ADD_" + (i+1)) != null ) {					  
					   programidData.add((String)t.get("P_ADD_"+(i+1)));
					}else{
					   programidData.add("N");
					}										
					if(t.get("P_DELETE_" + (i+1)) != null ) {					  
					   programidData.add((String)t.get("P_DELETE_"+(i+1)));
					}else{
					   programidData.add("N");
					}
					if(t.get("P_UPDATE_" + (i+1)) != null ) {					  
					   programidData.add((String)t.get("P_UPDATE_"+(i+1)));
					}else{
					   programidData.add("N");
					}
					if(t.get("P_QUERY_" + (i+1)) != null ) {					  
					   programidData.add((String)t.get("P_QUERY_"+(i+1)));
					}else{
					   programidData.add("N");
					}
					if(t.get("P_PRINT_" + (i+1)) != null ) {					  
					   programidData.add((String)t.get("P_PRINT_"+(i+1)));
					}else{
					   programidData.add("N");
					}
					if(t.get("P_UPLOAD_" + (i+1)) != null ) {					  
					   programidData.add((String)t.get("P_UPLOAD_"+(i+1)));
					}else{
					   programidData.add("N");
					}
					if(t.get("P_DOWNLOAD_" + (i+1)) != null ) {					  
					   programidData.add((String)t.get("P_DOWNLOAD_"+(i+1)));
					}else{
					   programidData.add("N");
					}
					if(t.get("P_LOCK_" + (i+1)) != null ) {					  
					   programidData.add((String)t.get("P_LOCK_"+(i+1)));
					}else{
					   programidData.add("N");
					}
					if(t.get("P_OTHER_" + (i+1)) != null ) {					  
					   programidData.add((String)t.get("P_OTHER_"+(i+1)));
					}else{
					   programidData.add("N");
					}
					pgData.add(programidData);
		  		}	
		  		System.out.println("pgData.size="+pgData.size());
		  		for(int i=0;i<pgData.size();i++){
		  			paramList.clear() ;
		  		    System.out.println((List)pgData.get(i));		  		   	    
		  		    program_id = (String)((List)pgData.get(i)).get(0);		  		   
		  		    p_add = (String)((List)pgData.get(i)).get(1);
		  		    p_delete = (String)((List)pgData.get(i)).get(2);
		  		    p_update = (String)((List)pgData.get(i)).get(3);
		  		    p_query = (String)((List)pgData.get(i)).get(4);
		  		    p_print = (String)((List)pgData.get(i)).get(5);
		  		    p_upload = (String)((List)pgData.get(i)).get(6);
		  		    p_download = (String)((List)pgData.get(i)).get(7);
		  		    p_lock = (String)((List)pgData.get(i)).get(8);
		  		    p_other = (String)((List)pgData.get(i)).get(9);
		  		    if( p_add.equals("Y") || p_delete.equals("Y") || p_update.equals("Y")
		  		    ||  p_query.equals("Y") || p_print.equals("Y") || p_upload.equals("Y")
		  		    ||  p_download.equals("Y") || p_lock.equals("Y") || p_other.equals("Y")
		  		    ){
		  		       sqlCmd = "INSERT INTO WTT04 VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,sysdate)";
					   
		  		       paramList.add(muser_id) ;
		  		       paramList.add(program_id) ;
		  		       paramList.add(p_add) ;
		  		       paramList.add(p_delete) ;
		  		       paramList.add(p_update);
		  		       paramList.add(p_query) ;
		  		       paramList.add(p_print) ;
		  		       paramList.add(p_upload) ;
		  		       paramList.add(p_download) ;
		  		       paramList.add(p_lock) ;
		  		       paramList.add(p_other) ;
		  		       paramList.add(user_id) ;
		  		       paramList.add(user_name) ;
		  		       if(!this.updDbUsesPreparedStatement(sqlCmd,paramList)) {
		  		    	   throw new Exception();
		  		       }
					   //updateDBSqlList.add(sqlCmd); 		            	   				                  		 						
    	   		    }		  		    
		  		}//enf of pgData
		  		//=========================================================================
		  		//取得上傳權限=============================================================
		  		int row_upload = 0;
		  		if(t.get("row_upload") != null){		  		   
		  		   row_upload = Integer.parseInt((String)t.get("row_upload"));
		  		}  
		  		System.out.println("row_upload="+row_upload);
		  	    List uploadData = new LinkedList();
		  	    List upload_detailData = null;
		  		for ( int i = 0; i < row_upload; i++) {		  	    		  	  			  
		  		    upload_detailData = new LinkedList();
		  		    if(t.get("upload_isModify_" + (i+1)) != null ) {					  					   
					   if(t.get("uploadData_trans_type_" + (i+1)) != null ) {					  
					      upload_detailData.add((String)t.get("uploadData_trans_type_"+(i+1)));
					   }										
					   if(t.get("uploadData_report_no_" + (i+1)) != null ) {					  
					      upload_detailData.add((String)t.get("uploadData_report_no_"+(i+1)));
					  }
					  uploadData.add(upload_detailData);
					}
					
		  		}	
		  		System.out.println("uploadData.size()="+uploadData.size());
		  		for(int i=0;i<uploadData.size();i++){
		  			paramList.clear() ;
		  		    System.out.println("uploadData="+uploadData.get(i));
		  		    program_id = "WMFileUpload";
		  		    detail_type="1";
		  		    
		  		    transfer_type=(String)((List)uploadData.get(i)).get(0);
		  		    report_no = (String)((List)uploadData.get(i)).get(1);
		  		    
		  		    sqlCmd = "INSERT INTO WTT04_1D VALUES (?,?,?,?,?,?,?,sysdate)";
					 
		  		    paramList.add(muser_id) ;
		  		    paramList.add(program_id) ;
		  		    paramList.add(detail_type) ;
		  		    paramList.add(transfer_type) ;
		  		    paramList.add(report_no) ;
		  		    paramList.add(user_id) ;
		  		    paramList.add(user_name);
		  		    if(!this.updDbUsesPreparedStatement(sqlCmd,paramList)) {
		  		    	throw new Exception();
		  		    }
					 //  updateDBSqlList.add(sqlCmd); 		
		  		}
		  		//=========================================================================		  		
	            
	            //取得下載權限=============================================================
		  		int row_download = 0;
		  		if( t.get("row_download") != null){
		  		    row_download = Integer.parseInt((String)t.get("row_download"));
		  		}   
		  		System.out.println("row_download="+row_download);
		  	    List downloadData = new LinkedList();
		  	    List download_detailData = null;
		  		for ( int i = 0; i < row_download; i++) {		  	    		  	  			  
		  		    download_detailData = new LinkedList();
		  		    if(t.get("download_isModify_" + (i+1)) != null ) {		
		  		       System.out.println("add"+i);			  					   
					   if(t.get("downloadData_trans_type_" + (i+1)) != null ) {					  
					      download_detailData.add((String)t.get("downloadData_trans_type_"+(i+1)));
					   }										
					   if(t.get("downloadData_report_no_" + (i+1)) != null ) {					  
					      download_detailData.add((String)t.get("downloadData_report_no_"+(i+1)));
					  }
					  downloadData.add(download_detailData);
					}
					
		  		}	
		  		for(int i=0;i<downloadData.size();i++){
		  		    System.out.println(downloadData.get(i));
		  		}
		  		System.out.println("downloadData.size()="+downloadData.size());
		  		for(int i=0;i<downloadData.size();i++){
		  			paramList.clear() ;
		  		    System.out.println("downloadData="+downloadData.get(i));
		  		    program_id = "WMFileDownload";
		  		    detail_type="2";
		  		    
		  		    transfer_type=(String)((List)downloadData.get(i)).get(0);
		  		    report_no = (String)((List)downloadData.get(i)).get(1);
		  		    
		  		    sqlCmd = "INSERT INTO WTT04_1D VALUES (?,?,?,?,?,?,?,sysdate)";
				
		  		    paramList.add(muser_id) ;
		  		    paramList.add(program_id);
		  		    paramList.add(detail_type) ;
		  		    paramList.add(transfer_type) ;
		  		    paramList.add(report_no) ;
		  		    paramList.add(user_id) ;
		  		    paramList.add(user_name) ;
		  		    if(!this.updDbUsesPreparedStatement(sqlCmd,paramList)){
		  		    	throw new Exception() ;
		  		    }
					   //updateDBSqlList.add(sqlCmd); 		
		  		}
		  		//=========================================================================		  		
	            
	            //取得編輯權限=============================================================
		  		int row_edit = 0 ;
		  		if(t.get("row_edit") != null){
		  		   row_edit = Integer.parseInt((String)t.get("row_edit"));
		  		}
		  		System.out.println("row_edit="+row_edit);
		  	    List editData = new LinkedList();
		  	    List edit_detailData = null;
		  		for ( int i = 0; i < row_edit; i++) {		  	    		  	  			  
		  		    edit_detailData = new LinkedList();
		  		    if(t.get("edit_isModify_" + (i+1)) != null ) {					  					   
					   if(t.get("editData_trans_type_" + (i+1)) != null ) {					  
					      edit_detailData.add((String)t.get("editData_trans_type_"+(i+1)));
					   }										
					   if(t.get("editData_report_no_" + (i+1)) != null ) {					  
					      edit_detailData.add((String)t.get("editData_report_no_"+(i+1)));
					  }
					  editData.add(edit_detailData);
					}
					
		  		}	
		  		System.out.println("editData.size()="+editData.size());
		  		for(int i=0;i<editData.size();i++){
		  			paramList.clear() ;
		  		    System.out.println("editData="+editData.get(i));
		  		    program_id = "WMFileEdit";
		  		    detail_type="3";
		  		    
		  		    transfer_type=(String)((List)editData.get(i)).get(0);
		  		    report_no = (String)((List)editData.get(i)).get(1);
		  		    
		  		    sqlCmd = "INSERT INTO WTT04_1D VALUES (?,?,?,?,?,?,?,sysdate)";
					  
		  		    paramList.add(muser_id) ;
		  		    paramList.add(program_id) ;
		  		    paramList.add(detail_type) ;
		  		    paramList.add(transfer_type) ;
		  		    paramList.add(report_no);
		  		    paramList.add(user_id) ;
		  		    paramList.add(user_name) ;
		  		    if(!this.updDbUsesPreparedStatement(sqlCmd,paramList)){
		  		    	throw new Exception();
		  		    }
					//   updateDBSqlList.add(sqlCmd); 		
		  		}
		  		//=========================================================================		  		
	            
	            //取得查詢權限=============================================================
		  		int row_query = 0;
		  		if(t.get("row_query") != null){
		  		   row_query = Integer.parseInt((String)t.get("row_query"));
		  		}  
		  		System.out.println("row_query="+row_query);
		  	    List queryData = new LinkedList();
		  	    List query_detailData = null;
		  		for ( int i = 0; i < row_query; i++) {		  	    		  	  			  
		  		    query_detailData = new LinkedList();
		  		    if(t.get("query_isModify_" + (i+1)) != null ) {					  					   
					   if(t.get("queryData_trans_type_" + (i+1)) != null ) {					  
					      query_detailData.add((String)t.get("queryData_trans_type_"+(i+1)));
					   }										
					   if(t.get("queryData_report_no_" + (i+1)) != null ) {					  
					      query_detailData.add((String)t.get("queryData_report_no_"+(i+1)));
					  }
					  queryData.add(query_detailData);
					}
					
		  		}	
		  		System.out.println("queryData.size()="+queryData.size());
		  		for(int i=0;i<queryData.size();i++){
		  			paramList.clear() ;
		  		    System.out.println("queryData="+queryData.get(i));
		  		    program_id = "WMFileQuery";
		  		    detail_type="4";
		  		    
		  		    transfer_type=(String)((List)queryData.get(i)).get(0);
		  		    report_no = (String)((List)queryData.get(i)).get(1);
		  		    
		  		    sqlCmd = "INSERT INTO WTT04_1D VALUES (?,?,?,?,?,?,?,sysdate)";
					
		  		    paramList.add(muser_id) ;
		  		    paramList.add(program_id) ;
		  		    paramList.add(detail_type) ;
		  		    paramList.add(transfer_type) ;
		  		    paramList.add(report_no);
		  		    paramList.add(user_id) ;
		  		    paramList.add(user_name);
		  		    if(!this.updDbUsesPreparedStatement(sqlCmd,paramList)) {
		  		    	throw new Exception() ;
		  		    }
					   //updateDBSqlList.add(sqlCmd); 		
		  		}
		  		//=========================================================================		  		
	            errMsg = errMsg + "相關資料寫入資料庫成功";		
			    /*if(DBManager.updateDB(updateDBSqlList)){					 
				   errMsg = errMsg + "相關資料寫入資料庫成功";					
			    }else{
			 	   errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
			    }*/				    
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";							
		}	

		return errMsg;
	} 
	
	
	public String UpdateDB(HttpServletRequest request,String muser_id,String lguser_id,String lguser_name,String menu) throws Exception{    	
		String sqlCmd = "";		
		String errMsg="";
	    String user_id=lguser_id;
	    String user_name=lguser_name;			 
	    String program_id="";
	    String p_add="";
	    String p_delete="";
	    String p_update="";
	    String p_query="";
	    String p_print="";
	    String p_upload="";
	    String p_download="";
	    String p_lock="";
	    String p_other="";   	    
	    String transfer_type="";
	    String report_no="";
	    String detail_type="";
		//List updateDBSqlList = new LinkedList();
		List paramList =new ArrayList() ;
		List dbData = null;
		try {			    			    			    
			    System.out.println("insert begin");
			    //取出form裡的所有變數=================================== 
		  		Enumeration ep = request.getParameterNames();
		  		Enumeration ea = request.getAttributeNames();
		  		Hashtable t = new Hashtable();
		  		String name = "";
		  
		  		for ( ; ep.hasMoreElements() ; ) {
			   		name = (String)ep.nextElement();
			   		t.put( name, request.getParameter(name) );		
			   		System.out.println(name+"="+request.getParameter(name));	   
		  		}		
		  		System.out.println("insertuser="+muser_id);
		  		//取得程式基本功能權限=======================================================  		  		
		  		int row_program =Integer.parseInt((String)t.get("row_program"));		  				  		
		  		System.out.println("row_program="+row_program);
		  	    List pgData = new LinkedList();
		  	    List programidData = null;
		  		for ( int i = 0; i < row_program; i++) {		  	    		  	  			  
		  		    programidData = new LinkedList();
		  		    if(t.get("program_id_" + (i+1)) != null ) {					  
					   programidData.add((String)t.get("program_id_"+(i+1)));
					}		  		   
					if(t.get("P_ADD_" + (i+1)) != null ) {					  
					   programidData.add((String)t.get("P_ADD_"+(i+1)));
					}else{
					   programidData.add("N");
					}										
					if(t.get("P_DELETE_" + (i+1)) != null ) {					  
					   programidData.add((String)t.get("P_DELETE_"+(i+1)));
					}else{
					   programidData.add("N");
					}
					if(t.get("P_UPDATE_" + (i+1)) != null ) {					  
					   programidData.add((String)t.get("P_UPDATE_"+(i+1)));
					}else{
					   programidData.add("N");
					}
					if(t.get("P_QUERY_" + (i+1)) != null ) {					  
					   programidData.add((String)t.get("P_QUERY_"+(i+1)));
					}else{
					   programidData.add("N");
					}
					if(t.get("P_PRINT_" + (i+1)) != null ) {					  
					   programidData.add((String)t.get("P_PRINT_"+(i+1)));
					}else{
					   programidData.add("N");
					}
					if(t.get("P_UPLOAD_" + (i+1)) != null ) {					  
					   programidData.add((String)t.get("P_UPLOAD_"+(i+1)));
					}else{
					   programidData.add("N");
					}
					if(t.get("P_DOWNLOAD_" + (i+1)) != null ) {					  
					   programidData.add((String)t.get("P_DOWNLOAD_"+(i+1)));
					}else{
					   programidData.add("N");
					}
					if(t.get("P_LOCK_" + (i+1)) != null ) {					  
					   programidData.add((String)t.get("P_LOCK_"+(i+1)));
					}else{
					   programidData.add("N");
					}
					if(t.get("P_OTHER_" + (i+1)) != null ) {					  
					   programidData.add((String)t.get("P_OTHER_"+(i+1)));
					}else{
					   programidData.add("N");
					}
					pgData.add(programidData);
		  		}	
		  		System.out.println("pgData.size="+pgData.size());
		  		//將原本的程式功能清掉===============================================
		  		sqlCmd = "select * from WTT04 where muser_id='"+muser_id+"'";		  		       
		  		//94.07.28 MIS只能清FR/BR/TC/AS/ZZ;BOAF只能清FX/WM/ZZ========================
		  		//95.10.16 MIS add DS/AN/CG by 2295
		  		//95.11.10 BOAF add DS by 2295
		  		if(menu.equals("MIS")){
		  		   sqlCmd += " and (program_id like 'BR%' or program_id like 'FR%' or program_id like 'TC%' or program_id like 'AS%' or program_id like 'ZZ%' or program_id like 'DS%' or program_id like 'AN%' or program_id like 'CG%' or program_id like 'MC%' or program_id like 'WR%' or program_id like 'TM%' or program_id like 'FL%' or program_id = 'RptFileUpload')";//98.08.12 add MIS.MC //102.10.16 add MIS.WR //105.09.06 add MIS.TM/FL
		  		}else{
		  		   sqlCmd += " and (program_id like 'FX%' or program_id like 'WM%' or program_id like 'ZZ%' or program_id like 'DS%' or program_id = 'RptFileUpload' or program_id like 'BR%' or program_id like 'TM%')";//103.01.20 add BOAF.BR //105.09.06 add BOAF.TM
		  		}
		  		//===========================================================================
		  		dbData = DBManager.QueryDB_SQLParam(sqlCmd,null,"");
		  		if(dbData.size() != 0){    
		  		   paramList.clear() ;
		  		   System.out.println("dbData.size="+dbData.size());
		  		   sqlCmd = "DELETE WTT04 where muser_id=? ";	
		  		   paramList.add(muser_id) ;
		  		   //95.10.16 MIS add DS/AN/CG by 2295
		  		   //95.11.10 BOAF add DS by 2295
		  		   if(menu.equals("MIS")){
		  		      sqlCmd += " and (program_id like 'BR%' or program_id like 'FR%' or program_id like 'TC%' or program_id like 'AS%' or program_id like 'ZZ%' or program_id like 'DS%' or program_id like 'AN%' or program_id like 'CG%' or program_id like 'MC%' or program_id like 'WR%' or program_id like 'TM%' or program_id like 'FL%' or program_id = 'RptFileUpload')";//98.08.12 add MIS.MC//102.10.16 add MIS.WR //105.09.06 add MIS.TM/FL
		  		   }else{
		  		      sqlCmd += " and (program_id like 'FX%' or program_id like 'WM%' or program_id like 'ZZ%' or program_id like 'DS%' or program_id = 'RptFileUpload' or program_id like 'BR%' or program_id like 'TM%')";//103.01.20 add BOAF.BR //105.09.06 add BOAF.TM
		  		   }
  		           //updateDBSqlList.add(sqlCmd);
  		           if(!this.updDbUsesPreparedStatement(sqlCmd,paramList)) {
  		        	   throw new Exception();
  		           }
		  		}
		  		//===================================================================
		  		//將原本所授權的報表清掉================================================
		  		//94.07.28 fix 不是MIS時.才清除申報報表
		  		if(!menu.equals("MIS")){
		  			paramList.clear() ;
		  		    sqlCmd = "SELECT * from WTT04_1D WHERE muser_id =? ";
		  		    paramList.add(muser_id) ;
		  		    dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");
		  		    if(dbData.size() != 0){       		  		       
		  		       paramList.clear() ;
		  		       sqlCmd = "DELETE WTT04_1D WHERE muser_id =? ";
		  		       paramList.add(muser_id) ;
				       //updateDBSqlList.add(sqlCmd);
				       if(!this.updDbUsesPreparedStatement(sqlCmd,paramList)){
				    	   throw new Exception() ;
				       }
				    }
				}
		  		//=======================================================================
		  		for(int i=0;i<pgData.size();i++){
		  		    System.out.println((List)pgData.get(i));		  		   	    
		  		    program_id = (String)((List)pgData.get(i)).get(0);		  		   
		  		    p_add = (String)((List)pgData.get(i)).get(1);
		  		    p_delete = (String)((List)pgData.get(i)).get(2);
		  		    p_update = (String)((List)pgData.get(i)).get(3);
		  		    p_query = (String)((List)pgData.get(i)).get(4);
		  		    p_print = (String)((List)pgData.get(i)).get(5);
		  		    p_upload = (String)((List)pgData.get(i)).get(6);
		  		    p_download = (String)((List)pgData.get(i)).get(7);
		  		    p_lock = (String)((List)pgData.get(i)).get(8);
		  		    p_other = (String)((List)pgData.get(i)).get(9);
		  		    if( p_add.equals("Y") || p_delete.equals("Y") || p_update.equals("Y")
		  		    ||  p_query.equals("Y") || p_print.equals("Y") || p_upload.equals("Y")
		  		    ||  p_download.equals("Y") || p_lock.equals("Y") || p_other.equals("Y")
		  		    ){//有任一項為Y時,則Insert該program_id的權限		  
		  		       paramList.clear() ;
		  		       sqlCmd = "INSERT INTO WTT04 VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,sysdate)";
					   paramList.add(muser_id) ;
		  		       paramList.add(program_id) ;
		  		       paramList.add(p_add) ;
		  		       paramList.add(p_delete) ;
		  		       paramList.add(p_update) ;
		  		       paramList.add(p_query) ;
		  		       paramList.add(p_print);
		  		       paramList.add(p_upload) ;
		  		       paramList.add(p_download) ;
		  		       paramList.add(p_lock) ;
		  		       paramList.add(p_other) ;
		  		       paramList.add(user_id) ;
		  		       paramList.add(user_name) ;
					   //updateDBSqlList.add(sqlCmd); 		  
					   if(!this.updDbUsesPreparedStatement(sqlCmd,paramList)) {
						   throw new Exception() ;
					   }
    	   		    }		  		    
		  		}//enf of pgData
		  		//=========================================================================
		  		//取得上傳權限=============================================================
		  		if(!menu.equals("MIS")){//不是MIS時才做=======================================
		  		int row_upload = 0;
		  		if(t.get("row_upload") != null){		  		   
		  		   row_upload = Integer.parseInt((String)t.get("row_upload"));
		  		}  
		  		System.out.println("row_upload="+row_upload);
		  	    List uploadData = new LinkedList();
		  	    List upload_detailData = null;
		  		for ( int i = 0; i < row_upload; i++) {		  	    		  	  			  
		  		    upload_detailData = new LinkedList();
		  		    if(t.get("upload_isModify_" + (i+1)) != null ) {					  					   
					   if(t.get("uploadData_trans_type_" + (i+1)) != null ) {					  
					      upload_detailData.add((String)t.get("uploadData_trans_type_"+(i+1)));
					   }										
					   if(t.get("uploadData_report_no_" + (i+1)) != null ) {					  
					      upload_detailData.add((String)t.get("uploadData_report_no_"+(i+1)));
					  }
					  uploadData.add(upload_detailData);
					}
					
		  		}	
		  		System.out.println("uploadData.size()="+uploadData.size());
		  		for(int i=0;i<uploadData.size();i++){
		  		    System.out.println("uploadData="+uploadData.get(i));
		  		    program_id = "WMFileUpload";
		  		    detail_type="1";
		  		    
		  		    transfer_type=(String)((List)uploadData.get(i)).get(0);
		  		    report_no = (String)((List)uploadData.get(i)).get(1);
		  		    paramList.clear( ) ;						       
		  		    sqlCmd = "INSERT INTO WTT04_1D VALUES (?,?,?,?,?,?,?,sysdate) ";
					
		  		    paramList.add(muser_id)  ;
		  		    paramList.add(program_id) ;
		  		    paramList.add(detail_type) ;
		  		    paramList.add(transfer_type) ;
		  		    paramList.add(report_no) ;
		  		    paramList.add(user_id) ;
		  		    paramList.add(user_name) ;
		  		    if(!this.updDbUsesPreparedStatement(sqlCmd,paramList)) {
		  		    	throw new Exception();
		  		    }
					//updateDBSqlList.add(sqlCmd); 		
		  		}
		  		//=========================================================================		  		
	            
	            //取得下載權限=============================================================
		  		int row_download = 0;
		  		if( t.get("row_download") != null){
		  		    row_download = Integer.parseInt((String)t.get("row_download"));
		  		}   
		  		System.out.println("row_download="+row_download);
		  	    List downloadData = new LinkedList();
		  	    List download_detailData = null;
		  		for ( int i = 0; i < row_download; i++) {		  	    		  	  			  
		  		    download_detailData = new LinkedList();
		  		    if(t.get("download_isModify_" + (i+1)) != null ) {		
		  		       System.out.println("add"+i);			  					   
					   if(t.get("downloadData_trans_type_" + (i+1)) != null ) {					  
					      download_detailData.add((String)t.get("downloadData_trans_type_"+(i+1)));
					   }										
					   if(t.get("downloadData_report_no_" + (i+1)) != null ) {					  
					      download_detailData.add((String)t.get("downloadData_report_no_"+(i+1)));
					  }
					  downloadData.add(download_detailData);
					}
					
		  		}	
		  		for(int i=0;i<downloadData.size();i++){
		  		    System.out.println(downloadData.get(i));
		  		}
		  		System.out.println("downloadData.size()="+downloadData.size());
		  		for(int i=0;i<downloadData.size();i++){
		  		    System.out.println("downloadData="+downloadData.get(i));
		  		    program_id = "WMFileDownload";
		  		    detail_type="2";
		  		    
		  		    transfer_type=(String)((List)downloadData.get(i)).get(0);
		  		    report_no = (String)((List)downloadData.get(i)).get(1);	
		  		    paramList.clear() ;
		  		    sqlCmd = "INSERT INTO WTT04_1D VALUES (?,?,?,?,?,?,?,sysdate)";
					paramList.add(muser_id) ;
		  		    paramList.add(program_id) ;
		  		    paramList.add(detail_type) ;
		  		    paramList.add(transfer_type) ;
		  		    paramList.add(report_no) ;
		  		    paramList.add(user_id);
					paramList.add(user_name) ;
					if(!this.updDbUsesPreparedStatement(sqlCmd,paramList)) {
						throw new Exception();
					}
		  		    // updateDBSqlList.add(sqlCmd); 		
		  		}
		  		//=========================================================================		  		
	            
	            //取得編輯權限=============================================================
		  		int row_edit = 0 ;
		  		if(t.get("row_edit") != null){
		  		   row_edit = Integer.parseInt((String)t.get("row_edit"));
		  		}
		  		System.out.println("row_edit="+row_edit);
		  	    List editData = new LinkedList();
		  	    List edit_detailData = null;
		  		for ( int i = 0; i < row_edit; i++) {		  	    		  	  			  
		  		    edit_detailData = new LinkedList();
		  		    if(t.get("edit_isModify_" + (i+1)) != null ) {					  					   
					   if(t.get("editData_trans_type_" + (i+1)) != null ) {					  
					      edit_detailData.add((String)t.get("editData_trans_type_"+(i+1)));
					   }										
					   if(t.get("editData_report_no_" + (i+1)) != null ) {					  
					      edit_detailData.add((String)t.get("editData_report_no_"+(i+1)));
					  }
					  editData.add(edit_detailData);
					}
					
		  		}	
		  		System.out.println("editData.size()="+editData.size());
		  		for(int i=0;i<editData.size();i++){
		  		    System.out.println("editData="+editData.get(i));
		  		    program_id = "WMFileEdit";
		  		    detail_type="3";
		  		    
		  		    transfer_type=(String)((List)editData.get(i)).get(0);
		  		    report_no = (String)((List)editData.get(i)).get(1);		  		    
		  		    paramList.clear() ;
		  		    sqlCmd = "INSERT INTO WTT04_1D VALUES (?,?,?,?,?,?,?,sysdate)";
					      	
		  		    paramList.add(muser_id) ;
		  		    paramList.add(program_id);
		  		    paramList.add(detail_type) ;
		  		    paramList.add(transfer_type) ;
		  		    paramList.add(report_no);
	  		    	paramList.add(user_id) ;
		  		    paramList.add(user_name) ;
					//   updateDBSqlList.add(sqlCmd);
					if(!this.updDbUsesPreparedStatement(sqlCmd,paramList)){
						throw new Exception();
					}
		  		}
		  		//=========================================================================		  		
	            
	            //取得查詢權限=============================================================
		  		int row_query = 0;
		  		if(t.get("row_query") != null){
		  		   row_query = Integer.parseInt((String)t.get("row_query"));
		  		}  
		  		System.out.println("row_query="+row_query);
		  	    List queryData = new LinkedList();
		  	    List query_detailData = null;
		  		for ( int i = 0; i < row_query; i++) {		  	    		  	  			  
		  		    query_detailData = new LinkedList();
		  		    if(t.get("query_isModify_" + (i+1)) != null ) {					  					   
					   if(t.get("queryData_trans_type_" + (i+1)) != null ) {					  
					      query_detailData.add((String)t.get("queryData_trans_type_"+(i+1)));
					   }										
					   if(t.get("queryData_report_no_" + (i+1)) != null ) {					  
					      query_detailData.add((String)t.get("queryData_report_no_"+(i+1)));
					  }
					  queryData.add(query_detailData);
					}
					
		  		}	
		  		System.out.println("queryData.size()="+queryData.size());
		  		for(int i=0;i<queryData.size();i++){
		  		    System.out.println("queryData="+queryData.get(i));
		  		    program_id = "WMFileQuery";
		  		    detail_type="4";
		  		    
		  		    transfer_type=(String)((List)queryData.get(i)).get(0);
		  		    report_no = (String)((List)queryData.get(i)).get(1);
		  		    paramList.clear() ;
		  		    sqlCmd = "INSERT INTO WTT04_1D VALUES (?,?,?,?,?,?,?,sysdate)";
					
		  		    paramList.add(muser_id) ;
		  		    paramList.add(program_id);
		  		    paramList.add(detail_type) ;
		  		    paramList.add(transfer_type);
		  		    paramList.add(report_no) ;
		  		    paramList.add(user_id) ;
		  		    paramList.add(user_name) ;
		  		    if(!this.updDbUsesPreparedStatement(sqlCmd,paramList)) {
		  		    	throw new Exception();
		  		    }
					 //  updateDBSqlList.add(sqlCmd); 		
		  		}
		  		//=========================================================================		  		
	            }//不是MIS時.才做	
	            errMsg = "相關資料寫入資料庫成功";	
			   /* if(DBManager.updateDB(updateDBSqlList)){					 
				   errMsg = errMsg + "相關資料寫入資料庫成功";					
			    }else{
			 	   errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
			    }*/				    
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = "相關資料寫入資料庫失敗";					
		}	

		return errMsg;
	} 
	//94.03.08 fix 刪除時,整個muser_id的都刪除 
	public String DeleteDB(HttpServletRequest request,String muser_id,String lguser_id,String lguser_name) throws Exception{    	
	    String sqlCmd = "";		
		String errMsg="";
	    String user_id=lguser_id;
	    String user_name=lguser_name;			 
	    //List updateDBSqlList = new LinkedList();
	    List paramList = new ArrayList() ;
	    try {			    			    			    
				sqlCmd = "Delete WTT04 where muser_id=?";		  
				paramList.add(muser_id) ;
		  		//updateDBSqlList.add(sqlCmd);
		  		this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
		  		paramList.clear() ;
		  		sqlCmd = " select * from WTT04_1D where muser_id=? ";
		  		paramList.add(muser_id) ;
		  		List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");
		  		paramList.clear() ;
		  		if(dbData != null && dbData.size() != 0){
		  		   sqlCmd = "DELETE WTT04_1D WHERE muser_id=? ";
		  		   paramList.add(muser_id) ;
				   //updateDBSqlList.add(sqlCmd);
				   this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
				}   
		  		errMsg = errMsg + "相關資料寫入資料庫成功";
				/*if(DBManager.updateDB(updateDBSqlList)){					 
				   errMsg = errMsg + "相關資料寫入資料庫成功";					
			    }else{
			 	   errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
			    }*/				    
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";				
		}	

		return errMsg;		 
	}
	
	public String DeleteDB_old(HttpServletRequest request,String muser_id,String lguser_id,String lguser_name) throws Exception{    	
		String sqlCmd = "";		
		String errMsg="";
	    String user_id=lguser_id;
	    String user_name=lguser_name;			 
	    String program_id="";
	    String p_add="";
	    String p_delete="";
	    String p_update="";
	    String p_query="";
	    String p_print="";
	    String p_upload="";
	    String p_download="";
	    String p_lock="";
	    String p_other="";   	    
	    String transfer_type="";
	    String report_no="";
	    String detail_type="";
		//List updateDBSqlList = new LinkedList();
		List paramList =new ArrayList() ;
		try {			    			    			    
			    System.out.println("delete begin");
			    
			    //取出form裡的所有變數=================================== 
		  		Enumeration ep = request.getParameterNames();
		  		Enumeration ea = request.getAttributeNames();
		  		Hashtable t = new Hashtable();
		  		String name = "";
		  
		  		for ( ; ep.hasMoreElements() ; ) {
			   		name = (String)ep.nextElement();
			   		t.put( name, request.getParameter(name) );		
			   		System.out.println(name+"="+request.getParameter(name));	   
		  		}		
		  		System.out.println("deleteuser="+muser_id);
		  		//取得程式基本功能權限=======================================================  		  		
		  		int row_program =Integer.parseInt((String)t.get("row_program"));		  				  		
		  		System.out.println("row_program="+row_program);
		  	    List pgData = new LinkedList();
		  	    List programidData = null;
		  		for ( int i = 0; i < row_program; i++) {		  	    		  	  			  
		  		    programidData = new LinkedList();
		  		    if(t.get("program_id_" + (i+1)) != null ) {					  
					   programidData.add((String)t.get("program_id_"+(i+1)));
					}		  		   
					if(t.get("P_ADD_" + (i+1)) != null ) {					  
					   programidData.add((String)t.get("P_ADD_"+(i+1)));
					}else{
					   programidData.add("N");
					}										
					if(t.get("P_DELETE_" + (i+1)) != null ) {					  
					   programidData.add((String)t.get("P_DELETE_"+(i+1)));
					}else{
					   programidData.add("N");
					}
					if(t.get("P_UPDATE_" + (i+1)) != null ) {		
					   System.out.println("P_UPDATE+"+(i+1)+"="+(String)t.get("P_UPDATE_"+(i+1)));			  
					   programidData.add((String)t.get("P_UPDATE_"+(i+1)));
					}else{
					   programidData.add("N");
					}
					if(t.get("P_QUERY_" + (i+1)) != null ) {					  
					   programidData.add((String)t.get("P_QUERY_"+(i+1)));
					}else{
					   programidData.add("N");
					}
					if(t.get("P_PRINT_" + (i+1)) != null ) {					  
					   programidData.add((String)t.get("P_PRINT_"+(i+1)));
					}else{
					   programidData.add("N");
					}
					if(t.get("P_UPLOAD_" + (i+1)) != null ) {					  
					   programidData.add((String)t.get("P_UPLOAD_"+(i+1)));
					}else{
					   programidData.add("N");
					}
					if(t.get("P_DOWNLOAD_" + (i+1)) != null ) {					  
					   programidData.add((String)t.get("P_DOWNLOAD_"+(i+1)));
					}else{
					   programidData.add("N");
					}
					if(t.get("P_LOCK_" + (i+1)) != null ) {					  
					   programidData.add((String)t.get("P_LOCK_"+(i+1)));
					}else{
					   programidData.add("N");
					}
					if(t.get("P_OTHER_" + (i+1)) != null ) {					  
					   programidData.add((String)t.get("P_OTHER_"+(i+1)));
					}else{
					   programidData.add("N");
					}
					pgData.add(programidData);
		  		}	
		  		System.out.println("pgData.size="+pgData.size());
		  		for(int i=0;i<pgData.size();i++){
		  		    System.out.println((List)pgData.get(i));		  		   	    
		  		    program_id = (String)((List)pgData.get(i)).get(0);		  		   
		  		    p_add = (String)((List)pgData.get(i)).get(1);
		  		    p_delete = (String)((List)pgData.get(i)).get(2);
		  		    p_update = (String)((List)pgData.get(i)).get(3);
		  		    p_query = (String)((List)pgData.get(i)).get(4);
		  		    p_print = (String)((List)pgData.get(i)).get(5);
		  		    p_upload = (String)((List)pgData.get(i)).get(6);
		  		    p_download = (String)((List)pgData.get(i)).get(7);
		  		    p_lock = (String)((List)pgData.get(i)).get(8);
		  		    p_other = (String)((List)pgData.get(i)).get(9);
		  		    if( p_add.equals("N") && p_delete.equals("N") && p_update.equals("N")
		  		    &&  p_query.equals("N") && p_print.equals("N") && p_upload.equals("N")
		  		    &&  p_download.equals("N") && p_lock.equals("N") && p_other.equals("N")
		  		    ){//全部都為N時,則將此程式代號的權限刪除===========================
		  		      paramList.clear() ;
		  		      sqlCmd = "Delete WTT04 where muser_id=?"
		  		             + " and program_id=?";
		  		      paramList.add(muser_id) ;
		  		      paramList.add(program_id) ;
		  		      //updateDBSqlList.add(sqlCmd);
		  		      if(!this.updDbUsesPreparedStatement(sqlCmd,paramList)) {
		  		    	  throw new Exception();
		  		      }
		  		    }
		  		    if( p_add.equals("Y") || p_delete.equals("Y") || p_update.equals("Y")
		  		    ||  p_query.equals("Y") || p_print.equals("Y") || p_upload.equals("Y")
		  		    ||  p_download.equals("Y") || p_lock.equals("Y") || p_other.equals("Y")
		  		    ){//有任一項為Y時,則Update該program_id的權限
		  		       paramList.clear() ;
		  		       sqlCmd = "Update WTT04 SET "
					          + "p_add = ?" 
					       	  + ",p_delete=?" 
					          + ",p_update=?" 
					          + ",p_query=?" 
					          + ",p_print=?" 
					          + ",p_upload=?" 
					          + ",p_download=?" 
					          + ",p_lock=?" 
					          + ",p_other=?" 					       
					          + ",user_id=?"					       
					          + ",user_name=?"
					          + ",update_date = sysdate"
					          + " where muser_id=?"
					          + " and program_id=?";
		  		       paramList.add(p_add) ;
		  		       paramList.add(p_delete) ;
		  		       paramList.add(p_update) ;
		  		       paramList.add(p_query) ;
		  		       paramList.add(p_print) ;
		  		       paramList.add(p_upload);
		  		       paramList.add(p_download);
		  		       paramList.add(p_lock) ;
		  		       paramList.add(p_other) ;
		  		       paramList.add(user_id) ;
		  		       paramList.add(user_name) ;
		  		       paramList.add(muser_id) ;
		  		       paramList.add(program_id) ;
					   //updateDBSqlList.add(sqlCmd); 	
					   if(!this.updDbUsesPreparedStatement(sqlCmd,paramList)){
						   throw new Exception();
					   }
    	   		    }		  		    
		  		}//enf of pgData
		  		//=========================================================================
		  		//取得欲刪除的上傳權限=============================================================
		  		int row_upload = 0;
		  		if(t.get("row_upload") != null){
		  		   row_upload = Integer.parseInt((String)t.get("row_upload"));
		  		}   
		  		System.out.println("row_upload="+row_upload);
		  	    List uploadData = new LinkedList();
		  	    List upload_detailData = null;
		  		for ( int i = 0; i < row_upload; i++) {		  	    		  	  			  
		  		    upload_detailData = new LinkedList();
		  		    if(t.get("upload_isModify_" + (i+1)) != null ) {					  					   
					   if(t.get("uploadData_trans_type_" + (i+1)) != null ) {					  
					      upload_detailData.add((String)t.get("uploadData_trans_type_"+(i+1)));
					   }										
					   if(t.get("uploadData_report_no_" + (i+1)) != null ) {					  
					      upload_detailData.add((String)t.get("uploadData_report_no_"+(i+1)));
					  }
					  uploadData.add(upload_detailData);
					}
					
		  		}	
		  		System.out.println("uploadData.size()="+uploadData.size());
		  		
		  		for(int i=0;i<uploadData.size();i++){		  		
		  		    System.out.println("uploadData="+uploadData.get(i));
		  		    program_id = "WMFileUpload";
		  		    detail_type="1";
		  		    
		  		    transfer_type=(String)((List)uploadData.get(i)).get(0);
		  		    report_no = (String)((List)uploadData.get(i)).get(1);
		  		    paramList.clear() ;
		  		    sqlCmd = "DELETE WTT04_1D WHERE muser_id='"+muser_id+"'" 			        	   		  		        
					       + " and program_id = '" + program_id +"'" 
					       + " and detail_type = '" + detail_type + "'" 					       
					       + " and transfer_type = '" + transfer_type + "'" 
					       + " and report_no = '" + report_no + "'";
		  		    paramList.add(muser_id) ;
		  		    paramList.add(program_id) ;
		  		    paramList.add(detail_type) ;
		  		    paramList.add(transfer_type) ;
		  		    paramList.add(report_no) ;
					//updateDBSqlList.add(sqlCmd); 		
					if(!this.updDbUsesPreparedStatement(sqlCmd,paramList)){
						throw new Exception();
					}
		  		}
		  		//=========================================================================		  		
	            
	            //取得欲刪除的下載權限=============================================================
		  		int row_download = 0 ;
		  		if(t.get("row_download") != null){
		  		  row_download = Integer.parseInt((String)t.get("row_download"));
		  		} 
		  		System.out.println("row_download="+row_download);
		  	    List downloadData = new LinkedList();
		  	    List download_detailData = null;
		  		for ( int i = 0; i < row_download; i++) {		  	    		  	  			  
		  		    download_detailData = new LinkedList();
		  		    if(t.get("download_isModify_" + (i+1)) != null ) {		
		  		       System.out.println("add"+i);			  					   
					   if(t.get("downloadData_trans_type_" + (i+1)) != null ) {					  
					      download_detailData.add((String)t.get("downloadData_trans_type_"+(i+1)));
					   }										
					   if(t.get("downloadData_report_no_" + (i+1)) != null ) {					  
					      download_detailData.add((String)t.get("downloadData_report_no_"+(i+1)));
					  }
					  downloadData.add(download_detailData);
					}
					
		  		}	
		  		for(int i=0;i<downloadData.size();i++){
		  		    System.out.println(downloadData.get(i));
		  		}
		  		System.out.println("downloadData.size()="+downloadData.size());
		  		for(int i=0;i<downloadData.size();i++){
		  		    System.out.println("downloadData="+downloadData.get(i));
		  		    program_id = "WMFileDownload";
		  		    detail_type="2";
		  		    
		  		    transfer_type=(String)((List)downloadData.get(i)).get(0);
		  		    report_no = (String)((List)downloadData.get(i)).get(1);
		  		    paramList.clear() ;
		  		    sqlCmd = "DELETE WTT04_1D WHERE muser_id=?" 			        	   		  		        
					       + " and program_id = ?" 
					       + " and detail_type = ?" 					       
					       + " and transfer_type = ?" 
					       + " and report_no = ?";
		  		    paramList.add(muser_id) ;
		  		    paramList.add(program_id) ;
		  		    paramList.add(detail_type) ;
		  		    paramList.add(transfer_type) ;
		  		    paramList.add(report_no) ;
		  		    if(!this.updDbUsesPreparedStatement(sqlCmd,paramList)){
		  		    	throw new Exception();
		  		    }
					//updateDBSqlList.add(sqlCmd); 		
		  		}
		  		//=========================================================================		  		
	            
	            //取得欲刪除的編輯權限=============================================================
		  		int row_edit = 0 ;
		  		if(t.get("row_edit") != null){
		  		    row_edit = Integer.parseInt((String)t.get("row_edit"));
		  		}   
		  		System.out.println("row_edit="+row_edit);
		  	    List editData = new LinkedList();
		  	    List edit_detailData = null;
		  		for ( int i = 0; i < row_edit; i++) {		  	    		  	  			  
		  		    edit_detailData = new LinkedList();
		  		    if(t.get("edit_isModify_" + (i+1)) != null ) {					  					   
					   if(t.get("editData_trans_type_" + (i+1)) != null ) {					  
					      edit_detailData.add((String)t.get("editData_trans_type_"+(i+1)));
					   }										
					   if(t.get("editData_report_no_" + (i+1)) != null ) {					  
					      edit_detailData.add((String)t.get("editData_report_no_"+(i+1)));
					  }
					  editData.add(edit_detailData);
					}
					
		  		}	
		  		System.out.println("editData.size()="+editData.size());
		  		for(int i=0;i<editData.size();i++){
		  		    System.out.println("editData="+editData.get(i));
		  		    program_id = "WMFileEdit";
		  		    detail_type="3";
		  		    
		  		    transfer_type=(String)((List)editData.get(i)).get(0);
		  		    report_no = (String)((List)editData.get(i)).get(1);
		  		    paramList.clear() ;
		  		    sqlCmd = "DELETE WTT04_1D WHERE muser_id=?" 			        	   		  		        
					       + " and program_id = ?" 
					       + " and detail_type = ?" 					       
					       + " and transfer_type =?" 
					       + " and report_no = ?";
		  		    paramList.add(muser_id) ;
		  		    paramList.add(program_id) ;
		  		    paramList.add(detail_type) ;
		  		    paramList.add(transfer_type) ;
		  		    paramList.add(report_no) ;
		  		    if(!this.updDbUsesPreparedStatement(sqlCmd,paramList)) {
		  		    	throw new Exception();
		  		    }
				   //updateDBSqlList.add(sqlCmd); 		
		  		}
		  		//=========================================================================		  		
	            
	            //取得欲刪除的查詢權限=============================================================
		  		int row_query = 0 ;
		  		if(t.get("row_query") != null){
		  		  row_query = Integer.parseInt((String)t.get("row_query"));
		  		}  
		  		System.out.println("row_query="+row_query);
		  	    List queryData = new LinkedList();
		  	    List query_detailData = null;
		  		for ( int i = 0; i < row_query; i++) {		  	    		  	  			  
		  		    query_detailData = new LinkedList();
		  		    if(t.get("query_isModify_" + (i+1)) != null ) {					  					   
					   if(t.get("queryData_trans_type_" + (i+1)) != null ) {					  
					      query_detailData.add((String)t.get("queryData_trans_type_"+(i+1)));
					   }										
					   if(t.get("queryData_report_no_" + (i+1)) != null ) {					  
					      query_detailData.add((String)t.get("queryData_report_no_"+(i+1)));
					  }
					  queryData.add(query_detailData);
					}
					
		  		}	
		  		System.out.println("queryData.size()="+queryData.size());
		  		for(int i=0;i<queryData.size();i++){
		  		    System.out.println("queryData="+queryData.get(i));
		  		    program_id = "WMFileQuery";
		  		    detail_type="4";
		  		    
		  		    transfer_type=(String)((List)queryData.get(i)).get(0);
		  		    report_no = (String)((List)queryData.get(i)).get(1);
		  		    paramList.clear() ;
		  		    sqlCmd = "DELETE WTT04_1D WHERE muser_id=?" 			        	   		  		        
					       + " and program_id = ?" 
					       + " and detail_type = ?" 					       
					       + " and transfer_type =?" 
					       + " and report_no = ?";
		  		    paramList.add(muser_id) ;
		  		    paramList.add(program_id) ;
		  		    paramList.add(detail_type) ;
		  		    paramList.add(transfer_type) ;
		  		    paramList.add(report_no) ;
		  		    if(!this.updDbUsesPreparedStatement(sqlCmd,paramList)) {
		  		    	throw new Exception();
		  		    }
					//updateDBSqlList.add(sqlCmd); 		
		  		}
		  		//=========================================================================		  		
	            errMsg = errMsg + "相關資料寫入資料庫成功";		
			    /*if(DBManager.updateDB(updateDBSqlList)){					 
				   errMsg = errMsg + "相關資料寫入資料庫成功";					
			    }else{
			 	   errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
			    }*/				    
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