<%
//102.06.07 created by 2968
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>

<%
	RequestDispatcher rd = null;
	String actMsg = "";	
	String alertMsg = "";	
	String webURL = "";	
	boolean doProcess = false;	
	
	//取得session資料,取得成功時,才繼續往下執行===================================================
	if(session.getAttribute("muser_id") == null){//session timeout	
      System.out.println("ZZ006W login timeout");   
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
	String sn_docno0 =( request.getParameter("sn_docno0")==null ) ? "" : (String)request.getParameter("sn_docno0");
	String sn_docno2 =( request.getParameter("sn_docno2")==null ) ? "" : (String)request.getParameter("sn_docno2");
	String sn_docno1 =( request.getParameter("sn_docno1")==null ) ? "" : (String)request.getParameter("sn_docno1");
	String pgId =( request.getParameter("pgId")==null ) ? "" : (String)request.getParameter("pgId");
	String rpt_id =( request.getParameter("rpt_id")==null ) ? "" : (String)request.getParameter("rpt_id");
	String rpt_no =( request.getParameter("rpt_no")==null ) ? "" : (String)request.getParameter("rpt_no");
	System.out.println("act="+act);	
	if(session.getAttribute("muser_id") == null){	
      System.out.println("ZZ006W login timeout");   
	  rd = application.getRequestDispatcher( "/pages/reLogin.jsp?url=LoginError.jsp?timeout=true" );         	   
      rd.forward(request,response);
    }
   
	//登入者資訊
	String lguser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		
	String lguser_name = ( session.getAttribute("muser_name")==null ) ? "" : (String)session.getAttribute("muser_name");		
	String lguser_type = ( session.getAttribute("muser_type")==null ) ? "" : (String)session.getAttribute("muser_type");				
	session.setAttribute("nowtbank_no",null);//94.01.05 fix 沒有Bank_List,把所點選的Bank_no清除======
	
    if(!CheckPermission(request)){//無權限時,導向到LoginError.jsp
        rd = application.getRequestDispatcher( LoginErrorPgName );        
    }else{            
    	//set next jsp 	    	
    	if(act.equals("List") || act.equals("Qry")){    	       	        	     
        	if(act.equals("List")){
        	   rd = application.getRequestDispatcher( ListPgName +"?act=List");                	
        	}    
        	if(act.equals("Qry")){    	    	    
        	   List qList = getQryResult(sn_docno0,sn_docno2,sn_docno1,pgId);    	    
    	       request.setAttribute("qList",qList);
    	       rd = application.getRequestDispatcher( ListPgName +"?act=Qry&sn_docno0="+sn_docno0+"&sn_docno2="+sn_docno2+"&sn_docno1="+sn_docno1+"&pgId="+pgId);            	        	        	    	    
    	    }   
        }else if(act.equals("Edit")){  
            List mtnList = getMtnList(rpt_no);    	    
    	    request.setAttribute("maintainInfo",mtnList);
    	 	rd = application.getRequestDispatcher( EditPgName +"?act=Edit&rpt_id="+rpt_id);       
        }else if(act.equals("Update")){
    	    actMsg = UpdateDB(request);    	    	        	    
        	rd = application.getRequestDispatcher( nextPgName );         
        }else if(act.equals("Delete")){
    	    actMsg = DeleteDB(request);    	    	        	    
        	rd = application.getRequestDispatcher( nextPgName +"?act=List&goPages=ZZ006W.jsp" );                 
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
    private final static String EditPgName = "/pages/ZZ006W_Edit.jsp";    
    private final static String ListPgName = "/pages/ZZ006W_List.jsp";        
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    private boolean CheckPermission(HttpServletRequest request){//檢核權限    	    
		boolean CheckOK=false;
    	HttpSession session = request.getSession();            
        Properties permission = ( session.getAttribute("ZZ006W")==null ) ? new Properties() : (Properties)session.getAttribute("ZZ006W");				                
        if(permission == null){
        	System.out.println("ZZ006W.permission == null");
        }else{
            System.out.println("ZZ006W.permission.size ="+permission.size());
        }
        //只要有Query的權限,就可以進入畫面
        if(permission != null && permission.get("Q") != null && permission.get("Q").equals("Y")){            
        	CheckOK = true;//Query
        }
        System.out.println("CheckOk="+CheckOK);
        return CheckOK;
    }           
    //取得查詢結果
    private List getQryResult(String sn_docno0,String sn_docno2,String sn_docno1,String pgId){
        StringBuffer sqlCmd = new StringBuffer();		
        List paramList = new ArrayList();
        List DBSqlList = new LinkedList();
        List<List> updateDBList = new ArrayList<List>();//0:sql 1:data
        List updateDBSqlList = new ArrayList();
        List<List> updateDBDataList = new ArrayList<List>();//儲存參數的List
		sqlCmd.append("select rpt_main.rpt_no,"); //--報表編號
		sqlCmd.append(" 	  rpt_main.rpt_id, "); //--功能代碼
		sqlCmd.append(" 	  rpt_main.rpt_name, ");//--報表名稱
		if(!"".equals(sn_docno2)||!"".equals(sn_docno1)){ //查詢條件有輸入報表欄位名稱 or 科目代碼才加入此SQL
			sqlCmd.append("   col_name || '=' || rpt_rule as col_name, ");
		}
		sqlCmd.append(" 	  rpt_main.rpt_path "); //--報表功能路徑
		sqlCmd.append("  from rpt_main left join wtt03_1 on rpt_main.rpt_id = wtt03_1.program_id ");
		if(!"".equals(sn_docno2)||!"".equals(sn_docno1)){
			sqlCmd.append("  left join rpt_detail on rpt_main.rpt_id = rpt_detail.rpt_id ");
		}
		sqlCmd.append(" where 1=1 ");
		if(!"".equals(sn_docno0)){
			sqlCmd.append("   and rpt_main.rpt_name like ? ");// 'UI.報表名稱 ex:經營指標.每個字前後加% ex:%經%營%指%標%'
			String tmpStr ="";
			for(int i=0;i<sn_docno0.length();i++){
			    tmpStr += ("%"+sn_docno0.substring(i, i+1));
			}
			paramList.add(tmpStr+"%");
		}
		if(!"".equals(sn_docno2)){
			sqlCmd.append("   and rpt_detail.col_name like ? ");
			paramList.add("%"+sn_docno2+"%");
		}
		if(!"".equals(sn_docno1)){
			sqlCmd.append("   and rpt_rule like ? ");
			paramList.add("%"+sn_docno1+"%");
		}
		if(!"".equals(pgId)){
			sqlCmd.append("   and rpt_id like ? ");
			paramList.add("%"+pgId+"%");
		}
		sqlCmd.append(" order by program_id");
		if(!"".equals(sn_docno2)||!"".equals(sn_docno1)){
			sqlCmd.append(" ,col_name ");
		}
		List qData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"rpt_no,rpt_id,rpt_name,col_name,rpt_path"); 
		System.out.println("### getQryResult.size="+qData.size());
        return qData;
    }
    
    private List getMtnList(String rpt_no){
        StringBuffer sqlCmd = new StringBuffer();		
        List paramList = new ArrayList();
        List DBSqlList = new LinkedList();
        List<List> updateDBList = new ArrayList<List>();//0:sql 1:data
        List updateDBSqlList = new ArrayList();
        List<List> updateDBDataList = new ArrayList<List>();//儲存參數的List
        sqlCmd.append("select rpt_main.rpt_no,rpt_main.rpt_id,");//功能代碼
        sqlCmd.append("	      rpt_name,");//報表抬頭名稱
        sqlCmd.append("       rpt_path,");//報表功能路徑
        sqlCmd.append("       col_name || '=' || rpt_rule as col_name ");//報表內欄位名稱
        sqlCmd.append("  from rpt_main left join rpt_detail on rpt_main.rpt_id =rpt_detail.rpt_id ");
        sqlCmd.append(" where rpt_main.rpt_no= ? ");
        paramList.add(rpt_no);
		List mtnData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"rpt_no,rpt_id,rpt_name,rpt_path,col_name"); 
		System.out.println("### getMtnList.size="+mtnData.size());
        return mtnData;
    }
	
    private boolean haveData(String rpt_no){
        StringBuffer sqlCmd = new StringBuffer();		
        List paramList = new ArrayList();
        List DBSqlList = new LinkedList();
        List<List> updateDBList = new ArrayList<List>();//0:sql 1:data
        List updateDBSqlList = new ArrayList();
        List<List> updateDBDataList = new ArrayList<List>();//儲存參數的List
        boolean flag = false;
        sqlCmd.append("select * from rpt_main ");
        sqlCmd.append(" where rpt_main.rpt_no= ? ");
        paramList.add(rpt_no);
		List qData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,""); 
		System.out.println("### haveData.size="+qData.size());
		if(qData.size()>0){
		    flag =true;
		}
		return flag;
    }
	private String UpdateDB(HttpServletRequest request) throws Exception{    	
	    StringBuffer sqlCmd = new StringBuffer();		
        List paramList = new ArrayList();
        List DBSqlList = new LinkedList();
        List<List> updateDBList = new ArrayList<List>();//0:sql 1:data
        List updateDBSqlList = new ArrayList();
        List<List> updateDBDataList = new ArrayList<List>();//儲存參數的List
		String errMsg="";
		String rpt_no=((String)request.getParameter("rpt_no")==null)?"":(String)request.getParameter("rpt_no");
		String rpt_id=((String)request.getParameter("rpt_id")==null)?"":(String)request.getParameter("rpt_id");
		String rpt_name=((String)request.getParameter("rpt_name")==null)?"":(String)request.getParameter("rpt_name");
		String rpt_path=((String)request.getParameter("rpt_path")==null)?" ":(String)request.getParameter("rpt_path");
		String rpt_idOri=((String)request.getParameter("rpt_idOri")==null)?"":(String)request.getParameter("rpt_idOri");
		try {
			if(haveData(rpt_no)==true){
			    sqlCmd.append(" update rpt_main set rpt_id=?, rpt_name=?, rpt_path =? where rpt_no =? ");
         		paramList.add(rpt_id);
         		paramList.add(rpt_name);
         		paramList.add(rpt_path);
         		paramList.add(rpt_no);
         		updateDBSqlList.add(sqlCmd.toString()) ;
			    updateDBDataList.add(paramList) ;
			    updateDBSqlList.add(updateDBDataList) ;
			    updateDBList.add(updateDBSqlList) ;
			    sqlCmd.setLength(0) ;
			    paramList = new ArrayList () ;
			    updateDBSqlList = new ArrayList() ;
			    updateDBDataList = new ArrayList<List> () ;
			    sqlCmd.append(" update rpt_detail set rpt_id=? where rpt_id =? ");
         		paramList.add(rpt_id);
         		paramList.add(rpt_idOri);
         		updateDBSqlList.add(sqlCmd.toString()) ;
			    updateDBDataList.add(paramList) ;
			    updateDBSqlList.add(updateDBDataList) ;
			    updateDBList.add(updateDBSqlList) ;
			    if(DBManager.updateDB_ps(updateDBList)){
			        errMsg = errMsg + "相關資料寫入資料庫成功";
      			}else{
      			  errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
      			}
			}else{
			    errMsg = errMsg + "此筆資料不存在無法修改<br>";
			}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗<br>[Exception Error]:<br>"+e.getMessage();								
		}	

		return errMsg;
	} 
    
    private String DeleteDB(HttpServletRequest request) throws Exception{    	
        StringBuffer sqlCmd = new StringBuffer();		
        List paramList = new ArrayList();
        List DBSqlList = new LinkedList();
        List<List> updateDBList = new ArrayList<List>();//0:sql 1:data
        List updateDBSqlList = new ArrayList();
        List<List> updateDBDataList = new ArrayList<List>();//儲存參數的List
		String errMsg="";
		String rpt_no=((String)request.getParameter("rpt_no")==null)?"":(String)request.getParameter("rpt_no");
		String rpt_id=((String)request.getParameter("rpt_id")==null)?"":(String)request.getParameter("rpt_id");
		String rpt_name=((String)request.getParameter("rpt_name")==null)?"":(String)request.getParameter("rpt_name");
		String rpt_path=((String)request.getParameter("rpt_path")==null)?" ":(String)request.getParameter("rpt_path");
		String rpt_idOri=((String)request.getParameter("rpt_idOri")==null)?"":(String)request.getParameter("rpt_idOri");
		try {
		    
			if(haveData(rpt_no)==true){
			    sqlCmd.append(" delete rpt_main where rpt_no =? ");
         		paramList.add(rpt_no);
         		updateDBSqlList.add(sqlCmd.toString()) ;
			    updateDBDataList.add(paramList) ;
			    updateDBSqlList.add(updateDBDataList) ;
			    updateDBList.add(updateDBSqlList) ;
			    sqlCmd.setLength(0) ;
			    paramList = new ArrayList () ;
			    updateDBSqlList = new ArrayList() ;
			    updateDBDataList = new ArrayList<List> () ;
			    sqlCmd.append(" delete rpt_detail where rpt_id =? ");
         		paramList.add(rpt_idOri);
         		updateDBSqlList.add(sqlCmd.toString()) ;
			    updateDBDataList.add(paramList) ;
			    updateDBSqlList.add(updateDBDataList) ;
			    updateDBList.add(updateDBSqlList) ;
			    if(DBManager.updateDB_ps(updateDBList)){
			        errMsg = errMsg + "刪除資料庫相關資料成功";
      			}else{
      			  errMsg = errMsg + "刪除資料庫相關資料失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
      			}
			}else{
			    errMsg = errMsg + "此筆資料不存在無法修改<br>";
			}	
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗<br>[Exception Error]:<br>"+e.getMessage();								
		}	

		return errMsg;
	}
	
%>    