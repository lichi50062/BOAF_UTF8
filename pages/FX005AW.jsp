<%
//94.11.14 first designed by 4180
//95.04.04 依增修功能案會議記錄fix by 2495
//99.12.03 fix sqlInjection by 2808
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
	String webURL_Y = "";
	boolean doProcess = false;

	//取得session資料,取得成功時,才繼續往下執行===================================================
	if(session.getAttribute("muser_id") == null){//session timeout
      System.out.println("FX005AW login timeout");
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
	String seq_no = (request.getParameter("seq_no")==null)?"":(String)request.getParameter("seq_no");
	System.out.println("-------------seq_no:"+seq_no);
	String checkseq =(request.getParameter("CHECKSEQ")==null)?"":(String)request.getParameter("CHECKSEQ");
	
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
String sort_type = ( request.getParameter("sort_type")==null ) ? " " : (String)request.getParameter("sort_type");	
	if(sort_type.equals(" ")){
		sort_type="0";
	}
	System.out.println("sort_type:"+sort_type);
     if(!Utility.CheckPermission(request,"FX005AW")){//無權限時,導向到LoginError.jsp     
        rd = application.getRequestDispatcher( LoginErrorPgName );
    }else{
    	if(act.equals("new")){
		rd = application.getRequestDispatcher( EditPgName +"?act=new");
    	}else if(act.equals("Edit")){
    		  List dbData = getWLX05_ATM_SETUP(bank_no,seq_no);
    	    request.setAttribute("WLX05_ATM_SETUP",dbData);
        	request.setAttribute("maintainInfo","select * from WLX05_ATM_SETUP WHERE bank_no='" + bank_no+"'");
        	rd = application.getRequestDispatcher( EditPgName +"?act=Edit");
    	}else if(act.equals("List")){
    		System.out.println("----------sort_type:"+sort_type);
    		 if(sort_type.equals("0"))  
    		 {
    		  List dbData = getWLX05_ATM_SETUP(bank_no,"");
    		  request.setAttribute("WLX05_ATM_SETUP",dbData);
        	rd = application.getRequestDispatcher( ListPgName +"?bank_no="+bank_no);    	
				 }				 
				 if(sort_type.equals("1"))
    		  {
    		  	List 	dbData = getWLX05_ATM_SETUP_SORT2(bank_no,"","PROPERTY_NO","CANCEL_TYPE");
    		  	request.setAttribute("WLX05_ATM_SETUP",dbData);
        	  rd = application.getRequestDispatcher( ListPgName +"?bank_no="+bank_no);  
    		  }
  
    		  if(sort_type.equals("2"))
    		  {
    		  	List dbData = getWLX05_ATM_SETUP_SORT2(bank_no,"","MACHINE_NAME","CANCEL_TYPE");
    		  	request.setAttribute("WLX05_ATM_SETUP",dbData);
        		rd = application.getRequestDispatcher( ListPgName +"?bank_no="+bank_no);  
    		  }
    		  if(sort_type.equals("3"))
    		  {
    		  	List dbData = getWLX05_ATM_SETUP_SORT2(bank_no,"","HSIEN_ID","SETUP_DATE");
    		  	request.setAttribute("WLX05_ATM_SETUP",dbData);
        		rd = application.getRequestDispatcher( ListPgName +"?bank_no="+bank_no);  
    		  }
    		  if(sort_type.equals("4"))
    		  {
    		  	List dbData = getWLX05_ATM_SETUP_SORT2(bank_no,"","SETUP_DATE","HSIEN_ID");
    		  	request.setAttribute("WLX05_ATM_SETUP",dbData);
        		rd = application.getRequestDispatcher( ListPgName +"?bank_no="+bank_no);  
    		  }
    		  
    		  if(sort_type.equals("5"))
    		  {
    		  	List dbData = getWLX05_ATM_SETUP_SORT3(bank_no,"","CANCEL_TYPE","SETUP_DATE","HSIEN_ID");
    		 		request.setAttribute("WLX05_ATM_SETUP",dbData);
        		rd = application.getRequestDispatcher( ListPgName +"?bank_no="+bank_no);  
    		  }
				  
				 
		}else if(act.equals("Insert")){
    	        actMsg = InsertDB(request,bank_no,lguser_id,lguser_name);
        	    rd = application.getRequestDispatcher( nextPgName+"?FX=FX005AW" );
    	}else if(act.equals("Update")){
    	        actMsg = UpdateDB(request,bank_no,lguser_id,lguser_name);
        	rd = application.getRequestDispatcher( nextPgName+"?FX=FX005AW");
    	}else if(act.equals("Delete")){
    	        actMsg = DeleteDB(request,bank_no,lguser_id,lguser_name);
        	rd = application.getRequestDispatcher( nextPgName+"?FX=FX005AW" );
    	}else if(act.equals("Add_Check_Property_No")){
    	        String property_no =  ( request.getParameter("property_no")==null ) ? "" : (String)request.getParameter("property_no");						   
							bank_no =  ( request.getParameter("bank_no")==null ) ? "" : (String)request.getParameter("bank_no");								 
							String cancel_type =  ( request.getParameter("cancel_type")==null ) ? "" : (String)request.getParameter("cancel_type");							
							System.out.println("Add_Check_Property_No_cancel_type ="+cancel_type);
							String site_name =  ( request.getParameter("site_name")==null ) ? "" : (String)request.getParameter("site_name");						   
							System.out.println("site_name ="+site_name);
							String addr =  ( request.getParameter("addr")==null ) ? "" : (String)request.getParameter("addr");						   
							System.out.println("addr ="+addr);
							String setup_date_y =  ( request.getParameter("setup_date_y")==null ) ? "" : (String)request.getParameter("setup_date_y");						   
							System.out.println("setup_date_y ="+setup_date_y);
							String setup_date_m =  ( request.getParameter("setup_date_m")==null ) ? "" : (String)request.getParameter("setup_date_m");						   
							System.out.println("setup_date_m ="+setup_date_m);
							String setup_date_d =  ( request.getParameter("setup_date_d")==null ) ? "" : (String)request.getParameter("setup_date_d");						   
							System.out.println("setup_date_d ="+setup_date_d);
							String machine_name =  ( request.getParameter("machine_name")==null ) ? "" : (String)request.getParameter("machine_name");						   
							System.out.println("machine_name ="+machine_name);
							
							List dbData_Check_Property_No = Check_Add_Property_No(property_no,bank_no,cancel_type);							
							String check_var = (((DataObject)dbData_Check_Property_No.get(0)).getValue("check_var1")).toString();
							System.out.println("check_var = "+check_var); 
							if(!check_var.equals("0"))
							{	
							  						  
							  if((cancel_type.equals("0")||cancel_type.equals("1")))
							  {
							  	System.out.println("Add_Check_Property_No_cancel_type ="+cancel_type);
							    actMsg = InsertDB(request,bank_no,lguser_id,lguser_name);
        	    	  rd = application.getRequestDispatcher( nextPgName+"?FX=FX005AW" );						  
							  }
							  
							  if(cancel_type.equals(""))
							  { 
									alertMsg = "該機器編號己裝設，尚未遷移/裁撤，不可重覆建檔!!!";
									webURL_Y = "/pages/FX005AW_Edit.jsp?act=new&property_no="+property_no+"&site_name="+site_name+"&addr="+addr+"&setup_date_y="+setup_date_y+"&setup_date_m="+setup_date_m+"&setup_date_d="+setup_date_d+"&machine_name="+machine_name;
									rd = application.getRequestDispatcher( nextPgName );			   
							  }
							}else
						  {
						  	actMsg = InsertDB(request,bank_no,lguser_id,lguser_name);
        	    	rd = application.getRequestDispatcher( nextPgName+"?FX=FX005AW" );
						  }	
    	}
    	else if(act.equals("Modify_Check_Property_No")){
    		      
    	        String property_no =  ( request.getParameter("property_no")==null ) ? "" : (String)request.getParameter("property_no");						   
							bank_no =  ( request.getParameter("bank_no")==null ) ? "" : (String)request.getParameter("bank_no");								 
							String cancel_type =  ( request.getParameter("cancel_type")==null ) ? "" : (String)request.getParameter("cancel_type");							
							String site_name =  ( request.getParameter("site_name")==null ) ? "" : (String)request.getParameter("site_name");						   
							System.out.println("site_name ="+site_name);
							String addr =  ( request.getParameter("addr")==null ) ? "" : (String)request.getParameter("addr");						   
							System.out.println("addr ="+addr);
							String setup_date_y =  ( request.getParameter("setup_date_y")==null ) ? "" : (String)request.getParameter("setup_date_y");						   
							System.out.println("setup_date_y ="+setup_date_y);
							String setup_date_m =  ( request.getParameter("setup_date_m")==null ) ? "" : (String)request.getParameter("setup_date_m");						   
							System.out.println("setup_date_m ="+setup_date_m);
							String setup_date_d =  ( request.getParameter("setup_date_d")==null ) ? "" : (String)request.getParameter("setup_date_d");						   
							System.out.println("setup_date_d ="+setup_date_d);
							String machine_name =  ( request.getParameter("machine_name")==null ) ? "" : (String)request.getParameter("machine_name");						   
							System.out.println("ssssssssssssssss machine_name ="+machine_name);
							seq_no =  ( request.getParameter("seq_no")==null ) ? "" : (String)request.getParameter("seq_no");						   
							System.out.println("seq_no ="+seq_no);
							
							List dbData_Check_Property_No = Check_Modify_Property_No(property_no,bank_no,cancel_type,seq_no);
							String check_var = ((DataObject)dbData_Check_Property_No.get(0)).getValue("check_var1").toString();
							System.out.println("check_var = "+check_var); 
							
							if(!check_var.equals("0")){
								alertMsg = "該機器編號己裝設，尚未遷移/裁撤，不可重覆建檔!!!";
								webURL_Y = "/pages/FX005AW_Edit.jsp?act=Edit&property_no="+property_no+"&site_name="+site_name+"&addr="+addr+"&setup_date_y="+setup_date_y+"&setup_date_m="+setup_date_m+"&setup_date_d="+setup_date_d+"&machine_name="+machine_name+"&seq_no="+seq_no;
								rd = application.getRequestDispatcher( nextPgName );
								   
							}else{
						  		actMsg = UpdateDB(request,bank_no,lguser_id,lguser_name);
        	    				rd = application.getRequestDispatcher( nextPgName+"?FX=FX005AW" );
						    }	
						  
						  
    	}
    	
    	
    	//2005/11/16 新增載入視窗顯示 by 4180
    	else if(act.equals("Load")){
    		if(checkseq.equals("")){
    		  List dbData = getWLX05_ATM_SETUP(bank_no,"");
    		  request.setAttribute("WLX05_ATM_SETUP",dbData);
        		rd = application.getRequestDispatcher( LoadPgName +"?bank_no="+bank_no); 
        	}else{
        		List dbData = getWLX05_ATM_SETUP(bank_no,checkseq);
    	    	request.setAttribute("WLX05_ATM_SETUP",dbData);
        		request.setAttribute("maintainInfo","select * from WLX05_ATM_SETUP WHERE bank_no='" + bank_no+"'");
        	rd = application.getRequestDispatcher( EditPgName +"?act=new&loaddata=ok");
        		
        	}
        }
    	request.setAttribute("actMsg",actMsg);
    	request.setAttribute("alertMsg",alertMsg);
    	request.setAttribute("webURL_Y",webURL_Y);
    }

	try {
        	//forward to next present jsp
        	rd.forward(request, response);
    	} catch (NullPointerException npe) { }
    }//end of doProcess
%>


<%!
    private final static String nextPgName = "/pages/ActMsg.jsp";
    private final static String EditPgName = "/pages/FX005AW_Edit.jsp";
    private final static String ListPgName = "/pages/FX005AW_List.jsp";
    private final static String LoadPgName = "/pages/FX005AW_Load.jsp";
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    
    //94.04.01 add 主key更改為bank_no+seq_no by 2295
    //99.12.08 fix 縣市合併問題 by 2808
    private List getWLX05_ATM_SETUP(String bank_no,String seq_no){
			String cd01Table = Integer.parseInt(Utility.getYear()) > 99 ? "cd01" :"cd01_99" ;
			
    		//程序為顯示畫面，查詢條件
  
		    List paramList = new ArrayList() ;
			String sqlCmd = " SELECT tem_file.*, NVL(CD01.HSIEN_NAME, ' ')  "+
			                " AS HSIEN_NAME ,  NVL(CD02.AREA_NAME, ' ') AS AREA_NAME "+
                            " from "+
                            " (select * from WLX05_ATM_SETUP where bank_no= ? ";
            paramList.add(bank_no) ;
            if(!"".equals(seq_no)) {
            	sqlCmd += " and seq_no= ?)tem_file " ;
            	paramList.add(seq_no) ;	
            }else {
            	sqlCmd += ") tem_file ";
            }
            //       sqlCmd =(seq_no.equals(""))?sqlCmd+") tem_file":sqlCmd+" and seq_no="+seq_no +") tem_file";                   
                   sqlCmd+=" LEFT JOIN "+cd01Table+" cd01 ON tem_file.HSIEN_ID  = cd01.HSIEN_ID"+
					       " LEFT JOIN CD02 ON tem_file.AREA_ID = cd02.AREA_ID"+
						   " ORDER BY SETUP_DATE DESC";
		
           List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"seq_no,setup_date,cancel_date,update_date,rank");
           return dbData;
    
    }
    
    
    private List getWLX05_ATM_SETUP_SORT2(String bank_no,String seq_no,String sort1,String sort2){

    		String cd01Table = Integer.parseInt(Utility.getYear()) > 99 ? "cd01" :"cd01_99" ;
    		//程序為顯示畫面，查詢條件
  
			List paramList = new ArrayList() ;
			String sqlCmd = " SELECT tem_file.*, NVL(CD01.HSIEN_NAME, ' ')  "+
			                " AS HSIEN_NAME ,  NVL(CD02.AREA_NAME, ' ') AS AREA_NAME "+
                            " from "+
                            " (select * from WLX05_ATM_SETUP where bank_no= ? ";
            paramList.add(bank_no) ;
            if("".equals(seq_no)) {
            	sqlCmd += " ) tem_file " ;
            }else {
            	sqlCmd += " and seq_no = ? ) tem_file " ;
            	paramList.add(seq_no) ;
            }
             //      sqlCmd =(seq_no.equals(""))?sqlCmd+") tem_file":sqlCmd+" and seq_no="+seq_no +") tem_file";                   
                   sqlCmd+=" LEFT JOIN "+cd01Table+" cd01 ON tem_file.HSIEN_ID  = cd01.HSIEN_ID"+
					       " LEFT JOIN CD02 ON tem_file.AREA_ID = cd02.AREA_ID"+
						   " ORDER BY "+sort1+","+sort2+" DESC";
		   //System.out.println("getWLX05_ATM_SETUP_SORT2 sqlCmd:"+sqlCmd);
           List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"SEQ_NO,SITE_NAME,PROPERTY_NO,MACHINE_NAME,HSIEN_NAME,AREA_NAME,ADDR,SETUP_DATE,CANCEL_TYPE,CANCEL_DATE,COMMENT_M,USER_ID,USER_NAME,UPDATE_DATE");
           return dbData;
    
    }
    
    
    private List getWLX05_ATM_SETUP_SORT3(String bank_no,String seq_no,String sort1,String sort2,String sort3){
    	//程序為顯示畫面，查詢條件
    	String cd01Table = Integer.parseInt(Utility.getYear()) > 99 ? "cd01" :"cd01_99" ;
    	List paramList = new ArrayList() ;
			String sqlCmd = " SELECT tem_file.*, NVL(CD01.HSIEN_NAME, ' ')  "+
			                " AS HSIEN_NAME ,  NVL(CD02.AREA_NAME, ' ') AS AREA_NAME "+
                            " from "+
                            " (select * from WLX05_ATM_SETUP where bank_no= ? ";
            paramList.add(bank_no) ;
            if(!"".equals(seq_no)) {
            	sqlCmd += " and seq_no=? ) tem_file " ;
			    paramList.add(seq_no) ;
            }else {
            	sqlCmd += " ) tem_file " ;
            }
            //       sqlCmd =(seq_no.equals(""))?sqlCmd+") tem_file":sqlCmd+" and seq_no="+seq_no +") tem_file";                   
                   sqlCmd+=" LEFT JOIN "+cd01Table+" cd01 ON tem_file.HSIEN_ID  = cd01.HSIEN_ID"+
					       " LEFT JOIN CD02 ON tem_file.AREA_ID = cd02.AREA_ID"+
						   " ORDER BY "+sort1+" DESC ,"+sort2+","+sort3;
		
           List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"SEQ_NO,SITE_NAME,PROPERTY_NO,MACHINE_NAME,HSIEN_NAME,AREA_NAME,ADDR,SETUP_DATE,CANCEL_TYPE,CANCEL_DATE,COMMENT_M,USER_ID,USER_NAME,UPDATE_DATE");
           return dbData;
    
    }
    
    
    //95.04.17 add 檢查機器編碼是否重複 by 2495
    private List Check_Add_Property_No(String property_no,String bank_no,String cancel_type){
	  		List paramList =new ArrayList() ;
			String sqlCmd = " select count(*) as check_var1  from WLX05_ATM_SETUP"+
			                " where bank_no= ? and "+
			                " nvl(CANCEL_TYPE,' ') = ' ' and "+
			                " nvl(PROPERTY_NO,' ') =? ";
        
            paramList.add(bank_no) ;
            paramList.add(property_no) ;
           List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"check_var1");
           return dbData;
    }
    
    private List Check_Modify_Property_No(String property_no,String bank_no,String cancel_type,String seq_no){
    	    List paramList =new ArrayList() ;
			String sqlCmd = " select count(*) as check_var1  from WLX05_ATM_SETUP"+
			                " where bank_no= ? and "+
			                " seq_no <> ? and "+
			                " nvl(CANCEL_TYPE,' ') = ' ' and "+
			                " nvl(PROPERTY_NO,' ') = ? ";
           paramList.add(bank_no) ;
           paramList.add(seq_no);
           paramList.add(property_no);
           List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"check_var1");
           return dbData;
    }
    
    
    public String InsertDB(HttpServletRequest request,String bank_no,String lguser_id,String lguser_name) throws Exception{

	StringBuffer sqlCmd = new StringBuffer() ;
	String errMsg="";
	String site_name=((String)request.getParameter("site_name")==null)?"-":(String)request.getParameter("site_name");
	String property_no=((String)request.getParameter("property_no")==null)?"-":(String)request.getParameter("property_no");
	String machine_name=((String)request.getParameter("machine_name")==null)?"-":(String)request.getParameter("machine_name");
	String hsien_id_area_id=((String)request.getParameter("HSIEN_ID_AREA_ID")==null)?"":(String)request.getParameter("HSIEN_ID_AREA_ID");
	String hsien_id = hsien_id_area_id.substring(0,hsien_id_area_id.indexOf("/"));
	String area_id = hsien_id_area_id.substring(hsien_id_area_id.indexOf("/")+1,hsien_id_area_id.length());
	String addr=((String)request.getParameter("addr")==null)?"-":(String)request.getParameter("addr");
	String setup_date=(String)request.getParameter("SETUP_DATE");
	String cancel_type=((String)request.getParameter("cancel_type")==null)?"N":(String)request.getParameter("cancel_type");
	String cancel_date=(String)request.getParameter("CANCEL_DATE");		
	String comment_m = ((String)request.getParameter("comment_m")==null)?"N":(String)request.getParameter("comment_m");
	String user_id=lguser_id;
	String user_name=lguser_name;

		try {
   		//List updateDBSqlList = new LinkedList();
   		List max_seq_number = new LinkedList();
   		List paramList = new ArrayList() ;
   		//先去取得seq number
			sqlCmd.append("select to_char(max(SEQ_NO)+1) as maxq from WLX05_ATM_SETUP where bank_no = ?");			    
			paramList.add(bank_no) ;
			max_seq_number = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");			  
			String max_seq =(String)((DataObject)max_seq_number.get(0)).getValue("maxq");
			max_seq=(max_seq==null)?"1":max_seq;
			sqlCmd.setLength(0) ;
			paramList.clear() ;
			sqlCmd.append(" insert into WLX05_ATM_SETUP VALUES ( ");
			sqlCmd.append(" ?,?,?,?,?,?,?,?,to_date(?,'YYYY/MM/DD'),?,");
			sqlCmd.append("to_date(?,'YYYY/MM/DD'),?,?,?,sysdate ) ");
			paramList.add(bank_no) ;
			paramList.add(max_seq) ;
			paramList.add(site_name) ;
			paramList.add(property_no) ;
			paramList.add(machine_name) ;
			paramList.add(hsien_id) ;
			paramList.add(area_id) ;
			paramList.add(addr) ;
			paramList.add(setup_date) ;
			paramList.add(cancel_type) ;
			paramList.add(cancel_date) ;
			paramList.add(comment_m) ;
			paramList.add(user_id) ;
			paramList.add(user_name) ;
		  /*sqlCmd = "INSERT INTO WLX05_ATM_SETUP VALUES('"
		  +bank_no+"','"
		  +max_seq+"','"
		  +site_name+"','"
		  +property_no+"','"
		  +machine_name+"','"
		  +hsien_id+"','"
		  +area_id+"','"
		  +addr+"',"
		  +"to_date('"+setup_date+"','YYYY/MM/DD'),'"
		  +cancel_type+"',"
		  +"to_date('"+cancel_date+"','YYYY/MM/DD'),'"
		  +comment_m+"','"
		  +user_id+"','"
		  +user_name
		  +"',sysdate)";*/

           //     updateDBSqlList.add(sqlCmd);
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
	//94.04.01 add 主key更改為bank_no+seq_no by 2295
	public String UpdateDB(HttpServletRequest request,String bank_no,String lguser_id,String lguser_name) throws Exception{
		StringBuffer sqlCmd = new StringBuffer();
		String sqlCmdLog = "";
		String errMsg="";

		String seq_no = ((String)request.getParameter("editseq_no")==null)?"":(String)request.getParameter("editseq_no");
		String site_name=((String)request.getParameter("site_name")==null)?"-":(String)request.getParameter("site_name");
		String property_no=((String)request.getParameter("property_no")==null)?"-":(String)request.getParameter("property_no");
		String machine_name=((String)request.getParameter("machine_name")==null)?"-":(String)request.getParameter("machine_name");
		String hsien_id_area_id=((String)request.getParameter("HSIEN_ID_AREA_ID")==null)?"":(String)request.getParameter("HSIEN_ID_AREA_ID");
		String hsien_id = "".equals(hsien_id_area_id)?"":hsien_id_area_id.substring(0,hsien_id_area_id.indexOf("/"));
		String area_id = "".equals(hsien_id_area_id)?"":hsien_id_area_id.substring(hsien_id_area_id.indexOf("/")+1,hsien_id_area_id.length());
		String addr=((String)request.getParameter("addr")==null)?"-":(String)request.getParameter("addr");
		String setup_date=(String)request.getParameter("SETUP_DATE");
		String cancel_type=((String)request.getParameter("cancel_type")==null)?"N":(String)request.getParameter("cancel_type");
		String cancel_date=(String)request.getParameter("CANCEL_DATE");		
		String comment_m = ((String)request.getParameter("comment_m")==null)?"N":(String)request.getParameter("comment_m");
		String user_id=lguser_id;
		String user_name=lguser_name;

		try {
				//List updateDBSqlList = new LinkedList();
				List paramList = new ArrayList() ;
				 sqlCmd.append("SELECT * FROM WLX05_ATM_SETUP WHERE bank_no=? AND seq_no=? ");
				 paramList.add(bank_no) ;
				 paramList.add(seq_no) ;
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
			    paramList.clear() ;
				sqlCmd.setLength(0) ;
				if (data.size() == 0){
				    System.out.println("seq_no="+seq_no);
				    System.out.println("bank_no="+bank_no);
				    errMsg = errMsg + "此筆資料不存在無法修改<br>";
				}else{

				/*寫入log檔中*/

				sqlCmdLog = "INSERT INTO WLX05_ATM_SETUP_LOG "
		             +"select bank_no,seq_no,site_name,"
		             +"property_no,machine_name,hsien_id,area_id,"
		             +"addr,setup_date,cancel_type,cancel_date,comment_m,"
		             +"user_id,user_name,update_date,?,?,sysdate,'U' from WLX05_ATM_SETUP where bank_no=? and seq_no=?";
		        
		        paramList.add(user_id) ;
		        paramList.add(user_name) ;
		        paramList.add(bank_no);
		        paramList.add(seq_no) ;
		        //updateDBSqlList.add(sqlCmdLog);  
		        this.updDbUsesPreparedStatement(sqlCmdLog,paramList) ;
		        sqlCmd.setLength(0) ;
		        paramList.clear() ;
				    sqlCmd.append("UPDATE WLX05_ATM_SETUP SET ") ;
				    sqlCmd.append("bank_no=?");paramList.add(bank_no);
				    sqlCmd.append(",seq_no=?");paramList.add(seq_no);
				    sqlCmd.append(",site_name=?");paramList.add( site_name ) ;
				    sqlCmd.append(",property_no=?");paramList.add(property_no ) ;
				    sqlCmd.append(",machine_name=?");paramList.add(machine_name ) ;
				    sqlCmd.append(",hsien_id=?");paramList.add(hsien_id ) ;
				    sqlCmd.append(",area_id=?");paramList.add(area_id );
				    sqlCmd.append(",addr=?");paramList.add(addr ) ;
				    sqlCmd.append(",setup_date=to_date(?,'YYYY/MM/DD')");paramList.add(setup_date);
				    sqlCmd.append(",cancel_type=?");paramList.add(cancel_type) ;
				    sqlCmd.append(",cancel_date=to_date(?,'YYYY/MM/DD')");paramList.add(cancel_date) ;
				    sqlCmd.append(",comment_m=?");paramList.add(comment_m);
				    sqlCmd.append(",user_id=?");paramList.add(user_id);
   					sqlCmd.append(",user_name=?");paramList.add(user_name);
   					sqlCmd.append(",update_date=sysdate");
   			        sqlCmd.append(" where bank_no=? and seq_no=? ");
   			        paramList.add(bank_no) ;
   			        paramList.add(seq_no) ;
					
		            //updateDBSqlList.add(sqlCmd);
					if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)) {
					//if(DBManager.updateDB(updateDBSqlList)){
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
    //94.04.01 add 主key更改為bank_no+seq_no by 2295
    public String DeleteDB(HttpServletRequest request,String bank_no,String lguser_id,String lguser_name) throws Exception{
		StringBuffer sqlCmd = new StringBuffer();
		String sqlCmdLog = "";
		String errMsg="";
		String seq_no = ((String)request.getParameter("editseq_no")==null)?"":(String)request.getParameter("editseq_no");
		String user_id=lguser_id;
	  	String user_name=lguser_name;


		try {
			   //List updateDBSqlList = new LinkedList();
			   List paramList =new ArrayList() ;
			   sqlCmd.append("SELECT * FROM WLX05_ATM_SETUP WHERE bank_no=? AND seq_no=? ");
			   paramList.add(bank_no) ;
			   paramList.add(seq_no) ;
			   List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
			

				if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法刪除<br>";
				}else{
					sqlCmdLog = "INSERT INTO WLX05_ATM_SETUP_LOG "
		             +"select bank_no,seq_no,site_name,"
		             +"property_no,machine_name,hsien_id,area_id,"
		             +"addr,setup_date,cancel_type,cancel_date,comment_m,"
		             +"user_id,user_name,update_date,? ,?,sysdate,'D' from WLX05_ATM_SETUP where bank_no=? and seq_no=?";
		    		 paramList.clear() ;
		    		 paramList.add(user_id) ;
		    		 paramList.add(user_name) ;
		    		 paramList.add(bank_no) ;
		    		 paramList.add(seq_no);
		        	//updateDBSqlList.add(sqlCmdLog);  
 					 
 					if(!this.updDbUsesPreparedStatement(sqlCmdLog,paramList)){					
				   		errMsg = errMsg + "WLX05_ATM_SETUP_LOG相關資料刪除失敗";
					}
 					paramList.clear() ;
 					sqlCmd.setLength(0) ;
				     sqlCmd.append(" delete WLX05_ATM_SETUP where bank_no=?  and seq_no=? ");
				     paramList.add(bank_no) ;
				     paramList.add(seq_no) ;
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
