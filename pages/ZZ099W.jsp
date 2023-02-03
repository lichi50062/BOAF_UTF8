<%
// 99.12.12 fix sqlInjection by 2808
//100.01.06 fix 密碼解密作業 by 2295
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
<%@ page import="java.util.Enumeration" %>
<%@ page import="java.util.HashMap" %>
<%@ page import="java.util.Hashtable" %>
<%@ page import="java.util.StringTokenizer" %>

<%
	RequestDispatcher rd=null;
	String actMsg="";
	String alertMsg="";
	String webURL="";
	boolean doProcess=false;

	//取得session資料,取得成功時,才繼續往下執行===================================================
	if(session.getAttribute("muser_id")==null){//session timeout
      System.out.println("ZZ099W login timeout");
	   rd=application.getRequestDispatcher("/pages/reLogin.jsp?url=LoginError.jsp?timeout=true");
	   try{
          rd.forward(request,response);
       }catch(Exception e){
          System.out.println("forward Error:"+e+e.getMessage());
       }
    }else{
      doProcess=true;
    }
	//若muser_id有資料時,表示登入成功
	if (doProcess) {
		String act=(request.getParameter("act")==null) ? "" : (String)request.getParameter("act");

		System.out.println("act="+act);

		//登入者資訊
		String lguser_id=(session.getAttribute("muser_id")==null) ? "" : (String)session.getAttribute("muser_id");
		String lguser_name=(session.getAttribute("muser_name")==null) ? "" : (String)session.getAttribute("muser_name");
		String lguser_type=(session.getAttribute("muser_type")==null) ? "" : (String)session.getAttribute("muser_type");
		String lguser_tbank_no=(session.getAttribute("tbank_no")==null) ? "" : (String)session.getAttribute("tbank_no");
		session.setAttribute("nowtbank_no",null);//94.01.05 fix 沒有Bank_List,把所點選的Bank_no清除======

		String bank_type=(session.getAttribute("bank_type")==null) ? "" : (String)session.getAttribute("bank_type");
	    String upd_code=(request.getParameter("UPD_CODE")==null) ? "" : (String)request.getParameter("UPD_CODE");

		String gp_orgcate=(request.getParameter("orgcate")==null) ? "" : (String)request.getParameter("orgcate");
		String gp_tbankno=(request.getParameter("TBANK_NO")==null) ? "" : (String)request.getParameter("TBANK_NO");
		String gp_bankno=(request.getParameter("BANK_NO")==null) ? "" : (String)request.getParameter("BANK_NO");
		String gp_utPW=(request.getParameter("user_pw")==null) ? "" : (String)request.getParameter("user_pw");

		//無權限時,導向到LoginError.jsp
    	if (!CheckPermission(request)) {
			rd=application.getRequestDispatcher(LoginErrorPgName);
		}
			else {
			//set next jsp
	    	if (act.equals("List")) {
	        	rd=application.getRequestDispatcher(ListPgName+"?act=List");
			}
				else if (act.equals("Qry") || act.equals("Delete")) {

		    	    String getPW="";
		    	    List getpwList=getPW(lguser_id);
		    	    //100.01.06 fix 密碼解密作業 by 2295
		    	    getPW=Utility.decode((String)((DataObject)getpwList.get(0)).getValue("muser_password"));					
		    	    if (gp_utPW!=null && gp_utPW.equals(getPW)) {

						if (act.equals("Qry")) {
							List bn01List=getBN01(gp_tbankno);
							request.setAttribute("bn01List",bn01List);
							List bn02List=getBN02(gp_tbankno,gp_bankno);
							request.setAttribute("bn02List",bn02List);
							List ba01List=getBA01(gp_tbankno,gp_bankno);
							request.setAttribute("ba01List",ba01List);
			    	    	List orgList=getQryResult(gp_orgcate,gp_tbankno,gp_bankno);
			    	    	request.setAttribute("orgList",orgList);
			    	    	rd=application.getRequestDispatcher(ListPgName+"?act=Qry&org_cate="+gp_orgcate+"&tbank_no="+gp_tbankno+"&bank_no="+gp_bankno+"&upd_code="+upd_code);
		    	    	}
							else if (act.equals("Delete")) {
								actMsg=DeleteDB(request,lguser_id,lguser_name,gp_orgcate);
								alertMsg="返回「總分支機構組織強制維護」查詢主頁？";
								webURL=(String)session.getAttribute("aftDel");
								webURL=ListPgName+"?act=List";
								rd=application.getRequestDispatcher(nextPgName);
							}
		    	    }
		    	    	else {
		    	    		actMsg="輸入密碼與目前使用者密碼未符合，請點選「回上一頁」並重新輸入！";
		    	    		rd=application.getRequestDispatcher(nextPgName);
		    	    	}
	    		}
	    	request.setAttribute("actMsg",actMsg);
			request.setAttribute("alertMsg",alertMsg);
	      	request.setAttribute("webURL_Y",webURL);

	      	}//end of else

		//forward to next present jsp
		try {
	        rd.forward(request, response);
	    }
	    catch (NullPointerException npe) {
	    }
    }//end of doProcess
%>


<%!
    private final static String nextPgName="/pages/ActMsg.jsp";
    private final static String ListPgName="/pages/ZZ099W_List.jsp";
    private final static String LoginErrorPgName="/pages/LoginError.jsp";


//	//檢核權限
    private boolean CheckPermission(HttpServletRequest request){
		boolean CheckOK=false;
		HttpSession session=request.getSession();
		Properties permission=(session.getAttribute("ZZ099W")==null) ? new Properties() : (Properties)session.getAttribute("ZZ099W");
		if (permission==null) {
			System.out.println("ZZ099W.permission==null");
		}
			else {
				System.out.println("ZZ099W.permission.size="+permission.size());
			}
		//只要有Query的權限,就可以進入畫面
		if (permission!=null && permission.get("Q")!=null && permission.get("Q").equals("Y")) {
			CheckOK=true;//Query
		}
	        	return CheckOK;
    }



    //取BA01
    private List getBA01(String p_tbankno,String p_bankno) {
		String sqlCmd="";
		List paramList =new ArrayList() ;
		if (p_tbankno!=null && p_bankno.length()==0) {
			sqlCmd=" select bank_kind,bank_no,bank_name,pbank_no from BA01 where bank_no=? ";
			paramList.add(p_tbankno) ;
		}
			else if (p_tbankno!=null && p_bankno.length()>0) {
				sqlCmd=" select bank_kind,bank_no,bank_name,pbank_no from BA01 where bank_no=? ";
				paramList.add(p_bankno) ;
			}

		List dbData=DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");
		return dbData;
    }



    //取BN01
    private List getBN01(String p_tbankno) {
		String sqlCmd="";
		List paramList =new ArrayList() ;
		sqlCmd=" select bank_no as pbank_no,bank_name as pbank_name from BN01 where bank_type in ('6','7') and bank_no=? ";
		paramList.add(p_tbankno) ;
		List dbData=DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");
		return dbData;
    }



    //取BN02
    private List getBN02(String p_tbankno,String p_bankno) {
		String sqlCmd="";
		List paramList = new ArrayList() ;
		sqlCmd=" select tbank_no as pbank_no,bank_no,bank_name as bank_name from BN02 where bank_type in ('6','7') and bank_no=? ";
		paramList.add(p_bankno) ;
		List dbData=DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");
		return dbData;
    }



	//使用者密碼
	private List getPW(String p_muid) {
		List paramList =new ArrayList() ;
		String sqlCmd=" select muser_password from wtt01 where muser_id=? ";
		paramList.add(p_muid) ;
		List dbData=DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");
		return dbData;
	}



    //查詢
	private List getQryResult(String p_orgcate,String p_tbankno,String p_bankno) {
		String sqlCmd=new String("");
		List paramList =new ArrayList() ;
		if (p_orgcate.equals("0")) {
			sqlCmd=" select '(1)BN01'||bank_name as bankdesc,'BN01' as itemcode,'總機構配賦資料' as itemname,count(*) as itemcnt from bn01 where bank_no=? and bank_type in ('6','7') group by '(1)BN01'||bank_name union "+
			" select '(2)BA01'||bank_name,'BA01','總分支機構清單（總）',count(*) from ba01 where pbank_no=? and bank_type in ('6','7') and bank_kind='0' group by '(2)BA01'||bank_name union "+
			" select bank_no,'WLX01','總機構基本資料維護',count(*) from wlx01 where bank_no=? group by bank_no union "+
			" select bank_no,'WLX01_M','總機構高階主管基本資料',count(*) from wlx01_m where bank_no=? group by bank_no union "+
			" select bank_no,'WLX04','理監事基本資料維護',count(*) from wlx04 where bank_no=? group by bank_no ";
            paramList.add(p_tbankno) ;
            paramList.add(p_tbankno) ;
            paramList.add(p_tbankno) ;
            paramList.add(p_tbankno);
            paramList.add(p_tbankno) ;
		}
		if (p_orgcate.equals("1")) {
			sqlCmd=" select '(1)BN01'||bank_name as bankdesc,'BN01' as itemcode,'總機構配賦資料' as itemname,count(*) as itemcnt from bn01 where bank_no=? and bank_type in ('6','7') group by '(1)BN01'||bank_name union "+
			" select '(1)BN02'||bank_name,'BN02','分支機構配賦資料',count(*) from bn02 where tbank_no=? and bank_no=? and bank_type in ('6','7') group by '(1)BN02'||bank_name union "+
			" select '(2)BA01'||bank_name,'BA01','總分支機構清單（總）',count(*) from ba01 where bank_no=? and bank_type in ('6','7') and bank_kind='0' group by '(2)BA01'||bank_name union "+
			" select '(2)BA01'||bank_name,'BA01','總分支機構清單（分）',count(*) from ba01 where bank_no=? and bank_type in ('6','7') and bank_kind='1' group by '(2)BA01'||bank_name union "+
			" select bank_no,'WLX02','國內營業分支機構基本資料維護',count(*) from wlx02 where tbank_no=? and bank_no=? group by bank_no union "+
			" select bank_no,'WLX02_M','國內營業分支機構負責人基本資料維護',count(*) from wlx02 where bank_no=? group by bank_no";
			paramList.add(p_bankno) ;
			paramList.add(p_bankno) ;
			paramList.add(p_bankno) ;
			paramList.add(p_bankno) ;
			paramList.add(p_bankno) ;
			paramList.add(p_bankno) ;
			paramList.add(p_bankno) ;
			paramList.add(p_bankno) ;
		}
		System.out.println("p_orgcate="+p_orgcate);
		System.out.println("p_tbankno="+p_tbankno);
		System.out.println("p_bankno="+p_bankno);
		List dbData=DBManager.QueryDB_SQLParam(sqlCmd,paramList,"itemcnt");
		return dbData;
    }


	//刪除
	public String DeleteDB(HttpServletRequest request,String lguser_id,String lguser_name,String p_orgcate) throws Exception {
		String sqlCmd="";
		String errMsg="";
		String user_id=lguser_id;
	    String user_name=lguser_name;
	    String delOrgcate=p_orgcate;
		String delIC="";
		String delTBankno=(request.getParameter("TBANK_NO")==null ) ? "" : (String)request.getParameter("TBANK_NO");
		String delBankno=(request.getParameter("BANK_NO")==null ) ? "" : (String)request.getParameter("BANK_NO");
		List paramList = new ArrayList() ;
		try {
			//取出form裡的所有變數
		  	Enumeration ep=request.getParameterNames();
		  	Enumeration ea=request.getAttributeNames();
		  	Hashtable t=new Hashtable();
		  	String name="";

		  	for (;ep.hasMoreElements();) {
				name=(String)ep.nextElement();
				t.put(name,request.getParameter(name));
		  	}
		  	int row=Integer.parseInt((String)t.get("row"));
		  	System.out.println("row="+row);
		  	List deleteData=new LinkedList();
		  	for (int i=0;i<row;i++) {
				if (t.get("isModify_"+(i+1))!=null) {
					deleteData.add((String)t.get("isModify_"+(i+1)));
				}
		  	}
		  	System.out.println("deleteData.size="+deleteData.size());

			List updateDBSqlList=new LinkedList();
			List data=null;

			for (int i=0;i<deleteData.size();i++) {
				delIC=(String)deleteData.get(i);
				if (delIC.equals("BN01") && delOrgcate.equals("0")) {
					paramList.clear() ;
					sqlCmd=" insert into bn04_log (tbank_no,bank_no,bank_name,bn_type,bank_type,add_user,add_name,add_date,bank_b_name, "+
						   " kind_1,kind_2,bn_type2,exchange_no,user_id,user_name,update_date,user_id_c,user_name_c,update_date_c,update_kind_c, "+
						   " update_type_c) select bank_no,bank_no,bank_name,bn_type,bank_type,add_user,add_name,add_date,bank_b_name, "+
						   " kind_1,kind_2,bn_type2,exchange_no,user_id,user_name,update_date,?,?,sysdate,'0','D' "+
						   " from bn01 where bank_no=? and bank_type in ('6','7') ";
					paramList.add(user_id) ;
					paramList.add(user_name) ;
					paramList.add(delTBankno) ;
					this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
					paramList.clear() ;
					//System.out.println("write log: "+sqlCmd);
					//updateDBSqlList.add(sqlCmd);
					sqlCmd=" delete from bn01 where bank_no=? and bank_type in ('6','7') ";
					paramList.add(delTBankno) ;
					//System.out.println("do job: "+sqlCmd);
					//updateDBSqlList.add(sqlCmd);
					this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
				}
				if (delIC.equals("BN02")) {
					if (delOrgcate.equals("0")) {
						paramList.clear() ;
						sqlCmd=" insert into bn04_log (tbank_no,bank_no,bank_name,bn_type,bank_type,add_user,add_name,add_date,bank_b_name, "+
							   " kind_1,kind_2,bn_type2,exchange_no,user_id,user_name,update_date,user_id_c,user_name_c,update_date_c,update_kind_c, "+
							   " update_type_c) select bank_no,bank_no,bank_name,bn_type,bank_type,add_user,add_name,add_date,bank_b_name, "+
							   " kind_1,kind_2,bn_type2,exchange_no,user_id,user_name,update_date,?,?,sysdate,'0','D' "+
							   " from bn02 where tbank_no=? and bank_type in ('6','7') ";
						paramList.add(user_id) ;
						paramList.add(user_name);
						paramList.add(delTBankno);
						this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
						paramList.clear() ;
						//System.out.println("write log:"+sqlCmd);
						//updateDBSqlList.add(sqlCmd);
						sqlCmd=" delete from bn02 where tbank_no=?  and bank_type in ('6','7') ";
						paramList.add(delTBankno) ;
						//System.out.println("do job:"+sqlCmd);
						//updateDBSqlList.add(sqlCmd);
						this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
					}
					if (delOrgcate.equals("1")) {
						sqlCmd=" insert into bn04_log (tbank_no,bank_no,bank_name,bn_type,bank_type,add_user,add_name,add_date,bank_b_name, "+
							   " kind_1,kind_2,bn_type2,exchange_no,user_id,user_name,update_date,user_id_c,user_name_c,update_date_c,update_kind_c, "+
							   " update_type_c) select tbank_no,bank_no,bank_name,bn_type,bank_type,add_user,add_name,add_date,bank_b_name, "+
							   " kind_1,kind_2,bn_type2,exchange_no,user_id,user_name,update_date,?,?,sysdate,'1','D' "+
							   " from bn02 where tbank_no=? and bank_no=? and bank_type in ('6','7') ";
						paramList.add(user_id) ;
						paramList.add(user_name) ;
						paramList.add(delTBankno);
						paramList.add(delBankno);
						//System.out.println("write log:"+sqlCmd);
						//updateDBSqlList.add(sqlCmd);
						this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
						paramList.clear() ;
						sqlCmd=" delete from bn02 where tbank_no=? and bank_no=? and bank_type in ('6','7') ";
						paramList.add(delTBankno) ;
						paramList.add(delBankno) ;
						//System.out.println("do job:"+sqlCmd);
						//updateDBSqlList.add(sqlCmd);
						this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
					}
				}
				paramList.clear();
				if (delIC.equals("BA01")) {
					if (delOrgcate.equals("0")) {
						sqlCmd=" delete from ba01 where pbank_no=?  and bank_type in ('6','7') and bank_kind='0' ";
						paramList.add(delTBankno) ;
					}
					if (delOrgcate.equals("1")){
						sqlCmd=" delete from ba01 where pbank_no=? and bank_no=? and bank_type in ('6','7') and bank_kind='1' ";
						paramList.add(delTBankno) ;
						paramList.add(delBankno) ;
					}
					//System.out.println("SQL="+sqlCmd);
					//updateDBSqlList.add(sqlCmd);
					this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
				}
				if (delIC.equals("WLX01") && delOrgcate.equals("0")) {
					paramList.clear();
					sqlCmd=" insert into wlx01_log select * from wlx01 where bank_no='"+delTBankno+"' ";
					//System.out.println("write log:"+sqlCmd);
					//updateDBSqlList.add(sqlCmd);
					paramList.add(delTBankno) ;
					this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
					paramList.clear();
					sqlCmd=" delete from wlx01 where bank_no=?  ";
					paramList.add(delTBankno) ;
					//System.out.println("do job:"+sqlCmd);
					//updateDBSqlList.add(sqlCmd);
					this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
				}
				if (delIC.equals("WLX01_M") && delOrgcate.equals("0")) {
					paramList.clear();
					sqlCmd=" insert into wlx01_m_log select * from wlx01_m where bank_no=? ";
					paramList.add(delTBankno) ;
					this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
					//System.out.println("write log:"+sqlCmd);
					//updateDBSqlList.add(sqlCmd);
					paramList.clear() ;
					sqlCmd=" delete from wlx01_M where bank_no='"+delTBankno+"' ";
					//System.out.println("do job:"+sqlCmd);
					//updateDBSqlList.add(sqlCmd);
					paramList.add(delTBankno) ;
					this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
				}
				if (delIC.equals("WLX04") && delOrgcate.equals("0")) {
					paramList.clear() ;
					sqlCmd=" insert into wlx04_log select * from wlx04 where bank_no=? ";
					paramList.add(delTBankno) ;
					//System.out.println("write log:"+sqlCmd);
					//updateDBSqlList.add(sqlCmd);
					this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
					paramList.clear() ;
					paramList.add(delTBankno) ;
					sqlCmd=" delete from wlx04 where bank_no=? ";
					//System.out.println("do job:"+sqlCmd);
					//updateDBSqlList.add(sqlCmd);
					this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
				}
				if (delIC.equals("WLX02")) {
					
					if (delOrgcate.equals("0")) {
						paramList.clear() ;
						sqlCmd=" insert into wlx02_log select * from wlx02 where tbank_no=? ";
						paramList.add(delTBankno) ;
						//System.out.println("write log:"+sqlCmd);
						//updateDBSqlList.add(sqlCmd);
						this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
						paramList.clear() ;
						paramList.add(delTBankno) ;
						sqlCmd=" delete from wlx02 where tbank_no=? ";
						//System.out.println("do job:"+sqlCmd);
						//updateDBSqlList.add(sqlCmd);
						this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
					}
					if (delOrgcate.equals("1")) {
						paramList.clear() ;
						sqlCmd=" insert into wlx02_log select * from wlx02 where tbank_no=? and bank_no=? ";
						paramList.add(delTBankno) ;
						paramList.add(delBankno);
						//System.out.println("write log:"+sqlCmd);
						//updateDBSqlList.add(sqlCmd);
						this.updDbUsesPreparedStatement(sqlCmd,paramList ) ;
						paramList.clear() ;
						sqlCmd=" delete from wlx02 where tbank_no=? and bank_no=? ";
						paramList.add(delTBankno) ;
						paramList.add(delBankno);
						//System.out.println("do job:"+sqlCmd);
						//updateDBSqlList.add(sqlCmd);
						this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
					}
				}
				if (delIC.equals("WLX02_M") && delOrgcate.equals("1")) {
					paramList.clear() ;
					sqlCmd=" insert into wlx02_m_log select * from wlx02_m where bank_no=? ";
					paramList.add(delBankno) ;
					//System.out.println("write log:"+sqlCmd);
					//updateDBSqlList.add(sqlCmd);
					this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
					paramList.clear() ;
					sqlCmd=" delete from wlx02_m where bank_no=?  ";
					paramList.add(delBankno) ;
					//System.out.println("do job:"+sqlCmd);
					//updateDBSqlList.add(sqlCmd);
					this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
				}
			}
		if (DBManager.updateDB(updateDBSqlList)) {
			errMsg=errMsg+"相關資料於資料庫刪除成功！";
		}else {
				errMsg=errMsg+"相關資料於資料庫刪除失敗！";
		}
		}
		catch (Exception e) {
			System.out.println(e+":"+e.getMessage());
			errMsg=errMsg+"相關資料於資料庫刪除失敗！";
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