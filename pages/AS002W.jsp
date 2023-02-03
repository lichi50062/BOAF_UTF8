<%
// 93.12.26 create by 2295
// 94.01.05 fix 沒有Bank_List,把所點選的Bank_no清除 by 2295
// 94.01.07 fix ba01增加欄位 by 2295
// 99.04.21 fix 刪除分支機構代號時,一併刪除分支機構基本資料 by 2295
// 99.10.14 fix 當有分支機構基本資料時,才需刪除分支機構基本資料 by 2295
//          fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
// 99.10.15 fix 區分100年/99年分支機構基本資料 by 2295
//101.06.29 add 若為全部類別時,可不下總機構,只下分支機構代號,查詢 by 2295
// 				  fix 修改時,若總機構代號有資料時,只顯示該筆總機構資料 by 2295
//101.08.16 add 若不為全部類別時,可不下總機構,只下分支機構代號,查詢 by 2295
//104.07.30 add 調整.只下單一家農會名稱時,查詢結果會顯示全部資料 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>

<%@include file="./include/Header.include" %>

<%

	String nowact = Utility.getTrimString(dataMap.get("nowact"));
	String bank_type = Utility.getTrimString(dataMap.get("BANK_TYPE"));
	if(bank_type.equals(""))
	    bank_type = Utility.getTrimString(dataMap.get("bank_type"));

	String tbank_no = Utility.getTrimString(dataMap.get("TBANK_NO"));
	String bank_no = Utility.getTrimString(dataMap.get("BANK_NO"));
	System.out.println("AS002W.bank_type="+bank_type);
	System.out.println("AS002W.tbank_no="+tbank_no);
	System.out.println("AS002W.bank_no="+bank_no);
	
	session.setAttribute("nowtbank_no",null);//94.01.05 fix 沒有Bank_List,把所點選的Bank_no清除======

    if(!Utility.CheckPermission(request,report_no)){//無權限時,導向到LoginError.jsp
        rd = application.getRequestDispatcher( LoginErrorPgName );
    }else{
    	//set next jsp
    	if(act.equals("new") || act.equals("getData") || act.equals("List") || act.equals("Qry")){
    	    if(act.equals("new")){
    	      bank_type="1";
    	    }
    	    List tbank_no_list = getTbank_No(bank_type,"");
    	    if(tbank_no.equals("") && (tbank_no_list.size() != 0)){
    	    		if(act.equals("new") || act.equals("getData")){//101.06.29 add 當為新增,或getdata時,總機構代號才預設取得第一筆資料顯示
               tbank_no= (String)((DataObject)tbank_no_list.get(0)).getValue("bank_no");
               System.out.println("new/getDate.bank_no="+tbank_no);
              } 
            }
    	    request.setAttribute("tbank_no",tbank_no_list);
    	    if(act.equals("new")){
        	   rd = application.getRequestDispatcher( EditPgName +"?act=new");
        	}
        	if(act.equals("List")){
        	   rd = application.getRequestDispatcher( ListPgName +"?act=List");
        	}
        	if(act.equals("Qry")){
        	   List BN02List = getQryResult(bank_type,tbank_no,bank_no);
    	       request.setAttribute("BN02List",BN02List);
    	       rd = application.getRequestDispatcher( ListPgName +"?act=Qry&bank_type="+bank_type+"&tbank_no="+tbank_no+"&bank_no="+bank_no);
    	    }
        	if(act.equals("getData")){
        	    if(nowact.equals("new") || nowact.equals("Edit")){
        	       rd = application.getRequestDispatcher( EditPgName +"?act="+nowact+"&bank_type="+bank_type+"&tbank_no="+tbank_no);
        	    }
        	    if(nowact.equals("List") || nowact.equals("Qry")){
        	       rd = application.getRequestDispatcher( ListPgName +"?act="+nowact+"&bank_type="+bank_type+"&tbank_no="+tbank_no);
        	    }
        	}
        }else if(act.equals("Edit")){
        	List tbank_no_list = getTbank_No(bank_type,tbank_no);//101.06.29 若總機構代號有資料時,只顯示該筆總機構資料
    	    if(tbank_no.equals("") && (tbank_no_list.size() != 0)){
               tbank_no= (String)((DataObject)tbank_no_list.get(0)).getValue("bank_no");
            }
    	    request.setAttribute("tbank_no",tbank_no_list);
            List BN02 = getBN02(tbank_no,bank_no);
            request.setAttribute("BN02",BN02);
            if(BN02.size() != 0){
               bank_type = (String)((DataObject)BN02.get(0)).getValue("bank_type");
               tbank_no = (String)((DataObject)BN02.get(0)).getValue("tbank_no");
            }
            //設定異動者資訊======================================================================
            Calendar now = Calendar.getInstance();
            String YEAR  = String.valueOf(now.get(Calendar.YEAR)-1911); //回覆值為西元年故需-1911取得民國年;
	    	//99.10.15 add 查詢年度100年以前.縣市別不同===============================         
  	    	String wlx01_m_year = (Integer.parseInt(YEAR) < 100)?"99":"100"; 
			request.setAttribute("maintainInfo","select * from BN02 WHERE m_year='"+wlx01_m_year+"' and tbank_no='" + tbank_no+"' and bank_no='"+bank_no+"'");
			//=======================================================================================================================
    	 	rd = application.getRequestDispatcher( EditPgName +"?act=Edit&bank_type="+bank_type+"&tbank_no="+tbank_no+"&bank_no='"+bank_no+"'");
    	}else if(act.equals("Insert")){
    	    actMsg = InsertDB(request,lguser_id,lguser_name);
        	rd = application.getRequestDispatcher( nextPgName );
        }else if(act.equals("Update")){
    	    actMsg = UpdateDB(request,lguser_id,lguser_name);
        	rd = application.getRequestDispatcher( nextPgName );
        }else if(act.equals("Delete")){
    	    actMsg = DeleteDB(request,lguser_id,lguser_name);
        	rd = application.getRequestDispatcher( nextPgName );
        }
    	request.setAttribute("actMsg",actMsg);
    }

%>

<%@include file="./include/Tail.include" %>

<%!
    private final static String report_no = "AS002W";
    private final static String nextPgName = "/pages/ActMsg.jsp";
    private final static String EditPgName = "/pages/"+report_no+"_Edit.jsp";
    private final static String ListPgName = "/pages/"+report_no+"_List.jsp";
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";

    //取得總機構代碼
    private List getTbank_No(String bank_type,String tbank_no){
    		//查詢條件
    		List paramList = new ArrayList();
    		StringBuffer sqlCmd = new StringBuffer();
    		Calendar now = Calendar.getInstance();
         	String YEAR  = String.valueOf(now.get(Calendar.YEAR)-1911); //回覆值為西元年故需-1911取得民國年;
    		//99.10.15 add 查詢年度100年以前.縣市別不同===============================	     
	      	String wlx01_m_year = (Integer.parseInt(YEAR) < 100)?"99":"100"; 
	        //===================================================================== 	      
    		sqlCmd.append(" select bank_type, bank_no,bank_name from bn01 where m_year=?");
    		if(!tbank_no.equals("")){
            sqlCmd.append(" and bank_no=?");     		
    		}
    		sqlCmd.append(" order by bank_type,bank_no ");
    		paramList.add(wlx01_m_year);
    		if(!tbank_no.equals("")) paramList.add(tbank_no);
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
        return dbData;
    }

    //取得BN02分支機構代號資料
    private List getBN02(String tbank_no,String bank_no){
    		//查詢條件
    		List paramList = new ArrayList();
    		StringBuffer sqlCmd = new StringBuffer();
    		sqlCmd.append(" select * from BN02 where tbank_no=? and bank_no=?");
    		paramList.add(tbank_no);
    		paramList.add(bank_no);

            List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"update_date");
            return dbData;
    }
    //取得查詢結果
    //101.06.29 若為全部類別時,可不下總機構條件,只下分支機構代號,查詢
    private List getQryResult(String bank_type,String tbank_no,String bank_no){
    		//查詢條件
    		Calendar now = Calendar.getInstance();
         	String YEAR  = String.valueOf(now.get(Calendar.YEAR)-1911); //回覆值為西元年故需-1911取得民國年;
    		//99.10.15 add 查詢年度100年以前.縣市別不同===============================	     
	      	String wlx01_m_year = (Integer.parseInt(YEAR) < 100)?"99":"100"; 
	        //===================================================================== 	 
    		List paramList = new ArrayList();
    		StringBuffer sqlCmd = new StringBuffer();
    		sqlCmd.append("select cdshareno.cmuse_id,cdshareno.cmuse_name,bn01.bank_no as tbank_no,bn01.bank_name as tbank_name,bn02.bank_no,bn02.bank_name,bn02.update_date ");
			sqlCmd.append("from (select * from bn02 where m_year=?)bn02 LEFT JOIN cdshareno on bank_type = cmuse_id and cmuse_div='001' ");
			sqlCmd.append("LEFT JOIN (select * from bn01 where m_year=?)bn01 on bn02.tbank_no = bn01.bank_no");
			paramList.add(wlx01_m_year);
			paramList.add(wlx01_m_year);
			sqlCmd.append(" where 1=1 ");
			if(!bank_type.equals("Z")){//不為全部類別
			    sqlCmd.append(" and bn02.bank_type=? ");
			    paramList.add(bank_type);
			}
			if(!tbank_no.equals("")){//101.06.28 add
			  	sqlCmd.append(" and bn02.tbank_no=?");
			  	paramList.add(tbank_no);
			}
			
			if(!bank_no.equals("")){//101.06.28 add
			  	sqlCmd.append(" and bn02.bank_no=?");
			  	paramList.add(bank_no);
			}
			  
			sqlCmd.append(" order by bn02.bank_type,bn02.tbank_no,bn02.bank_no ");
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"update_date");
            return dbData;
    }

	public String InsertDB(HttpServletRequest request,String lguser_id,String lguser_name) throws Exception{		
		String errMsg="";
		String tbank_no=((String)request.getParameter("TBANK_NO")==null)?" ":(String)request.getParameter("TBANK_NO");
		String bank_no=((String)request.getParameter("BANK_NO")==null)?" ":(String)request.getParameter("BANK_NO");
		String bank_name=((String)request.getParameter("BANK_NAME")==null)?" ":(String)request.getParameter("BANK_NAME");
		String bn_type="1";
		String bank_type=((String)request.getParameter("BANK_TYPE")==null)?"":(String)request.getParameter("BANK_TYPE");
		String add_user=lguser_id;
		String add_name=lguser_name;
		String bank_b_name=((String)request.getParameter("BANK_B_NAME")==null)?"":(String)request.getParameter("BANK_B_NAME");
		String kind_1="";
		String kind_2="";
		String bn_type2="";
		String exchange_no=((String)request.getParameter("EXCHANGE_NO")==null)?"":(String)request.getParameter("EXCHANGE_NO");
		String user_id=lguser_id;
	    String user_name=lguser_name;
		
		List paramList = new ArrayList();
		StringBuffer sqlCmd = new StringBuffer();		
		List updateDBList = new ArrayList();//0:sql 1:data		
		List updateDBSqlList = new ArrayList();
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data
		
		try {
				Calendar now = Calendar.getInstance();
         	    String YEAR  = String.valueOf(now.get(Calendar.YEAR)-1911); //回覆值為西元年故需-1911取得民國年;
    		    //99.10.15 add 查詢年度100年以前.縣市別不同===============================	     
	      	    String wlx01_m_year = (Integer.parseInt(YEAR) < 100)?"99":"100"; 
	            //===================================================================== 	
				sqlCmd.append("SELECT * FROM BA01 WHERE m_year=? and bank_no=?");
				paramList.add(wlx01_m_year);
				paramList.add(bank_no);

			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
				System.out.println("BA01.size="+data.size());

				if (data.size() != 0){
				    errMsg +=  "此筆分支機構資料已存在無法新增<br>";
				}else{
				    sqlCmd.delete(0, sqlCmd.length());
			        sqlCmd.append("INSERT INTO BN02 VALUES (?,?,?,?,?,?,?,sysdate,?,?,?,?,?,?,?,sysdate,?)");
					dataList.add(tbank_no);       
					dataList.add(bank_no);
					dataList.add(bank_name);
					dataList.add(bn_type);
					dataList.add(bank_type);
					dataList.add(add_user);
					dataList.add(add_name);
					dataList.add(bank_b_name);
					dataList.add(kind_1);
					dataList.add(kind_2);
					dataList.add(bn_type2);
					dataList.add(exchange_no);
					dataList.add(user_id);
					dataList.add(user_name);
					dataList.add(wlx01_m_year);
					
					updateDBDataList.add(dataList);//1:傳內的參數List
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				    
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);

   					sqlCmd.delete(0, sqlCmd.length());
   					dataList =  new ArrayList();//儲存參數的data
   					updateDBDataList = new ArrayList();//儲存參數的List	
   					updateDBSqlList = new ArrayList();							
					
	            	sqlCmd.append("INSERT INTO BA01 VALUES (?,?,?,'1',?,' ',null,?,?,sysdate,?)");
					dataList.add(bank_no);
					dataList.add(bank_name);
					dataList.add(bank_type);
					dataList.add(tbank_no);
					dataList.add(user_id);
					dataList.add(user_name);
					dataList.add(wlx01_m_year);
					
					updateDBDataList.add(dataList);//1:傳內的參數List
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				    
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
					
				    if(DBManager.updateDB_ps(updateDBList)){
				       errMsg += "相關資料寫入資料庫成功";
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


	private String UpdateDB(HttpServletRequest request,String lguser_id,String lguser_name) throws Exception{		
		String errMsg="";
		String tbank_no=((String)request.getParameter("TBANK_NO")==null)?"":(String)request.getParameter("TBANK_NO");
		String bank_no=((String)request.getParameter("BANK_NO")==null)?"":(String)request.getParameter("BANK_NO");
		String bank_name=((String)request.getParameter("BANK_NAME")==null)?" ":(String)request.getParameter("BANK_NAME");
		String bn_type="1";
		String bank_type=((String)request.getParameter("BANK_TYPE")==null)?"":(String)request.getParameter("BANK_TYPE");
		String add_user=lguser_id;
		String add_name=lguser_name;
		String bank_b_name=((String)request.getParameter("BANK_B_NAME")==null)?"":(String)request.getParameter("BANK_B_NAME");
		String kind_1="";
		String kind_2="";
		String bn_type2="";
		String exchange_no=((String)request.getParameter("EXCHANGE_NO")==null)?"":(String)request.getParameter("EXCHANGE_NO");
		String user_id=lguser_id;
	    String user_name=lguser_name;
	    
	    List paramList = new ArrayList();
		StringBuffer sqlCmd = new StringBuffer();		
		List updateDBList = new ArrayList();//0:sql 1:data		
		List updateDBSqlList = new ArrayList();
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data


		try {
				Calendar now = Calendar.getInstance();
         	    String YEAR  = String.valueOf(now.get(Calendar.YEAR)-1911); //回覆值為西元年故需-1911取得民國年;
    		    //99.10.15 add 查詢年度100年以前.縣市別不同===============================	     
	      	    String wlx01_m_year = (Integer.parseInt(YEAR) < 100)?"99":"100"; 
	            //===================================================================== 	
			
				sqlCmd.append("SELECT * FROM BA01 WHERE m_year=? and bank_no=?");
				paramList.add(wlx01_m_year);
				paramList.add(bank_no);

			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
				System.out.println("BA01.size="+data.size());

				if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法修改<br>";
				}else{
				    //insert BN04_LOG===================================================
				    sqlCmd.delete(0, sqlCmd.length());
				    sqlCmd.append(" INSERT INTO BN04_LOG ");
					sqlCmd.append(" select tbank_no,bank_no,bank_name,bn_type,bank_type,");
					sqlCmd.append(" add_user,add_name,add_date,bank_b_name,kind_1,kind_2,");
					sqlCmd.append(" bn_type2,exchange_no,user_id,user_name,update_date,");
					sqlCmd.append(" ?,?,sysdate,'1','U'");
					sqlCmd.append(" from BN02");
					sqlCmd.append(" WHERE m_year=? and bank_no=?");					
					dataList.add(user_id);
					dataList.add(user_name);
					dataList.add(wlx01_m_year);
					dataList.add(bank_no);
					updateDBDataList.add(dataList);//1:傳內的參數List
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				    
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
				    //UPDATE BN02=========================================================================
				    sqlCmd.delete(0, sqlCmd.length());
   					dataList =  new ArrayList();//儲存參數的data
   					updateDBDataList = new ArrayList();//儲存參數的List	
   					updateDBSqlList = new ArrayList();	   					
				    sqlCmd.append("UPDATE BN02 SET ");
				    sqlCmd.append(" bank_name=?");
					sqlCmd.append(",bank_b_name=?");
					sqlCmd.append(",exchange_no=?");
					sqlCmd.append(",user_id=?");
					sqlCmd.append(",user_name=?");
					sqlCmd.append(",update_date=sysdate");
					sqlCmd.append(" where m_year=? and tbank_no=? and bank_no=?");
					dataList.add(bank_name);
					dataList.add(bank_b_name);
					dataList.add(exchange_no);
					dataList.add(user_id);
					dataList.add(user_name);
					dataList.add(wlx01_m_year);
					dataList.add(tbank_no);
					dataList.add(bank_no);
					updateDBDataList.add(dataList);//1:傳內的參數List
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				    
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);

	            	//UPDATE BA01=========================================================================
	            	sqlCmd.delete(0, sqlCmd.length());
   					dataList =  new ArrayList();//儲存參數的data
   					updateDBDataList = new ArrayList();//儲存參數的List	
   					updateDBSqlList = new ArrayList();			
	            	sqlCmd.append("UPDATE BA01 SET ");
				    sqlCmd.append(" bank_name=?");
					sqlCmd.append(" where m_year=? and bank_no=?");
					dataList.add(bank_name);			
					dataList.add(wlx01_m_year);			   
					dataList.add(bank_no);						   
					updateDBDataList.add(dataList);//1:傳內的參數List
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				    
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
					
					if(DBManager.updateDB_ps(updateDBList)){
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

    private String DeleteDB(HttpServletRequest request,String lguser_id,String lguser_name) throws Exception{		
		String errMsg="";
		String tbank_no=((String)request.getParameter("TBANK_NO")==null)?" ":(String)request.getParameter("TBANK_NO");
		String bank_no=((String)request.getParameter("BANK_NO")==null)?" ":(String)request.getParameter("BANK_NO");
		String user_id=lguser_id;
	    String user_name=lguser_name;

    	List paramList = new ArrayList();
		StringBuffer sqlCmd = new StringBuffer();		
		List updateDBList = new ArrayList();//0:sql 1:data		
		List updateDBSqlList = new ArrayList();
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data

		try {
				Calendar now = Calendar.getInstance();
         	    String YEAR  = String.valueOf(now.get(Calendar.YEAR)-1911); //回覆值為西元年故需-1911取得民國年;
    		    //99.10.15 add 查詢年度100年以前.縣市別不同===============================	     
	      	    String wlx01_m_year = (Integer.parseInt(YEAR) < 100)?"99":"100"; 
	            //===================================================================== 	
			
				sqlCmd.append("SELECT * FROM BA01 WHERE m_year=? and bank_no=?");
				paramList.add(wlx01_m_year);
				paramList.add(bank_no);
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
				System.out.println("BA01.size="+data.size());
				if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法刪除<br>";
				}else{
				    //insert BN04_LOG===================================================
				    sqlCmd.delete(0, sqlCmd.length());
				    sqlCmd.append(" INSERT INTO BN04_LOG ");
					sqlCmd.append(" select tbank_no,bank_no,bank_name,bn_type,bank_type,");
					sqlCmd.append(" add_user,add_name,add_date,bank_b_name,kind_1,kind_2,");
					sqlCmd.append(" bn_type2,exchange_no,");
					sqlCmd.append(" user_id,user_name,update_date,");
					sqlCmd.append(" ?,?,sysdate,'1','D'");
					sqlCmd.append(" from BN02");
					sqlCmd.append(" WHERE m_year=? and tbank_no=? and bank_no=?");
					dataList.add(user_id);
					dataList.add(user_name);
					dataList.add(wlx01_m_year);
					dataList.add(tbank_no);
					dataList.add(bank_no);
					updateDBDataList.add(dataList);//1:傳內的參數List
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				    
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
				    //DELETE BN02=========================================================================
				    sqlCmd.delete(0, sqlCmd.length());
   					dataList =  new ArrayList();//儲存參數的data
   					updateDBDataList = new ArrayList();//儲存參數的List	
   					updateDBSqlList = new ArrayList();		
				    sqlCmd.append("DELETE from BN02 where m_year=? and tbank_no = ? and bank_no=?");
				    dataList.add(wlx01_m_year);
				    dataList.add(tbank_no);
				    dataList.add(bank_no);
		            updateDBDataList.add(dataList);//1:傳內的參數List
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				    
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
					//DELETE BA01=========================================================================
					sqlCmd.delete(0, sqlCmd.length());
   					dataList =  new ArrayList();//儲存參數的data
   					updateDBDataList = new ArrayList();//儲存參數的List	
   					updateDBSqlList = new ArrayList();
		            sqlCmd.append("DELETE from BA01 where m_year=? and bank_no = ?");
		            dataList.add(wlx01_m_year);
		            dataList.add(bank_no);
		            updateDBDataList.add(dataList);//1:傳內的參數List
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				    
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
		            //99.04.21 fix 刪除分支機構代號時,一併刪除分支機構基本資料
		            //99.10.14 fix 當有分支機構基本資料時,才需刪除分支機構基本資料
		            sqlCmd.delete(0, sqlCmd.length());
		            paramList = new ArrayList();
		            sqlCmd.append(" select * from WLX02 where m_year=? and tbank_no = ? and bank_no=?");
		            paramList.add(wlx01_m_year);
		            paramList.add(tbank_no);
		            paramList.add(bank_no);
		            data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");		            
				    System.out.println("WLX02.size="+data.size());
				    if (data.size() != 0){
				       sqlCmd.delete(0, sqlCmd.length());
   					   updateDBDataList = new ArrayList();//儲存參數的List	
   					   updateDBSqlList = new ArrayList();
				       sqlCmd.append("DELETE from WLX02 where m_year=? and tbank_no = ? and bank_no=?");
				       updateDBDataList.add(paramList);//1:傳內的參數List
					   updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				    
					   updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					   updateDBList.add(updateDBSqlList);		            
		            }
					if(DBManager.updateDB_ps(updateDBList)){
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

%>    