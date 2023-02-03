<%
// 99.12.14 fix SQLInjection by 2479
//100.05.10 fix 取得機構名稱  by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.util.*,java.io.*" %>
<%@ page import="org.apache.poi.poifs.filesystem.*,org.apache.poi.hssf.usermodel.*" %>
<%@ page import="com.tradevan.util.Utility"%>
<%@ page import="com.tradevan.util.DBManager"%>
<%@ page import="com.tradevan.util.dao.DataObject"%>
<%@ page import="com.tradevan.util.report.FR009W_Excel" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Enumeration" %>
<%@ page import="java.util.*" %>
<%@include file="./include/Header.include" %>

<%
        response.setContentType("application/msexcel;charset=UTF-8");//以上這行設定本網頁為excel格式的網頁
        act = dataMap.get("act")==null?"view" : (String)dataMap.get("act") ;
        String bank_type = Utility.getTrimString(dataMap.get("bank_type")) ;
        String BANK_DATA =  Utility.getTrimString(dataMap.get("BANK_NO"));			  
        String BANK_NO = BANK_DATA.substring(0,BANK_DATA.indexOf("/"));
        String BANK_NAME = BANK_DATA.substring(BANK_DATA.indexOf("/")+8,BANK_DATA.length());
        String S_YEAR =Utility.getTrimString(dataMap.get("S_YEAR")) ;
        String S_MONTH = Utility.getTrimString(dataMap.get("S_MONTH")) ; 
        String Unit = Utility.getTrimString(dataMap.get("Unit")) ; 
        
        
        String filename="信用部淨值占風險性資產比率計算表.xls";
        
        System.out.println("FR009W_Excel.S_YEAR="+S_YEAR);
        System.out.println("FR009W_Excel.act="+act);
      

   		if(!Utility.CheckPermission(request,report_no)){//無權限時,導向到LoginError.jsp
        	rd = application.getRequestDispatcher( LoginErrorPgName );
    	}else{
	    	//set next jsp
	    	if(act.equals("Qry")){
	    	   rd = application.getRequestDispatcher( QryPgName + "?bank_type="+bank_type);
	    	}else if(act.equals("view")){
      		   //以上這行設定傳送到前端瀏覽器時的檔名為test1.xls
      		   //就是靠這一行，讓前端瀏覽器以為接收到一個excel檔
      			response.setHeader("Content-disposition","inline; filename=view.xls");
   			}else if (act.equals("download")){
      			response.setHeader("Content-Disposition","attachment; filename=download.xls");
   			}
   			if(act.equals("view") || act.equals("download")){
   				try{
   				
   				    List paramList = new ArrayList();
            		HashMap h = new HashMap();
            		boolean isEmpty = true; 
            		paramList.add(S_YEAR);
                    paramList.add(S_MONTH);
                    paramList.add(BANK_NO);
				    String sqlCmd = " select amt from A05 where acc_code='910500' "
				    	+ " and   m_year =? and   m_month = ? "
				    	+ " and   a05.bank_code = ? ";
                    List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"amt");        
				    //List dbData = DBManager.QueryDB(sqlCmd, "amt");
				    if(dbData != null && dbData.size() > 0) {
				    	h.put("910500", ((DataObject)dbData.get(0)).getValue("amt").toString());
				    } else {
				    	h.put("910500","0");
				    }
				                  				
				    sqlCmd = " select a05.m_year,a05.m_month,ncacno.acc_code,ncacno.acc_name," + " substr(ncacno.acc_code,length(ncacno.acc_code)) as acc_code_l," + " a05_assumed.assumed as assumed,"
				    		+ " decode(substr(ncacno.acc_code,length(ncacno.acc_code)),'P',nvl(a05.amt,0)/1000,nvl(a05.amt,0))as amt,"
				    		+ " decode(substr(ncacno.acc_code,length(ncacno.acc_code)),'P',nvl(a05.amt,0)/1000,nvl(a05_assumed.assumed,1)*nvl(a05.amt,0))" + " as amt_assumed," + " nvl(a05.amt_name,'') as amt_name"
				    		+ " from ncacno join a05 on ncacno.acc_code=a05.acc_code left join a05_assumed on a05.acc_code=a05_assumed.acc_code" + " where acc_tr_type='A05' and ncacno.acc_code like '92%'" 
                            + " and   m_year = ? and   m_month = ? "
				    		+ " and   a05.bank_code = ? ";
                    dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"m_year,m_month,amt,amt_assumed");        
				    //dbData = DBManager.QueryDB(sqlCmd, "m_year,m_month,amt,amt_assumed");
                    
				    h.put("space", "");
				    
				    String tmp = "";
				    for (int k = 0; k < dbData.size(); k++) {
				    	isEmpty = false;
				    	//System.out.println("k= " + k);
				    	DataObject obj = (DataObject) dbData.get(k);
				    	String accCode = (String) obj.getValue("acc_code") != null ? (String) obj.getValue("acc_code") : "";
				    	System.out.println("accCode= " + accCode);
				    	if (accCode.indexOf("N") > 0) {
				    		String amtName = obj.getValue("amt_name") != null ? (String) obj.getValue("amt_name") : "";
				    		System.out.println("amtName= " + amtName);
				    		tmp += amtName;
				    		h.put(accCode, amtName);
				    	}
				    	else if (accCode.indexOf("P") > 0) {
				    		String amtAssumed = obj.getValue("amt_assumed").toString();
				    		System.out.println("amtAssumed= " + amtAssumed);
				    		double d = 0;
				    		try {
				    			d= Double.parseDouble(amtAssumed);
				    		}catch (Exception e) {
				    			d=0;
				    		}
				    		h.put(accCode, Double.toString(d));
				    	}
				    	else {
				    		String amt = obj.getValue("amt").toString();
				    		System.out.println("amt= " + amt);
				    		h.put(accCode, amt);
				    	}
				    }            		
              		
	    			actMsg =FR009W_Excel.createRpt(S_YEAR, S_MONTH, BANK_NAME,  h, isEmpty);
	    			System.out.println("createRpt="+actMsg);
	    			System.out.println("filename="+Utility.getProperties("reportDir")+System.getProperty("file.separator")+filename);
					FileInputStream fin = new FileInputStream(Utility.getProperties("reportDir")+System.getProperty("file.separator")+filename);
					ServletOutputStream out1 = response.getOutputStream();
					byte[] line = new byte[8196];
					int getBytes=0;
					while( ((getBytes=fin.read(line,0,8196)))!=-1 ){
						out1.write(line,0,getBytes);
						out1.flush();
	    			}

					fin.close();
					out1.close();

				}catch(Exception e){
	   				System.out.println(e.getMessage());
				}
   			}
   		}
   		request.setAttribute("actMsg",actMsg);
%>
<%@include file="./include/Tail.include" %>

<%!
    private final static String report_no = "FR009W" ;
    private final static String nextPgName = "/pages/ActMsg.jsp";
    private final static String QryPgName = "/pages/"+report_no+"_Qry.jsp";
    private final static String RptCreatePgName = "/pages/"+report_no+"_Excel.jsp";
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";    
%>