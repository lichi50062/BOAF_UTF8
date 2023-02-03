<%
//create 農漁會信用部財業務資料_動態 by 2968
//105.03.23 add for 地方主管機關使用 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.util.*,java.io.*" %>
<%@ page import="org.apache.poi.poifs.filesystem.*,org.apache.poi.hssf.usermodel.*" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.report.RptWR002W" %>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@include file="./include/Header.include" %>

<%
	response.setContentType("application/msexcel;charset=UTF-8");//以上這行設定本網頁為excel格式的網頁
	String BANK_NO = Utility.getTrimString(dataMap.get("BANK_NO"));
	String bank_type = ( dataMap.get("bank_type")==null ) ? (String)session.getAttribute("bank_type") : Utility.getTrimString(dataMap.get("bank_type"));			
	String S_YEAR = Utility.getTrimString(dataMap.get("S_YEAR"));     
	String S_MONTH = Utility.getTrimString(dataMap.get("S_MONTH")); 
	String Unit = Utility.getTrimString(dataMap.get("Unit")); 
    String filename="農漁會信用部財業務資料_動態.xls";
   
    if(!Utility.CheckPermission(request,report_no)){//無權限時,導向到LoginError.jsp
    	rd = application.getRequestDispatcher( LoginErrorPgName );
    }else{
	 	//set next jsp
	 	if(act.equals("Qry")){
	 	   rd = application.getRequestDispatcher( QryPgName + "?bank_type="+bank_type);
	 	   session.setAttribute("nowbank_type",bank_type);//100.06.24
	       //request.setAttribute("TBank",getBankList(request) );//按照直轄市在前.其他縣市在後排序.
	       request.setAttribute("TBank",Utility.getBankList(request) );//105.03.23 add for 地方主管機關使用	       
	 	}else if(act.equals("view")){
    	   //以上這行設定傳送到前端瀏覽器時的檔名為test1.xls
    	   //就是靠這一行，讓前端瀏覽器以為接收到一個excel檔
    		response.setHeader("Content-disposition","inline; filename=view.xls");
    	}else if (act.equals("download")){
    		response.setHeader("Content-Disposition","attachment; filename=download.xls");
    	}
    	if(act.equals("view") || act.equals("download")){
    		try{
	 			actMsg =RptWR002W.createRpt(S_YEAR,S_MONTH,Unit,BANK_NO);
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
    private final static String report_no = "WR002W";
    private final static String nextPgName = "/pages/ActMsg.jsp";
    private final static String QryPgName = "/pages/"+report_no+"_Qry.jsp";
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";  
    
    public static List getBankList(HttpServletRequest request){
        HttpSession session = request.getSession();
        List tbankList = null;
        List paramList = new ArrayList();
		StringBuffer sql = new StringBuffer();
        try{
            //95.11.13 取得登入者資訊=================================================================================================
            String muser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		
            String muser_name = ( session.getAttribute("muser_name")==null ) ? "" : (String)session.getAttribute("muser_name");		
            String muser_type = ( session.getAttribute("muser_type")==null ) ? "" : (String)session.getAttribute("muser_type");			
            String muser_bank_type = ( session.getAttribute("bank_type")==null ) ? "" : (String)session.getAttribute("bank_type");			
            String muser_tbank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");			
            //==============================================================================================================    	    
            String bank_type = (session.getAttribute("nowbank_type")==null)?"6":(String)session.getAttribute("nowbank_type");	    
    	
            sql.append(" select bn01.m_year,bn01.bn_type,bn01.bank_type,wlx01.hsien_id,bn01.bank_no,bn01.bank_name from bn01 ");
            sql.append(" LEFT JOIN WLX01 on bn01.bank_no = WLX01.bank_no");      
            sql.append(" and bn01.bank_type in (?,?)");   
            paramList.add("6");
            paramList.add("7");
            sql.append(" ,v_bank_location e");
            sql.append(" where wlx01.m_year=bn01.m_year ");            
            sql.append("   and bn01.m_year = e.m_year ");
            sql.append("   and bn01.bank_no = e.bank_no ");
            sql.append(" order by e.m_year,e.fr001w_output_order,bn01.bank_type,bn01.bank_no ");                       
            //====================================================================================================================
            tbankList =  DBManager.QueryDB_SQLParam(sql.toString(),paramList,"m_year");     
        }catch(Exception e){
            System.out.println("Utility.getBankList Error:"+e.getMessage());
        }
        return tbankList;
    }
%>
