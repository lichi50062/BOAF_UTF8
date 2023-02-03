<%
// 94.01.20 create by 2295 
// 94.01.31 add 權限 by 2295
// 94.03.23 add 營運中/已裁撤 by 2295
// 99.12.23 add 查詢日期for100年以後取得新縣市別.100年以前取得舊縣市別 by 2295
// 101.06   add 報表欄位 by 2968
//102.06.27 add 操作歷程寫入log by2968
//103.01.21 add BOAF/MIS共用畫面 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.report.reportUtil" %>
<%@ page import="java.util.*" %>
<%@ page import="java.io.*" %>
<%@include file="./include/Header.include" %>

<%
	String bank_type = Utility.getTrimString(dataMap.get("bank_type"));    
    String firstStatus = Utility.getTrimString(dataMap.get("firstStatus"));	
	if(firstStatus.equals("true")){//若從Menu點選時,先清空session裡的資料
	   session.setAttribute("S_YEAR",null);//99.12.23 add
	   session.setAttribute("S_MONTH",null);//99.12.23 add
	   session.setAttribute("CANCEL_NO",null);   
	   session.setAttribute("HSIEN_ID",null);   
	   session.setAttribute("BankList",null);    	
	   session.setAttribute("btnFieldList",null);	
	   session.setAttribute("SortList",null);   	   
	   session.setAttribute("SortBy",null);   
	   session.setAttribute("excelaction",null);   
	   session.setAttribute("nowbank_type",null);   
	   session.setAttribute("position_code",null);   
	}
    if(!Utility.CheckPermission(request,report_no)){//無權限時,導向到LoginError.jsp//95.11.03 fix 使用Utility.CheckPermission檢核權限 by 2295
        rd = application.getRequestDispatcher( LoginErrorPgName );        
    }else{            
    	//set next jsp 	
    	//將選取的縣市別,金融機構代碼,報表欄位,排序的報表欄位寫到session=======================================        	
    	if(!Utility.getTrimString(dataMap.get("BankList")).equals("")){
    	   session.setAttribute("BankList",(String)request.getParameter("BankList"));    	
    	}       	
    	if(!Utility.getTrimString(dataMap.get("btnFieldList")).equals("")){
    	   session.setAttribute("btnFieldList",(String)request.getParameter("btnFieldList"));	
    	}    	    	
    	if(!Utility.getTrimString(dataMap.get("SortList")).equals("")){
    	   System.out.println(report_no+".SortList="+(String)request.getParameter("SortList"));
           session.setAttribute("SortList",(String)request.getParameter("SortList"));   
        }
        //94.03.23 add 營運中/已裁撤=============================================================        
        if(!Utility.getTrimString(dataMap.get("CANCEL_NO")).equals("")){
           session.setAttribute("CANCEL_NO",(String)request.getParameter("CANCEL_NO"));   
        }
        //=======================================================================================    
        //99.12.23 add 查詢日期====================================================================
        if(!Utility.getTrimString(dataMap.get("S_YEAR")).equals("")){
           session.setAttribute("S_YEAR",(String)request.getParameter("S_YEAR"));   
        }
        if(!Utility.getTrimString(dataMap.get("S_MONTH")).equals("")){
           session.setAttribute("S_MONTH",(String)request.getParameter("S_MONTH"));   
        }    
        //========================================================================================
        if(!Utility.getTrimString(dataMap.get("HSIEN_ID")).equals("")){
           session.setAttribute("HSIEN_ID",(String)request.getParameter("HSIEN_ID"));   
        }
        if(!Utility.getTrimString(dataMap.get("SortBy")).equals("")){
           session.setAttribute("SortBy",(String)request.getParameter("SortBy"));   
        }
        if(!Utility.getTrimString(dataMap.get("excelaction")).equals("")){
           session.setAttribute("excelaction",(String)request.getParameter("excelaction"));   
        }                
        if(!Utility.getTrimString(dataMap.get("bank_type")).equals("")){
           session.setAttribute("nowbank_type",(String)request.getParameter("bank_type"));   
        }
        //95.11.10 add 儲存DS_bank_type. menu的機構類別===================================================================
        if(!Utility.getTrimString(dataMap.get("DS_bank_type")).equals("")){           
           session.setAttribute("DS_bank_type",Utility.getTrimString(dataMap.get("DS_bank_type")));
           session.setAttribute("nowbank_type",Utility.getTrimString(dataMap.get("DS_bank_type")));//100.05.04
        }
        if(firstStatus.equals("true") && (!Utility.getTrimString(dataMap.get("bank_type")).equals(""))){           
           session.setAttribute("DS_bank_type",Utility.getTrimString(dataMap.get("bank_type"))); 
        }
        //=================================================================================================
        if(act.equals("RptColumn")){//報表欄位   
			List FieldList = Utility.getFileData(Utility.getProperties("schemaDir")+System.getProperty("file.separator")+"WLX01.TXT");	
        	//List FieldList = Utility.getFileData("D:\\WorkSpace\\pboaf1\\MIS\\schemaDir"+System.getProperty("file.separator")+"WLX01.TXT");						
       	 	System.out.println("FieldListRoot ===> "+Utility.getProperties("schemaDir")+System.getProperty("file.separator")+"WLX01.TXT");
        	String FieldListString = "";			
			for(int i=0;i<FieldList.size();i++){			    			    
			    FieldListString += Utility.ISOtoUTF8((String)FieldList.get(i));
			    if(i < FieldList.size() - 1) FieldListString +=",";   
			}
			FieldListString = FieldListString.substring(1,FieldListString.length());//104.03.09 add 去掉第一個字元?
			request.setAttribute("FieldList",FieldListString); 		
			System.out.println(report_no+".FieldList ===>"+FieldListString);	
        	rd = application.getRequestDispatcher( RptColumnPgName );        
        }else if(act.equals("BankList")){//金融機構
            System.out.println("btnFieldList="+(String)request.getParameter("btnFieldList"));            
            rd = application.getRequestDispatcher( BankListPgName );        
    	}else if(act.equals("RptOrder")){//排序欄位       	     	       	        	    
        	rd = application.getRequestDispatcher( RptOrderPgName ); 
        }else if(act.equals("RptStyle")){//報表格式            
        	rd = application.getRequestDispatcher( RptStylePgName );         
        }else if(act.equals("SaveRpt") || act.equals("ReadRpt")){//儲存格式檔|讀取格式檔        
            if(act.equals("SaveRpt")){//儲存格式檔
               actMsg = reportUtil.saveReport_BR(request,lguser_id,report_no);            
               alertMsg = "儲存";
            }   
            if(act.equals("ReadRpt")){//讀取格式檔   
               actMsg = reportUtil.readReport_BR(request,lguser_id,report_no);            
               alertMsg = "讀取";
            } 
			if(actMsg.equals("")){
				alertMsg += "報表格式成功";    	           	       	    	       	    
    	    }else{    	        
    	       	alertMsg += "報表格式失敗:"+actMsg;    	           	       	
    	       	actMsg = "";
    	    }
    	    webURL_Y = "/pages/"+report_no+".jsp?act=RptStyle";    	       	    	       	   	   			       
        	rd = application.getRequestDispatcher( nextPgName );         
        }else if(act.equals("createRpt")){//產生Excel報表
            String excelAction = (String)session.getAttribute("excelaction");
            this.InsertWlXOPERATE_LOG(request,lguser_id,report_no,"P");
            rd = application.getRequestDispatcher( RptCreatePgName +"?act="+excelAction);         
        }
        
    	request.setAttribute("actMsg",actMsg);    
    	request.setAttribute("alertMsg",alertMsg);
    	request.setAttribute("webURL_Y",webURL_Y);
    	request.setAttribute("webURL_N",webURL_N);   
    }        
    
%>
<%@include file="./include/Tail.include" %>

<%!
    private final static String report_no = "BR002W";
    private final static String nextPgName = "/pages/ActMsg.jsp";    
    private final static String BankListPgName = "/pages/"+report_no+"_BankList.jsp";    
    private final static String RptColumnPgName = "/pages/"+report_no+"_RptColumn.jsp";        
    private final static String RptOrderPgName = "/pages/"+report_no+"_RptOrder.jsp";        
    private final static String RptStylePgName = "/pages/"+report_no+"_RptStyle.jsp";        
    private final static String RptCreatePgName = "/pages/"+report_no+"_Excel.jsp";        
    private final static String LoginErrorPgName = "/pages/LoginError.jsp"; 
    
    public String InsertWlXOPERATE_LOG(HttpServletRequest request,String lguser_id,String program_id,String update_type) throws Exception{    	
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList() ;
		String errMsg="";
	    try {
	        sqlCmd.append(" INSERT INTO WlXOPERATE_LOG(muser_id,use_Date,program_id,ip_address,update_type)");
	        sqlCmd.append("                     VALUES(?,sysdate,?,?,?) ");
	        paramList.add(lguser_id);
	        paramList.add(program_id);
	        paramList.add(request.getRemoteAddr());//ipAddress
	        paramList.add(update_type);//操作類別 I-新增，U-異動，D-刪除，Q-明細，P-列印
	        if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){
				errMsg = errMsg + "相關資料寫入資料庫成功";					
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