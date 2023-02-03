<%
//101.08.06 created 農信保--保證案件月報表(M201) by 2968
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.report.reportUtil" %>
<%@ page import="java.util.*" %>
<%@include file="./include/Header.include" %>

<%	
	String last_act = Utility.getTrimString(dataMap.get("last_act"));//95.12.08	
	String bank_type = Utility.getTrimString(dataMap.get("bank_type"));    
    String firstStatus = Utility.getTrimString(dataMap.get("firstStatus"));
    String template = Utility.getTrimString(dataMap.get("template"));
    String template_list = Utility.getTrimString(dataMap.get("template_list"));
    String menu = Utility.getTrimString(dataMap.get("menu"));//顯示BOAF menu用		    		
	if(bank_type.equals("")){
	   bank_type = (session.getAttribute("nowbank_type")==null)?"":(String)session.getAttribute("nowbank_type");
	} 
	System.out.println(report_no+".act="+act);
	System.out.println(report_no+".bank_type="+bank_type);
	if(firstStatus.equals("true")){//若從Menu點選時,先清空session裡的資料	   
	   session.setAttribute("btnFieldList",null);		  
	   session.setAttribute("excelaction",null);   
	   session.setAttribute("nowbank_type",null);   
	   session.setAttribute("S_YEAR",null);   
	   session.setAttribute("S_MONTH",null);   
	   session.setAttribute("E_YEAR",null);   
	   session.setAttribute("E_MONTH",null);   
	   session.setAttribute("Unit",null);
	   System.out.println("clear all session");
	}
	
    if(!Utility.CheckPermission(request,report_no)){//無權限時,導向到LoginError.jsp//95.11.03 fix 使用Utility.CheckPermission檢核權限 by 2295
        rd = application.getRequestDispatcher( LoginErrorPgName );        
    }else{            
    	//set next jsp 	
    	//將選取的報表欄位寫到session============================================================================================================ 
    	if(!Utility.getTrimString(dataMap.get("btnFieldList")).equals("")){
    	   session.setAttribute("btnFieldList",Utility.getTrimString(dataMap.get("btnFieldList")));	
    	}    	
       	//=======================================================================================================================================    	
        if(!Utility.getTrimString(dataMap.get("excelaction")).equals("")){
           session.setAttribute("excelaction",(String)request.getParameter("excelaction"));   
        }     
           
        if(!Utility.getTrimString(dataMap.get("bank_type")).equals("")){
           //95.09.06 此次所選的bank_type與上次不同清除已勾選的勾選的報表欄位/sort報表欄位
           if(session.getAttribute("nowbank_type") != null && !((String)session.getAttribute("nowbank_type")).equals(bank_type)){                           
              session.setAttribute("btnFieldList",null);//清除已勾選的勾選的報表欄位)              
           }                 
           session.setAttribute("nowbank_type",Utility.getTrimString(dataMap.get("bank_type")));   
        }else{//request無所選bank_type
           if(session.getAttribute("nowbank_type") != null){    
              bank_type = (String)session.getAttribute("nowbank_type");
           }
        }   
        //95.11.10 add 儲存DS_bank_type. menu的機構類別===================================================================
        if(!Utility.getTrimString(dataMap.get("DS_bank_type")).equals("")){           
           session.setAttribute("DS_bank_type",Utility.getTrimString(dataMap.get("DS_bank_type")));
           session.setAttribute("nowbank_type",Utility.getTrimString(dataMap.get("DS_bank_type")));//100.05.04
        }
        if(firstStatus.equals("true") && (!Utility.getTrimString(dataMap.get("bank_type")).equals(""))){           
           session.setAttribute("DS_bank_type",Utility.getTrimString(dataMap.get("bank_type"))); 
        }              
        //===============================================================================================================        
        //金額單位
        if(!Utility.getTrimString(dataMap.get("Unit")).equals("")){
           session.setAttribute("Unit",(String)request.getParameter("Unit"));   
        }
        //年-begin
        if(!Utility.getTrimString(dataMap.get("S_YEAR")).equals("")){
           session.setAttribute("S_YEAR",(String)request.getParameter("S_YEAR"));
        }
        //年-end
        if(!Utility.getTrimString(dataMap.get("E_YEAR")).equals("")){
           session.setAttribute("E_YEAR",(String)request.getParameter("E_YEAR"));   
        }
        //年月區間
        //月-begin
        if(!Utility.getTrimString(dataMap.get("S_MONTH")).equals("")){
           session.setAttribute("S_MONTH",(String)request.getParameter("S_MONTH"));   
        }
        //月-end
        if(!Utility.getTrimString(dataMap.get("E_MONTH")).equals("")){
           session.setAttribute("E_MONTH",(String)request.getParameter("E_MONTH"));   
        }        
        //=================================================================================================
        if(act.equals("RptColumn")){//報表欄位    
            List FieldList = null;        
	        StringBuffer sql = new StringBuffer();		
	        List paramList = new ArrayList();
	        DataObject bean = null;
            sql.append(" select cano,rulename from rptrule_acgf where acc_tr_type=? and type='0' order by to_number(output_order) ");
     		paramList.add("M201");
     		FieldList = DBManager.QueryDB_SQLParam(sql.toString(),paramList,""); 
			
			String FieldListString = "";			
			for(int i=0;i<FieldList.size();i++){			    			    
			    bean = (DataObject)FieldList.get(i);
			    FieldListString += (String)bean.getValue("cano")+"+"+(String)bean.getValue("rulename");
			    if(i < FieldList.size() - 1) FieldListString +=",";   
			}
			request.setAttribute("FieldList",FieldListString); 		
			System.out.println(report_no+".FieldList="+FieldListString);	
        	rd = application.getRequestDispatcher( RptColumnPgName );
        }else if(act.equals("RptStyle")){//報表格式     
            List templateList = reportUtil.getTemplateList(lguser_id,bank_type,report_no);//取得範本資料//95.11.03使用reportUtil 
            request.setAttribute("templateList",templateList);       
        	rd = application.getRequestDispatcher( RptStylePgName );         
        }else if(act.equals("SaveRpt") || act.equals("ReadRpt") || act.equals("DeleteRpt")){//儲存格式檔|讀取格式檔|刪除格式檔        
            if(act.equals("SaveRpt")){//儲存格式檔
               actMsg = reportUtil.saveReport(request,lguser_id,template,report_no);//95.11.03使用reportUtil             
               alertMsg = "儲存";
            }   
            if(act.equals("DeleteRpt")){//刪除格式檔
               //templast_list ex-->test:A111111111:20060824.txt               
               List tmpList = Utility.getStringTokenizerData(template_list,":");
               System.out.println((String)tmpList.get(1)+":"+(String)tmpList.get(0)+":"+(String)tmpList.get(2)+":"+report_no);
               if(tmpList != null && tmpList.size() != 0){
                  actMsg = reportUtil.deleteReport(request,(String)tmpList.get(1),(String)tmpList.get(0),(String)tmpList.get(2),report_no);//95.11.03使用reportUtil             
               }else{
                 actMsg = "範本刪除失敗";
               }
               alertMsg = "刪除";
            }
            if(act.equals("ReadRpt")){//讀取格式檔   
               //templast_list ex-->test:A111111111:20060824.txt               
               List tmpList = Utility.getStringTokenizerData(template_list,":");
               if(tmpList != null && tmpList.size() != 0){
                  actMsg = reportUtil.readReport(request,(String)tmpList.get(1),(String)tmpList.get(0),(String)tmpList.get(2),report_no);//95.11.03使用reportUtil             
               }else{
                 actMsg = "範本讀取失敗";
               }
               alertMsg = "讀取";
            } 
			if(actMsg.equals("")){
				alertMsg += "報表格式成功";    	           	       	    	       	    
    	    }else{    	        
    	       	alertMsg += "報表格式失敗:"+actMsg;    	           	       	
    	       	actMsg = "";
    	    }
    	    if(act.equals("ReadRpt")){
    	       webURL_Y = "/pages/"+report_no+".jsp?act=RptColumn&last_act=ReadRpt"; 	       	    	       	   	   			       
    	    }else{
    	       webURL_Y = "/pages/"+report_no+".jsp?act=RptStyle";    	       	    	       	   	   			       
    	    }   
        	rd = application.getRequestDispatcher( nextPgName );         
        }else if(act.equals("createRpt")){//產生Excel報表
            String excelAction = (String)session.getAttribute("excelaction");
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
    private final static String report_no = "DS053W";
    private final static String nextPgName = "/pages/ActMsg.jsp"; 
    private final static String RptColumnPgName = "/pages/"+report_no+"_RptColumn.jsp";
    private final static String RptStylePgName = "/pages/"+report_no+"_RptStyle.jsp";  
    private final static String RptCreatePgName = "/pages/"+report_no+"_Excel.jsp";  
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";    
%>    