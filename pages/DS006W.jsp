<%
// 95.10.20 create by 2295 
// 95.10.23 fix A06.農會.全部.都讀取bank_type=6的設定檔.漁會.全部.都讀取bank_type=6的設定檔
// 95.11.03 fix 使用reportUtil執行儲存報表格式檔.讀取報表格式檔.刪除報表格式檔.取得範本資料 by 2295
// 		   fix 使用Utility.CheckPermission檢核權限 by 2295
// 95.11.14 add 儲存DS_bank_type. menu的機構類別 by 2295
// 95.12.05 add 增加年月區間 by 2295
// 95.12.07 add 金融機構代號若本來選全部->各信用部 or 各信用部->全部,清空已選報表欄位/排序欄位 by 2295
// 95.12.08 add 不為讀取完報表格式檔,才判斷是否清除已選報表欄位/排序欄位 by 2295
// 99.05.14 fix 套用Header.include by2808 
//100.05.04 fix 儲存session中的nowbank_type=DS_bank_type by 2295
//103.02.10 add 103/01以後,漁會套用新表格(增加/異動科目代號) by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.report.reportUtil" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Enumeration" %>
<%@ page import="java.util.HashMap" %>
<%@ page import="java.util.Hashtable" %>
<%@ page import="java.io.FileOutputStream" %>
<%@ page import="java.io.BufferedOutputStream" %>
<%@ page import="java.io.File" %>
<%@ page import="java.io.PrintStream" %>
<%@ page import="java.io.*" %>
<%@ page import="java.util.*" %>
<%@include file="./include/Header.include" %>

<%
	String last_act = ( request.getParameter("last_act")==null ) ? "" : (String)request.getParameter("last_act");//95.12.08			    	
	String bank_type = ( request.getParameter("bank_type")==null ) ? "" : (String)request.getParameter("bank_type");			    	
	//String acc_div = ( request.getParameter("acc_div")==null ) ? "01" : (String)request.getParameter("acc_div");			    	
	String firstStatus = ( request.getParameter("firstStatus")==null ) ? "" : (String)request.getParameter("firstStatus");			    	
	String template = ( request.getParameter("template")==null ) ? "" : (String)request.getParameter("template");
	String template_list = ( request.getParameter("template_list")==null ) ? "" : (String)request.getParameter("template_list");	
	if(bank_type.equals("")){
	   bank_type = (session.getAttribute("nowbank_type")==null)?"":(String)session.getAttribute("nowbank_type");
	}   
	String nowBank_List = "";//95.12.07 add
	String lastBank_List = "";//95.12.07 add
	System.out.println(report_no+".act="+act);
	System.out.println(report_no+".bank_type="+bank_type);
	if(firstStatus.equals("true")){//若從Menu點選時,先清空session裡的資料
	   session.setAttribute("CANCEL_NO",null);   
	   session.setAttribute("HSIEN_ID",null);   
	   session.setAttribute("BankList",null);    	
	   session.setAttribute("btnFieldList",null);	
	   session.setAttribute("SortList",null);   	   
	   session.setAttribute("SortBy",null);   
	   session.setAttribute("excelaction",null);   
	   session.setAttribute("nowbank_type",null);   
	   session.setAttribute("S_YEAR",null);   
	   session.setAttribute("S_MONTH",null);   
	   session.setAttribute("E_YEAR",null);   
	   session.setAttribute("E_MONTH",null);   
	   session.setAttribute("Unit",null);   
	   //session.setAttribute("acc_div",null);   	   
	}
    if(!Utility.CheckPermission(request,report_no)){//無權限時,導向到LoginError.jsp//95.11.03 fix 使用Utility.CheckPermission檢核權限 by 2295
        rd = application.getRequestDispatcher( LoginErrorPgName );        
    }else{            
    	//set next jsp 	
    	//將選取的縣市別,金融機構代碼,報表欄位,排序的報表欄位寫到session=======================================    
    	if(last_act.equals("ReadRpt")){//95.12.08 add讀取完報表格式檔時,回到BankList不做此判斷(金融機構代號若本來選全部->各信用部 or 各信用部->全部,清空已選報表欄位/排序欄位)
    	   if(request.getParameter("BankList")	!= null && !((String)request.getParameter("BankList")).equals("")){
    	      session.setAttribute("BankList",(String)request.getParameter("BankList"));    	
    	   }    	   
    	}else{//不為讀取格式檔時
    		//95.12.07 add 取得前一次所選取的金融機構代碼=====================================================================
    		if(session.getAttribute("BankList") != null){        	   
           		lastBank_List = (String)session.getAttribute("BankList");
        	}
        	//==============================================================================================================
    		if(request.getParameter("BankList")	!= null && !((String)request.getParameter("BankList")).equals("")){    	   
    	   		nowBank_List = (String)request.getParameter("BankList");//95.12.06 add 取得目前所選取的金融機構代碼    	   
    	   		session.setAttribute("BankList",(String)request.getParameter("BankList"));    	    	   
    		}else{
    	   		session.setAttribute("BankList",null);//95.12.07 add 目前無所選取的金融機構代碼時,則session先清成null
    		}   
    	}
    	if(request.getParameter("btnFieldList")	!= null && !((String)request.getParameter("btnFieldList")).equals("")){
    	   session.setAttribute("btnFieldList",(String)request.getParameter("btnFieldList"));	
    	}
    	if(request.getParameter("SortList")	!= null && !((String)request.getParameter("SortList")).equals("")){
    	   System.out.println(report_no+".SortList="+(String)request.getParameter("SortList"));
           session.setAttribute("SortList",(String)request.getParameter("SortList"));   
        }
        if(!last_act.equals("ReadRpt")){//95.12.08 add不為讀取完報表格式檔,判斷是否清除已選報表欄位/排序欄位    
        	//95.12.07 add 金融機構代號若本來選全部->各信用部 or 各信用部->全部,清空已選報表欄位/排序欄位============================================
        	System.out.println("nowBank_List="+nowBank_List);
        	System.out.println("lastBank_List="+lastBank_List);
        	if((nowBank_List.indexOf("ALL") != -1 && lastBank_List.indexOf("ALL") == -1 )//目前所選的Bank_List,有選全部;上次選的Bank_List.沒選全部    	   
        	|| (nowBank_List.indexOf("ALL") == -1 && lastBank_List.indexOf("ALL") != -1 ))//目前所選的Bank_List,沒選全部;上次選的Bank_List.有選全部    	   
    		{        
    	   		session.setAttribute("btnFieldList",null);//清除已勾選的勾選的報表欄位)
           		session.setAttribute("SortList",null);//清除已勾選的欲sort的報表欄位                 
    		}
    	}
        //營運中/已裁撤=============================================================
        if(request.getParameter("CANCEL_NO") != null && !((String)request.getParameter("CANCEL_NO")).equals("")){
           session.setAttribute("CANCEL_NO",(String)request.getParameter("CANCEL_NO"));   
        }
        //=======================================================================================
        if(request.getParameter("HSIEN_ID")	!= null && !((String)request.getParameter("HSIEN_ID")).equals("")){
           session.setAttribute("HSIEN_ID",(String)request.getParameter("HSIEN_ID"));   
        }
        if(request.getParameter("SortBy")	!= null && !((String)request.getParameter("SortBy")).equals("")){
           session.setAttribute("SortBy",(String)request.getParameter("SortBy"));   
        }
        if(request.getParameter("excelaction")	!= null && !((String)request.getParameter("excelaction")).equals("")){
           session.setAttribute("excelaction",(String)request.getParameter("excelaction"));   
        }     
           
        if(request.getParameter("bank_type") != null && !((String)request.getParameter("bank_type")).equals("")){
           //95.09.06 此次所選的bank_type與上次不同清除已勾選的勾選的報表欄位/sort報表欄位
           if(session.getAttribute("nowbank_type") != null && !((String)session.getAttribute("nowbank_type")).equals(bank_type)){                           
              session.setAttribute("btnFieldList",null);//清除已勾選的勾選的報表欄位)
              session.setAttribute("SortList",null);//清除已勾選的欲sort的報表欄位
           }                 
           session.setAttribute("nowbank_type",(String)request.getParameter("bank_type"));   
        }else{//request無所選bank_type
           if(session.getAttribute("nowbank_type") != null){    
              bank_type = (String)session.getAttribute("nowbank_type");
           }
        }   
        
        //95.11.14 add 儲存DS_bank_type. menu的機構類別===================================================================
        if(request.getParameter("DS_bank_type") != null && !((String)request.getParameter("DS_bank_type")).equals("")){           
           session.setAttribute("DS_bank_type",(String)request.getParameter("DS_bank_type"));  
           session.setAttribute("nowbank_type",Utility.getTrimString(dataMap.get("DS_bank_type")));//100.05.04 
        }
        if(firstStatus.equals("true") && (request.getParameter("bank_type") != null && !((String)request.getParameter("bank_type")).equals(""))){           
           session.setAttribute("DS_bank_type",(String)request.getParameter("bank_type"));              
        }
        //===============================================================================================================
           
        //金額單位
        if(request.getParameter("Unit") != null && !((String)request.getParameter("Unit")).equals("")){
           session.setAttribute("Unit",(String)request.getParameter("Unit"));   
        }
        //年-begin
        if(request.getParameter("S_YEAR") != null && !((String)request.getParameter("S_YEAR")).equals("")){
           session.setAttribute("S_YEAR",(String)request.getParameter("S_YEAR"));   
        }
        //年-end
        if(request.getParameter("E_YEAR") != null && !((String)request.getParameter("E_YEAR")).equals("")){
           session.setAttribute("E_YEAR",(String)request.getParameter("E_YEAR"));   
        }
        //95.12.05 增加年月區間
        //月-begin
        if(request.getParameter("S_MONTH") != null && !((String)request.getParameter("S_MONTH")).equals("")){
           session.setAttribute("S_MONTH",(String)request.getParameter("S_MONTH"));   
        }
        //月-end
        if(request.getParameter("E_MONTH") != null && !((String)request.getParameter("E_MONTH")).equals("")){
           session.setAttribute("E_MONTH",(String)request.getParameter("E_MONTH"));   
        }               
        //=================================================================================================
        if(act.equals("RptColumn")){//報表欄位   
        	//103.02.10 add 103/01以後,漁會套用新表格(增加/異動科目代號) 
			String strYymm="";
    		if( bank_type.equals("7") && (Integer.parseInt(Utility.getTrimString(dataMap.get("S_YEAR"))) * 100 + Integer.parseInt(Utility.getTrimString(dataMap.get("S_MONTH"))) >= 10301) ){
    			strYymm = "_10301";
    		}           
			//List FieldList = Utility.getFileData(Utility.getProperties("schemaDir")+System.getProperty("file.separator")+"A06_"+bank_type+"_master_"+acc_div+".TXT");						
			//System.out.println(Utility.getProperties("schemaDir")+System.getProperty("file.separator")+"A06_"+bank_type+"_master_"+acc_div+".TXT");
			List FieldList = Utility.getFileData(Utility.getProperties("schemaDir")+System.getProperty("file.separator")+"A06_"+bank_type+"_master"+strYymm+".TXT");						
			/*A06.農會.全部.都讀取bank_type=6的設定檔.漁會.全部.都讀取bank_type=6的設定檔
			//95.09.06 add 若金融機構代號選為全部時,則讀取bank_type=ALL的設定檔
			System.out.println(report_no+".bank_type="+bank_type);
			System.out.println((String)session.getAttribute("BankList"));
            if(!bank_type.equals("ALL") && session.getAttribute("BankList") != null && ((String)session.getAttribute("BankList")).indexOf("ALL") != -1){
              FieldList = Utility.getFileData(Utility.getProperties("schemaDir")+System.getProperty("file.separator")+"A06_ALL_master.TXT");						
            }
            */
			System.out.println(Utility.getProperties("schemaDir")+System.getProperty("file.separator")+"A06_"+bank_type+"_master.TXT");
			String FieldListString = "";			
			for(int i=0;i<FieldList.size();i++){			    			    
			    FieldListString += Utility.ISOtoUTF8((String)FieldList.get(i));
			    if(i < FieldList.size() - 1) FieldListString +=",";   
			}
			FieldListString = FieldListString.substring(1,FieldListString.length());//104.03.09 add 去掉第一個字元?
			request.setAttribute("FieldList",FieldListString); 		
			System.out.println(report_no+".FieldList="+FieldListString);	
        	rd = application.getRequestDispatcher( RptColumnPgName );        
        }else if(act.equals("BankList")){//金融機構
            System.out.println("btnFieldList="+(String)request.getParameter("btnFieldList"));            
            rd = application.getRequestDispatcher( BankListPgName );        
    	}else if(act.equals("RptOrder")){//排序欄位       	     	       	        	    
        	rd = application.getRequestDispatcher( RptOrderPgName );         
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
    	       webURL_Y = "/pages/"+report_no+".jsp?act=BankList&last_act=ReadRpt";//95.12.08       	       	    	       	   	   			       
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
    private final static String report_no = "DS006W";
    private final static String nextPgName = "/pages/ActMsg.jsp";    
    private final static String BankListPgName = "/pages/"+report_no+"_BankList.jsp";    
    private final static String RptColumnPgName = "/pages/"+report_no+"_RptColumn.jsp";        
    private final static String RptOrderPgName = "/pages/"+report_no+"_RptOrder.jsp";        
    private final static String RptStylePgName = "/pages/"+report_no+"_RptStyle.jsp";        
    private final static String RptCreatePgName = "/pages/"+report_no+"_Excel.jsp";        
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";    
%>    