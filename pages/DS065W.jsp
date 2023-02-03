<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.report.reportUtil" %>
<%@ page import="java.util.*" %>
<%@include file="./include/Header.include" %>
<%
	String last_act = Utility.getTrimString(dataMap.get("last_act"));
	String bank_type = Utility.getTrimString(dataMap.get("bank_type"));    
    String firstStatus = Utility.getTrimString(dataMap.get("firstStatus"));
    String template = Utility.getTrimString(dataMap.get("template"));
    String template_list = Utility.getTrimString(dataMap.get("template_list"));
    String menu = Utility.getTrimString(dataMap.get("menu"));//顯示BOAF menu用		    		
	if(bank_type.equals("")){
	   bank_type = (session.getAttribute("nowbank_type")==null)?"":(String)session.getAttribute("nowbank_type");
	}   
	String nowBank_List = "";
	String lastBank_List = "";
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
	   session.setAttribute("Unit",null);   
	   session.setAttribute("loan_item",null);
	   System.out.println("clear all session");
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
    	   if(!Utility.getTrimString(dataMap.get("BankList")).equals("")){    	   
    	      nowBank_List = Utility.getTrimString(dataMap.get("BankList"));//95.12.06 add 取得目前所選取的金融機構代碼    	   
    	      session.setAttribute("BankList",Utility.getTrimString(dataMap.get("BankList")));    	    	   
    	   }else{
    	      session.setAttribute("BankList",null);//95.12.07 add 目前無所選取的金融機構代碼時,則session先清成null
    	   }
    	}   
    	if(!Utility.getTrimString(dataMap.get("btnFieldList")).equals("")){
    	   session.setAttribute("btnFieldList",Utility.getTrimString(dataMap.get("btnFieldList")));	
    	}
    	if(!Utility.getTrimString(dataMap.get("SortList")).equals("")){
           session.setAttribute("SortList",Utility.getTrimString(dataMap.get("SortList")));   
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
    	//=======================================================================================================================================
        //營運中/已裁撤=============================================================
        if(!Utility.getTrimString(dataMap.get("CANCEL_NO")).equals("")){
           session.setAttribute("CANCEL_NO",(String)request.getParameter("CANCEL_NO"));   
        }
        //=======================================================================================
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
           //95.09.06 此次所選的bank_type與上次不同清除已勾選的勾選的報表欄位/sort報表欄位
           if(session.getAttribute("nowbank_type") != null && !((String)session.getAttribute("nowbank_type")).equals(bank_type)){                           
              session.setAttribute("btnFieldList",null);//清除已勾選的勾選的報表欄位)
              session.setAttribute("SortList",null);//清除已勾選的欲sort的報表欄位
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
            //session.setAttribute("nowbank_type",Utility.getTrimString(dataMap.get("DS_bank_type")));//100.05.04
         }
        
         if(firstStatus.equals("true") && (!Utility.getTrimString(dataMap.get("bank_type")).equals(""))){           
            session.setAttribute("DS_bank_type",Utility.getTrimString(dataMap.get("bank_type"))); 
         }
        //===============================================================================================================
        //報表類別
        /*if(request.getParameter("acc_div") != null && !((String)request.getParameter("acc_div")).equals("")){
           //95.08.23 此次所選的acc_div與上次不同清除已勾選的勾選的報表欄位/sort報表欄位
           if(session.getAttribute("acc_div") != null && !((String)session.getAttribute("acc_div")).equals(acc_div)){              
              session.setAttribute("btnFieldList",null);//清除已勾選的勾選的報表欄位)
              session.setAttribute("SortList",null);//清除已勾選的欲sort的報表欄位
           }   
           session.setAttribute("acc_div",(String)request.getParameter("acc_div"));              
        }else{//request無所選acc_div
           if(session.getAttribute("acc_div") != null){    
              acc_div = (String)session.getAttribute("acc_div");
           }
        }*/   
        //金額單位
        if(!Utility.getTrimString(dataMap.get("Unit")).equals("")){
           session.setAttribute("Unit",(String)request.getParameter("Unit"));   
        }
        //年-begin
        if(!Utility.getTrimString(dataMap.get("S_YEAR")).equals("")){
           session.setAttribute("S_YEAR",(String)request.getParameter("S_YEAR"));
        }
        if(!Utility.getTrimString(dataMap.get("S_MONTH")).equals("")){
           session.setAttribute("S_MONTH",(String)request.getParameter("S_MONTH"));
        }
        if(!Utility.getTrimString(dataMap.get("loan_item")).equals("")){
            session.setAttribute("loan_item",(String)request.getParameter("loan_item"));   
        }
        //=================================================================================================
        if(act.equals("RptColumn")){//報表欄位              
            List FieldList = null;        
	        StringBuffer sql = new StringBuffer();		
	        List paramList = new ArrayList();
	        DataObject bean = null;
            sql.append(" select cano,rulename from rptrule_ds where acc_tr_type=? and type=? order by to_number(output_order) ");
     		paramList.add(report_no);
     		paramList.add("0");
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
        }else if(act.equals("BankList")){//金融機構
            System.out.println("BankList="+Utility.getTrimString(dataMap.get("BankList")));
            System.out.println("btnFieldList="+Utility.getTrimString(dataMap.get("btnFieldList"))); 
            rd = application.getRequestDispatcher( BankListPgName );        
    	}else if(act.equals("RptOrder")){//排序欄位       	     	       	        	    
        	rd = application.getRequestDispatcher( RptOrderPgName );         
        }else if(act.equals("RptStyle")){//報表格式     
            List templateList = reportUtil.getTemplateList(lguser_id,bank_type,report_no);//取得範本資料//95.11.03使用reportUtil 
            request.setAttribute("templateList",templateList);
            //避免點選[儲存格式檔]後.再點選至[金融機構]頁籤.原點選的金融機構清單為空白
            if(!"".equals(lastBank_List) && "".equals(nowBank_List) && "".equals(Utility.getTrimString(dataMap.get("BankList")))){
                session.setAttribute("BankList", lastBank_List);
            }
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
    private final static String report_no = "DS065W";
    private final static String nextPgName = "/pages/ActMsg.jsp";    
    private final static String BankListPgName = "/pages/"+report_no+"_BankList.jsp";    
    private final static String RptColumnPgName = "/pages/"+report_no+"_RptColumn.jsp";  
    private final static String RptOrderPgName = "/pages/"+report_no+"_RptOrder.jsp";  
    private final static String RptStylePgName = "/pages/"+report_no+"_RptStyle.jsp";  
    private final static String RptCreatePgName = "/pages/"+report_no+"_Excel.jsp";  
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";    
%>    