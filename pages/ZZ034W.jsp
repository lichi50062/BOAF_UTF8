<%
// 95.06.06-19 create by 2295
// 95.08.16 fix 拿掉A06/A04跨表檢核
//              各逾放期數合計之科目【960500】欄位的金額，應等於A04.【840731】+ A04.【840732】+A04.【840733】+A04.【840734】+ A04.【840735】
//              的合計申報金額不跨表檢核 by 2295
// 95.09.15 add 990000.970000轉金額的用Float取代Integer by 2295
// 98.08.28 add A04.840740=A01.990000=A06.970000(逾期放款合計) by 2205
// 98.09.14 add A01A02A99跨表檢核 by 2295
// 98.09.24 add 執行全部跨表檢核,並將其執行結果,發mail通知 by 2295
// 98.09.29 add 中央存保可查詢所有農漁會信用部資料 by 2295
// 98.12.28 fix A02.990120=A01.310000+A01.320000-A01.310800 by 2295
// 99.02.10 fix 因A02.990120/A02.990130為平均值.故該公式不納入檢核 by 2295
// 99.02.22 fix A01A02A99跨表檢核-5~5均算檢核成功 by 2295 
//100.01.11 fix sqlInjection/根據查詢年度.100年以後取得新縣市別.100年以前取得舊縣市別 by 2295
//102.02.05 fix a02有增加amt_name by 2295
//102.08.22 add A01A06跨表檢核,103.01以後,漁會套用新科目代號 by 2295
//104.01.22 fix 金庫.黃小姐來電通知.不需發e-mail,已將排程暫停 by 2295
//105.03.31 add 存保檢核相對科目金額之比對 2968
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Enumeration" %>
<%@ page import="java.util.HashMap" %>
<%@ page import="java.util.Hashtable" %>
<%@ page import="java.util.StringTokenizer" %>
<%@ page import="java.util.*,java.io.*" %>
<%
	RequestDispatcher rd = null;
	String actMsg = "";	
	String alertMsg = "";			
	String webURL_Y = "";	
	String webURL_N = "";		
	boolean doProcess = false;	
	HashMap dataMap = null;
	String mail_subject="";
	String mail_content="";
	
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");				
	System.out.println("act="+act);	
	
	//取得session資料,取得成功時,才繼續往下執行===================================================
	if(session.getAttribute("muser_id") == null && !act.equals("checkALL")){//session timeout	
      System.out.println("ZZ034W login timeout");   
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
	
	
   
	//登入者資訊
	String lguser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		
	String lguser_name = ( session.getAttribute("muser_name")==null ) ? "" : (String)session.getAttribute("muser_name");		
	String lguser_type = ( session.getAttribute("muser_type")==null ) ? "" : (String)session.getAttribute("muser_type");				
	String lguser_tbank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");			
	session.setAttribute("nowtbank_no",null);//94.01.05 fix 沒有Bank_List,把所點選的Bank_no清除======
	
	String bank_type = ( session.getAttribute("bank_type")==null ) ? "" : (String)session.getAttribute("bank_type");				
	String bank_type_list = ( request.getParameter("BANK_TYPE_List")==null ) ? "" : (String)request.getParameter("BANK_TYPE_List");
	String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? "" : (String)request.getParameter("S_YEAR");				
    String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? "" : (String)request.getParameter("S_MONTH");				
    String unit = ( request.getParameter("Unit")==null ) ? "" : (String)request.getParameter("Unit");				
    String upd_code = ( request.getParameter("UPD_CODE")==null ) ? "" : (String)request.getParameter("UPD_CODE");				
    String checkoption = ( request.getParameter("CheckOption")==null ) ? "" : (String)request.getParameter("CheckOption");				
    
    System.out.println("ZZ034W.bank_type="+bank_type);	
    System.out.println("ZZ034W.bank_type_list="+bank_type_list);	    
    System.out.println("ZZ034W.S_YEAR="+S_YEAR);	
    System.out.println("ZZ034W.S_MONTH="+S_MONTH);	
    System.out.println("ZZ034W.checkoption="+checkoption);	
    System.out.println("ZZ034W.upd_code="+upd_code);
    
    String tmpbank_type="";    
    if(lguser_id.equals("A111111111") || bank_type.equals("2") || bank_type.equals("1")){
       tmpbank_type = "Z";			    
	}else{
	   tmpbank_type = bank_type;				
	}	  
	
	//申報上個月份的報表
	String YEAR  = Utility.getYear();
   	String MONTH = Utility.getMonth(); 
	
    if(!Utility.CheckPermission(request,report_no) && !act.equals("checkALL")){//無權限時,導向到LoginError.jsp
        rd = application.getRequestDispatcher( LoginErrorPgName );        
    }else{            
    	//set next jsp 	   
    	List file_list = getUpdCode("032");//取得欲比對的申報檔案
    	List upd_code_list = getUpdCode("033");//取得檢核結果   
    	request.setAttribute("file_list",file_list); 
    	request.setAttribute("upd_codeList",upd_code_list); 
    	if( act.equals("Qry")){    	  
    	   if(!S_YEAR.equals("")){
               S_YEAR = String.valueOf(Integer.parseInt(S_YEAR));
            } 
            if(!S_MONTH.equals("")){
                S_MONTH = String.valueOf(Integer.parseInt(S_MONTH));
            }      	        	    
    	    rd = application.getRequestDispatcher( ListPgName +".jsp?act=Qry");            	        	        	    	            
    	}else if( act.equals("List")){    	          	     
    	    List returnList = null;
    	    String[] title = null;
    	    if(checkoption.equals("1")){//A01A05
    	       dataMap = getQryResultA01A05(bank_type_list,S_YEAR,S_MONTH,unit,upd_code,bank_type,lguser_tbank_no,lguser_id);    	       	        	    
    	       returnList = checkResult_A01A05(dataMap,upd_code,"0");
    	       title = (String[])returnList.get(0);
    	       request.setAttribute("zz034wList",returnList.get(1));    	
    	       rd = application.getRequestDispatcher( ListPgName+ "_A01A05.jsp?act=List&BANK_TYPE_List="+bank_type_list+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&unit="+unit+"&upd_code="+upd_code+"&checkoption="+checkoption+"&tmpbank_type="+tmpbank_type);            	        	        	    	                	       
    	    }else if(checkoption.equals("2")){//A01A06   
    	       dataMap = getQryResultA01A06(bank_type_list,S_YEAR,S_MONTH,unit,upd_code,bank_type,lguser_tbank_no,lguser_id);    	       	        	    
    	       returnList = checkResult_A01A06(dataMap,upd_code,bank_type_list,"0",S_YEAR,S_MONTH);
    	       title = (String[])returnList.get(0);
    	       request.setAttribute("zz034wList",returnList.get(1));    	
    	       rd = application.getRequestDispatcher( ListPgName+ "_A01A06.jsp?act=List&BANK_TYPE_List="+bank_type_list+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&unit="+unit+"&upd_code="+upd_code+"&checkoption="+checkoption+"&tmpbank_type="+tmpbank_type);            	        	        	    	                	       
    	    }else if(checkoption.equals("4")){//A01/A02/A99
    	       dataMap = getQryResultA01A02A99(bank_type_list,S_YEAR,S_MONTH,unit,upd_code,bank_type,lguser_tbank_no,lguser_id);    	       	        	    
    	       returnList = checkResult_A01A02A99(dataMap,upd_code,"0");
    	       title = (String[])returnList.get(0);
    	       request.setAttribute("zz034wList",returnList.get(1));    	
    	       rd = application.getRequestDispatcher( ListPgName+ "_A01A02A99.jsp?act=List&BANK_TYPE_List="+bank_type_list+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&unit="+unit+"&upd_code="+upd_code+"&checkoption="+checkoption+"&tmpbank_type="+tmpbank_type);            	        	        	    	                	       
    	    }else if(checkoption.equals("5")){//存保檢核
     	       dataMap = getQryResultA10(bank_type_list,S_YEAR,S_MONTH,unit,upd_code,bank_type,lguser_tbank_no,lguser_id);    	       	        	    
     	       returnList = checkResult_A10(dataMap,upd_code,"0");
     	       title = (String[])returnList.get(0);
     	       request.setAttribute("zz034wList",returnList.get(1));    	
     	       rd = application.getRequestDispatcher( ListPgName+ "_A10.jsp?act=List&BANK_TYPE_List="+bank_type_list+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&unit="+unit+"&upd_code="+upd_code+"&checkoption="+checkoption+"&tmpbank_type="+tmpbank_type);            	        	        	    	                	       
     	    }/*else if(checkoption.equals("3")){//A04A06//95.08.16 拿掉A06/A04比對   
    	    }/*else if(checkoption.equals("3")){//A04A06//95.08.16 拿掉A06/A04比對   
    	       returnList = getQryResultA04A06(bank_type_list,S_YEAR,S_MONTH,unit,upd_code,bank_type,lguser_tbank_no,lguser_id);    	       	        	    
    	       request.setAttribute("zz034wList",returnList);    	
    	       rd = application.getRequestDispatcher( ListPgName+ "_A04A06.jsp?act=List&BANK_TYPE_List="+bank_type_list+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&unit="+unit+"&upd_code="+upd_code+"&checkoption="+checkoption+"&tmpbank_type="+tmpbank_type);            	        	        	    	                	       
    	    }*/       	        	        	    
    	    
    	    /*List dbData = null;
    	    if(checkoption.equals("1")){//A01A05
    	       dbData = getQryResultA01A05(bank_type_list,S_YEAR,S_MONTH,upd_code,bank_type,lguser_tbank_no,lguser_id);    	       	        	    
    	    }else if(checkoption.equals("2")){//A01A06   
    	       dbData = getQryResultA01A05(bank_type_list,S_YEAR,S_MONTH,upd_code,bank_type,lguser_tbank_no,lguser_id);    	       	        	    
    	    }    	    
    	    String[] title = getTitleIdx(dbData);    	    
    	    */
    	    
    	    request.setAttribute("titleLength",title);    	        	     	        	        	         	        	        	    
        }else if( act.equals("checkALL")){
            String mail_text = "";
            String mail_data = "";
            DataObject bean = null;
            List returnList = null;
            List mail_data_list = null;
            List fileCheckData = null;
            HashMap dataMap_A01A05 = null;
            HashMap dataMap_A01A06_6 = null;
            HashMap dataMap_A01A06_7 = null;
            HashMap dataMap_A01A02A99 = null;
            HashMap dataMap_A10 = null;
            
            bank_type_list="ALL";//A01/A05=ALL;A01/A06=6,7;A01/A02/A99=ALL;A10=ALL            
            unit="1";
            upd_code="";
            lguser_id="A111111111";            
            if(S_YEAR.equals("") && S_MONTH.equals("")){
      		   S_YEAR = YEAR;
      		   S_MONTH = MONTH;
   			}
            dataMap_A01A05 = getQryResultA01A05(bank_type_list,S_YEAR,S_MONTH,unit,upd_code,bank_type,lguser_tbank_no,lguser_id);
            dataMap_A01A06_6 = getQryResultA01A06("6",S_YEAR,S_MONTH,unit,upd_code,bank_type,lguser_tbank_no,lguser_id);    	       	        	        	       	        	    
            dataMap_A01A06_7 = getQryResultA01A06("7",S_YEAR,S_MONTH,unit,upd_code,bank_type,lguser_tbank_no,lguser_id);    	       	        	        	       	        	    
            dataMap_A01A02A99 = getQryResultA01A02A99(bank_type_list,S_YEAR,S_MONTH,unit,upd_code,bank_type,lguser_tbank_no,lguser_id);
            dataMap_A10 = getQryResultA10(bank_type_list,S_YEAR,S_MONTH,unit,upd_code,bank_type,lguser_tbank_no,lguser_id);
            fileCheckData = getFileCheckData();
            for(int i=0;i<fileCheckData.size();i++){
                bean = (DataObject)fileCheckData.get(i);
                mail_text = "";
                if(((String)bean.getValue("cmuse_id")).equals("1")){//A01/A05相對科目金額之比對
    	             returnList = checkResult_A01A05(dataMap_A01A05,(String)bean.getValue("input_order"),"1");
    				 if(returnList != null) mail_text = (String)returnList.get(0);
    	    		 System.out.println("A01/A05.mail_text="+mail_text);
    	    	}else if(((String)bean.getValue("cmuse_id")).equals("2")){//A01/A06相對科目金額之比對	    	    	     
    	             returnList = checkResult_A01A06(dataMap_A01A06_6,(String)bean.getValue("input_order"),"6","1",S_YEAR,S_MONTH);//A01/A06有區分.農漁會科目代號
    	             if(returnList != null) mail_text += "農會-"+(String)returnList.get(0)+"\n                     ";
    	             returnList = checkResult_A01A06(dataMap_A01A06_7,(String)bean.getValue("input_order"),"7","1",S_YEAR,S_MONTH);//A01/A06有區分.農漁會科目代號
    	             if(returnList != null) mail_text += "漁會-"+(String)returnList.get(0);
    	    		 System.out.println("A01/A06.mail_text="+mail_text);
    	        }else if(((String)bean.getValue("cmuse_id")).equals("4")){//A01/A02/A99相對科目金額之比對			                                         
    	             returnList = checkResult_A01A02A99(dataMap_A01A02A99,(String)bean.getValue("input_order"),"1");
    	             if(returnList != null) mail_text = (String)returnList.get(0);
    	    		 System.out.println("A01A02A99.mail_text="+mail_text);
    	        }else if(((String)bean.getValue("cmuse_id")).equals("5")){//A10存保檢核相對科目金額之比對	                                         
	   	             returnList = checkResult_A10(dataMap_A10,(String)bean.getValue("input_order"),"1");
	   	             if(returnList != null) mail_text = (String)returnList.get(0);
	   	    		 System.out.println("A10.mail_text="+mail_text);
    	    	}
    	    	if(mail_text.equals("")) continue;
    	    	mail_subject = getMailSubject(S_YEAR,S_MONTH,(String)bean.getValue("cmuse_name"));    	    	
    	    	mail_content = getMailContent(S_YEAR,S_MONTH,(String)bean.getValue("cmuse_name"),mail_text);    	    	
    	    	mail_data_list = getMail((String)bean.getValue("cmuse_id"),(String)bean.getValue("input_order"));//getMail(String kind,String subkind)
    	    	mail_data = "";
    	    	if(mail_data_list != null){
    	    	   for(int j=0;j<mail_data_list.size();j++){
    	    	       bean = (DataObject)mail_data_list.get(j);
    	    	       mail_data +=(String)bean.getValue("m_email");
    	    	       if(mail_data.length() > 0 && j < (mail_data_list.size()-1)) mail_data += ",";
    	    	   }
    	    	} 
    	    	System.out.println("mail_data="+mail_data);
    	    	//檢核結果,發e-mail
       	        Utility.sendMail(mail_data,mail_subject,mail_content);
    	    
    	    	
    	    	/*
    	    	returnList = checkResult_A01A02A99(dataMap,"1",bank_type_list,"1");//金額不符
    	        if(returnList != null) mail_text += (String)returnList.get(0);
    	        System.out.println("金額不符.mail_text="+mail_text);	 
    	    	returnList = checkResult_A01A02A99(dataMap,"2",bank_type_list,"1");//金額為0
    	        if(returnList != null) mail_text += (String)returnList.get(0);
    	        System.out.println("金額為0.mail_text="+mail_text);
    	        returnList = checkResult_A01A02A99(dataMap,"3",bank_type_list,"1");//未申報
    	        if(returnList != null) mail_text += (String)returnList.get(0);
    	        System.out.println("未申報.mail_text="+mail_text);
    	    	*/
    	    		
    	    
    	    }
    	    
    	       
        }
       
    	request.setAttribute("actMsg",actMsg);
    	request.setAttribute("alertMsg",alertMsg);
    	request.setAttribute("webURL_Y",webURL_Y);    
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
    private final static String report_no = "ZZ034W";
    private final static String nextPgName = "/pages/ActMsg.jsp";        
    private final static String ListPgName = "/pages/ZZ034W_List";        
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    
    private List getUpdCode(String cmuse_div){//取得欲比對的申報檔案/檢核結果    	
          List paramList = new ArrayList();
          paramList.add(cmuse_div);
    	  List dbData = DBManager.QueryDB_SQLParam( "select * from cdshareno where cmuse_div=?",paramList,""); 
          return dbData;
    }           
    private List getFileCheckData(){   
     	String sqlCmd = " select cd2.cmuse_id,cd2.cmuse_name,cd1.input_order,cd1.cmuse_name as sub_name"
					  + " from " 
					  + " (select * from cdshareno  where cmuse_div='033')cd1 left join "
				      + " (select * from cdshareno  where cmuse_div='032')cd2 on cd1.identify_no= cd2.cmuse_id "
				      //+ " where cd2.cmuse_id='4'"
					  + " order by cd1.output_order ";
		List dbData = DBManager.QueryDB_SQLParam( sqlCmd,null,""); 
        return dbData;
    }      
    private List getMail(String kind,String subkind){//取得欲傳送的e-mail
        List paramList = new ArrayList();
        paramList.add(kind);
        paramList.add(subkind);
         String sqlCmd = "select sendmail.muser_id,m_email,kind,subkind"
				       + " from sendmail left join muser_data on sendmail.muser_id=muser_data.muser_id"
				       + " where kind=? and subkind=?";
		 List dbData = DBManager.QueryDB_SQLParam( sqlCmd,paramList,""); 
         return dbData;
    }
    private String getMailSubject(String S_YEAR,String S_MONTH,String checkFile){
         String mail_subject = "";
         mail_subject = S_YEAR + "年" + S_MONTH + "月-申報資料跨表檢核(ZZ034W)-"+checkFile+"-檢核結果通知";
         return mail_subject;
    }     
    
    private String getMailContent(String S_YEAR,String S_MONTH,String checkFile,String checkResult){
         System.out.println("date1="+Utility.getDateFormat("yyyy/MM/dd"));
          System.out.println("date2="+Utility.getCHTdate(Utility.getDateFormat("yyyy/MM/dd"),1));
         String mail_content = "";
       	 mail_content = "比對的申報檔案："+checkFile+"\n";
       	 mail_content += "檢核結果："+checkResult+ "\n";	       	
	     mail_content += "申報年月：" + S_YEAR + "年" + S_MONTH + "月\n";
	     mail_content += "檢核日期：" + Utility.getCHTdate(Utility.getDateFormat("yyyy/MM/dd"),1) + "\n\n";
	     mail_content += "煩請至網際網路申報系統查詢其詳細檢核結果!!";
	     return mail_content;
    }     			    
    //取得查詢結果
    /*private List getQryResultA01A05(String bank_type_list,String S_YEAR,String S_MONTH,String upd_code,String bank_type,String tbank_no,String lguser_id){    	   
    		//查詢條件        		
    		String sqlCmd = "";
    		String ruleA01 = "";
    		String ruleA05 = "";
    		String condition = "";     		  		
    		String tmpbank_type="";    		    		
    		String tmpbank_no="";    		
    		if(lguser_id.equals("A111111111") || bank_type.equals("2") || bank_type.equals("1")){
    		    tmpbank_type = "Z";
			    tmpbank_no = "9999999";
		    }else{
		        tmpbank_type = bank_type;
				tmpbank_no = tbank_no;
			}	
			
            //select a01.s_report_name,
			//	   a01.bank_no as a01_bank_no,a05.bank_no as a05_bank_no,
	   		//	   a01.acc_code as a01_acc_code,a05.acc_code as a05_acc_code,
	   		//	   a01.acc_range as a01_acc_range,a05.acc_range as a05_acc_range,
	   		//	   a01.amt as a01amt,a05.amt as a05amt
		 	//from 
		    //    (
			//	select * from 
			//	(
 			//		select NVL((BN01.BANK_NO || BN01.BANK_NAME), ' ')  as S_Report_Name,  
    		//			   BN01.BANK_NO, BN01.BANK_NAME, bn01.bank_type ,  
  			//			   nvl(wlx01.CENTER_NO,' ') as CENTER_NO, 
			//			   nvl(wlx01.M2_NAME,' ')   as M2_NAME, 
			//			   wlx01.CANCEL_no, wlx01.CANCEL_date,
			//			   nvl(cd01.hsien_id,' ') as  hsien_id,  
	    	//			   a01.acc_code,decode(A01.acc_code,'320100',A01_320100.amt,a01.amt) as amt,
	    	//			   a01.acc_range,
	    	//			   nvl(cd01.hsien_name,'OTHER')  as  hsien_name, 
			//			   cd01.FR001W_output_order     as  FR001W_output_order 
 			//	   from  (select * from cd01 where  cd01.hsien_id <> 'Y') cd01 
     		//			  left join wlx01 on wlx01.hsien_id=cd01.hsien_id  
	 		//			  left join (select BANK_NO, BANK_NAME,  bank_type  
	    	//						 from bn01 where bn01.bank_type in('6','7'))  bn01  
			//					on wlx01.bank_no=bn01.bank_no  
 			//			  left join 
 	    	//						(select  bank_code,m_year,m_month,a01.acc_code,amt,
		 	//								 decode(acc_code,'310100','01','310200','02','310300','03','310400','04','310500','05','310600','06','310800','07','320100','08','130200','09') as acc_range 
		 	//						 from  A01 
		 	//						 where M_Year =  94  and M_month =  8
		 	//						 and  a01.acc_code in('310100','310200','310300','310400','310500','310600','310800','320100','130200')		 
		 	//						 order by a01.bank_code,acc_range	
        	//						) A01
  	  		//					on bn01.bank_no = A01.Bank_Code	  	 
 			//			  left join
 	    	//						(select m_year,m_month,bank_code,sum(amt) as amt 
		 	//						 from (select  * FROM  A01 
         	//   							   where M_Year =  94  and M_month =  8
			//   							   and  acc_code in('320100','320200')) A01_320100_320200
		 	//						 group by m_year,m_month,bank_code	
		 	//						 order by m_year,m_month,bank_code		  
		 	//						)A01_320100	
 	  		//					on bn01.bank_no = A01_320100.Bank_Code
			//	)Temp_Output  where BANK_NO <> ' '
			//order by  Bank_Type, FR001W_output_order, BANK_NO,acc_range) A01,
			//(
			//	select * from 
			//	(
 			//	select NVL((BN01.BANK_NO || BN01.BANK_NAME), ' ')  as S_Report_Name,  
    		//		   BN01.BANK_NO, BN01.BANK_NAME, bn01.bank_type ,  
  			//		   nvl(wlx01.CENTER_NO,' ') as CENTER_NO, 
			//		   nvl(wlx01.M2_NAME,' ')   as M2_NAME, 
			//		   wlx01.CANCEL_no, wlx01.CANCEL_date,
			//		   nvl(cd01.hsien_id,' ') as  hsien_id,  
	   		//		   a05.acc_code,decode(A05.acc_code,'910401',A05_910401.amt,a05.amt) as amt,
	   		//		   a05.acc_range,
	    	//		   nvl(cd01.hsien_name,'OTHER')  as  hsien_name, 
			//		   cd01.FR001W_output_order     as  FR001W_output_order 
 			//	from  (select * from cd01 where  cd01.hsien_id <> 'Y') cd01 
     		//		   left join wlx01 on wlx01.hsien_id=cd01.hsien_id  
	 		//		   left join (select BANK_NO, BANK_NAME,  bank_type  
	    	//					  from bn01 where bn01.bank_type in('6','7'))  bn01  
			//				on wlx01.bank_no=bn01.bank_no  
 			//		   left join 
 	    	//					(select  bank_code,m_year,m_month,a05.acc_code,amt,
		 	//					 decode(acc_code,'910101','01','910102','02','910103','03','910104','04','910105','05','910106','06','910107','07','910108','08','910401','09') as acc_range
		 	//					 from  A05 		 
		 	//					 where M_Year =  94  and M_month =  8
		 	//					 and  a05.acc_code in('910101','910102','910103','910104','910105','910106','910107','910108','910401')		 
		 	//					 order by a05.bank_code,acc_range	
        	//					) A05
  	  		//				on bn01.bank_no = A05.Bank_Code	  	 
 			//		   left join
 	    	//					(   select m_year,m_month,bank_code,sum(amt) as amt 
		 	//					    from (select  * FROM  A05 
         	//   						   where M_Year =  94  and M_month =  8
			//   						   and  acc_code in('910401','910402','910403','910404')) A05_910401_910404
		 	//					    group by m_year,m_month,bank_code
		 	//					    order by m_year,m_month,bank_code		  
		 	//				    )    A05_910401
	 		//				on bn   01.bank_no = A05_910401.Bank_Code		
			//    )Temp_Output  whe   re BANK_NO <> ' '		  
			//	order by  Bank_Type, FR001W_output_order, BANK_NO,acc_range) A05
			//where a01.bank_no = a05.bank_no
			//and (a01.acc_range = a05.acc_range)
			//and ((a01.amt is not null and a05.amt is not null) and (a01.amt <> a05.amt)) --A01/A05跨表檢核金額不符
			//and ((a01.amt is not null and a05.amt is not null) and (a01.amt = 0 and  a05.amt = 0)) --A01/A05跨表檢核金額均為0
			//and (a01.amt is null or a05.amt is null) -- A01/A05跨表檢核尚未申報
			//										    不能加 and (a01.acc_range = a05.acc_range)
            //            			                    因為a01 or a05 acc_range is null
			//group by a01.s_report_name,a01.bank_no,a05.bank_no,a01.acc_code,a01.acc_range,a01.amt,a05.acc_code,a05.acc_range,a05.amt
			//order by a01.s_report_name,a01.bank_no,a01.acc_range,a05.acc_range
            //========================================================================================================================================
			
			
			ruleA01 = " select * from "
              	   + " ( "
               	   + "	select NVL((BN01.BANK_NO || BN01.BANK_NAME), ' ')  as S_Report_Name," 
                   + "		   BN01.BANK_NO, BN01.BANK_NAME, bn01.bank_type ," 
                   + "		   nvl(wlx01.CENTER_NO,' ') as CENTER_NO,"
              	   + "	   	   nvl(wlx01.M2_NAME,' ')   as M2_NAME, "
              	   + "		   wlx01.CANCEL_no, wlx01.CANCEL_date,"
              	   + "	   	   nvl(cd01.hsien_id,' ') as  hsien_id,"
              	   + "		   a01.acc_code,decode(A01.acc_code,'320100',A01_320100.amt,a01.amt) as amt,"
              	   + "		   a01.acc_range,"
              	   + "		   nvl(cd01.hsien_name,'OTHER')  as  hsien_name,"
              	   + "	  	   cd01.FR001W_output_order     as  FR001W_output_order"
               	   + "	from  (select * from cd01 where  cd01.hsien_id <> 'Y') cd01 "
                   + "	left join wlx01 on wlx01.hsien_id=cd01.hsien_id " 
              	   + "	left join (select BANK_NO, BANK_NAME,  bank_type " 
              	   + "			   from bn01 where bn01.bank_type in('6','7'))  bn01 " 
              	   + "		 on wlx01.bank_no=bn01.bank_no"
               	   + "	left join "
               	   + "			 (select  bank_code,m_year,m_month,a01.acc_code,amt,"
              	   + "			  decode(acc_code,'310100','01','310200','02','310300','03','310400','04','310500','05',"
              	   + "					 '310600','06','310800','07','320100','08','130200','09') as acc_range "
              	   + "			  from  A01"
              	   + "			  where M_Year = "+S_YEAR + " and M_month = "+String.valueOf(Integer.parseInt(S_MONTH))
              	   + "	 		  and   a01.acc_code in('310100','310200','310300','310400','310500','310600','310800','320100','130200')"		 
              	   + "	  	 	  order by a01.bank_code,acc_range"	
                   + "		     ) A01 "
                   + "		 on bn01.bank_no = A01.Bank_Code"	  	 
               	   + " left join "
               	   + "		    (select m_year,m_month,bank_code,sum(amt) as amt "
              	   + "			 from (select  * FROM  A01 "
                   + "		     where M_Year = "+S_YEAR + " and M_month = "+String.valueOf(Integer.parseInt(S_MONTH))
              	   + "		     and  acc_code in('320100','320200')) A01_320100_320200 "
              	   + "			 group by m_year,m_month,bank_code"
              	   + "			 order by m_year,m_month,bank_code"		  
              	   + "			)A01_320100"
               	   + "		on bn01.bank_no = A01_320100.Bank_Code"
               	   + " ) Temp_Output "
               	   + " where BANK_NO <> ' ' and "
               	   + " ( "
				   + " ('"+tmpbank_type+"' = 'Z') or "
			       + " (('"+tmpbank_type+"' = '6' or '"+tmpbank_type+"' = '7') and  '"+tmpbank_no+"' = BANK_NO) or "
			       + " ('"+tmpbank_type+"' = '8' and '"+tmpbank_no+"' = CENTER_NO) or"
			       + " ('"+tmpbank_type+"' = 'B' and '"+tmpbank_no+"' = M2_NAME) "
				   + " )    and "
				   + " ( "
				   + " ('"+bank_type_list+"'= 'ALL') or "
 			       + " (('"+bank_type_list+"'= '6') and (Bank_Type = '6')) or "
				   + " (('"+bank_type_list+"'= '7') and (Bank_Type = '7')) "
				   + " )"                                                                       
				   + " order by  Bank_Type, FR001W_output_order, BANK_NO,acc_range";              	   
            ruleA05 = " select * from "
              	   + " ("
               	   + "	select NVL((BN01.BANK_NO || BN01.BANK_NAME), ' ')  as S_Report_Name,"  
                   + "		   BN01.BANK_NO, BN01.BANK_NAME, bn01.bank_type ,"  
                   + "			nvl(wlx01.CENTER_NO,' ') as CENTER_NO," 
              	   + "          nvl(wlx01.M2_NAME,' ')   as M2_NAME, "
              	   + "          wlx01.CANCEL_no, wlx01.CANCEL_date,"
              	   + "			nvl(cd01.hsien_id,' ') as  hsien_id,"
              	   + "          a05.acc_code,decode(A05.acc_code,'910401',A05_910401.amt,a05.amt) as amt,"
              	   + "          a05.acc_range,"
              	   + "			nvl(cd01.hsien_name,'OTHER')  as  hsien_name," 
              	   + "          cd01.FR001W_output_order     as  FR001W_output_order "
               	   + " from  (select * from cd01 where  cd01.hsien_id <> 'Y') cd01 "
                   + " left join wlx01 on wlx01.hsien_id=cd01.hsien_id " 
              	   + " left join (select BANK_NO, BANK_NAME,  bank_type " 
              	   + "			  from bn01 where bn01.bank_type in('6','7'))  bn01 " 
              	   + "		on wlx01.bank_no=bn01.bank_no " 
               	   + " left join "
               	   + "		(select bank_code,m_year,m_month,a05.acc_code,amt,"
              	   + "				decode(acc_code,'910101','01','910102','02','910103','03','910104','04',"
              	   + "				'910105','05','910106','06','910107','07','910108','08','910401','09') as acc_range "
              	   + "		 from  A05" 		 
              	   + "		 where M_Year = "+S_YEAR + " and M_month = "+String.valueOf(Integer.parseInt(S_MONTH))              		 
              	   + "		 and  a05.acc_code in('910101','910102','910103','910104','910105','910106','910107','910108','910401')"		 
              	   + "		 order by a05.bank_code,acc_range"	
                   + " 		 ) A05 "
                   + "       on bn01.bank_no = A05.Bank_Code "	  	 
               	   + " left join "
               	   + "		(select m_year,m_month,bank_code,sum(amt) as amt "
              	   + "		 from (select  * FROM  A05 "
              	   + "		       where M_Year = "+S_YEAR + " and M_month = "+String.valueOf(Integer.parseInt(S_MONTH))              		                        
              	   + "	           and  acc_code in('910401','910402','910403','910404')) A05_910401_910404"
              	   + "		       group by m_year,m_month,bank_code"
              	   + "		       order by m_year,m_month,bank_code"		  
              	   + "			   ) A05_910401"
              	   + "		on bn01.bank_no = A05_910401.Bank_Code"		
              	   + " )Temp_Output "
              	   + " where BANK_NO <> ' ' and "	
              	   + " ( "
				   + " ('"+tmpbank_type+"' = 'Z') or "
			       + " (('"+tmpbank_type+"' = '6' or '"+tmpbank_type+"' = '7') and  '"+tmpbank_no+"' = BANK_NO) or "
			       + " ('"+tmpbank_type+"' = '8' and '"+tmpbank_no+"' = CENTER_NO) or"
			       + " ('"+tmpbank_type+"' = 'B' and '"+tmpbank_no+"' = M2_NAME) "
				   + " )    and "
				   + " ( "
				   + " ('"+bank_type_list+"'= 'ALL') or "
 			       + " (('"+bank_type_list+"'= '6') and (Bank_Type = '6')) or "
				   + " (('"+bank_type_list+"'= '7') and (Bank_Type = '7')) "
				   + " )"                                                                        	  
				   + " order by  Bank_Type, FR001W_output_order, BANK_NO,acc_range "; 
             
			condition = " where a01.bank_no = a05.bank_no ";
			if(upd_code.equals("0") || upd_code.equals("1")){//因為未申報時,a01 or a05 acc_range is null   
			   condition += " and (a01.acc_range = a05.acc_range)";
			}		 
			if(upd_code.equals("0")){//A01/A05跨表檢核金額不符
			   condition += " and ((a01.amt is not null and a05.amt is not null) and (a01.amt <> a05.amt)) ";
		    }else if(upd_code.equals("1")){//A01/A05跨表檢核金額均為0
			   condition += " and ((a01.amt is not null and a05.amt is not null) and (a01.amt = 0 and  a05.amt = 0)) ";
			}else if(upd_code.equals("2")){//A01/A05跨表檢核尚未申報
			   condition += " and (a01.amt is null or a05.amt is null) ";			
			}			
			sqlCmd = " select a01.s_report_name,"
				   + " 		  a01.bank_no as a01_bank_no,a05.bank_no as a05_bank_no,"
	   			   + " 		  a01.acc_code as a01_acc_code,a05.acc_code as a05_acc_code,"
	   			   + " 		  a01.acc_range as a01_acc_range,a05.acc_range as a05_acc_range,"
	   			   + " 		  a01.amt as a01amt,a05.amt as a05amt"
		 		   + " from (" + ruleA01 + ")A01,("+ruleA05+")A05"
		 		   + condition 
		           + " group by a01.s_report_name,a01.bank_no,a05.bank_no,a01.acc_code,a01.acc_range,a01.amt,a05.acc_code,a05.acc_range,a05.amt"
		    	   + " order by a01.s_report_name,a01.bank_no,a01.acc_range,a05.acc_range";

			List dbData = DBManager.QueryDB( sqlCmd,"a01amt,a05amt");   
			
            return dbData;
    }
    
    private String[] getTitleidx(List dbData){
        String[] titleLength={"0","0","0","0","0","0","0","0","0"};//欲顯示的檢核有誤title idx
        String[] acc_code_A01 = {"310100","310200","310300","310400","310500","310600","310800","320100","130200"};
        String[] acc_code_A05 = {"910101","910102","910103","910104","910105","910106","910107","910108","910401"};            	
    	String acc_code="";
    	for(int idx=0;idx<dbData.size();idx++){    	    
    	    if(((DataObject)dbData.get(idx)).getValue("a01_acc_code") != null){
    	        acc_code = (String)((DataObject)dbData.get(idx)).getValue("a01_acc_code");  
    	        for(int a01idx=0;a01idx<acc_code_A01.length;a01idx++){
    	            if(acc_code.equals(acc_code_A01[a01idx])){
    	               titleLength[a01idx]="1";//要顯示的title index    	        
    	            }   
    	        }
    	    }else if(((DataObject)dbData.get(idx)).getValue("a05_acc_code") != null){   
    	        acc_code = (String)((DataObject)dbData.get(idx)).getValue("a05_acc_code");  
    	        for(int a05idx=0;a05idx<acc_code_A05.length;a05idx++){
    	            if(acc_code.equals(acc_code_A05[a05idx])){
    	               titleLength[a05idx]="1";//要顯示的title index    	        
    	            }   
    	        }
    	    }
    	}
    	return titleLength;
    }
    */
    private List checkResult_A01A05(HashMap dataMap,String upd_code,String checktype){
        String[] titleLength={"0","0","0","0","0","0","0","0","0"};//欲顯示的檢核有誤title idx
        String[] acc_code_A01 = {"310100","310200","310300","310400","310500","310600","310800","320100","130200"};
        String[] acc_code_A05 = {"910101","910102","910103","910104","910105","910106","910107","910108","910401"};
        String[] title = {"事業資金","事業公積","法定公積","特別公積","捐贈公積","資產公積","統一農(漁)貸公積","累積盈虧(A05加計「上期損益」","聯營出資(投資)"};
        String[][] Amt = new String[9][3];//[n][0]->A01.amt [n][1]->A05.amt [n][2]->showflag           		      
    	List A01Data = (List)dataMap.get("A01");
    	List A05Data = (List)dataMap.get("A05");
    	List showData = new LinkedList();//回傳的檢核有誤data
    	List showData_detail = new LinkedList();//檢核有誤的each row data
    	Properties A01 = new Properties();		
    	Properties A05 = new Properties();		    	
    	boolean addflag=false; 
    	List checktype_List = new LinkedList(); 
    	List upd_code_list = null;   	
    	List paramList = new ArrayList();
    	for(int j=0;j<8;j++){
    	    Amt[j][0]="";
    	    Amt[j][0]="";
    	    Amt[j][2]="0";
    	}   
    	for(int a01idx=0;a01idx<A01Data.size();a01idx++){
    	    System.out.println("a01idx="+a01idx);
    	    System.out.println("A01.bank_no="+(String)((List)A01Data.get(a01idx)).get(0));
    	    System.out.println("A01.s_report_name="+(String)((List)A01Data.get(a01idx)).get(1));
    	    System.out.println("A05.bank_no="+(String)((List)A05Data.get(a01idx)).get(0));
    	    System.out.println("A05.s_report_name="+(String)((List)A05Data.get(a01idx)).get(1));
    	    addflag=false;
    	    if(((String)((List)A01Data.get(a01idx)).get(0)).equals((String)((List)A05Data.get(a01idx)).get(0))){//bank_no 一樣    	           
    	        A01 = (Properties)((List)A01Data.get(a01idx)).get(2);//A01.amt集合
    	        A05 = (Properties)((List)A05Data.get(a01idx)).get(2);//A05.amt集合
    	        for(int i=0;i<acc_code_A01.length;i++){    	            
    	            if(A01.getProperty(acc_code_A01[i]) != null && A05.getProperty(acc_code_A05[i]) != null){
    	               //System.out.println("get"+acc_code_A01[i]+"="+A01.getProperty(acc_code_A01[i]));
    	               //System.out.println("get"+acc_code_A05[i]+"="+A05.getProperty(acc_code_A05[i]));
    	               Amt[i][0]=A01.getProperty(acc_code_A01[i]);//A01.amt
    	               Amt[i][1]=A05.getProperty(acc_code_A05[i]);//A05.amt    	                 	                  
    	               if(upd_code.equals("1") && (!A01.getProperty(acc_code_A01[i]).equals(A05.getProperty(acc_code_A05[i])))){//A01/A05跨表檢核金額不符
    	                  System.out.println("titleLength["+i+"]set=1");
    	                  titleLength[i]="1";//要顯示的title index    	                      
    	                  Amt[i][2]="1";
    	                  System.out.println("Amt["+i+"][0]='"+Amt[i][0]+"':"+"Amt["+i+"][1]='"+Amt[i][1]+"':Amt["+i+"][2]='"+Amt[i][2]+"'");    	               
    	                  addflag=true;
    	                  if(checktype.equals("1")){//98.09.23 add 排程檢核,發e-mail
    	                     paramList = new ArrayList();
    	                     paramList.add(upd_code);    	                  
    	                  	 upd_code_list = DBManager.QueryDB_SQLParam( "select * from cdshareno where cmuse_div='033' and identify_no='1' and input_order=?",paramList,"");     	
    	                  	 if(upd_code_list != null) checktype_List.add((String)((DataObject)upd_code_list.get(0)).getValue("cmuse_name"));
    	                  	 return checktype_List;
    	                  }
    	               }else if(upd_code.equals("2") && (A01.getProperty(acc_code_A01[i]).equals("0") && A05.getProperty(acc_code_A05[i]).equals("0"))){//A01/A05跨表檢核金額均為0
    	                  System.out.println("titleLength["+i+"]set=1");
    	                  titleLength[i]="1";//要顯示的title index    	                      
    	                  Amt[i][2]="1";
    	                  addflag=true;
    	                  System.out.println("Amt["+i+"][0]='"+Amt[i][0]+"':"+"Amt["+i+"][1]='"+Amt[i][1]+"':Amt["+i+"][2]='"+Amt[i][2]+"'");    	               
    	                  if(checktype.equals("1")){//98.09.23 add 排程檢核,發e-mail
    	                     paramList = new ArrayList();
     	                     paramList.add(upd_code);    
    	                  	 upd_code_list = DBManager.QueryDB_SQLParam( "select * from cdshareno where cmuse_div='033' and identify_no='1' and input_order=?",paramList,"");     	
    	                  	 if(upd_code_list != null) checktype_List.add((String)((DataObject)upd_code_list.get(0)).getValue("cmuse_name"));
    	                  	 return checktype_List;
    	                  }
    	               }
    	            }else{//a01 is null or a05 is null
    	               Amt[i][0]=A01.getProperty(acc_code_A01[i]);//A01.amt
    	               Amt[i][1]=A05.getProperty(acc_code_A05[i]);//A05.amt    	                 	                      	                   	            
    	               if(upd_code.equals("3")){//A01/A05跨表檢核尚未申報
    	                  System.out.println("titleLength["+i+"]set=1");
    	                  titleLength[i]="1";//要顯示的title index    	                      
    	                  Amt[i][2]="1";
    	                  addflag=true;
    	                  System.out.println("Amt["+i+"][0]='"+Amt[i][0]+"':"+"Amt["+i+"][1]='"+Amt[i][1]+"':Amt["+i+"][2]='"+Amt[i][2]+"'");    	               
    	                  if(checktype.equals("1")){//98.09.23 add 排程檢核,發e-mail
    	                      paramList = new ArrayList();
      	                     paramList.add(upd_code);
    	                  	 upd_code_list = DBManager.QueryDB_SQLParam( "select * from cdshareno where cmuse_div='033' and identify_no='1' and input_order=?",paramList,"");     	
    	                  	 if(upd_code_list != null) checktype_List.add((String)((DataObject)upd_code_list.get(0)).getValue("cmuse_name"));
    	                  	 return checktype_List;
    	                  }
    	               }
    	            }
    	            
    	        }//end of acc_code_A01
    	        if(addflag){
    	           addflag=false;
    	           System.out.println("add "+(String)((List)A01Data.get(a01idx)).get(1));
    	           showData_detail.add((String)((List)A01Data.get(a01idx)).get(1));//s_report_name    
    	           showData_detail.add(Amt);
    	           showData.add(showData_detail);
    	           showData_detail = new LinkedList();
    	        }  
    	             	           
    	        //for(int j=0;j<8;j++){
    	        //    System.out.println("'"+Amt[j][0]+"':'"+Amt[j][1]+"'");    	              
    	        //}    	          
    	        Amt = new String[9][3]; 
    	        for(int j=0;j<8;j++){
    	            Amt[j][0]="";
    	            Amt[j][0]="";
    	            Amt[j][2]="0";
    	        }   
    	    }//end of a01.bank_no=a05.bank_no
    	}//a01 loop 
    	if(checktype.equals("1")){
           return null;
        }else{
    	   List returnList = new LinkedList();
    	   returnList.add(titleLength);
    	   returnList.add(showData);
    	   return returnList;
    	}
    }
    //98.08.28 add A04.840740=A01.990000=A06.970000(逾期放款合計) 
    //102.08.22 add 103.01以後,漁會套用新科目代號
    private List checkResult_A01A06(HashMap dataMap,String upd_code,String bank_type,String checktype,String S_YEAR,String S_MONTH){
        String[] titleLength={"0","0","0","0","0","0","0","0","0","0","0","0","0","0","0"};//欲顯示的檢核有誤title idx        
        String[] acc_code_A01_6 = {"990000","120101","120102","120301","120302","120700","120401","120402","120501","120502","120601","120602","120603","120604","120600"};		          
        String[] acc_code_A01_7 = {"990000","120101","120102","120401","120402","120700","120201","120202","120501","120502","120601","120602","120603","120604","120600"};		 
        String[] acc_code_A06_6 = {"970000","120101","120102","120301","120302","120700","120401","120402","120501","120502","120601","120602","120603","120604","120600"};		  
        String[] acc_code_A06_7 = {"970000","120101","120102","120401","120402","120700","120201","120202","120501","120502","120601","120602","120603","120604","120600"};		 
        String[] acc_code_A04_67 = {"840740"};//98.08.28 add
        String[] acc_code_A01 = null;
        String[] acc_code_A06 = null;
        String[] acc_code_A04 = null;//98.08.28 add
        String[] title = {"逾期放款合計","一般放款(無擔保放款)","一般放款(擔保放款)","無擔保透支","擔保透支","內部融資(透支)","統一農貸(無擔保)","統一農貸(擔保)","專案放款(無擔保)","專案放款(擔保)","農發基金-農建放款","農發基金-農機放款","農發基金-購地放款","農發基金-農宅放款","農發基金放款小計"};
        String[][] Amt = new String[15][4];//[n][0]->A01.amt [n][1]->A06.amt [n][2]-->A04.amt [n][3]->showflag           		      
    	List A01Data = (List)dataMap.get("A01");
    	List A06Data = (List)dataMap.get("A06");
    	List A04Data = (List)dataMap.get("A04");//98.08.28 add
    	List showData = new LinkedList();//回傳的檢核有誤data
    	List showData_detail = new LinkedList();//檢核有誤的each row data
    	Properties A01 = new Properties();		
    	Properties A06 = new Properties();		
    	Properties A04 = new Properties();//98.08.28 add
    	List checktype_List = new LinkedList(); 
    	List upd_code_list = null;   		    	
    	boolean addflag=false;    	
    	List paramList = new ArrayList();
    	for(int j=0;j<15;j++){
    	    Amt[j][0]="";
    	    Amt[j][1]="";
    	    Amt[j][2]="";
    	    Amt[j][3]="0";
    	}   
    	acc_code_A01 = (bank_type.equals("6"))?acc_code_A01_6:acc_code_A01_7;
    	acc_code_A06 = (bank_type.equals("6"))?acc_code_A06_6:acc_code_A06_7;
    	acc_code_A04 = acc_code_A04_67;//98.08.28 add
    	
    	if(bank_type.equals("7") && (Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH) >= 10301)) {
           acc_code_A01 = acc_code_A01_6;
        }
    	for(int a01idx=0;a01idx<A01Data.size();a01idx++){
    	    System.out.println("a01idx="+a01idx);
    	    System.out.println("A01.bank_no="+(String)((List)A01Data.get(a01idx)).get(0));
    	    System.out.println("A01.s_report_name="+(String)((List)A01Data.get(a01idx)).get(1));
    	    System.out.println("A06.bank_no="+(String)((List)A06Data.get(a01idx)).get(0));
    	    System.out.println("A06.s_report_name="+(String)((List)A06Data.get(a01idx)).get(1));
    	    System.out.println("A04.bank_no="+(String)((List)A04Data.get(a01idx)).get(0));//98.08.28 add
    	    System.out.println("A04.s_report_name="+(String)((List)A04Data.get(a01idx)).get(1));//98.08.28 add
    	    addflag=false;
    	    if((((String)((List)A01Data.get(a01idx)).get(0)).equals((String)((List)A06Data.get(a01idx)).get(0))) && 
    	       (((String)((List)A01Data.get(a01idx)).get(0)).equals((String)((List)A04Data.get(a01idx)).get(0)))
    	      ){//bank_no 一樣    	           
    	        A01 = (Properties)((List)A01Data.get(a01idx)).get(2);//A01.amt集合
    	        A06 = (Properties)((List)A06Data.get(a01idx)).get(2);//A06.amt集合
    	        A04 = (Properties)((List)A04Data.get(a01idx)).get(2);//A04.amt集合
    	        for(int i=0;i<acc_code_A01.length;i++){    	          
    	              
    	             if(A01.getProperty(acc_code_A01[i]) != null){
    	                System.out.println("A01.get"+acc_code_A01[i]+"="+A01.getProperty(acc_code_A01[i]));
    	             }else{
    	                System.out.println("A01.get "+acc_code_A01[i]+" is null");
    	             } 
    	             if(A06.getProperty(acc_code_A06[i]) != null){    	             
    	               System.out.println("A06.get"+acc_code_A06[i]+"="+A06.getProperty(acc_code_A06[i])); 
    	             }else{
    	               System.out.println("A06.get "+acc_code_A06[i]+" is null");
    	             } 
    	             if(i==0 && A04.getProperty(acc_code_A04[i]) != null){//98.08.28 add    	             
    	               System.out.println("A04.get"+acc_code_A04[i]+"="+A04.getProperty(acc_code_A04[i])); 
    	             }
    	             
    	            if(A01.getProperty(acc_code_A01[i]) != null && A06.getProperty(acc_code_A06[i]) != null){
    	               System.out.println("A01.get"+acc_code_A01[i]+"="+A01.getProperty(acc_code_A01[i]));
    	               System.out.println("A06.get"+acc_code_A06[i]+"="+A06.getProperty(acc_code_A06[i]));
    	               if(i==0 && A04.getProperty(acc_code_A04[i]) != null){//98.08.28 add
    	                  System.out.println("A04.get"+acc_code_A04[i]+"="+A04.getProperty(acc_code_A04[i]));
    	               }
    	               Amt[i][0]=A01.getProperty(acc_code_A01[i]);//A01.amt
    	               Amt[i][1]=A06.getProperty(acc_code_A06[i]);//A06.amt   
    	               if(i==0) Amt[i][2]=A04.getProperty(acc_code_A04[i]);//A04.amt//98.08.28 add   	                 	                  
    	               if(upd_code.equals("1")){//A01/A06跨表檢核金額不符
    	                  //95.09.15 add 990000.970000轉金額的用Float取代Integer by 2295
    	                  if(( acc_code_A01[i].equals("990000") && ((Float.parseFloat(A01.getProperty(acc_code_A01[i])) != Float.parseFloat(A06.getProperty(acc_code_A06[i]))) || ( i ==0 && Float.parseFloat(A01.getProperty(acc_code_A01[i])) != Float.parseFloat(A04.getProperty(acc_code_A04[i])))))//990000 = 970000 = 840740
    	                  || (!acc_code_A01[i].equals("990000") && (Float.parseFloat(A01.getProperty(acc_code_A01[i])) < Float.parseFloat(A06.getProperty(acc_code_A06[i]))))){//A06 不可大於A01
    	                  	  System.out.println("titleLength["+i+"]set=1");
    	                  	  titleLength[i]="1";//要顯示的title index    	                      
    	                  	  Amt[i][3]="1";
    	                  	  System.out.println("Amt["+i+"][0]='"+Amt[i][0]+"':"+"Amt["+i+"][1]='"+Amt[i][1]+"':Amt["+i+"][2]='"+Amt[i][2]+"':Amt["+i+"][3]='"+Amt[i][3]+"'");    	               
    	                  	  addflag=true;
    	                  	  if(checktype.equals("1")){//98.09.23 add 排程檢核,發e-mail
    	                  	     paramList = new ArrayList();
    	                  	     paramList.add(upd_code);
    	                  	     upd_code_list = DBManager.QueryDB_SQLParam( "select * from cdshareno where cmuse_div='033' and identify_no='2' and input_order=?",paramList,"");     	
    	                  	     if(upd_code_list != null) checktype_List.add((String)((DataObject)upd_code_list.get(0)).getValue("cmuse_name"));
    	                  	     return checktype_List;
    	                  	  }
    	                  }
    	               }else if(upd_code.equals("2") && (A01.getProperty(acc_code_A01[i]).equals("0") && A06.getProperty(acc_code_A06[i]).equals("0"))){//A01/A06跨表檢核金額均為0
    	                  System.out.println("titleLength["+i+"]set=1");
    	                  titleLength[i]="1";//要顯示的title index    	                      
    	                  Amt[i][3]="1";
    	                  addflag=true;
    	                  System.out.println("Amt["+i+"][0]='"+Amt[i][0]+"':"+"Amt["+i+"][1]='"+Amt[i][1]+"':Amt["+i+"][2]='"+Amt[i][2]+"':Amt["+i+"][3]='"+Amt[i][3]+"'");     	               
    	                  if(checktype.equals("1")){//98.09.23 add 排程檢核,發e-mail
    	                     paramList = new ArrayList();
 	                  	     paramList.add(upd_code);
    	                  	 upd_code_list = DBManager.QueryDB_SQLParam( "select * from cdshareno where cmuse_div='033' and identify_no='2' and input_order=?",paramList,"");     	
    	                  	 if(upd_code_list != null) checktype_List.add((String)((DataObject)upd_code_list.get(0)).getValue("cmuse_name"));
    	                  	 return checktype_List;
    	                  }
    	               }
    	            }else{//a01 is null or a06 is null
    	               Amt[i][0]=A01.getProperty(acc_code_A01[i]);//A01.amt
    	               Amt[i][1]=A06.getProperty(acc_code_A06[i]);//A06.amt    	                 	                      	                   	            
    	               if(upd_code.equals("3")){//A01/A06跨表檢核尚未申報
    	                  System.out.println("titleLength["+i+"]set=1");
    	                  titleLength[i]="1";//要顯示的title index    	                      
    	                  Amt[i][3]="1";
    	                  addflag=true;
    	                  System.out.println("Amt["+i+"][0]='"+Amt[i][0]+"':"+"Amt["+i+"][1]='"+Amt[i][1]+"':Amt["+i+"][2]='"+Amt[i][2]+"':Amt["+i+"][3]='"+Amt[i][3]+"'");    	               
    	                  if(checktype.equals("1")){//98.09.23 add 排程檢核,發e-mail
    	                     paramList = new ArrayList();
  	                  	     paramList.add(upd_code);
    	                  	 upd_code_list = DBManager.QueryDB_SQLParam( "select * from cdshareno where cmuse_div='033' and identify_no='2' and input_order=?",paramList,"");     	
    	                  	 if(upd_code_list != null) checktype_List.add((String)((DataObject)upd_code_list.get(0)).getValue("cmuse_name"));
    	                  	 return checktype_List;
    	                  }
    	               }
    	            }
    	            
    	        }//end of acc_code_A01
    	        if(addflag){
    	           addflag=false;
    	           System.out.println("add "+(String)((List)A01Data.get(a01idx)).get(1));
    	           showData_detail.add((String)((List)A01Data.get(a01idx)).get(1));//s_report_name    
    	           showData_detail.add(Amt);
    	           showData.add(showData_detail);
    	           showData_detail = new LinkedList();
    	        }  
    	             	           
    	        //for(int j=0;j<8;j++){
    	        //    System.out.println("'"+Amt[j][0]+"':'"+Amt[j][1]+"'");    	              
    	        //}    	          
    	        Amt = new String[15][4]; 
    	        for(int j=0;j<15;j++){
    	            Amt[j][0]="";
    	            Amt[j][1]="";
    	            Amt[j][2]="";
    	            Amt[j][3]="0";
    	        }   
    	    }//end of a01.bank_no=a06.bank_no
    	}//a01 loop 
    	if(checktype.equals("1")){
           return null;
        }else{
    	   List returnList = new LinkedList();
    	   returnList.add(titleLength);
    	   returnList.add(showData);
    	   return returnList;
    	}
    }
    
    //98.09.14 add A01A02A99跨表檢核
    //99.02.10 fix 因A02.990120/A02.990130為平均值.故該公式不納入檢核 by 2295
    //99.02.22 fix A01A02A99跨表檢核-5~5均算檢核成功 by 2295 
    private List checkResult_A01A02A99(HashMap dataMap,String upd_code,String checktype){
        String[] titleLength={"0","0","0","0","0","0","0","0","0","0","0","0","0","0"};//欲顯示的檢核有誤title idx        
        String[] acc_code_A02 = {"990120","990130","990210","990210","220000","120000","990630","990810","990000","840760","990510","992150","992150","992150"};		          
        String[] acc_code_A99 = {"310000","140000","120700","990220","992130","992140","220900","140000","992510","992610","990511","990510","992550","992650"};		  
        String[] title = {"信用部淨值","信用部固定資產淨值","內部融資","內部融資","存款","放款","公庫存款","固定資產","逾期放款金額","應予觀察放款","非會員無擔保消費性貸款","無擔保消費性貸款","無擔保消費性 貸款中之逾期放款","無擔保消費性貸款 中之應予觀察放款"};
        String[][] Amt = new String[14][4];//[n][0]->A02.amt [n][1]->A99.amt [n][2]->showflag           		      
    	List A02Data = (List)dataMap.get("A02");
    	List A99Data = (List)dataMap.get("A99");    	
    	List showData = new LinkedList();//回傳的檢核有誤data
    	List showData_detail = new LinkedList();//檢核有誤的each row data
    	Properties A02 = new Properties();		
    	Properties A99 = new Properties();		
    	List checktype_List = new LinkedList(); 
    	List upd_code_list = null;   	
    	boolean addflag=false;    	
    	List paramList = new ArrayList();
    	for(int j=0;j<14;j++){
    	    Amt[j][0]="";
    	    Amt[j][1]="";    	   
    	    Amt[j][2]="0";
    	}   
    	
    	double	amt_L = 0.0, amt_R = 0.0;
    	for(int a02idx=0;a02idx<A02Data.size();a02idx++){
    	    System.out.println("a02idx="+a02idx);
    	    System.out.println("A02.bank_no="+(String)((List)A02Data.get(a02idx)).get(0));
    	    System.out.println("A02.s_report_name="+(String)((List)A02Data.get(a02idx)).get(1));
    	    System.out.println("A99.bank_no="+(String)((List)A99Data.get(a02idx)).get(0));
    	    System.out.println("A99.s_report_name="+(String)((List)A99Data.get(a02idx)).get(1));
    	   
    	    addflag=false;
    	    if(((String)((List)A02Data.get(a02idx)).get(0)).equals((String)((List)A99Data.get(a02idx)).get(0))){//bank_no 一樣    	           
    	        A02 = (Properties)((List)A02Data.get(a02idx)).get(2);//A02.amt集合
    	        A99 = (Properties)((List)A99Data.get(a02idx)).get(2);//A99.amt集合
    	        //99.02.10 fix A02.990120/A02.990130公式不納入檢核    	        
    	        for(int i=2;i<acc_code_A02.length;i++){
    	             if(A02.getProperty(acc_code_A02[i]) != null){
    	                System.out.println("A02.get"+acc_code_A02[i]+"="+A02.getProperty(acc_code_A02[i]));
    	             }else{
    	                System.out.println("A02.get "+acc_code_A02[i]+" is null");
    	             } 
    	             if(A99.getProperty(acc_code_A99[i]) != null){    	             
    	               System.out.println("A99.get"+acc_code_A99[i]+"="+A99.getProperty(acc_code_A99[i])); 
    	             }else{
    	               System.out.println("A99.get "+acc_code_A99[i]+" is null");
    	             } 
    	             
    	             
    	            if(A02.getProperty(acc_code_A02[i]) != null && A99.getProperty(acc_code_A99[i]) != null){
    	               System.out.println("A02.get"+acc_code_A02[i]+"="+A02.getProperty(acc_code_A02[i]));
    	               System.out.println("A99.get"+acc_code_A99[i]+"="+A99.getProperty(acc_code_A99[i]));
    	               
    	               Amt[i][0]=A02.getProperty(acc_code_A02[i]);//A02.amt
    	               Amt[i][1]=A99.getProperty(acc_code_A99[i]);//A99.amt   
    	               amt_L = Double.parseDouble(A02.getProperty(acc_code_A02[i]));
				       amt_R = Double.parseDouble(A99.getProperty(acc_code_A99[i])); 	                 	                  
    	               if(upd_code.equals("1")){//A02/A99跨表檢核金額不符                                            
    	                  if((!(acc_code_A02[i].equals("990210") || acc_code_A02[i].equals("992150")) && Math.abs(amt_L - amt_R) > 5  ) //99.02.22 fix -5~5均算檢核成功 
    	                  || ( (acc_code_A02[i].equals("990210") || acc_code_A02[i].equals("992150")) && (Float.parseFloat(A02.getProperty(acc_code_A02[i])) < Float.parseFloat(A99.getProperty(acc_code_A99[i]))))){//A99 不可大於A02
    	                  	  System.out.println("titleLength["+i+"]set=1");
    	                  	  titleLength[i]="1";//要顯示的title index    	                      
    	                  	  Amt[i][2]="1";
    	                  	  System.out.println("Amt["+i+"][0]='"+Amt[i][0]+"':"+"Amt["+i+"][1]='"+Amt[i][1]+"':Amt["+i+"][2]='"+Amt[i][2]+"'");    	               
    	                  	  System.out.println("Math.abs="+Math.abs(amt_L - amt_R));
    	                  	  addflag=true;
    	                  	  if(checktype.equals("1")){//98.09.22 add 排程檢核,發e-mail
    	                  	     paramList = new ArrayList();
    	                  	     paramList.add(upd_code);
    	                  	     upd_code_list = DBManager.QueryDB_SQLParam( "select * from cdshareno where cmuse_div='033' and identify_no='4' and input_order=?",paramList,"");     	
    	                  	     if(upd_code_list != null) checktype_List.add((String)((DataObject)upd_code_list.get(0)).getValue("cmuse_name"));
    	                  	     return checktype_List;
    	                  	  }
    	                  }
    	               }else if(upd_code.equals("2") && (A02.getProperty(acc_code_A02[i]).equals("0") && A99.getProperty(acc_code_A99[i]).equals("0"))){//A02/A99跨表檢核金額均為0
    	                  System.out.println("titleLength["+i+"]set=1");
    	                  titleLength[i]="1";//要顯示的title index    	                      
    	                  Amt[i][2]="1";
    	                  addflag=true;
    	                  if(checktype.equals("1")){//98.09.22 add 排程檢核,發e-mail
    	                     paramList = new ArrayList();
 	                  	     paramList.add(upd_code);
    	                  	 upd_code_list = DBManager.QueryDB_SQLParam( "select * from cdshareno where cmuse_div='033' and identify_no='4' and input_order=?",paramList,"");     	
    	                  	 if(upd_code_list != null) checktype_List.add((String)((DataObject)upd_code_list.get(0)).getValue("cmuse_name"));
    	                  	 return checktype_List;
    	                  }
    	                  System.out.println("Amt["+i+"][0]='"+Amt[i][0]+"':"+"Amt["+i+"][1]='"+Amt[i][1]+"':Amt["+i+"][2]='"+Amt[i][2]+"'");     	               
    	               }
    	            }else{//a02 is null or a99 is null
    	               Amt[i][0]=A02.getProperty(acc_code_A02[i]);//A02.amt
    	               Amt[i][1]=A99.getProperty(acc_code_A99[i]);//A99.amt    	                 	                      	                   	            
    	               if(upd_code.equals("3")){//A02/A99跨表檢核尚未申報
    	                  System.out.println("titleLength["+i+"]set=1");
    	                  titleLength[i]="1";//要顯示的title index    	                      
    	                  Amt[i][2]="1";
    	                  addflag=true;
    	                  if(checktype.equals("1")){//98.09.22 add 排程檢核,發e-mail
    	                     paramList = new ArrayList();
  	                  	     paramList.add(upd_code);
    	                  	 upd_code_list = DBManager.QueryDB_SQLParam( "select * from cdshareno where cmuse_div='033' and identify_no='4' and input_order=?",paramList,"");     	
    	                  	 if(upd_code_list != null) checktype_List.add((String)((DataObject)upd_code_list.get(0)).getValue("cmuse_name"));
    	                  	 return checktype_List;
    	                  }
    	                  System.out.println("Amt["+i+"][0]='"+Amt[i][0]+"':"+"Amt["+i+"][1]='"+Amt[i][1]+"':Amt["+i+"][2]='"+Amt[i][2]+"'");    	               
    	               }
    	            }
    	            
    	        }//end of acc_code_A02
    	        if(addflag){
    	           addflag=false;
    	           System.out.println("add "+(String)((List)A02Data.get(a02idx)).get(1));
    	           showData_detail.add((String)((List)A02Data.get(a02idx)).get(1));//s_report_name    
    	           showData_detail.add(Amt);
    	           showData.add(showData_detail);
    	           showData_detail = new LinkedList();
    	        }   	           
    	        	          
    	        Amt = new String[14][3]; 
    	        for(int j=0;j<14;j++){
    	            Amt[j][0]="";
    	            Amt[j][1]="";
    	            Amt[j][2]="0";
    	        }   
    	    }//end of a02.bank_no=a99.bank_no
    	}//a02 loop 
    	if(checktype.equals("1")){
           return null;
        }else{
    	   List returnList = new LinkedList();
    	   returnList.add(titleLength);
    	   returnList.add(showData);
    	   return returnList;
    	}
    }
    
    private List checkResult_A10(HashMap dataMap,String upd_code,String checktype){
        String[] titleLength={"0","0","0","0","0","0","0","0","0","0","0","0","0","0","0","0","0","0"};//欲顯示的檢核有誤title idx
        String[] acc_code_L = {"190000","220000","120000","220900","140000","990000","840760","990510","992150",
        							"992150","992150","990610","120800","990021","990002","990001","990006","990005"};
        String[] acc_code_R = {"400000","992130","992140","990630","990810","992510","992610","990511","990510",
        							"992550","992650","990612","990025","990025","840760","990005","990010","120000"};
        String[] title = {"A01資產負債表之資產合計=負債合計+淨值合計",
        		"A01資產負債表中之存款總額=法定比率分析統計表中(19)正會員存款總額+(20)贊助會員存款總額=(21)非會員存款總額(含公庫存款)",
        		"A01資產負債表中之放款淨額+備抵呆帳-放款+備抵呆帳-催收項項-內部融資=法定比率分析統計表中(22)正會員放款總額+(23)贊助會員放款總額+(24)非會員放款總額(不含內部融資)",
        		"A01資產負債表中之1/2公庫存款=法定比率分析統計表中(25)1/2公庫存款",
        		"A01資產負債表中之固定資產淨額=法定比率分析統計表中(27)信用部固定資產淨額",
        		"A01資產負債表中之逾期放款總額=A04應予觀察放款總計表中之逾期放款總額=法定比率分析統計表中(38)正會員放款中之逾期放款+(39)贊助會員放款中之逾期放款+(40)非會員放款中之逾期放款+(41)內部融資中之逾期放款",
        		"A04應予觀察放款統計表中之應予觀察放款總額=法定比率分析統計表中(42)正會員放款中之應予觀察放款+(43)贊助會員放款中之應予觀察放款+(44)非會員放款中之應予觀察放款+(45)內部融資中之應予觀察放款",
        		"法定比率分析統計表中(50)非會員無擔保消費性貸款=(51)非會員無擔保消費性政策貸款+(52)非會員無擔保消費性非政策貸款",
        		"法定比率分析統計表中(53)無擔保消費性貸款 > (50)非會員無擔保消費性貸款",
        		"法定比率分析統計表中(53)無擔保消費性貸款 > (54)無擔保消費性貸款中之逾期放款",
        		"法定比率分析統計表中(53)無擔保消費性貸款 > (55)無擔保消費性貸款中之應予觀察放款",
        		"法定比率分析統計表中(24)非會員放款總額(不含內部融資) > (56)非會員放款中之政策性農業專案貸款",
        		"A01資產負債表中之備抵呆帳-放款+ 備抵呆帳-催收款項=A10應予評估資產彙總表中之放款帳列備抵呆帳合計",
        		"A10應予評估資產彙總表中之放款帳列備抵呆帳一類+放款帳列備抵呆帳二類+放款帳列備抵呆帳三類+放款帳列備抵呆帳四類=放款帳列備抵呆帳合計",
        		"A10應予評估資產彙總表中之應予評估放款二類+應予評估放款三類+應予評估放款四類 >=A04應予觀察放款總計表中之應予觀察放款總額+A01資產負債表中之逾期放款總額",
        		"A10應予評估資產彙總表中之應予評估放款一類+應予評估放款二類+應予評估放款三類+應予評估放款四類=應予評估放款合計",
        		"A10應予評估資產彙總表中之應予評估投資一類+應予評估投資二類+應予評估投資三類+應予評估投資四類=應予評估投資合計",
        		"A10應予評估資產彙總表中之應予評估放款合計=A01資產負債表中之放款總額-A02(990611)對直轄市、縣(市)政府、離島地區鄉(鎮、市)公所辦理之授信總額 "};
        //第06項檢核,增加檢核19,
        //rightSQL比leftSQL多一筆acc_code,rightSQL排序後撈出最後一筆acc_code剛好為第06項檢核的840740，故加到amt陣列19
        String[][] Amt = new String[19][3];//[n][0]->dataL.amt [n][1]->dataR.amt [n][2]->showflag           		      
    	List dataL = (List)dataMap.get("dataL");
    	List dataR = (List)dataMap.get("dataR");
    	List showData = new LinkedList();//回傳的檢核有誤data
    	List showData_detail = new LinkedList();//檢核有誤的each row data
    	Properties lProp = new Properties();		
    	Properties rProp = new Properties();		    	
    	boolean addflag=false; 
    	List checktype_List = new LinkedList(); 
    	List upd_code_list = null;   	
    	List paramList = new ArrayList();
    	for(int j=0;j<acc_code_R.length+1;j++){
    	    Amt[j][0]="";
    	    Amt[j][1]="";
    	    Amt[j][2]="0";
    	}   
    	for(int leftIdx=0;leftIdx<dataL.size();leftIdx++){
    	    System.out.println("leftIdx="+leftIdx);
    	    System.out.println("left.bank_no="+(String)((List)dataL.get(leftIdx)).get(0));
    	    System.out.println("left.s_report_name="+(String)((List)dataL.get(leftIdx)).get(1));
    	    System.out.println("right.bank_no="+(String)((List)dataR.get(leftIdx)).get(0));
    	    System.out.println("right.s_report_name="+(String)((List)dataR.get(leftIdx)).get(1));
    	    if(((String)((List)dataL.get(leftIdx)).get(0)).equals((String)((List)dataR.get(leftIdx)).get(0))){//bank_no 一樣    	           
    	        lProp = (Properties)((List)dataL.get(leftIdx)).get(2);//left.amt集合
    	        rProp = (Properties)((List)dataR.get(leftIdx)).get(2);//right.amt集合
    	        boolean addList = false;
    	        for(int i=0;i<acc_code_L.length;i++){ 
    	        	Amt[i][0]="";
    	    	    Amt[i][1]="";
    	    	    Amt[i][2]="0";
    	            if(lProp.getProperty(acc_code_L[i]) != null && rProp.getProperty(acc_code_R[i]) != null){
    	               Amt[i][0]=lProp.getProperty(acc_code_L[i]).toString();//left.amt
    	               Amt[i][1]=rProp.getProperty(acc_code_R[i]).toString();//right.amt   
    	               //System.out.println("get"+acc_code_L[i]+"="+Amt[i][0]);
    	               //System.out.println("get"+acc_code_R[i]+"="+Amt[i][1]);
    	               addflag = false;
    	               if( i==8 || i==11 ){//存表跨表檢核金額不符
						   if(!(Long.parseLong(Amt[i][0])>Long.parseLong(Amt[i][1]))){
							   addflag=true;
    	            	   }
    	               }else if( i==9 || i==10 || i==14){
						   if(!(Long.parseLong(Amt[i][0])>=Long.parseLong(Amt[i][1]))){
							   addflag=true;
    	            	   }
    	               }else{
    	            	   if(i==5){
	    	            	   Amt[acc_code_L.length][1]= rProp.getProperty("840740").toString();
    	            	   }
	    	               if(i==5){
		    	            	if(!(Amt[i][0]).equals(rProp.getProperty("840740").toString())
		    	            		 ||!(Amt[i][0]).equals(Amt[i][1])
		    	            		 ||!(rProp.getProperty("840740").toString()).equals(Amt[i][1])){
		    	            		   addflag=true;
	    	            	    }
	    	               }else{
	    	            	   if(!Amt[i][0].equals(Amt[i][1])){
		    	            	   addflag=true;
		    	               }
	    	            	   
	    	               }
    	               }
    	               if(addflag){
    	            	   //System.out.println("titleLength["+i+"]set=1");
	    	               titleLength[i]="1";//要顯示的title index
	    	               Amt[i][2]="1";
	    	               //System.out.println("Amt["+i+"][0]='"+Amt[i][0]+"':"+"Amt["+i+"][1]='"+Amt[i][1]+"':Amt["+i+"][2]='"+Amt[i][2]+"'"); 
	    	               if(i==5){System.out.println("Amt["+acc_code_L.length+"][0]='"+Amt[i][0]+"':"+"Amt["+acc_code_L.length+"][1]='"+Amt[acc_code_L.length][1]+"':Amt["+acc_code_L.length+"][2]='"+Amt[acc_code_L.length][2]+"'");}
	    	               if(checktype.equals("1")){//98.09.23 add 排程檢核,發e-mail
	    	                   paramList = new ArrayList();
	    	                   paramList.add(upd_code);    	                  
	    	                   upd_code_list = DBManager.QueryDB_SQLParam( "select * from cdshareno where cmuse_div='033' and identify_no='1' and input_order=?",paramList,"");     	
	    	                   if(upd_code_list != null) checktype_List.add((String)((DataObject)upd_code_list.get(0)).getValue("cmuse_name"));
	    	                   return checktype_List;
	    	               }
	    	               addList = true;
    	               }
    	            }else{//lProp is null or rProp is null
    	               Amt[i][0]=lProp.getProperty(acc_code_L[i]);//lProp.amt
    	               Amt[i][1]=rProp.getProperty(acc_code_R[i]);//rProp.amt    	                 	                      	                   	            
    	            }
    	        }//end of acc_code_L
    	        if(addList){
    	           addflag=false;
    	           //System.out.println("add "+(String)((List)dataL.get(leftIdx)).get(1));
    	           showData_detail.add((String)((List)dataL.get(leftIdx)).get(1));//s_report_name    
    	           showData_detail.add(Amt);
    	           showData.add(showData_detail);
    	           showData_detail = new LinkedList();
    	        }  
    	             	           
    	        //for(int j=0;j<18;j++){
    	        //    System.out.println("'"+Amt[j][0]+"':'"+Amt[j][1]+"'");    	              
    	        //}    	          
    	        Amt = new String[acc_code_R.length+1][3]; 
    	        for(int j=0;j<acc_code_R.length+1;j++){
    	            Amt[j][0]="";
    	            Amt[j][1]="";
    	            Amt[j][2]="0";
    	        }   
    	    }//end of a01.bank_no=a05.bank_no
    	}//left loop 
    	if(checktype.equals("1")){
           return null;
        }else{
    	   List returnList = new LinkedList();
    	   returnList.add(titleLength);
    	   returnList.add(showData);
    	   return returnList;
    	}
    }
    
    //取得查詢結果
    //100.01.11 fix sqlInjection/根據查詢年度.100年以後取得新縣市別.100年以前取得舊縣市別 
    private HashMap getQryResultA01A05(String bank_type_list,String S_YEAR,String S_MONTH,String unit,String upd_code,String bank_type,String tbank_no,String lguser_id){    	   
    		//查詢條件        		
    		String sqlCmd = "";
    		String rule_1_A01 = "";
    		String rule_1_A05 = "";
    		String rule_2 = "";    		
    		String tmpbank_type="";
    		String tmpbank_no="";
    		List A01DataListALL = new LinkedList();
    		List A05DataListALL = new LinkedList();
    		//100.01.11 add 查詢年度100年以前.縣市別不同===============================
  	    	String cd01_table = (Integer.parseInt(S_YEAR) < 100)?"cd01_99":""; 
  	    	String wlx01_m_year = (Integer.parseInt(S_YEAR) < 100)?"99":"100"; 
  	    	List paramList = new ArrayList(); 
  	    	List rule_1_A01_paramList = new ArrayList(); 
  	    	List rule_1_A05_paramList = new ArrayList(); 
  	    	//=====================================================================    
    		if(lguser_id.equals("A111111111") || bank_type.equals("2") || bank_type.equals("1") || bank_type.equals("3")){//98.09.29 add 中央存保
    		    tmpbank_type = "Z";
			    tmpbank_no = "9999999";
		    }else{
		        tmpbank_type = bank_type;
				tmpbank_no = tbank_no;
			}	
			
            //A01===================================================================================================
            /*  select * from 
              (
               select NVL((BN01.BANK_NO || BN01.BANK_NAME), ' ')  as S_Report_Name,  
                  	BN01.BANK_NO, BN01.BANK_NAME, bn01.bank_type ,  
                		nvl(wlx01.CENTER_NO,' ') as CENTER_NO, 
              		nvl(wlx01.M2_NAME,' ')   as M2_NAME, 
              		wlx01.CANCEL_no, wlx01.CANCEL_date,
              		nvl(cd01.hsien_id,' ') as  hsien_id,  
              	    a01.acc_code,decode(A01.acc_code,'320100',A01_320100.amt,a01.amt) as amt,
              	    a01.acc_range,
              	    nvl(cd01.hsien_name,'OTHER')  as  hsien_name, 
              		cd01.FR001W_output_order     as  FR001W_output_order 
               from  (select * from cd01 where  cd01.hsien_id <> 'Y') cd01 
                   	left join wlx01 on wlx01.hsien_id=cd01.hsien_id  
              	 	left join (select BANK_NO, BANK_NAME,  bank_type  
              	    from bn01 where bn01.bank_type in('6','7'))  bn01  
              		on wlx01.bank_no=bn01.bank_no  
               left join 
               	    (select  bank_code,m_year,m_month,a01.acc_code,round(amt/1,0),
              		 decode(acc_code,'310100','01','310200','02','310300','03','310400','04','310500','05','310600','06','310800','07','320100','08','130200','09') as acc_range 
              		 from  A01 
              		 where M_Year =  94  and M_month =  8
              		 and  a01.acc_code in('310100','310200','310300','310400','310500','310600','310800','320100','130200')		 
              		 order by a01.bank_code,acc_range	
                      ) A01
                	  on bn01.bank_no = A01.Bank_Code	  	 
               left join
               	    (select m_year,m_month,bank_code,round(sum(amt)/1,0) as amt 
              		 from (select  * FROM  A01 
                       	   where M_Year =  94  and M_month =  8
              			   and  acc_code in('320100','320200')) A01_320100_320200
              		 group by m_year,m_month,bank_code
              		 order by m_year,m_month,bank_code		  
              		 )A01_320100	
               	  on bn01.bank_no = A01_320100.Bank_Code
              )Temp_Output  where BANK_NO <> ' '
              order by  Bank_Type, FR001W_output_order, BANK_NO,acc_range
              //========================================================================================================
              //A05=====================================================================================================
              select * from 
              (
               select NVL((BN01.BANK_NO || BN01.BANK_NAME), ' ')  as S_Report_Name,  
                  	BN01.BANK_NO, BN01.BANK_NAME, bn01.bank_type ,  
                		nvl(wlx01.CENTER_NO,' ') as CENTER_NO, 
              		nvl(wlx01.M2_NAME,' ')   as M2_NAME, 
              		wlx01.CANCEL_no, wlx01.CANCEL_date,
              		nvl(cd01.hsien_id,' ') as  hsien_id,  
              	   a05.acc_code,decode(A05.acc_code,'910401',A05_910401.amt,a05.amt) as amt,
              	   a05.acc_range,
              	    nvl(cd01.hsien_name,'OTHER')  as  hsien_name, 
              		cd01.FR001W_output_order     as  FR001W_output_order 
               from  (select * from cd01 where  cd01.hsien_id <> 'Y') cd01 
                   	left join wlx01 on wlx01.hsien_id=cd01.hsien_id  
              	 	left join (select BANK_NO, BANK_NAME,  bank_type  
              	    from bn01 where bn01.bank_type in('6','7'))  bn01  
              		on wlx01.bank_no=bn01.bank_no  
               left join 
               	    (select  bank_code,m_year,m_month,a05.acc_code,round(amt/1,0),
              		 decode(acc_code,'910101','01','910102','02','910103','03','910104','04','910105','05','910106','06','910107','07','910108','08','910401','09') as acc_range
              		 from  A05 		 
              		 where M_Year =  94  and M_month =  8
              		 and  a05.acc_code in('910101','910102','910103','910104','910105','910106','910107','910108','910401')		 
              		 order by a05.bank_code,acc_range	
                      ) A05
                	  on bn01.bank_no = A05.Bank_Code	  	 
               left join
               	    (select m_year,m_month,bank_code,round(sum(amt)/1,0) as amt 
              		 from (select  * FROM  A05 
                       	   where M_Year =  94  and M_month =  8
              			   and  acc_code in('910401','910402','910403','910404')) A05_910401_910404
              		 group by m_year,m_month,bank_code
              		 order by m_year,m_month,bank_code		  
              		 ) A05_910401
              	 on bn01.bank_no = A05_910401.Bank_Code		
              )Temp_Output  where BANK_NO <> ' '		  
              order by  Bank_Type, FR001W_output_order, BANK_NO,acc_range
            //========================================================================================================================================
			*/
			
			rule_1_A01 = " select * from "
              	   + " ( "
               	   + "	select NVL((BN01.BANK_NO || BN01.BANK_NAME), ' ')  as S_Report_Name," 
                   + "		   BN01.BANK_NO, BN01.BANK_NAME, bn01.bank_type ," 
                   + "		   nvl(wlx01.CENTER_NO,' ') as CENTER_NO,"
              	   + "	   	   nvl(wlx01.M2_NAME,' ')   as M2_NAME, "
              	   + "		   wlx01.CANCEL_no, wlx01.CANCEL_date,"
              	   + "	   	   nvl(cd01.hsien_id,' ') as  hsien_id,"
              	   + "		   a01.acc_code,decode(A01.acc_code,'320100',A01_320100.amt,a01.amt) as amt,"
              	   + "		   a01.acc_range,"
              	   + "		   nvl(cd01.hsien_name,'OTHER')  as  hsien_name,"
              	   + "	  	   cd01.FR001W_output_order     as  FR001W_output_order"
               	   + "	from  (select * from "+cd01_table+" cd01 where  cd01.hsien_id <> 'Y') cd01 "
                   + "	left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id " 
              	   + "	left join (select BANK_NO, BANK_NAME,  bank_type " 
              	   + "			   from (select * from bn01 where m_year=?)bn01 where bn01.bank_type in('6','7'))  bn01 " 
              	   + "		 on wlx01.bank_no=bn01.bank_no"
               	   + "	left join "
               	   + "			 (select  bank_code,m_year,m_month,a01.acc_code,"               	   
               	   + "            round(amt/?,0) as amt,"
              	   + "			  decode(acc_code,'310100','01','310200','02','310300','03','310400','04','310500','05',"
              	   + "					 '310600','06','310800','07','320100','08','130200','09') as acc_range "
              	   + "			  from  A01"
              	   + "			  where M_Year = ? and M_month = ?"
              	   + "	 		  and   a01.acc_code in('310100','310200','310300','310400','310500','310600','310800','320100','130200')"		 
              	   + "	  	 	  order by a01.bank_code,acc_range"	
                   + "		     ) A01 "
                   + "		 on bn01.bank_no = A01.Bank_Code"	  	 
               	   + " left join "
               	   + "		    (select m_year,m_month,bank_code,"
               	   + "           round(sum(amt)/?,0) as amt "
              	   + "			 from (select  * FROM  A01 "
                   + "		     where M_Year = ? and M_month = ?"
              	   + "		     and  acc_code in('320100','320200')) A01_320100_320200 "
              	   + "			 group by m_year,m_month,bank_code"
              	   + "			 order by m_year,m_month,bank_code"		  
              	   + "			)A01_320100"
               	   + "		on bn01.bank_no = A01_320100.Bank_Code"
              	   + " ) Temp_Output  ";
            
            rule_1_A01_paramList.add(wlx01_m_year);//wlx01
			rule_1_A01_paramList.add(wlx01_m_year);//bn01 
			rule_1_A01_paramList.add(unit);
			rule_1_A01_paramList.add(S_YEAR);	   
			rule_1_A01_paramList.add(String.valueOf(Integer.parseInt(S_MONTH)));
			rule_1_A01_paramList.add(unit);
			rule_1_A01_paramList.add(S_YEAR);
			rule_1_A01_paramList.add(String.valueOf(Integer.parseInt(S_MONTH)));
			
            rule_1_A05 = " select * from "
              		   + " ("
               		   + "	select NVL((BN01.BANK_NO || BN01.BANK_NAME), ' ')  as S_Report_Name,"  
                  	   + "		   BN01.BANK_NO, BN01.BANK_NAME, bn01.bank_type ,"  
                	   + "			nvl(wlx01.CENTER_NO,' ') as CENTER_NO," 
              		   + "          nvl(wlx01.M2_NAME,' ')   as M2_NAME, "
              		   + "          wlx01.CANCEL_no, wlx01.CANCEL_date,"
              		   + "			nvl(cd01.hsien_id,' ') as  hsien_id,"
              	       + "          a05.acc_code,decode(A05.acc_code,'910401',A05_910401.amt,a05.amt) as amt,"
              	       + "          a05.acc_range,"
              	       + "			nvl(cd01.hsien_name,'OTHER')  as  hsien_name," 
              		   + "          cd01.FR001W_output_order     as  FR001W_output_order "
               		   + " from  (select * from "+cd01_table+" cd01 where  cd01.hsien_id <> 'Y') cd01 "
                   	   + " left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id " 
              	 	   + " left join (select BANK_NO, BANK_NAME,  bank_type " 
              	       + "			  from (select * from bn01 where m_year=?)bn01 where bn01.bank_type in('6','7'))  bn01 " 
              		   + "		on wlx01.bank_no=bn01.bank_no " 
               	       + " left join "
               	       + "		(select bank_code,m_year,m_month,a05.acc_code,"
               	       + "              round(amt/?,0) as amt,"
              		   + "				decode(acc_code,'910101','01','910102','02','910103','03','910104','04',"
              		   + "				'910105','05','910106','06','910107','07','910108','08','910401','09') as acc_range "
              		   + "		 from  A05" 		 
              		   + "		 where M_Year = ? and M_month = ? "            		 
              		   + "		 and  a05.acc_code in('910101','910102','910103','910104','910105','910106','910107','910108','910401')"		 
              		   + "		 order by a05.bank_code,acc_range"	
                       + " 		 ) A05 "
                	   + "       on bn01.bank_no = A05.Bank_Code "	  	 
               	  	   + " left join "
               	       + "		(select m_year,m_month,bank_code,"
               	       + "       round(sum(amt)/?,0) as amt "
              		   + "		 from (select  * FROM  A05 "
              		   + "		       where M_Year = ? and M_month = ?"
              		   + "	           and  acc_code in('910401','910402','910403','910404')) A05_910401_910404"
              		   + "		       group by m_year,m_month,bank_code"
              		   + "		       order by m_year,m_month,bank_code"		  
              		   + "			   ) A05_910401"
              	 	   + "		on bn01.bank_no = A05_910401.Bank_Code"		
              		   + " )Temp_Output ";
            rule_1_A05_paramList.add(wlx01_m_year);//wlx01
			rule_1_A05_paramList.add(wlx01_m_year);//bn01  
			rule_1_A05_paramList.add(unit);
			rule_1_A05_paramList.add(S_YEAR);	   
			rule_1_A05_paramList.add(String.valueOf(Integer.parseInt(S_MONTH)));
			rule_1_A05_paramList.add(unit);
			rule_1_A05_paramList.add(S_YEAR);
			rule_1_A05_paramList.add(String.valueOf(Integer.parseInt(S_MONTH)));
			rule_2 = " where 	BANK_NO <> ' '		and "
				   + " ( "
				   + " (? = 'Z') or "
			       + " ((? = '6' or ? = '7') and  ? = BANK_NO) or "
			       + " (? = '8' and ? = CENTER_NO) or"
			       + " (? = 'B' and ? = M2_NAME) "
				   + " )    and "
				   + " ( "
				   + " (?= 'ALL') or "
 			       + " ((?= '6') and (Bank_Type = '6')) or "
				   + " ((?= '7') and (Bank_Type = '7')) "
				   + " )"                                                                               
				   + " order by  Bank_Type,BANK_NO,acc_range,FR001W_output_order ";
			paramList.add(tmpbank_type);
			paramList.add(tmpbank_type);
			paramList.add(tmpbank_type);
			paramList.add(tmpbank_no);
			paramList.add(tmpbank_type);
			paramList.add(tmpbank_no);
			paramList.add(tmpbank_type);
			paramList.add(tmpbank_no);
			paramList.add(bank_type_list);
			paramList.add(bank_type_list);
			paramList.add(bank_type_list);		
			for(int i=0;i<paramList.size();i++){
	            rule_1_A01_paramList.add(paramList.get(i));		
			}
			for(int i=0;i<paramList.size();i++){
	            rule_1_A05_paramList.add(paramList.get(i));		
			}	   
			List dbData = DBManager.QueryDB_SQLParam( rule_1_A01 + rule_2,rule_1_A01_paramList,"cancel_date,amt");   
			String tmpBank_no="";
			String tmpBank_no_name="";
			Properties A01Data = new Properties();		
			List A01DataList = new LinkedList();
			if(dbData != null && dbData.size() > 0){
			   for(int i=0;i<dbData.size();i++){			 			          
			       if(i==0){
			           tmpBank_no = (String)((DataObject)dbData.get(i)).getValue("bank_no");     			       			       			       
			           tmpBank_no_name = (String)((DataObject)dbData.get(i)).getValue("s_report_name");
			       }    
			       if(((String)((DataObject)dbData.get(i)).getValue("bank_no")).equals(tmpBank_no)){//與前一個bank_no相同時,儲存acc_code,amt			          			       			       
			       	  if(((DataObject)dbData.get(i)).getValue("acc_code") != null){//95.06.19 acc_code == null時,只會有一筆資料	            			            			             			             
			             A01Data.setProperty((String)((DataObject)dbData.get(i)).getValue("acc_code"),(((DataObject)dbData.get(i)).getValue("amt")).toString());			       			       			            
			             System.out.println("A01.set("+(String)((DataObject)dbData.get(i)).getValue("acc_code")+")="+(((DataObject)dbData.get(i)).getValue("amt")).toString());
			          }   
			         
			       }else{ 
			          System.out.println("A01 add Bank_no="+tmpBank_no);
			          A01DataList.add(tmpBank_no);
			          A01DataList.add(tmpBank_no_name);
			          A01DataList.add(A01Data);
			          A01DataListALL.add(A01DataList);//欲回傳的A01 data
			          A01DataList = new LinkedList();//單一機構的data
			          A01Data = new Properties();//A01.acc_code|amt				          
			          tmpBank_no = (String)((DataObject)dbData.get(i)).getValue("bank_no");
			          tmpBank_no_name = (String)((DataObject)dbData.get(i)).getValue("s_report_name");
			          System.out.println("now tmpbank_no="+tmpBank_no);
			          if(((DataObject)dbData.get(i)).getValue("acc_code") != null){//95.06.19 acc_code == null時,只會有一筆資料	            			            			             			             
			             A01Data.setProperty((String)((DataObject)dbData.get(i)).getValue("acc_code"),(((DataObject)dbData.get(i)).getValue("amt")).toString());   
			             System.out.println("A01.set("+(String)((DataObject)dbData.get(i)).getValue("acc_code")+")="+(((DataObject)dbData.get(i)).getValue("amt")).toString());
			          }
			       }			       
			   }			   
			   //將最後一筆寫回A01DataListALL
			   System.out.println("A01 add last Bank_no="+tmpBank_no);				  
			   A01DataList.add(tmpBank_no);
			   A01DataList.add(tmpBank_no_name);
			   A01DataList.add(A01Data);
			   A01DataListALL.add(A01DataList);			   			   
			}
			dbData = DBManager.QueryDB_SQLParam( rule_1_A05 + rule_2,rule_1_A05_paramList,"cancel_date,amt");   
			tmpBank_no="";	
			tmpBank_no_name="";		
			Properties A05Data = new Properties();		
			List A05DataList = new LinkedList();
			if(dbData != null && dbData.size() > 0){
			   for(int i=0;i<dbData.size();i++){			 			         
			       if(i==0){
			          tmpBank_no = (String)((DataObject)dbData.get(i)).getValue("bank_no");     			       			       			       
			          tmpBank_no_name = (String)((DataObject)dbData.get(i)).getValue("s_report_name");
			       }   
			       if(((String)((DataObject)dbData.get(i)).getValue("bank_no")).equals(tmpBank_no)){//與前一個bank_no相同時,儲存acc_code,amt			          			       			       
			          if(((DataObject)dbData.get(i)).getValue("acc_code") != null){//95.06.19 acc_code == null時,只會有一筆資料
			             A05Data.setProperty((String)((DataObject)dbData.get(i)).getValue("acc_code"),(((DataObject)dbData.get(i)).getValue("amt")).toString());			       			       			            
			             System.out.println("A05.set("+(String)((DataObject)dbData.get(i)).getValue("acc_code")+")="+(((DataObject)dbData.get(i)).getValue("amt")).toString());
			          }   
			          
			       }else{ 
			          System.out.println("A05 add bank_no="+tmpBank_no);
			          A05DataList.add(tmpBank_no);
			          A05DataList.add(tmpBank_no_name);
			          A05DataList.add(A05Data);
			          A05DataListALL.add(A05DataList);//欲回傳的A05 data
			          A05DataList = new LinkedList();//單一機構的data
			          A05Data = new Properties();//A05.acc_code|amt				          
			          tmpBank_no = (String)((DataObject)dbData.get(i)).getValue("bank_no");
			          tmpBank_no_name = (String)((DataObject)dbData.get(i)).getValue("s_report_name");
			          System.out.println("now tmpbank_no="+tmpBank_no);
			          if(((DataObject)dbData.get(i)).getValue("acc_code") != null){//95.06.19 acc_code == null時,只會有一筆資料	            			            			
			            A05Data.setProperty((String)((DataObject)dbData.get(i)).getValue("acc_code"),(((DataObject)dbData.get(i)).getValue("amt")).toString());   
			            System.out.println("A05.set("+(String)((DataObject)dbData.get(i)).getValue("acc_code")+")="+(((DataObject)dbData.get(i)).getValue("amt")).toString());
			          }  			          
			       }			       
			   }			   			   
			   
			   //將最後一筆寫回A05DataListALL			      
			   System.out.println("A05 add last Bank_no="+tmpBank_no);	
			   A05DataList.add(tmpBank_no);
			   A05DataList.add(tmpBank_no_name);
			   A05DataList.add(A05Data);
			   A05DataListALL.add(A05DataList);			   			   
			}
    		
    		HashMap h = new HashMap();
    		h.put("A01",A01DataListALL);
    		h.put("A05",A05DataListALL);
            return h;
    }	    
    
    
    //取得查詢結果
    //100.01.11 fix sqlInjection/根據查詢年度.100年以後取得新縣市別.100年以前取得舊縣市別 
    //103.08.22 fix 103.01漁會套用新科目代號 
    private HashMap getQryResultA01A06(String bank_type_list,String S_YEAR,String S_MONTH,String unit,String upd_code,String bank_type,String tbank_no,String lguser_id){    	   
    		//查詢條件        		
    		String sqlCmd = "";
    		String rule_1_A01 = "";
    		String rule_1_A06 = "";
    		String rule_1_A04 = "";
    		String rule_2 = "";    		
    		String tmpbank_type="";
    		String tmpbank_no="";
    		List A01DataListALL = new LinkedList();
    		List A06DataListALL = new LinkedList();
    		List A04DataListALL = new LinkedList();
    		//100.01.11 add 查詢年度100年以前.縣市別不同===============================
  	    	String cd01_table = (Integer.parseInt(S_YEAR) < 100)?"cd01_99":""; 
  	    	String wlx01_m_year = (Integer.parseInt(S_YEAR) < 100)?"99":"100"; 
  	    	List paramList = new ArrayList(); 
  	    	List rule_1_A01_paramList = new ArrayList(); 
  	    	List rule_1_A06_paramList = new ArrayList(); 
  	    	List rule_1_A04_paramList = new ArrayList(); 
  	    	//=====================================================================    
    		if(lguser_id.equals("A111111111") || bank_type.equals("2") || bank_type.equals("1") || bank_type.equals("3")){//98.09.29 add 中央存保
    		    tmpbank_type = "Z";
			    tmpbank_no = "9999999";
		    }else{
		        tmpbank_type = bank_type;
				tmpbank_no = tbank_no;
			}	
			
            //A06===================================================================================================
            /*  select * from 
				(
 				  select NVL((BN01.BANK_NO || BN01.BANK_NAME), ' ')  as S_Report_Name,  
    					 BN01.BANK_NO, BN01.BANK_NAME, bn01.bank_type ,  
  						 nvl(wlx01.CENTER_NO,' ') as CENTER_NO, 
						 nvl(wlx01.M2_NAME,' ')   as M2_NAME, 
						 wlx01.CANCEL_no, wlx01.CANCEL_date,
						 nvl(cd01.hsien_id,' ') as  hsien_id,  
	    				 a06.acc_code,round(a06.amt/1,0),
	    				 a06.acc_range,
	    				 nvl(cd01.hsien_name,'OTHER')  as  hsien_name, 
						 cd01.FR001W_output_order     as  FR001W_output_order 
 				 from  (select * from cd01 where  cd01.hsien_id <> 'Y') cd01 
     					left join wlx01 on wlx01.hsien_id=cd01.hsien_id  
	 					left join (select BANK_NO, BANK_NAME,  bank_type  
	    						   from bn01 where bn01.bank_type in('6','7'))  bn01  
							 on wlx01.bank_no=bn01.bank_no  
 					    left join 
 	    						(select  bank_code,m_year,m_month,a06.acc_code,amt_total as amt,
		 								 decode(acc_code,'970000','01','120101','02','120102','03','120301','04','120302','05','120700','06','120401','07','120402','08','120501','09','120502','10','120601','11','120602','12','120603','13','120604','14','120600','15') as acc_range 
		 						 from  A06 
		 						 where M_Year =  94  and M_month =  8
		 						 and  a06.acc_code in('970000','120101','120102','120301','120302','120700','120401','120402','120501','120502','120601','120602','120603','120604','120600')		 
		 						 order by a06.bank_code,acc_range	
        						) A06
  	  						on bn01.bank_no = A06.Bank_Code
				)Temp_Output  where BANK_NO <> ' '
				order by  Bank_Type, FR001W_output_order, BANK_NO,acc_range) a06, -- A06資料
              //========================================================================================================
              //A01=====================================================================================================
              select * from 
			  (
 				select NVL((BN01.BANK_NO || BN01.BANK_NAME), ' ')  as S_Report_Name,  
    				   BN01.BANK_NO, BN01.BANK_NAME, bn01.bank_type ,  
  					   nvl(wlx01.CENTER_NO,' ') as CENTER_NO, 
					   nvl(wlx01.M2_NAME,' ')   as M2_NAME, 
					   wlx01.CANCEL_no, wlx01.CANCEL_date,
					   nvl(cd01.hsien_id,' ') as  hsien_id,  
	    			   a01.acc_code,round(a01.amt/1,0) as amt,
	    			   a01.acc_range,
	    			   nvl(cd01.hsien_name,'OTHER')  as  hsien_name, 
					   cd01.FR001W_output_order     as  FR001W_output_order 
 			   from  (select * from cd01 where  cd01.hsien_id <> 'Y') cd01 
     				  left join wlx01 on wlx01.hsien_id=cd01.hsien_id  
	 				  left join (select BANK_NO, BANK_NAME,  bank_type  
	    			  			 from bn01 where bn01.bank_type in('6','7'))  bn01  
						   on wlx01.bank_no=bn01.bank_no  
 					  left join 
 	    						(select  bank_code,m_year,m_month,a01.acc_code,amt,
		 								 decode(acc_code,'990000','01','120101','02','120102','03','120301','04','120302','05','120700','06','120401','07','120402','08','120501','09','120502','10','120601','11','120602','12','120603','13','120604','14','120600','15') as acc_range 
		 						 from  A01 
		 						 where M_Year =  94  and M_month =  8
		 						 and  a01.acc_code in('990000','120101','120102','120301','120302','120700','120401','120402','120501','120502','120601','120602','120603','120604','120600')		 
		 						 order by a01.bank_code,acc_range	
        						) A01
  	  						on bn01.bank_no = A01.Bank_Code
			)Temp_Output  where BANK_NO <> ' '
			order by  Bank_Type, FR001W_output_order, BANK_NO,acc_range
            //========================================================================================================================================
            //A04=====================================================================================================
            select * from 
			(
 			  select NVL((BN01.BANK_NO || BN01.BANK_NAME), ' ')  as S_Report_Name,  
    				 BN01.BANK_NO, BN01.BANK_NAME, bn01.bank_type ,  
  					 nvl(wlx01.CENTER_NO,' ') as CENTER_NO, 
					 nvl(wlx01.M2_NAME,' ')   as M2_NAME, 
					 wlx01.CANCEL_no, wlx01.CANCEL_date,
					 nvl(cd01.hsien_id,' ') as  hsien_id,  
	    			 a04.acc_code,round(a04.amt/1,0),
	    			 a04.acc_range,
	    			 nvl(cd01.hsien_name,'OTHER')  as  hsien_name, 
					 cd01.FR001W_output_order     as  FR001W_output_order 
 			 from  (select * from cd01 where  cd01.hsien_id <> 'Y') cd01 
     				left join wlx01 on wlx01.hsien_id=cd01.hsien_id  
	 				left join (select BANK_NO, BANK_NAME,  bank_type  
	    					   from bn01 where bn01.bank_type in('6','7'))  bn01  
						 on wlx01.bank_no=bn01.bank_no  
 				    left join 
 	    					(select  bank_code,m_year,m_month,a04.acc_code,amt,
		 							 decode(acc_code,'840740','01') as acc_range 
		 					 from  a04
		 					 where M_Year =  94  and M_month =  8
		 					 and  a04.acc_code in('840740')		 
		 					 order by a04.bank_code,acc_range	
        					) A04
  	  					on bn01.bank_no = A04.Bank_Code
			)Temp_Output  where BANK_NO <> ' '
			order by  Bank_Type, FR001W_output_order, BANK_NO,acc_range) a04, -- A04資料840740逾期放款總額
			//========================================================================================================================================
			*/
			
			rule_1_A01 = " select * from "
			  		   + " ("
 					   + "	select NVL((BN01.BANK_NO || BN01.BANK_NAME), ' ')  as S_Report_Name, " 
    				   + "  	   BN01.BANK_NO, BN01.BANK_NAME, bn01.bank_type ,  "
  					   + "		   nvl(wlx01.CENTER_NO,' ') as CENTER_NO, "
					   + "		   nvl(wlx01.M2_NAME,' ')   as M2_NAME, "
					   + "		   wlx01.CANCEL_no, wlx01.CANCEL_date,"
					   + "		   nvl(cd01.hsien_id,' ') as  hsien_id, " 
	    			   + "		   a01.acc_code,round(a01.amt/?,0) as amt,"
	    			   + "		   a01.acc_range,"
	    			   + "		   nvl(cd01.hsien_name,'OTHER')  as  hsien_name, "
					   + "		   cd01.FR001W_output_order     as  FR001W_output_order "
 			   		   + "	from  (select * from "+cd01_table+" cd01 where  cd01.hsien_id <> 'Y') cd01 "
     				   + "	left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id  "
	 				   + "	left join (select BANK_NO, BANK_NAME,  bank_type  "
	    			   + "			   from (select * from bn01 where m_year=?)bn01 where bn01.bank_type in(?))  bn01  "
					   + "		 on wlx01.bank_no=bn01.bank_no  "
 					   + "	left join "
 	    			   + "			(select  bank_code,m_year,m_month,a01.acc_code,amt,"; 	    				   
 	    	if(bank_type_list.equals("6")){//農會		   
		 	rule_1_A01 += "					 decode(acc_code,'990000','01','120101','02','120102','03','120301','04','120302',"
		 			   + "							'05','120700','06','120401','07','120402','08','120501','09','120502','10',"
		 			   + "							'120601','11','120602','12','120603','13','120604','14','120600','15') as acc_range ";
		 	}else if(bank_type_list.equals("7")){//漁會
		 		if(Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH) >= 10301){
		 	rule_1_A01 += "					 decode(acc_code,'990000','01','120101','02','120102','03','120301','04','120302',"
		 				+ "							'05','120700','06','120401','07','120402','08','120501','09','120502','10',"
		 				+ "						    '120601','11','120602','12','120603','13','120604','14','120600','15') as acc_range ";		 	
		 	    }else{		 	    	
		 	rule_1_A01 += "					 decode(acc_code,'990000','01','120101','02','120102','03','120401','04','120402',"
		 				+ "							'05','120700','06','120201','07','120202','08','120501','09','120502','10',"
		 				+ "						    '120601','11','120602','12','120603','13','120604','14','120600','15') as acc_range ";
		 		}		
		 	}		   
		 	rule_1_A01 += "			 from  A01 "
		 			   + "			 where M_Year = ? and M_month = ?";
		 	if(bank_type_list.equals("6")){//農會		   		   
		 	rule_1_A01 += "			 and  a01.acc_code in('990000','120101','120102','120301','120302','120700','120401','120402','120501','120502','120601','120602','120603','120604','120600') ";		 
		 	}else if(bank_type_list.equals("7")){//漁會		   		   
		 	rule_1_A01 += "			 and  a01.acc_code in('990000','120101','120102','120301','120302','120401','120402','120700','120201','120202','120501','120502','120601','120602','120603','120604','120600') ";		 
		 	}
		 	rule_1_A01 += "			 order by a01.bank_code,acc_range	"
        			   + "			) A01 "
  	  				   + "		 on bn01.bank_no = A01.Bank_Code "
					   + " )Temp_Output";
		 	rule_1_A01_paramList.add(unit);
			rule_1_A01_paramList.add(wlx01_m_year);//wlx01
			rule_1_A01_paramList.add(wlx01_m_year);//bn01  
			rule_1_A01_paramList.add(bank_type_list);
			rule_1_A01_paramList.add(S_YEAR);
			rule_1_A01_paramList.add(String.valueOf(Integer.parseInt(S_MONTH)));	
			   
			rule_1_A06 = " select * from "
					   + " ( "
 				  	   + "	select NVL((BN01.BANK_NO || BN01.BANK_NAME), ' ')  as S_Report_Name,  "
    				   + "		   BN01.BANK_NO, BN01.BANK_NAME, bn01.bank_type ,  "
  					   + "		   nvl(wlx01.CENTER_NO,' ') as CENTER_NO, "
					   + "		   nvl(wlx01.M2_NAME,' ')   as M2_NAME, "
					   + "		   wlx01.CANCEL_no, wlx01.CANCEL_date,"
					   + "	  	   nvl(cd01.hsien_id,' ') as  hsien_id, " 
	    			   + "	 	   a06.acc_code,round(a06.amt/?,0) as amt,"
	    			   + "		   a06.acc_range,"
	    			   + "		   nvl(cd01.hsien_name,'OTHER')  as  hsien_name, "
					   + "		   cd01.FR001W_output_order     as  FR001W_output_order "
 				 	   + "	from  (select * from "+cd01_table+" cd01 where  cd01.hsien_id <> 'Y') cd01 "
     				   + "	left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id " 
	 				   + "	left join (select BANK_NO, BANK_NAME,  bank_type  "
	    			   + "			   from (select * from bn01 where m_year=?)bn01 where bn01.bank_type in(?))  bn01 " 
					   + "		 on wlx01.bank_no=bn01.bank_no "
 					   + "  left join "
 	    			   + "			(select  bank_code,m_year,m_month,a06.acc_code,amt_total as amt,";
 	    	
 	    	if(bank_type_list.equals("6")){//農會			   
            rule_1_A06 += "					 decode(acc_code,'970000','01','120101','02','120102','03','120301','04','120302','05','120700'"
                       + "					 ,'06','120401','07','120402','08','120501','09','120502','10','120601','11','120602','12','120603'"
                       + "					 ,'13','120604','14','120600','15') as acc_range ";
            }else if(bank_type_list.equals("7")){//漁會			               
            rule_1_A06 += "					 decode(acc_code,'970000','01','120101','02','120102','03','120401','04','120402','05','120700',"
            			+ "                  '06','120201','07','120202','08','120501','09','120502','10','120601','11','120602','12','120603'"
            			+ "                  ,'13','120604','14','120600','15') as acc_range ";		 
            }
		 	rule_1_A06 += "			 from  A06 "
		 			   + "			 where M_Year = ? and M_month = ?";
			if(bank_type_list.equals("6")){//農會			   		 			   
		 	rule_1_A06 += "			 and  a06.acc_code in('970000','120101','120102','120301','120302','120700','120401','120402','120501','120502','120601','120602','120603','120604','120600')";		 
		 	}else if(bank_type_list.equals("7")){//漁會			               
		 	rule_1_A06 += "			 and  a06.acc_code in('970000','120101','120102','120401','120402','120700','120201','120202','120501','120502','120601','120602','120603','120604','120600')";		 
		 	}		   
		 	rule_1_A06 += "			 order by a06.bank_code,acc_range "	
        			   + "			) A06 "
  	  				   + "		 on bn01.bank_no = A06.Bank_Code "
					   + " )Temp_Output ";
		 	rule_1_A06_paramList.add(unit);
			rule_1_A06_paramList.add(wlx01_m_year);//wlx01
			rule_1_A06_paramList.add(wlx01_m_year);//bn01  
			rule_1_A06_paramList.add(bank_type_list);		   
			rule_1_A06_paramList.add(S_YEAR);
			rule_1_A06_paramList.add(String.valueOf(Integer.parseInt(S_MONTH)));
					   
			rule_1_A04 = " select * from "
			  		   + " ("
 					   + "	select NVL((BN01.BANK_NO || BN01.BANK_NAME), ' ')  as S_Report_Name, " 
    				   + "  	   BN01.BANK_NO, BN01.BANK_NAME, bn01.bank_type ,  "
  					   + "		   nvl(wlx01.CENTER_NO,' ') as CENTER_NO, "
					   + "		   nvl(wlx01.M2_NAME,' ')   as M2_NAME, "
					   + "		   wlx01.CANCEL_no, wlx01.CANCEL_date,"
					   + "		   nvl(cd01.hsien_id,' ') as  hsien_id, " 
	    			   + "		   a04.acc_code,round(a04.amt/?,0) as amt,"
	    			   + "		   a04.acc_range,"
	    			   + "		   nvl(cd01.hsien_name,'OTHER')  as  hsien_name, "
					   + "		   cd01.FR001W_output_order     as  FR001W_output_order "
 			   		   + "	from  (select * from "+cd01_table+" cd01 where  cd01.hsien_id <> 'Y') cd01 "
     				   + "	left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id  "
	 				   + "	left join (select BANK_NO, BANK_NAME,  bank_type  "
	    			   + "			   from (select * from bn01 where m_year=?)bn01 where bn01.bank_type in(?))  bn01  "
					   + "		 on wlx01.bank_no=bn01.bank_no  "
 					   + "	left join "
 	    			   + "			(select  bank_code,m_year,m_month,a04.acc_code,amt,"
 		 	           + "					 decode(acc_code,'840740','01') as acc_range "		 			   
	 	               + "			 from  A04 "
		 			   + "			 where M_Year = ? and M_month = ?"
		 			   + "			 and  a04.acc_code in('840740') "
		 	           + "			 order by a04.bank_code,acc_range "
        			   + "			) A04 "
  	  				   + "		 on bn01.bank_no = A04.Bank_Code "
					   + " )Temp_Output";		
			rule_1_A04_paramList.add(unit);
			rule_1_A04_paramList.add(wlx01_m_year);//wlx01
			rule_1_A04_paramList.add(wlx01_m_year);//bn01  
			rule_1_A04_paramList.add(bank_type_list);		   
			rule_1_A04_paramList.add(S_YEAR);
			rule_1_A04_paramList.add(String.valueOf(Integer.parseInt(S_MONTH)));		   
			rule_2 = " where 	BANK_NO <> ' '		and "
				   + " ( "
				   + " (? = 'Z') or "
			       + " ((? = '6' or ? = '7') and  ? = BANK_NO) or "
			       + " (? = '8' and ? = CENTER_NO) or"
			       + " (? = 'B' and ? = M2_NAME) "
				   + " )    and "
				   + " ( "
				   + " (?= 'ALL') or "
 			       + " ((?= '6') and (Bank_Type = '6')) or "
				   + " ((?= '7') and (Bank_Type = '7')) "
				   + " )"                                                                               
				   + " order by  Bank_Type, BANK_NO,acc_range";
			paramList.add(tmpbank_type);
			paramList.add(tmpbank_type);
			paramList.add(tmpbank_type);
			paramList.add(tmpbank_no);
			paramList.add(tmpbank_type);
			paramList.add(tmpbank_no);
			paramList.add(tmpbank_type);
			paramList.add(tmpbank_no);
			paramList.add(bank_type_list);
			paramList.add(bank_type_list);
			paramList.add(bank_type_list);			
			
			for(int i=0;i<paramList.size();i++){
	            rule_1_A01_paramList.add(paramList.get(i));		
			}
			for(int i=0;i<paramList.size();i++){
	            rule_1_A06_paramList.add(paramList.get(i));		
			}	 
			for(int i=0;i<paramList.size();i++){
	            rule_1_A04_paramList.add(paramList.get(i));		
			}
			 	   
			List dbData = DBManager.QueryDB_SQLParam( rule_1_A01 + rule_2,rule_1_A01_paramList,"cancel_date,amt");   
			String tmpBank_no="";
			String tmpBank_no_name="";
			Properties A01Data = new Properties();		
			List A01DataList = new LinkedList();
			
			if(dbData != null && dbData.size() > 0){
			   for(int i=0;i<dbData.size();i++){			 			        
			       if(i==0){
			           tmpBank_no = (String)((DataObject)dbData.get(i)).getValue("bank_no");     			       			       			       
			           tmpBank_no_name = (String)((DataObject)dbData.get(i)).getValue("s_report_name");
			       }    
			       if(((String)((DataObject)dbData.get(i)).getValue("bank_no")).equals(tmpBank_no)){//與前一個bank_no相同時,儲存acc_code,amt			          			       			       
			          if(((DataObject)dbData.get(i)).getValue("acc_code") != null){//95.06.19 acc_code == null時,只會有一筆資料	            			            			             			             
			            A01Data.setProperty((String)((DataObject)dbData.get(i)).getValue("acc_code"),(((DataObject)dbData.get(i)).getValue("amt")).toString());			       			       			            
			            System.out.println("A01.set("+(String)((DataObject)dbData.get(i)).getValue("acc_code")+")="+(((DataObject)dbData.get(i)).getValue("amt")).toString());
			          }  
			       }else{ 
			          System.out.println("A01 add Bank_no="+tmpBank_no);			          
			          A01DataList.add(tmpBank_no);
			          A01DataList.add(tmpBank_no_name);
			          A01DataList.add(A01Data);
			          A01DataListALL.add(A01DataList);//欲回傳的A01 data
			          A01DataList = new LinkedList();//單一機構的data
			          A01Data = new Properties();//A01.acc_code|amt	
			          tmpBank_no = (String)((DataObject)dbData.get(i)).getValue("bank_no");
			          tmpBank_no_name = (String)((DataObject)dbData.get(i)).getValue("s_report_name");
			          if(((DataObject)dbData.get(i)).getValue("acc_code") != null){//95.06.19 acc_code == null時,只會有一筆資料	            			            			             			             
			             A01Data.setProperty((String)((DataObject)dbData.get(i)).getValue("acc_code"),(((DataObject)dbData.get(i)).getValue("amt")).toString());   
			             System.out.println("A01.set("+(String)((DataObject)dbData.get(i)).getValue("acc_code")+")="+(((DataObject)dbData.get(i)).getValue("amt")).toString());
			          }   			          
			          System.out.println("now tmpbank_no="+tmpBank_no);
			       }			       
			   }				   			   
			   //將最後一筆寫回A01DataListALL
			   System.out.println("A01 add last Bank_no="+tmpBank_no);		   
			   A01DataList.add(tmpBank_no);
			   A01DataList.add(tmpBank_no_name);
			   A01DataList.add(A01Data);
			   A01DataListALL.add(A01DataList);			   			      
			}
			dbData = DBManager.QueryDB_SQLParam( rule_1_A06 + rule_2,rule_1_A06_paramList,"cancel_date,amt");   
			tmpBank_no="";	
			tmpBank_no_name="";		
			Properties A06Data = new Properties();		
			List A06DataList = new LinkedList();
			if(dbData != null && dbData.size() > 0){
			   for(int i=0;i<dbData.size();i++){			 			             
			       if(i==0){
			          tmpBank_no = (String)((DataObject)dbData.get(i)).getValue("bank_no");     			       			       			       
			          tmpBank_no_name = (String)((DataObject)dbData.get(i)).getValue("s_report_name");
			       }   
			       if(((String)((DataObject)dbData.get(i)).getValue("bank_no")).equals(tmpBank_no)){//與前一個bank_no相同時,儲存acc_code,amt			          			       			       
			          if(((DataObject)dbData.get(i)).getValue("acc_code") != null){//95.06.19 acc_code == null時,只會有一筆資料	            			            			             			             
			             A06Data.setProperty((String)((DataObject)dbData.get(i)).getValue("acc_code"),(((DataObject)dbData.get(i)).getValue("amt")).toString());			       			       			            
			             System.out.println("A06.set("+(String)((DataObject)dbData.get(i)).getValue("acc_code")+")="+(((DataObject)dbData.get(i)).getValue("amt")).toString());
			          }   
			       }else{ 
			          System.out.println("A06 add bank_no="+tmpBank_no);
			          A06DataList.add(tmpBank_no);
			          A06DataList.add(tmpBank_no_name);
			          A06DataList.add(A06Data);
			          A06DataListALL.add(A06DataList);//欲回傳的A05 data
			          A06DataList = new LinkedList();//單一機構的data
			          A06Data = new Properties();//A05.acc_code|amt	
			          tmpBank_no = (String)((DataObject)dbData.get(i)).getValue("bank_no");
			          tmpBank_no_name = (String)((DataObject)dbData.get(i)).getValue("s_report_name");
			          if(((DataObject)dbData.get(i)).getValue("acc_code") != null){//95.06.19 acc_code == null時,只會有一筆資料	            			            			             			             
			             A06Data.setProperty((String)((DataObject)dbData.get(i)).getValue("acc_code"),(((DataObject)dbData.get(i)).getValue("amt")).toString());   
			             System.out.println("A06.set("+(String)((DataObject)dbData.get(i)).getValue("acc_code")+")="+(((DataObject)dbData.get(i)).getValue("amt")).toString());
			          }   
			          System.out.println("now tmpbank_no="+tmpBank_no);
			       }			       
			   }			   
			   //將最後一筆寫回A06DataListALL			   
			   System.out.println("A06 add last bank_no="+tmpBank_no);
			   A06DataList.add(tmpBank_no);
			   A06DataList.add(tmpBank_no_name);
			   A06DataList.add(A06Data);
			   A06DataListALL.add(A06DataList);			   
			}
			
			dbData = DBManager.QueryDB_SQLParam( rule_1_A04 + rule_2,rule_1_A04_paramList,"cancel_date,amt");   
			tmpBank_no="";	
			tmpBank_no_name="";		
			Properties A04Data = new Properties();		
			List A04DataList = new LinkedList();
			if(dbData != null && dbData.size() > 0){
			   for(int i=0;i<dbData.size();i++){			 			             
			       if(i==0){
			          tmpBank_no = (String)((DataObject)dbData.get(i)).getValue("bank_no");     			       			       			       
			          tmpBank_no_name = (String)((DataObject)dbData.get(i)).getValue("s_report_name");
			       }   
			       if(((String)((DataObject)dbData.get(i)).getValue("bank_no")).equals(tmpBank_no)){//與前一個bank_no相同時,儲存acc_code,amt			          			       			       
			          if(((DataObject)dbData.get(i)).getValue("acc_code") != null){//95.06.19 acc_code == null時,只會有一筆資料	            			            			             			             
			             A04Data.setProperty((String)((DataObject)dbData.get(i)).getValue("acc_code"),(((DataObject)dbData.get(i)).getValue("amt")).toString());			       			       			            
			             System.out.println("A04.set("+(String)((DataObject)dbData.get(i)).getValue("acc_code")+")="+(((DataObject)dbData.get(i)).getValue("amt")).toString());
			          }   
			       }else{ 
			          System.out.println("A04 add bank_no="+tmpBank_no);
			          A04DataList.add(tmpBank_no);
			          A04DataList.add(tmpBank_no_name);
			          A04DataList.add(A04Data);
			          A04DataListALL.add(A04DataList);//欲回傳的A04 data
			          A04DataList = new LinkedList();//單一機構的data
			          A04Data = new Properties();//A04.acc_code|amt	
			          tmpBank_no = (String)((DataObject)dbData.get(i)).getValue("bank_no");
			          tmpBank_no_name = (String)((DataObject)dbData.get(i)).getValue("s_report_name");
			          if(((DataObject)dbData.get(i)).getValue("acc_code") != null){//95.06.19 acc_code == null時,只會有一筆資料	            			            			             			             
			             A04Data.setProperty((String)((DataObject)dbData.get(i)).getValue("acc_code"),(((DataObject)dbData.get(i)).getValue("amt")).toString());   
			             System.out.println("A04.set("+(String)((DataObject)dbData.get(i)).getValue("acc_code")+")="+(((DataObject)dbData.get(i)).getValue("amt")).toString());
			          }   
			          System.out.println("now tmpbank_no="+tmpBank_no);
			       }			       
			   }			   
			   //將最後一筆寫回A04DataListALL			   
			   System.out.println("A04 add last bank_no="+tmpBank_no);
			   A04DataList.add(tmpBank_no);
			   A04DataList.add(tmpBank_no_name);
			   A04DataList.add(A04Data);
			   A04DataListALL.add(A04DataList);			   
			}
			
    		System.out.println("A01DataListALL.size="+A01DataListALL.size());
    		System.out.println("A06DataListALL.size="+A06DataListALL.size());
    		System.out.println("A04DataListALL.size="+A04DataListALL.size());
    		HashMap h = new HashMap();
    		h.put("A01",A01DataListALL);
    		h.put("A06",A06DataListALL);
    		h.put("A04",A04DataListALL);
            return h;
    }
    
    //取得查詢結果(目前無使用)
    //100.01.11 fix sqlInjection/根據查詢年度.100年以後取得新縣市別.100年以前取得舊縣市別 
    private List getQryResultA04A06(String bank_type_list,String S_YEAR,String S_MONTH,String unit,String upd_code,String bank_type,String tbank_no,String lguser_id){    	   
    		//查詢條件        		
    		String sqlCmd = "";
    		String rule_1_A04 = "";
    		String rule_1_A06 = "";
    		String rule_1_ALL = "";
    		String rule_2 = "";    		
    		String tmpbank_type="";
    		String tmpbank_no="";    		
    		//100.01.11 add 查詢年度100年以前.縣市別不同===============================
  	    	String cd01_table = (Integer.parseInt(S_YEAR) < 100)?"cd01_99":""; 
  	    	String wlx01_m_year = (Integer.parseInt(S_YEAR) < 100)?"99":"100"; 
  	    	List paramList = new ArrayList(); 
  	    	//=====================================================================    
    	
    		if(lguser_id.equals("A111111111") || bank_type.equals("2") || bank_type.equals("1")){
    		    tmpbank_type = "Z";
			    tmpbank_no = "9999999";
		    }else{
		        tmpbank_type = bank_type;
				tmpbank_no = tbank_no;
			}	
			
            //===================================================================================================
            /*  
			select a04.s_report_name,
	   			   a04.bank_no as a04_bank_no,
	   			   a06.bank_no as a06_bank_no,	  
	   			   a04.a04_84073xamt,
	   			   a06.a06_960500amt 
		    from A04,A06
			where a04.bank_no = a06.bank_no 
			and ((a06.a06_960500amt is not null and a04.a04_84073xamt is not null) and (a06.a06_960500amt > a04.a04_84073xamt)) -- A01/A06逾放及放款額檢核有不符
			and ((a06.a06_960500amt is not null and a04.a04_84073xamt is not null) and (a06.a06_960500amt =0 and  a04.a04_84073xamt = 0)) -- A01/A06逾放款合計金額均為0
			and (a06.a06_960500amt is null or a04.a04_84073xamt is null) -- A01/A06逾放款合計金額尚未申報,                               
			group by a04.s_report_name,a04.bank_no,a06.bank_no,a04.a04_84073xamt,a06.a06_960500amt
			order by a04.s_report_name,a04.bank_no,a06.bank_no	
			//A04
 			select * from 
			(
 			  select NVL((BN01.BANK_NO || BN01.BANK_NAME), ' ')  as S_Report_Name,  
    		 		 BN01.BANK_NO, BN01.BANK_NAME, bn01.bank_type ,  
  				 	 nvl(wlx01.CENTER_NO,' ') as CENTER_NO, 
					 nvl(wlx01.M2_NAME,' ')   as M2_NAME, 
					 wlx01.CANCEL_no, wlx01.CANCEL_date,
					 nvl(cd01.hsien_id,' ') as  hsien_id,
        		 	 round(a04.amt/1,0) as a04_84073xamt,
	    			 nvl(cd01.hsien_name,'OTHER')  as  hsien_name, 
					 cd01.FR001W_output_order     as  FR001W_output_order 
 			  from  (select * from cd01 where  cd01.hsien_id <> 'Y') cd01 
     		  left join wlx01 on wlx01.hsien_id=cd01.hsien_id  
	 		  left join (select BANK_NO, BANK_NAME,  bank_type  
	    			     from bn01 where bn01.bank_type in('6','7'))  bn01  
					on wlx01.bank_no=bn01.bank_no  
 			  left join       
 	    			    (select  bank_code,m_year,m_month,
		 				 sum(decode(acc_code,'840731',amt,'840732',amt,'840733',amt,'840734',amt,'840735',amt,0)) as amt  
		 				 from  A04 
		 				 where M_Year =  95  and M_month =  5
		 				 group by a04.bank_code,m_year,m_month		 
		 				 order by a04.bank_code,m_year,m_month		 
        				) A04
  	  			   on bn01.bank_no = A04.Bank_Code
 			)Temp_Output  where BANK_NO <> ' '	
  			order by  Bank_Type,BANK_NO,FR001W_output_order
			//A06
			select * from
			(
			  select NVL((BN01.BANK_NO || BN01.BANK_NAME), ' ')  as S_Report_Name,  
    				 BN01.BANK_NO, BN01.BANK_NAME, bn01.bank_type ,  
  					 nvl(wlx01.CENTER_NO,' ') as CENTER_NO, 
					 nvl(wlx01.M2_NAME,' ')   as M2_NAME, 
					 wlx01.CANCEL_no, wlx01.CANCEL_date,
					 nvl(cd01.hsien_id,' ') as  hsien_id,  
	    			 a06.acc_code,
	    			 round(a06.amt/1,0) as a06_960500amt,       
	    			 nvl(cd01.hsien_name,'OTHER')  as  hsien_name, 
					 cd01.FR001W_output_order     as  FR001W_output_order 
 			 from  (select * from cd01 where  cd01.hsien_id <> 'Y') cd01 
    		 left join wlx01 on wlx01.hsien_id=cd01.hsien_id  
	 		 left join (select BANK_NO, BANK_NAME,  bank_type  
	    			    from bn01 where bn01.bank_type in('6','7'))  bn01  
					on wlx01.bank_no=bn01.bank_no
 			 left join 	  
 	    			   (select  bank_code,m_year,m_month,acc_code,amt_total as amt  
		  				from  A06 
		 			    where M_Year =  95  and M_month =  2		 
		 				and acc_code='960500' 		 
		 				order by a06.bank_code,m_year,m_month		 
        				) A06
			  	   on bn01.bank_no = A06.Bank_Code
			 )Temp_Output  where BANK_NO <> ' '	 
			 order by  Bank_Type,BANK_NO,FR001W_output_order
            //========================================================================================================================================
			*/
			rule_1_A04 = " select * from "
					   + " ( "
 			  		   + " select NVL((BN01.BANK_NO || BN01.BANK_NAME), ' ')  as S_Report_Name,  "
    		 		   + " 		  BN01.BANK_NO, BN01.BANK_NAME, bn01.bank_type , " 
  				 	   + "	  	  nvl(wlx01.CENTER_NO,' ') as CENTER_NO, "
					   + "		  nvl(wlx01.M2_NAME,' ')   as M2_NAME, "
					   + "		  wlx01.CANCEL_no, wlx01.CANCEL_date, "
					   + "		  nvl(cd01.hsien_id,' ') as  hsien_id, "
        		 	   + "	  	  round(a04.amt/?,0) as a04_84073xamt, "
	    			   + "		  nvl(cd01.hsien_name,'OTHER')  as  hsien_name, "
					   + "		  cd01.FR001W_output_order     as  FR001W_output_order "
 			  		   + "	from  (select * from "+cd01_table+" cd01 where  cd01.hsien_id <> 'Y') cd01 "
     		  		   + "	left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id  "
	 		  		   + "	left join (select BANK_NO, BANK_NAME,  bank_type  "
	    			   + " 			   from (select * from bn01 where m_year=?)bn01 where bn01.bank_type in('6','7'))  bn01 " 
					   + "       on wlx01.bank_no=bn01.bank_no  "
 			  		   + "	left join (select  bank_code,m_year,m_month,"
		 			   + "		  	   sum(decode(acc_code,'840731',amt,'840732',amt,'840733',amt,'840734',amt,'840735',amt,0)) as amt  "
		 			   + "			   from  A04 "
		 			   + "			   where M_Year = ? and M_month = ?"
		               + "             group by a04.bank_code,m_year,m_month "		 
		 			   + "			   order by a04.bank_code,m_year,m_month "		 
        			   + "			  ) A04 "
  	  			   	   + "	     on bn01.bank_no = A04.Bank_Code "
 					   + "	)Temp_Output  where BANK_NO <> ' '	";
 					   
 			rule_1_A06 = " select * from "
					   + " ( "
 			  		   + " select NVL((BN01.BANK_NO || BN01.BANK_NAME), ' ')  as S_Report_Name,  "
    		 		   + " 		  BN01.BANK_NO, BN01.BANK_NAME, bn01.bank_type , " 
  				 	   + "	  	  nvl(wlx01.CENTER_NO,' ') as CENTER_NO, "
					   + "		  nvl(wlx01.M2_NAME,' ')   as M2_NAME, "
					   + "		  wlx01.CANCEL_no, wlx01.CANCEL_date, "
					   + "		  nvl(cd01.hsien_id,' ') as  hsien_id, "
        		 	   + "	  	  a06.acc_code,"
        		 	   + "        round(a06.amt/?,0) as a06_960500amt, "
	    			   + "		  nvl(cd01.hsien_name,'OTHER')  as  hsien_name, "
					   + "		  cd01.FR001W_output_order     as  FR001W_output_order "
 			  		   + "	from  (select * from "+cd01_table+" cd01 where  cd01.hsien_id <> 'Y') cd01 "
     		  		   + "	left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id  "
	 		  		   + "	left join (select BANK_NO, BANK_NAME,  bank_type  "
	    			   + " 			   from (select * from bn01 where m_year=?)bn01 where bn01.bank_type in('6','7'))  bn01 " 
					   + "       on wlx01.bank_no=bn01.bank_no  "
					   + "	left join (select  bank_code,m_year,m_month,acc_code,amt_total as amt " 
		  			   + "		 	   from  A06 "
		 			   + "			   where M_Year = ? and M_month = ?"
		 			   + "		   	   and acc_code='960500'" 		 
		 			   + "			   order by a06.bank_code,m_year,m_month "		 
        			   + "			  ) A06 "
			  	   	   + "		 on bn01.bank_no = A06.Bank_Code "
 			  		   + "	)Temp_Output  where BANK_NO <> ' '	";
 			rule_1_ALL = " and "
				   	   + " ( "
				   	   + " (? = 'Z') or "
			       	   + " ((? = '6' or ? = '7') and  ? = BANK_NO) or "
			       	   + " (? = '8' and ? = CENTER_NO) or"
			       	   + " (? = 'B' and ? = M2_NAME) "
				   	   + " )    and "
				   	   + " ( "
				   	   + " (?= 'ALL') or "
 			       	   + " ((?= '6') and (Bank_Type = '6')) or "
				   	   + " ((?= '7') and (Bank_Type = '7')) "
				   	   + " )"           
 					   + " order by  Bank_Type,BANK_NO,FR001W_output_order ";	   
 			rule_2 = " where a04.bank_no = a06.bank_no ";                                                                    				   
			
			
			
			if(upd_code.equals("1")){//A06/A04逾放金額檢核有不符  	   
			   rule_2 += " and ((a06.a06_960500amt is not null and a04.a04_84073xamt is not null) and (a06.a06_960500amt > a04.a04_84073xamt)) "; 
  			}else if(upd_code.equals("2")){//A06/A04逾放金額均為0
  			   rule_2 += " and ((a06.a06_960500amt is not null and a04.a04_84073xamt is not null) and (a06.a06_960500amt =0 and  a04.a04_84073xamt = 0))";
  			}else if(upd_code.equals("3")){//A06/A04逾放金額尚未申報  	    	
  			   rule_2 += " and (a06.a06_960500amt is null or a04.a04_84073xamt is null) ";
  			}	
  			rule_2 += " group by a04.s_report_name,a04.bank_no,a06.bank_no,a04.a04_84073xamt,a06.a06_960500amt "
			       + "  order by a04.s_report_name,a04.bank_no,a06.bank_no ";	
			
			//rule_1_A04
			paramList.add(unit);
			paramList.add(wlx01_m_year);//wlx01
			paramList.add(wlx01_m_year);//bn01  
			paramList.add(S_YEAR);
			paramList.add(String.valueOf(Integer.parseInt(S_MONTH)));
			//rule_1_ALL
			paramList.add(tmpbank_type);	   		   
			paramList.add(tmpbank_type);
			paramList.add(tmpbank_type);
			paramList.add(tmpbank_no);
			paramList.add(tmpbank_type);
			paramList.add(tmpbank_no);
			paramList.add(tmpbank_type);
			paramList.add(tmpbank_no);
			paramList.add(bank_type_list);
			paramList.add(bank_type_list);
			paramList.add(bank_type_list);
			//rule_1_A06
			paramList.add(unit);
			paramList.add(wlx01_m_year);//wlx01
			paramList.add(wlx01_m_year);//bn01  
			paramList.add(S_YEAR);
			paramList.add(String.valueOf(Integer.parseInt(S_MONTH)));
			//rule_1_ALL
			paramList.add(tmpbank_type);	   		   
			paramList.add(tmpbank_type);
			paramList.add(tmpbank_type);
			paramList.add(tmpbank_no);
			paramList.add(tmpbank_type);
			paramList.add(tmpbank_no);
			paramList.add(tmpbank_type);
			paramList.add(tmpbank_no);
			paramList.add(bank_type_list);
			paramList.add(bank_type_list);
			paramList.add(bank_type_list);
			
			sqlCmd = " select a04.s_report_name, "
	   			   + " 		  a04.bank_no as a04_bank_no,"
	   			   + " 		  a06.bank_no as a06_bank_no, "	  
	   			   + " 		  a04.a04_84073xamt,"
	   			   + " 		  a06.a06_960500amt "
		           + " from (" + rule_1_A04 + rule_1_ALL + ")a04,(" + rule_1_A06 + rule_1_ALL + ")a06"
				   + rule_2;
			List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"a04_84073xamt,a06_960500amt");   			
			System.out.println("A04A06.size()="+dbData.size());
			return dbData;
    }
    
    //100.01.11 fix sqlInjection/根據查詢年度.100年以後取得新縣市別.100年以前取得舊縣市別 
    private HashMap getQryResultA01A02A99(String bank_type_list,String S_YEAR,String S_MONTH,String unit,String upd_code,String bank_type,String tbank_no,String lguser_id){    	   
    		//查詢條件        		
    		String sqlCmd = "";
    		String rule_1_L = "";
    		String rule_1_R = "";
    		String rule_2 = "";    		
    		String tmpbank_type="";
    		String tmpbank_no="";
    		List A02DataListALL = new LinkedList();
    		List A99DataListALL = new LinkedList();
    		//100.01.11 add 查詢年度100年以前.縣市別不同===============================
  	    	String cd01_table = (Integer.parseInt(S_YEAR) < 100)?"cd01_99":""; 
  	    	String wlx01_m_year = (Integer.parseInt(S_YEAR) < 100)?"99":"100"; 
  	    	List paramList = new ArrayList(); 
  	    	List rule_1_L_paramList = new ArrayList(); 
  	    	List rule_1_R_paramList = new ArrayList(); 
  	    	//=====================================================================   
    		if(lguser_id.equals("A111111111") || bank_type.equals("2") || bank_type.equals("1") || bank_type.equals("3")){//98.09.29 add 中央存保
    		    tmpbank_type = "Z";
			    tmpbank_no = "9999999";
		    }else{
		        tmpbank_type = bank_type;
				tmpbank_no = tbank_no;
			}	
			
           /*
              //跨表檢核左式=================================================================================================== 
              select * from 
              ( 
              	select NVL((BN01.BANK_NO || BN01.BANK_NAME), ' ')  as S_Report_Name, 
              		   BN01.BANK_NO, BN01.BANK_NAME, bn01.bank_type ,
              		   nvl(wlx01.CENTER_NO,' ') as CENTER_NO,
              	   	   nvl(wlx01.M2_NAME,' ')   as M2_NAME, 
              		   wlx01.CANCEL_no, wlx01.CANCEL_date,
              	   	   nvl(cd01.hsien_id,' ') as  hsien_id,
              		   a01.acc_code,amt,
              		   a01.acc_range,
              		   nvl(cd01.hsien_name,'OTHER')  as  hsien_name,
              	  	   cd01.FR001W_output_order     as  FR001W_output_order
              	from  (select * from cd01 where  cd01.hsien_id <> 'Y') cd01 
              	left join wlx01 on wlx01.hsien_id=cd01.hsien_id  
              	left join (select BANK_NO, BANK_NAME,  bank_type 
              			   from bn01 where bn01.bank_type in('6','7'))  bn01 
              		 on wlx01.bank_no=bn01.bank_no
              	left join (
                             select bank_code,m_year,m_month,a02.acc_code,round(decode(acc_code,'990630',amt/2,amt/1),0) as amt,
                                    decode(acc_code,'990120','01','990130','02','990210','03','990630','07','990810','08','990510','11') as acc_range 
                             from  A02
                             where M_Year =  94  and M_month =  8 --and bank_code='6230012'
                             and  a02.acc_code in('990120','990130','990210','990630','990810','990510')		
                             union 
                             select bank_code,m_year,m_month,a04.acc_code,round(amt/1,0) as amt,
                                    decode(acc_code,'840760','10') as acc_range 
                             from  A04
                             where M_Year =  94  and M_month =  8 --and bank_code='6230012'
                             and  a04.acc_code in('840760')	
                             union 
                             select bank_code,m_year,m_month,a99.acc_code,round(amt/1,0) as amt,
                                    decode(acc_code,'992150','12') as acc_range 
                             from  A99
                             where M_Year =  94  and M_month =  8 --and bank_code='6230012'
                             and  a99.acc_code in('992150')	 
                             union
                             select bank_code,m_year,m_month,a01.acc_code,round(amt/1,0) as amt,
                                    decode(acc_code,'220000','05','990000','09') as acc_range 
                             from  A01
                             where M_Year =  94  and M_month =  8 --and bank_code='6230012'
                             and  a01.acc_code in('220000','990000')
                             union
                             select bank_code,m_year,m_month,'120000' as acc_code,
                             	   round(sum(decode(a01.acc_code,'120000',amt, '120800',amt,'150300',amt,0)) /1,0)- 
                                    round(sum(decode(acc_code,'120700',amt))/1,0) as amt,	   
                                    '06' as acc_range 
                             from  A01
                             where M_Year =  94  and M_month =  8 --and bank_code='6230012'
                             and  a01.acc_code in('120000','120800','150300','120700')
                             group by bank_code,m_year,m_month
                             order by bank_code,acc_range
                            )a01 on bn01.bank_no = A01.Bank_Code
              ) Temp_Output where BANK_NO <> ' '	
              order by  Bank_Type,BANK_NO,FR001W_output_order
              //========================================================================================================
              //跨表檢核右式 			
              select * from 
              ( 
              	select NVL((BN01.BANK_NO || BN01.BANK_NAME), ' ')  as S_Report_Name, 
              		   BN01.BANK_NO, BN01.BANK_NAME, bn01.bank_type ,
              		   nvl(wlx01.CENTER_NO,' ') as CENTER_NO,
              	   	   nvl(wlx01.M2_NAME,' ')   as M2_NAME, 
              		   wlx01.CANCEL_no, wlx01.CANCEL_date,
              	   	   nvl(cd01.hsien_id,' ') as  hsien_id,
              		   a01.acc_code,amt,
              		   a01.acc_range,
              		   nvl(cd01.hsien_name,'OTHER')  as  hsien_name,
              	  	   cd01.FR001W_output_order     as  FR001W_output_order
              	from  (select * from cd01 where  cd01.hsien_id <> 'Y') cd01 
              	left join wlx01 on wlx01.hsien_id=cd01.hsien_id  
              	left join (select BANK_NO, BANK_NAME,  bank_type 
              			   from bn01 where bn01.bank_type in('6','7'))  bn01 
              		 on wlx01.bank_no=bn01.bank_no
              	left join (		
                             select bank_code,m_year,m_month,a01.acc_code,round(decode(acc_code,'220900',amt/2,amt/1),0) as amt,
                                    decode(acc_code,'140000','02','120700','03','220900','07') as acc_range 
                             from  A01
                             where M_Year =  98  and M_month =  12 --and bank_code='6230012'
                             and  a01.acc_code in('140000','120700','220900')
                             union
                             select bank_code,m_year,m_month,'310000' as acc_code,
                             	   round(sum(decode(acc_code,'310000',amt,'320000',amt,0)) /1,0)-  //98.12.28 A02.990120=A01.310000+A01.320000-A01.310800
                                    round(sum(decode(acc_code,'310800',amt))/1,0) as amt,
                                    '01' as acc_range 
                             from  A01
                             where M_Year =  98  and M_month =  12 --and bank_code='6230012'
                             and  a01.acc_code in('310000','320000','310800')
                             group by bank_code,m_year,m_month
                             union
                             select bank_code,m_year,m_month,'990511' as acc_code,sum(amt) as amt,
                                    '11' as acc_range 
                             from  A02
                             where M_Year =  98  and M_month =  12 --and bank_code='6230012'
                             and  a02.acc_code in('990511','990512')
                             group by bank_code,m_year,m_month
                             union
                             select bank_code,m_year,m_month,a02.acc_code,round(amt/1,0) as amt,
                                    decode(acc_code,'990220','04','990510','12') as acc_range 
                             from  A02
                             where M_Year =  98  and M_month =  12 --and bank_code='6230012'
                             and  a02.acc_code in('990220','990510')
                             union
                             select bank_code,m_year,m_month,'992130' as acc_code,
                             	   round(sum(amt) /1,0) as amt,
                                    '05' as acc_range 
                             from  (select * from A99 
							        where M_Year =  98  and M_month =  12 and bank_code='5030019'
                                    and  a99.acc_code = '992130'
								    union 
									select * from A02
									where M_Year =  98  and M_month =  12 and bank_code='5030019'
                                    and  a02.acc_code in('990420','990620')
							 )A99
                             group by bank_code,m_year,m_month
                             union
                             select bank_code,m_year,m_month,'992140' as acc_code,
                             	   round(sum(amt) /1,0) as amt,
                                    '06' as acc_range 
                             from  (select * from A99 
							        where M_Year =  98  and M_month =  12 and bank_code='5030019'
                                    and  a99.acc_code = '992140'
								    union 
									select * from A02
									where M_Year =  98  and M_month =  12 and bank_code='5030019'
                                    and  a02.acc_code in('990410','990610')
							 )A99                            
                             group by bank_code,m_year,m_month
                             union
                             select bank_code,m_year,m_month,'992510' as acc_code,
                             	   round(sum(amt) /1,0) as amt,
                                    '09' as acc_range 
                             from  A99
                             where M_Year =  98  and M_month =  12 --and bank_code='6230012'
                             and  a99.acc_code in('992510','992520','992530','992540')
                             group by bank_code,m_year,m_month
                             union
                             select bank_code,m_year,m_month,'992610' as acc_code,
                             	   round(sum(amt) /1,0) as amt,
                                    '10' as acc_range 
                             from  A99
                             where M_Year =  98  and M_month =  12 --and bank_code='6230012'
                             and  a99.acc_code in('992610','992620','992630','992640')
                             group by bank_code,m_year,m_month
                             union
                             select bank_code,m_year,m_month,a99.acc_code,round(amt/1,0) as amt,
                                    decode(acc_code,'992550','13','992650','14') as acc_range 
                             from  A99
                             where M_Year =  98  and M_month =  12 --and bank_code='6230012'
                             and  a99.acc_code in('992550','992650')
                             order by bank_code,acc_range
              			 )a01 on bn01.bank_no = A01.Bank_Code
              ) Temp_Output where BANK_NO <> ' '	
              order by  Bank_Type,BANK_NO,FR001W_output_order
            //========================================================================================================================================
			*/
			
			rule_1_L = " select * from "
              	   + " ( "
               	   + "	select NVL((BN01.BANK_NO || BN01.BANK_NAME), ' ')  as S_Report_Name," 
                   + "		   BN01.BANK_NO, BN01.BANK_NAME, bn01.bank_type ," 
                   + "		   nvl(wlx01.CENTER_NO,' ') as CENTER_NO,"
              	   + "	   	   nvl(wlx01.M2_NAME,' ')   as M2_NAME, "
              	   + "		   wlx01.CANCEL_no, wlx01.CANCEL_date,"
              	   + "	   	   nvl(cd01.hsien_id,' ') as  hsien_id,"
              	   + "		   a01.acc_code, round(amt/?,0) as amt,"
              	   + "		   a01.acc_range,"
              	   + "		   nvl(cd01.hsien_name,'OTHER')  as  hsien_name,"
              	   + "	  	   cd01.FR001W_output_order     as  FR001W_output_order"
               	   + "	from  (select * from "+cd01_table+" cd01 where  cd01.hsien_id <> 'Y') cd01 "
                   + "	left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id " 
              	   + "	left join (select BANK_NO, BANK_NAME,  bank_type " 
              	   + "			   from (select * from bn01 where m_year=?)bn01 where bn01.bank_type in('6','7'))  bn01 " 
              	   + "		 on wlx01.bank_no=bn01.bank_no"
               	   + "	left join "
               	   + "			 ( select bank_code,m_year,m_month,a02.acc_code,"
               	   + "			   round(decode(acc_code,'990630',amt/2,amt/1),0) as amt,"
                   + "   		   decode(acc_code,'990120','01','990130','02','990210','03','990630','07','990810','08','990510','11') as acc_range " 
               	   + " 			   from  A02"
                   + "		       where M_Year = ? and M_month = ?"
                   + "             and  a02.acc_code in('990120','990130','990210','990630','990810','990510') "		
                   + "             union "
                   + " 			   select bank_code,m_year,m_month,a04.acc_code,round(amt/1,0) as amt,"
                   + "		       decode(acc_code,'840760','10') as acc_range "
               	   + "			   from  A04 "
                   + "		       where M_Year = ? and M_month = ?"
                   + "			   and  a04.acc_code in('840760') "
                   + "			   union "
                   + "			   select bank_code,m_year,m_month,a99.acc_code,round(amt/1,0) as amt,"
                   + "			   decode(acc_code,'992150','12') as acc_range "
               	   + "			   from  A99"
                   + "		       where M_Year = ? and M_month = ?"
               	   + "			   and  a99.acc_code in('992150') "
               	   + "			   union "
               	   + "			   select bank_code,m_year,m_month,a01.acc_code,round(amt/1,0) as amt,"
                   + "			   decode(acc_code,'220000','05','990000','09') as acc_range "
               	   + "			   from  A01 "
                   + "		       where M_Year = ? and M_month = ?"
               	   + "			   and  a01.acc_code in('220000','990000')"
               	   + "			   union "
               	   + "			   select bank_code,m_year,m_month,'120000' as acc_code,"
               	   + "			   round(sum(decode(a01.acc_code,'120000',amt, '120800',amt,'150300',amt,0)) /1,0)- "
                   + "			   round(sum(decode(acc_code,'120700',amt))/1,0) as amt,'06' as acc_range "
               	   + "			   from  A01 "
                   + "		       where M_Year = ? and M_month = ?"
                   + "			   and  a01.acc_code in('120000','120800','150300','120700') "
               	   + "			   group by bank_code,m_year,m_month"
               	   + " 			   order by bank_code,acc_range"
                   + "		     ) A01 "
                   + "		 on bn01.bank_no = A01.Bank_Code"
              	   + " ) Temp_Output  ";
          rule_1_L_paramList.add(unit);
          rule_1_L_paramList.add(wlx01_m_year);//wlx01
		  rule_1_L_paramList.add(wlx01_m_year);//bn01  	 	
		  for(int i=1;i<=5;i++){
		      rule_1_L_paramList.add(S_YEAR);   
		      rule_1_L_paramList.add(String.valueOf(Integer.parseInt(S_MONTH)));
		  }  	 	   
		  		  
          rule_1_R = " select * from "
              	   + " ( "
               	   + "	select NVL((BN01.BANK_NO || BN01.BANK_NAME), ' ')  as S_Report_Name," 
                   + "		   BN01.BANK_NO, BN01.BANK_NAME, bn01.bank_type ," 
                   + "		   nvl(wlx01.CENTER_NO,' ') as CENTER_NO,"
              	   + "	   	   nvl(wlx01.M2_NAME,' ')   as M2_NAME, "
              	   + "		   wlx01.CANCEL_no, wlx01.CANCEL_date,"
              	   + "	   	   nvl(cd01.hsien_id,' ') as  hsien_id,"
              	   + "		   a01.acc_code, round(amt/?,0) as amt,"
              	   + "		   a01.acc_range,"
              	   + "		   nvl(cd01.hsien_name,'OTHER')  as  hsien_name,"
              	   + "	  	   cd01.FR001W_output_order     as  FR001W_output_order"
               	   + "	from  (select * from "+cd01_table+" cd01 where  cd01.hsien_id <> 'Y') cd01 "
                   + "	left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id " 
              	   + "	left join (select BANK_NO, BANK_NAME,  bank_type " 
              	   + "			   from (select * from bn01 where m_year=?)bn01 where bn01.bank_type in('6','7'))  bn01 " 
              	   + "		 on wlx01.bank_no=bn01.bank_no"
               	   + "	left join "
               	   + "		     ( select bank_code,m_year,m_month,a01.acc_code,round(decode(acc_code,'220900',amt/2,amt/1),0) as amt,"
                   + "		       decode(acc_code,'140000','02','120700','03','220900','07') as acc_range "
               	   + "			   from  A01"
                   + "		       where M_Year = ? and M_month = ?"
               	   + "			   and  a01.acc_code in('140000','120700','220900')"
               	   + "			   union "
               	   + "			   select bank_code,m_year,m_month,'310000' as acc_code,";
           if(bank_type_list.equals("7") && (Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH) >= 10301)){
           	 rule_1_R += "			round(sum(decode(acc_code,'310000',amt,'320000',amt,0)) /1,0) as amt, ";   //102.08.22 漁會無310800 A02.990120=A01.310000+A01.320000               	                 
           }else{
           	 rule_1_R += "			round(sum(decode(acc_code,'310000',amt,'320000',amt,0)) /1,0)- "   //98.12.28 A02.990120=A01.310000+A01.320000-A01.310800
               	      + "			round(sum(decode(acc_code,'310800',amt))/1,0) as amt, "; 
           }		
            rule_1_R += "		   '01' as acc_range " 
               	   + "			   from  A01"
                   + "		       where M_Year = ? and M_month = ?"
               	   + "			   and  a01.acc_code in('310000','320000','310800')"
               	   + "			   group by bank_code,m_year,m_month"
               	   + "			   union "
               	   + "			   select bank_code,m_year,m_month,'990511' as acc_code,sum(amt) as amt,'11' as acc_range "
               	   + "			   from  A02"
                   + "		       where M_Year = ? and M_month = ?"
               	   + " 			   and  a02.acc_code in('990511','990512')"
               	   + "			   group by bank_code,m_year,m_month"
               	   + "			   union"
               	   + "			   select bank_code,m_year,m_month,a02.acc_code,round(amt/1,0) as amt,"
                   + "			   decode(acc_code,'990220','04','990510','12') as acc_range "
               	   + "			   from  A02"
                   + "		       where M_Year = ? and M_month = ?"
               	   + "			   and  a02.acc_code in('990220','990510')"
               	   + "			   union "
               	   + "			   select bank_code,m_year,m_month,'992130' as acc_code,"
               	   + "			   round(sum(amt) /1,0) as amt,'05' as acc_range "
               	   + "			   from  (select * from A99 "
				   + "		              where M_Year = ? and M_month = ?"
                   + "                    and  a99.acc_code = '992130'"
				   + "			          union "
				   + "			          select m_year,m_month,bank_code,acc_code,amt from A02 "
				   + "		              where M_Year = ? and M_month = ?"
                   + "                    and  a02.acc_code in('990420','990620')"
				   + "			         )A99"
               	   + "			   group by bank_code,m_year,m_month"
               	   + "			   union"
               	   + "			   select bank_code,m_year,m_month,'992140' as acc_code,"
               	   + "			   round(sum(amt) /1,0) as amt,'06' as acc_range "
               	   + "			   from  (select * from A99 "
				   + "		              where M_Year = ? and M_month = ?"
                   + "                    and  a99.acc_code = '992140'"
				   + "			          union "
				   + "			          select m_year,m_month,bank_code,acc_code,amt from A02 "
				   + "		              where M_Year = ? and M_month = ?"
                   + "                    and  a02.acc_code in('990410','990610')"
				   + "			         )A99"
               	   + "			   group by bank_code,m_year,m_month"
                   + "			   union"
               	   + "			   select bank_code,m_year,m_month,'992510' as acc_code,"
               	   + "			   round(sum(amt) /1,0) as amt,'09' as acc_range "
               	   + "			   from  A99"
                   + "		       where M_Year = ? and M_month = ?"
               	   + "			   and  a99.acc_code in('992510','992520','992530','992540')"
               	   + "			   group by bank_code,m_year,m_month"
               	   + "			   union"
              	   + "			   select bank_code,m_year,m_month,'992610' as acc_code,"
               	   + "			   round(sum(amt) /1,0) as amt,'10' as acc_range "
               	   + "			   from  A99"
                   + "		       where M_Year = ? and M_month = ?"
                   + "			   and  a99.acc_code in('992610','992620','992630','992640')"
               	   + "			   group by bank_code,m_year,m_month"
               	   + "			   union "
               	   + "			   select bank_code,m_year,m_month,a99.acc_code,round(amt/1,0) as amt,"
                   + "		       decode(acc_code,'992550','13','992650','14') as acc_range "
               	   + "			   from  A99"
				   + "		       where M_Year = ? and M_month = ?"
               	   + "			   and  a99.acc_code in('992550','992650')"
               	   + "			   order by bank_code,acc_range"
                   + "		     ) A01 "
                   + "		 on bn01.bank_no = A01.Bank_Code"
              	   + " ) Temp_Output  ";
            rule_1_R_paramList.add(unit);
            rule_1_R_paramList.add(wlx01_m_year);//wlx01
		    rule_1_R_paramList.add(wlx01_m_year);//bn01
		    for(int i=1;i<=11;i++){
		       rule_1_R_paramList.add(S_YEAR);   
		       rule_1_R_paramList.add(String.valueOf(Integer.parseInt(S_MONTH)));
		    }  	 	   
		    	    
		    
			rule_2 = " where 	BANK_NO <> ' '		and "
				   + " ( "
				   + " (? = 'Z') or "
			       + " ((? = '6' or ? = '7') and  ? = BANK_NO) or "
			       + " (? = '8' and ? = CENTER_NO) or"
			       + " (? = 'B' and ? = M2_NAME) "
				   + " )    and "
				   + " ( "
				   + " (?= 'ALL') or "
 			       + " ((?= '6') and (Bank_Type = '6')) or "
				   + " ((?= '7') and (Bank_Type = '7')) "
				   + " )"                                                                               
				   + " order by  Bank_Type,BANK_NO,acc_range,FR001W_output_order ";    
			//rule_1_L使用	   
			rule_1_L_paramList.add(tmpbank_type);	   
			rule_1_L_paramList.add(tmpbank_type);
			rule_1_L_paramList.add(tmpbank_type);
			rule_1_L_paramList.add(tmpbank_no);
			rule_1_L_paramList.add(tmpbank_type);
			rule_1_L_paramList.add(tmpbank_no);
			rule_1_L_paramList.add(tmpbank_type);
			rule_1_L_paramList.add(tmpbank_no);
			rule_1_L_paramList.add(bank_type_list);
			rule_1_L_paramList.add(bank_type_list);
			rule_1_L_paramList.add(bank_type_list);
			
			//rule_1_R使用
			rule_1_R_paramList.add(tmpbank_type);	   
			rule_1_R_paramList.add(tmpbank_type);
			rule_1_R_paramList.add(tmpbank_type);
			rule_1_R_paramList.add(tmpbank_no);
			rule_1_R_paramList.add(tmpbank_type);
			rule_1_R_paramList.add(tmpbank_no);
			rule_1_R_paramList.add(tmpbank_type);
			rule_1_R_paramList.add(tmpbank_no);
			rule_1_R_paramList.add(bank_type_list);
			rule_1_R_paramList.add(bank_type_list);
			rule_1_R_paramList.add(bank_type_list);
			
			List dbData = DBManager.QueryDB_SQLParam( rule_1_L + rule_2,rule_1_L_paramList,"cancel_date,amt");   
			String tmpBank_no="";
			String tmpBank_no_name="";
			Properties A02Data = new Properties();		
			List A02DataList = new LinkedList();
			if(dbData != null && dbData.size() > 0){
			   for(int i=0;i<dbData.size();i++){			 			          
			       if(i==0){
			           tmpBank_no = (String)((DataObject)dbData.get(i)).getValue("bank_no");     			       			       			       
			           tmpBank_no_name = (String)((DataObject)dbData.get(i)).getValue("s_report_name");
			       }    
			       if(((String)((DataObject)dbData.get(i)).getValue("bank_no")).equals(tmpBank_no)){//與前一個bank_no相同時,儲存acc_code,amt			          			       			       
			       	  if(((DataObject)dbData.get(i)).getValue("acc_code") != null){//95.06.19 acc_code == null時,只會有一筆資料	            			            			             			             
			             A02Data.setProperty((String)((DataObject)dbData.get(i)).getValue("acc_code"),(((DataObject)dbData.get(i)).getValue("amt")).toString());			       			       			            
			             System.out.println("A02.set("+(String)((DataObject)dbData.get(i)).getValue("acc_code")+")="+(((DataObject)dbData.get(i)).getValue("amt")).toString());
			          }   
			         
			       }else{ 
			          System.out.println("A02 add Bank_no="+tmpBank_no);
			          A02DataList.add(tmpBank_no);
			          A02DataList.add(tmpBank_no_name);
			          A02DataList.add(A02Data);
			          A02DataListALL.add(A02DataList);//欲回傳的A02 data
			          A02DataList = new LinkedList();//單一機構的data
			          A02Data = new Properties();//A02.acc_code|amt				          
			          tmpBank_no = (String)((DataObject)dbData.get(i)).getValue("bank_no");
			          tmpBank_no_name = (String)((DataObject)dbData.get(i)).getValue("s_report_name");
			          System.out.println("now tmpbank_no="+tmpBank_no);
			          if(((DataObject)dbData.get(i)).getValue("acc_code") != null){//95.06.19 acc_code == null時,只會有一筆資料	            			            			             			             
			             A02Data.setProperty((String)((DataObject)dbData.get(i)).getValue("acc_code"),(((DataObject)dbData.get(i)).getValue("amt")).toString());   
			             System.out.println("A02.set("+(String)((DataObject)dbData.get(i)).getValue("acc_code")+")="+(((DataObject)dbData.get(i)).getValue("amt")).toString());
			          }
			       }			       
			   }			   
			   //將最後一筆寫回A02DataListALL
			   System.out.println("A02 add last Bank_no="+tmpBank_no);				  
			   A02DataList.add(tmpBank_no);
			   A02DataList.add(tmpBank_no_name);
			   A02DataList.add(A02Data);
			   A02DataListALL.add(A02DataList);			   			   
			}
			dbData = DBManager.QueryDB_SQLParam( rule_1_R + rule_2,rule_1_R_paramList,"cancel_date,amt");   
			tmpBank_no="";	
			tmpBank_no_name="";		
			Properties A99Data = new Properties();		
			List A99DataList = new LinkedList();
			if(dbData != null && dbData.size() > 0){
			   for(int i=0;i<dbData.size();i++){			 			         
			       if(i==0){
			          tmpBank_no = (String)((DataObject)dbData.get(i)).getValue("bank_no");     			       			       			       
			          tmpBank_no_name = (String)((DataObject)dbData.get(i)).getValue("s_report_name");
			       }   
			       if(((String)((DataObject)dbData.get(i)).getValue("bank_no")).equals(tmpBank_no)){//與前一個bank_no相同時,儲存acc_code,amt			          			       			       
			          if(((DataObject)dbData.get(i)).getValue("acc_code") != null){//95.06.19 acc_code == null時,只會有一筆資料
			             A99Data.setProperty((String)((DataObject)dbData.get(i)).getValue("acc_code"),(((DataObject)dbData.get(i)).getValue("amt")).toString());			       			       			            
			             System.out.println("A99.set("+(String)((DataObject)dbData.get(i)).getValue("acc_code")+")="+(((DataObject)dbData.get(i)).getValue("amt")).toString());
			          }   
			          
			       }else{ 
			          System.out.println("A99 add bank_no="+tmpBank_no);
			          A99DataList.add(tmpBank_no);
			          A99DataList.add(tmpBank_no_name);
			          A99DataList.add(A99Data);
			          A99DataListALL.add(A99DataList);//欲回傳的A99 data
			          A99DataList = new LinkedList();//單一機構的data
			          A99Data = new Properties();//A99.acc_code|amt				          
			          tmpBank_no = (String)((DataObject)dbData.get(i)).getValue("bank_no");
			          tmpBank_no_name = (String)((DataObject)dbData.get(i)).getValue("s_report_name");
			          System.out.println("now tmpbank_no="+tmpBank_no);
			          if(((DataObject)dbData.get(i)).getValue("acc_code") != null){//95.06.19 acc_code == null時,只會有一筆資料	            			            			
			            A99Data.setProperty((String)((DataObject)dbData.get(i)).getValue("acc_code"),(((DataObject)dbData.get(i)).getValue("amt")).toString());   
			            System.out.println("A99.set("+(String)((DataObject)dbData.get(i)).getValue("acc_code")+")="+(((DataObject)dbData.get(i)).getValue("amt")).toString());
			          }  			          
			       }			       
			   }			   			   
			   
			   //將最後一筆寫回A99DataListALL			      
			   System.out.println("A99 add last Bank_no="+tmpBank_no);	
			   A99DataList.add(tmpBank_no);
			   A99DataList.add(tmpBank_no_name);
			   A99DataList.add(A99Data);
			   A99DataListALL.add(A99DataList);			   			   
			}
    		
    		HashMap h = new HashMap();
    		h.put("A02",A02DataListALL);
    		h.put("A99",A99DataListALL);
            return h;
    }	
    private HashMap getQryResultA10(String bank_type_list,String S_YEAR,String S_MONTH,String unit,String upd_code,String bank_type,String tbank_no,String lguser_id){    	   
    		//查詢條件        		
    		String sqlCmd = "";
    		String rule_L = "";
    		String rule_R = "";
    		//String rule_2 = "";    		
    		String tmpbank_type="";
    		String tmpbank_no="";
    		List DataListALL_L = new LinkedList();
    		List DataListALL_R = new LinkedList();
    		//100.01.11 add 查詢年度100年以前.縣市別不同===============================
  	    	String cd01_table = (Integer.parseInt(S_YEAR) < 100)?"cd01_99":""; 
  	    	String wlx01_m_year = (Integer.parseInt(S_YEAR) < 100)?"99":"100"; 
  	    	List paramList = new ArrayList(); 
  	    	List rule_L_paramList = new ArrayList(); 
  	    	List rule_R_paramList = new ArrayList(); 
  	    	//=====================================================================    
    		if(lguser_id.equals("A111111111") || bank_type.equals("2") || bank_type.equals("1") || bank_type.equals("3")){//98.09.29 add 中央存保
    		    tmpbank_type = "Z";
			    tmpbank_no = "9999999";
		    }else{
		        tmpbank_type = bank_type;
				tmpbank_no = tbank_no;
			}	
			
            //左式公式SQL===================================================================================================
            /*  select * from 
				( 
				    select NVL((BN01.BANK_NO || BN01.BANK_NAME), ' ')  as S_Report_Name, 
				           BN01.BANK_NO, BN01.BANK_NAME, bn01.bank_type ,
				           nvl(wlx01.CENTER_NO,' ') as CENTER_NO,
				           nvl(wlx01.M2_NAME,' ')   as M2_NAME, 
				           wlx01.CANCEL_no, wlx01.CANCEL_date,
				           nvl(cd01.hsien_id,' ') as  hsien_id,
				           a01.acc_code,
				round(amt/'UI.金額單位',0) as amt,
				           a01.acc_range,
				           nvl(cd01.hsien_name,'OTHER')  as  hsien_name,
				           cd01.FR001W_output_order     as  FR001W_output_order
				    from  (select * from cd01 where  cd01.hsien_id <> 'Y') cd01 
				    left join (select * from wlx01 where m_year=100)wlx01 on wlx01.hsien_id=cd01.hsien_id  
				    left join (select BANK_NO, BANK_NAME,  bank_type 
				               from bn01 where m_year=100 and bn01.bank_type in('6','7'))bn01 on wlx01.bank_no=bn01.bank_no
				    left join (        
				               select bank_code,m_year,m_month,
				                     a01.acc_code,round(decode(acc_code,'220900',amt/2,amt/1),0) as amt,
				                      decode(acc_code,'190000','01','220000','02','220900','04','140000','05','990000','06') as acc_range 
				               from  A01
				               where M_Year = 'UI.查詢年度 ex:100'  and M_month =  'UI.查詢月份 ex:1' 
				               and  a01.acc_code in('190000','220000','220900','140000','990000')
				               union
				               select bank_code,m_year,m_month,'120000' as acc_code,
				                      round(sum(decode(acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /1,0)-  
				                      round(sum(decode(acc_code,'120700',amt))/1,0) as amt,
				                      '03' as acc_range 
				               from  A01
				               where M_Year = 'UI.查詢年度 ex:100'  and M_month = 'UI.查詢月份 ex:1' 
				               and  a01.acc_code in('120000','120800','150300','120700')
				               group by bank_code,m_year,m_month
				               union
				               select bank_code,m_year,m_month,'120800' as acc_code,
				                      round(sum(decode(acc_code,'120800',amt,'150300',amt,0)) /1,0) as amt,
				                      '13' as acc_range 
				               from  A01
				               where M_Year = 'UI.查詢年度 ex:100'  and M_month = 'UI.查詢月份 ex:1' 
				               and  a01.acc_code in('120800','150300')
				               group by bank_code,m_year,m_month
				               union
				               select bank_code,m_year,m_month,'840760' as acc_code,sum(amt) as amt,
				                      '07' as acc_range 
				               from  A04
				               where M_Year =  'UI.查詢年度 ex:100'  and M_month =  'UI.查詢月份 ex:1'
				               and  a04.acc_code in('840760') 
				               group by bank_code,m_year,m_month
				               union
				               select bank_code,m_year,m_month,a02.acc_code,round(amt/1,0) as amt,
				                      decode(acc_code,'990510','08','990610','12') as acc_range 
				               from  A02
				               where M_Year = 'UI.查詢年度 ex:100'  and M_month =  'UI.查詢月份 ex:1' 
				               and  a02.acc_code in('990510','990610')
				               union
				               select bank_code,m_year,m_month,a99.acc_code,round(amt/1,0) as amt,
				                      decode(acc_code,'992150','09') as acc_range 
				               from  A99
				               where M_Year = 'UI.查詢年度 ex:100'  and M_month =  'UI.查詢月份 ex:1' 
				               and  a99.acc_code in('992150')
				               union
				               select bank_code,m_year,m_month,a99.acc_code,round(amt/1,0) as amt,
				                      decode(acc_code,'992150','10') as acc_range 
				               from  A99
				               where M_Year = 'UI.查詢年度 ex:100'  and M_month = 'UI.查詢月份 ex:1'  
				               and  a99.acc_code in('992150')
				               union
				               select bank_code,m_year,m_month,a99.acc_code,round(amt/1,0) as amt,
				                      decode(acc_code,'992150','11') as acc_range 
				               from  A99
				               where M_Year = 'UI.查詢年度 ex:100'  and M_month = 'UI.查詢月份 ex:1' 
				               and  a99.acc_code in('992150')
				               union
				               select bank_code,m_year,m_month,'990021' as acc_code,
				                      round(sum(loan1_baddebt+loan2_baddebt+loan3_baddebt+loan4_baddebt)/'1',0) as loan_baddebt_sum,
				                      '14' as acc_range 
				               from A10
				               where M_Year = 'UI.查詢年度 ex:100'  and M_month = 'UI.查詢月份 ex:1'
				               group by bank_code,m_year,m_month
				               union
				               select bank_code,m_year,m_month,'990002' as acc_code,
				                      round(sum(loan2_amt+loan3_amt+loan4_amt)/'1',0) as loan_sum,
				                      '15' as acc_range 
				               from A10
				               where M_Year = 'UI.查詢年度 ex:100'  and M_month = 'UI.查詢月份 ex:1' 
				               group by bank_code,m_year,m_month
				               union
				               select bank_code,m_year,m_month,'990001' as acc_code,
				                      round(sum(loan1_amt+loan2_amt+loan3_amt+loan4_amt)/'1',0) as loan_sum,
				                      '16' as acc_range 
				               from A10
				               where M_Year ='UI.查詢年度 ex:100'  and M_month = 'UI.查詢月份 ex:1'
				               group by bank_code,m_year,m_month
				               union
				               select bank_code,m_year,m_month,'990006' as acc_code,
				                      round(sum(invest1_amt+invest2_amt+invest3_amt+invest4_amt)/'1',0) as loan_sum,
				                      '17' as acc_range 
				               from A10
				               where M_Year ='UI.查詢年度 ex:100'  and M_month ='UI.查詢月份 ex:1'
				               group by bank_code,m_year,m_month
				               union
				               select bank_code,m_year,m_month,'990005' as acc_code,
				                      round(sum(loan1_amt+loan2_amt+loan3_amt+loan4_amt)/'1',0) as loan_sum,
				                      '18' as acc_range 
				               from A10
				               where M_Year ='UI.查詢年度 ex:100'  and M_month ='UI.查詢月份 ex:1' 
				               group by bank_code,m_year,m_month
				             )a01 on bn01.bank_no = A01.Bank_Code
				) Temp_Output 
				where BANK_NO <> '  '
				AND (   ('Z' = 'Z')
				     OR ( ('Z' = '6' OR 'Z' = '7') AND '9999999' = BANK_NO)
				     OR ('Z' = '8' AND '9999999' = CENTER_NO)
				     OR ('Z' = 'B' AND '9999999' = M2_NAME))
				AND (   ('6' = 'ALL')
				     OR ( ('6' = '6') AND (Bank_Type = '6'))
				     OR ( ('6' = '7') AND (Bank_Type = '7')))
				order by  Bank_Type,BANK_NO,acc_range,FR001W_output_order

              //========================================================================================================
              //右式公式SQL:=====================================================================================================
              select * from 
				( 
				    select NVL((BN01.BANK_NO || BN01.BANK_NAME), ' ')  as S_Report_Name, 
				           BN01.BANK_NO, BN01.BANK_NAME, bn01.bank_type ,
				           nvl(wlx01.CENTER_NO,' ') as CENTER_NO,
				           nvl(wlx01.M2_NAME,' ')   as M2_NAME, 
				           wlx01.CANCEL_no, wlx01.CANCEL_date,
				           nvl(cd01.hsien_id,' ') as  hsien_id,
				           a01.acc_code,
				round(amt/'UI.金額單位',0) as amt,
				           a01.acc_range,
				           nvl(cd01.hsien_name,'OTHER')  as  hsien_name,
				           cd01.FR001W_output_order    as  FR001W_output_order
				    from  (select * from cd01 where  cd01.hsien_id <> 'Y') cd01 
				    left join (select * from wlx01 where m_year=100)wlx01 on wlx01.hsien_id=cd01.hsien_id  
				    left join (select BANK_NO, BANK_NAME,  bank_type 
				               from bn01 where m_year=100 and bn01.bank_type in('6','7'))  bn01 
				         on wlx01.bank_no=bn01.bank_no
				    left join (      
				                select bank_code,m_year,m_month,
				                       acc_code,round(amt/1,0) as amt,
				                      decode(acc_code,'400000','01') as acc_range 
				               from  A01
				               where M_Year = 'UI.查詢年度 ex:100'  and M_month ='UI.查詢月份 ex:1'
				               and  a01.acc_code in('400000')
				               union
				               select bank_code,m_year,m_month,'992130' as acc_code,
				                      round(sum(amt) /1,0) as amt,
				                      '02' as acc_range 
				               from  (select * from A99 
				                      where M_Year = 'UI.查詢年度 ex:100'  and M_month ='UI.查詢月份 ex:1'
				                      and  a99.acc_code = '992130'
				                      union 
				                      select m_year,m_month,bank_code,acc_code,amt from A02
				                      where M_Year = 'UI.查詢年度 ex:100'  and M_month ='UI.查詢月份 ex:1'
				                      and  a02.acc_code in('990420','990620')
				               )A99
				               group by bank_code,m_year,m_month
				               union
				               select bank_code,m_year,m_month,'992140' as acc_code,
				                      round(sum(amt) /1,0) as amt,
				                      '03' as acc_range 
				               from  (select * from A99 
				                      where M_Year = 'UI.查詢年度 ex:100'  and M_month ='UI.查詢月份 ex:1'
				                      and  a99.acc_code = '992140'
				                      union 
				                      select m_year,m_month,bank_code,acc_code,amt from A02
				                      where M_Year = 'UI.查詢年度 ex:100'  and M_month ='UI.查詢月份 ex:1'
				                      and  a02.acc_code in('990410','990610')
				               )A99
				               group by bank_code,m_year,m_month                            
				               union
				               select bank_code,m_year,m_month,a02.acc_code,
				                      round(decode(acc_code,'990630',amt/2,amt/1),0) as amt,
				                      decode(acc_code,'990630','04','990810','05','990510','09','990612','12') as acc_range 
				               from  A02
				               where M_Year ='UI.查詢年度 ex:100'  and M_month ='UI.查詢月份 ex:1'
				               and  a02.acc_code in('990630','990810','990510','990612')
				               union
				               select bank_code,m_year,m_month,'992510' as acc_code,
				                      round(sum(decode(acc_code,'992510',amt,'992520',amt,'992530',amt,'992540',amt,0)) /1,0) as amt,
				                      '06' as acc_range 
				               from  A99
				               where M_Year ='UI.查詢年度 ex:100'  and M_month ='UI.查詢月份 ex:1'
				               and  a99.acc_code in('992510','992520','992530','992540')
				               group by bank_code,m_year,m_month
				               union
				               select bank_code,m_year,m_month,'992610' as acc_code,
				                      round(sum(decode(acc_code,'992610',amt,'992620',amt,'992630',amt,'992640',amt,0)) /1,0) as amt,
				                      '07' as acc_range 
				               from  A99
				               where M_Year ='UI.查詢年度 ex:100'  and M_month ='UI.查詢月份 ex:1'
				               and  a99.acc_code in('992610','992620','992630','992640')
				               group by bank_code,m_year,m_month
				               union
				               select bank_code,m_year,m_month,'990511' as acc_code,
				                      round(sum(decode(acc_code,'990511',amt,'990512',amt,0)) /1,0) as amt,
				                      '08' as acc_range 
				               from  A02
				               where M_Year ='UI.查詢年度 ex:100'  and M_month ='UI.查詢月份 ex:1'
				               and  a02.acc_code in('990511','990512')
				               group by bank_code,m_year,m_month
				               union
				               select bank_code,m_year,m_month,a99.acc_code,round(amt/1,0) as amt,
				                      decode(acc_code,'992550','10','992650','11') as acc_range 
				               from  A99
				               where M_Year = 'UI.查詢年度 ex:100'  and M_month ='UI.查詢月份 ex:1'
				               and  a99.acc_code in('992550','992650')
				               union
				               select bank_code,m_year,m_month,'990025' as acc_code,
				                      round(sum(loan1_baddebt+loan2_baddebt+loan3_baddebt+loan4_baddebt)/'1',0) as loan_baddebt_sum,
				                      '13' as acc_range 
				               from A10
				               where M_Year ='UI.查詢年度 ex:100'  and M_month ='UI.查詢月份 ex:1' 
				               group by bank_code,m_year,m_month
				               union
				               select bank_code,m_year,m_month,'990025' as acc_code,
				                      round(sum(loan1_baddebt+loan2_baddebt+loan3_baddebt+loan4_baddebt)/'1',0) as loan_baddebt_sum,
				                      '14' as acc_range 
				               from A10
				               where M_Year ='UI.查詢年度 ex:100'  and M_month ='UI.查詢月份 ex:1'
				               group by bank_code,m_year,m_month
				               union
				               select bank_code,m_year,m_month,'840760' as acc_code,
				                      round(sum(amt) /1,0) as amt,
				                      '15' as acc_range 
				               from  (select * from A04 
				                      where M_Year ='UI.查詢年度 ex:100'  and M_month ='UI.查詢月份 ex:1'
				                      and  a04.acc_code = '840760'
				                      union 
				                      select m_year,m_month,bank_code,acc_code,amt from A01
				                      where M_Year ='UI.查詢年度 ex:100'  and M_month ='UI.查詢月份 ex:1'
				                      and  a01.acc_code in('990000')
				               )A04
				               group by bank_code,m_year,m_month      
				               union
				               select bank_code,m_year,m_month,'990005' as acc_code,
				                      round(sum(loan1_amt+loan2_amt+loan3_amt+loan4_amt)/'1',0) as loan_sum,
				                      '16' as acc_range 
				               from A10
				               where M_Year ='UI.查詢年度 ex:100'  and M_month ='UI.查詢月份 ex:1'
				               group by bank_code,m_year,m_month
				               union
				               select bank_code,m_year,m_month,'990010' as acc_code,
				                      round(sum(invest1_amt+invest2_amt+invest3_amt+invest4_amt)/'1',0) as loan_sum,
				                      '17' as acc_range 
				               from A10
				               where M_Year ='UI.查詢年度 ex:100'  and M_month = 'UI.查詢月份 ex:1'
				               group by bank_code,m_year,m_month
				               union
				               select bank_code,m_year,m_month,'120000' as acc_code,
				                        round(sum(decode(acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /1,0)-  
				                        round(sum(decode(acc_code,'990611',amt))/1,0) as amt,
				                       '18' as acc_range 
				              from (
				              select * from  A01
				              where M_Year ='UI.查詢年度 ex:100'  and M_month ='UI.查詢月份 ex:1' 
				               and  a01.acc_code in('120000','120800','150300')
				              union 
				              select m_year,m_month,bank_code,acc_code,amt from A02
				              where M_Year ='UI.查詢年度 ex:100'  and M_month ='UI.查詢月份 ex:1'
				               and  a02.acc_code in('990611')
				              )A01 
				              group by bank_code,m_year,m_month
				              union
				              select bank_code,m_year,m_month,'840740' as acc_code,sum(amt) as amt,
				                       '19' as acc_range --第06項檢核,增加檢核19 
				              from  A04
				              where M_Year ='UI.查詢年度 ex:100'  and M_month ='UI.查詢月份 ex:1'
				              and  a04.acc_code in('840740') 
				              group by bank_code,m_year,m_month
				             )a01 on bn01.bank_no = A01.Bank_Code
				) Temp_Output 
				where BANK_NO <> '  '
				AND (   ('Z' = 'Z')
				     OR ( ('Z' = '6' OR 'Z' = '7') AND '9999999' = BANK_NO)
				     OR ('Z' = '8' AND '9999999' = CENTER_NO)
				     OR ('Z' = 'B' AND '9999999' = M2_NAME))
				AND (   ('6' = 'ALL')
				     OR ( ('6' = '6') AND (Bank_Type = '6'))
				     OR ( ('6' = '7') AND (Bank_Type = '7')))
				order by  Bank_Type,BANK_NO,acc_range,FR001W_output_order

            //========================================================================================================================================
			*/
			
			rule_L = " select * from " 
					+ " ( " 
					+ "     select NVL((BN01.BANK_NO || BN01.BANK_NAME), ' ')  as S_Report_Name, " 
					+ "            BN01.BANK_NO, BN01.BANK_NAME, bn01.bank_type , "
					+ "            nvl(wlx01.CENTER_NO,' ') as CENTER_NO, "
					+ "            nvl(wlx01.M2_NAME,' ')   as M2_NAME, " 
					+ "            wlx01.CANCEL_no, wlx01.CANCEL_date, "
					+ "            nvl(cd01.hsien_id,' ') as  hsien_id, "
					+ "            a01.acc_code, "
					+ " 		   round(amt/?,0) as amt,"
					//+ " 		   decode(amt,'',0,round(amt/?,0)) as amt,"
					+ "            a01.acc_range, "
					+ "            nvl(cd01.hsien_name,'OTHER')  as  hsien_name, "
					+ "            cd01.FR001W_output_order     as  FR001W_output_order "
					+ "     from  (select * from "+cd01_table+" cd01 where  cd01.hsien_id <> 'Y') cd01 " 
					+ "     left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id "  
					+ "     left join (select BANK_NO, BANK_NAME,  bank_type  "
					+ "                from bn01 where m_year=? and bn01.bank_type in('6','7'))bn01 on wlx01.bank_no=bn01.bank_no "
					+ "     left join (  "       
					+ "                select bank_code,m_year,m_month, "
					+ "                      a01.acc_code,round(decode(acc_code,'220900',amt/2,amt/1),0) as amt, "
					+ "                       decode(acc_code,'190000','01','220000','02','220900','04','140000','05','990000','06') as acc_range " 
					+ "                from  A01 "
					+ "                where M_Year = ?  and M_month = ? " 
					+ "                and  a01.acc_code in('190000','220000','220900','140000','990000') "
					+ "                union "
					+ "                select bank_code,m_year,m_month,'120000' as acc_code, "
					+ "                       round(sum(decode(acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /1,0)-  " 
					+ "                       round(sum(decode(acc_code,'120700',amt))/1,0) as amt, "
					+ "                       '03' as acc_range " 
					+ "                from  A01 "
					+ "                where M_Year = ? and M_month = ? "
					+ "                and  a01.acc_code in('120000','120800','150300','120700') "
					+ "                group by bank_code,m_year,m_month "
					+ "                union "
					+ "                select bank_code,m_year,m_month,'120800' as acc_code, "
					+ "                       round(sum(decode(acc_code,'120800',amt,'150300',amt,0)) /1,0) as amt, "
					+ "                       '13' as acc_range  "
					+ "                from  A01 "
					+ "                where M_Year = ?  and M_month = ? " 
					+ "                and  a01.acc_code in('120800','150300') "
					+ "                group by bank_code,m_year,m_month "
					+ "                union "
					+ "                select bank_code,m_year,m_month,'840760' as acc_code,sum(amt) as amt, "
					+ "                       '07' as acc_range  "
					+ "                from  A04 "
					+ "                where M_Year = ? and M_month = ?  "
					+ "                and  a04.acc_code in('840760') " 
					+ "                group by bank_code,m_year,m_month "
					+ "                union "
					+ "                select bank_code,m_year,m_month,a02.acc_code,round(amt/1,0) as amt, "
					+ "                       decode(acc_code,'990510','08','990610','12') as acc_range " 
					+ "                from  A02 "
					+ "                where M_Year = ? and M_month = ? " 
					+ "                and  a02.acc_code in('990510','990610') "
					+ "                union "
					+ "                select bank_code,m_year,m_month,a99.acc_code,round(amt/1,0) as amt, "
					+ "                       decode(acc_code,'992150','09') as acc_range " 
					+ "                from  A99 "
					+ "                where M_Year =? and M_month = ? "
					+ "                and  a99.acc_code in('992150') "
					+ "                union "
					+ "                select bank_code,m_year,m_month,a99.acc_code,round(amt/1,0) as amt, "
					+ "                       decode(acc_code,'992150','10') as acc_range  "
					+ "                from  A99 "
					+ "                where M_Year =? and M_month = ? "  
					+ "                and  a99.acc_code in('992150') "
					+ "                union "
					+ "                select bank_code,m_year,m_month,a99.acc_code,round(amt/1,0) as amt, "
					+ "                       decode(acc_code,'992150','11') as acc_range " 
					+ "                from  A99 "
					+ "                where M_Year =?  and M_month =? " 
					+ "                and  a99.acc_code in('992150') "
					+ "                union "
					+ "                select bank_code,m_year,m_month,'990021' as acc_code, "
					+ "                       round(sum(loan1_baddebt+loan2_baddebt+loan3_baddebt+loan4_baddebt)/'1',0) as loan_baddebt_sum, "
					+ "                       '14' as acc_range " 
					+ "                from A10 "
					+ "                where M_Year = ?  and M_month = ? "
					+ "                group by bank_code,m_year,m_month "
					+ "                union "
					+ "                select bank_code,m_year,m_month,'990002' as acc_code, "
					+ "                       round(sum(loan2_amt+loan3_amt+loan4_amt)/'1',0) as loan_sum, "
					+ "                       '15' as acc_range " 
					+ "                from A10 "
					+ "                where M_Year =?  and M_month = ? " 
					+ "                group by bank_code,m_year,m_month "
					+ "                union "
					+ "                select bank_code,m_year,m_month,'990001' as acc_code, "
					+ "                       round(sum(loan1_amt+loan2_amt+loan3_amt+loan4_amt)/'1',0) as loan_sum, "
					+ "                       '16' as acc_range " 
					+ "                from A10 "
					+ "                where M_Year =?  and M_month = ? "
					+ "                group by bank_code,m_year,m_month "
					+ "                union "
					+ "                select bank_code,m_year,m_month,'990006' as acc_code, "
					+ "                       round(sum(invest1_amt+invest2_amt+invest3_amt+invest4_amt)/'1',0) as loan_sum, "
					+ "                       '17' as acc_range " 
					+ "                from A10 "
					+ "                where M_Year =?  and M_month =? "
					+ "                group by bank_code,m_year,m_month "
					+ "                union "
					+ "                select bank_code,m_year,m_month,'990005' as acc_code, "
					+ "                       round(sum(loan1_amt+loan2_amt+loan3_amt+loan4_amt)/'1',0) as loan_sum, "
					+ "                       '18' as acc_range  "
					+ "                from A10 "
					+ "                where M_Year =?  and M_month =?  "
					+ "                group by bank_code,m_year,m_month "
					+ "              )a01 on bn01.bank_no = A01.Bank_Code "
					+ " ) Temp_Output ";
            rule_L_paramList.add(unit);
			rule_L_paramList.add(wlx01_m_year);
			rule_L_paramList.add(wlx01_m_year);
			for(int i=0;i<=12;i++){
				rule_L_paramList.add(S_YEAR);	   
				rule_L_paramList.add(String.valueOf(Integer.parseInt(S_MONTH)));
			}
			
            rule_R = " select * from " 
            		+ " (  "
            		+ "     select NVL((BN01.BANK_NO || BN01.BANK_NAME), ' ')  as S_Report_Name, " 
            		+ "            BN01.BANK_NO, BN01.BANK_NAME, bn01.bank_type , "
            		+ "            nvl(wlx01.CENTER_NO,' ') as CENTER_NO, "
            		+ "            nvl(wlx01.M2_NAME,' ')   as M2_NAME, " 
            		+ "            wlx01.CANCEL_no, wlx01.CANCEL_date, "
            		+ "            nvl(cd01.hsien_id,' ') as  hsien_id, "
            		+ "            a01.acc_code, "
            		+ " 		   round(amt/?,0) as amt,"
            		//+ " 		   decode(amt,'',0,round(amt/?,0)) as amt,"
            		+ "            a01.acc_range, "
            		+ "            nvl(cd01.hsien_name,'OTHER')  as  hsien_name, "
            		+ "            cd01.FR001W_output_order    as  FR001W_output_order "
            		+ "     from  (select * from "+cd01_table+" cd01 where  cd01.hsien_id <> 'Y') cd01 " 
            		+ "     left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id "  
            		+ "     left join (select BANK_NO, BANK_NAME,  bank_type  "
            		+ "                from bn01 where m_year=? and bn01.bank_type in('6','7'))  bn01 " 
            		+ "          on wlx01.bank_no=bn01.bank_no "
            		+ "     left join (  "     
            		+ "                 select bank_code,m_year,m_month, "
            		+ "                        acc_code,round(amt/1,0) as amt, "
            		+ "                       decode(acc_code,'400000','01') as acc_range " 
            		+ "                from  A01 "
            		+ "                where M_Year = ? and M_month =? "
            		+ "                and  a01.acc_code in('400000') "
            		+ "                union "
            		+ "                select bank_code,m_year,m_month,'992130' as acc_code, "
            		+ "                       round(sum(amt) /1,0) as amt, "
            		+ "                       '02' as acc_range " 
            		+ "                from  (select * from A99 " 
            		+ "                       where M_Year = ? and M_month =? "
            		+ "                       and  a99.acc_code = '992130' "
            		+ "                       union " 
            		+ "                       select m_year,m_month,bank_code,acc_code,amt from A02 "
            		+ "                       where M_Year = ? and M_month =? "
            		+ "                       and  a02.acc_code in('990420','990620') "
            		+ "                )A99 "
            		+ "                group by bank_code,m_year,m_month "
            		+ "                union "
            		+ "                select bank_code,m_year,m_month,'992140' as acc_code, "
            		+ "                       round(sum(amt) /1,0) as amt, "
            		+ "                       '03' as acc_range "
            		+ "                from  (select * from A99 " 
            		+ "                       where M_Year =?  and M_month =? "
            		+ "                       and  a99.acc_code = '992140' "
            		+ "                       union  "
            		+ "                       select m_year,m_month,bank_code,acc_code,amt from A02 "
            		+ "                       where M_Year =?  and M_month =? "
            		+ "                       and  a02.acc_code in('990410','990610') "
            		+ "                )A99 "
            		+ "                group by bank_code,m_year,m_month "                            
            		+ "                union "
            		+ "                select bank_code,m_year,m_month,a02.acc_code, "
            		+ "                       round(decode(acc_code,'990630',amt/2,amt/1),0) as amt, "
            		+ "                       decode(acc_code,'990630','04','990810','05','990510','09','990612','12') as acc_range " 
            		+ "                from  A02 "
            		+ "                where M_Year =?  and M_month =? "
            		+ "                and  a02.acc_code in('990630','990810','990510','990612') "
            		+ "                union "
            		+ "                select bank_code,m_year,m_month,'992510' as acc_code, "
            		+ "                       round(sum(decode(acc_code,'992510',amt,'992520',amt,'992530',amt,'992540',amt,0)) /1,0) as amt, "
            		+ "                       '06' as acc_range  "
            		+ "                from  A99 "
            		+ "                where M_Year =? and M_month =? "
            		+ "                and  a99.acc_code in('992510','992520','992530','992540') "
            		+ "                group by bank_code,m_year,m_month "
            		+ "                union "
            		+ "                select bank_code,m_year,m_month,'992610' as acc_code, "
            		+ "                       round(sum(decode(acc_code,'992610',amt,'992620',amt,'992630',amt,'992640',amt,0)) /1,0) as amt, "
            		+ "                       '07' as acc_range " 
            		+ "                from  A99 "
            		+ "                where M_Year =? and M_month =? "
            		+ "                and  a99.acc_code in('992610','992620','992630','992640') "
            		+ "                group by bank_code,m_year,m_month "
            		+ "                union "
            		+ "                select bank_code,m_year,m_month,'990511' as acc_code, "
            		+ "                       round(sum(decode(acc_code,'990511',amt,'990512',amt,0)) /1,0) as amt, "
            		+ "                       '08' as acc_range " 
            		+ "                from  A02 "
            		+ "                where M_Year =? and M_month =? "
            		+ "                and  a02.acc_code in('990511','990512') "
            		+ "                group by bank_code,m_year,m_month "
            		+ "                union "
            		+ "                select bank_code,m_year,m_month,a99.acc_code,round(amt/1,0) as amt, "
            		+ "                       decode(acc_code,'992550','10','992650','11') as acc_range " 
            		+ "                from  A99 "
            		+ "                where M_Year =? and M_month =? "
            		+ "                and  a99.acc_code in('992550','992650') "
            		+ "                union "
            		+ "                select bank_code,m_year,m_month,'990025' as acc_code, "
            		+ "                       round(sum(loan1_baddebt+loan2_baddebt+loan3_baddebt+loan4_baddebt)/'1',0) as loan_baddebt_sum, "
            		+ "                       '13' as acc_range " 
            		+ "                from A10 "
            		+ "                where M_Year =? and M_month =? "
            		+ "                group by bank_code,m_year,m_month "
            		+ "                union "
            		+ "                select bank_code,m_year,m_month,'990025' as acc_code, "
            		+ "                       round(sum(loan1_baddebt+loan2_baddebt+loan3_baddebt+loan4_baddebt)/'1',0) as loan_baddebt_sum, "
            		+ "                       '14' as acc_range " 
            		+ "                from A10 "
            		+ "                where M_Year =? and M_month =? "
            		+ "                group by bank_code,m_year,m_month "
            		+ "                union "
            		+ "                select bank_code,m_year,m_month,'840760' as acc_code, "
            		+ "                       round(sum(amt) /1,0) as amt, "
            		+ "                       '15' as acc_range " 
            		+ "                from  (select * from A04 " 
            		+ "                       where M_Year =?  and M_month =? "
            		+ "                       and  a04.acc_code = '840760' "
            		+ "                       union " 
            		+ "                       select m_year,m_month,bank_code,acc_code,amt from A01 "
            		+ "                       where M_Year =?  and M_month =? "
            		+ "                       and  a01.acc_code in('990000') "
            		+ "                )A04 "
            		+ "                group by bank_code,m_year,m_month "      
            		+ "                union "
            		+ "                select bank_code,m_year,m_month,'990005' as acc_code, "
            		+ "                       round(sum(loan1_amt+loan2_amt+loan3_amt+loan4_amt)/'1',0) as loan_sum, "
            		+ "                       '16' as acc_range  "
            		+ "                from A10 "
            		+ "                where M_Year =?  and M_month =? "
            		+ "                group by bank_code,m_year,m_month "
            		+ "                union "
            		+ "                select bank_code,m_year,m_month,'990010' as acc_code, "
            		+ "                       round(sum(invest1_amt+invest2_amt+invest3_amt+invest4_amt)/'1',0) as loan_sum, "
            		+ "                       '17' as acc_range  "
            		+ "                from A10 "
            		+ "                where M_Year =? and M_month = ? "
            		+ "                group by bank_code,m_year,m_month "
            		+ "                union "
            		+ "                select bank_code,m_year,m_month,'120000' as acc_code, "
            		+ "                         round(sum(decode(acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /1,0)-  " 
            		+ "                         round(sum(decode(acc_code,'990611',amt))/1,0) as amt, "
            		+ "                        '18' as acc_range  "
            		+ "               from ( "
            		+ "               select * from  A01 "
            		+ "               where M_Year =?  and M_month =?  "
            		+ "                and  a01.acc_code in('120000','120800','150300') "
            		+ "               union  "
            		+ "               select m_year,m_month,bank_code,acc_code,amt from A02 "
            		+ "               where M_Year =? and M_month =? "
            		+ "                and  a02.acc_code in('990611') "
            		+ "               )A01  "
            		+ "               group by bank_code,m_year,m_month "
            		+ "               union "
            		+ "               select bank_code,m_year,m_month,'840740' as acc_code,sum(amt) as amt, "
            		+ "                        '19' as acc_range "//--第06項檢核,增加檢核19 
            		+ "               from  A04 "
            		+ "               where M_Year =? and M_month =? "
            		+ "               and  a04.acc_code in('840740') " 
            		+ "               group by bank_code,m_year,m_month "
            		+ "              )a01 on bn01.bank_no = A01.Bank_Code "
            		+ " ) Temp_Output  ";
            rule_R_paramList.add(unit);
			rule_R_paramList.add(wlx01_m_year);
			rule_R_paramList.add(wlx01_m_year);
			for(int i=0;i<=18;i++){
				rule_R_paramList.add(S_YEAR);	   
				rule_R_paramList.add(String.valueOf(Integer.parseInt(S_MONTH)));
			}
			
			String rule_2 = " where 	BANK_NO <> ' '		and "
						   + " ( "
						   + " (? = 'Z') or "
					       + " ((? = '6' or ? = '7') and  ? = BANK_NO) or "
					       + " (? = '8' and ? = CENTER_NO) or"
					       + " (? = 'B' and ? = M2_NAME) "
						   + " )    and "
						   + " ( "
						   + " (?= 'ALL') or "
		 			       + " ((?= '6') and (Bank_Type = '6')) or "
						   + " ((?= '7') and (Bank_Type = '7')) "
						   + " )"                                                                               
						   + " order by  Bank_Type,BANK_NO,acc_range,FR001W_output_order ";
			paramList.add(tmpbank_type);
			paramList.add(tmpbank_type);
			paramList.add(tmpbank_type);
			paramList.add(tmpbank_no);
			paramList.add(tmpbank_type);
			paramList.add(tmpbank_no);
			paramList.add(tmpbank_type);
			paramList.add(tmpbank_no);
			paramList.add(bank_type_list);
			paramList.add(bank_type_list);
			paramList.add(bank_type_list);		
			for(int i=0;i<paramList.size();i++){
	            rule_L_paramList.add(paramList.get(i));
	            rule_R_paramList.add(paramList.get(i));
			}
			List dbData = DBManager.QueryDB_SQLParam( rule_L + rule_2,rule_L_paramList,"cancel_date,amt");   
			String tmpBank_no="";
			String tmpBank_no_name="";
			Properties lData = new Properties();		
			List DataList_L = new LinkedList();
			if(dbData != null && dbData.size() > 0){
			   for(int i=0;i<dbData.size();i++){			 			          
			       if(i==0){
			           tmpBank_no = (String)((DataObject)dbData.get(i)).getValue("bank_no");     			       			       			       
			           tmpBank_no_name = (String)((DataObject)dbData.get(i)).getValue("s_report_name");
			       }    
			       if(((String)((DataObject)dbData.get(i)).getValue("bank_no")).equals(tmpBank_no)){//與前一個bank_no相同時,儲存acc_code,amt			          			       			       
			       	  if(((DataObject)dbData.get(i)).getValue("acc_code") != null){//95.06.19 acc_code == null時,只會有一筆資料	            			            			             			             
			             lData.setProperty((String)((DataObject)dbData.get(i)).getValue("acc_code"),((DataObject)dbData.get(i)).getValue("amt")==null?"":(((DataObject)dbData.get(i)).getValue("amt")).toString());			       			       			            
			             //System.out.println("leftData.set("+(String)((DataObject)dbData.get(i)).getValue("acc_code")+")="+(((DataObject)dbData.get(i)).getValue("amt")).toString());
			          }   
			       }else{ 
			          System.out.println("left add Bank_no="+tmpBank_no);
			          DataList_L.add(tmpBank_no);
			          DataList_L.add(tmpBank_no_name);
			          DataList_L.add(lData);
			          DataListALL_L.add(DataList_L);//欲回傳的left data
			          DataList_L = new LinkedList();//單一機構的data
			          lData = new Properties();//A01.acc_code|amt				          
			          tmpBank_no = (String)((DataObject)dbData.get(i)).getValue("bank_no");
			          tmpBank_no_name = (String)((DataObject)dbData.get(i)).getValue("s_report_name");
			          System.out.println("now tmpbank_no="+tmpBank_no);
			          if(((DataObject)dbData.get(i)).getValue("acc_code") != null){//95.06.19 acc_code == null時,只會有一筆資料	            			            			             			             
			             lData.setProperty((String)((DataObject)dbData.get(i)).getValue("acc_code"),((DataObject)dbData.get(i)).getValue("amt")==null?"":(((DataObject)dbData.get(i)).getValue("amt")).toString());   
			             //System.out.println("leftData.set("+(String)((DataObject)dbData.get(i)).getValue("acc_code")+")="+(((DataObject)dbData.get(i)).getValue("amt")).toString());
			          }
			       }			       
			   }			   
			   //將最後一筆寫回DataListALL_L
			   System.out.println("left add last Bank_no="+tmpBank_no);				  
			   DataList_L.add(tmpBank_no);
			   DataList_L.add(tmpBank_no_name);
			   DataList_L.add(lData);
			   DataListALL_L.add(DataList_L);			   			   
			}
			dbData = DBManager.QueryDB_SQLParam( rule_R + rule_2, rule_R_paramList, "cancel_date,amt");   
			tmpBank_no="";	
			tmpBank_no_name="";		
			Properties rData = new Properties();		
			List DataList_R = new LinkedList();
			if(dbData != null && dbData.size() > 0){
			   for(int i=0;i<dbData.size();i++){			 			         
			       if(i==0){
			          tmpBank_no = (String)((DataObject)dbData.get(i)).getValue("bank_no");     			       			       			       
			          tmpBank_no_name = (String)((DataObject)dbData.get(i)).getValue("s_report_name");
			       }   
			       if(((String)((DataObject)dbData.get(i)).getValue("bank_no")).equals(tmpBank_no)){//與前一個bank_no相同時,儲存acc_code,amt			          			       			       
			          if(((DataObject)dbData.get(i)).getValue("acc_code") != null){//95.06.19 acc_code == null時,只會有一筆資料
			             rData.setProperty((String)((DataObject)dbData.get(i)).getValue("acc_code"),((DataObject)dbData.get(i)).getValue("amt")==null?"":(((DataObject)dbData.get(i)).getValue("amt")).toString());			       			       			            
			             //System.out.println("rigthData.set("+(String)((DataObject)dbData.get(i)).getValue("acc_code")+")="+(((DataObject)dbData.get(i)).getValue("amt")).toString());
			          }   
			          
			       }else{ 
			          System.out.println("rigth add bank_no="+tmpBank_no);
			          DataList_R.add(tmpBank_no);
			          DataList_R.add(tmpBank_no_name);
			          DataList_R.add(rData);
			          DataListALL_R.add(DataList_R);//欲回傳的rigth data
			          DataList_R = new LinkedList();//單一機構的data
			          rData = new Properties();//rigth .acc_code|amt				          
			          tmpBank_no = (String)((DataObject)dbData.get(i)).getValue("bank_no");
			          tmpBank_no_name = (String)((DataObject)dbData.get(i)).getValue("s_report_name");
			          System.out.println("now tmpbank_no="+tmpBank_no);
			          if(((DataObject)dbData.get(i)).getValue("acc_code") != null){//95.06.19 acc_code == null時,只會有一筆資料	            			            			
			            rData.setProperty((String)((DataObject)dbData.get(i)).getValue("acc_code"),((DataObject)dbData.get(i)).getValue("amt")==null?"":(((DataObject)dbData.get(i)).getValue("amt")).toString());   
			            //System.out.println("rightData.set("+(String)((DataObject)dbData.get(i)).getValue("acc_code")+")="+(((DataObject)dbData.get(i)).getValue("amt")).toString());
			          }  			          
			       }			       
			   }			   			   
			   
			   //將最後一筆寫回DataListALL_R			      
			   System.out.println("rigth add last Bank_no="+tmpBank_no);	
			   DataList_R.add(tmpBank_no);
			   DataList_R.add(tmpBank_no_name);
			   DataList_R.add(rData);
			   DataListALL_R.add(DataList_R);			   			   
			}
    		
    		HashMap h = new HashMap();
    		h.put("dataL",DataListALL_L);
    		h.put("dataR",DataListALL_R);
            return h;
    }	    
%>    