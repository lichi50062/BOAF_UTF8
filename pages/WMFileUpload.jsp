F<%
// 93.12.07 add update WML01_UPLOAD by 2295
// 93.12.17 add 權限檢核 by 2295
// 93.12.18 fix 若有已點選的tbank_no,則以已點選的tbank_no為主 by 2295
// 93.12.23 add 超過登入時間,請重新登入 by 2295
// 94.02.15 fix 檔案上傳成功時,回上一頁為act=new+Report_no+S_YEAR+S_MONTH by 2295
// 94.05.17 add 不為共用中心上傳檔案時,檔案內容只能是該機構代號 by 2295 							   							    
//          add 未加入共用中心的機構代號不能上傳 by 2295
// 		   add 檔案年月須與資料年月相符 by 2295
// 94.06.17 add 把bank_type寫到session by 2295
// 94.06.20 fix 當filename已存在時.更新錯誤的問題 by 2295
// 95.04.07 fix 提到上頭來check資料格式for數字 by 2295
// 96.07.10 add A08.不檢核科目代號. by 2295
// 97.01.02 add A09.不檢核科目代號. by 2295
// 97.06.13 add A10.不檢核科目代號. by 2295
// 99.09.24 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
//101.07.20 add 農信保報表.M106/M201/M206上傳檔案/更新檢核狀態/上傳至農金局DB主機 by 2295
//102.04.25 add A02.增加上傳.990421/990621農金局文號.長度最多80 by 2295
//106.02.13 add A02.增加上傳.990422/990622農金局文號.長度最多80 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.Utility_WM" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.report.*" %>
<%@ page import="com.tradevan.util.transfer.*" %>
<%@ page import="com.oreilly.servlet.MultipartRequest" %>
<%@ page import="java.util.*" %>
<%@ page import="java.io.*" %>
<%@ page import="java.text.*" %>
<%@include file="./include/Header.include" %>
<%
	String CopyResult = "";	
	String fileName="";
	
	
	String Report_no = Utility.getTrimString(dataMap.get("Report_no"));
	String S_YEAR = Utility.getTrimString(dataMap.get("S_YEAR"));
	String S_MONTH = Utility.getTrimString(dataMap.get("S_MONTH"));	
	String WMdataDir = Utility.getProperties("WMdataDir")+System.getProperty("file.separator")+Report_no;
	String WMTempBKDir = Utility.getProperties("WMTempBKDir")+System.getProperty("file.separator")+Report_no;
	String user_id = (session.getAttribute("muser_id") == null)?"":(String)session.getAttribute("muser_id");
	String user_name = (session.getAttribute("muser_name") == null)?"":(String)session.getAttribute("muser_name"); 
	
	//======================================================================================================================
	//fix 93.12.18 若有已點選的tbank_no,則以已點選的tbank_no為主============================================================
	String bank_code = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");				
	String nowtbank_no = Utility.getTrimString(dataMap.get("tbank_no")); 
	if(!nowtbank_no.equals("")){
	   session.setAttribute("nowtbank_no",nowtbank_no);//將已點選的tbank_no寫入session	   
	}   
	bank_code = ( session.getAttribute("nowtbank_no")==null ) ? bank_code : (String)session.getAttribute("nowtbank_no");			
	//=======================================================================================================================
	String bank_type = Utility.getTrimString(dataMap.get("bank_type")); 	
	System.out.println("request.bank_type="+bank_type);
	if(bank_type.equals("")){	   
	   bank_type=(String)session.getAttribute("nowbank_type");
	   System.out.println("session.nowbank_type="+bank_type);
	}
	if(!bank_type.equals("") && bank_type != null){
	    session.setAttribute("nowbank_type",bank_type);//94.06.17
	}    
	
	int UploadSize = Integer.parseInt(Utility.getProperties("UploadSize"));
	if(S_MONTH.indexOf("0") == 0 ){		//modify by 2354 12.14
	   S_MONTH = S_MONTH.substring(S_MONTH.indexOf("0")+1,S_MONTH.length());
	}
	System.out.println("act="+act);	
	System.out.println("Report_no="+Report_no);			
	
    if(!Utility.CheckPermission(request,report_no)){//無權限時,導向到LoginError.jsp
        rd = application.getRequestDispatcher( LoginErrorPgName );        
    }else{        
        actMsg = actMsg + mkdir(Report_no);//新增目錄
    	//set next jsp 	
    	if(act.equals("new")){//起始page    	    
        	rd = application.getRequestDispatcher( EditPgName );
    	}else if(act.equals("Upload")){//上傳檔案
    	    System.out.println("Upload Dir="+WMdataDir);    	    
    	    fileName = Utility.parseFileName(request.getParameter("FileName"));
    	    //檢查檔案在WML01_LOCK中,是否已被LOCK
    	    if(!CheckFileLock(S_YEAR,S_MONTH,bank_code,Report_no)){//該檔沒有被LOCK住
    	    	//檢核檔案是否存在Input的目錄下
    	    	if(Utility.CheckFileExist(WMdataDir,request.getParameter("FileName"))){    	       	
    	        	//若欲上傳的檔案已存在dataDir的目錄下,先將檔案存到TempBKDir的目錄,並更改檔名加上.yyyyMMddHHmmssSSS    	        
    	    	    MultipartRequest multi = new MultipartRequest(request, WMTempBKDir, UploadSize  * 1024);    			        	    	 
    				if(Report_no.equals("M03") || Report_no.equals("M05")){	//modify by 2354 12.15
    					actMsg = checkReportDataM03M05(Report_no,WMdataDir+System.getProperty("file.separator")+fileName);
    				}else{    				    
    	    	    	actMsg = actMsg + checkReportData(Report_no,WMTempBKDir+System.getProperty("file.separator")+fileName,bank_type,bank_code,S_YEAR,S_MONTH);
    	    	    }
    	    	    if(actMsg.equals("")){//檢查檔案內容
    	        		String newFile = fileName+"."+ Utility.getDateFormat("yyyyMMddHHmmssSSS");	        
    	        		CopyResult = Utility.CopyFile(WMTempBKDir+System.getProperty("file.separator")+fileName,WMTempBKDir+System.getProperty("file.separator")+ newFile);
    	        		if(CopyResult.equals("0")){
    	           	   		File tmpFile = new File(WMTempBKDir+System.getProperty("file.separator")+fileName);
    	           	   		if(tmpFile.exists()) tmpFile.delete();
    	               		alertMsg = fileName+"該資料檔已上傳,是否欲覆蓋原檔案??";    	           	       	
    	       	   	   		webURL_Y = "/pages/WMFileUpload.jsp?act=OverWrite&FileName="+newFile+"&Report_no="+Report_no+"&Report_no="+Report_no+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&test=nothing";
    	       	   	   		webURL_N = "/pages/WMFileUpload.jsp?act=Delete&FileName="+newFile+"&Report_no="+Report_no+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&test=nothing";
    	        		}else{
    	               		actMsg = CopyResult;
    	            	}   	        
    	            }else{
    	                File tmpFile = new File(WMTempBKDir+System.getProperty("file.separator")+fileName);
    	           	   	if(tmpFile.exists()) tmpFile.delete();    	    	           	   	
    	           	   	actMsg = "該資料檔"+fileName+"無法上傳成功<br>錯誤原因:<br>"+actMsg;               	    	           	   	    	       	   	   	
    	            }	
    	    	}else{//欲上傳的檔案不在dataDir的目錄下,直接將檔案儲存至dataDir的目錄    	    	    
    				MultipartRequest multi = new MultipartRequest(request, WMdataDir, UploadSize * 1024);    			        	
    				//M03及 M05之處理流程與其他不同  Add by winnin 12.14
    				if(Report_no.equals("M03") || Report_no.equals("M05")){
    					actMsg = checkReportDataM03M05(Report_no,WMdataDir+System.getProperty("file.separator")+fileName);
    				}else{
    					actMsg = checkReportData(Report_no,WMdataDir+System.getProperty("file.separator")+fileName,bank_type,bank_code,S_YEAR,S_MONTH);
    				}    				
    	    	    if(actMsg.equals("")){//檢查檔案內容成功    	    	        
    	    		   actMsg = multi.getFilesystemName("UpFileName")+"檔案上傳成功<br>";
    	    		   actMsg = actMsg + updateDB(fileName,user_id,user_name,S_YEAR,S_MONTH,bank_code);//更新WML01_UPLOAD    	    	        
    	    		}else{
    	    		    File tmpFile = new File(WMdataDir+System.getProperty("file.separator")+fileName);
    	           	   	if(tmpFile.exists()) tmpFile.delete();    	
    	    		   	actMsg = "該資料檔"+fileName+"無法上傳成功<br>錯誤原因:<br>"+actMsg;               	    	           	   	    	       	   	   	
    	    		}	
    	    	} 	
    	    }else{
    	        //alertMsg = Utility.parseFileName(request.getParameter("FileName"))+"該資料檔已被鎖住,無法再上傳該檔案";    	           	       	
    	        alertMsg = "該資料檔已被鎖住,無法再上傳該檔案";    	           	       	    	        
    	    }
    	    
    	    //rd = application.getRequestDispatcher( EditPgName +"?Report_no="+Report_no+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&test=nothing");
        	rd = application.getRequestDispatcher( nextPgName +"?Report_no="+Report_no+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&test=nothing");
    	}else if (act.equals("OverWrite")){
    		String tmpbkFile = request.getParameter("FileName");
			String dataFile = "";
		
			if(tmpbkFile.lastIndexOf(".") != -1){
			   dataFile = tmpbkFile.substring(0,tmpbkFile.lastIndexOf("."));
			}
			System.out.println("OverWrite file="+tmpbkFile);
			System.out.println("dataFile="+dataFile);
			File tmpFile = new File(WMdataDir+System.getProperty("file.separator")+dataFile);
			if(tmpFile.exists()) tmpFile.delete();
			CopyResult = Utility.CopyFile(WMTempBKDir+System.getProperty("file.separator")+tmpbkFile,Utility.getProperties("WMdataDir")+System.getProperty("file.separator")+Report_no+System.getProperty("file.separator")+dataFile);
    		System.out.println("copyfile="+CopyResult);
    		
    		if(CopyResult.equals("0")){
    		   tmpFile = new File(WMTempBKDir+System.getProperty("file.separator")+tmpbkFile);
    		   if(tmpFile.exists()){
    		      System.out.println("tmpFile.exists=true");
    		      tmpFile.delete();    	
    		   }   
    		   actMsg = dataFile+"檔案上傳成功<br>";	   
    		   actMsg = actMsg + updateDB(dataFile,user_id,user_name,S_YEAR,S_MONTH,bank_code);//更新WML01_UPLOAD    	    	        
    		}else{
    		   actMsg = CopyResult;
    		}
    		rd = application.getRequestDispatcher( nextPgName +"?Report_no="+Report_no+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&test=nothing");
    	}else if (act.equals("Delete")){    	    
			System.out.println("Delete file="+request.getParameter("FileName"));			
			File tmpFile = new File(WMTempBKDir+System.getProperty("file.separator")+request.getParameter("FileName"));
			if(tmpFile.exists()) tmpFile.delete();
			rd = application.getRequestDispatcher( EditPgName +"?Report_no="+Report_no+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&test=nothing" );
			System.out.println(EditPgName +"?Report_no="+Report_no+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&test=nothing");
			System.out.println("delete ok");
    	}
    
    	request.setAttribute("actMsg",actMsg);
    	request.setAttribute("alertMsg",alertMsg);
    	request.setAttribute("webURL_Y",webURL_Y);
    	request.setAttribute("webURL_N",webURL_N);    	
    }

%>
<%@include file="./include/Tail.include" %>

<%!
    private final static String report_no = "WMFileUpload";
    private final static String nextPgName = "/pages/ActMsg.jsp";
    private final static String EditPgName = "/pages/"+report_no+"_Edit.jsp";
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    private File logfile;
	private FileOutputStream logos=null;
	private BufferedOutputStream logbos = null;
	private PrintStream logps = null;
	private Date nowlog = new Date();
	private SimpleDateFormat logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");
	private SimpleDateFormat logfileformat = new SimpleDateFormat("yyyyMMddHHmmss");
	private Calendar logcalendar;
	
    private boolean CheckFileLock(String S_YEAR,String S_MONTH,String bank_code,String Report_no){//檢核此檔在WML01_LOCK/WML_LOCK中有無被Lock
    		boolean lock = false;
    		StringBuffer sqlCmd = new StringBuffer();
			List paramList = new ArrayList();
			
			sqlCmd.append("select lock_status from WML01_LOCK where m_year=? and m_month=? and bank_code=? and report_no=?");
			paramList.add(S_YEAR);
			paramList.add(S_MONTH);
			paramList.add(bank_code);
			paramList.add(Report_no);			
    		List WML01_LOCK = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");  
    		
    		sqlCmd = new StringBuffer();
			paramList = new ArrayList(); 
			sqlCmd.append("select lock_status from WML01 where m_year=? and m_month=? and bank_code=? and report_no=?");
			paramList.add(S_YEAR);
			paramList.add(S_MONTH);
			paramList.add(bank_code);
			paramList.add(Report_no);   
    		List WML01 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");            
    		
    		if(WML01_LOCK.size() > 0){    		
    		  if((((String)((DataObject)WML01_LOCK.get(0)).getValue("lock_status")) != null) &&(((String)((DataObject)WML01_LOCK.get(0)).getValue("lock_status")).equals("Y"))){
    		    System.out.println("WML01_lock true");
    		    lock = true;
    		  }  
    		}  
    		if(WML01.size() > 0){    		
    		  if((((String)((DataObject)WML01.get(0)).getValue("lock_status")) != null) && ((String)((DataObject)WML01.get(0)).getValue("lock_status")).equals("Y")){
    		    System.out.println("WML01 true");
    		    lock = true;
    		  }  
    		}
    		
           return lock;
    }
    private String mkdir(String Report_no){//檢核目錄不存在時,則新增該目錄    		
    		String mkdirOK="";
    		try{
    			File dataDir = new File(Utility.getProperties("WMdataDir")+System.getProperty("file.separator")+Report_no);        
	        	File dataBKDir = new File(Utility.getProperties("WMdataBKDir")+System.getProperty("file.separator")+Report_no);        
    	    	File TempBKDir = new File(Utility.getProperties("WMTempBKDir")+System.getProperty("file.separator")+Report_no);        
        		if(!dataDir.exists()){
         			if(!Utility.mkdirs(Utility.getProperties("WMdataDir")+System.getProperty("file.separator")+Report_no)){
         		   		mkdirOK=mkdirOK+Utility.getProperties("WMdataDir")+System.getProperty("file.separator")+Report_no+"目錄新增失敗";
         			}    
        		}
        		if(!dataBKDir.exists()){
         			if(!Utility.mkdirs(Utility.getProperties("WMdataBKDir")+System.getProperty("file.separator")+Report_no)){
         		   		mkdirOK=mkdirOK+Utility.getProperties("WMdataBKDir")+System.getProperty("file.separator")+Report_no+"目錄新增失敗";
         			}    
        		}
        		if(!TempBKDir.exists()){
         			if(!Utility.mkdirs(Utility.getProperties("WMTempBKDir")+System.getProperty("file.separator")+Report_no)){
         		   		mkdirOK=mkdirOK+Utility.getProperties("WMTempBKDir")+System.getProperty("file.separator")+Report_no+"目錄新增失敗";
         			}    
        		}
        	}catch(Exception e){
        		System.out.println("目錄新增失敗"+e.getMessage()) ;
        		mkdirOK="目錄新增失敗";
        	}	 
        	return mkdirOK;
    }
    
    private String checkReportData(String Report_no,String filename,String bank_type,String bank_code,String S_YEAR,String S_MONTH){    		    		
    		String errMsg = "";
    		byte[] byte1 ;	//add by winnin 2004.12.09
    		String tmpBankCode="";
    		StringBuffer sqlCmd = new StringBuffer();
			List paramList = new ArrayList();
			String wlx01_m_year = (Integer.parseInt(S_YEAR) < 100)?"99":"100";
			try{
				String	txtline			= null;
				int[][]	checkNumber		= ListArray.getRowIsNumber(Report_no);//用來check欄位數字
				int[][] checkBankNo 	= ListArray.getBankNo(Report_no);//用來check總機構代號
				int[][] checkBankCode 	= ListArray.getBankCode(Report_no);	//用來check總機構代號 add by winnin 12.15
				int rowlength = Integer.parseInt(ListArray.getRowLength(Report_no));//用來check每列長度
				List dbData ;	//modify by 2354 12.15
				String errMsgSub="";	//add by 2354 2004.12.15
				Map values = new HashMap();//總機構代號List
				Map bank_codevalues = new HashMap();//加入共用中心的機構代號List
				
				System.out.println("bank_type="+bank_type);
				System.out.println("bank_code="+bank_code);
				System.out.println("S_YEAR="+S_YEAR);
				System.out.println("S_MONTH="+S_MONTH);
				if(bank_type.equals("8")){//94.05.17取得加入該共用中心的機構代號
				   sqlCmd.append(" select bn01.bank_no,bn01.bank_name");
				   sqlCmd.append(" from (select * from bn01 where m_year=?)bn01  LEFT JOIN (select * from wlx01 where m_year=?)WLX01 on bn01.bank_no = WLX01.bank_no ");
				   sqlCmd.append(" where bn01.bn_type <> '2' ");
				   sqlCmd.append(" and WLX01.center_no = ?");
				   sqlCmd.append(" order by bank_no");
				   paramList.add(wlx01_m_year);
				   paramList.add(wlx01_m_year);
				   paramList.add(bank_code);
				   dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
				   if(dbData != null && dbData.size() != 0){
				      for(int i=0;i<dbData.size();i++){
					    bank_codevalues.put((String)((DataObject)dbData.get(i)).getValue("bank_no"),"true");
					  }
				   }
				   System.out.println("bank_codeList.size()="+dbData.size());
				}   
				//根據Report_no來決定總機構代號List modify by 2354 2004.12.15
				if(Report_no.substring(0,1).equals("A")){
					dbData = DBManager.QueryDB_SQLParam("select bank_no from bn01 where bn_type <> '2'",null,"");
	 				for(int i=0;i<dbData.size();i++){
					    values.put((String)((DataObject)dbData.get(i)).getValue("bank_no"),"true");
					}
					errMsgSub="總機構代號";
				}else if (Report_no.equals("M01")){		//add by 2354
					dbData = DBManager.QueryDB_SQLParam("select distinct guarantee_item_no from m00_guarantee_item",null,"");
	 				for(int i=0;i<dbData.size();i++){
					    values.put((String)((DataObject)dbData.get(i)).getValue("guarantee_item_no"),"true");
					}
					errMsgSub="保證項目代碼";
				}else if (Report_no.equals("M02")){		//add by 2354
					dbData = DBManager.QueryDB_SQLParam("select distinct loan_unit_no from m00_loan_unit",null,"");
	 				for(int i=0;i<dbData.size();i++){
					    values.put((String)((DataObject)dbData.get(i)).getValue("loan_unit_no"),"true");
					}
					errMsgSub="貸款機構代碼";
				}else if (Report_no.equals("M04")){		//add by 2354
					dbData = DBManager.QueryDB_SQLParam("select distinct loan_use_no from m00_loan_use",null,"");
	 				for(int i=0;i<dbData.size();i++){
					    values.put((String)((DataObject)dbData.get(i)).getValue("loan_use_no"),"true");
					}
					errMsgSub="貸款用途代碼";
				}else if (Report_no.equals("M06") || Report_no.equals("M07")){		//add by 2354
					dbData = DBManager.QueryDB_SQLParam("select distinct area_no from m00_area",null,"");
	 				for(int i=0;i<dbData.size();i++){
					    values.put((String)((DataObject)dbData.get(i)).getValue("area_no"),"true");
					}
					errMsgSub="地區別代碼";
				}
				/*
				Set key = values.keySet();			
				Iterator it = key.iterator();			
				while (it.hasNext()) {
				    System.out.println("bank_no:"+it.next());					
				}*/			
				for(int i=0;i<checkNumber.length;i++){
					System.out.println(checkNumber[i][0]+":"+checkNumber[i][1]);
				}
				
		    	File tmpFile = new File(filename);
		    	if(tmpFile.length() == 0){
		    	   errMsg =errMsg + "檔案內容無資料";
		    	   return errMsg;
		    	}
		    	FileReader	f = new FileReader(tmpFile);
				LineNumberReader in	= new LineNumberReader(f);				
				
				doLoop:
				while ((txtline = in.readLine()) != null) {
						if (txtline.trim().equals("")){
							continue doLoop;
						}else{
						    byte1 = txtline.getBytes("BIG5");  //add by winnin 2004.12.09						    
						    tmpBankCode = txtline.substring(checkBankCode[0][0], checkBankCode[0][1]).trim();
						    System.out.println(tmpBankCode+":byte1.length="+byte1.length);
						    //102.04.25 add A02.增加上傳.990421/990621農金局文號.長度最多80
						    if ((!Report_no.equals("A05") && !Report_no.equals("A02") && byte1.length != rowlength) || 
						         (Report_no.equals("A05") && byte1.length!=32 && byte1.length != rowlength) || 
						         (Report_no.equals("A02") && 
						         ((tmpBankCode.equals("990421") && byte1.length > 98) || 
						          (tmpBankCode.equals("990422") && byte1.length > 98) || 
						          (tmpBankCode.equals("990621") && byte1.length > 98) ||       
						          (tmpBankCode.equals("990622") && byte1.length > 98) ||       
						          (!tmpBankCode.equals("990421") && !tmpBankCode.equals("990422") && !tmpBankCode.equals("990621") && !tmpBankCode.equals("990622") && byte1.length != rowlength)))
						       ) 
						    {		//modify by winnin 2004.12.23
								//errMsg = errMsg + "第" + in.getLineNumber() + "行資料長度不符<br>";//106.02.13		
								errMsg = errMsg +"("+tmpBankCode+ ")第" + in.getLineNumber() + "行資料長度不符<br>";		
							}else if((String)values.get(txtline.substring(checkBankNo[0][0], checkBankNo[0][1])) == null){							    
							    errMsg = errMsg + "第" + in.getLineNumber() + "行:"+errMsgSub+txtline.substring(checkBankNo[0][0], checkBankNo[0][1])+"不存在<br>";  //modify by 2354 2004.12.15							    
							}else if(Report_no.substring(0,1).equals("A")){ 							    
							    //94.05.17未加入共用中心的機構代號不能上傳=================================================
							    if(bank_type.equals("8") && ((String)bank_codevalues.get(txtline.substring(checkBankNo[0][0], checkBankNo[0][1])) == null)){
							       System.out.println("不屬於該共用中心的機構代號:"+txtline.substring(checkBankNo[0][0], checkBankNo[0][1]));							    							    							       
							       errMsg = errMsg + "第" + in.getLineNumber() + "行:"+errMsgSub+txtline.substring(checkBankNo[0][0], checkBankNo[0][1])+"未加入貴共用中心<br>";
							    }
							    //94.05.17不為共用中心上傳檔案時,檔案內容只能是該機構代號====================================  							   							    
							    if((!bank_type.equals("8")) &&  (!txtline.substring(checkBankNo[0][0], checkBankNo[0][1]).equals(bank_code))){							    
							       System.out.println("不屬於檔案名稱的機構代號:"+txtline.substring(checkBankNo[0][0], checkBankNo[0][1]));							    							    							       
							       errMsg = errMsg + "第" + in.getLineNumber() + "行:總機構代號"+txtline.substring(checkBankNo[0][0], checkBankNo[0][1])+"與上傳檔名的總機構代號不符合<br>"; 							    
							    }
							    //94.05.17檔案年月須與資料年月相符=========================================================
							    if((Integer.parseInt(txtline.substring(0,3)) != Integer.parseInt(S_YEAR))
							    || (Integer.parseInt(txtline.substring(3,5)) != Integer.parseInt(S_MONTH)))
							    {
							      System.out.println("不符合檔案年月的資料:"+txtline.substring(0,3)+"年"+txtline.substring(3,5)+"月");
							      errMsg = errMsg + "第" + in.getLineNumber() + "行:檔案年月與資料內容的年月不符合<br>";
							    }
							    
							    //modify by winnin 2004.12.09 begin ======================= 								
							    //fix 95.04.07 提到上頭來check資料格式 by 2295							    
								for(int i=0;i<checkNumber.length;i++){		
								    if(Report_no.equals("A08")){//96.07.10 add 
								       tmpBankCode = "A08";
								    }else if(Report_no.equals("A09")){//97.01.02 add 
								       tmpBankCode = "A09";
								    }else if(Report_no.equals("A10")){//97.06.13 add 
								       tmpBankCode = "A10";								       
								    }else{
									    tmpBankCode = txtline.substring(checkBankCode[0][0], checkBankCode[0][1]).trim();	// add by winnin 2004.12.10
									}
									// add by winnin 2004.12.09 N: 為名稱性欄位
									// 102.04.25 add A02.990421/990621.為名稱性欄位									
									//System.out.println("tmpBankCode.substring(tmpBankCode.length()-1)="+tmpBankCodetmd.substring(tmpBankCode.length()-1));
									if(!tmpBankCode.substring(tmpBankCode.length()-1).equals("N") && !tmpBankCode.equals("990421") && !tmpBankCode.equals("990422") && !tmpBankCode.equals("990621") && !tmpBankCode.equals("990622")){	//項目代號末一碼為N時或A02.990421/990621/990422/990622, 項目金額為文字性欄位
										System.out.println("tmpBankCode="+tmpBankCode);
										//System.out.println("data="+txtline.substring(checkNumber[i][0], checkNumber[i][1]));
										if (!Utility_WM.isNumber(txtline.substring(checkNumber[i][0], checkNumber[i][1]))) {
											System.out.println("tmpBankCodeerror");
								 			errMsg = errMsg + "第" + in.getLineNumber() + "行(" + checkNumber[i][0] + ":" 
								  				   + checkNumber[i][1] + ")資料格式不符<br>";							
										}										
								  	}//end of /項目代號末一碼為N時, 項目金額為文字性欄位
								}//end of for --檢核科目代號	 																 
								//modify by winnin 2004.12.09 end ======================= 
							}else{
								//modify by winnin 2004.12.09 begin ======================= 
								System.out.println("check 資料格式數字欄位");								
								for(int i=0;i<checkNumber.length;i++){	
								    if(!Report_no.equals("A08")){//96.07.10 add 
									    tmpBankCode=txtline.substring(checkBankCode[0][0], checkBankCode[0][1]).trim();	// add by winnin 2004.12.10
									}else{
										tmpBankCode="A08";
									}
									// add by winnin 2004.12.09 N: 為名稱性欄位
									//System.out.println("tmpBankCode="+tmpBankCode);
								 	//System.out.println("tmpBankCode.substring(tmpBankCode.length()-1)="+tmpBankCodetmd.substring(tmpBankCode.length()-1));
									if(!tmpBankCode.substring(tmpBankCode.length()-1).equals("N")){	//項目代號末一碼為N時, 項目金額為文字性欄位
									   //System.out.println("data="+txtline.substring(checkNumber[i][0], checkNumber[i][1]));
									   if (!Utility_WM.isNumber(txtline.substring(checkNumber[i][0], checkNumber[i][1]))) {
								 		   errMsg = errMsg + "第" + in.getLineNumber() + "行(" + checkNumber[i][0] + ":" 
								  			      + checkNumber[i][1] + ")資料格式不符<br>";							
									   }										
								  	}
								}//end of for --檢核科目代號		 																 
								//modify by winnin 2004.12.09 end ======================= 
							}
						}
				}		
				in.close();
				f.close();				
			}catch(Exception e){
				//errMsg = errMsg + e.getMessage();
				System.out.println(e.getMessage());
			}    
    		return errMsg;
    }
    
	/* write by winnin 12.14 */
    private String checkReportDataM03M05(String Report_no,String filename){    		    		
    		String errMsg = "";
    		String[] reportItem;
    		byte[] byte1 ;
    		
			try{
				String	txtline		  = null;
				int[][]	checkNumber   = null;
				int[][] checkBankNo   = null;
				int[][] checkBankCode = null;
				int[] rowlength;
				List dbData ;
				String errMsgSub="";
				Map valuesM03_NOTE 	= new HashMap();
				Map valuesM03 		= new HashMap();
				Map valuesM05_NOTE 	= new HashMap();
				Map valuesM05_TOTACC= new HashMap();
				Map valuesM05       = new HashMap();
				System.out.println("In M03M05  Report_no="+Report_no);
				
				if(Report_no.equals("M03")){
					//將值塞入 valuesM03
					dbData = DBManager.QueryDB_SQLParam("select data_range from m00_data_range_item where report_no='M03' and data_range_type='C'",null,"");
					for(int i=0;i<dbData.size();i++){
					    valuesM03.put((String)((DataObject)dbData.get(i)).getValue("data_range"),"true");
					}
					
					//將值塞入 valuesM03_NOTE
					dbData = DBManager.QueryDB_SQLParam("select data_range from m00_data_range_item where report_no='M03' and data_range_type='S'",null,"");
					for(int i=0;i<dbData.size();i++){
					    valuesM03_NOTE.put((String)((DataObject)dbData.get(i)).getValue("data_range"),"true");
					}
				}else{
					//將值塞入 valuesM05
					dbData = DBManager.QueryDB_SQLParam("select loan_unit_no || data_range as \"data_range\" from m00_data_range_item,m00_loan_unit where report_no='M05' and data_range_type='C'",null,"");
					for(int i=0;i<dbData.size();i++){
					    valuesM05.put((String)((DataObject)dbData.get(i)).getValue("data_range"),"true");
					}

					//將值塞入 valuesM05_NOTE
					dbData = DBManager.QueryDB_SQLParam("select data_range from m00_data_range_item where report_no='M05' and data_range_type='N'",null,"");
					for(int i=0;i<dbData.size();i++){
					    valuesM05_NOTE.put((String)((DataObject)dbData.get(i)).getValue("data_range"),"true");
					}
				
					//將值塞入 valuesM05_TOTACC
					dbData = DBManager.QueryDB_SQLParam("select loan_unit_no || '0ET' as \"data_range\" from m00_loan_unit",null,"");
					for(int i=0;i<dbData.size();i++){
					    valuesM05_TOTACC.put((String)((DataObject)dbData.get(i)).getValue("data_range"),"true");
					}
				}

				System.out.println("Report_no="+Report_no);
				System.out.println("Report_no.substring(0,1)="+Report_no.substring(0,1));
				/*
				Set key = values.keySet();
				Iterator it = key.iterator();			
				while (it.hasNext()) {
				    System.out.println("bank_no:"+it.next());					
				}			
				for(int i=0;i<checkNumber.length;i++){
					System.out.println(checkNumber[i][0]+":"+checkNumber[i][1]);
				}
				*/
		    	File tmpFile = new File(filename);
		    	if(tmpFile.length() == 0){
		    	   errMsg =errMsg + "檔案內容無資料";
		    	   return errMsg;
		    	}
		    	FileReader	f = new FileReader(tmpFile);
				LineNumberReader in	= new LineNumberReader(f);				
				
				doLoop:
				while ((txtline = in.readLine()) != null) {
					System.out.println("in M03M05 while");
					if (txtline.trim().equals("")){
						continue doLoop;
					}else{
						byte1 = txtline.getBytes("BIG5");
						System.out.println("Report_no="+Report_no+":rowlength="+byte1.length);
						reportItem = ListArray.getReportItem(Report_no,byte1.length);
						for(int j=0;j<reportItem.length;j++){
							System.out.println("reportItem["+j+"]="+reportItem[j]);
						}
						if(reportItem == null || reportItem[0].equals("")){
							errMsg = errMsg + "第" + in.getLineNumber() + "行資料長度不符<br>";
						}else{
							if((reportItem[3].equals("M03") 		&& ((String)valuesM03.get(txtline.substring(Integer.parseInt(reportItem[1]),Integer.parseInt(reportItem[2]))) == null)) ||
							   (reportItem[3].equals("M03_NOTE") 	&& ((String)valuesM03_NOTE.get(txtline.substring(Integer.parseInt(reportItem[1]),Integer.parseInt(reportItem[2]))) == null)) ||
							   (reportItem[3].equals("M05") 		&& ((String)valuesM05.get(txtline.substring(Integer.parseInt(reportItem[1]),Integer.parseInt(reportItem[2]))) == null)) ||
							   (reportItem[3].equals("M05_NOTE") 	&& ((String)valuesM05_NOTE.get(txtline.substring(Integer.parseInt(reportItem[1]),Integer.parseInt(reportItem[2]))) == null)) ||
							   (reportItem[3].equals("M05_TOTACC") 	&& ((String)valuesM05_TOTACC.get(txtline.substring(Integer.parseInt(reportItem[1]),Integer.parseInt(reportItem[2]))) == null))
							  ){
							  	System.out.println("lkj");
								errMsg = errMsg + "第" + in.getLineNumber() + "行長度:"+byte1.length+", 代號:"+txtline.substring(Integer.parseInt(reportItem[1]),Integer.parseInt(reportItem[2]))+"不存在<br>"; 
							}else{
								checkNumber	  = ListArray.getRowIsNumber(reportItem[0]);
								for(int i=0;i<checkNumber.length;i++){	
									if (!Utility_WM.isNumber(txtline.substring(checkNumber[i][0], checkNumber[i][1]))) {
								 		errMsg = errMsg + "第" + in.getLineNumber() + "行(" + checkNumber[i][0] + ":" 
								  			   + checkNumber[i][1] + ")資料格式不符<br>";							
								  	}
								}	 								 
							}
						}
					}
				}		
				in.close();
				f.close();				
			}catch(Exception e){
				//errMsg = errMsg + e.getMessage();
				System.out.println(e.toString());
				System.out.println(e.getMessage());
			}    
    		return errMsg;
    }

    private String updateDB(String filename,String user_id,String user_name,String m_year,String m_month,String bank_code){       
		String errMsg="";		
		List paramList = new ArrayList();
		StringBuffer sql = new StringBuffer();		
		List updateDBList = new ArrayList();//0:sql 1:data		
		List updateDBSqlList = new ArrayList();
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data
		String updateAcgfMsg="";
		try{
    		sql.append(" select * from WML01_UPLOAD where filename=?");
    		paramList.add(filename);    	    	    	
        	List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"");   
        	sql.delete(0, sql.length());	
        	if(dbData.size() != 0){//有資料,update userid
           		sql.append("UPDATE WML01_UPLOAD SET muser_id=?,muser_name=? where filename=?");
           		dataList = new ArrayList<String>();//傳內的參數List		 	           				   
				dataList.add(user_id); 
				dataList.add(user_name); 
				dataList.add(filename);   
        	}else{//無資料,直接insert
           		sql.append("INSERT INTO WML01_UPLOAD VALUES(?,?,?)");
           		dataList = new ArrayList<String>();//傳內的參數List		 	           				   
				dataList.add(filename);   
				dataList.add(user_id); 
				dataList.add(user_name); 
        	}    
        	
        	updateDBSqlList.add(sql.toString());//0:欲執行的sql	
			updateDBDataList.add(dataList);//1:傳內的參數List
			updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
			updateDBList.add(updateDBSqlList);        	
        	
        	if(DBManager.updateDB_ps(updateDBList)){						 
				errMsg = errMsg + "相關資料寫入資料庫成功";								
			}else{
				errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
			}
			//101.07.20 add更新農信保資料/上傳至農金局DB主機
			if(filename.substring(0,4).equals("M106") || filename.substring(0,4).equals("M201") || filename.substring(0,4).equals("M206")){
				updateAcgfMsg = updateACGF(filename,user_id,user_name,m_year,m_month,filename.substring(0,4),bank_code);
				if(errMsg.equals("相關資料寫入資料庫成功")){				
					errMsg = updateAcgfMsg;
				}else{
					errMsg += updateAcgfMsg;
				}		
			}
			    
		}catch(Exception e){
		   System.out.println(e+":"+e.getMessage());
			errMsg = errMsg + "相關資料寫入資料庫失敗";					 
		}	
		return errMsg;
    }
    //更新農信檢核狀態 101.07.20 add 
    private String updateACGF(String filename,String user_id,String user_name,String m_year,String m_month,String report_no,String bank_code){       
		String errMsg="";		
		List paramList = new ArrayList();
		StringBuffer sql = new StringBuffer();		
		List updateDBList = new ArrayList();//0:sql 1:data		
		List updateDBSqlList = new ArrayList();
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data
		List bank_codeList = new LinkedList();
		String input_method = "F";//檔案上傳
		String add_date="";//申報日期
		String upd_code="U";//檢核結果
		String upd_method = "M";//更新方式  M:人工
		String common_center="";//由共用中心傳入
		String checkResult="true";
		boolean updateOK=false;
		SimpleDateFormat emailformat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");	
		String WMdataDir = "";
		String WMdataBKDir = "";		
		List filename_List = new LinkedList();
		SimpleDateFormat bkformat = new SimpleDateFormat("yyyyMMddHHmmss");
		
		try{
			
    		WMdataDir = Utility.getProperties("WMdataDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");
		    WMdataBKDir = Utility.getProperties("WMdataBKDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");		
	
    		bank_codeList.add(bank_code);			  	
		    File tmpFile = new File(WMdataDir+filename);		
		    Date filedate = new Date(tmpFile.lastModified());
		    add_date=bkformat.format(filedate);
		    Date today = new Date();
    	    int	batch_no = today.hashCode();
    	       
		    List dbData =  Utility.getWML01(m_year,m_month,(String)bank_codeList.get(0),report_no);
		    //寫入WML01/WML01_LOG     
		    List getUpdateDBList = Utility.Insert_UpdateWML01(dbData,m_year, m_month,bank_code,report_no,input_method,add_date,user_id,user_name,common_center,upd_method,upd_code,batch_no);
		    for(int j=0;j<getUpdateDBList.size();j++){
		    	updateDBList.add(getUpdateDBList.get(j));
		    }
		    Utility.sendMailNotification(bank_code,report_no, m_year,m_month, checkResult,input_method,emailformat.format(filedate),filename,user_id,"");
						
    	    sql.append(" select * from WML01_M_UPLOAD where m_year=? and m_month=? and bank_code=? and report_no=?");
    		paramList.add(m_year);    	
    		paramList.add(m_month);    	
    		paramList.add(bank_code);    	
    		paramList.add(report_no);
            dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"");   
        	sql.delete(0, sql.length());	
        	if(dbData.size() != 0){//有資料,update userid
           	   sql.append("UPDATE WML01_M_UPLOAD SET filename=? where m_year=? and m_month=? and bank_code=? and report_no=?");
           	   dataList = new ArrayList<String>();//傳內的參數List		 	           				   				     
			   dataList.add(filename);  
			   dataList.add(m_year); 
			   dataList.add(m_month);  
			   dataList.add(bank_code);
			   dataList.add(report_no);
        	}else{//無資料,直接insert
           	   sql.append("INSERT INTO WML01_M_UPLOAD VALUES(?,?,?,?,?)");
           	   dataList = new ArrayList<String>();//傳內的參數List		 	           				   				     
			   dataList.add(m_year); 
			   dataList.add(m_month);  
			   dataList.add(bank_code);
			   dataList.add(report_no);
			   dataList.add(filename);  
        	}    
    	    updateDBSqlList.add(sql.toString());//0:欲執行的sql	
			updateDBDataList.add(dataList);//1:傳內的參數List
			updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
			updateDBList.add(updateDBSqlList);        	
					 			
			//上傳檔案至農金局DB主機  
			String putMsg = uploadAcgfRpt(report_no,filename);
			if(putMsg.equals("")){//上傳成功
				
			   errMsg = parseAcgfRpt.doParserRpt(report_no,filename,m_year,m_month);
			   System.out.println("test1.errMsg="+errMsg);
			   if(errMsg.indexOf("匯入資料庫成功") != -1){
			      String CopyResult = Utility.CopyFile(WMdataDir+System.getProperty("file.separator")+filename,WMdataBKDir+System.getProperty("file.separator")+ filename);
       		      if(CopyResult.equals("0")){//copy成功時,才將檔案刪除,避免使用rename造成的錯誤
          	      	tmpFile = new File(WMdataDir+System.getProperty("file.separator")+filename);
          	      	if(tmpFile.exists()) tmpFile.delete();              		
       		      }
       		      //將WML01_UPLOAD..filename對應的使用者帳號.姓名刪除						     
			      updateDBList.add(Utility.deleteWML01_UPLOAD(filename));
			      if(!DBManager.updateDB_ps(updateDBList)){
			      	errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
			      }	
			   }
       	    }else{
       			errMsg += putMsg;
       		}	
       		
		}catch(Exception e){
		   System.out.println(e+":"+e.getMessage());
			errMsg = errMsg + "相關資料寫入資料庫失敗";					 
		}	
		return errMsg;
    }
    
    //農信保財報.上傳至DB主機 101.07.20 add
    private String uploadAcgfRpt(String report_no,String fileName){
        String errMsg="";
        String CopyResult = "";
        String sqlCmd="";
        String putMsg="";
        List filename_List = new LinkedList();
        File tmpFile = null;
        List updateDBSqlList = new LinkedList();
       	try{
    		String rptIP=Utility.getProperties("rptIP");
	        String rptID=Utility.getProperties("rptID");
	        String rptPwd=Utility.getProperties("rptPwd");
         	filename_List.add(fileName);

       	    File logDir  = new File(Utility.getProperties("logDir"));
   			if(!logDir.exists()){
        	    if(!Utility.mkdirs(Utility.getProperties("logDir"))){
    	   	 	  System.out.println("目錄新增失敗");
        	    }
   			}
            logfile = new File(logDir + System.getProperty("file.separator") + "uploadAcgfRpt.log");
   			System.out.println("logfile filename="+logDir + System.getProperty("file.separator") +"uploadAcgfRpt.log");
   			logos = new FileOutputStream(logfile,true);
   			logbos = new BufferedOutputStream(logos);
   			logps = new PrintStream(logbos);
    		System.out.println("=============執行農信保財務報表上傳至DB Server開始===========");
    		logfile = new File(logDir + System.getProperty("file.separator") + "uploadAcgfRpt.log");
  			System.out.println("logfile filename="+logDir + System.getProperty("file.separator") +"uploadAcgfRpt.log");
   			logos = new FileOutputStream(logfile,true);
   			logbos = new BufferedOutputStream(logos);
   			logps = new PrintStream(logbos);

           	//上傳至農金局DB主機(農信保財務資料)===============================================================================
			//putFiles_multiSubDir(String server_host, String username, String password,
			//   				   String remote_path(主機端), String local_path(local端), String workDir(子目錄名稱), List filename(上傳檔案List),
			//   				   PrintStream logps)
   			//上傳至農金局DB主機(金庫財務資料)===============================================================================
			putMsg = RptGenerateALL.putFiles_multiSubDir(rptIP, rptID, rptPwd,Utility.getProperties("acgfDir"), Utility.getProperties("WMdataDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator"),report_no,filename_List,logps);

            if(putMsg == null){//上傳檔案成功
               System.out.println("檔案上傳成功");
    	    }else{
    	       errMsg += putMsg;
    	    }
    	    System.out.println("errMsg = "+errMsg);
   			logps.flush();   			
    	    System.out.println("=============執行農信保報表上傳至DB Server結束===========");
		}catch(Exception e){
		   System.out.println(e+":"+e.getMessage());
			errMsg = errMsg + "[uploadAcgfRpt Error]";
		}
		return errMsg;
    }
%>    