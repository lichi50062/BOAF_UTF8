<%
// 94.10.21 create by 2495
// 99.12.12 fix sqlInjection by 2808
//102.12.03 add pdf上傳 by 2295	
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.dao.DAOFactory" %>
<%@ page import="com.tradevan.util.dao.RdbCommonDao" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.Date" %>
<%@ page import="java.util.*" %>
<%@ page import="com.oreilly.servlet.MultipartRequest"%>
<%@ page import="java.util.List" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.math.BigDecimal" %>
<%@ page import="java.io.*" %>
<%@ page import="com.tradevan.util.UpdateTimeOut" %>


<%
	RequestDispatcher rd = null;	
	String actMsg = "";	
	String alertMsg = "";	
	String webURL = "";	
	boolean doProcess = false;
		
	//取得session資料,取得成功時,才繼續往下執行===================================================
	if(session.getAttribute("muser_id") == null)
	{//session timeout	
			System.out.println("ZZ091W.jsp login timeout");   
	   	rd = application.getRequestDispatcher( "/pages/reLogin.jsp?url=LoginError.jsp?timeout=true" );         	   
	   	try{
          rd.forward(request,response);
      }catch(Exception e){
          System.out.println("forward Error:"+e+e.getMessage());
      }
  }
  else
  {
      doProcess = true;

  }    
	if(doProcess)
	{//若muser_id資料時,表示登入成功====================================================================
	  String muser_id = (String)session.getAttribute("muser_id");	
		String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");		
		String seq_no = ( request.getParameter("seq_no")==null ) ? "" : (String)request.getParameter("seq_no");	
		String headmark = ( request.getParameter("head_title")==null ) ? "" : (String)request.getParameter("head_title");
		System.out.println("headmark:"+headmark);
		String notifyUrl = ( request.getParameter("notifyUrl")==null ) ? "" : (String)request.getParameter("notifyUrl");
		System.out.println("notifyUrl:"+notifyUrl);	
		String StartYear = ( request.getParameter("StartYear")==null ) ? "0" : (String)request.getParameter("StartYear");
		System.out.println("StartYear:"+StartYear);
		String StartMonth = ( request.getParameter("StartMonth")==null ) ? "" : (String)request.getParameter("StartMonth");	
		System.out.println("StartMonth:"+StartMonth);
		String StartDay = ( request.getParameter("StartDay")==null ) ? "" : (String)request.getParameter("StartDay");	
		System.out.println("StartDay:"+StartDay);
		String StartHour = ( request.getParameter("StartHour")==null ) ? "" : (String)request.getParameter("StartHour");	
		System.out.println("StartHour:"+StartHour);
		String EndYear = ( request.getParameter("EndYear")==null ) ? "" : (String)request.getParameter("EndYear");	
		System.out.println("EndYear:"+EndYear);
		String EndMonth = ( request.getParameter("EndMonth")==null ) ? "" : (String)request.getParameter("EndMonth");		
		System.out.println("EndMonth:"+EndMonth);
		String EndDay = ( request.getParameter("EndDay")==null ) ? "" : (String)request.getParameter("EndDay");	
	  System.out.println("EndDay:"+EndDay);
		String notify_date ="";
		String notify_end_date = "";
											
    if(!CheckPermission(request))
    {//無權限時,導向到LoginError.jsp
        rd = application.getRequestDispatcher( LoginErrorPgName );        
    }
    else
    {            
    	//set next jsp 	    	
    	if(act.equals("new"))
    	{    	      	   
    	   rd = application.getRequestDispatcher( EditPgName +"?act=new");                	
      }
      else if(act.equals("List"))
      {
           rd = application.getRequestDispatcher( ListPgName +"?act="+act);                	              	        	        	    	    
      }      
      else if(act.equals("Edit"))
      {
           List WLX_Notify = Get_WLX_Notify(seq_no);    	      
           request.setAttribute("WLX_Notify",WLX_Notify);      
           //102.12.03設定異動者資訊======================================================================
		   request.setAttribute("maintainInfo","select * from WLX_Notify where  seq_no =" + seq_no);								       
		   //=======================================================================================================================       	         	     	   	        	    
           rd = application.getRequestDispatcher( EditPgName +"?act=Edit");      
    	}
    	else if(act.equals("Clear")){
    		  List WLX_Notify = Get_WLX_Notify(seq_no);
        	List WTT01_USERNAME = TakeUserName(muser_id);		       			
		      String tseq_no=" ",theadmark=" ",tnotify_date=" ",tnotify_url=" ",tnotify_end_date=" ",tappend_file=" ",tuser_id=" ",tuser_name=" ",tupdate_date=" ",tappend_file2=" ",tappend_file3=" ",tuser_id_c=" ",tuesr_name_c=" ",tupdate_type_c="U";
          tseq_no = seq_no;       	   		    	      												       	   		           	   		    
		      theadmark = (String)((DataObject)WLX_Notify.get(0)).getValue("headmark");		      
 		      tnotify_date = (((DataObject)WLX_Notify.get(0)).getValue("notify_date")==null ) ? "" : (String)((DataObject)WLX_Notify.get(0)).getValue("notify_date");	
		      tappend_file = (((DataObject)WLX_Notify.get(0)).getValue("append_file")==null ) ? "" : (String)((DataObject)WLX_Notify.get(0)).getValue("append_file");			      
		      tnotify_end_date = (((DataObject)WLX_Notify.get(0)).getValue("notify_end_date")==null ) ? "" : (String)((DataObject)WLX_Notify.get(0)).getValue("notify_end_date");			       	
		      tuser_id = (String)((DataObject)WLX_Notify.get(0)).getValue("user_id");		        	   		  
		      tuser_name = (((DataObject)WLX_Notify.get(0)).getValue("user_name")==null ) ? "" : (String)((DataObject)WLX_Notify.get(0)).getValue("user_name");	
		      tupdate_date = (((DataObject)WLX_Notify.get(0)).getValue("update_date")==null ) ? "" : (String)((DataObject)WLX_Notify.get(0)).getValue("update_date");	        	   		  		        	   		  
		      tuser_id_c = muser_id;		      	   		    	
		      tuesr_name_c = (String)((DataObject)WTT01_USERNAME.get(0)).getValue("muser_name"); 		      	   		    		        	   		  		     	   		     	  
		      tnotify_url = (((DataObject)WLX_Notify.get(0)).getValue("notify_url")==null ) ? "" : (String)((DataObject)WLX_Notify.get(0)).getValue("notify_url");		
		  		String msg = InsertWLX_Notify_LOG(tseq_no,theadmark,tnotify_date,tappend_file,tnotify_end_date,tuser_id,tuser_name,tupdate_date,tuser_id_c,tuesr_name_c,tupdate_type_c,tnotify_url);        	   		          	   		          	   		   	  	
        	//String saveDirectory = "C:\\Sun\\WebServer6.1\\BOAF\\exp\\wlx\\WLX_Notify\\";
		  		String saveDirectory = Utility.getProperties("notifyDir");
		  		saveDirectory = saveDirectory+tappend_file;
		  		File objFile_del = new File(saveDirectory);
		  		objFile_del.delete();
		  
        	UpdateWLX_Notify_AppendFile(seq_no);
        	WLX_Notify = Get_WLX_Notify(seq_no);
        	request.setAttribute("WLX_Notify",WLX_Notify);
        	rd = application.getRequestDispatcher( EditPgName +"?act=Edit");      
    	}
    	
     	else if(act.equals("del")){
     		   List WLX_SEQNO = TakeMaxSeqno(); 
        	 int seqcount = Integer.parseInt( (((DataObject)WLX_SEQNO.get(0)).getValue("maxseqno")==null )? "-1" : ((DataObject)WLX_SEQNO.get(0)).getValue("maxseqno").toString());        		  					
           if(seqcount!= (-1))
           {	           
           		if(request.getParameter("isDelete")==null)
           		{
           			System.out.println(request.getParameter("isDelete"));  
           		}
           		else
           		{	
           			String Star[] = request.getParameterValues("isDelete");
           			System.out.println("刪除公告筆數: "+Star.length);
					 			for(int i=0;i<Star.length;i++)
					 			{	
		        	   	String tseq_no=" ",theadmark=" ",tnotify_url=" ",tnotify_date=" ",tnotify_end_date=" ",tappend_file=" ",tuser_id=" ",tuser_name=" ",tupdate_date=" ",tappend_file2=" ",tappend_file3=" ",tuser_id_c=" ",tuesr_name_c=" ",tupdate_type_c="D";					      	   		
        	   		  List WLX_Notify = Get_WLX_Notify(Star[i]);
        	   		  List WTT01_USERNAME = TakeUserName(muser_id);
		       	   		tseq_no = Star[i];    theadmark = (String)((DataObject)WLX_Notify.get(0)).getValue("headmark");
		        	   	tnotify_date = (((DataObject)WLX_Notify.get(0)).getValue("notify_date")==null ) ? "" : (String)((DataObject)WLX_Notify.get(0)).getValue("notify_date");	
		        	   	tappend_file = (((DataObject)WLX_Notify.get(0)).getValue("append_file")==null ) ? "" : (String)((DataObject)WLX_Notify.get(0)).getValue("append_file");	
		        	    tnotify_end_date = (((DataObject)WLX_Notify.get(0)).getValue("notify_end_date")==null ) ? "" : (String)((DataObject)WLX_Notify.get(0)).getValue("notify_end_date");	
		        	   	tuser_id = (String)((DataObject)WLX_Notify.get(0)).getValue("user_id");		        	   		  
		        	   	tuser_name = (((DataObject)WLX_Notify.get(0)).getValue("user_name")==null ) ? "" : (String)((DataObject)WLX_Notify.get(0)).getValue("user_name");	
		        	   	tupdate_date = (((DataObject)WLX_Notify.get(0)).getValue("update_date")==null ) ? "" : (String)((DataObject)WLX_Notify.get(0)).getValue("update_date");	
		        	   	tuser_id_c = muser_id;
		      	   		tuesr_name_c = (String)((DataObject)WTT01_USERNAME.get(0)).getValue("muser_name"); 		      	   		    		        	   		  
		     	   		  tnotify_url = (((DataObject)WLX_Notify.get(0)).getValue("notify_url")==null ) ? "" : (String)((DataObject)WLX_Notify.get(0)).getValue("notify_url");		
									String msg = InsertWLX_Notify_LOG(tseq_no,theadmark,tnotify_date,tappend_file,tnotify_end_date,tuser_id,tuser_name,tupdate_date,tuser_id_c,tuesr_name_c,tupdate_type_c,tnotify_url);        	   		          	   		          	   		 
		     	   		       
		     	  			String saveDirectory = Utility.getProperties("notifyDir");
									saveDirectory = saveDirectory+tappend_file;
									File objFile_del = new File(saveDirectory);
		  			      objFile_del.delete();
									Del_WLX_Notify(Star[i]); 		      
        	      }
        	 }
        	 }
        	 rd = application.getRequestDispatcher( ListPgName +"?act=List");     
    	}
    	else if(act.equals("insert")){  
		  		String saveDirectory = Utility.getProperties("notifyDir");
    			int maxPostSize = 100*1024*1024;   		  
    		  int count = 0;
    		  File objFile = new File(saveDirectory);
    		  if(!objFile.exists())
    		  		objFile.mkdir();   		     
    			//宣告上傳檔案名稱
    			String FileName = null;
    			//支援中文檔名
    			String enCoding = "MS950";
    			//產生一個新的MultipartRequest的物件,multi
    			MultipartRequest  multi = new MultipartRequest(request,saveDirectory,maxPostSize,enCoding);    		   		    		   
    		  //取得所有上傳之檔案輸入型態名稱
    		  Enumeration filesname = multi.getFileNames();
    		  String append_file1="",appfile_link="",reFileName="";
    		  	while(filesname.hasMoreElements())
    		   	{	      		                		   		    
    		    	String nextFilename = (String)filesname.nextElement();
    		     	FileName = multi.getFilesystemName(nextFilename); 
    		     	System.out.println("上傳之檔案FileName ="+FileName);
    		     	  		     	   		     			  
    		     			if( !(FileName == null))
    		     			{
											Calendar calendar=new GregorianCalendar();
											int year=calendar.get(Calendar.YEAR);
											String str_year=Integer.toString(year);
											int month=calendar.get(Calendar.MONTH);
											String str_month=Integer.toString(month);
											int day=calendar.get(Calendar.DATE);
											String str_day=Integer.toString(day);
											int hour=calendar.get(Calendar.HOUR_OF_DAY);
											String str_hour=Integer.toString(hour);
											int minute=calendar.get(Calendar.MINUTE);
											String str_minute=Integer.toString(minute);
											int second=calendar.get(Calendar.SECOND);
											String str_second=Integer.toString(second);
											int index = FileName.indexOf('.');      
      								String file_type = FileName.substring(index+1,index+4);
						
		       						if(file_type.equals("doc"))
		      							reFileName=str_year+str_month+str_day+str_hour+str_minute+str_second+".doc";
		       						if(file_type.equals("xls"))
		      							reFileName=str_year+str_month+str_day+str_hour+str_minute+str_second+".xls";
		       						if(file_type.equals("pps"))
		      							reFileName=str_year+str_month+str_day+str_hour+str_minute+str_second+".pps";
		       						if(file_type.equals("zip"))
		       							reFileName=str_year+str_month+str_day+str_hour+str_minute+str_second+".zip";
		       						if(file_type.equals("rar"))
		       							reFileName=str_year+str_month+str_day+str_hour+str_minute+str_second+".rar";
						
    		     					append_file1 = FileName;		
    		     					appfile_link = saveDirectory+append_file1;
						
											File f1 = new File(appfile_link);
											append_file1 = reFileName;		
					    		    appfile_link = saveDirectory+reFileName;
											File f2 = new File(appfile_link);
											f1.renameTo(f2);
						
											System.out.println("上傳之檔案appfile_link ="+FileName);
    		     			}
    		     		  else
    		     		  {
    		     		     	append_file1 =" ";   		
    		     			   	appfile_link = " ";
    		     		  }
    		   } 
    		   
    		   int iStartYear = Integer.parseInt(StartYear)+1911;
								StartYear = Integer.toString(iStartYear);		 					 
			    		  notify_date = StartMonth+"/"+StartDay+"/"+StartYear+" "+StartHour+":00:00";			    		   			    		   
			    		  System.out.println("Notify_DATE ="+notify_date); 
    		   int iEndYear = Integer.parseInt(EndYear)+1911;
								EndYear = Integer.toString(iEndYear);					 
			    		  notify_end_date =EndMonth+"/"+EndDay+"/"+EndYear;
			    		  System.out.println("Notify_End_DATE ="+notify_end_date);		 
    		   List WTT01_USERNAME = TakeUserName(muser_id);    		         		       		    	    		   
    		   String user_name = (String)((DataObject)WTT01_USERNAME.get(0)).getValue("muser_name");
    		   List WLX_SEQNO = TakeMaxSeqno();      
    		   int seqcount = Integer.parseInt( (((DataObject)WLX_SEQNO.get(0)).getValue("maxseqno")==null )? "-1" : ((DataObject)WLX_SEQNO.get(0)).getValue("maxseqno").toString()); 
   		     seqcount++;		  	       	        	     
    	     actMsg = InsertWLX_Notify(seqcount,headmark,muser_id,user_name,append_file1,notify_date,notify_end_date,appfile_link,notifyUrl);   	        	     	        	    
        	 rd = application.getRequestDispatcher( ListPgName +"?act=List");          	              
			}    	 
    	else if(act.equals("Update")){	     		  
	  	   	String tnotify_url=" ",theadmark=" ",tnotify_date=" ",tnotify_end_date=" ",tappend_file=" ",tappfile_link=" ",tuser_id=" ",tuser_name=" ",tupdate_date=" ",tuser_id_c=" ",tuesr_name_c=" ",tupdate_type_c="U";
    	    List WLX_Notify = Get_WLX_Notify(seq_no);    		  
        	List WTT01_USERNAME = TakeUserName(muser_id);
 				  theadmark = (((DataObject)WLX_Notify.get(0)).getValue("headmark")==null ) ? "" : (String)((DataObject)WLX_Notify.get(0)).getValue("headmark");	
		      tnotify_date = (((DataObject)WLX_Notify.get(0)).getValue("notify_date")==null ) ? "" : (String)((DataObject)WLX_Notify.get(0)).getValue("notify_date");	
		      tappend_file = (((DataObject)WLX_Notify.get(0)).getValue("append_file")==null ) ? " " : (String)((DataObject)WLX_Notify.get(0)).getValue("append_file");	
		      tappfile_link = (((DataObject)WLX_Notify.get(0)).getValue("appfile_link")==null ) ? " " : (String)((DataObject)WLX_Notify.get(0)).getValue("appfile_link");	
		      tnotify_end_date = (((DataObject)WLX_Notify.get(0)).getValue("notify_end_date")==null ) ? "" : (String)((DataObject)WLX_Notify.get(0)).getValue("notify_end_date");	
		      tuser_id = (String)((DataObject)WLX_Notify.get(0)).getValue("user_id");		        	   		  
		      tuser_name = (((DataObject)WLX_Notify.get(0)).getValue("user_name")==null ) ? "" : (String)((DataObject)WLX_Notify.get(0)).getValue("user_name");	
		      tupdate_date = (((DataObject)WLX_Notify.get(0)).getValue("update_date")==null ) ? "" : (String)((DataObject)WLX_Notify.get(0)).getValue("update_date");	
		      tuser_id_c = muser_id;
		      tuesr_name_c = (String)((DataObject)WTT01_USERNAME.get(0)).getValue("muser_name"); 		      	   		    		        	   		  
		     	tnotify_url = (((DataObject)WLX_Notify.get(0)).getValue("notify_url")==null ) ? "" : (String)((DataObject)WLX_Notify.get(0)).getValue("notify_url");	
		      String msg = InsertWLX_Notify_LOG(seq_no,theadmark,tnotify_date,tappend_file,tnotify_end_date,tuser_id,tuser_name,tupdate_date,tuser_id_c,tuesr_name_c,tupdate_type_c,tnotify_url);        	   		          	   		          	   		 									    		         		       		      		   
    		    														  
	    		String saveDirectory = Utility.getProperties("notifyDir");
    		  int maxPostSize = 100*1024*1024;   		  
    		  File objFile = new File(saveDirectory);		  
    		  File FileAry[] = objFile.listFiles();
    		  if(FileAry.length!=0)
    		  {      		   
    			//宣告上傳檔案名稱
    			String FileName = null;
    			//支援中文檔名
    			String enCoding = "MS950";
    			//產生一個新的MultipartRequest的物件,multi
    			MultipartRequest  multi = new MultipartRequest(request,saveDirectory,maxPostSize,enCoding);
    		  //取得所有上傳之檔案輸入型態名稱
    		  Enumeration filesname = multi.getFileNames();
    		  if(filesname.hasMoreElements())
    		  {	      		                		   		    
    		    	String nextFilename = (String)filesname.nextElement();
    		     	FileName = multi.getFilesystemName(nextFilename); 
    		     	System.out.println("上傳之更新檔案FileName ="+FileName);
    		    
    		    	if( !(FileName == null))
    		     	{
    		     		tappend_file = FileName;   		
    		     		tappfile_link = saveDirectory+tappend_file;
					
						Calendar calendar=new GregorianCalendar();
						int year=calendar.get(Calendar.YEAR);
						String str_year=Integer.toString(year);
						int month=calendar.get(Calendar.MONTH);
						String str_month=Integer.toString(month);
						int day=calendar.get(Calendar.DATE);
						String str_day=Integer.toString(day);
						int hour=calendar.get(Calendar.HOUR_OF_DAY);
						String str_hour=Integer.toString(hour);
						int minute=calendar.get(Calendar.MINUTE);
						String str_minute=Integer.toString(minute);
						int second=calendar.get(Calendar.SECOND);
						String str_second=Integer.toString(second);
						int index = FileName.indexOf('.');      
      					String file_type = FileName.substring(index+1,index+4);
						String	reFileName="";
					
       	  			    if(file_type.equals("doc")){
      				       reFileName=str_year+str_month+str_day+str_hour+str_minute+str_second+".doc";
      			        }				
       					if(file_type.equals("xls")){
      					   reFileName=str_year+str_month+str_day+str_hour+str_minute+str_second+".xls";
      					}		
       					if(file_type.equals("pps")){
      					   reFileName=str_year+str_month+str_day+str_hour+str_minute+str_second+".pps";
      					}		
       					if(file_type.equals("zip")){
       					   reFileName=str_year+str_month+str_day+str_hour+str_minute+str_second+".zip";
       					}		
       					if(file_type.equals("rar")){
       					   reFileName=str_year+str_month+str_day+str_hour+str_minute+str_second+".rar";
       					}	
       					//102.12.03 add pdf上傳 by 2295	
       					if(file_type.equals("pdf")){
       					   reFileName=str_year+str_month+str_day+str_hour+str_minute+str_second+".pdf";	
       					}			
						File f1 = new File(tappfile_link);
						tappend_file = reFileName;		
    		     		tappfile_link = saveDirectory+reFileName;
						File f2 = new File(tappfile_link);
						f1.renameTo(f2);
						System.out.println("上傳之檔案appfile_link ="+FileName);
    		     }
    		     else
    		     {
    		        tappend_file =" ";   		
    		     		tappfile_link = " ";
    		     }
    		 }
    	} 
    		   int iStartYear = Integer.parseInt(StartYear)+1911;
								 StartYear = Integer.toString(iStartYear);		 					 
			    		   notify_date = StartMonth+"/"+StartDay+"/"+StartYear+" "+StartHour+":00:00";			    		   			    		   
			    		   System.out.println("Notify_DATE ="+notify_date); 
    		   int iEndYear = Integer.parseInt(EndYear)+1911;
								 EndYear = Integer.toString(iEndYear);					 
			    		   notify_end_date =EndMonth+"/"+EndDay+"/"+EndYear;
			    		   System.out.println("Notify_End_DATE ="+notify_end_date);	
		 			 int seqcount = Integer.parseInt(seq_no);
		 			 
     if(tappend_file.equals(" "))
				UpdateWLX_Notify_Nofile(seqcount,headmark,tuser_id_c,tuesr_name_c,notify_date,notify_end_date,notifyUrl);
		 else             
		 		UpdateWLX_Notify(seqcount,headmark,tuser_id_c,tuesr_name_c,tappend_file,tappfile_link,notify_date,notify_end_date,notifyUrl);
		 rd = application.getRequestDispatcher( ListPgName +"?act=List");              
     }  
    	request.setAttribute("actMsg",actMsg);    
    }            
	  try
	  {
        //forward to next present jsp
        rd.forward(request, response);
    } 
    catch(NullPointerException npe){} 
    
 }//end of doProcess
%>


<%!
    private final static String nextPgName = "/pages/ActMsg.jsp";    
    private final static String EditPgName = "/pages/ZZ091W_Edit.jsp";    
    private final static String ListPgName = "/pages/ZZ091W_List.jsp";        
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    private boolean CheckPermission(HttpServletRequest request){//檢核權限    	    
    	   boolean CheckOK=false;
    	    HttpSession session = request.getSession();            
            Properties permission = ( session.getAttribute("WLX_Notify")==null ) ? new Properties() : (Properties)session.getAttribute("WLX_Notify");				                
            if(permission == null){
              System.out.println("WLX_Notify.permission == null");
            }else{
               System.out.println("WLX_Notify.permission.size ="+permission.size());
               
            }
           
        	return true;
    }           
    
    //取得該標題之公告資料
    private List Get_WLX_Notify(String seq_no){
    		//查詢條件    
    		List paramList =new ArrayList() ;
    		StringBuffer sqlCmd = new StringBuffer() ;
    		sqlCmd.append("select seq_no,headmark,notify_url,to_char(notify_date,'mm/dd/yyyy hh24:mi:ss') as notify_date");
    		sqlCmd.append(",to_char(notify_end_date,'mm/dd/yyyy hh24:mi:ss') as notify_end_date,append_file,user_id");
    		sqlCmd.append(",user_name,to_char(update_date,'mm/dd/yyyy hh24:mi:ss') as update_date,appfile_link ");
    		sqlCmd.append(" from WLX_Notify where  seq_no =? ") ;
    		paramList.add(seq_no) ;
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"seq_no");            
        return dbData;
    }
    
    //取得該序號之最大值
    private List TakeMaxSeqno(){
    		//查詢條件    
    		String sqlCmd = "select max(seq_no) as maxseqno from WLX_Notify";  		
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,null,"maxseqno");            
        return dbData;
    }
    
    //取得該user_id之user_name
    private List TakeUserName(String muser_id){
    		//查詢條件    
    		List paramList = new ArrayList() ;
    		String sqlCmd = "select *  from WTT01 where  muser_id =? ";  
    		paramList.add(muser_id) ;
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
        return dbData;
    }
    
 
    //insert新公告    
    private String InsertWLX_Notify(int seq_no,String headmark,String muser_id,String user_name,String append_file1,String notify_date,String notify_end_date,String appfile_link,String notify_url)throws Exception{   	    
			String	errMsg=" ";
			List updateDBSqlList = new LinkedList();	
			StringBuffer sqlCmd = new StringBuffer () ;
			List paramList =new ArrayList() ;
			sqlCmd.append("INSERT INTO WLX_Notify(seq_no,headmark,notify_date,notify_end_date,user_id,user_name,update_date,append_file,appfile_link,notify_url) ");
			sqlCmd.append(" VALUES (?,?,TO_TIMESTAMP(?,'mm/dd/yyyy hh24:mi:ss')" 						   																				 	
			       + ",TO_DATE(?,'mm/dd/yyyy')" 
			       + ",?" 	
			       + ",?" 
			       + ",sysdate"			      
			       + ",?" 																       
			       + ",?"
			       + ",? )" );
			paramList.add(seq_no) ;
			paramList.add(headmark);
			paramList.add(notify_date) ;
			paramList.add(notify_end_date) ;
			paramList.add(muser_id) ;
			paramList.add(user_name) ;			
			paramList.add(append_file1) ;
			paramList.add(appfile_link);
			paramList.add(notify_url) ;
		  //updateDBSqlList.add(sqlCmd); 	
		  //DBManager.updateDB(updateDBSqlList,"ZZ091W.jsp");
		  this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList) ;
		  return errMsg;	
    }
    
    
    private void Del_WLX_Notify(String seq_no) throws Exception{
	  //List updateDBSqlList = new LinkedList();
	  List paramList =  new ArrayList() ;
	  String sqlCmd = "Delete from WLX_Notify  where seq_no=? ";			   
 	  paramList.add(seq_no) ;	  			  
	  //updateDBSqlList.add(sqlCmd);                     
	  //DBManager.updateDB(updateDBSqlList,"");
	  this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
	  }
	  
		//insert新公告LOG
	  private String InsertWLX_Notify_LOG(String tseq_no,String theadmark,String tnotify_date,String tappend_file,String tnotify_end_date,String tuser_id,String tuser_name
			  ,String tupdate_date,String tuser_id_c,String tuesr_name_c,String tupdate_type_c,String tnotify_url) throws Exception
	  {	   
	  	String	errMsg=" ";
	  	List paramList =new ArrayList() ;
			//List updateDBSqlList = new LinkedList();				      		
			String sqlCmd = "INSERT INTO WLX_Notify_LOG(seq_no,headmark,append_file,notify_end_date,user_id,user_name,user_id_c,user_name_c,notify_url,notify_date,update_date_c,update_type_c)"
			+" VALUES (?,?"
									+ ",?"
									+ ",TO_TIMESTAMP(?,'mm/dd/yyyy hh24:mi:ss')"
									+ ",?"
									+ ",?"
									+ ",?"
									+ ",?"	
									+ ",?"
			                        + " ,TO_TIMESTAMP(?,'mm/dd/yyyy hh24:mi:ss')"
									+ ",sysdate"												
			                         + ",?)"	;	
		  paramList.add(tseq_no) ;	            
		  paramList.add(theadmark) ;
		  paramList.add(tappend_file) ;
		  paramList.add(tnotify_end_date) ;
		  paramList.add(tuser_id);
		  paramList.add(tuser_name) ;
		  paramList.add(tuser_id_c) ;
		  paramList.add(tuesr_name_c) ;
		  paramList.add(tnotify_url);
		  paramList.add(tnotify_date) ;
		  paramList.add(tupdate_type_c) ;
		  this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
		  //updateDBSqlList.add(sqlCmd); 	
		  //DBManager.updateDB(updateDBSqlList,"ZZ091W.jsp");		  
		  return errMsg;	  
	}

        //Update公告的附加檔案
	  private String UpdateWLX_Notify_AppendFile(String seq_no) throws Exception{	   
	  	String	errMsg=" ",empty=" ";
		//	List updateDBSqlList = new LinkedList();
		List paramList =new ArrayList() ;
			String sqlCmd = "UPDATE WLX_Notify SET "				   	   						    	   						   
					+ "append_file=' '" 	
					+ ",appfile_link=' '" 				   									       				            		 						       
					+ " where seq_no=? ";
			paramList.add(seq_no) ;
			this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
		//  updateDBSqlList.add(sqlCmd); 	
		//  DBManager.updateDB(updateDBSqlList,"ZZ091W.jsp");		  
		  return errMsg;	  
	  }
	
	 //Update公告內容(有要更新檔案)
	  private String UpdateWLX_Notify(int seq_no,String headmark,String muser_id,String user_name,String append_file,String appfile_link,String notify_date
			  ,String notify_end_date,String notify_Url) throws Exception{	   
	  	  String	errMsg=" ",empty=" ";
		  //List updateDBSqlList = new LinkedList();
		  List paramList =new ArrayList () ;
		  String sqlCmd = "UPDATE WLX_Notify SET "				   	   						    	   						   
					+ "headmark=?"
					+ ",user_id=?"
					+ ",user_name=?"
					+ ",append_file=?"
					+ ",appfile_link=?"
					+ ",notify_date=TO_TIMESTAMP(?,'mm/dd/yyyy hh24:mi:ss')"
					+ ",notify_end_date=TO_TIMESTAMP(?,'mm/dd/yyyy hh24:mi:ss')"
					+ ",notify_Url=?"					   									       				            		 						       
					+ " where seq_no=?";	
		  paramList.add(headmark) ;
		  paramList.add(muser_id) ;
		  paramList.add(user_name) ;
		  paramList.add(append_file);
		  paramList.add(appfile_link) ;
		  paramList.add(notify_date);
		  paramList.add(notify_end_date);
		  paramList.add(notify_Url) ;
		  paramList.add(seq_no);
		  this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
		  //updateDBSqlList.add(sqlCmd); 	
		  //DBManager.updateDB(updateDBSqlList,"ZZ091W.jsp");		  
		  return errMsg;	  
	  }
	//Update公告內容(沒有要更新檔案)
	  private String UpdateWLX_Notify_Nofile(int seq_no,String headmark,String muser_id,String user_name,String notify_date,String notify_end_date,String notify_Url)
	  throws Exception{	   
	  	String	errMsg=" ",empty=" ";
		  //List updateDBSqlList = new LinkedList();
		  List paramList = new ArrayList() ;
		  String sqlCmd = "UPDATE WLX_Notify SET "				   	   						    	   						   
					+ "headmark=?"
					+ ",user_id=?"
					+ ",user_name=?"
					+ ",notify_date=TO_TIMESTAMP(?,'mm/dd/yyyy hh24:mi:ss')"
					+ ",notify_end_date=TO_TIMESTAMP(?,'mm/dd/yyyy hh24:mi:ss')"
					+ ",notify_Url=?"					   									       				            		 						       
					+ " where seq_no=?";
		  paramList.add(headmark) ;
		  paramList.add(muser_id) ;
		  paramList.add(user_name);
		  paramList.add(notify_date) ;
		  paramList.add(notify_end_date);
		  paramList.add(notify_Url) ;
		  paramList.add(seq_no);
		  //updateDBSqlList.add(sqlCmd); 	
		  //DBManager.updateDB(updateDBSqlList,"ZZ091W.jsp");	
		  this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
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