<%
// 98.07.30 add 農業金庫財務資料上傳作業 by 2295
// 98.09.10 add 1.金庫上傳檔案至農金局DB主機 by 2295
//              2.瀏覽該檔案時,至農金局DB主機抓回 by 2295
//              3.至農金局DB主機下載回來,並上傳至檢查局 by 2295
// 98.09.28 add Download取得遠端主機目錄,增加上一層serverRptDir by 2295
// 99.10.13 add 使用PreparedStatement;並列印轉換後的SQL;套用QueryDB_SQLParam by 2295
//100.06.01 fix 日期顯示 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.Utility_WM" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.report.*" %>
<%@ page import="com.tradevan.util.ftp.MyFTPClient" %>
<%@ page import="com.oreilly.servlet.MultipartRequest" %>
<%@ page import="java.lang.Integer" %>
<%@ page import="java.util.*,java.io.*" %>
<%@ page import="java.text.SimpleDateFormat" %>
<%@include file="./include/Header.include" %>
<%
	   String CopyResult = "";
	   String fileName="";
	   String mail_subject="";
	   String mail_content="";
	   DataObject bean = null;
	   String update_date = "";

       RPT_CODE = Utility.getTrimString(dataMap.get("RPT_CODE"));
	   RPT_CODE = RPT_CODE.equals("")?(Utility.getTrimString(dataMap.get("Report_no"))):RPT_CODE;
	   S_YEAR = Utility.getTrimString(dataMap.get("S_YEAR"));
	   S_MONTH = Utility.getTrimString(dataMap.get("S_MONTH"));
	   String user_id = (session.getAttribute("muser_id") == null)?"":(String)session.getAttribute("muser_id");
	   String user_name = (session.getAttribute("muser_name") == null)?"":(String)session.getAttribute("muser_name");
	   String firstStatus = ( request.getParameter("firstStatus")==null ) ? "" : (String)request.getParameter("firstStatus");
	   String rptLine = Utility.getTrimString(dataMap.get("rptLine"));
       String rptIP=Utility.getProperties("rptIP");
	   String rptID=Utility.getProperties("rptID");
	   String rptPwd=Utility.getProperties("rptPwd");

       if(!Utility.CheckPermission(request,report_no)){//無權限時,導向到LoginError.jsp
           rd = application.getRequestDispatcher( LoginErrorPgName );
       }else{
           if(act.equals("Qry")){
    	       rd = application.getRequestDispatcher( EditPgName +"?act=Qry&firstStatus="+firstStatus);
    	   }else if(act.equals("List")){
    	       List dbData = getQryResult(RPT_CODE,S_YEAR,S_MONTH);
    	       request.setAttribute("reportList",dbData);
    	       rd = application.getRequestDispatcher( EditPgName +"?act=Qry&RPT_CODE="+RPT_CODE+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH);
       	   }else if(act.equals("Upload")){//金庫上傳檔案至農金局DB主機
       	       System.out.println("Upload Dir="+AgriBankDir);
       	       fileName = Utility.parseFileName(request.getParameter("FileName"));
       	       ver = subdir(RPT_CODE,S_YEAR,S_MONTH);
       	       AgriBankDir = Utility.getProperties("AgriBankDir")+System.getProperty("file.separator")+RPT_CODE+System.getProperty("file.separator")+S_YEAR+S_MONTH+System.getProperty("file.separator")+ver;
       	       actMsg = actMsg + mkdir(AgriBankDir);//新增目錄

       	       //int maxPostSize = 5 * 1024 * 1024;
               //MultipartRequest multi = new MultipartRequest(request, saveDirectory, maxPostSize, "UTF-8");
       	       int UploadSize = Integer.parseInt(Utility.getProperties("UploadSize"));
       	       MultipartRequest multi = new MultipartRequest(request, AgriBankDir, UploadSize  * 1024, "UTF-8");

       	       if(actMsg.equals("")){

       	          AgriBankDir_remote = RPT_CODE+"/"+S_YEAR+S_MONTH+"/"+ver;
       	          //上傳至農金局DB主機
       	          actMsg = uploadAgriRpt(RPT_CODE,AgriBankDir,AgriBankDir_remote,fileName);

       	          if(actMsg.equals("")){
       	             actMsg = multi.getFilesystemName("UpFileName")+"<br>檔案上傳至農金局已完成<br>";
       	             actMsg += updateDB(RPT_CODE,S_YEAR,S_MONTH,fileName,user_id,user_name);//更新rpt_dirf
       	             if(actMsg.indexOf("寫入資料庫成功") != -1){
       	                List dbData = getQryResult(RPT_CODE,S_YEAR,S_MONTH);
       	                bean = (DataObject)dbData.get(0);
       	                if((String)bean.getValue("s_updatedate") != null){
                    	   update_date = (String)bean.getValue("s_updatedate");
                    	   //100.06.01 fix 日期顯示 by 2295
                    	   if(update_date.startsWith("1")){
                              update_date = update_date.substring(0,3)+"/"+update_date.substring(3,5)+"/"+update_date.substring(5,7)+ " " + update_date.substring(8,update_date.length());	                    		
                           }else{
                          	  update_date = update_date.substring(0,2)+"/"+update_date.substring(2,4)+"/"+update_date.substring(4,6)+ " " + update_date.substring(7,update_date.length());	                    		
                           }                         	   
                    	   System.out.println("update_date"+update_date);
                        }
       	                mail_subject = S_YEAR + "年" + S_MONTH + "月-"+(String)bean.getValue("rpt_name")+ "-檔案上傳完成";
       	                mail_content = "報表名稱："+(String)bean.getValue("rpt_name")+"\n";
       	                mail_content += "上傳檔名："+(String)bean.getValue("rpt_fname")+ "\n";
	       	            mail_content += "申報年月：" + S_YEAR + "年" + S_MONTH + "月\n";
	       		        mail_content += "申報日期：" + update_date + "\n\n";
	                    mail_content += "煩請至MIS管理系統及檢查缺失追蹤管理系統平台確認報表正確性!!";
       	                //上傳完成後.發e-mail
       	                Utility.sendMail(Utility.getProperties("feb_Agri_Addr"),mail_subject,mail_content);
       	             }//end of write db
       	          }//end of uploadAgriRpt success
       	       }

       	       rd = application.getRequestDispatcher( nextPgName+"?goPages=RptFileUpload.jsp&act=List&Report_no="+RPT_CODE+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&test=nothing");
           }else if(act.equals("UploadFeb")){
               //1.至農金局DB主機下載檔案
               //2.上傳該檔案至檢查局
           	   File logDir  = new File(Utility.getProperties("logDir"));
   			   if(!logDir.exists()){
        		   if(!Utility.mkdirs(Utility.getProperties("logDir"))){
    	   			  System.out.println("目錄新增失敗");
        		   }
   			   }
               logfile = new File(logDir + System.getProperty("file.separator") + "RptFileUpload."+ logfileformat.format(nowlog));
   			   System.out.println("logfile filename="+logDir + System.getProperty("file.separator") +"RptFileUpload."+ logfileformat.format(nowlog));
   			   logos = new FileOutputStream(logfile,true);
   			   logbos = new BufferedOutputStream(logos);
   			   logps = new PrintStream(logbos);
    		   System.out.println("=============執行金庫財務報表上傳至檢查局開始===========");
    		   logfile = new File(logDir + System.getProperty("file.separator") + "RptFileUpload."+ logfileformat.format(nowlog));
  			   System.out.println("logfile filename="+logDir + System.getProperty("file.separator") +"RptFileUpload."+ logfileformat.format(nowlog));
   			   logos = new FileOutputStream(logfile,true);
   			   logbos = new BufferedOutputStream(logos);
   			   logps = new PrintStream(logbos);

  			   actMsg += uploadFeb(rptLine,logps);
   			   System.out.println("actMsg = "+actMsg);
   			   logcalendar = Calendar.getInstance();
   			   nowlog = logcalendar.getTime();
   			   logps.println(logformat.format(nowlog)+actMsg);
   			   logps.flush();
   			   System.out.println("=============執行金庫財務報表上傳至檢查局結束===========");
   			   out.print(actMsg);
    	       rd = application.getRequestDispatcher( nextPgName+"?goPages=RptFileUpload.jsp&act=List&Report_no="+RPT_CODE+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&test=nothing");
       	   }else if(act.equals("Download")){//下載已上傳至農金局DB主機的檔案
       	       ver = ( request.getParameter("rpt_version")==null ) ? "" : (String)request.getParameter("rpt_version");

    	       File ClientRptDir = new File(Utility.getProperties("ClientRptDir"));
	           if(!ClientRptDir.exists()){
         		   if(!Utility.mkdirs(Utility.getProperties("ClientRptDir"))){
         	     		actMsg=actMsg+Utility.getProperties("ClientRptDir")+"目錄新增失敗";
         		   }
        	   }
        	   if(actMsg.equals("")){
        	      List uploadfile = getUploadfile(RPT_CODE,S_YEAR,S_MONTH,ver);
        	      bean = (DataObject)uploadfile.get(0);
        	      System.out.println((String)bean.getValue("rpt_fname"));
        	      List filename_List = new LinkedList();
        	      filename_List.add(Utility.toBig5Convert((String)bean.getValue("rpt_fname")));

    	       	  MyFTPClient ftpC = new MyFTPClient(rptIP, rptID, rptPwd);
    	       	  System.out.println("AgriBankDir_remote="+Utility.getProperties("serverRptDir")+Utility.getProperties("AgriBankDir_remote")+"/"+RPT_CODE+"/"+S_YEAR+S_MONTH+"/"+ver);
    	       	  System.out.println("ClientRptDir="+Utility.getProperties("ClientRptDir")+System.getProperty("file.separator"));
    	       	  //MyFTPClient.getFiles(String remote_path, String local_path,List filename)
               	  actMsg = ftpC.getFiles(Utility.getProperties("serverRptDir")+Utility.getProperties("AgriBankDir_remote")+"/"+RPT_CODE+"/"+S_YEAR+S_MONTH+"/"+ver, Utility.getProperties("ClientRptDir")+System.getProperty("file.separator"),filename_List);
               	  ftpC=null;
               	  rd = application.getRequestDispatcher( downloadPgName+"?rpt_code="+RPT_CODE+"&m_year="+S_YEAR+"&m_month="+S_MONTH+"&rpt_version="+ver+"&test=nothing");
               }
		   }

       	   request.setAttribute("actMsg",actMsg);
       	   request.setAttribute("alertMsg",alertMsg);
       	   request.setAttribute("webURL_Y",webURL_Y);
       	   request.setAttribute("webURL_N",webURL_N);

       	   rd.forward(request,response);
       }


%>

<%!
    private final static String report_no = "RptFileUpload";
    private final static String nextPgName = "/pages/ActMsg.jsp";
    private final static String EditPgName = "/pages/"+report_no+"_Edit.jsp";
    private final static String downloadPgName = "/pages/DownloadFile.jsp";
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    private String S_YEAR="";
    private String S_MONTH="";
    private String RPT_CODE="";
    private String AgriBankDir = "";
    private String AgriBankDir_remote = "";
    private String ver = "";
    private File logfile;
	private FileOutputStream logos=null;
	private BufferedOutputStream logbos = null;
	private PrintStream logps = null;
	private Date nowlog = new Date();
	private SimpleDateFormat logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");
	private SimpleDateFormat logfileformat = new SimpleDateFormat("yyyyMMddHHmmss");
	private Calendar logcalendar;

    private String mkdir(String filedir){//檢核目錄不存在時,則新增該目錄
    		String mkdirOK="";
    		String ver="";
    		try{

    			File AgriBankDir = new File(filedir);
	        	if(!AgriBankDir.exists()){
         			if(!Utility.mkdirs(filedir)){
         		   		mkdirOK=mkdirOK+filedir+"目錄新增失敗";
         			}
        		}

        	}catch(Exception e){
        		System.out.println("目錄新增失敗"+e.getMessage()) ;
        		mkdirOK="目錄新增失敗";
        	}
        	return mkdirOK;
    }

	private String subdir(String report_no,String m_year,String m_month){
	    List paramList = new ArrayList();
	    String sqlCmd = "";
	    String report_ver = "001";
	    try{
    		sqlCmd = " select max(rpt_version) as rpt_version "
				   + " from agrirpt_dirf "
				   + " where rpt_code=?"
				   + " and m_year=?"
				   + " and m_month=?";
    		paramList.add(report_no);
    		paramList.add(m_year);
    		paramList.add(m_month);
        	List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"rpt_version");
        	if(dbData.size() != 0){
           	   report_ver = String.valueOf(Integer.parseInt((((DataObject)dbData.get(0)).getValue("rpt_version")).toString())+1);
           	   if(report_ver.length()==1) report_ver = "00"+report_ver;
           	   if(report_ver.length()==2) report_ver = "0"+report_ver;
        	}

		}catch(Exception e){
		   System.out.println(e+":"+e.getMessage());
		}
		return report_ver;
	}

	//取得查詢結果
    private List getQryResult(String report_no,String m_year,String m_month){
    		//查詢條件
    		StringBuffer sqlCmd = new StringBuffer();
			List paramList = new ArrayList();
    		sqlCmd.append(" select agrirpt_dirf.m_year,agrirpt_dirf.m_month,");
    		sqlCmd.append("       (to_char(agrirpt_dirf.m_year)|| LPAD(to_char(agrirpt_dirf.m_month),2,'0')) as s_yymm, ");
	   		sqlCmd.append(" 		  agrirpt_dirf.rpt_fname,agrirpt_dirf.rpt_version,rpt_nof.RPT_NAME,feb_flag, ");
	   		sqlCmd.append("	  	 ((to_char(agrirpt_dirf.UPDATE_DATE,'yyyymmdd')-19110000)  ||  to_char(agrirpt_dirf.UPDATE_DATE,' hh24:mi'))  as s_UpdateDate,");
	   		sqlCmd.append("		  decode(agrirpt_dirf.uploaddate,null,'',((to_char(agrirpt_dirf.uploaddate,'yyyymmdd')-19110000)  ||  to_char(agrirpt_dirf.uploaddate,' hh24:mi')))  as s_uploaddate ");
			sqlCmd.append(" from agrirpt_dirf left join rpt_nof on agrirpt_dirf.RPT_CODE = rpt_nof.RPT_CODE ");
			sqlCmd.append(" where agrirpt_dirf.rpt_code = ?");
			sqlCmd.append(" and to_char(agrirpt_dirf.m_year * 100 + agrirpt_dirf.m_month) = ?");
			sqlCmd.append(" order by rpt_version desc");
			paramList.add(report_no);
			paramList.add(m_year+m_month);
			
    		List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
    		if(dbData != null && dbData.size() != 0){
    		   System.out.println("dbData.size="+dbData.size());
    		}else{
    		   System.out.println("dbData is null or size = 0");
    		}

			return dbData;
	}
    private String updateDB(String report_no,String m_year,String m_month,String filename,String user_id,String user_name){        
        List updateDBList = new ArrayList();//0:sql 1:data		
	    List updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data
        StringBuffer sqlCmd = new StringBuffer();
		String errMsg="";
		try{
    		File tmpFile = new File(AgriBankDir+System.getProperty("file.separator")+filename);
        	sqlCmd.append("insert into agrirpt_dirf values(?,?,?,?,?,?,'N',null,?,?,sysdate)");
        	dataList.add(report_no);
        	dataList.add(m_year);
        	dataList.add(m_month);
        	dataList.add(filename);
        	dataList.add(ver);
        	dataList.add(String.valueOf(tmpFile.length()));
        	dataList.add(user_id);
        	dataList.add(user_name);
        	updateDBDataList.add(dataList);//1:傳內的參數List
        	
        	updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql	        	
			updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
			updateDBList.add(updateDBSqlList);
        	if(DBManager.updateDB_ps(updateDBList)){
				errMsg = errMsg + "相關資料寫入資料庫成功";
			}else{
				errMsg = errMsg + "相關資料寫入資料庫失敗";
			}
		}catch(Exception e){
		   System.out.println(e+":"+e.getMessage());
			errMsg = errMsg + "相關資料寫入資料庫失敗";
		}
		return errMsg;
    }

    //金庫財報上傳至檢查局
    private String uploadFeb(String rptLine,PrintStream logps){
        String errMsg="";
        String CopyResult = "";
	    String fileName="";
        StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();			
        String putMsg="";
        List filename_List = new LinkedList();
        File tmpFile = null;
        List updateDBList = new ArrayList();//0:sql 1:data		
	    List updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data
		try{
    		String rptIP_feb=Utility.getProperties("rptIP_feb");
	        String rptID_feb=Utility.getProperties("rptID_feb");
	        String rptPwd_feb=Utility.getProperties("rptPwd_feb");
	        String rptIP=Utility.getProperties("rptIP");
	        String rptID=Utility.getProperties("rptID");
	        String rptPwd=Utility.getProperties("rptPwd");
    		File febxlsDir_AgriBank = new File(Utility.getProperties("febxlsDir_AgriBank"));
	       	if(!febxlsDir_AgriBank.exists()){
         		if(!Utility.mkdirs(Utility.getProperties("febxlsDir_AgriBank"))){
         		 	errMsg+=Utility.getProperties("febxlsDir_AgriBank")+"目錄新增失敗";
         		}
        	}
        	File febBKxlsDir_AgriBank = new File(Utility.getProperties("febBKxlsDir_AgriBank"));
	       	if(!febBKxlsDir_AgriBank.exists()){
         		if(!Utility.mkdirs(Utility.getProperties("febBKxlsDir_AgriBank"))){
         		 	errMsg+=Utility.getProperties("febBKxlsDir_AgriBank")+"目錄新增失敗";
         		}
        	}
    		S_YEAR = rptLine.substring(0,rptLine.indexOf(":"));
		  	System.out.print("S_YEAR="+S_YEAR);
		  	rptLine = rptLine.substring(rptLine.indexOf(":")+1,rptLine.length());
		  	S_MONTH = rptLine.substring(0,rptLine.indexOf(":"));
		  	System.out.print(":S_MONTH="+S_MONTH);
		  	rptLine = rptLine.substring(rptLine.indexOf(":")+1,rptLine.length());
         	RPT_CODE = rptLine.substring(0,rptLine.indexOf(":"));
         	System.out.print(":RPT_CODE="+RPT_CODE);
         	rptLine = rptLine.substring(rptLine.indexOf(":")+1,rptLine.length());
         	ver = rptLine.substring(0,rptLine.indexOf(":")+4);
         	System.out.print(":ver="+ver);
         	sqlCmd.append(" select rpt_fname from agrirpt_dirf ");
			sqlCmd.append(" where rpt_code=?");
			sqlCmd.append(" and m_year=?");
			sqlCmd.append(" and m_month=?");
			sqlCmd.append(" and rpt_version=?");
            paramList.add(RPT_CODE);				   
            paramList.add(S_YEAR);
            paramList.add(S_MONTH);
            paramList.add(ver);
         	List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
         	if(dbData != null && dbData.size() != 0){
         	   fileName = (String)((DataObject)dbData.get(0)).getValue("rpt_fname");
         	}
         	filename_List.add(fileName);//欲上傳至檢查局的filename

        	List filename_List_iso8859 = new LinkedList();
        	filename_List_iso8859.add(Utility.toBig5Convert(fileName));//從農金局DB主機.get回來的file

    	    MyFTPClient ftpC = new MyFTPClient(rptIP, rptID, rptPwd);
    	    System.out.println("AgriBankDir_remote="+Utility.getProperties("serverRptDir")+Utility.getProperties("AgriBankDir_remote")+"/"+RPT_CODE+"/"+S_YEAR+S_MONTH+"/"+ver);
    	    System.out.println("febxlsDir_AgriBank="+Utility.getProperties("febxlsDir_AgriBank")+System.getProperty("file.separator"));
    	    //MyFTPClient.getFiles(String remote_path, String local_path,List filename)
    	    //從農金局DB主機.get欲上傳至檢查局的檔案=========================================================================================================
            errMsg = ftpC.getFiles(Utility.getProperties("serverRptDir")+Utility.getProperties("AgriBankDir_remote")+"/"+RPT_CODE+"/"+S_YEAR+S_MONTH+"/"+ver, Utility.getProperties("febxlsDir_AgriBank")+System.getProperty("file.separator"),filename_List_iso8859);
            ftpC=null;
            System.out.println("errMsg="+errMsg);
            //==============================================================================================================================================
   			if(errMsg == null){
   			    System.out.println(fileName+"檔案下載成功");
   			    errMsg = "";
         	    //將欲上傳的金庫報表.另存至febxlsDir_AgriBank(傳至檢查局的金庫報表)============================================
         	    AgriBankDir = Utility.getProperties("AgriBankDir")+System.getProperty("file.separator")+RPT_CODE+System.getProperty("file.separator")+S_YEAR+S_MONTH+System.getProperty("file.separator")+ver;
    	        ftpC = new MyFTPClient(rptIP_feb, rptID_feb, rptPwd_feb);
                //putFiles(server_host,username,password,remote_path,local_path,workDir,filename,logps)
                //從MIS APServer將已下載回來的檔案,上傳至檢查局(金庫財務資料)=================================================================
                putMsg = RptGenerateALL.putFiles(rptIP_feb, rptID_feb, rptPwd_feb,Utility.getProperties("feb_serverRptDir_"+RPT_CODE), Utility.getProperties("febxlsDir_AgriBank")+System.getProperty("file.separator"),Utility.getProperties("feb_serverRptDir_"+RPT_CODE),filename_List,logps);
                ftpC=null;
                //=============================================================================================================================
                if(putMsg == null){//上傳檔案成功
                   System.out.println("檔案上傳成功");
                   //將febxlsDir_AgriBank目錄下產生完成的檔案.並上傳成功的.搬移至所對應的bkdir=================================
                   tmpFile = new File(Utility.getProperties("febxlsDir_AgriBank")+System.getProperty("file.separator")+fileName);
                   if(!tmpFile.isDirectory() && tmpFile.exists()){
             	      CopyResult = Utility.CopyFile(Utility.getProperties("febxlsDir_AgriBank")+System.getProperty("file.separator")+fileName,Utility.getProperties("febBKxlsDir_AgriBank")+System.getProperty("file.separator")+fileName);
             	      System.out.println(Utility.getProperties("febxlsDir_AgriBank")+System.getProperty("file.separator")+fileName+" copy to "+Utility.getProperties("febBKxlsDir_AgriBank")+System.getProperty("file.separator")+fileName+" success ?? "+CopyResult);
             	      if(CopyResult.equals("0")) tmpFile.delete();
                   }

                   errMsg +=fileName+"<br>檔案上傳至檢查局完成";
                   sqlCmd.delete(0,sqlCmd.length());
                   sqlCmd.append(" update agrirpt_dirf set feb_flag='Y',uploaddate=sysdate ");
                   sqlCmd.append(" where rpt_code=?");
				   sqlCmd.append(" and m_year=?");
				   sqlCmd.append(" and m_month=?");
				   sqlCmd.append(" and rpt_version=?");
				   updateDBDataList.add(paramList);//1:傳內的參數List
				   
				   updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql		
				   updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
				   updateDBList.add(updateDBSqlList);
				  	 
				  
        		   if(DBManager.updateDB_ps(updateDBList)){
					   errMsg+=",相關資料寫入資料庫成功";
				   }else{
					   errMsg+= ",相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
				   }
                }else{//end of 上傳檔案成功
                   errMsg += fileName+"<br>上傳至Sever未成功"+putMsg;
                }
    	    }
		}catch(Exception e){
		   System.out.println(e+":"+e.getMessage());
			errMsg = errMsg + "[uploadFeb Error]";
		}
		return errMsg;
    }

    //金庫財報.上傳至DB主機
    private String uploadAgriRpt(String RPT_CODE,String AgriBankDir,String AgriBankDir_remote,String fileName){
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
            logfile = new File(logDir + System.getProperty("file.separator") + "RptFileUpload."+ logfileformat.format(nowlog));
   			System.out.println("logfile filename="+logDir + System.getProperty("file.separator") +"RptFileUpload."+ logfileformat.format(nowlog));
   			logos = new FileOutputStream(logfile,true);
   			logbos = new BufferedOutputStream(logos);
   			logps = new PrintStream(logbos);
    		System.out.println("=============執行金庫財務報表上傳至DB Server開始===========");
    		logfile = new File(logDir + System.getProperty("file.separator") + "RptFileUpload."+ logfileformat.format(nowlog));
  			System.out.println("logfile filename="+logDir + System.getProperty("file.separator") +"RptFileUpload."+ logfileformat.format(nowlog));
   			logos = new FileOutputStream(logfile,true);
   			logbos = new BufferedOutputStream(logos);
   			logps = new PrintStream(logbos);


   			//putFiles_multiSubDir(String server_host, String username, String password,
			//   				   String remote_path(主機端), String local_path(local端), String workDir(子目錄名稱), List filename(上傳檔案List),
			//   				   PrintStream logps)
   			//上傳至農金局DB主機(金庫財務資料)===============================================================================
			putMsg = RptGenerateALL.putFiles_multiSubDir(rptIP, rptID, rptPwd,Utility.getProperties("AgriBankDir_remote"), AgriBankDir+System.getProperty("file.separator"),AgriBankDir_remote,filename_List,logps);

            if(putMsg == null){//上傳檔案成功
               System.out.println("檔案上傳成功");
    	    }else{
    	       errMsg += putMsg;
    	    }
    	    System.out.println("=============執行金庫財務報表上傳至DB Server結束===========");
		}catch(Exception e){
		   System.out.println(e+":"+e.getMessage());
			errMsg = errMsg + "[uploadAgriRpt Error]";
		}
		return errMsg;
    }
    private List getUploadfile(String rpt_code,String m_year,String m_month,String rpt_version){
    	StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();			
    	sqlCmd.append(" select rpt_fname");
    	sqlCmd.append(" from agrirpt_dirf");
		sqlCmd.append(" where rpt_code=?");
		sqlCmd.append(" and m_year=?");
		sqlCmd.append(" and m_month=?");
		sqlCmd.append(" and rpt_version=?");
		paramList.add(rpt_code);
		paramList.add(m_year);
		paramList.add(m_month);
		paramList.add(rpt_version);

        List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
        return dbData;
    }
%>    