<%
// 93.12.17 add 權限檢核 by 2295
// 93.12.20 fix 若有已點選的bank_type,則以已點選的bank_type為主 by 2295
// 93.12.22 加入A02~M07之下載功能
// 93.12.22 加入A02~A05,M05之查詢功能
// 93.12.23 add 超過登入時間,請重新登入 by 2295
// 94.05.13 add 區分農/漁會科目代號 by 2295
// 94.07.13 add A02區分漁會 by 2295
// 94.07.14 fix 只有A01/A02/A03才區分農/漁會 by 2295
// 95.01.16 add F01查詢 by 2295
// 95.04.10 add A06 by 2295
// 95.05.26 add A99 by 2295
// 96.07.11 add A08 by 2295
// 96.11.16 fix A99漁會無法查詢資料 by 2295
// 97.01.02 add 97/01以後,套用新表格(增加/異動科目代號) by 2295
//          add get A09 by 2295
// 97.06.13 add A10 by 2295
// 99.09.27 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
//101.08.09 add 農信保原始檔下載 by 2295
//102.04.18 add 103/01以後,漁會套用新表格(增加/異動科目代號) by 2295  
//103.02.11 add 103/01以後,A06漁會套用新表格(增加/異動科目代號) by 2295
//104.01.09 add A12 by 2968
//104.10.12 add A13 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DownLoad" %>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.ftp.MyFTPClient" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="java.util.*,java.io.*" %>
<%@include file="./include/Header.include" %>
<%

	String S_YEAR = Utility.getTrimString(dataMap.get("S_YEAR"));
	String S_MONTH = Utility.getTrimString(dataMap.get("S_MONTH"));
    String Report_no = Utility.getTrimString(dataMap.get("Report_no"));

	//fix 93.12.18 若有已點選的tbank_no,則以已點選的tbank_no為主============================================================
	String bank_code = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");
	String nowtbank_no =  ( request.getParameter("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");
	if(!nowtbank_no.equals("")){
	   session.setAttribute("nowtbank_no",nowtbank_no);//將已點選的tbank_no寫入session
	}
	bank_code = ( session.getAttribute("nowtbank_no")==null ) ? bank_code : (String)session.getAttribute("nowtbank_no");
	//=======================================================================================================================
    //fix 93.12.20 若有已點選的bank_type,則以已點選的bank_type為主============================================================
	String bank_type = Utility.getTrimString(dataMap.get("bank_type"));
	bank_type = ( session.getAttribute("nowbank_type")==null ) ? bank_type : (String)session.getAttribute("nowbank_type");
	//=======================================================================================================================
	StringBuffer sqlCmd = new StringBuffer();
	List paramList = new ArrayList();
	if(S_MONTH.indexOf("0") == 0){	//modify by 2354 2004.12.23
	   S_MONTH = S_MONTH.substring(S_MONTH.indexOf("0")+1,S_MONTH.length());
	}
	System.out.println("S_MONTH="+S_MONTH);
	String rptIP=Utility.getProperties("rptIP");
	String rptID=Utility.getProperties("rptID");
	String rptPwd=Utility.getProperties("rptPwd");
	String filename = "";//農信保下載檔名
    if(!Utility.CheckPermission(request,report_no)){//無權限時,導向到LoginError.jsp
        rd = application.getRequestDispatcher( LoginErrorPgName );
    }else{
    	//set next jsp
    	if(act.equals("new")){
        	rd = application.getRequestDispatcher( QryPgName+"?bank_type="+bank_type );
    	}else if(act.equals("Query")){
    		 //modify and add by 2354 12.23
    		 List data_div01=null;
    	     if(Report_no.equals("A01")){
    	    	data_div01=getData_A01_A05(S_YEAR,S_MONTH,bank_code,bank_type,"01");
    	    	List data_div02=getData_A01_A05(S_YEAR,S_MONTH,bank_code,bank_type,"02");
    	    	request.setAttribute("data_div01",data_div01);
    	    	request.setAttribute("data_div02",data_div02);
    	    //============add by 2354 12.22 begin ---
    	    }else if(Report_no.equals("A02")){
    	    	data_div01=getData_A01_A05(S_YEAR,S_MONTH,bank_code,bank_type,"04");
    	    	request.setAttribute("data_div01",data_div01);
    	    }else if(Report_no.equals("A03")){
    	    	data_div01=getData_A01_A05(S_YEAR,S_MONTH,bank_code,bank_type,"05");
    	    	request.setAttribute("data_div01",data_div01);
			}else if(Report_no.equals("A04")){
    	    	data_div01=getData_A01_A05(S_YEAR,S_MONTH,bank_code,bank_type,"06");
    	    	request.setAttribute("data_div01",data_div01);
			}else if(Report_no.equals("A05")){
    	    	data_div01=getData_A01_A05(S_YEAR,S_MONTH,bank_code,bank_type,"07");
    	    	request.setAttribute("data_div01",data_div01);
    	    }else if(Report_no.equals("A06")){//95.04.10 add by 2295
    	    	data_div01=getData_A01_A05(S_YEAR,S_MONTH,bank_code,bank_type,"08");
    	    	request.setAttribute("data_div01",data_div01);
    	    }else if(Report_no.equals("A08")){//96.07.11 add by 2295
    	    	data_div01=getData_A08(S_YEAR,S_MONTH,bank_code);
    	    	request.setAttribute("data_div01",data_div01);
    	    }else if(Report_no.equals("A09")){//97.01.02 add by 2295
    	    	data_div01=getData_A09(S_YEAR,S_MONTH,bank_code);
    	    	request.setAttribute("data_div01",data_div01);
    	    }else if(Report_no.equals("A10")){//97.06.13 add by 2295
    	    	data_div01=getData_A10(S_YEAR,S_MONTH,bank_code);
    	    	request.setAttribute("data_div01",data_div01);
    	    }else if(Report_no.equals("A12")){//104.01.09 add by 2295
    	    	data_div01=getData_A12(S_YEAR,S_MONTH,bank_code);
    	    	request.setAttribute("data_div01",data_div01);
    	    }else if(Report_no.equals("A13")){//104.10.12 add by 2295
    	    	data_div01=getData_A01_A05(S_YEAR,S_MONTH,bank_code,bank_type,"12");
    	    	List data_div02=getData_A01_A05(S_YEAR,S_MONTH,bank_code,bank_type,"13");
    	    	request.setAttribute("data_div01",data_div01);
    	    	request.setAttribute("data_div02",data_div02);	
    	    }else if(Report_no.equals("A99")){//95.05.26 add by 2295
    	    	data_div01=getData_A01_A05(S_YEAR,S_MONTH,bank_code,bank_type,"99");
    	    	request.setAttribute("data_div01",data_div01);
			}else if(Report_no.equals("B01")){
    	    	data_div01=getData_B01(S_YEAR,S_MONTH,bank_code);
    	    	request.setAttribute("data_div01",data_div01);
			}else if(Report_no.equals("B02")){
    	    	data_div01=getData_B02(S_YEAR,S_MONTH,bank_code);
    	    	request.setAttribute("data_div01",data_div01);
    	    }else if(Report_no.equals("B03")){
				System.out.println("The condition is Report_no.equals(\"B03\") --Begin ");
    	    	data_div01=getData_B03(S_YEAR,S_MONTH,bank_code,Report_no,1);
    	    	request.setAttribute("data_div01",data_div01);
    	    	List data_div02=getData_B03(S_YEAR,S_MONTH,bank_code,Report_no,2);
    	    	request.setAttribute("data_div02",data_div02);
    	    	List data_div03=getData_B03(S_YEAR,S_MONTH,bank_code,Report_no,3);
    	    	request.setAttribute("data_div03",data_div03);
    	    	List data_div04=getData_B03(S_YEAR,S_MONTH,bank_code,Report_no,4);
    	    	request.setAttribute("data_div04",data_div04);
    	    	System.out.println("The condition is Report_no.equals(\"B03\") --End ");
			}else if(Report_no.equals("M01")){
    	    	data_div01=getData_M01(S_YEAR,S_MONTH,bank_code);
    	    	request.setAttribute("data_div01",data_div01);
			}else if(Report_no.equals("M02")){
    	    	data_div01=getData_M02(S_YEAR,S_MONTH,bank_code);
    	    	request.setAttribute("data_div01",data_div01);
			}else if(Report_no.equals("M03")){
				System.out.println("aaS_MONTH="+S_MONTH);
    	    	data_div01=getData_M03(S_YEAR,S_MONTH,bank_code,"C");
    	    	request.setAttribute("data_div01",data_div01);
    	    	List data_div02=getData_M03(S_YEAR,S_MONTH,bank_code,"S");
    	    	request.setAttribute("data_div02",data_div02);
			}else if(Report_no.equals("M04")){
    	    	data_div01=getData_M04(S_YEAR,S_MONTH,bank_code);
    	    	request.setAttribute("data_div01",data_div01);
			}else if(Report_no.equals("M05")){
    	    	data_div01=getData_M05(S_YEAR,S_MONTH,bank_code,"C");
    	    	request.setAttribute("data_div01",data_div01);
    	    	List data_div02=getData_M05(S_YEAR,S_MONTH,bank_code,"");
    	    	request.setAttribute("data_div02",data_div02);
    	    	List data_div03=getData_M05(S_YEAR,S_MONTH,bank_code,"N");
    	    	request.setAttribute("data_div03",data_div03);
			}else if(Report_no.equals("M06") || Report_no.equals("M07")){
    	    	data_div01=getData_M06_M07(S_YEAR,S_MONTH,bank_code,Report_no);
    	    	request.setAttribute("data_div01",data_div01);
			}else if(Report_no.equals("M08")){
    	    	data_div01=getData_M08(S_YEAR,S_MONTH,bank_code,Report_no);
    	    	request.setAttribute("data_div01",data_div01);
    	    }else if(Report_no.equals("F01")){//95.01.16 add F01查詢
    	        data_div01=getData_F01(S_YEAR,S_MONTH,bank_code);
			    request.setAttribute("data_div01",data_div01);
    	    }
    	    //modify by 2354 12.23
    	    if((data_div01.size() == 0)){
				actMsg = actMsg + "無資料可供查詢";
				request.setAttribute("actMsg",actMsg);
				rd = application.getRequestDispatcher( nextPgName );
    	    }else{
    	    	String path = ViewPgName +Report_no+".jsp?act=Query&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&test=nothing";
    	    	if(Report_no.equals("A10")) path+="&width=729";
        		rd = application.getRequestDispatcher(path);
    	    }
    	    //============add by 2354 12.23 end ---
    	}else if(act.equals("Download")){
    		List data=null;
    		//============modify by 2354 12.22 begin ---
    		if(Report_no.equals("A01") || Report_no.equals("A02") || Report_no.equals("A03") || Report_no.equals("A04") || Report_no.equals("A05") || Report_no.equals("A06") || Report_no.equals("A99") || Report_no.equals("A08") || Report_no.equals("A09") || Report_no.equals("A10") || Report_no.equals("A12") || Report_no.equals("A13")){
    		    if(Report_no.equals("A08")){
    		       data = getData_A08(S_YEAR,S_MONTH,bank_code);
    		    }else if(Report_no.equals("A09")){
    		       System.out.println("A09");
    		       data = getData_A09(S_YEAR,S_MONTH,bank_code);
    		    }else if(Report_no.equals("A10")){
    		       System.out.println("A10");
    		       data = getData_A10(S_YEAR,S_MONTH,bank_code);
    		    }else if(Report_no.equals("A12")){
     		       System.out.println("A12");
     		       data = getData_A12(S_YEAR,S_MONTH,bank_code);
    		    }else{
    		    	sqlCmd.append("select * from "+Report_no+" where m_year=? and m_month=? and bank_code=? order by acc_code ");
					paramList.add(S_YEAR);
					paramList.add(S_MONTH);
					paramList.add(bank_code);
    			    data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");
    			}
	    	}else if(Report_no.equals("M01") || Report_no.equals("M02") || Report_no.equals("M03") || Report_no.equals("M04") || Report_no.equals("M05") || Report_no.equals("M06") || Report_no.equals("M07")){
	    			sqlCmd.append("select * from "+Report_no+" where m_year=? and m_month=?");
					paramList.add(S_YEAR);
					paramList.add(S_MONTH);
    			    data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");	    	
	    	}else if(Report_no.equals("M106") || Report_no.equals("M201") || Report_no.equals("M206")){//101.08.09 add 農信保原始檔下載
    	            File ClientRptDir = new File(Utility.getProperties("ClientRptDir"));
	                if(!ClientRptDir.exists()){
         		        if(!Utility.mkdirs(Utility.getProperties("ClientRptDir"))){
         	          		actMsg=actMsg+Utility.getProperties("ClientRptDir")+"目錄新增失敗";
         		        }
        	        }
        	        if(actMsg.equals("")){
        	        	sqlCmd.append("select filename from WML01_M_UPLOAD where m_year=? and m_month=? and bank_code = ? and report_no = ?");
						paramList.add(S_YEAR);
						paramList.add(S_MONTH);
						paramList.add(bank_code);
						paramList.add(Report_no);
    			    	data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");	
    			    	
    			    	if(data.size() != 0){
    			    		filename = (String)((DataObject)data.get(0)).getValue("filename");
    			    		System.out.println("filename="+filename);
        	           		List filename_List = new LinkedList();
        	           		filename_List.add(Utility.toBig5Convert(filename));
                    
    	            	  	MyFTPClient ftpC = new MyFTPClient(rptIP, rptID, rptPwd);
    	            	  	//MyFTPClient.getFiles(String remote_path, String local_path,List filename)
                    	  	actMsg = ftpC.getFiles(Utility.getProperties("serverRptDir")+Utility.getProperties("acgfDir")+"/"+Report_no, Utility.getProperties("ClientRptDir")+System.getProperty("file.separator"),filename_List);
                    	  	ftpC=null;
                    	    
    			        }	
                    }
		   }
	    		
	       
	    	//============modify by 2354 12.22 end ---
    		if(data.size() == 0){
    		   actMsg = actMsg + "無資料可提供下載";
    		   request.setAttribute("actMsg",actMsg);
    		   rd = application.getRequestDispatcher( nextPgName );
    		}else{
    		   if(Report_no.equals("M106") || Report_no.equals("M201") || Report_no.equals("M206")){//101.08.09 add 農信保原始檔下載    		   	  
	     	      rd = application.getRequestDispatcher( ExcelPgName+"?Report_no="+Report_no+"&filename="+filename+"&test=nothing");	     	     
	           }else{	
	        	  String path = PrintDataPgName +"?Report_no="+Report_no+"&M_YEAR="+S_YEAR+"&M_MONTH="+S_MONTH+"&bank_code="+bank_code+"&test=nothing";
			      rd = application.getRequestDispatcher(path);
    		   }
    	    }
    	} 
    }

%>

<%@include file="./include/Tail.include" %>

<%!
    private final static String report_no = "WMFileDownload";
    private final static String nextPgName = "/pages/ActMsg.jsp";
    private final static String QryPgName = "/pages/"+report_no+"_Qry.jsp";
    private final static String ViewPgName = "/pages/WMFileEdit_";
    private final static String ListPgName = "/pages/"+report_no+"_List.jsp";
    private final static String PrintDataPgName = "/pages/"+report_no+"_PrintData.jsp";
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    private final static String downloadPgName = "/pages/DownloadFile.jsp";
    private final static String ExcelPgName = "/pages/WMFileDownload_Excel.jsp";
    //94.05.13 add 區分農/漁會科目代號 by 2295
    //95.05.26 add A99 by 2295
    //102.04.18 add 103/01以後,A01漁會套用新表格(增加/異動科目代號)   
    //103.02.11 add 103/01以後,A06漁會套用新表格(增加/異動科目代號)   
    private List getData_A01_A05(String S_YEAR,String S_MONTH,String bank_code,String bank_type,String acc_div){
    		//查詢條件
    		List dbData =null;
    		String Report_NO="";
    		String ncacno = "ncacno";
    		String ncacno_7 = "ncacno_7";
    		StringBuffer sqlCmd = new StringBuffer();
			List paramList = new ArrayList();
    		System.out.println("acc_div="+acc_div);
    		System.out.println("bank_type="+bank_type);

    		if(bank_type.equals("6") || (bank_type.equals("7") && !acc_div.equals("01") && !acc_div.equals("02") && !acc_div.equals("05") && !acc_div.equals("08") && !acc_div.equals("99") && !acc_div.equals("12")  && !acc_div.equals("13"))){
    		   if(acc_div.equals("01") || acc_div.equals("02")){
    		      Report_NO="A01";
    		      //97.01.02 add 97/01以後,套用新表格(增加/異動科目代號)
    		      if(Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH) >= 9701){
    		         ncacno = "ncacno_rule";
    		      }
    		   }
    		   else if(acc_div.equals("04"))                    Report_NO="A02";
    		   else if(acc_div.equals("05"))                    Report_NO="A03";
    		   else if(acc_div.equals("06"))		            Report_NO="A04";
    		   else if(acc_div.equals("07"))					Report_NO="A05";
    		   else if(acc_div.equals("08"))					Report_NO="A06";//95.04.10 add by 2295
    		   else if(acc_div.equals("99"))					Report_NO="A99";//95.05.26 add by 2295
    		   if(acc_div.equals("12") || acc_div.equals("13")){
    		      Report_NO="A13";    		     
    		   }
    		   if(Report_NO.equals("A05")){
    			  sqlCmd.append(" SELECT * ");
                  sqlCmd.append(" FROM ncacno LEFT JOIN "+Report_NO+" a01_a05 on A01_A05.acc_code = ncacno.acc_code ");
                  sqlCmd.append(" LEFT JOIN A05_ASSUMED ON A05_ASSUMED.acc_code = ncacno.acc_code ");
                  sqlCmd.append(" WHERE NCACNO.acc_div=?");
                  sqlCmd.append("   AND (A01_A05.m_year=? OR A01_A05.m_year IS NULL) ");
                  sqlCmd.append("   AND (A01_A05.m_month=? OR A01_A05.m_month IS NULL) ");
                  sqlCmd.append("   AND (A01_A05.bank_code=? OR A01_A05.bank_code IS NULL) ");
                  sqlCmd.append(" ORDER BY ncacno.acc_range");
                  paramList.add(acc_div);
                  paramList.add(S_YEAR);
                  paramList.add(S_MONTH);
                  paramList.add(bank_code);
                    
            	  dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt,assumed");
               }else{//94.07.13 add A02區分漁會//94.07.14只有A01/A02/A03才區分農/漁會
                  if(bank_type.equals("7") && (Report_NO.equals("A01") || Report_NO.equals("A02") || Report_NO.equals("A03") || Report_NO.equals("A06") || Report_NO.equals("A99")) ){
                  	 sqlCmd.delete(0,sqlCmd.length());
                  	 paramList = new ArrayList();
                  	 sqlCmd.append("select * from "+Report_NO+" A01_A05  LEFT JOIN ncacno_7 ON A01_A05.acc_code = ncacno_7.acc_code where A01_A05.m_year=? and A01_A05.m_month=? and A01_A05.bank_code=? and ncacno_7.acc_div=? order by ncacno_7.acc_range");
                     paramList.add(S_YEAR);
                     paramList.add(S_MONTH);
                     paramList.add(bank_code);
                     paramList.add(acc_div);
                     dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt,amt_3month,amt_6month,amt_1year,amt_2year,amt_over2year,amt_total");
                  }else{
            	     System.out.println("sqlcmd="+"select * from "+Report_NO+" A01_A05  LEFT JOIN "+ncacno+" ON A01_A05.acc_code = "+ncacno+".acc_code where A01_A05.m_year="+S_YEAR+" and A01_A05.m_month="+S_MONTH+" and A01_A05.bank_code='"+bank_code+"' and "+ncacno+".acc_div='"+acc_div+"' order by "+ncacno+".acc_range");
            	     sqlCmd.delete(0,sqlCmd.length());
                  	 sqlCmd.append("select * from "+Report_NO+" A01_A05  LEFT JOIN "+ncacno+" ON A01_A05.acc_code = "+ncacno+".acc_code where A01_A05.m_year=? and A01_A05.m_month=? and A01_A05.bank_code=? and "+ncacno+".acc_div=? order by "+ncacno+".acc_range");
                  	 paramList = new ArrayList();
                     paramList.add(S_YEAR);
                     paramList.add(S_MONTH);
                     paramList.add(bank_code);
                     paramList.add(acc_div);                     
            	     dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt,amt_3month,amt_6month,amt_1year,amt_2year,amt_over2year,amt_total");
            	  }
               }
            }else{
               if(acc_div.equals("01") || acc_div.equals("02")) Report_NO="A01";
               else if(acc_div.equals("04"))                    Report_NO="A02";
    		   else if(acc_div.equals("05"))                    Report_NO="A03";
    		   else if(acc_div.equals("08"))                    Report_NO="A06";
    		   else if(acc_div.equals("99"))                    Report_NO="A99";//96.11.16 fix by 2295
    		   if(acc_div.equals("12") || acc_div.equals("13")){
    		      Report_NO="A13";    		     
    		   }
               if(Report_NO.equals("A01") || Report_NO.equals("A06")){//漁會
                  //102.04.18 add 103/01以後,A01漁會套用新表格(增加/異動科目代號)   
                  //103.02.11 add 103/01以後,A06漁會套用新表格(增加/異動科目代號)                 
      		      if(Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH) >= 10301){
      		         ncacno_7 = "ncacno_7_rule";
      		      }
               }
    		   
    		   System.out.println("sqlcmd="+"select * from "+Report_NO+" A01_A05  LEFT JOIN "+ncacno_7+" ON A01_A05.acc_code = "+ncacno_7+".acc_code where A01_A05.m_year="+S_YEAR+" and A01_A05.m_month="+S_MONTH+" and A01_A05.bank_code='"+bank_code+"' and "+ncacno_7+".acc_div='"+acc_div+"' order by "+ncacno_7+".acc_range");
    		   sqlCmd.delete(0,sqlCmd.length());
    		   paramList = new ArrayList();
               sqlCmd.append("select * from "+Report_NO+" A01_A05  LEFT JOIN "+ncacno_7+" ON A01_A05.acc_code = "+ncacno_7+".acc_code where A01_A05.m_year=? and A01_A05.m_month=? and A01_A05.bank_code=? and "+ncacno_7+".acc_div=? order by "+ncacno_7+".acc_range");
               paramList.add(S_YEAR);
               paramList.add(S_MONTH);
               paramList.add(bank_code);
               paramList.add(acc_div);     
               dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt,amt_3month,amt_6month,amt_1year,amt_2year,amt_over2year,amt_total");
            }
            return dbData;
    }
    //96.07.10 add get A08 by 2295
    private List getData_A08(String S_YEAR,String S_MONTH,String bank_code){
    		//查詢條件
    		List dbData =null;
    		StringBuffer sqlCmd = new StringBuffer();
			List paramList = new ArrayList();
			sqlCmd.append("select * from a08 where m_year=? and m_month=? and bank_code=?");
			paramList.add(S_YEAR);
			paramList.add(S_MONTH);
			paramList.add(bank_code);
    		dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,warnaccount_cnt,limitaccount_cnt,erroraccount_cnt,otheraccount_cnt,depositaccount_tcnt");
            return dbData;
    }
    //97.01.02 add get A09 by 2295
    private List getData_A09(String S_YEAR,String S_MONTH,String bank_code){
    		//查詢條件
    		List dbData =null;
    		StringBuffer sqlCmd = new StringBuffer();
			List paramList = new ArrayList();
			sqlCmd.append("select * from a09 where m_year=? and m_month=? and bank_code=?");
			paramList.add(S_YEAR);
			paramList.add(S_MONTH);
			paramList.add(bank_code);
			
    		dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,over_cnt,over_amt,push_over_amt,totalamt,push_totalamt,over_total_rate");
            return dbData;
    }
    //97.06.13 add get A10 by 2295 
    //103.02.26 fix by 2968
    private List getData_A10(String S_YEAR,String S_MONTH,String bank_code){
    		//查詢條件
    		List dbData =null;
    		StringBuffer sqlCmd = new StringBuffer();
			List paramList = new ArrayList();
			sqlCmd.append("select * from a10 where m_year=? and m_month=? and bank_code=?");
			paramList.add(S_YEAR);
			paramList.add(S_MONTH);
			paramList.add(bank_code);
    		dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,loan1_amt,loan2_amt,loan3_amt,loan4_amt,invest1_amt,invest2_amt,invest3_amt,invest4_amt,other1_amt,other2_amt,other3_amt,other4_amt"
														    				+",loan1_baddebt,loan2_baddebt,loan3_baddebt,loan4_baddebt,build1_baddebt,build2_baddebt,build3_baddebt,build4_baddebt"
																			+",baddebt_flag,baddebt_noenough,baddebt_delay,baddebt_104,baddebt_105,baddebt_106,baddebt_107,baddebt_108");
    		return dbData;
    }
    
    //104.01.09 add get A12 by 2968
    private List getData_A12(String S_YEAR,String S_MONTH,String bank_code){
    		//查詢條件
    		List dbData =null;
    		StringBuffer sqlCmd = new StringBuffer();
			List paramList = new ArrayList();
			sqlCmd.append("select * from a12 where m_year=? and m_month=? and bank_code=?");
			paramList.add(S_YEAR);
			paramList.add(S_MONTH);
			paramList.add(bank_code);
    		dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,baddebt_amt,loss_amt,profit_amt");
            return dbData;
    }
  
    //Method modify by jei 93.12.14
	private List getData_B01(String S_YEAR,String S_MONTH,String bank_code){
   		//查詢條件
   	    List dbData =null;
    	StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		System.out.println("This Mothod is \"getData_B01\"....Begin ");
		sqlCmd.append(" select * ");
		sqlCmd.append(" from   B01 a,b00_fund_item b ");
		sqlCmd.append(" where  a.fund_master_no = b.fund_master_no ");
		sqlCmd.append(" and    a.fund_sub_no = b.fund_sub_no ");
		sqlCmd.append(" and    a.fund_next_no = b.fund_next_no ");
		sqlCmd.append(" and    m_year = ?");
		sqlCmd.append(" and    m_month = ?");
		sqlCmd.append(" order by input_order ");
		paramList.add(S_YEAR);
		paramList.add(S_MONTH);
		
		dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,fund_master_no,fund_sub_no,fund_next_no,budget_amt,credit_pay_amt,credit_pay_rate,remark");

		System.out.println("This Mothod is \"getData_B01\"....End ");
            return dbData;
    }

    //Method modify by jei 93.12.16
	private List getData_B02(String S_YEAR,String S_MONTH,String bank_code){
   		//查詢條件
   	    List dbData =null;
    	StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		System.out.println("This Mothod is \"getData_B02\"....Begin ");
		sqlCmd.append(" select * ");
		sqlCmd.append(" from   B02 a,b00_run_item b ");
		sqlCmd.append(" where  a.run_master_no = b.run_master_no ");
		sqlCmd.append(" and    a.run_sub_no = b.run_sub_no ");
		sqlCmd.append(" and    a.run_next_no = b.run_next_no ");
		sqlCmd.append(" and    m_year = ?");
		sqlCmd.append(" and    m_month = ?");
		sqlCmd.append(" order by input_order ");
		paramList.add(S_YEAR);
		paramList.add(S_MONTH);
		dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,run_master_no,run_sub_no,run_next_no,loan_cnt_year,loan_amt_year,loan_cnt_totacc,loan_amt_totacc,loan_cnt_bal,loan_amt_bal_subtot,loan_amt_bal_fund,loan_amt_bal_bank");

		System.out.println("This Mothod is \"getData_B02\"....End ");
        return dbData;
    }

    //Method modify by egg 93.12.10
	private List getData_M01(String S_YEAR,String S_MONTH,String bank_code){
   			//查詢條件
   		    List dbData =null;
    		StringBuffer sqlCmd = new StringBuffer();
		    List paramList = new ArrayList();
			sqlCmd.append(" select 	c.guarantee_item_no,a.guarantee_item_name,c.data_range,b.data_range_name, ");
			sqlCmd.append("        	c.guarantee_cnt,c.loan_amt,c.guarantee_amt,c.loan_bal,c.guarantee_bal,c.over_notpush_cnt, ");
			sqlCmd.append("        	c.over_notpush_bal,c.over_okpush_cnt,c.over_okpush_bal,c.over_notpush_bal, ");
			sqlCmd.append("        	c.repay_tot_cnt,c.repay_tot_amt,c.repay_bal_cnt,c.repay_bal_amt ");
			sqlCmd.append(" from 	m00_guarantee_item a,m00_data_range_item b,m01 c ");
			sqlCmd.append(" where 	c.guarantee_item_no         = a.guarantee_item_no ");
			sqlCmd.append(" and 	substr(c.data_range,4,2)= b.data_range ");
			sqlCmd.append(" and 	b.report_no  = 'M01' ");
			sqlCmd.append(" and 	c.m_year     = ?");
			sqlCmd.append(" and 	c.m_month    = ?");
			sqlCmd.append(" order by a.input_order,b.input_order ");
			paramList.add(S_YEAR);
			paramList.add(S_MONTH);
			dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"guarantee_cnt,loan_amt,guarantee_amt,loan_bal,guarantee_bal,over_notpush_cnt,over_notpush_bal,over_okpush_cnt,over_okpush_bal,repay_tot_cnt,repay_tot_amt,repay_bal_cnt,repay_bal_amt");
        return dbData;
	}

	//jei 931210
	private List getData_M02(String S_YEAR,String S_MONTH,String bank_code){
		//查詢條件
		List dbData =null;
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		sqlCmd.append(" select c.loan_unit_no,a.loan_unit_name,c.data_range,b.data_range_name, ");
		sqlCmd.append("        c.guarantee_cnt,c.loan_amt,c.guarantee_amt,c.loan_bal,c.guarantee_bal,c.over_notpush_cnt, ");
		sqlCmd.append("        c.over_notpush_bal,c.over_okpush_cnt,c.over_okpush_bal,c.repay_tot_cnt, ");
		sqlCmd.append("        c.repay_tot_amt,c.repay_bal_cnt,c.repay_bal_amt ");
		sqlCmd.append(" from   m00_loan_unit a,m00_data_range_item b,M02 c ");
		sqlCmd.append(" where  c.loan_unit_no=a.loan_unit_no ");
		sqlCmd.append(" and substr(c.data_range,4,2)=b.data_range ");
		sqlCmd.append(" and b.report_no = 'M02' ");
		sqlCmd.append(" and c.m_year = ?");
		sqlCmd.append(" and c.m_month = ?");
		paramList.add(S_YEAR);
		paramList.add(S_MONTH);
		sqlCmd.append("order by a.input_order,b.input_order");
		dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,guarantee_cnt,loan_amt,guarantee_amt,loan_bal,guarantee_bal,over_notpush_cnt,over_notpush_bal,over_okpush_cnt,over_okpush_bal,repay_tot_cnt,repay_tot_amt,repay_bal_cnt,repay_bal_amt");

		return dbData;
	}

	//Method modify by egg 93.12.12
	private List getData_M03(String S_YEAR,String S_MONTH,String bank_code,String data_range_type){
		System.out.println("getData_M03()--Begin");
		System.out.println("S_YEAR="+S_YEAR+"S_MONTH="+S_MONTH);
    	//查詢條件
    	List dbData =null;
    	StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
    	if(data_range_type.equals("C")){	// M03的資料
			sqlCmd.append(" select 	M03.div_no,a.data_range,a.data_range_name,M03.guarantee_cnt_month, ");
			sqlCmd.append("        	M03.loan_amt_month, M03.guarantee_amt_month, ");
			sqlCmd.append("        	M03.guarantee_cnt_year,M03.loan_amt_year,M03.guarantee_amt_year, ");
			sqlCmd.append("        	M03.guarantee_bal_totacc,M03.guarantee_bal_totacc_over,M03.repay_bal_totacc ");
			sqlCmd.append(" from 	M03,m00_data_range_item a");
			sqlCmd.append(" where 	M03.div_no=a.data_range ");
			sqlCmd.append(" and		a.data_range_type = 'C' ");
			sqlCmd.append(" and     a.report_no ='M03' ");
			sqlCmd.append(" and 	M03.m_year=? and M03.m_month=?");
			sqlCmd.append(" order 	by a.input_order ");
			paramList.add(S_YEAR);
			paramList.add(S_MONTH);
			dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,guarantee_cnt_month,loan_amt_month,guarantee_amt_month,guarantee_cnt_year,loan_amt_year,guarantee_amt_year,guarantee_bal_totacc,guarantee_bal_totacc_over,repay_bal_totacc");
		}else if(data_range_type.equals("S")){	// M03_NOTE的資料
			System.out.println("S_MONTH in M03 = "+S_MONTH);
			sqlCmd.append("select m_year,m_month,note_no as data_range,note_amt_rate ");
			sqlCmd.append("from M03_note,m00_data_range_item ");
			sqlCmd.append("where note_no=data_range ");
			sqlCmd.append("  and data_range_type='S' ");
			sqlCmd.append("  and m_year=?");
			sqlCmd.append("  and m_month=?");
			sqlCmd.append("order by input_order ");
			paramList.add(S_YEAR);
			paramList.add(S_MONTH);         
			dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,note_amt_rate");
		}
		System.out.println("getData_M03()--End");
		return dbData;
	}

	//jei 931212
	private List getData_M04(String S_YEAR,String S_MONTH,String bank_code){
		//查詢條件
		List dbData =null;
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		System.out.println("This Mothod is \"getData_M04\"....Begin ");
		sqlCmd.append(" select b.loan_use_no,a.loan_use_name,guarantee_no_month,guarantee_no_month_p, ");
		sqlCmd.append("        loan_amt_month,loan_amt_month_p,guarantee_amt_month,guarantee_amt_month_p,guarantee_no_year, ");
		sqlCmd.append("        guarantee_no_year_p,loan_amt_year,loan_amt_year_p,guarantee_amt_year, ");
		sqlCmd.append("        guarantee_amt_year_p,guarantee_no_totacc,guarantee_no_totacc_p, ");
		sqlCmd.append("        loan_amt_totacc,loan_amt_totacc_p,guarantee_amt_totacc,guarantee_amt_totacc_p ");
		sqlCmd.append(" from   m00_loan_use a,M04 b ");
		sqlCmd.append(" where  b.loan_use_no = a.loan_use_no ");
		sqlCmd.append("        and b.m_year = ?");
		sqlCmd.append("        and b.m_month = ?");
		sqlCmd.append("order by a.input_order");
		paramList.add(S_YEAR);      
		paramList.add(S_MONTH);     

		dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,guarantee_no_month,guarantee_no_month_p,loan_amt_month,loan_amt_month_p,guarantee_amt_month,guarantee_amt_month_p,guarantee_no_year,guarantee_no_year_p,loan_amt_year,loan_amt_year_p,guarantee_amt_year,guarantee_amt_year_p,guarantee_no_totacc,guarantee_no_totacc_p,loan_amt_totacc,loan_amt_totacc_p,guarantee_amt_totacc,guarantee_amt_totacc_p");

		System.out.println("This Mothod is \"getData_M04\"....End ");
		return dbData;
	}

    //============add by 2354 12.22  ---
	private List getData_M05(String S_YEAR,String S_MONTH,String bank_code,String data_range_type){
    		//查詢條件
    		List dbData =null;
    		StringBuffer sqlCmd = new StringBuffer();
			List paramList = new ArrayList();
			paramList.add(S_YEAR);      
			paramList.add(S_MONTH); 
    		if(data_range_type.equals("C")){	// M05的資料
				sqlCmd.append(" select m05.loan_unit_no,m00_loan_unit.loan_unit_name, ");
				sqlCmd.append("        m00_data_range_item.data_range,m00_data_range_item.data_range_name, ");
				sqlCmd.append("        m05.period_no,m05.item_no, ");
				sqlCmd.append("        repay_cnt,repay_amt,run_notgood_cnt,run_notgood_amt,turn_out_cnt, ");
				sqlCmd.append("        turn_out_amt,diease_cnt,dieaserepay_amt,disaster_cnt,disaster_amt, ");
				sqlCmd.append("        corun_out_cnt,corun_out_amt,other_cnt,other_amt ");
				sqlCmd.append(" from m05,m00_loan_unit,m00_data_range_item ");
				sqlCmd.append(" where m05.loan_unit_no=m00_loan_unit.loan_unit_no ");
				sqlCmd.append("   and m05.period_no||m05.item_no=m00_data_range_item.data_range ");
				sqlCmd.append("   and m00_data_range_item.data_range_type='C' ");
				sqlCmd.append("   and m05.period_no || m05.item_no = m00_data_rage_item.data_range ");
				sqlCmd.append("   and m05.m_year=? and m05.m_month=?");
				sqlCmd.append(" order by m00_loan_unit.input_order ,m00_data_range_item.input_order ");
				dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,repay_cnt,repay_amt,run_notgood_cnt,run_notgood_amt,turn_out_cnt,turn_out_amt,diease_cnt,dieaserepay_amt,disaster_cnt,disaster_amt,corun_out_cnt,corun_out_amt,other_cnt,other_amt");
			}else if(data_range_type.equals("N")){	// M05_NOTE的資料
				sqlCmd.append("select m_year,m_month,note_no,note_amt_rate ");
				sqlCmd.append("from m05_note,m00_data_range_item ");
				sqlCmd.append("where note_no=data_range ");
				sqlCmd.append("  and data_range_type='N' ");
				sqlCmd.append("  and m_year=?");
				sqlCmd.append("  and m_month=?");
				sqlCmd.append("order  by input_order ");
				dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,note_amt_rate");
			}else{ 	//M05_totacc
				sqlCmd.append("select m_year,m_month,m05_totacc.loan_unit_no,loan_unit_name, ");
				sqlCmd.append("       fix_no,guarantee_no_totacc,guarantee_amt_totacc ");
				sqlCmd.append("from m05_totacc,m00_loan_unit ");
				sqlCmd.append("where m05_totacc.loan_unit_no=m00_loan_unit.loan_unit_no ");
				sqlCmd.append("  and m_year=?");
				sqlCmd.append("  and m_month=?");
				sqlCmd.append("order by input_order ");
				dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,guarantee_no_totacc,guarantee_amt_totacc");
			}
            return dbData;
	}

	//Method modify by egg 93.12.14
	private List getData_M06_M07(String S_YEAR,String S_MONTH,String bank_code,String Report_no){
		List dbData =null;
    	StringBuffer sqlCmd = new StringBuffer(); 
		List paramList = new ArrayList(); 
		paramList.add(S_YEAR);      
		paramList.add(S_MONTH); 
    	// 取得M06(M07)的資料
		sqlCmd.append(" select * from " + Report_no + ",m00_area where " + Report_no + ".area_no=m00_area.area_no and m_year=? and m_month=?");
		sqlCmd.append(" order by input_order ");
		dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,guarantee_no_month,guarantee_amt_month,loan_amt_month,guarantee_no_year,guarantee_amt_year,loan_amt_year,guarantee_no_totacc,guarantee_amt_totacc,loan_amt_totacc,guarantee_bal_no,guarantee_bal_amt,guarantee_bal_p,loan_bal");
		
		return dbData;
	}

	//Method modify by egg 93.12.16
	private List getData_M08(String S_YEAR,String S_MONTH,String bank_code,String Report_no){
		List dbData =null;
    	StringBuffer sqlCmd = new StringBuffer(); 
		List paramList = new ArrayList(); 
		paramList.add(S_YEAR);      
		paramList.add(S_MONTH); 
    	// 取得M08的資料
		sqlCmd.append(" select 	c.id_no,a.id_name,c.data_range,b.data_range_name,c.guarantee_no_month,c.loan_amt_month,");
		sqlCmd.append(" 			c.guarantee_amt_month,c.guarantee_bal_month,c.guarantee_bal_p");
		sqlCmd.append(" from 		m00_id_item a,m00_data_range_item b,M08 c");
		sqlCmd.append(" where 	c.id_no=a.id_no");
		sqlCmd.append(" and 		substr(c.data_range,4,2)=b.data_range");
		sqlCmd.append(" and 		b.report_no='M08'");
		sqlCmd.append(" and 		c.m_year=?");
		sqlCmd.append(" and 		c.m_month=?");
		sqlCmd.append(" order by a.input_order,b.input_order");
		dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,guarantee_no_month,loan_amt_month,guarantee_amt_month,guarantee_bal_month,guarantee_bal_p");		
		return dbData;
	}

	//Method modify by egg 93.12.15
	private List getData_B03(String S_YEAR,String S_MONTH,String bank_code,String Report_no,int form_no){
		List dbData =null;
    	StringBuffer sqlCmd = new StringBuffer(); 
		List paramList = new ArrayList(); 
		paramList.add(S_YEAR);      
		paramList.add(S_MONTH); 

    	// 取得B03的資料
    	if(form_no == 1){	// B03_1的資料
    		System.out.println("form_no=1");
    		sqlCmd.append(" select * from " + Report_no + "_1 a,b00_funs_item b where a.funs_master_no=b.funs_master_no ");
    		sqlCmd.append(" and   a.funs_sub_no=b.funs_sub_no and   a.funs_next_no=b.funs_next_no ");
    		sqlCmd.append(" and   m_year=? and   m_month=? order by input_order");
			dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,loan_cnt_totacc,loan_amt_totacc_fund,loan_amt_totacc_bank,loan_amt_totacc_tot,loan_cnt_bal,loan_amt_bal_fund,loan_amt_bal_bank,loan_amt_bal_tot");
		}else if(form_no == 2){	// B03_2的資料
			System.out.println("form_no=2");
			sqlCmd.append(" select * from " + Report_no + "_2 a,b00_funs_item b where a.funs_master_no=b.funs_master_no ");
    		sqlCmd.append(" and   a.funs_sub_no=b.funs_sub_no and   a.funs_next_no=b.funs_next_no ");
    		sqlCmd.append(" and   m_year=? and   m_month=? order by input_order");
			dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,loan_amt_bal,loan_amt_over,loan_rate_over");
		}else if(form_no == 3){	// B03_3的資料
			System.out.println("form_no=3");
			sqlCmd.append(" select * from " + Report_no + "_3 a,b00_funo_item b where a.funo_master_no=b.funo_master_no ");
    		sqlCmd.append(" and   a.funo_sub_no=b.funo_sub_no and   a.funo_next_no=b.funo_next_no ");
    		sqlCmd.append(" and   m_year=? and  m_month=? order by input_order");
			dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,funo_master_no,funo_sub_no,funo_next_no,funo_amt,funo_rate");
		}else if(form_no == 4){	// B03_4的資料
			System.out.println("form_no=4");
			sqlCmd.append(" select * from " + Report_no + "_4 a,b00_bank_no b where a.bank_no=b.bank_no ");
    		sqlCmd.append(" and   m_year=? and  m_month=? order by input_order");
			dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,machine_cnt,machine_amt,land_cnt,land_amt,house_cnt,house_amt,build_cnt,build_amt,tot_cnt,tot_amt");
		}

		
		return dbData;
	}
	//95.01.16 add F01查詢
	private List getData_F01(String S_YEAR,String S_MONTH,String bank_code){
   		//查詢條件
   	    List dbData =null;
    	StringBuffer sqlCmd = new StringBuffer(); 
		List paramList = new ArrayList(); 
		
		sqlCmd.append(" select dep_type,acct_type,acct_cnt_tm,bal_lm,dep_tm,wtd_tm,bal_tm");
		sqlCmd.append(" from   F01 ");
		sqlCmd.append(" where  m_year = ?");
		sqlCmd.append(" and    m_month = ?");
		sqlCmd.append(" and    bank_code =?");
		sqlCmd.append(" order by dep_type,acct_type ");
		paramList.add(S_YEAR);      
		paramList.add(S_MONTH); 
		paramList.add(bank_code); 
		dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"acct_cnt_tm,bal_lm,dep_tm,wtd_tm,bal_tm");
        return dbData;
   }
%>    