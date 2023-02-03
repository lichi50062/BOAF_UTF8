<%
// 93.12.17 add 權限檢核
// 93.12.20 fix 若有已點選的bank_type,則以已點選的bank_type為主 by 2295
// 93.12.23 add 超過登入時間,請重新登入 by 2295
// 94.01.05 fix 不是super user只能做屬於自己機構代號的人工檢核 by 2295     
// 94.01.05 fix 沒有Bank_List,把所點選的Bank_no清除 by 2295
// 94.06.21 fix 若該目錄下無資料時,顯示無上傳檔案以供檢核 by 2295
// 95.04.14 add A04檢核 by 2295
// 95.05.26 add A99人工檢核 by 2295
// 96.07.11 add A08檢核 by 2295
// 97.01.02 add A09檢核 by 2295
// 97.06.13 add A10檢核 by 2295
// 99.09.23 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
//104.10.12 add A13檢核 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DownLoad" %>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.UpdateA01" %>
<%@ page import="com.tradevan.util.UpdateA02" %>	
<%@ page import="com.tradevan.util.UpdateA03" %>	
<%@ page import="com.tradevan.util.UpdateA04" %>	
<%@ page import="com.tradevan.util.UpdateA05" %>	
<%@ page import="com.tradevan.util.UpdateA06" %>	
<%@ page import="com.tradevan.util.UpdateA08" %>
<%@ page import="com.tradevan.util.UpdateA09" %>
<%@ page import="com.tradevan.util.UpdateA10" %>
<%@ page import="com.tradevan.util.UpdateA13" %>
<%@ page import="com.tradevan.util.UpdateA99" %>	
<%@ page import="com.tradevan.util.UpdateB01" %>	
<%@ page import="com.tradevan.util.UpdateB03" %>	
<%@ page import="com.tradevan.util.UpdateM01" %>	
<%@ page import="com.tradevan.util.UpdateM02" %>	
<%@ page import="com.tradevan.util.UpdateM03" %>
<%@ page import="com.tradevan.util.UpdateM04" %>	
<%@ page import="com.tradevan.util.UpdateM05" %>
<%@ page import="com.tradevan.util.UpdateM06" %>	
<%@ page import="com.tradevan.util.UpdateM07" %>	
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%@ page import="java.io.File" %>
<%@ page import="java.lang.Integer" %>
<%@include file="./include/Header.include" %>

<%
    String S_YEAR = Utility.getTrimString(dataMap.get("S_YEAR"));
	String S_MONTH = Utility.getTrimString(dataMap.get("S_MONTH"));
	String YM = Utility.getTrimString(dataMap.get("YM"));
	String Report_no = Utility.getTrimString(dataMap.get("Report_no"));
	String bank_code = Utility.getTrimString(dataMap.get("bank_code"));	
	String WMdataDir = Utility.getProperties("WMdataDir")+System.getProperty("file.separator")+Report_no;	
	//=======================================================================================================================
	//fix 93.12.20 若有已點選的bank_type,則以已點選的bank_type為主============================================================
	String bank_type = Utility.getTrimString(dataMap.get("bank_type"));	
	//=======================================================================================================================	
	//登入者資訊	
	String tbank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");		
	session.setAttribute("nowtbank_no",null);//94.01.05 fix 沒有Bank_List,把所點選的Bank_no清除======
	System.out.println("act="+act);
	System.out.println("Report_no="+Report_no);
	
	StringBuffer sqlCmd = new StringBuffer();
	List paramList = new ArrayList();
	//99.09.23 add 查詢年度100年以前.縣市別不同===============================
	String year = Utility.getCHTYYMMDD("yy");
    String cd01_table = (Integer.parseInt(year) < 100)?"cd01_99":"";
    String wlx01_m_year = (Integer.parseInt(year) < 100)?"99":"100";
    //=====================================================================
    
		
    if(!Utility.CheckPermission(request,report_no)){//無權限時,導向到LoginError.jsp
        rd = application.getRequestDispatcher( LoginErrorPgName );        
    }else{            
    	//set next jsp 
    	if(act.equals("List")){ 
    	    List dbData = getBank_Type("");
    	    request.setAttribute("bank_type",dbData);    	    
        	rd = application.getRequestDispatcher( ListPgName ); 
    	}else if(act.equals("Query")){    	    	     
    	     // 94.01.05 fix 不是super user只能做屬於自己機構代號的人工檢核 by 2295     
    	     sqlCmd.append("select bank_no, bank_name from bn01 where m_year=? and bank_type=? and bn_type <> '2'");
    	     paramList.add(wlx01_m_year);
    	     paramList.add(bank_type);
    	     if(!lguser_type.equals("S")){//不是super user只能做屬於自己機構代號的人工檢核
    	         sqlCmd.append(" and bank_no =?");
    	         paramList.add(tbank_no);
    	     }
    	     sqlCmd.append(" ORDER BY bank_no");    	     
    	     //===============================================================================
    	     List bn01Data  = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
    	     List bank_typeData = getBank_Type(bank_type);
    	     request.setAttribute("bn01Data",bn01Data);
    	     request.setAttribute("bank_name",(String)((DataObject)bank_typeData.get(0)).getValue("cmuse_name"));
    	     rd = application.getRequestDispatcher( QryPgName +"?bank_type="+bank_type+"&test=nothing");     	
    	}else if(act.equals("Check")){
    	    Date today = new Date();		
    	    int	batch_no = today.hashCode();
    	    System.out.println("batch_no="+batch_no);
    		File tmpFile = new File(WMdataDir);
    		String[] fileList = tmpFile.list();    		
    		//94.06.21 fix 若該目錄下無資料時,顯示無上傳檔案以供檢核
    		if(fileList == null || fileList.length == 0){
    		   actMsg = "無上傳檔案以供檢核";
    		   request.setAttribute("actMsg",actMsg);
          	   rd = application.getRequestDispatcher( nextPgName );    		          	  
    		}else{
    		   System.out.println("fileList.length="+fileList.length);    		   
    		
    		List checkFile = new LinkedList();
    		for(int i=0;i<fileList.length;i++){//把要check的file加到checkFile的list
    		    File checkfile = new File(WMdataDir+System.getProperty("file.separator")+fileList[i]);    		    
    			if (!checkfile.isDirectory()){	    			   
    			     System.out.println("filename="+WMdataDir+System.getProperty("file.separator")+fileList[i]);
    			     if((!Report_no.equals("")) && ((fileList[i].substring(0,3)).equals(Report_no))){
    			          if(bank_code.equals("all")){//所有機構代碼
							 System.out.println("bank_code==all");
							 sqlCmd = new StringBuffer();
							 paramList = new ArrayList();
							 sqlCmd.append("select bank_no, bank_name from bn01 where bank_type=? and bn_type <> '2' ORDER BY bank_no");
							 paramList.add(bank_type);
					         List bn01Data  = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");	 
					            			       					   
					         bank_codeLoop:
					         for(int j=0;j<bn01Data.size();j++){					       
					             if(Report_no.substring(0,1).equals("M") || (fileList[i].substring(3,10)).equals((String)((DataObject)bn01Data.get(j)).getValue("bank_no"))){	//modify by 2354 2004.12.17
					                 if(YM.equals("1")){//年月全部
					                    checkFile.add(fileList[i]);
										System.out.println("add file "+fileList[i]);	
					                    break bank_codeLoop;
    			                     }else{
	    			                    int start1=10;
	    			                    int start2=13;
	    			                    int start3=15;
	    			                    if(Report_no.substring(0,1).equals("M")){
	    			                    	start1=3;
	    			                    	start2=6;
	    			                    	start3=8;
	    			                    }
	    			                    if(((!S_YEAR.equals("")) && (Integer.parseInt(fileList[i].substring(start1,start2)) == Integer.parseInt(S_YEAR)))
	    			    	          	 	&& ((!S_MONTH.equals("")) && ((fileList[i].substring(start2,start3)).equals(S_MONTH)))){
	    			    	            	checkFile.add(fileList[i]);	
	    			    	            	System.out.println("add file "+fileList[i]);				                	 
	    			    	            	break bank_codeLoop;
	    			                   	}
	    			                   	//modify by 2354 12.20 === end
					                 }//end of YM
					              }//end of bank_no
					         }//end of for    			       	
    			    	  }else{//end of bank_code==all begin of bank_code != all
    			    	     if(Report_no.substring(0,1).equals("M") || (fileList[i].substring(3,10)).equals(bank_code)){ //modify by 2354 12.20
					             if(YM.equals("1")){//年月全部
					                 checkFile.add(fileList[i]);					                
					                 System.out.println("allYM");
    			                 }else{
    			                    System.out.println("singleYM");
    			                    System.out.println("S_YEAR="+S_YEAR);
    			                    System.out.println("S_MONTH="+S_MONTH);
	    			                int start1=10;
	    			                int start2=13;
	    			                int start3=15;
	    			                if(Report_no.substring(0,1).equals("M")){
	    			                	start1=3;
	    			                	start2=6;
	    			                	start3=8;
	    			                }
	    			                if(((!S_YEAR.equals("")) && (Integer.parseInt(fileList[i].substring(start1,start2)) == Integer.parseInt(S_YEAR)))
	    			    	         	&& ((!S_MONTH.equals("")) && ((fileList[i].substring(start2,start3)).equals(S_MONTH)))){
	    			    	        	checkFile.add(fileList[i]);	
	    			    	        	System.out.println("add file "+fileList[i]);				                	 
	    			                }
					             }//end of YM
					         }//end of bank_no 
    			          }//enf of bank_code     			     
    			     }//enf of Report_no     
    			}//end of is file
    		}//end of for	
    		
    		for(int i=0;i<checkFile.size();i++){	        			    
    		    Report_no=((String)checkFile.get(i)).substring(0,3);
    		    if(Report_no.substring(0,1).equals("M")){	// add by 2354 MXX報表檔名: M01YYYMM
    		    	bank_code="9700002";
    		    	S_YEAR=((String)checkFile.get(i)).substring(3,6);
    		    	S_MONTH=((String)checkFile.get(i)).substring(6,8);
    		    }else{
    		    	bank_code=((String)checkFile.get(i)).substring(3,10);
    		    	S_YEAR=((String)checkFile.get(i)).substring(10,13);
    		    	S_MONTH=((String)checkFile.get(i)).substring(13,15);
    		    }

    		    System.out.println("Report_no="+Report_no);
    		    System.out.println("bank_code="+bank_code);
    		    System.out.println("S_YEAR="+S_YEAR);
    		    System.out.println("S_MONTH="+S_MONTH); 	   		    
    			System.out.println("begin check i="+i);
    			try{    					
    				//moidfy  and add by winnin 2004.11.17
    				if(Report_no.equals("A01")){
    					actMsg = UpdateA01.doParserReport_A01(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);    						    				    						     
    					//doParserReport_A01(String Report_no, String m_year,String m_month,String filename, String srcbank_code,String upd_method,String input_method,String bank_type)
    				}else if(Report_no.equals("A02")){
	    				actMsg = UpdateA02.doParserReport_A02(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);
    				}else if(Report_no.equals("A03")){
	    				actMsg = UpdateA03.doParserReport_A03(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);
    				}else if(Report_no.equals("A04")){
	    				actMsg = UpdateA04.doParserReport_A04(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);
    				}else if(Report_no.equals("A05")){
	    				actMsg = UpdateA05.doParserReport_A05(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);
	    			}else if(Report_no.equals("A06")){
	    				actMsg = UpdateA06.doParserReport_A06(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);	
	    			}else if(Report_no.equals("A08")){
	    				actMsg = UpdateA08.doParserReport_A08(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);    					
	    			}else if(Report_no.equals("A09")){
	    				actMsg = UpdateA09.doParserReport_A09(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);    						    				
	    			}else if(Report_no.equals("A10")){
	    				actMsg = UpdateA10.doParserReport_A10(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);    						    					    				
	    			}else if(Report_no.equals("A13")){
    					actMsg = UpdateA13.doParserReport_A13(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);    						    				    						         				
	    			}else if(Report_no.equals("A99")){
	    				actMsg = UpdateA99.doParserReport_A99(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);		
    				}else if(Report_no.equals("B01")){
	    				actMsg = UpdateB01.doParserReport_B01(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);
    				}else if(Report_no.equals("B03")){
	    				actMsg = UpdateB03.doParserReport_B03(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);
    				}else if(Report_no.equals("M01")){
	    				actMsg = UpdateM01.doParserReport_M01(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);
    				}else if(Report_no.equals("M02")){
	    				actMsg = UpdateM02.doParserReport_M02(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);
    				}else if(Report_no.equals("M03")){
	    				actMsg = UpdateM03.doParserReport_M03(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);
    				}else if(Report_no.equals("M04")){
	    				actMsg = UpdateM04.doParserReport_M04(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);
	    			}else if(Report_no.equals("M05")){
	    				System.out.println("@@ WMFileCheck.jsp Start call UpdateM05.doParserReport_M05");
	    				actMsg = UpdateM05.doParserReport_M05(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);
	    			}else if(Report_no.equals("M06")){
	    				actMsg = UpdateM06.doParserReport_M06(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);
	    			}else if(Report_no.equals("M07")){
	    				actMsg = UpdateM07.doParserReport_M07(Report_no,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),((String)checkFile.get(i)),bank_code,"M","F",bank_type,"","",batch_no);
	    	
	    			}
	    				
    			}catch(Exception e){
    			    System.out.println(e.getMessage());
    			}	
    			System.out.println("end check i="+i);							
    		}//end of for	   		
    		if(checkFile == null || checkFile.size() == 0){
    		   actMsg = "無符合條件之上傳檔案以供檢核";
    		   request.setAttribute("actMsg",actMsg);
          	   rd = application.getRequestDispatcher( nextPgName );
    		}
    		
    		if(!actMsg.equals("")){//執行檢核失敗
    		   request.setAttribute("actMsg",actMsg);
          	   rd = application.getRequestDispatcher( nextPgName );
          	}else{//執行檢核功
          	   sqlCmd = new StringBuffer();
			   paramList = new ArrayList();
			   sqlCmd.append("SELECT bank_code, bank_name, wml01.m_year || '/' || substr(100 + m_month, 2) as mixdate, input_method, ");
			   sqlCmd.append("upd_code, wml01.m_year, m_month FROM WML01 LEFT JOIN (select * from bn01 where m_year=?)bn01 ON WML01.bank_code = bn01.bank_no ");
    		   sqlCmd.append(" WHERE batch_no=? ORDER BY  mixdate DESC,bank_code");
    		   paramList.add(wlx01_m_year);
    		   paramList.add(batch_no);
          	   List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month"); 	        		
    			request.setAttribute("dbData",dbData);							      		
    		    rd = application.getRequestDispatcher( StatusPgName +"?Report_no="+Report_no+"&test=nothing");            		    		    	  
          	} 
          	
          	}//end of dir have file
    	}//end of check
    	request.setAttribute("actMsg",actMsg);
    }        
     
%>


<%@include file="./include/Tail.include" %>

<%!
    private final static String report_no = "WMFileCheck";
    private final static String nextPgName = "/pages/ActMsg.jsp";
    private final static String QryPgName = "/pages/"+report_no+"_Qry.jsp";
    private final static String StatusPgName = "/pages/"+report_no+"_Status.jsp";
    private final static String ListPgName = "/pages/"+report_no+"_List.jsp";        
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    
    private List getBank_Type(String bank_type){
    		StringBuffer sqlCmd = new StringBuffer();
			List paramList = new ArrayList();
    		//查詢條件    		
    		sqlCmd.append("select cmuse_id,cmuse_name from CDShareNO where cmuse_div='001'");
    		if(!bank_type.equals("")){
    		   sqlCmd.append(" and cmuse_id=?");
    		   paramList.add(bank_type);
    		}
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");            
            return dbData;
    } 
%>    