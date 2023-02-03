<%
// 93.12.07 fix ORDER BY a.m_year desc ,a.m_month desc  by tv2295
// 93.12.17 add 權限檢核 by 2295
// 93.12.18 fix 若有已點選的tbank_no,則以已點選的tbank_no為主 by 2295
// 93.12.20 fix 若有已點選的bank_type,則以已點選的bank_type為主 by 2295
// 93.12.23 add 超過登入時間,請重新登入 by 2295
// 94.02.21 add 加上lock_status from wml01_a_v,若lock_status為'Y'..則不能編輯 by 2295
//          fix 該申報資料沒有被LOCK住才能新增 by 2295
// 94.05.27 add 加上A02區分農漁會 by 2295
// 94.11.14 add 檢核失敗時,prompt資料檢核有誤,詳情請至檢核狀態區查閱 by 2295
// 		   add 上月資料匯入 by 2295
// 94.11.15 add 新增/修改/刪除成功後.可回申報狀態List by 2295
//          add alert Z內容值皆為0 by 2295
// 94.12.25 fix 按新增時.檢查ok而且不是INI者.直接執行"上月資料匯入" by 2295
// 95.01.16 fix 若上月資料皆為0時,可申報下月資料 by 2295
// 95.01.17 fix Insert/Update/Delete 加上bank_type by 2295
// 95.04.11 add A06 Insert/Update/Delete by 2295
// 95.04.19 fix A01~A05改用preparestatment by 2295
// 95.05.11 add 加上A05區分農漁會 by 2295
// 95.05.15 fix 修改a05線上編輯 sql by 2295
// 95.05.17 add A02漁會以ncacno/ncacno_7為主 by 2295
// 95.05.18 add get A99/Inser/Update/Delete by 2295
// 95.07.17 add F01 不能新增 FIX BY 2495
// 96.02.05 fix A06.查詢wlx_apply_ini年月條件 by 2295
// 96.12.03 add A09.基本放款利率計價之舊貸案件資料 by 2295
// 96.12.19 add 97/01以後,套用新表格(增加/異動科目代號) by 2295
// 97.05.07 add A09.佔放款總額的比率(A-B)/(C-D)..去除小數點.取到小數位第2位.四捨5入 * 100存入資料庫 by 2295
// 97.06.13 add A10.應予評估資產彙總資料 by 2295
// 98.04.20 fix (F01/A06/A99/A08/A09/A10)若資料庫中有data時,會直接代出來做update,增加該檔沒有被LOCK住時,才能做update by 2295 //104.01.08 add A12 by 2968
// 99.09.30 合併F01/B01/B02/M01/M02/M04/M06/M07/M08至deleteA01_A10_B01_B02_M01_M08_F01 by 2295
// 99.10.01 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
// 99.10.01 合併A08/A09/A10/A12 的getData_A08_A12 by 2295 //104.01.08 add A12
//100.02.01 fix A10檢查年月條件 by 2295
//102.01.15 add A02增加勾選990421/990621的農金局回文函號 by 2295
//102.04.18 add 103/01以後,漁會A01套用新表格(增加/異動科目代號) by 2295
//102.04.26 add deleteA01_A10_B01_B02_M01_M08_F01-->a02.amt_name by 2295
//102.10.07 add getData_A01_A05,bank_type預設為6農會 by 2295
//103.02.10 add 103/01以後.A06.套用新表格(增加/異動科目代號) by 2295
//104.01.08 add A12 by 2968
//104.02.12 fix 調整A12寫入資料庫回傳訊息 by 2295
//104.02.13 add A02增加勾選990422/990622的農金局回文函號 by 2295
//104.03.17 add A10增加欄位 by 2968
//104.10.12 add A13增加欄位 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.util.*" %>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.UpdateA01" %>
<%@ page import="com.tradevan.util.UpdateA02" %> //add by winnin 2004.12.9
<%@ page import="com.tradevan.util.UpdateA03" %> //add by winnin 2004.12.23
<%@ page import="com.tradevan.util.UpdateA04" %> //add by winnin 2004.12.10
<%@ page import="com.tradevan.util.UpdateA05" %> //add by winnin 2004.12.10
<%@ page import="com.tradevan.util.UpdateA06" %> //add by 2295 2006.04.13
<%@ page import="com.tradevan.util.UpdateB01" %> //add by egg 2004.12.21
<%@ page import="com.tradevan.util.UpdateB02" %> //add by jei 2004.12.21
<%@ page import="com.tradevan.util.UpdateB03" %> //add by egg 2004.12.21
<%@ page import="com.tradevan.util.UpdateM01" %> //add by winnin 2004.12.20
<%@ page import="com.tradevan.util.UpdateM02" %> //add by winnin 2004.12.20
<%@ page import="com.tradevan.util.UpdateM03" %> //add by winnin 2004.12.20
<%@ page import="com.tradevan.util.UpdateM04" %> //add by egg 2004.12.21
<%@ page import="com.tradevan.util.UpdateM05" %> //add by winnin 2004.12.20
<%@ page import="com.tradevan.util.UpdateM06" %> //add by egg 2004.12.21
<%@ page import="com.tradevan.util.UpdateM07" %> //add by egg 2004.12.21
<%@ page import="com.tradevan.util.UpdateM08" %> //add by jei 2004.12.21
<%@ page import="com.tradevan.util.UpdateF01" %> //add by 2295 2005.11.10
<%@ page import="com.tradevan.util.UpdateA99" %> //add by 2295 2006.05.25
<%@ page import="com.tradevan.util.UpdateA08" %> //add by 2295 2007.07.10
<%@ page import="com.tradevan.util.UpdateA09" %> //add by 2295 2007.12.03
<%@ page import="com.tradevan.util.UpdateA10" %> //add by 2295 2008.06.13
<%@ page import="com.tradevan.util.UpdateA12" %> //add by 2968 2015.01.13
<%@ page import="com.tradevan.util.UpdateA13" %> //add by 2295 2015.10.12
<%@include file="./include/Header.include" %>
<%
   	String Report_no = Utility.getTrimString(dataMap.get("Report_no"));
	String S_YEAR = Utility.getTrimString(dataMap.get("S_YEAR"));
	String S_MONTH = Utility.getTrimString(dataMap.get("S_MONTH"));

	//fix 93.12.17 從 session取出登入者資訊=======================================================================================
	String userid = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");
	String username = ( session.getAttribute("muser_name")==null ) ? "" : (String)session.getAttribute("muser_name");
	//======================================================================================================================
	//fix 93.12.18 若有已點選的tbank_no,則以已點選的tbank_no為主============================================================
	String bank_code = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");
	String nowtbank_no = Utility.getTrimString(dataMap.get("tbank_no"));
	if(!nowtbank_no.equals("")){
	   session.setAttribute("nowtbank_no",nowtbank_no);//將已點選的tbank_no寫入session
	}
	bank_code = ( session.getAttribute("nowtbank_no")==null ) ? bank_code : (String)session.getAttribute("nowtbank_no");
	//=======================================================================================================================
	//fix 93.12.20 若有已點選的bank_type,則以已點選的bank_type為主============================================================
	String bank_type = Utility.getTrimString(dataMap.get("bank_type"));
	bank_type = ( session.getAttribute("nowbank_type")==null ) ? bank_type : (String)session.getAttribute("nowbank_type");
	//=======================================================================================================================
	if(session.getAttribute("nowbank_type") == null)
	   System.out.println("nowbank_type == null");
	else  System.out.println((String)session.getAttribute("nowbank_type"));


  	if(!Utility.CheckPermission(request,report_no)){//無權限時,導向到LoginError.jsp
        rd = application.getRequestDispatcher( LoginErrorPgName );
    }else{
    	//set next jsp
    	if(act.equals("List")){
        	rd = application.getRequestDispatcher( ListPgName+"?act=List&test=nothing" );
        }else if(act.equals("Status")){
        	//*94.11.14移至WMFileEdit_Status.jsp-->List dbData = getStatusData(bank_code,Report_no);
            //*94.11.14移至WMFileEdit_Status.jsp-->request.setAttribute("dbData",dbData);
    	    request.setAttribute("actMsg",actMsg);
        	rd = application.getRequestDispatcher( StatusPgName +"?act=Status&Report_no="+Report_no+"&bank_code="+bank_code+"&test=nothing");
        }else if(act.equals("new") || act.equals("getLastMonthData")){
            if(Report_no.equals("F01") || Report_no.equals("A06") || Report_no.equals("A99") || Report_no.equals("A08") || Report_no.equals("A10")){//95.05.26 add A99 //96.07.10 add A08 //97.06.13 add A10 
               if(Report_no.equals("A06") || Report_no.equals("A10")){
				  List datahasLast=getLastMonthData(S_YEAR,S_MONTH,bank_code,Report_no,bank_type);
			      request.setAttribute("datahasLast",datahasLast);
    	       }
			   
			   StringBuffer sqlCmd = new StringBuffer();
			   List paramList = new ArrayList();
			   
               String nextPg = "false";
               sqlCmd.append(" select count(*) as cnt ");
        	   sqlCmd.append(" from WLX_APPLY_INI  aa ");
        	   sqlCmd.append(" where aa.REPORT_NO = ?");
        	   sqlCmd.append(" and (m_year * 100 + m_month) > ? ");//96.02.05 fix 年月條件 by 2295//100.02.01 fix A10檢查年月條件 by 2295
        	   paramList.add(Report_no);
        	   paramList.add(S_YEAR+S_MONTH);
        	   
     	       List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"cnt");
     	       System.out.println("WLX_APPLY_INI.size()="+(((DataObject)dbData.get(0)).getValue("cnt")).toString());
     	       if(dbData != null && ((DataObject)dbData.get(0)).getValue("cnt") != null){
     	          if(Integer.parseInt((((DataObject)dbData.get(0)).getValue("cnt")).toString()) > 0 ){
     	             actMsg = "你鍵入的申報年月超出申報項目起始年月";
     	             request.setAttribute("actMsg",actMsg);
        			 rd = application.getRequestDispatcher( nextPgName );
        			 nextPg="true";
     	          }else{
     	             sqlCmd.delete(0,sqlCmd.length());
     	             paramList = new ArrayList();
     	             sqlCmd.append(" select count(*)  as  cnt ");
        			 sqlCmd.append(" from WML01  aa");
					 sqlCmd.append(" where aa.M_YEAR  = ?");
					 sqlCmd.append("   and aa.M_MONTH = ?");
					 sqlCmd.append("   and aa.BANK_CODE  = ?");
					 sqlCmd.append("   and aa.REPORT_NO = ?");
					 paramList.add(S_YEAR);
					 paramList.add(String.valueOf(Integer.parseInt(S_MONTH)));
					 paramList.add(bank_code);
					 paramList.add(Report_no);
					 
					 dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"cnt");
					 if(Integer.parseInt((((DataObject)dbData.get(0)).getValue("cnt")).toString()) > 0 ){
					    //有已申報過的資料,直接進修改畫面
					    List data_div01=null;
					    if(Report_no.equals("F01")){
					       data_div01 = getData_F01(S_YEAR,S_MONTH,bank_code);
					       if(!checkExistMonth(S_YEAR,S_MONTH,bank_code,"NEXT",Report_no)){//無下月申報的資料
			                  rd = application.getRequestDispatcher( EditPgName +Report_no+".jsp?act=Edit&NextMonthCreated=false&bank_type="+bank_type+"&S_YEAR="+S_YEAR+"&S_MONTH="+String.valueOf(Integer.parseInt(S_MONTH))+"&test=nothing");
			               }else{
			                  rd = application.getRequestDispatcher( EditPgName +Report_no+".jsp?act=Edit&NextMonthCreated=true&bank_type="+bank_type+"&S_YEAR="+S_YEAR+"&S_MONTH="+String.valueOf(Integer.parseInt(S_MONTH))+"&test=nothing");
  			               }
					    }else if(Report_no.equals("A06")){
					       data_div01=getData_A01_A05(S_YEAR,S_MONTH,bank_code,bank_type,"08");
					       rd = application.getRequestDispatcher( EditPgName +Report_no+".jsp?act=Edit&bank_type="+bank_type+"&S_YEAR="+S_YEAR+"&S_MONTH="+String.valueOf(Integer.parseInt(S_MONTH))+"&test=nothing");
					    }else if(Report_no.equals("A99")){
					       data_div01=getData_A01_A05(S_YEAR,S_MONTH,bank_code,bank_type,"99");
					       rd = application.getRequestDispatcher( EditPgName +Report_no+".jsp?act=Edit&bank_type="+bank_type+"&S_YEAR="+S_YEAR+"&S_MONTH="+String.valueOf(Integer.parseInt(S_MONTH))+"&test=nothing");
					    }else if(Report_no.equals("A08") || Report_no.equals("A09") || Report_no.equals("A10") || Report_no.equals("A12")){//97.06.13 add A10 
					       data_div01=getData_A08_A12(S_YEAR,S_MONTH,bank_code,Report_no);
					       String path = EditPgName +Report_no+".jsp?act=Edit&bank_type="+bank_type+"&S_YEAR="+S_YEAR+"&S_MONTH="+String.valueOf(Integer.parseInt(S_MONTH))+"&test=nothing";
					       if(Report_no.equals("A10")) path+="&width=729";
					       rd = application.getRequestDispatcher(path);
					    }
    	    	        request.setAttribute("data_div01",data_div01);
    	    	        request.setAttribute("actMsg",actMsg);

        				nextPg="true";
        				System.out.println("F01/A06/A99/A08/A09/A10/A12 Update");
					 }else{
					    System.out.println("F01/A06/A99/A08/A09/A10/A12 new");
					    String WLX_APPLY_INI[] = getWLX_APPLY_INI(Report_no);
					    String F01_APPLY_INI="false";
					    if(S_YEAR.equals(WLX_APPLY_INI[0]) && String.valueOf(Integer.parseInt(S_MONTH)).equals(WLX_APPLY_INI[1])){
					       F01_APPLY_INI="true";//為農金局設定的起始年月
					       System.out.println("申報年月與農金局設定的起始年月相同");
					    }else{
					       System.out.println("申報年月與農金局設定的起始年月不相同");
					       sqlCmd.delete(0,sqlCmd.length());
     	             	   paramList = new ArrayList();
					       sqlCmd.append(" select count(*)  as  cnt ");
        			       sqlCmd.append(" from WML01  aa");
						   sqlCmd.append(" where aa.M_YEAR  = ?");//起始年
						   sqlCmd.append("   and aa.M_MONTH = ?");//起始月
						   sqlCmd.append("   and aa.BANK_CODE  = ?");
						   sqlCmd.append("   and aa.REPORT_NO = ?");
						   paramList.add(WLX_APPLY_INI[0]);
						   paramList.add(WLX_APPLY_INI[1]);
						   paramList.add(bank_code);
						   paramList.add(Report_no);
					      dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"cnt");
					      if(Integer.parseInt((((DataObject)dbData.get(0)).getValue("cnt")).toString()) == 0 ){
					         alertMsg = "農金局規定本案起始的申報年月為("+WLX_APPLY_INI[0]+","+WLX_APPLY_INI[1]+")"
					         		+ ",與你第一次將申報年月("+S_YEAR+","+S_MONTH+")不一致"
					                + ",如果你是("+WLX_APPLY_INI[0]+","+WLX_APPLY_INI[1]+")之後才新增成立的單位,"
					                + "亦請以成立當月為起始申報年月??";
       					     F01_APPLY_INI="true";//為新增的總機構單位視同為ini
					         webURL_Y = EditPgName +Report_no+".jsp?act=new&bank_type="+bank_type+"&S_YEAR="+S_YEAR+"&S_MONTH="+String.valueOf(Integer.parseInt(S_MONTH))+"&F01_APPLY_INI="+F01_APPLY_INI+"&test=nothing";
					         webURL_N = StatusPgName +"?act=Status&Report_no="+Report_no+"&bank_code="+bank_code+"&test=nothing";
    	    				 request.setAttribute("actMsg",actMsg);
    	       	   	   		 request.setAttribute("alertMsg",alertMsg);
     	             		 request.setAttribute("webURL_Y",webURL_Y);
    				         request.setAttribute("webURL_N",webURL_N);
        			 		 rd = application.getRequestDispatcher( nextPgName );
        			 		 nextPg="true";
					      }else{//表示不為INI新增
					         System.out.println("不為INI新增");
					         //檢查前一個月是否有申報成功的資料
					         if(Report_no.equals("F01")){
					            if(!checkExistMonth(S_YEAR,S_MONTH,bank_code,"LAST",Report_no)){//無申報成功的資料
					               actMsg = "你鍵入的申報年月的前一個月資料尚有錯誤或未鍵檔,請先修正完妥後方可再辦理新增";
     	             		       request.setAttribute("actMsg",actMsg);
        			 		       rd = application.getRequestDispatcher( nextPgName );
        			 		       nextPg="true";
					            }else{//有上個月申報成功的資料
					             //if(act.equals("getLastMonthData")){//94.11.14 上月資料匯入 by 2295
					              List data_div01=getLastMonthData(S_YEAR,S_MONTH,bank_code,Report_no,bank_type);
			                      request.setAttribute("data_div01",data_div01);
    	                          //}
					            }
					        }else if(Report_no.equals("A06")){
					          if(act.equals("getLastMonthData")){//94.11.14 上月資料匯入 by 2295
					              List data_div01=getLastMonthData(S_YEAR,S_MONTH,bank_code,Report_no,bank_type);
			                      request.setAttribute("data_div01",data_div01);
    	                      }
					        }else if(Report_no.equals("A08")){//96.07.10 上月資料匯入 by 2295 
					           List data_div01=getLastMonthData(S_YEAR,S_MONTH,bank_code,Report_no,bank_type);
			                   request.setAttribute("data_div01",data_div01);
					        }else if(Report_no.equals("A10")){
					          //if(act.equals("getLastMonthData")){//97.06.13 上月資料匯入 by 2295
					           	 List data_div01=getLastMonthData(S_YEAR,S_MONTH,bank_code,Report_no,bank_type);
			                   	 request.setAttribute("data_div01",data_div01);
			                   	 rd = application.getRequestDispatcher( EditPgName +Report_no+".jsp?act=new&bank_type="+bank_type+"&S_YEAR="+S_YEAR+"&S_MONTH="+String.valueOf(Integer.parseInt(S_MONTH))+"&F01_APPLY_INI="+F01_APPLY_INI+"&test=nothing&width=729");
			                   	 nextPg = "true";
			                  //}
					        }
					      }
					    }//F01_APPLY_INI="false"
					    request.setAttribute("actMsg",actMsg);
					    if(!nextPg.equals("true")){//沒有設下一頁時,才加預設的next pg by 2295
        				   rd = application.getRequestDispatcher( EditPgName +Report_no+".jsp?act="+act+"&bank_type="+bank_type+"&S_YEAR="+S_YEAR+"&S_MONTH="+String.valueOf(Integer.parseInt(S_MONTH))+"&F01_APPLY_INI="+F01_APPLY_INI+"&test=nothing");
        				}
        				//rd = application.getRequestDispatcher( EditPgName +Report_no+".jsp?act=getLastMonth&NextMonthCreated=false&test=nothing");
					 }//F01/A06 new
     	          }
     	       }
     	    }else{
    	       request.setAttribute("actMsg",actMsg);
        	   rd = application.getRequestDispatcher( EditPgName +Report_no+".jsp?act=new&bank_type="+bank_type+"&test=nothing");
        	}
    	}else if(act.equals("Edit")){
    	    String NextMonthCreated = "true";//94.11.11 add
    	    if(Report_no.equals("A01")){
    	    	List data_div01=getData_A01_A05(S_YEAR,S_MONTH,bank_code,bank_type,"01");
    	    	List data_div02=getData_A01_A05(S_YEAR,S_MONTH,bank_code,bank_type,"02");
    	    	request.setAttribute("data_div01",data_div01);
    	    	request.setAttribute("data_div02",data_div02);
    	    }else if(Report_no.equals("A02")){
    	    	List data_div01=getData_A01_A05(S_YEAR,S_MONTH,bank_code,bank_type,"04");
    	    	request.setAttribute("data_div01",data_div01);
    	    }else if(Report_no.equals("A03")){
    	    	List data_div01=getData_A01_A05(S_YEAR,S_MONTH,bank_code,bank_type,"05");
    	    	request.setAttribute("data_div01",data_div01);
			}else if(Report_no.equals("A04")){
    	    	List data_div01=getData_A01_A05(S_YEAR,S_MONTH,bank_code,bank_type,"06");
    	    	request.setAttribute("data_div01",data_div01);
			}else if(Report_no.equals("A05")){
    	    	List data_div01=getData_A01_A05(S_YEAR,S_MONTH,bank_code,bank_type,"07");
    	    	request.setAttribute("data_div01",data_div01);
    	    }else if(Report_no.equals("A06")){//95.04.10 add by 2295
    	    	List data_div01=getData_A01_A05(S_YEAR,S_MONTH,bank_code,bank_type,"08");
    	    	request.setAttribute("data_div01",data_div01);
    	    }else if(Report_no.equals("A08") || Report_no.equals("A09") || Report_no.equals("A10") || Report_no.equals("A12")){//96.07.10 add by 2295 
    	    	List data_div01=getData_A08_A12(S_YEAR,S_MONTH,bank_code,Report_no);
    	    	request.setAttribute("data_div01",data_div01);    	    
    	    }else if(Report_no.equals("A99")){//95.05.18 add by 2295
    	    	List data_div01=getData_A01_A05(S_YEAR,S_MONTH,bank_code,bank_type,"99");
    	    	request.setAttribute("data_div01",data_div01);
    	    }else if(Report_no.equals("A13")){//104.10.12 add by 2295
    	    	List data_div01=getData_A01_A05(S_YEAR,S_MONTH,bank_code,bank_type,"12");
    	    	List data_div02=getData_A01_A05(S_YEAR,S_MONTH,bank_code,bank_type,"13");
    	    	request.setAttribute("data_div01",data_div01);
    	    	request.setAttribute("data_div02",data_div02);	
			}else if(Report_no.equals("B01")){
    	    	List data_div01=getData_B01(S_YEAR,S_MONTH,bank_code);
    	    	request.setAttribute("data_div01",data_div01);
			}else if(Report_no.equals("B02")){
    	    	List data_div01=getData_B02(S_YEAR,S_MONTH,bank_code);
    	    	request.setAttribute("data_div01",data_div01);
    	    }else if(Report_no.equals("B03")){
				System.out.println("The condition is Report_no.equals(\"B03\") --Begin ");
    	    	List data_div01=getData_B03(S_YEAR,S_MONTH,bank_code,Report_no,1);
    	    	request.setAttribute("data_div01",data_div01);
    	    	List data_div02=getData_B03(S_YEAR,S_MONTH,bank_code,Report_no,2);
    	    	request.setAttribute("data_div02",data_div02);
    	    	List data_div03=getData_B03(S_YEAR,S_MONTH,bank_code,Report_no,3);
    	    	request.setAttribute("data_div03",data_div03);
    	    	List data_div04=getData_B03(S_YEAR,S_MONTH,bank_code,Report_no,4);
    	    	request.setAttribute("data_div04",data_div04);
    	    	System.out.println("The condition is Report_no.equals(\"B03\") --End ");
			}else if(Report_no.equals("M01")){
    	    	List data_div01=getData_M01(S_YEAR,S_MONTH,bank_code);
    	    	request.setAttribute("data_div01",data_div01);
			}else if(Report_no.equals("M02")){
    	    	List data_div01=getData_M02(S_YEAR,S_MONTH,bank_code);
    	    	request.setAttribute("data_div01",data_div01);
			}else if(Report_no.equals("M03")){
    	    	List data_div01=getData_M03(S_YEAR,S_MONTH,bank_code,"C");
    	    	request.setAttribute("data_div01",data_div01);
    	    	List data_div02=getData_M03(S_YEAR,S_MONTH,bank_code,"S");
    	    	request.setAttribute("data_div02",data_div02);
			}else if(Report_no.equals("M04")){
    	    	List data_div01=getData_M04(S_YEAR,S_MONTH,bank_code);
    	    	request.setAttribute("data_div01",data_div01);
			}else if(Report_no.equals("M05")){
    	    	List data_div01=getData_M05(S_YEAR,S_MONTH,bank_code,"C");
    	    	request.setAttribute("data_div01",data_div01);
    	    	List data_div02=getData_M05(S_YEAR,S_MONTH,bank_code,"");
    	    	request.setAttribute("data_div02",data_div02);
    	    	List data_div03=getData_M05(S_YEAR,S_MONTH,bank_code,"N");
    	    	request.setAttribute("data_div03",data_div03);
			}else if(Report_no.equals("M06") || Report_no.equals("M07")){
    	    	List data_div01=getData_M06_M07(S_YEAR,S_MONTH,bank_code,Report_no);
    	    	request.setAttribute("data_div01",data_div01);
			}else if(Report_no.equals("M08")){
    	    	List data_div01=getData_M08(S_YEAR,S_MONTH,bank_code,Report_no);
    	    	request.setAttribute("data_div01",data_div01);
			}else if(Report_no.equals("F01")){
			    if(!checkExistMonth(S_YEAR,S_MONTH,bank_code,"NEXT",Report_no)){//無下月申報的資料
			        NextMonthCreated = "false";
			    }
			    //NextMonthCreated = "false";//94.11.11 for test
			    List data_div01=getData_F01(S_YEAR,S_MONTH,bank_code);
			    request.setAttribute("data_div01",data_div01);
			}
    	    request.setAttribute("actMsg",actMsg);
    	    String path = ".jsp?act=Edit&NextMonthCreated="+NextMonthCreated+"&bank_type="+bank_type+"&test=nothing";
    	    if(Report_no.equals("A10")) path+="&width=729";
    	    rd = application.getRequestDispatcher( EditPgName +Report_no+path);
    	}else if(act.equals("Insert")){
    	    if(!CheckFileLock(S_YEAR,S_MONTH,bank_code,Report_no)){//94.02.21該檔沒有被LOCK住
      	       if(Report_no.equals("A01") || Report_no.equals("A02") || Report_no.equals("A03") || Report_no.equals("A04") || Report_no.equals("A05")){
      	        	actMsg = insertA01_A05(request,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),bank_code,bank_type,Report_no,userid,username);
      	       }else if(Report_no.equals("B01")){
      	        	actMsg = insertB01(request,S_YEAR,S_MONTH,bank_code,Report_no,userid,username);
      	       }else if(Report_no.equals("B02")){
      	        	actMsg = insertB02(request,S_YEAR,S_MONTH,bank_code,Report_no,userid,username);
      	       }else if(Report_no.equals("B03")){
      	        	actMsg = insertB03(request,S_YEAR,S_MONTH,bank_code,Report_no,userid,username);
      	       }else if(Report_no.equals("M01")){
      	        	actMsg = insertM01(request,S_YEAR,S_MONTH,bank_code,Report_no,userid,username);
      	       }else if(Report_no.equals("M02")){
      	        	actMsg = insertM02(request,S_YEAR,S_MONTH,bank_code,Report_no,userid,username);
      	       }else if(Report_no.equals("M03")){
      	        	actMsg = insertM03(request,S_YEAR,S_MONTH,bank_code,Report_no,userid,username);
      	       }else if(Report_no.equals("M04")){
      	        	actMsg = insertM04(request,S_YEAR,S_MONTH,bank_code,Report_no,userid,username);
      	       }else if(Report_no.equals("M05")){
      	        	actMsg = insertM05(request,S_YEAR,S_MONTH,bank_code,Report_no,userid,username);
      	       }else if(Report_no.equals("M06") || Report_no.equals("M07")){
      	        	actMsg = insertM06_M07(request,S_YEAR,S_MONTH,bank_code,Report_no,userid,username);
      	       }else if(Report_no.equals("M08")){
      	        	actMsg = insertM08(request,S_YEAR,S_MONTH,bank_code,Report_no,userid,username);
    	       }else if(Report_no.equals("F01") || Report_no.equals("A06") || Report_no.equals("A99") || Report_no.equals("A08") || Report_no.equals("A09") || Report_no.equals("A10") || Report_no.equals("A12") || Report_no.equals("A13")){
    	            //94.11.10 add by 2295
    	            //95.05.18 add A99 by 2295
    	            if(Report_no.equals("F01")){
      	        	    actMsg = insertF01(request,S_YEAR,S_MONTH,bank_code,Report_no,userid,username);
      	        	}else if(Report_no.equals("A06")){//95.04.14 add A06 by 2295
      	        	    actMsg = insertA06(request,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),bank_code,bank_type,Report_no,userid,username);
      	        	}else if(Report_no.equals("A99")){//95.05.18 add A99 by 2295
      	        	    actMsg = insertA01_A05(request,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),bank_code,bank_type,Report_no,userid,username);
      	        	}if(Report_no.equals("A08")){
      	        	    actMsg = insertA08(request,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),bank_code,bank_type,Report_no,userid,username);
					}if(Report_no.equals("A09")){//96.12.10 add A09 by 2295
      	        	    actMsg = insertA09(request,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),bank_code,bank_type,Report_no,userid,username);
      	        	}if(Report_no.equals("A10")){//97.06.13 add A10 by 2295
      	        	    actMsg = insertA10(request,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),bank_code,bank_type,Report_no,userid,username);
      	        	}if(Report_no.equals("A12")){
      	        	    actMsg = insertA12(request,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),bank_code,bank_type,Report_no,userid,username);
      	        	}else if(Report_no.equals("A13")){//104.10.12 add A13 by 2295
      	        	    actMsg = insertA01_A05(request,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),bank_code,bank_type,Report_no,userid,username);    
      	        	}
      	        	//94.11.14 add by 2295 檢核失敗時,prompt資料檢核有誤,詳情請至檢核狀態區查閱
      	        	String updCode=actMsg.substring(0,1);
      	        	System.out.println("updCode="+updCode);
      	        	actMsg = actMsg.substring(2,actMsg.length());
      	        	if(updCode.equals("E")){//E檢核失敗
      	        	   alertMsg = "『建檔完成』『資料檢核有誤,詳情請至檢核狀態區查閱』";
      	        	}else if(updCode.equals("Z")){//94.11.15 add Z內容值皆為0
      	        	   alertMsg = "『建檔完成』『申報內容值皆為0,詳情請至檢核狀態區查閱』";
      	        	}
      	        	webURL_Y = StatusPgName +"?act=Status&Report_no="+Report_no+"&bank_code="+bank_code+"&test=nothing";
      	        	webURL_N = StatusPgName +"?act=Status&Report_no="+Report_no+"&bank_code="+bank_code+"&test=nothing";
    	    		request.setAttribute("actMsg",actMsg);
    	       	   	request.setAttribute("alertMsg",alertMsg);
     	            request.setAttribute("webURL_Y",webURL_Y);
    				request.setAttribute("webURL_N",webURL_N);
        			rd = application.getRequestDispatcher( nextPgName );
    	       }
    	    }else{//已被鎖住
    	       actMsg = "該基準日申報資料已被鎖住,無法新增";
    	    }
    	    request.setAttribute("actMsg",actMsg);
        	//rd = application.getRequestDispatcher( nextPgName );
    	    //94.11.15 add 新增成功後.可回申報狀態List
    	    //95.01.17 fix 加上bank_type
        	rd = application.getRequestDispatcher( nextPgName +"?goPages=WMFileEdit_Status.jsp&act=Status&Report_no="+Report_no+"&bank_code="+bank_code+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&bank_type="+bank_type);
    	}else if(act.equals("Update")){
    		if(Report_no.equals("A01") || Report_no.equals("A02") || Report_no.equals("A03") || Report_no.equals("A04") || Report_no.equals("A05") ){
    	     	actMsg = updateA01_A05(request,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),bank_code,Report_no,userid,username);
    	    }else if(Report_no.equals("B01")){
    	     	actMsg = updateB01(request,S_YEAR,S_MONTH,bank_code,Report_no,userid,username);
    	    }else if(Report_no.equals("B02")){
    	     	actMsg = updateB02(request,S_YEAR,S_MONTH,bank_code,Report_no,userid,username);
    	    }else if(Report_no.equals("M01")){
    	     	actMsg = updateM01(request,S_YEAR,S_MONTH,bank_code,Report_no,userid,username);
    	    }else if(Report_no.equals("M02")){
    	     	actMsg = updateM02(request,S_YEAR,S_MONTH,bank_code,Report_no,userid,username);
    	    }else if(Report_no.equals("M03")){
    	     	actMsg = updateM03(request,S_YEAR,S_MONTH,bank_code,Report_no,userid,username);
    	    }else if(Report_no.equals("M04")){
    	     	actMsg = updateM04(request,S_YEAR,S_MONTH,bank_code,Report_no,userid,username);
    	    }else if(Report_no.equals("M05")){
    	     	actMsg = updateM05(request,S_YEAR,S_MONTH,bank_code,Report_no,userid,username);
    	    }else if(Report_no.equals("M06") || Report_no.equals("M07")){
    	     	actMsg = updateM06_M07(request,S_YEAR,S_MONTH,bank_code,Report_no,userid,username);
    	    }else if(Report_no.equals("M08")){
    	     	actMsg = updateM08(request,S_YEAR,S_MONTH,bank_code,Report_no,userid,username);
    	    }else if(Report_no.equals("B03")){
    	     	actMsg = updateB03(request,S_YEAR,S_MONTH,bank_code,Report_no,userid,username);
    	    }else if(Report_no.equals("F01") || Report_no.equals("A06") || Report_no.equals("A99") || Report_no.equals("A08") || Report_no.equals("A09") || Report_no.equals("A10") || Report_no.equals("A12") || Report_no.equals("A13")){//94.11.11 add by 2295  //96.12.03 add by 2295 //97.06.13 //104.01.08
    	        if(!CheckFileLock(S_YEAR,S_MONTH,bank_code,Report_no)){//98.04.20 add 該檔沒有被LOCK住
    	           if(Report_no.equals("F01")){
    	     	      actMsg = updateF01(request,S_YEAR,S_MONTH,bank_code,Report_no,userid,username);
    	     	   }else if(Report_no.equals("A06")){ //95.04.14 add A06 by 2295
    	     	      actMsg = updateA06(request,S_YEAR,S_MONTH,bank_code,Report_no,userid,username);
    	     	   }else if(Report_no.equals("A99")){ //95.05.18 add A99 by 2295
    	     	      actMsg = updateA01_A05(request,S_YEAR,S_MONTH,bank_code,Report_no,userid,username);
    	     	   }if(Report_no.equals("A08")){
    	     	      actMsg = updateA08(request,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),bank_code,Report_no,userid,username);
    	     	   }if(Report_no.equals("A09")){
    	     	      actMsg = updateA09(request,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),bank_code,Report_no,userid,username);
    	     	   }if(Report_no.equals("A10")){//97.06.13 add A10 by 2295
    	     	      actMsg = updateA10(request,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),bank_code,Report_no,userid,username);
    	     	   }if(Report_no.equals("A12")){
    	     	      actMsg = updateA12(request,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),bank_code,Report_no,userid,username);
    	     	   }else if(Report_no.equals("A13")){ //104.10.12 add A13 by 2295
    	     	      actMsg = updateA01_A05(request,S_YEAR,S_MONTH,bank_code,Report_no,userid,username);   
    	     	   }
    	     	}else{//已被鎖住
    	          actMsg = "該基該基準日申報資料已被鎖住,無法新增";
    	        }
    	     	//94.11.14 add by 2295 檢核失敗時,prompt資料檢核有誤,詳情請至檢核狀態區查閱
      	        String updCode=actMsg.substring(0,1);
      	        System.out.println("updCode="+updCode);
      	        actMsg = actMsg.substring(2,actMsg.length());
      	        if(updCode.equals("E")){//E檢核失敗
      	           alertMsg = "『修改建檔完成』『資料檢核有誤,詳情請至檢核狀態區查閱』";
      	        }
      	        webURL_Y = StatusPgName +"?act=Status&Report_no="+Report_no+"&bank_code="+bank_code+"&test=nothing";
      	        webURL_N = StatusPgName +"?act=Status&Report_no="+Report_no+"&bank_code="+bank_code+"&test=nothing";
    	    	request.setAttribute("actMsg",actMsg);
    	       	request.setAttribute("alertMsg",alertMsg);
     	        request.setAttribute("webURL_Y",webURL_Y);
    			request.setAttribute("webURL_N",webURL_N);
        		rd = application.getRequestDispatcher( nextPgName );
    	    }
	   	    request.setAttribute("actMsg",actMsg);
        	//rd = application.getRequestDispatcher( nextPgName );
        	//94.11.15 add 修改成功後.可回申報狀態List
        	//95.01.17 fix 加上bank_type
        	rd = application.getRequestDispatcher( nextPgName +"?goPages=WMFileEdit_Status.jsp&act=Status&Report_no="+Report_no+"&bank_code="+bank_code+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&bank_type="+bank_type);
    	}else if(act.equals("Delete")){//99.09.30 合併F01/B01/B02/M01/M02/M04/M06/M07/M08
    	    if(Report_no.equals("A01") || Report_no.equals("A02") || Report_no.equals("A03") || Report_no.equals("A04") || Report_no.equals("A05") || Report_no.equals("A06") || Report_no.equals("A99") || Report_no.equals("A08") || Report_no.equals("A09") || Report_no.equals("A10")  || Report_no.equals("A12") || Report_no.equals("F01") || Report_no.equals("A13")
    	    || Report_no.equals("B01") || Report_no.equals("B02") || Report_no.equals("M01") || Report_no.equals("M02") || Report_no.equals("M04") || Report_no.equals("M06") || Report_no.equals("M07")  || Report_no.equals("M08") )    	        	    
    	    {//95.04.14 add A06 by 2295//95.05.18 add A99 by 2295//96.12.03 add A09 by 2295 //97.06.13 add A10 by 2295
    	     	actMsg = deleteA01_A10_B01_B02_M01_M08_F01(request,String.valueOf(Integer.parseInt(S_YEAR)),String.valueOf(Integer.parseInt(S_MONTH)),bank_code,Report_no,userid,username);
    	    }else if(Report_no.equals("B03")){
    	    	System.out.println("Before call deleteB03 method S_MONTH=[" + S_MONTH +"]");
    	     	actMsg = deleteB03(request,S_YEAR,S_MONTH,bank_code,Report_no,userid,username);
    	    }else if(Report_no.equals("M03")){
    	     	actMsg = deleteM03(request,S_YEAR,S_MONTH,bank_code,Report_no,userid,username);
    	    }else if(Report_no.equals("M05")){
    	     	actMsg = deleteM05(request,S_YEAR,S_MONTH,bank_code,Report_no,userid,username);
    	    }
    	        	    
    	    request.setAttribute("actMsg",actMsg);
    	    //94.11.15 add 刪除成功後.可回申報狀態List
    	    //95.01.17 fix 加上bank_type
        	rd = application.getRequestDispatcher( nextPgName +"?goPages=WMFileEdit_Status.jsp&act=Status&Report_no="+Report_no+"&bank_code="+bank_code+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&bank_type="+bank_type);
    	}
    }
%>

<%@include file="./include/Tail.include" %>

<%!
    private final static String report_no = "WMFileEdit";
    private final static String nextPgName = "/pages/ActMsg.jsp";
    private String EditPgName = "/pages/"+report_no+"_";
    private final static String ListPgName = "/pages/"+report_no+"_List.jsp";
    private final static String StatusPgName = "/pages/"+report_no+"_Status.jsp";
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";

    //95.04.19 改用preparestatement by 2295
    //95.05.18 add insert A99 by 2295
    public String insertA01_A05(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String bank_type,String Report_no,String userid,String username) throws Exception{
    	String[]	Amt			  = request.getParameterValues("amt");
    	String[]	Acc_Code	  = request.getParameterValues("acc_code");
    	String txt990421 = request.getParameter("txt990421");
    	String txt990621 = request.getParameter("txt990621");
    	String txt990422 = request.getParameter("txt990422");//104.02.13
    	String txt990622 = request.getParameter("txt990622");//104.02.13
		//String bank_type = ( request.getAttribute("bank_type")==null ) ? "6" : (String)request.getAttribute("bank_type");
		String sqlCmd = "";
		String errMsg="";
		int zerodata=0;
		System.out.println("Acc_Code.size="+Acc_Code.length);
		//System.out.println("Amt.size="+Amt.length);
		try {
				List updateDBList = new LinkedList();//0:sql 1:data
				List updateDBSqlList = new LinkedList();
				List updateDBDataList = new LinkedList();//儲存參數的List
				List dataList = new LinkedList();//儲存參數的data
				List paramList = new ArrayList();
				sqlCmd = "SELECT * FROM "+Report_no+" WHERE m_year= ? AND m_month= ? AND bank_code= ?";
				paramList.add(S_YEAR);
				paramList.add(S_MONTH);
				paramList.add(bank_code);
			    List data = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"m_year,m_month,amt");
				System.out.println(""+Report_no+".size="+data.size());

				if (data.size() != 0){
				    errMsg = errMsg + "此筆資料已存在無法新增<br>";
				}else{
				    if(Report_no.equals("A05") || Report_no.equals("A02")){//102.01.15 add by 2295
				       updateDBSqlList.add("INSERT INTO "+Report_no+" VALUES(?,?,?,?,?,?)");
				       System.out.println("INSERT INTO "+Report_no+" VALUES(?,?,?,?,?,?)");				   
				    }else{
				       updateDBSqlList.add("INSERT INTO "+Report_no+" VALUES(?,?,?,?,?)");
				       System.out.println("INSERT INTO "+Report_no+" VALUES(?,?,?,?,?)");
				    }
					for (int i = 0; i < Acc_Code.length; i++) {
					     dataList = new LinkedList();
						 dataList.add(S_YEAR);//m_year
						 dataList.add(S_MONTH);//m_month
						 dataList.add(bank_code);//bank_code
						 dataList.add(Acc_Code[i].trim());//acc_code
						 if(Report_no.equals("A01") || Report_no.equals("A99")){
						   dataList.add(Utility.setNoCommaFormat(Amt[i]));//amt
						 }else if(Report_no.equals("A02")){//102.01.15 add by 2295
		            	 	dataList.add(Utility.setNoCommaFormat(Amt[i]));//amt
		            	 	if("990421".equals(Acc_Code[i])){
		            	 	   dataList.add(txt990421);//amt_name
		            	 	}else if("990621".equals(Acc_Code[i])){
		            		   dataList.add(txt990621);//amt_name
		            		}else if("990422".equals(Acc_Code[i])){//104.02.13
		            	 	   dataList.add(txt990422);//amt_name
		            	 	}else if("990622".equals(Acc_Code[i])){//104.02.13
		            		   dataList.add(txt990622);//amt_name   
		            		}else{
		            		   dataList.add("");//amt_name
		            		}		  
		            	 }else if(Report_no.equals("A03")||Report_no.equals("A04")){		//modify by 2354 2004.12.23
							if(Acc_Code[i].substring(Acc_Code[i].length()-1).equals("P")){
							   dataList.add(Utility.setNoPercentFormat(Amt[i]));//amt
		            		}else{
		            		   dataList.add(Utility.setNoCommaFormat(Amt[i]));//amt
			            	}
						 }else if(Report_no.equals("A05")){
							if(Acc_Code[i].substring(Acc_Code[i].length()-1).equals("P")){
							   dataList.add(Utility.setNoPercentFormat(Amt[i]));//amt
							   dataList.add("");//amt_name
		            		}else if(Acc_Code[i].substring(Acc_Code[i].length()-1).equals("N")){
		            		   dataList.add("0");//amt=0
							   dataList.add(Amt[i]);//amt_name
		            		}else{
		            		   dataList.add(Utility.setNoCommaFormat(Amt[i]));//amt
							   dataList.add("");//amt_name
			            	}
			             }else if(Report_no.equals("A13")){//104.10.12 add
							if(Acc_Code[i].equals("995400") || Acc_Code[i].equals("993810")){
							   dataList.add(Utility.setNoPercentFormat(Amt[i]));//amt							 
		            		}else{
		            		   dataList.add(Utility.setNoCommaFormat(Amt[i]));//amt							  
			            	}
						 }
						 updateDBDataList.add(dataList);
						 //System.out.println("add db data i="+i);
	            	}

	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
	            	System.out.println("updateDBDataList add");
	            	if(DBManager.updateDB_ps(updateDBList)){
	            	   System.out.println(Report_no+"Insert ok");
					   errMsg = errMsg + "相關資料寫入資料庫成功";

						//93.12.01 add 2295
						if(Report_no.equals("A01")){
					    	Date today = new Date();
    	                	int	batch_no = today.hashCode();
					    	errMsg = errMsg + UpdateA01.doParserReport_A01(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
						}else if(Report_no.equals("A02")){//add by winnin 2004.12.10
					    	Date today = new Date();
    	                	int	batch_no = today.hashCode();
					    	errMsg = errMsg + UpdateA02.doParserReport_A02(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
						}else if(Report_no.equals("A03")){//add by winnin 2004.12.23
					    	Date today = new Date();
    	                	int	batch_no = today.hashCode();
					    	errMsg = errMsg + UpdateA03.doParserReport_A03(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
						}else if(Report_no.equals("A04")){//add by winnin 2004.12.10
					    	Date today = new Date();
    	                	int	batch_no = today.hashCode();
					    	errMsg = errMsg + UpdateA04.doParserReport_A04(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
						}else if(Report_no.equals("A05")){//add by winnin 2004.12.10
					    	Date today = new Date();
    	                	int	batch_no = today.hashCode();
					    	errMsg = errMsg + UpdateA05.doParserReport_A05(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
						}else if(Report_no.equals("A99")){//add by 2295 2005.5.18
							Date today = new Date();
    	                	int	batch_no = today.hashCode();
					    	errMsg = errMsg + UpdateA99.doParserReport_A99(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
						    if(errMsg.startsWith("U")){//檢核成功
					           errMsg = "U:相關資料寫入資料庫成功";
					        }
					    }else if(Report_no.equals("A13")){//add by 2295 104.10.12
							Date today = new Date();
    	                	int	batch_no = today.hashCode();
					    	errMsg = errMsg + UpdateA13.doParserReport_A13(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
						    if(errMsg.startsWith("U")){//檢核成功
					           errMsg = "U:相關資料寫入資料庫成功";
					        }    
					    }
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
					}
    	   		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}

		return errMsg;
	}

    public String insertA06(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String bank_type/*93.12.18 add by 2295*/,String Report_no,String userid,String username) throws Exception{
    	String[]	Amt_3month	  = request.getParameterValues("amt_3month");
    	String[]	Amt_6month	  = request.getParameterValues("amt_6month");
    	String[]	Amt_1year	  = request.getParameterValues("amt_1year");
    	String[]	Amt_2year	  = request.getParameterValues("amt_2year");
    	String[]	Amt_over2year = request.getParameterValues("amt_over2year");
    	String[]	Amt_total	  = request.getParameterValues("amt_total");
		String[]	Acc_Code	  = request.getParameterValues("acc_code");

		String errMsg="";
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		List updateDBList = new ArrayList();//0:sql 1:data
		List updateDBSqlList = new ArrayList();
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data

		System.out.println("Acc_Code.size="+Acc_Code.length);
		try {
				sqlCmd.append("SELECT * FROM "+Report_no+" WHERE m_year=? AND m_month=? AND bank_code=?");
				paramList.add(S_YEAR);
				paramList.add(S_MONTH);
				paramList.add(bank_code);
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt_3month,amt_6month,amt_1year,amt_2year,amt_over2year,amt_total");
				System.out.println(""+Report_no+".size="+data.size());

				if (data.size() != 0){
				    errMsg = errMsg + "此筆資料已存在無法新增<br>";
				}else{
				    sqlCmd.delete(0,sqlCmd.length());
				    sqlCmd.append("INSERT INTO "+Report_no+" VALUES (?,?,?,?,?,?,?,?,?,?)");
					for(int i = 0; i < Acc_Code.length; i++) {
					    dataList =  new ArrayList();//儲存參數的data
					    dataList.add(S_YEAR);
					    dataList.add(S_MONTH);
					    dataList.add(bank_code);
					    dataList.add(Acc_Code[i].trim());
					    dataList.add(Utility.setNoCommaFormat(Amt_3month[i]));
					    dataList.add(Utility.setNoCommaFormat(Amt_6month[i]));
					    dataList.add(Utility.setNoCommaFormat(Amt_1year[i]));
					    dataList.add(Utility.setNoCommaFormat(Amt_2year[i]));
					    dataList.add(Utility.setNoCommaFormat(Amt_over2year[i]));
					    dataList.add(Utility.setNoCommaFormat(Amt_total[i]));
					    updateDBDataList.add(dataList);//1:傳內的參數List
	            	}

	            	updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
				    updateDBList.add(updateDBSqlList);
	            	//寫入資料庫
	            	if(DBManager.updateDB_ps(updateDBList)){
						//errMsg = errMsg + "相關資料寫入資料庫成功";
					    Date today = new Date();
    	                int	batch_no = today.hashCode();
					    errMsg = errMsg + UpdateA06.doParserReport_A06(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
					    if(errMsg.startsWith("U")){//檢核成功
					      errMsg = "U:相關資料寫入資料庫成功";
					    }
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
					}
    	   		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}

		return errMsg;
	}
    //Method add by jei 93.12.14
    public String insertB01(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String Report_no,String userid,String username) throws Exception{
		String[]	fund_master_no  	= request.getParameterValues("fund_master_no_c");
		String[]	fund_sub_no 		= request.getParameterValues("fund_sub_no_c");
		String[]	fund_next_no		= request.getParameterValues("fund_next_no_c");
		String[]	budget_amt			= request.getParameterValues("budget_amt");
		String[]	credit_pay_amt		= request.getParameterValues("credit_pay_amt");
		String[]	credit_pay_rate		= request.getParameterValues("credit_pay_rate");
		String[]	remark      		= request.getParameterValues("remark");

		String bank_type = ( request.getAttribute("bank_type")==null ) ? "3" : (String)request.getAttribute("bank_type");
        String errMsg ="";
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		List updateDBList = new ArrayList();//0:sql 1:data
		List updateDBSqlList = new ArrayList();
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data

		try {
				sqlCmd.append("SELECT * FROM B01 WHERE m_year=? AND m_month=?");
				paramList.add(S_YEAR);
				paramList.add(S_MONTH);
			    List data1 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");


				if (data1.size() != 0 ){
				    errMsg = errMsg + "此筆資料已存在無法新增<br>";
				}else{
			        sqlCmd.delete(0,sqlCmd.length());
				    sqlCmd.append("INSERT INTO B01 VALUES (?,?,?,?,?,?,?,?,?)");

					for (int i = 0; i < fund_master_no.length; i++) {
						/*System.out.println("fund_master_no.length=[" + fund_master_no.length + "]");
						System.out.println("fund_master_no=[" + fund_master_no[i] + "]");
						System.out.println("fund_sub_no=[" + fund_sub_no[i] + "]");
						System.out.println("fund_next_no=[" + fund_next_no[i] + "]");
						System.out.println("budget_amt=[" + budget_amt[i] + "]");
						System.out.println("credit_pay_amt=[" + credit_pay_amt[i] + "]");
						System.out.println("credit_pay_rate=[" + credit_pay_rate[i] + "]");
						System.out.println("remark=[" + remark[i] + "]");*/
						dataList =  new ArrayList();//儲存參數的data
					    dataList.add(S_YEAR);
					    dataList.add(S_MONTH);
					    dataList.add(fund_master_no[i]);
					    dataList.add(fund_sub_no[i]);
					    dataList.add(fund_next_no[i]);
					    dataList.add(Utility.setNoCommaFormat(budget_amt[i]));
					    dataList.add(Utility.setNoCommaFormat(credit_pay_amt[i]));
					    dataList.add(Utility.setNoPercentFormat(credit_pay_rate[i]));
					    dataList.add(remark[i]);
					    updateDBDataList.add(dataList);//1:傳內的參數List
					}
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);

					//寫入資料庫
	            	if(DBManager.updateDB_ps(updateDBList)){
						errMsg = errMsg + "相關資料寫入資料庫成功";
						//93.12.21 add egg
					    Date today = new Date();
    	                int	batch_no = today.hashCode();
					    errMsg = errMsg + UpdateB01.doParserReport_B01(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
					}
        		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}

		return errMsg;
	}

    //Method add by jei 93.12.16
    public String insertB02(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String Report_no,String userid,String username) throws Exception{
        String[]	run_master_no     	= request.getParameterValues("run_master_no_c");
		String[]	run_sub_no 	    	= request.getParameterValues("run_sub_no_c");
		String[]	run_next_no	    	= request.getParameterValues("run_next_no_c");
		String[]	loan_cnt_year		= request.getParameterValues("loan_cnt_year");
		String[]	loan_amt_year		= request.getParameterValues("loan_amt_year");
		String[]	loan_cnt_totacc		= request.getParameterValues("loan_cnt_totacc");
		String[]	loan_amt_totacc  	= request.getParameterValues("loan_amt_totacc");
		String[]	loan_cnt_bal		= request.getParameterValues("loan_cnt_bal");
		String[]	loan_amt_bal_subtot	= request.getParameterValues("loan_amt_bal_subtot");
		String[]	loan_amt_bal_fund	= request.getParameterValues("loan_amt_bal_fund");
		String[]	loan_amt_bal_bank  	= request.getParameterValues("loan_amt_bal_bank");

		String bank_type = ( request.getAttribute("bank_type")==null ) ? "3" : (String)request.getAttribute("bank_type");
        String errMsg ="";
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		List updateDBList = new ArrayList();//0:sql 1:data
		List updateDBSqlList = new ArrayList();
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data
		try {

				//宣告sqlCmd字串
				sqlCmd.append("SELECT * FROM B02 WHERE m_year=? AND m_month=?");
				paramList.add(S_YEAR);
				paramList.add(S_MONTH);
				//保留原數值型態資料
			    List data1 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");

				if (data1.size() != 0 ){
				    errMsg = errMsg + "此筆資料已存在無法新增<br>";
				}else{
			        sqlCmd.delete(0,sqlCmd.length());
				    sqlCmd.append("INSERT INTO B02 VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)");
				    
					for (int i = 0; i < run_master_no.length; i++) {
						System.out.println("run_master_no.length=[" + run_master_no.length + "]");
						System.out.println("run_master_no=[" + run_master_no[i] + "]");
						System.out.println("run_sub_no=[" + run_sub_no[i] + "]");
						System.out.println("run_next_no=[" + run_next_no[i] + "]");
						System.out.println("loan_cnt_year=[" + loan_cnt_year[i] + "]");
						System.out.println("loan_amt_year=[" + loan_amt_year[i] + "]");
						System.out.println("loan_cnt_totacc=[" + loan_cnt_totacc[i] + "]");
						System.out.println("loan_amt_totacc=[" + loan_amt_totacc[i] + "]");
						dataList =  new ArrayList();//儲存參數的data
					    dataList.add(S_YEAR);
					    dataList.add(S_MONTH);
					    dataList.add(run_master_no[i]);
					    dataList.add(run_sub_no[i]);
					    dataList.add(run_next_no[i]);
					    dataList.add(Utility.setNoCommaFormat(loan_cnt_year[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_year[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_cnt_totacc[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_totacc[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_cnt_bal[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_bal_subtot[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_bal_fund[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_bal_bank[i]));
					    updateDBDataList.add(dataList);//1:傳內的參數List					    
					}
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);

					//寫入資料庫
	            	if(DBManager.updateDB_ps(updateDBList)){
						errMsg = errMsg + "相關資料寫入資料庫成功";
						//93.12.21 add egg
					      Date today = new Date();
    	                  int	batch_no = today.hashCode();
					      errMsg = errMsg + UpdateB02.doParserReport_B02(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
					}
        		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}

		return errMsg;
	}

	//Method modify by egg 93.12.14
    public String insertB03(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String Report_no,String userid,String username) throws Exception{
		String errMsg="";
		//B03_1
    	String[]	funs_master_no_1		= request.getParameterValues("funs_master_no_1");
    	String[]	funs_sub_no_1			= request.getParameterValues("funs_sub_no_1");
    	String[]	funs_next_no_1			= request.getParameterValues("funs_next_no_1");
    	String[]	loan_cnt_totacc			= request.getParameterValues("loan_cnt_totacc");
    	String[]	loan_amt_totacc_fund	= request.getParameterValues("loan_amt_totacc_fund");
    	String[]	loan_amt_totacc_bank	= request.getParameterValues("loan_amt_totacc_bank");
    	String[]	loan_amt_totacc_tot		= request.getParameterValues("loan_amt_totacc_tot");
    	String[]	loan_cnt_bal			= request.getParameterValues("loan_cnt_bal");
    	String[]	loan_amt_bal_fund		= request.getParameterValues("loan_amt_bal_fund");
    	String[]	loan_amt_bal_bank		= request.getParameterValues("loan_amt_bal_bank");
    	String[]	loan_amt_bal_tot		= request.getParameterValues("loan_amt_bal_tot");

		//B03_2
    	String[]	funs_master_no_2		= request.getParameterValues("funs_master_no_2");
    	String[]	funs_sub_no_2			= request.getParameterValues("funs_sub_no_2");
    	String[]	funs_next_no_2			= request.getParameterValues("funs_next_no_2");
    	String[]	loan_amt_bal			= request.getParameterValues("loan_amt_bal");
    	String[]	loan_amt_over			= request.getParameterValues("loan_amt_over");
    	String[]	loan_rate_over			= request.getParameterValues("loan_rate_over");

    	//B03_3
    	String[]	funo_master_no			= request.getParameterValues("funo_master_no");
    	String[]	funo_sub_no				= request.getParameterValues("funo_sub_no");
    	String[]	funo_next_no			= request.getParameterValues("funo_next_no");
    	String[]	funo_amt				= request.getParameterValues("funo_amt");
    	String[]	funo_rate				= request.getParameterValues("funo_rate");

    	//B03_4
    	String[]	bank_no					= request.getParameterValues("bank_no");
    	String[]	machine_cnt				= request.getParameterValues("machine_cnt");
    	String[]	machine_amt				= request.getParameterValues("machine_amt");
    	String[]	land_cnt				= request.getParameterValues("land_cnt");
    	String[]	land_amt				= request.getParameterValues("land_amt");
    	String[]	house_cnt				= request.getParameterValues("house_cnt");
    	String[]	house_amt				= request.getParameterValues("house_amt");
    	String[]	build_cnt				= request.getParameterValues("build_cnt");
    	String[]	build_amt				= request.getParameterValues("build_amt");
    	String[]	tot_cnt					= request.getParameterValues("tot_cnt");
    	String[]	tot_amt					= request.getParameterValues("tot_amt");

		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		List updateDBList = new ArrayList();//0:sql 1:data
		List updateDBSqlList = new ArrayList();
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data

		String bank_type = ( request.getAttribute("bank_type")==null ) ? "3" : (String)request.getAttribute("bank_type");
		System.out.println("bank_type=["+bank_type+"]");

		try {
				//B03_1
				sqlCmd.append("SELECT * FROM " + Report_no + "_1 WHERE m_year=? AND m_month=?");
				paramList.add(S_YEAR);
				paramList.add(S_MONTH);
			    List data1 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");

				if (data1.size() != 0){
				    errMsg = errMsg + "此筆資料已存在無法新增<br>";
				}else{
					if(funs_master_no_1	==null) System.out.println("funs_master_no_1	is null");
					if(funs_sub_no_1	==null) System.out.println("funs_sub_no_1	is null");
					if(funs_next_no_1	==null) System.out.println("funs_master_no_1	is null");
					if(loan_cnt_totacc	==null) System.out.println("funs_next_no_1	is null");
					if(loan_amt_totacc_fund	==null) System.out.println("loan_amt_totacc_fund	is null");
					if(loan_amt_totacc_bank	==null) System.out.println("loan_amt_totacc_bank	is null");
					if(loan_amt_totacc_tot	==null) System.out.println("loan_amt_totacc_tot	is null");
					if(loan_cnt_bal	==null) System.out.println("loan_cnt_bal	is null");
					if(loan_amt_bal_fund	==null) System.out.println("loan_amt_bal_fund	is null");
					if(loan_amt_bal_bank	==null) System.out.println("loan_amt_bal_bank	is null");
					if(loan_amt_bal_tot	==null) System.out.println("loan_amt_bal_tot	is null");
					
					//Insert B03_1
					sqlCmd.delete(0,sqlCmd.length());
				    sqlCmd.append("INSERT INTO " + Report_no + "_1 VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)");
					for (int i = 0; i < loan_cnt_totacc.length; i++) {
					    dataList =  new ArrayList();//儲存參數的data
						dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						dataList.add(funs_master_no_1[i]);
						dataList.add(funs_sub_no_1[i]);
						dataList.add(funs_next_no_1[i]);
						dataList.add(Utility.setNoCommaFormat(loan_cnt_totacc[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_totacc_fund[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_totacc_bank[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_totacc_tot[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_cnt_bal[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_bal_fund[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_bal_bank[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_bal_tot[i]));
						updateDBDataList.add(dataList);//1:傳內的參數List	
					}
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
					
					//Insert B03_2
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO " + Report_no + "_2 VALUES (?,?,?,?,?,?,?,?)");
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List
					for (int i = 0; i < loan_amt_bal.length; i++) {
					    dataList =  new ArrayList();//儲存參數的data
						dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						dataList.add(funs_master_no_2[i]);
						dataList.add(funs_sub_no_2[i]);
						dataList.add(funs_next_no_2[i]);
						dataList.add(Utility.setNoCommaFormat(loan_amt_bal[i]));
						dataList.add(Utility.setNoCommaFormat(loan_amt_over[i]));
						dataList.add(Utility.setNoPercentFormat(loan_rate_over[i]));
						updateDBDataList.add(dataList);//1:傳內的參數List	
					}
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
					
					//Insert B03_3
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO " + Report_no + "_3 VALUES (?,?,?,?,?,?,?)");
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List
					System.out.println("funo_master_no.length=[" + funo_master_no.length + "]");
					for (int i = 0; i < funo_master_no.length; i++) {
					    dataList =  new ArrayList();//儲存參數的data
						dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						dataList.add(funo_master_no[i]);
						dataList.add(funo_sub_no[i]);
						dataList.add(funo_next_no[i]);
						dataList.add(Utility.setNoCommaFormat(funo_amt[i]));
						dataList.add(Utility.setNoPercentFormat(funo_rate[i]));
						updateDBDataList.add(dataList);//1:傳內的參數List							
					}
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
					//Insert B03_4
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO " + Report_no + "_4 VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)");
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List
					System.out.println("bank_no.length=[" + bank_no.length + "]");
					for (int i = 0; i < bank_no.length; i++) {
						dataList =  new ArrayList();//儲存參數的data
						
						dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						dataList.add(bank_no[i]);
						dataList.add(Utility.setNoCommaFormat(machine_cnt[i]));
						dataList.add(Utility.setNoCommaFormat(machine_amt[i]));
						dataList.add(Utility.setNoCommaFormat(land_cnt[i]));
						dataList.add(Utility.setNoCommaFormat(land_amt[i]));
						dataList.add(Utility.setNoCommaFormat(house_cnt[i]));
						dataList.add(Utility.setNoCommaFormat(house_amt[i]));
						dataList.add(Utility.setNoCommaFormat(build_cnt[i]));
						dataList.add(Utility.setNoCommaFormat(build_amt[i]));
						dataList.add(Utility.setNoCommaFormat(tot_cnt[i]));
						dataList.add(Utility.setNoCommaFormat(tot_amt[i]));
						
						updateDBDataList.add(dataList);//1:傳內的參數List		
					}
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
					
	            	if(DBManager.updateDB_ps(updateDBList)){	
						errMsg = errMsg + "相關資料寫入資料庫成功";
						//93.12.21 add egg
					    Date today = new Date();
    	                int	batch_no = today.hashCode();
					    errMsg = errMsg + UpdateB03.doParserReport_B03(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗";
					}
        		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}
		
		return errMsg;
	}

    //Method add by egg 93.12.10
    public String insertM01(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String Report_no,String userid,String username) throws Exception{
		String[]	guarantee_item_no_c	= request.getParameterValues("guarantee_item_no_c");
		String[]	data_range_c		= request.getParameterValues("data_range_c");
		String[]	guarantee_cnt		= request.getParameterValues("guarantee_cnt");
		String[]	loan_amt			= request.getParameterValues("loan_amt");
		String[]	guarantee_amt		= request.getParameterValues("guarantee_amt");
		String[]	loan_bal			= request.getParameterValues("loan_bal");
		String[]	guarantee_bal		= request.getParameterValues("guarantee_bal");
		String[]	over_notpush_cnt	= request.getParameterValues("over_notpush_cnt");
		String[]	over_notpush_bal	= request.getParameterValues("over_notpush_bal");
		String[]	over_okpush_cnt		= request.getParameterValues("over_okpush_cnt");
		String[]	over_okpush_bal		= request.getParameterValues("over_okpush_bal");
		String[]	repay_tot_cnt		= request.getParameterValues("repay_tot_cnt");
		String[]	repay_tot_amt		= request.getParameterValues("repay_tot_amt");
		String[]	repay_bal_cnt		= request.getParameterValues("repay_bal_cnt");
		String[]	repay_bal_amt		= request.getParameterValues("repay_bal_amt");
		//年份資料表示轉換93-->093
		String s_year="0000"+S_YEAR; s_year=s_year.substring(s_year.length()-3);
		String s_month="0000"+S_MONTH; s_month=s_month.substring(s_month.length()-2);
		System.out.println("轉換後s_year =[" + s_year + "]");
		System.out.println("轉換後s_month=[" + s_month + "]");

		String bank_type = ( request.getAttribute("bank_type")==null ) ? "4" : (String)request.getAttribute("bank_type");
		String errMsg ="";
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		List updateDBList = new ArrayList();//0:sql 1:data
		List updateDBSqlList = new ArrayList();
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data
		System.out.println("Input S_YEAR=" + S_YEAR + "Inpute bank_code=" + bank_code + "Report_no=" + Report_no);
		try {
				
				sqlCmd.append("SELECT * FROM M01 WHERE m_year=? AND m_month=?");
				paramList.add(S_YEAR);
				paramList.add(S_MONTH);
				//保留原數值型態資料
			    List data1 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");

				if (data1==null || data1.size() != 0 ){
				    errMsg = errMsg + "此筆資料已存在無法新增<br>";
				}else{
				    sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO M01 VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
					for (int i = 0; i < guarantee_item_no_c.length; i++) {
						dataList =  new ArrayList();//儲存參數的data
						dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						dataList.add(guarantee_item_no_c[i]);
						dataList.add(s_year + data_range_c[i]);
						dataList.add(Utility.setNoCommaFormat(guarantee_cnt[i]));
						dataList.add(Utility.setNoCommaFormat(loan_amt[i]));
						dataList.add(Utility.setNoCommaFormat(guarantee_amt[i]));
						dataList.add(Utility.setNoCommaFormat(loan_bal[i]));
						dataList.add(Utility.setNoCommaFormat(guarantee_bal[i]));
						dataList.add(Utility.setNoCommaFormat(over_notpush_cnt[i]));
						dataList.add(Utility.setNoCommaFormat(over_notpush_bal[i]));
						dataList.add(Utility.setNoCommaFormat(over_okpush_cnt[i]));
						dataList.add(Utility.setNoCommaFormat(over_okpush_bal[i]));
						dataList.add(Utility.setNoCommaFormat(repay_tot_cnt[i]));
						dataList.add(Utility.setNoCommaFormat(repay_tot_amt[i]));
						dataList.add(Utility.setNoCommaFormat(repay_bal_cnt[i]));
						dataList.add(Utility.setNoCommaFormat(repay_bal_amt[i]));
						updateDBDataList.add(dataList);//1:傳內的參數List		
					}
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
					
					if(DBManager.updateDB_ps(updateDBList)){	
						errMsg = errMsg + "相關資料寫入資料庫成功";
					    Date today = new Date();
    	                int	batch_no = today.hashCode();
					    errMsg = errMsg + UpdateM01.doParserReport_M01(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
					}
        		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}

		return errMsg;
	}

    //jei 931210
    public String insertM02(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String Report_no,String userid,String username) throws Exception{
		String[]	loan_unit_no    	= request.getParameterValues("loan_unit_no_c");
		String[]	data_range			= request.getParameterValues("data_range_c");
		String[]	guarantee_cnt		= request.getParameterValues("guarantee_cnt");
		String[]	loan_amt			= request.getParameterValues("loan_amt");
		String[]	guarantee_amt		= request.getParameterValues("guarantee_amt");
		String[]	loan_bal			= request.getParameterValues("loan_bal");
		String[]	guarantee_bal		= request.getParameterValues("guarantee_bal");
		String[]	over_notpush_cnt	= request.getParameterValues("over_notpush_cnt");
		String[]	over_notpush_bal	= request.getParameterValues("over_notpush_bal");
		String[]	over_okpush_cnt		= request.getParameterValues("over_okpush_cnt");
		String[]	over_okpush_bal		= request.getParameterValues("over_okpush_bal");
		String[]	repay_tot_cnt		= request.getParameterValues("repay_tot_cnt");
		String[]	repay_tot_amt		= request.getParameterValues("repay_tot_amt");
		String[]	repay_bal_cnt		= request.getParameterValues("repay_bal_cnt");
		String[]	repay_bal_amt		= request.getParameterValues("repay_bal_amt");

		String errMsg="";		
		String bank_type = ( request.getAttribute("bank_type")==null ) ? "4" : (String)request.getAttribute("bank_type");
		
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		List updateDBList = new ArrayList();//0:sql 1:data
		List updateDBSqlList = new ArrayList();
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data

		//年份資料表示轉換93-->093
		String s_year="0000"+S_YEAR; s_year=s_year.substring(s_year.length()-3);
		System.out.println("轉換後s_year =[" + s_year + "]");

		try {
				
				sqlCmd.append("SELECT * FROM M02 WHERE m_year=? AND m_month=?");
				paramList.add(S_YEAR);
				paramList.add(S_MONTH);
				//保留原數值型態資料
			    List data1 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");

				if (data1.size() != 0 ){
				    errMsg = errMsg + "此筆資料已存在無法新增<br>";
				}else{
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO M01 VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
					for (int i = 0; i < loan_unit_no.length; i++) {
						dataList =  new ArrayList();//儲存參數的data
						dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						dataList.add(loan_unit_no[i]);
						dataList.add(s_year + data_range[i]);
						dataList.add(Utility.setNoCommaFormat(guarantee_cnt[i]));
						dataList.add(Utility.setNoCommaFormat(loan_amt[i]));
						dataList.add(Utility.setNoCommaFormat(guarantee_amt[i]));
						dataList.add(Utility.setNoCommaFormat(loan_bal[i]));
						dataList.add(Utility.setNoCommaFormat(guarantee_bal[i]));
						dataList.add(Utility.setNoCommaFormat(over_notpush_cnt[i]));
					    dataList.add(Utility.setNoCommaFormat(over_notpush_bal[i]));
					    dataList.add(Utility.setNoCommaFormat(over_okpush_cnt[i]));
					    dataList.add(Utility.setNoCommaFormat(over_okpush_bal[i]));
					    dataList.add(Utility.setNoCommaFormat(repay_tot_cnt[i]));
					    dataList.add(Utility.setNoCommaFormat(repay_tot_amt[i]));
					    dataList.add(Utility.setNoCommaFormat(repay_bal_cnt[i]));
					    dataList.add(Utility.setNoCommaFormat(repay_bal_amt[i]));
						updateDBDataList.add(dataList);//1:傳內的參數List		
					}
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
					
					if(DBManager.updateDB_ps(updateDBList)){
						errMsg = errMsg + "相關資料寫入資料庫成功";
					    	Date today = new Date();
    	                	int	batch_no = today.hashCode();
					    	errMsg = errMsg + UpdateM02.doParserReport_M02(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
					}
        		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗<br>";
		}

		return errMsg;
	}

    //Method modify by egg 2004.12.12
    public String insertM03(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String Report_no,String userid,String username) throws Exception{
		String errMsg="";
		
		//M03
		String[]	div_no						= request.getParameterValues("div_no");
    	String[]	guarantee_cnt_month			= request.getParameterValues("guarantee_cnt_month");
    	String[]	loan_amt_month				= request.getParameterValues("loan_amt_month");
    	String[]	guarantee_amt_month			= request.getParameterValues("guarantee_amt_month");
    	String[]	guarantee_cnt_year			= request.getParameterValues("guarantee_cnt_year");
    	String[]	loan_amt_year				= request.getParameterValues("loan_amt_year");
    	String[]	guarantee_amt_year			= request.getParameterValues("guarantee_amt_year");
    	String[]	guarantee_bal_totacc		= request.getParameterValues("guarantee_bal_totacc");
    	String[]	guarantee_bal_totacc_over	= request.getParameterValues("guarantee_bal_totacc_over");
    	String[]	repay_bal_totacc			= request.getParameterValues("repay_bal_totacc");

		//M03 Note
		String[]	note_no						= request.getParameterValues("note_no");
    	String[]	note_amt_rate				= request.getParameterValues("note_amt_rate");

    	String bank_type = ( request.getAttribute("bank_type")==null ) ? "4" : (String)request.getAttribute("bank_type");
		System.out.println("Report_no=" + Report_no);
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		List updateDBList = new ArrayList();//0:sql 1:data
		List updateDBSqlList = new ArrayList();
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data

		try {
				sqlCmd.append("SELECT * FROM M03 WHERE m_year=? AND m_month=?");
				paramList.add(S_YEAR);
				paramList.add(S_MONTH);
			    List data1 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");
			    sqlCmd.delete(0,sqlCmd.length());
				sqlCmd.append("SELECT * FROM M03_NOTE WHERE m_year=? AND m_month=?");				
			    List data2 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");

				if (data1.size() != 0 || data2.size() != 0 ){
				    errMsg = errMsg + "此筆資料已存在無法新增<br>";
				}else{
					System.out.println("div_no.length="+div_no.length);
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO M03 VALUES (?,?,?,?,?,?,?,?,?,?,?,?)");
					for (int i = 0; i < div_no.length; i++) {
						dataList =  new ArrayList();//儲存參數的data     
						dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						dataList.add(div_no[i]);
						if(div_no[i].equals("MP") || div_no[i].equals("YP") || div_no[i].equals("NT1P")){							
					    	dataList.add(Utility.setNoPercentFormat(guarantee_cnt_month[i]));
					    	dataList.add(Utility.setNoPercentFormat(loan_amt_month[i]));
					    	dataList.add(Utility.setNoPercentFormat(guarantee_amt_month[i]));
					    	dataList.add(Utility.setNoPercentFormat(guarantee_cnt_year[i]));
					    	dataList.add(Utility.setNoPercentFormat(loan_amt_year[i]));
					    	dataList.add(Utility.setNoPercentFormat(guarantee_amt_year[i]));
					    	dataList.add(Utility.setNoPercentFormat(guarantee_bal_totacc[i]));
					    	dataList.add(Utility.setNoPercentFormat(guarantee_bal_totacc_over[i]));
					    	dataList.add(Utility.setNoPercentFormat(repay_bal_totacc[i]));
						}else{							
					    	dataList.add(Utility.setNoCommaFormat(guarantee_cnt_month[i]));
					    	dataList.add(Utility.setNoCommaFormat(loan_amt_month[i]));
					    	dataList.add(Utility.setNoCommaFormat(guarantee_amt_month[i]));
					    	dataList.add(Utility.setNoCommaFormat(guarantee_cnt_year[i]));
					    	dataList.add(Utility.setNoCommaFormat(loan_amt_year[i]));
					    	dataList.add(Utility.setNoCommaFormat(guarantee_amt_year[i]));
					    	dataList.add(Utility.setNoCommaFormat(guarantee_bal_totacc[i]));
					    	dataList.add(Utility.setNoCommaFormat(guarantee_bal_totacc_over[i]));
					    	dataList.add(Utility.setNoCommaFormat(repay_bal_totacc[i]));
						}
						updateDBDataList.add(dataList);//1:傳內的參數List		     
					}
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
					
					
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO M03_NOTE VALUES (?,?,?)");
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List
					for (int i = 0; i < note_no.length; i++) {
						System.out.println("note_no["+i+"]="+note_no[i]);
						dataList =  new ArrayList();//儲存參數的data     
						dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						dataList.add(note_no[i]);
						if(note_no[i].equals("NT1P")){							
					       dataList.add(Utility.setNoPercentFormat(note_amt_rate[i]));
                        }else{
						   dataList.add(Utility.setNoCommaFormat(note_amt_rate[i]));
						}
						updateDBDataList.add(dataList);//1:傳內的參數List		
					}
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
					
	            	if(DBManager.updateDB_ps(updateDBList)){	
						errMsg = errMsg + "相關資料寫入資料庫成功";
						//93.12.21 add egg
					    Date today = new Date();
    	                int	batch_no = today.hashCode();
					    errMsg = errMsg + UpdateM03.doParserReport_M03(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
					}
        		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗<br>";
		}
		return errMsg;
	}

    //jei 931212
    public String insertM04(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String Report_no,String userid,String username) throws Exception{
		String[]	loan_use_no         	= request.getParameterValues("loan_use_no_c");
		String[]	guarantee_no_month		= request.getParameterValues("guarantee_no_month");
		String[]	guarantee_no_month_p	= request.getParameterValues("guarantee_no_month_p");
		String[]	loan_amt_month	    	= request.getParameterValues("loan_amt_month");
		String[]	loan_amt_month_p		= request.getParameterValues("loan_amt_month_p");
		String[]	guarantee_amt_month		= request.getParameterValues("guarantee_amt_month");
		String[]	guarantee_amt_month_p	= request.getParameterValues("guarantee_amt_month_p");
		String[]	guarantee_no_year	    = request.getParameterValues("guarantee_no_year");
		String[]	guarantee_no_year_p		= request.getParameterValues("guarantee_no_year_p");
		String[]	loan_amt_year		    = request.getParameterValues("loan_amt_year");
		String[]	loan_amt_year_p	    	= request.getParameterValues("loan_amt_year_p");
		String[]	guarantee_amt_year	 	= request.getParameterValues("guarantee_amt_year");
		String[]	guarantee_amt_year_p	= request.getParameterValues("guarantee_amt_year_p");
		String[]	guarantee_no_totacc		= request.getParameterValues("guarantee_no_totacc");
		String[]	guarantee_no_totacc_p	= request.getParameterValues("guarantee_no_totacc_p");
		String[]	loan_amt_totacc		    = request.getParameterValues("loan_amt_totacc");
		String[]	loan_amt_totacc_p		= request.getParameterValues("loan_amt_totacc_p");
    	String[]	guarantee_amt_totacc	= request.getParameterValues("guarantee_amt_totacc");
		String[]	guarantee_amt_totacc_p	= request.getParameterValues("guarantee_amt_totacc_p");

		String errMsg="";		
		String bank_type = ( request.getAttribute("bank_type")==null ) ? "4" : (String)request.getAttribute("bank_type");

		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		List updateDBList = new ArrayList();//0:sql 1:data
		List updateDBSqlList = new ArrayList();
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data
		
		try {
				sqlCmd.append("SELECT * FROM M04 WHERE m_year=? AND m_month=?");
				paramList.add(S_YEAR);
				paramList.add(S_MONTH);
				//保留原數值型態資料
			    List data1 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");

				if (data1.size() != 0 ){
				    errMsg = errMsg + "此筆資料已存在無法新增<br>";
				}else{
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO M04 VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
					for (int i = 0; i < loan_use_no.length; i++) {
						dataList =  new ArrayList();//儲存參數的data     
						dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						dataList.add(loan_use_no[i]);
					    dataList.add(Utility.setNoCommaFormat(guarantee_no_month[i]));
					    dataList.add(Utility.setNoPercentFormat(guarantee_no_month_p[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_month[i]));
					    dataList.add(Utility.setNoPercentFormat(loan_amt_month_p[i]));
					    dataList.add(Utility.setNoCommaFormat(guarantee_amt_month[i]));
					    dataList.add(Utility.setNoPercentFormat(guarantee_amt_month_p[i]));
					    dataList.add(Utility.setNoCommaFormat(guarantee_no_year[i]));
					    dataList.add(Utility.setNoPercentFormat(guarantee_no_year_p[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_year[i]));
					    dataList.add(Utility.setNoPercentFormat(loan_amt_year_p[i]));
					    dataList.add(Utility.setNoCommaFormat(guarantee_amt_year[i]));
					    dataList.add(Utility.setNoPercentFormat(guarantee_amt_year_p[i]));
					    dataList.add(Utility.setNoCommaFormat(guarantee_no_totacc[i]));
					    dataList.add(Utility.setNoPercentFormat(guarantee_no_totacc_p[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_totacc[i]));
					    dataList.add(Utility.setNoPercentFormat(loan_amt_totacc_p[i]));
					    dataList.add(Utility.setNoCommaFormat(guarantee_amt_totacc[i]));
					    dataList.add(Utility.setNoPercentFormat(guarantee_amt_totacc_p[i]));
						updateDBDataList.add(dataList);//1:傳內的參數List		
					}
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
					
	            	if(DBManager.updateDB_ps(updateDBList)){	
						errMsg = errMsg + "相關資料寫入資料庫成功";
						//93.12.01 add egg
					    Date today = new Date();
    	                int	batch_no = today.hashCode();
					    errMsg = errMsg + UpdateM04.doParserReport_M04(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
					}
        		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}
		return errMsg;
	}

    //add by winnin 2004.12.09
    public String insertM05(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String Report_no,String userid,String username) throws Exception{
    	String[]	Loan_Unit_No	= request.getParameterValues("loan_unit_no");
    	String[]	Period_No		= request.getParameterValues("period_no");
    	String[]	Item_No			= request.getParameterValues("item_no");
    	String[]	Repay_Cnt		= request.getParameterValues("repay_cnt");
    	String[]	Repay_Amt		= request.getParameterValues("repay_amt");
    	String[]	Run_Notgood_Cnt	= request.getParameterValues("run_notgood_cnt");
    	String[]	Run_Notgood_Amt	= request.getParameterValues("run_notgood_amt");
    	String[]	Turn_Out_Cnt= request.getParameterValues("turn_out_cnt");
    	String[]	Turn_Out_Amt= request.getParameterValues("turn_out_amt");
    	String[]	Diease_Cnt		= request.getParameterValues("diease_cnt");
    	String[]	Dieaserepay_Amt	= request.getParameterValues("dieaserepay_amt");
    	String[]	Disaster_Cnt	= request.getParameterValues("disaster_cnt");
    	String[]	Disaster_Amt	= request.getParameterValues("disaster_amt");
    	String[]	Corun_Out_Cnt	= request.getParameterValues("corun_out_cnt");
    	String[]	Corun_Out_Amt	= request.getParameterValues("corun_out_amt");
    	String[]	Other_Cnt		= request.getParameterValues("other_cnt");
    	String[]	Other_Amt		= request.getParameterValues("other_amt");

		//M05_TOTACC
    	String[]	Loan_Unit_No_C	= request.getParameterValues("loan_unit_no_c");
    	String[]	Fix_No	= request.getParameterValues("fix_no_c");
    	String[]	Guarantee_No_Totacc	= request.getParameterValues("guarantee_no_totacc");
    	String[]	Guarantee_Amt_Totacc	= request.getParameterValues("guarantee_amt_totacc");

		//M05_NOTE
    	String[]	Note_No	= request.getParameterValues("note_no");
    	String[]	Note_Amt_Rate	= request.getParameterValues("note_amt_rate");

		String bank_type = ( request.getAttribute("bank_type")==null ) ? "4" : (String)request.getAttribute("bank_type");
		String errMsg="";
		
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		List updateDBList = new ArrayList();//0:sql 1:data
		List updateDBSqlList = new ArrayList();
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data

		try {
				sqlCmd.append("SELECT * FROM M05 WHERE m_year=? AND m_month=?");
				paramList.add(S_YEAR);
				paramList.add(S_MONTH);
			    List data1 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");
			    sqlCmd.delete(0,sqlCmd.length());
				sqlCmd.append("SELECT * FROM M05_TOTACC WHERE m_year=? AND m_month=?");
			    List data2 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");
				sqlCmd.append("SELECT * FROM M05_NOTE WHERE m_year=? AND m_month=?");
			    List data3 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");
				if (data1.size() != 0 || data2.size() != 0 || data3.size() != 0){
				    errMsg = errMsg + "此筆資料已存在無法新增<br>";
				}else{
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO M05 VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
					for (int i = 0; i < Period_No.length; i++) {
						dataList =  new ArrayList();//儲存參數的data   
						dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						dataList.add(Loan_Unit_No[i]);
						dataList.add(Period_No[i]);
						dataList.add(Item_No[i]);
					    dataList.add(Utility.setNoCommaFormat(Repay_Cnt[i]));
					    dataList.add(Utility.setNoCommaFormat(Repay_Amt[i]));
					    dataList.add(Utility.setNoCommaFormat(Run_Notgood_Cnt[i]));
					    dataList.add(Utility.setNoCommaFormat(Run_Notgood_Amt[i]));
					    dataList.add(Utility.setNoCommaFormat(Turn_Out_Cnt[i]));
					    dataList.add(Utility.setNoCommaFormat(Turn_Out_Amt[i]));
					    dataList.add(Utility.setNoCommaFormat(Diease_Cnt[i]));
					    dataList.add(Utility.setNoCommaFormat(Dieaserepay_Amt[i]));
					    dataList.add(Utility.setNoCommaFormat(Disaster_Cnt[i]));
					    dataList.add(Utility.setNoCommaFormat(Disaster_Amt[i]));
					    dataList.add(Utility.setNoCommaFormat(Corun_Out_Cnt[i]));
					    dataList.add(Utility.setNoCommaFormat(Corun_Out_Amt[i]));
					    dataList.add(Utility.setNoCommaFormat(Other_Cnt[i]));
					    dataList.add(Utility.setNoCommaFormat(Other_Amt[i]));
						updateDBDataList.add(dataList);//1:傳內的參數List	
					}
					
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
					
					//M05_TOTACC
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO M05_TOTACC VALUES (?,?,?,?,?,?)");
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List
					
					for (int i = 0; i < Loan_Unit_No_C.length; i++) {
						dataList =  new ArrayList();//儲存參數的data   
						dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						dataList.add(Loan_Unit_No_C[i]);
						dataList.add(Fix_No[i]);
						dataList.add(Utility.setNoCommaFormat(Guarantee_No_Totacc[i]));
						dataList.add(Utility.setNoCommaFormat(Guarantee_Amt_Totacc[i]));
						updateDBDataList.add(dataList);//1:傳內的參數List	
					}
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
					
					//M05_NOTE
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO M05_NOTE VALUES (?,?,?,?)");
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List
					for (int i = 0; i < Note_No.length; i++) {
						dataList =  new ArrayList();//儲存參數的data   
						dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						dataList.add(Note_No[i]);
						dataList.add(Utility.setNoCommaFormat(Note_Amt_Rate[i]));
						updateDBDataList.add(dataList);//1:傳內的參數List	
					}
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
					
	            	if(DBManager.updateDB_ps(updateDBList)){	
						errMsg = errMsg + "相關資料寫入資料庫成功";
					    Date today = new Date();
    	                int	batch_no = today.hashCode();
					    errMsg = errMsg + UpdateM05.doParserReport_M05(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗";
					}
        		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}

		return errMsg;
	}


    //Method modify by egg 93.12.13
    public String insertM06_M07(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String Report_no,String userid,String username) throws Exception{
		String errMsg="";
		String[]	area_no					= request.getParameterValues("area_no");
    	String[]	guarantee_no_month		= request.getParameterValues("guarantee_no_month");
    	String[]	guarantee_amt_month		= request.getParameterValues("guarantee_amt_month");
    	String[]	loan_amt_month			= request.getParameterValues("loan_amt_month");
    	String[]	guarantee_no_year		= request.getParameterValues("guarantee_no_year");
    	String[]	guarantee_amt_year		= request.getParameterValues("guarantee_amt_year");
    	String[]	loan_amt_year			= request.getParameterValues("loan_amt_year");
    	String[]	guarantee_no_totacc		= request.getParameterValues("guarantee_no_totacc");
    	String[]	guarantee_amt_totacc	= request.getParameterValues("guarantee_amt_totacc");
    	String[]	loan_amt_totacc			= request.getParameterValues("loan_amt_totacc");
    	String[]	guarantee_bal_no		= request.getParameterValues("guarantee_bal_no");
    	String[]	guarantee_bal_amt		= request.getParameterValues("guarantee_bal_amt");
    	String[]	guarantee_bal_p			= request.getParameterValues("guarantee_bal_p");
    	String[]	loan_bal				= request.getParameterValues("loan_bal");

		
		System.out.println("Report_no=["+Report_no+"]");
		String bank_type = ( request.getAttribute("bank_type")==null ) ? "4" : (String)request.getAttribute("bank_type");
		System.out.println("bank_type=["+bank_type+"]");
		
	    StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		List updateDBList = new ArrayList();//0:sql 1:data
		List updateDBSqlList = new ArrayList();
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data
		
		try {
				sqlCmd.append("SELECT * FROM " + Report_no + " WHERE m_year=? AND m_month=?");
				paramList.add(S_YEAR);
				paramList.add(S_MONTH);
			    List data1 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");

				if (data1.size() != 0){
				    errMsg = errMsg + "此筆資料已存在無法新增<br>";
				}else{
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO " + Report_no + " VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
					for (int i = 0; i < area_no.length; i++) {
						dataList =  new ArrayList();//儲存參數的data               
						dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						dataList.add(area_no[i]);
					    dataList.add(Utility.setNoCommaFormat(guarantee_no_month[i]));
					    dataList.add(Utility.setNoCommaFormat(guarantee_amt_month[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_month[i]));
					    dataList.add(Utility.setNoCommaFormat(guarantee_no_year[i]));
					    dataList.add(Utility.setNoCommaFormat(guarantee_amt_year[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_year[i]));
					    dataList.add(Utility.setNoCommaFormat(guarantee_no_totacc[i]));
					    dataList.add(Utility.setNoCommaFormat(guarantee_amt_totacc[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_totacc[i]));
					    dataList.add(Utility.setNoCommaFormat(guarantee_bal_no[i]));
					    dataList.add(Utility.setNoCommaFormat(guarantee_bal_amt[i]));
					    dataList.add(Utility.setNoPercentFormat(guarantee_bal_p[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_bal[i]));
						updateDBDataList.add(dataList);//1:傳內的參數List	
					}
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
					
	            	if(DBManager.updateDB_ps(updateDBList)){	
						errMsg = errMsg + "相關資料寫入資料庫成功";
						//93.12.21 add egg
						if(Report_no.equals("M06")){
					    	Date today = new Date();
    	                	int	batch_no = today.hashCode();
					    	errMsg = errMsg + UpdateM06.doParserReport_M06(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
						}else{
							Date today = new Date();
    	                	int	batch_no = today.hashCode();
					    	errMsg = errMsg + UpdateM07.doParserReport_M07(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
						}
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗";
					}
        		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}
		
		return errMsg;
	}

    //Method modify by egg 93.12.16
    public String insertM08(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String Report_no,String userid,String username) throws Exception{
		String errMsg="";		

    	String[]	id_no					= request.getParameterValues("id_no");
    	String[]	data_range				= request.getParameterValues("data_range");
    	String[]	guarantee_no_month		= request.getParameterValues("guarantee_no_month");
    	String[]	loan_amt_month			= request.getParameterValues("loan_amt_month");
    	String[]	guarantee_amt_month		= request.getParameterValues("guarantee_amt_month");
    	String[]	guarantee_bal_month		= request.getParameterValues("guarantee_bal_month");
    	String[]	guarantee_bal_p			= request.getParameterValues("guarantee_bal_p");

	

		System.out.println("Report_no=["+Report_no+"]");
		String bank_type = ( request.getAttribute("bank_type")==null ) ? "4" : (String)request.getAttribute("bank_type");
		System.out.println("S_YEAR =[" + S_YEAR + "]");
		System.out.println("S_MONTH=[" + S_MONTH + "]");
		//年份資料表示轉換93-->093
		String s_year="0000"+S_YEAR; s_year=s_year.substring(s_year.length()-3);
		String s_month="0000"+S_MONTH; s_month=s_month.substring(s_month.length()-2);
		System.out.println("轉換後s_year =[" + s_year + "]");
		System.out.println("轉換後s_month=[" + s_month + "]");
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		List updateDBList = new ArrayList();//0:sql 1:data
		List updateDBSqlList = new ArrayList();
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data

		try {
				sqlCmd.append("SELECT * FROM " + Report_no + " WHERE m_year= ? AND m_month=?");
			    paramList.add(S_YEAR);
				paramList.add(S_MONTH);
			    List data1 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");

				if (data1.size() != 0){
				    errMsg = errMsg + "此筆資料已存在無法新增<br>";
				}else{
					System.out.println("id_no.length=["+ id_no.length + "]");
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO " + Report_no + " VALUES (?,?,?,?,?,?,?,?,?)");
					for (int i = 0; i < id_no.length; i++) {
						dataList =  new ArrayList();//儲存參數的data       
						dataList.add(S_YEAR);
						dataList.add(S_MONTH); 
						dataList.add(id_no[i]);
						dataList.add(s_year + data_range[i]);
					    dataList.add(Utility.setNoCommaFormat(guarantee_no_month[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_month[i]));
					    dataList.add(Utility.setNoCommaFormat(guarantee_amt_month[i]));
					    dataList.add(Utility.setNoCommaFormat(guarantee_bal_month[i]));
					    dataList.add(Utility.setNoPercentFormat(guarantee_bal_p[i]));
						updateDBDataList.add(dataList);//1:傳內的參數List	     
					}

					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
					
	            	if(DBManager.updateDB_ps(updateDBList)){	
						errMsg = errMsg + "相關資料寫入資料庫成功";
					      Date today = new Date();
    	                  int	batch_no = today.hashCode();
					      errMsg = errMsg + UpdateM08.doParserReport_M08(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗";
					}
        		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}
		
		return errMsg;
	}

	//94.11.10 add by 2295
    public String insertF01(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String Report_no,String userid,String username) throws Exception{
        String titleidx[] = {"A","B","C","D","E"};        
		String bank_type = ( request.getAttribute("bank_type")==null ) ? "4" : (String)request.getAttribute("bank_type");
		String errMsg ="";	
		System.out.println("Input S_YEAR=" + S_YEAR + "Inpute bank_code=" + bank_code + "Report_no=" + Report_no);
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		List updateDBList = new ArrayList();//0:sql 1:data
		List updateDBSqlList = new ArrayList();
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data
		try {
				sqlCmd.append("select count(*) as count from f01 WHERE m_year=? AND m_month=? and bank_code=?");
				paramList.add(S_YEAR);				
				paramList.add(S_MONTH);
				paramList.add(bank_code);
				//保留原數值型態資料
			    List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"count");
				if (dbData!=null && (!(((DataObject)dbData.get(0)).getValue("count")).toString().equals("0")) ){
				    errMsg = errMsg + "此筆資料已存在無法新增<br>";
				}else{

				    //取出form裡的所有變數===================================
					Enumeration ep = request.getParameterNames();
					Hashtable t = new Hashtable();
					String name = "";
					for ( ; ep.hasMoreElements() ; ) {
				 		name = (String)ep.nextElement();
				 		if(request.getParameter(name) == null){
				 		  t.put( name, "" );
				 		}else{
				 		  t.put( name, request.getParameter(name) );
				 		}
				 		System.out.println(name+"="+request.getParameter(name));
					}
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO F01 VALUES (?,?,?,?,?,?,?,?,?,?)");
                    for(int j=0;j<titleidx.length;j++){//A,B,C,D,E
		                for(int k=1;k<=5;k++){
						        if(t.get(titleidx[j]+k+"1") != null ) {
						          /*
						          System.out.println((String)t.get(titleidx[j]+k+"1"));
						          System.out.println((String)t.get(titleidx[j]+k+"2"));
						          System.out.println((String)t.get(titleidx[j]+k+"3"));
						          System.out.println((String)t.get(titleidx[j]+k+"4"));
						          System.out.println((String)t.get(titleidx[j]+k+"5"));
						          */
						          dataList =  new ArrayList();//儲存參數的data       
						          dataList.add(S_YEAR);
						          dataList.add(S_MONTH);
						          dataList.add(bank_code);
						          dataList.add(titleidx[j]);
						          dataList.add(k);
						          dataList.add(Utility.setNoCommaFormat((String)t.get(titleidx[j]+k+"1")));
						          dataList.add(Utility.setNoCommaFormat((String)t.get(titleidx[j]+k+"2")));
						          dataList.add(Utility.setNoCommaFormat((String)t.get(titleidx[j]+k+"3")));
						          dataList.add(Utility.setNoCommaFormat((String)t.get(titleidx[j]+k+"4")));
						          dataList.add(Utility.setNoCommaFormat((String)t.get(titleidx[j]+k+"5")));
						          updateDBDataList.add(dataList);//1:傳內的參數List	     
						       }
					    }//end of 1~5
				    }//end of "A,B,C,D,E"

					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
					
	            	if(DBManager.updateDB_ps(updateDBList)){	
						//errMsg = errMsg + "相關資料寫入資料庫成功";
					    Date today = new Date();
    	        int	batch_no = today.hashCode();
					    System.out.println("DEBUG1");
					    errMsg = errMsg + UpdateF01.doParserReport_F01(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
					    System.out.println("DEBUG2");
					    if(errMsg.startsWith("U")){//檢核成功
					      errMsg = "U:相關資料寫入資料庫成功";
					    }
					    System.out.println("DEBUG3");

					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
					}
        		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}

		return errMsg;
	}

    //96.07.10 add insert A08/用preparestatement by 2295
    public String insertA08(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String bank_type,String Report_no,String userid,String username) throws Exception{
    	String WarnAccount_Cnt="";//警示帳戶」戶數（A）
        String LimitAccount_Cnt="";//「衍生管制帳戶」戶數（B）
  		String ErrorAccount_Cnt="";//「自行篩選有異常並已採資金流出管制措施之存款帳戶」戶數（不包括警示帳戶及衍生管制帳戶）（C）
  		String OtherAccount_Cnt="";//其他帳戶戶數（D）
  		String DepositAccount_TCnt="";
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		String errMsg="";
		int zerodata=0;
		List updateDBList = new LinkedList();//0:sql 1:data
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data
		try {
				WarnAccount_Cnt =  ( request.getParameter("WarnAccount_Cnt")==null ) ? "0" : (String)request.getParameter("WarnAccount_Cnt");
				LimitAccount_Cnt =  ( request.getParameter("LimitAccount_Cnt")==null ) ? "0" : (String)request.getParameter("LimitAccount_Cnt");
				ErrorAccount_Cnt =  ( request.getParameter("ErrorAccount_Cnt")==null ) ? "0" : (String)request.getParameter("ErrorAccount_Cnt");
				OtherAccount_Cnt =  ( request.getParameter("OtherAccount_Cnt")==null ) ? "0" : (String)request.getParameter("OtherAccount_Cnt");
				DepositAccount_TCnt =  ( request.getParameter("DepositAccount_TCnt")==null ) ? "0" : (String)request.getParameter("DepositAccount_TCnt");

				sqlCmd.append("SELECT * FROM "+Report_no+" WHERE m_year=? AND m_month=? AND bank_code=?");
				paramList.add(S_YEAR);				
				paramList.add(S_MONTH);
				paramList.add(bank_code);
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month");
				System.out.println(""+Report_no+".size="+data.size());

				if (data.size() != 0){
				    errMsg = errMsg + "此筆資料已存在無法新增<br>";
				}else{

				    updateDBSqlList.add("INSERT INTO "+Report_no+" VALUES(?,?,?,?,?,?,?,?)");
				    System.out.println("INSERT INTO "+Report_no+" VALUES(?,?,?,?,?,?,?,?)");


					dataList = new LinkedList();
					dataList.add(S_YEAR);//m_year
					dataList.add(S_MONTH);//m_month
					dataList.add(bank_code);//bank_code
					dataList.add(Utility.setNoCommaFormat(WarnAccount_Cnt));
					dataList.add(Utility.setNoCommaFormat(LimitAccount_Cnt));
					dataList.add(Utility.setNoCommaFormat(ErrorAccount_Cnt));
					dataList.add(Utility.setNoCommaFormat(OtherAccount_Cnt));
					dataList.add(Utility.setNoCommaFormat(DepositAccount_TCnt));
					updateDBDataList.add(dataList);


	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
	            	System.out.println("updateDBDataList add");
	            	if(DBManager.updateDB_ps(updateDBList)){
	            	   System.out.println(Report_no+"Insert ok");
					   errMsg = errMsg + "相關資料寫入資料庫成功";
					   Date today = new Date();
    	               int	batch_no = today.hashCode();
					   errMsg = errMsg + UpdateA08.doParserReport_A08(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
					   if(errMsg.startsWith("U")){//檢核成功
					      errMsg = "U:相關資料寫入資料庫成功";
					   }
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗";
					}
    	   		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}

		return errMsg;
	}


	//96.12.03 add insert A09/用preparestatement by 2295
    public String insertA09(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String bank_type,String Report_no,String userid,String username) throws Exception{
    	String over_cnt="";//剩餘件數
    	String over_amt="";//剩餘金額(A)
  		String PUSH_over_amt="";//剩餘金額-催收款(B)
  		String totalamt="";//全會放出總金額(C)
  		String PUSH_totalamt="";//全會放出總金額-催收款(D)
  		String Over_total_rate="";//佔放款總額的比率(A-B)/(C-D)
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		String errMsg="";
		int zerodata=0;
		List updateDBList = new LinkedList();//0:sql 1:data
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data
		try {
				over_cnt=   ( request.getParameter("over_cnt")==null ) ? "0" : (String)request.getParameter("over_cnt");//剩餘件數
				over_amt =  ( request.getParameter("over_amt")==null ) ? "0" : (String)request.getParameter("over_amt");//剩餘金額(A)
				PUSH_over_amt =  ( request.getParameter("PUSH_over_amt")==null ) ? "0" : (String)request.getParameter("PUSH_over_amt");//剩餘金額-催收款(B)
				totalamt =  ( request.getParameter("totalamt")==null ) ? "0" : (String)request.getParameter("totalamt");//全會放出總金額(C)
				PUSH_totalamt =  ( request.getParameter("PUSH_totalamt")==null ) ? "0" : (String)request.getParameter("PUSH_totalamt");//全會放出總金額-催收款(D)
				Over_total_rate =  ( request.getParameter("Over_total_rate")==null ) ? "0" : (String)request.getParameter("Over_total_rate");//佔放款總額的比率(A-B)/(C-D)

				System.out.println("over_cnt="+over_cnt);
				System.out.println("over_amt="+over_amt);
				System.out.println("PUSH_over_amt="+PUSH_over_amt);
				System.out.println("totalamt="+totalamt);
				System.out.println("PUSH_totalamt="+PUSH_totalamt);
				System.out.println("Over_total_rate="+Over_total_rate);
				sqlCmd.append("SELECT * FROM "+Report_no+" WHERE m_year=? AND m_month=? AND bank_code=?");
				paramList.add(S_YEAR);				
				paramList.add(S_MONTH);
				paramList.add(bank_code);
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month");
				System.out.println(""+Report_no+".size="+data.size());

				if (data.size() != 0){
				    errMsg = errMsg + "此筆資料已存在無法新增<br>";
				}else{

				    updateDBSqlList.add("INSERT INTO "+Report_no+" VALUES(?,?,?,?,?,?,?,?,?)");
				    System.out.println("INSERT INTO "+Report_no+" VALUES(?,?,?,?,?,?,?,?,?)");


					dataList = new LinkedList();
					dataList.add(S_YEAR);//m_year
					dataList.add(S_MONTH);//m_month
					dataList.add(bank_code);//bank_code
					dataList.add(Utility.setNoCommaFormat(over_cnt));
					dataList.add(Utility.setNoCommaFormat(over_amt));
					dataList.add(Utility.setNoCommaFormat(PUSH_over_amt));
					dataList.add(Utility.setNoCommaFormat(totalamt));
					dataList.add(Utility.setNoCommaFormat(PUSH_totalamt));
					//97.05.07去除小數點.
					float tmp = Float.parseFloat(Utility.setNoCommaFormat(Over_total_rate));
			        tmp = tmp * 100;
			        Over_total_rate = String.valueOf(tmp);

					dataList.add(Over_total_rate);
					updateDBDataList.add(dataList);


	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
	            	System.out.println("updateDBDataList add");
	            	if(DBManager.updateDB_ps(updateDBList)){
	            	   System.out.println(Report_no+"Insert ok");
					   errMsg = errMsg + "相關資料寫入資料庫成功";
					   Date today = new Date();
    	               int	batch_no = today.hashCode();
					   errMsg = errMsg + UpdateA09.doParserReport_A09(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
					   if(errMsg.startsWith("U")){//檢核成功
					      errMsg = "U:相關資料寫入資料庫成功";
					   }
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗";
					}
    	   		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}

		return errMsg;
	}


	//97.06.13 add insert A10/用preparestatement by 2295
    public String insertA10(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String bank_type,String Report_no,String userid,String username) throws Exception{
    	String loan1_amt="";//放款-第一類
    	String loan2_amt="";//放款-第二類
    	String loan3_amt="";//放款-第三類
    	String loan4_amt="";//放款-第四類
    	String invest1_amt="";//投資-第一類
    	String invest2_amt="";//投資-第二類
    	String invest3_amt="";//投資-第三類
    	String invest4_amt="";//投資-第四類
    	String other1_amt="";//其他-第一類
    	String other2_amt="";//其他-第二類
    	String other3_amt="";//其他-第三類
    	String other4_amt="";//其他-第四類
    	String loan1_baddebt="";//放款-帳列備抵呆帳-第一類
    	String loan2_baddebt="";//放款-帳列備抵呆帳-第二類
    	String loan3_baddebt="";//放款-帳列備抵呆帳-第三類
    	String loan4_baddebt="";//放款-帳列備抵呆帳-第四類
    	String build1_baddebt="";//建築貸款-帳列備抵呆帳-第一類
    	String build2_baddebt="";//建築貸款-帳列備抵呆帳-第二類
    	String build3_baddebt="";//建築貸款-帳列備抵呆帳-第三類
    	String build4_baddebt="";//建築貸款-帳列備抵呆帳-第四類
    	String baddebt_flag="";//備呆等於或大於應提最低標準
    	String baddebt_noenough="";//備呆不足額
    	String baddebt_delay="";//申請展延
    	String baddebt_104="";//104年底前提足備呆
    	String baddebt_105="";//105年底前提足備呆
    	String baddebt_106="";//106年底前提足備呆
    	String baddebt_107="";//107年底前提足備呆
    	String baddebt_108="";//108年底前提足備呆
    	
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		String errMsg="";
		int zerodata=0;
		List updateDBList = new LinkedList();//0:sql 1:data
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data
		try {


			loan1_amt=   ( request.getParameter("loan1_amt")==null ) ? "0" : (String)request.getParameter("loan1_amt");//放款-第一類
			loan2_amt=   ( request.getParameter("loan2_amt")==null ) ? "0" : (String)request.getParameter("loan2_amt");//放款-第二類
			loan3_amt =  ( request.getParameter("loan3_amt")==null ) ? "0" : (String)request.getParameter("loan3_amt");//放款-第三類
			loan4_amt =  ( request.getParameter("loan4_amt")==null ) ? "0" : (String)request.getParameter("loan4_amt");//放款-第四類
			invest1_amt =  ( request.getParameter("invest1_amt")==null ) ? "0" : (String)request.getParameter("invest1_amt");//投資-第一類
			invest2_amt =  ( request.getParameter("invest2_amt")==null ) ? "0" : (String)request.getParameter("invest2_amt");//投資-第二類
			invest3_amt =  ( request.getParameter("invest3_amt")==null ) ? "0" : (String)request.getParameter("invest3_amt");//投資-第三類
			invest4_amt =  ( request.getParameter("invest4_amt")==null ) ? "0" : (String)request.getParameter("invest4_amt");//投資-第四類
			other1_amt =  ( request.getParameter("other1_amt")==null ) ? "0" : (String)request.getParameter("other1_amt");//其他-第一類
		    other2_amt =  ( request.getParameter("other2_amt")==null ) ? "0" : (String)request.getParameter("other2_amt");//其他-第二類
			other3_amt =  ( request.getParameter("other3_amt")==null ) ? "0" : (String)request.getParameter("other3_amt");//其他-第三類
			other4_amt =  ( request.getParameter("other4_amt")==null ) ? "0" : (String)request.getParameter("other4_amt");//其他-第四類
			loan1_baddebt=   ( request.getParameter("loan1_baddebt")==null ) ? "0" : (String)request.getParameter("loan1_baddebt");//放款-帳列備抵呆帳-第一類
			loan2_baddebt=   ( request.getParameter("loan2_baddebt")==null ) ? "0" : (String)request.getParameter("loan2_baddebt");//放款-帳列備抵呆帳-第二類
			loan3_baddebt =  ( request.getParameter("loan3_baddebt")==null ) ? "0" : (String)request.getParameter("loan3_baddebt");//放款-帳列備抵呆帳-第三類
			loan4_baddebt =  ( request.getParameter("loan4_baddebt")==null ) ? "0" : (String)request.getParameter("loan4_baddebt");//放款-帳列備抵呆帳-第四類
			build1_baddebt=   ( request.getParameter("build1_baddebt")==null ) ? "0" : (String)request.getParameter("build1_baddebt");//建築貸款-帳列備抵呆帳-第一類
			build2_baddebt=   ( request.getParameter("build2_baddebt")==null ) ? "0" : (String)request.getParameter("build2_baddebt");//建築貸款-帳列備抵呆帳-第二類
			build3_baddebt =  ( request.getParameter("build3_baddebt")==null ) ? "0" : (String)request.getParameter("build3_baddebt");//建築貸款-帳列備抵呆帳-第三類
			build4_baddebt =  ( request.getParameter("build4_baddebt")==null ) ? "0" : (String)request.getParameter("build4_baddebt");//建築貸款-帳列備抵呆帳-第四類
			baddebt_flag = Utility.getTrimString(request.getParameter("baddebt_flag"));//備呆等於或大於應提最低標準
	  		//if("N".equals(baddebt_flag)){
	  			baddebt_noenough = ( request.getParameter("baddebt_noenough")==null || ((String)request.getParameter("baddebt_noenough")).equals("")) ? "0" : Utility.setNoCommaFormat((String)request.getParameter("baddebt_noenough"));//建築貸款-帳列備抵呆帳-第四類;//備呆不足額
		  		baddebt_delay = Utility.getTrimString(request.getParameter("baddebt_delay"));//申請展延
		  		baddebt_104 = ( request.getParameter("baddebt_104")==null  || ((String)request.getParameter("baddebt_104")).equals("")) ? "0" : Utility.setNoCommaFormat((String)request.getParameter("baddebt_104"));//104年底前提足備呆
			  	baddebt_105 = ( request.getParameter("baddebt_105")==null  || ((String)request.getParameter("baddebt_105")).equals("")) ? "0" : Utility.setNoCommaFormat((String)request.getParameter("baddebt_105"));//105年底前提足備呆
			  	baddebt_106 = ( request.getParameter("baddebt_106")==null  || ((String)request.getParameter("baddebt_106")).equals("")) ? "0" : Utility.setNoCommaFormat((String)request.getParameter("baddebt_106"));//106年底前提足備呆
			  	baddebt_107 = ( request.getParameter("baddebt_107")==null  || ((String)request.getParameter("baddebt_107")).equals("")) ? "0" : Utility.setNoCommaFormat((String)request.getParameter("baddebt_107"));//107年底前提足備呆
			  	baddebt_108 = ( request.getParameter("baddebt_108")==null  || ((String)request.getParameter("baddebt_108")).equals("")) ? "0" : Utility.setNoCommaFormat((String)request.getParameter("baddebt_108"));//108年底前提足備呆
	  		//}
	  		
		  	
	  		System.out.println("loan1_amt="+loan1_amt);
			System.out.println("loan2_amt="+loan2_amt);
			System.out.println("loan3_amt="+loan3_amt);
			System.out.println("loan4_amt="+loan4_amt);
			System.out.println("invest1_amt="+invest1_amt);
			System.out.println("invest2_amt="+invest2_amt);
			System.out.println("invest3_amt="+invest3_amt);
			System.out.println("invest4_amt="+invest4_amt);
			System.out.println("other1_amt="+other1_amt);
			System.out.println("other2_amt="+other2_amt);
			System.out.println("other3_amt="+other3_amt);
			System.out.println("other4_amt="+other4_amt);
			System.out.println("loan1_baddebt="+loan1_baddebt);
			System.out.println("loan2_baddebt="+loan2_baddebt);
			System.out.println("loan3_baddebt="+loan3_baddebt);
			System.out.println("loan4_baddebt="+loan4_baddebt);
			System.out.println("build1_baddebt="+build1_baddebt);
			System.out.println("build2_baddebt="+build2_baddebt);
			System.out.println("build3_baddebt="+build3_baddebt);
			System.out.println("build4_baddebt="+build4_baddebt);
			System.out.println("baddebt_flag="+baddebt_flag);
			System.out.println("baddebt_noenough="+baddebt_noenough);
			System.out.println("baddebt_delay="+baddebt_delay);
			System.out.println("baddebt_104="+baddebt_104);
			System.out.println("baddebt_105="+baddebt_105);
			System.out.println("baddebt_106="+baddebt_106);
			System.out.println("baddebt_107="+baddebt_107);
			System.out.println("baddebt_108="+baddebt_108);
				sqlCmd.append("SELECT * FROM "+Report_no+" WHERE m_year=? AND m_month=? AND bank_code=?");
				paramList.add(S_YEAR);				
				paramList.add(S_MONTH);
				paramList.add(bank_code);
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month");
				System.out.println(""+Report_no+".size="+data.size());

				if (data.size() != 0){
				    errMsg = errMsg + "此筆資料已存在無法新增<br>";
				}else{

				    updateDBSqlList.add("INSERT INTO "+Report_no+" VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
					dataList = new LinkedList();
					dataList.add(S_YEAR);//m_year
					dataList.add(S_MONTH);//m_month
					dataList.add(bank_code);//bank_code
					dataList.add(Utility.setNoCommaFormat(loan1_amt));
					dataList.add(Utility.setNoCommaFormat(loan2_amt));
					dataList.add(Utility.setNoCommaFormat(loan3_amt));
					dataList.add(Utility.setNoCommaFormat(loan4_amt));
					dataList.add(Utility.setNoCommaFormat(invest1_amt));
					dataList.add(Utility.setNoCommaFormat(invest2_amt));
					dataList.add(Utility.setNoCommaFormat(invest3_amt));
					dataList.add(Utility.setNoCommaFormat(invest4_amt));
					dataList.add(Utility.setNoCommaFormat(other1_amt));
					dataList.add(Utility.setNoCommaFormat(other2_amt));
					dataList.add(Utility.setNoCommaFormat(other3_amt));
					dataList.add(Utility.setNoCommaFormat(other4_amt));
					dataList.add(Utility.setNoCommaFormat(loan1_baddebt));
					dataList.add(Utility.setNoCommaFormat(loan2_baddebt));
					dataList.add(Utility.setNoCommaFormat(loan3_baddebt));
					dataList.add(Utility.setNoCommaFormat(loan4_baddebt));
					dataList.add(Utility.setNoCommaFormat(build1_baddebt));
					dataList.add(Utility.setNoCommaFormat(build2_baddebt));
					dataList.add(Utility.setNoCommaFormat(build3_baddebt));
					dataList.add(Utility.setNoCommaFormat(build4_baddebt));
					dataList.add(baddebt_flag);
					dataList.add(baddebt_noenough);
					dataList.add(baddebt_delay);
					dataList.add(baddebt_104);
					dataList.add(baddebt_105);
					dataList.add(baddebt_106);
					dataList.add(baddebt_107);
					dataList.add(baddebt_108);
					updateDBDataList.add(dataList);


	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
	            	System.out.println("updateDBDataList add");
	            	if(DBManager.updateDB_ps(updateDBList)){
	            	   System.out.println(Report_no+"Insert ok");
					   errMsg = errMsg + "相關資料寫入資料庫成功";
					   Date today = new Date();
    	               int	batch_no = today.hashCode();
					   errMsg = errMsg + UpdateA10.doParserReport_A10(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
					   if(errMsg.startsWith("U")){//檢核成功
					      errMsg = "U:相關資料寫入資料庫成功";
					   }
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗";
					}
    	   		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}

		return errMsg;
	}
    public String insertA12(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String bank_type,String Report_no,String userid,String username) throws Exception{
    	String baddebt_Amt="";//轉銷呆帳金額-本月減少備抵呆帳B1
        String loss_Amt="";//轉銷呆帳金額-本月直接認列損失B2
  		String profit_Amt="";//存款準備率降低所增加盈餘(本月增加金額)B3
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		String errMsg="";
		int zerodata=0;
		List updateDBList = new LinkedList();//0:sql 1:data
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data
		try {
				baddebt_Amt =  ( request.getParameter("baddebt_Amt")==null ) ? "0" : (String)request.getParameter("baddebt_Amt");
				loss_Amt =  ( request.getParameter("loss_Amt")==null ) ? "0" : (String)request.getParameter("loss_Amt");
				profit_Amt =  ( request.getParameter("profit_Amt")==null ) ? "0" : (String)request.getParameter("profit_Amt");

				sqlCmd.append("SELECT * FROM "+Report_no+" WHERE m_year=? AND m_month=? AND bank_code=?");
				paramList.add(S_YEAR);				
				paramList.add(S_MONTH);
				paramList.add(bank_code);
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month");
				System.out.println(""+Report_no+".size="+data.size());

				if (data.size() != 0){
				    errMsg = errMsg + "此筆資料已存在無法新增<br>";
				}else{
	            	
				    updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List
					dataList = new LinkedList();//儲存參數的data
					sqlCmd.delete(0,sqlCmd.length());			
				    updateDBSqlList.add("INSERT INTO "+Report_no+" VALUES(?,?,?,?,?,?)");
					dataList = new LinkedList();
					dataList.add(S_YEAR);//m_year
					dataList.add(S_MONTH);//m_month
					dataList.add(bank_code);//bank_code
					dataList.add(Utility.setNoCommaFormat(baddebt_Amt));
					dataList.add(Utility.setNoCommaFormat(loss_Amt));
					dataList.add(Utility.setNoCommaFormat(profit_Amt));
					updateDBDataList.add(dataList);
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
				
	            	if(DBManager.updateDB_ps(updateDBList)){
	            	   System.out.println(Report_no+"Insert ok");
					   //errMsg = errMsg + "相關資料寫入資料庫成功";
					   Date today = new Date();
    	               int	batch_no = today.hashCode();
					   errMsg = errMsg + UpdateA12.doParserReport_A12(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
					   System.out.println("A12.errMsg="+errMsg);
					   if(errMsg.startsWith("U") || errMsg.startsWith("Z")){//檢核成功//104.02.12
					      errMsg = "U:相關資料寫入資料庫成功";
					   }
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗";
					}
    	   		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}
		return errMsg;
    }
	//95.04.19 改用preparestatement by 2295
	//95.05.18 add A99 by 2295
	//102.01.15 add A02.amt_name by 2295
    public String updateA01_A05(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String Report_no,String userid,String username) throws Exception{
    	String[]	Amt			= request.getParameterValues("amt");
		String[]	Acc_Code	= request.getParameterValues("acc_code");
		String	txt990421	= (String)request.getParameter("txt990421");
		String	txt990621	= (String)request.getParameter("txt990621");
		String	txt990422	= (String)request.getParameter("txt990422");//104.02.13
		String	txt990622	= (String)request.getParameter("txt990622");//104.02.13
		String bank_type = ( request.getAttribute("bank_type")==null ) ? "6" : (String)request.getAttribute("bank_type");

		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		String errMsg="";

		try {
				List updateDBList = new LinkedList();//0:sql 1:data
				List updateDBSqlList = new LinkedList();
				List updateDBDataList = new LinkedList();//儲存參數的List
				List dataList = new LinkedList();//儲存參數的data
				sqlCmd.append("SELECT * FROM "+Report_no+" WHERE m_year=? AND m_month=? AND bank_code=?");
				paramList.add(S_YEAR);				
				paramList.add(S_MONTH);
				paramList.add(bank_code);
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");
				System.out.println(""+Report_no+".size="+data.size());

				if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法修改<br>";
				}else{
				    //93.12.7 insert LOG===================================================
				    sqlCmd.delete(0,sqlCmd.length());
				    sqlCmd.append("INSERT INTO "+Report_no+"_LOG ");
					sqlCmd.append(" select m_year,m_month,bank_code,acc_code,");

					if(Report_no.equals("A05")){
					   sqlCmd.append(" amt,amt_name ");
					}else{
					   sqlCmd.append(" amt");
					}
					sqlCmd.append(",?,?,sysdate,'U'");
					if(Report_no.equals("A02")){
						 sqlCmd.append(" ,amt_name");
					}	
					sqlCmd.append(" from "+Report_no);
					sqlCmd.append(" WHERE m_year=? AND m_month=? AND bank_code=?");
					dataList.add(userid);	   
					dataList.add(username);	
					dataList.add(S_YEAR);	
					dataList.add(S_MONTH);	
					dataList.add(bank_code);
					updateDBDataList.add(dataList);

					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);					
				    //=========================================================================
				    
				    
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List
					dataList = new LinkedList();//儲存參數的data
					sqlCmd.delete(0,sqlCmd.length());					
				    sqlCmd.append("delete "+Report_no+" WHERE m_year=? AND m_month=? AND bank_code=?");
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					dataList.add(bank_code);
					updateDBDataList.add(dataList);	 
					
				    updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	

				    updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List					
					sqlCmd.delete(0,sqlCmd.length());					
				    if(Report_no.equals("A05") || Report_no.equals("A02")){//102.01.15 add by 2295
				       sqlCmd.append("INSERT INTO "+Report_no+" VALUES(?,?,?,?,?,?)");
				       System.out.println("INSERT INTO "+Report_no+" VALUES(?,?,?,?,?,?)");				   
				    }else{
				       sqlCmd.append("INSERT INTO "+Report_no+" VALUES(?,?,?,?,?)");
				       System.out.println("INSERT INTO "+Report_no+" VALUES(?,?,?,?,?)");
				    }
					for (int i = 0; i < Acc_Code.length; i++) {
					     dataList = new LinkedList();
						 dataList.add(S_YEAR);//m_year
						 dataList.add(S_MONTH);//m_month
						 dataList.add(bank_code);//bank_code
						 dataList.add(Acc_Code[i].trim());//acc_code
						 if(Report_no.equals("A01") || Report_no.equals("A99") ){
						   dataList.add(Utility.setNoCommaFormat(Amt[i]));//amt
		            	 }else if(Report_no.equals("A02")){//102.01.15 add by 2295
		            	 	dataList.add(Utility.setNoCommaFormat(Amt[i]));//amt
		            	 	if("990421".equals(Acc_Code[i])){
		            	 	   dataList.add(txt990421);//amt_name
		            	 	}else if("990621".equals(Acc_Code[i])){
		            		   dataList.add(txt990621);//amt_name
		            		}else if("990422".equals(Acc_Code[i])){//104.02.13
		            	 	   dataList.add(txt990422);//amt_name
		            	 	}else if("990622".equals(Acc_Code[i])){//104.02.13
		            		   dataList.add(txt990622);//amt_name      
		            		}else{
		            		   dataList.add("");//amt_name
		            		}		
		            	 }else if(Report_no.equals("A03")||Report_no.equals("A04")){		//modify by 2354 2004.12.23
							if(Acc_Code[i].substring(Acc_Code[i].length()-1).equals("P")){
							   dataList.add(Utility.setNoPercentFormat(Amt[i]));//amt
		            		}else{
		            		   dataList.add(Utility.setNoCommaFormat(Amt[i]));//amt
			            	}
						 }else if(Report_no.equals("A05")){
							if(Acc_Code[i].substring(Acc_Code[i].length()-1).equals("P")){
							   dataList.add(Utility.setNoPercentFormat(Amt[i]));//amt
							   dataList.add("");//amt_name
		            		}else if(Acc_Code[i].substring(Acc_Code[i].length()-1).equals("N")){
		            		   dataList.add("0");//amt=0
							   dataList.add(Amt[i]);//amt_name
		            		}else{
		            		   dataList.add(Utility.setNoCommaFormat(Amt[i]));//amt
							   dataList.add("");//amt_name
			            	}
			              }else if(Report_no.equals("A13")){
							if(Acc_Code[i].equals("995400") || Acc_Code[i].equals("993810")){
							   dataList.add(Utility.setNoPercentFormat(Amt[i]));//amt							 
		            		}else{
		            		   dataList.add(Utility.setNoCommaFormat(Amt[i]));//amt						   
			            	}
						 }	
						 
						 updateDBDataList.add(dataList);
						 //System.out.println("add db data i="+i);
	            	}

	            	updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	
	            	
	            	if(DBManager.updateDB_ps(updateDBList)){
						errMsg = errMsg + "相關資料寫入資料庫成功";
						//93.12.01 add 2295
						if(Report_no.equals("A01")){
					    	Date today = new Date();
    	                	int	batch_no = today.hashCode();
					    	errMsg = errMsg + UpdateA01.doParserReport_A01(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
						}else if(Report_no.equals("A02")){//add by winnin 2004.12.10
					    	Date today = new Date();
    	                	int	batch_no = today.hashCode();
					    	errMsg = errMsg + UpdateA02.doParserReport_A02(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
						}else if(Report_no.equals("A03")){//add by winnin 2004.12.23
					    	Date today = new Date();
    	                	int	batch_no = today.hashCode();
					    	errMsg = errMsg + UpdateA03.doParserReport_A03(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
						}else if(Report_no.equals("A04")){//add by winnin 2004.12.10
					    	Date today = new Date();
    	                	int	batch_no = today.hashCode();
					    	errMsg = errMsg + UpdateA04.doParserReport_A04(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
						}else if(Report_no.equals("A05")){//add by winnin 2004.12.10
					    	Date today = new Date();
    	                	int	batch_no = today.hashCode();
					    	errMsg = errMsg + UpdateA05.doParserReport_A05(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
					    }else if(Report_no.equals("A99")){//add by 2295 2006.05.18
					    	Date today = new Date();
    	                	int	batch_no = today.hashCode();
					    	errMsg = errMsg + UpdateA99.doParserReport_A99(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
					    	if(errMsg.startsWith("U")){//檢核成功
					           errMsg = "U:相關資料寫入資料庫成功";
					        }
					     }else if(Report_no.equals("A13")){//add by 2295 104.10.12
					    	Date today = new Date();
    	                	int	batch_no = today.hashCode();
					    	errMsg = errMsg + UpdateA13.doParserReport_A13(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
					    	if(errMsg.startsWith("U")){//檢核成功
					           errMsg = "U:相關資料寫入資料庫成功";
					        }    
					    }
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗";
					}
        		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}

		return errMsg;
	}

	public String updateA06(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String Report_no,String userid,String username) throws Exception{
    	String[]	Amt_3month	  = request.getParameterValues("amt_3month");
    	String[]	Amt_6month	  = request.getParameterValues("amt_6month");
    	String[]	Amt_1year	  = request.getParameterValues("amt_1year");
    	String[]	Amt_2year	  = request.getParameterValues("amt_2year");
    	String[]	Amt_over2year = request.getParameterValues("amt_over2year");
    	String[]	Amt_total	  = request.getParameterValues("amt_total");
		String[]	Acc_Code	= request.getParameterValues("acc_code");
		String bank_type = ( request.getAttribute("bank_type")==null ) ? "6" : (String)request.getAttribute("bank_type");

		String errMsg="";
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		List updateDBList = new ArrayList();//0:sql 1:data
		List updateDBSqlList = new ArrayList();
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data
		try {
				sqlCmd.append("SELECT * FROM "+Report_no+" WHERE m_year= ? AND m_month= ? AND bank_code=?");
				paramList.add(S_YEAR);
				paramList.add(S_MONTH);
				paramList.add(bank_code);
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");
				System.out.println(""+Report_no+".size="+data.size());

				if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法修改<br>";
				}else{
				    //insert LOG===================================================
				    sqlCmd.delete(0,sqlCmd.length());	
				    sqlCmd.append(" INSERT INTO "+Report_no+"_LOG ");
					sqlCmd.append(" select m_year,m_month,bank_code,acc_code,");
					sqlCmd.append(" amt_3month,amt_6month,amt_1year,amt_2year,amt_over2year,amt_total");
					sqlCmd.append(",?,?,sysdate,'U'");
					sqlCmd.append(" from "+Report_no);
					sqlCmd.append(" WHERE m_year=? AND m_month=? AND bank_code=?");
					dataList.add(userid);
					dataList.add(username);
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					dataList.add(bank_code);
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	

					
				    //=========================================================================
				    updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					dataList =  new ArrayList();//儲存參數的data				
					sqlCmd.delete(0,sqlCmd.length());			
				    sqlCmd.append("delete "+Report_no+" WHERE m_year=? AND m_month=? AND bank_code=?");
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					dataList.add(bank_code);
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	

										 		 
				    
				    updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List					
					sqlCmd.delete(0,sqlCmd.length());		
				    
				    sqlCmd.append("INSERT INTO "+Report_no+" VALUES (?,?,?,?,?,?,?,?,?,?)");
					for (int i = 0; i < Acc_Code.length; i++) {
						dataList =  new ArrayList();//儲存參數的data
						dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						dataList.add(bank_code);						
						dataList.add(Acc_Code[i].trim());
		            	dataList.add(Utility.setNoCommaFormat(Amt_3month[i]));
		            	dataList.add(Utility.setNoCommaFormat(Amt_6month[i]));
		            	dataList.add(Utility.setNoCommaFormat(Amt_1year[i]));
		            	dataList.add(Utility.setNoCommaFormat(Amt_2year[i]));
		            	dataList.add(Utility.setNoCommaFormat(Amt_over2year[i]));
		            	dataList.add(Utility.setNoCommaFormat(Amt_total[i]));
	            		updateDBDataList.add(dataList);  
	            	}
	            	updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	
	            
					if(DBManager.updateDB_ps(updateDBList)){
						//errMsg = errMsg + "相關資料寫入資料庫成功";
					    Date today = new Date();
    	                int	batch_no = today.hashCode();
					    errMsg = errMsg + UpdateA06.doParserReport_A06(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
					    if(errMsg.startsWith("U")){//檢核成功
					      errMsg = "U:相關資料寫入資料庫成功";
					    }
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗";
					}
        		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}

		return errMsg;
	}

	//96.07.10 add A08/改用preparestatement by 2295
    public String updateA08(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String Report_no,String userid,String username) throws Exception{
    	String WarnAccount_Cnt="";//警示帳戶」戶數（A）
        String LimitAccount_Cnt="";//「衍生管制帳戶」戶數（B）
  		String ErrorAccount_Cnt="";//「自行篩選有異常並已採資金流出管制措施之存款帳戶」戶數（不包括警示帳戶及衍生管制帳戶）（C）
  		String OtherAccount_Cnt="";//其他帳戶戶數（D）
  		String DepositAccount_TCnt="";
  		String bank_type = ( request.getAttribute("bank_type")==null ) ? "6" : (String)request.getAttribute("bank_type");
		String errMsg="";
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();		
		List updateDBList = new LinkedList();//0:sql 1:data
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data

		try {
				WarnAccount_Cnt =  ( request.getParameter("WarnAccount_Cnt")==null ) ? "0" : (String)request.getParameter("WarnAccount_Cnt");
				LimitAccount_Cnt =  ( request.getParameter("LimitAccount_Cnt")==null ) ? "0" : (String)request.getParameter("LimitAccount_Cnt");
				ErrorAccount_Cnt =  ( request.getParameter("ErrorAccount_Cnt")==null ) ? "0" : (String)request.getParameter("ErrorAccount_Cnt");
				OtherAccount_Cnt =  ( request.getParameter("OtherAccount_Cnt")==null ) ? "0" : (String)request.getParameter("OtherAccount_Cnt");
				DepositAccount_TCnt =  ( request.getParameter("DepositAccount_TCnt")==null ) ? "0" : (String)request.getParameter("DepositAccount_TCnt");

				sqlCmd.append("SELECT * FROM "+Report_no+" WHERE m_year=? AND m_month=? AND bank_code=?");
				paramList.add(S_YEAR);
				paramList.add(S_MONTH);
				paramList.add(bank_code);
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month");
				System.out.println(""+Report_no+".size="+data.size());

				if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法修改<br>";
				}else{
				    //93.12.7 insert LOG===================================================
				    sqlCmd.delete(0,sqlCmd.length());	
				    sqlCmd.append(" INSERT INTO "+Report_no+"_LOG ");
					sqlCmd.append(" select m_year,m_month,bank_code,warnAccount_cnt,limitAccount_cnt,erroraccount_cnt,otheraccount_cnt,depositaccount_tcnt");
					sqlCmd.append(",?,?,sysdate,'U'");
					sqlCmd.append(" from "+Report_no);
					sqlCmd.append(" WHERE m_year=? AND m_month=? AND bank_code=?");
					dataList.add(userid);
					dataList.add(username);
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					dataList.add(bank_code);
					updateDBDataList.add(dataList); 	
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	
	            						   
					
				    //=========================================================================
				    updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					dataList =  new ArrayList();//儲存參數的data						
					sqlCmd.delete(0,sqlCmd.length());
				    sqlCmd.append("delete "+Report_no+" WHERE m_year=? AND m_month=? AND bank_code=?");
					dataList.add(S_YEAR); 		 
					dataList.add(S_MONTH);
					dataList.add(bank_code);
					updateDBDataList.add(dataList); 	
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	

				    updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List						
					sqlCmd.delete(0,sqlCmd.length());
				    sqlCmd.append("INSERT INTO "+Report_no+" VALUES(?,?,?,?,?,?,?,?)");
				    //System.out.println("INSERT INTO "+Report_no+" VALUES(?,?,?,?,?,?,?,?)");

					dataList = new LinkedList();//儲存參數的data	
					dataList.add(S_YEAR);//m_year
					dataList.add(S_MONTH);//m_month
					dataList.add(bank_code);//bank_code
					dataList.add(Utility.setNoCommaFormat(WarnAccount_Cnt));
					dataList.add(Utility.setNoCommaFormat(LimitAccount_Cnt));
					dataList.add(Utility.setNoCommaFormat(ErrorAccount_Cnt));
					dataList.add(Utility.setNoCommaFormat(OtherAccount_Cnt));
					dataList.add(Utility.setNoCommaFormat(DepositAccount_TCnt));
					updateDBDataList.add(dataList);

	            	updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	
	            	System.out.println("updateDBDataList add");
	            	if(DBManager.updateDB_ps(updateDBList)){
						errMsg = errMsg + "相關資料寫入資料庫成功";
					   	Date today = new Date();
    	               	int	batch_no = today.hashCode();
					   	errMsg = errMsg + UpdateA08.doParserReport_A08(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
					   	if(errMsg.startsWith("U")){//檢核成功
					       errMsg = "U:相關資料寫入資料庫成功";
					    }
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗";
					}
        		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}

		return errMsg;
	}

	//96.12.03 add A09/改用preparestatement by 2295
    public String updateA09(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String Report_no,String userid,String username) throws Exception{
    	String over_cnt="";//剩餘件數
    	String over_amt="";//剩餘金額(A)
  		String PUSH_over_amt="";//剩餘金額-催收款(B)
  		String totalamt="";//全會放出總金額(C)
  		String PUSH_totalamt="";//全會放出總金額-催收款(D)
  		String Over_total_rate="";//佔放款總額的比率(A-B)/(C-D)
  		String bank_type = ( request.getAttribute("bank_type")==null ) ? "6" : (String)request.getAttribute("bank_type");

		String errMsg="";
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		List updateDBList = new LinkedList();//0:sql 1:data
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data

		try {
				over_cnt =  ( request.getParameter("over_cnt")==null ) ? "0" : (String)request.getParameter("over_cnt");//剩餘件數
				over_amt =  ( request.getParameter("over_amt")==null ) ? "0" : (String)request.getParameter("over_amt");//剩餘金額(A)
				PUSH_over_amt =  ( request.getParameter("PUSH_over_amt")==null ) ? "0" : (String)request.getParameter("PUSH_over_amt");//剩餘金額-催收款(B)
				totalamt =  ( request.getParameter("totalamt")==null ) ? "0" : (String)request.getParameter("totalamt");//全會放出總金額(C)
				PUSH_totalamt =  ( request.getParameter("PUSH_totalamt")==null ) ? "0" : (String)request.getParameter("PUSH_totalamt");//全會放出總金額-催收款(D)
				Over_total_rate =  ( request.getParameter("Over_total_rate")==null ) ? "0" : (String)request.getParameter("Over_total_rate");//佔放款總額的比率(A-B)/(C-D)

				sqlCmd.append("SELECT * FROM "+Report_no+" WHERE m_year=? AND m_month=? AND bank_code=?");
				paramList.add(S_YEAR);
				paramList.add(S_MONTH);
				paramList.add(bank_code);
				
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month");
				System.out.println(""+Report_no+".size="+data.size());
				System.out.println("over_cnt="+over_cnt);
				System.out.println("over_amt="+over_amt);
				System.out.println("PUSH_over_amt="+PUSH_over_amt);
				System.out.println("totalamt="+totalamt);
				System.out.println("PUSH_totalamt="+PUSH_totalamt);
				System.out.println("Over_total_rate="+Over_total_rate);
				if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法修改<br>";
				}else{
				    //93.12.7 insert LOG===================================================
				    sqlCmd.delete(0,sqlCmd.length());
				    sqlCmd.append(" INSERT INTO "+Report_no+"_LOG ");
					sqlCmd.append(" select m_year,m_month,bank_code,over_cnt,over_amt,PUSH_over_amt,totalamt,PUSH_totalamt,Over_total_rate");
					sqlCmd.append(",?,?,sysdate,'U' from "+Report_no);
					sqlCmd.append(" WHERE m_year=? AND m_month=? AND bank_code=?");
					dataList.add(userid);
					dataList.add(username);
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					dataList.add(bank_code);
					updateDBDataList.add(dataList);
	            	updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
	            		
				    //=========================================================================
				    updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List						
					dataList = new LinkedList();
					sqlCmd.delete(0,sqlCmd.length());
				    sqlCmd.append("delete "+Report_no+" WHERE m_year=? AND m_month=? AND bank_code=?");
					dataList.add(S_YEAR);					 		 
					dataList.add(S_MONTH);
					dataList.add(bank_code);
					updateDBDataList.add(dataList);
	            	updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);

				    updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List						
					sqlCmd.delete(0,sqlCmd.length());
				    sqlCmd.append("INSERT INTO "+Report_no+" VALUES(?,?,?,?,?,?,?,?,?)");
				    //System.out.println("INSERT INTO "+Report_no+" VALUES(?,?,?,?,?,?,?,?,?)");

					dataList = new LinkedList();
					dataList.add(S_YEAR);//m_year
					dataList.add(S_MONTH);//m_month
					dataList.add(bank_code);//bank_code
					dataList.add(Utility.setNoCommaFormat(over_cnt));
					dataList.add(Utility.setNoCommaFormat(over_amt));
					dataList.add(Utility.setNoCommaFormat(PUSH_over_amt));
					dataList.add(Utility.setNoCommaFormat(totalamt));
					dataList.add(Utility.setNoCommaFormat(PUSH_totalamt));

					//97.05.07去除小數點.
					float tmp = Float.parseFloat(Utility.setNoCommaFormat(Over_total_rate));
			        tmp = tmp * 100;
			        Over_total_rate = String.valueOf(tmp);
					dataList.add(Over_total_rate);

					updateDBDataList.add(dataList);					
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
					
	            	
	            	System.out.println("updateDBDataList add");
	            	if(DBManager.updateDB_ps(updateDBList)){
						errMsg = errMsg + "相關資料寫入資料庫成功";
					   	Date today = new Date();
    	               	int	batch_no = today.hashCode();
					   	errMsg = errMsg + UpdateA09.doParserReport_A09(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
					   	if(errMsg.startsWith("U")){//檢核成功
					       errMsg = "U:相關資料寫入資料庫成功";
					    }
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗";
					}
        		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}

		return errMsg;
	}

	//97.06.13 add A10/改用preparestatement by 2295
    public String updateA10(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String Report_no,String userid,String username) throws Exception{
    	String loan1_amt="";//放款-第一類
		String loan2_amt="";//放款-第二類
  		String loan3_amt="";//放款-第三類
  		String loan4_amt="";//放款-第四類
  		String invest1_amt="";//投資-第一類
  		String invest2_amt="";//投資-第二類
  		String invest3_amt="";//投資-第三類
  		String invest4_amt="";//投資-第四類
  		String other1_amt="";//其他-第一類
  		String other2_amt="";//其他-第二類
  		String other3_amt="";//其他-第三類
  		String other4_amt="";//其他-第四類
  		String loan1_baddebt="";//放款-帳列備抵呆帳-第一類
  		String loan2_baddebt="";//放款-帳列備抵呆帳-第二類
  		String loan3_baddebt="";//放款-帳列備抵呆帳-第三類
  		String loan4_baddebt="";//放款-帳列備抵呆帳-第四類
  		String build1_baddebt="";//建築貸款-帳列備抵呆帳-第一類
  		String build2_baddebt="";//建築貸款-帳列備抵呆帳-第二類
  		String build3_baddebt="";//建築貸款-帳列備抵呆帳-第三類
  		String build4_baddebt="";//建築貸款-帳列備抵呆帳-第四類
  		String baddebt_flag="";//備呆等於或大於應提最低標準
  		String baddebt_noenough="";//備呆不足額
  		String baddebt_delay="";//申請展延
  		String baddebt_104="";//104年底前提足備呆
  		String baddebt_105="";//105年底前提足備呆
  		String baddebt_106="";//106年底前提足備呆
  		String baddebt_107="";//107年底前提足備呆
  		String baddebt_108="";//108年底前提足備呆
  		String bank_type = ( request.getAttribute("bank_type")==null ) ? "6" : (String)request.getAttribute("bank_type");

		String errMsg="";
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		List updateDBList = new LinkedList();//0:sql 1:data
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data
		System.out.println("updateA10 begin");
		try {
				loan1_amt=   ( request.getParameter("loan1_amt")==null ) ? "0" : (String)request.getParameter("loan1_amt");//放款-第一類
				loan2_amt=   ( request.getParameter("loan2_amt")==null ) ? "0" : (String)request.getParameter("loan2_amt");//放款-第二類
				loan3_amt =  ( request.getParameter("loan3_amt")==null ) ? "0" : (String)request.getParameter("loan3_amt");//放款-第三類
				loan4_amt =  ( request.getParameter("loan4_amt")==null ) ? "0" : (String)request.getParameter("loan4_amt");//放款-第四類
				invest1_amt =  ( request.getParameter("invest1_amt")==null ) ? "0" : (String)request.getParameter("invest1_amt");//投資-第一類
				invest2_amt =  ( request.getParameter("invest2_amt")==null ) ? "0" : (String)request.getParameter("invest2_amt");//投資-第二類
				invest3_amt =  ( request.getParameter("invest3_amt")==null ) ? "0" : (String)request.getParameter("invest3_amt");//投資-第三類
				invest4_amt =  ( request.getParameter("invest4_amt")==null ) ? "0" : (String)request.getParameter("invest4_amt");//投資-第四類
				other1_amt =  ( request.getParameter("other1_amt")==null ) ? "0" : (String)request.getParameter("other1_amt");//其他-第一類
			    other2_amt =  ( request.getParameter("other2_amt")==null ) ? "0" : (String)request.getParameter("other2_amt");//其他-第二類
				other3_amt =  ( request.getParameter("other3_amt")==null ) ? "0" : (String)request.getParameter("other3_amt");//其他-第三類
				other4_amt =  ( request.getParameter("other4_amt")==null ) ? "0" : (String)request.getParameter("other4_amt");//其他-第四類
				loan1_baddebt=   ( request.getParameter("loan1_baddebt")==null ) ? "0" : (String)request.getParameter("loan1_baddebt");//放款-帳列備抵呆帳-第一類
				loan2_baddebt=   ( request.getParameter("loan2_baddebt")==null ) ? "0" : (String)request.getParameter("loan2_baddebt");//放款-帳列備抵呆帳-第二類
				loan3_baddebt =  ( request.getParameter("loan3_baddebt")==null ) ? "0" : (String)request.getParameter("loan3_baddebt");//放款-帳列備抵呆帳-第三類
				loan4_baddebt =  ( request.getParameter("loan4_baddebt")==null ) ? "0" : (String)request.getParameter("loan4_baddebt");//放款-帳列備抵呆帳-第四類
				build1_baddebt=   ( request.getParameter("build1_baddebt")==null ) ? "0" : (String)request.getParameter("build1_baddebt");//建築貸款-帳列備抵呆帳-第一類
				build2_baddebt=   ( request.getParameter("build2_baddebt")==null ) ? "0" : (String)request.getParameter("build2_baddebt");//建築貸款-帳列備抵呆帳-第二類
				build3_baddebt =  ( request.getParameter("build3_baddebt")==null ) ? "0" : (String)request.getParameter("build3_baddebt");//建築貸款-帳列備抵呆帳-第三類
				build4_baddebt =  ( request.getParameter("build4_baddebt")==null ) ? "0" : (String)request.getParameter("build4_baddebt");//建築貸款-帳列備抵呆帳-第四類
				baddebt_flag = Utility.getTrimString(request.getParameter("baddebt_flag"));//備呆等於或大於應提最低標準
				//if("N".equals(baddebt_flag)){
		  			baddebt_noenough = ( request.getParameter("baddebt_noenough")==null ) ? "0" : Utility.setNoCommaFormat((String)request.getParameter("baddebt_noenough"));//建築貸款-帳列備抵呆帳-第四類;//備呆不足額
			  		baddebt_delay = Utility.getTrimString(request.getParameter("baddebt_delay"));//申請展延
					baddebt_104 = ( request.getParameter("baddebt_104")==null || Utility.setNoCommaFormat((String)request.getParameter("baddebt_104"))=="") ? "0" : Utility.setNoCommaFormat((String)request.getParameter("baddebt_104"));//104年底前提足備呆
			  		baddebt_105 = ( request.getParameter("baddebt_105")==null || Utility.setNoCommaFormat((String)request.getParameter("baddebt_105"))=="") ? "0" : Utility.setNoCommaFormat((String)request.getParameter("baddebt_105"));//105年底前提足備呆
			  		baddebt_106 = ( request.getParameter("baddebt_106")==null || Utility.setNoCommaFormat((String)request.getParameter("baddebt_106"))=="") ? "0" : Utility.setNoCommaFormat((String)request.getParameter("baddebt_106"));//106年底前提足備呆
			  		baddebt_107 = ( request.getParameter("baddebt_107")==null || Utility.setNoCommaFormat((String)request.getParameter("baddebt_107"))=="") ? "0" : Utility.setNoCommaFormat((String)request.getParameter("baddebt_107"));//107年底前提足備呆
			  		baddebt_108 = ( request.getParameter("baddebt_108")==null || Utility.setNoCommaFormat((String)request.getParameter("baddebt_108"))=="") ? "0" : Utility.setNoCommaFormat((String)request.getParameter("baddebt_108"));//108年底前提足備呆
		  		//}
				
				sqlCmd.append("SELECT * FROM "+Report_no+" WHERE m_year= ? AND m_month=? AND bank_code=?");
				paramList.add(S_YEAR);
				paramList.add(S_MONTH);
				paramList.add(bank_code);
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,loan1_amt,loan2_amt,loan3_amt,loan4_amt,invest1_amt,invest2_amt,invest3_amt,invest4_amt,other1_amt,other2_amt,other3_amt,other4_amt"
			    																  +",loan1_baddebt,loan2_baddebt,loan3_baddebt,loan4_baddebt,build1_baddebt,build2_baddebt,build3_baddebt,build4_baddebt"
			    																  +",baddebt_flag,baddebt_noenough,baddebt_delay,baddebt_104,baddebt_105,baddebt_106,baddebt_107,baddebt_108");
				System.out.println(""+Report_no+".size="+data.size());
				System.out.println("loan1_amt="+loan1_amt);
				System.out.println("loan2_amt="+loan2_amt);
				System.out.println("loan3_amt="+loan3_amt);
				System.out.println("loan4_amt="+loan4_amt);
				System.out.println("invest1_amt="+invest1_amt);
				System.out.println("invest2_amt="+invest2_amt);
				System.out.println("invest3_amt="+invest3_amt);
				System.out.println("invest4_amt="+invest4_amt);
				System.out.println("other1_amt="+other1_amt);
				System.out.println("other2_amt="+other2_amt);
				System.out.println("other3_amt="+other3_amt);
				System.out.println("other4_amt="+other4_amt);
				System.out.println("loan1_baddebt="+loan1_baddebt);
				System.out.println("loan2_baddebt="+loan2_baddebt);
				System.out.println("loan3_baddebt="+loan3_baddebt);
				System.out.println("loan4_baddebt="+loan4_baddebt);
				System.out.println("build1_baddebt="+build1_baddebt);
				System.out.println("build2_baddebt="+build2_baddebt);
				System.out.println("build3_baddebt="+build3_baddebt);
				System.out.println("build4_baddebt="+build4_baddebt);
				System.out.println("baddebt_flag="+baddebt_flag);
				System.out.println("baddebt_noenough="+baddebt_noenough);
				System.out.println("baddebt_delay="+baddebt_delay);
				System.out.println("baddebt_104="+baddebt_104);
				System.out.println("baddebt_105="+baddebt_105);
				System.out.println("baddebt_106="+baddebt_106);
				System.out.println("baddebt_107="+baddebt_107);
				System.out.println("baddebt_108="+baddebt_108);
				if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法修改<br>";
				}else{
				    //93.12.7 insert LOG===================================================
				    sqlCmd.delete(0,sqlCmd.length());
				    sqlCmd.append(" INSERT INTO "+Report_no+"_LOG ");
					sqlCmd.append(" select m_year,m_month,bank_code,loan1_amt,loan2_amt,loan3_amt,loan4_amt,invest1_amt,invest2_amt,invest3_amt,invest4_amt,other1_amt,other2_amt,other3_amt,other4_amt");
					sqlCmd.append(" ,loan1_baddebt,loan2_baddebt,loan3_baddebt,loan4_baddebt,build1_baddebt,build2_baddebt,build3_baddebt,build4_baddebt");
					sqlCmd.append(" ,baddebt_flag,baddebt_noenough,baddebt_delay,baddebt_104,baddebt_105,baddebt_106,baddebt_107,baddebt_108");
					sqlCmd.append(",?,?,sysdate,'U' from "+Report_no);
					sqlCmd.append(" WHERE m_year=? AND m_month=? AND bank_code=?");
					dataList.add(userid);
					dataList.add(username);
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					dataList.add(bank_code);
					updateDBDataList.add(dataList);					
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
				    //=========================================================================
				    updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List						
					sqlCmd.delete(0,sqlCmd.length());
				    sqlCmd.append("delete "+Report_no+" WHERE m_year=? AND m_month=? AND bank_code=?");
					dataList = new LinkedList();//儲存參數的data 		 
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					dataList.add(bank_code);					 		 
					updateDBDataList.add(dataList);					
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);

				    updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List						
					sqlCmd.delete(0,sqlCmd.length());
				    sqlCmd.append("INSERT INTO "+Report_no+" VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?");
				    sqlCmd.append(" ,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
					dataList = new LinkedList();
					dataList.add(S_YEAR);//m_year
					dataList.add(S_MONTH);//m_month
					dataList.add(bank_code);//bank_code
					dataList.add(Utility.setNoCommaFormat(loan1_amt));
					dataList.add(Utility.setNoCommaFormat(loan2_amt));
					dataList.add(Utility.setNoCommaFormat(loan3_amt));
					dataList.add(Utility.setNoCommaFormat(loan4_amt));
					dataList.add(Utility.setNoCommaFormat(invest1_amt));
					dataList.add(Utility.setNoCommaFormat(invest2_amt));
					dataList.add(Utility.setNoCommaFormat(invest3_amt));
					dataList.add(Utility.setNoCommaFormat(invest4_amt));
					dataList.add(Utility.setNoCommaFormat(other1_amt));
					dataList.add(Utility.setNoCommaFormat(other2_amt));
					dataList.add(Utility.setNoCommaFormat(other3_amt));
					dataList.add(Utility.setNoCommaFormat(other4_amt));
					dataList.add(Utility.setNoCommaFormat(loan1_baddebt));
					dataList.add(Utility.setNoCommaFormat(loan2_baddebt));
					dataList.add(Utility.setNoCommaFormat(loan3_baddebt));
					dataList.add(Utility.setNoCommaFormat(loan4_baddebt));
					dataList.add(Utility.setNoCommaFormat(build1_baddebt));
					dataList.add(Utility.setNoCommaFormat(build2_baddebt));
					dataList.add(Utility.setNoCommaFormat(build3_baddebt));
					dataList.add(Utility.setNoCommaFormat(build4_baddebt));
					dataList.add(baddebt_flag);
					dataList.add(baddebt_noenough);
					dataList.add(baddebt_delay);
					dataList.add(baddebt_104);
					dataList.add(baddebt_105);
					dataList.add(baddebt_106);
					dataList.add(baddebt_107);
					dataList.add(baddebt_108);
					updateDBDataList.add(dataList);
					
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
					
	            	
	            	System.out.println("updateDBDataList add");
	            	if(DBManager.updateDB_ps(updateDBList)){
						errMsg = errMsg + "相關資料寫入資料庫成功";
					   	Date today = new Date();
    	               	int	batch_no = today.hashCode();
					   	errMsg = errMsg + UpdateA10.doParserReport_A10(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
					   	if(errMsg.startsWith("U")){//檢核成功
					       errMsg = "U:相關資料寫入資料庫成功";
					    }
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗";
					}
        		}
		}catch (Exception e){
			System.out.println(e+":"+e.getMessage());
			errMsg = errMsg + "相關資料寫入資料庫失敗";
		}

		return errMsg;
	}
    public String updateA12(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String Report_no,String userid,String username) throws Exception{
    	String baddebt_Amt="";//轉銷呆帳金額-本月減少備抵呆帳B1
        String loss_Amt="";//轉銷呆帳金額-本月直接認列損失B2
  		String profit_Amt="";//存款準備率降低所增加盈餘(本月增加金額)B3
  		String bank_type = ( request.getAttribute("bank_type")==null ) ? "6" : (String)request.getAttribute("bank_type");
		String errMsg="";
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();		
		List updateDBList = new LinkedList();//0:sql 1:data
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data

		try {
				baddebt_Amt =  ( request.getParameter("baddebt_Amt")==null ) ? "0" : (String)request.getParameter("baddebt_Amt");
				loss_Amt =  ( request.getParameter("loss_Amt")==null ) ? "0" : (String)request.getParameter("loss_Amt");
				profit_Amt =  ( request.getParameter("profit_Amt")==null ) ? "0" : (String)request.getParameter("profit_Amt");

				sqlCmd.append("SELECT * FROM "+Report_no+" WHERE m_year=? AND m_month=? AND bank_code=?");
				paramList.add(S_YEAR);
				paramList.add(S_MONTH);
				paramList.add(bank_code);
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month");
				System.out.println(""+Report_no+".size="+data.size());

				if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法修改<br>";
				}else{
					//93.12.7 insert LOG===================================================
				    sqlCmd.delete(0,sqlCmd.length());	
				    sqlCmd.append(" INSERT INTO "+Report_no+"_LOG ");
					sqlCmd.append(" select m_year,m_month,bank_code,baddebt_Amt,loss_Amt,profit_Amt");
					sqlCmd.append(",?,?,sysdate,'U'");
					sqlCmd.append(" from "+Report_no);
					sqlCmd.append(" WHERE m_year=? AND m_month=? AND bank_code=?");
					dataList.add(userid);
					dataList.add(username);
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					dataList.add(bank_code);
					updateDBDataList.add(dataList); 	
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	
	            						   
					
				    //=========================================================================
				    updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					dataList =  new ArrayList();//儲存參數的data	
					sqlCmd.delete(0,sqlCmd.length());
				    sqlCmd.append("delete "+Report_no+" WHERE m_year=? AND m_month=? AND bank_code=?");
					dataList.add(S_YEAR); 		 
					dataList.add(S_MONTH);
					dataList.add(bank_code);
					updateDBDataList.add(dataList); 	
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	

				    updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List						
					sqlCmd.delete(0,sqlCmd.length());
				    sqlCmd.append("INSERT INTO "+Report_no+" VALUES(?,?,?,?,?,?)");

					dataList = new LinkedList();//儲存參數的data	
					dataList.add(S_YEAR);//m_year
					dataList.add(S_MONTH);//m_month
					dataList.add(bank_code);//bank_code
					dataList.add(Utility.setNoCommaFormat(baddebt_Amt));
					dataList.add(Utility.setNoCommaFormat(loss_Amt));
					dataList.add(Utility.setNoCommaFormat(profit_Amt));
					updateDBDataList.add(dataList);

	            	updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	
	            	System.out.println("updateDBDataList add");
	            	if(DBManager.updateDB_ps(updateDBList)){
						//errMsg = errMsg + "相關資料寫入資料庫成功";
					   	Date today = new Date();
    	               	int	batch_no = today.hashCode();
					   	errMsg = errMsg + UpdateA12.doParserReport_A12(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
					   	if(errMsg.startsWith("U") || errMsg.startsWith("Z")){//檢核成功//104.02.12
					       errMsg = "U:相關資料寫入資料庫成功";
					    }
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗";
					}
        		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}

		return errMsg;
	}
	//Method modiy by jei 93.12.14
	public String updateB01(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String Report_no,String userid,String username) throws Exception{
		String[]	fund_master_no  	= request.getParameterValues("fund_master_no_c");
		String[]	fund_sub_no 		= request.getParameterValues("fund_sub_no_c");
		String[]	fund_next_no		= request.getParameterValues("fund_next_no_c");
		String[]	budget_amt			= request.getParameterValues("budget_amt");
		String[]	credit_pay_amt		= request.getParameterValues("credit_pay_amt");
		String[]	credit_pay_rate		= request.getParameterValues("credit_pay_rate");
		String[]	remark      		= request.getParameterValues("remark");

		String errMsg="";		
		String bank_type = ( request.getAttribute("bank_type")==null ) ? "3" : (String)request.getAttribute("bank_type");
		
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		List updateDBList = new LinkedList();//0:sql 1:data
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data
		try {
				sqlCmd.append("SELECT * FROM B01 WHERE m_year=? AND m_month=?");
				paramList.add(S_YEAR);
				paramList.add(S_MONTH);
				
			    List data1 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");

				if (data1.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法修改<br>";
				}else{
					//insert B01_LOG
					System.out.println("insert B01_LOG....");
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" INSERT INTO B01_LOG ");
					sqlCmd.append(" select m_year,m_month,fund_master_no,fund_sub_no,fund_next_no,");
					sqlCmd.append("        budget_amt,credit_pay_amt,credit_pay_rate,remark,");
					sqlCmd.append("        ?,?,sysdate,'U'");
					sqlCmd.append(" from B01");
					sqlCmd.append(" WHERE m_year=? AND m_month=?");
					dataList.add(userid);					       
					dataList.add(username);
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					updateDBDataList.add(dataList);
					
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
					

					//delete B01
					System.out.println("delete B01_LOG....");
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					dataList = new LinkedList();//儲存參數的data 					
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" delete from B01 WHERE m_year=? AND m_month=?");
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					updateDBDataList.add(dataList);
					
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
	            	
					//Insert into B01
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
						
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO B01 VALUES (?,?,?,?,?,?,?,?,?) ");
					for (int i = 0; i < fund_master_no.length; i++) {
						dataList = new LinkedList();//儲存參數的data 
							
						dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						dataList.add(fund_master_no[i]);
						dataList.add(fund_sub_no[i]);
						dataList.add(fund_next_no[i]);
					    dataList.add(Utility.setNoCommaFormat(budget_amt[i]));
					    dataList.add(Utility.setNoCommaFormat(credit_pay_amt[i]));
					    dataList.add(Utility.setNoPercentFormat(credit_pay_rate[i]));
					    dataList.add(remark[i]);

						updateDBDataList.add(dataList);
					}
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
					
	            	if(DBManager.updateDB_ps(updateDBList)){
						errMsg = errMsg + "相關資料寫入資料庫成功";
						//93.12.21 add egg
					    Date today = new Date();
    	                int	batch_no = today.hashCode();
					    errMsg = errMsg + UpdateB01.doParserReport_B01(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗";
					}
        		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}
		return errMsg;
	}

	//Method modiy by jei 93.12.16
	public String updateB02(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String Report_no,String userid,String username) throws Exception{
        String[]	run_master_no     	= request.getParameterValues("run_master_no_c");
		String[]	run_sub_no 	    	= request.getParameterValues("run_sub_no_c");
		String[]	run_next_no	    	= request.getParameterValues("run_next_no_c");
		String[]	loan_cnt_year		= request.getParameterValues("loan_cnt_year");
		String[]	loan_amt_year		= request.getParameterValues("loan_amt_year");
		String[]	loan_cnt_totacc		= request.getParameterValues("loan_cnt_totacc");
		String[]	loan_amt_totacc  	= request.getParameterValues("loan_amt_totacc");
		String[]	loan_cnt_bal		= request.getParameterValues("loan_cnt_bal");
		String[]	loan_amt_bal_subtot	= request.getParameterValues("loan_amt_bal_subtot");
		String[]	loan_amt_bal_fund	= request.getParameterValues("loan_amt_bal_fund");
		String[]	loan_amt_bal_bank  	= request.getParameterValues("loan_amt_bal_bank");

		String errMsg="";		
		String bank_type = ( request.getAttribute("bank_type")==null ) ? "3" : (String)request.getAttribute("bank_type");

		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		List updateDBList = new LinkedList();//0:sql 1:data
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data
		try {
				
				sqlCmd.append("SELECT * FROM B02 WHERE m_year=? AND m_month=?");
				paramList.add(S_YEAR);
				paramList.add(S_MONTH);
			    List data1 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");

				if (data1.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法修改<br>";
				}else{
					//insert B02_LOG
					System.out.println("insert B02_LOG....");
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" INSERT INTO B02_LOG ");
					sqlCmd.append(" select m_year,m_month,run_master_no,run_sub_no,run_next_no,");
					sqlCmd.append("        loan_cnt_year,loan_amt_year,loan_cnt_totacc,loan_amt_totacc,");
					sqlCmd.append("        loan_cnt_bal,loan_amt_bal_subtot,loan_amt_bal_fund,loan_amt_bal_bank,");
					sqlCmd.append("        ?,?,sysdate,'U'");
					sqlCmd.append(" from B02");
					sqlCmd.append(" WHERE m_year=? AND m_month=?");
					dataList.add(userid);					       
					dataList.add(username);
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);

					//delete B02
					System.out.println("delete B02_LOG....");
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					dataList = new LinkedList();//儲存參數的data 					
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" delete from B02 WHERE m_year=? AND m_month=?");
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
					
					
					//Insert B02
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					dataList = new LinkedList();//儲存參數的data 					
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO B02 VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)");
					for (int i = 0; i < run_master_no.length; i++) {
						dataList = new LinkedList();//儲存參數的data 
						
						dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						dataList.add(run_master_no[i]);
						dataList.add(run_sub_no[i]);
						dataList.add(run_next_no[i]);					
					    dataList.add(Utility.setNoCommaFormat(loan_cnt_year[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_year[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_cnt_totacc[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_totacc[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_cnt_bal[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_bal_subtot[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_bal_fund[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_bal_bank[i]));

						updateDBDataList.add(dataList);
					}
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
					
					//for (int i=0;i<updateDBSqlList.size();i++){
					//	System.out.println(updateDBSqlList.get(i) );
					//}
	            	if(DBManager.updateDB_ps(updateDBList)){
						errMsg = errMsg + "相關資料寫入資料庫成功";
					      Date today = new Date();
    	                  int	batch_no = today.hashCode();
					      errMsg = errMsg + UpdateB02.doParserReport_B02(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗";
					}
        		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}
		return errMsg;
	}

	//Method modiy by egg 93.12.10
	public String updateM01(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String Report_no,String userid,String username) throws Exception{
		String[]	guarantee_item_no	= request.getParameterValues("guarantee_item_no_c");
		String[]	data_range			= request.getParameterValues("data_range_c");
		String[]	guarantee_cnt		= request.getParameterValues("guarantee_cnt");
		String[]	loan_amt			= request.getParameterValues("loan_amt");
		String[]	guarantee_amt		= request.getParameterValues("guarantee_amt");
		String[]	loan_bal			= request.getParameterValues("loan_bal");
		String[]	guarantee_bal		= request.getParameterValues("guarantee_bal");
		String[]	over_notpush_cnt	= request.getParameterValues("over_notpush_cnt");
		String[]	over_notpush_bal	= request.getParameterValues("over_notpush_bal");
		String[]	over_okpush_cnt		= request.getParameterValues("over_okpush_cnt");
		String[]	over_okpush_bal		= request.getParameterValues("over_okpush_bal");
		String[]	repay_tot_cnt		= request.getParameterValues("repay_tot_cnt");
		String[]	repay_tot_amt		= request.getParameterValues("repay_tot_amt");
		String[]	repay_bal_cnt		= request.getParameterValues("repay_bal_cnt");
		String[]	repay_bal_amt		= request.getParameterValues("repay_bal_amt");

		String errMsg="";
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		List updateDBList = new LinkedList();//0:sql 1:data
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data

		try {
				
				//宣告sqlcmd statment 字串
				sqlCmd.append("SELECT * FROM M01 WHERE m_year=? AND m_month=?");
				paramList.add(S_YEAR);
				paramList.add(S_MONTH);
			    List data1 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");

				if (data1.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法修改<br>";
				}else{
					//insert M01_LOG
					System.out.println("insert M01_LOG....");
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" INSERT INTO M01_LOG ");
					sqlCmd.append(" select m_year,m_month,guarantee_item_no,data_range,guarantee_cnt,loan_amt,");
					sqlCmd.append("        guarantee_amt,loan_bal,guarantee_bal,over_notpush_cnt,over_notpush_bal,");
					sqlCmd.append("        over_okpush_cnt,over_okpush_bal,repay_tot_cnt,repay_tot_amt,");
					sqlCmd.append("        repay_bal_cnt,repay_bal_amt ");
					sqlCmd.append("        ,?,?,sysdate,'U'");
					sqlCmd.append("   from M01");
					sqlCmd.append(" WHERE m_year=? AND m_month=?");
					       
					dataList.add(userid);					       
					dataList.add(username);
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);

					//delete M01
					System.out.println("delete M01_LOG....");
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					dataList = new LinkedList();//儲存參數的data 					
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("delete from M01 WHERE m_year=? AND m_month=?");
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
	            	
					//Insert M01
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					dataList = new LinkedList();//儲存參數的data 					
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO M01 VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
					for (int i = 0; i < guarantee_item_no.length; i++) {
						System.out.println("guarantee_item_no=" + guarantee_item_no[i]);
						System.out.println("data_range=" + data_range[i]);
						dataList = new LinkedList();//儲存參數的data 						
						dataList.add(S_YEAR);
						dataList.add(S_MONTH);						
						dataList.add(guarantee_item_no[i]);
						dataList.add(data_range[i]);
					    dataList.add(Utility.setNoCommaFormat(guarantee_cnt[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt[i]));
					    dataList.add(Utility.setNoCommaFormat(guarantee_amt[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_bal[i]));
					    dataList.add(Utility.setNoCommaFormat(guarantee_bal[i]));
					    dataList.add(Utility.setNoCommaFormat(over_notpush_cnt[i]));
					    dataList.add(Utility.setNoCommaFormat(over_notpush_bal[i]));
					    dataList.add(Utility.setNoCommaFormat(over_okpush_cnt[i]));
					    dataList.add(Utility.setNoCommaFormat(over_okpush_bal[i]));
					    dataList.add(Utility.setNoCommaFormat(repay_tot_cnt[i]));
					    dataList.add(Utility.setNoCommaFormat(repay_tot_amt[i]));
					    dataList.add(Utility.setNoCommaFormat(repay_bal_cnt[i]));
					    dataList.add(Utility.setNoCommaFormat(repay_bal_amt[i]));
						updateDBDataList.add(dataList); 
					}

					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
	            	
	            	if(DBManager.updateDB_ps(updateDBList)){
						errMsg = errMsg + "相關資料寫入資料庫成功";
					    Date today = new Date();
    	                int	batch_no = today.hashCode();
					    errMsg = errMsg + UpdateM01.doParserReport_M01(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W","",userid,username,batch_no);
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗";
					}
        		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}
		return errMsg;
	}

    //jei 931210
    public String updateM02(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String Report_no,String userid,String username) throws Exception{
		String[]	loan_unit_no    	= request.getParameterValues("loan_unit_no_c");
		String[]	data_range			= request.getParameterValues("data_range_c");
		String[]	guarantee_cnt		= request.getParameterValues("guarantee_cnt");
		String[]	loan_amt			= request.getParameterValues("loan_amt");
		String[]	guarantee_amt		= request.getParameterValues("guarantee_amt");
		String[]	loan_bal			= request.getParameterValues("loan_bal");
		String[]	guarantee_bal		= request.getParameterValues("guarantee_bal");
		String[]	over_notpush_cnt	= request.getParameterValues("over_notpush_cnt");
		String[]	over_notpush_bal	= request.getParameterValues("over_notpush_bal");
		String[]	over_okpush_cnt		= request.getParameterValues("over_okpush_cnt");
		String[]	over_okpush_bal		= request.getParameterValues("over_okpush_bal");
		String[]	repay_tot_cnt		= request.getParameterValues("repay_tot_cnt");
		String[]	repay_tot_amt		= request.getParameterValues("repay_tot_amt");
		String[]	repay_bal_cnt		= request.getParameterValues("repay_bal_cnt");
		String[]	repay_bal_amt		= request.getParameterValues("repay_bal_amt");

		String errMsg="";		
		String bank_type = ( request.getAttribute("bank_type")==null ) ? "4" : (String)request.getAttribute("bank_type");
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		List updateDBList = new LinkedList();//0:sql 1:data
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data
		
		try {
				
				sqlCmd.append("SELECT * FROM M02 WHERE m_year=? AND m_month=?");
				paramList.add(S_YEAR);
				paramList.add(S_MONTH);
			    List data1 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");

				if (data1.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法修改<br>";
				}else{
					//insert M02_LOG
					System.out.println("insert M02_LOG....");
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" INSERT INTO M02_LOG ");
					sqlCmd.append(" select m_year,m_month,loan_unit_no,data_range,guarantee_cnt,loan_amt,");
					sqlCmd.append("        guarantee_amt,loan_bal,guarantee_bal,over_notpush_cnt,over_notpush_bal,");
					sqlCmd.append("        over_okpush_cnt,over_okpush_bal,repay_tot_cnt,repay_tot_amt,");
					sqlCmd.append("        repay_bal_cnt,repay_bal_amt ");
					sqlCmd.append("        ,?,?,sysdate,'U'");
					sqlCmd.append("   from M02");
					sqlCmd.append(" WHERE m_year=? AND m_month=?");
					dataList.add(userid);					       
					dataList.add(username);
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);					

					//delete M02
					System.out.println("delete M02_LOG....");
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					dataList = new LinkedList();//儲存參數的data 					
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" delete from M02 WHERE m_year=? AND m_month=?");
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
					//Insert M02
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					dataList = new LinkedList();//儲存參數的data 					
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO M02 VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
					for (int i = 0; i < loan_unit_no.length; i++) {
						System.out.println("loan_unit_no=" + loan_unit_no[i]);
						System.out.println("data_range=" + data_range[i]);
						dataList = new LinkedList();//儲存參數的data 	
						
						dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						dataList.add(loan_unit_no[i]);
						dataList.add(data_range[i]);
					    dataList.add(Utility.setNoCommaFormat(guarantee_cnt[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt[i]));
					    dataList.add(Utility.setNoCommaFormat(guarantee_amt[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_bal[i]));
					    dataList.add(Utility.setNoCommaFormat(guarantee_bal[i]));
					    dataList.add(Utility.setNoCommaFormat(over_notpush_cnt[i]));
					    dataList.add(Utility.setNoCommaFormat(over_notpush_bal[i]));
					    dataList.add(Utility.setNoCommaFormat(over_okpush_cnt[i]));
					    dataList.add(Utility.setNoCommaFormat(over_okpush_bal[i]));
					    dataList.add(Utility.setNoCommaFormat(repay_tot_cnt[i]));
					    dataList.add(Utility.setNoCommaFormat(repay_tot_amt[i]));
					    dataList.add(Utility.setNoCommaFormat(repay_bal_cnt[i]));
					    dataList.add(Utility.setNoCommaFormat(repay_bal_amt[i]));
						updateDBDataList.add(dataList);  
					}

					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
	            	
	            	if(DBManager.updateDB_ps(updateDBList)){
						errMsg = errMsg + "相關資料寫入資料庫成功";
					    Date today = new Date();
    	                int	batch_no = today.hashCode();
					    errMsg = errMsg + UpdateM02.doParserReport_M02(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗";
					}
        		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}
		return errMsg;
	}

    //Method modiy by egg 93.12.12
    public String updateM03(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String Report_no,String userid,String username) throws Exception{
		String errMsg="";		
		String bank_type = ( request.getAttribute("bank_type")==null ) ? "4" : (String)request.getAttribute("bank_type");

		//M03
		String[]	div_no						= request.getParameterValues("div_no");
    	String[]	guarantee_cnt_month			= request.getParameterValues("guarantee_cnt_month");
    	String[]	loan_amt_month				= request.getParameterValues("loan_amt_month");
    	String[]	guarantee_amt_month			= request.getParameterValues("guarantee_amt_month");
    	String[]	guarantee_cnt_year			= request.getParameterValues("guarantee_cnt_year");
    	String[]	loan_amt_year				= request.getParameterValues("loan_amt_year");
    	String[]	guarantee_amt_year			= request.getParameterValues("guarantee_amt_year");
    	String[]	guarantee_bal_totacc		= request.getParameterValues("guarantee_bal_totacc");
    	String[]	guarantee_bal_totacc_over	= request.getParameterValues("guarantee_bal_totacc_over");
    	String[]	repay_bal_totacc			= request.getParameterValues("repay_bal_totacc");

		//M03 Note
		String[]	note_no						= request.getParameterValues("note_no");
    	String[]	note_amt_rate				= request.getParameterValues("note_amt_rate");
		System.out.println("note_no="+note_no);
		
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		List updateDBList = new LinkedList();//0:sql 1:data
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data
		try {
				sqlCmd.append("SELECT * FROM M03 WHERE m_year=? AND m_month=?");
				paramList.add(S_YEAR);
				paramList.add(S_MONTH);
			    List data1 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");
			    sqlCmd.delete(0,sqlCmd.length());
				sqlCmd.append("SELECT * FROM M03_NOTE WHERE m_year=? AND m_month=?");
			    List data2 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");
				
				if (data1.size() == 0 || data2.size()==0){
				    errMsg = errMsg + "此筆資料不存在無法修改<br>";
				}else{
					//insert M03_LOG
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" INSERT INTO M03_LOG ");
					sqlCmd.append(" select m_year,m_month,div_no,guarantee_cnt_month,loan_amt_month,guarantee_amt_month,");
					sqlCmd.append("        guarantee_cnt_year,loan_amt_year,guarantee_amt_year,guarantee_bal_totacc,");
					sqlCmd.append("        guarantee_bal_totacc_over,repay_bal_totacc");
					sqlCmd.append("        ,?,?,sysdate,'U'");
					sqlCmd.append(" from M03");
					sqlCmd.append(" WHERE m_year=? AND m_month=?");
					dataList.add(userid);					       
					dataList.add(username);
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);			
	            	
					//insert M03_NOTE_LOG
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List						
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" INSERT INTO M03_NOTE_LOG ");
					sqlCmd.append(" select m_year,m_month,note_no,note_amt_rate");
					sqlCmd.append("        ,?,?,sysdate,'U'");
					sqlCmd.append("   from M03_NOTE");
					sqlCmd.append(" WHERE m_year=? AND m_month=?");
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	

					//delete M03,M03_NOTE
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					dataList = new LinkedList();//儲存參數的data 					
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" delete from M03 WHERE m_year=? AND m_month=?");
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
	            	
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List											
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" delete from M03_NOTE WHERE m_year=? AND m_month=?");
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
	            	
	            	
	            	//Insert M03
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					dataList = new LinkedList();//儲存參數的data 					
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO M03 VALUES (?,?,?,?,?,?,?,?,?,?,?,?)");
					for (int i = 0; i < div_no.length; i++) {   
						dataList = new LinkedList();//儲存參數的data
						if(div_no[i].equals("MP") || div_no[i].equals("YP")){							
							dataList.add(S_YEAR);
							dataList.add(S_MONTH);
							dataList.add(div_no[i]);
					    	dataList.add(Utility.setNoPercentFormat(guarantee_cnt_month[i]));
					    	dataList.add(Utility.setNoPercentFormat(loan_amt_month[i]));
					    	dataList.add(Utility.setNoPercentFormat(guarantee_amt_month[i]));
					    	dataList.add(Utility.setNoPercentFormat(guarantee_cnt_year[i]));
					    	dataList.add(Utility.setNoPercentFormat(loan_amt_year[i]));
					    	dataList.add(Utility.setNoPercentFormat(guarantee_amt_year[i]));
					    	dataList.add(Utility.setNoPercentFormat(guarantee_bal_totacc[i]));
					    	dataList.add(Utility.setNoPercentFormat(guarantee_bal_totacc_over[i]));
					    	dataList.add(Utility.setNoPercentFormat(repay_bal_totacc[i]));
						}else{							
							dataList.add(S_YEAR);
							dataList.add(S_MONTH);							
							dataList.add(div_no[i]);
					    	dataList.add(Utility.setNoCommaFormat(guarantee_cnt_month[i]));
					    	dataList.add(Utility.setNoCommaFormat(loan_amt_month[i]));
					    	dataList.add(Utility.setNoCommaFormat(guarantee_amt_month[i]));
					    	dataList.add(Utility.setNoCommaFormat(guarantee_cnt_year[i]));
					    	dataList.add(Utility.setNoCommaFormat(loan_amt_year[i]));
					    	dataList.add(Utility.setNoCommaFormat(guarantee_amt_year[i]));
					    	dataList.add(Utility.setNoCommaFormat(guarantee_bal_totacc[i]));
					    	dataList.add(Utility.setNoCommaFormat(guarantee_bal_totacc_over[i]));
					    	dataList.add(Utility.setNoCommaFormat(repay_bal_totacc[i]));
						}
						updateDBDataList.add(dataList);   
					}
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
	            	
					System.out.println("note_no.length="+note_no.length);
					//Insert M03_NOTE
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					dataList = new LinkedList();//儲存參數的data 					
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO M03_NOTE VALUES (?,?,?,?)");
					for (int i = 0; i < note_no.length; i++) {
						System.out.println("note_no="+note_no[i]);
						dataList = new LinkedList();//儲存參數的data
						if(note_no[i].equals("NT1P")){
							dataList.add(S_YEAR);
							dataList.add(S_MONTH);
							dataList.add(note_no[i]);
							dataList.add(Utility.setNoPercentFormat(note_amt_rate[i]));
						}else{
							dataList.add(S_YEAR);
							dataList.add(S_MONTH);
							dataList.add(note_no[i]);
							dataList.add(Utility.setNoCommaFormat(note_amt_rate[i]));
						}
						updateDBDataList.add(dataList);  
					}
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	            	
					
	            	if(DBManager.updateDB_ps(updateDBList)){
						errMsg = errMsg + "相關資料寫入資料庫成功";
						//93.12.21 add egg
					    Date today = new Date();
    	                int	batch_no = today.hashCode();
					    errMsg = errMsg + UpdateM03.doParserReport_M03(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗";
					}
        		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}
		return errMsg;
	}

	//jei 931212
    public String updateM04(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String Report_no,String userid,String username) throws Exception{
		String[]	loan_use_no         	= request.getParameterValues("loan_use_no_c");
		String[]	guarantee_no_month		= request.getParameterValues("guarantee_no_month");
		String[]	guarantee_no_month_p	= request.getParameterValues("guarantee_no_month_p");
		String[]	loan_amt_month	    	= request.getParameterValues("loan_amt_month");
		String[]	loan_amt_month_p		= request.getParameterValues("loan_amt_month_p");
		String[]	guarantee_amt_month		= request.getParameterValues("guarantee_amt_month");
		String[]	guarantee_amt_month_p	= request.getParameterValues("guarantee_amt_month_p");
		String[]	guarantee_no_year	    = request.getParameterValues("guarantee_no_year");
		String[]	guarantee_no_year_p		= request.getParameterValues("guarantee_no_year_p");
		String[]	loan_amt_year		    = request.getParameterValues("loan_amt_year");
		String[]	loan_amt_year_p	    	= request.getParameterValues("loan_amt_year_p");
		String[]	guarantee_amt_year	 	= request.getParameterValues("guarantee_amt_year");
		String[]	guarantee_amt_year_p	= request.getParameterValues("guarantee_amt_year_p");
		String[]	guarantee_no_totacc		= request.getParameterValues("guarantee_no_totacc");
		String[]	guarantee_no_totacc_p	= request.getParameterValues("guarantee_no_totacc_p");
		String[]	loan_amt_totacc		    = request.getParameterValues("loan_amt_totacc");
		String[]	loan_amt_totacc_p		= request.getParameterValues("loan_amt_totacc_p");
    	String[]	guarantee_amt_totacc	= request.getParameterValues("guarantee_amt_totacc");
		String[]	guarantee_amt_totacc_p	= request.getParameterValues("guarantee_amt_totacc_p");

		String errMsg="";		
		String bank_type = ( request.getAttribute("bank_type")==null ) ? "4" : (String)request.getAttribute("bank_type");
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		List updateDBList = new LinkedList();//0:sql 1:data
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data
		
		try {
				
				sqlCmd.append("SELECT * FROM M04 WHERE m_year=? AND m_month=?");
			    paramList.add(S_YEAR);
				paramList.add(S_MONTH);
			    List data1 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");

				if (data1.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法修改<br>";
				}else{
					//insert M04_LOG
					System.out.println("insert M04_LOG....");
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" INSERT INTO M04_LOG ");
					sqlCmd.append(" select m_year,m_month,loan_use_no,guarantee_no_month,guarantee_no_month_p,loan_amt_month,");
					sqlCmd.append("        loan_amt_month_p,guarantee_amt_month,guarantee_amt_month_p,guarantee_no_year,");
					sqlCmd.append("        guarantee_no_year_p,loan_amt_year,loan_amt_year_p,guarantee_amt_year,guarantee_amt_year_p,");
					sqlCmd.append("        guarantee_no_totacc,guarantee_no_totacc_p,loan_amt_totacc,loan_amt_totacc_p,");
					sqlCmd.append("        guarantee_amt_totacc,guarantee_amt_totacc_p ");
					sqlCmd.append("        ,?,?,sysdate,'U'");
					sqlCmd.append("   from M04");
					sqlCmd.append(" WHERE m_year=? AND m_month=?");
					dataList.add(userid);					       
					dataList.add(username);
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);		

					//delete M04
					System.out.println("delete M04_LOG....");
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					dataList = new LinkedList();//儲存參數的data 					
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" delete from M04 WHERE m_year=? AND m_month=?");
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
					
					//Insert M04
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					dataList = new LinkedList();//儲存參數的data 					
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO M04 VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
					for (int i = 0; i < loan_use_no.length; i++) {
						System.out.println("loan_use_no=" + loan_use_no[i]);       
						dataList = new LinkedList();//儲存參數的data 			
						dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						dataList.add(loan_use_no[i]);
					    dataList.add(Utility.setNoCommaFormat(guarantee_no_month[i]));
					    dataList.add(Utility.setNoPercentFormat(guarantee_no_month_p[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_month[i]));
					    dataList.add(Utility.setNoPercentFormat(loan_amt_month_p[i]));
					    dataList.add(Utility.setNoCommaFormat(guarantee_amt_month[i]));
					    dataList.add(Utility.setNoPercentFormat(guarantee_amt_month_p[i]));
					    dataList.add(Utility.setNoCommaFormat(guarantee_no_year[i]));
					    dataList.add(Utility.setNoPercentFormat(guarantee_no_year_p[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_year[i]));
					    dataList.add(Utility.setNoPercentFormat(loan_amt_year_p[i]));
					    dataList.add(Utility.setNoCommaFormat(guarantee_amt_year[i]));
					    dataList.add(Utility.setNoPercentFormat(guarantee_amt_year_p[i]));
					    dataList.add(Utility.setNoCommaFormat(guarantee_no_totacc[i]));
					    dataList.add(Utility.setNoPercentFormat(guarantee_no_totacc_p[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_totacc[i]));
					    dataList.add(Utility.setNoPercentFormat(loan_amt_totacc_p[i]));
					    dataList.add(Utility.setNoCommaFormat(guarantee_amt_totacc[i]));
					    dataList.add(Utility.setNoPercentFormat(guarantee_amt_totacc_p[i]));
						updateDBDataList.add(dataList);    
					}

					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
	            	
	            	if(DBManager.updateDB_ps(updateDBList)){
						errMsg = errMsg + "相關資料寫入資料庫成功";
						//93.12.01 add egg
					    Date today = new Date();
    	                int	batch_no = today.hashCode();
					    errMsg = errMsg + UpdateM04.doParserReport_M04(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗";
					}
        		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}
		return errMsg;
	}

    //add by winnin 2004.12.09
    public String updateM05(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String Report_no,String userid,String username) throws Exception{
    	String[]	Loan_Unit_No	= request.getParameterValues("loan_unit_no");
    	String[]	Period_No		= request.getParameterValues("period_no");
    	String[]	Item_No			= request.getParameterValues("item_no");
    	String[]	Repay_Cnt		= request.getParameterValues("repay_cnt");
    	String[]	Repay_Amt		= request.getParameterValues("repay_amt");
    	String[]	Run_Notgood_Cnt	= request.getParameterValues("run_notgood_cnt");
    	String[]	Run_Notgood_Amt	= request.getParameterValues("run_notgood_amt");
    	String[]	Turn_Out_Cnt= request.getParameterValues("turn_out_cnt");
    	String[]	Turn_Out_Amt= request.getParameterValues("turn_out_amt");
    	String[]	Diease_Cnt		= request.getParameterValues("diease_cnt");
    	String[]	Dieaserepay_Amt	= request.getParameterValues("dieaserepay_amt");
    	String[]	Disaster_Cnt	= request.getParameterValues("disaster_cnt");
    	String[]	Disaster_Amt	= request.getParameterValues("disaster_amt");
    	String[]	Corun_Out_Cnt	= request.getParameterValues("corun_out_cnt");
    	String[]	Corun_Out_Amt	= request.getParameterValues("corun_out_amt");
    	String[]	Other_Cnt		= request.getParameterValues("other_cnt");
    	String[]	Other_Amt		= request.getParameterValues("other_amt");

		//M05_TOTACC
    	String[]	Loan_Unit_No_C	= request.getParameterValues("loan_unit_no_c");
    	String[]	Fix_No	= request.getParameterValues("fix_no_c");
    	String[]	Guarantee_No_Totacc	= request.getParameterValues("guarantee_no_totacc");
    	String[]	Guarantee_Amt_Totacc	= request.getParameterValues("guarantee_amt_totacc");

		//M05_NOTE
    	String[]	Note_No	= request.getParameterValues("note_no");
    	String[]	Note_Amt_Rate	= request.getParameterValues("note_amt_rate");

		String bank_type = ( request.getAttribute("bank_type")==null ) ? "4" : (String)request.getAttribute("bank_type");
		
		String errMsg="";
		
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		List updateDBList = new LinkedList();//0:sql 1:data
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data

		try {
				
				sqlCmd.append("SELECT * FROM M05 WHERE m_year=? AND m_month=?");
			    paramList.add(S_YEAR);
				paramList.add(S_MONTH);
			    List data1 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");
			    sqlCmd.delete(0,sqlCmd.length());
				sqlCmd.append("SELECT * FROM M05_TOTACC WHERE m_year=? AND m_month=?");
			    List data2 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");
			    sqlCmd.delete(0,sqlCmd.length());
				sqlCmd.append("SELECT * FROM M05_NOTE WHERE m_year=? AND m_month=?");
			    List data3 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");

				if (data1.size() == 0 || data2.size()==0 || data3.size()==0){
				    errMsg = errMsg + "此筆資料不存在無法修改<br>";
				}else{
					//insert M05_LOG、M05_TOTACC_LOG、M05_NOTE_LOG==================================================
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" INSERT INTO M05_LOG ");
					sqlCmd.append(" select m_year,m_month,loan_unit_no,period_no,item_no,");
					sqlCmd.append("        repay_cnt,repay_amt,run_notgood_cnt,run_notgood_amt,turn_out_cnt,");
					sqlCmd.append("        turn_out_amt,diease_cnt,dieaserepay_amt,disaster_cnt,disaster_amt,");
					sqlCmd.append("        corun_out_cnt,corun_out_amt,other_cnt,other_amt ");
					sqlCmd.append("        ,?,?,sysdate,'U'");
					sqlCmd.append("   from M05");
					sqlCmd.append(" WHERE m_year=? AND m_month=?");
					dataList.add(userid);					       
					dataList.add(username);
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	

					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List								
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" INSERT INTO M05_TOTACC_LOG ");
					sqlCmd.append(" select m_year,m_month,loan_unit_no,fix_no,guarantee_no_totacc,guarantee_amt_totacc");
					sqlCmd.append("        ,?,?,sysdate,'U'");
					sqlCmd.append("   from M05_TOTACC");
					sqlCmd.append(" WHERE m_year=? AND m_month=?");
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	
	            	
	            	updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List								
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" INSERT INTO M05_NOTE_LOG ");
					sqlCmd.append(" select m_year,m_month,note_no,note_amt_rate");
					sqlCmd.append("        ,?,?,sysdate,'U'");
					sqlCmd.append("   from M05_NOTE");
					sqlCmd.append(" WHERE m_year=? AND m_month=?");
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	

					//delete M05,M05_TOTACC,M05_NOTE
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					dataList = new LinkedList();//儲存參數的data 					
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" delete from M05 WHERE m_year=? AND m_month=?");
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
					
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" delete from M05_TOTACC WHERE m_year=? AND m_month=?");
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
	            	
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" delete from M05_NOTE WHERE m_year=? AND m_month=?");
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
					
					//Insert M05
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					dataList = new LinkedList();//儲存參數的data 					
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO M05 VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
					for (int i = 0; i < Period_No.length; i++) {
						dataList = new LinkedList();//儲存參數的data 	
						
						dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						dataList.add(Loan_Unit_No[i]);
						dataList.add(Period_No[i]);
						dataList.add(Item_No[i]);				
					    dataList.add(Utility.setNoCommaFormat(Repay_Cnt[i]));
					    dataList.add(Utility.setNoCommaFormat(Repay_Amt[i]));
					    dataList.add(Utility.setNoCommaFormat(Run_Notgood_Cnt[i]));
					    dataList.add(Utility.setNoCommaFormat(Run_Notgood_Amt[i]));
					    dataList.add(Utility.setNoCommaFormat(Turn_Out_Cnt[i]));
					    dataList.add(Utility.setNoCommaFormat(Turn_Out_Amt[i]));
					    dataList.add(Utility.setNoCommaFormat(Diease_Cnt[i]));
					    dataList.add(Utility.setNoCommaFormat(Dieaserepay_Amt[i]));
					    dataList.add(Utility.setNoCommaFormat(Disaster_Cnt[i]));
					    dataList.add(Utility.setNoCommaFormat(Disaster_Amt[i]));
					    dataList.add(Utility.setNoCommaFormat(Corun_Out_Cnt[i]));
					    dataList.add(Utility.setNoCommaFormat(Corun_Out_Amt[i]));
					    dataList.add(Utility.setNoCommaFormat(Other_Cnt[i]));
					    dataList.add(Utility.setNoCommaFormat(Other_Amt[i]));
						updateDBDataList.add(dataList);   
					}
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
	            	
	            	
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					dataList = new LinkedList();//儲存參數的data 					
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO M05_TOTACC VALUES (?,?,?,?,?,?)");
					for (int i = 0; i < Loan_Unit_No_C.length; i++) {
						dataList = new LinkedList();//儲存參數的data 
						dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						dataList.add(Loan_Unit_No_C[i]);
						dataList.add(Fix_No[i]);
					    dataList.add(Utility.setNoCommaFormat(Guarantee_No_Totacc[i]));
					    dataList.add(Utility.setNoCommaFormat(Guarantee_Amt_Totacc[i]));
						updateDBDataList.add(dataList);   
					}
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
	            	
	            	
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					dataList = new LinkedList();//儲存參數的data 					
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO M05_NOTE VALUES (?,?,?,?)");
					for (int i = 0; i < Note_No.length; i++) {
						dataList = new LinkedList();//儲存參數的data 
						dataList.add(S_YEAR);
						dataList.add(S_MONTH); 
						dataList.add(Note_No[i]);
						dataList.add(Utility.setNoCommaFormat(Note_Amt_Rate[i]));
						updateDBDataList.add(dataList);    
					}
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
					
	            	if(DBManager.updateDB_ps(updateDBList)){
						errMsg = errMsg + "相關資料寫入資料庫成功";
					    Date today = new Date();
    	                int	batch_no = today.hashCode();
					    errMsg = errMsg + UpdateM05.doParserReport_M05(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗";
					}
        		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}

		return errMsg;
	}

    //Method modify by egg 93.12.14
    public String updateM06_M07(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String Report_no,String userid,String username) throws Exception{
    	//資料宣告
    	String[]	area_no					= request.getParameterValues("area_no");
    	String[]	guarantee_no_month		= request.getParameterValues("guarantee_no_month");
    	String[]	guarantee_amt_month		= request.getParameterValues("guarantee_amt_month");
    	String[]	loan_amt_month			= request.getParameterValues("loan_amt_month");
    	String[]	guarantee_no_year		= request.getParameterValues("guarantee_no_year");
    	String[]	guarantee_amt_year		= request.getParameterValues("guarantee_amt_year");
    	String[]	loan_amt_year			= request.getParameterValues("loan_amt_year");
    	String[]	guarantee_no_totacc		= request.getParameterValues("guarantee_no_totacc");
    	String[]	guarantee_amt_totacc	= request.getParameterValues("guarantee_amt_totacc");
    	String[]	loan_amt_totacc			= request.getParameterValues("loan_amt_totacc");
    	String[]	guarantee_bal_no		= request.getParameterValues("guarantee_bal_no");
    	String[]	guarantee_bal_amt		= request.getParameterValues("guarantee_bal_amt");
    	String[]	guarantee_bal_p			= request.getParameterValues("guarantee_bal_p");
    	String[]	loan_bal				= request.getParameterValues("loan_bal");

		String bank_type = ( request.getAttribute("bank_type")==null ) ? "4" : (String)request.getAttribute("bank_type");
		
		String errMsg="";
		
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		List updateDBList = new LinkedList();//0:sql 1:data
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data
		
		
		try {
				sqlCmd.append("SELECT * FROM " + Report_no + " WHERE m_year=? AND m_month=?");
			    paramList.add(S_YEAR);
				paramList.add(S_MONTH);
			    List data1 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");

				if (data1.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法修改<br>";
				}else{
					//insert M06_LOG(M07_LOG)
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" INSERT INTO " + Report_no + "_LOG ");
					sqlCmd.append(" select m_year,m_month,area_no,guarantee_no_month,guarantee_amt_month,loan_amt_month,");
					sqlCmd.append("        guarantee_no_year,guarantee_amt_year,loan_amt_year,guarantee_no_totacc,guarantee_amt_totacc,");
					sqlCmd.append("        loan_amt_totacc,guarantee_bal_no,guarantee_bal_amt,guarantee_bal_p,loan_bal");
					sqlCmd.append("        ,?,?,sysdate,'U'");
					sqlCmd.append("   from " + Report_no);
					sqlCmd.append(" WHERE m_year=? AND m_month=?");
					dataList.add(userid);					       
					dataList.add(username);
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	

					//delete M06(M07)
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					dataList = new LinkedList();//儲存參數的data 					
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("delete from " + Report_no + " WHERE m_year=? AND m_month=?");
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
					
					
					//Insert M06(M07)
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					dataList = new LinkedList();//儲存參數的data 					
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO " + Report_no + " VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
					for (int i = 0; i < area_no.length; i++) {
						dataList = new LinkedList();//儲存參數的data 
						dataList.add(S_YEAR);
						dataList.add(S_MONTH); 
						dataList.add(area_no[i]);
					    dataList.add(Utility.setNoCommaFormat(guarantee_no_month[i]));
					    dataList.add(Utility.setNoCommaFormat(guarantee_amt_month[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_month[i]));
					    dataList.add(Utility.setNoCommaFormat(guarantee_no_year[i]));
					    dataList.add(Utility.setNoCommaFormat(guarantee_amt_year[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_year[i]));
					    dataList.add(Utility.setNoCommaFormat(guarantee_no_totacc[i]));
					    dataList.add(Utility.setNoCommaFormat(guarantee_amt_totacc[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_totacc[i]));
					    dataList.add(Utility.setNoCommaFormat(guarantee_bal_no[i]));
					    dataList.add(Utility.setNoCommaFormat(guarantee_bal_amt[i]));
					    dataList.add(Utility.setNoPercentFormat(guarantee_bal_p[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_bal[i]));
						updateDBDataList.add(dataList);     
					}

					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
					
	            	if(DBManager.updateDB_ps(updateDBList)){	            	
						errMsg = errMsg + "相關資料寫入資料庫成功";
						//93.12.21 add egg
						if(Report_no.equals("M06")){
					    	Date today = new Date();
    	                	int	batch_no = today.hashCode();
					    	errMsg = errMsg + UpdateM06.doParserReport_M06(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
						}else{
							Date today = new Date();
    	                	int	batch_no = today.hashCode();
					    	errMsg = errMsg + UpdateM07.doParserReport_M07(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
						}
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗";
					}
        		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}
		
    	return errMsg;
	}

    //Method modify by egg 93.12.16
    public String updateM08(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String Report_no,String userid,String username) throws Exception{
    	//資料宣告
    	String[]	id_no					= request.getParameterValues("id_no");
    	String[]	data_range				= request.getParameterValues("data_range");
    	String[]	guarantee_no_month		= request.getParameterValues("guarantee_no_month");
    	String[]	loan_amt_month			= request.getParameterValues("loan_amt_month");
    	String[]	guarantee_amt_month		= request.getParameterValues("guarantee_amt_month");
    	String[]	guarantee_bal_month		= request.getParameterValues("guarantee_bal_month");
    	String[]	guarantee_bal_p			= request.getParameterValues("guarantee_bal_p");

		String bank_type = ( request.getAttribute("bank_type")==null ) ? "" : (String)request.getAttribute("bank_type");
		
		String errMsg="";
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		List updateDBList = new LinkedList();//0:sql 1:data
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data	
		
		
		//年份資料表示轉換93-->093
		String s_year="0000"+S_YEAR; s_year=s_year.substring(s_year.length()-3);
		String s_month="0000"+S_MONTH; s_month=s_month.substring(s_month.length()-2);
		System.out.println("轉換後s_year =[" + s_year + "]");
		System.out.println("轉換後s_month=[" + s_month + "]");

		try {
				sqlCmd.append("SELECT * FROM " + Report_no + " WHERE m_year=? AND m_month=?");
			    paramList.add(S_YEAR);
				paramList.add(S_MONTH);
			    List data1 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");

				if (data1.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法修改<br>";
				}else{
					//insert M08_LOG
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" INSERT INTO " + Report_no + "_LOG ");
					sqlCmd.append(" select m_year,m_month,id_no,data_range,guarantee_no_month,loan_amt_month,guarantee_amt_month,");
					sqlCmd.append("        guarantee_bal_month,guarantee_bal_p");
					sqlCmd.append("        ,?,?,sysdate,'U'");
					sqlCmd.append("   from " + Report_no);
					sqlCmd.append(" WHERE m_year=? AND m_month=?");
					dataList.add(userid);					       
					dataList.add(username);
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	

					//delete M08
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					dataList = new LinkedList();//儲存參數的data 					
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("delete from " + Report_no + " WHERE m_year=? AND m_month=?");
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
					

					System.out.println("id_no.length=["+ id_no.length + "]");
					//Insert M08
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					dataList = new LinkedList();//儲存參數的data 					
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO " + Report_no + " VALUES (?,?,?,?,?,?,?,?,?)");
					for (int i = 0; i < id_no.length; i++) {
						dataList = new LinkedList();//儲存參數的data 
						dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						dataList.add(id_no[i]);
						dataList.add(data_range[i]);
					    dataList.add(Utility.setNoCommaFormat(guarantee_no_month[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_month[i]));
					    dataList.add(Utility.setNoCommaFormat(guarantee_amt_month[i]));
					    dataList.add(Utility.setNoCommaFormat(guarantee_bal_month[i]));
					    dataList.add(Utility.setNoPercentFormat(guarantee_bal_p[i]));
						updateDBDataList.add(dataList);     
					}

					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
					
	            	if(DBManager.updateDB_ps(updateDBList)){	
						errMsg = errMsg + "相關資料寫入資料庫成功";
					      Date today = new Date();
    	                  int  batch_no = today.hashCode();
					      errMsg = errMsg + UpdateM08.doParserReport_M08(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗";
					}
        		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}
		
    	return errMsg;
	}

    //Method modify by egg 93.12.14
    public String updateB03(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String Report_no,String userid,String username) throws Exception{
		String errMsg="";	

		//B03_1
    	String[]	funs_master_no_1		= request.getParameterValues("funs_master_no_1");
    	String[]	funs_sub_no_1			= request.getParameterValues("funs_sub_no_1");
    	String[]	funs_next_no_1			= request.getParameterValues("funs_next_no_1");
    	String[]	loan_cnt_totacc			= request.getParameterValues("loan_cnt_totacc");
    	String[]	loan_amt_totacc_fund	= request.getParameterValues("loan_amt_totacc_fund");
    	String[]	loan_amt_totacc_bank	= request.getParameterValues("loan_amt_totacc_bank");
    	String[]	loan_amt_totacc_tot		= request.getParameterValues("loan_amt_totacc_tot");
    	String[]	loan_cnt_bal			= request.getParameterValues("loan_cnt_bal");
    	String[]	loan_amt_bal_fund		= request.getParameterValues("loan_amt_bal_fund");
    	String[]	loan_amt_bal_bank		= request.getParameterValues("loan_amt_bal_bank");
    	String[]	loan_amt_bal_tot		= request.getParameterValues("loan_amt_bal_tot");

		//B03_2
    	String[]	funs_master_no_2		= request.getParameterValues("funs_master_no_2");
    	String[]	funs_sub_no_2			= request.getParameterValues("funs_sub_no_2");
    	String[]	funs_next_no_2			= request.getParameterValues("funs_next_no_2");
    	String[]	loan_amt_bal			= request.getParameterValues("loan_amt_bal");
    	String[]	loan_amt_over			= request.getParameterValues("loan_amt_over");
    	String[]	loan_rate_over			= request.getParameterValues("loan_rate_over");

    	//B03_3
    	String[]	funo_master_no			= request.getParameterValues("funo_master_no");
    	String[]	funo_sub_no				= request.getParameterValues("funo_sub_no");
    	String[]	funo_next_no			= request.getParameterValues("funo_next_no");
    	String[]	funo_amt				= request.getParameterValues("funo_amt");
    	String[]	funo_rate				= request.getParameterValues("funo_rate");

    	//B03_4
    	String[]	bank_no					= request.getParameterValues("bank_no");
    	String[]	machine_cnt				= request.getParameterValues("machine_cnt");
    	String[]	machine_amt				= request.getParameterValues("machine_amt");
    	String[]	land_cnt				= request.getParameterValues("land_cnt");
    	String[]	land_amt				= request.getParameterValues("land_amt");
    	String[]	house_cnt				= request.getParameterValues("house_cnt");
    	String[]	house_amt				= request.getParameterValues("house_amt");
    	String[]	build_cnt				= request.getParameterValues("build_cnt");
    	String[]	build_amt				= request.getParameterValues("build_amt");
    	String[]	tot_cnt					= request.getParameterValues("tot_cnt");
    	String[]	tot_amt					= request.getParameterValues("tot_amt");

		
		String bank_type = ( request.getAttribute("bank_type")==null ) ? "3" : (String)request.getAttribute("bank_type");
		System.out.println("Method Inpute paramater Report_no=[" + Report_no + "]");
		System.out.println("bank_type=["+bank_type+"]");
	
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		List updateDBList = new LinkedList();//0:sql 1:data
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data	
		try {
				
				sqlCmd.append("SELECT * FROM " + Report_no + "_1 WHERE m_year=? AND m_month=?");
			    paramList.add(S_YEAR);
				paramList.add(S_MONTH);
			    List data1 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");
			    sqlCmd.delete(0,sqlCmd.length());
			    sqlCmd.append("SELECT * FROM " + Report_no + "_2 WHERE m_year=? AND m_month=?");			   
			    List data2 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");
			    sqlCmd.delete(0,sqlCmd.length());
			    sqlCmd.append("SELECT * FROM " + Report_no + "_3 WHERE m_year=? AND m_month=?");			    
			    List data3 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");
			    sqlCmd.delete(0,sqlCmd.length());
			    sqlCmd.append("SELECT * FROM " + Report_no + "_4 WHERE m_year=? AND m_month=?");			    
			    List data4 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");

				if (data1.size() == 0 || data2.size()==0 || data3.size()==0 || data4.size()==0){
				    errMsg = errMsg + "此筆資料不存在無法修改<br>";
				}else{
					//insert B03_1_LOG
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" INSERT INTO " + Report_no + "_1_LOG ");
					sqlCmd.append(" select m_year, m_month, funs_master_no, funs_sub_no, funs_next_no, loan_cnt_totacc,");
					sqlCmd.append(" loan_amt_totacc_fund, loan_amt_totacc_bank, loan_amt_totacc_tot, loan_cnt_bal,");
					sqlCmd.append(" loan_amt_bal_fund, loan_amt_bal_bank, loan_amt_bal_tot");
					sqlCmd.append("  ,?,?,sysdate,'U'");
					sqlCmd.append(" from " + Report_no + "_1");
					sqlCmd.append(" WHERE m_year=? AND m_month=?");
					dataList.add(userid);					       
					dataList.add(username);
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	
	            	
					//insert B03_2_LOG
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List									
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" INSERT INTO " + Report_no + "_2_LOG ");
					sqlCmd.append(" select m_year, m_month, funs_master_no, funs_sub_no, funs_next_no, loan_amt_bal, loan_amt_over, loan_rate_over,");
					sqlCmd.append(" ?,?,sysdate,'U'");
					sqlCmd.append(" from " + Report_no + "_2");
					sqlCmd.append(" WHERE m_year=? AND m_month=?");
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	
	            	
					//insert B03_3_LOG
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List									
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" INSERT INTO " + Report_no + "_3_LOG ");
					sqlCmd.append(" select m_year, m_month, funo_master_no, funo_sub_no, funo_next_no, funo_amt, funo_rate");
					sqlCmd.append(" ,?,?,sysdate,'U'");
					sqlCmd.append(" from " + Report_no + "_3");
					sqlCmd.append(" WHERE m_year=? AND m_month=?");
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	
					//insert B03_4_LOG
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List									
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" INSERT INTO " + Report_no + "_4_LOG ");
					sqlCmd.append(" select m_year, m_month, bank_no, machine_cnt, machine_amt, land_cnt, land_amt,");
					sqlCmd.append(" house_cnt, house_amt, build_cnt, build_amt, tot_cnt, tot_amt,");
					sqlCmd.append(" ?,?,sysdate,'U'");
					sqlCmd.append(" from " + Report_no + "_4");
					sqlCmd.append(" WHERE m_year=? AND m_month=?");
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	

					//delete B03_1
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					dataList = new LinkedList();//儲存參數的data								
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" delete from " + Report_no + "_1 WHERE m_year=? AND m_month=?");
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	
	            	
					//delete B03_2
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" delete from " + Report_no + "_2 WHERE m_year=? AND m_month=?");
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	
					
					//delete B03_3
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" delete from " + Report_no + "_3 WHERE m_year=? AND m_month=?");
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);		
	            				
					//delete B03_4
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" delete from " + Report_no + "_4 WHERE m_year=? AND m_month=?");
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	
	            	
					//Insert B03_1
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					dataList = new LinkedList();//儲存參數的data								
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO " + Report_no + "_1 VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?) ");
					for (int i = 0; i < loan_cnt_totacc.length; i++) {
						dataList = new LinkedList();//儲存參數的data	
						
						dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						dataList.add(funs_master_no_1[i]);
						dataList.add(funs_sub_no_1[i]);
						dataList.add(funs_next_no_1[i]);
					    dataList.add(Utility.setNoCommaFormat(loan_cnt_totacc[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_totacc_fund[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_totacc_bank[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_totacc_tot[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_cnt_bal[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_bal_fund[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_bal_bank[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_bal_tot[i]));
						updateDBDataList.add(dataList);     
					}

					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
	            	
	            	
					//Update B03_2
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					dataList = new LinkedList();//儲存參數的data								
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO " + Report_no + "_2 VALUES (?,?,?,?,?,?,?,?) ");
					for (int i = 0; i < loan_amt_bal.length; i++) {
						dataList = new LinkedList();//儲存參數的data	
						
						dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						dataList.add(funs_master_no_2[i]);
						dataList.add(funs_sub_no_2[i]);
						dataList.add(funs_next_no_2[i]);
					    dataList.add(Utility.setNoCommaFormat(loan_amt_bal[i]));
					    dataList.add(Utility.setNoCommaFormat(loan_amt_over[i]));
					    dataList.add(Utility.setNoPercentFormat(loan_rate_over[i]));
						updateDBDataList.add(dataList);     
					}

					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
	            	
					//Update B03_3
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					dataList = new LinkedList();//儲存參數的data								
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO " + Report_no + "_3 VALUES (?,?,?,?,?,?,?) ");
					for (int i = 0; i < funo_amt.length; i++) {
						dataList = new LinkedList();//儲存參數的data	
						dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						dataList.add(funo_master_no[i]);
						dataList.add(funo_sub_no[i]);
						dataList.add(funo_next_no[i]);
					    dataList.add(Utility.setNoCommaFormat(funo_amt[i]));
					    dataList.add(Utility.setNoPercentFormat(funo_rate[i]));
						updateDBDataList.add(dataList);     
					}

					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);
	            	
					//Update B03_4
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					dataList = new LinkedList();//儲存參數的data								
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO " + Report_no + "_4 VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?) ");
					for (int i = 0; i < machine_cnt.length; i++) {
						dataList = new LinkedList();//儲存參數的data	
						dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						dataList.add(bank_no[i]);
					    dataList.add(Utility.setNoCommaFormat(machine_cnt[i]));
					    dataList.add(Utility.setNoCommaFormat(machine_amt[i]));
					    dataList.add(Utility.setNoCommaFormat(land_cnt[i]));
					    dataList.add(Utility.setNoCommaFormat(land_amt[i]));
					    dataList.add(Utility.setNoCommaFormat(house_cnt[i]));
					    dataList.add(Utility.setNoCommaFormat(house_amt[i]));
					    dataList.add(Utility.setNoCommaFormat(build_cnt[i]));
					    dataList.add(Utility.setNoCommaFormat(build_amt[i]));
					    dataList.add(Utility.setNoCommaFormat(tot_cnt[i]));
					    dataList.add(Utility.setNoCommaFormat(tot_amt[i]));
						updateDBDataList.add(dataList);     
					}

					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);

					if(DBManager.updateDB_ps(updateDBList)){
						errMsg = errMsg + "相關資料寫入資料庫成功";
						//93.12.21 add egg
					    Date today = new Date();
    	                int	batch_no = today.hashCode();
					    errMsg = errMsg + UpdateB03.doParserReport_B03(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W",bank_type,userid,username,batch_no);
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗";
					}
        		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}
		
    	return errMsg;
	}


    //94.11.11 add by 2295
	public String updateF01(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String Report_no,String userid,String username) throws Exception{
		String errMsg="";		
		String titleidx[] = {"A","B","C","D","E"};
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		List updateDBList = new LinkedList();//0:sql 1:data
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data	
		try {
				sqlCmd.append("SELECT * FROM F01 WHERE m_year=? AND m_month=? and bank_code=?");
				paramList.add(S_YEAR);
				paramList.add(S_MONTH);
			    paramList.add(bank_code);
			    List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month");

				if (dbData.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法修改<br>";
				}else{
					//insert F01_LOG
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" INSERT INTO F01_LOG ");
					sqlCmd.append(" select m_year,m_month,bank_code,dep_type,acct_type,acct_cnt_tm,bal_lm,dep_tm,wtd_tm,bal_tm");
					sqlCmd.append("        ,?,?,sysdate,'U'");
					sqlCmd.append(" from F01");
					sqlCmd.append(" WHERE m_year=? AND m_month=? and bank_code=?");

					dataList.add(userid);					       
					dataList.add(username);
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					dataList.add(bank_code);
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	

					//delete F01
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					dataList = new LinkedList();//儲存參數的data								
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" delete from F01 WHERE m_year=? AND m_month=? and bank_code=?");
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					dataList.add(bank_code);
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	

					//取出form裡的所有變數===================================
					Enumeration ep = request.getParameterNames();
					Hashtable t = new Hashtable();
					String name = "";
					for ( ; ep.hasMoreElements() ; ) {
				 		name = (String)ep.nextElement();
				 		if(request.getParameter(name) == null){
				 		  t.put( name, "" );
				 		}else{
				 		  t.put( name, request.getParameter(name) );
				 		}
				 		System.out.println(name+"="+request.getParameter(name));
					}
					//95.07.17 不能新增 FIX BY 2495
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					dataList = new LinkedList();//儲存參數的data								
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO F01 VALUES (?,?,?,?,?,?,?,?,?,?)");
                    for(int j=0;j<titleidx.length;j++){//A,B,C,D,E
		                for(int k=1;k<=5;k++){
						        if(t.get(titleidx[j]+k+"1") != null ) {
						        /*
						          System.out.println((String)t.get(titleidx[j]+k+"1"));
						          System.out.println((String)t.get(titleidx[j]+k+"2"));
						          System.out.println((String)t.get(titleidx[j]+k+"3"));
						          System.out.println((String)t.get(titleidx[j]+k+"4"));
						          System.out.println((String)t.get(titleidx[j]+k+"5"));
						         */ 
						          dataList = new LinkedList();//儲存參數的data
						          dataList.add(S_YEAR);
						          dataList.add(S_MONTH);
						          dataList.add(bank_code);
						          dataList.add(titleidx[j]);
						          dataList.add(k);
						          dataList.add(Utility.setNoCommaFormat((String)t.get(titleidx[j]+k+"1")));
						          dataList.add(Utility.setNoCommaFormat((String)t.get(titleidx[j]+k+"2")));
						          dataList.add(Utility.setNoCommaFormat((String)t.get(titleidx[j]+k+"3")));
						          dataList.add(Utility.setNoCommaFormat((String)t.get(titleidx[j]+k+"4")));
						          dataList.add(Utility.setNoCommaFormat((String)t.get(titleidx[j]+k+"5")));						         
						          updateDBDataList.add(dataList);   
						       }
					    }//end of 1~5
				    }//end of "A,B,C,D,E"
					
					
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);

					if(DBManager.updateDB_ps(updateDBList)){
						//errMsg = errMsg + "相關資料寫入資料庫成功";
					    Date today = new Date();
    	                int	batch_no = today.hashCode();
					    errMsg = errMsg + UpdateF01.doParserReport_F01(Report_no,S_YEAR,S_MONTH,"", bank_code,"M","W","",userid,username,batch_no);
					    if(errMsg.startsWith("U")){//檢核成功
					      errMsg = "U:相關資料寫入資料庫成功";
					    }
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗";
					}
        		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}
		return errMsg;
	}

    //96.12.03 add A09 by 2295
    //97.06.13 add A10 by 2295
    //99.09.30 原deleteA01_A05更名為deleteA01_A10_B01_B02_M01_M08_F01
    //B03.M03.M05不合併
    //102.04.26 add a02.amt_name by 2295
	public String deleteA01_A10_B01_B02_M01_M08_F01(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String Report_no,String userid,String username) throws Exception{										  

		String errMsg="";
		String input_method="";//93.12.18 add by 2295
		
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		List updateDBList = new LinkedList();//0:sql 1:data
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data	
		
		try {
			   	sqlCmd.append("SELECT * FROM WML01 WHERE m_year=? AND m_month=? AND bank_code=? AND report_no=?");
				paramList.add(S_YEAR);
				paramList.add(S_MONTH);
				paramList.add(bank_code);
				paramList.add(Report_no);
				
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,add_date,batch_no,update_date");
				System.out.println("WML01.size="+data.size());
				if (data.size() == 0){
					errMsg = errMsg + "無資料可刪除(WML01)<br>";
				}else{
				    input_method = (String)((DataObject)data.get(0)).getValue("input_method");
				    sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO WML01_LOG ");
					sqlCmd.append(" select m_year,m_month,bank_code,report_no,input_method,add_user,add_name,add_date,common_center,upd_method,upd_code ");
					sqlCmd.append(",batch_no,lock_status,user_id,user_name,update_date,?,?,sysdate,'D'");
					sqlCmd.append(" FROM WML01 WHERE m_year=? AND m_month=?");
					sqlCmd.append(" AND bank_code=? AND report_no=?");
					dataList.add(userid);					       
					dataList.add(username);
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					dataList.add(bank_code);
					dataList.add(Report_no);
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	

				   
				    updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					dataList = new LinkedList();//儲存參數的data								
					sqlCmd.delete(0,sqlCmd.length());
				    
				    sqlCmd.append("DELETE FROM WML01 WHERE m_year=? AND m_month=? AND bank_code=? AND report_no=? AND input_method=?");

				    dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					dataList.add(bank_code);
					dataList.add(Report_no);
					dataList.add(input_method);
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	
					if(Report_no.equals("A10")){
						sqlCmd.delete(0,sqlCmd.length());
						paramList.clear();
						sqlCmd.append("SELECT * FROM WML03 WHERE m_year=? AND m_month=? AND bank_code=? AND report_no=?");
						paramList.add(S_YEAR);
						paramList.add(S_MONTH);
						paramList.add(bank_code);
						paramList.add(Report_no);
						
					    List WML03data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,add_date,batch_no,update_date");
						System.out.println("WML03.size="+WML03data.size());
						if (WML03data.size() > 0){
							updateDBDataList = new ArrayList();//儲存參數的List
							updateDBSqlList = new ArrayList();//儲存參數的List	
							dataList = new LinkedList();//儲存參數的data								
							sqlCmd.delete(0,sqlCmd.length());
							sqlCmd.append("INSERT INTO WML03_LOG ");  
				            sqlCmd.append("select m_year,m_month,bank_code,report_no,serial_no,remark,user_id,user_name,update_date,?,?,sysdate,'D' from WML03  "); 
				            sqlCmd.append("where m_year=? and m_month=? and bank_code=? and report_no=? ");  
							dataList.add(userid);					       
							dataList.add(username);
							dataList.add(S_YEAR);
							dataList.add(S_MONTH);
							dataList.add(bank_code);
							dataList.add(Report_no);
							updateDBDataList.add(dataList);
							updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
			            	updateDBSqlList.add(updateDBDataList);
			            	updateDBList.add(updateDBSqlList);	

						   
						    updateDBDataList = new ArrayList();//儲存參數的List
							updateDBSqlList = new ArrayList();//儲存參數的List	
							dataList = new LinkedList();//儲存參數的data								
							sqlCmd.delete(0,sqlCmd.length());
							sqlCmd.append("delete WML03 where m_year=? and m_month=? and bank_code=? and report_no=? ");  
							dataList.add(S_YEAR);
							dataList.add(S_MONTH);
							dataList.add(bank_code);
							dataList.add(Report_no);
							updateDBDataList.add(dataList);
							updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
			            	updateDBSqlList.add(updateDBDataList);
			            	updateDBList.add(updateDBSqlList);	
						}
					}
					sqlCmd.delete(0,sqlCmd.length());
					paramList = new ArrayList();	
				    sqlCmd.append("SELECT * FROM "+Report_no+" WHERE m_year=? AND m_month=? ");
				    paramList.add(S_YEAR);
					paramList.add(S_MONTH);
					if(Report_no.indexOf("B") == -1 && Report_no.indexOf("M") == -1){//99.09.30 add B/M開頭的.沒有區分bank_code
				       sqlCmd.append(" AND bank_code=?");					
					   paramList.add(bank_code);
					}
					
					data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");
					System.out.println(""+Report_no+".size="+data.size());
					if (data.size() == 0){
						errMsg = errMsg + "無資料可刪除("+Report_no+")<br>";
					}else{
					    //Insert to LOG==========================================
					    updateDBDataList = new ArrayList();//儲存參數的List
						updateDBSqlList = new ArrayList();//儲存參數的List	
						dataList = new LinkedList();//儲存參數的data								
						sqlCmd.delete(0,sqlCmd.length());
						
						    sqlCmd.append(" INSERT INTO "+Report_no+"_LOG ");
						    if(Report_no.equals("A08")){
						       sqlCmd.append(" select m_year,m_month,bank_code,WarnAccount_Cnt,LimitAccount_Cnt,ErrorAccount_Cnt,OtherAccount_Cnt,DepositAccount_TCnt");
						    }else if(Report_no.equals("A09")){//96.12.03 add by 2295
						       sqlCmd.append(" select m_year,m_month,bank_code,over_cnt,over_amt,PUSH_over_amt,totalamt,PUSH_totalamt,Over_total_rate");
						    }else if(Report_no.equals("A10")){//97.06.13 add by 2295
						       sqlCmd.append(" select m_year,m_month,bank_code,loan1_amt,loan2_amt,loan3_amt,loan4_amt,invest1_amt,invest2_amt,invest3_amt,invest4_amt,other1_amt,other2_amt,other3_amt,other4_amt"
						    		          +",loan1_baddebt,loan2_baddebt,loan3_baddebt,loan4_baddebt,build1_baddebt,build2_baddebt,build3_baddebt,build4_baddebt"
										      +",baddebt_flag,baddebt_noenough,baddebt_delay,baddebt_104,baddebt_105,baddebt_106,baddebt_107,baddebt_108");
						    }else if(Report_no.equals("A12")){//97.06.13 add by 2295
							   sqlCmd.append(" select m_year,m_month,bank_code,baddebt_amt,loss_amt,profit_amt");
						    }else if(Report_no.equals("A05")){
							   sqlCmd.append(" select m_year,m_month,bank_code,acc_code,amt,amt_name");
							}else if(Report_no.equals("A06")){//95.04.11 add by 2295
							   sqlCmd.append(" select m_year,m_month,bank_code,acc_code,amt_3month,amt_6month,amt_1year,amt_2year,amt_over2year,amt_total");
							}else if(Report_no.equals("B01")){
							   sqlCmd.append(" select m_year,m_month,fund_master_no,fund_sub_no,fund_next_no,budget_amt,credit_pay_amt,credit_pay_rate,remark ");
							}else if(Report_no.equals("B02")){
							   sqlCmd.append(" select m_year,m_month,run_master_no,run_sub_no,run_next_no,loan_cnt_year,loan_amt_year,loan_cnt_totacc,loan_amt_totacc,loan_cnt_bal,loan_amt_bal_subtot,loan_amt_bal_fund,loan_amt_bal_bank,");
							}else if(Report_no.equals("M01")){
							   sqlCmd.append(" select m_year,m_month,guarantee_item_no,data_range,guarantee_cnt,loan_amt,");
						       sqlCmd.append("        guarantee_amt,loan_bal,guarantee_bal,over_notpush_cnt,over_notpush_bal,");
						       sqlCmd.append("        over_okpush_cnt,over_okpush_bal,repay_tot_cnt,repay_tot_amt,");
						       sqlCmd.append("        repay_bal_cnt,repay_bal_amt ");
						    }else if(Report_no.equals("M02")){ 
						       sqlCmd.append(" select m_year,m_month,loan_unit_no,data_range,guarantee_cnt,loan_amt,");
						       sqlCmd.append("        guarantee_amt,loan_bal,guarantee_bal,over_notpush_cnt,over_notpush_bal,");
						       sqlCmd.append("        over_okpush_cnt,over_okpush_bal,repay_tot_cnt,repay_tot_amt,");
						       sqlCmd.append("        repay_bal_cnt,repay_bal_amt ");
						    }else if(Report_no.equals("M04")){ 
						       sqlCmd.append(" select m_year,m_month,loan_use_no,guarantee_no_month,guarantee_no_month_p,loan_amt_month,");
						       sqlCmd.append("        loan_amt_month_p,guarantee_amt_month,guarantee_amt_month_p,guarantee_no_year,");
						       sqlCmd.append("        guarantee_no_year_p,loan_amt_year,loan_amt_year_p,guarantee_amt_year,guarantee_amt_year_p,");
						       sqlCmd.append("        guarantee_no_totacc,guarantee_no_totacc_p,loan_amt_totacc,loan_amt_totacc_p,");
						       sqlCmd.append("        guarantee_amt_totacc,guarantee_amt_totacc_p ");					   
						    }else if(Report_no.equals("M06") || Report_no.equals("M07")){
						       sqlCmd.append(" select m_year,m_month,area_no,guarantee_no_month,guarantee_amt_month,loan_amt_month,guarantee_no_year,");
							   sqlCmd.append("        guarantee_amt_year,loan_amt_year,guarantee_no_totacc,guarantee_amt_totacc,loan_amt_totacc,");
							   sqlCmd.append("        guarantee_bal_no,guarantee_bal_amt,guarantee_bal_p,loan_bal ");								   
						    }else if(Report_no.equals("M08")){ 
						       sqlCmd.append(" select m_year,m_month,id_no,data_range,guarantee_no_month,loan_amt_month,guarantee_amt_month,guarantee_bal_month,guarantee_bal_p");					  
						    }else if(Report_no.equals("F01")){    
						       sqlCmd.append(" select m_year,m_month,bank_code,dep_type,acct_type,acct_cnt_tm,bal_lm,dep_tm,wtd_tm,bal_tm");					   				        
							}else{
							   sqlCmd.append(" select m_year,m_month,bank_code,acc_code,amt");
							}						
						    if(Report_no.equals("A02")){//102.04.25 add a02.amt_name 
						       sqlCmd.append(",?,?,sysdate,'D',amt_name");
							}else{
							   sqlCmd.append(",?,?,sysdate,'D'");
							}
							
							sqlCmd.append(" from "+Report_no);
							sqlCmd.append(" WHERE m_year=? AND m_month=? ");
							
						    dataList.add(userid);					       
						    dataList.add(username);
						    dataList.add(S_YEAR);
						    dataList.add(S_MONTH);
						    if(Report_no.indexOf("B") == -1 && Report_no.indexOf("M") == -1){//99.09.30 add B/M開頭的.沒有區分bank_code
						       sqlCmd.append(" AND bank_code=?");
						       dataList.add(bank_code);
						    }   
						    updateDBDataList.add(dataList);
						    updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
		            	    updateDBSqlList.add(updateDBDataList);
		            	    updateDBList.add(updateDBSqlList);	
						
					    //====================================================================
					    updateDBDataList = new ArrayList();//儲存參數的List
						updateDBSqlList = new ArrayList();//儲存參數的List	
						dataList = new LinkedList();//儲存參數的data								
						sqlCmd.delete(0,sqlCmd.length());
						sqlCmd.append("delete "+Report_no+" WHERE m_year=? AND m_month=? ");
						dataList.add(S_YEAR);
					    dataList.add(S_MONTH);
					    if(Report_no.indexOf("B") == -1 && Report_no.indexOf("M") == -1){//99.09.30 add B/M開頭的.沒有區分bank_code
					       sqlCmd.append(" AND bank_code=?");
					       dataList.add(bank_code);
					    }   
					    updateDBDataList.add(dataList);
					    updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	    updateDBSqlList.add(updateDBDataList);
	            	    updateDBList.add(updateDBSqlList);	

						if(DBManager.updateDB_ps(updateDBList)){
							errMsg = errMsg + "相關資料刪除成功";
						}else{
				   			errMsg = errMsg + "相關資料刪除失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
						}
					}
				}

		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料刪除失敗";
		}

		return errMsg;
	}
	
	//Method modify by egg 93.12.15
	public String deleteB03(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String Report_no,String userid,String username) throws Exception{		
		String errMsg="";
		String input_method=""; //add by egg 2004.12.21
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		List updateDBList = new LinkedList();//0:sql 1:data
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data	
		
		//轉換為int型態
		System.out.println("S_MONTH=[" + String.valueOf(Integer.parseInt(S_MONTH)) + "]");
		try {
			    sqlCmd.append("SELECT * FROM WML01 WHERE m_year=? AND m_month=? AND bank_code=? AND report_no=?");

			    paramList.add(S_YEAR);
				paramList.add(S_MONTH);
				paramList.add(bank_code);
				paramList.add(Report_no);
				
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,add_date,batch_no,update_date");
				System.out.println("WML01.size="+data.size());
				if (data.size() == 0){
					errMsg = errMsg + "無資料可刪除(WML01)<br>";
				}else{
					input_method = (String)((DataObject)data.get( 0 ) ).getValue("input_method");

					//Insert WML01_LOG
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO WML01_LOG ");
					sqlCmd.append(" select m_year,m_month,bank_code,report_no,input_method,add_user,add_name,add_date,common_center,upd_method,upd_code ");
					sqlCmd.append(",batch_no,lock_status,user_id,user_name,update_date,?,?,sysdate,'D'");
					sqlCmd.append(" FROM WML01 WHERE m_year=? AND m_month=?");
					sqlCmd.append(" AND bank_code=? AND report_no=?");
					dataList.add(userid);					       
					dataList.add(username);
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					dataList.add(bank_code);
					dataList.add(Report_no);
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	
					//Delete WML01
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					dataList = new LinkedList();//儲存參數的data								
					sqlCmd.delete(0,sqlCmd.length());
				    sqlCmd.append("DELETE FROM WML01 WHERE m_year=? AND m_month=? AND bank_code=? AND report_no=? AND input_method=?");
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					dataList.add(bank_code);
					dataList.add(Report_no);
					dataList.add(input_method);
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	
	            	
				    sqlCmd.delete(0,sqlCmd.length());
					paramList = new ArrayList();	
				    sqlCmd.append("SELECT * FROM " + Report_no + "_1 WHERE m_year=? AND m_month=?");
				    paramList.add(S_YEAR);
				    paramList.add(S_MONTH);
					data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month");
					if (data.size() == 0){
						errMsg = errMsg + "無資料可刪除("+Report_no+")<br>";
					}else{
						//新增B03_1_LOG
						updateDBDataList = new ArrayList();//儲存參數的List
						updateDBSqlList = new ArrayList();//儲存參數的List	
						dataList = new LinkedList();//儲存參數的data								
						sqlCmd.delete(0,sqlCmd.length());
						sqlCmd.append(" INSERT INTO " + Report_no + "_1_LOG ");
					    sqlCmd.append(" select m_year, m_month, funs_master_no, funs_sub_no, funs_next_no, loan_cnt_totacc,");
					    sqlCmd.append(" loan_amt_totacc_fund, loan_amt_totacc_bank, loan_amt_totacc_tot, loan_cnt_bal,");
					    sqlCmd.append(" loan_amt_bal_fund, loan_amt_bal_bank, loan_amt_bal_tot");
					    sqlCmd.append(" ,?,?,sysdate,'U'");
					    sqlCmd.append(" from " + Report_no + "_1");
					    sqlCmd.append(" WHERE m_year=? AND m_month=?");
						dataList.add(userid);					       
					    dataList.add(username);
					    dataList.add(S_YEAR);
					    dataList.add(S_MONTH);
					    updateDBDataList.add(dataList);
					    updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	    updateDBSqlList.add(updateDBDataList);
	            	    updateDBList.add(updateDBSqlList);	
						
						//insert B03_2_LOG
						updateDBDataList = new ArrayList();//儲存參數的List
						updateDBSqlList = new ArrayList();//儲存參數的List							
						sqlCmd.delete(0,sqlCmd.length());
						sqlCmd.append(" INSERT INTO " + Report_no + "_2_LOG ");
					    sqlCmd.append(" select m_year, m_month, funs_master_no, funs_sub_no, funs_next_no, loan_amt_bal, loan_amt_over, loan_rate_over,");
					    sqlCmd.append(" ?,?,sysdate,'U'");
					    sqlCmd.append(" from " + Report_no + "_2");
					    sqlCmd.append(" WHERE m_year=? AND m_month=?");
						updateDBDataList.add(dataList);
					    updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	    updateDBSqlList.add(updateDBDataList);
	            	    updateDBList.add(updateDBSqlList);	
						
						
						//insert B03_3_LOG
						updateDBDataList = new ArrayList();//儲存參數的List
						updateDBSqlList = new ArrayList();//儲存參數的List							
						sqlCmd.delete(0,sqlCmd.length());
						sqlCmd.append(" INSERT INTO " + Report_no + "_3_LOG ");
					    sqlCmd.append(" select m_year, m_month, funo_master_no, funo_sub_no, funo_next_no, funo_amt, funo_rate");
					    sqlCmd.append(" ,?,?,sysdate,'U'");
					    sqlCmd.append(" from " + Report_no + "_3");
					    sqlCmd.append(" WHERE m_year=? AND m_month=?");
						updateDBDataList.add(dataList);
					    updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	    updateDBSqlList.add(updateDBDataList);
	            	    updateDBList.add(updateDBSqlList);	
	            	    
						//insert B03_4_LOG
						updateDBDataList = new ArrayList();//儲存參數的List
						updateDBSqlList = new ArrayList();//儲存參數的List	
						sqlCmd.append(" INSERT INTO " + Report_no + "_4_LOG ");
					    sqlCmd.append(" select m_year, m_month, bank_no, machine_cnt, machine_amt, land_cnt, land_amt,");
					    sqlCmd.append(" house_cnt, house_amt, build_cnt, build_amt, tot_cnt, tot_amt,");
					    sqlCmd.append(" ?,?,sysdate,'U'");
					    sqlCmd.append(" from " + Report_no + "_4");
					    sqlCmd.append(" WHERE m_year=? AND m_month=?");
						updateDBDataList.add(dataList);
					    updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	    updateDBSqlList.add(updateDBDataList);
	            	    updateDBList.add(updateDBSqlList);

						//delete B03_1
						updateDBDataList = new ArrayList();//儲存參數的List
						updateDBSqlList = new ArrayList();//儲存參數的List	
						dataList = new LinkedList();//儲存參數的data								
						sqlCmd.delete(0,sqlCmd.length());
						sqlCmd.append(" delete from " + Report_no + "_1 WHERE m_year=? AND m_month=?");
						dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						updateDBDataList.add(dataList);
					    updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	    updateDBSqlList.add(updateDBDataList);
	            	    updateDBList.add(updateDBSqlList);
	            	    
						//delete B03_2
						updateDBDataList = new ArrayList();//儲存參數的List
						updateDBSqlList = new ArrayList();//儲存參數的List	
						sqlCmd.delete(0,sqlCmd.length());
						sqlCmd.append(" delete from " + Report_no + "_2 WHERE m_year=? AND m_month=?");
						updateDBDataList.add(dataList);
					    updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	    updateDBSqlList.add(updateDBDataList);
	            	    updateDBList.add(updateDBSqlList);
						//delete B03_3
						updateDBDataList = new ArrayList();//儲存參數的List
						updateDBSqlList = new ArrayList();//儲存參數的List	
						sqlCmd.delete(0,sqlCmd.length());
						sqlCmd.append(" delete from " + Report_no + "_3 WHERE m_year=? AND m_month=?");
						updateDBDataList.add(dataList);
					    updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	    updateDBSqlList.add(updateDBDataList);
	            	    updateDBList.add(updateDBSqlList);
						//delete B03_4
						updateDBDataList = new ArrayList();//儲存參數的List
						updateDBSqlList = new ArrayList();//儲存參數的List	
						sqlCmd.delete(0,sqlCmd.length());
						sqlCmd.append(" delete from " + Report_no + "_4 WHERE m_year=? AND m_month=?");
						updateDBDataList.add(dataList);
					    updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	    updateDBSqlList.add(updateDBDataList);
	            	    updateDBList.add(updateDBSqlList);						

						if(DBManager.updateDB_ps(updateDBList)){
							errMsg = errMsg + "相關資料刪除成功";
						}else{
				   			errMsg = errMsg + "相關資料刪除失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
						}
					}
				}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料刪除失敗";
		}
		
		return errMsg;
	}	

	//Method modify by egg 93.12.12
	public String deleteM03(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String Report_no,String userid,String username) throws Exception{	
		String errMsg="";
		String input_method=""; //add by egg 2004.12.21
	
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		List updateDBList = new LinkedList();//0:sql 1:data
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data	
		try {
			    
				sqlCmd.append("SELECT * FROM WML01 WHERE m_year=? AND m_month=? AND bank_code=? AND report_no=?");
			    paramList.add(S_YEAR);
				paramList.add(S_MONTH);
				paramList.add(bank_code);
				paramList.add(Report_no);
				
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,add_date,batch_no,update_date");
				System.out.println("WML01.size="+data.size());
				if (data.size() == 0){
					errMsg = errMsg + "無資料可刪除(WML01)<br>";
				}else{
					input_method = (String)((DataObject)data.get( 0 ) ).getValue("input_method");
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO WML01_LOG ");
					sqlCmd.append(" select m_year,m_month,bank_code,report_no,input_method,add_user,add_name,add_date,common_center,upd_method,upd_code ");
					sqlCmd.append(",batch_no,lock_status,user_id,user_name,update_date,?,?,sysdate,'D'");
					sqlCmd.append(" FROM WML01 WHERE m_year=? AND m_month=?");
					sqlCmd.append(" AND bank_code=? AND report_no=?");
					dataList.add(userid);					       
					dataList.add(username);
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					dataList.add(bank_code);
					dataList.add(Report_no);
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	

					//Delete WML01
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					dataList = new LinkedList();//儲存參數的data								
					sqlCmd.delete(0,sqlCmd.length());
				    sqlCmd.append("DELETE FROM WML01 WHERE m_year=? AND m_month=? AND ");
					sqlCmd.append("bank_code=? AND report_no=? AND input_method=?");

				    dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					dataList.add(bank_code);
					dataList.add(Report_no);
					dataList.add(input_method);
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	

					sqlCmd.delete(0,sqlCmd.length());
					paramList = new ArrayList();					    
				    sqlCmd.append("SELECT * FROM M03 WHERE m_year=? AND m_month=?");				    
					paramList.add(S_YEAR);
				    paramList.add(S_MONTH);
					data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month");
					if (data.size() == 0){
						errMsg = errMsg + "無資料可刪除("+Report_no+")<br>";
					}else{
					    updateDBDataList = new ArrayList();//儲存參數的List
					    updateDBSqlList = new ArrayList();//儲存參數的List	
					    dataList = new LinkedList();//儲存參數的data								
					    sqlCmd.delete(0,sqlCmd.length());
						sqlCmd.append(" INSERT INTO M03_LOG ");
						sqlCmd.append(" select m_year,m_month,div_no,guarantee_cnt_month,loan_amt_month,");
						sqlCmd.append("        guarantee_amt_month,guarantee_cnt_year,loan_amt_year,guarantee_amt_year,guarantee_bal_totacc,");
						sqlCmd.append("        guarantee_bal_totacc_over,repay_bal_totacc");
						sqlCmd.append("        ,?,?,sysdate,'U'");
						sqlCmd.append("   from M03");
						sqlCmd.append(" WHERE m_year=? AND m_month=?");
						dataList.add(userid);					       
					    dataList.add(username);
					    dataList.add(S_YEAR);
					    dataList.add(S_MONTH);
					    updateDBDataList.add(dataList);
					    updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	    updateDBSqlList.add(updateDBDataList);
	            	    updateDBList.add(updateDBSqlList);	
	            	    
	            	    //Delete M03
						updateDBDataList = new ArrayList();//儲存參數的List
					    updateDBSqlList = new ArrayList();//儲存參數的List	
					    dataList = new LinkedList();//儲存參數的data								
					    sqlCmd.delete(0,sqlCmd.length());
						sqlCmd.append(" delete from M03 WHERE m_year=? AND m_month=?");
						dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						updateDBDataList.add(dataList);
					    updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	    updateDBSqlList.add(updateDBDataList);
	            	    updateDBList.add(updateDBSqlList);

						sqlCmd.delete(0,sqlCmd.length());
						sqlCmd.append("SELECT * FROM M03_NOTE WHERE m_year=? AND m_month=?");
						data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month");
						if (data.size() == 0){
							errMsg = errMsg + "無資料可刪除("+Report_no+")<br>";
						}else{
						    //INSERT M03_NOTE_LOG
						    updateDBDataList = new ArrayList();//儲存參數的List
					        updateDBSqlList = new ArrayList();//儲存參數的List	
					        dataList = new LinkedList();//儲存參數的data								
					        sqlCmd.delete(0,sqlCmd.length());
							sqlCmd.append(" INSERT INTO M03_NOTE_LOG ");
							sqlCmd.append(" select m_year,m_month,note_no,note_amt_rate");
							sqlCmd.append("        ,?,?,sysdate,'U'");
							sqlCmd.append("   from M03_NOTE");
							sqlCmd.append(" WHERE m_year=? AND m_month=?");
							dataList.add(userid);					       
					        dataList.add(username);
					        dataList.add(S_YEAR);
					        dataList.add(S_MONTH);
					        updateDBDataList.add(dataList);
					        updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	        updateDBSqlList.add(updateDBDataList);
	            	        updateDBList.add(updateDBSqlList);	
	            	        
	            	        // delete M03_NOTE
	            	        updateDBDataList = new ArrayList();//儲存參數的List
					        updateDBSqlList = new ArrayList();//儲存參數的List	
					        dataList = new LinkedList();//儲存參數的data								
					        sqlCmd.delete(0,sqlCmd.length());
							sqlCmd.append(" delete from M03_NOTE WHERE m_year=? AND m_month=?");
							dataList.add(S_YEAR);
					        dataList.add(S_MONTH);
					        updateDBDataList.add(dataList);
					        updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	        updateDBSqlList.add(updateDBDataList);
	            	        updateDBList.add(updateDBSqlList);	
	            	        
							if(DBManager.updateDB_ps(updateDBList)){
								errMsg = errMsg + "相關資料刪除成功";
							}else{
				   				errMsg = errMsg + "相關資料刪除失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
							}
						}
					}
				}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料刪除失敗";
		}

		return errMsg;
	}    

	//add by winnin 2004.12.09
	public String deleteM05(HttpServletRequest request,String S_YEAR,String S_MONTH,String bank_code,String Report_no,String userid,String username) throws Exception{		
		String errMsg="";
		String input_method=""; //add by winnin 2004.12.21
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		List updateDBList = new LinkedList();//0:sql 1:data
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data	
			
		try {
				sqlCmd.append("SELECT * FROM WML01 WHERE m_year=? AND m_month=? AND bank_code=? AND report_no=?");

			    paramList.add(S_YEAR);
				paramList.add(S_MONTH);
				paramList.add(bank_code);
				paramList.add(Report_no);
				
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,add_date,batch_no,update_date");
				System.out.println("WML01.size="+data.size());
				if (data.size() == 0){
					errMsg = errMsg + "無資料可刪除(WML01)<br>";
				}else{
					input_method = (String)((DataObject)data.get(0)).getValue("input_method");
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO WML01_LOG ");
					sqlCmd.append(" select m_year,m_month,bank_code,report_no,input_method,add_user,add_name,add_date,common_center,upd_method,upd_code ");
					sqlCmd.append(",batch_no,lock_status,user_id,user_name,update_date,?,?,sysdate,'D'");
					sqlCmd.append(" FROM WML01 WHERE m_year=? AND m_month=?");
					sqlCmd.append(" AND bank_code=? AND report_no=?");
					dataList.add(userid);					       
					dataList.add(username);
					dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					dataList.add(bank_code);
					dataList.add(Report_no);
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	
					
					//Delete WML01
					updateDBDataList = new ArrayList();//儲存參數的List
					updateDBSqlList = new ArrayList();//儲存參數的List	
					dataList = new LinkedList();//儲存參數的data								
					sqlCmd.delete(0,sqlCmd.length());
				    sqlCmd.append("DELETE FROM WML01 WHERE m_year=? AND m_month=? AND ");
					sqlCmd.append("bank_code=? AND report_no=? AND input_method=?");

				    dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					dataList.add(bank_code);
					dataList.add(Report_no);
					dataList.add(input_method);
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	updateDBSqlList.add(updateDBDataList);
	            	updateDBList.add(updateDBSqlList);	
					
				    
				    sqlCmd.delete(0,sqlCmd.length());
					paramList = new ArrayList();					    
				    sqlCmd.append("SELECT * FROM M05 WHERE m_year=? AND m_month=?");				    
					paramList.add(S_YEAR);
				    paramList.add(S_MONTH);
					data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month");
				   
					if (data.size() == 0){
						errMsg = errMsg + "無資料可刪除("+Report_no+")<br>";
					}else{
					    updateDBDataList = new ArrayList();//儲存參數的List
					    updateDBSqlList = new ArrayList();//儲存參數的List	
					    dataList = new LinkedList();//儲存參數的data								
					    sqlCmd.delete(0,sqlCmd.length());
						sqlCmd.append(" INSERT INTO M05_LOG ");
						sqlCmd.append(" select m_year,m_month,loan_unit_no,period_no,item_no,");
						sqlCmd.append("        repay_cnt,repay_amt,run_notgood_cnt,run_notgood_amt,turn_out_cnt,");
						sqlCmd.append("        turn_out_amt,diease_cnt,dieaserepay_amt,disaster_cnt,disaster_amt,");
						sqlCmd.append("        corun_out_cnt,corun_out_amt,other_cnt,other_amt ");
						sqlCmd.append("        ,?,?,sysdate,'U'");
						sqlCmd.append("   from M05");
						sqlCmd.append(" WHERE m_year=? AND m_month=?");
						dataList.add(userid);					       
					    dataList.add(username);
					    dataList.add(S_YEAR);
					    dataList.add(S_MONTH);
					    updateDBDataList.add(dataList);
					    updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	    updateDBSqlList.add(updateDBDataList);
	            	    updateDBList.add(updateDBSqlList);	
	            	    
						//delete M05
						updateDBDataList = new ArrayList();//儲存參數的List
					    updateDBSqlList = new ArrayList();//儲存參數的List	
					    dataList = new LinkedList();//儲存參數的data								
					    sqlCmd.delete(0,sqlCmd.length());
						sqlCmd.append(" delete from M05 WHERE m_year=? AND m_month=?");
						dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						updateDBDataList.add(dataList);
					    updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	    updateDBSqlList.add(updateDBDataList);
	            	    updateDBList.add(updateDBSqlList);

					    sqlCmd.delete(0,sqlCmd.length());
					    sqlCmd.append("SELECT * FROM M05_TOTACC WHERE m_year=? AND m_month=?");
						data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month");
						if (data.size() == 0){
							errMsg = errMsg + "無資料可刪除("+Report_no+")<br>";
						}else{
						    updateDBDataList = new ArrayList();//儲存參數的List
					        updateDBSqlList = new ArrayList();//儲存參數的List	
					        dataList = new LinkedList();//儲存參數的data								
					    	sqlCmd.delete(0,sqlCmd.length());
							sqlCmd.append(" INSERT INTO M05_TOTACC_LOG ");
							sqlCmd.append(" select m_year,m_month,loan_unit_no,fix_no,guarantee_no_totacc,guarantee_amt_totacc");
							sqlCmd.append("        ,?,?,sysdate,'U'");
							sqlCmd.append("   from M05_TOTACC");
							sqlCmd.append(" WHERE m_year=? AND m_month=?");
							dataList.add(userid);					       
					        dataList.add(username);
					        dataList.add(S_YEAR);
					        dataList.add(S_MONTH);
					        updateDBDataList.add(dataList);
					        updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	        updateDBSqlList.add(updateDBDataList);
	            	        updateDBList.add(updateDBSqlList);	
	            	    
							//delete M05_TOTACC
							updateDBDataList = new ArrayList();//儲存參數的List
					        updateDBSqlList = new ArrayList();//儲存參數的List	
					        dataList = new LinkedList();//儲存參數的data								
					    	sqlCmd.delete(0,sqlCmd.length());
							sqlCmd.append(" delete from M05_TOTACC WHERE m_year=? AND m_month=?");
							dataList.add(S_YEAR);
					        dataList.add(S_MONTH);
					        updateDBDataList.add(dataList);
					        updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	        updateDBSqlList.add(updateDBDataList);
	            	        updateDBList.add(updateDBSqlList);	

						    sqlCmd.delete(0,sqlCmd.length());
						    sqlCmd.append("SELECT * FROM M05_NOTE WHERE m_year=? AND m_month=?");
							data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month");
							if (data.size() == 0){
								errMsg = errMsg + "無資料可刪除("+Report_no+")<br>";
							}else{
							    updateDBDataList = new ArrayList();//儲存參數的List
					            updateDBSqlList = new ArrayList();//儲存參數的List	
					            dataList = new LinkedList();//儲存參數的data								
					    	    sqlCmd.delete(0,sqlCmd.length());
								sqlCmd.append(" INSERT INTO M05_NOTE_LOG ");
								sqlCmd.append(" select m_year,m_month,note_no,note_amt_rate");
								sqlCmd.append("        ,?,?,sysdate,'U'");
								sqlCmd.append("   from M05_NOTE");
								sqlCmd.append(" WHERE m_year=? AND m_month=?");
								dataList.add(userid);					       
					            dataList.add(username);
					            dataList.add(S_YEAR);
					            dataList.add(S_MONTH);
					            updateDBDataList.add(dataList);
					            updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	            updateDBSqlList.add(updateDBDataList);
	            	            updateDBList.add(updateDBSqlList);	 
	            	            
	            	            //delete M05_NOTE
								updateDBDataList = new ArrayList();//儲存參數的List
					            updateDBSqlList = new ArrayList();//儲存參數的List	
					            dataList = new LinkedList();//儲存參數的data								
					    	    sqlCmd.delete(0,sqlCmd.length());
								sqlCmd.append(" delete from M05_NOTE WHERE m_year=? AND m_month=?");
								dataList.add(S_YEAR);
					            dataList.add(S_MONTH);
					            updateDBDataList.add(dataList);
					            updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
	            	            updateDBSqlList.add(updateDBDataList);
	            	            updateDBList.add(updateDBSqlList);	 

								if(DBManager.updateDB_ps(updateDBList)){
									errMsg = errMsg + "相關資料刪除成功";
								}else{
				   					errMsg = errMsg + "相關資料刪除失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
								}
							}
						}
					}
				}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料刪除失敗";
		}

		return errMsg;
	}
	
    //95.05.18 add get A99 by 2295
    //102.10.07 bank_type預設為6農會 by 2295
    //103.02.10 add 103/01以後.A06.套用新表格(增加/異動科目代號)  
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
			if(bank_type == null || "".equals(bank_type)){//102.10.07 bank_type預設為6農會
			   bank_type="6";
			}	
    		if(bank_type.equals("6") || (bank_type.equals("7") && !acc_div.equals("01") && !acc_div.equals("02") && !acc_div.equals("04") && !acc_div.equals("05") && !acc_div.equals("07") && !acc_div.equals("08") && !acc_div.equals("99") && !acc_div.equals("12") && !acc_div.equals("13"))){
    		   if(acc_div.equals("01") || acc_div.equals("02")){
    		      Report_NO="A01";
    		      //96.12.19 add 97/01以後,套用新表格(增加/異動科目代號)
    		      if(Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH) >= 9701){
    		         ncacno = "ncacno_rule";
    		      }
    		   }    		   
    		   else if(acc_div.equals("04"))                    Report_NO="A02";
    		   else if(acc_div.equals("05"))                    Report_NO="A03";
    		   else if(acc_div.equals("06"))		            Report_NO="A04";
    		   else if(acc_div.equals("07"))					Report_NO="A05";
    		   else if(acc_div.equals("08"))					Report_NO="A06";//95.04.10 add by 2295
    		   else if(acc_div.equals("99"))					Report_NO="A99";//95.05.18 add by 2295
    		  
    		   if(acc_div.equals("12") || acc_div.equals("13")){//104.10.12 add
    		      Report_NO="A13";    		     
    		   }    
    		  
    		  
    		   if(Report_NO.equals("A05")){
    		      //95.05.15 fix 修改a05 sql
    			  sqlCmd.append("SELECT * ");
                  sqlCmd.append("  FROM ncacno LEFT JOIN "+Report_NO+" a01_a05 on A01_A05.acc_code = ncacno.acc_code ");
                  sqlCmd.append("  				  AND (A01_A05.m_year=? OR A01_A05.m_year IS NULL) ");
                  sqlCmd.append("  				  AND (A01_A05.m_month=? OR A01_A05.m_month IS NULL) ");
                  sqlCmd.append("   				  AND (A01_A05.bank_code=? OR A01_A05.bank_code IS NULL) ");
                  sqlCmd.append("       LEFT JOIN A05_ASSUMED ON A05_ASSUMED.acc_code = ncacno.acc_code ");
                  sqlCmd.append(" WHERE NCACNO.acc_div=? ");
                  sqlCmd.append(" ORDER BY ncacno.acc_range");
                  paramList.add(S_YEAR);
                  paramList.add(S_MONTH);
                  paramList.add(bank_code);
                  paramList.add(acc_div);
            	  dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt,assumed");
               }else if(Report_NO.equals("A02") || Report_NO.equals("A99") ){//95.05.17 add A02農會以ncacno為主//95.05.18 add A99農會以ncacno為主
                  sqlCmd.append(" select * from ncacno LEFT JOIN "+Report_NO+" A01_A05  ");
                  sqlCmd.append(" ON A01_A05.acc_code = ncacno.acc_code and A01_A05.m_year=? and A01_A05.m_month=? and A01_A05.bank_code=?");
                  sqlCmd.append(" where ncacno.acc_div=? order by ncacno.acc_range");
                  paramList.add(S_YEAR);
                  paramList.add(S_MONTH);
                  paramList.add(bank_code);
                  paramList.add(acc_div);
            	  dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");               
               }else{
            	  sqlCmd.append(" select * from "+Report_NO+" A01_A05  LEFT JOIN "+ncacno+" ON A01_A05.acc_code = "+ncacno+".acc_code ");
            	  sqlCmd.append(" where A01_A05.m_year=? and A01_A05.m_month=? and A01_A05.bank_code=? ");
            	  sqlCmd.append(" and "+ncacno+".acc_div=? order by "+ncacno+".acc_range");
            	  paramList.add(S_YEAR);
                  paramList.add(S_MONTH);
                  paramList.add(bank_code);
                  paramList.add(acc_div);
            	  dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt,amt_3month,amt_6month,amt_1year,amt_2year,amt_over2year,amt_total");
               }
            }else{
               if(acc_div.equals("01") || acc_div.equals("02")) Report_NO="A01";
               else if(acc_div.equals("04"))                    Report_NO="A02";//94.05.27 add加上A02區分農漁會
    		   else if(acc_div.equals("05"))                    Report_NO="A03";
    		   else if(acc_div.equals("07"))                    Report_NO="A05";//95.05.11 add加上A05區分農漁會
    		   else if(acc_div.equals("08"))                    Report_NO="A06";
    		   else if(acc_div.equals("99"))                    Report_NO="A99";//95.05.18 add加上A99區分農漁會
    		   if(acc_div.equals("12") || acc_div.equals("13")){//104.10.12 add
    		      Report_NO="A13";    		     
    		   }    
    		   if(Report_NO.equals("A05")){//95.05.11 add A05漁會
    		      //95.05.15 fix 修改a05 sql
    			  sqlCmd.append("SELECT * ");
                  sqlCmd.append("  FROM ncacno_7 LEFT JOIN "+Report_NO+" a01_a05 on A01_A05.acc_code = ncacno_7.acc_code ");
                  sqlCmd.append("  				 AND (A01_A05.m_year=? OR A01_A05.m_year IS NULL) ");
                  sqlCmd.append("   				 AND (A01_A05.m_month=? OR A01_A05.m_month IS NULL) ");
                  sqlCmd.append("   				 AND (A01_A05.bank_code=? OR A01_A05.bank_code IS NULL) ");
                  sqlCmd.append("       LEFT JOIN A05_ASSUMED ON A05_ASSUMED.acc_code = ncacno_7.acc_code ");
                  sqlCmd.append(" WHERE NCACNO_7.acc_div=?");
                  sqlCmd.append(" ORDER BY ncacno_7.acc_range");
                  paramList.add(S_YEAR);
                  paramList.add(S_MONTH);
                  paramList.add(bank_code);
                  paramList.add(acc_div);       
            	  dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt,assumed");
               }else if(Report_NO.equals("A02") || Report_NO.equals("A99")){//95.05.17 add A02漁會以ncacno_7為主//95.05.18 add A99漁會以ncacno_7為主
                  sqlCmd.append(" select * from  ncacno_7 LEFT JOIN "+Report_NO+" A01_A05 ");
                  sqlCmd.append(" ON A01_A05.acc_code = ncacno_7.acc_code and A01_A05.m_year=? and A01_A05.m_month=?");
                  sqlCmd.append(" and A01_A05.bank_code=? where ncacno_7.acc_div=? order by ncacno_7.acc_range");
                  
                  paramList.add(S_YEAR);
                  paramList.add(S_MONTH);
                  paramList.add(bank_code);
                  paramList.add(acc_div);
            	  dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");
               }else{
                  if(Report_NO.equals("A01") || Report_NO.equals("A06")){//漁會
                     //102.04.18 add 103/01以後,A01.套用新表格(增加/異動科目代號)   
                     //103.02.10 add 103/01以後.A06.套用新表格(增加/異動科目代號)   
        		     if(Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH) >= 10301){
        		         ncacno_7 = "ncacno_7_rule";
        		     }
                  }
    		      sqlCmd.append(" select * from "+Report_NO+" A01_A05  LEFT JOIN "+ncacno_7);
    		      sqlCmd.append(" ON A01_A05.acc_code = "+ncacno_7+".acc_code ");
    		      sqlCmd.append(" where A01_A05.m_year=? and A01_A05.m_month=? and A01_A05.bank_code=?");
    		      sqlCmd.append(" and "+ncacno_7+".acc_div=? order by "+ncacno_7+".acc_range");
    		      paramList.add(S_YEAR);
                  paramList.add(S_MONTH);
                  paramList.add(bank_code);
                  paramList.add(acc_div);
                  dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt,amt_3month,amt_6month,amt_1year,amt_2year,amt_over2year,amt_total");
               }
            }

            return dbData;
    }
    //96.07.10 add get A08 by 2295
    //96.12.03 add get A09 by 2295
    //97.06.13 add get A10 by 2295
    //99.10.01 add 合併A08/A09/A10 by 2295
    //104.01.08 add A12
    private List getData_A08_A12(String S_YEAR,String S_MONTH,String bank_code,String report_no){
    		//查詢條件
    		List dbData =null;
    		StringBuffer sqlCmd = new StringBuffer();
			List paramList = new ArrayList();	
			sqlCmd.append("select * from "+report_no+" where m_year=? and m_month=? and bank_code=?");
			paramList.add(S_YEAR);
            paramList.add(S_MONTH);
            paramList.add(bank_code);
            if(report_no.equals("A08")){
    		   dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,warnaccount_cnt,limitaccount_cnt,erroraccount_cnt,otheraccount_cnt,depositaccount_tcnt");
    		}else if(report_no.equals("A09")){ 
    		   dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,over_cnt,over_amt,push_over_amt,totalamt,push_totalamt,over_total_rate");
    		}else if(report_no.equals("A10")){ 
    		   dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,
    				     "m_year,m_month,loan1_amt,loan2_amt,loan3_amt,loan4_amt,invest1_amt,invest2_amt,invest3_amt,invest4_amt,other1_amt,other2_amt,other3_amt,other4_amt"
    				    +",loan1_baddebt,loan2_baddebt,loan3_baddebt,loan4_baddebt,build1_baddebt,build2_baddebt,build3_baddebt,build4_baddebt"
						+",baddebt_flag,baddebt_noenough,baddebt_delay,baddebt_104,baddebt_105,baddebt_106,baddebt_107,baddebt_108");
    		}else if(report_no.equals("A12")){ 
     		   dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,bank_code,baddebt_amt,loss_amt,profit_amt");
    		}
            return dbData;
    }
    
    //Method modify by jei 93.12.14
	private List getData_B01(String S_YEAR,String S_MONTH,String bank_code){
   		//查詢條件
   	    List dbData =null;
    	StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		
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
		
        return dbData;
    }

    //Method modify by jei 93.12.16
	private List getData_B02(String S_YEAR,String S_MONTH,String bank_code){
   		//查詢條件
   	    List dbData =null;
    	StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		
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
			sqlCmd.append(" and 		substr(c.data_range,4,2)= b.data_range ");
			sqlCmd.append(" and 		b.report_no             = 'M01' ");
			sqlCmd.append(" and 		c.m_year                = ?");
			sqlCmd.append(" and 		c.m_month               = ?");
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
		sqlCmd.append("        and substr(c.data_range,4,2)=b.data_range ");
		sqlCmd.append("        and b.report_no = 'M02' ");
		sqlCmd.append("        and c.m_year = ?");
		sqlCmd.append("        and c.m_month = ?");
		sqlCmd.append(" order by a.input_order,b.input_order") ;
		paramList.add(S_YEAR);    
		paramList.add(S_MONTH);
		dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,guarantee_cnt,loan_amt,guarantee_amt,loan_bal,guarantee_bal,over_notpush_cnt,over_notpush_bal,over_okpush_cnt,over_okpush_bal,repay_tot_cnt,repay_tot_amt,repay_bal_cnt,repay_bal_amt");

		return dbData;
	}

	//Method modify by egg 93.12.12
	private List getData_M03(String S_YEAR,String S_MONTH,String bank_code,String data_range_type){
		
    	//查詢條件
    	List dbData =null;
    	StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
    	if(data_range_type.equals("C")){	// M03的資料
			sqlCmd.append(" select 		M03.div_no,a.data_range,a.data_range_name,M03.guarantee_cnt_month, ");
			sqlCmd.append("        	M03.loan_amt_month, M03.guarantee_amt_month, ");
			sqlCmd.append("        	M03.guarantee_cnt_year,M03.loan_amt_year,M03.guarantee_amt_year, ");
			sqlCmd.append("        	M03.guarantee_bal_totacc,M03.guarantee_bal_totacc_over,M03.repay_bal_totacc ");
			sqlCmd.append(" from 		M03,m00_data_range_item a");
			sqlCmd.append(" where 	M03.div_no=a.data_range ");
			sqlCmd.append(" and		a.data_range_type = 'C' ");
			sqlCmd.append(" and       a.report_no ='M03' ");
			sqlCmd.append(" and 		M03.m_year=? and M03.m_month=?");
			sqlCmd.append(" order 	by a.input_order ");
			paramList.add(S_YEAR);    
			paramList.add(S_MONTH);
			dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,guarantee_cnt_month,loan_amt_month,guarantee_amt_month,guarantee_cnt_year,loan_amt_year,guarantee_amt_year,guarantee_bal_totacc,guarantee_bal_totacc_over,repay_bal_totacc");
		}else if(data_range_type.equals("S")){	// M03_NOTE的資料
			sqlCmd.append(" select m_year,m_month,note_no as \"data_range\",note_amt_rate ");
			sqlCmd.append(" from M03_note,m00_data_range_item ");
			sqlCmd.append(" where note_no=data_range ");
			sqlCmd.append("  and data_range_type='S' ");
			sqlCmd.append("  and m_year=?");
			sqlCmd.append("  and m_month=?");
			sqlCmd.append(" order by input_order ");
			paramList.add(S_YEAR);    
			paramList.add(S_MONTH);
			dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,note_amt_rate");
		}
		
		return dbData;
	}

	//jei 931212
	private List getData_M04(String S_YEAR,String S_MONTH,String bank_code){
		//查詢條件
		List dbData =null;
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		
		sqlCmd.append(" select b.loan_use_no,a.loan_use_name,guarantee_no_month,guarantee_no_month_p, ");
		sqlCmd.append("        loan_amt_month,loan_amt_month_p,guarantee_amt_month,guarantee_amt_month_p,guarantee_no_year, ");
		sqlCmd.append("        guarantee_no_year_p,loan_amt_year,loan_amt_year_p,guarantee_amt_year, ");
		sqlCmd.append("        guarantee_amt_year_p,guarantee_no_totacc,guarantee_no_totacc_p, ");
		sqlCmd.append("        loan_amt_totacc,loan_amt_totacc_p,guarantee_amt_totacc,guarantee_amt_totacc_p ");
		sqlCmd.append(" from   m00_loan_use a,M04 b ");
		sqlCmd.append(" where  b.loan_use_no = a.loan_use_no ");
		sqlCmd.append("        and b.m_year = ?");
		sqlCmd.append("        and b.m_month = ?");
		sqlCmd.append(" order by a.input_order");
		paramList.add(S_YEAR);    
		paramList.add(S_MONTH);
		dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,guarantee_no_month,guarantee_no_month_p,loan_amt_month,loan_amt_month_p,guarantee_amt_month,guarantee_amt_month_p,guarantee_no_year,guarantee_no_year_p,loan_amt_year,loan_amt_year_p,guarantee_amt_year,guarantee_amt_year_p,guarantee_no_totacc,guarantee_no_totacc_p,loan_amt_totacc,loan_amt_totacc_p,guarantee_amt_totacc,guarantee_amt_totacc_p");

		return dbData;
	}

	private List getData_M05(String S_YEAR,String S_MONTH,String bank_code,String data_range_type){
    		//查詢條件
    		List dbData =null;
    		StringBuffer sqlCmd = new StringBuffer();
			List paramList = new ArrayList();	
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
				sqlCmd.append("   and m05.period_no || m05.item_no = m00_data_range_item.data_range ");
				sqlCmd.append("   and m05.m_year=? and m05.m_month=?");
				sqlCmd.append(" order by m00_loan_unit.input_order ,m00_data_range_item.input_order ");
				paramList.add(S_YEAR);    
				paramList.add(S_MONTH);
				dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,repay_cnt,repay_amt,run_notgood_cnt,run_notgood_amt,turn_out_cnt,turn_out_amt,diease_cnt,dieaserepay_amt,disaster_cnt,disaster_amt,corun_out_cnt,corun_out_amt,other_cnt,other_amt");
			}else if(data_range_type.equals("N")){	// M05_NOTE的資料
				sqlCmd.append(" select m_year,m_month,note_no,note_amt_rate ");
				sqlCmd.append(" from m05_note,m00_data_range_item ");
				sqlCmd.append(" where note_no=data_range ");
				sqlCmd.append("  and data_range_type='N' ");
				sqlCmd.append("  and m_year=?");
				sqlCmd.append("  and m_month=?");
				sqlCmd.append(" order by input_order ") ;
				paramList.add(S_YEAR);    
				paramList.add(S_MONTH);
				dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,note_amt_rate");
			}else{ 	//M05_totacc
				sqlCmd.append(" select m_year,m_month,m05_totacc.loan_unit_no,loan_unit_name, ");
				sqlCmd.append("       fix_no,guarantee_no_totacc,guarantee_amt_totacc ");
				sqlCmd.append(" from m05_totacc,m00_loan_unit ");
				sqlCmd.append(" where m05_totacc.loan_unit_no=m00_loan_unit.loan_unit_no ");
				sqlCmd.append("  and m_year=?");
				sqlCmd.append("  and m_month=?");
				sqlCmd.append(" order by input_order ");
				paramList.add(S_YEAR);    
				paramList.add(S_MONTH);
				dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,guarantee_no_totacc,guarantee_amt_totacc");
			}
            return dbData;
	}

	//Method modify by egg 93.12.14
	private List getData_M06_M07(String S_YEAR,String S_MONTH,String bank_code,String Report_no){
		List dbData =null;
    	StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();

    	// 取得M06(M07)的資料
		sqlCmd.append(" select * from " + Report_no + ",m00_area where " + Report_no + ".area_no=m00_area.area_no and m_year=? and m_month=?");
		sqlCmd.append(" and m00_area.area_no <> '1' order by input_order ");
		paramList.add(S_YEAR);    
		paramList.add(S_MONTH);
		dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,guarantee_no_month,guarantee_amt_month,loan_amt_month,guarantee_no_year,guarantee_amt_year,loan_amt_year,guarantee_no_totacc,guarantee_amt_totacc,loan_amt_totacc,guarantee_bal_no,guarantee_bal_amt,guarantee_bal_p,loan_bal");
		
		return dbData;
	}

	//Method modify by egg 93.12.16
	private List getData_M08(String S_YEAR,String S_MONTH,String bank_code,String Report_no){
		List dbData =null;
    	StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	

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
		paramList.add(S_YEAR);    
		paramList.add(S_MONTH);
		dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,guarantee_no_month,loan_amt_month,guarantee_amt_month,guarantee_bal_month,guarantee_bal_p");

		return dbData;
	}

	//Method modify by egg 93.12.15
	private List getData_B03(String S_YEAR,String S_MONTH,String bank_code,String Report_no,int form_no){
		List dbData =null;
    	StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	

    	// 取得B03的資料
    	if(form_no == 1){	// B03_1的資料
    		System.out.println("form_no=1");
    		sqlCmd.append(" select * from " + Report_no + "_1 a,b00_funs_item b where a.funs_master_no=b.funs_master_no ");
    		sqlCmd.append(" and   a.funs_sub_no=b.funs_sub_no and   a.funs_next_no=b.funs_next_no ");
    		sqlCmd.append(" and   m_year=? and   m_month=? order by input_order");
			paramList.add(S_YEAR);    
			paramList.add(S_MONTH);
			dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,loan_cnt_totacc,loan_amt_totacc_fund,loan_amt_totacc_bank,loan_amt_totacc_tot,loan_cnt_bal,loan_amt_bal_fund,loan_amt_bal_bank,loan_amt_bal_tot");
		}else if(form_no == 2){	// B03_2的資料
			System.out.println("form_no=2");
			sqlCmd.append(" select * from " + Report_no + "_2 a,b00_funs_item b where a.funs_master_no=b.funs_master_no ");
    		sqlCmd.append(" and   a.funs_sub_no=b.funs_sub_no and   a.funs_next_no=b.funs_next_no ");
    		sqlCmd.append(" and   m_year=? and   m_month=? order by input_order");
			paramList.add(S_YEAR);    
			paramList.add(S_MONTH);
			dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,loan_amt_bal,loan_amt_over,loan_rate_over");
		}else if(form_no == 3){	// B03_3的資料
			System.out.println("form_no=3");
			sqlCmd.append(" select * from " + Report_no + "_3 a,b00_funo_item b where a.funo_master_no=b.funo_master_no ");
    		sqlCmd.append(" and   a.funo_sub_no=b.funo_sub_no and   a.funo_next_no=b.funo_next_no ");
    		sqlCmd.append(" and   m_year=? and   m_month=? order by input_order");
			paramList.add(S_YEAR);    
			paramList.add(S_MONTH);
			dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,funo_master_no,funo_sub_no,funo_next_no,funo_amt,funo_rate");
		}else if(form_no == 4){	// B03_4的資料
			System.out.println("form_no=4");
			sqlCmd.append(" select * from " + Report_no + "_4 a,b00_bank_no b where a.bank_no=b.bank_no ");
    		sqlCmd.append(" and   m_year=? and   m_month=? order by input_order");
			paramList.add(S_YEAR);    
			paramList.add(S_MONTH);
			dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,machine_cnt,machine_amt,land_cnt,land_amt,house_cnt,house_amt,build_cnt,build_amt,tot_cnt,tot_amt");
		}

		
		return dbData;
	}

	private List getData_F01(String S_YEAR,String S_MONTH,String bank_code){
   		//查詢條件
   	    List dbData =null;
    	StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		sqlCmd.append(" select dep_type,acct_type,acct_cnt_tm,bal_lm,dep_tm,wtd_tm,bal_tm");
		sqlCmd.append(" from   F01 ");
		sqlCmd.append(" where  m_year = ?");
		sqlCmd.append(" and    m_month = ?");
		sqlCmd.append(" and    bank_code = ?");
		sqlCmd.append(" order by dep_type,acct_type ");
		paramList.add(S_YEAR);    
		paramList.add(S_MONTH);
		paramList.add(bank_code);
		dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"acct_cnt_tm,bal_lm,dep_tm,wtd_tm,bal_tm");
        return dbData;
    }
	
    //94.02.21 add 檢查檔案有無被鎖住 by 2295
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
    		
    		sqlCmd.delete(0,sqlCmd.length());
			sqlCmd.append("select lock_status from WML01 where m_year=? and m_month=? and bank_code=? and report_no=?");			
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
    	   System.out.println("CheckFileLock ="+lock);
           return lock;
    }
    //94.11.07 add 取得WLX_APPLY_INI農金局設定的起始年月 by 2295
    private String[] getWLX_APPLY_INI(String Report_no){//取得WLX_APPLY_INI農金局設定的起始年月
    		   StringBuffer sqlCmd = new StringBuffer();
			   List paramList = new ArrayList();	
    	       String INI[] = new String[2];
    	       sqlCmd.append("select m_year,m_month from WLX_APPLY_INI where REPORT_NO = ?");
    	       paramList.add(Report_no);
     	       List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month");
     	       System.out.println("WLX_APPLY_INI.size()="+dbData.size());
     	       if(dbData != null && ((DataObject)dbData.get(0)) != null){
     	          INI[0] = (((DataObject)dbData.get(0)).getValue("m_year")).toString();
     	          INI[1] = (((DataObject)dbData.get(0)).getValue("m_month")).toString();
     	       }
     	       return INI;
    }
    //94.11.07 add 檢查前一個月是否有申報成功的資料 by 2295
    //			   檢查後一個月是否有申報的資料
    private boolean checkExistMonth(String S_YEAR,String S_MONTH,String bank_code,String code,String report_no){
            String temp_year="",temp_month="";
            StringBuffer sqlCmd = new StringBuffer();
			List paramList = new ArrayList();
			
            if(code.equals("LAST")){
               if(S_MONTH.equals("01")){//若畫面輸入的月分為1月份..則是申報上個年度的12月份
      		       temp_year = String.valueOf(Integer.parseInt(S_YEAR) - 1);
       		       temp_month = "12";
    		   }else{
    		       temp_year = S_YEAR;
      		       temp_month = String.valueOf(Integer.parseInt(S_MONTH) - 1);//申報上個月份的
    		   }
            }else if(code.equals("NEXT")){
              temp_month = String.valueOf(Integer.parseInt(S_MONTH)+1);
              if(temp_month.equals("13")){
      		     temp_year = String.valueOf(Integer.parseInt(S_YEAR) + 1);
       		     temp_month = "1";
    		  }else{
    		     temp_year=S_YEAR;
    		  }
            }

			
    		sqlCmd.append(" select count(*)  as  cnt ");
        	sqlCmd.append(" from WML01  aa");
			sqlCmd.append(" where aa.M_YEAR  = ?");
			sqlCmd.append("   and aa.M_MONTH = ?");
			sqlCmd.append("   and aa.BANK_CODE  = ?");
			sqlCmd.append("   and aa.REPORT_NO = ?");
			paramList.add(temp_year);			  
			paramList.add(temp_month);
			paramList.add(bank_code);
			paramList.add(report_no);
			if(code.equals("LAST")){
			   sqlCmd.append("   and (aa.UPD_CODE ='U' or aa.UPD_CODE ='Z')");//fix 若上月資料皆為0時,可申報下月資料
			}
			List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"cnt");
    		if(dbData != null && Integer.parseInt((((DataObject)dbData.get(0)).getValue("cnt")).toString()) > 0 ){
    		   System.out.println("有申報成功的"+code+"資料");
     	       return true;//有申報成功的上月份資料//有下個月的申報資料
     	    }else{
     	       System.out.println("無申報成功的"+code+"資料");
     	       return false;//無申報成功的上月份資料//無下個月的申報資料
     	    }
    }
    //94.11.14 add 上月資料匯入 by 2295
    private List getLastMonthData(String S_YEAR,String S_MONTH,String bank_code,String Report_no,String bank_type){
     		 StringBuffer sqlCmd = new StringBuffer();
			 List paramList = new ArrayList();
			 
             if(S_MONTH.equals("1") || S_MONTH.equals("01")){//若本月為1月份是..則是抓上個年度的12月份
       			S_YEAR = String.valueOf(Integer.parseInt(S_YEAR) - 1);
       			S_MONTH = "12";
    		 }else{
      			S_MONTH = String.valueOf(Integer.parseInt(S_MONTH) - 1);//申報上個月份的
    		 }
    		 List dbData = null;
    		
    		 if(Report_no.equals("F01")){
    	        sqlCmd.append(" select dep_type,acct_type,acct_cnt_tm,bal_tm as bal_lm,0 as dep_tm,0 as wtd_tm, bal_tm");
		       	sqlCmd.append(" from   "+Report_no);
		       	sqlCmd.append(" where  m_year = ?");
		       	sqlCmd.append(" and    m_month = ?");
		       	sqlCmd.append(" and    bank_code =?");
		       	sqlCmd.append(" order by dep_type,acct_type ");
		       	paramList.add(S_YEAR);
		       	paramList.add(S_MONTH);
		       	paramList.add(bank_code);
			    dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"acct_cnt_tm,bal_lm,dep_tm,wtd_tm,bal_tm");
			 }else if(Report_no.equals("A06")){
			    dbData = getData_A01_A05(S_YEAR,S_MONTH,bank_code,bank_type,"08");
			 }else if(Report_no.equals("A08") || Report_no.equals("A10")){//96.07.10 add A08 //97.06.13 add A10
			    dbData = getData_A08_A12(S_YEAR,S_MONTH,bank_code,Report_no);			 
			 }
     	     System.out.println("LastMonthData.size()="+dbData.size());
     	     return dbData;
    }    
%>
