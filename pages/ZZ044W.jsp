<%
// 99.12.17 add 金庫BIS資料轉檔作業 by 2295
//100.08.17 fix 增加100年度的22張報表 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.*" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.transfer.*" %>
<%@ page import="com.tradevan.util.ftp.MyFTPClient" %>
<%@ page import="java.util.*,java.io.*" %>
<%@include file="./include/Header.include" %>
<%
	String S_YEAR = Utility.getTrimString(dataMap.get("S_YEAR"));
	String S_MONTH = Utility.getTrimString(dataMap.get("S_MONTH"));
	String E_YEAR = Utility.getTrimString(dataMap.get("E_YEAR"));
	String E_MONTH = Utility.getTrimString(dataMap.get("E_MONTH"));
	String acc_tr_type = Utility.getTrimString(dataMap.get("acc_tr_type"));	
	String firstStatus = Utility.getTrimString(dataMap.get("firstStatus"));
 	String rptIP=Utility.getProperties("rptIP");
	String rptID=Utility.getProperties("rptID");
	String rptPwd=Utility.getProperties("rptPwd");
    if(!Utility.CheckPermission(request,report_no)){//無權限時,導向到LoginError.jsp
        rd = application.getRequestDispatcher( LoginErrorPgName );
    }else{
    	//set next jsp
    	List BISList = getBISList();//BIS報表名稱
    	request.setAttribute("BISList",BISList);
    	if(act.equals("Qry")){
    	    rd = application.getRequestDispatcher( ListPgName +"?act=Qry&firstStatus="+firstStatus);    	    
    	}else if(act.equals("List")){
    	    List dbData = getQryResult(S_YEAR,S_MONTH,E_YEAR,E_MONTH,acc_tr_type,act);
    	    request.setAttribute("reportList",dbData);
    	    rd = application.getRequestDispatcher( ListPgName +"?act=Qry&acc_tr_type="+acc_tr_type+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&E_YEAR="+E_YEAR+"&E_MONTH="+E_MONTH);
    	}else if(act.equals("parseAgriBankRpt") || act.equals("parseALLAgriBankRpt")){
    	    //下載已上傳至農金局DB主機的金庫BIS報表檔案
    	    StringBuffer sqlCmd = new StringBuffer();
			List paramList = new ArrayList();
			RPT_CODE = "AgriRpt03";//BIS報表
    		sqlCmd.append(" select agrirpt_dirf.m_year,agrirpt_dirf.m_month,");
    		sqlCmd.append("       (to_char(agrirpt_dirf.m_year)|| LPAD(to_char(agrirpt_dirf.m_month),2,'0')) as s_yymm, ");
	   		sqlCmd.append(" 		  agrirpt_dirf.rpt_fname,agrirpt_dirf.rpt_version,rpt_nof.RPT_NAME,feb_flag, ");
	   		sqlCmd.append("	  	 ((to_char(agrirpt_dirf.UPDATE_DATE,'yyyymmdd')-19110000)  ||  to_char(agrirpt_dirf.UPDATE_DATE,' hh24:mi'))  as s_UpdateDate,");
	   		sqlCmd.append("		  decode(agrirpt_dirf.uploaddate,null,'',((to_char(agrirpt_dirf.uploaddate,'yyyymmdd')-19110000)  ||  to_char(agrirpt_dirf.uploaddate,' hh24:mi')))  as s_uploaddate ");
			sqlCmd.append(" from agrirpt_dirf left join rpt_nof on agrirpt_dirf.RPT_CODE = rpt_nof.RPT_CODE ");
			sqlCmd.append(" where agrirpt_dirf.rpt_code = ?");
			sqlCmd.append(" and to_char(agrirpt_dirf.m_year * 100 + agrirpt_dirf.m_month) = ?");
			sqlCmd.append(" order by rpt_version desc");
			paramList.add(RPT_CODE);//BIS
			paramList.add(S_YEAR+S_MONTH);
			
    		List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
    		if(dbData != null && dbData.size() != 0){//已有上傳至農金局主機的BIS報表
    	       ver = (String)((DataObject)dbData.get(0)).getValue("rpt_version");//取得最新版本
    	       
    	       File ClientRptDir = new File(Utility.getProperties("AgriBank_transferDir"));
	           if(!ClientRptDir.exists()){
         		   if(!Utility.mkdirs(Utility.getProperties("AgriBank_transferDir"))){
         	     		actMsg=actMsg+Utility.getProperties("AgriBank_transferDir")+"目錄新增失敗";
         		   }
        	   }
        	   if(actMsg.equals("")){        	     
        	      List filename_List = new LinkedList();
        	      filename_List.add(Utility.toBig5Convert((String)((DataObject)dbData.get(0)).getValue("rpt_fname")));

    	       	  MyFTPClient ftpC = new MyFTPClient(rptIP, rptID, rptPwd);
    	       	  System.out.println("AgriBankDir_remote="+Utility.getProperties("serverRptDir")+Utility.getProperties("AgriBankDir_remote")+"/"+RPT_CODE+"/"+S_YEAR+S_MONTH+"/"+ver);
    	       	  System.out.println("AgriBank_transferDir="+Utility.getProperties("AgriBank_transferDir")+System.getProperty("file.separator"));
    	       	  //MyFTPClient.getFiles(String remote_path, String local_path,List filename)
               	  actMsg = ftpC.getFiles(Utility.getProperties("serverRptDir")+Utility.getProperties("AgriBankDir_remote")+"/"+RPT_CODE+"/"+S_YEAR+S_MONTH+"/"+ver, Utility.getProperties("AgriBank_transferDir")+System.getProperty("file.separator"),filename_List);
               	  ftpC=null;  
               	  if(actMsg == null){
               	     actMsg = Utility.CopyFile(Utility.getProperties("AgriBank_transferDir")+System.getProperty("file.separator")+(String)((DataObject)dbData.get(0)).getValue("rpt_fname"),Utility.getProperties("AgriBank_transferDir")+System.getProperty("file.separator")+"BIS_"+S_YEAR+S_MONTH+".xls");
               	     if(actMsg.equals("0")){//檔案copy成功
               	        actMsg = "";
               	        File tmp = new File(Utility.getProperties("AgriBank_transferDir")+System.getProperty("file.separator")+(String)((DataObject)dbData.get(0)).getValue("rpt_fname"));
               	        tmp.delete();
               	        parseAgriBankRpt a = new parseAgriBankRpt();	    	
               	        String BISMsg = "";
               	        if(act.equals("parseALLAgriBankRpt")){//全部轉檔
               	           for(int i=0;i<BISList.size();i++){
               	               BISMsg = "";
               	               BISMsg = a.doParserRpt(S_YEAR,S_MONTH,(String)((DataObject)BISList.get(i)).getValue("identify_no"));               			
               	               if(BISMsg.equals("")){
               	                  actMsg += (String)((DataObject)BISList.get(i)).getValue("identify_no")+",";
               	               }
               	           }
               	           dbData = getQryResult(S_YEAR,S_MONTH,E_YEAR,E_MONTH,acc_tr_type,act);
               			   if(dbData != null && dbData.size() > 0){
                  			  if(Integer.parseInt(((DataObject)dbData.get(0)).getValue("total").toString()) > 0){
                     		     actMsg += "金庫BIS報表資料轉檔完成!!";
                  		      }
               			   }else{
                  			  actMsg = "金庫BIS報表資料轉檔失敗!!";
               			   }
    	                }else{    	                
               			   a.doParserRpt(S_YEAR,S_MONTH,acc_tr_type);               			
               			   dbData = getQryResult(S_YEAR,S_MONTH,E_YEAR,E_MONTH,acc_tr_type,act);
               			   if(dbData != null && dbData.size() > 0){
                  			  if(Integer.parseInt(((DataObject)dbData.get(0)).getValue("total").toString()) > 0){
                     		     actMsg = acc_tr_type+"金庫BIS報表資料轉檔完成!!";
                  		      }
               			   }else{
                  			  actMsg =  acc_tr_type+"金庫BIS報表資料轉檔失敗!!";
               			   }
               			}
               			
               			out.print(actMsg);
               	     }               	     
               	  }
               	  System.out.println("actMsg = "+actMsg);            	  
               }     	
    		   
    	       rd = application.getRequestDispatcher( nextPgName);
    	    }
    	}
    	request.setAttribute("actMsg",actMsg);
    }

%>

<%@include file="./include/Tail.include" %>

<%!
    private final static String report_no = "ZZ044W";
    private final static String nextPgName = "/pages/ActMsg.jsp";
    private final static String ListPgName = "/pages/"+report_no+"_List.jsp";
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    private String RPT_CODE="";
    private String ver = "";
    
    //取得查詢結果
    private List getQryResult(String S_YEAR,String S_MONTH,String E_YEAR,String E_MONTH,String acc_tr_type,String act){
    		//查詢條件
    		List paramList = new ArrayList();
    		StringBuffer sqlCmd = new StringBuffer();

    		sqlCmd.append(" select m_year,m_month,(to_char(m_year)|| LPAD(to_char(m_month),2,'0')) as s_yymm,");
            sqlCmd.append(" decode(acc_tr_type,'1-A1','(1-A1)合格自有資本與風險性資產比率計算表',");
            sqlCmd.append(" '1-B','(1-B)自有資本計算表',");
            sqlCmd.append(" '1-B1','(1-B1)資本扣除項目總表',");
            sqlCmd.append(" '1-C','(1-C)信用風險加權風險性資產、作業風險暨市場風險資本計提彙總表',");
            sqlCmd.append(" '2-A','(2-A)信用風險加權風險性資產額彙總表',");
            sqlCmd.append(" '2-B','(2-B)信用風險加權風險性資產額計算總表',");
            sqlCmd.append(" '2-F','(2-F)信用風險標準法資本扣除項目彙總表',");
            sqlCmd.append(" '4-A-1','(4-A-1)資產證券化加權風險性資產計算總表',");
            sqlCmd.append(" '4-H','(4-H)資產證券化資本扣除項目彙總表',");
            sqlCmd.append(" 'R0521','(R0521)自有資本與風險性資產比率計算表(銀行本身及合併)',"); 
            sqlCmd.append(" '6-B1','(6-B1)權益證券風險─個別風險之資本計提計算表(國家別)',");  
            sqlCmd.append(" '6-B2','(6-B2)權益證券風險─一般市場風險之資本計提計算表(國家別)',");  
            sqlCmd.append(" '6-C','(6-C)外匯(含黃金)風險－市場風險應計提資本彙總表',");  
            sqlCmd.append(" '6-C1','(6-C1)外匯(含黃金)風險－各幣別淨部位彙總表',");  
            sqlCmd.append(" '6-C2','(6-C2)外匯(含黃金)風險－各幣別淨部位彙總表',");  
            sqlCmd.append(" '6-E','(6-E)選擇權採用簡易法計提資本計算表',");  
            sqlCmd.append(" '6-G','(6-G)市場風險資本扣除項目彙總表',");  
            sqlCmd.append(" '2-C','(2-C)表內交易之信用風險加權風險性資產額計算表',");  
            sqlCmd.append(" '2-D','(2-D)一般表外交易之信用風險加權風險性資產額計算表',");  
            sqlCmd.append(" '2-E','(2-E)交易對手信用風險加權風險性資產額計算表',");  
            sqlCmd.append(" '2-E1','(2-E1)交易對手信用風險暴險額或違約暴險額計算表',");  
            sqlCmd.append(" '2-E2','(2-E2)交易對手信用風險信用相當額計算表',");  
            sqlCmd.append(" '4-D-1','(4-D-1)表外項目信用約當金額－標準法(非創始銀行)',");  
            sqlCmd.append(" '4-D-2','(4-D-2)表外項目信用約當金額－標準法(創始銀行)',");  
            sqlCmd.append(" '5-A','(5-A)作業風險之資本計提計算表',");  
            sqlCmd.append(" '6-A','(6-A)利率風險-市場風險應計提資本彙總表',");  
            sqlCmd.append(" '6-A1(NTD)','(6-A1)利率風險-個別風險之資本計提計算表(新台幣)',");  
            sqlCmd.append(" '6-A1(USD)','(6-A1)利率風險-個別風險之資本計提計算表(美元)',");  
            sqlCmd.append(" '6-A2-a(NTD)','(6-A2-a)利率風險-一般市場風險之資本計提計算表(到期法)(新台幣)',");  
            sqlCmd.append(" '6-A2-a(USD)','(6-A2-a)利率風險-一般市場風險之資本計提計算表(到期法)(美元)',");  
            sqlCmd.append(" '6-B','(6-B)權益證券風險─市場風險應計提資本彙總表',");   
            sqlCmd.append(" acc_tr_type) as s_report_name");
            sqlCmd.append(" ,count(*) as total");
            sqlCmd.append(" from agribank_rpt"); 
            if(act.equals("parseAgriBankRpt") || act.equals("parseALLAgriBankRpt")){
               sqlCmd.append(" where to_char(m_year * 100 + m_month) = ?");			   
			   paramList.add(S_YEAR+S_MONTH);
            }else{
               sqlCmd.append(" where (m_year * 100 + m_month) >= ?");
			   sqlCmd.append(" and (m_year * 100 + m_month) <= ?");
			   paramList.add(S_YEAR+S_MONTH);
			   paramList.add(E_YEAR+E_MONTH);
			}
            if(!acc_tr_type.equals("ALL")){
              sqlCmd.append(" and  acc_tr_type=?");
              paramList.add(acc_tr_type);
            }
            sqlCmd.append(" group by acc_tr_type,m_year,m_month");
            sqlCmd.append(" order by m_year desc,m_month desc,acc_tr_type asc");

    		List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"total");
    		if(dbData != null && dbData.size() != 0){
    		   System.out.println("dbData.size="+dbData.size());
    		}else{
    		   System.out.println("dbData is null or size = 0");
    		}

			return dbData;
    }

    //取得BIS報表名稱
    private List getBISList(){
    		//查詢條件
    		List paramList = new ArrayList();
    		StringBuffer sqlCmd = new StringBuffer();

    		sqlCmd.append(" select identify_no,cmuse_name");
            sqlCmd.append(" from cdshareno ");
            sqlCmd.append(" where cmuse_div=?");
            sqlCmd.append(" order by identify_no");
	        paramList.add("036");	   
	        DataObject bean = null;	   
            List acc_tr_type_List = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");	
    		
    		if(acc_tr_type_List != null && acc_tr_type_List.size() != 0){
    		   System.out.println("acc_tr_type_List.size="+acc_tr_type_List.size());
    		}else{
    		   System.out.println("acc_tr_type_List is null or size = 0");
    		}

			return acc_tr_type_List;
    }
%>    