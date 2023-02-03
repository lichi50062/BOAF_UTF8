//94.04.20 fix 調整長度 by 2295
//94.04.22 add 百分比比對的區間在-1 ~ 1(1000)之間都OK by 2295
//         fix 若都為"0"值時,顯示檢核為"0" by 2295
//99.11.17 寫入WML03的參數移至共用create_dataList;add ALL SQL 套用DAO.preparestatment,並列印轉換後的SQL by 2295
//99.11.19 add 移至共用Utility.getWML03_count讀取WML03 count(*)
//					  Utility.getWML01讀取WML01 all data
//					  Utility.Insert_UpdateWML01當WML01不存在Insert,存在時Update
//					  Utility.deleteWML01_UPLOAD刪除上傳檔案紀錄 by 2295

package com.tradevan.util;

import java.io.*;
import java.util.*;
import java.text.SimpleDateFormat;
import com.tradevan.util.dao.DataObject;


public class UpdateM04{
	private static String errMsg = "";
	
	public String getErrMsg(){
		return errMsg;  
	}
	public static synchronized String doParserReport_M04(String report_no, String m_year,
		String m_month,String filename, String srcbank_code,String upd_method,String input_method,String bank_type,String szuser_id,String szuser_name,int batch_no) throws Exception {
		//boolean parserResult=true;
		errMsg = "";
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();		
		int		errCount = 0;		
		String  user_id="";//本次異動者帳號
		String  user_name="";//本次異動者姓名
		String[] user_data = {"",""};
		String  bank_code=srcbank_code;	
		String  add_date="";//申報日期
		String  upd_code="";//檢核結果
		String  common_center="";//由共用中心傳入
		boolean updateOK=false;
		SimpleDateFormat bkformat = new SimpleDateFormat("yyyyMMddHHmmss");
		SimpleDateFormat emailformat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");	
		String WMdataDir = Utility.getProperties("WMdataDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");
		String WMdataBKDir = Utility.getProperties("WMdataBKDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");		
		
		List M04List = new LinkedList();//M04細部資料
		List dbData = null;//其他querydb後,資料暫存的list
		String checkResult="true";
		boolean nonZero=false;//94.04.22 檢查M04的內容是否都為"0"
		File tmpFile = new File(WMdataDir+filename);		
		Date filedate = new Date(tmpFile.lastModified());
		//99.11.17 add 查詢年度100年以前.縣市別不同===============================
	    String cd01_table = (Integer.parseInt(m_year) < 100)?"cd01_99":"";
	    String wlx01_m_year = (Integer.parseInt(m_year) < 100)?"99":"100";
	    List updateDBList = new ArrayList();//0:sql 1:data
	    List updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data		
	    //=====================================================================
		if(input_method.equals("F")){
		  System.out.println("tmpFile.lastModified()="+tmpFile.lastModified());
		}
		add_date=bkformat.format(filedate);
		
		try { 
			if(input_method.equals("F")){//若為檔案上傳,先將檔案讀取出來
				M04List = getM04FileData(report_no,filename);
				user_data = Utility.getUser_Data(filename);//取得該檔案的異動者帳號.姓名
				user_id=user_data[0];
				user_name=user_data[1];
				System.out.println("userid="+user_id);
				System.out.println("username="+user_name);
				
				System.out.println("M04List.size="+M04List.size());
			}else{
				user_id = szuser_id;
				user_name = szuser_name;
			}

			errCount = 0;
			System.out.println("layer 1");
			dbData = Utility.getWML03_count(m_year,m_month,bank_code,report_no);//99.11.19 fix
			/*99.11.19移至Utility.getWML03_count讀取WML03.count(*)*/		
		    if(dbData.size() != 0){//WML03有資料時,先清掉
		    	sqlCmd.delete(0,sqlCmd.length());
				sqlCmd.append(" INSERT INTO WML03_LOG "); 
				sqlCmd.append(" select m_year,m_month,bank_code,report_no,serial_no,remark,user_id,user_name,update_date,?,?,sysdate,'D'");
				sqlCmd.append(" from WML03");
				sqlCmd.append(" where m_year=? AND m_month=? AND bank_code=? AND report_no=?");
				updateDBDataList = new ArrayList();
				dataList =  new ArrayList();
				dataList.add(user_id);
				dataList.add(user_name);	
				dataList.add(String.valueOf(Integer.parseInt(m_year)));	
				dataList.add(String.valueOf(Integer.parseInt(m_month)));	
				dataList.add(bank_code);	
				dataList.add(report_no);
				updateDBDataList.add(dataList);
				updateDBSqlList.add(sqlCmd.toString());
				updateDBSqlList.add(updateDBDataList);
				updateDBList.add(updateDBSqlList);
				
				sqlCmd.delete(0,sqlCmd.length());
				updateDBDataList = new ArrayList();
				updateDBSqlList = new ArrayList();
				
				sqlCmd.append("DELETE FROM WML03 WHERE m_year=? AND m_month=? AND bank_code=? AND report_no=?");
				dataList =  new ArrayList();				
				dataList.add(String.valueOf(Integer.parseInt(m_year)));	
				dataList.add(String.valueOf(Integer.parseInt(m_month)));	
				dataList.add(bank_code);	
				dataList.add(report_no);
				updateDBDataList.add(dataList);
				updateDBSqlList.add(sqlCmd.toString());
				updateDBSqlList.add(updateDBDataList);
				updateDBList.add(updateDBSqlList);	
			}
			
			//clear WML03(檢核其他錯誤)
			if(input_method.equals("F")){//檔案上傳

				if(errCount==0){	//資料期間無錯時才insert資料至DB中
					sqlCmd.delete(0,sqlCmd.length());
					paramList = new ArrayList();			
					sqlCmd.append("select * from "+report_no+" where m_year=? AND m_month=?") ;
					paramList.add(String.valueOf(Integer.parseInt(m_year)));
					paramList.add(String.valueOf(Integer.parseInt(m_month)));
					
					dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month");
					
			    	if(dbData.size() == 0){	//M04無資料時,才insert一筆Zero的資料
			    		System.out.println("InsertZeroM04");
			    		InsertZeroM04(m_year,m_month,bank_code);
			    	}else{//end of M04不存在時
			    		updateDBDataList = new ArrayList();	
						updateDBSqlList = new ArrayList();
						sqlCmd.delete(0,sqlCmd.length());
						sqlCmd.append(" INSERT INTO M04_LOG ");
						sqlCmd.append(" select m_year, m_month, loan_use_no, guarantee_no_month, guarantee_no_month_p, loan_amt_month,loan_amt_month_p, guarantee_amt_month, guarantee_amt_month_p, guarantee_no_year, guarantee_no_year_p, loan_amt_year, loan_amt_year_p, guarantee_amt_year, guarantee_amt_year_p, guarantee_no_totacc, guarantee_no_totacc_p, loan_amt_totacc, loan_amt_totacc_p, guarantee_amt_totacc, guarantee_amt_totacc_p,?,?,sysdate,'U'");
						sqlCmd.append(" from M04");
						sqlCmd.append(" WHERE m_year=? AND m_month=?"); 						
		    			dataList =  new ArrayList();
						dataList.add(user_id);
						dataList.add(user_name);
						dataList.add(m_year);
						dataList.add(m_month);
						
						updateDBDataList.add(dataList);
						updateDBSqlList.add(sqlCmd.toString());
						updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
						updateDBList.add(updateDBSqlList);
			    	}
			    	
			    	System.out.println("M04List.size="+M04List.size());
			    	nonZero = false;//94.04.24
			    	updateDBDataList = new ArrayList();	
					updateDBSqlList = new ArrayList();
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("UPDATE M04 SET guarantee_no_month     =?,");
					sqlCmd.append("               guarantee_no_month_p	 =?,");
					sqlCmd.append("               loan_amt_month         =?,");
					sqlCmd.append("               loan_amt_month_p       =?,");
					sqlCmd.append("               guarantee_amt_month    =?,");
					sqlCmd.append("               guarantee_amt_month_p  =?,");
					sqlCmd.append("               guarantee_no_year      =?,");
					sqlCmd.append("               guarantee_no_year_p    =?,");
					sqlCmd.append("               loan_amt_year       	 =?,");
					sqlCmd.append("               loan_amt_year_p        =?,");
					sqlCmd.append("               guarantee_amt_year     =?,");
					sqlCmd.append("               guarantee_amt_year_p	 =?,");
					sqlCmd.append("               guarantee_no_totacc	 =?,");
					sqlCmd.append("               guarantee_no_totacc_p  =?,");
					sqlCmd.append("               loan_amt_totacc        =?,");
					sqlCmd.append("               loan_amt_totacc_p    	 =?,");
					sqlCmd.append("               guarantee_amt_totacc	 =?,");
					sqlCmd.append("               guarantee_amt_totacc_p =? ");
				    sqlCmd.append(" WHERE m_year=? AND m_month=?"); 
				    sqlCmd.append(" AND loan_use_no=?");
			   
					
			    	for(int j=0;j<M04List.size();j++){//把上傳檔案裡的資料更新至M04
			    		//94.04.24 add 檢核上傳檔案內容的金額是否都為"0"
			    		if( Long.parseLong((String)((List)M04List.get(j)).get(1)) > 0 
		    			 ||	Long.parseLong((String)((List)M04List.get(j)).get(2)) > 0
			    		 || Long.parseLong((String)((List)M04List.get(j)).get(3)) > 0
						 || Long.parseLong((String)((List)M04List.get(j)).get(4)) > 0
						 || Long.parseLong((String)((List)M04List.get(j)).get(5)) > 0
						 || Long.parseLong((String)((List)M04List.get(j)).get(6)) > 0
						 || Long.parseLong((String)((List)M04List.get(j)).get(7)) > 0
						 || Long.parseLong((String)((List)M04List.get(j)).get(8)) > 0
						 || Long.parseLong((String)((List)M04List.get(j)).get(9)) > 0
						 || Long.parseLong((String)((List)M04List.get(j)).get(10)) > 0
						 || Long.parseLong((String)((List)M04List.get(j)).get(11)) > 0
						 || Long.parseLong((String)((List)M04List.get(j)).get(12)) > 0
						 || Long.parseLong((String)((List)M04List.get(j)).get(13)) > 0
						 || Long.parseLong((String)((List)M04List.get(j)).get(14)) > 0
						 || Long.parseLong((String)((List)M04List.get(j)).get(15)) > 0
						 || Long.parseLong((String)((List)M04List.get(j)).get(16)) > 0
						 || Long.parseLong((String)((List)M04List.get(j)).get(17)) > 0
						 || Long.parseLong((String)((List)M04List.get(j)).get(18)) > 0)
			    		{
			    		   nonZero = true;
			    		} 
						
			    		if(nonZero){
							dataList =  new ArrayList();   
							dataList.add((String)((List)M04List.get(j)).get( 1));     
							dataList.add((String)((List)M04List.get(j)).get( 2));
							dataList.add((String)((List)M04List.get(j)).get( 3));
							dataList.add((String)((List)M04List.get(j)).get( 4));
							dataList.add((String)((List)M04List.get(j)).get( 5));
							dataList.add((String)((List)M04List.get(j)).get( 6));
							dataList.add((String)((List)M04List.get(j)).get( 7));
							dataList.add((String)((List)M04List.get(j)).get( 8));
							dataList.add((String)((List)M04List.get(j)).get( 9));
							dataList.add((String)((List)M04List.get(j)).get(10));
							dataList.add((String)((List)M04List.get(j)).get(11));
							dataList.add((String)((List)M04List.get(j)).get(12));
							dataList.add((String)((List)M04List.get(j)).get(13));
							dataList.add((String)((List)M04List.get(j)).get(14));
							dataList.add((String)((List)M04List.get(j)).get(15));
							dataList.add((String)((List)M04List.get(j)).get(16));
							dataList.add((String)((List)M04List.get(j)).get(17));
							dataList.add((String)((List)M04List.get(j)).get(18));
							dataList.add(String.valueOf(Integer.parseInt(m_year)));
							dataList.add(String.valueOf(Integer.parseInt(m_month)));
							dataList.add((String)((List)M04List.get(j)).get(0));
							
							updateDBDataList.add(dataList);
			    		}//end of nonZero	
					}//end of M04List				
			    	if(updateDBDataList != null && updateDBDataList.size()!=0){						  
				 	   updateDBSqlList.add(sqlCmd.toString());
					   updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
				 	   updateDBList.add(updateDBSqlList);
					}	
			    	if(updateDBList != null && updateDBList.size()!=0){		
			    		/*
						for(int j=0;j<updateDBList.size();j++){
				    		System.out.println((String)((List)updateDBList.get(j)).get(0));
				    		System.out.println((List)((List)updateDBList.get(j)).get(1));
				    	}
				    	*/				    	
						updateOK=DBManager.updateDB_ps(updateDBList);
						updateDBList = new ArrayList();//0:sql 1:data					
					}
			    }//end of errCount = 0
			    System.out.println("檔案上傳.nonZero="+nonZero);			    
			}//end of if 檔案上傳	
			
			
			if(input_method.equals("W")){//94.04.22 線上編輯check是否值全部為"0"
		    	nonZero = false;
		    	sqlCmd.delete(0,sqlCmd.length());
				paramList = new ArrayList();
				
				sqlCmd.append("select guarantee_no_month,    ");
				sqlCmd.append("       guarantee_no_month_p,  ");
				sqlCmd.append("       loan_amt_month,        ");
				sqlCmd.append("       loan_amt_month_p ,     ");
				sqlCmd.append("       guarantee_amt_month,   ");
				sqlCmd.append("       guarantee_amt_month_p, ");
				sqlCmd.append("       guarantee_no_year    , ");
				sqlCmd.append("       guarantee_no_year_p  , ");
				sqlCmd.append("       loan_amt_year        , ");
				sqlCmd.append("       loan_amt_year_p      , ");
				sqlCmd.append("       guarantee_amt_year   , ");
				sqlCmd.append("       guarantee_amt_year_p , ");
				sqlCmd.append("       guarantee_no_totacc  , ");
				sqlCmd.append("       guarantee_no_totacc_p, ");
				sqlCmd.append("       loan_amt_totacc      , ");
				sqlCmd.append("       loan_amt_totacc_p    , ");
				sqlCmd.append("       guarantee_amt_totacc , ");
				sqlCmd.append("       guarantee_amt_totacc_p ");
				sqlCmd.append(" from	M04 ");
			    sqlCmd.append(" WHERE m_year=? AND m_month=?");
			    paramList.add(String.valueOf(Integer.parseInt(m_year)));
				paramList.add(String.valueOf(Integer.parseInt(m_month)));
				dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,guarantee_no_month, guarantee_no_month_p, loan_amt_month, loan_amt_month_p, guarantee_amt_month, guarantee_amt_month_p, guarantee_no_year, guarantee_no_year_p, loan_amt_year, loan_amt_year_p, guarantee_amt_year, guarantee_amt_year_p, guarantee_no_totacc, guarantee_no_totacc_p, loan_amt_totacc, loan_amt_totacc_p, guarantee_amt_totacc, guarantee_amt_totacc_p");     
		    	checkZeroLoop:
			    if(dbData != null && dbData.size() != 0){
			    	for(int zeroIdx=0;zeroIdx < dbData.size();zeroIdx++){			    		  
			    		if(Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("guarantee_no_month")).toString()) > 0
			    		||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("guarantee_no_month_p")).toString()) > 0
			    		||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("loan_amt_month")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("loan_amt_month_p")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("guarantee_amt_month")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("guarantee_amt_month_p")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("guarantee_no_year")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("guarantee_no_year_p")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("loan_amt_year")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("loan_amt_year_p")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("guarantee_amt_year")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("guarantee_amt_year_p")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("guarantee_no_totacc")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("guarantee_no_totacc_p")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("loan_amt_totacc")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("loan_amt_totacc_p")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("guarantee_amt_totacc")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("guarantee_amt_totacc_p")).toString()) > 0
			    		){
				    	   nonZero = true;
				    	   break checkZeroLoop;
				    	}
			    	}
			    }
		    	System.out.println("線上編輯.nonZero="+nonZero);
		    }//end of 線上編輯check non zero
		    
		    //94.04.22 fix 只要有非"0"值時,才執行檢核
			//WML03 (檢查合計和小計)
			if(nonZero && errCount == 0){//無其他錯誤時,才檢核合計金額
				sqlCmd.delete(0,sqlCmd.length());
				paramList = new ArrayList();
				sqlCmd.append("select decode(loan_use_no,'0','0','1') as \"type\", ");
                sqlCmd.append("      sum(guarantee_no_month) as \"guarantee_no_month\", ");
                sqlCmd.append("      sum(guarantee_no_month_p) as \"guarantee_no_month_p\", ");
                sqlCmd.append("      sum(loan_amt_month) as \"loan_amt_month\", ");
                sqlCmd.append("      sum(loan_amt_month_p) as \"loan_amt_month_p\", ");
                sqlCmd.append("      sum(guarantee_amt_month_p) as \"guarantee_amt_month\", ");
                sqlCmd.append("      sum(guarantee_amt_month_p) as \"guarantee_amt_month_p\", ");
                sqlCmd.append("      sum(guarantee_no_year) as \"guarantee_no_year\", ");
                sqlCmd.append("      sum(guarantee_no_year_p) as \"guarantee_no_year_p\", ");
                sqlCmd.append("      sum(loan_amt_year) as \"loan_amt_year\", ");
                sqlCmd.append("      sum(loan_amt_year_p) as \"loan_amt_year_p\", ");
                sqlCmd.append("      sum(guarantee_amt_year) as \"guarantee_amt_year\", ");
                sqlCmd.append("      sum(guarantee_amt_year_p) as \"guarantee_amt_year_p\", ");
                sqlCmd.append("      sum(guarantee_no_totacc) as \"guarantee_no_totacc\", ");
                sqlCmd.append("      sum(guarantee_no_totacc_p) as \"guarantee_no_totacc_p\", ");
                sqlCmd.append("      sum(loan_amt_totacc) as \"loan_amt_totacc\", ");
                sqlCmd.append("      sum(loan_amt_totacc_p) as \"loan_amt_totacc_p\", ");
                sqlCmd.append("      sum(guarantee_amt_totacc) as \"guarantee_amt_totacc\", ");
                sqlCmd.append("      sum(guarantee_amt_totacc_p) as \"guarantee_amt_totacc_p\" ");
                sqlCmd.append("from	M04 ");
                sqlCmd.append("where m_year=? and m_month=?");
                sqlCmd.append("group by decode(loan_use_no,'0','0','1')");
                sqlCmd.append("order by decode(loan_use_no,'0','0','1')");
                paramList.add(String.valueOf(Integer.parseInt(m_year)));
                paramList.add(String.valueOf(Integer.parseInt(m_month)));
                dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,guarantee_no_month, guarantee_no_month_p, loan_amt_month, loan_amt_month_p, guarantee_amt_month, guarantee_amt_month_p, guarantee_no_year, guarantee_no_year_p, loan_amt_year, loan_amt_year_p, guarantee_amt_year, guarantee_amt_year_p, guarantee_no_totacc, guarantee_no_totacc_p, loan_amt_totacc, loan_amt_totacc_p, guarantee_amt_totacc, guarantee_amt_totacc_p");		
					
				long guarantee_no_month		= Long.parseLong((((DataObject)dbData.get(0)).getValue("guarantee_no_month")		==null?"0":((DataObject)dbData.get(0)).getValue("guarantee_no_month")).toString());
				long guarantee_no_month_p 	= Long.parseLong((((DataObject)dbData.get(0)).getValue("guarantee_no_month_p")		==null?"0":((DataObject)dbData.get(0)).getValue("guarantee_no_month_p")).toString());
				long loan_amt_month			= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_amt_month")			==null?"0":((DataObject)dbData.get(0)).getValue("loan_amt_month")).toString());
				long loan_amt_month_p  		= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_amt_month_p")			==null?"0":((DataObject)dbData.get(0)).getValue("loan_amt_month_p")).toString());
				long guarantee_amt_month	= Long.parseLong((((DataObject)dbData.get(0)).getValue("guarantee_amt_month")		==null?"0":((DataObject)dbData.get(0)).getValue("guarantee_amt_month")).toString());
				long guarantee_amt_month_p	= Long.parseLong((((DataObject)dbData.get(0)).getValue("guarantee_amt_month_p")		==null?"0":((DataObject)dbData.get(0)).getValue("guarantee_amt_month_p")).toString());
				long guarantee_no_year		= Long.parseLong((((DataObject)dbData.get(0)).getValue("guarantee_no_year")			==null?"0":((DataObject)dbData.get(0)).getValue("guarantee_no_year")).toString());
				long guarantee_no_year_p	= Long.parseLong((((DataObject)dbData.get(0)).getValue("guarantee_no_year_p")		==null?"0":((DataObject)dbData.get(0)).getValue("guarantee_no_year_p")).toString());
				long loan_amt_year			= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_amt_year")				==null?"0":((DataObject)dbData.get(0)).getValue("loan_amt_year")).toString());
				long loan_amt_year_p		= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_amt_year_p")			==null?"0":((DataObject)dbData.get(0)).getValue("loan_amt_year_p")).toString());
				long guarantee_amt_year  	= Long.parseLong((((DataObject)dbData.get(0)).getValue("guarantee_amt_year")		==null?"0":((DataObject)dbData.get(0)).getValue("guarantee_amt_year")).toString());
				long guarantee_amt_year_p  	= Long.parseLong((((DataObject)dbData.get(0)).getValue("guarantee_amt_year_p")		==null?"0":((DataObject)dbData.get(0)).getValue("guarantee_amt_year_p")).toString());
				long guarantee_no_totacc  	= Long.parseLong((((DataObject)dbData.get(0)).getValue("guarantee_no_totacc")		==null?"0":((DataObject)dbData.get(0)).getValue("guarantee_no_totacc")).toString());
				long guarantee_no_totacc_p  = Long.parseLong((((DataObject)dbData.get(0)).getValue("guarantee_no_totacc_p")		==null?"0":((DataObject)dbData.get(0)).getValue("guarantee_no_totacc_p")).toString());
				long loan_amt_totacc  		= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_amt_totacc")			==null?"0":((DataObject)dbData.get(0)).getValue("loan_amt_totacc")).toString());
				long loan_amt_totacc_p  	= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_amt_totacc_p")			==null?"0":((DataObject)dbData.get(0)).getValue("loan_amt_totacc_p")).toString());
				long guarantee_amt_totacc  	= Long.parseLong((((DataObject)dbData.get(0)).getValue("guarantee_amt_totacc")		==null?"0":((DataObject)dbData.get(0)).getValue("guarantee_amt_totacc")).toString());
				long guarantee_amt_totacc_p = Long.parseLong((((DataObject)dbData.get(0)).getValue("guarantee_amt_totacc_p")	==null?"0":((DataObject)dbData.get(0)).getValue("guarantee_amt_totacc_p")).toString());
				
				long t_guarantee_no_month		= Long.parseLong((((DataObject)dbData.get(1)).getValue("guarantee_no_month")		==null?"0":((DataObject)dbData.get(1)).getValue("guarantee_no_month")).toString());
				long t_guarantee_no_month_p 	= Long.parseLong((((DataObject)dbData.get(1)).getValue("guarantee_no_month_p")	==null?"0":((DataObject)dbData.get(1)).getValue("guarantee_no_month_p")).toString());
				long t_loan_amt_month			= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_amt_month")			==null?"0":((DataObject)dbData.get(1)).getValue("loan_amt_month")).toString());
				long t_loan_amt_month_p  		= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_amt_month_p")		==null?"0":((DataObject)dbData.get(1)).getValue("loan_amt_month_p")).toString());
				long t_guarantee_amt_month		= Long.parseLong((((DataObject)dbData.get(1)).getValue("guarantee_amt_month")		==null?"0":((DataObject)dbData.get(1)).getValue("guarantee_amt_month")).toString());
				long t_guarantee_amt_month_p	= Long.parseLong((((DataObject)dbData.get(1)).getValue("guarantee_amt_month_p")	==null?"0":((DataObject)dbData.get(1)).getValue("guarantee_amt_month_p")).toString());
				long t_guarantee_no_year		= Long.parseLong((((DataObject)dbData.get(1)).getValue("guarantee_no_year")		==null?"0":((DataObject)dbData.get(1)).getValue("guarantee_no_year")).toString());
				long t_guarantee_no_year_p		= Long.parseLong((((DataObject)dbData.get(1)).getValue("guarantee_no_year_p")		==null?"0":((DataObject)dbData.get(1)).getValue("guarantee_no_year_p")).toString());
				long t_loan_amt_year			= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_amt_year")			==null?"0":((DataObject)dbData.get(1)).getValue("loan_amt_year")).toString());
				long t_loan_amt_year_p			= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_amt_year_p")			==null?"0":((DataObject)dbData.get(1)).getValue("loan_amt_year_p")).toString());
				long t_guarantee_amt_year  		= Long.parseLong((((DataObject)dbData.get(1)).getValue("guarantee_amt_year")		==null?"0":((DataObject)dbData.get(1)).getValue("guarantee_amt_year")).toString());
				long t_guarantee_amt_year_p  	= Long.parseLong((((DataObject)dbData.get(1)).getValue("guarantee_amt_year_p")	==null?"0":((DataObject)dbData.get(1)).getValue("guarantee_amt_year_p")).toString());
				long t_guarantee_no_totacc  	= Long.parseLong((((DataObject)dbData.get(1)).getValue("guarantee_no_totacc")		==null?"0":((DataObject)dbData.get(1)).getValue("guarantee_no_totacc")).toString());
				long t_guarantee_no_totacc_p  	= Long.parseLong((((DataObject)dbData.get(1)).getValue("guarantee_no_totacc_p")	==null?"0":((DataObject)dbData.get(1)).getValue("guarantee_no_totacc_p")).toString());
				long t_loan_amt_totacc  		= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_amt_totacc")			==null?"0":((DataObject)dbData.get(1)).getValue("loan_amt_totacc")).toString());
				long t_loan_amt_totacc_p  		= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_amt_totacc_p")		==null?"0":((DataObject)dbData.get(1)).getValue("loan_amt_totacc_p")).toString());
				long t_guarantee_amt_totacc  	= Long.parseLong((((DataObject)dbData.get(1)).getValue("guarantee_amt_totacc")	==null?"0":((DataObject)dbData.get(1)).getValue("guarantee_amt_totacc")).toString());
				long t_guarantee_amt_totacc_p 	= Long.parseLong((((DataObject)dbData.get(1)).getValue("guarantee_amt_totacc_p")	==null?"0":((DataObject)dbData.get(1)).getValue("guarantee_amt_totacc_p")).toString());
				
				updateDBDataList = new ArrayList();	//99.11.18
				String subErrMsg="";
				if(guarantee_no_month != t_guarantee_no_month ) { //比對筆數金額
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"當月份保證件數加總與合計列不合",user_id,user_name));								
					errCount++ ;						
				}
				//94.04.22 add 百分比比對的區間在-1 ~ 1之間都OK by 2295
				if(guarantee_no_month_p != t_guarantee_no_month_p && 
					(Math.abs(guarantee_no_month_p - t_guarantee_no_month_p)/1000) > 1) {
					System.out.println("guarantee_no_month_p="+guarantee_no_month_p);
					System.out.println("t_guarantee_no_month_p="+t_guarantee_no_month_p);
					System.out.println("當月份保證件數結構比加總與合計列不合="+Math.abs(guarantee_no_month_p - t_guarantee_no_month_p));
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"當月份保證件數結構比加總與合計列不合",user_id,user_name));
					errCount++ ;
				}
				if(loan_amt_month	!= t_loan_amt_month	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"當月份貸款金額加總與合計列不合",user_id,user_name));						
					errCount++ ;
				}
				if(loan_amt_month_p  	!= t_loan_amt_month_p  	&& 
					(Math.abs(loan_amt_month_p - t_loan_amt_month_p)/1000) > 1) {
					System.out.println("loan_amt_month_p="+loan_amt_month_p);
					System.out.println("t_loan_amt_month_p="+t_loan_amt_month_p);
					System.out.println("當月份貸款金額結構比加總與合計列不合="+Math.abs(loan_amt_month_p - t_loan_amt_month_p));
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"當月份貸款金額結構比加總與合計列不合",user_id,user_name));					
					errCount++ ;
				}
				if(guarantee_amt_month != t_guarantee_amt_month && 
				  (Math.abs(guarantee_amt_month - t_guarantee_amt_month)/1000) > 1) {
					System.out.println("guarantee_amt_month="+guarantee_amt_month);
					System.out.println("t_guarantee_amt_month="+t_guarantee_amt_month);
					System.out.println("當月份保證金額比加總與合計列不合="+Math.abs(guarantee_amt_month - t_guarantee_amt_month));
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"當月份保證金額比加總與合計列不合",user_id,user_name));					
					errCount++ ;						
				}
				if(guarantee_amt_month_p != t_guarantee_amt_month_p  && 
					(Math.abs(guarantee_amt_month_p - t_guarantee_amt_month_p)/1000) > 1 ) {
					System.out.println("guarantee_amt_month_p="+guarantee_amt_month_p);
					System.out.println("t_guarantee_amt_month_p="+t_guarantee_amt_month_p);
					System.out.println("當月份保證金額結構比加總與合計列不合="+Math.abs(guarantee_amt_month_p - t_guarantee_amt_month_p));
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"當月份保證金額結構比加總與合計列不合",user_id,user_name));					
					errCount++ ;
				}
				if(guarantee_no_year	!= t_guarantee_no_year) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"當年度保證件數加總與合計列不合",user_id,user_name));					
					errCount++ ;
				}
				if(guarantee_no_year_p	!= t_guarantee_no_year_p && 
					(Math.abs(guarantee_no_year_p - t_guarantee_no_year_p)/1000) > 1) {
					System.out.println("guarantee_no_year_p="+guarantee_no_year_p);
					System.out.println("t_guarantee_no_year_p="+t_guarantee_no_year_p);
					System.out.println("當年度保證件數結構比加總與合計列不合="+Math.abs(guarantee_no_year_p - t_guarantee_no_year_p));
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"當年度保證件數結構比加總與合計列不合",user_id,user_name));					
					errCount++ ;						
				}
				if(loan_amt_year	!= t_loan_amt_year) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"當年度貸款金額加總與合計列不合",user_id,user_name));					
					errCount++ ;						
				}
				if(loan_amt_year_p	!= t_loan_amt_year_p && 
					(Math.abs(loan_amt_year_p - t_loan_amt_year_p)/1000) > 1) {
					System.out.println("loan_amt_year_p="+loan_amt_year_p);
					System.out.println("t_loan_amt_year_p="+t_loan_amt_year_p);
					System.out.println("當年度貸款金額結構比加總與合計列不合="+Math.abs(loan_amt_year_p - t_loan_amt_year_p));
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"當年度貸款金額結構比加總與合計列不合",user_id,user_name));					
					errCount++ ;
				}
				if(guarantee_amt_year  	!= t_guarantee_amt_year  ) {
				    updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"當年度保證金額加總與合計列不合",user_id,user_name));					
					errCount++ ;
				}
				if(guarantee_amt_year_p  != t_guarantee_amt_year_p  && 
					(Math.abs(guarantee_amt_year_p - t_guarantee_amt_year_p)/1000) > 1	) {
					System.out.println("guarantee_amt_year_p="+guarantee_amt_year_p);
					System.out.println("t_guarantee_amt_year_p="+t_guarantee_amt_year_p);
					System.out.println("當年度保證金額結構比加總與合計列不合="+Math.abs(guarantee_amt_year_p - t_guarantee_amt_year_p));
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"當年度保證金額結構比加總與合計列不合",user_id,user_name));					
					errCount++ ;						
				}
				if(guarantee_no_totacc  != t_guarantee_no_totacc  	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"總累計保證件數加總與合計列不合",user_id,user_name));					
					errCount++ ;						
				}
				if(guarantee_no_totacc_p  != t_guarantee_no_totacc_p  && 
					(Math.abs(guarantee_no_totacc_p - t_guarantee_no_totacc_p)/1000) > 1 	) {
					System.out.println("guarantee_no_totacc_p="+guarantee_no_totacc_p);
					System.out.println("t_guarantee_no_totacc_p="+t_guarantee_no_totacc_p);
					System.out.println("總累計保證件數結構比加總與合計列不合="+Math.abs(guarantee_no_totacc_p - t_guarantee_no_totacc_p));
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"總累計保證件數結構比加總與合計列不合",user_id,user_name));					
					errCount++ ;
				}
				if(loan_amt_totacc  != t_loan_amt_totacc  	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"總累計貸款金額加總與合計列不合",user_id,user_name));					
					errCount++ ;						
				}
				if(loan_amt_totacc_p  != t_loan_amt_totacc_p  && 
				   (Math.abs(loan_amt_totacc_p - t_loan_amt_totacc_p)/1000) > 1	) {
					System.out.println("loan_amt_totacc_p="+loan_amt_totacc_p);
					System.out.println("t_loan_amt_totacc_p="+t_loan_amt_totacc_p);
					System.out.println("總累計貸款金額結構比加總與合計列不合="+Math.abs(loan_amt_totacc_p - t_loan_amt_totacc_p));
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"總累計貸款金額結構比加總與合計列不合",user_id,user_name));					
					errCount++ ;
				}
				if(guarantee_amt_totacc  != t_guarantee_amt_totacc  ) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"總累計保證金額加總與合計列不合",user_id,user_name));					
					errCount++ ;						
				}
				if(guarantee_amt_totacc_p  != t_guarantee_amt_totacc_p && 
				   (Math.abs(guarantee_amt_totacc_p - t_guarantee_amt_totacc_p)/1000) > 1 ) {
					System.out.println("guarantee_amt_totacc_p="+guarantee_amt_totacc_p);
					System.out.println("t_guarantee_amt_totacc_p="+t_guarantee_amt_totacc_p);
					System.out.println("總累計保證金額結構比加總與合計列不合="+Math.abs(guarantee_amt_totacc_p - t_guarantee_amt_totacc_p));
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"總累計保證金額結構比加總與合計列不合",user_id,user_name));					
					errCount++ ;						
				}
				if(updateDBDataList != null && updateDBDataList.size()!=0){
					updateDBSqlList = new ArrayList();	
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("INSERT INTO WML03 VALUES (?,?,?,?,?,?,?,?,sysdate)");
					updateDBSqlList.add(sqlCmd.toString());
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
				}
			}//enf of errCount==0
			dbData =  Utility.getWML01(m_year,m_month,bank_code,report_no);
			/*99.11.19移至Utility.getWML01讀取WML01 all data*/		
			upd_code = (errCount == 0)?"U":"E";//U檢核成功:E檢核失敗
			checkResult = (errCount == 0)?"true":"false";//true檢核成功:false檢核失敗
			//94.04.22 fix Z檢核為0====================================
			if(!nonZero && errCount == 0){
				upd_code = "Z";//Z檢核為0
				checkResult="Z";//Z檢核為0
			}
			//========================================================	
			List getUpdateDBList = Utility.Insert_UpdateWML01(dbData,m_year, m_month,bank_code,report_no,input_method,add_date,user_id,user_name,common_center,upd_method,upd_code,batch_no);//99.11.19 fix
			for(int j=0;j<getUpdateDBList.size();j++){
			    updateDBList.add(getUpdateDBList.get(j));
			}
			/*99.11.19移至Utility.Insert_UpdateWML01;wml01不存在Insert,存在時Update*/
			
			//傳送email至bank_cmml的e-mail信箱
			System.out.println("send email begin");
			if(input_method.equals("F")){//檔案上傳					
			    Utility.sendMailNotification(bank_code,report_no,
				m_year,m_month, checkResult,input_method,emailformat.format(filedate),filename,user_id,"");
			}else{//線上編輯  		
				Utility.sendMailNotification(bank_code,report_no,
				m_year,m_month, checkResult,input_method,Utility.getDateFormat("yyyy/MM/dd HH:mm:ss"),"",user_id,"");							
			}		
			System.out.println("send email end");
			
				
			if(input_method.equals("F")){//檔案上傳
			   //將WML01_UPLOAD..filename對應的使用者帳號.姓名刪除			
			   //99.11.17updateDBSqlList.add("DELETE FROM WML01_UPLOAD where filename='"+filename+"'");
				updateDBList.add(Utility.deleteWML01_UPLOAD(filename));//99.11.19
				/*99.11.19移至Utility.deleteWML01_UPLOAD*/
			}
			if(updateDBList != null && updateDBList.size()!=0){
				/*
				for(int j=0;j<updateDBList.size();j++){
		    		System.out.println((String)((List)updateDBList.get(j)).get(0));
		    		System.out.println((List)((List)updateDBList.get(j)).get(1));
		    	}
		    	*/
				updateOK=DBManager.updateDB_ps(updateDBList);
			}
			System.out.println("update OK??"+updateOK);
			if(!updateOK){
				//parserResult=false;
				errMsg = errMsg + "UpdateM04.doParserReport_M04 UpdateDB Error:"+DBManager.getErrMsg()+"<br>";
				System.out.println(DBManager.getErrMsg());
			}
			
			if(input_method.equals("F")){//檔案上傳
			   String CopyResult = Utility.CopyFile(WMdataDir+System.getProperty("file.separator")+filename,WMdataBKDir+System.getProperty("file.separator")+ filename);
       		   if(CopyResult.equals("0")){//copy成功時,才將檔案刪除,避免使用rename造成的錯誤
          	   		tmpFile = new File(WMdataDir+System.getProperty("file.separator")+filename);
          	   		if(tmpFile.exists()) tmpFile.delete();              		
       		   }	   				
			}	
		}catch (Exception e) {			
			errMsg = errMsg + "UpdateM04.doParserReport_M04 Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateM04.doParserReport_M04="+e.getMessage());			
		}
		
		return errMsg;
		
	}

	//讀取上傳檔案的資料
	private static List getM04FileData(String report_no,String filename){
			String	txtline	 = null;			
			List M04List = new LinkedList();
			List detail = null;
			List dbData = null;	
			
			try {
				String WMdataDir = Utility.getProperties("WMdataDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");
				String WMdataBKDir = Utility.getProperties("WMdataBKDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");		
				File tmpFile = new File(WMdataDir+filename);		
				Date filedate = new Date(tmpFile.lastModified());
					
				FileReader	f		= new FileReader(tmpFile);
				LineNumberReader in	= new LineNumberReader(f);
				doLoop://將txt檔案儲存至LinkedList - M04List
				while ((txtline = in.readLine()) != null) {
						if (txtline.trim().equals("")){
							continue doLoop;
						}else{
							detail=new LinkedList();
							
							detail.add(txtline.substring(0, 1));//貸款用途
			                detail.add(txtline.substring(1, 8));//當月份保證件數
			                detail.add(txtline.substring(8, 15));//當月份保證件數結構比%
			                detail.add(txtline.substring(15, 29));//當月份貸款金額  
			                detail.add(txtline.substring(29, 36));//當月份貸款金額結構比% 
			                detail.add(txtline.substring(36, 50));//當月份保證金額
			                detail.add(txtline.substring(50, 57));//當月份保證金額結構比%
			                detail.add(txtline.substring(57, 64));//當年度保證件數
			                detail.add(txtline.substring(64, 71));//當年度保證件數結構比%
			                detail.add(txtline.substring(71, 85));//當年度貸款金額     
			                detail.add(txtline.substring(85, 92));//當年度貸款金額結構比%   
			                detail.add(txtline.substring(92, 106));//當年度保證金額
			                detail.add(txtline.substring(106, 113));//當年度保證金額結構比%
			                detail.add(txtline.substring(113, 120));//總累計保證件數
			                detail.add(txtline.substring(120, 127));//當年度保證件數結構比%
			                detail.add(txtline.substring(127, 141));//當年度貸款金額 
			                detail.add(txtline.substring(141, 148));//當年度貸款金額結構比%     
			                detail.add(txtline.substring(148, 162));//當年度保證金額
			                detail.add(txtline.substring(162, 169));//當年度保證金額結構比%				
							M04List.add(detail);
						}
				}	
				in.close();
				f.close();
			}catch(Exception e){
				errMsg = errMsg + "UpdateM04.getM04FileData Error:"+e.getMessage()+"<br>";
			}
			
			return M04List;
	}
	
	//Insert "0" 至M04
	//99.11.17 add 套用DAO.preparestatment,並列印轉換後的SQL by 2295
	private static boolean InsertZeroM04(String m_year,String m_month,String bank_code){
		List paramList = new ArrayList();
		StringBuffer sqlCmd = new StringBuffer();	
		List updateDBList = new LinkedList();//0:sql 1:data		
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data
		boolean updateOK=false;		
		try{
			sqlCmd.append(" select loan_use_no as \"loan_use_no\" ,loan_use_name ");
			sqlCmd.append(" from m00_loan_use "); 
			sqlCmd.append(" order by input_order ");
			List data_div01 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),null,"update_date");	    
    		System.out.println("data_div01.size="+data_div01.size());
    		
    		sqlCmd.delete(0, sqlCmd.length());
			sqlCmd.append("INSERT INTO M04 ( M_YEAR,M_MONTH,LOAN_USE_NO ) VALUES (?,?,?)");
		
			for(int d1=0;d1<data_div01.size();d1++){
				String sm_year="";
				sm_year="0000"+m_year; sm_year=sm_year.substring(sm_year.length()-3);
				dataList = new LinkedList();//儲存參數的data
				dataList.add(String.valueOf(Integer.parseInt(m_year)));
				dataList.add(String.valueOf(Integer.parseInt(m_month)));
				dataList.add((String)((DataObject)data_div01.get(d1)).getValue("loan_use_no"));				
				updateDBDataList.add(dataList);//1:傳內的參數List
			}
			if(updateDBDataList != null && updateDBDataList.size()!=0){
			   updateDBSqlList.add(sqlCmd.toString());
			   updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
			   updateDBList.add(updateDBSqlList);
			}
				
			updateOK=DBManager.updateDB_ps(updateDBList);
			System.out.println("M04 Insert Zero OK??"+updateOK);
			if(!updateOK){
				errMsg = errMsg + "M04 Insert Zero Fail:"+DBManager.getErrMsg()+"<br>";
				System.out.println(DBManager.getErrMsg());
			}
		}catch(Exception e){
			errMsg = errMsg + "UpdateM04.InsertZeroM04 Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateM04.InsertZeroM04 Error:"+e.getMessage());
		}
    	
    	return updateOK;
	}	
}


