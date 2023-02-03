//94.04.22 fix 都為"0"值時,顯示檢核為"0" by 2295
//94.06.22 add 百分比比對的區間在-1 ~ 1之間都OK by 2295
//99.11.18 寫入WML03的參數移至共用create_dataList;add ALL SQL 套用DAO.preparestatment,並列印轉換後的SQL by 2295
//99.11.19 add 移至共用Utility.getWML03_count讀取WML03 count(*)
//					  Utility.getWML01讀取WML01 all data
//					  Utility.Insert_UpdateWML01當WML01不存在Insert,存在時Update
//					  Utility.deleteWML01_UPLOAD刪除上傳檔案紀錄 by 2295

package com.tradevan.util;

import java.io.*;
import java.util.*;
import java.text.SimpleDateFormat;
import com.tradevan.util.dao.DataObject;


public class UpdateM06{
	private static String errMsg = "";
	
	public String getErrMsg(){
		return errMsg;  
	}
	public static synchronized String doParserReport_M06(String report_no, String m_year,
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
		
		List M06List = new LinkedList();//M06細部資料
		List dbData = null;//其他querydb後,資料暫存的list
		String checkResult="true";
		boolean nonZero=false;//94.04.22 檢查內容是否都為"0"
		File tmpFile = new File(WMdataDir+filename);		
		Date filedate = new Date(tmpFile.lastModified());
		//99.11.18 add 查詢年度100年以前.縣市別不同===============================
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
				M06List = getM06FileData(report_no,filename);
				user_data = Utility.getUser_Data(filename);//取得該檔案的異動者帳號.姓名
				user_id=user_data[0];
				user_name=user_data[1];
				System.out.println("userid="+user_id);
				System.out.println("username="+user_name);
				
				System.out.println("M06List.size="+M06List.size());
			}else{
				user_id = szuser_id;
				user_name = szuser_name;
			}
			
			errCount = 0;
			//判斷WML03資料是否存在，存在則刪除舊資料，並insert至WML03_LOG中
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
				//資料期間無錯時才insert資料至DB中
				if(errCount==0){	
					sqlCmd.delete(0,sqlCmd.length());
					paramList = new ArrayList();			
					sqlCmd.append("select * from "+report_no+" where m_year=? AND m_month=?") ;
					paramList.add(String.valueOf(Integer.parseInt(m_year)));
					paramList.add(String.valueOf(Integer.parseInt(m_month)));
					
					dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month");
					if(dbData.size() == 0){	//M06無資料時,才insert一筆Zero的資料
			    		System.out.println("InsertZeroM06");
			    		InsertZeroM06(m_year,m_month,bank_code);
			    	}else{//end of M06不存在時
			    		updateDBDataList = new ArrayList();	
						updateDBSqlList = new ArrayList();
						sqlCmd.delete(0,sqlCmd.length());
						sqlCmd.append(" INSERT INTO M06_LOG ");
						sqlCmd.append(" select m_year, m_month, area_no, guarantee_no_month, guarantee_amt_month, loan_amt_month, guarantee_no_year, guarantee_amt_year, loan_amt_year, guarantee_no_totacc, guarantee_amt_totacc, loan_amt_totacc, guarantee_bal_no, guarantee_bal_amt, guarantee_bal_p, loan_bal");
						sqlCmd.append(" ,?,?,sysdate,'U'");
						sqlCmd.append(" from M06");
						sqlCmd.append(" WHERE m_year= ? AND m_month=?"); 
						sqlCmd.append(" AND area_no=?");					
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
			    	
				    nonZero = false;//94.04.24
				    updateDBDataList = new ArrayList();	
					updateDBSqlList = new ArrayList();
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("UPDATE M06 SET guarantee_no_month    =?,");
					sqlCmd.append("               guarantee_amt_month	=?,");
					sqlCmd.append("               loan_amt_month        =?,");
					sqlCmd.append("               guarantee_no_year     =?,");
					sqlCmd.append("               guarantee_amt_year    =?,");
					sqlCmd.append("               loan_amt_year 		=?,");
					sqlCmd.append("               guarantee_no_totacc   =?,");
					sqlCmd.append("               guarantee_amt_totacc  =?,");
					sqlCmd.append("               loan_amt_totacc      	=?,");
					sqlCmd.append("               guarantee_bal_no      =?,");
					sqlCmd.append("               guarantee_bal_amt    	=?,");
					sqlCmd.append("               guarantee_bal_p		=?,");
					sqlCmd.append("               loan_bal				=? ");
				    sqlCmd.append(" WHERE m_year=? AND m_month=?"); 
				    sqlCmd.append(" AND area_no=?");
		
			    	System.out.println("M06List.size="+M06List.size());
			    	
			    	for(int j=0;j<M06List.size();j++){//把上傳檔案裡的資料更新至M06
			    		//94.04.24 add 檢核上傳檔案內容的金額是否都為"0"
			    		if( Long.parseLong((String)((List)M06List.get(j)).get(1)) > 0
					     || Long.parseLong((String)((List)M06List.get(j)).get(2)) > 0
					     || Long.parseLong((String)((List)M06List.get(j)).get(3)) > 0					      
		    			 ||	Long.parseLong((String)((List)M06List.get(j)).get(4)) > 0
						 || Long.parseLong((String)((List)M06List.get(j)).get(5)) > 0
						 || Long.parseLong((String)((List)M06List.get(j)).get(6)) > 0
						 || Long.parseLong((String)((List)M06List.get(j)).get(7)) > 0
						 || Long.parseLong((String)((List)M06List.get(j)).get(8)) > 0
						 || Long.parseLong((String)((List)M06List.get(j)).get(9)) > 0
						 || Long.parseLong((String)((List)M06List.get(j)).get(10)) > 0
			    		 || Long.parseLong((String)((List)M06List.get(j)).get(11)) > 0
			    		 || Long.parseLong((String)((List)M06List.get(j)).get(12)) > 0
			    		 || Long.parseLong((String)((List)M06List.get(j)).get(13)) > 0)
			    		{
			    		   nonZero = true;
			    		}
			    		if(nonZero){
							dataList =  new ArrayList();
							dataList.add((String)((List)M06List.get(j)).get( 1));     
							dataList.add((String)((List)M06List.get(j)).get( 2));
							dataList.add((String)((List)M06List.get(j)).get( 3));
							dataList.add((String)((List)M06List.get(j)).get( 4));
							dataList.add((String)((List)M06List.get(j)).get( 5));
							dataList.add((String)((List)M06List.get(j)).get( 6));
							dataList.add((String)((List)M06List.get(j)).get( 7));
							dataList.add((String)((List)M06List.get(j)).get( 8));
							dataList.add((String)((List)M06List.get(j)).get( 9));
							dataList.add((String)((List)M06List.get(j)).get(10));
							dataList.add((String)((List)M06List.get(j)).get(11));
							dataList.add((String)((List)M06List.get(j)).get(12));
							dataList.add((String)((List)M06List.get(j)).get(13));
							dataList.add(String.valueOf(Integer.parseInt(m_year)));
							dataList.add(String.valueOf(Integer.parseInt(m_month)));
							dataList.add((String)((List)M06List.get(j)).get(0));
							
							updateDBDataList.add(dataList);
			    		}//end of nonZero
			    	}//end of M06List
			    	if(updateDBDataList != null && updateDBDataList.size()!=0){						  
					   updateDBSqlList.add(sqlCmd.toString());
					   updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					   updateDBList.add(updateDBSqlList);
					}	
				    			    	
			    	//寫入小計資料處理
			    	System.out.println("寫入小計資料處理...");
			    	updateDBDataList = new ArrayList();	
					updateDBSqlList = new ArrayList();
					sqlCmd.delete(0,sqlCmd.length());
			    	sqlCmd.append("UPDATE M06 ");
					sqlCmd.append("SET guarantee_no_month   =(select sum(guarantee_no_month)   from m06 where m_year=? and m_month=? and area_no not in ('0','1','2','A','E')), ");
					sqlCmd.append("    guarantee_amt_month  =(select sum(guarantee_amt_month)  from m06 where m_year=? and m_month=? and area_no not in ('0','1','2','A','E')), ");
					sqlCmd.append("    loan_amt_month       =(select sum(loan_amt_month)       from m06 where m_year=? and m_month=? and area_no not in ('0','1','2','A','E')), ");
					sqlCmd.append("    guarantee_no_year    =(select sum(guarantee_no_year)    from m06 where m_year=? and m_month=? and area_no not in ('0','1','2','A','E')), ");
					sqlCmd.append("    guarantee_amt_year   =(select sum(guarantee_amt_year)   from m06 where m_year=? and m_month=? and area_no not in ('0','1','2','A','E')), ");
					sqlCmd.append("    loan_amt_year        =(select sum(loan_amt_year)        from m06 where m_year=? and m_month=? and area_no not in ('0','1','2','A','E')), ");
					sqlCmd.append("    guarantee_no_totacc  =(select sum(guarantee_no_totacc)  from m06 where m_year=? and m_month=? and area_no not in ('0','1','2','A','E')), ");
					sqlCmd.append("    guarantee_amt_totacc =(select sum(guarantee_amt_totacc) from m06 where m_year=? and m_month=? and area_no not in ('0','1','2','A','E')), ");
					sqlCmd.append("    loan_amt_totacc      =(select sum(loan_amt_totacc)      from m06 where m_year=? and m_month=? and area_no not in ('0','1','2','A','E')), ");
					sqlCmd.append("    guarantee_bal_no     =(select sum(guarantee_bal_no)     from m06 where m_year=? and m_month=? and area_no not in ('0','1','2','A','E')), ");
					sqlCmd.append("    guarantee_bal_amt    =(select sum(guarantee_bal_amt)    from m06 where m_year=? and m_month=? and area_no not in ('0','1','2','A','E')), ");
					sqlCmd.append("    guarantee_bal_p      =(select sum(guarantee_bal_p)      from m06 where m_year=? and m_month=? and area_no not in ('0','1','2','A','E')), ");
					sqlCmd.append("    loan_bal             =(select sum(loan_bal)             from m06 where m_year=? and m_month=? and area_no not in ('0','1','2','A','E'))  ");
					sqlCmd.append("where m_year=? AND m_month=?");
					sqlCmd.append("and area_no='2'");
					dataList =  new ArrayList();
					for(int i=0;i<=13;i++){
					    dataList.add(String.valueOf(Integer.parseInt(m_year)));
					    dataList.add(String.valueOf(Integer.parseInt(m_month)));
					}
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
					
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
					System.out.println("檔案上傳.nonZero="+nonZero);
			    	System.out.println("M06 update data OK??"+updateOK);			    	
			    }
			}//end of if 檔案上傳	
			if(input_method.equals("W")){//94.04.22 線上編輯check是否值全部為"0"
		    	nonZero = false;
		    	sqlCmd.delete(0,sqlCmd.length());
				paramList = new ArrayList();
		    	sqlCmd.append(" select guarantee_no_month,guarantee_amt_month,loan_amt_month,");
		    	sqlCmd.append("        guarantee_no_year,guarantee_amt_year,loan_amt_year,guarantee_no_totacc,guarantee_amt_totacc,");
		    	sqlCmd.append("        loan_amt_totacc,guarantee_bal_no,guarantee_bal_amt,guarantee_bal_p,loan_bal");
		    	sqlCmd.append("   from M06");
		    	sqlCmd.append(" WHERE m_year=? AND m_month=?");
		    	paramList.add(String.valueOf(Integer.parseInt(m_year)));
				paramList.add(String.valueOf(Integer.parseInt(m_month)));      
				dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"guarantee_no_month,guarantee_amt_month,loan_amt_month,guarantee_no_year,guarantee_amt_year,loan_amt_year,guarantee_no_totacc,guarantee_amt_totacc,loan_amt_totacc,guarantee_bal_no,guarantee_bal_amt,guarantee_bal_p,loan_bal");     
		    	checkZeroLoop:
			    if(dbData != null && dbData.size() != 0){
			    	for(int zeroIdx=0;zeroIdx < dbData.size();zeroIdx++){			    		  
			    		if(	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("guarantee_no_month")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("guarantee_amt_month")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("loan_amt_month")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("guarantee_no_year")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("guarantee_amt_year")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("loan_amt_year")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("guarantee_no_totacc")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("guarantee_amt_totacc")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("loan_amt_totacc")).toString()) > 0												
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("guarantee_bal_no")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("guarantee_bal_amt")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("guarantee_bal_p")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("loan_bal")).toString()) > 0						
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
				//檢核台灣省小計
				sqlCmd.append("select decode(area_no,'2','0','1') as \"type\", ");
				sqlCmd.append("      sum(guarantee_no_month) as \"guarantee_no_month\", ");
                sqlCmd.append("      sum(guarantee_amt_month) as \"guarantee_amt_month\", ");
                sqlCmd.append("      sum(loan_amt_month) as \"loan_amt_month\", ");
                sqlCmd.append("      sum(guarantee_no_year) as \"guarantee_no_year\", ");
                sqlCmd.append("      sum(guarantee_amt_year) as \"guarantee_amt_year\", ");
                sqlCmd.append("      sum(loan_amt_year) as \"loan_amt_year\", ");
                sqlCmd.append("      sum(guarantee_no_totacc) as \"guarantee_no_totacc\", ");
                sqlCmd.append("      sum(guarantee_amt_totacc) as \"guarantee_amt_totacc\", ");
                sqlCmd.append("      sum(loan_amt_totacc) as \"loan_amt_totacc\", ");
                sqlCmd.append("      sum(guarantee_bal_no) as \"guarantee_bal_no\", ");
                sqlCmd.append("      sum(guarantee_bal_amt) as \"guarantee_bal_amt\", ");
                sqlCmd.append("      sum(guarantee_bal_p) as \"guarantee_bal_p\", ");
                sqlCmd.append("      sum(loan_bal) as \"loan_bal\" ");
                sqlCmd.append(" from	M06 ");
                sqlCmd.append("where m_year=? and m_month=?");
                sqlCmd.append("and area_no not in ('A','E','0') ");
                sqlCmd.append("group by decode(area_no,'2','0','1') ");
                sqlCmd.append("order by decode(area_no,'2','0','1') desc");
                paramList.add(String.valueOf(Integer.parseInt(m_year)));
                paramList.add(String.valueOf(Integer.parseInt(m_month)));
                dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,guarantee_no_month,guarantee_amt_month,loan_amt_month,guarantee_no_year,guarantee_amt_year,loan_amt_year,guarantee_no_totacc,guarantee_amt_totacc,loan_amt_totacc,guarantee_bal_no,guarantee_bal_amt,guarantee_bal_p,loan_bal");

				long ts_guarantee_no_month		= Long.parseLong((((DataObject)dbData.get(0)).getValue("guarantee_no_month")	==null?"0":((DataObject)dbData.get(0)).getValue("guarantee_no_month")).toString());
				long ts_guarantee_amt_month 	= Long.parseLong((((DataObject)dbData.get(0)).getValue("guarantee_amt_month")	==null?"0":((DataObject)dbData.get(0)).getValue("guarantee_amt_month")).toString());
				long ts_loan_amt_month			= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_amt_month")		==null?"0":((DataObject)dbData.get(0)).getValue("loan_amt_month")).toString());
				long ts_guarantee_no_year  	= Long.parseLong((((DataObject)dbData.get(0)).getValue("guarantee_no_year")		==null?"0":((DataObject)dbData.get(0)).getValue("guarantee_no_year")).toString());
				long ts_guarantee_amt_year		= Long.parseLong((((DataObject)dbData.get(0)).getValue("guarantee_amt_year")	==null?"0":((DataObject)dbData.get(0)).getValue("guarantee_amt_year")).toString());
				long ts_loan_amt_year			= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_amt_year")			==null?"0":((DataObject)dbData.get(0)).getValue("loan_amt_year")).toString());
				long ts_guarantee_no_totacc	= Long.parseLong((((DataObject)dbData.get(0)).getValue("guarantee_no_totacc")	==null?"0":((DataObject)dbData.get(0)).getValue("guarantee_no_totacc")).toString());
				long ts_guarantee_amt_totacc	= Long.parseLong((((DataObject)dbData.get(0)).getValue("guarantee_amt_totacc")	==null?"0":((DataObject)dbData.get(0)).getValue("guarantee_amt_totacc")).toString());
				long ts_loan_amt_totacc		= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_amt_totacc")		==null?"0":((DataObject)dbData.get(0)).getValue("loan_amt_totacc")).toString());
				long ts_guarantee_bal_no		= Long.parseLong((((DataObject)dbData.get(0)).getValue("guarantee_bal_no")		==null?"0":((DataObject)dbData.get(0)).getValue("guarantee_bal_no")).toString());
				long ts_guarantee_bal_amt  	= Long.parseLong((((DataObject)dbData.get(0)).getValue("guarantee_bal_amt")		==null?"0":((DataObject)dbData.get(0)).getValue("guarantee_bal_amt")).toString());
				long ts_guarantee_bal_p  		= Long.parseLong((((DataObject)dbData.get(0)).getValue("guarantee_bal_p")		==null?"0":((DataObject)dbData.get(0)).getValue("guarantee_bal_p")).toString());
				long ts_loan_bal  				= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_bal")				==null?"0":((DataObject)dbData.get(0)).getValue("loan_bal")).toString());
				
				long s_guarantee_no_month	= Long.parseLong((((DataObject)dbData.get(1)).getValue("guarantee_no_month")	==null?"0":((DataObject)dbData.get(1)).getValue("guarantee_no_month")).toString());
				long s_guarantee_amt_month 	= Long.parseLong((((DataObject)dbData.get(1)).getValue("guarantee_amt_month")	==null?"0":((DataObject)dbData.get(1)).getValue("guarantee_amt_month")).toString());
				long s_loan_amt_month		= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_amt_month")		==null?"0":((DataObject)dbData.get(1)).getValue("loan_amt_month")).toString());
				long s_guarantee_no_year  	= Long.parseLong((((DataObject)dbData.get(1)).getValue("guarantee_no_year")		==null?"0":((DataObject)dbData.get(1)).getValue("guarantee_no_year")).toString());
				long s_guarantee_amt_year	= Long.parseLong((((DataObject)dbData.get(1)).getValue("guarantee_amt_year")	==null?"0":((DataObject)dbData.get(1)).getValue("guarantee_amt_year")).toString());
				long s_loan_amt_year		= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_amt_year")			==null?"0":((DataObject)dbData.get(1)).getValue("loan_amt_year")).toString());
				long s_guarantee_no_totacc	= Long.parseLong((((DataObject)dbData.get(1)).getValue("guarantee_no_totacc")	==null?"0":((DataObject)dbData.get(1)).getValue("guarantee_no_totacc")).toString());
				long s_guarantee_amt_totacc	= Long.parseLong((((DataObject)dbData.get(1)).getValue("guarantee_amt_totacc")	==null?"0":((DataObject)dbData.get(1)).getValue("guarantee_amt_totacc")).toString());
				long s_loan_amt_totacc		= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_amt_totacc")		==null?"0":((DataObject)dbData.get(1)).getValue("loan_amt_totacc")).toString());
				long s_guarantee_bal_no		= Long.parseLong((((DataObject)dbData.get(1)).getValue("guarantee_bal_no")		==null?"0":((DataObject)dbData.get(1)).getValue("guarantee_bal_no")).toString());
				long s_guarantee_bal_amt  	= Long.parseLong((((DataObject)dbData.get(1)).getValue("guarantee_bal_amt")		==null?"0":((DataObject)dbData.get(1)).getValue("guarantee_bal_amt")).toString());
				long s_guarantee_bal_p  	= Long.parseLong((((DataObject)dbData.get(1)).getValue("guarantee_bal_p")		==null?"0":((DataObject)dbData.get(1)).getValue("guarantee_bal_p")).toString());
				long s_loan_bal  			= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_bal")				==null?"0":((DataObject)dbData.get(1)).getValue("loan_bal")).toString());
				
				updateDBDataList = new ArrayList();//99.11.18	
				
				if(ts_guarantee_no_month != s_guarantee_no_month	) { //比對筆數金額
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"台灣省本月份保證件數加總與小計列不合",user_id,user_name));								
					errCount++ ;						
				}
				if(ts_guarantee_amt_month != s_guarantee_amt_month ) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"台灣省本月份保證金額加總與小計列不合",user_id,user_name));								
					errCount++ ;						
				}
				if(ts_loan_amt_month	!= s_loan_amt_month	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"台灣省本月份融資金額加總與小計列不合",user_id,user_name));								
					errCount++ ;
				}
				if(ts_guarantee_no_year  	!= s_guarantee_no_year  ) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"台灣省本年度保證件數加總與小計列不合",user_id,user_name));								
					errCount++ ;						
				}
				if(ts_guarantee_amt_year != s_guarantee_amt_year   ) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"台灣省本年度保證金額加總與小計列不合",user_id,user_name));								
					errCount++ ;						
				}
				if(ts_loan_amt_year != s_loan_amt_year   ) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"台灣省本年度融資金額加總與小計列不合",user_id,user_name));								
					errCount++ ;						
				}
				if(ts_guarantee_no_totacc	!= s_guarantee_no_totacc	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"台灣省累計保證件數加總與小計列不合",user_id,user_name));								
					errCount++ ;						
				}
				if(ts_guarantee_amt_totacc	!= s_guarantee_amt_totacc	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"台灣省累計保證金額加總與小計列不合",user_id,user_name));								
					errCount++ ;						
				}
				if(ts_loan_amt_totacc	!= s_loan_amt_totacc	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"台灣省累計融資金額加總與小計列不合",user_id,user_name));								
					errCount++ ;
				}
				if(ts_guarantee_bal_no	!= s_guarantee_bal_no    ) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"台灣省保證餘額件數加總與小計列不合",user_id,user_name));								
					errCount++ ;						
				}
				if(ts_guarantee_bal_amt  != s_guarantee_bal_amt  ) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"台灣省保證餘額金額加總與小計列不合",user_id,user_name));								
					errCount++ ;						
				}
				//94.06.22 add 百分比比對的區間在-1 ~ 1之間都OK by 2295
				if(ts_guarantee_bal_p  != s_guarantee_bal_p && 
					(Math.abs(ts_guarantee_bal_p - s_guarantee_bal_p)/1000) > 1) {
					System.out.println("ts_guarantee_bal_p="+ts_guarantee_bal_p);
					System.out.println("s_guarantee_bal_p="+s_guarantee_bal_p);
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"台灣省保證餘額比率加總與小計列不合",user_id,user_name));								
					errCount++ ;
				}
				if(ts_loan_bal  	!= s_loan_bal  	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"台灣省融資餘額加總與小計列不合",user_id,user_name));								
					errCount++ ;						
				}
					                       
				//檢核全部合計
				sqlCmd.delete(0,sqlCmd.length());
				sqlCmd.append("select decode(area_no,'0','0','1') as \"type\", ");
                sqlCmd.append("      sum(guarantee_no_month) as \"guarantee_no_month\", ");
                sqlCmd.append("      sum(guarantee_amt_month) as \"guarantee_amt_month\", ");
                sqlCmd.append("      sum(loan_amt_month) as \"loan_amt_month\", ");
                sqlCmd.append("      sum(guarantee_no_year) as \"guarantee_no_year\", ");
                sqlCmd.append("      sum(guarantee_amt_year) as \"guarantee_amt_year\", ");
                sqlCmd.append("      sum(loan_amt_year) as \"loan_amt_year\", ");
                sqlCmd.append("      sum(guarantee_no_totacc) as \"guarantee_no_totacc\", ");
                sqlCmd.append("      sum(guarantee_amt_totacc) as \"guarantee_amt_totacc\", ");
                sqlCmd.append("      sum(loan_amt_totacc) as \"loan_amt_totacc\", ");
                sqlCmd.append("      sum(guarantee_bal_no) as \"guarantee_bal_no\", ");
                sqlCmd.append("      sum(guarantee_bal_amt) as \"guarantee_bal_amt\", ");
                sqlCmd.append("      sum(guarantee_bal_p) as \"guarantee_bal_p\", ");
                sqlCmd.append("      sum(loan_bal) as \"loan_bal\" ");
                sqlCmd.append(" from	M06 ");
                sqlCmd.append("where m_year=? and m_month=?");
                sqlCmd.append("and area_no in ('A','E','2','0') ");
                sqlCmd.append("group by decode(area_no,'0','0','1') ");
                sqlCmd.append("order by decode(area_no,'0','0','1') desc");
                dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,guarantee_no_month,guarantee_amt_month,loan_amt_month,guarantee_no_year,guarantee_amt_year,loan_amt_year,guarantee_no_totacc,guarantee_amt_totacc,loan_amt_totacc,guarantee_bal_no,guarantee_bal_amt,guarantee_bal_p,loan_bal");		
					
				long guarantee_no_month		= Long.parseLong((((DataObject)dbData.get(0)).getValue("guarantee_no_month")	==null?"0":((DataObject)dbData.get(0)).getValue("guarantee_no_month")).toString());
				long guarantee_amt_month 	= Long.parseLong((((DataObject)dbData.get(0)).getValue("guarantee_amt_month")	==null?"0":((DataObject)dbData.get(0)).getValue("guarantee_amt_month")).toString());
				long loan_amt_month			= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_amt_month")		==null?"0":((DataObject)dbData.get(0)).getValue("loan_amt_month")).toString());
				long guarantee_no_year  	= Long.parseLong((((DataObject)dbData.get(0)).getValue("guarantee_no_year")		==null?"0":((DataObject)dbData.get(0)).getValue("guarantee_no_year")).toString());
				long guarantee_amt_year		= Long.parseLong((((DataObject)dbData.get(0)).getValue("guarantee_amt_year")	==null?"0":((DataObject)dbData.get(0)).getValue("guarantee_amt_year")).toString());
				long loan_amt_year			= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_amt_year")			==null?"0":((DataObject)dbData.get(0)).getValue("loan_amt_year")).toString());
				long guarantee_no_totacc	= Long.parseLong((((DataObject)dbData.get(0)).getValue("guarantee_no_totacc")	==null?"0":((DataObject)dbData.get(0)).getValue("guarantee_no_totacc")).toString());
				long guarantee_amt_totacc	= Long.parseLong((((DataObject)dbData.get(0)).getValue("guarantee_amt_totacc")	==null?"0":((DataObject)dbData.get(0)).getValue("guarantee_amt_totacc")).toString());
				long loan_amt_totacc		= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_amt_totacc")		==null?"0":((DataObject)dbData.get(0)).getValue("loan_amt_totacc")).toString());
				long guarantee_bal_no		= Long.parseLong((((DataObject)dbData.get(0)).getValue("guarantee_bal_no")		==null?"0":((DataObject)dbData.get(0)).getValue("guarantee_bal_no")).toString());
				long guarantee_bal_amt  	= Long.parseLong((((DataObject)dbData.get(0)).getValue("guarantee_bal_amt")		==null?"0":((DataObject)dbData.get(0)).getValue("guarantee_bal_amt")).toString());
				long guarantee_bal_p  		= Long.parseLong((((DataObject)dbData.get(0)).getValue("guarantee_bal_p")		==null?"0":((DataObject)dbData.get(0)).getValue("guarantee_bal_p")).toString());
				long loan_bal  				= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_bal")				==null?"0":((DataObject)dbData.get(0)).getValue("loan_bal")).toString());
				
				long t_guarantee_no_month	= Long.parseLong((((DataObject)dbData.get(1)).getValue("guarantee_no_month")	==null?"0":((DataObject)dbData.get(1)).getValue("guarantee_no_month")).toString());
				long t_guarantee_amt_month 	= Long.parseLong((((DataObject)dbData.get(1)).getValue("guarantee_amt_month")	==null?"0":((DataObject)dbData.get(1)).getValue("guarantee_amt_month")).toString());
				long t_loan_amt_month		= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_amt_month")		==null?"0":((DataObject)dbData.get(1)).getValue("loan_amt_month")).toString());
				long t_guarantee_no_year  	= Long.parseLong((((DataObject)dbData.get(1)).getValue("guarantee_no_year")		==null?"0":((DataObject)dbData.get(1)).getValue("guarantee_no_year")).toString());
				long t_guarantee_amt_year	= Long.parseLong((((DataObject)dbData.get(1)).getValue("guarantee_amt_year")	==null?"0":((DataObject)dbData.get(1)).getValue("guarantee_amt_year")).toString());
				long t_loan_amt_year		= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_amt_year")			==null?"0":((DataObject)dbData.get(1)).getValue("loan_amt_year")).toString());
				long t_guarantee_no_totacc	= Long.parseLong((((DataObject)dbData.get(1)).getValue("guarantee_no_totacc")	==null?"0":((DataObject)dbData.get(1)).getValue("guarantee_no_totacc")).toString());
				long t_guarantee_amt_totacc	= Long.parseLong((((DataObject)dbData.get(1)).getValue("guarantee_amt_totacc")	==null?"0":((DataObject)dbData.get(1)).getValue("guarantee_amt_totacc")).toString());
				long t_loan_amt_totacc		= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_amt_totacc")		==null?"0":((DataObject)dbData.get(1)).getValue("loan_amt_totacc")).toString());
				long t_guarantee_bal_no		= Long.parseLong((((DataObject)dbData.get(1)).getValue("guarantee_bal_no")		==null?"0":((DataObject)dbData.get(1)).getValue("guarantee_bal_no")).toString());
				long t_guarantee_bal_amt  	= Long.parseLong((((DataObject)dbData.get(1)).getValue("guarantee_bal_amt")		==null?"0":((DataObject)dbData.get(1)).getValue("guarantee_bal_amt")).toString());
				long t_guarantee_bal_p  	= Long.parseLong((((DataObject)dbData.get(1)).getValue("guarantee_bal_p")		==null?"0":((DataObject)dbData.get(1)).getValue("guarantee_bal_p")).toString());
				long t_loan_bal  			= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_bal")				==null?"0":((DataObject)dbData.get(1)).getValue("loan_bal")).toString());
				
				if(guarantee_no_month != t_guarantee_no_month	) { //比對筆數金額
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"本月份保證件數加總與合計列不合",user_id,user_name));								
					errCount++ ;					
				}
				if(guarantee_amt_month != t_guarantee_amt_month ) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"本月份保證金額加總與合計列不合",user_id,user_name));								
					errCount++ ;
				}
				if(loan_amt_month	!= t_loan_amt_month	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"本月份融資金額加總與合計列不合",user_id,user_name));								
					errCount++ ;					
				}
				if(guarantee_no_year  	!= t_guarantee_no_year  ) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"本年度保證件數加總與合計列不合",user_id,user_name));								
					errCount++ ;					
				}
				if(guarantee_amt_year != t_guarantee_amt_year   ) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"本年度保證金額加總與合計列不合",user_id,user_name));								
					errCount++ ;					
				}
				if(loan_amt_year != t_loan_amt_year   ) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"本年度融資金額加總與合計列不合",user_id,user_name));								
					errCount++ ;					
				}
				if(guarantee_no_totacc	!= t_guarantee_no_totacc	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"累計保證件數加總與合計列不合",user_id,user_name));								
					errCount++ ;					
				}
				if(guarantee_amt_totacc	!= t_guarantee_amt_totacc	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"累計保證金額加總與合計列不合",user_id,user_name));								
					errCount++ ;
				}
				if(loan_amt_totacc	!= t_loan_amt_totacc	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"累計融資金額加總與合計列不合",user_id,user_name));								
					errCount++ ;
				}
				if(guarantee_bal_no	!= t_guarantee_bal_no    ) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"保證餘額件數加總與合計列不合",user_id,user_name));								
					errCount++ ;
				}
				if(guarantee_bal_amt  != t_guarantee_bal_amt  ) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"保證餘額金額加總與合計列不合",user_id,user_name));								
					errCount++ ;					
				}
				//94.06.22 add 百分比比對的區間在-1 ~ 1之間都OK by 2295
				if(guarantee_bal_p  != t_guarantee_bal_p  && 
				   (Math.abs(guarantee_bal_p - t_guarantee_bal_p)/1000) > 1) {
					System.out.println("guarantee_bal_p="+guarantee_bal_p);
					System.out.println("t_guarantee_bal_p="+t_guarantee_bal_p);
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"保證餘額比率加總與合計列不合",user_id,user_name));								
					errCount++ ;
				}
				if(loan_bal  	!= t_loan_bal  	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"融資餘額加總與合計列不合",user_id,user_name));								
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
			    //99.11.18updateDBSqlList.add("DELETE FROM WML01_UPLOAD where filename='"+filename+"'");
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
				errMsg = errMsg + "UpdateM06.doParserReport_M06 UpdateDB Error:"+DBManager.getErrMsg()+"<br>";
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
			errMsg = errMsg + "UpdateM06.doParserReport_M06 Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateM06.doParserReport_M06="+e.getMessage());			
		}
		
		return errMsg;
		
	}

	//讀取上傳檔案的資料
	private static List getM06FileData(String report_no,String filename){
			String	txtline	 = null;			
			List M06List = new LinkedList();
			List detail = null;
			List dbData = null;	
			
			try {
				String WMdataDir = Utility.getProperties("WMdataDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");
				String WMdataBKDir = Utility.getProperties("WMdataBKDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");		
				File tmpFile = new File(WMdataDir+filename);		
				Date filedate = new Date(tmpFile.lastModified());
					
				FileReader	f		= new FileReader(tmpFile);
				LineNumberReader in	= new LineNumberReader(f);
				doLoop://將txt檔案儲存至LinkedList - M06List
				while ((txtline = in.readLine()) != null) {
						if (txtline.trim().equals("")){
							continue doLoop;
						}else{
							detail=new LinkedList();
							detail.add(txtline.substring(  0,  1));	//地區
							detail.add(txtline.substring(  1,  8));	//本月份保證件數
							detail.add(txtline.substring(  8, 22));	//本月份保證金額
							detail.add(txtline.substring( 22, 36));	//本月份融資金額         
							detail.add(txtline.substring( 36, 43));	//本年度保證件數       
							detail.add(txtline.substring( 43, 57));	//本年度保證金額
							detail.add(txtline.substring( 57, 71));	//本年度融資金額
							detail.add(txtline.substring( 71, 78));	//累計保證件數
							detail.add(txtline.substring( 78, 92));	//累計保證金額
							detail.add(txtline.substring( 92,106));	//累計融資金額        
							detail.add(txtline.substring(106,113));	//保證餘額件數
							detail.add(txtline.substring(113,127));	//保證餘額金額
							detail.add(txtline.substring(127,134));	//保證餘額結構比
							detail.add(txtline.substring(134,148));	//融資餘額
							
							M06List.add(detail);
						}
				}	
				in.close();
				f.close();
			}catch(Exception e){
				errMsg = errMsg + "UpdateM06.getM06FileData Error:"+e.getMessage()+"<br>";
			}
			return M06List;
	}
	
	//Insert "0" 至M06
	//99.11.18 add 套用DAO.preparestatment,並列印轉換後的SQL by 2295
	private static boolean InsertZeroM06(String m_year,String m_month,String bank_code){
		List paramList = new ArrayList();
		StringBuffer sqlCmd = new StringBuffer();	
		List updateDBList = new LinkedList();//0:sql 1:data		
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data
		boolean updateOK=false;
		
		try{
			sqlCmd.append(" select area_no as \"area_no\" ,area_name ");
			sqlCmd.append(" from m00_area "); 
			sqlCmd.append(" order by input_order ");
			List data_div01 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),null,"update_date");//	    
    		System.out.println("data_div01.size="+data_div01.size());
    		sqlCmd.delete(0, sqlCmd.length());
			sqlCmd.append("INSERT INTO M06 ( M_YEAR,M_MONTH,AREA_NO ) VALUES (?,?,?)");
		
			for(int d1=0;d1<data_div01.size();d1++){
				//Insert M06
				dataList = new LinkedList();//儲存參數的data
				dataList.add(String.valueOf(Integer.parseInt(m_year)));
				dataList.add(String.valueOf(Integer.parseInt(m_month)));
				dataList.add((String)((DataObject)data_div01.get(d1)).getValue("area_no"));				
				updateDBDataList.add(dataList);//1:傳內的參數List				
			}
			if(updateDBDataList != null && updateDBDataList.size()!=0){
			   updateDBSqlList.add(sqlCmd.toString());
			   updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
			   updateDBList.add(updateDBSqlList);
			}
					
			updateOK=DBManager.updateDB_ps(updateDBList);
			System.out.println("M06 Insert Zero OK??"+updateOK);
			if(!updateOK){
				errMsg = errMsg + "M06 Insert Zero Fail:"+DBManager.getErrMsg()+"<br>";
				System.out.println(DBManager.getErrMsg());
			}
		}catch(Exception e){
			errMsg = errMsg + "UpdateM06.InsertZeroM06 Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateM06.InsertZeroM06 Error:"+e.getMessage());
		}
    	
    	return updateOK;
	}	
}


