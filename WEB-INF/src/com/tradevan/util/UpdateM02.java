//94.04.22 fix 都為"0"值時,顯示檢核為"0" by 2295
//99.11.17 add 寫入WML03的參數移至共用create_dataList;ALL SQL 套用DAO.preparestatment,並列印轉換後的SQL by 2295
//99.11.19 add 移至共用Utility.getWML03_count讀取WML03 count(*)
//					  Utility.getWML01讀取WML01 all data
//					  Utility.Insert_UpdateWML01當WML01不存在Insert,存在時Update
//					  Utility.deleteWML01_UPLOAD刪除上傳檔案紀錄 by 2295

package com.tradevan.util;

import java.io.*;
import java.util.*;
import java.text.SimpleDateFormat;
import com.tradevan.util.dao.DataObject;


public class UpdateM02 {
	private static String errMsg = "";
	public String getErrMsg(){
		return errMsg;  
	}
	public static synchronized String doParserReport_M02(String report_no, String m_year,
			String m_month,String filename, String srcbank_code,String upd_method,String input_method,String bank_type,String szuser_id,String szuser_name,int batch_no)
			throws Exception {
		//boolean parserResult=true;
		errMsg = "";
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();		
		int		errCount = 0;		
		String  user_id="";//本次異動者帳號
		String  user_name="";//本次異動者姓名
		String[] user_data = {"",""};
		String  bank_code=srcbank_code;//農業信用保證基金 	
		String  add_date="";//申報日期
		String  upd_code="";//檢核結果
		String  common_center="";//由共用中心傳入
		boolean updateOK=false;
		SimpleDateFormat bkformat = new SimpleDateFormat("yyyyMMddHHmmss");
		SimpleDateFormat emailformat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");	
		String WMdataDir = Utility.getProperties("WMdataDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");
		String WMdataBKDir = Utility.getProperties("WMdataBKDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");		
				
		List M02List = new LinkedList();//M02細部資料
		List dbData = null;//其他querydb後,資料暫存的list
		String checkResult="true";
		boolean nonZero=false;//94.04.22 檢查內容是否都為"0"
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
				M02List = getM02FileData(report_no,filename);
				user_data = Utility.getUser_Data(filename);//取得該檔案的異動者帳號.姓名
				user_id=user_data[0];
				user_name=user_data[1];
				System.out.println("userid="+user_id);
				System.out.println("username="+user_name);
			}else{
				user_id = szuser_id;
				user_name = szuser_name;
			}
			
			errCount = 0;
			//clear WML03(檢核其他錯誤)
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
				/*99.11.17
				sqlCmd = " INSERT INTO WML03_LOG " 
					   + " select m_year,m_month,bank_code,report_no,serial_no,remark,user_id,user_name,update_date"
					   + ",'"+user_id+"','"+user_name+"',sysdate,'D'"
					   + " from WML03"
					   + " where m_year=" + String.valueOf(Integer.parseInt(m_year)) + " AND m_month=" + String.valueOf(Integer.parseInt(m_month)) + " AND " 
					   + " bank_code='" + bank_code + "' AND report_no='" + report_no + "'";
				updateDBSqlList.add(sqlCmd);			
				sqlCmd = "DELETE FROM WML03 WHERE m_year=" + String.valueOf(Integer.parseInt(m_year)) + " AND m_month=" + String.valueOf(Integer.parseInt(m_month)) + " AND " +
					   "bank_code='" + bank_code + "' AND report_no='" + report_no + "'";
				updateDBSqlList.add(sqlCmd);
				*/
			}
			
			if(input_method.equals("F")){//檔案上傳				
				/* 取得所有保證項目代碼, 已於上傳時檢核*/
				//dbData = DBManager.QueryDB("select distinct loan_unit_no from m00_guarantee_item","");
	 			//for(int i=0;i<dbData.size();i++){
				//    values.put((String)((DataObject)dbData.get(i)).getValue("bank_no"),"true");
				//}
				String sm_year="",sm_month="";
				sm_year="0000"+m_year; sm_year=sm_year.substring(sm_year.length()-3);
				sm_month="0000"+m_month; sm_month=sm_month.substring(sm_month.length()-2);
				
				updateDBDataList = new ArrayList();	
				updateDBSqlList = new ArrayList();	
				sqlCmd.delete(0,sqlCmd.length());
				sqlCmd.append("INSERT INTO WML03 VALUES (?,?,?,?,?,?,?,?,sysdate)");
				
				for(int j=0;j<M02List.size();j++){					
					String dataRange=(String)((List)M02List.get(j)).get(1);
					if(!dataRange.equals(sm_year+"MM") && 
					   !dataRange.equals(sm_year+"MT") && 
					   !dataRange.equals(sm_year+"YY") &&
					   !dataRange.equals(sm_year+"YT") && 
					   !dataRange.equals(sm_year+"TT") && 
					   !dataRange.equals(sm_year+"ET")
					  ){
						dataList =  new ArrayList();	
						dataList.add(String.valueOf(Integer.parseInt(m_year)));
						dataList.add(String.valueOf(Integer.parseInt(m_month)));
						dataList.add(bank_code);
						dataList.add(report_no);
						dataList.add(String.valueOf(errCount));
						dataList.add("資料期間[" + (String)((List)M02List.get(j)).get(1) +"]不存在");
						dataList.add(user_id);
						dataList.add(user_name);
						updateDBDataList.add(dataList);//1:傳內的參數List		
						errCount++ ;
					}
				}//end of M02List
				
				if(updateDBDataList != null && updateDBDataList.size()!=0){
					updateDBSqlList.add(sqlCmd.toString());
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
				}
				
				if(errCount==0){	//資料期間無錯時才insert資料至DB中     
					sqlCmd.delete(0,sqlCmd.length());
					paramList = new ArrayList();
					sqlCmd.append("select * from "+report_no+" where m_year=? AND m_month=?" );
					paramList.add(String.valueOf(Integer.parseInt(m_year)));
					paramList.add(String.valueOf(Integer.parseInt(m_month)));
					dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month");
					
					if(dbData.size() == 0){//M02無資料時,才insert一筆Zero的資料
			    		System.out.println("InsertZeroM02");
			    		InsertZeroM02(m_year,m_month,bank_code);
			    	}else{//end of M02不存在時
			    		updateDBDataList = new ArrayList();	
						updateDBSqlList = new ArrayList();
						sqlCmd.delete(0,sqlCmd.length());
						
			    		sqlCmd.append(" INSERT INTO M02_LOG ");
			    		sqlCmd.append(" select m_year,m_month,loan_unit_no,data_range,guarantee_cnt,loan_amt,");
			    		sqlCmd.append("        guarantee_amt,loan_bal,guarantee_bal,over_notpush_cnt,over_notpush_bal,");
			    		sqlCmd.append("        over_okpush_cnt,over_okpush_bal,repay_tot_cnt,repay_tot_amt,");
			    		sqlCmd.append("        repay_bal_cnt,repay_bal_amt,?,?,sysdate,'U'");
			    		sqlCmd.append("   from M02");
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
			    	
			    	nonZero = false;//94.04.24
			    	updateDBDataList = new ArrayList();	
					updateDBSqlList = new ArrayList();
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("UPDATE M02 SET guarantee_cnt =?,");
					sqlCmd.append("               loan_amt            =?,");
					sqlCmd.append("               guarantee_amt       =?,");
					sqlCmd.append("               loan_bal            =?,");
					sqlCmd.append("               guarantee_bal       =?,");
				    sqlCmd.append("               over_notpush_cnt    =?,");
					sqlCmd.append("               over_notpush_bal    =?,");
					sqlCmd.append("               over_okpush_cnt     =?,");
					sqlCmd.append("               over_okpush_bal     =?,");
					sqlCmd.append("               repay_tot_cnt       =?,");
					sqlCmd.append("               repay_tot_amt       =?,");
					sqlCmd.append("               repay_bal_cnt       =?,");
					sqlCmd.append("               repay_bal_amt		 =?");
					sqlCmd.append(" WHERE m_year=? AND m_month=?"); 
				    sqlCmd.append("   AND loan_unit_no=? ");
					sqlCmd.append("   AND data_range=?");
					
			    	for(int j=0;j<M02List.size();j++){//把上傳檔案裡的資料更新至M02				
			    		String dataRange=(String)((List)M02List.get(j)).get(1);
						if(dataRange.equals(sm_year+sm_month)) 	dataRange = sm_year+"MM";
						if(dataRange.equals(sm_year+"00"))		dataRange = sm_year+"YY";
						//94.04.24 add 檢核上傳檔案內容的金額是否都為"0"
			    		if( Long.parseLong((String)((List)M02List.get(j)).get(2)) > 0
			    		 || Long.parseLong((String)((List)M02List.get(j)).get(3)) > 0
						 || Long.parseLong((String)((List)M02List.get(j)).get(4)) > 0
						 || Long.parseLong((String)((List)M02List.get(j)).get(5)) > 0
						 || Long.parseLong((String)((List)M02List.get(j)).get(6)) > 0
						 || Long.parseLong((String)((List)M02List.get(j)).get(7)) > 0
						 || Long.parseLong((String)((List)M02List.get(j)).get(8)) > 0
						 || Long.parseLong((String)((List)M02List.get(j)).get(9)) > 0
						 || Long.parseLong((String)((List)M02List.get(j)).get(10)) > 0
						 || Long.parseLong((String)((List)M02List.get(j)).get(11)) > 0
						 || Long.parseLong((String)((List)M02List.get(j)).get(12)) > 0
						 || Long.parseLong((String)((List)M02List.get(j)).get(13)) > 0
						 || Long.parseLong((String)((List)M02List.get(j)).get(14)) > 0)
			    		{
			    		   nonZero = true;
			    		}
			    		
			    		if(nonZero){
							dataList =  new ArrayList();
							dataList.add((String)((List)M02List.get(j)).get( 2));
							dataList.add((String)((List)M02List.get(j)).get( 3));
							dataList.add((String)((List)M02List.get(j)).get( 4));
							dataList.add((String)((List)M02List.get(j)).get( 5));
							dataList.add((String)((List)M02List.get(j)).get( 6));
							dataList.add((String)((List)M02List.get(j)).get( 7));
							dataList.add((String)((List)M02List.get(j)).get( 8));
							dataList.add((String)((List)M02List.get(j)).get( 9));
							dataList.add((String)((List)M02List.get(j)).get(10));
							dataList.add((String)((List)M02List.get(j)).get(11));
							dataList.add((String)((List)M02List.get(j)).get(12));
							dataList.add((String)((List)M02List.get(j)).get(13));
							dataList.add((String)((List)M02List.get(j)).get(14));
							dataList.add(String.valueOf(Integer.parseInt(m_year)));
							dataList.add(String.valueOf(Integer.parseInt(m_month)));
							dataList.add((String)((List)M02List.get(j)).get(0));
							dataList.add(dataRange);
							updateDBDataList.add(dataList);
						}//end of M02List					
			    		
			    	}
			    	if(updateDBDataList != null && updateDBDataList.size()!=0){						  
			 		  updateDBSqlList.add(sqlCmd.toString());
					  updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					  updateDBList.add(updateDBSqlList);
					}
			    }
			}//end of if 檔案上傳	
			
			if(input_method.equals("W")){//94.04.22 線上編輯check是否值全部為"0"
		    	nonZero = false;
		    	sqlCmd.delete(0,sqlCmd.length());
				paramList = new ArrayList();
				sqlCmd.append("select * from M02 where m_year=? AND m_month=?" );
				paramList.add(String.valueOf(Integer.parseInt(m_year)));
				paramList.add(String.valueOf(Integer.parseInt(m_month)));
				dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,guarantee_cnt,loan_amt,guarantee_amt,loan_bal,guarantee_bal,over_notpush_cnt,over_notpush_bal,over_okpush_cnt,over_okpush_bal,repay_tot_cnt,repay_tot_amt,repay_bal_cnt,repay_bal_amt");
		    	
				checkZeroLoop:
			    if(dbData != null && dbData.size() != 0){
			    	for(int zeroIdx=0;zeroIdx < dbData.size();zeroIdx++){			    		  
			    		if(	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("guarantee_cnt")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("loan_amt")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("guarantee_amt")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("loan_bal")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("guarantee_bal")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("over_notpush_cnt")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("over_notpush_bal")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("over_okpush_cnt")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("over_okpush_bal")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("repay_tot_cnt")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("repay_tot_amt")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("repay_bal_cnt")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("repay_bal_amt")).toString()) > 0						
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
				sqlCmd.append("select decode(loan_unit_no,'0','0','1') as \"type\", ");
                sqlCmd.append("       a.data_range, b.input_order,");
                sqlCmd.append("       data_range_name || '加總與總計不合' as \"errmsg\", ");
                sqlCmd.append("       sum(guarantee_cnt) as \"guarantee_cnt\", ");
                sqlCmd.append("       sum(loan_amt) as \"loan_amt\", ");
                sqlCmd.append("       sum(guarantee_amt) as \"guarantee_amt\", ");
                sqlCmd.append("       sum(loan_bal) as \"loan_bal\", ");
                sqlCmd.append("       sum(guarantee_bal) as \"guarantee_bal\", ");
                sqlCmd.append("       sum(over_notpush_cnt) as \"over_notpush_cnt\", ");
                sqlCmd.append("       sum(over_notpush_bal) as \"over_notpush_bal\", ");
                sqlCmd.append("       sum(over_okpush_bal) as \"over_okpush_bal\", ");
                sqlCmd.append("       sum(over_okpush_bal) as \"over_okpush_bal\", ");
                sqlCmd.append("       sum(repay_tot_cnt) as \"repay_tot_cnt\", ");
                sqlCmd.append("       sum(repay_tot_amt) as \"repay_tot_amt\", ");
                sqlCmd.append("       sum(repay_bal_cnt) as \"repay_bal_cnt\", ");
                sqlCmd.append("       sum(repay_bal_amt) as \"repay_bal_amt\" ");
                sqlCmd.append("from M02 a,m00_data_range_item b ");
                sqlCmd.append("where m_year=? and m_month=?");
                sqlCmd.append("  and substr(a.data_range,4,2)=b.data_range ");
                sqlCmd.append("  and b.report_no='M02' ");
                sqlCmd.append("group by decode(loan_unit_no,'0','0','1'),a.data_range,b.input_order,b.data_range_name || '加總與總計不合' ");
                sqlCmd.append("order by decode(loan_unit_no,'0','0','1') desc,b.input_order,b.data_range_name || '加總與總計不合' ");
                paramList.add(String.valueOf(Integer.parseInt(m_year)));
                paramList.add(String.valueOf(Integer.parseInt(m_month)));
                dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,guarantee_cnt,loan_amt,guarantee_amt,loan_bal,guarantee_bal,over_notpush_cnt,over_notpush_bal,over_okpush_cnt,over_okpush_bal,repay_tot_cnt,repay_tot_amt,repay_bal_cnt,repay_bal_amt");
           		                
                updateDBDataList = new ArrayList();	//99.11.18
				for(int j=0;j<3;j++){
					long guarantee_cnt		= Long.parseLong((((DataObject)dbData.get(j)).getValue("guarantee_cnt")		==null?"0":((DataObject)dbData.get(j)).getValue("guarantee_cnt")).toString());
					long loan_amt			= Long.parseLong((((DataObject)dbData.get(j)).getValue("loan_amt")			==null?"0":((DataObject)dbData.get(j)).getValue("loan_amt")).toString());
					long guarantee_amt 		= Long.parseLong((((DataObject)dbData.get(j)).getValue("guarantee_amt")		==null?"0":((DataObject)dbData.get(j)).getValue("guarantee_amt")).toString());
					long loan_bal			= Long.parseLong((((DataObject)dbData.get(j)).getValue("loan_bal")			==null?"0":((DataObject)dbData.get(j)).getValue("loan_bal")).toString());
					long guarantee_bal  	= Long.parseLong((((DataObject)dbData.get(j)).getValue("guarantee_bal")		==null?"0":((DataObject)dbData.get(j)).getValue("guarantee_bal")).toString());
					long over_notpush_cnt	= Long.parseLong((((DataObject)dbData.get(j)).getValue("over_notpush_cnt")	==null?"0":((DataObject)dbData.get(j)).getValue("over_notpush_cnt")).toString());
					long over_notpush_bal	= Long.parseLong((((DataObject)dbData.get(j)).getValue("over_notpush_bal")	==null?"0":((DataObject)dbData.get(j)).getValue("over_notpush_bal")).toString());
					long over_okpush_cnt	= Long.parseLong((((DataObject)dbData.get(j)).getValue("over_okpush_cnt")		==null?"0":((DataObject)dbData.get(j)).getValue("over_okpush_cnt")).toString());
					long over_okpush_bal	= Long.parseLong((((DataObject)dbData.get(j)).getValue("over_okpush_bal")		==null?"0":((DataObject)dbData.get(j)).getValue("over_okpush_bal")).toString());
					long repay_tot_cnt		= Long.parseLong((((DataObject)dbData.get(j)).getValue("repay_tot_cnt")		==null?"0":((DataObject)dbData.get(j)).getValue("repay_tot_cnt")).toString());
					long repay_tot_amt		= Long.parseLong((((DataObject)dbData.get(j)).getValue("repay_tot_amt")		==null?"0":((DataObject)dbData.get(j)).getValue("repay_tot_amt")).toString());
					long repay_bal_cnt  	= Long.parseLong((((DataObject)dbData.get(j)).getValue("repay_bal_cnt")		==null?"0":((DataObject)dbData.get(j)).getValue("repay_bal_cnt")).toString());
					long repay_bal_amt  	= Long.parseLong((((DataObject)dbData.get(j)).getValue("repay_bal_amt")		==null?"0":((DataObject)dbData.get(j)).getValue("repay_bal_amt")).toString());

					long t_guarantee_cnt	= Long.parseLong((((DataObject)dbData.get(j+3)).getValue("guarantee_cnt")		==null?"0":((DataObject)dbData.get(j+3)).getValue("guarantee_cnt")).toString());
					long t_loan_amt			= Long.parseLong((((DataObject)dbData.get(j+3)).getValue("loan_amt")			==null?"0":((DataObject)dbData.get(j+3)).getValue("loan_amt")).toString());
					long t_guarantee_amt 	= Long.parseLong((((DataObject)dbData.get(j+3)).getValue("guarantee_amt")		==null?"0":((DataObject)dbData.get(j+3)).getValue("guarantee_amt")).toString());
					long t_loan_bal			= Long.parseLong((((DataObject)dbData.get(j+3)).getValue("loan_bal")			==null?"0":((DataObject)dbData.get(j+3)).getValue("loan_bal")).toString());
					long t_guarantee_bal  	= Long.parseLong((((DataObject)dbData.get(j+3)).getValue("guarantee_bal")		==null?"0":((DataObject)dbData.get(j+3)).getValue("guarantee_bal")).toString());
					long t_over_notpush_cnt	= Long.parseLong((((DataObject)dbData.get(j+3)).getValue("over_notpush_cnt")	==null?"0":((DataObject)dbData.get(j+3)).getValue("over_notpush_cnt")).toString());
					long t_over_notpush_bal	= Long.parseLong((((DataObject)dbData.get(j+3)).getValue("over_notpush_bal")	==null?"0":((DataObject)dbData.get(j+3)).getValue("over_notpush_bal")).toString());
					long t_over_okpush_cnt	= Long.parseLong((((DataObject)dbData.get(j+3)).getValue("over_okpush_cnt")	==null?"0":((DataObject)dbData.get(j+3)).getValue("over_okpush_cnt")).toString());
					long t_over_okpush_bal	= Long.parseLong((((DataObject)dbData.get(j+3)).getValue("over_okpush_bal")	==null?"0":((DataObject)dbData.get(j+3)).getValue("over_okpush_bal")).toString());
					long t_repay_tot_cnt	= Long.parseLong((((DataObject)dbData.get(j+3)).getValue("repay_tot_cnt")		==null?"0":((DataObject)dbData.get(j+3)).getValue("repay_tot_cnt")).toString());
					long t_repay_tot_amt	= Long.parseLong((((DataObject)dbData.get(j+3)).getValue("repay_tot_amt")		==null?"0":((DataObject)dbData.get(j+3)).getValue("repay_tot_amt")).toString());
					long t_repay_bal_cnt  	= Long.parseLong((((DataObject)dbData.get(j+3)).getValue("repay_bal_cnt")		==null?"0":((DataObject)dbData.get(j+3)).getValue("repay_bal_cnt")).toString());
					long t_repay_bal_amt  	= Long.parseLong((((DataObject)dbData.get(j+3)).getValue("repay_bal_amt")		==null?"0":((DataObject)dbData.get(j+3)).getValue("repay_bal_amt")).toString());
					
					String subErrMsg="";
					
					if(guarantee_cnt!=t_guarantee_cnt) {
						subErrMsg="保證件數        :"+(String)((DataObject)dbData.get(j)).getValue("errmsg");
					}
					if(subErrMsg.length()!=0){
						updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,subErrMsg,user_id,user_name));								
						errCount++ ;						
						subErrMsg = "";
					}
					if(loan_amt	!= t_loan_amt) {
						subErrMsg="貸款金額        :"+(String)((DataObject)dbData.get(j)).getValue("errmsg");
					}
					if(subErrMsg.length()!=0){
					   updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,subErrMsg,user_id,user_name));								
					   errCount++ ;						
					   subErrMsg = "";
					}
					if(guarantee_amt != t_guarantee_amt ) {
						subErrMsg="保證金額        :"+(String)((DataObject)dbData.get(j)).getValue("errmsg");
					}
					if(subErrMsg.length()!=0){
						updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,subErrMsg,user_id,user_name));								
						errCount++ ;						
						subErrMsg = "";
					}
					if(loan_bal	!= t_loan_bal) {
						subErrMsg="貸款餘額        :"+(String)((DataObject)dbData.get(j)).getValue("errmsg");
					}
					if(subErrMsg.length()!=0){
						updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,subErrMsg,user_id,user_name));								
						errCount++ ;						
						subErrMsg = "";
					}
					if(guarantee_bal  != t_guarantee_bal ) {
						subErrMsg="保證餘額        :"+(String)((DataObject)dbData.get(j)).getValue("errmsg");
					}
					if(subErrMsg.length()!=0){
						updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,subErrMsg,user_id,user_name));								
						errCount++ ;						
						subErrMsg = "";
					}
					if(over_notpush_cnt != t_over_notpush_cnt ) {
						subErrMsg="逾期未轉催收件數:"+(String)((DataObject)dbData.get(j)).getValue("errmsg");
					}
					if(subErrMsg.length()!=0){
						updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,subErrMsg,user_id,user_name));								
						errCount++ ;						
						subErrMsg = "";
					}
					if(over_notpush_bal != t_over_notpush_bal ) {
						subErrMsg="逾期未轉催收餘額:"+(String)((DataObject)dbData.get(j)).getValue("errmsg");
					}
					if(subErrMsg.length()!=0){
						updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,subErrMsg,user_id,user_name));								
						errCount++ ;						
						subErrMsg = "";
					}
					if(over_okpush_cnt	!= t_over_okpush_cnt) {
						subErrMsg="逾期已轉催收件數:"+(String)((DataObject)dbData.get(j)).getValue("errmsg");
					}
					if(subErrMsg.length()!=0){
						updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,subErrMsg,user_id,user_name));								
						errCount++ ;						
						subErrMsg = "";
					}
					if(over_okpush_bal	!= t_over_okpush_bal	) {
						subErrMsg="逾期已轉催收餘額:"+(String)((DataObject)dbData.get(j)).getValue("errmsg");
					}
					if(subErrMsg.length()!=0){
						updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,subErrMsg,user_id,user_name));								
						errCount++ ;						
						subErrMsg = "";
					}
					if(repay_tot_cnt	!= t_repay_tot_cnt) {
						subErrMsg="代位清償總額件數:"+(String)((DataObject)dbData.get(j)).getValue("errmsg");
					}
					if(subErrMsg.length()!=0){
						updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,subErrMsg,user_id,user_name));								
						errCount++ ;						
						subErrMsg = "";
					}
					if(repay_tot_amt	!= t_repay_tot_amt	) {
						subErrMsg="代位清償總金額  :"+(String)((DataObject)dbData.get(j)).getValue("errmsg");
					}
					if(subErrMsg.length()!=0){
						updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,subErrMsg,user_id,user_name));								
						errCount++ ;						
						subErrMsg = "";
					}
					if(repay_bal_cnt  	!= t_repay_bal_cnt  ) {
						subErrMsg="代位清償淨額件數:"+(String)((DataObject)dbData.get(j)).getValue("errmsg");
					}
					if(subErrMsg.length()!=0){
						updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,subErrMsg,user_id,user_name));								
						errCount++ ;						
						subErrMsg = "";
					}
					if(repay_bal_amt  	!= t_repay_bal_amt  	) {
						subErrMsg="代位清償淨額    :"+(String)((DataObject)dbData.get(j)).getValue("errmsg");
					}
					if(subErrMsg.length()!=0){
						updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,subErrMsg,user_id,user_name));								
						errCount++ ;						
						subErrMsg = "";
					}
				}//end of for
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
			/*99.11.19移至Utility.getWML01讀取WML01 all data	*/			
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
				errMsg = errMsg + "UpdateM02.doParserReport_M02 UpdateDB Error:"+DBManager.getErrMsg()+"<br>";
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
			errMsg = errMsg + "UpdateM02.doParserReport_M02 Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateM02.doParserReport_M02="+e.getMessage());			
		}
		return errMsg;		
	}

	//讀取上傳檔案的資料
	//99.11.17 add 套用DAO.preparestatment,並列印轉換後的SQL by 2295
	private static List getM02FileData(String report_no,String filename){
			String	txtline	 = null;			
			List M02List = new LinkedList();
			List detail = null;
			List dbData = null;			
			
			try {
				String WMdataDir = Utility.getProperties("WMdataDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");
				String WMdataBKDir = Utility.getProperties("WMdataBKDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");		
				File tmpFile = new File(WMdataDir+filename);		
				Date filedate = new Date(tmpFile.lastModified());
					
				FileReader	f		= new FileReader(tmpFile);
				LineNumberReader in	= new LineNumberReader(f);
				doLoop://將txt檔案儲存至LinkedList
				while ((txtline = in.readLine()) != null) {
						if (txtline.trim().equals("")){
							continue doLoop;
						}else{
							detail=new LinkedList();
							detail.add(txtline.substring(  0,  1));	//保證項目
							detail.add(txtline.substring(  1,  6));	//資料期間
							detail.add(txtline.substring(  6, 13));	//保證件數         
							detail.add(txtline.substring( 13, 27));	//貸款金額         
							detail.add(txtline.substring( 27, 41));	//保證金額         
							detail.add(txtline.substring( 41, 55));	//貸款餘額         
							detail.add(txtline.substring( 55, 69));	//保證餘額         
							detail.add(txtline.substring( 69, 76));	//逾期未轉催收件數 
							detail.add(txtline.substring( 76, 87));	//逾期未轉催收餘額 
							detail.add(txtline.substring( 87, 94));	//逾期已轉催收件數 
							detail.add(txtline.substring( 94,105));	//逾期已轉催收餘額 
							detail.add(txtline.substring(105,112));	//代位清償總額件數 
							detail.add(txtline.substring(112,123));	//代位清償總金額   
							detail.add(txtline.substring(123,130));	//代位清償淨額件數 
							detail.add(txtline.substring(130,141));	//代位清償淨額     

							M02List.add(detail);
						}
				}	
				in.close();
				f.close();
			}catch(Exception e){
				errMsg = errMsg + "UpdateM02.getM02FileData Error:"+e.getMessage()+"<br>";
			}
			return M02List;
	}
	
	//Insert "0" 至M02
	//99.11.17 add 套用DAO.preparestatment,並列印轉換後的SQL by 2295
	private static boolean InsertZeroM02(String m_year,String m_month,String bank_code){
		List paramList = new ArrayList();
		StringBuffer sqlCmd = new StringBuffer();	
		List updateDBList = new LinkedList();//0:sql 1:data		
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data
		boolean updateOK=false;
		
		try{
			sqlCmd.append("select a.loan_unit_no,b.data_range ");
			sqlCmd.append("from m00_loan_unit a,m00_data_range_item b ");
			sqlCmd.append("where ((a.loan_unit_no<>'0' and b.data_range_type<>'T') or ");
			sqlCmd.append("       (a.loan_unit_no='0' and b.data_range_type='T') ");
			sqlCmd.append("      )  ");
			sqlCmd.append("  and b.report_no='M02' ");
			sqlCmd.append("order by a.input_order,b.input_order ");
			
			List data_div01 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),null,"update_date");//資負表 	    
			sqlCmd.delete(0, sqlCmd.length());
			sqlCmd.append("INSERT INTO M02( M_YEAR,M_MONTH,loan_unit_no,DATA_RANGE ) VALUES (?,?,?,?)");
			for(int d1=0;d1<data_div01.size();d1++){
				String sm_year="";
				sm_year="0000"+m_year; sm_year=sm_year.substring(sm_year.length()-3);
				dataList = new LinkedList();//儲存參數的data
				dataList.add(String.valueOf(Integer.parseInt(m_year)));
				dataList.add(String.valueOf(Integer.parseInt(m_month)));
				dataList.add((String)((DataObject)data_div01.get(d1)).getValue("loan_unit_no"));
				dataList.add(sm_year+(String)((DataObject)data_div01.get(d1)).getValue("data_range"));
				updateDBDataList.add(dataList);//1:傳內的參數List				
			}
			if(updateDBDataList != null && updateDBDataList.size()!=0){
			   updateDBSqlList.add(sqlCmd.toString());
			   updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
			   updateDBList.add(updateDBSqlList);
			}
			updateOK=DBManager.updateDB_ps(updateDBList);
			System.out.println("M02 Insert Zero OK??"+updateOK);
			if(!updateOK){
				errMsg = errMsg + "M02 Insert Zero Fail:"+DBManager.getErrMsg()+"<br>";
				System.out.println(DBManager.getErrMsg());
			}
		}catch(Exception e){
			errMsg = errMsg + "UpdateM02.InsertZeroM02 Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateM02.InsertZeroM02 Error:"+e.getMessage());
		}
    	return updateOK;
	}		
}
