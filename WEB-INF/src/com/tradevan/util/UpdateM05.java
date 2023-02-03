//93.02.23 to fix upload error by egg
//94.04.22 fix 都為"0"值時,顯示檢核為"0" by 2295
//94.05.17 fix 當M05List長度為18時,才檢核是否都為"0" by 2295
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


public class UpdateM05 {
	private static String errMsg = "";
	public String getErrMsg(){
		return errMsg;  
	}
	public static synchronized String doParserReport_M05(String report_no, String m_year,
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
		String  bank_code=srcbank_code;//農業信用保證基金 	9700002";//農業信用保證基金 
		String  add_date="";//申報日期
		String  upd_code="";//檢核結果
		String  common_center="";//由共用中心傳入
		boolean updateOK=false;
		SimpleDateFormat bkformat = new SimpleDateFormat("yyyyMMddHHmmss");
		SimpleDateFormat emailformat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");	
		String WMdataDir = Utility.getProperties("WMdataDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");
		String WMdataBKDir = Utility.getProperties("WMdataBKDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");	
		
		List M05List = new LinkedList();//M05細部資料
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
				M05List = getM05FileData(report_no,filename);
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
			}
			
			if(input_method.equals("F")){//檔案上傳
				sqlCmd.delete(0,sqlCmd.length());
				paramList = new ArrayList();			
				sqlCmd.append("select * from "+report_no+" where m_year=? AND m_month=?") ;
				paramList.add(String.valueOf(Integer.parseInt(m_year)));
				paramList.add(String.valueOf(Integer.parseInt(m_month)));
				
				dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month");
				if(dbData.size() == 0){//M05無資料時,才insert一筆Zero的資料
					System.out.println("InsertZeroM05");
					InsertZeroM05(m_year,m_month,bank_code);
				}else{//end of M05不存在時
					updateDBDataList = new ArrayList();	
					updateDBSqlList = new ArrayList();
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" INSERT INTO M05_LOG ");
					sqlCmd.append(" select m_year,m_month,loan_unit_no,period_no,item_no,");
					sqlCmd.append("        repay_cnt,repay_amt,run_notgood_cnt,run_notgood_amt,turn_out_cnt,");
					sqlCmd.append("        turn_out_amt,diease_cnt,dieaserepay_amt,disaster_cnt,disaster_amt,");
					sqlCmd.append("        corun_out_cnt,corun_out_amt,other_cnt,other_amt ");
					sqlCmd.append("        ,?,?,sysdate,'U'");
					sqlCmd.append("   from M05");
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
					
					updateDBDataList = new ArrayList();	
					updateDBSqlList = new ArrayList();
					sqlCmd.delete(0,sqlCmd.length());
		    		sqlCmd.append(" INSERT INTO M05_TOTACC_LOG ");
					sqlCmd.append(" select m_year,m_month,loan_unit_no,fix_no,guarantee_no_totacc,guarantee_amt_totacc");
					sqlCmd.append("        ,?,?,sysdate,'U'");
					sqlCmd.append("   from M05_TOTACC");
					sqlCmd.append(" WHERE m_year=? AND m_month=?"); 
							
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
					
					updateDBDataList = new ArrayList();	
					updateDBSqlList = new ArrayList();
					sqlCmd.delete(0,sqlCmd.length());					
					sqlCmd.append(" INSERT INTO M05_NOTE_LOG ");
					sqlCmd.append(" select m_year,m_month,note_no,note_amt_rate,?,?,sysdate,'U'");
					sqlCmd.append("   from M05_NOTE");
					sqlCmd.append(" WHERE m_year=? AND m_month=?"); 
					updateDBDataList.add(dataList);
					updateDBSqlList.add(sqlCmd.toString());
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);	
				}
			    nonZero = false;//94.04.24
			    List updateDBDataList_M05 = new ArrayList();
			    List updateDBDataList_M05Totacc = new ArrayList();	
			    List updateDBDataList_M05Note = new ArrayList();	
			    for(int j=0;j<M05List.size();j++){//把上傳檔案裡的資料更新至M05				
			    	String tmp_reportNo=(String)((List)M05List.get(j)).get(0);
			    	//94.04.24 add 檢核上傳檔案內容的金額是否都為"0"
			    	//94.05.17 fix 當M05List長度為18時,才檢核是否都為"0"			    	
			    	if(((List)M05List.get(j)).size() == 18){			    		
		    		    if(  ((!((String)((List)M05List.get(j)).get(4)).equals("")) && Long.parseLong((String)((List)M05List.get(j)).get(4)) > 0 )
		    		      || ((!((String)((List)M05List.get(j)).get(5)).equals("")) && Long.parseLong((String)((List)M05List.get(j)).get(5)) > 0 )
		    		      || ((!((String)((List)M05List.get(j)).get(6)).equals("")) && Long.parseLong((String)((List)M05List.get(j)).get(6)) > 0 )
					      || ((!((String)((List)M05List.get(j)).get(7)).equals("")) && Long.parseLong((String)((List)M05List.get(j)).get(7)) > 0 )
					      || ((!((String)((List)M05List.get(j)).get(8)).equals("")) && Long.parseLong((String)((List)M05List.get(j)).get(8)) > 0 )
					      || ((!((String)((List)M05List.get(j)).get(9)).equals("")) && Long.parseLong((String)((List)M05List.get(j)).get(9)) > 0 )
					      || ((!((String)((List)M05List.get(j)).get(10)).equals("")) && Long.parseLong((String)((List)M05List.get(j)).get(10)) > 0 )
		    		      || ((!((String)((List)M05List.get(j)).get(11)).equals("")) && Long.parseLong((String)((List)M05List.get(j)).get(11)) > 0 )
		    		      || ((!((String)((List)M05List.get(j)).get(12)).equals("")) && Long.parseLong((String)((List)M05List.get(j)).get(12)) > 0 )
		    		      || ((!((String)((List)M05List.get(j)).get(13)).equals("")) && Long.parseLong((String)((List)M05List.get(j)).get(13)) > 0 )
		    		      || ((!((String)((List)M05List.get(j)).get(14)).equals("")) && Long.parseLong((String)((List)M05List.get(j)).get(14)) > 0 )
		    		      || ((!((String)((List)M05List.get(j)).get(15)).equals("")) && Long.parseLong((String)((List)M05List.get(j)).get(15)) > 0 )
		    		      || ((!((String)((List)M05List.get(j)).get(16)).equals("")) && Long.parseLong((String)((List)M05List.get(j)).get(16)) > 0 )
		    		      || ((!((String)((List)M05List.get(j)).get(17)).equals("")) && Long.parseLong((String)((List)M05List.get(j)).get(17)) > 0 ))
		    		    {
		    		       nonZero = true;
		    		     }//end of > 0
			    	}//end of size=18
			    	if(nonZero){
			    	   if(tmp_reportNo.equals("M05")){
			    	   		dataList =  new ArrayList();
					     	dataList.add((String)((List)M05List.get(j)).get( 4));
					     	dataList.add((String)((List)M05List.get(j)).get( 5));
					     	dataList.add((String)((List)M05List.get(j)).get( 6));
					     	dataList.add((String)((List)M05List.get(j)).get( 7));
					     	dataList.add((String)((List)M05List.get(j)).get( 8));
					     	dataList.add((String)((List)M05List.get(j)).get( 9));
					     	dataList.add((String)((List)M05List.get(j)).get(10));    
					     	dataList.add((String)((List)M05List.get(j)).get(11));     
					     	dataList.add((String)((List)M05List.get(j)).get(12));         
					     	dataList.add((String)((List)M05List.get(j)).get(13));         
					     	dataList.add((String)((List)M05List.get(j)).get(14)); 		
					     	dataList.add((String)((List)M05List.get(j)).get(15)); 		
					     	dataList.add((String)((List)M05List.get(j)).get(16)); 		
					     	dataList.add((String)((List)M05List.get(j)).get(17)); 						
					     	dataList.add(String.valueOf(Integer.parseInt(m_year)));
					     	dataList.add(String.valueOf(Integer.parseInt(m_month)));
					     	dataList.add((String)((List)M05List.get(j)).get(1));	
					     	dataList.add((String)((List)M05List.get(j)).get(2));
					     	dataList.add((String)((List)M05List.get(j)).get(3));					
					     	updateDBDataList_M05.add(dataList);	
				       }else if(tmp_reportNo.equals("M05_TOTACC")){	//M05_TOTACC
				       		dataList =  new ArrayList();
				       		dataList.add((String)((List)M05List.get(j)).get( 3));
				       		dataList.add((String)((List)M05List.get(j)).get( 4));
					     	dataList.add(String.valueOf(Integer.parseInt(m_year)));
					     	dataList.add(String.valueOf(Integer.parseInt(m_month)));
					     	dataList.add((String)((List)M05List.get(j)).get(1));	
					     	dataList.add((String)((List)M05List.get(j)).get(2));					
					     	updateDBDataList_M05Totacc.add(dataList);			    		
					   }else{//M05Note
					   		dataList =  new ArrayList();
					   		dataList.add((String)((List)M05List.get(j)).get( 2));				    	
					     	dataList.add(String.valueOf(Integer.parseInt(m_year)));
					     	dataList.add(String.valueOf(Integer.parseInt(m_month)));
					     	dataList.add((String)((List)M05List.get(j)).get(1));
					     	updateDBDataList_M05Note.add(dataList);
					   }
			    	}//end of nonZero
			    }//end of M05List
			    if(updateDBDataList_M05 != null && updateDBDataList_M05.size()!=0){	
			    	updateDBSqlList = new ArrayList();
					sqlCmd.delete(0,sqlCmd.length());
			        sqlCmd.append("UPDATE M05 SET repay_cnt       	=?,");
				    sqlCmd.append("               repay_amt         =?,");
				    sqlCmd.append("               run_notgood_cnt   =?,");
				    sqlCmd.append("               run_notgood_amt	=?,");
				    sqlCmd.append("               turn_out_cnt      =?,");
				    sqlCmd.append("               turn_out_amt   	=?,");
				    sqlCmd.append("               diease_cnt    	=?,");
				    sqlCmd.append("               dieaserepay_amt	=?,");
				    sqlCmd.append("               disaster_cnt		=?,");
				    sqlCmd.append("               disaster_amt		=?,");
				    sqlCmd.append("               corun_out_cnt		=?,");
				    sqlCmd.append("               corun_out_amt		=?,");
				    sqlCmd.append("               other_cnt			=?,");
				    sqlCmd.append("               other_amt			=? ");
			        sqlCmd.append(" WHERE m_year=? AND m_month=?"); 
			        sqlCmd.append("   AND loan_unit_no=?");
			        sqlCmd.append("   AND period_no=?");
			        sqlCmd.append("   AND item_no=?");
			        updateDBSqlList.add(sqlCmd.toString());
					updateDBSqlList.add(updateDBDataList_M05);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
			    }
			    if(updateDBDataList_M05Totacc != null && updateDBDataList_M05Totacc.size()!=0){	
			    	updateDBSqlList = new ArrayList();
					sqlCmd.delete(0,sqlCmd.length());				
	    	        sqlCmd.append("UPDATE M05_TOTACC SET guarantee_no_totacc  =?,");
			        sqlCmd.append("                      guarantee_amt_totacc =? ");
		            sqlCmd.append(" WHERE m_year=? AND m_month=?");
		            sqlCmd.append("   AND loan_unit_no=?");
		            sqlCmd.append("   AND fix_no=?");
		            updateDBSqlList.add(sqlCmd.toString());
					updateDBSqlList.add(updateDBDataList_M05Totacc);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
				}
			    if(updateDBDataList_M05Note != null && updateDBDataList_M05Note.size()!=0){	
			    	updateDBSqlList = new ArrayList();
					sqlCmd.delete(0,sqlCmd.length());
	    	        sqlCmd.append("UPDATE M05_NOTE SET note_amt_rate =?");
		            sqlCmd.append(" WHERE m_year=? AND m_month=?"); 
		            sqlCmd.append("   AND note_no=?");
		            updateDBSqlList.add(sqlCmd.toString());
					updateDBSqlList.add(updateDBDataList_M05Note);//0:sql 1:參數List
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
			    System.out.println("檔案上傳.nonZero="+nonZero);			    
			    System.out.println("M05 update data OK??"+updateOK);			   
			}//end of if 檔案上傳	
			
			System.out.println("檢查合計和小計資料開始");
			System.out.println("input_method="+input_method);
			if(input_method.equals("W")){//94.04.22 線上編輯check是否值全部為"0"
		    	nonZero = false;
		    	sqlCmd.delete(0,sqlCmd.length());
				paramList = new ArrayList();
				sqlCmd.append("select * from M05 where m_year=? AND m_month=?" );
				paramList.add(String.valueOf(Integer.parseInt(m_year)));
				paramList.add(String.valueOf(Integer.parseInt(m_month)));
				dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"repay_cnt,repay_amt,run_notgood_cnt,run_notgood_amt,turn_out_cnt,turn_out_amt,diease_cnt,dieaserepay_amt,disaster_cnt,disaster_amt,corun_out_cnt,corun_out_amt,other_cnt,other_amt");
		    	   
		    	checkZeroLoop:
			    if(dbData != null && dbData.size() != 0){
			    	for(int zeroIdx=0;zeroIdx < dbData.size();zeroIdx++){			    		  
			    		if(	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("repay_cnt")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("repay_amt")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("run_notgood_cnt")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("run_notgood_amt")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("turn_out_cnt")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("turn_out_amt")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("diease_cnt")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("dieaserepay_amt")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("disaster_cnt")).toString()) > 0												
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("disaster_amt")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("corun_out_cnt")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("corun_out_amt")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("other_cnt")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("other_amt")).toString()) > 0
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
				sqlCmd.append("select decode(loan_unit_no,'0','0','1') as \"type\", period_no, item_no, ");
				sqlCmd.append("       decode(period_no,'0MT','本月','0YT','本年','0TT','累計',period_no) || ");
				sqlCmd.append("       decode(item_no,'RP','代償','RC','獲償','RT','小計',item_no) as \"errmsg\" , ");
				sqlCmd.append("       sum(repay_cnt) as \"repay_cnt\", ");
				sqlCmd.append("       sum(repay_amt) as \"repay_amt\", ");
				sqlCmd.append("       sum(run_notgood_cnt) as \"run_notgood_cnt\", ");
				sqlCmd.append("       sum(run_notgood_amt) as \"run_notgood_amt\", ");
				sqlCmd.append("       sum(turn_out_cnt) as \"turn_out_cnt\", ");
				sqlCmd.append("       sum(turn_out_amt) as \"turn_out_amt\", ");
				sqlCmd.append("       sum(diease_cnt) as \"diease_cnt\", ");
				sqlCmd.append("       sum(dieaserepay_amt) as \"dieaserepay_amt\", ");
				sqlCmd.append("       sum(disaster_cnt) as \"disaster_cnt\", ");
				sqlCmd.append("       sum(disaster_amt) as \"disaster_amt\", ");
				sqlCmd.append("       sum(corun_out_cnt) as \"corun_out_cnt\", ");
				sqlCmd.append("       sum(corun_out_amt) as \"corun_out_amt\", ");
				sqlCmd.append("       sum(other_cnt) as \"other_cnt\", ");
				sqlCmd.append("       sum(other_amt) as \"other_amt\"  ");
				sqlCmd.append("  from m05 ");
				sqlCmd.append(" where m_year=?");
				sqlCmd.append(" and period_no != 'NT1' ");	//93.02.23 add by egg to fix upload error
				sqlCmd.append(" and m_month=?");
				sqlCmd.append(" group by decode(loan_unit_no,'0','0','1'),period_no,item_no ");
				sqlCmd.append(" order by decode(loan_unit_no,'0','0','1') desc,period_no,item_no ");
				paramList.add(String.valueOf(Integer.parseInt(m_year)));
	            paramList.add(String.valueOf(Integer.parseInt(m_month)));	           	
                dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,repay_cnt,repay_amt,run_notgood_cnt,run_notgood_amt,turn_out_cnt,turn_out_amt,diease_cnt,dieaserepay_amt,disaster_cnt,disaster_amt,corun_out_cnt,corun_out_amt,other_cnt,other_amt");
                
                updateDBDataList = new ArrayList();//99.11.19	
                //fix 13 to 9 93.02.23 by egg
				for(int j=0;j<9;j++){
					System.out.println("目前j="+j);
					long repay_cnt			= Long.parseLong((((DataObject)dbData.get(j)).getValue("repay_cnt")			==null?"0":((DataObject)dbData.get(j)).getValue("repay_cnt")).toString());
					long repay_amt       	= Long.parseLong((((DataObject)dbData.get(j)).getValue("repay_amt")			==null?"0":((DataObject)dbData.get(j)).getValue("repay_amt")).toString());
					long run_notgood_cnt    = Long.parseLong((((DataObject)dbData.get(j)).getValue("run_notgood_cnt")	==null?"0":((DataObject)dbData.get(j)).getValue("run_notgood_cnt")).toString());
					long run_notgood_amt	= Long.parseLong((((DataObject)dbData.get(j)).getValue("run_notgood_amt")	==null?"0":((DataObject)dbData.get(j)).getValue("run_notgood_amt")).toString());
					long turn_out_cnt 		= Long.parseLong((((DataObject)dbData.get(j)).getValue("turn_out_cnt")		==null?"0":((DataObject)dbData.get(j)).getValue("turn_out_cnt")).toString());
					long turn_out_amt    	= Long.parseLong((((DataObject)dbData.get(j)).getValue("turn_out_amt")		==null?"0":((DataObject)dbData.get(j)).getValue("turn_out_amt")).toString());
					long diease_cnt      	= Long.parseLong((((DataObject)dbData.get(j)).getValue("diease_cnt")		==null?"0":((DataObject)dbData.get(j)).getValue("diease_cnt")).toString());
					long dieaserepay_amt 	= Long.parseLong((((DataObject)dbData.get(j)).getValue("dieaserepay_amt")	==null?"0":((DataObject)dbData.get(j)).getValue("dieaserepay_amt")).toString());
					long disaster_cnt    	= Long.parseLong((((DataObject)dbData.get(j)).getValue("disaster_cnt")		==null?"0":((DataObject)dbData.get(j)).getValue("disaster_cnt")).toString());
					long disaster_amt     	= Long.parseLong((((DataObject)dbData.get(j)).getValue("disaster_amt")		==null?"0":((DataObject)dbData.get(j)).getValue("disaster_amt")).toString());
					long corun_out_cnt   	= Long.parseLong((((DataObject)dbData.get(j)).getValue("corun_out_cnt")		==null?"0":((DataObject)dbData.get(j)).getValue("corun_out_cnt")).toString());
					long corun_out_amt   	= Long.parseLong((((DataObject)dbData.get(j)).getValue("corun_out_amt")		==null?"0":((DataObject)dbData.get(j)).getValue("corun_out_amt")).toString());
					long other_cnt   		= Long.parseLong((((DataObject)dbData.get(j)).getValue("other_cnt")			==null?"0":((DataObject)dbData.get(j)).getValue("other_cnt")).toString());
					long other_amt   		= Long.parseLong((((DataObject)dbData.get(j)).getValue("other_amt")			==null?"0":((DataObject)dbData.get(j)).getValue("other_amt")).toString());
					String errmsg			= (String)((DataObject)dbData.get(j)).getValue("errmsg");
                      
					long t_repay_cnt       	= Long.parseLong((((DataObject)dbData.get(j+9)).getValue("repay_cnt")		==null?"0":((DataObject)dbData.get(j+9)).getValue("repay_cnt")).toString());
					long t_repay_amt       	= Long.parseLong((((DataObject)dbData.get(j+9)).getValue("repay_amt")		==null?"0":((DataObject)dbData.get(j+9)).getValue("repay_amt")).toString());
					long t_run_notgood_cnt 	= Long.parseLong((((DataObject)dbData.get(j+9)).getValue("run_notgood_cnt")	==null?"0":((DataObject)dbData.get(j+9)).getValue("run_notgood_cnt")).toString());
					long t_run_notgood_amt	= Long.parseLong((((DataObject)dbData.get(j+9)).getValue("run_notgood_amt")	==null?"0":((DataObject)dbData.get(j+9)).getValue("run_notgood_amt")).toString());
					long t_turn_out_cnt 	= Long.parseLong((((DataObject)dbData.get(j+9)).getValue("turn_out_cnt")	==null?"0":((DataObject)dbData.get(j+9)).getValue("turn_out_cnt")).toString());
					long t_turn_out_amt    	= Long.parseLong((((DataObject)dbData.get(j+9)).getValue("turn_out_amt")	==null?"0":((DataObject)dbData.get(j+9)).getValue("turn_out_amt")).toString());
					long t_diease_cnt      	= Long.parseLong((((DataObject)dbData.get(j+9)).getValue("diease_cnt")		==null?"0":((DataObject)dbData.get(j+9)).getValue("diease_cnt")).toString());
					long t_dieaserepay_amt 	= Long.parseLong((((DataObject)dbData.get(j+9)).getValue("dieaserepay_amt")	==null?"0":((DataObject)dbData.get(j+9)).getValue("dieaserepay_amt")).toString());
					long t_disaster_cnt    	= Long.parseLong((((DataObject)dbData.get(j+9)).getValue("disaster_cnt")	==null?"0":((DataObject)dbData.get(j+9)).getValue("disaster_cnt")).toString());
					long t_disaster_amt     = Long.parseLong((((DataObject)dbData.get(j+9)).getValue("disaster_amt")	==null?"0":((DataObject)dbData.get(j+9)).getValue("disaster_amt")).toString());
					long t_corun_out_cnt   	= Long.parseLong((((DataObject)dbData.get(j+9)).getValue("corun_out_cnt")	==null?"0":((DataObject)dbData.get(j+9)).getValue("corun_out_cnt")).toString());
					long t_corun_out_amt   	= Long.parseLong((((DataObject)dbData.get(j+9)).getValue("corun_out_amt")	==null?"0":((DataObject)dbData.get(j+9)).getValue("corun_out_amt")).toString());
					long t_other_cnt   		= Long.parseLong((((DataObject)dbData.get(j+9)).getValue("other_cnt")		==null?"0":((DataObject)dbData.get(j+9)).getValue("other_cnt")).toString());
					long t_other_amt   		= Long.parseLong((((DataObject)dbData.get(j+9)).getValue("other_amt")		==null?"0":((DataObject)dbData.get(j+9)).getValue("other_amt")).toString());
					
					System.out.println("repay_cnt="+repay_cnt+",t_repay_cnt="+t_repay_cnt);
					System.out.println("repay_amt="+repay_amt+",t_repay_amt="+t_repay_amt);
					if((checkRule(repay_cnt       	,t_repay_cnt       	,"代償案件件數  :"+errmsg,m_year,m_month,user_id,user_name,errCount,bank_code,report_no)).size() > 0){
					   updateDBDataList.add(checkRule(repay_cnt       	,t_repay_cnt       	,"代償案件件數  :"+errmsg,m_year,m_month,user_id,user_name,errCount,bank_code,report_no));
					   errCount++;
					}  
					
					if((checkRule(repay_amt       	,t_repay_amt       	,"代償案件金額  :"+errmsg,m_year,m_month,user_id,user_name,errCount,bank_code,report_no)).size() > 0){
					   updateDBDataList.add(checkRule(repay_amt       	,t_repay_amt       	,"代償案件金額  :"+errmsg,m_year,m_month,user_id,user_name,errCount,bank_code,report_no));
					   errCount++;
					} 
					
					if((checkRule(run_notgood_cnt 	,t_run_notgood_cnt 	,"經營不善件數  :"+errmsg,m_year,m_month,user_id,user_name,errCount,bank_code,report_no)).size() > 0){
					   updateDBDataList.add(checkRule(run_notgood_cnt 	,t_run_notgood_cnt 	,"經營不善件數  :"+errmsg,m_year,m_month,user_id,user_name,errCount,bank_code,report_no));
					   errCount++;
					} 
					
					if((checkRule(run_notgood_amt	,t_run_notgood_amt	,"經營不善金額  :"+errmsg,m_year,m_month,user_id,user_name,errCount,bank_code,report_no)).size() > 0){
					   updateDBDataList.add(checkRule(run_notgood_amt	,t_run_notgood_amt	,"經營不善金額  :"+errmsg,m_year,m_month,user_id,user_name,errCount,bank_code,report_no));
					   errCount++;
					} 
					
					if((checkRule(turn_out_cnt		,t_turn_out_cnt	,"週轉失靈件數  :"+errmsg,m_year,m_month,user_id,user_name,errCount,bank_code,report_no)).size() > 0){
					   updateDBDataList.add(checkRule(turn_out_cnt		,t_turn_out_cnt	,"週轉失靈件數  :"+errmsg,m_year,m_month,user_id,user_name,errCount,bank_code,report_no));
					   errCount++;
					} 
					
					if((checkRule(turn_out_amt    	,t_turn_out_amt    	,"週轉失靈金額  :"+errmsg,m_year,m_month,user_id,user_name,errCount,bank_code,report_no)).size() > 0){
					   updateDBDataList.add(checkRule(turn_out_amt    	,t_turn_out_amt    	,"週轉失靈金額  :"+errmsg,m_year,m_month,user_id,user_name,errCount,bank_code,report_no));
					   errCount++;
					} 
					
					if((checkRule(diease_cnt      	,t_diease_cnt      	,"疾病或死亡件數:"+errmsg,m_year,m_month,user_id,user_name,errCount,bank_code,report_no)).size() > 0){
					   updateDBDataList.add(checkRule(diease_cnt      	,t_diease_cnt      	,"疾病或死亡件數:"+errmsg,m_year,m_month,user_id,user_name,errCount,bank_code,report_no));
					   errCount++;
					} 
					
					if((checkRule(dieaserepay_amt 	,t_dieaserepay_amt 	,"疾病或死亡金額:"+errmsg,m_year,m_month,user_id,user_name,errCount,bank_code,report_no)).size() > 0){
					   updateDBDataList.add(checkRule(dieaserepay_amt 	,t_dieaserepay_amt 	,"疾病或死亡金額:"+errmsg,m_year,m_month,user_id,user_name,errCount,bank_code,report_no));
					   errCount++;
					} 
					
					if((checkRule(disaster_cnt    	,t_disaster_cnt    	,"災歉件數      :"+errmsg,m_year,m_month,user_id,user_name,errCount,bank_code,report_no)).size() > 0){
					   updateDBDataList.add(checkRule(disaster_cnt    	,t_disaster_cnt    	,"災歉件數      :"+errmsg,m_year,m_month,user_id,user_name,errCount,bank_code,report_no));
					   errCount++;
					} 
					if((checkRule(disaster_amt     ,t_disaster_amt     ,"災歉金額      :"+errmsg,m_year,m_month,user_id,user_name,errCount,bank_code,report_no)).size() > 0){
					   updateDBDataList.add(checkRule(disaster_amt     ,t_disaster_amt     ,"災歉金額      :"+errmsg,m_year,m_month,user_id,user_name,errCount,bank_code,report_no));
					   errCount++;
					} 
					if((checkRule(corun_out_cnt   	,t_corun_out_cnt   	,"兼業失利件數  :"+errmsg,m_year,m_month,user_id,user_name,errCount,bank_code,report_no)).size() > 0){
					   updateDBDataList.add(checkRule(corun_out_cnt   	,t_corun_out_cnt   	,"兼業失利件數  :"+errmsg,m_year,m_month,user_id,user_name,errCount,bank_code,report_no));
					   errCount++;
					} 
					if((checkRule(corun_out_amt   	,t_corun_out_amt   	,"兼業失利金額  :"+errmsg,m_year,m_month,user_id,user_name,errCount,bank_code,report_no)).size() > 0){
					   updateDBDataList.add(checkRule(corun_out_amt   	,t_corun_out_amt   	,"兼業失利金額  :"+errmsg,m_year,m_month,user_id,user_name,errCount,bank_code,report_no));
					   errCount++;
					} 
					if((checkRule(other_cnt   		,t_other_cnt   		,"其他件數      :"+errmsg,m_year,m_month,user_id,user_name,errCount,bank_code,report_no)).size() > 0){
					   updateDBDataList.add(checkRule(other_cnt   		,t_other_cnt   		,"其他件數      :"+errmsg,m_year,m_month,user_id,user_name,errCount,bank_code,report_no));
					   errCount++;
					} 
					if((checkRule(other_amt   		,t_other_amt   		,"其他金額      :"+errmsg,m_year,m_month,user_id,user_name,errCount,bank_code,report_no)).size() > 0){
					   updateDBDataList.add(checkRule(other_amt   		,t_other_amt   		,"其他金額      :"+errmsg,m_year,m_month,user_id,user_name,errCount,bank_code,report_no));
					   errCount++;
					} 
				}//end of for
				
				//檢核M05_TOTACC
				System.out.println("檢核M05_TOTACC...Start..");
				sqlCmd.delete(0,sqlCmd.length());
				sqlCmd.append("select decode(loan_unit_no,'0','0','1') as \"type\", ");
				sqlCmd.append("       sum(guarantee_no_totacc) as \"guarantee_no_totacc\", ");
				sqlCmd.append("       sum(guarantee_amt_totacc) as \"guarantee_amt_totacc\" ");
				sqlCmd.append("  from m05_totacc ");
				sqlCmd.append(" where m_year=? and m_month=?");
				sqlCmd.append(" group by decode(loan_unit_no,'0','0','1') ");
				sqlCmd.append(" order by decode(loan_unit_no,'0','0','1') desc ");
                dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,guarantee_no_totacc,guarantee_amt_totacc");		
				long guarantee_no_totacc   	= Long.parseLong((((DataObject)dbData.get(0)).getValue("guarantee_no_totacc")	==null?"0":((DataObject)dbData.get(0)).getValue("guarantee_no_totacc")).toString());
				long guarantee_amt_totacc   = Long.parseLong((((DataObject)dbData.get(0)).getValue("guarantee_amt_totacc")	==null?"0":((DataObject)dbData.get(0)).getValue("guarantee_amt_totacc")).toString());
    
				long t_guarantee_no_totacc	= Long.parseLong((((DataObject)dbData.get(1)).getValue("guarantee_no_totacc")	==null?"0":((DataObject)dbData.get(1)).getValue("guarantee_no_totacc")).toString());
				long t_guarantee_amt_totacc	= Long.parseLong((((DataObject)dbData.get(1)).getValue("guarantee_amt_totacc")	==null?"0":((DataObject)dbData.get(1)).getValue("guarantee_amt_totacc")).toString());
				if((checkRule(guarantee_no_totacc       	,t_guarantee_no_totacc       	,"累計保證件數",m_year,m_month,user_id,user_name,errCount,bank_code,report_no)).size() > 0){
				   updateDBDataList.add(checkRule(guarantee_no_totacc       	,t_guarantee_no_totacc       	,"累計保證件數",m_year,m_month,user_id,user_name,errCount,bank_code,report_no));
				   errCount++;
				} 
				if((checkRule(guarantee_amt_totacc       	,t_guarantee_amt_totacc       	,"累計保證金額",m_year,m_month,user_id,user_name,errCount,bank_code,report_no)).size() > 0){
				   updateDBDataList.add(checkRule(guarantee_amt_totacc       	,t_guarantee_amt_totacc       	,"累計保證金額",m_year,m_month,user_id,user_name,errCount,bank_code,report_no));
				   errCount++;
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
			System.out.println("upd_code="+upd_code);
			System.out.println("checkResult="+checkResult);
			System.out.println("Insert WML01 Data Start ...");
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
				errMsg = errMsg + "UpdateM05.doParserReport_M05 UpdateDB Error:"+DBManager.getErrMsg()+"<br>";
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
			//parserResult=false;
			errMsg = errMsg + "UpdateM05.doParserReport_M05 Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateM05.doParserReport_M05="+e.getMessage());			
		}
		return errMsg;
		//return parserResult;
	}

	//讀取上傳檔案的資料
	private static List getM05FileData(String report_no,String filename){
		String	txtline	 = null;			
		List M05List = new LinkedList();
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
					byte[] byte1 = txtline.getBytes("BIG5");
					detail=new LinkedList();
					if(byte1.length==153){	//長度為153
						detail.add("M05");
						detail.add(txtline.substring(  0,  1));	//計算期間及項目（貸款機構代碼）
						detail.add(txtline.substring(  1,  4));	//計算期間及項目（計算期間）
						detail.add(txtline.substring(  4,  6));	//計算期間及項目（計算項目）
						detail.add(txtline.substring(  6, 13));	// 3代償案件件數
						detail.add(txtline.substring( 13, 27));	// 4代償案件金額
						detail.add(txtline.substring( 27, 34));	// 5經營不善件數
						detail.add(txtline.substring( 34, 48));	// 6經營不善金額
						detail.add(txtline.substring( 48, 55));	// 7週轉失靈件數
						detail.add(txtline.substring( 55, 69));	// 8週轉失靈金額
						detail.add(txtline.substring( 69, 76));	// 9疾病或死亡件數
						detail.add(txtline.substring( 76, 90));	//10疾病或死亡金額
						detail.add(txtline.substring( 90, 97));	//11災歉件數
						detail.add(txtline.substring( 97,111));	//12災歉金額
						detail.add(txtline.substring(111,118));	//13兼業失利件數
						detail.add(txtline.substring(118,132));	//14兼業失利金額
						detail.add(txtline.substring(132,139));	//15其他件數
						detail.add(txtline.substring(139,153));	//16其他金額
					}else if(byte1.length==25){	//長度為 25
						detail.add("M05_TOTACC");
						detail.add(txtline.substring( 0, 1));		//貸款機構
						detail.add(txtline.substring( 1, 4));		//固定編碼
						detail.add(txtline.substring( 4,11));	//累計保證件數
						detail.add(txtline.substring(11,25));	//累計保證金額
					}else{	//長度為 11,18
						detail.add("M05_NOTE");
						detail.add(txtline.substring(  0,  4));				//附註代碼
						detail.add(txtline.substring(  4,  byte1.length));	//附註金額
					}
					M05List.add(detail);
				}
			}	
			in.close();
			f.close();
		}catch(Exception e){
			errMsg = errMsg + "UpdateM05.getM05FileData Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateM05.getM05FileData Error:"+e.toString());
		}
		
		return M05List;
	}
	
	//Insert "0" 至M05
	//99.11.18 add 套用DAO.preparestatment,並列印轉換後的SQL by 2295
	private static boolean InsertZeroM05(String m_year,String m_month,String bank_code){
		List paramList = new ArrayList();
		StringBuffer sqlCmd = new StringBuffer();	
		List updateDBList = new LinkedList();//0:sql 1:data		
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data
		boolean updateOK=false;
		
		try{
			//新增資料至M05			
			sqlCmd.append("select loan_unit_no,substr(data_range,1,3) as \"period_no\",substr(data_range,4,2) as \"item_no\" ");
			sqlCmd.append("from m00_loan_unit,m00_data_range_item ");
			sqlCmd.append("where report_no='M05' ");
			List data_div01 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),null,"");//M05 	    
    	
			sqlCmd.delete(0, sqlCmd.length());
			sqlCmd.append("INSERT INTO M05 (M_YEAR,M_MONTH,LOAN_UNIT_NO,period_no,item_no ) VALUES (?,?,?,?,?)");
		
			for(int d1=0;d1<data_div01.size();d1++){
				dataList = new LinkedList();//儲存參數的data
				dataList.add(String.valueOf(Integer.parseInt(m_year)));
				dataList.add(String.valueOf(Integer.parseInt(m_month)));
				dataList.add((String)((DataObject)data_div01.get(d1)).getValue("loan_unit_no"));				
				dataList.add((String)((DataObject)data_div01.get(d1)).getValue("period_no"));
				dataList.add((String)((DataObject)data_div01.get(d1)).getValue("item_no"));
				updateDBDataList.add(dataList);//1:傳內的參數List	
			}
			if(updateDBDataList != null && updateDBDataList.size()!=0){
			   updateDBSqlList.add(sqlCmd.toString());
			   updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
			   updateDBList.add(updateDBSqlList);
			}
			
			//新增資料至M05_TOTACC			
			sqlCmd.delete(0, sqlCmd.length());
			sqlCmd.append("select loan_unit_no ");
			sqlCmd.append("from m00_loan_unit ");
			sqlCmd.append("order by input_order ");
			data_div01 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),null,"");//M05 	    
    	
			updateDBDataList = new LinkedList();
			updateDBSqlList = new LinkedList();
			sqlCmd.delete(0, sqlCmd.length());
			sqlCmd.append("INSERT INTO M05_TOTACC (M_YEAR,M_MONTH,LOAN_UNIT_NO,fix_no ) VALUES (?,?,?,?)");
		
			for(int d1=0;d1<data_div01.size();d1++){
				dataList = new LinkedList();//儲存參數的data
				dataList.add(String.valueOf(Integer.parseInt(m_year)));
				dataList.add(String.valueOf(Integer.parseInt(m_month)));
				dataList.add((String)((DataObject)data_div01.get(d1)).getValue("loan_unit_no"));				
				dataList.add("0ET");
				updateDBDataList.add(dataList);//1:傳內的參數List
			}
			if(updateDBDataList != null && updateDBDataList.size()!=0){
			   updateDBSqlList.add(sqlCmd.toString());
			   updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
			   updateDBList.add(updateDBSqlList);
			}
				
			
			//新增資料至M05_NOTE		
			sqlCmd.delete(0, sqlCmd.length());
			sqlCmd.append("select data_range ");
			sqlCmd.append("from m00_data_range_item ");                         
			sqlCmd.append("where report_no='M05'  ");
			sqlCmd.append("and data_range_type='N' ");
			sqlCmd.append("order by input_order ");
			data_div01 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),null,"");//M05備註 	    
			
			updateDBDataList = new LinkedList();
			updateDBSqlList = new LinkedList();
			sqlCmd.delete(0, sqlCmd.length());
			sqlCmd.append("INSERT INTO M05_NOTE (M_YEAR,M_MONTH,NOTE_NO ) VALUES (?,?,?)");
		
			for(int d1=0;d1<data_div01.size();d1++){
				dataList = new LinkedList();//儲存參數的data
				dataList.add(String.valueOf(Integer.parseInt(m_year)));
				dataList.add(String.valueOf(Integer.parseInt(m_month)));
				dataList.add((String)((DataObject)data_div01.get(d1)).getValue("data_range"));				
				updateDBDataList.add(dataList);//1:傳內的參數List	
				
			}
			if(updateDBDataList != null && updateDBDataList.size()!=0){
			   updateDBSqlList.add(sqlCmd.toString());
			   updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
			   updateDBList.add(updateDBSqlList);
			}
				
			updateOK=DBManager.updateDB_ps(updateDBList);
			
			System.out.println("M05/M05_TOTACC/M05_NOTE Insert Zero OK??"+updateOK);
			if(!updateOK){
				errMsg = errMsg + "M05_NOTE Insert Zero Fail:"+DBManager.getErrMsg()+"<br>";
				System.out.println(DBManager.getErrMsg());
			}
		}catch(Exception e){
			errMsg = errMsg + "UpdateM05.InsertZeroM05 Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateM05.InsertZeroM05 Error:"+e.getMessage());
		}		
    	return updateOK;
	}
	
	//運算各項比較值
	private static List checkRule(long num1,long num2,String errmsg,  String m_year,String m_month,String user_id,String user_name,int errCount,String bank_code,String report_no){
		List dataList =  new ArrayList();//儲存參數的data	
		if(num1!=num2){
			dataList = Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,errmsg+"實際加總與合計列不合",user_id,user_name);			
			System.out.println("num1="+num1+"num2="+num2);			
		}		
		return dataList;
	}
}

