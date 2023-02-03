//94.04.22 fix 都為"0"值時,顯示檢核為"0" by 2295
//94.06.22 fix 當M03List長度為11時,才檢核是否都為"0"
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

public class UpdateM03 {
	private static String errMsg = "";
	public String getErrMsg(){
		return errMsg;  
	}
	public static synchronized String doParserReport_M03(String report_no, String m_year,
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
		String  bank_code=srcbank_code;	
		String  add_date="";//申報日期
		String  upd_code="";//檢核結果
		String  common_center="";//由共用中心傳入
		boolean updateOK=false;
		SimpleDateFormat bkformat = new SimpleDateFormat("yyyyMMddHHmmss");
		SimpleDateFormat emailformat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");	
		String WMdataDir = Utility.getProperties("WMdataDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");
		String WMdataBKDir = Utility.getProperties("WMdataBKDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");		
		
		List M03List = new LinkedList();//M03細部資料
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
				M03List = getM03FileData(report_no,filename);
				user_data = Utility.getUser_Data(filename);//取得該檔案的異動者帳號.姓名
				user_id=user_data[0];
				user_name=user_data[1];
				System.out.println("user_id="+user_id);
				System.out.println("user_name="+user_name);
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
			
			//將資料 INSERT 至M03
			if(input_method.equals("F")){//檔案上傳
				sqlCmd.delete(0,sqlCmd.length());
				paramList = new ArrayList();			
				sqlCmd.append("select * from "+report_no+" where m_year=? AND m_month=?") ;
				paramList.add(String.valueOf(Integer.parseInt(m_year)));
				paramList.add(String.valueOf(Integer.parseInt(m_month)));
				
				dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month");
			    if(dbData.size() == 0){//M03無資料時,才insert一筆Zero的資料
			    	System.out.println("InsertZeroM03");
			    	InsertZeroM03(m_year,m_month,bank_code);
			    }else{//end of M03不存在時			    
			    	updateDBDataList = new ArrayList();	
					updateDBSqlList = new ArrayList();
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" INSERT INTO M03_LOG ");
					sqlCmd.append(" select m_year,m_month,div_no,guarantee_cnt_month,loan_amt_month,guarantee_amt_month,");
					sqlCmd.append("        guarantee_cnt_year,loan_amt_year,guarantee_amt_year,guarantee_bal_totacc,");
					sqlCmd.append("        guarantee_bal_totacc_over,repay_bal_totacc,?,?,sysdate,'U'");
					sqlCmd.append(" from M03");
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
					sqlCmd.append(" INSERT INTO M03_NOTE_LOG ");
					sqlCmd.append(" select m_year,m_month,note_no,note_amt_rate,?,?,sysdate,'U'");
					sqlCmd.append("   from M03_NOTE");
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
			    List updateDBDataList_M03 = new ArrayList();
			    List updateDBDataList_M03Note = new ArrayList();	
				
			    for(int j=0;j<M03List.size();j++){//把上傳檔案裡的資料更新至M03				
			    	String tmp_reportNo=(String)((List)M03List.get(j)).get(0);
			    	//94.04.24 add 檢核上傳檔案內容的金額是否都為"0"
			    	//94.06.22 fix 當M03List長度為11時,才檢核是否都為"0"			    	
			    	if(((List)M03List.get(j)).size() == 11){
			    		if(  ((!((String)((List)M03List.get(j)).get(2)).equals("")) && Long.parseLong((String)((List)M03List.get(j)).get(2)) > 0 )
				    	  || ((!((String)((List)M03List.get(j)).get(3)).equals("")) && Long.parseLong((String)((List)M03List.get(j)).get(3)) > 0 )
				    	  || ((!((String)((List)M03List.get(j)).get(4)).equals("")) && Long.parseLong((String)((List)M03List.get(j)).get(4)) > 0 )
						  || ((!((String)((List)M03List.get(j)).get(5)).equals("")) && Long.parseLong((String)((List)M03List.get(j)).get(5)) > 0 )
						   || ((!((String)((List)M03List.get(j)).get(6)).equals("")) && Long.parseLong((String)((List)M03List.get(j)).get(6)) > 0 )
						   || ((!((String)((List)M03List.get(j)).get(7)).equals("")) && Long.parseLong((String)((List)M03List.get(j)).get(7)) > 0 )
						   || ((!((String)((List)M03List.get(j)).get(8)).equals("")) && Long.parseLong((String)((List)M03List.get(j)).get(8)) > 0 )
				    	   || ((!((String)((List)M03List.get(j)).get(9)).equals("")) && Long.parseLong((String)((List)M03List.get(j)).get(9)) > 0 )
				    	   || ((!((String)((List)M03List.get(j)).get(10)).equals("")) && Long.parseLong((String)((List)M03List.get(j)).get(10)) > 0 ))
				    	    {
				    	         nonZero = true;
				    	    }//end of > 0		    		  
			    	}//end of size=11	 
			    	if(nonZero){
			    	  if(tmp_reportNo.equals("M03")){//M03
			    	  	dataList =  new ArrayList();
					  	dataList.add((String)((List)M03List.get(j)).get( 2));
					  	dataList.add((String)((List)M03List.get(j)).get( 3));
					  	dataList.add((String)((List)M03List.get(j)).get( 4));
					  	dataList.add((String)((List)M03List.get(j)).get( 5));
					  	dataList.add((String)((List)M03List.get(j)).get( 6));
					  	dataList.add((String)((List)M03List.get(j)).get( 7));
					  	dataList.add((String)((List)M03List.get(j)).get( 8));
					  	dataList.add((String)((List)M03List.get(j)).get( 9));
					  	dataList.add((String)((List)M03List.get(j)).get(10));						
					  	dataList.add(String.valueOf(Integer.parseInt(m_year)));
					  	dataList.add(String.valueOf(Integer.parseInt(m_month)));
					  	dataList.add((String)((List)M03List.get(j)).get(1));						
					  	updateDBDataList_M03.add(dataList);
				      }else{//M03_NOTE
				      	dataList =  new ArrayList();
					  	dataList.add((String)((List)M03List.get(j)).get( 2));									
					  	dataList.add(String.valueOf(Integer.parseInt(m_year)));
					  	dataList.add(String.valueOf(Integer.parseInt(m_month)));
					  	dataList.add((String)((List)M03List.get(j)).get(1));						
					  	updateDBDataList_M03Note.add(dataList);					
					  }			       
			    	}//end of nonzero
			    }//end of M03List
			    if(updateDBDataList_M03 != null && updateDBDataList_M03.size()!=0){	
			    	updateDBSqlList = new ArrayList();
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("UPDATE M03 SET guarantee_cnt_month =?,");
					sqlCmd.append("   loan_amt_month      =?,");
					sqlCmd.append("   guarantee_amt_month =?,");
					sqlCmd.append("   guarantee_cnt_year  =?,");
					sqlCmd.append("   loan_amt_year       	=?,");
					sqlCmd.append("   guarantee_amt_year   	=?,");
					sqlCmd.append("   guarantee_bal_totacc  =?,");
					sqlCmd.append("   guarantee_bal_totacc_over	=?,");
					sqlCmd.append("   repay_bal_totacc  =?");
				    sqlCmd.append(" WHERE m_year=? AND m_month=?"); 
				    sqlCmd.append("   AND div_no=?");
		  
			 		updateDBSqlList.add(sqlCmd.toString());
					updateDBSqlList.add(updateDBDataList_M03);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
				}
			    if(updateDBDataList_M03Note != null && updateDBDataList_M03Note.size()!=0){	
			    	updateDBSqlList = new ArrayList();
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("UPDATE M03_NOTE SET note_amt_rate =?");
					sqlCmd.append(" WHERE m_year=? AND m_month=?"); 
					sqlCmd.append("   AND note_no=?");					
			 		updateDBSqlList.add(sqlCmd.toString());
					updateDBSqlList.add(updateDBDataList_M03Note);//0:sql 1:參數List
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
			}//end of if 檔案上傳	
			if(input_method.equals("W")){//94.04.22 線上編輯check是否值全部為"0"
		    	nonZero = false;
		    	sqlCmd.delete(0,sqlCmd.length());
				paramList = new ArrayList();
				sqlCmd.append("select * from M03 where m_year=? AND m_month=?" );
				paramList.add(String.valueOf(Integer.parseInt(m_year)));
				paramList.add(String.valueOf(Integer.parseInt(m_month)));
				dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"guarantee_cnt_month,loan_amt_month,guarantee_amt_month,guarantee_cnt_year,loan_amt_year,guarantee_amt_year,guarantee_bal_totacc,guarantee_bal_totacc_over,repay_bal_totacc");
		    	   
		    	checkZeroLoop:
			    if(dbData != null && dbData.size() != 0){
			    	for(int zeroIdx=0;zeroIdx < dbData.size();zeroIdx++){			    		  
			    		if(	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("guarantee_cnt_month")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("loan_amt_month")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("guarantee_amt_month")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("guarantee_cnt_year")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("loan_amt_year")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("guarantee_amt_year")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("guarantee_bal_totacc")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("guarantee_bal_totacc_over")).toString()) > 0
						||	Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("repay_bal_totacc")).toString()) > 0												
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
				sqlCmd.append("select guarantee_cnt_month, ");
				sqlCmd.append("       loan_amt_month, ");
				sqlCmd.append("       guarantee_amt_month, ");
				sqlCmd.append("       guarantee_cnt_year, ");
				sqlCmd.append("       loan_amt_year, ");
				sqlCmd.append("       guarantee_amt_year, ");
				sqlCmd.append("       guarantee_bal_totacc, ");
				sqlCmd.append("       guarantee_bal_totacc_over, ");
				sqlCmd.append("       repay_bal_totacc ");
				sqlCmd.append("  from m03,m00_data_range_item ");
				sqlCmd.append(" where m03.div_no=data_range ");
				sqlCmd.append("   and report_no='M03' and data_range_type='C' ");
				sqlCmd.append("   and m_year=? AND m_month=?"); 
				sqlCmd.append(" order by input_order ");
				paramList.add(String.valueOf(Integer.parseInt(m_year)));
	            paramList.add(String.valueOf(Integer.parseInt(m_month)));
	            dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,guarantee_cnt_month,loan_amt_month,guarantee_amt_month,guarantee_cnt_year,loan_amt_year,guarantee_amt_year,guarantee_bal_totacc,guarantee_bal_totacc_over,repay_bal_totacc");
	           		
                
				long[] guarantee_cnt_month = {0,0,0,0,0,0,0};
				long[] loan_amt_month = {0,0,0,0,0,0,0};
				long[] guarantee_amt_month = {0,0,0,0,0,0,0};
				long[] guarantee_cnt_year = {0,0,0,0,0,0,0};
				long[] loan_amt_year = {0,0,0,0,0,0,0};
				long[] guarantee_amt_year = {0,0,0,0,0,0,0};
				long[] guarantee_bal_totacc = {0,0,0,0,0,0,0};
				long[] guarantee_bal_totacc_over = {0,0,0,0,0,0,0};
				long[] repay_bal_totacc = {0,0,0,0,0,0,0};
				
				updateDBDataList = new ArrayList();	//99.11.18
				for(int i=0;i<dbData.size();i++){
					guarantee_cnt_month[i]			= Long.parseLong((((DataObject)dbData.get(i)).getValue("guarantee_cnt_month")		==null?"0":((DataObject)dbData.get(i)).getValue("guarantee_cnt_month")).toString());
					loan_amt_month[i]				= Long.parseLong((((DataObject)dbData.get(i)).getValue("loan_amt_month")			==null?"0":((DataObject)dbData.get(i)).getValue("loan_amt_month")).toString());
					guarantee_amt_month[i] 			= Long.parseLong((((DataObject)dbData.get(i)).getValue("guarantee_amt_month")		==null?"0":((DataObject)dbData.get(i)).getValue("guarantee_amt_month")).toString());
					guarantee_cnt_year[i]			= Long.parseLong((((DataObject)dbData.get(i)).getValue("guarantee_cnt_year")		==null?"0":((DataObject)dbData.get(i)).getValue("guarantee_cnt_year")).toString());
					loan_amt_year[i]  				= Long.parseLong((((DataObject)dbData.get(i)).getValue("loan_amt_year")				==null?"0":((DataObject)dbData.get(i)).getValue("loan_amt_year")).toString());
					guarantee_amt_year[i]  			= Long.parseLong((((DataObject)dbData.get(i)).getValue("guarantee_amt_year")		==null?"0":((DataObject)dbData.get(i)).getValue("guarantee_amt_year")).toString());
					guarantee_bal_totacc[i]			= Long.parseLong((((DataObject)dbData.get(i)).getValue("guarantee_bal_totacc")		==null?"0":((DataObject)dbData.get(i)).getValue("guarantee_bal_totacc")).toString());
					guarantee_bal_totacc_over[i]	= Long.parseLong((((DataObject)dbData.get(i)).getValue("guarantee_bal_totacc_over")	==null?"0":((DataObject)dbData.get(i)).getValue("guarantee_bal_totacc_over")).toString());
					repay_bal_totacc[i]				= Long.parseLong((((DataObject)dbData.get(i)).getValue("repay_bal_totacc")			==null?"0":((DataObject)dbData.get(i)).getValue("repay_bal_totacc")).toString());
				}
				if((checkRule(guarantee_cnt_month,"當月份保證件數",m_year,m_month,user_id,user_name,errCount,bank_code,report_no)).size() > 0){
				   updateDBDataList.add(checkRule(guarantee_cnt_month,"當月份保證件數",m_year,m_month,user_id,user_name,errCount,bank_code,report_no));
				   errCount++;
				}   
				if((checkRule(loan_amt_month,"當月份貸款金額",m_year,m_month,user_id,user_name,errCount,bank_code,report_no)).size() > 0){
				   updateDBDataList.add(checkRule(loan_amt_month,"當月份貸款金額",m_year,m_month,user_id,user_name,errCount,bank_code,report_no));
				   errCount++;
				} 
				if((checkRule(guarantee_amt_month,"當月份保證金額",m_year,m_month,user_id,user_name,errCount,bank_code,report_no)).size() > 0){
				   updateDBDataList.add(checkRule(guarantee_amt_month,"當月份保證金額",m_year,m_month,user_id,user_name,errCount,bank_code,report_no));
				   errCount++;
				}
				if((checkRule(guarantee_cnt_year,"當年度保證件數",m_year,m_month,user_id,user_name,errCount,bank_code,report_no)).size() > 0){
				   updateDBDataList.add(checkRule(guarantee_cnt_year,"當年度保證件數",m_year,m_month,user_id,user_name,errCount,bank_code,report_no));
				   errCount++;
				}
				if((checkRule(loan_amt_year,"當年度貸款金額",m_year,m_month,user_id,user_name,errCount,bank_code,report_no)).size() > 0){
				   updateDBDataList.add(checkRule(loan_amt_year,"當年度貸款金額",m_year,m_month,user_id,user_name,errCount,bank_code,report_no));
				   errCount++;
				}
				if((checkRule(guarantee_amt_year,"當年度保證金額",m_year,m_month,user_id,user_name,errCount,bank_code,report_no)).size() > 0){
				   updateDBDataList.add(checkRule(guarantee_amt_year,"當年度保證金額",m_year,m_month,user_id,user_name,errCount,bank_code,report_no));
				   errCount++;
				} 
				if((checkRule(guarantee_bal_totacc,"總累計保證金額",m_year,m_month,user_id,user_name,errCount,bank_code,report_no)).size() > 0){
				   updateDBDataList.add(checkRule(guarantee_bal_totacc,"總累計保證金額",m_year,m_month,user_id,user_name,errCount,bank_code,report_no));
				   errCount++;
				}
				if((checkRule(guarantee_bal_totacc_over,"總累計逾期保證餘額",m_year,m_month,user_id,user_name,errCount,bank_code,report_no)).size() > 0){
				   updateDBDataList.add(checkRule(guarantee_bal_totacc_over,"總累計逾期保證餘額",m_year,m_month,user_id,user_name,errCount,bank_code,report_no));
				   errCount++;
				}
				if((checkRule(repay_bal_totacc,"總累計代位清償淨額",m_year,m_month,user_id,user_name,errCount,bank_code,report_no)).size() > 0){
				   updateDBDataList.add(checkRule(repay_bal_totacc,"總累計代位清償淨額",m_year,m_month,user_id,user_name,errCount,bank_code,report_no));
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
			System.out.println("dbData.size="+dbData.size());
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
				errMsg = errMsg + "UpdateM03.doParserReport_M03 UpdateDB Error:"+DBManager.getErrMsg()+"<br>";
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
			errMsg = errMsg + "UpdateM03.doParserReport_M03 Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateM03.doParserReport_M03="+e.getMessage());			
		}
		return errMsg;		
	}

	//讀取上傳檔案的資料
	private static List getM03FileData(String report_no,String filename){
			String	txtline	 = null;			
			List M03List = new LinkedList();
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
							if(byte1.length==114){	//長度為114
								detail.add("M03");
								detail.add(txtline.substring(  0,  2));	//區分              
								detail.add(txtline.substring(  2,  9));	//當月份保證件數    
								detail.add(txtline.substring(  9, 23));	//當月份貸款金額    
								detail.add(txtline.substring( 23, 37));	//當月份保證金額    
								detail.add(txtline.substring( 37, 44));	//當年度保證件數    
								detail.add(txtline.substring( 44, 58));	//當年度貸款金額    
								detail.add(txtline.substring( 58, 72));	//當年度保證金額    
								detail.add(txtline.substring( 72, 86));	//總累計保證餘額    
								detail.add(txtline.substring( 86,100));	//總累計逾期保證餘額
								detail.add(txtline.substring(100,114));	//總累計代位清償淨額
							}else{
								detail.add("M03_NOTE");
								detail.add(txtline.substring(  0,  4));				//附註代碼
								detail.add(txtline.substring(  4,  byte1.length));	//附註金額
							}
							M03List.add(detail);
						}
				}	
				in.close();
				f.close();
			}catch(Exception e){
				errMsg = errMsg + "UpdateM03.getM03FileData Error:"+e.getMessage()+"<br>";
				System.out.println("UpdateM03.getM03FileData Error:"+e.toString());
			}
			return M03List;
	}
	
	//Insert "0" 至M03
	//99.11.17 add 套用DAO.preparestatment,並列印轉換後的SQL by 2295
	private static boolean InsertZeroM03(String m_year,String m_month,String bank_code){
		List paramList = new ArrayList();
		StringBuffer sqlCmd = new StringBuffer();	
		List updateDBList = new LinkedList();//0:sql 1:data		
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data
		boolean updateOK=false;		
		try{
			//新增資料至M03
			sqlCmd.append(" select data_range ");
			sqlCmd.append(" from m00_data_range_item ");                         
			sqlCmd.append(" where report_no='M03'  ");
			sqlCmd.append(" and data_range_type='C' ");
			sqlCmd.append(" order by input_order ");
			List data_div01 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),null,"");//M03 	    
    	
			sqlCmd.delete(0, sqlCmd.length());
			sqlCmd.append("INSERT INTO M03 (M_YEAR,M_MONTH,div_no ) VALUES (?,?,?)");
		
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
			//新增資料至M03_NOTE
			sqlCmd.delete(0, sqlCmd.length());
			sqlCmd.append("select data_range ");
			sqlCmd.append("from m00_data_range_item ");                         
	        sqlCmd.append("where report_no='M03'  ");
			sqlCmd.append("and data_range_type='S' ");
			sqlCmd.append("order by input_order ");
			data_div01 = DBManager.QueryDB_SQLParam(sqlCmd.toString(),null,"");//M03備註 	    
			
			updateDBDataList = new LinkedList();
			updateDBSqlList = new LinkedList();
			sqlCmd.delete(0, sqlCmd.length());
			sqlCmd.append("INSERT INTO M03_NOTE (M_YEAR,M_MONTH,NOTE_NO ) VALUES (?,?,?)");
		
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
			
			System.out.println("M03_NOTE Insert Zero OK??"+updateOK);
			if(!updateOK){
				errMsg = errMsg + "M03_NOTE Insert Zero Fail:"+DBManager.getErrMsg()+"<br>";
				System.out.println(DBManager.getErrMsg());
			}
		}catch(Exception e){
			errMsg = errMsg + "UpdateM03.InsertZeroM03 Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateM03.InsertZeroM03 Error:"+e.getMessage());
		}
    	return updateOK;
	}
	
	//運算各項比較值
	private static List checkRule(long[] longg,String checkItem,String m_year,String m_month,String user_id,String user_name,int errCount,String bank_code,String report_no){
		List dataList =  new ArrayList();//儲存參數的data		
		System.out.println("longg[0]="+longg[0]);
		System.out.println("longg[1]="+longg[1]);
		System.out.println("longg[2]="+longg[2]);
		System.out.println("longg[3]="+longg[3]);
		System.out.println("longg[4]="+longg[4]);		
		if(longg[3]!=(longg[0]-longg[1])){
			dataList = Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,checkItem+"上下月份比較數與實際不合",user_id,user_name);			
		}else if((longg[1]==0 && longg[4]!=100) || (longg[1]!=0 && (longg[4]!=(Math.round((double)(float)longg[3]/(float)longg[1]*10000)*10)))){	//2005.2.16 by 2354改成只驗證到小數位第2位
			float tmp=Math.round((double)(float)longg[3]/(float)longg[1]*10000)*10;//longg[3]/longg[1];
			System.out.println("longg[3]/longg[1]="+tmp);
			if(longg[1]!=0) System.out.println("(Math.round((double)(longg[3]/longg[1]*10000))/100)="+(Math.round((double)(longg[3]/longg[1]*10000))/100));
			dataList = Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,checkItem+"上下月份比較%與實際不合",user_id,user_name);			
		}
		if(longg[5]!=(longg[0]-longg[2])){
			dataList = Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,checkItem+"上下年度比較數與實際不合",user_id,user_name);			
		}else if((longg[2]==0 && longg[6]!=100) || (longg[2]!=0 && (longg[6]!=(Math.round((double)(float)longg[5]/(float)longg[2]*10000)*10)))){	//2005.2.16 by 2354改成只驗證到小數位第2位			
			dataList = Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,checkItem+"上下年度比較%與實際不合",user_id,user_name);			
		}		
		return dataList;
	}	
}
