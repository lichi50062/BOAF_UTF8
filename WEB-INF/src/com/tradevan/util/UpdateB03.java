//99.11.16 add ALL SQL 套用DAO.preparestatment,並列印轉換後的SQL by 2295
//99.11.19 add 移至共用Utility.getWML03_count讀取WML03 count(*)
//Utility.getWML01讀取WML01 all data
//Utility.Insert_UpdateWML01當WML01不存在Insert,存在時Update
//寫入WML03的參數移至共用create_dataList by 2295
package com.tradevan.util;

import java.util.*;
import java.text.SimpleDateFormat;
import com.tradevan.util.dao.DataObject;

public class UpdateB03{
	private static String errMsg = "";
	
	public String getErrMsg(){
		return errMsg;  
	}
	
	public static synchronized String doParserReport_B03(String report_no, String m_year,
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
				
		List dbData = null;//其他querydb後,資料暫存的list
		String checkResult="true"; 
		
		//99.11.16 add 查詢年度100年以前.縣市別不同===============================
	    String cd01_table = (Integer.parseInt(m_year) < 100)?"cd01_99":"";
	    String wlx01_m_year = (Integer.parseInt(m_year) < 100)?"99":"100";
	    List updateDBList = new ArrayList();//0:sql 1:data
	    List updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data
		DataObject bean = null;
	    //=====================================================================
		
		try { 
			user_id = szuser_id;
			user_name = szuser_name;
			errCount = 0;
			
			dbData = Utility.getWML03_count(String.valueOf(Integer.parseInt(m_year)),String.valueOf(Integer.parseInt(m_month)),bank_code,report_no);//99.11.19 fix
			/*99.11.19移至Utility.getWML03_count讀取WML03.count(*)*/
			Utility.printLogTime("B03-1 time");	
			
			if(dbData.size() != 0){//WML03有資料時,先清掉
				sqlCmd.delete(0,sqlCmd.length());
				sqlCmd.append(" INSERT INTO WML03_LOG "); 
				sqlCmd.append(" select m_year,m_month,bank_code,report_no,serial_no,remark,user_id,user_name,update_date,?,?,sysdate,'D'");
				sqlCmd.append(" from WML03");
				sqlCmd.append(" where m_year=? AND m_month=? AND bank_code=? AND report_no=?");
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
				/*99.11.16
				sqlCmd = " INSERT INTO WML03_LOG " 
						   + " select m_year,m_month,bank_code,report_no,serial_no,remark,user_id,user_name,update_date"
						   + ",'"+user_id+"','"+user_name+"',sysdate,'D'"
						   + " from WML03"
						   + " where m_year=" + String.valueOf(Integer.parseInt(m_year)) + " AND m_month=" + String.valueOf(Integer.parseInt(m_month)) + " AND " 
						   + " bank_code='" + bank_code + "' AND report_no='" + report_no + "'";
					updateDBSqlList.add(sqlCmd);			
					sqlCmd = "DELETE FROM WML03 WHERE m_year=" + String.valueOf(Integer.parseInt(m_year)) + " AND m_month=" + String.valueOf(Integer.parseInt(m_month)) + " AND " +
						   "bank_code='" + bank_code + "' AND report_no='" + report_no + "'";
				*/
			}
			Utility.printLogTime("B03-2 time");	
						
			//Insert Error to WML03 (檢查合計和小計)
			updateDBDataList = new ArrayList();	
			if(errCount == 0){//無其他錯誤時,才檢核合計金額				
				String subErrMsg="";
				sqlCmd.delete(0,sqlCmd.length());
				paramList = new ArrayList();
				//B03_1 Check檢查農業發展基金貸款統計表之"農機"資料
				sqlCmd.append("select decode(funs_master_no||funs_sub_no||funs_next_no,'010100','0','1') as \"funs_no\", ");
				sqlCmd.append("sum(loan_cnt_totacc) as \"loan_cnt_totacc\", ");
                sqlCmd.append(" sum(loan_amt_totacc_fund) as \"loan_amt_totacc_fund\", ");
                sqlCmd.append(" sum(loan_amt_totacc_bank) as \"loan_amt_totacc_bank\", ");
                sqlCmd.append(" sum(loan_amt_totacc_tot) as \"loan_amt_totacc_tot\", ");
                sqlCmd.append(" sum(loan_cnt_bal) as \"loan_cnt_bal\", ");
                sqlCmd.append(" sum(loan_amt_bal_fund) as \"loan_amt_bal_fund\", ");
                sqlCmd.append(" sum(loan_amt_bal_bank) as \"loan_amt_bal_bank\", ");
                sqlCmd.append(" sum(loan_amt_bal_tot) as \"loan_amt_bal_tot\" ");
                sqlCmd.append(" from B03_1 ");
                sqlCmd.append("where m_year=? and m_month=?");
                sqlCmd.append("and funs_master_no||funs_sub_no||funs_next_no in ('010100','010101','010102') ");
                sqlCmd.append("group by decode(funs_master_no||funs_sub_no||funs_next_no,'010100','0','1') ");
                sqlCmd.append("order by decode(funs_master_no||funs_sub_no||funs_next_no,'010100','0','1') desc");
            	paramList.add(String.valueOf(Integer.parseInt(m_year)));
				paramList.add(String.valueOf(Integer.parseInt(m_month)));
                dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,loan_cnt_totacc, loan_amt_totacc_fund, loan_amt_totacc_bank, loan_amt_totacc_tot, loan_cnt_bal, loan_amt_bal_fund, loan_amt_bal_bank, loan_amt_bal_tot");
				long f_loan_cnt_totacc			= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_cnt_totacc")		==null?"0":((DataObject)dbData.get(0)).getValue("loan_cnt_totacc")).toString());
				long f_loan_amt_totacc_fund 	= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_amt_totacc_fund")	==null?"0":((DataObject)dbData.get(0)).getValue("loan_amt_totacc_fund")).toString());
				long f_loan_amt_totacc_bank		= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_amt_totacc_bank")	==null?"0":((DataObject)dbData.get(0)).getValue("loan_amt_totacc_bank")).toString());
				long f_loan_amt_totacc_tot		= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_amt_totacc_tot")	==null?"0":((DataObject)dbData.get(0)).getValue("loan_amt_totacc_tot")).toString());
				long f_loan_cnt_bal 			= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_cnt_bal")			==null?"0":((DataObject)dbData.get(0)).getValue("loan_cnt_bal")).toString());
				long f_loan_amt_bal_fund		= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_amt_bal_fund")		==null?"0":((DataObject)dbData.get(0)).getValue("loan_amt_bal_fund")).toString());
				long f_loan_amt_bal_bank		= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_amt_bal_bank")		==null?"0":((DataObject)dbData.get(0)).getValue("loan_amt_bal_bank")).toString());
				long f_loan_amt_bal_tot 		= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_amt_bal_tot")		==null?"0":((DataObject)dbData.get(0)).getValue("loan_amt_bal_tot")).toString());
				
				long ft_loan_cnt_totacc			= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_cnt_totacc")		==null?"0":((DataObject)dbData.get(1)).getValue("loan_cnt_totacc")).toString());
				long ft_loan_amt_totacc_fund 	= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_amt_totacc_fund")	==null?"0":((DataObject)dbData.get(1)).getValue("loan_amt_totacc_fund")).toString());
				long ft_loan_amt_totacc_bank	= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_amt_totacc_bank")	==null?"0":((DataObject)dbData.get(1)).getValue("loan_amt_totacc_bank")).toString());
				long ft_loan_amt_totacc_tot		= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_amt_totacc_tot")	==null?"0":((DataObject)dbData.get(1)).getValue("loan_amt_totacc_tot")).toString());
				long ft_loan_cnt_bal 			= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_cnt_bal")			==null?"0":((DataObject)dbData.get(1)).getValue("loan_cnt_bal")).toString());
				long ft_loan_amt_bal_fund		= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_amt_bal_fund")		==null?"0":((DataObject)dbData.get(1)).getValue("loan_amt_bal_fund")).toString());
				long ft_loan_amt_bal_bank		= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_amt_bal_bank")		==null?"0":((DataObject)dbData.get(1)).getValue("loan_amt_bal_bank")).toString());
				long ft_loan_amt_bal_tot 		= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_amt_bal_tot")		==null?"0":((DataObject)dbData.get(1)).getValue("loan_amt_bal_tot")).toString());

				if(f_loan_cnt_totacc != ft_loan_cnt_totacc	) { //比對筆數金額
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"貸放累計戶數與農機列不合",user_id,user_name));								
					errCount++ ;										
				}
				if(f_loan_amt_totacc_fund != ft_loan_amt_totacc_fund ) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"貸放累計基金加總與農機列不合",user_id,user_name));								
					errCount++ ;		
				}
				if(f_loan_amt_totacc_bank	!= ft_loan_amt_totacc_bank	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"貸放累計經辦機構加總與農機列不合",user_id,user_name));								
					errCount++ ;		
				}	
				if(f_loan_amt_totacc_tot	!= ft_loan_amt_totacc_tot	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"貸放累計合計加總與農機列不合",user_id,user_name));								
					errCount++ ;										
				}	
				if(f_loan_cnt_bal	!= ft_loan_cnt_bal	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"貸放餘額戶數加總與農機列不合",user_id,user_name));								
					errCount++ ;										
				}	
				if(f_loan_amt_bal_fund	!= ft_loan_amt_bal_fund	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"貸放餘額基金加總與農機列不合",user_id,user_name));								
					errCount++ ;										
				}	
				if(f_loan_amt_bal_bank	!= ft_loan_amt_bal_bank	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"貸放餘額經辦機構加總與農機列不合",user_id,user_name));								
					errCount++ ;					
				}	
				if(f_loan_amt_bal_tot	!= ft_loan_amt_bal_tot	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"貸放餘額合計加總與農機列不合",user_id,user_name));								
					errCount++ ;										
				}
				Utility.printLogTime("B03-3 time");	
				//B03_1 Check檢查農業發展基金貸款統計表之"加建"資料
				sqlCmd.delete(0,sqlCmd.length());	
				sqlCmd.append(" select decode(funs_master_no||funs_sub_no||funs_next_no,'010400','0','1') as \"funs_no\", ");
				sqlCmd.append(" sum(loan_cnt_totacc) as \"loan_cnt_totacc\", ");
                sqlCmd.append(" sum(loan_amt_totacc_fund) as \"loan_amt_totacc_fund\", ");
                sqlCmd.append(" sum(loan_amt_totacc_bank) as \"loan_amt_totacc_bank\", ");
                sqlCmd.append(" sum(loan_amt_totacc_tot) as \"loan_amt_totacc_tot\", ");
                sqlCmd.append(" sum(loan_cnt_bal) as \"loan_cnt_bal\", ");
                sqlCmd.append(" sum(loan_amt_bal_fund) as \"loan_amt_bal_fund\", ");
                sqlCmd.append(" sum(loan_amt_bal_bank) as \"loan_amt_bal_bank\", ");
                sqlCmd.append(" sum(loan_amt_bal_tot) as \"loan_amt_bal_tot\" ");
                sqlCmd.append(" from	B03_1 ");
                sqlCmd.append(" where m_year=? and m_month=?");
                sqlCmd.append(" and funs_master_no||funs_sub_no||funs_next_no in ('010400','010401','010402','010404','010405','010406','010407','010408') ");
                sqlCmd.append(" group by decode(funs_master_no||funs_sub_no||funs_next_no,'010400','0','1') ");
                sqlCmd.append(" order by decode(funs_master_no||funs_sub_no||funs_next_no,'010400','0','1') desc");
                
                dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,loan_cnt_totacc, loan_amt_totacc_fund, loan_amt_totacc_bank, loan_amt_totacc_tot, loan_cnt_bal, loan_amt_bal_fund, loan_amt_bal_bank, loan_amt_bal_tot");
				long b_loan_cnt_totacc			= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_cnt_totacc")		==null?"0":((DataObject)dbData.get(0)).getValue("loan_cnt_totacc")).toString());
				long b_loan_amt_totacc_fund 	= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_amt_totacc_fund")	==null?"0":((DataObject)dbData.get(0)).getValue("loan_amt_totacc_fund")).toString());
				long b_loan_amt_totacc_bank		= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_amt_totacc_bank")	==null?"0":((DataObject)dbData.get(0)).getValue("loan_amt_totacc_bank")).toString());
				long b_loan_amt_totacc_tot		= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_amt_totacc_tot")	==null?"0":((DataObject)dbData.get(0)).getValue("loan_amt_totacc_tot")).toString());
				long b_loan_cnt_bal 			= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_cnt_bal")			==null?"0":((DataObject)dbData.get(0)).getValue("loan_cnt_bal")).toString());
				long b_loan_amt_bal_fund		= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_amt_bal_fund")		==null?"0":((DataObject)dbData.get(0)).getValue("loan_amt_bal_fund")).toString());
				long b_loan_amt_bal_bank		= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_amt_bal_bank")		==null?"0":((DataObject)dbData.get(0)).getValue("loan_amt_bal_bank")).toString());
				long b_loan_amt_bal_tot 		= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_amt_bal_tot")		==null?"0":((DataObject)dbData.get(0)).getValue("loan_amt_bal_tot")).toString());
				
				long bt_loan_cnt_totacc			= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_cnt_totacc")		==null?"0":((DataObject)dbData.get(1)).getValue("loan_cnt_totacc")).toString());
				long bt_loan_amt_totacc_fund 	= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_amt_totacc_fund")	==null?"0":((DataObject)dbData.get(1)).getValue("loan_amt_totacc_fund")).toString());
				long bt_loan_amt_totacc_bank	= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_amt_totacc_bank")	==null?"0":((DataObject)dbData.get(1)).getValue("loan_amt_totacc_bank")).toString());
				long bt_loan_amt_totacc_tot		= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_amt_totacc_tot")	==null?"0":((DataObject)dbData.get(1)).getValue("loan_amt_totacc_tot")).toString());
				long bt_loan_cnt_bal 			= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_cnt_bal")			==null?"0":((DataObject)dbData.get(1)).getValue("loan_cnt_bal")).toString());
				long bt_loan_amt_bal_fund		= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_amt_bal_fund")		==null?"0":((DataObject)dbData.get(1)).getValue("loan_amt_bal_fund")).toString());
				long bt_loan_amt_bal_bank		= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_amt_bal_bank")		==null?"0":((DataObject)dbData.get(1)).getValue("loan_amt_bal_bank")).toString());
				long bt_loan_amt_bal_tot 		= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_amt_bal_tot")		==null?"0":((DataObject)dbData.get(1)).getValue("loan_amt_bal_tot")).toString());

				if(b_loan_cnt_totacc != bt_loan_cnt_totacc	) { //比對筆數金額
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"貸放累計戶數與加建列不合",user_id,user_name));								
					errCount++ ;
				}
				if(b_loan_amt_totacc_fund != bt_loan_amt_totacc_fund ) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"貸放累計基金加總與加建列不合",user_id,user_name));								
					errCount++ ;					
				}
				if(b_loan_amt_totacc_bank	!= bt_loan_amt_totacc_bank	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"貸放累計經辦機構加總與加建列不合",user_id,user_name));								
					errCount++ ;										
				}	
				if(b_loan_amt_totacc_tot	!= bt_loan_amt_totacc_tot	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"貸放累計合計加總與加建列不合",user_id,user_name));								
					errCount++ ;					
				}	
				if(b_loan_cnt_bal	!= bt_loan_cnt_bal	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"貸放餘額戶數加總與加建列不合",user_id,user_name));								
					errCount++ ;										
				}	
				if(b_loan_amt_bal_fund	!= bt_loan_amt_bal_fund	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"貸放餘額基金加總與加建列不合",user_id,user_name));								
					errCount++ ;					
				}	
				if(b_loan_amt_bal_bank	!= bt_loan_amt_bal_bank	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"貸放餘額經辦機構加總與加建列不合",user_id,user_name));								
					errCount++ ;										
				}	
				if(b_loan_amt_bal_tot	!= bt_loan_amt_bal_tot	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"貸放餘額合計加總與加建列不合",user_id,user_name));								
					errCount++ ;					
				}
				Utility.printLogTime("B03-4 time");	
				//B03_1-檢查農業發展基金貸款統計表之"合計"資料
				sqlCmd.delete(0,sqlCmd.length());	
				sqlCmd.append(" select decode(funs_master_no||funs_sub_no||funs_next_no,'019090','0','1') as \"funs_no\", ");
                sqlCmd.append(" sum(loan_cnt_totacc) as \"loan_cnt_totacc\", ");
                sqlCmd.append(" sum(loan_amt_totacc_fund) as \"loan_amt_totacc_fund\", ");
                sqlCmd.append(" sum(loan_amt_totacc_bank) as \"loan_amt_totacc_bank\", ");
                sqlCmd.append(" sum(loan_amt_totacc_tot) as \"loan_amt_totacc_tot\", ");
                sqlCmd.append(" sum(loan_cnt_bal) as \"loan_cnt_bal\", ");
                sqlCmd.append(" sum(loan_amt_bal_fund) as \"loan_amt_bal_fund\", ");
                sqlCmd.append(" sum(loan_amt_bal_bank) as \"loan_amt_bal_bank\", ");
                sqlCmd.append(" sum(loan_amt_bal_tot) as \"loan_amt_bal_tot\" ");
                sqlCmd.append(" from B03_1 ");
                sqlCmd.append(" where m_year=? and m_month=?");
                sqlCmd.append(" and funs_master_no||funs_sub_no||funs_next_no in ('010100','010200','010300','010400','019090') ");
                sqlCmd.append(" group by decode(funs_master_no||funs_sub_no||funs_next_no,'019090','0','1') ");
                sqlCmd.append(" order by decode(funs_master_no||funs_sub_no||funs_next_no,'019090','0','1') desc");
                dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,loan_cnt_totacc, loan_amt_totacc_fund, loan_amt_totacc_bank, loan_amt_totacc_tot, loan_cnt_bal, loan_amt_bal_fund, loan_amt_bal_bank, loan_amt_bal_tot");		
				long loan_cnt_totacc		= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_cnt_totacc")		==null?"0":((DataObject)dbData.get(0)).getValue("loan_cnt_totacc")).toString());
				long loan_amt_totacc_fund 	= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_amt_totacc_fund")	==null?"0":((DataObject)dbData.get(0)).getValue("loan_amt_totacc_fund")).toString());
				long loan_amt_totacc_bank	= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_amt_totacc_bank")	==null?"0":((DataObject)dbData.get(0)).getValue("loan_amt_totacc_bank")).toString());
				long loan_amt_totacc_tot	= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_amt_totacc_tot")	==null?"0":((DataObject)dbData.get(0)).getValue("loan_amt_totacc_tot")).toString());
				long loan_cnt_bal 			= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_cnt_bal")			==null?"0":((DataObject)dbData.get(0)).getValue("loan_cnt_bal")).toString());
				long loan_amt_bal_fund		= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_amt_bal_fund")		==null?"0":((DataObject)dbData.get(0)).getValue("loan_amt_bal_fund")).toString());
				long loan_amt_bal_bank		= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_amt_bal_bank")		==null?"0":((DataObject)dbData.get(0)).getValue("loan_amt_bal_bank")).toString());
				long loan_amt_bal_tot 		= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_amt_bal_tot")		==null?"0":((DataObject)dbData.get(0)).getValue("loan_amt_bal_tot")).toString());
				
				long t_loan_cnt_totacc		= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_cnt_totacc")		==null?"0":((DataObject)dbData.get(1)).getValue("loan_cnt_totacc")).toString());
				long t_loan_amt_totacc_fund = Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_amt_totacc_fund")	==null?"0":((DataObject)dbData.get(1)).getValue("loan_amt_totacc_fund")).toString());
				long t_loan_amt_totacc_bank	= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_amt_totacc_bank")	==null?"0":((DataObject)dbData.get(1)).getValue("loan_amt_totacc_bank")).toString());
				long t_loan_amt_totacc_tot	= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_amt_totacc_tot")	==null?"0":((DataObject)dbData.get(1)).getValue("loan_amt_totacc_tot")).toString());
				long t_loan_cnt_bal 		= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_cnt_bal")			==null?"0":((DataObject)dbData.get(1)).getValue("loan_cnt_bal")).toString());
				long t_loan_amt_bal_fund	= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_amt_bal_fund")		==null?"0":((DataObject)dbData.get(1)).getValue("loan_amt_bal_fund")).toString());
				long t_loan_amt_bal_bank	= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_amt_bal_bank")		==null?"0":((DataObject)dbData.get(1)).getValue("loan_amt_bal_bank")).toString());
				long t_loan_amt_bal_tot 	= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_amt_bal_tot")		==null?"0":((DataObject)dbData.get(1)).getValue("loan_amt_bal_tot")).toString());

				if(loan_cnt_totacc != t_loan_cnt_totacc	) { //比對筆數金額
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"貸放累計戶數與合計列不合",user_id,user_name));								
					errCount++ ;					
				}
				if(loan_amt_totacc_fund != t_loan_amt_totacc_fund ) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"貸放累計基金加總與合計列不合",user_id,user_name));								
					errCount++ ;										
				}
				if(loan_amt_totacc_bank	!= t_loan_amt_totacc_bank	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"貸放累計經辦機構加總與合計列不合",user_id,user_name));								
					errCount++ ;										
				}	
				if(loan_amt_totacc_tot	!= t_loan_amt_totacc_tot	) {
				    updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"貸放累計合計加總與合計列不合",user_id,user_name));								
					errCount++ ;					
				}	
				if(loan_cnt_bal	!= t_loan_cnt_bal	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"貸放餘額戶數加總與合計列不合",user_id,user_name));								
					errCount++ ;					
				}	
				if(loan_amt_bal_fund	!= t_loan_amt_bal_fund	) {
				    updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"貸放餘額基金加總與合計列不合",user_id,user_name));								
					errCount++ ;										
				}	
				if(loan_amt_bal_bank	!= t_loan_amt_bal_bank	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"貸放餘額經辦機構加總與合計列不合",user_id,user_name));								
					errCount++ ;					
				}	
				if(loan_amt_bal_tot	!= t_loan_amt_bal_tot	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"貸放餘額合計加總與合計列不合",user_id,user_name));								
					errCount++ ;										
				}
				Utility.printLogTime("B03-5 time");	
				//B03_2 Check
				sqlCmd.delete(0,sqlCmd.length());	
				sqlCmd.append(" select decode(funs_master_no||funs_sub_no||funs_next_no,'019090','0','1') as \"funs_no\", ");
                sqlCmd.append("      sum(loan_amt_bal) as \"loan_cnt_totacc\", ");
                sqlCmd.append("      sum(loan_amt_over) as \"loan_amt_totacc_fund\", ");
                sqlCmd.append("      sum(loan_rate_over) as \"loan_amt_totacc_bank\" ");
                sqlCmd.append(" from	B03_2 ");
                sqlCmd.append(" where m_year=? and m_month=?");
                sqlCmd.append(" group by decode(funs_master_no||funs_sub_no||funs_next_no,'019090','0','1') ");
                sqlCmd.append(" order by decode(funs_master_no||funs_sub_no||funs_next_no,'019090','0','1') desc");
                dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,loan_amt_bal,loan_amt_over,loan_rate_over");		
				long loan_amt_bal		= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_amt_bal")		==null?"0":((DataObject)dbData.get(0)).getValue("loan_amt_bal")).toString());
				long loan_amt_over 		= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_amt_over")		==null?"0":((DataObject)dbData.get(0)).getValue("loan_amt_over")).toString());
				long loan_rate_over		= Long.parseLong((((DataObject)dbData.get(0)).getValue("loan_rate_over")	==null?"0":((DataObject)dbData.get(0)).getValue("loan_rate_over")).toString());
				
				long t_loan_amt_bal		= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_amt_bal")		==null?"0":((DataObject)dbData.get(1)).getValue("loan_amt_bal")).toString());
				long t_loan_amt_over 	= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_amt_over")		==null?"0":((DataObject)dbData.get(1)).getValue("loan_amt_over")).toString());
				long t_loan_rate_over	= Long.parseLong((((DataObject)dbData.get(1)).getValue("loan_rate_over")	==null?"0":((DataObject)dbData.get(1)).getValue("loan_rate_over")).toString());
				
				if(loan_amt_bal != t_loan_amt_bal	) { //比對筆數金額
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"貸款餘額與合計列不合",user_id,user_name));								
					errCount++ ;
				}
				if(loan_amt_over != t_loan_amt_over ) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"逾期餘額加總與合計列不合",user_id,user_name));								
					errCount++ ;
				}
				if(loan_rate_over	!= t_loan_rate_over	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"逾放比率加總與合計列不合",user_id,user_name));								
					errCount++ ;										
				}	
				Utility.printLogTime("B03-6 time");	
				//B03_3 Check 基金來源總額
				sqlCmd.delete(0,sqlCmd.length());	
				sqlCmd.append(" select decode(funo_master_no||funo_sub_no||funo_next_no,'010100','0','1') as \"funo_no\", ");
				sqlCmd.append(" sum(FUNO_AMT) as \"FUNO_AMT\", ");
				sqlCmd.append(" sum(FUNO_RATE) as \"FUNO_RATE\" ");
				sqlCmd.append(" from b03_3 ");
				sqlCmd.append(" where m_year=? and m_month=?");
				sqlCmd.append(" and funo_master_no||funo_sub_no||funo_next_no in ('010100','010102','010103') ");
				sqlCmd.append(" group by decode(funo_master_no||funo_sub_no||funo_next_no,'010100','0','1') ");
				sqlCmd.append(" order by decode(funo_master_no||funo_sub_no||funo_next_no,'010100','0','1') desc");
				dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,funo_amt, funo_rate");
				long funo_amt	= Long.parseLong((((DataObject)dbData.get(0)).getValue("funo_amt")	==null?"0":((DataObject)dbData.get(0)).getValue("funo_amt")).toString());
				long funo_rate 	= Long.parseLong((((DataObject)dbData.get(0)).getValue("funo_rate")	==null?"0":((DataObject)dbData.get(0)).getValue("funo_rate")).toString());
				long t_funo_amt		= Long.parseLong((((DataObject)dbData.get(1)).getValue("funo_amt")	==null?"0":((DataObject)dbData.get(1)).getValue("funo_amt")).toString());
				long t_funo_rate 	= Long.parseLong((((DataObject)dbData.get(1)).getValue("funo_rate")	==null?"0":((DataObject)dbData.get(1)).getValue("funo_rate")).toString());
				if(funo_amt != t_funo_amt	) { //比對筆數金額
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"金額加總與基金來源總額列不合",user_id,user_name));								
					errCount++ ;		
				}
				if(funo_rate != t_funo_rate ) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"比率加總與基金來源總額列不合",user_id,user_name));								
					errCount++ ;					
				}
				Utility.printLogTime("B03-7 time");	
				//B03_3 Check 基金運用總額
				sqlCmd.delete(0,sqlCmd.length());	
				sqlCmd.append(" select decode(funo_master_no||funo_sub_no||funo_next_no,'010200','0','1') as \"funo_no\", ");
				sqlCmd.append(" sum(FUNO_AMT) as \"FUNO_AMT\", ");
				sqlCmd.append(" sum(FUNO_RATE) as \"FUNO_RATE\" ");
				sqlCmd.append(" from b03_3 ");
				sqlCmd.append(" where m_year=? and m_month=?");
				sqlCmd.append(" and funo_master_no||funo_sub_no||funo_next_no not in ('010000','010100','010102','010103') ");
				sqlCmd.append(" group by decode(funo_master_no||funo_sub_no||funo_next_no,'010200','0','1') ");
				sqlCmd.append(" order by decode(funo_master_no||funo_sub_no||funo_next_no,'010200','0','1') desc");
				dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,funo_amt, funo_rate");
				long b_funo_amt		= Long.parseLong((((DataObject)dbData.get(0)).getValue("funo_amt")	==null?"0":((DataObject)dbData.get(0)).getValue("funo_amt")).toString());
				long b_funo_rate 	= Long.parseLong((((DataObject)dbData.get(0)).getValue("funo_rate")	==null?"0":((DataObject)dbData.get(0)).getValue("funo_rate")).toString());
				long bt_funo_amt	= Long.parseLong((((DataObject)dbData.get(1)).getValue("funo_amt")	==null?"0":((DataObject)dbData.get(1)).getValue("funo_amt")).toString());
				long bt_funo_rate 	= Long.parseLong((((DataObject)dbData.get(1)).getValue("funo_rate")	==null?"0":((DataObject)dbData.get(1)).getValue("funo_rate")).toString());
				if(b_funo_amt != bt_funo_amt	) { //比對筆數金額
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"金額加總與基金運用總額列不合",user_id,user_name));								
					errCount++ ;					
				}
				if(b_funo_rate != bt_funo_rate ) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"比率加總與基金運用總額列不合",user_id,user_name));								
					errCount++ ;										
				}
				Utility.printLogTime("B03-8 time");	
				//B03_4 Check
				sqlCmd.delete(0,sqlCmd.length());
				sqlCmd.append(" select decode(bank_no,'90','0','1') as \"bank_no\", ");
                sqlCmd.append("      sum(machine_cnt) as \"machine_cnt\", ");
                sqlCmd.append("      sum(machine_amt) as \"machine_amt\", ");
                sqlCmd.append("      sum(land_cnt) as \"land_cnt\", ");
                sqlCmd.append("      sum(land_amt) as \"land_amt\", ");
                sqlCmd.append("      sum(house_cnt) as \"house_cnt\", ");
                sqlCmd.append("      sum(house_amt) as \"house_amt\", ");
                sqlCmd.append("      sum(build_cnt) as \"build_cnt\", ");
                sqlCmd.append("      sum(build_amt) as \"build_amt\", ");
                sqlCmd.append("      sum(tot_cnt) as \"tot_cnt\", ");
                sqlCmd.append("      sum(tot_amt) as \"tot_amt\" ");
                sqlCmd.append(" from	B03_4 ");
                sqlCmd.append(" where m_year=? and m_month=?");
                sqlCmd.append(" group by decode(bank_no,'90','0','1') ");
                sqlCmd.append(" order by decode(bank_no,'90','0','1') desc");
              
                dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,machine_cnt, machine_amt, land_cnt, land_amt, house_cnt, house_amt, build_cnt, build_amt, tot_cnt, tot_amt");		
				long machine_cnt	= Long.parseLong((((DataObject)dbData.get(0)).getValue("machine_cnt")	==null?"0":((DataObject)dbData.get(0)).getValue("machine_cnt")).toString());
				long machine_amt 	= Long.parseLong((((DataObject)dbData.get(0)).getValue("machine_amt")	==null?"0":((DataObject)dbData.get(0)).getValue("machine_amt")).toString());
				long land_cnt		= Long.parseLong((((DataObject)dbData.get(0)).getValue("land_cnt")		==null?"0":((DataObject)dbData.get(0)).getValue("land_cnt")).toString());
				long land_amt		= Long.parseLong((((DataObject)dbData.get(0)).getValue("land_amt")		==null?"0":((DataObject)dbData.get(0)).getValue("land_amt")).toString());
				long house_cnt 		= Long.parseLong((((DataObject)dbData.get(0)).getValue("house_cnt")		==null?"0":((DataObject)dbData.get(0)).getValue("house_cnt")).toString());
				long house_amt		= Long.parseLong((((DataObject)dbData.get(0)).getValue("house_amt")		==null?"0":((DataObject)dbData.get(0)).getValue("house_amt")).toString());
				long build_cnt		= Long.parseLong((((DataObject)dbData.get(0)).getValue("build_cnt")		==null?"0":((DataObject)dbData.get(0)).getValue("build_cnt")).toString());
				long build_amt 		= Long.parseLong((((DataObject)dbData.get(0)).getValue("build_amt")		==null?"0":((DataObject)dbData.get(0)).getValue("build_amt")).toString());
				long tot_cnt		= Long.parseLong((((DataObject)dbData.get(0)).getValue("tot_cnt")		==null?"0":((DataObject)dbData.get(0)).getValue("tot_cnt")).toString());
				long tot_amt 		= Long.parseLong((((DataObject)dbData.get(0)).getValue("tot_amt")		==null?"0":((DataObject)dbData.get(0)).getValue("tot_amt")).toString());
				
				long t_machine_cnt	= Long.parseLong((((DataObject)dbData.get(1)).getValue("machine_cnt")	==null?"0":((DataObject)dbData.get(1)).getValue("machine_cnt")).toString());
				long t_machine_amt 	= Long.parseLong((((DataObject)dbData.get(1)).getValue("machine_amt")	==null?"0":((DataObject)dbData.get(1)).getValue("machine_amt")).toString());
				long t_land_cnt		= Long.parseLong((((DataObject)dbData.get(1)).getValue("land_cnt")		==null?"0":((DataObject)dbData.get(1)).getValue("land_cnt")).toString());
				long t_land_amt		= Long.parseLong((((DataObject)dbData.get(1)).getValue("land_amt")		==null?"0":((DataObject)dbData.get(1)).getValue("land_amt")).toString());
				long t_house_cnt 	= Long.parseLong((((DataObject)dbData.get(1)).getValue("house_cnt")		==null?"0":((DataObject)dbData.get(1)).getValue("house_cnt")).toString());
				long t_house_amt	= Long.parseLong((((DataObject)dbData.get(1)).getValue("house_amt")		==null?"0":((DataObject)dbData.get(1)).getValue("house_amt")).toString());
				long t_build_cnt	= Long.parseLong((((DataObject)dbData.get(1)).getValue("build_cnt")		==null?"0":((DataObject)dbData.get(1)).getValue("build_cnt")).toString());
				long t_build_amt 	= Long.parseLong((((DataObject)dbData.get(1)).getValue("build_amt")		==null?"0":((DataObject)dbData.get(1)).getValue("build_amt")).toString());
				long t_tot_cnt		= Long.parseLong((((DataObject)dbData.get(1)).getValue("tot_cnt")		==null?"0":((DataObject)dbData.get(1)).getValue("tot_cnt")).toString());
				long t_tot_amt 		= Long.parseLong((((DataObject)dbData.get(1)).getValue("tot_amt")		==null?"0":((DataObject)dbData.get(1)).getValue("tot_amt")).toString());
				
				if(machine_cnt != t_machine_cnt	) { //比對筆數金額
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"農機戶數與合計列不合",user_id,user_name));								
					errCount++ ;		
				}
				if(machine_amt != t_machine_amt ) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"農機金額加總與合計列不合",user_id,user_name));								
					errCount++ ;										
				}
				if(land_cnt	!= t_land_cnt	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"購地戶數加總與合計列不合",user_id,user_name));								
					errCount++ ;										
				}	
				if(land_amt	!= t_land_amt	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"購地金額加總與合計列不合",user_id,user_name));								
					errCount++ ;		
				}	
				if(house_cnt	!= t_house_cnt	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"農宅戶數加總與合計列不合",user_id,user_name));								
					errCount++ ;					
				}	
				if(house_amt	!= t_house_amt	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"農宅金額加總與合計列不合",user_id,user_name));								
					errCount++ ;
				}	
				if(build_cnt	!= t_build_cnt	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"加建戶數加總與合計列不合",user_id,user_name));								
					errCount++ ;					
				}	
				if(build_amt	!= t_build_amt	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"加建金額加總與合計列不合",user_id,user_name));								
					errCount++ ;					
				}
				if(tot_cnt	!= t_tot_cnt	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"合計戶數加總與合計列不合",user_id,user_name));								
					errCount++ ;								
				}
				if(tot_amt	!= t_tot_amt	) {
					updateDBDataList.add(Utility.create_dataList(m_year,m_month,bank_code,report_no,errCount,"合計金額加總與合計列不合",user_id,user_name));								
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
			
			Utility.printLogTime("B03-9 time");	
			dbData =  Utility.getWML01(String.valueOf(Integer.parseInt(m_year)),String.valueOf(Integer.parseInt(m_month)),bank_code,report_no);
			/*99.11.19移至Utility.getWML01讀取WML01 all data*/				
			upd_code = (errCount == 0)?"U":"E";//U檢核成功:E檢核失敗
			checkResult = (errCount == 0)?"true":"false";//true檢核成功:false檢核失敗
			
			System.out.println("dbData.size="+dbData.size());
			
			//99.11.19 fix
			List getUpdateDBList = Utility.Insert_UpdateWML01(dbData,String.valueOf(Integer.parseInt(m_year)), String.valueOf(Integer.parseInt(m_month)),bank_code,report_no,input_method,add_date,user_id,user_name,common_center,upd_method,upd_code,batch_no);
			for(int i=0;i<getUpdateDBList.size();i++){
			    updateDBList.add(getUpdateDBList.get(i));
			}
			/*99.11.19移至Utility.Insert_UpdateWML01;wml01不存在Insert,存在時Update*/
			Utility.printLogTime("B03-11 time");	
			
			//傳送email至bank_cmml的e-mail信箱
			System.out.println("send email begin");		
			//線上編輯
			Utility.sendMailNotification(bank_code,report_no,
			m_year,m_month, checkResult,input_method,Utility.getDateFormat("yyyy/MM/dd HH:mm:ss"),"",user_id,"");
			System.out.println("send email end");
			
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
				errMsg = errMsg + "UpdateB03.doParserReport_B03 UpdateDB Error:"+DBManager.getErrMsg()+"<br>";
				System.out.println(DBManager.getErrMsg());
			}
			Utility.printLogTime("B03-12 time");		
		}catch (Exception e) {
			//parserResult=false;
			errMsg = errMsg + "UpdateB03.doParserReport_B03 Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateB03.doParserReport_B03="+e.getMessage());			
		}
		
		return errMsg;
		//return parserResult;
	}
}


