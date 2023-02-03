
//94.03.02 金額改用BigDecimal去轉 by 2295
//94.03.03若項目代號最後一碼為N時.取中文名稱18-48 by 2295
//        若不為N時,取金額18-32 
//94.03.07 fix 有負號("-")時,把之前的"0"去掉 by 2295
//94.03.18 fix 910500/91060P由程式計算 by 2295
//94.03.21 fix 910400無法檢核的問題 by 2295
//94.03.22 add 若被鎖定,則不執行檢核,直接send mail通知 by 2295
//94.03.22 fix 改善performance by 2295
//94.03.22 fix 科目代號不區分農/漁會
//94.03.28 fix 若A04的值,都為"0"值時,顯示檢核為"0" by 2295
//94.05.11 fix AccCode最後一碼為"N"儲存中文字的問題 by 2295
//94.07.11 fix 910400不為空值時才計算91060P by 2295
//94.07.13 fix 9207XP不為空值時,才做累加 by 2295
//		   fix 920500--->/1000變/100000 by 2295
//94.07.19 add 檢核910201+910202 < 910199 by 2295
//94.08.19 add 檢核910199 < 0 and 910299 != 0 ==>第一類資本合計(910199)為負數,第二類資本(910299)應為0元  by 2295
//94.09.19 fix 910199 < 0 && 910299 = 0
//95.03.20 add 共用中心傳送彙總檢核結果e-mail by 2295
//95.03.27 fix 共用中心傳送彙總e-mail格式 by 2295
//95.05.11 add 910500=920101+920201+920301+920401+920501+920601+920710+920720+920730+920740+920750+920801-910203
//             減910203(2_1)特定損失所提列之備抵呆帳、損失準備及營業準備	 by 2295
//95.05.12 add 檔案上傳.原910500 - 910203
//			   若原A05無該科目時.Insert 一筆有科目時,才Update A05並Insert到A05_LOG by 2295
//95.05.16 add 檔案上傳時先增加額外科目代號 by 2295
//95.05.29 fix 910500*1.25%->限額取到整數,小數點下以無條件捨去 by 2295
//95.05.30 add 910109 > 0時,910202-910203也要等於0 by 2295
//95.08.17 fix 910202-910203 可小於 0 ,add 910202-910203 < 0 ,則910202-910203等於0 by 2295
//96.04.19 若為檔案上傳批次寫入A05 ZeroData / 有修改的A05寫入A05_log by 2295
//96.07.05 (未上線只加在備註)add 910204 調整後備抵呆帳、損失準備及營業準備,修正檢核公式 by 2295
//96.08.15 寫入暫存table AXX_TMP/有異動的資料寫入A03.使用preparestatement by 2295
//		    批次寫入WML01_Log/WML02_Log/WML03_Log/WML03 by 2295
//96.11.21 fix 批次寫入WML03_LOG(檢核其他錯誤)(區分檔案上傳/線上編輯) by 2295
//96.11.23 fix 將910500/91060P寫入AXX_TMP暫存table by 2295
//96.12.20 刪除AXX_TMP暫存檔 by 2295
//99.11.15 add InsertWML03/ALL SQL 套用DAO.preparestatment,並列印轉換後的SQL by 2295
//99.11.19 add 移至共用Utility.getWML03_count讀取WML03 count(*)
//					  Utility.getCountZero讀取該申報資料all data 都為0的資料筆數
//					  Utility.getWML01讀取WML01 all data
//					  Utility.Insert_UpdateWML01當WML01不存在Insert,存在時Update
//					  Utility.deleteWML01_UPLOAD刪除上傳檔案紀錄 by 2295
//103.04.14 add 910500調整為四捨五入取至整數 by 2295 
//106.02.20 fix 910203加總數值
//106.01.10 fix 910202與A01的跨表檢核公式.區分農會為112600/漁會為112300 by 2295
//107.06.21 fix 農金局-施秀芬來電(910109 > 0時,910202-910203也要等於0)此項檢核已於103年更改規定.移除此檢核條件 by 2295
//108.01.04 fix 910500調整再減910109.910401.910402.910403.910404
//          fix 910204=min([910202-910203],910500*1.75%) by 2295
//108.01.28 add A05.910500不扣除910109/910401/910402/910403/910404 by 2295
//          fix 910299=if(910199<0,0,min([910201+910204],910199)) by 2295
//110.03.09 add A05.920901/920902;920601風險權數原為0.5改為0.45 by 2295
package com.tradevan.util;

import java.io.*;
import java.util.*;
import java.text.SimpleDateFormat;
import com.tradevan.util.dao.DataObject;
import java.lang.Long;
import java.math.BigDecimal;


public class UpdateA05 {
	private static String errMsg = "";
	public String getErrMsg(){
		return errMsg;  
	}
	public static synchronized String doParserReport_A05(String report_no, String m_year,
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
		String  bank_code="";//機構代碼
		String  add_date="";//申報日期
		String  upd_code="";//檢核結果
		String  common_center="";//由共用中心傳入
		boolean updateOK=false;
		SimpleDateFormat bkformat = new SimpleDateFormat("yyyyMMddHHmmss");
		SimpleDateFormat emailformat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");	
		String WMdataDir = Utility.getProperties("WMdataDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");
		String WMdataBKDir = Utility.getProperties("WMdataBKDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");
		List A05List = new LinkedList();//A05細部資料
		List bank_codeList = new LinkedList();//機構代碼list
		List ruleList = new LinkedList();//check rule後,欲Insert到WML02的sqlList		
		List dbData = null;//其他querydb後,資料暫存的list
		String checkResult="true"; 		
		boolean nonZero=false;//94.03.28 檢查A05的內容是否都為"0"
		File tmpFile = new File(WMdataDir+filename);		
		Date filedate = new Date(tmpFile.lastModified());
		//99.11.15 add 查詢年度100年以前.縣市別不同===============================
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
			Utility.printLogTime("A05 begin time");	
			
			if(input_method.equals("F")){//若為檔案上傳,先將檔案讀取出來
				List allList = getA05FileData(report_no,filename,m_year,m_month);//讀取檔案,並寫入AXX_TMP暫存table									
				bank_codeList = (List)allList.get(0);
				A05List = (List)allList.get(1);
				user_data = Utility.getUser_Data(filename);//取得該檔案的異動者帳號.姓名
				user_id=user_data[0];
				user_name=user_data[1];
				System.out.println("userid="+user_id);
				System.out.println("username="+user_name);
			}else{//線上編輯,將bank_code加到bank_codeList				
				bank_codeList.add(srcbank_code);
				user_id = szuser_id;
				user_name = szuser_name;
			}
			Utility.printLogTime("A05-1 time");	
			
			//若為"Y"表示由共用中心代傳入	
			sqlCmd.append("select bank_type from ba01 where bank_no=? and m_year=?");
			paramList.add(srcbank_code);
			paramList.add(wlx01_m_year);
			dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");	
			
			if(dbData.size() != 0){//ba01有此筆bank_no的資料
			   System.out.println("common_center dbdata.size="+dbData.size());	
			   if((((DataObject)dbData.get(0)).getValue("bank_type") != null) && ((String)((DataObject)dbData.get(0)).getValue("bank_type")).equals("8")){			   	 
			   	  common_center="Y";			
			   	  Utility.setSendMsg("");//95.03.20清空共用中心的mail msg
			   	  Utility.addMailNotification("",report_no,m_year,m_month, checkResult,"",emailformat.format(filedate),filename,user_id);//95.03.27 add 彙總title
			   }
			}
			System.out.println("common_center="+common_center);
			Utility.printLogTime("A05-2 time");
			//批次寫入WML03_LOG(檢核其他錯誤)(區分檔案上傳/線上編輯)
			Utility.InsertWML03_LOG(m_year,m_month,user_id,user_name,report_no,filename,input_method,bank_codeList);
			//96.08.15 若為檔案上傳批次寫入A05 ZeroData / 有修改的A05寫入A03_log
			if(input_method.equals("F")){//在A03中無資料的先將insert Zero data 到A03
				errMsg += Utility.InsertZeroAXX_List(m_year,m_month,bank_codeList,report_no," acc_tr_type = 'A05' and acc_div='07' ");//在A05中無資料的先將insert Zero data 到A05
		    	InsertZeroExtraA05_List(m_year,m_month,bank_codeList);
				Utility.InsertAXX_LOG(m_year,m_month,user_id,user_name,filename,report_no);//批次寫入(有異動)至A05_LOG
		    	errMsg += Utility.InsertWML03(m_year,m_month,user_id,user_name,report_no,filename);//農漁會.批次寫入檢核其他錯誤
		    	errMsg += Utility.updateAXX(m_year,m_month,report_no,filename);//將有異動的資料update A03		    	
			}//end of 檔案上傳
			//批次寫入WML02_LOG(檢核公式錯誤)(區分檔案上傳/線上編輯)
			Utility.InsertWML02_LOG(m_year,m_month,user_id,user_name,report_no,input_method,filename,bank_codeList);
			//批次寫入WML01_LOG(檢核結果)(區分檔案上傳/線上編輯)
			Utility.InsertWML01_LOG(m_year,m_month,user_id,user_name,report_no,input_method,filename,bank_codeList);
						
			List lockBank_CodeList = Utility.getLockBank_Code(m_year,m_month,report_no);
			
			bank_codeLoop:
			for(int i=0;i<bank_codeList.size();i++){
				//94.03.22 add 若被鎖定,則不執行檢核,直接send mail通知==================================
				if(lockBank_CodeList != null){
				   for(int lockidx=0;lockidx<lockBank_CodeList.size();lockidx++){					
				       if(((String)bank_codeList.get(i)).equals((String)lockBank_CodeList.get(lockidx))){
				          if(common_center.equals("Y")){//由共用中心傳入的,先add欲傳送mail的內容
					         Utility.addMailNotification((String)bank_codeList.get(i),report_no,
								     m_year,m_month, checkResult,"C",emailformat.format(filedate),filename,user_id);
				          }else{
				             Utility.sendMailNotification((String)bank_codeList.get(i),report_no,
									 m_year,m_month, checkResult,"C",emailformat.format(filedate),filename,user_id,"");
				          }
				           continue bank_codeLoop;	
				       }
				   }
				}
				//=================================================================================
				Utility.printLogTime("A05-3 time");
				
				System.out.println("bank_codeList["+i+"]="+(String)bank_codeList.get(i));
				errCount = 0;
				//94.03.22 找出該bank_code的bank_type by 2295=============================================
				sqlCmd.delete(0,sqlCmd.length());
				sqlCmd.append("select bank_type from ba01 where bank_no=? and m_year=?");
				paramList = new ArrayList();
				paramList.add((String)bank_codeList.get(i));
				paramList.add(wlx01_m_year);
				dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");		
				
				if(dbData != null && dbData.size() != 0){
					bank_type = (String)((DataObject)dbData.get(0)).getValue("bank_type");
					System.out.println("bank_type="+bank_type);
				}
				//=================================================================================
				//clear WML03(檢核其他錯誤)
				/*96.08.15移至InsertWML03*/
				Utility.printLogTime("A05-4 time");
				
				//clear WML02(檢核公式)
				dbData = Utility.getWML03_count(m_year,m_month,(String)bank_codeList.get(i),report_no);//99.11.19 fix
				/*99.11.19移至Utility.getWML03_count讀取WML03.count(*)				
				sqlCmd.delete(0,sqlCmd.length());
				sqlCmd.append("select count(*) as countdata from WML03 where m_year=? AND m_month=? AND bank_code=? AND report_no=?");
				paramList = new ArrayList();
				paramList.add(m_year);
				paramList.add(m_month);
				paramList.add((String)bank_codeList.get(i));
				paramList.add(report_no);
				dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"countdata");
				*/	
			    if(Integer.parseInt((((DataObject)dbData.get(0)).getValue("countdata")).toString()) == 0){//96.08.15檢查WML03有無其他錯誤,無其他錯誤時,才檢核公式
				    System.out.println("bank_code="+(String)bank_codeList.get(i)+".WML03.size()="+(((DataObject)dbData.get(0)).getValue("countdata")).toString());
				    /*96.08.15 移至InsertWML02_LOG*/
				    //96.04.18移至InsertZeroA05_List/InsertZeroExtraA05_List
				    //96.04.19 移至InsertA05_LOG
				    Utility.printLogTime("A05-5 time");
				    
				    nonZero = false;
				    dbData = Utility.getCountZero(m_year,m_month,(String)bank_codeList.get(i),report_no,"amt > 0");//99.11.19 fix
				    /*99.11.19移至Utility.getCountZero讀取該申報資料all data 都為0的資料筆數					
					sqlCmd.delete(0,sqlCmd.length());
					paramList = new ArrayList();
				    
					//99.11.15檢核內容是否都為0
					sqlCmd.append(" select count(*) as countzero from "+report_no+" where m_year=? AND m_month=? AND bank_code=? and amt > 0");                          
					
					paramList.add(m_year);
					paramList.add(m_month);
					paramList.add((String)bank_codeList.get(i));
					
					dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"countzero");
				    */
					if(dbData != null && dbData.size()==1){
					   System.out.println("nonZero.size()="+(((DataObject)dbData.get(0)).getValue("countzero")).toString());
					   if(Integer.parseInt((((DataObject)dbData.get(0)).getValue("countzero")).toString())>0){
					       nonZero = true;
					   }
					}
					
				    Utility.printLogTime("A05-6 time");
				    	
				    if(nonZero){//94.03.28 fix 若A05的值,只要有非"0"值時,才執行檢核
				       //執行檢核
				       ruleList = CheckRule(bank_type,report_no,m_year,m_month,(String)bank_codeList.get(i),user_id,user_name);
					   if(ruleList.size() > 0){//99.11.15有檢核失敗時,才加入
					       updateDBList.add((List)ruleList.get(0));
					       errCount += ((List)((List)ruleList.get(0)).get(1)).size();//99.09.27累計參數List的size
					   }
				       //99.11.15ruleList = CheckRule(bank_type,report_no,m_year,m_month,(String)bank_codeList.get(i),user_id,user_name);
				       //for(int ruleidx=0;ruleidx<ruleList.size();ruleidx++){
				       //	 updateDBSqlList.add((String)ruleList.get(ruleidx));
				       //	 errCount ++;
				       //}
				       Utility.printLogTime("A05-7 time");
				       //94.07.19 add 檢核910201+910202<910199==================================================
				       ruleList = new LinkedList();
				       ruleList = CheckOtherRule(errCount,bank_type,report_no,m_year,m_month,(String)bank_codeList.get(i),user_id,user_name);
					   if(ruleList.size() > 0){//99.11.15有檢核失敗時,才加入
					       updateDBList.add((List)ruleList.get(0));
					       errCount += ((List)((List)ruleList.get(0)).get(1)).size();//99.09.27累計參數List的size
					   }
				       //99.11.15ruleList = CheckOtherRule(errCount,bank_type,report_no,m_year,m_month,(String)bank_codeList.get(i),user_id,user_name);
				       //for(int ruleidx=0;ruleidx<ruleList.size();ruleidx++){
				       //	   updateDBSqlList.add((String)ruleList.get(ruleidx));
				       //	   errCount ++;
				       //}
				       Utility.printLogTime("A05-8 time");				       	
				       //====================================================================================
				       /*執行A01.A05跨表檢核
				       ruleList = null;
				       ruleList = CheckRuleA01_A05(errCount,bank_type,report_no,m_year,m_month,(String)bank_codeList.get(i),user_id,user_name);
				       for(int ruleidx=0;ruleidx<ruleList.size();ruleidx++){
				    	   updateDBSqlList.add((String)ruleList.get(ruleidx));
				    	   errCount ++;
				       }*/
				    }				
				}else{//enf of errCount==0
				   errCount = Integer.parseInt((((DataObject)dbData.get(0)).getValue("countdata")).toString());	
				}//enf of errCount==0
			    dbData =  Utility.getWML01(m_year,m_month,(String)bank_codeList.get(i),report_no);
				/*99.11.19移至Utility.getWML01讀取WML01 all data		
			    sqlCmd.delete(0,sqlCmd.length());
				paramList = new ArrayList();
				sqlCmd.append("select * from WML01 where m_year=? AND m_month=? AND bank_code=? AND report_no=?");
			 	paramList.add(m_year);
				paramList.add(m_month);
				paramList.add((String)bank_codeList.get(i));
				paramList.add(report_no);
				dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,add_date,batch_no,update_date");
				*/		    
				upd_code = (errCount == 0)?"U":"E";//U檢核成功:E檢核失敗				
				checkResult = (errCount == 0)?"true":"false";//true檢核成功:false檢核失敗
				//94.03.28 fix Z檢核為0====================================
				if(!nonZero && errCount == 0){
					upd_code = "Z";//Z檢核為0
					checkResult="Z";//Z檢核為0
				}
				//========================================================
				List getUpdateDBList = Utility.Insert_UpdateWML01(dbData,m_year, m_month,(String)bank_codeList.get(i),report_no,input_method,add_date,user_id,user_name,common_center,upd_method,upd_code,batch_no);//99.11.19 fix
				for(int j=0;j<getUpdateDBList.size();j++){
				    updateDBList.add(getUpdateDBList.get(j));
				}
				/*99.11.19移至Utility.Insert_UpdateWML01;wml01不存在Insert,存在時Update
				sqlCmd.delete(0,sqlCmd.length());
				updateDBDataList = new ArrayList();
				updateDBSqlList = new ArrayList();
				if(dbData.size() == 0){//WML01不存在時,Insert
					if(input_method.equals("F")){//檔案上傳
					   sqlCmd.append("INSERT INTO WML01 VALUES (?,?,?,?,?,?,?,to_date('"+add_date+"','YYYYMMDDHH24MISS'),?,?,?,?,null,?,?,sysdate)");
						
					   //99.11.15sqlCmd="INSERT INTO WML01 VALUES (" +m_year+","+m_month+",'"+(String)bank_codeList.get(i)+"','"+report_no+"','"+input_method+"','"+user_id+"','"+user_name+"',to_date('"+add_date+"','YYYYMMDDHH24MISS'),'"+common_center+"','"
					   //      + upd_method +"','"+upd_code+"',"+batch_no+",null,'"+user_id+"','"+user_name+"',sysdate)";
					}else{//線上編輯
					   sqlCmd.append("INSERT INTO WML01 VALUES (?,?,?,?,?,?,?,sysdate,?,?,?,?,null,?,?,sysdate)");
					   //99.11.15sqlCmd="INSERT INTO WML01 VALUES (" +m_year+","+m_month+",'"+(String)bank_codeList.get(i)+"','"+report_no+"','"+input_method+"','"+user_id+"','"+user_name+"',sysdate,'"+common_center+"','"
				       //      + upd_method +"','"+upd_code+"',"+batch_no+",null,'"+user_id+"','"+user_name+"',sysdate)";	
					}
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql		
					dataList = new ArrayList();//傳內的參數List
					dataList.add(m_year); 
					dataList.add(m_month); 
					dataList.add((String)bank_codeList.get(i)); 
					dataList.add(report_no); 
					dataList.add(input_method);
					dataList.add(user_id);
					dataList.add(user_name);
					dataList.add(common_center);
					dataList.add(upd_method);
					dataList.add(upd_code);
					dataList.add(String.valueOf(batch_no));
					dataList.add(user_id);
					dataList.add(user_name);
					updateDBDataList.add(dataList);//1:傳內的參數List
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
				}else{//WML01存在時,做update
					//96.08.15移至批次寫入InsertWML01_LOG
				    //batch_no = Integer.parseInt((((DataObject)dbData.get(0)).getValue("batch_no")).toString()) + 1;
					sqlCmd.delete(0,sqlCmd.length());
					updateDBDataList = new ArrayList();
					updateDBSqlList = new ArrayList();
					sqlCmd.append(" UPDATE WML01 SET "); 
					sqlCmd.append(" upd_method=?,upd_code=?,batch_no=?,user_id=?,user_name=?,update_date=sysdate");   
					sqlCmd.append(" where m_year=? AND m_month=? AND bank_code=? AND report_no=?");
				
					//99.11.15sqlCmd=" UPDATE WML01 SET " 
					//	  +" upd_method='"+upd_method+"',upd_code='"+upd_code+"',batch_no="+batch_no+",user_id='"+user_id+"',user_name='"+user_name+"',update_date=sysdate"   
					//	  +" where m_year=" + m_year + " AND m_month=" + m_month + " AND " 						
					//	  +" bank_code='" + (String)bank_codeList.get(i) + "' AND report_no='" + report_no + "'";
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql		
					dataList = new ArrayList();//傳內的參數List
					dataList.add(upd_method); 
					dataList.add(upd_code); 
					dataList.add(String.valueOf(batch_no));
					dataList.add(user_id);
					dataList.add(user_name);
					dataList.add(m_year);
					dataList.add(m_month);
					dataList.add((String)bank_codeList.get(i));
					dataList.add(report_no);					
					updateDBDataList.add(dataList);//1:傳內的參數List					
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
				}
				*/
				Utility.printLogTime("A05-9 time");
				//傳送email至bank_cmml的e-mail信箱
			    System.out.println("send email begin");
				if(input_method.equals("F")){//檔案上傳
				    if(common_center.equals("Y")){//由共用中心傳入的,先add欲傳送mail的內容
				       Utility.addMailNotification((String)bank_codeList.get(i),report_no,
					           m_year,m_month, checkResult,input_method,emailformat.format(filedate),filename,user_id);
				    }else{
				        Utility.sendMailNotification((String)bank_codeList.get(i),report_no,
								m_year,m_month, checkResult,input_method,emailformat.format(filedate),filename,user_id,"");    
				    }
				}else{//線上編輯  		
				    if(common_center.equals("Y")){//由共用中心傳入的,先add欲傳送mail的內容
					   Utility.addMailNotification((String)bank_codeList.get(i),report_no,
					           m_year,m_month, checkResult,input_method,Utility.getDateFormat("yyyy/MM/dd HH:mm:ss"),"",user_id);
				    }else{
				        Utility.sendMailNotification((String)bank_codeList.get(i),report_no,
								m_year,m_month, checkResult,input_method,Utility.getDateFormat("yyyy/MM/dd HH:mm:ss"),"",user_id,"");    
				    }
				}		
				System.out.println("send email end");
			}//end of for -->bank_codeList
			Utility.printLogTime("A05-10 time");
				
			if(input_method.equals("F")){//檔案上傳
			    //將WML01_UPLOAD..filename對應的使用者帳號.姓名刪除			
			    //99.11.15updateDBSqlList.add("DELETE FROM WML01_UPLOAD where filename='"+filename+"'");
				//99.11.19刪除上傳檔案紀錄
				updateDBList.add(Utility.deleteWML01_UPLOAD(filename));//99.11.19				
			    //96.12.20 刪除AXX_TMP暫存檔
			    Utility.deleteAXX_TMP(m_year,m_month,report_no,filename);			   
			}
			if(updateDBList != null && updateDBList.size()!=0){
				updateOK=DBManager.updateDB_ps(updateDBList);
			}
			System.out.println("update OK??"+updateOK);
			if(!updateOK){
				//parserResult=false;
				errMsg = errMsg + "UpdateA05.doParserReport_A05 UpdateDB Error:"+DBManager.getErrMsg()+"<br>";
				System.out.println(DBManager.getErrMsg());
			}
			
			if(input_method.equals("F")){//檔案上傳
			   String CopyResult = Utility.CopyFile(WMdataDir+System.getProperty("file.separator")+filename,WMdataBKDir+System.getProperty("file.separator")+ filename);
       		   if(CopyResult.equals("0")){//copy成功時,才將檔案刪除,避免使用rename造成的錯誤
          	   		tmpFile = new File(WMdataDir+System.getProperty("file.separator")+filename);
          	   		if(tmpFile.exists()) tmpFile.delete();              		
       		   }	   				
			}	
			if(common_center.equals("Y")){//95.03.20由共用中心傳入的,傳送彙總mail
			   if(input_method.equals("F")){//檔案上傳			
			      Utility.sendMailNotification(srcbank_code,report_no,
						  m_year,m_month, checkResult,input_method,emailformat.format(filedate),filename,user_id,Utility.getSendMsg());
			   }else{//線上編輯
			      Utility.sendMailNotification(srcbank_code,report_no,
						  m_year,m_month, checkResult,input_method,Utility.getDateFormat("yyyy/MM/dd HH:mm:ss"),"",user_id,Utility.getSendMsg());    
			   }
			}
			Utility.printLogTime("A05-11 time");
			
		}catch (Exception e) {
			//parserResult=false;
			errMsg = errMsg + "UpdateA05.doParserReport_A05 Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateA05.doParserReport_A05="+e.getMessage());
		}
		return errMsg;
		//return parserResult;
	}
	//讀取上傳檔案的資料
	//99.11.15 add 套用DAO.preparestatment,並列印轉換後的SQL 
	private static List getA05FileData(String report_no,String filename,String m_year,String m_month){
			String	txtline	 = null;			
			List A05List = new LinkedList();
			List bank_codeList = new LinkedList();
			List allList = new LinkedList();
			List detail = null;
			List dbData = null;			
			String tmpAcc_Code="";
			String tmpAmt="";
			//96.08.15 寫入AXX_TMP暫存table,使用preparestatement
			List paramList = new ArrayList();
			StringBuffer sqlCmd = new StringBuffer();	
			List updateDBList = new LinkedList();//0:sql 1:data		
			List updateDBSqlList = new LinkedList();
			List updateDBDataList = new LinkedList();//儲存參數的List
			List dataList = new LinkedList();//儲存參數的data
			Utility.printLogTime("getA05FileData begin time");
				
			try {
				String WMdataDir = Utility.getProperties("WMdataDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");
				String WMdataBKDir = Utility.getProperties("WMdataBKDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");		
				File tmpFile = new File(WMdataDir+filename);		
				Date filedate = new Date(tmpFile.lastModified());
				Map values = new HashMap();//96.04.10 add
				FileReader	f		= new FileReader(tmpFile);
				LineNumberReader in	= new LineNumberReader(f);
				
				//96.08.15 若AXX_TMP暫存table,有資料則先delete
				sqlCmd.append("select count(*) as countdata from AXX_TMP where m_year=? and m_month=? and filename=?");
				paramList.add(m_year);
				paramList.add(m_month);
				paramList.add(filename);
				dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"countdata");				
			
				if(dbData != null && (Integer.parseInt((((DataObject)dbData.get(0)).getValue("countdata")).toString()) != 0)){
					System.out.println("AXX_TMP.size()="+(((DataObject)dbData.get(0)).getValue("countdata")).toString());
					//List deleteList = new LinkedList();
					//deleteList.add("delete AXX_TMP where m_year="+m_year+" and m_month="+m_month+" and filename='"+filename+"'");
					//System.out.println("delete OK??"+DBManager.updateDB(deleteList));
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("delete AXX_TMP where m_year=? and m_month=? and filename=?");
					updateDBDataList.add(paramList);//1:傳內的參數List
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				    
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
				}
				
				//96.08.14 寫入AXX_TMP暫存table
				sqlCmd.delete(0, sqlCmd.length());
				dataList =  new ArrayList();//儲存參數的data
				updateDBDataList = new ArrayList();//儲存參數的List	
				updateDBSqlList = new ArrayList();	
				sqlCmd.append("INSERT INTO AXX_TMP  VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
				
				doLoop://將txt檔案儲存至LinkedList
				while ((txtline = in.readLine()) != null) {
						if (txtline.trim().equals("")){
							continue doLoop;
						}else{
							detail=new LinkedList();
							detail.add(m_year);
							detail.add(m_month);
							if(!bank_codeList.contains(txtline.substring(5,12))){
								bank_codeList.add(txtline.substring(5,12));
								
							}
							//96.11.26 fix 910500/91060P,由程式計算.即使有上傳.也不抓檔案內容 by 2295
							if(txtline.substring(12,18).equals("910500") || txtline.substring(12,18).equals("91060P")){
							   continue;	
							}
							detail.add(txtline.substring(5,12));//bank_code
							detail.add(txtline.substring(12,18));//科目代號
							tmpAcc_Code=txtline.substring(12,18);
							//94.03.03若項目代號最後一碼為N時.取中文名稱18-48
							//        若不為N時,取金額18-32 
							if(tmpAcc_Code.substring(tmpAcc_Code.length()-1).equals("N")){
							    System.out.println("中文的全長="+txtline.length());
							    detail.add("0");//金額 axx_tmp.amt
							    detail.add("0");//99.11.15 add amt1
								detail.add("0");//99.11.15 add amt2
								detail.add("0");//99.11.15 add amt3
								detail.add("0");//99.11.15 add amt4
								detail.add("0");//99.11.15 add amt5
								detail.add("0");//99.11.15 add amt6
								detail.add("0");//99.11.15 add amt7
								detail.add("0");//99.11.15 add amt8
								detail.add("0");//99.11.15 add amt9			
							    detail.add(Utility.ISOtoBig5((Utility.toBig5Convert(txtline)).substring(18,48)));//中文名稱 axx_tmp.amt_name							   
							    System.out.println("中文名稱="+Utility.ISOtoBig5((Utility.toBig5Convert(txtline)).substring(18,48)));
							    detail.add(report_no);//96.08.15
							    detail.add(filename);//96.08.15
							}else{
 							    //94.03.07 fix 有負號("-")時,把之前的"0"去掉.. 
								tmpAmt = txtline.substring(18,32);//金額
								if(tmpAmt.indexOf("-") != -1){
								   tmpAmt = tmpAmt.substring(tmpAmt.indexOf("-"),tmpAmt.length());
								}
								//=========================================									
							    detail.add(tmpAmt);//金額 axx_tmp.amt
							    detail.add("0");//99.11.15 add amt1
								detail.add("0");//99.11.15 add amt2
								detail.add("0");//99.11.15 add amt3
								detail.add("0");//99.11.15 add amt4
								detail.add("0");//99.11.15 add amt5
								detail.add("0");//99.11.15 add amt6
								detail.add("0");//99.11.15 add amt7
								detail.add("0");//99.11.15 add amt8
								detail.add("0");//99.11.15 add amt9		
							    detail.add("");//amt_name
							    detail.add(report_no);//96.08.15
							    detail.add(filename);//96.08.15
							}   
							//System.out.println(detail);
							updateDBDataList.add(detail);//1:傳內的參數List
							A05List.add(detail);
						}
				}	
				in.close();
				f.close();
				//96.08.15 寫入AXX_TMP暫存table
				updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				    
				updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
				updateDBList.add(updateDBSqlList);
				updateDBSqlList.add(A05List);
            	if(DBManager.updateDB_ps(updateDBList)){
            	   System.out.println("AXX_TMP Insert ok");				  	
            	}				
			}catch(Exception e){
				errMsg = errMsg + "UpdateA05.getA05FileData Error:"+e.getMessage()+"<br>";
				return null;
			}
			Utility.printLogTime("getA05FileData end time");
			
			sqlCmd.delete(0, sqlCmd.length());
			sqlCmd.append("select ncacno.acc_code,ncacno.acc_name,a05_assumed.assumed from ncacno left join a05_assumed on ncacno.acc_code=a05_assumed.acc_code where acc_tr_type = 'A05' and acc_div='07' order by acc_range");
			dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),null,"assumed,update_date");//資負表
			Properties prop = new Properties();
			System.out.println("dbData.size()="+dbData.size());
			//94.03.18 fix 計算910500/91060P
			for(int i=0;i<dbData.size();i++){
				if( ((String)((DataObject)dbData.get(i)).getValue("acc_code")).equals("920101")
				||  ((String)((DataObject)dbData.get(i)).getValue("acc_code")).equals("920201")
				||  ((String)((DataObject)dbData.get(i)).getValue("acc_code")).equals("920301")
				||  ((String)((DataObject)dbData.get(i)).getValue("acc_code")).equals("920401")
				||  ((String)((DataObject)dbData.get(i)).getValue("acc_code")).equals("920501")
				||  ((String)((DataObject)dbData.get(i)).getValue("acc_code")).equals("920601")
				||  ((String)((DataObject)dbData.get(i)).getValue("acc_code")).equals("920901")//110.03.09 add
				||  ((String)((DataObject)dbData.get(i)).getValue("acc_code")).equals("920902")//110.03.09 add
				){
					System.out.println("acc_code="+(String)((DataObject)dbData.get(i)).getValue("acc_code"));
					System.out.println("assumed="+(((DataObject)dbData.get(i)).getValue("assumed")).toString());
					prop.put((String)((DataObject)dbData.get(i)).getValue("acc_code"),(((DataObject)dbData.get(i)).getValue("assumed")).toString());//科目代號
				}
			}
			
			
			BigDecimal tmp=null;			
			String propt="";			
			String dotbefor="";
			String dotafter="";
			List List9207X0 = new LinkedList();
			List List9207X0tmp = new LinkedList();
			Properties prop91500P = new Properties();
			String str910400 = "";
			List List910500 = new LinkedList();
			List List91060P = new LinkedList();
			//BigDecimal count910500=new BigDecimal("0.0");
			double count910500=0.0;
			double count91060P=0.0;
			double amt910203=0.0;
			//double amt910109=0.0;//108.01.04 add
			//double amt910401=0.0;//108.01.04 add
			//double amt910402=0.0;//108.01.04 add
			//double amt910403=0.0;//108.01.04 add
			//double amt910404=0.0;//108.01.04 add
			System.out.println("A05List.size()="+A05List.size());
			List countA05List=new LinkedList();
			
			sqlCmd.delete(0, sqlCmd.length());			
			updateDBDataList = new ArrayList();//儲存參數的List	
			updateDBSqlList = new ArrayList();
			updateDBList = new LinkedList();
			sqlCmd.append("INSERT INTO AXX_TMP  VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
			
			
			for(int j=0;j<bank_codeList.size();j++){
				//計算920101.920201.920301.920401.920501.920601.920801.
				A05Loop:
			    for(int i=0;i<A05List.size();i++){			    	
				    if(((String)bank_codeList.get(j)).equals((String)((List)A05List.get(i)).get(2))){
				    	System.out.println("A05List.acc_code="+(String)((List)A05List.get(i)).get(3));
				    	//儲存910400
				    	if(((String)((List)A05List.get(i)).get(3)).equals("910400")){
				    		str910400 = (String)((List)A05List.get(i)).get(4);
				    		System.out.println("get 910400="+str910400);
				    	}
				    	//儲存92071P.92072P.92073P.92074P.92075P,ex:100->0.1,上傳的已經*1000
				    	if((((String)((List)A05List.get(i)).get(3)).equals("92071P"))
				    	|| (((String)((List)A05List.get(i)).get(3)).equals("92072P"))
						|| (((String)((List)A05List.get(i)).get(3)).equals("92073P"))
						|| (((String)((List)A05List.get(i)).get(3)).equals("92074P"))
						|| (((String)((List)A05List.get(i)).get(3)).equals("92075P"))
				    	){
				    		prop91500P.put((String)((List)A05List.get(i)).get(3),(String)((List)A05List.get(i)).get(4));
				    		System.out.println("prop91500P.add--"+(String)((List)A05List.get(i)).get(3)+"="+(String)((List)A05List.get(i)).get(4));
						  	continue A05Loop;
						}
				    	//儲存920710.920720.920730.920740.920750				    	
				    	if((((String)((List)A05List.get(i)).get(3)).equals("920710"))
						|| (((String)((List)A05List.get(i)).get(3)).equals("920720"))
						|| (((String)((List)A05List.get(i)).get(3)).equals("920730"))
						|| (((String)((List)A05List.get(i)).get(3)).equals("920740"))
						|| (((String)((List)A05List.get(i)).get(3)).equals("920750"))
						){	
				    		List9207X0tmp=null;
				    		List9207X0tmp=new LinkedList();
				    		List9207X0tmp.add((String)((List)A05List.get(i)).get(3));//科目代號
				    		List9207X0tmp.add((String)((List)A05List.get(i)).get(4));//金額
				    		System.out.println("List9207X0.add="+List9207X0tmp);
				    		List9207X0.add(List9207X0tmp);
						  	continue A05Loop;
						}
				    	
				    	if((((String)((List)A05List.get(i)).get(3)).equals("920101"))
				    	|| (((String)((List)A05List.get(i)).get(3)).equals("920201"))
						|| (((String)((List)A05List.get(i)).get(3)).equals("920301"))
						|| (((String)((List)A05List.get(i)).get(3)).equals("920401"))
						|| (((String)((List)A05List.get(i)).get(3)).equals("920501"))
						|| (((String)((List)A05List.get(i)).get(3)).equals("920601"))
						|| (((String)((List)A05List.get(i)).get(3)).equals("920901"))//110.03.09 add
						|| (((String)((List)A05List.get(i)).get(3)).equals("920902"))//110.03.09 add
						|| (((String)((List)A05List.get(i)).get(3)).equals("920801"))						 
				        ){
				    		if(((String)((List)A05List.get(i)).get(3)).equals("920801")){
				    			count910500 += Double.parseDouble((String)((List)A05List.get(i)).get(4));
				    			continue A05Loop;
				    		}
				    		tmp = BigDecimal.valueOf(Long.parseLong((String)((List)A05List.get(i)).get(4)));
				    		System.out.println((String)((List)A05List.get(i)).get(3)+".amt="+tmp.toString());
				    		propt=prop.getProperty((String)((List)A05List.get(i)).get(3));
				    		
				    		//資料庫裡的風險權數有0.10要把最後一位0去掉再轉
				    		if(propt.indexOf(".") != -1 && propt.lastIndexOf("0") !=-1){
                               if(propt.lastIndexOf("0") > propt.indexOf(".")){
                               	  propt = propt.substring(0,propt.lastIndexOf("0"));
                               }
				    		}
				    		System.out.println((String)((List)A05List.get(i)).get(3)+".風險權數="+propt);				    		
							tmp = tmp.multiply(new BigDecimal(Double.parseDouble(propt)));
							tmpAmt = tmp.toString();
							System.out.println((String)((List)A05List.get(i)).get(3)+".tmpAmt="+tmpAmt);
							if(tmpAmt.indexOf(".") != -1 && (tmpAmt.length()-4 > tmpAmt.indexOf("."))){
							   tmpAmt = tmpAmt.substring(0,tmpAmt.indexOf(".")+4);	
							}
							System.out.println((String)((List)A05List.get(i)).get(3)+".tmpAmt="+tmpAmt);
							count910500 += Double.parseDouble(tmpAmt);
							System.out.println("累加後的count910500="+count910500);
				    	}
				    	if((((String)((List)A05List.get(i)).get(3)).equals("910203"))){//95.05.11 add 儲存910203
				    	    amt910203 = Double.parseDouble((String)((List)A05List.get(i)).get(4));     				    	
				    	}
				    	/*108.01.28 add A05.910500不扣除910109/910401/910402/910403/910404 by 2295
				    	if((((String)((List)A05List.get(i)).get(3)).equals("910109"))){//108.01.04 add 儲存910109
                            amt910109 = Double.parseDouble((String)((List)A05List.get(i)).get(4));                          
                        }     
				    	if((((String)((List)A05List.get(i)).get(3)).equals("910401"))){//108.01.04 add 儲存910401
                            amt910401 = Double.parseDouble((String)((List)A05List.get(i)).get(4));                          
                        }     
				    	if((((String)((List)A05List.get(i)).get(3)).equals("910402"))){//108.01.04 add 儲存910402
                            amt910402 = Double.parseDouble((String)((List)A05List.get(i)).get(4));                          
                        }     
				    	if((((String)((List)A05List.get(i)).get(3)).equals("910403"))){//108.01.04 add 儲存910403
                            amt910403 = Double.parseDouble((String)((List)A05List.get(i)).get(4));                          
                        }    
				    	if((((String)((List)A05List.get(i)).get(3)).equals("910404"))){//108.01.04 add 儲存910404
                            amt910404 = Double.parseDouble((String)((List)A05List.get(i)).get(4));                          
                        }  
				    	*/
				    }
			    }//end of 920101.920201.920301.920401.920501.920601.920801.910203.910109.910401.910402.910403.910404
			   
			    //count 9207X0====================================================
			    //System.out.println("List9207X0.size()="+List9207X0.size());
			    	
			    for(int i=0;i<List9207X0.size();i++){			    	
	   		        tmp = BigDecimal.valueOf(Long.parseLong((String)((List)List9207X0.get(i)).get(1)));
	   		        //94.07.13 fix 9207XP不為空值時,才做累加================================================
	   		        if(((String)prop91500P.get(((String)((List)List9207X0.get(i)).get(0)).substring(0,5)+"P")).equals("")	   		        		
					|| (String)prop91500P.get(((String)((List)List9207X0.get(i)).get(0)).substring(0,5)+"P") == null){
	   		        	continue;
	   		        }
	   		        //======================================================================================
	   		        System.out.println("List9207X0.tmp="+tmp);
	   		        System.out.println("get prop91500P."+((String)((List)List9207X0.get(i)).get(0)).substring(0,5)+"P");
	   		        System.out.println("prop91500P="+(String)prop91500P.get(((String)((List)List9207X0.get(i)).get(0)).substring(0,5)+"P"));
	   		        System.out.println("prop91500P.double="+Double.parseDouble((String)prop91500P.get(((String)((List)List9207X0.get(i)).get(0)).substring(0,5)+"P")));
	   		        //94.07.13 fix /1000 -> / 100======================================================
	   		        tmp = tmp.multiply(new BigDecimal(Double.parseDouble((String)prop91500P.get(((String)((List)List9207X0.get(i)).get(0)).substring(0,5)+"P"))/1000));
	   		        System.out.println((String)((List)List9207X0.get(i)).get(0)+"="+(String)((List)List9207X0.get(i)).get(1)+"*"+(String)prop91500P.get(((String)((List)List9207X0.get(i)).get(0)).substring(0,5)+"P")+"/1000=List9207X0.tmp="+tmp);
	   		        //========================================================================================
				    System.out.println((String)((List)List9207X0.get(i)).get(0)+".tmp="+tmp);
				    count910500 += Double.parseDouble(tmp.toString());
				    System.out.println("累加後的count910500="+count910500);
			    }
			    //95.05.11 add 原910500 - 910203 ========================================
			    count910500 -= amt910203;
			    System.out.println("累加後的count910500(-910203)="+count910500);
			    /*108.01.28 add A05.910500不扣除910109/910401/910402/910403/910404 by 2295
			    count910500 -= amt910109;
                System.out.println("累加後的count910500(-910203-910109)="+count910500);
                count910500 -= amt910401;
                System.out.println("累加後的count910500(-910203-910109-910401)="+count910500);
                count910500 -= amt910402;
                System.out.println("累加後的count910500(-910203-910109-910401-910402)="+count910500);
                count910500 -= amt910403;
                System.out.println("累加後的count910500(-910203-910109-910401-910402-910403)="+count910500);
                count910500 -= amt910404;
                System.out.println("累加後的count910500(-910203-910109-910401-910402-910403-910404)="+count910500);
                */
			    //=======================================================
			    List910500=new LinkedList();
			    List910500.add((String)bank_codeList.get(j));
			    List910500.add("910500");
			    
			    System.out.println((String)bank_codeList.get(j)+".910500.Math.round="+Math.round(count910500));//103.04.14 add
			    //System.out.println((String)bank_codeList.get(j)+".910500.long="+(new Double(count910500)).longValue()); 
			    List910500.add(String.valueOf(Math.round(count910500)));//103.04.14 910500調整為四捨五入取至整數 			    
			    A05List.add(List910500);//要寫入AXX_TMP
			    //96.11.23 將910500寫入AXX_TMP暫存table				
				detail=new LinkedList();
				detail.add(m_year);
				detail.add(m_month);				
				detail.add((String)bank_codeList.get(j));//bank_code
				detail.add("910500");//科目代號
				detail.add(String.valueOf(Math.round(count910500)));//金額//103.04.14 910500調整為四捨五入取至整數 
				detail.add("0");//99.11.15 add amt1
				detail.add("0");//99.11.15 add amt2
				detail.add("0");//99.11.15 add amt3
				detail.add("0");//99.11.15 add amt4
				detail.add("0");//99.11.15 add amt5
				detail.add("0");//99.11.15 add amt6
				detail.add("0");//99.11.15 add amt7
				detail.add("0");//99.11.15 add amt8
				detail.add("0");//99.11.15 add amt9		
				detail.add("");//amt_name
				detail.add(report_no);//96.08.15
				detail.add(filename);//96.08.15
				updateDBDataList.add(detail);
				countA05List.add(detail);
			    System.out.println((String)bank_codeList.get(j)+".910400="+str910400);
			    //94.07.11 fix 910400不為空值時才計算91060P======================================
			    if(!str910400.equals("") && str910400 != null){
			        System.out.println((String)bank_codeList.get(j)+".910400.double="+Double.parseDouble(str910400));
			        count91060P = Double.parseDouble(str910400) / count910500;
			        System.out.println((String)bank_codeList.get(j)+".91060P.double="+count91060P);
			        count91060P = Math.round(count91060P * 100000);			    
			        List91060P=new LinkedList();
			        List91060P.add((String)bank_codeList.get(j));
			        List91060P.add("91060P");
			        List91060P.add(String.valueOf((new Double(count91060P)).longValue()));
			        System.out.println((String)bank_codeList.get(j)+".91060P="+(new Double(count91060P)).longValue());
			        A05List.add(List91060P);//要寫入AXX_TMP
			        //96.11.23 將91060P寫入AXX_TMP暫存table				
					detail=new LinkedList();
					detail.add(m_year);
					detail.add(m_month);					
					detail.add((String)bank_codeList.get(j));//bank_code
					detail.add("91060P");//科目代號
					detail.add(String.valueOf((new Double(count91060P)).longValue()));//金額
					detail.add("0");//99.11.15 add amt1
					detail.add("0");//99.11.15 add amt2
					detail.add("0");//99.11.15 add amt3
					detail.add("0");//99.11.15 add amt4
					detail.add("0");//99.11.15 add amt5
					detail.add("0");//99.11.15 add amt6
					detail.add("0");//99.11.15 add amt7
					detail.add("0");//99.11.15 add amt8
					detail.add("0");//99.11.15 add amt9		
					detail.add("");//amt_name
					detail.add(report_no);//96.08.15
					detail.add(filename);//96.08.15
					updateDBDataList.add(detail);
					countA05List.add(detail);
			    }
			    //=============================================================================
			    //inital
			    List9207X0 = new LinkedList();
				List9207X0tmp = new LinkedList();
				prop91500P = new Properties();
				str910400 = "";							
				count910500=0.0;
				count91060P=0.0;
				amt910203=0.0;				
			}//end of bank_codeList
			//==============================================================
			//96.11.23 將910500/91060P寫入AXX_TMP暫存table
			updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				    
			updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
			updateDBList.add(updateDBSqlList);
			updateDBSqlList.add(A05List);
        	
        	if(DBManager.updateDB_ps(updateDBList)){
        	   System.out.println("AXX_TMP Insert 910500/91060P ok");				  	
        	}
			allList.add(bank_codeList);
			allList.add(A05List);
			Utility.printLogTime("getA05FileData end-1 time");
				
			return allList;
	}
	
	//公式檢核
	//99.11.15 add 套用DAO.preparestatment,並列印轉換後的SQL 
	private static List CheckRule(String bank_type,String report_no,String m_year,String m_month,
								 String bank_code,String user_id,String user_name){
		
		String  amt="";		
		double	amt_L = 0.0, amt_R = 0.0,amt_tbl = 0.0,amt_assumed=0.0;
		String	quop = null;
		String  cano="";
		String acc_code="";
		String noop="";		
		List dbData = null;
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		//99.11.15 add 
		List updateDBList = new ArrayList();//0:sql 1:data		
		List updateDBSqlList = new ArrayList();
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data
		Date nowlog = new Date();			
		SimpleDateFormat logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");
		Calendar logcalendar = Calendar.getInstance();
		nowlog = logcalendar.getTime();		
		logformat.format(nowlog);
		logcalendar = Calendar.getInstance();
		nowlog = logcalendar.getTime();		
		System.out.println("CheckRule begin time="+logformat.format(nowlog));	
		try{
			//check ruleno1, ruleno2
			//sqlCmd = " select r1.cano,r1.quop,r2.left_flag,decode(r1.cano,'070060',A05.amt*nvl(A05_assumed.assumed,1),A05.amt) as \"amt\",r2.noop,A05.acc_code "
		    //95.05.11 070060不*風險權數
		    sqlCmd.append(" select r1.cano,r1.quop,r2.left_flag,A05.amt,r2.noop,A05.acc_code ");
		    sqlCmd.append(" from A05 left join A05_assumed on A05.acc_code=A05_assumed.acc_code ,ruleno1 r1,ruleno2 r2 ");
		    sqlCmd.append(" where A05.acc_code=r2.acc_code ");
		    sqlCmd.append(" and r1.CANO = r2.cano");
		    sqlCmd.append(" and A05.m_year=? and A05.m_month=?");
		    sqlCmd.append(" and A05.bank_code=?");
		    sqlCmd.append(" order by r1.cano,r2.left_flag,r2.nserial");
		    
			paramList.add(m_year);
			paramList.add(m_month);
			paramList.add(bank_code);
			dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"amt");
			cano=(String)((DataObject)dbData.get(0)).getValue("cano");
			quop = ((String)((DataObject)dbData.get(0)).getValue("quop") == null ? "" : ((String)((DataObject)dbData.get(0)).getValue("quop")).trim());
			System.out.println("check rule begin");
			amt_L = 0.0; amt_R = 0.0;
			amt_tbl=0.0;
			
			sqlCmd.delete(0,sqlCmd.length());
			sqlCmd.append("INSERT INTO WML02 VALUES (?,?,?,?,?,?,?,?,?,sysdate)");
			
			ruleLoop:
			for(int r=0;r<dbData.size();r++){				
				//System.out.println("r="+r);
				
				if(!(((String)((DataObject)dbData.get(r)).getValue("cano")).trim()).equals(cano)){//與前一個公式不同時
					//System.out.println("amt_L="+amt_L);
					//System.out.println("amt_R="+amt_R);
					if ((quop.equals("=") && (amt_L == amt_R))){
					} else if ((quop.equals(">") && (amt_L > amt_R))) {
					} else if ((quop.equals("<") && (amt_L < amt_R))) {
					} else if ((quop.equals(">=") && (amt_L >= amt_R))) {
					} else if ((quop.equals("<=") && (amt_L <= amt_R))) {
					} else if ((quop.equals("!=") && (amt_L != amt_R))) {
					}else {	    	
						System.out.println("WML02 have error:cano="+cano);
						dataList = new ArrayList();//傳內的參數List		 	           				   
						dataList.add(m_year); 
						dataList.add(m_month); 
						dataList.add(bank_code); 
						dataList.add(report_no); 
						dataList.add(cano); 
						dataList.add(String.valueOf(amt_L)); 
						dataList.add(String.valueOf(amt_R)); 
						dataList.add(user_id); 
						dataList.add(user_name); 
						updateDBDataList.add(dataList);//1:傳內的參數List						
					}
					//把這次的公式編號及quop儲存;並把金額清空
					cano=(String)((DataObject)dbData.get(r)).getValue("cano");
					quop = ((String)((DataObject)dbData.get(r)).getValue("quop") == null ? "" : ((String)((DataObject)dbData.get(r)).getValue("quop")).trim());					
					amt_L = 0.0; amt_R = 0.0;
					amt_tbl=0.0;
				}	

				
				/*94.03.17 cano='070500' or '070070' 不做檢核
				 *070500-->910500=920101+920201+920301+920401+920501+920601+920710+920720+920730+920740+920750+920801  
				 *070700-->91060P=910400/910500
				 *070600-->929901=920101+920201+920301+920401+920501+920601+920901+920902+920710+920720+920730+920740+920750+920801
				*/
				if((((String)((DataObject)dbData.get(r)).getValue("cano")).trim()).equals("070050")
				|| (((String)((DataObject)dbData.get(r)).getValue("cano")).trim()).equals("070070")){
					continue ruleLoop;
				}
				
				acc_code = (((DataObject)dbData.get(r)).getValue("acc_code")).toString();
				/*if(cano.equals("070060") && acc_code.substring(acc_code.length()-1).equals("P")){	//第7大項之風險性資產計算風險權數
					System.out.println("in layer1");
					amt_assumed = Double.parseDouble((((DataObject)dbData.get(r)).getValue("amt")).toString())/1000;
					amt_tbl = Double.parseDouble((((DataObject)dbData.get(r+1)).getValue("amt")).toString())*amt_assumed;
				}else{*/
					//System.out.println("in layer2");
  					amt_tbl = Double.parseDouble((((DataObject)dbData.get(r)).getValue("amt")).toString());
				//}
				
				if(cano.equals("070070") && acc_code.substring(acc_code.length()-1).equals("P")){
					amt_tbl = Double.parseDouble((((DataObject)dbData.get(r)).getValue("amt")).toString())/1000;
				}
				//System.out.println("acc_code="+acc_code);
				//System.out.println("noop="+noop);
				//System.out.println("amt_tbl="+amt_tbl);
				if(((String)((DataObject)dbData.get(r)).getValue("left_flag")).equals("0")){//左式					
					if ((noop == null) || (noop.equals(""))) {
						amt_L = amt_tbl;
					}else if (noop.equals("+")) {
						amt_L += amt_tbl;
					}else if (noop.equals("-")) {
						amt_L -= amt_tbl;
					}else if (noop.equals("*")) {
						amt_L *= amt_tbl;
					}else if (noop.equals("/")) {
						//System.out.println("now is /");
						//System.out.println("amt_L="+amt_L);
						if (Double.parseDouble((((DataObject)dbData.get(r)).getValue("amt")).toString()) != 0)
							amt_L /= amt_tbl;
						if(cano.equals("070070")){
							amt_R = (int)Math.round((double)(amt_R*100000));
							amt_R = amt_R/1000;
						}						
					}//end of noop
				}//end of left_flag=0-->左式金額加總	
				if(((String)((DataObject)dbData.get(r)).getValue("left_flag")).equals("1")){//右式
					if ((noop == null) || (noop.equals(""))) {
						amt_R = amt_tbl;
					}else if (noop.equals("+")) {
						amt_R += amt_tbl;
					}else if (noop.equals("-")) {
						amt_R -= amt_tbl;
					}else if (noop.equals("*")) {
						amt_R *= amt_tbl;
					}else if (noop.equals("/")) {
						//System.out.println("now is /");
						//System.out.println("amt_R="+amt_R);
						if (Double.parseDouble((((DataObject)dbData.get(r)).getValue("amt")).toString()) != 0)
							amt_R /= amt_tbl;
						if(cano.equals("070070")){
							amt_R = (int)Math.round((double)(amt_R*100000));							
							amt_R = amt_R/1000;							
						}						
					}//end of noop
				}//end of left_flag=0-->右式金額加總
				
				//94.03.04 fix 把最後一筆的檢核失敗寫入db=================================
				if(r==dbData.size()-1){
				   if ((quop.equals("=") && (amt_L == amt_R))){
				   } else if ((quop.equals(">") && (amt_L > amt_R))) {
				   } else if ((quop.equals("<") && (amt_L < amt_R))) {
				   } else if ((quop.equals(">=") && (amt_L >= amt_R))) {
				   } else if ((quop.equals("<=") && (amt_L <= amt_R))) {
				   } else if ((quop.equals("!=") && (amt_L != amt_R))) {
				   }else {	    	
				   		System.out.println("WML02 have error:cano="+cano);
				   		dataList = new ArrayList();//傳內的參數List		 	           				   
				   		dataList.add(m_year); 
				   		dataList.add(m_month); 
				   		dataList.add(bank_code); 
				   		dataList.add(report_no); 
				   		dataList.add(cano); 
				   		dataList.add(String.valueOf(amt_L)); 
				   		dataList.add(String.valueOf(amt_R)); 
				   		dataList.add(user_id); 
				   		dataList.add(user_name); 
				   		updateDBDataList.add(dataList);//1:傳內的參數List	
				   }
				}
				//==========================================================
				/*if(cano.equals("070060") && acc_code.substring(acc_code.length()-1).equals("P"))	//略過下一個為該資產之帳面金額
					r++;*/
				noop = (((DataObject)dbData.get(r)).getValue("noop") == null ? "" : ((String)((DataObject)dbData.get(r)).getValue("noop")).trim());
				//System.out.println("r="+r+"end");				
			} //end of for()
			System.out.println("WML02 have error.updateDBDataList.size()="+updateDBDataList.size());
			if(updateDBDataList.size() > 0){//99.11.15當有檢核失敗時,才加入
			   updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql	
			   updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
			   updateDBList.add(updateDBSqlList);
			}
			System.out.println("check rule end");
		}catch(Exception e){
			errMsg = errMsg + "UpdateA05.CheckRule Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateA05.CheckRule Error:"+e.getMessage());
		}
		
		return updateDBList;
	}
	
 	//96.04.18 ad 批次增加額外科目代號	
	//99.11.15 add 套用DAO.preparestatment,並列印轉換後的SQL by 2295 
	private static boolean InsertZeroExtraA05_List(String m_year,String m_month,List bank_code){
		StringBuffer sqlCmd = new StringBuffer();		
		List updateDBList = new LinkedList();//0:sql 1:data		
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data
		String wlx01_m_year = (Integer.parseInt(m_year) < 100)?"99":"100"; 
		boolean updateOK=false;		
		Utility.printLogTime("InsertZeroExtraA05_List begin time");	
			
		try{
			
			/*若無資料在A05時,則新增zero至A05
			INSERT INTO A05
			select 96,3,a.bank_no,acc_code ,0,'' 
			from ncacno,(select bank_no from bn01 where bank_no in ('6030016','6040017'))a,--目前需更新的bank_code 
			            (select bank_code from A05 where m_year=94 AND m_month=6 AND 
			             bank_code in ('6030016','6040017')group by bank_code)b --A05已有資料的 
			where acc_tr_type = 'A05' and acc_code in ('910203','910404') and acc_div='07' 
			and a.bank_no = b.bank_code(+)
			and b.bank_code is null --無資料在A05的bank_no
			order by a.bank_no,acc_range"
			*/
			sqlCmd.append(" INSERT INTO A05"); 
			sqlCmd.append(" select ?,?,a.bank_no,acc_code ,0,''"); 
			sqlCmd.append(" from ncacno,(select bank_no from bn01 where m_year=? and bank_no in ("+Utility.getCombinString(bank_code,",")+"))a");//目前需更新的bank_code 
			sqlCmd.append("            ,(select bank_code from A05");
			sqlCmd.append("              where m_year=? AND m_month=?");
			sqlCmd.append("              AND bank_code in ("+Utility.getCombinString(bank_code,",")+")group by bank_code)b");//A05已有資料的 
			sqlCmd.append(" where acc_tr_type = 'A05' and acc_code in ('910203','910404') and acc_div='07'"); 
			sqlCmd.append(" and a.bank_no = b.bank_code(+)");
			sqlCmd.append(" and b.bank_code is null ");//無資料在A05的bank_no
			sqlCmd.append(" order by a.bank_no,acc_range");		
			dataList.add(m_year);
			dataList.add(m_month);
			dataList.add(wlx01_m_year);
			dataList.add(m_year);
			dataList.add(m_month);			
			updateDBDataList.add(dataList);
			updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				    
			updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
			updateDBList.add(updateDBSqlList);
			updateOK=DBManager.updateDB_ps(updateDBList);			
			
			System.out.println("InsertZeroExtraA05_List OK??"+updateOK);
			if(!updateOK){
				//errMsg = errMsg + "InsertZeroExtraA05_List Fail:"+DBManager.getErrMsg()+"<br>";
				System.out.println(DBManager.getErrMsg());
			}
		}catch(Exception e){
			errMsg = errMsg + "UpdateA05.InsertZeroExtraA05_List Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateA05.InsertZeroExtraA05_List Error:"+e.getMessage());
		}
		Utility.printLogTime("InsertZeroExtraA05_List end time");	
		return updateOK;
	}
	
    //94.07.19 add 檢核910201+910202 < 910199
	//             if(910201+910202) < 910199{
	//                910299=910201+910202
	//             }else{910299=910199}       
	//94.08.19 add 檢核910199 < 0 and 910299 != 0 ==>第一類資本合計(910199)為負數,第二類資本(910299)應為0元
    //95.05.30 add 910109 > 0時,910202-910203也要等於0;107.06.21 fix 農金局-施秀芬來電.此項檢核已於103年更改規定.移除此檢核條件 
	//95.08.17 fix 910202-910203 可小於 0 ,add 910202-910203 < 0 ,則910202-910203等於0
	//99.11.15 add 套用DAO.preparestatment,並列印轉換後的SQL 
	//105.03.29 add 檢核 1.若A05.910202 不等於 a01.加總
	//                 2.若A05.910203不等於 a10.加總 
	//                 3.若A05.910204,不為(【A05.910202】＋【A05.910203】)』與『(【A05.910500】*1.25%）』相較後，取較小者)
	//108.01.04 add A05.910204=min([910202-910203],[910500 * 1.75%])
	private static List CheckOtherRule(int errCount,String bank_type,String report_no,String m_year,String m_month,
				 					   String bank_code,String user_id,String user_name){
	    BigDecimal tmp910109=null;//95.05.29 add
	    BigDecimal tmp910201=null;
		BigDecimal tmp910202=null;
		BigDecimal tmp910203=null;
		BigDecimal tmp910204=null;
		BigDecimal tmp910199=null;
		BigDecimal tmp910299=null;
		BigDecimal tmp910500=null;
		BigDecimal tmpZero=new BigDecimal("0");		
		BigDecimal tmp175=new BigDecimal("0.0175");//108.01.04 fix
		BigDecimal tmpAmt_910202=null;
		BigDecimal tmpAmt_910299=null;	
		
		String  amt="";
		List dbData = null;
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		//99.11.15 add 
		List updateDBList = new ArrayList();//0:sql 1:data		
		List updateDBSqlList = new ArrayList();
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data
		try{
			 sqlCmd.append("select acc_code,amt from A05 where bank_code=? and m_year=? and m_month=?");
			 paramList.add(bank_code);
			 paramList.add(m_year);
			 paramList.add(m_month);
		     dbData=DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"amt");
		     if(dbData != null && dbData.size() != 0){
		         for(int i=0;i<dbData.size();i++){
		             if(((String)((DataObject)dbData.get(i)).getValue("acc_code")).equals("910109")){//95.05.29 add
		         	 	 tmp910109=BigDecimal.valueOf(Long.parseLong((((DataObject)dbData.get(i)).getValue("amt")).toString()));		         	 	
		         	 }
		         	 if(((String)((DataObject)dbData.get(i)).getValue("acc_code")).equals("910201")){
		         	 	 tmp910201=BigDecimal.valueOf(Long.parseLong((((DataObject)dbData.get(i)).getValue("amt")).toString()));		         	 	
		         	 }
		         	 if(((String)((DataObject)dbData.get(i)).getValue("acc_code")).equals("910202")){		         
		         	 	 tmp910202=BigDecimal.valueOf(Long.parseLong((((DataObject)dbData.get(i)).getValue("amt")).toString()));
		         	 }
		         	 if(((String)((DataObject)dbData.get(i)).getValue("acc_code")).equals("910203")){		         
		         	 	 tmp910203=BigDecimal.valueOf(Long.parseLong((((DataObject)dbData.get(i)).getValue("amt")).toString()));
		         	 }
		         	 if(((String)((DataObject)dbData.get(i)).getValue("acc_code")).equals("910204")){		         
		         	 	 tmp910204=BigDecimal.valueOf(Long.parseLong((((DataObject)dbData.get(i)).getValue("amt")).toString()));
		         	 }
		         	 if(((String)((DataObject)dbData.get(i)).getValue("acc_code")).equals("910199")){		         	 
		         	 	 tmp910199=BigDecimal.valueOf(Long.parseLong((((DataObject)dbData.get(i)).getValue("amt")).toString()));
		         	 }
		         	 if(((String)((DataObject)dbData.get(i)).getValue("acc_code")).equals("910299")){		         	 
		         	 	tmp910299=BigDecimal.valueOf(Long.parseLong((((DataObject)dbData.get(i)).getValue("amt")).toString()));
		         	 }		         	 	
		         	 if(((String)((DataObject)dbData.get(i)).getValue("acc_code")).equals("910500")){		         	 
		         	 	tmp910500=BigDecimal.valueOf(Long.parseLong((((DataObject)dbData.get(i)).getValue("amt")).toString()));
		         	 }
		         }
		     }//end of get 910201.910202.910199.910299	
		     System.out.println("CheckOtherRule.bank_code="+bank_code);
		     System.out.println("910109="+tmp910109);
		     System.out.println("910201="+tmp910201);
			 System.out.println("910202="+tmp910202);
			 System.out.println("910203="+tmp910203);
			 System.out.println("910204="+tmp910204);
			 System.out.println("910199="+tmp910199);
			 System.out.println("910299="+tmp910299);
			 System.out.println("910500="+tmp910500);
			 System.out.println("errCount="+errCount);
			 tmpAmt_910202 = tmp910202.add(tmp910203.negate());//910202-910203
			 System.out.println("tmpAmt_910202(910202-910203)="+tmpAmt_910202);
			 //95.08.17 add 910202-910203 < 0 ,則910202-910203等於0
			 if(tmpAmt_910202.compareTo(tmpZero) == -1){//910202-910203 < 0 
			     tmpAmt_910202 = new BigDecimal("0");
			     System.out.println("setZero.tmpAmt_910202(910202-910203)="+tmpAmt_910202);
			 }
			 
			 sqlCmd.delete(0,sqlCmd.length());
			 paramList = new ArrayList();
			 
			 sqlCmd.append("select a01.bank_code,a01_sum,a10.above_loan_sum ");
			 sqlCmd.append("from ");
			 sqlCmd.append("(select bank_code,sum(amt) as a01_sum ");//--a01.加總
			 sqlCmd.append("from a01 ");
			 sqlCmd.append("where m_year=? and m_month=? and bank_code=? ");
			 if("6".equals(bank_type)){
			     //農會 112600    減:備抵跌價損失--有價證券  
			     sqlCmd.append("and acc_code in ('112600','110800','111000','120800','150300','130300','152300') ");
			 }else{
			     //漁會 112300    減:備抵跌價損失-有價證券  
			     sqlCmd.append("and acc_code in ('112300','110800','111000','120800','150300','130300','152300') ");
			 }
			 sqlCmd.append("group by bank_code)a01");     
			 sqlCmd.append(",");
			 sqlCmd.append("(select bank_code, ");        
			 sqlCmd.append("         round(sum(loan2_amt)*0.02/1,0)+ round(sum(loan3_amt)*0.5/1,0)+ round(sum(loan4_amt)/1,0)+ round(sum(property_loss)/1,0) as above_loan_sum ");//--a10.加總         
			 sqlCmd.append("from a10 ");
			 sqlCmd.append("where m_year=? and m_month=? and bank_code=? ");
			 sqlCmd.append("group by bank_code)a10 ");
			 sqlCmd.append("where a01.bank_code=a10.bank_code ");
			 paramList.add(m_year);
			 paramList.add(m_month);
			 paramList.add(bank_code);
			 paramList.add(m_year);
			 paramList.add(m_month);
			 paramList.add(bank_code);
		     List dbData1=DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"bank_code,a01_sum,above_loan_sum");
		     BigDecimal a01_sum=null;	
			 BigDecimal above_loan_sum=null;
			 if(dbData1 != null && dbData1.size() != 0){
		         for(int i=0;i<dbData1.size();i++){
		        	 a01_sum=BigDecimal.valueOf(Long.parseLong((((DataObject)dbData1.get(i)).getValue("a01_sum")).toString()));		         	 	
		             above_loan_sum=BigDecimal.valueOf(Long.parseLong((((DataObject)dbData1.get(i)).getValue("above_loan_sum")).toString()));		         	 	
		         }
		     }
			 System.out.println("a01_sum="+a01_sum);
			 System.out.println("above_loan_sum="+above_loan_sum);
			 
			 if(dbData.size()> 0 && dbData1.size()> 0){
			 
				 sqlCmd.delete(0,sqlCmd.length());
				 sqlCmd.append("INSERT INTO WML03 VALUES (?,?,?,?,?,?,?,?,sysdate)");
				  
				 // 95.05.30 add 910109 > 0時,910202-910203也要等於0
				 //107.06.21 fix 農金局-施秀芬來電.此項檢核已於103年更改規定.移除此檢核條件 
				 /*
			     if(tmp910109.compareTo(tmpZero) == 1){//910109 > 0
			        if(tmpAmt_910202.compareTo(tmpZero) != 0){//910202-910203 != 0
			            System.out.println("errCount="+errCount);
			            dataList = new ArrayList();//傳內的參數List	
			            dataList.add(m_year); 
						dataList.add(m_month); 
						dataList.add(bank_code);   
						dataList.add(report_no);  
						dataList.add(String.valueOf(errCount));  
						dataList.add("檢核發現(910109)["+Utility.setCommaFormat(tmp910109.toString())+"]元 大於0,而(【910202】-【910203】)金額為["+Utility.setCommaFormat(tmpAmt_910202.toString())+"]元,不為0");  
						dataList.add(user_id);  
						dataList.add(user_name);
						updateDBDataList.add(dataList);//1:傳內的參數List
						
			            //99.11.15sqlCmd = "INSERT INTO WML03 VALUES (" +
					    //       m_year + "," + m_month + ",'" + bank_code + "','" + report_no + "'," +
						//       errCount + ",'檢核發現(910109)["+Utility.setCommaFormat(tmp910109.toString())+"]元 大於0,"
						//       +"而(【910202】-【910203】)金額為["+Utility.setCommaFormat(tmpAmt_910202.toString())+"]元,不為0','"+user_id+"','"+user_name+"',sysdate)";
			            errCount++ ;	
			            updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql	
						updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
						updateDBList.add(updateDBSqlList);		           
			            return updateDBList;
			        }
			     }
			     */
			     /*95.08.17 910202 fix 可大於或小於910203
				 if(tmpAmt_910202.compareTo(tmpZero) == -1){//910202 < 910203
				    System.out.println("errCount="+errCount); 
					sqlCmd = "INSERT INTO WML03 VALUES (" +
					       m_year + "," + m_month + ",'" + bank_code + "','" + report_no + "'," +
						   errCount + ",'檢核發現【910202】["+Utility.setCommaFormat(tmp910202.toString())+"元]不應小於【910203】["+Utility.setCommaFormat(tmp910203.toString())+"元]','"+user_id+"','"+user_name+"',sysdate)";
						   errCount++ ;	
				    updateDBSqlList.add(sqlCmd);
				    return updateDBSqlList;
				 }else{//910202 > 910203*/ 			    
			     //95.05.29 fix 910500*1.25%->取到整數,小數點下以無條件捨去
				 //108.01.04 fix 910500*1.75%->取到整數,小數點下以無條件捨去
			     tmp910500 = new BigDecimal(tmp910500.multiply(tmp175).intValue());
			     System.out.println("910500*1.75%="+tmp910500);
			     //(910202-910203) > 910500 * 1.75% -> tmpAmt_910202 = 910500 * 1.75%
			     if(tmpAmt_910202.compareTo(tmp910500) == 1){				    
			        tmpAmt_910202 = tmp910500; 
			     }
			     
			     System.out.println("-910203="+tmp910203.negate());
			     System.out.println("tmpAmt_910202(MIN[(910202-910203):(910500*1.75%)])="+tmpAmt_910202);
			     
			     //108.01.28取消使用
			     /*
			     if(tmpAmt_910299.compareTo(tmp910199) == -1){//910201+tmp910202 < 910199
			         if(tmp910199.compareTo(tmpZero) == -1 && tmp910299.compareTo(tmpZero) != 0){//910199 < 0 && 910299 != 0
			            System.out.println("errCount="+errCount);
			            
			            dataList = new ArrayList();//傳內的參數List	
			            dataList.add(m_year); 
						dataList.add(m_month); 
						dataList.add(bank_code);   
						dataList.add(report_no);  
						dataList.add(String.valueOf(errCount));  
						dataList.add("第一類資本合計(910199)為負數,第二類資本(910299)應為0元");  
						dataList.add(user_id);  
						dataList.add(user_name);
						updateDBDataList.add(dataList);//1:傳內的參數List
			            
			            //99.11.15sqlCmd = "INSERT INTO WML03 VALUES (" +
					    //       m_year + "," + m_month + ",'" + bank_code + "','" + report_no + "'," +
						//       errCount + ",'第一類資本合計(910199)為負數,第二類資本(910299)應為0元','"+user_id+"','"+user_name+"',sysdate)";
			            errCount++ ;	
			            updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql	
						updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
						updateDBList.add(updateDBSqlList);		           
			            return updateDBList;			            
			         }else if(tmp910199.compareTo(tmpZero) == -1 && tmp910299.compareTo(tmpZero) == 0){//910199 < 0 && 910299 = 0			             
			         }else if(tmp910299.compareTo(tmpAmt_910299) == 0){//910299 = tmp910299(910201+[910202-910203])
			         }else{
			             System.out.println("errCount="+errCount);
			             dataList = new ArrayList();//傳內的參數List	
				         dataList.add(m_year); 
						 dataList.add(m_month); 
						 dataList.add(bank_code);   
						 dataList.add(report_no);  
					 	 dataList.add(String.valueOf(errCount));  
						 dataList.add("因為X值=『【910201】+MIN((【910202】-【910203】)：(【910500】*1.75%))』 ["+
						    	      Utility.setCommaFormat(tmpAmt_910299.toString())+"元]小於 (910199)["+Utility.setCommaFormat(tmp910199.toString())+
						    	      "元],檢核發現(910299)["+Utility.setCommaFormat(tmp910299.toString())+"元]不等於X值["+Utility.setCommaFormat(tmpAmt_910299.toString())+"]元");  
						 dataList.add(user_id);  
						 dataList.add(user_name);
						 updateDBDataList.add(dataList);//1:傳內的參數List
						 //updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
			             //99.11.15sqlCmd = "INSERT INTO WML03 VALUES (" +
						 //   	m_year + "," + m_month + ",'" + bank_code + "','" + report_no + "'," +
						 //   	errCount + ",'因為X值=『【910201】+MIN((【910202】-【910203】)：(【910500】-【910203】)*1.25%))』 ["+
						 //   	Utility.setCommaFormat(tmpAmt_910299.toString())+"元]小於 (910199)["+Utility.setCommaFormat(tmp910199.toString())+
						 //   	"元],檢核發現(910299)["+Utility.setCommaFormat(tmp910299.toString())+"元]不等於X值["+Utility.setCommaFormat(tmpAmt_910299.toString())+"]元','"+
						 //   	user_id+"','"+user_name+"',sysdate)";
			             errCount++ ;	
			             //updateDBSqlList.add(sqlCmd);
			         }    
			     }else{//910201+tmp910202 >= 910199
			         if(tmp910199.compareTo(tmpZero) == -1 && tmp910299.compareTo(tmpZero) != 0){//910199 < 0 && 910299 != 0
			             System.out.println("errCount="+errCount);
			             dataList = new ArrayList();//傳內的參數List	
			             dataList.add(m_year); 
						 dataList.add(m_month); 
						 dataList.add(bank_code);   
						 dataList.add(report_no);  
					 	 dataList.add(String.valueOf(errCount));  
						 dataList.add("第一類資本合計(910199)為負數,第二類資本(910299)應為0元");  
						 dataList.add(user_id);  
						 dataList.add(user_name);
						 updateDBDataList.add(dataList);//1:傳內的參數List
			             //99.11.15sqlCmd = "INSERT INTO WML03 VALUES (" +
			             //		m_year + "," + m_month + ",'" + bank_code + "','" + report_no + "'," +
			             //		errCount + ",'第一類資本合計(910199)為負數,第二類資本(910299)應為0元','"+user_id+"','"+user_name+"',sysdate)";
			             errCount++ ;	
			             updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql	
						 updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
						 updateDBList.add(updateDBSqlList);		           
				         return updateDBList;
			         }else if(tmp910199.compareTo(tmpZero) == -1 && tmp910299.compareTo(tmpZero) == 0){//910199 < 0 && 910299 = 0			             
			         }else if(tmp910299.compareTo(tmp910199) == 0){//910299 = 910199
			         }else{
			             System.out.println("errCount="+errCount);
			             dataList = new ArrayList();//傳內的參數List	
			             dataList.add(m_year); 
						 dataList.add(m_month); 
						 dataList.add(bank_code);   
						 dataList.add(report_no);  
					 	 dataList.add(String.valueOf(errCount));  
						 dataList.add("第一類資本合計(910199)為負數,第二類資本(910299)應為0元因為X值=『【910201】+MIN((【910202】-【910203】)：(【910500】*1.75%))』 ["+
						 		      Utility.setCommaFormat(tmpAmt_910299.toString())+"元]"+((tmpAmt_910299.compareTo(tmp910199) == 0)?"等於":"大於")+"(910199)["+Utility.setCommaFormat(tmp910199.toString())+
						    	      "元],檢核發現(910299)["+Utility.setCommaFormat(tmp910299.toString())+"元]不等於(910199)["+Utility.setCommaFormat(tmp910199.toString())+"]元");  
						 dataList.add(user_id);  
						 dataList.add(user_name);
						 updateDBDataList.add(dataList);//1:傳內的參數List
			             //99.11.5sqlCmd = "INSERT INTO WML03 VALUES (" +
						 //   	m_year + "," + m_month + ",'" + bank_code + "','" + report_no + "'," +
						 //   	errCount + ",'因為X值=『【910201】+MIN((【910202】-【910203】)：(【910500】-【910203】)*1.25%))』 ["+
						 //   	Utility.setCommaFormat(tmpAmt_910299.toString())+"元]"+((tmpAmt_910299.compareTo(tmp910199) == 0)?"等於":"大於")+"(910199)["+Utility.setCommaFormat(tmp910199.toString())+
						 //   	"元],檢核發現(910299)["+Utility.setCommaFormat(tmp910299.toString())+"元]不等於(910199)["+Utility.setCommaFormat(tmp910199.toString())+"]元','"+
						 //   	user_id+"','"+user_name+"',sysdate)";
			             errCount++ ;
			         }
			     }
			     */
			     if(tmp910199.compareTo(tmpZero) == -1 && tmp910299.compareTo(tmpZero) != 0){//910199 < 0 && 910299 != 0
			         System.out.println("errCount="+errCount);
			         dataList = new ArrayList();//傳內的參數List	
			         dataList.add(m_year); 
					 dataList.add(m_month); 
					 dataList.add(bank_code);   
					 dataList.add(report_no);  
				 	 dataList.add(String.valueOf(errCount));  
					 dataList.add("第一類資本合計(910199)為負數,第二類資本(910299)應為0元");  
					 dataList.add(user_id);  
					 dataList.add(user_name);
					 updateDBDataList.add(dataList);//1:傳內的參數List
			         //99.11.15sqlCmd = "INSERT INTO WML03 VALUES (" +
	             	 //		m_year + "," + m_month + ",'" + bank_code + "','" + report_no + "'," +
	             	 //		errCount + ",'第一類資本合計(910199)為負數,第二類資本(910299)應為0元','"+user_id+"','"+user_name+"',sysdate)";
			         errCount++ ;			         
			     }			
			     //910299 > 0,910299=min([910201+910204],910199) 108.01.28 add
			     tmpAmt_910299 = tmp910201.add(tmp910204);//108.01.28 fix
                 System.out.println("tmpAmt_910299(910201+910204)="+tmpAmt_910299);
                 System.out.println("tmpAmt_910299(910201+910204) compare tmp910199="+tmpAmt_910299.compareTo(tmp910199));             
			     if(tmp910299.compareTo(tmpZero) > 0){//910299 > 0,910299=min([910201+910204],910199)
			        if(tmpAmt_910299.compareTo(tmp910199) < 0 && tmp910299.compareTo(tmpAmt_910299) != 0){//[910201+910204] < 910199
			            System.out.println("errCount="+errCount);
	                     dataList = new ArrayList();//傳內的參數List 
	                     dataList.add(m_year); 
	                     dataList.add(m_month); 
	                     dataList.add(bank_code);   
	                     dataList.add(report_no);  
	                     dataList.add(String.valueOf(errCount));  
	                     dataList.add("第一類資本合計(910199)大於0元,檢核發現第二類資本(910299)["+tmp910299+"元]不等於min([910201+910204],910199)["+tmpAmt_910299+"]元");  
	                     dataList.add(user_id);  
	                     dataList.add(user_name);
	                     updateDBDataList.add(dataList);//1:傳內的參數List
	                     errCount++ ;  
			        }
			        if(tmpAmt_910299.compareTo(tmp910199) > 0 && tmp910299.compareTo(tmp910199) != 0){//[910201+910204] > 910199
                        System.out.println("errCount="+errCount);
                         dataList = new ArrayList();//傳內的參數List 
                         dataList.add(m_year); 
                         dataList.add(m_month); 
                         dataList.add(bank_code);   
                         dataList.add(report_no);  
                         dataList.add(String.valueOf(errCount));  
                         dataList.add("第一類資本合計(910199)大於0元,檢核發現第二類資本(910299)["+tmp910299+"元]不等於min([910201+910204],910199)["+tmp910199+"]元");  
                         dataList.add(user_id);  
                         dataList.add(user_name);
                         updateDBDataList.add(dataList);//1:傳內的參數List
                         errCount++ ;  
                    }
			     }
			     //若A05.910202 不等於 a01.加總 
			     if(tmp910202.compareTo(a01_sum) != 0){
			    	 System.out.println("errCount="+errCount);
			         dataList = new ArrayList();//傳內的參數List	
			         dataList.add(m_year); 
					 dataList.add(m_month); 
					 dataList.add(bank_code);   
					 dataList.add(report_no);  
				 	 dataList.add(String.valueOf(errCount));
				 	 //107.01.10 fix 區分備抵跌價損失-有價證券.農會為112600.漁會為112300
					 dataList.add("BIS計算表之「備抵呆帳、損失準備及營業準備」不等於資產負債表之「備抵呆帳、備抵跌價損失」等科目合計數<br>"+
							      "（【A05.910202】不等於【A01."+(bank_type.equals("6")?"112600":"112300")+"】＋【A01.110800】＋【A01.111000】【A01.120800】＋【A01.150300】＋【A01.130300】＋【A01.152300】["+Utility.setCommaFormat(a01_sum.toString())+"元] ）");  
					 dataList.add(user_id);  
					 dataList.add(user_name);
					 updateDBDataList.add(dataList);//1:傳內的參數List
			         errCount++ ;	
			     }
			     //若A05.910203不等於 a10.加總
			     if(tmp910203.compareTo(above_loan_sum) !=0){
			    	 System.out.println("errCount="+errCount);
			         dataList = new ArrayList();//傳內的參數List	
			         dataList.add(m_year); 
					 dataList.add(m_month); 
					 dataList.add(bank_code);   
					 dataList.add(report_no);  
				 	 dataList.add(String.valueOf(errCount));  
					 dataList.add("BIS計算表之「減項：特定損失所提列備抵呆帳、損失準備及營業準備」不等於A10應予評估資產彙總資料之「第二類放款*2%+第三類放款*50%+第四類放款*100%」之合計數加「非授信資產可能遭受損失」"+
							 	  "（【A05.910203】不等於【A10.第二類放款*2%】＋【A10. 第三類放款*50%】＋【A10.第四類放款*100%】＋【A10.非授信資產可能遭受損失】["+Utility.setCommaFormat(above_loan_sum.toString())+"元]）");  
					 dataList.add(user_id);  
					 dataList.add(user_name);
					 updateDBDataList.add(dataList);//1:傳內的參數List
			         errCount++ ;	
			     }
			     //若A05.910204,不為(【A05.910202】－【A05.910203】)』與『(【A05.910500】*1.75%）』相較後，取較小者)
			     if(tmp910204.compareTo(tmpAmt_910202) !=0){
			    	 System.out.println("errCount="+errCount);
			         dataList = new ArrayList();//傳內的參數List	
			         dataList.add(m_year); 
					 dataList.add(m_month); 
					 dataList.add(bank_code);   
					 dataList.add(report_no);  
				 	 dataList.add(String.valueOf(errCount));  
					 dataList.add("BIS計算表之「調整後備抵呆帳、損失準備及營業準備」非為『A05之「備抵呆帳、損失準備及營業準備」、「減項：特定損失所提列備抵呆帳、損失準備及營業準備」』與『A05之風險性資產總額*1.75%』相較後，取較小者"+
							 	  "（【A05.910204】不等於『(【A05.910202】－【A05.910203】)』與『(【A05.910500】*1.75%）』相較後，取較小者)["+Utility.setCommaFormat(tmpAmt_910202.toString())+"元]");  
					 dataList.add(user_id);  
					 dataList.add(user_name);
					 updateDBDataList.add(dataList);//1:傳內的參數List
			         errCount++ ;	
			     }
			 }
		}catch(Exception e){
			errMsg = errMsg + "UpdateA05.CheckOtherRule Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateA05.CheckOtherRule Error:"+e.getMessage());
		}
		if(updateDBDataList.size() > 0){//99.11.15當有檢核失敗時,才加入
		   updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql	
		   updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
		   updateDBList.add(updateDBSqlList);
		}
		return updateDBList;
	}
	
	//94.07.19 add 檢核910201+910202 < 910199
	//             if(910201+910202) < 910199{
	//                910299=910201+910202
	//             }else{910299=910199}       
	//94.08.19 add 檢核910199 < 0 and 910299 != 0 ==>第一類資本合計(910199)為負數,第二類資本(910299)應為0元
    //95.05.30 add 910109 > 0時,910202-910203也要等於0
	//95.08.17 fix 910202-910203 可小於 0 ,add 910202-910203 < 0 ,則910202-910203等於0
    //96.07.05 add 910204 調整後備抵呆帳、損失準備及營業準備,修正檢核公式 
	/*
	private static List CheckOtherRule(int errCount,String bank_type,String report_no,String m_year,String m_month,
				 					   String bank_code,String user_id,String user_name){
	    BigDecimal tmp910109=null;//95.05.29 add
	    BigDecimal tmp910201=null;
		BigDecimal tmp910202=null;
		BigDecimal tmp910203=null;
		BigDecimal tmp910204=null;//96.07.03 add
		BigDecimal tmp910199=null;
		BigDecimal tmp910299=null;
		BigDecimal tmp910500=null;
		BigDecimal tmpZero=new BigDecimal("0");		
		BigDecimal tmp125=new BigDecimal("0.0125");
		BigDecimal tmpAmt_910202=null;
		BigDecimal tmpAmt_910299=null;
		String	sqlCmd	 = null;	
		String  amt="";		
		List updateDBSqlList = new LinkedList();		
		List dbData = null;
		Date nowlog = new Date();			
		SimpleDateFormat logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");
		Calendar logcalendar = Calendar.getInstance();
		nowlog = logcalendar.getTime();		
		logformat.format(nowlog);
		logcalendar = Calendar.getInstance();
		nowlog = logcalendar.getTime();		
		System.out.println("CheckOtherRule begin time="+logformat.format(nowlog));	
		try{
		     dbData=DBManager.QueryDB("select acc_code,amt from A05 where bank_code='"+bank_code+"' and m_year="+m_year+" and m_month="+m_month,"amt");
		     if(dbData != null && dbData.size() != 0){
		         for(int i=0;i<dbData.size();i++){
		             if(((String)((DataObject)dbData.get(i)).getValue("acc_code")).equals("910109")){//95.05.29 add
		         	 	 tmp910109=BigDecimal.valueOf(Long.parseLong((((DataObject)dbData.get(i)).getValue("amt")).toString()));		         	 	
		         	 }
		         	 if(((String)((DataObject)dbData.get(i)).getValue("acc_code")).equals("910201")){
		         	 	 tmp910201=BigDecimal.valueOf(Long.parseLong((((DataObject)dbData.get(i)).getValue("amt")).toString()));		         	 	
		         	 }
		         	 if(((String)((DataObject)dbData.get(i)).getValue("acc_code")).equals("910202")){		         
		         	 	 tmp910202=BigDecimal.valueOf(Long.parseLong((((DataObject)dbData.get(i)).getValue("amt")).toString()));
		         	 }
		         	 if(((String)((DataObject)dbData.get(i)).getValue("acc_code")).equals("910203")){		         
		         	 	 tmp910203=BigDecimal.valueOf(Long.parseLong((((DataObject)dbData.get(i)).getValue("amt")).toString()));
		         	 }
		         	 if(((String)((DataObject)dbData.get(i)).getValue("acc_code")).equals("910204")){//96.07.03 add 910204 調整後備抵呆帳、損失準備及營業準備		         
		         	 	 tmp910204=BigDecimal.valueOf(Long.parseLong((((DataObject)dbData.get(i)).getValue("amt")).toString()));
		         	 }
		         	 if(((String)((DataObject)dbData.get(i)).getValue("acc_code")).equals("910199")){		         	 
		         	 	 tmp910199=BigDecimal.valueOf(Long.parseLong((((DataObject)dbData.get(i)).getValue("amt")).toString()));
		         	 }
		         	 if(((String)((DataObject)dbData.get(i)).getValue("acc_code")).equals("910299")){		         	 
		         	 	tmp910299=BigDecimal.valueOf(Long.parseLong((((DataObject)dbData.get(i)).getValue("amt")).toString()));
		         	 }		         	 	
		         	if(((String)((DataObject)dbData.get(i)).getValue("acc_code")).equals("910500")){		         	 
		         	 	tmp910500=BigDecimal.valueOf(Long.parseLong((((DataObject)dbData.get(i)).getValue("amt")).toString()));
		         	 }
		         }
		     }//end of get 910201.910202.910199.910299		     
		     System.out.println("CheckOtherRule.bank_code="+bank_code);
		     System.out.println("910109="+tmp910109);
		     System.out.println("910201="+tmp910201);
			 System.out.println("910202="+tmp910202);
			 System.out.println("910203="+tmp910203);	
			 System.out.println("910204="+tmp910204);
			 System.out.println("910199="+tmp910199);
			 System.out.println("910299="+tmp910299);
			 System.out.println("910500="+tmp910500);
			 System.out.println("errCount="+errCount);
			 //95.05.29 fix 910500*1.25%->取到整數,小數點下以無條件捨去
			 tmp910500 = new BigDecimal(tmp910500.multiply(tmp125).intValue());//910500 * 1.25%
		     System.out.println("910500*1.25%="+tmp910500);
		     //【910204】最大為910500 * 1.25%
		     if(tmp910204.compareTo(tmp910500) == 1){//910204 > 910500 * 1.25%
		     	sqlCmd = "INSERT INTO WML03 VALUES (" +
		               m_year + "," + m_month + ",'" + bank_code + "','" + report_no + "'," +
			           errCount + ",'檢核發現「調整後備抵呆帳、損失準備及營業準備」【910204】["+Utility.setCommaFormat(tmp910204.toString())+"元]"
				       + " 不得超過「【910500】修正後之風險性資產總額(H) 」×1.25%["+Utility.setCommaFormat(tmp910500.toString())+"元]','"+user_id+"','"+user_name+"',sysdate)";
		     	errCount++ ;	
		     	updateDBSqlList.add(sqlCmd);
		     	return updateDBSqlList;
		     }else{//910204 <= 910500 * 1.25%	
		     	if(tmp910204.compareTo(tmp910500) == -1){//910204 < 910500 * 1.25%		     
		     	   tmpAmt_910202 = tmp910202.add(tmp910203.negate());//910202-910203
				   System.out.println("tmpAmt_910202(910202-910203)="+tmpAmt_910202);
		     	}else{//910204 = 910500 * 1.25%	
		     	   tmpAmt_910202 = tmp910204;
		     	   System.out.println("tmpAmt_910202(910204)="+tmpAmt_910202);
		     	}
		     	if(tmpAmt_910202.compareTo(tmpZero) == -1){//tmpAmt_910202 < 0 
				     tmpAmt_910202 = new BigDecimal("0");
				     System.out.println("setZero.tmpAmt_910202="+tmpAmt_910202);
				}
		     }
		     
		     //910109 > 0時,910204或910202-910203也要等於0
		     if(tmp910109.compareTo(tmpZero) == 1){//910109 > 0
		        if(tmpAmt_910202.compareTo(tmpZero) != 0){//tmpAmt_910202 != 0
		            System.out.println("errCount="+errCount);
		            sqlCmd = "INSERT INTO WML03 VALUES (" +
				           m_year + "," + m_month + ",'" + bank_code + "','" + report_no + "'," +
					       errCount + ",'檢核發現(910109)["+Utility.setCommaFormat(tmp910109.toString())+"]元 大於0,"
					       +"而(【910204】 或 【910202】-【910203】)金額為["+Utility.setCommaFormat(tmpAmt_910202.toString())+"]元,不為0','"+user_id+"','"+user_name+"',sysdate)";
		            errCount++ ;	
		            updateDBSqlList.add(sqlCmd);
		            return updateDBSqlList;
		        }
		     }
		     tmpAmt_910299 = tmp910201.add(tmpAmt_910202);//Temp_910299 = 【910201】+ Temp_910202
		     
		     System.out.println("-910203="+tmp910203.negate());
		     System.out.println("tmpAmt_910202(MIN[(910204 or 910202-910203):(910500*1.25%)])="+tmpAmt_910202);
		     System.out.println("tmpAmt_910299(910201+tmpAmt_910202)="+tmpAmt_910299);
		     System.out.println("tmpAmt_910299 compare tmp910199="+tmpAmt_910299.compareTo(tmp910199));

		     if(tmpAmt_910299.compareTo(tmp910199) == -1){//910201+Temp_910202 < 910199
		         if(tmp910199.compareTo(tmpZero) == -1 && tmp910299.compareTo(tmpZero) != 0){//910199 < 0 && 910299 != 0
		            System.out.println("errCount="+errCount);
		            sqlCmd = "INSERT INTO WML03 VALUES (" +
				           m_year + "," + m_month + ",'" + bank_code + "','" + report_no + "'," +
					       errCount + ",'第一類資本合計(910199)為負數,第二類資本(910299)應為0元','"+user_id+"','"+user_name+"',sysdate)";
		            errCount++ ;	
		            updateDBSqlList.add(sqlCmd);		            
		         }else if(tmp910199.compareTo(tmpZero) == -1 && tmp910299.compareTo(tmpZero) == 0){//910199 < 0 && 910299 = 0			             
		         }else if(tmp910299.compareTo(tmpAmt_910299) == 0){//910299 = tmp910299(910201+Temp_910202)
		         }else{
		             System.out.println("errCount="+errCount);
		             sqlCmd = "INSERT INTO WML03 VALUES (" +
					    	m_year + "," + m_month + ",'" + bank_code + "','" + report_no + "'," +
					    	errCount + ",'因為X值=『【910201】+MIN(【910204】：(【910500】-【910203】)*1.25%)』 ["+
					    	Utility.setCommaFormat(tmpAmt_910299.toString())+"元]小於 (910199)["+Utility.setCommaFormat(tmp910199.toString())+
					    	"元],檢核發現(910299)["+Utility.setCommaFormat(tmp910299.toString())+"元]不等於X值["+Utility.setCommaFormat(tmpAmt_910299.toString())+"]元','"+
					    	user_id+"','"+user_name+"',sysdate)";
		             errCount++ ;	
		             updateDBSqlList.add(sqlCmd);
		         }    
		     }else{//910201+Temp_910202 >= 910199
		         if(tmp910199.compareTo(tmpZero) == -1 && tmp910299.compareTo(tmpZero) != 0){//910199 < 0 && 910299 != 0
		             System.out.println("errCount="+errCount);
		             sqlCmd = "INSERT INTO WML03 VALUES (" +
		             		m_year + "," + m_month + ",'" + bank_code + "','" + report_no + "'," +
		             		errCount + ",'第一類資本合計(910199)為負數,第二類資本(910299)應為0元','"+user_id+"','"+user_name+"',sysdate)";
		             errCount++ ;	
		             updateDBSqlList.add(sqlCmd);
		             return updateDBSqlList;
		         }else if(tmp910199.compareTo(tmpZero) == -1 && tmp910299.compareTo(tmpZero) == 0){//910199 < 0 && 910299 = 0			             
		         }else if(tmp910299.compareTo(tmp910199) == 0){//910299 = 910199
		         }else{
		             System.out.println("errCount="+errCount);
		             sqlCmd = "INSERT INTO WML03 VALUES (" +
					    	m_year + "," + m_month + ",'" + bank_code + "','" + report_no + "'," +
					    	errCount + ",'因為X值=『【910201】+MIN(【910204】：(【910500】-【910203】)*1.25%)』 ["+
					    	Utility.setCommaFormat(tmpAmt_910299.toString())+"元]"+((tmpAmt_910299.compareTo(tmp910199) == 0)?"等於":"大於")+"(910199)["+Utility.setCommaFormat(tmp910199.toString())+
					    	"元],檢核發現(910299)["+Utility.setCommaFormat(tmp910299.toString())+"元]不等於(910199)["+Utility.setCommaFormat(tmp910199.toString())+"]元','"+
					    	user_id+"','"+user_name+"',sysdate)";
		             errCount++ ;	
		             updateDBSqlList.add(sqlCmd);
		         }
		    }
		    if(tmp910199.compareTo(tmpZero) == -1 && tmp910299.compareTo(tmpZero) != 0){//910199 < 0 && 910299 != 0
		       System.out.println("errCount="+errCount);
		       sqlCmd = "INSERT INTO WML03 VALUES (" +
            		  m_year + "," + m_month + ",'" + bank_code + "','" + report_no + "'," +
            		  errCount + ",'第一類資本合計(910199)為負數,第二類資本(910299)應為0元','"+user_id+"','"+user_name+"',sysdate)";
		       errCount++ ;	
		       updateDBSqlList.add(sqlCmd);
		    }
		}catch(Exception e){
			errMsg = errMsg + "UpdateA05.CheckOtherRule Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateA05.CheckOtherRule Error:"+e.getMessage());
		}
		logcalendar = Calendar.getInstance();
		nowlog = logcalendar.getTime();		
		System.out.println("CheckOtherRule end time="+logformat.format(nowlog));	
		return updateDBSqlList;
	}
	*/
	
	
	 //A01.A05跨表檢核
	/*
	private static List CheckRuleA01_A05(int errCount,String bank_type,String report_no,String m_year,String m_month,
			   							 String bank_code,String user_id,String user_name){
		String	sqlCmd	 = null;	
		String  amt="";		
		List updateDBSqlList = new LinkedList();		
		List dbData = null;
		Properties A01Data = new Properties();		
		Properties A05Data = new Properties();
		String[] A01acc_code = {"310100","310200","310300","310400","310500","310600","310800","320100+320200","130200                     "};
		String[] A05acc_code = {"910101","910102","910103","910104","910105","910106","910107","910108       ","910401+910402+910403+910404"};
		BigDecimal tmpAmtA01=null;		
		BigDecimal tmpAmtA05=null;
		StringTokenizer st = null;	    

		try{
			sqlCmd = " select * "
			       + " from a01"
				   + " where a01.m_year="+m_year 
				   + " and a01.m_month="+m_month
				   + " and a01.bank_code='"+bank_code+"'"
				   + " order by acc_code";			
			dbData = DBManager.QueryDB(sqlCmd,"amt");
			for(int i=0;i<dbData.size();i++){
			    A01Data.put((String)((DataObject)dbData.get(i)).getValue("acc_code"),(((DataObject)dbData.get(i)).getValue("amt")).toString());			    
			}
			sqlCmd = " select * "
			       + " from a05"
				   + " where m_year="+m_year 
				   + " and   m_month="+m_month
				   + " and   bank_code='"+bank_code+"'"
				   + " order by acc_code";			
			dbData = DBManager.QueryDB(sqlCmd,"amt");			
			for(int i=0;i<dbData.size();i++){
			    A05Data.put((String)((DataObject)dbData.get(i)).getValue("acc_code"),(((DataObject)dbData.get(i)).getValue("amt")).toString());			    
			}
			
			for(int i=0;i<A01acc_code.length;i++){
			    if(A01acc_code[i].trim().indexOf("+") == -1){//A01單一科目檢核			        
			       if(A05acc_code[i].trim().indexOf("+") == -1){//A05單一科目檢核
			          if(!((String)A01Data.get(A01acc_code[i].trim())).equals((String)A05Data.get(A05acc_code[i].trim()))){
			              sqlCmd = "INSERT INTO WML03 VALUES (" +
						           m_year + "," + m_month + ",'" + bank_code + "','" + report_no + "'," +
							       errCount + ",'A05("+A05acc_code[i].trim()+")["+Utility.setCommaFormat((String)A05Data.get(A05acc_code[i].trim()))+"]不等於A01("+A01acc_code[i].trim()+")["+Utility.setCommaFormat((String)A01Data.get(A01acc_code[i].trim()))+"],'"+user_id+"','"+user_name+"',sysdate)";
					      errCount++ ;	
					      updateDBSqlList.add(sqlCmd);
			          }
			       }else{//A01單一科目.A05多個科目
			           st = new StringTokenizer(A05acc_code[i].trim());
			           while (st.hasMoreTokens()) {
			  	          tmpAmtA05=tmpAmtA05.add(BigDecimal.valueOf(Long.parseLong((String)A05Data.get(st.nextToken()))));
			  	      }
			          if(!((String)A01Data.get(A01acc_code[i].trim())).equals(tmpAmtA05.toString())){
				              sqlCmd = "INSERT INTO WML03 VALUES (" +
							           m_year + "," + m_month + ",'" + bank_code + "','" + report_no + "'," +
								       errCount + ",'A05("+A05acc_code[i].trim()+")["+Utility.setCommaFormat(tmpAmtA05.toString())+"]不等於A01("+A01acc_code[i].trim()+")["+Utility.setCommaFormat((String)A01Data.get(A01acc_code[i].trim()))+"],'"+user_id+"','"+user_name+"',sysdate)";
						      errCount++ ;	
						      updateDBSqlList.add(sqlCmd);
				     }
			       }
			    }else{//A01多個科目
			        st = new StringTokenizer(A01acc_code[i].trim());
			        while (st.hasMoreTokens()) {
			  	          tmpAmtA01=tmpAmtA01.add(BigDecimal.valueOf(Long.parseLong((String)A01Data.get(st.nextToken()))));
			  	    } 
			        if(A05acc_code[i].trim().indexOf("+") == -1){//A01多個科目.A05單一科目檢核
				          if(!(tmpAmtA01.toString()).equals((String)A05Data.get(A05acc_code[i].trim()))){
				              sqlCmd = "INSERT INTO WML03 VALUES (" +
							           m_year + "," + m_month + ",'" + bank_code + "','" + report_no + "'," +
								       errCount + ",'A05("+A05acc_code[i].trim()+")["+Utility.setCommaFormat((String)A05Data.get(A05acc_code[i].trim()))+"]不等於A01("+A01acc_code[i].trim()+")["+Utility.setCommaFormat(tmpAmtA01.toString())+"],'"+user_id+"','"+user_name+"',sysdate)";
						      errCount++ ;	
						      updateDBSqlList.add(sqlCmd);
				          }
				    }else{//A01多個科目.A05多個科目
				           st = new StringTokenizer(A05acc_code[i].trim());
				           while (st.hasMoreTokens()) {
				  	          tmpAmtA05=tmpAmtA05.add(BigDecimal.valueOf(Long.parseLong((String)A05Data.get(st.nextToken()))));
				  	      }
				          if(!(tmpAmtA01.toString()).equals(tmpAmtA05.toString())){
					           sqlCmd = "INSERT INTO WML03 VALUES (" +
							           m_year + "," + m_month + ",'" + bank_code + "','" + report_no + "'," +
								       errCount + ",'A05("+A05acc_code[i].trim()+")["+Utility.setCommaFormat(tmpAmtA05.toString())+"]不等於A01("+A01acc_code[i].trim()+")["+Utility.setCommaFormat(tmpAmtA01.toString())+"],'"+user_id+"','"+user_name+"',sysdate)";
							   errCount++ ;	
							   updateDBSqlList.add(sqlCmd);
					     }
				       }
			        
			    }
			}
		}catch(Exception e){
			errMsg = errMsg + "UpdateA01.CheckRuleA01_A05 Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateA01.CheckRuleA01_A05 Error:"+e.getMessage());
		}
		return updateDBSqlList;		
	}
	*/	
}
