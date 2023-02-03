//94.01.04 fix 漁會檢核 by 2295
//94.01.04 找出該bank_code的bank_type by 2295
//94.03.02 金額改用BigDecimal去轉 by 2295
//94.03.07 fix 有負號("-")時,把之前的"0"去掉 by 2295
//94.03.15 add 若被鎖定,則不執行檢核,直接send mail通知 by 2295
//94.03.25 fix 若A01的值,都為"0"值時,顯示檢核為"0" by 2295
//95.02.07 add -10~10均算檢核成功 by 2295
//95.03.20 add 共用中心傳送彙總檢核結果e-mail by 2295
//95.03.27 fix 共用中心傳送彙總e-mail格式 by 2295
//95.06.01 fix 抓A01的公式代號為01/02開頭的 by 2295
//95.08.21 fix 若為備抵及累計折舊不可為負值 by 2295
//96.04.19 若為檔案上傳批次寫入A01 ZeroData / 有修改的A01寫入A01_log by 2295
//96.04.30 寫入暫存table AXX_TMP/有異動的資料寫入A01.使用preparestatement by 2295
//96.08.16 批次寫入WML01_Log/WML02_Log/WML03_Log/WML03 by 2295
//96.12.17 add 97/01以後的資料.套用新的科目代號 by 2295
//99.09.24-27 add InsertWML03套用DAO.preparestatment,並列印轉換後的SQL by 2295
//99.11.10 add AXX_TMP增加amt1~amt9 by 2295
//99.11.19 add 移至共用Utility.getWML03_count讀取WML03 count(*)
//					  Utility.getCountZero讀取該申報資料all data 都為0的資料筆數
//				      Utility.getWML01讀取WML01 all data
//				      Utility.Insert_UpdateWML01當WML01不存在Insert,存在時Update
//					  Utility.deleteWML01_UPLOAD刪除上傳檔案紀錄 by 2295
//102.04.18 add InsertWML03.103/01以後的資料.漁會套用新的科目代號 by 2295
package com.tradevan.util;

import java.io.*;
import java.util.*;
import java.text.SimpleDateFormat;
import com.tradevan.util.dao.DataObject;


public class UpdateA01 {
	private static String errMsg = "";
	public String getErrMsg(){
		return errMsg;  
	}
	public static synchronized String doParserReport_A01(String report_no, String m_year,
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
		
		List A01updateDBSqlList = new LinkedList();
		List A01List = new LinkedList();//A01細部資料
		List bank_codeList = new LinkedList();//機構代碼list
		List ruleList = new LinkedList();//check rule後,欲Insert到WML02的sqlList		
		List dbData = null;//其他querydb後,資料暫存的list
		String checkResult="true"; 
		//List acc_code_6 = null;
		//List acc_code_7 = null;
		boolean nonZero=false;//94.03.25 檢查A01的內容是否都為"0"
		File tmpFile = new File(WMdataDir+filename);		
		Date filedate = new Date(tmpFile.lastModified());
		//99.09.24 add 查詢年度100年以前.縣市別不同===============================
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
			
			Utility.printLogTime("A01 begin time");	
			if(input_method.equals("F")){//若為檔案上傳,先將檔案讀取出來
				List allList = getA01FileData(report_no,filename,m_year,m_month);//讀取檔案,並寫入AXX_TMP暫存table												
				bank_codeList = (List)allList.get(0);
				A01List = (List)allList.get(1);
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
			
			Utility.printLogTime("A01-1 time");	
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
			System.out.println("bank_codeList.size="+bank_codeList.size());
			Utility.printLogTime("A01-2 time");	
			//批次寫入WML03_LOG(檢核其他錯誤)(區分檔案上傳/線上編輯)
			Utility.InsertWML03_LOG(m_year,m_month,user_id,user_name,report_no,filename,input_method,bank_codeList);
  		    //96.04.19 若為檔案上傳批次寫入A01 ZeroData / 有修改的A01寫入A01_log
			if(input_method.equals("F")){//在A01中無資料的先將insert Zero data 到A01
		    	errMsg += Utility.InsertZeroAXX_List(m_year,m_month,bank_codeList,report_no," acc_tr_type = 'A01' and (acc_div='01' or (acc_div='02' and acc_code !='320300')) ");//在A01中無資料的先將insert Zero data 到A01
		    	Utility.InsertAXX_LOG(m_year,m_month,user_id,user_name,filename,report_no);//批次寫入(有異動)至A01_LOG		    			    	
		    	InsertWML03(m_year,m_month,user_id,user_name,report_no,"6",filename);//農會.批次寫入檢核其他錯誤
		    	InsertWML03(m_year,m_month,user_id,user_name,report_no,"7",filename);//漁會批次寫入檢核其他錯誤
		    	errMsg += Utility.updateAXX(m_year,m_month,report_no,filename);//將有異動的資料update A01
			}//end of 檔案上傳
			//批次寫入WML02_LOG(檢核公式錯誤)(區分檔案上傳/線上編輯)
			Utility.InsertWML02_LOG(m_year,m_month,user_id,user_name,report_no,input_method,filename,bank_codeList);
			//批次寫入WML01_LOG(檢核結果)(區分檔案上傳/線上編輯)
			Utility.InsertWML01_LOG(m_year,m_month,user_id,user_name,report_no,input_method,filename,bank_codeList);
			List lockBank_CodeList = Utility.getLockBank_Code(m_year,m_month,report_no);
			bank_codeLoop:
			for(int i=0;i<bank_codeList.size();i++){
				//94.03.15 add 若被鎖定,則不執行檢核,直接send mail通知==================================
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
				System.out.println("bank_codeList["+i+"]="+(String)bank_codeList.get(i));
				//94.01.04 找出該bank_code的bank_type by 2295=============================================
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
				Utility.printLogTime("A01-3 time");	
				
				errCount = 0;
				//clear WML03(檢核其他錯誤)
				/*96.04.25移至InsertWML03*/
				Utility.printLogTime("A01-4 time");	
				
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
				//99.09.24dbData = DBManager.QueryDB("select count(*) as countdata from WML03 where m_year="+m_year+" AND m_month="+m_month+" AND bank_code='"+(String)bank_codeList.get(i) +"' AND report_no='" + report_no + "'","countdata");
				
				if(Integer.parseInt((((DataObject)dbData.get(0)).getValue("countdata")).toString()) == 0){//96.08.13檢查WML03有無其他錯誤,無其他錯誤時,才檢核公式
				   System.out.println("bank_code="+(String)bank_codeList.get(i)+".WML03.size()="+(((DataObject)dbData.get(0)).getValue("countdata")).toString());
				   //if(errCount == 0){//無其他錯誤時,才檢核公式
					/*96.04.19 移至InsertWML02_LOG */				   
				    Utility.printLogTime("A01-5 time");	
				    
					nonZero = false;
					
					dbData = Utility.getCountZero(m_year,m_month,(String)bank_codeList.get(i),report_no,"amt > 0");//99.11.19 fix
					/*99.11.19移至Utility.getCountZero讀取該申報資料all data 都為0的資料筆數
					sqlCmd.delete(0,sqlCmd.length());
					paramList = new ArrayList();
					
					
					 if(input_method.equals("W")){
					 	sqlCmd.append("select count(*) as countzero from "+report_no+" where m_year=? AND m_month=? AND bank_code=? and amt > 0");
					 	paramList.add(m_year);
						paramList.add(m_month);
						paramList.add((String)bank_codeList.get(i));

						//dbData = DBManager.QueryDB("select count(*) as countzero from A01 where m_year=" + m_year + " AND m_month=" + m_month + " AND " +				
						// 	    			       "bank_code='" + (String)bank_codeList.get(i) + "' and amt > 0","countzero");
					}else{
						sqlCmd.append("select count(*) as countzero from axx_tmp where m_year=? AND m_month=? AND bank_code=? and report_no=? and amt > 0 and filename=?");
					 	paramList.add(m_year);
						paramList.add(m_month);
						paramList.add((String)bank_codeList.get(i));
						paramList.add(report_no);
						paramList.add(filename);						
						//dbData = DBManager.QueryDB("select count(*) as countzero from axx_tmp where m_year=" + m_year + " AND m_month=" + m_month + " AND " +				
						//						   "bank_code='" + (String)bank_codeList.get(i) + "' and report_no='"+report_no+"' and amt > 0 and filename='"+filename+"'","countzero");				    	
					}	
					dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"countzero");
					*/	
					if(dbData != null && dbData.size()==1){
					   System.out.println("nonZero.size()="+(((DataObject)dbData.get(0)).getValue("countzero")).toString());
					   if(Integer.parseInt((((DataObject)dbData.get(0)).getValue("countzero")).toString())>0){
					       nonZero = true;
					   }
					}   
				   	System.out.println("nonZero="+nonZero);			    		    	
				   
				   	Utility.printLogTime("A01-6 time");	
				    
					//96.05.02
				    if(nonZero){//94.03.25 fix 若A01的值,只要有非"0"值時,才執行檢核
				       //執行檢核
				       ruleList = CheckRule(bank_type,report_no,m_year,m_month,(String)bank_codeList.get(i),user_id,user_name);
				       if(ruleList.size() > 0){//99.09.27有檢核失敗時,才加入
				       //System.out.println("((List)ruleList.get(0)(0)).size()="+(String)((List)ruleList.get(0)).get(0));
				       //System.out.println("((List)ruleList.get(1)(1)).size()="+((List)((List)ruleList.get(0)).get(1)).size());
				       updateDBList.add((List)ruleList.get(0));
				       errCount += ((List)((List)ruleList.get(0)).get(1)).size();//99.09.27累計參數List的size
				       
				       //99.09.24for(int ruleidx=0;ruleidx<ruleList.size();ruleidx++){
				    	   //updateDBSqlList.add((String)ruleList.get(ruleidx));
				    	   //errCount ++;
				       //}
				       }
				       Utility.printLogTime("A01-7 time");
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
				//99.09.24dbData = DBManager.QueryDB("select * from WML01 where m_year=" + m_year + " AND m_month=" + m_month + " AND " +						
	 			//		 "bank_code='" + (String)bank_codeList.get(i) + "' AND report_no='" + report_no + "'","m_year,m_month,add_date,batch_no,update_date");
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
				//99.11.19移至Utility.Insert_UpdateWML01;wml01不存在Insert,存在時Update
				/*
				if(dbData.size() == 0){//WML01不存在時,Insert	
					sqlCmd.delete(0,sqlCmd.length());
					updateDBDataList = new ArrayList();
					updateDBSqlList = new ArrayList();
					if(input_method.equals("F")){//檔案上傳
						//dataList_InsertWML01.add("to_date('"+add_date+"','YYYYMMDDHH24MISS')");
						sqlCmd.append("INSERT INTO WML01 VALUES (?,?,?,?,?,?,?,to_date('"+add_date+"','YYYYMMDDHH24MISS'),?,?,?,?,null,?,?,sysdate)");
						
					    //99.09.24sqlCmd="INSERT INTO WML01 VALUES (" +m_year+","+m_month+",'"+(String)bank_codeList.get(i)+"','"+report_no+"','"++"','"+user_id+"','"+user_name+"',to_date('"+add_date+"','YYYYMMDDHH24MISS'),'"+common_center+"','"
					    //      + upd_method +"','"+upd_code+"',"+batch_no+",null,'"+user_id+"','"+user_name+"',sysdate)";
					}else{//線上編輯
						//dataList_InsertWML01.add("sysdate");						
						sqlCmd.append("INSERT INTO WML01 VALUES (?,?,?,?,?,?,?,sysdate,?,?,?,?,null,?,?,sysdate)");						
					    //99.09.24sqlCmd="INSERT INTO WML01 VALUES (" +m_year+","+m_month+",'"+(String)bank_codeList.get(i)+"','"+report_no+"','"+input_method+"','"+user_id+"','"+user_name+"',sysdate,'"+common_center+"','"
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
					//96.04.26移至批次寫入InsertWML01_LOG
				    //batch_no = Integer.parseInt((((DataObject)dbData.get(0)).getValue("batch_no")).toString()) + 1;
					sqlCmd.delete(0,sqlCmd.length());
					updateDBDataList = new ArrayList();
					updateDBSqlList = new ArrayList();
					sqlCmd.append(" UPDATE WML01 SET "); 
					sqlCmd.append(" upd_method=?,upd_code=?,batch_no=?,user_id=?,user_name=?,update_date=sysdate");   
					sqlCmd.append(" where m_year=? AND m_month=? AND bank_code=? AND report_no=?");
							
					//99.09.27sqlCmd=" UPDATE WML01 SET " 
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
			Utility.printLogTime("A01-8 time");	
				
			if(input_method.equals("F")){//檔案上傳
			   //將WML01_UPLOAD..filename對應的使用者帳號.姓名刪除			
			    //99.09.27updateDBSqlList.add("DELETE FROM WML01_UPLOAD where filename='"+filename+"'");
				updateDBList.add(Utility.deleteWML01_UPLOAD(filename));//99.11.19 fix
				/*99.11.19移至Utility.deleteWML01_UPLOAD
				sqlCmd.delete(0,sqlCmd.length());
				updateDBDataList = new ArrayList();
				updateDBSqlList = new ArrayList();
				sqlCmd.append(" DELETE FROM WML01_UPLOAD where filename=?"); 
				updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql	
				dataList = new ArrayList();//傳內的參數List
				dataList.add(filename);					
				updateDBDataList.add(dataList);//1:傳內的參數List
				updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
				updateDBList.add(updateDBSqlList);
				*/
			   //96.12.18 刪除AXX_TMP暫存檔
			   Utility.deleteAXX_TMP(m_year,m_month,report_no,filename);
			   
			}
			//updateDBSqlList.add("DELETE FROM AXX_TMP where filename='"+filename+"'");
			
			//99.09.27if(updateDBSqlList != null && updateDBSqlList.size()!=0){
			//	updateOK=DBManager.updateDB(updateDBSqlList);
			//}
			if(updateDBList != null && updateDBList.size()!=0){
				updateOK=DBManager.updateDB_ps(updateDBList);
			}
			System.out.println("update OK??"+updateOK);
			if(!updateOK){
				//parserResult=false;
				errMsg = errMsg + "UpdateA01.doParserReport_A01 UpdateDB Error:"+DBManager.getErrMsg()+"<br>";
				System.out.println(DBManager.getErrMsg());
			}
			
			if(input_method.equals("F")){//檔案上傳
			   String CopyResult = Utility.CopyFile(WMdataDir+System.getProperty("file.separator")+filename,WMdataBKDir+System.getProperty("file.separator")+ filename);
       		   if(CopyResult.equals("0")){//copy成功時,才將檔案刪除,避免使用rename造成的錯誤
          	   		tmpFile = new File(WMdataDir+System.getProperty("file.separator")+filename);
          	   		if(tmpFile.exists()) tmpFile.delete();              		
       		   }	   				
			}	
			Utility.printLogTime("A01-9 time");	
			
			if(common_center.equals("Y")){//95.03.20由共用中心傳入的,傳送彙總mail
			    if(input_method.equals("F")){//檔案上傳
			       Utility.sendMailNotification(srcbank_code,report_no,
				           m_year,m_month, checkResult,input_method,emailformat.format(filedate),filename,user_id,Utility.getSendMsg());			    
			    }else{//線上編輯
				   Utility.sendMailNotification(srcbank_code,report_no,
				           m_year,m_month, checkResult,input_method,Utility.getDateFormat("yyyy/MM/dd HH:mm:ss"),"",user_id,Utility.getSendMsg());
			    }   
			}	
			Utility.printLogTime("A01-10 time");	
			
		}catch (Exception e) {
			//parserResult=false;
			errMsg = errMsg + "UpdateA01.doParserReport_A01 Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateA01.doParserReport_A01="+e.getMessage());
		}
		return errMsg;
		//return parserResult;
	}
	//讀取上傳檔案的資料
 	//99.11.08 add 套用DAO.preparestatment,並列印轉換後的SQL 
	//99.11.10 add amt1~amt9 by 2295
	private static List getA01FileData(String report_no,String filename,String m_year,String m_month){
			String	txtline	 = null;			
			//List A01List = new LinkedList();
			List bank_codeList = new LinkedList();
			List allList = new LinkedList();
			List detail = null;
			List dbData = null;			
			String tmpAmt="";
			List paramList = new ArrayList();
			StringBuffer sqlCmd = new StringBuffer();	
			boolean updateOK=false;
			//96.04.30 寫入AXX_TMP暫存table,使用preparestatement
			List updateDBList = new LinkedList();//0:sql 1:data		
			List updateDBSqlList = new LinkedList();
			List updateDBDataList = new LinkedList();//儲存參數的List
			List dataList = new LinkedList();//儲存參數的data
			try {
				String WMdataDir = Utility.getProperties("WMdataDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");
				String WMdataBKDir = Utility.getProperties("WMdataBKDir")+System.getProperty("file.separator")+report_no+System.getProperty("file.separator");		
				File tmpFile = new File(WMdataDir+filename);		
				Date filedate = new Date(tmpFile.lastModified());
					
				FileReader	f		= new FileReader(tmpFile);
				LineNumberReader in	= new LineNumberReader(f);
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
				
				//96.08.13 寫入AXX_TMP暫存table
				sqlCmd.delete(0, sqlCmd.length());
				dataList =  new ArrayList();//儲存參數的data
				updateDBDataList = new ArrayList();//儲存參數的List	
				updateDBSqlList = new ArrayList();	
				
				sqlCmd.append("INSERT INTO AXX_TMP  VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,'',?,?)");
				
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
								//System.out.println("add bank_code list="+txtline.substring(5,12));
							}
							
							detail.add(txtline.substring(5,12));//bank_code
							detail.add(txtline.substring(12,18));//科目代號
							//94.03.07 fix 有負號("-")時,把之前的"0"去掉.. 
							tmpAmt = txtline.substring(18,32);//金額
							if(tmpAmt.indexOf("-") != -1){
							   tmpAmt = tmpAmt.substring(tmpAmt.indexOf("-"),tmpAmt.length());
							}
							//=========================================
							detail.add(tmpAmt);
							detail.add("0");//99.11.11 add amt1
							detail.add("0");//99.11.11 add amt2
							detail.add("0");//99.11.11 add amt3
							detail.add("0");//99.11.11 add amt4
							detail.add("0");//99.11.11 add amt5
							detail.add("0");//99.11.11 add amt6
							detail.add("0");//99.11.11 add amt7
							detail.add("0");//99.11.11 add amt8
							detail.add("0");//99.11.11 add amt9							
							detail.add(report_no);//96.04.30
							detail.add(filename);//96.04.30
							//A01List.add(detail);
							updateDBDataList.add(detail);//1:傳內的參數List
						}
						//System.out.println(detail);
				}	
				in.close();
				f.close();
				//96.04.30 寫入AXX_TMP暫存table
				updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				    
				updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
				updateDBList.add(updateDBSqlList);
            	if(DBManager.updateDB_ps(updateDBList)){
            	   System.out.println("AXX_TMP Insert ok");				  	
            	}
            	
			}catch(Exception e){
				errMsg = errMsg + "UpdateA01.getA01FileData Error:"+e.getMessage()+"<br>";
			}
			updateDBSqlList = null;
			//System.out.println("bank_codeList.size="+bank_codeList.size());
			//System.out.println("A01List.size="+A01List.size());
			allList.add(bank_codeList);
			allList.add(updateDBDataList);
			return allList;
	}
	
	private static double countAMT(double totalAmt,String noop,double amt){
		//double returnAmt=0.0; 
		//System.out.println("totalAmt="+totalAmt);
		try{
			//System.out.println("totalAmt="+totalAmt);
			//System.out.println("noop="+noop);
			//System.out.println("amt="+amt);
		    if (noop.equals("+")) {
		    	totalAmt += amt;
		    }else if (noop.equals("-")) {
		    	totalAmt -= amt;
		    }else if (noop.equals("*")) {
			    totalAmt *= amt;
		    }else if (noop.equals("/")) {
			   if (amt != 0) totalAmt /= amt;
		    }
		    //returnAmt=totalAmt;
		}catch(Exception e){
			System.out.println("countAMT Error:"+e.getMessage());
		}
		
		return totalAmt;
	}
	
	//公式檢核
    //95.06.01 fix 抓A01的公式代號為01/02開頭的 by 2295
	//99.09.24 add 套用DAO.preparestatment,並列印轉換後的SQL 
	private static List CheckRule(String bank_type,String report_no,String m_year,String m_month,
								 String bank_code,String user_id,String user_name){			
		String  amt="";		
		double	amt_L = 0.0, amt_R = 0.0,amt_tbl = 0.0;
		String	quop = null;
		String  cano="";
		String noop="";
			
		List dbData = null;		
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		//99.09.24 add 
		List updateDBList = new ArrayList();//0:sql 1:data		
		List updateDBSqlList = new ArrayList();
		List updateDBDataList = new ArrayList();//儲存參數的List
		List dataList =  new ArrayList();//儲存參數的data
		
		try{
			//check ruleno1, ruleno2
			sqlCmd.append(" select r1.cano,r1.quop,r2.left_flag,a01.amt,r2.noop,A01.acc_code ");
			sqlCmd.append(" from a01,ruleno1 r1,ruleno2 r2 ");
			sqlCmd.append(" where a01.acc_code=r2.acc_code ");			   
			sqlCmd.append(" and r1.CANO = r2.cano");
			sqlCmd.append(" and r1.acc_type = r2.acc_type");
			sqlCmd.append(" and r1.acc_type = ?");
			sqlCmd.append(" and (r1.CANO like '01%' or r1.CANO like '02%') ");//95.06.01 fix 抓A01的公式代號為01/02開頭的
			sqlCmd.append(" and a01.m_year=?"); 
			sqlCmd.append(" and a01.m_month=?");
			sqlCmd.append(" and a01.bank_code=?");
			sqlCmd.append(" order by r1.cano,r2.left_flag,r2.nserial");	
			paramList.add(bank_type);
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
			
			for(int r=0;r<dbData.size();r++){	
				//System.out.println("r="+r);
				if(!(((String)((DataObject)dbData.get(r)).getValue("cano")).trim()).equals(cano)){//與前一個公式不同時				    
				    //if ((quop.equals("=") && (amt_L == amt_R))){
				    if ((quop.equals("=") &&  ( Math.abs(amt_L - amt_R) <= 10 ))){//95.02.07 add -10~10均算檢核成功       
				         //System.out.println("amt_L="+amt_L);
					     //System.out.println("amt_R="+amt_R);
					     //System.out.println("amt_L - amt_R="+Math.abs(amt_L - amt_R));					
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
					quop = ((String)((DataObject)dbData.get(r)).getValue("quop") == null ? "" : ((String)((DataObject)dbData.get(0)).getValue("quop")).trim());					
					amt_L = 0.0; amt_R = 0.0;
					amt_tbl=0.0;
				}	

				amt_tbl = Double.parseDouble((((DataObject)dbData.get(r)).getValue("amt")).toString());
				//System.out.println("acc_code="+(String)((DataObject)dbData.get(r)).getValue("acc_code"));
				//System.out.println("noop="+noop);
				//System.out.println("amt_tbl="+amt_tbl);
				if(((String)((DataObject)dbData.get(r)).getValue("left_flag")).equals("0")){//左式
					/*
					if ((noop == null) || (noop.equals(""))) {
						amt_L = amt_tbl;
					}else{
						amt_L = countAMT(amt_L,noop,amt_tbl);
					}*/
					
					if ((noop == null) || (noop.equals(""))) {
						amt_L = amt_tbl;
					}else if (noop.equals("+")) {
						amt_L += amt_tbl;
					}else if (noop.equals("-")) {
						amt_L -= amt_tbl;
					}else if (noop.equals("*")) {
						amt_L *= amt_tbl;
					}else if (noop.equals("/")) {
						if (Double.parseDouble((((DataObject)dbData.get(r)).getValue("amt")).toString()) != 0)
							amt_L /= amt_tbl;
					}//end of noop
					
				}//end of left_flag=0-->左式金額加總	
				if(((String)((DataObject)dbData.get(r)).getValue("left_flag")).equals("1")){//右式
					/*
					if ((noop == null) || (noop.equals(""))) {
						amt_R = amt_tbl;
					}else{
						amt_R = countAMT(amt_R,noop,amt_tbl);
					}*/
					
					if ((noop == null) || (noop.equals(""))) {
						amt_R = amt_tbl;
					}else if (noop.equals("+")) {
						amt_R += amt_tbl;
					}else if (noop.equals("-")) {
						amt_R -= amt_tbl;
					}else if (noop.equals("*")) {
						amt_R *= amt_tbl;
					}else if (noop.equals("/")) {
						if (Double.parseDouble((((DataObject)dbData.get(r)).getValue("amt")).toString()) != 0)
							amt_R /= amt_tbl;
					}//end of noop
					
				}//end of left_flag=0-->右式金額加總
				//94.03.04 fix 把最後一筆的檢核失敗寫入db=================================
				if(r==dbData.size()-1){
				   //if ((quop.equals("=") && (amt_L == amt_R))){
				   if ((quop.equals("=") &&  ( Math.abs(amt_L - amt_R) <= 10 ))){//95.02.07 add -10~10均算檢核成功
				       //System.out.println("amt_L="+amt_L);
				       //System.out.println("amt_R="+amt_R);
				       //System.out.println("amt_L - amt_R="+Math.abs(amt_L - amt_R));
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
				noop = (((DataObject)dbData.get(r)).getValue("noop") == null ? "" : ((String)((DataObject)dbData.get(r)).getValue("noop")).trim());
				//System.out.println("r="+r+"end");				
			} //end of for()	
			System.out.println("WML02 have error.updateDBDataList.size()="+updateDBDataList.size());
			if(updateDBDataList.size() > 0){//99.09.27當有檢核失敗時,才加入
			   updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql	
			   updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
			   updateDBList.add(updateDBSqlList);
			}
			System.out.println("check rule end");
		}catch(Exception e){
			errMsg = errMsg + "UpdateA01.CheckRule Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateA01.CheckRule Error:"+e.getMessage());
		}
		return updateDBList;
	}
		
	//96.04.25 add 批次寫入WML03(檢核其他錯誤)
	//99.09.24 add 套用DAO.preparestatment,並列印轉換後的SQL 
	//102.04.18 add 103/01以後的資料.漁會套用新的科目代號 by 2295
	private static boolean InsertWML03(String m_year,String m_month,String user_id,String user_name,String report_no,String bank_type,String filename){
		boolean updateOK=false;
		String ncacno = "";	
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		//List updateDBSqlList = new LinkedList();
		Date nowlog = new Date();			
		SimpleDateFormat logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");
		Calendar logcalendar = Calendar.getInstance();
		nowlog = logcalendar.getTime();		
		logcalendar = Calendar.getInstance();
		nowlog = logcalendar.getTime();		
		System.out.println("InsertWML03 begin time="+logformat.format(nowlog));	
		List updateDBList = new LinkedList();//0:sql 1:data		
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		String wlx01_m_year = (Integer.parseInt(m_year) < 100)?"99":"100"; 	
		List dataList =  new ArrayList();//儲存參數的data
		try{
			/*
			select * from 
			(select * from (
			select m_year,m_month,bank_code,bank_type,report_no,axx_tmp.acc_code,ncacno.acc_name,amt
			from (select m_year,m_month,bank_code,bank_type,report_no,acc_code,amt
			from axx_tmp,(select bank_type,bank_no,bank_name from bn01 where m_year=99)bn01 
			where m_year=96 and m_month=2	 
			and report_no='A01'
			and amt < 0	
			and axx_tmp.bank_code=bn01.BANK_NO and bank_type='6')axx_tmp left join (select acc_code,acc_name from ncacno where acc_tr_type='A01')ncacno
			on axx_tmp.acc_code = ncacno.acc_code
			where (acc_name like '%備抵%' or acc_name like '%累計折舊%')
			order by m_year,m_month,bank_code)a1 
			union
			select * from 
			(select m_year,m_month,bank_code,bank_type,report_no,axx_tmp.acc_code,ncacno.acc_name,amt
			from (select m_year,m_month,bank_code,bank_type,report_no,acc_code,amt
			from axx_tmp,(select bank_type,bank_no,bank_name from bn01 where m_year=99)bn01	 
			where m_year=96 and m_month=2 
			and report_no='A01'
			and axx_tmp.bank_code=bn01.BANK_NO and bank_type='6'
			--and bank_code in ('6190196','6200167')
			)axx_tmp,(select acc_code,acc_name from ncacno where acc_tr_type='A01')ncacno
			where axx_tmp.acc_code = ncacno.acc_code(+)
			order by m_year,m_month,bank_code)where acc_name is null)
			order by m_year,m_month,bank_code,acc_name,acc_code
			 */
			
			ncacno = bank_type.equals("6")?"ncacno":"ncacno_7";
			
			if(bank_type.equals("6") && (Integer.parseInt(m_year) * 100 + Integer.parseInt(m_month) >= 9701)){//96.12.17 add 97/01以後的資料.農會套用新的科目代號 by 2295
		    	ncacno = "ncacno_rule";
		    	System.out.println("ncacno_rule.m_year="+m_year);
		    	System.out.println("ncacno_rule.m_year="+m_month);
		    }
			if(bank_type.equals("7") && (Integer.parseInt(m_year) * 100 + Integer.parseInt(m_month) >= 10301)){//102.04.18 add 103/01以後的資料.漁會套用新的科目代號 by 2295
                ncacno = "ncacno_7_rule";
                System.out.println("ncacno_7_rule.m_year="+m_year);
                System.out.println("ncacno_7_rule.m_year="+m_month);
            }
			//檢核其他錯誤
			sqlCmd.append(" select * from"); 
			sqlCmd.append(" (");
			sqlCmd.append("  select * from(");
			sqlCmd.append(" 		 select m_year,m_month,bank_code,bank_type,report_no,axx_tmp.acc_code,"+ncacno+".acc_name,amt");
			sqlCmd.append("       from (select m_year,m_month,bank_code,bank_type,report_no,acc_code,amt");
			sqlCmd.append("             from axx_tmp,(select bank_type,bank_no,bank_name from bn01 where m_year=?)bn01");
			paramList.add(wlx01_m_year);
			sqlCmd.append(" 			   where axx_tmp.m_year=? AND m_month=?");
			sqlCmd.append(" 			   and report_no=?");
			sqlCmd.append("             and filename=?");
			paramList.add(m_year);		
			paramList.add(m_month);
			paramList.add(report_no);
			paramList.add(filename);
			sqlCmd.append("             and amt <0");			
			sqlCmd.append(" 		       and axx_tmp.bank_code=bn01.BANK_NO and bank_type=?)axx_tmp ");
			paramList.add(bank_type);
			sqlCmd.append("			   left join (select acc_code,acc_name from "+ncacno+" where acc_tr_type='A01')"+ncacno);
			sqlCmd.append("				    on axx_tmp.acc_code = "+ncacno+".acc_code");
			sqlCmd.append("		       where (acc_name like '%備抵%' or acc_name like '%累計折舊%')");
			sqlCmd.append("			   order by m_year,m_month,bank_code )a1" );
			sqlCmd.append(" union ");
			sqlCmd.append(" select * from(" );
			sqlCmd.append("        select m_year,m_month,bank_code,bank_type,report_no,axx_tmp.acc_code,"+ncacno+".acc_name,amt");
			sqlCmd.append("		  from (select m_year,m_month,bank_code,bank_type,report_no,acc_code,amt");
			sqlCmd.append("				from axx_tmp,(select bank_type,bank_no,bank_name from bn01 where m_year=?)bn01" );
			paramList.add(wlx01_m_year);			 
			sqlCmd.append("              where m_year=? and m_month=?");
			sqlCmd.append(" 			 and report_no=?");
			sqlCmd.append("              and filename=?");
			paramList.add(m_year);		
			paramList.add(m_month);
			paramList.add(report_no);
			paramList.add(filename);
			sqlCmd.append("		        and axx_tmp.bank_code=bn01.BANK_NO and bank_type=?)axx_tmp," );
			paramList.add(bank_type);
			sqlCmd.append(" 			    (select acc_code,acc_name from "+ncacno+" where acc_tr_type='A01')"+ncacno);
			sqlCmd.append("		        where axx_tmp.acc_code = "+ncacno+".acc_code(+)");
			sqlCmd.append("				order by m_year,m_month,bank_code)where acc_name is null)");
			sqlCmd.append(" order by m_year,m_month,bank_code,acc_name,acc_code");	
			
			List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt");
			int errCount=-1;
			DataObject obj = null;
			String prebank_code="";			
			if(dbData != null && dbData.size() != 0){
			   prebank_code=(String)((DataObject) dbData.get(0)).getValue("bank_code");
			   for(int i=0;i<dbData.size();i++){
				   obj = (DataObject) dbData.get(i);
				   if(prebank_code.equals((String)obj.getValue("bank_code"))){
					  errCount++;
				   }else{
				      prebank_code = (String)obj.getValue("bank_code");
					  errCount=0;
				   }
				   sqlCmd.delete(0,sqlCmd.length());
				   sqlCmd.append("INSERT INTO WML03 VALUES (?,?,?,?,?,?,?,?,sysdate)");
				   if(obj.getValue("acc_name") != null){
				      //若為備抵及累計折舊不可為負值					  		
					  dataList = new ArrayList();//傳內的參數List		 	           				   
					  dataList.add(obj.getValue("m_year").toString()); 
					  dataList.add(obj.getValue("m_month").toString()); 
					  dataList.add((String)obj.getValue("bank_code"));   
					  dataList.add(report_no);  
					  dataList.add(String.valueOf(errCount));  
					  dataList.add("科目代號[" + (String)obj.getValue("acc_code") +"]["+ obj.getValue("amt").toString()+"]不可為負值");  
					  dataList.add(user_id);  
					  dataList.add(user_name);
					  updateDBDataList.add(dataList);//1:傳內的參數List
				   }else{
				      //無此科目代號
				   	  dataList = new ArrayList();//傳內的參數List		 	           				   
					  dataList.add(obj.getValue("m_year").toString()); 
					  dataList.add(obj.getValue("m_month").toString()); 
					  dataList.add((String)obj.getValue("bank_code"));   
					  dataList.add(report_no);  
					  dataList.add(String.valueOf(errCount));  
					  dataList.add("科目代號[" + (String)obj.getValue("acc_code") + "]不存在");  
					  dataList.add(user_id);  
					  dataList.add(user_name);
					  updateDBDataList.add(dataList);//1:傳內的參數List
				  }
			   }//end of for
			   updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql
			   updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
			   updateDBList.add(updateDBSqlList);        	
			   updateOK=DBManager.updateDB_ps(updateDBList);
			   System.out.println("Insert WML03 OK??"+updateOK);
			   if(!updateOK){
			      errMsg = errMsg + "Insert WML03 Fail:"+DBManager.getErrMsg()+"<br>";
			      System.out.println("Insert WML03 Fail:"+DBManager.getErrMsg());
			   }
			}//end of 有其他檢核錯誤時
		}catch(Exception e){
			errMsg = errMsg + "UpdateA01.InsertWML03 Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateA01.InsertWML03 Error:"+e.getMessage());
		}
		logcalendar = Calendar.getInstance();
		nowlog = logcalendar.getTime();		
		System.out.println("InsertWML03end time="+logformat.format(nowlog));	
    	return updateOK;
	}
	
    //96.05.02 先將公式寫入暫存table,ruleno_tmp 
	/*
	private static boolean InsertRuleno_tmp(String bank_type,String report_no,String m_year,String m_month,List bank_code,String filename)
	{
		boolean updateOK=false;
		String	sqlCmd	 = null;		
		List dataList = new LinkedList();
		List updateDBList = new LinkedList();//0:sql 1:data		
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		Date nowlog = new Date();			
		SimpleDateFormat logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");
		Calendar logcalendar = Calendar.getInstance();
		nowlog = logcalendar.getTime();		
		logcalendar = Calendar.getInstance();
		nowlog = logcalendar.getTime();		
		System.out.println("InsertRuleno_tmp begin time="+logformat.format(nowlog));	
		try{
			//check ruleno1, ruleno2
			
//			 Insert into ruleno_tmp
//			 select a01.m_year,a01.m_month,a01.bank_code,'A01' as report_no,r1.cano,r1.quop,r2.left_flag,a01.amt,r2.noop,r2.nserial,A01.acc_code,'filename'  
//	         from a01,ruleno1 r1,ruleno2 r2  
//		     where a01.acc_code=r2.acc_code  				   
//		     and r1.CANO = r2.cano 
//		     and r1.acc_type = r2.acc_type 
//		     and r1.acc_type = '6' 
//		     and (r1.CANO like '01%' or r1.CANO like '02%') 
//		     and a01.m_year= 96
//		     and a01.m_month= 2
//			 and a01.bank_code in (select bank_no
//			 						from bn01
//			 						where bank_no in('5120011')
//			 						and bank_type='6')
//		     order by r1.cano,r2.left_flag,r2.nserial
		   
			sqlCmd = " Insert into ruleno_tmp "
				   + " select a01.m_year,a01.m_month,a01.bank_code,? as report_no,r1.cano,r1.quop,"
				   + "        r2.left_flag,a01.amt,r2.noop,r2.nserial,A01.acc_code,? "  
				   + " from a01,ruleno1 r1,ruleno2 r2 "  
				   + " where a01.acc_code=r2.acc_code"  				   
				   + " and r1.CANO = r2.cano" 
				   + " and r1.acc_type = r2.acc_type" 
				   + " and r1.acc_type = ?"
				   + " and (r1.CANO like '01%' or r1.CANO like '02%') "//95.06.01 fix 抓A01的公式代號為01/02開頭的
				   + " and a01.m_year=?"
				   + " and a01.m_month=?"
			       + " and a01.bank_code in (select bank_no from bn01"
 				   + "  		             where bank_no in ("+Utility.getCombinString(bank_code,",")+")" 
 				   + "		                 and bank_type=? )"				   
		           + " order by r1.cano,r2.left_flag,r2.nserial";
			updateDBSqlList.add(sqlCmd);
			dataList = new LinkedList();
			dataList.add(report_no);
			dataList.add(filename);
			dataList.add(bank_type);
			dataList.add(m_year);
			dataList.add(m_month);	
			dataList.add(bank_type);
			updateDBDataList.add(dataList);
			updateDBSqlList.add(updateDBDataList);
			updateDBList.add(updateDBSqlList);
			updateOK = DBManager.updateDB_ps(updateDBList);
        	if(updateOK){
        	   System.out.println("bank_type="+bank_type+":ruleno_tmp Insert ok");			   	
        	}			
		}catch(Exception e){
			errMsg = errMsg + "UpdateA01.CheckRule Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateA01.CheckRule Error:"+e.getMessage());
		}
		logcalendar = Calendar.getInstance();
		nowlog = logcalendar.getTime();		
		System.out.println("InsertRuleno_tmp end time="+logformat.format(nowlog));	
    	return updateOK;
	}
	*/
	//96.05.02 批次公式檢核   
	/*
	private static List CheckRule_Batch(String m_year,String m_month,String report_no,
								 String filename,String user_id,String user_name){
		String	sqlCmd	 = null;
		double	amt_L = 0.0, amt_R = 0.0;
		String bank_code="";
		String	quop = null;
		String  cano="";
		String noop="";
		List updateDBSqlList = new LinkedList();		
		List dbData = null;
		List updateDBList_ps = new LinkedList();//0:sql 1:data		
		List updateDBSqlList_ps = new LinkedList();
		List updateDBDataList_ps = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();
		DataObject obj;
		Date nowlog = new Date();			
		SimpleDateFormat logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");
		Calendar logcalendar = Calendar.getInstance();
		nowlog = logcalendar.getTime();		
		logcalendar = Calendar.getInstance();
		nowlog = logcalendar.getTime();		
		System.out.println("CheckRule_Batch begin time="+logformat.format(nowlog));	
		try{
			//check ruleno1, ruleno2
//			select a1.bank_code,a1.cano,a1.amt as leftamt,quop,a2.cano as a2cano,a2.totalamt as rigntamt
//			 from (select bank_code,cano,amt,quop,nserial
//			 from ruleno_tmp
//			 where left_flag='0'
//			 and m_year=96 and m_month=2 and report_no='A01'
//			 and filename='1111111'
//			 )a1,
//			 (select cano,(beginamt+addamt-minusamt) as totalamt
//			 from (select cano
//		     ,sum(decode(nserial,'1',amt,0)) as beginamt
//			 ,sum(decode(noop,'+',nextamt,0)) as addamt
//			 ,sum(decode(noop,'-',nextamt,0)) as minusamt
//			 from ( select a1.cano,a1.nserial,a1.amt,a1.noop,a2.nserial as a2nserial,a2.amt as nextamt
//				    from (select cano,amt,noop,nserial
//				    	  from ruleno_tmp
//						  where left_flag='1'
//						  and m_year=96 and m_month=2 and report_no='A01'
//	  					  and filename='1111111'
//						  )a1 left join (select cano,amt,noop,nserial
//						 	   			 from ruleno_tmp
//										 where left_flag='1' and nserial > 1
//										 and m_year=96 and m_month=2 and report_no='A01'
//	  									 and filename='1111111'
//										 )a2
//									on a1.nserial+1=a2.nserial and a1.cano = a2.cano
//                  )
//			 group by cano
//           )
//           )a2
//           where a1.cano = a2.cano
//           order by bank_code,cano
			
			sqlCmd = " select a1.bank_code,a1.cano,a1.amt as leftamt,quop,a2.cano as a2cano,a2.totalamt as rigntamt"
				   + " from (select bank_code,cano,amt,quop,nserial "
			       + " 	     from ruleno_tmp "
			       + "       where left_flag='0'"
			       + "       and   m_year=? and m_month=? and report_no=? "
				   + "       and filename=? )a1,"
			       + "       (select cano,(beginamt+addamt-minusamt) as totalamt "
			       + "        from (select cano "
				   + " 	                   ,sum(decode(nserial,'1',amt,0)) as beginamt "
				   + "		               ,sum(decode(noop,'+',nextamt,0)) as addamt "
				   + "		               ,sum(decode(noop,'-',nextamt,0)) as minusamt "
				   + "		        from ( select a1.cano,a1.nserial,a1.amt,a1.noop,a2.nserial as a2nserial," 
				   + "                            a2.amt as nextamt"
				   + "			           from (select cano,amt,noop,nserial"
				   + "			    	         from ruleno_tmp"
				   + "					         where left_flag='1'"
				   + " 					         and m_year=? and m_month=? and report_no=? "
				   + " 					         and filename=? )a1 left join (select cano,amt,noop,nserial "
				   + "					 				     			       from ruleno_tmp "
				   + "								  				           where left_flag='1' and nserial > 1 "
				   + " 												           and m_year=? and m_month=? and report_no=? "
				   + " 													       and filename=? )a2"
				   + "									            on a1.nserial+1=a2.nserial and a1.cano = a2.cano"
			       + "			         )"
				   + "		 	    group by cano"
			       + "		      )"
				   + "	    )a2"
				   + " where a1.cano = a2.cano "
				   + " order by bank_code,cano ";
			
			updateDBSqlList_ps.add(sqlCmd);//0:欲執行的sql
			dataList.add(m_year);
			dataList.add(m_month);
			dataList.add(report_no);
			dataList.add(filename);
			dataList.add(m_year);
			dataList.add(m_month);
			dataList.add(report_no);
			dataList.add(filename);
			dataList.add(m_year);
			dataList.add(m_month);
			dataList.add(report_no);
			dataList.add(filename);			
			updateDBDataList_ps.add(dataList);
			updateDBSqlList_ps.add(updateDBDataList_ps);//1:參數List
			updateDBList_ps.add(updateDBSqlList_ps);//0:欲執行的sql/1:參數List
			dbData = DBManager.QueryDB_ps(updateDBList_ps,"leftamt,rigntamt");
			for(int r=0;r<dbData.size();r++){	
				amt_L = 0.0; amt_R = 0.0;
				obj = (DataObject) dbData.get(r);
				cano=(String)obj.getValue("cano");
				quop = (obj.getValue("quop") == null ? "" : ((String)obj.getValue("quop")).trim());
				amt_L = Double.parseDouble(obj.getValue("leftamt").toString());
				amt_R = Double.parseDouble(obj.getValue("rigntamt").toString());
				bank_code = (String)obj.getValue("bank_code");
				if ((quop.equals("=") &&  ( Math.abs(amt_L - amt_R) <= 10 ))){//95.02.07 add -10~10均算檢核成功       
			         System.out.println("amt_L="+amt_L);
				     System.out.println("amt_R="+amt_R);
				     System.out.println("amt_L - amt_R="+Math.abs(amt_L - amt_R));					
				} else if ((quop.equals(">") && (amt_L > amt_R))) {
				} else if ((quop.equals("<") && (amt_L < amt_R))) {
				} else if ((quop.equals(">=") && (amt_L >= amt_R))) {
				} else if ((quop.equals("<=") && (amt_L <= amt_R))) {
				} else if ((quop.equals("!=") && (amt_L != amt_R))) {
				}else {	    	
					System.out.println("WML02 have error");
					sqlCmd = "INSERT INTO WML02 VALUES (" + m_year + "," + m_month + ",'" + bank_code + "','" +
					         report_no+"','"+cano + "'," + amt_L + "," + amt_R + ",'"+user_id+"','"+user_name+"',sysdate)";
					System.out.println(sqlCmd);
					updateDBSqlList.add(sqlCmd);
				}
			}//end of for
		}catch(Exception e){
			errMsg = errMsg + "UpdateA01.CheckRule_Batch Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateA01.CheckRule_Batch Error:"+e.getMessage());
		}
		logcalendar = Calendar.getInstance();
		nowlog = logcalendar.getTime();		
		System.out.println("CheckRule_Batch end time="+logformat.format(nowlog));	
		return updateDBSqlList;
	}
	*/
}
