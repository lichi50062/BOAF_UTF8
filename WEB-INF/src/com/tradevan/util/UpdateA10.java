//97.06.13 A10應予評估資產彙總資料 create by 2295
//99.11.12 若為檔案上傳批次寫入A10 ZeroData / 有修改的A10寫入A10_log by 2295
//99.11.12 寫入暫存table AXX_TMP/有異動的資料寫入A10.使用preparestatement by 2295
//99.11.12 批次寫入WML01_Log by 2295
//99.11.19 add 移至共用Utility.getWML03_count讀取WML03 count(*)
//				      Utility.getCountZero讀取該申報資料all data 都為0的資料筆數
//					  Utility.getWML01讀取WML01 all data
//					  Utility.Insert_UpdateWML01當WML01不存在Insert,存在時Update
//					  Utility.deleteWML01_UPLOAD刪除上傳檔案紀錄 by 2295
//104.02.26 註2的帳列備抵呆帳需檢核->寫入WML03(*A10使用者目前無使用檔案上傳) by 2968
package com.tradevan.util;

import java.io.*;
import java.util.*;
import java.text.SimpleDateFormat;
import com.tradevan.util.dao.DataObject;

public class UpdateA10 {
	private static String errMsg = "";
	public String getErrMsg(){
		return errMsg;  
	}
	public static synchronized String doParserReport_A10(String report_no, String m_year,
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
		
		List AXXupdateDBSqlList = new LinkedList();
		List AXXList = new LinkedList();//AXX細部資料
		List bank_codeList = new LinkedList();//機構代碼list
		List ruleList = new LinkedList();//check rule後,欲Insert到WML02的sqlList		
		List dbData = null;//其他querydb後,資料暫存的list
		String checkResult="true"; 
		List acc_code = null;		
		boolean nonZero=false;//94.03.28 檢查AXX的內容是否都為"0"
		File tmpFile = new File(WMdataDir+filename);		
		Date filedate = new Date(tmpFile.lastModified());
		//99.11.12 add 查詢年度100年以前.縣市別不同===============================
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
		/*
		BigDecimal loan2_amt_d=null;//放款-列二類
		BigDecimal loan3_amt_d=null;//放款-列三類     
		BigDecimal loan4_amt_d=null;//放款-列四類 
		BigDecimal invest2_amt_d=null;//投資-列二類   
		BigDecimal invest3_amt_d=null;//投資-列三類   
		BigDecimal invest4_amt_d=null;//投資-列四類
		BigDecimal other2_amt_d=null;//其他-列二類    
		BigDecimal other3_amt_d=null;//其他-列三類    
		BigDecimal other4_amt_d=null;//其他-列四類
	    */
		
		try { 
			if(input_method.equals("F")){//若為檔案上傳,先將檔案讀取出來
				List allList = getA10FileData(report_no,filename,m_year,m_month);						
				bank_codeList = (List)allList.get(0);
				AXXList = (List)allList.get(1);
				
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
			Utility.printLogTime("A10-1 time");	
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
			Utility.printLogTime("A10-2 time");	
			
			//批次寫入WML03_LOG(檢核其他錯誤)
			//99.11.12無檢核其他錯誤Utility.InsertWML03_LOG(m_year,m_month,user_id,user_name,report_no,filename,input_method,bank_codeList);
  		    //96.04.19 若為檔案上傳批次寫入A10 ZeroData / 有修改的A10寫入A10_log
			if(input_method.equals("F")){//在A10中無資料的先將insert Zero data 到A10
		    	errMsg += Utility.InsertZeroAXX_List(m_year,m_month,bank_codeList,report_no,"");//在A10中無資料的先將insert Zero data 到A10
		    	Utility.InsertAXX_LOG(m_year,m_month,user_id,user_name,filename,report_no);//批次寫入(有異動)至A10_LOG		    			    	
		    	//99.11.12 A10沒有科目代號,不需檢核科目代號存不存在errMsg += Utility.InsertWML03(m_year,m_month,user_id,user_name,report_no,filename);//農漁會.批次寫入檢核其他錯誤(科目代號不存在)		    	
		    	errMsg += Utility.updateAXX(m_year,m_month,report_no,filename);//將有異動的資料update A10
			}//end of 檔案上傳
			//批次寫入WML02_LOG(檢核公式錯誤)(區分檔案上傳/線上編輯)
			//99.11.12 A10沒有檢核公式Utility.InsertWML02_LOG(m_year,m_month,user_id,user_name,report_no,input_method,filename,bank_codeList);
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
				Utility.printLogTime("A10-3 time");	
				//clear WML03(檢核其他錯誤)
				nonZero = false;
				sqlCmd.delete(0,sqlCmd.length());
				paramList = new ArrayList();
				
				//99.11.12檢核內容是否都為0
				dbData = Utility.getCountZero(m_year,m_month,(String)bank_codeList.get(i),report_no,"(loan1_amt >0 or loan2_amt > 0 or loan3_amt > 0 or loan4_amt > 0 "+                        
				                              " or invest1_amt > 0 or invest2_amt > 0 or invest3_amt > 0 or invest4_amt > 0 "+
				                              " or other1_amt > 0 or other2_amt > 0 or other3_amt > 0 or other4_amt > 0 "+
				                              " or loan1_baddebt > 0 or loan2_baddebt > 0 or loan3_baddebt > 0 or loan4_baddebt > 0 "+
				                              " or build1_baddebt > 0 or build2_baddebt > 0 or build3_baddebt > 0 or build4_baddebt > 0 )");//104.02.26 fix
			    /*99.11.19移至Utility.getCountZero讀取該申報資料all data 都為0的資料筆數				
				sqlCmd.append(" select count(*) as countzero from "+report_no+" where m_year=? AND m_month=? AND bank_code=? ");
				sqlCmd.append(" and (loan2_amt > 0 or loan3_amt > 0 or loan4_amt > 0 or invest2_amt > 0");                          
				sqlCmd.append("  or invest3_amt  > 0 or invest4_amt > 0 or other2_amt > 0 or other3_amt > 0 or other4_amt > 0)");

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
				System.out.println("nonZero="+nonZero);
				Utility.printLogTime("A10-4 time");
				if(input_method.equals("F")){//檔案上傳
					/* 99.11.12 移至Utility.InsertWML03_LOG 
					   99.11.12 移至Utility.InsertZeroAXX_List
					dbData = DBManager.QueryDB("select * from WML03 where m_year=" + m_year + " AND m_month=" + m_month + " AND " +
							"bank_code='" + (String)bank_codeList.get(i) + "' AND report_no='" + report_no + "'","m_year,m_month,serial_no,update_date");
					if(dbData.size() != 0){//WML03有資料時,先清掉
						sqlCmd = " INSERT INTO WML03_LOG " 
							   + " select m_year,m_month,bank_code,report_no,serial_no,remark,user_id,user_name,update_date"
							   + ",'"+user_id+"','"+user_name+"',sysdate,'D'"
							   + " from WML03"
							   + " where m_year=" + m_year + " AND m_month=" + m_month + " AND " 
							   + " bank_code='" + (String)bank_codeList.get(i) + "' AND report_no='" + report_no + "'";
						updateDBSqlList.add(sqlCmd);			
						sqlCmd = "DELETE FROM WML03 WHERE m_year=" + m_year + " AND m_month=" + m_month + " AND " +
							   "bank_code='" + (String)bank_codeList.get(i) + "' AND report_no='" + report_no + "'";
						updateDBSqlList.add(sqlCmd);
					}
					
					//若為檔案上傳時,若a10無資料,先將insert Zero data 到A10
			    	dbData = DBManager.QueryDB("select * from A10 where m_year=" + m_year + " AND m_month=" + m_month + " AND " +				
					 	     "bank_code='" + (String)bank_codeList.get(i) + "'","m_year,m_month");
				    if(dbData.size() == 0){//A10無資料時,才insert一筆Zero的資料
				    	InsertZeroA10(m_year,m_month,(String)bank_codeList.get(i));
				    }//end of A10不存在時
				    AXXupdateDBSqlList = new LinkedList();
				    nonZero = false;

				    BigDecimal tmpzero =  new BigDecimal("0");
					for(int j=0;j<AXXList.size();j++){//把上傳檔案裡的資料更新至A10
						System.out.println("AXXList="+(List)AXXList.get(j));
						System.out.println("bank_code="+(String)((List)AXXList.get(j)).get(0));
						//94.03.22 fix 改用List的方式來判斷科目代號存不存在
						if(((String)((List)AXXList.get(j)).get(0)).equals((String)bank_codeList.get(i))){//bank_code相同
							System.out.println("bank_code="+(String)((List)AXXList.get(j)).get(0));	
							
							loan2_amt_d = new BigDecimal((String)((List)AXXList.get(j)).get(1));//放款-列2類
							loan3_amt_d = new BigDecimal((String)((List)AXXList.get(j)).get(2));//放款-列3類
							loan4_amt_d = new BigDecimal((String)((List)AXXList.get(j)).get(3));//放款-列4類
							invest2_amt_d = new BigDecimal((String)((List)AXXList.get(j)).get(4));//投資-列2類
							invest3_amt_d = new BigDecimal((String)((List)AXXList.get(j)).get(5));//投資-列3類
							invest4_amt_d = new BigDecimal((String)((List)AXXList.get(j)).get(6));//投資-列4類
							other2_amt_d = new BigDecimal((String)((List)AXXList.get(j)).get(7));//其他-列2類
							other3_amt_d = new BigDecimal((String)((List)AXXList.get(j)).get(8));//其他-列3類
							other4_amt_d = new BigDecimal((String)((List)AXXList.get(j)).get(9));//其他-列4類
							
							System.out.println("loan2_amt_d="+loan2_amt_d);
							System.out.println("loan3_amt_d="+loan3_amt_d);
							System.out.println("loan4_amt_d="+loan4_amt_d);
							System.out.println("invest2_amt_d="+invest2_amt_d);
							System.out.println("invest3_amt_d="+invest3_amt_d);
							System.out.println("invest4_amt_d="+invest4_amt_d);
							System.out.println("other2_amt_d="+other2_amt_d);
							System.out.println("other3_amt_d="+other3_amt_d);
							System.out.println("other4_amt_d="+other4_amt_d);
							
							//94.03.28 add 檢核上傳檔案內容的金額是否都為"0"
							if(loan2_amt_d.compareTo(tmpzero) != 0 || 
							   loan3_amt_d.compareTo(tmpzero) != 0 || 
							   loan4_amt_d.compareTo(tmpzero) != 0 || 
							   invest2_amt_d.compareTo(tmpzero) != 0 || 
							   invest3_amt_d.compareTo(tmpzero) != 0 ||
							   invest4_amt_d.compareTo(tmpzero) != 0 ||
							   other2_amt_d.compareTo(tmpzero) != 0 || 
							   other3_amt_d.compareTo(tmpzero) != 0 ||
							   other4_amt_d.compareTo(tmpzero) != 0 
							){
								nonZero = true;	
							}
							
				    		sqlCmd = " INSERT INTO A10_LOG " 
								   + " select m_year,m_month,bank_code,loan2_amt,loan3_amt,loan4_amt,invest2_amt,invest3_amt,invest4_amt,other2_amt,other3_amt,other4_amt"
								   + ",'"+user_id+"','"+user_name+"',sysdate,'U'"
								   + " from A10"
								   + " where m_year=" +m_year+" and m_month="+m_month+" and bank_code='"+(String)bank_codeList.get(i)+"'";
				    		AXXupdateDBSqlList.add(sqlCmd);
				    						    		
				    		sqlCmd=" UPDATE A10 SET "
							 	   + " loan2_amt="+loan2_amt_d 
							       +" ,loan3_amt="+loan3_amt_d
							       +" ,loan4_amt="+loan4_amt_d
							       +" ,invest2_amt="+invest2_amt_d
							       +" ,invest3_amt="+invest3_amt_d
							       +" ,invest4_amt="+invest4_amt_d
							       +" ,other2_amt="+other2_amt_d
							       +" ,other3_amt="+other3_amt_d
							       +" ,other4_amt="+other4_amt_d
								   +" where m_year=" +m_year+" and m_month="+m_month+" and bank_code='"+(String)bank_codeList.get(i)+"'";
					        AXXupdateDBSqlList.add(sqlCmd);
								
						}//end of bank_code相同時			
					}//end of AXXList
					System.out.println("檔案上傳.nonZero="+nonZero);
					updateOK = DBManager.updateDB(AXXupdateDBSqlList);
					System.out.println("A10 update data OK??"+updateOK);
					if(!updateOK){
					   	//parserResult=false;
					   	errMsg = errMsg + "UpdateA10.doParserReport_A10 UpdateA10 Error:"+DBManager.getErrMsg()+"<br>";
					   	System.out.println(DBManager.getErrMsg());
					}
					*/
				}//end of input_method=="F"-->檔案上傳
				
						    
				if(input_method.equals("W")){//94.03.28 線上編輯check是否值全部為"0"
					/*
				   nonZero = false;
				   dbData = DBManager.QueryDB("select m_year,m_month,bank_code,loan2_amt,loan3_amt,loan4_amt,invest2_amt,invest3_amt,invest4_amt,other2_amt,other3_amt,other4_amt from A10 where m_year=" + m_year + " AND m_month=" + m_month + " AND " +				
					 	     			      "bank_code='" + (String)bank_codeList.get(i) + "'","m_year,m_month,loan2_amt,loan3_amt,loan4_amt,invest2_amt,invest3_amt,invest4_amt,other2_amt,other3_amt,other4_amt");
				   checkZeroLoop:
				   if(dbData != null && dbData.size() != 0){
					  for(int zeroIdx=0;zeroIdx < dbData.size();zeroIdx++){
					  	  if((Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("loan2_amt")).toString()) > 0)
					      ||(Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("loan3_amt")).toString()) > 0)
						  ||(Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("loan4_amt")).toString()) > 0)	
						  ||(Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("invest2_amt")).toString()) > 0)	
						  ||(Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("invest3_amt")).toString()) > 0)
						  ||(Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("invest4_amt")).toString()) > 0)
						  ||(Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("other2_amt")).toString()) > 0)	
						  ||(Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("other3_amt")).toString()) > 0)
						  ||(Long.parseLong((((DataObject)dbData.get(zeroIdx)).getValue("other4_amt")).toString()) > 0)
						  )	
					      {
						  	   nonZero = true;
						   	   break checkZeroLoop;
						  }
					  }
				   }
				   System.out.println("線上編輯.nonZero="+nonZero);
				   */
				}//end of 線上編輯check non zero	
				
				
				//註2的帳列備抵呆帳需檢核，寫入其他錯誤WML03
				if(nonZero){
					//執行公式檢核
				    ruleList = CheckOtherRule(errCount,bank_type,report_no,m_year,m_month,(String)bank_codeList.get(i),user_id,user_name,input_method);				    
					System.out.println("ruleList.size()="+ruleList.size());
				    if(ruleList.size() > 0){//有檢核失敗時,才加入
						sqlCmd.setLength(0);
				    	paramList.clear();
				    	sqlCmd.append("select bank_code from WML03 where m_year=? and m_month=? and bank_code=? and report_no=? ");  
			        	dataList = new ArrayList();//傳內的參數List		 	           				   
			        	paramList.add(m_year); 
			        	paramList.add(m_month); 
			        	paramList.add((String)bank_codeList.get(i)); 
			        	paramList.add(report_no);
				    	List WML03Data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"bank_code");
				    	if(WML03Data.size() != 0){//WML03有資料時,先清掉
							List WML03updateDBDataList = new ArrayList();//儲存參數的List
							sqlCmd.setLength(0);
							dataList.clear();
				            sqlCmd.append("INSERT INTO WML03_LOG ");  
				            sqlCmd.append("select m_year,m_month,bank_code,report_no,serial_no,remark,user_id,user_name,update_date,?,?,sysdate,'D' from WML03  "); 
				            sqlCmd.append("where m_year=? and m_month=? and bank_code=? and report_no=? ");  
				            dataList.add(user_id); 
							dataList.add(user_name);
							dataList.add(m_year); 
							dataList.add(m_month); 
							dataList.add((String)bank_codeList.get(i)); 
							dataList.add(report_no);
							WML03updateDBDataList.add(dataList);
							updateDBSqlList.add(sqlCmd.toString());
							updateDBSqlList.add(WML03updateDBDataList);
				            updateDBList.add(updateDBSqlList);
				            
				            WML03updateDBDataList.clear();
							updateDBSqlList.clear();
							sqlCmd.setLength(0);
							dataList.clear();
				            sqlCmd.append("delete WML03 where m_year=? and m_month=? and bank_code=? and report_no=? ");  
				        	dataList = new ArrayList();//傳內的參數List		 	           				   
							dataList.add(m_year); 
							dataList.add(m_month); 
							dataList.add((String)bank_codeList.get(i)); 
							dataList.add(report_no);
							WML03updateDBDataList.add(dataList);
							updateDBSqlList.add(sqlCmd.toString());
							updateDBSqlList.add(WML03updateDBDataList);
				            updateDBList.add(updateDBSqlList);
							
						}
						
						updateDBList.add((List)ruleList.get(0));
						errCount += ((List)((List)ruleList.get(0)).get(1)).size();//累計參數List的size
					}   
				}
				Utility.printLogTime("A10-5 time");	
				
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
				dbData =  Utility.getWML01(m_year,m_month,(String)bank_codeList.get(i),report_no);
				Utility.printLogTime("A10-6 time");
				
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
							
					   //99.11.12sqlCmd="INSERT INTO WML01 VALUES (" +m_year+","+m_month+",'"+(String)bank_codeList.get(i)+"','"+report_no+"','"+input_method+"','"+user_id+"','"+user_name+"',to_date('"+add_date+"','YYYYMMDDHH24MISS'),'"+common_center+"','"
					   //      + upd_method +"','"+upd_code+"',"+batch_no+",null,'"+user_id+"','"+user_name+"',sysdate)";
					}else{//線上編輯
					   sqlCmd.append("INSERT INTO WML01 VALUES (?,?,?,?,?,?,?,sysdate,?,?,?,?,null,?,?,sysdate)");
					   //99.11.12sqlCmd="INSERT INTO WML01 VALUES (" +m_year+","+m_month+",'"+(String)bank_codeList.get(i)+"','"+report_no+"','"+input_method+"','"+user_id+"','"+user_name+"',sysdate,'"+common_center+"','"
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
					/移至Utility.InsertWML01_LOG					
					//sqlCmd = " INSERT INTO WML01_LOG " 
					//	   + " select m_year,m_month,bank_code,report_no,input_method,add_user,add_name,add_date,common_center"
					//	   + ",upd_method,upd_code,batch_no,lock_status,user_id,user_name,update_date"
					//	   + ",'"+user_id+"','"+user_name+"',sysdate,'U'"
					//	   + " from WML01"
					//	   + " where m_year=" + m_year + " AND m_month=" + m_month + " AND " 						
					//	   + " bank_code='" + (String)bank_codeList.get(i) + "' AND report_no='" + report_no + "'";
				    //updateDBSqlList.add(sqlCmd);
				    
				    //batch_no = Integer.parseInt((((DataObject)dbData.get(0)).getValue("batch_no")).toString()) + 1;
					sqlCmd.append(" UPDATE WML01 SET "); 
					sqlCmd.append(" upd_method=?,upd_code=?,batch_no=?,user_id=?,user_name=?,update_date=sysdate");   
					sqlCmd.append(" where m_year=? AND m_month=? AND bank_code=? AND report_no=?");
					
					//99.11.12sqlCmd=" UPDATE WML01 SET " 
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
				Utility.printLogTime("A10-7 time");
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
			Utility.printLogTime("A10-8 time");			
			if(input_method.equals("F")){//檔案上傳
			    //將WML01_UPLOAD..filename對應的使用者帳號.姓名刪除			
			    //99.11.12updateDBSqlList.add("DELETE FROM WML01_UPLOAD where filename='"+filename+"'");
				//99.11.19刪除上傳檔案紀錄
				updateDBList.add(Utility.deleteWML01_UPLOAD(filename));//99.11.19
				//99.11.12 刪除AXX_TMP暫存檔
				Utility.deleteAXX_TMP(m_year,m_month,report_no,filename);
			}
			if(updateDBList != null && updateDBList.size()!=0){
				updateOK=DBManager.updateDB_ps(updateDBList);
			}
			System.out.println("update OK??"+updateOK);
			if(!updateOK){
				//parserResult=false;
				errMsg = errMsg + "UpdateA10.doParserReport_A10 UpdateDB Error:"+DBManager.getErrMsg()+"<br>";
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
			Utility.printLogTime("A10-9 time");	
		}catch (Exception e) {
			//parserResult=false;
			errMsg = errMsg + "UpdateA10.doParserReport_A10 Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateA10.doParserReport_A10="+e.getMessage());
		}
		return errMsg;
		//return parserResult;
	}
	//讀取上傳檔案的資料
    //99.11.12 add 套用DAO.preparestatment,並列印轉換後的SQL 
	private static List getA10FileData(String report_no,String filename,String m_year,String m_month){
			String	txtline	 = null;			
			List AXXList = new LinkedList();
			List bank_codeList = new LinkedList();
			List allList = new LinkedList();
			List detail = null;
			List dbData = null;	
			String tmpAmt="";
			String loan2_amt="";//放款-列二類
	  		String loan3_amt="";//放款-列三類     
	  		String loan4_amt="";//放款-列四類
	  		String invest2_amt="";//投資-列二類   
	  		String invest3_amt="";//投資-列三類   
	  		String invest4_amt="";//投資-列四類
	  		String other2_amt="";//其他-列二類    
	  		String other3_amt="";//其他-列三類    
	  		String other4_amt="";//其他-列四類    
	  	    //99.11.12 寫入AXX_TMP暫存table,使用preparestatement
			List paramList = new ArrayList();
			StringBuffer sqlCmd = new StringBuffer();	
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
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append("delete AXX_TMP where m_year=? and m_month=? and filename=?");
					updateDBDataList.add(paramList);//1:傳內的參數List
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				    
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
				}
				
				//99.11.12 寫入AXX_TMP暫存table
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
							}
							detail.add(txtline.substring(5,12));//bank_code	
							detail.add("000000");//acc_code(A10無acc_code)
							//94.03.07 fix 有負號("-")時,把之前的"0"去掉.. 
							//放款-列2類=================================================
							tmpAmt = txtline.substring(12,26);
							if(tmpAmt.indexOf("-") != -1){
							   tmpAmt = tmpAmt.substring(tmpAmt.indexOf("-"),tmpAmt.length());
							}
							loan2_amt=tmpAmt;
							//放款-列3類==============================================
							tmpAmt = txtline.substring(26,40);
							if(tmpAmt.indexOf("-") != -1){
							   tmpAmt = tmpAmt.substring(tmpAmt.indexOf("-"),tmpAmt.length());
							}
							loan3_amt=tmpAmt;
							//放款-列4類
							tmpAmt = txtline.substring(40,54);
							if(tmpAmt.indexOf("-") != -1){
							   tmpAmt = tmpAmt.substring(tmpAmt.indexOf("-"),tmpAmt.length());
							}
							loan4_amt=tmpAmt;
							//投資-列2類======================================================
							tmpAmt = txtline.substring(54,68);
							if(tmpAmt.indexOf("-") != -1){
							   tmpAmt = tmpAmt.substring(tmpAmt.indexOf("-"),tmpAmt.length());
							}
							invest2_amt=tmpAmt;
							//投資-列3類===================================================
							tmpAmt = txtline.substring(68,82);
							if(tmpAmt.indexOf("-") != -1){
							   tmpAmt = tmpAmt.substring(tmpAmt.indexOf("-"),tmpAmt.length());
							}
							invest3_amt=tmpAmt;					
							//投資-列4類===================================================
							tmpAmt = txtline.substring(82,96);
							if(tmpAmt.indexOf("-") != -1){
							   tmpAmt = tmpAmt.substring(tmpAmt.indexOf("-"),tmpAmt.length());
							}
							invest4_amt=tmpAmt;		
							//其他-列2類===================================================
							tmpAmt = txtline.substring(96,110);
							if(tmpAmt.indexOf("-") != -1){
							   tmpAmt = tmpAmt.substring(tmpAmt.indexOf("-"),tmpAmt.length());
							}
							other2_amt=tmpAmt;	
							//其他-列3類===================================================
							tmpAmt = txtline.substring(110,124);
							if(tmpAmt.indexOf("-") != -1){
							   tmpAmt = tmpAmt.substring(tmpAmt.indexOf("-"),tmpAmt.length());
							}
							other3_amt=tmpAmt;	
							//其他-列4類===================================================
							tmpAmt = txtline.substring(124,138);
							if(tmpAmt.indexOf("-") != -1){
							   tmpAmt = tmpAmt.substring(tmpAmt.indexOf("-"),tmpAmt.length());
							}
							other4_amt=tmpAmt;	
							//=====================================================================							
							detail.add(loan2_amt);//放款-列2類(axx_tmp.amt)
							detail.add(loan3_amt);//放款-列3類(axx_tmp.amt1)
							detail.add(loan4_amt);//放款-列4類(axx_tmp.amt2)
							detail.add(invest2_amt);//投資-列2類(axx_tmp.amt3)
							detail.add(invest3_amt);//投資-列3類(axx_tmp.amt4)
							detail.add(invest4_amt);//投資-列4類(axx_tmp.amt5)
							detail.add(other2_amt);//其他-列2類(axx_tmp.amt6)
							detail.add(other3_amt);//其他-列3類(axx_tmp.amt7)
							detail.add(other4_amt);//其他-列4類(axx_tmp.amt8)
							detail.add("0");//amt9
							detail.add(report_no);//99.11.12
							detail.add(filename);//99.11.12							
							updateDBDataList.add(detail);//1:傳內的參數List
							AXXList.add(detail);
						}
				}	
				in.close();
				f.close();
				//99.11.12 寫入AXX_TMP暫存table
				updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				    
				updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
				
				updateDBList.add(updateDBSqlList);            	
            	if(DBManager.updateDB_ps(updateDBList)){
            	   System.out.println("AXX_TMP Insert ok");				  	
            	}
			}catch(Exception e){
				errMsg = errMsg + "UpdateA10.getA10FileData Error:"+e.getMessage()+"<br>";
			}
			allList.add(bank_codeList);
			allList.add(AXXList);
			return allList;
	}
	// 99.11.15 add 註2的帳列備抵呆帳需檢核，寫入其他錯誤WML03
	private static List CheckOtherRule(int errCount,String bank_type,String report_no,String m_year,String m_month,
				   String bank_code,String user_id,String user_name,String input_method){
		    List dbData = null;	
		    StringBuffer sqlCmd = new StringBuffer();
			List paramList = new ArrayList();
			String wlx01_m_year = (Integer.parseInt(m_year) < 100)?"99":"100"; 
			//99.11.15 add 
			List updateDBList = new ArrayList();//0:sql 1:data		
			List updateDBSqlList = new ArrayList();
			List updateDBDataList = new ArrayList();//儲存參數的List
			List dataList =  new ArrayList();//儲存參數的data
		    try{
		    	sqlCmd.append("select loan1_baddebt,");//--B1
		    	sqlCmd.append("       build1_baddebt,");//--B6
		    	sqlCmd.append("       case when loan1_baddebt < build1_baddebt  then 1  else 0 end as loan1_baddebt_status,");//--B1<B6=1
		    	sqlCmd.append("       loan2_baddebt,");//--B2
		    	sqlCmd.append("       build2_baddebt,");//--B7
		    	sqlCmd.append("       case when loan2_baddebt < build2_baddebt  then 1  else 0 end as loan2_baddebt_status,");//--B2<B7=1
		    	sqlCmd.append("       loan3_baddebt,");//--B3
		    	sqlCmd.append("       build3_baddebt,");//--B8
		    	sqlCmd.append("       case when loan3_baddebt < build3_baddebt  then 1  else 0 end as loan3_baddebt_status,");//--B3<B8=1
		    	sqlCmd.append("       loan4_baddebt,");//--B4
		    	sqlCmd.append("       build4_baddebt,");//--B9
		    	sqlCmd.append("       case when loan4_baddebt < build4_baddebt  then 1  else 0 end as loan4_baddebt_status,");//--B4<B9=1
		    	sqlCmd.append("       loan1_baddebt+loan2_baddebt+loan3_baddebt+loan4_baddebt as loan_baddebt_sum,");//--B5
		    	sqlCmd.append("       field_backup,");//--A01.120800+A01.150300         
		    	sqlCmd.append("       case when loan1_baddebt+loan2_baddebt+loan3_baddebt+loan4_baddebt != field_backup then 1 else 0 end as field_backup_status ");//--B5 != A01.120800+A01.150300 then 1
		    	sqlCmd.append("  from a10 left join (select bank_code,sum(decode(acc_code,'120800',amt,'150300',amt,0)) as field_backup "); 
		    	sqlCmd.append("  from a01 ");
		    	sqlCmd.append("  where m_year=? ");
		    	sqlCmd.append("  and m_month=? ");
		    	sqlCmd.append("  group by bank_code ");
		    	sqlCmd.append("  )a01 on a10.bank_code = a01.bank_code ");
		    	sqlCmd.append("  where m_year=? ");
		    	sqlCmd.append("  and m_month=? ");
		    	sqlCmd.append("  and a10.bank_code=? ");
		    	paramList.add(m_year);
		    	paramList.add(m_month);
		    	paramList.add(m_year);
		    	paramList.add(m_month);
		    	paramList.add(bank_code);	  
		    	dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"loan1_baddebt,build1_baddebt,loan1_baddebt_status,"+
																				"loan2_baddebt,build2_baddebt,loan2_baddebt_status,"+
																				"loan3_baddebt,build3_baddebt,loan3_baddebt_status,"+
																				"loan4_baddebt,build4_baddebt,loan4_baddebt_status,"+
																				"loan_baddebt_sum,field_backup,field_backup_status");
		        
		    	
		    	if(dbData != null && dbData.size() != 0){
		            System.out.println("errCount="+errCount);
		            //a.不符合B1>B6時
		            String loan1_baddebt_status = ((DataObject)dbData.get(0)).getValue("loan1_baddebt_status").toString();
		            if("1".equals(loan1_baddebt_status)){
		            	dataList = new ArrayList();//傳內的參數List		 	           				   
						dataList.add(m_year); 
						dataList.add(m_month); 
						dataList.add(bank_code); 
						dataList.add(report_no); 
						dataList.add(String.valueOf(errCount)); 
						dataList.add("帳列備抵呆帳-第一類放款["+((DataObject)dbData.get(0)).getValue("loan1_baddebt").toString() +"元]必須大於帳列備抵呆帳-第一類建築貸款["+((DataObject)dbData.get(0)).getValue("build1_baddebt").toString()+"元]");					 
						dataList.add(user_id); 
						dataList.add(user_name); 
						updateDBDataList.add(dataList);//1:傳內的參數List
		                errCount++ ;	
		            }
		            //b.不符合B2>B7時
		            String loan2_baddebt_status = ((DataObject)dbData.get(0)).getValue("loan2_baddebt_status").toString();
		            if("1".equals(loan2_baddebt_status)){
		            	dataList = new ArrayList();//傳內的參數List		 	           				   
						dataList.add(m_year); 
						dataList.add(m_month); 
						dataList.add(bank_code); 
						dataList.add(report_no); 
						dataList.add(String.valueOf(errCount)); 
						dataList.add("帳列備抵呆帳-第二類放款["+((DataObject)dbData.get(0)).getValue("loan2_baddebt").toString()+"元]必須大於帳列備抵呆帳-第二類建築貸款["+((DataObject)dbData.get(0)).getValue("build2_baddebt").toString()+"元]");					 
						dataList.add(user_id); 
						dataList.add(user_name); 
						updateDBDataList.add(dataList);//1:傳內的參數List
		                errCount++ ;	
		            }
					//c.不符合B3>B8時
		            String loan3_baddebt_status = ((DataObject)dbData.get(0)).getValue("loan3_baddebt_status").toString();
		            if("1".equals(loan3_baddebt_status)){
		            	dataList = new ArrayList();//傳內的參數List		 	           				   
						dataList.add(m_year); 
						dataList.add(m_month); 
						dataList.add(bank_code); 
						dataList.add(report_no); 
						dataList.add(String.valueOf(errCount)); 
						dataList.add("帳列備抵呆帳-第三類放款["+((DataObject)dbData.get(0)).getValue("loan3_baddebt").toString()+"元]必須大於帳列備抵呆帳-第三類建築貸款["+((DataObject)dbData.get(0)).getValue("build3_baddebt").toString()+"元]");					 
						dataList.add(user_id); 
						dataList.add(user_name); 
						updateDBDataList.add(dataList);//1:傳內的參數List
		                errCount++ ;	
		            }
					//d.不符合B4>B9時
		            String loan4_baddebt_status = ((DataObject)dbData.get(0)).getValue("loan4_baddebt_status").toString();
		            if("1".equals(loan4_baddebt_status)){
		            	dataList = new ArrayList();//傳內的參數List		 	           				   
						dataList.add(m_year); 
						dataList.add(m_month); 
						dataList.add(bank_code); 
						dataList.add(report_no); 
						dataList.add(String.valueOf(errCount)); 
						dataList.add("帳列備抵呆帳-第四類放款["+((DataObject)dbData.get(0)).getValue("loan4_baddebt").toString()+"元]必須大於帳列備抵呆帳-第四類建築貸款["+((DataObject)dbData.get(0)).getValue("build4_baddebt").toString()+"元]");					 
						dataList.add(user_id); 
						dataList.add(user_name); 
						updateDBDataList.add(dataList);//1:傳內的參數List
		                errCount++ ;	
		            }
					//e.當B5不等於field_backup
		            String field_backup_status = ((DataObject)dbData.get(0)).getValue("field_backup_status").toString();
		            if("1".equals(field_backup_status)){
		            	dataList = new ArrayList();//傳內的參數List		 	           				   
						dataList.add(m_year); 
						dataList.add(m_month); 
						dataList.add(bank_code); 
						dataList.add(report_no); 
						dataList.add(String.valueOf(errCount)); 
						dataList.add("帳列備抵呆帳-放款合計["+((DataObject)dbData.get(0)).getValue("loan_baddebt_sum").toString()+"元]不等於A01資產負債表之(120800)「帳列備抵呆帳-放款」+(150300)「備抵呆帳-催收款項」合計["+((DataObject)dbData.get(0)).getValue("field_backup").toString()+"元]");					 
						dataList.add(user_id); 
						dataList.add(user_name); 
						updateDBDataList.add(dataList);//1:傳內的參數List
		                errCount++ ;	
		            }
		        }
		        
		        
		    	
		        if(updateDBDataList.size() > 0){//101.06.29當有檢核失敗時,才加入
					sqlCmd.setLength(0);
		            sqlCmd.append("INSERT INTO WML03 VALUES (?,?,?,?,?,?,?,?,sysdate)");             
		            updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql   
		            updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
		            updateDBList.add(updateDBSqlList);
		         }
		    }catch(Exception e){
		        errMsg = errMsg + "UpdateA10.CheckOtherRule Error:"+e.getMessage()+"<br>";
		        System.out.println("UpdateA10.CheckOtherRule Error:"+e.getMessage());
		    }	    
		    return updateDBList;
		}
}
