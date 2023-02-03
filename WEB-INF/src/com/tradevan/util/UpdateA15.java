//109.06.12 create A15電子銀行及行動支付業務辦理情形檔案上傳資料檢核(含共用中心上傳) by 2295
package com.tradevan.util;

import java.io.*;
import java.util.*;
import java.text.SimpleDateFormat;
import com.tradevan.util.dao.DataObject;

public class UpdateA15 {
	private static String errMsg = "";
	public String getErrMsg(){
		return errMsg;  
	}
	public static synchronized String doParserReport_A15(String report_no, String m_year,
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
		//99.11.11 add 查詢年度100年以前.縣市別不同===============================
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
		
		int depositaccount_tcnt=0;
		
		int total=0;
		try { 
		    
			if(input_method.equals("F")){//若為檔案上傳,先將檔案讀取出來
				List allList = getA15FileData(report_no,filename,m_year,m_month);						
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
			Utility.printLogTime("A15-1 time");	
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
			
			Utility.printLogTime("A15-2 time");	
			
			
			//批次寫入WML03_LOG(檢核其他錯誤)(區分檔案上傳/線上編輯)
			//104.02.11無檢核其他錯誤Utility.InsertWML03_LOG(m_year,m_month,user_id,user_name,report_no,filename,input_method,bank_codeList);
  		    //96.04.19 若為檔案上傳批次寫入A01 ZeroData / 有修改的A01寫入A01_log
			if(input_method.equals("F")){//在A12中無資料的先將insert Zero data 到A12
		    	errMsg += Utility.InsertZeroAXX_List(m_year,m_month,bank_codeList,report_no," acc_tr_type = 'A15' and acc_div='15' ");//在A15中無資料的先將insert Zero data 到A12
		    	Utility.InsertAXX_LOG(m_year,m_month,user_id,user_name,filename,report_no);//批次寫入(有異動)至A01_LOG		    			    	
		    	//99.11.11 A15沒有科目代號,不需檢核科目代號存不存在errMsg += Utility.InsertWML03(m_year,m_month,user_id,user_name,report_no,filename);//農漁會.批次寫入檢核其他錯誤(科目代號不存在)		    	
		    	errMsg += Utility.updateAXX(m_year,m_month,report_no,filename);//將有異動的資料update A15
			}//end of 檔案上傳
			//批次寫入WML02_LOG(檢核公式錯誤)(區分檔案上傳/線上編輯)
			//104.02.11 A15沒有檢核公式Utility.InsertWML02_LOG(m_year,m_month,user_id,user_name,report_no,input_method,filename,bank_codeList);
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
				Utility.printLogTime("A15-3 time");	
				//clear WML03(檢核其他錯誤)
				/*if(input_method.equals("F")){//檔案上傳
				    nonZero = false;
				    sqlCmd.delete(0,sqlCmd.length());
					paramList = new ArrayList();
					//存款帳戶總戶數(E)：與警示帳戶(A)、衍生管制帳戶(B)、自行篩選有異常並己採資金流出管制措施之存款帳戶(C)、其他帳戶(D)的戶數加總不符
					sqlCmd.append(" select * from ( ");
					sqlCmd.append(" select bank_code,warnaccount_cnt+limitaccount_cnt+erroraccount_cnt+otheraccount_cnt as total,depositaccount_tcnt");
					sqlCmd.append(" from A12 ");
					sqlCmd.append(" where m_year=? and m_month=?");
					sqlCmd.append(" and bank_code=?");
					sqlCmd.append(" )where total <> depositaccount_tcnt ");
					paramList.add(m_year);
					paramList.add(m_month);
					paramList.add((String)bank_codeList.get(i));
					dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"total,depositaccount_tcnt");
					
					depositaccount_tcnt = 0;
					total = 0;
					if(dbData != null && dbData.size() != 0){
						depositaccount_tcnt = Integer.parseInt((((DataObject)dbData.get(0)).getValue("depositaccount_tcnt")).toString());
						total = Integer.parseInt((((DataObject)dbData.get(0)).getValue("total")).toString());
						System.out.println("have wml03 error="+(String)bank_codeList.get(i));
						sqlCmd.delete(0,sqlCmd.length());
						//存款帳戶總戶數(E)：與警示帳戶(A)、衍生管制帳戶(B)、自行篩選有異常並己採資金流出管制措施之存款帳戶(C)、其他帳戶(D)的戶數加總不符
						sqlCmd.append("INSERT INTO WML03 VALUES (?,?,?,?,?,?,?,?,sysdate)");
						dataList =  new ArrayList();
						dataList.add(m_year);
						dataList.add(m_month);
						dataList.add((String)bank_codeList.get(i));
						dataList.add(report_no);
						dataList.add("0");//errCount
						dataList.add("存款帳戶總戶數(E)[" +Utility.setCommaFormat(String.valueOf(depositaccount_tcnt)) +"]:與警示帳戶(A)、衍生管制帳戶(B)、自行篩選有異常並己採資金流出管制措施之存款帳戶(C)、其他帳戶(D)的戶數"+ "["+Utility.setCommaFormat(String.valueOf(total))+"]加總不符");
						dataList.add(user_id);
						dataList.add(user_name);
						updateDBDataList.add(dataList);//1:傳內的參數List		
						updateDBSqlList.add(sqlCmd.toString());
						updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
						updateDBList.add(updateDBSqlList);
						errCount++;
					}
					Utility.printLogTime("A12-4 time");					
				}//end of input_method=="F"-->檔案上傳
				*/								
				nonZero = false;
				dbData = Utility.getCountZero(m_year,m_month,(String)bank_codeList.get(i),report_no,"(month_amt > 0 or year_amt > 0)");
			    /*99.11.19移至Utility.getCountZero讀取該申報資料all data 都為0的資料筆數
				sqlCmd.delete(0,sqlCmd.length());
				paramList = new ArrayList();
				
				//99.11.11檢核內容是否都為0
				sqlCmd.append(" select count(*) as countzero from "+report_no+" where m_year=? AND m_month=? AND bank_code=? ");
				sqlCmd.append(" and (warnaccount_cnt > 0 or limitaccount_cnt > 0 or erroraccount_cnt >0 or otheraccount_cnt >0 or depositaccount_tcnt >0)");
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
				if(input_method.equals("F")){//檔案上傳
					updateYearAmt1(m_year,m_month,(String)bank_codeList.get(i),user_id,user_name);//更新累計開戶數(自開辦以來)
					if(Integer.parseInt(m_month) > 1){//1月以後的本年累計數值
				       updateYearAmt2(m_year,m_month,(String)bank_codeList.get(i),user_id,user_name);//更新本年累計戶數及金額				   	
					}					
				}
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
				Utility.printLogTime("A15-5 time");	
				System.out.println("errCount="+errCount);
				upd_code = (errCount == 0)?"U":"E";//U檢核成功:E檢核失敗
				System.out.println("upd_code="+upd_code);
				checkResult = (errCount == 0)?"true":"false";//true檢核成功:false檢核失敗
				//94.03.28 fix Z檢核為0====================================
				if(!nonZero && errCount == 0){
					upd_code = "Z";//Z檢核為0
					checkResult="Z";//Z檢核為0
				}
				//========================================================
				System.out.println("upd_code="+upd_code);
				List getUpdateDBList = Utility.Insert_UpdateWML01(dbData,m_year, m_month,(String)bank_codeList.get(i),report_no,input_method,add_date,user_id,user_name,common_center,upd_method,upd_code,batch_no);//99.11.19 fix
				for(int j=0;j<getUpdateDBList.size();j++){
				    updateDBList.add(getUpdateDBList.get(j));
				}
				/*99.11.19移至Utility.Insert_UpdateWML01;wml01不存在Insert,存在時Update
				if(dbData.size() == 0){//WML01不存在時,Insert
					sqlCmd.delete(0,sqlCmd.length());
					updateDBDataList = new ArrayList();
					updateDBSqlList = new ArrayList();
					if(input_method.equals("F")){//檔案上傳
					   sqlCmd.append("INSERT INTO WML01 VALUES (?,?,?,?,?,?,?,to_date('"+add_date+"','YYYYMMDDHH24MISS'),?,?,?,?,null,?,?,sysdate)");
						
					   //99.11.11sqlCmd="INSERT INTO WML01 VALUES (" +m_year+","+m_month+",'"+(String)bank_codeList.get(i)+"','"+report_no+"','"+input_method+"','"+user_id+"','"+user_name+"',to_date('"+add_date+"','YYYYMMDDHH24MISS'),'"+common_center+"','"
					   //      + upd_method +"','"+upd_code+"',"+batch_no+",null,'"+user_id+"','"+user_name+"',sysdate)";
					}else{//線上編輯
					   sqlCmd.append("INSERT INTO WML01 VALUES (?,?,?,?,?,?,?,sysdate,?,?,?,?,null,?,?,sysdate)");
					   //99.11.11sqlCmd="INSERT INTO WML01 VALUES (" +m_year+","+m_month+",'"+(String)bank_codeList.get(i)+"','"+report_no+"','"+input_method+"','"+user_id+"','"+user_name+"',sysdate,'"+common_center+"','"
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
					//移至Utility.InsertWML01_LOG
				    //batch_no = Integer.parseInt((((DataObject)dbData.get(0)).getValue("batch_no")).toString()) + 1;
					sqlCmd.delete(0,sqlCmd.length());
					updateDBDataList = new ArrayList();
					updateDBSqlList = new ArrayList();
					sqlCmd.append(" UPDATE WML01 SET "); 
					sqlCmd.append(" upd_method=?,upd_code=?,batch_no=?,user_id=?,user_name=?,update_date=sysdate");   
					sqlCmd.append(" where m_year=? AND m_month=? AND bank_code=? AND report_no=?");
				
					//99.11.11sqlCmd=" UPDATE WML01 SET " 
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
				Utility.printLogTime("A15-6 time");	
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
			
			Utility.printLogTime("A15-7 time");		
			if(input_method.equals("F")){//檔案上傳
			    //將WML01_UPLOAD..filename對應的使用者帳號.姓名刪除			
			    //99.11.11updateDBSqlList.add("DELETE FROM WML01_UPLOAD where filename='"+filename+"'");
				//99.11.19刪除上傳檔案紀錄
				updateDBList.add(Utility.deleteWML01_UPLOAD(filename));//99.11.19
			    //99.11.11 刪除AXX_TMP暫存檔
			    Utility.deleteAXX_TMP(m_year,m_month,report_no,filename);
			}
			if(updateDBList != null && updateDBList.size()!=0){
				updateOK=DBManager.updateDB_ps(updateDBList);
			}
			
			System.out.println("update OK??"+updateOK);
			if(!updateOK){
				//parserResult=false;
				errMsg = errMsg + "UpdateA15.doParserReport_A15 UpdateDB Error:"+DBManager.getErrMsg()+"<br>";
				System.out.println(DBManager.getErrMsg());
			}
			if(input_method.equals("W"))/*線上編輯*/ errMsg = upd_code +":" + errMsg;//104.02.11
			System.out.println("errMsg="+errMsg);
			Utility.printLogTime("A15-8 time");	
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
			Utility.printLogTime("A15-9 time");	
		}catch (Exception e) {
			//parserResult=false;
			errMsg = errMsg + "UpdateA15.doParserReport_A15 Error:"+e.getMessage()+"<br>";
			System.out.println("UpdateA15.doParserReport_A15="+e.getMessage());
		}
		return errMsg;
		//return parserResult;
	}
	//讀取上傳檔案的資料
	//99.11.11 add 套用DAO.preparestatment,並列印轉換後的SQL 
	private static List getA15FileData(String report_no,String filename,String m_year,String m_month){
			String	txtline	 = null;			
			List AXXList = new LinkedList();
			List bank_codeList = new LinkedList();
			List allList = new LinkedList();
			List detail = null;
			List dbData = null;			
			String tmpAmt="";
			String year_amt="";
			//寫入AXX_TMP暫存table,使用preparestatement
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
				
				if(Integer.parseInt(m_month) != 1){
					
				}
				
				//98.11.11 寫入AXX_TMP暫存table
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
							detail.add(txtline.substring(12,18));//科目代號
							//本月開戶數/本月交易筆數或金額
							//94.03.07 fix 有負號("-")時,把之前的"0"去掉.. 
							tmpAmt = txtline.substring(18,32);//本月開戶數/本月交易筆數或金額
							if(tmpAmt.indexOf("-") != -1){
							   tmpAmt = tmpAmt.substring(tmpAmt.indexOf("-"),tmpAmt.length());
							}
							detail.add(tmpAmt);
							tmpAmt="";
							//累計開戶數/本年累計交易筆數或金額							
							tmpAmt = txtline.substring(32,46);//累計開戶數/本年累計交易筆數或金額		
							if(tmpAmt.indexOf("-") != -1){
							   tmpAmt = tmpAmt.substring(tmpAmt.indexOf("-"),tmpAmt.length());
							}
							detail.add(tmpAmt);
							detail.add("0");//amt2
							detail.add("0");//amt3
							detail.add("0");//amt4
							detail.add("0");//amt5
							detail.add("0");//amt6
							detail.add("0");//amt7
							detail.add("0");//amt8
							detail.add("0");//amt9
							detail.add(report_no);//99.11.09
							detail.add(filename);//99.11.09							
							updateDBDataList.add(detail);//1:傳內的參數List
							
							AXXList.add(detail);
						}
				}	
				in.close();
				f.close();
				
				//99.11.11 寫入AXX_TMP暫存table
				updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				    
				updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
				
				updateDBList.add(updateDBSqlList);            	
            	if(DBManager.updateDB_ps(updateDBList)){
            	   System.out.println("AXX_TMP Insert ok");				  	
            	}
			}catch(Exception e){
				errMsg = errMsg + "UpdateA15.getA15FileData Error:"+e.getMessage()+"<br>";
			}
			allList.add(bank_codeList);
			allList.add(AXXList);
			return allList;
	}
	
	private static int updateYearAmt2(String m_year,String m_month,String bank_code,String user_id,String user_name){//更新A15 本年累計交易戶數/金額		
		List dbData = null;	
		List paramList = new ArrayList();
		StringBuffer sqlCmd = new StringBuffer();	
		List updateDBList = new LinkedList();//0:sql 1:data		
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data
		DataObject bean = null;
		int updateOK = -1;
		try {	
			sqlCmd.append("	select count(*) as data from a15 where m_year=? and m_month=? and bank_code=?");
			paramList.add(m_year);
			paramList.add(Integer.parseInt(m_month)-1);
			paramList.add(bank_code);
			dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"data");
			if(dbData != null && Integer.parseInt(((DataObject)dbData.get(0)).getValue("data").toString()) > 0){//有上月份資料時,才更新累計數值 
				paramList = new ArrayList();
				sqlCmd.delete(0,sqlCmd.length());
				sqlCmd.append(" select now.bank_code,now.acc_code,now.month_amt,last.year_amt,month_amt+year_amt as sum_amt");
				sqlCmd.append(" from"); 
				sqlCmd.append(" (select bank_code,acc_code,month_amt");
				sqlCmd.append(" from a15");
				sqlCmd.append(" where m_year=? and m_month=?");
				sqlCmd.append(" and bank_code=?");
				sqlCmd.append(" and acc_code not in ('900100','900200','900300','900400'))now,");//當月份資料
				sqlCmd.append(" (select bank_code,acc_code,year_amt");
				sqlCmd.append(" from a15");
				sqlCmd.append(" where m_year=? and m_month=?");
				sqlCmd.append(" and bank_code=?");
				sqlCmd.append(" and acc_code not in ('900100','900200','900300','900400'))last");//上月份資料
				sqlCmd.append(" where now.bank_code=last.bank_code");
				sqlCmd.append(" and now.acc_code=last.acc_code");
				sqlCmd.append(" order by bank_code,acc_code");
			
				paramList.add(m_year);
				paramList.add(m_month);
				paramList.add(bank_code);
				paramList.add(m_year);
				paramList.add(Integer.parseInt(m_month)-1);
				paramList.add(bank_code);
				dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"month_amt,year_amt,sum_amt");
			
			
				if(dbData != null && dbData.size() > 0){
					System.out.println("dbData.size()="+dbData.size());
					for(int i=0;i<dbData.size();i++){
						dataList = new LinkedList();//儲存參數的data
						bean = (DataObject)dbData.get(i);
						dataList.add((bean.getValue("sum_amt")).toString());//year_amt
						dataList.add(user_id);
						dataList.add(user_name);
						dataList.add(m_year);//m_year
						dataList.add(m_month);//m_month
						dataList.add((String)bean.getValue("bank_code"));
						dataList.add((String)bean.getValue("acc_code"));
						updateDBDataList.add(dataList);//1:傳內的參數List					
					}
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" update a15");
					sqlCmd.append(" set year_amt=?,user_id=?,user_name=?,update_date=sysdate");
					sqlCmd.append(" where m_year=?");
					sqlCmd.append(" and m_month=?");
					sqlCmd.append(" and bank_code=?");
					sqlCmd.append(" and acc_code=?");				
				
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				    
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
				}	
			     	
				if(DBManager.updateDB_ps(updateDBList)){
					System.out.println("A15 Update Year_Amt2 ok");		
					updateOK = 0;
				}
			}
		}catch(Exception e){
			errMsg = errMsg + "UpdateA15.updateYearAmt2 Error:"+e.getMessage()+"<br>";			
		}
	    return updateOK;	
	}
	
	
	private static int updateYearAmt_1month(String m_year,String m_month,String bank_code,String user_id,String user_name){//更新A15-1月份本年累計交易戶數/金額		
		List dbData = null;	
		List paramList = new ArrayList();
		StringBuffer sqlCmd = new StringBuffer();	
		List updateDBList = new LinkedList();//0:sql 1:data		
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data
		DataObject bean = null;
		int updateOK = -1;
		try {	
			sqlCmd.append("	select count(*) as data from a15 where m_year=? and m_month=? and bank_code=?");
			paramList.add(m_year);
			paramList.add(Integer.parseInt(m_month)-1);
			paramList.add(bank_code);
			dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"data");
			if(dbData != null && Integer.parseInt(((DataObject)dbData.get(0)).getValue("data").toString()) > 0){//有上月份資料時,才更新累計數值 
				paramList = new ArrayList();
				sqlCmd.delete(0,sqlCmd.length());
				sqlCmd.append(" select now.bank_code,now.acc_code,now.month_amt,last.year_amt,month_amt+year_amt as sum_amt");
				sqlCmd.append(" from"); 
				sqlCmd.append(" (select bank_code,acc_code,month_amt");
				sqlCmd.append(" from a15");
				sqlCmd.append(" where m_year=? and m_month=?");
				sqlCmd.append(" and bank_code=?");
				sqlCmd.append(" and acc_code not in ('900100','900200','900300','900400'))now,");//當月份資料
				sqlCmd.append(" (select bank_code,acc_code,year_amt");
				sqlCmd.append(" from a15");
				sqlCmd.append(" where m_year=? and m_month=?");
				sqlCmd.append(" and bank_code=?");
				sqlCmd.append(" and acc_code not in ('900100','900200','900300','900400'))last");//上月份資料
				sqlCmd.append(" where now.bank_code=last.bank_code");
				sqlCmd.append(" and now.acc_code=last.acc_code");
				sqlCmd.append(" order by bank_code,acc_code");
			
				paramList.add(m_year);
				paramList.add(m_month);
				paramList.add(bank_code);
				paramList.add(m_year);
				paramList.add(Integer.parseInt(m_month)-1);
				paramList.add(bank_code);
				dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"month_amt,year_amt,sum_amt");
			
			
				if(dbData != null && dbData.size() > 0){
					System.out.println("dbData.size()="+dbData.size());
					for(int i=0;i<dbData.size();i++){
						dataList = new LinkedList();//儲存參數的data
						bean = (DataObject)dbData.get(i);
						dataList.add((bean.getValue("sum_amt")).toString());//year_amt
						dataList.add(user_id);
						dataList.add(user_name);
						dataList.add(m_year);//m_year
						dataList.add(m_month);//m_month
						dataList.add((String)bean.getValue("bank_code"));
						dataList.add((String)bean.getValue("acc_code"));
						updateDBDataList.add(dataList);//1:傳內的參數List					
					}
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" update a15");
					sqlCmd.append(" set year_amt=?,user_id=?,user_name=?,update_date=sysdate");
					sqlCmd.append(" where m_year=?");
					sqlCmd.append(" and m_month=?");
					sqlCmd.append(" and bank_code=?");
					sqlCmd.append(" and acc_code=?");				
				
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				    
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
				}	
			     	
				if(DBManager.updateDB_ps(updateDBList)){
					System.out.println("A15 Update Year_Amt2 ok");		
					updateOK = 0;
				}
			}
		}catch(Exception e){
			errMsg = errMsg + "UpdateA15.updateYearAmt2 Error:"+e.getMessage()+"<br>";			
		}
	    return updateOK;	
	}
	
	/*
	private static int updateYearAmt_1month(String m_year,String m_month,String bank_code,String user_id,String user_name){//更新A15 1月份的本年累計交易戶數/金額		
		List dbData = null;	
		List paramList = new ArrayList();
		StringBuffer sqlCmd = new StringBuffer();	
		List updateDBList = new LinkedList();//0:sql 1:data		
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data
		DataObject bean = null;
		int updateOK = -1;
		try {	
			
			sqlCmd.delete(0,sqlCmd.length());
			sqlCmd.append(" update a15");
			sqlCmd.append(" set year_amt=month_amt,user_id=?,user_name=?,update_date=sysdate");
			sqlCmd.append(" where m_year=?");
			sqlCmd.append(" and m_month=?");
			sqlCmd.append(" and bank_code=?");
			sqlCmd.append(" and acc_code not in ('900100','900200','900300','900400') ");				
			dataList = new LinkedList();//儲存參數的data
			updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				    
			updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
			updateDBList.add(updateDBSqlList);
	
	     	
		   if(DBManager.updateDB_ps(updateDBList)){
			  System.out.println("A15 Update Year_Amt 1 month ok");		
			  updateOK = 0;
		   }
			
			
			
			
		}catch(Exception e){
			errMsg = errMsg + "UpdateA15.updateYearAmt-1month Error:"+e.getMessage()+"<br>";			
		}
	    return updateOK;	
	}
	*/
	
	private static int updateYearAmt1(String m_year,String m_month,String bank_code,String user_id,String user_name){//更新A15 自開辦以來累計開戶數		
		List dbData = null;	
		List paramList = new ArrayList();
		StringBuffer sqlCmd = new StringBuffer();	
		List updateDBList = new LinkedList();//0:sql 1:data		
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data
		DataObject bean = null;
		int updateOK = -1;
		String last_year="";
		String last_month="";
		try {
			
			sqlCmd.append("	select count(*) as data from a15 where m_year=? and m_month=? and bank_code=?");
			if(Integer.parseInt(m_month) == 1){
				last_year = String.valueOf(Integer.parseInt(m_year)-1);
				last_month="12";				   
			}else{
				last_year=m_year;
				last_month=String.valueOf(Integer.parseInt(m_month)-1);				
			}
			paramList.add(last_year);
			paramList.add(last_month);
			paramList.add(bank_code);
			dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"data");
			if(dbData != null && Integer.parseInt(((DataObject)dbData.get(0)).getValue("data").toString()) > 0){//有上月份資料時,才更新累計數值
				paramList = new ArrayList();
				sqlCmd.delete(0,sqlCmd.length());
				sqlCmd.append(" select now.bank_code,now.acc_code,now.month_amt,last.year_amt,month_amt+year_amt as sum_amt");
				sqlCmd.append(" from"); 
				sqlCmd.append(" (select bank_code,acc_code,month_amt");
				sqlCmd.append(" from a15");
				sqlCmd.append(" where m_year=? and m_month=?");
				sqlCmd.append(" and bank_code=?");
				sqlCmd.append(" and acc_code in ('900100','900200','900300','900400'))now,");//當月份資料
				sqlCmd.append(" (select bank_code,acc_code,year_amt");
				sqlCmd.append(" from a15");
				sqlCmd.append(" where m_year=? and m_month=?");
				sqlCmd.append(" and bank_code=?");
				sqlCmd.append(" and acc_code in ('900100','900200','900300','900400'))last");//上月份資料
				sqlCmd.append(" where now.bank_code=last.bank_code");
				sqlCmd.append(" and now.acc_code=last.acc_code");
				sqlCmd.append(" order by bank_code,acc_code");
			
				paramList.add(m_year);
				paramList.add(m_month);
				paramList.add(bank_code);				
				paramList.add(last_year);
				paramList.add(last_month);
				paramList.add(bank_code);
				
				dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"month_amt,year_amt,sum_amt");
			
			
				if(dbData != null && dbData.size() > 0){
					System.out.println("dbData.size()="+dbData.size());
					for(int i=0;i<dbData.size();i++){
						dataList = new LinkedList();//儲存參數的data
						bean = (DataObject)dbData.get(i);
						dataList.add((bean.getValue("sum_amt")).toString());//year_amt
						dataList.add(user_id);
						dataList.add(user_name);
						dataList.add(m_year);//m_year
						dataList.add(m_month);//m_month
						dataList.add((String)bean.getValue("bank_code"));
						dataList.add((String)bean.getValue("acc_code"));
						updateDBDataList.add(dataList);//1:傳內的參數List					
					}
					sqlCmd.delete(0,sqlCmd.length());
					sqlCmd.append(" update a15");				
					sqlCmd.append(" set year_amt=?,user_id=?,user_name=?,update_date=sysdate");
					sqlCmd.append(" where m_year=?");
					sqlCmd.append(" and m_month=?");
					sqlCmd.append(" and bank_code=?");
					sqlCmd.append(" and acc_code=?");				
				
					updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				    
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
				}	
			     	
				if(DBManager.updateDB_ps(updateDBList)){
					System.out.println("A15 Update Year_Amt1 ok");		
					updateOK = 0;
				}
        	
			}
		}catch(Exception e){
			errMsg = errMsg + "UpdateA15.updateYearAmt1 Error:"+e.getMessage()+"<br>";			
		}
	    return updateOK;	
	}
	
}
