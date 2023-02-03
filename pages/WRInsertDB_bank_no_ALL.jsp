<%
/*
 * 103.03.18 created by 2968  
 *           1.配合103.02需求，結合WR00InsertDB_bank_no_ALL.jsp、WR01InsertDB_bank_no_ALL.jsp、WR02InsertDB_bank_no_ALL.jsp 
 *           2.fix (1)各項警訊指標皆調整為5名；
 *                 (2)最後一列增加顯示警示項目加總及各項數
 *                 (3)統計上月∕上季∕上年度同期，警訊指標皆相同，只有比較期間不同/顯示的title不同
 * 103.04.08 add專案農貸部份 
 * 103.06.10 add 翁小姐調整上年度信用部決算淨值公式 原A02.990230調整為A02.990230-A02.990240-A99.992810 by 2295
 * 103.06.10 add 翁小姐調整第8項指標,當上月份之(建築放款/上年度信用部決算淨值)低於100%,但本月份之(建築放款/上年度信用部決算淨值)高於100%,才列警示 by 2295
 * 103.06.17 fix 翁小姐調整本月份之(建築放款/上年度信用部決算淨值)-上月份之(建築放款/上年度信用部決算淨值)公式 by 2295 
 * 103.07.01 fix 專案農貸.過濾無資料的機構 by 2295
 * 103.07.30 fix 本月份資料SQL by 2295
 * 103.11.25 fix 專案農貸.過濾無資料的機構 by 2295
 */
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%@ page import="java.io.*" %>
<%@ page import="java.lang.Integer" %>
<html>
<head>
<title></title>
</head>
<body>
產生(農.漁會-各別機構)WR_operation
<%
	System.out.println("=============產生WR_Operation開始===========");
	Map dataMap = Utility.saveSearchParameter(request);
   	String report_no = Utility.getTrimString(dataMap.get("report_no"));	
   	String s_year = Utility.getTrimString(dataMap.get("s_year"));	
   	String s_month = Utility.getTrimString(dataMap.get("s_month"));	   
   	String isDebug = Utility.getTrimString(dataMap.get("isDebug"));	
   	String lguser_id = Utility.getTrimString(dataMap.get("lguser_id"));  
   	String wr_rpt = "";
   	/*警示報表類別  0:與上月比較 1:與上季比較 2:與上年度同期比較 
   	            3:與上月比較(專案農貸) 4:與上季比較(專案農貸)  
   	            5:與上年度同期比較(專案農貸)*/
   	if("WR00".equals(report_no)) wr_rpt = "0";
	else if("WR01".equals(report_no)) wr_rpt = "1";
	else if("WR02".equals(report_no)) wr_rpt = "2";
	else if("WR03".equals(report_no)) wr_rpt = "3";
	else if("WR04".equals(report_no)) wr_rpt = "4";
	else if("WR05".equals(report_no)) wr_rpt = "5";
   	String errMsg = Generate(report_no,s_year,s_month,isDebug,lguser_id,wr_rpt);
   	System.out.println("errMsg = "+errMsg);
   	System.out.println("=============產生WR"+report_no+"_Operation結束===========");      
%>

<%!
public String Generate(String report_no,String s_year,String s_month,String isDebug,String lguser_id,String wr_rpt) throws Exception{    	
	File logfile;
	FileOutputStream logos=null;    	
	BufferedOutputStream logbos = null;
	PrintStream logps = null;		
	File logDir = null;
	//==================================================================
	String cd01_table = "";
	String wlx01_m_year = "";
	String last_year = "";
	String last_month = "";
	String lastSeason_year = "";
	String lastSeason_month = "";
	StringBuffer sqlCmd = new StringBuffer();        	
	List<String> paramList = new ArrayList<String>();				
	List<List> updateDBList = new ArrayList<List>();//0:sql 1:data		
	List updateDBSqlList = new ArrayList();
	List<List> updateDBDataList = new ArrayList<List>();//儲存參數的List
	List<String> dataList =  new ArrayList<String>();//儲存參數的data
	//============================================================================
	String errMsg=""; 		
	DataObject bean = null;
	DataObject rangeBean = null;
	DataObject rangeTmpBean = null;
	DataObject AVGBean = null;
	DataObject operTmpRangeBean = null;
	try{      
      
		logDir  = new File(Utility.getProperties("logDir"));
	  	if(!logDir.exists()){
			if(!Utility.mkdirs(Utility.getProperties("logDir"))){
				System.out.println("目錄新增失敗");
			}    
       	}
       	
	   	logfile = new File(logDir + System.getProperty("file.separator") + report_no +"_GenerateOperation_WR."+ Utility.getDateFormat("yyyyMMddHHmmss"));						 
	   	System.out.println("logfile filename="+logDir + System.getProperty("file.separator") + report_no +"_GenerateOperation_WR."+ Utility.getDateFormat("yyyyMMddHHmmss"));
	   	logos = new FileOutputStream(logfile,true);  		        	   
	   	logbos = new BufferedOutputStream(logos);
	   	logps = new PrintStream(logbos);	
	   	logps.println(Utility.getDateFormat("yyyy/MM/dd  HH:mm:ss  ")+" "+"產生"+s_year+"年"+s_month+"月 (農.漁會-各別機構)WR00_opeation 與上月比較 資料中");		    					    
	   	logps.flush();
       
       	//99.05.07 add 查詢年度100年以前.縣市別不同===============================
  	   	cd01_table = (Integer.parseInt(s_year) < 100)?"cd01_99":"";
  	   	wlx01_m_year = (Integer.parseInt(s_year) < 100)?"99":"100";
  	    last_year  = (Integer.parseInt(s_month) == 1) ? String.valueOf(Integer.parseInt(s_year) - 1) : s_year; 
		last_month = (Integer.parseInt(s_month) == 1) ? "12" : String.valueOf(Integer.parseInt(s_month) - 1);
		lastSeason_year  = (Integer.parseInt(s_month) == 3) ? String.valueOf(Integer.parseInt(s_year) - 1) : s_year; 
  	    if(Integer.parseInt(s_month) <= 3 && Integer.parseInt(s_month)>=1){
  	    	lastSeason_month = "12";
  	    }else if(Integer.parseInt(s_month) <= 6 && Integer.parseInt(s_month)>=4){
  	      	lastSeason_month = "3";
  	    }else if(Integer.parseInt(s_month) <= 9 && Integer.parseInt(s_month)>=7){
  	      	lastSeason_month = "6";
  	    }else if(Integer.parseInt(s_month) <= 12 && Integer.parseInt(s_month)>=10){
  	      	lastSeason_month = "9";
  	    }
		//=====================================================================
		List dbData = new ArrayList();
		if("WR00".equals(report_no)){
		    dbData = getDetailInfo(s_year,s_month,last_year,last_month,wlx01_m_year,wr_rpt);
		}else if("WR01".equals(report_no)){
		    dbData = getDetailInfo(s_year,s_month,lastSeason_year,lastSeason_month,wlx01_m_year,wr_rpt);
		}else if("WR02".equals(report_no)){
		    dbData = getDetailInfo(s_year,s_month,String.valueOf(Integer.parseInt(s_year)-1),s_month,wlx01_m_year,wr_rpt);
		}else if("WR03".equals(report_no)){
		    dbData = getLoanDetailInfo(s_year,s_month,last_year,last_month,wlx01_m_year);
		}else if("WR04".equals(report_no)){
		    dbData = getLoanDetailInfo(s_year,s_month,lastSeason_year,lastSeason_month,wlx01_m_year);
		}else if("WR05".equals(report_no)){
		    dbData = getLoanDetailInfo(s_year,s_month,String.valueOf(Integer.parseInt(s_year)-1),s_month,wlx01_m_year);
		}
    	if(isDebug.equals("true"))  System.out.println("***********dbData.size()="+dbData.size());
    	if(dbData.size() > 0){
    	    if((getOperationTmpInfo(s_year,s_month,wr_rpt)).size()>0){
		        updateDBSqlList = new ArrayList();		
		    	updateDBDataList = new ArrayList();
		    	updateDBList = new ArrayList();	
		        sqlCmd.delete(0, sqlCmd.length());
		        dataList = new ArrayList();//傳內的參數List
		    	sqlCmd.append("delete WR_OPERATION_TMP where m_year=? and m_month=? and wr_rpt=? ");
		    	dataList.add(s_year); 
		    	dataList.add(s_month);
		    	dataList.add(wr_rpt);
		    	updateDBDataList.add(dataList);//1:傳內的參數List
		    	updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql    				
		    	updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
		    	updateDBList.add(updateDBSqlList);
		    	if(!DBManager.updateDB_ps(updateDBList)){					 
		    		errMsg = errMsg + "刪除資料庫失敗<br>[DBManager.getErrMsg()]:<br>";
		    	}else{
		    	    if(isDebug.equals("true"))  System.out.println("***********刪除資料庫成功");
		    	}
	    	}
		    if(haveOperationInfo(s_year,s_month,wr_rpt)){
		        updateDBSqlList = new ArrayList();		
		    	updateDBDataList = new ArrayList();
		    	updateDBList = new ArrayList();	
		        sqlCmd.delete(0, sqlCmd.length());
		        dataList = new ArrayList();//傳內的參數List
		    	sqlCmd.append("delete WR_OPERATION where m_year=? and m_month=? and wr_rpt=? ");
		    	dataList.add(s_year); 
		    	dataList.add(s_month); 
		    	dataList.add(wr_rpt);
		    	updateDBDataList.add(dataList);//1:傳內的參數List
		    	updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql    				
		    	updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
		    	updateDBList.add(updateDBSqlList);
		    	if(!DBManager.updateDB_ps(updateDBList)){					 
		    		errMsg = errMsg + "刪除資料庫失敗<br>[DBManager.getErrMsg()]:<br>";
		    	}else{
		    	    if(isDebug.equals("true"))  System.out.println("***********刪除資料庫成功");
		    	}
		    }
        	String[] field = {"field_debit_rate","field_credit_rate","field_990610_rate","field_noassure_rate","field_120700_rate",
        	         		  "field_992710","field_992710_rate","field_992710_credit_rate","field_992710_990230_rate","field_last_992710_990230_rate",
        	        		  "field_diff_992710_990230_rate","field_over","field_diff_over","field_over_rate","field_diff_over_rate",
        	        		  "field_diff_992530","field_992550_992150_rate","field_992720_992710_rate","field_992610_cal","field_diff_992610_cal",
        	        		  "field_992610_cal_credit_rate","field_diff_992630","field_diff_992730","field_diff_backup","field_backup_over_rate",
        	        		  "field_diff_backup_credit_rate"};
        	String[] field_type = { "4", "4", "4", "4", "4", 
		                   			"2", "4", "4", "4", "4",
		                   			"4", "2", "2", "4", "4",
		                   			"2", "4", "4", "2", "2",
		                   			"4", "2", "2", "2", "4",
		                   			"4"};
        	
        	String[] loanField = {"field_diff_loan_bal_amt","field_diff_over6m_loan_bal_amt","field_diff_over6m_loan_rate","field_delay_loan_rate"};
			String[] loanField_type = { "2", "2", "4", "4"};
			
		    String bank_code = ""; //機構代號
		    String acc_code = "";//科目代號
		    String type = "";//資料類別 0:純數字 2:加總 4:利率 
		    String amt = "";//金額
		    String tmp_amt = "";//暫存金額
		    String warn_type = "N";//Y:有警示 N:無警示
		    String field_code = "";
		    String quop = "";
	        String remark = "";
	        String wr_range_serial = "";
	        String field_amt = "";
	        //1.將上列資料寫入WR_OPERATION_TMP 警示報表暫存檔
	        updateDBSqlList = new ArrayList();		
	    	updateDBDataList = new ArrayList();
	    	updateDBList = new ArrayList();
	        sqlCmd.delete(0, sqlCmd.length());
	        sqlCmd.append("insert into WR_OPERATION_TMP(M_YEAR,M_MONTH,WR_RPT,BANK_CODE,ACC_CODE,TYPE,AMT,warn_type) values(?,?,?,?,?,?,?,?)");
			updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql 
		    for(int i=0;i<dbData.size();i++){
		    	bean = (DataObject)dbData.get(i);
		    	bank_code = Utility.getTrimString(bean.getValue("bank_no"));
		    	if("WR00".equals(report_no)||"WR01".equals(report_no)||"WR02".equals(report_no)){
			        for(int j=0;j<field.length;j++){
					    /**********1.將上列資料寫入WR_OPERATION_TMP (警示報表暫存檔)**********/
					    boolean isIns = true;
					    acc_code = field[j];
					    type = field_type[j];
					    amt = (bean.getValue(acc_code)==null)?"":(bean.getValue(acc_code)).toString();//金額
					    if("field_120700_rate".equals(acc_code) || "field_992710".equals(acc_code) || "field_992710_credit_rate".equals(acc_code)
					           || "field_over".equals(acc_code) || "field_over_rate".equals(acc_code) || "field_992550_992150_rate".equals(acc_code)
					           || "field_992610_cal".equals(acc_code) || "field_992610_cal_credit_rate".equals(acc_code) || "field_backup_over_rate".equals(acc_code)){
					        isIns = false;
					    }
					    if(isIns){
					        dataList = new ArrayList();//傳內的參數List
						    dataList.add(s_year); 
						    dataList.add(s_month); 
						    dataList.add(wr_rpt); 
						    dataList.add(bank_code);
						    dataList.add(acc_code); 
						    dataList.add(type); 
						    //103.06.10 add 當上月份之(建築放款/上年度信用部決算淨值)低於100%,但本月份之(建築放款/上年度信用部決算淨值)高於100%,才列警示
						    if("field_992710_990230_rate".equals(acc_code)){
						    	tmp_amt = "";
						    	tmp_amt = (bean.getValue("field_last_992710_990230_rate")==null)?"":(bean.getValue("field_last_992710_990230_rate")).toString();//金額
						    	if(Double.parseDouble(tmp_amt) < 100 && Double.parseDouble(amt) > 100){
						    		dataList.add(amt);
							        dataList.add("Y");
							        //System.out.println("last="+tmp_amt+":now="+amt);
						        }else{
						        	dataList.add(amt);
							        dataList.add("N");
						        }		
						    }else{	
					        //if("field_992710_990230_warn_type".equals(acc_code)){
					        //    dataList.add("0");
							//    dataList.add(amt);
					        //}else{
					            dataList.add(amt);
							    dataList.add("N");
					        }
						    updateDBDataList.add(dataList);//1:傳內的參數List
					    }
			        }
		    	}else if("WR03".equals(report_no)||"WR04".equals(report_no)||"WR05".equals(report_no)){
		    	    System.out.println(report_no+"loanField.length="+loanField.length);
		    	    for(int j=0;j<loanField.length;j++){
					    /**********1.將上列資料寫入WR_OPERATION_TMP (警示報表暫存檔)**********/
					    acc_code = loanField[j];
					    type = loanField_type[j];
					    amt = (bean.getValue(acc_code)==null)?"":(bean.getValue(acc_code)).toString();//金額
					    dataList = new ArrayList();//傳內的參數List
						dataList.add(s_year); 
						dataList.add(s_month); 
						dataList.add(wr_rpt); 
						dataList.add(bank_code);
						dataList.add(acc_code); 
						dataList.add(type); 
					    dataList.add(amt);
						dataList.add("N");
						updateDBDataList.add(dataList);//1:傳內的參數List
			        }
		    	}
		        logps.println(Utility.getDateFormat("yyyy/MM/dd  HH:mm:ss  ")+" "+"機構代號:"+bank_code);
		    }
		    updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
		    updateDBList.add(updateDBSqlList);
		    if(!DBManager.updateDB_ps(updateDBList)){					 
				errMsg = errMsg + "相關資料寫入 WR_OPERATION_TMP(警示報表暫存檔) 失敗<br>[DBManager.getErrMsg()]:<br>";
			}else{
				if(isDebug.equals("true"))  System.out.println("***********相關資料寫入 WR_OPERATION_TMP(警示報表暫存檔) 成功");
			}
			
			//2.取得與上月比較警示報表範圍
			updateDBSqlList = new ArrayList();		
	    	updateDBDataList = new ArrayList();
	    	updateDBList = new ArrayList();
		    List operTmpInfo = getOperationTmpInfo(s_year,s_month,wr_rpt);
	  		String total_count = String.valueOf(operTmpInfo.size());
		  	sqlCmd.delete(0, sqlCmd.length());
			sqlCmd.append("update WR_OPERATION_TMP set WARN_TYPE=?, REMARK=?, WR_RANGE_SERIAL=? ");
			sqlCmd.append(" where m_year=? and m_month=? and wr_rpt=? and bank_code=? and acc_code=? ");
			updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql 
		  	List rangeInfo = getRangeInfo(wr_rpt);
		  	for(int r=0;r<rangeInfo.size();r++){
		  		rangeBean = (DataObject)rangeInfo.get(r);
		  		String rangeBean_acc_code = rangeBean.getValue("acc_code").toString();
		  		String rangeBean_remark = (rangeBean.getValue("remark")).toString();
		  		String rangeBean_wr_range_serial = (rangeBean.getValue("serial")).toString();
		  		String rangeBean_field_amt = (rangeBean.getValue("field_amt")==null)?"":(rangeBean.getValue("field_amt")).toString();
		  		String rangeBean_quop = (rangeBean.getValue("quop")).toString();
		  		if("WR00".equals(report_no)||"WR01".equals(report_no)||"WR02".equals(report_no)){//警示報表
			  		if(!("6".equals(rangeBean_wr_range_serial) || "8".equals(rangeBean_wr_range_serial)  
				  		     || "9".equals(rangeBean_wr_range_serial) || "10".equals(rangeBean_wr_range_serial) 
				  		     || "12".equals(rangeBean_wr_range_serial) || "15".equals(rangeBean_wr_range_serial)
				  		     || "17".equals(rangeBean_wr_range_serial) || "22".equals(rangeBean_wr_range_serial)
				  		     || "24".equals(rangeBean_wr_range_serial))){
				  		//取出符合該range的bank
				  		if(!"max".equals(rangeBean_quop) && !"min".equals(rangeBean_quop)){
				  		    String tmpStr = rangeBean_field_amt.substring(0, 1);
				  		    if(!("0".equals(tmpStr)||"1".equals(tmpStr)||"2".equals(tmpStr)||"3".equals(tmpStr)||"4".equals(tmpStr)
				  		          ||"5".equals(tmpStr)||"6".equals(tmpStr)||"7".equals(tmpStr)||"8".equals(tmpStr)||"9".equals(tmpStr))){
				  		      rangeBean_field_amt = getField_amtValue(s_year,s_month,wr_rpt,bank_code,rangeBean_field_amt);
				  		    }
				  		}
				  		
				  		List operTmpRangeInfo = getOperTmpRangeInfo(s_year,s_month,rangeBean_acc_code,rangeBean_quop,rangeBean_field_amt,wr_rpt);
				  		if(operTmpRangeInfo.size()>0){
				  		  	int size = operTmpRangeInfo.size();
					  		if("max".equals(rangeBean_quop) || "min".equals(rangeBean_quop)){
					  		  	if (Integer.parseInt(rangeBean_field_amt)<size) size=Integer.parseInt(rangeBean_field_amt);
					  		  	for(int k=0;k<size;k++){
					                operTmpRangeBean = (DataObject)operTmpRangeInfo.get(k);
					                String operTmpRangeBean_bank_code = Utility.getTrimString(operTmpRangeBean.getValue("bank_code"));
					                String operTmpRangeBean_acc_code = Utility.getTrimString(operTmpRangeBean.getValue("acc_code"));
					                boolean isIns = true;
					                //1.警示指標13欄按照排名,排前5名,但零不要顯示,也不算入警示 field_992720_992710_RATE
									//2.警示指標16欄按照排名,排前5名,零及負數不要顯示,也不算入警示 field_diff_992730
									if("field_992720_992710_rate".equals(operTmpRangeBean_acc_code)||"field_diff_992730".equals(operTmpRangeBean_acc_code)){
									    Double operTmpRangeBean_amt = Double.parseDouble((operTmpRangeBean.getValue("amt")).toString());
										if((("field_992720_992710_rate".equals(operTmpRangeBean_acc_code)&& operTmpRangeBean_amt==0)
										        ||("field_diff_992730".equals(operTmpRangeBean_acc_code)&&operTmpRangeBean_amt<=0))){
										    isIns = false;
										}
									}
					                if(isIns){
						                dataList = new ArrayList();//傳內的參數List
						                dataList.add("Y");
									    dataList.add(rangeBean_remark); 
									    dataList.add(rangeBean_wr_range_serial);
										dataList.add(s_year); 
									    dataList.add(s_month); 
									    dataList.add(wr_rpt); 
									    dataList.add(operTmpRangeBean_bank_code);
									    dataList.add(operTmpRangeBean_acc_code); 
									    updateDBDataList.add(dataList);//1:傳內的參數List
					                }
					    	    }
					  		}else{
					  		    for(int k=0;k<operTmpRangeInfo.size();k++){
					  		      	operTmpRangeBean = (DataObject)operTmpRangeInfo.get(k);
					                String operTmpRangeBean_bank_code = Utility.getTrimString(operTmpRangeBean.getValue("bank_code"));
					                String operTmpRangeBean_acc_code = Utility.getTrimString(operTmpRangeBean.getValue("acc_code"));
					                String operTmpRangeBean_warn_type = Utility.getTrimString(operTmpRangeBean.getValue("warn_type"));					                
					                //String field_last_992710_990230_rate = getField_amtValue(s_year,s_month,wr_rpt,bank_code,"field_last_992710_990230_rate");
					                //String field_992710_990230_rate = getField_amtValue(s_year,s_month,wr_rpt,bank_code,"field_992710_990230_rate");
					                System.out.println("bank_code="+operTmpRangeBean_bank_code+":acc_code="+operTmpRangeBean_acc_code);
						            dataList = new ArrayList();//傳內的參數List
						            //當欄位是=field_992710_990230_rate時,則增加檢核
						            //103.06.10取消-->field_992710_990230_warn_type若為'N'時,表示上月/上季/上年度同期沒有警示,本月有警示.才寫入WR_OPERATION.WARN_TYPE ='Y'
						            //103.06.10 add 當上月份之(建築放款/上年度信用部決算淨值)低於100%,但本月份之(建築放款/上年度信用部決算淨值)高於100%,才列警示
						            String field_992710_990230_warn_type = getField_992710_990230_Warn_Type(s_year,s_month,wr_rpt,operTmpRangeBean_bank_code);
						            if(!"field_992710_990230_rate".equals(rangeBean_acc_code)
						                    ||("field_992710_990230_rate".equals(rangeBean_acc_code) && "Y".equals(field_992710_990230_warn_type)))
						            {
						                dataList.add("Y");
						                dataList.add(rangeBean_remark); 
									    dataList.add(rangeBean_wr_range_serial);
										dataList.add(s_year); 
									    dataList.add(s_month); 
									    dataList.add(wr_rpt); 
									    dataList.add(operTmpRangeBean_bank_code);
									    dataList.add(operTmpRangeBean_acc_code); 
									    updateDBDataList.add(dataList);//1:傳內的參數List
								    }
					  		    }
					  		}
				  		} 
			    	}
		  		}else if("WR03".equals(report_no)||"WR04".equals(report_no)||"WR05".equals(report_no)){//專案農貸
		  			//取出符合該range的bank
			  		if(!"max".equals(rangeBean_quop) && !"min".equals(rangeBean_quop)){
			  		    String tmpStr = rangeBean_field_amt.substring(0, 1);
			  		    if(!("0".equals(tmpStr)||"1".equals(tmpStr)||"2".equals(tmpStr)||"3".equals(tmpStr)||"4".equals(tmpStr)
			  		          ||"5".equals(tmpStr)||"6".equals(tmpStr)||"7".equals(tmpStr)||"8".equals(tmpStr)||"9".equals(tmpStr))){
			  		      rangeBean_field_amt = getField_amtValue(s_year,s_month,wr_rpt,bank_code,rangeBean_field_amt);
			  		    }
			  		}
			  		
			  		List operTmpRangeInfo = getOperTmpRangeInfo(s_year,s_month,rangeBean_acc_code,rangeBean_quop,rangeBean_field_amt,wr_rpt);
			  		if(operTmpRangeInfo.size()>0){
			  		  	int size = operTmpRangeInfo.size();
				  		if("max".equals(rangeBean_quop) || "min".equals(rangeBean_quop)){
				  		  	if (Integer.parseInt(rangeBean_field_amt)<size) size=Integer.parseInt(rangeBean_field_amt);
				  		  	for(int k=0;k<size;k++){
				                operTmpRangeBean = (DataObject)operTmpRangeInfo.get(k);
				                String operTmpRangeBean_bank_code = Utility.getTrimString(operTmpRangeBean.getValue("bank_code"));
				                String operTmpRangeBean_acc_code = Utility.getTrimString(operTmpRangeBean.getValue("acc_code"));
				                dataList = new ArrayList();//傳內的參數List
				                dataList.add("Y");
							    dataList.add(rangeBean_remark); 
							    dataList.add(rangeBean_wr_range_serial);
								dataList.add(s_year); 
							    dataList.add(s_month); 
							    dataList.add(wr_rpt); 
							    dataList.add(operTmpRangeBean_bank_code);
							    dataList.add(operTmpRangeBean_acc_code); 
							    updateDBDataList.add(dataList);//1:傳內的參數List
				    	    }
				  		}else{
				  		    for(int k=0;k<operTmpRangeInfo.size();k++){
				  		      	operTmpRangeBean = (DataObject)operTmpRangeInfo.get(k);
				                String operTmpRangeBean_bank_code = Utility.getTrimString(operTmpRangeBean.getValue("bank_code"));
				                String operTmpRangeBean_acc_code = Utility.getTrimString(operTmpRangeBean.getValue("acc_code"));
				                String operTmpRangeBean_warn_type = Utility.getTrimString(operTmpRangeBean.getValue("warn_type"));
					            dataList = new ArrayList();//傳內的參數List
					            dataList.add("Y");
				                dataList.add(rangeBean_remark); 
							    dataList.add(rangeBean_wr_range_serial);
								dataList.add(s_year); 
							    dataList.add(s_month); 
							    dataList.add(wr_rpt); 
							    dataList.add(operTmpRangeBean_bank_code);
							    dataList.add(operTmpRangeBean_acc_code); 
							    updateDBDataList.add(dataList);//1:傳內的參數List
				  		    }
				  		}
			  		}
		  		}
		  	}
		  	updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
			updateDBList.add(updateDBSqlList);
		  	if(!DBManager.updateDB_ps(updateDBList)){					 
				errMsg = errMsg + "更新 WR_OPERATION_TMP 失敗<br>[DBManager.getErrMsg()]:<br>";
			}else{
				if(isDebug.equals("true"))  System.out.println("***********更新 WR_OPERATION_TMP 成功");
			}
		  	
		    //4.將WR_OPERATION_TMP逐筆寫入WR_OPERATION
		  	updateDBSqlList = new ArrayList();		
	    	updateDBDataList = new ArrayList();
	    	updateDBList = new ArrayList();
			sqlCmd.delete(0, sqlCmd.length());
			dataList = new ArrayList();//傳內的參數List
			sqlCmd.append("insert into WR_OPERATION(M_YEAR,M_MONTH,WR_RPT,BANK_CODE,ACC_CODE,TYPE,AMT,WARN_TYPE,REMARK,WR_RANGE_SERIAL) select M_YEAR,M_MONTH,WR_RPT,BANK_CODE,ACC_CODE,TYPE,AMT,WARN_TYPE,REMARK,WR_RANGE_SERIAL from WR_OPERATION_TMP where m_year=? and m_month=? and wr_rpt=? ");
			dataList.add(s_year); 
		    dataList.add(s_month); 
		    dataList.add(wr_rpt);
			updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql 
		    updateDBDataList.add(dataList);//1:傳內的參數List
			updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
			updateDBList.add(updateDBSqlList);
			if(!DBManager.updateDB_ps(updateDBList)){					 
				errMsg = errMsg + "更新 WR_OPERATION 失敗<br>[DBManager.getErrMsg()]:<br>";
			}else{
				if(isDebug.equals("true"))  System.out.println("***********更新 WR_OPERATION 成功");
			}
		    //5.資料產生後,寫入WR_OPERATION_LOG
	        updateDBSqlList = new ArrayList();		
			updateDBDataList = new ArrayList();
			updateDBList = new ArrayList();
	        sqlCmd.delete(0, sqlCmd.length());
	    	dataList = new ArrayList();//傳內的參數List
	    	if(haveOperationLogInfo(s_year,s_month,report_no)){
	    	    sqlCmd.append("update WR_OPERATION_LOG set TOTAL=?,USER_ID=?,USER_NAME=?,UPDATE_DATE=sysdate where m_year=? and m_month = ? and REPORT_NO=? and KIND_TYPE=? ");
	    	    dataList.add(total_count); 
	    		dataList.add(lguser_id); 
	    		dataList.add(lguser_id);
	    		dataList.add(s_year); 
	    	    dataList.add(s_month);
	    		dataList.add(report_no);
	    	    dataList.add("bank_no_ALL");
	    	}else{
	            sqlCmd.append("insert into WR_OPERATION_LOG(M_YEAR,M_MONTH,REPORT_NO,KIND_TYPE,TOTAL,USER_ID,USER_NAME,UPDATE_DATE) values(?,?,?,?,?,?,?,sysdate)");
	            dataList.add(s_year); 
	    	    dataList.add(s_month);
	    	    dataList.add(report_no);
	    	    dataList.add("bank_no_ALL");
	    	    dataList.add(total_count); 
	    		dataList.add(lguser_id); 
	    		dataList.add(lguser_id);
	        }
		    updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql 
		    updateDBDataList.add(dataList);//1:傳內的參數List
			updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
			updateDBList.add(updateDBSqlList);
			if(!DBManager.updateDB_ps(updateDBList)){					 
				errMsg = errMsg + "更新 WR_OPERATION_LOG 失敗<br>[DBManager.getErrMsg()]:<br>";
			}else{
				if(isDebug.equals("true"))  System.out.println("***********更新 WR_OPERATION_LOG 成功");
			}
		    //=================================================================================================================================
	    }
		logps.flush();	      	   		
	}catch (Exception e){
		System.out.println(e+":"+e.getMessage());
	    errMsg = errMsg + "相關資料寫入資料庫失敗 excetion";
	    logps.println(Utility.getDateFormat("yyyy/MM/dd  HH:mm:ss  ")+" "+"UpdateDB Error:"+e + "\n"+e.getMessage());	
		logps.flush();		   
	}finally{
		try{
			if (logos  != null) logos.close();
			if (logbos != null) logbos.close();
			if (logps  != null) logps.close();
		}catch(Exception ioe){
			System.out.println(ioe.getMessage());
		}
	}	
	return errMsg;
}
private List getDetailInfo(String s_year,String s_month,String compareYY,String compareMM,String wlx01_m_year,String wr_rpt){
    StringBuffer sqlCmd = new StringBuffer();        	
	List paramList = new ArrayList();
	sqlCmd.append("select a01.m_year,a01.m_month,a01.bank_no , a01.BANK_NAME,");
	sqlCmd.append(" 	  decode(a01_last.field_DEBIT,0,0,round((a01.field_DEBIT - a01_last.field_DEBIT)/  a01_last.field_DEBIT *100 ,2))  as   field_DEBIT_RATE,");// --增減金額/存款總額%
	sqlCmd.append(" 	  decode(a01_last.field_CREDIT,0,0,round((a01.field_CREDIT - a01_last.field_CREDIT)/  a01_last.field_CREDIT *100 ,2))  as   field_CREDIT_RATE,");// --增減金額/放款總額%
	sqlCmd.append(" 	  decode(a01_last.field_990610,0,0,round((a01.field_990610 - a01_last.field_990610)/  a01_last.field_990610 *100 ,2))  as   field_990610_RATE,");// --增減金額/非會員放款%
	sqlCmd.append(" 	  decode(a01_last.field_NOASSURE,0,0,round((a01.field_NOASSURE - a01_last.field_NOASSURE)/  a01_last.field_NOASSURE *100 ,2))  as   field_NOASSURE_RATE,");// --增減金額/無擔保放款 
	sqlCmd.append(" 	  decode(a01_last.field_120700,0,0,round((a01.field_120700 - a01_last.field_120700)/  a01_last.field_120700 *100 ,2))  as field_120700_RATE,");//--增加金額/內部融資
	sqlCmd.append(" 	  round(a01.field_992710 /1,0)  as field_992710,");//  --建築放款
	sqlCmd.append(" 	  decode(a01_last.field_992710,0,0,round((a01.field_992710 - a01_last.field_992710)/  a01_last.field_992710 *100 ,2))  as field_992710_RATE,");//--增加金額/建築放款
	sqlCmd.append(" 	  decode(a01.field_CREDIT,0,0,round(a01.field_992710 /  a01.field_CREDIT *100 ,2))  as field_992710_CREDIT_RATE,");//--建築放款/放款  
    //--decode(a01.field_CREDIT,0,0,round(a01.field_992710 /  a01.field_CREDIT *100 ,2)) - 
    //--decode(a01_last.field_CREDIT,0,0,round(a01_last.field_992710 /  a01_last.field_CREDIT *100 ,2)) 
    //--as field_diff_992710_CREDIT_RATE,   --本月份之(建築放款/放款)-上月份之(建築放款/放款)     
    sqlCmd.append("       decode((a01.field_990230-a01.field_990240-a01.field_992810),0,0,round(a01.field_992710 /  (a01.field_990230-a01.field_990240-a01.field_992810) *100 ,2))  as   field_992710_990230_RATE,");//--本月份的建築放款/上年度信用部決算淨值 103.06.10 add
    sqlCmd.append("       decode((a01_last.field_990230-a01_last.field_990240-a01_last.field_992810),0,0,round(a01_last.field_992710 /  (a01_last.field_990230-a01_last.field_990240-a01_last.field_992810) *100 ,2)) as field_last_992710_990230_RATE, ");//--上月份的建築放款/上年度信用部決算淨值 103.06.10 add       
    //sqlCmd.append(" 	  decode(wr_operation_last.warn_type,null,'N',wr_operation_last.warn_type) as field_992710_990230_warn_type,");//--上個月.建築放款/上年度信用部決算淨值.警示註記 103.02.26 add
    //--decode(a01.field_NET,0,0,round(a01.field_992710 /  a01.field_NET *100/302 ,2))  as   field_992710_NET_AVG_RATE,--建築放款/淨值.同業平均數
    sqlCmd.append(" 	  decode((a01.field_990230-a01.field_990240-a01.field_992810),0,0,round(a01.field_992710 /  (a01.field_990230-a01.field_990240-a01.field_992810) *100 ,2)) - ");
    sqlCmd.append(" 	  decode((a01_last.field_990230-a01_last.field_990240-a01_last.field_992810),0,0,round(a01_last.field_992710 /  (a01_last.field_990230-a01_last.field_990240-a01_last.field_992810) *100 ,2))  ");
    sqlCmd.append(" 	  as   field_diff_992710_990230_RATE, ");// --本月份之(建築放款/上年度信用部決算淨值)-上月份之(建築放款/上年度信用部決算淨值) //103.06.17調整公式
    sqlCmd.append(" 	  round(a01.field_OVER /1,0)  as field_OVER,");//  --逾期放款       
    sqlCmd.append(" 	  a01.field_OVER - a01_last.field_OVER as field_diff_OVER,");//--逾期放款增加金額
    //--decode(a01_last.field_OVER,0,0,round((a01.field_OVER - a01_last.field_OVER)/  a01_last.field_OVER *100 ,2))  as field_OVER_RATE,--增加金額/逾期放款
    sqlCmd.append(" 	  decode(a01.field_CREDIT,0,0,round(a01.field_OVER /  a01.field_CREDIT *100 ,2))  as   field_OVER_RATE,");//--逾放比率    
    sqlCmd.append(" 	  decode(a01.field_CREDIT,0,0,round(a01.field_OVER /  a01.field_CREDIT *100 ,2)) - ");
    sqlCmd.append(" 	  decode(a01_last.field_CREDIT,0,0,round(a01_last.field_OVER /  a01_last.field_CREDIT *100 ,2)) as field_diff_OVER_RATE,");//--本月份逾放比率-上月份逾放比率
    sqlCmd.append(" 	  a01.field_992530 - a01_last.field_992530 as field_diff_992530,");//  --逾放-非會員增加金額     
    //--a01.field_992550 - a01_last.field_992550 as field_diff_992550, --逾放-無擔保消費性放款增加金額
    sqlCmd.append(" 	  decode(a01.field_992150,0,0,round(a01.field_992550 /  a01.field_992150 *100 ,2))  as field_992550_992150_RATE,");// --無擔保消費性放款中之逾放A99.992550/無擔保消費性放款A99.992150(V)
    //--a01.field_992720 - a01_last.field_992720 as field_diff_992720,  --逾放-建築放款增加金額
    sqlCmd.append(" 	  decode(a01.field_992710,0,0,round(a01.field_992720 /  a01.field_992710 *100 ,2))  as field_992720_992710_RATE,");//--建築放款中之逾放/建築放款=A99.992720/A99.992710
    sqlCmd.append(" 	  a01.field_992610_cal,");//--應予觀察放款 
    sqlCmd.append(" 	  a01.field_992610_cal - a01_last.field_992610_cal as field_diff_992610_cal,");//--應予觀察放款增加金額
    sqlCmd.append(" 	  decode(a01.field_CREDIT,0,0,round(a01.field_992610_cal/  a01.field_CREDIT *100 ,2))  as field_992610_cal_CREDIT_RATE,");//--應予觀察放款/放款
    sqlCmd.append(" 	  a01.field_992630 - a01_last.field_992630 as field_diff_992630,");//--應予觀察放款-非會員增加金額
    sqlCmd.append(" 	  a01.field_992730 - a01_last.field_992730 as field_diff_992730,");//--應予觀察放款-建築放款增加金額
    sqlCmd.append(" 	  a01.field_BACKUP - a01_last.field_BACKUP as field_diff_BACKUP,");//--備抵呆帳增加金額       
    sqlCmd.append(" 	  decode(a01.field_OVER,0,0,round(a01.field_BACKUP /  a01.field_OVER *100 ,2)) as   field_BACKUP_OVER_RATE,");//--備抵呆帳/逾期放款,若逾期放款為0者,需排除
    //--decode(a01.field_OVER,0,0,round(a01.field_BACKUP /  a01.field_OVER *100 ,2)) - 
    //--decode(a01_last.field_OVER,0,0,round(a01_last.field_BACKUP /  a01_last.field_OVER *100 ,2)) 
    //--as field_diff_BACKUP_OVER_RATE ");//--本月份之備抵呆帳覆蓋率-上月份備抵呆帳覆蓋率,<100才要算
    sqlCmd.append(" 	  decode(a01.field_CREDIT,0,0,round(a01.field_BACKUP /  a01.field_CREDIT *100 ,2)) - ");
    sqlCmd.append(" 	  decode(a01_last.field_CREDIT,0,0,round(a01_last.field_BACKUP /  a01_last.field_CREDIT *100 ,2)) as field_diff_BACKUP_CREDIT_RATE ");//--本月份之放款覆蓋率-上月份放款覆蓋率(備抵呆帳/放款)                  
    sqlCmd.append(" from ( select a01.m_year,a01.m_month, a01.bank_no ,   a01.BANK_NAME,");
    sqlCmd.append(" 	          SUM(field_DEBIT)    as field_DEBIT ,");
    sqlCmd.append(" 	          SUM(field_CREDIT)   as field_CREDIT,");  
    sqlCmd.append(" 	          SUM(field_NOASSURE) as field_NOASSURE,");     
    sqlCmd.append(" 	          SUM(field_120700)   as field_120700,");   
	//--SUM(field_NET)      as field_NET, 
	sqlCmd.append(" 	          SUM(field_OVER)   as field_OVER,");     
	sqlCmd.append(" 	          SUM(field_BACKUP) as field_BACKUP,"); 
	sqlCmd.append(" 	          SUM(field_990610) as field_990610,");
	sqlCmd.append(" 	          SUM(field_990230) as field_990230,");
	sqlCmd.append(" 	          SUM(field_990240) as field_990240,");
    sqlCmd.append(" 	          SUM(field_992810) as field_992810,");
	sqlCmd.append(" 	          SUM(field_992610_cal) as field_992610_cal,");
	sqlCmd.append(" 	          SUM(field_992710) as field_992710,");
	sqlCmd.append(" 	          SUM(field_992530) as field_992530,");
	sqlCmd.append(" 	          SUM(field_992510) as field_992510,");
	sqlCmd.append(" 	          SUM(field_992550) as field_992550,");
	sqlCmd.append(" 	          SUM(field_992150) as field_992150,");
	sqlCmd.append(" 	          SUM(field_992720) as field_992720,");
	sqlCmd.append(" 	          SUM(field_992630) as field_992630,");
	sqlCmd.append(" 	          SUM(field_992730) as field_992730 "); 
	sqlCmd.append("        from ");
	sqlCmd.append(" 	        ( select a01.m_year,a01.m_month, bn01.bank_no , bn01.BANK_NAME,");           
	sqlCmd.append(" 	                 round(sum(decode(a01.acc_code,'220000',amt,0)) /1,0) as field_DEBIT,");
	sqlCmd.append(" 	                 round(sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /1,0) as  field_CREDIT,"); 
	sqlCmd.append(" 	                 round(sum( decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code, '120101',amt,'120301',amt, '120401',amt, '120501',amt,0),'7',decode(a01.acc_code, '120101',amt,'120401', amt, '120201',amt, '120501',amt,0)),");  
	sqlCmd.append(" 	                                             '103',decode(a01.acc_code, '120101',amt,'120301',amt, '120401',amt, '120501',amt,0),0) ) /1,0) as  field_NOASSURE,");  
	sqlCmd.append(" 	                 round(sum(decode(a01.acc_code,'120700',amt,0)) /1,0) as field_120700,");          
    //--round(sum(decode(bn01.bank_type,'6',decode(a01.acc_code,'310000',amt,'320000',amt,0),'7',decode(a01.acc_code,'300000',amt,0),0)) /1,0)     as field_NET,");
	sqlCmd.append(" 	                 round(sum(decode(a01.acc_code,'990000',amt,0)) /1,0) as field_OVER,");
	sqlCmd.append(" 	                 round(sum(decode(a01.acc_code, '120800',amt,'150300',amt,0)) /1,0) as  field_BACKUP ");
	sqlCmd.append(" 	            from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
	sqlCmd.append(" 	            left join (select (CASE WHEN (a01.m_year <= 102) THEN '102' ");                             
	sqlCmd.append(" 	                                    WHEN (a01.m_year > 102) THEN '103' ");                             
	sqlCmd.append(" 	                                    ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01 ");
	sqlCmd.append(" 	                        where m_year=? and m_month=? ");
	sqlCmd.append(" 	                       ) a01  on  bn01.bank_no = a01.bank_code ");
	sqlCmd.append(" 	            group by a01.m_year,a01.m_month,bn01.bank_no,bn01.BANK_NAME ");
	sqlCmd.append(" 	          ) a01,"); 
	paramList.add(wlx01_m_year);
	paramList.add(s_year);
	paramList.add(s_month);
	sqlCmd.append(" 	          (select a02.m_year,a02.m_month, bn01.bank_no as bank_code,  bn01.bank_name,");        
	sqlCmd.append(" 	                  round(sum(decode(a02.acc_code,'990610',amt,0)) /1,0) as field_990610,");
	sqlCmd.append(" 	                  round(sum(decode(a02.acc_code,'990230',amt,0)) /1,0) as field_990230,");  
	sqlCmd.append(" 	                  round(sum(decode(a02.acc_code,'990240',amt,0)) /1,0) as field_990240 ");       
	sqlCmd.append(" 	             from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
	sqlCmd.append(" 	             left join (select * from a02 where m_year = ? and m_month=? ) a02 on  bn01.bank_no = a02.bank_code ");    
	sqlCmd.append(" 	            group by a02.m_year,a02.m_month,bn01.bank_no,bn01.BANK_NAME ");
	sqlCmd.append(" 	          ) a02,");
	paramList.add(wlx01_m_year);
	paramList.add(s_year);
	paramList.add(s_month);
	sqlCmd.append(" 	          ( select a99.m_year,a99.m_month, bn01.bank_no as bank_code, bn01.BANK_NAME,");           
	sqlCmd.append(" 	                   round(sum(decode(a99.acc_code,'992150',amt,0)) /1,0) as field_992150,");   
	sqlCmd.append(" 	                   round(sum(decode(a99.acc_code,'992510',amt,0)) /1,0) as field_992510,"); 
	sqlCmd.append(" 	                   round(sum(decode(a99.acc_code,'992530',amt,0)) /1,0) as field_992530,");  
	sqlCmd.append(" 	                   round(sum(decode(a99.acc_code,'992550',amt,0)) /1,0) as field_992550,"); 
	sqlCmd.append(" 	                   round(sum(decode(a99.acc_code,'992610',amt,'992620',amt,'992630',amt,'992640',amt,'992650',amt,0)) /1,0) as field_992610_cal,");// --應予觀察放款
	sqlCmd.append(" 	                   round(sum(decode(a99.acc_code,'992630',amt,0)) /1,0) as field_992630,");
	sqlCmd.append(" 	                   round(sum(decode(a99.acc_code,'992710',amt,0)) /1,0) as field_992710,");// --建築放款
	sqlCmd.append(" 	                   round(sum(decode(a99.acc_code,'992720',amt,0)) /1,0) as field_992720,");             
	sqlCmd.append(" 	                   round(sum(decode(a99.acc_code,'992730',amt,0)) /1,0) as field_992730,");  
	sqlCmd.append(" 	                   round(sum(decode(a99.acc_code,'992810',amt,0)) /1,0) as field_992810 ");  
	sqlCmd.append(" 	              from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
	sqlCmd.append(" 	              left join (select * from a99 where m_year = ? and m_month=?)a99 on bn01.bank_no = a99.bank_code ");
	sqlCmd.append(" 	             group by a99.m_year,a99.m_month,bn01.bank_no,bn01.BANK_NAME ");
	sqlCmd.append(" 	            ) a99 ");
	paramList.add(wlx01_m_year);
	paramList.add(s_year);
	paramList.add(s_month);
	sqlCmd.append("        where a01.bank_no=a02.bank_code(+) and a01.bank_no = a99.bank_code(+) ");
	sqlCmd.append("        and a01.bank_no <> ' ' "); 
	sqlCmd.append("        GROUP BY a01.m_year,a01.m_month,a01.bank_no,a01.BANK_NAME ");
	sqlCmd.append(" ) a01,");//--本月份資料
	sqlCmd.append(" ( select a01.m_year,a01.m_month, a01.bank_no, a01.BANK_NAME,"); 
	sqlCmd.append("          SUM(field_DEBIT)  field_DEBIT ,");
	sqlCmd.append("          SUM(field_CREDIT) field_CREDIT,");   
	sqlCmd.append("          SUM(field_NOASSURE) field_NOASSURE,");     
	sqlCmd.append("          SUM(field_120700) field_120700,"); 
	//sqlCmd.append("          SUM(field_NET)   field_NET,"); 
	sqlCmd.append("          SUM(field_OVER)  field_OVER,   ");  
	sqlCmd.append("          SUM(field_BACKUP)  field_BACKUP,  ");
	sqlCmd.append("          SUM(field_990610) as field_990610,");
	sqlCmd.append("          SUM(field_990230) as field_990230,");
	sqlCmd.append("          SUM(field_990240) as field_990240,");
    sqlCmd.append("          SUM(field_992810) as field_992810,");
	sqlCmd.append("          SUM(field_992610_cal) as field_992610_cal,");
	sqlCmd.append("          SUM(field_992710) as field_992710,");
	sqlCmd.append("          SUM(field_992530) as field_992530,");
	sqlCmd.append("          SUM(field_992550) as field_992550,");
	sqlCmd.append("          SUM(field_992150) as field_992150,");
	sqlCmd.append("          SUM(field_992720) as field_992720,");
	sqlCmd.append("          SUM(field_992630) as field_992630,");
	sqlCmd.append("          SUM(field_992730) as field_992730 "); 
	sqlCmd.append("        from ");
	sqlCmd.append("        ( select a01.m_year,a01.m_month, bn01.bank_no , bn01.BANK_NAME,");           
	sqlCmd.append("            round(sum(decode(a01.acc_code,'220000',amt,0)) /1,0) as field_DEBIT,");
	sqlCmd.append("            round(sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /1,0) as  field_CREDIT,"); 
	sqlCmd.append("            round(sum( decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code, '120101',amt,'120301',amt, '120401',amt, '120501',amt,0),'7',decode(a01.acc_code, '120101',amt,'120401', amt, '120201',amt, '120501',amt,0)),");  
	sqlCmd.append("                                        '103',decode(a01.acc_code, '120101',amt,'120301',amt, '120401',amt, '120501',amt,0),0) ) /1,0) as  field_NOASSURE,");  
	sqlCmd.append("            round(sum(decode(a01.acc_code,'120700',amt,0)) /1,0) as field_120700,");
	//sqlCmd.append("            round(sum(decode(bn01.bank_type,'6',decode(a01.acc_code,'310000',amt,'320000',amt,0),'7',decode(a01.acc_code,'300000',amt,0),0)) /1,0) as field_NET,");
	sqlCmd.append("            round(sum(decode(a01.acc_code,'990000',amt,0)) /1,0) as field_OVER,");
	sqlCmd.append("            round(sum(decode(a01.acc_code, '120800',amt,'150300',amt,0)) /1,0) as  field_BACKUP ");
	sqlCmd.append("          from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
	sqlCmd.append("          left join (select (CASE WHEN (a01.m_year <= 102) THEN '102' ");                             
	sqlCmd.append("                                   WHEN (a01.m_year > 102) THEN '103' ");                             
	sqlCmd.append("                                   ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01 ");
	sqlCmd.append("                      where m_year=? and m_month=? "); 
	sqlCmd.append("                      ) a01  on  bn01.bank_no = a01.bank_code ");
	sqlCmd.append("          group by a01.m_year,a01.m_month,bn01.bank_no,bn01.BANK_NAME ");
	sqlCmd.append("        ) a01,"); 
	paramList.add(wlx01_m_year);
	paramList.add(compareYY);
	paramList.add(compareMM);
	sqlCmd.append("        (select a02.m_year,a02.m_month, bn01.bank_no as bank_code,  bn01.bank_name,");        
	sqlCmd.append("           round(sum(decode(a02.acc_code,'990610',amt,0)) /1,0) as field_990610,");
	sqlCmd.append("           round(sum(decode(a02.acc_code,'990230',amt,0)) /1,0) as field_990230,");
	sqlCmd.append("           round(sum(decode(a02.acc_code,'990240',amt,0)) /1,0) as field_990240 ");  
	sqlCmd.append("         from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
	sqlCmd.append("         left join (select * from a02 where m_year = ? and m_month=?) a02 on bn01.bank_no = a02.bank_code ");    
	sqlCmd.append("         group by a02.m_year,a02.m_month,bn01.bank_no,bn01.BANK_NAME ");
	sqlCmd.append("        ) a02,");  
	paramList.add(wlx01_m_year);
	paramList.add(compareYY);
	paramList.add(compareMM);
	sqlCmd.append("        ( select a99.m_year,a99.m_month, bn01.bank_no as bank_code, bn01.BANK_NAME,");           
	sqlCmd.append("            round(sum(decode(a99.acc_code,'992150',amt,0)) /1,0) as field_992150,");   
	sqlCmd.append("            round(sum(decode(a99.acc_code,'992510',amt,0)) /1,0) as field_992510,"); 
	sqlCmd.append("            round(sum(decode(a99.acc_code,'992530',amt,0)) /1,0) as field_992530,");  
	sqlCmd.append("            round(sum(decode(a99.acc_code,'992550',amt,0)) /1,0) as field_992550,"); 
	sqlCmd.append("            round(sum(decode(a99.acc_code,'992610',amt,'992620',amt,'992630',amt,'992640',amt,'992650',amt,0)) /1,0) as field_992610_cal,"); //--應予觀察放款
	sqlCmd.append("            round(sum(decode(a99.acc_code,'992630',amt,0)) /1,0) as field_992630,");
	sqlCmd.append("            round(sum(decode(a99.acc_code,'992710',amt,0)) /1,0) as field_992710,"); //--建築放款
	sqlCmd.append("            round(sum(decode(a99.acc_code,'992720',amt,0)) /1,0) as field_992720,");             
	sqlCmd.append("            round(sum(decode(a99.acc_code,'992730',amt,0)) /1,0) as field_992730,");  
	sqlCmd.append("            round(sum(decode(a99.acc_code,'992810',amt,0)) /1,0) as field_992810 ");  
	sqlCmd.append("          from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
	sqlCmd.append("          left join (select * from a99 where m_year = ? and m_month=?)a99 on bn01.bank_no = a99.bank_code ");
	sqlCmd.append("          group by a99.m_year,a99.m_month,bn01.bank_no,bn01.BANK_NAME ");
	sqlCmd.append("        ) a99 ");
	paramList.add(wlx01_m_year);
	paramList.add(compareYY);
	paramList.add(compareMM);
	sqlCmd.append("        where a01.bank_no=a02.bank_code(+) and a01.bank_no = a99.bank_code(+) ");
	sqlCmd.append("        and a01.bank_no <> ' ' "); 
	sqlCmd.append("        GROUP BY a01.m_year,a01.m_month,a01.bank_no,a01.BANK_NAME ");
	sqlCmd.append(" ) a01_last ");//上月份資料/上季資料/上年度同期  103.02.27 add
	/*103.06.10 取消
	sqlCmd.append(" ( select m_year,m_month,bank_code,warn_type ");//--103.02.26 add
	sqlCmd.append("     from wr_operation ");
	sqlCmd.append("    where m_year = ? and m_month = ? ");
	sqlCmd.append("      and wr_rpt = ? ");//103.06.03 add 警示報表類別
	sqlCmd.append("      and acc_code = 'field_992710_990230_rate' ");//--建築放款/上年度信用部決算淨值)
	sqlCmd.append("    group by m_year,m_month,bank_code,warn_type ");
	sqlCmd.append(" )wr_operation_last ");//--上月份/上季/上年度同期警訊資料
	paramList.add(compareYY);
	paramList.add(compareMM);
	paramList.add(wr_rpt);//103.06.03 add 警示報表類別
	*/
	sqlCmd.append(" where a01.bank_no=a01_last.bank_no(+)"); 
	
	List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,
	        "m_year,m_month,bank_no,bank_name,"
	        +"field_debit_rate,field_credit_rate,field_990610_rate,field_noassure_rate,field_120700_rate,"
	        +"field_992710,field_992710_rate,field_992710_credit_rate,field_992710_990230_rate,field_last_992710_990230_rate,"
	        +"field_diff_992710_990230_rate,field_over,field_diff_over,field_over_rate,field_diff_over_rate,"
	        +"field_diff_992530,field_992550_992150_rate,field_992720_992710_rate,field_992610_cal,field_diff_992610_cal,"
	        +"field_992610_cal_credit_rate,field_diff_992630,field_diff_992730,field_diff_backup,field_backup_over_rate,"
	        +"field_diff_backup_credit_rate");
	return dbData;
}
private List getLoanDetailInfo(String s_year,String s_month,String compareYY,String compareMM,String wlx01_m_year){
    StringBuffer sqlCmd = new StringBuffer();        	
	List paramList = new ArrayList();
	sqlCmd.append("select * ");
    sqlCmd.append("from ( ");
	sqlCmd.append("select agri_loan.m_year,");//--WR_OPERATION_TMP資料年度
	sqlCmd.append("       agri_loan.m_month,");//--WR_OPERATION_TMP資料月份
	sqlCmd.append("       agri_loan.bank_no,");//--WR_OPERATION_TMP機構代號
	sqlCmd.append("       agri_loan.BANK_NAME,");    
	sqlCmd.append("       agri_loan.field_loan_bal_amt - agri_loan_last.field_loan_bal_amt as field_diff_loan_bal_amt,");//--專案農貸總額:本月(7)貸放餘額-上月(7)貸放餘額       
	sqlCmd.append("       agri_loan.field_over6m_loan_bal_amt - agri_loan_last.field_over6m_loan_bal_amt as field_diff_over6m_loan_bal_amt,");//--專案農貸逾期放款:本月(9)逾期六個月放款餘額-金額-上月(9)逾期六個月放款餘額-金額       
	sqlCmd.append("       decode(agri_loan.field_loan_bal_amt,0,0,round(agri_loan.field_over6m_loan_bal_amt /  agri_loan.field_loan_bal_amt *100 ,2)) - ");
	sqlCmd.append("       decode(agri_loan_last.field_loan_bal_amt,0,0,round(agri_loan_last.field_over6m_loan_bal_amt /  agri_loan_last.field_loan_bal_amt *100 ,2)) ");  
	sqlCmd.append("       as   field_diff_over6m_loan_rate,");//--專案農貸逾放比率:本月專案農貸逾放比率-上月專案農貸逾放比率((9)逾期六個月放款餘額-金額/(7)貸放餘額 )
	sqlCmd.append("       decode(agri_loan.field_loan_bal_cnt,0,0,round(agri_loan.field_delay_loan_cnt_sum /  agri_loan.field_loan_bal_cnt *100 ,2))  as field_delay_loan_rate ");//--當年度累計核准延期還款件數占尚有餘額專案農貸總件數比率:(10)當月核准延期還款-件數.1月~當月累計/(6)貸放餘額-戶數              
	sqlCmd.append("  from ( select agri_loan.m_year,agri_loan.m_month, agri_loan.bank_no ,   agri_loan.BANK_NAME,");
	sqlCmd.append("                field_loan_bal_cnt ,");//--(6)貸放餘額-戶數
	sqlCmd.append("                field_loan_bal_amt,"); //--(7)貸放餘額 
	sqlCmd.append("                field_over6m_loan_bal_amt,");//--(9)逾期六個月放款餘額-金額     
	sqlCmd.append("                SUM(field_delay_loan_cnt_sum)   as field_delay_loan_cnt_sum ");//--(10)當月核准延期還款-件數              
	sqlCmd.append("           from ");
	sqlCmd.append("               (select agri_loan.m_year,agri_loan.m_month, bn01.bank_no , bn01.BANK_NAME,");        
	sqlCmd.append("                       round(sum(agri_loan.loan_bal_cnt) /1,0) as field_loan_bal_cnt,");//--(6)貸放餘額-戶數
	sqlCmd.append("                       round(sum(agri_loan.loan_bal_amt) /1,0) as field_loan_bal_amt,");//--(7)貸放餘額
	sqlCmd.append("                       round(sum(agri_loan.over6m_loan_bal_amt) /1,0) as field_over6m_loan_bal_amt ");//--(9)逾期六個月放款餘額-金額          
	sqlCmd.append("                  from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
	sqlCmd.append("                  left join (select * from agri_loan  where m_year=? and m_month=?) agri_loan  on  bn01.bank_no = agri_loan.bank_code ");
	paramList.add(wlx01_m_year);
	paramList.add(s_year);
	paramList.add(s_month);
	sqlCmd.append("                 group by agri_loan.m_year,agri_loan.m_month,bn01.bank_no,bn01.BANK_NAME ");
	sqlCmd.append("                ) agri_loan,");
	sqlCmd.append("               (select agri_loan.m_year,agri_loan.m_month, bn01.bank_no , bn01.BANK_NAME,");        
	sqlCmd.append("                       sum(agri_loan.delay_loan_cnt) as field_delay_loan_cnt_sum ");//--(10)當月核准延期還款-件數        
	sqlCmd.append("        	  from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
	sqlCmd.append("        	  left join (select * from agri_loan  where (to_char(m_year * 100 + m_month) >= ? "); 
	sqlCmd.append("		   					and to_char(m_year * 100 + m_month) <= ?) ");  
	paramList.add(wlx01_m_year);
	paramList.add(s_year+"01");
	paramList.add(s_year+s_month);
	sqlCmd.append("                  	     ) agri_loan  on  bn01.bank_no = agri_loan.bank_code ");
	sqlCmd.append("        	 group by agri_loan.m_year,agri_loan.m_month,bn01.bank_no,bn01.BANK_NAME ");//--.1月~當月累計 
	sqlCmd.append("       		)agri_loan_sum ");
	sqlCmd.append("       	  where agri_loan.bank_no=agri_loan_sum.bank_no(+) ");
	sqlCmd.append("       	    and agri_loan.bank_no <> ' ' "); 
	sqlCmd.append("          GROUP BY agri_loan.m_year,agri_loan.m_month,agri_loan.bank_no,agri_loan.BANK_NAME,field_loan_bal_cnt,field_loan_bal_amt,field_over6m_loan_bal_amt ");
	sqlCmd.append("	) agri_loan,");//--本月份資料
	sqlCmd.append("	( select agri_loan.m_year,agri_loan.m_month, agri_loan.bank_no ,   agri_loan.BANK_NAME,");
	sqlCmd.append("         	 SUM(field_loan_bal_cnt)    as field_loan_bal_cnt ,");
	sqlCmd.append("         	 SUM(field_loan_bal_amt)   as field_loan_bal_amt,");   
	sqlCmd.append("         	 SUM(field_over6m_loan_bal_amt) as field_over6m_loan_bal_amt,");
	sqlCmd.append("         	 decode(SUM(field_loan_bal_amt),0,0,round(SUM(field_over6m_loan_bal_amt) / SUM(field_loan_bal_amt) *100 ,2)) ");  
	sqlCmd.append("		 as   field_over6m_loan_RATE ");//--專案農貸逾放比率=(9)逾期六個月放款餘額-金額/(7)貸放餘額                   
	sqlCmd.append("  	   from ");
	sqlCmd.append("  		(select agri_loan.m_year,agri_loan.m_month, bn01.bank_no , bn01.BANK_NAME,");        
	sqlCmd.append("          		round(sum(agri_loan.loan_bal_cnt) /1,0) as field_loan_bal_cnt,"); //--(6)貸放餘額-戶數
	sqlCmd.append("          		round(sum(agri_loan.loan_bal_amt) /1,0) as field_loan_bal_amt,"); //--(7)貸放餘額
	sqlCmd.append("          		round(sum(agri_loan.over6m_loan_bal_amt) /1,0) as field_over6m_loan_bal_amt ");//--(9)逾期六個月放款餘額-金額          
	sqlCmd.append("   		   from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
	sqlCmd.append("   		   left join (select * from agri_loan  where m_year=? and m_month=?) agri_loan  on  bn01.bank_no = agri_loan.bank_code ");
	paramList.add(wlx01_m_year);
	paramList.add(compareYY);
	paramList.add(compareMM);
	sqlCmd.append("   		  group by agri_loan.m_year,agri_loan.m_month,bn01.bank_no,bn01.BANK_NAME ");
	sqlCmd.append("  		) agri_loan ");
	sqlCmd.append("  	  where agri_loan.bank_no <> ' ' "); 
	sqlCmd.append("  	  GROUP BY agri_loan.m_year,agri_loan.m_month,agri_loan.bank_no,agri_loan.BANK_NAME ");
	sqlCmd.append("	) agri_loan_last ");//--上月份資料
	sqlCmd.append(" where agri_loan.bank_no=agri_loan_last.bank_no(+) ");
	sqlCmd.append(" )a ");
    sqlCmd.append(" where field_diff_loan_bal_amt is not null ");//103.11.25 過濾無資料的機構 by 2295	
	List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,
	        "m_year,m_month,bank_no,bank_name,"
	        +"field_diff_loan_bal_amt,field_diff_over6m_loan_bal_amt,field_diff_over6m_loan_rate,field_delay_loan_rate");
	return dbData;
}
private List getOperationTmpInfo(String m_year,String m_month,String wr_rpt){
    StringBuffer sqlCmd = new StringBuffer();        	
	List paramList = new ArrayList();
	boolean flg = false;
	sqlCmd.append("select * from WR_OPERATION_TMP where m_year= ? and m_month=? and wr_rpt = ? ");
	paramList.add(m_year);
	paramList.add(m_month);
	paramList.add(wr_rpt);
	List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
	return dbData;
}
private boolean haveOperationInfo(String m_year,String m_month,String wr_rpt){
    StringBuffer sqlCmd = new StringBuffer();        	
	List paramList = new ArrayList();
	boolean flg = false;
	sqlCmd.append("select * from WR_OPERATION where m_year= ? and m_month=? and wr_rpt = ? ");
	paramList.add(m_year);
	paramList.add(m_month);
	paramList.add(wr_rpt);
	List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
	if(dbData!=null && dbData.size()>0){
	    flg = true;
	}
	return flg;
}
private boolean haveOperationLogInfo(String m_year,String m_month,String wr_rpt){
    StringBuffer sqlCmd = new StringBuffer();        	
	List paramList = new ArrayList();
	boolean flg = false;
	sqlCmd.append("select * from WR_OPERATION_Log where m_year= ? and m_month=? and report_no=? ");
	paramList.add(m_year);
	paramList.add(m_month);
	paramList.add(wr_rpt);
	List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
	if(dbData!=null && dbData.size()>0){
	    flg = true;
	}
	return flg;
}
private List getRangeInfo(String wr_rpt){
    StringBuffer sqlCmd = new StringBuffer();        	
	List paramList = new ArrayList();				
	sqlCmd.append("select wr_rpt,acc_code,quop,field_amt,remark,serial from wr_range where wr_rpt= ? order by serial ");
	paramList.add(wr_rpt);//0:與上月比較 1:與上季比較 2:與上年度同期比較
	List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"wr_rpt,acc_code,quop,field_amt,remark,serial");
	return dbData;
}
private List getOperTmpRangeInfo(String m_year,String m_month,String acc_code,String quop,String field_amt,String wr_rpt){
    StringBuffer sqlCmd = new StringBuffer();        	
	List paramList = new ArrayList();
	sqlCmd.append("select M_YEAR,M_MONTH,WR_RPT,BANK_CODE,ACC_CODE,TYPE,AMT,WARN_TYPE,REMARK,WR_RANGE_SERIAL ");
	sqlCmd.append("  from WR_OPERATION_TMP where m_year= ? and m_month=? and wr_rpt= ? and acc_code= ? ");
	paramList.add(m_year);
	paramList.add(m_month);
	paramList.add(wr_rpt);
	paramList.add(acc_code);
	if(!"max".equals(quop) && !"min".equals(quop)){
	    sqlCmd.append("  and amt ").append(quop).append(" ").append(field_amt);
	}else{
	    if("max".equals(quop)){
		    sqlCmd.append("  order by amt desc ");
		}else if("min".equals(quop)){
		    sqlCmd.append("  order by amt asc ");
		}
	}
	List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,wr_rpt,bank_code,acc_code,type,amt,warn_type,remark,wr_range_serial");
	return dbData;
}

private String getField_amtValue(String m_year,String m_month,String wr_rpt,String bank_code,String acc_code){
    StringBuffer sqlCmd = new StringBuffer();        	
	List paramList = new ArrayList();
	String field_amtValue="0";
	sqlCmd.append("select amt ");
	sqlCmd.append("  from WR_OPERATION_TMP where m_year= ? and m_month=? and wr_rpt= ? and bank_code = ? and acc_code= ? ");
	paramList.add(m_year);
	paramList.add(m_month);
	paramList.add(wr_rpt);
	paramList.add(bank_code);
	paramList.add(acc_code);
	List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"amt");
	if(dbData!=null && dbData.size()>0){
	    field_amtValue = (((DataObject)dbData.get(0)).getValue("amt")).toString();
	}
	return field_amtValue;
}
private String getField_992710_990230_Warn_Type(String m_year,String m_month,String wr_rpt,String bank_code){
    StringBuffer sqlCmd = new StringBuffer();        	
	List paramList = new ArrayList();
	String warn_type="";
	sqlCmd.append("select warn_type ");
	sqlCmd.append("  from WR_OPERATION_TMP where m_year= ? and m_month=? and wr_rpt= ? and bank_code = ? and acc_code= ? ");
	paramList.add(m_year);
	paramList.add(m_month);
	paramList.add(wr_rpt);
	paramList.add(bank_code);
	paramList.add("field_992710_990230_rate");
	List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"warn_type");
	if(dbData!=null && dbData.size()>0){
	    warn_type = (((DataObject)dbData.get(0)).getValue("warn_type")).toString();
	}
	return warn_type;
}
%>
</body>
</html>