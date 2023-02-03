<%
// 95.09.05 add 將全体農漁會縣市別小計寫入a01_opeartion by 2295
// 95.11.30 add 排程產生A01_opeation by 2295
// 95.12.20 fix A01 SQL by 2295
// 96.04.30 fix 存放比率若為負數,以0顯示 by 2295
// 98.08.25 add 寫入OPERATION_LOG by 2295
// 99.05.04 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
//102.06.20 add 103/01以後.A01漁會.套用新科目代號/計算公式 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.io.*" %>
<%@ page import="java.util.*" %>
<%@ page import="java.lang.Integer" %>
<html>
<head>
<title></title>
</head>
<body>
產生(全体農漁會-縣市別小計)A01_operation
<%
   System.out.println("=============產生A01_Operation開始===========");
   Map dataMap =Utility.saveSearchParameter(request);
   String report_no = Utility.getTrimString(dataMap.get("report_no"));	
   String s_year = Utility.getTrimString(dataMap.get("s_year"));	
   String s_month = Utility.getTrimString(dataMap.get("s_month"));	   
   String isDebug = Utility.getTrimString(dataMap.get("isDebug"));	
   String lguser_id = Utility.getTrimString(dataMap.get("lguser_id"));   
   
   String errMsg = Generate(report_no,s_year,s_month,isDebug,lguser_id);
   System.out.println("errMsg = "+errMsg);
   System.out.println("=============產生A01_Operation結束===========");
%>
<%!
public String Generate(String Report_no,String s_year,String s_month,String isDebug,String lguser_id) throws Exception{
		File logfile;
		FileOutputStream logos=null;
		BufferedOutputStream logbos = null;
		PrintStream logps = null;		
	    File logDir = null;
	    String errMsg="";
        String bank_type="ALL";//全体農漁會
        //99.04.29 add==================================================================
		String cd01_table = "";
    	String wlx01_m_year = "";
    	StringBuffer sqlCmd = new StringBuffer();        	
		List<String> paramList = new ArrayList<String>();				
		List<List> updateDBList = new ArrayList<List>();//0:sql 1:data		
		List updateDBSqlList = new ArrayList();
		List<List> updateDBDataList = new ArrayList<List>();//儲存參數的List
		List<String> dataList =  new ArrayList<String>();//儲存參數的data
		//============================================================================
try{
      logDir  = new File(Utility.getProperties("logDir"));
	  if(!logDir.exists()){
          if(!Utility.mkdirs(Utility.getProperties("logDir"))){
     		  System.out.println("目錄新增失敗");
     	  }
       }

	   logfile = new File(logDir + System.getProperty("file.separator") + Report_no +"_GenerateOperation."+ Utility.getDateFormat("yyyyMMddHHmmss"));
	   System.out.println("logfile filename="+logDir + System.getProperty("file.separator") + Report_no +"_GenerateOperation."+ Utility.getDateFormat("yyyyMMddHHmmss"));
	   logos = new FileOutputStream(logfile,true);
	   logbos = new BufferedOutputStream(logos);
	   logps = new PrintStream(logbos);       
	   logps.println(Utility.getDateFormat("yyyy/MM/dd  HH:mm:ss  ")+" "+"產生"+s_year+"年"+s_month+"月 (全体農漁會-縣市別小計)A01_opeation 資料中");
	   logps.flush();

       //99.04.29 add 查詢年度100年以前.縣市別不同===============================
  	   cd01_table = (Integer.parseInt(s_year) < 100)?"cd01_99":"";
  	   wlx01_m_year = (Integer.parseInt(s_year) < 100)?"99":"100";
  	   //=====================================================================

      //全体農漁會各縣市小計
sqlCmd.append(" select "+s_year+" as m_year,"+s_month+" as m_month,'ALL' as bank_type,hsien_id , hsien_name, FR001W_output_order, ");
sqlCmd.append("       decode(bank_no,' ','ALL',bank_no) as bank_no , bank_name, ");
sqlCmd.append("	      COUNT_SEQ,  ");
sqlCmd.append("       field_SEQ, ");
sqlCmd.append("       field_DEBIT,field_CREDIT, ");
sqlCmd.append("       decode(sign(field_DC_RATE - 0),-1,0,field_DC_RATE) as field_DC_RATE,");//96.04.30 fix 存放比率若為負數,以0顯示
sqlCmd.append("       field_120700,      field_OVER, ");
sqlCmd.append("       field_OVER_RATE, ");
sqlCmd.append("       field_320300,   field_TRANSFER, ");
sqlCmd.append("       field_TRANSFER_RATE, ");
sqlCmd.append("       field_310000,   field_NET, ");
sqlCmd.append("       field_FIXNET_RATE, ");
sqlCmd.append("       field_CHECK_RATE, ");
sqlCmd.append("       field_150200,   field_BACKUP,   field_NOASSURE, ");
sqlCmd.append("       field_MODIFYNET, ");
sqlCmd.append("       field_CAPTIAL_RATE ");
sqlCmd.append(" from (select a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order,a01.bank_no , a01.bank_name, COUNT_SEQ, ");
sqlCmd.append("      		 field_SEQ, ");
sqlCmd.append("      		 round(field_DEBIT /1,0)  as field_DEBIT, ");
sqlCmd.append("      		 round(field_CREDIT /1,0)  as field_CREDIT, ");
sqlCmd.append("         	 	 decode(a01.fieldI_Y,0,0, ");
sqlCmd.append("          	 round( ");
sqlCmd.append("               (a01.fieldI_XA                                          + ");
sqlCmd.append("               decode(sign(a01.fieldI_XB1 - a01.fieldI_XB2),-1,0, ");
sqlCmd.append("                          (a01.fieldI_XB1 - a01.fieldI_XB2))           + ");
sqlCmd.append("               decode(sign(a01.fieldI_XC1 - a01.fieldI_XC2),-1,0, ");
sqlCmd.append("                          (a01.fieldI_XC1 - a01.fieldI_XC2))           + ");
sqlCmd.append("               decode(sign(a01.fieldI_XD1 - a01.fieldI_XD2),-1,0, ");
sqlCmd.append("                          (a01.fieldI_XD1 - a01.fieldI_XD2))           + ");
sqlCmd.append("               decode(sign(a01.fieldI_XE1 - a01.fieldI_XE2),-1,0, ");
sqlCmd.append("                          (a01.fieldI_XE1 - a01.fieldI_XE2))           - ");
sqlCmd.append("               decode(sign(a01.fieldI_XF1 - a01.fieldI_XF3 -  a01.fieldI_XF2),-1,0, ");
sqlCmd.append("                          (a01.fieldI_XF1 - a01.fieldI_XF3 - a01.fieldI_XF2)) ");
sqlCmd.append("               ) ");
sqlCmd.append("               /    a01.fieldI_Y * 100,2))        as     field_DC_RATE , ");
sqlCmd.append("      		 round(field_120700 /1,0)  as field_120700, ");
sqlCmd.append("      		 round(field_OVER /1,0)  as field_OVER, ");
sqlCmd.append("         	 decode(a01.field_CREDIT,0,0,round(a01.field_OVER /  a01.field_CREDIT *100 ,2))  as   field_OVER_RATE, ");
sqlCmd.append("      		 round(field_320300 /1,0)  as field_320300, ");
sqlCmd.append("              round(field_TRANSFER /1,0)  as field_TRANSFER, ");
sqlCmd.append("            	 decode(field_DEPOSITBANK,0,0,round(a01.field_DEPOSITBANK_AA / a01.field_DEPOSITBANK *100 ,2))  as   field_TRANSFER_RATE, ");
sqlCmd.append("      		 round(field_310000 /1,0)  as field_310000, ");
sqlCmd.append("      		 round(field_NET /1,0)  as field_NET,decode(field_NET,0,0,round(a01.field_140000 /  a01.field_NET *100 ,2))  as   field_FIXNET_RATE, ");
sqlCmd.append("         	 decode(field_DEBIT,0,0,round(a01.field_CHECK /  a01.field_DEBIT *100 ,2))  as   field_CHECK_RATE, ");
sqlCmd.append("      		 round(field_150200 /1,0)  as field_150200, ");
sqlCmd.append("      		 round(field_BACKUP /1,0)  as field_BACKUP, ");
sqlCmd.append("      		 round(field_NOASSURE /1,0)  as field_NOASSURE, ");
sqlCmd.append("      		 round(a01.field_MODIFYNET / COUNT_SEQ,0)  as   field_MODIFYNET, ");
sqlCmd.append("      		 round(field_CAPTIAL / COUNT_SEQ /  1000  ,2)  as   field_CAPTIAL_RATE ");
sqlCmd.append("	  from (select  ' '  AS  hsien_id ,  ' 總計 '   AS hsien_name,  '001'  AS FR001W_output_order,'ALL' as bank_type, ");
sqlCmd.append("	     			' ' AS  bank_no ,     ' '   AS  bank_name, ");
sqlCmd.append("	      		    COUNT(*)  AS  COUNT_SEQ, ");
sqlCmd.append("					'A99'  as  field_SEQ, ");
sqlCmd.append("                  SUM(field_120700)   	field_120700 , ");
sqlCmd.append("                  SUM(field_320300)   	field_320300 , ");
sqlCmd.append("                  SUM(field_TRANSFER)  	field_TRANSFER , ");
sqlCmd.append("                  SUM(field_DEPOSITBANK_AA) field_DEPOSITBANK_AA, ");
sqlCmd.append("                  SUM(field_DEPOSITBANK)  field_DEPOSITBANK, ");
sqlCmd.append("                  SUM(field_310000)  	field_310000, ");
sqlCmd.append("                  SUM(field_NET)       	field_NET, ");
sqlCmd.append("                  SUM(field_140000)    	field_140000, ");
sqlCmd.append("                  SUM(field_CHECK)       field_CHECK, ");
sqlCmd.append("                  SUM(field_150200)     	field_150200, ");
sqlCmd.append("                  SUM(field_BACKUP)     	field_BACKUP, ");
sqlCmd.append("                  SUM(field_NOASSURE)   	field_NOASSURE, ");
sqlCmd.append("                  SUM(field_MODIFYNET)  	field_MODIFYNET,");
sqlCmd.append("                  SUM(field_OVER)        field_OVER,  ");
sqlCmd.append("                  SUM(field_DEBIT)       field_DEBIT, ");
sqlCmd.append("                  SUM(field_CREDIT)      field_CREDIT,");
sqlCmd.append("                  SUM(fieldI_XA)         fieldI_XA,   ");
sqlCmd.append("                  SUM(fieldI_XB1)        fieldI_XB1,  ");
sqlCmd.append("                  SUM(fieldI_XB2)        fieldI_XB2,  ");
sqlCmd.append("                  SUM(fieldI_XC1)        fieldI_XC1,  ");
sqlCmd.append("                  SUM(fieldI_XC2)        fieldI_XC2,  ");
sqlCmd.append("                  SUM(fieldI_XD1)        fieldI_XD1,  ");
sqlCmd.append("                  SUM(fieldI_XD2)        fieldI_XD2,  ");
sqlCmd.append("                  SUM(fieldI_XE1)        fieldI_XE1,  ");
sqlCmd.append("                  SUM(fieldI_XE2)        fieldI_XE2,  ");
sqlCmd.append("                  SUM(fieldI_XF1)        fieldI_XF1,  ");
sqlCmd.append("                  SUM(fieldI_XF3)        fieldI_XF3,  ");
sqlCmd.append("                  SUM(fieldI_XF2)        fieldI_XF2,  ");
sqlCmd.append("                  SUM(fieldI_Y)          fieldI_Y,    ");
sqlCmd.append("                  SUM(nvl(a05.amt,0))  as field_CAPTIAL ");
sqlCmd.append(" 			from ( select nvl(cd01.hsien_id,' ')       as  hsien_id , ");
sqlCmd.append("                           nvl(cd01.hsien_name,'OTHER') as  hsien_name,");
sqlCmd.append("                           cd01.FR001W_output_order     as  FR001W_output_order, ");
sqlCmd.append("                           bn01.bank_no ,  bn01.BANK_NAME, ");
sqlCmd.append("   				          round(sum(decode(a01.acc_code,'120700',amt,0)) /1,0)     as field_120700, ");  
sqlCmd.append("                           round(sum(decode(a01.acc_code,'320300',amt,0)) /1,0)     as field_320300, "); 
sqlCmd.append("					          round(sum( decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'110324',amt,'110325',amt,0),'7',decode(a01.acc_code,'110254',amt,'110255',amt,0)), ");
sqlCmd.append("	           		                                       '103',decode(a01.acc_code,'110324',amt,'110325',amt,0),0) ) /1,0)     as field_TRANSFER, "); 
sqlCmd.append("	            	          round(sum( decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'110320',amt,0),'7',decode(a01.acc_code,'110250',amt,0)), ");
sqlCmd.append("	            	                                      '103',decode(a01.acc_code,'110320',amt,0),0)) /1,0)     as field_DEPOSITBANK_AA,   ");
sqlCmd.append("	            	          round(sum( decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'110300',amt,0),'7',decode(a01.acc_code,'110200',amt,0)), ");
sqlCmd.append("					                                      '103',decode(a01.acc_code,'110300',amt,0),0) ) /1,0)     as field_DEPOSITBANK, ");
sqlCmd.append("					          round(sum(decode(a01.acc_code,'310000',amt,0)) /1,0)     as field_310000,   ");
sqlCmd.append("					          round(sum(decode(bank_type,'6',decode(a01.acc_code,'310000',amt,'320000',amt,0),'7',decode(a01.acc_code,'300000',amt,0),0)) /1,0)     as field_NET,  "); 
sqlCmd.append("					          round(sum(decode(a01.acc_code,'140000',amt,0)) /1,0)     as field_140000,   ");
sqlCmd.append("               	          round(sum(decode(a01.acc_code,'220100',amt, '220200',amt, '220300',amt, '220400',amt, '220500',amt,0)) /1,0) as field_CHECK,   ");
sqlCmd.append("               	          round(sum(decode(a01.acc_code,'150200',amt,0)) /1,0)     as field_150200,   ");
sqlCmd.append("                	          round(sum(decode(a01.acc_code, '120800',amt,'150300',amt,0)) /1,0) as  field_BACKUP, ");                            
sqlCmd.append("                           round(sum( decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code, '120101',amt,'120301',amt, '120401',amt, '120501',amt,0),'7',decode(a01.acc_code, '120101',amt,'120401', amt, '120201',amt, '120501',amt,0)),  ");
sqlCmd.append("                                                       '103',decode(a01.acc_code, '120101',amt,'120301',amt, '120401',amt, '120501',amt,0),0) ) /1,0) as  field_NOASSURE,  ");
sqlCmd.append("                           round((sum(decode(bank_type,'6',decode(a01.acc_code, '310000',amt,'320000',amt, '120800',amt,'150300',amt,0),'7',decode(a01.acc_code, '300000',amt, '120800',amt, '150300',amt,0))) -   ");
sqlCmd.append("              	          round(sum(decode(a01.acc_code,'990000',amt,0)) * 1.25 * 0.7,0))/1,0) as  field_MODIFYNET,   ");
sqlCmd.append("   				          round(sum(decode(a01.acc_code,'990000',amt,0)) /1,0)        as field_OVER,   ");
sqlCmd.append("                           round(sum(decode(a01.acc_code,'220000',amt,0)) /1,0)        as field_DEBIT,  "); 
sqlCmd.append("					          round(sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /1,0) as  field_CREDIT,    ");                           
sqlCmd.append("	           		          round(sum(decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'120101',amt,'120102',amt,  ");
sqlCmd.append("	            	                                                                        '120200',amt,'120301',amt,'120302',amt,'120700',amt,'150200',amt,0)  "); 
sqlCmd.append("	            	                                                           ,'7',decode(a01.acc_code,'120101',amt,'120102',amt,  ");
sqlCmd.append("	            	                                                                        '120300',amt,'120401',amt,'120402',amt,'120700',amt,'150200',amt,0)), ");
sqlCmd.append("					                                     '103',decode(a01.acc_code,'120101',amt,'120102',amt,'120200',amt,'120301',amt, "); 
sqlCmd.append("					                                                                '120302',amt,'120700',amt,'150200',amt,0),0)) /1,0)     as fieldI_XA,  "); 
sqlCmd.append("					          round(sum(decode(year_type,'102',decode(bank_type,'6',decode(a01.acc_code,'120401',amt,'120402',amt,0),'7',decode(a01.acc_code,'120201',amt,'120202',amt,0)), ");
sqlCmd.append("					                                     '103',decode(a01.acc_code,'120401',amt,'120402',amt,0),0)) /1,0)     as fieldI_XB1, ");
sqlCmd.append("               	          sum(decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'240305',amt,'251300',amt,'310800',amt, 0),'7',decode(a01.acc_code,'240205',amt, '310800',amt,0)), ");
sqlCmd.append("               	                               '103',decode(bank_type,'6',decode(a01.acc_code,'240305',amt,'251300',amt,'310800',amt, 0),'7',decode(a01.acc_code,'240305',amt, '251200',amt,0)),0))  as fieldI_XB2, ");
sqlCmd.append("                	          round(sum(decode(a01.acc_code,'120501',amt,'120502',amt,0)) /1,0)     as fieldI_XC1, ");  
sqlCmd.append("                           round(sum( decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'240301',amt,'240302',amt,'240303',amt,'240304',amt,0), ");  
sqlCmd.append("                                                                              '7',decode(a01.acc_code,'240201',amt,'240202',amt,'240203',amt,'240204',amt,0)), ");
sqlCmd.append("                                                       '103', decode(a01.acc_code,'240301',amt,'240302',amt,'240303',amt,'240304',amt,0),0)) /1,0)     as fieldI_XC2, ");
sqlCmd.append("              	          round(sum(decode(a01.acc_code,'120600',amt,0)) /1,0)              as fieldI_XD1, "); 
sqlCmd.append("   				          round(sum( decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'240200',amt,0), ");
sqlCmd.append("                                                                              '7',decode(a01.acc_code,'240300',amt,0)), ");
sqlCmd.append("					                                      '103',decode(a01.acc_code,'240200',amt,0),0) ) /1,0)   as fieldI_XD2,  ");
sqlCmd.append("	           		          round(sum(decode(a01.acc_code,'150100',amt,0)) /1,0)              as fieldI_XE1,  "); 
sqlCmd.append("	            	          round(sum(decode(a01.acc_code,'250100',amt,0)) /1,0)              as fieldI_XE2,  "); 
sqlCmd.append("	            	          round(sum(decode(a01.acc_code,'310000',amt,'320000',amt,0)) /1,0) as fieldI_XF1,  "); 
sqlCmd.append("	            	          round(sum( decode(YEAR_TYPE,'102',decode(a01.acc_code,'310800',amt,0), ");
sqlCmd.append("					                                      '103',decode(bank_type,'6',decode(a01.acc_code,'310800',amt,0),'7',0,0),0)) /1,0)  as fieldI_XF3, ");  
sqlCmd.append("					          round(sum(decode(a01.acc_code,'140000',amt,0)) /1,0)              as fieldI_XF2,  "); 
sqlCmd.append("					          round((sum(decode(a01.acc_code,'220100',amt,'220200',amt,  "); 
sqlCmd.append("					                                         '220300',amt,'220400',amt,  "); 
sqlCmd.append("               	                                         '220500',amt,'220600',amt,  "); 
sqlCmd.append("               	                                         '220700',amt,'220800',amt,   ");
sqlCmd.append("                	                                         '220900',amt,'221000',amt,0))-  ");
sqlCmd.append("                           round(sum(decode(a01.acc_code,'220900',amt,0))/2,0)) /1,0)   as fieldI_Y "); 
sqlCmd.append("                    from (select * from "+cd01_table+" cd01 where cd01.hsien_id <> 'Y') cd01 ");
sqlCmd.append("                    left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id ");
sqlCmd.append("                    left join (select * from bn01 where m_year=?)bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type in ('6','7') ");
sqlCmd.append("                    left join (select (CASE WHEN (a01.m_year <= 102) THEN '102' ");                                 
sqlCmd.append("                                            WHEN (a01.m_year > 102) THEN '103'");                                   
sqlCmd.append("                                            ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01 ");
sqlCmd.append("                               where  a01.m_year = ? and a01.m_month = ?) a01  on  bn01.bank_no = a01.bank_code ");
paramList.add(wlx01_m_year);
paramList.add(wlx01_m_year);
paramList.add(s_year);
paramList.add(s_month);
sqlCmd.append("                    group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no ,  bn01.BANK_NAME ");
sqlCmd.append("                 ) a01, (select * from a05 where a05.m_year=? and a05.m_month = ? and  a05.ACC_code = '91060P') a05 ");
paramList.add(s_year);
paramList.add(s_month);
sqlCmd.append("			where a01.bank_no=a05.bank_code(+)  and a01.bank_no <> ' ' ");
sqlCmd.append("			) a01 ");
sqlCmd.append(" 	  UNION ALL ");
sqlCmd.append(" 	  		select a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order, ");
sqlCmd.append("            	       a01.bank_no , a01.BANK_NAME,COUNT_SEQ, ");
sqlCmd.append("      		       field_SEQ, ");
sqlCmd.append("      		       round(field_DEBIT /1,0)  as field_DEBIT, ");
sqlCmd.append("      		       round(field_CREDIT /1,0)  as field_CREDIT,");
sqlCmd.append("         		   decode(a01.fieldI_Y,0,0, ");
sqlCmd.append("          	       round( ");
sqlCmd.append("                    		(a01.fieldI_XA                                          + ");
sqlCmd.append("                       		 decode(sign(a01.fieldI_XB1 - a01.fieldI_XB2),-1,0, ");
sqlCmd.append("                             		    (a01.fieldI_XB1 - a01.fieldI_XB2))           + ");
sqlCmd.append("                    		 decode(sign(a01.fieldI_XC1 - a01.fieldI_XC2),-1,0, ");
sqlCmd.append("                             			(a01.fieldI_XC1 - a01.fieldI_XC2))           + ");
sqlCmd.append("                    		 decode(sign(a01.fieldI_XD1 - a01.fieldI_XD2),-1,0, ");
sqlCmd.append("                             		    (a01.fieldI_XD1 - a01.fieldI_XD2))           + ");
sqlCmd.append("                    		 decode(sign(a01.fieldI_XE1 - a01.fieldI_XE2),-1,0, ");
sqlCmd.append("                             		    (a01.fieldI_XE1 - a01.fieldI_XE2))           - ");
sqlCmd.append("                    		 decode(sign(a01.fieldI_XF1 - a01.fieldI_XF3 -  a01.fieldI_XF2),-1,0, ");
sqlCmd.append("                             		    (a01.fieldI_XF1 - a01.fieldI_XF3 - a01.fieldI_XF2)) ");
sqlCmd.append("                     ) /    a01.fieldI_Y * 100,2))        as     field_DC_RATE , ");
sqlCmd.append("      		        round(field_120700 /1,0)  as field_120700, ");
sqlCmd.append("      		        round(field_OVER /1,0)  as field_OVER, ");
sqlCmd.append("         		    decode(a01.field_CREDIT,0,0,round(a01.field_OVER /  a01.field_CREDIT *100 ,2))  as   field_OVER_RATE, ");
sqlCmd.append("      		        round(field_320300 /1,0)  as field_320300, ");
sqlCmd.append("      		        round(field_TRANSFER /1,0)  as field_TRANSFER, ");
sqlCmd.append("        		        decode(field_DEPOSITBANK,0,0,round(a01.field_DEPOSITBANK_AA / a01.field_DEPOSITBANK *100 ,2))  as   field_TRANSFER_RATE, ");
sqlCmd.append("      		        round(field_310000 /1,0)  as field_310000, ");
sqlCmd.append("      		        round(field_NET /1,0)  as field_NET, ");
sqlCmd.append("         		    decode(field_NET,0,0,round(a01.field_140000 /  a01.field_NET *100 ,2))  as   field_FIXNET_RATE, ");
sqlCmd.append("         		    decode(field_DEBIT,0,0,round(a01.field_CHECK /  a01.field_DEBIT *100 ,2))  as   field_CHECK_RATE, ");
sqlCmd.append("      		        round(field_150200 /1,0)  as field_150200, ");
sqlCmd.append("      		        round(field_BACKUP /1,0)  as field_BACKUP, ");
sqlCmd.append("      		        round(field_NOASSURE /1,0)  as field_NOASSURE, ");
sqlCmd.append("	  			        round(field_MODIFYNET /1 ,0)  as field_MODIFYNET, ");
sqlCmd.append("   	  		        round(field_CAPTIAL /  1000 ,2)  as   field_CAPTIAL_RATE ");
sqlCmd.append(" 		   from (select a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order, ");
sqlCmd.append("      		     		a01.bank_no,a01.bank_name, ");
sqlCmd.append("                 		1  AS  COUNT_SEQ, ");
sqlCmd.append(" 				 		'A01'  as  field_SEQ, ");
sqlCmd.append("                         SUM(field_120700)   	field_120700 , ");
sqlCmd.append("                         SUM(field_320300)   	field_320300 , ");
sqlCmd.append("                         SUM(field_TRANSFER)  	field_TRANSFER , ");
sqlCmd.append("                         SUM(field_DEPOSITBANK_AA) field_DEPOSITBANK_AA, ");
sqlCmd.append("                         SUM(field_DEPOSITBANK)  field_DEPOSITBANK, ");
sqlCmd.append("                         SUM(field_310000)  	   	field_310000, ");
sqlCmd.append("                         SUM(field_NET)       	field_NET, ");
sqlCmd.append("                         SUM(field_140000)    	field_140000, ");
sqlCmd.append("                         SUM(field_CHECK)        field_CHECK, ");
sqlCmd.append("                         SUM(field_150200)     	field_150200, ");
sqlCmd.append("                         SUM(field_BACKUP)     	field_BACKUP, ");
sqlCmd.append("                         SUM(field_NOASSURE)   	field_NOASSURE, ");
sqlCmd.append("                         SUM(field_MODIFYNET)  	field_MODIFYNET, ");
sqlCmd.append("                         SUM(field_OVER)         field_OVER, ");
sqlCmd.append("                         SUM(field_DEBIT)        field_DEBIT, ");
sqlCmd.append("                         SUM(field_CREDIT)       field_CREDIT, ");
sqlCmd.append("                         SUM(fieldI_XA)          fieldI_XA, ");
sqlCmd.append("                         SUM(fieldI_XB1)         fieldI_XB1, ");
sqlCmd.append("                         SUM(fieldI_XB2)         fieldI_XB2, ");
sqlCmd.append("                         SUM(fieldI_XC1)         fieldI_XC1, ");
sqlCmd.append("                         SUM(fieldI_XC2)         fieldI_XC2, ");
sqlCmd.append("                         SUM(fieldI_XD1)         fieldI_XD1, ");
sqlCmd.append("                         SUM(fieldI_XD2)         fieldI_XD2, ");
sqlCmd.append("                         SUM(fieldI_XE1)         fieldI_XE1, ");
sqlCmd.append("                         SUM(fieldI_XE2)         fieldI_XE2, ");
sqlCmd.append("                         SUM(fieldI_XF1)         fieldI_XF1, ");
sqlCmd.append("                         SUM(fieldI_XF3)         fieldI_XF3, ");
sqlCmd.append("                         SUM(fieldI_XF2)         fieldI_XF2, ");
sqlCmd.append("                         SUM(fieldI_Y)           fieldI_Y, ");
sqlCmd.append("                 		SUM(nvl(a05.amt,0))  as field_CAPTIAL ");
sqlCmd.append(" 				from ( select nvl(cd01.hsien_id,' ')       as  hsien_id , ");
sqlCmd.append("        					      nvl(cd01.hsien_name,'OTHER')  as  hsien_name, ");
sqlCmd.append("        					      cd01.FR001W_output_order     as  FR001W_output_order, ");
sqlCmd.append("        					      bn01.bank_no ,  bn01.bank_name, ");
sqlCmd.append("   				              round(sum(decode(a01.acc_code,'120700',amt,0)) /1,0)     as field_120700, ");  
sqlCmd.append("                               round(sum(decode(a01.acc_code,'320300',amt,0)) /1,0)     as field_320300, "); 
sqlCmd.append("					              round(sum( decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'110324',amt,'110325',amt,0),'7',decode(a01.acc_code,'110254',amt,'110255',amt,0)), ");
sqlCmd.append("	           		                                           '103',decode(a01.acc_code,'110324',amt,'110325',amt,0),0) ) /1,0)     as field_TRANSFER, "); 
sqlCmd.append("	            	              round(sum( decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'110320',amt,0),'7',decode(a01.acc_code,'110250',amt,0)), ");
sqlCmd.append("	            	                                          '103',decode(a01.acc_code,'110320',amt,0),0)) /1,0)     as field_DEPOSITBANK_AA,   ");
sqlCmd.append("	            	              round(sum( decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'110300',amt,0),'7',decode(a01.acc_code,'110200',amt,0)), ");
sqlCmd.append("					                                          '103',decode(a01.acc_code,'110300',amt,0),0) ) /1,0)     as field_DEPOSITBANK, ");
sqlCmd.append("					              round(sum(decode(a01.acc_code,'310000',amt,0)) /1,0)     as field_310000,   ");
sqlCmd.append("					              round(sum(decode(bank_type,'6',decode(a01.acc_code,'310000',amt,'320000',amt,0),'7',decode(a01.acc_code,'300000',amt,0),0)) /1,0)     as field_NET,  "); 
sqlCmd.append("					              round(sum(decode(a01.acc_code,'140000',amt,0)) /1,0)     as field_140000,   ");
sqlCmd.append("               	              round(sum(decode(a01.acc_code,'220100',amt, '220200',amt, '220300',amt, '220400',amt, '220500',amt,0)) /1,0) as field_CHECK,   ");
sqlCmd.append("               	              round(sum(decode(a01.acc_code,'150200',amt,0)) /1,0)     as field_150200,   ");
sqlCmd.append("                	              round(sum(decode(a01.acc_code, '120800',amt,'150300',amt,0)) /1,0) as  field_BACKUP, ");                            
sqlCmd.append("                               round(sum( decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code, '120101',amt,'120301',amt, '120401',amt, '120501',amt,0),'7',decode(a01.acc_code, '120101',amt,'120401', amt, '120201',amt, '120501',amt,0)),  ");
sqlCmd.append("                                                           '103',decode(a01.acc_code, '120101',amt,'120301',amt, '120401',amt, '120501',amt,0),0) ) /1,0) as  field_NOASSURE,  ");
sqlCmd.append("                               round((sum(decode(bank_type,'6',decode(a01.acc_code, '310000',amt,'320000',amt, '120800',amt,'150300',amt,0),'7',decode(a01.acc_code, '300000',amt, '120800',amt, '150300',amt,0))) -   ");
sqlCmd.append("              	              round(sum(decode(a01.acc_code,'990000',amt,0)) * 1.25 * 0.7,0))/1,0) as  field_MODIFYNET,   ");
sqlCmd.append("   				              round(sum(decode(a01.acc_code,'990000',amt,0)) /1,0)        as field_OVER,   ");
sqlCmd.append("                               round(sum(decode(a01.acc_code,'220000',amt,0)) /1,0)        as field_DEBIT,  "); 
sqlCmd.append("					              round(sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /1,0) as  field_CREDIT,    ");                           
sqlCmd.append("	           		              round(sum(decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'120101',amt,'120102',amt,  ");
sqlCmd.append("	            	                                                                            '120200',amt,'120301',amt,'120302',amt,'120700',amt,'150200',amt,0)  "); 
sqlCmd.append("	            	                                                               ,'7',decode(a01.acc_code,'120101',amt,'120102',amt,  ");
sqlCmd.append("	            	                                                                            '120300',amt,'120401',amt,'120402',amt,'120700',amt,'150200',amt,0)), ");
sqlCmd.append("					                                         '103',decode(a01.acc_code,'120101',amt,'120102',amt,'120200',amt,'120301',amt, "); 
sqlCmd.append("					                                                                    '120302',amt,'120700',amt,'150200',amt,0),0)) /1,0)     as fieldI_XA,  "); 
sqlCmd.append("					              round(sum(decode(year_type,'102',decode(bank_type,'6',decode(a01.acc_code,'120401',amt,'120402',amt,0),'7',decode(a01.acc_code,'120201',amt,'120202',amt,0)), ");
sqlCmd.append("					                                         '103',decode(a01.acc_code,'120401',amt,'120402',amt,0),0)) /1,0)     as fieldI_XB1, ");
sqlCmd.append("               	              sum(decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'240305',amt,'251300',amt,'310800',amt, 0),'7',decode(a01.acc_code,'240205',amt, '310800',amt,0)), ");
sqlCmd.append("               	                                   '103',decode(bank_type,'6',decode(a01.acc_code,'240305',amt,'251300',amt,'310800',amt, 0),'7',decode(a01.acc_code,'240305',amt, '251200',amt,0)),0))  as fieldI_XB2, ");
sqlCmd.append("                	              round(sum(decode(a01.acc_code,'120501',amt,'120502',amt,0)) /1,0)     as fieldI_XC1, ");  
sqlCmd.append("                               round(sum( decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'240301',amt,'240302',amt,'240303',amt,'240304',amt,0), ");  
sqlCmd.append("                                                                                  '7',decode(a01.acc_code,'240201',amt,'240202',amt,'240203',amt,'240204',amt,0)), ");
sqlCmd.append("                                                           '103', decode(a01.acc_code,'240301',amt,'240302',amt,'240303',amt,'240304',amt,0),0)) /1,0)     as fieldI_XC2, ");
sqlCmd.append("              	              round(sum(decode(a01.acc_code,'120600',amt,0)) /1,0)              as fieldI_XD1, "); 
sqlCmd.append("   				              round(sum( decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'240200',amt,0), ");
sqlCmd.append("                                                                                  '7',decode(a01.acc_code,'240300',amt,0)), ");
sqlCmd.append("					                                          '103',decode(a01.acc_code,'240200',amt,0),0) ) /1,0)   as fieldI_XD2,  ");
sqlCmd.append("	           		              round(sum(decode(a01.acc_code,'150100',amt,0)) /1,0)              as fieldI_XE1,  "); 
sqlCmd.append("	            	              round(sum(decode(a01.acc_code,'250100',amt,0)) /1,0)              as fieldI_XE2,  "); 
sqlCmd.append("	            	              round(sum(decode(a01.acc_code,'310000',amt,'320000',amt,0)) /1,0) as fieldI_XF1,  "); 
sqlCmd.append("	            	              round(sum( decode(YEAR_TYPE,'102',decode(a01.acc_code,'310800',amt,0), ");
sqlCmd.append("					                                          '103',decode(bank_type,'6',decode(a01.acc_code,'310800',amt,0),'7',0,0),0)) /1,0)  as fieldI_XF3, ");  
sqlCmd.append("					              round(sum(decode(a01.acc_code,'140000',amt,0)) /1,0)              as fieldI_XF2,  "); 
sqlCmd.append("					              round((sum(decode(a01.acc_code,'220100',amt,'220200',amt,  "); 
sqlCmd.append("					                                             '220300',amt,'220400',amt,  "); 
sqlCmd.append("               	                                             '220500',amt,'220600',amt,  "); 
sqlCmd.append("               	                                             '220700',amt,'220800',amt,   ");
sqlCmd.append("                	                                             '220900',amt,'221000',amt,0))-  ");
sqlCmd.append("                               round(sum(decode(a01.acc_code,'220900',amt,0))/2,0)) /1,0)   as fieldI_Y "); 
sqlCmd.append("     			  from  (select * from "+cd01_table+" cd01 where cd01.hsien_id <> 'Y') cd01 ");
sqlCmd.append("        			  left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id ");
sqlCmd.append("        			  left join (select * from bn01 where m_year=?)bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type in ('6','7')");
sqlCmd.append("                   left join (select (CASE WHEN (a01.m_year <= 102) THEN '102' ");                                 
sqlCmd.append("                                           WHEN (a01.m_year > 102) THEN '103'");                                   
sqlCmd.append("                                           ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01 ");
sqlCmd.append("                              where  a01.m_year = ? and a01.m_month = ?) a01  on  bn01.bank_no = a01.bank_code ");
paramList.add(wlx01_m_year);
paramList.add(wlx01_m_year);
paramList.add(s_year);
paramList.add(s_month);
sqlCmd.append("        			  group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no ,  bn01.BANK_NAME ");
sqlCmd.append("				) a01, (select * from a05 where a05.m_year= ? and a05.m_month  = ? and  a05.ACC_code = '91060P') a05 ");
paramList.add(s_year);
paramList.add(s_month);
sqlCmd.append("			 where   a01.bank_no=a05.bank_code(+)  and a01.bank_no <> ' ' ");
sqlCmd.append("			GROUP  BY a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order,a01.bank_no , a01.BANK_NAME ");
sqlCmd.append("		) a01 ");
sqlCmd.append("	  UNION ALL ");
sqlCmd.append("      select a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order, ");
sqlCmd.append("             a01.bank_no ,         a01.bank_name, ");
sqlCmd.append("             COUNT_SEQ, ");
sqlCmd.append("             field_SEQ, ");
sqlCmd.append("      	   round(field_DEBIT /1,0)  as field_DEBIT, ");
sqlCmd.append("      	   round(field_CREDIT /1,0)  as field_CREDIT, ");
sqlCmd.append("         	   decode(a01.fieldI_Y,0,0, ");
sqlCmd.append("             round( ");
sqlCmd.append("                (a01.fieldI_XA                                          + ");
sqlCmd.append("                   decode(sign(a01.fieldI_XB1 - a01.fieldI_XB2),-1,0, ");
sqlCmd.append("                           (a01.fieldI_XB1 - a01.fieldI_XB2))           + ");
sqlCmd.append("                decode(sign(a01.fieldI_XC1 - a01.fieldI_XC2),-1,0, ");
sqlCmd.append("                           (a01.fieldI_XC1 - a01.fieldI_XC2))           + ");
sqlCmd.append("                decode(sign(a01.fieldI_XD1 - a01.fieldI_XD2),-1,0, ");
sqlCmd.append("                           (a01.fieldI_XD1 - a01.fieldI_XD2))           + ");
sqlCmd.append("                decode(sign(a01.fieldI_XE1 - a01.fieldI_XE2),-1,0, ");
sqlCmd.append("                           (a01.fieldI_XE1 - a01.fieldI_XE2))           - ");
sqlCmd.append("                decode(sign(a01.fieldI_XF1 - a01.fieldI_XF3 -  a01.fieldI_XF2),-1,0, ");
sqlCmd.append("                           (a01.fieldI_XF1 - a01.fieldI_XF3 - a01.fieldI_XF2)) ");
sqlCmd.append("                ) ");
sqlCmd.append("               /    a01.fieldI_Y * 100,2))        as     field_DC_RATE , ");
sqlCmd.append("      	    round(field_120700 /1,0)  as field_120700, ");
sqlCmd.append("      	    round(field_OVER /1,0)  as field_OVER,decode(a01.field_CREDIT,0,0,round(a01.field_OVER /  a01.field_CREDIT *100 ,2))  as   field_OVER_RATE, ");
sqlCmd.append("      	    round(field_320300 /1,0)  as field_320300, ");
sqlCmd.append("      	    round(field_TRANSFER /1,0)  as field_TRANSFER, ");
sqlCmd.append("        		decode(field_DEPOSITBANK,0,0,round(a01.field_DEPOSITBANK_AA / a01.field_DEPOSITBANK *100 ,2))  as   field_TRANSFER_RATE, ");
sqlCmd.append("      		round(field_310000 /1,0)  as field_310000, ");
sqlCmd.append("      		round(field_NET /1,0)  as field_NET,decode(field_NET,0,0,round(a01.field_140000 /  a01.field_NET *100 ,2))  as   field_FIXNET_RATE, ");
sqlCmd.append("         		decode(field_DEBIT,0,0,round(a01.field_CHECK /  a01.field_DEBIT *100 ,2))  as   field_CHECK_RATE, ");
sqlCmd.append("      		round(field_150200 /1,0)  as field_150200, ");
sqlCmd.append("      		round(field_BACKUP /1,0)  as field_BACKUP, ");
sqlCmd.append("      		round(field_NOASSURE /1,0)  as field_NOASSURE, ");
sqlCmd.append("    			round(a01.field_MODIFYNET / COUNT_SEQ,0)  as   field_MODIFYNET, ");
sqlCmd.append("    			round(field_CAPTIAL / COUNT_SEQ /  1000  ,2)  as   field_CAPTIAL_RATE ");
sqlCmd.append("       from ( ");
sqlCmd.append("        	select a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order, ");
sqlCmd.append("            	   ' ' AS  bank_no ,     ' '   AS  bank_name, ");
sqlCmd.append("             	   COUNT(*)  AS  COUNT_SEQ, ");
sqlCmd.append("        		   'A90'  as  field_SEQ, ");
sqlCmd.append("                 SUM(field_120700)   	field_120700 , ");
sqlCmd.append("                 SUM(field_320300)   	field_320300 , ");
sqlCmd.append("                 SUM(field_TRANSFER)  	field_TRANSFER , ");
sqlCmd.append("                 SUM(field_DEPOSITBANK_AA) field_DEPOSITBANK_AA, ");
sqlCmd.append("                 SUM(field_DEPOSITBANK)  field_DEPOSITBANK, ");
sqlCmd.append("                 SUM(field_310000)  	   	field_310000, ");
sqlCmd.append("                 SUM(field_NET)       	field_NET, ");
sqlCmd.append("                 SUM(field_140000)    	field_140000, ");
sqlCmd.append("                 SUM(field_CHECK)        field_CHECK, ");
sqlCmd.append("                 SUM(field_150200)     	field_150200, ");
sqlCmd.append("                 SUM(field_BACKUP)     	field_BACKUP, ");
sqlCmd.append("                 SUM(field_NOASSURE)   	field_NOASSURE, ");
sqlCmd.append("                 SUM(field_MODIFYNET)  	field_MODIFYNET, ");
sqlCmd.append("                 SUM(field_OVER)         field_OVER, ");
sqlCmd.append("                 SUM(field_DEBIT)        field_DEBIT, ");
sqlCmd.append("                 SUM(field_CREDIT)       field_CREDIT, ");
sqlCmd.append("                 SUM(fieldI_XA)          fieldI_XA,  ");
sqlCmd.append("                 SUM(fieldI_XB1)         fieldI_XB1, ");
sqlCmd.append("                 SUM(fieldI_XB2)         fieldI_XB2, ");
sqlCmd.append("                 SUM(fieldI_XC1)         fieldI_XC1, ");
sqlCmd.append("                 SUM(fieldI_XC2)         fieldI_XC2, ");
sqlCmd.append("                 SUM(fieldI_XD1)         fieldI_XD1, ");
sqlCmd.append("                 SUM(fieldI_XD2)         fieldI_XD2, ");
sqlCmd.append("                 SUM(fieldI_XE1)         fieldI_XE1, ");
sqlCmd.append("                 SUM(fieldI_XE2)         fieldI_XE2, ");
sqlCmd.append("                 SUM(fieldI_XF1)         fieldI_XF1, ");
sqlCmd.append("                 SUM(fieldI_XF3)         fieldI_XF3, ");
sqlCmd.append("                 SUM(fieldI_XF2)         fieldI_XF2, ");
sqlCmd.append("                 SUM(fieldI_Y)           fieldI_Y,   ");
sqlCmd.append("                 SUM(nvl(a05.amt,0)) as field_CAPTIAL ");
sqlCmd.append(" 			from ( select nvl(cd01.hsien_id,' ')   as  hsien_id , ");
sqlCmd.append("        				  nvl(cd01.hsien_name,'OTHER') as  hsien_name,");
sqlCmd.append("        				  cd01.FR001W_output_order     as  FR001W_output_order, ");
sqlCmd.append("        				  bn01.bank_no ,  bn01.bank_name, ");
sqlCmd.append("   				      round(sum(decode(a01.acc_code,'120700',amt,0)) /1,0)     as field_120700, ");  
sqlCmd.append("                       round(sum(decode(a01.acc_code,'320300',amt,0)) /1,0)     as field_320300, "); 
sqlCmd.append("					      round(sum( decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'110324',amt,'110325',amt,0),'7',decode(a01.acc_code,'110254',amt,'110255',amt,0)), ");
sqlCmd.append("	           		                                   '103',decode(a01.acc_code,'110324',amt,'110325',amt,0),0) ) /1,0)     as field_TRANSFER, "); 
sqlCmd.append("	            	      round(sum( decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'110320',amt,0),'7',decode(a01.acc_code,'110250',amt,0)), ");
sqlCmd.append("	            	                                  '103',decode(a01.acc_code,'110320',amt,0),0)) /1,0)     as field_DEPOSITBANK_AA,   ");
sqlCmd.append("	            	      round(sum( decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'110300',amt,0),'7',decode(a01.acc_code,'110200',amt,0)), ");
sqlCmd.append("					                                  '103',decode(a01.acc_code,'110300',amt,0),0) ) /1,0)     as field_DEPOSITBANK, ");
sqlCmd.append("					      round(sum(decode(a01.acc_code,'310000',amt,0)) /1,0)     as field_310000,   ");
sqlCmd.append("					      round(sum(decode(bank_type,'6',decode(a01.acc_code,'310000',amt,'320000',amt,0),'7',decode(a01.acc_code,'300000',amt,0),0)) /1,0)     as field_NET,  "); 
sqlCmd.append("					      round(sum(decode(a01.acc_code,'140000',amt,0)) /1,0)     as field_140000,   ");
sqlCmd.append("               	      round(sum(decode(a01.acc_code,'220100',amt, '220200',amt, '220300',amt, '220400',amt, '220500',amt,0)) /1,0) as field_CHECK,   ");
sqlCmd.append("               	      round(sum(decode(a01.acc_code,'150200',amt,0)) /1,0)     as field_150200,   ");
sqlCmd.append("                	      round(sum(decode(a01.acc_code, '120800',amt,'150300',amt,0)) /1,0) as  field_BACKUP, ");                            
sqlCmd.append("                       round(sum( decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code, '120101',amt,'120301',amt, '120401',amt, '120501',amt,0),'7',decode(a01.acc_code, '120101',amt,'120401', amt, '120201',amt, '120501',amt,0)),  ");
sqlCmd.append("                                                   '103',decode(a01.acc_code, '120101',amt,'120301',amt, '120401',amt, '120501',amt,0),0) ) /1,0) as  field_NOASSURE,  ");
sqlCmd.append("                       round((sum(decode(bank_type,'6',decode(a01.acc_code, '310000',amt,'320000',amt, '120800',amt,'150300',amt,0),'7',decode(a01.acc_code, '300000',amt, '120800',amt, '150300',amt,0))) -   ");
sqlCmd.append("              	      round(sum(decode(a01.acc_code,'990000',amt,0)) * 1.25 * 0.7,0))/1,0) as  field_MODIFYNET,   ");
sqlCmd.append("   				      round(sum(decode(a01.acc_code,'990000',amt,0)) /1,0)        as field_OVER,   ");
sqlCmd.append("                       round(sum(decode(a01.acc_code,'220000',amt,0)) /1,0)        as field_DEBIT,  "); 
sqlCmd.append("					      round(sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /1,0) as  field_CREDIT,    ");                           
sqlCmd.append("	           		      round(sum(decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'120101',amt,'120102',amt,  ");
sqlCmd.append("	            	                                                                    '120200',amt,'120301',amt,'120302',amt,'120700',amt,'150200',amt,0)  "); 
sqlCmd.append("	            	                                                       ,'7',decode(a01.acc_code,'120101',amt,'120102',amt,  ");
sqlCmd.append("	            	                                                                    '120300',amt,'120401',amt,'120402',amt,'120700',amt,'150200',amt,0)), ");
sqlCmd.append("					                                 '103',decode(a01.acc_code,'120101',amt,'120102',amt,'120200',amt,'120301',amt, "); 
sqlCmd.append("					                                                            '120302',amt,'120700',amt,'150200',amt,0),0)) /1,0)     as fieldI_XA,  "); 
sqlCmd.append("					      round(sum(decode(year_type,'102',decode(bank_type,'6',decode(a01.acc_code,'120401',amt,'120402',amt,0),'7',decode(a01.acc_code,'120201',amt,'120202',amt,0)), ");
sqlCmd.append("					                                 '103',decode(a01.acc_code,'120401',amt,'120402',amt,0),0)) /1,0)     as fieldI_XB1, ");
sqlCmd.append("               	      sum(decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'240305',amt,'251300',amt,'310800',amt, 0),'7',decode(a01.acc_code,'240205',amt, '310800',amt,0)), ");
sqlCmd.append("               	                           '103',decode(bank_type,'6',decode(a01.acc_code,'240305',amt,'251300',amt,'310800',amt, 0),'7',decode(a01.acc_code,'240305',amt, '251200',amt,0)),0))  as fieldI_XB2, ");
sqlCmd.append("                	      round(sum(decode(a01.acc_code,'120501',amt,'120502',amt,0)) /1,0)     as fieldI_XC1, ");  
sqlCmd.append("                       round(sum( decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'240301',amt,'240302',amt,'240303',amt,'240304',amt,0), ");  
sqlCmd.append("                                                                          '7',decode(a01.acc_code,'240201',amt,'240202',amt,'240203',amt,'240204',amt,0)), ");
sqlCmd.append("                                                   '103', decode(a01.acc_code,'240301',amt,'240302',amt,'240303',amt,'240304',amt,0),0)) /1,0)     as fieldI_XC2, ");
sqlCmd.append("              	      round(sum(decode(a01.acc_code,'120600',amt,0)) /1,0)              as fieldI_XD1, "); 
sqlCmd.append("   				      round(sum( decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code,'240200',amt,0), ");
sqlCmd.append("                                                                          '7',decode(a01.acc_code,'240300',amt,0)), ");
sqlCmd.append("					                                  '103',decode(a01.acc_code,'240200',amt,0),0) ) /1,0)   as fieldI_XD2,  ");
sqlCmd.append("	           		      round(sum(decode(a01.acc_code,'150100',amt,0)) /1,0)              as fieldI_XE1,  "); 
sqlCmd.append("	            	      round(sum(decode(a01.acc_code,'250100',amt,0)) /1,0)              as fieldI_XE2,  "); 
sqlCmd.append("	            	      round(sum(decode(a01.acc_code,'310000',amt,'320000',amt,0)) /1,0) as fieldI_XF1,  "); 
sqlCmd.append("	            	      round(sum( decode(YEAR_TYPE,'102',decode(a01.acc_code,'310800',amt,0), ");
sqlCmd.append("					                                  '103',decode(bank_type,'6',decode(a01.acc_code,'310800',amt,0),'7',0,0),0)) /1,0)  as fieldI_XF3, ");  
sqlCmd.append("					      round(sum(decode(a01.acc_code,'140000',amt,0)) /1,0)              as fieldI_XF2,  "); 
sqlCmd.append("					      round((sum(decode(a01.acc_code,'220100',amt,'220200',amt,  "); 
sqlCmd.append("					                                     '220300',amt,'220400',amt,  "); 
sqlCmd.append("               	                                     '220500',amt,'220600',amt,  "); 
sqlCmd.append("               	                                     '220700',amt,'220800',amt,   ");
sqlCmd.append("                	                                     '220900',amt,'221000',amt,0))-  ");
sqlCmd.append("                       round(sum(decode(a01.acc_code,'220900',amt,0))/2,0)) /1,0)   as fieldI_Y "); 
sqlCmd.append("       			from  (select * from "+cd01_table+" cd01 where cd01.hsien_id <> 'Y') cd01 ");
sqlCmd.append("        			left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id ");
sqlCmd.append("        			left join (select * from bn01 where m_year=?)bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type in ('6','7') ");
sqlCmd.append("                 left join (select (CASE WHEN (a01.m_year <= 102) THEN '102' ");                                 
sqlCmd.append("                                          WHEN (a01.m_year > 102) THEN '103'");                                   
sqlCmd.append("                                          ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01 ");
sqlCmd.append("                            where  a01.m_year = ? and a01.m_month = ?) a01  on  bn01.bank_no = a01.bank_code ");
paramList.add(wlx01_m_year);
paramList.add(wlx01_m_year);
paramList.add(s_year);
paramList.add(s_month);
sqlCmd.append("        			group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no ,  bn01.BANK_NAME "); 
sqlCmd.append("					) a01,( select * from a05 where a05.m_year= ? and a05.m_month  = ? and  a05.acc_code = '91060P') a05 "); 
paramList.add(s_year);
paramList.add(s_month);
sqlCmd.append("				 where   a01.bank_no=a05.bank_code(+)  and a01.bank_no <> ' ' "); 
sqlCmd.append("				 GROUP  BY a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order "); 
sqlCmd.append("			) a01 "); 
sqlCmd.append("	)  a01  ORDER by    FR001W_output_order, field_SEQ,  hsien_id ,  bank_no ");


	String[] field = {"count_seq", "field_debit", "field_credit", "field_dc_rate", "field_120700",
	                  "field_over", "field_over_rate", "field_320300", "field_transfer", "field_transfer_rate",
	                  "field_310000", "field_net", "field_fixnet_rate", "field_check_rate","field_150200",
	                  "field_backup","field_noassure","field_captial_rate"};
	String[] field_type = {"2", "2", "2", "4", "2","2", "4", "2", "2", "4","2", "2", "4", "4","2","2","2","4"};

    List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,count_seq,field_seq,field_debit,field_credit,field_dc_rate,field_120700,field_over,field_over_rate,field_320300,field_transfer,field_transfer_rate,field_310000,field_net,field_fixnet_rate,field_check_rate,field_backup,field_150200,field_noassure,field_captial_rate");
    if(isDebug.equals("true")) System.out.println("dbData.size()="+dbData.size());
    //m_year,m_month,bank_code,acc_code,hsien_id,type,amt
    String field_seq="";
    DataObject bean = null;
    for(int i=0;i<dbData.size();i++){
        bean=(DataObject)dbData.get(i);
        field_seq = bean.getValue("field_seq").toString();
        if(!field_seq.equals("A90")) continue;//A90才是縣市別統計
    	
	    logps.println(Utility.getDateFormat("yyyy/MM/dd  HH:mm:ss  ")+" "+"縣市別:"+(String)bean.getValue("hsien_id")+(String)bean.getValue("hsien_name"));
	    logps.flush();
	    
        for(int j=0;j<field.length;j++){
            dataList = new ArrayList<String>();//傳內的參數List	        	 	           				   
		    dataList.add(s_year); 
		    dataList.add(s_month); 
		    dataList.add((String)bean.getValue("bank_type")); 
		    dataList.add((String)bean.getValue("bank_no")); 
		    dataList.add(field[j]); 
		    dataList.add((String)bean.getValue("hsien_id")); 
		    dataList.add(field_type[j]); //type=4-->利率. type=2-->加總
		    dataList.add((bean.getValue(field[j])).toString()); 
		    updateDBDataList.add(dataList);//1:傳內的參數List	
        }
    }//儲存參數List
    
	//1.寫入A01_OPERATION==================================================================================================
	sqlCmd.delete(0, sqlCmd.length());
	sqlCmd.append("insert into a01_operation values(?,?,?,?,?,?,?,?)");
    updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql    				
	updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
	updateDBList.add(updateDBSqlList);		
	//===================================================================================================================
	String total_count="0";//總筆數
	total_count=String.valueOf(updateDBDataList.size());
    //2.寫入OPERATION_LOG 98.08.25 add ===================================================================================================
    updateDBSqlList = new ArrayList();		
	updateDBDataList = new ArrayList<List>();//儲存參數的List	
    sqlCmd.delete(0, sqlCmd.length());
	sqlCmd.append("insert into OPERATION_LOG values(?,?,?,?,?,?,?,sysdate)");
    dataList = new ArrayList<String>();//傳內的參數List	        	 	           				   
	dataList.add(s_year); 
	dataList.add(s_month); 
	dataList.add(Report_no); 
	dataList.add("hsien_id_ALL"); 
	dataList.add(total_count); 
	dataList.add(lguser_id); 
	dataList.add(lguser_id); 
	updateDBDataList.add(dataList);//1:傳內的參數List
	
	updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql    				
	updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
	updateDBList.add(updateDBSqlList);	
    //=================================================================================================================================

    if(DBManager.updateDB_ps(updateDBList)){
       errMsg = errMsg + "相關資料寫入資料庫成功";
	}else{
	   errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
	}	
	
	if(errMsg.equals("相關資料寫入資料庫成功")){
       logps.println(Utility.getDateFormat("yyyy/MM/dd  HH:mm:ss  ")+" "+"產生 A01_opeation 完成");
    }else{
       logps.println(Utility.getDateFormat("yyyy/MM/dd  HH:mm:ss  ")+" "+"執行 A01_opeation 失敗:"+errMsg);
    }
	logps.flush();
   }catch (Exception e){
		System.out.println(e+":"+e.getMessage());
	    errMsg = errMsg + "相關資料寫入資料庫失敗";
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
%>
</body>
</html>