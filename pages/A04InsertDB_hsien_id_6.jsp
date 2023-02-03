<%
// 95.10.17 add 將農會縣市別小計寫入a04_opeartion by 2295
// 			   bn01.bank_type in ('7')-->漁會
// 								 ('6')-->農會
// 95.11.30 add 排程產生A01_opeation by 2295
// 98.08.25 add 寫入OPERATION_LOG by 2295
// 99.11.05 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
//103.08.18 fix field_RATE_840760_OVER_A04欄位太長,造成sql error(改成field_840760_OVER_A04) by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%@ page import="java.io.*" %>
<%@ page import="java.lang.Integer" %>
<%@ page import="java.text.DecimalFormat" %>
<%@ page import="java.text.SimpleDateFormat" %>
<html>
<head>
<title></title>
</head>
<body>
產生(農會-縣市別小計)A04_operation
<%
   System.out.println("=============產生A04_Operation開始===========");
   String report_no = ( request.getParameter("report_no")==null ) ? "" : (String)request.getParameter("report_no");
   String s_year = ( request.getParameter("s_year")==null ) ? "" : (String)request.getParameter("s_year");
   String s_month = ( request.getParameter("s_month")==null ) ? "" : (String)request.getParameter("s_month");
   String isDebug = ( request.getParameter("isDebug")==null ) ? "" : (String)request.getParameter("isDebug");
   String lguser_id = ( request.getParameter("lguser_id")==null ) ? "" : (String)request.getParameter("lguser_id");
   String errMsg = Generate(report_no,s_year,s_month,isDebug,lguser_id);
   System.out.println("errMsg = "+errMsg);
   System.out.println("=============產生A04_Operation結束===========");
%>

<%!
public String Generate(String Report_no,String s_year,String s_month,String isDebug,String lguser_id) throws Exception{
		File logfile;
		FileOutputStream logos=null;
		BufferedOutputStream logbos = null;
		PrintStream logps = null;
		Date nowlog = new Date();
		SimpleDateFormat logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");
		SimpleDateFormat logfileformat = new SimpleDateFormat("yyyyMMddHHmmss");
	    Calendar logcalendar;
	    File logDir = null;
	    String errMsg="";
 		String bank_type="6";
        //99.11.05 add==================================================================
		String cd01_table = "";
    	String wlx01_m_year = "";
    	StringBuffer sqlCmd = new StringBuffer();
		List<String> paramList = new ArrayList<String>();
		List<List> updateDBList = new ArrayList<List>();//0:sql 1:data
		List updateDBSqlList = new ArrayList();
		List<List> updateDBDataList = new ArrayList<List>();//儲存參數的List
		List<String> dataList =  new ArrayList<String>();//儲存參數的data
		DataObject bean = null;
		//============================================================================
 try{
 	    logDir  = new File(Utility.getProperties("logDir"));
	    if(!logDir.exists()){
          if(!Utility.mkdirs(Utility.getProperties("logDir"))){
     		  System.out.println("目錄新增失敗");
     	  }
        }

	   logfile = new File(logDir + System.getProperty("file.separator") + Report_no +"_GenerateOperation."+ logfileformat.format(nowlog));
	   System.out.println("logfile filename="+logDir + System.getProperty("file.separator") + Report_no +"_GenerateOperation."+ logfileformat.format(nowlog));
	   logos = new FileOutputStream(logfile,true);
	   logbos = new BufferedOutputStream(logos);
	   logps = new PrintStream(logbos);
       logcalendar = Calendar.getInstance();
	   nowlog = logcalendar.getTime();
	   logps.println(logformat.format(nowlog)+" "+"產生"+s_year+"年"+s_month+"月 (農會-縣市別小計)A04_opeation 資料中");
	   logps.flush();

	   //99.11.05 add 查詢年度100年以前.縣市別不同===============================
  	   cd01_table = (Integer.parseInt(s_year) < 100)?"cd01_99":"";
  	   wlx01_m_year = (Integer.parseInt(s_year) < 100)?"99":"100";
  	   //=====================================================================

      //各縣市小計
      //按縣市別
	  //農會
sqlCmd.append(" select distinct g.hsien_id,g.hsien_name,g.fr001w_output_order,                       ");
sqlCmd.append(" 	   f.field_OVER_A01,f.field_CREDIT_A01,f.field_OVER_RATE_A01,                    ");
sqlCmd.append(" 	   f.field_OVER_A04,                                                             ");
sqlCmd.append(" 	   f.field_CREDIT_A04,                                                           ");
sqlCmd.append(" 	   f.field_OVER_RATE_A04,                                                        ");
sqlCmd.append(" 	   f.field_840760_A04,                                                           ");
sqlCmd.append(" 	   f.field_RATE_840760_A04,                                                      ");
sqlCmd.append("		   f.field_840760_OVER_A04,                                                 ");
sqlCmd.append("		   f.field_840710_A04_A,                                                         ");
sqlCmd.append("		   f.field_840720_A04_B,                                                         ");
sqlCmd.append("		   f.field_840731_A04_a,                                                         ");
sqlCmd.append("		   f.field_840732_A04_b,                                                         ");
sqlCmd.append("		   f.field_840733_A04_c,                                                         ");
sqlCmd.append("		   f.field_840734_A04_d,                                                         ");
sqlCmd.append("		   f.field_840735_A04_e	                                                         ");
sqlCmd.append("	from (select nvl(e.hsien_id,' ') as hsien_id,nvl(e.hsien_name,'其他') as hsien_name, ");
sqlCmd.append("		         round(sum(field_OVER_A01)/1,0) as field_OVER_A01,                       ");
sqlCmd.append("		         round(sum(field_CREDIT_A01)/1,0) as field_CREDIT_A01,                   ");
sqlCmd.append(" 		     decode(round(sum(field_OVER_A01)/1,0),0,0,round(round(sum(field_OVER_A01)/1,0) /  round(sum(field_CREDIT_A01)/1,0) *100 ,2))  as   field_OVER_RATE_A01,         ");
sqlCmd.append(" 		     round(sum(field_OVER_A04)/1,0) as field_OVER_A04,                                                                                                               ");
sqlCmd.append(" 		     round(sum(field_CREDIT_A04)/1,0) as field_CREDIT_A04,                                                                                                           ");
sqlCmd.append(" 		     decode(round(sum(field_CREDIT_A04)/1,0),0,0,round(round(sum(field_OVER_A04)/1,0) /  round(sum(field_CREDIT_A04)/1,0) *100 ,2))  as   field_OVER_RATE_A04,       ");
sqlCmd.append(" 		     round(sum(field_840760_A04)/1,0) as field_840760_A04,                                                                                                           ");
sqlCmd.append(" 		     decode(round(sum(field_CREDIT_A04)/1,0) ,0,0,round(round(sum(field_840760_A04)/1,0) /  round(sum(field_CREDIT_A04)/1,0)  *100 ,2))  as   field_RATE_840760_A04, ");
sqlCmd.append(" 		     decode(round(sum(field_CREDIT_A04)/1,0),0,0,round((round(sum(field_OVER_A04)/1,0) +round(sum(field_840760_A04)/1,0)) /  round(sum(field_CREDIT_A04)/1,0) *100 ,2))  as  field_840760_OVER_A04,  ");
sqlCmd.append(" 		     round(sum(field_840710_A04_A)/1,0) as field_840710_A04_A,                                                          ");
sqlCmd.append(" 		     round(sum(field_840720_A04_B)/1,0) as field_840720_A04_B,                                                          ");
sqlCmd.append(" 		     round(sum(field_840731_A04_a)/1,0) as field_840731_A04_a,                                                          ");
sqlCmd.append(" 		     round(sum(field_840732_A04_b)/1,0) as field_840732_A04_b,                                                          ");
sqlCmd.append(" 		     round(sum(field_840733_A04_c)/1,0) as field_840733_A04_c,                                                          ");
sqlCmd.append(" 		     round(sum(field_840734_A04_d)/1,0) as field_840734_A04_d,                                                          ");
sqlCmd.append(" 		     round(sum(field_840735_A04_e)/1,0) as field_840735_A04_e                                                           ");
sqlCmd.append("       from (select a04.hsien_id,a04.hsien_name,a04.FR001W_output_order,a04.bank_type,a04.bank_no ,a04.BANK_NAME,                ");
sqlCmd.append(" 		           round(field_OVER_A01 /1,0)    as field_OVER_A01,                                                             ");
sqlCmd.append(" 		           round(field_CREDIT_A01 /1,0)  as field_CREDIT_A01,                                                           ");
sqlCmd.append(" 		           decode(field_CREDIT_A01,0,0,round(field_OVER_A01 /  field_CREDIT_A01 *100 ,2))  as   field_OVER_RATE_A01,    ");
sqlCmd.append(" 		           round(field_OVER_A04 /1,0)    as field_OVER_A04,                                                             ");
sqlCmd.append(" 		           round(field_CREDIT_A04 /1,0)  as field_CREDIT_A04,                                                           ");
sqlCmd.append(" 		           decode(field_CREDIT_A04,0,0,round(field_OVER_A04 /  field_CREDIT_A04 *100 ,2))  as   field_OVER_RATE_A04,    ");
sqlCmd.append(" 		           round(field_840760_A04 /1,0)  as field_840760_A04,                                                           ");
sqlCmd.append(" 		           decode(field_CREDIT_A04,0,0,round(field_840760_A04 /  field_CREDIT_A04 *100 ,2))  as   field_RATE_840760_A04,");
sqlCmd.append(" 		           decode(field_CREDIT_A04,0,0,round((field_OVER_A04 + field_840760_A04) /  field_CREDIT_A04 *100 ,2))  as  field_840760_OVER_A04,");
sqlCmd.append(" 		           round(field_840710_A04_A /1,0)  as field_840710_A04_A,                                                       ");
sqlCmd.append(" 		           round(field_840720_A04_B /1,0)  as field_840720_A04_B,                                                       ");
sqlCmd.append(" 		           round(field_840731_A04_a /1,0)  as field_840731_A04_a,                                                       ");
sqlCmd.append(" 		           round(field_840732_A04_b /1,0)  as field_840732_A04_b,                                                       ");
sqlCmd.append(" 		           round(field_840733_A04_c /1,0)  as field_840733_A04_c,                                                       ");
sqlCmd.append(" 		           round(field_840734_A04_d /1,0)  as field_840734_A04_d,                                                       ");
sqlCmd.append(" 		           round(field_840735_A04_e /1,0)  as field_840735_A04_e                                                        ");
sqlCmd.append(" 			from ( select nvl(cd01.hsien_id,' ') as  hsien_id , nvl(cd01.hsien_name,'OTHER')  as  hsien_name,                   ");
sqlCmd.append(" 			              cd01.FR001W_output_order  as  FR001W_output_order,                                                    ");
sqlCmd.append(" 			              bn01.bank_type,bn01.bank_no , bn01.BANK_NAME,                                                         ");
sqlCmd.append(" 			              sum(decode(acc_code,'990000',amt,0))  as field_OVER_A01,                                              ");
sqlCmd.append(" 			              sum(decode(acc_code,'120000',amt,'120800',amt,'150300',amt,0))  	as  field_CREDIT_A01,               ");
sqlCmd.append(" 			              sum(decode(acc_code,'840740',amt,0))         as field_OVER_A04,                                       ");
sqlCmd.append(" 			              sum(decode(acc_code,'840750',amt,0))         as field_CREDIT_A04,                                     ");
sqlCmd.append(" 			              sum(decode(acc_code,'840760',amt,0))         as field_840760_A04,                                     ");
sqlCmd.append(" 			              sum(decode(acc_code,'840710',amt,0))         as field_840710_A04_A,                                   ");
sqlCmd.append(" 			              sum(decode(acc_code,'840720',amt,0))         as field_840720_A04_B,                                   ");
sqlCmd.append(" 			              sum(decode(acc_code,'840731',amt,0))         as field_840731_A04_a,                                   ");
sqlCmd.append(" 			              sum(decode(acc_code,'840732',amt,0))         as field_840732_A04_b,                                   ");
sqlCmd.append("          		          sum(decode(acc_code,'840733',amt,0))         as field_840733_A04_c,                                   ");
sqlCmd.append("          		          sum(decode(acc_code,'840734',amt,0))         as field_840734_A04_d,                                   ");
sqlCmd.append("          		          sum(decode(acc_code,'840735',amt,0))         as field_840735_A04_e                                    ");
sqlCmd.append("          	        from  (select * from "+cd01_table+" cd01 where cd01.hsien_id <> 'Y') cd01                                   ");
sqlCmd.append("          	        left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id                         ");
sqlCmd.append("          	        left join (select * from bn01 where m_year=?)bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type in (?)  ");
sqlCmd.append("          	        left join ((select * from a04 where m_year = ? and m_month = ?)                                             ");
sqlCmd.append("          	                    union (select * from a01 where m_year = ? and m_month = ?                                       ");
sqlCmd.append("          	                     and acc_code in('990000', '120000','120800','150300'))) a04                                    ");
sqlCmd.append("          	                  on  bn01.bank_no = a04.bank_code                                                                  ");
sqlCmd.append("          	        group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_type,       ");
sqlCmd.append("          	        		 bn01.bank_no , bn01.BANK_NAME) a04                                                                 ");
sqlCmd.append("                  where a04.bank_no <> ' ' order by a04.hsien_id ,  a04.hsien_name,  a04.FR001W_output_order, a04.bank_type,a04.bank_no, a04.BANK_NAME  ");
sqlCmd.append("   ) d,(select * from v_bank_location where m_year=?)e                                            ");
sqlCmd.append("   where e.bank_type=?                                                                              ");
sqlCmd.append("   and d.bank_no(+)=e.bank_no                                                                       ");
sqlCmd.append("   and (e.hsien_id>'Y' or e.hsien_id<'Y')                                                           ");
sqlCmd.append("   group by nvl(e.hsien_id,' '),nvl(e.hsien_name,'其他'),e.fr001w_output_order) f,(select * from v_bank_location where m_year=?)g ");
sqlCmd.append("  where f.hsien_id(+)=g.hsien_id order by g.fr001w_output_order                                     ");
paramList.add(wlx01_m_year);
paramList.add(wlx01_m_year);
paramList.add(bank_type);
paramList.add(s_year);
paramList.add(s_month);
paramList.add(s_year);
paramList.add(s_month);
paramList.add(wlx01_m_year);
paramList.add(bank_type);
paramList.add(wlx01_m_year);
    List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"field_OVER_A01,field_CREDIT_A01,field_OVER_RATE_A01,field_OVER_A04,field_CREDIT_A04,field_OVER_RATE_A04,field_840760_A04,field_RATE_840760_A04,field_840760_OVER_A04,field_840710_A04_A,field_840720_A04_B,field_840731_A04_a,field_840732_A04_b,field_840733_A04_c,field_840734_A04_d,field_840735_A04_e");
    if(isDebug.equals("true")) System.out.println("dbData.size()="+dbData.size());
    //1.field_OVER_A01=逾期放款[990000]
    //2.field_CREDIT_A01=總放款(放款含催收)[120000+120800+150300]
	//3.field_OVER_RATE_A01=逾放比率
	//4.field_OVER_A04=1.逾期放款[840740]
	//5.field_CREDIT_A04=2.總放款(放款含催收)[840750]
	//6.field_OVER_RATE_A04=3.逾放比率
	//7.field_840760_A04=4.應予觀察放款金額[840760]
	//8.field_RATE_840760_A04=5.應予觀察放款占總放款比率
	//9.field_840760_OVER_A04=6.逾期放款及應予觀察放款占總放款比率
	//10.field_840710_A04_A=放款本金未超過清償期三個月，惟利息未按期繳納超過三個月至六個月者840710(A)
	//11.field_840720_A04_B=中長期分期償債放款，未按期攤還超過三個月至六個月者840720(B)
	//12.field_840731_A04_a=協議分期償還放款，協議條件符合規定，且借款戶依協議條件按期履約未滿六個月者840731(a)
	//13.field_840732_A04_b=已獲信用保證基金同意理賠款項或有足額存單或存款備償(須辦妥質權設定且徵得發單行拋棄抵銷權同意書)，而約定待其他債務人財產處分後再予沖償者840732(b)
	//14.field_840733_A04_c=已確定分配之債權，惟尚未接獲分配款者840733(c)
	//15.field_840734_A04_d=債務人兼擔保品提供人死亡，於辦理繼承期間屆期而未清償之放款840734(d)
	//16.field_840735_A04_e=其他840735(e)
    String[] field = {"field_over_a01","field_credit_a01","field_over_rate_a01","field_over_a04","field_credit_a04","field_over_rate_a04",
					  "field_840760_a04","field_rate_840760_a04","field_840760_over_a04","field_840710_a04_a","field_840720_a04_b",
					  "field_840731_a04_a","field_840732_a04_b","field_840733_a04_c","field_840734_a04_d","field_840735_a04_e"};
	String[] field_type = {"2","2","4","2","2","4","2","4","4","2","2","2","2","2","2","2"};


    //m_year,m_month,bank_code,acc_code,hsien_id,type,amt
    for(int i=0;i<dbData.size();i++){        
        bean=(DataObject)dbData.get(i);       
    	logcalendar = Calendar.getInstance();
	    nowlog = logcalendar.getTime();
	    logps.println(logformat.format(nowlog)+" "+"縣市別:"+(String)bean.getValue("hsien_id"));
	    logps.flush();		
	    for(int j=0;j<field.length;j++){	       
            dataList = new ArrayList<String>();//傳內的參數List
		    dataList.add(s_year);
		    dataList.add(s_month);
		    dataList.add(bank_type);
		    dataList.add("ALL");
		    dataList.add(field[j]);
		    dataList.add((String)bean.getValue("hsien_id"));
		   	dataList.add(field_type[j]); //type=4-->利率. type=2-->加總
		    dataList.add((bean.getValue(field[j])).toString());
		    dataList.add("");
		    updateDBDataList.add(dataList);//1:傳內的參數List		    
        }
    }//儲存參數List

    //1.寫入A01_OPERATION==================================================================================================
	sqlCmd.delete(0, sqlCmd.length());
	sqlCmd.append("insert into a04_operation values(?,?,?,?,?,?,?,?,?)");
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
	dataList.add("hsien_id_6");
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
	logcalendar = Calendar.getInstance();
	nowlog = logcalendar.getTime();
	if(errMsg.equals("相關資料寫入資料庫成功")){
       logps.println(logformat.format(nowlog)+" "+"產生 A04_opeation 完成");
    }else{
       logps.println(logformat.format(nowlog)+" "+"執行 A04_opeation 失敗:"+errMsg);
    }
	logps.flush();
   }catch (Exception e){
		System.out.println(e+":"+e.getMessage());
	    errMsg = errMsg + "相關資料寫入資料庫失敗";
		logcalendar = Calendar.getInstance();
		nowlog = logcalendar.getTime();
	    logps.println(logformat.format(nowlog)+" "+"UpdateDB Error:"+e + "\n"+e.getMessage());
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