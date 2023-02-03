<%
// 95.10.02 add 將農漁會-各別機構寫入a03_opeartion by 2295
// 			   bn01.bank_type in ('7')-->漁會
// 								 ('6')-->農會
// 95.11.29 add 排程產生A03_opeation by 2295
// 98.08.25 add 寫入OPERATION_LOG by 2295
// 99.11.04 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
//102.06.21 add 改寫原本查詢SQL.並合併農.漁會各別機構 by 2295                                                                  
//102.06.21 add 103/01以後.A01漁會.套用新科目代號/計算公式 by 2295 
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
產生(農漁會-各別機構)A03_operation
<%
  System.out.println("=============產生A03_Operation開始===========");
  String report_no = ( request.getParameter("report_no")==null ) ? "" : (String)request.getParameter("report_no");
  String s_year = ( request.getParameter("s_year")==null ) ? "" : (String)request.getParameter("s_year");
  String s_month = ( request.getParameter("s_month")==null ) ? "" : (String)request.getParameter("s_month");
  String isDebug = ( request.getParameter("isDebug")==null ) ? "" : (String)request.getParameter("isDebug");
  String lguser_id = ( request.getParameter("lguser_id")==null ) ? "" : (String)request.getParameter("lguser_id");
  String errMsg = Generate(report_no,s_year,s_month,isDebug,lguser_id);
  System.out.println("errMsg = "+errMsg);
  System.out.println("=============產生A03_Operation結束===========");


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
      	String s_year_last="";
	  	String s_month_last="";
	  	String div="2";
	  	String bank_type="6";
	  	//99.11.04 add==================================================================
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
	   logps.println(logformat.format(nowlog)+" "+"產生"+s_year+"年"+s_month+"月 (農漁會-各別機構)A03_opeation 資料中");
	   logps.flush();

	  div=(Integer.parseInt(s_year)==94 && Integer.parseInt(s_month)==6)?"1":"2";
	  if (Integer.parseInt(s_month)==1) {
		  s_year_last=String.valueOf(Integer.parseInt(s_year)-1);
		  s_month_last="12";
	  }else {
		  s_year_last=s_year;
		  s_month_last=String.valueOf(Integer.parseInt(s_month)-1);
	  }

       //99.11.04 add 查詢年度100年以前.縣市別不同===============================
  	   cd01_table = (Integer.parseInt(s_year) < 100)?"cd01_99":"";
  	   wlx01_m_year = (Integer.parseInt(s_year) < 100)?"99":"100";
  	   //=====================================================================

      //各別機構
	  //農會
	  sqlCmd.append(" select e.fr001w_output_order,bank_code,d.bank_name, e.bank_type,e.hsien_id, ");
	  sqlCmd.append("    round(sum(field_b)/ 1,0) as fieldB ,  ");//--平均放款總額
      sqlCmd.append("    round(sum(field_c)/ 1,0) as fieldC ,  ");//--平均內部融資
      sqlCmd.append("    round(sum(field_d)/ 1,0) as fieldD ,  ");//--平均統一農(漁)貸
      sqlCmd.append("    round(sum(field_e)/ 1,0) as fieldE ,  ");//--放款利息收入
      sqlCmd.append("    round(sum(field_f)/ 1,0) as fieldF ,  ");//--內部融資利息收入
      sqlCmd.append("    round(sum(field_g)/ 1,0) as fieldG ,  ");//--統一農(漁)貸利息收入
      sqlCmd.append("    decode(sum(field_b),0,0,round(sum(field_e) / sum(field_b) *100 ,2))  as fieldH,  ");//--平均利率
      sqlCmd.append("    decode(sum(field_c),0,0,round(sum(field_f) / sum(field_c) *100 ,2))  as fieldI,  ");//--內部融資利率
      sqlCmd.append("    decode(sum(field_d),0,0,round(sum(field_g) / sum(field_d) *100 ,2))  as fieldJ,  ");//--統一農貸平均利率
      sqlCmd.append("    decode(sum(field_r),0,0,round(sum(field_k) / sum(field_r) *100 ,2))  as fieldK,  ");//--綜合平均存款利率
      sqlCmd.append("    decode(sum(field_u),0,0,round(sum(field_l) / sum(field_u) *100 ,2))  as fieldL,  ");//--活期存款.平均利率
      sqlCmd.append("    decode(sum(field_v),0,0,round(sum(field_m) / sum(field_v) *100 ,2))  as fieldM,  ");//--活儲存款.平均利率
      sqlCmd.append("    decode(sum(field_w),0,0,round(sum(field_n) / sum(field_w) *100 ,2))  as fieldN,  ");//--員工活期儲蓄存款.平均利率
      sqlCmd.append("    decode(sum(field_x),0,0,round(sum(field_o) / sum(field_x) *100 ,2))  as fieldO,  ");//--定期存款.平均利率
      sqlCmd.append("    decode(sum(field_y),0,0,round(sum(field_p) / sum(field_y) *100 ,2))  as fieldP,  ");//--定期儲蓄存款.平均利率
      sqlCmd.append("    decode(sum(field_z),0,0,round(sum(field_q) / sum(field_z) *100 ,2))  as fieldQ,  ");//--員工定期儲蓄存款.平均利率     
      sqlCmd.append("    round(sum(field_r) / 1,0) as fieldR , ");//--存款總額.合計
      sqlCmd.append("	 round(sum(field_s) / 1,0) as fieldS , ");//--支票存款
      sqlCmd.append("	 round(sum(field_t) / 1,0) as fieldT , ");//--保付支票
      sqlCmd.append("	 round(sum(field_u) / 1,0) as fieldU , ");//--活期存款
      sqlCmd.append("	 round(sum(field_v) / 1,0) as fieldV , ");//--活儲存款
      sqlCmd.append("	 round(sum(field_w) / 1,0) as fieldW , ");//--員工活期儲蓄存款
      sqlCmd.append("    round(sum(field_x) / 1,0) as fieldX , ");//--定期存款
      sqlCmd.append("    round(sum(field_y) / 1,0) as fieldY , ");//--定期儲蓄存款
      sqlCmd.append("    round(sum(field_z) / 1,0) as fieldZ , ");//--員工定期儲蓄存款
      sqlCmd.append("    round(sum(field_za)/ 1,0) as fieldZA, ");//--公庫存款
      sqlCmd.append("    round(sum(field_zb)/ 1,0) as fieldZB  ");//--本會支票
      sqlCmd.append(" from (select * from "+cd01_table+" cd01)cd01 ");
      sqlCmd.append(" left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id ");  
      paramList.add(wlx01_m_year);
      sqlCmd.append(" left join ( ");
      sqlCmd.append(" select a01.bank_code,a01.bank_name, ");        
      sqlCmd.append(" field_b,field_c,field_d,field_e,field_f,field_g,field_k,field_l,field_m,field_n,field_o,field_p,  ");      
      sqlCmd.append(" field_q,field_r,field_s,field_t,field_u,field_v,field_w,field_x,field_y,field_z,field_za,field_zb ");
      sqlCmd.append(" from (      ");   
      sqlCmd.append(" 		select bank_code,bank_name,");
      sqlCmd.append(" 		       round((sum(decode(acc_code,'120000',amt,'150300',amt,0))- sum(decode(acc_code,'150200',amt,0)))/"+div+",0) as field_b,");
      sqlCmd.append(" 		       round(sum(decode(acc_code,'120700',amt,0))/"+div+",0)  as field_c,");
      sqlCmd.append(" 		       round((decode(YEAR_TYPE,'102',decode(bank_type,'6',sum(decode(acc_code,'120401',amt,'120402',amt,0)),");
      sqlCmd.append(" 		                                                      '7',sum(decode(acc_code,'120201',amt,'120202',amt,0)),0), ");
      sqlCmd.append(" 		                               '103',sum(decode(acc_code,'120401',amt,'120402',amt,0)),0))/"+div+",0)  as field_d,");
      sqlCmd.append("			   round(sum(decode(acc_code,'220000',amt,0))/"+div+",0) as field_r, ");
      sqlCmd.append("			   round(sum(decode(acc_code,'220100',amt,0))/"+div+",0) as field_s, ");
      sqlCmd.append("			   round(sum(decode(acc_code,'220200',amt,0))/"+div+",0) as field_t, ");
      sqlCmd.append("			   round(sum(decode(acc_code,'220300',amt,0))/"+div+",0) as field_u, ");
      sqlCmd.append("			   round(sum(decode(acc_code,'220400',amt,0))/"+div+",0) as field_v, ");
      sqlCmd.append(" 		       round(sum(decode(acc_code,'220500',amt,0))/"+div+",0) as field_w, ");
      sqlCmd.append(" 		       round(sum(decode(acc_code,'220600',amt,0))/"+div+",0) as field_x, ");
      sqlCmd.append(" 		       round(sum(decode(acc_code,'220700',amt,0))/"+div+",0) as field_y, ");
      sqlCmd.append(" 		       round(sum(decode(acc_code,'220800',amt,0))/"+div+",0) as field_z, ");
      sqlCmd.append(" 		       round(sum(decode(acc_code,'220900',amt,0))/"+div+",0) as field_za,");
      sqlCmd.append(" 		       round(sum(decode(acc_code,'221000',amt,0))/"+div+",0) as field_zb ");
      sqlCmd.append(" 		     from (select  (CASE WHEN (a01.m_year <= 102) THEN '102' ");
      sqlCmd.append("                                                   WHEN (a01.m_year > 102) THEN '103' ");
      sqlCmd.append("                                              ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01)a01 ");
      sqlCmd.append("            left join (select * from ba01 where m_year=?)ba01 on a01.bank_code = ba01.bank_no  ");
      paramList.add(wlx01_m_year);   
      sqlCmd.append("            where (a01.m_year=? and a01.m_month=? or a01.m_year=? and m_month=?)");
      paramList.add(s_year_last);     
      paramList.add(s_month_last);
      paramList.add(s_year);      
      paramList.add(s_month);           
      sqlCmd.append(" 		     and acc_code in ('120000','150300','150200','120700','120401','120402','120201','120202',");
      sqlCmd.append(" 		     				  '420100','520100','220000','220100','220200','220300','220400','220500',");
      sqlCmd.append(" 		     				  '220600','220700','220800','220900','221000')                           ");
      sqlCmd.append(" 		    group by YEAR_TYPE,bank_type,bank_name,bank_code      ");
      sqlCmd.append("             order by bank_code                                  ");
      sqlCmd.append(" 		    )a01,                                                 ");
      sqlCmd.append(" 		    (select bank_code,bank_name,                          ");
      sqlCmd.append("			      sum(decode(acc_code,'420100',amt,0)) as field_e,");
      sqlCmd.append("			      sum(decode(acc_code,'420170',amt,0)) as field_f,");
      sqlCmd.append("			      decode(bank_type,'6',sum(decode(acc_code,'420140',amt,0)),");
      sqlCmd.append("			                       '7',sum(decode(acc_code,'420120',amt,0)),0) as field_g,");
      sqlCmd.append("			      sum(decode(acc_code,'520100',amt,0)) as field_k,");
      sqlCmd.append("  			      sum(decode(acc_code,'520110',amt,0)) as field_l,");
      sqlCmd.append("  			      sum(decode(acc_code,'520130',amt,0)) as field_m,");
      sqlCmd.append("  			      sum(decode(acc_code,'840640',amt,0)) as field_n,");
      sqlCmd.append("  			      sum(decode(acc_code,'520120',amt,0)) as field_o,");
      sqlCmd.append("  			      sum(decode(acc_code,'520140',amt,0)) as field_p,");
      sqlCmd.append("  			      sum(decode(acc_code,'840670',amt,0)) as field_q ");
      sqlCmd.append("  			from (select bank_code,acc_code,-amt as amt");
      sqlCmd.append("  			      from a01                             ");
      sqlCmd.append("  			      where m_year=? and m_month=?  ");
      paramList.add(s_year_last);     
 	  paramList.add(s_month_last);    
      sqlCmd.append("  			      and acc_code in ('420100','520100')  ");
      sqlCmd.append("  			      union all                            ");
      sqlCmd.append("  			      select bank_code,acc_code,amt        ");
      sqlCmd.append("  			      from a01                             ");
      sqlCmd.append("  			      where m_year=? and m_month=?   ");
      paramList.add(s_year);          
 	  paramList.add(s_month);   
      sqlCmd.append("  			      and acc_code in ('420100','520100')  ");
      sqlCmd.append("  			      union all                            ");
      sqlCmd.append("  			      select bank_code,acc_code,-amt as amt ");
      sqlCmd.append("  			      from a03 where m_year=? and m_month=?");
      paramList.add(s_year_last);     
 	  paramList.add(s_month_last);  
      sqlCmd.append("	 			      and acc_code in ('420120','420140','420170','520110','520130','840640','520120','520140','840670')");
      sqlCmd.append("	 			      union all ");
      sqlCmd.append("	 			      select bank_code,acc_code,amt ");
      sqlCmd.append("	 			      from a03 where m_year=? and m_month=?");
      paramList.add(s_year);          
 	  paramList.add(s_month);  
      sqlCmd.append("	 			      and acc_code in ('420120','420140','420170','520110','520130','840640','520120','520140','840670')");
      sqlCmd.append("  			     )a03 left join (select * from ba01 where m_year=?)ba01 on a03.bank_code = ba01.bank_no ");
      paramList.add(wlx01_m_year);   
      sqlCmd.append("  			     group by bank_type,bank_name,bank_code ");
      sqlCmd.append("	 			 order by bank_code ");
      sqlCmd.append("	 			)a01_1 where a01.bank_code=a01_1.bank_code  ");
      sqlCmd.append("	     ) d on d.bank_code=wlx01.bank_no  ,(select * from v_bank_location where m_year=?) e  ");
      paramList.add(wlx01_m_year);   
      sqlCmd.append("	     where e.bank_type in ('6','7') and d.bank_code=e.bank_no ");
      sqlCmd.append("	     group by e.fr001w_output_order,d.bank_code,d.bank_name,e.bank_type,e.hsien_id ");  
      sqlCmd.append("	     order by e.fr001w_output_order,d.bank_code ");

    List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"fieldB,fieldC,fieldD,fieldE,fieldF,fieldG,fieldH,fieldI,fieldJ,fieldK,fieldL,fieldM,fieldN,fieldO,fieldP,fieldQ,fieldR,fieldS,fieldT,fieldU,fieldV,fieldW,fieldX,fieldY,fieldZ,fieldZA,fieldZB");

    if(isDebug.equals("true")) System.out.println("dbData.size()="+dbData.size());

	//0.B:field_120000_150200_150300=平均放款總額 1.C:field_120700=平均內部融資 2.D:field_120401_120402=平均統一農貸
    //3.E:field_420100=放款利息收入 4.F:field_420170=內部融資利息收入 5.G:field_420140=統一農(漁)貸利息收入
    //6.H:field_420100/(120000_150200_150300)=平均利率 7.I:field_420170/120700=內部融資利率 8.J:field_420140/(120401_120402)=統一農貸平均利率
    //9.K:field_520100/220000=綜合平均存款利率 10.L:field_520110/220300=活期存款  11.M:field_520130/220400=活儲存款
	//12.N:field_840640/220500=員工活期儲蓄存款 13.O:field_520120/220600=定期存款 14.P:field_520140/220700=定期儲蓄存款
    //15.Q:field_840670/220800=員工定期儲蓄存款 16.R:field_220000=合計 17.S:field_220100=支票存款
    //18.T:field_220200=保付支票 19.U:field_220300=活期存款  20.V:field_220400=活儲存款 21.W:field_220500=員工活期儲蓄存款
    //22.X:field_220600=定期存款 23.Y:field_220700=定期儲蓄存款 24.Z:field_220800=員工定期儲蓄存款 25.ZA:field_220900=公庫存款
    //26.ZB:field_221000=本會支票
	//彈性報表:欄位名稱
	String[] field_name_6 = {"field_1200_1502_1503","field_120700","field_120401_120402","field_420100","field_420170",
						   "field_420140","field_4201_b","field_420170_120700","field_420140_d","field_520100_220000",
						   "field_520110_220300","field_520130_220400","field_840640_220500","field_520120_220600",
						   "field_520140_220700","field_840670_220800","field_220000","field_220100","field_220200",
						   "field_220300","field_220400","field_220500","field_220600","field_220700","field_220800",
						   "field_220900","field_221000"};
						   
	//0.B:field_120000_150200_150300=平均放款總額 1.C:field_120700=平均內部融資 2.D:field_120201_120202=平均統一漁貸
    //3.E:field_420100=放款利息收入 4.F:field_420170=內部融資利息收入 5.G:field_420120=統一農(漁)貸利息收入
    //6.H:field_420100/(120000_150200_150300)=平均利率 7.I:field_420170/120700=內部融資利率 8.J:field_420120/(120201_120202)=統一農/漁貸平均利率
    //9.K:field_520100/220000=綜合平均存款利率 10.L:field_520110/220300=活期存款  11.M:field_520130/220400=活儲存款
	//12.N:field_840640/220500=員工活期儲蓄存款 13.O:field_520120/220600=定期存款 14.P:field_520140/220700=定期儲蓄存款
    //15.Q:field_840670/220800=員工定期儲蓄存款 16.R:field_220000=合計 17.S:field_220100=支票存款
    //18.T:field_220200=保付支票 19.U:field_220300=活期存款  20.V:field_220400=活儲存款 21.W:field_220500=員工活期儲蓄存款
    //22.X:field_220600=定期存款 23.Y:field_220700=定期儲蓄存款 24.Z:field_220800=員工定期儲蓄存款 25.ZA:field_220900=公庫存款
    //26.ZB:field_221000=本會支票					   
	String[] field_name_7 = {"field_1200_1502_1503","field_120700","field_120201_120202","field_420100","field_420170",
						   "field_420120","field_4201_b","field_420170_120700","field_420120_d","field_520100_220000",
						   "field_520110_220300","field_520130_220400","field_840640_220500","field_520120_220600",
						   "field_520140_220700","field_840670_220800","field_220000","field_220100","field_220200",
						   "field_220300","field_220400","field_220500","field_220600","field_220700","field_220800",
						   "field_220900","field_221000"};					   
						   
    //讀取dbdata的欄位名稱
    String[] field = {"fieldb", "fieldc", "fieldd", "fielde", "fieldf", "fieldg", "fieldh", "fieldi", "fieldj",
	                  "fieldk", "fieldl", "fieldm", "fieldn", "fieldo", "fieldp", "fieldq", "fieldr", "fields",
	                  "fieldt", "fieldu", "fieldv", "fieldw", "fieldx", "fieldy", "fieldz", "fieldza", "fieldzb"};
    //類別type=4-->利率. type=0-->實際金額	                  
	String[] field_type = {"1","1","1","2","2","2","4","4","4","4","4","4","4","4","4","4","2","2","2","2","2","2","2","2","2","2","2"};


    //m_year,m_month,bank_code,acc_code,hsien_id,type,amt
    for(int i=0;i<dbData.size();i++){
        bean=(DataObject)dbData.get(i);
        logcalendar = Calendar.getInstance();
	    nowlog = logcalendar.getTime();
	    logps.println(logformat.format(nowlog)+" "+"農漁會別:"+(String)bean.getValue("bank_type")+";機構代號:"+(String)((DataObject)dbData.get(i)).getValue("bank_code"));
	    logps.flush();

		for(int j=0;j<field.length;j++){
            dataList = new ArrayList<String>();//傳內的參數List
		    dataList.add(s_year);
		    dataList.add(s_month);
		    dataList.add((String)bean.getValue("bank_type"));
		    dataList.add((String)bean.getValue("bank_code"));
		    if(((String)bean.getValue("bank_type")).equals("6")){
		       dataList.add(field_name_6[j]);
			}else if(((String)bean.getValue("bank_type")).equals("7")){
		       dataList.add(field_name_7[j]);
		    }   
		    dataList.add((String)bean.getValue("hsien_id"));
		    dataList.add(field_type[j]); //type=4-->利率. type=0-->實際金額
		    dataList.add((bean.getValue(field[j])).toString());
		    dataList.add("");	
		    updateDBDataList.add(dataList);//1:傳內的參數List
        }
    }//儲存參數List

    //1.寫入A01_OPERATION==================================================================================================
	sqlCmd.delete(0, sqlCmd.length());
	sqlCmd.append("insert into a03_operation values(?,?,?,?,?,?,?,?,?)");
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
	dataList.add("bank_no_ALL");
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
       logps.println(logformat.format(nowlog)+" "+"產生 A03_opeation 完成");
    }else{
       logps.println(logformat.format(nowlog)+" "+"執行 A03_opeation 失敗:"+errMsg);
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