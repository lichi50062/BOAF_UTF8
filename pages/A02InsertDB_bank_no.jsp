<%
// 95.09.19-20 add 將農.漁會明細寫入a02_opeartion by 2295
// 95.09.25 fix 倍改成%-->(3)10倍=1000%,(10)1.5倍=150%
// 95.11.29 add 排程產生A02_opeation by 2295
// 97.01.29 add 990612非會員政策性農業專案貸款 K.非會員授信總額(990610)-(990511)-(990611)-(990612)
// 98.08.25 add 寫入OPERATION_LOG by 2295
// 99.01.25 fix 計算上年度信用部決算淨值/全体農(漁)會決算淨值時,需扣除992810投資全國農業金庫尚未攤提之損失[99.1月起] by 2295
// 99.11.04 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
//101.08.23 add field_990812_990813_990814違反信用部固定資產淨額不得超過上年度信用部決算淨值不在此限的原因項目 by 2295
//102.01.16 add (四)信用部逾放比< 1%  且 BIS > 10%且備抵呆帳覆蓋率高100%,已申請經主管機關同意者,得不超過200%
//          add (六)逾放比< 2% 且 BIS > 8%,已申請經主管機關同意者,得不超過150% by 2295
//102.09.23 add (四)信用部逾放比< 1%  且 BIS > 10%若備呆占狹義逾期放款比率(備抵呆帳/狹義逾放)=0,但備抵呆帳 > 0 且狹義逾放=0時,已申請經主管機關同意者,得不超過200% by 2295
//103.01.06 add (二)2.1/2.2及(八)上年度信用部決算淨值為負數時,要顯示違反 by 2295
//          add (九)逾新台幣100萬且超過前一年度信用部決算淨值5%,才算違反 by 2295
//104.02.13 add (一)月底存放比不得超過80%之規定,取消(三)限制
//          add (四)信用部逾放比< 1%  且 BIS > 10%且備抵呆帳覆蓋率高100%且 > 全體信用部備抵呆帳覆蓋率平均值且備呆占放款比率> 2%,已申請經主管機關同意者,不受限制) by 2295          
//          add (六)逾放比< 1% 且 BIS > 10% 且備抵呆帳覆蓋率高100%且 > 全體信用部備抵呆帳覆蓋率平均值且備呆占放款比率> 2%,已申請經主管機關同意者,得不超過 200%) by 2295       
//          add (七)調整為50% by 2295
//104.02.16 add (十四)對鄉(鎮、市)公所授信未經其所隸屬之縣政府保證之限額 by 2295
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
產生A02_operation
<%
   System.out.println("=============產生A02_Operation開始===========");
   String report_no = ( request.getParameter("report_no")==null ) ? "" : (String)request.getParameter("report_no");
   String s_year = ( request.getParameter("s_year")==null ) ? "" : (String)request.getParameter("s_year");
   String s_month = ( request.getParameter("s_month")==null ) ? "" : (String)request.getParameter("s_month");
   String isDebug = ( request.getParameter("isDebug")==null ) ? "" : (String)request.getParameter("isDebug");
   String lguser_id = ( request.getParameter("lguser_id")==null ) ? "" : (String)request.getParameter("lguser_id");
   String errMsg = Generate(report_no,s_year,s_month,isDebug,lguser_id);
   System.out.println("errMsg = "+errMsg);
   System.out.println("=============產生A02_Operation結束===========");
   //明細
   /* select 100 as m_year,1 as m_month,  
             a02.HSIEN_ID, nvl(a02.hsien_div_1,' '),a02.bank_type,a02.bank_code as bank_no, 
             amt990110,amt990120,amt990130,amt990140,amt990150,amt990210,amt990230,amt990240,amt990220, 
             amt990310,amt990630,amt990320,amt990410,amt990420,amt990421,amt990512,amt990610,amt990511,amt990611,amt990612, 
             amt990620,amt990621,amt990710,amt990720,amt990810,amt990811,amt990812,amt990813,amt990814,amt990910,amt990920,amt991020,amt991110, 
             amt991120,amt991210,amt991220,amt991310,amt991320,amt992810, 
             a05.amt as a05bis, 
             a01_operation.field_over_rate as a01field_over_rate, 
             a01_operation.field_backup_over_rate as a01field_backup_over_rate 
      from   
           (select wlx01.HSIEN_ID,wlx01.hsien_div_1,bn01.bank_type,a02.bank_code,  
                   sum(decode(a02.acc_code,'990110',amt,0)) as  amt990110, 
                   sum(decode(a02.acc_code,'990120',amt,0)) as  amt990120, 
                   sum(decode(a02.acc_code,'990130',amt,0)) as  amt990130, 
                   sum(decode(a02.acc_code,'990140',amt,0)) as  amt990140, 
                   sum(decode(a02.acc_code,'990150',amt,0)) as  amt990150, 
                   sum(decode(a02.acc_code,'990210',amt,0)) as  amt990210, 
                   sum(decode(a02.acc_code,'990230',amt,0)) as  amt990230, 
                   sum(decode(a02.acc_code,'990240',amt,0)) as  amt990240, 
                   sum(decode(a02.acc_code,'990220',amt,0)) as  amt990220, 
                   sum(decode(a02.acc_code,'990310',amt,0)) as  amt990310, 
                   sum(decode(a02.acc_code,'990630',amt,0)) as  amt990630, 
                   sum(decode(a02.acc_code,'990320',amt,0)) as  amt990320, 
                   sum(decode(a02.acc_code,'990410',amt,0)) as  amt990410, 
                   sum(decode(a02.acc_code,'990420',amt,0)) as  amt990420, 
                   sum(decode(a02.acc_code,'990421',amt,0)) as  amt990421, --102.01.15 add
                   sum(decode(a02.acc_code,'990512',amt,0)) as  amt990512, 
                   sum(decode(a02.acc_code,'990610',amt,0)) as  amt990610, 
                   sum(decode(a02.acc_code,'990511',amt,0)) as  amt990511, 
                   sum(decode(a02.acc_code,'990611',amt,0)) as  amt990611, 
                   sum(decode(a02.acc_code,'990612',amt,0)) as  amt990612, --97.01.29 add
                   sum(decode(a02.acc_code,'990620',amt,0)) as  amt990620, 
                   sum(decode(a02.acc_code,'990621',amt,0)) as  amt990621, --102.01.15 add
                   sum(decode(a02.acc_code,'990710',amt,0)) as  amt990710, 
                   sum(decode(a02.acc_code,'990720',amt,0)) as  amt990720, 
                   sum(decode(a02.acc_code,'990810',amt,0)) as  amt990810, 
                   sum(decode(a02.acc_code,'990811',amt,0)) as  amt990811, --101.08.23 add
                   sum(decode(a02.acc_code,'990812',amt,0)) as  amt990812, --101.08.23 add
                   sum(decode(a02.acc_code,'990813',amt,0)) as  amt990813, --101.08.23 add
                   sum(decode(a02.acc_code,'990814',amt,0)) as  amt990814, --101.08.23 add
                   sum(decode(a02.acc_code,'990910',amt,0)) as  amt990910, 
                   sum(decode(a02.acc_code,'990920',amt,0)) as  amt990920, 
                   sum(decode(a02.acc_code,'991020',amt,0)) as  amt991020, 
                   sum(decode(a02.acc_code,'991110',amt,0)) as  amt991110, 
                   sum(decode(a02.acc_code,'991120',amt,0)) as  amt991120, 
                   sum(decode(a02.acc_code,'991210',amt,0)) as  amt991210, 
                   sum(decode(a02.acc_code,'991220',amt,0)) as  amt991220, 
                   sum(decode(a02.acc_code,'991310',amt,0)) as  amt991310, 
                   sum(decode(a02.acc_code,'991320',amt,0)) as  amt991320, 
                   sum(decode(a02.acc_code,'992810',amt,0)) as  amt992810  
              from (select * from bn01 where m_year=100)bn01 
              left join (select * from a02 where a02.m_year=100 and a02.m_month=1                         
                         union 
                         select  m_year,m_month,bank_code,acc_code,amt,'' from a99 where a99.m_year=100 and a99.m_month=1
                         )a02 on a02.bank_code = bn01.bank_no 
              left join (select * from wlx01 where m_year=100)wlx01 on bn01.bank_no = wlx01.BANK_NO 
              where bn01.bank_type in ('6','7')  --and bank_code='6030016'
              and wlx01.HSIEN_ID is not null  
              group by wlx01.HSIEN_ID, wlx01.hsien_div_1,bn01.bank_type,a02.bank_code 
              order by wlx01.HSIEN_ID, wlx01.hsien_div_1,bn01.bank_type,a02.bank_code)a02 
          ,(select m_year,m_month,bank_code,acc_code,'4' as type,round(amt/1000,3) as amt,'' as violate from a05  
            where  a05.m_year =100 and a05.m_month = 1
            and a05.acc_code in ('91060P'))a05 
          ,(select m_year,m_month,bank_code, 
                   sum(decode(acc_code,'field_over_rate',amt,0)) as  field_over_rate, 
                   sum(decode(acc_code,'field_backup_over_rate',amt,0)) as  field_backup_over_rate    
            from a01_operation  
            where m_year=100 and m_month=1
            and acc_code in('field_over_rate','field_backup_over_rate')  
            group by m_year,m_month,bank_code  ) a01_operation            
         where a02.bank_code = a05.bank_code  
         and a02.bank_code=a01_operation.bank_code  
      */

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
	    String hsien_div="";
        String violate_remark="";
        
        DecimalFormat dft = new DecimalFormat("#.##");
        DecimalFormat dft_new = new DecimalFormat("###");
	    DecimalFormat dft_int = new DecimalFormat("#,###");
       
        long amt990110=0;
        long amt990120=0;
        long amt990130=0;
        long amt990140=0;
        long amt990150=0;
        long amt990210=0;
        long amt990230=0;
        long amt990240=0;
        long amt990220=0;
        long amt990310=0;
        long amt990630=0;
        long amt990320=0;
        long amt990410=0;
        long amt990420=0;
        long amt990421=0;//102.01.15 add
        long amt990422=0;//104.02.13 add
        long amt990512=0;
        long amt990610=0;
        long amt990511=0;
	    long amt990611=0;
	    long amt990612=0;//97.01.29 add
	    long amt990620=0;
	    long amt990621=0;//102.01.15 add
	    long amt990622=0;//104.02.13 add
	    long amt990710=0;
	    long amt990720=0;
	    long amt990810=0;
	    long amt990811=0;//101.08.23 add
	    long amt990812=0;//101.08.23 add
	    long amt990813=0;//101.08.23 add
	    long amt990814=0;//101.08.23 add
	    long amt990910=0;
	    long amt990920=0;
	    long amt991020=0;
	    long amt991110=0;
	    long amt991120=0;
	    long amt991210=0;
	    long amt991220=0;
	    long amt991310=0;
	    long amt991320=0;
	    long amt992810=0;//99.01.25 add
	    long amt996114=0;//104.02.16 add
	    long amt996115=0;//104.02.16 add
	    long a01field_over=0;//102.09.23 add
	    long a01field_backup=0;//102.09.23 add
	    double a05bis=0;
	    double a01field_over_rate=0;
	    double a01field_backup_over_rate=0;//102.01.15 add
	    double a01field_backup_credit_rate=0;//104.02.13 add
	    double a01field_backup_over_rate_avg=0;//104.02.13 add
	    double field_X=0;
	    String field_V_X="";
	    double field_x=0;
	    String field_W_x="";
	    String field_a_X="";
	    String field_b_x="";
	    double field_d=0;
	    String field_c_d="";
	    double field_g=0;
	    String field_f_g="";
        double field_month_dc_rate=0.0;
        double tmp_A = 0.0;
        String tmp_value = "";
        double tmp990230_990240_992810=0.0;//103.01.06 add 上年度信用部決算淨值
	    //99.11.04 add==================================================================
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
	   logfile = new File(logDir + System.getProperty("file.separator") + Report_no +"_GenerateOperation."+ logfileformat.format(nowlog));
	   System.out.println("logfile filename="+logDir + System.getProperty("file.separator") + Report_no +"_GenerateOperation."+ logfileformat.format(nowlog));
	   logos = new FileOutputStream(logfile,true);
	   logbos = new BufferedOutputStream(logos);
	   logps = new PrintStream(logbos);
       logcalendar = Calendar.getInstance();
	   nowlog = logcalendar.getTime();
	   logps.println(logformat.format(nowlog)+" "+"產生"+s_year+"年"+s_month+"月 (農.漁會-各別機構)A02_opeation 資料中");
	   logps.flush();
       //99.11.04 add 查詢年度100年以前.縣市別不同===============================
  	   cd01_table = (Integer.parseInt(s_year) < 100)?"cd01_99":"";
  	   wlx01_m_year = (Integer.parseInt(s_year) < 100)?"99":"100";
  	   //=====================================================================
       sqlCmd.append(" select "+s_year+" as m_year,"+s_month+" as m_month, ");
       sqlCmd.append(" 		a02.HSIEN_ID, nvl(a02.hsien_div_1,' '),a02.bank_type,a02.bank_code as bank_no,");
       sqlCmd.append(" 		amt990110,amt990120,amt990130,amt990140,amt990150,amt990210,amt990230,amt990240,amt990220,");
	   sqlCmd.append(" 		amt990310,amt990630,amt990320,amt990410,amt990420,amt990421,amt990422,amt990512,amt990610,amt990511,amt990611,amt990612,");
	   sqlCmd.append(" 		amt990620,amt990621,amt990622,amt990710,amt990720,amt990810,amt990811,amt990812,amt990813,amt990814,amt990910,amt990920,amt991020,amt991110,");
	   sqlCmd.append(" 		amt991120,amt991210,amt991220,amt991310,amt991320,amt992810,amt996114,amt996115,");	   
       sqlCmd.append(" 		a05.amt as a05bis,");
	   sqlCmd.append("		a01_operation.field_over_rate as a01field_over_rate,");
	   sqlCmd.append("		a01_operation.field_backup_over_rate as a01field_backup_over_rate,");//備呆占狹義逾期放款比率 
	   sqlCmd.append("		a01_operation.field_over as a01field_over,");
	   sqlCmd.append("		a01_operation.field_backup as a01field_backup,");
	   sqlCmd.append("		a01_operation.field_backup_credit_rate as a01field_backup_credit_rate,");//104.02.13 add 備呆占放款比率
	   sqlCmd.append("		a01_operation_sum.field_backup_over_rate as a01field_backup_over_rate_avg");//104.02.13 add全體信用部備呆占狹義逾期放款比率平均值
	   sqlCmd.append(" from  ");
 	   sqlCmd.append("      (select wlx01.HSIEN_ID,wlx01.hsien_div_1,bn01.bank_type,a02.bank_code, ");
	   sqlCmd.append("              sum(decode(a02.acc_code,'990110',amt,0)) as  amt990110,");
	   sqlCmd.append("              sum(decode(a02.acc_code,'990120',amt,0)) as  amt990120,");
	   sqlCmd.append("              sum(decode(a02.acc_code,'990130',amt,0)) as  amt990130,");
	   sqlCmd.append("              sum(decode(a02.acc_code,'990140',amt,0)) as  amt990140,");
	   sqlCmd.append("              sum(decode(a02.acc_code,'990150',amt,0)) as  amt990150,");
	   sqlCmd.append("              sum(decode(a02.acc_code,'990210',amt,0)) as  amt990210,");
	   sqlCmd.append("              sum(decode(a02.acc_code,'990230',amt,0)) as  amt990230,");
	   sqlCmd.append("              sum(decode(a02.acc_code,'990240',amt,0)) as  amt990240,");
       sqlCmd.append("              sum(decode(a02.acc_code,'990220',amt,0)) as  amt990220,");
	   sqlCmd.append("              sum(decode(a02.acc_code,'990310',amt,0)) as  amt990310,");
	   sqlCmd.append("              sum(decode(a02.acc_code,'990630',amt,0)) as  amt990630,");
	   sqlCmd.append("              sum(decode(a02.acc_code,'990320',amt,0)) as  amt990320,");
	   sqlCmd.append("              sum(decode(a02.acc_code,'990410',amt,0)) as  amt990410,");
	   sqlCmd.append("              sum(decode(a02.acc_code,'990420',amt,0)) as  amt990420,");
	   sqlCmd.append("              sum(decode(a02.acc_code,'990421',amt,0)) as  amt990421,");//102.01.15 add
	   sqlCmd.append("              sum(decode(a02.acc_code,'990422',amt,0)) as  amt990422,");//104.02.13 add
	   sqlCmd.append("              sum(decode(a02.acc_code,'990512',amt,0)) as  amt990512,");
	   sqlCmd.append("              sum(decode(a02.acc_code,'990610',amt,0)) as  amt990610,");
	   sqlCmd.append("              sum(decode(a02.acc_code,'990511',amt,0)) as  amt990511,");
       sqlCmd.append("              sum(decode(a02.acc_code,'990611',amt,0)) as  amt990611,");
       sqlCmd.append("              sum(decode(a02.acc_code,'990612',amt,0)) as  amt990612,");//97.01.29 add
	   sqlCmd.append("              sum(decode(a02.acc_code,'990620',amt,0)) as  amt990620,");
	   sqlCmd.append("              sum(decode(a02.acc_code,'990621',amt,0)) as  amt990621,");//102.01.15 add
	   sqlCmd.append("              sum(decode(a02.acc_code,'990622',amt,0)) as  amt990622,");//104.02.13 add
	   sqlCmd.append("              sum(decode(a02.acc_code,'990710',amt,0)) as  amt990710,");
	   sqlCmd.append("              sum(decode(a02.acc_code,'990720',amt,0)) as  amt990720,");
	   sqlCmd.append("              sum(decode(a02.acc_code,'990810',amt,0)) as  amt990810,");
	   sqlCmd.append("              sum(decode(a02.acc_code,'990811',amt,0)) as  amt990811,");//101.08.23 add
	   sqlCmd.append("              sum(decode(a02.acc_code,'990812',amt,0)) as  amt990812,");//101.08.23 add
	   sqlCmd.append("              sum(decode(a02.acc_code,'990813',amt,0)) as  amt990813,");//101.08.23 add
	   sqlCmd.append("              sum(decode(a02.acc_code,'990814',amt,0)) as  amt990814,");//101.08.23 add
	   sqlCmd.append("              sum(decode(a02.acc_code,'990910',amt,0)) as  amt990910,");
	   sqlCmd.append("              sum(decode(a02.acc_code,'990920',amt,0)) as  amt990920,");
	   sqlCmd.append("              sum(decode(a02.acc_code,'991020',amt,0)) as  amt991020,");
       sqlCmd.append("              sum(decode(a02.acc_code,'991110',amt,0)) as  amt991110,");
	   sqlCmd.append("              sum(decode(a02.acc_code,'991120',amt,0)) as  amt991120,");
	   sqlCmd.append("              sum(decode(a02.acc_code,'991210',amt,0)) as  amt991210,");
	   sqlCmd.append("              sum(decode(a02.acc_code,'991220',amt,0)) as  amt991220,");
	   sqlCmd.append("              sum(decode(a02.acc_code,'991310',amt,0)) as  amt991310,");
	   sqlCmd.append("              sum(decode(a02.acc_code,'991320',amt,0)) as  amt991320,");
	   sqlCmd.append("              sum(decode(a02.acc_code,'992810',amt,0)) as  amt992810,");
	   sqlCmd.append("              sum(decode(a02.acc_code,'996114',amt,0)) as  amt996114,");//104.02.16 add
	   sqlCmd.append("              sum(decode(a02.acc_code,'996115',amt,0)) as  amt996115 ");//104.02.16 add
	   sqlCmd.append(" 		from (select * from bn01 where m_year=?)bn01 left join (select * from a02 where a02.m_year=? and a02.m_month=?");
	   paramList.add(wlx01_m_year);
	   paramList.add(s_year);
	   paramList.add(s_month);
	   sqlCmd.append("                             union");
	   sqlCmd.append("                             select  m_year,m_month,bank_code,acc_code,amt,'' from a99 where a99.m_year=? and a99.m_month=?");
	   paramList.add(s_year);
	   paramList.add(s_month);
	   sqlCmd.append("                            )a02 on a02.bank_code = bn01.bank_no");
	   sqlCmd.append(" 		left join (select * from wlx01 where m_year=?)wlx01 on bn01.bank_no = wlx01.BANK_NO");
	   paramList.add(wlx01_m_year);
	   sqlCmd.append(" 		where bn01.bank_type in ('6','7')");// and bn01.bank_no in('6220077')"
	   sqlCmd.append(" 		and wlx01.HSIEN_ID is not null ");
	   sqlCmd.append(" 		group by wlx01.HSIEN_ID, wlx01.hsien_div_1,bn01.bank_type,a02.bank_code");
	   sqlCmd.append(" 		order by wlx01.HSIEN_ID, wlx01.hsien_div_1,bn01.bank_type,a02.bank_code)a02");
	   sqlCmd.append("	 ,(select m_year,m_month,bank_code,acc_code,'4' as type,round(amt/1000,3) as amt,'' as violate from a05 ");
  	   sqlCmd.append("	  where  a05.m_year =? and a05.m_month = ?");
  	   paramList.add(s_year);
	   paramList.add(s_month);
   	   sqlCmd.append("	  and a05.acc_code in ('91060P'))a05");
   	   sqlCmd.append("    ,(select m_year,m_month,bank_code,");
   	   sqlCmd.append("     sum(decode(acc_code,'field_over_rate',amt,0)) as  field_over_rate,");
       sqlCmd.append("     sum(decode(acc_code,'field_backup_over_rate',amt,0)) as  field_backup_over_rate, ");  
       sqlCmd.append("     sum(decode(acc_code,'field_over',amt,0)) as  field_over, ");//逾放金額  102.09.23 add
       sqlCmd.append("     sum(decode(acc_code,'field_backup',amt,0)) as  field_backup, "); //備抵呆帳 102.09.23 add
       sqlCmd.append("     sum(decode(acc_code,'field_backup_credit_rate',amt,0)) as  field_backup_credit_rate "); //備呆占放款比率 104.02.13 add 
       sqlCmd.append("     from a01_operation ");
       sqlCmd.append("     where m_year=? and m_month=?");
       paramList.add(s_year);
	   paramList.add(s_month);
       sqlCmd.append("     and acc_code in('field_over_rate','field_backup_over_rate','field_over','field_backup','field_backup_credit_rate') ");
       sqlCmd.append("     group by m_year,m_month,bank_code  ) a01_operation, ");
       sqlCmd.append("    (select m_year,m_month,");   	 
       sqlCmd.append("     sum(decode(acc_code,'field_backup_over_rate',amt,0)) as  field_backup_over_rate ");//104.02.13 add備呆占狹義逾期放款比率         
       sqlCmd.append("     from a01_operation ");
       sqlCmd.append("     where m_year=? and m_month=?");
       paramList.add(s_year);
	   paramList.add(s_month);
       sqlCmd.append("     and acc_code in('field_backup_over_rate') ");
       sqlCmd.append("     and bank_type='ALL' ");
       sqlCmd.append("     and bank_code='ALL' ");
       sqlCmd.append("     and hsien_id=' '");
       sqlCmd.append("     group by m_year,m_month ) a01_operation_sum ");   	   
   	   /*102.01.15
	   sqlCmd.append("	 ,(select m_year,m_month,bank_code,acc_code,type,amt,'' as violate from a01_operation ");
  	   sqlCmd.append(" 	   where m_year=? and m_month=?");
  	   paramList.add(s_year);
	   paramList.add(s_month);
  	   sqlCmd.append("	   and acc_code in('field_over_rate')) a01_operation ");
  	  */ 
	   sqlCmd.append("	where a02.bank_code = a05.bank_code ");
	   sqlCmd.append("    and a02.bank_code=a01_operation.bank_code ");

    List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt990110,amt990120,amt990130,amt990140,amt990150,amt990210,amt990230,amt990240,amt990220,"
			    + "amt990310,amt990630,amt990320,amt990410,amt990420,amt990421,amt990422,amt990512,amt990610,amt990511,amt990611,amt990612,"
			    + "amt990620,amt990621,amt990622,amt990710,amt990720,amt990810,amt990811,amt990812,amt990813,amt990814,amt990910,amt990920,amt991020,amt991110,"
			    + "amt991120,amt991210,amt991220,amt991310,amt991320,amt992810,amt996114,amt996115,a05bis,a01field_over_rate,a01field_backup_over_rate,a01field_over,a01field_backup,a01field_backup_credit_rate,a01field_backup_over_rate_avg");
    System.out.println("dbData.size()="+dbData.size());
    DataObject bean = null;
    for(int i=0;i<dbData.size();i++){
        bean=(DataObject)dbData.get(i);
        logcalendar = Calendar.getInstance();
	    nowlog = logcalendar.getTime();
	    logps.println(logformat.format(nowlog)+" "+"機構代號:"+(String)((DataObject)dbData.get(i)).getValue("bank_no"));
	    logps.flush();

        amt990110 = Long.parseLong( ( bean.getValue("amt990110")).toString());
        amt990120 = Long.parseLong( ( bean.getValue("amt990120")).toString());
        amt990130 = Long.parseLong( ( bean.getValue("amt990130")).toString());
        amt990140 = Long.parseLong( ( bean.getValue("amt990140")).toString());
        amt990150 = Long.parseLong( ( bean.getValue("amt990150")).toString());
        System.out.println("amt990150="+amt990150);
        amt990210 = Long.parseLong( ( bean.getValue("amt990210")).toString());
        amt990230 = Long.parseLong( ( bean.getValue("amt990230")).toString());
        amt990240 = Long.parseLong( ( bean.getValue("amt990240")).toString());
        amt990220 = Long.parseLong( ( bean.getValue("amt990220")).toString());
        amt990310 = Long.parseLong( ( bean.getValue("amt990310")).toString());
        amt990630 = Long.parseLong( ( bean.getValue("amt990630")).toString());
        amt990320 = Long.parseLong( ( bean.getValue("amt990320")).toString());
        amt990410 = Long.parseLong( ( bean.getValue("amt990410")).toString());
        amt990420 = Long.parseLong( ( bean.getValue("amt990420")).toString());
        amt990421 = Long.parseLong( ( bean.getValue("amt990421")).toString());//102.01.15 add
        amt990422 = Long.parseLong( ( bean.getValue("amt990422")).toString());//104.02.13 add
        System.out.println("amt990420="+amt990420);
        amt990512 = Long.parseLong( ( bean.getValue("amt990512")).toString());
        amt990610 = Long.parseLong( ( bean.getValue("amt990610")).toString());
        amt990511 = Long.parseLong( ( bean.getValue("amt990511")).toString());
        amt990611 = Long.parseLong( ( bean.getValue("amt990611")).toString());
        amt990612 = Long.parseLong( ( bean.getValue("amt990612")).toString());//97.01.29 add
        amt990620 = Long.parseLong( ( bean.getValue("amt990620")).toString());
        amt990621 = Long.parseLong( ( bean.getValue("amt990621")).toString());//102.01.15 add
        amt990622 = Long.parseLong( ( bean.getValue("amt990622")).toString());//104.02.13 add
        amt990710 = Long.parseLong( ( bean.getValue("amt990710")).toString());
        amt990720 = Long.parseLong( ( bean.getValue("amt990720")).toString());
        amt990810 = Long.parseLong( ( bean.getValue("amt990810")).toString());
        amt990811 = Long.parseLong( ( bean.getValue("amt990811")).toString());//101.08.23 add
        amt990812 = Long.parseLong( ( bean.getValue("amt990812")).toString());//101.08.23 add
        amt990813 = Long.parseLong( ( bean.getValue("amt990813")).toString());//101.08.23 add
        amt990814 = Long.parseLong( ( bean.getValue("amt990814")).toString());//101.08.23 add
        amt990910 = Long.parseLong( ( bean.getValue("amt990910")).toString());
        System.out.println("amt990910="+amt990910);
        amt990920 = Long.parseLong( ( bean.getValue("amt990920")).toString());
        amt991020 = Long.parseLong( ( bean.getValue("amt991020")).toString());
        amt991110 = Long.parseLong( ( bean.getValue("amt991110")).toString());
        amt991120 = Long.parseLong( ( bean.getValue("amt991120")).toString());
        amt991210 = Long.parseLong( ( bean.getValue("amt991210")).toString());
        amt991220 = Long.parseLong( ( bean.getValue("amt991220")).toString());
        amt991310 = Long.parseLong( ( bean.getValue("amt991310")).toString());
        amt991320 = Long.parseLong( ( bean.getValue("amt991320")).toString());
        System.out.println("amt991320="+amt991320);
        amt992810 = Long.parseLong( ( bean.getValue("amt992810")).toString());//99.01.25
        amt996114 = Long.parseLong( ( bean.getValue("amt996114")).toString());//104.02.16
        amt996115 = Long.parseLong( ( bean.getValue("amt996115")).toString());//104.02.16
        System.out.println("amt992810="+amt992810);
        a01field_over = Long.parseLong( ( bean.getValue("a01field_over")).toString());
        a01field_backup = Long.parseLong( ( bean.getValue("a01field_backup")).toString());
        a05bis=0;
        a01field_over_rate =0;
        a01field_backup_over_rate = 0;
        a05bis = Double.parseDouble( ( bean.getValue("a05bis")).toString());
        a01field_over_rate = Double.parseDouble( ( bean.getValue("a01field_over_rate")).toString());
		a01field_backup_over_rate = Double.parseDouble( ( bean.getValue("a01field_backup_over_rate")).toString());		
		a01field_backup_credit_rate = Double.parseDouble( ( bean.getValue("a01field_backup_credit_rate")).toString());//104.02.13 add	
		a01field_backup_over_rate_avg = Double.parseDouble( ( bean.getValue("a01field_backup_over_rate_avg")).toString());//104.02.13 add		
        hsien_div = (String)bean.getValue("hsien_div_1");
        tmp_A = 0.0;
	    field_X=0.0;
	    field_V_X="N";
	    field_x=0.0;
	    field_W_x="N";
	    field_a_X="N";
	    field_b_x="N";
	    field_d=0.0;
	    field_c_d="N";
	    field_g=0.0;
	    field_f_g="N";
        field_month_dc_rate=0.0;
        //1.鄉鎮地區農會不得超過80%，直轄市及省、縣轄市農會不得超過78%
        //field_month_dc_rate=B.月底存放比率
        //104.02.13 一律調整為80%
        if(isDebug.equals("true")){
           System.out.println("bank_code="+(String)bean.getValue("bank_no"));
        }
        if(amt990120 > amt990130){
           if(isDebug.equals("true")) System.out.println("(A-B+C)/(D-E*1/2)");
           field_month_dc_rate = (double)(((double)amt990110-(double)amt990120+(double)amt990130)/((double)amt990140-(double)amt990150/2));

        }else{
           if(isDebug.equals("true")) System.out.println("A/(D-E*1/2)");
           field_month_dc_rate = (double)((double)amt990110/((double)amt990140-(double)amt990150/2));
        }
        if(isDebug.equals("true")) System.out.println("field_month_dc_rate="+field_month_dc_rate);
        if(isDebug.equals("true")) System.out.println("field_month_dc_rate="+Math.round( field_month_dc_rate * 10000));
        long tmp_B =Math.round( field_month_dc_rate * 10000);
        if(isDebug.equals("true")) System.out.println(((double)tmp_B)/100);
        field_month_dc_rate=((double)tmp_B)/100;
        if(isDebug.equals("true")) System.out.println("field_month_dc_rate="+field_month_dc_rate);
        violate_remark="N";
        if(hsien_div == null || hsien_div.equals("")){
           violate_remark="N";
        /*104.02.13 調整都以80%為限   
        }else if(hsien_div.equals("1")){//直轄市
           if(field_month_dc_rate > 78){
              violate_remark="Y";
           }
        }else if(hsien_div.equals("2")){//鄉鎮地區
        */
    	}else{	
           if(field_month_dc_rate > 80){
              violate_remark="Y";
           }
        }
        dataList = new ArrayList<String>();//傳內的參數List
		dataList.add(s_year);
		dataList.add(s_month);
		dataList.add((String)bean.getValue("bank_type"));
		dataList.add((String)bean.getValue("bank_no"));
		dataList.add("field_month_dc_rate");
		dataList.add((String)bean.getValue("hsien_id"));
		dataList.add("4"); //type=4-->利率. type=2-->加總
		dataList.add(String.valueOf(field_month_dc_rate));
		dataList.add(violate_remark);
		updateDBDataList.add(dataList);//1:傳內的參數List
		/*
        sqlCmd = "insert into a02_operation values("+s_year+","+s_month+",'"+(String)bean.getValue("bank_type")+"','"
               + (String)bean.getValue("bank_no")+"','field_month_dc_rate','"
               + (String)bean.getValue("hsien_id")+"','4',"
               + field_month_dc_rate+",'"+violate_remark+"')";
        //System.out.println("test1="+sqlCmd);
        updateDBSqlList.add(sqlCmd);
        */
        //2.1違反內部融資不得超過信用部上年度決算淨值60%之規定
        //field_990230-990240=D.上年度信用部決算淨值(扣除檢查應補提未提足之備抵呆帳)(990230)-(990240)
        //99.01.25 fix 99.01開始.上年度信用部決算淨值扣除992810投資全國農業金庫尚未攤提之損失
        //103.01.06 add 上年度信用部決算淨值為負數時,要顯示違反
        tmp_A=0.0;
        tmp_A=(double)amt990230-(double)amt990240-(double)amt992810;
        tmp990230_990240_992810 = tmp_A;//103.01.06 add
        if(isDebug.equals("true")) System.out.println("field_990230-990240="+tmp_A);
        dataList = new ArrayList<String>();//傳內的參數List
		dataList.add(s_year);
		dataList.add(s_month);
		dataList.add((String)bean.getValue("bank_type"));
		dataList.add((String)bean.getValue("bank_no"));
		dataList.add("field_990230-990240");
		dataList.add((String)bean.getValue("hsien_id"));
		dataList.add("4"); //type=4-->利率. type=2-->加總
		dataList.add(String.valueOf(tmp_A));
		dataList.add("");
		updateDBDataList.add(dataList);//1:傳內的參數List
		/*
        sqlCmd = "insert into a02_operation values("+s_year+","+s_month+",'"+(String)((DataObject)dbData.get(i)).getValue("bank_type")+"','"
               + (String)((DataObject)dbData.get(i)).getValue("bank_no")+"','field_990230-990240','"
               + (String)((DataObject)dbData.get(i)).getValue("hsien_id")+"','0',"
               + tmp_A+",'')";
        //System.out.println("test2="+sqlCmd);
        updateDBSqlList.add(sqlCmd);
        */
        //field_990210/(990230-990240)=C/D
        tmp_A = 0.0;
        if(((double)amt990230-(double)amt990240-(double)amt992810) != 0){        	
           tmp_A=(double)amt990210/((double)amt990230-(double)amt990240-(double)amt992810);
        }
        if(isDebug.equals("true")) System.out.println("field_990210/(990230-990240-992810)="+tmp_A);
        tmp_B =Math.round( tmp_A * 10000);
        if(isDebug.equals("true")) System.out.println(((double)tmp_B)/100);
        tmp_A=((double)tmp_B)/100;
        if(isDebug.equals("true")) System.out.println("field_990210/(990230-990240-992810)="+ tmp_A);
        violate_remark="N";
        if(tmp_A > 60){
           violate_remark="Y";
        }
        if(tmp990230_990240_992810 < 0.0){//103.01.06 add 上年度信用部決算淨值為負數時,要顯示違反
    	   violate_remark="Y";
    	}	
        dataList = new ArrayList<String>();//傳內的參數List
        dataList.add(s_year);
        dataList.add(s_month);
        dataList.add((String)bean.getValue("bank_type"));
        dataList.add((String)bean.getValue("bank_no"));
        dataList.add("field_990210/(990230-990240)");
        dataList.add((String)bean.getValue("hsien_id"));
        dataList.add("4"); //type=4-->利率. type=2-->加總
        dataList.add(String.valueOf(tmp_A));
        dataList.add(violate_remark);
        updateDBDataList.add(dataList);//1:傳內的參數List
		/*
        sqlCmd = "insert into a02_operation values("+s_year+","+s_month+",'"+(String)((DataObject)dbData.get(i)).getValue("bank_type")+"','"
               + (String)((DataObject)dbData.get(i)).getValue("bank_no")+"','field_990210/(990230-990240)','"
               + (String)((DataObject)dbData.get(i)).getValue("hsien_id")+"','4',"
               + tmp_A+",'"+violate_remark+"')";
        //System.out.println("test3="+sqlCmd);
        updateDBSqlList.add(sqlCmd);
        */
        //2.2違反內部融資(中長期)不得超過信用部上年度決算淨值30%之規定
        //103.01.06 add 上年度信用部決算淨值為負數時,要顯示違反
        //field_990220/(990230-990240)=E/D
        tmp_A = 0.0;
        //99.01.25 fix 99.01開始.上年度信用部決算淨值扣除992810投資全國農業金庫尚未攤提之損失
        if(((double)amt990230-(double)amt990240-(double)amt992810) != 0 ){
           tmp_A=(double)amt990220/((double)amt990230-(double)amt990240-(double)amt992810);
        }
        if(isDebug.equals("true")) System.out.println("field_990220/(990230-990240-992810)="+ tmp_A);
        tmp_B =Math.round( tmp_A * 10000);
        if(isDebug.equals("true")) System.out.println(((double)tmp_B)/100);
        tmp_A=((double)tmp_B)/100;
        if(isDebug.equals("true")) System.out.println("field_990220/(990230-990240-992810)="+ tmp_A);
        violate_remark="N";
        if(tmp_A > 30){
           violate_remark="Y";
        }
        if(tmp990230_990240_992810 < 0.0){//103.01.06 add 上年度信用部決算淨值為負數時,要顯示違反
    	   violate_remark="Y";
    	}	
        dataList = new ArrayList<String>();//傳內的參數List
        dataList.add(s_year);
        dataList.add(s_month);
        dataList.add((String)bean.getValue("bank_type"));
        dataList.add((String)bean.getValue("bank_no"));
        dataList.add("field_990220/(990230-990240)");
        dataList.add((String)bean.getValue("hsien_id"));
        dataList.add("4"); //type=4-->利率. type=2-->加總
        dataList.add(String.valueOf(tmp_A));
        dataList.add(violate_remark);
        updateDBDataList.add(dataList);//1:傳內的參數List
        /*
        sqlCmd = "insert into a02_operation values("+s_year+","+s_month+",'"+(String)((DataObject)dbData.get(i)).getValue("bank_type")+"','"
               + (String)((DataObject)dbData.get(i)).getValue("bank_no")+"','field_990220/(990230-990240)','"
               + (String)((DataObject)dbData.get(i)).getValue("hsien_id")+"','4',"
               + tmp_A+",'"+violate_remark+"')";
        //System.out.println("test4="+sqlCmd);
        updateDBSqlList.add(sqlCmd);
		*/

        //3.違反非會員存款總額不得超過上年度農漁會全體決算淨值10倍之規定
        //field_990310-(990630)/2=F.非會員存款總額(990310)-1/2(990630)
        //95.09.25 fix 倍改成%-->10倍=1000%
        //104.02.13 add 取消(三)限制
        tmp_A = 0.0;
        tmp_A=(double)amt990310-((double)amt990630/2);
        if(isDebug.equals("true")) System.out.println("field_990310-(990630)/2="+ tmp_A);
        dataList = new ArrayList<String>();//傳內的參數List
        dataList.add(s_year);
        dataList.add(s_month);
        dataList.add((String)bean.getValue("bank_type"));
        dataList.add((String)bean.getValue("bank_no"));
        dataList.add("field_990310-(990630)/2");
        dataList.add((String)bean.getValue("hsien_id"));
        dataList.add("0"); //type=4-->利率. type=2-->加總
        dataList.add(String.valueOf(tmp_A));
        dataList.add("");
        updateDBDataList.add(dataList);//1:傳內的參數List
        /*
        sqlCmd = "insert into a02_operation values("+s_year+","+s_month+",'"+(String)((DataObject)dbData.get(i)).getValue("bank_type")+"','"
               + (String)((DataObject)dbData.get(i)).getValue("bank_no")+"','field_990310-(990630)/2','"
               + (String)((DataObject)dbData.get(i)).getValue("hsien_id")+"','0',"
               + tmp_A+",'')";
        //System.out.println("test5="+sqlCmd);
        updateDBSqlList.add(sqlCmd);
        */
        //field_f/g=F/G
        if(isDebug.equals("true")) System.out.println("990310-(990630)/2="+tmp_A);
        tmp_A=0.0;
        if(isDebug.equals("true")) System.out.println("amt990320="+amt990320);
        //99.01.25 fix 99.01開始.全体農/漁會上年度決算淨值扣除992810投資全國農業金庫尚未攤提之損失
        if((amt990320-amt992810)!=0){
           tmp_A = ((double)amt990310-((double)amt990630/2)) / (amt990320-amt992810);
           if(isDebug.equals("true")) System.out.println("field_990310-(990630)/2 / (990320-992810)="+tmp_A);
           tmp_value = dft.format(tmp_A);
           if(isDebug.equals("true")) System.out.println("tmp_value="+tmp_value);
           //95.09.25 fix 倍改成%-->10倍=1000%
           tmp_A = Double.parseDouble(tmp_value);
           tmp_B =Math.round( tmp_A * 10000);
           if(isDebug.equals("true")) System.out.println(((double)tmp_B)/100);
           tmp_A=((double)tmp_B)/100;
           if(isDebug.equals("true")) System.out.println("field_990310-(990630)/2 / (990320-992810)="+ tmp_A);
        }
        violate_remark="N";
        //95.09.25 fix 倍改成%-->10倍=1000%
        /*104.02.13 add 取消(三)限制
        if(tmp_A > 1000){
           violate_remark="Y";
        }
        */
        dataList = new ArrayList<String>();//傳內的參數List
        dataList.add(s_year);
        dataList.add(s_month);
        dataList.add((String)bean.getValue("bank_type"));
        dataList.add((String)bean.getValue("bank_no"));
        dataList.add("field_f/g");
        dataList.add((String)bean.getValue("hsien_id"));
        dataList.add("4"); //type=4-->利率. type=2-->加總
        dataList.add(String.valueOf(tmp_A));
        dataList.add(violate_remark);
        updateDBDataList.add(dataList);//1:傳內的參數List
        /*
        sqlCmd = "insert into a02_operation values("+s_year+","+s_month+",'"+(String)((DataObject)dbData.get(i)).getValue("bank_type")+"','"
               + (String)((DataObject)dbData.get(i)).getValue("bank_no")+"','field_f/g','"
               + (String)((DataObject)dbData.get(i)).getValue("hsien_id")+"','4',"
               + tmp_A+",'"+violate_remark+"')";
        //System.out.println("test6="+sqlCmd);
        updateDBSqlList.add(sqlCmd);
        */
        //4.違反其對全部贊助會員授信總額占贊助會員存款總額之比率不得超過100%之規定
        //(信用部逾放比< 2%  且 BIS > 8%者 得不超過 150%)
        //102.01.15 add 信用部逾放比< 1%  且 BIS > 10%且備抵呆帳覆蓋率高100%,已申請經主管機關同意者,得不超過200%			      
        //104.02.13 add 信用部逾放比< 1%  且 BIS > 10%且備抵呆帳覆蓋率高100%且 > 全體信用部備抵呆帳覆蓋率平均值且備呆占放款比率> 2%,已申請經主管機關同意者,不受限制)
        //field_990410/990420=H/I
        tmp_A = 0.0;
        if(amt990420 != 0){
           tmp_A = (double)amt990410 / (double)amt990420;
           if(isDebug.equals("true")) System.out.println("(double)amt990410 / (double)amt990420="+tmp_A);
           tmp_B =Math.round( tmp_A * 10000);
           if(isDebug.equals("true")) System.out.println(((double)tmp_B)/100);
           tmp_A=((double)tmp_B)/100;
           if(isDebug.equals("true")) System.out.println("(double)amt990410 / (double)amt990420="+ tmp_A);
        }
        violate_remark="N";
        if(a01field_over_rate < 2 && a05bis > 8){
           if(a01field_over_rate < 1 && a05bis > 10){
           	  if((a01field_backup_over_rate == 0 && a01field_backup > 0 && a01field_over == 0 )//102.09.23 add 若備呆占狹義逾期放款比率(備抵呆帳/狹義逾放)=0,但備抵呆帳 > 0 且狹義逾放=0時,可>200
           	   ||(a01field_backup_over_rate > 100)){//102.01.15 add 990421  
           	   	if(a01field_backup_over_rate > a01field_backup_over_rate_avg && a01field_backup_credit_rate > 2 && amt990422 > 0){//104.02.13 add
           	   	   if(isDebug.equals("true")) System.out.println("a01field_over_rate < 1 && a05bis > 10 && a01field_backup_over_rate > 100 && a01field_backup_over_rate > a01field_backup_over_rate_avg && a01field_backup_credit_rate > 2 && amt990422 > 0-->a01field_over_rate="+a01field_over_rate+":a05bis="+a05bis+":a01field_backup_over_rate="+a01field_backup_over_rate+":a01field_backup_over_rate_avg="+a01field_backup_over_rate_avg+":a01field_backup_credit_rate="+a01field_backup_credit_rate+":amt990422="+amt990422);          
           	    }else if(amt990421 > 0){//104.02.13 add        	  
                  if(isDebug.equals("true")) System.out.println("a01field_over_rate < 1 && a05bis > 10 && a01field_backup_over_rate > 100-->a01field_over_rate="+a01field_over_rate+":a05bis="+a05bis+":a01field_backup_over_rate="+a01field_backup_over_rate+":amt990421="+amt990421);          
                  if(tmp_A > 200){
                     violate_remark="Y";
                  }
                }else {
                  if(isDebug.equals("true")) System.out.println("a01field_over_rate < 1 && a05bis > 10 && a01field_backup_over_rate > 100-->a01field_over_rate="+a01field_over_rate+":a05bis="+a05bis+":a01field_backup_over_rate="+a01field_backup_over_rate+":amt990421="+amt990421);          
                  if(tmp_A > 150){
                     violate_remark="Y";
                  }                              
                }		  
              }else{
                if(isDebug.equals("true")) System.out.println("a01field_over_rate < 1 && a05bis > 10 && a01field_backup_over_rate < 100-->a01field_over_rate="+a01field_over_rate+":a05bis="+a05bis+":a01field_backup_over_rate="+a01field_backup_over_rate+":amt990421="+amt990421);          
                if(tmp_A > 150){
                 violate_remark="Y";
                }              
              }	              
           }else{   
              if(isDebug.equals("true")) System.out.println("a01field_over_rate < 2 && a05bis > 8-->a01field_over_rate="+a01field_over_rate+":a05bis="+a05bis);          
              if(tmp_A > 150){
                 violate_remark="Y";
              }
           }       
        }else if(tmp_A > 100){
           violate_remark="Y";
        }
        dataList = new ArrayList<String>();//傳內的參數List
        dataList.add(s_year);
        dataList.add(s_month);
        dataList.add((String)bean.getValue("bank_type"));
        dataList.add((String)bean.getValue("bank_no"));
        dataList.add("field_990410/990420");
        dataList.add((String)bean.getValue("hsien_id"));
        dataList.add("4"); //type=4-->利率. type=2-->加總
        dataList.add(String.valueOf(tmp_A));
        dataList.add(violate_remark);
        updateDBDataList.add(dataList);//1:傳內的參數List
        //104.02.13 add全體信用部備呆占狹義逾期放款比率平均值
        dataList = new ArrayList<String>();//傳內的參數List
        dataList.add(s_year);
        dataList.add(s_month);
        dataList.add((String)bean.getValue("bank_type"));
        dataList.add((String)bean.getValue("bank_no"));
        dataList.add("field_backup_over_rate_avg");
        dataList.add((String)bean.getValue("hsien_id"));
        dataList.add("4"); //type=4-->利率. type=2-->加總
        dataList.add(String.valueOf(a01field_backup_over_rate_avg));
        dataList.add("N");
        updateDBDataList.add(dataList);//1:傳內的參數List
        
        /*
        sqlCmd = "insert into a02_operation values("+s_year+","+s_month+",'"+(String)((DataObject)dbData.get(i)).getValue("bank_type")+"','"
               + (String)((DataObject)dbData.get(i)).getValue("bank_no")+"','field_990410/990420','"
               + (String)((DataObject)dbData.get(i)).getValue("hsien_id")+"','4',"
               + tmp_A+",'"+violate_remark+"')";
        //System.out.println("test7="+sqlCmd);
        updateDBSqlList.add(sqlCmd);
        */
        //5.違反非會員無擔保消費性貸款占農漁會上年度決算淨值之比率不得超過100%之規定
        //field_990512/990320=J/G
        tmp_A = 0.0;
        //99.01.25 fix 99.01開始.全体農/漁會上年度決算淨值扣除992810投資全國農業金庫尚未攤提之損失
        if((amt990320-amt992810)!=0){
           tmp_A=(double)amt990512/((double)amt990320-(double)amt992810);
           if(isDebug.equals("true")) System.out.println("(double)amt990512/((double)amt990320-(double)amt992810)="+tmp_A);
           tmp_B =Math.round( tmp_A * 10000);
           if(isDebug.equals("true")) System.out.println(((double)tmp_B)/100);
           tmp_A=((double)tmp_B)/100;
           if(isDebug.equals("true")) System.out.println("(double)amt990512/((double)amt990320-(double)amt992810)="+ tmp_A);
        }
        violate_remark="N";
        if(tmp_A > 100){
           violate_remark="Y";
        }
        dataList = new ArrayList<String>();//傳內的參數List
        dataList.add(s_year);
        dataList.add(s_month);
        dataList.add((String)bean.getValue("bank_type"));
        dataList.add((String)bean.getValue("bank_no"));
        dataList.add("field_990512/990320");
        dataList.add((String)bean.getValue("hsien_id"));
        dataList.add("4"); //type=4-->利率. type=2-->加總
        dataList.add(String.valueOf(tmp_A));
        dataList.add(violate_remark);
        updateDBDataList.add(dataList);//1:傳內的參數List
        /*
        sqlCmd = "insert into a02_operation values("+s_year+","+s_month+",'"+(String)((DataObject)dbData.get(i)).getValue("bank_type")+"','"
               + (String)((DataObject)dbData.get(i)).getValue("bank_no")+"','field_990512/990320','"
               + (String)((DataObject)dbData.get(i)).getValue("hsien_id")+"','4',"
               + tmp_A+",'"+violate_remark+"')";
        //System.out.println("test8="+sqlCmd);
        updateDBSqlList.add(sqlCmd);
        */
        //6.違反非會員授信總額占非會員存款總額比率不得超過100%之規定
        //field_990610-990511-990611=K.非會員授信總額(990610)-(990511)-(990611)
        //102.01.15 add 逾放比< 2% 且 BIS > 8%,已申請經主管機關同意者,得不超過150%
        //104.02.13 add 逾放比< 1% 且 BIS > 10% 且備抵呆帳覆蓋率高100%且 > 全體信用部備抵呆帳覆蓋率平均值且備呆占放款比率> 2%,已申請經主管機關同意者,得不超過 200%)        
        tmp_A = 0.0;
        //97.01.29 add 990612非會員政策性農業專案貸款 990610 - 990511 - 990611 - 990612
        tmp_A =(double)amt990610-(double)amt990511-(double)amt990611-(double)amt990612;
        if(isDebug.equals("true")) System.out.println("(990610)-(990511)-(990611)-(990612)="+tmp_A);
        dataList = new ArrayList<String>();//傳內的參數List
        dataList.add(s_year);
        dataList.add(s_month);
        dataList.add((String)bean.getValue("bank_type"));
        dataList.add((String)bean.getValue("bank_no"));
        dataList.add("field_990610-990511-990611");
        dataList.add((String)bean.getValue("hsien_id"));
        dataList.add("0"); //type=4-->利率. type=2-->加總
        dataList.add(String.valueOf(tmp_A));
        dataList.add("");
        updateDBDataList.add(dataList);//1:傳內的參數List
        /*
        sqlCmd = "insert into a02_operation values("+s_year+","+s_month+",'"+(String)((DataObject)dbData.get(i)).getValue("bank_type")+"','"
               + (String)((DataObject)dbData.get(i)).getValue("bank_no")+"','field_990610-990511-990611','"
               + (String)((DataObject)dbData.get(i)).getValue("hsien_id")+"','0',"
               + tmp_A+",'')";
        //System.out.println("test9="+sqlCmd);
        updateDBSqlList.add(sqlCmd);
        */
        //field_k/990620=K/L
        tmp_A = 0.0;
        if(amt990620 != 0){
           //97.01.29 add 990612非會員政策性農業專案貸款 990610 - 990511 - 990611 - 990612
           tmp_A =((double)amt990610-(double)amt990511-(double)amt990611-(double)amt990612)/(double)amt990620;
           if(isDebug.equals("true")) System.out.println("(990610)-(990511)-(990611)-(990612)/990620="+tmp_A);
           tmp_B =Math.round( tmp_A * 10000);
           if(isDebug.equals("true")) System.out.println(((double)tmp_B)/100);
           tmp_A=((double)tmp_B)/100;
           if(isDebug.equals("true")) System.out.println("(990610)-(990511)-(990611)-(990612)/990620="+ tmp_A);
        }
        violate_remark="N";
        if(a01field_over_rate < 2 && a05bis > 8 ){
        	 if(a01field_over_rate < 1 && a05bis > 10){
           	  if((a01field_backup_over_rate == 0 && a01field_backup > 0 && a01field_over == 0 && amt990622 > 0)//備呆占狹義逾期放款比率(備抵呆帳/狹義逾放)=0,但備抵呆帳 > 0 且狹義逾放=0時,可>200
           	   ||(a01field_backup_over_rate > 100 && amt990622 > 0)){//104.02.13 add 990622           	  
                  if(isDebug.equals("true")) System.out.println("a01field_over_rate < 1 && a05bis > 10 && a01field_backup_over_rate > 100-->a01field_over_rate="+a01field_over_rate+":a05bis="+a05bis+":a01field_backup_over_rate="+a01field_backup_over_rate+":amt990622="+amt990622);          
                  if(tmp_A > 200){//逾放比< 1% 且 BIS > 10% 且備抵呆帳覆蓋率高100%且 > 全體信用部備抵呆帳覆蓋率平均值且備呆占放款比率> 2%,已申請經主管機關同意者,得不超過 200%
                     violate_remark="Y";
                  }
              }else if(amt990621 > 0){//逾放比< 1% 且 BIS > 10%,990621已申請經主管機關同意者,得不超過150%
                if(isDebug.equals("true")) System.out.println("a01field_over_rate < 1 && a05bis > 10 && a01field_backup_over_rate < 100-->a01field_over_rate="+a01field_over_rate+":a05bis="+a05bis+":a01field_backup_over_rate="+a01field_backup_over_rate+":amt990621="+amt990621);          
                if(tmp_A > 150){
                 violate_remark="Y";
                }              
              }else{
              	if(isDebug.equals("true")) System.out.println("a01field_over_rate < 1 && a05bis > 10 && a01field_backup_over_rate < 100-->a01field_over_rate="+a01field_over_rate+":a05bis="+a05bis+":a01field_backup_over_rate="+a01field_backup_over_rate+":amt990621="+amt990621+":amt990622="+amt990622);          
                if(tmp_A > 100){
                   violate_remark="Y";
                }
              }		              
           }else if(amt990621 > 0){//逾放比< 2% 且 BIS > 8%,990621已申請經主管機關同意者,得不超過150%
           	  if(isDebug.equals("true")) System.out.println("a01field_over_rate < 2 && a05bis > 8 && amt990621 > 0-->a01field_over_rate="+a01field_over_rate+":a05bis="+a05bis+":amt990621="+amt990621);          
              if(tmp_A > 150){
                 violate_remark="Y";
              }
           }else{//逾放比< 2% 且 BIS > 8%,未申請經主管機關同意者,100%
              if(isDebug.equals("true")) System.out.println("a01field_over_rate < 2 && a05bis > 8 && amt990621 = 0-->a01field_over_rate="+a01field_over_rate+":a05bis="+a05bis+":amt990621="+amt990621);          
              if(tmp_A > 100){
                 violate_remark="Y";
              }
           }   
        }else if(tmp_A > 100){
           violate_remark="Y";
        }
        dataList = new ArrayList<String>();//傳內的參數List
        dataList.add(s_year);
        dataList.add(s_month);
        dataList.add((String)bean.getValue("bank_type"));
        dataList.add((String)bean.getValue("bank_no"));
        dataList.add("field_k/990620");
        dataList.add((String)bean.getValue("hsien_id"));
        dataList.add("4"); //type=4-->利率. type=2-->加總
        dataList.add(String.valueOf(tmp_A));
        dataList.add(violate_remark);
        updateDBDataList.add(dataList);//1:傳內的參數List
        /*
        sqlCmd = "insert into a02_operation values("+s_year+","+s_month+",'"+(String)((DataObject)dbData.get(i)).getValue("bank_type")+"','"
               + (String)((DataObject)dbData.get(i)).getValue("bank_no")+"','field_k/990620','"
               + (String)((DataObject)dbData.get(i)).getValue("hsien_id")+"','4',"
               + tmp_A+",'"+violate_remark+"')";
        //System.out.println("test10="+sqlCmd);
        updateDBSqlList.add(sqlCmd);
        */
        //7.違反自用住宅放款總額占定期性存款總額比率不得超過40%之規定
        //field_990710/990720=M/N
        //104.02.13 調整為50% 
        tmp_A = 0.0;
        if(amt990720 != 0){
           tmp_A =(double)amt990710/(double)amt990720;
           if(isDebug.equals("true")) System.out.println("990710/990720="+tmp_A);
           tmp_B =Math.round( tmp_A * 10000);
           if(isDebug.equals("true")) System.out.println(((double)tmp_B)/100);
           tmp_A=((double)tmp_B)/100;
           if(isDebug.equals("true")) System.out.println("990710/990720="+ tmp_A);
        }
        violate_remark="N";
        if(tmp_A > 50){
           violate_remark="Y";
        }
        dataList = new ArrayList<String>();//傳內的參數List
        dataList.add(s_year);
        dataList.add(s_month);
        dataList.add((String)bean.getValue("bank_type"));
        dataList.add((String)bean.getValue("bank_no"));
        dataList.add("field_990710/990720");
        dataList.add((String)bean.getValue("hsien_id"));
        dataList.add("4"); //type=4-->利率. type=2-->加總
        dataList.add(String.valueOf(tmp_A));
        dataList.add(violate_remark);
        updateDBDataList.add(dataList);//1:傳內的參數List
        /*
        sqlCmd = "insert into a02_operation values("+s_year+","+s_month+",'"+(String)((DataObject)dbData.get(i)).getValue("bank_type")+"','"
               + (String)((DataObject)dbData.get(i)).getValue("bank_no")+"','field_990710/990720','"
               + (String)((DataObject)dbData.get(i)).getValue("hsien_id")+"','4',"
               + tmp_A+",'"+violate_remark+"')";
        //System.out.println("test11="+sqlCmd);
        updateDBSqlList.add(sqlCmd);
        */
        //8.違反信用部固定資產淨額不得超過上年度信用部決算淨值
        //103.01.06 add 上年度信用部決算淨值為負數時,要顯示違反
        //field_990810/(990230-990240)=P/D
        tmp_A = 0.0;
        //99.01.25 fix 99.01開始.上年度信用部決算淨值扣除992810投資全國農業金庫尚未攤提之損失
        if(((double)amt990230-(double)amt990240-(double)amt992810) != 0){
           tmp990230_990240_992810 = (double)amt990230-(double)amt990240-(double)amt992810;//103.01.06 add
           tmp_A =(double)amt990810/((double)amt990230-(double)amt990240-(double)amt992810);
           if(isDebug.equals("true")) System.out.println("990810/(990230-990240-992810)="+tmp_A);
           tmp_B =Math.round( tmp_A * 10000);
           if(isDebug.equals("true")) System.out.println(((double)tmp_B)/100);
           tmp_A=((double)tmp_B)/100;
           if(isDebug.equals("true")) System.out.println("990810/(990230-990240-992810)="+ tmp_A);
        }
        violate_remark="N";
        if(tmp_A > 100 && amt990812 == 0 && amt990813 == 0 & amt990814 == 0){
           violate_remark="Y";
        }
        if(tmp990230_990240_992810 < 0.0){//103.01.06 add 上年度信用部決算淨值為負數時,要顯示違反
    	   violate_remark="Y";
    	}	
        dataList = new ArrayList<String>();//傳內的參數List
        dataList.add(s_year);
        dataList.add(s_month);
        dataList.add((String)bean.getValue("bank_type"));
        dataList.add((String)bean.getValue("bank_no"));
        dataList.add("field_990810/(990230-990240)");
        dataList.add((String)bean.getValue("hsien_id"));
        dataList.add("4"); //type=4-->利率. type=2-->加總
        dataList.add(String.valueOf(tmp_A));
        dataList.add(violate_remark);
        updateDBDataList.add(dataList);//1:傳內的參數List
        //101.08.23 add
        //違反信用部固定資產淨額不得超過上年度信用部決算淨值不在此限的原因項目：
		//一、因購置或汰換安全維護或營業相關設備，經中央主管機關核准
		//二、因固定資產重估增值
		//三、因淨值降低
        //field_990812_990813_990814
        String field_990812_990813_990814 = "";
        if(amt990812 == 1) field_990812_990813_990814 += "1";
        if(amt990813 == 1) field_990812_990813_990814 += "2";
        if(amt990814 == 1) field_990812_990813_990814 += "3";
        if("".equals(field_990812_990813_990814)) field_990812_990813_990814 = "0";
        dataList = new ArrayList<String>();//傳內的參數List
        dataList.add(s_year);
        dataList.add(s_month);
        dataList.add((String)bean.getValue("bank_type"));
        dataList.add((String)bean.getValue("bank_no"));
        dataList.add("field_990812_990813_990814");
        dataList.add((String)bean.getValue("hsien_id"));
        dataList.add("2"); //type=4-->利率. type=2-->加總
        dataList.add(field_990812_990813_990814);
        dataList.add("");
        updateDBDataList.add(dataList);//1:傳內的參數List
        /*
        sqlCmd = "insert into a02_operation values("+s_year+","+s_month+",'"+(String)((DataObject)dbData.get(i)).getValue("bank_type")+"','"
               + (String)((DataObject)dbData.get(i)).getValue("bank_no")+"','field_990810/(990230-990240)','"
               + (String)((DataObject)dbData.get(i)).getValue("hsien_id")+"','4',"
               + tmp_A+",'"+violate_remark+"')";
        //System.out.println("test12="+sqlCmd);
        updateDBSqlList.add(sqlCmd);
        */
        //9.違反外幣資產與外幣負債差額絕對值不得超過上年底信用部帳面價值5%之限制
        //103.01.06 add 逾新台幣100萬且超過前一年度信用部決算淨值5%才算違反 
        //field_|990910-990920|/990230=|Q-R|/S
        tmp_A = 0.0;
        //99.01.25 fix 99.01開始.上年度信用部決算淨值扣除992810投資全國農業金庫尚未攤提之損失
        if((amt990230-amt992810) != 0){
           tmp_A=Math.abs((double)amt990910-(double)amt990920)/((double)amt990230-(double)amt992810);
           if(isDebug.equals("true")) System.out.println("|990910-990920|/(990230-992810)="+tmp_A);
           tmp_B =Math.round( tmp_A * 10000);
           if(isDebug.equals("true")) System.out.println(((double)tmp_B)/100);
           tmp_A=((double)tmp_B)/100;
           if(isDebug.equals("true")) System.out.println("|990910-990920|/(990230-992810)="+ tmp_A);
        }
        violate_remark="N";
        if(tmp_A > 5){          
           if(Math.abs((double)amt990910-(double)amt990920) > 1000000){//103.01.06 add 超過100萬者,才算違反
           	  violate_remark="Y";
           }
        }
        dataList = new ArrayList<String>();//傳內的參數List
        dataList.add(s_year);
        dataList.add(s_month);
        dataList.add((String)bean.getValue("bank_type"));
        dataList.add((String)bean.getValue("bank_no"));
        dataList.add("field_|990910-990920|/990230");
        dataList.add((String)bean.getValue("hsien_id"));
        dataList.add("4"); //type=4-->利率. type=2-->加總
        dataList.add(String.valueOf(tmp_A));
        dataList.add(violate_remark);
        updateDBDataList.add(dataList);//1:傳內的參數List
        /*
        sqlCmd = "insert into a02_operation values("+s_year+","+s_month+",'"+(String)((DataObject)dbData.get(i)).getValue("bank_type")+"','"
               + (String)((DataObject)dbData.get(i)).getValue("bank_no")+"','field_|990910-990920|/990230','"
               + (String)((DataObject)dbData.get(i)).getValue("hsien_id")+"','4',"
               + tmp_A+",'"+violate_remark+"')";
        //System.out.println("test13="+sqlCmd);
        updateDBSqlList.add(sqlCmd);
        */
        //10.違反擔保授信總額不得超過上一年度農會決算淨值1.5倍之限制
        //field_991020/990320=U/G
        tmp_A = 0.0;
        //99.01.25 fix 99.01開始.上年度全体農/漁會決算淨值扣除992810投資全國農業金庫尚未攤提之損失
        if((amt990320-amt992810) != 0){
           tmp_A=(double)amt991020/((double)amt990320-(double)amt992810);
           if(isDebug.equals("true")) System.out.println("991020/(990320-992810)="+tmp_A);
           tmp_value = dft.format(tmp_A);
           if(isDebug.equals("true")) System.out.println("tmp_value="+tmp_value);
           tmp_A = Double.parseDouble(tmp_value);
           //95.09.25 fix 倍改成%-->1.5倍=150%
           tmp_B =Math.round( tmp_A * 10000);
           if(isDebug.equals("true")) System.out.println(((double)tmp_B)/100);
           tmp_A=((double)tmp_B)/100;
           if(isDebug.equals("true")) System.out.println("field_991020/(990320-992810)="+ tmp_A);
        }
        violate_remark="N";
        //95.09.25 fix 倍改成%-->1.5倍=150%
        if(tmp_A > 150){
           violate_remark="Y";
        }
        dataList = new ArrayList<String>();//傳內的參數List
        dataList.add(s_year);
        dataList.add(s_month);
        dataList.add((String)bean.getValue("bank_type"));
        dataList.add((String)bean.getValue("bank_no"));
        dataList.add("field_991020/990320");
        dataList.add((String)bean.getValue("hsien_id"));
        dataList.add("4"); //type=4-->利率. type=2-->加總
        dataList.add(String.valueOf(tmp_A));
        dataList.add(violate_remark);
        updateDBDataList.add(dataList);//1:傳內的參數List
        /*
        sqlCmd = "insert into a02_operation values("+s_year+","+s_month+",'"+(String)((DataObject)dbData.get(i)).getValue("bank_type")+"','"
               + (String)((DataObject)dbData.get(i)).getValue("bank_no")+"','field_991020/990320','"
               + (String)((DataObject)dbData.get(i)).getValue("hsien_id")+"','4',"
               + tmp_A+",'"+violate_remark+"')";
        //System.out.println("test14="+sqlCmd);
        updateDBSqlList.add(sqlCmd);
        */
        //11.1違反放款最高限額不得超過上一年度農會信用部決算淨值25%之限制
        //field_X=公式_X
        //99.01.25 fix 99.01開始.上年度信用部決算淨值扣除992810投資全國農業金庫尚未攤提之損失
        double tmp990230_990240 = (double)amt990230-(double)amt990240-(double)amt992810;
        if((tmp990230_990240*0.25) > 9000000){
            field_X = tmp990230_990240 * 0.25;
        }else if((tmp990230_990240*0.25) <=9000000 && (tmp990230_990240*0.25) >6000000){
            field_X = 9000000;
        }else{
            field_X = 6000000;
        }
        if(isDebug.equals("true")) System.out.println("field_X="+field_X);
        tmp_value = dft_new.format(field_X);
        if(isDebug.equals("true")) System.out.println("tmp_value="+tmp_value);
        field_X = Double.parseDouble(tmp_value);
        dataList = new ArrayList<String>();//傳內的參數List
        dataList.add(s_year);
        dataList.add(s_month);
        dataList.add((String)bean.getValue("bank_type"));
        dataList.add((String)bean.getValue("bank_no"));
        dataList.add("field_X");
        dataList.add((String)bean.getValue("hsien_id"));
        dataList.add("0"); //type=4-->利率. type=2-->加總
        dataList.add(String.valueOf(tmp_value));
        dataList.add("");
        updateDBDataList.add(dataList);//1:傳內的參數List
        /*
        sqlCmd = "insert into a02_operation values("+s_year+","+s_month+",'"+(String)((DataObject)dbData.get(i)).getValue("bank_type")+"','"
               + (String)((DataObject)dbData.get(i)).getValue("bank_no")+"','field_X','"
               + (String)((DataObject)dbData.get(i)).getValue("hsien_id")+"','0',"
               + tmp_value+",'')";
        //System.out.println("test15="+sqlCmd);
        updateDBSqlList.add(sqlCmd);
        */
        //field_V_X=V > X
        if(amt991110 > field_X){
           field_V_X="Y";
        }
        if(isDebug.equals("true")) System.out.println("field_V_X="+field_V_X);
        dataList = new ArrayList<String>();//傳內的參數List
        dataList.add(s_year);
        dataList.add(s_month);
        dataList.add((String)bean.getValue("bank_type"));
        dataList.add((String)bean.getValue("bank_no"));
        dataList.add("field_V_X");
        dataList.add((String)bean.getValue("hsien_id"));
        dataList.add("0"); //type=4-->利率. type=2-->加總
        dataList.add("0");
        dataList.add(field_V_X);
        updateDBDataList.add(dataList);//1:傳內的參數List
        /*
        sqlCmd = "insert into a02_operation values("+s_year+","+s_month+",'"+(String)((DataObject)dbData.get(i)).getValue("bank_type")+"','"
               + (String)((DataObject)dbData.get(i)).getValue("bank_no")+"','field_V_X','"
               + (String)((DataObject)dbData.get(i)).getValue("hsien_id")+"','0',0"
               + ",'"+field_V_X+"')";
        //System.out.println("test16="+sqlCmd);
        updateDBSqlList.add(sqlCmd);
        */
        //11.2違反無擔保放款最高限額不得超過上一年度農會信用部決算淨值5%之限制
        //field_x=公式_x
        if((tmp990230_990240*0.05)>2000000){
           field_x = tmp990230_990240*0.05;
        }else{
           field_x = 2000000;
        }
        if(isDebug.equals("true")) System.out.println("field_x="+field_x);
        tmp_value = dft_new.format(field_x);
        if(isDebug.equals("true")) System.out.println("tmp_value="+tmp_value);
        field_x = Double.parseDouble(tmp_value);
        dataList = new ArrayList<String>();//傳內的參數List
        dataList.add(s_year);
        dataList.add(s_month);
        dataList.add((String)bean.getValue("bank_type"));
        dataList.add((String)bean.getValue("bank_no"));
        dataList.add("field_x");
        dataList.add((String)bean.getValue("hsien_id"));
        dataList.add("0"); //type=4-->利率. type=2-->加總
        dataList.add(String.valueOf(tmp_value));
        dataList.add("");
        updateDBDataList.add(dataList);//1:傳內的參數List
        /*
        sqlCmd = "insert into a02_operation values("+s_year+","+s_month+",'"+(String)((DataObject)dbData.get(i)).getValue("bank_type")+"','"
               + (String)((DataObject)dbData.get(i)).getValue("bank_no")+"','field_x','"
               + (String)((DataObject)dbData.get(i)).getValue("hsien_id")+"','0',"
               + tmp_value+",'')";
        //System.out.println("test17="+sqlCmd);
        updateDBSqlList.add(sqlCmd);
        */
        //field_W_x=W > x
        if(amt991120 > field_x){
           field_W_x="Y";
        }
        if(isDebug.equals("true")) System.out.println("field_W_x="+field_W_x);
        if(isDebug.equals("true")) System.out.println("field_V_X="+field_V_X);
        dataList = new ArrayList<String>();//傳內的參數List
        dataList.add(s_year);
        dataList.add(s_month);
        dataList.add((String)bean.getValue("bank_type"));
        dataList.add((String)bean.getValue("bank_no"));
        dataList.add("field_W_x");
        dataList.add((String)bean.getValue("hsien_id"));
        dataList.add("0"); //type=4-->利率. type=2-->加總
        dataList.add("0");
        dataList.add(field_W_x);
        updateDBDataList.add(dataList);//1:傳內的參數List
        /*
        sqlCmd = "insert into a02_operation values("+s_year+","+s_month+",'"+(String)((DataObject)dbData.get(i)).getValue("bank_type")+"','"
               + (String)((DataObject)dbData.get(i)).getValue("bank_no")+"','field_W_x','"
               + (String)((DataObject)dbData.get(i)).getValue("hsien_id")+"','0',0"
               + ",'"+field_W_x+"')";
        //System.out.println("test18="+sqlCmd);
        updateDBSqlList.add(sqlCmd);
        */
        //12.1違反放款最高限額不得超過上一年度農會信用部決算淨值25%之限制
        //field_a_X=a > X
        if(amt991210 > field_X ){
           field_a_X = "Y";
        }
        if(isDebug.equals("true")) System.out.println("field_a_X="+field_a_X);
        dataList = new ArrayList<String>();//傳內的參數List
        dataList.add(s_year);
        dataList.add(s_month);
        dataList.add((String)bean.getValue("bank_type"));
        dataList.add((String)bean.getValue("bank_no"));
        dataList.add("field_a_X");
        dataList.add((String)bean.getValue("hsien_id"));
        dataList.add("0"); //type=4-->利率. type=2-->加總
        dataList.add("0");
        dataList.add(field_a_X);
        updateDBDataList.add(dataList);//1:傳內的參數List
        /*
        sqlCmd = "insert into a02_operation values("+s_year+","+s_month+",'"+(String)((DataObject)dbData.get(i)).getValue("bank_type")+"','"
               + (String)((DataObject)dbData.get(i)).getValue("bank_no")+"','field_a_X','"
               + (String)((DataObject)dbData.get(i)).getValue("hsien_id")+"','0',0"
               + ",'"+field_a_X+"')";
        //System.out.println("test19="+sqlCmd);
        updateDBSqlList.add(sqlCmd);
        */
        //12.2違反無擔保放款最高限額不得超過上一年度農會信用部決算淨值5%之限制
        //field_b_x=b > x
        if(amt991220 > field_x){
           field_b_x="Y";
        }
        if(isDebug.equals("true")) System.out.println("field_b_x="+field_b_x);
        dataList = new ArrayList<String>();//傳內的參數List
        dataList.add(s_year);
        dataList.add(s_month);
        dataList.add((String)bean.getValue("bank_type"));
        dataList.add((String)bean.getValue("bank_no"));
        dataList.add("field_b_x");
        dataList.add((String)bean.getValue("hsien_id"));
        dataList.add("0"); //type=4-->利率. type=2-->加總
        dataList.add("0");
        dataList.add(field_b_x);
        updateDBDataList.add(dataList);//1:傳內的參數List
        /*
        sqlCmd = "insert into a02_operation values("+s_year+","+s_month+",'"+(String)((DataObject)dbData.get(i)).getValue("bank_type")+"','"
               + (String)((DataObject)dbData.get(i)).getValue("bank_no")+"','field_b_x','"
               + (String)((DataObject)dbData.get(i)).getValue("hsien_id")+"','0',0"
               + ",'"+field_b_x+"')";
        //System.out.println("test19_1="+sqlCmd);
        updateDBSqlList.add(sqlCmd);
        */
        //13.1違反放款最高限額不得超過上一年度農會信用部決算淨值12.5%之限制
        //field_d=公式_d
        if((tmp990230_990240*0.125)>9000000){
           field_d = tmp990230_990240*0.125;
        }else if((tmp990230_990240*0.125)<=9000000 &&(tmp990230_990240*0.125)>6000000){
           field_d = 9000000;
        }else{
           field_d = 6000000;
        }
        if(isDebug.equals("true")) System.out.println("field_d="+field_d);
        field_d = Math.round(field_d);
        if(isDebug.equals("true")) System.out.println("field_d.round="+field_d);
        /*
        tmp_value = dft_new.format(field_d);
        if(isDebug.equals("true")) System.out.println("tmp_value="+tmp_value);
        field_d = Double.parseDouble(tmp_value);
        */
        dataList = new ArrayList<String>();//傳內的參數List
        dataList.add(s_year);
        dataList.add(s_month);
        dataList.add((String)bean.getValue("bank_type"));
        dataList.add((String)bean.getValue("bank_no"));
        dataList.add("field_d");
        dataList.add((String)bean.getValue("hsien_id"));
        dataList.add("0"); //type=4-->利率. type=2-->加總
        dataList.add(String.valueOf(field_d));
        dataList.add("");
        updateDBDataList.add(dataList);//1:傳內的參數List
        /*
        sqlCmd = "insert into a02_operation values("+s_year+","+s_month+",'"+(String)((DataObject)dbData.get(i)).getValue("bank_type")+"','"
               + (String)((DataObject)dbData.get(i)).getValue("bank_no")+"','field_d','"
               + (String)((DataObject)dbData.get(i)).getValue("hsien_id")+"','0',"
               + field_d+",'')";
        //System.out.println("test20="+sqlCmd);
        updateDBSqlList.add(sqlCmd);
        */
        //field_c_d=c > d
        if(isDebug.equals("true")) System.out.println("amt991310="+amt991310);
        if(amt991310 > field_d){
           field_c_d ="Y";
        }
        if(isDebug.equals("true")) System.out.println("field_c_d="+field_c_d);
        dataList = new ArrayList<String>();//傳內的參數List
        dataList.add(s_year);
        dataList.add(s_month);
        dataList.add((String)bean.getValue("bank_type"));
        dataList.add((String)bean.getValue("bank_no"));
        dataList.add("field_c_d");
        dataList.add((String)bean.getValue("hsien_id"));
        dataList.add("0"); //type=4-->利率. type=2-->加總
        dataList.add("0");
        dataList.add(field_c_d);
        updateDBDataList.add(dataList);//1:傳內的參數List
        /*
        sqlCmd = "insert into a02_operation values("+s_year+","+s_month+",'"+(String)((DataObject)dbData.get(i)).getValue("bank_type")+"','"
               + (String)((DataObject)dbData.get(i)).getValue("bank_no")+"','field_c_d','"
               + (String)((DataObject)dbData.get(i)).getValue("hsien_id")+"','0',0"
               + ",'"+field_c_d+"')";
        //System.out.println("test21="+sqlCmd);
        updateDBSqlList.add(sqlCmd);
        */
        //13.2違反無擔保放款最高限額不得超過上一年度農會信用部決算淨值2.5%之限制
        //field_g=公式_g
        if((tmp990230_990240*0.025)>2000000){
           field_g = tmp990230_990240*0.025;
        }else{
           field_g = 2000000;
        }
        if(isDebug.equals("true")) System.out.println("field_g="+field_g);
        tmp_value = dft_new.format(field_g);
        if(isDebug.equals("true")) System.out.println("tmp_value="+tmp_value);
        field_g = Double.parseDouble(tmp_value);
        dataList = new ArrayList<String>();//傳內的參數List
        dataList.add(s_year);
        dataList.add(s_month);
        dataList.add((String)bean.getValue("bank_type"));
        dataList.add((String)bean.getValue("bank_no"));
        dataList.add("field_g");
        dataList.add((String)bean.getValue("hsien_id"));
        dataList.add("0"); //type=4-->利率. type=2-->加總
        dataList.add(String.valueOf(tmp_value));
        dataList.add("");
        updateDBDataList.add(dataList);//1:傳內的參數List
        /*
        sqlCmd = "insert into a02_operation values("+s_year+","+s_month+",'"+(String)((DataObject)dbData.get(i)).getValue("bank_type")+"','"
               + (String)((DataObject)dbData.get(i)).getValue("bank_no")+"','field_g','"
               + (String)((DataObject)dbData.get(i)).getValue("hsien_id")+"','0',"
               + tmp_value+",'')";
        //System.out.println("test22="+sqlCmd);
        updateDBSqlList.add(sqlCmd);
        */
        //field_f_g=f  >  g
        if(amt991320 > field_g){
           field_f_g = "Y";
        }
        if(isDebug.equals("true")) System.out.println("field_f_g="+field_f_g);
        dataList = new ArrayList<String>();//傳內的參數List
        dataList.add(s_year);
        dataList.add(s_month);
        dataList.add((String)bean.getValue("bank_type"));
        dataList.add((String)bean.getValue("bank_no"));
        dataList.add("field_f_g");
        dataList.add((String)bean.getValue("hsien_id"));
        dataList.add("0"); //type=4-->利率. type=2-->加總
        dataList.add("0");
        dataList.add(field_f_g);
        updateDBDataList.add(dataList);//1:傳內的參數List
        /*
        sqlCmd = "insert into a02_operation values("+s_year+","+s_month+",'"+(String)((DataObject)dbData.get(i)).getValue("bank_type")+"','"
               + (String)((DataObject)dbData.get(i)).getValue("bank_no")+"','field_f_g','"
               + (String)((DataObject)dbData.get(i)).getValue("hsien_id")+"','0',0"
               + ",'"+field_f_g+"')";
        System.out.println("test23="+sqlCmd);
        updateDBSqlList.add(sqlCmd);
		*/
		
		//14.對鄉(鎮、市)公所授信未經其所隸屬之縣政府保證之限額-違反對鄉(鎮、市)公所授信未經其所隸屬之縣政府保證,及對直轄市、縣(市)政府投資經營之公營事業,其授信經該直轄市、縣(市)政府保證,兩者合計不得超過上一年度信用部決算淨值       
        //field_996114_996115/(990230-990240)=(i+j)/D
        tmp_A = 0.0;
        //99.01.25 fix 99.01開始.上年度信用部決算淨值扣除992810投資全國農業金庫尚未攤提之損失
        if(((double)amt990230-(double)amt990240-(double)amt992810) != 0 ){
           tmp_A=((double)amt996114+(double)amt996115)/((double)amt990230-(double)amt990240-(double)amt992810);
           if(isDebug.equals("true")) System.out.println("field_996114_996115/(990230-990240-992810)="+tmp_A);
           tmp_B =Math.round( tmp_A * 10000);
           if(isDebug.equals("true")) System.out.println(((double)tmp_B)/100);
           tmp_A=((double)tmp_B)/100;
           if(isDebug.equals("true")) System.out.println("field_996114_996115/(990230-990240-992810)="+ tmp_A);
        }
        
        violate_remark="N";
        if(tmp_A > 100){
           violate_remark="Y";
        }
        dataList = new ArrayList<String>();//傳內的參數List
        dataList.add(s_year);
        dataList.add(s_month);
        dataList.add((String)bean.getValue("bank_type"));
        dataList.add((String)bean.getValue("bank_no"));
        dataList.add("field_996114_996115/(990230-990240)");
        dataList.add((String)bean.getValue("hsien_id"));
        dataList.add("4"); //type=4-->利率. type=2-->加總
        dataList.add(String.valueOf(tmp_A));
        dataList.add(violate_remark);
        updateDBDataList.add(dataList);//1:傳內的參數List
		
    }//end of for


	//1.寫入A02_OPERATION==================================================================================================
	sqlCmd.delete(0, sqlCmd.length());
	sqlCmd.append("insert into a02_operation values(?,?,?,?,?,?,?,?,?)");
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
	dataList.add("bank_no"); 
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
       logps.println(logformat.format(nowlog)+" "+"產生 A02_opeation 完成");
    }else{
       logps.println(logformat.format(nowlog)+" "+"執行 A02_opeation 失敗:"+errMsg);
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
