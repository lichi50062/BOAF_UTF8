<%
// 95.10.20 create by 2295
// 95.12.05 add 增加年月區間,可根據年月做sort,拿掉可選欄位做sort by 2295
// 99.05.14 fix 縣市合併 & sql injection
//100.05.12 fix 有挑選排序欄位年月時,查詢SQL error by 2295 
//103.02.10 add 103/01以後,漁會套用新表格(增加/異動科目代號) by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.util.*,java.io.*" %>
<%@ page import="org.apache.poi.poifs.filesystem.*,org.apache.poi.hssf.usermodel.*" %>
<%@ page import="org.apache.poi.hssf.util.Region" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.report.Report01" %>
<%@ page import="com.tradevan.util.report.HssfStyle" %>
<%@ page import="com.tradevan.util.report.reportUtil" %>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="java.util.List" %>
<%@ page import="java.text.SimpleDateFormat" %>
<%@ page import="java.util.StringTokenizer" %>
<%@ page import="java.lang.StringBuffer" %>
<%@ page import="java.lang.Short" %>
<%@ page import="java.lang.Math" %>
<%@ page import="java.text.DecimalFormat" %>

<%
   response.setContentType("application/msexcel;charset=UTF-8");//以上這行設定本網頁為excel格式的網頁
   String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");
   System.out.println("act="+act);
   if(act.equals("view")){
      //以上這行設定傳送到前端瀏覽器時的檔名為test1.xls
      //就是靠這一行，讓前端瀏覽器以為接收到一個excel檔
      response.setHeader("Content-disposition","inline; filename=view.xls");
   }else if (act.equals("download")){
      response.setHeader("Content-Disposition","attachment; filename=download.xls");
   }
%>
<%
	DecimalFormat dft = new DecimalFormat("#.##");
	String actMsg = "";
	FileOutputStream fileOut = null;
    HSSFCellStyle defaultStyle;
    HSSFCellStyle rightStyle;
    HSSFCellStyle noBorderDefaultStyle;
    HSSFCellStyle noBorderLeftStyle;
    HSSFCellStyle titleStyle;
	HSSFCellStyle columnStyle;
	HSSFCellStyle noBoderStyle;
	HSSFRow row;
	HSSFRow acc_code_row;//讀取acc_code的row
	HSSFCell cell = null;//宣告一個儲存格
	String titleName = "信用部";
    reportUtil reportUtil = new reportUtil();
    String BankList = "";//儲存bank_code/bank_name
    String btnFieldList = "";//儲存所選取的大類acc_code/名稱
    String SortList = "";//排序的acc_code
    String CANCEL_NO = "";//裁撤別
    String Unit = "";//列印單位
    String S_YEAR = "";//年
    String E_YEAR = "";//年
    String S_MONTH = "";//月
    String E_MONTH = "";//月
    List BankList_data = null;//儲存bank_code/bank_name的集合
    List btnFieldList_data = null;
    List SortList_data = null;
	int i = 0;
	int j= 0;
	String lguser_name = "測試使用者";
	String bank_type="";
	String hasBankListALL="false";
	String u_year = "99" ;
	String cd01Table = "cd01_99" ;
	DataObject bean = null ;
	String stryymm="";
	String ncacno="ncacno";
	try{
			bank_type = ((String)session.getAttribute("nowbank_type")).equals("")?"6":(String)session.getAttribute("nowbank_type");
			System.out.println("bank_type="+bank_type);
			//儲存報表的目錄================================================================
        	File reportDir = new File(Utility.getProperties("reportDir"));
    		if(!reportDir.exists()){
     			if(!Utility.mkdirs(Utility.getProperties("reportDir"))){
     	   			actMsg +=Utility.getProperties("reportDir")+"目錄新增失敗";
     			}
    		}
    		//==============================================================================
    		//營運中/已裁撤
			if(session.getAttribute("CANCEL_NO") != null && !((String)session.getAttribute("CANCEL_NO")).equals("")){
		  		CANCEL_NO = (String)session.getAttribute("CANCEL_NO");
			}
    		//金融機構
			if(session.getAttribute("BankList") != null && !((String)session.getAttribute("BankList")).equals("")){
		   		BankList = (String)session.getAttribute("BankList");
		   		BankList_data = Utility.getReportData(BankList);
		   		System.out.println("BankList_data.size()="+BankList_data.size());
		   		System.out.println("BankList_data="+BankList_data);
			}
			//報表欄位
			if(session.getAttribute("btnFieldList") != null && !((String)session.getAttribute("btnFieldList")).equals("")){
		   		btnFieldList = (String)session.getAttribute("btnFieldList");
		   		btnFieldList_data = Utility.getReportData(btnFieldList);
		   		System.out.println("btnFieldList_data.size()="+btnFieldList_data.size());
		   		System.out.println("btnFieldList_data="+btnFieldList_data);
			}
			//排序欄位
			if(session.getAttribute("SortList") != null && !((String)session.getAttribute("SortList")).equals("")){
		  		SortList = (String)session.getAttribute("SortList");
		  		SortList_data = Utility.getReportData(SortList);
		   		System.out.println("SortList_data.size()="+SortList_data.size());
		   		System.out.println("SortList_data="+SortList_data);
			}

        	//機構類別
			if(session.getAttribute("nowbank_type") != null && !((String)session.getAttribute("nowbank_type")).equals("")){
			    if(((String)session.getAttribute("nowbank_type")).equals("6")){
			       titleName = "農會" + titleName;
			    }else if(((String)session.getAttribute("nowbank_type")).equals("7")){
			       titleName = "漁會" + titleName;
			    }else{
			       titleName = "農漁會" + titleName;
			    }
			}

			titleName += "逾期放款資料";

			//年
			if(session.getAttribute("S_YEAR") != null && !((String)session.getAttribute("S_YEAR")).equals("")){
		  		S_YEAR = (String)session.getAttribute("S_YEAR");
		  		if(Integer.parseInt(S_YEAR) > 99 ) {
		  			u_year = "100" ;
		  			cd01Table = "cd01" ;
		  		}
			}
			
			//年
			if(session.getAttribute("E_YEAR") != null && !((String)session.getAttribute("E_YEAR")).equals("")){
		  		E_YEAR = (String)session.getAttribute("E_YEAR");
			}
			//95.12.05 增加年月區間
			//月
			if(session.getAttribute("S_MONTH") != null && !((String)session.getAttribute("S_MONTH")).equals("")){
		  		S_MONTH = (String)session.getAttribute("S_MONTH");		  		
			}
			//月
			if(session.getAttribute("E_MONTH") != null && !((String)session.getAttribute("E_MONTH")).equals("")){
		  		E_MONTH = (String)session.getAttribute("E_MONTH");
			}
			//金額單位
			if(session.getAttribute("Unit") != null && !((String)session.getAttribute("Unit")).equals("")){
		  		Unit = (String)session.getAttribute("Unit");
			}

		    //103.02.10 add 103/01以後,漁會套用新表格(增加/異動科目代號) 
    		if( bank_type.equals("7") && (Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH) >= 10301) ){
    			stryymm = "_10301";
    		}
    		ncacno = bank_type.equals("6")?"ncacno":"ncacno_7";
		    if(bank_type.equals("7")){   
	           //103.02.10 add 103/01以後,漁會套用新表格(增加/異動科目代號) 
	           if(Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH) >= 10301){
                  ncacno = "ncacno_7_rule";
               }
            }  		
			//讀取欄位大類所包含的細項===================================================================================
        	Properties prop_column = new Properties();
			prop_column.load(new FileInputStream(Utility.getProperties("schemaDir")+System.getProperty("file.separator")+"A06_"+bank_type+"_detail"+stryymm+".TXT"));
			//=======================================================================================================================
			//取出欄位細項將資料存入MAP-->key=大類acc_code,value=細項acc_code=============================================================
			HashMap h_column = new HashMap();//儲存column大類,及其細項的acc_code
			List detail_column = new LinkedList();
			String column_tmp = "";
			String selectacc_code = "";//選取的detail科目代號
			String a06_operation_field_sum="";
			String a06_operation_field="";
			String a06_field_sum="";
			String ori_field="";
			String a06acc_code="";
			String a06_field_sum_hsien_id = "";//縣市別用
			String a06_field_hsien_id = "";//縣市別用
			int columnLength=0;//column個數
			for(i=0;i<btnFieldList_data.size();i++){
			    column_tmp = "";
			    column_tmp = (String)prop_column.get((String)((List)btnFieldList_data.get(i)).get(0));
			    //System.out.println("column_tmp="+column_tmp);
			    if(!column_tmp.equals("")){
			        detail_column = Utility.getStringTokenizerData(column_tmp,"+");
			        //System.out.println(detail_column);
			        if(detail_column != null && detail_column.size() != 0){
			           columnLength += detail_column.size();//累加總欄位個數
              		   for(j=0;j<detail_column.size();j++){
              		   	   a06acc_code = (String)detail_column.get(j);
            	 		   selectacc_code +="'"+a06acc_code+"'";
            	 		   //95.09.26 add=================================================================================================================
            	 		   if(a06acc_code.length() <= 6){
            	 		      ori_field+= a06acc_code;
            	 		      a06_field_sum += " ,sum(decode(acc_code,'"+a06acc_code+"',amt,0)) as amt"+a06acc_code;
            	 		   }else{
            	 		      ori_field += a06acc_code;
            	 		      //sum(decode(acc_code,'120101',amt_3month,0)) as amt_3month_120101
            	 		      if(a06_field_sum.indexOf(" ,sum(decode(acc_code,'"+a06acc_code.substring(a06acc_code.length()-6,a06acc_code.length())+"',"+a06acc_code.substring(0,a06acc_code.length()-7)+",0)) as "+a06acc_code) == -1){
            	 		         a06_field_sum += " ,sum(decode(acc_code,'"+a06acc_code.substring(a06acc_code.length()-6,a06acc_code.length())+"',"+a06acc_code.substring(0,a06acc_code.length()-7)+",0)) as "+a06acc_code;
            	 		      }

            	 		      //sum(d.amt_3month_120101) as amt_3month_120101,
            	 		      if(a06_field_sum_hsien_id.indexOf(" ,sum(d."+a06acc_code+") as "+a06acc_code) == -1){
            	 		         a06_field_sum_hsien_id += " ,sum(d."+a06acc_code+") as "+a06acc_code;
            	 		      }
            	 		      //a06.amt_3month_120101,
            	 		      if(a06_field_hsien_id.indexOf(" ,a06."+a06acc_code) == -1){
            	 		         a06_field_hsien_id += " ,a06."+a06acc_code;
            	 		      }
            	 		   }
            	 		   //====================================================================================================================================
            	 		   if(j < detail_column.size()-1){
            	 		      selectacc_code +=",";
            	 		      ori_field +=",";
            	 		   }
               		   }
               		   //System.out.println("select acc_code="+selectacc_code);
            	   }else{

            	   }
			       h_column.put((String)((List)btnFieldList_data.get(i)).get(0),detail_column);
			       if(i < btnFieldList_data.size()-1){
			          selectacc_code +=",";
			          ori_field +=",";
			       }
			    }
			}
			System.out.println("select acc_code="+selectacc_code);
			System.out.println("h_column.size()="+h_column.size());
        	//讀取報表欄位名稱===================================================================================
        	Properties prop_column_name = new Properties();
			prop_column_name.load(new FileInputStream(Utility.getProperties("schemaDir")+System.getProperty("file.separator")+"A06_"+bank_type+"_column"+stryymm+".TXT"));
			//====================================================================================================
			String selectBank_no = "";//選取的金融機構代號
            //金融機構代號=============================================================
            if(BankList_data != null && BankList_data.size() != 0){
               for(i=0;i<BankList_data.size();i++){
                   //95.09.04 判斷機構代號是否為ALL:全部===============================================
			       if(((String)((List)BankList_data.get(i)).get(0)).equals("ALL")){
			          hasBankListALL="true";

			       }
            	   selectBank_no +="'"+(String)((List)BankList_data.get(i)).get(0)+"'";
            	   if(i < BankList_data.size()-1) selectBank_no +=",";
               }
               System.out.println("select bank_no="+selectBank_no);
            }
            //==============================================================================
            String order = "";//排序欄位
            //排序欄位=========================================================================
            if(SortList_data != null && SortList_data.size() != 0){
            	for(i=0;i<SortList_data.size();i++){
            		//order += (String)((List)SortList_data.get(i)).get(0);
            		//95.09.26 add====================================================
            		if(hasBankListALL.equals("true")){//選全部時
            		   order += (String)((List)SortList_data.get(i)).get(0);
            		}else{
            		   if(((String)((List)SortList_data.get(i)).get(0)).length() <= 6){
            		      order += "amt"+(String)((List)SortList_data.get(i)).get(0);
            		   }else{
            		      order += (String)((List)SortList_data.get(i)).get(0);
            		   }
            		}
            		//===============================================================
            		if(i < SortList_data.size() -1 ) order +=",";
	            }
	            System.out.println("order="+order);
            }
            //====================================================================================
  			/*
  			//各別機構,按bank_code,m_year,m_month排序
            select a06.bank_code,bn01.bank_name,m_year || '/' || m_month as m_yearmonth,
                   wlx01.hsien_id,cd01.hsien_name,a06.amt_3month_150200
            from (
            	 	select m_year,m_month,bank_code,
             		       sum(decode(acc_code,'150200',amt_3month,0)) as amt_3month_150200
            	    from (
            		 	 select A06.m_year, A06.m_month,a06.bank_code,
            	  		  	     ncacno.acc_range,A06.acc_code,ncacno.acc_name,
            	  		  	     round(amt_3month/1,0) as amt_3month, --未滿3個月
            	  		  	     round(amt_6month/1,0) as amt_6month, --3個月~6個月
            	  		  	     round(amt_1year/1,0) as amt_1year, --6個月~1年
            	  		  	     round(amt_2year/1,0) as amt_2year, --1年~2年
            	  		  	     round(amt_over2year/1,0) as amt_over2year, --2年以上
            	  		  	     round(amt_total/1,0) as amt_total -- 逾放合計
            	  		  	     ,round(A01.amt/1,0) as amt -- a01放款合計
                          from A06 LEFT JOIN ncacno ON A06.acc_code = ncacno.acc_code and ncacno.acc_div='08'
                          	       LEFT JOIN  (select m_year,m_month,bank_code,
            					               decode(acc_code,'990000','970000',acc_code) as acc_code
                                               ,amt from a01
            								   )A01 ON A01.bank_code = A06.bank_code
                          		   			        and A01.m_year = A06.m_year and A01.m_month = A06.m_month
            									    and A01.acc_code = A06.acc_code
                          where (to_char(m_year * 100 + m_month) >= 9406 and to_char(m_year * 100 + m_month) <= 9508)
                          and A06.bank_code in ('6030016','6040017')
                          order by ncacno.acc_range
            			 )a06
              		where (to_char(m_year * 100 + m_month) >= 9406 and to_char(m_year * 100 + m_month) <= 9508)
              		and bank_code in  ('6030016','6180302','6180313')
            		group by m_year,m_month,bank_code
            	 )a06,bn01 left join wlx01 on bn01.bank_no = wlx01.bank_no
            			   left join cd01 on wlx01.HSIEN_ID = cd01.HSIEN_ID
            where a06.bank_code = bn01.bank_no
            and bn01.bank_no in ('6030016','6180302','6180313') and bn01.bn_type <> '2'
            and a06.bank_code = bn01.bank_no
            group by a06.bank_code,bn01.bank_name,a06.m_year,a06.m_month,wlx01.hsien_id,cd01.hsien_name,a06.amt_3month_150200
            order by a06.bank_code,a06.m_year,a06.m_month
            漁會把a01取代成
            (select m_year,m_month,bank_code,acc_code,amt
            from a01
            where acc_code not in ('120700','120900') and m_year=94 and m_month=6
            union
            select m_year,m_month,bank_code,'120700' as acc_code,sum(amt) amt from a01
            where acc_code in('120700','120900') and m_year=94 and m_month=6
            group by m_year,m_month,bank_code
            order by m_year,m_month,bank_code ) a01

            //縣市別.按縣市.年月sort
            select e.FR001W_OUTPUT_ORDER,nvl(e.hsien_id,' ') as hsien_id,
            	   nvl(e.hsien_name,'其他') as hsien_name,
            	   m_year || '/' || m_month as m_yearmonth,
             	   sum(d.amt_3month_120101) as amt_3month_120101,
            	   sum(d.amt_120101) as amt_120101
            from
            	 (select wlx01.hsien_id,cd01.hsien_name,cd01.FR001W_OUTPUT_ORDER,a06.bank_code,bn01.bank_name,a06.m_year,a06.m_month,
            			 a06.amt_3month_120101,a06.amt_120101
              	  from (
            	 	 	 select m_year,m_month,bank_code,
            			   	    sum(decode(acc_code,'120101',amt_3month,0)) as amt_3month_120101,
            				 	sum(decode(acc_code,'120101',amt,0)) as amt_120101
            			 from (
            				   select m_year,m_month,bank_code,acc_code,amt_3month,amt_6month,amt_1year,amt_2year,amt_over2year,amt_total,amt
            				   from (select A06.m_year, A06.m_month,a06.bank_code,
            				   		 	    ncacno.acc_range,A06.acc_code,ncacno.acc_name,
            				  		  	    round(amt_3month/1,0) as amt_3month, --未滿3個月
            				  		  	    round(amt_6month/1,0) as amt_6month, --3個月~6個月
            				  		  	    round(amt_1year/1,0) as amt_1year, --6個月~1年
              	  	  		  		  	    round(amt_2year/1,0) as amt_2year, --1年~2年
              	  	  		  		  	    round(amt_over2year/1,0) as amt_over2year, --2年以上
                        	  		  	    round(amt_total/1,0) as amt_total, -- 逾放合計
            								0 as amt
                                     from A06 LEFT JOIN ncacno ON A06.acc_code = ncacno.acc_code and ncacno.acc_div='08',bn01
            						 where (to_char(m_year * 100 + m_month) >= 9406 and to_char(m_year * 100 + m_month) <= 9408)
            						 and a06.bank_code = bn01.bank_no
            						 and bn01.bank_type='6'
            						 and bn01.bn_type <> '2'
            						 and a06.bank_code = bn01.bank_no
            						 union all
            						 select A01.m_year, A01.m_month,a01.bank_code,
            				  		    	ncacno.acc_range,A01.acc_code,ncacno.acc_name,
            							 	0 as amt_3month, --未滿3個月
            				  		  	    0 as amt_6month, --3個月~6個月
            				  		  	   	0 as amt_1year, --6個月~1年
              	  	  		  		  	   	0 as amt_2year, --1年~2年
              	  	  		  		  	   	0 as amt_over2year, --2年以上
                        	  		  	   	0 as amt_total, -- 逾放合計
                        	  		  	   	round(A01.amt/1,0) as amt -- a01放款合計
                                     from (select m_year,m_month,bank_code,
            							   		  decode(acc_code,'990000','970000',acc_code) as acc_code,amt from a01
                        				   )A01 LEFT JOIN ncacno ON A01.acc_code = ncacno.acc_code and ncacno.acc_div='08',bn01
            						 where (to_char(m_year * 100 + m_month) >= 9406 and to_char(m_year * 100 + m_month) <= 9408)
            						 and a01.bank_code = bn01.bank_no
            						 and bn01.bn_type <> '2'
            						 and bn01.bank_type='6'
            						 and a01.bank_code = bn01.bank_no
            						 )c
            					 )a06
            					group by m_year,m_month,bank_code
            				 )a06,bn01 left join wlx01 on bn01.bank_no = wlx01.bank_no
            					   	   left join cd01 on wlx01.HSIEN_ID = cd01.HSIEN_ID
            				where a06.bank_code = bn01.bank_no
            				and bn01.bn_type <> '2'
            				and bn01.bank_type in ('6')
            				and a06.bank_code = bn01.bank_no
            	) d ,v_bank_location e
            where d.bank_code(+) =e.bank_no
            and (e.hsien_id>'Y' or e.hsien_id<'Y')
            and m_year is not null and m_month is not null
            group by e.FR001W_OUTPUT_ORDER,nvl(e.hsien_id,' '),nvl(e.hsien_name,'其他'),m_year,m_month
            order by e.fr001w_output_order,m_year,m_month
  			*/

            String column = "";//選取欄位
            //String condition = "";//其他條件
            StringBuffer condition = new StringBuffer() ;
            List conditionList = new ArrayList() ;
            //add 營運中/已裁撤============================
			condition.append(" and bn01.bn_type "+(CANCEL_NO.equals("N")?"<>":"=")+" ? and a06.bank_code = bn01.bank_no ");
            conditionList.add("2") ;
			//======================================================
			StringBuffer sqlCmd = new StringBuffer() ;
			List sqlCmdList = new ArrayList() ;
            StringBuffer sqlCmd_sum = new StringBuffer() ;//縣市別小計
            List sqlCmd_sumList = new ArrayList();
            StringBuffer a06_table = new StringBuffer() ;
            List a06_tableList = new ArrayList() ;
			StringBuffer a01_table = new StringBuffer() ;
			List a01_tableList = new ArrayList() ;
            //漁會的a01取代成
            if(bank_type.equals("7")){
              if(Integer.parseInt(S_YEAR) * 100 + Integer.parseInt(S_MONTH) >= 10301){
              	a01_table.append(" a01 ") ;
              }else{              	 		
                 a01_table.append(" (select m_year,m_month,bank_code,acc_code,amt "
				           + "	  from a01 "
				     		+ "   where acc_code not in (?,?)"
				     		+ "   and   (to_char(m_year * 100 + m_month) >= ? and to_char(m_year * 100 + m_month) <= ? )"
				           + "   union"
				           + "   select m_year,m_month,bank_code,'120700' as acc_code,sum(amt) amt from a01 "
				           + "   where acc_code in(?,?) "
				           + "   and   (to_char(m_year * 100 + m_month) >= ? and to_char(m_year * 100 + m_month) <= ? )"
				           + "   group by m_year,m_month,bank_code "
				           + "   order by m_year,m_month,bank_code ) a01 ");
                a01_tableList.add("120700") ;
                a01_tableList.add("120900") ;
                a01_tableList.add(S_YEAR+S_MONTH) ;
                a01_tableList.add(E_YEAR+E_MONTH) ;
                a01_tableList.add("120700") ;
                a01_tableList.add("120900") ;
                a01_tableList.add(S_YEAR+S_MONTH) ;
                a01_tableList.add(E_YEAR+E_MONTH) ;
              }
			}else {
				a01_table.append(" a01 ") ;
			}
            
           	a06_table.append(" (	"
			 		  + "  select m_year,m_month,bank_code"+ a06_field_sum
				      + "  from ( "
					  + "		  select A06.m_year, A06.m_month,a06.bank_code,"
				  	  + "		  	     "+ncacno+ ".acc_range,A06.acc_code,"+ncacno+ ".acc_name,"
				  	  + "		  	     round(amt_3month/1,0) as amt_3month,"//--未滿3個月
				  	  + "		  	     round(amt_6month/1,0) as amt_6month,"//--3個月~6個月
				  	  + "		  	     round(amt_1year/1,0) as amt_1year,"//--6個月~1年
  	  	  		  	  + "		  	     round(amt_2year/1,0) as amt_2year,"//--1年~2年
  	  	  		  	  + "		  	     round(amt_over2year/1,0) as amt_over2year,"//--2年以上
            	  	  + "		  	     round(amt_total/1,0) as amt_total"//-- 逾放合計
            	  	  + "		  	     ,round(A01.amt/1,0) as amt"//-- a01放款合計
                      + "		  from A06 LEFT JOIN "+ncacno+ " ON A06.acc_code = "+ncacno+ ".acc_code and "+ncacno+ ".acc_div= ? "
			          + "	    	       LEFT JOIN  (select m_year,m_month,bank_code,"
					  + "			   		           decode(acc_code,'990000','970000',acc_code) as acc_code,amt from "+a01_table.toString()
            		  + "							   )A01 ON A01.bank_code = A06.bank_code"
                      + "    					        and A01.m_year = A06.m_year and A01.m_month = A06.m_month"
            		  + "							    and A01.acc_code = A06.acc_code"
            		  + "               where (to_char(A06.m_year * 100 + A06.m_month) >= ? and to_char(A06.m_year * 100 + A06.m_month) <= ? )");
           	a06_tableList.add("08");
           	for(Object obj : a01_tableList) {
           		a06_tableList.add(obj) ;
           	}
           	a06_tableList.add(S_YEAR+S_MONTH);
           	a06_tableList.add(E_YEAR+E_MONTH);
			if(hasBankListALL.equals("false")){
			   a06_table.append(" 	 and   A06.bank_code in ("+selectBank_no+")");
			}
            a06_table.append("	     order by "+ncacno+ ".acc_range "
					  + "	   )a06 "
					  + "  where (to_char(m_year * 100 + m_month) >= ? and to_char(m_year * 100 + m_month) <= ? )");
            a06_tableList.add(S_YEAR+S_MONTH);
           	a06_tableList.add(E_YEAR+E_MONTH);
			if(hasBankListALL.equals("false")){
			   a06_table.append("  and   bank_code in ("+selectBank_no+")");
			}
			a06_table.append("  group by m_year,m_month,bank_code"
				 	  + " )a06 ");
           	
		    sqlCmd.append(" select a06.bank_code,bn01.bank_name,a06.m_year || '/' || a06.m_month as m_yearmonth,"		 	  	
			       + "        wlx01.hsien_id,cd01.hsien_name,a06.* "
		           + " from "+a06_table.toString()+",(select * from bn01 where m_year=? )bn01 left join (select  *  from wlx01 where m_year =? ) wlx01 on bn01.bank_no = wlx01.bank_no "
			       + "                         left join ").append(cd01Table).append(" cd01 on wlx01.HSIEN_ID = cd01.HSIEN_ID "
			       + " where a06.bank_code = bn01.bank_no "
			       + " and bn01.bank_no in ("+selectBank_no+")"
			       + condition.toString()
			       + " group by a06.bank_code,bn01.bank_name,a06.m_year,a06.m_month,wlx01.hsien_id,cd01.hsien_name,"+ori_field );
		   for(Object obj : a06_tableList) {
			  sqlCmdList.add(obj) ;
		   }
		   sqlCmdList.add(u_year) ;
		   sqlCmdList.add(u_year) ;
		   for(Object obj : conditionList) {
			   sqlCmdList.add(obj) ;
		   }
		   
		   sqlCmd_sum.append(" select e.FR001W_OUTPUT_ORDER,nvl(e.hsien_id,' ') as hsien_id,"
	   			   	   + " 		  nvl(e.hsien_name,'"+"其他"+"') as hsien_name,"
	   			   	   + "		  d.m_year || '/' || d.m_month as m_yearmonth"
	   			   	   +a06_field_sum_hsien_id
            		   + " from "
            	 	   + "      (select wlx01.hsien_id,cd01.hsien_name,a06.bank_code,bn01.bank_name,a06.m_year,a06.m_month " + a06_field_hsien_id
              	  	   + "		 from ( "
            	 	   + "		   	   select m_year,m_month,bank_code " + a06_field_sum
            		   + "			   from ( "
            		   + "		  		     select m_year,m_month,bank_code,acc_code,amt_3month,amt_6month,amt_1year,amt_2year,amt_over2year,amt_total,amt "
            		   + "					 from (select A06.m_year, A06.m_month,a06.bank_code,"
            		   + "		   		 	   	 		  "+ncacno+ ".acc_range,A06.acc_code,"+ncacno+ ".acc_name,"
            		   + "	  		  	    			  round(amt_3month/1,0) as amt_3month,"// --未滿3個月
            		   + "	  		  	    			  round(amt_6month/1,0) as amt_6month,"// --3個月~6個月
            		   + "	  		  	    			  round(amt_1year/1,0) as amt_1year,"// --6個月~1年
              	  	   + "	  		  	    			  round(amt_2year/1,0) as amt_2year,"// --1年~2年
              	  	   + "	  		  	    			  round(amt_over2year/1,0) as amt_over2year,"// --2年以上
                       + "	  		  	    			  round(amt_total/1,0) as amt_total,"// -- 逾放合計
            		   + "								  0 as amt "
                       + "              		  from A06 LEFT JOIN "+ncacno+ " ON A06.acc_code = "+ncacno+ ".acc_code and "+ncacno+ ".acc_div=?,(select * from bn01 where m_year=? )bn01"
                       + " 						  where A06.m_year = ?"
				       + "  					  and   A06.m_month = ?"
            		   + "					 	  and a06.bank_code = bn01.bank_no and bn01.bank_type=? and bn01.bn_type <> ? and a06.bank_code = bn01.bank_no "
            		   + "			 			  union all "
            		   + "						  select A01.m_year, A01.m_month,a01.bank_code,"+ncacno+ ".acc_range,A01.acc_code,"+ncacno+ ".acc_name,"
            		   + "				 				0 as amt_3month,"// --未滿3個月
            		   + "	  		  	    			0 as amt_6month,"// --3個月~6個月
            		   + "	  		  	   				0 as amt_1year,"// --6個月~1年
              	  	   + "	  		  	   				0 as amt_2year,"// --1年~2年
              	  	   + "	  		  	   				0 as amt_over2year,"// --2年以上
                       + "  	  		  	   			0 as amt_total,"// -- 逾放合計
                       + "			 	  		  	   	round(A01.amt/1,0) as amt"// -- a01放款合計
                       + "			              from (select m_year,m_month,bank_code,decode(acc_code,'990000','970000',acc_code) as acc_code,amt from " + a01_table.toString()
                       + " 				   				)A01 LEFT JOIN "+ncacno+ " ON A01.acc_code = "+ncacno+ ".acc_code and "+ncacno+ ".acc_div= ? ,(select * from bn01 where m_year = ?)bn01 "
            		   + " 						  where A01.m_year = ? "
				       + "  					  and   A01.m_month = ? "
            		   + "						  and a01.bank_code = bn01.bank_no and bn01.bn_type <> ? and bn01.bank_type=? and a01.bank_code = bn01.bank_no "
            		   + "			 			)c "
            		   + "		           )a06 "
            		   + "			   group by a06.m_year,a06.m_month,bank_code "
            		   + "	 	 	 )a06,(select * from bn01 where m_year =? )bn01 left join (select * from wlx01 where m_year = ? )wlx01 on bn01.bank_no = wlx01.bank_no "
            		   + "				   	   left join ").append(cd01Table).append(" cd01 on wlx01.HSIEN_ID = cd01.HSIEN_ID "
            		   + "		where a06.bank_code = bn01.bank_no "
            		   + "		and bn01.bn_type <> ? and bn01.bank_type in ('"+bank_type+"') and a06.bank_code = bn01.bank_no "
            		   + "	    ) d ,(select * from v_bank_location where m_year=?)e  "
            		   + " where d.bank_code(+) =e.bank_no "
            		   + " and (e.hsien_id> ? or e.hsien_id< ? ) "
            		   + " and  d.m_year is not null and d.m_month is not null"
					   + " group by e.FR001W_OUTPUT_ORDER,nvl(e.hsien_id,' '),nvl(e.hsien_name,'"+"其他"+"'),d.m_year,d.m_month");
					   
			sqlCmd_sumList.add("08");
			sqlCmd_sumList.add(u_year);
			sqlCmd_sumList.add(S_YEAR) ;
			sqlCmd_sumList.add(String.valueOf(Integer.parseInt(S_MONTH))) ;
			sqlCmd_sumList.add(bank_type) ;
			sqlCmd_sumList.add("2") ;
			for(Object obj : a01_tableList) {
				sqlCmd_sumList.add(obj) ;
			}
			sqlCmd_sumList.add("08");
			sqlCmd_sumList.add(u_year);
			
			sqlCmd_sumList.add(S_YEAR) ;
			sqlCmd_sumList.add(String.valueOf(Integer.parseInt(S_MONTH))) ;
			
			sqlCmd_sumList.add("2") ;
			sqlCmd_sumList.add(bank_type) ;
			sqlCmd_sumList.add(u_year);
			sqlCmd_sumList.add(u_year);
			sqlCmd_sumList.add("2") ;
			sqlCmd_sumList.add(u_year);
			sqlCmd_sumList.add("Y") ;
			sqlCmd_sumList.add("Y") ;
			if(!order.equals("")){
			    //100.05.12 fix 有挑選排序欄位年月時,查詢SQL error by 2295 				    
			    sqlCmd.append(" order by "+order.replaceAll("order_YYMM", "a06.m_year,m_month"));//各別機構
			    sqlCmd_sum.append(" order by " + order.replaceAll("order_YYMM", "d.m_year,d.m_month"));//全部
  				if( session.getAttribute("SortBy") != null && !((String)session.getAttribute("SortBy")).equals("")){
  		            sqlCmd.append(" " + ((String)session.getAttribute("SortBy")));
  		            sqlCmd_sum.append(" " + ((String)session.getAttribute("SortBy")));
  		         }
  		    }else{
  		       sqlCmd.append(" order by a06.bank_code,a06.m_year,a06.m_month");
  		       sqlCmd_sum.append("  order by e.fr001w_output_order,d.m_year,d.m_month ");
  		    }
            //System.out.println("sqlCmd="+sqlCmd);
            //讀取報表欄位名稱===================================================================================
        	Properties p = new Properties();
			p.load(new FileInputStream(Utility.getProperties("schemaDir")+System.getProperty("file.separator")+"A06_"+bank_type+"_column.TXT"));
			//====================================================================================================

            //Creating Cells
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.createSheet( "report" ); //建立sheet，及名稱
            wb.setSheetName(0, titleName, HSSFWorkbook.ENCODING_UTF_16 );
            HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
            //設定頁面符合列印大小
            //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
            //sheet.setAutobreaks(true); //自動分頁
            sheet.setAutobreaks( false );
            ps.setScale( ( short )100 ); //列印縮放百分比
            ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
            ps.setLandscape( true ); // 設定橫印
            //ps.setFitWidth((short)14);
            HSSFFooter footer = sheet.getFooter();
            //設定樣式和位置(請精減style物件的使用量，以免style物件太多excel報表無法開啟)
			defaultStyle = reportUtil.getDefaultStyle(wb);//有框內文置中
			rightStyle = reportUtil.getRightStyle(wb);//有框內文置右
    		noBorderDefaultStyle = reportUtil.getNoBorderDefaultStyle(wb);//無框內文置中
    		noBorderLeftStyle = reportUtil.getNoBorderLeftStyle(wb);//無框內文置左
			reportUtil.setDefaultStyle(defaultStyle);
			reportUtil.setNoBorderDefaultStyle(noBorderDefaultStyle);
    		titleStyle = reportUtil.getTitleStyle(wb); //標題用
    		columnStyle = reportUtil.getColumnStyle(wb);//報表欄位名稱用--有框內文置中
    		noBoderStyle = reportUtil.getNoBoderStyle(wb);//無框置右
    		//============================================================================
            //設定表頭===============================================================================
            row = sheet.createRow( ( short )1 );
            reportUtil.createCell( wb, row, ( short )1, titleName, titleStyle );

            for(i=2;i<columnLength+5;i++){
              reportUtil.createCell( wb, row, ( short )i, "", noBorderDefaultStyle );
            }
            sheet.addMergedRegion( new Region( ( short )1, ( short )1,
                                               ( short )1,
                                               ( short )(columnLength+3)) );
            //列印年月=======================================================================================
            row = sheet.createRow( ( short )2 );
            row.setHeight((short) 0x200);
            //95.12.01 add 查詢年月區間
            reportUtil.createCell( wb, row, ( short )1, S_YEAR + "年" + S_MONTH + "月~"+E_YEAR + "年" + E_MONTH + "月", titleStyle );            
            for(i=2;i<columnLength+5;i++){
               reportUtil.createCell( wb, row, ( short )i, "", noBorderDefaultStyle );
            }
            sheet.addMergedRegion( new Region( ( short )2, ( short )1,
                                               ( short )2,
                                               ( short )(columnLength+3) ) );
            //======================================================================================================
            row = sheet.createRow( ( short )3 );
            String printTime = Utility.getDateFormat("  HH:mm:ss");
            String printDate = Utility.getDateFormat("yyyy/MM/dd");
            reportUtil.createCell( wb, row, ( short )1, "列印日期："+Utility.getCHTdate(printDate, 1)+printTime, noBoderStyle );
            sheet.addMergedRegion( new Region( ( short )3, ( short )1,
                                               ( short )3,
                                               ( short )(columnLength+3) ) );
            row = sheet.createRow( ( short )4 );
            //列印單位=======================================================================================
            //System.out.println("unit_name="+Utility.getUnitName(Unit));
            //System.out.println("columnLength="+columnLength);
            reportUtil.createCell( wb, row, ( short )1, "單位：新台幣"+Utility.getUnitName(Unit)+"、％", noBorderLeftStyle );
            sheet.addMergedRegion( new Region( ( short )4, ( short )1,
                                               ( short )4,
                                               ( short )2) );
            //設定列印人員==========================================================
            reportUtil.createCell( wb, row, ( short )4, "列印人員："+lguser_name, noBoderStyle );
            sheet.addMergedRegion( new Region( ( short )4, ( short )4,
                                               ( short )4,
                                               ( short )(columnLength+3) ) );
            //報表欄位=======================================================================
            //列印單位代號+機構名稱
            for(i=5;i<8;i++){
                row = sheet.createRow( ( short )i );
                reportUtil.createCell( wb, row, ( short )1, "單位代號", columnStyle );
                reportUtil.createCell( wb, row, ( short )2, "單位名稱", columnStyle );
                reportUtil.createCell( wb, row, ( short )3, "查詢年月", columnStyle );//95.12.05 add 查詢年月
            }
            sheet.addMergedRegion( new Region( ( short )5, ( short )1,
                                               ( short )7,
                                               ( short )1) );
            sheet.addMergedRegion( new Region( ( short )5, ( short )2,
                                               ( short )7,
                                               ( short )2) );
		    sheet.addMergedRegion( new Region( ( short )5, ( short )3,
                                               ( short )7,
                                               ( short )3) );                                               
            row = sheet.createRow( ( short )5 );//大類表頭
            int columnIdx = 4;
            for(i=0;i<btnFieldList_data.size();i++){
               //System.out.println("["+i+"]i="+(String)((List)btnFieldList_data.get(i)).get(1));
               //System.out.println("columnIdx="+columnIdx);
               //設定表頭欄位
               for(j=columnIdx;j<((List)h_column.get(((List)btnFieldList_data.get(i)).get(0))).size() + columnIdx;j++){
                  reportUtil.createCell( wb, row, ( short )j, (String)((List)btnFieldList_data.get(i)).get(1), columnStyle );
               }
               sheet.addMergedRegion( new Region( ( short )5, ( short )columnIdx,
                                               ( short )5,
                                               ( short )(((List)h_column.get(((List)btnFieldList_data.get(i)).get(0))).size() + columnIdx - 1)) );
               columnIdx +=  ((List)h_column.get(((List)btnFieldList_data.get(i)).get(0))).size();
            }

            row = sheet.createRow( ( short ) 6);//細項表頭
            //row.setHeightInPoints(90);//設定細項表頭高度
            columnIdx = 4;
            for(i=0;i<btnFieldList_data.size();i++){
               //System.out.println("["+i+"]i="+(String)((List)btnFieldList_data.get(i)).get(1));
               //System.out.println("columnIdx="+columnIdx);
               detail_column = (List)h_column.get(((List)btnFieldList_data.get(i)).get(0));//取出該大項的細類
               //設定細項表頭欄位
               for(j=0 ;j<detail_column.size();j++){
                  //System.out.println((String)detail_column.get(j)+"="+Utility.ISOtoBig5((String)prop_column_name.get((String)detail_column.get(j))));
                  reportUtil.createCell( wb, row, ( short )columnIdx, Utility.ISOtoUTF8((String)prop_column_name.get((String)detail_column.get(j))), columnStyle );
                  columnIdx ++;
               }
            }
            row = sheet.createRow( ( short ) 7);//細項-科目代號
            columnIdx = 4;
            for(i=0;i<btnFieldList_data.size();i++){
               //System.out.println("["+i+"]i="+(String)((List)btnFieldList_data.get(i)).get(1));
               //System.out.println("columnIdx="+columnIdx);
               detail_column = (List)h_column.get(((List)btnFieldList_data.get(i)).get(0));//取出該大項的細類
               //設定細項表頭欄位
               for(j=0 ;j<detail_column.size();j++){
                  //System.out.println((String)detail_column.get(j)+"="+Utility.ISOtoBig5((String)prop_column_name.get((String)detail_column.get(j))));
                  reportUtil.createCell( wb, row, ( short )columnIdx, (String)detail_column.get(j), columnStyle );
                  columnIdx ++;
               }
            }

            //wb.setRepeatingRowsAndColumns( 0, 1, 8, 1, 3 ); //設定表頭 為固定 先設欄的起始再設列的起始
            wb.setRepeatingRowsAndColumns(0, 1, columnLength+3, 1, 7); //設定表頭 為固定 先設欄的起始再設列的起始

  			//System.out.println("DS006W_Excel.sqlCmd="+sqlCmd);
  			//System.out.println("ori_field="+ori_field);
  			List dbData = null;
  			if(hasBankListALL.equals("false")){  			  
			  dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),sqlCmdList,ori_field+"m_yearmonth");
			  //System.out.println("DS006W_Excel.sqlCmd="+sqlCmd);
			}else{//95.09.04 add 金融機構代號=ALL全部
			  dbData = DBManager.QueryDB_SQLParam(sqlCmd_sum.toString(),sqlCmd_sumList,ori_field+"m_yearmonth");
			  //System.out.println("DS006W_Excel.sqlCmd_sum="+sqlCmd_sum);
			}

			short rowNo = ( short )8;//資料起始列
			//無資料時,顯示訊息========================================================================
			if(dbData == null || dbData.size() == 0){
			   	row = sheet.createRow( rowNo );
                row.setHeight((short) 0x120);
                reportUtil.createCell( wb, row, ( short )1,"無資料" ,noBorderDefaultStyle );
                sheet.addMergedRegion( new Region( ( short )8, ( short )1,
                                               ( short )8,
                                               ( short )(columnLength+5)) );
			}else{
			//有Data時,將DBData寫入===============================================================================================
            acc_code_row = sheet.getRow(7);
            short lastCellNum = acc_code_row.getLastCellNum();
            //System.out.println("lastCellNum="+lastCellNum);
            columnIdx = 1;
            double amt_d = 0.0;
            float amt_f = 0;
            //String amt_type="";
            String amt="";
            String acc_code = "";
            for(i=0;i<dbData.size();i++){
            	bean = (DataObject) dbData.get(i) ;
                acc_code = "";
                row = sheet.createRow( rowNo );
                //System.out.println("rowNo="+rowNo);
                row.setHeight((short) 0x120);
                if(hasBankListALL.equals("false")){
                   //System.out.println("bank_code="+(String) ((DataObject) dbData.get(i)).getValue("bank_code"));
                   reportUtil.createCell( wb, row, ( short )columnIdx, (String) bean.getValue("bank_code"), defaultStyle );//單位代號
                   columnIdx++;
                   reportUtil.createCell( wb, row, ( short )columnIdx, (String) bean.getValue("bank_name"), defaultStyle );//機構名稱
                   columnIdx++;                   
                }else{//縣市別小計
                   //System.out.println("hsien_id="+(String) ((DataObject) dbData.get(i)).getValue("hsien_id"));
                   reportUtil.createCell( wb, row, ( short )columnIdx, (String) bean.getValue("hsien_id"), defaultStyle );//單位代號
                   columnIdx++;
                   reportUtil.createCell( wb, row, ( short )columnIdx, (String) bean.getValue("hsien_name"), defaultStyle );//機構名稱
                   columnIdx++;
                }
                //95.12.05 add 查詢年月
                reportUtil.createCell( wb, row, ( short )columnIdx, bean.getValue("m_yearmonth").toString(), defaultStyle );//查詢年月
                columnIdx++;
                for(int cellIdx =4;cellIdx < (new Short(lastCellNum)).intValue();cellIdx++){
                     amt="";
                     //amt_type = "";
                     cell = acc_code_row.getCell((short)cellIdx);
                     acc_code = cell.getStringCellValue().toLowerCase();
                     //System.out.println("acc_code="+acc_code);
                     if(((DataObject) dbData.get(i)).getValue(acc_code) != null){
                        amt =(((DataObject) dbData.get(i)).getValue(acc_code)).toString();
                     }
                     //amt_type = (((DataObject) dbData.get(i)).getValue("type_"+acc_code)).toString();
                     //System.out.println("amt="+amt);
                     //System.out.println("amt_type="+amt_type);

                     //if(!amt_type.equals("4")){//該值不為利率時.再除以單位4捨五入
                        amt = Utility.getRound(amt,Unit);
                     //}

                     amt = Utility.setCommaFormat(amt);
                     reportUtil.createCell( wb, row, ( short )columnIdx, amt, rightStyle );
                     columnIdx ++;
                }
                columnIdx = 1;
                rowNo++;
            }

            }//end of 有data


            //95.10.02 add 合併acc_code的欄位名稱=================================================================
            columnIdx = 4;
            for(i=0;i<btnFieldList_data.size();i++){
                //System.out.println("["+i+"]i="+(String)((List)btnFieldList_data.get(i)).get(0));
                //System.out.println("columnIdx="+columnIdx);
                //設定表頭欄位.把中間值的acc_code合併成一個欄位只顯示中文名稱=====================================
                 detail_column = (List)h_column.get(((List)btnFieldList_data.get(i)).get(0));//取出該大項的細類
               	//設定細項表頭欄位
               	for(j=0 ;j<detail_column.size();j++){
                     //System.out.println((String)detail_column.get(j)+"="+Utility.ISOtoBig5((String)prop_column_name.get((String)detail_column.get(j))));
                     if(((String)detail_column.get(j)).length() > 6 ){//95.10.02 add 合併acc_code長度為大於6的.欄位名稱
                  	    row = sheet.getRow(6);
                  	    reportUtil.createCell( wb, row, ( short )columnIdx,Utility.ISOtoUTF8((String)prop_column_name.get((String)detail_column.get(j))), columnStyle );
                  	    sheet.addMergedRegion( new Region( ( short )6, ( short )columnIdx,
                                                           ( short )7,
                                               			   ( short )columnIdx) );
                  		columnIdx ++;
               		}
                }
            }

            //設定寬度============================================================
            for ( i = 1; i <= columnLength+3; i++ ) {
                if(i==2){
                  sheet.setColumnWidth( ( short )i,
                                        ( short ) ( 256 * ( 25 + 4 ) ) );//機構名稱
                }else{
                  sheet.setColumnWidth( ( short )i,
                                        ( short ) ( 256 * ( 21 + 4 ) ) );
                }
            }
			//======================================================================================
            //設定涷結欄位
            //sheet.createFreezePane(0,1,0,1);
            footer.setCenter( "Page:" + HSSFFooter.page() + " of " +
                             HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));

            // Write the output to a file
            fileOut = new FileOutputStream( Utility.getProperties("reportDir")+System.getProperty("file.separator")+ titleName+".xls" );
            wb.write( fileOut );
            fileOut.close();

            FileInputStream fin = new FileInputStream(Utility.getProperties("reportDir")+System.getProperty("file.separator")+ titleName+".xls");
			ServletOutputStream out1 = response.getOutputStream();
			byte[] line = new byte[8196];
			int getBytes=0;
			while( ((getBytes=fin.read(line,0,8196)))!=-1 ){
				out1.write(line,0,getBytes);
				out1.flush();
	    	}

			fin.close();
			out1.close();
        } catch ( Exception e ) {
            e.printStackTrace();

        } finally {
            try {
                if ( fileOut != null ) {
                    fileOut.close();
                }
            } catch ( Exception e ) {
                  System.out.println(e.getMessage() );
            }
        }
%>	    		