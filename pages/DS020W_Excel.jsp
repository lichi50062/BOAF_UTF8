<%
// 99.07.30 create 信用部從業人員參加金融相關業務進修情形明細資料彈性報表 by 2660
// 99.08.20 fix 100年度區分縣市別才要加m_year條件 by 2295
// 99.08.31 fix 上課不足時數為空值時,顯示空白,調整欄位寬度 by 2295
//102.11.05 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
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
  if (act.equals("view")) {
    //以上這行設定傳送到前端瀏覽器時的檔名為test1.xls
    //就是靠這一行，讓前端瀏覽器以為接收到一個excel檔
    response.setHeader("Content-disposition","inline; filename=view.xls");
  } else if (act.equals("download")) {
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
  HSSFRow acc_code_row; //讀取acc_code的row
  HSSFCell cell = null; //宣告一個儲存格

  String titleName = "";
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
  String bank_type = "";
  String hasBankListALL = "false";
  String acc_code = "";

  try {
    bank_type = ((String)session.getAttribute("nowbank_type")).equals("")?"6":(String)session.getAttribute("nowbank_type");
    System.out.println("bank_type="+bank_type);
    //儲存報表的目錄================================================================
    File reportDir = new File(Utility.getProperties("reportDir"));
    if (!reportDir.exists()) {
      if (!Utility.mkdirs(Utility.getProperties("reportDir"))) {
        actMsg +=Utility.getProperties("reportDir")+"目錄新增失敗";
      }
    }
    //==============================================================================
    //營運中/已裁撤
    if (session.getAttribute("CANCEL_NO") != null && !((String)session.getAttribute("CANCEL_NO")).equals("")) {
      CANCEL_NO = (String)session.getAttribute("CANCEL_NO");
    }
    //金融機構
    if (session.getAttribute("BankList") != null && !((String)session.getAttribute("BankList")).equals("")) {
      BankList = (String)session.getAttribute("BankList");
      BankList_data = Utility.getReportData(BankList);
      System.out.println("BankList_data.size()="+BankList_data.size());
      System.out.println("BankList_data="+BankList_data);
    }
    //報表欄位
    if (session.getAttribute("btnFieldList") != null && !((String)session.getAttribute("btnFieldList")).equals("")) {
      btnFieldList = (String)session.getAttribute("btnFieldList");
      btnFieldList_data = Utility.getReportData(btnFieldList);
      System.out.println("btnFieldList_data.size()="+btnFieldList_data.size());
      System.out.println("btnFieldList_data="+btnFieldList_data);
    }
    //排序欄位
    if (session.getAttribute("SortList") != null && !((String)session.getAttribute("SortList")).equals("")) {
      SortList = (String)session.getAttribute("SortList");
      SortList_data = Utility.getReportData(SortList);
      System.out.println("SortList_data.size()="+SortList_data.size());
      System.out.println("SortList_data="+SortList_data);
    }

    titleName += "信用部從業人員參加金融相關業務進修情形明細資料";

    //年
    if (session.getAttribute("S_YEAR") != null && !((String)session.getAttribute("S_YEAR")).equals("")) {
      S_YEAR = (String)session.getAttribute("S_YEAR");
    }
    //年
    if (session.getAttribute("E_YEAR") != null && !((String)session.getAttribute("E_YEAR")).equals("")) {
      E_YEAR = (String)session.getAttribute("E_YEAR");
    }
    //95.12.05 增加年月區間
    //月
    if (session.getAttribute("S_MONTH") != null && !((String)session.getAttribute("S_MONTH")).equals("")) {
      S_MONTH = (String)session.getAttribute("S_MONTH");
    }
    //月
    if (session.getAttribute("E_MONTH") != null && !((String)session.getAttribute("E_MONTH")).equals("")) {
      E_MONTH = (String)session.getAttribute("E_MONTH");
    }
    //金額單位
    if (session.getAttribute("Unit") != null && !((String)session.getAttribute("Unit")).equals("")) {
      Unit = (String)session.getAttribute("Unit");
    }

    //讀取欄位大類所包含的細項===================================================================================
    Properties prop_column = new Properties();
    prop_column.load(new FileInputStream(Utility.getProperties("schemaDir")+System.getProperty("file.separator")+"WLX_TRAINNING_detail.TXT"));
    //=======================================================================================================================
    //取出欄位細項將資料存入MAP-->key=大類acc_code,value=細項acc_code=============================================================
    HashMap h_column = new HashMap();//儲存column大類,及其細項的acc_code
    List detail_column = new LinkedList();
    String column_tmp = "";
    String selectacc_code = "";//選取的detail科目代號
    String a08_field_sum = "";
    String ori_field = "";
    String a08acc_code = "";			
    int columnLength = 0;//column個數
    for (i=0; i < btnFieldList_data.size(); i++) {
      column_tmp = "";
      column_tmp = (String)prop_column.get((String)((List)btnFieldList_data.get(i)).get(0));
      System.out.println("column_tmp="+column_tmp);
      if (!column_tmp.equals("")) {
        detail_column = Utility.getStringTokenizerData(column_tmp,"+");
        //System.out.println(detail_column);
        if (detail_column != null && detail_column.size() != 0) {
          columnLength += detail_column.size();//累加總欄位個數
          for(j=0; j<detail_column.size(); j++) {
            //wlx07_m_creditacc_code = (String)detail_column.get(j);
            selectacc_code +="'"+a08acc_code+"'";
            //95.09.26 add=================================================================================================================
            ori_field += a08acc_code;
            //====================================================================================================================================
            if (j < detail_column.size()-1) {
              selectacc_code +=",";
              ori_field += ",";
            }
          }
          //System.out.println("select acc_code="+selectacc_code);
        } else {
        }
        h_column.put((String)((List)btnFieldList_data.get(i)).get(0),detail_column);
        if (i < btnFieldList_data.size()-1) {
          selectacc_code +=",";
          ori_field +=",";
        }
      }
    }
    System.out.println("select acc_code="+selectacc_code);
    System.out.println("h_column.size()="+h_column.size());
    //讀取報表欄位名稱===================================================================================
    Properties prop_column_name = new Properties();
    prop_column_name.load(new FileInputStream(Utility.getProperties("schemaDir")+System.getProperty("file.separator")+"WLX_TRAINNING_column.TXT"));
    //====================================================================================================
    String selectBank_no = "";//選取的金融機構代號
    //金融機構代號=============================================================
    if (BankList_data != null && BankList_data.size() != 0) {
      for (i=0; i<BankList_data.size(); i++) {
        //95.09.04 判斷機構代號是否為ALL:全部===============================================
        if (((String)((List)BankList_data.get(i)).get(0)).equals("ALL")) {
          hasBankListALL = "true";
        }
        selectBank_no += "'" + (String)((List)BankList_data.get(i)).get(0) + "'";
        if (i < BankList_data.size()-1) selectBank_no += ",";
      }
      System.out.println("select bank_no="+selectBank_no);
    }
    //==============================================================================
    String order = "";//排序欄位
    //排序欄位=========================================================================
    if (SortList_data != null && SortList_data.size() != 0) {
      for (i=0; i<SortList_data.size(); i++) {
        order += (String)((List)SortList_data.get(i)).get(0);            		              		
        if (i < SortList_data.size() -1 ) order += ",";
      }
      System.out.println("order="+order);
    }
    //====================================================================================
  			/*
  			//各別機構,按bank_code,m_year,m_month排序
            select bank_no,bank_name,
                   m_yearmonth,
            	   hsien_id,hsien_name,fr001w_output_order,
            	   warnaccount_cnt,
				   limitaccount_cnt,
				   erroraccount_cnt,
				   otheraccount_cnt,
				   depositaccount_tcnt
            from
            (
              select bn01.bank_no,bn01.bank_name,
                     m_year || '/' || decode(substr(to_char(m_month),2),'','0','') || decode(substr(to_char(m_month),1),'1','1',substr(to_char(m_month),1)) as m_yearmonth,
              		 m_year,m_month,
              		 nvl(cd01.hsien_id,' ')       as  hsien_id ,
                     nvl(cd01.hsien_name,'OTHER') as  hsien_name,
                     cd01.FR001W_output_order     as  FR001W_output_order,
                     warnaccount_cnt,
				   	 limitaccount_cnt,
				   	 erroraccount_cnt,
				   	 otheraccount_cnt,
				   	 depositaccount_tcnt
              from  (select * from cd01 where cd01.hsien_id <> 'Y') cd01
              left join wlx01 on wlx01.hsien_id=cd01.hsien_id
              left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type='6'
              left join (select * from A08
              	   		 where (to_char(m_year * 100 + m_month) >= 9501 and to_char(m_year * 100 + m_month) <= 9810) 
            			 and bank_code in ('6060019','6030016')
              	   		)A08 on bn01.bank_no = A08.bank_code
              where A08.bank_code  <>  ' ' 
              group by bn01.bank_no,bn01.bank_name,
                       m_year || '/' || decode(substr(to_char(m_month),2),'','0','') || decode(substr(to_char(m_month),1),'1','1',substr(to_char(m_month),1)),
              		   m_year,m_month,nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,
                       warnaccount_cnt,limitaccount_cnt,erroraccount_cnt,otheraccount_cnt,depositaccount_tcnt	
            ) A08
            order by fr001w_output_order,bank_no,m_year,m_month
            
            //縣市別.按縣市.年月sort(拿掉總計)
            select *
            from
            (
            select  cd01.FR001W_output_order as FR001W_output_order,
            		nvl(cd01.hsien_id,' ') as  hsien_id ,
                 	nvl(cd01.hsien_name,'OTHER')  as  hsien_name,
            		m_year || '/' || decode(substr(to_char(m_month),2),'','0','') || decode(substr(to_char(m_month),1),'1','1',substr(to_char(m_month),1)) as m_yearmonth,
            		m_year,m_month,       	
            		sum(A08.warnaccount_cnt)  as warnaccount_cnt,           		
            		sum(A08.limitaccount_cnt) as limitaccount_cnt,            		
            		sum(A08.erroraccount_cnt)  as erroraccount_cnt,                    
                	sum(A08.otheraccount_cnt)  as otheraccount_cnt,
             		sum(A08.depositaccount_tcnt )  as depositaccount_tcnt 
            from  (select * from cd01 where cd01.hsien_id <> 'Y') cd01
            left join wlx01 on wlx01.hsien_id=cd01.hsien_id
            left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type='6'
            left join (select * from A08 
             	   	   where (to_char(m_year * 100 + m_month) >= 9501 and to_char(m_year * 100 + m_month) <= 9810) 
            		  )A08  on  bn01.bank_no = A08.bank_code
            where m_year is not null
            and   m_month is not null			
            group by cd01.FR001W_output_order,nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),
            		 m_year || '/' || decode(substr(to_char(m_month),2),'','0','') || decode(substr(to_char(m_month),1),'1','1',substr(to_char(m_month),1)),
            		 m_year,m_month)A08   
            order by fr001w_output_order,m_yearmonth,m_year,m_month             
            */            
    String column = "";//選取欄位
    String condition = "";//其他條件
    //add 營運中/已裁撤============================
    condition += " and bn01.bn_type "+(CANCEL_NO.equals("N")?"<>":"=")+" '2' and wlx07_m_credit.bank_no = bn01.bank_no ";
    //======================================================
    String sqlCmd="";

    sqlCmd = "Select m_year," //年度
           + "m_month," //月份
           + "TO_CHAR(m_year,999) || TO_CHAR(m_month,99) as M_YEARMONTH,"
           + "tbank_no," //總機構代號
           + "hsien_id," //縣市別代號
           + "hsien_name," //縣市名稱
           + "FR001W_OUTPUT_ORDER,"
           + "bank_no," //總分支機構代號(金融機構代號)
           + "bank_name," //金融機構名稱
           + "name," //學員姓名
           + "position_code,position_name," //職稱
           + "train_place,train_place_name," //訓練機構
           + "course_name," //課程名稱
           + "course_hour," //課程時數
           + "F_TRANSCHINESEDATE(begin_date) || '~' || F_TRANSCHINESEDATE(end_date) as course_period," //上課期間
           + "licno," //課程名稱
           + "other_trainplace," //其他訓練機構備註
           + "count_person," //其他訓練機構備註
           + "count_hour," //上課時數合計
           + "wlx02count," //分支機構家數
           + "total_hour," //年度所需總時數
           + "count_hour - total_hour as less_hour " //進修不足時數
           + "From ( Select WT.m_year,WT.m_month,hsien_id,hsien_name,FR001W_OUTPUT_ORDER,WT.tbank_no,WT.bank_no,bank_name,name,"
                  + "position_code,F_TRANSCODE('034',position_code) as position_name,"
                  + "train_place,F_TRANSCODE('035',train_place) as train_place_name,"
                  + "course_name,course_hour,begin_date,end_date,"
                  + "licno,other_trainplace,person.count_person,sumhour.count_hour,"
                  + "wlx02.wlx02count,(wlx02.wlx02count+1) * 16 as total_hour "
                  + "From wlx_trainning WT "
                  + "Left Join ba01 On WT.bank_no = ba01.bank_no "
                  ///99.08.20,100年後才加m_year+ "Left Join ba01 On WT.m_year = ba01.m_year And WT.bank_no = ba01.bank_no "
                  + "Left Join ( Select tbank_no,WT.m_year,count(name) as count_person " //同一個年度/總機構代號統計人次
                  + "From wlx_trainning WT Left Join ba01 on WT.bank_no = ba01.bank_no "                  
                              //99.08.20,100年後才加m_year+ "From wlx_trainning WT Left Join ba01 on WT.m_year = ba01.m_year And WT.bank_no = ba01.bank_no "
                              + "Group By tbank_no,WT.m_year) person "
                  + "On WT.tbank_no = person.tbank_no And WT.m_year = person.m_year "
                  + "Left Join ( Select WT.tbank_no,WT.m_year,sum(course_hour) as count_hour " //同一個年度/總機構代號統計上課時數
                              + "From wlx_trainning WT "
                              + "Group By tbank_no,m_year) sumhour "
                  + "On WT.tbank_no = sumhour.tbank_no And WT.m_year = sumhour.m_year "
                  + "Left Join ( Select tbank_no,count(*) as wlx02count " //國內營業分支機構家數
                              + "From wlx02 "
                              + "Where CANCEL_NO <> 'Y' OR CANCEL_NO IS NULL "
                              //99.08.20,100年後才加m_year+ "Where m_year = 99 And CANCEL_NO <> 'Y' OR CANCEL_NO IS NULL "
                              + "Group By tbank_no) wlx02 "
                  + "On wlx02.tbank_no = WT.tbank_no "
                  + "Left Join wlx01 on WT.tbank_no = wlx01.bank_no " //總機構地址
                  + "Left Join cd01 on wlx01.HSIEN_ID = cd01.HSIEN_ID " //對應的縣市別 	
                  + "Where (TO_CHAR(m_year * 100 + m_month) >= " + S_YEAR + S_MONTH //查詢年月區間
                  + " And TO_CHAR(m_year * 100 + m_month) <= " + E_YEAR + E_MONTH + ") ";
    if (hasBankListALL.equals("false")) {
      sqlCmd += "And WT.tbank_no In (" + selectBank_no + ") "; //所挑選的總機構代號
    }
    sqlCmd += "Order By WT.m_year,WT.m_month,cd01.FR001W_OUTPUT_ORDER,WT.tbank_no,WT.seq_no)";

    if (!order.equals("")) {
      //各別機構
      sqlCmd += " order by " + order ;
      if (session.getAttribute("SortBy") != null && !((String)session.getAttribute("SortBy")).equals("")) {
        sqlCmd += " " + ((String)session.getAttribute("SortBy"));
      }
    } else {
      sqlCmd += " Order By m_year,m_month,tbank_no,bank_no";
    }
    //System.out.println("sqlCmd="+sqlCmd);

    //讀取報表欄位名稱===================================================================================
    Properties p = new Properties();
    p.load(new FileInputStream(Utility.getProperties("schemaDir")+System.getProperty("file.separator")+"WLX_TRAINNING_column.TXT"));
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

    for(i=2; i<columnLength+5; i++){
      reportUtil.createCell( wb, row, ( short )i, "", noBorderDefaultStyle );
    }
    sheet.addMergedRegion( new Region( ( short )1, ( short )1,
                                       ( short )1, ( short )(columnLength+4)));
    //列印年月=======================================================================================
    row = sheet.createRow( ( short )2 );
    row.setHeight((short) 0x200);
    //95.12.06 add 查詢年月區間
    reportUtil.createCell( wb, row, ( short )1, S_YEAR + "年" + S_MONTH + "月~"+E_YEAR + "年" + E_MONTH + "月", titleStyle );                                    
    for (i=2; i < columnLength+5; i++) {
      reportUtil.createCell( wb, row, ( short )i, "", noBorderDefaultStyle );
    }
    sheet.addMergedRegion( new Region( ( short )2, ( short )1,
                                       ( short )2, ( short )(columnLength+4) ) );
    //======================================================================================================
    row = sheet.createRow( ( short )3 );
    String printTime = Utility.getDateFormat("  HH:mm:ss");
    String printDate = Utility.getDateFormat("yyyy/MM/dd");
    reportUtil.createCell( wb, row, ( short )1, "列印日期："+Utility.getCHTdate(printDate, 1)+printTime, noBoderStyle );
    sheet.addMergedRegion( new Region( ( short )3, ( short )1,
                                       ( short )3, ( short )(columnLength+4) ) );
    row = sheet.createRow( ( short )4 );
    //列印單位=======================================================================================
    //System.out.println("unit_name="+Utility.getUnitName(Unit));
    //System.out.println("columnLength="+columnLength);
    //98.04.21戶數無須列印單位
    //reportUtil.createCell( wb, row, ( short )1, Utility.ISOtoBig5("單位：新台幣")+Utility.getUnitName(Unit)+Utility.ISOtoBig5("、％"), noBorderLeftStyle );
    //sheet.addMergedRegion( new Region( ( short )4, ( short )1,
    //                                   ( short )4,
    //                                   ( short )2) );
    //設定列印人員==========================================================
    reportUtil.createCell( wb, row, ( short )1, "列印人員："+lguser_name, noBoderStyle );
    sheet.addMergedRegion( new Region( ( short )4, ( short )1,
                                       ( short )4, ( short )(columnLength+4) ) );
    //報表欄位=======================================================================
    //列印單位代號+機構名稱
    for (i=5; i<8; i++) {
      row = sheet.createRow( ( short )i );
      reportUtil.createCell( wb, row, ( short )1, "年度", columnStyle );
      reportUtil.createCell( wb, row, ( short )2, "月份", columnStyle );
      reportUtil.createCell( wb, row, ( short )3, "金融機構代號", columnStyle );
      reportUtil.createCell( wb, row, ( short )4, "金融機構名稱", columnStyle );
      for (j=5; j<columnLength+5; j++) { 
        reportUtil.createCell( wb, row, ( short )j, "", columnStyle );
      }
    }
    sheet.addMergedRegion( new Region( ( short )5, ( short )1,
                                       ( short )7, ( short )1) );
    sheet.addMergedRegion( new Region( ( short )5, ( short )2,
                                       ( short )7, ( short )2) );
    sheet.addMergedRegion( new Region( ( short )5, ( short )3,
                                       ( short )7, ( short )3) );
    sheet.addMergedRegion( new Region( ( short )5, ( short )4,
                                       ( short )7, ( short )4) );

    int columnIdx = 5;

    //wb.setRepeatingRowsAndColumns( 0, 1, 8, 1, 3 ); //設定表頭 為固定 先設欄的起始再設列的起始
    wb.setRepeatingRowsAndColumns(0, 1, columnLength+3, 1, 7); //設定表頭 為固定 先設欄的起始再設列的起始

    //System.out.println("DS020W_Excel.sqlCmd="+sqlCmd);
    //System.out.println("ori_field="+ori_field);
    List dbData = null;
    dbData = DBManager.QueryDB_SQLParam(sqlCmd,null,"m_year,m_month,course_hour,wlx02count,total_hour,count_hour,less_hour");
    //System.out.println("DS020W_Excel.sqlCmd="+sqlCmd);

    short rowNo = ( short )8;//資料起始列
    short start_rowNo = (short)8;
    short end_rowNo = (short)8;
    //無資料時,顯示訊息========================================================================
    if (dbData == null || dbData.size() == 0) {
      row = sheet.createRow( rowNo );
      row.setHeight((short) 0x120);
      reportUtil.createCell( wb, row, ( short )1,"無資料" ,noBorderDefaultStyle );
      sheet.addMergedRegion( new Region( ( short )8, ( short )1,
                                         ( short )8, ( short )(columnLength+5)) );
    } else {
      //有Data時,將DBData寫入===============================================================================================
      acc_code_row = sheet.getRow(7);
      short lastCellNum = acc_code_row.getLastCellNum();
      //System.out.println("lastCellNum="+lastCellNum);
      columnIdx = 1;
      double amt_d = 0.0;
      float amt_f = 0;
      String amt="";
      String prtbank_code = "";
      String tmptbank_no = "";
      DataObject bean = null;
      for (i=0; i<dbData.size(); i++) {
        acc_code = "";
        bean = (DataObject) dbData.get(i);
        row = sheet.createRow( rowNo );
        //System.out.println("rowNo="+rowNo);
        row.setHeight((short) 0x120);
        reportUtil.createCell( wb, row, ( short )columnIdx, bean.getValue("m_year").toString(), defaultStyle );//查詢年月
        columnIdx++;
        reportUtil.createCell( wb, row, ( short )columnIdx, bean.getValue("m_month").toString(), defaultStyle );//查詢年月
        columnIdx++;

          //System.out.println("bank_code="+(String) bean.getValue("bank_no"));
          reportUtil.createCell( wb, row, ( short )columnIdx, (String) bean.getValue("bank_no"), defaultStyle );//單位代號
          columnIdx++;
          reportUtil.createCell( wb, row, ( short )columnIdx, (String) bean.getValue("bank_name"), defaultStyle );//機構名稱
          columnIdx++;
          prtbank_code = (String) bean.getValue("bank_no");

        for (j=0; j < btnFieldList_data.size(); j++) {
          column_tmp = "";
          column_tmp = (String)prop_column.get((String)((List)btnFieldList_data.get(j)).get(0));

          if (column_tmp.equals("course_hour")) {
            reportUtil.createCell( wb, row, ( short )columnIdx, bean.getValue(column_tmp).toString(), defaultStyle );
          } else if (column_tmp.equals("count_hour")) {
            if (tmptbank_no.equals("") || !tmptbank_no.equals((String) bean.getValue("tbank_no"))) {
              reportUtil.createCell( wb, row, ( short )columnIdx, bean.getValue(column_tmp).toString(), defaultStyle );
              end_rowNo = --rowNo;
              rowNo++;
              System.out.println(tmptbank_no+":start_rowNo="+start_rowNo);
              System.out.println("end_rowNo="+end_rowNo);
              if(end_rowNo >= start_rowNo){
                 sheet.addMergedRegion( new Region( start_rowNo, ( short )columnIdx, end_rowNo, ( short )columnIdx) );    
              }
            } else {
              reportUtil.createCell( wb, row, ( short )columnIdx, "", defaultStyle );
            }
          } else if (column_tmp.equals("less_hour")) {
            if (tmptbank_no.equals("") || !tmptbank_no.equals((String) bean.getValue("tbank_no"))) {
              if(bean.getValue(column_tmp) == null){
                 reportUtil.createCell( wb, row, ( short )columnIdx, "", defaultStyle );
              }else{
                 reportUtil.createCell( wb, row, ( short )columnIdx, bean.getValue(column_tmp).toString(), defaultStyle );
                 if(end_rowNo >= start_rowNo){
                    sheet.addMergedRegion( new Region( start_rowNo, ( short )columnIdx,end_rowNo, ( short )columnIdx) );      
                 }
                 System.out.println(tmptbank_no+":start_rowNo="+start_rowNo);
                 System.out.println("end_rowNo="+end_rowNo);
                 start_rowNo = end_rowNo;                                                 
              }
            } else {
              reportUtil.createCell( wb, row, ( short )columnIdx, "", defaultStyle );
            }            
          } else {
            reportUtil.createCell( wb, row, ( short )columnIdx, (String) bean.getValue(column_tmp), defaultStyle );
          }
          columnIdx++;
        }
        tmptbank_no = (String) bean.getValue("tbank_no");
        columnIdx = 1;
        rowNo++;
      }
    }//end of 有data

    //95.10.02 add 合併acc_code的欄位名稱=================================================================
    columnIdx = 5;
    for (i=0; i<btnFieldList_data.size(); i++) {
      //System.out.println("["+i+"]i="+(String)((List)btnFieldList_data.get(i)).get(0));
      //System.out.println("columnIdx="+columnIdx);
      //設定表頭欄位.把中間值的acc_code合併成一個欄位只顯示中文名稱=====================================
      detail_column = (List)h_column.get(((List)btnFieldList_data.get(i)).get(0));//取出該大項的細類
      //設定細項表頭欄位
      for (j=0; j<detail_column.size(); j++) {
        acc_code = (String)detail_column.get(j);
        //System.out.println((String)detail_column.get(j)+"="+Utility.ISOtoBig5((String)prop_column_name.get((String)detail_column.get(j))));
        row = sheet.getRow(5);
        reportUtil.createCell( wb, row, ( short )columnIdx,Utility.ISOtoUTF8((String)prop_column_name.get(acc_code)), columnStyle );
        sheet.addMergedRegion( new Region( ( short )5, ( short )columnIdx,
                                           ( short )7, ( short )columnIdx) );
        columnIdx ++;
      }
    }
    
    column_tmp = ""; 
    //設定寬度============================================================
    for ( i = 1; i <= columnLength+4; i++ ) {
       if(i==4){
          sheet.setColumnWidth( ( short )i,( short ) ( 256 * ( 25 + 4 ) ) );//金融機構名稱4
       }else if(i==7){ //訓練機構       
          sheet.setColumnWidth( ( short )i,( short ) ( 256 * ( 40 + 4 ) ) );//訓練機構7
       }else if(i==8){ //課程名稱       
          sheet.setColumnWidth( ( short )i,( short ) ( 256 * ( 25 + 4 ) ) );//課程名稱8   
       }else if(i==9){ //上課期間       
          sheet.setColumnWidth( ( short )i,( short ) ( 256 * ( 30 + 4 ) ) );//上課期間9     
       }else if(i==11){ //證書字號     
          sheet.setColumnWidth( ( short )i,( short ) ( 256 * ( 25 + 4 ) ) );//證書字號11     
       }else if(i==12){ //其他訓練機構備註  
          sheet.setColumnWidth( ( short )i,( short ) ( 256 * ( 25 + 4 ) ) );//其他訓練機構備註 12             
       }else{
               sheet.setColumnWidth( ( short )i,( short ) ( 256 * ( 15 + 4 ) ) );
       } 
    }

    //======================================================================================
    //設定涷結欄位
    //sheet.createFreezePane(0,1,0,1);
    footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
    footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));

    // Write the output to a file
    fileOut = new FileOutputStream( Utility.getProperties("reportDir")+System.getProperty("file.separator")+ titleName+".xls" );
    wb.write( fileOut );
    fileOut.close();

    FileInputStream fin = new FileInputStream(Utility.getProperties("reportDir")+System.getProperty("file.separator")+ titleName+".xls");
    ServletOutputStream out1 = response.getOutputStream();
    byte[] line = new byte[8196];
    int getBytes=0;
    while (((getBytes=fin.read(line,0,8196)))!=-1 ) {
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