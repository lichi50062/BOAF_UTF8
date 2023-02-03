<%
//98.10.13 create 各農漁會信用部申報者名稱.通訊錄 by 2295
//99.12.11 fix sqlInjection by 2808
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.util.*,java.io.*" %>
<%@ page import="org.apache.poi.poifs.filesystem.*,org.apache.poi.hssf.usermodel.*" %>
<%@ page import="org.apache.poi.hssf.util.Region" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.report.reportUtil" %>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>

<%
   response.setContentType("application/msexcel;charset=UTF-8");//以上這行設定本網頁為excel格式的網頁   
   String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");	
   String report_no = ( request.getParameter("REPORT_NO")==null ) ? "" : (String)request.getParameter("REPORT_NO");
   String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? "" : (String)request.getParameter("S_YEAR");				
   String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? "" : (String)request.getParameter("S_MONTH");				
   String lguser_name = ( session.getAttribute("muser_name")==null ) ? "" : (String)session.getAttribute("muser_name");	  
   System.out.println("ZZ031W_Excel.act="+act);
   System.out.println("ZZ031W_Excel.report_no="+report_no);
   System.out.println("ZZ031W_Excel.S_YEAR="+S_YEAR);
   System.out.println("ZZ031W_Excel.S_MONTH="+S_MONTH);   
   response.setHeader("Content-Disposition","attachment; filename=download.xls");      
%>
<%
	
	String actMsg = "";
	FileOutputStream fileOut = null;      	
    HSSFCellStyle defaultStyle;   
    HSSFCellStyle noBorderDefaultStyle;    
    HSSFCellStyle titleStyle;
	HSSFCellStyle columnStyle;
	HSSFCellStyle noBoderStyle;
	HSSFRow row;	
	HSSFCell cell = null;//宣告一個儲存格
	String titleName = "各農漁會信用部申報者通訊錄";
	DataObject bean = null;
    reportUtil reportUtil = new reportUtil();  
    int columnIdx = 1;
	int i = 0;		
	try{
			//儲存報表的目錄================================================================
        	File reportDir = new File(Utility.getProperties("reportDir"));       
    		if(!reportDir.exists()){
     			if(!Utility.mkdirs(Utility.getProperties("reportDir"))){
     	   			actMsg +=Utility.getProperties("reportDir")+"目錄新增失敗";
     			}    
    		}
    		//===============================================================================
    		/*  			
            select wml01.bank_code,bn01.bank_name,wml01.user_id,wml01.user_name,muser_data.m_telno
			from wml01 
		    left join muser_data on wml01.user_id = muser_data.muser_id 
			left join bn01 on wml01.bank_code = bn01.bank_no
			where report_no='A01' and m_year=94 and m_month=8
			order by bank_code
  			*/  			
			
            String sqlCmd="";          
			List paramList =new ArrayList() ;
			String yy = Integer.parseInt(S_YEAR)>99?"100":"99";
			sqlCmd = " select wml01.bank_code,bn01.bank_name,wml01.user_id,wml01.user_name,muser_data.m_telno "
				   + " from wml01 "
		    	   + " left join muser_data on wml01.user_id = muser_data.muser_id "
				   + " left join (select * from bn01 where m_year=? )bn01 on wml01.bank_code = bn01.bank_no "
				   + " where report_no=? "
				   + " and wml01.m_year= ?"
				   + " and wml01.m_month=?"
				   + " order by bank_code ";
			paramList.add(yy) ;
			paramList.add(report_no) ;
			paramList.add(S_YEAR);
			paramList.add(S_MONTH) ;
			//System.out.println("sqlCmd="+sqlCmd);
            
            //Creating Cells
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.createSheet( "report" ); //建立sheet，及名稱
            wb.setSheetName(0, titleName+"("+report_no+")", HSSFWorkbook.ENCODING_UTF_16 );
            HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定            
            //設定頁面符合列印大小              
            sheet.setAutobreaks( false );
            ps.setScale( ( short )100 ); //列印縮放百分比
            ps.setPaperSize( ( short )9 ); //設定紙張大小 A4            
            ps.setLandscape( false ); // 設定橫印            
            HSSFFooter footer = sheet.getFooter();            
            //設定樣式和位置(請精減style物件的使用量，以免style物件太多excel報表無法開啟)
			defaultStyle = reportUtil.getDefaultStyle(wb);//有框內文置中
			
    		noBorderDefaultStyle = reportUtil.getNoBorderDefaultStyle(wb);//無框內文置中    		
			reportUtil.setDefaultStyle(defaultStyle);
			reportUtil.setNoBorderDefaultStyle(noBorderDefaultStyle);			
    		titleStyle = reportUtil.getTitleStyle(wb); //標題用
    		columnStyle = reportUtil.getColumnStyle(wb);//報表欄位名稱用--有框內文置中			                                               
    		noBoderStyle = reportUtil.getNoBoderStyle(wb);//無框置右			                                               
    		//============================================================================                        
            //設定表頭(各農漁會信用部申報者通訊錄)===============================================================================
            row = sheet.createRow( ( short )1 );
            reportUtil.createCell( wb, row, ( short )1, titleName+"("+report_no+")", titleStyle );
            
            for(i=2;i<6;i++){
              reportUtil.createCell( wb, row, ( short )i, "", noBorderDefaultStyle );
            }            
            sheet.addMergedRegion( new Region( ( short )1, ( short )1,( short )1,( short )5) );
            //======================================================================================================                                                                     
            row = sheet.createRow( ( short )2 );            
            String printTime = Utility.getDateFormat("  HH:mm:ss");
            String printDate = Utility.getDateFormat("yyyy/MM/dd");                                    
            reportUtil.createCell( wb, row, ( short )1, "列印日期："+Utility.getCHTdate(printDate, 1)+printTime, noBoderStyle );
            sheet.addMergedRegion( new Region( ( short )2, ( short )1,
                                               ( short )2,
                                               ( short )5 ) );            
                                   
            //設定列印人員========================================================== 
            row = sheet.createRow( ( short )3 );            
            reportUtil.createCell( wb, row, ( short )1, "列印人員："+lguser_name, noBoderStyle );
            sheet.addMergedRegion( new Region( ( short )3, ( short )1,
                                               ( short )3,
                                               ( short )5 ) );            
            //報表欄位=======================================================================                   
            row = sheet.createRow( ( short )4 );
            reportUtil.createCell( wb, row, ( short )1, "機構代號", columnStyle );               
            reportUtil.createCell( wb, row, ( short )2, "機構名稱", columnStyle );   
            reportUtil.createCell( wb, row, ( short )3, "申報者代號", columnStyle );
            reportUtil.createCell( wb, row, ( short )4, "申報者姓名", columnStyle );
            reportUtil.createCell( wb, row, ( short )5, "聯絡電話", columnStyle );
                        
            wb.setRepeatingRowsAndColumns(0, 1, 5, 1, 4); //設定表頭 為固定 先設欄的起始再設列的起始
  			
  			List dbData = null;  					  
			dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
			paramList.clear() ;
			short rowNo = ( short )5;//資料起始列     
			//無資料時,顯示訊息========================================================================			
			if(dbData == null || dbData.size() == 0){
			   	row = sheet.createRow( rowNo );     
                row.setHeight((short) 0x120); 
                reportUtil.createCell( wb, row, ( short )1,"無資料" ,noBorderDefaultStyle ); 
                sheet.addMergedRegion( new Region( ( short )rowNo, ( short )1,
                                               ( short )rowNo,
                                               ( short )5) );
			}else{
			    //將DBData寫入===============================================================================================                 
                for(i=0;i<dbData.size();i++){            
                    row = sheet.createRow( rowNo );  
                    row.setHeight((short) 0x120);  
                    bean = (DataObject)dbData.get(i); 
                   
                    reportUtil.createCell( wb, row, ( short )columnIdx, (String) bean.getValue("bank_code"), defaultStyle );//1.機構代號
                    columnIdx++;
                    reportUtil.createCell( wb, row, ( short )columnIdx, (String) bean.getValue("bank_name"), defaultStyle );//2.機構名稱
                    columnIdx++;
                    reportUtil.createCell( wb, row, ( short )columnIdx, (String) bean.getValue("user_id"), defaultStyle );//3.申報者代號
                    columnIdx++;
                    reportUtil.createCell( wb, row, ( short )columnIdx, (String) bean.getValue("user_name"), defaultStyle );//4.申報者名稱
                    columnIdx++;
                    reportUtil.createCell( wb, row, ( short )columnIdx, (String) bean.getValue("m_telno"), defaultStyle );//5.電話
                    columnIdx++;
                    columnIdx = 1;
                    rowNo++;                
                }
            }//end of 有data   
           
            
            //設定寬度============================================================            
            for ( i = 1; i <= 5; i++ ) {                
                if(i==2){
                  sheet.setColumnWidth( ( short )i,
                                        ( short ) ( 256 * ( 25 + 4 ) ) );//機構名稱
                }else{
                  sheet.setColumnWidth( ( short )i,
                                        ( short ) ( 256 * ( 15 + 4 ) ) );
                }                        
            }
			//======================================================================================
            //設定涷結欄位
            //sheet.createFreezePane(0,1,0,1);
            footer.setCenter( "Page:" + HSSFFooter.page() + " of " +HSSFFooter.numPages() );		                                 
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