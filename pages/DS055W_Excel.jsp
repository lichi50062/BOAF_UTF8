<%
//102.07.09 created N年內及目前月份經營統計資料 by 2968
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
<%@ page import="java.text.*" %>
<%@ page import="java.lang.*" %>
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
	NumberFormat nf = NumberFormat.getInstance();
	nf.setMinimumFractionDigits(2);                        // 若小數點不足二位，則補足二位
	String actMsg = "";
	FileOutputStream fileOut = null;      	
    HSSFCellStyle defaultStyle;
    HSSFCellStyle rightStyle;
    HSSFCellStyle noBorderDefaultStyle;
    HSSFCellStyle noBorderLeftStyle;
    HSSFCellStyle titleStyle;
	HSSFCellStyle columnStyle;
	HSSFCellStyle noBoderStyle;
	HSSFCellStyle leftStyle;
	HSSFRow row;
	HSSFRow acc_code_row;//讀取acc_code的row
	HSSFCell cell = null;//宣告一個儲存格
	String titleName = "";
    reportUtil reportUtil = new reportUtil();
    String BankList = "";//儲存bank_code/bank_name
    String btnFieldList = "";//儲存所選取的大類acc_code/名稱 
    String Unit = "";//列印單位
    String S_YEAR = "";//年
    String S_YEAR1 = "";//N年內
    String S_MONTH = "";//月
    String bank_type = "";//農(漁)會別
    String bank_type_name = "";
    String field_debit="";
    String field_120700="";
    String field_840740_rate_avg="";
    String field_840740_rate_max="";
    String field_over_rate_avg="";
    String field_over_rate_max="";
    String tmp="";
    String tmpYear="";
    List btnFieldList_data = null;
    List SortList_data = null;    
	int i = 0;
	int j= 0;
	int tmpCell=0;
	String lguser_name = ( session.getAttribute("muser_name")==null ) ? "" : (String)session.getAttribute("muser_name");
	String report_no="DS055W";
	System.out.println("=================DS055W.Excel.jsp start===================");		
	try{	      	
			//儲存報表的目錄================================================================
        	File reportDir = new File(Utility.getProperties("reportDir"));       
    		if(!reportDir.exists()){
     			if(!Utility.mkdirs(Utility.getProperties("reportDir"))){
     	   			actMsg +=Utility.getProperties("reportDir")+"目錄新增失敗";
     			}    
    		}
    		//==============================================================================    		
    		
			//報表欄位
			if(session.getAttribute("btnFieldList") != null && !((String)session.getAttribute("btnFieldList")).equals("")){
		   		btnFieldList = (String)session.getAttribute("btnFieldList");
		   		btnFieldList_data = Utility.getReportData(btnFieldList);
		   		System.out.println("btnFieldList_data.size()="+btnFieldList_data.size());		   
		   		System.out.println("btnFieldList_data="+btnFieldList_data);		   
			}
			//年
			if(session.getAttribute("S_YEAR") != null && !((String)session.getAttribute("S_YEAR")).equals("")){
		  		S_YEAR = (String)session.getAttribute("S_YEAR");		  				   
			}
			//月
			if(session.getAttribute("S_MONTH") != null && !((String)session.getAttribute("S_MONTH")).equals("")){
		  		S_MONTH = (String)session.getAttribute("S_MONTH");		  				   
			}
			//N年內
			if(session.getAttribute("S_YEAR1") != null && !((String)session.getAttribute("S_YEAR1")).equals("")){
			    S_YEAR1 = (String)session.getAttribute("S_YEAR1");		  				   
			}
			//金額單位
			if(session.getAttribute("Unit") != null && !((String)session.getAttribute("Unit")).equals("")){
		  		Unit = (String)session.getAttribute("Unit");		  				   
			}
			//農(漁)會別
			if(session.getAttribute("nowbank_type") != null && !((String)session.getAttribute("nowbank_type")).equals("")){
			    bank_type = (String)session.getAttribute("nowbank_type");		  				   
			}
			if("6".equals(bank_type)){
			    bank_type_name = "農會";
	        }else if("7".equals(bank_type)){
	            bank_type_name = "漁會";
	        }else if("ALL".equals(bank_type)){
	            bank_type_name = "農漁會";
	        }
			titleName += (bank_type_name+S_YEAR1+"年內及目前月份經營統計資料");
			//讀取欄位大類所包含的細項===================================================================================        	
        	Properties prop_column = new Properties();
			StringBuffer sql = new StringBuffer();		
	        List paramList = new ArrayList();
	        DataObject bean = null;
            sql.append(" select acc_code,acc_name from  ncacno_ds where acc_tr_type=? order by acc_range ");
     		paramList.add(report_no);
     		List qData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,""); 
			for(i=0;i<qData.size();i++){
			    bean = (DataObject)qData.get(i);
			    prop_column.setProperty((String)bean.getValue("acc_code"),(String)bean.getValue("acc_code"));  
	        }		
			//=======================================================================================================================			
			//取出欄位細項將資料存入MAP-->key=大類acc_code,value=細項acc_code=============================================================
			HashMap h_column = new HashMap();//儲存column大類,及其細項的acc_code			
			List detail_column = new LinkedList();
			String column_tmp = "";			
			String selectacc_code = "";//選取的detail科目代號		
			String field_sum="";			
			String ori_field="";
			int columnLength=0;//column個數
			for(i=0;i<btnFieldList_data.size();i++){			    
			    column_tmp = "";
			    column_tmp = (String)prop_column.get((String)((List)btnFieldList_data.get(i)).get(0));
			    if(!column_tmp.equals("")){
			        detail_column = Utility.getStringTokenizerData(column_tmp,"+");
			        if(detail_column != null && detail_column.size() != 0){               
			           columnLength += detail_column.size();//累加總欄位個數
              		   for(j=0;j<detail_column.size();j++){
            	 		   selectacc_code +=(String)detail_column.get(j); 
            	 		   ori_field += (String)detail_column.get(j);
            	 		   if(j < detail_column.size()-1){
            	 		      selectacc_code +=",";            	 		      
            	 		      ori_field +=",";            	 		      
            	 		   }  
               		   }                              		   
            	   }  			      
			       h_column.put((String)((List)btnFieldList_data.get(i)).get(0),detail_column);			       
			       if(i < btnFieldList_data.size()-1){
			          selectacc_code +=",";
			          ori_field +=",";            	
			       }   
			    }
			}
			System.out.println("h_column.size()="+h_column.size());
			System.out.println("select acc_code="+selectacc_code);        		
  			
            //String column = "";//選取欄位           
			//======================================================
            StringBuffer sqlCmd =new StringBuffer();            
            List sqlCmd_paramList = new ArrayList();
            int parami = 0;
            sqlCmd.append("select a01.m_yearmonth,a01.bank_type ");
            for(int k=1;k<=columnLength;k++){
                if("field_debit".equals(splitStr(selectacc_code, k-1))){
                    sqlCmd.append(",field_debit,type_field_debit ");
                }else if("field_120700".equals(splitStr(selectacc_code, k-1))){
                    sqlCmd.append(",field_120700,type_field_120700 ");
                }else if("field_840740_rate_avg".equals(splitStr(selectacc_code, k-1))){
                    sqlCmd.append(",field_840740_rate_avg,'4' as type_field_840740_rate_avg ");
                    field_840740_rate_avg="Y";
                }else if("field_840740_rate_max".equals(splitStr(selectacc_code, k-1))){
                    sqlCmd.append(",field_840740_rate_max,'4' as type_field_840740_rate_max ");
                    field_840740_rate_max="Y";
                }else if("field_over_rate_avg".equals(splitStr(selectacc_code, k-1))){
                    sqlCmd.append(",field_over_rate_avg,'4' as type_field_over_rate_avg ");
                    field_over_rate_avg="Y";
                }else if("field_over_rate_max".equals(splitStr(selectacc_code, k-1))){
                    sqlCmd.append(",field_over_rate_max,'4' as type_field_over_rate_max ");
                    field_over_rate_max="Y";                    
                }else{
                    sqlCmd.append(","+splitStr(selectacc_code, k-1)+" ");
                }
            }
			sqlCmd.append("from "); 
			sqlCmd.append("( "); 
			sqlCmd.append(" select  m_year ||  LPAD(m_month,2,'0') as m_yearmonth,bank_type ");
			for(int k=1;k<=columnLength;k++){
                if("field_debit".equals(splitStr(selectacc_code, k-1))){
                    sqlCmd.append("                    ,round(sum(decode(acc_code,?,amt,0)) /?,0) as field_debit ");//
    				sqlCmd.append("                    ,round(sum(decode(acc_code,?,type,'')) /1,2) as type_field_debit ");//type:2為金額,4為利率
    				sqlCmd_paramList.add("field_debit");
    				sqlCmd_paramList.add(Unit);
    				sqlCmd_paramList.add("field_debit");
                }else if("field_120700".equals(splitStr(selectacc_code, k-1))){
                    sqlCmd.append("                    ,round(sum(decode(acc_code,?,amt,0)) /?,0) as field_120700 ");//
    				sqlCmd.append("                    ,round(sum(decode(acc_code,?,type,'')) /1,2) as type_field_120700 ");//type:2為金額,4為利率
    				sqlCmd_paramList.add("field_120700");
    				sqlCmd_paramList.add(Unit);
    				sqlCmd_paramList.add("field_120700");
                }else if( !"field_840740_rate_avg".equals(splitStr(selectacc_code, k-1))
                        && !"field_840740_rate_max".equals(splitStr(selectacc_code, k-1))
                        && !"field_over_rate_avg".equals(splitStr(selectacc_code, k-1))
                        && !"field_over_rate_max".equals(splitStr(selectacc_code, k-1))){
                    if("field_credit".equals(splitStr(selectacc_code, k-1)) 
                        || "field_320300".equals(splitStr(selectacc_code, k-1))
                        || "field_320000".equals(splitStr(selectacc_code, k-1))
                        || "field_transfer".equals(splitStr(selectacc_code, k-1))
                        || "field_310000".equals(splitStr(selectacc_code, k-1))
                        || "field_net".equals(splitStr(selectacc_code, k-1))
                        || "field_150200".equals(splitStr(selectacc_code, k-1))
                        || "field_noassure".equals(splitStr(selectacc_code, k-1))
                        || "field_990611".equals(splitStr(selectacc_code, k-1))
                        || "field_over".equals(splitStr(selectacc_code, k-1))
                        || "field_840740".equals(splitStr(selectacc_code, k-1))
                        || "field_backup".equals(splitStr(selectacc_code, k-1))
                        || "field_840760".equals(splitStr(selectacc_code, k-1))){
                        sqlCmd.append("                    ,round(sum(decode(acc_code,?,amt,0)) /?,0) as ").append(splitStr(selectacc_code, k-1));
                        sqlCmd.append("					   ,'2' as type_").append(splitStr(selectacc_code, k-1));
                        sqlCmd_paramList.add(splitStr(selectacc_code, k-1));
                        sqlCmd_paramList.add(Unit);
                    }else{
                        sqlCmd.append("                    ,round(sum(decode(acc_code,?,amt,0)) /1,2) as ").append(splitStr(selectacc_code, k-1));
                        sqlCmd.append("					   ,'4' as type_").append(splitStr(selectacc_code, k-1));
                        sqlCmd_paramList.add(splitStr(selectacc_code, k-1));
                    }
                }
            }
			sqlCmd.append("            from a01_operation ");
			sqlCmd.append("            where to_char(m_year * 100 + m_month) in ( ");
			if(!"".equals(S_YEAR1)){
    	  		for(int k=Integer.parseInt(S_YEAR)-Integer.parseInt(S_YEAR1);k<Integer.parseInt(S_YEAR);k++){
    	  		    sqlCmd.append("?,");
    	  		  	sqlCmd_paramList.add(k+"12");
    	  		    if(k==Integer.parseInt(S_YEAR)-Integer.parseInt(S_YEAR1)){
    	  		        tmp= String.valueOf(k) ;
    	  		    }else{
    	  		        tmp= tmp+","+String.valueOf(k);
    	  		    }
    	  		}
	  		}
			sqlCmd.append("?) ");
			sqlCmd_paramList.add(S_YEAR+addZeroForNum(S_MONTH,2));
			if(!"".equals(tmp)){
	  		    tmp= tmp+","+S_YEAR;
	  		}else{
	  		    tmp = S_YEAR;
	  		}
			sqlCmd.append("            and bank_code='ALL' and (type='2' or type='4') ");
			sqlCmd.append("            and bank_type in (?) ");
			sqlCmd_paramList.add(bank_type);
			sqlCmd.append("            and hsien_id =' ' ");
			sqlCmd.append("            group by m_year,m_month,bank_type ");
			sqlCmd.append("            order by m_year,m_month "); 
			sqlCmd.append(")a01 ");
			//--有挑選到field_840740_rate_avg or field_840740_rate_max or field_over_rate_avg or field_over_rate_max才加入
			if("Y".equals(field_840740_rate_avg)||"Y".equals(field_840740_rate_max)||"Y".equals(field_over_rate_avg)||"Y".equals(field_over_rate_max)){
			    sqlCmd.append(", (select t1.m_yearmonth,t1.bank_type,field_840740_rate_avg,field_840740_rate_max,field_over_rate_avg,field_over_rate_max ");
			    sqlCmd.append("from "); 
			    sqlCmd.append("( ");   
			    sqlCmd.append("  select m_year || LPAD(m_month,2,'0') as m_yearmonth ");
			    if("ALL".equals(bank_type)){
			    	sqlCmd.append("    ,'ALL' as bank_type ");
			    	sqlCmd.append("    ,round(avg(amt),2) as field_840740_rate_avg ");
			    }else{
			        sqlCmd.append("    ,bank_type ");
				    sqlCmd.append("    ,round(decode(acc_code,'field_840740_rate',avg(amt),0),2) as field_840740_rate_avg ");
			    }
			    sqlCmd.append("    from a01_operation ");
			    sqlCmd.append("   where to_char(m_year * 100 + m_month) in (");
			    if(!"".equals(S_YEAR1)){
	    	  		for(int k=Integer.parseInt(S_YEAR)-Integer.parseInt(S_YEAR1);k<Integer.parseInt(S_YEAR);k++){
	    	  		    sqlCmd.append("?,");
	    	  		  	sqlCmd_paramList.add(k+"12");
	    	  		}
		  		}
				sqlCmd.append("?) ");
				sqlCmd_paramList.add(S_YEAR+addZeroForNum(S_MONTH,2));
			    sqlCmd.append("     and bank_code !='ALL' ");
				if("ALL".equals(bank_type)){
				    sqlCmd.append("     and bank_type in ('6','7') ");
				}else{
				    sqlCmd.append("     and bank_type in (?) ");
				    sqlCmd_paramList.add(bank_type);
				}
			    sqlCmd.append("     and acc_code in (?) ");
			    sqlCmd_paramList.add("field_840740_rate");
			    sqlCmd.append("   group by m_year,m_month ");
			    if(!"ALL".equals(bank_type)){
			        sqlCmd.append(" ,bank_type,acc_code ");
				}
			    sqlCmd.append("   order by m_year,m_month ");  
			    sqlCmd.append("  )t1, ");
			    sqlCmd.append(" (select m_year || LPAD(m_month,2,'0') as m_yearmonth ");
			    if("ALL".equals(bank_type)){
			    	sqlCmd.append("    ,'ALL' as bank_type ");
			    	sqlCmd.append("    ,round(avg(amt),2) as field_over_rate_avg ");
			    }else{
			        sqlCmd.append("    ,bank_type ");
				    sqlCmd.append("    ,round(decode(acc_code,'field_over_rate',avg(amt),0),2) as field_over_rate_avg ");
			    }
			    sqlCmd.append("    from a01_operation ");
			    sqlCmd.append("   where to_char(m_year * 100 + m_month) in (");
			    if(!"".equals(S_YEAR1)){
	    	  		for(int k=Integer.parseInt(S_YEAR)-Integer.parseInt(S_YEAR1);k<Integer.parseInt(S_YEAR);k++){
	    	  		    sqlCmd.append("?,");
	    	  		  	sqlCmd_paramList.add(k+"12");
	    	  		}
		  		}
				sqlCmd.append("?) ");
				sqlCmd_paramList.add(S_YEAR+addZeroForNum(S_MONTH,2));
			    sqlCmd.append("     and bank_code !='ALL' ");
				
				if("ALL".equals(bank_type)){
				    sqlCmd.append("     and bank_type in ('6','7') ");
				}else{
				    sqlCmd.append("     and bank_type in (?) ");
				    sqlCmd_paramList.add(bank_type);
				}
			    sqlCmd.append("     and acc_code in (?) ");
			    sqlCmd_paramList.add("field_over_rate");
			    sqlCmd.append("   group by m_year,m_month ");
			    if(!"ALL".equals(bank_type)){
			        sqlCmd.append("   ,bank_type,acc_code ");
			    }
			    sqlCmd.append("   order by m_year,m_month ");  
			    sqlCmd.append(" )t2, ");
			    sqlCmd.append(" (select m_year || LPAD(m_month,2,'0') as m_yearmonth ");
			    if("ALL".equals(bank_type)){
			    	sqlCmd.append("    ,'ALL' as bank_type ");
			    	sqlCmd.append("    , round(max(amt),2) as field_over_rate_max ");
			    }else{
			        sqlCmd.append("    ,bank_type ");
				    sqlCmd.append("    ,round(decode(acc_code,'field_over_rate',max(amt),0),2) as field_over_rate_max ");
			    }
			    sqlCmd.append("   from a01_operation ");
			    sqlCmd.append("  where to_char(m_year * 100 + m_month) in (");
			    if(!"".equals(S_YEAR1)){
	    	  		for(int k=Integer.parseInt(S_YEAR)-Integer.parseInt(S_YEAR1);k<Integer.parseInt(S_YEAR);k++){
	    	  		    sqlCmd.append("?,");
	    	  		  	sqlCmd_paramList.add(k+"12");
	    	  		}
		  		}
				sqlCmd.append("?) ");
				sqlCmd_paramList.add(S_YEAR+addZeroForNum(S_MONTH,2));
			    sqlCmd.append("    and bank_code !='ALL' ");
				
				if("ALL".equals(bank_type)){
				    sqlCmd.append("    and bank_type in ('6','7') ");
				}else{
				    sqlCmd.append("    and bank_type in (?) ");
				    sqlCmd_paramList.add(bank_type);
				}
			    sqlCmd.append("    and acc_code in (?) ");
			    sqlCmd_paramList.add("field_over_rate");
			    sqlCmd.append("  group by m_year,m_month ");
			    if(!"ALL".equals(bank_type)){
			        sqlCmd.append("  ,bank_type,acc_code ");
			    }
			    sqlCmd.append("  order by m_year,m_month "); 
			    sqlCmd.append(" )t3, ");       
			    sqlCmd.append(" (select m_year || LPAD(m_month,2,'0') as m_yearmonth ");
			    if("ALL".equals(bank_type)){
			    	sqlCmd.append("    ,'ALL' as bank_type ");
			    	sqlCmd.append("    , round(max(amt),2) as field_840740_rate_max ");
			    }else{
			        sqlCmd.append("    ,bank_type ");
				    sqlCmd.append("    ,round(decode(acc_code,'field_840740_rate',max(amt),0),2) as field_840740_rate_max ");
			    }
			    sqlCmd.append("    from a01_operation ");
			    sqlCmd.append("   where to_char(m_year * 100 + m_month) in (");
			    if(!"".equals(S_YEAR1)){
	    	  		for(int k=Integer.parseInt(S_YEAR)-Integer.parseInt(S_YEAR1);k<Integer.parseInt(S_YEAR);k++) {
	    	  		    sqlCmd.append("?,");
	    	  		  	sqlCmd_paramList.add(k+"12");
	    	  		}
		  		}
				sqlCmd.append("?) ");
				sqlCmd_paramList.add(S_YEAR+addZeroForNum(S_MONTH,2));
			    sqlCmd.append("     and bank_code !='ALL' ");
				if("ALL".equals(bank_type)){
				    sqlCmd.append("     and bank_type in ('6','7') ");
				}else{
				    sqlCmd.append("     and bank_type in (?) ");
					sqlCmd_paramList.add(bank_type);
				}
			    sqlCmd.append("     and acc_code in (?) ");
			    sqlCmd_paramList.add("field_840740_rate");
			    sqlCmd.append("   group by m_year,m_month ");
			    if(!"ALL".equals(bank_type)){
			        sqlCmd.append("  ,bank_type,acc_code ");
			    }
			    sqlCmd.append("   order by m_year,m_month "); 
			    sqlCmd.append(" )t4 ");
			    sqlCmd.append("  where t1.m_yearmonth=t2.m_yearmonth(+) and t1.bank_type=t2.bank_type(+) ");
			    sqlCmd.append("    and t1.m_yearmonth = t3.m_yearmonth(+) and t1.bank_type=t3.bank_type(+) ");
			    sqlCmd.append("    and t1.m_yearmonth=t4.m_yearmonth(+) and t1.bank_type=t4.bank_type(+) ");
			    sqlCmd.append(")a01_1 ");
			    sqlCmd.append("where a01.m_yearmonth=a01_1.m_yearmonth(+) and a01.bank_type=a01_1.bank_type(+) ");
			}
			List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),sqlCmd_paramList,"m_yearmonth,"+ori_field);
            System.out.println("dbData.size()="+dbData.size());
			//讀取報表欄位名稱===================================================================================        	
        	Properties prop_column_name = new Properties();
			sql = new StringBuffer();		
	        paramList = new ArrayList();
	        bean = null;
            sql.append(" select acc_code,acc_name from ncacno_ds where acc_tr_type=? order by acc_range ");
     		paramList.add(report_no);
     		qData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,""); 
			for(i=0;i<qData.size();i++){
			    bean = (DataObject)qData.get(i);
			    prop_column_name.setProperty((String)bean.getValue("acc_code"),(String)bean.getValue("acc_name"));  
			}	
			//====================================================================================================
			
            //Creating Cells
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.createSheet( "report" ); //建立sheet，及名稱
            wb.setSheetName(0, titleName, HSSFWorkbook.ENCODING_UTF_16 );
            HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定            
            //設定頁面符合列印大小          
            sheet.setAutobreaks( false );
            ps.setScale( ( short )100 ); //列印縮放百分比
            ps.setPaperSize( ( short )9 ); //設定紙張大小 A4            
            ps.setLandscape( true ); // 設定橫印           
            HSSFFooter footer = sheet.getFooter();            
            //設定樣式和位置(請精減style物件的使用量，以免style物件太多excel報表無法開啟)
			defaultStyle = reportUtil.getDefaultStyle(wb);//有框內文置中
			rightStyle = reportUtil.getRightStyle(wb);//有框內文置右
			leftStyle = reportUtil.getLeftStyle(wb);//有框內文置右
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
            
            for(i=2;i<columnLength+2;i++){
              reportUtil.createCell( wb, row, ( short )i, "", noBorderDefaultStyle );
            }            
            sheet.addMergedRegion( new Region( ( short )1, ( short )1,
                                               ( short )1,
                                               ( short )(Integer.parseInt(S_YEAR1)+2)) );
                                                           
            //列印年月=======================================================================================
            row = sheet.createRow( ( short )2 );            
            row.setHeight((short) 0x200);
            //95.12.01 add 查詢年月區間
            reportUtil.createCell( wb, row, ( short )1, S_YEAR + "年" + S_MONTH + "月", titleStyle );
            for(i=2;i<columnLength+2;i++){
               reportUtil.createCell( wb, row, ( short )i, "", noBorderDefaultStyle );
            }            
            sheet.addMergedRegion( new Region( ( short )2, ( short )1,
                                               ( short )2,
                                               ( short )(Integer.parseInt(S_YEAR1)+2) ) );
            //======================================================================================================                                                                     
            row = sheet.createRow( ( short )3 );            
            String printTime = Utility.getDateFormat("  HH:mm:ss");
            String printDate = Utility.getDateFormat("yyyy/MM/dd");                                    
            reportUtil.createCell( wb, row, ( short )1, "列印日期："+Utility.getCHTdate(printDate, 1)+printTime, noBoderStyle );
            sheet.addMergedRegion( new Region( ( short )3, ( short )1,
                                               ( short )3,
                                               ( short )(Integer.parseInt(S_YEAR1)+2) ) );            
            row = sheet.createRow( ( short )4 );                        
            //列印單位=======================================================================================            
            //System.out.println("columnLength="+columnLength);
            reportUtil.createCell( wb, row, ( short )1, "單位：新台幣"+Utility.getUnitName(Unit)+"、％", noBorderLeftStyle );                                                
            /*
            sheet.addMergedRegion( new Region( ( short )4, ( short )1,
                                               ( short )4,
                                               ( short )2) );          
            */                                   
            //設定列印人員==========================================================            
            reportUtil.createCell( wb, row, ( short )2, "列印人員："+lguser_name, noBoderStyle );
            sheet.addMergedRegion( new Region( ( short )4, ( short )2,
                                               ( short )4,
                                               ( short )(Integer.parseInt(S_YEAR1)+2) ) );            
            //報表欄位=======================================================================
            //列印單位代號+機構名稱            
            for(i=5;i<8;i++){
                row = sheet.createRow( ( short )i );                
                reportUtil.createCell( wb, row, ( short )1, "項目", columnStyle );
                
            }           
            sheet.addMergedRegion( new Region( ( short )5, ( short )1,
                                               ( short )7,
                                               ( short )1) ); 
            int columnIdx = 2;                                                                                         
            //大類表頭
            for(i=0;i<=Integer.parseInt(S_YEAR1);i++){
               	int cellNum=2+i;
               	row = sheet.createRow( ( short )5 );
               	if(i==Integer.parseInt(S_YEAR1)){
               	 	reportUtil.createCell( wb, row, ( short )cellNum, splitStr(tmp, i)+"年"+S_MONTH+"月底", columnStyle );
               		tmpCell = cellNum;
               	}else{
               	 	reportUtil.createCell( wb, row, ( short )cellNum, splitStr(tmp, i)+"年底", columnStyle );
               	}
               	row = sheet.createRow( ( short )6 );
               	reportUtil.createCell( wb, row, ( short )cellNum,"", columnStyle );
               	row = sheet.createRow( ( short )7 );
               	reportUtil.createCell( wb, row, ( short )cellNum, "", columnStyle );
               	sheet.addMergedRegion( new Region( ( short )5, ( short )cellNum,( short )7,( short )cellNum ));
               	//columnIdx ++;
            }
            //設定表頭欄位(選項)
            for(i=0;i<btnFieldList_data.size();i++){
               int rowNum=8+i;
               row = sheet.createRow( ( short )rowNum );
               cell=row.createCell((short)1);
               reportUtil.createCell( wb, row, (short)1, (String)((List)btnFieldList_data.get(i)).get(1), leftStyle );
   	           sheet.setColumnWidth((short)1, (short)(((String)((List)btnFieldList_data.get(i)).get(1)).length()));
   	           
            }
			//將DBData寫入========================================================================			
			System.out.println("S_YEAR+S_MONTH="+S_YEAR+S_MONTH);
			int cellNum = (short)2;
			if(dbData.size() > 0){
			    int isInsYY = 0;
			    for(int n=Integer.parseInt(S_YEAR)-Integer.parseInt(S_YEAR1);n<=Integer.parseInt(S_YEAR);n++){
                    for(int t=0;t<dbData.size();t++){
                        int rowNum = (short)8;
                    	bean = (DataObject)dbData.get(t);
                    	if((String.valueOf(n)+"12").equals(bean.getValue("m_yearmonth").toString())){
                    		System.out.println("12test");
                    	    for(short k=0;k<btnFieldList_data.size();k++){
                    	        row = sheet.createRow( rowNum );
                                row.setHeight((short) 0x120);
                                if("field_credit".equals(splitStr(ori_field,k)) 
                                        || "field_320300".equals(splitStr(ori_field,k))
                                        || "field_320000".equals(splitStr(ori_field,k))
                                        || "field_transfer".equals(splitStr(ori_field,k))
                                        || "field_310000".equals(splitStr(ori_field,k))
                                        || "field_net".equals(splitStr(ori_field,k))
                                        || "field_150200".equals(splitStr(ori_field,k))
                                        || "field_noassure".equals(splitStr(ori_field,k))
                                        || "field_990611".equals(splitStr(ori_field,k))
                                        || "field_over".equals(splitStr(ori_field,k))
                                        || "field_840740".equals(splitStr(ori_field,k))
                                        || "field_backup".equals(splitStr(ori_field,k))
                                        || "field_840760".equals(splitStr(ori_field,k))
                                        || "field_debit".equals(splitStr(ori_field,k))
                                        || "field_120700".equals(splitStr(ori_field,k))){
                                    reportUtil.createCell( wb, row, (short)cellNum, Utility.setCommaFormat(bean.getValue(splitStr(ori_field,k)).toString()), rightStyle );
                                }else{
	                    	    	reportUtil.createCell( wb, row, (short)cellNum, Utility.setCommaFormat((bean.getValue(splitStr(ori_field,k))==null)?"":bean.getValue(splitStr(ori_field,k)).toString()), rightStyle );
                                }
	                    	    rowNum++;
                    	    }
                    	    isInsYY = n;
                    	}else{
                    		System.out.println("test1="+bean.getValue("m_yearmonth").toString());
                    	    if((S_YEAR+S_MONTH).equals(bean.getValue("m_yearmonth").toString())){
                    	    	System.out.println("01test");
	                    	    for(short k=0;k<btnFieldList_data.size();k++){
	                    	        row = sheet.createRow( rowNum );
	                                row.setHeight((short) 0x120);
	                                if("field_credit".equals(splitStr(ori_field,k)) 
	                                        || "field_320300".equals(splitStr(ori_field,k))
	                                        || "field_320000".equals(splitStr(ori_field,k))
	                                        || "field_transfer".equals(splitStr(ori_field,k))
	                                        || "field_310000".equals(splitStr(ori_field,k))
	                                        || "field_net".equals(splitStr(ori_field,k))
	                                        || "field_150200".equals(splitStr(ori_field,k))
	                                        || "field_noassure".equals(splitStr(ori_field,k))
	                                        || "field_990611".equals(splitStr(ori_field,k))
	                                        || "field_over".equals(splitStr(ori_field,k))
	                                        || "field_840740".equals(splitStr(ori_field,k))
	                                        || "field_backup".equals(splitStr(ori_field,k))
	                                        || "field_840760".equals(splitStr(ori_field,k))
	                                        || "field_debit".equals(splitStr(ori_field,k))
	                                        || "field_120700".equals(splitStr(ori_field,k))){
	                                    reportUtil.createCell( wb, row, (short)tmpCell, Utility.setCommaFormat(bean.getValue(splitStr(ori_field,k)).toString()), rightStyle );
	                                }else{
		                    	    	reportUtil.createCell( wb, row, (short)tmpCell, Utility.setCommaFormat((bean.getValue(splitStr(ori_field,k))==null)?"":bean.getValue(splitStr(ori_field,k)).toString()), rightStyle );
	                                }
		                    	    rowNum++;
	                    	    }
	                    	    isInsYY = n;
                    	    }
                    	}
                    }
                    if(isInsYY != n){
                        for(int y=isInsYY+1;y<=Integer.parseInt(S_YEAR);y++){
                            int rowNum = (short)8;
                            for(short k=0;k<btnFieldList_data.size();k++){
    	               	        row = sheet.createRow( rowNum );
    	                        row.setHeight((short) 0x120);
    	                   	    reportUtil.createCell( wb, row, (short)cellNum, "", rightStyle );
    	                   	    rowNum++;
    	           	        } 
                            isInsYY = n;
                        }
                    }
                    cellNum++;
			    }
			}else{
			    for(int n=Integer.parseInt(S_YEAR)-Integer.parseInt(S_YEAR1);n<=Integer.parseInt(S_YEAR);n++){
			        int rowNum = (short)8;
			        if(!(tmpYear).equals(String.valueOf(n))){
	        	 	   	for(short k=0;k<btnFieldList_data.size();k++){
	               	        row = sheet.createRow( rowNum );
	                        row.setHeight((short) 0x120);
	                   	    reportUtil.createCell( wb, row, (short)cellNum, "", rightStyle );
	                   	    rowNum++;
	           	       }
	        	 	}
				    tmpYear = String.valueOf(n);
				    cellNum++;
			    }
			}
            //設定寬度============================================================
            short columnWidth = 256 * ( 15 + 4 );
            boolean haveField_992611 = false;
            boolean haveField_84760_rate = false;
            boolean haveField_backup = false;
            boolean haveField_captial_rate  = false;
			for(i=0;i<btnFieldList_data.size();i++){
		    	if("field_990611".equals(splitStr(ori_field, i))){
		            haveField_992611 = true;
		    	}else if("field_840760_rate".equals(splitStr(ori_field, i))){
		            haveField_84760_rate = true;
		    	}else if("field_backup_over_rate".equals(splitStr(ori_field, i))||
		    	        "field_backup_840740_rate".equals(splitStr(ori_field, i))){
		    	    haveField_backup = true;
		    	}else if("field_captial_rate".equals(splitStr(ori_field, i))){
		    	    haveField_captial_rate = true;
		    	}
	        }
            for ( i = 1; i <= Integer.parseInt(S_YEAR1)+3; i++ ) { 
                if(i==1){
                    if(haveField_992611==true){
                        columnWidth = 256 * ( 15 + 32 );
                    }else{
                        if(haveField_84760_rate == true){
                            columnWidth = 256 * ( 15 + 22 ); 
                        }else if(haveField_backup ==true || haveField_captial_rate ==true){
                            columnWidth = 256 * ( 15 + 12 );
                        }else{
                            columnWidth = 256 * ( 15 + 4 );
                        }
                    }
                	sheet.setColumnWidth( ( short )i,columnWidth );   
                }else{
                  	sheet.setColumnWidth( ( short )i,( short ) ( 256 * ( 15 + 4 ) ) );  
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
			System.out.println("=================DS055W.Excel.jsp End===================");
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
<%!
public static String addZeroForNum(String str, int strLength) {
    int strLen = str.length();
    if (strLen < strLength) {
        while (strLen < strLength) {
            StringBuffer sb = new StringBuffer();
            sb.append("0").append(str);//左補零
            //sb.append(str).append("0");//右補零
            str = sb.toString();
            strLen = str.length();
        }
    }
    return str;
}
public static String splitStr(String str, int i) {
    String[] strarray=str.split(",");
    return strarray[i];
}
%>	    		