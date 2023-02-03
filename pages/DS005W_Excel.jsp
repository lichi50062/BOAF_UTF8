<%
// 95.10.18-19 create by 2295
// 95.12.04 add 增加年月區間,可根據年月做sort,拿掉可選欄位做sort by 2295
// 95.12.05 fix 讓User可根據Excel做其他運用.單位代號.機構名稱repeat by 2295
// 99.05.14 fix 縣市合併& sql Injection by 2808
//100.05.12 fix 有挑選排序欄位年月時,查詢SQL error by 2295 
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
<%@ page import="java.util.*" %>
<%@ page import="java.text.SimpleDateFormat" %>
<%@ page import="java.util.StringTokenizer" %>
<%@ page import="java.lang.StringBuffer" %>
<%@ page import="java.lang.Short" %>
<%@ page import="java.lang.Math" %>
<%@ page import="java.text.DecimalFormat" %>
<%@ page import="java.text.MessageFormat" %>

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
	String acc_code = "";
	String u_year = "99" ;//縣市合併用
	String cd01Table = "cd01_99" ;//縣市合併用
	DataObject bean = null ;
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
			
			titleName += "資本適足率資料";		  		
			
			//年
			if(session.getAttribute("S_YEAR") != null && !((String)session.getAttribute("S_YEAR")).equals("")){
		  		S_YEAR = (String)session.getAttribute("S_YEAR");	
		  		if(Integer.parseInt(S_YEAR) >99) {
		  			u_year = "100" ;
		  			cd01Table = "cd01" ;
		  		}
			}
			//年
			if(session.getAttribute("E_YEAR") != null && !((String)session.getAttribute("E_YEAR")).equals("")){
		  		E_YEAR = (String)session.getAttribute("E_YEAR");		  				   
			}
			//95.12.01 增加年月區間
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
			
			//讀取欄位大類所包含的細項===================================================================================        	
        	Properties prop_column = new Properties();
			prop_column.load(new FileInputStream(Utility.getProperties("schemaDir")+System.getProperty("file.separator")+"A05_"+bank_type+"_detail.TXT"));			
			//System.out.println("load="+Utility.getProperties("schemaDir")+System.getProperty("file.separator")+"A05_"+bank_type+"_detail.TXT");
			//System.out.println("prop_column.size()="+prop_column.size());
			//Enumeration e = prop_column.propertyNames();
			// while (e.hasMoreElements())
	        //    System.out.println(e.nextElement());
			//=======================================================================================================================			
			//取出欄位細項將資料存入MAP-->key=大類acc_code,value=細項acc_code=============================================================
			HashMap h_column = new HashMap();//儲存column大類,及其細項的acc_code			
			List detail_column = new LinkedList();
			String column_tmp = "";			
			String selectacc_code = "";//選取的detail科目代號
			String a05_operation_field_sum="";
			String a05_operation_field="";
			String a05_field_sum="";
			String a05_field="";
			String ori_field="";
			String a05acc_code="";
			String maxa05_acc_code="";
			String select_ori_acc_code="";
			String otherGroup="";
			int columnLength=0;//column個數			
			for(i=0;i<btnFieldList_data.size();i++){			    
			    column_tmp = "";
			    column_tmp = (String)prop_column.get((String)((List)btnFieldList_data.get(i)).get(0));
			    //System.out.println("get="+(String)((List)btnFieldList_data.get(i)).get(0));
			    //System.out.println("column_tmp="+column_tmp);
			    if(!column_tmp.equals("")){
			        detail_column = Utility.getStringTokenizerData(column_tmp,"+");
			        //System.out.println(detail_column);
			        if(detail_column != null && detail_column.size() != 0){               
			           columnLength += detail_column.size();//累加總欄位個數
              		   for(j=0;j<detail_column.size();j++){
            	 		   selectacc_code +="'"+(String)detail_column.get(j)+"'";            	
            	 		   //95.09.26 add================================================================================================================= 
            	 		   a05acc_code = (String)detail_column.get(j);
            	 		   if(a05acc_code.length() <= 6){            	 		      
            	 		      if(select_ori_acc_code.indexOf(a05acc_code) == -1){
            	 		         select_ori_acc_code += "'"+a05acc_code+"',";            	 		         
            	 		      }
            	 		      ori_field+= "amt"+a05acc_code;                    	 		      
            	 		      if(a05acc_code.endsWith("N")){//decode(acc_code,'92071N',amt_name,'') as amt92071N--中文名稱
            	 		         a05_field_sum +=  " ,decode(acc_code,'"+a05acc_code+"',amt_name,'') as amt"+a05acc_code;
            	 		         otherGroup += ",decode(acc_code,'"+a05acc_code+"',amt_name,'')";
            	 		      }else if(a05_field_sum.indexOf(" ,sum(decode(acc_code,'"+a05acc_code+"',amt,0)) as amt"+a05acc_code) == -1){
            	 		         a05_field_sum += " ,sum(decode(acc_code,'"+a05acc_code+"',amt,0)) as amt"+a05acc_code;            	            	 		          		               						            	 		                  	 		      
            	 		      }   
            	 		      if(maxa05_acc_code.indexOf(a05acc_code) == -1){
            	 		         maxa05_acc_code += "max(a05.amt"+a05acc_code+") as amt"+a05acc_code+","; 
            	 		      }
            	 		   }else{
            	 		      //sum(decode(acc_code,'920101',amt,0)) as amt920101,			 
			 			   	  //sum(decode(acc_code,'920101',assumed,0)) as assumed920101,
 			 			   	  //sum(decode(acc_code,'920101',amt_assumed,0)) as amt_assumed920101, 			 			   		            	 		      
            	 		      if(select_ori_acc_code.indexOf(a05acc_code.substring(a05acc_code.length()-6,a05acc_code.length())) == -1){
            	 		         select_ori_acc_code += "'"+a05acc_code.substring(a05acc_code.length()-6,a05acc_code.length())+"',";
            	 		      }
            	 		      
            	 		      ori_field += a05acc_code;            	 		                            	 		      
            	 		      if(a05_field_sum.indexOf(" ,sum(decode(acc_code,'"+a05acc_code.substring(a05acc_code.length()-6,a05acc_code.length())+"',"+a05acc_code.substring(0,a05acc_code.length()-6)+",0)) as "+a05acc_code) == -1){
            	 		         a05_field_sum += " ,sum(decode(acc_code,'"+a05acc_code.substring(a05acc_code.length()-6,a05acc_code.length())+"',"+a05acc_code.substring(0,a05acc_code.length()-6)+",0)) as "+a05acc_code; 							                
            	 		      }
            	 		      if(maxa05_acc_code.indexOf(a05acc_code) == -1){
            	 		        maxa05_acc_code += "max(a05."+a05acc_code+") as "+a05acc_code+","; 
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
			}//end of for
			select_ori_acc_code += "'910500'";
			if(a05_field_sum.indexOf("910500") != -1){
               maxa05_acc_code += "max(a05.amt910500) as amt910500 ";
            }else{
              maxa05_acc_code += " '0' ";
            }
            System.out.println("select acc_code="+selectacc_code);               		   
			System.out.println("h_column.size()="+h_column.size());
        	//讀取報表欄位名稱===================================================================================        	
        	Properties prop_column_name = new Properties();
			prop_column_name.load(new FileInputStream(Utility.getProperties("schemaDir")+System.getProperty("file.separator")+"A05_"+bank_type+"_column.TXT"));			
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
            select  a05.bank_code,bn01.bank_name,m_year || '/' || m_month as m_yearmonth,
            		wlx01.hsien_id,cd01.hsien_name,
            		max(a05.amt920101) as amt920101, max(a05.amt920720) as amt920720,max(a05.amt92072N) as amt92072N
            from (
            		select m_year,m_month,bank_code,
             			   sum(decode(acc_code,'920101',amt,0)) as amt920101,			 
             			   sum(decode(acc_code,'920101',assumed,0)) as assumed920101,
             			   sum(decode(acc_code,'920101',amt_assumed,0)) as amt_assumed920101,
             			   sum(decode(acc_code,'92072P',amt,0)) as amt92072P,
             			   sum(decode(acc_code,'920720',amt,0)) as amt920720,
             			   decode(acc_code,'92072N',amt_name,'') as amt92072N			 
            	   from (  select * 
            	           from 
            	           (select bank_code,m_year,m_month, ncacno.acc_code, ncacno.acc_name,
            			   	 		    '' as acc_code_l,
            							0 as assumed,				    
            			    			decode(substr( ncacno .acc_code,length( ncacno .acc_code)),'P',nvl(a05.amt,0)/1000,round(nvl(a05.amt,0)/ 1 ,0)) as amt,
            			    			0 as  amt_assumed,
            			    			nvl(a05.amt_name,'') as amt_name 
            	    		from  ncacno  left join a05 on  ncacno.acc_code=a05.acc_code and  ncacno.acc_code like '91%'
            				where (to_char(m_year * 100 + m_month) >= 9406 and to_char(m_year * 100 + m_month) <= 9409)
            				and   a05.bank_code in ('6030016','6230023','6060086')
            				and   acc_tr_type='A05' 
            				order by  ncacno .acc_range)
            				union
            				(select a05.bank_code,a05.m_year,a05.m_month,ncacno.acc_code,ncacno.acc_name,
            						substr(ncacno.acc_code,length(ncacno.acc_code)) as acc_code_l,
            						a05_assumed.assumed as assumed,
            						decode(substr(ncacno.acc_code,length(ncacno.acc_code)),'P',nvl(a05.amt,0)/1000,nvl(a05.amt,0))as amt,
            						decode(substr(ncacno.acc_code,length(ncacno.acc_code)),'P',nvl(a05.amt,0)/1000,nvl(a05_assumed.assumed,1)*nvl(a05.amt,0)) as amt_assumed,
            						nvl(a05.amt_name,'') as amt_name
            				from ncacno join a05 on ncacno.acc_code=a05.acc_code 
            			 				left join a05_assumed on a05.acc_code=a05_assumed.acc_code 
            				where (to_char(m_year * 100 + m_month) >= 9406 and to_char(m_year * 100 + m_month) <= 9409)
            				and	  acc_tr_type='A05' and ncacno.acc_code like '92%'
            				and   a05.bank_code in ('6030016','6230023','6060086'))
            	  		 ) a05
            	   where (to_char(m_year * 100 + m_month) >= 9406 and to_char(m_year * 100 + m_month) <= 9409)		 
            	    and  bank_code in  ('6030016','6230023','6060086')
             		and  acc_code in ('920101','920201','920501','92072P','920720','92072N')					 
            		group by m_year,m_month,bank_code,decode(acc_code,'92072N',amt_name,'') 
            	)a05 ,bn01 left join wlx01  on bn01.bank_no = wlx01.bank_no
            	  		  left join cd01 on wlx01.HSIEN_ID = cd01.HSIEN_ID  
            where a05.bank_code = bn01.bank_no
            and bn01.bank_no in ('6030016','6230023','6060086') and bn01.bn_type <> '2' 
            and a05.bank_code = bn01.bank_no
            group by a05.bank_code,bn01.bank_name,a05.m_year,a05.m_month,wlx01.hsien_id,cd01.hsien_name
            order by a05.bank_code,a05.m_year,a05.m_month	
  			*/
  			
            String column = "";//選取欄位           
            //String condition = "";//其他條件
            StringBuffer condition = new StringBuffer() ;
            List conditionList = new ArrayList () ;
            //add 營運中/已裁撤============================             
			condition.append(" and bn01.bn_type "+(CANCEL_NO.equals("N")?"<>":"=")+" ?  and a05.bank_code = bn01.bank_no ");
            conditionList.add("2") ;
			//======================================================
            //String sqlCmd="";
			StringBuffer sqlCmd = new StringBuffer() ;
			List sqlCmdList = new ArrayList () ; 
            String sqlCmd_sum="";//縣市別小計
            //String a05_table="";//a05
            StringBuffer a05_table = new StringBuffer() ; 
            List a05_tableList = new ArrayList() ;
			a05_table.append(" ( "
      				  + "	select m_year,m_month,bank_code" + a05_field_sum		     			   
	  			   	  + "	from (  select * "
	  			      + "		     from "
	  			      + "			     (select bank_code,m_year,m_month, ncacno.acc_code, ncacno.acc_name,"
        			  + "	 	   	 		    '' as acc_code_l,"
	    		      + "						0 as assumed,"				    
 	   				  + "	 	    			decode(substr( ncacno .acc_code,length( ncacno .acc_code)),'P',nvl(a05.amt,0)/1000,round(nvl(a05.amt,0)/ 1 ,0)) as amt,"
	    			  + "		    			0 as  amt_assumed,"
 	   	    		  + "		    			nvl(a05.amt_name,'') as amt_name "
				      + "				 from  ncacno  left join a05 on  ncacno.acc_code=a05.acc_code and  ncacno.acc_code like ? "
				      + "                and   (to_char(m_year * 100 + m_month) >= ? and to_char(m_year * 100 + m_month) <= ?)"					     					    					  				      
				      + "				 and   a05.bank_code in ("+selectBank_no+")"	
					  + "				 where acc_tr_type= ?  "
					  + "				 order by  ncacno .acc_range)"
					  + "		 		 union "
					  + "				(select a05.bank_code,a05.m_year,a05.m_month,ncacno.acc_code,ncacno.acc_name,"
	   				  + "						substr(ncacno.acc_code,length(ncacno.acc_code)) as acc_code_l,"
	   				  + "						a05_assumed.assumed as assumed,"
	   				  + "						decode(substr(ncacno.acc_code,length(ncacno.acc_code)),'P',nvl(a05.amt,0)/1000,nvl(a05.amt,0))as amt,"
	   				  + "						decode(substr(ncacno.acc_code,length(ncacno.acc_code)),'P',nvl(a05.amt,0)/1000,nvl(a05_assumed.assumed,1)*nvl(a05.amt,0)) as amt_assumed,"
	   				  + "						nvl(a05.amt_name,'') as amt_name "
					  + "				from ncacno join a05 on ncacno.acc_code=a05.acc_code "
	 				  + "			 				left join a05_assumed on a05.acc_code=a05_assumed.acc_code "
	 				  + "               where (to_char(m_year * 100 + m_month) >= ? and to_char(m_year * 100 + m_month) <= ? )"					     					    					  
					  + "				 and   acc_tr_type=? and ncacno.acc_code like ? "
				      + "				 and   a05.bank_code in ("+selectBank_no+"))"						  
				  	  + "  		  ) a05 "
				  	  + "  where (to_char(m_year * 100 + m_month) >= ? and to_char(m_year * 100 + m_month) <= ? )"					     					    					  				  	  
			          + "  and   bank_code in ("+selectBank_no+")"				      
				      + "  and   acc_code in ("+select_ori_acc_code+")"    
  				   	  + "  group by m_year,m_month,bank_code "+otherGroup
		 			  + " )a05"); 
			a05_tableList.add("91%") ;
			a05_tableList.add(S_YEAR+S_MONTH) ;
			a05_tableList.add(E_YEAR+E_MONTH) ;
			a05_tableList.add("A05") ;
			a05_tableList.add(S_YEAR+S_MONTH) ;
			a05_tableList.add(E_YEAR+E_MONTH) ;
			a05_tableList.add("A05") ;
			a05_tableList.add("92%") ;
			a05_tableList.add(S_YEAR+S_MONTH) ;
			a05_tableList.add(E_YEAR+E_MONTH) ;
			
			sqlCmd.append("  select a05.bank_code,bn01.bank_name,a05.m_year || '/' || a05.m_month as m_yearmonth,"
            	   + "	       wlx01.hsien_id,cd01.hsien_name,"+ maxa05_acc_code	  				
				   + "	from " + a05_table.toString() + ",(select * from bn01 where m_year=? )bn01 left join (select * from wlx01 where m_year=? ) wlx01  on bn01.bank_no = wlx01.bank_no "
				   + "								left join ").append(cd01Table).append(" cd01 on wlx01.HSIEN_ID = cd01.HSIEN_ID " 
				   + " where a05.bank_code = bn01.bank_no "
			       + " and bn01.bank_no in ("+selectBank_no+")"
			       + condition.toString()			
			       + " group by a05.bank_code,bn01.bank_name,a05.m_year,a05.m_month,wlx01.hsien_id,cd01.hsien_name ");
           for(Object obj : a05_tableList) {
        	   sqlCmdList.add(obj);
           }
           sqlCmdList.add(u_year) ;
           sqlCmdList.add(u_year) ;
           for(Object obj : conditionList) {
        	   sqlCmdList.add(obj);
           }
				   
			if(!order.equals("")){	   			
			    //100.05.12 fix 有挑選排序欄位年月時,查詢SQL error by 2295 	
			    //各別機構
			    sqlCmd.append(" order by "+order.replaceAll("order_YYMM", "a05.m_year,m_month") ); 
			    //sqlCmd_sum += " order by " + order;			    			    
  				if( session.getAttribute("SortBy") != null && !((String)session.getAttribute("SortBy")).equals("")){	   
  		            sqlCmd.append(" " + ((String)session.getAttribute("SortBy")));	  		            
  		            //sqlCmd_sum += " " + ((String)session.getAttribute("SortBy"));	  		            
  		         }
  		    }else{
  		       sqlCmd.append(" order by a05.bank_code,a05.m_year,a05.m_month ");
  		       //sqlCmd_sum += " order by cd01.hsien_id";
  		    }	       			      
            //System.out.println("sqlCmd="+sqlCmd);
            //讀取報表欄位名稱===================================================================================        	
        	Properties p = new Properties();
			p.load(new FileInputStream(Utility.getProperties("schemaDir")+System.getProperty("file.separator")+"A05_"+bank_type+"_column.TXT"));			
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
                reportUtil.createCell( wb, row, ( short )3, "查詢年月", columnStyle );//95.12.04 add 查詢年月
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
            row.setHeightInPoints(33);//設定大類表頭高度
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
            //row.setHeightInPoints(63);//設定細項表頭高度
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
  			  			
  			//System.out.println("DS005W_Excel.sqlCmd="+sqlCmd);	   
  			System.out.println("ori_field="+ori_field); 
  			List dbData = null;
  			if(hasBankListALL.equals("false")){    			  	  
			  //dbData = DBManager.QueryDB(sqlCmd,ori_field+"m_yearmonth");
			  dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),sqlCmdList,ori_field+"m_yearmonth") ;
	        }else{//95.09.04 add 金融機構代號=ALL全部  			 
			  dbData = DBManager.QueryDB_SQLParam(sqlCmd_sum.toString(),null,ori_field+"m_yearmonth");  	          
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
            String amt_type="";
            String amt="";    
            String prtbank_code="";       
            for(i=0;i<dbData.size();i++){   
            	bean = (DataObject) dbData.get(i) ;
                acc_code = "";
                row = sheet.createRow( rowNo );  
                //System.out.println("rowNo="+rowNo);   
                row.setHeight((short) 0x120);    
                //System.out.println("bank_code="+(String) ((DataObject) dbData.get(i)).getValue("bank_code"));                
                /*95.12.05 fix 讓User可根據Excel做其他運用.單位代號.機構名稱repeat
                if(prtbank_code.equals((String) ((DataObject) dbData.get(i)).getValue("bank_code"))){
                   reportUtil.createCell( wb, row, ( short )columnIdx, " ", defaultStyle );//單位代號
                   columnIdx++;
                   reportUtil.createCell( wb, row, ( short )columnIdx, " ", defaultStyle );//機構名稱
                   columnIdx++;
                }else{*/
                   reportUtil.createCell( wb, row, ( short )columnIdx, (String) bean.getValue("bank_code"), defaultStyle );//單位代號
                   columnIdx++;
                   reportUtil.createCell( wb, row, ( short )columnIdx, (String)bean.getValue("bank_name"), defaultStyle );//機構名稱
                   columnIdx++;
                //}
                //95.12.04 add 查詢年月
                reportUtil.createCell( wb, row, ( short )columnIdx, bean.getValue("m_yearmonth").toString(), defaultStyle );//查詢年月
                columnIdx++;
                for(int cellIdx =4;cellIdx < (new Short(lastCellNum)).intValue();cellIdx++){                    
                     amt="";                                          
                     cell = acc_code_row.getCell((short)cellIdx);                    
                     acc_code = cell.getStringCellValue().toLowerCase();                     
                     //System.out.println("acc_code="+acc_code);                      
                     if(hasBankListALL.equals("true")){  		
                        amt = bean.getValue(acc_code).toString();
                     }else{                     
                        if(cell.getStringCellValue().length() == 6){               
                           if(bean.getValue("amt"+acc_code) != null){                                                           
                              amt = bean.getValue("amt"+acc_code).toString();                                                
                           }
                        }else{                  
                           amt = bean.getValue(acc_code).toString();
                        }
                     }                                                  
                     //System.out.println("amt="+amt);                     
                     if(acc_code.toUpperCase().equals("91060P")){//91060P   
                        reportUtil.createCell( wb, row, ( short )columnIdx, amt, rightStyle );               
                        //System.out.println(" is 91060P");                        
                     }else if(acc_code.startsWith("assumed") || acc_code.toUpperCase().endsWith("P")){//風險權數                        
                        amt_d = Double.parseDouble(amt); 
                        //System.out.println("amt_d="+amt_d);
                        reportUtil.createCell( wb, row, ( short )columnIdx, Double.parseDouble(amt)*100 + "%", rightStyle );                                    
 					 }else if(acc_code.toUpperCase().endsWith("N")){//9207XN中文名稱
 					    reportUtil.createCell( wb, row, ( short )columnIdx, amt, rightStyle );               
                     }else{//除以單位4捨五入
                        if(acc_code.equals("amt_assumed920710")
                        || acc_code.equals("amt_assumed920720")
                        || acc_code.equals("amt_assumed920730")
                        || acc_code.equals("amt_assumed920740")
                        || acc_code.equals("amt_assumed920750")){
                           //System.out.println("amt_d="+amt_d);
                           amt_d = amt_d * Double.parseDouble(amt);
                           //System.out.println("amt_d="+amt_d);
                           amt = String.valueOf(amt_d);
                           amt_d = 0.0; 
                        }
                        
                        reportUtil.createCell( wb, row, ( short )columnIdx, toFormatWithoutZero(Double.parseDouble(amt),Unit), rightStyle ); 
                        //amt = Utility.getRound(amt,Unit);                        
                        //amt = Utility.setCommaFormat(amt);
                        //reportUtil.createCell( wb, row, ( short )columnIdx, amt, rightStyle );               
                     }
                     
                     columnIdx ++;
                }
                columnIdx = 1;
                rowNo++;
                //prtbank_code = (String) ((DataObject) dbData.get(i)).getValue("bank_code");
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
               	     acc_code = (String)detail_column.get(j);
                     //System.out.println(acc_code +"="+Utility.ISOtoBig5((String)prop_column_name.get(acc_code)));                                           
                  	 if(acc_code.length() == 6 && !acc_code.startsWith("9207")){ //add 合併acc_code長度為=6的.欄位名稱                  	                      	    
               		    row = sheet.getRow(5);
               		    reportUtil.createCell( wb, row, ( short )columnIdx,Utility.ISOtoUTF8((String)prop_column_name.get(acc_code)), columnStyle );               
               		    sheet.addMergedRegion( new Region( ( short )5, ( short )columnIdx,
                                                           ( short )6,
                                               			   ( short )columnIdx) );                                              
                  		columnIdx ++;
               		}
               		if(acc_code.length() > 6 ){//合併acc_code長度為大於6的.欄位名稱               		                      	
               		   if(acc_code.startsWith("amt_assumed9207")){               		 
               		      row = sheet.getRow(6);
                  	      reportUtil.createCell( wb, row, ( short )columnIdx,Utility.ISOtoUTF8((String)prop_column_name.get(acc_code)), columnStyle );               
                  	      sheet.addMergedRegion( new Region( ( short )6, ( short )columnIdx,
                                                           ( short )7,
                                               			   ( short )columnIdx) );                                              
                  		  columnIdx ++;               		
               		    }else if(!acc_code.startsWith("9207")){               		      
                  	      row = sheet.getRow(6);
                  	      reportUtil.createCell( wb, row, ( short )columnIdx,Utility.ISOtoUTF8((String)prop_column_name.get(acc_code)), columnStyle );               
                  	      sheet.addMergedRegion( new Region( ( short )6, ( short )columnIdx,
                                                           ( short )7,
                                               			   ( short )columnIdx) );                                              
                  		  columnIdx ++;               		
               		    }
               		}
               		if(acc_code.length() == 6 && acc_code.startsWith("9207")) columnIdx ++;
               		acc_code = "";
                }//end of 細項表頭欄位                                              
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
<%!
private static String toFormatWithoutZero(double d,String Unit) {
		if(!Unit.equals("1")){
		   d = Double.parseDouble(Utility.getRound(String.valueOf(d),Unit));  
		}   
		d = toRound(d, 4);
		MessageFormat mf = new MessageFormat("{0,number,##,###.####}");
		Object[] obj = {new Double(d)};
		String str = mf.format(obj);		
		if(str.equals("0") || str.equals("0.0")) {
			str = "0";
		}
		return   str;
}

//四捨五入 num 到小數以下 p 位數
private static double toRound(double num, int p) {
	return (Math.floor(num * (Math.pow(10, p)) + 0.5)) / Math.pow(10, p);
}
%>