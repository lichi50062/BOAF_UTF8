<%
//93.12.22 加入A02~M07之下載功能 2354
//93.12.23 fix 會多一列空白列 by 2295
//93.12.24 fix 最後一行多出空白字串 by 2295
//93.03.03 fix M01,M02產生之09312-->093MM及09300-->093YY by EGG
//95.04.10 add A06 by 2295
//95.05.26 add A99 by 2295
//96.07.11 add A08 by 2295
//97.01.02 add A09 by 2295
//97.06.13 add A10 by 2295
//99.09.27 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
//104.01.09 add A12 by 2968
//104.02.26 fix A10報表 by 2968
//104.10.12 add A13 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DownLoad" %>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="java.util.*" %>
<%@ page import="java.io.*" %>


<%
		Map dataMap =Utility.saveSearchParameter(request);
        String M_YEAR = Utility.getTrimString(dataMap.get("M_YEAR"));
        String M_MONTH = Utility.getTrimString(dataMap.get("M_MONTH"));
        String Report_no = Utility.getTrimString(dataMap.get("Report_no"));
        String bank_code = Utility.getTrimString(dataMap.get("bank_code"));

		
		System.out.println("save filename="+DownLoad.getFileName(Report_no, bank_code, M_YEAR, M_MONTH));
		System.out.println("DLId.substring(DLId.length()-1).equals(\"A\")="+Report_no.substring(0,1));

		StringBuffer sqlCmd = new StringBuffer();
	    List paramList = new ArrayList();

		//response.setContentType("multipart/form-data;charset=8859_1;Content-Transfer-Encoding=8bit;");
		response.setContentType("multipart/form-data;charset=BIG5;Content-Transfer-Encoding=8bit;");
	    response.setHeader("Content-Disposition","filename=" +
	    			DownLoad.getFileName(Report_no, bank_code, M_YEAR, M_MONTH) +
	    			";size=100;charset=8859_1;Content-Transfer-Encoding=8bit;");

		//response.setContentType("www/unknown");
  		//response.setHeader("Content-Disposition", "filename=test1.txt;");

 	    //先清buffer================================
 	    out.clear();
		out.clearBuffer();
		paramList.add(M_YEAR);
 	   	paramList.add(M_MONTH);
 	    paramList.add(bank_code);
		//============modify by 2354 12.22 begin ---
 	    if(Report_no.equals("A01") || Report_no.equals("A02") || Report_no.equals("A03") || Report_no.equals("A04") || Report_no.equals("A05") || Report_no.equals("A06") || Report_no.equals("A99") || Report_no.equals("A08") || Report_no.equals("A09") || Report_no.equals("A10")|| Report_no.equals("A12") || Report_no.equals("A13")){
 	    	List data = null; 	    	
 	    	if(Report_no.equals("A08")){
 	    	   sqlCmd.append("select * from a08 where m_year=? and m_month=? and bank_code=?"); 	    	   
 	    	   data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,warnaccount_cnt,limitaccount_cnt,erroraccount_cnt,otheraccount_cnt,depositaccount_tcnt");
 	    	}else if(Report_no.equals("A09")){
 	    	   sqlCmd.append("select * from a09 where m_year=? and m_month=? and bank_code=?");
 	    	   data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,over_cnt,over_amt,push_over_amt,totalamt,push_totalamt,over_total_rate");
 	    	}else if(Report_no.equals("A10")){
 	    	   sqlCmd.append("select * from a10 where m_year=? and m_month=? and bank_code=?");
 	    	   data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,loan1_amt,loan2_amt,loan3_amt,loan4_amt,invest1_amt,invest2_amt,invest3_amt,invest4_amt,other1_amt,other2_amt,other3_amt,other4_amt"
														 	    			  +",loan1_baddebt,loan2_baddebt,loan3_baddebt,loan4_baddebt,build1_baddebt,build2_baddebt,build3_baddebt,build4_baddebt"
																		      +",baddebt_flag,baddebt_noenough,baddebt_delay,baddebt_104,baddebt_105,baddebt_106,baddebt_107,baddebt_108");
 	    	}else if(Report_no.equals("A12")){
  	    	   sqlCmd.append("select * from a12 where m_year=? and m_month=? and bank_code=?");
  	    	   data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,baddebt_amt,loss_amt,profit_amt");
 	    	}else{
 	    	   sqlCmd.append("select * from "+Report_no+" where m_year=? and m_month=? and bank_code=? order by acc_code ");
 	    	   data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,amt,amt_3month,amt_6month,amt_1year,amt_2year,amt_over2year,amt_total");
    		}

			//===============================================
 	    	int i=0;
 	    	if(data.size() != 0){
 	    	      while(i < data.size()){
 	    	      	   out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("m_year")).toString(), "L", "0", 3));
				   	   out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("m_month")).toString(), "L", "0", 2));
				   	   out.print(DownLoad.fillStuff((String)((DataObject)data.get(i)).getValue("bank_code"), "R", "0", 7));
				   	   if(Report_no.equals("A08")){//96.07.11 add A08
				   	       //warnaccount_cnt,limitaccount_cnt,erroraccount_cnt,otheraccount_cnt,depositaccount_tcnt
				   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("warnaccount_cnt")).toString(), "L", "0", 0, 14));
				   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("limitaccount_cnt")).toString(), "L", "0", 0, 14));
				   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("erroraccount_cnt")).toString(), "L", "0", 0, 14));
				   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("otheraccount_cnt")).toString(), "L", "0", 0, 14));
				   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("depositaccount_tcnt")).toString(), "L", "0", 0, 14));
				       }else if(Report_no.equals("A09")){//97.01.02 add A09
				           //over_cnt,over_amt,push_over_amt,totalamt,push_totalamt,over_total_rate
				   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("over_cnt")).toString(), "L", "0", 0, 14));
				   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("over_amt")).toString(), "L", "0", 0, 14));
				   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("push_over_amt")).toString(), "L", "0", 0, 14));
				   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("totalamt")).toString(), "L", "0", 0, 14));
				   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("push_totalamt")).toString(), "L", "0", 0, 14));
				   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("over_total_rate")).toString(), "L", "0", 0, 14));
				   	   }else if(Report_no.equals("A10")){//97.06.13 add A10
				   		   String yymm = M_YEAR+M_MONTH;
				   		   if(Integer.parseInt(yymm)>=10403){
							   out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("loan1_amt")).toString(), "L", "0", 0, 14));
					   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("loan2_amt")).toString(), "L", "0", 0, 14));
					   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("loan3_amt")).toString(), "L", "0", 0, 14));
					   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("loan4_amt")).toString(), "L", "0", 0, 14));
					   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("invest1_amt")).toString(), "L", "0", 0, 14));
					   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("invest2_amt")).toString(), "L", "0", 0, 14));
					   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("invest3_amt")).toString(), "L", "0", 0, 14));
					   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("invest4_amt")).toString(), "L", "0", 0, 14));
					   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("other1_amt")).toString(), "L", "0", 0, 14));
					   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("other2_amt")).toString(), "L", "0", 0, 14));
					   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("other3_amt")).toString(), "L", "0", 0, 14));
					   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("other4_amt")).toString(), "L", "0", 0, 14));
					   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("loan1_baddebt")).toString(), "L", "0", 0, 14));
					   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("loan2_baddebt")).toString(), "L", "0", 0, 14));
					   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("loan3_baddebt")).toString(), "L", "0", 0, 14));
					   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("loan4_baddebt")).toString(), "L", "0", 0, 14));
					   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("build1_baddebt")).toString(), "L", "0", 0, 14));
					   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("build2_baddebt")).toString(), "L", "0", 0, 14));
					   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("build3_baddebt")).toString(), "L", "0", 0, 14));
					   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("build4_baddebt")).toString(), "L", "0", 0, 14));
					   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("baddebt_flag")).toString(), "L", "0", 0, 14));
					   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("baddebt_noenough")).toString(), "L", "0", 0, 14));
					   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("baddebt_delay")).toString(), "L", "0", 0, 14));
					   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("baddebt_104")).toString(), "L", "0", 0, 14));
					   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("baddebt_105")).toString(), "L", "0", 0, 14));
					   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("baddebt_106")).toString(), "L", "0", 0, 14));
					   		   out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("baddebt_107")).toString(), "L", "0", 0, 14));
					   		   out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("baddebt_108")).toString(), "L", "0", 0, 14));
				   		   }else{
					   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("loan2_amt")).toString(), "L", "0", 0, 14));
					   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("loan3_amt")).toString(), "L", "0", 0, 14));
					   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("loan4_amt")).toString(), "L", "0", 0, 14));
					   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("invest2_amt")).toString(), "L", "0", 0, 14));
					   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("invest3_amt")).toString(), "L", "0", 0, 14));
					   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("invest4_amt")).toString(), "L", "0", 0, 14));
					   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("other2_amt")).toString(), "L", "0", 0, 14));
					   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("other3_amt")).toString(), "L", "0", 0, 14));
					   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("other4_amt")).toString(), "L", "0", 0, 14));
				   		   }
				   	   }else if(Report_no.equals("A12")){//104.01.09 add A12
				   		   //baddebt_amt,loss_amt,profit_amt
				   		   out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("baddebt_amt")).toString(), "L", "0", 0, 14));
				   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("loss_amt")).toString(), "L", "0", 0, 14));
				   	       out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("profit_amt")).toString(), "L", "0", 0, 14));
				       }else{
				   	       out.print((String)((DataObject)data.get(i)).getValue("acc_code"));
				   	       String tmpAcc_Code=(String)((DataObject)data.get(i)).getValue("acc_code");
				   	       if(!tmpAcc_Code.substring(tmpAcc_Code.length()-1).equals("N")){
				   	          if(Report_no.equals("A06")){//95.04.10 add by 2295
				   	             out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("amt_3month")).toString(), "L", "0", 0, 14));
				   	             out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("amt_6month")).toString(), "L", "0", 0, 14));
				   	             out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("amt_1year")).toString(), "L", "0", 0, 14));
				   	             out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("amt_2year")).toString(), "L", "0", 0, 14));
				   	             out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("amt_over2year")).toString(), "L", "0", 0, 14));
				   	             out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("amt_total")).toString(), "L", "0", 0, 14));
				   	          }else{
				   	             out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("amt")).toString(), "L", "0", 0, 14));
				   	          }
				   	       }else{
				   	        out.print(DownLoad.fillStuff((String)((DataObject)data.get(i)).getValue("amt_name"), "R", " ", 30));
				   	       }
				   	   }
				   	   //System.out.println("i="+i);
				   	   //System.out.println("data.size()="+data.size());
				   	   if(i != (data.size()-1)) {
				   	      out.print("\n");
				   	   }
				   	   i++;
				   }
 	    	}
 	    //============modify by 2354 12.22 end ---
        //============add by 2354 12.22 begin ---
 	    }else if(Report_no.equals("M01")){
			sqlCmd.append(" select 	c.guarantee_item_no,a.guarantee_item_name,");
			         //"          substr(c.data_range,1,3) || decode(substr(c.data_range,4,5),'MM',LPAD(c.m_month,2,'0'),'YY','00',substr(c.data_range,4,5)) as \"data_range\", " +
			         //Modify by egg to fix 00-->YY,12-->MM
			sqlCmd.append("          substr(c.data_range,1,3) || decode(substr(c.data_range,4,5),'MM','MM','YY','YY',substr(c.data_range,4,5)) as \"data_range\", ");
			sqlCmd.append("        	c.guarantee_cnt,c.loan_amt,c.guarantee_amt,c.loan_bal,c.guarantee_bal,c.over_notpush_cnt, ");
			sqlCmd.append("        	c.over_notpush_bal,c.over_okpush_cnt,c.over_okpush_bal,c.over_notpush_bal, ");
			sqlCmd.append("        	c.repay_tot_cnt,c.repay_tot_amt,c.repay_bal_cnt,c.repay_bal_amt ");
			sqlCmd.append(" from 	m00_guarantee_item a,m00_data_range_item b,m01 c ");
			sqlCmd.append(" where 	c.guarantee_item_no         = a.guarantee_item_no ");
			sqlCmd.append(" and 		substr(c.data_range,4,2)= b.data_range ");
			sqlCmd.append(" and 		b.Report_no             = 'M01' ");
			sqlCmd.append(" and 		c.m_year                = ?");
			sqlCmd.append(" and 		c.m_month               = ?");
			sqlCmd.append(" order by a.input_order,b.input_order ");
 	    	List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"guarantee_cnt,loan_amt,guarantee_amt,loan_bal,guarantee_bal,over_notpush_cnt,over_notpush_bal,over_okpush_cnt,over_okpush_bal,repay_tot_cnt,repay_tot_amt,repay_bal_cnt,repay_bal_amt");
			//===============================================
 	    	int i=0;
 	    	System.out.println("data.size() ="+data.size() );
 	    	if(data.size() != 0){
 	    		while(i < data.size()){
					out.print(DownLoad.fillStuff((String)((DataObject)data.get(i)).getValue("guarantee_item_no"), "R", " ", 1));
					out.print(DownLoad.fillStuff((String)((DataObject)data.get(i)).getValue("data_range"), "R", " ", 5));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_cnt")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("loan_amt")).toString(), "L", "0", 0, 14));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_amt")).toString(), "L", "0", 0, 14));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("loan_bal")).toString(), "L", "0", 0, 14));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_bal")).toString(), "L", "0", 0, 14));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("over_notpush_cnt")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("over_notpush_bal")).toString(), "L", "0", 0, 11));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("over_okpush_cnt")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("over_okpush_bal")).toString(), "L", "0", 0, 11));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("repay_tot_cnt")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("repay_tot_amt")).toString(), "L", "0", 0, 11));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("repay_bal_cnt")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("repay_bal_amt")).toString(), "L", "0", 0, 11));

					if(i != (data.size()-1)) {out.print("\n");}
					i++;
				}
			}
 	    }else if(Report_no.equals("M02")){
			sqlCmd.append(" select c.loan_unit_no,a.loan_unit_name, ");
			         //"        substr(c.data_range,1,3) || decode(substr(c.data_range,4,5),'MM',LPAD(c.m_month,2,'0'),'YY','00',substr(c.data_range,4,5)) as \"data_range\", ");
			         //Modify by egg to fix 00-->YY,12-->MM
			sqlCmd.append("        substr(c.data_range,1,3) || decode(substr(c.data_range,4,5),'MM','MM','YY','YY',substr(c.data_range,4,5)) as \"data_range\", ");
			sqlCmd.append("        c.guarantee_cnt,c.loan_amt,c.guarantee_amt,c.loan_bal,c.guarantee_bal,c.over_notpush_cnt, ");
			sqlCmd.append("        c.over_notpush_bal,c.over_okpush_cnt,c.over_okpush_bal,c.repay_tot_cnt, ");
			sqlCmd.append("        c.repay_tot_amt,c.repay_bal_cnt,c.repay_bal_amt ");
			sqlCmd.append(" from   m00_loan_unit a,m00_data_range_item b,M02 c ");
			sqlCmd.append(" where  c.loan_unit_no=a.loan_unit_no ");
			sqlCmd.append("        and substr(c.data_range,4,2)=b.data_range ");
			sqlCmd.append("        and b.report_no = 'M02' ");
			sqlCmd.append("        and c.m_year = ?");
			sqlCmd.append("        and c.m_month = ?");
			sqlCmd.append("order by a.input_order,b.input_order");
 	    	List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"guarantee_cnt,loan_amt,guarantee_amt,loan_bal,guarantee_bal,over_notpush_cnt,over_notpush_bal,over_okpush_cnt,over_okpush_bal,repay_tot_cnt,repay_tot_amt,repay_bal_cnt,repay_bal_amt");
			//===============================================
 	    	int i=0;
 	    	System.out.println("data.size() ="+data.size() );
 	    	if(data.size() != 0){
 	    		while(i < data.size()){
					out.print(DownLoad.fillStuff((String)((DataObject)data.get(i)).getValue("loan_unit_no"), "R", " ", 1));
					out.print(DownLoad.fillStuff((String)((DataObject)data.get(i)).getValue("data_range"), "R", " ", 5));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_cnt")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("loan_amt")).toString(), "L", "0", 0, 14));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_amt")).toString(), "L", "0", 0, 14));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("loan_bal")).toString(), "L", "0", 0, 14));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_bal")).toString(), "L", "0", 0, 14));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("over_notpush_cnt")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("over_notpush_bal")).toString(), "L", "0", 0, 11));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("over_okpush_cnt")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("over_okpush_bal")).toString(), "L", "0", 0, 11));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("repay_tot_cnt")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("repay_tot_amt")).toString(), "L", "0", 0, 11));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("repay_bal_cnt")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("repay_bal_amt")).toString(), "L", "0", 0, 11));

					if(i != (data.size()-1)) {out.print("\n");}
					i++;
				}
			}

		}else if(Report_no.equals("M03")){
			sqlCmd.append(" select  M03.div_no,a.data_range,a.data_range_name,M03.guarantee_cnt_month, ");
			sqlCmd.append("        	M03.loan_amt_month, M03.guarantee_amt_month, ");
			sqlCmd.append("        	M03.guarantee_cnt_year,M03.loan_amt_year,M03.guarantee_amt_year, ");
			sqlCmd.append("        	M03.guarantee_bal_totacc,M03.guarantee_bal_totacc_over,M03.repay_bal_totacc ");
			sqlCmd.append(" from 	M03,m00_data_range_item a");
			sqlCmd.append(" where 	M03.div_no=a.data_range ");
			sqlCmd.append(" and		a.data_range_type = 'C' ");
			sqlCmd.append(" and     a.report_no ='M03' ");
			sqlCmd.append(" and 	M03.m_year= ? and M03.m_month= ?");
			sqlCmd.append(" order 	by a.input_order ");
			List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,guarantee_cnt_month,loan_amt_month,guarantee_amt_month,guarantee_cnt_year,loan_amt_year,guarantee_amt_year,guarantee_bal_totacc,guarantee_bal_totacc_over,repay_bal_totacc");
			//===============================================
 	    	int i=0;
 	    	if(data.size() != 0){
 	    		while(i < data.size()){
					out.print(DownLoad.fillStuff((String)((DataObject)data.get(i)).getValue("div_no"), "R", " ", 2));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_cnt_month")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("loan_amt_month")).toString(), "L", "0", 0, 14));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_amt_month")).toString(), "L", "0", 0, 14));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_cnt_year")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("loan_amt_year")).toString(), "L", "0", 0, 14));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_amt_year")).toString(), "L", "0", 0, 14));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_bal_totacc")).toString(), "L", "0", 0, 14));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_bal_totacc_over")).toString(), "L", "0", 0, 14));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("repay_bal_totacc")).toString(), "L", "0", 0, 14));

					out.print("\n");
					i++;
				}
 	    	}
			sqlCmd.append("select m_year,m_month,note_no as \"data_range\",note_amt_rate ");
			sqlCmd.append("from M03_note,m00_data_range_item ");
			sqlCmd.append("where note_no=data_range ");
			sqlCmd.append("  and data_range_type='S' ");
			sqlCmd.append("  and m_year= ?");
			sqlCmd.append("  and m_month= ?");
			sqlCmd.append("order by input_order ");
			data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,note_amt_rate");
			//===============================================
 	    	i=0;
 	    	if(data.size() != 0){
 	    		while(i < data.size()){
					String tmpNoteNO = (String)((DataObject)data.get(i)).getValue("data_range");
					out.print(DownLoad.fillStuff(tmpNoteNO, "R", " ", 4));
					if(tmpNoteNO.equals("NT1P")){
						out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("note_amt_rate")).toString(), "L", "0", 0, 7));
					}else{
						out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("note_amt_rate")).toString(), "L", "0", 0, 14));
					}
					if(i != (data.size()-1)) {out.print("\n");}
					i++;
				}
 	    	}

 	    }else if(Report_no.equals("M04")){
			sqlCmd.append(" select b.loan_use_no,a.loan_use_name,guarantee_no_month,guarantee_no_month_p, ");
			sqlCmd.append("        loan_amt_month,loan_amt_month_p,guarantee_amt_month,guarantee_amt_month_p,guarantee_no_year, ");
			sqlCmd.append("        guarantee_no_year_p,loan_amt_year,loan_amt_year_p,guarantee_amt_year, ");
			sqlCmd.append("        guarantee_amt_year_p,guarantee_no_totacc,guarantee_no_totacc_p, ");
			sqlCmd.append("        loan_amt_totacc,loan_amt_totacc_p,guarantee_amt_totacc,guarantee_amt_totacc_p ");
			sqlCmd.append(" from   m00_loan_use a,M04 b ");
			sqlCmd.append(" where  b.loan_use_no = a.loan_use_no ");
			sqlCmd.append("        and b.m_year = ?");
			sqlCmd.append("        and b.m_month = ?");
			sqlCmd.append("order by a.input_order");
 	    	List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,guarantee_no_month,guarantee_no_month_p,loan_amt_month,loan_amt_month_p,guarantee_amt_month,guarantee_amt_month_p,guarantee_no_year,guarantee_no_year_p,loan_amt_year,loan_amt_year_p,guarantee_amt_year,guarantee_amt_year_p,guarantee_no_totacc,guarantee_no_totacc_p,loan_amt_totacc,loan_amt_totacc_p,guarantee_amt_totacc,guarantee_amt_totacc_p");
			//===============================================
 	    	int i=0;
 	    	System.out.println("data.size() ="+data.size() );
 	    	if(data.size() != 0){
 	    		while(i < data.size()){
					out.print(DownLoad.fillStuff((String)((DataObject)data.get(i)).getValue("loan_use_no"), "R", " ", 1));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_no_month")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_no_month_p")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("loan_amt_month")).toString(), "L", "0", 0, 14));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("loan_amt_month_p")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_amt_month")).toString(), "L", "0", 0, 14));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_amt_month_p")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_no_year")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_no_year_p")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("loan_amt_year")).toString(), "L", "0", 0, 14));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("loan_amt_year_p")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_amt_year")).toString(), "L", "0", 0, 14));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_amt_year_p")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_no_totacc")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_no_totacc_p")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("loan_amt_totacc")).toString(), "L", "0", 0, 14));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("loan_amt_totacc_p")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_amt_totacc")).toString(), "L", "0", 0, 14));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_amt_totacc_p")).toString(), "L", "0", 0, 7));

					if(i != (data.size()-1)) {out.print("\n");}
					i++;
				}
			}

		}else if(Report_no.equals("M05")){
			sqlCmd.append(" select m05.loan_unit_no,m00_loan_unit.loan_unit_name, ");
			sqlCmd.append("        m00_data_range_item.data_range,m00_data_range_item.data_range_name, ");
			sqlCmd.append("        m05.period_no,m05.item_no, ");
			sqlCmd.append("        repay_cnt,repay_amt,run_notgood_cnt,run_notgood_amt,turn_out_cnt, ");
			sqlCmd.append("        turn_out_amt,diease_cnt,dieaserepay_amt,disaster_cnt,disaster_amt, ");
			sqlCmd.append("        corun_out_cnt,corun_out_amt,other_cnt,other_amt ");
			sqlCmd.append(" from m05,m00_loan_unit,m00_data_range_item ");
			sqlCmd.append(" where m05.loan_unit_no=m00_loan_unit.loan_unit_no ");
			sqlCmd.append("   and m05.period_no||m05.item_no=m00_data_range_item.data_range ");
			sqlCmd.append("   and m00_data_range_item.data_range_type='C' ");
			sqlCmd.append("   and m05.period_no || m05.item_no = m00_data_range_item.data_range ");
			sqlCmd.append("   and m05.m_year = ? and m05.m_month= ?");
			sqlCmd.append(" order by m00_loan_unit.output_order ,m00_data_range_item.output_order ");
			List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,repay_cnt,repay_amt,run_notgood_cnt,run_notgood_amt,turn_out_cnt,turn_out_amt,diease_cnt,dieaserepay_amt,disaster_cnt,disaster_amt,corun_out_cnt,corun_out_amt,other_cnt,other_amt");
			//===============================================
 	    	int i=0;
 	    	if(data.size() != 0){
 	    		while(i < data.size()){
					out.print(DownLoad.fillStuff((String)((DataObject)data.get(i)).getValue("loan_unit_no"), "R", " ", 1));
					out.print(DownLoad.fillStuff((String)((DataObject)data.get(i)).getValue("period_no"), "R", " ", 3));
					out.print(DownLoad.fillStuff((String)((DataObject)data.get(i)).getValue("item_no"), "R", " ", 2));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("repay_cnt")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("repay_amt")).toString(), "L", "0", 0, 14));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("run_notgood_cnt")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("run_notgood_amt")).toString(), "L", "0", 0, 14));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("turn_out_cnt")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("turn_out_amt")).toString(), "L", "0", 0, 14));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("diease_cnt")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("dieaserepay_amt")).toString(), "L", "0", 0, 14));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("disaster_cnt")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("disaster_amt")).toString(), "L", "0", 0, 14));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("corun_out_cnt")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("corun_out_amt")).toString(), "L", "0", 0, 14));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("other_cnt")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("other_amt")).toString(), "L", "0", 0, 14));
					out.print("\n");
					i++;
				}
 	    	}
			sqlCmd.append("select m_year,m_month,note_no,note_amt_rate ");
			sqlCmd.append("from m05_note,m00_data_range_item ");
			sqlCmd.append("where note_no=data_range ");
			sqlCmd.append("  and data_range_type='N' ");
			sqlCmd.append("  and m_year=?");
			sqlCmd.append("  and m_month=?");
			sqlCmd.append("order by input_order ");
			data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,note_amt_rate");
			//===============================================
 	    	i=0;
 	    	if(data.size() != 0){
 	    		while(i < data.size()){
					String tmpNoteNO = (String)((DataObject)data.get(i)).getValue("note_no");
					out.print(DownLoad.fillStuff(tmpNoteNO, "R", " ", 4));
					if(tmpNoteNO.equals("NT11")){
						out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("note_amt_rate")).toString(), "L", "0", 0, 7));
					}else{
						out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("note_amt_rate")).toString(), "L", "0", 0, 14));
					}
					out.print("\n");
					i++;
				}
 	    	}
			sqlCmd.append("select m_year,m_month,m05_totacc.loan_unit_no,loan_unit_name, ");
			sqlCmd.append("       fix_no,guarantee_no_totacc,guarantee_amt_totacc ");
			sqlCmd.append("from m05_totacc,m00_loan_unit ");
			sqlCmd.append("where m05_totacc.loan_unit_no=m00_loan_unit.loan_unit_no ");
			sqlCmd.append("  and m_year= ?");
			sqlCmd.append("  and m_month=?");
			sqlCmd.append("order by input_order ");
			data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,guarantee_no_totacc,guarantee_amt_totacc");
			//===============================================
 	    	i=0;
 	    	if(data.size() != 0){
 	    		while(i < data.size()){
					out.print(DownLoad.fillStuff((String)((DataObject)data.get(i)).getValue("loan_unit_no"), "R", " ", 1));
					out.print(DownLoad.fillStuff((String)((DataObject)data.get(i)).getValue("fix_no"), "R", " ", 3));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_no_totacc")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_amt_totacc")).toString(), "L", "0", 0, 14));
					if(i != (data.size()-1)) {out.print("\n");}
					i++;
				}
 	    	}

 	    }else if(Report_no.equals("M06") || Report_no.equals("M07")){
			sqlCmd.append(" select * from " + Report_no + ",m00_area where " + Report_no + ".area_no=m00_area.area_no and m_year=? and m_month=?");
			sqlCmd.append(" order by input_order ");
 	    	List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,guarantee_no_month,guarantee_amt_month,loan_amt_month,guarantee_no_year,guarantee_amt_year,loan_amt_year,guarantee_no_totacc,guarantee_amt_totacc,loan_amt_totacc,guarantee_bal_no,guarantee_bal_amt,guarantee_bal_p,loan_bal");
			//===============================================
 	    	int i=0;
 	    	System.out.println("data.size() ="+data.size() );
 	    	if(data.size() != 0){
 	    		while(i < data.size()){
					out.print(DownLoad.fillStuff((String)((DataObject)data.get(i)).getValue("area_no"), "R", " ", 1));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_no_month")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_amt_month")).toString(), "L", "0", 0, 14));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("loan_amt_month")).toString(), "L", "0", 0, 14));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_no_year")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_amt_year")).toString(), "L", "0", 0, 14));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("loan_amt_year")).toString(), "L", "0", 0, 14));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_no_totacc")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_amt_totacc")).toString(), "L", "0", 0, 14));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("loan_amt_totacc")).toString(), "L", "0", 0, 14));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_bal_no")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_bal_amt")).toString(), "L", "0", 0, 14));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("guarantee_bal_p")).toString(), "L", "0", 0, 7));
					out.print(DownLoad.fillStuff((((DataObject)data.get(i)).getValue("loan_bal")).toString(), "L", "0", 0, 14));
					if(i != (data.size()-1)) {out.print("\n");}
					i++;
				}
			}
        //============add by 2354 12.22 end ---
		}
		//93.12.24 fix by 2295
		if (out != null) {
            out.flush();
            out.close();
        }

%>        	