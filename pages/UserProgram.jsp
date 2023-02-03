<%
// 93.12.22 fix FX004W(農業行庫相關機構通訊錄)不顯示bank_list
// 93.12.23 add superuser可以看到bank_list
// 93.12.28 fix ZZ003W(程式_歸屬系統類別維護)只有super user才能使用
// 94.01.10 fix 若Muser_Data裡的e-mail為空白時,須強制輸入基本資料 by 2295
// 94.01.12 fix 農會.漁會不顯示FX004W(機構基本資料維護) by 2295
// 94.01.14 fix bank_type=共用中心8/地方主管機關B/農金局局內2/super user 可看到Bank_List by 2295
// 94.03.02 fix 為共用中心時,只能看到其共用中心的資料,不能看到其他共用中心的 by 2295
// 94.03.30 add 中央存保也可看到Bank_List by 2295
//          fix WMFileUpload/Download/Query/Edit 為共用中心時,且不是super user 時,只能看到其共用中心 by 2295
// 94.04.18 fix WMFileDeleteBatch若為共用中心時,直接進內頁 by 2295
// 94.05.18 fix 申報資料批次刪除僅開放給共用中心 by 2295
//          add 地方主管機關FX001W/FX002W/FX003W直接進去內頁 by 2295
//          fix 若登入者為地方主管機關,只能看到其地方主管機關,且不是super user  by 2295
// 94.06.20 fix 為中央存保時,可看到全部機構,且不是super user時 by 2295
// 94.08.15 fix 開放登入者為不為共用中心/農會/漁會/農漁會地方主管機關時,點選農會/漁會時可看到全部機關 by 2295
// 95.11.09 DS決策支援分析系統直接進入內頁 by 2295
// 95.11.27 FR0066WB直接進入內頁 by 2295
// 98.06.25 add DS開頭為彈性報表子目錄下 by 2295
// 98.08.12 RptFileUpload直接進入內頁 by 2295
//103.01.21 add BR開頭為基本資料表子目錄下 by 2295
//105.03.14 add 縣市政府增加財務報表子目錄 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@include file="common.jsp"%>
<%
    RequestDispatcher rd = null;
    String muser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");
    String muser_type = ( session.getAttribute("muser_type")==null ) ? "" : (String)session.getAttribute("muser_type");
    String bank_type = ( session.getAttribute("bank_type")==null ) ? "" : (String)session.getAttribute("bank_type");
    String tbank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");
    String muser_i_o = ( session.getAttribute("muser_i_o")==null ) ? "" : (String)session.getAttribute("muser_i_o");
    System.out.println("muser_type="+muser_type);
    if(session.getAttribute("muser_id") == null){
       System.out.println("UserProgram.login timeout");
	   rd = application.getRequestDispatcher( "/pages/reLogin.jsp?url=LoginError.jsp?timeout=true" );
       rd.forward(request,response);
    }
    String sqlCmd = "";
    List paramList = new ArrayList();
    sqlCmd = " select WTT02.bank_type,cdshareno.CMUSE_NAME,cdshareno.INPUT_ORDER,"
     	   + " WTT03_2.program_id,WTT03_1.PROGRAM_NAME,WTT03_1.url_id,"
	   	   + " WTT04.P_ADD,WTT04.P_DELETE,WTT04.P_UPDATE,"
	   	   + " WTT04.P_QUERY,WTT04.P_PRINT,WTT04.P_UPLOAD,"
	   	   + " WTT04.P_DOWNLOAD,WTT04.P_LOCK,WTT04.P_OTHER"
		   + " from WTT02"
		   + " LEFT JOIN cdshareno on WTT02.BANK_TYPE = cdshareno.CMUSE_ID and cdshareno.CMUSE_DIV='016'"
		   + " LEFT JOIN WTT03_2  on WTT03_2.PROGRAM_TYPE = WTT02.BANK_TYPE "
		   + " LEFT JOIN WTT04 on WTT03_2.PROGRAM_ID=WTT04.PROGRAM_ID "
		   + " LEFT JOIN WTT03_1 on WTT03_1.PROGRAM_ID=WTT04.PROGRAM_ID"
		   + " where WTT02.muser_id=?"
		   + " and WTT03_2.WEB_TYPE='1'"
		   + " and WTT04.muser_id = WTT02.muser_id"
		   + " order by cdshareno.INPUT_ORDER,WTT04.program_id";
    paramList.add(muser_id);
    List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");
    if(dbData != null){
       System.out.println("dbData.size()="+dbData.size());
    }
%>
<html><head>
<title>財政部金融局-金融機構網際網路權限分配子系統</title>
<link rel='stylesheet' href='js/ftie4style.css'>
<script src='js/ftiens4.js'></script>
<img src="images/L_menu_Bg1.gif" width="150" height="82">
<script>
<%
//94.01.10 fix 若Muser_Data裡的e-mail為空白時,須強制輸入基本資料
String UpdateMuser_Data = null;
if(session.getAttribute("UpdateMuser_Data") != null){
   UpdateMuser_Data = (String)session.getAttribute("UpdateMuser_Data");
}

if((UpdateMuser_Data != null && UpdateMuser_Data.equals("true")) ||
   ((dbData == null || dbData.size() == 0 || (dbData != null && dbData.size() != 0 && (!((String)((DataObject)dbData.get(0)).getValue("bank_type")).equals("Z"))))
&& (!muser_id.equals("superBOAF"))) )
{
   out.println("foldersTree = gFld('<font size=2>管理系統</font>', ' ');");
   out.println("insDoc(foldersTree, gLnk(0, '<font size=2>使用者基本資料維護(ZZ025W)</font><font size=1><br>&nbsp;</font>', 'ZZ025W.jsp?act=Edit&bank_type=Z'));");
   System.out.println("print initializeDocument");
   out.println("initializeDocument()");
   out.println("foldersTree.setState(0)");
}
if(( dbData == null || dbData.size() == 0 ) && (muser_id.equals("superBOAF"))){
   out.println("foldersTree = gFld('<font size=2>管理系統</font>', ' ');");
   out.println("insDoc(foldersTree, gLnk(0, '<font size=2>使用者帳號管理(ZZBOAF)</font><font size=1><br>&nbsp;</font>', 'ZZBOAF.jsp?act=List'));");
   System.out.println("print initializeDocument");
   out.println("initializeDocument()");
   out.println("foldersTree.setState(0)");
}

if( ( UpdateMuser_Data == null || !UpdateMuser_Data.equals("true")) && (dbData != null && dbData.size() != 0)){
   String Program_Type="";
   String program_id="";
   String program_id_pre="";//98.06.25 add 
   String program_name="";
   String program_url="";
   String program_type_name="";
   Properties permission = null;

   for(int i=0;i<dbData.size();i++){
       //System.out.print("i="+i+":");
       //System.out.print("Program_Type="+Program_Type);
       //System.out.print("bank_type="+(String)((DataObject)dbData.get(i)).getValue("bank_type")+":");
       //System.out.println("program_id="+program_id);

       program_id=(String)((DataObject)dbData.get(i)).getValue("program_id");
       program_name=(String)((DataObject)dbData.get(i)).getValue("program_name");

       program_url=(String)((DataObject)dbData.get(i)).getValue("url_id");
       program_type_name=(String)((DataObject)dbData.get(i)).getValue("cmuse_name");
       //取得該program_id的權限==============================================================
       permission =  ( session.getAttribute(program_id)==null ) ? new Properties() : (Properties)session.getAttribute(program_id);
       //====================================================================================
       if(program_id.equals("ZZ003W") && (muser_type.equals("A") || muser_type.equals(" "))){
          continue;
       }

       if(Program_Type.equals((String)((DataObject)dbData.get(i)).getValue("bank_type"))){
          if(permission != null && permission.get("Q") != null && permission.get("Q").equals("Y")){          	
          	
             if((!((String)((DataObject)dbData.get(i)).getValue("bank_type")).equals("Z"))
             &&  ( bank_type.equals("8")/*共用中心*/ || bank_type.equals("3")/*中央存保*/ || bank_type.equals("B")/*地方主管機關*/
                  || muser_i_o.equals("I")/*局內*/ || bank_type.equals("1")/*全國農業金庫*/ || bank_type.equals("4")/*農業信用保證基金*/
                  || bank_type.equals("5")/*農民銀行*/ || bank_type.equals("9")/*合作金庫*/ || bank_type.equals("A")/*銀行*/
                  || muser_type.equals("S")/*super user*/)){
                 //94.05.18 申報資料批次刪除只有共用中心可使用=========================================
                 if(program_id.equals("WMFileDeleteBatch") && !Program_Type.equals("8")){
                    continue;
                 }            
                 
                 if(program_id.startsWith("DS") || program_id.startsWith("FR0066WB") || program_id.startsWith("RptFileUpload")  || program_id.startsWith("WMGenerateOperation") || program_id.startsWith("BR") || program_id.startsWith("FR") || program_id.startsWith("WR")){
                   //95.11.09 DS決策支援分析系統直接進入內頁//95.11.27 FR0066WB直接進入內頁//98.08.12 RptFileUpload直接進入內頁//103.01.20 BR開頭接進入內頁 //105.03.14 FR開頭直接進入內頁
       		       program_url = (program_url.indexOf("?") != -1)?program_url+"&":program_url+"?";
       		       if(program_id.startsWith("FR") || program_id.startsWith("WR")){
       		       	  if(program_id.equals("FR001WB") || program_id.equals("FR001WB_ALL") || program_id.equals("FR006W") || program_id.equals("FR023W")){
       		       	  	 program_url = program_url + "bank_type=B:"+tbank_no;//105.03.15 add
       		       	  }else{       		       	  		
       		       	     program_url = program_url + "bank_type=6";//105.03.14 add
       		       	  }
       		       }else{	
                      program_url = program_url + "DS_bank_type="+(String)((DataObject)dbData.get(i)).getValue("bank_type");
                   }
                   //System.out.println("program_url="+program_url);                   
       		     }else if(!program_id.equals("FX004W") || !program_id.equals("WMFileEdit") || !program_id.equals("WMFileUpload") || !program_id.equals("WMFileQuery") || !program_id.equals("WMFileDownload")){
                     program_url = "Bank_List.jsp?bank_type="+((String)((DataObject)dbData.get(i)).getValue("bank_type"))+"&link_name="+program_id;
                     //System.out.println("program_url="+program_url);
                 }
                 //94.03.02 fix為共用中心時,只能看到其共用中心,且不是super user 時===============================================================
                 //94.05.18 fix為地方主管機關時,只能看到其地方主管機關,且不是super user 時===============================================================
                 if(program_id.equals("WMFileDownload") && (Program_Type.equals("8") || Program_Type.equals("B")) && !muser_type.equals("S")/*super user*/){
       		        program_url = "WMFileDownload.jsp?act=new&tbank_no="+tbank_no+"&bank_type="+((String)((DataObject)dbData.get(i)).getValue("bank_type"));
       		     }
       		     if(program_id.equals("WMFileEdit") && (Program_Type.equals("8") || Program_Type.equals("B")) && !muser_type.equals("S")/*super user*/){
       		        program_url = "WMFileEdit.jsp?act=List&tbank_no="+tbank_no+"&bank_type="+((String)((DataObject)dbData.get(i)).getValue("bank_type"));
       		     }
       		     if(program_id.equals("WMFileQuery") && (Program_Type.equals("8") || Program_Type.equals("B")) && !muser_type.equals("S")/*super user*/){
       		        program_url = "WMFileQuery.jsp?act=new&tbank_no="+tbank_no+"&bank_type="+((String)((DataObject)dbData.get(i)).getValue("bank_type"));
       		     }
       		     if(program_id.equals("WMFileUpload") && (Program_Type.equals("8") || Program_Type.equals("B")) && !muser_type.equals("S")/*super user*/){
       		        program_url = "WMFileUpload.jsp?act=new&tbank_no="+tbank_no+"&bank_type="+((String)((DataObject)dbData.get(i)).getValue("bank_type"));
       		     }
       		     //94.06.20 fix為中央存保時,可看到全部機構,且不是super user時
       		     if(program_id.equals("WMFileDownloadBatch") && (Program_Type.equals("8") || Program_Type.equals("B") 
       		     || Program_Type.equals("3") || Program_Type.equals("2") || Program_Type.equals("9") || Program_Type.equals("A") || Program_Type.equals("1")/*全國農業金庫*/ 
       		     || Program_Type.equals("4")/*農業信用保證基金*/ || Program_Type.equals("5")/*農民銀行*/ || Program_Type.equals("9")/*合作金庫*/ || Program_Type.equals("A")/*銀行*/) 
       		     && !muser_type.equals("S")/*super user*/){       		     
       		        program_url = "WMFileDownloadBatch_Qry.jsp?firstStatus=true&tbank_no="+tbank_no+"&bank_type="+((String)((DataObject)dbData.get(i)).getValue("bank_type"));
       		     }
       		     if(program_id.equals("WMFileDeleteBatch") && (Program_Type.equals("8") || Program_Type.equals("B")) && !muser_type.equals("S")/*super user*/){
       		        program_url = "WMFileDeleteBatch_Qry.jsp?firstStatus=true&tbank_no="+tbank_no+"&bank_type="+((String)((DataObject)dbData.get(i)).getValue("bank_type"));
       		     }
       		     //94.05.18 add 若登入者為地方主管機關FX001W/FX002W/FX003W直接進去內頁=====================================
       		     if(program_id.equals("FX001W") && Program_Type.equals("B") && !bank_type.equals("2") && !muser_type.equals("S")/*super user*/){
       		        program_url = "FX001W.jsp?act=new&tbank_no="+tbank_no+"&bank_type="+((String)((DataObject)dbData.get(i)).getValue("bank_type"));
       		     }
       		     if(program_id.equals("FX002W") && Program_Type.equals("B") && !bank_type.equals("2") && !muser_type.equals("S")/*super user*/){
       		        program_url = "FX002W.jsp?act=List&tbank_no="+tbank_no+"&bank_type="+((String)((DataObject)dbData.get(i)).getValue("bank_type"));
       		     }
       		     if(program_id.equals("FX003W") && Program_Type.equals("B") && !bank_type.equals("2") && !muser_type.equals("S")/*super user*/){
       		        program_url = "FX003W.jsp?act=List&tbank_no="+tbank_no+"&bank_type="+((String)((DataObject)dbData.get(i)).getValue("bank_type"));
       		     }   
       		     //System.out.println("program_url="+program_url);    		    
             }else{
                 program_url = (program_url.indexOf("?") != -1)?program_url+"&":program_url+"?";
                 program_url = program_url + "bank_type="+(String)((DataObject)dbData.get(i)).getValue("bank_type");
             }

             if(program_id.equals("FX004W")){
                if(Program_Type.equals("6") || Program_Type.equals("7")){
                   continue;
                }
                if(muser_type.equals("S") || (!((String)((DataObject)dbData.get(i)).getValue("bank_type")).equals(bank_type))){
      		       program_url = program_url+"&bank_type="+((String)((DataObject)dbData.get(i)).getValue("bank_type"));
      		    }else{
      		       program_url = "FX004W.jsp?act=Edit&bank_no="+tbank_no;
      		    }
       		 }//end of FX004W
       		 //System.out.println("program_id="+program_id);
       		 //System.out.println("program_id_pre="+program_id_pre);
       		 if(program_id.startsWith("DS") && !program_id_pre.startsWith("DS")){//98.06.25 add DS開頭為彈性報表子目錄下		    
       		    out.println("aux1 = insFld(foldersTree, gFld('<font size=2>彈性報表</font>', ''));") ;       		    
       		 }
       		 if(program_id.startsWith("BR") && !program_id_pre.startsWith("BR")){//103.01.21 add BR開頭為基本資料報表子目錄下		    
       		    out.println("aux1 = insFld(foldersTree, gFld('<font size=2>基本資料表</font>', ''));") ;       		    
       		 }
       		 if(program_id.startsWith("FR") && !program_id_pre.startsWith("FR")){//105.03.14 add FR開頭為財務報表子目錄下		    
       		    out.println("aux1 = insFld(foldersTree, gFld('<font size=2>財務報表</font>', ''));") ;       		    
       		 }
       		 if(program_id.startsWith("WR") && !program_id_pre.startsWith("WR")){//105.03.23 add WR開頭為財務報表子目錄下		    
       		    out.println("aux1 = insFld(foldersTree, gFld('<font size=2>警示報表</font>', ''));") ;       		    
       		 }
       		 /*
       		 if(program_id.startsWith("TM") && !program_id_pre.startsWith("TM")){//105.9.5 add TM開頭為協助措施子目錄下		    
       		    out.println("aux1 = insFld(foldersTree, gFld('<font size=2>協助措施</font>', ''));") ;       		    
       		 }
       		 */
       		 if(program_id.startsWith("DS") || program_id.startsWith("BR") || program_id.startsWith("FR") || program_id.startsWith("WR")){//105.03.14 add       		    
       		    out.println("insDoc(aux1, gLnk(0, '<font size=2>"+program_name+"("+program_id+")</font><font size=1><br>&nbsp;</font>', '"+program_url+"'));");       		        		    
       		 }else{
			    out.println("insDoc(foldersTree, gLnk(0, '<font size=2>"+program_name+"("+program_id+")</font><font size=1><br>&nbsp;</font>', '"+program_url+"'));");
			 }
			 program_id_pre = program_id;//98.06.25 add
          }
       }else{
          //System.out.println("Program_Type !==");
          if(i != 0){
             //System.out.println("print initializeDocument");
             out.println("initializeDocument()");
			 out.println("foldersTree.setState(0)");
          }
          Program_Type = (String)((DataObject)dbData.get(i)).getValue("bank_type");
          out.println("foldersTree = gFld('<font size=2>"+program_type_name+"</font>', ' ');");
          if(permission != null && permission.get("Q") != null && permission.get("Q").equals("Y")){
             //94.05.18 申報資料批次刪除只有共用中心可使用=========================================
             if(program_id.equals("WMFileDeleteBatch") && !Program_Type.equals("8")){
                continue;
             }
             if((!((String)((DataObject)dbData.get(i)).getValue("bank_type")).equals("Z"))             
             &&  (bank_type.equals("8")/*共用中心*/ || bank_type.equals("3")/*中央存保*/ || bank_type.equals("B")/*地方主管機關*/ 
             || muser_i_o.equals("I")/*局內*/ || bank_type.equals("1")/*全國農業金庫*/ || bank_type.equals("4")/*農業信用保證基金*/ 
             || bank_type.equals("5")/*農民銀行*/ || bank_type.equals("9")/*合作金庫*/ || bank_type.equals("A")/*銀行*/
             || muser_type.equals("S")/*super user*/)){
                 if(program_id.startsWith("DS") || program_id.startsWith("FR0066WB") || program_id.startsWith("RptFileUpload") || program_id.startsWith("WMGenerateOperation") || program_id.startsWith("BR")  || program_id.startsWith("FR") || program_id.startsWith("WR")){
                    //95.11.09 DS決策支援分析系統直接進入內頁//95.11.27 FR0066WB直接進入內頁 //98.08.12 RptFileUpload直接進入內頁 //103.01.21 BR直接進入內頁 //105.03.14 FR直接進入內頁            
       		        program_url = (program_url.indexOf("?") != -1)?program_url+"&":program_url+"?";
       		      
                   if(program_id.startsWith("FR") || program_id.startsWith("WR")){
       		       	  if(program_id.equals("FR001WB") || program_id.equals("FR001WB_ALL") || program_id.equals("FR006W") || program_id.equals("FR023W")){
       		       	  	 program_url = program_url + "bank_type=B:"+tbank_no;//105.03.15 add
       		       	  }else{       		       	  		
       		       	     program_url = program_url + "bank_type=6";//105.03.14 add
       		       	  }
       		       }else{	
                      program_url = program_url + "DS_bank_type="+(String)((DataObject)dbData.get(i)).getValue("bank_type");
                   }
                   //System.out.println("program_url="+program_url);
       		     }else if(!program_id.equals("FX004W") || !program_id.equals("WMFileEdit") || !program_id.equals("WMFileUpload") || !program_id.equals("WMFileQuery") || !program_id.equals("WMFileDownload")){
                    program_url = "Bank_List.jsp?bank_type="+((String)((DataObject)dbData.get(i)).getValue("bank_type"))+"&link_name="+program_id;
                 }
                 //94.03.02 fix為共用中心時,只能看到其共用中心,且不是super user 時===============================================================
                 //94.05.18 fix為地方主管機關時,只能看到其地方主管機關,且不是super user 時===============================================================
                 if(program_id.equals("WMFileDownload") && (Program_Type.equals("8") || Program_Type.equals("B")) && !muser_type.equals("S")/*super user*/){
       		        program_url = "WMFileDownload.jsp?act=new&tbank_no="+tbank_no+"&bank_type="+((String)((DataObject)dbData.get(i)).getValue("bank_type"));
       		     }
       		     if(program_id.equals("WMFileEdit") && (Program_Type.equals("8") || Program_Type.equals("B")) && !muser_type.equals("S")/*super user*/){
       		        program_url = "WMFileEdit.jsp?act=List&tbank_no="+tbank_no+"&bank_type="+((String)((DataObject)dbData.get(i)).getValue("bank_type"));
       		     }
       		     if(program_id.equals("WMFileQuery") && (Program_Type.equals("8") || Program_Type.equals("B")) && !muser_type.equals("S")/*super user*/){
       		        program_url = "WMFileQuery.jsp?act=new&tbank_no="+tbank_no+"&bank_type="+((String)((DataObject)dbData.get(i)).getValue("bank_type"));
       		     }
       		     if(program_id.equals("WMFileUpload") && (Program_Type.equals("8") || Program_Type.equals("B")) && !muser_type.equals("S")/*super user*/){
       		        program_url = "WMFileUpload.jsp?act=new&tbank_no="+tbank_no+"&bank_type="+((String)((DataObject)dbData.get(i)).getValue("bank_type"));
       		     }
       		     //94.06.20 fix為中央存保時,可看到全部機構,且不是super user時
       		     if(program_id.equals("WMFileDownloadBatch") && (Program_Type.equals("8") || Program_Type.equals("B") || Program_Type.equals("3") 
       		     || Program_Type.equals("2") || Program_Type.equals("9") || Program_Type.equals("A") || Program_Type.equals("1")/*全國農業金庫*/ 
       		     || Program_Type.equals("4")/*農業信用保證基金*/ || Program_Type.equals("5")/*農民銀行*/ || Program_Type.equals("9")/*合作金庫*/ 
       		     || Program_Type.equals("A")/*地土銀行*/) && !muser_type.equals("S")/*super user*/){       		     
       		        program_url = "WMFileDownloadBatch_Qry.jsp?firstStatus=true&tbank_no="+tbank_no+"&bank_type="+((String)((DataObject)dbData.get(i)).getValue("bank_type"));
       		     }

       		     if(program_id.equals("WMFileDeleteBatch") && (Program_Type.equals("8") || Program_Type.equals("B")) && !muser_type.equals("S")/*super user*/){
       		        program_url = "WMFileDeleteBatch_Qry.jsp?firstStatus=true&tbank_no="+tbank_no+"&bank_type="+((String)((DataObject)dbData.get(i)).getValue("bank_type"));
       		     }
       		     //94.05.18 add 若登入者為地方主管機關FX001W/FX002W/FX003W直接進去內頁=====================================
       		     if(program_id.equals("FX001W") && Program_Type.equals("B") && !bank_type.equals("2") && !muser_type.equals("S")/*super user*/){
       		        program_url = "FX001W.jsp?act=new&tbank_no="+tbank_no+"&bank_type="+((String)((DataObject)dbData.get(i)).getValue("bank_type"));
       		     }
       		     if(program_id.equals("FX002W") && Program_Type.equals("B") && !bank_type.equals("2") && !muser_type.equals("S")/*super user*/){
       		        program_url = "FX002W.jsp?act=List&tbank_no="+tbank_no+"&bank_type="+((String)((DataObject)dbData.get(i)).getValue("bank_type"));
       		     }
       		     if(program_id.equals("FX003W") && Program_Type.equals("B") && !bank_type.equals("2") && !muser_type.equals("S")/*super user*/){
       		        program_url = "FX003W.jsp?act=List&tbank_no="+tbank_no+"&bank_type="+((String)((DataObject)dbData.get(i)).getValue("bank_type"));
       		     }
             }else{
                program_url = (program_url.indexOf("?") != -1)?program_url+"&":program_url+"?";
                program_url = program_url + "bank_type="+(String)((DataObject)dbData.get(i)).getValue("bank_type");

             }
             if(program_id.equals("FX004W")){
                if(Program_Type.equals("6") || Program_Type.equals("7")){
                   continue;
                }
                if(muser_type.equals("S") || (!((String)((DataObject)dbData.get(i)).getValue("bank_type")).equals(bank_type))){
      		       program_url = program_url+"&bank_type="+((String)((DataObject)dbData.get(i)).getValue("bank_type"));
      		    }else{
      		       program_url = "FX004W.jsp?act=Edit&bank_no="+tbank_no;
      		    }
       		 }//end of FX004W
       		 if(program_id.startsWith("DS") && !program_id_pre.startsWith("DS")){//98.06.25 add DS開頭為彈性報表子目錄下		    
       		    out.println("aux1 = insFld(foldersTree, gFld('<font size=2>彈性報表</font>', ''));");       		     
       		 }
			 if(program_id.startsWith("BR") && !program_id_pre.startsWith("BR")){//103.01.21 add BR開頭為基本資料彈性報表子目錄下		    
       		    out.println("aux1 = insFld(foldersTree, gFld('<font size=2>基本資料表</font>', ''));");       		     
       		 }
       		 if(program_id.startsWith("FR") && !program_id_pre.startsWith("FR")){//105.03.14 add FR開頭為財務報表子目錄下		    
       		    out.println("aux1 = insFld(foldersTree, gFld('<font size=2>財務報表</font>', ''));");       		     
       		 }
       		 if(program_id.startsWith("WR") && !program_id_pre.startsWith("WR")){//105.03.23 add WR開頭為財務報表子目錄下		    
       		    out.println("aux1 = insFld(foldersTree, gFld('<font size=2>警示報表</font>', ''));") ;       		    
       		 }
       		 /*
       		 if(program_id.startsWith("TM") && !program_id_pre.startsWith("TM")){//105.09.05 add TM開頭為協助措施子目錄下		    
       		    out.println("aux1 = insFld(foldersTree, gFld('<font size=2>協助措施</font>', ''));") ;       		    
       		 }
       		 */
       		 if(program_id.startsWith("DS") || program_id.startsWith("BR") || program_id.startsWith("FR") || program_id.startsWith("WR")){ //105.03.23 add  	    
       		    out.println("insDoc(aux1, gLnk(0, '<font size=2>"+program_name+"("+program_id+")</font><font size=1><br>&nbsp;</font>', '"+program_url+"'));");       		    
       		 }else{
                out.println("insDoc(foldersTree, gLnk(0, '<font size=2>"+program_name+"("+program_id+")</font><font size=1><br>&nbsp;</font>', '"+program_url+"'));");
             }
             program_id_pre = program_id;//98.06.25
          }
       }
   }//end of for
   out.println("initializeDocument();");
   out.println("foldersTree.setState(0);");
}//end of dbData.size != 0
%>
drawlayer();
</script></head>
<body bgcolor="#6FBEBB" leftmargin="0" topmargin="0">
</body>
</html>
