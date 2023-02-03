// 97.11.21 add 增加傳入日期 by 2295
// 97.11.28 fix 增加傳入日期為區間 by 2295
// 97.12.01 fix 若資料重覆時,為update by 2295
// 98.01.08 有失敗時.發送e-mail by 2295
// 99.06.07 fix 金庫代碼.在檢查局為018.農金局為0180012 by 2295 
//101.09.21 add exd08,檢查缺失摘要長度過長時(超過2000),分割為多筆,做base64存入DB by 2295
//105.10.18 add 檢查報告屬專案農貸查核缺失匯入(個案缺失及整体性缺失) by 2295
//106.03.24 add 檢查報告專案農貸,增加[核准貸放]解析字元 by 2295
//108.12.26 因檢查局改成https,調整取檔方式,TLS1.2不能使用java1.6版本(目前排程使用java 1.8/網頁使用1.6) by 2295
//109.02.17 fix 金資中心.在檢查局為600.農金局為6000000 by 2295     
package com.tradevan.util.xml;

import javax.mail.Message;
import javax.mail.Multipart;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import javax.net.ssl.TrustManagerFactory;
import java.security.cert.X509Certificate;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.ParserConfigurationException;
import org.xml.sax.SAXException;
import org.xml.sax.ErrorHandler;
import org.xml.sax.SAXParseException;
import org.w3c.dom.Document;
import org.w3c.dom.NodeList;
import org.w3c.dom.Node;
import java.io.*;
import java.net.URL;
import java.net.URLConnection;
import java.util.*;
import java.util.Date;
import java.security.KeyStore;
import java.sql.*;
import java.text.SimpleDateFormat;
import com.tradevan.util.Utility;
import com.tradevan.base64.*; 

import javax.net.ssl.HostnameVerifier;
import javax.net.ssl.HttpsURLConnection;
import javax.net.ssl.SSLContext;
import javax.net.ssl.SSLSession;
import javax.net.ssl.SSLSocketFactory;
import javax.net.ssl.TrustManager;
import javax.net.ssl.X509TrustManager;

public class parseDOM implements ErrorHandler {
  static Document document;  
  static List exd05 = new LinkedList();
  static List exd08 = new LinkedList();
  static List trans_exd08 = new LinkedList();
  static List trans_exDef = new LinkedList();
  static List ba01 = new LinkedList();
  static Properties recordData = new Properties();
  static Properties exDefData = new Properties();
  static Properties ba01Data = new Properties();
  static String val = "";
  
  static File logfile3;
  static FileOutputStream logos3=null;      
  static BufferedOutputStream logbos3 = null;
  static PrintStream logps3 = null;
  static Date nowlog = new Date();
  static SimpleDateFormat logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");        
  static SimpleDateFormat logfileformat = new SimpleDateFormat("yyyyMMddHHmmss");
  static Calendar logcalendar;
  static File logDir = null;
  static final String driverName = "oracle.jdbc.driver.OracleDriver";
  static String dbURL = "";//jdbc:oracle:thin:@172.20.5.22:1521:TBOAF"; 
  static String userName = ""; 
  static String userPwd = "";
  static Connection dbConn = null;
  static ResultSet rs = null;
  static Statement stat = null;
  static File logfile,logfile1;
  static FileOutputStream logos,logos1=null;        
  static BufferedOutputStream logbos,logbos1 = null;
  static PrintStream logps=null,logps1 = null;
  static DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
  static String sqlCmd="";
  static String dateStr = "";
  static String dateStr_end = "";  
  static String loan_date = "",memo = "";
  
  public static void main(String args[]) {      
    parseDOM a = new parseDOM();
    //a.parse_exd08("2008-12-01","");
    
    //System.out.println("args[0]="+args[0]);
    //a.parse_exd05("","");    
    //a.parse_exd08("","");
    //recordData = new Properties();
    //recordData.setProperty("content","五、辦理農業發展基金放款，對借款人有無重覆申貸或逾申貸餘額上限之查核結果，有未載明於相關徵信文件供參者，如：97.5.7貸放劉新川輔導漁業經營貸款7,680千元、97.6.28貸放游燊彥農家綜合貸款250千元及98.4.10貸放李春男農家綜合貸款300千元，核與「農業發展基金貸款作業規範」第12點規定不符。");
    //recordData.setProperty("content","四、內部管理面事項 (一)辦理內部稽核，經查有下列待改善事項： 1、辦理處分承受擔保品損失之轉銷作業，有未由稽核人員查明授信有無依據法令及業務規章辦理之情事，如：103.8.18日轉銷廖于翔案承受擔保品處分損失560千元，核與行政院農業委員會93.7.2農授金字第0935013489號函規定不符 2、有查核結果與事實不符或流於形式者： (1)對出納業務查核項目「營業廳報警系統有無按時測試每月至少2次，並做成紀錄？」之查核結果為「有」，惟實際有未落實辦理者，如：102.9.30本部一般內部稽核（102.2及102.6均未辦理警鈴測試）。 (2)對查核項目「銀行法32、33條有關授信明細表及33條之1有利害關係者資料表之建檔是否確實？」，查核結果為「確實填報建檔利害關係者資料表」，惟實際有未建檔者，如：103.9.30本部一般內部稽核。 (3)對專案農貸查核項目「貸款用途屬資本支出者所徵提之憑證，其應載事項是否完備？」，查核結果為「是」，惟有抽查實為週轉金支出案件者，如：103.9.30本部一般內部稽核，所抽查洪啟能、鄭翠華及田淑麗等案均為農家綜合貸款。 (4)對存款開戶查核項目「採用電腦印鑑比對系統者，其印鑑管理有無符合內部控制作業規定？」之查核結果勾選「有」，惟實際並未採用電腦印鑑比對系統者，如：103.9.30加祿分部一般內部稽核。");
    //recordData.setProperty("content","(二)辦理內部稽核及自行查核，經查有下列待改善事項： 1、有未將申報主管機關(含中央存保公司)各項資料之正確性，納入內部稽核查核項目，如：104.3.11本部一般內部稽核，核與行政院農業委員會農業金融局104.1.26農金二字第1045070062號函規定不符。 2、有查核之相關規定已廢除，未配合及時修正查核項目之情形，如：104.4.21西嶼分部一般內部稽核，查核項目「久未往來之存戶有否適時轉入靜止戶處理？靜止戶餘額是否相符？」，惟有關轉入靜止戶之規定，業經全面取消。 3、有查核結果與事實不符情事： (1)對查核項目「是否於規定期限內辦理貸款資金用途查驗？」之查核結果為「是」，惟所抽查農業發展基金貸款-綜合農家貸款案件，依規定免辦理查驗，實際亦未辦理查驗，如：104.3.11本部一般內部稽核。 (2)對查核項目「是否設置能拍攝營業廳全部、金庫室、保管箱區域進出口及重要處所之監視錄影系統？」，查核結果為「是」，惟本部攝錄範圍未涵蓋大出納收付區域，核與「金融機構安全維護管理辦法」第5條第6款規定不符，如：104.12.8本部出納自行查核。 4、辦理一般內部稽核，報告內容有未揭露自行查核辦理情形，如：104.3.11本部、104.4.13白沙分部及104.4.21西嶼分部，核與「農會漁會信用部內部控制及稽核制度實施辦法」第17條規定不符。");
    //recordData.setProperty("content","(四)辦理內部稽核及自行查核，經查有下列待改善事項： 1、對各營業單位之變現性資產查核，有僅於一般內部稽核時辦理，未適度增加其查核頻率者，如：103年度山腳、宏竹、外社、海湖、中福及五福等分部，請依行政院農業委員會96.6.27農授金字第0965013307號函規定，適度增加變現性資產之查核頻率。 2、辦理自行查核，有由徵授信經辦人員查核自身經辦之業務者，如：104.10.6本部ㄧ般自行查核，由徵授信簡麗卿辦理授信業務查核，核與「農會漁會信用部內部控制及稽核制度實施辦法」第21條第2項規定不符。 3、辦理一般內部稽核，有未涵蓋各項業務控制與內部管理者，如：104.7.3本部一般內部稽核，未將總務事項納入查核範圍，核與「農會漁會信用部內部控制及稽核制度實施辦法」第17條第2項規定不符。 4、有查核結果與事實不符者： (1)對查核項目「銀行法第33條之1有利害關係者資料表建檔是否確實？」，查核結果為「是｣，惟實際有負責人之「有利害關係者」資料未建檔，如：104.7.3本部一般內部稽核。 (2)對農業發展基金貸款查核項目「辦理貸款資金用途查驗是否依規定格式填寫查驗報告？」，查核結果為「是」，惟所抽樣案件均係「農家綜合貸款」，無須查驗，如：104.10.6本部ㄧ般自行查核。 (3)對電腦連線作業查核項目「…對其使用者代號之授予與交易權限之設定是否符合牽制原則？」，其查核結果為「是」，惟經查有櫃員1人同時擁有存款、放款、匯兌及代收等業務電腦作業權限，未符內部牽制原則，如：104.10.6山腳、新興及大竹等分部一般自行查核。");
    //recordData.setProperty("content","(二)辦理聯貸案件，經查有下列待改善事項：1、參貸聯貸案件，有先回覆主辦行同意參貸，再辦理徵信、授信審議委員會審議及核貸作業，作業程序顛倒，徵信與審核作業流於形式，如：105.8.9核准參貸陳新基一般放款(擔保)15000千元(全國農業金庫主辦，新北市土城區農會為管理行，聯貸總額754000千元，105.7.14先回復主辦行同意參貸，105.7.27始繕製徵信調查報告，105.8.5提授信審議委員會審議通過及105.8.9由總幹事核准。2、辦理土地融資，對開發案後續建築所需自籌資金，有未洽主辦行瞭解借戶整體資力及籌資能力是否足以支應，以確保興建進度能如期完成，上次檢查已發現有類似情形，如：105.8.9核准參貸陳新基一般放款(擔保)15000千元 (需自籌資金543036千元)。3、辦理擔保品鑑估，有逕以主辦行提供之擔保品訪價紀錄作為查定價格，未再自行蒐集資料查估，以評估鑑價之合理性，上次檢查已發現有類似情形，如：105.8.9核准參貸陳新基一般放款(擔保)15000千元及105.9.27核准貸放中華民國農會農漁會事業發展貸款100000千元(全國農業金庫擔任主辦行兼管理行，聯貸總額400000千元) 。4、參貸建築放款聯貸案件，有未洽請管理行定期提供借戶相關實地查勘及覆審資料供參，貸後管理欠落實，如：104.4.23核准參貸豐邑百貨(股)公司一般放款(擔保)14500千元(全國農業金庫主辦兼管理行，聯貸總額1772000千元)、103.9.10核准參貸吳俊億一般放款（擔保）14000千元(全國農業金庫及中華民國農會【兼管理行】共同主辦，聯貸總額121800千元)。");    
    //recordData.setProperty("content","一、辦理業務及內部稽核，經查有下列應改善事項，上次檢查已提列檢查意見，仍未就制度面辦理改善者： （一）辦理「有利害關係者」資料建檔，對負責人之親屬有未建檔者，如：詹永旺及詹黃蘭花（理事郭錦城配偶之父母），核與行政院農業委員會102.4.2農授金字第1025070350號函規定不符。 （二）辦理更正交易（EC），有重新認證處未由原放行主管核章確認者，如：本部104.6.25傳票編號＃100-1(放行主管為信用部主任)。 （三）對可查詢個人資料之電腦或設備，有未限制可攜式儲存媒體(如：USB磁碟)之使用並採取適當控管措施，客戶資料易遭外洩，不利個人資料保護者，如：聯徵中心信用查詢專用電腦。 （四）有調用非稽核人員辦理內部稽核，雖經總幹事核准，惟查核範圍涉及內部牽制與授信風險控管，已超逾法定得協查範圍，有妨礙內部稽核獨立功能之發揮者，如：104.6.30本部一般業務內部稽核，指派非稽核人員游秀美辦理政策性農業專案貸款查核，及指派非稽核人員許佳琪辦理電腦連線查核，渠等協查事項均已超逾法定「協助核算有關資料金（餘）額」之範圍，核與「農會漁會信用部內部控制及稽核制度實施辦法｣第15條規定不符者。");
    //recordData.setProperty("content","（二）企劃稽核部主任(胡貞如)有代理總幹事核准放款案件，未專任辦理稽核事務之情形，如：104.10.26代理總幹事核准貸放張淑惠農業天然災害低利貸款3000千元，核與「農會漁會信用部內部控制及稽核制度實施辦法」第6條規定不符。");
    /*   
    a.parseExDef(recordData);
    System.out.println("trans_exDef.size()="+trans_exDef.size());
    for(int i=0;i<trans_exDef.size();i++){
        recordData = (Properties)trans_exDef.get(i);
        System.out.println("record["+i+"]"); 
        System.out.println("ex_kind="+recordData.getProperty("ex_kind"));
        System.out.println("loan_name="+recordData.getProperty("loan_name"));
        System.out.println("loan_date="+recordData.getProperty("loan_date"));
        System.out.println("loan_item="+recordData.getProperty("loan_item"));
        System.out.println("loan_amt="+recordData.getProperty("loan_amt"));     
        System.out.println("memo="+recordData.getProperty("memo"));   
    }
    */
  
    if(args[0] != null){
       if(args[0].equals("exd05")){    	    
    	   if(args.length ==3  && (args[1] != null && args[2] != null)){
    		 a.parse_exd05(args[1],args[2]);  
    	   }else{
             a.parse_exd05("","");
    	   }  
       }    
       if(args[0].equals("exd08")){
    	   if(args.length ==3  && (args[1] != null && args[2] != null)){
    		 a.parse_exd08(args[1],args[2]);    
    	   }else{
             a.parse_exd08("","");
    	   }  
       }
    }
  
  } 
  
  public static String parse_exd05(String dateStr_src,String dateStr_src_end) {     
     String errMsg = "";
     try{
         exd05 = new LinkedList();       
         recordData = new Properties();
         dateStr = Utility.getDateFormat("yyyy-MM-dd");
         dateStr_end = dateStr;//起始.結束日期.皆為預設日期.當天
         if(!dateStr_src.equals("")){//起始日期
            dateStr = dateStr_src;
         }
         if(!dateStr_src_end.equals("")){//結束日期
            dateStr_end = dateStr_src_end;
         }
         logDir  = new File(Utility.getProperties("logDir"));
         if(!logDir.exists()){
            if(!Utility.mkdirs(Utility.getProperties("logDir"))){
               System.out.println("目錄新增失敗");
            }    
         }
         logfile3 = new File(logDir + System.getProperty("file.separator") + "parseDOM_exd05."+ logfileformat.format(nowlog));                       
         System.out.println("logfile filename="+logDir + System.getProperty("file.separator") +"parseDOM_exd05."+ logfileformat.format(nowlog));
         logos3 = new FileOutputStream(logfile3,true);                         
         logbos3 = new BufferedOutputStream(logos3);
         logps3 = new PrintStream(logbos3);             
     
         /*101.09.18--暫時拿掉,從網站上下載xml 
         //下載xml檔案.暫存至xlsDir之下   
         printMsg(logps3,"執行xml檔案下載"); 
         errMsg += "執行xml檔案下載<br>";
          
         logfile = new File(Utility.getProperties("xlsDir")+System.getProperty("file.separator")+"exd05.xml");
         logfile1 = new File(Utility.getProperties("xlsDir")+System.getProperty("file.separator")+"exd05_1.xml");
         logos = new FileOutputStream(logfile,false);                      
         logbos = new BufferedOutputStream(logos);
         logps = new PrintStream(logbos);       
         logos1 = new FileOutputStream(logfile1,false);                        
         logbos1 = new BufferedOutputStream(logos1);
         logps1= new PrintStream(logbos1);      
         //URL myURL=new URL("http://localhost:81/pages/exd05.xml");//測試
         
         //108.12.26
         X509TrustManager sunJSSEX509TrustManager;
         // 加載 Keytool 生成的証書文件
         char[] kphrase;
         String p = "changeit";
         kphrase = p.toCharArray();
         File file = new File("C:/jdk1.6.0_10/jre/lib/security/cacerts");
         System.out.println("Loading KeyStore " + file + "...");
         InputStream in_ca = new FileInputStream(file);
         KeyStore ks = KeyStore.getInstance(KeyStore.getDefaultType());
         ks.load(in_ca, kphrase);
         in_ca.close();
         for(int pi=0;pi < kphrase.length;pi++){
        	 kphrase[pi]=' ';
         }
         //System.out.println("test1="+ks.getCertificate("febsr11-RootCA")); //檢查局的root憑證     
         //System.out.println("test2="+ks.getCertificate("examweb"));//網站的憑證                  
         TrustManagerFactory tmf;
         tmf = TrustManagerFactory.getInstance(TrustManagerFactory.getDefaultAlgorithm());
         
         tmf.init(ks);
         TrustManager tms [] = tmf.getTrustManagers();
         System.out.println("parseDOM:java.version="+(System.getProperties()).getProperty("java.version"));  
         // 使用構造好的 TrustManager 訪問相應的 https 站點
         //SSLContext sslContext = SSLContext.getInstance("SSL", "SunJSSE");
         SSLContext sslContext = SSLContext.getInstance("TLSv1.2");
         sslContext.init(null, tms, new java.security.SecureRandom());
         SSLSocketFactory ssf = sslContext.getSocketFactory();
         
         URL myURL = new URL(Utility.getProperties("febWebSite")+"t=exd05&datef="+dateStr+"&datet="+dateStr_end);//正式
         //URL myURL=new URL(Utility.getProperties("febWebSite")+"t=exd05&datef=2005-01-01&datet=2008-10-29");
         System.out.println(Utility.getProperties("febWebSite")+"t=exd05&datef="+dateStr+"&datet="+dateStr_end);
         
         HttpsURLConnection httpsConn = (HttpsURLConnection) myURL.openConnection();
         
         //設定HttpURLConnection引數 [java]  view plain  copy
         //設定請求的方法為"POST",預設是GET    
         //httpsConn.setRequestMethod("POST");    
         
         //設定是否向httpUrlConnection輸出,因為這個是post請求,引數要放在    
         // http正文內,因此需要設為true, 預設情況下是false;    
         httpsConn.setDoOutput(true);               
         //設定是否從httpUrlConnection讀入,預設情況下是true;    
         httpsConn.setDoInput(true);               
         //Post 請求不能使用快取    
         httpsConn.setUseCaches(false);    
         
         //設定傳送的內容型別是可序列化的java物件    
         //如果不設此項,在傳送序列化物件時,當WEB服務預設的不是這種型別時可能拋java.io.EOFException)    
         httpsConn.setRequestProperty("Content-type", "application/x-java-serialized-object");    
         httpsConn.setRequestProperty("Keep-alive","true");
        
         //連線,從上述url.openConnection()至此的配置必須要在connect之前完成        
         
         httpsConn.setSSLSocketFactory(ssf);        
         //httpsConn.getConnectTimeout();
         //httpsConn.setConnectTimeout(50000);
         //httpsConn.setReadTimeout(50000);
        
         httpsConn.setUseCaches(false);// 取消緩存
         httpsConn.setRequestProperty("Content-type","application/x-www-form-urlencoded;charset=utf-8");
         //httpsConn.setRequestMethod("GET");
         
         //System.out.println("httpsConn.getReadTimeout()="+httpsConn.getReadTimeout());
         //System.out.println("httpsConn.getExpiration()="+httpsConn.getExpiration());
         
         BufferedReader infromURL=new BufferedReader(new InputStreamReader(httpsConn.getInputStream()));
         
         //108.12.26舊的取檔方式
         //URLConnection raoURL;
         //raoURL=myURL.openConnection();
         //raoURL.setDoInput(true);
         //raoURL.setDoOutput(true);
         //raoURL.setUseCaches(false);
         //BufferedReader infromURL=new BufferedReader(new InputStreamReader(raoURL.getInputStream()));
         //
         
         String raoInputString;     
         //下載後..第一個字元會出現亂碼.必須截取掉==========================================================
         logps.println("?<?xml version='1.0' encoding='utf-8' standalone='no'?>".trim());
         logps.flush();         
         int datai=0;
         boolean haveRecords = false;
         boolean records_multi = false;
         while((raoInputString=infromURL.readLine())!=null){              
              if( datai > 0){
                 logps.println(raoInputString);                                 
                 logps.flush();
              }
              datai++;
         }
         infromURL.close();
         if(logos != null) logos.close();
         if(logbos != null) logbos.close();
         if(logps != null) logps.close();   
         
         FileReader f = new FileReader(logfile);
         LineNumberReader in = new LineNumberReader(f);             
         String txtline=""; 
         datai=0;
         doLoop:
         while ((txtline = in.readLine()) != null) {
              if (!txtline.trim().equals("")){
                  if( datai > 0){
                    logps1.println(txtline);                                    
                    logps1.flush();                 
                  }else{
                    txtline = txtline.substring(1,txtline.length());
                    logps1.println(txtline);                                    
                    logps1.flush(); 
                  }
                  printMsg(logps3,txtline); 
                  if(txtline.trim().indexOf("<record>") != -1) haveRecords = true;
                  datai++;
              }
         }  
         in.close();
         f.close(); 
         if(logos1 != null) logos1.close();
         if(logbos1 != null) logbos1.close();
         if(logps1 != null) logps1.close();
         //=================================================================================
         printMsg(logps3,"xml檔案下載完成");          
         /*101.09.18--暫時拿掉,從網站上下載xml
         */
         
         boolean haveRecords = true;//101.09.18測試用
         boolean records_multi = false;//101.09.18測試用
         errMsg += "xml檔案下載完成<br>";
         if(haveRecords){//有資料時  
             //設定剖析的參數
             dbf.setIgnoringComments(true);
             dbf.setIgnoringElementContentWhitespace(true);
             dbf.setCoalescing(true);
             
             DocumentBuilder db = dbf.newDocumentBuilder();
             //讀入XML文件
             document = db.parse(new File(Utility.getProperties("xlsDir")+System.getProperty("file.separator")+"exd05_1.xml"));
             
             //取得根元素
             Node root = document.getDocumentElement();
             System.out.println("根元素:"+root.getNodeName());
             
             //取得機構資料
             dbURL = Utility.getProperties("BOAFDBURL");
             userName = Utility.getProperties("rptID");
             userPwd = Utility.getProperties("rptPwd");
             Class.forName(driverName); 
             dbConn = DriverManager.getConnection(dbURL, userName, userPwd); 
             System.out.println("Connection Successful!");
             stat = dbConn.createStatement(); 
             rs = stat.executeQuery("select bank_no,bank_name,bank_type,bank_kind,pbank_no from ba01");         
             while (rs.next()) {            
                //System.out.println("have data");
                ba01Data = new Properties();                
                ba01Data.setProperty("bank_no",rs.getString("bank_no"));
                ba01Data.setProperty("bank_type",rs.getString("bank_type"));
                ba01Data.setProperty("pbank_no",rs.getString("pbank_no"));
                ba01.add(ba01Data);          
             }       
             printMsg(logps3,"取得機構資料完成");
             errMsg += "取得機構資料完成<br>";
             //顯示指定元素的所有子節點
             NodeList tagNodes = document.getElementsByTagName("record");
             for(int i=0;i<tagNodes.getLength();i++){
                 System.out.println("record("+i+"):");               
                 NodeList childs = tagNodes.item(i).getChildNodes();
                 //顯示所有子節點 
                 for(int j=0;j<childs.getLength();j++){
                     if(childs.item(j).getNodeType() == Node.ELEMENT_NODE){
                        System.out.print(" +-"+childs.item(j).getNodeName());
                        try{
                            System.out.println("/"+childs.item(j).getFirstChild().getNodeValue());
                            val = childs.item(j).getFirstChild().getNodeValue();                               
                        }catch(Exception e){val = "";}; 
                        
                        //99.06.07 fix 金庫代碼.在檢查局為018.農金局為0180012                        
                        if(childs.item(j).getNodeName().equals("bank_no") && val.equals("018")){//金庫代碼:檢查局機構代號018.農金局為0180012
                           val = "0180012";//機構代號為018時,轉為0180012存入bank_no
                        } 	
                        
                        //109.02.17 fix 金資中心.在檢查局為600.農金局為6000000                        
                        if(childs.item(j).getNodeName().equals("bank_no") && val.equals("600")){//金資中心代碼:檢查局機構代號600.農金局為6000000
                           val = "6000000";//機構代號為600時,轉為6000000存入bank_no
                        }
                        recordData.setProperty(childs.item(j).getNodeName(),val);
                        
                        if(childs.item(j).getNodeName().equals("bank_no")){
                            ba01Loop:
                            for(int k=0;k<ba01.size();k++){
                                ba01Data = (Properties)ba01.get(k);                                
                                if(ba01Data.getProperty("bank_no").equals(val)){
                                   recordData.setProperty("bank_type",ba01Data.getProperty("bank_type")); 
                                   recordData.setProperty("pbank_no",ba01Data.getProperty("pbank_no"));
                                   break ba01Loop;
                                }
                            }
                        }
                        //System.out.println("set prop="+childs.item(j).getNodeName()+":"+val);  
                    }
                }
                exd05.add(recordData);
                recordData = new Properties();
             }
             
             System.out.println("exd05.size()="+exd05.size());
             printMsg(logps3,"exd05共下載"+exd05.size()+"筆資料");
             errMsg += "exd05共下載"+exd05.size()+"筆資料<br>";
             List updateSQL = new LinkedList();
             for(int i=0;i<exd05.size();i++){
                 recordData = (Properties)exd05.get(i);
                 System.out.println("record["+i+"]"); 
                 System.out.println("reportno="+recordData.getProperty("reportno"));
                 System.out.println("exam_id="+recordData.getProperty("exam_id"));
                 System.out.println("deptid="+recordData.getProperty("deptid"));
                 System.out.println("upd_user="+recordData.getProperty("upd_user"));
                 System.out.println("upd_date="+recordData.getProperty("upd_date"));             
                 System.out.println("bank_no="+recordData.getProperty("bank_no"));
                 System.out.println("bank_type="+recordData.getProperty("bank_type"));
                 System.out.println("pbank_no="+recordData.getProperty("pbank_no"));
                 System.out.println("base_date="+recordData.getProperty("base_date"));
                 //check資料有無重覆==============================================================================
                 records_multi = false;
                 rs = stat.executeQuery("select count(*) as datacount from exreportf where reportno='"+recordData.getProperty("reportno")+"'");
                 System.out.println("select count(*) as datacount from exreportf where reportno='"+recordData.getProperty("reportno")+"'");
                 while (rs.next()) {            
                    System.out.println("exd05 multi data");
                    if(rs.getInt("datacount") > 0){
                       records_multi = true;
                       printMsg(logps3,"reportno["+recordData.getProperty("reportno")+"]資料已存在");
                    }
                 } 
                 //==============================================================================================
                 //檢查報告檔
                 if(records_multi){//資料重覆時
                    sqlCmd = " update exreportf set "
                           + " report_in_date = To_date('"+recordData.getProperty("base_date")+"','mm/dd/yyyy'),"                       
                           + " report_en_date = To_date('"+recordData.getProperty("upd_date")+"','yyyy-mm-dd hh24-mi-ss'),"                 
                           + " base_date = To_date('"+recordData.getProperty("base_date")+"','mm/dd/yyyy'),"
                           + " disp_id = '"+recordData.getProperty("deptid")+"',"
                           + " originunt_id = '"+recordData.getProperty("exam_id")+"',ch_type='1',"
                           + " bank_type = '"+Utility.getTrimString(recordData.getProperty("bank_type"))+"',"
                           + " tbank_no = '"+Utility.getTrimString(recordData.getProperty("pbank_no"))+"',"
                           + " bank_no = '"+recordData.getProperty("bank_no")+"',"
                           + " user_name ='"+recordData.getProperty("upd_user")+"',"
                           + " update_date = To_date('"+recordData.getProperty("upd_date")+"','yyyy-mm-dd hh24-mi-ss')"
                           + " where reportno = '"+recordData.getProperty("reportno")+ "'";
                 }else{
                    sqlCmd = "insert into exreportf("
                           + "reportno,report_in_date,report_en_date,base_date," 
                           + "disp_id,originunt_id,ch_type,bank_type,tbank_no,bank_no, user_name,update_date"
                           + ")values("
                           + "'"+recordData.getProperty("reportno")+"',"
                           + "To_date('"+recordData.getProperty("base_date")+"','mm/dd/yyyy'),"                     
                           + "To_date('"+recordData.getProperty("upd_date")+"','yyyy-mm-dd hh24-mi-ss'),"                   
                           + "To_date('"+recordData.getProperty("base_date")+"','mm/dd/yyyy'),"
                           + "'"+recordData.getProperty("deptid")+"',"
                           + "'"+recordData.getProperty("exam_id")+"','1',"
                           + "'"+Utility.getTrimString(recordData.getProperty("bank_type"))+"',"
                           + "'"+Utility.getTrimString(recordData.getProperty("pbank_no"))+"',"
                           + "'"+recordData.getProperty("bank_no")+"',"
                           + "'"+recordData.getProperty("upd_user")+"',"
                           + "To_date('"+recordData.getProperty("upd_date")+"','yyyy-mm-dd hh24-mi-ss'))";
                 }
                 updateSQL.add(sqlCmd);             
             }
             
             if(updateSQL.size() >= 1){
                dbConn.setAutoCommit(false);        
                for(int idx = 0;idx < updateSQL.size();idx++){
                    stat.addBatch((String)updateSQL.get(idx));
                    System.out.println((String)updateSQL.get(idx));
                    printMsg(logps3,(String)updateSQL.get(idx));
                }
                int[] rowCount  = stat.executeBatch();
                System.out.println("rowCount="+rowCount.length);
                
                int i=0;
                boolean updateOK=true;
                while(i < rowCount.length){
                    if(rowCount[i] <= 0){
                        System.out.println("i="+i+":"+rowCount[i]+":sql="+(String)updateSQL.get(i));                    
                        updateOK = false;
                    }
                    printMsg(logps3,"i="+i+":updateOK="+rowCount[i]+":sql="+(String)updateSQL.get(i));
                    i++;
                }
                dbConn.commit();
                if(updateOK){
                   printMsg(logps3,"批次寫入資料庫成功");
                   errMsg += "批次寫入資料庫成功<br>";
                }else{
                   printMsg(logps3,"批次寫入資料庫失敗");
                   errMsg += "批次寫入資料庫失敗<br>";
                   send_Mail("","批次寫入資料庫失敗");
                }
             }
         }else{
            printMsg(logps3,"無資料可供下載");
            errMsg += "無資料可供下載<br>";
         }
     }catch(SAXException se){
        //剖析過程錯誤
        Exception e = se;
        if(se.getException() != null) e = se.getException();
        send_Mail("",e+e.getMessage());
        e.printStackTrace();        
     }catch(ParserConfigurationException pe){
        send_Mail("",pe+pe.getMessage());
        //剖析器設定錯誤
        pe.printStackTrace();       
     }catch(IOException ie){
        send_Mail("",ie+ie.getMessage());
        //檔案處理錯誤
        ie.printStackTrace();
     }catch(Exception e){
        System.out.println(e+e.getMessage());
        printMsg(logps3,e+e.getMessage());
        errMsg += e+e.getMessage()+"<br>";
        send_Mail("",e+e.getMessage());
     }finally{
        try{
            if (rs != null){
                rs.close();
                rs = null;//104.10.06 add
            }
                
            if (stat != null){
                stat.close();
                stat = null;//104.10.06 add
            }
            
            if (dbConn != null){//106.06.07 add
               if(!dbConn.isClosed()){//104.10.06        
                dbConn.close();
                dbConn = null;
               }
            }
        }catch(SQLException sqle){
               System.out.println(sqle+sqle.getMessage());
               printMsg(logps3,sqle+sqle.getMessage());
        }
     }
     return errMsg;
}           
  
  public static String parse_exd08(String dateStr_src,String dateStr_src_end) {     
     String errMsg = "";
     try{    
         exd08 = new LinkedList();   
         trans_exd08 = new LinkedList();
         recordData = new Properties();
         dateStr = Utility.getDateFormat("yyyy-MM-dd");
         dateStr_end = dateStr;//起始.結束日期.皆為預設日期.當天
         if(!dateStr_src.equals("")){//起始日期
            dateStr = dateStr_src;
         }
         if(!dateStr_src_end.equals("")){//結束日期
            dateStr_end = dateStr_src_end;
         }
         logDir  = new File(Utility.getProperties("logDir"));
         if(!logDir.exists()){
            if(!Utility.mkdirs(Utility.getProperties("logDir"))){
               System.out.println("目錄新增失敗");
            }    
         }
         logfile3 = new File(logDir + System.getProperty("file.separator") + "parseDOM_exd08."+ logfileformat.format(nowlog));                       
         System.out.println("logfile filename="+logDir + System.getProperty("file.separator") +"parseDOM_exd08."+ logfileformat.format(nowlog));
         logos3 = new FileOutputStream(logfile3,true);                         
         logbos3 = new BufferedOutputStream(logos3);
         logps3 = new PrintStream(logbos3); 
         
         //101.09.18--暫時拿掉,從網站上下載xml
         //下載xml檔案.暫存至xlsDir之下
         
         printMsg(logps3,"執行xml檔案下載"); 
         errMsg += "執行xml檔案下載<br>";
         logfile = new File(Utility.getProperties("xlsDir")+System.getProperty("file.separator")+"exd08.xml");       
         logos = new FileOutputStream(logfile,false);                      
         logbos = new BufferedOutputStream(logos);
         logps = new PrintStream(logbos);       
         //URL myURL=new URL("http://localhost:81/pages/exd08.xml");//測試         
         
         //108.12.26
         X509TrustManager sunJSSEX509TrustManager;
         // 加載 Keytool 生成的證書文件
         char[] kphrase;
         String p = "changeit";
         kphrase = p.toCharArray();
         File file = new File("C:/jdk1.6.0_10/jre/lib/security/cacerts");
         System.out.println("Loading KeyStore " + file + "...");
         InputStream in_ca = new FileInputStream(file);
         KeyStore ks = KeyStore.getInstance(KeyStore.getDefaultType());
         ks.load(in_ca, kphrase);
         in_ca.close();
         for(int pi=0;pi < kphrase.length;pi++){
        	 kphrase[pi]=' ';
         }
         //System.out.println("test1="+ks.getCertificate("febsr11-RootCA")); //檢查局的root憑證     
         //System.out.println("test2="+ks.getCertificate("examweb"));//網站的憑證                  
         TrustManagerFactory tmf;
         tmf = TrustManagerFactory.getInstance(TrustManagerFactory.getDefaultAlgorithm());
         
         tmf.init(ks);
         TrustManager tms [] = tmf.getTrustManagers();
         //使用構造好的 TrustManager 訪問相應的 https 站點
         //SSLContext sslContext = SSLContext.getInstance("SSL", "SunJSSE");
         SSLContext sslContext = SSLContext.getInstance("TLSv1.2");
         sslContext.init(null, tms, new java.security.SecureRandom());
         SSLSocketFactory ssf = sslContext.getSocketFactory();
         
         URL myURL = new URL(Utility.getProperties("febWebSite")+"t=exd08&datef="+dateStr+"&datet="+dateStr_end);//正式
         //URL myURL=new URL(Utility.getProperties("febWebSite")+"t=exd08&datef=2005-01-01&datet=2008-10-29");
         System.out.println(Utility.getProperties("febWebSite")+"t=exd08&datef="+dateStr+"&datet="+dateStr_end);
         
         HttpsURLConnection httpsConn = (HttpsURLConnection) myURL.openConnection();
         
         //設定HttpURLConnection引數 [java]  view plain  copy
         //設定請求的方法為"POST",預設是GET    
         //httpsConn.setRequestMethod("POST");    
         
         //設定是否向httpUrlConnection輸出,因為這個是post請求,引數要放在    
         // http正文內,因此需要設為true, 預設情況下是false;    
         httpsConn.setDoOutput(true);               
         //設定是否從httpUrlConnection讀入,預設情況下是true;    
         httpsConn.setDoInput(true);               
         //Post 請求不能使用快取    
         httpsConn.setUseCaches(false);    
         
         //設定傳送的內容型別是可序列化的java物件    
         //如果不設此項,在傳送序列化物件時,當WEB服務預設的不是這種型別時可能拋java.io.EOFException)    
         httpsConn.setRequestProperty("Content-type", "application/x-java-serialized-object");    
         httpsConn.setRequestProperty("Keep-alive","true");
        
         //連線,從上述url.openConnection()至此的配置必須要在connect之前完成
         
         httpsConn.setSSLSocketFactory(ssf);        
         //httpsConn.getConnectTimeout();
         //httpsConn.setConnectTimeout(50000);
         //httpsConn.setReadTimeout(50000);
        
         httpsConn.setRequestProperty("Content-type","application/x-www-form-urlencoded;charset=utf-8");
         //httpsConn.setRequestMethod("GET");
         
         //System.out.println("httpsConn.getReadTimeout()="+httpsConn.getReadTimeout());
         //System.out.println("httpsConn.getExpiration()="+httpsConn.getExpiration());
         
         BufferedReader infromURL=new BufferedReader(new InputStreamReader(httpsConn.getInputStream()));
         BufferedInputStream biStream = new BufferedInputStream(httpsConn.getInputStream()); 
         
         /*108.12.26舊的取檔方式
         URLConnection raoURL;
         raoURL=myURL.openConnection();
         raoURL.setDoInput(true);
         raoURL.setDoOutput(true);
         raoURL.setUseCaches(false);
         
         BufferedReader infromURL=new BufferedReader(new InputStreamReader(raoURL.getInputStream()));
         BufferedInputStream biStream = new BufferedInputStream(raoURL.getInputStream()); 
         */
                  
         String raoInputString;  
         
         //byte[] szData = new byte[raoURL.getInputStream().available()];
         byte[] szData = new byte[128];
         boolean haveRecords = false;
         boolean records_multi = false;
         //logps.write((new String(szData,"ISO8859_1")).getBytes("UTF8"));
              
         System.out.println("BufferedInputStream begin======================================");
         //logps.write((new String(szData,"ISO8859_1")).getBytes("UTF8"));
         //System.out.println((new String(szData,"ISO8859_1")).getBytes("UTF8"));   
         int byteCount = biStream.read(szData);
         while(byteCount != -1){
               System.out.println("get data="+byteCount);
               //System.out.println(new String(szData));    
               if(new String(szData,"ISO8859_1").indexOf("<record>") != -1) haveRecords = true;
               if(new String(szData,"ISO8859_1").indexOf("</records>") != -1){//若為xml</records>檔案結尾時,則只取到</records>字串寫入
                  logps.print(new String(szData,"ISO8859_1").substring(0,new String(szData,"ISO8859_1").indexOf("</records>")+10));
               }else{
                  //System.out.println(new String(szData,"UTF8"));                                
                  logps.write(szData,0,byteCount);                
                  //logps.write(new String(szData,"utf-8").getBytes());               
                  //System.out.println(new String(szData,"UTF8"));
               }
               byteCount = biStream.read(szData);
               
               //for(int i=0;i<szData.length;i++){//清空byte array
                   //szData[i]=0;
               //}
                           
         } 
         logps.flush();
         System.out.println("BufferedInputStream end======================================");
         
         infromURL.close();
         if(logos != null) logos.close();
         if(logbos != null) logbos.close();
         if(logps != null) logps.close();
         
         printMsg(logps3,"xml檔案下載完成"); 
         errMsg += "xml檔案下載完成<br>";         
         //101.09.18--暫時拿掉,從網站上下載xml                   
         
         //boolean haveRecords = true;//for 測試用
         //boolean records_multi = false;//for 測試用
         if(haveRecords){//有資料時  
            //設定剖析的參數
            dbf.setIgnoringComments(true);
            dbf.setIgnoringElementContentWhitespace(true);
            dbf.setCoalescing(true);
            
            DocumentBuilder db = dbf.newDocumentBuilder();
            //讀入XML文件
            document = db.parse(new File(Utility.getProperties("xlsDir")+System.getProperty("file.separator")+"exd08.xml"));
            
            //取得根元素
            Node root = document.getDocumentElement();
            System.out.println("根元素:"+root.getNodeName());         
            
            //顯示指定元素的所有子節點
            NodeList tagNodes = document.getElementsByTagName("record");
            for(int i=0;i<tagNodes.getLength();i++){
                 System.out.println("record("+i+"):");               
                 NodeList childs = tagNodes.item(i).getChildNodes();
                 //顯示所有子節點 
                 for(int j=0;j<childs.getLength();j++){
                     if(childs.item(j).getNodeType() == Node.ELEMENT_NODE){
                        //System.out.print(" +-"+childs.item(j).getNodeName());
                        try{
                            //System.out.println("/"+childs.item(j).getFirstChild().getNodeValue());
                            val = childs.item(j).getFirstChild().getNodeValue();
                            //val = val.replaceAll("\n","");//97.11.21把跳行符號取代掉 by 2295
                            //System.out.println(val);                                                  
                        }catch(Exception e){val = "";}; 
                        //99.06.07 fix 金庫代碼.在檢查局為018.農金局為0180012                        
                        if(childs.item(j).getNodeName().equals("bank_no") && val.equals("018")){//金庫代碼:檢查局機構代號018.農金局為0180012
                           val = "0180012";//機構代號為018時,轉為0180012存入bank_no
                        } 
                        //109.02.17 fix 金資中心.在檢查局為600.農金局為6000000                        
                        if(childs.item(j).getNodeName().equals("bank_no") && val.equals("600")){//金資中心代碼:檢查局機構代號600.農金局為6000000
                           val = "6000000";//機構代號為600時,轉為6000000存入bank_no
                        }
                        recordData.setProperty(childs.item(j).getNodeName(),val); 
                        if(childs.item(j).getNodeName().equals("bank_no")){
                            ba01Loop:
                            for(int k=0;k<ba01.size();k++){
                                ba01Data = (Properties)ba01.get(k);
                                if(ba01Data.getProperty("bank_no").equals(val)){
                                   recordData.setProperty("bank_type",ba01Data.getProperty("bank_type")); 
                                   recordData.setProperty("pbank_no",ba01Data.getProperty("pbank_no"));
                                   break ba01Loop;
                                }
                            }
                        }
                        //System.out.println("set prop="+childs.item(j).getNodeName()+":"+val);  
                    }
                }
                exd08.add(recordData);
                recordData = new Properties();
            }
            
            System.out.println("exd08.size()="+exd08.size());
           
            printMsg(logps3,"exd08共下載"+exd08.size()+"筆資料");
            errMsg += "exd08共下載"+exd08.size()+"筆資料<br>";
           
            List updateSQL = new LinkedList();
            StringBuffer tmpBuf = new StringBuffer();
            //List remove_exd08 = new LinkedList();
            //int exd08Len = exd08.size();
            System.out.println("begin.exd08.size()="+exd08.size());
            List content_list = null;
            boolean checkSQL = false;
            
            dbURL = Utility.getProperties("BOAFDBURL");
            userName = Utility.getProperties("rptID");
            userPwd = Utility.getProperties("rptPwd");
            
            Class.forName(driverName); 
            dbConn = DriverManager.getConnection(dbURL, userName, userPwd); 
            System.out.println("Connection Successful!");
            stat = dbConn.createStatement();
            
            for(int i=0;i<exd08.size();i++){
                recordData = (Properties)exd08.get(i);
                Properties addData = new Properties();
                
                /*
                System.out.println("record["+i+"]"); 
                System.out.println("serial="+recordData.getProperty("serial"));
                System.out.println("reportno="+recordData.getProperty("reportno"));
                System.out.println("item_no="+recordData.getProperty("item_no"));
                System.out.println("content="+recordData.getProperty("content"));
                System.out.println("content.length="+recordData.getProperty("content").getBytes().length);
                System.out.println("comment="+recordData.getProperty("comment"));            
                System.out.println("fault_id="+recordData.getProperty("fault_id"));
                System.out.println("oppinion="+recordData.getProperty("oppinion"));
                System.out.println("act_id="+recordData.getProperty("act_id"));
                System.out.println("digest="+recordData.getProperty("digest"));
                System.out.println("rec_docno="+recordData.getProperty("rec_docno"));
                System.out.println("rec_date="+recordData.getProperty("rec_date"));
                System.out.println("verify="+recordData.getProperty("verify"));
                System.out.println("upd_user="+recordData.getProperty("upd_user"));
                System.out.println("upd_date="+recordData.getProperty("upd_date"));
                */ 
                
                //105.10.17 add 檢查報告屬專案農貸查核缺失匯入frm_exDef=================================================
                try{
                int def_seq=0;
                int exMasterCount=0;
                parseExDef(recordData);
                System.out.println("trans_exDef.size()="+trans_exDef.size());                
                printMsg(logps3,"trans_exDef.size()="+trans_exDef.size());
                for(int k=0;k<trans_exDef.size();k++){
                    printMsg(logps3,"k=["+k+"]:"+trans_exDef.size());
                    exDefData = (Properties)trans_exDef.get(k);
                    def_seq=1;
                    System.out.println("record["+k+"]"); 
                    System.out.println("bank_no="+exDefData.getProperty("bank_no"));
                    System.out.println("reportno="+exDefData.getProperty("reportno"));
                    System.out.println("ex_kind="+exDefData.getProperty("ex_kind"));
                    System.out.println("loan_name="+exDefData.getProperty("loan_name"));
                    System.out.println("loan_date="+exDefData.getProperty("loan_date"));
                    System.out.println("loan_item="+exDefData.getProperty("loan_item"));
                    System.out.println("loan_amt="+exDefData.getProperty("loan_amt"));     
                    System.out.println("memo="+exDefData.getProperty("memo"));   
                    
                    printMsg(logps3,"record["+k+"]"); 
                    printMsg(logps3,"bank_no="+exDefData.getProperty("bank_no"));
                    printMsg(logps3,"reportno="+exDefData.getProperty("reportno"));
                    printMsg(logps3,"ex_kind="+exDefData.getProperty("ex_kind"));
                    printMsg(logps3,"loan_name="+exDefData.getProperty("loan_name"));
                    printMsg(logps3,"loan_date="+exDefData.getProperty("loan_date"));
                    printMsg(logps3,"loan_item="+exDefData.getProperty("loan_item"));
                    printMsg(logps3,"loan_amt="+exDefData.getProperty("loan_amt"));     
                    printMsg(logps3,"memo="+exDefData.getProperty("memo"));  
                    
                    //取得該檢查報告缺失的最大序號
                    rs = stat.executeQuery("select decode(max(def_seq),null,0,max(def_seq))+1 as def_seq from frm_exdef where ex_no='"+exDefData.getProperty("reportno")+"' and  bank_no='"+exDefData.getProperty("bank_no")+"' order by def_seq ");
                    System.out.println("select decode(max(def_seq),null,0,max(def_seq))+1 as def_seq from frm_exdef where ex_no='"+exDefData.getProperty("reportno")+"' and  bank_no='"+exDefData.getProperty("bank_no")+"' order by def_seq ");
                       
                    while (rs.next()) {        
                        System.out.println("have data:"+rs.getObject("def_seq"));                        
                        def_seq = Integer.parseInt(rs.getString("def_seq"));
                    }  
                   
                    //寫入frm_exMaster查核缺失主檔
                    rs = stat.executeQuery("select count(*) as data from frm_exMaster where ex_type='FEB' and ex_no='"+exDefData.getProperty("reportno")+"' and  bank_no='"+exDefData.getProperty("bank_no")+"'");
                    System.out.println("select count(*) as data from frm_exMaster where ex_type='FEB' and ex_no='"+exDefData.getProperty("reportno")+"' and  bank_no='"+exDefData.getProperty("bank_no")+"'");
                      
                    while (rs.next()) {        
                        System.out.println("have data:"+rs.getInt("data"));                        
                        exMasterCount = rs.getInt("data");
                    }
                    if(exMasterCount == 0){//若無查核缺失主檔時,才新增資料
                        sqlCmd = "insert into frm_exMaster("
                                + "ex_type,ex_no,bank_no," 
                                + "case_status,user_id,user_name,update_date"
                                + ")values("
                                + "'FEB',"
                                + "'"+exDefData.getProperty("reportno")+"',"
                                + "'"+exDefData.getProperty("bank_no")+"',"
                                + "'1',"
                                + "'BOAF000001',"
                                + "'BOAF000001',"
                                + "sysdate)";
                        dbConn.setAutoCommit(false);
                        stat.execute(sqlCmd);
                        printMsg(logps3,sqlCmd);
                        dbConn.commit();  
                    }
                    
                    //專案農貨.查核缺失
                    if(exDefData.getProperty("ex_kind").equals("A")){//整体性缺失
                        sqlCmd = "insert into frm_exdef("
                                + "ex_no,bank_no,def_seq,ex_kind,ex_result," 
                                + "case_status,memo,user_id,user_name,update_date"
                                + ")values("
                                + "'"+exDefData.getProperty("reportno")+"',"
                                + "'"+exDefData.getProperty("bank_no")+"',"
                                + def_seq+","
                                + "'"+exDefData.getProperty("ex_kind")+"',"
                                + "'1',"                              
                                + "'1',"                  
                                + "'"+exDefData.getProperty("memo")+"',"
                                + "'BOAF000001',"
                                + "'BOAF000001',"
                                + "sysdate)";  
                    }else{
                        sqlCmd = "insert into frm_exdef("
                                + "ex_no,bank_no,def_seq,ex_kind,ex_result,loan_name,loan_date," 
                                + "loan_item,loan_amt,case_status,memo,user_id,user_name,update_date"
                                + ")values("
                                + "'"+exDefData.getProperty("reportno")+"',"
                                + "'"+exDefData.getProperty("bank_no")+"',"
                                + def_seq+","
                                + "'"+exDefData.getProperty("ex_kind")+"',"
                                + "'1',"
                                + "'"+exDefData.getProperty("loan_name")+"',"                        
                                + "To_date('"+exDefData.getProperty("loan_date")+"','yyyy/mm/dd')"+","
                                + "'"+exDefData.getProperty("loan_item")+"',"
                                + exDefData.getProperty("loan_amt")+","
                                + "'1',"                  
                                + "'"+exDefData.getProperty("memo")+"',"
                                + "'BOAF000001',"
                                + "'BOAF000001',"
                                + "sysdate)";
                    }
                   
                    dbConn.setAutoCommit(false);
                    stat.execute(sqlCmd);
                    printMsg(logps3,sqlCmd);
                    dbConn.commit(); 
                }
                }catch(Exception e){
                    System.out.println("ExDef Error:"+e+e.getMessage());
                    printMsg(logps3,"ExDef Error:"+e+e.getMessage());
                }
                //============================================================================================
                 
                int parseLen = 2000;
                int iLen = recordData.getProperty("content").getBytes("UTF-8").length;
                if(recordData.getProperty("content").getBytes("UTF-8").length > parseLen){
                    System.out.println("reportno="+recordData.getProperty("reportno"));
                    System.out.println("item_no="+recordData.getProperty("item_no"));
                    System.out.println("content="+recordData.getProperty("content"));
                    System.out.println("iLen="+iLen);
                    if(iLen > parseLen){
                       content_list = Utility.parseLenBase64encode(recordData.getProperty("content"),2000);
                    }
                    for(int idx=0;idx<content_list.size();idx++){
                        addData = new Properties();
                        addData.setProperty("reportno_seq",recordData.getProperty("serial"));
                        addData.setProperty("reportno",recordData.getProperty("reportno"));
                        addData.setProperty("item_no",recordData.getProperty("item_no"));
                        addData.setProperty("content",(String)content_list.get(idx));
                        addData.setProperty("comment",recordData.getProperty("comment"));            
                        addData.setProperty("fault_id",recordData.getProperty("fault_id"));
                        addData.setProperty("oppinion",recordData.getProperty("oppinion"));
                        addData.setProperty("act_id",recordData.getProperty("act_id"));
                        addData.setProperty("digest",recordData.getProperty("digest"));
                        addData.setProperty("rec_docno",recordData.getProperty("rec_docno"));
                        addData.setProperty("rec_date",recordData.getProperty("rec_date"));
                        addData.setProperty("verify",recordData.getProperty("verify"));
                        addData.setProperty("upd_user",recordData.getProperty("upd_user"));
                        addData.setProperty("upd_date",recordData.getProperty("upd_date"));  
                        addData.setProperty("serial",String.valueOf(idx));  
                        trans_exd08.add(addData);
                        System.out.println("add "+addData.getProperty("reportno")+":item_no="+addData.getProperty("item_no"));                        
                    }
                }else{
                    addData = new Properties();
                    addData.setProperty("reportno_seq",recordData.getProperty("serial"));
                    addData.setProperty("reportno",recordData.getProperty("reportno"));
                    addData.setProperty("item_no",recordData.getProperty("item_no"));
                    addData.setProperty("content",recordData.getProperty("content"));
                    addData.setProperty("comment",recordData.getProperty("comment"));            
                    addData.setProperty("fault_id",recordData.getProperty("fault_id"));
                    addData.setProperty("oppinion",recordData.getProperty("oppinion"));
                    addData.setProperty("act_id",recordData.getProperty("act_id"));
                    addData.setProperty("digest",recordData.getProperty("digest"));
                    addData.setProperty("rec_docno",recordData.getProperty("rec_docno"));
                    addData.setProperty("rec_date",recordData.getProperty("rec_date"));
                    addData.setProperty("verify",recordData.getProperty("verify"));
                    addData.setProperty("upd_user",recordData.getProperty("upd_user"));
                    addData.setProperty("upd_date",recordData.getProperty("upd_date"));  
                    addData.setProperty("serial","0");  
                    trans_exd08.add(addData);                    
                }
            }//end of exd08Len 
            System.out.println("after.exd08.size()="+trans_exd08.size());
            
            for(int i=0;i<trans_exd08.size();i++){
                recordData = (Properties)trans_exd08.get(i);
                //check資料有無重覆==============================================================================
                records_multi = false;
                rs = stat.executeQuery("select count(*) as datacount from exdefgoodf where reportno ='"+recordData.getProperty("reportno")+"' and reportno_seq="+recordData.getProperty("reportno_seq"));
                System.out.println("select count(*) as datacount from exdefgoodf where reportno ='"+recordData.getProperty("reportno")+"' and reportno_seq="+recordData.getProperty("reportno_seq"));
                
                while (rs.next()) {         
                   System.out.println("exd08 multi data");
                   if(rs.getInt("datacount") > 0){
                       records_multi = true;
                       printMsg(logps3,"reportno["+recordData.getProperty("reportno")+"],reportno_seq["+recordData.getProperty("reportno_seq")+"]資料已存在");
                    }    
                } 
                //==============================================================================================                
                if(records_multi){//資料重覆時
                    //缺失改善
                    sqlCmd = " delete exdefgoodf where reportno ='"+recordData.getProperty("reportno")+"' and reportno_seq="+recordData.getProperty("reportno_seq");
                    /*
                    sqlCmd = " update exdefgoodf set "
                           + " item_no = '"+recordData.getProperty("item_no")+"',"
                           + " ex_content = '"+recordData.getProperty("content")+"',"
                           + " commentt = '"+recordData.getProperty("comment")+"',"
                           + " fault_id = '"+recordData.getProperty("fault_id")+"',"
                           + " audit_oppinion = '"+recordData.getProperty("oppinion")+"',"
                           + " act_id = '"+recordData.getProperty("act_id")+"',"
                           + " audit_result = '"+recordData.getProperty("verify")+"',"                  
                           + " user_name = '"+recordData.getProperty("upd_user")+"',"
                           + " update_date = To_date('"+recordData.getProperty("upd_date")+"','yyyy-mm-dd hh24-mi-ss')"
                           + " where reportno = '"+recordData.getProperty("reportno")+"'"
                           + " and   reportno_seq = "+recordData.getProperty("serial")
                           + " and item_no = '"+recordData.getProperty("item_no")+"'";
                    */     
                    checkSQL = false; 
                    for(int sqlidx=0;sqlidx<updateSQL.size();sqlidx++){
                        if(((String)updateSQL.get(sqlidx)).equals(sqlCmd)){
                            checkSQL = true;
                        }
                    }
                    if(!checkSQL) updateSQL.add(sqlCmd);  
                    sqlCmd = " delete exdg_historyf "
                           + " where reportno = '"+recordData.getProperty("reportno")+"'"
                           + " and   reportno_seq = "+recordData.getProperty("reportno_seq");
                    //缺失改善歷程紀錄(除了原content會分割,其餘資料皆相同)
                    /*
                    sqlCmd = " update exdg_historyf set "
                           + " digest = '"+recordData.getProperty("digest")+"',"
                           + " rt_docno = '"+recordData.getProperty("rec_docno")+"',"
                           + " rt_date = To_date('"+recordData.getProperty("rec_date")+"','mm/dd/yyyy'),"                   
                           + " audit_result = '"+recordData.getProperty("verify")+"',"                  
                           + " user_name = '"+recordData.getProperty("upd_user")+"',"
                           + " update_date = To_date('"+recordData.getProperty("upd_date")+"','yyyy-mm-dd hh24-mi-ss')"
                           + " where reportno = '"+recordData.getProperty("reportno")+"'"
                           + " and   reportno_seq = "+recordData.getProperty("serial");
                    */       
                    checkSQL = false; 
                    for(int sqlidx=0;sqlidx<updateSQL.size();sqlidx++){
                        if(((String)updateSQL.get(sqlidx)).equals(sqlCmd)){
                            checkSQL = true;
                        }
                    }
                    if(!checkSQL) updateSQL.add(sqlCmd);  
                }
                
                
                //缺失改善
                sqlCmd = "insert into exdefgoodf("
                       + "reportno,reportno_seq,item_no,ex_content,commentt,fault_id," 
                       + "audit_oppinion,act_id,audit_result,user_name,update_date,serial"
                       + ")values("
                       + "'"+recordData.getProperty("reportno")+"',"
                       + recordData.getProperty("reportno_seq")+","
                       + "'"+recordData.getProperty("item_no")+"',"
                       + "'"+recordData.getProperty("content")+"',"
                       + "'"+recordData.getProperty("comment")+"',"
                       + "'"+recordData.getProperty("fault_id")+"',"
                       + "'"+recordData.getProperty("oppinion")+"',"
                       + "'"+recordData.getProperty("act_id")+"',"
                       + "'"+recordData.getProperty("verify")+"',"                  
                       + "'"+recordData.getProperty("upd_user")+"',"
                       + "To_date('"+recordData.getProperty("upd_date")+"','yyyy-mm-dd hh24-mi-ss')"+","
                       + recordData.getProperty("serial")+")";
                updateSQL.add(sqlCmd);  
                //System.out.println("serial="+recordData.getProperty("serial"));
                //缺失改善歷程紀錄
                if(recordData.getProperty("serial").equals("0")){
                   sqlCmd = "insert into exdg_historyf("
                          + "reportno,reportno_seq,digest,rt_docno,rt_date," 
                          + "audit_result,user_name,update_date"
                          + ")values("
                          + "'"+recordData.getProperty("reportno")+"',"
                          + recordData.getProperty("reportno_seq")+","
                          + "'"+recordData.getProperty("digest")+"',"
                          + "'"+recordData.getProperty("rec_docno")+"',"
                          + "To_date('"+recordData.getProperty("rec_date")+"','mm/dd/yyyy'),"                  
                          + "'"+recordData.getProperty("verify")+"',"                  
                          + "'"+recordData.getProperty("upd_user")+"',"
                          + "To_date('"+recordData.getProperty("upd_date")+"','yyyy-mm-dd hh24-mi-ss'))";
                               
                   updateSQL.add(sqlCmd);
                }
                 
            }
            //寫入DB          
            if(updateSQL.size() >= 1){
               dbConn.setAutoCommit(false);        
               for(int idx = 0;idx < updateSQL.size();idx++){
                    stat.addBatch((String)updateSQL.get(idx));
               	 	System.out.println((String)updateSQL.get(idx));
                	printMsg(logps3,(String)updateSQL.get(idx));
               }
               int[] rowCount  = stat.executeBatch();
               System.out.println("rowCount="+rowCount.length);
               
               int i=0;
               boolean updateOK=true;
               while(i < rowCount.length){
                  if(rowCount[i] <= 0){
                    System.out.println("i="+i+":"+rowCount[i]+":sql="+(String)updateSQL.get(i));                    
                    updateOK = false;
                  }
                  printMsg(logps3,"i="+i+":updateOK="+rowCount[i]+":sql="+(String)updateSQL.get(i));
                  i++;
               }
               dbConn.commit();
               if(updateOK){                  
                  printMsg(logps3,"批次寫入資料庫成功");
                  errMsg += "批次寫入資料庫成功<br>";
               }else{
                  printMsg(logps3,"批次寫入資料庫失敗");
                  errMsg += "批次寫入資料庫失敗<br>";                   
                  send_Mail("","批次寫入資料庫失敗"); 
               }
            }
         }else{
            printMsg(logps3,"無資料可下載");
            errMsg += "無資料可下載<br>";
         }
        
     }catch(SAXException se){
        //剖析過程錯誤
        Exception e = se;
        if(se.getException() != null) e = se.getException();
        send_Mail("",e+e.getMessage());
        e.printStackTrace();        
     }catch(ParserConfigurationException pe){
        //剖析器設定錯誤
        pe.printStackTrace();       
     }catch(IOException ie){
        //檔案處理錯誤
        ie.printStackTrace();
     }catch(Exception e){
        System.out.println(e+e.getMessage());
        printMsg(logps3,e+e.getMessage());
        errMsg += e+e.getMessage()+"<br>";
        send_Mail("",e+e.getMessage());
     }finally{
        try{
            if (rs != null){
                rs.close();
                rs = null;//104.10.06 add
            }
            if (stat != null){
                stat.close();
                stat = null;//104.10.06 add
            }
            if (dbConn != null){//106.05.25 add
               if(!dbConn.isClosed()){//104.10.06        
                  dbConn.close();
                  dbConn = null;//104.10.06 add
               }
            }
        }catch(SQLException sqle){
               System.out.println(sqle+sqle.getMessage());
               printMsg(logps3,sqle+sqle.getMessage());
        }
     }
     return errMsg;
}
  
  /*檢查報告屬專案農貸查核缺失匯入*/
  private static void parseExDef(Properties recordData){
      List loan_item_list = new LinkedList();
      List item_list = null;
      String content= "";
      String content_src= "";
      String loan_name="";//借款人名稱
      //String loan_date="";//貸款日期
      String loan_item="";//貸款種類
      String loan_amt ="";//貸款金額
      //String memo = "";//原解析文字      
      int idx=0;
      Properties exDefData = null;     
      try {  
          //取得專案農貸貸款種類資料
          rs = stat.executeQuery("select loan_item,loan_item_name from frm_loan_item order by input_order");  
          //18種專案農貸貸款種類
          while (rs.next()) {        
              //System.out.println("have data:"+rs.getString("loan_item_name"));  
              item_list = new LinkedList();
              item_list.add(rs.getString("loan_item"));
              item_list.add(rs.getString("loan_item_name"));
              loan_item_list.add(item_list);
          }   
          content = recordData.getProperty("content");//缺失摘要(個案缺失使用)
          content_src = recordData.getProperty("content");//缺失摘要(整体性缺失使用)
          /*個案缺失*/
          System.out.println("個案缺失.content="+content);
          //iLength = content.length();
          System.out.println("content.length()="+content.length());
          System.out.println("loan_item_list.size()="+loan_item_list.size());
          List loan_list = new LinkedList();
          List loan_name_list = null;
          trans_exDef = new LinkedList();
          boolean haveAdd=true;
          int loan_item_idx=0;
          int loan_item_idx_s = 0;
          int loan_item_idx_e = 0;
          int loan_item_list_idx = 0;
          int firstitem=content.length();
          //找出第一個符合專案農貸18種類別關鍵字
          for(int i=0;i<loan_item_list.size();i++){      
              //System.out.println(i+"="+(String)((List)loan_item_list.get(i)).get(1));
              //System.out.println("now.frstitem="+firstitem+":loan_item_list_idx="+loan_item_list_idx+":i="+loan_item_list_idx);
              if(content.indexOf((String)((List)loan_item_list.get(i)).get(1)) != -1 && content.indexOf((String)((List)loan_item_list.get(i)).get(1)) < firstitem){
                  firstitem=content.indexOf((String)((List)loan_item_list.get(i)).get(1));
                  loan_item_list_idx = i;
              }
              //System.out.println("now.frstitem="+firstitem+":loan_item_list_idx="+loan_item_list_idx+":i="+loan_item_list_idx);
          }
          
          System.out.println("loan_item_list_idx="+loan_item_list_idx+":frst_item="+(String)((List)loan_item_list.get(loan_item_list_idx)).get(1));
          contentLoop:
          while(haveAdd && content.length() !=0){
              haveAdd=false;     
              if(content.indexOf((String)((List)loan_item_list.get(loan_item_list_idx)).get(1)) != -1){//專案農貸相關字                      
                  System.out.println("loan_itemLoop.content="+content);
                  System.out.println((String)((List)loan_item_list.get(loan_item_list_idx)).get(1)+"="+content.indexOf((String)((List)loan_item_list.get(loan_item_list_idx)).get(1)));
                  //105.10.18 add
                  loan_item_idx = content.indexOf((String)((List)loan_item_list.get(loan_item_list_idx)).get(1));
                  //依據專案農貸關鍵字回往找貸放
                  loan_item_idx_s = findBackwardChar(content,"貸放",content.indexOf((String)((List)loan_item_list.get(loan_item_list_idx)).get(1)));
                  //依據專案農貸關鍵字回後找千元
                  loan_item_idx_e = findForwardChar(content,"千元",content.indexOf((String)((List)loan_item_list.get(loan_item_list_idx)).get(1)));
                  System.out.println("loan_item_idx="+loan_item_idx+"loan_item_idx_s="+loan_item_idx_s+"loan_item_idx_e="+loan_item_idx_e);                      
                  if(loan_item_idx_s != -1 && loan_item_idx_e !=-1 && loan_item_idx_s < loan_item_idx){
                      loan_name_list = new LinkedList();
                      loan_name_list.add(loan_item_list_idx);//專案農貸.項目位置
                      if(loan_item_idx_s-10 < 0){
                          System.out.println("add:["+loan_item_list_idx+"]="+content.substring(0,loan_item_idx+loan_item_idx_e+2));
                          loan_name_list.add(content.substring(0,loan_item_idx+loan_item_idx_e+2));//個案失缺內容
                          loan_list.add(loan_name_list);
                          content = content.substring(loan_item_idx+loan_item_idx_e+2,content.length());
                          System.out.println("now.content="+content);  
                          haveAdd=true;
                       }else{
                           System.out.println("add:["+loan_item_list_idx+"]="+content.substring(loan_item_idx_s-10,loan_item_idx+loan_item_idx_e+2));
                           loan_name_list.add(content.substring(loan_item_idx_s-10,loan_item_idx+loan_item_idx_e+2));//個案失缺內容                           
                           loan_list.add(loan_name_list);
                           content = content.substring(loan_item_idx+loan_item_idx_e+2,content.length());
                           System.out.println("now.content="+content);
                           haveAdd=true;
                       }                      
                  }    
              }//end of 專案農貸相關字
              
              //找出第一個符合專案農貸18種類別關鍵字
              loan_item_list_idx = 0;
              firstitem=content.length();
              haveAdd=false;
              //找出第一個符合專案農貸18種類別關鍵字
              for(int i=0;i<loan_item_list.size();i++){      
                  //System.out.println(i+"="+(String)((List)loan_item_list.get(i)).get(1));
                  //System.out.println("now.frstitem="+firstitem+":loan_item_list_idx="+loan_item_list_idx+":i="+loan_item_list_idx);
                  if(content.indexOf((String)((List)loan_item_list.get(i)).get(1)) != -1 && content.indexOf((String)((List)loan_item_list.get(i)).get(1)) < firstitem){
                      firstitem=content.indexOf((String)((List)loan_item_list.get(i)).get(1));
                      loan_item_list_idx = i;
                      haveAdd=true;
                  }
                  //System.out.println("now.frstitem="+firstitem+":loan_item_list_idx="+loan_item_list_idx+":i="+loan_item_list_idx);
              }
              
              System.out.println("loan_item_list_idx="+loan_item_list_idx+":frst_item="+(String)((List)loan_item_list.get(loan_item_list_idx)).get(1));
             
          }
          //解釋個案字串 ex:98.4.10貸放李春男農家綜合貸款300千元
          System.out.println("loan_list.size()="+loan_list.size());
          String Temp="";
          List parse_list=null;
          loan_item_idx=0;  
          for(int i=0;i<loan_list.size();i++){            
              System.out.println("loan_list["+i+"]="+(String)((List)loan_list.get(i)).get(1));
              Temp = (String)((List)loan_list.get(i)).get(1);
              memo = (String)((List)loan_list.get(i)).get(1);
              loan_item_idx = (Integer)((List)loan_list.get(i)).get(0);
              if(Temp.indexOf("核准貸放") != -1){//106.03.24 add
                  loan_date=Temp.substring(0,Temp.indexOf("核准貸放"));
                  if(loan_date.indexOf(".") != -1){
                      parse_list = Utility.getStringTokenizerData(loan_date,".");                     
                      //System.out.println("parse_list.size()="+parse_list.size());
                      parseLoan_date(parse_list);//貸放日期轉換格式為原97.10.10轉換為yyyy/mm/dd                     
                  }
                  System.out.println("loan_date="+loan_date);
                  if(Temp.indexOf((String)((List)loan_item_list.get(loan_item_idx)).get(1)) != -1){//18種專案農貸相關字        
                     loan_name=Temp.substring(Temp.indexOf("核准貸放")+4,Temp.indexOf((String)((List)loan_item_list.get(loan_item_idx)).get(1)));                        
                     loan_item=(String)((List)loan_item_list.get(loan_item_idx)).get(0);                        
                     loan_amt =Temp.substring(Temp.indexOf((String)((List)loan_item_list.get(loan_item_idx)).get(1))+((String)((List)loan_item_list.get(loan_item_idx)).get(1)).length(),Temp.indexOf("千元"));
                     if(loan_amt.indexOf(",") != -1){
                        parse_list = Utility.getStringTokenizerData(loan_amt,",");
                        loan_amt = "";
                        //System.out.println("parse_list.size()="+parse_list.size());
                        if(parse_list.size() > 0){
                           for(idx=0;idx<parse_list.size();idx++){
                               loan_amt += (String)parse_list.get(idx);
                           }        
                        }
                     }
                     System.out.println("loan_date="+loan_date);
                     System.out.println("loan_name="+loan_name);
                     System.out.println("loan_item="+loan_item+":"+(String)((List)loan_item_list.get(loan_item_idx)).get(1));
                     System.out.println("loan_amt="+loan_amt);
                          
                     exDefData = new Properties();
                     exDefData.setProperty("bank_no",recordData.getProperty("bank_no"));
                     exDefData.setProperty("reportno",recordData.getProperty("reportno"));
                     exDefData.setProperty("ex_kind", "C");
                     exDefData.setProperty("loan_name", loan_name);
                     exDefData.setProperty("loan_date", loan_date);
                     exDefData.setProperty("loan_item", loan_item);
                     exDefData.setProperty("loan_amt", String.valueOf(Integer.parseInt(loan_amt)*1000));
                     exDefData.setProperty("memo", memo);
                     trans_exDef.add(exDefData);
                  }
              }else  if(Temp.indexOf("貸放") != -1){
                 loan_date=Temp.substring(0,Temp.indexOf("貸放"));
                 if(loan_date.indexOf(".") != -1){
                     parse_list = Utility.getStringTokenizerData(loan_date,".");                     
                     //System.out.println("parse_list.size()="+parse_list.size());
                     parseLoan_date(parse_list);//貸放日期轉換格式為原97.10.10轉換為yyyy/mm/dd                     
                 }
                 System.out.println("loan_date="+loan_date);
                 if(Temp.indexOf((String)((List)loan_item_list.get(loan_item_idx)).get(1)) != -1){//18種專案農貸相關字        
                    loan_name=Temp.substring(Temp.indexOf("貸放")+2,Temp.indexOf((String)((List)loan_item_list.get(loan_item_idx)).get(1)));                        
                    loan_item=(String)((List)loan_item_list.get(loan_item_idx)).get(0);                        
                    loan_amt =Temp.substring(Temp.indexOf((String)((List)loan_item_list.get(loan_item_idx)).get(1))+((String)((List)loan_item_list.get(loan_item_idx)).get(1)).length(),Temp.indexOf("千元"));
                    if(loan_amt.indexOf(",") != -1){
                       parse_list = Utility.getStringTokenizerData(loan_amt,",");
                       loan_amt = "";
                       //System.out.println("parse_list.size()="+parse_list.size());
                       if(parse_list.size() > 0){
                          for(idx=0;idx<parse_list.size();idx++){
                              loan_amt += (String)parse_list.get(idx);
                          }        
                       }
                    }
                    System.out.println("loan_date="+loan_date);
                    System.out.println("loan_name="+loan_name);
                    System.out.println("loan_item="+loan_item+":"+(String)((List)loan_item_list.get(loan_item_idx)).get(1));
                    System.out.println("loan_amt="+loan_amt);
                         
                    exDefData = new Properties();
                    exDefData.setProperty("bank_no",recordData.getProperty("bank_no"));
                    exDefData.setProperty("reportno",recordData.getProperty("reportno"));
                    exDefData.setProperty("ex_kind", "C");
                    exDefData.setProperty("loan_name", loan_name);
                    exDefData.setProperty("loan_date", loan_date);
                    exDefData.setProperty("loan_item", loan_item);
                    exDefData.setProperty("loan_amt", String.valueOf(Integer.parseInt(loan_amt)*1000));
                    exDefData.setProperty("memo", memo);
                    trans_exDef.add(exDefData);
                 }
              }
          }
          
          content = content_src;
          /*整体性缺失*/
          item_list = new LinkedList();
          item_list.add("");
          item_list.add("專案農貸");
          loan_item_list.add(item_list);
          item_list = new LinkedList();
          item_list.add("");
          item_list.add("政策性農業專案貸款");
          loan_item_list.add(item_list);
          item_list = new LinkedList();
          item_list.add("");
          item_list.add("農業發展基金貸款");
          loan_item_list.add(item_list);
          
          
          
          System.out.println("整体性缺失.content="+content);
          //iLength = content.length();
          System.out.println("content.length()="+content.length());
          System.out.println("loan_item_list.size()="+loan_item_list.size());
          //List loan_list = new LinkedList();
          loan_item_idx=0;
          loan_item_idx_s = 0;
          loan_item_list_idx = 0;
          firstitem=content.length();
          //找出第一個符合專案農貸18種類別關鍵字及專案農貸、政策性農業專案貸款、農業發展基金貸款共21種鍵字
          for(int i=0;i<loan_item_list.size();i++){      
              //System.out.println(i+"="+(String)((List)loan_item_list.get(i)).get(1));
              //System.out.println("now.frstitem="+firstitem+":loan_item_list_idx="+loan_item_list_idx+":i="+loan_item_list_idx);
              if(content.indexOf((String)((List)loan_item_list.get(i)).get(1)) != -1 && content.indexOf((String)((List)loan_item_list.get(i)).get(1)) < firstitem){
                  firstitem=content.indexOf((String)((List)loan_item_list.get(i)).get(1));
                  loan_item_list_idx = i;
                  haveAdd=true;
              }
              //System.out.println("now.frstitem="+firstitem+":loan_item_list_idx="+loan_item_list_idx+":i="+loan_item_list_idx);
          }
          
          System.out.println("loan_item_list_idx="+loan_item_list_idx+":frstitem="+(String)((List)loan_item_list.get(loan_item_list_idx)).get(1));
         
          if(haveAdd && content.length() !=0){   
                haveAdd=false;
                if(content.indexOf((String)((List)loan_item_list.get(loan_item_list_idx)).get(1)) != -1){//專案農貸相關字                      
                   System.out.println("loan_itemLoop.content="+content);
                   System.out.println((String)((List)loan_item_list.get(loan_item_list_idx)).get(1)+"="+content.indexOf((String)((List)loan_item_list.get(loan_item_list_idx)).get(1)));
                   loan_item_idx = content.indexOf((String)((List)loan_item_list.get(loan_item_list_idx)).get(1));
                   //依據專案農貸21種關鍵字往前找)或 、     
                   loan_item_idx_s = findBackwardChar(content,")",content.indexOf((String)((List)loan_item_list.get(loan_item_list_idx)).get(1)));
                   
                   if(loan_item_idx_s != -1  && loan_item_idx_s < loan_item_idx){
                      content = content.substring(loan_item_idx_s-2,content.length());  
                   }else{                             
                      loan_item_idx_s = findBackwardChar(content,"、",content.indexOf((String)((List)loan_item_list.get(loan_item_list_idx)).get(1)));
                      if(loan_item_idx_s != -1 && content.indexOf("、") < loan_item_idx){
                         content = content.substring(loan_item_idx_s-2,content.length());
                      }   
                   }
                   System.out.println("loan_item_idx="+loan_item_idx);
                   System.out.println("loan_item_idx_s="+loan_item_idx_s);
                   System.out.println("now.content="+content);
                   if(content.indexOf("內部稽核") != -1){//有內部稽核的關鍵字..
                      loan_item_idx = content.indexOf("內部稽核");
                      loan_item_idx_s = findForwardChar(content,"。",content.indexOf("內部稽核"));
                      System.out.println("內部稽核.loan_item_idx="+loan_item_idx);
                      System.out.println("內部稽核.loan_item_idx_s="+loan_item_idx_s);
                      if(content.indexOf("內部稽核") != -1  && (loan_item_idx+loan_item_idx_s) > content.indexOf("內部稽核")){
                         content = content.substring(0,(loan_item_idx+loan_item_idx_s)+1);  
                      }
                      System.out.println("now.content="+content);  
                      haveAdd=true;                            
                   }
                         
                   if(content.indexOf("自行查核") != -1){//有內部稽核的關鍵字..
                      loan_item_idx = content.indexOf("自行查核");
                      loan_item_idx_s = findForwardChar(content,"。",content.indexOf("自行查核"));
                      System.out.println("自行查核.loan_item_idx="+loan_item_idx);
                      System.out.println("自行查核.loan_item_idx_s="+loan_item_idx_s);
                      if(content.indexOf("自行查核") != -1  && (loan_item_idx+loan_item_idx_s) > content.indexOf("自行查核")){
                         content = content.substring(0,(loan_item_idx+loan_item_idx_s)+1);  
                      }
                      System.out.println("now.content="+content);
                      haveAdd=true;                         
                    }
                }//end of 專案農貸21種關鍵字
                 
                if(haveAdd){
                   exDefData = new Properties();
                   exDefData.setProperty("bank_no",recordData.getProperty("bank_no"));
                   exDefData.setProperty("reportno",recordData.getProperty("reportno"));
                   exDefData.setProperty("ex_kind", "A");                             
                   exDefData.setProperty("memo", content);
                   trans_exDef.add(exDefData);
                }
          }//end of 有找到21種專案農貸關鍵字      
      }catch (Exception e){
          System.out.println("parseExDef Error:"+e.getMessage());
      }      
  }
  
  
  
  public void fatalError(SAXParseException spe) throws SAXException {
    System.out.println("Fatal error at line "+spe.getLineNumber());
    System.out.println(spe.getMessage());
    throw spe;
  }

  public void warning(SAXParseException spe) {
    System.out.println("Warning at line "+spe.getLineNumber());
    System.out.println(spe.getMessage());
  }

  public void error(SAXParseException spe) {
    System.out.println("Error at line "+spe.getLineNumber());
    System.out.println(spe.getMessage());
  }
  private static void printMsg(PrintStream logps,String errRptMsg){
    if(!errRptMsg.equals("")){
       logcalendar = Calendar.getInstance(); 
       nowlog = logcalendar.getTime();
       logps.println(logformat.format(nowlog)+errRptMsg);
       logps.flush();
    }
  }
  //98.01.08 有失敗時.發送e-mail
  public static void send_Mail(String Subject,String messageText){      
    try {     
         String SMTP_HOST = Utility.getProperties("SMTP_Host").trim();
         String FROM_ADDR = Utility.getProperties("From_Addr").trim();
         String SEND_MAIL = Utility.getProperties("Send_Mail").trim();
         String UserID     = Utility.getProperties("UserID").trim();
         String PWD    = Utility.getProperties("PWD").trim();
         String feb_Addr1   = Utility.getProperties("feb_Addr1").trim();
         String feb_Addr2   = Utility.getProperties("feb_Addr2").trim();
         String feb_Addr3   = Utility.getProperties("feb_Addr3").trim();
         String feb_Addr4   = Utility.getProperties("feb_Addr4").trim();
         String Auth        = (Utility.getProperties("Auth")==null?"false":Utility.getProperties("Auth")).trim();
         boolean auth = Boolean.getBoolean(Auth);
         boolean sessionDebug = true;
         if(SEND_MAIL.equals("true")){
            if(Subject.equals("")){
                Subject = "檢查報告資料下載作業失敗";
            }
            //設定所要用的Mail 伺服器和所使用的傳送協定
            Properties props = System.getProperties();
            if (SMTP_HOST != null) props.put("mail.smtp.host", SMTP_HOST);
            if (auth) props.put("mail.smtp.auth", "true");
            //產生新的Session 服務
            Session mailSession = Session.getInstance(props, null);
            if (sessionDebug) mailSession.setDebug(true);    
            Message msg = new MimeMessage(mailSession);       
            //設定傳送郵件的發信人
            msg.setFrom(new InternetAddress(FROM_ADDR));
            //設定傳送郵件至收信人的信箱
            msg.setRecipients(Message.RecipientType.TO,InternetAddress.parse(feb_Addr1+","+feb_Addr2+","+feb_Addr3, false));
            msg.setRecipients(Message.RecipientType.CC,InternetAddress.parse(feb_Addr4, false));            
            //設定信中的主題 
            msg.setSubject(Subject);
            //設定送信的時間
            msg.setSentDate(new Date());
            Multipart mp = new MimeMultipart();
            msg.setText(messageText);   
            //Transport.send(msg);      
            Transport transport = mailSession.getTransport("smtp"); 
            transport.connect(SMTP_HOST, UserID, PWD); 
            transport.sendMessage(msg, msg.getAllRecipients()); 
            transport.close();
         }
        return ;
    }catch (Exception e){
      System.out.println("send_Mail Error:"+e.getMessage());
    }
  }
  
  private static String parseNumber(String srcStr){     
      srcStr=srcStr.substring(1,srcStr.length());
      boolean parseResut=false;
      while(!parseResut){
          try{
              if(Integer.parseInt(srcStr) > 0){
                  parseResut = true;
             }
          }catch(Exception e){
           System.out.println("parseNumber Error1="+srcStr);   
           srcStr=srcStr.substring(1,srcStr.length());
           System.out.println("parseNumber Error2="+srcStr);   
          }
      }
      return  srcStr;
  }
  
  
  private static int findBackwardChar(String srcStr,String searchStr,int loan_idx){     
      srcStr=srcStr.substring(0,loan_idx);
      int searchIdx=0;
      //boolean parseResut=false;
      //while(!parseResut){
          try{
              if(srcStr.lastIndexOf(searchStr) != -1){
                  searchIdx = srcStr.lastIndexOf(searchStr);
             }
          }catch(Exception e){
           System.out.println("findBackwardChar Error="+e.getMessage());   
           //srcStr=srcStr.substring(1,srcStr.length());
           //System.out.println("parseNumber Error2="+srcStr);   
          }
      //}
      return  searchIdx;
  }
  
  private static int findForwardChar(String srcStr,String searchStr,int loan_idx){     
      System.out.println("findForwardChar.srtStr1="+srcStr);
      srcStr=srcStr.substring(loan_idx,srcStr.length());
      System.out.println("findForwardChar.srtStr2="+srcStr);
      int searchIdx=0;
      //boolean parseResut=false;
      //while(!parseResut){
          try{
              if(srcStr.indexOf(searchStr) != -1){
                  searchIdx = srcStr.indexOf(searchStr);
             }
          }catch(Exception e){
           System.out.println("findForwardChar Error="+e.getMessage());   
           //srcStr=srcStr.substring(1,srcStr.length());
           //System.out.println("parseNumber Error2="+srcStr);   
          }
      //}
      return  searchIdx;
  }
  
  private static void parseLoan_date(List parse_list){//貸款日期轉換格式為yyyy/mm/dd
      System.out.println("parse_list.size()="+parse_list.size());
      try{
      if(parse_list.size() > 0){
          for(int idx=0;idx<parse_list.size();idx++){
              System.out.println("begin-idx="+idx+":"+(String)parse_list.get(idx));
              String tmp="";
              tmp=(String)parse_list.get(idx);
              try{
                 if(Integer.parseInt((String)parse_list.get(idx)) > 0){
                    System.out.println("ok-idx="+idx+":"+(String)parse_list.get(idx));
                 }
              }catch(Exception e){
                  System.out.println("Error-idx="+idx+":"+(String)parse_list.get(idx));
                  tmp=(String)parse_list.get(idx);
                  tmp=parseNumber(tmp);
                  System.out.println("idx="+idx+":"+tmp);
              }
              if(idx==0){
                 loan_date = String.valueOf(Integer.parseInt(tmp)+1911);    
                 memo = memo.substring(memo.indexOf(tmp),memo.length());
              }else{
                  loan_date += "/"+(tmp.length()==1?"0":"")+tmp;
              }
          }        
      }
     }catch(Exception e){
         System.out.println("parseLoan_date Error:"+e.getMessage());
     }
  }
  /*
  private static SSLSocketFactory createSSLSocketFactory() throws Exception {
      TrustManager[] trustAllCerts = new TrustManager[] { new X509TrustManager() {

          public X509Certificate[] getAcceptedIssuers() {
              return null;
          }

          public void checkClientTrusted(X509Certificate[] certs, String authType) {
          }

          public void checkServerTrusted(X509Certificate[] certs, String authType) {
          }
      } };

      SSLContext ctx = SSLContext.getInstance("TLS");
      ctx.init(null, trustAllCerts, null);
      return ctx.getSocketFactory();
  }
  */
  /*
  private static List parseLenBase64encode(String strSource,int parseLen){
      List content_list = new LinkedList();
      
      System.out.println("strSource="+strSource);
      byte[] srcDate = strSource.getBytes();
      String base64Data = Base64.encode(strSource.getBytes());
      String substr = "";
      //int parseLen = 2000;
      boolean checkLen = true;
      //System.out.println("base64Data="+base64Data);
      while(checkLen){
         if(base64Data.length() > parseLen){
             substr = base64Data.substring(0,parseLen);
             //System.out.println("substr="+substr);
             content_list.add(substr);
             //System.out.println("add="+substr);
             base64Data = base64Data.substring(parseLen,base64Data.length());
             //System.out.println("last="+base64Data);
         }else{
             checkLen = false;
         }    
      }
      content_list.add(base64Data);
      System.out.println("content_list.size()="+content_list.size());
      
      return content_list;      
  }
*/
}

