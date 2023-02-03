/*
  106.03.09 add 排程產檔至官網提供官網公示揭露查詢 by 2968
  106.06.19 add 調整為UTF8格式 by 2295 
  106.07.13 fix 調整信用部分部主任.position_code='1' by 2295
  107.05.16 add 排程產檔至官網提供官網公示揭露查詢 by 6417
  107.07.04 add 內部控制制度聲明書 by 2295
*/
package com.tradevan.util;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class AutoOffice {

	private final static String OPEN_DIR = "openDir";
	private final static String LOG_DIR = "logDir";

	private Calendar calendar = null;
	private Date currenTime = null;
	private SimpleDateFormat simpleDateFormat = null;
	private File file = null;

	public static void main(String[] args) {
		System.out.println("AutoOffice begin");
		AutoOffice a = new AutoOffice();
		a.exeOffice();
		System.out.println("AutoOffice end");
	}

	public void exeOffice() {

		FileOutputStream logoFileOutput = null;
		BufferedOutputStream logBufferOutput = null;
		PrintStream logPrintStream = null;

		try {

			File logfile = this.getLogFile();
			logoFileOutput = new FileOutputStream(logfile, true);
			logBufferOutput = new BufferedOutputStream(logoFileOutput);
			logPrintStream = new PrintStream(logBufferOutput);

			String openDirPath = Utility.getProperties(OPEN_DIR);
			List<DataObject> dbData = this.getDataObjectList();
			logPrintStream.println(this.getLogTime() + " 開始產檔");
			if (dbData.size() == 0) {
				File outputFile = new File(openDirPath + "office.txt");
				outputFile.delete();
			} else {
				String tbank_no = "", bank_no = "", exchange_no = "", bank_type = "", bank_name = "", hsien_name = "", area_name = "", addr = "", telno = "", start_date = "", agree_date = "", agree_no = "", name = "", business_item = "", file_type = "", append_file = "", append_link = "",fileString = "";
				String file_type_m = "", append_file_m = "", append_link_m = "",fileString_m = "";
				String result = "";
				String newLineSign = "\\n";
				String separationSign = "@@@";
				Writer out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(openDirPath + "office.txt"), "UTF-8"));// 106.06.19
				for (int index = 0; index < dbData.size(); index++) {
					DataObject obj = (DataObject) dbData.get(index);
					tbank_no = this.getDataObjectValue(obj, "tbank_no", newLineSign);
					bank_no = this.getDataObjectValue(obj, "bank_no", newLineSign);
					exchange_no = this.getDataObjectValue(obj, "exchange_no", newLineSign);
					bank_type = this.getDataObjectValue(obj, "bank_type", newLineSign);
					bank_name = this.getDataObjectValue(obj, "bank_name", newLineSign);
					hsien_name = this.getDataObjectValue(obj, "hsien_name", newLineSign);
					area_name = this.getDataObjectValue(obj, "area_name", newLineSign);
					addr = this.getDataObjectValue(obj, "addr", newLineSign);
					telno = this.getDataObjectValue(obj, "telno", newLineSign);
					start_date = this.getDataObjectValue(obj, "start_date", newLineSign);
					agree_date = this.getDataObjectValue(obj, "agree_date", newLineSign);
					agree_no = this.getDataObjectValue(obj, "agree_no", newLineSign);
					name = this.getDataObjectValue(obj, "name", newLineSign);
					business_item = this.getDataObjectValue(obj, "business_item", newLineSign);
					//其他檔案
					file_type = this.getDataObjectValue(obj, "file_type", "");
					file_type = "".equals(file_type)?file_type:file_type+newLineSign;
					append_file = this.getDataObjectValue(obj, "append_file", "");
					append_file = "".equals(append_file)?append_file:append_file+newLineSign;
					append_link = this.getDataObjectValue(obj, "append_link", "");
					if("".equals(append_file)&&"".equals(file_type)){
						fileString = newLineSign;
					}else{
						fileString = file_type + append_file ;						
					}
					if (!"".equals(append_link)) {
                        this.copyFile(append_link, openDirPath);
                    }
					//內部控制制度聲明書 107.07.04 add======================================================================
                    file_type_m = this.getDataObjectValue(obj, "file_type_m", "");
                    file_type_m = "".equals(file_type_m)?file_type_m:file_type_m+newLineSign;
                    append_file_m = this.getDataObjectValue(obj, "append_file_m", "");
                    append_file_m = "".equals(append_file_m)?append_file_m:append_file_m+newLineSign;
                    append_link_m = this.getDataObjectValue(obj, "append_link_m", "");
                    if("".equals(append_file_m)&&"".equals(file_type_m)){
                        fileString_m = newLineSign;
                    }else{
                        fileString_m = file_type_m + append_file_m ;                      
                    }
					if (!"".equals(append_link_m)) {
						this.copyFile(append_link_m, openDirPath);
					}
					//===============================================================================================
					result = tbank_no + separationSign + bank_no + separationSign + exchange_no + separationSign + bank_type + separationSign + bank_name + separationSign + hsien_name + separationSign + area_name + separationSign + addr + separationSign + telno + separationSign + start_date + separationSign + agree_date + separationSign + agree_no + separationSign + name + separationSign
							+ business_item  + separationSign + fileString_m + fileString + separationSign;
					//System.out.println("Result =" + result);
					logPrintStream.println(this.getLogTime() + " " + result);
					out.write(result);
					out.write(13);
					out.write(10);
				}
				//System.out.println("4");
				out.close();
			}
			logPrintStream.println(this.getLogTime() + " 產檔結束");
			logPrintStream.flush();
		} catch (Exception e) {
			System.out.println(e.getMessage());
			logPrintStream.println(this.getLogTime() + " " + "AutoOffice Error:" + e + "\n" + e.getMessage());
			logPrintStream.flush();
		} finally {
			try {
				if (logoFileOutput != null)
					logoFileOutput.close();
				if (logBufferOutput != null)
					logBufferOutput.close();
				if (logPrintStream != null)
					logPrintStream.close();
			} catch (Exception ioe) {
				System.out.println(ioe.getMessage());
			}
		}
	}

	private String getDataObjectValue(DataObject obj, String value, String nullValue) {
		value = ((obj).getValue(value) == null) ? nullValue : ((String) (obj).getValue(value));
		return value;
	}

	private File getLogFile() throws Exception {
		File logDir = new File(Utility.getProperties(LOG_DIR));
		if (!logDir.exists()) {
			if (!Utility.mkdirs(Utility.getProperties(LOG_DIR))) {
				System.out.println("目錄新增失敗");
			}
		}
		this.simpleDateFormat = new SimpleDateFormat("yyyyMMddHHmmss");
		logDir = new File(logDir + System.getProperty("file.separator") + "AutoOffice." + this.simpleDateFormat.format(this.getCurrentTime()));
		return logDir;
	}

	private String getLogTime() {
		currenTime = this.getCurrentTime();
		this.simpleDateFormat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");
		String logTime = this.simpleDateFormat.format(currenTime);
		return logTime;
	}

	private List<DataObject> getDataObjectList() {
		// 100.07.21 add 查詢年度100年以前.縣市別不同===============================
		String cd01_table = (Integer.parseInt(Utility.getYear()) < 100) ? "cd01_99" : "";
		System.out.println("cd01_table :" + cd01_table);
		String cd02_table = (Integer.parseInt(Utility.getYear()) < 100) ? "cd02_99" : "";
		System.out.println("cd02_table :" + cd02_table);
		String wlx01_m_year = (Integer.parseInt(Utility.getYear()) < 100) ? "99" : "100";
		System.out.println("wlx01_m_year :" + wlx01_m_year);
		// =====================================================================

		StringBuffer sqlCmd = new StringBuffer();
		sqlCmd.append("select ");//
		sqlCmd.append("     bn01.bank_no as tbank_no,");// --總機構代號
		sqlCmd.append("		'-1' as bank_no,");// --分支機構代號
		sqlCmd.append("		exchange_no,");// --通儲代號
		sqlCmd.append(" 	decode(bank_type,'6','農會','7','漁會','') as bank_type,");// --金融機構類別
		sqlCmd.append(" 	bank_name,");// --機構名稱
		sqlCmd.append("     hsien_name,");// --縣市別
		sqlCmd.append("     area_name,");// --鄉鎮別
		sqlCmd.append("     addr,");// --地址
		sqlCmd.append("     wlx01.telno,");// --電話
		sqlCmd.append("     F_TRANSCHINESEDATE(start_date) as start_date,");// --開始營業日
																			// 107.05.16
																			// add
		sqlCmd.append("     F_TRANSCHINESEDATE(agree_date) as agree_date,");// --最近地方主管機關同意/備查信用部分部主任函日期
																			// 107.05.16
																			// add
		sqlCmd.append("     agree_no,");// --最近地方主管機關同意/備查信用部分部主任函文號
		sqlCmd.append("     name,");// --信用部主任姓名
		sqlCmd.append("     replace(F_TRANSBUSITEM(business_item,business_item_extra),'、','\\n') as business_item, ");// --營業項目
		sqlCmd.append("     wlx01_uploadM.file_type as file_type_m,");//--內部控制制度聲明書-名稱
		sqlCmd.append("     wlx01_uploadM.append_file as append_file_m,");//--內部控制制度聲明書-檔名
		sqlCmd.append("     wlx01_uploadM.append_link as append_link_m,");//--內部控制制度聲明書檔案實際位置107.07.03 add
		sqlCmd.append("     wlx01_upload.file_type,");// --其他檔案-名稱
		sqlCmd.append("     wlx01_upload.append_file,");// --其他檔案-檔名
		sqlCmd.append("     wlx01_upload.append_link ");// --其他檔案實際位置
		sqlCmd.append(" from  ");
		sqlCmd.append(" 	(select * from wlx01 where m_year=?  and wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL)wlx01 ");
		sqlCmd.append(" 	left join (select * from bn01 where m_year=? and bank_type in ('6','7') and bn_type <> '2')bn01  on wlx01.bank_no=bn01.bank_no ");
		sqlCmd.append(" 	left join (select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y')cd01 on wlx01.hsien_id=cd01.hsien_id ");
		sqlCmd.append(" 	left join ").append(cd02_table).append(" cd02 on wlx01.hsien_id = cd02.hsien_id and wlx01.area_id = cd02.area_id ");
		sqlCmd.append(" 	left join (select * from wlx01_m where position_code='4' and abdicate_code <> 'Y')wlx01_m on wlx01.bank_no=wlx01_m.bank_no ");
		sqlCmd.append("     left join (");
		sqlCmd.append("     select wlx01_upload.*");
		sqlCmd.append("     from  (select bank_no,max(seq_no) as seq_no");
		sqlCmd.append("     from wlx01_upload where file_kind='M'"); 
		sqlCmd.append("     group by bank_no)wlx01_upload_max");
		sqlCmd.append("     left join wlx01_upload on wlx01_upload_max.bank_no = wlx01_upload.bank_no and wlx01_upload_max.seq_no=wlx01_upload.seq_no)wlx01_uploadM on wlx01.bank_no = wlx01_uploadM.bank_no");//--總機構最新上傳內部控制聲明書 107.07.03 add
		sqlCmd.append(" 	left join (");
		sqlCmd.append("     select wlx01_upload.*");
		sqlCmd.append("     from (select bank_no,max(seq_no) as seq_no");
		sqlCmd.append("     from wlx01_upload where file_kind='O' and open_flag='Y'");//107.07.04 add 其他類型檔案,設定為已揭露才可上傳
		sqlCmd.append("     group by bank_no)wlx01_upload_max");
		sqlCmd.append(" 	left join wlx01_upload on wlx01_upload_max.bank_no = wlx01_upload.bank_no and wlx01_upload_max.seq_no=wlx01_upload.seq_no)wlx01_upload on wlx01.bank_no = wlx01_upload.bank_no"); // 總機構最新上傳檔案
		sqlCmd.append(" where ");
		sqlCmd.append(" 	bank_type in ('6','7') ");
		sqlCmd.append(" union ");
		sqlCmd.append(" select ");
		sqlCmd.append(" 	bn02.tbank_no,");
		sqlCmd.append(" 	bn02.bank_no,");
		sqlCmd.append(" 	exchange_no,");
		sqlCmd.append(" 	decode(bank_type,'6','農會','7','漁會','') as bank_type,");
		sqlCmd.append(" 	bank_name,");
		sqlCmd.append(" 	hsien_name,");
		sqlCmd.append(" 	area_name,");
		sqlCmd.append(" 	addr,");
		sqlCmd.append(" 	wlx02.telno,");
		sqlCmd.append(" 	F_TRANSCHINESEDATE(start_date) as start_date,");
		sqlCmd.append(" 	F_TRANSCHINESEDATE(refer_date) as refer_date,");// --最近地方主管機關備查信用部分部主任函日期
																			// 107.05.16
																			// add
		sqlCmd.append(" 	refer_no,");// --最近地方主管機關備查信用部分部主任函文號 107.05.16 add
		sqlCmd.append(" 	name,");
		sqlCmd.append(" 	replace(F_TRANSBUSITEM(business_item,business_item_extra),'、','\\n') as business_item,'','','','','','' ");
		sqlCmd.append(" from ");
		sqlCmd.append(" 	(select * from wlx02 where m_year=?  and wlx02.cancel_no <> 'Y' OR wlx02.cancel_no IS NULL)wlx02 ");
		sqlCmd.append(" 	left join (select * from bn02 where m_year=? and bank_type in ('6','7') and bn_type <> '2')bn02  on wlx02.bank_no=bn02.bank_no ");
		sqlCmd.append(" 	left join (select * from ").append(cd01_table).append(" cd01 where cd01.hsien_id <> 'Y')cd01 on wlx02.hsien_id=cd01.hsien_id ");
		sqlCmd.append(" 	left join ").append(cd02_table).append(" cd02 on wlx02.hsien_id = cd02.hsien_id and wlx02.area_id = cd02.area_id ");
		sqlCmd.append(" 	left join (select * from wlx02_m where position_code='1' and abdicate_code <> 'Y')wlx02_m on wlx02.bank_no=wlx02_m.bank_no ");
		sqlCmd.append(" where ");
		sqlCmd.append(" 	bank_type in ('6','7') ");
		sqlCmd.append(" order by tbank_no,bank_no ");
		System.out.println("sqlCmd="+sqlCmd);
		List<String> paramList = new ArrayList<String>();
		paramList.add(wlx01_m_year);
		paramList.add(wlx01_m_year);
		paramList.add(wlx01_m_year);
		paramList.add(wlx01_m_year);

		List<DataObject> dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(), paramList, "tbank_no,bank_no,exchange_no,bank_type,bank_name,hsien_name,area_name,addr,telno,start_date,agree_date,agree_no,name,business_item,file_type_m,append_file_m,append_link_m,file_type,append_file,append_link");
		return dbData;
	}

	private Date getCurrentTime() {
		this.calendar = Calendar.getInstance();
		this.currenTime = calendar.getTime();
		return currenTime;
	}

	private void copyFile(String resource, String destination) {
		this.file = new File(resource);
		InputStream input = null;
		OutputStream output = null;

		try {
			
			System.out.println(resource);
			System.out.println(destination + file.getName());
			input = new FileInputStream(resource);
			output = new FileOutputStream(destination + file.getName());
			byte[] buffer = new byte[1024];
			int length;
			while ((length = input.read(buffer)) > 0) {
				output.write(buffer, 0, length);
			}
			System.out.println("檔案複製完成");
		} catch (FileNotFoundException e) {
			System.out.println("檔案不存在,請檢查檔案");
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				if (input != null) {
					input.close();
				}
				if (output != null) {
					output.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
		}

	}
}
