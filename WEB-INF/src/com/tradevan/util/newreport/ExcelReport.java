package com.tradevan.util.newreport;

import com.tradevan.util.Utility;
import java.util.*;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.hssf.util.*;
import java.io.*;
import java.text.*;
import java.util.*;

public class ExcelReport
    implements Report {
  public HSSFCellStyle DEFAULT_STYLE; //有框內文置中
  public HSSFCellStyle LEFT_STYLE; //有框內文置左
  public HSSFCellStyle RIGHT_STYLE; //有框內文置右

  public HSSFCellStyle NO_BORDER_DEFAULT_STYLE; //無框內文置中
  public HSSFCellStyle NO_BORDER_RIGHT_STYLE; //無框置右
  public HSSFCellStyle NO_BORDER_LIFT_STYLE; //無框置左

  public HSSFCellStyle TITLE_STYLE; //標題用
  public HSSFCellStyle COLUMN_STYLE; //欄位名
  public HSSFCellStyle SMALL_FONT_STYLE; //有框小字

  private HSSFWorkbook wb;
  private HSSFSheet sheet;
  private HSSFPrintSetup ps;
  private HSSFFooter footer;
  private String filename = "";
  private String reportname = "";
  private int[] columnLen = null;

  ExcelReport() {
    if (filename == null) {
      filename = "";
    }
    if (reportname == null || reportname == "") {
      reportname = "Report";
    }
  }
  /**
   * 初始化sheet 並設定邊界
   */
  public void initReport() {
    wb = new HSSFWorkbook();
    sheet = wb.createSheet(reportname);
    sheet.setMargin(HSSFSheet.TopMargin, 0.4);
    sheet.setMargin(HSSFSheet.BottomMargin, 0.4);
    sheet.setMargin(HSSFSheet.RightMargin, 0.4);
    sheet.setMargin(HSSFSheet.LeftMargin, 0.4);
    footer = sheet.getFooter();
    ps = sheet.getPrintSetup();
    ps.setHeaderMargin(0.31);
    ps.setFooterMargin( (double) 0);
  }

  public void initReport(String reportname) {
    wb = new HSSFWorkbook();
    sheet = wb.createSheet(reportname);
    footer = sheet.getFooter();
    ps = sheet.getPrintSetup();
  }

  public void initReport(String reportname, String filename) {
    wb = new HSSFWorkbook();
    sheet = wb.createSheet(reportname);
    footer = sheet.getFooter();
    ps = sheet.getPrintSetup();
    this.filename = filename;
  }

  public void initStyle() {
    //設定樣式和位置(請精減style物件的使用量，以免style物件太多excel報表無法開啟)
    DEFAULT_STYLE = wb.createCellStyle(); //有框內文置中
    DEFAULT_STYLE = BoafHssfStyle.setStyle(DEFAULT_STYLE, wb.createFont(),
                                           new String[] {
                                           "BORDER", "PHC", "PVC", "F12",
                                           "WRAP"});
    LEFT_STYLE = wb.createCellStyle(); //有框內文置左
    LEFT_STYLE = BoafHssfStyle.setStyle(LEFT_STYLE, wb.createFont(),
                                        new String[] {
                                        "BORDER", "PHL", "PVC", "F12",
                                        "WRAP"});
    RIGHT_STYLE = wb.createCellStyle(); //有框內文置右
    RIGHT_STYLE = BoafHssfStyle.setStyle(RIGHT_STYLE, wb.createFont(),
                                         new String[] {
                                         "BORDER", "PHR", "PVC", "F12",
                                         "WRAP"});

    NO_BORDER_DEFAULT_STYLE = wb.createCellStyle(); //無框內文置中
    NO_BORDER_DEFAULT_STYLE = BoafHssfStyle.setStyle(NO_BORDER_DEFAULT_STYLE,
        wb.createFont(), new String[] {
        "PHC", "PVC", "F12", "WRAP"});

    NO_BORDER_RIGHT_STYLE = wb.createCellStyle(); //無框置右
    NO_BORDER_RIGHT_STYLE = BoafHssfStyle.setStyle(NO_BORDER_RIGHT_STYLE,
        wb.createFont(),
        new String[] {
        "PHR", "PVC", "F12", "WRAP"});
    NO_BORDER_LIFT_STYLE = wb.createCellStyle(); //無框置左
    NO_BORDER_LIFT_STYLE = BoafHssfStyle.setStyle(NO_BORDER_LIFT_STYLE,
                                                  wb.createFont(),
                                                  new String[] {
                                                  "PHL", "PVC", "F12", "WRAP"});

    //自定需求style
    TITLE_STYLE = wb.createCellStyle(); //標題用
    TITLE_STYLE = BoafHssfStyle.setStyle(TITLE_STYLE, wb.createFont(),
                                         new String[] {
                                         "PHC", "PVC", "F18"});
    SMALL_FONT_STYLE = wb.createCellStyle(); //有框小字
    SMALL_FONT_STYLE = BoafHssfStyle.setStyle(SMALL_FONT_STYLE, wb.createFont(),
                                              new String[] {
                                              "BORDER", "PHC", "PVC", "F08",
                                              "WRAP"});

    COLUMN_STYLE = wb.createCellStyle();/*欄位名稱*/
    COLUMN_STYLE = BoafHssfStyle.setStyle(COLUMN_STYLE, wb.createFont(),
                                          new String[] {"BORDER", "PHC", "PVC",
                                          "F14",
                                          "WRAP"});
  }

  public void doFooter() {
    // 設定cell 寬度
    if (columnLen.length > 0) {
      sheet.setColumnWidth( (short) 0, (short) 0);
      for (int i = 1; i <= columnLen.length; i++) {
        sheet.setColumnWidth( (short) i,
                             (short) (256 * (columnLen[i - 1] + 4)));
      }
    }

    //設定涷結欄位
    //sheet.createFreezePane(0,1,0,1);
    //設定頁尾
    footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
    footer.setCenter("Page:" + HSSFFooter.page() + " of " +
                     HSSFFooter.numPages());
  }

  public String outputFileName() throws IOException {
    /* Write the output to a file */
    String fName = null;
    try {
      fName = Utility.getProperties("reportDir") +
          System.getProperty("file.separator") + Utility.ISOtoBig5(filename) +
          ".xls";
    }
    catch (Exception ex) {
    }
    FileOutputStream fileOut = new FileOutputStream(fName);
    System.out.println("fileOut = " + fName);
    wb.write(fileOut);
    fileOut.close();
    return fName;
  }

  public String createReport() throws IOException {
    initReport();
    initStyle();
    doReport();
    doFooter();
    return outputFileName();
  }

  public void doReport() {
  }

  public void setBorder(HSSFCellStyle style) {
    style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
    style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
    style.setBorderRight(HSSFCellStyle.BORDER_THIN);
    style.setBorderTop(HSSFCellStyle.BORDER_THIN);
    style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
    style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
  }

  public void createCell(HSSFWorkbook wb, HSSFRow row, short column,
                         String value) {
    createCell(wb, row, column, value, true);
  }

  public void createCell(HSSFWorkbook wb, HSSFRow row, short column,
                         String value, boolean hasBorder) {
    HSSFCell cell = row.createCell(column);
    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
    cell.setCellValue(value);
    int len = value.getBytes().length;
    if (len > columnLen[column - 1]) {
      columnLen[column - 1] = len;
    }
    HSSFCellStyle style = wb.createCellStyle();
    if (hasBorder) {
      setBorder(style);
      //style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
    }
    cell.setCellStyle(style);
  }

  public void createCell(HSSFWorkbook wb, HSSFRow row, short column,
                         String value, HSSFCellStyle style) {
    HSSFCell cell = row.createCell(column);
    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
    cell.setCellValue(value);
    cell.setCellStyle(style);
  }

  public void createCell(HSSFWorkbook wb, HSSFRow row, short column,
                         int value) {
    createCell(wb, row, column, String.valueOf(value), this.DEFAULT_STYLE);
  }

  public void createCell(HSSFWorkbook wb, HSSFRow row, short column,
                         int value, boolean hasBorder) {
    if (hasBorder == true) {
      createCell(wb, row, column, String.valueOf(value),
                 this.DEFAULT_STYLE);
    }
    else {
      createCell(wb, row, column, String.valueOf(value),
                 this.NO_BORDER_DEFAULT_STYLE);
    }
  }

  private String toBig5(String s) throws UnsupportedEncodingException {
    return new String(s.getBytes("ISO-8859-1"), "Big5");
  }

  private String toISO8859_1(String s) throws UnsupportedEncodingException {
    return new String(s.getBytes("Big5"), "ISO-8859-1");
  }

  public HSSFPrintSetup getHSSFPrintSetup() {
    return ps;
  }

  public HSSFFooter getHSSFFooter() {
    return footer;
  }

  public HSSFSheet getHSSFSheet() {
    return sheet;
  }
  /**
   * 取得下一個Row，並設並列高
   * @param sheet HSSFSheet
   * @param rowNo short
   * @return HSSFRow
   */
  public HSSFRow getHSSFRow(HSSFSheet sheet, short rowNo) {
    HSSFRow row = sheet.createRow(rowNo);
    row.setHeightInPoints( (float) 20.5);
    return row;
  }

  public void setColumnWidth(int[] width) {
    this.columnLen = width;
  }

  public HSSFWorkbook getHSSFWorkbook() {
    return wb;
  }

  /*
       程式:將7碼日期轉換成中文格式
       操作:將日期拆開，組成中文年月日
       傳入:0920101
       回傳:92年01月1日
       (修正傳入值少於7碼的情形，return 原字串回去)
   */
  public String toChtDate(String inputDate) {

    String strY = "";
    String strM = "";
    String strD = "";

    int len = inputDate.length();
    String outputDate = inputDate;

    if (inputDate != null && len > 5) {
      strY = inputDate.substring(0, len - 4);
      strM = inputDate.substring(len - 4, len - 2);
      strD = inputDate.substring(len - 2);

      if (strY.substring(0, 1).equals("0")) {
        strY = strY.substring(1, strY.length());
      }
      if (strM.substring(0, 1).equals("0")) {
        strM = strM.substring(1, strM.length());
      }
      if (strD.substring(0, 1).equals("0")) {
        strD = strD.substring(1, strD.length());
      }

      outputDate = strY + "年" + strM + "月" + strD + "日";
    }
    return outputDate;
  }

  /*
       程式:將5碼年月轉換成中文格式
       操作:將日期拆開，組成中文年月
       傳入:09201
       回傳:92年01月
       (修正傳入值少於5碼的情形，return 原字串回去)
   */
  public String toChtYearMon(String inputDate) {

    String strY = "";
    String strM = "";

    int len = inputDate.length();
    String outputDate = inputDate;

    if (inputDate != null && len >= 5) {
      strY = inputDate.substring(0, len - 2);
      strM = inputDate.substring(len - 2, len);

      if (strY.substring(0, 1).equals("0")) {
        strY = strY.substring(1, strY.length());
      }
      if (strM.substring(0, 1).equals("0")) {
        strM = strM.substring(1, strM.length());
      }

      outputDate = strY + "年" + strM + "月";
    }
    return outputDate;
  }

  // 日期格式轉換 yyy/mm/dd
  // original: 7位日期字串,  singe: 分割字元
  public static String getDateFormat1(String original, String singe) {
    String formatDate = "";
    if (original != null && original.length() == 7) {
      formatDate = Integer.parseInt(original.substring(0, 3))
          + singe
          + original.substring(3, 5)
          + singe
          + original.substring(5, 7);
    }
    else {
      formatDate = original;
    }
    return formatDate;
  }

  // 取得目前系統日期時間
  public String getDataTimeFormat() {
    Calendar rnow = Calendar.getInstance();
    String YYYY = Integer.toString(rnow.get(Calendar.YEAR) - 1911);
    YYYY = (YYYY.length() == 3) ? YYYY : "0" + YYYY;
    String MM = Integer.toString(rnow.get(Calendar.MONTH) + 1);
    MM = (MM.length() == 2) ? MM : "0" + MM;
    String DD = Integer.toString(rnow.get(Calendar.DAY_OF_MONTH));
    DD = (DD.length() == 2) ? DD : "0" + DD;
    String HH = Integer.toString(rnow.get(Calendar.HOUR_OF_DAY));
    HH = (HH.length() == 2) ? HH : "0" + HH;
    String MI = Integer.toString(rnow.get(Calendar.MINUTE));
    MI = (MI.length() == 2) ? MI : "0" + MI;
    String SS = Integer.toString(rnow.get(Calendar.SECOND));
    SS = (SS.length() == 2) ? SS : "0" + SS;

    return YYYY + MM + DD;
  }

  // 取得目前系統日期時間分秒
  public String getDataTimeMinSecFormat() {
    Calendar rnow = Calendar.getInstance();
    String YYYY = Integer.toString(rnow.get(Calendar.YEAR) - 1911);
    YYYY = (YYYY.length() == 3) ? YYYY : "0" + YYYY;
    String MM = Integer.toString(rnow.get(Calendar.MONTH) + 1);
    MM = (MM.length() == 2) ? MM : "0" + MM;
    String DD = Integer.toString(rnow.get(Calendar.DAY_OF_MONTH));
    DD = (DD.length() == 2) ? DD : "0" + DD;
    String HH = Integer.toString(rnow.get(Calendar.HOUR_OF_DAY));
    HH = (HH.length() == 2) ? HH : "0" + HH;
    String MI = Integer.toString(rnow.get(Calendar.MINUTE));
    MI = (MI.length() == 2) ? MI : "0" + MI;
    String SS = Integer.toString(rnow.get(Calendar.SECOND));
    SS = (SS.length() == 2) ? SS : "0" + SS;

    return YYYY + MM + DD + HH + MI + SS;
  }

  public void setFileName(String name) {
    filename = name;
  }

  public short drawASpaceRow(short rowNo, short start, int total,
                             HSSFCellStyle style) {
    HSSFRow row = getHSSFSheet().createRow(rowNo++);
    fillSpaceCell(row, start, total, style);
    return rowNo;
  }

  public void fillSpaceCell(HSSFRow row, short start, int total,
                            HSSFCellStyle style) {
    HSSFWorkbook wb = getHSSFWorkbook();
    for (short i = start; i < total + 1; i++) {
      createCell(wb, row, i, "", style);
    }
  }

}
