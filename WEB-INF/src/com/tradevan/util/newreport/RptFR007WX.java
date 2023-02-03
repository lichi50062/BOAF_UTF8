package com.tradevan.util.newreport;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.Region;

import java.util.*;
import com.tradevan.util.bean.DetailRepBean;
import com.tradevan.util.Utility;
public class RptFR007WX
    extends ExcelReport {

  private List listData = null;
  private String reportId = "";

  private short rowNo = (short) 1;
  DetailRepBean bean;

  public RptFR007WX() {
  }

  public void doReport() { // 116030, 116060
    // 指定輸出檔名
    setFileName("Rpt_" + reportId + "_Report");
    HSSFWorkbook wb = getHSSFWorkbook(); // 建立工作區
    HSSFSheet sheet = getHSSFSheet(); //建立資料表
    HSSFPrintSetup ps = getHSSFPrintSetup(); //取得印表機設定
    // 設定欄位的寬度
    int[] cellWidth = {
        2, 2, 8, 27, 18, 20,
        5, 20, 20, 5, 19, 5,
        5, 17, 17, 17, 17, 17,
        17, 15};
    setColumnWidth(cellWidth);
    String a = "";

    //自動分頁
    sheet.setAutobreaks(false);
    //列印縮放百分比
    ps.setScale( (short) 40);
    //設定紙張大小 A4
    ps.setPaperSize( (short) 9);
    // 設定橫印
    ps.setLandscape(true);

    //設定rowNo從1開始
    rowNo = (short) 1;
    //設定表頭 為固定 先設欄的起始再設列的起始
//    setRepeatingRowsAndColumns(int sheetIndex, int startColumn, int endColumn, int startRow, int endRow) { }
    wb.setRepeatingRowsAndColumns(0, 1, 20, 1, 5);
    String seqno = "", oldSeqno = "", h1 = "";
    int index = 0;
    ArrayList s;
    while (index < listData.size()) {
      bean = (DetailRepBean) listData.get(index);

      seqno = (String) bean.getSeqNo();
      if (!seqno.equals(oldSeqno) && !oldSeqno.equals("")) {
        if (rowNo != (short) 1) { // 全部的第一頁不用換頁
          sheet.setRowBreak(rowNo++); // 換頁
        }
      }

      // 用h1判斷是否換了欄位格式
      h1 = (String) bean.getH1();
      if (h1.equals("100")) { // 報表名稱欄位
        HSSFRow row = sheet.createRow(rowNo);
        createCell(wb, row, (short) 1, "報表編號：" + reportId, NO_BORDER_LIFT_STYLE); // 報表編號
        for(int i=2;i<=20;i++){
          createCell(wb, row, (short)i, "", NO_BORDER_DEFAULT_STYLE);
        }
        sheet.addMergedRegion(new Region(rowNo, (short) 1, rowNo, (short) 20));
        rowNo++;

        row = sheet.createRow(rowNo);
        createCell(wb, row, (short) 1, (String) bean.getC21(), TITLE_STYLE); // 報表名稱
        for (int i = 2; i <= 20; i++) {
          createCell(wb, row, (short) i, "", NO_BORDER_DEFAULT_STYLE);
        }
        sheet.addMergedRegion(new Region(rowNo, (short) 1, rowNo, (short) 20));
        rowNo++;

      }
      else if (h1.equals("150")) { // 查詢條件欄位
        HSSFRow row = sheet.createRow(rowNo);
        createCell(wb, row, (short) 1, (String) bean.getC20(),
                   NO_BORDER_RIGHT_STYLE); // 查詢期間
        for (int i = 2; i <=17; i++) {
          createCell(wb, row, (short) i, "", NO_BORDER_DEFAULT_STYLE);
        }
        createCell(wb, row, (short) 18, bean.getC21(), NO_BORDER_DEFAULT_STYLE);
        createCell(wb, row, (short) 19, "", NO_BORDER_DEFAULT_STYLE);
        createCell(wb, row, (short) 20, "", NO_BORDER_DEFAULT_STYLE);
        sheet.addMergedRegion(new Region(rowNo, (short) 2, rowNo, (short) 17));
        sheet.addMergedRegion(new Region(rowNo, (short) 18, rowNo, (short) 20));
        rowNo++;
        createHeader(wb, sheet); // 建立欄位表頭
      }
      else if (h1.equals("200")) { // 表身欄位
        if(index==listData.size()-1){
          HSSFRow row = getHSSFRow(sheet,rowNo);
          for(int i=1;i<20;i++){
             createCell(wb, row, (short) i, "", NO_BORDER_DEFAULT_STYLE);
          }
          rowNo++;
        }
        HSSFRow row = getHSSFRow(sheet,rowNo);
        createCell(wb, row, (short) 1, (String) bean.getC1(), DEFAULT_STYLE);
        createCell(wb, row, (short) 2, (String) bean.getC2(), DEFAULT_STYLE);
        createCell(wb, row, (short) 3, (String) bean.getC3(), DEFAULT_STYLE);
        createCell(wb, row, (short) 4, (String) bean.getC21(), LEFT_STYLE);
        createCell(wb, row, (short) 5, Utility.setCommaFormat((String) bean.getN1()), RIGHT_STYLE);
        createCell(wb, row, (short) 6, Utility.setCommaFormat((String) bean.getN2()), RIGHT_STYLE);
        createCell(wb, row, (short) 7, Utility.setCommaFormat((String) bean.getN3()), RIGHT_STYLE);
        createCell(wb, row, (short) 8, Utility.setCommaFormat((String) bean.getN4()), RIGHT_STYLE);
        createCell(wb, row, (short) 9, Utility.setCommaFormat((String) bean.getN5()), RIGHT_STYLE);
        createCell(wb, row, (short) 10, Utility.setCommaFormat((String) bean.getN6()), RIGHT_STYLE);
        createCell(wb, row, (short) 11, Utility.setCommaFormat((String) bean.getN7()), RIGHT_STYLE);
        createCell(wb, row, (short) 12, Utility.setCommaFormat((String) bean.getN8()), RIGHT_STYLE);
        createCell(wb, row, (short) 13, Utility.setCommaFormat((String) bean.getN9()), RIGHT_STYLE);
        createCell(wb, row, (short) 14, Utility.setCommaFormat((String) bean.getN10()), RIGHT_STYLE);
        createCell(wb, row, (short) 15, Utility.setCommaFormat((String) bean.getN11()), RIGHT_STYLE);
        createCell(wb, row, (short) 16, Utility.setCommaFormat((String) bean.getN12()), RIGHT_STYLE);
        createCell(wb, row, (short) 17, Utility.setCommaFormat((String) bean.getN13()), RIGHT_STYLE);
        createCell(wb, row, (short) 18, Utility.setCommaFormat((String) bean.getN14()), RIGHT_STYLE);
        createCell(wb, row, (short) 19, Utility.setCommaFormat((String) bean.getN15()), RIGHT_STYLE);
        createCell(wb, row, (short) 20, Utility.setCommaFormat((String) bean.getN16()), RIGHT_STYLE);
        rowNo++;
      }
      //用seqno判斷是否要換頁

      oldSeqno = seqno; // 紀錄舊的seqno
      index++;
    }
  }

  private void createHeader(HSSFWorkbook wb, HSSFSheet sheet) {
    HSSFRow row = sheet.createRow(rowNo);
    createCell(wb, row, (short) 1, "", NO_BORDER_DEFAULT_STYLE);
    createCell(wb, row, (short) 2, "", NO_BORDER_DEFAULT_STYLE);
    createCell(wb, row, (short) 3, "", NO_BORDER_DEFAULT_STYLE);
    createCell(wb, row, (short) 4, "", NO_BORDER_DEFAULT_STYLE);
    createCell(wb, row, (short) 5, "A01", COLUMN_STYLE);
    createCell(wb, row, (short) 6, "", COLUMN_STYLE);
    createCell(wb, row, (short) 7, "", COLUMN_STYLE);
    createCell(wb, row, (short) 8, "A04", COLUMN_STYLE);
    createCell(wb, row, (short) 9, "", COLUMN_STYLE);
    createCell(wb, row, (short) 10, "", COLUMN_STYLE);
    createCell(wb, row, (short) 11, "", COLUMN_STYLE);
    createCell(wb, row, (short) 12, "", COLUMN_STYLE);
    createCell(wb, row, (short) 13, "", COLUMN_STYLE);
    createCell(wb, row, (short) 14, "", COLUMN_STYLE);
    createCell(wb, row, (short) 15, "", COLUMN_STYLE);
    createCell(wb, row, (short) 16, "A04_逾期已達列報逾放條件而准免列報者(C)", COLUMN_STYLE);
    createCell(wb, row, (short) 17, "", COLUMN_STYLE);
    createCell(wb, row, (short) 18, "", COLUMN_STYLE);
    createCell(wb, row, (short) 19, "", COLUMN_STYLE);
    createCell(wb, row, (short) 20, "", COLUMN_STYLE);
    sheet.addMergedRegion(new Region(rowNo, (short) 1, rowNo, (short) 4));
    sheet.addMergedRegion(new Region(rowNo, (short) 5, rowNo, (short) 7));
    sheet.addMergedRegion(new Region(rowNo, (short) 8, rowNo, (short) 15));
    sheet.addMergedRegion(new Region(rowNo, (short) 16, rowNo, (short) 20));
    rowNo++;

    row = sheet.createRow(rowNo);
    createCell(wb, row, (short) 1, "逾放不符(*)", COLUMN_STYLE);
    createCell(wb, row, (short) 2, "總訪款不符(*)", COLUMN_STYLE);
    createCell(wb, row, (short) 3, "單位代號", COLUMN_STYLE);
    createCell(wb, row, (short) 4, "單位名稱", COLUMN_STYLE);
    createCell(wb, row, (short) 5, "逾期放款\n[990000]", COLUMN_STYLE);
    createCell(wb, row, (short) 6, "總放款\n(放款含催收)\n[120000+\n120800+\n150300]",
               COLUMN_STYLE);
    createCell(wb, row, (short) 7, "逾放\n比率", COLUMN_STYLE);
    createCell(wb, row, (short) 8, "1.逾期放款\n[840740]", COLUMN_STYLE);
    createCell(wb, row, (short) 9, "2.總放款\n（放款含催收）\n[840750]", COLUMN_STYLE);
    createCell(wb, row, (short) 10, "3.逾放\n比率", COLUMN_STYLE);
    createCell(wb, row, (short) 11, "4.應予觀察\n放款金額\n[840760]", COLUMN_STYLE);
    createCell(wb, row, (short) 12, "5.應予觀察放款佔總放款比率", COLUMN_STYLE);
    createCell(wb, row, (short) 13, "6.逾期放款及應予觀察放款佔總放款比率", COLUMN_STYLE);
    createCell(wb, row, (short) 14, "放款本金未超過清償期三個月，惟利息未按期繳納超過三個月至六個月者840710(A)",
               COLUMN_STYLE);
    createCell(wb, row, (short) 15, "中長期分期償債放款，未按期攤還超過三個月至六個月者840720(B)",
               COLUMN_STYLE);
    createCell(wb, row, (short) 16,
               "協議分期償還放款，協議條件符合規定，且借款戶依協議條件按期履約未滿六個月者840731(a)", COLUMN_STYLE);
    createCell(wb, row, (short) 17,
        "已獲信用保證基金同意理賠款項或有足額存單或存款備償(須辦妥質權設定且徵得發單行拋棄抵銷權同意書)，而約定待其他債務人財產處分後再予沖償者840732(b)",
               COLUMN_STYLE);
    createCell(wb, row, (short) 18, "已確定分配之債權，惟尚未接獲分配款者840733( c)",
               COLUMN_STYLE);
    createCell(wb, row, (short) 19, "債務人兼擔保品提供人死亡，於辦理繼承期間屆期而未清償之放款840734(d)",
               COLUMN_STYLE);
    createCell(wb, row, (short) 20, "其他\n840735(e)", COLUMN_STYLE);
    rowNo++;
  }

  public void setReportId(String reportId) {
    this.reportId = reportId;
  }

  public void setListData(List list) {
    listData = list;
  }
}
