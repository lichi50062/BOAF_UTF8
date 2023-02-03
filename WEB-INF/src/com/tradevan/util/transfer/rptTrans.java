
package com.tradevan.util.transfer;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import com.tradevan.util.Utility;
//呼叫rptconvert webserver要用的
import com.rpt.convert.common.ConvertConstants;
import com.rpt.convert.common.RptConvertFactory;
import com.rpt.convert.service.RptConvert;

/**
 * @author 2295
 *
 * TODO To change the template for this generated type comment go to
 * Window - Preferences - Java - Code Style - Code Templates
 */
public class rptTrans {
    
	public static void main(String[] args){
		rptTrans t = new rptTrans();
		t.transOutputFormat("","","");
	}
	
	
	/***
     * 報表格式輸出轉換.
     * 
     * @param printStyle 轉換格式 pdf , ods, odt
     * @param fName 原輸出格式fileName
     * @param fileDir 原輸出格式檔案目錄
     * @return
     */
	
    public String transOutputFormat (String printStyle, String fName , String fileDir) {
        String newFileName = "";
        
        try {
        	System.out.println("transOutputFormat.printStyle="+printStyle);
            File file = new File((fileDir.equals("")?Utility.getProperties("reportDir"):fileDir)+System.getProperty("file.separator")+fName);
            FileInputStream fis = new FileInputStream(file);
            BufferedInputStream inputStream = new BufferedInputStream(fis);
            byte[] fileBytes = new byte[(int) file.length()];
            inputStream.read(fileBytes);
            inputStream.close();
            RptConvert rptCv = RptConvertFactory.getRptConvertService(Utility.getProperties("rptTransServer")); 
            
            //byte[] retBytes = null;
            if(printStyle.equalsIgnoreCase("pdf")) {            	
                //excel to pdf
                newFileName = rptCv.changeSubFileName(fName, ConvertConstants.DOT_PDF); //將檔案名稱轉成新的副檔名               
                //System.out.println("newFileName="+newFileName);                
                byte[] convertByte = rptCv.convertExcelToPdf(fileBytes, fName) ;               
                if(convertByte != null) {
                    System.out.println("convertByte="+convertByte.length);
                }               
                rptCv.outputFile(convertByte, (fileDir.equals("")?Utility.getProperties("reportDir"):fileDir)+System.getProperty("file.separator")+newFileName) ;
            }else if(printStyle.equalsIgnoreCase("wpdf")) {
                    //word to pdf
                    System.out.println("now is wpdf begin");
                    newFileName = rptCv.changeSubFileName(fName, ConvertConstants.DOT_PDF); //將檔案名稱轉成新的副檔名
                    System.out.println("newFileName="+newFileName);                
                    byte[] convertByte = rptCv.convertWordToPdf(fileBytes, fName) ;
                    System.out.println("now is wpdf end");
                    if(convertByte != null) {
                        System.out.println("convertByte="+convertByte.length);
                    }
                    rptCv.outputFile(convertByte, (fileDir.equals("")?Utility.getProperties("reportDir"):fileDir)+System.getProperty("file.separator")+newFileName) ;
                    System.out.println("transfile to ="+(fileDir.equals("")?Utility.getProperties("reportDir"):fileDir)+System.getProperty("file.separator")+newFileName);
            } else if(printStyle.equalsIgnoreCase("ods")) {
                //excel to ods
                newFileName = rptCv.changeSubFileName(fName, ConvertConstants.DOT_ODS); 
                //System.out.println("newFileName="+newFileName);
                byte[] convertByte = rptCv.convertExcelToOds(fileBytes, fName);
                if(convertByte != null) {
                    System.out.println("convertByte="+convertByte.length);
                }
                rptCv.outputFile(convertByte, (fileDir.equals("")?Utility.getProperties("reportDir"):fileDir)+System.getProperty("file.separator")+newFileName) ;                       
            } else if(printStyle.equalsIgnoreCase("odt")) {
                //excel to ods
            	System.out.println("now is odt trans");
                newFileName = rptCv.changeSubFileName(fName, ConvertConstants.DOT_ODT); 
                System.out.println("newFileName="+newFileName);
                byte[] convertByte = rptCv.convertWordToOdt(fileBytes, fName);
                if(convertByte != null) {
                    System.out.println("convertByte="+convertByte.length);
                }
                rptCv.outputFile(convertByte, (fileDir.equals("")?Utility.getProperties("reportDir"):fileDir)+System.getProperty("file.separator")+newFileName) ;                       
            }
            
        }catch(Exception e) {  
            System.out.println("rptTrans.transOutputFormat Error");
            e.printStackTrace();
        }
        
        return  newFileName ;
    }      
    
}
