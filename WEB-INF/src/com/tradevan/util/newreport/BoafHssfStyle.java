package com.tradevan.util.newreport;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.Region;

import java.util.*;

public class BoafHssfStyle {

        public static HSSFCellStyle setStyle(HSSFCellStyle style, HSSFFont font, String[] styleSet){

            for(int i=0;i<styleSet.length;i++){
            String setKey = styleSet[i];
            if(setKey.equals("LEFTBORDER")){
        //    	System.out.println("LEFTborder");
                style.setBorderLeft  (HSSFCellStyle.BORDER_THIN);
            }
            if(setKey.equals("LEFTBOTTOMBORDER")){
         //   	System.out.println("LEFTBOTTOMborder");
                style.setBorderLeft  (HSSFCellStyle.BORDER_THIN);
                style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            }
            if(setKey.equals("RIGHTBORDER")){
        //    	System.out.println("RIGHTborder");
                style.setBorderRight (HSSFCellStyle.BORDER_THIN);
            }
            if(setKey.equals("RIGHTBOTTOMBORDER")){
        //    	System.out.println("RIGHTBOTTOMborder");
                style.setBorderRight (HSSFCellStyle.BORDER_THIN);
                style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            }
            if(setKey.equals("BOTTOMBORDER")){
       //     	System.out.println("BOTTOMBORDER");
                style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            }
            if(setKey.equals("BORDER")){
        //    	System.out.println("border");
                style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
                style.setBorderLeft  (HSSFCellStyle.BORDER_THIN);
                style.setBorderRight (HSSFCellStyle.BORDER_THIN);
                style.setBorderTop   (HSSFCellStyle.BORDER_THIN);
            }
            if(setKey.equals("PVT")){style.setVerticalAlignment(HSSFCellStyle.VERTICAL_TOP);}
            if(setKey.equals("PVC")){style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);}
            if(setKey.equals("PVB")){style.setVerticalAlignment(HSSFCellStyle.VERTICAL_BOTTOM);}
            if(setKey.equals("PHL")){style.setAlignment(HSSFCellStyle.ALIGN_LEFT);}
            if(setKey.equals("PHC")){style.setAlignment(HSSFCellStyle.ALIGN_CENTER);}
            if(setKey.equals("PHR")){style.setAlignment(HSSFCellStyle.ALIGN_RIGHT);}
            if(setKey.equals("PHJ")){style.setAlignment(HSSFCellStyle.ALIGN_JUSTIFY);}
            if(setKey.substring(0,1).equals("F")){
                font.setFontName("Courier New");//字型
                short fontsize = (short)Integer.parseInt(setKey.substring(1,3));
                font.setFontHeightInPoints(fontsize);
                style.setFont(font);
            }
            if(setKey.equals("WRAP")){style.setWrapText(true);}
        }
                return style;
        }

}
