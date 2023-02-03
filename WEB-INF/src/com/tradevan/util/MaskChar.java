package com.tradevan.util;

public class MaskChar {
    
    /***************************************************************************
     * 字串遮蔽功能
     * @param str              --欲遮蔽之字串 ex:A123456789
     * @param startChar        --遮蔽起始位置 ex:4
     * @param masklength       --遮蔽長度 ex:4
     * @param maskChar         --遮蔽符號 ex:*
     * @return Mask characters --ex:A12****789
     * @throws Exception
     */
    public static String maskChar(String str,int startChar,int masklength,String maskChar){
        String cover="";
        try{
            if(str.length()>0){
                for(int i=1;i<=masklength;i++) cover+=maskChar;
                str = str.substring(0, startChar-1)+cover
                        +str.substring(startChar-1+masklength, str.length());
            }
        }catch(Exception e){
            System.out.println("maskChar Exception e ="+e);
        } 
        return str;
    }
    
}
