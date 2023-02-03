package com.tradevan.util.sql;

import java.sql.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.io.*;


import com.tradevan.util.Utility;
import com.tradevan.util.sql.StoredProcedure.OutputParameter;
//import com.tradevan.util.*;

public class SQLUtils {
	/*
	static File logfile;
    static FileOutputStream logos=null;      
    static BufferedOutputStream logbos = null;
    static PrintStream logps = null;
    static Date nowlog = new Date();
    static SimpleDateFormat logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");        
    static SimpleDateFormat logfileformat = new SimpleDateFormat("yyyyMMddHHmmss");
    static Calendar logcalendar;
    static File logDir = null;
    */
    public static void printQuerySQL(String sql, List pramlist) {
        StringBuffer sb = new StringBuffer(sql);
        if (pramlist != null && pramlist.size() > 0) {
            int start = 0;
            int i = 1;
            //System.out.println("pramlist.size()="+pramlist.size());
            while (sb.indexOf("?", start) != -1) {
                start = sb.indexOf("?", start);

                String pram = "";
                
                SQLParameter obj = (SQLParameter) pramlist.get(i++);
         
                if (obj instanceof OutputParameter) {
                    pram = "returnCode";
                } else {
                    if (obj.getValue() instanceof java.lang.String) {
                        pram = obj.getValue() != null ? (String) obj.getValue() : "";
                    } else {
                        pram = obj != null ? obj.getValue().toString() : "null";
                    }
                   
                }                
                sb.replace(start, start + 1, "'" + pram + "'");         
                
            }
        }
        System.out.println("\nSQL Statement : \n" + sb.toString() + "\n");
        /*
        try{
        logDir  = new File(Utility.getProperties("logDir"));
        logfile = new File(logDir + System.getProperty("file.separator") + "FR001WB_ALL");                       
	    logos = new FileOutputStream(logfile,true);                         
	    logbos = new BufferedOutputStream(logos);
	    logps = new PrintStream(logbos);   
        System.out.println("\nSQL Statement : \n" + sb.toString() + "\n");
        logps.println(sb.toString());
        logps.flush();
        }catch(Exception e){}
        */
    }

  
    public static void closeStatement(Statement stmt) {
        if(stmt == null) {
            return ;
        }
        
        try {
            stmt.close();
            stmt = null;
        } catch(Exception e) {
            e.printStackTrace();
        }
    }
    
    public static void closeResultSet(ResultSet rs) {
        if(rs == null) {
            return ;
        }
        
        try {
            
            rs.close();
            rs = null;
        } catch(Exception e) {
            e.printStackTrace();
        }
    }
    
    public static void closeConnection(Connection conn) {
        if(conn == null) {
            return ;
        }
        
        try {
            if(!conn.isClosed()) {
                conn.close();
            }
            conn = null;  
            
        } catch(Exception e) {
            e.printStackTrace();
        }
    }
    
   
}
