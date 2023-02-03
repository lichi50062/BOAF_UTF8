package com.tradevan.util;

import java.io.*;
import java.util.*;

/**
 * <p>Title: rtf字元取代副程式</p>
 * <p>Description: </p>
 * <p>Company: TV</p>
 * @author:周文瑞
 * @version 1.0
 */

public class RTFMixer
{

    String rtfSrcName = null;    //來源檔檔名
    String rtfSource  = null;    //來源檔內容
    String rtfTagetName = null;  //目的檔檔名

    public RTFMixer(String rtfSrcName, String rtfTagetName)
    {
        this.rtfSrcName   = rtfSrcName;
        this.rtfTagetName = rtfTagetName;
    }

    //讀取來源rtf
    private void readSource() throws Exception
    {
        FileReader fr       = null;
        BufferedReader br   = null;
        StringBuffer source = new StringBuffer();
        String sourceTmp    = "";

        try{
            //System.out.println("source:"+rtfSrcName);
            fr = new FileReader(rtfSrcName);
            br = new BufferedReader(fr);
            while((sourceTmp = br.readLine()) != null)
                source.append(sourceTmp);

            rtfSource = source.toString();
        }
        catch(Exception e)
        {
            throw e;
        }
        finally
        {
            if (br != null)
                br.close();
            if (fr != null)
                fr.close();
        }
    }

    //
    private void writeTarget() throws Exception
    {
        FileWriter fw     = null;
        BufferedWriter bw = null;
        try{
            fw = new FileWriter(rtfTagetName);
            bw = new BufferedWriter(fw);
            bw.write(rtfSource);
            bw.flush();
        }
        catch(Exception e)
        {
            throw e;
        }
        finally
        {
            if (bw != null)
                bw.close();
            if (fw != null)
                fw.close();
        }
    }

    public void mix(Properties props)
    {
        //String[] keys = new String[props.size()];
        ArrayList keys = new ArrayList();
        try{
            this.readSource(); //read file
            if (rtfSource != null)
            {
                Enumeration e = props.propertyNames();
                while (e.hasMoreElements())
                    keys.add(e.nextElement());

                StringBuffer rtfSourceTmp = new StringBuffer(rtfSource);
                int index = 0;
                for(int i=0; i<keys.size(); i++)
                {
                    index = rtfSource.indexOf((String)keys.get(i));
                    if (index >= 0)
                    {
                        rtfSourceTmp = rtfSourceTmp.replace(index,
                                                            index + ((String)keys.get(i)).length(),
                                                            this.format(props.getProperty((String)keys.get(i))));
                    }
                    rtfSource = rtfSourceTmp.toString();
                }
                rtfSource = rtfSourceTmp.toString();
                this.writeTarget();
            }
        }
        catch(Exception e)
        {
            e.printStackTrace();
        }
    }

    private String format(String words)  //代換textarea的換行字元
    {
        StringBuffer sb = new StringBuffer();
        ArrayList wordsTmp = new ArrayList();
        StringTokenizer st = new StringTokenizer(words,"\n");
        while(st.hasMoreTokens())
            wordsTmp.add(st.nextToken());

        for (int i=0; i<wordsTmp.size(); i++)
        {
            //if (((String)wordsTmp.get(i)).charAt(0) == ' ')
                //sb.append(" ");
            sb.append(wordsTmp.get(i));
            if (i < wordsTmp.size()-1)
                sb.append("\\par ");
        }
        //System.out.println(sb);
        return sb.toString();
    }

}
