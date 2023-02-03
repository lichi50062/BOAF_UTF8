//97.09.22-23 fix �R��uploadDir�U���ɮ�
//                1.���R��uploadDir�U���ɮ�
//                2.�R��uploadDir�U�ؿ����ɮ�.ex:bank�U���ɮ�
//                3.�R��uploadDir�ؿ��U���ؿ��U���ɮ�.ex:bank\about���ɮ� by 2295
package UploadHtml;
import java.io.*;
import java.util.*;
import java.util.zip.*;


public class Unzip {
  //private static String dirpath = "/APSCGF/JAVA/htmlDir.properties";//������
  private static String dirpath = "D:\\workProject\\BOAF\\WEB-INF\\classes\\UploadHtml\\htmlDir.properties";//���ժ�
  public static final void copyInputStream(InputStream in, OutputStream out)  throws IOException
  {
    byte[] buffer = new byte[1024];
    int len;

    while((len = in.read(buffer)) >= 0)
      out.write(buffer, 0, len);

    in.close();
    out.close();
  }
  
  public static final void main(String[] args) {
  	 if(args.length != 1) {
        System.err.println("Usage: Unzip zipfile");
        return;
     }
  	 Unzip unzip = new Unzip();
  	 System.out.println(unzip.exeUnZIP(args[0]));
  	 /*
  	 try{
  	     File flag_txt = new File(unzip.getProperties("homeDir")+System.getProperty("file.separator")+"flag.txt");
  	     FileWriter fw=null;
         BufferedWriter  bw =null;
         fw = new FileWriter(flag_txt);
         bw = new BufferedWriter(fw);
         bw.close();
         System.out.println("create flag.txt");
  	 }catch(Exception e){
  	 	System.out.println("unzip Error:"+e+e.getMessage());
  	 }
  	 */
  	  
  }
  
  public static String exeUnZIP(String filename) {
    Enumeration entries;
    ZipFile zipFile;  

    try {
    	System.out.println("filename="+getProperties("homeDir")+System.getProperty("file.separator")+filename);
    	zipFile = new ZipFile(getProperties("homeDir")+System.getProperty("file.separator")+filename);
    	entries = zipFile.entries();
    	while(entries.hasMoreElements()) {
    		ZipEntry entry = (ZipEntry)entries.nextElement();

    		if(entry.isDirectory()) {
    			// Assume directories are stored parents first then children.
    			System.err.println("Extracting directory: " + entry.getName());
    			// This is not robust, just for demonstration purposes.
    			(new File(getProperties("homeDir")+System.getProperty("file.separator")+entry.getName())).mkdir();
    			continue;
    		}
            /*
           if(entry.getName().indexOf("\\") != -1){
         	 System.err.println("Extracting file1: " + entry.getName().substring(0,entry.getName().indexOf("\\")-1)+ISOtoBig5(entry.getName().substring(entry.getName().indexOf("\\"),entry.getName().length())));
           }
           System.err.println("Extracting file2: " + toBig5Inverse(entry.getName()));
           System.err.println("Extracting file3: " + toBig5Convert(entry.getName()));
           System.err.println("Extracting file4: " + toBig5Inverse(toBig5Convert(entry.getName())));
           */
    		System.err.println("Extracting file: " + entry.getName());
    		copyInputStream(zipFile.getInputStream(entry),
    				new BufferedOutputStream(new FileOutputStream(getProperties("homeDir")+System.getProperty("file.separator")+ISOtoBig5(entry.getName()))));
    	}

    	zipFile.close();
    	/*
    	File flag_txt = new File(getProperties("homeDir")+System.getProperty("file.separator")+"flag.txt");
 	    FileWriter fw=null;
        BufferedWriter  bw =null;
        fw = new FileWriter(flag_txt);
        bw = new BufferedWriter(fw);
        bw.close();
        System.out.println("create flag.txt");
        */
    	return new String("unzip complete");
    } catch (Exception e) {
      System.err.println("Unhandled exception:"+e.getMessage());
      e.printStackTrace();
      return "";
    }
  }
  
//==========================================================================================	
	public  static String toBig5Convert(String s) //=====�त��X============
	{  
  		try {
    			return new String(s.getBytes(),"ISO8859_1");
  		}catch(UnsupportedEncodingException e) {
    			return s;
  		}	
	}
//==========================================================================================  	
	public  static String toBig5Inverse(String s) //========���त��X
	{  
  		try {
    			return new String(s.getBytes("ISO8859_1"));
  		}catch(UnsupportedEncodingException e) {
    			return s;
  		}	
	}  
//==========================================================================================
  
  public static String ISOtoBig5(String s) {
  		if (s == null) s = "";
		else s = s.trim();        
    		try {
      			return new String(s.getBytes("ISO8859_1"), "Big5");
    		}
    		catch (UnsupportedEncodingException e) {
      			return s;
    		}
        
    }
  /******************************************************************8
   * get property from PathConfig.dat
   */
  public static String getProperties(String key) throws Exception{
  	String value	= "";
  	try {
  		Properties p = new Properties();
  		p.load(new FileInputStream(dirpath));
  		value = (String) p.get(key);
  	}catch (FileNotFoundException e) {
  		System.out.println("Unzip.FileNotFoundException["+e.getMessage()+"]");
  		throw new Exception("Unzip.FileNotFoundException["+e.getMessage()+"]");
  	}catch (Exception e) {
  		System.out.println("Unzip.IOException["+e.getMessage()+"]");
  		throw new Exception("Unzip.IOException["+e.getMessage()+"]");
  	}
  	
  	return value;
  }
  
  public static boolean deleteUploadDir(){
  	boolean result = false;
  	File uploadFile = null;
    String[] fname1;	
	List fname_sub = new LinkedList();//�x�suploadDir�U���l�ؿ� ex.bank....
	List fname_sub1 = new LinkedList();//�x�sbank�U���l�ؿ� ex.bank\about
	File tmpFile = null;
	File subfile = null;
	String nowDir = "";
  	try{
  		 uploadFile = new File(getProperties("homeDir"));
		 if(uploadFile.exists() && uploadFile.isDirectory()){
           fname1= uploadFile.list(); //====�C�X���ؿ��U���Ҧ��ɮ�===================
           //�R��uploadDir�U���ɮ�.�ڥؿ����ɮ�(�u�R���ɮ�.���R���ؿ�)
           for(int c=0;c<fname1.length;c++){
    	          tmpFile = new File(getProperties("homeDir")+System.getProperty("file.separator")+fname1[c]);
    	          //System.out.println(getProperties("homeDir")+System.getProperty("file.separator")+fname1[c]);                  	            
                  if(!tmpFile.isDirectory() && tmpFile.exists()){
                  	  System.out.println(getProperties("homeDir")+System.getProperty("file.separator")+fname1[c]+" delete ??"+tmpFile.delete());                  	                    	
                  }else if(tmpFile.isDirectory()){//uploadDir�U���l�ؿ�
                  	  fname_sub.add(fname1[c]);//�O��uploadDir�U���l�ؿ�
                  	  System.out.println("fname_sub(top1)="+fname1[c]);
                  }
           }                      
		 }//end of uploadDir�ڥؿ��U���ɮ�  
		 //fname1 = null;
		 //97.09.22�R��uploadDir�U�l�ؿ����ɮ�
		 System.out.println("fname_sub.size()="+fname_sub.size());
		 for(int i=0;i<fname_sub.size();i++){
		 	 uploadFile = new File(getProperties("homeDir")+System.getProperty("file.separator")+fname_sub.get(i));
		 	 System.out.print("fname_sub(top1)["+i+"]="+getProperties("homeDir")+System.getProperty("file.separator")+fname_sub.get(i));
			 if(uploadFile.exists() && uploadFile.isDirectory()){
	           fname1= uploadFile.list(); //====�C�X���ؿ��U���Ҧ��ɮ�===================
	           System.out.println(":fname1.size="+fname1.length);
	           //�R��uploadDir�U�l�ؿ����ɮ� ex.bank�U���Ҧ��ɮ�.���]�A�ؿ�
	           for(int c=0;c<fname1.length;c++){
	    	       tmpFile = new File(getProperties("homeDir")+System.getProperty("file.separator")+fname_sub.get(i)+System.getProperty("file.separator")+fname1[c]);    	          
	               if(!tmpFile.isDirectory() && tmpFile.exists()){
	               	  System.out.println(getProperties("homeDir")+System.getProperty("file.separator")+fname_sub.get(i)+System.getProperty("file.separator")+fname1[c]+" delete ??"+tmpFile.delete());	               	  
	               }else if(tmpFile.isDirectory()){//uploadDir�U���l�ؿ�.���l�ؿ� ex:bank\about
	               	   fname_sub1.add(getProperties("homeDir")+System.getProperty("file.separator")+fname_sub.get(i)+System.getProperty("file.separator")+fname1[c]);//�O��uploadDir�U���l�ؿ�.���l�ؿ� ex:bank\about
	               	   System.out.println("fname_sub(top2)="+fname1[c]);
	               }
	           }                       
		     }
		 }
		 //fname1 = null;
		 //�R��uploadDir�U�l�ؿ����l�ؿ��ɮ�
		 for(int i=0;i<fname_sub1.size();i++){
			 uploadFile = new File((String)fname_sub1.get(i));
			 System.out.print("fname_sub(top2)["+i+"]="+(String)fname_sub1.get(i));
			 if(uploadFile.exists() && uploadFile.isDirectory()){
		        fname1= uploadFile.list(); //====�C�X���ؿ��U���Ҧ��ɮ�===================
		        System.out.println(":fname2.size="+fname1.length);
		        //�R��uploadDir�U�l�ؿ�.�l�ؿ����ɮ� ex.bank\about�U���Ҧ��ɮ�.���]�A�ؿ�
		        for(int c=0;c<fname1.length;c++){
		    	    tmpFile = new File(fname_sub1.get(i)+System.getProperty("file.separator")+fname1[c]);    	          
		            if(!tmpFile.isDirectory() && tmpFile.exists()){
		                System.out.println(fname_sub1.get(i)+System.getProperty("file.separator")+fname1[c]+" delete ??"+tmpFile.delete());                  	                    	
		            }
		        }                       
			}
		 }
		 /*
		 for(int i=1;i<=Integer.parseInt(getProperties("dir_count"));i++){		   	   
		   	   nowDir = getProperties("dir"+i);
		   	   if(!nowDir.equals("")){
		   	      uploadFile = new File(getProperties("homeDir")+System.getProperty("file.separator")+nowDir);		   	      
		   	      //System.out.print("deleteFile="+getProperties("homeDir")+System.getProperty("file.separator")+nowDir);		   	         	   
		   	      if(uploadFile.exists() && uploadFile.isDirectory()){
	                 fname1= uploadFile.list(); //====�C�X���ؿ��U���Ҧ��ɮ�===================
	                 System.out.println("fname1.length="+fname1.length);
	                 //�R��uploadDir.subdir�U���ɮ�
	                 for(int c=0;c<fname1.length;c++){
	            	     tmpFile = new File(getProperties("homeDir")+System.getProperty("file.separator")+nowDir+System.getProperty("file.separator")+fname1[c]);
	            	     System.out.println(getProperties("homeDir")+System.getProperty("file.separator")+nowDir+System.getProperty("file.separator")+fname1[c]);
	                     if(!tmpFile.isDirectory() && tmpFile.exists()){
	                     	System.out.println(" delete ??"+tmpFile.delete());
	                     }
	                 }	                 
		   	      }else{//end of �ӥؿ��s�b
		   	        System.out.println(getProperties("homeDir")+System.getProperty("file.separator")+nowDir+" has not exists");
		   	      }
		   	   }else{
		   	   	  System.out.println("dir"+i+" has not found");		   	      
		   	   }		   	   	   	   
		 }//end of for.uploadDir(�l�ؿ��U�������ɮ�)
		 */
		 result = true;
  	}catch(Exception e){
  		System.out.println(e+e.getMessage());  		
  	}
  	return result;
  }
}