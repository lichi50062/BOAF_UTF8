package com.tradevan.util.transfer;

import java.io.File;


import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;


public class WordConverter {
	
	public static void main(String[] args){
		WordConverter t = new WordConverter();
		
		if(args[0].equalsIgnoreCase("odt")){
		t.convertWordToOdt("C:\\Sun\\WebServer6.1\\MIS\\reportDir\\農會各項法定比率表.rtf","C:\\Sun\\WebServer6.1\\MIS\\reportDir\\農會各項法定比率表.odt");
		}
		if(args[0].equalsIgnoreCase("pdf")){
			t.convertWordToOdt("C:\\Sun\\WebServer6.1\\MIS\\reportDir\\農會各項法定比率表.rtf","DC:\\Sun\\WebServer6.1\\MIS\\reportDir\\農會各項法定比率表.pdf");
	   }
	}
	
	public void convertWordToOdt(String inputFilePath , String outputFilePath) {
		ComThread.InitSTA(true);  
		ActiveXComponent app = null;
		Dispatch doc = null;
		try{
			app = new ActiveXComponent("Word.Application");
			app.setProperty("Visible", new Variant(false));

			Dispatch docs = app.getProperty("Documents").toDispatch();
			
			doc = Dispatch.invoke(docs,"Open",Dispatch.Method,new Object[] { 
					inputFilePath , new Variant(false),new Variant(true) }, new int[1]).toDispatch(); 
			System.out.println("test1");
			File tofile = new File(outputFilePath);
			if (tofile.exists()) {
				tofile.delete();
			}
			System.out.println("test2");
			//Dispatch.invoke(doc, "SaveAs", Dispatch.Method, new Object[] { outputFilePath, new Variant(17) }, new int[1]);//wordtopdf
			Dispatch.invoke(doc, "SaveAs", Dispatch.Method, new Object[] { outputFilePath, new Variant(23) }, new int[1]);//wordtoodt
			System.out.println("test3");
		} catch(Exception e) {
			e.printStackTrace();
		} finally {
			Dispatch.call(doc,"Close",new Variant(false)); 
			System.out.println("test4");
			if (app != null) {
				app.invoke("Quit", new Variant[] {});
				System.out.println("test5");
			}
            //關閉winword.exe
			ComThread.Release();
			System.out.println("test8");
		}
	}
	
	public void convertWordToPdf(String inputFilePath , String outputFilePath) {
		ComThread.InitSTA(true);  
		ActiveXComponent app = null;
		Dispatch doc = null;
		try{
			app = new ActiveXComponent("Word.Application");
			app.setProperty("Visible", new Variant(false));

			Dispatch docs = app.getProperty("Documents").toDispatch();
			
			doc = Dispatch.invoke(docs,"Open",Dispatch.Method,new Object[] { 
					inputFilePath , new Variant(false),new Variant(true) }, new int[1]).toDispatch(); 
			System.out.println("test1");
			File tofile = new File(outputFilePath);
			if (tofile.exists()) {
				tofile.delete();
			}
			System.out.println("test2");
			Dispatch.invoke(doc, "SaveAs", Dispatch.Method, new Object[] { outputFilePath, new Variant(17) }, new int[1]);//wordtopdf
			
			System.out.println("test3");
		} catch(Exception e) {
			e.printStackTrace();
		} finally {
			Dispatch.call(doc,"Close",new Variant(false)); 
			System.out.println("test4");
			if (app != null) {
				app.invoke("Quit", new Variant[] {});
				System.out.println("test5");
			}
            //關閉winword.exe
			ComThread.Release();
			System.out.println("test8");
		}
	}
}
