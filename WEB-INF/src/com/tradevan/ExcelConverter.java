package com.tradevan;


import java.io.File;

//import org.slf4j.Logger;
//import org.slf4j.LoggerFactory;
//import RptConvertConstants.ConvertConstants;

import com.artofsolving.jodconverter.DocumentConverter;
import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.OpenOfficeDocumentConverter;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class ExcelConverter {
	
	public void convertExceltoPdf(String inputFilePath , String outputFilePath) {
		ComThread.InitSTA(true);
		ActiveXComponent ax = new ActiveXComponent("Excel.Application");
		try{
			ax.setProperty("Visible", new Variant(false));
			ax.setProperty("AutomationSecurity", new Variant(3));
			
			Dispatch excels = ax.getProperty("Workbooks").toDispatch();
			Dispatch excel = Dispatch.invoke(excels , "Open" , Dispatch.Method , new Object[] {
				inputFilePath , new Variant(false) , new Variant(false) } , new int[9]).toDispatch();
			Dispatch.invoke(excel , "ExportAsFixedFormat" , Dispatch.Method , new Object[] {
					new Variant(0) , outputFilePath , new Variant(0) } , new int[1]);
			Dispatch.call(excel , "Close" , new Variant(false));
		} catch(Exception e) {
			e.printStackTrace();
		} finally {
			if(ax != null) {
				ax.invoke("Quit",new Variant[]{});
				ax = null;
			}
			ComThread.Release();
		}
	}

	public void convertExceltoPdf(String inputFilePath , String outputFilePath) {
		ComThread.InitSTA(true);
		ActiveXComponent ax = new ActiveXComponent("Excel.Application");
		try{
			ax.setProperty("Visible", new Variant(false));
			ax.setProperty("AutomationSecurity", new Variant(3));

			Dispatch excels = ax.getProperty("Workbooks").toDispatch();
			Dispatch excel = Dispatch.invoke(excels , "Open" , Dispatch.Method , new Object[] {
					inputFilePath , new Variant(false) , new Variant(false) } , new int[9]).toDispatch();
			Dispatch.invoke(excel , "ExportAsFixedFormat" , Dispatch.Method , new Object[] {
					new Variant(0) , outputFilePath , new Variant(0) } , new int[1]);
			Dispatch.call(excel , "Close" , new Variant(false));
		} catch(Exception e) {
			e.printStackTrace();
		} finally {
			if(ax != null) {
				ax.invoke("Quit",new Variant[]{});
				ax = null;
			}
			ComThread.Release();
		}
	}

	public void convertExceltoOds(String inputFilePath , String outputFilePath) {
		OpenOfficeConnection connection = new SocketOpenOfficeConnection("127.0.0.1" , 8100);
		try {
			File inputFile = new File(inputFilePath);
			File outputFile = new File(outputFilePath);
			connection.connect();
			DocumentConverter converter = new OpenOfficeDocumentConverter(connection);
			converter.convert(inputFile, outputFile);
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (connection.isConnected()){
				connection.disconnect();
				System.out.println("openOffice disconnect");
			}
		}
	}
}
