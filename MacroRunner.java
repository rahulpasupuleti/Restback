package com.elp.macro;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class MacroRunner {

	/**
	 * @param args
	 * @throws IOException 
	 * @throws InvalidFormatException 
	 */
	public static void main(String[] args) throws IOException, InvalidFormatException {
		String elpHome = System.getenv("ELP_Home");
		File file = new File(elpHome+"\\PlotJavaResultV8.xlsm");
        //String macroName = "Macro";
        //callExcelMacro(file, macroName);
        System.out.println("done");
        String macroName2 = "uploadFile";
        callExcelMacro(file, macroName2);
        System.out.println("done2");
        String macroName3 = "PrintPlotResult";
        callExcelMacro(file, macroName3);
        System.out.println("Completed");
	}

	private static void callExcelMacro(File file, String macroName) {
        ComThread.InitSTA(true);
        final ActiveXComponent excel = new ActiveXComponent("Excel.Application");
        try{
            excel.setProperty("EnableEvents", new Variant(false));

            Dispatch workbooks = excel.getProperty("Workbooks")
                    .toDispatch();

            Dispatch workBook = Dispatch.call(workbooks, "Open",
                    file.getAbsolutePath()).toDispatch();

            // Calls the macro
            //Variant V1 = new Variant("\'"+file.getName()+"\'"+ macroName);
            Variant V1 =  new Variant(macroName);
            //Variant V1 = new Variant( file.getName() + macroName);
            Variant result = Dispatch.call(excel, "Run", V1);

            // Saves and closes
            Dispatch.call(workBook, "Save");

            com.jacob.com.Variant f = new com.jacob.com.Variant(true);
            Dispatch.call(workBook, "Close", f);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            excel.invoke("Quit", new Variant[0]);
            ComThread.Release();
        }
    }
}
