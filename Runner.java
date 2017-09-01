
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;


public class Runner {

	public static void main(String[] args) throws InvalidFormatException {
		generatePDF("C:\\ELP\\ELP_Runs\\10109797_10031710_0731160740.xls");
	}
	public static void generatePDF(String xlsInput) throws InvalidFormatException{

		try {
			System.out.println(System.getenv("ProgramFiles(x86)"));
			String elpHome = System.getenv("ELP_Home");
			File file = new File(elpHome+"\\PlotJavaResultV6.xlsm");
			File inputXLS = new File(xlsInput);
			//XSSFSheet sheet;
	        FileInputStream filexl = new FileInputStream(file);
	   	 
	        //Create Workbook instance holding reference to .xls file
	        org.apache.poi.ss.usermodel.Workbook workbook = WorkbookFactory.create(filexl);
	        //XSSFWorkbook workbook = new XSSFWorkbook(filexl);

	        //Get desired sheet from the workbook
	        //sheet = workbook.getSheet("PlotsSheetSystem");
	        org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheet("PlotsSheetSystem");
			 
	        //Model Data same for all Revision
	        System.out.println(sheet.getRow(0).getCell(1).toString());
	         //sheet.getRow(0).getCell(1).setCellValue("E:\\Bala\\Rahul\\xls");
	         CellReference cr = new CellReference("A1");
	         Row row = sheet.getRow(cr.getRow());
	         Cell cell = row.getCell(cr.getCol());
	         cell.setCellValue(inputXLS.getParentFile().getAbsolutePath());
	         
	         CellReference b1 = new CellReference("B1");
	         Row b1Row = sheet.getRow(b1.getRow());
	         Cell b1Cell = b1Row.getCell(b1.getCol());
	         b1Cell.setCellValue(inputXLS.getAbsolutePath());
	         
	         System.out.println("cell+"+cell.getStringCellValue());
	         System.out.println( sheet.getRow(1).getCell(1)+"after");
	         filexl.close();
	         FileOutputStream f2 = new FileOutputStream(file);
	         workbook.write(f2);
	         f2.close();
			
			
			String dllPath = null;
			if(System.getenv("ProgramFiles(x86)") != null){
				dllPath = elpHome+"\\64bit";
			}else{
				dllPath = elpHome+"\\32bit";
			}
			System.out.println(dllPath);
			//Process process = Runtime.getRuntime().exec("java -cp P:\\self\\workspace\\MacroRunner\\bin\\;. com.elp.macro.MacroRunner");
			Process process = Runtime.getRuntime().exec("java -Djava.library.path="+dllPath+" -jar "+elpHome+"\\macro.jar");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	
	}

}
