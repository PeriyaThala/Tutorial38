package Writeexcel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
//import org.apache.poi.ss.usermodel.Cell;
//import org.apache.poi.ss.usermodel.Row;

public class writing {

	public static void main(String[] args) throws IOException {
//		Apache poi is a library used to read and write data in excel files
//		We have to configure build path now to add appache poi 
//		steps -> folder right click -> build path -> configure build path -> libraries -> class path -> add external jars -> apply and close
//		we have to manually create workbook and sheets initially 
		
//		Step 1:
		
		HSSFWorkbook wbb = new HSSFWorkbook();
		HSSFSheet Sheet1 = wbb.createSheet("SheetName");
		HSSFRow r0 = Sheet1.createRow(0);
		HSSFCell c0 = r0.createCell(0);
		c0.setCellValue("VAnShika");
		
//		Step 2:
		
		File f = new File("D:\\Tutor38\\src\\readexcel\\readexcelfile.xls");
		
//	    Step 3:
		
		FileOutputStream fos = new FileOutputStream(f);
		wbb.write(fos);
		fos.close();
		wbb.close();
		System.out.println("written successfully");
		

	}

}
