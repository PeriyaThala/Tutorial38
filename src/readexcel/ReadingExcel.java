package readexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadingExcel {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		
//		Class 39 read in excel file 

		File f = new File("D:\\Tutor38\\src\\readexcel\\ReadExcel");
		FileInputStream fis = new FileInputStream(f);

		Workbook wb = WorkbookFactory.create(fis);
		Sheet sheet0 = wb.getSheetAt(0);

		/*
		 * Row r0 = sheet0.getRow(0); Cell c0 = r0.getCell(0);
		 * System.out.println("read successfully");
		 */
//		instead of above steps to manually get data - we can use "for each" loop

		for (Row r : sheet0) {
			for (Cell c : r) {
				switch (c.getCellType()) {
				case STRING:
//					System.out.println(c.getStringCellValue());
//					break;
					
//					instead of above -> if you use println it will print in line but to print in table use below format in all places
					System.out.print(c.getStringCellValue()+"  "); // instead of "  " you can also use "\t"
					break;
					
				case BOOLEAN:
					System.out.println(c.getBooleanCellValue());
					break;
				case NUMERIC:
					System.out.println(c.getNumericCellValue());
					break;
				default:
					break;
				}
			}
		}

		fis.close();

	}

}
