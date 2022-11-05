

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Example {
	//to count no.of sheets

	public static void main(String[] args) throws IOException {
FileInputStream fis = new FileInputStream("ExcelTestData.xlsx");
		
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		
		int sheetCount = workbook.getNumberOfSheets();
		
		System.out.println(sheetCount);
		
	}

}
