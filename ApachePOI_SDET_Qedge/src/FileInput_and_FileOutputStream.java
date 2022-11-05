import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FileInput_and_FileOutputStream {

	public static void main(String[] args) throws IOException {
		
		FileInputStream fi = new FileInputStream("testdata.xlsx");		
		Workbook wb = new XSSFWorkbook(fi);
		
		wb.createSheet("DemoSheet");
		
		FileOutputStream fo = new FileOutputStream("result.xlsx");
		wb.write(fo);
		
		wb.close();
		fi.close();
		fo.close();
		
		//----------------my prac
		FileInputStream fi9=new FileInputStream("");
		Workbook book=new XSSFWorkbook(fi9);
		Sheet sheet = book.getSheet("sheet1");
		int row_count=sheet.getLastRowNum();
		int cell_count = sheet.getRow(1).getLastCellNum();
		Cell cell = sheet.getRow(1).getCell(0);
		String cell_data = cell.getStringCellValue();
		System.out.println(cell_data);
		
		for(int i=0;i<row_count;i++)
		{
			Cell cell1 = sheet.getRow(i).getCell(0);
			Cell cell2 = sheet.getRow(i).getCell(1);
			Cell cell3 = sheet.getRow(i).getCell(2);
			String data1 = cell1.getStringCellValue();
			boolean data2 = cell2.getBooleanCellValue();
			double data3 = cell3.getNumericCellValue();
			System.out.println(data1+"   "+data2+"    "+data3);
		}
		book.close();
		fi9.close();
		
		/*
		Script to Count No. of Rows in XlSheet
		-------------------------------------------------------------------------------*/
		
		FileInputStream fi1 = new FileInputStream("TestData.xlsx");
				Workbook wb1 = new XSSFWorkbook(fi1);
				Sheet ws1 = wb.getSheet("LoginData");
				Sheet ws2 = wb1.getSheet("EmpData");
				
				int sheet1_rowcount = ws1.getLastRowNum();
				int sheet2_rowcount = ws2.getLastRowNum();
				
				System.out.println(sheet1_rowcount);
				System.out.println(sheet2_rowcount);
				
				wb1.close();
				fi1.close();
		/*------------------------------------------------------------------------------
		Script to Count No. of Columns in a XlSheet Row
		------------------------------------------------------------------------------*/
				FileInputStream fi2 = new FileInputStream("TestData.xlsx");
				Workbook wb2 = new XSSFWorkbook(fi2);
				Sheet ws = wb2.getSheet("LoginData");
				Row r = ws.getRow(0);
				short colcount = r.getLastCellNum();
				System.out.println(colcount);
				wb2.close();
				fi2.close();
		/*-----------------------------------------------------------------------------
		Script to Read data from XlSheet Cells
		-----------------------------------------------------------------------------*/
		FileInputStream fi3 = new FileInputStream("TestData.xlsx");
				Workbook wb3 = new XSSFWorkbook(fi3);
				Sheet ws3 = wb3.getSheet("EmpData");
				
				Row r3 =  ws.getRow(1);
				Cell c1 = r3.getCell(1);
				Cell c2 = r3.getCell(2);
				Cell c3 = r3.getCell(3);
				
				String empname =  c1.getStringCellValue();
				double salary =  c2.getNumericCellValue();
				boolean status = c3.getBooleanCellValue();
				
				System.out.println(empname+"  "+salary+"  "+status);
				
				wb3.close();
				fi3.close();
				
		/*------------------------------------------------------------------------------
		Script to Read all Rows of data present in a XLSheet
		---------------------------------------------------------------------------*/
				FileInputStream fi4 = new FileInputStream("TestData.xlsx");
				Workbook wb4 = new XSSFWorkbook(fi4);
				Sheet ws4 = wb4.getSheet("EmpData");
				
				int rowcount = ws4.getLastRowNum();
				Row row;
				Cell c11,c21,c31;
						
				for(int i=1;i<=rowcount;i++)
				{
					row = ws.getRow(i);
					c11 = row.getCell(1);
					c21= row.getCell(2);
					c31 = row.getCell(3);
					
					String empname1 = c11.getStringCellValue();
					double sal = c21.getNumericCellValue();
					boolean status1 = c31.getBooleanCellValue();
					System.out.println(empname1+"  "+sal+"   "+status1);
				}
						
				wb4.close();
				fi4.close();
	
}
}
