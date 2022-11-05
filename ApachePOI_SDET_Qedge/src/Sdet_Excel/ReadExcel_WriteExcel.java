package Sdet_Excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel_WriteExcel {

	public static void main(String[] args) throws IOException {
		//read data from excel
		FileInputStream fi=new FileInputStream("TestData.xlsx");
		Workbook book=new XSSFWorkbook(fi);
		Sheet sheet = book.getSheet("AdminLoginValidData"); //or book.getSheetAt(0);
		
		int rows = sheet.getLastRowNum();
		int cells = sheet.getRow(1).getLastCellNum();
		
		for(int i=0;i<=rows;i++)
		{
			for(int j=0;j<cells;j++)
			{
				Cell cell = sheet.getRow(i).getCell(j);
				switch(cell.getCellType())
				{
				case STRING: System.out.println(cell.getStringCellValue()); break;
				case NUMERIC:System.out.println(cell.getNumericCellValue()); break;
				case BOOLEAN:System.out.println(cell.getBooleanCellValue()); break;
			
				}
				System.out.print("|");
			}
			System.out.println();
		}
		
		//write data into excel
		Workbook book1=new XSSFWorkbook();
		Sheet sheet1 = book1.createSheet("Datasheet");
		
		//creating data
		Object[][]  empdata= {{"EmpID","Name","Job"},{101,"Yashu","Tester"},{102,"Abhi","Consultant"},{103,"Mahesh","Hero"}};
		int rows1 = empdata.length;
		int cells1=empdata[0].length;
		
		for(int i=0;i<rows1;i++)
		{
			Row row1 = sheet1.createRow(i);
			for(int j=0;j<cells1;j++)
			{
				Cell cell1 = row1.createCell(j);
				Object value = empdata[i][j];
				if(value instanceof String)
					cell1.setCellValue((String)value);
				if(value instanceof Integer)
					cell1.setCellValue((Integer)value);
				if(value instanceof Boolean)
					cell1.setCellValue((Boolean)value);
			}
		}
		FileOutputStream fo=new FileOutputStream("Employee.xlsx");
		book1.write(fo);
		book1.close();
		fo.close();
		
		

	}

}
