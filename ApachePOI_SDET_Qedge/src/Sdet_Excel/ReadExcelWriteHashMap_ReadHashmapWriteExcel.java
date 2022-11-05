package Sdet_Excel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelWriteHashMap_ReadHashmapWriteExcel {

	public static void main(String[] args) throws IOException {
		

		//Read from Excel and write into Hashmap
		FileInputStream fi=new FileInputStream("TestData.xlsx");
		Workbook book=new XSSFWorkbook(fi);
		Sheet sheet = book.getSheet("AdminLoginValidData");
		int rows = sheet.getLastRowNum();
		
		
		HashMap<String,String> data=new HashMap<String,String> ();
		
		//reading data from excel to hashmap
		for(int i=0;i<=rows;i++)
		{
		String key = sheet.getRow(i).getCell(0).getStringCellValue();
		String value = sheet.getRow(i).getCell(1).getStringCellValue();
		data.put(key, value);
		}
         //reading hashmap data
		for(Map.Entry entry:data.entrySet())
		{
		System.out.println( entry.getKey()+"  "+entry.getValue());	
		}
		
		
		//Read data from Hashmap and write into Excel
		Workbook book1=new XSSFWorkbook();
		Sheet sheet1 = book1.createSheet("data");
		
		
		Map<String,String> data1=new HashMap<String,String>();
		data1.put("101", "John");
		data1.put("102", "David");
		data1.put("103", "Scot");
		data1.put("104", "Mary");
		
		int rowno=0;
		for(Map.Entry entry1:data1.entrySet())
		{
			Row row1=sheet1.createRow(rowno++);
			row1.createCell(0).setCellValue((String)entry1.getKey());
			row1.createCell(1).setCellValue((String)entry1.getValue());
		}
		
		FileOutputStream fo=new FileOutputStream("Student.xlsx");
		book1.write(fo);
		fo.close();
		System.out.println("Excel file written successfully");
	}



}
