import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FillColor_Createcell {

	public static void main(String[] args) throws IOException {
		/*-----------------------------------------------------------------------------
		Script to handle null pointer exception that occurs when no data present in
		XlSheet
		-----------------------------------------------------------------------------
		*/FileInputStream fi = new FileInputStream("TestData.xlsx");
				Workbook wb = new XSSFWorkbook(fi);
				Sheet ws = wb.getSheet("EmpData");
				Row r = ws.getRow(1);
				String data;
				try {
					data = r.getCell(0).getStringCellValue();
					System.out.println(data);
				}catch (Exception e) {
					data="";
					System.out.println("data not found");
				}
				
				/*String data;
				try 
				{
					Cell c = r.getCell(1);
					data = c.getStringCellValue();
					System.out.println(data);
				} catch (Exception e) 
				{
					data = "";
					System.out.println("No data found!");
				}
*/		/*----------------------------------------------------------------------------
		Script to write data into XLSheet Cells
		----------------------------------------------------------------------------
		*/		FileInputStream fi2 = new FileInputStream("TestData.xlsx");
				Workbook wb2 = new XSSFWorkbook(fi2);
				Sheet ws2 = wb2.getSheet("EmpData");
				Row r2 = ws2.getRow(1);
		
				
				Cell c = r.createCell(4);
				c.setCellValue("Pass");
				
				FileOutputStream fo = new FileOutputStream("Result.xlsx");
				wb.write(fo);
				wb.close();
				fi.close();
				fo.close();
		/*----------------------------------------------------------------------------
		Script to fill XlSheet Cell Color with Green
		-----------------------------------------------------------------------------
			*/	FileInputStream fi3 = new FileInputStream("TestData.xlsx");
				Workbook wb3 = new XSSFWorkbook(fi3);
				Sheet ws3 = wb.getSheet("EmpData");
				Row r3 = ws.getRow(1);
				Cell c3 = r.createCell(4);
				c.setCellValue("Fail");
				
				CellStyle style = wb.createCellStyle();
				style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
				style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				c.setCellStyle(style);
				
				
				FileOutputStream fo3 = new FileOutputStream("Result.xlsx");
				wb.write(fo);
				wb.close();
				fi.close();
				fo.close();

		/*----------------------------------------------------------------------------------
		Script to fill XLSheet Cell color with Red
		-----------------------------------------------------------------------------
		*/FileInputStream fi4 = new FileInputStream("TestData.xlsx");
				Workbook wb4 = new XSSFWorkbook(fi);
				Sheet ws4 = wb.getSheet("EmpData");
				Row r4 = ws.getRow(1);
				Cell c4= r.createCell(4);
				c.setCellValue("Fail");
				
				CellStyle style4 = wb.createCellStyle();
				style.setFillForegroundColor(IndexedColors.RED.getIndex());
				style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				c.setCellStyle(style);
				
				
				FileOutputStream fo4 = new FileOutputStream("Result.xlsx");
				wb.write(fo);
				wb.close();
				fi.close();
				fo.close();

	


	}

}
