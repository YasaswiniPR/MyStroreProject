package Sdet_Excel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Xlutils {

	public static FileInputStream fi;
	public static FileOutputStream fo;
	public static Workbook book;
	public static Sheet sheet;
	public static Row row; 
	public static Cell cell; 
	public static CellStyle style;
	
 	public static int getRowCount(String xlfile,String xlsheet) throws IOException
	{
		fi = new FileInputStream(xlfile);
		book = new XSSFWorkbook(fi);
		sheet = book.getSheet(xlsheet);
		int rowcount = sheet.getLastRowNum();
		book.close();
		fi.close();
		return rowcount;		
	}	
		
	public static short getCellCount(String xlfile,String xlsheet,int rownum) throws IOException
	{
		fi = new FileInputStream(xlfile);
		book = new XSSFWorkbook(fi);
		sheet = book.getSheet(xlsheet);
		row = sheet.getRow(rownum);
		short cellcount = row.getLastCellNum();
		book.close();
		fi.close();
		return cellcount;		
	}
	
	public static String getCellData(String xlfile,String xlsheet,int rownum,int colnum) throws IOException
	{
		fi = new FileInputStream(xlfile);
		book = new XSSFWorkbook(fi);
		sheet = book.getSheet(xlsheet);
		row = sheet.getRow(rownum);
        cell = row.getCell(colnum);
		
        DataFormatter formatter=new DataFormatter();
		String data;
		try 
		{   //return data from the cell in string format
			data = formatter.formatCellValue(cell);//returns the formatted value of a cell as a string regardless of cell type
		} catch (Exception e) 
		{
			data = "";
			System.out.println("No Data Found!");
		}
		book.close();
		fi.close();
		return data;
	}
	
	
	
	public static void setCellData(String xlfile,String xlsheet,int rownum,int colnum,String data) throws IOException
	{
		fi = new FileInputStream(xlfile);
		book = new XSSFWorkbook(fi);
		sheet = book.getSheet(xlsheet);
		row = sheet.getRow(rownum);
		cell = row.createCell(colnum);
		cell.setCellValue(data);
		fo = new FileOutputStream(xlfile);
		book.write(fo);
		book.close();
		fi.close();
		fo.close();
	}
	
	public static void fillGreenColor(String xlfile,String xlsheet,int rownum,int colnum) throws IOException
	{
		fi = new FileInputStream(xlfile);
		book = new XSSFWorkbook(fi);
		sheet = book.getSheet(xlsheet);
		row = sheet.getRow(rownum);
		cell = row.getCell(colnum);
		
		style = book.createCellStyle();
		style.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);		
		cell.setCellStyle(style);
		
		fo = new FileOutputStream(xlfile);
		book.write(fo);
		book.close();
		fi.close();
		fo.close();				
	}
	
	public static void fillRedColor(String xlfile,String xlsheet,int rownum,int colnum) throws IOException
	{
		fi = new FileInputStream(xlfile);
		book = new XSSFWorkbook(fi);
		sheet = book.getSheet(xlsheet);
		row = sheet.getRow(rownum);
		cell = row.getCell(colnum);
		
		style = book.createCellStyle();
		style.setFillForegroundColor(IndexedColors.RED.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);		
		cell.setCellStyle(style);
		
		fo = new FileOutputStream(xlfile);
		book.write(fo);
		book.close();
		fi.close();
		fo.close();				
	}
	

}
