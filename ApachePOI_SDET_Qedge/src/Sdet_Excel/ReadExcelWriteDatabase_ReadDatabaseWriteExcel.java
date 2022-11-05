package Sdet_Excel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelWriteDatabase_ReadDatabaseWriteExcel {

	public static void main(String[] args) throws IOException, SQLException {
		
		//	read  from excel and write into database
		Connection con = DriverManager.getConnection("jdbc:mysql://localhost:3306/world", "root", "12R61a0574@");

		Statement stmt = con.createStatement();

		stmt.execute("create table places16(ID int,Name varchar(25),CountryCode varchar(25),District varchar(25),Population int);");
	
		FileInputStream fi=new FileInputStream("city.xlsx");
		Workbook book=new XSSFWorkbook(fi);
		Sheet sheet = book.getSheet("City data");
		int rows = sheet.getLastRowNum();
		
		
		for(int i=0;i<=rows;i++)
		{
			DataFormatter formatter=new DataFormatter();
			 String city_id=formatter.formatCellValue(sheet.getRow(i).getCell(0));
			 String name= formatter.formatCellValue(sheet.getRow(i).getCell(1));
			 String c_code=formatter.formatCellValue(sheet.getRow(i).getCell(2));
			 String dist=formatter.formatCellValue(sheet.getRow(i).getCell(3));
			 String pop=formatter.formatCellValue(sheet.getRow(i).getCell(4));
			     /*
			    String city_id = sheet.getRow(i).getCell(0).getStringCellValue();
				String name=sheet.getRow(i).getCell(1).getStringCellValue();
				String c_code=sheet.getRow(i).getCell(2).getStringCellValue();
				String dist=sheet.getRow(i).getCell(3).getStringCellValue();
				String pop=sheet.getRow(i).getCell(4).getStringCellValue();
			    */
				stmt.execute("insert into places values('"+city_id+"','"+name+"','"+c_code+"','"+dist+"','"+pop+"')");
		
		
		}
		System.out.println("Done");
			book.close();
			fi.close();
			con.close();
			
			
			
			//	read data from database and write into Excel
			Connection con1 = DriverManager.getConnection("jdbc:mysql://localhost:3306/world", "root", "12R61a0574@");

			Statement stmt1 = con1.createStatement();

			ResultSet rs = stmt1.executeQuery("select * from city");
			
			Workbook book1=new XSSFWorkbook();
			Sheet sheet1 = book1.createSheet("City data");
			
			Row row = sheet1.createRow(0);
			row.createCell(0).setCellValue("ID");
			row.createCell(1).setCellValue("Name");
			row.createCell(2).setCellValue("CountryCode");
			row.createCell(3).setCellValue("District");
			row.createCell(4).setCellValue("Population");
			
			int r=1;
			while(rs.next())
			{
				double city_id = rs.getDouble("ID");
				String name = rs.getString("Name");
				String ccode = rs.getString("CountryCode");
				String dist = rs.getString("District");
				double pop = rs.getDouble("Population");
				
				row=sheet1.createRow(r++);
			    row.createCell(0).setCellValue(city_id);
			    row.createCell(1).setCellValue(name);
			    row.createCell(2).setCellValue(ccode);
			    row.createCell(3).setCellValue(dist);
			    row.createCell(4).setCellValue(pop);
			    
			}
			FileOutputStream fo=new FileOutputStream("city.xlsx");
			book1.write(fo);
			fo.close();
			System.out.println("Excel file written successfully");

			con.close();
		}
		
		
	}


