package com.mavenProject;

import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class ExcelProject {

	public static void main(String[] args) throws Exception {
		FileInputStream in = new FileInputStream("Student1.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(in);// POI and POI-OOxml
		XSSFSheet sh =wb.getSheet("Sheet1");
		for(Row r : sh) 
		{
			
			for(Cell c : r) {
				if (c.getCellType()==CellType.NUMERIC) {
					System.out.print(c.getNumericCellValue() +"\t");
				
					
				}
				else if (c.getCellType()==CellType.STRING) {
					System.out.print(c.getStringCellValue() +"\t" );
				}
			
				
			}
			System.out.println();
			
		}
		wb.close();
		

	}

}
