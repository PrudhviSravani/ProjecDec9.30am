package org.test.cts;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Example {
	public static void main(String[] args) throws IOException {
		File exceLoc = new File("C:\\Users\\Dell\\Desktop\\Prudhvi\\CompanyDetails\\Excel\\TestData 1.xlsx");
		FileInputStream stream = new FileInputStream(exceLoc);
		Workbook w = new XSSFWorkbook(stream);
		Sheet s = w.getSheet("Datas");
		
		for (int j = 0; j < s.getPhysicalNumberOfRows(); j++) {
			Row r = s.getRow(j);
			for (int i = 0; i <r.getPhysicalNumberOfCells(); i++) {
				Cell c = r.getCell(i);
				System.out.println(c);
				
			}
			
		}
			
		}
		
		
	
		

}
