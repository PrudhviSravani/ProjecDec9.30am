
 package org.test.cts;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FaceBook {
	public static void main(String[] args) throws IOException {
		File excelLoc = new File("C:\\Users\\Dell\\Desktop\\Prudhvi\\CompanyDetails\\Excel\\Test Data.xlsx");
		FileInputStream stream = new FileInputStream(excelLoc);
		Workbook w = new XSSFWorkbook(stream);
		Sheet s = w.getSheet("sheet1");
		
		for (int i = 0; i < s.getPhysicalNumberOfRows(); i++) {
			Row r = s.getRow(i);
			for (int j = 0; j < r.getPhysicalNumberOfCells(); j++) {
				Cell c = r.getCell(j);
				System.out.println(c);
				
			}
			
		}
		
		
	}
	

}

