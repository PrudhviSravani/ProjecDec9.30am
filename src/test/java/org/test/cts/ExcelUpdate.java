package org.test.cts;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUpdate {
	public static void main(String[] args) throws IOException {
		File loc = new File ("C:\\Users\\Dell\\Desktop\\Prudhvi\\CompanyDetails\\Excel\\srav.xlsx");
		FileInputStream stream = new FileInputStream(loc);
		Workbook w = new XSSFWorkbook(stream);
		Sheet s = w.getSheet("sheet0");
        Row r =s.getRow(5);
		Cell c = r.getCell(5);
		String s1 = c.getStringCellValue();
		if(s1.equals("Nishanthi")) {
		c.setCellValue("rahavana");
		}
		FileOutputStream o = new FileOutputStream(loc);
		w.write(o);
		System.out.println("Updated Suceesfully");
		
		
		
		
	}
}


