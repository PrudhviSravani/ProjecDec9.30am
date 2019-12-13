package org.test.cts;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelCreate {
	public static void main(String[] args) throws IOException {
		File loc = new File ("C:\\Users\\Dell\\Desktop\\Prudhvi\\CompanyDetails\\Excel\\srav1234.xlsx");
		Workbook w = new XSSFWorkbook();
		Sheet s = w.createSheet("srav");
	    Row r =s.createRow(5);
		Cell c = r.createCell(4);
	    c.setCellValue("nishanthi");
		FileOutputStream o = new FileOutputStream(loc);
		w.write(o);
		System.out.println("wrote Suceesfully");
		
		
		
		
	}
}
