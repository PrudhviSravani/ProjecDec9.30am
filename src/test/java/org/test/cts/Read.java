package org.test.cts;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Read {
	public static void main(String[] args) throws IOException {
		File exceLoc = new File("C:\\Users\\Dell\\Desktop\\Prudhvi\\CompanyDetails\\Excel\\hjk.xlsx");
		FileInputStream stream = new FileInputStream(exceLoc);
		Workbook w = new XSSFWorkbook(stream);
		Sheet s = w.getSheet("sheet1");
		
		for (int j = 0; j < s.getPhysicalNumberOfRows(); j++) {
			Row r = s.getRow(j);
			for (int i = 0; i <r.getPhysicalNumberOfCells(); i++) {
				Cell c = r.getCell(i);
				int type =c.getCellType();
				
				//1==1
				if (type==1) {
					String StringCellValue =c.getStringCellValue();
					System.out.println(StringCellValue);
				
				}
				//0==0
				else if(type==0) {
					//checking whether the cell is a date or numeric once
					if(DateUtil.isCellDateFormatted(c)) {
					Date dateCellValue =c.getDateCellValue();
					//converting date in to string
					SimpleDateFormat sim = new SimpleDateFormat("dd-MMM-yy");
					String format = sim.format(dateCellValue);
					System.out.println(format);
				}
				
				else {
					double numericCellValue = c.getNumericCellValue();
					//converting double into long
					//typecasting
					long l = (long)numericCellValue;
					//converting long into string
					String v= String.valueOf(l);
					System.out.println(v);
					
					
				}
					
				}	
				}
				}
			}
			}
			
		