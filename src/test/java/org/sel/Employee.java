package org.sel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Employee {

	public static void main(String[] args) throws IOException {
		
	File f = new File("C:\\Users\\RAM SARATH KUMAR\\Downloads\\ADACTIN.xlsx");
	
	FileInputStream stream = new FileInputStream(f);
	
	Workbook w = new XSSFWorkbook(stream);
	
	Sheet sheet = w.getSheet("Sheet1");
	
	for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
		
		Row row = sheet.getRow(i);
		
		for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
			Cell cell = row.getCell(j);
			int cellType = cell.getCellType();
			
			if(cellType==1) {
				String stringCellValue = cell.getStringCellValue();
				
				System.out.println(stringCellValue);
				
			}
			
			else if(DateUtil.isCellDateFormatted(cell)) {
				
				Date dateCellValue = cell.getDateCellValue();
				
				SimpleDateFormat format = new SimpleDateFormat("dd-mm-yyyy");
				
				String format2 = format.format(dateCellValue);
				
				System.out.println(format2);
			}
			
			else {
				
				double numericCellValue = cell.getNumericCellValue();
				
				long l = (long)numericCellValue;
				
				System.out.println(l);
			}
			
		}
	}
	
	
	
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
	}
	
}
