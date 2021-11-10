package org.sel;

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

public class Write {

	public static void main(String[] args) throws IOException {
		
		File f = new File("C:\\Users\\RAM SARATH KUMAR\\Downloads\\ADACTIN.xlsx");
		
		FileInputStream str = new FileInputStream(f);
		
		Workbook w = new XSSFWorkbook(str);
		
		Sheet createSheet = w.createSheet("dai");
		
		Row createRow = createSheet.createRow(1);
		
		
		Cell createCell = createRow.createCell(1);
		
		createCell.setCellValue("Java");
		
		
		FileOutputStream stre = new FileOutputStream(f);
		
		System.out.println("done");
	}
	
	
	
}
