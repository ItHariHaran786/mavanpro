package org.maventask;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelREgForm {
public static void main(String[] args) throws IOException {
	File file = new File("C:\\Users\\sivasakthi\\eclipse-workspace\\maventask\\TestInfo\\Book2.xlsx");
	FileInputStream st = new FileInputStream(file);
	Workbook workbook = new XSSFWorkbook(st);
	Sheet sheet = workbook.getSheet("sheet1");
	for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
		Row row = sheet.getRow(i);
		for (int j = 0; j < sheet.getPhysicalNumberOfRows(); j++) {
			Cell cell = sheet.getCell();
			switch (cell.getCellType()) {
            case STRING:
                System.out.print(cell.getStringCellValue());
                break;
            case BOOLEAN:
                System.out.print(cell.getBooleanCellValue());
                break;
            case NUMERIC:
                System.out.print(cell.getNumericCellValue());
                break;
		
		
			
		}
			System.out.println(cell);
	}
	
}
}
}

