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

public class TestDataExcel {
public static void main(String[] args) throws IOException {
	File file = new File("C:\\Users\\sivasakthi\\eclipse-workspace\\maventask\\TestInfo\\Book1.xlsx");
	FileInputStream stream = new FileInputStream(file);
	Workbook workbook = new XSSFWorkbook(stream);
	Sheet sheet = workbook.getSheet("sheet1");
	Row row = sheet.getRow(1);
	Cell cell = row.getCell(1);
	String sc = cell.getStringCellValue();
	System.out.println(sc);
	Row row2 = sheet.getRow(2);
	Cell cell2 = row2.getCell(1);
	double nc = cell2.getNumericCellValue();
	System.out.println(nc);
}
}
