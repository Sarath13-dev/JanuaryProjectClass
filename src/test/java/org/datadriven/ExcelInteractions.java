package org.datadriven;

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

public class ExcelInteractions {
public static void main(String[] args) throws IOException {
	//create object for File
	File f = new File("D:\\GREENS\\mor30.xlsx");
	
	//to read data
	FileInputStream stream = new FileInputStream(f);
	System.out.println("work done by perf");
	//create object for Workbook
	Workbook w = new XSSFWorkbook(stream);
	
	//to get Sheet from Workbook
	Sheet sheet = w.getSheet("abcd");
	
	//to find how many number of rows filled with data
	//int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
	
	//to iterate or seperate each row
	for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
		Row row = sheet.getRow(i);
		
		//to find how many cells in a row filled with data
		for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
			Cell cell = row.getCell(j);
			//to determine type of data stored in cell
			int cellType = cell.getCellType();
			if(cellType==1) {
				String stringCellValue = cell.getStringCellValue();
				System.out.println(stringCellValue);
			}
			else if(DateUtil.isCellDateFormatted(cell)) {
				Date dateCellValue = cell.getDateCellValue();
				System.out.println(dateCellValue);
				//to change date format
				SimpleDateFormat s = new SimpleDateFormat("MMM/dd/yy");
				String format = s.format(dateCellValue);
				System.out.println(format);
			}
			
			
			else {
				double numericCellValue = cell.getNumericCellValue();
				//type conversion
				long l = (long)numericCellValue;
				System.out.println(l);
			}
		}
		
	}
	
	
	
}
}
