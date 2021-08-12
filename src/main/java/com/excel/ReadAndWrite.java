package com.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadAndWrite {
	
	public static void main(String[] args) throws IOException {
		
		File f = new File ("C:\\Users\\thenswami\\eclipse-workspace\\BaratAb\\src\\main\\resources\\Source_excel\\BBF.xlsx");
		
		FileInputStream f1 = new FileInputStream (f);
		Workbook w = new XSSFWorkbook(f1);
		Sheet s = w.getSheet("Baru");
		
		for (int i = 1; i < s.getPhysicalNumberOfRows(); i++) {
			
			Row r = s.getRow(i);
			
			for (int j = 0; j < r.getPhysicalNumberOfCells(); j++) {
				Cell c = r.getCell(j);
				System.out.println(c);
			
			
		}
		
			
		}
		
		
		
		
		
	
		
		
	}

}
