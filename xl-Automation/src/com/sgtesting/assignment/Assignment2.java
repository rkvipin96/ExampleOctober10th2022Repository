package com.sgtesting.assignment;

import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Assignment2 {

	public static void main(String[] args) {
		FileOutputStream fout = null;
		Workbook wb = null;
		Sheet sh = null;
		Row row = null;
		Cell cell = null;
		
		try {
			wb = new XSSFWorkbook();
			sh = wb.createSheet("Flowernames");
			row = sh.createRow(9);
			for(int i = 0 ; i<20 ; i++)
			{
				cell=row.createCell(i);
				cell.setCellValue("Flower"+ i);
			}
			
			fout = new FileOutputStream("A:\\ExcelAutomation\\Assignment\\Assignment2.xlsx");
			wb.write(fout);
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		finally
		{
			try {
				fout.close();
				wb.close();
				
			} catch (Exception e2) {
				e2.printStackTrace();
			}
		}
		

	}

}
