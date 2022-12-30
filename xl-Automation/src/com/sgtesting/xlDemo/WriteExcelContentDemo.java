package com.sgtesting.xlDemo;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class WriteExcelContentDemo {

	public static void main(String[] args) {
		FileOutputStream fout = null;
		Workbook wb = null;
		Sheet sh = null;
		Row row = null;
		Cell cell = null;
		
		try
		{
			wb = new XSSFWorkbook();
			sh=wb.createSheet("Information");
			row = sh.createRow(0);
			cell = row.createCell(0);
			cell.setCellValue("Username");
			cell=row.createCell(1);
			cell.setCellValue("Password");
			
			row=sh.createRow(1);
			cell=row.createCell(0);
			cell.setCellValue("admin");
			cell=row.createCell(1);
			cell.setCellValue("manager");
			
			fout = new FileOutputStream("A:\\ExcelAutomation\\Data.xlsx");
			wb.write(fout);
			
		}catch(Exception e)
		{
			e.printStackTrace();
		}
		finally {
			try {
				fout.close();
				wb.close();
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		
		

	}

}
