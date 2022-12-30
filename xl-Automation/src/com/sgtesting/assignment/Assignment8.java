package com.sgtesting.assignment;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Assignment8 {

	public static void main(String[] args) {
		FileInputStream fin = null;
		FileOutputStream fout = null;
		Workbook wb = null;
		Sheet sh1 = null;
		Sheet sh2 = null;
		Row rowsh1 = null;
		Row rowsh2 = null;
		Cell cellsh1 = null;
		Cell cellsh2 = null;

		try {
			fin = new FileInputStream("A:\\ExcelAutomation\\Assignment\\Assignment8.xlsx");
			wb = new XSSFWorkbook(fin);
			sh1 = wb.getSheet("Sheet1");
			sh2 = wb.getSheet("Sheet2");
			if(sh2 == null)
			{
				sh2 = wb.createSheet("Sheet2");
			}
			int rc = sh1.getPhysicalNumberOfRows();
			rowsh2 = sh2.getRow(9);
			if(rowsh2 == null)
			{
				rowsh2=sh2.createRow(9);
			}
			for(int i=0 ; i<rc ; i++)
			{
				rowsh1 = sh1.getRow(i);				
				cellsh1 = rowsh1.getCell(0)	;
				cellsh2 = rowsh2.getCell(i);
				if(cellsh2 == null)
				{
					cellsh2 = rowsh2.createCell(i);
				}
				String data = cellsh1.getStringCellValue();
				cellsh2.setCellValue(data);
			}
			rowsh2 = sh2.createRow(10);
			for(int i=0 ; i<rc ; i++)
			{
				rowsh1 = sh1.getRow(i);
				
				cellsh1 = rowsh1.getCell(1)	;
				cellsh2 = rowsh2.getCell(i);
				if(cellsh2 == null)
				{
					cellsh2 = rowsh2.createCell(i);
				}
				String data = cellsh1.getStringCellValue();
				cellsh2.setCellValue(data);
			}
		
			fout = new FileOutputStream("A:\\ExcelAutomation\\Assignment\\Assignment8.xlsx");
			wb.write(fout);
			
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		finally
		{
			try {
				fin.close();
				fout.close();
				wb.close();
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
	}

}
