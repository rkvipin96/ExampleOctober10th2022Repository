package com.sgtesting.assignment;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Assignment9 {

	public static void main(String[] args) {
		FileInputStream fin = null;
		FileOutputStream fout = null;
		Workbook wb = null;
		Workbook wb2 = null;
		Sheet sh1 = null;
		Sheet sh2 = null;
		Row rowsh1 = null;
		Row rowsh2 = null;
		Cell cell1 = null;
		Cell cell2 = null;
		
		try {
			fin = new FileInputStream("A:\\ExcelAutomation\\Assignment\\inputAssign9.xlsx");
			wb = new XSSFWorkbook(fin);
			wb2 = new XSSFWorkbook();
			sh1 = wb.getSheet("Sheet1");
			sh2 = wb2.createSheet("Sheet1");
			int rc = sh1.getPhysicalNumberOfRows();
			for(int i= 0 ; i<rc ; i++)
			{
				rowsh1 = sh1.getRow(i);
				rowsh2 = sh2.createRow(i);
				int cc = rowsh1.getPhysicalNumberOfCells();
				for(int j = 0 ; j<cc ; j++)
				{
					cell1 = rowsh1.getCell(j);
					cell2 = rowsh2.createCell(j);
					String data = cell1.getStringCellValue();
					cell2.setCellValue(data);
				}
			}
			
			fout = new FileOutputStream("A:\\ExcelAutomation\\Assignment\\Assignment9.xlsx");
			wb2.write(fout);
		} catch (Exception e) {
			e.printStackTrace();
		}
		finally
		{
			try {
				fin.close();
				fout.close();
				wb.close();
				wb2.close();
			} catch (Exception e) {
				e.printStackTrace();
			}
		}

	}

}
