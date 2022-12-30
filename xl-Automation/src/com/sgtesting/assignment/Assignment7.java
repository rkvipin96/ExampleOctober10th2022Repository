package com.sgtesting.assignment;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Assignment7 {

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
			fin = new FileInputStream("A:\\ExcelAutomation\\Assignment\\Assignment7.xlsx");
			wb = new XSSFWorkbook(fin);
			sh1 = wb.getSheet("Sheet1");
			sh2 = wb.getSheet("Sheet2");
			if(sh2 == null)
			{
				sh2 = wb.createSheet("Sheet2");
						
			}
			int rc = sh1.getPhysicalNumberOfRows();
			for(int i=0 ; i< rc ; i++)
			{
				rowsh1 = sh1.getRow(i);
				rowsh2 = sh2.getRow(i);
				if(rowsh2 == null)
				{
					rowsh2 = sh2.createRow(i);
				}
				int cc = rowsh1.getPhysicalNumberOfCells();
				int k = 8;
				for(int j = 0 ; j<cc ; j++)
				{
					cellsh1 = rowsh1.getCell(j);
					cellsh2 = rowsh2.getCell(k);
					if(cellsh2 == null)
					{
						cellsh2 = rowsh2.createCell(k);
					}
					
					String data = cellsh1.getStringCellValue();
					cellsh2.setCellValue(data);	
					k++;				
				}
			}
			
			fout = new FileOutputStream("A:\\ExcelAutomation\\Assignment\\Assignment7.xlsx");
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
