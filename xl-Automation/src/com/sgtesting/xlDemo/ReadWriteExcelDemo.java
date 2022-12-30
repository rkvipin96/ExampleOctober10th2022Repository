package com.sgtesting.xlDemo;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadWriteExcelDemo {

	public static void main(String[] args) {
		FileInputStream fin=null;
		FileOutputStream fout = null;
		Workbook wb = null;
		Sheet sh1= null;
		Sheet sh2 = null;
		Row rowsh1 = null;
		Row rowsh2 = null;
		Cell cellsh1 = null;
		Cell cellsh2 = null;
		
		try {
			fin = new FileInputStream("A:\\ExcelAutomation\\Customer.xlsx");
			wb = new XSSFWorkbook(fin);
			sh1 = wb.getSheet("Sheet1");
			sh2 = wb.getSheet("Sheet2");
			if(sh2 == null)
			{
				sh2 = wb.createSheet("Sheet2");
			}
			else
			{
				sh2 = wb.createSheet("Sheet3");
			}
			int rc = sh1.getPhysicalNumberOfRows();
			for(int i = 0 ; i<rc ; i++)
			{
				rowsh1 = sh1.getRow(i);
				rowsh2 = sh2.getRow(i);
				if(rowsh2 == null)
				{
					rowsh2 = sh2.createRow(i);
				}
				int cc = rowsh1.getPhysicalNumberOfCells();
				for(int j=0; j<cc ; j++)
				{
					cellsh1 = rowsh1.getCell(j);
					cellsh2 = rowsh2.getCell(j);
					if(cellsh2 == null)
					{
						cellsh2 = rowsh2.createCell(j);
					}
					String data = cellsh1.getStringCellValue();
					cellsh2.setCellValue(data);
				}
								
			}
			
			fout = new FileOutputStream("A:\\ExcelAutomation\\Customer.xlsx");
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
