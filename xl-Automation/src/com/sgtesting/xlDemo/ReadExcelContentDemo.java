package com.sgtesting.xlDemo;

import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelContentDemo {

	public static void main(String[] args) {
		FileInputStream fin = null;
		Workbook wb = null;
		Sheet sh = null;
		Row row = null;
		Cell cell = null;
		
		try
		{
			fin = new FileInputStream("A:\\ExcelAutomation\\Customer.xlsx");
			wb = new XSSFWorkbook(fin);
			sh = wb.getSheet("Sheet1");
			int rc = sh.getPhysicalNumberOfRows();
			for(int i = 0 ; i<rc ; i++)
			{
				row = sh.getRow(i);
				int cc = row.getPhysicalNumberOfCells();
				for(int c=0 ; c<cc; c++)
				{
					cell = row.getCell(c);
					String data = cell.getStringCellValue();
					System.out.print(data + "    ");
					
				}
				System.out.printf("\n");
			}
		}catch(Exception e)
		{
			e.printStackTrace();
		}
		finally
		{
			
			try
			{
				fin.close();
				wb.close();
			}catch(Exception e)
			{
				e.printStackTrace();
			}
			
		}
	}

}
