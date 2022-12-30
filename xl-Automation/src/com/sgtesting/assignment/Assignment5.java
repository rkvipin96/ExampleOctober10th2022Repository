package com.sgtesting.assignment;

import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Assignment5 {

	public static void main(String[] args) {
		FileOutputStream fout = null;
		Workbook wb = null;
		Sheet sh = null;
		Row row = null;
		Cell cell = null;
		
		try {
			wb = new XSSFWorkbook();
			sh=wb.createSheet("FlowerData");
			for(int i = 0; i<20 ; i++)
			{
				row = sh.createRow(i);
				for(int j = 0 ; j < 2 ; j++)
				{
					cell = row.createCell(j);
					if(j == 0)
					{
						cell.setCellValue("Flower" + i);
					}
					else
					{
						cell.setCellValue("Colour" + i);
					}
				}
			}
			
			fout = new FileOutputStream("A:\\ExcelAutomation\\Assignment\\Assignment5.xlsx");
			wb.write(fout);
				
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		finally
		{
			try
			{
				fout.close();
				wb.close();
			}catch(Exception e)
			{
				e.printStackTrace();
			}
		}

	}

}
