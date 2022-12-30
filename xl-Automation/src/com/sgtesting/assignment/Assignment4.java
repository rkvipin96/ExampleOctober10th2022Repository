package com.sgtesting.assignment;

import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class Assignment4 {

	public static void main(String[] args) {
		FileOutputStream fout = null;
		Workbook wb = null;
		Sheet sh = null;
		Row row = null;
		Cell cell = null;
		
		try {
			wb = new XSSFWorkbook();
			sh = wb.createSheet();
			for(int i = 0 ; i<20 ; i++)
			{
				row = sh.createRow(i);
				
				for(int j = 0 ; j<5 ; j++)
				{
					cell=row.createCell(j);
					if(j==4)
					{
						cell.setCellValue("Vegetable" + i);
					}
				}
			
			}
			
			fout= new FileOutputStream("A:\\ExcelAutomation\\Assignment\\Assignment4.xlsx");
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
