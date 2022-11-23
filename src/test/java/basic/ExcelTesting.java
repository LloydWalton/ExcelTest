package basic;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelTesting {

	public static void main(String[] args) throws IOException {
		
		
		
		FileInputStream fis = new FileInputStream("D:\\Test\\Book1.xlsx"); // Access the file
		
		XSSFWorkbook workbook=new XSSFWorkbook(fis); //complete access to file
		
		System.out.println(workbook.getNumberOfSheets()); //get number of sheets
		
		
		int sheetCount=workbook.getNumberOfSheets();
		for(int i=0;i<sheetCount;i++)
		{
			//System.out.println(workbook.getSheetName(i)); //sheet name
			
			if(workbook.getSheetName(i).equalsIgnoreCase("Sheet1"))
			{
				XSSFSheet sheet=workbook.getSheetAt(i); //Located sheet
				
				
				Iterator<Row> rows=sheet.iterator();
				Row firstRow=rows.next();
				
				
				
				
				Iterator<Cell> cells=firstRow.cellIterator();
				while(cells.hasNext())
				{
					Cell value=cells.next();
					if(value.getStringCellValue().equalsIgnoreCase("Name"))
					{
						
					}
				}
				
				
				
				
				


			}
		}
		

	}

}
