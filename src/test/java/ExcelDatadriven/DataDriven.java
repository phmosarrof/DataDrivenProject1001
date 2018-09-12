package ExcelDatadriven;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.FileNotFoundException;

import java.util.Iterator;

public class DataDriven {

	public static void main(String[] args) throws IOException   {

		FileInputStream fis = new FileInputStream("D://Downloads//DataDrivenTest.xlsx");

		XSSFWorkbook workbook = new XSSFWorkbook(fis);

		int sheets = workbook.getNumberOfSheets();

		for (int i=0; i < sheets; i++) 
		{
			if (workbook.getSheetName(i).equalsIgnoreCase("Product"))
			{
				XSSFSheet sheet = workbook.getSheetAt(i);

				Iterator <Row> rows = sheet.iterator();
				Row firstrow = rows.next();

				Iterator <Cell> cel = firstrow.cellIterator();

				int k = 0;
				int column;

				while (cel.hasNext())
				{

					Cell value = cel.next();

					{
						if(value.getStringCellValue().equalsIgnoreCase("name"));
						{

							column = k ;
						}
						k++;

					}
					System.out.println(column);
				}
			}
		}
	}
}