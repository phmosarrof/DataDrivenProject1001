package ExcelDatadriven;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PracticeDD {

	public static void main(String[] args) throws IOException {

		FileInputStream fis = new FileInputStream("D://Downloads//DataDrivenTest.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);

		int sheets = workbook.getNumberOfSheets();
		int k = 0;
		int column ;
		for(int i=0;i<sheets;i++) 
			{
			if(workbook.getSheetName(i).equalsIgnoreCase("product")) 
			{
				XSSFSheet sheet = workbook.getSheetAt(i);

				Iterator <Row> rows = sheet.iterator();
				Row firstrow  = rows.next();
				Iterator <Cell> cel = firstrow.cellIterator();
				
				while (cel.hasNext()) 
				{
					Cell value = cel.next();
					{
						if (value.getStringCellValue().equalsIgnoreCase("name")) 
						{
							column=k;
						}

						k++;	

					}


				}

			}


		}

	}
}
