package ExcelDatadriven;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class DataDriven {

	public static void main(String[] args) throws IOException   {

		FileInputStream fis = new FileInputStream("D://Downloads//DataDrivenTest.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		int sheets = workbook.getNumberOfSheets();
		for (int i=0; i < sheets; i++) {
			if (workbook.getSheetName(i).equalsIgnoreCase("Product"))
			{
				XSSFSheet sheet = workbook.getSheetAt(i);
			}
			
		}
	}

}
