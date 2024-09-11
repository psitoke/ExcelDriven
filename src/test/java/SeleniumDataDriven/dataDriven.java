package SeleniumDataDriven;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dataDriven {

	public ArrayList<String> getData(String testcaseName) throws IOException {

		ArrayList<String> a = new ArrayList<String>();

		FileInputStream fis = new FileInputStream(
				"C:\\Users\\pranavsitoke\\Downloads\\Pranav_PersonalDocs\\Selenium_Automation\\DataDriven Excel\\TestDataFile.xlsx");

		XSSFWorkbook wb = new XSSFWorkbook(fis);
		int sheets = wb.getNumberOfSheets();

		for (int i = 0; i < sheets; i++) {

			if (wb.getSheetName(i).equalsIgnoreCase("testdata")) {
				XSSFSheet sheet = wb.getSheetAt(i);

				java.util.Iterator<Row> rows = sheet.iterator();
				Row firstrow = rows.next();

				java.util.Iterator<Cell> ce = firstrow.cellIterator();

				int k = 0;
				int coloumn = 0;
				while (ce.hasNext()) {
					Cell value = ce.next();

					if (value.getStringCellValue().equalsIgnoreCase("TestCase")) {
						coloumn = k;
					}
					k++;
				}
				System.out.println(coloumn);

				while (rows.hasNext()) {
					Row r = rows.next();

					if (r.getCell(coloumn).getStringCellValue().equalsIgnoreCase(testcaseName)) {
						java.util.Iterator<Cell> cv = r.cellIterator();

						while (cv.hasNext()) {
							a.add(cv.next().getStringCellValue());
						}
					}
				
				}

			}

		}return a;
		}

public static void main(String[] args) throws IOException {
			// TODO Auto-generated method stub
		
		
	}

}
