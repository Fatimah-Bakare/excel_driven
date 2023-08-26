package rahulshettyacademy.ExcelDriven;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDriven {

	public ArrayList<String> getData(String testcaseName) throws IOException {

		ArrayList<String> a = new ArrayList<String>();

		FileInputStream fis = new FileInputStream("C:\\Users\\fatim\\Desktop\\ExcelDriven.xlsx");
		XSSFWorkbook workBook = new XSSFWorkbook(fis);

		int sheets = workBook.getNumberOfSheets();
		System.out.println(sheets);
		for (int i = 0; i < sheets; i++) {
			if (workBook.getSheetName(i).equalsIgnoreCase("testdata")) {
				XSSFSheet sheet = workBook.getSheetAt(i);

				// Identify the Testcases column by scanning the entire 1st row
				Iterator<Row> rows = sheet.iterator();// sheet is collection of rows
				Row firstRow = rows.next();
				Iterator<Cell> cell = firstRow.cellIterator();// row is collection of cells
				int k = 0;
				int column = 0;
				while (cell.hasNext()) {
					Cell value = cell.next();
					if (value.getStringCellValue().equalsIgnoreCase("Testcases")) {

						column = k;

					}
					k++;
				}
				System.out.println(column);

				// once column is identified, scan the entire testcase column to identify
				// Purchase testcase row
				while (rows.hasNext()) {
					Row row = rows.next();
					if (row.getCell(column).getStringCellValue().equalsIgnoreCase(testcaseName)) {
						// after you grab Purchase testcase row = pull all the data of that row and feed
						// it into test
						Iterator<Cell> cellValue = row.cellIterator();
						while (cellValue.hasNext()) {
							// data is getting stored in ArrayList a:
							Cell c = cellValue.next();
							if (c.getCellType()==CellType.STRING) 
							{
								a.add(c.getStringCellValue());
							}
							else 
							{
								a.add(NumberToTextConverter.toText(c.getNumericCellValue()));
							}
						}
					}
				}
			}
		}
		return a;
	}

	public static void main(String[] args) throws IOException {

	}

}
