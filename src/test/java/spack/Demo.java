package spack;

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

public class Demo {

	public static void main(String[] args) throws IOException {

		ArrayList<String> alist = getDataFromExcelFile("Register");

		for (String a : alist) {

			System.out.println(a);

		}
	}

	public static ArrayList<String> getDataFromExcelFile(String testName) throws IOException {

		ArrayList<String> alist = new ArrayList<String>();

		FileInputStream fis = new FileInputStream("C:\\Users\\User\\Desktop\\ExcelTestData.xlsx");
		@SuppressWarnings("resource")
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		int sheetCount = workbook.getNumberOfSheets();

		for (int i = 0; i < sheetCount; i++) {

			if (workbook.getSheetName(i).equalsIgnoreCase("SheetA")) {

				XSSFSheet sheet = workbook.getSheetAt(i);

				Iterator<Row> rows = sheet.iterator();

				Row firstrow = rows.next();

				Iterator<Cell> firstrowCells = firstrow.iterator();

				int c = 0;
				int TestColumnPosition = 0;

				while (firstrowCells.hasNext()) {

					Cell firstrowCell = firstrowCells.next();

					if (firstrowCell.getStringCellValue().equalsIgnoreCase("Tests")) {

						TestColumnPosition = c;

					}
					c++;
				}

				while (rows.hasNext()) {

					Row row = rows.next();

					Cell cell = row.getCell(TestColumnPosition);

					if (cell.getStringCellValue().equalsIgnoreCase(testName)) {

						Iterator<Cell> cells = row.iterator();

						cells.next();

						while (cells.hasNext()) {

							Cell currentCell = cells.next();

							if (currentCell.getCellType() == CellType.STRING) {

								System.out.println();
								alist.add(currentCell.getStringCellValue());

							} else if (currentCell.getCellType() == CellType.NUMERIC)

								alist.add(NumberToTextConverter.toText(currentCell.getNumericCellValue()));

						}

					}
				}

			}

		}

		return alist;

	}

}
