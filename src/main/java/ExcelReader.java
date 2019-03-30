import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;

public class ExcelReader {

	public static final String FILE_PATH = "./katt.xls";

	public static void main(String[] args) throws IOException, InvalidFormatException {

		// Initialize workbook from define path
		Workbook workbook = WorkbookFactory.create(new File(FILE_PATH));

		// create file to write in results
		FileWriter fileWriterFirmy = new FileWriter("firmy.txt");
		PrintWriter printWriterFirmy = new PrintWriter(fileWriterFirmy);

		long startTime = System.currentTimeMillis();

		// foreach and thread
		new Thread(() -> {
			int inx = 0;

			for (Sheet sheet : workbook) {
				printWriterFirmy.println(sheet.getSheetName() + ";" + inx++);  
			}

			long endTime1 = System.currentTimeMillis();
			System.out.println("Time 1thread: " + (endTime1 - startTime));
			printWriterFirmy.close();
		}).start();

		
		FileWriter fileWriter = new FileWriter("katalogik.txt");
		PrintWriter printWriter = new PrintWriter(fileWriter);

		new Thread(() -> {
			for (int y = 0; y < workbook.getNumberOfSheets(); y++) {
				Sheet sheet = workbook.getSheetAt(y);
				for (int j = 1; j < sheet.getLastRowNum(); j++) {
					Row row = sheet.getRow(j);
					// System.out.println(row.getCell(2));
					if (!(row.getCell(2).getCellType() == Cell.CELL_TYPE_BLANK)) {
						for (int i = 0; i < 6; i++) {
							Cell cell = row.getCell(i);
							if (cell != null) {
								printWriter.print(printCellValue(cell));
								if (getMergedRegionForCell(cell) != null) {
									printWriter.print(";" + row.getCell(i));
									i = 4;
								}
							} else {
								printWriter.print(cell);
							}
							if (i != 5)
								printWriter.print(";");
						}
						;
						printWriter.print(";" + y);
						printWriter.println("");
					}
				}
			}
			long endTime2 = System.currentTimeMillis();
			System.out.println("Time 2thread: " + (endTime2 - startTime));
			printWriter.close();
		}).start();

		workbook.close();
		// Closing the workbook
		

		long endTimeF = System.currentTimeMillis();
		System.out.println("ALL: " + (endTimeF - startTime));

	}

	public static CellRangeAddress getMergedRegionForCell(Cell c) {
		Sheet s = c.getRow().getSheet();
		for (CellRangeAddress mergedRegion : s.getMergedRegions()) {
			if (mergedRegion.isInRange(c.getRowIndex(), c.getColumnIndex())) {
				return mergedRegion;
			}
		}
		return null;
	}

	private static String printCellValue(Cell cell) {
		switch (cell.getCellTypeEnum()) {
		case BOOLEAN:
			System.out.print(cell.getBooleanCellValue());
			break;
		case STRING:
			return (cell.getRichStringCellValue().getString());
		case NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				System.out.print(cell.getDateCellValue());
			} else {
				return Double.toString(cell.getNumericCellValue());
			}
			break;
		case FORMULA:
			System.out.print(cell.getCellFormula());
			break;
		case BLANK:
			return ("");
		default:
			return ("");
		}
		return ("");
	}
}
