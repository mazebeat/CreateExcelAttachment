/**
 * 
 */
package cl.intelidata;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author DIEGOPC
 * 
 */
public class CreateExcelAttachment {

	private static final Logger LOGGER = Logger
			.getLogger(CreateExcelAttachment.class.getName());

	/**
	 * 
	 * @param args
	 * @throws IOException
	 */
	public static void main(String[] args) throws IOException {
		if (args.length != 2) {
			LOGGER.warning("The number of parameters is incorrect: [fileRead] [filter]");
			System.out
					.println("The number of parameters is incorrect: [fileRead] [filter]");
			System.exit(0);
		}
		LOGGER.info("Init process...");
		System.out.println("Init process...");
		File fileRead = new File(args[0]);
		process(fileRead, args[1]);
		// File fileRead = new File("C:/base_ejemplo.xlsx");
		// process(fileRead, "MMRivera@entel.cl");
	}
	
	/**
	 * 
	 * @param fileRead
	 * @param filter
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	public static void process(File fileRead, String filter)
			throws IOException, FileNotFoundException {
		try {
			int countRows = 1;
			FileInputStream file = new FileInputStream(fileRead);
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFWorkbook workbook2 = new XSSFWorkbook();
			XSSFSheet sheet = workbook.getSheetAt(0);
			XSSFSheet sheet2 = workbook2.createSheet();

			for (int i = 0; i <= sheet.getLastRowNum(); i++) {
				Row row = sheet.getRow(i);
				if (i == 0) {
					XSSFRow fila = sheet2.createRow(i);
					LOGGER.info("Creating header...");
					System.out.println("Creating header...");
					for (int c = 0; c < 7; c++) {
						XSSFCell cell = fila.createCell(c);
						cell.setCellValue(row.getCell(c).getStringCellValue());

						sheet2.autoSizeColumn(c);
					}
					LOGGER.info("Header created successfully");
					System.out.println("Header created successfully");
				} else {
					if (i == 1) {
						LOGGER.info("Creating body...");
						System.out.println("Creating body...");
					}
					if (row.getCell(1).getStringCellValue()
							.equalsIgnoreCase(filter)) {
						XSSFRow fila = sheet2.createRow(countRows);
						for (int c = 0; c < 7; c++) {
							XSSFCell cell = fila.createCell(c);
							if (row.getCell(c).getCellType() != Cell.CELL_TYPE_BLANK) {
								row.getCell(c).setCellType(
										Cell.CELL_TYPE_STRING);
								cell.setCellValue(readCellValue(row.getCell(c)));

								sheet2.autoSizeColumn(c);
							}
						}
						countRows++;
					}
				}
			}
			LOGGER.info("Created " + (int) (countRows - 1) + "  rows");
			System.out.println("Created " + (int) (countRows - 1) + " rows");
			LOGGER.info("Body created successfully");
			System.out.println("Body created successfully");
			createFile(filter, workbook2);
		} catch (Exception ex) {
			LOGGER.severe("Exception occur " + ex);
			System.out.println("Exception occur: " + ex);
		} finally {
			LOGGER.info("Finish process");
			System.out.println("Finish process");
		}
	}

	/**
	 * 
	 * @param cell
	 * @return
	 */
	public static String readCellValue(Cell cell) {
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_BLANK:
			return "";
		case Cell.CELL_TYPE_BOOLEAN:
			return String.valueOf(cell.getBooleanCellValue());
		case Cell.CELL_TYPE_ERROR:
			return String.valueOf(cell.getErrorCellValue());
		case Cell.CELL_TYPE_NUMERIC:
		    if (DateUtil.isCellDateFormatted(cell)) {
		        return cell.getDateCellValue().toString();
		    } else {
		        return Double.toString(cell.getNumericCellValue());
		    }		    
		case Cell.CELL_TYPE_STRING:
			return cell.getStringCellValue();
		case Cell.CELL_TYPE_FORMULA:
		    return cell.getCellFormula();	
		default:
			return "Unknown type";
		}
	}

	/**
	 * 
	 * @param fileWrite
	 * @param workbook
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public static void createFile(String fileWrite, XSSFWorkbook workbook)
			throws FileNotFoundException, IOException {
		FileOutputStream out = new FileOutputStream("C:/" + fileWrite + ".xlsx");
		try {
			workbook.write(out);
			out.flush();
			LOGGER.info("Excel written successfully");
			System.out.println("Excel written successfully");
		} finally {
			out.close();
		}

	}

	/**
	 * 
	 * @param worksheet
	 * @param sourceRowNum
	 * @param destinationRowNum
	 * 
	 *            Usage: HSSFWorkbook workbook = new HSSFWorkbook(new
	 *            FileInputStream("c:/input.xls")); HSSFSheet sheet =
	 *            workbook.getSheet("Sheet1"); copyRow(workbook, sheet, 0, 1);
	 *            FileOutputStream out = new FileOutputStream("c:/output.xls");
	 *            workbook.write(out); out.close();
	 */
	private static void copyRow(Sheet worksheet, int sourceRowNum,
			int destinationRowNum) {
		// Get the source / new row
		Row newRow = worksheet.getRow(destinationRowNum);
		Row sourceRow = worksheet.getRow(sourceRowNum);

		// If the row exist in destination, push down all rows by 1 else create
		// a new row
		if (newRow != null) {
			worksheet
					.shiftRows(destinationRowNum, worksheet.getLastRowNum(), 1);
		} else {
			newRow = worksheet.createRow(destinationRowNum);
		}

		// Loop through source columns to add to new row
		for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
			// Grab a copy of the old/new cell
			Cell oldCell = sourceRow.getCell(i);
			Cell newCell = newRow.createCell(i);

			// If the old cell is null jump to next cell
			if (oldCell == null) {
				newCell = null;
				continue;
			}

			// Use old cell style
			newCell.setCellStyle(oldCell.getCellStyle());

			// If there is a cell comment, copy
			if (newCell.getCellComment() != null) {
				newCell.setCellComment(oldCell.getCellComment());
			}

			// If there is a cell hyperlink, copy
			if (oldCell.getHyperlink() != null) {
				newCell.setHyperlink(oldCell.getHyperlink());
			}

			// Set the cell data type
			newCell.setCellType(oldCell.getCellType());

			// Set the cell data value
			switch (oldCell.getCellType()) {
			case Cell.CELL_TYPE_BLANK:
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				newCell.setCellValue(oldCell.getBooleanCellValue());
				break;
			case Cell.CELL_TYPE_ERROR:
				newCell.setCellErrorValue(oldCell.getErrorCellValue());
				break;
			case Cell.CELL_TYPE_FORMULA:
				newCell.setCellFormula(oldCell.getCellFormula());
				break;
			case Cell.CELL_TYPE_NUMERIC:
				newCell.setCellValue(oldCell.getNumericCellValue());
				break;
			case Cell.CELL_TYPE_STRING:
				newCell.setCellValue(oldCell.getRichStringCellValue());
				break;
			}
		}
	}
}
