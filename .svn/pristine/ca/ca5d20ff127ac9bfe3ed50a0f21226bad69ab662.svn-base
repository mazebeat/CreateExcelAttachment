/**
 * 
 */
package cl.intelidata;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import cl.intelidata.utils.ExcelLib;
import cl.intelidata.utils.FileLib;
import cl.intelidata.utils.StringLib;

/**
 * @author DIEGOPC
 * 
 */
public class CreateAllExcelAttachment {

	/**
	 * @param args
	 */
	public static void main(String[] args) throws IOException {
		// Valida la cantidad de argumentos de entrada
		if (args.length != 2) {
			StringLib
					.generateWarning("The number of parameters is incorrect: [fileToRead] [srcDestiny]");
			System.exit(0);
		}

		// Seteamos las rutas a las carpetas correspondientes
		FileLib.setSRCEXCEL(args[1]); // "C:/CreateAttachment/Excel"

		// Se captura el archivo de entrada
		File fileRead = new File(args[0]); // "C:/base_ejemplo.xlsx"

		// Se procesa el archivo matriz
		StringLib.generateInfo("Init process...");
		process(fileRead);

		// Se limpia la carpeta donde se alojan los archivos excel
		FileLib.cleanFolder(FileLib.getSRCEXCEL());
	}

	/**
	 * Procesa la información obtenida del archivo matriz
	 * 
	 * @param srcFileRead
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	public static void process(File srcFileRead) throws IOException,
			FileNotFoundException {
		try {
			int countRows = 1;
			String mailAdmin, dateProcess = null;
			FileInputStream file = new FileInputStream(srcFileRead);
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheetAt(0);

			List<String> listMails = new ArrayList<String>(
					createListMails(sheet));

			Iterator<String> it = listMails.iterator();

			while (it.hasNext()) {
				mailAdmin = (String) it.next();

				XSSFWorkbook workbook2 = new XSSFWorkbook();
				XSSFSheet sheet2 = workbook2.createSheet();

				for (int i = 0; i <= sheet.getLastRowNum(); i++) {
					Row row = sheet.getRow(i);
					if (i == 0) {
						XSSFRow fila = sheet2.createRow(i);

						StringLib.generateInfo("Creating header for "
								+ mailAdmin + " ...");

						if (row.getCell(0).getCellType() != Cell.CELL_TYPE_BLANK) {
							XSSFCell cell = fila.createCell(0);
							row.getCell(0).setCellType(Cell.CELL_TYPE_STRING);
							cell.setCellValue(ExcelLib.readCellValue(row
									.getCell(0)));
							sheet2.autoSizeColumn(0);

							XSSFCellStyle style = workbook2.createCellStyle();
							style.setFillForegroundColor(IndexedColors.DARK_BLUE
									.getIndex());
							style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

							style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
							style.setBottomBorderColor(IndexedColors.BLACK
									.getIndex());
							style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
							style.setLeftBorderColor(IndexedColors.BLACK
									.getIndex());
							style.setBorderRight(HSSFCellStyle.BORDER_THIN);
							style.setRightBorderColor(IndexedColors.BLACK
									.getIndex());
							style.setBorderTop(HSSFCellStyle.BORDER_THIN);
							style.setTopBorderColor(IndexedColors.BLACK
									.getIndex());
							cell.setCellStyle(style);

							XSSFFont font = workbook2.createFont();
							font.setColor(IndexedColors.WHITE.getIndex());
							style.setFont(font);
						}

						if (row.getCell(4).getCellType() != Cell.CELL_TYPE_BLANK) {
							XSSFCell cell = fila.createCell(1);
							row.getCell(4).setCellType(Cell.CELL_TYPE_STRING);
							cell.setCellValue(ExcelLib.readCellValue(row
									.getCell(4)));
							sheet2.autoSizeColumn(1);

							XSSFCellStyle style = workbook2.createCellStyle();
							style.setFillForegroundColor(IndexedColors.DARK_BLUE
									.getIndex());
							style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

							style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
							style.setBottomBorderColor(IndexedColors.BLACK
									.getIndex());
							style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
							style.setLeftBorderColor(IndexedColors.BLACK
									.getIndex());
							style.setBorderRight(HSSFCellStyle.BORDER_THIN);
							style.setRightBorderColor(IndexedColors.BLACK
									.getIndex());
							style.setBorderTop(HSSFCellStyle.BORDER_THIN);
							style.setTopBorderColor(IndexedColors.BLACK
									.getIndex());
							cell.setCellStyle(style);

							XSSFFont font = workbook2.createFont();
							font.setColor(IndexedColors.WHITE.getIndex());
							style.setFont(font);
						}

						if (row.getCell(5).getCellType() != Cell.CELL_TYPE_BLANK) {
							XSSFCell cell = fila.createCell(2);
							row.getCell(5).setCellType(Cell.CELL_TYPE_STRING);
							cell.setCellValue(ExcelLib.readCellValue(row
									.getCell(5)));
							sheet2.autoSizeColumn(2);

							XSSFCellStyle style = workbook2.createCellStyle();
							style.setFillForegroundColor(IndexedColors.DARK_BLUE
									.getIndex());
							style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

							style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
							style.setBottomBorderColor(IndexedColors.BLACK
									.getIndex());
							style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
							style.setLeftBorderColor(IndexedColors.BLACK
									.getIndex());
							style.setBorderRight(HSSFCellStyle.BORDER_THIN);
							style.setRightBorderColor(IndexedColors.BLACK
									.getIndex());
							style.setBorderTop(HSSFCellStyle.BORDER_THIN);
							style.setTopBorderColor(IndexedColors.BLACK
									.getIndex());
							cell.setCellStyle(style);

							XSSFFont font = workbook2.createFont();
							font.setColor(IndexedColors.WHITE.getIndex());
							style.setFont(font);
						}

						if (row.getCell(6).getCellType() != Cell.CELL_TYPE_BLANK) {
							XSSFCell cell = fila.createCell(3);
							row.getCell(6).setCellType(Cell.CELL_TYPE_STRING);
							cell.setCellValue(ExcelLib.readCellValue(row
									.getCell(6)));
							sheet2.autoSizeColumn(3);

							XSSFCellStyle style = workbook2.createCellStyle();
							style.setFillForegroundColor(IndexedColors.DARK_BLUE
									.getIndex());
							style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

							style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
							style.setBottomBorderColor(IndexedColors.BLACK
									.getIndex());
							style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
							style.setLeftBorderColor(IndexedColors.BLACK
									.getIndex());
							style.setBorderRight(HSSFCellStyle.BORDER_THIN);
							style.setRightBorderColor(IndexedColors.BLACK
									.getIndex());
							style.setBorderTop(HSSFCellStyle.BORDER_THIN);
							style.setTopBorderColor(IndexedColors.BLACK
									.getIndex());
							cell.setCellStyle(style);

							XSSFFont font = workbook2.createFont();
							font.setColor(IndexedColors.WHITE.getIndex());
							style.setFont(font);
						}

						// for (int c = 0; c < 7; c++) {
						// XSSFCell cell = fila.createCell(c);
						// cell.setCellValue(row.getCell(c).getStringCellValue());
						// sheet2.autoSizeColumn(c);
						//
						// XSSFCellStyle style = workbook2.createCellStyle();
						// style.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
						// style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
						//
						// style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
						// style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
						// style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
						// style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
						// style.setBorderRight(HSSFCellStyle.BORDER_THIN);
						// style.setRightBorderColor(IndexedColors.BLACK.getIndex());
						// style.setBorderTop(HSSFCellStyle.BORDER_THIN);
						// style.setTopBorderColor(IndexedColors.BLACK.getIndex());
						// cell.setCellStyle(style);
						//
						// XSSFFont font = workbook2.createFont();
						// font.setColor(IndexedColors.WHITE.getIndex());
						// style.setFont(font);
						// }

						countRows = 1;

						StringLib.generateInfo("Header created successfully");
					} else {
						if (i == 1) {
							StringLib.generateInfo("Creating body " + mailAdmin
									+ " ...");
						}

						if (row.getCell(1).getStringCellValue()
								.equalsIgnoreCase(mailAdmin)) {

							XSSFRow fila = sheet2.createRow(countRows);

							if (row.getCell(0).getCellType() != Cell.CELL_TYPE_BLANK) {
								XSSFCell cell = fila.createCell(0);
								row.getCell(0).setCellType(
										Cell.CELL_TYPE_STRING);
								cell.setCellValue(ExcelLib.readCellValue(row
										.getCell(0)));
								sheet2.autoSizeColumn(0);

								XSSFCellStyle style = workbook2
										.createCellStyle();
								style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
								style.setBottomBorderColor(IndexedColors.BLACK
										.getIndex());
								style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
								style.setLeftBorderColor(IndexedColors.BLACK
										.getIndex());
								style.setBorderRight(HSSFCellStyle.BORDER_THIN);
								style.setRightBorderColor(IndexedColors.BLACK
										.getIndex());
								style.setBorderTop(HSSFCellStyle.BORDER_THIN);
								style.setTopBorderColor(IndexedColors.BLACK
										.getIndex());
								cell.setCellStyle(style);
							}

							if (row.getCell(4).getCellType() != Cell.CELL_TYPE_BLANK) {
								XSSFCell cell = fila.createCell(1);
								row.getCell(4).setCellType(
										Cell.CELL_TYPE_STRING);
								cell.setCellValue(ExcelLib.readCellValue(row
										.getCell(4)));
								sheet2.autoSizeColumn(1);

								XSSFCellStyle style = workbook2
										.createCellStyle();
								style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
								style.setBottomBorderColor(IndexedColors.BLACK
										.getIndex());
								style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
								style.setLeftBorderColor(IndexedColors.BLACK
										.getIndex());
								style.setBorderRight(HSSFCellStyle.BORDER_THIN);
								style.setRightBorderColor(IndexedColors.BLACK
										.getIndex());
								style.setBorderTop(HSSFCellStyle.BORDER_THIN);
								style.setTopBorderColor(IndexedColors.BLACK
										.getIndex());
								cell.setCellStyle(style);
							}

							if (row.getCell(5).getCellType() != Cell.CELL_TYPE_BLANK) {
								XSSFCell cell = fila.createCell(2);
								row.getCell(5).setCellType(
										Cell.CELL_TYPE_STRING);
								cell.setCellValue(ExcelLib.readCellValue(row
										.getCell(5)));
								sheet2.autoSizeColumn(2);

								XSSFCellStyle style = workbook2
										.createCellStyle();
								style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
								style.setBottomBorderColor(IndexedColors.BLACK
										.getIndex());
								style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
								style.setLeftBorderColor(IndexedColors.BLACK
										.getIndex());
								style.setBorderRight(HSSFCellStyle.BORDER_THIN);
								style.setRightBorderColor(IndexedColors.BLACK
										.getIndex());
								style.setBorderTop(HSSFCellStyle.BORDER_THIN);
								style.setTopBorderColor(IndexedColors.BLACK
										.getIndex());
								cell.setCellStyle(style);
							}

							if (row.getCell(6).getCellType() != Cell.CELL_TYPE_BLANK) {
								XSSFCell cell = fila.createCell(3);
								row.getCell(6).setCellType(
										Cell.CELL_TYPE_STRING);
								cell.setCellValue(ExcelLib.readCellValue(row
										.getCell(6)));
								sheet2.autoSizeColumn(3);

								XSSFCellStyle style = workbook2
										.createCellStyle();
								style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
								style.setBottomBorderColor(IndexedColors.BLACK
										.getIndex());
								style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
								style.setLeftBorderColor(IndexedColors.BLACK
										.getIndex());
								style.setBorderRight(HSSFCellStyle.BORDER_THIN);
								style.setRightBorderColor(IndexedColors.BLACK
										.getIndex());
								style.setBorderTop(HSSFCellStyle.BORDER_THIN);
								style.setTopBorderColor(IndexedColors.BLACK
										.getIndex());
								cell.setCellStyle(style);
							}

							dateProcess = ExcelLib
									.readCellValue(row.getCell(0));

							// for (int c = 0; c < 7; c++) {
							// XSSFCell cell = fila.createCell(c);
							//
							// if (row.getCell(c).getCellType() !=
							// Cell.CELL_TYPE_BLANK) {
							// row.getCell(c).setCellType(Cell.CELL_TYPE_STRING);
							// cell.setCellValue(ExcelLib.readCellValue(row.getCell(c)));
							// sheet2.autoSizeColumn(c);
							// }
							//
							// if (c == 0) {
							// dateProcess =
							// ExcelLib.readCellValue(row.getCell(c));
							// }
							// XSSFCellStyle style =
							// workbook2.createCellStyle();
							// style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
							// style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
							// style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
							// style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
							// style.setBorderRight(HSSFCellStyle.BORDER_THIN);
							// style.setRightBorderColor(IndexedColors.BLACK.getIndex());
							// style.setBorderTop(HSSFCellStyle.BORDER_THIN);
							// style.setTopBorderColor(IndexedColors.BLACK.getIndex());
							// cell.setCellStyle(style);
							// }
							countRows++;
						}
					}
				}

				StringLib.generateInfo("Created " + (countRows - 1) + "  rows");
				StringLib.generateInfo("Body created successfully");

				if (dateProcess != null) {
					String nameFile = FileLib.createNameFile(mailAdmin,
							dateProcess);
					FileLib.createFile(nameFile, workbook2);
				}
			}
		} catch (Exception ex) {
			StringLib.generateAlert("Exception occur " + ex);
		} finally {
			StringLib.generateInfo("Finish process");
		}
	}

	/**
	 * Crea una lista con valores unicos con el campo Correo_Administrador del
	 * archivo matriz
	 * 
	 * @param sheet
	 * @return
	 */
	public static List<String> createListMails(Sheet sheet) {
		List<String> listMails = new ArrayList<String>();

		for (int i = 1; i <= sheet.getLastRowNum(); i++) {
			Row row = sheet.getRow(i);
			if (row.getCell(1).getCellType() != Cell.CELL_TYPE_BLANK) {
				listMails.add(row.getCell(1).getStringCellValue());
			}
		}

		Set<String> setMails = StringLib.sortList(listMails);

		return new ArrayList<String>(setMails);
	}

}
