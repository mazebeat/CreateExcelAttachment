/**
 * 
 */
package cl.intelidata.utils;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author DIEGOPC
 * 
 */
public class FileLib {

	private static String	SRCEXCEL;
	private static String	SRCZIP;

	/**
	 * 
	 * @return
	 */
	public static String getSRCEXCEL() {
		return SRCEXCEL;
	}

	/**
	 * 
	 * @param sRCEXCEL
	 */
	public static void setSRCEXCEL(String sRCEXCEL) {
		SRCEXCEL = sRCEXCEL;
	}

	/**
	 * 
	 * @return
	 */
	public static String getSRCZIP() {
		return SRCZIP;
	}

	/**
	 * 
	 * @param sRCZIP
	 */
	public static void setSRCZIP(String sRCZIP) {
		SRCZIP = sRCZIP;
	}

	/**
	 * Crea el nombre del archivo de salida según lo acordado con Wladimir Cea
	 * => Correo_Administrador + Fecha_proceso
	 * 
	 * @param mailAdmin
	 * @param dateProcess
	 * @return
	 */
	public static String createNameFile(String mailAdmin, String dateProcess) {
		String a = mailAdmin.replaceAll("[@.]", "_");		
		String name = a.concat(dateProcess).toLowerCase();

		return name;
	}

	/**
	 * Crea el archivo xlsx para luego comprimirlo
	 * 
	 * @param fileWrite
	 * @param workbook
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public static boolean createFile(String fileWrite, XSSFWorkbook workbook)
			throws FileNotFoundException, IOException {
		String srcTemp = SRCEXCEL + "/" + fileWrite + ".xlsx"; // "C:/fileWrite.xlsx"
		String srcFileWrite;
		File f = new File(srcTemp);
		if (f.exists()) {
			srcFileWrite = SRCEXCEL + "/" + fileWrite + "_1.xlsx"; // "C:/fileWrite_1.xlsx"
		} else {
			srcFileWrite = SRCEXCEL + "/" + fileWrite + ".xlsx"; // "C:/fileWrite.xlsx"
		}
		FileOutputStream out = new FileOutputStream(srcFileWrite);
		try {

			workbook.write(out);
			out.flush();
			ZipLib.zip(srcFileWrite, SRCEXCEL + "/" + fileWrite + ".zip");
			StringLib.generateInfo("Excel written successfully");
		} catch (Exception ex) {
			StringLib.generateAlert("Exception occur " + ex);
			return false;
		} finally {
			out.close();
		}

		return true;
	}

	/**
	 * Elimina un archivo dentro de un directorio
	 * 
	 * @param dir
	 */
	public static void delteFile(String file) {
		File f = new File(file);
		if (f.delete())
			StringLib.generateInfo("El fichero " + file
					+ " ha sido borrado correctamente");
		else
			StringLib.generateInfo("El fichero " + file
					+ " no se ha podido borrar");
	}

	/**
	 * Elimina todo el contenido de un directorio
	 * 
	 * @param dir
	 */
	public static void cleanFolder(String dir) {
		File directorio = new File(dir);
		File f;
		if (directorio.isDirectory()) {
			if (!directorio.exists()) {
				directorio.mkdirs();
			}

			String[] files = directorio.list();
			if (files.length > 0) {
				for (String archivo : files) {
					f = new File(dir + File.separator + archivo);
					if (archivo.contains(".xlsx")) {
						f.delete();
						f.deleteOnExit();
					}
				}
			}
		}
	}
}
