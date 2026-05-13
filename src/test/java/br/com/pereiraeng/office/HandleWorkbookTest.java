package br.com.pereiraeng.office;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.junit.jupiter.api.Assertions.assertFalse;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URISyntaxException;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

public class HandleWorkbookTest {

	@Test
	public void testWriteAndReadWorkbookTest() throws IOException {
		File file = new File("newWorkbook.xlsx");
		assertFalse(file.exists());

		XSSFWorkbook workbook = Office.getWB(file);

		XSSFSheet sheet = workbook.createSheet();

		XSSFRow row = sheet.createRow(0);
		row.createCell(1).setCellValue("col1");
		row.createCell(2).setCellValue("col2");

		row = sheet.createRow(1);
		row.createCell(0).setCellValue("row1");
		row.createCell(1).setCellValue(1.);
		row.createCell(2).setCellValue(2.);

		row = sheet.createRow(2);
		row.createCell(0).setCellValue("row2");
		row.createCell(1).setCellValue(3.);
		row.createCell(2).setCellValue(4.);

		FileOutputStream outputStream = new FileOutputStream(file);
		workbook.write(outputStream);
		outputStream.close();
		workbook.close();

		assertTrue(file.exists());

		workbook = Office.getWB(file);
		sheet = workbook.getSheetAt(0);

		row = sheet.getRow(0);
		assertEquals("col1", row.getCell(1).getStringCellValue());
		assertEquals("col2", row.getCell(2).getStringCellValue());

		row = sheet.getRow(1);
		assertEquals("row1", row.getCell(0).getStringCellValue());
		assertEquals(1, row.getCell(1).getNumericCellValue());
		assertEquals(2, row.getCell(2).getNumericCellValue());

		row = sheet.getRow(2);
		assertEquals("row2", row.getCell(0).getStringCellValue());
		assertEquals(3, row.getCell(1).getNumericCellValue());
		assertEquals(4, row.getCell(2).getNumericCellValue());

		workbook.close();

		assertTrue(file.delete());
	}

	@Test
	public void editWorkbookTest() throws IOException {
		String filename = "simpleWorkbook.xlsx";
		copyToRoot(filename);

		File file = new File(filename);

		XSSFWorkbook workbook = Office.getWB(file, true);

		XSSFSheet sheet = workbook.getSheetAt(0);

		XSSFRow row = sheet.getRow(0);
		row.getCell(1).setCellValue("col3");
		row.getCell(2).setCellValue("col4");

		row = sheet.getRow(1);
		row.getCell(0).setCellValue("row3");
		row.getCell(1).setCellValue(5.);
		row.getCell(2).setCellValue(6.);

		row = sheet.getRow(2);
		row.getCell(0).setCellValue("row4");
		row.getCell(1).setCellValue(7.);
		row.getCell(2).setCellValue(8.);

		FileOutputStream outputStream = new FileOutputStream(file);
		workbook.write(outputStream);
		outputStream.close();
		workbook.close();

		assertTrue(file.exists());

		workbook = Office.getWB(file);
		sheet = workbook.getSheetAt(0);

		row = sheet.getRow(0);
		assertEquals("col3", row.getCell(1).getStringCellValue());
		assertEquals("col4", row.getCell(2).getStringCellValue());

		row = sheet.getRow(1);
		assertEquals("row3", row.getCell(0).getStringCellValue());
		assertEquals(5., row.getCell(1).getNumericCellValue());
		assertEquals(6., row.getCell(2).getNumericCellValue());

		row = sheet.getRow(2);
		assertEquals("row4", row.getCell(0).getStringCellValue());
		assertEquals(7., row.getCell(1).getNumericCellValue());
		assertEquals(8., row.getCell(2).getNumericCellValue());

		workbook.close();

		assertTrue(file.delete());
	}

	private static void copyToRoot(String resourceName) {
		URL resourceUrl = ClassLoader.getSystemResource(resourceName);
		if (resourceUrl == null) {
			throw new IllegalArgumentException("Resource not found: " + resourceName);
		}
		try {
			Files.copy(new File(resourceUrl.toURI()).toPath(), new File(resourceName).toPath(),
					StandardCopyOption.REPLACE_EXISTING);
		} catch (URISyntaxException | IOException e) {
			e.printStackTrace();
		}

	}
}
