package br.com.pereiraeng.office;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.nio.file.Files;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

class Html2XlsxTest {

	@Test
	void testHtml2Xlsx() {
		String filename = "src/test/resources/SDRO_DIARIO_2023_06_29_HTML_09_ProducaoTermicaUsina.html";

		StringBuilder html = new StringBuilder();
		try {
			String str;
			BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream(filename)));
			while ((str = br.readLine()) != null)
				html.append(str);
			br.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
		File newFile = new File("html.xlsx");
		Office.exportExcel(newFile, html.toString(), 3);

		XSSFWorkbook workbook = Office.getWB(newFile);
		XSSFSheet sheet = workbook.getSheetAt(0);

		XSSFRow row = sheet.getRow(0);
		assertEquals("Subsistema", row.getCell(0).getStringCellValue());
		assertEquals("MWmed no Dia", row.getCell(1).getStringCellValue());
		assertEquals("MWmed no Mêsaté o Dia", row.getCell(2).getStringCellValue());
		assertEquals("MWmed no Anoaté o Dia", row.getCell(3).getStringCellValue());

		row = sheet.getRow(1);
		assertEquals("Norte", row.getCell(0).getStringCellValue());
		assertEquals("1.812,59", row.getCell(1).getStringCellValue());
		assertEquals("2.039,86", row.getCell(2).getStringCellValue());
		assertEquals("1.257,24", row.getCell(3).getStringCellValue());

		row = sheet.getRow(2);
		assertEquals("Nordeste", row.getCell(0).getStringCellValue());
		assertEquals("411,16", row.getCell(1).getStringCellValue());
		assertEquals("465,11", row.getCell(2).getStringCellValue());
		assertEquals("519,91", row.getCell(3).getStringCellValue());

		row = sheet.getRow(3);
		assertEquals("Sul", row.getCell(0).getStringCellValue());
		assertEquals("1.057,15", row.getCell(1).getStringCellValue());
		assertEquals("1.232,55", row.getCell(2).getStringCellValue());
		assertEquals("871,90", row.getCell(3).getStringCellValue());

		row = sheet.getRow(4);
		assertEquals("Sudeste/Centro-Oeste", row.getCell(0).getStringCellValue());
		assertEquals("6.084,29", row.getCell(1).getStringCellValue());
		assertEquals("6.393,70", row.getCell(2).getStringCellValue());
		assertEquals("4.667,57", row.getCell(3).getStringCellValue());

		row = sheet.getRow(5);
		assertEquals("Sistema Interligado Nacional", row.getCell(0).getStringCellValue());
		assertEquals("9.365,20", row.getCell(1).getStringCellValue());
		assertEquals("10.131,22", row.getCell(2).getStringCellValue());
		assertEquals("7.316,63", row.getCell(3).getStringCellValue());

		try {
			workbook.close();
			Files.delete(newFile.toPath());
		} catch (IOException exception) {
			exception.printStackTrace();
		}
	}

}
