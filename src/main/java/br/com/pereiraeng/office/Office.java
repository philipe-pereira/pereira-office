package br.com.pereiraeng.office;

import java.awt.Dimension;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.RandomAccessFile;
import java.math.BigInteger;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.EmptyFileException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbookFactory;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDocument1;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation;

import br.com.pereiraeng.core.StringUtils;
import br.com.pereiraeng.core.TimeUtils;
import br.com.pereiraeng.core.collections.ArrayUtils;
import br.com.pereiraeng.html.HTML;
import br.com.pereiraeng.math.timeseries.Reg;
import br.com.pereiraeng.math.timeseries.RegP;
import br.com.pereiraeng.math.timeseries.SrT;

public class Office {

	// ==================================================================
	// ============================== EXCEL =============================
	// ==================================================================

	public static final int MAX_COLUMNS_EXCEL = 16384;

	public static final String FALSO = "FALSE";

	// funções lógicas

	public static final String SE = "IF", E = "AND", PROCV = "VLOOKUP", EERRO = "ISERROR", SEERRO = "IFERROR",
			CONTSE = "COUNTIF";

	// funções matemáticas

	public static final String MULT = "PRODUCT", MIN = "MIN", MAX = "MAX", SOMA = "SUM", COS = "COS", RAIZ = "SQRT",
			MEDIASE = "AVERAGEIF", ABS = "ABS";

	// função de manipulação de textos

	public static final String CONCATENAR = "CONCATENATE";

	public static final String HYPERLINK = "HYPERLINK";

	/**
	 * Número máximo de caracteres do nome de uma aba do MS-Excel
	 */
	public static final int MAX_SHEET_LENGTH = 31;

	// ---------------------- XSSF - Arrays e matrizes ----------------------

	/**
	 * Função que transforma uma matriz em uma planilha do Excel
	 * 
	 * @param file         arquivo a ser criado com a planilha Excel
	 * @param table        matriz
	 * @param columnHeader vetor com o cabeçalho das colunas
	 * @param rowHeader    vetor com o cabeçalho das linhas
	 */
	public static void export(File file, Object[][] table, Object[] columnHeader, Object[] rowHeader) {
		export(file, null, table, columnHeader, rowHeader);
	}

	/**
	 * Função que transforma uma matriz em uma planilha do Excel, indicando-se o
	 * nome da aba
	 * 
	 * @param file         arquivo a ser criado com a planilha Excel
	 * @param sheetName    nome da aba
	 * @param table        matriz
	 * @param columnHeader vetor com o cabeçalho das colunas
	 * @param rowHeader    vetor com o cabeçalho das linhas
	 */
	public static void export(File file, String sheetName, Object[][] table, Object[] columnHeader,
			Object[] rowHeader) {
		XSSFWorkbook wb = getWB(file);

		XSSFSheet sh = null;
		if (sheetName == null)
			sh = wb.createSheet();
		else
			sh = wb.createSheet(sheetName);

		// cabeçalho das colunas
		XSSFRow row = sh.createRow(0);
		for (int i = 0; i < columnHeader.length; i++) {
			XSSFCell cell = row.createCell(i + 1);
			setCell(cell, columnHeader[i]);
		}

		for (int j = 0; j < table.length; j++) {
			row = sh.createRow(j + 1);

			// cabeçalho das linhas
			XSSFCell cell = row.createCell(0);
			setCell(cell, rowHeader[j]);

			for (int i = 0; i < table[j].length; i++) {
				cell = row.createCell(i + 1);
				setCell(cell, table[j][i]);
			}
		}

		try {
			FileOutputStream out = new FileOutputStream(file);
			wb.write(out);
			out.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	// =========================== LOAD TABLE ===========================

	/**
	 * Função que carrega nas listas o conteúdo da planilha Excel
	 * 
	 * @param file    arquivo da planilha Excel
	 * @param header  lista de palavras do cabeçalho
	 * @param content lista de vetor de objetos, cada vetor representando uma linha
	 */
	public static void getTable(File file, List<String> header, List<Object[]> content) {
		XSSFWorkbook wb = getWB(file);

		XSSFSheet sheet = wb.getSheetAt(0);

		// header
		XSSFRow row = sheet.getRow(0);
		int cn = row.getLastCellNum();
		for (int i = 0; i < cn; i++) {
			XSSFCell c = row.getCell(i);
			if (c != null)
				header.add(c.getStringCellValue());
		}

		// cells
		int rn = sheet.getLastRowNum();
		for (int r = 1; r <= rn; r++) {
			row = sheet.getRow(r);
			if (row != null) {
				Object[] ss = new Object[cn];
				boolean nnr = false;
				for (int i = 0; i < cn; i++) {
					XSSFCell cell = row.getCell(i);
					if (cell != null) {
						CellType type = cell.getCellType();
						boolean be = type != CellType.BLANK && type != CellType.ERROR;
						nnr |= be;
						if (be) {
							if (type == CellType.NUMERIC)
								ss[i] = cell.getNumericCellValue();
							else
								ss[i] = cell.getStringCellValue();
						} else
							ss[i] = null;
					} else
						ss[i] = null;
				}
				if (nnr) // se a linha é não-nula
					content.add(ss);
			}
		}

		try {
			wb.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	// ============================= REG -> XLSX =============================

	/**
	 * Função que exporta os dados de um objeto {@link RegP} para uma planilha Excel
	 * 
	 * @param file arquivo com terminação .xlsx para onde a planilha será exportada
	 * @param reg  objeto registro comportando os dados a serem exportados
	 * @param days se <code>true</code> os dias são organizados em diferentes
	 *             colunas, senão em uma coluna única
	 * @return
	 *         <ol start="0">
	 *         <li>não gerou a planilha;</i>
	 *         <li>gerou a planilha no formato XLSX;</i>
	 *         <li>gerou a planilha no formato CSV.</i>
	 *         </ol>
	 */
	public static int exportExcel(File file, Reg reg, boolean days) {
		XSSFWorkbook wb = new XSSFWorkbook();

		// estilos
		CellStyle numStyle = wb.createCellStyle();
		numStyle.setDataFormat((short) 2);

		// nomes das etiquetas
		String[] labels = new String[reg.length()];
		System.arraycopy(reg.getLabels(), 0, labels, 0, labels.length);

		// You can use all alphanumeric characters but not the following special
		// characters: \ , / , * , ? , : , [ , ].
		for (int i = 0; i < labels.length; i++)
			if (labels[i] != null)
				labels[i] = labels[i].replace('/', 'd').replace(':', 'd').replace('*', 'x');

		// nomes não poder ter mais de 31 caracteres (o excel faz uma truncagem qnd há,
		// o que pode fazer com que haja colisões...)
		for (int i = 0; i < labels.length; i++) {
			if (labels[i] != null ? labels[i].length() > MAX_SHEET_LENGTH : false) { // truncar
				labels[i] = labels[i].substring(0, 31);
				int k = 1;
				while (ArrayUtils.hasDuplicate(labels, i) != -1)
					labels[i] = String.format("%s (%d)", labels[i].substring(0, k > 9 ? 26 : 27), k++);
			}
		}

		int ok = 1;
		try {
			if (days) {
				// organizar por dias -> várias abas, cada uma com uma tag

				// estilos
				CellStyle dateStyle = wb.createCellStyle();
				dateStyle.setDataFormat((short) 14); // 0xe, "m/d/yy"

				CellStyle timeStyle = wb.createCellStyle();
				timeStyle.setDataFormat((short) 20); // 0x14, "h:mm"

				// espaçamento
				int block = -1;
				if (reg instanceof RegP)
					block = ((RegP) reg).getFreq() * 60;
				else
					block = reg.getMinFreq();

				for (int j = 0; j < labels.length; j++) {
					XSSFSheet sh = wb.createSheet(labels[j] == null ? String.format("Med %03d", j) : labels[j]);

					int day, d = -1, col = 0;
					for (Entry<Integer, float[]> e : reg.entrySet()) {
						Calendar c = TimeUtils.toCalendar(e.getKey());
						XSSFRow row = null;

						// ver se troca de coluna
						day = c.get(Calendar.DAY_OF_MONTH);
						if (day != d) {
							d = day;
							col++;

							// cabeçalho das colunas
							row = sh.getRow(0);
							if (row == null)
								row = sh.createRow(0);
							XSSFCell cell = row.createCell(col);
							cell.setCellStyle(dateStyle);
							cell.setCellValue(c);
						}

						int r = calendar2row(c, block) + 1;

						// cria ou pega a linha
						row = sh.getRow(r);
						if (row == null) {
							row = sh.createRow(r);
							// cabeçalho das linhas
							XSSFCell cell = row.createCell(0);
							cell.setCellStyle(timeStyle);
							cell.setCellValue(c);
						}

						float value = e.getValue()[j];
						if (!Float.isNaN(value)) {
							XSSFCell cell = row.createCell(col);
							cell.setCellStyle(numStyle);
							cell.setCellValue(value);
						}
					}
				}
			} else { // coluna única -> uma aba, cada coluna com uma tag

				// estilos
				CellStyle datetimeStyle = wb.createCellStyle();
				datetimeStyle.setDataFormat((short) 22); // 0x16, "m/d/yy h:mm"

				XSSFSheet sh = wb.createSheet("Medições");

				// cabeçalho
				XSSFRow row = sh.createRow(0);
				row.createCell(0).setCellValue("Data e hora");
				for (int j = 0; j < labels.length; j++)
					row.createCell(j + 1).setCellValue(labels[j] == null ? String.format("Med %03d", j) : labels[j]);

				// valores
				int r = 1;
				for (Map.Entry<Integer, float[]> e : reg.entrySet()) {
					row = sh.createRow(r++);

					Calendar c = TimeUtils.toCalendar(e.getKey());

					XSSFCell cell = row.createCell(0);
					cell.setCellStyle(datetimeStyle);
					cell.setCellValue(c);

					for (int j = 0; j < labels.length; j++) {
						float value = e.getValue()[j];
						if (!Float.isNaN(value)) {
							cell = row.createCell(j + 1);
							cell.setCellStyle(numStyle);
							cell.setCellValue(value);
						}
					}
				}
			}
		} catch (OutOfMemoryError e) {
			// se gastou mais memória do que podia, tenta CSV
			System.gc();
			ok = 2;
		}

		if (ok == 1) {
			try {
				FileOutputStream out = new FileOutputStream(file);
				wb.write(out);
				out.close();
				wb.close();
			} catch (IOException e) {
				e.printStackTrace();
				ok = 0;
			}
		} else
			ok = exportExcelCSV(file, reg) ? 2 : 0;

		return ok;
	}

	public static void export(File xlsx, String sheetName, SrT<Double> series) {
		XSSFWorkbook wb = new XSSFWorkbook();

		XSSFSheet sheet = wb.createSheet(sheetName);

		String[] labels = series.getLabels();

		XSSFRow row = sheet.createRow(0);
		for (int i = 0; i < labels.length; i++)
			row.createCell(i + 1).setCellValue(labels[i]);

		int r = 1;
		for (Entry<Double, float[]> e : series.entrySet()) {
			row = sheet.createRow(r);

			row.createCell(0).setCellValue(e.getKey());
			float[] values = e.getValue();
			for (int i = 0; i < values.length; i++)
				row.createCell(i + 1).setCellValue(values[i]);

			r++;
		}

		try {
			FileOutputStream fos = new FileOutputStream(xlsx);
			wb.write(fos);
			wb.close();
			fos.close();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	// ============================= XLSX -> REG =============================

	public static RegP importExcel(File file) {
		XSSFWorkbook wb = Office.getWB(file);

		XSSFSheet sheet = wb.getSheetAt(0);
		int rows = sheet.getLastRowNum();

		XSSFRow row = sheet.getRow(0);
		LinkedList<String> tags = new LinkedList<>();
		for (int col = 1; col < row.getLastCellNum(); col++)
			tags.add(row.getCell(col).getStringCellValue());

		RegP out = new RegP(tags.size());

		for (int i = 1; i <= rows; i++) {
			row = sheet.getRow(i);
			Date d = row.getCell(0).getDateCellValue();
			for (int col = 1; col < row.getLastCellNum(); col++) {
				out.put(TimeUtils.date2Calendar(d), col - 1, (float) row.getCell(col).getNumericCellValue());
			}
		}

		try {
			wb.close();
		} catch (IOException e) {
			e.printStackTrace();
		}

		return out;
	}

	// ============================= REG -> CSV =============================

	public static boolean exportExcelCSV(File file, Reg regs) {
		RandomAccessFile raf = null;

		try {
			raf = new RandomAccessFile(file.getAbsoluteFile() + ".csv", "rw");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			return false;
		}

		// nomes das etiquetas
		String[] labels = regs.getLabels();

		// coluna única -> uma aba, cada coluna com uma tag

		// cabeçalho
		try {
			for (int j = 0; j < labels.length; j++)
				raf.writeBytes(";" + (labels[j] == null ? String.format("Med %03d", j) : labels[j]));

			for (Map.Entry<Integer, float[]> e : regs.entrySet()) {
				Calendar c = TimeUtils.toCalendar(e.getKey());
				raf.writeBytes(String.format("\r\n%1$td-%1$tm-%1$tY %1$tT", c));

				float[] reg = e.getValue();
				for (int j = 0; j < labels.length; j++) {
					if (Float.isNaN(reg[j]))
						raf.writeBytes(";-");
					else
						raf.writeBytes(String.format(";%g", reg[j]));
				}
			}

			raf.close();
		} catch (IOException e) {
			e.printStackTrace();
			return false;
		} catch (OutOfMemoryError e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	// =========================== HTML table -> XLSX ===========================

	/**
	 * Função que exporta uma tabela inteira de um documento HTML para uma planilha
	 * do Excel
	 * 
	 * @param file       arquivo XLSX de destino
	 * @param html       código HTML
	 * @param tableIndex índice da tabela, zero-based
	 */
	public static void exportExcel(File file, String html, int tableIndex) {
		// procurar no HTML inteiro onde começa e termina a tabela
		int[] limiters = StringUtils.getLimits(html, 0, html.length(), tableIndex, HTML.PATTERN_TABLE_OPENING_TAG,
				HTML.PATTERN_TABLE_CLOSING_TAG);

		XSSFWorkbook wb = Office.getWB(file);
		XSSFSheet sh = wb.createSheet();
		int rowCount = 0;

		int start = limiters[0];
		while (true) {
			// próxima linha
			int[] sublimiters = StringUtils.getLimits(html, start, limiters[1], 0, HTML.PATTERN_ROW_OPENING_TAG,
					HTML.PATTERN_ROW_CLOSING_TAG);

			if (sublimiters == null)
				break;

			// ler colunas
			String[] columns = StringUtils.getContent(html, sublimiters[0], sublimiters[1], HTML.PATTERN_COLUMN);
			for (int i = 0; i < columns.length; i++) {
				columns[i] = columns[i].replaceAll("<!--.+?-->", "");
				columns[i] = columns[i].replaceAll("<.+?>", "");
			}

			// cabeçalho das colunas
			XSSFRow row = sh.createRow(rowCount++);
			for (int column = 0; column < columns.length; column++) {
				XSSFCell cell = row.createCell(column);
				Office.setCell(cell, columns[column].trim());
			}

			// a próxima linha começa no final desta
			start = sublimiters[1];
		}

		// fechar arquivo
		try {
			FileOutputStream out = new FileOutputStream(file);
			wb.write(out);
			out.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	// =========================== AUXILIARES ===========================

	public static void setCell(XSSFCell cell, Object object) {
		if (object == null)
			return;
		if (object instanceof String)
			cell.setCellValue((String) object);
		else if (object instanceof Number)
			cell.setCellValue(((Number) object).doubleValue());
		else if (object instanceof Calendar)
			cell.setCellValue((Calendar) object);
		else if (object instanceof Date)
			cell.setCellValue((Date) object);
		else if (object instanceof Boolean)
			cell.setCellValue((Boolean) object ? "S" : "N");
		else
			cell.setCellValue(object.toString());
	}

	public static Object getCell(XSSFCell cell, XSSFWorkbook workbook) {
		switch (cell.getCellType()) {
		case STRING:
			return cell.getStringCellValue();
		case BOOLEAN:
			return cell.getBooleanCellValue();
		case NUMERIC:
			return cell.getNumericCellValue();
		case FORMULA:
			FormulaEvaluator evaluator = null;
			if (workbook != null)
				evaluator = workbook.getCreationHelper().createFormulaEvaluator();
			if (evaluator != null)
				return getValue(evaluator.evaluate(cell));
			else
				return cell.getCellFormula();
		default:
			return null;
		}
	}

	private static Object getValue(CellValue value) {
		switch (value.getCellType()) {
		case STRING:
			return value.getStringValue();
		case NUMERIC:
			return value.getNumberValue();
		case BOOLEAN:
			return value.getBooleanValue();
		default:
			return null;
		}
	}

	/**
	 * Função que abre uma planilha do MS-Excel (2003-...) no modo somente leitura
	 * ou cria uma novo Workbook, em modo escrita, caso o arquivo não exista
	 * 
	 * @param filePath sequência de caracteres indicando o caminho até a planilha
	 * @return objeto da planilha
	 */
	public static XSSFWorkbook getWB(String filePath) {
		return getWB(new File(filePath));
	}

	/**
	 * Função que abre uma planilha do MS-Excel (2003-...) no modo somente leitura
	 * ou cria uma novo Workbook, em modo escrita, caso o arquivo não exista
	 * 
	 * @param file arquivo da planilha
	 * @return objeto da planilha
	 */
	public static XSSFWorkbook getWB(File file) {
		return getWB(file, false);
	}

	/**
	 * Função que abre uma planilha do MS-Excel (2003-...)
	 * 
	 * @param file    arquivo da planilha
	 * @param rewrite <code>true</code> para modo escrita e leitura,
	 *                <code>false</code> para somente leitura
	 * @return objeto da planilha
	 */
	public static XSSFWorkbook getWB(File file, boolean rewrite) {
		if (file.exists()) {
			if (rewrite) { // se o arquivo já existir, sobreescrever
				try {
					return new XSSFWorkbook(new FileInputStream(file));
				} catch (IOException e2) {
					e2.printStackTrace();
				}
			} else { // mode somente leitura
				try {
					Workbook workbook = WorkbookFactory.create(file, null, true);
					if (workbook instanceof XSSFWorkbook) {
						return (XSSFWorkbook) workbook;
					} else {
						workbook.close();
					}
				} catch (EncryptedDocumentException | IOException e) {
					e.printStackTrace();
				}
			}
			return null;
		} else // se o arquivo não existe, criar
			return new XSSFWorkbook();
	}

	public static HSSFWorkbook getHWB(File file, boolean rewrite) {
		if (file.exists()) {
			if (rewrite) { // se o arquivo já existir, sobreescrever
				try {
					return new HSSFWorkbook(new FileInputStream(file));
				} catch (IOException e2) {
					e2.printStackTrace();
					return null;
				}
			} else { // mode somente leitura
				try {
					return HSSFWorkbookFactory.createWorkbook(new POIFSFileSystem(file));
				} catch (EmptyFileException | IOException e1) {
					e1.printStackTrace();
					return null;
				}
			}
		} else // se o arquivo não existe, criar
			return new HSSFWorkbook();
	}

	public static XSSFRow getOrCreateRow(XSSFSheet sheet, int r) {
		XSSFRow out = sheet.getRow(r);
		return out != null ? out : sheet.createRow(r);
	}

	public static XSSFCell getOrCreateCell(XSSFRow row, int c) {
		XSSFCell out = row.getCell(c);
		return out != null ? out : row.createCell(c);
	}

	public static XSSFSheet getOrCreateSheet(XSSFWorkbook wb, String s) {
		XSSFSheet out = wb.getSheet(s);
		return out != null ? out : wb.createSheet(s);
	}

	/**
	 * Função que a partir do número da coluna retorna a letra correpondente da
	 * coluna da planilha
	 * 
	 * @param i número da coluna
	 * @return identificação da coluna
	 */
	public static String cl(int i) {
		return (i > 25 ? ((char) (i / 26 + 64)) + "" : "") + ((char) (i % 26 + 65)) + "";
	}

	/**
	 * Função que converte um horário do dia em um número inteiro representativo. É
	 * a função inversa de {@link ExportExcel#row2calendar(int, int)}.
	 * 
	 * @param c     objeto {@link Calendar} que representa o horário
	 * @param block tempo, em segundos, entre duas medições do {@link RegP#getFreq()
	 *              registro periódico}
	 * @return inteiro que representa um horário do dia
	 */
	public static int calendar2row(Calendar c, int block) {
		return (c.get(Calendar.HOUR_OF_DAY) * 3600 + c.get(Calendar.MINUTE) * 60 + c.get(Calendar.SECOND)) / block;
	}

	/**
	 * Função que converte um número inteiro em um objeto que representa uma hora. É
	 * a função inversa de {@link ExportExcel#calendar2row(Calendar, int )}.
	 * 
	 * @param row   inteiro que representa um horário do dia
	 * @param block tempo, em segundos, entre duas medições do {@link RegP#getFreq()
	 *              registro periódico}
	 * @return objeto {@link Calendar} que representa o horário
	 */
	public static Calendar row2calendar(int row, int block) {
		Calendar c = new GregorianCalendar(1970, 0, 1, 0, 0);
		c.add(Calendar.SECOND, row * block);
		return c;
	}

	/**
	 * Função que converte o número que designa tempo no Excel (número decimal de
	 * dias desde 1/1/1900) no objeto {@link Calendar} correspondente
	 * 
	 * @param numeric número que designa tempo no Excel
	 * @return {@link Calendar} correspondente
	 */
	public static Calendar numeric2calendar(double numeric) {
		Calendar out = Calendar.getInstance();
		out.setTimeInMillis((long) (numeric * 86400000L) - 2209150800000L);
		return out;
	}

	// ==================================================================
	// ============================== WORD ==============================
	// ==================================================================

	public static final Dimension A4_POINT = new Dimension(595, 842);

	public static void changePaperOrientation(XWPFDocument doc) {
		CTDocument1 document = doc.getDocument();
		CTBody body = document.getBody();

		if (!body.isSetSectPr())
			body.addNewSectPr();

		CTSectPr section = body.getSectPr();

		if (!section.isSetPgSz())
			section.addNewPgSz();

		CTPageSz pageSize = section.getPgSz();

		pageSize.setW(BigInteger.valueOf(A4_POINT.height * 20));
		pageSize.setH(BigInteger.valueOf(A4_POINT.width * 20));

		pageSize.setOrient(STPageOrientation.LANDSCAPE);
	}

	public static XWPFTableCell getOrCreateCell(XWPFTableRow xwpfTableRow, int c) {
		XWPFTableCell cell = xwpfTableRow.getCell(c);
		if (cell == null)
			cell = xwpfTableRow.addNewTableCell();
		return cell;
	}

	public static void spanCellsAcrossRow(XWPFTableRow row, int colNum, int span) {
		XWPFTableCell cell = row.getCell(colNum);
		CTTc cttc = cell.getCTTc();
		if (!cttc.isSetTcPr())
			cttc.addNewTcPr();
		CTTcPr tcPr = cttc.getTcPr();
		tcPr.addNewGridSpan();
		tcPr.getGridSpan().setVal(BigInteger.valueOf((long) span));
	}

	public static void mergeCellsVertically(XWPFTable table, int col, int fromRow, int toRow) {
		for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
			XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
			if (rowIndex == fromRow) // The first merged cell is set with RESTART merge value
				cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
			else // Cells which join (merge) the first one, are set with CONTINUE
				cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
		}
	}

	public static void mergeCellsVertically(XWPFTable table, int fromRow, int toRow, int[] cols) {
		for (int rowIndex = fromRow, i = 0; rowIndex <= toRow; rowIndex++, i++) {
			XWPFTableCell cell = table.getRow(rowIndex).getCell(cols[i]);
			if (rowIndex == fromRow) // The first merged cell is set with RESTART merge value
				cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
			else // Cells which join (merge) the first one, are set with CONTINUE
				cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
		}
	}

	public static XWPFParagraph addBl(XWPFDocument doc, String text) {
		return addBl(doc, text, true);
	}

	public static XWPFParagraph addBl(XWPFDocument doc, String text, boolean center) {
		XWPFParagraph paragraph = new XWPFParagraph(CTP.Factory.newInstance(), doc);

		paragraph.setSpacingAfter(0);
		paragraph.setSpacingAfterLines(0);
		paragraph.setSpacingBetween(1.);
		if (center)
			paragraph.setAlignment(ParagraphAlignment.CENTER);

		XWPFRun run = paragraph.createRun();
		run.setFontFamily("Arial");
		run.setFontSize(10);

		String[] ts = text.split("\n");
		run.setText(ts[0]);
		for (int i = 1; i < ts.length; i++) {
			run.addBreak();
			run.setText(ts[i]);
		}
		return paragraph;
	}

	public static XWPFParagraph format(XWPFDocument doc, String text) {
		XWPFParagraph paragraph = new XWPFParagraph(CTP.Factory.newInstance(), doc);

		paragraph.setSpacingAfter(0);
		paragraph.setSpacingAfterLines(0);
		paragraph.setSpacingBetween(1.);
		paragraph.setAlignment(ParagraphAlignment.CENTER);

		XWPFRun run = paragraph.createRun();
		run.setFontFamily("Arial");
		run.setFontSize(10);

		run.setText(text);

		return paragraph;
	}

	public static void setColumnWidth(XWPFTable table, int rowB, int rowE, int column, int width) {
		if (width < 1)
			return;
		for (int r = rowB; r < rowE; r++) {
			XWPFTableRow row = table.getRow(r);
			XWPFTableCell cell = row.getCell(column);
			CTTblWidth cellWidth = cell.getCTTc().addNewTcPr().addNewTcW();
			CTTcPr pr = cell.getCTTc().addNewTcPr();
			pr.addNewNoWrap();
			cellWidth.setW(BigInteger.valueOf(width));
		}
	}

	public static void changePaperMargins(XWPFDocument doc, int margin) {
		CTSectPr sectPr = doc.getDocument().getBody().addNewSectPr();
		CTPageMar pageMar = sectPr.addNewPgMar();
		pageMar.setLeft(BigInteger.valueOf(margin));
		pageMar.setTop(BigInteger.valueOf(margin));
		pageMar.setRight(BigInteger.valueOf(margin));
		pageMar.setBottom(BigInteger.valueOf(margin));
	}

	// ==================================================================
	// ========================== POWER POINT ===========================
	// ==================================================================

	public static final Dimension SLIDE_POINT = new Dimension(717, 538);

}