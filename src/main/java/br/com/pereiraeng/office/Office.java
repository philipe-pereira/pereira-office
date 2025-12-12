package br.com.pereiraeng.office;

import java.awt.Dimension;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.RandomAccessFile;
import java.math.BigInteger;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import javax.swing.AbstractListModel;
import javax.swing.JTable;
import javax.swing.table.AbstractTableModel;
import javax.swing.tree.DefaultMutableTreeNode;

import org.apache.poi.EmptyFileException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbookFactory;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory;
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

import br.com.pereiraeng.math.timeseries.Reg;
import br.com.pereiraeng.math.timeseries.RegP;
import br.com.pereiraeng.core.StringUtils;
import br.com.pereiraeng.core.TimeUtils;
import br.com.pereiraeng.core.collections.ArrayUtils;
import br.com.pereiraeng.html.HTML;

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

	// ----------------------------- SEM O XSSF -----------------------------

	/**
	 * Função que cria um <code>String</code> no formato a ser inserido num arquivo
	 * do Excel a partir de uma conjunto de células de uma tabela
	 * 
	 * @param cell células da tabela
	 * @return <code>String</code> formatado
	 */
	public static String gerarExcel(String[][] cell) {
		String s = "";
		for (int i = 0; i < cell.length; i++) {
			for (int j = 0; j < cell[i].length; j++)
				s += (cell[i][j] + "\t");
			s += "\n";
		}
		return s;
	}

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

	// ---------------------- XSSF - TableModel ----------------------

	/**
	 * Função que transforma uma tabela gráfica {@link JTable} em uma planilha do
	 * Excel
	 * 
	 * @param file       arquivo a ser criado com a planilha Excel
	 * @param tableModel modelo de tabela contendo o conteúdo e o cabeçalho das
	 *                   colunas
	 */
	public static void export(File file, AbstractTableModel tableModel) {
		export(file, "001", tableModel);
	}

	public static void export(File file, String sheetName, AbstractTableModel tableModel) {
		export(file, sheetName, tableModel, null);
	}

	/**
	 * Função que transforma uma tabela gráfica {@link JTable} em uma planilha do
	 * Excel
	 * 
	 * @param file           arquivo a ser criado com a planilha Excel
	 * @param sheetName      nome da folha da planilha
	 * @param tableModel     modelo de tabela contendo o conteúdo da tabela e o
	 *                       cabeçalho das colunas
	 * @param rowHeaderModel modelo da lista contendo o cabeçalho das linhas
	 */
	public static void export(File file, String sheetName, AbstractTableModel tableModel,
			AbstractListModel<?> rowHeaderModel) {
		export(file, new String[] { sheetName }, new AbstractTableModel[] { tableModel },
				new AbstractListModel[] { rowHeaderModel });
	}

	/**
	 * Função que transforma tabelas gráficas {@link JTable} em uma planilha do
	 * Excel
	 * 
	 * @param file            arquivo a ser criado com a planilha Excel
	 * @param tableModels     vetor de modelos de tabelas contendo o conteúdo e o
	 *                        cabeçalho das colunas
	 * @param rowHeaderModels vetor de modelos da lista contendo o cabeçalho das
	 *                        linhas
	 */
	public static void export(File file, String[] sheetNames, AbstractTableModel[] tableModels,
			AbstractListModel<?>[] rowHeaderModels) {
		XSSFWorkbook wb = getWB(file);

		for (int k = 0; k < tableModels.length; k++)
			writeSheet(wb, sheetNames[k], tableModels[k], rowHeaderModels[k]);

		try {
			FileOutputStream out = new FileOutputStream(file);
			wb.write(out);
			out.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private static XSSFSheet writeSheet(XSSFWorkbook wb, String sheetName, AbstractTableModel tableModel,
			AbstractListModel<?> rowHeaderModel) {
		// se houver o cabeçalho das linhas, deslocar uma coluna para ele
		int rh = rowHeaderModel != null ? 1 : 0;

		XSSFSheet sh = wb.createSheet(sheetName);

		// cabeçalho das colunas
		XSSFRow row = sh.createRow(0);
		for (int i = 0; i < tableModel.getColumnCount(); i++) {
			XSSFCell cell = row.createCell(i + rh);
			String columnName = tableModel.getColumnName(i);
			cell.setCellValue(columnName);
		}

		for (int j = 0; j < tableModel.getRowCount(); j++) {
			row = sh.createRow(j + 1);

			// cabeçalho das linhas
			if (rowHeaderModel != null)
				setCell(row.createCell(0), rowHeaderModel.getElementAt(j));

			// conteúdo
			for (int i = 0; i < tableModel.getColumnCount(); i++)
				setCell(row.createCell(i + rh), tableModel.getValueAt(j, i));
		}

		return sh;
	}

	// =========================== TREE-TABLE ===========================

	/**
	 * Função que exporta uma tabela-árvore para uma planilha Excel (a arborescência
	 * será convertida em uma série de linhas, onde os nós superiores mesclam seus
	 * filhos)
	 * 
	 * @param file       arquivo a ser criado com a planilha do MS-Excel
	 * @param sheetName  nome da folha da planilha
	 * @param attm       modelo de tabela-árvore contendo o conteúdo da tabela e o
	 *                   cabeçalho das colunas
	 * @param treeLevels nome das colunas da árvore (o tamanho deste vetor deve ser
	 *                   igual à profundidade da árvore)
	 */
	public static void export(File file, String sheetName, AdvancedTreeTableModel attm, String... treeLevels) {
		String[] valueColumns = new String[attm.getColumnCount() - 1];
		for (int j = 0; j < valueColumns.length; j++)
			valueColumns[j] = attm.getColumnName(j + 1);
		export(file, sheetName, (DefaultMutableTreeNode) attm.getRoot(), attm.getTableData(), valueColumns, treeLevels);
	}

	public static void main(String[] args) {
		DefaultMutableTreeNode root = new DefaultMutableTreeNode("1");
		DefaultMutableTreeNode node2 = new DefaultMutableTreeNode("2");
		root.add(node2);
		DefaultMutableTreeNode node3 = new DefaultMutableTreeNode("3");
		node2.add(node3);
		DefaultMutableTreeNode node4 = new DefaultMutableTreeNode("4");
		node3.add(node4);
		DefaultMutableTreeNode node5 = new DefaultMutableTreeNode("5");
		node2.add(node5);

		Map<Object, Object[]> table = new HashMap<>();
		table.put("1", new Object[] { "Um", 1, "Un" });
		table.put("2", new Object[] { "Dois", 2, "Deux" });
		table.put("3", new Object[] { "Três", 3, "Trois" });
		table.put("4", new Object[] { "Quatro", 4, "Quatre" });
		table.put("5", new Object[] { "Cinco", 5, "Cinq" });

		export(new File("test.xlsx"), "test", root, table, null);
	}

	/**
	 * Função que exporta uma tabela-árvore para uma planilha Excel (a arborescência
	 * será convertida em uma série de linhas, onde os nós superiores mesclam seus
	 * filhos)
	 * 
	 * @param file         arquivo a ser criado com a planilha Excel do MS-Excel
	 * @param sheetName    nome da folha da planilha
	 * @param treeLevels   nome das colunas da árvore (o tamanho deste vetor deve
	 *                     ser igual à profundidade da árvore)
	 * @param root         raiz da árvore
	 * @param data         tabela com os dados de cada nó
	 * @param valueColumns demais colunas, respectivas aos valores de cada nó (por
	 *                     ser <code>null</code>, e neste caso não haverá cabeçalho
	 *                     para os valores)
	 */
	public static void export(File file, String sheetName, DefaultMutableTreeNode root, Map<Object, Object[]> data,
			String[] valueColumns, String... treeLevels) {
		// dados

		int depth = root.getDepth();
		if (treeLevels.length == 0) {
			treeLevels = new String[depth];
			for (int i = 0; i < treeLevels.length; i++)
				treeLevels[i] = String.format("Nível %02d", i);
		}

		// workbook e sheet
		XSSFWorkbook wb = Office.getWB(file);

		XSSFSheet sh = wb.createSheet(sheetName);
		sh.createFreezePane(depth, 1);

		// cabeçalho da coluna da árvore
		XSSFRow row = sh.createRow(0);

		XSSFCell cell = null;
		for (int k = 0; k < treeLevels.length; k++) {
			cell = row.createCell(k);
			cell.setCellValue(treeLevels[k]);
		}

		if (valueColumns != null)
			for (int j = 0; j < valueColumns.length; j++)
				row.createCell(treeLevels.length + j).setCellValue(valueColumns[j] == null);

		// conteúdo
		DefaultMutableTreeNode[] nodes = new DefaultMutableTreeNode[depth];
		int[] rst = new int[depth];
		int[] starts = new int[depth];
		writeLine(sh, root, data, 1, nodes, rst, starts);

		// escrever arquivo
		try {
			FileOutputStream out = new FileOutputStream(file);
			wb.write(out);
			out.close();
		} catch (IOException exc) {
			exc.printStackTrace();
		}
	}

	private static int writeLine(XSSFSheet sh, DefaultMutableTreeNode node, Map<Object, Object[]> data, int rowIndex,
			DefaultMutableTreeNode[] nodes, int[] rst, int[] starts) {
		if (node.isLeaf()) {
			// se for uma folha, contém valores...

			int level = node.getLevel();
			Object obj = node.getUserObject();

			Object[] values = data.get(obj);

			XSSFRow row = sh.createRow(rowIndex);

			// folha
			XSSFCell cell = row.createCell(level - 1);
			Office.setCell(cell, obj);

			// medições
			for (int m = 0; m < values.length; m++) {
				cell = row.createCell(nodes.length + m);
				Office.setCell(cell, values[m]);
			}

			// valor da célula mesclada
			for (int i = rst.length - 1; i > 0; i--) {
				if (rst[i] == 0) { // quando a coluna atual começa...
					starts[i - 1] = rowIndex;
//					row.createCell(i - 1).setCellValue(nodes[i - 1].toString());
				} else
					break;
			}

			// mesclagem
//			for (int i = rst.length - 1; i > 0; i--) {
//				if (rst[i] == nodes[i - 1].getChildCount() - 1) {
//					// quando a coluna atual termina...
//					if (starts[i - 1] != rowIndex)
//						sh.addMergedRegion(new CellRangeAddress(starts[i - 1], rowIndex, i - 1, i - 1));
//				} else
//					break;
//			}
			rowIndex++;
			return rowIndex;
		} else {
			// se o nó possuir filhos...
			for (int l = 0; l < node.getChildCount(); l++) {
				// ler recursivamente os filhos
				DefaultMutableTreeNode child = (DefaultMutableTreeNode) node.getChildAt(l);
				int level = child.getLevel() - 1;
				nodes[level] = child;
				rst[level] = l;
				rowIndex = writeLine(sh, child, data, rowIndex, nodes, rst, starts);
			}
			return rowIndex;
		}
	}

	// =========================== SQL <-> EXCEL ===========================

	/**
	 * Função que exporta para o Excel os resultados de uma busca SQL
	 * 
	 * @param file arquivo de destino
	 * @param rs   resultado da busca
	 */
	public static void export(File file, ResultSet rs) {
		XSSFWorkbook wb = getWB(file);

		XSSFSheet sh = wb.createSheet();

		try {
			int col = rs.getMetaData().getColumnCount();

			// cabeçalho das colunas
			XSSFRow row = sh.createRow(0);
			for (int i = 0; i < col; i++)
				row.createCell(i).setCellValue(rs.getMetaData().getColumnName(i + 1));

			// conteúdo
			int r = 1;
			while (rs.next()) {
				row = sh.createRow(r);
				for (int i = 0; i < col; i++)
					setCell(row.createCell(i), rs.getObject(i + 1));
				r++;
			}
		} catch (SQLException e) {
			e.printStackTrace();
		}

		try {
			FileOutputStream out = new FileOutputStream(file);
			wb.write(out);
			out.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	/**
	 * Função que carrega o conteúdo da planilha Excel numa base de dados SQL. O
	 * nome da tabela SQL será o nome do arquivo
	 * 
	 * @param file arquivo da planilha Excel
	 * @param sql  objeto conector da base de dados
	 */
	public static void transferSQL(File file, SQLadapter sql) {
		XSSFWorkbook wb = getWB(file);

		XSSFSheet sheet = wb.getSheetAt(0);

		// header
		XSSFRow row = sheet.getRow(0);
		int cn = row.getLastCellNum();
		String[] cols = new String[cn];
		for (int i = 0; i < cols.length; i++)
			cols[i] = row.getCell(i).getStringCellValue().replaceAll("\\s+", "_");

		// tipos
		row = sheet.getRow(1);
		CellType[] types = new CellType[cn];
		for (int i = 0; i < cols.length; i++)
			types[i] = row.getCell(i).getCellType();

		// criar tabela
		String table = file.getName().replaceAll("\\s+", "_");
		StringBuilder s = new StringBuilder(String.format("CREATE TABLE IF NOT EXISTS `%s` (", table));
		for (int i = 0; i < cols.length; i++)
			s.append(String.format("`%s` %s, ", cols[i], types[i] == CellType.NUMERIC ? "float" : "text"));
		s.setLength(s.length() - 2);
		s.append(") ENGINE=InnoDB DEFAULT CHARSET=utf8;");
		sql.update(s.toString());

		// cells
		int rn = sheet.getLastRowNum();
		s = new StringBuilder(
				String.format("INSERT INTO `%s`(`%s`) VALUES ", table, StringUtils.addSeparator(cols, "`, `")));
		for (int r = 1; r <= rn; r++) {
			row = sheet.getRow(r);
			if (row != null) {
				StringBuilder ss = new StringBuilder();
				boolean nnr = false;
				for (int i = 0; i < cols.length; i++) {
					XSSFCell cell = row.getCell(i);
					if (cell != null) {
						CellType type = cell.getCellType();
						boolean be = type != CellType.BLANK && type != CellType.ERROR;
						nnr |= be;
						if (be) {
							if (types[i] == CellType.NUMERIC)
								ss.append(cell.getNumericCellValue());
							else {
								ss.append("'");
								ss.append(cell.getStringCellValue());
								ss.append("'");
							}
						} else
							ss.append("NULL");
					} else
						ss.append("NULL");
					ss.append(",");
				}
				if (nnr) { // se a linha é não-nula
					s.append("(");
					s.append(ss.substring(0, ss.length() - 1));
					s.append("), ");
				}
			}
		}
		sql.update(s.substring(0, s.length() - 2));

		try {
			wb.close();
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

	// ============================= CHART -> XLSX =============================

	public static void export(File file, Chart<?> chart) {
		XSSFWorkbook wb = new XSSFWorkbook();

		// nomes das etiquetas
		List<?> labels = chart.getKeyArray();

		for (int l = 0; l < labels.size(); l++) {
			XSSFSheet sh = wb.createSheet(labels.get(l).toString());

			XSSFRow row = sh.createRow(0);
			row.createCell(0).setCellValue("X");
			row.createCell(1).setCellValue("Y");

			Object obj = (Object) labels.get(l);
			Plotable plotable = chart.get(obj);
			if (plotable instanceof Cloud) {
				Cloud c = (Cloud) plotable;
				double[][] xy = c.getCoordinates();

				for (int j = 0; j < xy[0].length; j++) {
					row = sh.createRow(j + 1);
					row.createCell(0).setCellValue(xy[0][j]);
					row.createCell(1).setCellValue(xy[1][j]);
				}
			} else if (plotable instanceof CurveFamily) {
				CurveFamily cf = (CurveFamily) plotable;
				for (int k = 0; k < cf.size(); k++) {
					cf.setIndex(k);
					double[][] xy = cf.getCoordinates();

					for (int j = 0; j < xy[0].length; j++) {
						row = sh.createRow(j + 1);
						row.createCell(2 * k).setCellValue(xy[0][j]);
						row.createCell(2 * k + 1).setCellValue(xy[1][j]);
					}
				}
			}
		}

		try {
			FileOutputStream out = new FileOutputStream(file);
			wb.write(out);
			out.close();
			wb.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
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
	 * 
	 * @param filePath sequência de caracteres indicando o caminho até a planilha
	 * @return objeto da planilha
	 */
	public static XSSFWorkbook getWB(String filePath) {
		return getWB(new File(filePath));
	}

	/**
	 * Função que abre uma planilha do MS-Excel (2003-...) no modo somente leitura
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
					return null;
				}
			} else { // mode somente leitura TODO testar isso!
				try {
					return XSSFWorkbookFactory.createWorkbook(OPCPackage.create(file));
				} catch (EmptyFileException | IOException e1) {
					e1.printStackTrace();
					return null;
				}
			}
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

	/**
	 * Função que exporta para o Excel o conteúdo de toda uma tabela da base de
	 * dados
	 * 
	 * @param folder diretório onde será criado a planilha do Excel
	 * @param table  nome da tabela do banco de dados a ser exportada
	 */
	public void sql2xlsx(SQLadapter sql, String folder, String table) {
		ResultSet rs = null;

		XSSFWorkbook wb = new XSSFWorkbook();

		XSSFSheet sh = wb.createSheet("DB");
		XSSFRow row = sh.createRow(0);

		try {
			String query = "SELECT * FROM " + table;

			rs = sql.query(query);

			// cabeçalho
			ResultSetMetaData rsmd = rs.getMetaData();
			int columns = rsmd.getColumnCount();
			for (int j = 1; j <= columns; j++) {
				XSSFCell cell = row.createCell(j - 1);
				cell.setCellValue(rsmd.getColumnName(j));
			}

			// para cada entrada da BD
			int y = 0;

			while (rs.next()) {
				y++;
				row = sh.createRow(y);

				for (int j = 1; j <= columns; j++) {
					XSSFCell cell = row.createCell(j - 1);
					cell.setCellValue(rs.getString(j));
				}
			}

			rs.getStatement().close();
			rs.close();

			// gerar arquivo do Excel

			File f = new File(folder + "/" + table + ".xlsx");
			FileOutputStream out = new FileOutputStream(f);
			wb.write(out);
			out.close();
		} catch (SQLException | IOException e) {
			e.printStackTrace();
		}
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