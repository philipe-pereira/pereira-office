package br.com.pereiraeng.office;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import java.util.LinkedList;

import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.ooxml.util.SAXHelper;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

@Deprecated
public class SheetHandler extends DefaultHandler {

	private SharedStringsTable sst;
	private LargeSheet ls;

	private int sheetNumber;
	private String cell;
	private int row;

	private boolean nextIsString;
	private StringBuilder lastContents;
	private LinkedList<String> rowContents;

	public SheetHandler(LargeSheet ls, boolean rowRead, File file) {
		this.ls = ls;
		if (rowRead)
			this.rowContents = new LinkedList<>();

		try {
			XSSFReader reader = new XSSFReader(OPCPackage.open(file));
			this.sst = reader.getSharedStringsTable();

			XMLReader parser = SAXHelper.newXMLReader();
			parser.setContentHandler(this);

			Iterator<InputStream> sheets = reader.getSheetsData();
			int c = 0;
			while (sheets.hasNext()) {
				InputStream sheet = sheets.next();
				InputSource sheetSource = new InputSource(sheet);
				this.setSheetNumber(c);
				parser.parse(sheetSource);
				sheet.close();
				c++;
			}
		} catch (IOException | OpenXML4JException | SAXException | ParserConfigurationException e) {
			e.printStackTrace();
		}
	}

	private void setSheetNumber(int sheetNumber) {
		this.sheetNumber = sheetNumber;
	}

	@Override
	public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
		// c => cellS

		if (name.equals("c")) {
			// Print the cell reference
			this.cell = attributes.getValue("r");

			// Figure out if the value is an index in the SST
			String cellType = attributes.getValue("t");
			nextIsString = cellType != null && cellType.equals("s");
		}

		if (name.equals("row"))
			this.row = Integer.parseInt(attributes.getValue("r"));

		// Clear contents cache
		lastContents = new StringBuilder();
	}

	@Override
	public void characters(char[] ch, int start, int length) {
		lastContents.append(new String(ch, start, length));
	}

	@Override
	public void endElement(String uri, String localName, String name) throws SAXException {
		// Process the last contents as required.
		// Do now, as characters() may be called more than once
		if (nextIsString) {
			int idx = Integer.parseInt(lastContents.toString());
			lastContents = new StringBuilder(sst.getEntryAt(idx).getT().toString());
			nextIsString = false;
		}

		switch (name) {
		case "v": // contents of a cell
			if (rowContents != null)
				rowContents.add(lastContents.toString());
			else
				ls.cellContent(sheetNumber, this.cell, lastContents.toString());
			break;

		case "row": // fim da linha
			if (rowContents != null) {
				ls.rowContent(sheetNumber, row, rowContents.toArray(new String[rowContents.size()]));
				rowContents.clear();
			}
			break;
		}
	}
}