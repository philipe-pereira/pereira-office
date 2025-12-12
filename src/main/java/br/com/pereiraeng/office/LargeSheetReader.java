package br.com.pereiraeng.office;

import java.io.File;

/**
 * ESSA CLASSE FICOU UMA MERDA: MUITA TROCA DE THREAD! CREIO QUE ISSO FEZ FUDER
 * O BAGULHO!
 * 
 * VER: http://en.wikipedia.org/wiki/Context_switch
 * 
 * @author Philipe Pereira
 *
 */
@Deprecated
public class LargeSheetReader implements Runnable, LargeSheet {

	private File file;
	private boolean rowRead;
	private int sheet;

	private int currentSheet;
	private String currentCell;
	private int currentRow;
	private String[] content;

	private Thread t;

	public LargeSheetReader(File file, boolean rowRead) {
		this(file, rowRead, -1);
	}

	public LargeSheetReader(File file, boolean rowRead, int sheet) {
		this.file = file;
		this.rowRead = rowRead;
		if (this.rowRead)
			content = new String[1];
		this.sheet = sheet;

		this.t = new Thread(this);
		this.t.start();

		synchronized (file) {
			try {
				file.wait();
			} catch (InterruptedException e) {
				e.printStackTrace();
			}
		}
	}

	public boolean hasNext() {
		return t.isAlive();
	}

	public String[] nextRow() {
		swapThread(this, file);

		return content;
	}

	public String nextCell() {
		swapThread(this, file);

		return content[0];
	}

	public int getCurrentSheet() {
		return currentSheet;
	}

	public int getCurrentRow() {
		return currentRow;
	}

	public String getCurrentCell() {
		return currentCell;
	}

	// --------------------------------------------------------------

	@Override
	public void cellContent(int currentSheet, String cell, String content) {
		this.currentSheet = currentSheet;
		if (currentSheet == this.sheet || this.sheet < 0) {
			this.currentCell = cell;
			this.content[0] = content;

			swapThread(file, this);
		}
	}

	@Override
	public void rowContent(int currentSheet, int row, String[] content) {
		this.currentSheet = currentSheet;
		if (currentSheet == this.sheet || this.sheet < 0) {
			this.currentRow = row;
			this.content = content;

			swapThread(file, this);
		}
	}

	@Override
	public void run() {
		new SheetHandler(this, rowRead, file);
	}

	/**
	 * Função que faz a troca do Thread
	 * 
	 * @param release se for {@link #file}, solta a execução do Thread principal e
	 *                segura a do leitor
	 * @param hold    se for {@link #file}, segura a execução do Thread principal e
	 *                solta a do leitor
	 */
	private void swapThread(Object release, Object hold) {
		synchronized (release) {
			release.notify();
		}
		synchronized (hold) {
			try {
				hold.wait();
			} catch (InterruptedException e) {
				e.printStackTrace();
			}
		}
	}
}
