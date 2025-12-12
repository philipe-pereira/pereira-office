package br.com.pereiraeng.office;

public interface LargeSheet {
	public void cellContent(int sheet, String cell, String content);

	public void rowContent(int sheet, int row, String[] content);
}
