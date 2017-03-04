package com.accenture.util.avaya.phraseset;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Hashtable;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Utility class for editing Avaya dialog designer component phraseset.
 * 
 * 2016-11-05
 * 
 * @author daryl.v.lapuz
 * 
 */
public class PhrasesetEditor {

	public static void main(String[] args) {

		String workingDir = System.getProperty("user.dir");
		File workbook = new File(workingDir, "input.xlsx");
		List<Hashtable<String, String>> tableList = getTableList(workbook);
		System.out.println(tableList);
	}

	private static List<Hashtable<String, String>> getTableList(File inputFile) {

		List<Hashtable<String, String>> tableList = new ArrayList<Hashtable<String, String>>();

		XSSFWorkbook inputWorkbook = null;

		try {

			inputWorkbook = new XSSFWorkbook(inputFile);
			XSSFSheet inputSheet = inputWorkbook.getSheet("Input");

			int rowIndex = 1;
			XSSFRow row;
			Hashtable<String, String> table = null;

			while ((row = inputSheet.getRow(rowIndex)) != null) {
				table = new Hashtable<String, String>();

				XSSFRow header = inputSheet.getRow(0);
				Iterator<Cell> headerIter = header.cellIterator();

				int columnIndex = 0;

				while (headerIter.hasNext()) {

					Cell headerCell = headerIter.next();
					String strHeader = headerCell.getStringCellValue();

					Cell cell = row.getCell(columnIndex);

					String value = "";
					if (cell != null) {
						value = cell.getStringCellValue();
					}
					table.put(strHeader, value);
					columnIndex++;
				}
				tableList.add(table);
				rowIndex++;
			}

		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				inputWorkbook.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return tableList;
	}

}
