package com.excelcounter;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;

public class CellsCounter {

	private File all;

	private ArrayList<Order> orders = new ArrayList<>();
	private ArrayList<Integer> values = new ArrayList<>();

	public CellsCounter(File all) {
		this.all = all;
	}

	private int TSEH_5_redCellsCount = 0;
	private int TSEH_10_redCellsCount = 0;
	private int TSEH_51_redCellsCount = 0;
	private int TSEH_121_redCellsCount = 0;
	private int TSEH_217_redCellsCount = 0;
	private int TSEH_317_redCellsCount = 0;
	private int TSEH_416_redCellsCount = 0;
	private int TSEH_517_redCellsCount = 0;
	private int VVMMZ_redCellsCount = 0;
	private int UVK_redCellsCount = 0;
	private int OMZK_redCellsCount = 0;
	private int OMO_redCellsCount = 0;
	private int MMZ02_redCellsCount = 0;

	public void run() {
		readMain(all);
	}

	private void readMain(File all) {
		try {
			FileInputStream allFileInputStream = new FileInputStream(all);
			XSSFWorkbook allWorkBook = new XSSFWorkbook(allFileInputStream);
			XSSFSheet allSheet = allWorkBook.getSheetAt(0);

			System.out.println("*********************");
			for (Row row : allSheet) {
				for (Cell cell : row) {
					if (cell.getCellType() == CellType.STRING) {
						if (cell.getStringCellValue().startsWith("(") || cell.getStringCellValue().startsWith("16")) {
							String orderNumber = cell.getStringCellValue();
							Order order = new Order(orderNumber);
							System.out.println(cell.getStringCellValue() + "\n");
							readColumn(cell.getColumnIndex(), allSheet, order);
							orders.add(order);
							System.out.println("*********************");
						}
					}
				}
			}

			allFileInputStream.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private void readColumn(int columnNum, XSSFSheet sheet, Order order) {
		for (int rowNum = 0; rowNum < sheet.getPhysicalNumberOfRows(); rowNum++) {
			Row row = sheet.getRow(rowNum);
			Cell cell = row.getCell(columnNum);
			XSSFCellStyle cs = (XSSFCellStyle) cell.getCellStyle();
			XSSFFont font = cs.getFont();

			if (font.getBold() && font.getFontHeightInPoints() == 8) {
				System.out.println(row.getCell(0));
				count(rowNum, columnNum, sheet, order);
			}
		}
	}

	private void count(int rowNum, int columnNum, XSSFSheet sheet, Order order) {
		int redCount = 0;
		int yellowCount = 0;
		for (int currentRowNum = rowNum + 1; currentRowNum < sheet.getPhysicalNumberOfRows(); currentRowNum++) {
			Row row = sheet.getRow(currentRowNum);
			Cell cell = row.getCell(columnNum);

			XSSFCellStyle cs = (XSSFCellStyle) cell.getCellStyle();
			XSSFFont font = cs.getFont();

			String redARGBHEX = "FFFFC0CB";
			String yellowARGBHEX = "FFFFEC8B";

			if (cell.getCellType() != CellType.BLANK) {
				if (cs.getFillForegroundColorColor().getARGBHex().equals(redARGBHEX)) {
					redCount++;
				}
				else if (cs.getFillForegroundColorColor().getARGBHex().equals(yellowARGBHEX)) {
					yellowCount++;
				}
			}

			if (font.getBold() && (font.getFontHeightInPoints() == 8 || font.getFontHeightInPoints() == 10)) {
				System.out.println("К: " + redCount);
				System.out.println("Ж: " + yellowCount);
				System.out.println("-----------------");
				return;
			}
		}
	}
}