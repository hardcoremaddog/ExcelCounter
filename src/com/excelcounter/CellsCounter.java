package com.excelcounter;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;

public class CellsCounter {

	private File all;
	private File table;

	private ArrayList<Carriage> carriages = new ArrayList<>();
	private ArrayList<Integer> values = new ArrayList<>();

	public CellsCounter(File all, File table) {
		this.all = all;
		this.table = table;
	}

	private final String TSEH_5 = "Сварочно - сборочный цех № 5";
	private final String TSEH_10 = "Электромонтажный цех № 10";
	private final String TSEH_51 = "Цех инструмента и оснастки № 51";
	private final String TSEH_121 = "Прессово-заготовительный цех № 121";
	private final String TSEH_217 = "Вагоносборочный цех № 217";
	private final String TSEH_317 = "Цех по сборке тележек и кузовов  № 317";
	private final String TSEH_416 = "Механосборочный  цех № 416";
	private final String TSEH_517 = "Цех окраски вагонов № 517";
	private final String VVMMZ = "Производственный участок";
	private final String UVK = "Управление внешней комплектации";
	private final String OMZK = "Отдел межзаводской кооперации";
	private final String OMO = "Отдел материального обеспечения";
	private final String MMZ02 = "02ММЗ ММЗ Цех 02";

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
		readMain(all, table);
	}

	private void readMain(File all, File table) {
		try {
			FileInputStream allFileInputStream = new FileInputStream(all);
			FileInputStream tableFileInputStream = new FileInputStream(table);

			XSSFWorkbook allWorkBook = new XSSFWorkbook(allFileInputStream);
			XSSFWorkbook tableWorkBook = new XSSFWorkbook(tableFileInputStream);

			XSSFSheet allSheet = allWorkBook.getSheetAt(0);
			XSSFSheet tableSheet = allWorkBook.getSheetAt(0);

			System.out.println("*********************");
			for (Row row : allSheet) {
				for (Cell cell : row) {
					if (cell.getCellType() == CellType.STRING) {
						if (cell.getStringCellValue().startsWith("(")) {
							String carriageName = cell.getStringCellValue();
							Carriage carriage = new Carriage(carriageName);
							System.out.println(cell.getStringCellValue() + "\n");
							readColumn(cell.getColumnIndex(), allSheet, carriage);
							carriages.add(carriage);
							System.out.println("*********************");
						}
					}
				}
			}

			allFileInputStream.close();
			tableFileInputStream.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private void readColumn(int columnNum, XSSFSheet sheet, Carriage carriage) {
		int rowNum = 0;
		int count = 0;
		for (Row row : sheet) {
			Cell cell = row.getCell(columnNum);
			XSSFCellStyle cs = (XSSFCellStyle) cell.getCellStyle();
			XSSFFont font = cs.getFont();

			byte[] redRBG = new byte[] {(byte) 255, (byte) 192, (byte) 203};
			byte[] yellowRGB = new byte[] {(byte) 255, (byte) 236, (byte) 139};

			if (font.getBold() && font.getFontHeightInPoints() == 8) {
				System.out.println(row.getCell(0));
				System.out.println(count(rowNum, columnNum, sheet));
			}
			rowNum++;
		}
//		System.out.println("Количество строк: " + count + "\n");
	}

	private int count(int rowNum, int columnNum, XSSFSheet sheet) {
		int currentRowNum = 0;
		int count = 0;
		for (Row row : sheet) {
			if (currentRowNum <= rowNum) {
				currentRowNum++;
				continue;
			}

			Cell cell = row.getCell(columnNum);
			XSSFCellStyle cs = (XSSFCellStyle) cell.getCellStyle();
			XSSFFont font = cs.getFont();

			if (font.getBold() && font.getFontHeightInPoints() == 8) {
				return count;
			}
			count++;
		}
		return 0;
	}
}
