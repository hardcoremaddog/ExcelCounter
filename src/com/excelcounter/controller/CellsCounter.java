package com.excelcounter.controller;

import com.excelcounter.model.Department;
import com.excelcounter.model.Order;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class CellsCounter {

	private File all;
	private File table;
	private ArrayList<Order> orders = new ArrayList<>();

	private final String TSEH_5 = "Сварочно - сборочный цех № 5";
	private final String TSEH_10 = "Электромонтажный цех № 10";
	private final String TSEH_51 = "Цех инструмента и оснастки № 51";
	private final String TSEH_121 = "Прессово-заготовительный цех № 121";
	private final String TSEH_217 = "Вагоносборочный цех № 217";
	private final String TSEH_317 = "Цех по сборке тележек и кузовов  № 317";
	private final String TSEH_416 = "Механосборочный  цех № 416";
	private final String TSEH_417 = "Цех изделий малых серий  № 417";
	private final String TSEH_517 = "Цех окраски вагонов № 517";
	private final String VVMMZ = "Производственный участок";
	private final String UVK = "Управление внешней комплектации";
	private final String OMZK = "Отдел межзаводской кооперации";
	private final String OMO = "Отдел материального обеспечения";
	private final String MMZ02 = "02ММЗ ММЗ Цех 02";

	public CellsCounter(File all) {
		this.all = all;
	}

	public CellsCounter(File all, File table) {
		this.all = all;
		this.table = table;
	}

	public void run(int param) {
		readMain(all);

		for (Order order : orders) {
			System.out.println("=====================");
			System.out.println(order.getName());
			System.out.println("=====================");
			for (Department department : order.getDepartments()) {
				System.out.println("--------------------------------");
				System.out.println(department.getName());
				System.out.println("Красных позиций: " + department.getRedCellsCount());
				System.out.println("Желтых позиций: " + department.getYellowCellsCount());
			}
			System.out.println("--------------------------------");
			System.out.println("\n \n \n");
		}

		if (table != null) {
			if (param != 0) {
				switch (param) {
					case 1: {
						writeResultTo765Table(table);
						break;
					}
					case 2: {
						writeResultTo753Table(table);
						break;
					}
					case 3: {
						writeResultToOrdersTable(table);
						break;
					}
				}
			}
		}
	}

	private void readMain(File all) {
		try {
			FileInputStream allFileInputStream = new FileInputStream(all);
			XSSFWorkbook allWorkBook = new XSSFWorkbook(allFileInputStream);
			XSSFSheet allSheet = allWorkBook.getSheetAt(0);

			for (int i = 0; i < allSheet.getPhysicalNumberOfRows(); i++) {
				Row row = allSheet.getRow(i);
				for (Cell cell : row) {
					if (cell.getCellType() == CellType.STRING) {
						if (cell.getStringCellValue().startsWith("(")
								|| cell.getStringCellValue().startsWith("При")
								|| cell.getStringCellValue().startsWith("16-")
								|| cell.getStringCellValue().startsWith("17-")
								|| cell.getStringCellValue().startsWith("18-")
								|| cell.getStringCellValue().startsWith("19-")
								|| cell.getStringCellValue().startsWith("20-")
								|| cell.getStringCellValue().startsWith("21-")
								|| cell.getStringCellValue().startsWith("22-")
								|| cell.getStringCellValue().startsWith("16_")
								|| cell.getStringCellValue().startsWith("17_")
								|| cell.getStringCellValue().startsWith("18_")
								|| cell.getStringCellValue().startsWith("19_")
								|| cell.getStringCellValue().startsWith("20_")
								|| cell.getStringCellValue().startsWith("21_")
								|| cell.getStringCellValue().startsWith("22_")) {
							String orderNumber = cell.getStringCellValue();
							Order order = new Order(orderNumber);
							readColumn(cell.getColumnIndex(), allSheet, order);
							orders.add(order);
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
				String departmentName = row.getCell(0).getStringCellValue();
				Department department = new Department(departmentName);
				count(rowNum, columnNum, sheet, order, department);
			}
		}
	}

	private void count(int rowNum, int columnNum, XSSFSheet sheet, Order order, Department department) {
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
				} else if (cs.getFillForegroundColorColor().getARGBHex().equals(yellowARGBHEX)) {
					yellowCount++;
				}
			}

			if (font.getBold() && (font.getFontHeightInPoints() == 8 || font.getFontHeightInPoints() == 10)) {
				department.setRedCellsCount(redCount);
				department.setYellowCellsCount(yellowCount);
				order.getDepartments().add(department);
				return;
			}
		}
	}

	private void writeResultTo765Table(File table) {
		try {
			FileInputStream tableFileInputStream = new FileInputStream(table);
			XSSFWorkbook tableWorkbook = new XSSFWorkbook(tableFileInputStream);
			XSSFSheet tableSheet = tableWorkbook.getSheetAt(0);

			for (int i = 4; i < 70; i++) {
				Row row = tableSheet.getRow(i);
				Cell orderCell = row.getCell(1);

				for (Order order : orders) {
					switch (orderCell.getCellType()) {
						case STRING: {
							if (orderCell.getStringCellValue().contains(order.getName())) {
								for (Department department : order.getDepartments()) {
									switch (department.getName()) {
										case UVK: {
											row.getCell(2).setCellValue(department.getYellowCellsCount());
											row.getCell(6).setCellValue(department.getRedCellsCount());
											break;
										}
										case OMO: {
											row.getCell(3).setCellValue(department.getYellowCellsCount());
											row.getCell(7).setCellValue(department.getRedCellsCount());
											break;
										}
										case OMZK: {
											row.getCell(4).setCellValue(department.getYellowCellsCount());
											row.getCell(8).setCellValue(department.getRedCellsCount());
											break;
										}
										case TSEH_5: {
											row.getCell(10).setCellValue(department.getRedCellsCount());
											break;
										}
										case TSEH_10: {
											row.getCell(11).setCellValue(department.getRedCellsCount());
											break;
										}
										case TSEH_51: {
											row.getCell(12).setCellValue(department.getRedCellsCount());
											break;
										}
										case TSEH_121: {
											row.getCell(13).setCellValue(department.getRedCellsCount());
											break;
										}
										case TSEH_217: {
											row.getCell(14).setCellValue(department.getRedCellsCount());
											break;
										}
										case TSEH_317: {
											row.getCell(15).setCellValue(department.getRedCellsCount());
											break;
										}
										case TSEH_416: {
											row.getCell(16).setCellValue(department.getRedCellsCount());
											break;
										}
										case TSEH_517: {
											row.getCell(17).setCellValue(department.getRedCellsCount());
											break;
										}
										case VVMMZ: {
											row.getCell(18).setCellValue(department.getRedCellsCount());
											break;
										}
									}
								}
							}
							break;
						}
					}
				}
			}

			FileOutputStream tableFileOutputStream = new FileOutputStream(table);
			tableWorkbook.write(tableFileOutputStream);

			tableWorkbook.close();
			tableFileInputStream.close();
			tableFileOutputStream.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private void writeResultTo753Table(File table) {
		try {
			FileInputStream tableFileInputStream = new FileInputStream(table);
			XSSFWorkbook tableWorkbook = new XSSFWorkbook(tableFileInputStream);
			XSSFSheet tableSheet = tableWorkbook.getSheetAt(0);

			for (int i = 4; i < 50; i++) {
				Row row = tableSheet.getRow(i);
				Cell orderCell = row.getCell(1);

				for (Order order : orders) {
					switch (orderCell.getCellType()) {
						case STRING: {
							if (orderCell.getStringCellValue().contains(order.getName())) {
								for (Department department : order.getDepartments()) {
									switch (department.getName()) {
										case UVK: {
											row.getCell(2).setCellValue(department.getYellowCellsCount());
											row.getCell(6).setCellValue(department.getRedCellsCount());
											break;
										}
										case OMO: {
											row.getCell(3).setCellValue(department.getYellowCellsCount());
											row.getCell(7).setCellValue(department.getRedCellsCount());
											break;
										}
										case OMZK: {
											row.getCell(4).setCellValue(department.getYellowCellsCount());
											row.getCell(8).setCellValue(department.getRedCellsCount());
											break;
										}
										case TSEH_5: {
											row.getCell(10).setCellValue(department.getRedCellsCount());
											break;
										}
										case TSEH_10: {
											row.getCell(11).setCellValue(department.getRedCellsCount());
											break;
										}
										case TSEH_51: {
											row.getCell(12).setCellValue(department.getRedCellsCount());
											break;
										}
										case TSEH_121: {
											row.getCell(13).setCellValue(department.getRedCellsCount());
											break;
										}
										case TSEH_217: {
											row.getCell(14).setCellValue(department.getRedCellsCount());
											break;
										}
										case TSEH_317: {
											row.getCell(15).setCellValue(department.getRedCellsCount());
											break;
										}
										case TSEH_416: {
											row.getCell(16).setCellValue(department.getRedCellsCount());
											break;
										}
										case TSEH_417: {
											row.getCell(17).setCellValue(department.getRedCellsCount());
											break;
										}
										case TSEH_517: {
											row.getCell(18).setCellValue(department.getRedCellsCount());
											break;
										}
										case VVMMZ: {
											row.getCell(19).setCellValue(department.getRedCellsCount());
											break;
										}
									}
								}
							}
							break;
						}
					}
				}
			}

			FileOutputStream tableFileOutputStream = new FileOutputStream(table);
			tableWorkbook.write(tableFileOutputStream);

			tableWorkbook.close();
			tableFileInputStream.close();
			tableFileOutputStream.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private void writeResultToOrdersTable(File table) {
		try {
			FileInputStream tableFileInputStream = new FileInputStream(table);
			XSSFWorkbook tableWorkbook = new XSSFWorkbook(tableFileInputStream);
			XSSFSheet tableSheet = tableWorkbook.getSheet("Сводная таблица");

			for (int i = 3; i < 100; i++) {
				Row row = tableSheet.getRow(i);
				Cell orderCell = row.getCell(2);

				for (Order order : orders) {
					switch (orderCell.getCellType()) {
						case STRING: {
							if (orderCell.getStringCellValue().contains(order.getName())) {
								for (Department department : order.getDepartments()) {
									switch (department.getName()) {
										case UVK: {
											row.getCell(6).setCellValue(department.getYellowCellsCount());
											row.getCell(3).setCellValue(department.getRedCellsCount());
											break;
										}
										case OMO: {
											row.getCell(7).setCellValue(department.getYellowCellsCount());
											row.getCell(4).setCellValue(department.getRedCellsCount());
											break;
										}
										case OMZK: {
											row.getCell(8).setCellValue(department.getYellowCellsCount());
											row.getCell(5).setCellValue(department.getRedCellsCount());
											break;
										}
										case TSEH_5: {
											row.getCell(27).setCellValue(department.getYellowCellsCount());
											row.getCell(19).setCellValue(department.getRedCellsCount());
											break;
										}
										case TSEH_10: {
											row.getCell(28).setCellValue(department.getYellowCellsCount());
											row.getCell(20).setCellValue(department.getRedCellsCount());
											break;
										}
										case TSEH_51: {
											row.getCell(29).setCellValue(department.getYellowCellsCount());
											row.getCell(21).setCellValue(department.getRedCellsCount());
											break;
										}
										case TSEH_121: {
											row.getCell(30).setCellValue(department.getYellowCellsCount());
											row.getCell(22).setCellValue(department.getRedCellsCount());
											break;
										}
										case TSEH_217: {
											row.getCell(31).setCellValue(department.getYellowCellsCount());
											row.getCell(23).setCellValue(department.getRedCellsCount());
											break;
										}
										case TSEH_317: {
											row.getCell(32).setCellValue(department.getYellowCellsCount());
											row.getCell(24).setCellValue(department.getRedCellsCount());
											break;
										}
										case TSEH_416: {
											row.getCell(33).setCellValue(department.getYellowCellsCount());
											row.getCell(25).setCellValue(department.getRedCellsCount());
											break;
										}

										case TSEH_517: {
											row.getCell(34).setCellValue(department.getYellowCellsCount());
											row.getCell(26).setCellValue(department.getRedCellsCount());
											break;
										}
									}
								}
							}
							break;
						}
					}
				}
			}

			FileOutputStream tableFileOutputStream = new FileOutputStream(table);
			tableWorkbook.write(tableFileOutputStream);

			tableWorkbook.close();
			tableFileInputStream.close();
			tableFileOutputStream.close();
		} catch (
				IOException e) {
			e.printStackTrace();
		}
	}
}