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
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;


public class CellsCounter {

	private File all;
	private ArrayList<Order> orders = new ArrayList<>();

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

	public CellsCounter(File all) {
		this.all = all;
	}

	public void run() {
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
	}

	private void readMain(File all) {
		try {
			FileInputStream allFileInputStream = new FileInputStream(all);
			XSSFWorkbook allWorkBook = new XSSFWorkbook(allFileInputStream);
			XSSFSheet allSheet = allWorkBook.getSheetAt(0);

			for (Row row : allSheet) {
				for (Cell cell : row) {
					if (cell.getCellType() == CellType.STRING) {
						if (cell.getStringCellValue().startsWith("(") || cell.getStringCellValue().startsWith("16")) {
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

	private List<Department> getSortedDepartmentsListFor765Table(ArrayList<Department> departments) {
		Department[] sortedDepartments = new Department[12];

		for (Department department : departments) {
			switch (department.getName()) {
				case UVK: {
					sortedDepartments[0] = department;
					break;
				}
				case OMO: {
					sortedDepartments[1] = department;
					break;
				}

				case OMZK: {
					sortedDepartments[2] = department;
					break;
				}

				case TSEH_5: {
					sortedDepartments[3] = department;
					break;
				}

				case TSEH_10: {
					sortedDepartments[4] = department;
					break;
				}

				case TSEH_51: {
					sortedDepartments[5] = department;
					break;
				}

				case TSEH_121: {
					sortedDepartments[6] = department;
					break;
				}

				case TSEH_217: {
					sortedDepartments[7] = department;
					break;
				}

				case TSEH_317: {
					sortedDepartments[8] = department;
					break;
				}

				case TSEH_416: {
					sortedDepartments[9] = department;
					break;
				}

				case TSEH_517: {
					sortedDepartments[10] = department;
					break;
				}

				case VVMMZ: {
					sortedDepartments[11] = department;
					break;
				}
			}
		}
		return Arrays.asList(sortedDepartments);
	}
}