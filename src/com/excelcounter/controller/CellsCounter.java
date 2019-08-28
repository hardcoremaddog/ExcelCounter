package com.excelcounter.controller;

import com.excelcounter.model.*;
import com.excelcounter.util.MapUtil;
import com.excelcounter.view.AdvancedGUI;
import com.excelcounter.view.GUI;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Files;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;

public class CellsCounter {
    GUI gui;
    AdvancedGUI advancedGUI;

    private boolean win95colors;

    private List<File> allFiles;
    private List<File> directories;
    private List<Day> days = new ArrayList<>();

    private File all;
    private File table;
    private File sbyt = null;
    private List<Order> orders = new ArrayList<>();
    private List<OrderSbyt> ordersSbyt = new ArrayList<>();

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
    private final String OMZK_new = "Отдел материального обеспечения и сервисных закупок";
    private final String OMO = "Отдел материального обеспечения";
    private final String MMZ02 = "02ММЗ ММЗ Цех 02";

    private List<Department> departmentsListFull = new ArrayList<>();

    public CellsCounter(File all, GUI gui) {
        this.all = all;
        this.gui = gui;
    }

    public CellsCounter(File all, File sbyt, GUI gui) {
        this.all = all;
        this.sbyt = sbyt;
        this.gui = gui;
    }

    public CellsCounter(File all, File table, File sbyt, GUI gui) {
        this.all = all;
        this.table = table;
        this.sbyt = sbyt;
        this.gui = gui;
    }

    public CellsCounter(List<File> allFiles, List<File> directories, File table, AdvancedGUI advancedGUI) {
        initDepartmentsList();
        this.allFiles = allFiles;
        this.directories = directories;
        this.table = table;
        this.advancedGUI = advancedGUI;
    }

    private void initDepartmentsList() {
        this.departmentsListFull.add(new Department(TSEH_5));
        this.departmentsListFull.add(new Department(TSEH_10));
        this.departmentsListFull.add(new Department(TSEH_51));
        this.departmentsListFull.add(new Department(TSEH_121));
        this.departmentsListFull.add(new Department(TSEH_217));
        this.departmentsListFull.add(new Department(TSEH_317));
        this.departmentsListFull.add(new Department(TSEH_417));
        this.departmentsListFull.add(new Department(TSEH_517));
        this.departmentsListFull.add(new Department(VVMMZ));
        this.departmentsListFull.add(new Department(UVK));
        this.departmentsListFull.add(new Department(OMZK));
        this.departmentsListFull.add(new Department(OMZK_new));
        this.departmentsListFull.add(new Department(OMO));
        this.departmentsListFull.add(new Department(MMZ02));
    }

    public void run(int param, boolean showResult, boolean win95colors) {
        this.win95colors = win95colors;

        if (allFiles != null) {
            if (param == 4) {
                readTables(allFiles);
            }

            if (param == 5) {
                readMainWithDSE(allFiles);
            }
        }

        if (all != null) {
            readMain(all);
        }

        if (sbyt != null) {
            readSbyt(sbyt);
        }

        if (showResult) {
            if (all != null) {
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

            if (sbyt != null) {
                for (OrderSbyt orderSbyt : ordersSbyt) {
                    System.out.println("=====================");
                    System.out.println(orderSbyt.getName());
                    System.out.println("=====================");
                    System.out.println("--------------------------------");
                    System.out.println("Непереданных поцизий: " + orderSbyt.getRedCells());
                    System.out.println("--------------------------------");
                    System.out.println("\n");
                }
            }

            if (allFiles != null) {
                if (param == 4) {
                    directories.sort(Comparator.comparing(File::getName));

                    //TODO fix non-dates period show
//                String firstDirectoryName = directories.get(0).getName();
//                String lastDirectoryName = directories.get(directories.size() - 1).getName();
//                String leafName = firstDirectoryName + "-" + lastDirectoryName
//                System.out.println("\nСформирован отчет за период: " + leafName);

                    for (Day day : days) {
                        System.out.println("\n");
                        System.out.println("=====================");
                        System.out.println("Имя файла: " + day.getFileName());
                        System.out.println("День номер: " + day.getDayNumber());
                        System.out.println("=====================");
                        System.out.println("--------------------------------");
                        System.out.println("ТМЦ: " + day.getTmcCount());
                        System.out.println("ДСЕ: " + day.getDseCount());
                        System.out.println("--------------------------------");
                    }
                }

                if (param == 5) {
                    //print overall result and write to result.txt file
                    StringBuilder sb = new StringBuilder();
                    for (Department department : departmentsListFull) {

                        if (department.getDseRepeatCountMap().size() > 0) {
                            sb.append("\n \n");
                            sb.append("\n-------------------------------------------");
                            System.out.println("-------------------------------------------");
                            sb.append("\n").append(department.getName());
                            System.out.println(department.getName());
                            sb.append("\n-------------------------------------------");
                            System.out.println("-------------------------------------------");

                            Map<String, Integer> sortedMap = department.getDseRepeatCountMap().entrySet().stream()
                                    .sorted(Map.Entry.comparingByValue(Comparator.reverseOrder()))
                                    .collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue, (e1, e2) -> e1, LinkedHashMap::new));

                            for (Map.Entry<String, Integer> entry : sortedMap.entrySet()) {
                                sb.append("\n").append(entry.getKey()).append(" [Кол-во повторений: ").append(entry.getValue()).append("]");
                                System.out.println(entry.getKey() + " [Кол-во повторений: " + entry.getValue() + "]");
                            }
                        }
                    }

                    try {
                        //write to text file
                        File textResult = new File("C:\\textResult.doc");
                        BufferedWriter bw = new BufferedWriter(new FileWriter(textResult));
                        bw.write(sb.toString());
                        bw.close();
                        System.out.println("\n\nРезультат подсчета записан в текстовый файл по пути: C:\\textResult.doc");

                        //write to xlsx file
                        File xlsxResult = new File("C:\\xlsxResult.xlsx");
                        writeRepeatCountResultIntoExcel(xlsxResult);
                        System.out.println("Результат подсчета записан в книгу Excel по пути: C:\\xlsxResult.xlsx");

                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
            }
        }

        if (table != null) {
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
        if (gui != null) {
            gui.progressBar.setValue(100);
        }

        if (advancedGUI != null) {
            advancedGUI.progressBar.setValue(100);
        }
    }

    public void writeRepeatCountResultIntoExcel(File file) throws FileNotFoundException, IOException {
        int rowCount = 0;

        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("result");

        for (Department department : departmentsListFull) {
            Row row = sheet.createRow(rowCount);
            Cell departmentNameCell = row.createCell(0);
            departmentNameCell.setCellValue(department.getName());

            Map<String, Integer> sortedMap = department.getDseRepeatCountMap().entrySet().stream()
                    .sorted(Map.Entry.comparingByValue(Comparator.reverseOrder()))
                    .collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue, (e1, e2) -> e1, LinkedHashMap::new));

            for (Map.Entry<String, Integer> entry : sortedMap.entrySet()) {
                rowCount++;

                Cell nomenclatureCell = sheet.createRow(rowCount).createCell(0);
                Cell repeatCountCell = sheet.getRow(rowCount).createCell(1);

                nomenclatureCell.setCellValue(entry.getKey());
                repeatCountCell.setCellValue(entry.getValue());
            }

        }

        // Меняем размер столбца
        sheet.autoSizeColumn(1);

        // Записываем всё в файл
        workbook.write(new FileOutputStream(file));
        workbook.close();
    }

    private void readMainWithDSE(List<File> allFiles) {
        for (File file : allFiles) {
            try (XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(file))) {
                XSSFSheet sheet = workbook.getSheetAt(0);

                for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
                    Row row = sheet.getRow(i);

                    for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
                        Cell cell = row.getCell(j);

                        XSSFCellStyle cs = (XSSFCellStyle) cell.getCellStyle();
                        XSSFFont font = cs.getFont();

                        if (font.getBold() && font.getFontHeightInPoints() == 8 && cell.getCellType() == CellType.STRING) {
                            String departmentName = cell.getStringCellValue();
                            readDSE(i + 1, sheet, departmentName);
                        }
                    }
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private void readDSE(int rowStart, XSSFSheet sheet, String departmentName) {
        for (int i = rowStart; i < sheet.getPhysicalNumberOfRows(); i++) {

            Cell dseVendorCell = sheet.getRow(i).getCell(0);
            Cell dseNomenclatureCell = sheet.getRow(i).getCell(3);

            XSSFCellStyle cs = (XSSFCellStyle) dseVendorCell.getCellStyle();
            XSSFFont font = cs.getFont();

            if (font.getBold()) {
                return;
            }

            DataFormatter formatter = new DataFormatter();

            String vendorValue = formatter.formatCellValue(dseVendorCell);
            String nomenclatureValue = formatter.formatCellValue(dseNomenclatureCell);

            for (Department department : departmentsListFull) {
                if (department.getName().equals(departmentName)) {
                    Map<String, Integer> dseRepeatCountMap = department.getDseRepeatCountMap();
                    if (dseRepeatCountMap.containsKey(nomenclatureValue)) {
                        dseRepeatCountMap.put(nomenclatureValue, dseRepeatCountMap.get(nomenclatureValue) + 1);
                    } else dseRepeatCountMap.put(nomenclatureValue, 1);
                }
            }
        }
    }

    private void readTables(List<File> allFiles) {
        int dayCount = 1;

        int cellRowIndex = 0;
        int cellColumnIndex = 0;

        for (File allFile : allFiles) {
            try (XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(allFile))) {
                XSSFSheet sheet = workbook.getSheet("Сводная таблица");

                for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
                    Row row = sheet.getRow(i);

                    for (Cell cell : row) {
                        XSSFCellStyle cs = (XSSFCellStyle) cell.getCellStyle();
                        XSSFFont font = cs.getFont();

                        if (cell.getCellType() == CellType.STRING
                                && font.getFontHeightInPoints() == 7
                                && cell.getStringCellValue().contains("Кол-во красных строк")) {
                            cellRowIndex = cell.getRowIndex();
                            cellColumnIndex = cell.getColumnIndex();
                        }
                    }
                }

                if (cellRowIndex != 0 && cellColumnIndex != 0) {
                    int tmcCount = (int) sheet.getRow(cellRowIndex + 2).getCell(cellColumnIndex).getNumericCellValue();
                    int dseCount = (int) sheet.getRow(cellRowIndex + 3).getCell(cellColumnIndex).getNumericCellValue();

                    Day day = new Day(String.valueOf(dayCount));
                    day.setFileName(allFile.getName());
                    day.setTmcCount(tmcCount);
                    day.setDseCount(dseCount);
                    days.add(day);
                }
            } catch (IOException e) {
                e.printStackTrace();
            } catch (IllegalStateException e) {
                System.out.println(allFile.getName() + " является файлом старого типа и не содержит необходимых данных.");
            }
            dayCount++;
        }
    }

    private void readMain(File all) {
        try (XSSFWorkbook allWorkBook = new XSSFWorkbook(new FileInputStream(all))) {
            XSSFSheet allSheet = allWorkBook.getSheetAt(0);

            Row row = allSheet.getRow(0);
            for (Cell cell : row) {
                //если это обычный номер заказа
                if (cell.getCellType() == CellType.STRING) {
                    XSSFCellStyle cs = (XSSFCellStyle) cell.getCellStyle();
                    if (cs.getAlignment() == HorizontalAlignment.CENTER) {
                        String orderNumber = cell.getStringCellValue();
                        Order order = new Order(orderNumber);
                        readColumn(cell.getColumnIndex(), allSheet, order);
                        orders.add(order);
                    }
                    //если это ебучий ЗИП
                } else if (cell.getCellType() == CellType.NUMERIC
                        && (cell.getNumericCellValue() == 765
                        || cell.getNumericCellValue() == 7654
                        || cell.getNumericCellValue() == 7655
                        || cell.getNumericCellValue() == 75300
                        || cell.getNumericCellValue() == 75310
                        || cell.getNumericCellValue() == 75311)) {
                    XSSFCellStyle cs = (XSSFCellStyle) cell.getCellStyle();
                    if (cs.getAlignment() == HorizontalAlignment.CENTER) {
                        DataFormatter formatter = new DataFormatter();
                        String orderNumber = formatter.formatCellValue(cell);

                        //todo ебучий ноль вываливается, ебучая запятая добавляется, лютый говнокод
                        StringBuilder sb = new StringBuilder(orderNumber);
                        if (orderNumber.contains("765")) {
                            sb.insert(orderNumber.indexOf('-') + 1, "0");
                        } else if (orderNumber.contains("753")) {
                            sb.insert(orderNumber.indexOf('П') + 1, "0");
                        }

                        orderNumber = sb.toString().replaceAll(",", "");
                        for (Order o : orders) {
                            if (o.getName().equals(orderNumber)) {
                                sb = new StringBuilder(orderNumber);
                                if (cell.getNumericCellValue() == 765) {
                                    sb.insert(11, "0");
                                } else {
                                    sb.insert(12, "0");
                                }
                            }
                        }
                        orderNumber = sb.toString().replaceAll(",", "");

                        Order order = new Order(orderNumber);
                        readColumn(cell.getColumnIndex(), allSheet, order);
                        orders.add(order);
                    }
                }
            }
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

    private void readSbyt(File sbyt) {
        try (XSSFWorkbook sbytWorkBook = new XSSFWorkbook(new FileInputStream(sbyt))) {
            XSSFSheet sbytSheet = sbytWorkBook.getSheetAt(0);

            for (int i = 3; i < sbytSheet.getPhysicalNumberOfRows(); i++) {
                Row row = sbytSheet.getRow(i);
                Cell cell = row.getCell(0);

                if (cell.getCellType() == CellType.STRING) {
                    XSSFCellStyle cs = (XSSFCellStyle) cell.getCellStyle();

                    if (cs.getFillForegroundColorColor() == null) {
                        continue;
                    }

                    String blueARGBHex = "FFC6E2FF";

                    if (cs.getFillForegroundColorColor().getARGBHex().equals(blueARGBHex)) {
                        String orderNumber = cell.getStringCellValue().substring(0, cell.getStringCellValue().lastIndexOf(',')).trim();
                        OrderSbyt orderSbyt = new OrderSbyt(orderNumber);
                        countSbyt(i + 1, sbytSheet, orderSbyt);
                        ordersSbyt.add(orderSbyt);
                    }
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
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

            String redARGBHEX;
            String yellowARGBHEX;

            if (win95colors) {
                redARGBHEX = "00DD9CB3";
                yellowARGBHEX = "00FFFFC0";
            } else {
                redARGBHEX = "FFFFC0CB";
                yellowARGBHEX = "FFFFEC8B";
            }

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

    private void countSbyt(int rowNum, XSSFSheet sheet, OrderSbyt orderSbyt) {
        int cellsCount = 0;

        for (int i = rowNum; i < sheet.getPhysicalNumberOfRows(); i++) {
            Row row = sheet.getRow(i);
            Cell cell = row.getCell(0);

            XSSFCellStyle cs = (XSSFCellStyle) cell.getCellStyle();

            String blueARGBHEX = "FFC6E2FF";
            String lightBlueARGBHEX = "FFDCF1FF";
            String darkBlueARGBHEX = "FF4A62B9";

            if (cell.getCellType() != CellType.BLANK) {
                if (cs.getFillForegroundColorColor() == null) {
                    continue;
                }

                if (cs.getFillForegroundColorColor().getARGBHex().equals(lightBlueARGBHEX)) {
                    cellsCount++;
                }

                if (cs.getFillForegroundColorColor().getARGBHex().equals(blueARGBHEX)
                        || cs.getFillForegroundColorColor().getARGBHex().equals(darkBlueARGBHEX)) {
                    orderSbyt.setRedCells(cellsCount);
                    return;
                }
            }
        }
    }

    private void writeResultTo765Table(File table) {
        try (XSSFWorkbook tableWorkbook = new XSSFWorkbook(new FileInputStream(table))) {
            XSSFSheet tableSheet = tableWorkbook.getSheet("Сводная таблица");

            writeDateOfNow(tableSheet);

            for (int i = 4; i < tableSheet.getPhysicalNumberOfRows(); i++) {
                Row row = tableSheet.getRow(i);
                Cell orderCell = row.getCell(1);

                for (Order order : orders) {
                    //todo fix this bad solution
                    //null row NEP protection
                    try {
                        orderCell.getCellType();
                    } catch (NullPointerException e) {
                        break;
                    }

                    if (orderCell.getCellType() == CellType.STRING
                            && orderCell.getStringCellValue().trim().equals(order.getName().trim())) {
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
                                case OMZK:
                                case OMZK_new: {
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
                }
            }
            recalculateAndWrite(tableWorkbook);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void writeResultTo753Table(File table) {
        try (XSSFWorkbook tableWorkbook = new XSSFWorkbook(new FileInputStream(table))) {
            XSSFSheet tableSheet = tableWorkbook.getSheet("Сводная таблица");

            writeDateOfNow(tableSheet);

            for (int i = 4; i < tableSheet.getPhysicalNumberOfRows(); i++) {
                Row row = tableSheet.getRow(i);
                Cell orderCell = row.getCell(1);

                for (Order order : orders) {
                    //todo fix this bad solution
                    //null row NEP protection
                    try {
                        orderCell.getCellType();
                    } catch (NullPointerException e) {
                        break;
                    }

                    if (orderCell.getCellType() == CellType.STRING
                            && orderCell.getStringCellValue().trim().equals(order.getName().trim())) {
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
                                case OMZK:
                                case OMZK_new: {
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
                }
            }
            recalculateAndWrite(tableWorkbook);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void writeResultToOrdersTable(File table) {
        try (XSSFWorkbook tableWorkbook = new XSSFWorkbook(new FileInputStream(table))) {
            XSSFSheet tableSheet = tableWorkbook.getSheet("Сводная таблица");

            writeDateOfNow(tableSheet);

            for (int i = 3; i < tableSheet.getPhysicalNumberOfRows(); i++) {
                Row row = tableSheet.getRow(i);

                Cell orderCell = row.getCell(2);

                for (Order order : orders) {
                    //todo fix this bad solution
                    //null row NEP protection
                    try {
                        orderCell.getCellType();
                    } catch (NullPointerException e) {
                        break;
                    }
                    if (orderCell.getCellType() == CellType.STRING
                            && (order.getName().trim().equals(orderCell.getStringCellValue().trim()))) {
                        for (Department department : order.getDepartments()) {
                            switch (department.getName()) {
                                case UVK: {
                                    row.getCell(7).setCellValue(department.getYellowCellsCount());
                                    row.getCell(4).setCellValue(department.getRedCellsCount());
                                    break;
                                }
                                case OMO: {
                                    row.getCell(8).setCellValue(department.getYellowCellsCount());
                                    row.getCell(5).setCellValue(department.getRedCellsCount());
                                    break;
                                }
                                case OMZK:
                                case OMZK_new: {
                                    row.getCell(9).setCellValue(department.getYellowCellsCount());
                                    row.getCell(6).setCellValue(department.getRedCellsCount());
                                    break;
                                }
                                case TSEH_5: {
                                    row.getCell(28).setCellValue(department.getYellowCellsCount());
                                    row.getCell(20).setCellValue(department.getRedCellsCount());
                                    break;
                                }
                                case TSEH_10: {
                                    row.getCell(29).setCellValue(department.getYellowCellsCount());
                                    row.getCell(21).setCellValue(department.getRedCellsCount());
                                    break;
                                }
                                case TSEH_51: {
                                    row.getCell(30).setCellValue(department.getYellowCellsCount());
                                    row.getCell(22).setCellValue(department.getRedCellsCount());
                                    break;
                                }
                                case TSEH_121: {
                                    row.getCell(31).setCellValue(department.getYellowCellsCount());
                                    row.getCell(23).setCellValue(department.getRedCellsCount());
                                    break;
                                }
                                case TSEH_217: {
                                    row.getCell(32).setCellValue(department.getYellowCellsCount());
                                    row.getCell(24).setCellValue(department.getRedCellsCount());
                                    break;
                                }
                                case TSEH_317: {
                                    row.getCell(33).setCellValue(department.getYellowCellsCount());
                                    row.getCell(25).setCellValue(department.getRedCellsCount());
                                    break;
                                }
                                case TSEH_416: {
                                    row.getCell(34).setCellValue(department.getYellowCellsCount());
                                    row.getCell(26).setCellValue(department.getRedCellsCount());
                                    break;
                                }

                                case TSEH_517: {
                                    row.getCell(35).setCellValue(department.getYellowCellsCount());
                                    row.getCell(27).setCellValue(department.getRedCellsCount());
                                    break;
                                }
                            }
                        }
                    }
                }

                for (OrderSbyt orderSbyt : ordersSbyt) {
                    //todo fix this bad solution
                    //null row NEP protection
                    try {
                        orderCell.getCellType();
                    } catch (NullPointerException e) {
                        break;
                    }
                    if (orderCell.getCellType() == CellType.STRING
                            && orderCell.getStringCellValue().trim().equals(orderSbyt.getName().trim())) {
                        row.getCell(36).setCellValue(orderSbyt.getRedCells());
                    }
                }
            }
            recalculateAndWrite(tableWorkbook);
        } catch (
                IOException e) {
            e.printStackTrace();
        }
    }

    private void writeResultToParetoTable(File table) {
        //TODO write realization
    }

    private void writeDateOfNow(XSSFSheet tableSheet) {
        Cell dateCell = tableSheet.getRow(0).getCell(0);
        LocalDate dateTime = LocalDate.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd.MM.yyyy");
        dateCell.setCellValue(dateTime.format(formatter));
    }

    private void recalculateFormulas(XSSFWorkbook workbook) {
        FormulaEvaluator evaluator = workbook
                .getCreationHelper()
                .createFormulaEvaluator();
        for (Sheet sheet : workbook) {
            for (Row r : sheet) {
                for (Cell c : r) {
                    if (c.getCellType() == CellType.FORMULA) {
                        evaluator.evaluateFormulaCell(c);
                    }
                }
            }
        }
    }

    private void recalculateAndWrite(XSSFWorkbook workbook) {
        recalculateFormulas(workbook);

        try (FileOutputStream tableFileOutputStream = new FileOutputStream(table)) {
            workbook.write(tableFileOutputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}