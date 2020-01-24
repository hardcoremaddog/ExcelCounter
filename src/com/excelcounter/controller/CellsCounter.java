package com.excelcounter.controller;

import com.excelcounter.model.*;
import com.excelcounter.view.AdvancedGUI;
import com.excelcounter.view.GUI;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class CellsCounter {
    private GUI gui;
    private AdvancedGUI advancedGUI;

    private boolean win95colors;

    private List<File> allFiles;
    private List<File> directories;
    private List<Day> days = new ArrayList<>();
    private List<ExcelFile> excelFiles = new ArrayList<>();

    private File all;
    private File allNomenclatureTableFile;
    private File table;
    private File ostatki;
    private File primenyaemost;
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
    private final String OMZK_new2 = "Отдел межзаводской кооперации*";
    private final String OMO = "Отдел материального обеспечения";
    private final String MMZ02 = "02ММЗ ММЗ Цех 02";

    private final String ELMASH = "Электромашинный участок";
    private final String TSEH_SB_VAG = "Цех сборки вагонов";
    private final String UCH_SB_TEL = "Участок сборки тележек";
    private final String KOMPL_TSEH = "Комплектовочный цех";
    private final String UCH_OKR_VAG = "Участок окраски вагонов";
    private final String UCH_MEH_OBR = "Участок механической обработки и ремонта комплектующих";

    private List<Department> departmentListFull = new ArrayList<>();
    private Map<String, Department> departmentMap = new HashMap<>();
    private Set<String> warningOrders = new TreeSet<>();

    private void initDepartmentsList() {
        departmentListFull.add(new Department(TSEH_5));
        departmentListFull.add(new Department(TSEH_10));
        departmentListFull.add(new Department(TSEH_51));
        departmentListFull.add(new Department(TSEH_121));
        departmentListFull.add(new Department(TSEH_217));
        departmentListFull.add(new Department(TSEH_317));
        departmentListFull.add(new Department(TSEH_416));
        departmentListFull.add(new Department(TSEH_417));
        departmentListFull.add(new Department(TSEH_517));
        departmentListFull.add(new Department(VVMMZ));
        departmentListFull.add(new Department(UVK));
        departmentListFull.add(new Department(OMZK));
        departmentListFull.add(new Department(OMZK_new));
        departmentListFull.add(new Department(OMZK_new2));
        departmentListFull.add(new Department(OMO));
        departmentListFull.add(new Department(MMZ02));
        departmentListFull.add(new Department(ELMASH));
        departmentListFull.add(new Department(TSEH_SB_VAG));
        departmentListFull.add(new Department(UCH_SB_TEL));
        departmentListFull.add(new Department(KOMPL_TSEH));
        departmentListFull.add(new Department(UCH_OKR_VAG));
        departmentListFull.add(new Department(UCH_MEH_OBR));
    }

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

    public CellsCounter(List<File> allFiles, List<File> directories, File allNomenclatureTableFile, AdvancedGUI advancedGUI) {
        initDepartmentsList();
        this.allFiles = allFiles;
        this.directories = directories;
        this.allNomenclatureTableFile = allNomenclatureTableFile;
        this.advancedGUI = advancedGUI;
    }

    public void finalWorkWithData(int param, boolean showResult, boolean win95colors) {
        this.win95colors = win95colors;

        //all table red and yellow cells count
        if (all != null) {
            readMain(all);
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
            }
        }

        //transfer to sbyt control count
        if (sbyt != null) {
            readSbyt(sbyt);

            if (showResult) {
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
            }
        }

        //table type
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

        if (allFiles != null) {
            //paretoTable
            if (param == 4) {
                readTables(allFiles);
                directories.sort(Comparator.comparing(File::getName));
                for (Day day : days) {
                    System.out.println("\n");
                    System.out.println("=====================");
                    System.out.println("Имя файла: " + day.getFileName());
                    System.out.println("=====================");
                    System.out.println("--------------------------------");
                    System.out.println("ТМЦ: " + day.getTmcCount());
                    System.out.println("ДСЕ: " + day.getDseCount());
                    System.out.println("--------------------------------");
                }
            }

            //repeatCount
            if (param == 5) {
                readMainWithDSE(allNomenclatureTableFile);
                readMainWithDSE(allFiles);

                String username = System.getProperty("user.name");

                String filePrefix = "";
                if (allNomenclatureTableFile.getName().contains("765")) {
                    filePrefix = "765";
                } else if (allNomenclatureTableFile.getName().contains("753")) {
                    filePrefix = "753";
                }

                File directory = new File("Z:\\Общая для МВМ\\!ПДУ_ОД\\Результаты подсчета повторений\\" + username);
                File resultFile = new File(directory.toPath() + "\\" + filePrefix + "RepeatCountResult.xlsx");
                try {
                    writeResultIntoRepeatCountTable(directory, resultFile);
                } catch (IOException e) {
                    e.printStackTrace();
                }
                System.out.println("\nГотово!");
                System.out.println("Результат записан в файл по пути: " + resultFile.getAbsolutePath());
            }

            //releaseAnalysis
            if (param == 6) {
                readMainForReleaseAnalysis();

                for (String departmentName : departmentMap.keySet()) {
                    Department department = departmentMap.get(departmentName);

                    System.out.println("\n==================================================================");
                    System.out.println("\n" + department.getName());
                    readOstatki(ostatki, department);

                    for (OperationLvl operationLvl : department.getOperationsLvlList()) {
                        System.out.println(operationLvl.getName());
                        for (Order order : operationLvl.getOrdersList()) {
                            if (order.getPrimenyaemostCount() != order.getFaktReleaseCount()) {
                                System.out.println(order.getName() + "     " + order.getZakazPokypatelya() + " Применяемость: " + order.getPrimenyaemostCount() + " Факт: " + order.getFaktReleaseCount());
                                warningOrders.add(order.getZakazPokypatelya());
                            }
                        }
                    }
                }
                System.out.println("\n==================================================================");
                System.out.println("\nСписок заказов, по которым возникли проблемы:");
                for (String order : warningOrders) {
                    System.out.println(order);
                }
            }
        }

        //all tasks done
        if (gui != null) {
            gui.progressBar.setValue(100);
        }

        if (advancedGUI != null) {
            advancedGUI.progressBar.setValue(100);
        }
    }

    private void readMainForReleaseAnalysis() {
        //init files
        for (File file : allFiles) {
            if (file.getName().contains("Остатки")) {
                ostatki = file;
            } else if (file.getName().contains("Применяемость")) {
                primenyaemost = file;
            }
        }
        readPrimenyaemost(primenyaemost);
    }

    private void readPrimenyaemost(File file) {
        try (XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(file))) {
            XSSFSheet sheet = workbook.getSheetAt(0);

            for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
                Row row = sheet.getRow(i);

                Cell departmentCell = row.getCell(0);
                Cell operationLvlCell = row.getCell(3);

                XSSFCellStyle cs = (XSSFCellStyle) departmentCell.getCellStyle();
                XSSFFont font = cs.getFont();

                final String departmentLineColorARGBHex = "FFC6E2FF";
                final String darkBlueARGBHEX = "FF4A62B9";

                String departmentName;

                if (font.getBold() && font.getFontHeightInPoints() == 8 && departmentCell.getCellType() == CellType.STRING) {
                    if (cs.getFillForegroundColorColor().getARGBHex().equals(departmentLineColorARGBHex)) {

                        departmentName = departmentCell.getStringCellValue().replaceAll("\\*","");
                        Department department = new Department(departmentName);

                        OperationLvl operationLvl = new OperationLvl(operationLvlCell.getStringCellValue());

                        for (int j = departmentCell.getRowIndex() + 1; j < sheet.getPhysicalNumberOfRows(); j++) {
                            Row rowDeep = sheet.getRow(j);
                            Cell orderCell = rowDeep.getCell(0);
                            Cell primenyaemostCountCell = rowDeep.getCell(6);
                            XSSFCellStyle csDeep = (XSSFCellStyle) orderCell.getCellStyle();

                            if (font.getBold() && font.getFontHeightInPoints() == 8 && orderCell.getCellType() == CellType.STRING) {
                                if (csDeep.getFillForegroundColorColor() != null && !csDeep.getFillForegroundColorColor().getARGBHex().equals("FFDCF1FF")) {
                                    if (csDeep.getFillForegroundColorColor().getARGBHex().equals(departmentLineColorARGBHex)
                                            || csDeep.getFillForegroundColorColor().getARGBHex().equals(darkBlueARGBHEX)) {
                                        break;
                                    }
                                    Order order = new Order(orderCell.getStringCellValue());
                                    order.setPrimenyaemostCount(primenyaemostCountCell.getNumericCellValue());
                                    operationLvl.getOrdersList().add(order);
                                }
                            }
                        }

                        department.getOperationsLvlList().add(operationLvl);
                        readOperPlans(department);
                        departmentMap.put(departmentName, department);
                    }
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void readOperPlans(Department department) {
        try (XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(allNomenclatureTableFile))) {
            XSSFSheet sheet = workbook.getSheetAt(0);

            for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
                Row row = sheet.getRow(i);
                Cell orderCell = row.getCell(6);
                Cell zakazPokypatelyaCell = row.getCell(14);

                switch (orderCell.getCellType()) {
                    case STRING: {
                        for (OperationLvl operationLvl : department.getOperationsLvlList()) {
                            for (Order order : operationLvl.getOrdersList()) {
                                if (order.getName().trim().equals(orderCell.getStringCellValue().trim())) {
                                    order.setZakazPokypatelya(zakazPokypatelyaCell.getStringCellValue());
                                }
                            }
                        }
                        break;
                    }
                    case NUMERIC: {
                        for (OperationLvl operationLvl : department.getOperationsLvlList()) {
                            for (Order order : operationLvl.getOrdersList()) {
                                if (order.getName().trim().equals(String.valueOf(orderCell.getNumericCellValue()).trim())) {
                                    order.setZakazPokypatelya(zakazPokypatelyaCell.getStringCellValue());
                                }
                            }
                        }
                        break;
                    }
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void readOstatki(File file, Department department) {
        try (XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(file))) {
            XSSFSheet sheet = workbook.getSheetAt(0);

            for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
                Row row = sheet.getRow(i);

                Cell departmentCell = row.getCell(0);

                XSSFCellStyle cs = (XSSFCellStyle) departmentCell.getCellStyle();
                XSSFFont font = cs.getFont();

                final String departmentLineColorARGBHex = "FFC6E2FF";
                final String violetOperationLvlColorRGBHEx = "FFE6E6FA";

                String departmentName;
                if (font.getBold() && font.getFontHeightInPoints() == 8 && departmentCell.getCellType() == CellType.STRING) {
                    if (cs.getFillForegroundColorColor().getARGBHex().equals(departmentLineColorARGBHex)) {
                        departmentName = departmentCell.getStringCellValue().replaceAll("\\*","");
                        if (department.getName().trim().equals(departmentName.trim())) {
                            for (int j = departmentCell.getRowIndex() + 1; j < sheet.getPhysicalNumberOfRows(); j++) {
                                Row rowDeep = sheet.getRow(j);

                                Cell firstCellDeep = rowDeep.getCell(0);
                                Cell operationLvlCell = rowDeep.getCell(5);

                                XSSFCellStyle csDeep = (XSSFCellStyle) firstCellDeep.getCellStyle();

                                if (csDeep.getFillForegroundColorColor() != null) {
                                    String currentARGBHex = csDeep.getFillForegroundColorColor().getARGBHex();
                                    if (csDeep.getFillForegroundColorColor().getARGBHex().equals(departmentLineColorARGBHex)) {
                                        break;
                                    }

                                    if (currentARGBHex.equals(violetOperationLvlColorRGBHEx)) {
                                        for (OperationLvl operationLvl : department.getOperationsLvlList()) {
                                            if (operationLvl.getName().trim().equals(operationLvlCell.getStringCellValue().trim())) {
                                                for (int k = rowDeep.getRowNum() + 1; k < sheet.getPhysicalNumberOfRows(); k++) {
                                                    Row rowDeepDeep = sheet.getRow(k);
                                                    Cell zakazPokypatelyaDeepDeep = rowDeepDeep.getCell(0);
                                                    XSSFCellStyle csDeepDeep = (XSSFCellStyle) zakazPokypatelyaDeepDeep.getCellStyle();

                                                    if (csDeepDeep.getFillForegroundColorColor() != null) {
                                                        if (csDeepDeep.getFillForegroundColorColor().getARGBHex().equals(violetOperationLvlColorRGBHEx)) {
                                                            break;
                                                        }
                                                    }
                                                    if (zakazPokypatelyaDeepDeep.getCellType() == CellType.STRING && zakazPokypatelyaDeepDeep.getStringCellValue().contains("покупателя")) {
                                                        String currentCellZakazPokypatelya = zakazPokypatelyaDeepDeep.getStringCellValue();

                                                        for (Order order : operationLvl.getOrdersList()) {
                                                            if (order.getZakazPokypatelya() != null && order.getZakazPokypatelya().trim().equals(currentCellZakazPokypatelya.trim())) {
                                                                Cell prihodCell = rowDeepDeep.getCell(10);
                                                                order.setFaktReleaseCount(prihodCell.getNumericCellValue());
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void readMainWithDSE(File file) {
        try (XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(file))) {
            XSSFSheet sheet = workbook.getSheetAt(0);

            for (int i = 0; i < sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);

                for (int j = 0; j < row.getLastCellNum(); j++) {
                    Cell cell = row.getCell(j);

                    XSSFCellStyle cs = (XSSFCellStyle) cell.getCellStyle();
                    XSSFFont font = cs.getFont();

                    if (font.getBold() && font.getFontHeightInPoints() == 8 && cell.getCellType() == CellType.STRING) {
                        String departmentName = cell.getStringCellValue().replaceAll("\\*","");
                        readDSE(i + 1, sheet, departmentName, new ExcelFile(file.getName()), false);
                    }
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void readMainWithDSE(List<File> allFiles) {
        for (File file : allFiles) {

            ExcelFile excelFile = new ExcelFile(file.getName());

            try (XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(file))) {
                XSSFSheet sheet = workbook.getSheetAt(0);

                for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
                    Row row = sheet.getRow(i);

                    for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
                        Cell cell = row.getCell(j);

                        XSSFCellStyle cs = (XSSFCellStyle) cell.getCellStyle();
                        XSSFFont font = cs.getFont();

                        if (font.getBold() && font.getFontHeightInPoints() == 8 && cell.getCellType() == CellType.STRING) {
                            String departmentName = cell.getStringCellValue().replaceAll("\\*", "");
                            excelFile.getDepartmentList().add(new Department(departmentName));
                            readDSE(i + 1, sheet, departmentName, excelFile, true);
                        }
                    }
                }
            } catch (IOException e) {
                e.printStackTrace();
            }

            excelFiles.add(excelFile);
        }
    }

    private void readDSE(int rowStart, XSSFSheet sheet, String departmentName, ExcelFile excelFile, boolean manyFiles) {
        for (int i = rowStart; i < sheet.getPhysicalNumberOfRows(); i++) {

            Cell dseVendorCell = sheet.getRow(i).getCell(0);
            Cell dseNomenclatureCell = sheet.getRow(i).getCell(6);

            XSSFCellStyle cs = (XSSFCellStyle) dseVendorCell.getCellStyle();
            XSSFFont font = cs.getFont();

            if (font.getBold()) {
                return;
            }

            DataFormatter formatter = new DataFormatter();

            String nomenclatureValue = formatter.formatCellValue(dseNomenclatureCell);

            if (manyFiles) {
                for (Department department : excelFile.getDepartmentList()) {
                    if (department.getName().equals(departmentName)) {

                        final String redARGBHEX = "FFFFC0CB";

                        //if this dse red
                        if (cs.getFillForegroundColorColor().getARGBHex().equals(redARGBHEX)) {
                            department.getDseTreeSet().add(nomenclatureValue);
                        }
                    }
                }
            } else {
                for (Department department : departmentListFull) {
                    if (department.getName().equals(departmentName)) {
                        department.getDseTreeSet().add(nomenclatureValue);
                    }
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
                //if it is regular orderNumber
                if (cell.getCellType() == CellType.STRING) {
                    XSSFCellStyle cs = (XSSFCellStyle) cell.getCellStyle();
                    if (cs.getAlignment() == HorizontalAlignment.CENTER) {
                        String orderNumber = cell.getStringCellValue();
                        Order order = new Order(orderNumber);
                        readColumn(cell.getColumnIndex(), allSheet, order);
                        orders.add(order);
                    }
                    //if it ZIP
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
                String departmentName = row.getCell(0).getStringCellValue().replaceAll("\\*","");
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

                    final String blueARGBHex = "FFC6E2FF";
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
                            try {
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
                                    case OMZK_new:
                                    case OMZK_new2: {
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
                                    //TODO add radio button SPB and realise method for it
                                    case ELMASH: {
                                        row.getCell(20).setCellValue(department.getRedCellsCount());
                                        break;
                                    }

                                    case TSEH_SB_VAG: {
                                        row.getCell(21).setCellValue(department.getRedCellsCount());
                                        break;
                                    }

                                    case UCH_SB_TEL: {
                                        row.getCell(22).setCellValue(department.getRedCellsCount());
                                        break;
                                    }

                                    case KOMPL_TSEH: {
                                        row.getCell(23).setCellValue(department.getRedCellsCount());
                                        break;
                                    }

                                    case UCH_MEH_OBR: {
                                        row.getCell(24).setCellValue(department.getRedCellsCount());
                                        break;
                                    }

                                    case UCH_OKR_VAG: {
                                        row.getCell(25).setCellValue(department.getRedCellsCount());
                                        break;
                                    }
                                }
                            } catch (NullPointerException e) {
                                //it's probably SPB, go ahead
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
                                case OMZK_new:
                                case OMZK_new2: {
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
                                case OMZK_new:
                                case OMZK_new2: {
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

    private void writeResultIntoRepeatCountTable(File directory, File file) throws IOException {
        XSSFWorkbook book;
        if (!Files.exists(file.toPath())) {
            Files.createDirectories(directory.toPath());
            Files.createFile(file.toPath());
            book = new XSSFWorkbook();
        } else {
            try (FileInputStream fis = new FileInputStream(file)) {
                book = new XSSFWorkbook(fis);
            }
        }

        String sheetName = directories.get(0).getName()
                + "-" + directories.get(directories.size() - 1).getName();
        Sheet sheet;
        if (book.getSheet(sheetName) != null) {
            sheet = book.getSheet(sheetName);
        } else {
            sheet = book.createSheet(sheetName);
        }

        int excelFileNameColumnNum = 1;

        sheet.setDefaultColumnWidth(7);
        sheet.createRow(0);

        boolean firstCycle = true;

        for (ExcelFile excelFile : excelFiles) {
            Cell excelFileNameCell = sheet.getRow(0).createCell(excelFileNameColumnNum);

            String filename = excelFile.getFileName();
            Pattern pattern = Pattern.compile("^.+(\\d{2}[.]\\d{2}[.]\\d{2}).+$");
            Matcher matcher = pattern.matcher(filename);

            String date;
            if (matcher.find()) {
                date = matcher.group(1);
            } else {
                date = "Не удалось прочитать дату из названия файла";
            }
            excelFileNameCell.setCellValue(date);

            int rowNum = 0;
            for (Department department : departmentListFull) {
                rowNum++;
                Row departmentRow;
                if (firstCycle) {
                    if (department.getDseTreeSet().size() != 0) {
                        departmentRow = sheet.createRow(rowNum);
                        Cell departmentCell = departmentRow.createCell(0);
                        departmentCell.setCellValue(department.getName());
                    }
                }
                rowNum++;
                for (String nomenclature : department.getDseTreeSet()) {
                    Row nomenclatureRow;
                    if (firstCycle) {
                        nomenclatureRow = sheet.createRow(rowNum);
                        Cell nomenclatureCell = nomenclatureRow.createCell(0);
                        nomenclatureCell.setCellValue(nomenclature);
                    } else {
                        nomenclatureRow = sheet.getRow(rowNum);
                    }
                    rowNum++;
                    boolean depMatch = excelFile.getDepartmentList().stream()
                            .anyMatch(d -> d.getName().equals(department.getName()));

                    if (depMatch) {
                        Department fileDepartment = excelFile.getDepartmentList().stream()
                                .filter(d -> d.getName().equals(department.getName()))
                                .findAny()
                                .orElse(null);

                        if (fileDepartment != null) {
                            boolean match = fileDepartment.getDseTreeSet().stream()
                                    .anyMatch(n -> n.equals(nomenclature));

                            Cell valueCell = nomenclatureRow.createCell(excelFileNameColumnNum);

                            if (match) {
                                valueCell.setCellValue(1);
                            } else {
                                valueCell.setCellValue(0);
                            }
                        }
                    }
                }
            }
            excelFileNameColumnNum++;
            firstCycle = false;
        }
        sheet.setRowSumsBelow(false);
        sheet.createFreezePane(0, 1);
        sheet.setColumnWidth(0, 17500);

        table = file;
        recalculateAndWrite(book);
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