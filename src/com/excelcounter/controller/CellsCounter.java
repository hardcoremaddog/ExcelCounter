package com.excelcounter.controller;

import com.excelcounter.model.StuntRow;
import com.excelcounter.view.GUI;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Set;
import java.util.TreeSet;

public class CellsCounter {
    private GUI gui;

    private List<StuntRow> stuntRowsList = new ArrayList<>();

    private Set<String> countriesOriginUniqueSet = new TreeSet<>();
    private Set<String> countriesImportUniqueSet = new TreeSet<>();
    private Set<String> productUniqueSet = new TreeSet<>();
    private Set<String> viewOfProductUniqueSet = new TreeSet<>();

    private File all;

    public CellsCounter(File all, GUI gui) {
        this.all = all;
        this.gui = gui;
    }

    public void finalWorkWithData() {
        if (all != null) {
            try {
                readAndWriteStunt(all);
            } catch (IOException ex) {
                ex.printStackTrace();
            }

            gui.progressBar.setValue(100);
            gui.allFilePathLabel.setText("                                             (ﾉ◕ヮ◕)ﾉ*:･ﾟ✧ ХОБА!");
        }

        if (gui != null) {
            gui.progressBar.setValue(100);
        }
    }

    private void readAndWriteStunt(File file) throws IOException {
        FileInputStream fin = new FileInputStream(file);

        Workbook workbook = WorkbookFactory.create(fin);

        Sheet sheet = workbook.getSheetAt(0);
        for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
            Row row = sheet.getRow(i);

            if (row.getCell(0) == null || row.getCell(0).getCellType() == CellType.STRING) {
                continue;
            }

            StuntRow stuntRow = new StuntRow();

            stuntRow.setCountryOrigin(row.getCell(2).getStringCellValue());
            countriesOriginUniqueSet.add(row.getCell(2).getStringCellValue());

            stuntRow.setCountryImport(row.getCell(4).getStringCellValue());
            countriesImportUniqueSet.add(row.getCell(4).getStringCellValue());

            stuntRow.setProduct(row.getCell(6).getStringCellValue());
            productUniqueSet.add(row.getCell(6).getStringCellValue());

            stuntRow.setViewOfProduct(row.getCell(7).getStringCellValue());
            viewOfProductUniqueSet.add(row.getCell(7).getStringCellValue());

            stuntRow.setCargoTotalWeight(row.getCell(15).getNumericCellValue());
            stuntRowsList.add(stuntRow);
        }

        workbook.setSheetName(workbook.getSheetIndex(workbook.getSheetAt(0)), "Ввоз");

        int rowIndex = 2;
        int nameCellIndex = 1;
        int totalWeighIndex = 2;

        double totalWeightFull = 0;

        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 14);
        style.setFont(font);

        //by countriesOrigin
        countriesOriginUniqueSet.remove("");
        if (countriesOriginUniqueSet.size() > 1) {
            Sheet sheetCountriesOrigin;
            if (workbook.getSheet("По странам происхождения") != null) {
                sheetCountriesOrigin = workbook.getSheet("По странам происхождения");
            } else {
                sheetCountriesOrigin = workbook.createSheet("По странам происхождения");
            }
            sheetCountriesOrigin.setColumnWidth(1, 10000);
            sheetCountriesOrigin.setColumnWidth(2, 5000);

            for (String country : countriesOriginUniqueSet) {
                Row row = sheetCountriesOrigin.createRow(rowIndex);
                Cell countryNameCell = row.createCell(nameCellIndex);
                Cell totalWeighCell = row.createCell(totalWeighIndex);

                double countryTotalWeight = 0;
                for (StuntRow stuntRow : stuntRowsList) {
                    if (stuntRow.getCountryOrigin().equals(country)) {
                        countryTotalWeight += stuntRow.getCargoTotalWeight();
                    }
                }
                if (countryTotalWeight > 0) {
                    countryNameCell.setCellValue(country);
                    totalWeighCell.setCellValue(countryTotalWeight / 1000);
                }
                rowIndex++;
            }

            for (int i = 0; i < sheetCountriesOrigin.getPhysicalNumberOfRows() + 2; i++) {
                Row row = sheetCountriesOrigin.getRow(i);

                try {
                    totalWeightFull += row.getCell(2).getNumericCellValue();
                } catch (Exception e) {
                    //it's ok
                }
            }

            Row row = sheetCountriesOrigin.getRow(2);
            Cell nameCell = row.getCell(1);
            Cell totalCell = row.getCell(2);

            nameCell.setCellValue("Страна происхождения");
            totalCell.setCellValue(totalWeightFull);

            nameCell.setCellStyle(style);
            totalCell.setCellStyle(style);
        } else {
            String country = countriesOriginUniqueSet.iterator().next();
            System.out.println("Страна происхождения только одна, " + country + ". Смысл считать объем по одной стране? \nНичего не пишу короче.");
        }

        //by countriesImport
        countriesImportUniqueSet.remove("");
        if (countriesImportUniqueSet.size() > 1) {
            Sheet sheetCountriesImport;
            if (workbook.getSheet("По странам-имортерам") != null) {
                sheetCountriesImport = workbook.getSheet("По странам-имортерам");
            } else {
                sheetCountriesImport = workbook.createSheet("По странам-имортерам");
            }
            sheetCountriesImport.setColumnWidth(1, 10000);
            sheetCountriesImport.setColumnWidth(2, 5000);

            rowIndex = 2;
            for (String country : countriesImportUniqueSet) {
                Row row = sheetCountriesImport.createRow(rowIndex);
                Cell countryNameCell = row.createCell(nameCellIndex);
                Cell totalWeighCell = row.createCell(totalWeighIndex);

                double productTotalWeight = 0;
                for (StuntRow stuntRow : stuntRowsList) {
                    if (stuntRow.getCountryImport().equals(country)) {
                        productTotalWeight += stuntRow.getCargoTotalWeight();
                    }
                }

                if (productTotalWeight > 0) {
                    countryNameCell.setCellValue(country);
                    totalWeighCell.setCellValue(productTotalWeight / 1000);
                }
                rowIndex++;
            }

            Row row = sheetCountriesImport.getRow(2);
            Cell nameCell = row.getCell(1);
            Cell totalCell = row.getCell(2);

            nameCell.setCellValue("Страна-импортер");
            totalCell.setCellValue(totalWeightFull);

            nameCell.setCellStyle(style);
            totalCell.setCellStyle(style);
        } else {
            String country = countriesImportUniqueSet.iterator().next();
            System.out.println("Страна-импортер только одна, " + country + ". Смысл считать объем по одной стране? \nНичего не пишу короче.");
        }

        //by product
        Sheet sheetProducts;
        if (workbook.getSheet("По продукции") != null) {
            sheetProducts = workbook.getSheet("По продукции");
        } else {
            sheetProducts = workbook.createSheet("По продукции");
        }
        sheetProducts.setColumnWidth(1, 38000);
        sheetProducts.setColumnWidth(2, 5000);

        rowIndex = 2;
        for (String product : productUniqueSet) {

            Row row = sheetProducts.createRow(rowIndex);
            Cell countryNameCell = row.createCell(nameCellIndex);
            Cell totalWeighCell = row.createCell(totalWeighIndex);

            double productTotalWeight = 0;
            for (StuntRow stuntRow : stuntRowsList) {
                if (stuntRow.getProduct().equals(product)) {
                    productTotalWeight += stuntRow.getCargoTotalWeight();
                }
            }

            if (productTotalWeight > 0) {
                countryNameCell.setCellValue(product);
                totalWeighCell.setCellValue(productTotalWeight / 1000);
            }
            rowIndex++;
        }

        Row row = sheetProducts.getRow(2);
        Cell nameCell = row.getCell(1);
        Cell totalCell = row.getCell(2);

        nameCell.setCellValue("Продукт");
        totalCell.setCellValue(totalWeightFull);

        nameCell.setCellStyle(style);
        totalCell.setCellStyle(style);

        //by viewOfProduct
        Sheet sheetViewOfProducts;
        if (workbook.getSheet("По виду") != null) {
            sheetViewOfProducts = workbook.getSheet("По виду");
        } else {
            sheetViewOfProducts = workbook.createSheet("По виду");
        }
        sheetViewOfProducts.setColumnWidth(1, 30000);
        sheetViewOfProducts.setColumnWidth(2, 5000);

        rowIndex = 2;
        for (String viewOfProduct : viewOfProductUniqueSet) {
            row = sheetViewOfProducts.createRow(rowIndex);
            Cell countryNameCell = row.createCell(nameCellIndex);
            Cell totalWeighCell = row.createCell(totalWeighIndex);
            double viewTotalWeight = 0;
            for (StuntRow stuntRow : stuntRowsList) {
                if (stuntRow.getViewOfProduct().equals(viewOfProduct)) {
                    viewTotalWeight += stuntRow.getCargoTotalWeight();
                }
            }
            if (viewTotalWeight > 0) {
                countryNameCell.setCellValue(viewOfProduct);
                totalWeighCell.setCellValue(viewTotalWeight / 1000);
            }
            rowIndex++;
        }
        row = sheetViewOfProducts.getRow(2);
        nameCell = row.getCell(1);
        totalCell = row.getCell(2);

        nameCell.setCellValue("Вид продукта");
        totalCell.setCellValue(totalWeightFull);

        nameCell.setCellStyle(style);
        totalCell.setCellStyle(style);

        //todo count percent
        for (int i = 1; i < workbook.getNumberOfSheets(); i++) {
            sheet = workbook.getSheetAt(i);
            sheet.getRow(2).createCell(3).setCellValue("%");
            sheet.getRow(2).getCell(3).setCellStyle(style);
            for (int j = 3; j < sheet.getPhysicalNumberOfRows() + 2; j++) {
                Row row1 = sheet.getRow(j);

                Cell percentCell = row1.createCell(3);
                percentCell.setCellValue(row1.getCell(2).getNumericCellValue() / totalWeightFull);

                CellStyle percentStyle = workbook.createCellStyle();
                percentStyle.setDataFormat(workbook.createDataFormat().getFormat("0.00%"));
                percentCell.setCellStyle(percentStyle);
            }
        }
        FileOutputStream fos = new FileOutputStream(file);

        workbook.write(fos);
        fos.close();
        fin.close();

        System.out.println("Работа программы завершена без ошибок.");
    }
}