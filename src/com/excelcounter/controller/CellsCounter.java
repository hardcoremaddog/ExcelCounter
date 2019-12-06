package com.excelcounter.controller;

import com.excelcounter.model.*;
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

    private List<StuntRow> stuntRowsList = new ArrayList<>();

    private Set<String> countriesUniqueSet = new TreeSet<>();
    private Set<String> productUniqueSet = new TreeSet<>();
    private Set<String> viewOfProductUniqueSet = new TreeSet<>();

    private File all;

    public CellsCounter(File all, GUI gui) {
        this.all = all;
        this.gui = gui;
    }

    public void finalWorkWithData(int param, boolean showResult, boolean win95colors) {
        //todo Stuntman func
        if (all != null && param == 7) {
            readStuntFile(all);
            writeStuntFile(all);

            gui.progressBar.setValue(100);
            gui.allFilePathLabel.setText("                                             (ﾉ◕ヮ◕)ﾉ*:･ﾟ✧ ХОБА!");
        }

        //all tasks done
        if (gui != null) {
            gui.progressBar.setValue(100);
        }
    }

    private void readStuntFile(File file) {
        try (XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(file))) {

            XSSFSheet sheet = workbook.getSheetAt(0);
            for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
                Row row = sheet.getRow(i);

                if (row.getCell(0) == null || row.getCell(0).getCellType() == CellType.STRING) {
                    continue;
                }

                Cell numberPPCell = row.getCell(0);
                Cell yearCell = row.getCell(1);
                Cell countryCell = row.getCell(2);
                Cell productCell = row.getCell(6);
                Cell viewOfProductCell = row.getCell(7);
                Cell cargoTotalWeightCell = row.getCell(15);

                StuntRow stuntRow = new StuntRow((int) numberPPCell.getNumericCellValue());
                stuntRow.setYear((int) yearCell.getNumericCellValue());

                stuntRow.setCountry(countryCell.getStringCellValue());
                countriesUniqueSet.add(countryCell.getStringCellValue());

                stuntRow.setProduct(productCell.getStringCellValue());
                productUniqueSet.add(productCell.getStringCellValue());

                stuntRow.setViewOfProduct(viewOfProductCell.getStringCellValue());
                viewOfProductUniqueSet.add(viewOfProductCell.getStringCellValue());

                stuntRow.setCargoTotalWeight((int) cargoTotalWeightCell.getNumericCellValue());

                stuntRowsList.add(stuntRow);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void writeStuntFile(File all) {
        try (XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(all))) {

            workbook.setSheetName(workbook.getSheetIndex(workbook.getSheetAt(0)), "Ввоз");

            XSSFSheet sheetCountries;
            if (workbook.getSheet("По странам") != null) {
                sheetCountries = workbook.getSheet("По странам");
            } else {
                sheetCountries = workbook.createSheet("По странам");
            }
            sheetCountries.setColumnWidth(1, 10000);
            sheetCountries.setColumnWidth(2, 5000);

            XSSFSheet sheetProducts;
            if (workbook.getSheet("По продукции") != null) {
                sheetProducts = workbook.getSheet("По продукции");
            } else {
                sheetProducts = workbook.createSheet("По продукции");
            }
            sheetProducts.setColumnWidth(1, 38000);
            sheetProducts.setColumnWidth(2, 5000);

            XSSFSheet sheetViewOfProducts;
            if (workbook.getSheet("По виду") != null) {
                sheetViewOfProducts = workbook.getSheet("По виду");
            } else {
                sheetViewOfProducts = workbook.createSheet("По виду");
            }
            sheetViewOfProducts.setColumnWidth(1, 30000);
            sheetViewOfProducts.setColumnWidth(2, 5000);

            //by countries
            int rowIndex = 2;
            int nameCellIndex = 1;
            int totalWeighIndex = 2;
            for (String country : countriesUniqueSet) {

                Row row = sheetCountries.createRow(rowIndex);
                Cell countryNameCell = row.createCell(nameCellIndex);
                Cell totalWeighCell = row.createCell(totalWeighIndex);

                double countryTotalWeight = 0;
                for (StuntRow stuntRow : stuntRowsList) {
                    if (stuntRow.getCountry().equals(country)) {
                        countryTotalWeight += stuntRow.getCargoTotalWeight();
                    }
                }
                if (countryTotalWeight > 0) {
                    countryNameCell.setCellValue(country);
                    totalWeighCell.setCellValue(countryTotalWeight / 1000);
                }
                rowIndex++;
            }

            //by product
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

            //by viewOfProduct
            rowIndex = 2;
            for (String viewOfProduct : viewOfProductUniqueSet) {

                Row row = sheetViewOfProducts.createRow(rowIndex);
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

            double totalWeightFull = 0;
            for (int i = 0; i < sheetCountries.getPhysicalNumberOfRows() + 2; i++) {
                Row row = sheetCountries.getRow(i);

                try {
                    totalWeightFull += row.getCell(2).getNumericCellValue();
                } catch (Exception e) {
                    //it's ok
                }
            }

            //by country
            Row row = sheetCountries.getRow(2);

            Cell nameCell = row.getCell(1);
            Cell totalCell = row.getCell(2);

            nameCell.setCellValue("Страны");
            totalCell.setCellValue(totalWeightFull);

            CellStyle style = workbook.createCellStyle();
            Font font = workbook.createFont();
            font.setBold(true);
            font.setFontHeightInPoints((short) 14);
            style.setFont(font);

            nameCell.setCellStyle(style);
            totalCell.setCellStyle(style);

            //by product
            row = sheetProducts.getRow(2);
            nameCell = row.getCell(1);
            totalCell = row.getCell(2);

            nameCell.setCellValue("Продукт");
            totalCell.setCellValue(totalWeightFull);

            nameCell.setCellStyle(style);
            totalCell.setCellStyle(style);

            //by viewOfProduct
            row = sheetViewOfProducts.getRow(2);
            nameCell = row.getCell(1);
            totalCell = row.getCell(2);

            nameCell.setCellValue("Вид продукта");
            totalCell.setCellValue(totalWeightFull);

            nameCell.setCellStyle(style);
            totalCell.setCellStyle(style);

//            //todo count percent
//            for (int i = 1; i < workbook.getNumberOfSheets(); i++) {
//                Sheet sheet = workbook.getSheetAt(i);
//                sheet.getRow(2).createCell(3).setCellValue("%");
//                sheet.getRow(2).getCell(3).setCellStyle(style);
//                for (int j = 3; j < sheet.getPhysicalNumberOfRows() + 2; j++) {
//                    Row row1 = sheet.getRow(j);
//
//                    Cell percentCell = row1.createCell(3);
//                    percentCell.setCellValue(row1.getCell(2).getNumericCellValue() / totalWeightFull);
//
//                    CellStyle percentStyle = workbook.createCellStyle();
//                    percentStyle.setDataFormat(workbook.createDataFormat().getFormat("0.00%"));
//                    percentCell.setCellStyle(percentStyle);
//                }
//            }

            try (FileOutputStream tableFileOutputStream = new FileOutputStream(all)) {
                workbook.write(tableFileOutputStream);
            } catch (IOException e) {
                System.out.println("Данные посчитаны, но не были записаны, " +
                        "\nт.к. выбранный файл поврежден, либо открыт в системе.");
                System.out.println("Работа программы завершена с ошибкой.");
                return;
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("Работа программы завершена без ошибок.");
    }
}