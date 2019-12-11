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

    private Set<String> namesUniqueList = new TreeSet<>();


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
            gui.allFilePathLabel.setText("                                  (ﾉ◕ヮ◕)ﾉ*:･ﾟ✧ QUAS, WEX, EXORT!!");
        }
    }

    private void readAndWriteStunt(File file) throws IOException {
        //get data
        FileInputStream fin = new FileInputStream(file);
        Workbook workbook = WorkbookFactory.create(fin);
        Sheet sheet = workbook.getSheetAt(0);
        for (int i = 0; i < sheet.getLastRowNum() + 1; i++) {
            Row row = sheet.getRow(i);

            StuntRow stuntRow = new StuntRow();
            stuntRow.setName(row.getCell(0).getStringCellValue());
            namesUniqueList.add(row.getCell(0).getStringCellValue());
            stuntRow.setCargoTotalWeight(row.getCell(1).getNumericCellValue());

            stuntRowsList.add(stuntRow);
        }
        workbook.setSheetName(workbook.getSheetIndex(workbook.getSheetAt(0)), "Данные");


        //create sheetResult
        double totalWeightFull = 0;
        String sheetResultName = "Результат";
        Sheet sheetResult;
        if (workbook.getSheet(sheetResultName) != null) {
            sheetResult = workbook.getSheet(sheetResultName);
        } else {
            sheetResult = workbook.createSheet(sheetResultName);
        }
        sheetResult.setColumnWidth(0, 20000);
        sheetResult.setColumnWidth(1, 5000);


        //count totalByUniqueName
        int rowIndex = 0;
        for (String name : namesUniqueList) {
            Row row = sheetResult.createRow(rowIndex);
            Cell countryNameCell = row.createCell(0);
            Cell totalWeighCell = row.createCell(1);

            double countryTotalWeight = 0;
            for (StuntRow stuntRow : stuntRowsList) {
                if (stuntRow.getName().equals(name)) {
                    countryTotalWeight += stuntRow.getCargoTotalWeight();
                }
            }
            if (countryTotalWeight > 0) {
                countryNameCell.setCellValue(name);
                if (gui.checkBox.isSelected()) {
                    totalWeighCell.setCellValue(countryTotalWeight / 1000);
                } else {
                    totalWeighCell.setCellValue(countryTotalWeight);
                }
            }
            rowIndex++;
        }

        for (int i = 0; i < namesUniqueList.size(); i++) {
            Row row = sheetResult.getRow(i);

            try {
                totalWeightFull += row.getCell(1).getNumericCellValue();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        Row row = sheetResult.createRow(namesUniqueList.size() + 1);
        Cell totalCell = row.createCell(1);
        totalCell.setCellValue(totalWeightFull);


        //count %
        sheet = workbook.getSheet(sheetResultName);
        for (int j = 0; j < sheet.getLastRowNum(); j++) {
            Row row1 = sheet.getRow(j);
            if (row1 == null) continue;

            Cell totalNameCell = row.createCell(0);
            totalNameCell.setCellValue("Итого:");
            Cell percentCell = row1.createCell(2);
            percentCell.setCellValue(row1.getCell(1).getNumericCellValue() / totalWeightFull);

            CellStyle percentStyle = workbook.createCellStyle();
            percentStyle.setDataFormat(workbook.createDataFormat().getFormat("0.00%"));
            percentCell.setCellStyle(percentStyle);
        }


        //write result and say good bye
        FileOutputStream fos = new FileOutputStream(file);
        workbook.write(fos);
        fos.close();
        fin.close();

        String username = System.getProperty("user.name");

        System.out.println("______________________________________");
        System.out.println("Работа программы завершена без ошибок.");
        System.out.println(username + ", хорошего дня!");
    }
}