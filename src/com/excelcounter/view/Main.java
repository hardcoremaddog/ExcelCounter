package com.excelcounter.view;

import javax.swing.*;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;

public class Main {
    public static final String VERSION = "v0.3a";

    public static void main(String[] args) {
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (ClassNotFoundException | UnsupportedLookAndFeelException | IllegalAccessException | InstantiationException e) {
            e.printStackTrace();
        }

        GUI app = new GUI();
        app.setVisible(true);

        System.out.println("\"FriendlyExcelJavaWorker " + VERSION + "\"" +
                "\nby Alexey Zheludov" +
                "\nCopyright (c) 2019 MDDG Software, All rights reserved.\n");
    }

    protected static void checkFile(File file) throws IOException {
        if (!file.exists() && !file.isFile()) {
            throw new FileNotFoundException("Неверно указан путь до книги excel с данными. Книги по этому пути не существует!");
        }
        if (!file.getName().endsWith(".xls") && !file.getName().endsWith(".xlsx") && !file.getName().endsWith(".xlsm")) {
            throw new UnsupportedOperationException("Файл не является книгой excel формата .xls, .xlsx или .xlsm");
        }
    }
}