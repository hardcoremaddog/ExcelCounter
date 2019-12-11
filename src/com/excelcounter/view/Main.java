package com.excelcounter.view;

import javax.swing.*;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;

public class Main {
    public static final String VERSION = "v0.4";

    public static void main(String[] args) {
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (ClassNotFoundException | UnsupportedLookAndFeelException | IllegalAccessException | InstantiationException e) {
            e.printStackTrace();
        }

        GUI app = new GUI();
        app.setVisible(true);

        System.out.println("\"EasyXLCounter " + VERSION + "\"" +
                "\nby Alexey Zheludov" +
                "\nCopyright (c) 2019 MDDG Software, All rights reserved.\n");

        System.out.println("В книге с данными, на первом листе необходимо иметь:" +
                "\n -Первый столбец: страна, вид, тип, субъект и т.д." +
                "\n -Второй столбец: объем в КГ");
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