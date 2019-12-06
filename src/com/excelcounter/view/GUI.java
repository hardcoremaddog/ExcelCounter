package com.excelcounter.view;

import com.excelcounter.controller.CellsCounter;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;

public class GUI extends JFrame {

    private GUI gui = this;

    private File all;
    private File table;
    private File sbyt;

    private JButton allFileChooserButton = new JButton("Выбрать книгу .xlsx с данными");
    private JButton sbytFileChooseButton = new JButton("Выбрать книгу .xlsx КПНС");
    private JButton tableFileChooseButton = new JButton("Выбрать книгу .xlsx/.xlsm с таблицей");

    private JButton advancedGUIShowButton = new JButton("Перейти на дополнительный интерфейс");

    private JLabel allFilePathLabel = new JLabel();
    private JLabel sbytFilePathLabel = new JLabel();
    private JLabel tableFilePathLabel = new JLabel();

    private JLabel tableTypeLabel = new JLabel("Тип таблицы для записи данных");
    private JRadioButton table765Radio = new JRadioButton("765 и СПб");
    private JRadioButton table753Radio = new JRadioButton("753");
    private JRadioButton tableOrdersRadio = new JRadioButton("Заказы");

    private JRadioButton stuntmanMikeRadio = new JRadioButton("Stuntman Mike");

    private JCheckBox check = new JCheckBox("Вывести результат подсчета в консоль", true);
    private JCheckBox win95colors = new JCheckBox("Win95 cell colors", false);

    private JButton startWork = new JButton("Посчитать и записать");

    public JProgressBar progressBar = new JProgressBar();

    GUI() {
        super("ExcelCounter " + Main.VERSION);
        this.setBounds(100, 100, 900, 350);
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        Container container = this.getContentPane();
        container.setLayout(new GridLayout(5, 4, 1, 1));

        container.add(allFileChooserButton);
        container.add(sbytFileChooseButton);
        container.add(tableFileChooseButton);
        container.add(allFilePathLabel);
        container.add(sbytFilePathLabel);
        container.add(tableFilePathLabel);

        ButtonGroup group = new ButtonGroup();
        group.add(table765Radio);
        group.add(table753Radio);
        group.add(tableOrdersRadio);

        group.add(stuntmanMikeRadio);

        container.add(tableTypeLabel);
        container.add(table765Radio);
        container.add(table753Radio);
        container.add(tableOrdersRadio);

        stuntmanMikeRadio.setSelected(true);
        container.add(stuntmanMikeRadio);

        allFileChooserButton.addActionListener(new allFileChooseButtonActionListener());
        sbytFileChooseButton.addActionListener(new sbytFileChooseButtonActionListener());
        tableFileChooseButton.addActionListener(new TableFileChooseButtonActionListener());
        advancedGUIShowButton.addActionListener(new AdvancedGUIShowButtonActionListener());
        advancedGUIShowButton.setDefaultCapable(true);
        startWork.addActionListener(new CountButtonEventListener(this));
        container.add(check);
//        container.add(win95colors);

        progressBar.setStringPainted(true);
        progressBar.setMinimum(0);
        progressBar.setMaximum(100);
        container.add(advancedGUIShowButton);
        container.add(progressBar);
        container.add(startWork);
    }

    private void setXLSXFilter(JFileChooser fileChooser) {
        FileNameExtensionFilter filter = new FileNameExtensionFilter("XLSX/XLSM files", "xlsx", "xlsm");
        fileChooser.setFileFilter(filter);
    }

    private void checkFile(JFileChooser fileChooser) {
        try {
            Main.checkFile(fileChooser.getSelectedFile());
        } catch (IOException e1) {
            e1.printStackTrace();
        }
    }

    class allFileChooseButtonActionListener implements ActionListener {
        @Override
        public void actionPerformed(ActionEvent e) {
            progressBar.setValue(0);
            JFileChooser allFileChooser = new JFileChooser();
            allFileChooser.setMultiSelectionEnabled(false);
            setXLSXFilter(allFileChooser);
            int ret = allFileChooser.showDialog(null, "Выбрать файл книги с данными");
            if (ret == JFileChooser.APPROVE_OPTION) {
                checkFile(allFileChooser);
                all = allFileChooser.getSelectedFile();
                allFilePathLabel.setText(all.getName());
            }
        }
    }

    class sbytFileChooseButtonActionListener implements ActionListener {
        @Override
        public void actionPerformed(ActionEvent e) {
            progressBar.setValue(0);
            JFileChooser sbytFileChooser = new JFileChooser();
            setXLSXFilter(sbytFileChooser);
            int ret = sbytFileChooser.showDialog(null, "Выбрать файл книги контроля передачи на сбыт");
            if (ret == JFileChooser.APPROVE_OPTION) {
                checkFile(sbytFileChooser);
                sbyt = sbytFileChooser.getSelectedFile();
                sbytFilePathLabel.setText(sbyt.getName());
            }
        }
    }

    class TableFileChooseButtonActionListener implements ActionListener {
        @Override
        public void actionPerformed(ActionEvent e) {
            progressBar.setValue(0);
            JFileChooser tableFileChooser = new JFileChooser();
            setXLSXFilter(tableFileChooser);
            int ret = tableFileChooser.showDialog(null, "Выбрать файл книги с таблицей");
            if (ret == JFileChooser.APPROVE_OPTION) {
                checkFile(tableFileChooser);
                table = tableFileChooser.getSelectedFile();
                tableFilePathLabel.setText(table.getName());
            }
        }
    }

    class CountButtonEventListener implements ActionListener {
        private GUI gui;

        private CountButtonEventListener(GUI gui) {
            this.gui = gui;
        }

        @Override
        public void actionPerformed(ActionEvent e) {
            CellsCounter cellsCounter;

            if (all != null) {
                cellsCounter = new CellsCounter(all, gui);
            } else {
                System.out.println("Файл не выбран!");
                return;
            }

            cellsCounter.finalWorkWithData(7, true, false);
        }
    }

    private class AdvancedGUIShowButtonActionListener implements ActionListener {
        @Override
        public void actionPerformed(ActionEvent e) {
            AdvancedGUI advancedGUI = new AdvancedGUI();
            advancedGUI.setVisible(true);
            gui.setVisible(false);
        }
    }
}
