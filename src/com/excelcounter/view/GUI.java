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

    private File all;

    private JButton allFileChooserButton = new JButton("Выбрать .xlsx файл с данными");
    public JLabel allFilePathLabel = new JLabel();

    public JCheckBox checkBox = new JCheckBox("Результат в тоннах");

    private JButton startWork = new JButton("ПОСЧИТАТЬ ДАННЫЕ");

    public JProgressBar progressBar = new JProgressBar();

    GUI() {
        super("EasyXLCounter " + Main.VERSION);
        this.setBounds(100, 100, 400, 250);
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        Container container = this.getContentPane();
        container.setLayout(new GridLayout(5, 1, 1, 1));

        container.add(allFileChooserButton);
        container.add(allFilePathLabel);

        allFileChooserButton.addActionListener(new allFileChooseButtonActionListener());

        startWork.addActionListener(new CountButtonEventListener(this));

        progressBar.setStringPainted(true);
        progressBar.setMinimum(0);
        progressBar.setMaximum(100);
        container.add(progressBar);
        container.add(startWork);

        checkBox.setEnabled(true);
        checkBox.setSelected(true);
        container.add(checkBox);
    }

    private void setXLSXFilter(JFileChooser fileChooser) {
        FileNameExtensionFilter filter = new FileNameExtensionFilter("XLS/XLSX/XLSM files", "xls", "xlsx", "xlsm");
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
                allFilePathLabel.setText(" Выбран файл: " + all.getAbsolutePath());
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
            cellsCounter.finalWorkWithData();
        }
    }
}
