package com.excelcounter.view;

import com.excelcounter.controller.CellsCounter;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.nio.file.FileVisitResult;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.SimpleFileVisitor;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class AdvancedGUI extends JFrame {

    private AdvancedGUI advancedGUI = this;

    private List<File> directories = new ArrayList<>();
    private List<File> filesForRepeat = new ArrayList<>();
    private List<File> filesForPareto = new ArrayList<>();
    private List<File> filesForReleaseCount = new ArrayList<>();

    private File allNomenclatureTableFile;

    public JProgressBar progressBar = new JProgressBar();

    private JButton mainGUIShowButton = new JButton("Вернуться на основной интерфейс");

    private JButton tablesFileChooserButton = new JButton("Выбрать книги .xlsx с данными");
    private JButton allNomenclatureTableFileChooseButton = new JButton("Полный список номенклатуры/опер. планов");

    private JRadioButton repeatDseRadio = new JRadioButton("Анализ повторений ДСЕ");
    private JRadioButton repeatTmcRadio = new JRadioButton("Анализ повторений ТМЦ");
    private JRadioButton paretoRadio = new JRadioButton("Парето (без записи)");
    private JRadioButton releaseRadio = new JRadioButton("Анализ выпуска ДСЕ");

    private JLabel tablesFilePathLabel = new JLabel();
    private JLabel allNomenclatureTableFilePathLabel = new JLabel();

    private JButton startWork = new JButton("Посчитать и записать");

    AdvancedGUI() {
        super("ExcelCounter " + Main.VERSION);
        this.setBounds(100, 100, 900, 350);
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        Container container = this.getContentPane();
        container.setLayout(new GridLayout(5, 3, 1, 1));

        mainGUIShowButton.addActionListener(new MainGUIShowButton());
        tablesFileChooserButton.addActionListener(new allFileChooseButtonActionListener());
        allNomenclatureTableFileChooseButton.addActionListener(new AllNomenclatureTableFileChooseButtonActionListener());
        startWork.addActionListener(new CountButtonEventListener(this));

        repeatDseRadio.setSelected(true);

        ButtonGroup group = new ButtonGroup();
        group.add(repeatDseRadio);
        group.add(repeatTmcRadio);
        group.add(releaseRadio);
        group.add(paretoRadio);


        progressBar.setStringPainted(true);
        progressBar.setMinimum(0);
        progressBar.setMaximum(100);

        container.add(new JLabel());
        container.add(tablesFileChooserButton);
        container.add(allNomenclatureTableFileChooseButton);

        container.add(new JLabel());
        container.add(tablesFilePathLabel);
        container.add(allNomenclatureTableFilePathLabel);

        container.add(new JLabel());
        container.add(new JLabel());

        container.add(paretoRadio);
        container.add(repeatDseRadio);
        container.add(repeatTmcRadio);
        container.add(releaseRadio);

        container.add(mainGUIShowButton);
        container.add(progressBar);
        container.add(startWork);
    }

    class allFileChooseButtonActionListener implements ActionListener {
        @Override
        public void actionPerformed(ActionEvent e) {
            //clear
            progressBar.setValue(0);
            filesForPareto.clear();
            filesForRepeat.clear();
            tablesFilePathLabel.setText("");

            JFileChooser allFileChooser = new JFileChooser();
            allFileChooser.setMultiSelectionEnabled(true);
            allFileChooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
            setXLSXFilter(allFileChooser);
            int ret = allFileChooser.showDialog(null, "Выбрать файлы книги со сводными таблицами");
            if (ret == JFileChooser.APPROVE_OPTION) {
                directories = Arrays.asList(allFileChooser.getSelectedFiles());
                if (directories.size() > 0) {
                    tablesFilePathLabel.setText("Файлы выбраны");

                    MyFileVisitor myFileVisitor = new MyFileVisitor();

                    if (repeatDseRadio.isSelected()) {
                        myFileVisitor.setPartOfName("Производство");
                    } else if (repeatTmcRadio.isSelected()) {
                        myFileVisitor.setPartOfName("Покупка");
                    } else if (paretoRadio.isSelected()) {
                        myFileVisitor.setPartOfName("Сводная");
                    } else if (releaseRadio.isSelected()) {
                        myFileVisitor.setPartOfName("partForRelease");
                    }

                    for (File file : directories) {
                        try {
                            Files.walkFileTree(file.toPath(), myFileVisitor);
                        } catch (IOException ex) {
                            ex.printStackTrace();
                        }
                    }
                } else {
                    tablesFilePathLabel.setText("Файлы не выбраны!");
                }
            }
        }
    }

    class AllNomenclatureTableFileChooseButtonActionListener implements ActionListener {
        @Override
        public void actionPerformed(ActionEvent e) {
            progressBar.setValue(0);
            JFileChooser tableFileChooser = new JFileChooser();
            setXLSXFilter(tableFileChooser);
            int ret = tableFileChooser.showDialog(null, "Выбрать файл книги с полным списком номенклатуры/опер. планов");
            if (ret == JFileChooser.APPROVE_OPTION) {
                allNomenclatureTableFile = tableFileChooser.getSelectedFile();
                allNomenclatureTableFilePathLabel.setText(allNomenclatureTableFile.getName());
            }
        }
    }

    public class MyFileVisitor extends SimpleFileVisitor<Path> {
        private String partOfName;

        private void setPartOfName(String partOfName) {
            this.partOfName = partOfName;
        }

        @Override
        public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) {
            if (partOfName != null && file.getFileName().toString().contains(partOfName)) {
                if (paretoRadio.isSelected()) {
                    filesForPareto.add(file.toFile());
                } else if (repeatDseRadio.isSelected() || repeatTmcRadio.isSelected()) {
                    filesForRepeat.add(file.toFile());
                }
            }

            if (partOfName != null && partOfName.equalsIgnoreCase("partForRelease")) {
                filesForReleaseCount.add(file.toFile());
            }

            return FileVisitResult.CONTINUE;
        }
    }

    private void setXLSXFilter(JFileChooser fileChooser) {
        FileNameExtensionFilter filter = new FileNameExtensionFilter("XLSX/XLSM files", "xlsx", "xlsm");
        fileChooser.setFileFilter(filter);

    }

    private class MainGUIShowButton implements ActionListener {
        @Override
        public void actionPerformed(ActionEvent e) {
            GUI gui = new GUI();
            gui.setVisible(true);
            advancedGUI.setVisible(false);
        }
    }

    class CountButtonEventListener implements ActionListener {
        private AdvancedGUI advancedGUI;

        private CountButtonEventListener(AdvancedGUI advancedGUI) {
            this.advancedGUI = advancedGUI;
        }

        @Override
        public void actionPerformed(ActionEvent e) {
            String errMsgAll = "Необходимо выбрать по крайней мере один файл книги с данными!";
            String errMsgAllNomenclature = "Необходимо выбрать файл книги c полным списком номенклатуры!";

            CellsCounter cellsCounter;

            if (paretoRadio.isSelected()) {
                if (filesForPareto.size() == 0 || directories.size() == 0) {
                    System.out.println(errMsgAll);
                } else {
                    cellsCounter = new CellsCounter(filesForPareto, directories, allNomenclatureTableFile, advancedGUI);
                    cellsCounter.finalWorkWithData(4, true, false);
                }
            } else if (repeatDseRadio.isSelected() || repeatTmcRadio.isSelected()) {
                if (filesForRepeat.size() == 0 || directories.size() == 0) {
                    System.out.println(errMsgAll);
                } else {
                    if (allNomenclatureTableFile == null) {
                        System.out.println(errMsgAllNomenclature);
                    } else {
                        cellsCounter = new CellsCounter(filesForRepeat, directories, allNomenclatureTableFile, advancedGUI);
                        cellsCounter.finalWorkWithData(5, true, false);
                    }
                }
            } else if (releaseRadio.isSelected()) {
                if (filesForReleaseCount.size() == 0) {
                    System.out.println(errMsgAll);
                } else {
                    cellsCounter = new CellsCounter(filesForReleaseCount, null, allNomenclatureTableFile, advancedGUI);
                    cellsCounter.finalWorkWithData(6, true, false);
                }
            }
        }
    }
}


