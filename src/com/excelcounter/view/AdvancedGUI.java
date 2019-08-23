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
	private List<File> files = new ArrayList<>();
	private File table;

	public JProgressBar progressBar = new JProgressBar();

	private JButton mainGUIShowButton = new JButton("Вернуться на основной интерфейс");

	private JButton allFileChooserButton = new JButton("Выбрать книги .xlsx со сводными таблицами");
	private JButton tableFileChooseButton = new JButton("Выбрать книгу .xlsx/.xlsm с таблицей для записи данных");

	private JRadioButton paretoRadio = new JRadioButton("Парето");
	private JRadioButton repeatRadio = new JRadioButton("Анализ повторений");

	private JLabel allFilePathLabel = new JLabel();
	private JLabel tableFilePathLabel = new JLabel();

	private JButton startWork = new JButton("Посчитать и записать");

	AdvancedGUI() {
		super("ExcelCounter" + Main.VERSION);
		this.setBounds(100, 100, 900, 350);
		this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

		Container container = this.getContentPane();
		container.setLayout(new GridLayout(5, 4, 1, 1));

		mainGUIShowButton.addActionListener(new MainGUIShowButton());
		allFileChooserButton.addActionListener(new allFileChooseButtonActionListener());
		tableFileChooseButton.addActionListener(new TableFileChooseButtonActionListener());
		startWork.addActionListener(new CountButtonEventListener(this));

		paretoRadio.setSelected(true);

		ButtonGroup group = new ButtonGroup();
		group.add(paretoRadio);
		group.add(repeatRadio);

		progressBar.setStringPainted(true);
		progressBar.setMinimum(0);
		progressBar.setMaximum(100);
		container.add(mainGUIShowButton);
		container.add(new JLabel());
		container.add(allFileChooserButton);
		container.add(allFilePathLabel);
		container.add(tableFileChooseButton);
		container.add(tableFilePathLabel);
		container.add(paretoRadio);
		container.add(repeatRadio);
		container.add(startWork);

		container.add(progressBar);
	}

	class allFileChooseButtonActionListener implements ActionListener {
		@Override
		public void actionPerformed(ActionEvent e) {
			progressBar.setValue(0);
			files.clear();
			JFileChooser allFileChooser = new JFileChooser();
			allFileChooser.setMultiSelectionEnabled(true);
			allFileChooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
			setXLSXFilter(allFileChooser);
			int ret = allFileChooser.showDialog(null, "Выбрать файлы книги со сводными таблицами");
			if (ret == JFileChooser.APPROVE_OPTION) {
				directories = Arrays.asList(allFileChooser.getSelectedFiles());
				if (directories.size() > 0) {
					allFilePathLabel.setText("Файлы выбраны.");

					MyFileVisitor myFileVisitor = new MyFileVisitor();

					for (File file : directories) {
						try {
							Files.walkFileTree(file.toPath(), myFileVisitor);
						} catch (IOException ex) {
							ex.printStackTrace();
						}
					}
				} else {
					allFilePathLabel.setText("Файлы не выбраны!");
				}
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
				table = tableFileChooser.getSelectedFile();
				tableFilePathLabel.setText(table.getName());
			}
		}
	}

	public class MyFileVisitor extends SimpleFileVisitor<Path> {
		String partOfName = "Сводная";

		@Override
		public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) {
			if (partOfName != null && file.getFileName().toString().contains(partOfName)) {
				files.add(file.toFile());
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
			CellsCounter cellsCounter;
			if (files.size() == 0 || directories.size() == 0) {
				System.out.println("Необходимо выбрать по крайней мере один файл книги с данными!");
				return;
			} else {
				cellsCounter = new CellsCounter(files, directories, table, advancedGUI);
			}

			int param;

			if (paretoRadio.isSelected()) {
				param = 4;
			} else if (repeatRadio.isSelected()) {
				param = 5;
				cellsCounter.run(param, true, false);
			}

			cellsCounter.run(4, true, false);
		}
	}
}


