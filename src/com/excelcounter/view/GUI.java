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
	private JButton tableFileChooseButton = new JButton("Выбрать книгу .xlsx с таблицей");

	private JButton advancedGUIShowButton = new JButton("Перейти на дополнительный интерфейс");

	private JLabel allFilePathLabel = new JLabel();
	private JLabel sbytFilePathLaber = new JLabel();
	private JLabel tableFilePathLabel = new JLabel();

	private JLabel tableTypeLabel = new JLabel("Тип таблицы для записи данных");
	private JRadioButton table765Radio = new JRadioButton("765");
	private JRadioButton table753Radio = new JRadioButton("753");
	private JRadioButton tableOrdersRadio = new JRadioButton("Заказы");

	private JCheckBox check = new JCheckBox("Вывести результат подсчета в консоль", true);

	private JButton startWork = new JButton("Посчитать и записать");

	public JProgressBar progressBar = new JProgressBar();

	GUI() {
		super("ExcelCounter v0.6a");
		this.setBounds(100, 100, 900, 350);
		this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

		Container container = this.getContentPane();
		container.setLayout(new GridLayout(5, 4, 1, 1));

		container.add(allFileChooserButton);
		container.add(sbytFileChooseButton);
		container.add(tableFileChooseButton);
		container.add(allFilePathLabel);
		container.add(sbytFilePathLaber);
		container.add(tableFilePathLabel);

		ButtonGroup group = new ButtonGroup();
		group.add(table765Radio);
		group.add(table753Radio);
		group.add(tableOrdersRadio);

		container.add(tableTypeLabel);
		container.add(table765Radio);
		table765Radio.setSelected(true);
		container.add(table753Radio);
		container.add(tableOrdersRadio);

		allFileChooserButton.addActionListener(new allFileChooseButtonActionListener());
		sbytFileChooseButton.addActionListener(new sbytFileChooseButtonActionListener());
		tableFileChooseButton.addActionListener(new TableFileChooseButtonActionListener());
		advancedGUIShowButton.addActionListener(new AdvancedGUIShowButtonActionListener());
		advancedGUIShowButton.setDefaultCapable(true);
		startWork.addActionListener(new CountButtonEventListener(this));
		container.add(check);
		container.add(startWork);

		//fakeLabels
		progressBar.setStringPainted(true);
		progressBar.setMinimum(0);
		progressBar.setMaximum(100);
		container.add(progressBar);
		container.add(new JLabel());

		container.add(advancedGUIShowButton);
	}

	private void setXLSXFilter(JFileChooser fileChooser) {
		FileNameExtensionFilter filter = new FileNameExtensionFilter("XLSX files", "xlsx");
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
			setXLSXFilter(allFileChooser);
			int ret = allFileChooser.showDialog(null, "Выбрать файл книги с данными");
			if (ret == JFileChooser.APPROVE_OPTION) {
				checkFile(allFileChooser);
				all = allFileChooser.getSelectedFile();
				allFilePathLabel.setText(all.getAbsolutePath());
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
				sbytFilePathLaber.setText(sbyt.getAbsolutePath());
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
				tableFilePathLabel.setText(table.getAbsolutePath());
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
			if (all == null && sbyt == null) {
				System.out.println("Необходимо выбрать по крайней мере один файл книги с данными!");
				return;
			} else if (table != null) {
				cellsCounter = new CellsCounter(all, table, sbyt, gui);
			} else {
				cellsCounter = new CellsCounter(all, sbyt, gui);
			}

			int param;
			if (table753Radio.isSelected()) {
				param = 2;
			} else if (tableOrdersRadio.isSelected()) {
				param = 3;
			} else {
				param = 1;
			}
			cellsCounter.run(param, gui.check.isSelected());
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
