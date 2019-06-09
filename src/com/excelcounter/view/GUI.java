package com.excelcounter.view;

import com.excelcounter.controller.CellsCounter;

import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;

public class GUI extends JFrame {

	private File all;
	private File table;

	private JButton startWork = new JButton("Посчитать и записать");

	private JButton allFileChooserButton = new JButton("Выбрать книгу .xlsx с данными");
	private JButton tableFileChooseButton = new JButton("Выбрать книгу .xlsx с таблицей");
	private JLabel allFilePathLabel = new JLabel();
	private JLabel tableFilePathLabel = new JLabel();

	private JLabel tableTypeLabel = new JLabel("Тип таблицы для записи данных");
	private JRadioButton table765Radio = new JRadioButton("765");
	private JRadioButton table753Radio = new JRadioButton("753");
	private JRadioButton tableOrdersRadio = new JRadioButton("Заказы");

	private JCheckBox check = new JCheckBox("Вывести результат подсчета в консоль", true);

	public GUI() {
		super("ExcelCounter v0.4 [Парсер .xlsx таблиц из \"светофора\"]");
		this.setBounds(100, 100, 700, 350);
		this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

		Container container = this.getContentPane();
		container.setLayout(new GridLayout(5, 4, 4, 2));

		container.add(allFileChooserButton);
		container.add(tableFileChooseButton);
		container.add(allFilePathLabel);
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

		container.add(check);

		allFileChooserButton.addActionListener(new allFileChooseButtonActionListener());
		tableFileChooseButton.addActionListener(new TableFileChooseButtonActionListener());
		startWork.addActionListener(new CountButtonEventListener(this));
		container.add(startWork);
	}

	private void setXLSXFilter(JFileChooser fileChooser) {
		FileNameExtensionFilter filter = new FileNameExtensionFilter("XLSX files", "xlsx");
		fileChooser.setFileFilter(filter);
	}

	class allFileChooseButtonActionListener implements ActionListener {

		@Override
		public void actionPerformed(ActionEvent e) {
			JFileChooser allFileChooser = new JFileChooser();
			setXLSXFilter(allFileChooser);
			int ret = allFileChooser.showDialog(null, "Выбрать файл книги с данными");
			if (ret == JFileChooser.APPROVE_OPTION) {
				all = allFileChooser.getSelectedFile();
				allFilePathLabel.setText(all.getAbsolutePath());
			}
		}
	}

	class TableFileChooseButtonActionListener implements ActionListener {
		@Override
		public void actionPerformed(ActionEvent e) {
			JFileChooser tableFileChooser = new JFileChooser();
			setXLSXFilter(tableFileChooser);
			int ret = tableFileChooser.showDialog(null, "Выбрать файл книги с таблицей");
			if (ret == JFileChooser.APPROVE_OPTION) {
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
			if (all == null) {
				System.out.println("Необходимо выбрать файл книги с данными!");
				return;
			}

			if (table != null) {
				cellsCounter = new CellsCounter(all, table);
			} else {
				cellsCounter = new CellsCounter(all);
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
}



