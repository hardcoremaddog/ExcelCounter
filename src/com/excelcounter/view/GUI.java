package com.excelcounter.view;

import com.excelcounter.controller.CellsCounter;

import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import javax.swing.*;

public class GUI extends JFrame {
	private JButton startWork = new JButton("Посчитать и записать");

	private JTextField allPathField = new JTextField("", 5);
	private JTextField tablePathField = new JTextField("", 5);

	private JLabel allPathFieldLabel = new JLabel("Полный путь до книги с данными");
	private JLabel tablePathFieldLabel = new JLabel("Полный путь до книги с таблицей (опционально)");
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
		container.setLayout(new GridLayout(5, 2, 2, 2));
		container.add(allPathFieldLabel);
		container.add(tablePathFieldLabel);

		container.add(allPathField);
		container.add(tablePathField);

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

		startWork.addActionListener(new CountButtonEventListener(this));
		container.add(startWork);
	}

	class CountButtonEventListener implements ActionListener {

		private GUI gui;

		private CountButtonEventListener(GUI gui) {
			this.gui = gui;
		}

		@Override
		public void actionPerformed(ActionEvent e) {

			if (gui.allPathField.getText().isEmpty()) {
				System.out.println("Необходимо указать путь до книги с данными!");
			} else {
				CellsCounter cellsCounter;
				if (!gui.tablePathField.getText().isEmpty()) {
					cellsCounter = new CellsCounter(new File(gui.allPathField.getText()), new File(gui.tablePathField.getText()));
				} else {
					cellsCounter = new CellsCounter(new File(gui.allPathField.getText()));
				}

				int param;
				if (table753Radio.isSelected()) {
					param = 2;
				}
				else if (tableOrdersRadio.isSelected()) {
					param = 3;
				} else {
					param = 1;
				}

				cellsCounter.run(param, gui.check.isSelected());
			}
		}
	}
}



