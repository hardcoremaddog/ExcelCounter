package com.excelcounter.view;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

public class AdvancedGUI extends JFrame {
	private AdvancedGUI advancedGUI = this;
	private File[] files;

	private JButton mainGUIShowButton = new JButton("Вернуться на основной интерфейс");

	private JCheckBox check = new JCheckBox("Вывести результат подсчета в консоль", true);

	private JButton startWork = new JButton("Посчитать и записать");

	AdvancedGUI() {
		super("ExcelCounter v0.6a");
		this.setBounds(100, 100, 900, 350);
		this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

		Container container = this.getContentPane();
		container.setLayout(new GridLayout(5, 1, 1, 1));

		mainGUIShowButton.addActionListener(new MainGUIShowButton());
		container.add(mainGUIShowButton);
	}

	private void setXLSXFilter(JFileChooser fileChooser) {
		FileNameExtensionFilter filter = new FileNameExtensionFilter("XLSX files", "xlsx");
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
}
