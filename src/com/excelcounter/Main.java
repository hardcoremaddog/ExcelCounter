package com.excelcounter;

import java.io.*;

public class Main {
//	private File all;
//	private File table;

	private final File all = new File("Z:\\Общая для МВМ\\!ПДУ_ОД\\!Укомплектованность оперативных планов\\!765\\2019\\05.Май\\test\\full.xlsx");
	private final File table = new File("Z:\\Общая для МВМ\\!ПДУ_ОД\\!Укомплектованность оперативных планов\\!765\\2019\\05.Май\\test\\[765] Сводная таблица [29.05.19].xlsx");

	public static void main(String[] args) throws IOException {

		BufferedReader br = new BufferedReader(new InputStreamReader(System.in));

		Main main = new Main();
//
//		System.out.println("Введите полный путь к книге с данными: ");
//		main.all = new File(br.readLine());
//		if (!main.all.exists() && !main.all.isFile()) {
//			throw new FileNotFoundException("Неверно указан путь до книги excel с данными. Книги по этому пути не существует!");
//		}
//		if (!main.all.getName().endsWith(".xlsx")) {
//			throw new UnsupportedOperationException("Файл не является книгой excel формата .xlsx");
//		}
//
//		System.out.println("Введите полный путь до книги с таблицей: ");
//		main.table = new File(br.readLine());
//		if (!main.table.exists() && !main.table.isFile()) {
//			throw new FileNotFoundException("Неверно указан путь до сводной таблицы. Таблицы по этому пути не существует!");
//		}
//		if (!main.table.getName().endsWith(".xlsx")) {
//			throw new UnsupportedOperationException("Файл не является книгой excel формата .xlsx");
//		}

		CellsCounter cellsCounter = new CellsCounter(main.all, main.table);
		System.out.println("Всё ок! Начинаю работу...");
		cellsCounter.run();
	}
}
