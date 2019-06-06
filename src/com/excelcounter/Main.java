package com.excelcounter;

import java.io.*;

public class Main {
//	private File all;
//	private File table;

	//home
	private static final File all = new File("D:\\1\\all3.xlsx");

	public static void main(String[] args) throws IOException {

		BufferedReader br = new BufferedReader(new InputStreamReader(System.in));
//
//		System.out.println("Введите полный путь к книге с данными: ");
//		main.all = new File(br.readLine());
//		if (!main.all.exists() && !main.all.isFile()) {
//			throw new FileNotFoundException("Неверно указан путь до книги excel с данными. Книги по этому пути не существует!");
//		}
//		if (!main.all.getName().endsWith(".xlsx")) {
//			throw new UnsupportedOperationException("Файл не является книгой excel формата .xlsx");
//		}

		CellsCounter cellsCounter = new CellsCounter(all);
		System.out.println("Всё ок! Начинаю работу...");
		cellsCounter.run();
	}
}