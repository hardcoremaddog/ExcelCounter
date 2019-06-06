package com.excelcounter;

import com.excelcounter.controller.CellsCounter;

import java.io.*;

public class Main {
	private static File all;

	public static void main(String[] args) throws IOException {

		BufferedReader br = new BufferedReader(new InputStreamReader(System.in));

		System.out.println("Введите полный путь к книге с данными: ");
		all = new File(br.readLine());
		if (!all.exists() && !all.isFile()) {
			throw new FileNotFoundException("Неверно указан путь до книги excel с данными. Книги по этому пути не существует!");
		}
		if (!all.getName().endsWith(".xlsx")) {
			throw new UnsupportedOperationException("Файл не является книгой excel формата .xlsx");
		}

		CellsCounter cellsCounter = new CellsCounter(all);
		System.out.println("Всё ок! Начинаю работу...");
		cellsCounter.run();
	}
}