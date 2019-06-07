package com.excelcounter;

import com.excelcounter.controller.CellsCounter;

import java.io.*;

public class Main {
	private static File all;
	private static File table;

	public static void main(String[] args) throws IOException {
		BufferedReader br = new BufferedReader(new InputStreamReader(System.in));

		System.out.println("Введите полный путь к книге с данными: ");
		String allPath = br.readLine();
		if (allPath.isEmpty()) {
			System.out.println("Путь не указан! \nРабота программы завершена.");
			System.exit(0);
		} else {
			all = new File(allPath);
			if (!all.exists() && !all.isFile()) {
				throw new FileNotFoundException("Неверно указан путь до книги excel с данными. Книги по этому пути не существует!");
			}
			if (!all.getName().endsWith(".xlsx")) {
				throw new UnsupportedOperationException("Файл не является книгой excel формата .xlsx");
			}
		}

		int param = 0;
		System.out.println("Введите полный путь к книге с таблицей, куда необходимо записать данные: ");
		String tablePath = br.readLine();
		if (tablePath.isEmpty()) {
			System.out.println("Путь не указан! \nЗапись данных в таблицу не будет осуществлена. \nРезультат работы программы будет выведен на экран.\n");
		} else {
			table = new File(tablePath);
			if (!table.exists() && !table.isFile()) {
				throw new FileNotFoundException("Неверно указан путь до книги excel с данными. Книги по этому пути не существует!");
			}
			if (!table.getName().endsWith(".xlsx")) {
				throw new UnsupportedOperationException("Файл не является книгой excel формата .xlsx");
			}

			while (true) {
				System.out.println("Введите номер, соответствующий типу таблицы, куда будут записаны данные:" +
						"\n 1 - 765" +
						"\n 2 - 753" +
						"\n 3 - Заказы");

				try {
					param = Integer.parseInt(br.readLine());
				} catch (NumberFormatException e) {
					System.out.println("Необходимо ввести номер!");
					continue;
				}
				if (param < 1 || param > 3) {
					System.out.println("Указан неверный номер, попробуйте еще раз.");
				} else {
					break;
				}
			}
		}

		CellsCounter cellsCounter;
		if (table == null) {
			cellsCounter = new CellsCounter(all);
		} else {
			cellsCounter = new CellsCounter(all, table);
		}

		System.out.println("ОК. Файлы прошли проверку. Начинаю работу...");
		cellsCounter.run(param);
	}
}