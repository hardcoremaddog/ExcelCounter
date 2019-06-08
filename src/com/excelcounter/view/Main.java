package com.excelcounter.view;

import com.excelcounter.controller.CellsCounter;

import java.io.*;

public class Main {
	private static File all;
	private static File table;

	public static void main(String[] args) throws IOException {
		while (true) {
			GUI app = new GUI();
			app.setVisible(true);

			BufferedReader br = new BufferedReader(new InputStreamReader(System.in));

			System.out.println("ExcelCounter v0.4" +
					"\nby Alexey Zheludov" +
					"\nCopyright © 2019 MDDG Software, All rights reserved.\n" +

					" \nПриложение разработано исключительно для нужд МВМ ПДУ." +
					" \nИспользование приложения в других целях не гарантирует его корректную работу," +
					" \nа автор не несет ответственности за испорченные, в ходе работы приложения, файлы!" +

					"\nРаботать в приложении можно как с помощью консоли, так и в пользовательском интерфейсе. \n\n");

			System.out.println("Введите полный путь к книге с данными: ");
			String allPath = br.readLine();
			if (allPath.isEmpty()) {
				System.out.println("Путь не указан! \nРабота программы завершена.");
				System.exit(0);
			} else {
				all = new File(allPath);
				checkFile(all);
			}

			int param = 0;
			System.out.println("Введите полный путь к книге с таблицей, куда необходимо записать данные: ");
			String tablePath = br.readLine();
			if (tablePath.isEmpty()) {
				System.out.println("Путь не указан! \nЗапись данных в таблицу не будет осуществлена. \nРезультат работы программы будет выведен на экран.\n");
			} else {
				table = new File(tablePath);

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
			cellsCounter.run(param, true);

			while (true) {
				System.out.println("Введите \"ДА\" и нажмите Enter, если хотите продолжить, или просто нажмите Enter, если хотите завешить работу программы");
				String answer = br.readLine();

				if (answer.equalsIgnoreCase("ДА")) {
					break;
				} else {
					if (!answer.isEmpty())
					System.out.println("Не понял... Еще раз!");
				}

				if (answer.isEmpty()) {
					System.out.println("Было приятно работать с Вами! До скорого :)");
					try {
						Thread.sleep(2000);
					} catch (InterruptedException e) {
						e.printStackTrace();
					}
					System.exit(0);
				}
			}
		}
	}

	private static void checkFile(File file) throws IOException {
		if (!file.exists() && !file.isFile()) {
			throw new FileNotFoundException("Неверно указан путь до книги excel с данными. Книги по этому пути не существует!");
		}
		if (!file.getName().endsWith(".xlsx")) {
			throw new UnsupportedOperationException("Файл не является книгой excel формата .xlsx");
		}
	}
}