package com.excelcounter.view;

import javax.swing.*;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;

public class Main {
//	private static File all;
//	private static File table;
	public static final String VERSION = "0.7a";

	public static void main(String[] args) throws IOException {
		try {
			UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
		} catch (ClassNotFoundException | UnsupportedLookAndFeelException | IllegalAccessException | InstantiationException e) {
			e.printStackTrace();
		}

		GUI app = new GUI();
		app.setVisible(true);

		System.out.println("ExcelCounter " +  VERSION +
				"\nby Alexey Zheludov" +
				"\nCopyright (c) 2019 MDDG Software, All rights reserved.\n" +

				"\nПриложение разработано исключительно для подсчета" +
				"\nкрасных и желтых строк в формах отчета" +
				"\n\"Укомплектованность оперативных планов\" и" +
				"\nподсчета количества непереданных позиций на сбыт" +
				"\nв формах отчета \"Контроль передачи на сбыт\"" +
				"\nа также записи результата в книги Excel." +
				"\nПриложение работает ТОЛЬКО с книгами формата .xlsx и .xlsm." +
				"\nИспользование приложения в других целях не гарантирует" +
				"\nего корректную работу, а автор не несет ответственности" +
				"\nза испорченные, в ходе работы приложения, файлы!" +
				"\n\nВ виду элементарности функционала и наличия" +
				"\nпростого пользовательского интерфейса - " +
				"\nдокументация и инструкции по работе с приложением" +
				"\nне предусмотрены.");

//				"\nРаботать в приложении можно как с помощью консоли, так и в пользовательском интерфейсе. \n\n")

//		while (true) {
//			BufferedReader br = new BufferedReader(new InputStreamReader(System.in));


//			System.out.println("Введите полный путь к книге с данными: ");
//			String allPath = br.readLine();
//			if (allPath.isEmpty()) {
//				System.out.println("Путь не указан! \nРабота программы завершена.");
//				System.exit(0);
//			} else {
//				all = new File(allPath);
//				checkFile(all);
//			}

//			int param = 0;
//			System.out.println("Введите полный путь к книге с таблицей, куда необходимо записать данные: ");
//			String tablePath = br.readLine();
//			if (tablePath.isEmpty()) {
//				System.out.println("Путь не указан! \nЗапись данных в таблицу не будет осуществлена. \nРезультат работы программы будет выведен на экран.\n");
//			} else {
//				table = new File(tablePath);
//
//				while (true) {
//					System.out.println("Введите номер, соответствующий типу таблицы, куда будут записаны данные:" +
//							"\n 1 - 765" +
//							"\n 2 - 753" +
//							"\n 3 - Заказы");
//
//					try {
//						param = Integer.parseInt(br.readLine());
//					} catch (NumberFormatException e) {
//						System.out.println("Необходимо ввести номер!");
//						continue;
//					}
//					if (param < 1 || param > 3) {
//						System.out.println("Указан неверный номер, попробуйте еще раз.");
//					} else {
//						break;
//					}
//				}
//			}

//			CellsCounter cellsCounter;
//			if (all != null && table != null) {
//				cellsCounter = new CellsCounter(all, table);
//			} else if (all != null) {
//				cellsCounter = new CellsCounter(all);
//			} else {
//				continue;
//			}
//
//			System.out.println("ОК. Файлы прошли проверку. Начинаю работу...");
//			cellsCounter.run(param, true);
//
//			while (true) {
//				System.out.println("Введите \"ДА\" и нажмите Enter, если хотите продолжить, или просто нажмите Enter, если хотите завешить работу программы");
//				String answer = br.readLine();
//
//				if (answer.equalsIgnoreCase("ДА")) {
//					break;
//				} else {
//					if (!answer.isEmpty())
//					System.out.println("Не понял... Еще раз!");
//				}
//
//				if (answer.isEmpty()) {
//					System.out.println("Было приятно работать с Вами! До скорого :)");
//					try {
//						Thread.sleep(2000);
//					} catch (InterruptedException e) {
//						e.printStackTrace();
//					}
//					System.exit(0);
//				}
//			}
//		}
	}

	protected static void checkFile(File file) throws IOException {
		if (!file.exists() && !file.isFile()) {
			throw new FileNotFoundException("Неверно указан путь до книги excel с данными. Книги по этому пути не существует!");
		}
		if (!file.getName().endsWith(".xlsx") && !file.getName().endsWith(".xlsm")) {
			throw new UnsupportedOperationException("Файл не является книгой excel формата .xlsx или .xlsm");
		}
	}
}