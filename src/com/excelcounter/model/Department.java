package com.excelcounter.model;

import java.util.ArrayList;
import java.util.List;

public class Department {

	private String name;
	private String fileName;

	private int redCellsCount;
	private int yellowCellsCount;

	private List<String> dseList = new ArrayList<>();

	public List<String> getDseList() {
		return dseList;
	}

	public Department(String name) {
		this.name = name;
	}

	public String getFileName() {
		return fileName;
	}

	public void setFileName(String fileName) {
		this.fileName = fileName;
	}

	public String getName() {
		return name;
	}

	public int getRedCellsCount() {
		return redCellsCount;
	}

	public void setRedCellsCount(int redCellsCount) {
		this.redCellsCount = redCellsCount;
	}

	public int getYellowCellsCount() {
		return yellowCellsCount;
	}

	public void setYellowCellsCount(int yellowCellsCount) {
		this.yellowCellsCount = yellowCellsCount;
	}
}
