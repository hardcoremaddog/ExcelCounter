package com.excelcounter.model;

public class Department {

	private String name;

	private int redCellsCount;
	private int yellowCellsCount;

	public Department(String name) {
		this.name = name;
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
