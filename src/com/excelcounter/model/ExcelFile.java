package com.excelcounter.model;

import java.util.ArrayList;
import java.util.List;

public class ExcelFile {

	private String fileName;
	private List<Department> departmentList;

	public ExcelFile(String fileName) {
		this.fileName = fileName;
		this.departmentList = new ArrayList<>();
	}

	public String getFileName() {
		return fileName;
	}

	public List<Department> getDepartmentList() {
		return departmentList;
	}
}
