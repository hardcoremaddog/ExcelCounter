package com.excelcounter.model;

import java.util.ArrayList;
import java.util.List;

public class Day {

    private String dayNumber;
    private String fileName;

    private int tmcCount;
    private int dseCount;

    private List<Department> departmentList;

    public Day(String dayNumber) {
        this.dayNumber = dayNumber;
        this.departmentList = new ArrayList<>();
    }

    public List<Department> getDepartmentList() {
        return departmentList;
    }

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public String getDayNumber() {
        return dayNumber;
    }

    public int getTmcCount() {
        return tmcCount;
    }

    public void setTmcCount(int tmcCount) {
        this.tmcCount = tmcCount;
    }

    public int getDseCount() {
        return dseCount;
    }

    public void setDseCount(int dseCount) {
        this.dseCount = dseCount;
    }
}
