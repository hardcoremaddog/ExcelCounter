package com.excelcounter.model;

import java.util.Set;
import java.util.TreeSet;

public class Department {

    private String name;
    private String fileName;

    private int redCellsCount;
    private int yellowCellsCount;

    private Set<String> dseTreeSet = new TreeSet<>();

    public Set<String> getDseTreeSet() {
        return dseTreeSet;
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
