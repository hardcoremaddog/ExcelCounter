package com.excelcounter.model;

import java.util.HashMap;
import java.util.Map;

public class Department {

    private String name;

    private int redCellsCount;
    private int yellowCellsCount;

    private Map<String, Integer> dseRepeatCountMap;

    public Department(String name) {
        this.name = name;
        this.dseRepeatCountMap = new HashMap<>();
    }

    public Map<String, Integer> getDseRepeatCountMap() {
        return dseRepeatCountMap;
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
