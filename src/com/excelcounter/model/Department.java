package com.excelcounter.model;

import java.util.ArrayList;
import java.util.List;
import java.util.Set;
import java.util.TreeSet;

public class Department {

    private String name;
    private String fileName;

    private int redCellsCount;
    private int yellowCellsCount;

    private List<OperationLvl> operationsLvlList = new ArrayList<>();

    private List<Order> ordersList = new ArrayList<>();
    private Set<String> dseTreeSet = new TreeSet<>();

    //getters and setters
    public Set<String> getDseTreeSet() {
        return dseTreeSet;
    }

    public void setDseTreeSet(Set<String> dseTreeSet) {
        this.dseTreeSet = dseTreeSet;
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

    public List<OperationLvl> getOperationsLvlList() {
        return operationsLvlList;
    }

    public List<Order> getOrdersList() {
        return ordersList;
    }
}
