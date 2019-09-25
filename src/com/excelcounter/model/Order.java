package com.excelcounter.model;

import java.util.ArrayList;

public class Order {

    private String name;
    private ArrayList<Department> departments = new ArrayList<>();

    public Order(String name) {
        this.name = name;
    }

    private int redCells;
    private int yellowCells;

    public ArrayList<Department> getDepartments() {
        return departments;
    }

    public String getName() {
        return name;
    }

    public int getRedCells() {
        return redCells;
    }

    public void setRedCells(int redCells) {
        this.redCells = redCells;
    }

    public int getYellowCells() {
        return yellowCells;
    }

    public void setYellowCells(int yellowCells) {
        this.yellowCells = yellowCells;
    }
}
