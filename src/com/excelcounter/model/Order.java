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

    private double primenyaemostCount;
    private double faktReleaseCount = 0;
    private String zakazPokypatelya;

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

    public String getZakazPokypatelya() {
        return zakazPokypatelya;
    }

    public void setZakazPokypatelya(String zakazPokypatelya) {
        this.zakazPokypatelya = zakazPokypatelya;
    }

    public double getPrimenyaemostCount() {
        return primenyaemostCount;
    }

    public void setPrimenyaemostCount(double primenyaemostCount) {
        this.primenyaemostCount = primenyaemostCount;
    }

    public double getFaktReleaseCount() {
        return faktReleaseCount;
    }

    public void setFaktReleaseCount(double faktReleaseCount) {
        this.faktReleaseCount = faktReleaseCount;
    }
}


