package com.excelcounter.model;

public class Day {

    private String fileName;

    private int tmcCount;
    private int dseCount;

    public Day(String fileName) {
        this.fileName = fileName;
    }

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
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
