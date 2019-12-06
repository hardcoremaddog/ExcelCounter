package com.excelcounter.model;

public class StuntRow {

    private int numberPP;

    private int year;
    private String country;
    private String product;
    private String viewOfProduct;
    private double cargoTotalWeight;

    public StuntRow(int numberPP) {
        this.numberPP = numberPP;
    }

    public int getYear() {
        return year;
    }

    public void setYear(int year) {
        this.year = year;
    }

    public String getCountry() {
        return country;
    }

    public void setCountry(String country) {
        this.country = country;
    }

    public String getProduct() {
        return product;
    }

    public void setProduct(String product) {
        this.product = product;
    }

    public String getViewOfProduct() {
        return viewOfProduct;
    }

    public void setViewOfProduct(String viewOfProduct) {
        this.viewOfProduct = viewOfProduct;
    }

    public double getCargoTotalWeight() {
        return cargoTotalWeight;
    }

    public void setCargoTotalWeight(double cargoTotalWeight) {
        this.cargoTotalWeight = cargoTotalWeight;
    }

    public void setCargoTotalWeight(int cargoTotalWeight) {
        this.cargoTotalWeight = cargoTotalWeight;
    }
}
