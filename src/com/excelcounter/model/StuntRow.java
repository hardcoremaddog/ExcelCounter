package com.excelcounter.model;

public class StuntRow {

    private String country;
    private String product;
    private String viewOfProduct;
    private double cargoTotalWeight;

    public StuntRow() {

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
}
