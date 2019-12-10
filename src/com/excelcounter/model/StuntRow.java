package com.excelcounter.model;

public class StuntRow {

    private String countryOrigin;
    private String countryImport;
    private String product;
    private String viewOfProduct;
    private double cargoTotalWeight;

    public String getCountryOrigin() {
        return countryOrigin;
    }

    public void setCountryOrigin(String countryOrigin) {
        this.countryOrigin = countryOrigin;
    }

    public String getCountryImport() {
        return countryImport;
    }

    public void setCountryImport(String countryImport) {
        this.countryImport = countryImport;
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
