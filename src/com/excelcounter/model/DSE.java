package com.excelcounter.model;

public class DSE {
    private String vendorCode;
    private String nomenclature;

    public DSE(String vendorCode, String nomenclature) {
        this.vendorCode = vendorCode;
        this.nomenclature = nomenclature;
    }

    public String getVendorCode() {
        return vendorCode;
    }

    public String getNomenclature() {
        return nomenclature;
    }
}
