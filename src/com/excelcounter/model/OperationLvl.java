package com.excelcounter.model;

import java.util.ArrayList;
import java.util.List;

public class OperationLvl {

    public OperationLvl(String name) {
        this.name = name;
    }

    private String name;
    private List<Order> ordersList = new ArrayList<>();

    public String getName() {
        return name;
    }

    public List<Order> getOrdersList() {
        return ordersList;
    }
}
