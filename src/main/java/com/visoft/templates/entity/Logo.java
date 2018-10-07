package com.visoft.templates.entity;

public class Logo {

    private int countCells;
    private String path;

    public Logo() {

    }

    public Logo(int countCells, String path) {
        this.countCells = countCells;
        this.path = path;
    }

    public int getCountCells() {
        return countCells;
    }

    public void setCountCells(int countCells) {
        this.countCells = countCells;
    }

    public String getPath() {
        return path;
    }

    public void setPath(String path) {
        this.path = path;
    }

    @Override
    public String toString() {
        return "Logo{" +
                "countCells=" + countCells +
                ", path='" + path + '\'' +
                '}';
    }
}