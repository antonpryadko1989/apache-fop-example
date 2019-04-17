package com.visoft.dto;

import java.util.Arrays;

public class ReportLogo {
    private int countCells;
    private byte[] data;

    public ReportLogo(int countCells, byte[] data) {
        this.countCells = countCells;
        this.data = data;
    }

    public ReportLogo() {
    }

    public int getCountCells() {
        return countCells;
    }

    public void setCountCells(int countCells) {
        this.countCells = countCells;
    }

    public byte[] getData() {
        return data;
    }

    public void setData(byte[] data) {
        this.data = data;
    }

    @Override
    public String toString() {
        return "ReportLogo{" +
                "countCells=" + countCells +
                ", data=" + Arrays.toString(data) +
                '}';
    }
}
