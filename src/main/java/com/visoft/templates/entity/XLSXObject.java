package com.visoft.templates.entity;

import org.apache.poi.ss.usermodel.Sheet;

public class XLSXObject {
    private int rowNum;
    private Sheet sheet;

    public XLSXObject(int rowNum, Sheet sheet) {
        this.rowNum = rowNum;
        this.sheet = sheet;
    }

    public int getRowNum() {
        return rowNum;
    }

    public Sheet getSheet() {
        return sheet;
    }

    public void setSheet(Sheet sheet) {
        this.sheet = sheet;
    }

    public int increment(){
        return increment(1);
    }

    public int increment(int step){
        return rowNum = rowNum + step;
    }

}
