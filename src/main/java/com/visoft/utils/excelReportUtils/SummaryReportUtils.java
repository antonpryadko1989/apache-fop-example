package com.visoft.utils.excelReportUtils;

import com.visoft.templates.entity.XLSXObject;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.util.List;

import static com.visoft.utils.excelReportUtils.XLSXReportUtils.mergeCells;

public class SummaryReportUtils {

    public static void addSummaryReportHeader(XLSXObject xlsxObject,
                                              int from,
                                              int to,
                                              XSSFCellStyle style,
                                              String value,
                                              int height){
        Row row = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        for (int i = from; i<=to; i++){
            row.createCell(i).setCellStyle(style);
        }
        row.getCell(from).setCellValue(value);
        mergeCells(xlsxObject.getSheet(), xlsxObject.getRowNum(), xlsxObject.getRowNum(), from, to);
        row.setHeightInPoints(height);
        xlsxObject.increment();
    }

    public static void addSummaryReportMainInfoRow(XLSXObject xlsxObject,
                                                   int index1, int index2, int index3, int index4, int height,
                                                   String value1, String value2, String value3, String value4,
                                                   XSSFCellStyle style1, XSSFCellStyle style2, XSSFCellStyle style3, XSSFCellStyle style4){
        Row row = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        row.createCell(index1).setCellStyle(style1);
        row.createCell(index2).setCellStyle(style2);
        row.createCell(index3).setCellStyle(style3);
        row.createCell(index4).setCellStyle(style4);
        row.getCell(index1).setCellValue(value1);
        row.getCell(index2).setCellValue(value2);
        row.getCell(index3).setCellValue(value3);
        row.getCell(index4).setCellValue(value4);
        row.setHeightInPoints(height);
        xlsxObject.increment();
    }

    public static void addSummaryReportRow(XLSXObject xlsxObject,
                                           List<String> rowList,
                                           XSSFCellStyle style,
                                           boolean filterFlag){
        Row row = xlsxObject.getSheet().createRow(xlsxObject.getRowNum());
        for (int i = 0; i < rowList.size(); i++) {
            row.createCell(i).setCellStyle(style);
            row.getCell(i).setCellValue(rowList.get(i));
        }
        if(filterFlag){
            row.setHeightInPoints(96);
            xlsxObject.getSheet().setAutoFilter(new CellRangeAddress(xlsxObject.getRowNum(),xlsxObject.getRowNum(),0, rowList.size() - 1));
        }
        xlsxObject.increment();
    }
}
