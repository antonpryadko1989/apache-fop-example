package com.visoft.utils.excelReportUtils;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class SummaryReportConverter {

    public static List<String> convertMapToSummaryReportRow(Map<String, String> currentRow, List<String> currentList, int i){
        List<String> result = new ArrayList<>();
        result.add(String.valueOf(i));
        for (String key: currentList) {
            result.add(currentRow.get(key));
        }
        return result;
    }

    public static XSSFWorkbook changeDirectionRightToLeft(XSSFWorkbook workbook) {
        XSSFSheet firstSheet;
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            firstSheet = workbook.getSheetAt(i);
            firstSheet.getCTWorksheet().getSheetViews().getSheetViewArray(0).setRightToLeft(true);
        }
        return workbook;
    }

}
