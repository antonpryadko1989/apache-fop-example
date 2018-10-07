package com.visoft.services.impl;

import com.visoft.cellStyleUtil.FontParams;
import com.visoft.services.Input2OutputService;
import com.visoft.services.XLSXBuilder;
import com.visoft.templates.entity.TemplateDTO;
import com.visoft.templates.entity.XLSXObject;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.io.*;

import static com.visoft.cellStyleUtil.CellStyleUtil.*;
import static com.visoft.cellStyleUtil.FontParams.setFontParams;
import static com.visoft.services.Const.*;
import static com.visoft.services.POIXls.*;


public class ChecklistXlsx implements XLSXBuilder {

    private Input2OutputService input2OutputService = new Input2OutputService();

    @Override
    public StreamingResponseBody buildXLSX(TemplateDTO template) throws IOException {
        boolean notEmpty = template.templateBodyListNotEmpty(template);
        XSSFWorkbook workbook = new XSSFWorkbook();
        int startRowNum = 1;
        XLSXObject xlsxObject = new XLSXObject(startRowNum, workbook.createSheet(template.getSheetName()));
        setILColumnWidths(xlsxObject.getSheet());

        FontParams fontCalibri12bold = setFontParams(FONT_NAME_CALIBRI, true, (short) (12 * FONT_RATE));
        FontParams fontCalibri12 = setFontParams(FONT_NAME_CALIBRI, false, (short) (12 * FONT_RATE));

        XSSFCellStyle lTMediumRBThinCalibri12Bold = getLTMediumRBThinFont(xlsxObject, fontCalibri12bold);
        XSSFCellStyle tMediumLRBThinCalibri12Bold = getTMediumLRBThinFont(xlsxObject, fontCalibri12bold);
        XSSFCellStyle tRMediumLBThinCalibri12Bold = getTRMediumLBThinFont(xlsxObject, fontCalibri12bold);
        XSSFCellStyle bMediumLTRThinCalibri12Bold = getBMediumLTRThinFont(xlsxObject, fontCalibri12bold);
        XSSFCellStyle rBMediumLTThinCalibri12Bold = getRBMediumLTThinFont(xlsxObject, fontCalibri12bold);
        XSSFCellStyle lBMediumTRThinCalibri12Bold = getLBMediumTRThinFont(xlsxObject, fontCalibri12bold);
        XSSFCellStyle lTBMediumRThinCalibri12Bold = getLTBMediumRThinFont(xlsxObject, fontCalibri12bold);
        XSSFCellStyle tBMediumLRThinCalibri12Bold = getTBMediumLRThinFont(xlsxObject, fontCalibri12bold);
        XSSFCellStyle tRBMediumLThinCalibri12bold = getTRBMediumLThinFont(xlsxObject, fontCalibri12bold);
        XSSFCellStyle allBordersThinCalibri12Bold = setAllBordersThin(xlsxObject, fontCalibri12bold);
        XSSFCellStyle lMediumTRBThinCalibri12Bold = getLMediumTRBThinFont(xlsxObject, fontCalibri12bold);

        XSSFCellStyle tMediumLRBThinCalibri12 = getTMediumLRBThinFont(xlsxObject, fontCalibri12);
        XSSFCellStyle tRMediumLBThinCalibri12 = getTRMediumLBThinFont(xlsxObject, fontCalibri12);
        XSSFCellStyle allBordersThinCalibri12 = setAllBordersThin(xlsxObject, fontCalibri12);
        XSSFCellStyle rMediumLTBThinCalibri12 = getRMediumLTBThinFont(xlsxObject, fontCalibri12);
        XSSFCellStyle bMediumLTRThinCalibri12 = getBMediumLTRThinFont(xlsxObject, fontCalibri12);
        XSSFCellStyle rBMediumLTThinCalibri12 = getRBMediumLTThinFont(xlsxObject, fontCalibri12);
        XSSFCellStyle lTMediumRBThinCalibri12 = getLTMediumRBThinFont(xlsxObject, fontCalibri12);
        XSSFCellStyle lMediumTRBThinCalibri12 = getLMediumTRBThinFont(xlsxObject, fontCalibri12);
        XSSFCellStyle lBMediumRTThinCalibri12 = getLBMediumRTThinFont(xlsxObject, fontCalibri12);
        XSSFCellStyle lTBMediumRThinCalibri12 = getLTBMediumRThinFont(xlsxObject, fontCalibri12);
        XSSFCellStyle tBMediumLRThinCalibri12 = getTBMediumLRThinFont(xlsxObject, fontCalibri12);
        XSSFCellStyle tRBMediumLThinCalibri12 = getTRBMediumLThinFont(xlsxObject, fontCalibri12);

        addLogo(xlsxObject, template.getBody());
        addHeaderInfo(xlsxObject, template.getBody().getBodyElements(),
                lTMediumRBThinCalibri12Bold, tMediumLRBThinCalibri12Bold, tRMediumLBThinCalibri12Bold,
                lBMediumTRThinCalibri12Bold, bMediumLTRThinCalibri12Bold, rBMediumLTThinCalibri12Bold);
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addMainInfo(xlsxObject, template,
                lTMediumRBThinCalibri12Bold, tMediumLRBThinCalibri12, tMediumLRBThinCalibri12Bold, tRMediumLBThinCalibri12,
                lMediumTRBThinCalibri12Bold, allBordersThinCalibri12, allBordersThinCalibri12Bold, rMediumLTBThinCalibri12,
                lBMediumTRThinCalibri12Bold, bMediumLTRThinCalibri12, bMediumLTRThinCalibri12Bold, rBMediumLTThinCalibri12);
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        if(notEmpty) {
            if (template.getBody().getBodyLists().get(DRAWINGS_VAL) != null) {
                addDrawings(xlsxObject, DRAWING_NO, DRAWING_VERSION_REVISION, DRAWING_NAME,
                        lTBMediumRThinCalibri12Bold, tBMediumLRThinCalibri12Bold, tRBMediumLThinCalibri12bold);
                addDrawingsList(xlsxObject, template.getBody().getBodyLists().get(DRAWINGS_VAL),
                        lTMediumRBThinCalibri12, tMediumLRBThinCalibri12, tRMediumLBThinCalibri12,
                        lMediumTRBThinCalibri12, allBordersThinCalibri12, rMediumLTBThinCalibri12,
                        lBMediumRTThinCalibri12, bMediumLTRThinCalibri12, rBMediumLTThinCalibri12,
                        lTBMediumRThinCalibri12, tBMediumLRThinCalibri12, tRBMediumLThinCalibri12);
                addEmptyStringWithGivenHeight(xlsxObject, 9);
            }
            if (template.getBody().getBodyLists().get(CHECKLIST_ELEMENTS_VAL) != null) {
                addChecklistElements(xlsxObject, template.getBody().getBodyLists().get(CHECKLIST_ELEMENTS_VAL),
                        lTMediumRBThinCalibri12Bold, tMediumLRBThinCalibri12, tMediumLRBThinCalibri12Bold, tRMediumLBThinCalibri12,
                        lMediumTRBThinCalibri12Bold, allBordersThinCalibri12, allBordersThinCalibri12Bold, rMediumLTBThinCalibri12,
                        lBMediumTRThinCalibri12Bold, bMediumLTRThinCalibri12, bMediumLTRThinCalibri12Bold, rBMediumLTThinCalibri12,
                        lTBMediumRThinCalibri12Bold, tBMediumLRThinCalibri12, tBMediumLRThinCalibri12Bold, tRBMediumLThinCalibri12);
                addEmptyStringWithGivenHeight(xlsxObject, 9);
            }
            if (template.getBody().getBodyLists().get(CHECKLIST_ITEMS_VAL) != null) {
                addChecklistItem(xlsxObject, CHECKLIST_ITEMS_WORKDEFINITION, CHECKLIST_ITEMS_RESPONSIBLE_PARTY, CHECKLIST_ITEMS_NAME,
                        CHECKLIST_ITEMS_SIGNATURE, CHECKLIST_ITEMS_DATE, CHECKLIST_ITEMS_NOTES,
                        lTBMediumRThinCalibri12Bold, tBMediumLRThinCalibri12Bold, tBMediumLRThinCalibri12Bold, tRBMediumLThinCalibri12bold, true);
                addChecklistItemsList(xlsxObject, template.getBody().getBodyLists().get(CHECKLIST_ITEMS_VAL),
                        lTMediumRBThinCalibri12Bold, tMediumLRBThinCalibri12Bold, tMediumLRBThinCalibri12, tRMediumLBThinCalibri12,
                        lMediumTRBThinCalibri12Bold, allBordersThinCalibri12Bold, allBordersThinCalibri12, rMediumLTBThinCalibri12,
                        lBMediumTRThinCalibri12Bold, bMediumLTRThinCalibri12Bold, bMediumLTRThinCalibri12, rBMediumLTThinCalibri12,
                        lTBMediumRThinCalibri12Bold, tBMediumLRThinCalibri12Bold, tBMediumLRThinCalibri12, tRBMediumLThinCalibri12);
            }
        }
        return input2OutputService.writeWorkBook(workbook);
    }
}