package com.visoft.services.impl;

import com.visoft.services.Input2OutputService;
import com.visoft.services.XLSXBuilder;
import com.visoft.templates.entity.TemplateDTO;
import com.visoft.templates.entity.XLSXObject;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.io.*;

import static com.visoft.services.PoiXSSFCellStyleService.getChecklistReportStyles;
import static com.visoft.utils.Const.*;
import static com.visoft.services.POIXls.*;


public class ChecklistReportImpl implements XLSXBuilder {

    private Input2OutputService input2OutputService = new Input2OutputService();

    @Override
    public StreamingResponseBody buildXLSX(TemplateDTO template) throws IOException {
        boolean notEmpty = template.templateBodyListNotEmpty(template);
        XSSFWorkbook workbook = new XSSFWorkbook();
        int startRowNum = 1;
        XLSXObject xlsxObject = new XLSXObject(startRowNum, workbook.createSheet(template.getSheetName()));
        setILColumnWidths(xlsxObject.getSheet());
        XSSFCellStyle[] styles = getChecklistReportStyles(xlsxObject);

        addLogo(
                xlsxObject, template.getBody()
        );
        addHeaderInfo(
                xlsxObject, template.getBody().getBodyElements(),
                styles[0], styles[1], styles[2], styles[5], styles[3], styles[4]
        );
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addMainInfo(
                xlsxObject, template,
                styles[0],  styles[11], styles[1],  styles[12], styles[10], styles[13],
                styles[9],  styles[14], styles[5],  styles[15], styles[3],  styles[16]
        );
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        if(notEmpty) {
            if (template.getBody().getBodyLists().get(DRAWINGS_VAL) != null) {
                addDrawings(
                        xlsxObject, DRAWING_NO, DRAWING_VERSION_REVISION, DRAWING_NAME,
                        styles[6], styles[7], styles[8]
                );
                addDrawingsList(
                        xlsxObject, template.getBody().getBodyLists().get(DRAWINGS_VAL),
                        styles[17], styles[11], styles[12], styles[18], styles[13], styles[14],
                        styles[19], styles[15], styles[16], styles[20], styles[21], styles[22]
                );
                addEmptyStringWithGivenHeight(xlsxObject, 9);
            }
            if (template.getBody().getBodyLists().get(CHECKLIST_ELEMENTS_VAL) != null) {
                addChecklistElements(
                        xlsxObject, template.getBody().getBodyLists().get(CHECKLIST_ELEMENTS_VAL),
                        styles[0],  styles[11], styles[1],  styles[12], styles[10], styles[13],
                        styles[9],  styles[14], styles[5],  styles[15], styles[3],  styles[16],
                        styles[6],  styles[21], styles[7],  styles[22]
                );
                addEmptyStringWithGivenHeight(xlsxObject, 9);
            }
            if (template.getBody().getBodyLists().get(CHECKLIST_ITEMS_VAL) != null) {
                addChecklistItem(
                        xlsxObject, CHECKLIST_ITEMS_WORKDEFINITION, CHECKLIST_ITEMS_RESPONSIBLE_PARTY, CHECKLIST_ITEMS_NAME,
                        CHECKLIST_ITEMS_SIGNATURE, CHECKLIST_ITEMS_DATE, CHECKLIST_ITEMS_NOTES,
                        styles[6], styles[7], styles[7], styles[8], true);
                addChecklistItemsList(
                        xlsxObject, template.getBody().getBodyLists().get(CHECKLIST_ITEMS_VAL),
                        styles[0],  styles[1],  styles[11], styles[12], styles[10], styles[9],  styles[13], styles[14],
                        styles[5],  styles[3],  styles[15], styles[16], styles[6],  styles[7],  styles[21], styles[22]
                );
            }
        }
        return input2OutputService.writeWorkBook(workbook);
    }
}