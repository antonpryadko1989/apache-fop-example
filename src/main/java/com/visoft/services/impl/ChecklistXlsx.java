package com.visoft.services.impl;

import com.visoft.services.XLSXBuilder;
import com.visoft.templates.entity.TemplateDTO;
import com.visoft.templates.entity.XLSXObject;
import org.apache.commons.collections.CollectionUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.xssf.usermodel.*;

import java.util.*;

import static com.visoft.cellStyleUtil.CellStyleUtil.*;
import static com.visoft.cellStyleUtil.FontParams.setFontParams;
import static com.visoft.services.Const.*;
import static com.visoft.services.POIXls.*;

public class ChecklistXlsx implements XLSXBuilder {
    @Override
    public String buildXLSX(TemplateDTO template) {
        long start = System.currentTimeMillis();
        XSSFWorkbook workbook = new XSSFWorkbook();
        int startRowNum = 1;
        boolean notEmpty = template.templateBodyListNotEmpty(template);
        XLSXObject xlsxObject = new XLSXObject(startRowNum,
                workbook.createSheet(template.getSheetName()));
        setILColumnWidths(xlsxObject.getSheet());
        XSSFCellStyle borderThinCalibri12Bold = setAllBordersThin(xlsxObject,
                setFontParams(FONT_NAME_CALIBRI, true,
                        (short) (12 * FONT_RATE)));
        XSSFCellStyle leftBorderMediumCalibri12Bold =
                setBordersAndFont(xlsxObject,
                        setFontParams(FONT_NAME_CALIBRI, true,
                                (short) (12 * FONT_RATE)),
                        BorderStyle.MEDIUM, BorderStyle.THIN,
                        BorderStyle.THIN, BorderStyle.THIN);
        XSSFCellStyle rightBorderMediumCalibri12Bold =
                setBordersAndFont(xlsxObject,
                        setFontParams(FONT_NAME_CALIBRI, true,
                                (short) (12 * FONT_RATE)),
                        BorderStyle.THIN, BorderStyle.THIN,
                        BorderStyle.MEDIUM, BorderStyle.THIN);
        XSSFCellStyle rightBorderMediumCalibri12 =
                setBordersAndFont(xlsxObject,
                        setFontParams(FONT_NAME_CALIBRI, false,
                                (short) (12 * FONT_RATE)),
                        BorderStyle.THIN, BorderStyle.THIN,
                        BorderStyle.MEDIUM, BorderStyle.THIN);
        XSSFCellStyle borderThinCalibri12 = setAllBordersThin(xlsxObject,
                setFontParams(FONT_NAME_CALIBRI, false,
                        (short) (12 * FONT_RATE)));
        addLogo(xlsxObject,
                template.getBody().getBodyElements().get(LOGO_PATH),
                template.getBody().getBodyElements().get(LOGO_VAL));
        addHeaderInfo(xlsxObject,
                template.getBody().getBodyElements(),
                leftBorderMediumCalibri12Bold,
                borderThinCalibri12Bold,
                rightBorderMediumCalibri12Bold);
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        addMainInfo(xlsxObject,
                template,
                borderThinCalibri12Bold,
                borderThinCalibri12);
        addEmptyStringWithGivenHeight(xlsxObject, 9);
        if(notEmpty){
            if(template.getBody().getBodyLists().get(DRAWINGS_VAL)!=null){
                addDrawings(xlsxObject, DRAWING_NO, DRAWING_VERSION_REVISION,
                        DRAWING_NAME, borderThinCalibri12Bold, true);
                addDrawingsList(xlsxObject,
                        template.getBody().getBodyLists().get(DRAWINGS_VAL),
                        borderThinCalibri12);
                addEmptyStringWithGivenHeight(xlsxObject, 9);
            }
            if(template.getBody().getBodyLists().get(CHECKLIST_ELEMENTS_VAL)!=null){
                addChecklistElements(xlsxObject,
                        template.getBody().getBodyLists().get(CHECKLIST_ELEMENTS_VAL),
                        leftBorderMediumCalibri12Bold,
                        borderThinCalibri12,
                        borderThinCalibri12Bold);
                addEmptyStringWithGivenHeight(xlsxObject, 9);
            }
            if(template.getBody().getBodyLists().get(CHECKLIST_ITEMS_VAL)!=null){
                addChecklistItems(xlsxObject,
                        template.getBody().getBodyLists().get(CHECKLIST_ITEMS_VAL),
                        leftBorderMediumCalibri12Bold,
                        borderThinCalibri12Bold,
                        borderThinCalibri12,
                        rightBorderMediumCalibri12Bold,
                        rightBorderMediumCalibri12);
            }
        }
        System.out.println(CHECKLIST + " " +
                (System.currentTimeMillis() - start));
        return writeWorkBook(workbook);
    }

    private void addChecklistItems(XLSXObject xlsxObject,
                                   List<Map<String, String>>
                                           checklistItems,
                                   XSSFCellStyle styleOne,
                                   XSSFCellStyle styleTwo,
                                   XSSFCellStyle styleThree,
                                   XSSFCellStyle styleFour,
                                   XSSFCellStyle styleFive) {
        if (!CollectionUtils.isEmpty(checklistItems)) {
            addChecklistItem(xlsxObject,
                    CHECKLIST_ITEMS_WORKDEFINITION,
                    CHECKLIST_ITEMS_RESPONSIBLE_PARTY,
                    CHECKLIST_ITEMS_NAME,
                    CHECKLIST_ITEMS_SIGNATURE,
                    CHECKLIST_ITEMS_DATE,
                    CHECKLIST_ITEMS_NOTES,
                    styleOne, styleTwo, styleTwo, styleFour, true);
            addChecklistItemsList(xlsxObject, checklistItems,
                    styleOne, styleTwo, styleThree, styleFive);
        }
    }
}