package com.visoft.services;

import com.visoft.cellStyleUtil.FontParams;
import com.visoft.templates.entity.XLSXObject;
import com.visoft.utils.Const;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import static com.visoft.cellStyleUtil.CellStyleUtil.*;
import static com.visoft.cellStyleUtil.CellStyleUtil.getLBMediumTRThinFont;
import static com.visoft.cellStyleUtil.CellStyleUtil.getRBMediumLTThinFont;
import static com.visoft.cellStyleUtil.FontParams.setFontParams;
import static com.visoft.utils.excelReportUtils.FontConst.FONT_NAME_ARIAL;
import static com.visoft.utils.excelReportUtils.FontConst.FONT_RATE;

public interface XSSFCellStyleService {

    FontParams[] fonts = {
            /*0*/setFontParams(FONT_NAME_ARIAL, true, true, (short) (20 * Const.FONT_RATE)),
            /*1*/setFontParams(FONT_NAME_ARIAL, true, (short) (20 * FONT_RATE)),
            /*2*/setFontParams(FONT_NAME_ARIAL, true, (short) (12 * FONT_RATE)),
            /*3*/setFontParams(FONT_NAME_ARIAL, false, (short) (12 * FONT_RATE))
    };

    static XSSFCellStyle[] getHandlingTrafficAccidentReportStyles(XLSXObject xlsxObject) {
        return new XSSFCellStyle[]{
                /* 0*/setAllBordersByStyle(xlsxObject, fonts[1], BorderStyle.MEDIUM, VerticalAlignment.BOTTOM, HorizontalAlignment.CENTER),
                /* 1*/setAlignmentByParam(xlsxObject, fonts[1], Alignment.CENTER, BorderStyle.MEDIUM),
                /* 2*/setAllBordersByStyle(xlsxObject, fonts[3], BorderStyle.THIN),
                /* 3*/getLTMediumRBThinFont(xlsxObject, fonts[2]),
                /* 4*/getTMediumLRBThinFont(xlsxObject, fonts[2]),
                /* 5*/getTRMediumLBThinFont(xlsxObject, fonts[2]),
                /* 6*/getLBMediumRTThinFont(xlsxObject, fonts[2]),
                /* 7*/getBMediumLTRThinFont(xlsxObject, fonts[2]),
                /* 8*/getRBMediumLTThinFont(xlsxObject, fonts[2]),
                /* 9*/getLBMediumTRThinFont(xlsxObject, fonts[2]),
                /*10*/getLTBMediumRThinFont(xlsxObject, fonts[2]),
                /*11*/getTRBMediumLThinFont(xlsxObject, fonts[2]),
                /*12*/getTBMediumLRThinFont(xlsxObject, fonts[2]),
                /*13*/getLMediumTRBThinFont(xlsxObject, fonts[2]),
                /*14*/getTMediumLRBThinFont(xlsxObject, fonts[3]),
                /*15*/getTRMediumLBThinFont(xlsxObject, fonts[3]),
                /*16*/getBMediumLTRThinFont(xlsxObject, fonts[3]),
                /*17*/getRBMediumLTThinFont(xlsxObject, fonts[3]),
                /*18*/getTBMediumLRThinFont(xlsxObject, fonts[3]),
                /*19*/getTRBMediumLThinFont(xlsxObject, fonts[3]),
                /*20*/getRMediumLTBThinFont(xlsxObject, fonts[3]),
                /*21*/getLTMediumRBThin(xlsxObject)
        };
    }

    static XSSFCellStyle[] getFailesSafetyReportStyles(XLSXObject xlsxObject) {
        return new XSSFCellStyle[]{
                /* 0*/setAllBordersByStyle(xlsxObject, fonts[1], BorderStyle.MEDIUM, VerticalAlignment.BOTTOM, HorizontalAlignment.CENTER),
                /* 1*/setAllBordersByStyle(xlsxObject, fonts[3], BorderStyle.THIN, VerticalAlignment.CENTER, HorizontalAlignment.CENTER),
                /* 2*/setAllBordersByStyle(xlsxObject, fonts[2], BorderStyle.THIN, VerticalAlignment.CENTER, HorizontalAlignment.CENTER),
                /* 3*/getLTMediumRBThinFont(xlsxObject, fonts[2]),
                /* 4*/getTMediumLRBThinFont(xlsxObject, fonts[2]),
                /* 5*/getTRMediumLBThinFont(xlsxObject, fonts[2]),
                /* 6*/getLBMediumRTThinFont(xlsxObject, fonts[2]),
                /* 7*/getBMediumLTRThinFont(xlsxObject, fonts[2]),
                /* 8*/getRBMediumLTThinFont(xlsxObject, fonts[2]),
                /* 9*/getLMediumTRBThinFont(xlsxObject, fonts[2]),
                /*10*/getLBMediumTRThinFont(xlsxObject, fonts[2]),
                /*11*/getLTBMediumRThinFont(xlsxObject, fonts[2]),
                /*12*/getTRBMediumLThinFont(xlsxObject, fonts[2]),
                /*13*/getTBMediumLRThinFont(xlsxObject, fonts[2]),
                /*14*/getTMediumLRBThinFont(xlsxObject, fonts[3]),
                /*15*/getTRMediumLBThinFont(xlsxObject, fonts[3]),
                /*16*/getRMediumLTBThinFont(xlsxObject, fonts[3]),
                /*17*/getBMediumLTRThinFont(xlsxObject, fonts[3]),
                /*18*/getRBMediumLTThinFont(xlsxObject, fonts[3]),
                /*19*/getTRBMediumLThinFont(xlsxObject, fonts[3]),
                /*20*/getTBMediumLRThinFont(xlsxObject, fonts[3]),
                /*21*/getLTMediumRBThin(xlsxObject)
        };
    }

    static XSSFCellStyle[] getTreatmentOfImpairedByMonitoringReportStyles(XLSXObject xlsxObject) {
        return new XSSFCellStyle[]{
                /* 0*/setAllBordersByStyle(xlsxObject, fonts[1], BorderStyle.MEDIUM, VerticalAlignment.BOTTOM, HorizontalAlignment.CENTER),
                /* 1*/setAllBordersByStyle(xlsxObject, fonts[3], BorderStyle.THIN, VerticalAlignment.CENTER, HorizontalAlignment.CENTER),
                /* 2*/setAllBordersByStyle(xlsxObject, fonts[2], BorderStyle.THIN, VerticalAlignment.CENTER, HorizontalAlignment.CENTER),
                /* 3*/getLTMediumRBThinFont(xlsxObject, fonts[2]),
                /* 4*/getTMediumLRBThinFont(xlsxObject, fonts[2]),
                /* 5*/getTRMediumLBThinFont(xlsxObject, fonts[2]),
                /* 6*/getLBMediumRTThinFont(xlsxObject, fonts[2]),
                /* 7*/getBMediumLTRThinFont(xlsxObject, fonts[2]),
                /* 8*/getRBMediumLTThinFont(xlsxObject, fonts[2]),
                /* 9*/getLMediumTRBThinFont(xlsxObject, fonts[2]),
                /*10*/getLBMediumTRThinFont(xlsxObject, fonts[2]),
                /*11*/getLTBMediumRThinFont(xlsxObject, fonts[2]),
                /*12*/getTRBMediumLThinFont(xlsxObject, fonts[2]),
                /*13*/getTBMediumLRThinFont(xlsxObject, fonts[2]),
                /*14*/getTMediumLRBThinFont(xlsxObject, fonts[3]),
                /*15*/getTRMediumLBThinFont(xlsxObject, fonts[3]),
                /*16*/getRMediumLTBThinFont(xlsxObject, fonts[3]),
                /*17*/getBMediumLTRThinFont(xlsxObject, fonts[3]),
                /*18*/getRBMediumLTThinFont(xlsxObject, fonts[3]),
                /*19*/getTRBMediumLThinFont(xlsxObject, fonts[3]),
                /*20*/getTBMediumLRThinFont(xlsxObject, fonts[3]),
                /*21*/getLTMediumRBThin(xlsxObject)
        };
    }

    static XSSFCellStyle[] getHandlingTrafficAccidentReportAdditionalDocumentsStyles(XLSXObject xlsxObject){
        return new XSSFCellStyle[]{
                /* 0*/setAlignmentByParam(xlsxObject, fonts[1], Alignment.RIGHT, BorderStyle.MEDIUM),
                /* 1*/getLTBMediumRThinFont(xlsxObject, fonts[2]),
                /* 2*/getTRBMediumLThinFont(xlsxObject, fonts[2]),
                /* 3*/getLTBMediumRThinFont(xlsxObject, fonts[3]),
                /* 4*/getTRBMediumLThinFont(xlsxObject, fonts[3]),
                /* 5*/getLTMediumRBThinFont(xlsxObject, fonts[3]),
                /* 6*/getTRMediumLBThinFont(xlsxObject, fonts[3]),
                /* 7*/getLMediumTRBThinFont(xlsxObject, fonts[3]),
                /* 8*/getRMediumLTBThinFont(xlsxObject, fonts[3]),
                /* 9*/getLBMediumTRThinFont(xlsxObject, fonts[3]),
                /*10*/getRBMediumLTThinFont(xlsxObject, fonts[3])
        };
    }

    static XSSFCellStyle[] getFailesSafetyReportAdditionalDocumentsStyles(XLSXObject xlsxObject){
        return new XSSFCellStyle[]{
                /* 0*/setAlignmentByParam(xlsxObject, fonts[1], Alignment.RIGHT, BorderStyle.MEDIUM),
                /* 1*/getLTBMediumRThinFont(xlsxObject, fonts[2]),
                /* 2*/getTRBMediumLThinFont(xlsxObject, fonts[2]),
                /* 3*/getLTBMediumRThinFont(xlsxObject, fonts[3]),
                /* 4*/getTRBMediumLThinFont(xlsxObject, fonts[3]),
                /* 5*/getLTMediumRBThinFont(xlsxObject, fonts[3]),
                /* 6*/getTRMediumLBThinFont(xlsxObject, fonts[3]),
                /* 7*/getLMediumTRBThinFont(xlsxObject, fonts[3]),
                /* 8*/getRMediumLTBThinFont(xlsxObject, fonts[3]),
                /* 9*/getLBMediumTRThinFont(xlsxObject, fonts[3]),
                /*10*/getRBMediumLTThinFont(xlsxObject, fonts[3])
        };
    }

    static XSSFCellStyle[] getTreatmentOfImpairedByMonitoringReportStylesAdditionalDocumentsStyles(XLSXObject xlsxObject){
        return new XSSFCellStyle[]{
                /* 0*/setAlignmentByParam(xlsxObject, fonts[1], Alignment.RIGHT, BorderStyle.MEDIUM),
                /* 1*/getLTBMediumRThinFont(xlsxObject, fonts[2]),
                /* 2*/getTRBMediumLThinFont(xlsxObject, fonts[2]),
                /* 3*/getLTBMediumRThinFont(xlsxObject, fonts[3]),
                /* 4*/getTRBMediumLThinFont(xlsxObject, fonts[3]),
                /* 5*/getLTMediumRBThinFont(xlsxObject, fonts[3]),
                /* 6*/getTRMediumLBThinFont(xlsxObject, fonts[3]),
                /* 7*/getLMediumTRBThinFont(xlsxObject, fonts[3]),
                /* 8*/getRMediumLTBThinFont(xlsxObject, fonts[3]),
                /* 9*/getLBMediumTRThinFont(xlsxObject, fonts[3]),
                /*10*/getRBMediumLTThinFont(xlsxObject, fonts[3])
        };
    }

    static XSSFCellStyle[] getHandlingTrafficAccidentSummaryReportStyles(XLSXObject xlsxObject){
        return new XSSFCellStyle[]{
                /* 0*/setAlignmentAndColorByParam(xlsxObject, fonts[2], HorizontalAlignment.CENTER, VerticalAlignment.CENTER, BorderStyle.THIN, IndexedColors.LIGHT_YELLOW),
                /* 1*/setAlignmentByParam(xlsxObject, fonts[0], VerticalAlignment.BOTTOM, HorizontalAlignment.CENTER, BorderStyle.NONE),
                /* 2*/getLTMediumRBThinFont(xlsxObject, fonts[2]),
                /* 3*/getTMediumLRBThinFont(xlsxObject, fonts[2]),
                /* 4*/getLBMediumTRThinFont(xlsxObject, fonts[2]),
                /* 5*/getBMediumLTRThinFont(xlsxObject, fonts[2]),
                /* 6*/getTMediumLRBThinFont(xlsxObject, fonts[3]),
                /* 7*/getTRMediumLBThinFont(xlsxObject, fonts[3]),
                /* 8*/getBMediumLTRThinFont(xlsxObject, fonts[3]),
                /* 9*/getRBMediumLTThinFont(xlsxObject, fonts[3]),
                /*10*/setAllBordersThin(xlsxObject, fonts[3])
        };
    }

    static XSSFCellStyle[] getFailesSafetySummaryReportStyles(XLSXObject xlsxObject){
        return new XSSFCellStyle[]{
                /* 0*/setAlignmentAndColorByParam(xlsxObject, fonts[2], HorizontalAlignment.CENTER, VerticalAlignment.CENTER, BorderStyle.THIN, IndexedColors.LIGHT_YELLOW),
                /* 1*/setAlignmentByParam(xlsxObject, fonts[0], VerticalAlignment.BOTTOM, HorizontalAlignment.CENTER, BorderStyle.NONE),
                /* 2*/getLTMediumRBThinFont(xlsxObject, fonts[2]),
                /* 3*/getTMediumLRBThinFont(xlsxObject, fonts[2]),
                /* 4*/getLBMediumTRThinFont(xlsxObject, fonts[2]),
                /* 5*/getBMediumLTRThinFont(xlsxObject, fonts[2]),
                /* 6*/getTMediumLRBThinFont(xlsxObject, fonts[3]),
                /* 7*/getTRMediumLBThinFont(xlsxObject, fonts[3]),
                /* 8*/getBMediumLTRThinFont(xlsxObject, fonts[3]),
                /* 9*/getRBMediumLTThinFont(xlsxObject, fonts[3]),
                /*10*/setAllBordersThin(xlsxObject, fonts[3])
        };
    }

    static XSSFCellStyle[] getSafetySupervisoryTreatmentAfterMonitoringSummaryReportStyles(XLSXObject xlsxObject) {
        return new XSSFCellStyle[] {
                /* 0*/setAlignmentAndColorByParam(xlsxObject, fonts[2], HorizontalAlignment.CENTER, VerticalAlignment.CENTER, BorderStyle.THIN, IndexedColors.LIGHT_YELLOW),
                /* 1*/setAlignmentByParam(xlsxObject, fonts[0], VerticalAlignment.BOTTOM, HorizontalAlignment.CENTER, BorderStyle.NONE),
                /* 2*/getLTMediumRBThinFont(xlsxObject, fonts[2]),
                /* 3*/getTMediumLRBThinFont(xlsxObject, fonts[2]),
                /* 4*/getLBMediumTRThinFont(xlsxObject, fonts[2]),
                /* 5*/getBMediumLTRThinFont(xlsxObject, fonts[2]),
                /* 6*/getTMediumLRBThinFont(xlsxObject, fonts[3]),
                /* 7*/getTRMediumLBThinFont(xlsxObject, fonts[3]),
                /* 8*/getLBMediumTRThinFont(xlsxObject, fonts[3]),
                /* 9*/getRBMediumLTThinFont(xlsxObject, fonts[3]),
                /*10*/setAllBordersThin(xlsxObject, fonts[3])
        };
    }
}