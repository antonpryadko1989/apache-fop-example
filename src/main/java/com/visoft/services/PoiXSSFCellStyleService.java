package com.visoft.services;

import com.visoft.cellStyleUtil.FontParams;
import com.visoft.templates.entity.XLSXObject;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import static com.visoft.cellStyleUtil.CellStyleUtil.*;
import static com.visoft.cellStyleUtil.FontParams.setFontParams;
import static com.visoft.utils.Const.FONT_NAME_CALIBRI;
import static com.visoft.utils.Const.FONT_RATE;

public interface PoiXSSFCellStyleService {

    FontParams fontCalibri20bold = setFontParams(FONT_NAME_CALIBRI, true, (short) (20 * FONT_RATE));
    FontParams fontCalibri12bold = setFontParams(FONT_NAME_CALIBRI, true, (short) (12 * FONT_RATE));
    FontParams fontCalibri12 = setFontParams(FONT_NAME_CALIBRI, false, (short) (12 * FONT_RATE));

    static XSSFCellStyle[] getNCRReportStyles(XLSXObject xlsxObject) {
        return new XSSFCellStyle[]{
                /* 0*/setAlignmentByParam(xlsxObject, fontCalibri20bold, Alignment.CENTER, BorderStyle.MEDIUM),
                /* 1*/setAlignmentByParam(xlsxObject, fontCalibri20bold, Alignment.RIGHT, BorderStyle.MEDIUM),
                /* 2*/getLTMediumRBThinFont(xlsxObject, fontCalibri12bold),
                /* 3*/getTMediumLRBThinFont(xlsxObject, fontCalibri12bold),
                /* 4*/getTRMediumLBThinFont(xlsxObject, fontCalibri12bold),
                /* 5*/getBMediumLTRThinFont(xlsxObject, fontCalibri12bold),
                /* 6*/getRBMediumLTThinFont(xlsxObject, fontCalibri12bold),
                /* 7*/getLBMediumTRThinFont(xlsxObject, fontCalibri12bold),
                /* 8*/getLTBMediumRThinFont(xlsxObject, fontCalibri12bold),
                /* 9*/getTBMediumLRThinFont(xlsxObject, fontCalibri12bold),
                /*10*/getTRBMediumLThinFont(xlsxObject, fontCalibri12bold),
                /*11*/setAllBordersThin(xlsxObject, fontCalibri12bold),
                /*12*/getLMediumTRBThinFont(xlsxObject, fontCalibri12bold),
                /*13*/getLBMediumRTThinFont(xlsxObject, fontCalibri12bold),
                /*14*/getTMediumLRBThinFont(xlsxObject, fontCalibri12),
                /*15*/getTRMediumLBThinFont(xlsxObject, fontCalibri12),
                /*16*/setAllBordersThin(xlsxObject, fontCalibri12),
                /*17*/getRMediumLTBThinFont(xlsxObject, fontCalibri12),
                /*18*/getBMediumLTRThinFont(xlsxObject, fontCalibri12),
                /*19*/getRBMediumLTThinFont(xlsxObject, fontCalibri12),
                /*20*/getLMediumTRBThinFont(xlsxObject, fontCalibri12),
                /*21*/getTBMediumLRThinFont(xlsxObject, fontCalibri12),
                /*22*/getTRBMediumLThinFont(xlsxObject, fontCalibri12)
        };
    }

    static XSSFCellStyle[] getPOCReportStyles(XLSXObject xlsxObject) {
        return new XSSFCellStyle[]{
                /* 0*/setAlignmentByParam(xlsxObject, fontCalibri20bold, Alignment.RIGHT, BorderStyle.MEDIUM),
                /* 1*/getLTMediumRBThinFont(xlsxObject, fontCalibri12bold),
                /* 2*/getTMediumLRBThinFont(xlsxObject, fontCalibri12bold),
                /* 3*/getTRMediumLBThinFont(xlsxObject, fontCalibri12bold),
                /* 4*/getBMediumLTRThinFont(xlsxObject, fontCalibri12bold),
                /* 5*/getRBMediumLTThinFont(xlsxObject, fontCalibri12bold),
                /* 6*/getLBMediumTRThinFont(xlsxObject, fontCalibri12bold),
                /* 7*/getLTBMediumRThinFont(xlsxObject, fontCalibri12bold),
                /* 8*/getTBMediumLRThinFont(xlsxObject, fontCalibri12bold),
                /* 9*/getTRBMediumLThinFont(xlsxObject, fontCalibri12bold),
                /*10*/setAllBordersThin(xlsxObject, fontCalibri12bold),
                /*11*/getLMediumTRBThinFont(xlsxObject, fontCalibri12bold),
                /*12*/getLBMediumRTThinFont(xlsxObject, fontCalibri12bold),
                /*13*/getTMediumLRBThinFont(xlsxObject, fontCalibri12),
                /*14*/getTRMediumLBThinFont(xlsxObject, fontCalibri12),
                /*15*/setAllBordersThin(xlsxObject, fontCalibri12),
                /*16*/getRMediumLTBThinFont(xlsxObject, fontCalibri12),
                /*17*/getBMediumLTRThinFont(xlsxObject, fontCalibri12),
                /*18*/getRBMediumLTThinFont(xlsxObject, fontCalibri12),
                /*19*/getLTMediumRBThinFont(xlsxObject, fontCalibri12),
                /*20*/getLMediumTRBThinFont(xlsxObject, fontCalibri12),
                /*21*/getLBMediumRTThinFont(xlsxObject, fontCalibri12),
                /*22*/getLTBMediumRThinFont(xlsxObject, fontCalibri12),
                /*23*/getTBMediumLRThinFont(xlsxObject, fontCalibri12),
                /*24*/getTRBMediumLThinFont(xlsxObject, fontCalibri12)
        };
    }

    static XSSFCellStyle[] getChecklistReportStyles(XLSXObject xlsxObject) {
        return new XSSFCellStyle[]{
                /* 0*/getLTMediumRBThinFont(xlsxObject, fontCalibri12bold),
                /* 1*/getTMediumLRBThinFont(xlsxObject, fontCalibri12bold),
                /* 2*/getTRMediumLBThinFont(xlsxObject, fontCalibri12bold),
                /* 3*/getBMediumLTRThinFont(xlsxObject, fontCalibri12bold),
                /* 4*/getRBMediumLTThinFont(xlsxObject, fontCalibri12bold),
                /* 5*/getLBMediumTRThinFont(xlsxObject, fontCalibri12bold),
                /* 6*/getLTBMediumRThinFont(xlsxObject, fontCalibri12bold),
                /* 7*/getTBMediumLRThinFont(xlsxObject, fontCalibri12bold),
                /* 8*/getTRBMediumLThinFont(xlsxObject, fontCalibri12bold),
                /* 9*/setAllBordersThin(xlsxObject, fontCalibri12bold),
                /*10*/getLMediumTRBThinFont(xlsxObject, fontCalibri12bold),
                /*11*/getTMediumLRBThinFont(xlsxObject, fontCalibri12),
                /*12*/getTRMediumLBThinFont(xlsxObject, fontCalibri12),
                /*13*/setAllBordersThin(xlsxObject, fontCalibri12),
                /*14*/getRMediumLTBThinFont(xlsxObject, fontCalibri12),
                /*15*/getBMediumLTRThinFont(xlsxObject, fontCalibri12),
                /*16*/getRBMediumLTThinFont(xlsxObject, fontCalibri12),
                /*17*/getLTMediumRBThinFont(xlsxObject, fontCalibri12),
                /*18*/getLMediumTRBThinFont(xlsxObject, fontCalibri12),
                /*19*/getLBMediumRTThinFont(xlsxObject, fontCalibri12),
                /*20*/getLTBMediumRThinFont(xlsxObject, fontCalibri12),
                /*21*/getTBMediumLRThinFont(xlsxObject, fontCalibri12),
                /*22*/getTRBMediumLThinFont(xlsxObject, fontCalibri12)
        };
    }

    static XSSFCellStyle[] getPreliminaryMaterialsInspectionReportStyles(XLSXObject xlsxObject) {
        return new XSSFCellStyle[]{
                /* 0*/setAlignmentByParam(xlsxObject, fontCalibri20bold, Alignment.RIGHT, BorderStyle.MEDIUM),
                /* 1*/getLTMediumRBThinFont(xlsxObject, fontCalibri12bold),
                /* 2*/getTMediumLRBThinFont(xlsxObject, fontCalibri12bold),
                /* 3*/getTRMediumLBThinFont(xlsxObject, fontCalibri12bold),
                /* 4*/getBMediumLTRThinFont(xlsxObject, fontCalibri12bold),
                /* 5*/getLBMediumRTThinFont(xlsxObject, fontCalibri12bold),
                /* 6*/getRBMediumLTThinFont(xlsxObject, fontCalibri12bold),
                /* 7*/getLMediumTRBThinFont(xlsxObject, fontCalibri12bold),
                /* 8*/setAllBordersThin(xlsxObject, fontCalibri12bold),
                /* 9*/getLBMediumTRThinFont(xlsxObject, fontCalibri12bold),
                /*10*/getLTBMediumRThinFont(xlsxObject, fontCalibri12bold),
                /*11*/getTBMediumLRThinFont(xlsxObject, fontCalibri12bold),
                /*12*/getTRBMediumLThinFont(xlsxObject, fontCalibri12bold),
                /*13*/getTRBMediumLThinFont(xlsxObject, fontCalibri12bold),
                /*14*/getTBMediumLRThinFont(xlsxObject, fontCalibri12),
                /*15*/getTRBMediumLThinFont(xlsxObject, fontCalibri12),
                /*16*/getRBMediumLTThinFont(xlsxObject, fontCalibri12),
                /*17*/getTMediumLRBThinFont(xlsxObject, fontCalibri12),
                /*18*/getTRMediumLBThinFont(xlsxObject, fontCalibri12),
                /*19*/getBMediumLTRThinFont(xlsxObject, fontCalibri12),
                /*20*/getRMediumLTBThinFont(xlsxObject, fontCalibri12),
                /*21*/setAllBordersThin(xlsxObject, fontCalibri12)
        };
    }

    static XSSFCellStyle[] getApprovalOfSuppliersReportStyles(XLSXObject xlsxObject) {
        return new XSSFCellStyle[]{
                /* 0*/setAlignmentByParam(xlsxObject, fontCalibri20bold, Alignment.RIGHT, BorderStyle.MEDIUM),
                /* 1*/getLTMediumRBThinFont(xlsxObject, fontCalibri12bold),
                /* 2*/getTMediumLRBThinFont(xlsxObject, fontCalibri12bold),
                /* 3*/getTRMediumLBThinFont(xlsxObject, fontCalibri12bold),
                /* 4*/getBMediumLTRThinFont(xlsxObject, fontCalibri12bold),
                /* 5*/getRBMediumLTThinFont(xlsxObject, fontCalibri12bold),
                /* 6*/getLBMediumTRThinFont(xlsxObject, fontCalibri12bold),
                /* 7*/getLTBMediumRThinFont(xlsxObject, fontCalibri12bold),
                /* 8*/getTBMediumLRThinFont(xlsxObject, fontCalibri12bold),
                /* 9*/getTRBMediumLThinFont(xlsxObject, fontCalibri12bold),
                /*10*/setAllBordersThin(xlsxObject, fontCalibri12bold),
                /*11*/getLMediumTRBThinFont(xlsxObject, fontCalibri12bold),
                /*12*/getLBMediumRTThinFont(xlsxObject, fontCalibri12bold),
                /*13*/getTMediumLRBThinFont(xlsxObject, fontCalibri12),
                /*14*/getTRMediumLBThinFont(xlsxObject, fontCalibri12),
                /*15*/setAllBordersThin(xlsxObject, fontCalibri12),
                /*16*/getRMediumLTBThinFont(xlsxObject, fontCalibri12),
                /*17*/getBMediumLTRThinFont(xlsxObject, fontCalibri12),
                /*18*/getRBMediumLTThinFont(xlsxObject, fontCalibri12),
                /*19*/getTBMediumLRThinFont(xlsxObject, fontCalibri12),
                /*20*/getTRBMediumLThinFont(xlsxObject, fontCalibri12)
        };
    }

    static XSSFCellStyle[] getApprovalOfSubcontractorsReportStyles(XLSXObject xlsxObject) {
        return new XSSFCellStyle[]{
                /* 0*/setAlignmentByParam(xlsxObject, fontCalibri20bold, Alignment.RIGHT, BorderStyle.MEDIUM),
                /* 1*/getLTMediumRBThinFont(xlsxObject, fontCalibri12bold),
                /* 2*/getTMediumLRBThinFont(xlsxObject, fontCalibri12bold),
                /* 3*/getTRMediumLBThinFont(xlsxObject, fontCalibri12bold),
                /* 4*/getBMediumLTRThinFont(xlsxObject, fontCalibri12bold),
                /* 5*/getRBMediumLTThinFont(xlsxObject, fontCalibri12bold),
                /* 6*/getLBMediumTRThinFont(xlsxObject, fontCalibri12bold),
                /* 7*/getLTBMediumRThinFont(xlsxObject, fontCalibri12bold),
                /* 8*/getTBMediumLRThinFont(xlsxObject, fontCalibri12bold),
                /* 9*/getTRBMediumLThinFont(xlsxObject, fontCalibri12bold),
                /*10*/setAllBordersThin(xlsxObject, fontCalibri12bold),
                /*11*/getLMediumTRBThinFont(xlsxObject, fontCalibri12bold),
                /*12*/getLBMediumRTThinFont(xlsxObject, fontCalibri12bold),
                /*13*/getTMediumLRBThinFont(xlsxObject, fontCalibri12),
                /*14*/getTRMediumLBThinFont(xlsxObject, fontCalibri12),
                /*15*/setAllBordersThin(xlsxObject, fontCalibri12),
                /*16*/getRMediumLTBThinFont(xlsxObject, fontCalibri12),
                /*17*/getBMediumLTRThinFont(xlsxObject, fontCalibri12),
                /*18*/getRBMediumLTThinFont(xlsxObject, fontCalibri12),
                /*19*/getTBMediumLRThinFont(xlsxObject, fontCalibri12),
                /*20*/getTRBMediumLThinFont(xlsxObject, fontCalibri12)
        };
    }
}
