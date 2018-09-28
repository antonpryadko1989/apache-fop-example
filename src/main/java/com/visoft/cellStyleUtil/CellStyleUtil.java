package com.visoft.cellStyleUtil;

import com.visoft.services.Alignment;
import com.visoft.templates.entity.XLSXObject;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import static org.apache.poi.ss.usermodel.BorderStyle.*;

public class CellStyleUtil {

    private static Font createFont(XLSXObject xlsxObject, FontParams fontParams) {
        Font font = xlsxObject.getSheet().getWorkbook().createFont();
        font.setFontName(fontParams.getFontName());
        font.setBold(fontParams.isBold());
        font.setFontHeight(fontParams.getFontHeight());
        return font;
    }

    public static XSSFCellStyle  setBordersAndFont(XLSXObject xlsxObject,
                                                   FontParams fontParams,
                                                   BorderStyle left,
                                                   BorderStyle top,
                                                   BorderStyle right,
                                                   BorderStyle bottom) {
        XSSFCellStyle style = (XSSFCellStyle) xlsxObject.getSheet()
                .getWorkbook().createCellStyle();
        style.setFont(createFont(xlsxObject, fontParams));
        style.setWrapText(true);
        style.setBorderLeft(left);
        style.setBorderTop(top);
        style.setBorderRight(right);
        style.setBorderBottom(bottom);
        setAlignments(style);
        return style;
    }


    public static XSSFCellStyle  setBordersAndFont(XLSXObject xlsxObject,
                                                   FontParams fontParams,
                                                   BorderStyle left,
                                                   BorderStyle top,
                                                   BorderStyle right,
                                                   BorderStyle bottom,
                                                   VerticalAlignment
                                                           verticalAlignment,
                                                   HorizontalAlignment
                                                           horizontalAlignment) {
        XSSFCellStyle style = (XSSFCellStyle) xlsxObject.getSheet()
                .getWorkbook().createCellStyle();
        style.setFont(createFont(xlsxObject, fontParams));
        style.setWrapText(true);
        style.setBorderLeft(left);
        style.setBorderTop(top);
        style.setBorderRight(right);
        style.setBorderBottom(bottom);
        setAlignments(style, verticalAlignment, horizontalAlignment);
        return style;
    }

    public static XSSFCellStyle  setAllBordersThin(XLSXObject xlsxObject,
                                                   FontParams fontParams) {
        return setBordersAndFont(xlsxObject, fontParams,
                BorderStyle.THIN, BorderStyle.THIN,
                BorderStyle.THIN, BorderStyle.THIN);
    }


    public static XSSFCellStyle  setAllBordersByStyle(XLSXObject xlsxObject,
                                               FontParams fontParams,
                                               BorderStyle borderStyle) {
        return setBordersAndFont(xlsxObject, fontParams,
                borderStyle, borderStyle,
                borderStyle, borderStyle);
    }

    public static XSSFCellStyle  setAllBordersByStyle(XLSXObject xlsxObject,
                                                      FontParams fontParams,
                                                      BorderStyle borderStyle,
                                                      VerticalAlignment
                                                              verticalAlignment,
                                                      HorizontalAlignment
                                                              horizontalAlignment){
        return setBordersAndFont(xlsxObject, fontParams,
                borderStyle, borderStyle,
                borderStyle, borderStyle,
                verticalAlignment, horizontalAlignment);
    }

    public static XSSFCellStyle setRightBorderMedium(XLSXObject xlsxObject){
        XSSFCellStyle style = (XSSFCellStyle) xlsxObject.getSheet().getWorkbook().createCellStyle();
        style.setBorderRight(MEDIUM);
        return style;
    }

    public static XSSFCellStyle setAlignmentByParam(XLSXObject xlsxObject, FontParams fontParams, Alignment alignment) {
        XSSFCellStyle style = (XSSFCellStyle) xlsxObject.getSheet().getWorkbook().createCellStyle();
        style.setFont(createFont(xlsxObject, fontParams));
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        if(alignment.equals(Alignment.LEFT)){
            style.setAlignment(HorizontalAlignment.LEFT);
        }
        if(alignment.equals(Alignment.CENTER)){
            style.setAlignment(HorizontalAlignment.CENTER);
        }
        if(alignment.equals(Alignment.RIGHT)){
            style.setAlignment(HorizontalAlignment.RIGHT);
        }
        return style;
    }

    public static CellStyle setBordersLeftRightBottomThin(XLSXObject xlsxObject, FontParams fontParams) {
        XSSFCellStyle style = (XSSFCellStyle) xlsxObject.getSheet().getWorkbook().createCellStyle();
        style.setFont(createFont(xlsxObject, fontParams));
        style.setWrapText(true);
        style.setBorderBottom(THIN);
        style.setBorderLeft(THIN);
        style.setBorderRight(THIN);
        setAlignments(style);
        return style;
    }

    private static void
    setAlignments (XSSFCellStyle style){
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
    }

    private static void
    setAlignments (XSSFCellStyle style, VerticalAlignment verticalAlignment,
                   HorizontalAlignment alignment){
        style.setAlignment(alignment);
        style.setVerticalAlignment(verticalAlignment);
    }

    private static XSSFCellStyle
    setAlignmentsCenterBottom (XSSFCellStyle style){
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.BOTTOM);
        return style;
    }
}
