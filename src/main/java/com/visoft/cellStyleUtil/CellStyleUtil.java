package com.visoft.cellStyleUtil;

import com.visoft.services.Alignment;
import com.visoft.templates.entity.XLSXObject;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import static org.apache.poi.ss.usermodel.BorderStyle.MEDIUM;
import static org.apache.poi.ss.usermodel.BorderStyle.THIN;

public class CellStyleUtil {

    private static Font createFont(XLSXObject xlsxObject, FontParams fontParams) {
        Font font = xlsxObject.getSheet().getWorkbook().createFont();
        font.setFontName(fontParams.getFontName());
        font.setBold(fontParams.isBold());
        font.setFontHeight(fontParams.getFontHeight());
        if(fontParams.isUnderline()){
            font.setUnderline(XSSFFont.U_SINGLE);
        }
        return font;
    }

    private static XSSFCellStyle setBordersAndFont(XLSXObject xlsxObject,
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

    private static XSSFCellStyle setBordersAndFont(XLSXObject xlsxObject,
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

    private static XSSFCellStyle setBorders(XLSXObject xlsxObject,
                                            BorderStyle left,
                                            BorderStyle top,
                                            BorderStyle right,
                                            BorderStyle bottom) {
        XSSFCellStyle style = (XSSFCellStyle) xlsxObject.getSheet()
                .getWorkbook().createCellStyle();
        style.setBorderLeft(left);
        style.setBorderTop(top);
        style.setBorderRight(right);
        style.setBorderBottom(bottom);
        return style;
    }

    public static XSSFCellStyle setAllBordersByBorderStyle (XSSFSheet sh,
                                                            BorderStyle bs) {
        XSSFCellStyle style = sh.getWorkbook().createCellStyle();
        style.setBorderLeft(bs);
        style.setBorderTop(bs);
        style.setBorderRight(bs);
        style.setBorderBottom(bs);
        return style;
    }

    public static XSSFCellStyle  setAllBordersThin(XLSXObject xlsxObject,
                                                   FontParams fontParams) {
        return setBordersAndFont(xlsxObject, fontParams,
                BorderStyle.THIN, BorderStyle.THIN,
                BorderStyle.THIN, BorderStyle.THIN);
    }


    public static XSSFCellStyle setAllBordersByStyle(XLSXObject xlsxObject,
                                               FontParams fontParams,
                                               BorderStyle borderStyle) {
        return setBordersAndFont(xlsxObject, fontParams,
                borderStyle, borderStyle,
                borderStyle, borderStyle);
    }

    public static XSSFCellStyle  setAllBordersByStyle(XLSXObject xlsxObject,
                                                      FontParams fontParams,
                                                      BorderStyle borderStyle,
                                                      VerticalAlignment verticalAlignment,
                                                      HorizontalAlignment horizontalAlignment){
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

    public static XSSFCellStyle setAlignmentByParam(XLSXObject xlsxObject,
                                                    FontParams fontParams,
                                                    Alignment alignment,
                                                    BorderStyle borderStyle) {
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
        style.setBorderLeft(borderStyle);
        style.setBorderTop(borderStyle);
        style.setBorderRight(borderStyle);
        style.setBorderBottom(borderStyle);
        return style;
    }

    public static XSSFCellStyle setAlignmentByParam(XLSXObject xlsxObject,
                                                    FontParams fontParams,
                                                    VerticalAlignment verticalAlignment,
                                                    HorizontalAlignment horizontalAlignment,
                                                    BorderStyle borderStyle) {
        XSSFCellStyle style = (XSSFCellStyle) xlsxObject.getSheet().getWorkbook().createCellStyle();
        style.setFont(createFont(xlsxObject, fontParams));
        style.setVerticalAlignment(verticalAlignment);
        style.setAlignment(horizontalAlignment);
        style.setBorderLeft(borderStyle);
        style.setBorderTop(borderStyle);
        style.setBorderRight(borderStyle);
        style.setBorderBottom(borderStyle);
        return style;
    }

    public static XSSFCellStyle setAlignmentAndColorByParam(XLSXObject xlsxObject,
                                                            FontParams fontParams,
                                                            HorizontalAlignment horizontalAlignment,
                                                            VerticalAlignment verticalAlignment,
                                                            BorderStyle borderStyle,
                                                            IndexedColors color) {
        XSSFCellStyle style = (XSSFCellStyle) xlsxObject.getSheet().getWorkbook().createCellStyle();
        style.setFont(createFont(xlsxObject, fontParams));
        style.setFillForegroundColor(color.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setVerticalAlignment(verticalAlignment);
        style.setAlignment(horizontalAlignment);
        style.setBorderLeft(borderStyle);
        style.setBorderTop(borderStyle);
        style.setBorderRight(borderStyle);
        style.setBorderBottom(borderStyle);
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

    private static XSSFCellStyle setAlignmentsCenterBottom (XSSFCellStyle style){
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.BOTTOM);
        return style;
    }

    public static XSSFCellStyle getLTMediumRBThinFont(XLSXObject xlsxObject, FontParams font) {
        return setBordersAndFont(xlsxObject, font,
                BorderStyle.MEDIUM, BorderStyle.MEDIUM,
                BorderStyle.THIN, BorderStyle.THIN);
    }

    public static XSSFCellStyle getRMediumLTBThinFont(XLSXObject xlsxObject, FontParams font) {
        return setBordersAndFont(xlsxObject, font,
                BorderStyle.THIN, BorderStyle.THIN,
                BorderStyle.MEDIUM, BorderStyle.THIN);
    }

    public static XSSFCellStyle getTMediumLRBThinFont(XLSXObject xlsxObject, FontParams font) {
        return setBordersAndFont(xlsxObject, font,
                BorderStyle.THIN, BorderStyle.MEDIUM,
                BorderStyle.THIN, BorderStyle.THIN);
    }

    public static XSSFCellStyle getTRMediumLBThinFont(XLSXObject xlsxObject, FontParams font){
        return setBordersAndFont(xlsxObject, font,
                BorderStyle.THIN, BorderStyle.MEDIUM,
                BorderStyle.MEDIUM, BorderStyle.THIN);
    }

    public static XSSFCellStyle getLBMediumTRThinFont(XLSXObject xlsxObject, FontParams font){
        return setBordersAndFont(
                xlsxObject, font,
                BorderStyle.MEDIUM, BorderStyle.THIN,
                BorderStyle.THIN, BorderStyle.MEDIUM);
    }

    public static XSSFCellStyle getBMediumLTRThinFont(XLSXObject xlsxObject, FontParams font){
        return setBordersAndFont(xlsxObject, font,
                BorderStyle.THIN, BorderStyle.THIN,
                BorderStyle.THIN, BorderStyle.MEDIUM);
    }

    public static XSSFCellStyle getRBMediumLTThinFont(XLSXObject xlsxObject, FontParams font){
        return setBordersAndFont(xlsxObject, font,
                BorderStyle.THIN, BorderStyle.THIN,
                BorderStyle.MEDIUM, BorderStyle.MEDIUM);
    }

    public static XSSFCellStyle getLTBMediumRThinFont(XLSXObject xlsxObject, FontParams font){
        return setBordersAndFont(xlsxObject, font,
                BorderStyle.MEDIUM, BorderStyle.MEDIUM,
                BorderStyle.THIN, BorderStyle.MEDIUM);
    }

    public static XSSFCellStyle getTBMediumLRThinFont(XLSXObject xlsxObject, FontParams font){
        return setBordersAndFont(xlsxObject, font,
                BorderStyle.THIN, BorderStyle.MEDIUM,
                BorderStyle.THIN, BorderStyle.MEDIUM);
    }

    public static XSSFCellStyle getTRBMediumLThinFont(XLSXObject xlsxObject, FontParams font){
        return setBordersAndFont(xlsxObject, font,
                BorderStyle.THIN, BorderStyle.MEDIUM,
                BorderStyle.MEDIUM, BorderStyle.MEDIUM);
    }

    public static XSSFCellStyle getLMediumTRBThinFont(XLSXObject xlsxObject, FontParams font) {
        return setBordersAndFont(xlsxObject, font,
                BorderStyle.MEDIUM, BorderStyle.THIN,
                BorderStyle.THIN, BorderStyle.THIN);
    }

    public static XSSFCellStyle getLBMediumRTThinFont(XLSXObject xlsxObject, FontParams font) {
        return setBordersAndFont(xlsxObject, font,
                BorderStyle.MEDIUM, BorderStyle.THIN,
                BorderStyle.THIN, BorderStyle.MEDIUM);
    }

    public static XSSFCellStyle getLTMediumRBThin(XLSXObject xlsxObject){
        return setBorders(xlsxObject,
                BorderStyle.MEDIUM, BorderStyle.MEDIUM,
                BorderStyle.THIN, BorderStyle.THIN);
    }
}
