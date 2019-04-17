package com.visoft.cellStyleUtil;

public class FontParams {
    private String fontName;
    private boolean isBold;
    private boolean underline;
    private short fontHeight;

    private FontParams(String fontName, boolean isBold, short fontHeight) {
        this.fontName = fontName;
        this.isBold = isBold;
        this.fontHeight = fontHeight;
    }

    private FontParams(String fontName, boolean isBold, boolean underline, short fontHeight) {
        this.fontName = fontName;
        this.isBold = isBold;
        this.underline = underline;
        this.fontHeight = fontHeight;
    }

    public String getFontName() {
        return fontName;
    }

    public void setFontName(String fontName) {
        this.fontName = fontName;
    }

    public boolean isBold() {
        return isBold;
    }

    public boolean isUnderline() {
        return underline;
    }

    public void setUnderline(boolean underline) {
        this.underline = underline;
    }

    public void setBold(boolean bold) {
        isBold = bold;
    }

    public short getFontHeight() {
        return fontHeight;
    }

    public void setFontHeight(short fontHeight) {
        this.fontHeight = fontHeight;
    }

    public static FontParams setFontParams(String fontName, boolean isBold, short fontHeight) {
        return new FontParams(fontName, isBold, fontHeight);
    }

    public static FontParams setFontParams(String fontName, boolean isBold, boolean underscore, short fontHeight) {
        return new FontParams(fontName, isBold, underscore, fontHeight);
    }


}
