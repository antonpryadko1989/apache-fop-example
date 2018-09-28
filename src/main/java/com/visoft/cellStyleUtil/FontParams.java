package com.visoft.cellStyleUtil;

public class FontParams {
    private String fontName;
    private boolean isBold;
    private short fontHeight;

    private FontParams(String fontName, boolean isBold, short fontHeight) {
        this.fontName = fontName;
        this.isBold = isBold;
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


}
