package com.visoft.templates.regulars;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class TemplateValue {
    private static final String translationKeyValue = "#([^#]*)#";

    private static Pattern hashTagEnclosedPattern = Pattern.compile(translationKeyValue);

    public static String isValue (String value) {
        Matcher matcher = hashTagEnclosedPattern.matcher(value);
        if (matcher.find()) {
            return  matcher.group(1);
        }
        else return "";
    }
}
