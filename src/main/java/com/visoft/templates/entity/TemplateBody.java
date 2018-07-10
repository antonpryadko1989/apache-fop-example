package com.visoft.templates.entity;

import java.util.List;
import java.util.Map;

public class TemplateBody {

    private Map<String, String> bodyElements;
    private Map<String, List<Map<String, String>>> bodyLists;

    public TemplateBody() {
    }

    public TemplateBody(Map<String, String> bodyElements, Map<String, List<Map<String, String>>> bodyLists) {
        this.bodyElements = bodyElements;
        this.bodyLists = bodyLists;
    }

    public Map<String, String> getBodyElements() {
        return bodyElements;
    }

    public void setBodyElements(Map<String, String> bodyElements) {
        this.bodyElements = bodyElements;
    }

    public Map<String, List<Map<String, String>>> getBodyLists() {
        return bodyLists;
    }

    public void setBodyLists(Map<String, List<Map<String, String>>> bodyLists) {
        this.bodyLists = bodyLists;
    }
}
