package com.visoft.templates.entity;

import java.util.List;
import java.util.Map;

public class Body {

    private Map<String, String> bodyElements;
    private Map<String, List<Map<String, String>>> bodyLists;
    private Logo logo;

    @Override
    public String toString() {
        return "TemplateBody{" +
                "bodyElements=" + bodyElements +
                ", bodyLists=" + bodyLists +
                ", logo=" + logo +
                '}';
    }

    public Logo getLogo() {
        return logo;
    }

    public void setLogo(Logo logo) {
        this.logo = logo;
    }

    public Body() {
    }

    public Body(Map<String, String> bodyElements,
                Map<String, List<Map<String, String>>> bodyLists,
                Logo logo) {
        this.bodyElements = bodyElements;
        this.bodyLists = bodyLists;
        this.logo = logo;
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
