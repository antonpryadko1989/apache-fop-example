package com.visoft.dto;

import java.util.List;
import java.util.Map;

public class ReportBody {
    private Map<String, String> bodyElements;
    private Map<String, List<Map<String, String>>> bodyLists;
    private ReportLogo reportLogo;

    public ReportBody() {
    }

    public ReportBody(Map<String, String> bodyElements, Map<String, List<Map<String, String>>> bodyLists, ReportLogo reportLogo) {
        this.bodyElements = bodyElements;
        this.bodyLists = bodyLists;
        this.reportLogo = reportLogo;
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

    public ReportLogo getReportLogo() {
        return reportLogo;
    }

    public void setReportLogo(ReportLogo reportLogo) {
        this.reportLogo = reportLogo;
    }

    @Override
    public String toString() {
        return "ReportBody{" +
                "bodyElements=" + bodyElements +
                ", bodyLists=" + bodyLists +
                ", reportLogo=" + reportLogo +
                '}';
    }
}
