package com.visoft.dto;

import java.util.List;
import java.util.Map;

public class SummaryReportDTO {
    private String name;
    private Map<String, String> headers;
    private List<Map<String, String>> values;

    public SummaryReportDTO() {
    }

    public SummaryReportDTO(String name, Map<String, String> headers, List<Map<String, String>> values) {
        this.name = name;
        this.headers= headers;
        this.values = values;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Map<String, String> getHeaders() {
        return headers;
    }

    public void setHeaders(Map<String, String> headers) {
        this.headers = headers;
    }

    public List<Map<String, String>> getValues() {
        return values;
    }

    public void setValues(List<Map<String, String>> values) {
        this.values = values;
    }

    @Override
    public String toString() {
        return "SummaryReportDTO{" +
                "name='" + name + '\'' +
                ", headers=" + headers +
                ", values=" + values +
                '}';
    }
}
