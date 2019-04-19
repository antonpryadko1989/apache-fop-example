package com.visoft.dto;

import java.util.List;

public class DeficienciesReportDTO {

    private String reportName;
    private String outputName;
    private ReportBody reportBody;
    private List<byte[]> docs;

    public DeficienciesReportDTO() {
    }

    public DeficienciesReportDTO(String reportName, String outputName, ReportBody reportBody) {
        this.reportName = reportName;
        this.outputName = outputName;
        this.reportBody = reportBody;
    }

    public DeficienciesReportDTO(String reportName, String outputName, ReportBody reportBody, List<byte[]> docs) {
        this.reportName = reportName;
        this.outputName = outputName;
        this.reportBody = reportBody;
        this.docs = docs;

    }

    public String getReportName() {
        return reportName;
    }

    public void setReportName(String reportName) {
        this.reportName = reportName;
    }

    public String getOutputName() {
        return outputName;
    }

    public void setOutputName(String outputName) {
        this.outputName = outputName;
    }

    public ReportBody getReportBody() {
        return reportBody;
    }

    public void setReportBody(ReportBody reportBody) {
        this.reportBody = reportBody;
    }

    public List<byte[]> getDocs() {
        return docs;
    }

    public void setDocs(List<byte[]> docs) {
        this.docs = docs;
    }

    @Override
    public String toString() {
        return "DeficienciesReportDTO{" +
                "reportName='" + reportName + '\'' +
                ", outputName='" + outputName + '\'' +
                ", reportBody=" + reportBody +
                ", docs=" + docs +
                '}';
    }

    public boolean reportBodyListNotEmpty(DeficienciesReportDTO reportDTO) {
        boolean result;
        try {
            result = reportDTO.getReportBody().getBodyLists() != null;
        }catch (NullPointerException e){
            result = false;
        }
        return result;
    }
}
