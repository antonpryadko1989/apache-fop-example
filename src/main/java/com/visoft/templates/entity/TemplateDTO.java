package com.visoft.templates.entity;

public class TemplateDTO {
    private String templateName;
    private String projectId;
    private String outPutName;
    private TemplateBody body;

    public TemplateDTO() {
    }

    public TemplateDTO(String templateName, String projectId, String outPutName, TemplateBody body) {
        this.templateName = templateName;
        this.projectId = projectId;
        this.outPutName = outPutName;
        this.body = body;

    }

    public String getTemplateName() {
        return templateName;
    }

    public void setTemplateName(String templateName) {
        this.templateName = templateName;
    }

    public String getProjectId() {
        return projectId;
    }

    public void setProjectId(String projectId) {
        this.projectId = projectId;
    }

    public String getOutPutName() {
        return outPutName;
    }

    public void setOutPutName(String outPutName) {
        this.outPutName = outPutName;
    }

    public void setBody(TemplateBody body) {
        this.body = body;
    }

    public TemplateBody getBody() {
        return body;

    }

    @Override
    public String toString() {
        return "TemplateDTO{" +
                "templateName='" + templateName + '\'' +
                ", projectId='" + projectId + '\'' +
                ", outPutName='" + outPutName + '\'' +
                ", body=" + body +
                '}';
    }
}