package com.visoft.templates.entity;

import java.util.Map;

public class TemplateDTO {
    private String templateName;
    private String projectId;
    private String outPutName;
    private Map<String, String> templateBody;
    private Object body;

    public TemplateDTO() {
    }

    public TemplateDTO(String templateName, String projectId, String outPutName, Map<String, String> templateBody, Object body){
        this.templateName = templateName;
        this.projectId = projectId;
        this.outPutName = outPutName;
        this.templateBody = templateBody;
        this.body = body;
    }

    public String getTemplateName() {
        return templateName;
    }

    public void setTemplateName(String templateName) {
        this.templateName = templateName;
    }

    public void setProjectId(String projectId) {
        this.projectId = projectId;
    }

    public void setOutPutName(String outPutName) {
        this.outPutName = outPutName;
    }

    public void setTemplateBody(Map<String, String> templateBody) {
        this.templateBody = templateBody;
    }

    public void setBody(Object body) {
        this.body = body;
    }

    public String getProjectId() {
        return projectId;
    }

    public String getOutPutName() {
        return outPutName;
    }

    public Map<String, String> getTemplateBody() {
        return templateBody;
    }

    public Object getBody() {
        return body;
    }
}
