package com.visoft.templates.entity;

import com.visoft.templates.enums.ConfigType;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static com.visoft.utils.Const.*;

public class TemplateDTO {
    private String templateName;
    private String projectId;
    private String outPutName;
    private String sheetName;
    private Body body;
    private ConfigType configType;


    public TemplateDTO() {
    }

    public TemplateDTO(String templateName,
                       String projectId,
                       String outPutName,
                       String sheetName,
                       Body body,
                       ConfigType configType) {
        this.templateName = templateName;
        this.projectId = projectId;
        this.outPutName = outPutName;
        this.sheetName = sheetName;
        this.body = body;
        this.configType = configType;
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

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public Body getBody() {
        return body;
    }

    public void setBody(Body body) {
        this.body = body;
    }

    public ConfigType getConfigType() {
        return configType;
    }

    public void setConfigType(ConfigType configType) {
        this.configType = configType;
    }

    public boolean templateBodyListNotEmpty(TemplateDTO template) {
        return template.getBody().getBodyLists() != null;
    }

    @Override
    public String toString() {
        return "TemplateDTO{" +
                "templateName='" + templateName + '\'' +
                ", projectId='" + projectId + '\'' +
                ", outPutName='" + outPutName + '\'' +
                ", sheetName='" + sheetName + '\'' +
                ", body=" + body +
                ", configType=" + configType +
                '}';
    }

    public static TemplateDTO getDefaultTemplate(String templateName) {
        TemplateDTO template = new TemplateDTO();
        template.setOutPutName(templateName);
        template.setConfigType(ConfigType.ZERO);
//        template.setConfigType(ConfigType.ONE);
//        template.setConfigType(ConfigType.TWO);
        template.setSheetName(templateName);
        template.setBody(setDefaultBody(templateName));
        return template;
    }

    private static Body setDefaultBody(String templateName) {
        Body body = new Body();
        body.setLogo(setDefaultLogo());
        switch (templateName) {
            case NCR:
                body.setBodyElements(
                        setDefaultNCRBodyElements(
                                setDefaultBodyElements()));
                break;
            case POC:
                body.setBodyElements(
                        setDefaultPOCBodyElements(
                                setDefaultBodyElements()));
                break;
            case APPROVAL_OF_SUBCONTRACTORS:
                body.setBodyElements(
                        setDefaultApprovalOfSubcontractorsBodyElements(
                                setDefaultBodyElements()));
                break;
            case APPROVAL_OF_SUPPLIERS:
                body.setBodyElements(
                        setDefaultApprovalOfSuppliersBodyElements(
                                setDefaultBodyElements()));
                break;
            case PRELIMINARY_MATERIALS_INSPECTION:
                body.setBodyElements(
                        setDefaultPreliminaryMaterialsInspectionBodyElements(
                                setDefaultBodyElements()));
                break;
            case CHECKLIST:
                body.setBodyElements(
                        setDefaultChecklistBodyElements(
                                setDefaultBodyElements()));
                break;
            default:
                break;
        }
        body.setBodyLists(setDefaultBodyLists());
        return body;
    }

    private static Logo setDefaultLogo() {
        Logo logo = new Logo();
//        logo.setCountCells(98);
        logo.setPath(LOGO_PATH);
        return logo;
    }

    private static Map<String, String> setDefaultChecklistBodyElements(
            Map<String, String> result) {
        result.putAll(fillMap(CHECKLIST_BODY));
        return result;
    }

    private static Map<String, String> setDefaultPreliminaryMaterialsInspectionBodyElements(
            Map<String, String> result) {
        result.putAll(fillMap(PRELIM_MATERIALS_INS_APPROVAL_BODY));
        return result;
    }

    private static Map<String, String>
    setDefaultApprovalOfSuppliersBodyElements(Map<String, String> result) {
        result.putAll(fillMap(APP_OF_SUPPL_APPROVAL_BODY));
        return result;
    }

    private static Map<String, String>
    setDefaultApprovalOfSubcontractorsBodyElements(Map<String, String> result) {
        result.putAll(fillMap(APP_OF_SUBCONTR_APPROVAL_BODY));
        return result;
    }

    private static Map<String, String>
    setDefaultPOCBodyElements(Map<String, String> bodyElements) {
        bodyElements.putAll(fillMap(POC_BODY));
        return bodyElements;
    }

    private static Map<String, String> setDefaultNCRBodyElements(
            Map<String, String> bodyElements) {
        bodyElements.putAll(fillMap(NCR_BODY));
        return bodyElements;
    }

    private static Map<String, String> setDefaultBodyElements() {
        return fillMap(DEF_BODY_ELEMENTS);
    }

    private static Map<String, List<Map<String, String>>> setDefaultBodyLists() {
        Map<String, List<Map<String, String>>> bodyLists = new HashMap<>();
        bodyLists.put(ADDITIONAL_DOCUMENTS_VAL, setDefaultAdditionalDocuments());
        bodyLists.put(DRAWINGS_VAL, setDefaultDrawings());
        bodyLists.put(APPROVALS_VAL, setDefaultApprovals());
        bodyLists.put(CERTIFICATIONS_VAL, setDefaultCertifications());
        bodyLists.put(PRELIMINARY_INSPECTION_RESULTS_VAL,
                setDefaultPreliminaryInspectionResults());
        bodyLists.put(CHECKLIST_ELEMENTS_VAL, setDefaulChecklistElements());
        bodyLists.put(CHECKLIST_ITEMS_VAL, setDefaultChecklistItems());
        return bodyLists;
    }

    private static List<Map<String, String>> setDefaulChecklistElements() {
        List<Map<String, String>> list = new ArrayList<>();
        for (int i = 0; i < 5; i++) {
            list.add(fillMap(CHECKLIST_ELEMENT_KEY_VALUE, i));
        }
        return list;
    }


    private static List<Map<String, String>> setDefaultApprovals() {
        List<Map<String, String>> result = new ArrayList<>();
        /*  add QA1 */
        result.add(fillApprovalInMap(
                APPROVALS_LAST_IN_LIST_ROLE_VAL,
                APPROVALS_NAME_VAL + " 1",
                APPROVALS_SIGNATURE_VAL + " 1",
                APPROVALS_DATE_VAL + " 1",
                APPROVALS_STATUS_VAL + " 1"));

        /*  add QA2 */
        result.add(fillApprovalInMap(APPROVALS_LAST_IN_LIST_ROLE_VAL,
                APPROVALS_NAME_VAL + " 2",
                APPROVALS_SIGNATURE_VAL + " 2",
                APPROVALS_DATE_VAL + " 2",
                APPROVALS_STATUS_VAL + " 2"));

        /* add QC1  */
        result.add(fillApprovalInMap(APPROVALS_FIRST_IN_LIST_ROLE_VAL,
                APPROVALS_NAME_VAL + " 1",
                APPROVALS_SIGNATURE_VAL + " 1",
                APPROVALS_DATE_VAL + " 1",
                APPROVALS_STATUS_VAL + " 1"));

        /*  add Other1 */
        result.add(fillApprovalInMap("Some Role 1",
                APPROVALS_NAME_VAL + " 1",
                APPROVALS_SIGNATURE_VAL + " 1",
                APPROVALS_DATE_VAL + " 1",
                APPROVALS_STATUS_VAL + " 1"));

        /*  add QC2 */
        result.add(fillApprovalInMap(APPROVALS_FIRST_IN_LIST_ROLE_VAL,
                APPROVALS_NAME_VAL + " 2",
                APPROVALS_SIGNATURE_VAL + " 2",
                APPROVALS_DATE_VAL + " 2",
                APPROVALS_STATUS_VAL + " 2"));

        /*  add Other2 */
        result.add(fillApprovalInMap("Some Role 2",
                APPROVALS_NAME_VAL + " 2",
                APPROVALS_SIGNATURE_VAL + " 2",
                APPROVALS_DATE_VAL + " 2",
                APPROVALS_STATUS_VAL + " 2"));

        /*  add Other3 */
        result.add(fillApprovalInMap("Some Role 3",
                APPROVALS_NAME_VAL + " 3",
                APPROVALS_SIGNATURE_VAL + " 3",
                APPROVALS_DATE_VAL + " 3",
                APPROVALS_STATUS_VAL + " 3"));
        return result;
    }

    private static List<Map<String, String>> setDefaultCertifications() {
        List<Map<String, String>> result = new ArrayList<>();
        for (int i = 0; i < 5; i++) {
            result.add(fillMap(CERTIFICATIONS_ARR, i));
        }
        return result;
    }

    private static List<Map<String, String>> setDefaultPreliminaryInspectionResults() {
        List<Map<String, String>> result = new ArrayList<>();
        for (int i = 0; i < 5; i++) {
            result.add(fillMap(PRELIM_INSPEC_RESULT, i));
        }
        return result;
    }

    private static List<Map<String, String>> setDefaultChecklistItems() {
        List<Map<String, String>> list = new ArrayList<>();
        for (int i = 0; i < 4; i++) {
            list.add(fillMap(CHECKLIST_ITEMS_ARR, i));
        }
        return list;
    }

    private static Map<String, String> fillMap(String[] strings) {
        Map<String, String> result = new HashMap<>();
        for (String a : strings) {
            result.put(a, a);
        }
        return result;
    }

    private static Map<String, String> fillMap(String[] strings,
                                               int i) {
        Map<String, String> result = new HashMap<>();
        for (String a : strings) {
            result.put(a, a + " " + i);
        }
        return result;
    }

    private static Map<String, String> fillApprovalInMap(String first,
                                                         String second,
                                                         String third,
                                                         String fourth,
                                                         String fifth) {
        Map<String, String> result = new HashMap<>();
        result.put(APPROVALS_ROLE_VAL, first);
        result.put(APPROVALS_NAME_VAL, second);
        result.put(APPROVALS_SIGNATURE_VAL, third);
        result.put(APPROVALS_DATE_VAL, fourth);
        result.put(APPROVALS_STATUS_VAL, fifth);
        return result;
    }

    private static List<Map<String, String>> setDefaultDrawings() {
        List<Map<String, String>> result = new ArrayList<>();
        for (int i = 0; i < 5; i++) {
            result.add(fillMap(DRAWINGS_ARR, i));
        }
        return result;
    }

    private static List<Map<String, String>> setDefaultAdditionalDocuments() {
        List<Map<String, String>> result = new ArrayList<>();
        for (int i = 0; i < 10; i++) {


            result.add(fillMap(CERTIFICATIONS_ARR, i));
        }
        return result;
    }

}