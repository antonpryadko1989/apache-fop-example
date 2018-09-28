package com.visoft.services;

import java.util.HashMap;
import java.util.Map;

public interface Const {
//    int[] cellLengthFont20IL = {1, 10, 10, 9, 9, 10, 8, 9, 6};
    int[] cellLengthFont18IL = {1, 12, 12, 10, 10, 12, 9, 10, 7};
    int[] cellLengthFont12IL = {2, 18, 18, 16, 16, 18, 14, 18 ,13};
    int IL_DEF_ROW_HEIGHT_FONT_18 = 35;
    int IL_DEF_ROW_HEIGHT_FONT_12 = 24;
    String TEMPLATES = "E:/WORK/TEMPLATES/";
    int FONT_RATE = 20;
    int COLUMN_WIDTHS_RATE = 36;
//    String FONT_NAME_ARIAL = "Arial";
    String FONT_NAME_CALIBRI = "Calibri";
    String NCR_TEMPLATE_NAME = "טופס אי התאמה";
    String POC_TEMPLATE_NAME = "דו\"ח קטעי ניסוי";
    /**
     *  APP_OF_SUBCONTR  - Approval Of Sub-Contractors
     */
    String APP_OF_SUBCONTR_TEMPLATE_NAME =
            "אישור קבלני משנה";
    /**
     *  APP_OF_SUPPL - Approval Of Suppliers
     */
    String APP_OF_SUPPL_TEMPLATE_NAME =
            "אישור ספקים";
    /**
     *  PRELIM_MATERIALS_INS - Preliminary Materials Inspection
     */
    String PRELIM_MATERIALS_INS_TEMPLATE_NAME =
            "בקרה מקדימה לחומרים";
    String LOGO_VAL = "logoValue";
    String LOGO_PATH = "logoPath";
    String CHECKLIST_FORM_NO = "Form N \n" + "רשימת תיוג מס'";
    String CHECKLIST_QPN = "QP N / מס'נוהל";
    String CHECKLIST_TEMPLATE_NAME = "Procedure Name / שם הנוהל";
    String VERSION = "Version / מהדורה";
    String DATE = "Date /\n תאריך";
    String MAIN_CONTRACTOR = "Main Contractor /  קבלן ראשי";
    String PROJECT_NAME = "Project Name / שם הפרויקט";
    String MANAGEMENT_COMPANY  = "Management Company / חברת ניהול";
    String CONTRACT_NO = "Contract No. / חוזה מסי";
    String QC_COMPANY = "QC Company / חברת בקרת איכות";
    String QA_COMPANY = "QA Company / חברת הבטחת איכות";
    String NCR_NUMBER = "אי התאמה מס' / NCR Number";
    String NCR_OPENED = "NCR opened / נפתח ע\"י";
    String POSITION = "Position / תפקיד";
    String NAME = "Name / שם";
    String DATE_OF_NCR = "Date of NCR / תאריך הפתיחה";
    String QA_QC = "QA / QC";
    String QCM = "QCM";
    String QAM = "QAM";

    String DRAWING_NO = "Drawing No / מס תוכנית";
    String DRAWING_VERSION_REVISION = "Version / Revision / מהדורה";
    String DRAWING_NAME = "Drawing Name / שם תוכנית";

    String DETAILS_OF_STRUCTURE = "Details of Structure / פרטי המבנה";
    String ELEMENT_STATION_ROAD_TUNNEL_BRIDGE = "Element (Station, Road, " +
            "Tunnel, Bridge) / (כביש , רמפה , גשר, מנהרה )";
    String STRUCTURE = "Structure / מבנה";
    String ELEMENT = "Element / אלמנט";
    String FROM_SECTION = "From section / מחתך";
    String TILL_SECTION = "Till section / עד חתך";
    String SIDE = "Side / הסט";
    String NCR_LEVEL = "NCR Level / דרגה";
    String ELEMENT_LAYER = "Element/Layer / אלמנט/שכבה";
    String BLUE_BOOK_CODE = "Blue Book Code/ קוד ספר כחול";
    String NUMBER_OF_DAYS_LATE = "Number Of Days Late/ מס ימי עיכוב לסגירה";
    String SUB_ELEMENT = "Sub-Element / תת אלמנט";
    String[] NCR_ZERO_ROW_HEADER = {ELEMENT_STATION_ROAD_TUNNEL_BRIDGE,
            STRUCTURE, ELEMENT, FROM_SECTION, TILL_SECTION, SIDE, NCR_LEVEL,
            NUMBER_OF_DAYS_LATE};
    String[] NCR_ONE_ROW_HEADER = {STRUCTURE, ELEMENT_LAYER, FROM_SECTION,
            TILL_SECTION, SIDE, NCR_LEVEL, BLUE_BOOK_CODE};
    String[] NCR_TWO_ROW_HEADER = {ELEMENT_STATION_ROAD_TUNNEL_BRIDGE,
            STRUCTURE, ELEMENT_LAYER, SUB_ELEMENT, FROM_SECTION, TILL_SECTION,
            SIDE, NCR_LEVEL};
    String EXPECTED_CLOSING_DATE = "Expected Closing Date / תאריך סגירה משוער";
    String UPDATED_EXPECTED_CLOSING_DATE = "Updated Expected Closing Date / " +
            "תאריך משוער לסגירה מקורי";
    String SUB_PROJECT = "Sub Project / תת פרויקט";
    String NCR_DESCRIPTION = "NCR description / תאור אי ההתאמה";
    String RESPONSIBLE_PARTY = "Responsible Party (Design ,Contractor," +
            "Supplier) / גורם אחראי לליקוי (תכנון, ביצוע, ספק)";
    String CORRECTIVE_ACTION = "Corrective action / פעולה מתקנת נדרשת";
    String DESCRIPTION_OF_PERFORMED_CORRECTIVE_ACTION =
            "Description of performed corrective action / פעולה מתקנת שבוצעה";
    String RESPONSIBLE_PERSON = "Responsible person / שם האחראי";
    String REMARKS = "Remarks / הערות";
    String APPROXIMATE_CLOSING_DATE_NCR_TWO =
            "Approximate closing date / תאריך סגירת אי התאמה משוער-מסוכם";
    String APPROXIMATE_CLOSING_DATE_NCR_ONE =
            "Approximate closing date / תאריך סגירה משוער";
    String RESPONSIBLE_PARTY_DESIGN_CONTRACTOR_SUPPLIER =
            "Responsible Party (Design ,Contractor,Supplier) " +
                    "/ גורם אחראי לליקוי (תכנון, ביצוע, ספק)";
    String ORIGINAL_EXPECTED_CLOSING_DATE = "Original Expected Closing Date / " +
            "תאריך סגירה משוער על פי החלטת מנה\"פ";
    String DAMAGE = "שבר / Damage";
    String QUALITY_IMPACT = "Quality Impact / השפעה על איכות";
    String RESPONSIBLE_PARTY_ROLE = "Responsible Party Role / גורם מטפל";
    String IS_OPENED_BEFORE = "Is Opened Before / אי התאמה חוזרת מאותו סוג";
    String OPENED_BEFORE_NCR_NO = "Opened Before NCR No / מס' אי התאמה חוזרת";
    String TYPE_OF_TEST = "Type Of Test / סוג בדיקה";
    String ACCEPTANCE_OF_CORRECTIVE_ACTION =
            "Acceptance of corrective action / אישור ביצוע פעילות מתקנת";
    String NCR_CLOSED = "NCR closed / נסגרה ע\"י";
    String CLOSING_DATE = "Closing Date / תאריך סגירה";

    String CHECKLIST_FORM_NO_VAL = "formNo";
    String CHECKLIST_QPN_VAL = "qpn";
    String TEMPLATE_NAME_VAL = "templateName";
    String VERSION_VAL = "version";
    String DATE_VAL = "date";
    String MAIN_CONTRACTOR_VAL = "mainContractor";
    String PROJECT_NAME_VAL = "projectName";
    String MANAGEMENT_COMPANY_VAL  = "managementCompany";
    String CONTRACT_NO_VAL = "contractNo";
    String QC_COMPANY_VAL = "qcCompany";
    String QA_COMPANY_VAL = "qaCompany";
    String NCR_NUMBER_VAL = "ncrNumber";
    String QC_OPENED_NAME_VAL = "ncrOpenQCName";
    String QC_OPENED_DATE_OF_NCR_VAL = "ncrOpenQCDate";
    String QA_OPENED_NAME_VAL = "ncrOpenQAName";
    String QA_OPENED_DATE_OF_NCR_VAL = "ncrOpenQADate";
    String DRAWING_NO_VAL = "drawingNo";
    String DRAWING_VERSION_REVISION_VAL = "versionRevision";
    String DRAWING_NAME_VAL = "drawingName";
    String ELEMENT_STATION_ROAD_TUNNEL_BRIDGE_VAL =
            "elementStationRoadTunnelBridge";
    String STRUCTURE_VAL = "structure";
    String ELEMENT_VAL = "element";
    String FROM_SECTION_VAL = "fromSection";
    String TILL_SECTION_VAL = "tillSection";
    String SIDE_VAL = "side";
    String NCR_LEVEL_VAL = "ncrLevel";
    String ELEMENT_LAYER_VAL = "elementLayer";
    String BLUE_BOOK_CODE_VAL = "blueBookCode";
    String SUB_ELEMENT_VAL = "SubElement";
    String NUMBER_OF_DAYS_LATE_VAL = "numberOfDaysLate";
    String EXPECTED_CLOSING_DATE_VAL = "expectedClosingDate";
    String UPDATED_EXPECTED_CLOSING_DATE_VAL = "updatedExpectedClosingDate";
    String SUB_PROJECT_VAL = "subProject";
    String NCR_DESCRIPTION_VAL = "ncrDescription";
    String CORRECTIVE_ACTION_VAL = "correctiveAction";
    String DESCRIPTION_OF_PERFORMED_CORRECTIVE_ACTION_VAL =
            "descriptionOfPerformedCorrectiveAction";
    String RESPONSIBLE_PERSON_VAL = "responsiblePerson";
    String REMARKS_VAL = "remarks";
    String APPROXIMATE_CLOSING_DATE_VAL = "approximateClosingDate";
    String RESPONSIBLE_PARTY_DESIGN_CONTRACTOR_SUPPLIER_VAL =
            "responsiblePartyDesignContractorSupplier";
    String DAMAGE_VAL = "damage";
    String QUALITY_IMPACT_VAL = "qualityImpact";
    String ORIGINAL_EXPECTED_CLOSING_DATE_VAL  =
            "originalExpectedClosingDate";
    String RESPONSIBLE_PARTY_ROLE_VAL = "responsiblePartyRole";
    String IS_OPENED_BEFORE_VAL = "isOpenedBefore";
    String OPENED_BEFORE_NCR_NO_VAL = "openedBeforeNcrNo";
    String TYPE_OF_TEST_VAL = "typeOfTest";
    String QC_CLOSED_NAME_VAL = "ncrClosedQCName";
    String QC_CLOSED_DATE_OF_NCR_VAL = "ncrClosedQCDate";
    String QA_CLOSED_NAME_VAL = "ncrClosedQAName";
    String QA_CLOSED_DATE_OF_NCR_VAL = "ncrClosedQADate";
    String DRAWINGS_VAL = "drawings";

    String POC_REPORT_NO = "Report No. / קטע מס'";
    String POC_TYPE_OF_TEST  = "Type of Test / הוכחת היכולת לפעולה מסוג";
    String POC_ELEMENT = "Element / שם האלמנט";
    String POC_FROM_SECTION  = "From Section / מחתך";
    String POC_UP_TO_SECTION  = "Up To Section / עד חתך";
    String POC_SIDE = "Side / צד";
    String POC_PARTICIPANTS_IN_TEST = "Participants in Test / משתתפים בקטע ניסוי";
    String POC_MATERIAL_FOR_USE = "Material for Use / חומרים לשימוש";
    String POC_TOOLS_USED = "Tools Used / הכלים בהם משתמשים";
    String POC_DATE_OF_EXECUTION = "Date of Execution / תאריך ביצוע";
    String POC_DESCRIPTION_OF_TEST = "Description of Test / תיאור קטע ניסוי";
    String POC_CONCLUSIONS_OF_TEST = "Conclusions of Test / מסקנות קטע ניסוי";
    String POC_CORRECTIVE_ACTION = "Corrective Action / פעולה מתקנת (במידה ונדרשת)";

    String POC_REPORT_NO_VAL = "reportNo";
    String POC_TYPE_OF_TEST_VAL  = "typeOfTest";
    String POC_ELEMENT_VAL = "element";
    String POC_FROM_SECTION_VAL = "fromSection";
    String POC_UP_TO_SECTION_VAL = "tillSection";
    String POC_SIDE_VAL = "side";
    String POC_PARTICIPANTS_IN_TEST_VAL = "participantsInTest";
    String POC_MATERIAL_FOR_USE_VAL = "materialForUse";
    String POC_TOOLS_USED_VAL = "toolsUsed";
    String POC_DATE_OF_EXECUTION_VAL = "dateOfExecution";
    String POC_DESCRIPTION_OF_TEST_VAL = "descriptionOfTest";
    String POC_CONCLUSIONS_OF_TEST_VAL = "conclusionsOfTest";
    String POC_CORRECTIVE_ACTION_VAL = "correctiveAction";

    String APP_OF_SUBCONTR_APPROVAL_NO = "Approval No. / אישור מס'";
    String APP_OF_SUBCONTR_SUB_CONTRACTOR_NAME = "Sub-Contractor name";
    String APP_OF_SUBCONTR_FIELD_OF_ACTIVITY  =
            "Field of Activity / תחום פעילות";
    String APP_OF_SUBCONTR_SUBPROJECT = "Subproject / תת פרויקט";
    String APP_OF_SUBCONTR_CONTACT =
            "Contact Person and telephone / אנשי קשר וטלפון";

    String APP_OF_SUBCONTR_APPROVAL_NO_VAL = "approvalNo";
    String APP_OF_SUBCONTR_SUB_CONTRACTOR_NAME_VAL = "SubContractorName";
    String APP_OF_SUBCONTR_FIELD_OF_ACTIVITY_VAL  =
            "fieldOfActivity";
    String APP_OF_SUBCONTR_SUBPROJECT_VAL = "subProject";
    String APP_OF_SUBCONTR_CONTACT_VAL = "contact";

    String APP_OF_SUPPL_APPROVAL_NO = "Approval No. / אישור מס'";
    String APP_OF_SUPPL_SUPPLIER_NAME  = "Supplier Name / שם ספק";
    String APP_OF_SUPPL_SUBPROJECT = "Subproject / תת פרויקט";
    String APP_OF_SUPPL_CONTACT = "Contact Person and telephone / אנשי קשר וטלפון";
    String APP_OF_SUPPL_SUPPLIED_MATERIALS = "Supplied Materials / חומר מסופק";

    String APP_OF_SUPPL_APPROVAL_NO_VAL = "approvalNo";
    String APP_OF_SUPPL_SUPPLIER_NAME_VAL = "supplierName";
    String APP_OF_SUPPL_SUBPROJECT_VAL = "subProject";
    String APP_OF_SUPPL_CONTACT_VAL = "contact";
    String APP_OF_SUPPL_SUPPLIED_MATERIALS_VAL = "suppliedMaterials";

    String PRELIM_MATERIALS_INS_APPROVAL_NO = "Approval No. / אישור מס'";
    String PRELIM_MATERIALS_INS_SUPPLIER_NAME = "Supplier Name / שם ספק";
    String PRELIM_MATERIALS_INS_SUBPROJECT = "Subproject / תת פרויקט";
    String PRELIM_MATERIALS_INS_SOURCE_OF_MATERIAL =
            "Source Of Material / מקור החומר";
    String PRELIM_MATERIALS_INS_MATERIAL_SUPPLIED =
            "Material Supplied / חומר מסופק";
    String PRELIM_MATERIALS_INS_MATERIAL_USE =
            "Material Use / ייעוד השימוש בחומר";

    String PRELIM_MATERIALS_INS_APPROVAL_NO_VAL = "approvalNo";
    String PRELIM_MATERIALS_INS_SUPPLIER_NAME_VAL = "supplierName";
    String PRELIM_MATERIALS_INS_SUBPROJECT_VAL = "subProject";
    String PRELIM_MATERIALS_INS_SOURCE_OF_MATERIAL_VAL =
            "sourceOfMaterial";
    String PRELIM_MATERIALS_INS_MATERIAL_SUPPLIED_VAL =
            "materialSupplied";
    String PRELIM_MATERIALS_INS_MATERIAL_USE_VAL =
            "materialUse";

    String APPROVALS  = "Approvals / אישורים";
    String APPROVALS_ROLE  = "Role \n תפקיד";
    String APPROVALS_NAME  = "Name \n שם";
    String APPROVALS_SIGNATURE = "Signature \n חתימה";
    String APPROVALS_DATE = "Date \n תאריך";
    String APPROVALS_STATUS = "Status (Approved / Not Approved) \n" +
            "סטטוס (מאושר / לא מאושר)";
    String APPROVALS_VAL = "approvals";       // List name
    String APPROVALS_ROLE_VAL = "role";
    String APPROVALS_NAME_VAL = "name";
    String APPROVALS_SIGNATURE_VAL = "signature";
    String APPROVALS_DATE_VAL = "date";
    String APPROVALS_STATUS_VAL = "status";
    String APPROVALS_FIRST_IN_LIST_ROLE_VAL = "Quality Control (QC)";
    String APPROVALS_LAST_IN_LIST_ROLE_VAL = "Quality Assurance (QA)";
    Map<String, String> POC_APPROVALS_HEAD = new HashMap<String, String>(){
        {
            put(APPROVALS_ROLE_VAL, APPROVALS_FIRST_IN_LIST_ROLE_VAL);
        }
    };
    Map<String, String> POC_APPROVALS_TAIL = new HashMap<String, String>(){
        {
            put(APPROVALS_ROLE_VAL, APPROVALS_LAST_IN_LIST_ROLE_VAL);
        }
    };

    String ADDITIONAL_DOCUMENTS = "Additional Documents / מסמכים נוספים";
    String ADDITIONAL_DOCUMENTS_ITEM = "Item \n" + "פרטים";
    String ADDITIONAL_DOCUMENTS_EXISTS =
            "Exists/Does not exist\n" + "קיים / לא קיים";
    String ADDITIONAL_DOCUMENTS_CERTIFICATE_NO =
            "Certificate No. / אישור מס";
    String ADDITIONAL_DOCUMENTS_EXPIRATION = "Expiration \n" + "תוקף";
    String ADDITIONAL_DOCUMENTS_ATTACHED_DOCUMENTS =
            "Attached Documents\n" + "מסמכים מצורפים";

    String ADDITIONAL_DOCUMENTS_VAL = "additionalDocuments";
    String ADDITIONAL_DOCUMENTS_ITEM_VAL = "item";
    String ADDITIONAL_DOCUMENTS_EXISTS_VAL = "exists";
    String ADDITIONAL_DOCUMENTS_CERTIFICATE_NO_VAL = "certificateNo";
    String ADDITIONAL_DOCUMENTS_EXPIRATION_VAL = "expiration";
    String ADDITIONAL_DOCUMENTS_ATTACHED_DOCUMENTS_VAL = "attachedDocuments";

    String CERTIFICATIONS = "Certifications / תעודות";
    String CERTIFICATIONS_ITEM = "Item \n פרטים";
    String CERTIFICATIONS_EXISTS =
            "Exists/Does not exist\n קיים / לא קיים";
    String CERTIFICATIONS_CERTIFICATE_NO = "Certificate No. / אישור מס";
    String CERTIFICATIONS_EXPIRATION = "Expiration \n תוקף";
    String CERTIFICATIONS_ATTACHED_DOCUMENTS =
            "Attached Documents \n מסמכים מצורפים";
    String CERTIFICATIONS_VAL = "certifications";
    String CERTIFICATIONS_ITEM_VAL = "item";
    String CERTIFICATIONS_EXISTS_VAL = "exists";
    String CERTIFICATIONS_CERTIFICATE_NO_VAL = "certificateNo";
    String CERTIFICATIONS_EXPIRATION_VAL = "expiration";
    String CERTIFICATIONS_ATTACHED_DOCUMENTS_VAL = "attachedDocuments";


    String PRELIMINARY_INSPECTION_RESULTS =
            "Prelimitary Inspection Results / סיכום בדיקה מקדימה";
    String PRELIM_INSPEC_RESULT_TYPE_OF_INSPECTION =
            "Type of Inspection / סוג הבדיקה";
    String PRELIM_INSPEC_RESULT_SPEC_REQUIREMENTS =
            "Spec Requirements / דרישות מפרטיות";
    String PRELIM_INSPEC_RESULT_INSPECTION_RESULTS =
            "Inspection Results / תוצאות בדיקה";
    String PRELIM_INSPEC_RESULT_CERTIFICATE_NO =
            "Certificate N / oתעודה מס'";
    String PRELIM_INSPEC_RESULT_PASS_FAIL =
            "Pass- Fail / עבר - נכשל";
    String PRELIM_INSPEC_RESULT_COMMENTS =
            "Comments / הערות";

    String PRELIMINARY_INSPECTION_RESULTS_VAL = "prelimitaryInspectionResults";
    String PRELIM_INSPEC_RESULT_TYPE_OF_INSPECTION_VAL = "typeOfInspection";
    String PRELIM_INSPEC_RESULT_SPEC_REQUIREMENTS_VAL = "specRequirements";
    String PRELIM_INSPEC_RESULT_INSPECTION_RESULTS_VAL = "inspectionResults";
    String PRELIM_INSPEC_RESULT_CERTIFICATE_NO_VAL = "certificateNo";
    String PRELIM_INSPEC_RESULT_PASS_FAIL_VAL = "passFail";
    String PRELIM_INSPEC_RESULT_COMMENTS_VAL = "comments";

    String CHECKLIST_ELEMENTS_VAL = "elements";
    String CHECKLIST_ELEMENT_KEY_VAL = "elementKey";
    String CHECKLIST_ELEMENT_VALUE_VAL = "elementValue";

    String CHECKLIST_ITEMS_WORKDEFINITION = "Work Definition / תאור העבודה לבקרה";
    String CHECKLIST_ITEMS_RESPONSIBLE_PARTY = "Responsible Party / אחריות";
    String CHECKLIST_ITEMS_NAME = "Name / שם";
    String CHECKLIST_ITEMS_SIGNATURE = "Signature / חתימה";
    String CHECKLIST_ITEMS_DATE = "Date / תאריך";
    String CHECKLIST_ITEMS_NOTES = "Notes /\n הערות";

    String CHECKLIST_ITEMS_VAL = "items";
    String CHECKLIST_ITEMS_WORK_DEFINITION_VAL = "workDefinition";
    String CHECKLIST_ITEMS_RESPONSIBLE_PARTY_VAL = "responsibleParty";
    String CHECKLIST_ITEMS_NAME_VAL = "name";
    String CHECKLIST_ITEMS_SIGNATURE_VAL = "signature";
    String CHECKLIST_ITEMS_DATE_VAL = "date";
    String CHECKLIST_ITEMS_NOTES_VAL = "notes";


    String NCR = "NCR";
    String POC = "POC";
    String CHECKLIST = "CHECKLIST";
    String APPROVAL_OF_SUBCONTRACTORS = "APPROVAL_OF_SUBCONTRACTORS";
    String APPROVAL_OF_SUPPLIERS = "APPROVAL_OF_SUPPLIERS";
    String PRELIMINARY_MATERIALS_INSPECTION = "PRELIMINARY_MATERIALS_INSPECTION";


    String[] POC_BODY = {POC_REPORT_NO_VAL, POC_TYPE_OF_TEST_VAL,
            POC_ELEMENT_VAL, POC_FROM_SECTION_VAL,
            POC_UP_TO_SECTION_VAL, POC_SIDE_VAL,
            POC_PARTICIPANTS_IN_TEST_VAL, POC_MATERIAL_FOR_USE_VAL,
            POC_TOOLS_USED_VAL, POC_DATE_OF_EXECUTION_VAL,
            POC_DESCRIPTION_OF_TEST_VAL, POC_CONCLUSIONS_OF_TEST_VAL,
            POC_CORRECTIVE_ACTION_VAL
    };

    String[] NCR_BODY = {NCR_NUMBER_VAL, QC_OPENED_NAME_VAL,
            QC_OPENED_DATE_OF_NCR_VAL, QA_OPENED_NAME_VAL,
            QA_OPENED_DATE_OF_NCR_VAL, ELEMENT_STATION_ROAD_TUNNEL_BRIDGE_VAL,
            STRUCTURE_VAL, ELEMENT_VAL, FROM_SECTION_VAL, TILL_SECTION_VAL,
            SIDE_VAL, ELEMENT_LAYER_VAL, BLUE_BOOK_CODE_VAL, NCR_LEVEL_VAL,
            NUMBER_OF_DAYS_LATE_VAL, EXPECTED_CLOSING_DATE_VAL,
            UPDATED_EXPECTED_CLOSING_DATE_VAL, SUB_PROJECT_VAL,
            NCR_DESCRIPTION_VAL,
            RESPONSIBLE_PARTY_DESIGN_CONTRACTOR_SUPPLIER_VAL,
            CORRECTIVE_ACTION_VAL,
            DESCRIPTION_OF_PERFORMED_CORRECTIVE_ACTION_VAL,
            RESPONSIBLE_PERSON_VAL, REMARKS_VAL,
            QC_CLOSED_NAME_VAL, QC_CLOSED_DATE_OF_NCR_VAL,
            QA_CLOSED_NAME_VAL, QA_CLOSED_DATE_OF_NCR_VAL,
    };

    String[] DEF_BODY_ELEMENTS = {LOGO_VAL, TEMPLATE_NAME_VAL, VERSION_VAL,
            DATE_VAL, MAIN_CONTRACTOR_VAL, MANAGEMENT_COMPANY_VAL,
            QC_COMPANY_VAL, PROJECT_NAME_VAL, CONTRACT_NO_VAL, QA_COMPANY_VAL
    };

    String[] APP_OF_SUBCONTR_APPROVAL_BODY = {APP_OF_SUBCONTR_APPROVAL_NO_VAL,
            APP_OF_SUBCONTR_SUB_CONTRACTOR_NAME_VAL,
            APP_OF_SUBCONTR_FIELD_OF_ACTIVITY_VAL,
            APP_OF_SUBCONTR_SUBPROJECT_VAL, APP_OF_SUBCONTR_CONTACT_VAL
    };

    String[] APP_OF_SUPPL_APPROVAL_BODY = {APP_OF_SUPPL_SUPPLIER_NAME_VAL,
            APP_OF_SUPPL_SUBPROJECT_VAL, APP_OF_SUPPL_CONTACT_VAL,
            APP_OF_SUPPL_SUPPLIED_MATERIALS_VAL
    };

    String[] PRELIM_MATERIALS_INS_APPROVAL_BODY = {
            PRELIM_MATERIALS_INS_APPROVAL_NO_VAL,
            PRELIM_MATERIALS_INS_SUPPLIER_NAME_VAL,
            PRELIM_MATERIALS_INS_SUBPROJECT_VAL,
            PRELIM_MATERIALS_INS_SOURCE_OF_MATERIAL_VAL,
            PRELIM_MATERIALS_INS_MATERIAL_SUPPLIED_VAL,
            PRELIM_MATERIALS_INS_MATERIAL_USE_VAL
        };

    String[] CHECKLIST_BODY = {
            CHECKLIST_FORM_NO_VAL,
            CHECKLIST_QPN_VAL
    };

    String[] CHECKLIST_ELEMENT_KEY_VALUE = {
            CHECKLIST_ELEMENT_KEY_VAL,
            CHECKLIST_ELEMENT_VALUE_VAL
    };

    String[] CERTIFICATIONS_ARR = {
            CERTIFICATIONS_ITEM_VAL,
            CERTIFICATIONS_EXISTS_VAL,
            CERTIFICATIONS_CERTIFICATE_NO_VAL,
            CERTIFICATIONS_EXPIRATION_VAL ,
            CERTIFICATIONS_ATTACHED_DOCUMENTS_VAL

    };

    String[] PRELIM_INSPEC_RESULT = {
            PRELIM_INSPEC_RESULT_TYPE_OF_INSPECTION_VAL,
            PRELIM_INSPEC_RESULT_SPEC_REQUIREMENTS_VAL,
            PRELIM_INSPEC_RESULT_INSPECTION_RESULTS_VAL,
            PRELIM_INSPEC_RESULT_CERTIFICATE_NO_VAL,
            PRELIM_INSPEC_RESULT_PASS_FAIL_VAL,
            PRELIM_INSPEC_RESULT_COMMENTS_VAL
    };

    String[] CHECKLIST_ITEMS_ARR = {
            CHECKLIST_ITEMS_WORK_DEFINITION_VAL,
            CHECKLIST_ITEMS_RESPONSIBLE_PARTY_VAL,
            CHECKLIST_ITEMS_NAME_VAL,
            CHECKLIST_ITEMS_SIGNATURE_VAL,
            CHECKLIST_ITEMS_DATE_VAL,
            CHECKLIST_ITEMS_NOTES_VAL
    };

    String[] DRAWINGS_ARR = {
            DRAWING_NO_VAL,
            DRAWING_VERSION_REVISION_VAL,
            DRAWING_NAME_VAL
    };

    /*
    * 1 index of column B
    * 6 index of column G
    * return merging columns length for 18 font size
     */
    static int getDefLengthBGFont18IL(){
        int count = 0;
        for (int i = 1; i <= 6; i++){
            count += cellLengthFont18IL[i];
        }
        return count;
    }

    /*
   * 1 index of column B
   * 4 index of column E
   * return merging columns length for 12 font size
    */
    static int getDefLengthBEFont12IL(){
        int count = 0;
        for (int i = 1; i <= 4; i++){
            count += cellLengthFont12IL[i];
        }
        return count;
    }

    /*
    * 5 index of column F
    * 8 index of column I
    * return merging columns length for 12 font size
    */
    static int getDefLengthFIFont12IL(){
        int count = 0;
        for (int i = 5; i <= 8; i++){
            count += cellLengthFont12IL[i];
        }
        return count;
    }

    /*
    * 3 index of column D
    * 6 index of column G
    * return merging columns length for 12 font size
    */
    static int getDefLengthDGFont12IL(){
        int count = 0;
        for (int i = 3; i <= 6; i++){
            count += cellLengthFont12IL[i];
        }
        return count;
    }

    /*
     * 3 index of column D
     * 8 index of column I
     * return merging columns length for 12 font size
     */
    static int getDefLengthDIFont12IL(){
        int count = 0;
        for (int i = 3; i <= 8; i++){
            count += cellLengthFont12IL[i];
        }
        return count;
    }

    /*
    * 1 index of column B
    * 2 index of column C
    *  return merging columns length for 12 font size
    */
    static int getDefLengthBCFont12IL(){
        return cellLengthFont12IL[1] +
               cellLengthFont12IL[2];
    }

    /*
    * 3 index of column D
    * 4 index of column E
    *  return merging columns length for 12 font size
    */
    static int getDefLengthDEFont12IL(){
        return cellLengthFont12IL[3] +
               cellLengthFont12IL[4];
    }

    /*
    * 5 index of column F
    * 6 index of column G
    *  return merging columns length for 12 font size
    */
    static int getDefLengthFGFont12IL(){
        return cellLengthFont12IL[5] +
               cellLengthFont12IL[6];
    }

    /*
    * 7 index of column H
    * 6 index of column I
    *  return merging columns length for 12 font size
    */
    static int getDefLengthHIFont12IL(){
        return cellLengthFont12IL[7] +
               cellLengthFont12IL[8];
    }

    /*
     * 1 index of column B
     * 2 index of column C
     * 3 index of column D
     *  return merging columns length for 12 font size
     */
    static int getDefLengthBCDFont12IL(){
        return cellLengthFont12IL[1] +
               cellLengthFont12IL[2] +
               cellLengthFont12IL[3];
    }

    /*
     * 4 index of column E
     * 5 index of column F
     *  return merging columns length for 12 font size
     */
    static int getDefLengthEFFont12IL(){
        return cellLengthFont12IL[4] +
               cellLengthFont12IL[5];
    }

    /*
     * 6 index of column G
     * 7 index of column H
     * 8 index of column I
     *  return merging columns length for 12 font size
     */
    static int getDefLengthGHIFont12IL(){
        return cellLengthFont12IL[6] +
               cellLengthFont12IL[7] +
               cellLengthFont12IL[8];
    }
}