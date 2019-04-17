package com.visoft.utils.excelReportUtils;

public interface XLSXServiceConst {
    int COLUMN_WIDTHS_RATE = 36;

    String NOT_ENOUGH_INFORMATION = "NOT_ENOUGH_INFORMATION";
    String ERROR = "error";
    String NOT_SUPPORTED = "NOT_SUPPORTED";

    String HANDLING_TRAFFIC_ACCIDENT = "HANDLING_TRAFFIC_ACCIDENT";
    String FAILES_SAFETY = "FAILES_SAFETY";
    String TREATMENT_OF_IMPAIRED_BY_MONITORING_REPORT = "TREATMENT_OF_IMPAIRED_BY_MONITORING_REPORT";
    String HANDLING_TRAFFIC_ACCIDENT_REPORT_SHEET_NAME = "טופס תאונת דרכים";
    String FAILES_SAFETY_REPORT_SHEET_NAME = "טופס טיפול ליקוי בטיחות";
    String TREATMENT_OF_IMPAIRED_BY_MONITORING_REPORT_SHEET_NAME = "טופס טיפול בליקוי לאחר ניטור";

    String HANDLING_TRAFFIC_ACCIDENT_SUMMARY_REPORT = "HANDLING_TRAFFIC_ACCIDENT_SUMMARY_REPORT";
    String HANDLING_TRAFFIC_ACCIDENT_SUMMARY_REPORT_SHEET_NAME = "ריכוז תאונות דרכים";
    String HANDLING_TRAFFIC_ACCIDENT_SUMMARY_REPORT_HEADER_NAME = " ריכוז טיפול בתאונות דרכים";
    int[] HANDLING_TRAFFIC_ACCIDENT_SUMMARY_REPORT_COLUMN_WIDTHS = {
            110, 138, 198, 190, 384, 404, 398, 344,
            321, 472, 506, 222, 252, 282, 268, 324
    };

    String FAILES_SAFETY_SUMMARY_REPORT = "FAILES_SAFETY_SUMMARY_REPORT";
    String FAILES_SAFETY_SUMMARY_REPORT_SHEET_NAME = "ריכוז ליקוי בטיחות לאחר ניטור";
    String FAILES_SAFETY_SUMMARY_REPORT_HEADER_NAME = "טופס ריכוז -טיפול בליקוי בטיחות";
    int[] FAILES_SAFETY_SUMMARY_REPORT_COLUMN_WIDTHS = {
            116, 132, 156, 132, 168, 342, 276, 300, 314,
            310, 360, 386, 252, 408, 258, 304, 284, 178
    };

    String SAFETY_SUPERVISORY_TREATMENT_AFTER_MONITORING_SUMMARY_REPORT = "SAFETY_SUPERVISORY_TREATMENT_AFTER_MONITORING_SUMMARY_REPORT";
    String SAFETY_SUPERVISORY_TREATMENT_AFTER_MONITORING_SUMMARY_REPORT_SHEET_NAME = "ריכוז ליקוי בטיחות לאחר ניטור";
    String SAFETY_SUPERVISORY_TREATMENT_AFTER_MONITORING_SUMMARY_REPORT_HEADER_NAME = "טופס ריכוז -טיפול בליקוי בטיחות לאחר ניטור";
    int[] SAFETY_SUPERVISORY_TREATMENT_AFTER_MONITORING_SUMMARY_REPORT_COLUMN_WIDTHS = {
            140, 188, 263, 178, 330, 275, 258, 435, 228, 355,
            501, 278, 278, 390, 293, 428, 258, 195, 386, 240
    };

    int[] DEFAULT_REPORT_COLUMN_WIDTHS = {
            10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10,
            10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10
    };

    String HANDLING_TRAFFIC_ACCIDENT_REPORT_NAME = "טופס טיפול באירוע תאונת דרכים";
    String FAILES_SAFETY_REPORT_NAME = "טופס טיפול בליקוי בטיחות";
    String TREATMENT_OF_IMPAIRED_BY_MONITORING_REPORT_NAME = "טופס טיפול בליקוי לפי דוח ניטור";
    String VERSION = "מהדורה";
    String DATE = "תאריך";

    String MAIN_CONTRACTOR = "קבלן ראשי";
    String PROJECT_NAME = "שם הפרויקט";
    String QC_COMPANY = "חברת בקרת איכות";
    String QA_COMPANY = "חברת הבטחת איכות";
    String MANAGEMENT_COMPANY = "חוזה מסי";
    String CONTRACT_NUMBER = "חברת ניהול";

    String SAFETY_DISADVANTAGES_NUMBER = "ליקוי בטיחות מס'";
    String FAILES_SAFETY_NUMBER = "ליקוי בטיחות מס'";
    String ROLE = "תפקיד";
    String OPENED_BY = "נתפח ע\"י";
    String OPENED_DATE = "תאיריך פתיחה";
    String OPENED_HOUR = "שעת פתיחה";
    String ACCIDENT_DETAILS = "פרטי תאונה";
    String SAFETY_EVENT_INFORMATION = "פרטי אירוע בטיחות";
    String ELEMENT = "אלמנט";
    String LOCATION = "מיקום התאונה";
    String SIDE = "הסט";

    /*TREATMENT_OF_IMPAIRED_BY_MONITORING*/
    String MONITORING_DATE = "תאריך ניטור";
    String TEST_CODE = "קוד בדיקה";

    /*HANDLING_TRAFFIC_ACCIDENT */
    String DETAILS_OF_ACCIDENT = "פירוט תאונה";
    String INVOLVED_CARS = "רכבים מעורבים";
    String TELEPHONE_REPORTING_FOR_PROJECT_MANAGEMENT = "דיווח טלפוני להנהלת הפרויקט";
    String CALL_CENTER_REPORTING_OPERATIONS_TEAM  = "דיווח טלפוני לצוות תפעול";
    String REPORT_BY_TELEPHONE_THE_POLICE = "דווח טלפוני למשטרה";
    String TELEPHONE_REPORTING_AND_EMERGENCY_POWERS = "דיווח טלפוני לכוחות חירום";
    String SPECIAL_PATROL_UNIT_STAFF_INFO = "פרטי צוות היס\"מ";

    /*FAILES_SAFETY*/
    String SAFETY_FLAW_DETAILS = "פירוט ליקוי בטיחות";
    String TELEPHONE_REPORTING_OPERATIONS_TEAM = "דיווח טלפוני לצוות תפעול";
    String IMAGE_BEFORE = "תמונה ממקום האירוע לפני הטיפול";
    String CORRECTIVE_ACTION = "טיפול הנדרש";
    String RESPONSIBLE_PERSON = " גורם המטפל";
    String DESCRIPTION_OF_PERFORMED_CORRECTIVE_ACTION = " פירוט ביצוע פעולה מתקנת";
    String FINISH_DATE = "תאריך סיום";
    String END_TIME = "שעת סיום";
    String TELEPHONE_REPORTING_TO_THE_OPERATIONS_TEAM_FOR_TERMINATION_OF_TREATMENT = "דיווח טלפוני למרכז התפעול לסיום הטיפול";
    String IMAGE_AFTER = "תמונה ממקום האירוע אחרי הטיפול";

    String REMARKS = "הערות";

    String HANDLING_TRAFFIC_ACCIDENT_ADDITIONAL_DOCUMENTS = " מסמכים נוספים (תמונות או כל מסמך אחר)";
    String HANDLING_TRAFFIC_ACCIDENT_ADDITIONAL_DOCUMENT_TYPE = "פרטים";
    String HANDLING_TRAFFIC_ACCIDENT_ADDITIONAL_DOCUMENT_DESCRIPTION = "מסמכים מצורפים";

    String FAILES_SAFETY_ADDITIONAL_DOCUMENTS = "מסמכים נוספים";
    String FAILES_SAFETY_ADDITIONAL_DOCUMENT_TYPE = "פרטים";
    String FAILES_SAFETY_ADDITIONAL_DOCUMENT_DESCRIPTION = "מסמכים מצורפים";

    String TREATMENT_OF_IMPAIRED_BY_MONITORING_ADDITIONAL_DOCUMENTS = " מסמכים נוספים";
    String TREATMENT_OF_IMPAIRED_BY_MONITORING_ADDITIONAL_DOCUMENT_TYPE = "פרטים";
    String TREATMENT_OF_IMPAIRED_BY_MONITORING_ADDITIONAL_DOCUMENT_DESCRIPTION = "מסמכים מצורפים";

    String REPORT_NAME_VAL = "reportName";
    String REPORT_VERSION_VAL = "version";
    String REPORT_DATE_VAL = "date";

    String MAIN_CONTRACTOR_VAL = "mainContractor";
    String PROJECT_NAME_VAL = "projectName";
    String QC_COMPANY_VAL =  "qcCompany";
    String QA_COMPANY_VAL = "qaCompany";
    String MANAGEMENT_COMPANY_VAL = "managementCompany";
    String CONTRACT_NUMBER_VAL = "contractNumber";

    String REPORT_NUMBER_VAL = "reportNumber";
    String ROLE_VAL = "role";
    String OPENED_BY_VAL = "openedBy";
    String OPENED_DATE_VAL = "openedDate";
    String OPENED_HOUR_VAL = "openedHour";
    String ELEMENT_VAL = "element";
    String LOCATION_VAL = "location";
    String SIDE_VAL = "side";

    /*TREATMENT_OF_IMPAIRED_BY_MONITORING*/
    String TEST_CODE_VAL = "testCode";
    String MONITORING_DATE_VAL = "monitoringDate";

    /*HANDLING_TRAFFIC_ACCIDENT */
    String DETAILS_OF_ACCIDENT_VAL = "detailsOfAccident";
    String INVOLVED_CARS_VAL = "involvedCars";
    String TELEPHONE_REPORTING_FOR_PROJECT_MANAGEMENT_VAL = "telephoneReportingForProjectManagement";
    String CALL_CENTER_REPORTING_OPERATIONS_TEAM_VAL = "callCenterReportingOperationsTeam";
    String REPORT_BY_TELEPHONE_THE_POLICE_VAL = "reportByTelephoneThePolice";
    String TELEPHONE_REPORTING_AND_EMERGENCY_POWERS_VAL = "telephoneReportingAndEmergencyPowers";
    String SPECIAL_PATROL_UNIT_STAFF_INFO_VAL = "specialPatrolUnitStaffInfo";

    /*FAILES_SAFETY*/
    String SAFETY_FLAW_DETAILS_VAL = "safetyFlawDetails";
    String TELEPHONE_REPORTING_OPERATIONS_TEAM_VAL = "telephoneReportingOperationsTeam";
    String IMAGE_BEFORE_VAL = "imageBefore";
    String CORRECTIVE_ACTION_VAL = "correctiveAction";
    String RESPONSIBLE_PERSON_VAL = "responsiblePerson";
    String DESCRIPTION_OF_PERFORMED_CORRECTIVE_ACTION_VAL = "descriptionOfPerformedCorrectiveAction";
    String FINISH_DATE_VAL = "finishDate";
    String END_TIME_VAL = "endTime";
    String TELEPHONE_REPORTING_TO_THE_OPERATIONS_TEAM_FOR_TERMINATION_OF_TREATMENT_VAL = "telephoneReportingToTheOperationsTeamForTerminationOfTreatment";
    String IMAGE_AFTER_VAL = "imageAfter";

    String REMARKS_VAL = "remarks";

    String QC_NAME_VAL = "qcName";
    String QC_APPROVAL_DATE_VAL = "qcApprovalDate";
    String FOREMAN_NAME_VAL = "foremanName";
    String FOREMAN_APPROVAL_DATE_VAL = "foremanApprovalDate";

    String FAILES_SAFETY_ACCEPTANCE_OF_CORRECTIVE_ACTION = "אישור ביצוע פעילות מתקנת";
    String ACCEPTANCE_OF_CORRECTIVE_ACTION_ROLE = "תפקיד";
    String ACCEPTANCE_OF_CORRECTIVE_ACTION_ROLE_QC = "QC";
    String ACCEPTANCE_OF_CORRECTIVE_ACTION_ROLE_FOREMAN = "מנהל עבודה";

    String ACCEPTANCE_OF_CORRECTIVE_ACTION_NAME = "שם";
    String ACCEPTANCE_OF_CORRECTIVE_ACTION_DATE = " תאריך אישור";

    String ADDITIONAL_DOCUMENTS_VAL = "additionalDocuments";
    String ADDITIONAL_DOCUMENT_TYPE_VAL = "type";
    String ADDITIONAL_DOCUMENT_DESCRIPTION_VAL = "description";

    String[] FAILES_SAFETY_REPORT_BODY_KEYS = {
            SAFETY_FLAW_DETAILS,
            TELEPHONE_REPORTING_FOR_PROJECT_MANAGEMENT,
            TELEPHONE_REPORTING_OPERATIONS_TEAM,
            IMAGE_BEFORE,
            CORRECTIVE_ACTION,
            RESPONSIBLE_PERSON,
            DESCRIPTION_OF_PERFORMED_CORRECTIVE_ACTION,
            FINISH_DATE,
            END_TIME,
            TELEPHONE_REPORTING_TO_THE_OPERATIONS_TEAM_FOR_TERMINATION_OF_TREATMENT,
            IMAGE_AFTER,
            REMARKS
    };
    String[]  FAILES_SAFETY_REPORT_BODY_VALUES = {
            SAFETY_FLAW_DETAILS_VAL,
            TELEPHONE_REPORTING_FOR_PROJECT_MANAGEMENT_VAL,
            TELEPHONE_REPORTING_OPERATIONS_TEAM_VAL,
            IMAGE_BEFORE_VAL,
            CORRECTIVE_ACTION_VAL,
            RESPONSIBLE_PERSON_VAL,
            DESCRIPTION_OF_PERFORMED_CORRECTIVE_ACTION_VAL,
            FINISH_DATE_VAL,
            END_TIME_VAL,
            TELEPHONE_REPORTING_TO_THE_OPERATIONS_TEAM_FOR_TERMINATION_OF_TREATMENT_VAL,
            IMAGE_AFTER_VAL,
            REMARKS_VAL
    };

    String[] HANDLING_TRAFFIC_ACCIDENT_REPORT_BODY_KEYS = {
            DETAILS_OF_ACCIDENT,
            INVOLVED_CARS,
            TELEPHONE_REPORTING_FOR_PROJECT_MANAGEMENT,
            CALL_CENTER_REPORTING_OPERATIONS_TEAM,
            REPORT_BY_TELEPHONE_THE_POLICE,
            TELEPHONE_REPORTING_AND_EMERGENCY_POWERS,
            SPECIAL_PATROL_UNIT_STAFF_INFO,
            REMARKS
    };
    String[]  HANDLING_TRAFFIC_ACCIDENT_REPORT_BODY_VALUES = {
            DETAILS_OF_ACCIDENT_VAL,
            INVOLVED_CARS_VAL,
            TELEPHONE_REPORTING_FOR_PROJECT_MANAGEMENT_VAL,
            CALL_CENTER_REPORTING_OPERATIONS_TEAM_VAL,
            REPORT_BY_TELEPHONE_THE_POLICE_VAL,
            TELEPHONE_REPORTING_AND_EMERGENCY_POWERS_VAL,
            SPECIAL_PATROL_UNIT_STAFF_INFO_VAL,
            REMARKS_VAL
    };

    String[] TREATMENT_OF_IMPAIRED_BY_MONITORING_REPORT_BODY_KEYS = {
            MONITORING_DATE,
            TEST_CODE,
            SAFETY_FLAW_DETAILS,
            TELEPHONE_REPORTING_FOR_PROJECT_MANAGEMENT,
            TELEPHONE_REPORTING_OPERATIONS_TEAM,
            IMAGE_BEFORE,
            CORRECTIVE_ACTION,
            RESPONSIBLE_PERSON,
            DESCRIPTION_OF_PERFORMED_CORRECTIVE_ACTION,
            FINISH_DATE,
            END_TIME,
            TELEPHONE_REPORTING_TO_THE_OPERATIONS_TEAM_FOR_TERMINATION_OF_TREATMENT,
            IMAGE_AFTER,
            REMARKS
    };

    String[] TREATMENT_OF_IMPAIRED_BY_MONITORING_REPORT_BODY_VALUES = {
            MONITORING_DATE_VAL,
            TEST_CODE_VAL,
            SAFETY_FLAW_DETAILS_VAL,
            TELEPHONE_REPORTING_FOR_PROJECT_MANAGEMENT_VAL,
            TELEPHONE_REPORTING_OPERATIONS_TEAM_VAL,
            IMAGE_BEFORE_VAL,
            CORRECTIVE_ACTION_VAL,
            RESPONSIBLE_PERSON_VAL,
            DESCRIPTION_OF_PERFORMED_CORRECTIVE_ACTION_VAL,
            FINISH_DATE_VAL,
            END_TIME_VAL,
            TELEPHONE_REPORTING_TO_THE_OPERATIONS_TEAM_FOR_TERMINATION_OF_TREATMENT_VAL,
            IMAGE_AFTER_VAL,
            REMARKS_VAL
    };
}