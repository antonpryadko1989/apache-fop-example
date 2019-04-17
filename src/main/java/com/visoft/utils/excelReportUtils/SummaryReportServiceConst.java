package com.visoft.utils.excelReportUtils;

import java.util.Arrays;
import java.util.List;

public interface SummaryReportServiceConst extends XLSXServiceConst{

    String NUMBER = "מסי";
    String FAILSE_SAFETY_NO = "מסי ליקוי";
    String FAILES_SAFETY_SUMMARY_REPORT_LOCATION = "מיקום הליקוי";
    String SPECIAL_PATROL_UNIT_STAFF_INFO = "פרטי צוות היס\"מ";
    String OPENED_NAME = "נפתחה  על יד";
    String STATUS = "סטטוס (טופל/ בטיפול)";
    String OPENING_DATE = "תאריך פתיחה";
    String FINISH_DATE = "תאריך  סיום";

    String STATUS_VAL = "status";
    String FINISH_DATE_VAL = "finishDate";

    List<String> HANDLING_TRAFFIC_ACCIDENT_SUMMARY_REPORT_TITLE = Arrays.asList(
            NUMBER,
            FAILSE_SAFETY_NO,
            OPENED_DATE,
            OPENED_HOUR,
            OPENED_NAME,
            SPECIAL_PATROL_UNIT_STAFF_INFO,
            LOCATION,
            ELEMENT,
            SIDE,
            DETAILS_OF_ACCIDENT,
            INVOLVED_CARS,
            TELEPHONE_REPORTING_FOR_PROJECT_MANAGEMENT,
            CALL_CENTER_REPORTING_OPERATIONS_TEAM,
            REPORT_BY_TELEPHONE_THE_POLICE,
            TELEPHONE_REPORTING_AND_EMERGENCY_POWERS,
            REMARKS
    );

    List<String> HANDLING_TRAFFIC_ACCIDENT_SUMMARY_REPORT_ROW_VALUES = Arrays.asList(
            REPORT_NUMBER_VAL,
            OPENED_DATE_VAL,
            OPENED_HOUR_VAL,
            OPENED_BY_VAL,
            SPECIAL_PATROL_UNIT_STAFF_INFO_VAL,
            LOCATION_VAL,
            ELEMENT_VAL,
            SIDE_VAL,
            DETAILS_OF_ACCIDENT_VAL,
            INVOLVED_CARS_VAL,
            TELEPHONE_REPORTING_FOR_PROJECT_MANAGEMENT_VAL,
            CALL_CENTER_REPORTING_OPERATIONS_TEAM_VAL,
            REPORT_BY_TELEPHONE_THE_POLICE_VAL,
            TELEPHONE_REPORTING_AND_EMERGENCY_POWERS_VAL,
            REMARKS_VAL
    );

    List<String> FAILES_SAFETY_SUMMARY_REPORT_TITLE = Arrays.asList(
            NUMBER,
            FAILSE_SAFETY_NO,
            OPENING_DATE,
            OPENED_HOUR,
            OPENED_NAME,
            FAILES_SAFETY_SUMMARY_REPORT_LOCATION,
            ELEMENT,
            SIDE,
            SAFETY_FLAW_DETAILS,
            TELEPHONE_REPORTING_FOR_PROJECT_MANAGEMENT,
            TELEPHONE_REPORTING_OPERATIONS_TEAM,
            CORRECTIVE_ACTION,
            RESPONSIBLE_PERSON,
            DESCRIPTION_OF_PERFORMED_CORRECTIVE_ACTION,
            FINISH_DATE,
            END_TIME,
            REMARKS,
            STATUS
    );

    List<String> FAILES_SAFETY_SUMMARY_REPORT_ROW_VALUES = Arrays.asList(
            REPORT_NUMBER_VAL,
            OPENED_DATE_VAL,
            OPENED_HOUR_VAL,
            OPENED_BY_VAL,
            LOCATION_VAL,
            ELEMENT_VAL,
            SIDE_VAL,
            SAFETY_FLAW_DETAILS_VAL,
            TELEPHONE_REPORTING_FOR_PROJECT_MANAGEMENT_VAL,
            TELEPHONE_REPORTING_OPERATIONS_TEAM_VAL,
            CORRECTIVE_ACTION_VAL,
            RESPONSIBLE_PERSON_VAL,
            DESCRIPTION_OF_PERFORMED_CORRECTIVE_ACTION_VAL,
            FINISH_DATE_VAL,
            END_TIME_VAL,
            REMARKS_VAL,
            STATUS_VAL
    );

    List<String> SAFETY_SUPERVISORY_TREATMENT_AFTER_MONITORING_SUMMARY_REPORT_TITLE = Arrays.asList(
            NUMBER,
            FAILSE_SAFETY_NO,
            OPENING_DATE,
            OPENED_HOUR,
            OPENED_NAME,
            MONITORING_DATE,
            TEST_CODE,
            FAILES_SAFETY_SUMMARY_REPORT_LOCATION,
            ELEMENT,
            SIDE,
            SAFETY_FLAW_DETAILS,
            TELEPHONE_REPORTING_FOR_PROJECT_MANAGEMENT,
            TELEPHONE_REPORTING_OPERATIONS_TEAM,
            CORRECTIVE_ACTION,
            RESPONSIBLE_PERSON,
            DESCRIPTION_OF_PERFORMED_CORRECTIVE_ACTION,
            FINISH_DATE,
            END_TIME,
            REMARKS,
            STATUS
    );

    List<String> SAFETY_SUPERVISORY_TREATMENT_AFTER_MONITORING_SUMMARY_REPORT_ROW_VALUES = Arrays.asList(
            REPORT_NUMBER_VAL,
            OPENED_DATE_VAL,
            OPENED_HOUR_VAL,
            OPENED_BY_VAL,
            MONITORING_DATE_VAL,
            TEST_CODE_VAL,
            LOCATION_VAL,
            ELEMENT_VAL,
            SIDE_VAL,
            SAFETY_FLAW_DETAILS_VAL,
            TELEPHONE_REPORTING_FOR_PROJECT_MANAGEMENT_VAL,
            TELEPHONE_REPORTING_OPERATIONS_TEAM_VAL,
            CORRECTIVE_ACTION_VAL,
            RESPONSIBLE_PERSON_VAL,
            DESCRIPTION_OF_PERFORMED_CORRECTIVE_ACTION_VAL,
            FINISH_DATE_VAL,
            END_TIME_VAL,
            REMARKS_VAL,
            STATUS_VAL
    );
}