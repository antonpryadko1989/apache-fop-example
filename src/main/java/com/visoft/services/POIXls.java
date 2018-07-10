package com.visoft.services;

import com.visoft.exceptions.FileConvertException;
import com.visoft.exceptions.PathValidationException;
import com.visoft.templates.entity.TemplateBody;
import com.visoft.templates.entity.TemplateDTO;
import com.visoft.templates.regulars.TemplateValue;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.tika.metadata.Metadata;
import org.apache.tika.mime.MediaType;
import org.apache.tika.parser.AutoDetectParser;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.io.*;
import java.nio.file.*;
import java.util.*;

@Service
public class POIXls {


    @Value("${templates.repository}")
    private String templatesRepository;

    @Value("${app.XLSX}")
    private String appXLSX;

    @Value("${app.XLS}")
    private String appXLS;

    @Autowired
    private Input2OutputService resultService;

    private static final String TEMPLATES = "E:/WORK/TEMPLATES/";

    private static final String NCR_TEMPLATE_NAME = "טופס אי התאמה";
    private static final String VERSION = "Version / מהדורה";
    private static final String DATE = "Date / תאריך";
    private static final String MAIN_CONTRACTOR = "Main Contractor /  קבלן ראשי";
    private static final String PROJECT_NAME = "Project Name / שם הפרויקט";
    private static final String MANAGEMENT_COMPANY  = "Management Company / חברת ניהול";
    private static final String CONTRACT_NO = "Contract No. / חוזה מסי";
    private static final String QC_COMPANY = "QC Company / חברת בקרת איכות";
    private static final String QA_COMPANY = "QA Company / חברת הבטחת איכות";
    private static final String NCR_NUMBER = "אי התאמה מס' / NCR Number";
    private static final String NCR_OPENED= "NCR opened / נפתח ע\"י";
    private static final String POSITION = "Position / תפקיד";
    private static final String NAME = "Name / שם";
    private static final String DATE_OF_NCR = "Date of NCR / תאריך הפתיחה";
    private static final String QA_QC = "QA / QC";
    private static final String QCM = "QCM";
    private static final String QAM = "QAM";
    private static final String DETAILS_OF_STRUCTURE = "Details of Structure / פרטי המבנה";
    private static final String ELEMENT_STATION_ROAD_TUNNEL_BRIDGE = "Element (Station, Road, Tunnel, Bridge) / (כביש , רמפה , גשר, מנהרה )";
    private static final String STRUCTURE = "Structure / מבנה";
    private static final String ELEMENT = "Element / אלמנט";
    private static final String FROM_SECTION = "From section / מחתך";
    private static final String TILL_SECTION = "Till section / עד חתך";
    private static final String SIDE = "Side / הסט";
    private static final String NCR_LEVEL = "NCR Level / דרגה";
    private static final String NUMBER_OF_DAYS_LATE = "Number Of Days Late/ מס ימי עיכוב לסגירה";

    private static final String EXPECTED_CLOSING_DATE = "Expected Closing Date / תאריך סגירה משוער";
    private static final String UPDATED_EXPECTED_CLOSING_DATE = "Updated Expected Closing Date / תאריך משוער לסגירה מקורי";
    private static final String SUB_PROJECT = "Sub Project / תת פרויקט";
    private static final String NCR_DESCRIPTION = "NCR description / תאור אי ההתאמה";
    private static final String RESPONSIBLE_PARTY = "Responsible Party (Design ,Contractor,Supplier) / גורם אחראי לליקוי (תכנון, ביצוע, ספק)";
    private static final String CORRECTIVE_ACTION = "Corrective action / פעולה מתקנת נדרשת";
    private static final String DESCRIPTION_OF_PERFORMED_CORRECTIVE_ACTION = "Description of performed corrective action / פעולה מתקנת שבוצעה";
    private static final String RESPONSIBLE_PERSON = "Responsible person / שם האחראי";
    private static final String REMARKS = "Remarks / הערות";


    private static final String TEMPLATE_NAME_VAL = "templateName";
    private static final String VERSION_VAL = "version";
    private static final String DATE_VAL = "date";
    private static final String MAIN_CONTRACTOR_VAL = "mainContractor";
    private static final String PROJECT_NAME_VAL = "projectName";
    private static final String MANAGEMENT_COMPANY_VAL  = "managementCompany";
    private static final String CONTRACT_NO_VAL = "contractNo";
    private static final String QC_COMPANY_VAL = "qcCompany";
    private static final String QA_COMPANY_VAL = "qaCompany";
    private static final String NCR_NUMBER_VAL = "ncrNumber";
    private static final String QC_NAME_VAL = "ncrOpenQCName";
    private static final String QC_DATE_OF_NCR_VAL = "ncrOpenQCDate";
    private static final String QA_NAME_VAL = "ncrOpenQAName";
    private static final String QA_DATE_OF_NCR_VAL = "ncrOpenQADate";
    private static final String ELEMENT_STATION_ROAD_TUNNEL_BRIDGE_VAL = "elementStationRoadTunnelBridge";
    private static final String STRUCTURE_VAL = "structure";
    private static final String ELEMENT_VAL = "element";
    private static final String FROM_SECTION_VAL = "fromSection";
    private static final String TILL_SECTION_VAL = "tillSection";
    private static final String SIDE_VAL = "side";
    private static final String NCR_LEVEL_VAL = "ncrLevel";
    private static final String NUMBER_OF_DAYS_LATE_VAL = "numberOfDaysLate";

    private static final String EXPECTED_CLOSING_DATE_VAL = "expectedClosingDate";
    private static final String UPDATED_EXPECTED_CLOSING_DATE_VAL = "updatedExpectedClosingDate";
    private static final String SUB_PROJECT_VAL = "subProject";
    private static final String NCR_DESCRIPTION_VAL = "ncrDescription";
    private static final String RESPONSIBLE_PARTY_VAL = "responsibleParty";
    private static final String CORRECTIVE_ACTION_VAL = "correctiveAction";
    private static final String DESCRIPTION_OF_PERFORMED_CORRECTIVE_ACTION_VAL = "descriptionOfPerformedCorrectiveAction";
    private static final String RESPONSIBLE_PERSON_VAL = "responsiblePerson";
    private static final String REMARKS_VAL = "remarks";

    public static void main(String[] args) {
        TemplateDTO template = getDefaultTemplate();
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet(template.getOutPutName());
        setColumnWidths(sheet);
        int rowNum = 2;
        Row row = sheet.createRow(rowNum);
        addHeaderInfo(sheet, rowNum, NCR_TEMPLATE_NAME, VERSION, DATE);
        rowNum++;
        addHeaderInfo(sheet, rowNum, template.getBody().getBodyElements().get(TEMPLATE_NAME_VAL),
                template.getBody().getBodyElements().get(VERSION_VAL),
                template.getBody().getBodyElements().get(DATE_VAL));
        rowNum = rowNum + 2;

        setValuesToRow(sheet,row, rowNum, MAIN_CONTRACTOR, template.getBody().getBodyElements().get(MAIN_CONTRACTOR_VAL),
                PROJECT_NAME, template.getBody().getBodyElements().get(PROJECT_NAME_VAL));
        rowNum++;
        setValuesToRow(sheet, row, rowNum, MANAGEMENT_COMPANY, template.getBody().getBodyElements().get(MANAGEMENT_COMPANY_VAL),
                CONTRACT_NO, template.getBody().getBodyElements().get(CONTRACT_NO_VAL));
        rowNum++;
        setValuesToRow(sheet, row, rowNum, QC_COMPANY, template.getBody().getBodyElements().get(QC_COMPANY_VAL),
                QA_COMPANY, template.getBody().getBodyElements().get(QA_COMPANY_VAL));
        rowNum = rowNum + 2;

        row = sheet.createRow(rowNum);
        mergeCells(sheet, rowNum, rowNum, 1, 4);
        mergeCells(sheet, rowNum, rowNum, 5, 8);
        row.createCell(1).setCellValue(NCR_NUMBER);
        row.createCell(5).setCellValue(template.getBody().getBodyElements().get(NCR_NUMBER_VAL));
        rowNum = rowNum + 2;
        addQaQc(sheet, rowNum, template.getBody());
        rowNum = rowNum + 3;
        row = sheet.createRow(rowNum);
        mergeCells(sheet, rowNum, rowNum, 1, 8);
        row.createCell(1).setCellValue(DETAILS_OF_STRUCTURE);
        rowNum++;
        row = sheet.createRow(rowNum);
        String[] rowArrayHeader = {ELEMENT_STATION_ROAD_TUNNEL_BRIDGE, STRUCTURE, ELEMENT,
                FROM_SECTION, TILL_SECTION, SIDE, NCR_LEVEL, NUMBER_OF_DAYS_LATE};
        setValuesToRow(sheet, row, rowNum, rowArrayHeader);
        rowNum++;
        row = sheet.createRow(rowNum);
        String[] rowArrayBody = {template.getBody().getBodyElements().get(ELEMENT_STATION_ROAD_TUNNEL_BRIDGE_VAL),
                template.getBody().getBodyElements().get(STRUCTURE_VAL),
                template.getBody().getBodyElements().get(ELEMENT_VAL),
                template.getBody().getBodyElements().get(FROM_SECTION_VAL),
                template.getBody().getBodyElements().get(TILL_SECTION_VAL),
                template.getBody().getBodyElements().get(SIDE_VAL),
                template.getBody().getBodyElements().get(NCR_LEVEL_VAL),
                template.getBody().getBodyElements().get(NUMBER_OF_DAYS_LATE_VAL)};
        setValuesToRow(sheet, row, rowNum, rowArrayBody);
        rowNum = rowNum + 2;
        row = sheet.createRow(rowNum);
        setValuesToRow(sheet, row, rowNum, EXPECTED_CLOSING_DATE, template.getBody().getBodyElements().get(EXPECTED_CLOSING_DATE_VAL));
        rowNum++;
        row = sheet.createRow(rowNum);
        setValuesToRow(sheet, row, rowNum, UPDATED_EXPECTED_CLOSING_DATE, template.getBody().getBodyElements().get(UPDATED_EXPECTED_CLOSING_DATE_VAL ));
        rowNum++;row = sheet.createRow(rowNum);
        setValuesToRow(sheet, row, rowNum, SUB_PROJECT, template.getBody().getBodyElements().get(SUB_PROJECT_VAL));
        rowNum++;
        row = sheet.createRow(rowNum);
        setValuesToRow(sheet, row, rowNum, NCR_DESCRIPTION, template.getBody().getBodyElements().get(NCR_DESCRIPTION_VAL));
        rowNum++;
        row = sheet.createRow(rowNum);
        setValuesToRow(sheet, row, rowNum, RESPONSIBLE_PARTY, template.getBody().getBodyElements().get(RESPONSIBLE_PARTY_VAL));
        rowNum++;
        row = sheet.createRow(rowNum);
        setValuesToRow(sheet, row, rowNum, CORRECTIVE_ACTION, template.getBody().getBodyElements().get(CORRECTIVE_ACTION_VAL));
        rowNum++;
        row = sheet.createRow(rowNum);
        setValuesToRow(sheet, row, rowNum, DESCRIPTION_OF_PERFORMED_CORRECTIVE_ACTION, template.getBody().getBodyElements().get(DESCRIPTION_OF_PERFORMED_CORRECTIVE_ACTION_VAL));
        rowNum++;
        row = sheet.createRow(rowNum);
        setValuesToRow(sheet, row, rowNum, RESPONSIBLE_PERSON, template.getBody().getBodyElements().get(RESPONSIBLE_PERSON_VAL));
        rowNum++;
        row = sheet.createRow(rowNum);
        setValuesToRow(sheet, row, rowNum, REMARKS, template.getBody().getBodyElements().get(REMARKS_VAL));
        rowNum++;


        row = sheet.createRow(rowNum+1);
        row.createCell(1).setCellValue("Hello Word!");
        try {
            FileOutputStream fileOut = new FileOutputStream(TEMPLATES + "XLSX/" + System.currentTimeMillis() +".xlsx");
            workbook.write(fileOut);
            fileOut.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("OK");
    }

    private static Sheet setValuesToRow(Sheet sheet, Row row, int rowNum, String[] rowArray) {
        for(int i = 0; i < rowArray.length; i++){
            row.createCell(i + 1).setCellValue(rowArray[i]);
        }
        return sheet;
    }

    private static Sheet addQaQc(Sheet sheet, int rowNum, TemplateBody body) {
        Row row = sheet.createRow(rowNum);
        setValuesToRow(sheet, row, rowNum, NCR_OPENED, POSITION, NAME, DATE_OF_NCR);
        rowNum++;
        row = sheet.createRow(rowNum);
        mergeCells(sheet, rowNum, rowNum + 1, 1, 2);
        row.createCell(1).setCellValue(QA_QC);
        setValuesToRow(sheet, row, rowNum, QCM, body.getBodyElements().get(QC_NAME_VAL), body.getBodyElements().get(QC_DATE_OF_NCR_VAL));
        rowNum++;
        row = sheet.createRow(rowNum);
        setValuesToRow(sheet, row, rowNum, QAM, body.getBodyElements().get(QA_NAME_VAL), body.getBodyElements().get(QA_DATE_OF_NCR_VAL));
        return sheet;
    }

    private static Sheet setValuesToRow(Sheet sheet, Row row, int rowNum, String first,String second, String third) {
        mergeCells(sheet, rowNum, rowNum, 3, 4);
        row.createCell(3).setCellValue(first);
        mergeCells(sheet, rowNum, rowNum, 5, 6);
        row.createCell(5).setCellValue(second);
        mergeCells(sheet, rowNum, rowNum, 7, 8);
        row.createCell(7).setCellValue(third);
        return sheet;
    }

    private static Sheet setValuesToRow(Sheet sheet,Row row, int rowNum, String first, String second) {
        mergeCells(sheet, rowNum, rowNum, 1, 2);
        row.createCell(1).setCellValue(first);
        mergeCells(sheet, rowNum, rowNum, 3, 8);
        row.createCell(3).setCellValue(second);
        return sheet;
    }

    private static Sheet addHeaderInfo(Sheet sheet, int rowNum, String bg, String h, String i) {
        Row row = sheet.createRow(rowNum);
        mergeCells(sheet, rowNum, rowNum, 1, 6);
        row.createCell(1).setCellValue(bg);
        row.createCell(7).setCellValue(h);
        row.createCell(8).setCellValue(i);
        return sheet;
    }

    private static Sheet setValuesToRow(Sheet sheet,Row row, int rowNum, String first, String second, String third, String fourth) {
        mergeCells(sheet, rowNum, rowNum, 1, 2);
        row.createCell(1).setCellValue(first);
        mergeCells(sheet, rowNum, rowNum, 3, 4);
        row.createCell(3).setCellValue(second);
        mergeCells(sheet, rowNum, rowNum, 5, 6);
        row.createCell(5).setCellValue(third);
        mergeCells(sheet, rowNum, rowNum, 7, 8);
        row.createCell(7).setCellValue(fourth);
        return sheet;
    }

    private static Sheet mergeCells(Sheet sheet, int fromFirstRow, int toLastRow , int fromFirstColumn, int toLastColumn) {
        sheet.addMergedRegion(new CellRangeAddress(
                fromFirstRow, toLastRow, fromFirstColumn, toLastColumn
        ));
        return sheet;
    }


    private static TemplateDTO getDefaultTemplate() {
        TemplateDTO template = new TemplateDTO();
        template.setOutPutName("Bla Bla");
        TemplateBody body = setDefaultBody();
        template.setBody(body);
        return template;
    }

    private static TemplateBody setDefaultBody() {
        TemplateBody body = new TemplateBody();
        body.setBodyElements(setDefaultBodyElements());
        body.setBodyLists(setDefaultBodyLists());
        return body;
    }

    private static Map<String,List<Map<String,String>>> setDefaultBodyLists() {
        Map<String, List<Map<String, String>>> bodyLists = new HashMap<>();
        bodyLists.put("additionalDocuments", setDefaultAdditionalDocuments());
        return bodyLists;
    }


    private static List<Map<String,String>> setDefaultAdditionalDocuments() {
        List<Map<String, String>> additionalDocuments = new ArrayList<>();
        for(int i = 0; i < 5; i++){
            Map<String, String> addDoc1 = new HashMap<>();
            addDoc1.put("item", "i" + i);
            addDoc1.put("exists", "e" + i);
            addDoc1.put("certificateNo", "cerNo" + i);
            addDoc1.put("expiration", "exp" + i);
            addDoc1.put("attachedDocuments", "Doc" + i);
            additionalDocuments.add(addDoc1);
        }
        return  additionalDocuments;
    }

    private static Map<String,String> setDefaultBodyElements() {
        Map<String,String> bodyElements = new HashMap<>();

        bodyElements.put("templateName", "templateName");
        bodyElements.put("version", "version");
        bodyElements.put("date", "date");
        bodyElements.put("mainContractor", "mainContractor");
        bodyElements.put("managementCompany", "managementCompany");
        bodyElements.put("qcCompany", "qcCompany");
        bodyElements.put("projectName", "projectName");
        bodyElements.put("contractNo", "contractNo");
        bodyElements.put("qaCompany", "qaCompany");
        bodyElements.put("ncrNumber", "ncrNumber");
        bodyElements.put("ncrOpenQCName", "ncrOpenQCName");
        bodyElements.put("ncrOpenQCDate", "ncrOpenQCDate");
        bodyElements.put("ncrOpenQAName", "ncrOpenQAName");
        bodyElements.put("ncrOpenQADate", "ncrOpenQADate");
        bodyElements.put("elementStationRoadTunnelBridge", "elementStationRoadTunnelBridge");
        bodyElements.put("structure", "structure");
        bodyElements.put("element", "element");
        bodyElements.put("fromSection", "fromSection");
        bodyElements.put("tillSection", "tillSection");
        bodyElements.put("side", "side");
        bodyElements.put("ncrLevel", "ncrLevel");
        bodyElements.put("numberOfDaysLate", "numberOfDaysLate");
        bodyElements.put("expectedClosingDate", "expectedClosingDate");
        bodyElements.put("updatedExpectedClosingDate", "updatedExpectedClosingDate");
        bodyElements.put("subProject", "subProject");
        bodyElements.put("ncrDescription", "ncrDescription");
        bodyElements.put("responsibleParty", "responsibleParty");
        bodyElements.put("correctiveAction", "correctiveAction");
        bodyElements.put("descriptionOfPerformedCorrectiveAction", "descriptionOfPerformedCorrectiveAction");
        bodyElements.put("responsiblePerson", "responsiblePerson");
        bodyElements.put("remarks", "remarks");
        return bodyElements;
    }

    private static Sheet setColumnWidths(Sheet sheet) {
        sheet.setColumnWidth(0, 21*36);
        sheet.setColumnWidth(1,152*36);
        sheet.setColumnWidth(2,152*36);
        sheet.setColumnWidth(3,136*36);
        sheet.setColumnWidth(4,136*36);
        sheet.setColumnWidth(5,152*36);
        sheet.setColumnWidth(6,124*36);
        sheet.setColumnWidth(7,144*36);
        sheet.setColumnWidth(8,96*36);

        return sheet;
    }

    StreamingResponseBody getExcelFile(TemplateDTO template) throws IOException {
//        File f = Paths.get(templatesRepository, template.getProjectId(), template.getTemplateName()).toFile();
//        if (f.isDirectory()) {
//            throw new PathValidationException("error", "It's directory");
//        }
//        if (!f.exists()) {
//            throw new PathValidationException("error", "File not exist");
//        }
//        Workbook workbook = getWorkbook(Paths.get(templatesRepository, template.getProjectId(), template.getTemplateName()).toString());
//        Sheet sheet = workbook.getSheetAt(0);
//        for (Row row : sheet) {
//            Iterator<Cell> cellIterator = row.cellIterator();
//            while (cellIterator.hasNext()) {
//                Cell cell = cellIterator.next();
//                if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
//                    String newCellValueKey = TemplateValue.isValue(cell.getStringCellValue());
//                    if (newCellValueKey.length() > 0) {
//                        cell.setCellValue(template.getTemplateBody().get(newCellValueKey));
//                    }
//                }
//            }
//        }
//        ByteArrayOutputStream out = new ByteArrayOutputStream();
//        workbook.write(out);
//        InputStream inputStream = new ByteArrayInputStream(out.toByteArray());
//        return resultService.getOutput(inputStream);
        return null;
    }

    private Workbook getWorkbook(String filePath)
            throws IOException {
        InputStream inputFile = new FileInputStream(filePath);
        byte[] bytes = Files.readAllBytes(Paths.get(filePath));
        final ByteArrayInputStream is = new ByteArrayInputStream(bytes);
        final MediaType mediaType = new AutoDetectParser().getDetector().detect(is, new Metadata());
        Workbook workbook = null;
        if (mediaType.toString().equals(appXLSX)) {
            workbook = new XSSFWorkbook(inputFile);
            return  workbook;
        } else if (mediaType.toString().equals(appXLS)) {
            workbook = new HSSFWorkbook(inputFile);
            return workbook;
        } else {
            throw new FileConvertException("The specified file is not Excel file");
        }
    }
}