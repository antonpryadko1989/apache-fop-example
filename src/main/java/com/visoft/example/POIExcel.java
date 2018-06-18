package com.visoft.example;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Date;
import java.util.Iterator;

import com.visoft.templates.entity.TemplateDTO;
import com.visoft.templates.regulars.TemplateValue;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.tika.metadata.Metadata;
import org.apache.tika.mime.MediaType;
import org.apache.tika.parser.AutoDetectParser;
import org.dom4j.DocumentException;

public class POIExcel {

    private static final String APP_XLS = "application/x-tika-msoffice";
    private static final String APP_XLSX = "application/x-tika-ooxml";
    private static final String PATH_TO_TEMPLATE_REPOSITORY = "E:\\WORK\\TEMPLATES";
    private static final String PATH_TO_TEMP_TEMPLATE_REPOSITORY = "E:\\WORK\\1\\1";


    public String getTemplate(TemplateDTO template) throws IOException, DocumentException {
        String filePath = PATH_TO_TEMPLATE_REPOSITORY + File.separator + template.getProjectId() + File.separator + template.getTemplateName();
        byte[] bytes = Files.readAllBytes(Paths.get(filePath));
        final ByteArrayInputStream is = new ByteArrayInputStream(bytes);
        final MediaType mediaType = new AutoDetectParser().getDetector().detect(is, new Metadata());
        if (mediaType.toString().equals(APP_XLS)) {
            System.out.println("xls");
            return setValuesIntoXLSTamplate(template);
        } else if (mediaType.toString().equals(APP_XLSX)) {
            System.out.println("xlsx");
            return setValuesIntoXLSXTamplate(template);
        } else {
            return null;
        }
    }

    private static String setValuesIntoXLSXTamplate(TemplateDTO template) throws IOException {
        String filePath = PATH_TO_TEMPLATE_REPOSITORY + File.separator + template.getProjectId() + File.separator + template.getTemplateName();
        InputStream inputFile = new FileInputStream(filePath);
        XSSFWorkbook wb = new XSSFWorkbook(inputFile);
        XSSFSheet sheet = wb.getSheetAt(0);
        Iterator<Row> rowIterator = sheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Iterator <Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                if (cell.getCellType() == Cell.CELL_TYPE_STRING){
                    String newCellValueKey = TemplateValue.isValue(cell.getStringCellValue());
                    if(newCellValueKey.length() > 0) {
                        System.out.println(template.getTemplateBody().get(newCellValueKey));
                        cell.setCellValue(template.getTemplateBody().get(newCellValueKey));
                    }
                }
            }
        }
        String fileName = "E:\\WORK\\TEMP\\" + new Date().getTime() + "_" + template.getOutPutName() ;
        FileOutputStream out = new FileOutputStream(new File(fileName));
        wb.write(out);
        out.close();
        return fileName;
    }

    private static String setValuesIntoXLSTamplate(TemplateDTO template) throws IOException {
        String filePath = PATH_TO_TEMPLATE_REPOSITORY + File.separator + template.getProjectId() + File.separator + template.getTemplateName();
        InputStream inputFile = new FileInputStream(filePath);
        HSSFWorkbook wb = new HSSFWorkbook(inputFile);
        HSSFSheet sheet = wb.getSheetAt(0);
        Iterator<Row> rowIterator = sheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Iterator <Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                if (cell.getCellType() == Cell.CELL_TYPE_STRING){
                    String newCellValueKey = TemplateValue.isValue(cell.getStringCellValue());
                    if(newCellValueKey.length() > 0) {
                        System.out.println(template.getTemplateBody().get(newCellValueKey));
                        cell.setCellValue(template.getTemplateBody().get(newCellValueKey));
                    }
                }
            }
        }
        String fileName = PATH_TO_TEMP_TEMPLATE_REPOSITORY + new Date().getTime() + "_" + template.getOutPutName() ;
        FileOutputStream out = new FileOutputStream(new File(fileName));
        wb.write(out);
        wb.write(out);
        out.close();
        return fileName;
    }
}
