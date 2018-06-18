package com.visoft;

import com.visoft.example.POIExcel;
import com.visoft.templates.entity.TemplateDTO;
import org.dom4j.DocumentException;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class Application {


    public static void main(String[] args) throws IOException, DocumentException {
        TemplateDTO temp = new TemplateDTO();
        Map<String, String> bodyMap = new HashMap<>();
        temp.setProjectId("1");
        temp.setTemplateName("TEMP_1.xlsx");
        temp.setOutPutName("TEMP_1.xlsx");
        bodyMap.put("1", "01");
        bodyMap.put("2", "22");
        bodyMap.put("3", "333");
        bodyMap.put("4", "4444");
        bodyMap.put("5", "55555");
        temp. setTemplateBody(bodyMap);


        POIExcel poi = new POIExcel();
        System.out.println(poi.getTemplate(temp));
    }

}
