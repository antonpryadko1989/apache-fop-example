package com.visoft;

import com.sun.org.apache.xml.internal.serialize.XMLSerializer;
import com.visoft.example.TestObject;
import com.visoft.templates.entity.TemplateDTO;
import flexjson.JSON;
import flexjson.JSONSerializer;
import org.dom4j.DocumentException;
import org.json.JSONObject;
import org.json.XML;

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
//        temp.setBody(new TestObject("21","32"));


        JSONObject jsonObj = new JSONObject("{\"phonetype\":\"N95\",\"cat\":\"WP\",\"docs\":[{\"a\":\"ff\",\"bb\":854},{\"bb\":78},{\"a\":\"ff\",\"bb\":78}]}");
        temp.setBody(jsonObj);
//        POIExcel poi = new POIExcel();
//        System.out.println(poi.getTemplate(temp));
        System.out.println(json2XML(temp));
    }
    private static String json2XML(Object str) {
        JSONObject json = new JSONObject(str);
        String xml = XML.toString(json);
        return xml;
    }
}