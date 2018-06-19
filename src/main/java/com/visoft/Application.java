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


    public static String object2XML(TemplateDTO temp) throws IOException, DocumentException {
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
        temp.setBody("dsaad");


//        JSONObject jsonObj = new JSONObject("{\"phonetype\":\"N95\",\"cat\":\"WP\",\"docs\":[{\"a\":\"ff\",\"bb\":854},{\"bb\":78},{\"a\":\"ff\",\"bb\":78}]}");
//        temp.setBody(jsonObj);
        JSONObject json = new JSONObject(temp);
        return XML.toString(json);

    }
}