package com.visoft.services.impl;

import org.apache.commons.collections4.map.LinkedMap;

import java.util.Map;

class ReportBodyConverter {

    static LinkedMap<String, String> convertReportBodyToReportBodyRows(Map<String, String> currentMap,
                                                                       String[] keys,
                                                                       String[] values) {
        LinkedMap<String, String> result = new LinkedMap<>();
        for (int i = 0; i < keys.length; i++) {
            result.put(keys[i], currentMap.get(values[i]));
        }
        return result;
    }
}