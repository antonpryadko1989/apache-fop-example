package com.visoft.exceptions;

import java.util.Map;

public class DeficienciesReportException extends RuntimeException{

    private String exception;
    private Map<String, String> validationException;

    public DeficienciesReportException(String message, String exception){
        super(message);
        this.exception = exception;
    }

    public DeficienciesReportException(String message, Map<String, String> validationException){
        super(message);
        this.validationException = validationException;
    }

    public String getException() {
        return exception;
    }

    public Map<String, String> getValidationException() {
        return validationException;
    }
}
