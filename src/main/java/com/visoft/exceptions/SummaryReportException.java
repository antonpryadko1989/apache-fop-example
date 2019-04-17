package com.visoft.exceptions;

import java.util.Map;

public class SummaryReportException extends RuntimeException{
    private String exception;
    private Map<String, String> validationException;

    public SummaryReportException(String message, String exception){
        super(message);
        this.exception = exception;
    }

    public SummaryReportException(String message, Map<String, String> validationException){
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
