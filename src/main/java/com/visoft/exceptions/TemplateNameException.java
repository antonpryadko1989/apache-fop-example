package com.visoft.exceptions;

import java.util.Map;

public class TemplateNameException extends RuntimeException{
    private String exception;

    public void setException(String exception) {
        this.exception = exception;
    }

    public TemplateNameException(String message, String exception){
        super(message);
        this.exception = exception;
    }

    public TemplateNameException(String message) {
        super(message);
    }

    public String getException (){
        return exception;
    }
}
