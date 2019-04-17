package com.visoft.exceptions;

public class PreviewLogoException  extends RuntimeException{

    private int exception;

    public PreviewLogoException(String message, int exception){
        super(message);
        this.exception = exception;
    }

    public PreviewLogoException(String message) {
        super(message);
    }

    public int getException (){
        return exception;
    }
}
