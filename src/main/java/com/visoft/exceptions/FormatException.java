package com.visoft.exceptions;

public class FormatException  extends RuntimeException{
    private String exception;

    public FormatException(String message, String exception){
        super(message);
        this.exception = exception;
    }

    public FormatException(String message) {
        super(message);
    }

    public String getException() {
        return exception;
    }
}