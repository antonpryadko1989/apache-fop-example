package com.visoft.api;

import com.visoft.exceptions.PathValidationException;
import com.visoft.exceptions.TemplateValidationException;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.servlet.mvc.method.annotation.ResponseEntityExceptionHandler;

import java.util.HashMap;
import java.util.Map;

@ControllerAdvice
public class ErrorController extends ResponseEntityExceptionHandler {

    @ExceptionHandler(TemplateValidationException.class)
    public ResponseEntity<Object> validationError(TemplateValidationException temp){
        Map<String, Map<String, String>> errorMessage = new HashMap<>();
        errorMessage.put(temp.getMessage(), temp.getException());
        return new ResponseEntity<>(errorMessage, new HttpHeaders(), HttpStatus.BAD_REQUEST);
    }

    @ExceptionHandler(PathValidationException.class)
    public ResponseEntity<Object> validationError(PathValidationException temp){
        Map<String, String> errorMessage = new HashMap<>();
        errorMessage.put(temp.getMessage(), temp.getException());
        return new ResponseEntity<>(errorMessage, new HttpHeaders(), HttpStatus.BAD_REQUEST);
    }
}
