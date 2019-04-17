package com.visoft.services;

import com.visoft.templates.entity.TemplateDTO;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.io.IOException;

public interface XLSXBuilder {

    StreamingResponseBody buildXLSX(TemplateDTO template) throws IOException;
}
