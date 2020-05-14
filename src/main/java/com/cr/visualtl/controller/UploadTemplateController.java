package com.cr.visualtl.controller;

import com.cr.visualtl.model.ResponseTarget;
import com.cr.visualtl.service.UploadTemplate;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

@Controller
public class UploadTemplateController {
    @Autowired
    private UploadTemplate uploadTemplate;

    @RequestMapping("/upload")
    @ResponseBody
    public ResponseTarget uploadTemplate(@RequestParam("templateFile") MultipartFile templateFile){
        return uploadTemplate.upload(templateFile);
    }
}
