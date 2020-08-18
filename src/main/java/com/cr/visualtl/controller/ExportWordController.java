package com.cr.visualtl.controller;

import com.cr.visualtl.model.ResponseTarget;
import com.cr.visualtl.service.ExportWord;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

@Controller
public class ExportWordController {
    @Autowired
    private ExportWord exportWord;

    @RequestMapping("/export")
    @ResponseBody
    public ResponseTarget export() throws Exception{
        return exportWord.exportWord();
    }
}
