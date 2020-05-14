package com.cr.visualtl.service.impl;

import com.cr.visualtl.config.Constant;
import com.cr.visualtl.model.ResponseTarget;
import com.cr.visualtl.service.UploadTemplate;
import com.cr.visualtl.util.CrFileUtils;
import com.cr.visualtl.util.fwt.FreemarkerWordTemplateUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import java.io.*;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;

@Service
public class UploadTemplateImpl implements UploadTemplate {
    @Autowired
    private HttpServletRequest httpServletRequest;

    @Override
    public ResponseTarget upload(MultipartFile templateFile) {
        // 存储用户模板
        String templateFilePath = Paths.get(Constant.TEMPLATEFOLDER, templateFile.getOriginalFilename()).toString();
        String ftlTemplateFilePath = templateFilePath.substring(0, templateFilePath.lastIndexOf(46))+".ftl";
        String ftlTemplateFileName = ftlTemplateFilePath.substring(ftlTemplateFilePath.lastIndexOf(File.separator)+1);
        httpServletRequest.getSession().setAttribute("ftlTemplateFileName", ftlTemplateFileName);
        httpServletRequest.getSession().setAttribute("templateFileName", templateFile.getOriginalFilename());
        try {
            templateFile.transferTo(new File(templateFilePath));
        } catch (IOException e) {
            e.printStackTrace();
        }
        Map<String, String> dataMap = new HashMap<>();
        dataMap.put("test1", "<w:r><w:t><#if (1>0)>\\${(scoreList_item2)!}</#if></w:t></w:r>");
        File  docTemplate = new File(templateFilePath);
        // 解析为Freemarker模板
        String docText;
        docText = FreemarkerWordTemplateUtils.word2ftl(dataMap, docTemplate);
        // 存储Freemarker模板
        boolean converted = CrFileUtils.convertText2file(docText, ftlTemplateFilePath);
        if(converted){
            return new ResponseTarget(true, "上传成功");
        }
        return new ResponseTarget(false, "上传失败");
    }
}
