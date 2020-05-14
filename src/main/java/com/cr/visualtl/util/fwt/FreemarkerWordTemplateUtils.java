package com.cr.visualtl.util.fwt;

import com.cr.visualtl.util.CrFileUtils;
import freemarker.template.Configuration;
import freemarker.template.Template;

import java.io.*;
import java.util.Map;
import java.util.Objects;
import java.util.regex.Matcher;

// Freemarker成文工具类  YJ_20200513
public class FreemarkerWordTemplateUtils {
    public static String word2ftl(Map<String, String> map, File docFile){
        String docText;
        docText = CrFileUtils.convertFile2text(docFile);
        // 解析自定义字段
        docText = Objects.requireNonNull(docText).replaceAll(FwtElement.WORD_BOOKMARK_CUSTOM_FIELD_PATTERN.pattern(), "<w:r><w:t>\\$\\{($2)!\\}</w:t></w:r>");
        // 解析系统自带字段
        docText = Objects.requireNonNull(docText).replaceAll(FwtElement.WORD_BOOKMARK_SYSTEM_FIELD_PATTERN.pattern(), "<w:r><w:t>\\$\\{($2)!\\}</w:t></w:r>");
        // 解析系统固化的Freemarker语法块
        StringBuffer docBuffer = new StringBuffer();
        Matcher blockMatcher = FwtElement.WORD_BOOKMARK_SYSTEM_BLOCK_PATTERN.matcher(Objects.requireNonNull(docText));
        int last = -1;
        while(blockMatcher.find()){
            String groupValue = blockMatcher.group(2);
            String blockValue = map.get(groupValue);
            blockMatcher.appendReplacement(docBuffer, blockValue);
            last = blockMatcher.end();
        }
        // 将替换后的部分补全
        if(docBuffer.length()>0 && last>0){
            docText = docBuffer.toString() + docText.substring(last);
        }
        return docText;
    }

    public static boolean exportSimpleWord(Map<?,?> dataMap, String templateDir, String templateFileName, String outPath){
        Configuration configuration = new Configuration(Configuration.DEFAULT_INCOMPATIBLE_IMPROVEMENTS);
        configuration.setDefaultEncoding("utf-8");
        try{
            configuration.setDirectoryForTemplateLoading(new File(templateDir));
            Template template = configuration.getTemplate(templateFileName);
            Writer outWriter = new OutputStreamWriter(new FileOutputStream(outPath));
            template.process(dataMap, outWriter);
            outWriter.flush();
            outWriter.close();
        }catch (Exception e){
            e.printStackTrace();
            return false;
        }
        return true;
    }
}
