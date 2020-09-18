package com.cr.visualtl.util.fwt;

import lombok.extern.slf4j.Slf4j;

import java.util.regex.Matcher;

import static org.springframework.util.Assert.notNull;

@Slf4j
public class SeparatorTemplateParser implements TemplateParser {
    /**
     * 待解析的文本
     */
    private String text;


    SeparatorTemplateParser(String docText){
        this.text = docText;
    }

    @Override
    public String invoke() {
        String docText = text;
        Matcher separatorMatcher = FwtElement.OBJECT_SEPARATOR_PATTERN.matcher(docText);
        int last = -1;
        StringBuffer docBuffer = new StringBuffer();
        while (separatorMatcher.find()){
            String itemValue = separatorMatcher.group();
            itemValue = itemValue.replace("_", ".");
            separatorMatcher.appendReplacement(docBuffer, itemValue);
            last = separatorMatcher.end();
        }
        // 将替换后的部分补全
        if(docBuffer.length()>0 && last>0){
            docText = docBuffer.toString() + docText.substring(last);
        }
        notNull(docText, "成文模板解析异常");
        return docText;
    }
}
