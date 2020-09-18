package com.cr.visualtl.util.fwt;

import com.cr.visualtl.util.StringUtils;
import lombok.extern.slf4j.Slf4j;

import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.regex.Matcher;

import static org.springframework.util.Assert.notNull;

@Slf4j
public class DateTemplateParser implements TemplateParser {
    /**
     * 待解析的文本
     */
    private String text;

    /**
     * 存储待删除的书签
     */
    private List<String> removeTagList;

    /**
     * 书签字典
     */
    private Map<String, String> bookmarkMap;

    DateTemplateParser(String docText, List<String> removeTagList, Map<String, String> bookmarkMap){
        this.text = docText;
        this.removeTagList = removeTagList;
        this.bookmarkMap = bookmarkMap;
    }

    @Override
    public String invoke() {
        String docText = text;
        Matcher dateMatcher = FwtElement.WORD_DATE_PATTERN.matcher(Objects.requireNonNull(docText));
        int last = -1;
        StringBuffer docBuffer = new StringBuffer();
        while(dateMatcher.find()){
            String itemValue = dateMatcher.group();
            String itemId = dateMatcher.group(1);
            String fieldValue = dateMatcher.group(2);
            fieldValue = fieldValue.substring(FwtElement.CR_DATE_.length());
            int firstUnderscoreIndex = fieldValue.indexOf("_");
            String format = fieldValue.substring(0,firstUnderscoreIndex);
            fieldValue = fieldValue.substring(firstUnderscoreIndex+1);
            if(Objects.equals(format, FwtElement.DEFAULT_DATE_FORMAT)){
                fieldValue = String.format("%s?string(\"%s\")", fieldValue, FwtElement.DEFAULT_DATE_PATTERN);
            }else{
                String datePattern = bookmarkMap.get(format);
                if(StringUtils.isNotBlank(datePattern)){
                    fieldValue = String.format("%s?string(\"%s\")", fieldValue, datePattern);
                }
            }
            itemValue += String.format("<w:r><w:t>\\$\\{%s\\}</w:t></w:r>", fieldValue);
            dateMatcher.appendReplacement(docBuffer, itemValue);
            last = dateMatcher.end();
            String removeTag1 = String.format("<aml:annotation\\s*?aml:id=\"%s\"\\s*?w:type=\"Word.Bookmark.Start\"\\s*?w:name=\"%s\"\\s*?/>", itemId, dateMatcher.group(2));
            String removeTag2 = String.format("<aml:annotation\\s*?aml:id=\"%s\"\\s+?w:type=\"Word.Bookmark.End\"\\s*/>", itemId);
            removeTagList.add(removeTag1);
            removeTagList.add(removeTag2);
        }
        // 将替换后的部分补全
        if(docBuffer.length()>0 && last>0){
            docText = docBuffer.toString() + docText.substring(last);
        }
        // 删除已解析标签，避免重复解析
        docText = TemplateParser.removeParsedTag(docText, removeTagList);
        notNull(docText, "成文模板解析异常");
        return docText;
    }
}
