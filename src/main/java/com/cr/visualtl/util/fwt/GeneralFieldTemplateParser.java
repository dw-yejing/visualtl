package com.cr.visualtl.util.fwt;

import lombok.extern.slf4j.Slf4j;

import java.util.List;
import java.util.Objects;
import java.util.regex.Matcher;

import static org.springframework.util.Assert.notNull;

@Slf4j
public class GeneralFieldTemplateParser implements TemplateParser {
    /**
     * 待解析的文本
     */
    private String text;

    /**
     * 存储待删除的书签
     */
    private List<String> removeTagList;

    GeneralFieldTemplateParser(String docText, List<String> removeTagList){
        this.text = docText;
        this.removeTagList = removeTagList;
    }

    @Override
    public String invoke() {
        String docText = text;
        Matcher generalFieldMatcher = FwtElement.WORD_BOOKMARK_GENERAL_FIELD_PATTERN.matcher(Objects.requireNonNull(docText));
        int last = -1;
        StringBuffer docBuffer = new StringBuffer();
        while(generalFieldMatcher.find()){
            String itemValue = generalFieldMatcher.group();
            String fieldValue = generalFieldMatcher.group(2);
            String itemId = generalFieldMatcher.group(1);
            itemValue += String.format("<w:r><w:t>\\$\\{(%s)!\\}</w:t></w:r>", fieldValue);
            generalFieldMatcher.appendReplacement(docBuffer, itemValue);
            last = generalFieldMatcher.end();
            String removeTag1 = String.format("<aml:annotation\\s*?aml:id=\"%s\"\\s*?w:type=\"Word.Bookmark.Start\"\\s*?w:name=\"%s\"\\s*?/>", itemId, fieldValue);
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
