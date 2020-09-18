package com.cr.visualtl.util.fwt;

import lombok.extern.slf4j.Slf4j;
import org.springframework.util.Assert;

import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.regex.Matcher;

import static org.springframework.util.Assert.notNull;

@Slf4j
public class BlockTemplateParser implements TemplateParser {
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

    BlockTemplateParser(String docText, List<String> removeTagList, Map<String, String> bookmarkMap){
        this.text = docText;
        this.removeTagList = removeTagList;
        this.bookmarkMap = bookmarkMap;
    }

    @Override
    public String invoke() {
        String docText = text;
        StringBuffer docBuffer = new StringBuffer();
        Matcher blockMatcher = FwtElement.WORD_BOOKMARK_SYSTEM_BLOCK_PATTERN.matcher(Objects.requireNonNull(docText));
        int last = -1;
        while(blockMatcher.find()){
            String itemValue = blockMatcher.group();
            String itemId = blockMatcher.group(1);
            String fieldValue = blockMatcher.group(2);
            fieldValue = fieldValue.substring(FwtElement.CR_BLOCK_.length());
            String blockValue = bookmarkMap.get(fieldValue);
            itemValue += blockValue;
            blockMatcher.appendReplacement(docBuffer, itemValue);
            last = blockMatcher.end();
            String removeTag1 = String.format("<aml:annotation\\s*?aml:id=\"%s\"\\s*?w:type=\"Word.Bookmark.Start\"\\s*?w:name=\"%s%s\"\\s*?/>", itemId, FwtElement.CR_BLOCK_, fieldValue);
            String removeTag2 = String.format("<aml:annotation\\s*?aml:id=\"%s\"\\s+?w:type=\"Word.Bookmark.End\"\\s*/>", itemId);
            removeTagList.add(removeTag1);
            removeTagList.add(removeTag2);
        }
        // 将替换后的部分补全
        if(docBuffer.length()>0 && last>0){
            docText = docBuffer.toString() + docText.substring(last);
        }
        Assert.notNull(docText, "成文模板解析异常");
        // 删除已解析标签，避免重复解析
        docText = TemplateParser.removeParsedTag(docText, removeTagList);
        notNull(docText, "成文模板解析异常");
        return docText;
    }
}
