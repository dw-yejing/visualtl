package com.cr.visualtl.util.fwt;

import com.alibaba.fastjson.JSON;
import lombok.extern.slf4j.Slf4j;
import org.dom4j.Document;
import org.dom4j.DocumentHelper;
import org.dom4j.Element;
import org.dom4j.Node;
import org.springframework.util.Assert;

import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;

import static org.springframework.util.Assert.notNull;

@Slf4j
public class TrTemplateParser implements TemplateParser {
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

    TrTemplateParser(String docText, List<String> removeTagList, Map<String, String> bookmarkMap){
        this.text = docText;
        this.removeTagList = removeTagList;
        this.bookmarkMap = bookmarkMap;
    }

    @Override
    public String invoke() {
        String docText = text;
        Matcher rowMatcher = FwtElement.WORD_ROW_BOOKMARK_START_PATTERN.matcher(docText);
        while(rowMatcher.find()){
            String itemId = rowMatcher.group(1);
            String itemValue = rowMatcher.group(2).substring(FwtElement.CR_C_ROW_.length());
            String removeTag1 = String.format("<aml:annotation\\s*?aml:id=\"%s\"\\s*?w:type=\"Word.Bookmark.Start\"\\s*?w:name=\"%s%s\"\\s*?/>", itemId, FwtElement.CR_C_ROW_, itemValue);
            String removeTag2 = String.format("<aml:annotation\\s*?aml:id=\"%s\"\\s+?w:type=\"Word.Bookmark.End\"\\s*/>", itemId);
            removeTagList.add(removeTag1);
            removeTagList.add(removeTag2);
            // 擦除自定义语法块书签，并解析为目标dom
            try {
                docText = generateCustomDomTable(docText, itemValue);
            } catch (Exception e) {
                e.printStackTrace();
                log.error("发生错误，解析失败");
            }
        }
        Assert.notNull(docText, "成文模板解析异常");
        // 删除已解析标签，避免重复解析
        docText = TemplateParser.removeParsedTag(docText, removeTagList);
        notNull(docText, "成文模板解析异常");
        return docText;
    }

    private String generateCustomDomTable(String docText, String itemValue) throws Exception{
        String ftlTagStr = bookmarkMap.get(itemValue);
        FwtElement.FtlTagWrapper ftlTagWrapper = JSON.parseObject(ftlTagStr, FwtElement.FtlTagWrapper.class);
        Document doc = DocumentHelper.parseText(docText);
        String tableBookmarkStartXPath = String.format("//aml:annotation[@w:type='Word.Bookmark.Start'][@w:name='%s%s']", FwtElement.CR_C_ROW_, itemValue);
        Node tableBookmarkStartNode = doc.selectSingleNode(tableBookmarkStartXPath);
        Element target = TemplateParser.getTargetNode(tableBookmarkStartNode, FwtElement.WORD_TR_TAG);
        Assert.notNull(target, "Ftl模板解析异常");
        Element wrapperNode = target.getParent();
        @SuppressWarnings("unchecked")
        List<Element> elementList = wrapperNode.elements();
        // 获取目标节点的索引  YJ_20200817
        int index = 0;
        for(int i=0; i<elementList.size(); i++){
            if(elementList.get(i) == target){
                index = i;
                break;
            }
        }
        Element ftlElement1 = DocumentHelper.createElement("ftl");
        Element ftlElement2 = DocumentHelper.createElement("ftl");
        ftlElement1.addAttribute("content", ftlTagWrapper.startTag);
        elementList.add(index, ftlElement1);
        ftlElement2.addAttribute("content", ftlTagWrapper.endTag);
        elementList.add(index+2, ftlElement2);
        return doc.asXML();
    }
}
