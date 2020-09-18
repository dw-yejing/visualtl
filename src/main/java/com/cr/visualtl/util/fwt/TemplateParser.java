package com.cr.visualtl.util.fwt;

import org.dom4j.Element;
import org.dom4j.Node;

import java.util.*;

public interface TemplateParser {

    String invoke();

    /**
     * 获取目标父节点  YJ_20200814
     * @param fromNode 原节点
     * @param targetNodeName 父节点名称
     */
    static Element getTargetNode(Node fromNode, String targetNodeName){
        Element parentNode = fromNode.getParent();
        if(parentNode == null){
            return null;
        }
        String nodeName = parentNode.getQualifiedName();
        if(Objects.equals(nodeName, targetNodeName)){
            return parentNode;
        }else {
            return getTargetNode(parentNode, targetNodeName);
        }
    }

    /**
     * 删除已解析标签  YJ_20200908
     */
    static String removeParsedTag(String docText, List<String> removeTagList){
        // 删除书签
        for(String item : removeTagList){
            docText = docText.replaceAll(item, "");
        }
        removeTagList.clear();
        return docText;
    }
}
