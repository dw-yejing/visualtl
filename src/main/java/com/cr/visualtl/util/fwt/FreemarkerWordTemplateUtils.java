package com.cr.visualtl.util.fwt;

import cn.hutool.core.io.IoUtil;
import com.alibaba.fastjson.JSON;
import com.cr.visualtl.util.StringUtils;
import freemarker.template.Configuration;
import freemarker.template.Template;
import lombok.extern.slf4j.Slf4j;
import org.dom4j.*;
import org.springframework.util.Assert;
import org.springframework.web.util.HtmlUtils;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.regex.Matcher;

// Freemarker成文工具类  YJ_20200513
@Slf4j
public class FreemarkerWordTemplateUtils {

    public static String doc2ftl(Map<String, String> map, InputStream inputStream) throws Exception{
        // 生成ftl模板
        String docText = IoUtil.read(inputStream, StandardCharsets.UTF_8);
        // 解析自定义语法块  YJ_20200814
        StringBuffer docBuffer = new StringBuffer();
        int last = -1;
        // region
        // 存储待删除的cr_c_*书签
        List<String> removeTagList = new ArrayList<>();
        // Table Level
        Matcher tableMatcher = FwtElement.WORD_TABLE_BOOKMARK_START_PATTERN.matcher(docText);
        while(tableMatcher.find()){
            String itemId = tableMatcher.group(1);
            String itemValue = tableMatcher.group(2).substring(FwtElement.CR_C_TABLE_.length());
            String removeTag1 = String.format("<aml:annotation\\s*?aml:id=\"%s\"\\s*?w:type=\"Word.Bookmark.Start\"\\s*?w:name=\"%s%s\"\\s*?/>", itemId, FwtElement.CR_C_TABLE_, itemValue);
            String removeTag2 = String.format("<aml:annotation\\s*?aml:id=\"%s\"\\s*?w:type=\"Word.Bookmark.End\"\\s*/>", itemId);
            removeTagList.add(removeTag1);
            removeTagList.add(removeTag2);
            // 擦除自定义语法块书签，并解析为目标dom
            docText = generateCustomDom(docText, itemValue, map, FwtElement.CR_C_TABLE_);
        }
        Assert.notNull(docText, "成文模板解析异常");
        // Row Level
        Matcher rowMatcher = FwtElement.WORD_ROW_BOOKMARK_START_PATTERN.matcher(docText);
        while(rowMatcher.find()){
            String itemId = rowMatcher.group(1);
            String itemValue = rowMatcher.group(2).substring(FwtElement.CR_C_ROW_.length());
            String removeTag1 = String.format("<aml:annotation\\s*?aml:id=\"%s\"\\s*?w:type=\"Word.Bookmark.Start\"\\s*?w:name=\"%s%s\"\\s*?/>", itemId, FwtElement.CR_C_ROW_, itemValue);
            String removeTag2 = String.format("<aml:annotation\\s*?aml:id=\"%s\"\\s+?w:type=\"Word.Bookmark.End\"\\s*/>", itemId);
            removeTagList.add(removeTag1);
            removeTagList.add(removeTag2);
            // 擦除自定义语法块书签，并解析为目标dom
            docText = generateCustomDom(docText, itemValue, map, FwtElement.CR_C_ROW_);
        }
        Assert.notNull(docText, "成文模板解析异常");
        // Cell Level
        Matcher cellMatcher = FwtElement.WORD_CELL_BOOKMARK_START_PATTERN.matcher(docText);
        while(cellMatcher.find()){
            String itemId = cellMatcher.group(1);
            String itemValue = cellMatcher.group(2).substring(FwtElement.CR_C_CELL_.length());
            String removeTag1 = String.format("<aml:annotation\\s*?aml:id=\"%s\"\\s*?w:type=\"Word.Bookmark.Start\"\\s*?w:name=\"%s%s\"\\s*?/>", itemId, FwtElement.CR_C_CELL_, itemValue);
            String removeTag2 = String.format("<aml:annotation\\s*?aml:id=\"%s\"\\s+?w:type=\"Word.Bookmark.End\"\\s*/>", itemId);
            removeTagList.add(removeTag1);
            removeTagList.add(removeTag2);
            // 擦除自定义语法块书签，并解析为目标dom
            docText = generateCustomDom(docText, itemValue, map, FwtElement.CR_C_CELL_);
        }
        Assert.notNull(docText, "成文模板解析异常");
        // Para Level
        Matcher paraMatcher = FwtElement.WORD_PARA_BOOKMARK_START_PATTERN.matcher(docText);
        while(paraMatcher.find()){
            String itemId = paraMatcher.group(1);
            String itemValue = paraMatcher.group(2).substring(FwtElement.CR_C_PARA_.length());
            String removeTag1 = String.format("<aml:annotation\\s*?aml:id=\"%s\"\\s*?w:type=\"Word.Bookmark.Start\"\\s*?w:name=\"%s%s\"\\s*?/>", itemId, FwtElement.CR_C_PARA_, itemValue);
            String removeTag2 = String.format("<aml:annotation\\s*?aml:id=\"%s\"\\s+?w:type=\"Word.Bookmark.End\"\\s*/>", itemId);
            removeTagList.add(removeTag1);
            removeTagList.add(removeTag2);
            // 擦除自定义语法块书签，并解析为目标dom
            docText = generateCustomDom(docText, itemValue, map, FwtElement.CR_C_PARA_);
        }
        Assert.notNull(docText, "成文模板解析异常");
        // Run level
        Matcher runMatcher = FwtElement.WORD_RUN_BOOKMARK_START_PATTERN.matcher(docText);
        while(runMatcher.find()){
            String itemId = runMatcher.group(1);
            String itemValue = runMatcher.group(2).substring(FwtElement.CR_C_RUN_.length());
            String removeTag1 = String.format("<aml:annotation\\s*?aml:id=\"%s\"\\s*?w:type=\"Word.Bookmark.Start\"\\s*?w:name=\"%s%s\"\\s*?/>", itemId, FwtElement.CR_C_RUN_, itemValue);
            String removeTag2 = String.format("<aml:annotation\\s*?aml:id=\"%s\"\\s+?w:type=\"Word.Bookmark.End\"\\s*/>", itemId);
            removeTagList.add(removeTag1);
            removeTagList.add(removeTag2);
            // 擦除自定义语法块书签，并解析为目标dom
            docText = generateCustomDom(docText, itemValue, map, FwtElement.CR_C_RUN_);
        }
        Assert.notNull(docText, "成文模板解析异常");

        // 将xml解析为ftl语法文件
        Matcher xmlTemporaryMatcher = FwtElement.TEMPORARY_XML_PATTERN.matcher(docText);
        while(xmlTemporaryMatcher.find()){
            String content = xmlTemporaryMatcher.group(1);
            //反转义Html实体字符
            content = HtmlUtils.htmlUnescape(content);
            xmlTemporaryMatcher.appendReplacement(docBuffer, content);
            last = xmlTemporaryMatcher.end();
        }
        // 将替换后的部分补全
        if(docBuffer.length()>0 && last>0){
            docText = docBuffer.toString() + docText.substring(last);
        }
        // 删除cr_c_*书签
        for(String item : removeTagList){
            docText = docText.replaceAll(item, "");
        }
        // endregion

        // 解析普通字段
        docText = Objects.requireNonNull(docText).replaceAll(FwtElement.WORD_BOOKMARK_GENERAL_FIELD_PATTERN.pattern(), "<w:r><w:t>\\$\\{($2)!\\}</w:t></w:r>");
        // 解析系统固化的Freemarker语法块(if语法)
        docBuffer = new StringBuffer();
        Matcher blockMatcher = FwtElement.WORD_BOOKMARK_SYSTEM_BLOCK_PATTERN.matcher(Objects.requireNonNull(docText));
        last = -1;
        while(blockMatcher.find()){
            String groupValue = blockMatcher.group(2);
            groupValue = groupValue.substring(FwtElement.CR_BLOCK_.length());
            String blockValue = map.get(groupValue);
            blockMatcher.appendReplacement(docBuffer, blockValue);
            last = blockMatcher.end();
        }
        // 将替换后的部分补全
        if(docBuffer.length()>0 && last>0){
            docText = docBuffer.toString() + docText.substring(last);
        }
        // 修正对象分割符（将"_"变为"."）
        Matcher separatorMatcher = FwtElement.OBJECT_SEPARATOR_PATTERN.matcher(docText);
        last = -1;
        docBuffer = new StringBuffer();
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
        return docText;
    }

    /**
     *  擦除自定义语法块书签，并解析为目标dom  YJ_20200814
     */
    private static String generateCustomDom(String docText, String itemValue, Map<String,String> map, String type) throws Exception{
        if(Objects.equals(type, FwtElement.CR_C_TABLE_)){
            return generateCustomDomTable(docText, itemValue, map);
        }
        if(Objects.equals(type, FwtElement.CR_C_ROW_)){
            return generateCustomDomRow(docText, itemValue, map);
        }
        if(Objects.equals(type, FwtElement.CR_C_CELL_)){
            return generateCustomDomCell(docText, itemValue, map);
        }
        if(Objects.equals(type, FwtElement.CR_C_PARA_)){
            return generateCustomDomPara(docText, itemValue, map);
        }
        if(Objects.equals(type, FwtElement.CR_C_RUN_)){
            return generateCustomDomRun(docText, itemValue, map);
        }
        return null;
    }

    private static String generateCustomDomTable(String docText, String itemValue, Map<String,String> map) throws Exception{
        String ftlTagStr = map.get(itemValue);
        FwtElement.FtlTagWrapper ftlTagWrapper = JSON.parseObject(ftlTagStr, FwtElement.FtlTagWrapper.class);
        Document doc = DocumentHelper.parseText(docText);
        String tableBookmarkStartXPath = String.format("//aml:annotation[@w:type='Word.Bookmark.Start'][@w:name='%s%s']", FwtElement.CR_C_TABLE_, itemValue);
        Node tableBookmarkStartNode = doc.selectSingleNode(tableBookmarkStartXPath);
        Element target = getTargetNode(tableBookmarkStartNode, FwtElement.WORD_TBL_TAG);
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

    private static String generateCustomDomRow(String docText, String itemValue, Map<String,String> map) throws Exception{
        String ftlTagStr = map.get(itemValue);
        FwtElement.FtlTagWrapper ftlTagWrapper = JSON.parseObject(ftlTagStr, FwtElement.FtlTagWrapper.class);
        Document doc = DocumentHelper.parseText(docText);
        String tableBookmarkStartXPath = String.format("//aml:annotation[@w:type='Word.Bookmark.Start'][@w:name='%s%s']", FwtElement.CR_C_ROW_, itemValue);
        Node tableBookmarkStartNode = doc.selectSingleNode(tableBookmarkStartXPath);
        Element target = getTargetNode(tableBookmarkStartNode, FwtElement.WORD_TR_TAG);
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

    private static String generateCustomDomCell(String docText, String itemValue, Map<String,String> map) throws Exception{
        String ftlTagStr = map.get(itemValue);
        FwtElement.FtlTagWrapper ftlTagWrapper = JSON.parseObject(ftlTagStr, FwtElement.FtlTagWrapper.class);
        Document doc = DocumentHelper.parseText(docText);
        String tableBookmarkStartXPath = String.format("//aml:annotation[@w:type='Word.Bookmark.Start'][@w:name='%s%s']", FwtElement.CR_C_CELL_, itemValue);
        Node tableBookmarkStartNode = doc.selectSingleNode(tableBookmarkStartXPath);
        Element target = getTargetNode(tableBookmarkStartNode, FwtElement.WORD_TC_TAG);
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

    private static String generateCustomDomPara(String docText, String itemValue, Map<String,String> map) throws Exception{
        String ftlTagStr = map.get(itemValue);
        FwtElement.FtlTagWrapper ftlTagWrapper = JSON.parseObject(ftlTagStr, FwtElement.FtlTagWrapper.class);
        Document doc = DocumentHelper.parseText(docText);
        String tableBookmarkStartXPath = String.format("//aml:annotation[@w:type='Word.Bookmark.Start'][@w:name='%s%s']", FwtElement.CR_C_PARA_, itemValue);
        Node tableBookmarkStartNode = doc.selectSingleNode(tableBookmarkStartXPath);
        Element target = getTargetNode(tableBookmarkStartNode, FwtElement.WORD_P_TAG);
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

    private static String generateCustomDomRun(String docText, String itemValue, Map<String,String> map) throws Exception{
        String ftlTagStr = map.get(itemValue);
        FwtElement.FtlTagWrapper ftlTagWrapper = JSON.parseObject(ftlTagStr, FwtElement.FtlTagWrapper.class);
        Document doc = DocumentHelper.parseText(docText);
        String tableBookmarkStartXPath = String.format("//aml:annotation[@w:type='Word.Bookmark.Start'][@w:name='%s%s']", FwtElement.CR_C_RUN_, itemValue);
        Node tableBookmarkStartNode = doc.selectSingleNode(tableBookmarkStartXPath);
        Element target = getTargetNode(tableBookmarkStartNode, FwtElement.WORD_R_TAG);
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


    /**
     * 获取目标父节点  YJ_20200814
     * @param fromNode 原节点
     * @param targetNodeName 父节点名称
     */
    private static Element getTargetNode(Node fromNode, String targetNodeName){
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

    public static void exportSimpleWord2Stream(ExportWordWrapper exportWordWrapper){
        String fileName = exportWordWrapper.outFileName;
        String ftlTemplateLocation = exportWordWrapper.ftlTemplateLocation;
        InputStream ftlInputStream = exportWordWrapper.ftlInputStream;
        Map<?,?> dataMap = exportWordWrapper.dataMap;
        HttpServletResponse httpServletResponse = exportWordWrapper.httpServletResponse;
        HttpServletRequest httpServletRequest = exportWordWrapper.httpServletRequest;
        try{
            httpServletResponse.setCharacterEncoding("utf-8");
            if (httpServletRequest.getHeader("User-Agent").toLowerCase().indexOf("firefox") > 0) {
                fileName = new String(fileName.getBytes(StandardCharsets.UTF_8), "ISO8859-1"); // firefox浏览器
            } else if (httpServletRequest.getHeader("User-Agent").toUpperCase().indexOf("MSIE") > 0) {
                fileName = URLEncoder.encode(fileName, "UTF-8");// IE浏览器
            }else if (httpServletRequest.getHeader("User-Agent").toUpperCase().indexOf("CHROME") > 0) {
                fileName = new String(fileName.getBytes(StandardCharsets.UTF_8), "ISO8859-1");// 谷歌
            }
            httpServletResponse.addHeader("content-type", "application/octet-stream");
            httpServletResponse.setHeader("content-disposition", "attachment;filename=" + fileName);
        }catch (Exception e){
            e.printStackTrace();
        }
        String ftlFileName = StringUtils.get32UUID()+".ftl";
        String ftlFilePath = Paths.get(ftlTemplateLocation, ftlFileName).toString();
        File ftlFolder = new File(ftlTemplateLocation);
        if(!ftlFolder.exists()){
            boolean maked = ftlFolder.mkdirs();
            if(!maked){
                String msg = String.format("目录【%s】，新建失败", ftlTemplateLocation);
                log.error(msg);
            }
        }
        try(OutputStream outputStream = httpServletResponse.getOutputStream();
            OutputStream ftlOutputStream = new FileOutputStream(ftlFilePath)){
            // 下载ftl模板，并本地临时存储
            IoUtil.copy(ftlInputStream, ftlOutputStream);
            ftlOutputStream.flush();
            FreemarkerWordTemplateUtils.exportSimpleWord2Stream(dataMap, ftlTemplateLocation, fileName, outputStream);
            // 删除临时ftl模板
            boolean deleted = new File(ftlFilePath).delete();
            if(!deleted){
                String msg = String.format("文档【%s】,未删除成功", ftlFilePath);
                log.info(msg);
            }
        }catch (Exception e){
            e.printStackTrace();
        }
    }

    /**
     * 利用ftl模板，成文到输出流  YJ_20200803
     */
    public static void exportSimpleWord2Stream(Map<?,?> dataMap, String templateDir, String templateFileName, OutputStream outputStream){

        Configuration configuration = new Configuration(Configuration.DEFAULT_INCOMPATIBLE_IMPROVEMENTS);
        configuration.setDefaultEncoding("utf-8");
        try{
            configuration.setDirectoryForTemplateLoading(new File(templateDir));
            Template template = configuration.getTemplate(templateFileName);
            Writer outWriter = new OutputStreamWriter(outputStream);
            template.process(dataMap, outWriter);
            outWriter.flush();
            outWriter.close();
        }catch (Exception e){
            e.printStackTrace();
        }
    }

    @SuppressWarnings("unused")
    public static void exportSimpleWord(Map<?,?> dataMap, String ftlTemplateLocation,InputStream ftlInputStream, String outPath){
        String ftlFileName = StringUtils.get32UUID()+".ftl";
        String ftlFilePath = Paths.get(ftlTemplateLocation, ftlFileName).toString();
        File ftlFolder = new File(ftlTemplateLocation);
        if(!ftlFolder.exists()){
            boolean maked = ftlFolder.mkdirs();
            if(!maked){
                String msg = String.format("目录【%s】，新建失败", ftlTemplateLocation);
                log.error(msg);
            }
        }
        File outFile = new File(outPath);
        File outFileFolder = outFile.getParentFile();
        if(!outFileFolder.exists()){
            boolean maked = outFileFolder.mkdirs();
            if(!maked){
                String msg = String.format("目录【%s】，新建失败", ftlTemplateLocation);
                log.error(msg);
            }
        }
        try(OutputStream outputStream = new FileOutputStream(outPath);
            OutputStream ftlOutputStream = new FileOutputStream(ftlFilePath)){
            // 下载ftl模板，并本地临时存储
            IoUtil.copy(ftlInputStream, ftlOutputStream);
            ftlOutputStream.flush();
            FreemarkerWordTemplateUtils.exportSimpleWord(dataMap, ftlTemplateLocation, ftlFileName, outPath);
            // 删除临时ftl模板
            boolean deleted = new File(ftlFilePath).delete();
            if(!deleted){
                String msg = String.format("文档【%s】,未删除成功", ftlFilePath);
                log.info(msg);
            }
        }catch (Exception e){
            e.printStackTrace();
        }
    }

    /**
     * 利用ftl模板，成文到指定文件  YJ_20200803
     */
    private static void exportSimpleWord(Map<?,?> dataMap, String templateDir, String templateFileName, String outPath){
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
        }
    }

    /**
     * 成文参数封装类  YJ_20200818
     * @Param outFileName 成文的Word名称
     * @Param ftlTemplateLocation 下载ftl的暂存地址
     * @Param ftlInputStream 下载的ftl文件流
     * @Param dataMap 成文的数据字典
     * @Param httpServletResponse 响应信息
     * @Param httpServletRequest 请求信息
     */
    public static class ExportWordWrapper{
        public String outFileName;
        public String ftlTemplateLocation;
        public InputStream ftlInputStream;
        public Map<String, Object> dataMap;
        public  HttpServletResponse httpServletResponse;
        public HttpServletRequest httpServletRequest;
    }
}
