package com.cr.visualtl.util.fwt;

import cn.hutool.core.io.IoUtil;
import com.cr.visualtl.util.StringUtils;
import freemarker.template.Configuration;
import freemarker.template.Template;
import lombok.extern.slf4j.Slf4j;
import org.springframework.web.util.HtmlUtils;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;

// Freemarker成文工具类  YJ_20200513
@Slf4j
public class FreemarkerWordTemplateUtils {

    public static String doc2ftl(Map<String, String> map, InputStream inputStream){
        // 生成ftl模板
        String docText = IoUtil.read(inputStream, StandardCharsets.UTF_8);
        // 解析自定义语法块  YJ_20200814
        StringBuffer docBuffer = new StringBuffer();
        int last = -1;
        // 解析自定义标签
        // region
        // 存储待删除的cr_c_*书签
        List<String> removeTagList = new ArrayList<>();
        // Table Level
        TemplateParser templateParser = new TblTemplateParser(docText, removeTagList, map);
        docText = templateParser.invoke();
        // Row Level
        templateParser = new TrTemplateParser(docText, removeTagList, map);
        docText = templateParser.invoke();
        // Cell Level
        templateParser = new TcTemplateParser(docText, removeTagList, map);
        docText = templateParser.invoke();
        // Para Level
        templateParser = new ParaTemplateParser(docText, removeTagList, map);
        docText = templateParser.invoke();
        // Run level
        templateParser = new RunTemplateParser(docText, removeTagList, map);
        docText = templateParser.invoke();
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
        // endregion

        // 解析系统固化的Freemarker语法块(if语法)
        templateParser = new BlockTemplateParser(docText, removeTagList, map);
        docText = templateParser.invoke();

        // 格式化时间
        templateParser = new DateTemplateParser(docText, removeTagList, map);
        docText = templateParser.invoke();

        // 解析普通字段
        templateParser = new GeneralFieldTemplateParser(docText, removeTagList);
        docText = templateParser.invoke();

        // 修正对象分割符（将"_"变为"."）
        templateParser = new SeparatorTemplateParser(docText);
        docText = templateParser.invoke();

        return docText;
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
     * @Param outputSteam 非Servlet响应输出流
     */
    public static class ExportWordWrapper{
        public String outFileName;
        public String ftlTemplateLocation;
        public InputStream ftlInputStream;
        public Map<String, Object> dataMap;
        public  HttpServletResponse httpServletResponse;
        public HttpServletRequest httpServletRequest;
        public OutputStream outputStream;
    }
}
