package com.cr.visualtl.service.impl;

import com.cr.visualtl.config.Constant;
import com.cr.visualtl.model.ResponseTarget;
import com.cr.visualtl.service.ExportWord;
import com.cr.visualtl.util.fwt.FreemarkerWordTemplateUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import javax.servlet.http.HttpServletRequest;
import java.io.FileInputStream;
import java.io.InputStream;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;

@Service
public class ExportWordImpl implements ExportWord {
    @Autowired
    private HttpServletRequest httpServletRequest;

    @Override
    public ResponseTarget exportWord() throws Exception{
        Map<String, Object> dataMap = new HashMap<>();
        dataMap.put("contractNo", "Z2017-031");
        dataMap.put("meetingHost", "tom");
        dataMap.put("scoreListitem1", "1");
        dataMap.put("commentListitem1", "备注1");
        dataMap.put("scoreListitem2", "2");
        dataMap.put("commentListitem2", "备注2");
        dataMap.put("scoreListitem3", "3");
        dataMap.put("commentListitem3", "备注3");
        dataMap.put("scoreListitem4", "4");
        dataMap.put("commentListitem4", "备注4");
        dataMap.put("scoreListitem5", "5");
        dataMap.put("commentListitem5", "备注5");
        dataMap.put("scoreListitem6", "6");
        dataMap.put("commentListitem6", "备注6");
        dataMap.put("scoreListitem7", "7");
        dataMap.put("commentListitem7", "备注7");
        dataMap.put("scoreListitem8", "8");
        dataMap.put("commentListitem8", "备注8");
        dataMap.put("scoreListitem9", "9");
        dataMap.put("commentListitem9", "备注9");
        dataMap.put("scoreListitem10", "10");
        dataMap.put("commentListitem10", "备注10");
        String templateDir = Constant.TEMPLATEFOLDER;
        String templateFileName = (String)httpServletRequest.getSession().getAttribute("ftlTemplateFileName");
        InputStream ftlInputStream = new FileInputStream(Paths.get(templateDir, templateFileName).toFile());
        String outPath = Paths.get(Constant.OUTFOLDER, (String)httpServletRequest.getSession().getAttribute("templateFileName")).toString();
        FreemarkerWordTemplateUtils.exportSimpleWord(dataMap, templateDir, ftlInputStream, outPath);
        return new ResponseTarget(true, "导出成功");
    }
}
