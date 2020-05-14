package com.cr.visualtl.service.impl;

import com.cr.visualtl.config.Constant;
import com.cr.visualtl.model.ResponseTarget;
import com.cr.visualtl.service.ExportWord;
import com.cr.visualtl.util.fwt.FreemarkerWordTemplateUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import javax.servlet.http.HttpServletRequest;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;

@Service
public class ExportWordImpl implements ExportWord {
    @Autowired
    private HttpServletRequest httpServletRequest;

    @Override
    public ResponseTarget exportWord() {
        Map<String, Object> dataMap = new HashMap<>();
        dataMap.put("contractNo", "Z2017-031");
        dataMap.put("meetingHost", "tom");
        dataMap.put("scoreList_item1", "1");
        dataMap.put("commentList_item1", "备注1");
        dataMap.put("scoreList_item2", "2");
        dataMap.put("commentList_item2", "备注2");
        dataMap.put("scoreList_item3", "3");
        dataMap.put("commentList_item3", "备注3");
        dataMap.put("scoreList_item4", "4");
        dataMap.put("commentList_item4", "备注4");
        dataMap.put("scoreList_item5", "5");
        dataMap.put("commentList_item5", "备注5");
        dataMap.put("scoreList_item6", "6");
        dataMap.put("commentList_item6", "备注6");
        dataMap.put("scoreList_item7", "7");
        dataMap.put("commentList_item7", "备注7");
        dataMap.put("scoreList_item8", "8");
        dataMap.put("commentList_item8", "备注8");
        dataMap.put("scoreList_item9", "9");
        dataMap.put("commentList_item9", "备注9");
        dataMap.put("scoreList_item10", "10");
        dataMap.put("commentList_item10", "备注10");
        String templateDir = Constant.TEMPLATEFOLDER;
        String templateFileName = (String)httpServletRequest.getSession().getAttribute("ftlTemplateFileName");
        String outPath = Paths.get(Constant.OUTFOLDER, (String)httpServletRequest.getSession().getAttribute("templateFileName")).toString();
        boolean exported = FreemarkerWordTemplateUtils.exportSimpleWord(dataMap, templateDir, templateFileName, outPath);
        return new ResponseTarget(exported, exported? "导出成功": "导出失败");
    }
}
