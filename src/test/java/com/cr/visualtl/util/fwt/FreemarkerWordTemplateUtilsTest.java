package com.cr.visualtl.util.fwt;

import org.junit.jupiter.api.Test;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

class FreemarkerWordTemplateUtilsTest {

    @Test
    void word2ftl() {
    }

    @Test
    void regex() {
        String a = "123456789";
        Pattern pattern = Pattern.compile("((?i))(\\d)");
        Matcher matcher = pattern.matcher(a);
        StringBuffer sb = new StringBuffer();
        while(matcher.find()){
            String b = matcher.group();
            matcher.appendReplacement(sb, Integer.valueOf(b)%2+"");
            System.out.println(b);
        }
        System.out.println(sb.toString());

    }
}