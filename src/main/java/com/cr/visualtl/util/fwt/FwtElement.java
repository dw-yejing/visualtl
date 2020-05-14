package com.cr.visualtl.util.fwt;

import java.util.regex.Pattern;

/**
 * 基于openxml的用户模板的相关常量  YJ_20200513
 */
public class FwtElement {
    // 书签
    public static final String WORD_BOOKMARK_START = "Word.Bookmark.Start";
    public static final String WORD_BOOKMARK_END = "Word.Bookmark.End";

    // 系统设定的字段值（Freemarker的代替字段）
    public static final String CUSTOM_FIELD_ = "CUSTOM_FIELD_";
    public static final String SYSTEM_FIELD_ = "SYSTEM_FIELD_";
    public static final String SYSTEM_BLOCK_ = "SYSTEM_BLOCK_";

    // openxml匹配正则
    // 开始书签
    public static final Pattern WORD_BOOKMARK_START_PATTERN = Pattern.compile("<aml:annotation [.]*? w:type=\"Word.Bookmark.Start\" w:name=\"([.]+?)\" />");
    // 结束书签
    public static final Pattern WORD_BOOKMARK_END_PATTERN = Pattern.compile("<aml:annotation [.]*? w:type=\"Word.Bookmark.End\" />");
    // 书签
    public static final Pattern WORD_BOOKMARK_PATTERN = Pattern.compile("<aml:annotation [.]*? w:type=\"Word.Bookmark.Start\" w:name=\"([.]+?)\" />[\r\n\\s]*?<aml:annotation [.]*? w:type=\"Word.Bookmark.End\" />");
    // 自定义字段
    public static final Pattern WORD_BOOKMARK_CUSTOM_FIELD_PATTERN = Pattern.compile("<aml:annotation[^>]*?w:type=\"Word.Bookmark.Start\"[^>]*?w:name=\"((?i)CUSTOM_FIELD_)(.+?)\"\\s*/>[^>]*?<aml:annotation[^>]*?w:type=\"Word.Bookmark.End\"\\s*/>");
    // 系统字段
    public static final Pattern WORD_BOOKMARK_SYSTEM_FIELD_PATTERN = Pattern.compile("<aml:annotation[^>]*?w:type=\"Word.Bookmark.Start\"[^>]*?w:name=\"((?i)SYSTEM_FIELD_)(.+?)\"\\s*/>[^>]*?<aml:annotation[^>]*?w:type=\"Word.Bookmark.End\"\\s*/>");
    // 系统固化块
    public static final Pattern WORD_BOOKMARK_SYSTEM_BLOCK_PATTERN = Pattern.compile("<aml:annotation[^>]+?w:type=\"Word.Bookmark.Start\"[^>]*?w:name=\"((?i)SYSTEM_BLOCK_)(.+?)\"\\s*/>[^>]*?<aml:annotation[^>]*?w:type=\"Word.Bookmark.End\"\\s*/>");
}
