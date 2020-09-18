package com.cr.visualtl.util.fwt;

import java.util.regex.Pattern;

/**
 * 基于openxml的用户模板的相关常量  YJ_20200513
 */
public class FwtElement {
    // tag prefix
    // region
    // 系统设定的字段值（Freemarker的代替字段）
    static final String CR_BLOCK_ = "crblock_";

    // 时间字段
    static final String CR_DATE_ = "crdate_";
    // 默认时间格式
    static final String DEFAULT_DATE_FORMAT = "d";
    static final String DEFAULT_DATE_PATTERN = "yyyy-MM-dd";

    // 系统自定义标签1  Celll Level
    static final String CR_C_CELL_ = "cr_c_tc_";

    // 系统自定义标签2 Row Level
    static final String CR_C_ROW_ = "cr_c_tr_";

    // 系统自定义标签3  Table level
    static final String CR_C_TABLE_ = "cr_c_tbl_";

    // 系统自定义标签4 Run Level
    static final String CR_C_RUN_ = "cr_c_run_";

    // 系统自定义标签5 Paragraph Level
    static final String CR_C_PARA_ = "cr_c_p_";

    // Word标签1 <w:tbl>
    static final String WORD_TBL_TAG = "w:tbl";

    // Word标签2 <w:tr>
    static final String WORD_TR_TAG = "w:tr";

    // Word标签3 <w:tc>
    static final String WORD_TC_TAG = "w:tc";

    // Word标签4 <w:p>
    static final String WORD_P_TAG = "w:p";

    // Word标签5 <w:r>
    static final String WORD_R_TAG = "w:r";
    // endregion

    // bookmark pattern
    // region
    // openxml匹配正则
    // 开始书签
    private static final String wordBookmarkStart = "<aml:annotation[\\s]*?aml:id=\"([\\d]+?)\"[\\s]*?w:type=\"Word.Bookmark.Start\"[^>]*?w:name=\"%s\"\\s*/>[^>]*?";

    // 普通字段
    private static final String generalBookmarkName = String.format("((?i)((?!%s)(.+?)))", CR_BLOCK_);
    private static String generalField = String.format(wordBookmarkStart, generalBookmarkName);
    static final Pattern WORD_BOOKMARK_GENERAL_FIELD_PATTERN = Pattern.compile(generalField);

    // 系统固化块（if语法）
    private static String blockBookmarkName =  String.format("((?i)((?=%s)(.+?)))", CR_BLOCK_);
    private static String blockField = String.format(wordBookmarkStart, blockBookmarkName);
    static final Pattern WORD_BOOKMARK_SYSTEM_BLOCK_PATTERN = Pattern.compile(blockField);

    /**
     * 格式化时间
     */
    private static String dateBookmarkName = String.format("((?i)((?=%s)([\\da-zA-Z_]+)))", CR_DATE_);
    private static String dateField = String.format(wordBookmarkStart, dateBookmarkName);
    static final Pattern WORD_DATE_PATTERN = Pattern.compile(dateField);

    // Cell Level
    private static String cellBookmarkName = String.format("((?i)((?=%s)(.+?)))", CR_C_CELL_);
    private static String cellBookmarkStart = String.format(wordBookmarkStart, cellBookmarkName);
    static final Pattern WORD_CELL_BOOKMARK_START_PATTERN = Pattern.compile(cellBookmarkStart);

    // Row Level
    private static String rowBookmarkName = String.format("((?i)((?=%s)(.+?)))", CR_C_ROW_);
    private static String rowBookmarkStart = String.format(wordBookmarkStart, rowBookmarkName);
    static final Pattern WORD_ROW_BOOKMARK_START_PATTERN = Pattern.compile(rowBookmarkStart);

    // Table Level
    private static String  tableBookmarkName = String.format("((?i)((?=%s)(.+?)))", CR_C_TABLE_);
    private static String tableBookmarkStart = String.format(wordBookmarkStart, tableBookmarkName);
    static final Pattern WORD_TABLE_BOOKMARK_START_PATTERN = Pattern.compile(tableBookmarkStart);

    // Run Level
    private static String runBookmarkName = String.format("((?i)((?=%s)(.+?)))", CR_C_RUN_);
    private static String runBookmarkStart = String.format(wordBookmarkStart, runBookmarkName);
    static final Pattern WORD_RUN_BOOKMARK_START_PATTERN = Pattern.compile(runBookmarkStart);

    //Para(Paragraph) Level
    private static String paraBookmarkName = String.format("((?i)((?=%s)(.+?)))", CR_C_PARA_);
    private static String paraBookmarkStart = String.format(wordBookmarkStart, paraBookmarkName);
    static final Pattern WORD_PARA_BOOKMARK_START_PATTERN = Pattern.compile(paraBookmarkStart);
    // endregion

    // other pattern
    // region
    // ftl临时匹配模式  YJ_20200817
    static Pattern TEMPORARY_XML_PATTERN = Pattern.compile("<ftl\\s+?content=\"([\\s\\S]+?)\"\\s*?/>");

    // object separator  YJ_20200805
    private static String objectSeparator = "(?<=\\$\\{)[^}]+?(?=})";
    static final Pattern OBJECT_SEPARATOR_PATTERN = Pattern.compile(objectSeparator);
    //endregion

    // tag wrapper
    // region
    public static class FtlTagWrapper{
        public String startTag;

        public String endTag;
    }


    public static FtlTagWrapper initFtlWrapper(){
        return new FtlTagWrapper();
    }
    // endregion
}
