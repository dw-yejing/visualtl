package com.cr.visualtl.util;

import java.util.UUID;

public class StringUtils {

    public static String get32UUID(){
        return UUID.randomUUID().toString().replace("-", "");
    }
}
