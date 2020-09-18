package com.cr.visualtl.util;

import java.util.UUID;

public class StringUtils {

    public static String get32UUID(){
        return UUID.randomUUID().toString().replace("-", "");
    }

    public static boolean isBlank(CharSequence str) {
        int length;
        if (str != null && (length = str.length()) != 0) {
            for(int i = 0; i < length; ++i) {
                if (str.charAt(i)!=' ') {
                    return false;
                }
            }

            return true;
        } else {
            return true;
        }
    }

    public static boolean isNotBlank(CharSequence str) {
        return !isBlank(str);
    }
}
