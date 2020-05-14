package com.cr.visualtl.util;

import java.io.*;

public class CrFileUtils {
    private static int DEFAULT_BUFFER_SIZE = 2048;

    public static String convertFile2text(File file) {
        StringBuilder sb = new StringBuilder();
        try(InputStreamReader inputStreamReader = new InputStreamReader(new FileInputStream(file))) {
            char[] buffer = new char[DEFAULT_BUFFER_SIZE];
            int len;
            while ((len = inputStreamReader.read(buffer)) != -1) {
                sb.append(buffer, 0, len);
            }
        }catch (Exception e){
            e.printStackTrace();
            return null;
        }
        return sb.toString();
    }

    public static boolean convertText2file(String text, String filepath){
        try(ByteArrayInputStream bais = new ByteArrayInputStream(text.getBytes());
            FileOutputStream fos = new FileOutputStream(filepath)) {
            int len;
            byte[] buffer = new byte[DEFAULT_BUFFER_SIZE];
            while ((len = bais.read(buffer)) != -1) {
                fos.write(buffer, 0, len);
            }
            fos.flush();
        }catch (Exception e){
            e.printStackTrace();
            return false;
        }
        return true;
    }
}
