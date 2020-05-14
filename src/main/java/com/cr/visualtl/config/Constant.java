package com.cr.visualtl.config;

import java.nio.file.Paths;

public class Constant {
     public static final String TEMPLATEFOLDER = Paths.get(System.getProperty("user.dir"),"template").toString();

     public static final String OUTFOLDER = Paths.get(System.getProperty("user.dir"), "word").toString();
}
