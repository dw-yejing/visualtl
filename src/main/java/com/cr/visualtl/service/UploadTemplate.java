package com.cr.visualtl.service;

import com.cr.visualtl.model.ResponseTarget;
import org.springframework.web.multipart.MultipartFile;

public interface UploadTemplate {
    ResponseTarget upload(MultipartFile templateFile);
}
