package com.hashedin.mailServer.service;

import com.hashedin.mailServer.utils.EmailDetails;

import java.io.FileNotFoundException;
import java.io.IOException;

public interface EmailService {
    String sendSimpleMail(EmailDetails details);

    String sendMailWithAttachment(EmailDetails details);

    String generateExcel(Object details) throws FileNotFoundException, IOException;
}
