package com.hashedin.mailServer.controller;

import com.hashedin.mailServer.service.EmailService;
import com.hashedin.mailServer.utils.EmailDetails;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;
import java.util.Map;

@RestController
public class EmailController {

    @Autowired private EmailService emailService;

    @PostMapping(value="/sendMail", consumes = { "application/json" })
    public String sendMail(@RequestBody EmailDetails details)
    {
        System.out.println("Code---->"+details);
        String status = emailService.sendSimpleMail(details);
        return status;
    }

    @PostMapping(value="/sendMailWithAttachment",consumes = { "application/json" })
    public String sendMailWithAttachment(@RequestBody EmailDetails details)
    {
        String status= emailService.sendMailWithAttachment(details);
        return status;
    }

    @PostMapping(value="/generateExcel",consumes = { "application/json" })
    public String generateExcel(@RequestBody Object details) throws IOException {
        System.out.println("Controller details---->"+details);
        String status = emailService.generateExcel(details);
        return status;
    }
}

