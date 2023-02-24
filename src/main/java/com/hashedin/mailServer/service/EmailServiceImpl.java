package com.hashedin.mailServer.service;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;
import javax.mail.MessagingException;
import javax.mail.internet.MimeMessage;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.hashedin.mailServer.utils.EmailDetails;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.FileSystemResource;
import org.springframework.mail.SimpleMailMessage;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.springframework.stereotype.Service;

@Service
public class EmailServiceImpl implements EmailService {

    @Autowired private JavaMailSender javaMailSender;

    @Value("${spring.mail.username}") private String sender;

    public String sendSimpleMail(EmailDetails details)
    {

        try {

            SimpleMailMessage mailMessage = new SimpleMailMessage();

            mailMessage.setFrom(sender);
            mailMessage.setTo(details.getRecipient());
            mailMessage.setText(details.getMsgBody());
            mailMessage.setSubject(details.getSubject());

            javaMailSender.send(mailMessage);
            return "Mail Sent Successfully...";
        }

        catch (Exception e) {
            return "Error while Sending Mail";
        }
    }

    public String sendMailWithAttachment(EmailDetails details)
    {
        MimeMessage mimeMessage = javaMailSender.createMimeMessage();
        MimeMessageHelper mimeMessageHelper;

        try {

            mimeMessageHelper = new MimeMessageHelper(mimeMessage, true);
            mimeMessageHelper.setFrom(sender);
            mimeMessageHelper.setTo(details.getRecipient());
            mimeMessageHelper.setText(details.getMsgBody());
            mimeMessageHelper.setSubject(details.getSubject());

            FileSystemResource file = new FileSystemResource(new File(details.getAttachment()));

            mimeMessageHelper.addAttachment(file.getFilename(), file);
            javaMailSender.send(mimeMessage);
            return "Mail sent Successfully";
        }

        catch (MessagingException e) {
            // Display message when exception occurred
            return "Error while sending mail!!!";
        }
    }

    public String generateExcel(Object details) throws FileNotFoundException, IOException {

        System.out.println("Service impl details---->"+details);
        ObjectMapper om = new ObjectMapper();
        try {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("Details");

        ArrayList al1 = new ArrayList();
        al1 = (ArrayList) details;
        System.out.println("al1---->"+al1);

        int colNum = 0;
        Row row = sheet.createRow(0);
        int rowNum = 1;
        if (rowNum==1){
            Cell apiNameCell1 = row.createCell(colNum++);
            Cell methodCell1 = row.createCell(colNum++);
            Cell baseUrlCell1 = row.createCell(colNum++);
            Cell pathCell1 = row.createCell(colNum++);
            Cell payloadJsonCell1 = row.createCell(colNum++);
            Cell pathParamCell1 = row.createCell(colNum++);
            Cell requestParamCell1 = row.createCell(colNum++);
            Cell responseTimeCell1 = row.createCell(colNum++);
            Cell expectedStatusCell1 = row.createCell(colNum++);
            Cell responseStatusCell1 = row.createCell(colNum++);
            Cell passedOrFailedCell1 = row.createCell(colNum++);
            Cell responsePayloadCell1 = row.createCell(colNum++);

            apiNameCell1.setCellValue("API");
            methodCell1.setCellValue("Method");
            baseUrlCell1.setCellValue("Base URL");
            pathCell1.setCellValue("Path");
            payloadJsonCell1.setCellValue("Payload");
            pathParamCell1.setCellValue("Path Param");
            requestParamCell1.setCellValue("Request Param");
            responseTimeCell1.setCellValue("Response Time");
            expectedStatusCell1.setCellValue("Expected Status");
            responseStatusCell1.setCellValue("Response Status");
            passedOrFailedCell1.setCellValue("Passed Or Fail");
            responsePayloadCell1.setCellValue("Response Payload");
        }

        int i = 0;
        colNum = 0;
        rowNum = 1;
        JsonNode rowNode;
        while(i < al1.size()) {
            String jsonStr = om.writeValueAsString(al1.get(i));
            rowNode=om.readTree(jsonStr);
            Row bodyRow = sheet.createRow(rowNum++);
            Cell apiNameCell = bodyRow.createCell(colNum++);
            Cell methodCell = bodyRow.createCell(colNum++);
            Cell baseUrlCell = bodyRow.createCell(colNum++);
            Cell pathCell = bodyRow.createCell(colNum++);
            Cell payloadJsonCell = bodyRow.createCell(colNum++);
            Cell pathParamCell = bodyRow.createCell(colNum++);
            Cell requestParamCell = bodyRow.createCell(colNum++);
            Cell responseTimeCell = bodyRow.createCell(colNum++);
            Cell expectedStatusCell = bodyRow.createCell(colNum++);
            Cell responseStatusCell = bodyRow.createCell(colNum++);
            Cell passedOrFailedCell = bodyRow.createCell(colNum++);
            Cell responsePayloadCell = bodyRow.createCell(colNum++);

            apiNameCell.setCellValue(rowNode.get("apiName").asText());
            methodCell.setCellValue(rowNode.get("method").asText());
            baseUrlCell.setCellValue(rowNode.get("baseUrl").asText());
            pathCell.setCellValue(rowNode.get("path").asText());
            payloadJsonCell.setCellValue(rowNode.get("payloadJson").asText());
            pathParamCell.setCellValue(rowNode.get("pathParam").asText());
            requestParamCell.setCellValue(rowNode.get("requestParam").asText());
            responseTimeCell.setCellValue(rowNode.get("responseTime").asText());
            expectedStatusCell.setCellValue(rowNode.get("expectedStatus").asText());
            responseStatusCell.setCellValue(rowNode.get("responseStatus").asText());
            passedOrFailedCell.setCellValue(rowNode.get("passedOrFailed").asText());
            responsePayloadCell.setCellValue(rowNode.get("responsePayload").asText());


            colNum = 0;
            i+=1;
        }
        Date date = new Date();
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy_MM_dd__HH_mm_a");
        String formattedDate = sdf.format(date);
        String path = "C:/MailingServer/Latest/response_"+formattedDate+".xlsx";
        FileOutputStream outputStream = new FileOutputStream(path);
        wb.write(outputStream);
        wb.close();
        return "Excel file generated";
        } catch (JsonProcessingException e1) {
            e1.printStackTrace();
            return "JsonProcessingException";
        } catch (IOException e1) {
            e1.printStackTrace();
            return "IOException";
        }


    }
}

