package com.emploi.controllers;

import java.io.File;

import javax.mail.MessagingException;
import javax.mail.internet.MimeMessage;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.springframework.stereotype.Component;


@Component
public class SmtpMailSender {
 
@Autowired
private JavaMailSender javaMailsender;	

public void send(String to, String subject, String body, File file ) throws MessagingException{
	MimeMessage message = javaMailsender.createMimeMessage();
	MimeMessageHelper helper;
	
	helper = new MimeMessageHelper(message, true);
	helper.setSubject(subject);
	helper.setTo(to);
	helper.setText(body, true);
	helper.addAttachment(file.getName(), file);
	
	javaMailsender.send(message);
}
}
