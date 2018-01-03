package org.apache.tika.parser.mbox;

import com.google.common.base.Charsets;
import com.pff.PSTAttachment;
import com.pff.PSTException;
import com.pff.PSTMessage;
import com.pff.PSTRecipient;
import org.apache.cxf.jaxrs.ext.multipart.InputStreamDataSource;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.mail.Header;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Session;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.InternetHeaders;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import javax.mail.internet.MimePart;
import javax.mail.util.ByteArrayDataSource;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.Enumeration;
import java.util.Properties;

/**
 * Utility class to convert emails from PSTMessage format to MimeMessage.
 *
 * Created by Ben on 19/12/17.
 */
public class MimeConverter {

    public static MimeMessage convertToMimeMessage(PSTMessage message) throws MessagingException, IOException, PSTException {
        Properties p = System.getProperties();
        Session session = Session.getInstance(p);

        MimeMessage mimeMessage = new MimeMessage(session);

        processHeaders(message, mimeMessage);

        if (message.getNumberOfAttachments() == 0) {
            processContent(message, mimeMessage);
        } else {
            // create root part to hold content and attachments
            MimeMultipart rootMultipart = new MimeMultipart();

            MimeBodyPart contentBodyPart = new MimeBodyPart();
            processContent(message, contentBodyPart);
            rootMultipart.addBodyPart(contentBodyPart);

            processAttachments(message, rootMultipart);
            mimeMessage.setContent(rootMultipart);
        }

        // make sure the headers are consistent with the content
        mimeMessage.saveChanges();

        return mimeMessage;
    }

    private static void processHeaders(PSTMessage message, MimeMessage mimeMessage) throws MessagingException, PSTException, IOException {
        if (!isEmptyOrNull(message.getTransportMessageHeaders())) {
            byte[] hb = message.getTransportMessageHeaders().getBytes("UTF-8");
            InternetHeaders headers = new InternetHeaders(new ByteArrayInputStream(hb));
            headers.removeHeader("Content-Type");

            Enumeration allHeaders = headers.getAllHeaders();
            while (allHeaders.hasMoreElements()) {
                Header header = (Header) allHeaders.nextElement();
                mimeMessage.addHeader(header.getName(), header.getValue());
            }
        } else {
            mimeMessage.setSubject(message.getSubject());
            mimeMessage.setSentDate(message.getClientSubmitTime());

            InternetAddress fromMailbox = new InternetAddress();

            if (!isEmptyOrNull(message.getSenderName())) {
                fromMailbox.setPersonal(message.getSenderName());
            } else {
                fromMailbox.setPersonal(message.getSenderEmailAddress());
            }
            fromMailbox.setAddress(message.getSenderEmailAddress());

            mimeMessage.setFrom(fromMailbox);

            for (int i = 0; i < message.getNumberOfRecipients(); i++) {
                PSTRecipient recipient = message.getRecipient(i);

                switch (recipient.getRecipientType()) {
                    case PSTRecipient.MAPI_TO:
                        mimeMessage.setRecipient(MimeMessage.RecipientType.TO, new InternetAddress(recipient.getEmailAddress(), recipient.getDisplayName()));
                        break;
                    case PSTRecipient.MAPI_CC:
                        mimeMessage.setRecipient(Message.RecipientType.CC, new InternetAddress(recipient.getEmailAddress(), recipient.getDisplayName()));
                        break;
                    case PSTRecipient.MAPI_BCC:
                        mimeMessage.setRecipient(Message.RecipientType.BCC, new InternetAddress(recipient.getEmailAddress(), recipient.getDisplayName()));
                        break;
                }
            }
        }
    }

    private static void processContent(PSTMessage message, MimePart root) throws MessagingException, IOException {

        // no plain text or html body
        if (isEmptyOrNull(message.getBody()) && isEmptyOrNull(message.getBodyHTML())) {
            // set to empty string
            root.setText("", Charsets.UTF_8.name());
        }

        // plain text body only
        if (!isEmptyOrNull(message.getBody()) && isEmptyOrNull(message.getBodyHTML())) {
            root.setText(message.getBody(), Charsets.UTF_8.name());
        }

        // html text body only
        if (isEmptyOrNull(message.getBody()) && !isEmptyOrNull(message.getBodyHTML())) {
            root.setDataHandler(new DataHandler(new ByteArrayDataSource(message.getBodyHTML(), "text/html; charset=utf-8")));
        }

        // plain text body and html body
        if (!isEmptyOrNull(message.getBody()) && !isEmptyOrNull(message.getBodyHTML())) {
            // create a multipart-alternative
            MimeMultipart contentMultipart = new MimeMultipart("alternative");

            MimeBodyPart plainBodyPart = new MimeBodyPart();
            plainBodyPart.setText(message.getBody(), Charsets.UTF_8.name());
            contentMultipart.addBodyPart(plainBodyPart);

            MimeBodyPart htmlBodyPart = new MimeBodyPart();
            htmlBodyPart.setDataHandler(new DataHandler(new ByteArrayDataSource(message.getBodyHTML(), "text/html; charset=utf-8")));
            contentMultipart.addBodyPart(htmlBodyPart);

            root.setContent(contentMultipart);
        }

    }

    private static void processAttachments(PSTMessage message, MimeMultipart rootMultipart) throws PSTException, IOException, MessagingException {
        for (int i = 0; i < message.getNumberOfAttachments(); i++) {
            PSTAttachment attachment = message.getAttachment(i);

            if (attachment != null && attachment.getFileInputStream() != null) {
                MimeBodyPart attachmentBodyPart = new MimeBodyPart();

                if (!isEmptyOrNull(attachment.getMimeTag())) {
                    DataSource source = new InputStreamDataSource(attachment.getFileInputStream(), attachment.getMimeTag());
                    attachmentBodyPart.setDataHandler(new DataHandler(source));
                } else {
                    DataSource source = new InputStreamDataSource(attachment.getFileInputStream(), "application/octet-stream");
                    attachmentBodyPart.setDataHandler(new DataHandler(source));
                }
                attachmentBodyPart.setHeader("Content-Transfer-Encoding", "base64");

                attachmentBodyPart.setContentID(attachment.getContentId());

                String fileName = "";
                if (!isEmptyOrNull(attachment.getLongFilename())) {
                    fileName = attachment.getLongFilename();
                } else if (!isEmptyOrNull(attachment.getDisplayName())) {
                    fileName = attachment.getDisplayName();
                } else if (!isEmptyOrNull(attachment.getFilename())) {
                    fileName = attachment.getFilename();
                }
                attachmentBodyPart.setFileName(fileName);

                rootMultipart.addBodyPart(attachmentBodyPart); // add attachment to root
            }
        }
    }

    private static boolean isEmptyOrNull(String s) {
        return s == null || s.isEmpty();
    }

}
