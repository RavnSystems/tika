/*
 * Licensed to the Apache Software Foundation (ASF) under one or more
 * contributor license agreements.  See the NOTICE file distributed with
 * this work for additional information regarding copyright ownership.
 * The ASF licenses this file to You under the Apache License, Version 2.0
 * (the "License"); you may not use this file except in compliance with
 * the License.  You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package org.apache.tika.parser.mbox;

import static java.lang.String.valueOf;
import static java.util.Collections.singleton;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Set;

import com.pff.PSTFile;
import com.pff.PSTFolder;
import com.pff.PSTMessage;
import org.apache.tika.exception.TikaException;
import org.apache.tika.extractor.EmbeddedDocumentExtractor;
import org.apache.tika.extractor.EmbeddedDocumentUtil;
import org.apache.tika.io.IOUtils;
import org.apache.tika.io.TikaInputStream;
import org.apache.tika.metadata.Metadata;
import org.apache.tika.mime.MediaType;
import org.apache.tika.parser.AbstractParser;
import org.apache.tika.parser.ParseContext;

import org.apache.tika.sax.XHTMLContentHandler;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.ContentHandler;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.AttributesImpl;

import javax.mail.internet.MimeMessage;

/**
 * Parser for MS Outlook PST email storage files
 */
public class OutlookPSTParser extends AbstractParser {

    private static final Logger LOGGER = LoggerFactory.getLogger(OutlookPSTParser.class);

    private static final long serialVersionUID = 620998217748364063L;

    public static final MediaType MS_OUTLOOK_PST_MIMETYPE = MediaType.application("vnd.ms-outlook-pst");
    private static final Set<MediaType> SUPPORTED_TYPES = singleton(MS_OUTLOOK_PST_MIMETYPE);

    private static AttributesImpl createAttribute(String attName, String attValue) {
        AttributesImpl attributes = new AttributesImpl();
        attributes.addAttribute("", attName, attName, "CDATA", attValue);
        return attributes;
    }


    public Set<MediaType> getSupportedTypes(ParseContext context) {
        return SUPPORTED_TYPES;
    }

    public void parse(InputStream stream, ContentHandler handler, Metadata metadata, ParseContext context)
            throws IOException, SAXException, TikaException {

        metadata.set(Metadata.CONTENT_TYPE, MS_OUTLOOK_PST_MIMETYPE.toString());

        XHTMLContentHandler xhtml = new XHTMLContentHandler(handler, metadata);
        xhtml.startDocument();

        TikaInputStream in = TikaInputStream.get(stream);
        PSTFile pstFile = null;
        try {
            pstFile = new PSTFile(in.getFile().getPath());
            metadata.set(Metadata.CONTENT_LENGTH, valueOf(pstFile.getFileHandle().length()));
            boolean isValid = pstFile.getFileHandle().getFD().valid();
            metadata.set("isValid", valueOf(isValid));
            if (isValid) {
                // Use the delegate parser to parse the contained document
                EmbeddedDocumentExtractor embeddedExtractor = EmbeddedDocumentUtil.getEmbeddedDocumentExtractor(context);

                parseFolder(xhtml, pstFile.getRootFolder(), embeddedExtractor);
            }
        } catch (Exception e) {
            throw new TikaException(e.getMessage(), e);
        } finally {
            if (pstFile != null && pstFile.getFileHandle() != null) {
                try {
                    pstFile.getFileHandle().close();
                } catch (IOException e) {
                    //swallow closing exception
                }
            }
        }

        xhtml.endDocument();
    }

    private void parseFolder(XHTMLContentHandler handler, PSTFolder pstFolder, EmbeddedDocumentExtractor embeddedExtractor)
            throws Exception {
        if (pstFolder.getContentCount() > 0) {
            PSTMessage pstMail = (PSTMessage) pstFolder.getNextChild();
            while (pstMail != null) {
                AttributesImpl attributes = new AttributesImpl();
                attributes.addAttribute("", "class", "class", "CDATA", "embedded");
                attributes.addAttribute("", "id", "id", "CDATA", pstMail.getInternetMessageId());
                handler.startElement("div", attributes);
                handler.element("h1", pstMail.getSubject());

                // parse the message
                parserMailItem(handler, pstMail, embeddedExtractor);

                handler.endElement("div");

                pstMail = (PSTMessage) pstFolder.getNextChild();
            }
        }

        if (pstFolder.hasSubfolders()) {
            for (PSTFolder pstSubFolder : pstFolder.getSubFolders()) {
                handler.startElement("div", createAttribute("class", "email-folder"));
                handler.element("h1", pstSubFolder.getDisplayName());
                parseFolder(handler, pstSubFolder, embeddedExtractor);
                handler.endElement("div");
            }
        }
    }

    private void parserMailItem(XHTMLContentHandler handler, PSTMessage pstMail,
                                EmbeddedDocumentExtractor embeddedExtractor) throws TikaException {

        try {
            MimeMessage message = MimeConverter.convertToMimeMessage(pstMail);

            ByteArrayOutputStream output = new ByteArrayOutputStream();

            // configure the correct class loader to use in an OSGI context
            ClassLoader contextClassLoader = Thread.currentThread().getContextClassLoader();
            Thread.currentThread().setContextClassLoader(MimeMessage.class.getClassLoader());
            try {
                message.writeTo(output);
            } finally {
                Thread.currentThread().setContextClassLoader(contextClassLoader);
            }

            final Metadata metadata = new Metadata();
            metadata.set(Metadata.RESOURCE_NAME_KEY, pstMail.getInternetMessageId());
            metadata.set(Metadata.EMBEDDED_RELATIONSHIP_ID, pstMail.getInternetMessageId());

            final String rfc822 = output.toString("UTF-8");
            try (InputStream is = IOUtils.toInputStream(rfc822, "UTF-8")) {
                embeddedExtractor.parseEmbedded(is, handler, metadata, false);
            }
        } catch (Throwable t) {
            LOGGER.error("Failed to parse mail item", t);
        }
    }

}