package org.apache.tika.parser.mail;

import java.io.Serializable;

public class RFC822ParserConfig implements Serializable {

    private static final long serialVersionUID = 0L;

    private boolean isPlainTextBody = false;

    /*
     * When enabled, only plain text parts found in the message body
     * will be parsed. All other content types found in body will be ignored
     */
    public void plainTextBody() {
        this.isPlainTextBody = true;
    }

    public boolean isPlainTextBody() {
        return isPlainTextBody;
    }
}
