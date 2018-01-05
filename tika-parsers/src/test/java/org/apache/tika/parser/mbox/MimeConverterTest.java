package org.apache.tika.parser.mbox;

import com.pff.PSTFile;
import com.pff.PSTFolder;
import com.pff.PSTMessage;
import org.junit.Ignore;
import org.junit.Test;

import javax.mail.internet.MimeMessage;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Vector;

/**
 * Example usage of the MimeConverter class.
 *
 * The following test takes a pst file as an input, and a target directory
 * to write the emails in RFC822 format to.
 *
 * Created by Ben on 19/12/17.
 */
public class MimeConverterTest {

    @Test
    @Ignore
    public void test() {
        try {
            String inputPST = "..."; // input pst file
            String outputDir = "..."; // output directory

            if (!Files.exists(Paths.get(inputPST))) {
                throw new IOException("Could not find input pst file " + inputPST);
            }

            PSTFile pstFile = new PSTFile(inputPST);
            processFolder(pstFile.getRootFolder(), Paths.get(outputDir));
        } catch (Exception err) {
            err.printStackTrace();
        }
    }

    private void processFolder(PSTFolder folder, Path outputDir) throws Exception {

        // go through the folders...
        if (folder.hasSubfolders()) {
            Vector<PSTFolder> childFolders = folder.getSubFolders();
            for (PSTFolder childFolder : childFolders) {
                processFolder(childFolder, outputDir.resolve(childFolder.getDisplayName()));
            }
        }

        // and now the emails for this folder
        if (folder.getContentCount() > 0) {
            PSTMessage message = (PSTMessage)folder.getNextChild();

            int n = 1;
            while (message != null) {
                MimeMessage mimeMessage = MimeConverter.convertToMimeMessage(message);

                Files.createDirectories(outputDir);
                Path destFile = outputDir.resolve(n + ".eml");
                try (FileOutputStream fos = new FileOutputStream(destFile.toFile())) {
                    try {
                        mimeMessage.writeTo(fos);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }

                message = (PSTMessage)folder.getNextChild();
                n++;
            }
        }
    }

}
