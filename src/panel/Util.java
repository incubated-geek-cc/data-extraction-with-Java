package panel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.logging.Logger;

public class Util {

    public static void copy(File src, File dest) throws IOException {
        InputStream is = null;
        OutputStream os = null;
        try {
            is = new FileInputStream(src);
            os = new FileOutputStream(dest);
            // buffer size 1K
            byte[] buf = new byte[1024];
            int bytesRead;
            while ((bytesRead = is.read(buf)) > 0) {
                os.write(buf, 0, bytesRead);
            }
        } finally {
            is.close();
            os.close();
        }
    }

    public static String getCurrentTimeStamp() {
        SimpleDateFormat sdf = new SimpleDateFormat("(dd-MMM-yyyy_hhmmaa)");
        Date date = new Date();
        String timestamp = sdf.format(date);

        return timestamp;
    }

    private static void addTextToOutputLogs(Logger LOGGER, String logString) {
        LOGGER.info(() -> logString);
    }

    public static void outputConsoleLogsBreakline(Logger LOGGER, String consoleCaption) {
        String logString = "";

        int charLimit = 180;
        if (consoleCaption.length() > charLimit) {
            logString = consoleCaption.substring(0, charLimit - 4) + " ...";
        } else {
            String result = "";

            if (consoleCaption.isEmpty()) {
                for (int i = 0; i < charLimit; i++) {
                    result += "=";
                }
                logString = result;
            } else {
                charLimit = (charLimit - consoleCaption.length() - 1);
                for (int i = 0; i < charLimit; i++) {
                    result += "-";
                }
                logString = consoleCaption + " " + result;
            }
        }
        logString = logString + "\n";
        addTextToOutputLogs(LOGGER, logString);
    }
}
