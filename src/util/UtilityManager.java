package util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.invoke.MethodHandles;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.text.DefaultCaret;

public class UtilityManager {

    private static final String LOGGER_NAME = MethodHandles.lookup().lookupClass().getName();
    private static final Logger LOGGER = Logger.getLogger(LOGGER_NAME);

    private final JTextArea LOG_TEXT_AREA;
    private final JScrollPane JSCROLL_PANEL_OUTPUT_LOGS;

    public UtilityManager(JTextArea LOG_TEXT_AREA, JScrollPane JSCROLL_PANEL_OUTPUT_LOGS) {
        this.LOG_TEXT_AREA = LOG_TEXT_AREA;
        this.JSCROLL_PANEL_OUTPUT_LOGS = JSCROLL_PANEL_OUTPUT_LOGS;
        JSCROLL_PANEL_OUTPUT_LOGS.setHorizontalScrollBar(null);
        DefaultCaret caret = (DefaultCaret) LOG_TEXT_AREA.getCaret();
        caret.setUpdatePolicy(DefaultCaret.ALWAYS_UPDATE);

        LOGGER.setUseParentHandlers(false);
        LOGGER.setLevel(Level.ALL);
        LOGGER.addHandler(new TextAreaHandler(new TextAreaOutputStream(LOG_TEXT_AREA)));
    }

    public Logger getLogger() {
        return LOGGER;
    }

    public void updateLogs() {
        JSCROLL_PANEL_OUTPUT_LOGS.getVerticalScrollBar().setValue(JSCROLL_PANEL_OUTPUT_LOGS.getVerticalScrollBar().getMaximum());
        JSCROLL_PANEL_OUTPUT_LOGS.getVerticalScrollBar().paint(JSCROLL_PANEL_OUTPUT_LOGS.getVerticalScrollBar().getGraphics());
        LOG_TEXT_AREA.scrollRectToVisible(LOG_TEXT_AREA.getVisibleRect());
        LOG_TEXT_AREA.paint(LOG_TEXT_AREA.getGraphics());
    }

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

    public static String getColumnName(int columnNumber) {
        String columnName = "";
        int dividend = columnNumber + 1;
        int modulus;

        while (dividend > 0) {
            modulus = (dividend - 1) % 26;
            columnName = (char) (65 + modulus) + columnName;
            dividend = (int) ((dividend - modulus) / 26);
        }

        return columnName;
    }

    public static String getDateStrFromExcelNumber(String cellStrValue) {
        String cellStrValueResult = cellStrValue;
        try {
            double cellValDouble = Double.parseDouble(cellStrValue); // 
            cellStrValueResult = convertRawValueToDate(new BigDecimal(cellValDouble).longValue());
        } catch (NumberFormatException nfe) {
            nfe.printStackTrace();
        }
        return cellStrValueResult;
    }

    public static String convertRawValueToDate(Long dateLong) {
        Long daysToAdd = dateLong - 25569 - 1;
        LocalDate localDate = LocalDate.of(1970, 01, 02).plusDays(daysToAdd);
        ZoneId defaultZoneId = ZoneId.systemDefault();
        Date resultDate = Date.from(localDate.atStartOfDay(defaultZoneId).toInstant());

        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        String resultDateStr = sdf.format(resultDate);

        return resultDateStr;
    }

    public static String getCurrentTimeStamp() {
        SimpleDateFormat sdf = new SimpleDateFormat("(dd-MMM-yyyy_hhmmaa)");
        Date date = new Date();
        String timestamp = sdf.format(date);

        return timestamp;
    }

    private static void addTextToOutputLogs(String logString) {
        LOGGER.info(() -> logString);
    }

    public void outputConsoleLogsBreakline(String consoleCaption) {
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
        addTextToOutputLogs(logString);
    }
}
