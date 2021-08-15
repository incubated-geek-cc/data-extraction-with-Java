package panel;

import au.com.bytecode.opencsv.CSVWriter;
import com.auxilii.msgparser.Message;
import com.auxilii.msgparser.MsgParser;
import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.text.DefaultCaret;
import java.awt.event.ActionEvent;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.lang.invoke.MethodHandles;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;
import static panel.Util.copy;
import static panel.Util.getCurrentTimeStamp;
import static panel.Util.outputConsoleLogsBreakline;

public class OutlookEmailToCSVPanel extends JPanel {

    private final JFrame APP_FRAME;
    private static final String LOGGER_NAME = MethodHandles.lookup().lookupClass().getName();
    private static final Logger LOGGER = Logger.getLogger(LOGGER_NAME);

    // input files selected
    DefaultListModel jListInputFilesSelectedModel = new DefaultListModel<>();
    private static JList<String> jListInputFilesSelected;
    private static JScrollPane jScrollPane1FileListItems;

    //initial value, minimum value, maximum value, step
    private static SpinnerModel jSpinnerInputHeaderCountModel = new SpinnerNumberModel(13, 1, 100, 1);
    private static JSpinner jSpinnerInputHeaderCount;

    // OUTPUT LOGS
    private static final JTextArea LOG_TEXT_AREA = new JTextArea();
    private static JScrollPane jScrollPane1OutputFileLogs;

    private static JLabel jLabelFileChooserText;
    private static JLabel jLabelInputHeaderCount;
    private static JLabel jLabelOutputFileLogsTitle;
    private static JLabel jLabelFileListSelected;

    private static JButton jButtonSelectInputFiles;
    private static JButton jButtonResetAll;
    private static JButton jButtonRemoveSelectedFiles;
    private static JButton jButtonRun;

    // LIST OF FILE ITEMS - INPUT FILES TO COMPILE INTO ARCHIVE
    private static final ArrayList<File> INPUT_FILES = new ArrayList<File>();
    private static int noOfHeaders = 13;
    private static File outputArchiveZip = null;

    public OutlookEmailToCSVPanel(JFrame APP_FRAME) {
        super();
        this.APP_FRAME = APP_FRAME;
        LOGGER.setUseParentHandlers(false);
        LOGGER.setLevel(Level.ALL);
        LOGGER.addHandler(new TextAreaHandler(new TextAreaOutputStream(LOG_TEXT_AREA)));

        LOGGER.info(() -> "Welcome to Outlook Email Data Extractor.");

        // INPUT FILES SELECTED
        jListInputFilesSelected = new JList<>(jListInputFilesSelectedModel);
        jScrollPane1FileListItems = new JScrollPane(jListInputFilesSelected);

        // INPUT NO OF HEADERS
        jSpinnerInputHeaderCount = new JSpinner(jSpinnerInputHeaderCountModel);

        // OUTPUT LOGS
        LOG_TEXT_AREA.setEditable(false);
        LOG_TEXT_AREA.setWrapStyleWord(true);
        jScrollPane1OutputFileLogs = new JScrollPane(LOG_TEXT_AREA);

        updateLogs();
        jScrollPane1OutputFileLogs.setHorizontalScrollBar(null);

        DefaultCaret caret = (DefaultCaret) LOG_TEXT_AREA.getCaret();
        caret.setUpdatePolicy(DefaultCaret.ALWAYS_UPDATE);

        // ACTIONABLE BUTTONS
        jButtonSelectInputFiles = new JButton("Choose File(s)");
        jButtonRemoveSelectedFiles = new JButton("Remove File");
        jButtonResetAll = new JButton("Reset All");

        jButtonRun = new JButton("Run >>");

        jLabelFileChooserText = new JLabel("Select input file(s)");
        jLabelInputHeaderCount = new JLabel("No. of header field(s)");

        jLabelFileListSelected = new JLabel("List of input files selected:");
        jLabelOutputFileLogsTitle = new JLabel("Output File Log(s):");

        // set components properties
        jButtonRun.setEnabled(false);

        //add components
        add(jLabelFileChooserText);
        add(jButtonSelectInputFiles);

        add(jLabelInputHeaderCount);
        add(jSpinnerInputHeaderCount);

        add(jButtonRemoveSelectedFiles);
        add(jLabelFileListSelected);
        add(jScrollPane1FileListItems);

        add(jButtonRun);
        add(jLabelOutputFileLogsTitle);
        add(jScrollPane1OutputFileLogs);
        add(jButtonResetAll);

        // set component bounds (only needed by Absolute Positioning)
        jLabelFileChooserText.setBounds(20, 15, 795, 30);
        jButtonSelectInputFiles.setBounds(160, 15, 130, 30);

        jLabelInputHeaderCount.setBounds(20, 50, 795, 30);
        jSpinnerInputHeaderCount.setBounds(160, 50, 130, 30);

        jButtonRemoveSelectedFiles.setBounds(665, 15, 130, 30);
        jLabelFileListSelected.setBounds(395, 15, 200, 30);
        jScrollPane1FileListItems.setBounds(395, 50, 400, 195);

        jButtonRun.setBounds(20, 215, 130, 30);
        jLabelOutputFileLogsTitle.setBounds(20, 255, 775, 30);
        jScrollPane1OutputFileLogs.setBounds(20, 285, 775, 220);

        jButtonResetAll.setBounds(665, 515, 130, 30);

        jButtonSelectInputFiles.addActionListener((java.awt.event.ActionEvent evt) -> {
            selectInputFilesAction(evt);
        });

        jButtonResetAll.addActionListener((java.awt.event.ActionEvent evt) -> {
            resetAllAction(evt);
        });

        jButtonRun.addActionListener((java.awt.event.ActionEvent evt) -> {
            runAppAction(evt);
        });

        jButtonRemoveSelectedFiles.addActionListener((java.awt.event.ActionEvent evt) -> {
            removeListItemAction(evt);
            if (INPUT_FILES.isEmpty()) {
                jButtonRun.setEnabled(false);
            }
        });
    }

    private void selectInputFilesAction(ActionEvent e) {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Select Input File(s)");

        fileChooser.setMultiSelectionEnabled(true);
        fileChooser.setAcceptAllFileFilterUsed(false);

        FileNameExtensionFilter filter = new FileNameExtensionFilter("Outlook File (.msg)", "msg");
        fileChooser.addChoosableFileFilter(filter);
        filter = new FileNameExtensionFilter("Outlook File (.eml)", "eml");
        fileChooser.addChoosableFileFilter(filter);

        int option = fileChooser.showOpenDialog(APP_FRAME);
        if (option == JFileChooser.APPROVE_OPTION) {
            jListInputFilesSelectedModel = (DefaultListModel) jListInputFilesSelected.getModel();
            File[] selectedFiles = fileChooser.getSelectedFiles();
            for (File selectedFile : selectedFiles) { // FOR-EACH FILE
                String selectedFileName = selectedFile.getName();
                jListInputFilesSelectedModel.addElement(selectedFileName);
                INPUT_FILES.add(selectedFile);
            }
            if (INPUT_FILES.size() > 0) {
                jButtonRun.setEnabled(true);
            }
        }
    }

    private void removeListItemAction(ActionEvent e) {
        jListInputFilesSelectedModel = (DefaultListModel) jListInputFilesSelected.getModel();
        int[] selectedInputFiles = jListInputFilesSelected.getSelectedIndices();

        for (int i : selectedInputFiles) {
            jListInputFilesSelectedModel.remove(i);
            INPUT_FILES.remove(i);
        }
    }

    private void resetAllAction(ActionEvent e) {
        jButtonSelectInputFiles.setEnabled(true);
        jButtonRun.setEnabled(false);

        jListInputFilesSelectedModel.clear();
        INPUT_FILES.clear();
        LOG_TEXT_AREA.setText("");
        LOGGER.info(() -> "Welcome to Outlook Email Data Extractor.");
    }

    private void runAppAction(ActionEvent e) {
        noOfHeaders = (Integer) jSpinnerInputHeaderCount.getValue();
        jButtonSelectInputFiles.setEnabled(false);
        jButtonResetAll.setEnabled(false);
        jButtonRemoveSelectedFiles.setEnabled(false);

        outputConsoleLogsBreakline(LOGGER, "");
        outputConsoleLogsBreakline(LOGGER, "Initialising Outlook Email Data Extractor");
        outputConsoleLogsBreakline(LOGGER, "");
        updateLogs();
        try {
            outputConsoleLogsBreakline(LOGGER, "Reading in Outlook files");
            updateLogs();
            // ================================================= READ IN FILES ================================
            inputOutlook(INPUT_FILES);
            JFileChooser saveFileChooser = new JFileChooser();
            saveFileChooser.setDialogTitle("Save Output As...");
            saveFileChooser.setDialogType(JFileChooser.SAVE_DIALOG);

            saveFileChooser.setSelectedFile(outputArchiveZip);
            saveFileChooser.setFileFilter(new FileNameExtensionFilter("ZIP (*.zip)", "zip"));

            int option = saveFileChooser.showSaveDialog(APP_FRAME);
            if (option == JFileChooser.APPROVE_OPTION) {
                File selectedFile = saveFileChooser.getSelectedFile();
                if (selectedFile != null) {
                    if (!selectedFile.getName().toLowerCase().endsWith(".zip")) {
                        selectedFile = new File(selectedFile.getParentFile(), selectedFile.getName() + ".zip");
                    }
                    copy(outputArchiveZip, selectedFile);
                    Desktop.getDesktop().open(selectedFile);
                    outputArchiveZip.delete();
                }
            }
        } catch (IOException ex) {
            LOGGER.log(Level.SEVERE, null, ex);
        }

        jButtonRun.setEnabled(false);
        jButtonRemoveSelectedFiles.setEnabled(true);
        jButtonResetAll.setEnabled(true);
        jButtonSelectInputFiles.setEnabled(true);
    }

    private static void inputOutlook(ArrayList<File> outlookFiles) throws IOException, FileNotFoundException {
        String outlookFileName = "";
        String outlookFilePath = "";

        outputArchiveZip = new File("OutlookEmailToCSV_" + getCurrentTimeStamp() + ".zip");
        try (FileOutputStream fos = new FileOutputStream(outputArchiveZip)) {
            ZipOutputStream zipOut = new ZipOutputStream(fos);

            FileOutputStream os = null;
            File outputFile = null;
            CSVWriter writer = null;

            for (File outlookFile : outlookFiles) {
                outlookFileName = outlookFile.getName();
                outlookFilePath = outlookFile.getAbsolutePath();
                
                outputConsoleLogsBreakline(LOGGER, "Processing "+outlookFileName);
                updateLogs();
                
                MsgParser msgp = new MsgParser();
                String bodyText = "";
                if (outlookFile.getName().endsWith(".msg")) {
                    Message msg = msgp.parseMsg(outlookFile);
                    bodyText = msg.getBodyText();
                } else if (outlookFile.getName().endsWith(".eml")) {
                    bodyText = new String(Files.readAllBytes(outlookFile.toPath()), StandardCharsets.UTF_8);
                }
                String[] strArr = bodyText.split("\\r\\n");
                String outputCsvFileName = outlookFileName + ".csv";
                outputFile = new File(outputCsvFileName);
                os = new FileOutputStream(outputFile);
                os.write(0xef);
                os.write(0xbb);
                os.write(0xbf);

                /*char textDelimiter = (char) jComboBoxDelimiterChoice.getSelectedItem();
                char textQualifier = (char) jComboBoxTextQualifierChoice.getSelectedItem();*/
                writer = new CSVWriter(new OutputStreamWriter(os), ',', '"');
                int counter = 0;
                ArrayList<String> values = new ArrayList<String>();
                for (String str : strArr) {
                    str = str.trim();
                    if (str.isEmpty()) {
                        str = "-";
                    }
                    if (counter > 0) {
                        if (counter % noOfHeaders == 0) {
                            String[] valuesArr = new String[values.size()];
                            for (int v = 0; v < values.size(); v++) {
                                valuesArr[v] = values.get(v);
                            }
                            writer.writeNext(valuesArr);
                            values = new ArrayList<String>();
                        }
                    }
                    values.add(str);
                    counter++;
                }
                writer.close();
                // CSV file has been written to
                outputConsoleLogsBreakline(LOGGER, outputFile.getName() + " data has been extracted.");
                updateLogs();

                File fileToZip = new File(outputFile.getAbsolutePath());
                FileInputStream fis = new FileInputStream(fileToZip);
                ZipEntry zipEntry = new ZipEntry(fileToZip.getName());
                zipOut.putNextEntry(zipEntry);

                byte[] bytes = new byte[1024];
                int length;
                while ((length = fis.read(bytes)) >= 0) {
                    zipOut.write(bytes, 0, length);
                }
                fis.close();
                outputFile.delete();
            } // for-loop each ExcelFile
            zipOut.close();
        }
    }

    private static void updateLogs() {
        jScrollPane1OutputFileLogs.getVerticalScrollBar().setValue(jScrollPane1OutputFileLogs.getVerticalScrollBar().getMaximum());
        jScrollPane1OutputFileLogs.getVerticalScrollBar().paint(jScrollPane1OutputFileLogs.getVerticalScrollBar().getGraphics());
        LOG_TEXT_AREA.scrollRectToVisible(LOG_TEXT_AREA.getVisibleRect());
        LOG_TEXT_AREA.paint(LOG_TEXT_AREA.getGraphics());
    }
}
