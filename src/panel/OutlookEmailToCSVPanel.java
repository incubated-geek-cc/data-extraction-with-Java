package panel;

import util.UtilityManager;
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
import java.awt.event.ActionEvent;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.util.logging.Level;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;
import static util.UtilityManager.copy;
import static util.UtilityManager.getCurrentTimeStamp;

public class OutlookEmailToCSVPanel extends JPanel {
    private final UtilityManager UTILITY_MGR;
    private final JFrame APP_FRAME;

    // input files selected
    DefaultListModel jListInputFilesSelectedModel = new DefaultListModel<>();
    private static JList<String> jListInputFilesSelected;
    private static JScrollPane jScrollPane1FileListItems;

    //initial value, minimum value, maximum value, step
    private static final SpinnerModel JSPINNER_INPUT_HEADER_COUNT_MODEL = new SpinnerNumberModel(13, 1, 100, 1);
    private static JSpinner jSpinnerInputHeaderCount;

    // OUTPUT LOGS
    private static JTextArea LOG_TEXT_AREA;
    private static JScrollPane JSCROLL_PANEL_OUTPUT_LOGS;

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
        LOG_TEXT_AREA = new JTextArea();
        LOG_TEXT_AREA.setEditable(false);
        LOG_TEXT_AREA.setWrapStyleWord(true);
        JSCROLL_PANEL_OUTPUT_LOGS = new JScrollPane(LOG_TEXT_AREA);
        UTILITY_MGR=new UtilityManager(LOG_TEXT_AREA,JSCROLL_PANEL_OUTPUT_LOGS); // so all logs are handled by the same panel
        
        UTILITY_MGR.getLogger().info(() -> "Welcome to Outlook Email Data Extractor.");

        // INPUT FILES SELECTED
        jListInputFilesSelected = new JList<>(jListInputFilesSelectedModel);
        jScrollPane1FileListItems = new JScrollPane(jListInputFilesSelected);

        // INPUT NO OF HEADERS
        jSpinnerInputHeaderCount = new JSpinner(JSPINNER_INPUT_HEADER_COUNT_MODEL);
        
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
        add(JSCROLL_PANEL_OUTPUT_LOGS);
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
        JSCROLL_PANEL_OUTPUT_LOGS.setBounds(20, 285, 775, 220);

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
        UTILITY_MGR.getLogger().info(() -> "Welcome to Outlook Email Data Extractor.");
    }

    private void runAppAction(ActionEvent e) {
        noOfHeaders = (Integer) jSpinnerInputHeaderCount.getValue();
        jButtonSelectInputFiles.setEnabled(false);
        jButtonResetAll.setEnabled(false);
        jButtonRemoveSelectedFiles.setEnabled(false);

        UTILITY_MGR.outputConsoleLogsBreakline("");
        UTILITY_MGR.outputConsoleLogsBreakline("Initialising Outlook Email Data Extractor");
        UTILITY_MGR.outputConsoleLogsBreakline("");
        UTILITY_MGR.updateLogs();
        try {
            UTILITY_MGR.outputConsoleLogsBreakline("Reading in Outlook files");
            UTILITY_MGR.updateLogs();
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
            UTILITY_MGR.getLogger().log(Level.SEVERE, null, ex);
        }

        jButtonRun.setEnabled(false);
        jButtonRemoveSelectedFiles.setEnabled(true);
        jButtonResetAll.setEnabled(true);
        jButtonSelectInputFiles.setEnabled(true);
    }

    private void inputOutlook(ArrayList<File> outlookFiles) throws IOException, FileNotFoundException {
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
                
                UTILITY_MGR.outputConsoleLogsBreakline("Processing "+outlookFileName);
                UTILITY_MGR.updateLogs();
                
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
                UTILITY_MGR.outputConsoleLogsBreakline(outputFile.getName() + " data has been extracted.");
                UTILITY_MGR.updateLogs();

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
}
