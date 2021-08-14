package panel;

import java.awt.*;
import java.io.File;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.text.DefaultCaret;
import java.awt.event.ActionEvent;
import java.io.IOException;
import java.lang.invoke.MethodHandles;
import java.util.HashMap;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.pdfbox.multipdf.PDFMergerUtility;
import org.apache.pdfbox.pdmodel.PDDocument;
import static panel.Util.copy;
import static panel.Util.getCurrentTimeStamp;
import static panel.Util.outputConsoleLogsBreakline;

public class MergePdfFilesPanel extends JPanel {

    private final JFrame APP_FRAME;
    private static final String LOGGER_NAME = MethodHandles.lookup().lookupClass().getName();
    private static final Logger LOGGER = Logger.getLogger(LOGGER_NAME);

    // input files selected
    DefaultListModel jListInputFilesSelectedModel = new DefaultListModel<>();
    private static JList<String> jListInputFilesSelected;
    private static JScrollPane jScrollPane1FileListItems;

    // OUTPUT LOGS
    private static final JTextArea LOG_TEXT_AREA = new JTextArea();
    private static JScrollPane jScrollPane1OutputFileLogs;

    private static JLabel jLabelFileChooserText;

    private static JLabel jLabelOutputFileLogsTitle;
    private static JLabel jLabelFileListSelected;

    private static JButton jButtonSelectInputFiles;
    private static JButton jButtonResetAll;
    private static JButton jButtonRemoveSelectedFiles;

    private static JButton jButtonRun;

    // LIST OF FILE ITEMS - INPUT FILES TO MERGE INTO OUTPUT PDF
    private static final ArrayList<File> INPUT_FILES = new ArrayList<File>();

    private static File outputMergedPdf = null;
    private static PDFMergerUtility PDFMerger;

    public MergePdfFilesPanel(JFrame APP_FRAME) {
        super();
        this.APP_FRAME = APP_FRAME;
        LOGGER.setUseParentHandlers(false);
        LOGGER.setLevel(Level.ALL);
        LOGGER.addHandler(new TextAreaHandler(new TextAreaOutputStream(LOG_TEXT_AREA)));

        LOGGER.info(() -> "Welcome to Pdf Merger.");

        // INPUT FILES SELECTED
        jListInputFilesSelected = new JList<>(jListInputFilesSelectedModel);
        jScrollPane1FileListItems = new JScrollPane(jListInputFilesSelected);

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

        jLabelFileListSelected = new JLabel("List of input files selected:");
        jLabelOutputFileLogsTitle = new JLabel("Output File Log(s):");

        // set components properties
        jButtonRun.setEnabled(false);

        //add components
        add(jLabelFileChooserText);
        add(jButtonSelectInputFiles);
        
        add(jButtonRemoveSelectedFiles);
        add(jScrollPane1FileListItems);
        
        add(new JSeparator());  
        
        add(jButtonRun);
        add(jLabelOutputFileLogsTitle);
        add(jScrollPane1OutputFileLogs);
        add(jButtonResetAll);

        //set component bounds (only needed by Absolute Positioning)
        jLabelFileChooserText.setBounds(20, 15, 795, 30);
        jButtonSelectInputFiles.setBounds(160, 15, 130, 30);

        //jButtonResetAll.setBounds(685, 35, 130, 30);
        jButtonRemoveSelectedFiles.setBounds(665, 15, 130, 30);//535, 35, 130, 30);

        jLabelFileListSelected.setBounds(395, 15, 400, 30);
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

        FileNameExtensionFilter filter = new FileNameExtensionFilter("Pdf File (.pdf)", "pdf");
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
        LOGGER.info(() -> "Welcome to Pdf Merger.");
    }

    private void runAppAction(ActionEvent e) {
        jButtonSelectInputFiles.setEnabled(false);
        jButtonResetAll.setEnabled(false);
        jButtonRemoveSelectedFiles.setEnabled(false);

        outputConsoleLogsBreakline(LOGGER, "");
        outputConsoleLogsBreakline(LOGGER, "Initialising Pdf Merger App");
        outputConsoleLogsBreakline(LOGGER, "");

        try {
            outputConsoleLogsBreakline(LOGGER, "Reading in pdf files");
            // ================================================= READ IN FILES ================================
            PDFMerger = new PDFMergerUtility();
            String destinationFileName="output_" + getCurrentTimeStamp() + ".pdf";
            outputMergedPdf = new File(destinationFileName);
            
            JFileChooser saveFileChooser = new JFileChooser();
            saveFileChooser.setDialogTitle("Save Output As...");
            saveFileChooser.setDialogType(JFileChooser.SAVE_DIALOG);
            
            saveFileChooser.setSelectedFile(outputMergedPdf);
            saveFileChooser.setFileFilter(new FileNameExtensionFilter("Pdf file (*.pdf)", "pdf"));
            
            int option = saveFileChooser.showSaveDialog(APP_FRAME);
            if (option == JFileChooser.APPROVE_OPTION) {
                File selectedFile = saveFileChooser.getSelectedFile();
                if (selectedFile != null) {
                    if (!selectedFile.getName().toLowerCase().endsWith(".pdf")) {
                        selectedFile = new File(selectedFile.getParentFile(), selectedFile.getName() + ".pdf");
                    }
                    PDFMerger.setDestinationFileName(destinationFileName);
                    // merge PDF Documents here
                    mergePdfDocuments(INPUT_FILES);
                    copy(outputMergedPdf, selectedFile);
                    Desktop.getDesktop().open(selectedFile);
                    outputMergedPdf.delete();
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

    private static void mergePdfDocuments(ArrayList<File> pdfFilesToMerge) throws FileNotFoundException {
        ArrayList<PDDocument> pdfDocs = new ArrayList<PDDocument>();
        
        for (File pdfFile : pdfFilesToMerge) {
            HashMap<String, Object> files = new HashMap<String, Object>();
            PDDocument pdfDocument = null;
            try {
                pdfDocument = PDDocument.load(pdfFile);
                outputConsoleLogsBreakline(LOGGER, "Reading " + pdfFile.getName());
                updateLogs();
            } catch (IOException ex) {
                LOGGER.log(Level.SEVERE, null, ex);
            }
            files.put("File", pdfFile);
            files.put("PDF", pdfDocument);
            File file = (File) files.get("File");

            PDDocument doc = (PDDocument) files.get("PDF");
            PDFMerger.addSource(file);
            pdfDocs.add(doc);
        }
        try {
            PDFMerger.mergeDocuments();
        } catch (IOException ex) {
            LOGGER.log(Level.SEVERE, null, ex);
        }
        outputConsoleLogsBreakline(LOGGER, "Pdf Documents have been merged successfully.");

        for (PDDocument pdfDoc : pdfDocs) {
            try {
                pdfDoc.close();
            } catch (IOException ex) {
                LOGGER.log(Level.SEVERE, null, ex);
            }
        }
    }

    private static void updateLogs() {
        jScrollPane1OutputFileLogs.getVerticalScrollBar().setValue(jScrollPane1OutputFileLogs.getVerticalScrollBar().getMaximum());
        jScrollPane1OutputFileLogs.getVerticalScrollBar().paint(jScrollPane1OutputFileLogs.getVerticalScrollBar().getGraphics());
        LOG_TEXT_AREA.scrollRectToVisible(LOG_TEXT_AREA.getVisibleRect());
        LOG_TEXT_AREA.paint(LOG_TEXT_AREA.getGraphics());
    }
}
