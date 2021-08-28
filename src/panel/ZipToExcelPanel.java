package panel;

import util.UtilityManager;
import au.com.bytecode.opencsv.CSVReader;
import java.awt.*;
import java.io.File;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.awt.event.ActionEvent;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.util.Enumeration;
import java.util.logging.Level;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import static util.UtilityManager.copy;
import static util.UtilityManager.getCurrentTimeStamp;

public class ZipToExcelPanel extends JPanel {
    private final UtilityManager UTILITY_MGR;
    private final JFrame APP_FRAME;

    // input files selected
    DefaultListModel jListInputFilesSelectedModel = new DefaultListModel<>();
    private static JList<String> jListInputFilesSelected;
    private static JScrollPane jScrollPane1FileListItems;

    // OUTPUT LOGS
    private static JTextArea LOG_TEXT_AREA;
    private static JScrollPane JSCROLL_PANEL_OUTPUT_LOGS;

    private static JLabel jLabelFileChooserText;

    private static JLabel jLabelOutputFileLogsTitle;
    private static JLabel jLabelFileListSelected;
    
    private static JButton jButtonSelectInputFiles;
    private static JButton jButtonResetAll;
    private static JButton jButtonRemoveSelectedFiles;

    private static JButton jButtonRun;

    // LIST OF FILE ITEMS - INPUT FILES TO COMPILE INTO ARCHIVE
    private static final ArrayList<File> INPUT_FILES = new ArrayList<File>();
    
    private static String workbookName;
    private static File outputCompiledExcel;
    private static Workbook workbook;

    public ZipToExcelPanel(JFrame APP_FRAME) {
        super();
        this.APP_FRAME = APP_FRAME;
        LOG_TEXT_AREA = new JTextArea();
        LOG_TEXT_AREA.setEditable(false);
        LOG_TEXT_AREA.setWrapStyleWord(true);
        JSCROLL_PANEL_OUTPUT_LOGS = new JScrollPane(LOG_TEXT_AREA);
        UTILITY_MGR=new UtilityManager(LOG_TEXT_AREA,JSCROLL_PANEL_OUTPUT_LOGS); // so all logs are handled by the same panel
        
        UTILITY_MGR.getLogger().info(() ->  "Welcome to Zip to Excel Data Extractor.");
        
        // INPUT FILES SELECTED
        jListInputFilesSelected = new JList<>(jListInputFilesSelectedModel);
        jScrollPane1FileListItems = new JScrollPane(jListInputFilesSelected);
 
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
        add(jLabelFileListSelected);
        add(jScrollPane1FileListItems);

        add(jButtonRun);
        add(jLabelOutputFileLogsTitle);
        add(JSCROLL_PANEL_OUTPUT_LOGS);
        add(jButtonResetAll);

        // set component bounds (only needed by Absolute Positioning)
        jLabelFileChooserText.setBounds(20, 15, 795, 30);
        jButtonSelectInputFiles.setBounds(160, 15, 130, 30);

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

        FileNameExtensionFilter filter = new FileNameExtensionFilter("Zip (.zip)", "zip");
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
        UTILITY_MGR.getLogger().info(() -> "Welcome to Zip to Excel Data Extractor.");
    }

    private void runAppAction(ActionEvent e) {
        workbook = new XSSFWorkbook();
        
        jButtonSelectInputFiles.setEnabled(false);
        jButtonResetAll.setEnabled(false);
        jButtonRemoveSelectedFiles.setEnabled(false);

        UTILITY_MGR.outputConsoleLogsBreakline("");
        UTILITY_MGR.outputConsoleLogsBreakline("Initialising Zip to Excel Data Extractor");
        UTILITY_MGR.outputConsoleLogsBreakline("");
        UTILITY_MGR.updateLogs();
        
        try {
            UTILITY_MGR.outputConsoleLogsBreakline("Reading in Zip files");
            // ================================================= READ IN FILES ================================
            inputZipFiles(INPUT_FILES);
            JFileChooser saveFileChooser = new JFileChooser();
            saveFileChooser.setDialogTitle("Save Output As...");
            saveFileChooser.setDialogType(JFileChooser.SAVE_DIALOG);

            saveFileChooser.setSelectedFile(outputCompiledExcel);
            saveFileChooser.setFileFilter(new FileNameExtensionFilter("Excel (*.xlsx)", "xlsx"));

            int option = saveFileChooser.showSaveDialog(APP_FRAME);
            if (option == JFileChooser.APPROVE_OPTION) {
                File selectedFile = saveFileChooser.getSelectedFile();
                if (selectedFile != null) {
                    if (!selectedFile.getName().toLowerCase().endsWith(".xlsx")) {
                        selectedFile = new File(selectedFile.getParentFile(), selectedFile.getName() + ".xlsx");
                    }
                    copy(outputCompiledExcel, selectedFile);
                    Desktop.getDesktop().open(selectedFile);
                    outputCompiledExcel.delete();
                }
            }
        } catch (EncryptedDocumentException ex) {
            UTILITY_MGR.getLogger().log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            UTILITY_MGR.getLogger().log(Level.SEVERE, null, ex);
        }

        jButtonRun.setEnabled(false);
        jButtonRemoveSelectedFiles.setEnabled(true);
        jButtonResetAll.setEnabled(true);
        jButtonSelectInputFiles.setEnabled(true);
    }

    private void inputZipFiles(ArrayList<File> zipFiles) 
            throws EncryptedDocumentException, 
            IOException, 
            FileNotFoundException {

        for (File zipFile : zipFiles) {
            ZipFile zipArchiveFile =  new ZipFile(zipFile);
            readZipFile(zipArchiveFile);
        } // for-loop each ZipFile
    }
    
    private void readZipFile(ZipFile zipFile) throws IOException { // for each zip archive
        String zipPath=zipFile.getName();
        String nameOfSheet=zipPath.substring(zipPath.lastIndexOf('\\')+1, zipPath.lastIndexOf("."));
        UTILITY_MGR.outputConsoleLogsBreakline(("Processing "+nameOfSheet+" Archive"));
        UTILITY_MGR.updateLogs();
        
        if(nameOfSheet.length()>30) {
            nameOfSheet = nameOfSheet.substring(0,30);
        }
        Sheet sheet = workbook.getSheet(nameOfSheet);
        if(sheet ==null) {
            sheet = workbook.createSheet(nameOfSheet);
        } else {
            nameOfSheet=nameOfSheet.substring(0,28) + "2";
        }
        
        sheet.setZoom(100);
        InputStream stream = null;
        InputStreamReader isr = null;
        CSVReader reader = null;
        
        ZipEntry entry = null;
        Enumeration<? extends ZipEntry> entries = zipFile.entries();
        
        while(entries.hasMoreElements()){
            entry = entries.nextElement();
            stream = zipFile.getInputStream(entry);
            isr = new InputStreamReader(stream, StandardCharsets.UTF_8);
            
            reader = new CSVReader(isr, ',', '"');
            String[] nextLine;
            int nrCounter=0;
            
            while ((nextLine= reader.readNext()) != null) {
                Row newRow = sheet.createRow(nrCounter++);
                for(int arrIndex=0;arrIndex<nextLine.length;arrIndex++) {
                    String str=nextLine[arrIndex];
                    if(str==null || str.isEmpty()) {
                        newRow.createCell(arrIndex).setCellValue("");
                    } else {
                        try {
                            Double cellValue = Double.parseDouble(str);
                            Cell newCell = newRow.createCell(arrIndex);
                            newCell.setCellType(CellType.NUMERIC);
                            newCell.setCellValue(cellValue);
                        } catch(NumberFormatException nfe) {
                            newRow.createCell(arrIndex).setCellValue(str);
                        }
                    }
                }
            }
        }
        workbookName = "ZipToExcel_" + getCurrentTimeStamp() + ".xlsx";
        outputCompiledExcel = new File(workbookName);
        FileOutputStream out = new FileOutputStream(outputCompiledExcel);
        workbook.write(out);
        out.close();
    }
}
