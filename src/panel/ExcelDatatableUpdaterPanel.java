package panel;

import au.com.bytecode.opencsv.CSVReader;
import java.awt.Desktop;
import util.UtilityManager;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import javax.swing.filechooser.FileNameExtensionFilter;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.awt.event.ActionEvent;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;
import javax.swing.DefaultListModel;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JList;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTable;
import static util.UtilityManager.copy;
import static util.UtilityManager.getColumnName;
import static util.UtilityManager.getCurrentTimeStamp;

public class ExcelDatatableUpdaterPanel extends JPanel {

    DataFormatter dataFormatter = new DataFormatter();
    private final JFrame APP_FRAME;
    private final UtilityManager UTILITY_MGR;
    private JFileChooser FILE_CHOOSER;
    DefaultListModel jListInputFilesSelectedModel = new DefaultListModel<>();
    private static JList<String> jListInputFilesSelected;
    private static JScrollPane jScrollPane1FileListItems;

    // OUTPUT LOGS UTILITY_MGR
    private final JTextArea LOG_TEXT_AREA;
    private final JScrollPane JSCROLL_PANEL_OUTPUT_LOGS;

    private static JLabel jLabelSelectUpdateFileChooserText;
    private static JLabel jLabelCSVFileChooserText;

    private static JLabel jLabelFileChooserText;
    private static JLabel jLabelOutputFileLogsTitle;
    private static JLabel jLabelFileListSelected;

    private static JButton jButtonSelectUpdateFile;
    private static JButton jButtonSelectInputCSVData;
    private static JButton jButtonAppendData;

    private static JButton jButtonSelectInputFiles;
    private static JButton jButtonResetAll;
    private static JButton jButtonRemoveSelectedFiles;

    private static JButton jButtonRun;

    private static JLabel jLabelUpdateFileName;
    private static JLabel jLabelCSVFileName;

    // LIST OF FILE ITEMS - INPUT FILES TO COMPILE INTO ARCHIVE
    private static final ArrayList<File> INPUT_FILES = new ArrayList<File>();
    private static File outputArchiveZip = null;

    private static File fileToAppend = null;
    private static File dataToAppend = null;

    public ExcelDatatableUpdaterPanel(JFrame APP_FRAME) {
        super();
        this.APP_FRAME = APP_FRAME;

        // INPUT FILES SELECTED
        jLabelSelectUpdateFileChooserText = new JLabel("Select File to Append:");
        jButtonSelectUpdateFile = new JButton("Choose File");
        jLabelUpdateFileName = new JLabel("(No File Chosen)");

        jLabelCSVFileChooserText = new JLabel("Upload Data:");
        jButtonSelectInputCSVData = new JButton("Choose File");
        jLabelCSVFileName = new JLabel("(No File Chosen)");
        jButtonAppendData = new JButton("Append Data >>");

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

        LOG_TEXT_AREA = new JTextArea();
        LOG_TEXT_AREA.setEditable(false);
        LOG_TEXT_AREA.setWrapStyleWord(true);
        JSCROLL_PANEL_OUTPUT_LOGS = new JScrollPane(LOG_TEXT_AREA);
        UTILITY_MGR = new UtilityManager(LOG_TEXT_AREA, JSCROLL_PANEL_OUTPUT_LOGS); // so all logs are handled by the same panel

        UTILITY_MGR.getLogger().info(() -> "Welcome to Excel Datatable Updater.");
        // set components properties
        jButtonRun.setEnabled(false);
        jButtonAppendData.setEnabled(false);
        
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

        // set component bounds (only needed by Absolute Positioning
        jLabelSelectUpdateFileChooserText.setBounds(20, 15, 130, 30);
        jButtonSelectUpdateFile.setBounds(160, 15, 130, 30);
        jLabelUpdateFileName.setBounds(310, 15, 400, 30);

        jLabelCSVFileChooserText.setBounds(20, 50, 130, 30);
        jButtonSelectInputCSVData.setBounds(160, 50, 130, 30);
        jLabelCSVFileName.setBounds(310, 50, 400, 30);

        jButtonAppendData.setBounds(665, 50, 130, 30);

        add(jLabelSelectUpdateFileChooserText);
        add(jButtonSelectUpdateFile);
        add(jLabelUpdateFileName);

        add(jLabelCSVFileChooserText);
        add(jButtonSelectInputCSVData);
        add(jLabelCSVFileName);

        add(jButtonAppendData);

        jLabelFileChooserText.setBounds(20, 90, 795, 30);
        jButtonSelectInputFiles.setBounds(160, 90, 130, 30);

        jButtonRemoveSelectedFiles.setBounds(665, 90, 130, 30);
        jLabelFileListSelected.setBounds(395, 90, 200, 30);
        jScrollPane1FileListItems.setBounds(395, 125, 400, 155);

        jButtonRun.setBounds(20, 215, 130, 30);
        jLabelOutputFileLogsTitle.setBounds(20, 255, 775, 30);
        JSCROLL_PANEL_OUTPUT_LOGS.setBounds(20, 285, 775, 220);
        jButtonResetAll.setBounds(665, 515, 130, 30);

        jButtonSelectUpdateFile.addActionListener((java.awt.event.ActionEvent evt) -> {
            selectUpdateFileAction(evt);
        });

        jButtonSelectInputCSVData.addActionListener((java.awt.event.ActionEvent evt) -> {
            selectInputCSVDataAction(evt);
        });

        jButtonSelectInputFiles.addActionListener((java.awt.event.ActionEvent evt) -> {
            selectInputFilesAction(evt);
        });

        jButtonResetAll.addActionListener((java.awt.event.ActionEvent evt) -> {
            resetAllAction(evt);
        });

        jButtonRun.addActionListener((java.awt.event.ActionEvent evt) -> {
            runAppAction(evt);
        });

        jButtonAppendData.addActionListener((java.awt.event.ActionEvent evt) -> {
            runAppendAction(evt);
        });

        jButtonRemoveSelectedFiles.addActionListener((java.awt.event.ActionEvent evt) -> {
            removeListItemAction(evt);
            if (INPUT_FILES.isEmpty()) {
                jButtonRun.setEnabled(false);
            }
        });
    }

    private void selectInputFilesAction(ActionEvent e) {
        FILE_CHOOSER = new JFileChooser();
        FILE_CHOOSER.setDialogTitle("Select Input File(s)");

        FILE_CHOOSER.setMultiSelectionEnabled(true);
        FILE_CHOOSER.setAcceptAllFileFilterUsed(false);

        FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel File (.xlsx)", "xlsx");
        FILE_CHOOSER.addChoosableFileFilter(filter);

        int option = FILE_CHOOSER.showOpenDialog(APP_FRAME);
        if (option == JFileChooser.APPROVE_OPTION) {
            jListInputFilesSelectedModel = (DefaultListModel) jListInputFilesSelected.getModel();
            File[] selectedFiles = FILE_CHOOSER.getSelectedFiles();
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

    private void selectUpdateFileAction(ActionEvent e) {
        FILE_CHOOSER = new JFileChooser();
        FILE_CHOOSER.setDialogTitle("Select Excel File to Upudate");

        FILE_CHOOSER.setMultiSelectionEnabled(false);
        FILE_CHOOSER.setAcceptAllFileFilterUsed(false);

        FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel File (.xlsx)", "xlsx");
        FILE_CHOOSER.addChoosableFileFilter(filter);

        int option = FILE_CHOOSER.showOpenDialog(APP_FRAME);
        if (option == JFileChooser.APPROVE_OPTION) {
            File selectedFile = FILE_CHOOSER.getSelectedFile();
            fileToAppend = selectedFile.getAbsoluteFile();
            if (fileToAppend != null) {
                jLabelUpdateFileName.setText(fileToAppend.getName());
            }
            if (fileToAppend != null && dataToAppend != null) {
                jButtonAppendData.setEnabled(true);
            }
        }
    }

    private void selectInputCSVDataAction(ActionEvent e) {
        FILE_CHOOSER = new JFileChooser();
        FILE_CHOOSER.setDialogTitle("Select CSV Data File");

        FILE_CHOOSER.setMultiSelectionEnabled(false);
        FILE_CHOOSER.setAcceptAllFileFilterUsed(false);

        FileNameExtensionFilter filter = new FileNameExtensionFilter("CSV File (.csv)", "csv");
        FILE_CHOOSER.addChoosableFileFilter(filter);

        int option = FILE_CHOOSER.showOpenDialog(APP_FRAME);
        if (option == JFileChooser.APPROVE_OPTION) {
            File selectedFile = FILE_CHOOSER.getSelectedFile();
            dataToAppend = selectedFile;
            if (dataToAppend != null) {
                jLabelCSVFileName.setText(dataToAppend.getName());
            }
            if (fileToAppend != null && dataToAppend != null) {
                jButtonAppendData.setEnabled(true);
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
        jButtonAppendData.setEnabled(false);
        jListInputFilesSelectedModel.clear();
        INPUT_FILES.clear();
        LOG_TEXT_AREA.setText("");
        jLabelUpdateFileName.setText("No File Chosen");
        jLabelCSVFileName.setText("No File Chosen");
        UTILITY_MGR.getLogger().info(() -> "Welcome to Excel Datatable Updater.");
    }

    private void runAppendAction(ActionEvent e) {
        UTILITY_MGR.outputConsoleLogsBreakline("");
        UTILITY_MGR.outputConsoleLogsBreakline("Append to Excel Datatable");
        UTILITY_MGR.outputConsoleLogsBreakline("");

        try {
            XSSFWorkbook workbookToUpdate = new XSSFWorkbook(new FileInputStream(fileToAppend));
            XSSFSheet sheetToUpdate = workbookToUpdate.getSheetAt(0);
            int lastRowIndex = sheetToUpdate.getLastRowNum();

            FileInputStream fis = new FileInputStream(dataToAppend);
            InputStreamReader isr = new InputStreamReader(fis, StandardCharsets.UTF_8);
            CSVReader reader = new CSVReader(isr, ',', '"');

            int nrCounter = 0;
            String[] nextLine;
            while ((nextLine = reader.readNext()) != null) {
                XSSFRow newRow = sheetToUpdate.createRow(lastRowIndex + 1 + nrCounter);
                int ncCounter = 0;
                for (String nextLineStr : nextLine) {
                    XSSFCell newCell = newRow.createCell(ncCounter);
                    newCell.setCellValue(nextLineStr);
                    ncCounter++;
                }
                nrCounter++;
            }

            reader.close();
            isr.close();
            
            FileOutputStream out = new FileOutputStream(fileToAppend);
            workbookToUpdate.write(out);
            out.flush();
            out.close();
            workbookToUpdate.close();

            UTILITY_MGR.outputConsoleLogsBreakline("");
            UTILITY_MGR.outputConsoleLogsBreakline("Data is appended.");
            UTILITY_MGR.outputConsoleLogsBreakline("");
        } catch (EncryptedDocumentException | IOException ex) {
            UTILITY_MGR.getLogger().log(Level.SEVERE, null, ex);
        }
    }

    private void runAppAction(ActionEvent e) {
        jButtonSelectInputFiles.setEnabled(false);
        jButtonResetAll.setEnabled(false);
        jButtonRemoveSelectedFiles.setEnabled(false);

        UTILITY_MGR.outputConsoleLogsBreakline("");
        UTILITY_MGR.outputConsoleLogsBreakline("Initialising Excel Datatable Updator");
        UTILITY_MGR.outputConsoleLogsBreakline("");

        try {
            UTILITY_MGR.outputConsoleLogsBreakline("Reading in excel files");
            UTILITY_MGR.updateLogs();
            // ================================================= READ IN FILES ================================
            inputExcel(INPUT_FILES);
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
        } catch (EncryptedDocumentException | IOException ex) {
            UTILITY_MGR.getLogger().log(Level.SEVERE, null, ex);
        }

        jButtonRun.setEnabled(false);
        jButtonRemoveSelectedFiles.setEnabled(true);
        jButtonResetAll.setEnabled(true);
        jButtonSelectInputFiles.setEnabled(true);
    }

    private void inputExcel(ArrayList<File> excelFiles) throws EncryptedDocumentException, IOException, FileNotFoundException {
        Workbook workbook = null;

        outputArchiveZip = new File("ExcelDatatableUpdated_" + getCurrentTimeStamp() + ".zip");
        try (FileOutputStream fos = new FileOutputStream(outputArchiveZip)) {
            ZipOutputStream zipOut = new ZipOutputStream(fos);
            File outputFile = null;

            for (File excelFile : excelFiles) {
                workbook = new XSSFWorkbook(new FileInputStream(excelFile));
                outputFile = extendXSSFDatatableRange(workbook, excelFile);

                UTILITY_MGR.outputConsoleLogsBreakline(excelFile.getName() + " datatables have been updated.");
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

    private File extendXSSFDatatableRange(Workbook workbook, File excelFile) throws IOException {
        String excelFileName = excelFile.getName();
        excelFileName = excelFileName.substring(0, excelFileName.indexOf(".xls"));
        int noOfSheets = workbook.getNumberOfSheets();

        for (int s = 0; s < noOfSheets; s++) {
            XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(s);
            workbook.setSheetHidden(s, false);
            List<XSSFTable> tables = sheet.getTables();

            int tableFirstRowNum = 0;
            int tableLastRowNum = 0;

            int firstCellNum = 0;
            int lastCellNum = 0;

            String firstCellRef = "";
            String lastCellRef = "";

            int sheetLastRowNum = sheet.getLastRowNum() + 1;

            for (XSSFTable t : tables) {
                XSSFTable currentTable = t;
                CTTable ctTable = currentTable.getCTTable();
                String originalTableRef = ctTable.getRef();

                tableFirstRowNum = t.getStartRowIndex() + 1;
                tableLastRowNum = t.getEndRowIndex() + 1;

                firstCellNum = t.getStartColIndex();
                lastCellNum = t.getEndColIndex();

                firstCellRef = getColumnName(firstCellNum);
                lastCellRef = getColumnName(lastCellNum);

                System.out.println("Current Datatable Range:");
                System.out.println("[R" + tableFirstRowNum + "]..[R" + tableLastRowNum + "]");
                System.out.println("[C" + firstCellNum + "]..[C" + lastCellNum + "]");
                System.out.println("[" + firstCellRef + "]..[" + lastCellRef + "]");

                String newRefStr = firstCellRef + tableFirstRowNum + ":" + lastCellRef + sheetLastRowNum;
                ctTable.setRef(newRefStr);
                currentTable.updateReferences();

                System.out.println("Table Range Reference has been updated from: " + originalTableRef + " to " + newRefStr);
                UTILITY_MGR.outputConsoleLogsBreakline("Table Range Reference has been updated from: " + originalTableRef + " to " + newRefStr);
                UTILITY_MGR.updateLogs();
            }
        }
        try {
            XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
        } catch (Exception e) {
            e.printStackTrace();
        }
        String outputExcelFileName = excelFileName + ".xlsx";
        File outputFile = new File(outputExcelFileName);
        FileOutputStream os = new FileOutputStream(outputFile);

        workbook.write(os);
        workbook.close();

        os.flush();
        os.close();

        return outputFile;
    }
}
