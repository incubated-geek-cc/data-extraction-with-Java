package gui;

import panel.ExcelToCSVPanel;
import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.event.KeyEvent;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTabbedPane;
import panel.ExcelDatatableUpdaterPanel;
import panel.MergePdfFilesPanel;
import panel.OutlookEmailToCSVPanel;
import panel.ZipToExcelPanel;

public class MainFrame {

    private static final JFrame APP_FRAME = new JFrame("Data & Document Utilities :: Prototype v1.0");
    private static final JTabbedPane TABBED_PANE = new JTabbedPane();
    private static final int[] KEY_EVENTS = {
        KeyEvent.VK_1,
        KeyEvent.VK_2,
        KeyEvent.VK_3,
        KeyEvent.VK_4,
        KeyEvent.VK_5,
        KeyEvent.VK_6,
        KeyEvent.VK_7,
        KeyEvent.VK_8,
        KeyEvent.VK_9
    };
    
    public static void main(String[] args) {
      initUI();
   }

   private static void initUI() {    
      APP_FRAME.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
      APP_FRAME.getContentPane().add(TABBED_PANE, BorderLayout.CENTER);
      createUI();
      APP_FRAME.setSize(835, 625);      
      APP_FRAME.setLocationRelativeTo(null);  
      APP_FRAME.setVisible(true);
   }

   private static void createUI() {
      ExcelToCSVPanel panel_0 = new ExcelToCSVPanel(APP_FRAME);
      addTabPanel(panel_0, "Excel to CSV", "Convert Excel File(s) to CSV", 0);
      
      OutlookEmailToCSVPanel panel_1 = new OutlookEmailToCSVPanel(APP_FRAME);
      addTabPanel(panel_1, "Outlook to CSV", "Extract Outlook Table Content to CSV", 1);
      
      ZipToExcelPanel panel_2 = new ZipToExcelPanel(APP_FRAME);
      addTabPanel(panel_2, "Extract Zip Data to Excel", "Extract Archived CSV(s) to Excel", 2);
      
      ExcelDatatableUpdaterPanel panel_3 = new ExcelDatatableUpdaterPanel(APP_FRAME);
      addTabPanel(panel_3, "Update Datables in Excel", "Update Datatable Ranges in Excel", 2);
      
      MergePdfFilesPanel panel_4 = new MergePdfFilesPanel(APP_FRAME);
      addTabPanel(panel_4, "Merge Pdf Docs", "Merge multiple Pdf File(s) into one", 3);
   }
   
   private static void addTabPanel(JPanel panel, String tabLabel, String tabTooltip, int tabIndex) {
      JLabel filler = new JLabel(tabLabel);
      filler.setHorizontalAlignment(JLabel.CENTER);
      panel.setPreferredSize(new Dimension(835, 625));
      panel.setLayout(null);
      panel.add(filler);
      TABBED_PANE.addTab(tabLabel, null, panel, tabTooltip);
      TABBED_PANE.setMnemonicAt(tabIndex, KEY_EVENTS[tabIndex]);
   }
}
