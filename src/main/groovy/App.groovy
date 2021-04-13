import java.io.File;
import org.apache.poi.poi.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import java.nio.file.Files;
import groovy.io.FileType;
import groovy.swing.SwingBuilder;
import javax.swing.*;
import java.awt.*;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import java.text.SimpleDateFormat
// Obs fjerne denne før compile 
//@Grab(group='org.apache.poi', module='poi-ooxml', version='4.0.0')
//@Grab(group='org.apache.poi', module='poi', version='4.0.0')
// Previous imports: @Grab(group='org.apache.poi', module='poi', version='4.0.0')
//@Grab(group='org.apache.poi', module='poi-ooxml', version='4.0.0')




// Read in the file:
def readFile = {f -> 
    InputStream inp = new FileInputStream( new File(f) )
    Workbook wb = WorkbookFactory.create( inp );                                                                     
    return wb
}
// Get the sheet:
def getSheet = {wb, sheetName -> 
    Sheet sheet = wb.getSheet(sheetName);
    return sheet
    }


// Get the contents as a list:
def parseXLtoList = {sheet, count, sessionID ->
    def output = []
    def tempstr = []
    while (sheet.getRow(count) != null) {
        sId = sheet.getRow(count).getCell(3).toString()
        if (sId.contains(',')) {
        sample = sId.tokenize(',')[0]
        year = "20" + sId.tokenize(',')[1].tokenize(' ')[0]
        extrID = sId.tokenize(' ')[1]
        userID = sessionID.tokenize('_')[-1]
        
        def sdf = new SimpleDateFormat("dd.MM.yyyy")
        def date = sdf.format(new Date())
        
       
        tempstr = [sheet.getRow(count).getCell(1).toString(), sample, year, extrID,  sheet.getRow(count).getCell(12).toString().replace('.', ','), sheet.getRow(count).getCell(16).toString().replace('.', ','), userID, date];
        output.add(tempstr)
        }
    count += 1;
    }
    return output
}

//hvis 280/260 under 1.6 

// Results:
def addToResultsFiles = {result_summary, sessionID ->
    // Write to the files:
    FileOutputStream fout = new FileOutputStream(sessionID + ".xls");  //Open FileOutputStream to write updates          
    // Build the Excel File
    ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
    HSSFWorkbook workBook = new HSSFWorkbook();
    // Create the spreadsheet
    HSSFSheet spreadSheet = workBook.createSheet("Hello_World");
    // Create the cells and write to the file
    HSSFCell cell;
    //Sett inn i excelark:
    parseXLtoList(result_summary, 5, sessionID).eachWithIndex {it, index ->
        if (it[2].toString() != "Ctrl DNA" && !it[2].toString().startsWith("Blank") ) {
            HSSFRow row = spreadSheet.createRow((short) 0 + index);
            it.eachWithIndex {itt, index2 ->
                cell = row.createCell(index2);
                cell.setCellValue(new HSSFRichTextString(itt.toString()));
            }
        }  
    }

    workBook.write(outputStream);
    outputStream.writeTo(fout);
    outputStream.close();
    fout.close();
}


// GUI helpers:
def OpenReport = { text ->
    // Sets initial path to project dir
    def initialPath = System.getProperty("user.dir");
    JFileChooser fc = new JFileChooser(initialPath);
    fc.setFileSelectionMode(JFileChooser.FILES_ONLY);
    fc.setDialogTitle(text);
    int result = fc.showOpenDialog( null );
    switch ( result ) {
        case JFileChooser.APPROVE_OPTION:
            File file = fc.getSelectedFile();
            def path =  fc.getCurrentDirectory()
            return [file, path]
        break;
        case JFileChooser.CANCEL_OPTION:
        case JFileChooser.ERROR_OPTION:
            break;
    }
}


// GUI:
def myapp = new SwingBuilder()

def process = {
    
    // Hente rapport-fil:
    def report = OpenReport.call("Velg rapport")
    
    def fname = report[0].toString().tokenize('\\')[-1]
    def path = report[1].toString()
    wb = readFile(path + '\\' + fname)
    // Get the sessionID from the filename:
    def sessionID = fname.tokenize('.')[0]
    result_summary = getSheet(wb, "Result summary")
    addToResultsFiles(result_summary, sessionID)
    JOptionPane.showMessageDialog(null, "Ferdig!!");
    
    
}


def buttonPanel = {
    myapp.panel(constraints : BorderLayout.SOUTH) {
        button(text : 'Åpne fil', actionPerformed : process ) // Endret fra OpenReportFromLC midlertidig
   } 
} 


def mainPanel = {
   myapp.panel(layout : new BorderLayout()) {
      label(text : 'Åpne en rapport fra SkanIT', horizontalAlignment : JLabel.CENTER, constraints : BorderLayout.CENTER)
      buttonPanel()   
   }
} 



def myframe = myapp.frame(title : 'Plateleser rapportverktøy v0.1', location : [100, 100],
   size : [400, 300], defaultCloseOperation : WindowConstants.EXIT_ON_CLOSE) {
      mainPanel()
     
   } 

myframe.setVisible(true)


/*  
    Dette er fra da alt skulle inn i en samlefil. Nå skrives i stedet alt til en fil etter hver kjøring.
    
    else { // Write the control and blank to the kontrol_blank-file
            row = c_sheet.createRow(c_lastnum);
            it.eachWithIndex {itt, index2 ->
                c_cell = row.createCell(index2);
                c_cell.setCellValue(itt.toString());
            }
            c_lastnum += 1;
        }  



        // Finn første rad som er tom
    def lastnum = sheet.getLastRowNum() + 1
    def c_lastnum = c_sheet.getLastRowNum() + 1
    
    
    FileInputStream inputStream = new FileInputStream(new File("resultater_plateleser.xlsx"));
    Workbook workbook = WorkbookFactory.create(inputStream);
    Sheet sheet = workbook.getSheetAt(0);
    Cell cell = null; // declare a Cell object
    // Control and blank:
    FileInputStream c_inputStream = new FileInputStream(new File("kontroll_blank_log.xlsx"));
    Workbook c_workbook = WorkbookFactory.create(c_inputStream);
    Sheet c_sheet = c_workbook.getSheetAt(0);
    Cell c_cell = null; // declare a Cell object


    // CLose the streams:
    inputStream.close();
    c_inputStream.close();
*/ 