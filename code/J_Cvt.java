import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.openxml4j.opc.OPCPackage;

import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

public class J_Cvt {
    private static FileWriter fout = null; // for output use
    private static String iFilePath = "";  // input excel file name
    private static String oFilePath = "";  // output txt file name
    private static int sheetNo = 0;        // the sheet number of input excel file to be processed
    private static int[] outputCol;        // excel file output column
    private static int outputColNum = 0;   // the number of output column
    private static int skipRow = 0;        // skip the n row of sheet in excel
    private static String prefixFlag = ""; // prefix flag
    private static String extension = "";  // input excel file extension

    private static int maxOutputColumn = 0;// the max output column number
    private static String[] rowDataSet;    // data of one row

    enum XSSFDataType {
        BOOL, ERROR, FORMULA, INLINESTR, SSTINDEX, NUMBER,
    }

    public static void main(String[] args) throws Exception {
        /* check input parameter */

        if (args == null || args.length < 5) {
            throw new Exception("Usage: J_Cvt [input file] [output file] [excel file ouput column] [sheet no.] [skip row] [prefix string]");
        }

        /* pre-process the parameter */
        
        // input file path
        iFilePath = args[0];
        int ext = iFilePath.lastIndexOf(".");
        if (ext == -1) {
            extension = "";
        } else {
            extension = iFilePath.substring(ext + 1).toUpperCase();
        }
        // output file path
        oFilePath = args[1];
        // output column order
        String[] outputColA = args[2].split(",");
        outputColNum = outputColA.length;
        outputCol = new int[outputColNum];
        for (int i = 0; i < outputColNum; i++) {
            outputCol[i] = convertCol(outputColA[i].toUpperCase());
            if (outputCol[i] > maxOutputColumn) maxOutputColumn = outputCol[i];
        }
        // input file sheet number
        sheetNo = Integer.parseInt(args[3]);
        // input file skip row
        skipRow = Integer.parseInt(args[4]);
        // prefix flag
        if (args.length >= 6) prefixFlag = args[5];

        System.out.println("Parse the file [" + iFilePath + "] with sheet [" + args[3] + "] of column [" + args[2] +
                           "] and skip [" + args[4] + "] row, you will get result in file [" + oFilePath + "]");

        /* initial parameter */

        fout = new FileWriter(oFilePath);
        rowDataSet = new String[maxOutputColumn + 1];

        /* process the excel */

        J_Cvt tool = new J_Cvt();
        try {
            if ("XLS".equals(extension)) {
                tool.processXlsFile(iFilePath, sheetNo);
            } else if ("XLSX".equals(extension)) {
                tool.processXlsxFile(iFilePath, sheetNo);
            } else {
                throw new Exception("Input file format wrong!");
            }
        } catch (GetContentEndException e) {
            // Already get all content
        }
        fout.close();

        return;
    }

    /**
     * Get index of column (A-0, Z-25, AA-26, AZ-51, ZZ-701)
     */
    private static int convertCol(String col) throws Exception {
        if (col.length() == 1) {
            return col.charAt(0) - 'A';
        } else if (col.length() == 2) {
            return (col.charAt(0) - 'A' + 1) * 26 + (col.charAt(1) - 'A');
        }
        return 0;
    }

   /**
    * Write one row data in order of ouput columns to output file
    */
    private static void writeOneRowToFile(int currentRow) throws Exception {
        // prefix flag
        if (!"".equals(prefixFlag)) fout.write(prefixFlag + "|");
        // content
        for (int i = 0; i < outputColNum; i++) {
            fout.write(rowDataSet[outputCol[i]] + "|");
            rowDataSet[outputCol[i]] = "";
        }
        fout.write("\n");
    }

    /** 
     * Read the sheet of excel xls file
     */
    public void processXlsFile(String filePath, int sheetNo) throws Exception {
        InputStream readFile = new FileInputStream(filePath);
        HSSFWorkbook wb = new HSSFWorkbook(readFile);
        HSSFSheet sheet = wb.getSheetAt(sheetNo - 1);
        HSSFRow row;
        HSSFCell cell;
        Iterator<org.apache.poi.ss.usermodel.Row> rows = sheet.rowIterator();

        while (rows.hasNext()) {
            row = (HSSFRow)rows.next();
            if (row.getRowNum() + 1 <= skipRow) continue;
            int s = row.getFirstCellNum();
            int t = row.getLastCellNum();
            String lastContents = "";
            for (int i = 0; i < outputColNum; i++) {
                cell = row.getCell(outputCol[i]);
                if (cell == null) {
                    lastContents = "";
                } else {
                    // get cell style
                    HSSFCellStyle style = cell.getCellStyle();
                    int formatIndex = style.getDataFormat();
                    String formatString = style.getDataFormatString();
                    if (formatString == null) {
                        formatString = BuiltinFormats.getBuiltinFormat(formatIndex);
                    }
                    // get content
                    switch (cell.getCellTypeEnum()) {
                    case BOOLEAN:
                        lastContents = cell.getBooleanCellValue() ? "TRUE" : "FALSE";
                        break;
                    case STRING:
                        lastContents = cell.getStringCellValue();
                        break;
                    case NUMERIC:
                        if (HSSFDateUtil.isCellDateFormatted(cell)) { // date
                            Date date = HSSFDateUtil.getJavaDate(cell.getNumericCellValue());
                            lastContents = formateDateToString(date);
                        } else if (formatString != null) {
                            DataFormatter formatter = new DataFormatter();
                            lastContents = formatter.formatRawCellContents(Double.parseDouble(cell.toString()),
                                                                           formatIndex,
                                                                           formatString);
                        } else {
                            lastContents = cell.toString();
                        }
                        break;
                    case FORMULA:
                        lastContents = cell.getCellFormula();
                        break;
                    default:
                        lastContents = cell.toString();
                        break;
                    }
                }
                // trim the output string
                if (lastContents == null) {
                    lastContents = "";
                } else {
                    lastContents = lastContents.trim();
                }
                rowDataSet[outputCol[i]] = lastContents; // store the needed column data
            }

            // write one row data
            if (rowDataSet[outputCol[0]] == null || "".equals(rowDataSet[outputCol[0]])) { // reach footprint
                throw new GetContentEndException();
            }
            writeOneRowToFile(row.getRowNum());
        }
    }

    /** 
     * Read the sheet of excel xlsx file
     */
    public void processXlsxFile(String filePath, int sheetNo) throws Exception {
        OPCPackage pkg = OPCPackage.open(filePath);
        XSSFReader r = new XSSFReader(pkg);
        StylesTable styles = r.getStylesTable();
        SharedStringsTable sst = r.getSharedStringsTable();

        XMLReader parser = fetchSheetParser(sst, styles);

        // To look up the Sheet Name / Sheet Order / rID,
        //   you need to process the core Workbook stream.
        // Normally it's of the form rId# or rSheet#
        InputStream sheet2 = r.getSheet("rId" + sheetNo);
        InputSource sheetSource = new InputSource(sheet2);
        parser.parse(sheetSource);
        sheet2.close();
    }
        
    private static String formateDateToString(Date date) {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        return sdf.format(date);
    }

    public XMLReader fetchSheetParser(SharedStringsTable sst, StylesTable styles) throws SAXException {
        XMLReader parser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
        ContentHandler handler = new SheetHandler(sst, styles);
        parser.setContentHandler(handler);
        return parser;
    }

    /**
     * After get all content needed, will throw this exception
     */
    private static class GetContentEndException extends SAXException {
        private static final long serialVersionUID = 2016122300L;
        private GetContentEndException() {
            super();
        }
    }

    /**
     * An individual SAX handler that parse the xlsx file
     */
    private static class SheetHandler extends DefaultHandler {
        private SharedStringsTable sst;
        private XSSFDataType nextDataType; // the data type of the cell
        private StylesTable stylesTable;   // the cell format table of the sheet
        private StringBuffer value;        // string buffer of the current reading cell
        private String lastContents = "";  // the content of the cell
        private int currentCol = 0;        // current row
        private int currentRow = 0;        // current column [0 for A, 1 for B, ..., 25 for Z, 26 for AA]
        private boolean vIsOpen = false;   // set when V start element is seen

        private final String REGEX = "[A-Z]*"; // for handling the reference of the cell to seek the column [A1 -> A]

        // Used to format numeric cell values
        private short formatIndex;
        private String formatString;
        private final DataFormatter formatter;

        private SheetHandler(SharedStringsTable sst, StylesTable styles) {
            this.sst = sst;
            this.stylesTable = styles;
            this.nextDataType = XSSFDataType.NUMBER;
            this.value = new StringBuffer();
            this.formatter = new DataFormatter();
        }

        // override
        public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
            if ("inlineStr".equals(name) || "v".equals(name)) {
                vIsOpen = true;
                value.setLength(0);
            } else if (name.equals("c")) {
                // c => cell
                // Print the cell reference
                try {
                    // get column index
                    String ref = attributes.getValue("r");
                    Pattern p = Pattern.compile(REGEX);
                    Matcher m = p.matcher(ref);
                    if (m.find()) {
                        String tempCol = ref.substring(0, m.end());
                        String tempRow = ref.substring(m.end(), ref.length());
                        currentCol = convertCol(tempCol);
                        currentRow = Integer.parseInt(tempRow);
                    }
                    if (currentCol <= maxOutputColumn) rowDataSet[currentCol] = ""; // initialize the needed data

                    // set default
                    nextDataType = XSSFDataType.NUMBER;
                    formatIndex = -1;
                    formatString = null;
                    String cellType = attributes.getValue("t");
                    String cellDataType = attributes.getValue("s");

                    // Figure out if the value is an index in the SST
                    if ("s".equals(cellType)) {
                        nextDataType = XSSFDataType.SSTINDEX;
                    } else if ("b".equals(cellType)) {
                        nextDataType = XSSFDataType.BOOL;
                    } else if ("e".equals(cellType)) {
                        nextDataType = XSSFDataType.ERROR;
                    } else if ("inlineStr".equals(cellType)) {
                        nextDataType = XSSFDataType.INLINESTR;
                    } else if ("str".equals(cellType)) {
                        nextDataType = XSSFDataType.FORMULA;
                    } else if (cellDataType != null) {
                        // It's a number but almost certainly one with a special style or format
                        int styleIndex = Integer.parseInt(cellDataType);
                        XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);
                        formatIndex = style.getDataFormat();
                        formatString = style.getDataFormatString();
                        if (formatString == null) {
                            formatString = BuiltinFormats.getBuiltinFormat(formatIndex);
                        }
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                    return;
                }
            }
            // Clear contents cache
            lastContents = "";
        }

        // override
        public void endElement(String uri, String localName, String name) throws SAXException {
            // Process the last contents as required
            // Do now, as characters() may be called more than once
            try {
                if (name.equals("v")) {
                    // v => contents of a cell
                    // Output after we've seen the string contents
                    switch (nextDataType) {
                    case BOOL:
                        char first = value.charAt(0);
                        lastContents = first == '0' ? "FALSE" : "TRUE";
                        break;
                    case SSTINDEX:
                        int idx = Integer.parseInt(value.toString());
                        lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
                        break;
                    case INLINESTR:
                        lastContents = new XSSFRichTextString(value.toString()).toString();
                        break;
                    case NUMBER:
                        String n = value.toString();
                        if (HSSFDateUtil.isADateFormat(formatIndex, n)) { // date format
                            Double d = Double.parseDouble(n);
                            Date date = HSSFDateUtil.getJavaDate(d);
                            lastContents = formateDateToString(date);
                        } else if (formatString != null) {
                            lastContents = formatter.formatRawCellContents(Double.parseDouble(n), formatIndex, formatString);
                        } else {
                            lastContents = n;
                        }
                        break;
                    default:
                        lastContents = value.toString();
                        break;
                    }

                    // trim the output string
                    if (lastContents == null) {
                        lastContents = "";
                    } else {
                        lastContents = lastContents.trim();
                    }
                    if (currentCol <= maxOutputColumn) rowDataSet[currentCol] = lastContents; // store the needed column data
                } else if (name.equals("row")) { // end of a line
                    if (currentRow > skipRow) {
                        if (rowDataSet[outputCol[0]] == null || "".equals(rowDataSet[outputCol[0]])) { // reach footprint
                            throw new GetContentEndException();
                        }
                        writeOneRowToFile(currentRow);
                    }
                }
            } catch (GetContentEndException e) {
                throw e;
            } catch (Exception e) {
                e.printStackTrace();
                return;
            }
        }

        // override
        public void characters(char[] ch, int start, int length) throws SAXException {
            if (vIsOpen) value.append(ch, start, length);
        }
    } // private static class SheetHandler extends DefaultHandler
}

