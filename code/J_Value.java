import java.io.File;
import java.io.FileInputStream;
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

public class J_Value {
    private static String iFilePath = ""; // input excel file name
    private static int sheetNo = 0;       // the sheet number of input excel file to be processed

    private static int outputCol = 0;     // output column (A-0,B-1,...,Z-25,AA-26,...)
    private static int outputRow = 0;     // output row    (from 1 to n)
    private static String cellData = "";  // data of cell
    private static String extension = ""; // input excel file extension

    enum XSSFDataType {
        BOOL, ERROR, FORMULA, INLINESTR, SSTINDEX, NUMBER,
    }

    private final static String REGEX = "[A-Z]*"; // for handling the reference of the cell to seek the column [A1 -> A]

    public static void main(String[] args) throws Exception {
        if (args == null || args.length < 3) {
            throw new Exception("Usage: J_Value [input file] [excel file ouput column] [sheet no.]");
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
        // output column
        int[] outputColA = getCellPosition(args[1].toUpperCase());
        outputCol = outputColA[0];
        outputRow = outputColA[1];
        // input file sheet number
        sheetNo = Integer.parseInt(args[2]);

        /* process the excel */

        J_Value tool = new J_Value();
        try {
            if ("XLS".equals(extension)) {
                tool.processXlsFile(iFilePath, sheetNo);
            } else if ("XLSX".equals(extension)) {
                tool.processXlsxFile(iFilePath, sheetNo);
            } else {
                throw new Exception("Input file format wrong!");
            }
        } catch (GetCellDataException e) {
            // Already get cell content
            System.out.print(cellData);
        }

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

    private static int[] getCellPosition(String ref) throws Exception {
        int[] res = new int[2];
        Pattern p = Pattern.compile(REGEX);
        Matcher m = p.matcher(ref);
        if (m.find()) {
            String tempCol = ref.substring(0, m.end());
            String tempRow = ref.substring(m.end(), ref.length());
            res[0] = convertCol(tempCol);
            res[1] = Integer.parseInt(tempRow);
        } else {
            res[0] = -1;
            res[1] = -1;
        }
        return res;
    }
        
    private static String formateDateToString(Date date) {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        return sdf.format(date);
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
            if (row.getRowNum() + 1 != outputRow) continue;
            // get data
            cell = row.getCell(outputCol);
            if (cell == null) {
                cellData = "";
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
                    cellData = cell.getBooleanCellValue() ? "TRUE" : "FALSE";
                    break;
                case STRING:
                    cellData = cell.getStringCellValue();
                    break;
                case NUMERIC:
                    if (HSSFDateUtil.isCellDateFormatted(cell)) { // date
                        Date date = HSSFDateUtil.getJavaDate(cell.getNumericCellValue());
                        cellData = formateDateToString(date);
                    } else if (formatString != null) {
                        DataFormatter formatter = new DataFormatter();
                        cellData = formatter.formatRawCellContents(Double.parseDouble(cell.toString()),
                                                                   formatIndex,
                                                                   formatString);
                    } else {
                        cellData = cell.toString();
                    }
                    break;
                case FORMULA:
                    cellData = cell.getCellFormula();
                    break;
                default:
                    cellData = cell.toString();
                    break;
                }
            }
            cellData = cellData.trim();
            throw new GetCellDataException();
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

    public XMLReader fetchSheetParser(SharedStringsTable sst, StylesTable styles) throws SAXException {
        XMLReader parser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
        ContentHandler handler = new SheetHandler(sst, styles);
        parser.setContentHandler(handler);
        return parser;
    }

    private static class GetCellDataException extends SAXException {
        private static final long serialVersionUID = 2016122301L;
        private GetCellDataException() {
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
                    String ref = attributes.getValue("r");
                    int[] refnum = getCellPosition(ref);
                    currentCol = refnum[0];
                    currentRow = refnum[1];

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
            if (!(currentCol == outputCol && currentRow == outputRow)) { return; }
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

                    // store the result and exit parsing
                    cellData = lastContents;
                    throw new GetCellDataException();
                }
            } catch (GetCellDataException e) {
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
    }
}

