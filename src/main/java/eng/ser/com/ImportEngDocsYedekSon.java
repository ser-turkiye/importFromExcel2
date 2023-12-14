//
// Source code recreated from a .class file by IntelliJ IDEA
// (powered by FernFlower decompiler)
//

package eng.ser.com;

import com.ser.blueline.*;
import com.ser.blueline.metaDataComponents.*;
import de.ser.doxis4.agentserver.UnifiedAgent;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

public class ImportEngDocsYedekSon extends UnifiedAgent {
    Logger log = LogManager.getLogger(this.getClass().getName());
    String nameDescriptor1 = "ccmPrjDocNumber";
    String nameDescriptorRev = "ccmPrjDocRevision";
    String searchClassName = "Search Engineering Documents";
    IDescriptor descriptor1;
    IDescriptor descriptor2;

    public ImportEngDocsYedekSon() {
    }

    protected Object execute() {
        this.log.info("Initiate the agent");
        if (this.getEventDocument() == null) {
            return this.resultError("Null Document object");
        } else {
            ISession ses = this.getSes();
            IDocument ldoc = this.getEventDocument();
            HashMap<Integer, String> flds = new HashMap();
            flds.put(1, "ccmPrjDocRevision");
            flds.put(2, "ObjectName");
            flds.put(3, "ccmPrjDocDiscipline");
            flds.put(4, "ccmPrjDocDocType");
            flds.put(5, "ccmPrjDocParentDoc");
            flds.put(6, "ccmPrjDocParentDocRevision");
            flds.put(7, "ccmPrjDocCode");
            flds.put(8, "ccmPrjDocClient");
            flds.put(9, "ccmPrjDocClientPrjNumber");
            flds.put(10, "ccmPRJCard_name");
            flds.put(11, "ccmPrjDocVendor");
            flds.put(12, "ccmPrjDocApprCode");
            flds.put(13, "ccmPrjDocIssueStatus");
            flds.put(14, "ccmPrjDocWBS2");
            flds.put(15, "ccmPrjDocDate");
            flds.put(16, "ccmPrjDocReqDate");
            flds.put(17, "ccmPrjDocDueDate");
            flds.put(18, "ccmPrjDocTransIncCode");
            flds.put(19, "ccmPrjDocTransOutCode");

            try {
                String excelPath = this.exportDocumentContent(ldoc, "C:/tmp2/bulk/import");
                FileInputStream fist = new FileInputStream(excelPath);
                Workbook wrkb = new XSSFWorkbook(fist);
                HashMap<String, Row> list = listOfDocuments(wrkb);
                fist.close();
                Iterator var8 = list.entrySet().iterator();

                while(var8.hasNext()) {
                    Map.Entry<String, Row> line = (Map.Entry)var8.next();
                    String docNo = (String)line.getKey();
                    Row row = (Row)line.getValue();
                    String docRev = row.getCell(1).getStringCellValue();
                    IDocument engDocument = this.getEngDocument(ses, docNo, docRev);
                    if (engDocument != null) {
                        this.log.info("*** ENGINERING DOC FOUND :" + docNo);
                        this.log.info("Doc no is " + docNo + " Values is " + row.toString());
                        Iterator var13 = flds.entrySet().iterator();

                        while(var13.hasNext()) {
                            Map.Entry<Integer, String> ffld = (Map.Entry)var13.next();
                            if (row.getCell((Integer)ffld.getKey()) == null) {
                                row.createCell((Integer)ffld.getKey());
                            }
                            String descValue = "";
                            Date descDateValue = new Date();
                            Logger var10000 = this.log;
                            int rowKey = (Integer)ffld.getKey();
                            String descName = (String)ffld.getValue();
                            String descDateStrVal = "";
                            if(descName.contains("Date")){
                                //DateFormat dt = new SimpleDateFormat("yyyyMMdd");
                                //descDateValue = row.getCell(rowKey).getDateCellValue();
                                //descValue = dt.format(descDateValue);
                                descValue = "20230604";
                            }else{
                                descValue = row.getCell(rowKey).getStringCellValue();
                            }
                            var10000.info("DESC::" + descName + " // VALUE: " + descValue);
                            //OLD:var10000.info("DESC::" + var10001 + " // VALUE: " + row.getCell((Integer)ffld.getKey()));
                            //OLD:engDocument.setDescriptorValue((String)ffld.getValue(), row.getCell((Integer)ffld.getKey()).toString());

                            engDocument.setDescriptorValue(descName, descValue);
                            //engDocument.setDescriptorValueTyped(descName,descValue);
                        }

                        engDocument.commit();
                    }
                }

                this.log.info("Finished");
                return this.resultSuccess("Ended successfully");
            } catch (IOException var15) {
                throw new RuntimeException(var15);
            }
        }
    }

    public static HashMap<String, Row> listOfDocuments(Workbook workbook) throws IOException {
        HashMap<String, Row> rtrn = new HashMap();
        Sheet sheet = workbook.getSheetAt(0);
        Iterator var3 = sheet.iterator();

        while(var3.hasNext()) {
            Row row = (Row)var3.next();
            if (row.getRowNum() != 0) {
                Cell cll1 = row.getCell(0);
                if (cll1 != null) {
                    String indx = cll1.getRichStringCellValue().getString();
                    if (!indx.equals("") && !indx.equals("Document Number")) {
                        rtrn.put(indx, row);
                    }
                }
            }
        }

        return rtrn;
    }

    public String exportDocumentContent(IDocument document, String exportPath) throws IOException {
        String expt = "";
        String documentID = document.getDocumentID().getID();
        documentID = documentID.replaceAll(":", ".");

        for(int representationConter = 0; representationConter < document.getRepresentationCount(); ++representationConter) {
            for(int partDocumentCounter = 0; partDocumentCounter < document.getPartDocumentCount(representationConter); ++partDocumentCounter) {
                IDocumentPart partDocument = document.getPartDocument(representationConter, partDocumentCounter);
                InputStream inputStream = partDocument.getRawDataAsStream();

                try {
                    IFDE fde = partDocument.getFDE();
                    if (fde.getFDEType() == 3) {
                        expt = exportPath + "/output_" + documentID + "." + ((IFileFDE)fde).getShortFormatDescription();
                        FileOutputStream fileOutputStream = new FileOutputStream(expt);

                        try {
                            byte[] bytes = new byte[2048];

                            int length;
                            while((length = inputStream.read(bytes)) > -1) {
                                fileOutputStream.write(bytes, 0, length);
                            }
                        } catch (Throwable var15) {
                            try {
                                fileOutputStream.close();
                            } catch (Throwable var14) {
                                var15.addSuppressed(var14);
                            }

                            throw var15;
                        }

                        fileOutputStream.close();
                    }
                } catch (Throwable var16) {
                    if (inputStream != null) {
                        try {
                            inputStream.close();
                        } catch (Throwable var13) {
                            var16.addSuppressed(var13);
                        }
                    }

                    throw var16;
                }

                if (inputStream != null) {
                    inputStream.close();
                }
            }
        }

        return expt;
    }

    public static HashMap<String, List<Object>> readExcelFile(String fileLocation) throws IOException {
        HashMap<String, List<Object>> hm = new HashMap();
        FileInputStream file = new FileInputStream(fileLocation);
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);
        Iterator var5 = sheet.iterator();

        while(var5.hasNext()) {
            Row row = (Row)var5.next();
            int colNo = 1;
            String docNo = "";
            List<Object> list = new ArrayList();

            for(Iterator var10 = row.iterator(); var10.hasNext(); ++colNo) {
                Cell cell = (Cell)var10.next();
                Object cellVal = getValue(cell, cell.getCellType());
                if (cellVal != null) {
                    list.add(getValue(cell, cell.getCellType()));
                    if (colNo == 1) {
                        docNo = cellVal.toString();
                    }
                } else {
                    list.add("");
                }
            }

            hm.put(docNo, list);
        }

        return hm;
    }

    public IQueryDlg findQueryDlgForQueryClass(IQueryClass queryClass) {
        IQueryDlg dlg = null;
        if (queryClass != null) {
            dlg = queryClass.getQueryDlg("default");
        }

        return dlg;
    }

    public IQueryParameter query(ISession session, IQueryDlg queryDlg, Map<String, String> descriptorValues) {
        IDocumentServer documentServer = session.getDocumentServer();
        ISerClassFactory classFactory = documentServer.getClassFactory();
        IQueryParameter queryParameter = null;
        IQueryExpression expression = null;
        IComponent[] components = queryDlg.getComponents();

        for(int i = 0; i < components.length; ++i) {
            if (components[i].getType() == IMaskedEdit.TYPE) {
                IControl control = (IControl)components[i];
                String descriptorId = control.getDescriptorID();
                String value = (String)descriptorValues.get(descriptorId);
                if (value != null && value.trim().length() > 0) {
                    IDescriptor descriptor = documentServer.getDescriptor(descriptorId, session);
                    IQueryValueDescriptor queryValueDescriptor = classFactory.getQueryValueDescriptorInstance(descriptor);
                    queryValueDescriptor.addValue(value);
                    IQueryExpression expr = queryValueDescriptor.getExpression();
                    if (expression != null) {
                        expression = classFactory.getExpressionInstance(expression, expr, 0);
                    } else {
                        expression = expr;
                    }
                }
            }
        }

        if (expression != null) {
            queryParameter = classFactory.getQueryParameterInstance(session, queryDlg, expression);
        }

        return queryParameter;
    }

    public IDocumentHitList executeQuery(ISession session, IQueryParameter queryParameter) {
        IDocumentServer documentServer = session.getDocumentServer();
        return documentServer.query(queryParameter, session);
    }

    public IDocument getEngDocument(ISession session, String docNo, String docRev) throws IOException {
        IDocument result = null;
        IDocumentServer documentServer = session.getDocumentServer();
        this.descriptor1 = documentServer.getDescriptorForName(session, this.nameDescriptor1);
        this.descriptor2 = documentServer.getDescriptorForName(session, this.nameDescriptorRev);
        IQueryClass queryClass = documentServer.getQueryClassByName(session, this.searchClassName);
        IQueryDlg queryDlg = this.findQueryDlgForQueryClass(queryClass);
        Map<String, String> searchDescriptors = new HashMap();
        searchDescriptors.put(this.descriptor1.getId(), docNo);
        if(Objects.equals(docRev, "")) {
            searchDescriptors.put(this.descriptor2.getId(), "");
            IQueryParameter queryParameter = this.query(session, queryDlg, searchDescriptors);
            if (queryParameter != null) {
                IDocumentHitList hitresult = this.executeQuery(session, queryParameter);
                IDocument[] hits = hitresult.getDocumentObjects();
                queryParameter.close();
                return hits != null && hits.length > 0 ? hits[0] : null;
            } else {
                return null;
            }
        }else{
            searchDescriptors.put(this.descriptor2.getId(), docRev);
            IQueryParameter queryParameter = this.query(session, queryDlg, searchDescriptors);
            if (queryParameter != null) {
                IDocumentHitList hitresult = this.executeQuery(session, queryParameter);
                IDocument[] hits = hitresult.getDocumentObjects();
                queryParameter.close();
                result = hits != null && hits.length > 0 ? hits[0] : null;
                if(result == null){
                    searchDescriptors.put(this.descriptor2.getId(), "");
                    queryParameter = this.query(session, queryDlg, searchDescriptors);
                    if (queryParameter != null) {
                        hitresult = this.executeQuery(session, queryParameter);
                        hits = hitresult.getDocumentObjects();
                        queryParameter.close();
                        return hits != null && hits.length > 0 ? hits[0] : null;
                    } else {
                        return null;
                    }
                }
                return  result;
            } else {
                return null;
            }
        }
    }

    public static Object getValue(Cell cell, CellType type) {
        switch (type) {
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return getLocalDateTime(cell.getDateCellValue().toString());
                } else {
                    double value = cell.getNumericCellValue();
                    if (value == Math.floor(value)) {
                        return (long)value;
                    }

                    return value;
                }
            case STRING:
                return cell.getStringCellValue();
            case FORMULA:
                return getValue(cell, cell.getCachedFormulaResultType());
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case _NONE:
                return null;
            case BLANK:
                return null;
            case ERROR:
                return null;
            default:
                return null;
        }
    }

    public static LocalDateTime getLocalDateTime(String strDate) {
        strDate = strDate.replace("TRT", "Europe/Istanbul");
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("EEE MMM dd HH:mm:ss zzzz yyyy");
        ZonedDateTime zdt = ZonedDateTime.parse(strDate, formatter);
        LocalDateTime ldt = zdt.toLocalDateTime();
        return ldt;
    }

    public static LocalDate strTodate(String strDate) {
        strDate = strDate.replace("TRT", "Europe/Istanbul");
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MMMM d, yyyy", Locale.ENGLISH);
        LocalDate date = LocalDate.parse(strDate, formatter);
        return date;
    }
}
