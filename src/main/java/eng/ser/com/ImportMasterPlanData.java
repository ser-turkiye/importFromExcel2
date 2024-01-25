//
// Source code recreated from a .class file by IntelliJ IDEA
// (powered by FernFlower decompiler)
//

package eng.ser.com;

import com.ser.blueline.*;
import com.ser.blueline.metaDataComponents.*;
import com.ser.blueline.modifiablemetadata.IArchiveFolderClassModifiable;
import com.ser.foldermanager.IFolder;
import com.ser.foldermanager.IFolderConnection;
import de.ser.doxis4.agentserver.UnifiedAgent;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

public class ImportMasterPlanData extends UnifiedAgent {
    Logger log = LogManager.getLogger(this.getClass().getName());
    String nameDescriptor1 = "ccmPrjDocNumber";
    String nameDescriptor2 = "ccmPrjDocIssueStatus";
    String nameDescriptorRev = "ccmPrjDocRevision";
    String searchClassName = "Search Engineering Documents";
    ISession ses = null;
    IDocumentServer srv = null;
    String prjCode = "";
    String prjDocNumber = "";
    String prjDocDiscipline = "";
    IDescriptor descriptor1;
    IDescriptor descriptor2;
    CellType cellType;
    public ImportMasterPlanData() {
    }

    protected Object execute() {
        this.log.info("Initiate the agent");
        if (this.getEventDocument() == null) {
            return this.resultError("Null Document object.");
        } else {
            ses = getSes();
            srv = ses.getDocumentServer();
            Integer cnt = 0;
            HashMap<Integer, String> flds = new HashMap();
            IDocument ldoc = this.getEventDocument();
            try {
                String excelPath = this.exportDocumentContent(ldoc, "C:/tmp2/bulk/import");
                FileInputStream fist = new FileInputStream(excelPath);
                this.log.info("Exported excel file to path:" + excelPath);
                Workbook wrkb = new XSSFWorkbook(fist);
                HashMap<String, Row> list = listOfDocuments(wrkb);
                List<String> fields = fieldsOfDocuments(wrkb);
                fist.close();

                for(String field : fields){
                    flds.put(cnt, field);
                    cnt++;
                }
                ///ilk dongude doc no, rev no guncelle
                Iterator var8 = list.entrySet().iterator();
                this.log.info("Start second loop.");
                var8 = list.entrySet().iterator();
                while(var8.hasNext()) {
                    List<String> nodeNames = new ArrayList<>();
                    Map.Entry<String, Row> line = (Map.Entry)var8.next();
                    prjCode = ldoc.getDescriptorValue("ccmPRJCard_code");
                    String docKey = (String)line.getKey();
                    Row row = (Row)line.getValue();
                    prjDocNumber = getValueFromRow(fields, row, "ccmPrjDocNumber");
                    prjDocDiscipline = getValueFromRow(fields, row, "ccmPrjDocIssueStatus");
                    IFolder MDPRecord = this.getMDPRecord(prjCode, prjDocNumber, prjDocDiscipline);
                    if (MDPRecord == null) {
                        MDPRecord = newMDPRecord();
                    }
                    if (MDPRecord != null) {
                        this.log.info("ENGINERING DOC FOUND :" + docKey);
                        MDPRecord.setDescriptorValueTyped("ccmPRJCard_code", prjCode);
                        Iterator var13 = flds.entrySet().iterator();

                        while(var13.hasNext()) {
                            Map.Entry<Integer, String> ffld = (Map.Entry)var13.next();
                            if (row.getCell((Integer)ffld.getKey()) == null) {
                                row.createCell((Integer)ffld.getKey());
                            }
                            String descValue = "";
                            Date descDateValue = new Date();

                            int rowKey = (Integer)ffld.getKey();
                            String descName = (String)ffld.getValue();
                            this.log.info("START SET DESCRIPTOR NAME :" + descName);
                            if(descName.contains("Date")){
                                if(row.getCell(rowKey).getCellType()==CellType.STRING) {
                                    String sDate1 = row.getCell(rowKey).getStringCellValue();
                                    Date date1 = new SimpleDateFormat("dd.MM.yyyy").parse(sDate1);
                                    DateFormat dt = new SimpleDateFormat("yyyyMMdd");
                                    descValue = dt.format(date1);
                                    MDPRecord.setDescriptorValueTyped(descName, descValue);
                                }else {
                                    DateFormat dt = new SimpleDateFormat("yyyyMMdd");
                                    descDateValue = row.getCell(rowKey).getDateCellValue();
                                    if (descDateValue != null) {
                                        descValue = dt.format(descDateValue);
                                        MDPRecord.setDescriptorValueTyped(descName, descValue);
                                    }
                                }
                            }else if(row.getCell(rowKey).getCellType()==CellType.NUMERIC) {
                                descValue = String.valueOf((int) row.getCell(rowKey).getNumericCellValue());
                                MDPRecord.setDescriptorValueTyped(descName, descValue);
                            }else{
                                descValue = row.getCell(rowKey).getStringCellValue();
                                MDPRecord.setDescriptorValueTyped(descName, descValue);
                            }
                            this.log.info("FINISH SET THIS VALUE: " + descValue);
                        }
                        MDPRecord.commit();
                    }
                }
                this.log.info("Import ProjectDoc from Excel Finished");
                return this.resultSuccess("Ended successfully");
            } catch (Exception e) {
                this.log.error("Exception Caught");
                this.log.error(e.getMessage());
                //return resultError(e.getMessage());
            }
        }
        return resultSuccess("Success");
    }
    public boolean existDCCGVList(String paramName, String key1, String key2) {
        boolean rtrn = false;
        IStringMatrix settingsMatrix = getDocumentServer().getStringMatrix(paramName, getSes());
        String rowValuePrjCode = "";
        String rowValueParamUserID = "";
        String rowValueParamDCC = "";
        String rowValueParamMyComp = "";
        for(int i = 0; i < settingsMatrix.getRowCount(); i++) {
            rowValuePrjCode = settingsMatrix.getValue(i, 0);
            rowValueParamUserID = settingsMatrix.getValue(i, 5);
            rowValueParamDCC = settingsMatrix.getValue(i, 6);
            rowValueParamMyComp = settingsMatrix.getValue(i, 7);

            //if (!Objects.equals(rowValuePrjCode, prjCode)){continue;}
            if (!Objects.equals(rowValueParamDCC, key1)){continue;}
            if (!Objects.equals(rowValueParamUserID, key2)){continue;}
            if (!Objects.equals(rowValueParamMyComp, "1")){continue;}

            return true;
        }
        return rtrn;
    }
    public static HashMap<String, Row> listOfDocuments(Workbook workbook) throws IOException {
        HashMap<String, Row> rtrn = new HashMap();
        Sheet sheet = workbook.getSheetAt(0);
        Iterator var3 = sheet.iterator();

        while(var3.hasNext()) {
            Row row = (Row)var3.next();
            if (row.getRowNum() != 0 && row.getRowNum() != 1) {
                Cell cll1 = row.getCell(0);
                if (cll1 != null) {
                    String indx = cll1.getRichStringCellValue().getString();
                    if (!indx.equals("") && !indx.equals("File Name")) {
                        rtrn.put(indx, row);
                    }
                }
            }
        }
        return rtrn;
    }
    public static List<String> fieldsOfDocuments(Workbook workbook) throws IOException {
        List<String> rtrn = new ArrayList<>();
        Sheet sheet = workbook.getSheetAt(0);
        String indx = "";
        Iterator var3 = sheet.iterator();
        Integer c = 0;
        while(var3.hasNext()) {
            Row row = (Row)var3.next();
            if (row.getRowNum() == 0) {
                while (c<35) { ///toplam fields sayısı
                    Cell cll1 = row.getCell(c);
                    if (cll1 != null) {
                        indx = cll1.getRichStringCellValue().getString();
                        if(Objects.equals(indx, "")){break;}
                        rtrn.add(indx);
                    }else{break;}
                    c++;
                }
            } else{break;}
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
    public IDocument getEngDocument(ISession session, String prjCode, String docKey) throws IOException {
        IDocument result = null;
        IDocumentServer documentServer = session.getDocumentServer();
        this.descriptor1 = documentServer.getDescriptorForName(session, "ccmPRJCard_code");
        this.descriptor2 = documentServer.getDescriptorForName(session, this.nameDescriptor1);
        IQueryClass queryClass = documentServer.getQueryClassByName(session, this.searchClassName);
        IQueryDlg queryDlg = this.findQueryDlgForQueryClass(queryClass);
        Map<String, String> searchDescriptors = new HashMap();
        searchDescriptors.put(this.descriptor1.getId(), prjCode);
        searchDescriptors.put(this.descriptor2.getId(), docKey);
        IQueryParameter queryParameter = this.query(session, queryDlg, searchDescriptors);
        if (queryParameter != null) {
            IDocumentHitList hitresult = this.executeQuery(session, queryParameter);
            IDocument[] hits = hitresult.getDocumentObjects();
            queryParameter.close();
            return hits != null && hits.length > 0 ? hits[0] : null;
        } else {
            return null;
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
    public IArchiveFolderClass createNewMDPRecord(){
        try{
            log.info("Creating new MDP record..");
            IMetaDataManager dataManager = srv.getMetaDataManager(getSes());
            IArchiveFolderClassModifiable archiveFolderClassModifiable = dataManager.createArchiveFolderClass(Constants.ClassIDs.MDPRecord);
            archiveFolderClassModifiable.commit();
            //ISerClassFactory classFactory = srv.getClassFactory();
            //IArchiveClass aClass = srv.getArchiveClass(id,getSes());
            //rtrn = classFactory.getDocumentInstance(getSes(), aClass, getSes().getDatabaseByName("PRJ_FOLDER"), null);
            return (IArchiveFolderClass) archiveFolderClassModifiable;
        }catch(Exception e){
            log.error("Exception caught...createNewMDPRecord..:" + e.getMessage());
            return null;
        }
    }
    private IFolder newMDPRecord() throws Exception{
        try{
            IFolderConnection folderConnection = ses.getFolderConnection();
            IFolder rtrn = folderConnection.createFolder();
            IArchiveFolderClass afc = srv.getArchiveFolderClass(Constants.ClassIDs.MDPRecord, ses);
            IDatabase db = ses.getDatabase(afc.getDefaultDatabaseID());
            rtrn.init(afc);
            rtrn.setDatabaseName(db.getDatabaseName());
            return rtrn;
        }catch(Exception e){
            log.error("Exception caught...newMDPRecord..:" + e.getMessage());
            return null;
        }
    }
    public IFolder getMDPRecord(String prjCode, String prjNumber, String prjDiscipline)  {
        StringBuilder builder = new StringBuilder();
        builder.append("TYPE = '").append(Constants.ClassIDs.MDPRecord).append("'")
                .append(" AND ")
                .append(Constants.Literals.ProjectNumberDescriptor).append(" = '").append(prjCode).append("'")
                .append(" AND ")
                .append(Constants.Literals.ProjectDocNumberDescriptor).append(" = '").append(prjNumber).append("'")
                .append(" AND ")
                .append(Constants.Literals.ProjectDisciplineDescriptor).append(" = '").append(prjDiscipline).append("'");
        String whereClause = builder.toString();
        System.out.println("Where Clause: " + whereClause);

        IInformationObject[] informationObjects = createQuery(new String[]{Constants.Databases.MDPRecordDB} , whereClause , 1);
        if(informationObjects.length < 1) {return null;}
        return (IFolder) informationObjects[0];
    }
    public IInformationObject[] createQuery(String[] dbNames , String whereClause , int maxHits){
        String[] databaseNames = dbNames;

        ISerClassFactory fac = getSrv().getClassFactory();
        IQueryParameter que = fac.getQueryParameterInstance(
                getSes() ,
                databaseNames ,
                fac.getExpressionInstance(whereClause) ,
                null,null);
        if(maxHits > 0) {
            que.setMaxHits(maxHits);
            que.setHitLimit(maxHits + 1);
            que.setHitLimitThreshold(maxHits + 1);
        }
        IDocumentHitList hits = que.getSession() != null? que.getSession().getDocumentServer().query(que, que.getSession()):null;
        if(hits == null) return null;
        else return hits.getInformationObjects();
    }
    public static String getValueFromRow(List<String> fields,Row row, String fieldName) throws IOException {
        String rtrn = "";

        int c = 0;
        for (String field :fields) { ///toplam fields sayısı
            if(Objects.equals(fieldName, field)) {
                Cell cll1 = row.getCell(c);
                if (cll1 != null) {
                    rtrn = cll1.getRichStringCellValue().getString();
                    break;
                }
            }
            c++;
        }
        return rtrn;
    }
}
