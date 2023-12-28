//
// Source code recreated from a .class file by IntelliJ IDEA
// (powered by FernFlower decompiler)
//

package eng.ser.com;

import com.ser.blueline.*;
import com.ser.blueline.metaDataComponents.*;
import com.ser.foldermanager.*;
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
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

public class ImportProjectDocs extends UnifiedAgent {
    Logger log = LogManager.getLogger(this.getClass().getName());
    String nameDescriptor1 = "ccmPrjDocFileName";
    String nameDescriptorRev = "ccmPrjDocRevision";
    String searchClassName = "Search Engineering Documents";
    ISession ses = null;
    IDocumentServer srv = null;
    String prjCode = "";
    IDescriptor descriptor1;
    IDescriptor descriptor2;

    public ImportProjectDocs() {
    }

    protected Object execute() {
        this.log.info("Initiate the agent");
        if (this.getEventDocument() == null) {
            return this.resultError("Null Document object.");
        } else {
            ses = getSes();
            srv = ses.getDocumentServer();

            IDocument ldoc = this.getEventDocument();
            IUser owner = getDocumentServer().getUser(getSes() , ldoc.getOwnerID());
            boolean isDCCMember = existDCCGVList("CCM_PARAM_CONTRACTOR-MEMBERS","DCC",owner.getID());

            HashMap<Integer, String> flds = new HashMap();
            //flds.put(0, "ccmPrjDocFileName");
            flds.put(1, "ccmPrjDocNumber");
            flds.put(2, "ccmPrjDocRevision");
            flds.put(3, "ObjectName");
            flds.put(4, "ccmPrjDocCategory");
            flds.put(5, "ccmPrjDocOiginator");
            flds.put(6, "ccmPrjDocDiscipline");
            flds.put(7, "ccmPrjDocDocType");
            flds.put(8, "ccmPrjDocParentDoc");
            flds.put(9, "ccmPrjDocParentDocRevision");
            flds.put(10, "ccmPrjDocVendor");
            flds.put(11, "ccmPrjDocApprCode");
            flds.put(12, "ccmPrjDocIssueStatus");
            flds.put(13, "ccmPrjDocWBS2");
            flds.put(14, "ccmPrjDocDate");
            flds.put(15, "ccmPrjDocReqDate");
            flds.put(16, "ccmPrjDocDueDate");
            flds.put(17, "ccmPrjDocTransIncCode");

            try {

                String excelPath = this.exportDocumentContent(ldoc, "C:/tmp2/bulk/import");
                FileInputStream fist = new FileInputStream(excelPath);
                Workbook wrkb = new XSSFWorkbook(fist);
                HashMap<String, Row> list = listOfDocuments(wrkb);
                fist.close();
                Iterator var8 = list.entrySet().iterator();

                while(var8.hasNext()) {
                    List<String> nodeNames = new ArrayList<>();
                    Map.Entry<String, Row> line = (Map.Entry)var8.next();
                    prjCode = ldoc.getDescriptorValue("ccmPRJCard_code");
                    String docKey = (String)line.getKey();
                    Row row = (Row)line.getValue();
                    String docRev = row.getCell(1).getStringCellValue();
                    IDocument engDocument = this.getEngDocument(ses, prjCode, docKey);
                    if (engDocument != null) {
                        this.log.info("*** ENGINERING DOC FOUND :" + docKey);
                        this.log.info("Doc no is " + docKey + " Values is " + row.toString());
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

                            if(descName.contains("Date")){
                                DateFormat dt = new SimpleDateFormat("yyyyMMdd");
                                descDateValue = row.getCell(rowKey).getDateCellValue();
                                if(descDateValue!=null) {
                                    descValue = dt.format(descDateValue);
                                    engDocument.setDescriptorValue(descName, descValue);
                                }
                            }else if(row.getCell(rowKey).getCellType()==CellType.NUMERIC) {
                                descValue = String.valueOf((int) row.getCell(rowKey).getNumericCellValue());
                                engDocument.setDescriptorValue(descName, descValue);
                            }else{
                                if(descName.equals("ccmPrjDocOiginator")){
                                    descValue = row.getCell(rowKey).getStringCellValue().toUpperCase();
                                }else {
                                    descValue = row.getCell(rowKey).getStringCellValue();
                                }
                                engDocument.setDescriptorValue(descName, descValue);
                            }
                            var10000.info("DESC::" + descName + " // VALUE: " + descValue);
                        }
                        engDocument.setDescriptorValue("ccmReleased","1");
                        engDocument.setDescriptorValue("ccmPrjDocStatus","10");
                        if(isDCCMember) {
                            engDocument.setDescriptorValue("ccmPrjDocStatus", "50");
                        }
                        engDocument.commit();
                    }
                }
                this.log.info("Import ProjectDoc from Excel Finished");
                return this.resultSuccess("Ended successfully");
            } catch (Exception var15) {
                throw new RuntimeException(var15);
            }
        }
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
    public void setParent(List<IDocument> documentList){
        for(IDocument engDocument : documentList){
            this.log.info("Start Link.....");
            String chkKeyPrnt = engDocument.getDescriptorValue("ccmPrjDocParentDoc");
            String chkKeyCrrs = engDocument.getDescriptorValue("ccmPrjDocTransIncCode");
            this.log.info("Start Link for Parent Number:" + chkKeyPrnt + " child number:" + engDocument.getDescriptorValue("ccmPrjDocNumber"));
            if (prjCode != null){
                if(chkKeyPrnt != null){
                    IDocument prntDocument = getEngDocumentByNumber(ses, prjCode, chkKeyPrnt);
                    this.log.info("Parent Doc ? " + prntDocument);
                    if (prntDocument != null && !prntDocument.getDescriptorValue("ccmPrjDocCategory").trim().equalsIgnoreCase("TRANSMITTAL")) {
                        if (!Objects.equals(prntDocument.getID(), engDocument.getID())) {
                            ILink lnk2 = srv.createLink(ses, prntDocument.getID(), (INodeGeneric) null, engDocument.getID());
                            lnk2.commit();
                            this.log.info("Created Link...");
                        }
                    }
                }
                if(chkKeyCrrs != null){
                    IDocument prntDocument = this.getEngDocumentByNumber(ses, prjCode, chkKeyCrrs);
                    this.log.info("Parent (CRRS) Doc ? " + prntDocument);
                    if (prntDocument != null && prntDocument.getDescriptorValue("ccmPrjDocCategory").trim().equalsIgnoreCase("TRANSMITTAL")) {
                        if (!Objects.equals(prntDocument.getID(), engDocument.getID())) {
                            ILink lnk2 = srv.createLink(ses, prntDocument.getID(), (INodeGeneric) null, engDocument.getID());
                            lnk2.commit();
                            this.log.info("Created Link...");
                        }
                    }
                }
            }
        }
    }
    public void removeReleaseOldEngDoc(IDocument doc1){
        IDocument result = null;
        ISession session = this.getSes();
        String searchClassName = "Search Engineering Documents";
        IDocumentServer documentServer = session.getDocumentServer();
        IDescriptor descriptor1 = documentServer.getDescriptorForName(session, "ccmPrjDocNumber");
        IDescriptor descriptor2 = documentServer.getDescriptorForName(session, "ccmPRJCard_code");
        //IDescriptor descriptor2 = documentServer.getDescriptorForName(session, "ccmPrjDocRevision");
        IDescriptor descriptor3 = documentServer.getDescriptorForName(session, "ccmReleased");
        IQueryClass queryClass = documentServer.getQueryClassByName(session, searchClassName);
        IQueryDlg queryDlg = this.findQueryDlgForQueryClass(queryClass);
        Map<String, String> searchDescriptors = new HashMap();
        searchDescriptors.put(descriptor1.getId(), doc1.getDescriptorValue("ccmPrjDocNumber"));
        searchDescriptors.put(descriptor2.getId(), doc1.getDescriptorValue("ccmPRJCard_code"));
        searchDescriptors.put(descriptor3.getId(), "1");
        IQueryParameter queryParameter = this.query(session, queryDlg, searchDescriptors);
        if (queryParameter != null) {
            IDocumentHitList hitresult = this.executeQuery(session, queryParameter);
            IDocument[] hits = hitresult.getDocumentObjects();
            queryParameter.close();
            for(IDocument ldoc : hits){
                String docID = doc1.getID();
                String chkID = ldoc.getID();
                if(!Objects.equals(docID, chkID)){
//                    if(ldoc.getMutability() == IMutability.FIXED_CONTENT) {
//                        getSes().getDocumentServer().updateMutability(getSes(), ldoc, IMutability.MUTABLE);
//                        ldoc.commit();
//                    }
                    ldoc.setDescriptorValue("ccmReleased","0");
                    ldoc.commit();
                    //getSes().getDocumentServer().updateMutability(getSes(), ldoc, IMutability.FIXED_CONTENT);
                }
            }
        }
    }
    public INode createNodesByList(List<String> fNames) throws Exception {
        IFolder prjFolder = getProjectFolder();
        if(prjFolder == null){
            throw new Exception("Project folder not found.");
        }
        INode prjDocNode = prjFolder.getNodeByID(Constants.ClassIDs.ProjectDocsFolder);
        if(prjDocNode == null){
            throw new Exception("Project Docs. folder not found.");
        }
        INode newNode = null;
        INodes childs = null;
        for(String fname : fNames) {
            if(newNode == null) {
                childs = (INodes) prjDocNode.getChildNodes();
            }else{
                childs = (INodes) newNode.getChildNodes();
            }
            newNode = childs.getItemByName(fname);
            if(newNode == null) {
                newNode = childs.addNew(FMNodeType.STATIC);
                newNode.setName(fname);
                prjFolder.commit();
            }
        }
        log.info("Add NewNode Final ?? : " + newNode);
        return newNode;
    }
    public INode createNewNode(INode parentNode, String newNodeName) throws Exception {
        IFolder prjFolder = getProjectFolder();
        if(prjFolder == null){
            throw new Exception("Project folder not found.");
        }
        //List<INode> nodesByName = prjFolder.getNodesByName("Project Documents");
        INode prjDocNode = prjFolder.getNodeByID(Constants.ClassIDs.ProjectDocsFolder);
        if(prjDocNode == null){
            throw new Exception("Project Docs. folder not found.");
        }
        prjDocNode.refreshNodes(true);
        prjDocNode.refresh(true);
        INode childNode = null;
        INodes childNodes = (parentNode != null ? (INodes) parentNode.getChildNodes() : (INodes) prjDocNode.getChildNodes());
        childNode = childNodes.getItemByName(newNodeName);
        if(childNode == null) {
            childNode = childNodes.addNew(FMNodeType.STATIC);
            childNode.setName(newNodeName);
        }

        log.info("Add NewNode Final ?? : " + childNode);
        return childNode;
    }
    private INode addNewNode(IFolder folder, String rootName, String nodeName) throws Exception {
        log.info("Add NewNode Start: " + rootName + " under new Node: " + nodeName);
        List<INode> nodesByName = folder.getNodesByName(rootName);
        INode iNode = nodesByName.get(0);
        INodes root = (INodes) iNode.getChildNodes();
        INode newNode = root.getItemByName(nodeName);
        log.info("Find NewNode ?? : " + newNode);
        if(newNode == null) {
            newNode = root.addNew(FMNodeType.STATIC);
            newNode.setName(nodeName);
            folder.commit();
        }
        log.info("Add NewNode Final ?? : " + newNode);
        return newNode;
    }
    private boolean addToRootNode(IFolder folder, String rootName, String nodeName, IDocument pdoc) throws Exception {
        log.info("Add2RootNode Start");
        boolean add2Node = false;
        List<INode> nodesByName = folder.getNodesByName(rootName);
        INode iNode = nodesByName.get(0);
        INodes root = (INodes) iNode.getChildNodes();
        INode newNode = root.getItemByName(nodeName);
        if(newNode != null) {
            log.info("Find Node : " + newNode.getID() + " /// " + nodeName);
            boolean isExistElement = false;
            log.info("Start ProjectDoc exit in folder: " + isExistElement);
            IElements nelements = newNode.getElements();
            for(int i=0;i<nelements.getCount2();i++) {
                IElement nelement = nelements.getItem2(i);
                String edocID = nelement.getLink();
                String pdocID = pdoc.getID();
                if(Objects.equals(pdocID, edocID)){
                    isExistElement = true;
                    break;
                }
            }
            log.info("Finish ProjectDoc exit in folder: " + isExistElement);
            if(!isExistElement) {
                add2Node = folder.addInformationObjectToNode(pdoc.getID(), newNode.getID());
                log.info("ProjectDoc add to root folder: " + newNode.getID());
                folder.commit();
            }
        }
        log.info("ProjectDoc add to root node result : " + add2Node);
        return add2Node;
    }
    private boolean addToNode(INode newNode, IDocument pdoc) throws Exception {
        log.info("Add2Node Start");
        boolean add2Node = false;
        IFolder prjFolder = getProjectFolder();
        if(prjFolder == null){
            throw new Exception("Project folder not found.");
        }
        prjFolder.refresh(true);

        if(newNode != null) {
            newNode.refresh(true);
            boolean isExistElement = false;
            log.info("Start ProjectDoc exit in folder: " + isExistElement);
            IElements nelements = newNode.getElements();
            for(int i=0;i<nelements.getCount2();i++) {
                IElement nelement = nelements.getItem2(i);
                String edocID = nelement.getLink();
                String pdocID = pdoc.getID();
                if(Objects.equals(pdocID, edocID)){
                    isExistElement = true;
                    break;
                }
            }
            log.info("Finish ProjectDoc exit in folder: " + isExistElement);
            if(!isExistElement) {
                add2Node = prjFolder.addInformationObjectToNode(pdoc.getID(), newNode.getID());
                log.info("ProjectDoc add to folder: " + newNode.getID());
                if (add2Node) {
                    pdoc.setDescriptorValue("ccmLinkedFolderID", newNode.getID());
                    pdoc.commit();
                    log.info("ProjectDoc setting new node ID: " + newNode.getID());
                }
                prjFolder.commit();
            }
        }
        return add2Node;
    }
    public IFolder getProjectFolder() throws Exception {
        log.info("Getting Project Folder");
        String projectNumber = getEventDocument().getDescriptorValue("ccmPRJCard_code");
        if(projectNumber == null) throw new Exception("Project Number is NULL");
        if(projectNumber.isEmpty()) throw new Exception("Project Number is NULL");
        StringBuilder whereClause = new StringBuilder();
        whereClause.append("TYPE = '")
                .append(Constants.ClassIDs.ProjectFolder)
                .append("' AND ")
                .append(Constants.Literals.ProjectNumberDescriptor)
                .append(" = '")
                .append(projectNumber).append("'");
        log.info("Attemptign Query");
        IInformationObject[] objects = createQuery(Constants.Literals.ProjectFolderDB , whereClause.toString() , 2);
        if(objects == null) throw new Exception("Not Folder with: " + projectNumber + " was found");
        if(objects.length < 1)throw new Exception("Not Folder with: " + projectNumber + " was found");
        return (IFolder) objects[0];
    }
    private IInformationObject[] createQuery(String dbName , String whereClause , int maxHits){
        String[] databaseNames = {dbName};

        ISerClassFactory fac = getDocumentServer().getClassFactory();
        IQueryParameter que = fac.getQueryParameterInstance(
                getSes() ,
                databaseNames ,
                fac.getExpressionInstance(whereClause) ,
                null,null);
        que.setMaxHits(maxHits);
        que.setHitLimit(maxHits + 1);
        que.setHitLimitThreshold(maxHits + 1);
        IDocumentHitList hits = que.getSession() != null? que.getSession().getDocumentServer().query(que, que.getSession()):null;
        if(hits == null) return null;
        else return hits.getInformationObjects();
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
                    if (!indx.equals("") && !indx.equals("File Name")) {
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
    public IDocument getEngDocumentByNumber(ISession session, String prjCode, String docKey) {
        IDocument result = null;
        log.info("Search Eng Document By PRJ Code:" + prjCode);
        log.info("Search Eng Document By Number:" + docKey);
        IDocumentServer documentServer = session.getDocumentServer();
        this.descriptor1 = documentServer.getDescriptorForName(session, "ccmPRJCard_code");
        this.descriptor2 = documentServer.getDescriptorForName(session, "ccmPrjDocNumber");
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
    public IDocument checkDublicateEngDocByFileName(IDocument doc1){
        IDocument result = null;
        ISession session = this.getSes();
        String searchClassName = "Search Engineering Documents";
        IDocumentServer documentServer = session.getDocumentServer();
        IDescriptor descriptor1 = documentServer.getDescriptorForName(session, "ccmPRJCard_code");
        IDescriptor descriptor2 = documentServer.getDescriptorForName(session, "ccmPrjDocFileName");
        IQueryClass queryClass = documentServer.getQueryClassByName(session, searchClassName);
        IQueryDlg queryDlg = this.findQueryDlgForQueryClass(queryClass);
        Map<String, String> searchDescriptors = new HashMap();
        searchDescriptors.put(descriptor1.getId(), doc1.getDescriptorValue("ccmPRJCard_code"));
        searchDescriptors.put(descriptor2.getId(), doc1.getDescriptorValue("ccmPrjDocFileName"));
        IQueryParameter queryParameter = this.query(session, queryDlg, searchDescriptors);
        if (queryParameter != null) {
            IDocumentHitList hitresult = this.executeQuery(session, queryParameter);
            IDocument[] hits = hitresult.getDocumentObjects();
            queryParameter.close();
            for(IDocument ldoc : hits){
                String docID = doc1.getID();
                String chkID = ldoc.getID();
                if(!Objects.equals(docID, chkID)){
                    result = ldoc;
                    break;
                }
            }
        }
        return result;
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

}
