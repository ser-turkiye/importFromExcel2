//
// Source code recreated from a .class file by IntelliJ IDEA
// (powered by FernFlower decompiler)
//

package eng.ser.com;

import com.ser.blueline.*;
import com.ser.blueline.bpm.IBpmService;
import com.ser.blueline.bpm.IWorkbasket;
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
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

public class ImportUsers extends UnifiedAgent {
    Logger log = LogManager.getLogger(this.getClass().getName());
    String nameDescriptor1 = "ccmPrjDocNumber";
    String nameDescriptorRev = "ccmPrjDocRevision";
    String searchClassName = "Search Engineering Documents";
    IDescriptor descriptor1;
    IDescriptor descriptor2;
    protected Object execute() {
        this.log.info("Initiate the agent");
        ISession ses = this.getSes();
        IDocumentServer srv = ses.getDocumentServer();
        ISerClassFactory classFactory = srv.getClassFactory();
        IBpmService bpmService = ses.getBpmService();
        boolean res = false;

        ses.refreshServerSessionCache();

        HashMap<Integer, String> flds = new HashMap();
        flds.put(0, "Team");
        flds.put(1, "Email");
        flds.put(2, "UserName");
        flds.put(3, "Name");
        flds.put(4, "Surname");
        flds.put(5, "Position");

        try {
            String excelPath = "C:/tmp2/bulk/users/GIBUsers.xlsx";
            FileInputStream fist = new FileInputStream(excelPath);
            Workbook wrkb = new XSSFWorkbook(fist);
            HashMap<String, Row> list = listOfDocuments(wrkb);
            fist.close();
            Iterator var8 = list.entrySet().iterator();
            String team, email, username, name, surname, position;
            String password = "Gib@2024";
            while(var8.hasNext()) {
                Map.Entry<String, Row> line = (Map.Entry)var8.next();
                //String rowNo = (String)line.getKey();
                Row row = (Row)line.getValue();
                team = row.getCell(1).getStringCellValue().trim();
                email = row.getCell(2).getStringCellValue().trim().replaceAll(" ","");
                email = email.replaceAll(" ","");
                username = row.getCell(3).getStringCellValue().trim();
                name = row.getCell(4).getStringCellValue().trim();
                surname = row.getCell(5).getStringCellValue().trim();
                position = row.getCell(6).getStringCellValue().trim();

                String unitName = team;
                IUnit unit = getDocumentServer().getUnitByName(getSes(), unitName);

                IUser user = getDocumentServer().getUserByLoginName(getSes(),username);
                if(user == null){
                    user = classFactory.createUserInstance(getSes(),username,password);
                    user.commit();
                }
                if(user!=null) {
                    IUser cuser = user.getModifiableCopy(getSes());
                    cuser.setLicenseType(LicenseType.NORMAL_USER);
                    cuser.setFirstName(name);
                    cuser.setLastName(surname);
                    cuser.setEMailAddress(email);
                    cuser.setDescription(position);
                    cuser.commit();
                    log.info("user created:" + cuser.getFullName());
                    IWorkbasket wb = bpmService.getWorkbasketByAssociatedOrgaElement((IOrgaElement) cuser);
                    if(wb == null) {
                        wb = bpmService.createWorkbasketObject((IOrgaElement) cuser);
                        wb.commit();
                        IWorkbasket wbCopy = wb.getModifiableCopy(getSes());
                        wbCopy.setNotifyEMail(email);
                        wbCopy.setOwner(cuser);
                        IRole admRole = getSes().getDocumentServer().getRoleByName(getSes(),"admins");
                        if(admRole != null) {
                            res = wbCopy.addAccessibleBy(admRole);
                        }
                        wbCopy.commit();
                    }
                    if(unit != null){
                        this.addToUnit(cuser,unit.getID());
                        log.info("add user:" + cuser.getFullName() + " to unit " + unitName);
                    }
                }
                ses.refreshServerSessionCache();
                log.info("----User Updated --- for (User):" + username);
            }
            log.info("Finished");
            return this.resultSuccess("Ended successfully");
        } catch (Exception e) {
            log.error("Exception Caught");
            log.error(e.getMessage());
            return resultError(e.getMessage());
        }
    }
    public void addToUnit(IUser user, String unitID) throws Exception {
        try {
            String[] unitIDs = (user != null ? user.getUnitIDs() : null);
            boolean isExist = Arrays.asList(unitIDs).contains(unitID);
            if(!isExist){
                List<String> rtrn = new ArrayList<String>(Arrays.asList(unitIDs));
                rtrn.add(unitID);
                IUser cuser = user.getModifiableCopy(getSes());
                String[] newUnitIDs = rtrn.toArray(new String[0]);
                cuser.setUnitIDs(newUnitIDs);
                cuser.commit();
            }
        }catch (Exception e){
            throw new Exception("Exeption Caught..addToRole : " + e);
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
