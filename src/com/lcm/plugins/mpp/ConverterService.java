package com.lcm.plugins.mpp;

import com.appiancorp.suiteapi.content.ContentConstants;
import com.appiancorp.suiteapi.content.ContentService;
import com.appiancorp.suiteapi.content.DocumentOutputStream;
import com.appiancorp.suiteapi.knowledge.Document;
import com.appiancorp.suiteapi.process.exceptions.SmartServiceException;
import net.sf.mpxj.ProjectFile;
import net.sf.mpxj.Relation;
import net.sf.mpxj.Task;
import net.sf.mpxj.Resource;
import net.sf.mpxj.ResourceAssignment;
import net.sf.mpxj.reader.UniversalProjectReader;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.commons.io.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.lang.reflect.Method;
import java.util.List;
import java.util.LinkedHashMap;
import java.util.Locale;
import java.util.Map;
import java.util.ResourceBundle;
import java.util.function.Function;
import java.util.stream.Collectors;
import java.util.logging.Logger;

public class ConverterService {

    private static final String BUNDLE_NAME = "com.lcm.plugins.mpp.MPPToExcelConversion"; // Full bundle name with package
    private static final Logger LOG = Logger.getLogger("Log", BUNDLE_NAME);
    
    public Long convertMPPToExcel(ContentService contentService, Long mppDocument, Long saveInFolder, String taskHeadersJson, String taskSheetName, 
                                  String resourceHeadersJson, String resourceSheetName, String assignmentHeadersJson, String assignmentSheetName) throws SmartServiceException {
        Long excelDocument = null;

        try {
            ProjectFile projectFile = readProjectFile(contentService, mppDocument);

            Map<String, Function<Task, String>> taskHeaders = parseHeaders(taskHeadersJson, Task.class);
            Map<String, Function<Resource, String>> resourceHeaders = null;
            Map<String, Function<ResourceAssignment, String>> assignmentHeaders = null;

            if (resourceHeadersJson != null && !resourceHeadersJson.isEmpty()) {
                resourceHeaders = parseHeaders(resourceHeadersJson, Resource.class);
            }

            if (assignmentHeadersJson != null && !assignmentHeadersJson.isEmpty()) {
                assignmentHeaders = parseHeaders(assignmentHeadersJson, ResourceAssignment.class);
            }

            SheetConfigDTO sheetConfig = new SheetConfigDTO(taskHeaders, taskSheetName, resourceHeaders, resourceSheetName, assignmentHeaders, assignmentSheetName);

            XSSFWorkbook workbook = ExcelWorkbookFactory.createWorkbookFromProjectFile(projectFile, sheetConfig);
            String excelFilePath = saveWorkbookToFile(workbook, contentService, mppDocument);
            excelDocument = createDocument(contentService, saveInFolder, excelFilePath);
            uploadDocument(contentService, excelDocument, excelFilePath);
        } catch (Exception e) {
            LOG.info("Error converting MPP to Excel");
            ResourceBundle bundle = ResourceBundle.getBundle(BUNDLE_NAME, Locale.getDefault());
            String errorMessage = bundle.getString("error.convertingMPPToExcel");
            throw new SmartServiceException(null, e, errorMessage, null);
        }
        return excelDocument;
    }

    private ProjectFile readProjectFile(ContentService contentService, Long mppDocument) throws Exception {
        try (InputStream inputStream = contentService.download(mppDocument, ContentConstants.VERSION_CURRENT, false)[0].getInputStream()) {
            return new UniversalProjectReader().read(inputStream);
        }
    }

    private <T> Map<String, Function<T, String>> parseHeaders(String headersJson, Class<T> type) throws Exception {
        ObjectMapper objectMapper = new ObjectMapper();
        Map<String, String> headersMap = objectMapper.readValue(headersJson, new TypeReference<LinkedHashMap<String, String>>() {});
        Map<String, Function<T, String>> headers = new LinkedHashMap<>();
        headersMap.forEach((key, value) -> headers.put(key, createMethodInvoker(type, value)));
        return headers;
    }

    private <T> Function<T, String> createMethodInvoker(Class<T> type, String methodName) {
        if ("getPredecessors".equals(methodName)) {
            return item -> {
                try {
                    Method method = type.getMethod(methodName);
                    List<Relation> relations = (List<Relation>) method.invoke(item);
                    return relations.stream()
                            .map(relation -> String.valueOf(relation.getTargetTask().getID()))
                            .collect(Collectors.joining(","));
                } catch (Exception e) {
                    return "";
                }
            };
        } else {
            return item -> {
                try {
                    return type.getMethod(methodName).invoke(item).toString();
                } catch (Exception e) {
                    return "";
                }
            };
        }
    }

    private String saveWorkbookToFile(XSSFWorkbook workbook, ContentService contentService, Long mppDocument) throws Exception {
        String excelFilePath = contentService.download(mppDocument, ContentConstants.VERSION_CURRENT, false)[0].getName().replace(".mpp", ".xlsx");
        try (FileOutputStream fileOut = new FileOutputStream(excelFilePath)) {
            workbook.write(fileOut);
        }
        return excelFilePath;
    }

    private Long createDocument(ContentService contentService, Long saveInFolder, String excelFilePath) throws Exception {
        Document document = new Document(saveInFolder, excelFilePath, "xlsx");
        document.setState(ContentConstants.STATE_ACTIVE_PUBLISHED);
        document.setFileSystemId(ContentConstants.ALLOCATE_FSID);
        return contentService.create(document, ContentConstants.UNIQUE_NONE);
    }

    private void uploadDocument(ContentService contentService, Long excelDocument, String excelFilePath) throws Exception {
        try (InputStream targetStream = new FileInputStream(excelFilePath);
             DocumentOutputStream out = contentService.download(excelDocument, ContentConstants.VERSION_CURRENT, false)[0].getOutputStream()) {
            IOUtils.copy(targetStream, out);
        }
    }
}
