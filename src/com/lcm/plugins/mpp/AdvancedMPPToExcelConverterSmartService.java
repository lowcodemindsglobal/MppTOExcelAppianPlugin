package com.lcm.plugins.mpp;

import com.appiancorp.suiteapi.process.exceptions.SmartServiceException;
import com.appiancorp.suiteapi.process.framework.AppianSmartService;
import com.appiancorp.suiteapi.common.Name;
import com.appiancorp.suiteapi.content.ContentService;
import com.appiancorp.suiteapi.knowledge.DocumentDataType;
import com.appiancorp.suiteapi.knowledge.FolderDataType;
import com.appiancorp.suiteapi.process.framework.Input;
import com.appiancorp.suiteapi.process.framework.Required;
import com.appiancorp.suiteapi.process.framework.SmartServiceContext;
import com.appiancorp.suiteapi.process.palette.PaletteInfo;

@PaletteInfo(paletteCategory = "Appian Smart Services", palette = "Document Management")
public class AdvancedMPPToExcelConverterSmartService extends AppianSmartService {

    private Long mppDocument;
    private Long saveInFolder;
    private Long excelDocument;
    private String taskHeadersJson;
    private String taskSheetName;
    private String resourceHeadersJson;
    private String resourceSheetName;
    private String assignmentHeadersJson;
    private String assignmentSheetName;

    private final ContentService contentService;

    @Input(required = Required.ALWAYS)
    @Name("mppDocument")
    @DocumentDataType
    public void setMppDocument(Long val) {
        this.mppDocument = val;
    }

    @Name("excelDocument")
    @DocumentDataType
    public Long getExcelDocument() {
        return excelDocument;
    }

    @Input(required = Required.ALWAYS)
    @Name("saveInFolder")
    @FolderDataType
    public void setSaveInFolder(Long val) {
        this.saveInFolder = val;
    }

    @Input(required = Required.ALWAYS)
    @Name("taskHeadersJson")
    public void setTaskHeadersJson(String taskHeadersJson) {
        this.taskHeadersJson = taskHeadersJson;
    }

    @Input(required = Required.ALWAYS)
    @Name("taskSheetName")
    public void setTaskSheetName(String taskSheetName) {
        this.taskSheetName = taskSheetName;
    }

    @Input(required = Required.OPTIONAL)
    @Name("resourceHeadersJson")
    public void setResourceHeadersJson(String resourceHeadersJson) {
        this.resourceHeadersJson = resourceHeadersJson;
    }

    @Input(required = Required.OPTIONAL)
    @Name("resourceSheetName")
    public void setResourceSheetName(String resourceSheetName) {
        this.resourceSheetName = resourceSheetName;
    }

    @Input(required = Required.OPTIONAL)
    @Name("assignmentHeadersJson")
    public void setAssignmentHeadersJson(String assignmentHeadersJson) {
        this.assignmentHeadersJson = assignmentHeadersJson;
    }

    @Input(required = Required.OPTIONAL)
    @Name("assignmentSheetName")
    public void setAssignmentSheetName(String assignmentSheetName) {
        this.assignmentSheetName = assignmentSheetName;
    }

    public AdvancedMPPToExcelConverterSmartService(SmartServiceContext smartServiceCtx, ContentService cs) {
        super();
        this.contentService = cs;
    }

    @Override
    public void run() throws SmartServiceException {
        ConverterService converterService = new ConverterService();
        excelDocument = converterService.convertMPPToExcel(contentService, mppDocument, saveInFolder, taskHeadersJson, taskSheetName,
                                                           resourceHeadersJson, resourceSheetName, assignmentHeadersJson, assignmentSheetName);
    }
}
