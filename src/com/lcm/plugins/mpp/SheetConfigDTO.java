package com.lcm.plugins.mpp;

import java.util.Map;
import java.util.function.Function;

import net.sf.mpxj.Resource;
import net.sf.mpxj.ResourceAssignment;
import net.sf.mpxj.Task;

public class SheetConfigDTO {
    private Map<String, Function<Task, String>> taskHeaders;
    private String taskSheetName;
    private Map<String, Function<Resource, String>> resourceHeaders;
    private String resourceSheetName;
    private Map<String, Function<ResourceAssignment, String>> assignmentHeaders;
    private String assignmentSheetName;

    public SheetConfigDTO(Map<String, Function<Task, String>> taskHeaders, String taskSheetName,
                          Map<String, Function<Resource, String>> resourceHeaders, String resourceSheetName,
                          Map<String, Function<ResourceAssignment, String>> assignmentHeaders, String assignmentSheetName) {
        this.taskHeaders = taskHeaders;
        this.taskSheetName = taskSheetName;
        this.resourceHeaders = resourceHeaders;
        this.resourceSheetName = resourceSheetName;
        this.assignmentHeaders = assignmentHeaders;
        this.assignmentSheetName = assignmentSheetName;
    }

    public Map<String, Function<Task, String>> getTaskHeaders() {
        return taskHeaders;
    }

    public String getTaskSheetName() {
        return taskSheetName;
    }

    public Map<String, Function<Resource, String>> getResourceHeaders() {
        return resourceHeaders;
    }

    public String getResourceSheetName() {
        return resourceSheetName;
    }

    public Map<String, Function<ResourceAssignment, String>> getAssignmentHeaders() {
        return assignmentHeaders;
    }

    public String getAssignmentSheetName() {
        return assignmentSheetName;
    }
}
