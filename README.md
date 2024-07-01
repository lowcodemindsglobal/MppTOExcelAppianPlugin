# Overview
The MPP to Excel Smart Service Plugin in Appian is a tool designed to facilitate the conversion of Microsoft Project files (MPP) into Excel spreadsheets. This plugin is particularly useful in business processes where project data needs to be extracted, manipulated, or analyzed in a more flexible and widely-used format like Excel.

## Purpose and Use Cases
  
The primary purpose of the MPP to Excel Smart Service Plugin is to convert project plans created in Microsoft Project into Excel format. This conversion allows for easier data manipulation, reporting, and integration with other business processes or systems. 
## Typical use cases include:
- Project Management: Simplifying the sharing and reviewing of project plans with stakeholders who may not have access to Microsoft Project.
- Data Analysis: Enabling detailed analysis of project data using Excel's advanced data manipulation and visualization features.
- Reporting: Facilitating the creation of custom reports and dashboards by integrating project data into broader reporting frameworks.
## Key Features
- File Conversion: The plugin takes an MPP file as input and converts it into an Excel file. This includes the extraction of various project elements such as tasks, resources, timelines, and milestones.
- Customization Options: Users can customize the conversion process to include specific data fields or formats as needed for their particular use case.
- Automation: The plugin can be integrated into Appian workflows, allowing for automated conversion of MPP files as part of a larger business process. This reduces

# Design Details
Here are the primary design patterns used:

Builder Pattern: Used implicitly within SheetConfigDTO. This pattern simplifies object creation by allowing step-by-step configuration and building of complex objects.

Strategy Pattern: The use of `Function<Task, String>` and similar function mappings for resources and assignments allows different strategies for extracting data from tasks, resources, and assignments. This pattern is evident in the creation of headers and the flexible handling of different data extraction methods.

- Factory Pattern: The ExcelWorkbookFactory class is a clear example of the Factory Pattern. It abstracts the creation of different sheets in an Excel workbook, making the creation logic reusable and encapsulated.

- Template Method Pattern: Within ExcelWorkbookFactory, the createSheet method acts as a template method, defining the skeleton of how a sheet is created while allowing the specific details (like the type of data) to be handled via parameterization.

- Data Transfer Object (DTO) Pattern: SheetConfigDTO is an example of the DTO pattern. It encapsulates the data required for configuring sheets, making the method signatures cleaner and the data flow more explicit.

- Service Layer Pattern: MPPToExcelConverterService serves as a service layer, handling business logic related to converting an MPP file to an Excel workbook. This pattern helps in separating business logic from the Appian smart service framework.

Detailed Pattern Usage:
- Builder Pattern:

Though not explicitly implemented as a builder, SheetConfigDTO allows for the construction of a configuration object that can be passed around, mimicking a builder's role by consolidating related parameters.
 - Strategy Pattern:

The mappings for headers `(Map<String, Function<Task, String>>, Map<String, Function<Resource, String>>, etc.)` allow different strategies for extracting data from different objects. This encapsulation of behavior aligns with the Strategy Pattern.
- Factory Pattern:

ExcelWorkbookFactory.createWorkbookFromProjectFile and the helper method createSheet encapsulate the logic for creating different sheets based on given configurations.

- Template Method Pattern:

The method createSheet in ExcelWorkbookFactory provides a template for creating a sheet, with the specific details of how to handle rows and cells being passed as parameters.
- Data Transfer Object (DTO) Pattern:

SheetConfigDTO encapsulates the configuration data required for creating sheets in the workbook, improving the clarity and cleanliness of method signatures.
- Service Layer Pattern:

MPPToExcelConverterService acts as a mediator between the Appian plugin and the actual business logic required to convert an MPP file to Excel, handling the orchestration of various tasks and encapsulating the logic in a service layer.

# Release info
- apache poi 5.2.1 cannot be used together with poi-ooxml-schemas-4.1.2.jar. It needs to be used with either poi-ooxml-lite-5.2.1.jar or poi-ooxml-full-5.2.1.jar.

## Project Structure
![image](https://github.com/Sangeerththan/MppTOExcelAppianPlugin/assets/25486160/4653d44a-a8d9-4f28-942d-e697b77eea53)
## Dependencies
```
<?xml version="1.0" encoding="UTF-8"?>
<classpath>
	<classpathentry kind="con" path="org.eclipse.jdt.launching.JRE_CONTAINER/org.eclipse.jdt.internal.debug.ui.launcher.StandardVMType/JavaSE-1.8">
		<attributes>
			<attribute name="module" value="true"/>
		</attributes>
	</classpathentry>
	<classpathentry kind="src" path="src"/>
	<classpathentry kind="lib" path="C:/Users/SangeerththanBalacha/Desktop/2024/Appian/PluginDevelopment/Template-Plugin/libraries/appian-plug-in-sdk.jar"/>
	<classpathentry kind="lib" path="src/META-INF/lib/apache-logging-log4j.jar"/>
	<classpathentry kind="lib" path="src/META-INF/lib/log4j-api-2.20.0.jar"/>
	<classpathentry kind="lib" path="src/META-INF/lib/mpxj-12.10.1.jar"/>
	<classpathentry kind="lib" path="src/META-INF/lib/poi-5.2.5.jar"/>
	<classpathentry kind="lib" path="src/META-INF/lib/commons-logging-1.2.jar"/>
	<classpathentry kind="lib" path="src/META-INF/lib/log4j-core-2.20.0.jar"/>
	<classpathentry kind="lib" path="src/META-INF/lib/commons-io-2.15.0.jar"/>
	<classpathentry kind="lib" path="src/META-INF/lib/xmlbeans-5.2.1.jar"/>
	<classpathentry kind="lib" path="src/META-INF/lib/commons-compress-1.26.2.jar"/>
	<classpathentry kind="lib" path="src/META-INF/lib/poi-ooxml-lite-5.2.5.jar"/>
	<classpathentry kind="lib" path="src/META-INF/lib/jackson-databind-2.17.1.jar"/>
	<classpathentry kind="lib" path="src/META-INF/lib/jackson-core-2.17.1.jar"/>
	<classpathentry kind="lib" path="src/META-INF/lib/poi-ooxml-5.2.5.jar"/>
	<classpathentry kind="output" path="bin"/>
</classpath>
```
![image](https://github.com/Sangeerththan/MppTOExcelAppianPlugin/assets/25486160/4991d03c-fd07-4cf4-a38f-e96acedbbfce)
- Exporting Jar In Eclipse
![image](https://github.com/Sangeerththan/MppTOExcelAppianPlugin/assets/25486160/2e4adf55-bd7b-43f4-b944-4f31d91f9c88)

## Configuration

- Configuration Provides a way to use headers with the respective functionalities
- Supported parameters: `Mpp Document`,`Save In Folder`,`Sheet Name`,`Headers Json`

![image](https://github.com/Sangeerththan/MppTOExcelAppianPlugin/assets/25486160/12b2293d-4c98-4024-930e-2fd32c023775)
![image](https://github.com/Sangeerththan/MppTOExcelAppianPlugin/assets/25486160/6d9ade26-aca4-4226-acea-502bc32ca7e4)
![image](https://github.com/Sangeerththan/MppTOExcelAppianPlugin/assets/25486160/f378b011-9902-4b96-acf4-10416a4fb014)
![image](https://github.com/Sangeerththan/MppTOExcelAppianPlugin/assets/25486160/b157021a-1a6c-41eb-97a3-e056b8aaf501)
![image](https://github.com/Sangeerththan/MppTOExcelAppianPlugin/assets/25486160/ba56c485-cb34-4b95-90cf-dee583adc978)
![image](https://github.com/Sangeerththan/MppTOExcelAppianPlugin/assets/25486160/3a64cdaf-1fd9-4dcd-9740-b910203083df)
![image](https://github.com/Sangeerththan/MppTOExcelAppianPlugin/assets/25486160/8742e69d-894d-462a-a95a-c152bbee9bfb)
![image](https://github.com/Sangeerththan/MppTOExcelAppianPlugin/assets/25486160/38ef29c7-e1ba-4bfb-a27f-7cb7e7f3cbd6)
![image](https://github.com/Sangeerththan/MppTOExcelAppianPlugin/assets/25486160/f5c177be-d18d-45d7-802a-f091acc0c415)
![image](https://github.com/Sangeerththan/MppTOExcelAppianPlugin/assets/25486160/9315d529-4b41-44ff-89f5-9738799c9af6)
![image](https://github.com/Sangeerththan/MppTOExcelAppianPlugin/assets/25486160/8dd49e1c-3285-48b0-ab04-d64f1e8eb2de)
![image](https://github.com/Sangeerththan/MppTOExcelAppianPlugin/assets/25486160/ad76a434-95e3-49a2-aa22-4299d6e40bbc)
![image](https://github.com/Sangeerththan/MppTOExcelAppianPlugin/assets/25486160/bb79c95d-ce12-4079-b55c-7e11c6948235)



-  Sample Task Headers Json,Assignments headers Json,resource Headers Json
  
 ```
a!toJson(
  {
    "id": "getID",
    "name": "getName",
    "start": "getStart",
    "end": "getFinish",
    "progress": "getPercentageComplete",
    "actualStart": "getActualStart",
    "resourceName": "getResourceNames",
    "resourceGroup": "getResourceGroup",
    "dependencies": "getPredecessors",
    "work": "getWork",
    "duration": "getDuration",
    "scheduledStart": "getScheduledStart"
  }
)

```


```
###Assignments has Task
a!toJson({ "task": "getTask" })
```


```
### Resources have name
a!toJson({ "resource": "getName" })
```
## List of Methods Supported In Tasks

```
getID(), 
getUniqueID(),
getName(),
 getStart(),
 getFinish(),
 getDuration(),
 getOutlineLevel(),
getPriority(),
getPercentComplete(),
 getNotes(),
 getResources(),
 getConstraintDate(),
 getConstraintType(),
 getConstraint(),
 getActualStart(),
 getActualFinish(),
 getBaselineStart(),
getBaselineFinish(),
 getBaselineDuration(),
 getPredecessors(),
 getSuccessors(),
 getBaselineStart(int),
 getBaselineFinish(int),
 getBaselineDuration(int),
 getNumberBaselines(),
 getBaselineNumber(),
 getCachedValue(int),
 getCachedValue(String),
 getCustomField(String),
getCustomFields(),
 getCost(),
getFixedCost(),
 getActualCost(),
 getRemainingCost(),
 getWork(),
getActualWork(),
 getRemainingWork(),
 getUnits(),
 getActualUnits(),
 getRemainingUnits(),
 getEarnedValueCost(),
getEarnedValueScheduleIndicator(),
 getPercentageComplete(),
 getActualPercentageComplete(),
 getPV(),
 getEV(),
 getAC(),
 getSPI(),
 getCPI(),
 getBCWS(),
getBCWP(),
 getIsMarked(),
 getOutlineNumber(),
 getOutlineSequenceNumber(),
isCritical(),
isMilestone(),
isSummary(),
 getTotalSlack(),
 getFreeSlack(),
getStartSlack(),
getDeadline(),
 getCalendar(),
 getTaskMode(),
 getExternalTask(),
 getResume(),
getSuspend(),
getManual(),
 getLevelingDelay(), getActualDuration(), getHasNotes(), getHideBar(), getBarStyle(),
 getVisible(), getRollup(), getIgnoreResourceCalendar(), getEarnedValueMethod(),
getHideBar(), getIgnoreWarnings(), getIgnoreResourceCalendar(), getIgnoreCost(),
 getResourceId(), getResourceInitials(), getResourceName(), getResourceUniqueID(),
getResourceGroup(), getResourceAssignmentId(), getResourceId(), getResourceInitials(),
 getResourceName(), getResourceUniqueID(), getResourceGroup(), getResourceAssignmentId()
```

