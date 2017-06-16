# Automatically open a task pane with a document

You can use add-in commands in your Office Add-in to extend the Office UI by adding buttons to the Office ribbon. When users click your command button, an action occurs, such as opening a task pane. Some scenarios require that a task pane open automatically when a document opens, without explicit user interaction. You can use the autoopen taskpane feature, introduced in the AddInCommands 1.1 requirement set, to automatically open a task pane when your scenario requires it. 


## How is the autoopen feature different from inserting a task pane? 

Add-ins that don't use add-in commands - for example, add-ins that run in Office 2013 - are inserted into the document by default, and persist in that document without explicit user or developer intent. As a result, when other users open the document, they are prompted to install the add-in, and the task pane opens. The challenge with this model is that in many cases, users don’t want the add-in to persist in the document. For example, a student who uses a dictionary add-in in a Word document might not want their classmates or teachers to be prompted to install that add-in when they open the document.  

With the autoopen feature, you can explicitly define or allow the user to define whether a specific task pane add-in persists in a specific document. 

## Support and availability
The autoopen feature is currently <!-- in **developer preview** and it is only --> supported in the following products and platforms.

|**Products**|**Platforms**|
|:-----------|:------------|
|<ul><li>Word</li><li>Excel</li><li>PowerPoint</li></ul>|<ul><li>Office for Windows Desktop. Build 16.0.811.1000+ (Insiders Fast)</li><li>Office for Mac. Build 15.34(170414)+ (Insiders Fast)</li><li>Office Online</li></ul>|


<!-- >**Note:** For Windows and Mac, you need to be on **[Insiders Fast](https://products.office.com/en-us/office-insider?tab=tab-1)** and have updates turned on to have access to this feature during the preview. The feature won't work if you are not part of Insiders Fast, even if you have a more recent build. -->

## Best practices for using the autoopen feature

- Use the autoopen feature when it will help make your add-in users more efficient. Consider the following scenarios:
	- The document needs the add-in in order to function properly. For example, a spreadsheet that includes stock values that are periodically refreshed by an add-in. The add-in should open automatically when the spreadsheet is opened to keep the values up to date. 
	- The user will most likely always use the add-in with a particular document. For example, an add-in that helps users fill in or change data inside a document by pulling information from a backend system. 
- Allow users to turn on or turn off the autoopen feature. Include an option in your UI for users to choose to no longer automatically open the add-in task pane.  
- Use requirement set detection to determine whether the autoopen feature is available, and provide a fallback behavior if it isn’t.
- Don't use the autoopen feature to artificially increase usage of your add-in. If it doesn’t make sense for your add-in to open automatically with certain documents, using this feature can annoy users. If Microsoft detects abuse of the autoopen feature,  your add-in might get rejected from the Office Store. 
- Don't use this feature as a way to generically pin panes of your add-in in place. The autoopen feature enables you to designate one pane of your add-in to open automatically with a document. If your add-in has multiple panes, you can only designate one to open automatically. 

## Implementation
To implement the autoopen feature, you:

- Specify the pane to be opened automatically.
- Tag documents to automatically open the task pane.


### Specify the pane to open
To specify the pane to open automatically, set a [TaskpaneId](https://dev.office.com/reference/add-ins/manifest/action#taskpaneid) value of ***Office.AutoShowTaskpaneWithDocument***. You can only set this value on one task pane. If you set this value on multiple task panes, the first occurrence of the value will be recognized and the others will be ignored. 
          
    <Action xsi:type="ShowTaskpane">
         <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
         <SourceLocation resid="Contoso.Taskpane.Url" />
    </Action>
     

#### 2.-Tagging a document to trigger auto-open
To trigger auto-open a document must be appropriately tagged. Documents that are not tagged will not trigger auto-open. You can tag a document in 2 main ways, choose the one that makes the most sense for your scenario:


##### Client side
Set the ***Office.AutoShowTaskpaneWithDocument*** setting to ***true*** using Office.js. Use this method if you need to tag the document as part of your add-in interaction (E.g. as soon as the user creates a binding, or clicks on a UI affordance on your add-in to indicate they want the pane to auto-open) 

    Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
    Office.context.document.settings.saveAsync();

##### Via Open XML
You can use Open XML to create or modify a document and add the appropriate Open Office XML markup that implements auto-open. A sample to show how to do this is at [Office-OOXML-EmbedAddin](https://github.com/OfficeDev/Office-OOXML-EmbedAddin). Here is some information you need to implement your own auto-open solution with Open XML.

There are two Open XML parts that need to be added to the document.

First, is a webextension part. The following is an example:

    <we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="[ADD-IN ID PER MANIFEST]">
      <we:reference id="[GUID or Office Store asset ID]" version="[your add-in version]" store="[Pointer to store or catalog]" storeType="[Store or catalog type]"/>
      <we:alternateReferences/>
      <we:properties>
    	<we:property name="Office.AutoShowTaskpaneWithDocument" value="true"/>
      </we:properties>
      <we:bindings/>
      <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
    </we:webextension>

The webextension part includes a property bag and a property named **Office.AutoShowTaskpaneWithDocument** that must be set to `true`.

The webextension part also includes a reference to the store or catalog with attributes for `id`, `storeType`, `store`, and `version`. There are seven store types, but only four of them are relevant to auto-opening add-ins. The values for the other three attributes depend on the value for `storeType` as shown in the following table. 

| When the `storeType` is: | The `id` should be:	|The `store` should be: | The `version` should be:|
|:---------------|:---------------|:---------------|:---------------|
|OMEX (The Office Store)|The Office Store asset ID of the add-in.\*|The locale of the Office Store; for example, "en-US".|The version in the Office Store catalog.\*|
|FileSystem (A network share)|The GUID of the add-in in the add-in manifest.|The path of the network share; for example, "\\\\MyComputer\\MySharedFolder".|The version in the add-in manifest.|
|EXCatalog (Centralized deployment from Exchange) |The GUID of the add-in in the add-in manifest.|"EXCatalog"|The version in the add-in manifest.
|Registry (System registry)|The GUID of the add-in in the add-in manifest.|"developer"|The version in the add-in manifest.|

>\* To find the asset ID and version of an add-in in the Office Store, navigate to the add-in's page in the Office Store. The asset ID will be in the address bar of the browser. The version will be in the **Details** section of the page.

For more information about the webextension markup, see [[MS-OWEXML] 2.2.5. WebExtensionReference](https://msdn.microsoft.com/en-us/library/hh695383(v=office.12).aspx).

The second Open XML part that is related to auto-open is a taskpane part. The following is an example.

    <wetp:taskpane dockstate="right" visibility="0" width="350" row="4" xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">
      <wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" />
    </wetp:taskpane>

The critical attribute here is `visibility` which in this example is set to "0". This means that the first time the document is opened after the two parts have been added, the user has to install the add-in, from the **Add-in** button on the ribbon. On every *subsequent* opening of the file, the taskpane with the add-in opens automatically. Also, setting `visibility` to "0" means that you can use the Office.js to give users the ability to turn on and off the auto-opening of the add-in. Specifically, your script sets the **Office.AutoShowTaskpaneWithDocument** document setting to `true` or `false`. (See above in the **Client side** section.) 

If `visibility` is set to "1", then the task pane opens automatically the very first time the document is opened after the two parts have been added. A prompt to trust the add-in appears in the task pane, and when the user grants trust, the add-in opens. On every *subsequent* opening of the file, the taskpane with the add-in opens automatically. However, when `visibility` is set to "1", there is no way for Office.js or, hence, your users to turn off the auto-opening of the add-in.  

Setting `visibility` to "1" could be a good choice when the add-in and the document's template or content are so closely tied that there's no scenario in which users would not want the add-in to open when the document is opened.

An easy way to figure out the XML you need to write is to first run your add-in and use the client side technique to write the value then save the document and inspect the XML that is generated. Office will detect the store or catalog of the add-in and fill the appropriate attribute values. There is even a tool that will generate the C# code to programmatically add the markup based on the example you produce with the client side technique: [Open XML SDK 2.5 Productivity Tool](https://www.microsoft.com/en-us/download/details.aspx?id=30425).


### Add-in installation requirement
It is important to highlight that the **pane that you designate will only automatically open IF** , by the time the user opens the document, your **add-in is already installed on the users device**.  If users open a document and they do not have your add-in already installed then nothing will happen, the setting will be ignored. 

If you require to also distribute the add-in with the document, so that users are prompted to install it, you also need to set the pane visibility property to 1, you can only do this via OpenXML.

## Samples
The folder in this repo contains a simple example that shows you how to specify what pane to open on your add-in manifest as well as how to tag a document via Office.js. Additional samples are in the works. 

![](http://i.imgur.com/JtHwr47.png)
