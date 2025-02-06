---
title: "Release Notes for Beta Channel"
ms.author: timda
author: TimDavenport
manager: TimDavenport
ms.audience: Win32 Fast
ms.topic: release-notes
ms.service: o365-proplus-itpro
ms.localizationpriority: medium
ms.collection: RelNotes_ProPlus
description: "Provides Insider Fast audience with the latest list of new features, fixes or known issues."
ms.date: 02/03/2025
---

# Release Notes for Beta Channel

This article contains release notes for Beta Channel builds of Word, Excel, PowerPoint, Outlook, Access, and Project for Windows desktop. Every week, we highlight interesting new features, important fixes, and any significant issues we want you to know about.

>[!Note]
>
> - We often roll out features to Beta Channel over a period of time. This allows us to ensure that things are working smoothly before releasing the feature to a wider audience. If you don't see something described, you'll get it eventually.

[//]: # (DO NOT REMOVE)

## Version 2502: February 03
*Version 2502 (Build 18526.20024)*


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates
### Excel

- **Reduce eye strain with Dark Mode:** Go to View and select Switch Modes to try Dark Mode, which darkens your entire sheet, including cells.


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

<br/>

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### PowerPoint

- We have decreased the opacity for watermark content markings inserted into presentations as a result of sensitivity labeling to match the opacity of watermark content markings in Word.



[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2502: January 21
*Version 2502 (Build 18516.20000)*

## Version 2502: January 17
*Version 2502 (Build 18514.20000)*


[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### PowerPoint

- We fixed an issue where opening and closing multiple embedded Excel objects from a PowerPoint file throws an OLEFormat.DoVerb: invalid request exception

- We fixed an issue where PowerPoint automatically closes when the system goes into hibernate or sleep mode.

### Word

- Resolved an issue in Word where accepting the security prompt for legacy text forms may clear it to the default text.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2502: January 07
*Version 2502 (Build 18502.20000)*


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates
### Word

- **Removed preference to open files in Reading View:** The option "Open email attachments and other uneditable files in reading view" has been removed. The experience now defaults to "Viewing Mode".


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

<br/>

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Word

- Resolved an issue where the Shape.AlternativeText property may return the alternative text for the group shape instead of the individual shapes.



[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2501: January 01
*Version 2501 (Build 18429.20000)*

## Version 2501: December 25
*Version 2501 (Build 18422.20000)*


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Word

- Resolved an issue where objects didn't display as expected when using the "hide all" option in the selection pane and using undo/redo.

- Resolved an issue where a user didn’t receive an error message when the save operation was unsuccessful on a Distributed File System.

### Office Suite

- Fixed a problem where PowerPoint would sometimes close unexpectedly when reading the slide content with a screen reader, particularly when navigating by heading in Slide Show on a slide with hyperlinks.

- Fixed a problem where the legend of a map chart is missing in print, print preview and export-as-PDF.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2501: December 10
*Version 2501 (Build 18407.20002)*

Security updates listed [here](microsoft365-apps-security-updates.md)


[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### PowerPoint

- The Title field in File > Info will no longer be filled in automatically when saving a file.



[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2501: December 03
*Version 2501 (Build 18330.20000)*

## Version 2412: November 27
*Version 2412 (Build 18324.20012)*


[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### OneNote

- We fixed an issue that was causing pages to have a blank title when sending a meeting request from Outlook to OneNote.

### Word

- Resolved an issue in Word where users received the message "The function you are trying to run contains macros" when opening documents with VBA disabled and web add-ins installed.

- Resolved an issue in Word where users may receive an HTTP error when calling Content Control JavaScript APIs.


[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2412: November 26
*Version 2412 (Build 18318.20000)*


[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Word

- Resolved an issue in Word where footnote numbering unexpectedly changed to the number 8 when using the "‘Restart each page" option.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2412: November 19
*Version 2412 (Build 18314.20000)*


[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Office Suite

- We fixed an issue where the application displayed empty or incorrect contents during scrolling or switching sheets in right-to-left (RTL) Excel.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2412: November 13
*Version 2412 (Build 18311.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Word

- Resolved an issue in Outlook where content in the preview pane was truncated.

- Resolved an issue where watermarks in a document weren't being printed.

- Resolved an issue where Styles unexpectedly change when coauthoring in a document based on a template.

### Office Suite

- We fixed an issue where the application wouldn't render the grid properly after switching from page break preview to normal view.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2412: November 07
*Version 2412 (Build 18305.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Modernized user-defined permissions experience:** We’ve improved the user experience for selecting which users should have which permissions when a sensitivity label configured for user-defined permissions is applied to a file or when configuring standalone Information Rights Management through the Restrict Access feature.

### PowerPoint

- **Modernized user-defined permissions experience:** We’ve improved the user experience for selecting which users should have which permissions when a sensitivity label configured for user-defined permissions is applied to a file or when configuring standalone Information Rights Management through the Restrict Access feature.

### Word

- **Modernized user-defined permissions experience:** We’ve improved the user experience for selecting which users should have which permissions when a sensitivity label configured for user-defined permissions is applied to a file or when configuring standalone Information Rights Management through the Restrict Access feature.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Word

- Resolved an issue where a document saved to a network shared folder and set to "Always Open Read-Only" would open in "Editing" mode.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2411: October 29
*Version 2411 (Build 18227.20000)*


[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### OneNote

- Fixed an issue in the application that disabled scrolling with a mouse scroll wheel.

- Fixed an issue that caused the application to repeatedly ask for login credentials.

### PowerPoint

- Fixed an issue in PowerPoint where embedded PowerPoint Presentations were inheriting the default label.

### Office Suite

- Fixed an issue that caused PowerPoint to stop responding while animating certain GIF images.

- We fixed an issue where some PowerPoint presentations created by third-party tools wouldn't open correctly and some content were removed.


[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2411: October 22
*Version 2411 (Build 18217.20000)*


[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### PowerPoint

- We have fixed an issue where embedded BMP images in the PowerPoint
slide were not opening.


### Word

- Resolved an issue where Word would give the error "Word could not re-establish a DDE connection to Microsoft Query to complete the current task" when selecting MS Query as the data source for “Mail Merge Create Data Source”.



[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2411: October 14
*Version 2411 (Build 18210.20000)*


[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Word

- Fixed an issue that was preventing emails with linked SVG content from saving or sending.


- Resolved an issue where opening a large document with many revisions would take an extended period of time.


- Resolved an issue where RTF files would not be proofed on open, resulting in a 0% rating from the Editor pane.


- Resolved an issue where updating a table of figures may cause unexpected changes to certain words, resulting in spelling mistakes.


- Resolved an issue where the Save button would be unresponsive if the file was saved to a folder including with a period (.) in the path name.


### Office Suite

- We fixed an issue where users experienced slowness when an Excel file contained many combo box form controls.



[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2411: October 08
*Version 2411 (Build 18201.20000)*

Security updates listed [here](microsoft365-apps-security-updates.md)

## Version 2410: October 08
*Version 2410 (Build 18129.20012)*

Security updates listed [here](microsoft365-apps-security-updates.md)

## Version 2410: October 04
*Version 2410 (Build 18129.20010)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Project

- Fixed an issue where signed macros would not run even if the Trust Center Macro Settings "Disable all macros except digitally signed macros" option was selected.

### Word

- Fixed an issue in Word where images would move when moused over.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2410: October 01
*Version 2410 (Build 18122.20000)*

## Version 2410: September 24
*Version 2410 (Build 18120.20002)*

## Version 2410: September 17
*Version 2410 (Build 18112.20000)*

## Version 2410: September 10
*Version 2410 (Build 18108.20000)*

Security updates listed [here](microsoft365-apps-security-updates.md)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### PowerPoint

- Fixed an issue in PowerPoint where saves to a OneDrive sync location would sometimes fail unexpectedly for users subject to a mandatory sensitivity labeling policy.

- We fixed an issue where PowerPoint (.pptx) files exported as .pdf truncated or omitted images.

### Word

- Resolved an issue where placeholders in an equation would show when printing or exporting to PDF.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2410: September 03
*Version 2410 (Build 18028.20004)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

## Version 2409: August 28

*Version 2409 (Build 18025.20006)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Information Protection: Improved DLP policy tip reliability:** Word, Excel, and PowerPoint for Windows now consistently show the same policy tips that SharePoint Online and OneDrive for Business sites have determined should be shown on a file.

### PowerPoint

- **Information Protection: Improved DLP policy tip reliability:** Word, Excel, and PowerPoint for Windows now consistently show the same policy tips that SharePoint Online and OneDrive for Business sites have determined should be shown on a file.

### Word

- **Information Protection: Improved DLP policy tip reliability:** Word, Excel, and PowerPoint for Windows now consistently show the same policy tips that SharePoint Online and OneDrive for Business sites have determined should be shown on a file.

### Office Suite

- **Protect sensitive files with dynamic watermarking:** Deter leakage of your highly confidential document content by rendering the reader’s email address and other information over the document. This is configurable on sensitivity labels.

- **Open locked records as read-only:** Files with retention labels marking them as locked records will now open as read-only to prevent user edits.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### PowerPoint

- The Title field in File > Info will no longer be filled in automatically when saving a file.

### Word

- Resolved an issue when modifying styles that the "Style based on" entry would be blank.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2409: August 20
*Version 2409 (Build 18015.20000)*

## Version 2409: August 13
*Version 2409 (Build 18011.20000)*

Security updates listed [here](microsoft365-apps-security-updates.md)

## Version 2409: August 08

*Version 2409 (Build 18006.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where the user would experience changes in column widths unexpectedly when using PivotTables.

- We fixed issues users could encounter applying conditional formatting via the Top10 Interop object in Add-in, which included requiring a restart and potential data loss.

### Word

- Resolved an issue where "Restricted Editing" may not prevent editing after uploading a document to SharePoint and reopening.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2408: July 31

*Version 2408 (Build 17928.20004)*


[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Office Suite

- We fixed an issue where the application stopped rendering shapes while switching sheets in Excel.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2408: July 30
*Version 2408 (Build 17925.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Word

- Fixed an issue to prevent opening OneDrive sync files exceeding 259 characters, which could result in unpredictable outcomes.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2408: July 23
*Version 2408 (Build 17920.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates
### Excel

- **Publish to Power BI retires starting August 19, 2024:** Publish to Power BI retires from Excel in Microsoft 365 starting August 19, 2024. We recommend users transition to Microsoft Power BI on the web (also known as the Power BI service), for their publishing needs.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

<br/>

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Word

- Resolved an issue where the ruler wouldn't show unit numbers or indents in right-to-left (RTL) layout.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2408: July 16
*Version 2408 (Build 17911.20000)*


[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Word

- Resolved an issue where preferences deployed with the Office Deployment Tool were not being set correctly.

### Office Suite

- We fixed an issue where the application displayed empty or incorrect contents during scrolling or switching sheets in right-to-left (RTL) Excel.

- Resolved a performance issue in PowerPoint when typing with autocorrect.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2407: July 08
*Version 2407 (Build 17830.20016)*


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Information Protection: Improved DLP policy tip reliability:** Word, Excel, and PowerPoint for Windows now consistently show the same policy tips that SharePoint Online and OneDrive for Business sites have determined should be shown on a file.

### PowerPoint

- **Information Protection: Improved DLP policy tip reliability:** Word, Excel, and PowerPoint for Windows now consistently show the same policy tips that SharePoint Online and OneDrive for Business sites have determined should be shown on a file.

### Word

- **Information Protection: Improved DLP policy tip reliability:** Word, Excel, and PowerPoint for Windows now consistently show the same policy tips that SharePoint Online and OneDrive for Business sites have determined should be shown on a file.

### Office Suite

- **Use sensitivity labels in Microsoft 365 apps when connected experiences are disabled:** Your organization can now disable connected experiences for privacy concerns without impacting data security policies, such as sensitivity labels.  Services associated with Microsoft Purview (e.g. sensitivity labels, rights management, ...) are no longer controlled by policy settings to manage privacy controls for Microsoft 365 Apps. Instead, these services will rely on their existing security admin controls in Purview portals.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues
### Excel

- We fixed an issue where coauthoring on text boxes could yield unexpected results.

- We fixed an issue where linked pictures could fail to update.

- We fixed an issue where the ChartArea.ClearContents method might unexpectedly fail.

### Word

- We fixed an issue in which certain customers who have installed 3rd party fonts on their machines could see garbled text with WordArt effects applied to paragraph text in Word.

- Resolved an issue where revisions may be skipped when reviewing using VBA.

- Resolved an issue where characters could appear corrupted when opening RTF emails.

- Resolved an issue where the Word process may continue running for users with the GENKO add-in after Word is called using OLE.

- Resolved an issue where Toolbar and Command bar customizations could become corrupted in a document.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2407: June 25
*Version 2407 (Build 17820.20000)*

## Version 2407: June 18

*Version 2407 (Build 17816.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Word

- Resolved as issue where the Actions pane was unavailable in Office 2016.

- Resolved an issue where the prompt given to Copilot is lost if the user encountered an error message. The prompt is now saved after clicking 'Generate'.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2407: June 11

*Version 2407 (Build 17809.20000)*

Security updates listed [here](microsoft365-apps-security-updates.md)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Word

- Resolved an issue where users were unable to launch Word with the DDE function (/x switch).

### Office Suite

- Resolved a performance issue when typing with autocorrect on in PowerPoint.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2407: June 04

*Version 2407 (Build 17730.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Prevent connected services that analyze content with sensitivity labels in files and emails:** You can configure a new setting for sensitivity labels that prevents content in Word, Excel, PowerPoint, and Outlook from being sent to Microsoft for content analysis as a privacy control.

- **Open locked records as read-only:** Files with retention labels marking them as locked records now open as read-only to prevent user edits.

### PowerPoint

- **Open locked records as read-only:** Files with retention labels marking them as locked records now open as read-only to prevent user edits.

- **Prevent connected services that analyze content with sensitivity labels in files and emails:** You can configure a new setting for sensitivity labels that prevents content in Word, Excel, PowerPoint, and Outlook from being sent to Microsoft for content analysis as a privacy control.

### Outlook

- **Prevent connected services that analyze content with sensitivity labels in files and emails:** You can configure a new setting for sensitivity labels that prevents content in Word, Excel, PowerPoint, and Outlook from being sent to Microsoft for content analysis as a privacy control.

- **Tables are more accessible for screen readers:** Screen readers should more sensibly read out information in tables.

### Word

- **Prevent connected services that analyze content with sensitivity labels in files and emails:** You can configure a new setting for sensitivity labels that prevents content in Word, Excel, PowerPoint, and Outlook from being sent to Microsoft for content analysis as a privacy control.

- **Open locked records as read-only:** Files with retention labels marking them as locked records now open as read-only to prevent user edits.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

## Version 2406: May 29

*Version 2406 (Build 17726.20006)*

## Version 2406: May 28

*Version 2406 (Build 17723.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

## Version 2406: May 21

*Version 2406 (Build 17716.20002)*

## Version 2406: May 17

*Version 2406 (Build 17715.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **PDF tagging updates on Excel for Windows:** Updating and reducing the accessibility tags added when a workbook is converted to a PDF.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

<br/>

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Outlook

- We fixed an issue that caused Outlook to exit unexpectedly when using Copilot Summarize.

### Word

- Resolved an issue in Word where embedded objects couldn't be opened if the file needed to be checked out before editing.

- Resolved an issue where embedded objects couldn't be opened after signing a document and marking it as final.

- Resolved an issue in Outlook where background images wouldn't render correctly, requiring the user to scroll or click into the message for the image to appear.

- Resolved an issue where multiple newlines were appearing between paragraphs in multi-line comments while coauthoring.

- Resolved an issue where the comment anchor range would still show after the comment was resolved.

- Resolved an issue where AutoRecovered documents would be saved in compatibility mode.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2406: May 07

*Version 2406 (Build 17630.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)


[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

## Version 2405: May 01

*Version 2405 (Build 17628.20006)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Outlook

- We fixed an issue that caused users to be unable to open some search results when searching their Online Archive.

- We fixed an issue that caused updates made by delegates to the "Show As" setting on a manager's calendar fail to be persisted after clicking "Send Update."

### PowerPoint

- The Title field in File > Info will no longer be filled in automatically when saving a file.

### Word

- Resolved an issue in Outlook where background images wouldn't render correctly, requiring the user to scroll or click into the message for the image to appear.

- Resolved an issue where some objects wouldn't be selected when using the "Select Drawing Objects" tool.

- Resolved an issue where attempting to reference a file using Draft with Copilot would give the error "Copilot wasn't able to apply the right sensitivity label to this file because it couldn't work with the content you referenced."

- Resolved an issue that caused footnote numbers to be missing from the Cross Reference list when track changes is turned on in a document and changes are made.

- Resolved an issue where the Enter and Backspace keys weren't performing their expected functions within a list, preventing users from exiting the list.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2405: April 25

*Version 2405 (Build 17616.20002)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### PowerPoint

- Improved experience in changing the source of a linked chart:
In a situation where a presentation contains a linked chart, which is not only stored at a SharePoint location, but also synced to a local folder, when the user tries to change the chart source via "File -> Info -> Edit Links to File -> Change Source...", the initial location in the Change Source dialog now shows the current location of the linked chart, instead of the last used location. PowerPoint can now convert the local sync folder location to its corresponding SharePoint URL and preserve it in the file if the user chooses a chart in a local sync folder as the new chart source.

### Word

- Resolved an issue where text inside the track change card wasn't properly visible when using dark grey theme.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2405: April 10

*Version 2405 (Build 17602.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### PowerPoint

- The Edit Photo Album option is now enabled in Slide Sorter view.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2404: April 02

*Version 2404 (Build 17531.20000)*

## Version 2404: March 26

*Version 2404 (Build 17521.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2404: March 19

*Version 2404 (Build 17514.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2404: March 14

*Version 2404 (Build 17512.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2404: March 13

*Version 2404 (Build 17509.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where the app closed unexpectedly when the user would attempt to open or import large text files.

- We fixed an issue where an Excel file was no longer able to be saved due to file corruption with conditional formatting.

- We fixed an issue where the app closed unexpectedly when opening files with pivot tables and table design in macro-enabled files.

- With multiple pane scenarios, make sure all the visible panes get drawn for scrolling operation.

- We fixed an issue where worksheets with right-to-left alignment and frozen panes would sometimes render incorrectly.

- We fixed an issue where attempting to programmatically set the Formula property on a Linked Picture resulted in a run-time error.

- We fixed an issue that caused errors when attempting to save files in environments where Default Labeling and Mandatory Labeling policies are configured and AIP document co-authoring is disabled.

- We fixed the issue of a shape location shifting when pasting a range with some shapes as a picture.

- We fixed an issue in Excel where the user was unable to set the default font.

- We fixed an issue where the app closed unexpectedly when a user switched Excel from an account where preview was started or enabled to an account where no preview had been started or enabled.

### Outlook

- We fixed an issue where users were unable to recall a message sent in Outlook.

- We fixed an issue where custom forms from MAPI form servers stopped responding.

### PowerPoint

- We fixed an issue in the PowerPoint presenter view where speaker notes from a previous slide would move to the current slide view if no speaker notes were available.

### Word

- We fixed an issue where certain Word documents would not open in compatibility mode.

- We fixed an issue where the accessibility pane was blank on saving the document in .doc format or when document was in compatibility mode.

- We fixed an issue where a notification wasn't disabling as expected in template (.dotx) files with personal information removal enabled in the Change Setting or the Trust Center Settings.

- We fixed an issue where the Word template (normal.dotm) was overwritten upon exit when the proofing language was set to Japanese.

- We fixed an issue where styles changed or were applied incorrectly when transitioning to WPM.

- We fixed an issue where opening certain Word documents would give error, "Word experienced an error trying to open the file."

### Office Suite

- We fixed an issue where the Office update installer would appear to be unresponsive.

- We fixed a performance issue when applying criteria to the ID field in a query to a linked ODBC table.

- We fixed an issues where Access didn't run digitally-signed macros.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2404: March 06

*Version 2404 (Build 17503.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed a rendering issue related to scrolling or zooming in charts.

### Office Suite

- We fixed an issue where the user wasn't prompted to save to the cloud on exit from Excel.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2403: February 28

*Version 2403 (Build 17425.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### PowerPoint

- We fixed an issue where PowerPoint closed unexpectedly when a user tried to open a copy of a presentation saved to the cloud.

- We fixed an issue where PowerPoint (.pptx) files exported as .pdf truncated or omitted images.

### Word

- We fixed an issue where completing Proofing with an Office Add-in's side pane open would result in the add-in's pane being closed.

### Office Suite

- We fixed an issue where the app wouldn't re-install when the device was offline, and the re-install was from a local install source.

- We fixed an issue where a mandatory sensitivity label wasn't applied or prompted for if an email was sent from the Drafts-folder.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2403: February 21

*Version 2403 (Build 17419.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue that caused errors when attempting to save files in environments where Default Labeling and Mandatory Labeling policies are configured and AIP document co-authoring is disabled.

- We fixed an issue where the input language of a keyboard converted from Chinese Traditional to English, an error appeared, and the app closed unexpectedly when the name of an Excel sheet was edited.

### Outlook

- We fixed an issue that affected all-day meetings with a reminder on the day of. The meeting didn't open, and an error message was displayed.

### Word

- We fixed an issue where characters became invisible when converted to Kanji.

- We fixed an issue where certain Word documents wouldn't open in compatibility mode.

- We fixed an issue where styles changed or were applied incorrectly when transitioning to WPM.

### Office Suite

- We fixed an issue where opened Office apps didn't automatically start when a laptop was reopened, and an error message appeared after manual relaunch.

- We fixed an issue where the user wasn'tified of an Alt-text violation after adding and deleting an image.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2403: February 13

*Version 2403 (Build 17404.20000)*

Security updates listed [here](microsoft365-apps-security-updates.md).

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### OneNote

- **Preview of New Sticky Notes app on Windows available from OneNote app:** The new Sticky Notes app allows you to create and recall notes like the classic Microsoft Sticky Notes app, while making it easier to capture screenshots, automatically add the source information to the notes, and more.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where attempting to programmatically set the Formula property on a Linked Picture resulted in a run-time error.

### OneNote

- We fixed an issue where, after switching from an ODM notebook to an ODP notebook with the notebooks pane pinned, the Math Assistant continued to appear in the Draw tab and prompted an error.

### Outlook

- We fixed an issue with user and shared mailboxes where a message sent from a user with saved to Sent Items folder enabled and where DelegateSentItemsStyle is set to 1 is saved in the shared mailbox Sent Items folder.

- We fixed an issue that caused expanded groups in the message list to become collapsed when changing which column they're arranged by.

- We fixed an issue where a Meeting window wouldn't close and the meeting wouldn't get added to the organizer's calendar if an image had previously been added into the body of the message and the organizer was attempting to reopen the meeting and add new attendees.

### Word

- We fixed an issue where the default font Yu Mincho Demibold would change to Yu Gothic when the user moved to the next line.

- We fixed an issue where pasting a list with a custom style into a document that was saved to the cloud with restricted formatting would show an "AutoSave Failed" error.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2403: February 06

*Version 2403 (Build 17330.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Outlook

- We fixed an issue with user and shared mailboxes where a message sent from a user with saved to Sent Items folder enabled and where DelegateSentItemsStyle is set to 1 is saved in the shared mailbox Sent Items folder.

- We fixed an issue that caused expanded groups in the message list to become collapsed when changing which column they're arranged by.

### Word

- We fixed an issue where the default font Yu Mincho Demibold would change to Yu Gothic when the user moved to the next line.

- We fixed an issue where pasting a list with a custom style into a document that was saved to the cloud with restricted formatting would show an "AutoSave Failed" error.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2402: January 31

*Version 2402 (Build 17328.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Word

- We fixed an issue where the default font Yu Mincho Demibold would change to Yu Gothic when the user moved to the next line.

- We fixed an issue where the error message “Compile error” would appear when opening a .docm file.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2402: January 23

*Version 2402 (Build 17318.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### PowerPoint

- We fixed an issue where users weren't able to create a bullet list with hyphens.

### Word

- We fixed an issue where the font would switch from Aptos to Times New Roman when typing in a new email.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2402: January 17

*Version 2402 (Build 17315.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Word

- We fixed an issue where the wrong paper size was chosen when two paper sizes in a printer were close to 1/6th of an inch of each other.

- We fixed an issue where only one document would open when trying to open several files containing SQL data from File Explorer or the desktop.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2402: January 09

*Version 2402 (Build 17304.20000)*

Security updates listed [here](microsoft365-apps-security-updates.md).

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2401: January 03

*Version 2401 (Build 17231.20008)*

## Version 2401: January 02

*Version 2401 (Build 17228.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### PowerPoint

- We fixed an issue where Visual Basic for Applications (VBA) RecordNarrationFromCurrentSlide failed with a runtime error.

- We fixed an issue where the 'Cell Margins' in the Table Layout ribbon item was disabled for users with only Edit Information Rights Management (IRM) permissions.

- We fixed an issue where the Add comments method in Visual Basic for Applications (VBA) showed a runtime error.

- We fixed an issue where the Notes and Slide layout would open with incorrect proportions when a file was opened from a protected view.

### Word

- We fixed an issue where comment cards may appear too wide and cut off text when changing or switching the screen in use.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2401: December 20

*Version 2401 (Build 17217.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Word

- We fixed an issue where comment cards may appear too wide and cut off text when changing or switching the screen in use.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2401: December 12

*Version 2401 (Build 17204.20000)*

Security updates listed [here](microsoft365-apps-security-updates.md).

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Sensitivity toolbar is now available when creating a copy of a document:** The sensitivity bar is now available in Word, Excel, and PowerPoint when users are creating copies of their documents in File / Save As. This helps users understand the security policies that apply to their document.

### PowerPoint

- **Sensitivity toolbar is now available when creating a copy of a document:** The sensitivity bar is now available in Word, Excel, and PowerPoint when users are creating copies of their documents in File / Save As. This helps users understand the security policies that apply to their document.

### Word

- **Sensitivity toolbar is now available when creating a copy of a document:** The sensitivity bar is now available in Word, Excel, and PowerPoint when users are creating copies of their documents in File / Save As. This helps users understand the security policies that apply to their document.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

## Version 2401: December 05

*Version 2401 (Build 17130.20000)*

## Version 2312: November 30

*Version 2312 (Build 17126.20000)*

## Version 2312: November 28

*Version 2312 (Build 17123.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Word

- We fixed an issue where the input language formatting would sometimes revert to a previously typed language when using Language AutoDetect.

### Office Suite

- We fixed an issue where the orientation of Scalable Vector Graphics (SVG) line markers wasn't rendering correctly.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2312: November 21

*Version 2312 (Build 17116.20002)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### PowerPoint

- We fixed an issue where restoring a previous version of a presentation wasn't working as expected when using Version History.

### Project

- We fixed an issue where the duration on manually scheduled tasks wasn't calculated correctly when importing from XML.

### Word

- We fixed an issue where the content control end tag is marked at the end of the document automatically if the document is edited in Word Online and then opened in Word desktop.

- We fixed an issue where timestamps weren't showing correctly or were missing in comments.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2312: November 14

*Version 2312 (Build 17108.20000)*

Security updates listed [here](microsoft365-apps-security-updates.md).

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Outlook

- **Outlook for Windows: Actionable card based Loop components:** Add an app to Microsoft 365 that supports link unfurling with an Actionable Card based Loop component.  Outlook will show the component when inserting supported links in your email messages. You are able to choose between the new embedded card or a link when sending email.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where the font wasn't being retained in the dropdown list when changing the text to a new font.

- We fixed an issue where certain Pivot Tables would load slowly.

### PowerPoint

- We fixed an issue where running the Visual Basic for Applications (VBA) macro View.Slide with multiple slides selected, returned the wrong Slide object.

### Office Suite

- We fixed an issue where an out-of-memory error would sometimes appear when saving a PowerPoint presentation using the 32-bit version of Office.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2312: November 07

*Version 2312 (Build 17031.20000)*

## Version 2311: October 31

*Version 2311 (Build 17029.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### OneNote

- We fixed an issue where the calendar date picker day (Sat., Sun., Mon., etc.) didn't align with the actual day number of the month.

### PowerPoint

- We fixed an issue where some Connected Service account users were prevented from adding comments.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2311:  October 31

*Version 2311 (Build 17023.20000)*

## Version 2311: October 25

*Version 2311 (Build 17022.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Reduce workbook size bloat from unwanted cell formatting with "Check Performance":** Reduce workbook size bloat from unwanted cell formatting with new "Check Performance" task pane user experience.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where certain Pivot Tables would load slowly.

### Word

- We fixed an issue where creating a new comment with the Comment Pane enabled would reset the selected comment to the top of the pane, causing the document to scroll to that comment's anchor.

- We fixed an issue where running Word's Spell Check through a Visual Studio solution would result in the 'Not in dictionary' area being grey and not editable when the Word window wasn't visible.

- We fixed an issue where the content control end tag is marked at the end of the document automatically if the document is edited in Word Online and then opened in Word desktop.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2311: October 17

*Version 2311 (Build 17010.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### OneNote

- We fixed an issue where the calendar date picker day (Sat., Sun., Mon., etc.) didn't align with the actual day number of the month.

### Word

- We made performance improvements for Macros where runtimes were slower in current builds compared to older versions.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2311: October 10

*Version 2311 (Build 17005.20000)*

Security updates listed [here](microsoft365-apps-security-updates.md).

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Outlook

- We fixed an issue where Outlook would close unexpectedly when a user clicked “Cancel” in the Wait on Send dialog if the message had attachments.

### Word

- We fixed an issue where revised content was reversed showing added content as deleted and deleted content as added.

### Office Suite

- We fixed an issue that caused unexpected black borders to appear around screen captures added with the Insert Screenshot functionality.

- We fixed an issue where an out-of-memory error would sometimes appear when saving a PowerPoint presentation using the 32-bit version of Office.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2311: October 03

*Version 2311 (Build 16926.20000)*

## Version 2310: September 28

*Version 2310 (Build 16924.20002)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Efficient reading in Excel using Narrator:** Announcements through the Windows inbox screen reader, Narrator, are much more succinct and efficient.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Word

- We fixed an issue where previously inserted custom bullets would reset after editing a document.

- We fixed an issue where if you saved a PDF in a language other than English and it had a hyperlink with a dash (-), the URL would get an extra dash that made the link not work.

- We fixed an issue where the layout of a table would unexpectedly change after saving and re-opening the file when the table was modified with Track Changes enabled.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2310: September 26

*Version 2310 (Build 16921.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

## Version 2310: September 22

*Version 2310 (Build 16918.20002)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### PowerPoint

- We fixed an issue where ActiveX controls do not display correctly after changing the display scale.

- We fixed an issue where the 'End with Black Slide' option wasn't working when slides advanced on a timer.

### Word

- We fixed an issue where deleting text within a Content Control would unexpectedly delete the Content Control when Track Changes was enabled.

- We fixed an issue where the content control end tag is marked at the end of the document automatically if the document is edited in Word Online and then opened in Word desktop.

- We fixed an issue where subheading numbering with a custom Style would disappear if the file was saved and re-opened.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2310: September 12

*Version 2310 (Build 16907.20000)*

Security updates listed [here](microsoft365-apps-security-updates.md).

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### PowerPoint

- We fixed an issue where clicking on Updates Available in the status bar doesn't merge in changes.

### Project

- We fixed an issue where a message that says, "You cannot delete or modify a project summary task" appears when opening a project.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2310: September 05

*Version 2310 (Build 16830.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

## Version 2309: September 05

*Version 2309 (Build 16827.20004)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

## Version 2309: August 29

*Version 2309 (Build 16827.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2309:  August 29

*Version 2309 (Build 16824.20002)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Word

- We fixed an issue where the Compare button was grayed out when a document was checked in and opened in desktop.

- We fixed an issue where a document with "view only" permissions was able to be saved.

- We fixed an issue where the error message "You are attempting to open a file type that has been blocked by your File Block settings in the Trust Center." would appear when trying to open a file type marked for opening in Protected View, when Application Guard was enabled for Office and the setting "Open selected file types in Protected View and allow editing" was enabled.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2309: August 22

*Version 2309 (Build 16818.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Outlook

- **Teams Meeting Apps now work in Outlook too:** Now, you can configure a meeting app while scheduling an invite in Outlook. This meeting app is ready to use when you chat or join the meeting on Teams.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Outlook

- We fixed an issue that caused the incorrect workdays to be displayed on shared-in calendars.

### Word

- We fixed an issue where base64 encoded .png images displayed the incorrect background color after switching folders.

- We fixed an issue where the header of a document would disappear when enforcing Microsoft Information Protection (MIP) labels if the document was printed before selecting the label.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2309: August 15

*Version 2309 (Build 16811.20004)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Office has a new default theme:** We've updated the default font and colors of the Office theme to be more modern and accessible.

### Outlook

- **Office has a new default theme:** We've updated the default font and colors of the Office theme to be more modern and accessible.

### PowerPoint

- **Office has a new default theme:** We've updated the default font and colors of the Office theme to be more modern and accessible.

### Word

- **Office has a new default theme:** We've updated the default font and colors of the Office theme to be more modern and accessible.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Office Suite

- We fixed an issue where the File > Open dialog wouldn't open when selecting Save As Picture to save an image.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2309: August 09

*Version 2309 (Build 16803.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Updated encryption for Microsoft Purview Information Protection:** Advanced Encryption Standard (AES) with 256-bit key length in Cipher Block Chaining mode (AES256-CBC) is now the default Microsoft Purview Information Protection encryption mechanism for Microsoft 365 Apps documents and emails. [Learn more](/microsoft-365/compliance/technical-reference-details-about-encryption#aes256-cbc-support-for-microsoft-365).

- **Track tasks with \@mentions:** Use @mentions in comments to create, assign and track tasks within your workbook. [Learn more](https://support.office.com/article/644bf689-31a0-4977-a4fb-afe01820c1fd)

### Outlook

- **Updated encryption for Microsoft Purview Information Protection:** Advanced Encryption Standard (AES) with 256-bit key length in Cipher Block Chaining mode (AES256-CBC) is now the default Microsoft Purview Information Protection encryption mechanism for Microsoft 365 Apps documents and emails. [Learn more](/microsoft-365/compliance/technical-reference-details-about-encryption#aes256-cbc-support-for-microsoft-365).

### PowerPoint

- **Updated encryption for Microsoft Purview Information Protection:** Advanced Encryption Standard (AES) with 256-bit key length in Cipher Block Chaining mode (AES256-CBC) is now the default Microsoft Purview Information Protection encryption mechanism for Microsoft 365 Apps documents and emails. [Learn more](/microsoft-365/compliance/technical-reference-details-about-encryption#aes256-cbc-support-for-microsoft-365).

### Word

- **Updated encryption for Microsoft Purview Information Protection:** Advanced Encryption Standard (AES) with 256-bit key length in Cipher Block Chaining mode (AES256-CBC) is now the default Microsoft Purview Information Protection encryption mechanism for Microsoft 365 Apps documents and emails. [Learn more](/microsoft-365/compliance/technical-reference-details-about-encryption#aes256-cbc-support-for-microsoft-365).

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Word

- We fixed an issue where updating Object Linking & Embedding (OLE) links would display the error message "Error! Not a valid link."

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2308: August 01

*Version 2308 (Build 16731.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### PowerPoint

- We fixed an issue where saving changes from a conflict dialog could result in duplicated slides.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2308: July 26

*Version 2308 (Build 16721.20002)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Performance improvement related to fonts:** If you do not use printer fonts or have not heard of printer fonts, unchecking the setting in File/Options/Advanced/[Include fonts that are stored on the printer] can help speed up font related operations such as choosing a font from font drop down, or formatting a cell by bolding/italicizing a cell's font, etc.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Outlook

- We fixed an issue where Outlook would prompt the user to save changes to a meeting when no changes were made.

### Word

- We fixed an issue where opening documents with track changes would cause the application to close unexpectedly if “Show paragraph” was enabled.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2308: July 18

*Version 2308 (Build 16717.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Desktop Excel Action Recorder:** You can now record Office Scripts in Desktop Excel! Perform the actions you want in Excel, and the recorder will create an Office Script to replay those actions for you.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Outlook

- We fixed an issue where users were unable to attach an Outlook item into an email when the item was selected from the Archive Mailbox (Auto Expanding Archive Mailbox).

### Word

- We fixed an issue where hovering on the anchored text tooltip for a comment, the tooltip would appear blank.

- We fixed an issue where copying text and pasting as Enhanced Metafile would cause the text to overflow to the next line.

- We fixed an issue where attempting to print more than one page of tasks in Outlook would result in an error message, "There is a problem with the selected printer".

### Office Suite

- We fixed an issue where importing customizations for the Quick Access Toolbar had no effect.

- We fixed an issue in Project where content was missing after scrolling.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2308: July 11

*Version 2308 (Build 16708.20004)*

Security updates listed [here](microsoft365-apps-security-updates.md).

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Word

- We fixed an issue where the content control end tag is marked at the end of the document automatically if the document is edited in Word Online and then opened in Word desktop.

- We fixed an issue where using the Repeating Section Content Control could cause the application to close unexpectedly.

- We fixed an issue where editing and saving a .doc file created with OpenOffice would cause the application to stop responding.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2308: July 05

*Version 2308 (Build 16628.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Office Suite

- We fixed an issue where broken SVG assets didn't fall back on their raster image counterparts.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2307: June 27

*Version 2307 (Build 16626.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Word

- We fixed an issue where the content control end tag is marked at the end of the document automatically if the document is edited in Word Online and then opened in Word desktop.

- We fixed an issue where multi-level lists may not retain alignment or indent values when saved and reopened.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2307: June 20

*Version 2307 (Build 16619.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Word

- We fixed an issue where the Header/Footer wasn't displaying correctly in Dark mode when inactive.

- We fixed an issue where the content control end tag is marked at the end of the document automatically if the document is edited in Word Online and then opened in Word desktop.

- We fixed an issue where table formatting was unexpectedly changed after dragging the column width to a different size in a table, saving, and reopening.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2307: June 13

*Version 2307 (Build 16610.20000)*

Security updates listed [here](microsoft365-apps-security-updates.md).

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Create richer analyses and summaries with PivotTables and data types:** PivotTables in Excel now support data types. Instead of just strings, bring data types such as images in cells to your PivotTable rows and columns, or use related properties from linked data types such as Stocks and Geography directly from the PivotTable pane.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where comments and notes weren't printed at the end of a sheet in a PDF file when using ExportAsFixedFormat.

### Outlook

- We fixed an issue where Outlook would close unexpectedly when saving a contact in the Address Book.

### Office Suite

- We fixed an issue where sensitivity labeling was unavailable for documents opened from SharePoint on-premise servers.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2307: June 06

*Version 2307 (Build 16602.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Office Suite

- **Ability to insert live camera feed in all slides with one click:** Now, one can insert a PowerPoint Cameo in all slides using the Insert > Cameo > All Slides command on the PowerPoint ribbon. Further, a user-customized camera style and location can be applied to all slides using the 'Apply to All Slides' option in the Camera Format ribbon.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

## Version 2306: May 30

*Version 2306 (Build 16529.20000)*

### Feature updates

### Outlook

- **Enabling Bookings Meta OS app in Win32:** Replacing the existing Bookings app with new app for Meta OS in Win32.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Word

- We fixed an issue where a “template not found” error would display when using Documents.Add add-ins.

### Office Suite

- We fixed an issue where the Mini Toolbar was cut off if the user's screen didn't have enough room to display the toolbar at its full size.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2306:  May 30

*Version 2306 (Build 16519.20004)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### PowerPoint

- **Present a local file in PowerPoint Live:** The "Present in Teams" button now supports presenting local files in PowerPoint Live in Microsoft Teams. Once you click the button, you are prompted to save the file to Microsoft cloud, and then automatically launch the presentation in PowerPoint Live.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

## Version 2306: May 25

*Version 2306 (Build 16519.20002)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an issue where OLEBD/ADODB interfaces exposed by the Access Database Engine from Excel may be slower in recent versions of Office.

### Excel

- We fixed an issue where Conditional Formatting rules weren't being preserved after closing and reopening a workbook.

### PowerPoint

- We fixed an issue where hyperlinks weren't working if they contained the "#" symbol when exporting to a .PDF file.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2306: May 09

*Version 2306 (Build 16505.20002)*

Security updates listed [here](microsoft365-apps-security-updates.md).

## Version 2305: May 03

*Version 2305 (Build 16501.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Protect your most sensitive content with double key encryption:** Double Key Encryption (DKE) provides organizations with a second encryption key for your most sensitive content. Learn more [about DKE](/microsoft-365/compliance/double-key-encryption) and whether it's suitable for your organization.

### Outlook

- **Accessibility Ribbon in Outlook for Windows:** The Accessibility Ribbon brings together in one place all the tools you need to make your emails accessible.

- **Protect your most sensitive content with double key encryption:** Double Key Encryption (DKE) provides organizations with a second encryption key for your most sensitive content. Learn more [about DKE](/microsoft-365/compliance/double-key-encryption) and whether it's suitable for your organization.

### PowerPoint

- **Protect your most sensitive content with double key encryption:** Double Key Encryption (DKE) provides organizations with a second encryption key for your most sensitive content. Learn more [about DKE](/microsoft-365/compliance/double-key-encryption) and whether it's suitable for your organization.

### Word

- **Protect your most sensitive content with double key encryption:** Double Key Encryption (DKE) provides organizations with a second encryption key for your most sensitive content. Learn more [about DKE](/microsoft-365/compliance/double-key-encryption) and whether it's suitable for your organization.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an issue where refreshing an ODBC connection to an Access database would cause the operation to stop working after refreshing the link several times.

### Excel

- We fixed an issue where OfficeJS commands didn't delete PivotTables.

- We fixed an issue where cell values remained visible on the grid after deleting the row, until the cells were scrolled off the screen.

### PowerPoint

- We fixed an issue where the PowerPoint Eyedropper Tool displayed the incorrect color when the Windows Accessibility Setting was set to Invert, either through Windows Color Filters, or the Windows Magnifier Tool.

### Word

- We fixed an issue where the background color of cells copied from Excel would be incorrect when pasted into Word.

- We fixed an issue where a dialog asking "This document contains links that may refer to other files. Do you want to update the document with the data from the linked files?" wouldn't appear for files with shortcuts.

### Office Suite

- We fixed an issue where the Insert and Link options weren't working as expected.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2305: April 25

*Version 2305 (Build 16421.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### PowerPoint

- **PowerPoint Accessibility Ribbon:** All of the tools to help you make your presentation accessible in one place.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an issue where the error message “Cannot open anymore databases” may appear when exporting too many reports to .PDF.

### Outlook

- We fixed an issue where the “From” field and Signatures weren't working while using Google Workspace Sync for Microsoft Outlook.

### PowerPoint

- We fixed an issue where autocorrect in PowerPoint marks and splits proofed words with incorrect spelling suggestions when using English (Sweden) as the Windows 10 language.

### Word

- We fixed an issue in which Word documents could freeze when co-authoring and trying to save.

### Office Suite

- We fixed an issue that prevented the description of sensitivity labels from being shown in tooltips when saving a new file.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2305: April 18

*Version 2305 (Build 16414.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an error that occurred when publishing as a web page (.htm or .html) to a cloud location.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2305: April 11

*Version 2305 (Build 16407.20000)*

Security updates listed [here](microsoft365-apps-security-updates.md).

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an issue that may cause the application to close unexpectedly when exporting from an SAS application to a Microsoft Office format.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2305: April 04

*Version 2305 (Build 16403.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an issue where a digital signature wouldn't carry over when saving an ACCDB as an ACCDE even if the user has access to the signing certificate.

### Excel

- We fixed an issue with scrolling the sheet tabs, where a scroll gesture would cause the worksheet to scroll rather than the tabs.

- We fixed an issue where the sheet tab scroll buttons wouldn't scroll left in some workbooks.

- We fixed an issue where Excel would stop responding when selecting the Clear option under Sort & Filter from the Editing group on the Ribbon.

- We fixed an issue where the arrow keys didn't work in a cell while in edit mode.

- We fixed an issue where Excel would stop working when charts were built on dynamic arrays from external links.

### Outlook

- We fixed an issue that caused users to see an empty From field drop down when their profile included an SMTP address with an asterisk in it.

- We fixed an issue where AllowPhraseMatch wasn't working when searching an uncached shared mailbox.

### PowerPoint

- We fixed an issue related to text selection when using an Input Method Editor for East Asian languages.

- We fixed an issue where the file couldn't be saved when Access Restriction and Digital Signature are applied.

- We fixed an issue where PowerPoint would stop responding for a few seconds when the user clicks on a complex shape.

### Project

- We fixed an issue where the task pane add-in API for Microsoft Project wasn't returning the proper value for the Summary, Milestone, and Active properties.

### Word

- We fixed a failure to export files to PDF if they were protected with restricted access or sensitivity labels that restrict rights.

- We fixed an issue where the right side of the email body was cut off when printed and in print preview.

- We fixed an issue where the Word macro function Application.PrintOut range:=wdPrintCurrentPage couldn't print the current page while using the macro in Draft view.

- We fixed an issue where Word would stop working when trying to save a new document to OneDrive.

### Office Suite

- We fixed an issue where the "connecting to" dialog box was shown during Save/Save-As when a UNC (Universal Naming Convention) path was set as the default save directory.

- We fixed an issue where right clicking a folder in Outlook could cause Office applications to close unexpectedly.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2304: March 30

*Version 2304 (Build 16327.20038)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2304: March 29

*Version 2304 (Build 16327.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue with scrolling the sheet tabs, where a scroll gesture would cause the worksheet to scroll rather than the tabs.

- We fixed an issue where the sheet tab scroll buttons wouldn't scroll left in some workbooks.

- We fixed an issue where Excel would stop working while trying to close a worksheet without saving changes.

### Outlook

- We fixed an issue that caused users to see an empty From field drop down when their profile included an SMTP address with an asterisk in it.

- We fixed an issue where AllowPhraseMatch wasn't working when searching an uncached shared mailbox.

### PowerPoint

- We fixed an issue related to text selection when using an Input Method Editor for East Asian languages.

- We fixed an issue when applying Access Restriction and Digital Signature the file couldn't be saved.

### Project

- We fixed an issue where the task pane add-in API for Microsoft Project wasn't returning the proper value for the Summary, Milestone, and Active properties.

### Word

- We fixed a failure to export files to PDF if they were protected with restricted access or sensitivity labels that restrict rights.

- We fixed an issue where the right side of the email body was cut off when printed and in print preview.

- We fixed an issue where the Word macro function Application.PrintOut range:=wdPrintCurrentPage couldn't print the current page while using the macro in Draft view.

### Office Suite

- We fixed an issue where Word would stop working when trying to save a new document to OneDrive.

- We fixed an issue where the "connecting to" dialog box was shown during Save/Save-As when a UNC (Universal Naming Convention) path was set as the default save directory.

- We fixed an issue where right clicking a folder in Outlook could cause Office applications to close unexpectedly.

- We fixed an issue where PowerPoint stopped responding for a few seconds when the user clicks on a complex shape.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2304: March 21

*Version 2304 (Build 16316.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where copying and pasting a pivot chart would cause Excel to close unexpectedly.

- We fixed an issue where a name conflict appeared randomly when files were opened from Teams and SharePoint sites.

### Outlook

- We fixed an issue that caused the application to close unexpectedly when opening the Tracking tab for meetings from the Manager's calendar.

### Word

- We fixed an issue where Word would default to the Big 32 Kai paper size when the current system locale is set to “Chinese (Simplified, China)”.

### Office Suite

- We fixed an issue where a "Downloading Fonts..." message in the status bar didn't refresh if a font request failed.

- We fixed an issue that prevented the description of sensitivity labels from being shown in tooltips when hovering with a mouse on a parent label.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2304: March 14

*Version 2304 (Build 16310.20000)*

Security updates listed [here](microsoft365-apps-security-updates.md).

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an issue where elements of a query design grid may appear on top of the subform datasheet when a subform control had its Source Object property set to a query, and the parent form was viewed in Form View.

### Excel

- We fixed an issue where some users were unable to expand/collapse column headers using the plus and minus icons.

### Outlook

- We fixed an issue where some settings didn't roam between machines when switching to Focused Inbox.

### Word

- We fixed an issue where the bullet character wasn't rendering correctly in the Table of Contents in Portable Document Format (PDF) files.

- We fixed an issue where bookmarks in text boxes would disappear from Rich Text Format documents if the name of the bookmark was specified as Unicode.

- We fixed an issue that caused Word to close unexpectedly when automation was calling the application.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2304: March 09

*Version 2304 (Build 16307.20006)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Access

- **Modern Web Browser Control:** The Access team is releasing a new web browser control version that supports Edge! Our existing, Internet Explorer browser control remains in the platform as-is.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue that caused the application to close unexpectedly when using Visual Basic for Applications (VBA) to modify a Waterfall Chart.

- We fixed an issue where large cell selection operations might fail in virtual machine environments.

### Outlook

- We fixed an issue where users were unable to login using their personal account.

### Word

- We fixed an issue where Viewing and Editing modes could prevent modifying editable areas in restricted documents.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2303: March 01

*Version 2303 (Build 16227.20004)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Blocking XLL add-ins from the Internet:** To address the increasing number of malware attacks in recent months, we are implementing measures that block XLL add-ins coming from the Internet.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an issue that may cause the application to close unexpectedly when exporting from an SAS application to a Microsoft Office format.

### Project

- We fixed an issue so that progress lines now scroll vertically together with their tasks when a Gantt type view is scrolled up and down.

### Word

- We fixed an issue where the error message "The last time you opened *filename*, it caused serious error. Do you still want to open it?" may appear when creating Word documents through automation using templates.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2303: February 21

*Version 2303 (Build 16216.20004)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Outlook

- We fixed an issue where users couldn't see the sensitivity label applied to an email when the label wasn't included in their label policy.

### Project

- Fixed an issue where the ACWP earned value didn't always roll-up as expected from sub tasks to summary tasks. A side effect of the problem is that other earned value calculations such as TCPI, which rely on ACWP, were incorrect.

### Office Suite

- We fixed an issue that when a user creates a custom ribbon and the label starts with two zero width joiners, the application closes unexpectedly.

- We decreased the time it takes to open the font dialog in PowerPoint when the default editing language is set to an East Asian language.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2303: February 14

*Version 2303 (Build 16206.20000)*

Security updates listed [here](microsoft365-apps-security-updates.md).

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Reducing unwanted fragmenting of conditional formatting rules:** Check out the speed up in workbooks with lots of unwanted fragmented conditional formatting rules. When a workbook is opened this feature merges those fragmented conditional formatting rules that are identical, within a contiguous range of cells, and with unchanged priority ordering. It excludes rules whose evaluation relies on a selection range like Above or Below average, Unique or Duplicate, Gradients, etc. and rules in PivotTables.

- **Hyperlinks in Modern Comments:** Add hyperlinks to your Modern Comments in Excel [Learn more](https://support.office.com/article/bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)<br />See details in [blog post](https://insider.office.com/blog/add-hyperlinks-to-your-threaded-comments-in-excel)

### Outlook

- **Enhanced discovery of sensitivity labels in Outlook:** Admins can configure sensitivity labels to appear on subject lines when composing e-mails, for enhanced user discovery. The labels icons can also be customized.

- **Inheritance of attachment labels to email messages:** For email messages with attachments, apply a label that matches the highest classification of those attachments

- **Block emails with sensitive labels:** Implement pop-up messages in Outlook that warn, justify, or block emails being sent based on sensitivity labels.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where the Ribbon options dropdown menu could disappear.

- We fixed an organization check for accounts to enable Power BI.

### Outlook

- We fixed an issue in the Outlook calendar, where clicking with the mouse on a timeslot selected a different timeslot.

### PowerPoint

- We fixed an issue where the Ribbon options dropdown menu could disappear.

### Project

- We fixed an issue that caused a manually-scheduled task's start date to be after its finish date when an XML file was opened in Project.

### Word

- We fixed an issue where the Ribbon options dropdown menu could disappear.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2303: February 07

*Version 2303 (Build 16202.20000)*

## Version 2302: February 03

*Version 2302 (Build 16130.20020)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Office apps can now audit protection properties applied to a document through sensitivity labels:** Added the missing audit fields related to the protection status of a document and how the protection was changed due to a specific event. That would include the following fields: protectionEventType, documentEncrypted, documentEncryptedBefore, protectionType, protectionTypeBefore, protectionTemplateId, protectionTemplateIdBefore, protectionOwner and protectionOwnerBefore.

### Outlook

- **Office apps can now audit protection properties applied to a document through sensitivity labels:** Added the missing audit fields related to the protection status of a document and how the protection was changed due to a specific event. That would include the following fields: protectionEventType, documentEncrypted, documentEncryptedBefore, protectionType, protectionTypeBefore, protectionTemplateId, protectionTemplateIdBefore, protectionOwner and protectionOwnerBefore.

### PowerPoint

- **Closed Captions for Audio in PowerPoint:** Add closed captions to audio objects in PowerPoint to make your presentations more accessible.

- **Office apps can now audit protection properties applied to a document through sensitivity labels:** Added the missing audit fields related to the protection status of a document and how the protection was changed due to a specific event. That would include the following fields: protectionEventType, documentEncrypted, documentEncryptedBefore, protectionType, protectionTypeBefore, protectionTemplateId, protectionTemplateIdBefore, protectionOwner and protectionOwnerBefore.

### Word

- **Office apps can now audit protection properties applied to a document through sensitivity labels:** Added the missing audit fields related to the protection status of a document and how the protection was changed due to a specific event. That would include the following fields: protectionEventType, documentEncrypted, documentEncryptedBefore, protectionType, protectionTypeBefore, protectionTemplateId, protectionTemplateIdBefore, protectionOwner and protectionOwnerBefore.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue that compiled VBA macros more often.

- Fixed an issue where large cell selection operations might fail in virtual machine environments.

### Word

- Fixed an issue when the last key was released quickly, sometimes AutoComplete wouldn't perform.

- Fixed an issue where attempting to open an .odt files with a table would result in an error message, "Word found unreadable content. Do you want to recover this document".

### Office Suite

- We fixed an issue in PowerPoint record window in case of audio or video hardware issue. In case either the audio or video hardware fails to provide picture or voice then we were waiting forever with an assumption that this should never happen when the initial communication with those devices are successful. The changes will no longer make that assumption and fail sooner, so that the user can toggle (off/on) the devices and try to record again.

- We fixed an issue that caused Mandatory Label warning bar to show even though the policy has Default label on it. The issue is for WXP documents in one drive sync folder in Win32. The fix checks if default label application was ever attempted.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2302: January 25

*Version 2302 (Build 16124.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Assign a sublabel as the default when a parent label is selected:** When using built-in sensitivity labels in Microsoft 365 Apps, admins can specify a sublabel to get applied automatically when a parent label is selected. This takes effect only when users select a parent label manually.

- **Reducing slowness and freezes when multiple workbooks are open:** This feature reduces slowness and freezes experienced when working in a workbook due to calculations occurring in other unrelated workbooks also open at the same time and in the same Excel.exe instance. It achieves this by optimizing global automatic recalculation to the workbook being worked in, and its interdependent workbooks also open at the same time.

- **Formula evaluation tooltips:** Select a part of your formula and you see a tooltip showing the current value of just the part you've selected.

### Outlook

- **Assign a sublabel as the default when a parent label is selected:** When using built-in sensitivity labels in Microsoft 365 Apps, admins can specify a sublabel to get applied automatically when a parent label is selected. This takes effect only when users select a parent label manually.

### PowerPoint

- **Assign a sublabel as the default when a parent label is selected:** When using built-in sensitivity labels in Microsoft 365 Apps, admins can specify a sublabel to get applied automatically when a parent label is selected. This takes effect only when users select a parent label manually.

### Word

- **Assign a sublabel as the default when a parent label is selected:** When using built-in sensitivity labels in Microsoft 365 Apps, admins can specify a sublabel to get applied automatically when a parent label is selected. This takes effect only when users select a parent label manually.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- An issue with #Deleted when utilizing the Datetime2(3) column in the clustered primary key has been fixed.

### Excel

- We fixed an issue where Charts in PowerPoint or Word do not update to the data series set by other users during coauthoring sessions.

### Outlook

- We fixed an issue that caused users to be unable to view Actionable messages, restart, or shut down Windows after cancelling a Windows restart operation.

- We fixed an issue that caused attendees to see the wrong end date on All Day Events when explicitly opening the item.

### PowerPoint

- We fixed an issue where PowerPoint wasn't responding and then closed unexpectedly when using the Compare function within Review menu.

### Project

- Fixes an issue where upon opening an MPP file, you see the following error:

   "An import error occurred.

   Check row 1 column 5.

   To continue importing with other error messages, click Yes.
   To continue importing with no error messages, click No.
   To stop importing, click Cancel."

- Fixed an issue where the task pane add-in API for Microsoft Project doesn't always set or get values for the following properties:

   ConstraintType
   EarnedValueMethod
   FixedCostAccrual
   OutlineLevel
   PercentComplete
   PercentWorkComplete
   Priority
   Status
   Type
   Active
   Milestone
   Summary

### Word

- Resolved an issue where selecting text inside a text box could cause the application to close suddenly.

- Resolved an issue where printing last section would print entire document.

- Resolved an issue where commenting in a document can disrupt the formatting of a table.

- Resolved an issue with Selection.Repeat command when used after Selection.TypeText and Selection.TypeParagraph commands.

- Resolved an issue where Check-out business bar not displaying when opening files from SPO with file path longer than 260 characters.

- Resolved an issue when linked objects update mode is set to manual, on file reopen the link will never update properly, regardless of clicking yes/no on the dialog.

- We fixed an issue where applying the blockcontentexecutionfrominternet policy would prevent files containing an embedded macro to no longer to run.

### Office Suite

- We fixed an issue where leading spaces were unexpectedly removed when importing SVG images in Office 365.

- We fixed an issue where fonts and data point shadows weren't rendering correctly when exporting charts from Excel and importing into PowerPoint using SVG Format and then converting to shape.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2302: January 17

*Version 2302 (Build 16110.20000)*

### Resolved issues

### Outlook

- We fixed an issue where some users were getting high CPU and heavy flicker while scrolling around in the calendar.

### Word

- Resolved an issue where printing last section would print entire document.

## Version 2302: January 13

*Version 2302 (Build 16107.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Office Suite

- **Automatically insert image captioning for Excel's images:** Insert image into Excel's spreadsheet and accessibility image captioning is automatically generated for you.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an issue where Datasheet displayed incorrect record when simultaneous entries were submitted to SQL Server.

- We fixed an issue where when opening a report in Print Preview on 64-bit Access caused an error message to appear.

### Excel

- We fixed an issue where Excel might close unexpectedly when loading foreign files while there are drawing objects in the active sheet.

### Word

- Resolved an issue where using Word VBA function InsertFile to insert .docx file would result in 'Error: 5097 - Word has encountered a problem'.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2301: January 03

*Version 2301 (Build 16029.20000)*

## Version 2301: December 30

*Version 2301 (Build 16026.20002)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an issue that could cause the TransferDatabase command to be blocked when connecting to a database, even when it's in a Trusted Location.

- We fixed an issue that caused current queries to fail because half and full width characters were compared incorrectly.

### PowerPoint

- We fixed an issue where presentations with redundant custom XML datastore items and an encrypted label failed to open.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2301: December 23

*Version 2301 (Build 16015.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Outlook

- We fixed an issue that caused users to see an NDR in Local Failures and to get an "Operation Failed" message when replying to some S/MIME messages.

### Word

- Fixed an issue where Restrict Editing wouldn't enforce properly and result in broken table layout.

- Resolved an issue where applying paragraph shading to text spanning across multiple lines draws the shading incorrectly in between lines until the first paragraph mark appears.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2301: December 16

*Version 2301 (Build 16012.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **PivotTable overlap improvements:** We have improved the experience when PivotTables overlap other content in your workbook.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue when enlarging the size of TreeMap Chart to a certain size in PowerPoint the user receives a message that the file couldn't be saved, and the file is opened in read-only mode.

- We fixed an issue when you right click on a chart and select edit option the application closes unexpectedly.

### PowerPoint

- We fixed an issue when using a Web add-In built on the Shared Runtime, PowerPoint would stop working.

- We fixed an Issue where PowerPoint would stop responding in some situations using touch screen.

- We fixed an issue where applying the blockcontentexecutionfrominternet policy would prevent files containing an embedded macro to no longer to run.

- We fixed an issue when clicking on a Presentation document with an object inserted in the PowerPoint Presentation document, a Security Alert (OLE Actions have been blocked) was displayed.

### Word

- We fixed an issue where Word opened in read-only mode when remember my credentials was checked and Certificate required.

- Fixed an issue where the Display for Review dropdown is enabled for read only documents.

- Fixed an issue with scrolling with a trackpad.

- Fixed an issue where tables render blurry in a Word doc.

- Fixed an issue when an image is copied from Outlook and pasted into Word using the Picture (PNG) format it displays cropped.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2301: December 09

*Version 2301 (Build 15929.20006)*

[//]: # (DO NOT REMOVE)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue when you right click on a chart and select edit option the application closes unexpectedly.

### PowerPoint

- We fixed an issue when clicking on a Presentation document with an object inserted in the PowerPoint Presentation document, a Security Alert (OLE Actions have been blocked) was displayed.

- We fixed an issue where applying the blockcontentexecutionfrominternet policy would prevent files containing an embedded macro to no longer to run.

### Word

- Fixed an issue where the Display for Review dropdown is enabled for read only documents.

- Fixed an issue with scrolling with a trackpad.

- Fixed an issue where tables render blurry in a Word doc.

- Fixed an issue when an image is copied from Outlook and pasted into Word using the Picture (PNG) format it displays cropped.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2212: December 02

*Version 2212 (Build 15928.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Access

- **Enable the ability to code sign your Microsoft Access database and VBA code:** This update enables the Tools/Digital Signature command within the VBA (Visual Basic for Applications) IDE (Integrated Development Environment) for current Microsoft Access database formats.  Signing a database allows VBA code in the database to be run even if Trust Center settings specify that only digitally signed code should be enabled.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Outlook

- We fixed an issue where the Outlook Ribbon is unresponsive when downloading Address Book from the File Menu.

- We fixed an issue that was causing users to see multiple copies of a shared calendar rendered in certain circumstances.

### Word

- We fixed an issue where users received a message that changes  be merged when accessing files as Read Only.

- Resolved an issue where tasks may not show full name of assigned user.

- Fixed an issue where after inserting comments on certain documents the application closes unexpectedly.

- Resolved an issue where Word would stop working when autosaving a template with a building block that contains a Rich Text content control with an image.

- Fixed an issue when exporting data from IBM DOORs application to Word the application would close unexpectedly.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2212: November 25

*Version 2212 (Build 15917.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Word

- We fixed an issue where no visual indicator would show which wrap text option was selected.

- We fixed an issue where Word Autoformat unexpectedly creates lots of revisions when creating new table row in a document.

- Resolved an issue where VBA Application.OrganizerCopy would fail for files with filenames containing URL escape characters.

### Office Suite

- Fixed an issue where users were unable to see the labels applied to documents/emails after sensitive labels had been removed from the label policy.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2212: November 18

*Version 2212 (Build 15911.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### PowerPoint

- We fixed an issue where a new presentation wasn't created when using the PowerPoint.createPresentation API.

### Word

- We fixed an issue where text was unexpectedly hidden by a shape even though the text wrapping option was selected.

### Office Suite

- We fixed an issue where applying the blockcontentexecutionfrominternet policy would prevent files containing an embedded macro to no longer to run.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2212: November 11

*Version 2212 (Build 15904.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Insert in-cell images with the new IMAGE function:** Your images can now be part of the worksheet, instead of floating on top. You can move and resize cells, sort and filter, and work with images within an Excel table.

- **Touch Improvements to Ribbon:** We've improved the spacing of buttons in the Ribbon when the device is being used in Tablet posture.

### PowerPoint

- **Touch Improvements to Ribbon:** We've improved the spacing of buttons in the Ribbon when the device is being used in Tablet posture.

### Word

- **Touch Improvements to Ribbon:** We've improved the spacing of buttons in the Ribbon when the device is being used in Tablet posture.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

## Version 2211: November 04

*Version 2211 (Build 15831.20012)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### PowerPoint

- **Save Media with Closed Captions:** Now when you save media to a file in PowerPoint, the closed captions associated with the media are also saved.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Word

- We fixed an issue where no visual indicator would show which wrap text option was selected.

- We fixed an issue where content would flash with every keystroke when typing in a tracked change in a table.

- We fixed an issue where printing would stop working.

- We fixed an issue where a paragraph before non-editable content couldn't be deleted.

### Office Suite

- We fixed an issue where a file, opened as protected, would remain in a temp folder after the user closed it and exited Office.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2211: October 28

*Version 2211 (Build 15822.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Word

- We fixed an issue where fonts could be corrupted in Styles, resulting in unrecognized character squares.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2211: October 21

*Version 2211 (Build 15813.20002)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an issue that caused the Find and Replace dialog box in the SQL view of a query to display black text on a black background, rendering it illegible.

### Excel

- We fixed an issue where multiple cells could be selected when clicking a cell while freeze panes are enabled and another app is listening for Focus-Changed events.

### Outlook

- We fixed an issue where Outlook couldn't display contacts in People View.

### Project

- We fixed an issue that didn't allow users to set Trusted Locations for opening files that contain macros in the Options dialog box.

### Word

- We fixed an issue where the Ruler didn't show correct margins for a paragraph after the Mirror margins setting had been applied.

- We fixed an issue that forced users to choose between saving a copy or discarding an unmergeable file, and we now enable them to either keep their own version of the file or keep the server version.

- We fixed an issue where choosing the Keep source formatting option when pasting content would result in characters such as a question mark (?) or a blank square being displayed.

- We fixed an issue that prevented files saved to the cloud from opening when using the Update linked information in a Word source document command.

### Office Suite

- We fixed an issue where a status message would appear and remain on the screen for several minutes when trying to save Office files or Outlook attachments to a network shared location.

- We fixed an issue where Visual Studio Online add-ins were getting disabled.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2211: October 14

*Version 2211 (Build 15806.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **New shortcut key to Paste Values:** You can now use the keyboard shortcut CTRL+SHIFT+V to “Paste Values” in Excel for Windows.<br />See details in [blog post](https://insider.office.com/blog/new-paste-options-when-using-keyboard-shortcuts)

- **Disable the Azure Information Protection Add-in by default:** Office apps will now automatically disable the legacy Azure Information Protection add-in and use the built-in sensitivity labels to view and apply labels powered by Microsoft Purview Information Protection.  

### Outlook

- **Labels based on document type:** Labels are defined so that they're connected to specific document types.

- **Disable the Azure Information Protection Add-in by default:** Office apps will now automatically disable the legacy Azure Information Protection add-in and use the built-in sensitivity labels to view and apply labels powered by Microsoft Purview Information Protection.

### PowerPoint

- **Disable the Azure Information Protection Add-in by default:** Office apps will now automatically disable the legacy Azure Information Protection add-in and use the built-in sensitivity labels to view and apply labels powered by Microsoft Purview Information Protection.

### Word

- **Disable the Azure Information Protection Add-in by default:** Office apps will now automatically disable the legacy Azure Information Protection add-in and use the built-in sensitivity labels to view and apply labels powered by Microsoft Purview Information Protection.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where changes made to links to a shared formula wouldn't update each cell in the shared formula.

- We fixed an issue that prevented content copied from Microsoft Paint to be pasted from the Clipboard using the Change Picture command.

### Outlook

- We fixed an issue where secondary, shared calendars will not scroll in sync with the primary calendar in the view.

### PowerPoint

- We fixed an issue that prevented content copied from Microsoft Paint to be pasted from the Clipboard using the Change Picture command.

### Project

- We fixed an issue where Project was unable to connect to Project Server when forms-based authentication was used.

### Word

- We fixed an issue that prevented content copied from Microsoft Paint to be pasted from the Clipboard using the Change Picture command.

- We fixed an issue where the Reviewing pane failed to open despite what was indicated in the Unsupported Content waning message.

### Office Suite

- We fixed an issue that resulted in text items being either reordered or incorrectly positioned when converting SVG graphics to shapes.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2211: October 07

*Version 2211 (Build 15729.20002)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Outlook

- We fixed an issue that prevented messages from being sent when using multiple Exchange accounts with one Outlook profile.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2210: September 30

*Version 2210 (Build 15726.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Automate tab in Excel for desktop:** You can now access the Automate tab not just in your web browser, but also in your spreadsheets in Excel for Windows and Mac.

### PowerPoint

- **Easily edit auto generated alt text:** Click edit on the caption bar to easily modify auto alt text [Learn more](https://support.office.com/article/6f7772b2-2f33-4bd2-8ca7-dae3b2b3ef25)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where text formatting caused Excel to stop working.

### Outlook

- We fixed an issue where Web view showed an error message on restart and cancel.

- We fixed an issue where the scroll bar wasn't visible while reading or composing a message.

### Project

- We fixed an issue where only one cross project linking option was applied, even when mutiple were selected.

### Word

- We fixed an issue where the Read Aloud feature wasn't available when Microsoft 365 was deployed on Windows Server 2016.

### Office Suite

- We fixed an issue where Excel files opened from a Sharepoint 2019 library intermittently opened as read-only in Excel for Windows.

- We fixed an issue that prevented locked Word and Excel files from opening.

- We fixed an issue that slowed down the saving of PowerPoint slides with SVG content to metafile formats.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2210: September 23

*Version 2210 (Build 15715.20014)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where a function would run out of resources or display the wrong error message when alttext had no value.

- We fixed an issue where app would stop responding and consume huge amounts of memory while doing touch zoom on RTL sheets.

### OneNote

- We fixed an issue where users couldn't use the panning tool with stylus.

### Outlook

- We fixed an issue where search would always return results from the primary account and not the selected account.

- We fixed an issue that caused attachment size limits to not be evaluated when drag/dropping an attachment into a mail message.

- We fixed an issue where the app closed unexpectedly when closing the Mail control panel.

### PowerPoint

- We fixed an issue where PowerPoint could be prevented from closing when accessed as an IOleObject.

### Word

- We fixed an issue where hyperlinks weren't working, and an error was displayed.

- We fixed an issue in interaction mode with application guard.

- We fixed an issue where after splitting document window, lower window didn't scroll to expected position below cursor.

### Office Suite

- We fixed an issue where the font size for labels increased unexpectedly in QAT when the icon sizes were updated.

- We fixed an issue where a customer certificate was incorrectly shown as being revoked.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2210: September 16

*Version 2210 (Build 15709.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Dynamic Array Integration with Charts:** This feature enables users to link charts to Dynamic array calculations, which can produce results of variable length. The chart will automatically update to capture all data when the array recalculates, rather than being fixed to a specific number of data points.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### OneNote

- We fixed an issue where, after inserting a shape or line onto the canvas, the user was unable to drag or resize the item.

### Outlook

- We fixed an issue where disconnecting IMAP IDLE session caused stopping of IMAP sync until rebooting.

- We fixed an issue where SharePoint documents, when attached as a copy, wouldn't download for URLs above a given size.

### Project

- We fixed an issue where it sometimes wasn't possible to synchronize resources to a SharePoint List when the resource's name included the system list separator.

### Word

- We fixed an issue with large URLs where a link couldn't be opened if its length exceeded a specific character limit.

- We fixed an issue where header styles would be removed upon co-authoring.

- We fixed an issue where if a footer was added by built-in labeling, it incorrectly moved existing, manually added footers.

### Office Suite

- We fixed an issue when running Convert to Shape on some SVG graphics that contain text.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2209: September 09

*Version 2209 (Build 15629.20058)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Prevent data leaks more easily with the new Sensitivity toolbar:** Sensitivity labels powered by Microsoft Purview Information Protection are now displayed alongside the filename in the app's title bar, allowing you to easily recognize and adhere to your organization's policies. The sensitivity toolbar is also available while saving new documents or renaming existing ones, helping you keep information security at your fingertips.

- **Copy Paste is more efficient:** Copy/paste now uses an index-based search, rather than a linear scan, for more optimal merging. The optimizations are especially noticeable on low-end devices with constrained hardware.

- **User-defined permissions now support domain name restrictions:** When you choose a sensitivity label configured for user-defined permissions, they can now use domain names to restrict file access to all individuals from that domain. For example, you can specify “<someone@example.com>” or “@example.com,” and permissions are restricted based on either then individual or all individuals within the example domain.

### Outlook

- **We added a registry key that hides the “Try the new Outlook” toggle:**

    Registry Key:

    >HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Options\General</br>
    >REG_DWORD "HideNewOutlookToggle”</br></br>
    >0 (default) - “Try the new Outlook” toggle, if available in selected update channel, is displayed to users.</br>
    >1 - “Try the new Outlook” toggle is hidden.

  To learn more about the new Outlook for Windows, please see [Getting Started with the new Outlook for Windows](https://support.microsoft.com/office/getting-started-with-the-new-outlook-for-windows-656bb8d9-5a60-49b2-a98b-ba7822bc7627). For more information on managing mailbox access to the new Outlook for Windows, please see [Clients and Mobile in Exchange Online](/exchange/clients-and-mobile-in-exchange-online/outlook-on-the-web/enable-disable-employee-access-new-outlook).

- **S/MIME as an Outcome for labelling - Win 32:** Providing S/MIME encryption and signing functionality as an outcome of labelling - Outlook Win32.

### PowerPoint

- **Prevent data leaks more easily with the new Sensitivity toolbar** Sensitivity labels powered by Microsoft Purview Information Protection are now displayed alongside the filename in the app's title bar, allowing you to easily recognize and adhere to your organization's policies. The sensitivity toolbar is also available while saving new documents or renaming existing ones, helping you keep information security at your fingertips.

- **User-defined permissions now support domain name restrictions:** When you choose a sensitivity label configured for user-defined permissions, they can now use domain names to restrict file access to all individuals from that domain. For example, you can specify “<someone@example.com>” or “@example.com,” and permissions are restricted based on either then individual or all individuals within the example domain.

### Word

- **Prevent data leaks more easily with the new Sensitivity toolbar** Sensitivity labels powered by Microsoft Purview Information Protection are now displayed alongside the filename in the app's title bar, allowing you to easily recognize and adhere to your organization's policies. The sensitivity toolbar is also available while saving new documents or renaming existing ones, helping you keep information security at your fingertips.

- **User-defined permissions now support domain name restrictions:** When you choose a sensitivity label configured for user-defined permissions, they can now use domain names to restrict file access to all individuals from that domain. For example, you can specify “<someone@example.com>” or “@example.com,” and permissions are restricted based on either then individual or all individuals within the example domain.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where the app closed unexpectedly when accessing files containing rich entities.

- We fixed an issue where Excel could close unexpectedly when using Office Scripts to manipulate PivotTable fields.

- We fixed an issue where an Excel file could become corrupt after setting formatting (such as fill color) on some cells in a PivotTable on the row or column axis and then moving those fields to the Filter area of the PivotTable.

- We fixed an issue with the sheet navigation buttons, which were flipped when using a sheet in right-to-left orientation.

- We fixed an issue where the app would close unexpectedly when selecting error #FIELD! cell in sheet and then opening Error checking OOUI field option for Field errors.

### Outlook

- We fixed an issue that caused the Customization Quick Access Toolbar file (.exportedUI) to fail to import when the simplified ribbon was in use.

- We fixed an issue that caused users to be unable to see the Due Date on search results in Outlook Desktop.

### Word

- We fixed an issue where the user couldn't click on a hyperlink in track changes.

- We fixed an issue where comments might not load.

- We fixed an issue where Word may not show the right file when all open file windows have been minimized.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2209: August 26

*Version 2209 (Build 15619.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Reducing unwanted fragmenting of conditional formatting rules:** Check out the improved performance and faster calculations when applying conditional formatting rules to cells. We've done so by reducing fragmentation of data when pasting copied cells into that a cell range.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an issue where the app closed unexpectedly if memory requirements grew too rapidly when using DAO or OLEDB interfaces to read Access databases.

### Excel

- We fixed an issue where the incorrect background color was displayed on certain menu gallery items.

- We fixed an issue that prevented a user from selecting table rows and columns after selecting the Insert Table command when a Coachmark was displayed on the screen.

- We fixed an issue where the password caching function didn't work with MSQuery x64 build.

### Outlook

- We fixed an accessibility issue where Narrator didn't read out-of-office status when automatic replies were enabled.

### PowerPoint

- We fixed an issue that caused cropping of pictures to not work properly.

- We fixed an issue where the user couldn't exit the app due to encryption rights.

### Word

- We fixed an issue where synchronization issues were resolved when query changes were cancelled.

- We fixed an issue where the app would stop responding when saving a document that contained 3D content.

- We fixed an issue where an app would show an error message when pasting a picture from Microsoft Paint.

- We fixed an issue where user was unable to copy rows of a table from older versions in Version History to the current version.

- We fixed an issue where when using Track Changes the app was unable to accept all changes.

- We fixed an issue where, after many documents had been closed using Task Manager and later restored when Word was open again, the View mode switcher in the status bar would be disabled in one of the documents, while being enabled in all the others.

- We fixed an issue where content in a rotated textbox became horizontal and WordArt Text Effects were suppressed when exported to PDF/A.

### Office Suite

- We fixed an issue where when trying to save Office files or Outlook attachments to network shared path, a status message would appear and remain on screen for several minutes.

- We fixed an issue where when switching commanding modes between Toolbar (preview) and Classic, the mode switch didn't apply to all active app windows.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2209: August 19

*Version 2209 (Build 15615.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Date Support for Pivot Tables Connected to Power BI:** In PivotTables that are connected to Power BI datasets, dates are now supported for analysis as the data type is no longer a string. For example, filtering data on specific timeframes is now possible.<br />See details in [blog post](https://insider.office.com/blog/improvements-to-the-connected-power-bi-experience)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue that caused the upload to OneDrive of a newly saved file to which a sensitivity label has been applied and an edit had been made, to fail.

- We fixed an issue where Excel would incorrectly use a fixed width rather than the full width space when converting text to column.

- We fixed an issue where the index variables passed to the lambda in MAKEARRAY were of the wrong type, resulting in an erroneous evaluation of the formula.

### Outlook

- We fixed an issue where linked images wouldn't display when replying to or forwarding a message.

### Project

- We fixed an issue where projects could no longer be opened because new task and assignment GUIDs were generated when the project was saved.

- We fixed an issue where a message was displayed erroneously warning customers about potentially unsafe links when opening a project that didn't contain any links.

### Word

- We fixed an issue where Outlook couldn't open a message that was sent using Outlook on the web and contained a comment that was copied from Word.

- We fixed an issue that prevented changes from being merged into a shared document due a busy server error.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2209: August 12

*Version 2209 (Build 15606.20002)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Word

- **Assign Tasks:** You can facilitate collaboration and stay organized by creating and assigning tasks directly from within your Word document using @mentions in comments. The people you assign the tasks to will receive email notifications, letting them know they need to take action. Ready to give it a try? Open an existing document stored on OneDrive or SharePoint. Highlight the text that contains the information you want to comment on, select New Comment, write your comment and type @ followed by the name of the team member you want to tag. Then, press Ctrl + Enter to post your comment and select the Assign to check box to convert your comment into a task.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue affecting only Japanese language users that caused macros to no longer work.

### Project

- We fixed an issue affecting only Hebrew language users where the Resources tab within the Task Information dialog box didn't display the four expected columns within the grid.

- We fixed an issue where inaccurate actual cost values could be viewed and reported when the timescale was modified.

### Word

- We fixed an issue in the business bar during merge conflict.

- We fixed an issue where outdated property values may be displayed under File > Info.

### Office Suite

- We fixed an issue where the Date Modified field was updated in SharePoint when the file was opened but not modified.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

## Version 2208: August 05

*Version 2208 (Build 15601.20028)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an issue where the Clean cache on close option didn't delete the cache when the database was closed.

### Excel

- We fixed an issue with the format for the headers in the task panes.

- We fixed an issue where Excel chart contents were missing when zooming above 170%.

### Outlook

- We fixed an issue that caused some mail-storage-related settings to not be applied as expected when configured through Cloud Settings.

- We fixed an issue where Outlook closed unexpectedly when using the Search People box.

- We fixed an issue where Skype for Business closed unexpectedly when trying to expand a group.

- We fixed an issue that caused a file to not open in Edge after a user chose the View in browser option in an Outlook message.

- We fixed an issue that caused actions performed on Outlook toast notifications to stop working as expected on Windows Server 2016.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

## Version 2208: July 29

*Version 2208 (Build 15522.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Office Suite

- **New Rating question added to Polls app:** Many Teams users wanted a way to quickly get feedback from meeting attendees on how satisfied they were or how well they understood the meeting content. So, we've added a new Rating question type to the Polls app within Microsoft Teams. Meeting owners can easily create and launch Rating polls to increase engagement and collect input from their meeting attendees, as well as share the results in real time.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### PowerPoint

- We fixed an issue where thumbnail content for pictures wouldn't render properly on the slide after clicking "Enable Content" on the security warning.

### Project

- We fixed an issue where scrolling down in Network Diagram view when the layout mode is set to manual doesn't always draw the nodes.

### Word

- We fixed an issue where users were unable to open links to locally managed files with bookmarks to files on SPO/OD4B.

- We fixed an issue where internal links within Track Changes were rendering as "Error! Hyperlink reference not valid."

- We fixed an issue where the app experienced delays when opening the HTML-file version of Outlook messages.

### Office Suite

- We fixed an issue where some users couldn't access the latest version of Excel files hosted on SPO and were seeing a "Refresh Recommended" error message.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2208: July 22

*Version 2208 (Build 15519.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)
[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Outlook

- We fixed an issue with the limit on the number of sensitivity label templates that was causing templates to not properly load.

- We fixed an issue in Outlook Search where the green Replied icon wasn't appearing next to emails that had replies.

### Word

- We fixed an issue where an Outlook (WordMail) message would flash with a white background while zooming in or out.

- We fixed an issue to add logging of user behavior when the Similarity Checker isn't available to them.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2208: July 15

*Version 2208 (Build 15511.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Improvements made to data models in Power BI:** We've made some upgrades to the analysis services components used in data modeling and Power Pivot in Excel. This is a maintenance update for the data modeling engine, and it won't noticeably change any interactions related to the Excel data model.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Outlook

- We fixed an issue affecting Windows Information Protection (WIP) when trying to retrieve EDP email addresses.

### Project

- We fixed an issue where local custom fields stopped working prompting "Delete Custom Fields" dialog box and couldn't be deleted by the organizer.

### Word

- We fixed an issue where "paste as link" might not update automatically.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2208: July 08

*Version 2208 (Build 15505.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Outlook

- **Get relevant alerts with new Notifications pane:** Don't let important information get buried in your inbox. The new Notifications pane in Outlook delivers notifications that are relevant to you in the context of your regular email. The pane gives you the ability to customize the types of notifications you wish to receive, including email and document @mentions, travel updates, deliveries, and more. [Learn more](https://insider.office.com/blog/new-notifications-pane-in-outlook-helps-you-stay-on-task)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### PowerPoint

- We fixed an issue by making minor updates to the smart Quick Access Toolbar.

### Project

- We fixed an issue where the actual hour totals in timesheets weren't always recalculated correctly.

- We fixed an issue so that progress lines now scroll vertically together with their tasks when a Gantt type view is scrolled up and down.

### Word

- We fixed an issue where the app was blocking access to a document after closing it, so that OneDrive couldn't sync the document until the app was exited.

- We fixed an issue where comments could disappear (starting on the second page) after exporting a document to PDF.

- We fixed an issue where when print multiple copies with a range of pages in a Word document, it might not work.

- We fixed an issue where some longer documents with fields might experience slower performance.

### Office Suite

- We fixed an issue where the protection prescribed by the security label and the actual protection applied to the document were misaligned.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2207: July 01

*Version 2207 (Build 15427.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Fixing how Excel “autoconverts” data entered by the user:** Give users the option to enable/disable specific data conversions as needed.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an issue that occurred while attempting to support Unicode characters in the unique identifier column.

### Excel

- We fixed an issue where, in certain cases, some parts of the background of the formula bar would incorrectly show as white when opening a workbook after starting Excel with the Start screen disabled.

- We fixed an issue where protected labels stopped working when the user reopened the app immediately after saving and closing.

- We fixed an issue where cell references in charts were displaying incorrectly.

### OneNote

- We fixed an issue that was causing the app to close unexpectedly.

- We fixed an issue where, after inserting a shape or line onto the canvas, the cursor didn't change when the shape or line was selected for moving or resizing.

### Outlook

- We fixed an issue that was preventing the Room Finder from loading in the GCC High environment.

- We fixed an issue where, if a user's default link permission type was set to "specific people," permissions on a link inserted using Outlook's Insert Link option weren't granted to the recipients as intended (and an error could be generated).

- We fixed an issue where the app would erroneously create a new group configuration for the other Exchange accounts in the profile.

- We fixed an issue where the app would close unexpectedly if a user double-clicked on the Tentative response button for a meeting from the reading pane when responses weren't requested.

### PowerPoint

- We fixed an issue where changes to custom document properties weren't always being reflected in the user interface.

- We fixed an issue where Table Style Options were sometimes being incorrectly disabled with Information Rights Management.

### Publisher

- We fixed an issue that occurred when applying 3D effects to pictures in the app.

### Word

- We fixed an issue where the UI sometimes stopped rendering on ARM64 devices.

- We fixed an issue that occurred when applying 3D effects to pictures in the app.

- We fixed an issue that was causing the app to not refresh correctly.

- We fixed an issue where the app would close unexpectedly when saving a file as an Adobe PDF.

- We fixed an issue where content was being hidden when the user double-clicked to hide white space.

- We fixed an issue where the footer added by built-in labeling incorrectly moved existing, manually added footers.

### Office Suite

- We fixed an issue where the app would close unexpectedly when using ink to do math.

- We fixed an issue where background uploads for CLP-protected files led to data loss or protection loss in some cases.

- We fixed an issue where, when working in files stored in SharePoint Online or OneDrive, the CPU memory incremented until it was consuming the machine's entire capacity.

- We fixed an issue where the ribbon options menu wasn't rendering in Excel.

- We fixed a visual glitch in Office apps that occurred when opening the Backstage by clicking the File tab.

- We fixed an issue where cropped images were displaying as blurred in Print Preview.

- We fixed an issue so that SVG formats from Adobe and Inkscape are properly recognized and inserted.

- We fixed an issue affecting the recognition and insertion of SVG files from various websites.

- We fixed an issue so that online users of the coauthoring feature would be informed when a sensitivity label that they don't have permission to access is applied by another user.

- We fixed an issue where, when opening Excel files stored on OneDrive and SharePoint Online, the app would close unexpectedly, and data loss could occur.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2207: June 17

*Version 2207 (Build 15407.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where the app would generate an error saying "The file couldn't be opened in Protected View" for certain files on certain computers.

### Outlook

- We fixed an issue where the app generated an error message saying "The attempted operation failed. An object couldn't be found" when searching an LDAP address book.

- We fixed an issue that caused the "Play on phone" button to be enabled in scenarios where it shouldn't have been enabled.

### Office Suite

- We fixed an issue where the size of characters was incorrect when using small caps with characters that include accents.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2207: June 10

*Version 2207 (Build 15402.20002)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Add table data from an image with Data from Picture:** Transcribing data from paper into Excel can be a slow and frustrating process. Wouldn't it be easier just to take a picture of the information and import it into your worksheet? Now you can, with the Data from Picture feature! To take advantage of this powerful feature, simply go to the **Data** tab and select **From Picture**, then choose the source; you can also review and correct the data, if necessary, before inserting it into your worksheet.

### Visio

- **See files others have shared with you:** It's now easier to find key documents in Visio with the Shared with Me list. Any documents shared with you will automatically show up in your list; typically, the most relevant documents to you appear at the top of the list. Note that you'll only see the specific files people have chosen to share with you, so that their other files remain secure. To experience this feature, simply click on the Home tab or the Open tab and select Shared with Me.

### Office Suite

- **Add SketchUp files to Office creations:** SketchUp is a popular 3D graphics program that makes it easy to create shareable conceptual designs, such as fully textured architectural models and other graphics used in industrial design, product design, and civil and mechanical engineering. Now, for the first time, SketchUp graphics (.skp files) can be integrated into your creations in Word, Excel, PowerPoint, and Outlook!<br />See details in [blog post](https://insider.office.com/blog/add-sketchup-files-to-office-creations)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Outlook

- We fixed an issue where a user created and shared a contact folder enabled for the recipient to edit contact information, but the folder was received with view-only permissions.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2206: June 03

*Version 2206 (Build 15330.20004)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)
[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed a performance issue where existing code from a user that previously ran without incident was taking much longer to run.

### Excel

- We fixed a visual glitch in the Formula bar.

- We fixed an issue where similar but distinct images may be handled incorrectly.

- We fixed an issue where the "Insert" and "Cancel" labels weren't visible in high-contrast themes.

### OneNote

- We fixed an issue where, after inserting a shape or line onto the canvas, the user was unable to drag or resize the item.

### Outlook

- We fixed an issue that was causing users to see multiple copies of a shared calendar rendered in certain circumstances.

### PowerPoint

- We fixed an issue where similar but distinct images may be handled incorrectly.

- We fixed an issue where the "Insert" and "Cancel" labels weren't visible in high-contrast themes.

- We fixed an issue where a blank slide was displayed when exporting images, icons, and videos to PDF.

### Project

- We fixed an issue where the user would open a project previously set to 100% complete only to find that the progress status had reverted to a lesser value.

### Word

- We fixed an issue where similar but distinct images may be handled incorrectly.

- We fixed an issue where the "Insert" and "Cancel" labels weren't visible in high-contrast themes.

- We fixed an issue where the interaction mode showed "Editing" when a document was in read-only mode.

- We fixed an issue where the user couldn't re-enable Autosave after it had been switched off.

- We fixed an issue that was preventing the application from discarding a document that had stopped responding, leading to a recurring sync process.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2206: May 27

*Version 2206 (Build 15321.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Outlook

- **Dictation toolbar redesigned for efficiency, cleaner look:** The Office dictation toolbar has been redesigned, featuring new visuals, a more responsive user interface, and a smaller size to stay out of the way of what matters—your content!

### Word

- **Dictation toolbar redesigned for efficiency, cleaner look:** The Office dictation toolbar has been redesigned, featuring new visuals, a more responsive user interface, and a smaller size to stay out of the way of what matters—your content!

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an issue impacting performance when importing .csv files into Access.

### Excel

- We resolved an issue where charts wouldn't insert in PowerPoint slideshows or Word documents when the default sensitivity label with encryption was applied.

- We fixed an issue related to shape anchoring when inserting rows into a worksheet that has right-to-left orientation.

- We fixed an issue when using Change Picture From Clipboard so that SVG content shows up properly.

- We fixed an issue related to converting icons to shapes to retain visibility after saving and reopening.

- We fixed an issue that was prompting the "No current record" error and other messages.

### Outlook

- We fixed an issue where references and reply-to headers were removed when replying to an email and the subject was changed.

### Word

- We fixed an issue where the Word Styles window/pane appeared blank.

- We fixed an issue with SVG objects so they can handle text-anchor end attributes and correctly maintain the current text position.

- We fixed an issue when using Change Picture From Clipboard so that SVG content shows up properly.

- We fixed a document protection issue where, when the exception list contained only emails, the "Find Next Region..." and "Show All Regions..." commands may stop working correctly.

- We fixed an issue where a red-line strikethrough on a deleted image didn't show up when the file was saved as a PDF with Track Changes turned on.

- We fixed an issue where an alternative font is displayed for certain special characters when exporting to PDF with the Chinese (Taiwan) Windows display language.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2206: May 20

*Version 2206 (Build 15313.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue that caused the user profile picture and the data types section in the Data tab to be missing in Excel after an Office update in the background (in the Windows lock screen).

- We fixed an issue that caused Excel to stop responding.

- We fixed an issue that caused Excel to close unexpectedly when showing live preview of a chart.

### Outlook

- We fixed an issue that caused the end time of a meeting to jump in the future (more than 12 hours ahead) when picking an end time greater or equal to 12 hours/0.5 days in meeting invites.

- We fixed an issue that caused users to see an error message when attempting to open a shared contacts folder for which they had editor or delegate permissions.

### Word

- We fixed in issue that caused users to see a memory error when they inserted a hyperlink in their documents.

- We fixed an issue that caused users to be unable to switch to Reviewing Mode in Interaction modes.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2206: May 13

*Version 2206 (Build 15310.20004)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Office Suite

- **Maximum number of responses extended to 5 million**The total number of responses allowed on an individual form or quiz has been increased from 50,000 to 5 million. This change eliminates the need to use multiple forms to collect more than 50,000 responses. Form owners can now gather up to 5 million responses on a single form, and can also export the results as a .csv file for further analysis.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We have fixed an issue in which the active worksheet would appear blank when running the code to turn off Application.ScreenUpdating in a VBA macro.

- We fixed an issue where the user couldn't update links in PowerPoint while a linked Excel file was open.

- We fixed an issue where the app would close unexpectedly when opening an .iqy file exported from a SharePoint list.

- We fixed an issue where shared workbooks in .xls format may improperly merge changes.

- We fixed an issue where an AMSI scan would cause the app to close unexpectedly.

### Outlook

- We fixed an issue when using the Replay All button resulted in an error stating the operation had failed.

- We fixed an issue that caused "No response required" to show on all forwarded meeting invites for delegates of room mailboxes when the Shared Calendar Improvements feature is enabled.

### Project

- We fixed an issue where columns within dialog boxes such as Task Information didn't appear in right-to-left languages such as Hebrew.

- We fixed an issue where Project was unable to connect to Project Server when forms based authentication was used with Azure Application Proxy.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

## Version 2206: May 06

*Version 2206 (Build 15227.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **AutoFilter faster and more efficient:** The AutoFilter function is now noticeably faster! These improvements were achieved by reducing memory usage and optimizing the calls made by the filter's comparison algorithm. The optimizations are especially noticeable on low-end devices that have less memory or slower CPU-memory throughput.

- **Optimized Excel recalculation on devices with constrained resources:** On resource-constrained devices (two cores or less and eight gigabytes of RAM or less), Excel has now by default made recalculation more optimal by running calculation on a single thread. In most cases, users should see noticeably faster calculation on these devices. Note that in some cases that require compute-intensive calculation, users may want to consider overriding this default by setting Number of calculation threads to “2.” For more details, see the Formulas section of the [Advanced Options](https://support.microsoft.com/office/advanced-options-33244b32-fe79-4579-91a6-48b3be0377c4) page on the Microsoft Support site.

- **Speeding up Formula Entry:** We've sped up the time required for entering a formula in a cell by reducing memory usage, making more efficient memory allocation calls, and optimizing the redrawing of the visible range around the cell edited. These optimizations are more noticeable on devices with slower memory or slower CPU-memory throughput, such as low-cost devices.

### OneNote

- **Take notes using just your voice with Dictate:** Using your voice to dictate content is a fast, hands-free way of putting your thoughts into a note—shown to be three times faster than typing! With Dictate for OneNote, now you can simply speak your thoughts to create content. To access this feature, just click on the Dictate microphone icon on the home tab and begin dictating your notes. [Learn more](https://support.office.com/article/2f5d1549-afe1-4abd-95ff-829a839e3d00)

### Office Suite

- **Easily access shared documents on the desktop:** The sign in to organization option allows guests to access documents that have been shared with them on their desktop Office applications, using a one-time passcode or third-party federation. This functionality, which was previously available in Office for the web, is now also an option for your Office desktop applications.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an issue where the previously created record appears blank when inserting records into a form that is bound to an ADO record set.

### Excel

- We fixed an issue where running a VBA script/add-in on a workbook with a chart sheet may cause the app to close unexpectedly.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2205: April 29

*Version 2205 (Build 15225.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Improvements to Power BI Request Access:** We've upgraded the experience for users who create PivotTables connected to Power BI datasets. Prior to the improved experience, when a user tried to refresh a PivotTable connected to a dataset that they do not have access to, nothing would happen. Now, a dialog box appears that informs the user that they do not have access to the dataset, and allows them to click a link to request access from the owner. Once access is granted, the user can go back to the PivotTable and refresh to continue their analysis.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an issue with the Access 2016 error "Your computer is out of disk space," generated when the app couldn't process large .csv files on linked tables.

- We fixed an issue so that modern charts will now show localized currency symbols when currency values are formatted.

- We fixed an issue where a reading pane could disappear when the user switched folders and was using a pinned web add-in.

### Excel

- We fixed an issue where the letter "j" wasn't being properly inserted.

- We fixed an issue where certain buttons in the ribbon wouldn't draw correctly when the window was resized.

- We are now removing duplicates from Data Validation AutoComplete dropdown lists.

- We fixed an issue that was causing Excel to consume excessive memory.

- We fixed an issue where the app would close unexpectedly as a result of PowerQuery data processing.

- We fixed an issue with LET functions where the name argument was the same as the column reference and the row was absolute.

- We fixed an issue with the "Relink Lists to New Site" dialog, where relinking to a new SharePoint list URL that exceeds 63 characters wasn't allowed.

- We fixed an issue where the close button was being disabled in the Macro Builder tab.

### Outlook

- We fixed an issue where the app was using over 500 threads for a specific task.

- We fixed an issue where users would receive an error message when clicking "Send\Receive All Folders" while working in offline mode.

### Word

- We fixed an issue with a OneDrive conflict that would occur when closing a label-protected file.

- We fixed an issue where comments were being incorrectly displayed in a loop when clicking on "previous" and "next."

- We fixed an issue where the cursor would scroll to the end of the document after deleting a paragraph mark.

- We fixed an issue where switching between Linear and Professional in LaTex equations could result in display errors.

### Office Suite

- We fixed an issue where four policies that were set to only take enable/disable status are now set to take numeric values and be written to HKCU.

- We fixed an issue related to specific example text, which would display poorly due to the rich-edit HTML ignoring the HTML small element and not resetting the character format masks at the end of a hyperlink.

- We fixed an issue that was causing a "Publish Package Failed" error during an update, which would cause a file-type association to be lost.

- We fixed an issue where users were unable to install M365 apps when PowerShell was disabled.

- We fixed an issue with OCPS client code by enabling config service for correct URLs in place of Federation Check.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2205: April 22

*Version 2205 (Build 15211.20000)*

## Version 2205: April 15

*Version 2205 (Build 15209.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Improved Recommended PivotTables experience:** Recommended PivotTables are now more intelligent and easier to use! The dialog box interface has been replaced by a redesigned panel, making it easier to view all of your options and simpler to change your data selection before inserting a recommended PivotTable.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where map charts would get recolored when adding labels.

- We fixed an issue with a problem that could occur when using the TEXTSPLIT function with certain arguments.

- We fixed an issue where the app stopped responding when referencing a defined name containing the formula =IF(1) in a cell.

### Outlook

- We fixed an issue where the app stopped responding when trying to respond to a specific contact, after calling ResolveContacts thousands of times.

### PowerPoint

- We fixed an issue where multi-line sensitivity label footers would be positioned so that only the first line would appear at the bottom of the slide.

- We fixed an issue where the user sees a blank block when inserting another PowerPoint file as a linked object.

### Word

- We fixed an issue with Modern Comments in Word where text copied from OneNote into a comment would insert an image.

- We fixed an issue so that an icon within a table is displayed properly when being resized.

- We fixed an issue where the Trust Center settings windows wouldn't open after posting a new Modern Comment in Word.

- We fixed an issue where, when closing a comment, pending bookmarks also closed.

### Office Suite

- We fixed an issue where customers with the updated visuals in Backstage wouldn't see their chosen Office background in certain apps and themes.

- We fixed an issue where, when a user paused the Read Aloud function, the app would skip to the next paragraph when restarted.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2205: April 08

*Version 2205 (Build 15130.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an issue where unintended line breaks could occur for languages with IVS characters.

### Excel

- We fixed an issue with selections not appearing or not visible after being chosen in some dropdown menus, especially in Dark Gray theme.

- We fixed an issue where similar but distinct images may be handled incorrectly.

- We fixed an issue that could cause an error when trying to navigate to a document stored in SharePoint.

### PowerPoint

- We fixed an issue with selections not appearing or not visible after being chosen in some dropdown menus, especially in Dark Gray theme.

- We fixed an issue where similar but distinct images may be handled incorrectly.

### Word

- We fixed an issue with selections not appearing or not visible after being chosen in some dropdown menus, especially in Dark Gray theme.

- We fixed an issue where similar but distinct images may be handled incorrectly.

- We fixed an issue where a merge conflict alert didn't offer the option of discarding the current changes.

### Office Suite

- We fixed an issue with Read Aloud that would cause the app to close unexpectedly.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2204: April 01

*Version 2204 (Build 15128.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue that caused Excel to consume excessive memory.

### Project

- We fixed an issue that prevented Gantt chart-type views from fully displaying; this could occur when the view had a drawing object applied to it.

### Word

- We fixed an issue where Live Preview was using a lower zoom level than it was set to.

- We fixed an issue where the app would sometimes stop responding when a repeating content control was added while Track Changes were turned on.

### Office Suite

- We fixed an issue with displaying line and fill properties in SVG files.

- We fixed an issue that would prevent Record Slide Show from working when animation timings were involved.

- We fixed an issue with sessions being disconnected from RDS session host.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2204: March 25

*Version 2204 (Build 15121.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an issue that could cause the app to close unexpectedly if memory requirements grow too rapidly for applications using DAO or OLEDB interfaces to read Access databases.

### Excel

- We fixed an issue with newly added labels on unprotected files being lost if opened too soon after saving.

- We fixed an issue where the Power BI data type was missing under Get Data (GCC tenant only).

- We fixed an issue with OnPrem where Get Data would become disabled in a document with a sensitive label and all permissions except full control.

- We fixed an issue related to selecting items in the online premium content library.

- We fixed an issue that caused the app to close unexpectedly when a chart was inserted by a Recommended Chart operation.

- We fixed an issue where a tooltip for the See All Charts button was misleading (“see all chart types”); the behavior was actually limited to Recommended Charts.

### Outlook

- We fixed a responsiveness issue when viewing animated GIFs.

- We fixed an issue that would cause folder items outside of the sync window to be lost when the folder is moved to the Online Archive.

- We fixed an issue that prevented a subset of messages from being moved when moving multiple mail items from the Focused to Other inbox or the Other to Focused inbox.

- We fixed an issue that would occur when replying to a message in a shared mailbox when the email sender is also on the To: or CC: line. The user's email address would be removed from the recipient list on the reply, resulting in the primary email sender not receiving any replies to the thread.

### PowerPoint

- We fixed an issue related to the selection of text in a shape with 3D properties applied.

- We fixed an issue related to selecting items in the online premium content library.

- We fixed an issue to make paste behavior consistent in Edit view and Presenter view.

### Project

- We fixed an issue that affected the display of customized views in Project.

- We fixed an issue where progress lines weren't being displayed when performing a Save As PDF command.

- We fixed an issue where the user was unable to open certain projects; the app would close unexpectedly instead.

- We fixed an issue where an asterisk (*) used as a separator in filters caused errors and unexpected behavior; replaced by a pipe (|).

- We fixed an issue where the app stops responding when leveling doesn't complete.

- We fixed an issue where the application opens to a screen that is no longer available, making the project hidden.

### Word

- We fixed an issue related to selecting items in the online premium content library.

- We fixed an issue where SVG files may be displayed as a broken link.

- We fixed an issue where GIFs with zero frame delay wouldn't animate when inserted into an email.

- We fixed an issue where the wrong font color was showing for comments inserted on Mac and in different themes on Windows.

- We fixed an issue where sections of text changed color to white after pasting a table from Excel and saving filtered HTML.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2204: March 18

*Version 2204 (Build 15109.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an issue that could cause the application to close unexpectedly when using the Access Database Engine OLEDB API with a database containing links to SharePoint lists.

### Excel

- We fixed an issue where a recovered file would be opened as read-only.

- We fixed an issue where the app could close unexpectedly when trying to change a link via the Edit Link dialog.

- We fixed an issue where the app wasn't translating named formats from Power BI.

- We fixed an issue where choosing the search command in context menus disabled common keyboard shortcut combinations.

### Outlook

- We fixed an issue where resending a message reverted the message body to plain text and all changes were lost (data loss).

- We fixed an issue affecting RMS-protected messages.

- A data change was made to include a new United Kingdom holiday in UK-based calendars.

- We fixed an issue that could cause users to be unable to disable the Conversation view.

- We fixed an issue where removed shared calendars reappeared when the REST shared calendar feature was enabled.

### PowerPoint

- We fixed an issue regarding the handling of multiple contiguous spaces properly within SVG files.

- We fixed a stability issue related to exporting a recorded presentation.

### Project

- We fixed an issue on systems with high DPI where nodes on a network diagram were viewed as overlapping or not to scale and didn't print correctly.

- We fixed an issue where saving or publishing a project from Project Client doesn't save changes to custom fields.

### Word

- We fixed an issue that disabled collaboration mode for a document.

- We fixed an issue where Word couldn't print visible optional hyphens (Alt+173) when used in equations.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2203: March 11

*Version 2203 (Build 15028.20022)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Faster filtering when cells contain unique or duplicate rules:** When your workbook contains many unique or duplicate conditional formatting rules, it can often slow down the app's performance. No longer! By optimizing the underlying comparison algorithm, we've enhanced the performance and sped up the filtering process.

### PowerPoint

- **Modern comments now available to consumer audience:** Modern comments in PowerPoint offer many enhancements that improve the collaboration process on presentation, including comments anchored to specific text, comments visible in the margin and in the comments pane, the ability to resolve threads, enhanced @mentions functionality, and more. Previously limited to commercial licenses of Office, modern comments are now available to the PowerPoint consumer audience as well!

### Word

- **Improved Coauthoring Error Recovery Experience:** Collaborating with others while working in Word is a vital productivity tool for many users, and disruptions can be extremely frustrating. We've now developed an improved recovery experience to quickly restore users back to a connected state after experiencing coauthoring errors. This automatic refresh synchronizes all changes among different authors, so you're seeing the most up-to-date document possible. Conflicts between different authors and changes not saved to the server will show as tracked changes.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an issue that caused the Display Name and Trendline Name for a chart data series to be unable to be edited in the Chart Settings pane.

### Excel

- We fixed an issue where the text fields in the custom filter dialog would autocomplete when you start typing a value.

- We fixed an issue where, when typing double-byte characters and there are two suggestions, one would be automatically inserted.

- We fixed an issue where the button form control would lose its position after inserting or deleting rows or columns.

### Outlook

- We fixed an issue where the app would stop responding due to a missing writer thread.

- We fixed an issue where a scroll-down button was missing in the "From" account picker dropdown menu.

- We fixed an issue where the default icon was displayed for attached messages when a custom icon was expected.

- We fixed an issue that caused users to experience performance issues when switching folders due to a corrupt view setting.

- We fixed an issue that caused the membership list to not display properly for contact cards.

### PowerPoint

- We fixed an issue where the default sensitivity label wasn't getting properly applied when creating a new document in PowerPoint.

- We fixed an issue that would cause Excel, Word, and PowerPoint to occasionally close unexpectedly when the user changed the sensitivity label for a document.

### Project

- We fixed an issue so that users should no longer experience problems when printing Progress Lines in the Gantt chart.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2203: February 25

*Version 2203 (Build 15018.20008)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where text appeared in a blank cell next to a data type image when the zoom level increased.

- We fixed an issue where using custom command bars could cause Excel to close unexpectedly.

- We fixed an issue where the tooltip for a sensitivity label was displayed too far from the label itself.

### Outlook

- We fixed an issue that caused users to see a one-off task appear unexpectedly after deleting an occurrence of a recurring task.

- We fixed an issue where rules for outgoing mail sent to a specific person wasn't working properly in Cached Exchange Mode.

- We fixed an issue where large images wouldn't appear in a message unless it was viewed in Draft mode.

### PowerPoint

- We fixed an issue where PowerPoint wouldn't show the company name under File > Info.

- We fixed an issue where the tooltip for a sensitivity label was displayed too far from the label itself.

### Word

- We fixed an issue where the tooltip for a sensitivity label was displayed too far from the label itself.

- We fixed an issue where large images wouldn't appear in a message unless it was viewed in Draft mode.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2203: February 18

*Version 2203 (Build 15012.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue to correct how the counties of Spain were appearing in map charts.

- We fixed an issue that was preventing the user from being able to paste a copied cell into the text field in the Custom Filter dialog.

- We fixed an issue to align how gradients fill on a series in a waterfall chart with how they fill on a series in a column chart.

### Outlook

- We fixed an issue where a received email wouldn't include a link to open the message.

### PowerPoint

- We fixed an issue where certain text color operations were incorrectly disabled when a document was encrypted with rights management.

### Project

- We fixed an issue where, when using the Physical % Complete tracking method, calculated BCWP values greater than 10,000,000,000 on a task weren't possible.

- We fixed an issue where the user would see a security dialog stating that the project had links to one or more data sources, even though the project had no active links. Now, the dialog appears only when there are active links.

- We fixed an issue where there was intermittent data loss when saving for various baseline fields.

- We fixed an issue where the save operation to a network drive would become very slow when using the Save As PDF (*.pdf) file type.

### Word

- We fixed an issue where the app would intermittently stop responding when checking for updates.

- We fixed an issue that was causing SVG images that contained external content to not show up in some cases.

- We fixed an issue with Track Changes where revision marks weren't working as expected.

- We fixed an issue where invoking Repeat through the user interface, or VBA after hitting Enter, would insert the whole current paragraph in addition to the newly inserted paragraph.

- We fixed an issue where the Save As dialog would appear unexpectedly while editing a document.

### Office Suite

- We fixed an issue that reduced the number of SVG image saves that didn't work.

- We fixed an issue where the user couldn't open a file in Application Guard.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2203: February 11

*Version 2203 (Build 15003.20004)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an issue where Power BI would sometimes fail to refresh data from an Access database if there were links to SharePoint lists in the database.

### Project

- We fixed an issue where local formatting settings were being overwritten when a user saved a server-based project as an MPP file.

### Word

- We fixed an issue where a file couldn't be opened in the Microsoft 365 version but could be opened in earlier versions of Word.

### Office Suite

- We fixed an issue that caused contact cards to fail to show.

- We fixed an issue with the Outlook preview pane that was preventing SVG images from rendering properly.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2202: February 04

*Version 2202 (Build 14931.20010)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where clicking the Accessibility checker pane no longer switches focus to the Accessibility ribbon if the user has navigated away from the ribbon.

- We fixed a typo in an alert that would appear when a formula was too long.

- We fixed an issue that would occur when there were hidden fields in a PivotTable.

- We fixed an issue where the menu icons in the ribbon looked misplaced.

### Outlook

- We fixed an issue where Outlook wouldn't sync the complete folder hierarchy for mailboxes with a large number of folders (e.g., more than 10,000 folders).

- We fixed an issue where clicking on multiple users' contact cards in the organization chart view caused the app to stop responding.

- We fixed an issue where a delegate was unable to tell which manager's shared calendar they were in when opening the calendar from a meeting invite in the new REST-based calendar sharing model.

- We fixed an issue where the Join button wasn't available when one instance of a meeting series was changed to a Teams meeting.

- We fixed an issue where the user would lose data when deleting an "empty" folder that contained data on the server side; a warning will now appear in this scenario.

- We fixed an issue (affecting 64-bit client machines only) where the app stopped responding when forwarding an RTF email.

### PowerPoint

- We fixed an issue where the menu icons in the ribbon looked misplaced.

### Project

- We fixed an issue so that users can now save projects to the Project Web Application, even if the resources in the offline file have names that match enterprise resources.

- We fixed an issue where users were getting into a state where saving wouldn't work. This fix allows users to always be able to save their work.

### Word

- We fixed an issue where a mail merge from an Outlook contact could generate an "Out of memory or system resources" error.

- We fixed an issue where Table of Contents page numbers changed to a single digit when the document was in Print Preview mode.

- We fixed an issue where the menu icons in the ribbon looked misplaced.

### Office Suite

- We fixed an issue that caused text to render incorrectly when using the small caps feature.

- We fixed an issue where, after the user received access to the visual refresh, a one-time popup (upon initial opening of Office apps) wasn't enabled as expected per the Coming Soon feature's group policy setting.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2202: January 26

*Version 2202 (Build 14922.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where the formatting of filtered charts would change after being saved in Excel for the web.

### Outlook

- We fixed an issue where, if a user changed the From field with no alterations made to the new compose, the label justification prompt was incorrectly shown.

- We fixed an issue where Outlook was merging contact information when the RunContactLinking regkey was set to 0.

- We fixed an issue where the Appointment Quick View (AQV) was removed from the hidden window check to enable initial scrolling.

### Project

- We fixed an issue where autofilters didn't work if the text information in an enterprise text type field contained a question mark (?) or asterisk (*). The error generated was, "The entry isn't valid. The test value cannot be used with the field for the data you want to find or filter for." Autofilters now allow the inclusion of values with reserved characters.

### Word

- We fixed an issue by changing non-white background colors that exactly matched highlight colors to have the highlight property.

- We fixed an issue where, upon selecting a specific page using the navigation pane, only the first part of the page was displayed in print layout mode.

- We fixed an issue where performance with large documents was slower in Word 365 or Windows 10 than in Word 2013 or Windows 7.

- We fixed an issue that updated the case statements to correspond to the data size.

### Office Suite

- We fixed an issue in existing WOPI protocol to support real-time coauthoring, where if a client lock on a document expires when attempting to refresh it, the client will acquire the lock.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2202: January 21

*Version 2202 (Build 14912.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Improved Power BI Connected PivotTables:** This work brings improvements to PivotTables connected to Power BI datasets, including the ability to create aggregations in Excel by dragging fields into the Values area.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an issue where some SQL Server drivers would show the error "String data, right truncation (#0)" when a user tried to add more than 8,000 bytes worth of data to a varchar(max) field.

### Outlook

- We fixed an issue where the list of folders in Outlook stopped responding to mouse clicks after a contact profile card was opened.

- We fixed an issue where a hyperlinked URL didn't change automatically when there was a change in the shared OneDrive link using Outlook.

### Word

- We fixed an issue where an error appeared on closing any document containing a list style with the "New documents based on this template" option enabled.

### Office Suite

- We fixed an issue where some default applications disappeared after the monthly Office update.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2202: January 14

*Version 2202 (Build 14907.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Data Validation dropdown list autocomplete:** Dropdown lists are a handy way to make data entry and validation more efficient in Excel. We've now added AutoComplete functionality, which automatically compares the text typed in a cell to all items in the dropdown list and displays only the items that match. You'll spend less time scrolling through lists, dealing with data validation errors, or writing complex code to handle this task.

### Outlook

- **Send emails from your account alias or proxy email address:** Outlook has traditionally supported receiving email at addresses other than your default address (known as a proxy address, or alias). Now you can send mail from those proxy accounts as well, by choosing the desired outgoing address.

### Word

- **Improved Coauthoring Error Recovery Experience:** Collaborating with others while working in Word is a vital productivity tool for many users, and disruptions can be extremely frustrating. We've now developed an improved recovery experience to quickly restore users back to a connected state after experiencing coauthoring errors. This automatic refresh synchronizes all changes among different authors, so you're seeing the most up-to-date document possible.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an issue that would prevent multiple users from opening a database when using network paths that include DFS Namespaces, short file names, or mapped drives. [Learn More](https://support.microsoft.com/office/error-in-access-when-opening-a-database-on-a-network-file-share-6cbc1560-62c2-46e7-9980-d079a46f5acc)

### Excel

- We fixed an issue with SVG files exported from Visio not displaying text correctly.

### Outlook

- We fixed an issue with Outlook Search that caused incomplete results when using the OnPrem service search.

- We fixed an issue where a message using Digital Rights Management (DRM) couldn't be opened.

### PowerPoint

- We fixed an issue where linked images might not load when opening a presentation programmatically in window-less mode.

- We fixed an issue with SVG files exported from Visio not displaying text correctly.

### Word

- We fixed an issue with SVG files exported from Visio not displaying text correctly.

- We fixed an issue that caused the app to close unexpectedly when selecting Go to Comment for the first comment in a document.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2202: January 07

*Version 2202 (Build 14901.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Project

- We fixed an issue where the Outline Grouping filtering wasn't working.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2201: December 31

*Version 2201 (Build 14822.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where Excel was unable to export workbooks to XPS if the user didn't have export permission.

- We fixed a performance issue when using automation-based tools, including slicers and accessibility tools.

### Outlook

- We fixed an issue where removed shared calendars reappeared when the REST shared calendar feature was enabled.

### Word

- We fixed an issue where comments were missing in Sidetrack but visible in the pane.

- We fixed an issue where zoom on live preview dismissed the contextual card.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2201: December 24

*Version 2201 (Build 14816.20004)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Office Suite

- **Haptics Feedback:** Enables Haptics feedback for the pen.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an issue that could cause an application to stop responding when connecting to an Access or Jet database using multiple threads in rapid succession.

### Outlook

- We implemented a design change request to prevent the Automatic Replies banner from appearing.

- We fixed an issue where Outlook was merging contact information when the RunContactLinking regkey was set to 0.

- We fixed an issue where Outlook would stop responding when navigating to a shared calendar.

### PowerPoint

- We fixed an issue where Reset Slide didn't respect specific picture fill properties in some cases.

### Project

- We fixed an issue that would cause the app to stop responding and resulted in data loss.

- We fixed an issue where time-phased data didn't appear for Material Resources when the assigned task's duration was set to zero after the task was completed.

- We fixed an issue where upgrading projects migrated from Windows Server 2008 to Windows Server 2016/2019 stops working on save.

### Word

- We fixed an issue where there was a potential for data loss without error when a user carried out a Save As action to a location using a server path in a CICO library synced by the OneDrive client.

- We fixed an issue where contact cards weren't working in Office.

- We fixed an issue where values in a table sometimes didn't refresh when entering new data.

### Office Suite

- We fixed an issue where the app stopped responding after an update to Office version 2109 (current channel).

- We fixed an issue where the user was unable to open files from the Shared with Me section on the app home page.

- We fixed an issue where Office shows an account error in the SSO and ADFS DRS environment.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2201: December 17

*Version 2201 (Build 14809.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an issue that could cause applications using the Access Database Engine OLEDB API with a database containing links to SharePoint lists to close unexpectedly.

### Excel

- We fixed an issue where AutoSave could be temporarily disabled after applying a sensitivity label to protect the document.

- We fixed an issue where you would be unable to select a value from a data validation dropdown list in a cell if that list contained blank values.

- We fixed an issue where the search results would get lost in the pivot table field list taskpane.

- We fixed an issue where, in a multi-monitor setup, some data in dialog boxes was being hidden from the user when selecting cells.

- We fixed an issue where, when you had a Microsoft Excel 97-2003 Worksheet object embedded inside another application (such as a Word document), using the Convert feature to convert it to a Microsoft Excel Worksheet (Office OpenXML) object didn't complete the conversion until you opened the embedded object and made a change to it. The object is completely converted when using the Convert feature now.

- We fixed an issue related to text anchoring in SVG rendering.

- We fixed an issue where an application using the OLEDB API with the ACE.OLEDB.12.0 or ACE.OLEDB.16.0 provider was closing unexpectedly.

### OneNote

- We fixed an issue where a normal lasso selection on tablet emulators would cause the app to close unexpectedly.

### Outlook

- We fixed an issue that was causing users to see garbled text in some fields when exporting contacts to a CSV.

- We fixed an issue where the sender of a mail wasn't being included when replying all, when the From address was different from the Reply To address.

- We fixed an issue that caused the app to close unexpectedly when opening a message via a reminder, when the user's download preference was set to Download Headers.

- We fixed an issue that caused users to be unable to add a shared calendar from their contacts with "Shared Calendar Improvements" enabled.

### PowerPoint

- We fixed an issue with selecting text via double-tap when using the touchpad.

- We fixed an issue related to text anchoring in SVG rendering.

- We fixed an issue that caused an error when inserting Vimeo videos into a presentation.

### Word

- We fixed an issue related to text anchoring in SVG rendering.

- We fixed an issue where the app closed unexpectedly when the user clicked on or used shortcut keys to turn the Read Aloud feature on and off repeatedly.

### Office Suite

- We fixed an issue where Microsoft Purview Information Protection sensitivity labeling audit data was no longer generated if the EnableAudit setting was turned off.

- We fixed an issue where the app stopped responding when Read Aloud was started and stopped in quick succession.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2112: December 10

*Version 2112 (Build 14729.20038)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Project

- We fixed an issue where the application would close unexpectedly when loading a customized report.

- We fixed an issue where manually scheduled tasks were rescheduled to earlier dates when the user opened projects that had been Saved As with a different name.

### Word

- We fixed an issue where the cursor disappears from a comment while scrolling in a multi-page document.

- We fixed an issue where a GUID was displayed rather than an author's name during collaboration.

- We fixed an issue where the application sometimes stopped responding when refreshing the table of contents.

- We fixed an issue where the Undo command wasn't working after changing the color of bullets in a document.

### Office Suite

- We fixed an issue where an Upload Blocked warning window appeared when a new file with an Encrypted Label was saved and re-opened in an environment where OneDrive Sync had been configured.

- We fixed an issue where a reboot was triggered to complete the removal of existing installations.

- We fixed an issue where Read Aloud would close unexpectedly when it was started and stopped in rapid succession.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2112: November 26

*Version 2112 (Build 14718.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where opening xlsm file in SpreedsheetCompare tool can cause the tool to stop responding.

- Refreshing the data for Pivot Tables can stop working when filtered values no longer exist in the data. A subsequent query statement generated without the invalid filtered values to retry the refresh request was temporarily disabled and is now reenabled.

- We fixed this regression that was a result of Mica on Windows 11 being turned on when Wincomp is off.

- We fixed an issue with the WEBSERVICE worksheet function, where Excel would stop working in rare cases when calculations are canceled by the user.

### Word

- We fixed an issue where the alignment of a previous comment changed from right to left when posting a new comment with right alignment.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2112: November 19

*Version 2112 (Build 14712.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### PowerPoint

- **Record:** Make your presentations more impactful by recording videos with narration. [Learn more](https://support.office.com/article/ddc4432c-79f6-4add-b85e-1009815d955c)

- **Export:** Bring all the components of the presentation together for easy sharing and viewing. Exported video includes all recorded timings, narrations, ink, and laser pointer gestures. Video also preserves animations, transitions, and media.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue so that the default control is the text field, and users can start typing as soon as the dialog opens.

- We fixed an issue where interacting with form controls could cause Excel to close unexpectedly.

### Word

- We fixed an issue where pasting text in a comment would unnecessarily add a new line.

### Office Suite

- We fixed an issue where the app would close unexpectedly upon opening a file.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2112: November 12

*Version 2112 (Build 14706.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where the filter dropdown arrows stopped working when the sheet was zoomed in.

- We fixed an issue where newly created PivotTables could lose custom settings if the data source range was changed.

- We fixed an issue where certain formula results weren't updated after their inputs had changed.

- We fixed an issue with the LOGEST worksheet function, where a transient overflow error wasn't handled and cleared, and was incorrectly picked up for unrelated subsequent calculations.

### Outlook

- We fixed an issue where received emails that users sent to themselves (the same email address) rendered a blank message body.

### Project

- We fixed an issue where progress lines aren't being drawn correctly when a summary task is at 0% complete.

- We fixed an issue where, when tasks are rescheduled in Project, manually scheduled tasks may be scheduled earlier than they should be.

### Word

- We fixed an issue where the cursor doesn't appear in the reply box when inserting a new comment if any pane is open on the right side of the document.

### Office Suite

- We fixed an issue where explanations for TrialNotificationBarDismissed, DateHigh, and DateLow DWORDS were removed.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2111: November 05

*Version 2111 (Build 14630.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### PowerPoint

- **PowerPoint Accessibility Ribbon:** Making your presentation accessible to people with disabilities requires knowledge, compassion, and special tools. The new Accessibility ribbon in PowerPoint helps you accomplish this by bringing all the tools you need together in one place.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where the original icon of the app wasn't displaying correctly if the source file was from a OneDrive location and the file name contained HTML encoding characters when adding an Excel, Word, or PowerPoint embedded object into an Excel file.

- We fixed an issue where apply name for defined name wasn't working for dynamic array.

- We fixed an issue where the text in an Excel cell  be seen in live preview under certain conditions.

- We fixed an issue where switching sheets may cause the sheet tabs to not display properly.

- We fixed a visual glitch when clicking and dragging to select a range of cells.

### Outlook

- We fixed an issue where the calendar experienced problems when doing a full sync.

- We fixed an issue where saving Outlook messages to a SharePoint document library would generate an error, due to an Internet connectivity test.

- We fixed an LDAP Address Book search error generated when there were two LDAP Address Books from different domains.

### PowerPoint

- We fixed an issue where a Save As operation with AutoSave off saved the changes to the original cloud file rather than creating a new version.

### Project

- We fixed an issue where paste links weren't updating as information in the project changed.

### Word

- We fixed an issue where writing a comment that mixes language directions (like English and Hebrew) resulted in the word order appearing incorrectly.

- We fixed an issue with large URLs where a link couldn't be opened if its length exceeded a specific character limit.

- We fixed an issue of a persisting zoom level when relaunching and opening documents saved at different zoom levels.

- We fixed an issue related to page width when closing the comments pane with page width zoom selected.

- We fixed an issue where anchor highlights aren't removed from a modern comment when the cursor moves off of it.

- We fixed an issue where Word goes unresponsive when selecting and right-clicking a content control locked against deletion.

### Office Suite

- We fixed an issue where Excel was unable to open a workbook saved from a workbook opened with Application Guard.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2111: October 29

*Version 2111 (Build 14623.20002)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where importing a query, then selecting a range in the new worksheet and loading from that range, would cause the query editor to open with the wrong table name.

- We fixed an issue causing rendering errors when teleporting great distances in the grid (such as when using Ctrl + arrow key).

### Outlook

- We fixed an issue where the wrong end date for an all-day event was appearing in Calendar view.

- We fixed an issue where the Windows Explorer would close unexpectedly when doing a “File - Send-To – Mail recipient.”

### PowerPoint

- We fixed an issue where a printout would be cut off (data loss) when the user changed a document in portrait orientation from a bigger paper size to a smaller paper size.

- We fixed an issue so the app will no longer prompt for a password to modify it if the user explicitly opens the file in read-only mode.

### Word

- We fixed an issue that caused auto-scroll mode to not function.

- We fixed an issue with co-authoring mode where a revision might not be saved.

- We fixed an issue where a graphic could be missing in certain documents.

- We fixed an issue where certain font styles weren't being mapped appropriately when pasting text as "Keep text only."

- We fixed an issue where a mail-merge formatted document couldn't be saved in .doc format.

### Office Suite

- We fixed an issue where the PowerPoint app would stop responding and the error message "ERROR_TOO_MANY_FILES" would appear.

- We fixed an issue so that Eyedropper functionality is enabled in documents opened with limited permissions.

- We fixed an issue in which animated GIFs cropped to a non-rectangular shape wouldn't animate in PowerPoint's Slide Show presentation mode.

- We added an option to the Record Slide Show dialog to remove the prompt to Export to Video on closing.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2111: October 22

*Version 2111 (Build 14613.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where a query update caused Excel to stop responding.

### Outlook

- We fixed an issue where Like Button didn't appear for Group emails if user opens an Inbox email and closes it before going to Groups.

- We fixed an issue where Outlook would display fewer contact lists/contact groups in suggestions than expected.

### PowerPoint

- We have fixed an issue where copying an embedded object caused the embedded object to save, as well as re-saving the containing document (which created a new version if AutoSave was enabled).

### Word

- We fixed an issue where some tables spanning multiple pages would bounce up and down (due to redrawing) in a no-margin page view.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2111: October 15

*Version 2111 (Build 14609.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Word

- **Improved Coauthoring Error Recovery Experience:** An improved recovery experience to quickly restore users back to a connected state after experiencing coauthoring errors.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where the word "Column" disappeared from the Cell Format ribbon dropdown menu on mouse over.

- We fixed an issue where localized characters were rendering too small in a worksheet tab (issue only present for users with the Fluent UI Coming Soon toggle enabled).

- We fixed an issue related to a workbook with many custom views with Freeze Panes that was causing Excel to stop responding immediately after launch.

- We fixed an issue where the user was seeing the Music data type button but the artists weren't converting.

- We fixed an issue where hitting the hot key "e" for Filter would conflict with Search on the context menu.

### Outlook

- We fixed a sync failure that was occurring when generating a preview.

- We fixed an issue where the app would stop responding when conducting a search.

- We fixed an issue where default sensitivity labels weren't getting applied on encrypted emails.

- We fixed an issue where an invite would go out to attendees but wouldn't be saved to the Organizer's calendar.

- We fixed an issue where an "Encrypt-Only" sensitivity label wouldn't process for all UI languages except EN.

- We fixed an issue that caused the Appointment quick view to be cropped when previewing meeting invitations.

### Project

- We fixed an issue where, when programmatically adding new tasks to a project via CSOM, the tasks might not be inserted at the correct location if the new task's summary task is collapsed.

### Word

- We fixed an issue where the app would stop responding when "at mentioning" (@mention) someone in a comment.

### Office Suite

- We fixed an issue where a "wdlor" + GUID query parameter could get added to the end of a link.

- We fixed an issue affecting the MSI Office 2007 catalyst detection logic, which was causing Visio and Project to be unintentionally removed.

- We fixed an issue where corrupted SVGs in Office documents fail to render, showing a red X, by substituting an uncorrupted bitmap version of the image.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2111: October 08

*Version 2111 (Build 14530.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Outlook

- **Update Photo Link in Contact Card:** Shows the Update Photo link in your own contact card

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where using a cell reference in a formula caused users to experience high CPU usage.

### PowerPoint

- We fixed an issue related to reordering bulleted list items using drag and drop.

### Word

- We fixed an issue where the application would stop responding when the user selected File > Save As.

- We fixed an issue where the app would close unexpectedly when the user clicked on comment text while dictating a reply to that comment.

- We fixed an issue where the Read Aloud playback sometimes jumps to a random location in the document.

### Office Suite

- We fixed an issue where an update would stop responding. The update would open Software Center as expected; it would start the "download and apply" phase, but then would stop responding and an error would appear: 0x87d00668 ("Software update still detected as actionable after apply").

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2110: October 01

*Version 2110 (Build 14527.20040)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where the row and column header colors on partial selection weren't discernible from the fishbowl color in Dark Gray theme (issue only present for users with the Fluent UI Coming Soon toggle enabled).

- We fixed an issue where users reported it was difficult to tell the selected tabs from the non-selected tabs in the Office light themes (issue only present for users with the Fluent UI Coming soon toggle enabled).

- We fixed an issue in the Excel macro language in which the alert dialog box didn't show in the proper type.

- We fixed an issue with number formats used in Data Type property values for non-English regional system settings.

- We fixed an issue where, in some cases, the top rows could appear duplicated in worksheets with Freeze Panes enabled.

- We fixed an issue where the BizUD Gothic Japanese font was rendering differently between Office applications.

### Outlook

- We fixed an issue that caused users of Outlook's Shared Calendar Improvements feature to experience high CPU usage.

- We fixed an issue where a delegate  see organizer details in Scheduling Assistant.

### PowerPoint

- We fixed an issue where the sub-menu opened in the wrong place in a slideshow.

### Word

- We fixed an issue where the document was appearing and disappearing in a flicker motion.

- We fixed an issue where emails forwarded with long image URLs wouldn't display.

- We fixed an issue where the BizUD Gothic Japanese font was rendering differently between Office applications.

### Office Suite

- We fixed an issue where Excel now has more control about the input events and users can explicitly decide when they want to start a PTP gesture.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2110: September 24

*Version 2110 (Build 14517.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where Excel's enclosed numeric characters show as question marks.

### Outlook

- We fixed an issue where the sensitivity tab was disabled in the frame-control window for some users.

- We fixed an issue where search didn't include items from an online archive mailbox.

### Project

- Fixed an issue where the beforetaskchange event would fire twice and included incorrect information when the predecessors field changed.

### Word

- We fixed an issue where changes failed to sync and progress was lost in both the synced file and the locally backed up file.

- We fixed an issue where Word would become slow to respond while using a high CPU percentage.

- We fixed an issue where autocorrect wasn't working in Modern Comments.

### Office Suite

- We fixed an issue where a user can't open an .xls/.ppt/.doc file in a folder that was synced from a SharePoint folder with read permission.

- We fixed an issue in Excel where the error ERROR_DISK_FULL would appear before maintenance would begin.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2110: September 17

*Version 2110 (Build 14509.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Default sensitivity labels now apply to existing unlabeled documents:** Fixed an issue where the default sensitivity label wouldn't be applied to existing unlabeled documents.

### PowerPoint

- **Default sensitivity labels now apply to existing unlabeled documents:** Fixed an issue where the default sensitivity label wouldn't be applied to existing unlabeled documents.

### Word

- **Default sensitivity labels now apply to existing unlabeled documents:** Fixed an issue where the default sensitivity label wouldn't be applied to existing unlabeled documents.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue that caused an insufficient memory warning to appear when copying and pasting content.

### PowerPoint

- We fixed an issue where ink wouldn't be displayed when a PowerPoint slide was pasted in another program.

### Project

- We fixed an issue that caused the Visual Basic Applications (VBA) OrganizerMoveItem method that is used to move custom field information from one project to another to not work properly when the Name parameter is omitted.

### Word

- We fixed an issue that caused the Save indicator to stop responding.

### Office Suite

- We fixed an issue that negatively impacted the typing speed in an email or document when an animated GIF was playing.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2110: September 10

*Version 2110 (Build 14503.20002)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **DLP policy tips in Word/Excel/PowerPoint:** Additional sensitive information types configured as part of OneDrive and SharePoint data loss prevention (DLP) policies can now be detected by the app to show a policy tip. This update also brings accuracy improvements and globalization support. [Learn more](/microsoft-365/compliance/sensitive-information-type-learn-about)

### Outlook

- **Use your voice to send email and @mention people:** In our increasingly busy world, dictating your emails in Outlook has become a very popular way to improve your efficiency. We've now made this feature even more powerful by adding specific voice commands, enabling you to add people to an email, mention (@name) someone in a message, and send the mail—all using only your voice.<br />See details in [blog post](https://insider.office.com/blog/use-voice-commands-to-speed-up-email-dictation-in-outlook)

### PowerPoint

- **DLP policy tips in Word/Excel/PowerPoint:** Additional sensitive information types configured as part of OneDrive and SharePoint data loss prevention (DLP) policies can now be detected by the app to show a policy tip. This update also brings accuracy improvements and globalization support. [Learn more](/microsoft-365/compliance/sensitive-information-type-learn-about)

### Word

- **DLP policy tips in Word/Excel/PowerPoint:** Additional sensitive information types configured as part of OneDrive and SharePoint data loss prevention (DLP) policies can now be detected by the app to show a policy tip. This update also brings accuracy improvements and globalization support. [Learn more](/microsoft-365/compliance/sensitive-information-type-learn-about)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where automatic sensitivity labeling wasn't working for a few GCC-H tenants.

### Outlook

- We fixed an issue where double-clicking to save an untrusted attachment would fail to save to network locations.

- We fixed an issue where messages created via Send To do not get default sensitivity labels.

### PowerPoint

- We fixed an issue where attempting to turn on AutoSave for files with unsupported extensions saved to OneDrive or SharePoint would show the Share dialog.

### Project

- We fixed an issue where, if the decimal separator isn't a period, enterprise resources can't be saved when an enterprise number custom field is updated.

### Office Suite

- We fixed an issue where the ampersand character was incorrectly shown as an underline in the data types card.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2109: September 03

*Version 2109 (Build 14430.20030)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Word

- **Office apps now support OpenDocument Format (ODF) 1.3:** ODF 1.3 brought many improvements to the OpenDocument format and these are now supported in Word, Excel, and PowerPoint (file extensions .odt, .ods, and .odp).<br />See details in [blog post](https://insider.office.com/blog/office-apps-now-support-opendocument-format-odf-1-3)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an issue where special characters were included in error messages.

### Excel

- We fixed an issue in which some character sets weren't displaying correctly in sheet tabs

- We fixed an issue where the Find/Replace dialog only saved history for Find and not Replace (the dialog wasn't saving the history of what was replaced when Replace occurred).

- We fixed a rendering issue in worksheets with Freeze Panes enabled for certain scrolling scenarios.

### Outlook

- We fixed an issue where the app stopped responding with large number of groups.

- We fixed an issue that caused Room Finder to fail to load.

- We fixed an issue where Mail subject prepopulated with unexpected characters when created from mail link in contact card.

- We fixed an issue that caused reminders to intermittently display late and show the wrong time in the dialog.

- We fixed an issue that caused deleted meeting invitations to intermittently re-surface for some users.

### PowerPoint

- We fixed an issue where the slide size can change during print preview.

### Project

- We fixed an issue where, when programmatically adding new tasks to a project, the tasks may not be inserted at the correct location if the new task's summary task is collapsed.

### Word

- We fixed an issue where, during the uploading of a file, the application stopped responding and the document wasn't syncing.

- We fixed an issue related to the app stops responding when calling DCompositionCreateDevice.

- We fixed an issue where the app wasn't showing the page numbers accurately for the Chapters/Documents in the inserted Table of Contents using RD field codes.

- We fixed an issue where typing Hiragana with the Japanese input method editor (IME), with the At Mention people picker open, caused the IME to stop working.

### Office Suite

- We fixed an issue where Read Aloud neural voice regression stopped responding.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2109: August 27

*Version 2109 (Build 14420.20004)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Outlook

- **Show multiple mini-months in Calendar To-Do Bar:** This feature allows you to show multiple mini-months both horizontally and vertically in the Calendar To-Do bar.

### PowerPoint

- **Office apps now support OpenDocument Format (ODF) 1.3:** ODF 1.3 brought many improvements to the OpenDocument format and these are now supported in Word, Excel, and PowerPoint (file extensions .odt, .ods, and .odp).

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where, under rare circumstances, Excel could stop responding while calculating a workbook.

- We fixed a stability issue related to graphics during recalculation.

### Outlook

- We fixed an issue where an unexpected informational notification tip appeared when creating appointments on meetings in Outlook which wasn't configured by tenant admins.

### Publisher

- We fixed a stability issue related to graphics during recalculation.

### Word

- We fixed a stability issue related to graphics during recalculation.

### Office Suite

- We fixed an issue where pictures copied from Word didn't preserve the Lock Aspect Ratio setting when pasted into drawing canvases in other Office applications.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2109: August 20

*Version 2109 (Build 14416.20006)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an error that was causing tables dropped into the query design or system relationship window to appear in a different location than where they were dropped.

### Excel

- We fixed an issue with the Insert Cells dialog in which double-clicking on one of the options wasn't applying the selected option and dismissing the dialog.

- We fixed an issue where scrolling using a mouse wheel or touch pad wasn't working if the last row or column in the sheet was hidden.

- We fixed a problem where the Analysis ToolPak add-in didn't work with certain Automation Security settings.

### Outlook

- We fixed an issue that was allowing users to be able to download protected attachments.

- We fixed an error caused when forwarding large or long-running recurring meetings as ICAL.

### Word

- We fixed an issue where enabling autosave could cause recent edits to temporarily disappear.

- We fixed an issue with the comments where icons, stickers, and illustrations weren't visible if they were pasted along with the text in a comment.

- We fixed an issue related to exporting a chart as an EMF, so the EMF has editable labels when inserted as a picture in Office.

- We fixed an issue where Word was failing to render the base-64 encoded, embedded GIFs in the email body.

- We fixed an issue with non-default ribbon configurations that could lead to the Style Gallery not functioning.

- We fixed an issue where the Resolve button was disabled when the focus was on a comment.

### Office Suite

- We fixed an issue related to the following scenario: when using Create Video to export a video from a presentation at the default OneDrive location, an error message would appear saying that the location isn't available.

- We improved behavior during file save to a location requiring user access approval. The Grant Access screen should now appear to allow user access approval.

- We fixed a performance issue related to inserting or editing rows on some versions of Windows.

- We fixed an issue where some documents were saving with an incomplete task pane.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2109: August 13

*Version 2109 (Build 14405.20002)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Word

- **Track just your changes:** Collaborating with others is a key part of producing great content in Word, and the Track Changes feature is an essential part of that process. Previously, when you turned on Track Changes, everyone's changes were automatically tracked. But sometimes you only want your own changes to be tracked, without forcing this setting on others. Now we've given you the power to turn on Track Changes only for yourself. To do so, go to the Review tab and open the dropdown menu on the Track Changes button; then select Just Mine.<br />See details in [blog post](https://insider.office.com/blog/track-just-your-changes-in-word)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an issue where attempting to use the DAO API from non-Office applications would stop responding with "The operating system isn't presently configured to run this application."

- We fixed an issue where, after an error is encountered when pasting records into a subform, data added to the subform is discarded when the form is closed.

### Excel

- We fixed an issue that caused the trend line on charts with logarithmic data to not be smooth.

### Outlook

- We fixed an issue where replying to a message from an external user with a sensitivity label and we do not apply our sensitivity label default.

### Word

- We fixed an issue where the Repeat as Header Row feature in a table was disabled in some cases.

### Office Suite

- We improved behavior during a file Save or Save As to a location requiring user access approval; the Grant Access screen should now appear, to allow user access approval.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2109: August 06

*Version 2109 (Build 14329.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Office Suite

- **Use WebP images in Word, Excel, Outlook, and PowerPoint:** WebP is a modern image format that offers better compression for publishing images to the web. We've now added support in Office apps for WebP images! [Learn More](https://insider.office.com/blog/add-webp-images-to-office-creations)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue with switching separators after a Selection.Parent.Copy call.

- We fixed an issue where Excel was clearing out the value for HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\Common\UserInfo\Company in non-MSI installs of Office.

### Word

- We fixed an issue with the Document.Save command.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2108: July 30

*Version 2108 (Build 14326.20004)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Want your workbook to take you places?:** Understand the layout of your workbook, see what elements exist, and navigate around quickly using the Navigation pane! [Learn more](https://support.office.com/article/ddd037e7-22e3-41f0-8bbd-07f5479e92bf)<br />See details in [blog post](https://insider.office.com/blog/see-the-big-picture-with-navigation-pane-for-excel)

- **Enjoy an improved scrolling experience:** Scrolling through your sheet just got smoother when navigating large or very wide cells. [Learn more](https://support.office.com/article/06fc34b8-64bb-4d78-9b62-34656d700f82)

### Outlook

- **Hero answers in search results:** Outlook Search Answers is the latest development in search enhancement for Outlook, helping you find exactly what you're looking for. Now it's even easier to find the people, files, and calendar events you want, with Outlook surfacing the best answer as a card at the top of the results.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue occurring when saving as XML data (.xml) from workbooks containing an XML map.

- We fixed an issue where scrolling with touch or a touchpad would revert back to the start of the spreadsheet.

### Outlook

- We fixed an issue that caused emails re-sent by a different user to appear to have been sent by the original sender, in organizations where SendFromAliasEnabled is set to True.

- We fixed an issue where some notifications in Outlook desktop weren't actionable when Office is installed on Windows Server 2016.

- We fixed an issue that caused delegates trying to view forwarded meeting requests in their sent items folder to see the manager's copy of the meeting rather than the delegate's sent item.

- We made a change to allow administrators to disable Always On Logging on a per-process basis via group policy.

- We fixed an issue that caused users to be able to download protected voicemail files.

### PowerPoint

- We fixed an issue to re-enable cropping for pictures inside SmartArt objects.

- We fixed an issue where, in certain situations, a user was unable to use Paste Special to paste chart content from Excel to PowerPoint.

### Project

- We fixed an issue where, when you applied a view or table, not all of the columns that were supposed to show up actually showed up.

- We fixed an issue where you couldn't create a visual report if the project had cross-project links and fixed cost.

### Word

- We fixed an issue where cursor may not be visible after a page break

- We fixed an issue where Word becomes unresponsive on updating 'Table of content' field via VBA when track changes is on.

- We fixed an issue which inserting GIF will show security notice.

### Office Suite

- We fixed an issue where Fire coauthoring state changed when transitioning to None.

- We fixed an issue where text was truncated in some languages due to increased text size.

- We fixed an issue with incorrect text wrapping in task panes for high-DPI displays.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2108: July 23

*Version 2108 (Build 14315.20008)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Outlook

- We fixed an issue where a user was unable to open or preview emails.

### PowerPoint

- We fixed an issue where an exception caused an unexpected close of the application.

### Word

- We fixed an issue where the application unexpectedly closed when inserting a link in a comment.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2108: July 16

*Version 2108 (Build 14312.20008)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Outlook

- **Link to Create Outlook.com Account:** When adding an account to Outlook, a link to create a new outlook.com account appears in the window.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where small data markers disappeared when the spreadsheet zoomed out.

- We fixed an issue where a different workbook was opened when a workbook had to be closed and re-opened because it was modified and checked in by a different user.

- We fixed an issue for protected files with no label metadata; the label is determined by the protection. Mandatory labeling now uses label metadata and label policy.

### Word

- We fixed an issue where proofing settings didn't persist.

- We fixed an issue where bullets could disappear from text during co-authoring.

- We fixed an issue related to saving files.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2108: July 09

*Version 2108 (Build 14301.20004)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Outlook

- **Turn off Suggested Replies:** Outlook [makes it easy to reply faster](https://insider.office.com/blog/reply-faster-using-suggested-replies-in-outlook) to emails by offering short suggested replies for messages that can be answered with just a few words. Some users may not want to see this option, so it's now possible to turn the feature off. To do so, select File > Options > Mail, go to the Replies and Forwards section, and clear the Show suggested replies check box.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where a PowerPoint-linked file became unavailable when the source .xlsx file is already running in the background and both files are opened from the ODB local synced folder.

- We fixed an issue where localized function names will now be recognized when opening CSV files.

### Word

- We fixed an issue related to print size being too small in Word Mobile.

- We fixed an issue where a file added to a SharePoint document library would inherit the setting "ShowDocument Information Panel" immediately after opening, and it would remain if the file was removed from SharePoint.

- We fixed an issue where a folder name not ending with "/" was being incorrectly interpreted by the URL parser.

### Office Suite

- We fixed an issue where added DropShadow properties on the search box caused the TitleBar to be too tall, creating layout bugs.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2107: July 02

*Version 2107 (Build 14228.20044)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Additional file types supported for the save-as scenario:** In addition to saving files, you can save files to other file types.

### Outlook

- **REST Forward Meeting Request:** Allows users to forward a previously declined meeting for REST shared calendars.

- **Read Aloud just got better:** The Read Aloud toolbar features new, natural sounding voice options

### PowerPoint

- **Additional file types supported for the save-as scenario:** In addition to saving files, you can save files to other file types.

### Word

- **Search with your voice:** Tap or click the microphone in the search bar to use your voice in Word to find commands, content, and more.

- **More natural voice options for Read Aloud:** Try out a new, more natural sounding voice in the Read Aloud toolbar. [Learn more](https://support.office.com/article/5a2de7f3-1ef4-4795-b24e-64fc2731b001)

- **Additional file types supported for the save-as scenario:** In addition to saving files, you can save files to other file types.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where some users would fail to open into edit mode on password protected files.

### Outlook

- We fixed an issue that caused the translation options to be disabled for some users.  Customers who experienced this bug would have seen their translation options disabled when navigating to File -> Options -> Language. Due to this, they would have been unable to change their preferred translation language and other translation related settings.

- We fixed an issue relating to "failed to load" response status. The default response flag was set to "None." We didn't show any strings in the UI when hovering over a calendar where we didn't have edit permissions.

- We fixed an issue where default text increase includes text scaling, so another call of LayoutChanged doesn't need to be used.

- We fixed an issue where mailtips weren't showing for one-off addresses.

- We added a registry key to allow for the Voicemail form to be displayed in the UI in Outlook Desktop, due to the retirement of Unified messaging in Exchange Online (<https://techcommunity.microsoft.com/t5/exchange-team-blog/retiring-unified-messaging-in-exchange-online/ba-p/608991>).
For users, enterprises, and organizations that want the Voicemail form to appear, the following registry key needs to be set:
[HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\Outlook\Addins] "AllowVoicemailForm"=dword:00000001

### Word

- We fixed an issue which improves integration with the new commenting pane in Word and JAWS, a popular screen-reading software.

- We fixed an issue relating to using a different CommentId than lTagNil for cleared selection and highlight.

- We fixed an issue where the unload queue would become unresponsive.

### Office Suite

- We fixed a localization issue where en-gb, fr-ca, and es-mx will now be matched with their respective parent versions.

- We fixed an issue where sharing settings between OMEX and ExCatalog were no longer possible, such as for web add-in settings updates to the webextension.xml, since a new webextension file is created. The previous one was only accessed when the add-in was deployed in the original method, or the new solution reference comparison was turned off.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2107: June 25

*Version 2107 (Build 14217.20002)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where chart axis values couldn't be changed if both the thousand and decimal separators use the same symbol.

### PowerPoint

- We fixed an issue related to SmartArt nodes having Change Shape disabled.

### Project

- We fixed an issue where engagements created in the Project Web App might not load properly in the Project desktop client if the resource name had special characters, such as a semicolon.

- We fixed an issue where, when the project option “Project should calculate costs” is disabled, then the time-phased cost values might not have been correctly baselined for cost-type resources.

- We fixed an issue where project-level enterprise custom fields with lookup tables weren't showing a value in the Project desktop client.

- We fixed an issue where saving a local project to Project Web App could change a previously saved baseline.

### Word

- We fixed an issue where enabling autosave could cause recent edits to temporarily disappear.

- We fixed an issue with scrolling in the comments pane.

- We fixed an issue where header/footer text wasn't clearly visible in print preview when the Office theme was set to black.

### Office Suite

- We fixed an issue where hyperlinks, including digits, would be broken when composing a message in Outlook in a right-to-left language.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2107: June 18

*Version 2107 (Build 14210.20004)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Outlook

- **Last Sign In / Suspicious Sign In:** Outlook now tells you when and where you last signed into your account, and alerts you if any suspicious sign-in activity is detected

- **Read messages with fewer distractions:** Make it easier to focus on messages with custom text spacing, page colors, column width, and line focus, by turning on Immersive Reader.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- Fixed an issue that may cause applications using the Access Database Engine OLEDB API with a database containing links to SharePoint lists to close unexpectedly.

- Fixed an issue that may cause applications using the Access Database Engine ODBC API to close unexpectedly.

### Word

- We fixed an issue where comments became read-only during collaboration.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2107: June 11

*Version 2107 (Build 14204.20006)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where extra entries appeared in the Excel add-in list for some users.

- We fixed an issue where a saved workbook would appear at the top of the Recent list when saving to an SPO document library.

### OneNote

- We fixed an issue where copying a link to a paragraph didn't always redirect to the correct page.

### Word

- We fixed an issue in which squares appeared when using the Microsoft Word Manuscript Paper add-in.

- We fixed an issue in which some pages in print preview were blank.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2106: June 04

*Version 2106 (Build 14131.20008)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Data type detection settings:** Configure the data type detection behavior when importing data from unstructured sources with Power Query in Excel

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### PowerPoint

- We fixed an issue that ensures a file saved using the 'Retry Save' button on the bus bar is added to the Recent List.

### Word

- We fixed an issue where character spacing increases for specific fonts when rotating them 90 degrees.

- We fixed an issue where comment replies were sometimes lost when coauthoring with multiple users.

### Office Suite

- We fixed an issue where the CLP feature previously caused unprompted saves to the SyncBacked file (File synced by OneDrive).

- We fixed an issue where user is unable to edit files stored in OnPrem servers.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2106: May 28

*Version 2106 (Build 14122.20004)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Outlook

- **Org Explorer:** View and navigate the organization chart.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We removed "RichValue" from Range.valueTypes. Linked Data Types will now return "Error" to match the value of "#VALUE!" returned by Range.values.

### Outlook

- We fixed an issue where users were unable to move items across folders in "non-business" licensed Outlook versions.

- This registry key disables the new Room Finder experience (the same experience as in Outlook for Web) and enables the legacy Room Finder with suggested times.

  Registry Key:
  
  > HKCU\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\Calendar </br>
  > REG_DWORD “ShowLegacyRoomFinder”</br></br>
  > 0 (default) - Outlook uses new Room Finder OWA Powered eXperience when user clicks 'Room Finder' button to search for available rooms  </br>
  > 1 - Outlook uses legacy Room Finder UI to search for available rooms</br>

### Project

- We fixed an issue where assignments on manually scheduled tasks could be moved to an incorrect date.

- We fixed an issue where the resource pool was unresponsive and couldn't be opened.

- We fixed an issue where an error was generated if you created a custom field formula that used the ProjectDate*/ProjectDur* functions with specific date or time parameters.

### Word

- We fixed an issue where some comments weren't saved when exporting a document to PDF.

- We fixed an issue that prevented the editing of a new comment in an unprotected area of a document when Restricted Editing is applied.

- We fixed an issue related to unneeded scrolling animation.

- We fixed an issue causing comments pane closed unexpectedly.

- We fixed an issue that was causing a mismatch between the Editor pane theme and the system theme.

- We fixed a performance issue related to working with large documents.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2106: May 21

*Version 2106 (Build 14117.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Outlook

- **Last Sign In / Suspicious Sign In:** Outlook now tells you when and where you last signed into your account, and alerts you if any suspicious sign-in activity is detected

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue that caused Dynamic Arrays to not update cell values when referenced by RealTimeData functions.

- We fixed an issue that impacted the performance of VLOOKUP and INDEX/MATCH when filling a large amount of data.

- We fixed an issue related to two-finger scrolling when using freeze panes.

- We fixed an issue related to out of memory issues when printing on large-format printers.

### PowerPoint

- We fixed an issue that prevented copying from the speaker notes pane while in Read Only mode.

### Word

- We fixed an issue where a Save As error message was displayed even after a user chose to discard changes.

- We fixed an issue that prevented images from being posted in modern comments.

- We fixed an issue where the Reviewing pane could scroll or appear to scroll but didn't align with selected comment.

- We fixed an issue that caused the selection in the document to not be cleared when clicking outside a newly created comment.

- We fixed an issue where comments aren't highlighted when selected.

- We fixed an issue where the wrong field is getting updated when running a macro if editing restrictions are applied.

- We fixed an issue where some Word files don't open because of corrupt bookmarks.

### Office Suite

- We fixed an issue where signing in with a different account could result in a unexpected close of the app.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2106: May 14

*Version 2106 (Build 14107.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Outlook

- We fixed an issue that caused users to see actionable messages either constantly refreshing or reverting back to headers after download when running in Download Headers Only mode.

- We fixed an issue that caused the people picker in Outlook to expand upwards rather than downwards for users with a perpetual license.

- We fixed an issue that caused users of custom domains to see a warning message about permissions when pasting a link into an email message.

### Word

- We fixed an issue where pressing key combinations such as ctrl + shift + @ wouldn't produce the expected accented character( in this case, 'å').

- We fixed an issue related to image compression.

- We fixed an issue where copying a mail attachment to an application other than Word would fail if the filename included DBCS characters.

- We fixed an issue where Word sometimes displayed a border around text that should have not been there.

### Office Suite

- We fixed an issue where OneDrive would display a merge error message when there was indeed no merge conflict.

- We fixed an issue related to z-order of SVG objects when converted to shapes.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2106: May 07

*Version 2106 (Build 14029.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue that prevented the Name Manager from opening books with a large number of hidden names.

### Outlook

- We fixed an issue that caused users to see copies of all of their sent items appearing in their Outbox folder.

- We fixed an issue that caused Outlook closed unexpectedly when using read aloud with other versions of Windows.

### Word

- We fixed an issue that caused Word closed unexpectedly when using read aloud with other versions of Windows.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2105: April 30

*Version 2105 (Build 14026.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Outlook

- **Accessibility Checker nudge when sending emails to large DL, external users:** We added the functionally to get prompted automatically, through a mailtip, of an accessibility violation while composing an email to large audiences, external users, etc. These settings live in the Ease of Access

### Visio

- **AWS stencils and shapes:** We now have stencils with the latest AWS shapes to help you create diagrams. [Learn more](https://support.office.com/article/138206bf-d10f-4583-9f31-885ce706af49)

### Word

- **Writing Goals:** Writing Goals for WinWord

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue that cause Excel to close unexpectedly when moving through comments in the Comments pane.

- We fixed an issue that caused date formatting to be displayed incorrectly in some languages when using add-ins.

- We fixed an issue which could cause Excel to close unexpectedly when using Paste Special with formats in certain situations.

### Project

- We fixed an issue where changes done through Planning Wizards weren't always captured by change events.

- We fixed an issue where users were unable to remove projects from the resource pool.

### Word

- We fixed an issue where text formatting remains after removing hyperlinks.

- We fixed an issue where comments aren't displayed after filtering by people.

- We fixed an issue where Word was unable to perform a Mail Merge with an Access database.

- We fixed an issue that caused the ability to collapse margins in a document containing multiple columns to be available.

- We fixed an issue where some characters aren't displayed correctly in table cells when there are comments in the document.

- We fixed an issue where the file format changes occurred when saving documents with the AIP add-in enabled.

- We fixed an issue where Word would become unresponsive when editing fields.

- We fixed an issue where users wouldn't be prompted to save documents when using a command (rather than the CTRL+S keyboard shortcut).

- We fixed an issue where the sensitivity label disappears from a file in Word after uploading the file to SharePoint Online.

### Office Suite

- We fixed an issue that caused the Dictation button to be misaligned when adding comments to a document.

- We fixed an issue where using High Contrast mode for extended periods of time would cause Outlook to close unexpectedly.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2105: April 23

*Version 2105 (Build 14014.20002)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Import data from dynamic arrays:** You can now import, shape and refresh data from dynamic arrays in the current workbook. [Learn more](https://support.office.com/article/205c6b06-03ba-4151-89a1-87a7eb36e531)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue to support backward compatibility with older versions of Excel. The issue may cause a file that is saved in a more recent version of Excel fail to load properly in older versions of Excel due to  functions such as IFERROR and XLOOKUP added to Excel since Office 2007.

- We fixed an issue where some files would occasionally fail to open in Protected View.

- We fixed an issue that caused the status bar to not indicate a Ready state for some users.

### Outlook

- We fixed an issue that caused name resolution to fail when sending on behalf of another user and resolving against an address book that isn't the Global Address List.

### Word

- We fixed a issue where placeholder text was clipped in comments when using right-to-left languages.

### Office Suite

- We fixed an issue where hyperlinks including digits would be broken when composing a message in Outlook in a right-to-left language.

- We fixed an issue where some Scalable Vector Graphics (SVG) didn't render correctly.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2105: April 16

*Version 2105 (Build 14007.20002)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue that caused Excel to close unexpectedly when using 32 bit Office on 64 bit Windows.

- We fixed an issue that caused Narrator to incorrectly read the properties of two buttons on the Header/Footer tab in the Page Setup dialog box.

### PowerPoint

- We fixed an issue related to linked pictures.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2105: April 09

*Version 2105 (Build 14002.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Accessibility ribbon:** Find all of the tools you need to create accessible content in one place - the Accessibility ribbon!

### Word

- **Proofing is now available for selected text within the Document:** With these changes you can now review spelling, grammar and other intelligent writing suggestions for just the selected text. Additionally you can also review suggestions for the whole document.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

## Version 2104: April 02

*Version 2104 (Build 13929.20016)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Outlook

- **Suggested Replies in Outlook for Windows:** When you receive an email message that can be answered by a short response, Outlook can suggest three responses you can use to reply with just a couple of clicks.

- **Turn on shared calendar improvements:** For shared calendars in Office 365, Outlook can update these calendars using the REST API. Turn on the preview for faster and more reliable updates to shared calendars.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an issue when an external application requests an accessibility interface, it will prevent us from shutting down until they release their reference.

### Word

- In this bug, specific policies weren't being honored by Office (a group of templates were being shown on the Home Page when they should have been disabled). With this fix, the policies are being honored.

- We fixed an issue in auto save.

- We made a fix in Application.OnTime where it might not trigger correctly.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2104: March 26

*Version 2104 (Build 13919.20002)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Outlook

- We fixed an issue that caused some users to experience Outlook to close unexpectedly when syncing folder hierarchy changes.

### Word

- We fixed an issue which copy and paste styles may not be same in pasted text.

### Office Suite

- Fixed a performance issue related to iteration of text.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2104: March 19

*Version 2104 (Build 13913.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- This change fixes an issue where in some cases running a SQL Server pass through query could result in an error message indicating that there was an "invalid cursor state".

### Excel

- We fixed a problem that was preventing the ability to paste as formulas on a protected sheet.

### Project

- Fixed an issue where if the date format is W4/4, the date picker may show the wrong day and year.

### Office Suite

- Fixes an issue with support of GDI+ LineJoinMiterClipped in Office.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2104: March 12

*Version 2104 (Build 13906.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### PowerPoint

- **Insert Flipgrid videos in PowerPoint:** PowerPoint will now support insertion of Flipgrid videos in slides.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue in which hyperlinks created using the HYPERLINK function wouldn't work if the file was saved as a PDF document.

### Word

- We fixed an issue with comments during coauthoring.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2103: March 05

*Version 2103 (Build 13901.20036)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **AutoSave and coauthoring on sensitive encrypted documents:** Don't trade off productivity for security. With Microsoft Purview Information Protection, documents that are encrypted with sensitivity labels can now be AutoSaved and coauthored with others in real time just like unencrypted documents can. Requires tenant opt-in (more info: <https://aka.ms/mipcoauth>).

### PowerPoint

- **AutoSave and coauthoring on sensitive encrypted documents:** Don't trade off productivity for security. With Microsoft Purview Information Protection, documents that are encrypted with sensitivity labels can now be AutoSaved and coauthored with others in real time just like unencrypted documents can. Requires tenant opt-in (more info: <https://aka.ms/mipcoauth>).

### Word

- **AutoSave and coauthoring on sensitive encrypted documents:** Don't trade off productivity for security. With Microsoft Purview Information Protection, documents that are encrypted with sensitivity labels can now be AutoSaved and coauthored with others in real time just like unencrypted documents can. Requires tenant opt-in (more info: <https://aka.ms/mipcoauth>).

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- Corrected an issue where the font would change unexpectedly when using a multiplication or divide sign with a Japanese font. We now continue to use the same font if it supports the character.

- We fixed an issue that caused some PivotTable formatting to corrupt the workbook when saving to the .xls or .xlt format.

- We fixed an issue where some notes were unexpectedly shown when opening a workbook.

### Outlook

- We fixed an issue that caused non-ASCII characters to export incorrectly when exporting to CSV.

- We fixed an issue that caused users to be unable to look up a contact group with Check Names  when composing mail.

### PowerPoint

- We fixed an issue where arrows in line charts weren't appearing as expected in PowerPoint slideshow mode.

### Word

- We fixed an issue where opening a file protected with a Microsoft Purview Information Protection label can stop responding indefinitely if the user isn't signed in to an identity that has access to the MIP protected label. The user is forced to cancel the open to show the sign-in prompt, and the open only succeeds after that point. Fixing this issue by allowing the sign-in prompt to be shown during open/download.

- We fixed an issue when using Dictation in the new Word Commenting, the Dictation button in the Comment card now correctly toggles on and off.

- Fixed an issue where there was no space being inserted between words when users dictated into their document.

- We fixed an issue when posting multi-line comments typed in RTL caused the 2nd and onward lines to be aligned to the left instead of the right.

- We fixed an issue where spell check switched between two different spelling correction context menus.

- We fixed an issue where columns might have overlapping text.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2103: February 26

*Version 2103 (Build 13819.20006)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue that prevented users from exporting an Excel workbook to PDF.

- We fixed an issue where some formatting could be lost when copying a sheet while coauthoring.

### Outlook

- We fixed an issue that caused users to see attachments getting duplicated when removing DRM protection.

### Project

- Fixed an issue where task splits may be wrongly created when saving a project from Project web app to a local file. This would happen if a task calendar with non-standard working times was being used.

- Fixed an issue where if the indicator column isn't in the first column spot, when you cut a summary task you aren't warned that the subtasks will also be removed.

- Fixed an issue where if a user selected the Add Yourself to a Task function on their Timesheet, the correctly resource availability units may not be used on the created assignment.

### Word

- We fixed an issue with alignment of multiple comments.

- We fixed an issue in Read Aloud task pane keyboard shortcuts.

### Office Suite

- OneDrive places are now filtered out as appropriate per the group policy setting.

- Fixes an issue related to unresponsiveness that may occur when loading EMF images.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2103: February 19

*Version 2103 (Build 13811.20002)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Outlook

- We fixed an issue that caused users to see attachments getting duplicated when removing DRM protection.

### Project

- Fixed an issue where if the indicator column isn't in the first column spot, when you cut a summary task you aren't warned that the subtasks will also be removed.

- Fixed an issue where if a user selected the Add Yourself to a Task function on their Timesheet, the correctly resource availability units may not be used on the created assignment.

### Word

- We fixed an issue in Read Aloud task pane keyboard shortcuts.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2103: February 12

*Version 2103 (Build 13806.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Outlook

- **Microsoft Search powered compose (To\CC\BCC) suggestions:** Adding people to the To\CC line is now powered by Microsoft Search.

- **Dictation is available in more languages:** Dictation now supports 7 new languages: Hindi, Russian, Polish, Portuguese (Portugal), Korean, Thai, Chinese (Taiwan)

### Word

- **Dictation is available in more languages:** Dictation now supports 7 new languages: Hindi, Russian, Polish, Portuguese (Portugal), Korean, Thai, Chinese (Taiwan)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- Fixed an issue where Excel would sometimes close unexpectedly when trying to show the Data Types card due to a not being able to retrieve an image.

### PowerPoint

- Fixes an issue with looping animations and audio bookmarks.

### Project

- Fixed an issue where a task that are 100% complete may revert back to 99% complete.

- Fixed an issue where Project may close unexpectedly if you are running JAWS and go to the task information dialog.

### Word

- We fixed an issue in AutoSave.

- We fixed an issue in resolving conflicts while in coauthoring.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2102: February 05

*Version 2102 (Build 13801.20004)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- You will now see selected tabs clearer in Access.

### Excel

- We fixed an issue where Excel would stop responding after selecting a data series in a chart.

- We fixed an issue where pressing Enter with certain keyboards on Android would add a new line rather than moving to the next cell.

- Fixed an issue related to pictures retaining their aspect ratio during a crop operation.

### Outlook

- We fixed an issue that caused mails to be sent as digitally signed after the user unchecked that option.

### PowerPoint

- Fixed an issue related to pictures retaining their aspect ratio during a crop operation.

### Word

- Fixed an issue related to pictures retaining their aspect ratio during a crop operation.

- We fixed an issue which the comment may be truncated with links.

- We fixed an issue with conflict mode when coauthoring.

- We fixed an issue in saving to SharePoint Online

- We fixed an issue in exporting Word document to PDF.

### Office Suite

- Fixed an issue where Office would in some circumstances present sensitivity labels for one signed-in account when it should present sensitivity labels for a different signed-in account.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2102: January 29

*Version 2102 (Build 13721.20008)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed a problem where Excel would unexpectedly quit when you added a Name in the Define Name dialog.

### Outlook

- We fixed an issue that caused the encryption icon to fail to display for emails sent using the Encrypt Only option.

### Project

- Fixed an issue where projects with long Cyrillic names couldn't be opened through Project Center.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2102: January 22

*Version 2102 (Build 13714.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Government customers: Apply sensitivity labels to your documents and emails:** Sensitivity labeling features are now available for customers in the GCC and GCC-H environments. [Learn more](/microsoft-365/compliance/sensitivity-labels)

### Outlook

- **Government customers: Apply sensitivity labels to your documents and emails:** Sensitivity labeling features are now available for customers in the GCC and GCC-H environments. [Learn more](/microsoft-365/compliance/sensitivity-labels)

### PowerPoint

- **Government customers: Apply sensitivity labels to your documents and emails:** Sensitivity labeling features are now available for customers in the GCC and GCC-H environments. [Learn more](/microsoft-365/compliance/sensitivity-labels)

### Word

- **Government customers: Apply sensitivity labels to your documents and emails:** Sensitivity labeling features are now available for customers in the GCC and GCC-H environments. [Learn more](/microsoft-365/compliance/sensitivity-labels)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- Fixes issue where certain charts using discontinuous ranges of cells wouldn't load when files are re-opened.

- Fixes an issue where Excel would fail to launch or close unexpectedly if certain Windows Security exploit protection settings (SimExec, CallerCheck) are in use.

### PowerPoint

- Fixed an issue related to displaying emojis with color.

### Word

- This fixes an issue that prevented real-time typing and presence from being restored after loosing internet connectivity for a period of time.

- We fixed an issue with coauthoring.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2102: January 15

*Version 2102 (Build 13707.20008)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Outlook

- **Share to Teams:** Share an email or a conversation from Outlook to a person or channel in Teams.

### Visio

- **Ready-made graphics for your diagrams:** Choose from a large library of icons, stock photo images, cutout people, and stickers that you can add to your Visio drawings. [Learn more](https://support.office.com/article/0ab844a5-289b-47f2-ba92-eeda168bc381)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Project

- Fixed an issue where when a cost resource was assigned to a milestone task, baseline cost didn't rollup correctly.

### Word

- Fixing a failure when running the VBA macro ExportAsFixedFormat2 fails with an error stating "Presentation (unknown member) illegal value".

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2102: January 08

*Version 2102 (Build 13704.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Outlook

- **Dictation just got better:** It's now easier to create content with your voice with the new dictation toolbar, voice commands, and auto-punctuation support

### Word

- **Dictation just got better:** It's now easier to create content with your voice with the new dictation toolbar, voice commands, and auto-punctuation support

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We have fixed an issue where Preview of embedded Excel range in PowerPoint shows incorrect size.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2101: January 01

*Version 2101 (Build 13624.20002)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Mandatory Labeling:** Users with the Mandatory Labeling policy set their Admins are required to label their documents and emails.

### PowerPoint

- **Mandatory Labeling:** Users with the Mandatory Labeling policy set their Admins are required to label their documents and emails.

### Word

- **Mandatory Labeling:** Users with the Mandatory Labeling policy set their Admins are required to label their documents and emails.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### OneNote

- This change addresses a rendering issue affecting OneNote.

### Outlook

- This change enables Outlook to take advantage of an Exchange server setting that suppresses the display of the Exchange Online Archive Mailbox to end users.

### PowerPoint

- This change addresses an issue with Merge Shapes working with text.

### Word

- Fix to make Modern comments more robust.

- Fixed an issue when editing commenting post with @mention.

- Comment drafts disappears when creating a new Word instance.

- Fix and issue with nested scrollbars in the comments pane.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2101: December 25

*Version 2101 (Build 13617.20002)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- Update so that decimal and thousands separators settings carryover when copying a chart from Excel and pasting into Word

- Fixed an issue where Excel would close unexpectedly when opening UNC files that have invalid file attributes (creation time, modified time, etc.)

- This change addresses an issue related to changing outline colors of SVG images.

### Outlook

- We fixed an issue that caused users to be unable to specify how long they wanted to allow access for when starting a mail merge from Word, resulting in them getting excess prompts.

- We fixed an issue that caused a Outlook to close unexpectedly for users of Redemption-based add-ins.

### PowerPoint

- This change addresses an issue related to changing outline colors of SVG images.

### Word

- This change addresses an issue related to changing outline colors of SVG images.

- Fixes an issue where the reply box on a comment card is off the screen.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2101: December 18

*Version 2101 (Build 13610.20002)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Send audit data about sensitivity labeling to M365 administrators:** When users apply, change, or remove sensitivity labels on their documents and emails, Office will send up audit data to the M365 audit backend for administrators to see. This is a silent functionality (no UI) for administrator benefit.

### Outlook

- **Send audit data about sensitivity labeling to M365 administrators:** When users apply, change, or remove sensitivity labels on their documents and emails, Office will send up audit data to the M365 audit backend for administrators to see. This is a silent functionality (no UI) for administrator benefit.

### PowerPoint

- **Send audit data about sensitivity labeling to M365 administrators:** When users apply, change, or remove sensitivity labels on their documents and emails, Office will send up audit data to the M365 audit backend for administrators to see. This is a silent functionality (no UI) for administrator benefit.

### Word

- **Send audit data about sensitivity labeling to M365 administrators:** When users apply, change, or remove sensitivity labels on their documents and emails, Office will send up audit data to the M365 audit backend for administrators to see. This is a silent functionality (no UI) for administrator benefit.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- Made a performance improvement when applying formatting styles to pivot tables.

### Outlook

- We fixed and issue that caused users to be unable to select more than one category to search.

- We fixed an issue that caused the start time of some calendar items to change unexpectedly when the event is copied from another appointment.

### Project

- We fixed an issue where users open projects which have supposedly been saved with updated information, but find the updates are is missing.

### Word

- Fix animation when typing on the bottom of a comment card.

- We fixed an issue where Word hangs when saving document to PDF with hidden text.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2101: December 11

*Version 2101 (Build 13604.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Outlook

- **Your Outlook settings in the cloud:** Choose your Outlook for Windows settings like Automatic Replies, Focused Inbox, and Privacy, and get to them on any PC.

### Word

- **Better collaboration with modern comments:** Add comments to objects, @mention colleagues, and resolve comment threads for a better collaboration experience. [Learn more](https://support.office.com/article/8d3f868a-867e-4df2-8c68-bf96671641e2)<br />See details in [blog post](https://insider.office.com/blog/modern-commenting-in-word)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- Fixed an issue where Excel would incorrectly show a message bar that a new version of the file is available and force the user to save their changes in a copy of the workbook or discard their changes.

- Fixed an issue with switching separators after a Selection.Parent.Copy call.

### Outlook

- Addressed an issue that caused plain text S/MIME messages to become garbled when sending.

### PowerPoint

- This change addresses an issue with looping background videos playback in Slide Show.

- We have fixed an issue where font size command, added in QAT, auto completes to the nearest defined font size while updating it.

### Word

- We fixed an issue around deleting modern comments in a content control that is marked as not editable.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2012: December 04

*Version 2012 (Build 13530.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Outlook

- **Build in time between back-to-back meetings:** Give attendees time to catch their breath or travel between locations by setting meetings to start 5-10 min late by default. [Learn more](https://support.office.com/article/be84396a-0903-4e25-b31c-1c99ce0dacf2)

### Visio

- **New Azure stencils and shapes:** We've added many more stencils to help you create up-to-date Azure diagrams. [Learn more](https://support.office.com/article/efbb25e7-c80e-42e1-b1ad-7ef630ff01b7)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- Fixed an issue where editing in languages that require use of IME would behave poorly when editing in overwrite mode.

- Fixed a broken hyperlink to a help article in an alert in case Autosave becomes disabled.

- Fixed an issue where Excel would close unexpectedly when copying and pasting data in formula view.

### Outlook

- We fixed an issue that caused SmartLinks to lose their formatting when saved to drafts.

### Project

- Fixed an issue where opening a project with many resources was taking a long time.

- Fixed an issue where specific projects could be opened if there was an issue with the project file in a specific part of load.

### Word

- Paste as plain text is often preferred to pasting as rich text. This context menu fix allows the user to paste as plain text. Else the user would have to copy into a plain-text editor like Notepad and then copy from Notepad into the desired target app

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2012: November 27

*Version 2012 (Build 13519.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- This fixes an issue where Power Pivot wasn't able to correctly import a tab-delimited text file.

### Outlook

- We fixed an issue that was causing users to experience some issues when sending Outlook mail from applications other than Outlook.

### PowerPoint

- This change addresses an issue related to timeouts experienced during ink analysis.

- This change addresses a grammatical error in the Create an Animated GIF user interface.

- Fixed an issue where some corrupt PowerPoint files weren't opening correctly, even after a document repair operation.

### Project

- Fixed an issue where users may see multiple unassigned assignments associated with a task.

- Fixed an issue where in large projects it can be very slow to enter a task name.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2012: November 20

*Version 2012 (Build 13512.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Outlook

- **Every meeting online:** Make it easier to schedule online meetings with a new setting to make all your meetings online by default.

### PowerPoint

- **Video Library:** Elevate your documents with a library of curated, royalty-free video footage available in-app

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### PowerPoint

- Fixed an issue where the presence indicator for an unknown coauthor doesn't get completely refreshed, once more information about the coauthor is available.

### Word

- We fixed an issue where Word hangs when saving document to PDF with hidden text.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2012: November 13

*Version 2012 (Build 13510.20004)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where the Insert Object command doesn't show the correct icon when inserting a file from OneDrive local sync folder.

### Outlook

- We fixed an issue that caused the To: field to be blank in task status reports.

### PowerPoint

- We fixed a VBA issue where Slide.Shapes.AddMediaObject2 closed unexpectedly with legacy video formats (MPG-1,Mpeg-2).

### Project

- We fixed an issue where you couldn't delete dependencies on the deliverables if the SharePoint site the deliverable was associated with no longer existed.

- We fixed an issue where users open projects which have supposedly been saved with updated information, but find the updates are is missing.

### Word

- We fixed an issue related to pictures becoming blurry while zooming.

- We fixed an issue where long hyperlinks were getting truncated.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2012: November 06

*Version 2012 (Build 13430.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Unhide many sheets at the same time:** No need to unhide one sheet at a time anymore -- unhide multiple hidden sheets at once. [Learn more](https://support.office.com/article/69f2701a-21f5-4186-87d7-341a8cf53344)

### Outlook

- **Same signature, all devices:** Your signature is stored in the cloud. Create it once and use it everywhere you use Outlook.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where some Ribbon elements weren't localized in Simplified Chinese.

- We fixed an issue where Excel terminated unexpectedly when updating.

### Outlook

- We fixed and issue where adding an attachment to a message opened from a zip file would fail.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2011: October 30

*Version 2011 (Build 13426.20004)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Improved Conditional Formatting dialogs:** Conditional Formatting dialogs are now resizable, and now you can duplicate the rule with a single click. [Learn more](https://support.office.com/article/fed60dfa-1d3f-4e13-9ecb-f1951ff89d7f)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an issue when using DAO from non-Office applications would cause the application to close unexpectedly.

### Excel

- Fixed a problem with Power Pivot when using a connection to an Oracle database.

- We fixed an issue where Excel terminated unexpectedly when the process of MTR calc and Group Policy Object update (e.g. via Remote Group Policy Refresh) was triggered.

- This change fixes a bug, which causes a failure in Excel on attempt to load an atomsvc file.

- We fixed an issue which Word appears to stop responding when insert Excel workbook into Word document.

### Outlook

- We fixed an issue where a mailbox owner wasn't able to manage Shared permission for their own Calendar as the option was greyed out.

- We fixed an issue where saving email templates as .OFT changed Chinese characters to question marks.

- We fixed an issue where Outlook wasn't able to create a message with restricted permission.

- Addressed an issue which caused Outlook to stop working sporadically when adding or saving attachments.

### PowerPoint

- We fixed an issue where grid lines were getting shifted from slides when closing design pane.

- We fixed an issue where the scroll bar in the slide starts adjusting itself after stopping screen recording with selection pane opened.

### Project

- Fixed an issue where when you save a project from PWA to a local mpp file, the ProjectBeforeTaskChangeEvent fires for data that wasn't actually changed by the user.

- Fixed an issue where resource engagements searched for a resource by name instead of GUID which would cause issues if there were multiple resources with the same name.

### Word

- We fixed an issue where clicking comment hint didn't zoom out to show comment card in view.

- We fixed a layout issue which the line between columns might have shifted.

- We fixed an issue which Word appears to stop responding when insert Excel workbook into Word document.

### Office Suite

- We fixed an issue in the Office Deployment Tool where configuration was failing when using the RemoveMSI feature with the Office 2007 "Microsoft Application Error Reporting" product present.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2011: October 23

*Version 2011 (Build 13415.20002)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### PowerPoint

- **Rehearse your presentation with Presenter Coach:** Get feedback on the things that help keep an audience engaged — like pacing, pitch, filler words, sensitive phrases, and more. [Learn more](https://support.office.com/article/cd7fc941-5c3b-498c-a225-83ef3f64f07b)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an issue where some users were seeing the "system resource exceeded" error when they tried to export a query from their synced OneDrive folder.

- We fixed an issue where 'auto'-switching between form windows was switching to another form.

### Outlook

- We fixed an issue when pasting a URL copied from meeting location to somewhere else (such as a browser), the URL contains a semicolon at the end.

### PowerPoint

- We fixed an issue when duplicating slideshow to secondary monitor, the slideshow may hide behind other window.

### Project

- Fixed an issue where Project may terminate unexpectedly on opening files where resource contours were specified in a certain manner.

### Word

- We fixed an issue in Track Changes which sometimes opening Word document might display error dialog.

- We fixed a print issue when sensitivity label with watermarks are applied.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2011: October 16

*Version 2011 (Build 13408.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **SVG Clipboard Support:** You can now paste SVG content from Office into 3rd party apps. [Learn more](https://support.office.com/article/69f29d39-194a-4072-8c35-dbe5e7ea528c)

### Outlook

- **SVG Clipboard Support:** You can now paste SVG content from Office into 3rd party apps. [Learn more](https://support.office.com/article/69f29d39-194a-4072-8c35-dbe5e7ea528c)

### PowerPoint

- **SVG Clipboard Support:** You can now paste SVG content from Office into 3rd party apps. [Learn more](https://support.office.com/article/69f29d39-194a-4072-8c35-dbe5e7ea528c)

### Word

- **SVG Clipboard Support:** You can now paste SVG content from Office into 3rd party apps. [Learn more](https://support.office.com/article/69f29d39-194a-4072-8c35-dbe5e7ea528c)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Outlook

- We fixed an issue where users were unable to delete appointments in Calendar of Microsoft 365 Groups in Basic Auth.

- We fixed an issue where starting Outlook failed while loading nickname cache.

### PowerPoint

- We fixed an issue where in content placeholder icon next to Pictures didn't have a tooltip.

- We have fixed an issue where Protected view of slide show, shown by pptsx file, allows screen capture of IRM protected document.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2011: October 09

*Version 2011 (Build 13406.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Create Power Platform dataflows from queries:** You can now export your queries into Power Query templates that can be used to create new Power Platform dataflows

### PowerPoint

- **Export animated GIF in a range:** Select a range of slides when exporting to animated GIF

- **Create GIFs with Transparent Backgrounds:** When exporting to an Animated GIF, a new option will allow you to make the background transparent.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where  Filename wasn't changing after a SaveAs operation with COM add-ins enabled.

- We fixed an issue when autosave fails with incorrect/misleading error message when there's a bad measure definition in the Excel data model.

- We fixed an issue where zooming in and out from the presentation area resulted in a gap between the zoomed selection marquee and the mouse pointer.

### Outlook

- We fixed an issue where quick print for image attachments resulted in error, "Windows can't find this picture. Check the location, and then try again".

- Fixes an issue that caused some users to see Outlook to start in an Offline state until they manually chose to work online.

### PowerPoint

- We fixed an issue where zooming in and out from the presentation area resulted in a gap between the zoomed selection marquee and the mouse pointer.

### Project

- Fixed an issue where the NewVal in the ProjectBeforeTaskChagne event doesn't have the correct value if a lag is changed within a Task Form type view.

- Fixed an issue where if you have a task list in a project site and group the task list, you won't be able to quick edit the task list.

- Fixed an issue where if you update an enterprise resource via CSOM, resource max units may be lost.

### Word

- We fixed an issue where zooming in and out from the presentation area resulted in a gap between the zoomed selection marquee and the mouse pointer.

### Office Suite

- We fixed an issue where SSO API interactive Sign-In was returning an error code.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2010: October 02

*Version 2010 (Build 13328.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Use the Advanced Dialog to Create Data Types:** The Advanced Dialog allows you to manually select the columns which combine the Data Type you are creating.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### OneNote

- We fixed an issue where a user was unable to select and copy notebook URL from textbox in OutSpace File > Info.

### Outlook

- Addresses an issue that caused automatically generated emails to be sent with a blank body when the subject is blank.

- We fixed an issue where the wrong folder guid is cached for folders.

- When a user copy-and-pastes an email address into the recipient field with the display name, the email address wasn't always parsed correctly and caused a warning about an invalid email address to appear.  It's been fixed so the name and email address are parsed correctly so the warning is no longer shown.

- We fixed an issue where online shared folders didn't return parent folder name. Instead of failing, it returned an empty path which incorrectly went to the primary account.

- We fixed an issue where track changes turned on after reopening draft from read-only preview pane.

- We fixed an issue where the Save As option wasn't available for classic attachments.

- We fixed an issue to provide a user a way to customize justification text when overriding a policy.

### PowerPoint

- We fixed an issue where PowerPoint wasn't exporting rectangle bullet points while exporting to PDF.

- We fixed an issue that If you are on the last slide and if you swipe to the next slide after pressing 'End Session' and before the summary shows up, the End-Session dialog is visible on the summary page as well.

### Project

- Fixed an issue where Project may close unexpectedly if you apply a group by to the Resource Usage or Sheet view and then insert a column.

- Fixed an issue where if you have custom fields with formulas and are using earned value, you may notice performance delays switching views and opening project/task details.

- Fixed an issue where the ConsolidateProjects VBA method may file if you try to add the same project multiple times and have AttachToSources set to false.

- Fixed an issue where if you have eventing code running and try to make changes through a Task Form view, clicking the OK button may not commit the changes.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2010: September 25

*Version 2010 (Build 13318.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Create Data Types with Power Query:** Create rich data types with Power Query from any data source

### Outlook

- **User Experience Updates for Tasks:** A visual refresh of task items

- **Save time while composing messages:** Outlook shows you writing suggestions that help you compose messages quickly. To accept the suggestion, just use the Tab key.

### Word

- **Microsoft Editor pane gets an update in Word for desktop:** We have upgraded the current experience with the Editor pane in Word for desktop clients.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an issue where scrollbar position wasn't set correctly when loading query/relationship window saved while scrolled.

- We fixed an issue where Add Table Task Pane wasn't displaying names containing '&' correctly.

### Excel

- We fixed an issue where multi-level category manual interval wasn't working in charts.

- We fixed an issue which could cause the app to stop responding when refreshing OLAP PivotTables.

- We fixed an issue where adding to a table used for Data Validation didn't update options for all sheets in the workbook.

### OneNote

- We fixed an issue where OneNote didn't honor High Contrast colors in the canvas for custom themes.

- We fixed an issue when you hover over green color in notebook color selector, the pop up reads "red chalk".

### PowerPoint

- We fixed an issue where linked excel chart incorrectly changes to excel sheet when user changes the source path to local OneDrive folder.

### Word

- We fixed an issue where links to workflow enabled files weren't opening as expected.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2010: September 18

*Version 2010 (Build 13312.20006)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Outlook

- **Proofread your messages with Editor:** You can now get grammar and other style suggestions in your emails for Outlook 64-bit users. Look for underlined words to see Editor's suggestions to refine your writing.

- **Break the language barrier with a built-in translator:** Add-ins for translation aren't required anymore! In a message, right-click to translate specific words, phrases, or the whole message.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- Fixed an issue with 2D Map Charts where using VBA to set the colors for the max, mid, and min values for a series wasn't working.

- Fixed an issue when the Office language was set to Spanish, in which data validation lists may not show all the items in the list.

- Fixed an issue which could cause an error that "Excel ran out of resources while attempting to calculate one or more formulas".

- Fixed an issue where ChartSheet closed unexpectedly in some cases when a formula is entered through formula bar.

### Outlook

- When a user copy-and-pastes an email address into the recipient field with the display name, the email address wasn't always parsed correctly and caused a warning about an invalid email address to appear.  It's been fixed so the name and email address are parsed correctly so the warning is no longer shown.

### Word

- Fixed an issue where a user tapping on a tracked change (insertion/deletion) would bring up a comment pop-out.

- We fixed an issue with deleting comment callouts in Word.

- We fixed an issue with Outlook with message set to Do not forward.

- We fixed an issue with saving Word document that contains citation and equation.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2010: September 11

*Version 2010 (Build 13304.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Access

- **Office can follow your Windows 10 Dark Mode setting:** Using Windows 10 in Dark Mode? Office can now switch themes to match automatically - just choose "Use system setting" as your Office Theme.

### Excel

- **Office can follow your Windows 10 Dark Mode setting:** Using Windows 10 in Dark Mode? Office can now switch themes to match automatically - just choose "Use system setting" as your Office Theme.

### OneNote

- **Office can follow your Windows 10 Dark Mode setting:** Using Windows 10 in Dark Mode? Office can now switch themes to match automatically - just choose "Use system setting" as your Office Theme.

### Outlook

- **Office can follow your Windows 10 Dark Mode setting:** Using Windows 10 in Dark Mode? Office can now switch themes to match automatically - just choose "Use system setting" as your Office Theme.

### PowerPoint

- **Office can follow your Windows 10 Dark Mode setting:** Using Windows 10 in Dark Mode? Office can now switch themes to match automatically - just choose "Use system setting" as your Office Theme.

### Project

- **Office can follow your Windows 10 Dark Mode setting:** Using Windows 10 in Dark Mode? Office can now switch themes to match automatically - just choose "Use system setting" as your Office Theme.

### Publisher

- **Office can follow your Windows 10 Dark Mode setting:** Using Windows 10 in Dark Mode? Office can now switch themes to match automatically - just choose "Use system setting" as your Office Theme.

### Visio

- **Office can follow your Windows 10 Dark Mode setting:** Using Windows 10 in Dark Mode? Office can now switch themes to match automatically - just choose "Use system setting" as your Office Theme.

### Word

- **Office can follow your Windows 10 Dark Mode setting:** Using Windows 10 in Dark Mode? Office can now switch themes to match automatically - just choose "Use system setting" as your Office Theme.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- Fixed an issue where there could be a noticeable delay when switching between worksheets with large amounts of data when 'Page Break Preview' was enabled.

### Outlook

- We fixed an issue with emails being hidden after turning Focused Inbox off and doing a sort.

### PowerPoint

- We fixed an issue where GIF's would animate only once in the editor and slide shows.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2010: September 04

*Version 2010 (Build 13301.20004)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Outlook

- **Pinning email:** This feature helps users keep track of mails they need to go back to you or have as a reminder by keeping them at the top of the message list.

- **Receive email suggestions when searching by person:** As you type your search terms in Outlook, you'll receive the most relevant emails surfaced in the suggestions.

- **Receive email suggestions when searching by person:** As you type your search terms in Outlook, you'll receive the most relevant emails surfaced in the suggestions.

- **Microsoft Editor gets an upgrade for Word and Outlook desktop clients:** We are introducing a new click-to-review model for Editor's spelling ,grammar and advanced style suggestions. This change also includes a new dedicated card surface for reviewing the suggestions.

### Word

- **Microsoft Editor gets an upgrade for Word and Outlook desktop clients:** We are introducing a new click-to-review model for Editor's spelling ,grammar and advanced style suggestions. This change also includes a new dedicated card surface for reviewing the suggestions.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where if you opened a file containing the LET function, it would display the alert:  "We found a problem with content in "your file.xlsx". Do you want us to try to recover as much as we can? If you trust the source of this workbook, click Yes".
- We fixed a unexpected app closure related to XLAM add-in references and named ranges.
- We fixed an issue where users couldn't modify a PivotTable filter because it was set to a value that was no longer present in an Analysis Services database.
- We fixed an issue where if a user applied a custom style to a dynamic array, they would get the error: "You can't change part of an array". This was a legacy restriction that has been removed.
- We fixed an issue where the Excel formula bar wouldn't render completely after connection to a device was lost, such as a remote session connect/disconnect or a monitor change.

### Outlook

- We fixed an issue which provides more flexibility in enabling / disabling the default logging options via Group Policy.
- We fixed an issue where the Legacy Domain Name for an email sender was preserved and displayed after a draft of the email was moved between mailboxes with assistant permissions and manager permissions.
- We fixed an issue that caused some users to see Outlook to start in an Offline state until they manually chose to work online.
- We fixed an issue where running the VBA code ActiveInspector.CommandBars.ExecuteMso("ShowSchedulingPage") after enabling the Single Line Ribbon (SLR) would result in a runtime error.
- We fixed an issue where the 'OK' and 'Cancel' buttons on the Automatic Replies dialog wouldn't be visible on a system with a high resolution (such as 1750 x 1920) combined with a large text size (such as 175%).
- We fixed a condition where sending a meeting request from an empty contact group to another contact group would result in a unexpected close.
- We fixed an issue that caused users to experience a unexpected close when opening certain very large emails.
- We fixed an issue where if Group Policy requires an add-in to be always enabled, then monitoring add-in's becomes unavailable in order to prevent users from disabling the add-in.

### PowerPoint

- We fixed an issue where videos weren't playing automatically in slideshows.
- We fixed an issue where after booting PowerPoint, inserting a slide and opening and closing the comments pane, the slides in the thumbnail pane displayed as being overlapped.

### Project

- We fixed an issue where if a resource has multiple cost rate tables, the remaining cost may not be calculated correctly.
- We fixed an issue where the Project finish date wasn't getting updated for projects connected to SharePoint tasks list.

### Word

- We fixed an issue where the Comment card would display a border around the comment text if the user clicked on the comment.
- We fixed an issue where the focus wouldn't go to the comment pane if the document was zoomed to 160% or more and the comment pane wasn't visible.
- We fixed an issue where the comments on one document would be displayed on other open documents after switching between the multiple open documents.
- We fixed an issue where if a user created a comment draft anchored to a line already containing committed comments, then the draft was arranged in the wrong order relative to the committed comment in the SideTrack.
- We fixed an issue where long links weren't being wrapped when saving a document to HTML format.
- We fixed an issue with macros in which AutoOpen runs before AutoExec.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2009: August 28

*Version 2009 (Build 13219.20004)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Outlook

- Addresses an issue that caused users to be able to send email content that had a "Do Not Forward" policy applied to OneNote when selecting more than one message.

### PowerPoint

- We fixed an issue where the functionality to insert a video was disabled.

### Word

- We fixed an issue where the user couldn't exit the Header/Footer when selecting a comment.
- We fixed an issue that prevented users from seeing comment threads that exceeded the sidetrack boundary because scrolling through the sidetrack wasn't working.
- We fixed an issue where searching for resolved comments in the sidetrack pane wasn't working.

### Office Suite

- We fixed an issue in the Office Deployment Tool where configuration was failing when using the RemoveMSI feature with the Office 2007 "Microsoft Application Error Reporting" product present.
- We fixed an issue in the Compress Picture dialog where some user-selected DPI settings aren't retained.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2009: August 21

*Version 2009 (Build 13212.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Action Pen in Excel:** Pen Tool to help you handwrite and make quick edits to your data

### Outlook

- **Delete conversation by message owner:** This feature allows you to delete a conversation by message owner.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an issue where connections to an ODBC database weren't working with third party applications.

### Excel

- We fixed an issue where using a macro to set the FormulaR1C1 property for a range, the cell references would be incorrect if a chart sheet was the active sheet.
- We fixed an issue where inking could cause Excel to become unresponsive.

### Outlook

- We fixed an issue where users can now disable IRM (Information Rights Management) for Outlook without having to disable it for the rest of the Office applications.

### Word

- We fixed an issue where Word could close unexpectedly after comments were deleted.
- We fixed an issue where in some cases, bullets aren't displaying correctly in email.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2009: August 14

*Version 2009 (Build 13205.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where if a user typed a formula name including the parenthesis and invoked help via F1, the help topic specific to that formula wouldn't be displayed.
- Fixed an issue where macros assigned to buttons were broken after restoring an older version of the file.

### Outlook

- This change fixes an issue where the Meeting page would continue to be displayed after the user switched tabs from the Meeting page to the Scheduling Assistant page.

### Word

- We fixed an issue where the bullet picture icon didn't display correctly.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2009: August 07

*Version 2009 (Build 13130.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Outlook

- We fixed an issue where the user account attributes in Active Directory for "otherTelephone" and "otherHomePhone" weren't mapped to the corresponding Outlook LDAP attributes.

### PowerPoint

- We fixed an issue where users were seeing the ribbon/title bar not being displayed under certain conditions.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2008: July 31

*Version 2008 (Build 13127.20002)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Insert your iPhone photos directly into Office:** HEIC pictures from your phone now insert seamlessly into Office. No conversion required.<br />See details in [blog post](https://insider.office.com/blog/insert-apple-photos-into-office-easily)

### Outlook

- **Insert your iPhone photos directly into Office:** HEIC pictures from your phone now insert seamlessly into Office. No conversion required.<br />See details in [blog post](https://insider.office.com/blog/insert-apple-photos-into-office-easily)

### PowerPoint

- **Insert your iPhone photos directly into Office:** HEIC pictures from your phone now insert seamlessly into Office. No conversion required.<br />See details in [blog post](https://insider.office.com/blog/insert-apple-photos-into-office-easily)

### Word

- **Insert your iPhone photos directly into Office:** HEIC pictures from your phone now insert seamlessly into Office. No conversion required.<br />See details in [blog post](https://insider.office.com/blog/insert-apple-photos-into-office-easily)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- This fix addresses the issue where trying to run certain queries have previously produced the error message 'Query is too complex.'

### Excel

- We fixed an issue where if the order of a chart series was changed, the corresponding checkbox aligned with the series wasn't reordered along with the series.
- We fixed an issue where a copy of an image with a radial gradient fill didn't match the original.

### Outlook

- This fix addresses an issue that caused users to be unable to add a signature when replying to a digitally rights managed message from an inspector window when the user didn't have Owner permissions on the message being replied to.
- This fix addresses an issue that was causing Outlook to fail to display line breaks properly in markdown content.

### PowerPoint

- We fixed an issue where a copy of an image with a radial gradient fill didn't match the original.
- We fixed an issue where the Forms button in PowerPoint didn't allow the creation of Forms when access to the Office Store wasn't permitted.

### Project

- We fixed an issue where for a SharePoint tasks list, the ribbon buttons on the second tab may be disabled.

### Word

- We fixed an issue where a copy of an image with a radial gradient fill didn't match the original.
- We fixed an issue where if a comment was added to track a change, the revisions pane would unexpectedly open.
- We fixed an issue where links to documents weren't being inserted to the comments box via the Insert -> Link dropdown.
- We fixed an issue where the hyperlink count in the VBA hyperlinks collection wasn't iterating correctly after adding an image containing a hyperlink.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2008: July 24

*Version 2008 (Build 13117.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Make polished Visio diagrams in Excel:** Create a flow chart or organizational chart by putting data on a worksheet. [Learn more](https://support.office.com/article/bee3b5aa-aaaf-4401-acc6-276b711c763c)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### OneNote

- We fixed an issue where the placeholder text in the Search edit box would overflow if the application window was resized to a small dimension.

### Word

- We fixed an issue where the placeholder text in the Search edit box would overflow if the application window was resized to a small dimension.
- We fixed an issue where the 'Specific People' option for Track Changes was disabled.
- We fixed an occasional instance where the app stopped responding while opening HTML files.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2008: July 17

*Version 2008 (Build 13115.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where any time a pivot chart with hidden leader lines was saved and reopened, the leader lines would become visible.

- We fixed an issue where charts weren't always updated as expected when 'ForceFullCalculation' was enabled via VBA for the workbook.

### Outlook

- We fixed an issue around creating multiple profiles in Outlook from the same email domain.
- Addresses an issue that caused the lock icon to fail to display in the header of S/MIME encrypted messages.
- Addresses an issue that caused attachments to get stripped from S/MIME messages when sent unencrypted.
- Addresses an issue that caused plain text S/MIME messages to become garbled when sending.
- Addresses an issue that caused attachments to become corrupted when sending an S/MIME email unencrypted.
- Addresses an issue that caused users to be unable to save OneDrive attachments from outside their tenant to their local computer when selecting the 'Save' option on the security dialog.
- Addresses an issue that caused recipients to be unable to save rights protected messages even when the save as permission was granted by the sender.
- Addresses an issue that caused the labels for some Advanced Search options to be truncated in some languages.

### Project

- We fixed an issue where tasks listed in the Task Board view weren't in sync with those in the Assign Resources dialog.
- We fixed an issue where if you copied and pasted a task that had multiple dependencies, not all dependencies were copied correctly.

### Word

- We fixed an issue where the 'Editor' command was disabled when the focus was in a comment text box.
- We fixed an issue where the 'Show Markup' command was disabled when the focus was in a comment text box.
- We fixed an issue in custom XML where the state of comments may be lost when opening the document.

### Office Suite

- We fixed an issue where after the user opened a new app window from the taskbar and created a new blank document, more files were created.
- We fixed an issue where if a user was editing a document but had lost permissions, we weren't notifying the user that they had to re-authenticate.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2008: July 10

*Version 2008 (Build 13102.20002)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Outlook

- We fixed an issue where the Allow Forwarding option was missing from the shared calendar meeting "Response Options" if Download shared folder wasn't checked.
- We fixed an issue that would display the print button in a disabled state even though the user had the appropriate print permissions.

### Project

- We fixed an issue where if you tried to save a PDF/XPS from Project to a SharePoint document library, nothing would happen.

### Word

- We fixed an issue where Word would stop responding after pasting some text and an image in to a comments box.
- We fixed an issue where the 'New comment' button would be disabled after deleting the last comment.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2007: July 03

*Version 2007 (Build 13029.20006)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Outlook

- **Create polls in Outlook with Quick Poll:** Easily create a poll, collect votes, and view results within an email [Learn more](https://support.office.com/article/46893563-ab12-4bd0-aff7-26f5a488fea0)

- **New room finder:** Search for conference rooms by different capabilities.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where data model tables created in certain versions of Excel couldn't be seen in 'Table Preview' even though the query associated with the table had not been edited.
- We fixed an issue where disabling 'Ignore Relative/Absolute' references in the Define Name \ Apply Names dialog would cause formulas to not work.
- We fixed an issue where clearing an advanced data filter could lose table formatting.
- We fixed an issue where the full path of an embedded PDF document would show in the document caption rather than just the filename.
- We fixed an issue where after disabling the Wolfram cloud connector and then saving and re-opening an Excel workbook, could result in a unexpected close.
- We fixed an issue where booting Excel with the Solver add-in enabled would result in a unexpected close.

### Outlook

- We fixed an issue where Outlook would stop responding if there were over 130 recipients on the 'To' line and we also improved the performance of rendering the text.
- We fixed an issue in the 'To Do Bar' where events that spanned more than two days, displayed the same end time for all subsequent days.
- We fixed an issue that caused users of Outlook to see their message list stop updating for several minutes after using shared calendars.

### PowerPoint

- We fixed an issue where pasting HTML to a text area on a slide would instead get pasted into a text box created at the top of the slide.
- We fixed an issue where selecting all slides in Presenter View, then exiting Presenter View using Alt+Tab and returning to the slide show and clicking 'End Show' would result in an unhandled exception.

### Project

- We fixed an issue where Project may close unexpectedly when opening certain XML files.
- We fixed an issue where you couldn't open a Project file from a SharePoint document library if the library was in modern mode.
- We fixed an issue where projects couldn't be opened in the Project desktop client from the Project Web App if the URL ended in '.com'.

### Word

- We fixed an issue during co-authoring mode when there is a merge conflict and the user has already chosen to discard changes, we no longer display the option to save or discard changes.
- We fixed an issue that, when attempting to save a file containing a macro under a new name, would cause it to be saved with a .docx extension and the filename 'WRO0004.docx', regardless of what the user entered, rendering the document unusable.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2007: June 26

*Version 2007 (Build 13020.20004)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an issue where the linked table manager would prompt for a primary key if a linked SQL table was refreshed.
- We fixed an issue where queries in the Query Editor scrolled out of view.
- We fixed an issue where query execution was taking approximately twice as long to complete than expected.

### Outlook

- We fixed an issue where users were unable to 'Send As' or 'Send on behalf' of a distribution list.
- We fixed an issue where inserting an image inline in a message, then saving the message as a draft would result in a resizing of the image.
- We fixed an issue that caused the body of an NDR message to change from Unicode to ASCII after editing the subject.

### Project

- We fixed an issue where Project Planner links in Government Community Cloud environments had been disabled.

### Office Suite

- We fixed an issue where text inserted in a Scalable Vector Graphic (SVG) was illegible after inserting it in a Word, Excel, or PowerPoint file, saving and closing the file, and then re-opening the file.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2007: June 19

*Version 2007 (Build 13012.20000)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where CustomUI XML for a custom ribbon tab was removed when saving a workbook to SharePoint/OneDrive.
- We fixed an issue where workbooks were read-only when the file only had read-only recommended.

### Outlook

- We fixed an issue where the Input Method Editor (IME) window would overlap the underlying text being entered via the IME when using multiple monitors with different resolutions.
- We fixed an issue that caused users to see the following error when closing an appointment that was previously saved: "The item cannot be saved because it was changed by another user or in another window. Do you want to make a copy in the default folder for the item?"
- We fixed an issue where dates in the mini calendar failed to display in bold for users in Japan.
- We fixed an issue that prevented calendar reminders from showing exact times for meetings coming up in less than a week.

### PowerPoint

- We fixed an issue where a user's presence color indicator wasn't getting refreshed in the co-authoring gallery during a live co-authoring session.

### Project

- We fixed an issue where if Fixed Duration tasks are at 100% complete but the Actual Finish isn't specified, the Task % Complete would display as less than 100%.

### Word

- We fixed an issue where the HTML hyperlink color wasn't being rendered correctly.

### Office Suite

- We fixed an issue where URLs that weren't http or https based weren't being displayed in the Most Recently Used list.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2007: June 12

*Version 2007 (Build 13006.20002)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Get Organization Data from Power BI using Data Types:** Excel data types from Power BI are now rolling out to Insiders in organizations with Office 365 E5/A5 or Microsoft 365 E5/A5. Getting the information you need and easily refreshing it's critical to many everyday workflows. We're giving you access to your company or organization information from Power BI as a data type in Excel, which expands your ability to bring in linked information in your spreadsheets. [Learn more](https://support.office.com/article/cd8938ce-f963-444d-b82a-7140848241e9)
See details in [blog post](https://insider.microsoft365.com/blog/use-power-bi-data-in-excel)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- We fixed an issue that caused Microsoft Access to fail to identify an Identity Column in a linked SQL Server table, which could cause rows to be reported as deleted incorrectly.

### Excel

- We fixed an issue where the major grid lines of radar charts couldn't be formatted correctly.

### Project

- We fixed an issue where the ProjectBeforeTaskChange event didn't fire when there was a change to the project summary task, either the project start/task field.
- We fixed an issue where a baseline reset or update could change time-phased budget cost/work resources and the baseline could reflect incorrect budget values.

### Word

- We fixed an issue where the ability to clear formatting within the Comments pane via the Clear Formatting button in the Office Ribbon wasn't working.
- We fixed an issue where changing the size of a table when the ruler isn't displayed caused other applications running in the background to start flashing.
- We fixed an issue where in co-authoring mode, comment replies would sometimes not show up in the comments pane but would be visible in the revisions pane.
- We fixed an issue where if Word had a list of more than 50 frequently opened documents, then after saving and opening a document, a revision history would be displayed even though no revisions were made to that document.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2006: June 05

*Version 2006 (Build 13001.20002)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Sort/filter while collaborating in Excel:** You can now sort and filter your Excel file while collaborating with others. This new feature prevents you from being impacted by other user's sorts and filters while coauthoring the document.

- **Create PivotTables from Datasets in Power BI within Excel:** You can create PivotTables in Excel that are connected to datasets stored in Power BI with a few clicks. Doing this allows you get the best of both PivotTables and Power BI. Calculate, summarize, and analyze your data with PivotTables from your secure Power BI datasets. [Learn more](https://support.office.com/article/31444a04-9c38-4dd7-9a45-22848c666884)

### Outlook

- **Quickly reopen items from previous session:** We added an option to quickly reopen items from a previous Outlook session. Whether Outlook closes unexpectedly or you close it, you'll now be able to quickly relaunch items when you reopen the app. This feature is on by default. To turn it off, go to Options > General > Start up Options.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where custom values on the chart axis wouldn't get applied correctly.
- We fixed an issue where worksheets containing multiple formulas with defined names was resulting in longer times when saving files.

### Outlook

- We fixed an issue where the IME (Input Method Editor) window would overlap the underlying text being entered via the IME when using multiple monitors with different resolutions.
- We fixed an issue where viewing a template when composing a new email message would result in a unexpected close of the app.
- We fixed an issue where users were unable to Exchange 2010 public folders after Outlook version 1911.
- We fixed an issue where the Categorize button for group calendars in the Office Ribbon was disabled.
- We fixed an issue that would cause users with conflicting contacts to experience unexpectedly closes in Outlook.
- We fixed an issue that resulted in the Online Archive dropdown in folder properties to be missing for users on high DPI monitors.
- We fixed an issue that caused users to experience a unexpected close in Outlook when working with hyperlinks in Plain Text emails.
- We fixed an issue that caused Outlook to be unable to parse long file names encoded with RFC2231.
- We fixed an issue that was causing Outlook users to experience intermittent hangs when using screen readers.

### PowerPoint

- We fixed an issue with opening server-configured files with forms-based authentication.
- We fixed an issue where PowerPoint files with embedded charts / workbooks could result in failures when saving the file.
- We fixed an issue where a Comment pane that had been closed by the user would re-open automatically.
- We fixed an issue where the slide editor from one slide would overlap on to the next slide.

### Project

- We fixed an issue that prevented orphaned tasks from being deleted or re-assigned after their parent plan was deleted.

### Word

- We fixed an issue where timestamps in Comment panes weren't based on the system locale time.
- We fixed an issue where comments between the web app and the desktop application weren't in sync.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2006: May 29

*Version 2006 (Build 12920.20000)*

### Feature updates

### Outlook

- **More buttons added to Outlook toast notifications:** Quick Action buttons now appear in Outlook toast notifications when running Outlook on Windows 10.

### PowerPoint

- **Improved Stream video performance in PowerPoint:** We've made improvements to the playback performance of Microsoft Stream videos to minimize video loading time and create a smooth viewing experience.  Use your corporate videos from Microsoft Stream to create better presentations.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where Excel would occasionally shut down when engaging OneDrive.
- We fixed an issue where clicking a bookmarked hyperlink within the same workbook would cause the workbook to be hidden.
- We fixed an issue where some copy and pasted chart links used mapped drive addresses rather than universal addresses.
- We fixed an issue where Excel could become unresponsive after using Ctrl+Shift+Arrow keys to scroll when the Excel window was shared through Teams.

### PowerPoint

- We fixed an issue where slides weren't centered after zooming using the mouse wheel.

### Word

- We fixed an issue where copy and pasting text to a comment pane wouldn't be displayed.
- We fixed an issue where comment hint bubbles appeared blurry at 100% zoom.
- We fixed an issue where enabling policy Word 2007 and later Binary Documents and Templates would cause some co-authoring cases to fail.
- We fixed an issue where files with long path names (greater than 32K) wouldn't open and an appropriate error message wasn't being displayed.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2006: May 22

*Version 2006 (Build 12914.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Save to Pinned Folders:** Pin your folders makes saving Office files easier. We received feedback that users want more control over the folders available when a new file is saved. We're excited to bring a new capability to you: pin your folders in the Save dialog. This new capability will make saving your Word, Excel, and PowerPoint files easier.
See details in [blog post](https://insider.microsoft365.com/blog/pin-your-folders-makes-saving-office-files-easier)

### PowerPoint

- **Save to Pinned Folders:** Pin your folders makes saving Office files easier. We received feedback that users want more control over the folders available when a new file is saved. We're excited to bring a new capability to you: pin your folders in the Save dialog. This new capability will make saving your Word, Excel, and PowerPoint files easier.
See details in [blog post](https://insider.microsoft365.com/blog/pin-your-folders-makes-saving-office-files-easier)

### Word

- **Save to Pinned Folders:** Pin your folders makes saving Office files easier. We received feedback that users want more control over the folders available when a new file is saved. We're excited to bring a new capability to you: pin your folders in the Save dialog. This new capability will make saving your Word, Excel, and PowerPoint files easier.
See details in [blog post](https://insider.microsoft365.com/blog/pin-your-folders-makes-saving-office-files-easier)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue where the error message: “This workbook is currently referenced by another and cannot be closed” would appear because add-ins were being loaded in alphabetical order rather than in a user specified order.
- We fixed an issue where memory was being corrupted when managing fonts between Excel and some third party assistive technology applications.
- We fixed an issue where Excel would close unexpectedly when add-ins ask for Host Items on worksheets that contain shapes with noSelect locks.

### Outlook

- We fixed an issue where the Folder.BeforeItemMove event didn't fire correctly when a user moved items between folders.
- We fixed an issue where users saw calendar items that spanned the midnight threshold as All day events.
- We fixed an issue where Outlook closed unexpectedly when two add-ins added a button to the same group in the ribbon.
- We fixed an issue where users were unable to share a calendar with a guest.

### PowerPoint

- We fixed an issue where zooming in and out from the presentation area resulted in a gap between the zoomed selection marquee and the mouse pointer.

### Project

- We fixed an issue where Project would close unexpectedly after clicking on Options.

### Word

- We fixed an issue where hyperlinks in comments weren't working.
- We fixed an issue where zooming in and out from the presentation area resulted in a gap between the zoomed selection marquee and the mouse pointer.
- We fixed an issue where pasting HTML into WordMail for Calendar wasn't working.
- We fixed an issue where replying to a comment in a coauthored session could sometimes cause Word to freeze.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2006: May 15

*Version 2006 (Build 12905.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Make a PDF connection:** Connect to, import, refresh data from a PDF. [Learn more](https://support.office.com/article/9967afd8-85ee-4df3-aa06-753bcc1a2724)

- **Auto-apply or recommend sensitivity labels:** Office can recommend or automatically apply a sensitivity label based on the sensitive content detected.

### Outlook

- **Find just what you need:** Narrow your search with options like folder, sender, date, attachment info, and more.

### PowerPoint

- **No need for a clicker: your earbuds have you covered:** Use your Surface Earbuds to control your PowerPoint presentations. Important: You must pair your Surface Earbuds in the Surface Audio app for Windows 10 in order to use gestures to control presentations. Instructions for getting started with the Surface Audio app on Windows 10 are available here. [Learn more](https://support.office.com/article/6319a6f3-ad69-44e6-a8ff-e79676423e4a)

- **Auto-apply or recommend sensitivity labels:** Office can recommend or automatically apply a sensitivity label based on the sensitive content detected.

### Word

- **Auto-apply or recommend sensitivity labels:** Office can recommend or automatically apply a sensitivity label based on the sensitive content detected.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- We fixed an issue that resulted in improved performance time for users when they deleted merged columns.
- <div>We fixed an issue that caused printer names to be duplicated in the list of available printers.</div>

### PowerPoint

- We fixed an issue where keyboard shortcuts and spell check wouldn't function as expected when using an English Switzerland (QWERTZ) keyboard.

### Word

- We fixed an issue where adding a new comment on a blank document wouldn't do anything.
- We fixed an issue where inserting or updating an Index in a document containing more than a hundred entries would result in the application closing unexpectedly.
- We fixed an issue where files with custom xml values opened extremely slowly.

### Office Suite

- We fixed an issue in Visual Basic for Applications in Microsoft Office where certain VBA projects that contain references to code libraries with DBCS characters in the library name or library path would be viewed by the Office application as corrupt on load.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2006: May 08

*Version 2006 (Build 12829.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Tell your stories with animated GIFs:** Animated GIFs are now supported in the Office editor - your documents just got snazzier.

### Outlook

- **Better results—in a jiffy:** We've updated the Search experience to make it smarter, faster, and more reliable than ever. [Learn more](https://support.office.com/article/96fee452-80cd-492d-a35c-5c37584b416b)

- **Tell your stories with animated GIFs:** Animated GIFs are now supported in the Office editor - your documents just got snazzier.

### PowerPoint

- **Tell your stories with animated GIFs:** Animated GIFs are now supported in the Office editor - your documents just got snazzier. [Learn more](https://support.office.com/article/3a04f755-25a9-42c4-8cc1-1da4148aef01)

### Word

- **Tell your stories with animated GIFs:** Animated GIFs are now supported in the Office editor - your documents just got snazzier.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Office Suite

- We have investigated and resolved the issue where an Office 365 ProPlus deployment via Intune is paused after an OS shutdown.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2005: May 01

*Version 2005 (Build 12827.20030)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Outlook

- **Improved links in email:** When you include a link to a file, the file name replaces the URL. You can change permissions so all recipients have access. [Learn more](https://support.office.com/article/02040f47-bd56-4806-8311-fc913fed54c0)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- Fixed an issue where chart data table could render values in a date axis incorrectly.
- Fixed an issue where page breaks couldn't be disabled after going into Page Layout or Page Break Preview.
- Fixed an issue where chart line styles could be lost after hiding and unhiding columns with series data.
- Inserting a column in a filtered list would take longer than expected.
- A unexpected close could occur when trying to list changes on a new sheet for a workbook using legacy "Shared Workbook" mode.
- Fixed the issue where custom formatting in Pivot charts may not be saved when the "Invert if negative" option was enabled.
- Fixed an issue where custom formatting for a single data point in a Pivot chart wasn't saved if the "Invert if negative" option was selected.
- This change fixes an issue where the '@' character uploaded in a CSV file, would result in the string after the '@' character to be converted to a formula.
- Fixed an issue where decimal values in the SEQUENCE function weren't rounded correctly.

### Outlook

- Addresses an issue that caused very long safelinks that users clicked on in the Outlook Desktop client to fail to load due to truncation.
- Fixed an issue where Outlook folders with names containing DBCS (Double Byte Character Set) characters would intermittently disappear when synchronizing with the server. For this to happen, Outlook had to be configured with an IMAP account and running on a system with the locale set to Japanese.

### PowerPoint

- Fixed an issue where if a user created a comment without posting it and closed the Comments pane, then opened a new window, navigated through multiple slides and, closed the window, and finally re-opened the Comments pane in the original presentation, the draft comments wouldn't be available.

### Project

- Fixed an issue where if Project is connected to the Project Web App and the decimal separator is a comma, TaskDependencies Add method fails when Lag is added.

### Word

- Fixed an issue where inserting comments on a document in collaboration mode wouldn't always work.
- This change fixes an issue where the People card would flash if the @ mention was clicked.
- Fixed the issue where closing a document with draft comments would prompt the user if they wanted to close the document without saving the draft comments. Cancelling the prompt would close the document rather than leaving it open.
- Fixed an issue where translating a posted comment would result in the error "Inserting translated text failed".
- In Web View/Immersive reader, clicking on a hint would scroll to the top even though it was already in view. This has been fixed.
- We fixed an issue that, when attempting to save a file containing a macro under a new name, would cause it to be saved with .docx extension and the filename WRO0004.docx, regardless of what the user entered, rendering the document unusable.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2005: April 24

*Version 2005 (Build 12816.20006)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- This change fixes an issue where the chart trendline R-Squared value (in the forced y-intercept case) was incorrect even though the LINEST function returns the correct value.
- This change fixes an issue where customized chart trendline formatting wasn't always being saved.

### Outlook

- Fixed an issue where the Categorize button for group calendars in the Office Ribbon was disabled.
- Fixed an issue where enterprise customers with group folders not implemented or not working, would result in Outlook displaying a "not responding" message.

### PowerPoint

- Fixed an issue where hovering over the asterisk (*) symbol didn't display the user name and date of the last person to update the document.

### Word

- Enabling the option "Show bookmarks" wouldn't display bookmarks. This has been fixed.
- This change fixes an issue where text with hyperlinks may not display if the option "Show field codes instead of their values" was enabled.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2005: April 17

*Version 2005 (Build 12810.20002)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- Increased the size of the cell reference edit controls on the Custom Error Bars dialog used with charts.
- Workbooks saved with a digital signature in Excel 2016 could have the signature invalidated upon opening in the current version of Excel.
- Fixed a problem with the scaling of checkboxes in form controls when printed.
- Application.Evaluate (VBA) wasn't working for User-defined functions in some cases.
- This change fixes an issue where conditional formatting (CF) information wasn't being saved to XLSB files correctly.

### OneNote

- Fixed an issue where line breaks were being stored as vertical tabs.

### Outlook

- Addresses an issue that caused users to be unable to add a Personal Contact Group as a Meeting attendee.
- Addresses an issue that caused meetings that are more than 2 months away to fail to display a meeting subject in the Scheduling Assistant.
- Addresses an issue that caused users to see message body truncation when forwarding large HTML messages.
- Added the ability to enforce S/MIME default signing configuration via group policy.
- Addresses an issue that caused delete rules created for mailboxes other than the user's primary mailbox to become invalid.
- Addresses an issue that caused attachments to get dropped when forwarding an encrypted message.

### Project

- When Predecessor/Successor data is edited within a Form view, an extra ProjectBeforeTaskChange event is fired.
- Fixed an issue where Project may close unexpectedly when changing the board status field on a project that is connected to a SharePoint task list.
- Fixed an issue where Project may close unexpectedly when saving projects created with older versions of Project.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2004: April 10

*Version 2004 (Build 12730.20024)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Access

- **Bypass the Show Table dialog box, go directly to a task pane, and streamline adding tables to relationships and queries.:** Add tables and queries by clicking four tabs, multi-selecting names, searching by text, and dragging from a task pane that stays open while you work.

### Excel

- **M365 Premium Content Picker:** Bring your documents to life! Explore 1000's of royalty free stock images, icons and stickers [Learn more](https://support.office.com/article/3c51edf4-22e1-460a-b372-9329a8724344)

### Outlook

- **M365 Premium Content Picker:** Bring your documents to life! Explore 1000's of royalty free stock images, icons and stickers [Learn more](https://support.office.com/article/3c51edf4-22e1-460a-b372-9329a8724344)

- **Keep your pictures high fidelity when sending them as part of an email:** A new Outlook setting is available to limit picture compression when you send pictures as part of the email contents

### PowerPoint

- **M365 Premium Content Picker:** Bring your documents to life! Explore 1000's of royalty free stock images, icons and stickers [Learn more](https://support.office.com/article/3c51edf4-22e1-460a-b372-9329a8724344)

- **Synchronize changes while you are presenting:** Synchronize changes whenever they're made even when the presentation is in slide show mode.

### Word

- **M365 Premium Content Picker:** Bring your documents to life! Explore 1000's of royalty free stock images, icons and stickers [Learn more](https://support.office.com/article/3c51edf4-22e1-460a-b372-9329a8724344)

- **Annotate your private copy:** Create hand written notes for your eyes by making a private copy of a shared document. Go to View > Create a Private Copy to get started.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- Fixed issues with resizing and refreshing tables in the task pane.

### Excel

- Fixed an issue where selecting a range of cells on a sheet would result in the selection of a single cell.

- Fixed an issue, which could cause Excel to stop responding when reducing the size of a chart with some x-axis ranges.

- Fixed an issue where inserting a user defined chart template as default would result in saving it as a column chart.

- Fixed an issue where Data labels on charts would display as blank when the underlying data cells didn't have a caption.

- Fixed an issue where an Excel sheet with R1C1 cell referencing enabled and is being coauthored / shared, hovering over the user presence icon doesn't display the active cell reference in R1C1 mode.

### Outlook

- Addresses an issue that caused categories to occasionally disappear from email messages.

- Addresses an issue that caused delegates to see different folder hierarchies on different machines for shared mailboxes.

- Addresses an issue that caused users to experience a close unexpectedly when attempting to view the properties of an Organizational Form.

- Addresses an issue that caused some reminders to fail to fire when changing the timezone on a machine.

### PowerPoint

- This change fixes an issue where the rendering of a legacy Excel chart embedded as an OLE object in PowerPoint or Word may not always display the chart title.

- We have fixed an issue when copying text from Excel to PowerPoint might change its formatting.

- This change fixes an issue where finding special characters using 'find whole words only' didn't always work as expected.

### Project

- Fixed an issue where the user couldn't enter time-phased Baseline Work when the setting to protect actual work is on.

### Word

- This change fixes an issue where hovering a cursor over a hint wouldn't highlight its card.

- This change fixes an issue that would cause the text in grouped shapes to disappear temporarily when using the Lasso selection tool.

- This change fixes an issue where the rendering of a legacy Excel chart embedded as an OLE object in PowerPoint or Word may not always display the chart title.

- This change fixes an issue in two page view, when creating a comment, the comment anchor didn't always come into view.

- Fixed an issue where if a paragraph whose style is an ancestor of a style linked to a list, then the numbering of that list could be lost.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2004: March 27

*Version 2004 (Build 12718.20010)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Outlook

- **New option to disable @ mention suggestions when composing mail:** Do you find the @ mention picker more annoying than useful? Now you can turn it off if you prefer.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Outlook

- Addresses an issue that caused the "Save to Cloud" button to be missing from Attachment Tools.
- Addresses an issue that caused users to be unable to add a signature when replying to a digitally rights managed message from an inspector window when the user doesn't have Owner permission on the message being replied to.
- Addresses an issue that caused users to be unable to add more attachments from a web location to a previously created meeting.

### PowerPoint

- This change fixes an error that could cause PowerPoint files containing emojis to fail when saving.

### Project

- Fixed an issue where if 'CustomFieldValueListGetItem' is executed and a lookup table for the custom field doesn't exist, an empty lookup table is created even though it shouldn't be.

### Word

- This change fixes an issue with multiple pages selected from the View menu, where the comments pane could be displayed as blank.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2004: March 20

*Version 2004 (Build 12711.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Outlook

- **Calendar visual refresh:** Last year, we brought you a refreshed mail experience, and, this year, it's the calendar's turn to get a facelift! The updates are fresh but familiar so, as a seasoned Outlook user, you can jump in and be more productive right away.

- **Help protect data in your group:** The Sensitivity label you choose when creating a group is applied to group email, documents, and team sites

### PowerPoint

- **Update slides during slide show:** Update slides changed by other authors during your presentation.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- This change addresses delays when processing images with malformed or invalid protocol information.

### Outlook

- This change addresses delays when processing images with malformed or invalid protocol information.

- This change fixes an issue where the latest changes to draft emails weren't being updated.

- Fixed an issue where right-mouse clicking on a file and using 'Send to' wouldn't work.

- Fixed an issue where if a user had a customized the search path for the Address book, Outlook's name resolution scope would be limited to the customized path rather than including the Global Address List (GAL).

- Fixed an issue where within a set of returned search results, sorting the results by Categories wouldn't display the Category colors.

### Project

- Fixed an issue where the 'ProjectBeforeTaskChange' Visual Basic Applications (VBA) event didn't fire when a user clicked the "Inactivate" button found on the Tasks Ribbon within the Scheduling grouping.

- If you set predecessor or successor details from within a Form type view, the
ProjectBeforeTaskChange Visual Basic Applications (VBA) event didn't always capture the changes. For example, if you deleted a dependency and clicked OK on the form, the event didn't fire. This behavior has been fixed.

- Fixed an issue where the latest values for the Actual Cost of Work Performed (ACWP) wouldn't be displayed after making a change, such as a date change.

- Fixed an issue where opening a project using the Most Recently Used (MRU) menu opened the project file with Read/Write access.

- This change fixes an issue where if you created a manual task with a start date and a time (but no duration), it would be displayed with an incorrect time on the timeline.

- Fixed an issue where printing a timeline using Hijri calendar would result in a month being skipped or duplicated in the print view.

- This change addresses an issue where working in Team Planner with GDI objects, could result in the over allocation of GDI objects and create low memory conditions.

### Word

- Fixed an issue where the functionality to post comments was disabled.

- This change addresses delays when processing images with malformed or invalid protocol information.

- This change addresses an issue where the account manager wouldn't dispatch messages resulting in the app to stop responding with third party applications.

- This change fixes an issue where the Table of Contents would get updated with heading styles, which weren't present in the document.

- Fixed an issue where digital signatures saved in Word documents would be removed when mailing the documents.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2004: March 13

*Version 2004 (Build 12703.20010)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Excel

- **Sensitivity labels**: You can now apply a sensitivity label that your organization has configured to prompt you for custom permissions. [Learn more](/microsoft-365/compliance/encryption-sensitivity-labels)

### PowerPoint

- **Sensitivity labels**: You can now apply a sensitivity label that your organization has configured to prompt you for custom permissions. [Learn more](/microsoft-365/compliance/encryption-sensitivity-labels)

### Word

- **Sensitivity labels**: You can now apply a sensitivity label that your organization has configured to prompt you for custom permissions. [Learn more](/microsoft-365/compliance/encryption-sensitivity-labels)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- Fixed an issue where international versions of Access were displaying English strings in the user interface.

### Excel

- Fixed a performance issue that users may have experienced when programmatically editing a large range of cells.
- Fixed a performance issue that occurred when opening csv files with Japanese environments.

### Outlook

- Addresses an issue that caused the "Last Modified"; date on a file to be updated when adding an attachment to a mail or saving an attachment from a mail by dragging and dropping it (as opposed to via a menu).
- Addresses an issue that caused hitting enter in the expanded find pane to fail to start a search, requiring instead that users click on the search button.
- Fixed an issue where search shows no information about users when the option to "Show user photographs when available" is disabled.

### Project

- Fixed an issue where summary task dates weren't always getting calculated correctly.
- Fixed an issue where the OnUndoOrRedo event doesn't fire without first running the OpenUndoTransaction method.

### Word

- Fixed an issue when typing or editing a comment and using Ctrl+A would result in selecting text in the canvas instead of selecting text just within the comment card.
- We fixed an issue in which the alignment of words in a document gets scrambled when tried to edit after printing using Quick Print.
- We fixed an issue when merging two documents into one document.
- Fixed an issue where marking revisions involving equations could result in a failure when saving the file.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2003: March 06

*Version 2003 (Build 12624.20086)*

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Outlook

- Fixed an issue where creating a rule with Outlook Web Access didn't persist to the Exchange server and resulted in a conflict.
- Fixed an issue with Outlook in dark mode wouldn't display the drop down list in the 'From:' field.
- Addresses an issue that caused users to be unable to attach a file to their mail message via the file explorer when that file was open in another application.

### PowerPoint

- Fixed an issue where the recommended thumbnails flash when hovering your mouse over the thumbnails. In some cases this could cause PowerPoint to close unexpectedly.

### Word

- Fixed an issue with Compare feature for documents that were protected for editing.

### Office Suite

- Fixed an issue Word/Excel/PowerPoint where the User Principal Name (UPN) is no longer case sensitive resulting in less failures when working with files on SharePoint.

- Fixed a cosmetic issue where the 'OK' button on the File \ Options dialog was being displayed as grayed out but functionality wasn't impacted.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

## Version 2003: February 28

*Version 2003 (Build 12619.20002)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Outlook

- **Incident Notification for IT Admins:** Office Apps Administrators are notified about Outlook and O365 Exchange incidents affecting their users with a new right-side panel notification in Outlook for Windows. [Learn more](https://support.office.com/article/46c07f08-1277-41ce-b353-4e205e9da333)

### PowerPoint

- **Improved ink to shape diagramming experience:** Draw better diagrams and have it convert so office object you can manipulate [Learn more](https://support.office.com/article/f304ef73-9514-450b-9bb9-28c6057020f2)

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- Fixed an issue where text in a slicer isn't scaled properly in Print Preview.

### Outlook

- Addresses an issue that caused the "Last Modified" date on a file to be updated when adding an attachment to a mail or saving an attachment from a mail by dragging and dropping it (as opposed to via a menu).

### Word

- Fixed an issue that when tabbing through a comment card, the focus on the comment edit box wouldn't be visible.

- Inserting a control (such as a Text Content Control) in an equation then saving and opening the file results in an un-readable content error.

- Fixed an issue where saving a previously password protected file to a cloud storage wouldn't work.

### Office Suite

- Fixes an issue when multiple documents are open in Word/Excel/PowerPoint from the same SharePoint library, only the first document opened are scanned for Policy compliance.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2003: February 21

*Version 2003 (Build 12615.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Office Suite

- **Pick the perfect color:** Use hex color codes to choose exactly the color you want for your font, text highlight, and more.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Excel

- Fixed an issue that users may have experienced when renaming pivot table measures.
- Fixed a performance issue that users may have experienced when using a VBA macro to clear the contents of a range.
- Fixed an issue that caused the UI to flash when users executed a macro that interacted with the ribbon.
- Fixed an issue where CSV files were loaded incorrectly when the first word in the file was TABLE.
- Fixed an issue where users may have experienced closes unexpectedly when switching between two workbooks that had different zoom levels.

### Outlook

- Addressed an issue that caused users to be unable to open public folder messages when Outlook was left running overnight.
- Fixed a race condition where the 'Allow' and 'Deny' buttons on the permissions page are disabled during the authentication workflow of adding a Gmail account.
- Addressed an issue that caused Outlook to unexpectedly generate logging output in some scenarios, even when logging was turned off.

### Word

- Fixed an issue where comment cards don't always get highlighted when a mouse pointer hovers over the comment card.
- During an active document co-authoring session, adding an image directly in to a comment card may result in the addition of a tag. This issue has been fixed.

### Office Suite

- When using Multichoice/Lookup/Managed-metadata properties with Word/Excel/PowerPoint documents and saving to a SharePoint Document Library, these properties were previously limited to 255 characters. When these properties exceeded 255 characters, such documents couldn't be saved. With this change, this limit has been increased to 2048 characters.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2003: February 14

*Version 2003 (Build 12607.20000)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Outlook

- **New experience for captive wifi networks:** Have you ever joined a wifi network that required a web page to sign in with? Outlook now detects this and helps you get connected.

### Word

- **Find Ink Editor in your drawing toolbox:** Select Draw and then choose the Ink Editor pen to edit your document with your finger or a digital pen. [Learn more](https://support.office.com/article/7edbcf8e-0004-484d-9b62-501a31c23ee9)

### Office Suite

- **Clearer status bar icons:** Status bar icons are now easier to see

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Outlook

- Addressed an issue that caused users to lose access to the "Free Busy Options" calendar permissions dialog.

- Fixed an issue that may result in the alert: "Sorry we're having trouble opening this item" when opening some recurring meeting instances sent from a different time zone.

- Addressed an issue that could cause users to be unable to reopen a .msg file after dragging and dropping an attachment from that message.

- Fixed an issue where after uploading a file attachment from Outlook to OneDrive could result in the file name being changed if the attachment's name contained parenthesis.

### PowerPoint

- Fixed an issue that could result in a failure to save a document in PowerPoint or Word containing an Excel chart.

### Word

- Fixed an issue where pictures in documents lose transparency when exported to PDF.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

## Version 2002: February 07

*Version 2002 (Build 12527.20040)*

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT START)

### Feature updates

### Access

- **Be more productive working in Query Designer, SQL view, and the Relationships window:** Right-click a table to open, design, size, and hide it. Search and replace text in SQL View. Select multiple tables in the Relationships window.

[//]: # (DO NOT REMOVE FEATUREDETAILS CONTENT END)

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT START)

### Resolved issues

### Access

- This update fixes an issue where using an ADODB. Recorder object in VB code may incorrectly report an error.

- This update fixes an issue that can cause Microsoft Access to fail to identify an Identity Column in a linked SQL Server table, which can cause rows to be reported as deleted incorrectly.

### Excel

- Fixed an issue where Excel would close unexpectedly when using Text To Columns with dynamic arrays.

### Outlook

- Fixed an issue where scrolling in calendar with month view, fails to show previous calendar events.

- Addresses an issue that caused users to experience a unexpected close when viewing more than 30 calendars in a Citrix environment.

### PowerPoint

- Fixed an issue where after closing a file, PowerPoint doesn't immediately remove it from the Presentations collection if there are any event handlers running. Hence the number of open presentations reported by the object model is incorrect, and shutdown of PowerPoint is prevented.

- Fixed an issue with highlighter : White texts with dark highlighter colors are printed as black in Grayscale.

### Word

- Updating and scrolling through a table of contents may sometimes display a gray area over the document.

- Fixed an issue where if a document is being coauthored, the draft version of a root comment may not be preserved.

- Fixed an issue where going back and forth between comment cards would sometimes display the initially selected comment with a selection highlight.

- Fixed an issue where using 'Browse' to save a file didn't work if a comment was written but not posted and the user tried to save the file.

- With SlideTrack enabled and the comments pane closed, Ctrl+Alt+M may not open the comments pane.

- Fixed an issue when adding @mention in a table could generate the error message: 'A table in this document has become corrupted'.

### Office Suite

- Resolves an issue that may have caused Norway Nynorsk (nn-no) proofing tools package to be installed incorrectly.

[//]: # (DO NOT REMOVE BUGDETAILS CONTENT END)

[//]: # (DO NOT MODIFY ADMIN CENTER METADATA CONTENT START)
[//]: # (|Win32|DevMain|Insiders| |16.0.18526.20024|version-2502-february-03|)
[//]: # (|Win32|DevMain|Insiders| |16.0.18516.20000|version-2502-january-21|)
[//]: # (|Win32|DevMain|Insiders| |16.0.18514.20000|version-2502-january-17|)
[//]: # (|Win32|DevMain|Insiders| |16.0.18502.20000|version-2502-january-07|)
[//]: # (|Win32|DevMain|Insiders| |16.0.18429.20000|version-2501-january-01|)
[//]: # (|Win32|DevMain|Insiders| |16.0.18422.20000|version-2501-december-25|)
[//]: # (|Win32|DevMain|Insiders| |16.0.18407.20002|version-2501-december-10|)
[//]: # (|Win32|DevMain|Insiders| |16.0.18330.20000|version-2501-december-03|)
[//]: # (|Win32|DevMain|Insiders| |16.0.18324.20012|version-2412-november-27|)
[//]: # (|Win32|DevMain|Insiders| |16.0.18318.20000|version-2412-november-26|)
[//]: # (DO NOT MODIFY ADMIN CENTER METADATA CONTENT END)
