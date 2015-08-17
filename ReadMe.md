# Basic CRUD operations in SharePoint Add-ins using the REST/OData APIs #

## Summary
Use the SharePoint REST/OData endpoints to perform create, read, update, and delete operations on lists and list items from a SharePoint Add-in.

### Applies to ###
-  SharePoint Online and on-premise SharePoint 2013 and later 

----------
## Prerequisites ##
This sample requires the following:


- A SharePoint 2013 (or later) development environment that is configured for add-in isolation and OAuth. (A SharePoint Online Developer Site is automatically configured. For an on premise development environment, see [Set up an on-premises development environment for SharePoint Add-ins](https://msdn.microsoft.com/library/office/fp179923.aspx) and [Use an Office 365 SharePoint site to authorize provider-hosted add-ins on an on-premises SharePoint site](https://msdn.microsoft.com/library/office/dn155905.aspx).) 


- Visual Studio and the Office Developer Tools for Visual Studio installed on your developer computer 


- Basic familiarity with RESTful web services and OData

## Description of the code ##
The code that uses the REST APIs is located in the Default.aspx.cs file of the SharePoint-Add-in-REST-OData-BasicDataOperationsWeb project. The Default.aspx page of the add-in appears after you install and launch the add-in and looks similar to the following.

![The add-in start page with a table listing all the list on the site by name and ID.](/description/fig1.gif) 



The sample demonstrates the following:


- How to read and write data to and from a SharePoint host web. This data conforms with the OData protocol to the REST endpoints where the list and list item entities are exposed. 



- How to parse Atom-formatted XML returned from these endpoints and how to construct JSON-formatted representations of the list and list item entities so that you can perform Create and Update operations on them. 


- Best practices for retrieving form digest and eTag values that are required for Create and Update operations on lists and list items. 


## To use the sample #

12. Open **Visual Studio** as an administrator.
13. Open the .sln file.
13. In **Solution Explorer**, highlight the SharePoint add-in project and replace the **Site URL** property with the URL of your SharePoint developer site.
14. Press F5.
15. After the add-in installs, the consent page opens. Click **Trust It**.
16. Enter a string in the text box beside the **Add List** button and click the button. In a moment, the page refreshes and the new list is in the table.
17. Click the ID of the list, and then click **Retrieve List Items**. There will initially be no items on the list. Some additional buttons will appear.
18. Add a string to the text box beside the **Add Item** button and press the button. The new item will appear in the table in the row for the list.
19. Add a string to the text box beside the **Change List Title** button and press the button. The title will change in the table.
20. Press the **Delete the List** button and the list is deleted.

**Do not delete any of the built-in SharePoint lists. If you mistakenly do so, recover the list from the SharePoint Recycle Bin.**

## Troubleshooting

<table border="0" cellspacing="5" cellpadding="5" frame="void" align="left" style="width:601px; height:212px">
<tbody>
<tr style="background-color:#a9a9a9">
<th align="left" scope="col"><strong><span style="font-size:small">Problem </span>
</strong></th>
<th align="left" scope="col"><strong><span style="font-size:small">Solution</span></strong></th>
</tr>
<tr valign="top">
<td><span style="font-size:small">Visual Studio does not open the browser after you press the F5 key.</span></td>
<td><span style="font-size:small">Set the SharePoint Add-in project as the startup project.</span></td>
</tr>
<tr valign="top">
<td><span style="font-size:small">HTTP error 405 <strong>Method not allowed</strong>.</span></td>
<td><span style="font-size:small">Locate the applicationhost.config file in <em>%userprofile%</em>\Documents\IISExpress\config.</span>
<p><span style="font-size:small">Locate the handler entry for <strong>StaticFile</strong>, and add the verbs
<strong>GET</strong>, <strong>HEAD</strong>, <strong>POST</strong>, <strong>DEBUG</strong>, and
<strong>TRACE</strong>.</span></p>
</td>
</tr>
</tbody>
</table>

## Questions and comments

We'd love to get your feedback on this sample. You can send your questions and suggestions to us in the [Issues](https://github.com/OfficeDev/SharePoint-Add-in-REST-OData-BasicDataOperations/issues) section of this repository.
  
<a name="resources"/>
## Additional resources

[Get to know the SharePoint 2013 REST service](https://msdn.microsoft.com/library/fp142380.aspx).

[Open Data Protocol](http://www.odata.org/)
 
[OData: JavaScript Object Notation (JSON) Format](http://www.odata.org/developers/protocols/json-format)

[OData: AtomPub Format](http://www.odata.org/developers/protocols/atom-format).

### Copyright ###

Copyright (c) Microsoft. All rights reserved.




