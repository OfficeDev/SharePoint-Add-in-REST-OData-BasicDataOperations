// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml;

namespace SharePoint_Add_in_REST_OData_BasicDataOperationsWeb
{
    public partial class Default : System.Web.UI.Page
    {
        SharePointContextToken contextToken;
        string accessToken;
        Uri sharepointUrl;

        //Create a namespace manager for parsing the ATOM XML returned by the queries.
        XmlNamespaceManager xmlnspm = new XmlNamespaceManager(new NameTable());       

        protected void Page_PreInit(object sender, EventArgs e)
        {
            Uri redirectUrl;
            switch (SharePointContextProvider.CheckRedirectionStatus(Context, out redirectUrl))
            {
                case RedirectionStatus.Ok:
                    return;
                case RedirectionStatus.ShouldRedirect:
                    Response.Redirect(redirectUrl.AbsoluteUri, endResponse: true);
                    break;
                case RedirectionStatus.CanNotRedirect:
                    Response.Write("An error occurred while processing your request.");
                    Response.End();
                    break;
            }
        }

        // The Page_load method gets the context token and the access token. The access token is used by all of the data retrieval methods.
        protected void Page_Load(object sender, EventArgs e)
        {
            string contextTokenString = TokenHelper.GetContextTokenFromRequest(Request);

            if (contextTokenString != null)
            {
                contextToken =
                    TokenHelper.ReadAndValidateContextToken(contextTokenString, Request.Url.Authority);

                sharepointUrl = new Uri(Request.QueryString["SPHostUrl"]);
                accessToken =
                    TokenHelper.GetAccessToken(contextToken, sharepointUrl.Authority).AccessToken;

                // In a production add-in, you should cache the access token somewhere, such as in a database
                // or ASP.NET Session Cache. (Do not put it in a cookie.) Your code should also check to see 
                // if it is expired before using it (and use the refresh token to get a new one when needed). 
                // For more information, see the MSDN topic at https://msdn.microsoft.com/library/office/dn762763.aspx
                // For simplicity, this sample does not follow these practices. 
                AddListButton.CommandArgument = accessToken;
                RefreshListButton.CommandArgument = accessToken;
                RetrieveListButton.CommandArgument = accessToken;
                AddItemButton.CommandArgument = accessToken;
                DeleteListButton.CommandArgument = accessToken;
                ChangeListTitleButton.CommandArgument = accessToken;
                RetrieveLists(accessToken);

            }
            else if (!IsPostBack)
            {
                Response.Write("Could not find a context token.");
            }
        }


        private void RetrieveLists(string accessToken)
        {
            if (IsPostBack)
            {
                sharepointUrl = new Uri(Request.QueryString["SPHostUrl"]);
            }

            AddItemButton.Visible = false;
            AddListItemBox.Visible = false;
            DeleteListButton.Visible = false;
            ChangeListTitleButton.Visible = false;
            ChangeListTitleBox.Visible = false;
            RetrieveListNameBox.Enabled = true;
            ListTable.Rows[0].Cells[1].Text = "List ID";

            //Add needed namespaces to the namespace manager.
            xmlnspm.AddNamespace("atom", "http://www.w3.org/2005/Atom");
            xmlnspm.AddNamespace("d", "http://schemas.microsoft.com/ado/2007/08/dataservices");
            xmlnspm.AddNamespace("m", "http://schemas.microsoft.com/ado/2007/08/dataservices/metadata");

            //Execute a REST request for all of the site's lists.
            HttpWebRequest listRequest =
                (HttpWebRequest)HttpWebRequest.Create(sharepointUrl.ToString() + "/_api/Web/lists");
            listRequest.Method = "GET";
            listRequest.Accept = "application/atom+xml";
            listRequest.ContentType = "application/atom+xml;type=entry";
            listRequest.Headers.Add("Authorization", "Bearer " + accessToken);
            HttpWebResponse listResponse = (HttpWebResponse)listRequest.GetResponse();
            StreamReader listReader = new StreamReader(listResponse.GetResponseStream());
            var listXml = new XmlDocument();
            listXml.LoadXml(listReader.ReadToEnd());

            var titleList = listXml.SelectNodes("//atom:entry/atom:content/m:properties/d:Title", xmlnspm);
            var idList = listXml.SelectNodes("//atom:entry/atom:content/m:properties/d:Id", xmlnspm);

            int listCounter = 0;
            foreach (XmlNode title in titleList)
            {
                TableRow tableRow = new TableRow();
                LiteralControl idClick = new LiteralControl();
                //Use Javascript to populate the RetrieveListNameBox control with the list id.
                string clickScript = "<a onclick=\"document.getElementById(\'RetrieveListNameBox\').value = '" + idList[listCounter].InnerXml + "';\" href=\"#\">" + idList[listCounter].InnerXml + "</a>";
                idClick.Text = clickScript;
                TableCell tableCell1 = new TableCell();
                tableCell1.Controls.Add(new LiteralControl(title.InnerXml));
                TableCell tableCell2 = new TableCell();
                tableCell2.Controls.Add(idClick);
                tableRow.Cells.Add(tableCell1);
                tableRow.Cells.Add(tableCell2);
                ListTable.Rows.Add(tableRow);
                listCounter++;
            }
        }

        private void RetrieveListItems(string accessToken, Guid listId)
        {
            if (IsPostBack)
            {
                sharepointUrl = new Uri(Request.QueryString["SPHostUrl"]);
            }

            //Adjust the visibility of controls on the page in light of the list-specific context.
            AddItemButton.Visible = true;
            AddListItemBox.Visible = true;
            DeleteListButton.Visible = true;
            ChangeListTitleButton.Visible = true;
            ChangeListTitleBox.Visible = true;
            RetrieveListNameBox.Enabled = false;
            ListTable.Rows[0].Cells[1].Text = "List Items";

            //Add needed namespaces to the namespace manager.
            xmlnspm.AddNamespace("atom", "http://www.w3.org/2005/Atom");
            xmlnspm.AddNamespace("d", "http://schemas.microsoft.com/ado/2007/08/dataservices");
            xmlnspm.AddNamespace("m", "http://schemas.microsoft.com/ado/2007/08/dataservices/metadata");

            //Execute a REST request to get the list name.
            HttpWebRequest listRequest =
                (HttpWebRequest)HttpWebRequest.Create(sharepointUrl.ToString() + "/_api/Web/lists(guid'" + listId + "')");
            listRequest.Method = "GET";
            listRequest.Accept = "application/atom+xml";
            listRequest.ContentType = "application/atom+xml;type=entry";
            listRequest.Headers.Add("Authorization", "Bearer " + accessToken);
            HttpWebResponse listResponse = (HttpWebResponse)listRequest.GetResponse();
            StreamReader listReader = new StreamReader(listResponse.GetResponseStream());
            var listXml = new XmlDocument();
            listXml.LoadXml(listReader.ReadToEnd());

            var listNameNode = listXml.SelectSingleNode("//atom:entry/atom:content/m:properties/d:Title", xmlnspm);
            string listName = listNameNode.InnerXml;

            //Execute a REST request to get all of the list's items.
            HttpWebRequest itemRequest =
                (HttpWebRequest)HttpWebRequest.Create(sharepointUrl.ToString() + "/_api/Web/lists(guid'" + listId + "')/Items");
            itemRequest.Method = "GET";
            itemRequest.Accept = "application/atom+xml";
            itemRequest.ContentType = "application/atom+xml;type=entry";
            itemRequest.Headers.Add("Authorization", "Bearer " + accessToken);
            HttpWebResponse itemResponse = (HttpWebResponse)itemRequest.GetResponse();
            StreamReader itemReader = new StreamReader(itemResponse.GetResponseStream());
            var itemXml = new XmlDocument();
            itemXml.LoadXml(itemReader.ReadToEnd());

            var itemList = itemXml.SelectNodes("//atom:entry/atom:content/m:properties/d:Title", xmlnspm);

            TableRow tableRow = new TableRow();
            TableCell tableCell1 = new TableCell();
            tableCell1.Controls.Add(new LiteralControl(listName));
            TableCell tableCell2 = new TableCell();

            foreach (XmlNode itemTitle in itemList)
            {
                tableCell2.Text += itemTitle.InnerXml + "<br>";
            }

            tableRow.Cells.Add(tableCell1);
            tableRow.Cells.Add(tableCell2);
            ListTable.Rows.Add(tableRow);
        }

        private void AddList(string accessToken, string newListName)
        {
            if (IsPostBack)
            {
                sharepointUrl = new Uri(Request.QueryString["SPHostUrl"]);
            }

            try
            {


                //Add pertinent namespace to the namespace manager.
                xmlnspm.AddNamespace("d", "http://schemas.microsoft.com/ado/2007/08/dataservices");

                //Execute a REST request to get the form digest. All POST requests that change the state of resources on the host
                //Web require the form digest in the request header.
                HttpWebRequest contextinfoRequest =
                    (HttpWebRequest)HttpWebRequest.Create(sharepointUrl.ToString() + "/_api/contextinfo");
                contextinfoRequest.Method = "POST";
                contextinfoRequest.ContentType = "text/xml;charset=utf-8";
                contextinfoRequest.ContentLength = 0;
                contextinfoRequest.Headers.Add("Authorization", "Bearer " + accessToken);

                HttpWebResponse contextinfoResponse = (HttpWebResponse)contextinfoRequest.GetResponse();
                StreamReader contextinfoReader = new StreamReader(contextinfoResponse.GetResponseStream(), System.Text.Encoding.UTF8);
                var formDigestXML = new XmlDocument();
                formDigestXML.LoadXml(contextinfoReader.ReadToEnd());
                var formDigestNode = formDigestXML.SelectSingleNode("//d:FormDigestValue", xmlnspm);
                string formDigest = formDigestNode.InnerXml;

                //Execute a REST request to add a list that has the user-supplied name.
                //The body of the REST request is ASCII encoded and inserted into the request stream.
                string listPostBody = "{'__metadata':{'type':'SP.List'}, 'Title':'" + newListName + "', 'BaseTemplate': 100}";
                byte[] listPostData = System.Text.Encoding.ASCII.GetBytes(listPostBody);

                HttpWebRequest listRequest =
                    (HttpWebRequest)HttpWebRequest.Create(sharepointUrl.ToString() + "/_api/lists");
                listRequest.Method = "POST";
                listRequest.ContentLength = listPostBody.Length;
                listRequest.ContentType = "application/json;odata=verbose";
                listRequest.Accept = "application/json;odata=verbose";
                listRequest.Headers.Add("Authorization", "Bearer " + accessToken);
                listRequest.Headers.Add("X-RequestDigest", formDigest);
                Stream listRequestStream = listRequest.GetRequestStream();
                listRequestStream.Write(listPostData, 0, listPostData.Length);
                listRequestStream.Close();
                HttpWebResponse listResponse = (HttpWebResponse)listRequest.GetResponse();

                RetrieveLists(accessToken);
            }
            catch (Exception e)
            {
                AddListNameBox.Text = e.Message;
            }
        }

        private void AddListItem(string accessToken, Guid listId, string newItemName)
        {
            if (IsPostBack)
            {
                sharepointUrl = new Uri(Request.QueryString["SPHostUrl"]);
            }

            try
            {

                //Add pertinent namespaces to the namespace manager.
                xmlnspm.AddNamespace("atom", "http://www.w3.org/2005/Atom");
                xmlnspm.AddNamespace("d", "http://schemas.microsoft.com/ado/2007/08/dataservices");
                xmlnspm.AddNamespace("m", "http://schemas.microsoft.com/ado/2007/08/dataservices/metadata");


                //Execute a REST request to get the form digest. All POST requests that change the state of resources on the host
                //Web require the form digest in the request header.
                HttpWebRequest contextinfoRequest =
                    (HttpWebRequest)HttpWebRequest.Create(sharepointUrl.ToString() + "/_api/contextinfo");
                contextinfoRequest.Method = "POST";
                contextinfoRequest.ContentType = "text/xml;charset=utf-8";
                contextinfoRequest.ContentLength = 0;
                contextinfoRequest.Headers.Add("Authorization", "Bearer " + accessToken);

                HttpWebResponse contextinfoResponse = (HttpWebResponse)contextinfoRequest.GetResponse();
                StreamReader contextinfoReader = new StreamReader(contextinfoResponse.GetResponseStream(), System.Text.Encoding.UTF8);
                var formDigestXML = new XmlDocument();
                formDigestXML.LoadXml(contextinfoReader.ReadToEnd());
                var formDigestNode = formDigestXML.SelectSingleNode("//d:FormDigestValue", xmlnspm);
                string formDigest = formDigestNode.InnerXml;

                //Execute a REST request to get the list name and the entity type name for the list.
                HttpWebRequest listRequest =
                    (HttpWebRequest)HttpWebRequest.Create(sharepointUrl.ToString() + "/_api/Web/lists(guid'" + listId + "')");
                listRequest.Method = "GET";
                listRequest.Accept = "application/atom+xml";
                listRequest.ContentType = "application/atom+xml;type=entry";
                listRequest.Headers.Add("Authorization", "Bearer " + accessToken);
                HttpWebResponse listResponse = (HttpWebResponse)listRequest.GetResponse();
                StreamReader listReader = new StreamReader(listResponse.GetResponseStream());
                var listXml = new XmlDocument();
                listXml.LoadXml(listReader.ReadToEnd());

                //The entity type name is the required type when you construct a request to add a list item.
                var entityTypeNode = listXml.SelectSingleNode("//atom:entry/atom:content/m:properties/d:ListItemEntityTypeFullName", xmlnspm);
                var listNameNode = listXml.SelectSingleNode("//atom:entry/atom:content/m:properties/d:Title", xmlnspm);
                string entityTypeName = entityTypeNode.InnerXml;
                string listName = listNameNode.InnerXml;

                //Execute a REST request to add an item to the list.
                string itemPostBody = "{'__metadata':{'type':'" + entityTypeName + "'}, 'Title':'" + newItemName + "'}}";
                Byte[] itemPostData = System.Text.Encoding.ASCII.GetBytes(itemPostBody);

                HttpWebRequest itemRequest =
                    (HttpWebRequest)HttpWebRequest.Create(sharepointUrl.ToString() + "/_api/Web/lists(guid'" + listId + "')/Items");
                itemRequest.Method = "POST";
                itemRequest.ContentLength = itemPostBody.Length;
                itemRequest.ContentType = "application/json;odata=verbose";
                itemRequest.Accept = "application/json;odata=verbose";
                itemRequest.Headers.Add("Authorization", "Bearer " + accessToken);
                itemRequest.Headers.Add("X-RequestDigest", formDigest);
                Stream itemRequestStream = itemRequest.GetRequestStream();

                itemRequestStream.Write(itemPostData, 0, itemPostData.Length);
                itemRequestStream.Close();

                HttpWebResponse itemResponse = (HttpWebResponse)itemRequest.GetResponse();
                RetrieveListItems(accessToken, listId);
            }
            catch (Exception e)
            {
                AddListItemBox.Text = e.Message;
            }
        }

        private void ChangeListTitle(string accessToken, Guid listId, string newListTitle)
        {
            if (IsPostBack)
            {
                sharepointUrl = new Uri(Request.QueryString["SPHostUrl"]);
            }

            //Add pertinent namespace to the namespace manager.
            xmlnspm.AddNamespace("d", "http://schemas.microsoft.com/ado/2007/08/dataservices");

            //Execute a REST request to get the form digest. All POST requests that change the state of resources on the host
            //Web require the form digest in the request header.
            HttpWebRequest contextinfoRequest =
                (HttpWebRequest)HttpWebRequest.Create(sharepointUrl.ToString() + "/_api/contextinfo");
            contextinfoRequest.Method = "POST";
            contextinfoRequest.ContentType = "text/xml;charset=utf-8";
            contextinfoRequest.ContentLength = 0;
            contextinfoRequest.Headers.Add("Authorization", "Bearer " + accessToken);

            HttpWebResponse contextinfoResponse = (HttpWebResponse)contextinfoRequest.GetResponse();
            StreamReader contextinfoReader = new StreamReader(contextinfoResponse.GetResponseStream(), System.Text.Encoding.UTF8);
            var formDigestXML = new XmlDocument();
            formDigestXML.LoadXml(contextinfoReader.ReadToEnd());
            var formDigestNode = formDigestXML.SelectSingleNode("//d:FormDigestValue", xmlnspm);
            string formDigest = formDigestNode.InnerXml;

            //Execute a REST request to get the ETag value, which needs to be sent with the delete request.
            HttpWebRequest getListEtagRequest =
                (HttpWebRequest)HttpWebRequest.Create(sharepointUrl.ToString() + "/_api/Web/lists(guid'" + listId + "')");
            getListEtagRequest.Method = "GET";
            getListEtagRequest.Accept = "application/atom+xml";
            getListEtagRequest.ContentType = "application/atom+xml;type=entry";
            getListEtagRequest.Headers.Add("Authorization", "Bearer " + accessToken);
            HttpWebResponse listETagResponse = (HttpWebResponse)getListEtagRequest.GetResponse();
            string eTag = listETagResponse.Headers["ETag"];

            //Execute a REST request to change the list title
            string listPostBody = "{'__metadata':{'type':'SP.List'}, 'Title':'" + newListTitle + "'}";
            byte[] listPostData = System.Text.Encoding.ASCII.GetBytes(listPostBody);

            HttpWebRequest listRequest =
                (HttpWebRequest)HttpWebRequest.Create(sharepointUrl.ToString() + "/_api/lists(guid'" + listId + "')");
            listRequest.Method = "POST";
            listRequest.ContentLength = listPostBody.Length;
            listRequest.ContentType = "application/json;odata=verbose";
            listRequest.Accept = "application/json;odata=verbose";
            listRequest.Headers.Add("Authorization", "Bearer " + accessToken);
            listRequest.Headers.Add("X-RequestDigest", formDigest);
            listRequest.Headers.Add("If-Match", eTag);
            listRequest.Headers.Add("X-Http-Method", "MERGE");
            Stream listRequestStream = listRequest.GetRequestStream();
            listRequestStream.Write(listPostData, 0, listPostData.Length);
            listRequestStream.Close();
            HttpWebResponse listResponse = (HttpWebResponse)listRequest.GetResponse();

            RetrieveListItems(accessToken, listId);
        }

        private void DeleteList(string accessToken, Guid listId)
        {
            if (IsPostBack)
            {
                sharepointUrl = new Uri(Request.QueryString["SPHostUrl"]);
            }

            //Add pertinent namespace to the namespace manager.
            xmlnspm.AddNamespace("d", "http://schemas.microsoft.com/ado/2007/08/dataservices");

            //Execute a REST request to get the form digest. All POST requests that change the state of resources on the host
            //Web require the form digest in the request header.
            HttpWebRequest contextinfoRequest =
                (HttpWebRequest)HttpWebRequest.Create(sharepointUrl.ToString() + "/_api/contextinfo");
            contextinfoRequest.Method = "POST";
            contextinfoRequest.ContentType = "text/xml;charset=utf-8";
            contextinfoRequest.ContentLength = 0;
            contextinfoRequest.Headers.Add("Authorization", "Bearer " + accessToken);

            HttpWebResponse contextinfoResponse = (HttpWebResponse)contextinfoRequest.GetResponse();
            StreamReader contextinfoReader = new StreamReader(contextinfoResponse.GetResponseStream(), System.Text.Encoding.UTF8);
            var formDigestXML = new XmlDocument();
            formDigestXML.LoadXml(contextinfoReader.ReadToEnd());
            var formDigestNode = formDigestXML.SelectSingleNode("//d:FormDigestValue", xmlnspm);
            string formDigest = formDigestNode.InnerXml;

            //Execute a REST request to get the ETag value, which needs to be sent with the delete request.
            HttpWebRequest getListEtagRequest =
                (HttpWebRequest)HttpWebRequest.Create(sharepointUrl.ToString() + "/_api/Web/lists(guid'" + listId + "')");
            getListEtagRequest.Method = "GET";
            getListEtagRequest.Accept = "application/atom+xml";
            getListEtagRequest.ContentType = "application/atom+xml;type=entry";
            getListEtagRequest.Headers.Add("Authorization", "Bearer " + accessToken);
            HttpWebResponse listETagResponse = (HttpWebResponse)getListEtagRequest.GetResponse();
            string eTag = listETagResponse.Headers["ETag"];

            //Execute a REST request to delete the list.
            HttpWebRequest deleteListRequest =
                    (HttpWebRequest)HttpWebRequest.Create(sharepointUrl.ToString() + "/_api/Web/lists(guid'" + listId + "')");
            deleteListRequest.Method = "POST";
            deleteListRequest.ContentLength = 0;
            deleteListRequest.ContentType = "text/xml;charset=utf-8";
            deleteListRequest.Headers.Add("X-RequestDigest", formDigest);
            deleteListRequest.Headers.Add("If-Match", eTag);
            deleteListRequest.Headers.Add("Authorization", "Bearer " + accessToken);
            deleteListRequest.Headers.Add("X-Http-Method", "DELETE");
            HttpWebResponse deleteListResponse = (HttpWebResponse)deleteListRequest.GetResponse();
            RetrieveListNameBox.Text = "";
            RetrieveLists(accessToken);
        }

        protected void AddList_Click(object sender, EventArgs e)
        {
            string commandAccessToken = ((Button)sender).CommandArgument;
            if (AddListNameBox.Text != "")
            {
                AddList(commandAccessToken, AddListNameBox.Text);
            }
            else
            {
                AddListNameBox.Text = "Enter a list title";
            }
        }

        protected void RefreshList_Click(object sender, EventArgs e)
        {
            string commandAccessToken = ((Button)sender).CommandArgument;
            RetrieveLists(commandAccessToken);
        }

        protected void RetrieveListButton_Click(object sender, EventArgs e)
        {
            string commandAccessToken = ((Button)sender).CommandArgument;
            Guid listId = new Guid();
            if (Guid.TryParse(RetrieveListNameBox.Text, out listId))
            {
                RetrieveListItems(commandAccessToken, listId);
            }
            else
            {
                RetrieveListNameBox.Text = "Enter a List GUID";
            }
        }

        protected void AddItemButton_Click(object sender, EventArgs e)
        {
            string commandAccessToken = ((Button)sender).CommandArgument;
            Guid listId = new Guid(RetrieveListNameBox.Text);
            if (AddListItemBox.Text != "")
            {
                AddListItem(commandAccessToken, listId, AddListItemBox.Text);
            }
            else
            {
                AddListItemBox.Text = "Enter an item title";
            }
        }

        protected void DeleteListButton_Click(object sender, EventArgs e)
        {
            string commandAccessToken = ((Button)sender).CommandArgument;
            Guid listId = new Guid(RetrieveListNameBox.Text);
            DeleteList(commandAccessToken, listId);
        }

        protected void ChangeListTitleButton_Click(object sender, EventArgs e)
        {
            string commandAccessToken = ((Button)sender).CommandArgument;
            Guid listId = new Guid(RetrieveListNameBox.Text);
            if (ChangeListTitleBox.Text != "")
            {
                ChangeListTitle(commandAccessToken, listId, ChangeListTitleBox.Text);
            }
            else
            {
                ChangeListTitleBox.Text = "Enter a new list title";
            }
        }
    }
}

/*
SharePoint Add-in REST/OData Basic Data Operations, https://github.com/OfficeDev/SharePoint-Add-in-REST-OData-BasicDataOperations
 
Copyright (c) Microsoft Corporation
All rights reserved. 
 
MIT License:
Permission is hereby granted, free of charge, to any person obtaining
a copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:
 
The above copyright notice and this permission notice shall be
included in all copies or substantial portions of the Software.
 
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.    
  
*/