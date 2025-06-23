<%@ WebHandler Language="C#" Class="EnquirySubmissionHandler" %>

using System;
using System.Web;
using System.Web.SessionState;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;
using System.Collections.Specialized;
using System.IO;
using System.Text;
using System.Web.Script.Serialization;

public class EnquirySubmissionHandler : IHttpHandler, IRequiresSessionState
{
    public void ProcessRequest(HttpContext context)
    {
        try
        {
            context.Response.ContentType = "application/json";
            context.Response.Headers.Add("Access-Control-Allow-Origin", "*");
            context.Response.Headers.Add("Access-Control-Allow-Methods", "POST, OPTIONS");
            context.Response.Headers.Add("Access-Control-Allow-Headers", "Content-Type");
            
            // Handle preflight OPTIONS request
            if (context.Request.HttpMethod == "OPTIONS")
            {
                context.Response.StatusCode = 200;
                context.Response.End();
                return;
            }
            
            if (context.Request.HttpMethod != "POST")
            {
                context.Response.StatusCode = 405;
                WriteJsonResponse(context, false, "Method not allowed");
                return;
            }

            // Check if this is a file-only upload request (HYBRID MODE)
            string fileUploadOnly = context.Request.Form["fileUploadOnly"];
            bool isFileOnlyMode = !string.IsNullOrEmpty(fileUploadOnly) && fileUploadOnly.ToLower() == "true";

            if (isFileOnlyMode)
            {
                // Handle file-only upload mode (for hybrid approach)
                HandleFileOnlyUpload(context);
                return;
            }

            // Original full form submission mode (legacy)
            HandleFullFormSubmission(context);
        }
        catch (Exception ex)
        {
            context.Response.StatusCode = 500;
            WriteJsonResponse(context, false, "Error: " + ex.Message);
        }
    }
    
    private void HandleFileOnlyUpload(HttpContext context)
    {
        try
        {
            // Get enquiry reference
            string enquiryRef = context.Request.Form["enquiryReference"];
            if (string.IsNullOrEmpty(enquiryRef))
            {
                context.Response.StatusCode = 400;
                WriteJsonResponse(context, false, "Enquiry reference is required for file upload");
                return;
            }

            // Check if files were uploaded
            if (context.Request.Files.Count == 0)
            {
                WriteJsonResponse(context, true, "No files to upload", enquiryRef);
                return;
            }

            // Run with elevated privileges for file uploads only
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                using (SPSite site = new SPSite("https://www.ifwg.co.za"))
                using (SPWeb web = site.OpenWeb())
                {
                    web.AllowUnsafeUpdates = true;
                    
                    try
                    {
                        // Get the document library
                        SPDocumentLibrary docLibrary = null;
                        
                        try
                        {
                            docLibrary = (SPDocumentLibrary)web.Lists["EnquiryFormDocuments"];
                        }
                        catch (Exception ex)
                        {
                            throw new Exception("Document library 'EnquiryFormDocuments' not found: " + ex.Message);
                        }
                        
                        // Handle file uploads
                        System.Collections.Generic.List<string> uploadedFiles = new System.Collections.Generic.List<string>();
                        int numberOfFiles = 0;
                        
                        foreach (string fileKey in context.Request.Files.AllKeys)
                        {
                            HttpPostedFile file = context.Request.Files[fileKey];
                            if (file != null && file.ContentLength > 0)
                            {
                                try
                                {
                                    // Validate file size (10MB limit)
                                    if (file.ContentLength > 10 * 1024 * 1024)
                                    {
                                        throw new Exception(string.Format("File {0} exceeds 10MB size limit", file.FileName));
                                    }
                                    
                                    // Create unique filename to avoid conflicts
                                    string fileName = string.Format("{0}_{1}_{2}", 
                                        enquiryRef, 
                                        DateTime.Now.ToString("yyyyMMdd_HHmmss"), 
                                        Path.GetFileName(file.FileName));
                                    
                                    // Read file content
                                    byte[] fileBytes = new byte[file.ContentLength];
                                    file.InputStream.Read(fileBytes, 0, file.ContentLength);
                                    
                                    // Upload to SharePoint document library
                                    SPFile uploadedFile = docLibrary.RootFolder.Files.Add(
                                        fileName, 
                                        fileBytes, 
                                        true
                                    );
                                    
                                    // Set file properties
                                    if (uploadedFile.Item != null)
                                    {
                                        uploadedFile.Item["Title"] = enquiryRef + " - " + Path.GetFileNameWithoutExtension(file.FileName);
                                        uploadedFile.Item.Update();
                                    }
                                    
                                    numberOfFiles++;
                                    uploadedFiles.Add(uploadedFile.ServerRelativeUrl);
                                    
                                    System.Diagnostics.Debug.WriteLine(string.Format("File uploaded successfully: {0}", fileName));
                                }
                                catch (Exception fileEx)
                                {
                                    System.Diagnostics.Debug.WriteLine(string.Format("File upload error for {0}: {1}", file.FileName, fileEx.Message));
                                    throw new Exception(string.Format("Failed to upload file {0}: {1}", file.FileName, fileEx.Message));
                                }
                            }
                        }
                        
                        web.AllowUnsafeUpdates = false;
                        
                        // Return success response with uploaded file details
                        var response = new
                        {
                            success = true,
                            message = string.Format("Successfully uploaded {0} file(s)", numberOfFiles),
                            enquiryReference = enquiryRef,
                            uploadedFiles = uploadedFiles.ToArray(),
                            fileCount = numberOfFiles,
                            timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                        };
                        
                        JavaScriptSerializer serializer = new JavaScriptSerializer();
                        string json = serializer.Serialize(response);
                        context.Response.Write(json);
                    }
                    catch (Exception innerEx)
                    {
                        web.AllowUnsafeUpdates = false;
                        throw innerEx;
                    }
                }
            });
        }
        catch (Exception ex)
        {
            context.Response.StatusCode = 500;
            WriteJsonResponse(context, false, "File upload error: " + ex.Message);
        }
    }

    private void HandleFullFormSubmission(HttpContext context)
    {
        // Original full form submission logic (legacy mode)
        SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                using (SPSite site = new SPSite("https://www.ifwg.co.za"))
                using (SPWeb web = site.OpenWeb())
                {
                    web.AllowUnsafeUpdates = true;
                    
                    try
                    {
                        // Get the list and library
                        SPList enquiryList = web.Lists["Enquiry Details"];
                        SPDocumentLibrary docLibrary = null;
                        
                        try
                        {
                            docLibrary = (SPDocumentLibrary)web.Lists["EnquiryFormDocuments"];
                        }
                        catch (Exception ex)
                        {
                            // Document library might not exist, that's okay
                            System.Diagnostics.Debug.WriteLine("Document library not found: " + ex.Message);
                        }
                        
                        // Create new list item
                        SPListItem newItem = enquiryList.Items.Add();
                        
                        // Get enquiry reference
                        string enquiryRef = GetFormValue(context, "enquiryReference");
                        if (string.IsNullOrEmpty(enquiryRef))
                        {
                            enquiryRef = GenerateEnquiryReference();
                        }
                        
                        // Map form fields to list columns (adjust field names to match your list)
                        SetListItemField(newItem, "Title", enquiryRef);
                        SetListItemField(newItem, "FullName", GetFormValue(context, "fullName"));
                        SetListItemField(newItem, "OrganisationName", GetFormValue(context, "organisationName"));
                        SetListItemField(newItem, "ContactNumber", GetFormValue(context, "contactNumber"));
                        SetListItemField(newItem, "EmailAddress", GetFormValue(context, "emailAddress"));
                        SetListItemField(newItem, "WebsiteAddress", GetFormValue(context, "websiteAddress"));
                        SetListItemField(newItem, "OperationLocation", GetFormValue(context, "operationLocation"));
                        SetListItemField(newItem, "CountriesOfOperation", GetFormValue(context, "countriesOfOperation"));
                        SetListItemField(newItem, "OperationLength", GetFormValue(context, "operationLength"));
                        SetListItemField(newItem, "PrimaryBusinessAreas", GetFormValue(context, "primaryBusinessAreas"));
                        SetListItemField(newItem, "ProductServiceCategory", GetFormValue(context, "productServiceCategory"));
                        SetListItemField(newItem, "OtherProductServiceCategory", GetFormValue(context, "otherProductServiceCategory"));
                        SetListItemField(newItem, "OperationalStatus", GetFormValue(context, "operationalStatus"));
                        SetListItemField(newItem, "RegulatoryStatus", GetFormValue(context, "regulatoryStatus"));
                        SetListItemField(newItem, "Regulators", GetFormValue(context, "regulators"));
                        SetListItemField(newItem, "OtherRegulator", GetFormValue(context, "otherRegulator"));
                        SetListItemField(newItem, "ProductServiceDescription", GetFormValue(context, "productServiceDescription"));
                        SetListItemField(newItem, "Questions", GetFormValue(context, "questions"));
                        SetListItemField(newItem, "AdditionalInformation", GetFormValue(context, "additionalInformation"));
                        SetListItemField(newItem, "FAQConfirmation", GetFormValue(context, "faqConfirmation"));
                        SetListItemField(newItem, "ConsentConfirmation", GetFormValue(context, "consentConfirmation"));
                        
                        // Handle file uploads
                        string fileUrls = "";
                        int numberOfFiles = 0;
                        
                        if (docLibrary != null && context.Request.Files.Count > 0)
                        {
                            StringBuilder fileUrlBuilder = new StringBuilder();
                            
                            foreach (string fileKey in context.Request.Files.AllKeys)
                            {
                                HttpPostedFile file = context.Request.Files[fileKey];
                                if (file != null && file.ContentLength > 0)
                                {
                                    try
                                    {
                                        // Create unique filename to avoid conflicts
                                        string fileName = string.Format("{0}_{1}_{2}", 
                                            enquiryRef, 
                                            DateTime.Now.ToString("yyyyMMdd_HHmmss"), 
                                            file.FileName);
                                        
                                        byte[] fileBytes = new byte[file.ContentLength];
                                        file.InputStream.Read(fileBytes, 0, file.ContentLength);
                                        
                                        SPFile uploadedFile = docLibrary.RootFolder.Files.Add(
                                            fileName, 
                                            fileBytes, 
                                            true
                                        );
                                        
                                        numberOfFiles++;
                                        if (fileUrlBuilder.Length > 0)
                                            fileUrlBuilder.Append("; ");
                                        fileUrlBuilder.Append(uploadedFile.ServerRelativeUrl);
                                    }
                                    catch (Exception fileEx)
                                    {
                                        // Log file upload error but don't fail the entire submission
                                        System.Diagnostics.Debug.WriteLine("File upload error: " + fileEx.Message);
                                    }
                                }
                            }
                            
                            fileUrls = fileUrlBuilder.ToString();
                        }
                        
                        // Set file-related fields
                        SetListItemField(newItem, "FileAttachments", numberOfFiles > 0 ? "Yes" : "No");
                        SetListItemField(newItem, "NumberOfAttachments", numberOfFiles.ToString());
                        if (!string.IsNullOrEmpty(fileUrls))
                        {
                            SetListItemField(newItem, "AttachedDocuments", fileUrls);
                        }
                        
                        // Update the list item
                        newItem.Update();
                        web.AllowUnsafeUpdates = false;
                        
                        WriteJsonResponse(context, true, "Enquiry submitted successfully", enquiryRef);
                    }
                    catch (Exception innerEx)
                    {
                        web.AllowUnsafeUpdates = false;
                        throw innerEx;
                    }
                }
            });
        }
    }
    
    private void SetListItemField(SPListItem item, string fieldName, string value)
    {
        try
        {
            if (item.Fields.ContainsField(fieldName))
            {
                item[fieldName] = value ?? "";
            }
        }
        catch (Exception ex)
        {
            // Log but don't fail for individual field errors
            System.Diagnostics.Debug.WriteLine($"Error setting field {fieldName}: {ex.Message}");
        }
    }
    
    private string GetFormValue(HttpContext context, string key)
    {
        return context.Request.Form[key] ?? "";
    }
    
    private string GenerateEnquiryReference()
    {
        DateTime now = DateTime.Now;
        return string.Format("ENQ-{0:yyyyMMdd}-{0:HHmmss}", now);
    }
    
    private void WriteJsonResponse(HttpContext context, bool success, string message, string enquiryRef = null)
    {
        var response = new
        {
            success = success,
            message = message,
            enquiryReference = enquiryRef,
            timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        };
        
        JavaScriptSerializer serializer = new JavaScriptSerializer();
        string json = serializer.Serialize(response);
        context.Response.Write(json);
    }

    public bool IsReusable
    {
        get { return false; }
    }
} 