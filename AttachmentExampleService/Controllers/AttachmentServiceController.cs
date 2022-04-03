/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

using AttachmentsService.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web.Http;
using System.Xml;
using System.Xml.Linq;

namespace AttachmentsService.Controllers
{
    public class AttachmentServiceController : ApiController
    {
        public ServiceResponse PostAttachments(ServiceRequest request)
        {
            ServiceResponse response = new ServiceResponse();

            try
            {
                response = GetAttachmentsFromExchangeServer(request);
            }
            catch (Exception ex)
            {
                response.isError = true;
                response.message = ex.Message;
            }

            return response;
        }

        // This method does the work of making an Exchange Web Services (EWS) request to get the 
        // attachments from the Exchange server. This implementation makes an individual
        // request for each attachment, and returns the count of attachments processed.
        private ServiceResponse GetAttachmentsFromExchangeServer(ServiceRequest request)
        {
            int processedCount = 0;
            List<string> attachmentNames = new List<string>();

            foreach (AttachmentDetails attachment in request.attachments)
            {
                // Prepare a web request object.
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                HttpWebRequest webRequest = WebRequest.CreateHttp(request.ewsUrl);
                webRequest.Headers.Add("Authorization", string.Format("Bearer {0}", request.attachmentToken));
                webRequest.PreAuthenticate = true;
                webRequest.AllowAutoRedirect = false;
                webRequest.Method = "POST";
                webRequest.ContentType = "text/xml; charset=utf-8";

                // Construct the SOAP message for the GetAttchment operation.
                byte[] bodyBytes = Encoding.UTF8.GetBytes(string.Format(GetAttachmentSoapRequest, attachment.id));
                webRequest.ContentLength = bodyBytes.Length;
                webRequest.KeepAlive = false;

                webRequest.ProtocolVersion = HttpVersion.Version10;

                webRequest.ServicePoint.ConnectionLimit = 1;
                Stream requestStream = webRequest.GetRequestStream();
                requestStream.Write(bodyBytes, 0, bodyBytes.Length);
                requestStream.Close();

                // Make the request to the Exchange server and get the response.
                HttpWebResponse webResponse = (HttpWebResponse)webRequest.GetResponse();

                // If the response is okay, create an XML document from the
                // response and process the request.
                if (webResponse.StatusCode == HttpStatusCode.OK)
                {
                    Stream responseStream = webResponse.GetResponseStream();

                    var responseEnvelope = XElement.Load(responseStream);

                    // This method simply writes the XML document to the
                    // trace output. Your service would perform its
                    // processing here.

                    if (responseEnvelope != null)
                    {
                        var processResult = ProcessXmlResponse(responseEnvelope);
                        attachmentNames.Add(string.Format("{0} {1}", attachment.name, processResult));

                    }
                    // Close the response stream.
                    responseStream.Close();
                    webResponse.Close();

                    processedCount++;
                    //attachmentNames.Add(attachment.name);
                }

            }
            ServiceResponse response = new ServiceResponse();
            response.attachmentNames = attachmentNames.ToArray();
            response.attachmentsProcessed = processedCount;

            return response;
        }
        // This method processes the response from the Exchange server.
        // In your application the bulk of the processing occurs here.
        private string ProcessXmlResponse(XElement responseEnvelope)
        {
            // First, check the response for web service errors.
            var errorCodes = from errorCode in responseEnvelope.Descendants
                              ("{http://schemas.microsoft.com/exchange/services/2006/messages}ResponseCode")
                             select errorCode;
            // Return the first error code found.
            foreach (var errorCode in errorCodes)
            {
                if (errorCode.Value != "NoError")
                {
                    return string.Format("Could not process result. Error: {0}", errorCode.Value);
                }
            }

            // No errors found, proceed with processing the content.
            // First, get and process file attachments.
            var fileAttachments = from fileAttachment in responseEnvelope.Descendants
                              ("{http://schemas.microsoft.com/exchange/services/2006/types}FileAttachment")
                                  select fileAttachment;
            foreach (var fileAttachment in fileAttachments)
            {
                var fileContent = fileAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Content");
                var fileData = System.Convert.FromBase64String(fileContent.Value);
                var s = new MemoryStream(fileData);
                // Process the file attachment here.
            }

            // Second, get and process item attachments.
            var itemAttachments = from itemAttachment in responseEnvelope.Descendants
                                  ("{http://schemas.microsoft.com/exchange/services/2006/types}ItemAttachment")
                                  select itemAttachment;
            foreach (var itemAttachment in itemAttachments)
            {
                var message = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Message");
                if (message != null)
                {
                    // Process a message here.
                    break;
                }
                var calendarItem = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}CalendarItem");
                if (calendarItem != null)
                {
                    // Process calendar item here.
                    break;
                }
                var contact = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Contact");
                if (contact != null)
                {
                    // Process contact here.
                    break;
                }
                var task = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Tontact");
                if (task != null)
                {
                    // Process task here.
                    break;
                }
                var meetingMessage = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingMessage");
                if (meetingMessage != null)
                {
                    // Process meeting message here.
                    break;
                }
                var meetingRequest = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingRequest");
                if (meetingRequest != null)
                {
                    // Process meeting request here.
                    break;
                }
                var meetingResponse = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingResponse");
                if (meetingResponse != null)
                {
                    // Process meeting response here.
                    break;
                }
                var meetingCancellation = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingCancellation");
                if (meetingCancellation != null)
                {
                    // Process meeting cancellation here.
                    break;
                }
            }

            return string.Empty;
        }
        private const string GetAttachmentSoapRequest =
    @"<?xml version=""1.0"" encoding=""utf-8""?>
<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""
xmlns:xsd=""http://www.w3.org/2001/XMLSchema""
xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/""
xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
<soap:Header>
<t:RequestServerVersion Version=""Exchange2013"" />
</soap:Header>
  <soap:Body>
    <GetAttachment xmlns=""http://schemas.microsoft.com/exchange/services/2006/messages""
    xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
      <AttachmentShape/>
      <AttachmentIds>
        <t:AttachmentId Id=""{0}""/>
      </AttachmentIds>
    </GetAttachment>
  </soap:Body>
</soap:Envelope>";
    }
}

// *********************************************************
//
// Outlook-Add-in-Javascript-GetAttachments, https://github.com/OfficeDev/Outlook-Add-in-Javascript-GetAttachments
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
// *********************************************************