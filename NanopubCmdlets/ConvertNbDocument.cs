// File: ConvertNpDocument.cs

using System;
using IO=System.IO;
using System.Xml.Linq;
using System.Management.Automation;

using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;

namespace NanopubCmdlets
{
    [Cmdlet("Convert", "NbDocument", DefaultParameterSetName="Standard")]
    public class ConvertNbDocument : PSCmdlet
    {
        private string path;
        private string xmlPath;
        private wdNanobook nanobook;

        [Alias("Fullname")]
        [Parameter(HelpMessage = "Der Pfad der Docx-Datei",
            Mandatory = true,
            ValueFromPipelineByPropertyName = true)
        ]

       
        // Pfad der Docx-Datei    
        public string Path
        {
            get { return path; }
            set
            {
                path = value;
            }
        }

        // Pfad der Xml-Ausgabedatei    
        [Parameter(HelpMessage = "Der Pfad der Xml-Ausgabedatei",
            ParameterSetName = "Xml")
        ]
        public string XmlPath
        {
            get { return xmlPath; }
            set
            {
                xmlPath = value;
            }
        }

        /// <summary>
        /// Hier passiert noch nichts
        /// </summary>
        protected override void BeginProcessing()
        {
            base.BeginProcessing();
        }

        /// <summary>
        /// Hier passiert noch nichts
        /// </summary>
        protected override void EndProcessing()
        {
            base.EndProcessing();
        }
        /// <summary>
        /// Wird für jedes Objekt in der Pipeline ausgeführt
        /// </summary>
        protected override void ProcessRecord()
        {
            string tagText = "";
            string tag = "";
            int i = 0;
            int subTopicCount = 0;

            try
            {
                if (!IO.Path.IsPathRooted(path))
                {
                    try
                    {
                        path = GetUnresolvedProviderPathFromPSPath(path);
                    }
                    catch (SystemException ex)
                    {
                        ParameterBindingException parEx = new ParameterBindingException("Fehler beim Binden des Path-Parameters", ex);
                        ErrorRecord errRecord = new ErrorRecord(parEx, "GetWordText.Path", ErrorCategory.InvalidOperation, null);
                        WriteError(errRecord);
                    }
                }
                // Neues Nanobook anlegen
                nanobook = new wdNanobook();

                WordprocessingDocument wdDoc = WordprocessingDocument.Open(path, false);
                foreach (Paragraph para in wdDoc.MainDocumentPart.Document.Body.Descendants<Paragraph>())
                {
                    i++;
                    // Ab hier beginnt das Text-Parsing
                    // Erster Absatz ist <nanobook>
                    tagText = para.InnerText.TrimStart();
                    if (tagText.Length > 1 && tagText.Substring(0,1) == "<" && tagText.Substring(1, 1) != "/")
                    {
                        tag = tagText.Substring(1, tagText.IndexOfAny(new char[] { ' ', '>' }) - 1).TrimEnd().ToLower();
                        switch (tag)
                        {
                            case "nanobook":
                                {
                                    string headerXml = para.InnerText.Replace('>', '/') + ">";
                                    // Alle Attribute auswerten
                                    XElement xHeader = XElement.Parse(headerXml);
                                    nanobook.Id = xHeader.Attribute("id").Value;
                                    nanobook.CreationDate = xHeader.Attribute("creationDate").Value;
                                    nanobook.LastUpdateDate = xHeader.Attribute("lastUpdateDate").Value;
                                    nanobook.Version = xHeader.Attribute("version").Value;
                                    nanobook.Author = xHeader.Attribute("author").Value;
                                    nanobook.DocumentXml = para.InnerText;
                                    break;
                                }
                            case "topic":
                                nanobook.DocumentXml += "<topic>";

                                break;
                            case "subtopic":
                                nanobook.DocumentXml += "<subtopic>";
                                subTopicCount++;
                              
                                break;
                            case "para":
                                nanobook.DocumentXml += "<para>";

                                break;
                            case "example":
                                nanobook.DocumentXml += "<example>";

                                break;
                            default:
                                {
 
                                    break;
                                }
                        }
                    }
                    else if (tagText.Length > 1 && tagText.Substring(1, 1) == "/")
                    {
                        nanobook.DocumentXml += tagText;
                        // Ist das Ende erreicht?
                        if (tagText == "</nanobook>")
                        {
                            break;
                        }
                    }
                    else
                    {
                        nanobook.DocumentXml += tagText;
                        nanobook.DocumentXml += "</" + tag + ">";
                    }
                }
                WriteObject(nanobook);
                wdDoc.Close();
                // Als Datei speichern?
                if (xmlPath != "")
                {
                    byte[] xmlBytes = System.Text.Encoding.Default.GetBytes(nanobook.DocumentXml);
                    try
                    {
                        using (IO.FileStream fs = new IO.FileStream(xmlPath, IO.FileMode.CreateNew))
                        {
                            fs.Write(xmlBytes, 0, xmlBytes.Length);
                        }
                    }
                    catch (SystemException ex)
                    {
                        ErrorRecord errRecord = new ErrorRecord(ex, "ConvertNpDocument.XmlSpeichern", ErrorCategory.InvalidOperation, null);
                        WriteError(errRecord);
                    }
                }
            }
            catch (SystemException ex)
            {
                ErrorRecord errRecord = new ErrorRecord(ex, "ConvertNpDocument.Allgemein", ErrorCategory.InvalidOperation, null);
                WriteError(errRecord);
            }

        }
    }
}
