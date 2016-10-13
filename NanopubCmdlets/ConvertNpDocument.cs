// File: ConvertNpDocument.cs

using System;
using IO=System.IO;
using System.Management.Automation;

using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;

namespace NanopubCmdlets
{
    [Cmdlet("Convert", "NpDocument")]
    public class ConvertNpDocument : PSCmdlet
    {
        private string path;

        [Alias("Fullname")]
        [Parameter(HelpMessage = "Der Pfad der Docx-Datei",
            Mandatory = true,
            ValueFromPipelineByPropertyName = true)]
        public string Path
        {
            get { return path; }
            set
            {
                path = value;
            }
        }

        /// <summary>
        /// Wird für jedes Objekt in der Pipeline ausgeführt
        /// </summary>
        protected override void ProcessRecord()
        {
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
                WordprocessingDocument wdDoc = WordprocessingDocument.Open(path, false);
                int i = 0;
                foreach (Paragraph para in wdDoc.MainDocumentPart.Document.Body.Descendants<Paragraph>())
                {
                    i++;
                    WriteObject(para.InnerText);
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
