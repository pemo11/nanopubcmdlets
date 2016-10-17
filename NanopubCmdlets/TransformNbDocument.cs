// File: TransformNbDocument.cs

using System;
using IO = System.IO;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Xsl;
using System.Reflection;

using System.Management.Automation;

namespace NanopubCmdlets
{
    [Cmdlet("Transform", "NbDocument")]
    public class TransformNbDocument : PSCmdlet
    {
        private string path;

        [Alias("Fullname")]
        [Parameter(HelpMessage = "Der Pfad der Xml-Datei",
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
        protected override void ProcessRecord()
        {
            string xmlText = "";
            string htmlText = "";
            try
            {
                using (IO.StreamReader sr = new IO.StreamReader(path))
                {
                    xmlText = sr.ReadToEnd();
                }
            }
            catch (SystemException ex)
            {
                ErrorRecord err = new ErrorRecord(ex, "TransformNbDocument.ReadXml", ErrorCategory.InvalidOperation, null);
                WriteError(err);
            }
            try
            {
                XslCompiledTransform xslTransform = new XslCompiledTransform();
                IO.Stream stXsl = Assembly.GetExecutingAssembly().GetManifestResourceStream("NanopubCmdlets.Nanobook1.xslt");
                IO.StreamReader srXsl = new IO.StreamReader(stXsl);
                XmlReader xmlReader = XmlReader.Create(srXsl);
                xslTransform.Load(xmlReader);
                IO.StringReader srXml = new IO.StringReader(xmlText);
                XmlReader xrXml = XmlReader.Create(srXml);
                IO.MemoryStream ms = new IO.MemoryStream();
                XmlWriterSettings xmlWriterSettings = new XmlWriterSettings();
                xmlWriterSettings.OmitXmlDeclaration = true;
                XmlWriter xwXml = XmlWriter.Create(ms, xmlWriterSettings);
                xslTransform.Transform(xrXml, xwXml);
                ms.Position = 0;
                IO.StreamReader srHtml = new IO.StreamReader(ms);
                htmlText = srHtml.ReadToEnd();
                WriteObject(htmlText);
            }
            catch (SystemException ex)
            {
                ErrorRecord err = new ErrorRecord(ex, "TransformNbDocument.XslTransform", ErrorCategory.InvalidOperation, null);
                WriteError(err);
            }
        }
    }
}
