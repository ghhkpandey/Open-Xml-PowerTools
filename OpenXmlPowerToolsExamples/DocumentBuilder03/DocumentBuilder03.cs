// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OpenXmlPowerTools;

class Program
{
    static void Main(string[] args)
    {
        //var n = DateTime.Now;
        //var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
        //tempDi.Create();

        //WmlDocument doc1 = new WmlDocument(@"..\..\..\Template.docx");
        //using (MemoryStream mem = new MemoryStream())
        //{
        //    mem.Write(doc1.DocumentByteArray, 0, doc1.DocumentByteArray.Length);
        //    using (WordprocessingDocument doc = WordprocessingDocument.Open(mem, true))
        //    {
        //        XDocument xDoc = doc.MainDocumentPart.GetXDocument();
        //        XElement frontMatterPara = xDoc.Root.Descendants(W.txbxContent).Elements(W.p).FirstOrDefault();
        //        frontMatterPara.ReplaceWith(
        //            new XElement(PtOpenXml.Insert,
        //                new XAttribute("Id", "Front")));
        //        XElement tbl = xDoc.Root.Element(W.body).Elements(W.tbl).FirstOrDefault();
        //        XElement firstCell = tbl.Descendants(W.tr).First().Descendants(W.p).First();
        //        firstCell.ReplaceWith(
        //            new XElement(PtOpenXml.Insert,
        //                new XAttribute("Id", "Liz")));
        //        XElement secondCell = tbl.Descendants(W.tr).Skip(1).First().Descendants(W.p).First();
        //        secondCell.ReplaceWith(
        //            new XElement(PtOpenXml.Insert,
        //                new XAttribute("Id", "Eric")));
        //        doc.MainDocumentPart.PutXDocument();
        //    }
        //    doc1.DocumentByteArray = mem.ToArray();
        //}

        //string outFileName = Path.Combine(tempDi.FullName, "Out.docx");
        //List<Source> sources = new List<Source>()
        //    {
        //        new Source(doc1, true),
        //        new Source(new WmlDocument(@"..\..\..\Insert-01.docx"), "Liz"),
        //        new Source(new WmlDocument(@"..\..\..\Insert-02.docx"), "Eric"),
        //        new Source(new WmlDocument(@"..\..\..\FrontMatter.docx"), "Front"),
        //    };
        //DocumentBuilder.BuildDocument(sources, outFileName);

        ReBuildTemplate();
    }

    private static void ReBuildTemplate()
    {
        Dictionary<string, string> tagToFile = new Dictionary<string, string>
        {
            { "DBB_RecipientName", @"DemoBuildingBlock3.docx" },
            { "DBB_Signature", @"DemoSignature3.docx" },
            { "DBH_HeaderFooter", @"DemoHeaderFooter3.docx" },
        };

        WmlDocument doc1 = new WmlDocument(@"DemoTemplate3.docx");
        using (MemoryStream mem = new MemoryStream())
        {
            mem.Write(doc1.DocumentByteArray, 0, doc1.DocumentByteArray.Length);
            using (WordprocessingDocument doc = WordprocessingDocument.Open(mem, true))
            {
                // Replace content controls whose tags match a tagToFile key with a PtOpenXml.Insert
                // element in the body, headers and footers.
                ReplaceTaggedContentControlsWithInserts(doc.MainDocumentPart.GetXDocument(), tagToFile);
                doc.MainDocumentPart.PutXDocument();

                foreach (HeaderPart headerPart in doc.MainDocumentPart.HeaderParts)
                {
                    ReplaceTaggedContentControlsWithInserts(headerPart.GetXDocument(), tagToFile);
                    headerPart.PutXDocument();
                }

                foreach (FooterPart footerPart in doc.MainDocumentPart.FooterParts)
                {
                    ReplaceTaggedContentControlsWithInserts(footerPart.GetXDocument(), tagToFile);
                    footerPart.PutXDocument();
                }
            }
            doc1.DocumentByteArray = mem.ToArray();
        }

        var n = DateTime.Now;
        var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
        tempDi.Create();

        string outFileName = Path.Combine(tempDi.FullName, "Out3.docx");
        List<Source> sources = new List<Source>
        {
            new Source(doc1, true),
        };
        foreach (KeyValuePair<string, string> entry in tagToFile)
            sources.Add(new Source(new WmlDocument(entry.Value), entry.Key, false)
            {
                KeepHeaderOrFooterOnly = entry.Key.StartsWith("DBH_")
            }
            );

        DocumentBuilder.BuildDocument(sources, outFileName, new DocumentBuilderSettings() { NormalizeStyleIds = true });
    }

    // Content element names that represent meaningful content inside a w:sdtContent and should
    // be replaced by the PtOpenXml.Insert placeholder.  Everything else (bookmarks, proofErr,
    // permission fences, tracked-change wrappers, etc.) is intentionally left untouched.
    private static readonly HashSet<XName> ContentElementNames = new HashSet<XName>
    {
        W.sdt,      // nested sdt
        W.p,       // block-level sdt  (paragraph)
        W.r,       // inline/run sdt   (run)
        W.tbl,     // table sdt
        W.tr,      // row sdt
        W.tc,      // cell sdt
        W.drawing, // inline drawing
        W.pict,    // VML picture
    };

    private static void ReplaceTaggedContentControlsWithInserts(XDocument xDoc, Dictionary<string, string> tagToFile)
    {
        // Materialise the list first so tree mutations inside the loop do not affect enumeration.
        List<XElement> sdtsToReplace = xDoc.Descendants(W.sdt)
            .Where(sdt =>
            {
                XElement sdtPr = sdt.Element(W.sdtPr);
                if (sdtPr == null) return false;
                XElement tagElement = sdtPr.Element(W.tag);
                if (tagElement == null) return false;
                string tagVal = (string)tagElement.Attribute(W.val);
                return tagVal != null && tagToFile.ContainsKey(tagVal);
            })
            .ToList();

        foreach (XElement sdt in sdtsToReplace)
        {
            string insertId = (string)sdt.Element(W.sdtPr).Element(W.tag).Attribute(W.val);
            XElement sdtContent = sdt.Element(W.sdtContent);

            if (sdtContent == null)
            {
                // No sdtContent yet — create one containing only the Insert element.
                sdt.Add(new XElement(W.sdtContent,
                    new XElement(PtOpenXml.Insert,
                        new XAttribute(PtOpenXml.Id, insertId))));
                continue;
            }

            // Remove only meaningful content nodes (paragraphs, runs, tables, rows, cells,
            // drawings, pictures).  Non-content nodes such as bookmarks, proofErr, permission
            // fences and tracked-change wrappers are preserved so the document remains valid.
            List<XElement> contentNodesToRemove = sdtContent
                .Elements()
                .Where(e => ContentElementNames.Contains(e.Name))
                .ToList();

            if (contentNodesToRemove.Count > 0)
            {
                // Insert the placeholder before the first content node, then remove them all.
                contentNodesToRemove[0].AddBeforeSelf(
                    new XElement(PtOpenXml.Insert,
                        new XAttribute(PtOpenXml.Id, insertId)));
                foreach (XElement node in contentNodesToRemove)
                    node.Remove();
            }
            else
            {
                // sdtContent exists but holds no recognisable content nodes — just prepend.
                sdtContent.AddFirst(
                    new XElement(PtOpenXml.Insert,
                        new XAttribute(PtOpenXml.Id, insertId)));
            }
        }
    }
}
