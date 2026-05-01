// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    // Marker annotation used to track that a package is inside a PowerTools block.
    internal class PowerToolsBlockMarker { }

    public static class PowerToolsBlockExtensions
    {
        /// <summary>
        /// Begins a PowerTools Block by (1) removing annotations and, unless the package was
        /// opened in read-only mode, (2) saving the package.
        /// </summary>
        /// <remarks>
        /// Removes <see cref="XDocument" /> and <see cref="XmlNamespaceManager" /> instances
        /// added by <see cref="PtOpenXmlExtensions.GetXDocument(OpenXmlPart)" />,
        /// <see cref="PtOpenXmlExtensions.GetXDocument(OpenXmlPart, out XmlNamespaceManager)" />,
        /// <see cref="PtOpenXmlExtensions.PutXDocument(OpenXmlPart)" />,
        /// <see cref="PtOpenXmlExtensions.PutXDocument(OpenXmlPart, XDocument)" />, and
        /// <see cref="PtOpenXmlExtensions.PutXDocumentWithFormatting(OpenXmlPart)" />.
        /// methods.
        /// </remarks>
        /// <param name="package">
        /// A <see cref="WordprocessingDocument" />, <see cref="SpreadsheetDocument" />,
        /// or <see cref="PresentationDocument" />.
        /// </param>
        public static void BeginPowerToolsBlock(this OpenXmlPackage package)
        {
            if (package == null) throw new ArgumentNullException("package");

            package.RemovePowerToolsAnnotations();
            package.Save();

            // Mark the package as being in a PowerTools block so that PutXDocument
            // defers stream writes and only updates the cached annotation.
            package.AddAnnotation(new PowerToolsBlockMarker());
        }

        /// <summary>
        /// Ends a PowerTools Block by flushing all cached XDocument annotations to the
        /// part streams and then reloading the root elements of all changed parts.
        /// </summary>
        /// <param name="package">
        /// A <see cref="WordprocessingDocument" />, <see cref="SpreadsheetDocument" />,
        /// or <see cref="PresentationDocument" />.
        /// </param>
        public static void EndPowerToolsBlock(this OpenXmlPackage package)
        {
            if (package == null) throw new ArgumentNullException("package");

            // Remove the PowerTools block marker first so that PutXDocument
            // will write to the stream normally during the flush below.
            package.RemoveAnnotations<PowerToolsBlockMarker>();

            foreach (OpenXmlPart part in package.GetAllParts())
            {
                if (part.Annotations<XDocument>().Any())
                {
                    // Flush the cached XDocument to the part stream.
                    part.PutXDocument();

                    // Reload the SDK's root element from the updated stream.
                    if (part.RootElement != null)
                        part.RootElement.Reload();
                }
            }
        }

        /// <summary>
        /// Returns true if the package is currently inside a PowerTools block.
        /// </summary>
        internal static bool IsInPowerToolsBlock(this OpenXmlPackage package)
        {
            return package != null && package.Annotations<PowerToolsBlockMarker>().Any();
        }

        private static void RemovePowerToolsAnnotations(this OpenXmlPackage package)
        {
            if (package == null) throw new ArgumentNullException("package");

            foreach (OpenXmlPart part in package.GetAllParts())
            {
                part.RemoveAnnotations<XDocument>();
                part.RemoveAnnotations<XmlNamespaceManager>();
            }
        }
    }
}
