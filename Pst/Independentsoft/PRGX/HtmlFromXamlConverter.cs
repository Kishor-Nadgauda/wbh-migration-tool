//---------------------------------------------------------------------------
// 
// File: HtmlFromXamlConverter.cs
//
// Copyright (C) Microsoft Corporation.  All rights reserved.
//
// Description: Prototype for Xaml - Html conversion 
//
//---------------------------------------------------------------------------

namespace HTMLConverter
{
  using System;
  using System.Diagnostics;
  using System.IO;
  using System.Text;
  using System.Text.RegularExpressions;
  using System.Threading;
  using System.Windows.Documents;
  using System.Xml;

  /// <summary>
  /// HtmlToXamlConverter is a static class that takes an HTML string
  /// and converts it into XAML
  /// </summary>
  public static class HtmlFromXamlConverter
  {
    // ---------------------------------------------------------------------
    //
    // Internal Methods
    //
    // ---------------------------------------------------------------------

    #region Internal Methods

    /// <summary>
    /// Main entry point for Xaml-to-Html converter.
    /// Converts a xaml string into html string.
    /// </summary>
    /// <param name="xamlString">
    /// Xaml strinng to convert.
    /// </param>
    /// <returns>
    /// Html string produced from a source xaml.
    /// </returns>
    public static string ConvertXamlToHtml(string xamlString)
    {
      StringBuilder htmlStringBuilder;
      XmlReader xamlReader;
      XmlWriter htmlWriter;
      XmlReaderSettings x;

      xamlReader = XmlReader.Create(new StringReader(xamlString));

      htmlStringBuilder = new StringBuilder(10000);
      htmlWriter = XmlWriter.Create(new StringWriter(htmlStringBuilder));

      if (!WriteFlowDocument(xamlReader, htmlWriter))
      {
        return "";
      }

      htmlWriter.Flush();

      string htmlString = htmlStringBuilder.ToString();

      return htmlString;
    }

    #endregion Internal Methods

    // ---------------------------------------------------------------------
    //
    // Private Methods
    //
    // ---------------------------------------------------------------------

    #region Private Methods
    /// <summary>
    /// Processes a root level element of XAML (normally it's FlowDocument element).
    /// </summary>
    /// <param name="xamlReader">
    /// XmlTextReader for a source xaml.
    /// </param>
    /// <param name="htmlWriter">
    /// XmlTextWriter producing resulting html
    /// </param>
    private static bool WriteFlowDocument(XmlReader xamlReader, XmlWriter htmlWriter)
    {
      if (!ReadNextToken(xamlReader))
      {
        // Xaml content is empty - nothing to convert
        return false;
      }

//      if (xamlReader.NodeType != XmlNodeType.Element || xamlReader.Name != "FlowDocument")
      if (xamlReader.NodeType != XmlNodeType.Element || xamlReader.Name != "Section")
      {
          // Root FlowDocument elemet is missing
          return false;
      }

      // Create a buffer StringBuilder for collecting css properties for inline STYLE attributes
      // on every element level (it will be re-initialized on every level).
      StringBuilder inlineStyle = new StringBuilder();

      htmlWriter.WriteStartElement("HTML");
      htmlWriter.WriteStartElement("HEAD");
      // handle things like Right Single Quote Mark Correctly
      htmlWriter.WriteStartElement("meta");
      htmlWriter.WriteAttributeString("http-equiv", "Content-Type");
      htmlWriter.WriteAttributeString("content", "text/html");
      htmlWriter.WriteAttributeString("charset", "UTF-8");
      htmlWriter.WriteEndElement();

      htmlWriter.WriteStartElement("STYLE");
      htmlWriter.WriteString("\ntable {border-collapse:collapse}\n");
      htmlWriter.WriteString("td {border:solid windowtext 1.0pt}\n");
     // htmlWriter.WriteString("body {background-color: lightgrey }\n"); // ("body {filter: contrast(10%);}\n"); ("body {filter: invert(50%);}\n");
      htmlWriter.WriteEndElement();
      htmlWriter.WriteEndElement();

      htmlWriter.WriteStartElement("BODY");



      WriteFormattingProperties(xamlReader, htmlWriter, inlineStyle);

      WriteElementContent(xamlReader, htmlWriter, inlineStyle);

      htmlWriter.WriteEndElement();
      htmlWriter.WriteEndElement();

      return true;
    }

    /// <summary>
    /// Reads attributes of the current xaml element and converts
    /// them into appropriate html attributes or css styles.
    /// </summary>
    /// <param name="xamlReader">
    /// XmlTextReader which is expected to be at XmlNodeType.Element
    /// (opening element tag) position.
    /// The reader will remain at the same level after function complete.
    /// </param>
    /// <param name="htmlWriter">
    /// XmlTextWriter for output html, which is expected to be in
    /// after WriteStartElement state.
    /// </param>
    /// <param name="inlineStyle">
    /// String builder for collecting css properties for inline STYLE attribute.
    /// </param>
    private static void WriteFormattingProperties(XmlReader xamlReader, XmlWriter htmlWriter, StringBuilder inlineStyle)
    {
      Debug.Assert(xamlReader.NodeType == XmlNodeType.Element);

      // Clear string builder for the inline style
      inlineStyle.Remove(0, inlineStyle.Length);

      if (!xamlReader.HasAttributes)
      {
        return;
      }

      bool borderSet = false;

      while (xamlReader.MoveToNextAttribute())
      {
        string css = null;

        switch (xamlReader.Name)
        {
          // Character fomatting properties
          // ------------------------------
          case "Background":
            css = "background-color:" + ParseXamlColor(xamlReader.Value) + ";";
            break;
          case "FontFamily":
            css = "font-family:" + xamlReader.Value + ";";
            break;
          case "FontStyle":
            css = "font-style:" + xamlReader.Value.ToLower() + ";";
            break;
          case "FontWeight":
            css = "font-weight:" + xamlReader.Value.ToLower() + ";";
            break;
          case "FontStretch":
            break;
          case "FontSize":
            css = "font-size:" + xamlReader.Value + ";";
            break;
          case "Foreground":
            css = "color:" + ParseXamlColor(xamlReader.Value) + ";";
            break;
          case "TextDecorations":
            css = "text-decoration:underline;";
            break;
          case "TextEffects":
            break;
          case "Emphasis":
            break;
          case "StandardLigatures":
            break;
          case "Variants":
            break;
          case "Capitals":
            break;
          case "Fraction":
            break;
          case "Typography.Variants":
            css = "vertical-align: super; font-size: 60%";
            break;

          // Paragraph formatting properties
          // -------------------------------
          case "Padding":
            css = "padding:" + ParseXamlThickness(xamlReader.Value) + ";";
            break;
          case "Margin":
            css = "margin:" + ParseXamlThickness(xamlReader.Value) + ";";
            break;
          case "BorderThickness":
            css = "border-width:" + ParseXamlThickness(xamlReader.Value) + ";";
            borderSet = true;
            break;
          case "CellSpacing":
            css = string.Format("padding:{0}in 5.4pt {0}in 5.4;", xamlReader.Value);
            break;
          case "BorderBrush":
            css = "border-color:" + ParseXamlColor(xamlReader.Value) + ";";
            borderSet = true;
            break;
          case "LineHeight":
            break;
          case "TextIndent":
            css = "text-indent:" + xamlReader.Value + ";";
            break;
          case "TextAlignment":
            css = "text-align:" + xamlReader.Value + ";";
            break;
          case "IsKeptTogether":
            break;
          case "IsKeptWithNext":
            break;
          case "ColumnBreakBefore":
            break;
          case "PageBreakBefore":
            break;
          case "FlowDirection":
            break;

          // Table attributes
          // ----------------
          case "Width":
            css = "width:" + xamlReader.Value + ";";
            break;
          case "ColumnSpan":
            htmlWriter.WriteAttributeString("COLSPAN", xamlReader.Value);
            break;
          case "RowSpan":
            htmlWriter.WriteAttributeString("ROWSPAN", xamlReader.Value);
            break;

          // additional attributes
          case "xml:space":
            if (xamlReader.Value.Equals("preserve", StringComparison.InvariantCultureIgnoreCase))
            {
              css = "white-space: pre;";
            } else
            {
              css = "white-space: normal;";
            }
            //htmlWriter.WriteAttributeString("xml:space", xamlReader.Value);
            break;
          case "xml:lang":
            css = ":lang(" + xamlReader.Value + ");";
            //htmlWriter.WriteAttributeString("xml:lang", xamlReader.Value);
            break;
          case "xmlns":
            break;

          case "NavigateUri":
            htmlWriter.WriteAttributeString("href", xamlReader.Value);
            break;

          default:
            break;
        }

        if (css != null)
        {
          inlineStyle.Append(css);
        }
      }

      if (borderSet)
      {
        inlineStyle.Append("border-style:solid windowtext 1.0pt;");
      }

      // Return the xamlReader back to element level
      xamlReader.MoveToElement();
      Debug.Assert(xamlReader.NodeType == XmlNodeType.Element);
    }

    /// <summary>
    /// Extract RTF text to HTML
    /// </summary>
    /// <param name="bodyRtf">The RTF in bytes</param>
    /// <param name="length">The length of the body to use</param>
    /// <returns>RTF converted to HTML</returns>
    public static string extractRtf2HtmlAsStaThread(byte[] bodyRtf, int length)
    {
      string retVal = string.Empty;
      Thread thread = new Thread(() => { retVal = HtmlFromXamlConverter.extractRtf2Html(bodyRtf, length); } );
      thread.SetApartmentState(ApartmentState.STA);
      thread.Start();
      thread.Join();
      return retVal;
    }

    static string strRegex = @"fontfamily=""+(?<font>[^\""]+)""+"; // encountering isses with multiple double quotes, sometimes two on just the lead/trail 
    static string strRegexApos = @"fontfamily=""+(?:\w*[^a-zA-Z0-9 ""]+)+\w*""+"; // find any font family name such as &apos or other control chracters and replace with arial
    static string strRegexFormFeed = @"\f"; // find any form feed and replace it with proper HTML
    static Regex myRegex = new Regex(strRegex, RegexOptions.IgnoreCase | RegexOptions.IgnorePatternWhitespace | RegexOptions.Compiled);
    static Regex myRegexApos = new Regex(strRegexApos, RegexOptions.IgnoreCase | RegexOptions.IgnorePatternWhitespace | RegexOptions.Compiled);
    static Regex myRegexFormFeed = new Regex(strRegexFormFeed, RegexOptions.IgnoreCase | RegexOptions.Compiled);
    static string strReplace = @"FontFamily=""${font}""";
    static string strReplaceApos = @"FontFamily=""Arial""";
    static string strReplaceFormFeed = "<br style=\"page-break-before: always\">";

    static int fid = 0;
    /// <summary>
    /// Extract RTF text to HTML
    /// </summary>
    /// <param name="bodyRtf">The RTF in bytes</param>
    /// <param name="length">The length of the body to use</param>
    /// <returns>RTF converted to HTML</returns>
    public static string extractRtf2Html(byte[] bodyRtf, int length)
    {
      
      string retVal = string.Empty;
      string xamlContent = string.Empty;
      string regexCleanup = string.Empty;
      string rtfBody = Encoding.UTF8.GetString(bodyRtf, 0, length);  // first pass try using the RTF as a text only version
      string wpsText = string.Empty;

      try
      {
        System.Windows.Documents.XamlRtfConverter xamlRtfConverter = new XamlRtfConverter();
//        using (Stream stream = new MemoryStream())
//        {
//          xamlRtfConverter.WpfPayload = WpfPayload.CreateWpfPayload(stream);
          xamlContent = xamlRtfConverter.ConvertRtfToXaml(bodyRtf);
//          stream.Position = 0;
//          using (StreamReader streamReader = new StreamReader(stream))
//          {
//            wpsText = streamReader.ReadToEnd();
//          }
//        }
        //System.Windows.DataObject dataObject = new System.Windows.DataObject(System.Windows.DataFormats.Rtf, bodyRtf);
        //var assembly = System.Reflection.Assembly.GetAssembly(typeof(System.Windows.FrameworkElement));
        //var xamlRtfConverterType = assembly.GetType("System.Windows.Documents.XamlRtfConverter");
        //var xamlRtfConverter = Activator.CreateInstance(xamlRtfConverterType, true);
        //var convertRtfToXaml = xamlRtfConverterType.GetMethod("ConvertRtfToXaml", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
        //xamlContent = (string)convertRtfToXaml.Invoke(xamlRtfConverter, new object[] { rtfBody });
        regexCleanup = myRegex.Replace(myRegexApos.Replace(xamlContent, strReplaceApos), strReplace);

        retVal = HTMLConverter.HtmlFromXamlConverter.ConvertXamlToHtml(regexCleanup);

        //File.WriteAllText(string.Format(@"c:\x\rtf\r{0}.txt", fid.ToString("D7")), string.Format("{0}\n\n\n{1}\n\n\n{2}", rtfString, xamlContent, retVal));
        //File.WriteAllText(string.Format(@"c:\x\rtf\r{0}.rtf", fid.ToString("D7")), rtfString);
        //File.WriteAllText(string.Format(@"c:\x\rtf\r{0}.htm", fid.ToString("D7")), retVal);
        //File.WriteAllText(string.Format(@"c:\x\rtf\r{0}.xaml", fid.ToString("D7")), xamlContent);
        //File.WriteAllText(string.Format(@"c:\x\rtf\r{0}.xaml1", fid.ToString("D7")), regexCleanup);
        //fid++;
      }
      catch (Exception ex)
      {
        retVal = string.Empty;
      }
//      finally
//      {
//        if (retVal.Length == 0)
//        {
//          try
//          {
//            System.Windows.DataObject dataObject =  new System.Windows.DataObject(System.Windows.DataFormats.Rtf, bodyRtf);

//            //var range = new TextRange(fd.ContentStart, fd.ContentEnd);
//            //range.Load(streamRtf, System.Windows.DataFormats.Rtf);
//            //using (var outputStream = new MemoryStream())
//            //{
//            // range.Save(outputStream, System.Windows.DataFormats.Xaml);
//            // outputStream.Position = 0;
//            //   using (var xamlStream = new StreamReader(outputStream))
//            //  {
//            //  var xaml = xamlStream.ReadToEnd();
//            var xaml = System.Windows.Clipboard.GetData(System.Windows.DataFormats.Xaml).ToString();
//                  regexCleanup = myRegex.Replace(myRegexApos.Replace(xaml, strReplaceApos), strReplace);
//                  retVal = HTMLConverter.HtmlFromXamlConverter.ConvertXamlToHtml(regexCleanup);
//              //  }
//             // }
//              //                RichTextBox richTextBox = new RichTextBox();
//              //              richTextBox.LoadFile(streamRtf, RichTextBoxStreamType.RichText);
//              //              rtfBody = richTextBox.Rtf;
//              //              var assembly = System.Reflection.Assembly.GetAssembly(typeof(System.Windows.FrameworkElement));
//              //              var xamlRtfConverterType = assembly.GetType("System.Windows.Documents.XamlRtfConverter");
//              //              var xamlRtfConverter = Activator.CreateInstance(xamlRtfConverterType, true);
//              //              var convertRtfToXaml = xamlRtfConverterType.GetMethod("ConvertRtfToXaml", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
//              //              xamlContent = (string)convertRtfToXaml.Invoke(xamlRtfConverter, new object[] { rtfBody });
////              regexCleanup = myRegex.Replace(myRegexApos.Replace(xamlContent, strReplaceApos), strReplace);

////              retVal = HTMLConverter.HtmlFromXamlConverter.ConvertXamlToHtml(regexCleanup);
//            //}
//          }
//          catch (Exception)
//          {
//            throw;
//          }
//        }
//      }

      return retVal;
    }

    private static string ParseXamlColor(string color)
    {
      if (color.StartsWith("#"))
      {
        // Remove transparancy value
        color = "#" + color.Substring(3);
      }
      return color;
    }

    private static string ParseXamlThickness(string thickness)
    {
      string[] values = thickness.Split(',');

      for (int i = 0; i < values.Length; i++)
      {
        double value;
        if (double.TryParse(values[i], out value))
        {
          values[i] = Math.Ceiling(value).ToString();
        }
        else
        {
          values[i] = "1";
        }
      }

      string cssThickness;
      switch (values.Length)
      {
        case 1:
          cssThickness = thickness;
          break;
        case 2:
          cssThickness = values[1] + " " + values[0];
          break;
        case 4:
          cssThickness = values[1] + " " + values[2] + " " + values[3] + " " + values[0];
          break;
        default:
          cssThickness = values[0];
          break;
      }

      return cssThickness;
    }

    /// <summary>
    /// Reads a content of current xaml element, converts it
    /// </summary>
    /// <param name="xamlReader">
    /// XmlTextReader which is expected to be at XmlNodeType.Element
    /// (opening element tag) position.
    /// </param>
    /// <param name="htmlWriter">
    /// May be null, in which case we are skipping the xaml element;
    /// witout producing any output to html.
    /// </param>
    /// <param name="inlineStyle">
    /// StringBuilder used for collecting css properties for inline STYLE attribute.
    /// </param>
    private static void WriteElementContent(XmlReader xamlReader, XmlWriter htmlWriter, StringBuilder inlineStyle)
    {
      Debug.Assert(xamlReader.NodeType == XmlNodeType.Element);

      bool elementContentStarted = false;

      if (xamlReader.IsEmptyElement)
      {
        if (htmlWriter != null && !elementContentStarted && inlineStyle.Length > 0)
        {
          // Output STYLE attribute and clear inlineStyle buffer.
          htmlWriter.WriteAttributeString("STYLE", inlineStyle.ToString());
          inlineStyle.Remove(0, inlineStyle.Length);
        }
        elementContentStarted = true;
      }
      else
      {
        while (ReadNextToken(xamlReader) && xamlReader.NodeType != XmlNodeType.EndElement)
        {
          switch (xamlReader.NodeType)
          {
            case XmlNodeType.Element:
              if (xamlReader.Name.Contains("."))
              {
                AddComplexProperty(xamlReader, inlineStyle);
              }
              else
              {
                if (htmlWriter != null && !elementContentStarted && inlineStyle.Length > 0)
                {
                  // Output STYLE attribute and clear inlineStyle buffer.
                  htmlWriter.WriteAttributeString("STYLE", inlineStyle.ToString());
                  inlineStyle.Remove(0, inlineStyle.Length);
                }
                elementContentStarted = true;
                WriteElement(xamlReader, htmlWriter, inlineStyle);
              }
              Debug.Assert(xamlReader.NodeType == XmlNodeType.EndElement || xamlReader.NodeType == XmlNodeType.Element && xamlReader.IsEmptyElement);
              break;
            case XmlNodeType.Comment:
              if (htmlWriter != null)
              {
                if (!elementContentStarted && inlineStyle.Length > 0)
                {
                  htmlWriter.WriteAttributeString("STYLE", inlineStyle.ToString());
                }
                htmlWriter.WriteComment(xamlReader.Value);
              }
              elementContentStarted = true;
              break;
            case XmlNodeType.CDATA:
            case XmlNodeType.Text:
            case XmlNodeType.SignificantWhitespace:
              if (htmlWriter != null)
              {
                if (!elementContentStarted && inlineStyle.Length > 0)
                {
                  htmlWriter.WriteAttributeString("STYLE", inlineStyle.ToString());
                }
                htmlWriter.WriteString(xamlReader.Value);
              }
              elementContentStarted = true;
              break;
            default:
              break;
          }
        }

        Debug.Assert(xamlReader.NodeType == XmlNodeType.EndElement);
      }
    }

    /// <summary>
    /// Conberts an element notation of complex property into
    /// </summary>
    /// <param name="xamlReader">
    /// On entry this XmlTextReader must be on Element start tag;
    /// on exit - on EndElement tag.
    /// </param>
    /// <param name="inlineStyle">
    /// StringBuilder containing a value for STYLE attribute.
    /// </param>
    private static void AddComplexProperty(XmlReader xamlReader, StringBuilder inlineStyle)
    {
      Debug.Assert(xamlReader.NodeType == XmlNodeType.Element);

      if (inlineStyle != null && xamlReader.Name.EndsWith(".TextDecorations"))
      {
        inlineStyle.Append("text-decoration:underline;");
      }

      // Skip the element representing the complex property
      WriteElementContent(xamlReader, /*htmlWriter:*/null, /*inlineStyle:*/null);
    }

    /// <summary>
    /// Converts a xaml element into an appropriate html element.
    /// </summary>
    /// <param name="xamlReader">
    /// On entry this XmlTextReader must be on Element start tag;
    /// on exit - on EndElement tag.
    /// </param>
    /// <param name="htmlWriter">
    /// May be null, in which case we are skipping xaml content
    /// without producing any html output
    /// </param>
    /// <param name="inlineStyle">
    /// StringBuilder used for collecting css properties for inline STYLE attributes on every level.
    /// </param>
    private static void WriteElement(XmlReader xamlReader, XmlWriter htmlWriter, StringBuilder inlineStyle)
    {
      Debug.Assert(xamlReader.NodeType == XmlNodeType.Element);

      if (htmlWriter == null)
      {
        // Skipping mode; recurse into the xaml element without any output
        WriteElementContent(xamlReader, /*htmlWriter:*/null, null);
      }
      else
      {
        string htmlElementName = null;
        string classAttribute = string.Empty;

        switch (xamlReader.Name)
        {
          case "Run":
          case "Span":
            htmlElementName = "SPAN";
            break;
          case "InlineUIContainer":
            htmlElementName = "SPAN";
            break;
          case "Bold":
            htmlElementName = "B";
            break;
          case "Italic":
            htmlElementName = "I";
            break;
          case "Paragraph":
            htmlElementName = "P";
            classAttribute = "MsoNormal";
            break;
          case "BlockUIContainer":
            htmlElementName = "DIV";
            classAttribute = "WordSection1";
            break;
          case "Section":
            htmlElementName = "DIV";
            classAttribute = "MsoNormal";
            break;
          case "Table":
            htmlElementName = "TABLE";
            classAttribute = "MsoNormalTable";
            break;
          case "TableColumn":
            htmlElementName = "COL";
            break;
          case "TableRowGroup":
            htmlElementName = "TBODY";
            break;
          case "TableRow":
            htmlElementName = "TR";
            break;
          case "TableCell":
            htmlElementName = "TD";
            break;
          case "List":
            string marker = xamlReader.GetAttribute("MarkerStyle");
            if (marker == null || marker == "None" || marker == "Disc" || marker == "Circle" || marker == "Square" || marker == "Box")
            {
              htmlElementName = "UL";
            }
            else
            {
              htmlElementName = "OL";
            }
            break;
          case "ListItem":
            htmlElementName = "LI";
            classAttribute = "MsoNormal";
            break;
          case "LineBreak":
            htmlElementName = "br";
            break;
          case "Hyperlink":
            htmlElementName = "A";
            break;
          default:
            htmlElementName = null; // Ignore the element
            break;
        }

        if (htmlWriter != null && htmlElementName != null)
        {
          htmlWriter.WriteStartElement(htmlElementName);

          if (classAttribute.Length > 0)
          {
           // htmlWriter.WriteAttributeString("class", classAttribute);
          }

          WriteFormattingProperties(xamlReader, htmlWriter, inlineStyle);

          WriteElementContent(xamlReader, htmlWriter, inlineStyle);

          htmlWriter.WriteEndElement();
        }
        else
        {
          // Skip this unrecognized xaml element
          WriteElementContent(xamlReader, /*htmlWriter:*/null, null);
        }
      }
    }

    // Reader advance helpers
    // ----------------------

    /// <summary>
    /// Reads several items from xamlReader skipping all non-significant stuff.
    /// </summary>
    /// <param name="xamlReader">
    /// XmlTextReader from tokens are being read.
    /// </param>
    /// <returns>
    /// True if new token is available; false if end of stream reached.
    /// </returns>
    private static bool ReadNextToken(XmlReader xamlReader)
    {
      while (xamlReader.Read())
      {
        Debug.Assert(xamlReader.ReadState == ReadState.Interactive, "Reader is expected to be in Interactive state (" + xamlReader.ReadState + ")");
        switch (xamlReader.NodeType)
        {
          case XmlNodeType.Element:
          case XmlNodeType.EndElement:
          case XmlNodeType.None:
          case XmlNodeType.CDATA:
          case XmlNodeType.Text:
          case XmlNodeType.SignificantWhitespace:
            return true;

          case XmlNodeType.Whitespace:
            if (xamlReader.XmlSpace == XmlSpace.Preserve)
            {
              return true;
            }
            // ignore insignificant whitespace
            break;

          case XmlNodeType.EndEntity:
          case XmlNodeType.EntityReference:
            //  Implement entity reading
            //xamlReader.ResolveEntity();
            //xamlReader.Read();
            //ReadChildNodes( parent, parentBaseUri, xamlReader, positionInfo);
            break; // for now we ignore entities as insignificant stuff

          case XmlNodeType.Comment:
            return true;
          case XmlNodeType.ProcessingInstruction:
          case XmlNodeType.DocumentType:
          case XmlNodeType.XmlDeclaration:
          default:
            // Ignorable stuff
            break;
        }
      }
      return false;
    }

    #endregion Private Methods

    // ---------------------------------------------------------------------
    //
    // Private Fields
    //
    // ---------------------------------------------------------------------

    #region Private Fields

    #endregion Private Fields
  }
}
