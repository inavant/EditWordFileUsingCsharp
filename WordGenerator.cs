using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;


namespace InaOfficeTools
{
    /// <summary>
    /// Main class to create word document
    /// </summary>
    public class WordGenerator
    {
        private WordprocessingDocument wpDocument { get; set; }

        public WordGenerator(WordprocessingDocument doc) {
            this.wpDocument = doc;
        }

        /// <summary>
        /// Inserts a line of text
        /// </summary>
        /// <param name="TagName">Document TAG where value will be inserted</param>
        /// <param name="TagValue">Value</param>
        public void UpdateTextoControlWord(string TagName, string TagValue) {


            if (TagValue != null)
            {
                var Elements = wpDocument.MainDocumentPart.Document.Descendants<SdtElement>().Where(run => run.SdtProperties.GetFirstChild<Tag>().Val.Value == TagName);
                foreach (SdtElement BElement in Elements)
                {
                    //Gets text nodes
                    List<Text> TextTemplate = BElement.Descendants<Text>().ToList<Text>();

                    for (int i = 0; i < TextTemplate.Count; i++)
                    {
                        //Edits first node
                        Text TagText = (Text)TextTemplate[i];
                        TagText.Text = TagValue;                      
                    }
                }

                foreach (var headerPart in wpDocument.MainDocumentPart.HeaderParts)
                {
                    //Gets the text in header
                    var hElements = headerPart.RootElement.Descendants<SdtElement>().Where(run => run.SdtProperties.GetFirstChild<Tag>().Val.Value == TagName);

                    foreach (SdtElement BElement in hElements)
                    {
                        //Gets text nodes
                        List<Text> TextTemplate = BElement.Descendants<Text>().ToList<Text>();

                        for (int i = 0; i < TextTemplate.Count; i++)
                        {
                            //Edits first node
                            Text TagText = (Text)TextTemplate[i];
                            TagText.Text = TagValue;                           
                        }
                    }

                }                

            }


        }

        /// <summary>
        /// Inserts bullets
        /// </summary>
        /// <param name="TagName">Document TAG where value will be inserted</param>
        /// <param name="itemsList">Bullets structure  to be inserted</param>
        public void UpdateBulletsControlWord(string TagName, List<BulletsConfigWordGenerator> itemsList)
        {
            var Elements = wpDocument.MainDocumentPart.Document.Descendants<SdtElement>().Where(run => run.SdtProperties.GetFirstChild<Tag>().Val.Value == TagName);

            SetNumberingConfiguration();
            List<Text> TextTemplate = null;
            foreach (SdtElement BElement in Elements)
            {
                TextTemplate = BElement.Descendants<Text>().ToList<Text>();
                for (int i = 0; i < TextTemplate.Count; i++)
                {
                    Run ContentControlRun = null;
                    if (BElement.Descendants<Run>().Any())
                    {
                        ContentControlRun = BElement.Descendants<Run>().FirstOrDefault();
                        ContentControlRun.RemoveAllChildren();

                        Run runItem = new Run();
                        foreach(BulletsConfigWordGenerator itemBullet in itemsList)
                            AppendListItem(runItem, itemBullet.Text, itemBullet.Lavel, itemBullet.MargingLeft, itemBullet.fontSize,  (itemBullet.IsUnderLine?UnderlineValues.Single: UnderlineValues.None), itemBullet.Isbold);                      

                        ContentControlRun.Append(runItem);

                    }
                }

            }

        }
        
        private void AppendListItem(Run body, string content, int level, string margin, string fontSize, UnderlineValues underLineType, bool IsBoldText)
        {

            Paragraph paragraph1 = new Paragraph();
            ParagraphProperties paragraphProperties = new ParagraphProperties();
            paragraphProperties.Append(new SpacingBetweenLines() { Line = "480", LineRule = LineSpacingRuleValues.AtLeast, Before = "0", After = "0" });

            NumberingProperties numberingProperties1 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference = new NumberingLevelReference() { Val = level };

            NumberingId numberingId = new NumberingId() { Val = 1 };
            NumberingFormat numberingFormat = new NumberingFormat() { Val = NumberFormatValues.UpperRoman };
            numberingProperties1.Append(numberingLevelReference);
            numberingProperties1.Append(numberingId);
            numberingProperties1.Append(numberingFormat);

            Run run = new Run();
            RunProperties RunProperty = new RunProperties();
            RunProperty.Append(new Color() { Val = "#000000" });
            RunProperty.Append(new FontSize() { Val = fontSize });
            RunProperty.Append(new RunStyle() { Val = "PlaceholderText" });
            RunProperty.Append(new RunFonts() { Ascii = "Arial" });
            if (IsBoldText)
                RunProperty.Append(new Bold());
            if (underLineType != UnderlineValues.None)
                RunProperty.Append(new Underline() { Val = underLineType });

            run.Append(RunProperty);
            Text text = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text.Text = content;
            run.Append(text);
            paragraph1.Append(paragraphProperties);
            paragraph1.Append(numberingProperties1);
            paragraph1.Append(run);
            body.Append(paragraph1);

        }

        /// <summary>
        /// Insert a table
        /// </summary>
        /// <param name="TblName">Name of existing Table</param>
        /// <param name="tableConf">Structure of Rows,Cells and Values to be inserted</param>
        public void UpdateTablaControlWord(string TblName, TableConfigWordGenerator tableConf) {

            TableCaption foundTableCaption = null;
            Table detailsTable = null;
            TableRow baseRow = null;

            foreach (var e in wpDocument.MainDocumentPart.Document.Body.ChildElements.Where(ce => ce.GetType() == typeof(Table)))
            {
                foundTableCaption = null;
                foundTableCaption = (TableCaption)e.ChildElements.Where(cet => cet.GetType() == typeof(TableProperties))
                                                                 .SelectMany(cet => cet.ChildElements)
                                                                 .FirstOrDefault(ce => ce.GetType() == typeof(TableCaption)
                                                                                       && ((TableCaption)ce).Val == TblName);



                if (foundTableCaption != null)
                {
                    detailsTable = (Table)e;
                    baseRow = (TableRow)detailsTable.ChildElements.LastOrDefault(ce => ce.GetType() == typeof(TableRow));
                    break;
                }
            }

            if (baseRow != null) 
            { 
                baseRow.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableCell>().ElementAt(0).Descendants<Paragraph>().First().Descendants<Text>().First().Text = "CELDA0";
                TableRow rowCopy = null;

                for (int idxRow=0; idxRow< tableConf.RowList.Count; idxRow++) 
                {
                    // Clones first row after the  second row
                    if (idxRow>0)
                        rowCopy = (TableRow)baseRow.CloneNode(true);

                    for (int idxCell = 0; idxCell < tableConf.RowList[idxRow].CellList.Count; idxCell++)
                    {
                        var cellValue = tableConf.RowList[idxRow].CellList[idxCell];
                        // At first row , clones the cell
                        if (idxRow == 0)
                        {
                            // The first row already exists, just fill it
                            if (idxCell == 0)
                            {
                                baseRow.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableCell>().ElementAt(0).Descendants<Paragraph>().First().Descendants<Text>().First().Text = cellValue.Text;
                            }
                            else
                            {
                                // Adds cell
                                TableCell cell2 = (TableCell)baseRow.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableCell>().ElementAt(0).CloneNode(true);
                                cell2.Descendants<Paragraph>().First().Descendants<Text>().First().Text = cellValue.Text;
                                baseRow.AppendChild(cell2);
                            }
                        }
                        else 
                        {  
                            // working with cloned row
                            rowCopy.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableCell>().ElementAt(idxCell).Descendants<Paragraph>().First().Descendants<Text>().First().Text = cellValue.Text;
                        }                    
                    }

                    if (idxRow > 0)
                        detailsTable.AppendChild(rowCopy);
                }  
            
            
            }

            


        }

        /// <summary>
        /// Set Bullet configuaration
        /// </summary>        
        private void SetNumberingConfiguration()
        {
            Guid chunkId = Guid.NewGuid();
            NumberingDefinitionsPart numberingPart = wpDocument.MainDocumentPart.NumberingDefinitionsPart;

            Level level1 = new Level() { LevelIndex = 0, TemplateCode = "04090015" };
            StartNumberingValue startNumberingValue1 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat1 = new NumberingFormat() { Val = NumberFormatValues.UpperRoman };
            LevelText levelText1 = new LevelText() { Val = "%1." };
            LevelJustification levelJustification1 = new LevelJustification() { Val = LevelJustificationValues.Left };
            PreviousParagraphProperties ParagraphProperties1 = new PreviousParagraphProperties();
            Indentation indentation1 = new Indentation() { Start = "360", Hanging = "360" };
            NumberingSymbolRunProperties numberingSymbolRunProperties = new NumberingSymbolRunProperties();
            RunFonts runFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Arial", HighAnsi = "Arial" };
            numberingSymbolRunProperties.Append(runFonts);
            numberingSymbolRunProperties.Append(new Bold());
            ParagraphProperties1.Append(indentation1);
            level1.Append(startNumberingValue1);
            level1.Append(numberingFormat1);
            level1.Append(levelText1);
            level1.Append(levelJustification1);
            level1.Append(ParagraphProperties1);
            level1.Append(numberingSymbolRunProperties);

            Level level2 = new Level() { LevelIndex = 1, TemplateCode = "0409000F" };
            StartNumberingValue startNumberingValue2 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat2 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText2 = new LevelText() { Val = "%1.%2." };
            LevelJustification levelJustification2 = new LevelJustification() { Val = LevelJustificationValues.Left };
            PreviousParagraphProperties ParagraphProperties2 = new PreviousParagraphProperties();
            Indentation indentation2 = new Indentation() { Start = "720", Hanging = "360" };
            NumberingSymbolRunProperties numberingSymbolRunProperties2 = new NumberingSymbolRunProperties();
            RunFonts runFonts2 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Arial", HighAnsi = "Arial" };
            numberingSymbolRunProperties2.Append(runFonts2);
            numberingSymbolRunProperties2.Append(new Bold());
            ParagraphProperties2.Append(indentation2);
            level2.Append(startNumberingValue2);
            level2.Append(numberingFormat2);
            level2.Append(levelText2);
            level2.Append(levelJustification2);
            level2.Append(ParagraphProperties2);
            level2.Append(numberingSymbolRunProperties2);

            Level level3 = new Level() { LevelIndex = 2, TemplateCode = "E8080B64" };
            StartNumberingValue startNumberingValue3 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat3 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText3 = new LevelText() { Val = "%3." };
            LevelJustification levelJustification3 = new LevelJustification() { Val = LevelJustificationValues.Left };
            PreviousParagraphProperties previousParagraphProperties3 = new PreviousParagraphProperties();
            Indentation indentation3 = new Indentation() { Start = "1080", Hanging = "360" };
            previousParagraphProperties3.Append(indentation3);
            NumberingSymbolRunProperties numberingSymbolRunProperties3 = new NumberingSymbolRunProperties();
            RunFonts runFonts3 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Arial", HighAnsi = "Arial" };
            numberingSymbolRunProperties3.Append(runFonts3);
            numberingSymbolRunProperties3.Append(new Bold());
            level3.Append(startNumberingValue3);
            level3.Append(numberingFormat3);
            level3.Append(levelText3);
            level3.Append(levelJustification3);
            level3.Append(previousParagraphProperties3);
            level3.Append(numberingSymbolRunProperties3);


            Level level4 = new Level() { LevelIndex = 3, TemplateCode = "E8080B64" };
            StartNumberingValue startNumberingValue4 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat4 = new NumberingFormat() { Val = NumberFormatValues.None };
            LevelText levelText4 = new LevelText() { Val = "-" };
            LevelJustification levelJustification4 = new LevelJustification() { Val = LevelJustificationValues.Left };
            PreviousParagraphProperties previousParagraphProperties4 = new PreviousParagraphProperties();
            Indentation indentation4 = new Indentation() { Start = "1080", Hanging = "360" };
            previousParagraphProperties4.Append(indentation4);
            NumberingSymbolRunProperties numberingSymbolRunProperties4 = new NumberingSymbolRunProperties();
            RunFonts runFonts4 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Arial", HighAnsi = "Arial" };
            numberingSymbolRunProperties4.Append(runFonts4);
            numberingSymbolRunProperties4.Append(new Bold());
            level4.Append(startNumberingValue4);
            level4.Append(numberingFormat4);
            //level4.Append(levelText4);
            level4.Append(levelJustification4);
            level4.Append(previousParagraphProperties4);
            level4.Append(numberingSymbolRunProperties4);

            Numbering NUMelement = new Numbering(
                new AbstractNum(
                    level1,
                    level2,
                    level3,
                    level4
                )
                { AbstractNumberId = 1 },
                new NumberingInstance(
                  new AbstractNumId() { Val = 1 }
                )
                { NumberID = 1 }
            );

            NUMelement.Save(numberingPart);
        }

    }

    /// <summary>
    /// Bullets configuration 
    /// </summary>
    public class BulletsConfigWordGenerator
    {
        public string Text { get; set; }
        public int Lavel { get; set; }
        public bool IsUnderLine { get; set; }
        public bool Isbold { get; set; }

        public string MargingLeft {
            get {
                switch (Lavel) {
                    case 0: return "720";
                        break;
                    case 1:
                        return "820";
                        break;
                    case 2:
                    case 3:
                        return "920";
                        break;
                    default: return "720";
                };
            }
        }
        public string fontSize { get; set; }

        public BulletsConfigWordGenerator(string text, int lavel, bool isUnderLine, bool isBool, string fontSize) {
            this.Text = text;
            this.Lavel = lavel;
            this.IsUnderLine = isUnderLine;
            this.Isbold = isBool;
            this.fontSize = fontSize;
        }

    }

    /// <summary>
    /// Table , Rows, Cells an values configurations
    /// </summary>
    public class TableConfigWordGenerator
    {
        public List<RowConfigWordGenerator> RowList { get; set; }
    }

    public class RowConfigWordGenerator
    {
        public List<CellConfigWordGenerator> CellList { get; set; }
    }

    public class CellConfigWordGenerator
    {
        public string Text { get; set; }
    }


}
