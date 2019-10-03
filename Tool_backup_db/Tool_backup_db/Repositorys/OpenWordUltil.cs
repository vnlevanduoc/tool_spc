using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Backup_Db_ATD.Repositorys
{
    public class OpenWordUltil
    {
        public void SetContentControlText(string document, Dictionary<string, string> values)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
            {
                if (document == null) throw new ArgumentNullException("document");
                if (values == null) throw new ArgumentNullException("values");

                foreach (SdtElement sdtElement in wordDoc.MainDocumentPart.Document.Descendants<SdtElement>())
                {
                    string tag = sdtElement.SdtProperties.Descendants<Tag>().First().Val;
                    if ((tag != null) && values.ContainsKey(tag))
                    {
                        sdtElement.Descendants<Text>().First().Text = values[tag] ?? string.Empty;
                        sdtElement.Descendants<Text>().Skip(1).ToList().ForEach(t => t.Remove());
                    }
                }
            }
        }

        public void EditTable(string document, int index_table, int start_rows, DataTable data)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
            {
                Body my_body = wordDoc.MainDocumentPart.Document.Body;
                Table my_table = my_body.Descendants<Table>().ToList()[index_table];

                TableProperties tblProp = new TableProperties(
                    new TableBorders(
                        new TopBorder()
                        { Val = new EnumValue<BorderValues>(BorderValues.BasicThinLines), Size = 1 },
                        new BottomBorder()
                        { Val = new EnumValue<BorderValues>(BorderValues.BasicThinLines), Size = 1 },
                        new LeftBorder()
                        { Val = new EnumValue<BorderValues>(BorderValues.BasicThinLines), Size = 1 },
                        new RightBorder()
                        { Val = new EnumValue<BorderValues>(BorderValues.BasicThinLines), Size = 1 },
                        new InsideHorizontalBorder()
                        { Val = new EnumValue<BorderValues>(BorderValues.BasicThinLines), Size = 1 },
                        new InsideVerticalBorder()
                        { Val = new EnumValue<BorderValues>(BorderValues.BasicThinLines), Size = 1 })
                );
                my_table.AppendChild<TableProperties>(tblProp);

                for (int i = 0; i < data.Rows.Count; i++)
                {
                    TableRow tr = new TableRow();
                    foreach (DataColumn dc in data.Columns)
                    {
                        TableCell my_table_cell = new TableCell(
                            new TableCellProperties(
                                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }
                            ),
                            new Paragraph(
                                new ParagraphProperties(
                                    new Justification() { Val = JustificationValues.Center }
                                ),
                                new Run(
                                    new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman" },
                                    new RunProperties(
                                        new FontSize { Val = "26" /*13px = 26*/ }
                                    ),
                                    new Text(data.Rows[i][dc].ToString())
                                )
                            )
                        );
                        tr.Append(my_table_cell);
                    }
                    my_table.Append(tr);
                }
            }
        }
    }
}
