using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLWorldDocument
{
  class Program
  {
    static void Main(string[] args)
    {
      using (WordprocessingDocument wordDocument =
                  WordprocessingDocument.Create(@"C:\Users\viet.to\source\repos\OpenXMLWorldDocument\TestDocs.docx", WordprocessingDocumentType.Document))
      {
        // Add a main document part. 
        MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
        // Create the document structure and add some text.
        Body docBody = new Body();
        Paragraph p = new Paragraph();
        Paragraph pagebreak = new Paragraph(new Run(new Break(){ Type = BreakValues.Page }));
        p.Append(new Run(new Text("Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent quam augue, tempus id metus in, laoreet viverra quam. Sed vulputate risus lacus, et dapibus orci porttitor non.")));
        //// Add the table to the body
        WordTable wt = new WordTable();
        Table tbl = wt.createTable();
        //// Create a new row
        TableRow tr0 = new TableRow();
        //// Add a cell to each column in the row
        TableCell itemNo = new TableCell(new Paragraph(new Run(new Text("Item No"))));
        TableCell dateTime = new TableCell(new Paragraph(new Run(new Text("Date Time"))));
        TableCell itemDescription = new TableCell(new Paragraph(new Run(new Text("Item Description"))));
        TableCell claimed = new TableCell(new Paragraph(new Run(new Text("Claimed"))));
        TableCell offered = new TableCell(new Paragraph(new Run(new Text("Offered"))));
        TableCell agreed = new TableCell(new Paragraph(new Run(new Text("Agreed"))));
        TableCell vat = new TableCell(new Paragraph(new Run(new Text("VAT"))));
        TableCell customVAT = new TableCell(new Paragraph(new Run(new Text("Custom VAT"))));
        TableCell accepted = new TableCell(new Paragraph(new Run(new Text("Accepted"))));
        tr0.Append(itemNo, dateTime, itemDescription, claimed, offered, agreed, vat, customVAT, accepted);
        //// Create a new row
        TableRow tr = new TableRow();
        //// Add a cell to each column in the row
        TableCell tcName1 = new TableCell(new Paragraph(new Run(new Text("A.1.1"))));
        TableCell tcId1 = new TableCell(new Paragraph(new Run(new Text("20/02/2012 to TBC"))));
        TableCell tcId11 = new TableCell(new Paragraph(new Run(new Text("Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent quam augue, tempus id metus in, laoreet viverra quam. Sed vulputate risus lacus, et dapibus orci porttitor non."))));
        TableCell tcId12 = new TableCell(new Paragraph(new Run(new Text("1000"))));
        TableCell tcId13 = new TableCell(new Paragraph(new Run(new Text("1500"))));
        TableCell tcId14 = new TableCell(new Paragraph(new Run(new Text("1500"))));
        TableCell tcId15 = new TableCell(new Paragraph(new Run(new Text("23%"))));
        TableCell tcId16 = new TableCell(new Paragraph(new Run(new Text("0"))));
        TableCell tcId17 = new TableCell(new Paragraph(new Run(new Text("Yes"))));
        //// Add the cells to the row
        tr.Append(tcName1, tcId1, tcId11, tcId12, tcId13, tcId14, tcId15, tcId16, tcId17);
        // Create a new row
        TableRow tr1 = new TableRow();
        //// Add a cell to each column in the row
        TableCell tcName2 = new TableCell(new Paragraph(new Run(new Text("A.1.2"))));
        TableCell tcId2 = new TableCell(new Paragraph(new Run(new Text("20/02/2012 to 30/04/2013"))));
        TableCell tcId21 = new TableCell(new Paragraph(new Run(new Text("Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent quam augue, tempus id metus in, laoreet viverra quam. Sed vulputate risus lacus, et dapibus orci porttitor non."))));
        TableCell tcId22 = new TableCell(new Paragraph(new Run(new Text("500"))));
        TableCell tcId23 = new TableCell(new Paragraph(new Run(new Text("700"))));
        TableCell tcId24 = new TableCell(new Paragraph(new Run(new Text("500"))));
        TableCell tcId25 = new TableCell(new Paragraph(new Run(new Text("0%"))));
        TableCell tcId26 = new TableCell(new Paragraph(new Run(new Text("150"))));
        TableCell tcId27 = new TableCell(new Paragraph(new Run(new Text("Yes"))));
        //// Add the cells to the row
        tr1.Append(tcName2, tcId2, tcId21, tcId22, tcId23, tcId24, tcId25, tcId26, tcId27);
        TableRow tr2 = new TableRow();
        //// Add a cell to each column in the row
        TableCell tcName3 = new TableCell(new Paragraph(new Run(new Text("A.1.2"))));
        TableCell tcId3 = new TableCell(new Paragraph(new Run(new Text("09/12/2014"))));
        TableCell tcId31 = new TableCell(new Paragraph(new Run(new Text("Drawing Personal Injuries Summons, engrossing and serving same"))));
        TableCell tcId32 = new TableCell(new Paragraph(new Run(new Text("75"))));
        TableCell tcId33 = new TableCell(new Paragraph(new Run(new Text("0"))));
        TableCell tcId34 = new TableCell(new Paragraph(new Run(new Text("0"))));
        TableCell tcId35 = new TableCell(new Paragraph(new Run(new Text("23%"))));
        TableCell tcId36 = new TableCell(new Paragraph(new Run(new Text("0"))));
        TableCell tcId37 = new TableCell(new Paragraph(new Run(new Text("No"))));
        //Add the cells to the row
        tr2.Append(tcName3, tcId3, tcId31, tcId32, tcId33, tcId34, tcId35, tcId36, tcId37);
        TableRow tr3 = new TableRow();
        //Add a cell to each column in the row
        TableCell tcName4= new TableCell(new Paragraph(new Run(new Text("A.1.2"))));
        //Create table cell properties
        TableCellProperties tcp = new TableCellProperties();
        //Create grid span
        GridSpan gs = new GridSpan(){Val = 8};
        //Add the grid span to the Table Cell property
        tcp.Append(gs);
        //Add table cell property to table cell
        tcName4.Append(tcp);
        TableCell tcId47 = new TableCell(new Paragraph(new Run(new Text("Total"))));
        //Add the cells to the row
        tr3.Append(tcName4, tcId47);

        var tr4 = new TableRow();
        //// Add a cell to each column in the row
        var tcName5 = new TableCell(new TableCellProperties(){TableCellVerticalAlignment = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }}, new ParagraphProperties( new Justification() {Val = JustificationValues.Center}),new Paragraph(new Run(new Text("Item No")){RunProperties = new RunProperties(){ Bold = new Bold() }}));
        var tcId5 = new TableCell(new TableCellProperties(){TableCellVerticalAlignment = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }}, new ParagraphProperties( new Justification() {Val = JustificationValues.Center}),new Paragraph(new Run(new Text("Date"))));
        var tcId51 = new TableCell(new TableCellProperties(){TableCellVerticalAlignment = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }}, new ParagraphProperties( new Justification() {Val = JustificationValues.Center}), new Paragraph(new Run(new Text("Work Done"))));
        var tcId52 = new TableCell(new TableCellProperties(){TableCellVerticalAlignment = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }}, new ParagraphProperties( new Justification() {Val = JustificationValues.Center}),new Paragraph(new Run(new Text("Claimed"))));
        var tcId53 = new TableCell(new TableCellProperties(){TableCellVerticalAlignment = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }}, new ParagraphProperties( new Justification() {Val = JustificationValues.Center}),new Paragraph(new Run(new Text("For Legal Costs Adjudicator's use only"))));
        //Add the cells to the row
        tr4.Append(tcName5, tcId5, tcId51, tcId52, tcId53);

        var tr5 = new TableRow();
        //// Add a cell to each column in the row
        var tcName6 = new TableCell(new Paragraph(new Run(new Text("D.1.1"))));
        var tcId6 = new TableCell(new Paragraph(new Run(new Text("17 January 2020 to Completion"))));
        var tcId61 = new TableCell(new Paragraph(new Run(new Text("Corresponding with Client by way of update in order to advise him of the terms of the final settlement and Order. Bespeaking a copy of the final Order and discharging the relevant Stamp Duty in that regard.Collating all of the Fee Notes and invoices and upon receipt of the settlement cheque, discharging the remaining outlays and accounting to Client therefor.Conducting a review of the paperwork, which underpinned the costs claim to include, inter-alia, Counsel's briefs, the Solicitor's working files (2 volumes), copy Orders, vouchers and fee notes and thereafter, preparing a detailed / statutory Bill of Costs for settlement and / or adjudication.Preparation for and attending Adjudication, completing bill and vouching and extracting Certificate of Adjudication.For all instructions and work not specifically provided for by way of Appendix W charge, having regard to LSRA 2015 and Order 99 of the RSC."))));
        var tcId62 = new TableCell(new ParagraphProperties( new Justification() {Val = JustificationValues.Right}), new Paragraph(new Run(new Text("€3,000.00"))));
        var tcId63 = new TableCell(new Paragraph(new Run(new Text(String.Empty))));
        var tcp2 = new TableCellProperties();
        TableCellVerticalAlignment tcVA2 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Bottom };
        tcp2.Append(tcVA2);
        tcId62.Append(tcp2);
        //Add the cells to the row
        tr5.Append(tcName6, tcId6, tcId61, tcId62, tcId63);

        Table tbl2 = wt.createTable();
        //Add the rows to the table
        tbl.AppendChild(tr0);
        tbl.AppendChild(tr);
        tbl.AppendChild(tr1);
        tbl.AppendChild(tr2);
        tbl.AppendChild(tr3);
        tbl2.AppendChild(tr4);
        tbl2.AppendChild(tr5);
        //Add the table to the body
        docBody.Append(tbl);
        docBody.AppendChild(pagebreak);
        docBody.AppendChild(tbl2);
        docBody.AppendChild(p);
        mainPart.Document = new Document(docBody);
      }
    }
  }
}