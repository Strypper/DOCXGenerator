using System;
using System.IO;
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
        p.Append(new Run(new Text("Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent quam augue, tempus id metus in, laoreet viverra quam. Sed vulputate risus lacus, et dapibus orci porttitor non.")));
        //// Add the table to the body
        Table tbl = new Table();
        // Set the style and width for the table.
        TableProperties tableProp = new TableProperties();
        TableStyle tableStyle = new TableStyle() { Val = "TableGrid" };
        //Table Width
        TableWidth tableWidth = new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct };
        // Apply Style
        tableProp.Append(tableStyle, tableWidth);
        tbl.AppendChild(tableProp);
        //// Create the table properties
        TableProperties tblProperties = new TableProperties();
        //// Create Table Borders
        TableBorders tblBorders = new TableBorders();
        TopBorder topBorder = new TopBorder();
        topBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
        topBorder.Color = "Black";
        tblBorders.AppendChild(topBorder);
        BottomBorder bottomBorder = new BottomBorder();
        bottomBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
        bottomBorder.Color = "Black";
        tblBorders.AppendChild(bottomBorder);
        RightBorder rightBorder = new RightBorder();
        rightBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
        rightBorder.Color = "Black";
        tblBorders.AppendChild(rightBorder);
        LeftBorder leftBorder = new LeftBorder();
        leftBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
        leftBorder.Color = "Black";
        tblBorders.AppendChild(leftBorder);
        InsideHorizontalBorder insideHBorder = new InsideHorizontalBorder();
        insideHBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
        insideHBorder.Color = "Black";
        tblBorders.AppendChild(insideHBorder);
        InsideVerticalBorder insideVBorder = new InsideVerticalBorder();
        insideVBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);
        insideVBorder.Color = "Red";
        tblBorders.AppendChild(insideVBorder);
        //// Add the table borders to the properties
        tblProperties.AppendChild(tblBorders);
        //// Add the table properties to the table
        tbl.AppendChild(tblProperties);
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
        //Add the rows to the table
        tbl.AppendChild(tr0);
        tbl.AppendChild(tr);
        tbl.AppendChild(tr1);
        tbl.AppendChild(tr2);
        tbl.AppendChild(tr3);
        //Add the table to the body
        docBody.Append(tbl);
        docBody.AppendChild(p);
        mainPart.Document = new Document(docBody);
      }
    }
  }
}