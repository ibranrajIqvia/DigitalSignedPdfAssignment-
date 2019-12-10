using SautinSoft.Document;
using SautinSoft.Document.Drawing;
using System.Linq;
using System;

namespace ConsoleApp1
{
    class Program
    {
        static string blankdoc = @"..\..\BlankDoc.docx";
        static string sign = @"..\..\sing.jpg";
        static string resultpdf = "Result.pdf";
        static void Main(string[] args)
        {
            Console.WriteLine("Processing...............");
            
            DocumentCore dcObj = DocumentCore.Load(blankdoc);


            Shape signatureShape = new Shape(dcObj, Layout.Floating(new HorizontalPosition(0f, LengthUnit.Millimeter, HorizontalPositionAnchor.LeftMargin),
                                   new VerticalPosition(0f, LengthUnit.Millimeter, VerticalPositionAnchor.TopMargin), new Size(3, 3)));
            ((FloatingLayout)signatureShape.Layout).WrappingStyle = WrappingStyle.InFrontOfText;
            signatureShape.Outline.Fill.SetEmpty();

            Paragraph firstPar = dcObj.GetChildElements(true).OfType<Paragraph>().FirstOrDefault();
            firstPar.Inlines.Add(signatureShape);


            Picture signaturePict = new Picture(dcObj, sign);

            signaturePict.Layout = Layout.Floating(
               new HorizontalPosition(2.5, LengthUnit.Centimeter, HorizontalPositionAnchor.Page),
               new VerticalPosition(4.5, LengthUnit.Centimeter, VerticalPositionAnchor.Page),
               new Size(10, 5, LengthUnit.Centimeter));

            PdfSaveOptions options = new PdfSaveOptions();


            options.DigitalSignature.CertificatePath = @"..\..\sautinsoft.pfx";
            options.DigitalSignature.CertificatePassword = "123456789";

           
            options.DigitalSignature.SignatureLine = signatureShape;           
            options.DigitalSignature.Signature = signaturePict;


            dcObj.Save(resultpdf, options);           
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(resultpdf) { UseShellExecute = true });

        }
    }
}
