using System;
using System.IO;
using DevExpress.Pdf;
using Microsoft.Office.Interop.Word;

namespace purify;

internal class Tests
{
    public static void TestPdf()
    {
        //About.ShowAbout(ProductKind.DXperienceUni);

        //var org = @"D:\tmp\1.pdf";
        var enc = @"D:\tmp\1-enc.pdf";
        var org2 = @"D:\tmp\1-restore.pdf";

        using var processor = new PdfDocumentProcessor();
        processor.PasswordRequested += Processor_PasswordRequested;
        processor.PasswordAttemptsLimit = 100;
        try
        {
            processor.LoadDocument(enc);
            Console.WriteLine(processor.Document.Producer);
            processor.SaveDocument(org2, new PdfSaveOptions() {EncryptionOptions = null});
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex);
        }

        //processor.LoadDocument(org);

        //var opt = new PdfEncryptionOptions
        //{
        //    UserPasswordString = "123456",
        //    Algorithm = PdfEncryptionAlgorithm.AES256
        //};

        //processor.SaveDocument(enc, new PdfSaveOptions {EncryptionOptions = opt});
    }

    public static void Processor_PasswordRequested(object sender, PdfPasswordRequestedEventArgs e)
    {
        Console.WriteLine($"input passwd for {e.FileName}:");
        //e.PasswordString = "1234567";
        e.PasswordString = Console.ReadLine();
    }

    public static void TestDocx()
    {
        var fname = @"D:\1.docx";
        var app = new Application();
        var doc = app.Documents.Open(fname);
        //doc.SaveAs2(@"D:\2.docx", WdSaveFormat.wdFormatDocument);
        doc.SaveAs2(@"D:\3.docx", WdSaveFormat.wdFormatDocumentDefault,
                    CompatibilityMode: WdCompatibilityMode.wdCurrent);
        doc.Close();
        app.Quit();
    }

    public static void TestPaths()
    {
        Console.WriteLine(Path.GetFullPath("d:"));
        Console.WriteLine(Path.GetFullPath("d:\\"));
        Console.WriteLine(Path.GetFullPath("d:\\ABC"));
        Console.WriteLine(Path.GetFullPath("d:\\ABC\\"));
        Console.WriteLine(Path.GetDirectoryName("d:"));
        Console.WriteLine(Path.GetDirectoryName("d:\\"));
        Console.WriteLine(Path.GetDirectoryName("d:\\ABC"));
        Console.WriteLine(Path.GetDirectoryName("d:\\ABC\\..\\cde\\ddd\\..\\last"));
        Console.WriteLine(Path.GetFullPath("d:\\ABC\\..\\cde\\ddd\\..\\last"));
        Console.WriteLine(Path.Combine("D:", "dir1"));
        Console.WriteLine(Path.Combine("D:\\", "dir1"));
        Console.WriteLine("finish");
    }
}