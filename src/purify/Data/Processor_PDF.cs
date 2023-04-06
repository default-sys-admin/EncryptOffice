using System;
using System.IO;
using DevExpress.Pdf;
using NLog;

namespace purify.Data;

public class Processor_PDF : IDocProcessor
{
    private static readonly Logger log = LogManager.GetCurrentClassLogger();

    public Action<PdfPasswordRequestedEventArgs> GetPassword;

    private readonly PdfDocumentProcessor processor;
    private OneFileOptions cur_opts;

    public string TargetExt => ".pdf";

    public Processor_PDF()
    {
        processor = new PdfDocumentProcessor();
        processor.PasswordRequested += Processor_PasswordRequested;
    }

    public void Dispose()
    {
        processor.Dispose();
    }

    private void Processor_PasswordRequested(object sender, PdfPasswordRequestedEventArgs e)
    {
        if (cur_opts?.Password != null)
        {
            e.PasswordString = cur_opts.Password;
            return;
        }

        if (GetPassword != null)
        {
            GetPassword(e);
        }
        else
        {
            Console.WriteLine($"Iput password for {e.FileName}");
            Console.Write(": ");
            e.PasswordString = Console.ReadLine();
        }

        // 将用户输入的pwd保存到opts中
        if (cur_opts != null)
            cur_opts.Password = e.PasswordString;
    }

    public void ProcessFile(OneFileOptions opts)
    {
        cur_opts = opts;
        processor.LoadDocument(opts.SrcFile);

        log.Info("已成功打开文件，处理中……");
        var doc = processor.Document;
        doc.Author = null;
        doc.Producer = null;
        doc.Title = null;
        doc.Subject = null;
        doc.Creator = null;
        doc.Keywords = null;

        var saveOptions = new PdfSaveOptions();

        if (opts.RemovePassword)
        {
            saveOptions.EncryptionOptions = null;
        }

        if (opts.NewPassword != null)
        {
            saveOptions.EncryptionOptions = new PdfEncryptionOptions
                                            {
                                                Algorithm = PdfEncryptionAlgorithm.AES256,
                                                UserPasswordString = opts.NewPassword,
                                            };
        }
        else
        {
            // 既没有设置新密码，也没有移除老密码，则保留原有密码配置
            if (!opts.RemovePassword && opts.Password != null)
            {
                saveOptions.EncryptionOptions = new PdfEncryptionOptions
                                                {
                                                    Algorithm = PdfEncryptionAlgorithm.AES256,
                                                    UserPasswordString = opts.Password,
                                                };
            }
            else
                saveOptions.EncryptionOptions = null;
        }

        string dstFilename;
        if (opts.OverwriteInplace)
        {
            dstFilename = opts.SrcFile;
            log.Info($"原地覆盖文件：{dstFilename}");
        }
        else
        {
            dstFilename = opts.DstFile;
            log.Info($"另存为文件：{dstFilename}");
        }

        // 确保输出路径存在
        var dir = Path.GetDirectoryName(dstFilename);
        if (dir != null) Directory.CreateDirectory(dir);

        processor.SaveDocument(dstFilename, saveOptions);

        processor.CloseDocument();
    }
}