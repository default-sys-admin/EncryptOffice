using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using NLog;

namespace purify.Data;

public class Processor_Word : IDocProcessor
{
    private static readonly Logger log = LogManager.GetCurrentClassLogger();

    private readonly Application app;

    public string TargetExt => ".docx";

    public Processor_Word()
    {
        app = new Application();
    }

    public void Dispose()
    {
        app.Quit();
        Marshal.ReleaseComObject(app);
    }

    public void ProcessFile(OneFileOptions opts)
    {
        var doc = opts.Password == null
                      ? app.Documents.Open(opts.SrcFile)
                      : app.Documents.Open(opts.SrcFile, PasswordDocument: opts.Password);

        log.Info("已成功打开文件，处理中……");
        doc.RemoveDocumentInformation(WdRemoveDocInfoType.wdRDIDocumentProperties);
        if (opts.RemovePassword)
            doc.Password = "";
        if (opts.NewPassword != null)
            doc.Password = opts.NewPassword;

        if (opts.OverwriteInplace)
        {
            log.Info($"原地覆盖文件：{opts.SrcFile}");
            doc.Save();
        }
        else
        {
            log.Info($"另存为文件：{opts.DstFile}");
            // 确保输出路径存在
            var dir = Path.GetDirectoryName(opts.DstFile);
            if (dir != null) Directory.CreateDirectory(dir);
            if (opts.OverwriteDst)
                File.Delete(opts.DstFile);
            doc.SaveAs2(opts.DstFile, WdSaveFormat.wdFormatDocumentDefault, CompatibilityMode: WdCompatibilityMode.wdCurrent);
        }

        doc.Close();
    }
}