using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using NLog;

namespace purify.Data;

public class Processor_PPT : IDocProcessor
{
    private static readonly Logger log = LogManager.GetCurrentClassLogger();

    private readonly Application app;

    public string TargetExt => ".pptx";

    public Processor_PPT()
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
                      ? app.Presentations.Open(opts.SrcFile)
                      : app.Presentations.Open(opts.SrcFile + $"::{opts.Password}::");

        log.Info("已成功打开文件，处理中……");
        doc.RemoveDocumentInformation(PpRemoveDocInfoType.ppRDIDocumentProperties);
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
            doc.SaveAs(opts.DstFile);
        }

        doc.Close();
    }
}