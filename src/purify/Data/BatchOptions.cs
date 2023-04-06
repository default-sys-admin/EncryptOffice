using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using DevExpress.Pdf;
using NLog;

namespace purify.Data;

public class BatchOptions
{
    private static readonly Logger log = LogManager.GetCurrentClassLogger();

    public BatchOptions()
    {
        PasswordsTxt = "19890604\r\n19970219";
        ApplyNewPasswordCheck();
    }

    /// <summary>
    /// 返回绝对路径，并始终保持末尾有反斜杠，以确保路径替换时不会出错
    /// </summary>
    string GetNormalizedDir(string dir)
    {
        if (Path.GetDirectoryName(dir) == null)
        {
            // 根目录需要特殊处理，否则 GetFullPath("D:") 这种情况会返回 D:\src\code\bin\debug\win64 这种结果
            return Path.EndsInDirectorySeparator(dir) ? dir : dir + Path.DirectorySeparatorChar;
        }

        var fullpath = Path.GetFullPath(dir);
        return Path.EndsInDirectorySeparator(fullpath) ? fullpath : fullpath + Path.DirectorySeparatorChar;
    }

    // 源目录
    public virtual string InputFolder { get; set; } = @"D:\data";
    public string InputFolderNorm => GetNormalizedDir(InputFolder);

    public void OnInputFolderChanged(string oldValue)
    {
        // 自动设置输出目录，特殊处理根目录的情况
        if (Path.GetDirectoryName(InputFolderNorm) == null)
            OutputFolder = Path.Combine(InputFolderNorm, "anonymous");
        else
            OutputFolder = InputFolderNorm.TrimEnd('\\') + "-anonymous";
    }

    // 输出目录
    public virtual string OutputFolder { get; set; } = @"D:\data-anonymous";
    public string OutputFolderNorm => GetNormalizedDir(OutputFolder);

    // 是否覆盖输出目录中的同名文件
    public virtual bool OverwriteOutputFile { get; set; } = false;

    // 要转换的类型
    public virtual bool CnvWord { get; set; } = true;
    public virtual bool CnvExcel { get; set; } = true;
    public virtual bool CnvPowerPoint { get; set; } = true;
    public virtual bool CnvPDF { get; set; } = true;

    // 密码本
    public virtual string PasswordsTxt { get; set; }

    public void OnPasswordsTxtChanged(string oldValue)
    {
        try
        {
            if (string.IsNullOrEmpty(PasswordsTxt))
            {
                // 默认无密码本
                Passwords = Array.Empty<string>();
                return;
            }

            Passwords = PasswordsTxt.Split('\r', '\n').Where(s => !string.IsNullOrEmpty(s)).ToArray();
            IsPasswordTxtValid = true;
        }
        catch (Exception ex)
        {
            log.Error(ex);

            // 出错了，用null表示
            Passwords = null;
            IsPasswordTxtValid = false;
        }
    }

    public virtual string[] Passwords { get; set; }
    public virtual bool IsPasswordTxtValid { get; set; }

    // 是否要去除已有的密码
    public virtual bool RemovePassword { get; set; }

    // 对原先没有密码的，是否自动加密
    public virtual bool AddPassword { get; set; }

    public void OnAddPasswordChanged(bool oldValue)
    {
        ApplyNewPasswordCheck();
    }

    public virtual string NewPassword { get; set; }

    public void OnNewPasswordChanged(string oldValue)
    {
        ApplyNewPasswordCheck();
    }

    public virtual bool IsNewPasswordValid { get; set; }

    private void ApplyNewPasswordCheck()
    {
        IsNewPasswordValid = !(AddPassword && string.IsNullOrEmpty(NewPassword));
    }

    // 避免线程重入
    public virtual bool IsRunning { get; set; }
    public virtual string RunBtnTxt { get; set; } = "Run!";
    public virtual bool StopButtonIsVisible { get; set; }
    public virtual bool IsStoppedByUser { get; set; }

    // log信息
    public virtual LogLevel CurrentLogLevel { get; set; } = LogLevel.Info;

    // 实际显示的log
    public virtual ObservableCollection<LogEventInfo> Logs { get; } = new();

    // 保存全部log，更改显示级别时重刷
    private readonly List<LogEventInfo> all_logs = new();

    public void OnCurrentLogLevelChanged(LogLevel oldValue)
    {
        Logs.Clear();
        foreach (var evt in all_logs.Where(l => l.Level >= CurrentLogLevel))
        {
            Logs.Add(evt);
        }
    }

    public void AppendLog(LogEventInfo evt)
    {
        all_logs.Add(evt);
        if (evt.Level >= CurrentLogLevel) Logs.Add(evt);

        // 超过1000条时干掉前100条
        if (all_logs.Count > 1000)
        {
            var preserved = all_logs.Skip(100).ToArray();
            all_logs.Clear();
            all_logs.AddRange(preserved);
        }

        while (Logs.Count > 1000)
        {
            Logs.RemoveAt(0);
        }
    }

    /// <summary>
    /// 实际执行操作
    /// </summary>
    public void Execute(Action<PdfPasswordRequestedEventArgs> pdfPwdCallback = null)
    {
        // 按类型逐个处理
        try
        {
            if (CnvWord)
            {
                log.Info("============ 处理Word文件 ================");
                var app = new Processor_Word();
                var exts = new[] {".doc", ".docx"};
                ProcessFileType(app, exts);
                app.Dispose();
            }

            if (!IsStoppedByUser)
            {
                if (CnvExcel)
                {
                    log.Info("============ 处理Excel文件 ================");
                    var app = new Processor_Excel();
                    var exts = new[] {".xls", ".xlsx"};
                    ProcessFileType(app, exts);
                    app.Dispose();
                }
            }

            if (!IsStoppedByUser)
            {
                if (CnvPowerPoint)
                {
                    log.Info("============ 处理PPT文件 ================");
                    var app = new Processor_PPT();
                    var exts = new[] {".ppt", ".pptx"};
                    ProcessFileType(app, exts);
                    app.Dispose();
                }
            }

            if (!IsStoppedByUser)
            {
                if (CnvPDF)
                {
                    log.Info("============ 处理PDF文件 ================");
                    var app = new Processor_PDF();
                    app.GetPassword = pdfPwdCallback; // PDF输入密码的专用回调函数
                    var exts = new[] {".pdf"};
                    ProcessFileType(app, exts);
                    app.Dispose();
                }
            }
        }
        catch (Exception ex)
        {
            log.Error(ex);
        }

        log.Info("Finished");
        IsRunning = false;
        RunBtnTxt = "Run!";
        StopButtonIsVisible = false;
        IsStoppedByUser = false;
    }

    /// <summary>
    /// 根据输入的文件名和model的配置，得到输出的文件名
    /// 例如，输入 D:\doc1\sub1\sub2\abc.doc
    /// model中的输入路径为 D:\doc1
    /// 输出路径为 D:\doc1\updated
    /// 则最终输出的文件名为 D:\doc1\updated\sub1\abc.docx 注意扩展名也要更新
    /// </summary>
    protected string CalcOutputFilename(string filename, string ext)
    {
        var fullname = Path.GetFullPath(filename);
        if (!fullname.StartsWith(InputFolderNorm))
            throw new Exception($"path error: {filename} not in {InputFolderNorm}");

        var outfile = fullname.Replace(InputFolderNorm, OutputFolderNorm);
        return Path.ChangeExtension(outfile, ext);
    }

    private void ProcessFileType(IDocProcessor app, string[] exts)
    {
        var filenames = new List<string>();

        foreach (var ext in exts)
        {
            var lst = Directory.EnumerateFiles(InputFolderNorm, "*" + ext, SearchOption.AllDirectories).ToArray();
            log.Info($"找到{lst.Length}个{ext}文件");
            filenames.AddRange(lst);
        }

        var cnt = 1;
        foreach (var filename in filenames)
        {
            // 检查是否提前停止
            if (IsStoppedByUser)
            {
                log.Warn("stopped by user");
                break;
            }

            log.Info($"处理文件[{cnt} of {filenames.Count}]: {filename}");
            cnt++;
            if (Path.GetFileName(filename).StartsWith("~$"))
            {
                log.Warn($"跳过残留的临时文件[{filename}]");
                continue;
            }

            try
            {
                // 开始处理单个文件
                ProcessOneFile(app, filename);
            }
            catch (Exception ex)
            {
                log.Error(ex);
            }
        }
    }

    private void ProcessOneFile(IDocProcessor app, string filename)
    {
        var opts = new OneFileOptions
                   {
                       SrcFile = filename,
                       DstFile = CalcOutputFilename(filename, app.TargetExt),
                       NewPassword = NewPassword,
                       OverwriteDst = OverwriteOutputFile,
                       RemovePassword = RemovePassword,
                   };
        opts.OverwriteInplace = Path.GetFullPath(filename).ToLowerInvariant() == Path.GetFullPath(opts.DstFile).ToLowerInvariant()
                                && OverwriteOutputFile;

        if (File.Exists(opts.DstFile) && !OverwriteOutputFile)
        {
            log.Error($"{opts.DstFile} 已存在，跳过");
            return;
        }

        // 先尝试文件名和密码本中的密码
        var pwd_candidates = Passwords.ToList();
        var bare_name = Path.GetFileNameWithoutExtension(filename);
        var ss = bare_name.Split('-');
        if (ss.Length > 1)
        {
            var pwd = ss[^1];
            if (!string.IsNullOrEmpty(pwd))
            {
                pwd_candidates.Insert(0, pwd);
            }
        }

        var success = false;
        foreach (var pwd in pwd_candidates)
        {
            try
            {
                opts.Password = pwd;
                app.ProcessFile(opts);
                success = true;
                break;
            }
            catch (Exception)
            {
                // 密码错误
            }
        }

        if (!success)
        {
            // 用户输入密码
            try
            {
                opts.Password = null;
                app.ProcessFile(opts);
                success = true;
            }
            catch (Exception)
            {
                // 密码错误
            }
        }

        if (!success)
        {
            log.Error($"无法处理文件：{filename}");
        }
    }
}