using System;
using System.Collections.Generic;
using System.IO;
using CommandLine;
using CommandLine.Text;
using NLog;
using purify.Data;

namespace purify;

internal class Program
{
    public const string Version = "0.7.2";

    private static readonly Logger log = LogManager.GetCurrentClassLogger();

    [STAThread]
    public static void Main(string[] args)
    {
        if (args.Length == 0)
        {
            // 无参数，提示选择运行模式
            Console.WriteLine("=============================");
            Console.WriteLine("[向导模式] purify guide");
            Console.WriteLine("[参数模式] purify --help");
            Console.WriteLine("=============================");
            RunInCommandLineMode(new[] { "--help" });
        }
        else if (args.Length == 1 && args[0] == "guide")
        {
            // 进入向导模式
            new Guided().Run();
        }
        else
        {
            // 命令行参数模式
            RunInCommandLineMode(args);
        }
    }

    private static void RunInCommandLineMode(string[] args)
    {
        var parser = new Parser(with => with.HelpWriter = null);
        var result = parser.ParseArguments<OneFileOptions>(args);
        result.WithParsed(RunOneFile)
              .WithNotParsed(errs => ShowErrors(result, errs));
    }

    private static void ShowErrors(ParserResult<OneFileOptions> result, IEnumerable<Error> errors)
    {
        var helpText = HelpText.AutoBuild(
            result,
            h =>
            {
                h.AdditionalNewLineAfterOption = false;
                h.Heading = $"purify v{Version}";
                h.Copyright = "参数说明：";
                h.AutoVersion = false;
                h.MaximumDisplayWidth = 120;
                return HelpText.DefaultParsingErrorsHandler(result, h);
            }, e => e);
        Console.WriteLine(helpText);

        Console.WriteLine(@"
示例:

  加密:
    purify --src D:\docs\1.docx --inplace --new-pwd asdf

  清除密码并另存为新文件:
    purify --src D:\docs\1.docx --dst D:\docs\2.docx --pwd asdf --remove-pwd

  修改密码，另存为新文件，如果目标文件已存在则覆盖
    purify --src D:\docs\1.docx --dst D:\docs\2.docx --pwd asdf --new-pwd thisisnewpwd --overwrite

  直接以向导模式启动:
    purify guide

"
        );
    }

    /// <summary>
    /// 单独处理一个文件
    /// </summary>
    public static void RunOneFile(OneFileOptions opts)
    {
        try
        {
            // convert to absolute fullpath
            opts.SrcFile = Path.GetFullPath(opts.SrcFile);
            if (!opts.OverwriteInplace)
                opts.DstFile = Path.GetFullPath(opts.DstFile);

            if (!opts.Validate())
                return;

            IDocProcessor p;
            var ext = Path.GetExtension(opts.SrcFile)?.ToLowerInvariant();
            if (ext == ".doc" || ext == ".docx")
                p = new Processor_Word();
            else if (ext == ".xls" || ext == ".xlsx")
                p = new Processor_Excel();
            else if (ext == ".ppt" || ext == ".pptx")
                p = new Processor_PPT();
            else if (ext == ".pdf")
                p = new Processor_PDF();
            else
            {
                log.Error($"无法处理的文件类型 [{ext}]");
                return;
            }

            log.Info($"处理文件: {opts.SrcFile}");
            p.ProcessFile(opts);
            p.Dispose();
        }
        catch (Exception ex)
        {
            var save = Console.ForegroundColor;
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(ex);
            Console.ForegroundColor = save;
        }
    }
}