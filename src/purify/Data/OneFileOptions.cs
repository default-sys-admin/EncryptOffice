using System.IO;
using CommandLine;
using NLog;

namespace purify.Data;

public class OneFileOptions
{
    private static readonly Logger log = LogManager.GetCurrentClassLogger();

    [Option("src", Required = true, HelpText = "源文件")]
    public string SrcFile { get; set; }

    [Option("inplace", HelpText = "覆盖源文件，启用此选项时，不能设置目标文件")]
    public bool OverwriteInplace { get; set; }

    [Option("dst", HelpText = "目标文件")]
    public string DstFile { get; set; }

    [Option("overwrite", HelpText = "如果目标文件已存在，是否覆盖")]
    public bool OverwriteDst { get; set; }

    [Option("pwd", HelpText = "源文件的密码")]
    public string Password { get; set; }

    [Option("new-pwd", HelpText = "新的密码，如果不设置新密码，且不移除原有密码，则保持密码不变")]
    public string NewPassword { get; set; }

    [Option("remove-pwd", HelpText = "移除密码，取消加密")]
    public bool RemovePassword { get; set; }

    public bool Validate()
    {
        if (!File.Exists(SrcFile))
        {
            log.Error($"--src [{SrcFile}] 不存在");
            return false;
        }

        if (DstFile == null && !OverwriteInplace)
        {
            log.Error("--inplace 和 --dst必须设置其中一个");
            return false;
        }

        if (OverwriteInplace && DstFile != null)
        {
            log.Error("--inplace 和 --dst 不能同时设置");
            return false;
        }

        if (File.Exists(DstFile) && !OverwriteDst)
        {
            log.Error($"目标文件 {DstFile} 已存在，如要覆盖，需要设置 --overwrite");
            return false;
        }

        if (RemovePassword && NewPassword != null)
        {
            log.Error("--remove-pwd 和 --new-pwd 不能同时设置");
            return false;
        }

        return true;
    }
}