using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NLog;
using purify.Data;

namespace purify;

internal class Guided
{
    private static readonly Logger log = LogManager.GetCurrentClassLogger();

    private readonly BatchOptions batchOpts = new();

    bool GetYesNoFromKeyboard(string prompt)
    {
        while (true)
        {
            Console.Write(prompt);
            var c = Console.ReadKey();
            Console.WriteLine();
            switch (c.KeyChar)
            {
            case 'y':
            case 'Y':
                return true;
            case 'n':
            case 'N':
                return false;
            default:
                // 输入错误，重新输入y/n
                continue;
            }
        }
    }

    public void Run()
    {
        Console.WriteLine("=====================================");
        Console.WriteLine("||             向导模式            ||");
        Console.WriteLine("=====================================");

        while (true)
        {
            Console.Write("源目录:");
            batchOpts.InputFolder = Console.ReadLine();
            if (Directory.Exists(batchOpts.InputFolder))
            {
                Console.WriteLine($"[源目录] {batchOpts.InputFolderNorm}");
                break;
            }

            Console.WriteLine("Error: 源目录不存在");
        }

        while (true)
        {
            Console.Write("输出目录:");
            batchOpts.OutputFolder = Console.ReadLine();
            Console.WriteLine($"[输出目录] {batchOpts.OutputFolderNorm}");
            if (Directory.Exists(batchOpts.OutputFolder))
            {
                batchOpts.OverwriteOutputFile = GetYesNoFromKeyboard("目标目录已存在，是否覆盖目录中的同名文件？(Y/N):");

                // 确认覆盖，否则重新选择输出目录
                if (batchOpts.OverwriteOutputFile)
                    break;
            }
            else
            {
                break;
            }
        }

        while (true)
        {
            // reset flags
            batchOpts.CnvWord = batchOpts.CnvExcel = batchOpts.CnvPowerPoint = batchOpts.CnvPDF = false;

            Console.Write("需要处理的文件类型（1-word 2-excel 3-ppt 4-pdf，可以多选，如134）:");
            var s = Console.ReadLine() + "";
            var is_valid = true;
            foreach (var c in s)
            {
                if (c == '1') batchOpts.CnvWord = true;
                else if (c == '2') batchOpts.CnvExcel = true;
                else if (c == '3') batchOpts.CnvPowerPoint = true;
                else if (c == '4') batchOpts.CnvPDF = true;
                else
                {
                    // 输入错误
                    Console.WriteLine($"错误输入：{c}");
                    is_valid = false;
                    break;
                }
            }

            if (is_valid)
                break;
        }

        // 输入密码本
        var pwd_list = new List<string>();
        Console.WriteLine("输入密码本，一行一个密码，直接回车结束输入:");
        while (true)
        {
            var pwd = Console.ReadLine();
            if (string.IsNullOrEmpty(pwd))
            {
                batchOpts.Passwords = pwd_list.ToArray();
                break;
            }

            pwd_list.Add(pwd);
        }

        // 设置新密码
        Console.Write("设置新密码，直接回车表示不设置新密码:");
        batchOpts.NewPassword = Console.ReadLine();
        batchOpts.AddPassword = !string.IsNullOrEmpty(batchOpts.NewPassword);

        // 如果设置了新密码，那么自动去除老密码，否则需要选择
        batchOpts.RemovePassword = batchOpts.AddPassword || GetYesNoFromKeyboard("是否去除已有的密码(Y/N):");

        // 设置完毕，确认执行
        Console.WriteLine("=====================================");
        Console.WriteLine("||                                 ||");
        Console.WriteLine("||        设置完毕，检查选项       ||");
        Console.WriteLine("||                                 ||");
        Console.WriteLine("=====================================");
        Console.WriteLine($"      [源目录] {batchOpts.InputFolderNorm}");
        Console.WriteLine($"    [输出目录] {batchOpts.OutputFolderNorm}");
        Console.WriteLine($"    [自动覆盖] {batchOpts.OverwriteOutputFile}");
        Console.WriteLine($"    [文件类型] "
                          + (batchOpts.CnvWord ? "word " : "")
                          + (batchOpts.CnvExcel ? "excel " : "")
                          + (batchOpts.CnvPowerPoint ? "ppt " : "")
                          + (batchOpts.CnvPDF ? "pdf " : ""));
        Console.WriteLine($"      [密码本] {batchOpts.Passwords.FirstOrDefault()}");
        foreach (var pwd in batchOpts.Passwords.Skip(1)) Console.WriteLine("               " + pwd);
        Console.WriteLine($"[去除已有密码] {batchOpts.RemovePassword}");
        Console.WriteLine($"      [新密码] {(batchOpts.AddPassword ? batchOpts.NewPassword : string.Empty)}");

        var confirm = GetYesNoFromKeyboard("确认执行？(Y/N):");
        if (confirm)
        {
            // 准备Ctrl+C中断回调，避免中间状态
            Console.CancelKeyPress += OnKeyboardBreak;
            batchOpts.Execute();
        }
    }

    private void OnKeyboardBreak(object sender, ConsoleCancelEventArgs e)
    {
        e.Cancel = true;
        batchOpts.IsStoppedByUser = true;
    }
}