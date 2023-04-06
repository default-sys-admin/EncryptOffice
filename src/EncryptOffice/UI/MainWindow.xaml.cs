using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using DevExpress.Mvvm.POCO;
using DevExpress.Pdf;
using DevExpress.Xpf.Core;
using Microsoft.WindowsAPICodePack.Dialogs;
using NLog;
using NLog.Config;
using purify.Data;

namespace EncryptOffice.UI;

public partial class MainWindow : DXWindow
{
    private static readonly Logger log = LogManager.GetCurrentClassLogger();

    private readonly BatchOptions model = ViewModelSource<BatchOptions>.Create();

    public MainWindow()
    {
        InitializeComponent();
        DataContext = model;
    }

    private void MainWindow_OnLoaded(object sender, RoutedEventArgs e)
    {
        // register listbox as target
        var config = LogManager.Configuration;
        var target = new ListBoxNLogTarget(loglb, model)
                     {Layout = @"${date:format=yyyy-mm-dd HH\:mm\:ss.fff} [${level}] ${message}"};
        config.AddTarget("textbox", target);
        config.LoggingRules.Add(new LoggingRule("*", LogLevel.Info, target));
        LogManager.Configuration = config;
    }

    private void ButtonSelectInput_OnClick(object sender, RoutedEventArgs e)
    {
        var dlg = new CommonOpenFileDialog();
        dlg.InitialDirectory = model.InputFolder;
        dlg.IsFolderPicker = true;
        if (dlg.ShowDialog() == CommonFileDialogResult.Ok)
        {
            model.InputFolder = dlg.FileName;
        }
    }

    private void ButtonSelectOutput_OnClick(object sender, RoutedEventArgs e)
    {
        var dlg = new CommonOpenFileDialog();
        dlg.InitialDirectory = model.OutputFolder;
        dlg.IsFolderPicker = true;
        if (dlg.ShowDialog() == CommonFileDialogResult.Ok)
        {
            model.OutputFolder = dlg.FileName;
        }
    }

    private void ButtonReadme_OnClick(object sender, RoutedEventArgs e)
    {
        var wnd = new ReadmeWindow {Owner = this, WindowStartupLocation = WindowStartupLocation.CenterOwner};
        wnd.ShowDialog();
    }

    /// <summary>
    /// 开始实际操作
    /// </summary>
    private void ButtonRun_OnClick(object sender, RoutedEventArgs e)
    {
        // 在此处进行各种检查，因为msgbox需要在UI线程内显示

        // 检查路径
        if (!Path.IsPathFullyQualified(model.InputFolderNorm))
        {
            MessageBox.Show("输入目录必须使用绝对路径", "error", MessageBoxButton.OK);
            return;
        }

        if (!Path.IsPathFullyQualified(model.OutputFolderNorm))
        {
            MessageBox.Show("输出目录必须使用绝对路径", "error", MessageBoxButton.OK);
            return;
        }

        if (!Directory.Exists(model.InputFolderNorm))
        {
            MessageBox.Show($"{model.InputFolderNorm} 不存在", "error", MessageBoxButton.OK);
            return;
        }

        if (model.InputFolderNorm == model.OutputFolderNorm)
        {
            if (model.OverwriteOutputFile)
            {
                var ret = MessageBox.Show("输出目录和输入目录相同，确认覆盖？", "目录相同", MessageBoxButton.YesNo);
                if (ret != MessageBoxResult.Yes)
                    return;
            }
            else
            {
                MessageBox.Show("输出目录和输入目录相同，请选中覆盖选项", "目录相同", MessageBoxButton.OK);
                return;
            }
        }
        else if (Directory.Exists(model.OutputFolderNorm) && model.OverwriteOutputFile)
        {
            var ret = MessageBox.Show($"{model.OutputFolderNorm} 已存在，确认覆盖？", "目录已存在", MessageBoxButton.YesNo);
            if (ret != MessageBoxResult.Yes)
                return;
        }

        // 避免重入
        model.IsRunning = true;
        model.RunBtnTxt = "Running...";
        model.StopButtonIsVisible = true;

        // 启动后台线程，执行实际操作
        Task.Run(() => { model.Execute(InputPdfPwd); });
    }

    private void ButtonStop_OnClick(object sender, RoutedEventArgs e)
    {
        model.IsStoppedByUser = true;
    }

    private void InputPdfPwd(PdfPasswordRequestedEventArgs e)
    {
        if (!Dispatcher.CheckAccess())
        {
            Dispatcher.BeginInvoke(new Action(() => InputPdfPwd(e))).Wait();
            return;
        }

        var dlg = new PdfPasswordWindow
                  {
                      Filename = e.FileName,
                      Owner = this,
                      WindowStartupLocation = WindowStartupLocation.CenterOwner
                  };
        if (dlg.ShowDialog() == true)
        {
            e.PasswordString = dlg.Password;
        }
    }
}