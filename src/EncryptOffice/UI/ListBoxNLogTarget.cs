using System;
using System.Globalization;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using NLog;
using NLog.Targets;
using purify.Data;

namespace EncryptOffice.UI;

public class ListBoxNLogTarget : TargetWithLayout
{
    private readonly ListBox listbox;
    private readonly BatchOptions model;

    public ListBoxNLogTarget(ListBox lb, BatchOptions batchOptions)
    {
        listbox = lb;
        model = batchOptions;
    }

    protected override void Write(LogEventInfo logEvent)
    {
        listbox.Dispatcher?.BeginInvoke(new Action<LogEventInfo>(AppendLine), logEvent);
    }

    private void AppendLine(LogEventInfo evt)
    {
        model.AppendLog(evt);
        listbox.ScrollIntoView(evt);
    }
}

/// <summary>
/// log级别的颜色的对应关系
/// </summary>
public class LevelToColorConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        var entry = value as LogEventInfo;
        if (entry == null) return Brushes.Transparent;
        if (entry.Level >= LogLevel.Error) return Brushes.LightPink;
        if (entry.Level == LogLevel.Warn) return Brushes.Yellow;
        return Brushes.Transparent;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return null;
    }
}