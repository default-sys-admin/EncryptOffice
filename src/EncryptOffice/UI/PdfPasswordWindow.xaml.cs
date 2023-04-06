using System.Windows;

namespace EncryptOffice.UI;

/// <summary>
/// Interaction logic for PdfPasswordWindow.xaml
/// </summary>
public partial class PdfPasswordWindow : Window
{
    public string Filename { get; set; }
    public string Password { get; set; }

    public PdfPasswordWindow()
    {
        InitializeComponent();
        this.DataContext = this;
    }

    private void ButtonBase_OnClick(object sender, RoutedEventArgs e)
    {
        Password = pwdBox.Password;
        DialogResult = true;
    }
}