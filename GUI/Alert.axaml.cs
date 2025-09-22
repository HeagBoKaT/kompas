using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Markup.Xaml;

namespace GUI;

public partial class Alert : Window
{
    public Alert(string message)
    {
        InitializeComponent();
        TextBlock.Text =  message;
    }
}