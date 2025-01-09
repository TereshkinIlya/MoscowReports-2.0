﻿using System.Windows;
using Infrastructure.Interfaces;

namespace MoscowReports_2._0
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow(IViewModel viewModel)
        {
            DataContext = viewModel;
            InitializeComponent();
        }
    }
}