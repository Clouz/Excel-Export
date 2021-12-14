﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office;
using Excel_Export;


namespace Finestra
{
    /// <summary>
    /// Logica di interazione per MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        Excel_Export.Export data;
        int _noOfErrorsOnScreen = 0;

        public MainWindow(Excel_Export.Export data)
        {
            this.data = data;
            //MessageBox.Show("Col: " + this.data.TotalCol.ToString() + "\nRow:" + this.data.TotalRow.ToString());
            //MessageBox.Show("Array Length: " + this.data.array.Length.ToString() + "\nIndex: " + this.data.ColumnIndex + "\nHeader: " + this.data.TableHeader);
            this.data.UpdateUniqueList();

            InitializeComponent();
            DataContext = this.data;

            ListaItems.ItemsSource = this.data.UniqueList;
        }

        private void Esegui_Click(object sender, RoutedEventArgs e)
        {
            data.ToNewFiles();
            this.Close();
        }

        private void OnValidationError(object sender, ValidationErrorEventArgs e)
        {
            if (e.Action == ValidationErrorEventAction.Added)
                _noOfErrorsOnScreen++;
            else
                _noOfErrorsOnScreen--;

            Esegui.IsEnabled = _noOfErrorsOnScreen > 0 ? false : true;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

        }

        private void Index_TextChanged(object sender, TextChangedEventArgs e)
        {
            data.UpdateUniqueList();
        }
    }
}
