using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ControlLibrary
{
    /// <summary>
    /// Логика взаимодействия для StonePropertyControl.xaml
    /// </summary>
    public partial class StonePropertyControl : UserControl, INotifyPropertyChanged
    {
        // property changed event
        public event PropertyChangedEventHandler PropertyChanged;

        public static readonly DependencyProperty FileFoundProperty = DependencyProperty.Register(
           "FileFound", typeof(bool), typeof(StonePropertyControl),
            new FrameworkPropertyMetadata(
               false, new PropertyChangedCallback(OnFileFoundChanged)
               ));

        public bool FileFound
        {
            get { return (bool)GetValue(FileFoundProperty); }
            set
            {
                SetValue(FileFoundProperty, value);
                OnPropertyChanged("FileFound");
            }
        }

        public StonePropertyControl()
        {
            InitializeComponent();
        }

        public void RefreshFound()
        {
            ChangeFound(FileFound);
        }

        private static void OnFileFoundChanged(DependencyObject sender, DependencyPropertyChangedEventArgs e)
        {
            if (e.Property.PropertyType.Equals(typeof(bool)))
            {
                ((StonePropertyControl)sender).ChangeFound((bool)e.NewValue);
            }
        }

        private void ChangeFound(bool value)
        {
            Color fColor = Color.FromRgb(0, 255, 0);
            Color nfColor = Color.FromRgb(255, 0, 0);
            string fText = "\xE73E";
            string nfText = "\xE711";
            SolidColorBrush b = new SolidColorBrush(fColor);
            lblFileFound.Content = fText;
            if (!value)
            {
                lblFileFound.Content = nfText;
                b = new SolidColorBrush(nfColor);
            }

            lblFileFound.Foreground = b;
        }

        internal void OnPropertyChanged(String property)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(property));
            }
        }
    }
}
