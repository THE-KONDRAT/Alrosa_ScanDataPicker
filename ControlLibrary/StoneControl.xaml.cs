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
    /// Логика взаимодействия для StoneControl.xaml
    /// </summary>
    public partial class StoneControl : UserControl, INotifyPropertyChanged
    {
        // property changed event
        public event PropertyChangedEventHandler PropertyChanged;

        public delegate void SetSelected(StoneControl sender);
        public event SetSelected OnSetSelected;

        public static readonly DependencyProperty FileFoundProperty = DependencyProperty.Register(
            "FileFound", typeof(bool), typeof(StoneControl),
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

        /*public static readonly DependencyProperty SelectedDP = DependencyProperty.Register(
            "Selected", typeof(bool), typeof(StoneControl),
             new FrameworkPropertyMetadata(
                false, new PropertyChangedCallback(OnFileFoundChanged)
                ));

        public bool Selected
        {
            get { return (bool)GetValue(SelectedDP); }
            set
            {
                SetValue(SelectedDP, value);
                OnPropertyChanged("Selected");
            }
        }*/

        private bool selected;

        public bool Selected
        {
            get { return selected; }
            set
            {
                selected = value;
                PaintSelected(selected);
                OnPropertyChanged("Selected");
            }
        }


        public StoneControl()
        {
            InitializeComponent();
        }

        private static void OnFileFoundChanged(DependencyObject sender, DependencyPropertyChangedEventArgs e)
        {
            if (e.Property.PropertyType.Equals(typeof(bool)))
            {
                ((StoneControl)sender).ChangeFound((bool)e.NewValue);
            }
        }

        public void RefreshFound()
        {
            ChangeFound(FileFound);
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

        private void PaintSelected(bool selected)
        {
            Color sColor = Color.FromRgb(160, 0, 160);
            Color nsColor = Color.FromRgb(18, 18, 98);
            SolidColorBrush b = new SolidColorBrush(sColor);
            if (!selected)
            {
                b = new SolidColorBrush(nsColor);
            }

            outputBorder.BorderBrush = b;
        }

        internal void OnPropertyChanged(String property)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(property));
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OnSetSelected?.Invoke(this);
        }
    }
}
