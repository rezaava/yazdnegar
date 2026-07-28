using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using ShivaNegar.Forms.Controls;

namespace ShivaNegar.Forms.Converters
{
    public class IsProgressedConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if ((values[0] is ContentPresenter &&
                values[1] is int) == false)
            {
                return Visibility.Collapsed;
            }
            bool checkNextItem = System.Convert.ToBoolean(parameter.ToString());
            ContentPresenter contentPresenter = values[0] as ContentPresenter;
            int progress = (int)values[1];
            ItemsControl itemsControl = ItemsControl.ItemsControlFromItemContainer(contentPresenter);
            int index = itemsControl.ItemContainerGenerator.IndexFromContainer(contentPresenter);
            if (checkNextItem == true)
            {
                index++;
            }
            WizardProgressBar wizardProgressBar = itemsControl.TemplatedParent as WizardProgressBar;
            int percent = (int)(((double)index / wizardProgressBar.Items.Count) * 100);
            if (percent < progress)
            {
                return Visibility.Visible;
            }
            return Visibility.Collapsed;
        }
        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }
}
