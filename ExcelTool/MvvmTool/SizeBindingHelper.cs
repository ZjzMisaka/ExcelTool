using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ExcelTool
{
    public static class SizeBindingHelper
    {
        public static readonly DependencyProperty ActiveProperty = DependencyProperty.RegisterAttached(
            "Active",
            typeof(bool),
            typeof(SizeBindingHelper),
            new FrameworkPropertyMetadata(OnActiveChanged));

        public static bool GetActive(FrameworkElement frameworkElement)
        {
            return (bool)frameworkElement.GetValue(ActiveProperty);
        }

        public static void SetActive(FrameworkElement frameworkElement, bool active)
        {
            frameworkElement.SetValue(ActiveProperty, active);
        }

        public static readonly DependencyProperty BoundActualWidthProperty = DependencyProperty.RegisterAttached(
            "BoundActualWidth",
            typeof(double),
            typeof(SizeBindingHelper));

        public static double GetBoundActualWidth(FrameworkElement frameworkElement)
        {
            return (double)frameworkElement.GetValue(BoundActualWidthProperty);
        }

        public static void SetBoundActualWidth(FrameworkElement frameworkElement, double width)
        {
            frameworkElement.SetValue(BoundActualWidthProperty, width);
        }

        public static readonly DependencyProperty BoundActualHeightProperty = DependencyProperty.RegisterAttached(
            "BoundActualHeight",
            typeof(double),
            typeof(SizeBindingHelper));

        public static double GetBoundActualHeight(FrameworkElement frameworkElement)
        {
            return (double)frameworkElement.GetValue(BoundActualHeightProperty);
        }

        public static void SetBoundActualHeight(FrameworkElement frameworkElement, double height)
        {
            frameworkElement.SetValue(BoundActualHeightProperty, height);
        }

        private static void OnActiveChanged(DependencyObject dependencyObject, DependencyPropertyChangedEventArgs e)
        {
            if (!(dependencyObject is FrameworkElement frameworkElement))
            {
                return;
            }

            if ((bool)e.NewValue)
            {
                frameworkElement.SizeChanged += OnFrameworkElementSizeChanged;
                UpdateObservedSizesForFrameworkElement(frameworkElement);
            }
            else
            {
                frameworkElement.SizeChanged -= OnFrameworkElementSizeChanged;
            }
        }

        private static void OnFrameworkElementSizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (sender is FrameworkElement frameworkElement)
            {
                UpdateObservedSizesForFrameworkElement(frameworkElement);
            }
        }

        private static void UpdateObservedSizesForFrameworkElement(FrameworkElement frameworkElement)
        {
            frameworkElement.SetCurrentValue(BoundActualWidthProperty, frameworkElement.ActualWidth);
            frameworkElement.SetCurrentValue(BoundActualHeightProperty, frameworkElement.ActualHeight);
        }
    }
}
