using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Media.Imaging;
using System.Windows.Interop;
using System.Windows;
using System.Windows.Controls;

namespace TFSProjectMigration
{
    class ImageHelpers
    {

        public static BitmapSource GetImageSource(System.Drawing.Bitmap bitmap)
        {
            return Imaging.CreateBitmapSourceFromHBitmap(bitmap.GetHbitmap(),
                IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
        }

        public static StackPanel CreateHeader(string title, ItemTypes type)
        {
            StackPanel panel = new StackPanel();
            panel.Orientation = Orientation.Horizontal;

            Label lbl = new Label();
            lbl.Content = title;

            Image img = new Image();

            switch (type)
            {
                case ItemTypes.TestCase:
                    img.Source = ImageHelpers.GetImageSource(Properties.Resources.TestCase);
                    break;
                case ItemTypes.TestSuite:
                    img.Source = ImageHelpers.GetImageSource(Properties.Resources.Suite);
                    break;
                case ItemTypes.TestPlan:
                    img.Source = ImageHelpers.GetImageSource(Properties.Resources.Plan);
                    break;
                case ItemTypes.TeamProject:
                    img.Source = ImageHelpers.GetImageSource(Properties.Resources.TeamProject);
                    break;
            }
            panel.Children.Add(img);
            panel.Children.Add(lbl);
            return panel;
        }
    }
}
