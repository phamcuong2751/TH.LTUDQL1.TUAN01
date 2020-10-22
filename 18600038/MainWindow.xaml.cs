using Microsoft.Win32;
using System;
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
using Aspose.Cells;
using System.Diagnostics;

namespace _18600038
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        private void OpenFileFolder_Click(object sender, RoutedEventArgs e)
        {
            var screen = new OpenFileDialog();
            string[] Column = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K" };
            if (screen.ShowDialog() == true)
            {
                var FileExcel = new Workbook(screen.FileName);
                var Sheets = FileExcel.Worksheets;
                var db = new MyStoreEntities();
                var _id = 1;

                foreach (var Sheet in Sheets)
                {
                    Debug.WriteLine(Sheet.Name);
                    var i = 2;
         
                    var Row = 3;
                    var Cell = Sheet.Cells[$"C3"];
                    var caterogy = new Category()
                    {
                        Id = _id,
                        Name = Sheet.Name
                    };
                    db.Category.Add(caterogy);
                    db.SaveChanges();
                    _id++;
                    
                    while (Cell.Value != null)
                    {
                        i = 2;                        
                        var sku = Sheet.Cells[$"{Column[i]}{Row}"].StringValue; i++;
                        var name = Sheet.Cells[$"{Column[i]}{Row}"].StringValue; i++;
                        var price = Sheet.Cells[$"{Column[i]}{Row}"].Value; i++;
                        var quality = Sheet.Cells[$"{Column[i]}{Row}"].Value; i++;

                        var product = new Product()
                        {
                            SKU = sku,
                            Name = name,
                           
                        };

                        caterogy.Product.Add(product);
                        db.SaveChanges();


                        Debug.WriteLine($"{sku} - {name} - {price} - {quality} \n");
                        Row++;
                        Cell = Sheet.Cells[$"B{Row}"];
                    }

                }
                MessageBox.Show("Notification", "Import Excel succesful!", MessageBoxButton.OK, MessageBoxImage.Information);
            }    
           
        }
    }
}
