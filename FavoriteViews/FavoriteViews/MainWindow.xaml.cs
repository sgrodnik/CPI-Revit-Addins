using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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

namespace FavoriteViews {

    public partial class MainWindow : Window {
        UIApplication uiApp;
        UIDocument uiDoc;
        Document doc;
        ObservableCollection<WPFElem> views = new ObservableCollection<WPFElem>();
        public MainWindow(ExternalCommandData commandData) {
            uiApp = commandData.Application;
            uiDoc = uiApp.ActiveUIDocument;
            doc = uiDoc.Document;
            InitializeComponent();
            ListBox1.ItemsSource = views;
        }

        private void btnAdd_Click(object sender, RoutedEventArgs e) {
            IEnumerable<Element> sel = uiDoc.Selection.GetElementIds().Select(id => doc.GetElement(id));
            var selWPFElems = sel.Select(el => new WPFElem(el));
            var viewIds = views.Select(el => el.Id);
            foreach (var item in selWPFElems) {
                if (viewIds.Contains(item.Id)) {
                    continue;
                }
                views.Add(item);
            }
        }

        private void ListBox1_SelectionChanged(object sender, SelectionChangedEventArgs e) {
            //if (ListBox1.SelectedItem != null)
            //    this.Title = (ListBox1.SelectedItem as WPFElem).Name;
        }

        private void btnRemove_Click(object sender, RoutedEventArgs e) {
            try {
                var selWPFElems = ListBox1.SelectedItems;
                var index = ListBox1.SelectedIndex;
                if (ListBox1.SelectedIndex != -1) {
                    for (int i = selWPFElems.Count - 1; i >= 0; i--)
                        views.Remove(selWPFElems[i] as WPFElem);
                }
                if (index > views.Count - 1) {
                    index = views.Count - 1;
                }
                ListBox1.SelectedIndex = index;
            } catch (Exception ex) {
                MessageBox.Show(ex.ToString(), "Exception");
            }
        }

        private void ListBox1_MouseDoubleClick(object sender, RoutedEventArgs e) {
            try {
                uiDoc.ActiveView = (ListBox1.SelectedItem as WPFElem).View;
            } catch (Exception ex) {
                MessageBox.Show(ex.ToString(), "Exception");
            }
        }

        private void btnCloseRest_Click(object sender, RoutedEventArgs e) {
            try {
                if (views.Count == 0) { return; }
                var OpenUIViews = uiDoc.GetOpenUIViews();
                foreach (var view in OpenUIViews) {
                    var arr = views.Select(el => el.Id);
                    if (!arr.Contains(view.ViewId.IntegerValue)) {
                        try {
                            view.Close();
                        } catch (Autodesk.Revit.Exceptions.InvalidOperationException) {
                            uiDoc.ActiveView = views.First().View;
                            view.Close();
                        }
                    }
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.ToString(), "Exception");
            }
        }
    }

    public class WPFElem {
        Element Element { get; set; }
        public WPFElem(Element el) {
            Element = el;
        }
        public string Name {
            get { return this.Element.Name; }
        }
        public string Type {
            get { return this.Element.LookupParameter("Семейство").AsValueString(); }
        }
        public int Id {
            get { return this.Element.Id.IntegerValue; }
        }
        public View View {
            get { return this.Element as View; }
        }
    }
}
