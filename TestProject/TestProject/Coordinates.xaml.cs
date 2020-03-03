using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Windows;
using System.Windows.Controls;
using Form1 = TestProject.Form1;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Esri.ArcGISRuntime.Geometry;
using Esri.ArcGISRuntime.Mapping;
using Esri.ArcGISRuntime.Symbology;
using Esri.ArcGISRuntime.UI;


namespace TestProject
{
    /// <summary>
    /// Логика взаимодействия для Coordinates.xaml
    /// </summary>
    public partial class Coordinates : Window
    {        
        static Microsoft.Office.Interop.Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
        private Form1 form1;
        

        public Coordinates(Form1 form1)
        {

            this.form1 = form1;


            InitializeComponent();
           
            

            

            initialize();
            
            
                       
        }

        private void initialize()
        {
                        
            MyMapView.GraphicsOverlays.Add(new GraphicsOverlay());
            MapPoint startingPoint = new MapPoint(0, 0, SpatialReferences.WebMercator);

            // Update the UI with the initial point
            UpdateUIFromMapPoint(startingPoint);

            // Subscribe to text change events
            TextBox1.TextChanged += InputTextChanged;

            // Subscribe to map tap events to enable tapping on map to update coordinates
            MyMapView.GeoViewTapped += (sender, args) => { UpdateUIFromMapPoint(args.Location); };
                        
        }

        private void InputTextChanged(object sender, TextChangedEventArgs e)
        {
            // Keep track of the last edited field
            TextBox1 = (TextBox)sender;
        }


        private void UpdateUIFromMapPoint(MapPoint selectedPoint)
        {
            try
            {
                // Check if the selected point can be formatted into coordinates.
                CoordinateFormatter.ToLatitudeLongitude(selectedPoint, LatitudeLongitudeFormat.DecimalDegrees, 0);
            }
            catch (Exception e)
            {
                // Check if the excpetion is because the coordinates are out of range.
                if (e.Message == "Invalid argument: coordinates are out of range")
                {
                    // Set all of the text fields to contain the error message.
                    TextBox1.Text = "Coordinates are out of range";
                    
                    // Clear the selectionss symbol.
                    MyMapView.GraphicsOverlays[0].Graphics.Clear();
                }
                return;
            }

            // Update the decimal degrees text
            TextBox1.Text = CoordinateFormatter.ToLatitudeLongitude(selectedPoint, LatitudeLongitudeFormat.DecimalDegrees, 4);
            form1.setCoords(TextBox1.Text);

            // Clear existing graphics overlays
            MyMapView.GraphicsOverlays[0].Graphics.Clear();

            // Create a symbol to symbolize the point
            SimpleMarkerSymbol symbol = new SimpleMarkerSymbol(SimpleMarkerSymbolStyle.Circle, System.Drawing.Color.Red, 5);

            // Create the graphic
            Graphic symbolGraphic = new Graphic(selectedPoint, symbol);

            // Add the graphic to the graphics overlay
            MyMapView.GraphicsOverlays[0].Graphics.Add(symbolGraphic);
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Form1.changeMapCounter();
        }
    }
}
