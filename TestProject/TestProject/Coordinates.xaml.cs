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
using System.Windows.Navigation;
using System.Windows.Shapes;
using Esri.ArcGISRuntime.Data;
using Esri.ArcGISRuntime.Geometry;
using Esri.ArcGISRuntime.Mapping;
using Esri.ArcGISRuntime.Symbology;
using Esri.ArcGISRuntime.UI;
using Esri.ArcGISRuntime.UI.Controls;
using Brushes = System.Windows.Media.Brushes;

namespace TestProject
{
    /// <summary>
    /// Логика взаимодействия для Coordinates.xaml
    /// </summary>
    public partial class Coordinates : Window
    {
        bool[] bools = new bool[5];
        Esri.ArcGISRuntime.Geometry.PointCollection polylinePoints = new Esri.ArcGISRuntime.Geometry.PointCollection(SpatialReferences.WebMercator);
        Esri.ArcGISRuntime.Geometry.PointCollection polygonLinePoints = new Esri.ArcGISRuntime.Geometry.PointCollection(SpatialReferences.WebMercator);
        Esri.ArcGISRuntime.Geometry.PointCollection polygonPoints = new Esri.ArcGISRuntime.Geometry.PointCollection(SpatialReferences.WebMercator);

        static Microsoft.Office.Interop.Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
        private Form1 form1;
        

        public Coordinates(Form1 form1)
        {

            this.form1 = form1;
            for (int i = 0; i < bools.Length; i++)
            {
                bools[i] = true;
            }
            InitializeComponent();   
            initialize();           
            
                       
        }

        private void initialize()
        {
                        
            MyMapView.GraphicsOverlays.Add(new GraphicsOverlay());
            
            // Subscribe to map tap events to enable tapping on map to update coordinates
            MyMapView.GeoViewTapped += (sender, args) => { UpdateUIFromMapPoint(args.Location); };
            MyMapView.GeoViewTapped += deleteElement;
            //MyMapView.GeoViewDoubleTapped += (sender, args) => { createElement(); };
            MyMapView.MouseRightButtonDown += (sender, args) => { createElement(); };
            updateMap(form1.getCoords());
        }


        private void UpdateUIFromMapPoint(MapPoint selectedPoint)
        {
            if (!bools[1])
            {
                addPoints(selectedPoint);
            }
            else if (!bools[2])
            {
                addLine(selectedPoint);
            }
            else if (!bools[3])
            {
                addPolygon(selectedPoint);
            }

        }

        private void addPoints(MapPoint selectedPoint)
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
                    // Clear the selectionss symbol.
                    //MyMapView.GraphicsOverlays[0].Graphics.Clear();
                }
                return;
            }

            // Clear existing graphics overlays
            //  MyMapView.GraphicsOverlays[0].Graphics.Clear();

            // Create a symbol to symbolize the point
            SimpleMarkerSymbol symbol = new SimpleMarkerSymbol(SimpleMarkerSymbolStyle.Circle, System.Drawing.Color.Red, 5);

            // Create the graphic
            Graphic symbolGraphic = new Graphic(selectedPoint, symbol);

            GraphicsOverlay currentOverlay = new GraphicsOverlay();

            // Add the graphic to the graphics overlay
            currentOverlay.Graphics.Add(symbolGraphic);

            MyMapView.GraphicsOverlays.Add(currentOverlay);

            updateText();
        }

        private void addLine(MapPoint selectedPoint)
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
                    // Clear the selectionss symbol.
                    //MyMapView.GraphicsOverlays[0].Graphics.Clear();
                }
                return;
            }

            // Create a symbol to symbolize the point
            SimpleMarkerSymbol symbol = new SimpleMarkerSymbol(SimpleMarkerSymbolStyle.Circle, System.Drawing.Color.Red, 5);

            // Create the graphic
            Graphic symbolGraphic = new Graphic(selectedPoint, symbol);

            // Add the graphic to the graphics overlay
            MyMapView.GraphicsOverlays[0].Graphics.Add(symbolGraphic);

            //Add point to polylinePoints cillection
            polylinePoints.Add(selectedPoint);

            var polyline = new Esri.ArcGISRuntime.Geometry.Polyline(polylinePoints);

            //Create symbol for polyline
            var polylineSymbol = new SimpleLineSymbol(SimpleLineSymbolStyle.Solid, System.Drawing.Color.Red, 3);

            //Create a polyline graphic with geometry and symbol
            var polylineGraphic = new Graphic(polyline, polylineSymbol);

            //Add polyline to graphics overlay
            MyMapView.GraphicsOverlays[0].Graphics.Add(polylineGraphic);
        }

        private void addPolygon(MapPoint selectedPoint)
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
                    // Clear the selectionss symbol.
                    //MyMapView.GraphicsOverlays[0].Graphics.Clear();
                }
                return;
            }

            // Create a symbol to symbolize the point
            SimpleMarkerSymbol symbol = new SimpleMarkerSymbol(SimpleMarkerSymbolStyle.Circle, System.Drawing.Color.Blue, 5);

            // Create the graphic
            Graphic symbolGraphic = new Graphic(selectedPoint, symbol);

            // Add the graphic to the graphics overlay
            MyMapView.GraphicsOverlays[0].Graphics.Add(symbolGraphic);

            //Add point to polylinePoints cillection
            polygonPoints.Add(selectedPoint);
            polygonLinePoints.Add(selectedPoint);

            var polygonLine = new Esri.ArcGISRuntime.Geometry.Polyline(polygonLinePoints);

            //Create symbol for polyline
            var polygonLineSymbol = new SimpleLineSymbol(SimpleLineSymbolStyle.Solid, System.Drawing.Color.Blue, 3);

            //Create a polyline graphic with geometry and symbol
            var polygonLineGraphic = new Graphic(polygonLine, polygonLineSymbol);

            //Add polyline to graphics overlay
            MyMapView.GraphicsOverlays[0].Graphics.Add(polygonLineGraphic);


        }

        private async void deleteElement(object sender, GeoViewInputEventArgs e)
        {
            double tolerance = 10d; // Use larger tolerance for touch
            int maximumResults = 1; // Only return one graphic  
            bool onlyReturnPopups = false; // Return more than popups

            if (!bools[4])
            {
                try
                {
                    foreach (GraphicsOverlay overlay in MyMapView.GraphicsOverlays)
                    {
                        // Use the following method to identify graphics in a specific graphics overlay
                        IdentifyGraphicsOverlayResult identifyResults = await MyMapView.IdentifyGraphicsOverlayAsync(
                            overlay,
                            e.Position,
                            tolerance,
                            onlyReturnPopups,
                            maximumResults);
                        // Check if we got results
                        if (identifyResults.Graphics.Count > 0)
                        {
                            //  Display to the user the identify worked.                            
                            overlay.Graphics.Clear();
                            break;
                        }
                    }
                    updateText();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString(), "Error");
                }
            }
        }

        private void createElement()
        {
            if (!bools[3] && polygonPoints.Count > 2)
            {
                var polygon = new Esri.ArcGISRuntime.Geometry.Polygon(polygonPoints);

                //Create symbol for polygon with outline
                var polygonSymbol = new SimpleFillSymbol(SimpleFillSymbolStyle.Solid, System.Drawing.Color.FromArgb(120, 0, 0, 100),
                 new SimpleLineSymbol(SimpleLineSymbolStyle.Solid, System.Drawing.Color.Blue, 3));

                //Create polygon graphic with geometry and symbol
                Graphic polygonGraphic = new Graphic(polygon, polygonSymbol);

                //Add polyline to graphics overlay
                //MyMapView.GraphicsOverlays[0].Graphics.Add(polygonGraphic);

                GraphicsOverlay currentOverlay = new GraphicsOverlay();

                currentOverlay.Graphics.Add(polygonGraphic);

                MyMapView.GraphicsOverlays.Add(currentOverlay);

                MyMapView.GraphicsOverlays[0].Graphics.Clear();
                polygonLinePoints.Clear();
                polygonPoints.Clear();
                updateText();
            }
            else if (!bools[2] && polylinePoints.Count > 1)
            {
                var polyline = new Esri.ArcGISRuntime.Geometry.Polyline(polylinePoints);

                //Create symbol for polyline
                var polylineSymbol = new SimpleLineSymbol(SimpleLineSymbolStyle.Solid, System.Drawing.Color.Red, 3);

                //Create a polyline graphic with geometry and symbol
                var polylineGraphic = new Graphic(polyline, polylineSymbol);

                //Add polyline to graphics overlay
                //MyMapView.GraphicsOverlays[0].Graphics.Add(polygonGraphic);

                GraphicsOverlay currentOverlay = new GraphicsOverlay();

                currentOverlay.Graphics.Add(polylineGraphic);

                MyMapView.GraphicsOverlays.Add(currentOverlay);

                MyMapView.GraphicsOverlays[0].Graphics.Clear();
                polylinePoints.Clear();
                updateText();
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            double mapScale = MyMapView.MapScale;
            double zoom = mapScale / 5;
            mapScale += zoom;
            MyMapView.SetViewpointScaleAsync(mapScale);
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            double mapScale = MyMapView.MapScale;
            double zoom = mapScale / 5;
            mapScale -= zoom;
            MyMapView.SetViewpointScaleAsync(mapScale);
        }

        private void MyHand_Click(object sender, RoutedEventArgs e)
        {
            if (bools[0])
            {
                bools[0] = false;
                MyHand.Background = Brushes.Aquamarine;
                unckeckBools(0);
            }

        }

        private void MyPoint_Click(object sender, RoutedEventArgs e)
        {
            if (bools[1])
            {
                bools[1] = false;
                MyPoint.Background = Brushes.Aquamarine;
                unckeckBools(1);
            }
        }

        private void MyLine_Click(object sender, RoutedEventArgs e)
        {
            if (bools[2])
            {
                bools[2] = false;
                MyLine.Background = Brushes.Aquamarine;
                unckeckBools(2);
            }
        }

        private void MyPoligon_Click(object sender, RoutedEventArgs e)
        {
            if (bools[3])
            {
                bools[3] = false;
                MyPoligon.Background = Brushes.Aquamarine;
                unckeckBools(3);
            }
        }

        private void MyDelete_Click(object sender, RoutedEventArgs e)
        {
            if (bools[4])
            {
                bools[4] = false;
                MyDelete.Background = Brushes.Aquamarine;
                unckeckBools(4);
            }
        }

        private void unckeckBools(int index)
        {
            for (int i = 0; i < bools.Length; i++)
            {
                if (i != index)
                {
                    bools[i] = true;
                    switch (i)
                    {
                        case 0: MyHand.Background = Brushes.LightGray; break;
                        case 1: MyPoint.Background = Brushes.LightGray; break;
                        case 2: MyLine.Background = Brushes.LightGray; polylinePoints.Clear(); break;
                        case 3: MyPoligon.Background = Brushes.LightGray; polygonPoints.Clear(); break;
                        case 4: MyDelete.Background = Brushes.LightGray; polygonPoints.Clear(); break;
                    }
                }
            }
            MyMapView.GraphicsOverlays[0].Graphics.Clear();
        }

        private void updateText()
        {
            string mapPoints = "";
            string mapPolylines = "";
            string mapPolygons = "";

            foreach (GraphicsOverlay overlay in MyMapView.GraphicsOverlays)
            {
                foreach (Graphic graphic in overlay.Graphics)
                {
                    if (graphic.Geometry.ToString().Contains("MapPoint"))
                    {
                        mapPoints += graphic.Geometry.ToJson() + "\b";
                    }
                    else if (graphic.Geometry.ToString().Contains("Polyline"))
                    {
                        mapPolylines += graphic.Geometry.ToJson() + "\b";
                    }
                    else if (graphic.Geometry.ToString().Contains("Polygon"))
                    {
                        mapPolygons += graphic.Geometry.ToJson() + "\b";
                    }
                }
            }
            MyTextBlock.Text = mapPoints + mapPolylines + mapPolygons;
        }

        private void updateMap(string coordinates)
        {
            try
            {
                string[] coordinatesArray = coordinates.Split('\b');
                foreach (string coord in coordinatesArray)
                {                    
                    try
                    {
                        if (coord.Contains(@"x"))
                        {
                            GraphicsOverlay newOverlay = new GraphicsOverlay();
                            SimpleMarkerSymbol symbol = new SimpleMarkerSymbol(SimpleMarkerSymbolStyle.Circle, System.Drawing.Color.Red, 5);
                            Graphic graphic = new Graphic(Esri.ArcGISRuntime.Geometry.Geometry.FromJson(coord), symbol);
                            newOverlay.Graphics.Add(graphic);
                            MyMapView.GraphicsOverlays.Add(newOverlay);
                        }
                        else if (coord.Contains("paths"))
                        {
                            GraphicsOverlay newOverlay = new GraphicsOverlay();
                            var polylineSymbol = new SimpleLineSymbol(SimpleLineSymbolStyle.Solid, System.Drawing.Color.Red, 3);
                            Graphic graphic = new Graphic(Esri.ArcGISRuntime.Geometry.Geometry.FromJson(coord), polylineSymbol);
                            newOverlay.Graphics.Add(graphic);
                            MyMapView.GraphicsOverlays.Add(newOverlay);
                        }
                        else if (coord.Contains("rings"))
                        {
                            GraphicsOverlay newOverlay = new GraphicsOverlay();
                            var polygonSymbol = new SimpleFillSymbol(SimpleFillSymbolStyle.Solid, System.Drawing.Color.FromArgb(120, 0, 0, 100),
                     new SimpleLineSymbol(SimpleLineSymbolStyle.Solid, System.Drawing.Color.Blue, 3));


                            Graphic graphic = new Graphic(Esri.ArcGISRuntime.Geometry.Geometry.FromJson(coord), polygonSymbol);

                            newOverlay.Graphics.Add(graphic);
                            MyMapView.GraphicsOverlays.Add(newOverlay);
                        }
                    }
                    catch (Exception)
                    {

                    }
                }

            }
            catch (Exception)
            {
                MessageBox.Show("Coordinates are absent");
            }
        }

        public void setCoordsText(string text)
        {
            MyTextBlock.Text = text;
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            form1.setCoords(MyTextBlock.Text);
            Form1.changeMapCounter();
        }
    }
}
