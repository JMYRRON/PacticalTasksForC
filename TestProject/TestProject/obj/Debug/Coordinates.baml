   M S B A M L   `   `   `  ����F  BTestProject, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null)clr-namespace:TestProjectTestProject  5          TestProject.Coordinates   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35^ ZWindowsFormsIntegration, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5      R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5      Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5      @d2http://schemas.microsoft.com/expression/blend/2008     5      +localclr-namespace:TestProject     5      [ WEsri.ArcGISRuntime, Version=100.7.0.0, Culture=neutral, PublicKeyToken=8fc3cc631e44ad866esri+http://schemas.esri.com/arcgis/runtime/2013  5      -   5          � 
ResizeMode$  NoResize=�5       � WindowStyle$ 
ToolWindow=�5
   7   .��5      � .+�5   
   ,  &Esri.ArcGISRuntime.UI.Controls.MapView  -   5          Name$ 	MyMapView��5      $��0,0,-0.4,-0.2q�6*   
   Map. 5      $  Esri.ArcGISRuntime.Mapping.Map     Basemap 5      (  "Esri.ArcGISRuntime.Mapping.Basemap  5         Name$ Basemap��6#      
BaseLayers. 5      1  +Esri.ArcGISRuntime.Mapping.ArcGISTiledLayer  &   Esri.ArcGISRuntime.Mapping.Layer   Name$ World Topographic��60      Source$T Nhttp://services.arcgisonline.com/ArcGIS/rest/services/World_Topo_Map/MapServer?�6I   5      5      5      �� 5   
   -   5        	 ��Name$	 TextBox1��5      $
��Left=�6$   $��22��6?   $��	10,10,0,0q�6K   
 �� TextWrapping$

 Wrap=�6^   $�� ��6r   $	��Top=�6z   $	��289��6�   5      5      