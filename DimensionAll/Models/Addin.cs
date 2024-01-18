using AlphaCAMMill;
using System;
using System.IO;
using Serilog;

namespace DimensionAll.Models
{
  public sealed class Addin
  {

    private readonly App _acamApp;
    private readonly string _runCountFilePath;
    
    public Addin(App acamApplication)
    {
      this._acamApp = acamApplication;
      string str = System.IO.Path.Combine(this._acamApp.LicomdirPath, "Licomdir", "ESE_TOOLS");
      if (!Directory.Exists(str))
        Directory.CreateDirectory(str);
      this._runCountFilePath = System.IO.Path.Combine(str, "run_count.txt");
    }
    
    public void DimensionAll()
    {
      
      Log.Information($"DimensionAll - Start");
      int andUpdateRunCount = this.GetAndUpdateRunCount();
      Drawing activeDrawing = this._acamApp.ActiveDrawing;
      if (activeDrawing.Geometries.Count == 0)
      {
        Log.Warning("No geometries found to measure.");
      }
      else
      {
        Log.Information($"DimensionAll - Processing {activeDrawing.Geometries.Count} geometries.");

        SetToolSideForAllGeometries(activeDrawing.Geometries); //Set tool side for all geometries

        foreach (object geometry in activeDrawing.Geometries)
        {
          if (geometry is IPath element && IsGeometryLayer(element))
          {
            if (!DetermineIfInternal(element) && element.Closed) // Add this condition to check if the element is external and closed
            {
              if (andUpdateRunCount % 2 == 1)
              {
                this.CreateDimensions(element);
                activeDrawing.ZoomAll();
                Log.Information("DimensionAll - Dimensions created and zoomed all.");
              }
              else
              {
                this.DeleteDimensions();
                Log.Information("DimensionAll - Dimensions deleted.");
              }
            }
            activeDrawing.Refresh();
          }
        }
      }
      Log.Information("DimensionAll - End");
    }

    private int GetAndUpdateRunCount()
    {
      int andUpdateRunCount = 1;
      if (File.Exists(this._runCountFilePath))
      {
        try
        {
          if (!int.TryParse(File.ReadAllText(this._runCountFilePath), out int result))
            throw new InvalidOperationException("The run count file contains invalid data.");
          andUpdateRunCount = result == 1 ? 2 : 1;
        }
        catch (Exception ex)
        { Log.Error("Error reading run count: " + ex.Message, ex);
          // In case of exception, should we re-throw, halt the execution, or continue? If we re-throw or halt, uncomment the line below
          // throw new IOException("Error reading run count: " + ex.Message, ex);
        }
      }
      try
      {
        File.WriteAllText(this._runCountFilePath, andUpdateRunCount.ToString());
      }
      catch (Exception ex)
      {
        Log.Error("Error writing run count: " + ex.Message, ex);
        // Should we re-throw or halt execution here in case of exception?
      }
      return andUpdateRunCount;
    }

private void CreateDimensions(IPath path)
{
    double offset = 10.0;
    int count = 0;
    int totalElements = path.GetElemCount();
    double minimumLength = 4.0; // Set your minimum length here

    bool previousWasArc = false;
    Element previousElement = null;

    Element element = path.GetFirstElem();
    while (count++ < totalElements && element != null)
    {
        Element nextElement = element.GetNext();

        double startX = element.StartXG;
        double startY = element.StartYG;
        double endX = element.EndXG;
        double endY = element.EndYG;

        ApplyOffsetIfNextElementIsArc(element, nextElement, ref endX, ref endY);

        if (previousWasArc && previousElement != null)
        {
          previousElement.GetDirection(startX, startY, out var startDirX, out var startDirY, out _);
            ApplyOffsetForPreviousArc(previousElement, ref startX, ref startY, startDirX, startDirY);
        }

        previousWasArc = element.IsArc;

        // Check if the line length is greater than the minimum length and if the element is not an arc
        if (!element.IsArc && element.Length > minimumLength)
        {
          element.GetDirection(startX, startY, out var startDirX, out var startDirY, out _);
            element.GetDirection(endX, endY, out var endDirX, out var endDirY, out _);

            var avgDirX = (startDirX + endDirX) / 2.0;
            var avgDirY = (startDirY + endDirY) / 2.0;

            CalculatePerpendicularDirections(element, avgDirX, avgDirY, out double perpX, out double perpY);

            double midX = (startX + endX) / 2.0;
            double midY = (startY + endY) / 2.0;
            double dimX = midX - offset * perpX;
            double dimY = midY - offset * perpY;

            try
            {
                this._acamApp.ActiveDrawing.Dimension.CreateAligned(startX, startY, endX, endY, dimX, dimY, dimX, dimY);
            }
            catch (Exception ex)
            {
                ShowErrorAndBreak($"Exception: {ex.Message}");
            }
        }
        previousElement = element;
        element = nextElement;
    }
}

    private void ShowErrorAndBreak(string errorMessage)
    {
      Log.Error(errorMessage);
    }
    
    private void ApplyOffsetIfNextElementIsArc(Element element, Element nextElement, ref double endX, ref double endY)
    {
      element.GetDirection(endX, endY, out var endDirX, out var endDirY, out _);

      if (nextElement != null && nextElement.IsArc)
      {
        endX += nextElement.Radius * endDirX;
        endY += nextElement.Radius * endDirY;
      }
    }
    private void ApplyOffsetForPreviousArc(Element previousElement, ref double startX, ref double startY, double startDirX, double startDirY)
    {
        startX -= previousElement.Radius * startDirX;
        startY -= previousElement.Radius * startDirY;
    }

    private void CalculatePerpendicularDirections(Element element, double avgDirX, double avgDirY, out double perpX, out double perpY)
    {
      if (element.IsArc)
      {
        perpX = (!element.CW && avgDirY < 0.0) || (element.CW && avgDirY >= 0.0)? -avgDirY: avgDirY;
        perpY = (!element.CW && avgDirY < 0.0) || (element.CW && avgDirY >= 0.0)? avgDirX: -avgDirX;
      }
      else
      {
        perpX = avgDirY;
        perpY = -avgDirX;
      }
    }


    private void DeleteDimensions()
    {
      this._acamApp.ActiveDrawing.Clear(false, false, false, true, false, false, false, false);
    }

    private bool IsGeometryLayer(IPath element)
    {
      Layer layer = element.GetLayer();
      return layer != null && layer.Name.StartsWith("APS GEOMETRY", StringComparison.OrdinalIgnoreCase);
    }
    
    private void SetToolSideForAllGeometries(Paths paths)
    {
      Log.Information($"SetToolSideForAllGeometries - Start, paths count: {paths.Count}");
      int closedPathsCount = 0;
      try
      {
        foreach (var path in paths)
        {
          if (path is IPath pathObject && pathObject.Closed)
          {
            closedPathsCount++;
            pathObject.Selected = true;
          }
        }
        Log.Information($"SetToolSideForAllGeometries - closed paths count: {closedPathsCount}");

        _acamApp.ActiveDrawing.SetToolSideAuto(AcamAutoToolSide.acamToolSideCUT);
      }
      catch (Exception ex)
      {
        Log.Error($"Error in SetToolSideAuto: paths count: {paths.Count}, closed paths: {closedPathsCount}", ex);
        //MessageBox.Show($"Error in SetToolSideAuto: {ex.Message}");
      }
      Log.Information("SetToolSideForAllGeometries - End");
    }
      
    private bool DetermineIfInternal(IPath element) //check if its internal
    {
      return element.ToolSide == AcamToolSide.acamRIGHT;
    }
    
    public void ResetRunCount()
    {
      try
      {
        File.WriteAllText(this._runCountFilePath, "2");
      }
      catch (Exception ex)
      {
        Log.Error("Error resetting run count: " + ex.Message, ex);
      }
    }
  }
}
