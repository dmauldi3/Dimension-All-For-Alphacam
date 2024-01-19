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
      _acamApp = acamApplication;
      string directoryPath = System.IO.Path.Combine(_acamApp.LicomdirPath, "Licomdir", "ESE_TOOLS");
      if (!Directory.Exists(directoryPath))
        Directory.CreateDirectory(directoryPath);
      _runCountFilePath = System.IO.Path.Combine(directoryPath, "run_count.txt");
    }

    public void DimensionAll()
    {
      Log.Information($"DimensionAll - Start");
      var currentRunCount = GetAndUpdateRunCount();
      Log.Information($"RunCount= {currentRunCount}");

      Drawing activeDrawing = this._acamApp.ActiveDrawing;
      if (activeDrawing.Geometries.Count == 0)
      {
        Log.Warning("No geometries found to measure.");
      }
      else
      {
        Log.Information($"DimensionAll - Processing {activeDrawing.Geometries.Count} geometries.");

        Log.Information("Calling SetToolSideForAllGeometries()");
        SetToolSideForAllGeometries(activeDrawing.Geometries); // Set tool side for all geometries
        Log.Information("Returned from SetToolSideForAllGeometries()");

        foreach (object geometry in activeDrawing.Geometries)
        {
          if (geometry is IPath element)
          {
            Log.Information("Processing a geometry item.");

            if (IsGeometryLayer(element))
            {
              Log.Information("The element is a {ElementType} path on the geometry layer and is external",
                element.Closed ? "closed" : "open");

              if (element.Closed)
              {
                Log.Information("The element is closed.");

                if (currentRunCount % 2 == 1)
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
              else
              {
                Log.Warning("The element is not closed, skipping dimension creation");
              }

              activeDrawing.Refresh();
            }
            else
            {
              Log.Warning("Skipping an element because it's not on the geometry layer.");
            }
          }
          else
          {
            Log.Information("Skipping a non-path geometry item");
          }
        }

        Log.Information("DimensionAll - End");
      }
    }

    private int GetAndUpdateRunCount()
    {
      int andUpdateRunCount = 1;
      try
      {
        if (File.Exists(this._runCountFilePath))
        {
          if (!int.TryParse(File.ReadAllText(this._runCountFilePath), out int result))
          {
            throw new InvalidOperationException("The run count file contains invalid data.");
          }
          andUpdateRunCount = result == 1 ? 2 : 1;
        }
        File.WriteAllText(this._runCountFilePath, andUpdateRunCount.ToString());
      }
      catch (Exception ex)
      {
        Log.Error(ex, "Error processing the run count: " + ex.Message);
        throw; // Re-throwing the exception
      }
      return andUpdateRunCount;
    }

    
void CreateDimensions(IPath path)
{
    Log.Information("Starting to create dimensions for path");

    double offset;
    int count = 0;
    int totalElements = path.GetElemCount();
    double minimumLength = 4.0; // Set your minimum length here

    bool previousWasArc = false;
    Element previousElement = null;

    Element element = path.GetFirstElem();
    while (count++ < totalElements && element != null)
    {
        Element nextElement = element.GetNext();

        Log.Debug("Processing element {ElementNumber} of {TotalElements}", count, totalElements);

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

        Log.Debug("Processing element {ElementNumber} of {TotalElements}, Start X {startX}, Start Y {startY}, End X {endX}, End Y {endY}", count, totalElements, Math.Round(startX, 3), Math.Round(startY, 3), Math.Round(endX, 3), Math.Round(endY, 3));
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
            
            // Adjust offset based on whether the path is internal
            offset = DetermineIfInternal(path) ? -10.0 : 10.0;

            double dimX = midX - offset * perpX;
            double dimY = midY - offset * perpY;

            Log.Information("Creating an aligned dimension for element {ElementNumber}, Dimension X {dimX}, Dimension Y {dimY}", count, Math.Round(dimX, 3), Math.Round(dimY, 3));
            
            double controlX = midX - offset * perpX * 1.5; // Example control point
            double controlY = midY - offset * perpY * 1.5; //

            try
            {
                this._acamApp.ActiveDrawing.Dimension.CreateAligned(startX, startY, endX, endY, dimX, dimY, controlX, controlY);
                Log.Debug("Dimension for element {ElementNumber} created successfully", count);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Failed to create dimension for element {ElementNumber}", count);
                ShowErrorAndBreak($"Exception: {ex.Message}");
            }
        } else {
          if (element.IsArc)
            Log.Information("Element {ElementNumber} is an arc, so no dimension is created", count);
          else
            Log.Information("Element {ElementNumber} has a length less than the minimum length {minimumLength}, so no dimension is created", count, minimumLength);
        }

        Log.Debug("Finished processing element {ElementNumber}", count);

        previousElement = element;
        element = nextElement;
    }

    Log.Information("Finished creating dimensions");
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
      Log.Debug($"CalculatePerpendicularDirections - Start, element: {element}, avgDirX: {avgDirX}, avgDirY: {avgDirY}");

      if (element.IsArc)
      {
        perpX = (!element.CW && avgDirY < 0.0) || (element.CW && avgDirY >= 0.0)? -avgDirY: avgDirY;
        perpY = (!element.CW && avgDirY < 0.0) || (element.CW && avgDirY >= 0.0)? avgDirX: -avgDirX;

        Log.Debug($"For Arc, perpX: {perpX}, perpY: {perpY}");
      }
      else
      {
        // Lines segments are always counterclockwise
        perpX = -avgDirY;
        perpY = avgDirX;

        Log.Debug($"For Non-Arc, perpX: {perpX}, perpY: {perpY}");
      }

      Log.Debug($"CalculatePerpendicularDirections - End");
    }
    

    private void DeleteDimensions()
    {
      this._acamApp.ActiveDrawing.Clear(false, false, false, true, false, false, false, false);
    }

    
    private bool IsGeometryLayer(IPath element)
    {
      Layer layer = element.GetLayer();
      bool isGeometryLayer = layer != null && layer.Name.StartsWith("APS GEOMETRY", StringComparison.OrdinalIgnoreCase);

      Log.Debug($"Element on layer: '{layer?.Name}', is considered as on GeometryLayer: {isGeometryLayer}");

      return isGeometryLayer;
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
      }
      Log.Information("SetToolSideForAllGeometries - End");
    }
      
    
    private bool DetermineIfInternal(IPath element) //check if its internal
    {
      bool isInternal = element.ToolSide == AcamToolSide.acamLEFT;

      Log.Debug($"Element with tool side: '{element.ToolSide}', is considering as IsInternal: {isInternal}");

      return isInternal;
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
