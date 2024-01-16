using AlphaCAMMill;
using System;
using System.IO;
using System.Windows;

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
      int andUpdateRunCount = this.GetAndUpdateRunCount();
      Drawing activeDrawing = this._acamApp.ActiveDrawing;
      if (activeDrawing.Geometries.Count == 0)
      {
        MessageBox.Show("No geometries found to measure.");
      }
      else
      {
        foreach (object geometry in activeDrawing.Geometries)
        {
          if (geometry is IPath element && IsGeometryLayer(element))
          {
            if (andUpdateRunCount % 2 == 1)
            {
              this.CreateDimensions(element);
              activeDrawing.ZoomAll();
            }
            else
            {
              this.DeleteDimensions();
            }
          }

          activeDrawing.Refresh();
        }
      }
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
        {
          throw new IOException("Error reading run count: " + ex.Message, ex);
        }
      }
      try
      {
        File.WriteAllText(this._runCountFilePath, andUpdateRunCount.ToString());
      }
      catch (Exception ex)
      {
        throw new IOException("Error writing run count: " + ex.Message, ex);
      }
      return andUpdateRunCount;
    }


    
    
    
    private void CreateDimensions(IPath path)
    {
      double offset = 10.0;
      Element previousElement = null;

      int count = 0;
      int totalElements = path.GetElemCount();
      
      bool previousWasArc = false; // Add this line

      Element element = path.GetFirstElem();
      while (element != null && count < totalElements)
      {
        // Skip if no new elements are found
        if (previousElement != null && element.IsSame(previousElement))
        {
          MessageBox.Show("Processing the same element again. No new elements found, exiting the loop.");
          break;
        }

        Element nextElement = element.GetNext(); 

        // Get parameters for creating dimensions
        double startX = element.StartXG;
        double startY = element.StartYG;
        double endX = element.EndXG;
        double endY = element.EndYG;

        // Get direction vectors
        element.GetDirection(startX, startY, out double startDirX, out double startDirY, out _);
        element.GetDirection(endX, endY, out double endDirX, out double endDirY, out _);

        // Extend the length of the line if followed by an arc
        if (nextElement != null && nextElement.IsArc)
        {
          endX += nextElement.Radius * endDirX;
          endY += nextElement.Radius * endDirY;
        }

        // Extend the start of the line if the previous element was an arc
        if (previousWasArc)
        {
          startX -= previousElement.Radius * startDirX;
          startY -= previousElement.Radius * startDirY;
        }

        // Remember if the current element is an arc for the next iteration
        previousWasArc = element.IsArc;

        // Join declaration and assignment
        double avgDirX = (startDirX + endDirX) / 2.0;
        double avgDirY = (startDirY + endDirY) / 2.0;

        double perpX, perpY; // Perpendicular

        if (element.IsArc)
        {
            // Arc has different winding direction compared to path
            // CW arc in a CW path or CCW arc in a CCW path
            if ( (!element.CW && avgDirY < 0.0) || (element.CW && avgDirY >= 0.0))
            {
                perpX = -avgDirY;
                perpY = avgDirX;
            }
            else
            {
                perpX = avgDirY;
                perpY = -avgDirX;
            }
        }
        else
        {
            // Line segment direction based on perpendicular
            perpX = avgDirY;
            perpY = -avgDirX;
        }

        // Compute middle position and dimension vectors
        double midX = (startX + endX) / 2.0;
        double midY = (startY + endY) / 2.0;
        double dimX = midX - offset * perpX;
        double dimY = midY - offset * perpY;

        try
        {
            // Create aligned dimension
            this._acamApp.ActiveDrawing.Dimension.CreateAligned(startX, startY, endX, endY, dimX, dimY, dimX, dimY);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Exception: {ex.Message}"); 
        }

        // Update for the next loop
        previousElement = element;
        element = nextElement; // Get the next element
        count++;
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

    public void ResetRunCount()
    {
      try
      {
        File.WriteAllText(this._runCountFilePath, "2");
      }
      catch (Exception ex)
      {
        MessageBox.Show("Error resetting run count: " + ex.Message);
      }
    }
  }
}
