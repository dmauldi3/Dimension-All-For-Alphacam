using AlphaCAMMill;
using System;
using System.IO;
using System.Windows;

namespace AlphacamAddinSample.Models
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
              this.CreateDimensions(element);
            else
              this.DeleteDimensions();
          }
        }
        activeDrawing.Refresh();
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

    private void CreateDimensions(IPath element)
    {
      double minXl = element.MinXL;
      double minYl = element.MinYL;
      double maxXl = element.MaxXL;
      double maxYl = element.MaxYL;
      double offset = 10.0;
      this.CreateAlignedDimension(minXl, minYl, minXl, maxYl, -offset);
      this.CreateAlignedDimension(minXl, maxYl, maxXl, maxYl, -offset);
      this.CreateAlignedDimension(maxXl, maxYl, maxXl, minYl, offset);
      this.CreateAlignedDimension(maxXl, minYl, minXl, minYl, offset);
    }

    private void DeleteDimensions()
    {
      this._acamApp.ActiveDrawing.Clear(false, false, false, true, false, false, false, false);
    }

    private void CreateAlignedDimension(
      double startX,
      double startY,
      double endX,
      double endY,
      double offset)
    {
      double xPosition = (startX + endX) / 2.0 + offset;
      double yPosition = (startY + endY) / 2.0 - offset;
      double num1 = endX - startX;
      double num2 = endY - startY;
      double num3 = Math.Sqrt(num1 * num1 + num2 * num2);
      double num4 = num1 / num3;
      double num5 = num2 / num3;
      double xControl = xPosition + offset * num5;
      double yControl = yPosition + offset * num4;
      this._acamApp.ActiveDrawing.Dimension.CreateAligned(startX, startY, endX, endY, xPosition, yPosition, xControl, yControl);
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
