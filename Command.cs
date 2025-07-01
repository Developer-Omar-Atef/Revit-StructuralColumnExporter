using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace KaitechColumnsReportAddin
{
    [Transaction(TransactionMode.ReadOnly)]

    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                using (var form = new ColumnReportForm())
                {
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        string selectedFolderPath = form.SelectedPath;
                        CollectAndExportData(doc, selectedFolderPath);
                    }
                }
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Error", $"An unexpected error occurred: {ex.Message}\n{ex.StackTrace}");
                return Result.Failed;
            }
        }
        private void CollectAndExportData(Document doc, string folderPath)
        {
            // 1. Collect Column Data
            List<ColumnData> columnsData = new List<ColumnData>();
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ICollection<Element> structuralColumns = collector
                .OfCategory(BuiltInCategory.OST_StructuralColumns)
                .WhereElementIsNotElementType()
                .ToElements();

            if (!structuralColumns.Any())
            {
                TaskDialog.Show("Info", "No structural columns found in the current model.");
                return;
            }

            foreach (Element elem in structuralColumns)
            {
                if (elem is FamilyInstance col)
                {
                    var data = new ColumnData();

                    // Family & Type
                    data.Family = col.Symbol?.Family?.Name ?? "N/A";
                    data.Type = col.Symbol?.Name ?? doc.GetElement(col.GetTypeId())?.Name ?? "N/A";
                    data.Id = col.Id.ToString();

                    // Location (Easting/Northing)
                    LocationPoint locPoint = col.Location as LocationPoint;
                    if (locPoint != null)
                    {
                        XYZ point = locPoint.Point;
                        //data.Easting = UnitUtils.ConvertFromInternalUnits(point.X, UnitTypeId.Meters).ToString("F3", CultureInfo.InvariantCulture);  // Revit 2021+
                        //data.Northing = UnitUtils.ConvertFromInternalUnits(point.Y, UnitTypeId.Meters).ToString("F3", CultureInfo.InvariantCulture); // Revit 2021+
                        data.Easting = UnitUtils.ConvertFromInternalUnits(point.X, DisplayUnitType.DUT_METERS).ToString("F3", CultureInfo.InvariantCulture);     // Revit 2021-
                        data.Northing = UnitUtils.ConvertFromInternalUnits(point.Y, DisplayUnitType.DUT_METERS).ToString("F3", CultureInfo.InvariantCulture);    // Revit 2021-
                    }
                    else
                    {
                        data.Easting = "N/A";
                        data.Northing = "N/A";
                    }

                    // Levels and Offsets
                    Parameter baseLevelParam = col.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM);
                    Parameter baseOffsetParam = col.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM);
                    Parameter topLevelParam = col.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM);
                    Parameter topOffsetParam = col.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM);

                    Level baseLevel = null;
                    double baseOffsetInternal = 0;
                    if (baseLevelParam != null && baseLevelParam.HasValue && baseLevelParam.AsElementId() != ElementId.InvalidElementId)
                    {
                        baseLevel = doc.GetElement(baseLevelParam.AsElementId()) as Level;
                        data.BaseLevel = baseLevel?.Name ?? "N/A";
                    }
                    else
                    {
                        data.BaseLevel = "N/A (Unconnected)";
                    }

                    if (baseOffsetParam != null && baseOffsetParam.HasValue)
                    {
                        baseOffsetInternal = baseOffsetParam.AsDouble();
                        //data.BaseOffset = UnitUtils.ConvertFromInternalUnits(baseOffsetInternal, UnitTypeId.Meters).ToString("F3", CultureInfo.InvariantCulture);         // Revit 2021+
                        data.BaseOffset = UnitUtils.ConvertFromInternalUnits(baseOffsetInternal, DisplayUnitType.DUT_METERS).ToString("F3", CultureInfo.InvariantCulture);  // Revit 2021-
                    }
                    else
                    {
                        data.BaseOffset = "0.000";
                    }

                    Level topLevel = null;
                    double topOffsetInternal = 0;
                    if (topLevelParam != null && topLevelParam.HasValue && topLevelParam.AsElementId() != ElementId.InvalidElementId)
                    {
                        topLevel = doc.GetElement(topLevelParam.AsElementId()) as Level;
                        data.TopLevel = topLevel?.Name ?? "N/A";
                    }
                    else
                    {
                        data.TopLevel = "N/A (Unconnected)";
                    }

                    if (topOffsetParam != null && topOffsetParam.HasValue)
                    {
                        topOffsetInternal = topOffsetParam.AsDouble();
                        //data.TopOffset = UnitUtils.ConvertFromInternalUnits(topOffsetInternal, UnitTypeId.Meters).ToString("F3", CultureInfo.InvariantCulture);        // Revit 2021+
                        data.TopOffset = UnitUtils.ConvertFromInternalUnits(topOffsetInternal, DisplayUnitType.DUT_METERS).ToString("F3", CultureInfo.InvariantCulture); // Revit 2021-
                    }
                    else
                    {
                        data.TopOffset = "0.000";
                    }

                    // Height
                    try
                    {
                        double baseElevationInternal = (baseLevel != null) ? baseLevel.ProjectElevation : 0;
                        double topElevationInternal = (topLevel != null) ? topLevel.ProjectElevation : 0;

                        // If levels are not defined, height calculation might be problematic or rely on other params
                        if (baseLevel == null && topLevel == null)
                        {
                            // Try to get instance length parameter if column is not level-bound
                            Parameter lengthParam = col.get_Parameter(BuiltInParameter.INSTANCE_LENGTH_PARAM);
                            if (lengthParam != null && lengthParam.HasValue)
                            {
                                //data.Height = UnitUtils.ConvertFromInternalUnits(lengthParam.AsDouble(), UnitTypeId.Meters).ToString("F3", CultureInfo.InvariantCulture);         // Revit 2021+
                                data.Height = UnitUtils.ConvertFromInternalUnits(lengthParam.AsDouble(), DisplayUnitType.DUT_METERS).ToString("F3", CultureInfo.InvariantCulture);  // Revit 2021-
                            }
                            else
                            {
                                data.Height = "N/A";
                            }
                        }
                        else
                        {
                            double heightInternal = (topElevationInternal + topOffsetInternal) - (baseElevationInternal + baseOffsetInternal);
                            //data.Height = UnitUtils.ConvertFromInternalUnits(heightInternal, UnitTypeId.Meters).ToString("F3", CultureInfo.InvariantCulture);         // Revit 2021+
                            data.Height = UnitUtils.ConvertFromInternalUnits(heightInternal, DisplayUnitType.DUT_METERS).ToString("F3", CultureInfo.InvariantCulture);  // Revit 2021-
                        }
                    }
                    catch
                    {
                        data.Height = "N/A (Calc Error)";
                    }


                    // Volume
                    Parameter volParam = col.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED);
                    if (volParam != null && volParam.HasValue)
                    {
                        //data.Volume = UnitUtils.ConvertFromInternalUnits(volParam.AsDouble(), UnitTypeId.CubicMeters).ToString("F3", CultureInfo.InvariantCulture);             // Revit 2021+
                        data.Volume = UnitUtils.ConvertFromInternalUnits(volParam.AsDouble(), DisplayUnitType.DUT_CUBIC_METERS).ToString("F3", CultureInfo.InvariantCulture);     // Revit 2021-
                    }
                    else
                    {
                        data.Volume = "N/A";
                    }

                    columnsData.Add(data);
                }
            }

            // 2. Export to Excel

            if (!columnsData.Any()) return; // No data to export

            string rvtFileName = doc.Title;
            if (string.IsNullOrEmpty(rvtFileName) || rvtFileName.ToLower().Contains(".rvt")) // Check if it's an unsaved project or already has .rvt
            {
                rvtFileName = Path.GetFileNameWithoutExtension(doc.PathName); // Get filename from path
                if (string.IsNullOrEmpty(rvtFileName)) rvtFileName = "UntitledProject"; // Default for new, unsaved projects
            }
            else // If doc.Title is just the name without .rvt (e.g. after detaching)
            {
                rvtFileName = Path.GetFileNameWithoutExtension(rvtFileName); // Ensure no extension if it was just the title
            }


            string timestamp = DateTime.Now.ToString("yyyy-MM-dd_hh-mm-ss tt", CultureInfo.InvariantCulture);
            string excelFileName = $"{rvtFileName}-Columns Report [{timestamp}].xlsx";
            string fullFilePath = Path.Combine(folderPath, excelFileName);

            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Columns Report");

                    // Headers
                    string[] headers = { "Family", "Type", "ID", "Easting (m)", "Northing (m)", "Base Level", "Base Offset (m)", "Top Level", "Top Offset (m)", "Height (m)", "Volume (m³)" };
                    for (int i = 0; i < headers.Length; i++)
                    {
                        worksheet.Cell(1, i + 1).Value = headers[i];
                        worksheet.Cell(1, i + 1).Style.Font.Bold = true;
                    }

                    // Data
                    int currentRow = 2; // ClosedXML is 1-based, data starts on row 2
                    foreach (var cd in columnsData)
                    {
                        worksheet.Cell(currentRow, 1).Value = cd.Family;
                        worksheet.Cell(currentRow, 2).Value = cd.Type;
                        worksheet.Cell(currentRow, 3).Value = cd.Id;
                        worksheet.Cell(currentRow, 4).Value = cd.Easting;
                        worksheet.Cell(currentRow, 5).Value = cd.Northing;
                        worksheet.Cell(currentRow, 6).Value = cd.BaseLevel;
                        worksheet.Cell(currentRow, 7).Value = cd.BaseOffset;
                        worksheet.Cell(currentRow, 8).Value = cd.TopLevel;
                        worksheet.Cell(currentRow, 9).Value = cd.TopOffset;
                        worksheet.Cell(currentRow, 10).Value = cd.Height;
                        worksheet.Cell(currentRow, 11).Value = cd.Volume;
                        currentRow++;
                    }

                    worksheet.Columns().AdjustToContents(); // Auto-fit all columns

                    workbook.SaveAs(fullFilePath);
                }

                TaskDialog.Show("Success", $"Report generated successfully at:\n{fullFilePath}");
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = fullFilePath,
                    UseShellExecute = true
                };
                Process.Start(psi);
            }
            catch (Exception ex)
            {
                string detailedErrorMessage = $"Error saving Excel file to '{fullFilePath}'.\n";
                detailedErrorMessage += $"Message: {ex.Message}\n";
                if (ex.InnerException != null)
                {
                    detailedErrorMessage += $"Inner Exception Message: {ex.InnerException.Message}\n";
                    detailedErrorMessage += $"Inner Exception StackTrace: {ex.InnerException.StackTrace}\n";
                }
                detailedErrorMessage += $"StackTrace: {ex.StackTrace}";
                TaskDialog.Show("Excel Save Error", detailedErrorMessage);
            }
        }
    }
}
