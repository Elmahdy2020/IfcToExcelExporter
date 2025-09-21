using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Xbim.Ifc;
using Xbim.Ifc4.Interfaces;
using Xbim.ModelGeometry.Scene;  
using Xbim.Common.Geometry;

namespace IfcExcelExporter
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "IFC files (*.ifc)|*.ifc";
                if (ofd.ShowDialog() != DialogResult.OK) return;

                string ifcPath = ofd.FileName;
                string excelPath = Path.Combine(Path.GetDirectoryName(ifcPath),
                    Path.GetFileNameWithoutExtension(ifcPath) + "_Export.xlsx");

                progressBar.Value = 0;
                lblStatus.Text = "Loading IFC...";

                using (var model = IfcStore.Open(ifcPath))
                {
                   
                    var context = new Xbim3DModelContext(model);
                    context.CreateContext();

                    var entities = model.Instances.OfType<IIfcRoot>()
                        .GroupBy(ent => ent.GetType().Name)
                        .ToList();

                    using (SpreadsheetDocument document = SpreadsheetDocument.Create(excelPath, SpreadsheetDocumentType.Workbook))
                    {
                        WorkbookPart workbookPart = document.AddWorkbookPart();
                        workbookPart.Workbook = new Workbook();
                        Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                        // Add styles
                        WorkbookStylesPart stylesPart = AddStyles(workbookPart);

                        uint sheetId = 1;
                        int groupCount = entities.Count;
                        int groupIndex = 0;

                        foreach (var group in entities)
                        {
                            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                            SheetData sheetData = new SheetData();
                            worksheetPart.Worksheet = new Worksheet(sheetData);

                         
                            var headers = new HashSet<string>
                            {
                                "GlobalId", "Name", "Type",
                                "Top Elevation", "Bottom Elevation", "Length", "Width", "Height"
                            };

                            foreach (var item in group)
                            {
                                foreach (var prop in item.ExpressType.Properties)
                                    headers.Add(prop.Value.Name);

                                if (item is IIfcObject obj)
                                {
                                    var psets = obj.IsDefinedBy
                                        .OfType<IIfcRelDefinesByProperties>()
                                        .Select(r => r.RelatingPropertyDefinition)
                                        .OfType<IIfcPropertySet>();
                                    foreach (var ps in psets)
                                        foreach (var p in ps.HasProperties)
                                            headers.Add(ps.Name + "." + p.Name);

                                    var quants = obj.IsDefinedBy
                                        .OfType<IIfcRelDefinesByProperties>()
                                        .Select(r => r.RelatingPropertyDefinition)
                                        .OfType<IIfcElementQuantity>();
                                    foreach (var qset in quants)
                                        foreach (var q in qset.Quantities)
                                            headers.Add(qset.Name + "." + q.Name);
                                }
                            }

                           
                            Row headerRow = new Row();
                            foreach (var h in headers)
                                headerRow.Append(ConstructCell(h, CellValues.String, 2)); 
                            sheetData.Append(headerRow);

                  
                            foreach (var item in group)
                            {
                                var values = new Dictionary<string, string>
                                {
                                    ["GlobalId"] = item.GlobalId.ToString() ?? "",
                                    ["Name"] = item.Name?.ToString() ?? "",
                                    ["Type"] = item.ExpressType.ExpressName
                                };

                              
                                if (item is IIfcProduct product)
                                {
                                    var shape = context.ShapeInstancesOf(product).FirstOrDefault();
                                    if (shape != null)
                                    {
                                        XbimRect3D box = shape.BoundingBox;
                                        double top = box.Z + box.SizeZ;
                                        double bottom = box.Z;
                                        values["Top Elevation"] = top.ToString("F3");
                                        values["Bottom Elevation"] = bottom.ToString("F3");
                                        values["Length"] = box.SizeX.ToString("F3");
                                        values["Width"] = box.SizeY.ToString("F3");
                                        values["Height"] = box.SizeZ.ToString("F3");
                                    }
                                }


                        
                                foreach (var prop in item.ExpressType.Properties)
                                {
                                    var val = prop.Value.PropertyInfo.GetValue(item);
                                    values[prop.Value.Name] = val?.ToString() ?? "";
                                }

                         
                                if (item is IIfcObject obj2)
                                {
                                    var psets = obj2.IsDefinedBy
                                        .OfType<IIfcRelDefinesByProperties>()
                                        .Select(r => r.RelatingPropertyDefinition)
                                        .OfType<IIfcPropertySet>();
                                    foreach (var ps in psets)
                                    {
                                        foreach (var p in ps.HasProperties)
                                        {
                                            string val = "";
                                            if (p is IIfcPropertySingleValue sv && sv.NominalValue != null)
                                                val = sv.NominalValue.Value?.ToString() ?? "";
                                            else if (p is IIfcPropertyEnumeratedValue ev)
                                                val = string.Join(",", ev.EnumerationValues?.Select(x => x?.Value?.ToString()).Where(x => !string.IsNullOrEmpty(x)) ?? Enumerable.Empty<string>());
                                            else if (p is IIfcPropertyListValue lv)
                                                val = string.Join(",", lv.ListValues?.Select(x => x?.Value?.ToString()).Where(x => !string.IsNullOrEmpty(x)) ?? Enumerable.Empty<string>());

                                            values[ps.Name + "." + p.Name] = val;
                                        }
                                    }

                                    var quants = obj2.IsDefinedBy
                                        .OfType<IIfcRelDefinesByProperties>()
                                        .Select(r => r.RelatingPropertyDefinition)
                                        .OfType<IIfcElementQuantity>();
                                    foreach (var qset in quants)
                                    {
                                        foreach (var q in qset.Quantities)
                                        {
                                            string val = "";
                                            if (q is IIfcQuantityLength ql) val = ql.LengthValue.ToString();
                                            if (q is IIfcQuantityArea qa) val = qa.AreaValue.ToString();
                                            if (q is IIfcQuantityVolume qv) val = qv.VolumeValue.ToString();
                                            if (q is IIfcQuantityCount qc) val = qc.CountValue.ToString();
                                            values[qset.Name + "." + q.Name] = val;
                                        }
                                    }
                                }

                    
                                Row row = new Row();
                                foreach (var h in headers)
                                {
                                    values.TryGetValue(h, out string val);
                                    row.Append(ConstructCell(val, CellValues.String, 1)); 
                                }
                                sheetData.Append(row);
                            }

                          
                            Sheet sheet = new Sheet()
                            {
                                Id = workbookPart.GetIdOfPart(worksheetPart),
                                SheetId = sheetId++,
                                Name = group.Key.Length > 25 ? group.Key.Substring(0, 25) : group.Key
                            };
                            sheets.Append(sheet);
                            AutoFitColumns(worksheetPart, sheetData);

                            groupIndex++;
                            progressBar.Value = (int)((groupIndex / (double)groupCount) * 100);
                            lblStatus.Text = $"Exporting {group.Key} ({groupIndex}/{groupCount})...";
                            Application.DoEvents();
                        }

                        workbookPart.Workbook.Save();

                    }
                }

                lblStatus.Text = "Export finished!";
                progressBar.Value = 100;
                MessageBox.Show("Excel file saved to:\n" + excelPath, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                OpenExcelFile(excelPath);
            }
        }
        private static void OpenExcelFile(string path)
        {
            try
            {
                var psi = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = path,
                    UseShellExecute = true 
                };
                System.Diagnostics.Process.Start(psi);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not open Excel automatically:\n" + ex.Message,
                    "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private static void AutoFitColumns(WorksheetPart worksheetPart, SheetData sheetData)
        {
            
            var columns = new Columns();

         
            int maxColumnCount = sheetData.Elements<Row>().Max(r => r.Elements<Cell>().Count());

            for (int col = 0; col < maxColumnCount; col++)
            {
                double maxWidth = 10;

                foreach (var row in sheetData.Elements<Row>())
                {
                    var cell = row.Elements<Cell>().ElementAtOrDefault(col);
                    if (cell != null && cell.CellValue != null)
                    {
                        string text = cell.CellValue.InnerText;
                        if (!string.IsNullOrEmpty(text))
                        {
                            
                            maxWidth = Math.Max(maxWidth, text.Length + 2);
                        }
                    }
                }

                columns.Append(new Column()
                {
                    Min = (UInt32)(col + 1),
                    Max = (UInt32)(col + 1),
                    Width = maxWidth,
                    CustomWidth = true
                });
            }

          
            worksheetPart.Worksheet.InsertAt(columns, 0);
        }

        private static Cell ConstructCell(string value, CellValues type, uint styleIndex)
        {
            return new Cell()
            {
                CellValue = new CellValue(value ?? ""),
                DataType = new EnumValue<CellValues>(type),
                StyleIndex = styleIndex
            };
        }

        private static WorkbookStylesPart AddStyles(WorkbookPart workbookPart)
        {
            WorkbookStylesPart sp = workbookPart.AddNewPart<WorkbookStylesPart>();
            Stylesheet ss = new Stylesheet();

           
            Fonts fonts = new Fonts(new Font(), new Font(new Bold())); 
           
            Fills fills = new Fills(
                new Fill(new PatternFill() { PatternType = PatternValues.None }),
                new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }),
                new Fill(new PatternFill(new ForegroundColor { Rgb = "FFCCFFFF" }) { PatternType = PatternValues.Solid }) 
            );
          
            Borders borders = new Borders(new Border(), new Border(
                new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                new DiagonalBorder()
            ));
          
            CellFormats cellFormats = new CellFormats(
                new CellFormat(),
                new CellFormat { BorderId = 1, ApplyBorder = true },
                new CellFormat { FontId = 1, FillId = 2, BorderId = 1, ApplyFill = true, ApplyFont = true, ApplyBorder = true } 
            );

            ss.Append(fonts);
            ss.Append(fills);
            ss.Append(borders);
            ss.Append(cellFormats);

            sp.Stylesheet = ss;
            sp.Stylesheet.Save();
            return sp;
        }
    }
}
