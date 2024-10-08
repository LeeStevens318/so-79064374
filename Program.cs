using System.Diagnostics;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using BlipFill = DocumentFormat.OpenXml.Drawing.Spreadsheet.BlipFill;
using FromMarker = DocumentFormat.OpenXml.Spreadsheet.FromMarker;
using NonVisualDrawingProperties = DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties;
using NonVisualPictureDrawingProperties = DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureDrawingProperties;
using Path = System.IO.Path;
using Picture = DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture;
using ShapeProperties = DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties;
using ToMarker = DocumentFormat.OpenXml.Spreadsheet.ToMarker;

string path = System.AppContext.BaseDirectory;

// Template File
var templatePath = Path.Join(path, "Assets", "Blank.xlsx");
var templateSheetName = "Sheet1";

// Output File
var destinationPath = Path.Join(path, "Assets", "File.xlsx");

// Image to Add
var image = new ExcelSheetImage
{
	Name = "Logo",
	ImagePath = Path.Join(path, "Assets", "image1.png"),
	Column = 1,
	Row = 1,
	Width = 100,
	Height = 100
};

AddImageToCell(templatePath, destinationPath, image, templateSheetName);


Process.Start(new ProcessStartInfo(destinationPath) { UseShellExecute = true });
return;


void AddImageToCell(string templatePath, string filePath, ExcelSheetImage image, string worksheetName)
{
	File.Copy(templatePath, filePath, overwrite: true);

	using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, true))
	{
		WorkbookPart workbookPart = document.WorkbookPart;
		Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == worksheetName);

		if (sheet == null)
			throw new ArgumentException("Worksheet not found.");

		WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);

		DrawingsPart drawingsPart;
		if (worksheetPart.DrawingsPart == null)
		{
			drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
			worksheetPart.Worksheet.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
		}
		else
		{
			drawingsPart = worksheetPart.DrawingsPart;
		}

		ImagePart imagePart = drawingsPart.AddImagePart(ImagePartType.Png);
		using (FileStream stream = new FileStream(image.ImagePath, FileMode.Open))
		{
			imagePart.FeedData(stream);
		}

		if (drawingsPart.WorksheetDrawing == null)
			drawingsPart.WorksheetDrawing = new WorksheetDrawing();

		WorksheetDrawing worksheetDrawing = drawingsPart.WorksheetDrawing;

		NonVisualDrawingProperties nonVisualProperties = new NonVisualDrawingProperties
		{
			Id = new UInt32Value((uint)(worksheetDrawing.ChildElements.Count + 1)),
			Name = image.Name
		};

		BlipFill blipFill = new BlipFill(
			new Blip { Embed = drawingsPart.GetIdOfPart(imagePart), CompressionState = BlipCompressionValues.Print },
			new Stretch(new FillRectangle())
		);

		Transform2D transform = new Transform2D(
			new Offset { X = 0, Y = 0 },
			new Extents
			{
				Cx = image.Width * 9525, Cy = image.Height * 9525
			} // Width & Height in EMUs (1 pixel = 9525 EMUs)
		);

		ShapeProperties shapeProperties = new ShapeProperties(transform,
			new PresetGeometry(new AdjustValueList()) { Preset = ShapeTypeValues.Rectangle });

		TwoCellAnchor twoCellAnchor = new TwoCellAnchor(
			new FromMarker(
				new ColumnId((image.Column - 1).ToString()),
				new ColumnOffset("0"),
				new RowId((image.Row - 1).ToString()),
				new RowOffset("0")
			),
			new ToMarker(
				new ColumnId(image.Column.ToString()), // Adjust if spanning multiple cells
				new ColumnOffset("0"),
				new RowId(image.Row.ToString()), // Adjust if spanning multiple rows
				new RowOffset("0")
			),
			new Picture(
				nonVisualProperties,
				new NonVisualPictureDrawingProperties(new PictureLocks { NoChangeAspect = true }),
				blipFill,
				shapeProperties
			),
			new ClientData()
		);

		worksheetDrawing.Append(twoCellAnchor);
		worksheetDrawing.Save(drawingsPart);
	}
}

public class ExcelSheetImage
{
	public string Name { get; set; }
	public string ImagePath { get; set; }
	public int Column { get; set; }
	public int Row { get; set; }
	public int Width { get; set; }
	public int Height { get; set; }
}