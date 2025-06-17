using Bentley.Interop.MicroStationDGN;
public class MicroStationElements
{
    private Element element;

    public string ElementType { get; private set; }
    public long ElementID { get; private set; }
    public string ModelName { get; private set; }
    public string LevelName { get; private set; }
    public int Color { get; private set; }
    public string LineStyle { get; private set; }
    public int LineWeight { get; private set; }
    public bool IsGraphical { get; private set; }

    public MicroStationElements(Element element)
    {
        this.element = element;
        this.ElementType = element.Type.ToString();
        this.ElementID = element.ID;
        this.ModelName = element.ModelReference.Name;
        this.IsGraphical = element.IsGraphical;

        if (this.IsGraphical)
        {
            // Access graphical properties
            this.LevelName = element.Level?.Name ?? "Unnamed Level";
            this.Color = element.Color;
            this.LineStyle = element.LineStyle?.Name ?? "Default";
            this.LineWeight = element.LineWeight;
            
        }
        else
        {
            // Set default or placeholder values for non-graphical elements
            this.LevelName = "N/A";
            this.Color = -1;
            this.LineStyle = "N/A";
            this.LineWeight = -1;
        }
    }

    public string GetSpecificPropertiesAsString()
    {
        // Implement code to get specific properties as string
        // For example, for text elements, get the text string
        if (element is TextElement textElement)
        {
            return $"Text: {textElement.Text}";
        }
        else if (element is LineElement lineElement)
        {
            Point3d start = lineElement.StartPoint;
            Point3d end = lineElement.EndPoint;
            return $"Start: ({start.X}, {start.Y}, {start.Z}), End: ({end.X}, {end.Y}, {end.Z})";
        }
        // Add other element types as needed
        else
        {
            return "";
        }
    }
}
