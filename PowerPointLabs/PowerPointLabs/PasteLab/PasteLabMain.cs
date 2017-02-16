using System.Windows;

using PowerPointLabs.Utils;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.PasteLab
{
    public class PasteLabMain
    {
        public static void PasteToFillSlide(Models.PowerPointSlide slide, float width, float height)
        {
            if (IsClipboardEmpty())
            {
                return;
            }

            PowerPoint.ShapeRange pastedObject = slide.Shapes.Paste();
            for (int i = 1; i <= pastedObject.Count; i++)
            {
                var shape = new PPShape(pastedObject[i]);
                shape.AbsoluteHeight = height;
                shape.AbsoluteWidth = width;
                shape.VisualTop = 0;
                shape.VisualLeft = 0;
            }
        }

        public static void PasteAndReplace(Models.PowerPointSlide slide, PowerPoint.Selection selection)
        {
            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }

            PowerPoint.Shape selectedShape = selection.ShapeRange[1];

            PowerPoint.Shape newShape = slide.Shapes.Paste()[1];
            newShape.Left = selectedShape.Left;
            newShape.Top = selectedShape.Top;

            foreach (PowerPoint.Effect eff in slide.TimeLine.MainSequence)
            {
                if (eff.Shape == selectedShape)
                {
                    PowerPoint.Effect newEff = slide.TimeLine.MainSequence.Clone(eff);
                    newEff.Shape = newShape;
                    eff.Delete();
                }
            }

            selectedShape.PickUp();
            newShape.Apply();
            selectedShape.Delete();
        }

        internal static bool IsClipboardEmpty()
        {
            return Clipboard.GetDataObject() == null;
        }
    }
}
