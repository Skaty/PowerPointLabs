using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerPointLabs.PasteLab
{
    class PasteLabMain
    {
        #pragma warning disable 0618
        public static void PasteToCursor(float x, float y)
        {
            PowerPointSlide curslide = PowerPointCurrentPresentationInfo.CurrentSlide;
            Shape pastedShape = curslide.Shapes.Paste()[1];
            pastedShape.Left = x;
            pastedShape.Top = y;
        }
        public static void PasteToFit()
        {
            PowerPointSlide curslide = PowerPointCurrentPresentationInfo.CurrentSlide;
            Shape pastedShape = curslide.Shapes.PasteSpecial(PpPasteDataType.ppPasteBitmap)[1];

            pastedShape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
            pastedShape.Left = 0;
            pastedShape.Top = 0;
            pastedShape.Height = PowerPointPresentation.Current.SlideHeight;
            pastedShape.Width = PowerPointPresentation.Current.SlideWidth;
        }
    }
}
