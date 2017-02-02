﻿using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.PasteLab
{
    class PasteLabMain
    {
        #pragma warning disable 0618
        public static void PasteToCursor(float x, float y)
        {
            PowerPointSlide curslide = PowerPointCurrentPresentationInfo.CurrentSlide;
            PowerPoint.ShapeRange pastedShape = curslide.Shapes.Paste();
            if (pastedShape.Count > 1)
            {
                pastedShape.Group();
            }

            pastedShape.Left = x;
            pastedShape.Top = y;

            if (pastedShape.Count > 1)
            {
                pastedShape.Ungroup();
            }
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

        internal static void PasteIntoSelectedGroup()
        {
            Presentation cur = Globals.ThisAddIn.Application.ActivePresentation;
            PowerPointSlide curslide = PowerPointCurrentPresentationInfo.CurrentSlide;
            PowerPoint.ShapeRange selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;

            var customLayout = cur.SlideMaster.CustomLayouts[2];
            var newSlide = cur.Slides.AddSlide(cur.Slides.Count + 1, customLayout);

            PowerPoint.ShapeRange pastedShapes = curslide.Shapes.Paste();

            selectedShapes.Copy();
            newSlide.Shapes.Paste();
            pastedShapes.Copy();

            List<int> order = new List<int>();

            foreach (Effect eff in curslide.TimeLine.MainSequence)
            {
                if (eff.Shape.Equals(selectedShapes[1]))
                {
                    order.Add(eff.Index);
                }
            }

            selectedShapes = selectedShapes.Ungroup();
            

            List<String> newShapeNames = new List<String>();

            foreach (PowerPoint.Shape shape in selectedShapes)
            {
                newShapeNames.Add(shape.Name);
            }
            
            foreach (PowerPoint.Shape shape in pastedShapes)
            {
                newShapeNames.Add(shape.Name);
            }

            PowerPoint.ShapeRange newShapeRange = curslide.Shapes.Range(newShapeNames.ToArray());
            Shape newGroupedShape = newShapeRange.Group();

            foreach (int curo in order)
            {
                Effect eff = newSlide.TimeLine.MainSequence[1];
                eff.Shape = newGroupedShape;

                if (curslide.TimeLine.MainSequence.Count == 0)
                {
                    Shape tempShape = curslide.Shapes.AddLine(0, 0, 1, 1);
                    Effect tempEff = curslide.TimeLine.MainSequence.AddEffect(tempShape, MsoAnimEffect.msoAnimEffectAppear);
                    eff.MoveAfter(tempEff);
                    tempEff.Delete();
                }
                else if (curslide.TimeLine.MainSequence.Count + 1 < curo)
                {
                    // out of range, assumed to be last
                    eff.MoveAfter(curslide.TimeLine.MainSequence[curslide.TimeLine.MainSequence.Count]);
                }
                else if (curo == 1)
                {
                    // first item!
                    eff.MoveBefore(curslide.TimeLine.MainSequence[1]);
                }
                else
                {
                    eff.MoveAfter(curslide.TimeLine.MainSequence[curo - 1]);
                }
            }

            newSlide.Delete();
        }

        internal static void PasteToPosition()
        {
            Presentation cur = Globals.ThisAddIn.Application.ActivePresentation;
            PowerPointSlide slideToPaste = PowerPointCurrentPresentationInfo.CurrentSlide;

            var customLayout = cur.SlideMaster.CustomLayouts[2];
            var newSlide = cur.Slides.AddSlide(cur.Slides.Count + 1, customLayout);

            PowerPoint.ShapeRange correctShapes = newSlide.Shapes.Paste();

            foreach (PowerPoint.Shape shape in correctShapes)
            {
                shape.Copy();
                PowerPoint.Shape pastedShape = slideToPaste.Shapes.Paste()[1];
                pastedShape.Top = shape.Top;
                pastedShape.Left = shape.Left;
            }

            newSlide.Delete();
        }

        internal static void ReplaceClipboardItem()
        {
            Presentation cur = Globals.ThisAddIn.Application.ActivePresentation;
            PowerPoint.ShapeRange selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;

            var customLayout = cur.SlideMaster.CustomLayouts[2];
            var newSlide = cur.Slides.AddSlide(cur.Slides.Count + 1, customLayout);

            PowerPoint.Shape oldShape = newSlide.Shapes.Paste()[1];

            selectedShapes.Copy();
            PowerPoint.Shape newShape = newSlide.Shapes.Paste()[1];

            foreach (Effect eff in newSlide.TimeLine.MainSequence)
            {
                eff.Shape = newShape;
            }

            newShape.Cut();
            newSlide.Delete();
        }
    }
}
