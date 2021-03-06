﻿using System.Collections.Generic;
using System.ComponentModel.Composition;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.Options.Interface;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.Options
{
    [Export(typeof(IStyleOptions))]
    [ExportMetadata("StyleOrder", 3)]
    class FrostedGlassTextBoxStyleOptions : BaseStyleOptions
    {
        public override List<StyleOption> GetOptionsForVariation()
        {
            var options = GetOptions();
            foreach (var option in options)
            {
                option.IsUseFrostedGlassTextBoxStyle = true;
                option.FontFamily = "Segoe UI";
            }
            return UpdateStyleName(
                options,
                TextCollection.PictureSlidesLabText.StyleNameFrostedGlassTextBox);
        }

        public override StyleOption GetDefaultOptionForPreview()
        {
            return new StyleOption
            {
                StyleName = TextCollection.PictureSlidesLabText.StyleNameFrostedGlassTextBox,
                FontFamily = "Segoe UI",
                IsUseFrostedGlassTextBoxStyle = true
            };
        }
    }
}
