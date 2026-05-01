// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Drawing;

namespace OpenXmlPowerTools
{
    // Internal Drawing helper used by WordprocessingMLDrawingUtil and FormattingAssembler.
    // Isolated here so that FormattingAssembler does not depend on MetricsGetter.
    public static class DrawingTextMetrics
    {
        private static Lazy<Graphics> Graphics { get; } = new Lazy<Graphics>(() =>
        {
            Image image = new Bitmap(1, 1);
            return System.Drawing.Graphics.FromImage(image);
        });

        private static int MeasureText(FontFamily ff, FontStyle fs, decimal sz, string text)
        {
            try
            {
                using (var f = new Font(ff, (float)sz / 2f, fs))
                {
                    var proposedSize = new System.Drawing.Size(int.MaxValue, int.MaxValue);
                    var sf = Graphics.Value.MeasureString(text, f, proposedSize);
                    return (int)sf.Width;
                }
            }
            catch
            {
                return 0;
            }
        }

        public static int GetTextWidth(FontFamily ff, FontStyle fs, decimal sz, string text)
        {
            try
            {
                return MeasureText(ff, fs, sz, text);
            }
            catch (ArgumentException)
            {
                try
                {
                    const FontStyle fs2 = FontStyle.Regular;
                    return MeasureText(ff, fs2, sz, text);
                }
                catch (ArgumentException)
                {
                    try
                    {
                        const FontStyle fs3 = FontStyle.Bold;
                        return MeasureText(ff, fs3, sz, text);
                    }
                    catch (ArgumentException)
                    {
                        var ff2 = new FontFamily("Times New Roman");
                        return MeasureText(ff2, fs, sz, text);
                    }
                }
            }
            catch (OverflowException)
            {
                return 0;
            }
        }
    }
}
