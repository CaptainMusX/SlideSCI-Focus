using System;
using System.Runtime.InteropServices;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SlideSCI
{
    /// <summary>
    /// Small shared boundary for accessing the active PowerPoint context.
    /// Feature code should use this guard before touching ActiveWindow/View/Selection
    /// because those COM properties are unavailable in several PowerPoint views.
    /// </summary>
    internal static class PowerPointContext
    {
        internal static bool TryGetActiveSlide(
            PowerPoint.Application application,
            out PowerPoint.Slide slide)
        {
            slide = null;
            try
            {
                if (application == null
                    || application.ActivePresentation == null
                    || application.ActiveWindow == null
                    || application.ActiveWindow.View == null)
                {
                    return false;
                }

                slide = application.ActiveWindow.View.Slide;
                return slide != null;
            }
            catch
            {
                slide = null;
                return false;
            }
        }

        internal static bool TryGetActiveSelection(
            PowerPoint.Application application,
            out PowerPoint.Selection selection)
        {
            selection = null;
            try
            {
                if (application == null || application.ActiveWindow == null)
                {
                    return false;
                }

                selection = application.ActiveWindow.Selection;
                return selection != null;
            }
            catch
            {
                selection = null;
                return false;
            }
        }

        internal static void Release(object comObject)
        {
            if (comObject == null || !Marshal.IsComObject(comObject)) return;

            try
            {
                Marshal.ReleaseComObject(comObject);
            }
            catch
            {
                // COM cleanup must never mask the original operation result.
            }
        }
    }
}
