using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WorshipHelperVSTO
{
    class SelectionManager
    {
        Application app = Globals.ThisAddIn.Application;

        public int GetNextSlideIndex()
        {
            var window = getMainWindow(); // i.e. not the presenter view
            var index = -1;

            // If there are no slides, it is 1
            if (app.ActivePresentation.Slides.Count == 0)
            {
                Debug.WriteLine("There are no slides, so insert index is 1");
                return 1;
            }

            // If in presentation mode, it is the index of the slide after the one currently shown
            if (app.SlideShowWindows.Count > 0)
            {
                index = app.ActivePresentation.SlideShowWindow.View.Slide.SlideIndex + 1;
                Debug.WriteLine($"We are presenting; insert index is {index}");
                return index;
            }

            // If in edit mode, and there is a selection, it is the end of the selection
            if (window.Selection.Type == PpSelectionType.ppSelectionSlides)
            {
                index = getLastSelectedIndex(window.Selection.SlideRange) + 1;
                Debug.WriteLine($"There is an active selection; insert index is {index}");
                return index;
            }

            // If there is no selection, toggle the view to force a selection
            Debug.WriteLine("There is no active selection; toggling view mode");
            toggleViewMode();
            index = window.Selection.SlideRange.SlideIndex + 1;
            Debug.WriteLine($"Insert index is {index}");
            return index;
        }

        public void GoToSlide(int index)
        {
            var window = getMainWindow();
            window.View.GotoSlide(index);
        }

        private int getLastSelectedIndex(SlideRange range)
        {
            int index = -1;
            foreach(Slide slide in range)
            {
                if (slide.SlideIndex > index) index = slide.SlideIndex;
            }
            return index;
        }

        private void toggleViewMode()
        {
            var activeWindow = app.ActiveWindow;
            activeWindow.ViewType = PpViewType.ppViewSlide;
            activeWindow.ViewType = PpViewType.ppViewNormal;
        }

        private DocumentWindow getMainWindow()
        {
            foreach (DocumentWindow win in app.ActivePresentation.Windows)
            {
                // There is probably a better way...
                if (!win.Caption.Contains("Presenter View"))
                {
                    return win;
                }
            }
            return null;
        }
    }
}
