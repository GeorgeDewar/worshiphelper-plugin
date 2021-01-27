using Microsoft.Office.Interop.PowerPoint;

namespace WorshipHelperVSTO
{
    class WindowManager
    {
        Application app = Globals.ThisAddIn.Application;

        public DocumentWindow GetMainWindow()
        {
            if (app.Presentations.Count == 0)
            {
                return null;
            }

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

        public DocumentWindow GetPresenterView()
        {
            if (app.Presentations.Count == 0)
            {
                return null;
            }
            
            foreach (DocumentWindow win in app.ActivePresentation.Windows)
            {
                // There is probably a better way...
                if (win.Caption.Contains("Presenter View"))
                {
                    return win;
                }
            }
            return null;
        }


    }
}
