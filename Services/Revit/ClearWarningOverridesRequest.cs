using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;

namespace Connections.Services.Revit
{
    public class ClearWarningOverridesRequest : IExternalEventRequest
    {
        private readonly IEnumerable<ElementId> _elementIds;
        private readonly Action _onComplete;

        public ClearWarningOverridesRequest(IEnumerable<ElementId> elementIds, Action onComplete)
        {
            _elementIds = elementIds;
            _onComplete = onComplete;
        }

        public void Execute(UIApplication app)
        {
            try
            {
                var doc = app.ActiveUIDocument?.Document;
                var view = app.ActiveUIDocument?.ActiveView;
                if (doc == null || view == null) return;

                var defaultOverrides = new OverrideGraphicSettings();

                using (var tx = new Transaction(doc, "Clear Warning Highlights"))
                {
                    tx.Start();
                    foreach (var id in _elementIds)
                        view.SetElementOverrides(id, defaultOverrides);
                    tx.Commit();
                }

                _onComplete?.Invoke();
            }
            catch { }
        }
    }
}
