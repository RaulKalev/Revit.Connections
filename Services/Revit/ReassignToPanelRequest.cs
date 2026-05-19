using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Connections.Services.Revit
{
    public class ReassignToPanelRequest : IExternalEventRequest
    {
        private readonly ElementId _panelId;
        private readonly string _circuitParamName;
        private readonly string _circuitParamValue;
        private readonly double _maxCableLengthMeters;
        private readonly int _connectionLimit;
        private readonly Action<string> _onComplete;
        private readonly Action<IEnumerable<ElementId>> _onHighlighted;

        public ReassignToPanelRequest(
            ElementId panelId,
            string circuitParamName,
            string circuitParamValue,
            double maxCableLengthMeters,
            int connectionLimit,
            Action<string> onComplete,
            Action<IEnumerable<ElementId>> onHighlighted = null)
        {
            _panelId = panelId;
            _circuitParamName = circuitParamName;
            _circuitParamValue = circuitParamValue;
            _maxCableLengthMeters = maxCableLengthMeters;
            _connectionLimit = connectionLimit;
            _onComplete = onComplete;
            _onHighlighted = onHighlighted;
        }

        public void Execute(UIApplication app)
        {
            var uidoc = app.ActiveUIDocument;
            var doc = uidoc.Document;
            var sb = new StringBuilder();

            try
            {
                var panel = doc.GetElement(_panelId) as FamilyInstance;
                if (panel == null)
                {
                    _onComplete?.Invoke("Panel not found in model.");
                    return;
                }

                // Let user select elements
                IList<Reference> refs;
                try
                {
                    refs = uidoc.Selection.PickObjects(
                        ObjectType.Element,
                        new ElectricalElementSelectionFilter(),
                        "Select elements to reassign to panel. Press Escape when done.");
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    _onComplete?.Invoke("Selection cancelled.");
                    return;
                }

                if (refs == null || refs.Count == 0)
                {
                    _onComplete?.Invoke("No elements selected.");
                    return;
                }

                var elements = refs
                    .Select(r => doc.GetElement(r.ElementId))
                    .Where(e => e != null)
                    .ToList();

                // Build a list of all circuits to reassign (all circuits per element; skip elements with no circuit)
                var circuitAssignments = new List<(ElectricalSystem Circuit, Element Element)>();
                foreach (var element in elements)
                {
                    var circuits = FindCircuitsForElement(doc, element);
                    if (circuits.Count == 0)
                    {
                        sb.AppendLine($"  {element.Name} (Id:{element.Id}) - No existing circuit found; skipped.");
                        continue;
                    }
                    foreach (var circuit in circuits)
                        circuitAssignments.Add((circuit, element));
                }

                if (circuitAssignments.Count == 0)
                {
                    sb.Insert(0, "No elements with existing circuits were found in selection.\n");
                    _onComplete?.Invoke(sb.ToString().TrimEnd());
                    return;
                }

                // Check connection limit: count circuits NOT already on target panel
                if (_connectionLimit > 0)
                {
                    int existingCount = new FilteredElementCollector(doc)
                        .OfClass(typeof(ElectricalSystem))
                        .Cast<ElectricalSystem>()
                        .Count(sys => sys.BaseEquipment?.Id == panel.Id);

                    int incomingCount = circuitAssignments
                        .Count(ca => ca.Circuit.BaseEquipment?.Id != panel.Id);

                    if (existingCount + incomingCount > _connectionLimit)
                    {
                        _onComplete?.Invoke(
                            $"\u26a0 Panel \u201c{panel.Name}\u201d would exceed its connection limit " +
                            $"({existingCount} existing + {incomingCount} incoming > {_connectionLimit}).\n" +
                            "Increase the limit or choose a different panel.");
                        return;
                    }
                }

                int successCount = 0;
                int failCount = 0;
                var highlightIds = new List<ElementId>();

                foreach (var (circuit, element) in circuitAssignments)
                {
                    // Skip if already on target panel
                    if (circuit.BaseEquipment?.Id == panel.Id)
                    {
                        sb.AppendLine($"  {element.Name} (Id:{element.Id}) - Already on panel \"{panel.Name}\"; skipped.");
                        continue;
                    }

                    try
                    {
                        using (var tx = new Transaction(doc, "Reassign Circuit to Panel"))
                        {
                            tx.Start();

                            circuit.SelectPanel(panel);

                            if (!string.IsNullOrWhiteSpace(_circuitParamName))
                                WriteCircuitParameter(circuit, _circuitParamName, _circuitParamValue ?? string.Empty);

                            tx.Commit();
                            successCount++;
                        }

                        // Cable length check after reassignment
                        if (_maxCableLengthMeters > 0)
                        {
                            CheckAndHighlightCableLength(doc, uidoc.ActiveView, circuit,
                                new List<Element> { element }, _maxCableLengthMeters, sb, highlightIds);
                        }
                    }
                    catch (Exception ex)
                    {
                        sb.AppendLine($"  {element.Name} (Id:{element.Id}) - Error: {ex.Message}");
                        failCount++;
                    }
                }

                sb.Insert(0, $"Panel: {panel.Name}\nReassign: {successCount} succeeded, {failCount} failed.\n");

                if (highlightIds.Count > 0)
                    _onHighlighted?.Invoke(highlightIds);

                _onComplete?.Invoke(sb.ToString().TrimEnd());
            }
            catch (Exception ex)
            {
                _onComplete?.Invoke($"Error: {ex.Message}");
            }
        }

        private static List<ElectricalSystem> FindCircuitsForElement(Document doc, Element element)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(ElectricalSystem))
                .Cast<ElectricalSystem>()
                .Where(sys =>
                {
                    foreach (Element e in sys.Elements)
                        if (e.Id == element.Id) return true;
                    return false;
                })
                .ToList();
        }

        private void CheckAndHighlightCableLength(
            Document doc,
            View activeView,
            ElectricalSystem circuit,
            List<Element> elements,
            double maxMeters,
            StringBuilder sb,
            List<ElementId> highlightIds)
        {
            try
            {
                var lengthParam = circuit.LookupParameter("Length");
                if (lengthParam == null || lengthParam.StorageType != StorageType.Double)
                    return;

                double lengthMeters = UnitUtils.ConvertFromInternalUnits(
                    lengthParam.AsDouble(),
                    UnitTypeId.Meters);

                sb.AppendLine($"  [CableCheck] Circuit length: {lengthMeters:F1} m (limit {maxMeters} m)");

                if (lengthMeters > maxMeters)
                {
                    sb.AppendLine($"  \u26a0 Cable length {lengthMeters:F1} m exceeds limit of {maxMeters} m \u2014 elements highlighted.");

                    var overrides = new OverrideGraphicSettings();
                    var orange = new Color(255, 140, 0);
                    overrides.SetProjectionLineColor(orange);
                    overrides.SetSurfaceForegroundPatternColor(orange);

                    using (var tx = new Transaction(doc, "Highlight Cable Length Exceeded"))
                    {
                        tx.Start();
                        foreach (var element in elements)
                            activeView.SetElementOverrides(element.Id, overrides);
                        tx.Commit();
                    }

                    highlightIds.AddRange(elements.Select(e => e.Id));
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine($"  [CableCheck] Error: {ex.Message}");
            }
        }

        private static void WriteCircuitParameter(ElectricalSystem circuit, string paramName, string value)
        {
            var param = circuit.LookupParameter(paramName);
            if (param != null && !param.IsReadOnly)
            {
                switch (param.StorageType)
                {
                    case StorageType.String:
                        param.Set(value);
                        break;
                    case StorageType.Integer:
                        if (int.TryParse(value, out int intVal))
                            param.Set(intVal);
                        break;
                    case StorageType.Double:
                        if (double.TryParse(value, System.Globalization.NumberStyles.Float,
                            System.Globalization.CultureInfo.InvariantCulture, out double dblVal))
                            param.Set(dblVal);
                        break;
                }
            }
        }
    }
}
