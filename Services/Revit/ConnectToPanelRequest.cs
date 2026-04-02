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
    public class ConnectToPanelRequest : IExternalEventRequest
    {
        private readonly ElementId _panelId;
        private readonly bool _connectIndividually;
        private readonly string _circuitParamName;
        private readonly string _circuitParamValue;
        private readonly double _maxCableLengthMeters;
        private readonly int _connectionLimit;
        private readonly Action<string> _onComplete;
        private readonly Action<IEnumerable<ElementId>> _onHighlighted;

        public ConnectToPanelRequest(
            ElementId panelId,
            bool connectIndividually,
            string circuitParamName,
            string circuitParamValue,
            double maxCableLengthMeters,
            int connectionLimit,
            Action<string> onComplete,
            Action<IEnumerable<ElementId>> onHighlighted = null)
        {
            _panelId = panelId;
            _connectIndividually = connectIndividually;
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

                // Check existing circuit count against the connection limit before prompting
                if (_connectionLimit > 0)
                {
                    int existingCount = new FilteredElementCollector(doc)
                        .OfClass(typeof(ElectricalSystem))
                        .Cast<ElectricalSystem>()
                        .Count(sys => sys.BaseEquipment?.Id == panel.Id);

                    if (existingCount >= _connectionLimit)
                    {
                        _onComplete?.Invoke(
                            $"\u26a0 Panel \u201c{panel.Name}\u201d has reached its connection limit ({existingCount} / {_connectionLimit}).\n" +
                            $"Increase the limit or choose a different panel.");
                        return;
                    }
                }

                // Let user select elements
                IList<Reference> refs;
                try
                {
                    refs = uidoc.Selection.PickObjects(
                        ObjectType.Element,
                        new ElectricalElementSelectionFilter(),
                        "Select elements to connect to panel. Press Escape when done.");
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

                // panel already resolved above

                int successCount = 0;
                int failCount = 0;

                if (_connectIndividually)
                {
                    // Connect each element to the panel individually (one circuit per element)
                    foreach (var element in elements)
                    {
                        try
                        {
                            using (var tx = new Transaction(doc, "Connect to Panel"))
                            {
                                tx.Start();

                                var systemType = GetElectricalSystemType(element);
                                if (systemType == null)
                                {
                                    sb.AppendLine($"  {element.Name} (Id:{element.Id}) - No electrical connector found.");
                                    failCount++;
                                    tx.RollBack();
                                    continue;
                                }

                                // Create electrical circuit using the element's own connector type
                                var elecSystem = ElectricalSystem.Create(doc,
                                    new List<ElementId> { element.Id },
                                    systemType.Value);

                                if (elecSystem != null)
                                {
                                    // Select the panel as the base equipment
                                    elecSystem.SelectPanel(panel);

                                    // Write circuit parameter if specified
                                    if (!string.IsNullOrWhiteSpace(_circuitParamName))
                                    {
                                        WriteCircuitParameter(elecSystem, _circuitParamName, _circuitParamValue ?? string.Empty);
                                    }

                                    successCount++;
                                    tx.Commit();

                                    // Check cable length against this circuit's actual routed length
                                    if (_maxCableLengthMeters > 0)
                                    {
                                        CheckAndHighlightCableLength(doc, uidoc.ActiveView, elecSystem,
                                            new List<Element> { element }, _maxCableLengthMeters, sb);
                                    }
                                }
                                else
                                {
                                    sb.AppendLine($"  {element.Name} (Id:{element.Id}) - Failed to create circuit.");
                                    failCount++;
                                    tx.RollBack();
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            sb.AppendLine($"  {element.Name} (Id:{element.Id}) - Error: {ex.Message}");
                            failCount++;
                        }
                    }
                }
                else
                {
                    // Connect all elements combined (one circuit for all)
                    try
                    {
                        using (var tx = new Transaction(doc, "Connect to Panel (Combined)"))
                        {
                            tx.Start();

                            // Detect system type from first element's connector
                            var systemType = GetElectricalSystemType(elements[0]);
                            if (systemType == null)
                            {
                                sb.AppendLine("No electrical connector found on selected elements.");
                                failCount = elements.Count;
                                tx.RollBack();
                                _onComplete?.Invoke(sb.ToString().TrimEnd());
                                return;
                            }

                            var elementIds = elements.Select(e => e.Id).ToList();
                            var elecSystem = ElectricalSystem.Create(doc,
                                elementIds,
                                systemType.Value);

                            if (elecSystem != null)
                            {
                                elecSystem.SelectPanel(panel);

                                if (!string.IsNullOrWhiteSpace(_circuitParamName))
                                {
                                    WriteCircuitParameter(elecSystem, _circuitParamName, _circuitParamValue ?? string.Empty);
                                }

                                successCount = elements.Count;
                                tx.Commit();

                                if (_maxCableLengthMeters > 0)
                                {
                                    CheckAndHighlightCableLength(doc, uidoc.ActiveView, elecSystem,
                                        elements, _maxCableLengthMeters, sb);
                                }
                            }
                            else
                            {
                                sb.AppendLine("Failed to create combined circuit.");
                                failCount = elements.Count;
                                tx.RollBack();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        sb.AppendLine($"Error creating combined circuit: {ex.Message}");
                        failCount = elements.Count;
                    }
                }

                // Build result summary
                var header = _connectIndividually ? "Individual" : "Combined";
                sb.Insert(0, $"{header} connection: {successCount} succeeded, {failCount} failed.\n");
                if (successCount > 0)
                    sb.Insert(0, $"Panel: {panel.Name}\n");

                _onComplete?.Invoke(sb.ToString().TrimEnd());
            }
            catch (Exception ex)
            {
                _onComplete?.Invoke($"Error: {ex.Message}");
            }
        }

        private void CheckAndHighlightCableLength(
            Document doc,
            View activeView,
            ElectricalSystem circuit,
            List<Element> elements,
            double maxMeters,
            StringBuilder sb)
        {
            try
            {
                // Read the circuit's actual Length parameter (computed by Revit after SelectPanel)
                // The "Length" parameter on ElectricalSystem is stored in internal Revit units (decimal feet)
                var lengthParam = circuit.LookupParameter("Length");
                if (lengthParam == null || lengthParam.StorageType != StorageType.Double)
                {
                    sb.AppendLine("  [CableCheck] Length parameter not found on circuit.");
                    return;
                }

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
                    _onHighlighted?.Invoke(elements.Select(e => e.Id));
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine($"  [CableCheck] Error: {ex.Message}");
            }
        }

        private static ElectricalSystemType? GetElectricalSystemType(Element element)
        {
            if (element is FamilyInstance fi)
            {
                var mep = fi.MEPModel;
                if (mep?.ConnectorManager?.Connectors != null)
                {
                    foreach (Connector c in mep.ConnectorManager.Connectors)
                    {
                        if (c.Domain == Domain.DomainElectrical)
                            return c.ElectricalSystemType;
                    }
                }
            }
            return null;
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

    internal class ElectricalElementSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            // Allow family instances with electrical connectors
            if (elem is FamilyInstance fi)
            {
                var mep = fi.MEPModel;
                if (mep?.ConnectorManager?.Connectors != null && mep.ConnectorManager.Connectors.Size > 0)
                {
                    foreach (Connector c in mep.ConnectorManager.Connectors)
                    {
                        if (c.Domain == Domain.DomainElectrical)
                            return true;
                    }
                }
            }
            return false;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
}
