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
        private readonly Action<string> _onComplete;

        public ConnectToPanelRequest(
            ElementId panelId,
            bool connectIndividually,
            string circuitParamName,
            string circuitParamValue,
            Action<string> onComplete)
        {
            _panelId = panelId;
            _connectIndividually = connectIndividually;
            _circuitParamName = circuitParamName;
            _circuitParamValue = circuitParamValue;
            _onComplete = onComplete;
        }

        public void Execute(UIApplication app)
        {
            var uidoc = app.ActiveUIDocument;
            var doc = uidoc.Document;
            var sb = new StringBuilder();

            try
            {
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

                var panel = doc.GetElement(_panelId) as FamilyInstance;
                if (panel == null)
                {
                    _onComplete?.Invoke("Panel not found in model.");
                    return;
                }

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
                                }
                                else
                                {
                                    sb.AppendLine($"  {element.Name} (Id:{element.Id}) - Failed to create circuit.");
                                    failCount++;
                                    tx.RollBack();
                                    continue;
                                }

                                tx.Commit();
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
                            }
                            else
                            {
                                sb.AppendLine("Failed to create combined circuit.");
                                failCount = elements.Count;
                            }

                            tx.Commit();
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
