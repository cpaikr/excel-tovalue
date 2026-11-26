namespace XlClean.Core.Styles;

using System.Xml.Linq;

/// <summary>
/// A canonical text form of an element for equality comparisons in the planner.
/// </summary>
/// <remarks>
/// Attributes are ordered by namespace and name, namespace declarations are dropped (names are
/// compared by namespace URI, not prefix), and children and text stay in document order. Nothing
/// is defaulted: an absent attribute and an explicit default differ, so equal text always means
/// equal content and the planner can only miss a merge, never invent one.
/// </remarks>
internal static class CanonicalXml
{
    public static string Of(XElement element, params XName[] excludedAttributes) =>
        Normalize(element, excludedAttributes).ToString(SaveOptions.DisableFormatting);

    private static XElement Normalize(XElement element, XName[] excludedAttributes)
    {
        XElement normalized = new(element.Name);
        foreach (XAttribute attribute in element.Attributes()
            .Where(attribute => !attribute.IsNamespaceDeclaration && !excludedAttributes.Contains(attribute.Name))
            .OrderBy(attribute => attribute.Name.NamespaceName, StringComparer.Ordinal)
            .ThenBy(attribute => attribute.Name.LocalName, StringComparer.Ordinal))
        {
            normalized.Add(new XAttribute(attribute.Name, attribute.Value));
        }

        foreach (XNode node in element.Nodes())
        {
            normalized.Add(node switch
            {
                XElement child => Normalize(child, []),
                XText text => new XText(text.Value),
                _ => null,
            });
        }

        return normalized;
    }
}
