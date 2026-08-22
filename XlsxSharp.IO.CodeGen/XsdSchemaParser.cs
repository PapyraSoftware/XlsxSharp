using System.Collections.Generic;
using XlsxSharp.IO.CodeGen.Model;
using XlsxSharp.IO.CodeGen.Model.Elements;
using XlsxSharp.IO.CodeGen.Model.SimpleTypes;
using XlsxSharp.IO.CodeGen.Model.TopLevel;

namespace XlsxSharp.IO.CodeGen;

/// <summary>
/// Parser to parse XSD of OOXML. It doesn't have to support anythings not found in the official XSD.
/// </summary>
public class XsdSchemaParser
{
    /// <summary>
    /// XSD namespace.
    /// </summary>
    private const string XsdNs = "http://www.w3.org/2001/XMLSchema";

    public Schema ParseSchema(XmlTreeReader reader)
    {
        Schema file = new();

        reader.Open("schema", XsdNs);

        // Read imports
        while (reader.TryOpen("import", XsdNs))
        {
            string ns = reader.GetString("namespace");
            string schemaLocation = reader.GetString("schemaLocation");
            reader.Close("import", XsdNs);

            file.Imports.Add(new ImportElement { Namespace = ns, SchemaLocation = schemaLocation });
        }

        while (!reader.TryClose("schema", XsdNs))
        {
            if (reader.TryOpen("complexType", XsdNs))
            {
                ComplexType complexType = ParseComplexType(reader);
                file.Entries.Add(complexType);
            }
            else if (reader.TryOpen("simpleType", XsdNs))
            {
                ISimpleType simpleType = ParseSimpleType(reader);
                file.Entries.Add(simpleType);
            }
            else if (reader.TryOpen("element", XsdNs))
            {
                string name = reader.GetString("name");
                string typeName = reader.GetString("type");
                reader.Close("element", XsdNs);

                file.Entries.Add(new ElementDefinition { Name = name, TypeName = typeName });
            }
            else if (reader.TryOpen("group", XsdNs))
            {
                string name = reader.GetString("name");
                IElementGroup elementGroup = ParseElementsGroup(reader);
                reader.Close("group", XsdNs);

                file.Entries.Add(new GroupDefinition { Name = name, Content = elementGroup });
            }
            else if (reader.TryOpen("attributeGroup", XsdNs))
            {
                AttributeGroupDefinition attributeGroup = ParseAttributeGroupDefinition(reader);
                file.Entries.Add(attributeGroup);
            }
            else
            {
                throw PartStructureException.ExpectedChoiceElementNotFound(reader);
            }
        }

        return file;
    }

    private static ComplexType ParseComplexType(XmlTreeReader reader)
    {
        string name = reader.GetString("name");
        bool? mixed = reader.GetOptionalBool("mixed");
        if (reader.TryOpen("sequence", XsdNs))
        {
            Occurrences occurrences = GetOccursAttributes(reader);
            List<IElementGroup> elements = [];
            do
            {
                IElementGroup element = ParseElementsGroup(reader);
                elements.Add(element);
            } while (!reader.TryClose("sequence", XsdNs));

            List<OneOf<AttributeElement, AttributeGroupReference>> attributes =
                ParseComplexTypeAttributes(reader);

            return new ComplexTypeSequence
            {
                Name = name,
                Attributes = attributes,
                Mixed = mixed,
                Sequence = new Sequence { Children = elements, Occurrences = occurrences },
            };
        }

        if (reader.TryOpen("choice", XsdNs))
        {
            Occurrences occurrences = GetOccursAttributes(reader);
            List<IElementGroup> choices = [];
            do
            {
                IElementGroup elementGroup = ParseElementsGroup(reader);
                choices.Add(elementGroup);
            } while (!reader.TryClose("choice", XsdNs));

            List<OneOf<AttributeElement, AttributeGroupReference>> attributes =
                ParseComplexTypeAttributes(reader);

            return new ComplexTypeChoice
            {
                Name = name,
                Attributes = attributes,
                Mixed = mixed,
                Choice = new Choice { Children = choices, Occurrences = occurrences },
            };
        }

        if (reader.TryOpen("simpleContent", XsdNs))
        {
            // simpleContent can't have attributes like complexType. It has them only in <extension> tag.
            (
                string baseTypeName,
                List<OneOf<AttributeElement, AttributeGroupReference>> extensionAttributes
            ) = ParseSimpleContent(reader);
            reader.Close("complexType", XsdNs);

            return new ComplexTypeSimpleContent
            {
                Name = name,
                Attributes = extensionAttributes,
                Mixed = mixed,
                BaseTypeName = baseTypeName,
            };
        }

        // Complex type that consists only from attributes
        List<OneOf<AttributeElement, AttributeGroupReference>> attr = ParseComplexTypeAttributes(
            reader
        );
        return new ComplexTypeElement
        {
            Name = name,
            Attributes = attr,
            Mixed = mixed,
        };
    }

    private static ISimpleType ParseSimpleType(XmlTreeReader reader)
    {
        string simpleTypeName = reader.GetString("name");
        if (reader.TryOpen("restriction", XsdNs))
        {
            Restriction restriction = ParseRestriction(reader);
            reader.Close("simpleType", XsdNs);

            return new SimpleType
            {
                Name = simpleTypeName,
                BaseTypeName = restriction.BaseTypeName,
                Restrictions = restriction.ValueRestrictions,
            };
        }

        if (reader.TryOpen("list", XsdNs))
        {
            string itemType = reader.GetString("itemType");
            reader.Close("list", XsdNs);
            reader.Close("simpleType", XsdNs);

            return new SimpleTypeList { Name = simpleTypeName, ItemType = itemType };
        }

        if (reader.TryOpen("union", XsdNs))
        {
            List<Restriction> unionRestrictions = [];
            while (reader.TryOpen("simpleType", XsdNs))
            {
                reader.Open("restriction", XsdNs);
                Restriction restriction = ParseRestriction(reader);
                reader.Close("simpleType", XsdNs);

                unionRestrictions.Add(restriction);
            }

            reader.Close("union", XsdNs);
            reader.Close("simpleType", XsdNs);

            return new SimpleTypeUnion
            {
                Name = simpleTypeName,
                RestrictionsUnion = unionRestrictions,
            };
        }

        throw PartStructureException.ExpectedChoiceElementNotFound(reader);
    }

    private static Restriction ParseRestriction(XmlTreeReader reader)
    {
        string baseType = reader.GetString("base");
        List<IValueRestriction> valueRestrictions = [];

        while (!reader.TryClose("restriction", XsdNs))
        {
            if (reader.TryOpen("enumeration", XsdNs))
            {
                string value = reader.GetString("value");
                valueRestrictions.Add(new RestrictEnumeration(value));
                reader.Close("enumeration", XsdNs);
            }
            else if (reader.TryOpen("length", XsdNs))
            {
                int length = reader.GetInt("value");
                valueRestrictions.Add(new RestrictLength(length));
                reader.Close("length", XsdNs);
            }
            else if (reader.TryOpen("minInclusive", XsdNs))
            {
                string minInclusive = reader.GetString("value");
                valueRestrictions.Add(new RestrictMinInclusive(minInclusive));
                reader.Close("minInclusive", XsdNs);
            }
            else if (reader.TryOpen("minExclusive", XsdNs))
            {
                string minExclusive = reader.GetString("value");
                valueRestrictions.Add(new RestrictMinExclusive(minExclusive));
                reader.Close("minExclusive", XsdNs);
            }
            else if (reader.TryOpen("maxInclusive", XsdNs))
            {
                string maxInclusive = reader.GetString("value");
                valueRestrictions.Add(new RestrictMaxInclusive(maxInclusive));
                reader.Close("maxInclusive", XsdNs);
            }
            else if (reader.TryOpen("maxExclusive", XsdNs))
            {
                string maxExclusive = reader.GetString("value");
                valueRestrictions.Add(new RestrictMaxExclusive(maxExclusive));
                reader.Close("maxExclusive", XsdNs);
            }
            else if (reader.TryOpen("pattern", XsdNs))
            {
                string pattern = reader.GetString("value");
                valueRestrictions.Add(new RestrictPattern(pattern));
                reader.Close("pattern", XsdNs);
            }
            else
            {
                throw PartStructureException.ExpectedChoiceElementNotFound(reader);
            }
        }

        return new Restriction { BaseTypeName = baseType, ValueRestrictions = valueRestrictions };
    }

    private static AttributeGroupDefinition ParseAttributeGroupDefinition(XmlTreeReader reader)
    {
        string name = reader.GetString("name");
        List<AttributeElement> attributes = [];

        while (reader.TryOpen("attribute", XsdNs))
        {
            AttributeElement attribute = ParseAttribute(reader);
            attributes.Add(attribute);
        }

        reader.Close("attributeGroup", XsdNs);

        return new AttributeGroupDefinition { Name = name, Attributes = attributes };
    }

    private static (
        string Base,
        List<OneOf<AttributeElement, AttributeGroupReference>> Attributes
    ) ParseSimpleContent(XmlTreeReader reader)
    {
        reader.Open("extension", XsdNs);
        string baseTypeName = reader.GetString("base");
        List<OneOf<AttributeElement, AttributeGroupReference>> extensionAttributes = [];

        while (!reader.TryClose("extension", XsdNs))
        {
            reader.Open("attribute", XsdNs);
            string name = reader.GetString("name");
            string type = reader.GetString("type");
            AttributeUseType use =
                reader.GetOptionalEnum<AttributeUseType>("use") ?? AttributeUseType.Default;
            string? defaultValue = reader.GetOptionalString("default");
            reader.Close("attribute", XsdNs);
            AttributeElement attribute = new()
            {
                Name = name,
                Type = type,
                Use = use,
                DefaultValue = defaultValue,
                RefName = null,
            };
            extensionAttributes.Add(attribute);
        }

        reader.Close("simpleContent", XsdNs);

        return (baseTypeName, extensionAttributes);
    }

    private static List<
        OneOf<AttributeElement, AttributeGroupReference>
    > ParseComplexTypeAttributes(XmlTreeReader reader)
    {
        List<OneOf<AttributeElement, AttributeGroupReference>> attributes = [];

        while (!reader.TryClose("complexType", XsdNs))
        {
            if (reader.TryOpen("attribute", XsdNs))
            {
                AttributeElement attribute = ParseAttribute(reader);
                attributes.Add(attribute);
            }
            else if (reader.TryOpen("attributeGroup", XsdNs))
            {
                string refName = reader.GetString("ref");
                reader.Close("attributeGroup", XsdNs);
                attributes.Add(new AttributeGroupReference { RefName = refName });
            }
            else
            {
                throw PartStructureException.ExpectedChoiceElementNotFound(reader);
            }
        }

        return attributes;
    }

    private static AttributeElement ParseAttribute(XmlTreeReader reader)
    {
        string? name = reader.GetOptionalString("name");
        string? type = reader.GetOptionalString("type");
        string? refName = reader.GetOptionalString("ref");
        string? defaultValue = reader.GetOptionalString("default");
        AttributeUseType use =
            reader.GetOptionalEnum<AttributeUseType>("use") ?? AttributeUseType.Default;
        reader.Close("attribute", XsdNs);

        return new AttributeElement
        {
            Name = name,
            RefName = refName,
            Type = type,
            Use = use,
            DefaultValue = defaultValue,
        };
    }

    private static IElementGroup ParseElementsGroup(XmlTreeReader reader)
    {
        if (reader.TryOpen("sequence", XsdNs))
        {
            Occurrences occurs = GetOccursAttributes(reader);
            List<IElementGroup> elements = [];
            do
            {
                IElementGroup element = ParseElementsGroup(reader);
                elements.Add(element);
            } while (!reader.TryClose("sequence", XsdNs));

            return new Sequence { Children = elements, Occurrences = occurs };
        }

        if (reader.TryOpen("choice", XsdNs))
        {
            Occurrences occurs = GetOccursAttributes(reader);
            List<IElementGroup> choices = [];
            do
            {
                IElementGroup choice = ParseElementsGroup(reader);
                choices.Add(choice);
            } while (!reader.TryClose("choice", XsdNs));

            return new Choice { Children = choices, Occurrences = occurs };
        }

        if (reader.TryOpen("element", XsdNs))
        {
            Occurrences occurrences = GetOccursAttributes(reader);

            string? refName = reader.GetOptionalString("ref");
            if (refName is not null)
            {
                reader.Close("element", XsdNs);

                return new ElementReference { RefName = refName, Occurrences = occurrences };
            }

            // name, type, min/maxOccurs
            string name = reader.GetString("name");
            string type = reader.GetString("type");
            reader.Close("element", XsdNs);

            return new ElementType
            {
                Name = name,
                TypeName = type,
                Occurrences = occurrences,
            };
        }

        if (reader.TryOpen("group", XsdNs))
        {
            string? refName = reader.GetOptionalString("ref");
            Occurrences occurrences = GetOccursAttributes(reader);

            // Element group reference
            if (refName is not null)
            {
                reader.Close("group", XsdNs);
                return new GroupReference { RefName = refName, Occurrences = occurrences };
            }

            throw PartStructureException.InvalidAttributeValue();
        }

        if (reader.TryOpen("any", XsdNs))
        {
            ProcessContents processContents =
                reader.GetOptionalEnum<ProcessContents>("processContents")
                ?? ProcessContents.Default;
            reader.Close("any", XsdNs);

            return new Any { ProcessContent = processContents };
        }

        throw PartStructureException.ExpectedChoiceElementNotFound(reader);
    }

    private static Occurrences GetOccursAttributes(XmlTreeReader reader)
    {
        int? minOccurs = reader.GetOptionalInt("minOccurs") ?? null;
        int? maxOccurs =
            reader.GetOptionalString("maxOccurs") == "unbounded"
                ? int.MaxValue
                : reader.GetOptionalInt("maxOccurs") ?? null;
        return new Occurrences(minOccurs, maxOccurs);
    }
}
