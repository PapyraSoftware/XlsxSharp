using Newtonsoft.Json;
using XlsxSharp.Parser.Ast;

namespace XlsxSharp.Parser.Function;

internal class AstNodeConverter : JsonConverter<AstNode>
{
    private readonly ReferenceStyle _style;

    internal AstNodeConverter(ReferenceStyle style)
    {
        this._style = style;
    }

    public override AstNode ReadJson(
        JsonReader reader,
        Type objectType,
        AstNode existingValue,
        bool hasExistingValue,
        JsonSerializer serializer
    )
    {
        throw new NotSupportedException("Deserialization of AST is not supported.");
    }

    public override void WriteJson(JsonWriter writer, AstNode value, JsonSerializer serializer)
    {
        this.WriteNode(writer, value);
    }

    private void WriteNode(JsonWriter writer, AstNode value)
    {
        writer.WriteStartObject();
        writer.WritePropertyName("type");
        writer.WriteValue(value.GetTypeString());

        writer.WritePropertyName("content");
        writer.WriteValue(value.GetDisplayString(this._style));

        writer.WritePropertyName("children");
        writer.WriteStartArray();
        foreach (AstNode child in value.Children)
        {
            this.WriteNode(writer, child);
        }

        writer.WriteEndArray();
        writer.WriteEndObject();
    }
}
