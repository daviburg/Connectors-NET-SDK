//------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All rights reserved.
//------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Azure.Connectors.Sdk;

/// <summary>
/// Envelope type for Connector Namespace trigger callback payloads.
/// The Connector Namespace wraps triggerBody() in one of two shapes:
/// <list type="bullet">
///   <item><description>Batch: <c>{"body":{"value":[...]}}</c> — an array of items.</description></item>
///   <item><description>Single: <c>{"body":{...item...}}</c> — a single item directly (e.g., OnNewEmailV3).</description></item>
/// </list>
/// Both shapes are handled transparently — <see cref="TriggerCallbackBody{T}.Value"/> always
/// contains the item(s) regardless of which shape the Connector Namespace delivered.
/// </summary>
/// <typeparam name="T">The connector-specific trigger item type (e.g., <c>GraphClientReceiveMessage</c> for Office 365 email triggers).</typeparam>
public class TriggerCallbackPayload<T>
{
    /// <summary>
    /// The body envelope containing the trigger items.
    /// </summary>
    [JsonPropertyName("body")]
    public TriggerCallbackBody<T>? Body { get; set; }
}

/// <summary>
/// Inner body of the Connector Namespace trigger callback, containing the trigger items.
/// Supports both batch (<c>{"value":[...]}</c>) and single-item (<c>{...item...}</c>) payloads.
/// The <see cref="TriggerCallbackBodyConverterFactory"/> automatically detects the shape and
/// normalizes single items into <see cref="Value"/> so consumers always iterate one collection.
/// </summary>
/// <typeparam name="T">The connector-specific trigger item type.</typeparam>
[JsonConverter(typeof(TriggerCallbackBodyConverterFactory))]
public class TriggerCallbackBody<T>
{
    /// <summary>
    /// The list of trigger items delivered by the connector trigger.
    /// For single-item triggers (e.g., OnNewEmailV3), this list contains exactly one element.
    /// Split-on is not supported — consumers must iterate this list.
    /// </summary>
    [JsonPropertyName("value")]
    public List<T>? Value { get; set; }
}

/// <summary>
/// JSON converter factory for <see cref="TriggerCallbackBody{T}"/> that handles both
/// batch (<c>{"value":[...]}</c>) and single-item (<c>{...item...}</c>) trigger payloads.
/// </summary>
public class TriggerCallbackBodyConverterFactory : JsonConverterFactory
{
    /// <inheritdoc/>
    public override bool CanConvert(Type typeToConvert)
    {
        return typeToConvert.IsGenericType &&
               typeToConvert.GetGenericTypeDefinition() == typeof(TriggerCallbackBody<>);
    }

    /// <inheritdoc/>
    public override JsonConverter CreateConverter(Type typeToConvert, JsonSerializerOptions options)
    {
        var itemType = typeToConvert.GetGenericArguments()[0];
        var converterType = typeof(TriggerCallbackBodyConverter<>).MakeGenericType(itemType);
        return (JsonConverter)Activator.CreateInstance(converterType)!;
    }
}

/// <summary>
/// JSON converter for <see cref="TriggerCallbackBody{T}"/> that auto-detects the payload shape.
/// If the JSON has a <c>"value"</c> property that is an array, it deserializes as a batch.
/// Otherwise, the entire JSON object is deserialized as a single <typeparamref name="T"/>
/// item and wrapped in a one-element list.
/// </summary>
/// <typeparam name="T">The connector-specific trigger item type.</typeparam>
internal class TriggerCallbackBodyConverter<T> : JsonConverter<TriggerCallbackBody<T>>
{
    /// <inheritdoc/>
    public override TriggerCallbackBody<T>? Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        if (reader.TokenType == JsonTokenType.Null)
        {
            return null;
        }

        using var document = JsonDocument.ParseValue(ref reader);
        var root = document.RootElement;

        // NOTE(daviburg): Batch shape — {"value": [...]}. This is the swagger-documented pattern
        // for TriggerBatchResponse definitions. Most triggers use this shape.
        if (root.TryGetProperty("value", out var valueElement) &&
            valueElement.ValueKind == JsonValueKind.Array)
        {
            var items = valueElement.Deserialize<List<T>>(options);
            return new TriggerCallbackBody<T> { Value = items };
        }

        // NOTE(daviburg): Single-item shape — the body IS the item (no "value" array wrapper).
        // Some triggers (e.g., Office365 OnNewEmailV3) deliver a single item directly in body
        // despite the swagger defining a TriggerBatchResponse. See:
        // https://github.com/Azure/Connectors-NET-SDK/issues/149
        var singleItem = root.Deserialize<T>(options);
        return new TriggerCallbackBody<T>
        {
            Value = singleItem != null ? new List<T> { singleItem } : new List<T>()
        };
    }

    /// <inheritdoc/>
    public override void Write(Utf8JsonWriter writer, TriggerCallbackBody<T> value, JsonSerializerOptions options)
    {
        // NOTE(daviburg): Always serialize using the batch shape for round-trip consistency.
        writer.WriteStartObject();
        writer.WritePropertyName("value");
        JsonSerializer.Serialize(writer, value.Value, options);
        writer.WriteEndObject();
    }
}
