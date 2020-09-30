using FasType.Abbreviations;
using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace FasType.Converters
{
    public class IEnumerableConverter : JsonConverter<IEnumerable<IAbbreviation>>
    {
        private readonly JsonConverter<IAbbreviation> _converter;

        public IEnumerableConverter(JsonSerializerOptions options)
        {
            // For performance, use the existing converter if available.
            _converter = (JsonConverter<IAbbreviation>)options.GetConverter(typeof(IAbbreviation)) ?? new IAbbreviationConverter();
        }

        public override bool CanConvert(Type typeToConvert)
        {
            bool isType = typeToConvert == typeof(IEnumerable<IAbbreviation>) || typeToConvert.GetInterface(typeof(IEnumerable<>).FullName) != null;
            return isType;// base.CanConvert(typeToConvert);
        }

        public override IEnumerable<IAbbreviation> Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            var list = new List<IAbbreviation>();
            if (reader.TokenType == JsonTokenType.StartArray)
            {
                while (reader.Read() && reader.TokenType != JsonTokenType.EndArray)
                {
                    if (reader.TokenType == JsonTokenType.StartObject)
                        reader.Read();
                    string typeString = reader.GetString();
                    Type type = Type.GetType(typeString);
                    reader.Read();
                    IAbbreviation abbrev = null;
                    if (_converter.CanConvert(type))
                        abbrev = _converter.Read(ref reader, type, options);
                    list.Add(abbrev);
                }
            }
            return list;
        }

        public override void Write(Utf8JsonWriter writer, IEnumerable<IAbbreviation> value, JsonSerializerOptions options)
        {
            writer.WriteStartArray();
            foreach (var abbrev in value)
            {
                _converter.Write(writer, abbrev, options);
            }
            writer.WriteEndArray();
        }
    }

    //public class ILookupConverterFactory : JsonConverterFactory
    //{
    //    public override bool CanConvert(Type typeToConvert)
    //    {
    //        if (!typeToConvert.IsGenericType)
    //        {
    //            return false;
    //        }
    //        var generics = typeToConvert.GetGenericArguments();
    //        var fullname = typeof(IAbbreviation).FullName;
    //        if (generics.Length < 2 || generics.All(t => t.GetInterface(fullname) == null && t != typeof(IAbbreviation)))
    //        {
    //            return false;
    //        }
    //        return typeToConvert.GetInterfaces().Any(t => t != typeof(ILookup<,>));
    //    }

    //    public override JsonConverter CreateConverter(Type typeToConvert, JsonSerializerOptions options)
    //    {
    //        var generics = typeToConvert.GetGenericArguments();

    //        Type keyType = generics[0];
    //        Type valueType = generics[1];

    //        JsonConverter converter = (JsonConverter)Activator.CreateInstance(
    //            typeof(InnerILookupConverter<>).MakeGenericType(
    //                new Type[] { valueType }),
    //            BindingFlags.Instance | BindingFlags.Public,
    //            binder: null,
    //            args: new object[] { options },
    //            culture: null);

    //        return converter;
    //    }

    //    private class InnerILookupConverter<TValue> 
    //        : JsonConverter<ILookup<string, TValue>>
    //        where TValue : IAbbreviation
    //    {
    //        private readonly JsonConverter<TValue> _dictionnaryConverter;
    //        private Type _valueType;

    //        public InnerILookupConverter(JsonSerializerOptions options)
    //        {                
    //            // For performance, use the existing converter if available.
    //            _dictionnaryConverter = (JsonConverter<TValue>)options.GetConverter(typeof(TValue));

    //            // Cache the key and value types.
    //            _valueType = typeof(TValue);
    //        }

    //        public override ILookup<string, TValue> Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    //        {
    //            var list = JsonSerializer.Deserialize<List<TValue>>(ref reader, options);

    //            var lookup = list.ToLookup(e => string.Join("", e.ShortForm.Take(2)), e => e);
    //            return lookup;
    //            //var lookup = dtos.SelectMany(grp => grp.Values.Select(v => new { grp.Key, v })).ToLookup(e => e.Key, e => e.v);
    //            //return lookup;
    //        }

    //        public override void Write(Utf8JsonWriter writer, ILookup<string, TValue> value, JsonSerializerOptions options)
    //        {
    //            var list = value.SelectMany(grp => grp.AsEnumerable()).ToList();
    //            JsonSerializer.Serialize(writer, list, options);
    //            //_dictionnaryConverter.Write(writer, dict, options);
    //            //throw new NotImplementedException();
    //        }
    //    }
    //}
}
